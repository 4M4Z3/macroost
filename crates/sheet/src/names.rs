//! `Names` and `Name`: defined names (named ranges) stored through
//! [`crate::Workbook::defined_names`]. `Workbook.Names` lists every name (sheet-scoped ones as
//! `Sheet1!Name`), `Worksheet.Names` only that sheet's; both in alphabetical order.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::address::{format_area, parse_reference, quote_sheet};
use crate::objects::{active_sheet, app_value, workbook_value};
use crate::range::{Kind, decode_range, range_value};
use crate::state::{Shared, data, err1004, with, with_mut};
use crate::workbook::SheetId;

/// `Workbook.Names` (`scope` `None`) or `Worksheet.Names`, or `Names(index)` with arguments.
pub(crate) fn names_member(app: &Shared, scope: Option<SheetId>, args: &Args) -> Result<Value> {
    let coll = NamesObj { app: app.clone(), scope };
    if args.is_empty() { Ok(Value::Object(Some(Rc::new(coll)))) } else { coll.call("Item", args) }
}

struct NamesObj {
    app: Shared,
    scope: Option<SheetId>,
}

/// Splits `Sheet1!X` / `'My Sheet'!X` into (sheet name, name).
fn split_scope(s: &str) -> (Option<String>, &str) {
    match s.rsplit_once('!') {
        Some((sh, n)) => {
            let sh = sh.strip_prefix('\'').and_then(|x| x.strip_suffix('\'')).unwrap_or(sh);
            (Some(sh.replace("''", "'")), n)
        }
        None => (None, s),
    }
}

/// Refers-to text: a Range object becomes its absolute qualified address; text loses its `=`.
fn refers_to_text(app: &Shared, v: &Value) -> Result<String> {
    if let Some((sheet, areas)) = decode_range(v) {
        return with(app, |st| {
            let s = quote_sheet(&st.book.sheet_name(sheet).unwrap_or_default());
            let parts: Vec<String> = areas.iter().map(|a| format!("{s}!{}", format_area(a, true, true))).collect();
            Ok(parts.join(","))
        });
    }
    let t = data(v)?.to_text()?;
    Ok(t.strip_prefix('=').unwrap_or(&t).to_owned())
}

fn valid_name(n: &str) -> bool {
    n.chars().next().is_some_and(|c| c.is_alphabetic() || c == '_' || c == '\\')
        && n.chars().all(|c| c.is_alphanumeric() || matches!(c, '_' | '.' | '\\'))
        && parse_reference(n).is_none()
        && !n.eq_ignore_ascii_case("r")
        && !n.eq_ignore_ascii_case("c")
}

/// The names a collection lists, sorted case-insensitively.
fn listed(app: &Shared, scope: Option<SheetId>) -> Result<Vec<(String, Option<SheetId>, String)>> {
    with(app, |st| {
        let mut v: Vec<_> =
            st.book.defined_names().into_iter().filter(|(_, s, _)| scope.is_none() || *s == scope).collect();
        v.sort_by_key(|(n, _, _)| n.to_lowercase());
        Ok(v)
    })
}

/// Adds or replaces a name; returns it.
pub(crate) fn add_name(app: &Shared, full: &str, default_scope: Option<SheetId>, refers: &str) -> Result<Value> {
    let (sheet, n) = split_scope(full);
    let scope = match sheet {
        Some(sh) => Some(with(app, |st| {
            st.sheet_by_name(&sh).ok_or_else(|| err1004("The name that you entered is not valid."))
        })?),
        None => default_scope,
    };
    if !valid_name(n) || refers.is_empty() {
        return Err(err1004("The name that you entered is not valid."));
    }
    // An unqualified reference belongs to the scope's sheet (or the active sheet).
    let refers = match parse_reference(refers) {
        Some(p) if p.sheet.is_none() => {
            let sheet = match scope {
                Some(s) => s,
                None => active_sheet(app)?,
            };
            let sh = with(app, |st| Ok(quote_sheet(&st.book.sheet_name(sheet).unwrap_or_default())))?;
            refers.split(',').map(|part| format!("{sh}!{}", part.trim())).collect::<Vec<_>>().join(",")
        }
        _ => refers.to_owned(),
    };
    with_mut(app, |st| {
        st.book.set_defined_name(n, scope, &refers);
        Ok(())
    })?;
    Ok(Value::Object(Some(Rc::new(NameObj { app: app.clone(), name: n.to_owned(), scope }))))
}

impl Object for NamesObj {
    fn type_name(&self) -> String {
        "Names".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let app = &self.app;
        let missing = || err1004("Application-defined or object-defined error");
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(i32::try_from(listed(app, self.scope)?.len()).unwrap_or(i32::MAX))),
            "" | "item" => {
                let idx = data(args.require(0, "Index")?)?;
                let names = listed(app, self.scope)?;
                let found = match &idx {
                    Value::String(s) => {
                        let (sheet, n) = split_scope(s);
                        let scope = match sheet {
                            Some(sh) => Some(with(app, |st| st.sheet_by_name(&sh).ok_or_else(missing))?),
                            None => self.scope,
                        };
                        // A bare name prefers the workbook scope, then any sheet's.
                        names
                            .iter()
                            .find(|(x, sc, _)| *sc == scope && x.eq_ignore_ascii_case(n))
                            .or_else(|| names.iter().find(|(x, _, _)| x.eq_ignore_ascii_case(n)))
                            .cloned()
                    }
                    other => {
                        let i = other.to_i64()?;
                        usize::try_from(i - 1).ok().and_then(|i| names.get(i)).cloned()
                    }
                };
                let (n, scope, _) = found.ok_or_else(missing)?;
                Ok(Value::Object(Some(Rc::new(NameObj { app: app.clone(), name: n, scope }))))
            }
            "add" => {
                let full = data(args.require(0, "Name")?)?.to_text()?;
                let refers = match (args.get(1, "RefersTo"), args.get(9, "RefersToR1C1")) {
                    (Some(v), _) => refers_to_text(app, v)?,
                    (None, Some(v)) => {
                        let t = data(v)?.to_text()?;
                        let a1 = crate::formula::r1c1_to_a1(&t, 1, 1);
                        a1.strip_prefix('=').unwrap_or(&a1).to_owned()
                    }
                    (None, None) => return Err(RuntimeError::standard(449)),
                };
                add_name(app, &full, self.scope, &refers)
            }
            "application" => Ok(app_value(app)),
            "parent" => Ok(workbook_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(listed(&self.app, self.scope)?
            .into_iter()
            .map(|(n, s, _)| {
                Value::Object(Some(Rc::new(NameObj { app: self.app.clone(), name: n, scope: s }) as Rc<dyn Object>))
            })
            .collect())
    }
}

pub(crate) struct NameObj {
    pub app: Shared,
    pub name: String,
    pub scope: Option<SheetId>,
}

impl NameObj {
    fn refers(&self) -> Result<String> {
        with(&self.app, |st| {
            st.book
                .defined_names()
                .into_iter()
                .find(|(n, s, _)| *s == self.scope && n.eq_ignore_ascii_case(&self.name))
                .map(|(_, _, r)| r)
                .ok_or_else(|| err1004("Application-defined or object-defined error"))
        })
    }
}

impl Object for NameObj {
    fn type_name(&self) -> String {
        "Name".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let app = &self.app;
        match name.to_ascii_lowercase().as_str() {
            "name" | "namelocal" => match self.scope {
                None => Ok(Value::from(self.name.as_str())),
                Some(s) => with(app, |st| {
                    let sh = quote_sheet(&st.book.sheet_name(s).unwrap_or_default());
                    Ok(Value::from(format!("{sh}!{}", self.name)))
                }),
            },
            "" | "refersto" | "value" | "referstolocal" => Ok(Value::from(format!("={}", self.refers()?))),
            "referstor1c1" | "referstor1c1local" => {
                Ok(Value::from(crate::formula::a1_to_r1c1(&format!("={}", self.refers()?), 1, 1)))
            }
            "referstorange" => {
                let refers = self.refers()?;
                let fail = || err1004("Application-defined or object-defined error");
                let p = parse_reference(&refers).ok_or_else(fail)?;
                let sheet = match (&p.sheet, self.scope) {
                    (Some(n), _) => with(app, |st| st.sheet_by_name(n).ok_or_else(fail))?,
                    (None, Some(s)) => s,
                    (None, None) => active_sheet(app)?,
                };
                Ok(range_value(app, sheet, p.areas, Kind::Cells))
            }
            "visible" => Ok(Value::Boolean(true)),
            "delete" => with_mut(app, |st| {
                st.book.set_defined_name(&self.name, self.scope, "");
                Ok(Value::Empty)
            }),
            "application" => Ok(app_value(app)),
            "parent" => Ok(workbook_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "refersto" | "value" | "referstolocal" => {
                let t = refers_to_text(&self.app, &value)?;
                if t.is_empty() {
                    return Err(err1004("The formula contains an error."));
                }
                with_mut(&self.app, |st| {
                    st.book.set_defined_name(&self.name, self.scope, &t);
                    Ok(())
                })
            }
            "visible" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}
