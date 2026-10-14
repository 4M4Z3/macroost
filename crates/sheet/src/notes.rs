//! Cell comments (`Comment`, `Comments`, `Range.AddComment`) and hyperlinks (`Hyperlink`,
//! `Hyperlinks`), stored per cell through the [`crate::Workbook`] trait.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::address::Rect;
use crate::range::{Kind, decode_range, range_value};
use crate::state::{Shared, data, err1004, opt_text, with, with_mut};
use crate::workbook::{Hyperlink, SheetId};

fn obj(o: impl Object + 'static) -> Value {
    Value::Object(Some(Rc::new(o)))
}

/// Cells with a comment inside `area`, row by row.
fn commented(app: &Shared, sheet: SheetId, area: Rect) -> Result<Vec<(u32, u32)>> {
    with(app, |st| {
        Ok(st
            .book
            .occupied_cells(sheet, area)
            .into_iter()
            .filter(|&(r, c)| st.book.cell_comment(sheet, r, c).is_some())
            .collect())
    })
}

/// `Range.Comment`: the first cell's comment or `Nothing`.
pub(crate) fn comment_of(app: &Shared, sheet: SheetId, r: u32, c: u32) -> Result<Value> {
    let has = with(app, |st| Ok(st.book.cell_comment(sheet, r, c).is_some()))?;
    Ok(if has { obj(CommentObj { app: app.clone(), sheet, r, c }) } else { Value::Object(None) })
}

/// `Range.AddComment(Text)`.
pub(crate) fn add_comment(app: &Shared, sheet: SheetId, r: u32, c: u32, args: &Args) -> Result<Value> {
    let text = opt_text(args, 0, "Text")?.unwrap_or_default();
    with_mut(app, |st| {
        if st.book.cell_comment(sheet, r, c).is_some() {
            return Err(err1004("Application-defined or object-defined error"));
        }
        st.book.set_cell_comment(sheet, r, c, Some(text))
    })?;
    Ok(obj(CommentObj { app: app.clone(), sheet, r, c }))
}

pub(crate) fn clear_comments(app: &Shared, sheet: SheetId, areas: &[Rect]) -> Result<()> {
    for a in areas {
        for (r, c) in commented(app, sheet, *a)? {
            with_mut(app, |st| {
                st.visible_comments.remove(&(sheet, r, c));
                st.book.set_cell_comment(sheet, r, c, None)
            })?;
        }
    }
    Ok(())
}

/// Removes the comments in `area` (inside an object-model call).
pub(crate) fn clear_comments_in(st: &mut crate::state::State, sheet: SheetId, area: Rect) -> Result<()> {
    for (r, c) in st.book.occupied_cells(sheet, area) {
        if st.book.cell_comment(sheet, r, c).is_some() {
            st.visible_comments.remove(&(sheet, r, c));
            st.book.set_cell_comment(sheet, r, c, None)?;
        }
    }
    Ok(())
}

/// `Worksheet.Comments` / `Comments(i)`.
pub(crate) fn comments_value(app: &Shared, sheet: SheetId, args: &Args) -> Result<Value> {
    let coll = CommentsObj { app: app.clone(), sheet };
    if args.is_empty() { Ok(obj(coll)) } else { coll.call("Item", args) }
}

struct CommentObj {
    app: Shared,
    sheet: SheetId,
    r: u32,
    c: u32,
}

impl Object for CommentObj {
    fn type_name(&self) -> String {
        "Comment".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let app = &self.app;
        let (s, r, c) = (self.sheet, self.r, self.c);
        match name.to_ascii_lowercase().as_str() {
            "text" | "" => {
                if let Some(t) = opt_text(args, 0, "Text")? {
                    let start = args.get(1, "Start").map(|v| v.to_i64()).transpose()?;
                    let overwrite = args.get(2, "Overwrite").map(|v| v.to_bool()).transpose()?.unwrap_or(true);
                    with_mut(app, |st| {
                        let old = st.book.cell_comment(s, r, c).unwrap_or_default();
                        let new = match start {
                            Some(p) if p > 1 => {
                                let keep: String = old.chars().take(p as usize - 1).collect();
                                let tail: String = if overwrite {
                                    old.chars().skip(p as usize - 1 + t.chars().count()).collect()
                                } else {
                                    old.chars().skip(p as usize - 1).collect()
                                };
                                format!("{keep}{t}{tail}")
                            }
                            _ => t.clone(),
                        };
                        st.book.set_cell_comment(s, r, c, Some(new))
                    })?;
                }
                with(app, |st| Ok(Value::from(st.book.cell_comment(s, r, c).unwrap_or_default())))
            }
            "delete" => with_mut(app, |st| {
                st.visible_comments.remove(&(s, r, c));
                st.book.set_cell_comment(s, r, c, None).map(|_| Value::Empty)
            }),
            "visible" => with(app, |st| Ok(Value::Boolean(st.visible_comments.contains(&(s, r, c))))),
            "author" => Ok(Value::from("")),
            "parent" => Ok(range_value(app, s, vec![Rect::cell(r, c)], Kind::Cells)),
            "next" | "previous" => {
                let all = commented(app, s, Rect::sheet())?;
                let i = all.iter().position(|&p| p == (r, c));
                let j = match (i, name.eq_ignore_ascii_case("next")) {
                    (Some(i), true) => all.get(i + 1),
                    (Some(i), false) => i.checked_sub(1).and_then(|j| all.get(j)),
                    _ => None,
                };
                Ok(match j {
                    Some(&(r, c)) => obj(CommentObj { app: app.clone(), sheet: s, r, c }),
                    None => Value::Object(None),
                })
            }
            "application" => Ok(crate::objects::app_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let key = (self.sheet, self.r, self.c);
        match name.to_ascii_lowercase().as_str() {
            "visible" => {
                let v = data(&value)?.to_bool()?;
                with_mut(&self.app, |st| {
                    if v {
                        st.visible_comments.insert(key);
                    } else {
                        st.visible_comments.remove(&key);
                    }
                    Ok(())
                })
            }
            "shape" | "author" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

struct CommentsObj {
    app: Shared,
    sheet: SheetId,
}

impl Object for CommentsObj {
    fn type_name(&self) -> String {
        "Comments".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let all = commented(&self.app, self.sheet, Rect::sheet())?;
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(all.len() as i32)),
            "" | "item" => {
                let i = data(args.require(0, "Index")?)?.to_i64()?;
                let &(r, c) = usize::try_from(i - 1)
                    .ok()
                    .and_then(|i| all.get(i))
                    .ok_or_else(|| err1004("Application-defined or object-defined error"))?;
                Ok(obj(CommentObj { app: self.app.clone(), sheet: self.sheet, r, c }))
            }
            "application" | "parent" => Ok(crate::objects::app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(commented(&self.app, self.sheet, Rect::sheet())?
            .into_iter()
            .map(|(r, c)| obj(CommentObj { app: self.app.clone(), sheet: self.sheet, r, c }))
            .collect())
    }
}

// ---------------------------------------------------------------- hyperlinks

/// A web address as the application stores it (`https://example.com` gains its `/`).
fn normalize_address(a: &str) -> String {
    let lower = a.to_ascii_lowercase();
    for scheme in ["http://", "https://"] {
        if lower.starts_with(scheme) {
            let rest = &a[scheme.len()..];
            if !rest.is_empty() && !rest.contains('/') && !rest.contains('?') && !rest.contains('#') {
                return format!("{a}/");
            }
        }
    }
    a.to_owned()
}

fn linked(app: &Shared, sheet: SheetId, area: Rect) -> Result<Vec<(u32, u32)>> {
    with(app, |st| {
        Ok(st
            .book
            .occupied_cells(sheet, area)
            .into_iter()
            .filter(|&(r, c)| st.book.cell_hyperlink(sheet, r, c).is_some())
            .collect())
    })
}

/// `Worksheet.Hyperlinks` (`area` `None`) or `Range.Hyperlinks`, or `Hyperlinks(i)`.
pub(crate) fn hyperlinks_value(app: &Shared, sheet: SheetId, area: Option<Vec<Rect>>, args: &Args) -> Result<Value> {
    let coll = HyperlinksObj { app: app.clone(), sheet, areas: area.unwrap_or_else(|| vec![Rect::sheet()]) };
    if args.is_empty() { Ok(obj(coll)) } else { coll.call("Item", args) }
}

pub(crate) fn clear_hyperlinks(app: &Shared, sheet: SheetId, areas: &[Rect]) -> Result<()> {
    for a in areas {
        for (r, c) in linked(app, sheet, *a)? {
            with_mut(app, |st| st.book.set_cell_hyperlink(sheet, r, c, None))?;
        }
    }
    Ok(())
}

struct HyperlinksObj {
    app: Shared,
    sheet: SheetId,
    areas: Vec<Rect>,
}

impl HyperlinksObj {
    fn cells(&self) -> Result<Vec<(u32, u32)>> {
        let mut out = Vec::new();
        for a in &self.areas {
            out.extend(linked(&self.app, self.sheet, *a)?);
        }
        Ok(out)
    }
}

impl Object for HyperlinksObj {
    fn type_name(&self) -> String {
        "Hyperlinks".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let app = &self.app;
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.cells()?.len() as i32)),
            "" | "item" => {
                let i = data(args.require(0, "Index")?)?.to_i64()?;
                let cells = self.cells()?;
                let &(r, c) =
                    usize::try_from(i - 1).ok().and_then(|i| cells.get(i)).ok_or_else(RuntimeError::subscript)?;
                Ok(obj(HyperlinkObj { app: app.clone(), sheet: self.sheet, r, c }))
            }
            "add" => {
                let anchor = args.require(0, "Anchor")?;
                let (sheet, areas) = decode_range(anchor).ok_or_else(RuntimeError::type_mismatch)?;
                let address = normalize_address(&opt_text(args, 1, "Address")?.unwrap_or_default());
                let sub = opt_text(args, 2, "SubAddress")?.unwrap_or_default();
                let tip = opt_text(args, 3, "ScreenTip")?.unwrap_or_default();
                let text = opt_text(args, 4, "TextToDisplay")?;
                let a = areas[0];
                with_mut(app, |st| {
                    for r in a.r1..=a.r2.min(a.r1 + 10_000) {
                        for c in a.c1..=a.c2.min(a.c1 + 10_000) {
                            let shown = match &text {
                                Some(t) => Some(t.clone()),
                                None if st.is_blank(sheet, r, c) => {
                                    Some(if address.is_empty() { sub.clone() } else { address.clone() })
                                }
                                None => None,
                            };
                            if let Some(t) = shown {
                                st.book.set_cell_value(sheet, r, c, Value::from(t))?;
                            }
                            st.book.set_cell_hyperlink(
                                sheet,
                                r,
                                c,
                                Some(Hyperlink {
                                    address: address.clone(),
                                    sub_address: sub.clone(),
                                    screen_tip: tip.clone(),
                                }),
                            )?;
                        }
                    }
                    Ok(())
                })?;
                Ok(obj(HyperlinkObj { app: app.clone(), sheet, r: a.r1, c: a.c1 }))
            }
            "delete" => clear_hyperlinks(app, self.sheet, &self.areas).map(|_| Value::Empty),
            "application" | "parent" => Ok(crate::objects::app_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(self
            .cells()?
            .into_iter()
            .map(|(r, c)| obj(HyperlinkObj { app: self.app.clone(), sheet: self.sheet, r, c }))
            .collect())
    }
}

struct HyperlinkObj {
    app: Shared,
    sheet: SheetId,
    r: u32,
    c: u32,
}

impl HyperlinkObj {
    fn link(&self) -> Result<Hyperlink> {
        with(&self.app, |st| {
            st.book.cell_hyperlink(self.sheet, self.r, self.c).ok_or_else(RuntimeError::object_required)
        })
    }
}

impl Object for HyperlinkObj {
    fn type_name(&self) -> String {
        "Hyperlink".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let app = &self.app;
        match name.to_ascii_lowercase().as_str() {
            "address" => Ok(Value::from(self.link()?.address)),
            "subaddress" => Ok(Value::from(self.link()?.sub_address)),
            "screentip" => Ok(Value::from(self.link()?.screen_tip)),
            "texttodisplay" | "name" => with(app, |st| Ok(Value::from(st.text(self.sheet, self.r, self.c)))),
            "range" | "parent" => Ok(range_value(app, self.sheet, vec![Rect::cell(self.r, self.c)], Kind::Cells)),
            "delete" => {
                with_mut(app, |st| st.book.set_cell_hyperlink(self.sheet, self.r, self.c, None)).map(|_| Value::Empty)
            }
            "follow" => Ok(Value::Empty),
            "type" => Ok(Value::Long(0)),
            "application" => Ok(crate::objects::app_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let v = data(&value)?.to_text()?;
        let mut l = self.link()?;
        match name.to_ascii_lowercase().as_str() {
            "address" => l.address = normalize_address(&v),
            "subaddress" => l.sub_address = v,
            "screentip" => l.screen_tip = v,
            "texttodisplay" => {
                return with_mut(&self.app, |st| st.book.set_cell_value(self.sheet, self.r, self.c, Value::from(v)));
            }
            _ => return Err(RuntimeError::no_member(name)),
        }
        with_mut(&self.app, |st| st.book.set_cell_hyperlink(self.sheet, self.r, self.c, Some(l)))
    }
}
