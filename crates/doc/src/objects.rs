//! `Application`, `Documents`, `Document`, the text-unit collections (`Paragraphs`, `Words`,
//! `Characters`, `Sentences`), `Paragraph`, `Styles`/`Style` and `Bookmarks`/`Bookmark`.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::constants::{BUILTIN_STYLES, builtin_style, is_character_style};
use crate::document::CharFormat;
use crate::range::{FontObj, FontTarget, ParagraphFormatObj, decode_range, range_value, selection_value, wrong_args};
use crate::state::{
    LiveSpan, Shared, State, Unit, data, no_member_of_collection, opt_i64, out_of_range, pos, req_i64, req_text,
    statistic, units_in, valid_bookmark_name, with, with_mut,
};
use crate::table::TablesObj;

pub(crate) fn app_value(app: &Shared) -> Value {
    Value::Object(Some(Rc::new(AppObj { app: app.clone() })))
}

/// The Document object of `id`, the same object for as long as a macro holds it.
pub(crate) fn doc_value(app: &Shared, id: u32) -> Result<Value> {
    if let Some(o) = with(app, |st| Ok(st.entry(id)?.object.as_ref().and_then(std::rc::Weak::upgrade)))? {
        return Ok(Value::Object(Some(o)));
    }
    let o: Rc<dyn Object> = Rc::new(DocumentObj { app: app.clone(), id });
    with_mut(app, |st| {
        st.entry_mut(id)?.object = Some(Rc::downgrade(&o));
        Ok(())
    })?;
    Ok(Value::Object(Some(o)))
}

fn object(o: impl Object + 'static) -> Value {
    Value::Object(Some(Rc::new(o)))
}

/// Calls `Item` on a collection when arguments follow its name (`Documents(1)`).
fn coll_or_item(coll: Value, args: &Args) -> Result<Value> {
    if args.is_empty() {
        return Ok(coll);
    }
    match &coll {
        Value::Object(Some(o)) => o.call("Item", args),
        _ => Err(RuntimeError::object_required()),
    }
}

/// A 1-based collection index.
fn index(v: &Value, count: usize) -> Result<usize> {
    let n = data(v)?.to_i64()?;
    if n < 1 || n as u64 > count as u64 {
        return Err(no_member_of_collection());
    }
    Ok(n as usize - 1)
}

// ---------------------------------------------------------------------------------------------
// Application
// ---------------------------------------------------------------------------------------------

pub(crate) struct AppObj {
    pub app: Shared,
}

impl Object for AppObj {
    fn type_name(&self) -> String {
        "Application".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "activedocument" => {
                let id = with(&self.app, State::active_id)?;
                doc_value(&self.app, id)
            }
            "documents" => coll_or_item(object(DocumentsObj { app: self.app.clone() }), args),
            "selection" => {
                let id = with(&self.app, State::active_id)?;
                selection_value(&self.app, id)
            }
            "application" => Ok(app_value(&self.app)),
            "name" => Ok(Value::from("macroost")),
            "version" => Ok(Value::from("16.0")),
            "screenupdating" => with(&self.app, |st| Ok(Value::Boolean(st.screen_updating))),
            "displayalerts" => with(&self.app, |st| Ok(Value::Long(st.display_alerts as i32))),
            "visible" => Ok(Value::Boolean(true)),
            "quit" | "screenrefresh" => Ok(Value::Empty),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let v = data(&value)?;
        match name.to_ascii_lowercase().as_str() {
            "screenupdating" => with_mut(&self.app, |st| {
                st.screen_updating = v.to_bool()?;
                Ok(())
            }),
            "displayalerts" => with_mut(&self.app, |st| {
                st.display_alerts = v.to_i64()?;
                Ok(())
            }),
            "statusbar" | "visible" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Documents / Document
// ---------------------------------------------------------------------------------------------

pub(crate) struct DocumentsObj {
    pub app: Shared,
}

impl DocumentsObj {
    /// Document ids, most recently opened first (the collection's order).
    fn ids(&self) -> Result<Vec<u32>> {
        with(&self.app, |st| Ok(st.docs.iter().rev().map(|d| d.id).collect()))
    }
}

impl Object for DocumentsObj {
    fn type_name(&self) -> String {
        "Documents".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.ids()?.len() as i32)),
            "" | "item" => {
                let key = data(args.require(0, "Index")?)?;
                let ids = self.ids()?;
                let id = if let Value::String(s) = &key {
                    with(&self.app, |st| {
                        st.docs
                            .iter()
                            .find(|d| d.doc.name().eq_ignore_ascii_case(s))
                            .map(|d| d.id)
                            .ok_or_else(no_member_of_collection)
                    })?
                } else {
                    ids[index(&key, ids.len())?]
                };
                doc_value(&self.app, id)
            }
            "add" => {
                let id = with_mut(&self.app, |st| Ok(st.add_untitled()))?;
                doc_value(&self.app, id)
            }
            "open" => Err(RuntimeError::new(5174, "Opening files is not available to macros in this host.")),
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        self.ids()?.into_iter().map(|id| doc_value(&self.app, id)).collect()
    }
}

pub(crate) struct DocumentObj {
    pub app: Shared,
    pub id: u32,
}

impl DocumentObj {
    fn len(&self) -> Result<usize> {
        with(&self.app, |st| Ok(st.doc(self.id)?.len()))
    }

    fn units(&self, unit: Unit, args: &Args) -> Result<Value> {
        let len = self.len()?;
        coll_or_item(object(UnitsObj { app: self.app.clone(), doc: self.id, unit, start: 0, end: len }), args)
    }
}

impl Object for DocumentObj {
    fn type_name(&self) -> String {
        "Document".into()
    }

    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "name" | "fullname" | "" => with(&self.app, |st| Ok(Value::from(st.doc(self.id)?.name()))),
            "path" => Ok(Value::from("")),
            "content" => range_value(&self.app, self.id, 0, self.len()?),
            "range" => {
                let len = self.len()?;
                let s = opt_i64(args, 0, "Start")?.map(pos).transpose()?.unwrap_or(0);
                let e = opt_i64(args, 1, "End")?.map(pos).transpose()?.unwrap_or(len);
                if s > e || e > len {
                    return Err(out_of_range());
                }
                range_value(&self.app, self.id, s, e)
            }
            "paragraphs" => self.units(Unit::Paragraph, args),
            "words" => self.units(Unit::Word, args),
            "characters" => self.units(Unit::Character, args),
            "sentences" => self.units(Unit::Sentence, args),
            "tables" => coll_or_item(object(TablesObj { app: self.app.clone(), doc: self.id, scope: None }), args),
            "bookmarks" => {
                coll_or_item(object(BookmarksObj { app: self.app.clone(), doc: self.id, scope: None }), args)
            }
            "styles" => coll_or_item(object(StylesObj { app: self.app.clone(), doc: self.id }), args),
            "computestatistics" => {
                let kind = req_i64(args, 0, "Statistic")?;
                with(&self.app, |st| {
                    let d = st.doc(self.id)?;
                    Ok(Value::Long(statistic(d, kind, 0, d.len())? as i32))
                })
            }
            "activate" => with_mut(&self.app, |st| {
                st.entry(self.id)?;
                st.active = self.id;
                Ok(Value::Empty)
            }),
            "select" => {
                let len = self.len()?;
                with_mut(&self.app, |st| {
                    st.active = self.id;
                    *st.entry_mut(self.id)?.selection.borrow_mut() = crate::state::Span { start: 0, end: len };
                    Ok(Value::Empty)
                })
            }
            "close" => with_mut(&self.app, |st| {
                st.close(self.id)?;
                Ok(Value::Empty)
            }),
            "save" => with_mut(&self.app, |st| {
                st.entry_mut(self.id)?.doc.save();
                st.log.push("Save".into());
                Ok(Value::Empty)
            }),
            "saved" => Ok(Value::Boolean(true)),
            "undo" | "redo" => Ok(Value::Boolean(false)),
            "application" | "parent" => Ok(app_value(&self.app)),
            "storylength" => Ok(Value::Long(self.len()? as i32)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "saved" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Text units
// ---------------------------------------------------------------------------------------------

/// `Paragraphs`, `Words`, `Characters` or `Sentences` of a span.
pub(crate) struct UnitsObj {
    pub app: Shared,
    pub doc: u32,
    pub unit: Unit,
    pub start: usize,
    pub end: usize,
}

impl UnitsObj {
    fn spans(&self) -> Result<Vec<(usize, usize)>> {
        with(&self.app, |st| Ok(units_in(st.doc(self.doc)?, self.unit, self.start, self.end)))
    }

    fn element(&self, (s, e): (usize, usize)) -> Result<Value> {
        if self.unit == Unit::Paragraph {
            let span = with_mut(&self.app, |st| st.span(self.doc, s, e))?;
            return Ok(object(ParagraphObj { app: self.app.clone(), doc: self.doc, span }));
        }
        range_value(&self.app, self.doc, s, e)
    }
}

impl Object for UnitsObj {
    fn type_name(&self) -> String {
        match self.unit {
            Unit::Paragraph => "Paragraphs",
            Unit::Word => "Words",
            Unit::Sentence => "Sentences",
            _ => "Characters",
        }
        .into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.spans()?.len() as i32)),
            "" | "item" => {
                let spans = self.spans()?;
                let i = index(args.require(0, "Index")?, spans.len())?;
                self.element(spans[i])
            }
            "first" => {
                self.spans()?.first().copied().map_or_else(|| Err(no_member_of_collection()), |s| self.element(s))
            }
            "last" => self.spans()?.last().copied().map_or_else(|| Err(no_member_of_collection()), |s| self.element(s)),
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        self.spans()?.into_iter().map(|s| self.element(s)).collect()
    }
}

pub(crate) struct ParagraphObj {
    pub app: Shared,
    pub doc: u32,
    pub span: LiveSpan,
}

impl ParagraphObj {
    fn format(&self) -> ParagraphFormatObj {
        ParagraphFormatObj { app: self.app.clone(), doc: self.doc, span: self.span.clone() }
    }

    fn neighbour(&self, next: bool) -> Result<Value> {
        let s = *self.span.borrow();
        let all = with(&self.app, |st| Ok(crate::state::units(st.doc(self.doc)?, Unit::Paragraph)))?;
        let i = all.iter().position(|u| u.0 <= s.start && s.start < u.1).ok_or_else(no_member_of_collection)?;
        let j = if next { Some(i + 1) } else { i.checked_sub(1) };
        match j.and_then(|j| all.get(j)) {
            Some(&(a, b)) => {
                let span = with_mut(&self.app, |st| st.span(self.doc, a, b))?;
                Ok(object(ParagraphObj { app: self.app.clone(), doc: self.doc, span }))
            }
            None => Ok(Value::Object(None)),
        }
    }
}

impl Object for ParagraphObj {
    fn type_name(&self) -> String {
        "Paragraph".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let l = name.to_ascii_lowercase();
        match l.as_str() {
            "range" => {
                let s = *self.span.borrow();
                range_value(&self.app, self.doc, s.start, s.end)
            }
            "format" | "paragraphformat" => Ok(object(self.format())),
            "next" => self.neighbour(true),
            "previous" => self.neighbour(false),
            "application" => Ok(app_value(&self.app)),
            "alignment" | "leftindent" | "rightindent" | "firstlineindent" | "spacebefore" | "spaceafter" | "style" => {
                self.format().call(name, args)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "alignment" | "leftindent" | "rightindent" | "firstlineindent" | "spacebefore" | "spaceafter" | "style" => {
                self.format().assign(name, args, value)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------------------------

/// The canonical spelling of a built-in style name, matched case-insensitively.
pub(crate) fn canonical_style(name: &str) -> Option<String> {
    BUILTIN_STYLES.iter().filter_map(|i| builtin_style(*i)).find(|n| n.eq_ignore_ascii_case(name)).map(str::to_owned)
}

/// The font a style gives text: the document defaults unless a macro changed the style.
pub(crate) fn style_font(st: &State, doc: u32, name: &str) -> Result<CharFormat> {
    let e = st.entry(doc)?;
    Ok(e.style_fonts
        .iter()
        .find(|(n, _)| n.eq_ignore_ascii_case(name))
        .map(|(_, f)| f.clone())
        .unwrap_or_else(|| e.doc.default_font()))
}

pub(crate) fn set_style_font(st: &mut State, doc: u32, name: &str, font: CharFormat) {
    if let Ok(e) = st.entry_mut(doc) {
        e.style_fonts.retain(|(n, _)| !n.eq_ignore_ascii_case(name));
        e.style_fonts.push((name.to_owned(), font));
    }
}

pub(crate) struct StylesObj {
    pub app: Shared,
    pub doc: u32,
}

impl Object for StylesObj {
    fn type_name(&self) -> String {
        "Styles".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(BUILTIN_STYLES.len() as i32)),
            "" | "item" => {
                let v = args.require(0, "Index")?;
                let name = crate::range::style_name(v)?;
                Ok(object(StyleObj { app: self.app.clone(), doc: self.doc, name }))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(BUILTIN_STYLES
            .iter()
            .filter_map(|i| builtin_style(*i))
            .map(|n| object(StyleObj { app: self.app.clone(), doc: self.doc, name: n.to_owned() }))
            .collect())
    }
}

pub(crate) struct StyleObj {
    pub app: Shared,
    pub doc: u32,
    pub name: String,
}

impl Object for StyleObj {
    fn type_name(&self) -> String {
        "Style".into()
    }

    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "" | "namelocal" | "name" => Ok(Value::from(self.name.as_str())),
            "font" => Ok(object(FontObj {
                app: self.app.clone(),
                doc: self.doc,
                target: FontTarget::Style(self.name.clone()),
            })),
            "type" => Ok(Value::Long(if is_character_style(&self.name) { 2 } else { 1 })),
            "builtin" => Ok(Value::Boolean(true)),
            "inuse" => Ok(Value::Boolean(true)),
            "basestyle" => Ok(Value::from(if self.name == "Normal" { "" } else { "Normal" })),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Bookmarks
// ---------------------------------------------------------------------------------------------

pub(crate) struct BookmarksObj {
    pub app: Shared,
    pub doc: u32,
    /// For `Range.Bookmarks`: only bookmarks overlapping this span.
    pub scope: Option<(usize, usize)>,
}

impl BookmarksObj {
    /// The bookmarks in the collection, sorted by name.
    fn list(&self) -> Result<Vec<crate::document::Bookmark>> {
        with(&self.app, |st| {
            let mut v: Vec<_> = st
                .doc(self.doc)?
                .bookmarks()
                .iter()
                .filter(|b| match self.scope {
                    None => true,
                    Some((s, e)) => {
                        if b.start == b.end || s == e {
                            b.start <= e && b.end >= s
                        } else {
                            b.start < e && b.end > s
                        }
                    }
                })
                .cloned()
                .collect();
            v.sort_by_key(|b| b.name.to_lowercase());
            Ok(v)
        })
    }

    fn item(&self, key: &Value) -> Result<Value> {
        let key = data(key)?;
        let list = self.list()?;
        let (stored, shown) = if let Value::String(s) = &key {
            let b = list.iter().find(|b| b.name.eq_ignore_ascii_case(s)).ok_or_else(no_member_of_collection)?;
            (b.name.clone(), s.to_string())
        } else {
            let b = &list[index(&key, list.len())?];
            (b.name.clone(), b.name.clone())
        };
        Ok(object(BookmarkObj { app: self.app.clone(), doc: self.doc, name: stored, shown }))
    }
}

impl Object for BookmarksObj {
    fn type_name(&self) -> String {
        "Bookmarks".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.list()?.len() as i32)),
            "" | "item" => self.item(args.require(0, "Index")?),
            "exists" => {
                let n = req_text(args, 0, "Name")?;
                Ok(Value::Boolean(self.list()?.iter().any(|b| b.name.eq_ignore_ascii_case(&n))))
            }
            "add" => {
                let n = req_text(args, 0, "Name")?;
                if !valid_bookmark_name(&n) {
                    return Err(RuntimeError::new(5828, "Bad bookmark name."));
                }
                let (s, e) = match args.get(1, "Range") {
                    Some(v) => {
                        let (d, s, e) = decode_range(v).ok_or_else(RuntimeError::type_mismatch)?;
                        if d != self.doc {
                            return Err(RuntimeError::new(4198, "Command failed"));
                        }
                        (s, e)
                    }
                    None => match self.scope {
                        Some(sc) => sc,
                        None => with(&self.app, |st| {
                            let sel = *st.entry(self.doc)?.selection.borrow();
                            Ok((sel.start, sel.end))
                        })?,
                    },
                };
                with_mut(&self.app, |st| {
                    let bms = st.entry_mut(self.doc)?.doc.bookmarks_mut();
                    bms.retain(|b| !b.name.eq_ignore_ascii_case(&n));
                    bms.push(crate::document::Bookmark { name: n.clone(), start: s, end: e });
                    Ok(())
                })?;
                Ok(object(BookmarkObj { app: self.app.clone(), doc: self.doc, name: n.clone(), shown: n }))
            }
            "defaultsorting" => Ok(Value::Long(0)),
            "showhidden" => Ok(Value::Boolean(false)),
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "defaultsorting" | "showhidden" => Ok(()),
            _ => Err(wrong_args()),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(self
            .list()?
            .into_iter()
            .map(|b| object(BookmarkObj { app: self.app.clone(), doc: self.doc, name: b.name.clone(), shown: b.name }))
            .collect())
    }
}

pub(crate) struct BookmarkObj {
    pub app: Shared,
    pub doc: u32,
    /// The stored name (lookups are case-insensitive).
    pub name: String,
    /// The name as the macro spelled it when it looked the bookmark up.
    pub shown: String,
}

impl BookmarkObj {
    fn span(&self) -> Result<(usize, usize)> {
        with(&self.app, |st| {
            st.doc(self.doc)?
                .bookmarks()
                .iter()
                .find(|b| b.name.eq_ignore_ascii_case(&self.name))
                .map(|b| (b.start, b.end))
                .ok_or_else(no_member_of_collection)
        })
    }
}

impl Object for BookmarkObj {
    fn type_name(&self) -> String {
        "Bookmark".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "" | "name" => Ok(Value::from(self.shown.as_str())),
            "range" => {
                let (s, e) = self.span()?;
                range_value(&self.app, self.doc, s, e)
            }
            "start" => Ok(Value::Long(self.span()?.0 as i32)),
            "end" => Ok(Value::Long(self.span()?.1 as i32)),
            "empty" => {
                let (s, e) = self.span()?;
                Ok(Value::Boolean(s == e))
            }
            "delete" => {
                self.span()?;
                with_mut(&self.app, |st| {
                    st.entry_mut(self.doc)?.doc.bookmarks_mut().retain(|b| !b.name.eq_ignore_ascii_case(&self.name));
                    Ok(Value::Empty)
                })
            }
            "select" => {
                let (s, e) = self.span()?;
                with_mut(&self.app, |st| {
                    st.active = self.doc;
                    *st.entry_mut(self.doc)?.selection.borrow_mut() = crate::state::Span { start: s, end: e };
                    Ok(Value::Empty)
                })
            }
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let l = name.to_ascii_lowercase();
        match l.as_str() {
            "start" | "end" => {
                let p = pos(data(&value)?.to_i64()?)?;
                with_mut(&self.app, |st| {
                    let e = st.entry_mut(self.doc)?;
                    let len = e.doc.len();
                    if p > len {
                        return Err(out_of_range());
                    }
                    let b = e
                        .doc
                        .bookmarks_mut()
                        .iter_mut()
                        .find(|b| b.name.eq_ignore_ascii_case(&self.name))
                        .ok_or_else(no_member_of_collection)?;
                    if l == "start" {
                        b.start = p;
                        b.end = b.end.max(p);
                    } else {
                        b.end = p;
                        b.start = b.start.min(p);
                    }
                    Ok(())
                })
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}
