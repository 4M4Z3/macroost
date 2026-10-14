//! `Range` and `Selection` (one implementation: a selection is the document's live selection
//! span plus the typing and cursor-movement methods), `Font` and `ParagraphFormat`.

use std::cell::RefCell;
use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::constants::builtin_style;
use crate::document::{CharFormat, Marker, ParaFormat};
use crate::find::FindObj;
use crate::objects::{BookmarksObj, StyleObj, UnitsObj, doc_value};
use crate::state::{
    DocEntry, FindState, LiveSpan, Shared, Span, State, TOGGLE, UNDEFINED, Unit, data, move_by, opt_i64, out_of_range,
    pos, req_text, statistic, text, units_in, with, with_mut,
};
use crate::table::{CellsObj, TablesObj};

pub(crate) struct RangeObj {
    pub app: Shared,
    pub doc: u32,
    pub span: LiveSpan,
    pub selection: bool,
    pub find: Rc<RefCell<FindState>>,
}

/// A new Range over `start..end` of document `doc`.
pub(crate) fn range_value(app: &Shared, doc: u32, start: usize, end: usize) -> Result<Value> {
    let span = with_mut(app, |st| st.span(doc, start, end))?;
    Ok(Value::Object(Some(Rc::new(RangeObj {
        app: app.clone(),
        doc,
        span,
        selection: false,
        find: Rc::new(RefCell::new(FindState::default())),
    }))))
}

/// The selection of document `doc`.
pub(crate) fn selection_value(app: &Shared, doc: u32) -> Result<Value> {
    let (span, find) = with(app, |st| {
        let e = st.entry(doc)?;
        Ok((e.selection.clone(), e.selection_find.clone()))
    })?;
    Ok(Value::Object(Some(Rc::new(RangeObj { app: app.clone(), doc, span, selection: true, find }))))
}

/// The document and span of a Range (or Selection) passed as an argument.
pub(crate) fn decode_range(v: &Value) -> Option<(u32, usize, usize)> {
    let Value::Object(Some(o)) = v else { return None };
    let r = o.as_any()?.downcast_ref::<RangeObj>()?;
    let s = *r.span.borrow();
    Some((r.doc, s.start, s.end))
}

pub(crate) fn wrong_args() -> RuntimeError {
    RuntimeError::new(450, "Wrong number of arguments or invalid property assignment")
}

fn long(v: i64) -> Value {
    Value::Long(v.clamp(i64::from(i32::MIN), i64::from(i32::MAX)) as i32)
}

impl RangeObj {
    fn get(&self) -> Span {
        *self.span.borrow()
    }

    fn set(&self, start: usize, end: usize) {
        let mut s = self.span.borrow_mut();
        s.start = start.min(end);
        s.end = end.max(start);
    }

    /// The current span clamped to the story.
    fn bounds(&self, st: &State) -> Result<(usize, usize, usize)> {
        let len = st.doc(self.doc)?.len();
        let s = self.get();
        let end = s.end.min(len);
        Ok((s.start.min(end), end, len))
    }

    fn edit<T>(&self, f: impl FnOnce(&mut DocEntry, usize, usize, usize) -> Result<T>) -> Result<T> {
        with_mut(&self.app, |st| {
            let (a, b, len) = self.bounds(st)?;
            let e = st.entry_mut(self.doc)?;
            f(e, a, b, len)
        })
    }

    fn read<T>(&self, f: impl FnOnce(&DocEntry, usize, usize, usize) -> Result<T>) -> Result<T> {
        with(&self.app, |st| {
            let (a, b, len) = self.bounds(st)?;
            f(st.entry(self.doc)?, a, b, len)
        })
    }

    fn units(&self, unit: Unit, args: &Args) -> Result<Value> {
        let s = self.get();
        let coll: Rc<dyn Object> =
            Rc::new(UnitsObj { app: self.app.clone(), doc: self.doc, unit, start: s.start, end: s.end });
        if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
    }

    fn font(&self) -> Value {
        Value::Object(Some(Rc::new(FontObj {
            app: self.app.clone(),
            doc: self.doc,
            target: FontTarget::Span(self.span.clone()),
        })))
    }

    fn paragraph_format(&self) -> Value {
        Value::Object(Some(Rc::new(ParagraphFormatObj {
            app: self.app.clone(),
            doc: self.doc,
            span: self.span.clone(),
        })))
    }

    /// Where `InsertAfter` puts text: after the range, but before a structure mark or the final
    /// paragraph mark that ends it.
    fn after_point(e: &DocEntry, a: usize, b: usize, len: usize) -> usize {
        if b > a && (b >= len || e.doc.item(b - 1).is_some_and(|i| i.marker != Marker::None)) { b - 1 } else { b }
    }

    fn insert_after(&self, t: &str) -> Result<()> {
        let (p, n) = self.edit(|e, a, b, len| {
            let p = Self::after_point(e, a, b, len);
            Ok((p, e.insert_text(p, t, None)))
        })?;
        let cur = self.get();
        self.set(cur.start, cur.end.max(p + n));
        Ok(())
    }

    fn insert_before(&self, t: &str) -> Result<()> {
        let (p, n) = self.edit(|e, a, _, _| Ok((a, e.insert_text(a, t, None))))?;
        let cur = self.get();
        self.set(p, cur.end.max(p + n));
        Ok(())
    }

    fn set_text(&self, t: &str) -> Result<()> {
        let (s, e) = self.edit(|e, a, b, _| Ok(e.replace(a, b, t, None)))?;
        self.set(s, e);
        Ok(())
    }

    fn move_limit(&self, len: usize) -> usize {
        if self.selection { len.saturating_sub(1) } else { len }
    }

    /// `MoveStart` / `MoveEnd`.
    fn move_edge(&self, args: &Args, end: bool) -> Result<Value> {
        let unit = Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(1))?;
        let count = opt_i64(args, 1, "Count")?.unwrap_or(1);
        let (a, b, np, moved) = self.read(|e, a, b, len| {
            let from = if end { b } else { a };
            let (np, moved) = move_by(e.doc.as_ref(), unit, from, count, len);
            Ok((a, b, np, moved))
        })?;
        if end {
            self.set(a.min(np), np);
        } else {
            self.set(np, b.max(np));
        }
        Ok(long(moved))
    }

    /// `Move`, `MoveRight`, `MoveLeft`: collapse, then move by units. For a selection with
    /// `extend`, only the end (moving right) or the start (moving left) moves.
    fn move_by_units(&self, unit: Unit, count: i64, extend: bool) -> Result<Value> {
        let (a, b, len) = with(&self.app, |st| self.bounds(st))?;
        let limit = self.move_limit(len);
        if extend {
            let (np, moved) = with(&self.app, |st| {
                let from = if count >= 0 { b } else { a };
                Ok(move_by(st.doc(self.doc)?, unit, from, count, len.min(limit + 1)))
            })?;
            if count >= 0 {
                self.set(a, np)
            } else {
                self.set(np, b)
            }
            return Ok(long(moved.abs()));
        }
        let mut count = count;
        let mut moved = 0;
        let mut p = if count >= 0 { b } else { a };
        if a != b && count != 0 {
            // Collapsing a non-empty range counts as one unit.
            moved += 1;
            count -= count.signum();
        }
        if count != 0 {
            let (np, m) = with(&self.app, |st| Ok(move_by(st.doc(self.doc)?, unit, p, count, limit)))?;
            p = np;
            moved += m.abs();
        }
        let p = p.min(limit);
        self.set(p, p);
        Ok(long(moved))
    }

    /// Collapsing to the end of a range that holds the final paragraph mark lands before it.
    fn collapse(&self, direction: i64) -> Result<()> {
        let s = self.get();
        if direction == 0 {
            let len = with(&self.app, |st| Ok(st.doc(self.doc)?.len()))?;
            let p = if s.end >= len && s.end > s.start { len - 1 } else { s.end };
            self.set(p, p);
        } else {
            self.set(s.start, s.start);
        }
        Ok(())
    }

    fn expand(&self, unit: Unit) -> Result<Value> {
        let (na, nb, a, b) = self.read(|e, a, b, _| {
            let us = units_in(e.doc.as_ref(), unit, a, b);
            let na = us.first().map(|u| u.0).unwrap_or(a);
            let nb = us.last().map(|u| u.1).unwrap_or(b);
            Ok((na.min(a), nb.max(b), a, b))
        })?;
        self.set(na, nb);
        Ok(long(((a - na) + (nb - b)) as i64))
    }

    /// `Delete(Unit, Count)`: a non-empty range is deleted (one unit); a collapsed one deletes
    /// `Count` units after it (before it for a negative count). Returns the units deleted.
    fn delete(&self, args: &Args) -> Result<Value> {
        let unit = Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(1))?;
        let count = opt_i64(args, 1, "Count")?.unwrap_or(1);
        let deleted = self.edit(|e, a, b, len| {
            if a < b {
                return Ok(i64::from(e.delete(a, b) > 0));
            }
            let (to, moved) = move_by(e.doc.as_ref(), unit, a, count, len.saturating_sub(1));
            let removed = if to >= a { e.delete(a, to) } else { e.delete(to, a) };
            Ok(if removed > 0 { moved.abs() } else { 0 })
        })?;
        if self.selection {
            let s = self.get();
            self.set(s.start, s.start);
        }
        Ok(long(deleted))
    }

    fn style(&self) -> Result<Value> {
        let name = self.read(|e, a, b, _| {
            let marks = e.para_marks(a, b);
            Ok(marks
                .first()
                .and_then(|m| e.doc.item(*m))
                .map(|i| i.para.style.clone())
                .unwrap_or_else(|| "Normal".into()))
        })?;
        Ok(Value::Object(Some(Rc::new(StyleObj { app: self.app.clone(), doc: self.doc, name }))))
    }

    fn set_style(&self, v: &Value) -> Result<()> {
        let name = style_name(v)?;
        self.edit(|e, a, b, _| {
            e.format_paras(a, b, &|p| p.style = name.clone());
            Ok(())
        })
    }

    // Selection-only members.

    fn type_text(&self, t: &str) -> Result<()> {
        let p = self.edit(|e, a, b, _| {
            if a < b {
                let font = e.doc.item(a).filter(|i| !i.ends_paragraph()).map(|i| i.font.clone());
                let (s, _) = e.replace(a, b, "", None);
                Ok(s + e.insert_text(s, t, font))
            } else {
                Ok(a + e.insert_text(a, t, None))
            }
        })?;
        self.set(p, p);
        Ok(())
    }

    fn home_end(&self, args: &Args, end: bool) -> Result<Value> {
        let unit = Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(5))?;
        let extend = opt_i64(args, 1, "Extend")?.unwrap_or(0) != 0;
        let (a, b, target) = self.read(|e, a, b, len| {
            let target = match unit {
                Unit::Story => {
                    if end && extend {
                        len
                    } else if end {
                        len.saturating_sub(1)
                    } else {
                        0
                    }
                }
                _ => {
                    let at = if end { b } else { a };
                    let us = units_in(e.doc.as_ref(), unit, at, at);
                    match us.first() {
                        Some((s, t)) => {
                            if end {
                                t.saturating_sub(1).max(*s)
                            } else {
                                *s
                            }
                        }
                        None => at,
                    }
                }
            };
            Ok((a, b, target))
        })?;
        let moved = if end { target as i64 - b as i64 } else { a as i64 - target as i64 };
        match (extend, end) {
            (true, true) => self.set(a, target),
            (true, false) => self.set(target, b),
            _ => self.set(target, target),
        }
        Ok(long(moved.max(0)))
    }
}

/// A style named by a string, a built-in index or a Style object.
pub(crate) fn style_name(v: &Value) -> Result<String> {
    if let Value::Object(Some(o)) = v
        && let Some(s) = o.as_any().and_then(|a| a.downcast_ref::<StyleObj>())
    {
        return Ok(s.name.clone());
    }
    let v = data(v)?;
    if v.is_numeric_type() {
        let n = v.to_i64()?;
        return builtin_style(n).map(str::to_owned).ok_or_else(crate::state::no_member_of_collection);
    }
    let name = v.to_text()?;
    crate::objects::canonical_style(&name)
        .ok_or_else(|| RuntimeError::new(5834, "Item with specified name does not exist."))
}

impl Object for RangeObj {
    fn type_name(&self) -> String {
        if self.selection { "Selection".into() } else { "Range".into() }
    }

    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let l = name.to_ascii_lowercase();
        match l.as_str() {
            "" | "text" => self.read(|e, a, b, _| Ok(Value::from(text(e.doc.as_ref(), a, b)))),
            "start" => Ok(long(self.get().start as i64)),
            "end" => Ok(long(self.get().end as i64)),
            "storylength" => with(&self.app, |st| Ok(long(st.doc(self.doc)?.len() as i64))),
            "characters" => self.units(Unit::Character, args),
            "words" => self.units(Unit::Word, args),
            "sentences" => self.units(Unit::Sentence, args),
            "paragraphs" => self.units(Unit::Paragraph, args),
            "font" => Ok(self.font()),
            "paragraphformat" | "format" => Ok(self.paragraph_format()),
            "bold" | "italic" | "underline" | "highlightcolorindex" => {
                let f = FontObj { app: self.app.clone(), doc: self.doc, target: FontTarget::Span(self.span.clone()) };
                let member = if l == "highlightcolorindex" { "Highlight" } else { name };
                f.call(member, args)
            }
            "style" => self.style(),
            "find" => Ok(Value::Object(Some(Rc::new(FindObj {
                app: self.app.clone(),
                doc: self.doc,
                owner: self.span.clone(),
                state: self.find.clone(),
            })))),
            "insertafter" => {
                self.insert_after(&req_text(args, 0, "Text")?)?;
                Ok(Value::Empty)
            }
            "insertbefore" => {
                self.insert_before(&req_text(args, 0, "Text")?)?;
                Ok(Value::Empty)
            }
            "insertparagraphafter" => {
                self.insert_after("\r")?;
                Ok(Value::Empty)
            }
            "insertparagraphbefore" => {
                self.insert_before("\r")?;
                Ok(Value::Empty)
            }
            "insertparagraph" => {
                self.set_text("\r")?;
                Ok(Value::Empty)
            }
            "delete" => self.delete(args),
            "collapse" => {
                self.collapse(opt_i64(args, 0, "Direction")?.unwrap_or(1))?;
                Ok(Value::Empty)
            }
            "move" => {
                let unit = Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(1))?;
                self.move_by_units(unit, opt_i64(args, 1, "Count")?.unwrap_or(1), false)
            }
            "movestart" => self.move_edge(args, false),
            "moveend" => self.move_edge(args, true),
            "expand" => self.expand(Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(2))?),
            "setrange" => {
                let s = pos(crate::state::req_i64(args, 0, "Start")?)?;
                let e = pos(crate::state::req_i64(args, 1, "End")?)?;
                let len = with(&self.app, |st| Ok(st.doc(self.doc)?.len()))?;
                self.set(s.min(len), e.min(len));
                Ok(Value::Empty)
            }
            "duplicate" => {
                let s = self.get();
                range_value(&self.app, self.doc, s.start, s.end)
            }
            "range" => {
                let s = self.get();
                range_value(&self.app, self.doc, s.start, s.end)
            }
            "select" => {
                let s = self.get();
                with_mut(&self.app, |st| {
                    st.active = self.doc;
                    let e = st.entry_mut(self.doc)?;
                    *e.selection.borrow_mut() = s;
                    Ok(())
                })?;
                Ok(Value::Empty)
            }
            "tables" => {
                let s = self.get();
                let coll: Rc<dyn Object> =
                    Rc::new(TablesObj { app: self.app.clone(), doc: self.doc, scope: Some((s.start, s.end)) });
                let v = Value::Object(Some(coll.clone()));
                if args.is_empty() { Ok(v) } else { coll.call("Item", args) }
            }
            "cells" => {
                let s = self.get();
                let coll: Rc<dyn Object> = Rc::new(CellsObj::in_range(&self.app, self.doc, s.start, s.end)?);
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "bookmarks" => {
                let s = self.get();
                let coll: Rc<dyn Object> =
                    Rc::new(BookmarksObj { app: self.app.clone(), doc: self.doc, scope: Some((s.start, s.end)) });
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "computestatistics" => {
                let kind = crate::state::req_i64(args, 0, "Statistic")?;
                self.read(|e, a, b, _| Ok(long(statistic(e.doc.as_ref(), kind, a, b)?)))
            }
            "document" => doc_value(&self.app, self.doc),
            "application" => Ok(crate::objects::app_value(&self.app)),
            "isequal" => {
                let other = decode_range(args.require(0, "Range")?).ok_or_else(RuntimeError::type_mismatch)?;
                let s = self.get();
                Ok(Value::Boolean(other == (self.doc, s.start, s.end)))
            }
            "inrange" => {
                let (d, os, oe) = decode_range(args.require(0, "Range")?).ok_or_else(RuntimeError::type_mismatch)?;
                let s = self.get();
                Ok(Value::Boolean(d == self.doc && os <= s.start && s.end <= oe))
            }
            "information" => {
                let what = crate::state::req_i64(args, 0, "Type")?;
                match what {
                    12 => {
                        let s = self.get();
                        self.read(|e, _, _, _| {
                            Ok(Value::Boolean(e.doc.item(s.start).is_some_and(|i| i.table.is_some())))
                        })
                    }
                    3 | 1 => Ok(Value::Long(1)),
                    _ => Err(RuntimeError::new(4198, "Command failed")),
                }
            }
            // Selection-only members.
            "typetext" if self.selection => {
                self.type_text(&req_text(args, 0, "Text")?)?;
                Ok(Value::Empty)
            }
            "typeparagraph" if self.selection => {
                self.type_text("\r")?;
                Ok(Value::Empty)
            }
            "typebackspace" if self.selection => {
                let p = self.edit(|e, a, b, _| {
                    if a < b {
                        e.delete(a, b);
                        Ok(a)
                    } else if a > 0 && e.delete(a - 1, a) == 1 {
                        Ok(a - 1)
                    } else {
                        Ok(a)
                    }
                })?;
                self.set(p, p);
                Ok(Value::Empty)
            }
            "homekey" if self.selection => self.home_end(args, false),
            "endkey" if self.selection => self.home_end(args, true),
            "moveright" | "moveleft" | "movedown" | "moveup" if self.selection => {
                let default_unit = if l == "moveright" || l == "moveleft" { 1 } else { 5 };
                let unit = Unit::from_constant(opt_i64(args, 0, "Unit")?.unwrap_or(default_unit))?;
                let count = opt_i64(args, 1, "Count")?.unwrap_or(1);
                let extend = opt_i64(args, 2, "Extend")?.unwrap_or(0) != 0;
                let count = if l == "moveleft" || l == "moveup" { -count } else { count };
                self.move_by_units(unit, count, extend)
            }
            "wholestory" if self.selection => {
                let len = with(&self.app, |st| Ok(st.doc(self.doc)?.len()))?;
                self.set(0, len);
                Ok(Value::Empty)
            }
            "type" if self.selection => {
                let s = self.get();
                Ok(Value::Long(if s.start == s.end { 1 } else { 2 }))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        let l = name.to_ascii_lowercase();
        match l.as_str() {
            "" | "text" => self.set_text(&data(&value)?.to_text()?),
            "start" | "end" => {
                let p = pos(data(&value)?.to_i64()?)?;
                let len = with(&self.app, |st| Ok(st.doc(self.doc)?.len()))?;
                if p > len {
                    return Err(out_of_range());
                }
                let s = self.get();
                if l == "start" {
                    self.set(p, s.end.max(p))
                } else {
                    self.set(s.start.min(p), p)
                }
                Ok(())
            }
            "bold" | "italic" | "underline" | "highlightcolorindex" => {
                let f = FontObj { app: self.app.clone(), doc: self.doc, target: FontTarget::Span(self.span.clone()) };
                let member = if l == "highlightcolorindex" { "Highlight" } else { name };
                f.assign(member, args, value)
            }
            "style" => self.set_style(&value),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Font
// ---------------------------------------------------------------------------------------------

pub(crate) enum FontTarget {
    Span(LiveSpan),
    /// The font of a style.
    Style(String),
}

pub(crate) struct FontObj {
    pub app: Shared,
    pub doc: u32,
    pub target: FontTarget,
}

fn tri(v: Option<Option<bool>>) -> Value {
    match v {
        Some(Some(true)) => Value::Long(-1),
        Some(Some(false)) | None => Value::Long(0),
        Some(None) => Value::Long(UNDEFINED),
    }
}

fn num(v: Option<Option<i32>>) -> Value {
    match v {
        Some(Some(n)) => Value::Long(n),
        Some(None) => Value::Long(UNDEFINED),
        None => Value::Long(0),
    }
}

impl FontObj {
    /// The formatting a property reads over: the range's text characters, leaving out
    /// paragraph marks, structure marks and white space when there is anything else.
    fn formats(&self) -> Result<Vec<CharFormat>> {
        with(&self.app, |st| match &self.target {
            FontTarget::Span(s) => {
                let e = st.entry(self.doc)?;
                let len = e.doc.len();
                let s = *s.borrow();
                let (a, b) = (s.start.min(len), s.end.min(len));
                let range = if a < b { a..b } else { a.min(len.saturating_sub(1))..a.min(len.saturating_sub(1)) + 1 };
                let items: Vec<_> = range.filter_map(|p| e.doc.item(p)).collect();
                let text: Vec<CharFormat> = items
                    .iter()
                    .filter(|i| !i.ends_paragraph() && !i.ch.is_whitespace())
                    .map(|i| i.font.clone())
                    .collect();
                if text.is_empty() { Ok(items.iter().map(|i| i.font.clone()).collect()) } else { Ok(text) }
            }
            FontTarget::Style(name) => Ok(vec![crate::objects::style_font(st, self.doc, name)?]),
        })
    }

    fn apply(&self, f: &dyn Fn(&mut CharFormat)) -> Result<()> {
        with_mut(&self.app, |st| match &self.target {
            FontTarget::Span(s) => {
                let s = *s.borrow();
                let e = st.entry_mut(self.doc)?;
                e.format_chars(s.start, s.end, f);
                Ok(())
            }
            FontTarget::Style(name) => {
                let mut font = crate::objects::style_font(st, self.doc, name)?;
                f(&mut font);
                crate::objects::set_style_font(st, self.doc, name, font);
                Ok(())
            }
        })
    }

    fn get_bool(&self, f: impl Fn(&CharFormat) -> bool) -> Result<Value> {
        Ok(tri(crate::state::uniform(self.formats()?.iter().map(f))))
    }

    fn set_bool(&self, value: &Value, get: fn(&CharFormat) -> bool, put: fn(&mut CharFormat, bool)) -> Result<()> {
        let v = data(value)?;
        let on =
            if v.is_numeric_type() && v.to_i64()? == TOGGLE { !self.formats()?.iter().all(get) } else { v.to_bool()? };
        self.apply(&move |c| put(c, on))
    }
}

impl Object for FontObj {
    fn type_name(&self) -> String {
        "Font".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let fm = || self.formats();
        match name.to_ascii_lowercase().as_str() {
            "bold" => self.get_bool(|c| c.bold),
            "italic" => self.get_bool(|c| c.italic),
            "strikethrough" => self.get_bool(|c| c.strike_through),
            "superscript" => self.get_bool(|c| c.superscript),
            "subscript" => self.get_bool(|c| c.subscript),
            "allcaps" => self.get_bool(|c| c.all_caps),
            "smallcaps" => self.get_bool(|c| c.small_caps),
            "hidden" => self.get_bool(|c| c.hidden),
            "underline" => Ok(num(crate::state::uniform(fm()?.iter().map(|c| c.underline)))),
            "color" | "textcolor" => Ok(num(crate::state::uniform(fm()?.iter().map(|c| c.color)))),
            "highlight" => Ok(num(crate::state::uniform(fm()?.iter().map(|c| c.highlight)))),
            "size" => Ok(match crate::state::uniform(fm()?.iter().map(|c| c.size)) {
                Some(Some(s)) => Value::Single(s),
                _ => Value::Single(UNDEFINED as f32),
            }),
            "name" => Ok(match crate::state::uniform(fm()?.iter().map(|c| c.name.clone())) {
                Some(Some(n)) => Value::from(n.as_str()),
                _ => Value::from(""),
            }),
            "reset" => {
                let def = with(&self.app, |st| Ok(st.doc(self.doc)?.default_font()))?;
                self.apply(&move |c| *c = def.clone())?;
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "bold" => self.set_bool(&value, |c| c.bold, |c, v| c.bold = v),
            "italic" => self.set_bool(&value, |c| c.italic, |c, v| c.italic = v),
            "strikethrough" => self.set_bool(&value, |c| c.strike_through, |c, v| c.strike_through = v),
            "superscript" => self.set_bool(
                &value,
                |c| c.superscript,
                |c, v| {
                    c.superscript = v;
                    if v {
                        c.subscript = false;
                    }
                },
            ),
            "subscript" => self.set_bool(
                &value,
                |c| c.subscript,
                |c, v| {
                    c.subscript = v;
                    if v {
                        c.superscript = false;
                    }
                },
            ),
            "allcaps" => self.set_bool(&value, |c| c.all_caps, |c, v| c.all_caps = v),
            "smallcaps" => self.set_bool(&value, |c| c.small_caps, |c, v| c.small_caps = v),
            "hidden" => self.set_bool(&value, |c| c.hidden, |c, v| c.hidden = v),
            "underline" => {
                let n = data(&value)?.to_i32()?;
                self.apply(&move |c| c.underline = n)
            }
            "color" | "textcolor" => {
                let n = data(&value)?.to_i32()?;
                self.apply(&move |c| c.color = n)
            }
            "highlight" => {
                let n = data(&value)?.to_i32()?;
                self.apply(&move |c| c.highlight = n)
            }
            "size" => {
                let s = data(&value)?.to_f64()?;
                if !(1.0..=1638.0).contains(&s) {
                    return Err(RuntimeError::new(5843, "The font size is out of range."));
                }
                // ORACLE: sizes are cut down to half points (11.3 reads back as 11).
                let s = ((s * 2.0).floor() / 2.0) as f32;
                self.apply(&move |c| c.size = s)
            }
            "name" => {
                let n = data(&value)?.to_text()?;
                self.apply(&move |c| c.name = n.clone())
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

// ---------------------------------------------------------------------------------------------
// ParagraphFormat
// ---------------------------------------------------------------------------------------------

pub(crate) struct ParagraphFormatObj {
    pub app: Shared,
    pub doc: u32,
    pub span: LiveSpan,
}

impl ParagraphFormatObj {
    fn formats(&self) -> Result<Vec<ParaFormat>> {
        let s = *self.span.borrow();
        with(&self.app, |st| {
            let e = st.entry(self.doc)?;
            Ok(e.para_marks(s.start, s.end).into_iter().filter_map(|m| e.doc.item(m).map(|i| i.para.clone())).collect())
        })
    }

    fn apply(&self, f: &dyn Fn(&mut ParaFormat)) -> Result<()> {
        let s = *self.span.borrow();
        with_mut(&self.app, |st| {
            st.entry_mut(self.doc)?.format_paras(s.start, s.end, f);
            Ok(())
        })
    }

    fn points(&self, f: impl Fn(&ParaFormat) -> f32) -> Result<Value> {
        Ok(match crate::state::uniform(self.formats()?.iter().map(f)) {
            Some(Some(v)) => Value::Single(v),
            Some(None) => Value::Single(UNDEFINED as f32),
            None => Value::Single(0.0),
        })
    }
}

impl Object for ParagraphFormatObj {
    fn type_name(&self) -> String {
        "ParagraphFormat".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "alignment" => Ok(num(crate::state::uniform(self.formats()?.iter().map(|p| p.alignment)))),
            "leftindent" => self.points(|p| p.left_indent),
            "rightindent" => self.points(|p| p.right_indent),
            "firstlineindent" => self.points(|p| p.first_line_indent),
            "spacebefore" => self.points(|p| p.space_before),
            "spaceafter" => self.points(|p| p.space_after),
            "style" => {
                let name = self.formats()?.first().map(|p| p.style.clone()).unwrap_or_else(|| "Normal".into());
                Ok(Value::Object(Some(Rc::new(StyleObj { app: self.app.clone(), doc: self.doc, name }))))
            }
            "reset" => {
                self.apply(&|p| {
                    let style = std::mem::take(&mut p.style);
                    *p = ParaFormat { style, ..ParaFormat::default() };
                })?;
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let l = name.to_ascii_lowercase();
        if l == "style" {
            let n = style_name(&value)?;
            return self.apply(&move |p| p.style = n.clone());
        }
        let v = data(&value)?;
        match l.as_str() {
            "alignment" => {
                let n = v.to_i32()?;
                if !(0..=9).contains(&n) {
                    return Err(out_of_range());
                }
                self.apply(&move |p| p.alignment = n)
            }
            "leftindent" | "rightindent" | "firstlineindent" | "spacebefore" | "spaceafter" => {
                let x = v.to_f64()? as f32;
                self.apply(&move |p| match l.as_str() {
                    "leftindent" => p.left_indent = x,
                    "rightindent" => p.right_indent = x,
                    "firstlineindent" => p.first_line_indent = x,
                    "spacebefore" => p.space_before = x,
                    _ => p.space_after = x,
                })
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}
