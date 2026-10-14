//! Application state shared by every object (`Rc<RefCell<State>>`) and the editing rules the
//! object model applies on top of the storage: live ranges that follow edits, bookmarks that
//! shrink and vanish, inherited formatting, text units (words, sentences, paragraphs) and table
//! geometry.

use std::cell::RefCell;
use std::rc::{Rc, Weak};

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::document::{CharFormat, Document, Item, Marker, MemoryDocument, ParaFormat};

pub(crate) type Shared = Rc<RefCell<State>>;

/// A live span of positions. Every Range object (and each document's selection) holds one; the
/// state moves them when the text changes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Span {
    pub start: usize,
    pub end: usize,
}

pub(crate) type LiveSpan = Rc<RefCell<Span>>;

/// Settings of a Find object.
#[derive(Clone, Debug)]
pub(crate) struct FindState {
    pub text: String,
    pub replacement: String,
    pub forward: bool,
    pub wrap: i64,
    pub match_case: bool,
    pub whole_word: bool,
    pub wildcards: bool,
    pub format: bool,
    pub found: bool,
}

impl Default for FindState {
    fn default() -> Self {
        Self {
            text: String::new(),
            replacement: String::new(),
            forward: true,
            wrap: 0,
            match_case: false,
            whole_word: false,
            wildcards: false,
            format: false,
            found: false,
        }
    }
}

pub(crate) struct DocEntry {
    pub id: u32,
    pub doc: Box<dyn Document>,
    pub spans: Vec<Weak<RefCell<Span>>>,
    pub selection: LiveSpan,
    pub selection_find: Rc<RefCell<FindState>>,
    /// The cached Document object, so `ActiveDocument Is Documents(1)` holds.
    pub object: Option<Weak<dyn Object>>,
    /// Style fonts a macro changed (`Styles("Normal").Font.Name = ...`).
    pub style_fonts: Vec<(String, CharFormat)>,
}

pub(crate) struct State {
    /// Open documents in the order they were opened or created.
    pub docs: Vec<DocEntry>,
    pub active: u32,
    /// The document holding the macros (`ThisDocument`).
    pub this_document: u32,
    next_id: u32,
    next_untitled: u32,
    pub log: Vec<String>,
    pub screen_updating: bool,
    pub display_alerts: i64,
}

/// `wdUndefined`: the value a property reads when the range holds mixed values.
pub(crate) const UNDEFINED: i32 = 9_999_999;
/// `wdToggle`.
pub(crate) const TOGGLE: i64 = 9_999_998;

impl State {
    pub fn new(doc: Box<dyn Document>) -> Self {
        let mut st = Self {
            docs: Vec::new(),
            active: 0,
            this_document: 0,
            next_id: 1,
            next_untitled: 1,
            log: Vec::new(),
            screen_updating: true,
            display_alerts: -1,
        };
        let id = st.open(doc);
        st.this_document = id;
        st
    }

    /// Adds a document and makes it active.
    pub fn open(&mut self, doc: Box<dyn Document>) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        self.docs.push(DocEntry {
            id,
            doc,
            spans: Vec::new(),
            selection: Rc::new(RefCell::new(Span { start: 0, end: 0 })),
            selection_find: Rc::new(RefCell::new(FindState::default())),
            object: None,
            style_fonts: Vec::new(),
        });
        self.active = id;
        id
    }

    pub fn add_untitled(&mut self) -> u32 {
        let name = format!("Document{}", self.next_untitled);
        self.next_untitled += 1;
        self.open(Box::new(MemoryDocument::new(&name)))
    }

    pub fn entry(&self, id: u32) -> Result<&DocEntry> {
        self.docs.iter().find(|d| d.id == id).ok_or_else(gone)
    }

    pub fn entry_mut(&mut self, id: u32) -> Result<&mut DocEntry> {
        self.docs.iter_mut().find(|d| d.id == id).ok_or_else(gone)
    }

    pub fn doc(&self, id: u32) -> Result<&dyn Document> {
        Ok(self.entry(id)?.doc.as_ref())
    }

    pub fn active_id(&self) -> Result<u32> {
        if self.docs.iter().any(|d| d.id == self.active) {
            Ok(self.active)
        } else {
            Err(RuntimeError::new(4248, "This command is not available because no document is open."))
        }
    }

    pub fn close(&mut self, id: u32) -> Result<()> {
        let i = self.docs.iter().position(|d| d.id == id).ok_or_else(gone)?;
        self.docs.remove(i);
        if self.active == id {
            self.active = self.docs.last().map(|d| d.id).unwrap_or(0);
        }
        Ok(())
    }

    /// Registers a new live span on document `id`.
    pub fn span(&mut self, id: u32, start: usize, end: usize) -> Result<LiveSpan> {
        let e = self.entry_mut(id)?;
        let s = Rc::new(RefCell::new(Span { start, end }));
        e.spans.retain(|w| w.strong_count() > 0);
        e.spans.push(Rc::downgrade(&s));
        Ok(s)
    }
}

pub(crate) fn gone() -> RuntimeError {
    RuntimeError::new(4605, "This method or property is not available because the document is closed.")
}

pub(crate) fn busy() -> RuntimeError {
    RuntimeError::new(4198, "Command failed")
}

pub(crate) fn out_of_range() -> RuntimeError {
    RuntimeError::new(4608, "Value out of range")
}

pub(crate) fn no_member_of_collection() -> RuntimeError {
    RuntimeError::new(5941, "The requested member of the collection does not exist.")
}

pub(crate) fn with<T>(app: &Shared, f: impl FnOnce(&State) -> Result<T>) -> Result<T> {
    let st = app.try_borrow().map_err(|_| busy())?;
    f(&st)
}

pub(crate) fn with_mut<T>(app: &Shared, f: impl FnOnce(&mut State) -> Result<T>) -> Result<T> {
    let mut st = app.try_borrow_mut().map_err(|_| busy())?;
    f(&mut st)
}

pub(crate) fn data(v: &Value) -> Result<Value> {
    v.clone().deref_default()
}

pub(crate) fn opt_i64(args: &Args, idx: usize, name: &str) -> Result<Option<i64>> {
    args.get(idx, name).map(|v| data(v)?.to_i64()).transpose()
}

pub(crate) fn opt_bool(args: &Args, idx: usize, name: &str) -> Result<Option<bool>> {
    args.get(idx, name).map(|v| data(v)?.to_bool()).transpose()
}

pub(crate) fn opt_text(args: &Args, idx: usize, name: &str) -> Result<Option<String>> {
    args.get(idx, name).map(|v| data(v)?.to_text()).transpose()
}

pub(crate) fn req_text(args: &Args, idx: usize, name: &str) -> Result<String> {
    data(args.require(idx, name)?)?.to_text()
}

pub(crate) fn req_i64(args: &Args, idx: usize, name: &str) -> Result<i64> {
    data(args.require(idx, name)?)?.to_i64()
}

pub(crate) fn pos(v: i64) -> Result<usize> {
    usize::try_from(v).map_err(|_| out_of_range())
}

// ---------------------------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------------------------

/// The characters of the story; structure marks are `'\u{7}'`.
pub(crate) fn chars(doc: &dyn Document) -> Vec<char> {
    (0..doc.len())
        .map(|p| doc.item(p).map(|i| if i.marker == Marker::None { i.ch } else { '\u{7}' }).unwrap_or('\r'))
        .collect()
}

/// The text a macro reads for positions `start..end`.
pub(crate) fn text(doc: &dyn Document, start: usize, end: usize) -> String {
    let mut s = String::new();
    for p in start..end.min(doc.len()) {
        if let Some(i) = doc.item(p) {
            if i.marker == Marker::None {
                s.push(i.ch);
            } else {
                s.push_str(i.text());
            }
        }
    }
    s
}

pub(crate) fn is_terminator(c: char) -> bool {
    c == '\r' || c == '\u{7}'
}

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

/// Text units.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Unit {
    Character,
    Word,
    Sentence,
    Paragraph,
    Story,
    Cell,
    Row,
    Table,
}

impl Unit {
    /// From the `wdUnits` constants.
    pub fn from_constant(n: i64) -> Result<Self> {
        Ok(match n {
            1 => Unit::Character,
            2 => Unit::Word,
            3 => Unit::Sentence,
            4 | 5 => Unit::Paragraph, // wdLine: lines need a layout; a paragraph is one line here
            6 => Unit::Story,
            12 => Unit::Cell,
            10 => Unit::Row,
            15 => Unit::Table,
            _ => return Err(RuntimeError::new(4120, "Bad parameter")),
        })
    }
}

/// The spans of every unit of the story, in order, covering it.
pub(crate) fn units(doc: &dyn Document, unit: Unit) -> Vec<(usize, usize)> {
    let c = chars(doc);
    let n = c.len();
    let mut out = Vec::new();
    match unit {
        Unit::Character => out.extend((0..n).map(|i| (i, i + 1))),
        Unit::Story => out.push((0, n)),
        Unit::Paragraph => {
            let mut start = 0;
            for (i, ch) in c.iter().enumerate() {
                if is_terminator(*ch) {
                    out.push((start, i + 1));
                    start = i + 1;
                }
            }
            if start < n {
                out.push((start, n));
            }
        }
        Unit::Word => {
            let mut i = 0;
            while i < n {
                let start = i;
                let ch = c[i];
                if is_terminator(ch) || ch == '\t' {
                    i += 1;
                    if ch == '\t' {
                        while i < n && c[i] == ' ' {
                            i += 1;
                        }
                    }
                } else if is_word_char(ch) {
                    while i < n
                        && (is_word_char(c[i])
                            || ((c[i] == '\'' || c[i] == '\u{2019}') && i + 1 < n && is_word_char(c[i + 1])))
                    {
                        i += 1;
                    }
                    while i < n && c[i] == ' ' {
                        i += 1;
                    }
                } else if ch == ' ' {
                    while i < n && c[i] == ' ' {
                        i += 1;
                    }
                } else {
                    // A run of punctuation is one word ("...", "?!"); other characters (a
                    // manual line break) stand alone.
                    i += 1;
                    while ch != '\u{b}' && i < n && !is_word_char(c[i]) && !c[i].is_whitespace() && !is_terminator(c[i])
                    {
                        i += 1;
                    }
                    while i < n && c[i] == ' ' {
                        i += 1;
                    }
                }
                out.push((start, i));
            }
        }
        Unit::Sentence => {
            let (mut i, mut start) = (0, 0);
            while i < n {
                let ch = c[i];
                if is_terminator(ch) {
                    i += 1;
                    // An end-of-row mark belongs to the sentence of the row's last cell.
                    let row_end = doc.item(i - 1).is_some_and(|it| it.marker == Marker::RowEnd);
                    match out.last_mut() {
                        Some(last) if row_end && start == i - 1 && last.1 == start => last.1 = i,
                        _ => out.push((start, i)),
                    }
                    start = i;
                    continue;
                }
                if matches!(ch, '.' | '!' | '?') {
                    let mut j = i;
                    while j < n && matches!(c[j], '.' | '!' | '?') {
                        j += 1;
                    }
                    while j < n && matches!(c[j], '"' | '\'' | ')' | ']' | '\u{2019}' | '\u{201D}') {
                        j += 1;
                    }
                    if j < n && c[j] == ' ' {
                        while j < n && c[j] == ' ' {
                            j += 1;
                        }
                        if j < n && is_terminator(c[j]) {
                            j += 1;
                        }
                        out.push((start, j));
                        start = j;
                    }
                    i = j;
                    continue;
                }
                i += 1;
            }
            if start < n {
                out.push((start, n));
            }
        }
        Unit::Cell | Unit::Row | Unit::Table => {
            for t in tables(doc) {
                match unit {
                    Unit::Table => out.push((t.start, t.end)),
                    Unit::Row => out.extend(t.rows.iter().map(|r| (r.start, r.end))),
                    _ => out.extend(t.rows.iter().flat_map(|r| r.cells.iter().copied())),
                }
            }
        }
    }
    out
}

/// Units overlapping `start..end`; for a collapsed span, the unit containing it.
pub(crate) fn units_in(doc: &dyn Document, unit: Unit, start: usize, end: usize) -> Vec<(usize, usize)> {
    let all = units(doc, unit);
    if start == end {
        let p = if start >= doc.len() { doc.len().saturating_sub(1) } else { start };
        return all.into_iter().filter(|(s, e)| *s <= p && p < *e).take(1).collect();
    }
    all.into_iter().filter(|(s, e)| *s < end && *e > start).collect()
}

/// Word-count statistics (`wdStatistic...`) over `start..end`.
pub(crate) fn statistic(doc: &dyn Document, kind: i64, start: usize, end: usize) -> Result<i64> {
    let c = chars(doc);
    let slice = &c[start.min(c.len())..end.min(c.len())];
    let is_space = |ch: char| ch.is_whitespace() || is_terminator(ch) || ch == '\u{b}' || ch == '\u{a0}';
    let tokens = || slice.split(|ch| is_space(*ch)).filter(|t| !t.is_empty());
    let paragraphs =
        || slice.split(|ch| is_terminator(*ch)).filter(|p| p.iter().any(|ch| !is_space(*ch))).count() as i64;
    Ok(match kind {
        0 => tokens().filter(|t| t.iter().any(|ch| ch.is_alphanumeric())).count() as i64,
        1 => paragraphs(),
        2 => 1,
        3 => slice.iter().filter(|ch| !is_space(**ch)).count() as i64,
        4 => paragraphs(),
        5 => slice.iter().filter(|ch| !is_terminator(**ch)).count() as i64,
        6 => 0,
        _ => return Err(RuntimeError::new(4120, "Bad parameter")),
    })
}

// ---------------------------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------------------------

#[derive(Clone, Debug)]
pub(crate) struct RowGeom {
    pub start: usize,
    /// After the end-of-row mark.
    pub end: usize,
    /// Each cell's span, including its end-of-cell mark.
    pub cells: Vec<(usize, usize)>,
}

#[derive(Clone, Debug)]
pub(crate) struct TableGeom {
    pub id: u32,
    pub start: usize,
    pub end: usize,
    pub rows: Vec<RowGeom>,
}

impl TableGeom {
    pub fn columns(&self) -> usize {
        self.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    }
}

/// The tables of the story in order.
pub(crate) fn tables(doc: &dyn Document) -> Vec<TableGeom> {
    let mut out: Vec<TableGeom> = Vec::new();
    let mut p = 0;
    let n = doc.len();
    while p < n {
        let Some(id) = doc.item(p).and_then(|i| i.table) else {
            p += 1;
            continue;
        };
        let start = p;
        let mut rows = Vec::new();
        let (mut row_start, mut cell_start, mut cells) = (p, p, Vec::new());
        while p < n {
            let Some(item) = doc.item(p) else { break };
            if item.table != Some(id) {
                break;
            }
            match item.marker {
                Marker::CellEnd => {
                    cells.push((cell_start, p + 1));
                    cell_start = p + 1;
                }
                Marker::RowEnd => {
                    rows.push(RowGeom { start: row_start, end: p + 1, cells: std::mem::take(&mut cells) });
                    row_start = p + 1;
                    cell_start = p + 1;
                }
                Marker::None => {}
            }
            p += 1;
        }
        out.push(TableGeom { id, start, end: p, rows });
    }
    out
}

pub(crate) fn table_by_id(doc: &dyn Document, id: u32) -> Result<TableGeom> {
    tables(doc).into_iter().find(|t| t.id == id).ok_or_else(no_member_of_collection)
}

// ---------------------------------------------------------------------------------------------
// Editing
// ---------------------------------------------------------------------------------------------

impl DocEntry {
    fn live(&self) -> Vec<LiveSpan> {
        let mut v: Vec<LiveSpan> = self.spans.iter().filter_map(Weak::upgrade).collect();
        v.push(self.selection.clone());
        v
    }

    /// Formatting for text inserted at `p`: the character before it, unless that ends a
    /// paragraph, else the item at `p`.
    pub fn inherited_font(&self, p: usize) -> CharFormat {
        if p > 0
            && let Some(prev) = self.doc.item(p - 1)
            && !prev.ends_paragraph()
        {
            return prev.font.clone();
        }
        self.doc.item(p).map(|i| i.font.clone()).unwrap_or_else(|| self.doc.default_font())
    }

    /// The paragraph format a new paragraph mark at `p` takes: that of the paragraph it splits.
    fn para_at(&self, p: usize) -> ParaFormat {
        (p..self.doc.len())
            .filter_map(|q| self.doc.item(q))
            .find(|i| i.ends_paragraph())
            .map(|i| if i.marker == Marker::None { i.para.clone() } else { ParaFormat::default() })
            .unwrap_or_default()
    }

    /// Items for `text` inserted at `p`.
    pub fn text_items(&self, p: usize, text: &str, font: Option<CharFormat>) -> Vec<Item> {
        let font = font.unwrap_or_else(|| self.inherited_font(p));
        let para = self.para_at(p);
        let table = self.doc.item(p).and_then(|i| i.table);
        let normalized = text.replace("\r\n", "\r").replace('\n', "\r");
        normalized
            .chars()
            .map(|ch| {
                let mut it = Item::char(ch, font.clone());
                it.table = table;
                if ch == '\r' {
                    it.para = para.clone();
                }
                it
            })
            .collect()
    }

    /// Inserts `items` at `p`, moving live spans and bookmarks. Returns the count inserted.
    pub fn insert_items(&mut self, p: usize, items: Vec<Item>) -> usize {
        let n = items.len();
        if n == 0 {
            return 0;
        }
        let p = p.min(self.doc.len().saturating_sub(1));
        self.doc.splice(p, p, items);
        // Positions after the insertion point move; a span starting or ending exactly there
        // stays put (so a range ending at `p` does not grow, and one starting there does).
        let shift = |x: &mut usize| {
            if *x > p {
                *x += n;
            }
        };
        for s in self.live() {
            let mut s = s.borrow_mut();
            shift(&mut s.start);
            shift(&mut s.end);
        }
        for b in self.doc.bookmarks_mut() {
            shift(&mut b.start);
            shift(&mut b.end);
        }
        n
    }

    pub fn insert_text(&mut self, p: usize, text: &str, font: Option<CharFormat>) -> usize {
        let items = self.text_items(p, text, font);
        let splits_at_end = items.iter().any(|i| i.ch == '\r')
            && self.doc.item(p).is_some_and(|i| i.ch == '\r' && i.marker == Marker::None);
        let n = self.insert_items(p, items);
        // A paragraph ended right before its mark continues in the style that follows a
        // heading or title (Normal).
        if splits_at_end
            && let Some(mark) = self.doc.item_mut(p + n)
            && next_style(&mark.para.style).is_some()
        {
            mark.para.style = next_style(&mark.para.style).unwrap_or("Normal").to_owned();
        }
        n
    }

    /// Removes the positions `a..b` that may be deleted: never structure marks, never the final
    /// paragraph mark. Returns how many were removed.
    pub fn delete(&mut self, a: usize, b: usize) -> usize {
        let len = self.doc.len();
        let b = b.min(len);
        let last = len.saturating_sub(1);
        let mut removed = 0;
        let mut q = b;
        while q > a {
            // Find the next run (from the right) of deletable positions.
            let deletable =
                |pos: usize, doc: &dyn Document| pos != last && doc.item(pos).is_some_and(|i| i.marker == Marker::None);
            if !deletable(q - 1, self.doc.as_ref()) {
                q -= 1;
                continue;
            }
            let run_end = q;
            while q > a && deletable(q - 1, self.doc.as_ref()) {
                q -= 1;
            }
            removed += run_end - q;
            self.remove_raw(q, run_end);
        }
        removed
    }

    /// Removes positions `a..b` unconditionally (table rows, whole tables).
    pub fn remove_raw(&mut self, a: usize, b: usize) {
        if b <= a {
            return;
        }
        let n = b - a;
        self.doc.splice(a, b, Vec::new());
        let map = |x: usize| {
            if x >= b {
                x - n
            } else if x > a {
                a
            } else {
                x
            }
        };
        for s in self.live() {
            let mut s = s.borrow_mut();
            s.start = map(s.start);
            s.end = map(s.end);
        }
        self.doc.bookmarks_mut().retain_mut(|bm| {
            let whole = bm.start < bm.end && a <= bm.start && bm.end <= b;
            bm.start = map(bm.start);
            bm.end = map(bm.end);
            !whole
        });
    }

    /// Replaces `a..b` with `text`; returns the new span of the text (plus the final paragraph
    /// mark when the replaced span held it, since that mark stays).
    pub fn replace(&mut self, a: usize, b: usize, text: &str, font: Option<CharFormat>) -> (usize, usize) {
        let len = self.doc.len();
        let held_final = b >= len;
        let font = font.or_else(|| {
            if a < b { self.doc.item(a).filter(|i| !i.ends_paragraph()).map(|i| i.font.clone()) } else { None }
        });
        self.delete(a, b);
        let p = a.min(self.doc.len().saturating_sub(1));
        let n = self.insert_text(p, text, font);
        (p, p + n + usize::from(held_final))
    }

    /// Sets the character formatting of `a..b` through `f`.
    pub fn format_chars(&mut self, a: usize, b: usize, f: &dyn Fn(&mut CharFormat)) {
        for p in a..b.min(self.doc.len()) {
            if let Some(i) = self.doc.item_mut(p) {
                f(&mut i.font);
            }
        }
    }

    /// The paragraph marks of the paragraphs overlapping `a..b`.
    pub fn para_marks(&self, a: usize, b: usize) -> Vec<usize> {
        units_in(self.doc.as_ref(), Unit::Paragraph, a, b).into_iter().map(|(_, e)| e - 1).collect()
    }

    pub fn format_paras(&mut self, a: usize, b: usize, f: &dyn Fn(&mut ParaFormat)) {
        for m in self.para_marks(a, b) {
            if let Some(i) = self.doc.item_mut(m) {
                f(&mut i.para);
            }
        }
    }
}

/// Moves `pos` by `count` units; returns the new position and the units actually moved.
pub(crate) fn move_by(doc: &dyn Document, unit: Unit, pos: usize, count: i64, limit: usize) -> (usize, i64) {
    let len = doc.len();
    if unit == Unit::Character {
        let target = (pos as i64 + count).clamp(0, limit as i64) as usize;
        return (target, target as i64 - pos as i64);
    }
    let mut bounds: Vec<usize> = units(doc, unit).into_iter().map(|(s, _)| s).collect();
    bounds.push(len);
    bounds.retain(|b| *b <= limit);
    bounds.dedup();
    let (mut p, mut moved) = (pos, 0i64);
    if count > 0 {
        for _ in 0..count {
            match bounds.iter().find(|b| **b > p) {
                Some(b) => {
                    p = *b;
                    moved += 1;
                }
                None => break,
            }
        }
    } else {
        for _ in 0..(-count) {
            match bounds.iter().rev().find(|b| **b < p) {
                Some(b) => {
                    p = *b;
                    moved -= 1;
                }
                None => break,
            }
        }
    }
    (p, moved)
}

/// The style of the paragraph after one in `style` when it is ended at its end.
fn next_style(style: &str) -> Option<&'static str> {
    let heading = style.strip_prefix("Heading ").is_some_and(|n| n.len() == 1 && n.as_bytes()[0].is_ascii_digit());
    (heading || matches!(style, "Title" | "Subtitle" | "Caption")).then_some("Normal")
}

/// Bookmark names: a letter first, then letters, digits and underscores, at most 40.
pub(crate) fn valid_bookmark_name(name: &str) -> bool {
    let mut c = name.chars();
    c.next().is_some_and(char::is_alphabetic)
        && c.all(|ch| ch.is_alphanumeric() || ch == '_')
        && name.chars().count() <= 40
}

/// Values of a property over several positions: one value, or mixed.
pub(crate) fn uniform<T: PartialEq + Clone>(values: impl IntoIterator<Item = T>) -> Option<Option<T>> {
    let mut it = values.into_iter();
    let first = it.next()?;
    if it.all(|v| v == first) { Some(Some(first)) } else { Some(None) }
}
