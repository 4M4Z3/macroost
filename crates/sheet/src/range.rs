//! `Range` and the format objects hanging off it (`Interior`, `Font`, `Borders`, `Border`,
//! `Areas`). A range is a descriptor (sheet, areas, kind); nothing is allocated per cell.

use std::cell::RefCell;
use std::rc::Rc;

use macroost_lang::{Args, Array, Object, Result, RuntimeError, Value};

use crate::address::{MAX_COLS, MAX_ROWS, Rect, format_area, format_area_r1c1, parse_reference, quote_sheet};
use crate::format::{PALETTE, nearest_color_index};
use crate::formula::{RefEdit, a1_to_r1c1, r1c1_to_a1, shift_formula};
use crate::objects::sheet_value;
use crate::state::{
    Input, LARGE, PasteKind, Shared, State, column_arg, constant_formula_text, data, err1004, opt_bool, opt_i64,
    parse_typed, with, with_mut,
};
use crate::workbook::{BorderLine, CellFormat, HAlign, SheetId, ShiftDirection};

/// What `Count`, `Item` and `For Each` walk: cells, rows or columns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Kind {
    Cells,
    Rows,
    Columns,
}

pub(crate) struct RangeObj {
    pub app: Shared,
    pub sheet: SheetId,
    pub areas: Vec<Rect>,
    pub kind: Kind,
}

/// Hidden member through which objects recognize a range passed as an argument.
pub(crate) const RANGE_REF: &str = "_macroostRangeRef";

pub(crate) fn range_value(app: &Shared, sheet: SheetId, areas: Vec<Rect>, kind: Kind) -> Value {
    let areas = if areas.is_empty() { vec![Rect::cell(1, 1)] } else { areas };
    Value::Object(Some(Rc::new(RangeObj { app: app.clone(), sheet, areas, kind })))
}

/// The (sheet, areas) of a Range object argument, if it is one of ours.
pub(crate) fn decode_range(v: &Value) -> Option<(SheetId, Vec<Rect>)> {
    let Value::Object(Some(o)) = v else { return None };
    let Ok(Value::String(s)) = o.call(RANGE_REF, &Args::none()) else { return None };
    let (sheet, rest) = s.split_once('|')?;
    let sheet: SheetId = sheet.parse().ok()?;
    let mut areas = Vec::new();
    for part in rest.split(';') {
        let n: Vec<u32> = part.split(',').filter_map(|x| x.parse().ok()).collect();
        if n.len() != 4 {
            return None;
        }
        areas.push(Rect::new(n[0], n[1], n[2], n[3]));
    }
    (!areas.is_empty()).then_some((sheet, areas))
}

/// Resolves address text against `default_sheet`; a qualifier names another sheet.
pub(crate) fn resolve_address(st: &State, default_sheet: SheetId, text: &str) -> Result<(SheetId, Vec<Rect>)> {
    let Some(p) = parse_reference(text) else {
        return resolve_name(st, default_sheet, text)
            .ok_or_else(|| err1004("Method 'Range' of object '_Global' failed"));
    };
    let sheet = match &p.sheet {
        Some(name) => st.sheet_by_name(name).ok_or_else(|| err1004("Method 'Range' of object '_Global' failed"))?,
        None => default_sheet,
    };
    Ok((sheet, p.areas))
}

/// The refers-to text of defined name `name` as seen from `sheet`: a sheet-scoped name wins over
/// a workbook-scoped one.
pub(crate) fn lookup_name(st: &State, sheet: SheetId, name: &str) -> Option<String> {
    let names = st.book.defined_names();
    let find = |scope: Option<SheetId>| {
        names.iter().find(|(n, s, _)| *s == scope && n.eq_ignore_ascii_case(name)).map(|(_, _, r)| r.clone())
    };
    find(Some(sheet)).or_else(|| find(None))
}

/// A defined name (`Total`, `Sheet1!Total`) whose refers-to text is a reference.
pub(crate) fn resolve_name(st: &State, default_sheet: SheetId, text: &str) -> Option<(SheetId, Vec<Rect>)> {
    let t = text.trim();
    let (scope_sheet, name) = match t.rsplit_once('!') {
        Some((s, n)) => {
            let s = s.trim_matches('\'').replace("''", "'");
            (st.sheet_by_name(&s)?, n)
        }
        None => (default_sheet, t),
    };
    if name.is_empty() {
        return None;
    }
    let refers = lookup_name(st, scope_sheet, name)?;
    let p = parse_reference(refers.trim_start_matches('='))?;
    let sheet = match &p.sheet {
        Some(n) => st.sheet_by_name(n)?,
        None => scope_sheet,
    };
    Some((sheet, p.areas))
}

/// A range argument given as a Range object or as address text.
pub(crate) fn range_arg(app: &Shared, default_sheet: SheetId, v: &Value) -> Result<(SheetId, Vec<Rect>)> {
    if let Some(r) = decode_range(v) {
        return Ok(r);
    }
    match v {
        Value::String(s) => with(app, |st| resolve_address(st, default_sheet, s)),
        Value::Object(None) => Err(RuntimeError::object_not_set()),
        _ => Err(RuntimeError::type_mismatch()),
    }
}

fn mixed<T: Into<Value>>(v: Option<T>) -> Value {
    v.map(Into::into).unwrap_or(Value::Null)
}

/// Common value of `f` over every cell of `areas`, or `None` when cells differ.
pub(crate) fn fold_format<T: PartialEq>(
    st: &State,
    sheet: SheetId,
    areas: &[Rect],
    f: impl Fn(&CellFormat) -> T,
) -> Option<T> {
    let mut acc: Option<T> = None;
    let mut push = |v: T| -> bool {
        match &acc {
            Some(a) if *a != v => false,
            Some(_) => true,
            None => {
                acc = Some(v);
                true
            }
        }
    };
    for a in areas {
        if a.area() <= LARGE {
            for r in a.r1..=a.r2 {
                for c in a.c1..=a.c2 {
                    if !push(f(&st.book.cell_format(sheet, r, c))) {
                        return None;
                    }
                }
            }
        } else {
            let occ = st.book.occupied_cells(sheet, *a);
            if (occ.len() as u64) < a.area() && !push(f(&CellFormat::default())) {
                return None;
            }
            for (r, c) in occ {
                if !push(f(&st.book.cell_format(sheet, r, c))) {
                    return None;
                }
            }
        }
    }
    acc
}

/// Applies `f` to the format of every cell of `areas` (huge areas: stored cells only).
pub(crate) fn update_format(
    st: &mut State,
    sheet: SheetId,
    areas: &[Rect],
    f: impl Fn(&mut CellFormat, u32, u32, &Rect),
) -> Result<()> {
    st.check_format(sheet)?;
    for a in areas {
        // Whole columns or rows keep a format of their own for cells not stored yet.
        if a.area() > LARGE {
            if a.is_entire_columns() {
                for c in a.c1..=a.c2 {
                    let mut cf = st.book.column_format(sheet, c).unwrap_or_default();
                    f(&mut cf, 1, c, a);
                    st.book.set_column_format(sheet, c, Some(cf));
                }
            } else if a.is_entire_rows() {
                for r in a.r1..=a.r2 {
                    let mut cf = st.book.row_format(sheet, r).unwrap_or_default();
                    f(&mut cf, r, 1, a);
                    st.book.set_row_format(sheet, r, Some(cf));
                }
            }
        }
        for (r, c) in st.write_cells(sheet, *a) {
            let mut cf = st.book.cell_format(sheet, r, c);
            let before = cf.clone();
            f(&mut cf, r, c, a);
            if cf != before {
                st.book.set_cell_format(sheet, r, c, cf)?;
            }
        }
    }
    st.saved = false;
    Ok(())
}

pub(crate) fn halign_constant(h: HAlign) -> i32 {
    match h {
        HAlign::General => 1,
        HAlign::Left => -4131,
        HAlign::Center => -4108,
        HAlign::Right => -4152,
        HAlign::Fill => 5,
        HAlign::Justify => -4130,
        HAlign::CenterAcrossSelection => 7,
        HAlign::Distributed => -4117,
    }
}

pub(crate) fn halign_from(n: i64) -> Result<HAlign> {
    Ok(match n {
        1 => HAlign::General,
        -4131 => HAlign::Left,
        -4108 => HAlign::Center,
        -4152 => HAlign::Right,
        5 => HAlign::Fill,
        -4130 => HAlign::Justify,
        7 => HAlign::CenterAcrossSelection,
        -4117 => HAlign::Distributed,
        _ => return Err(err1004("Unable to set the HorizontalAlignment property of the Range class")),
    })
}

/// Turns a 1-based (row, col) pair that may be out of the sheet into a cell rect.
fn cell_at(r: i64, c: i64) -> Result<Rect> {
    if r < 1 || c < 1 || r > i64::from(MAX_ROWS) || c > i64::from(MAX_COLS) {
        return Err(err1004("Application-defined or object-defined error"));
    }
    Ok(Rect::cell(r as u32, c as u32))
}

/// the reference spreadsheet's Ctrl+arrow rule from (r, c). `dir`: xlUp -4162, xlDown -4121, xlToLeft -4159,
/// xlToRight -4161.
fn end_from(st: &State, sheet: SheetId, r: u32, c: u32, dir: i64) -> Result<(u32, u32)> {
    let filled = |r: u32, c: u32| !st.is_blank(sheet, r, c);
    Ok(match dir {
        -4121 => {
            if r == MAX_ROWS {
                return Ok((r, c));
            }
            if filled(r, c) && filled(r + 1, c) {
                let mut x = r + 1;
                while x < MAX_ROWS && filled(x + 1, c) {
                    x += 1;
                }
                (x, c)
            } else {
                let next = st.content_cells(sheet, Rect::new(r + 1, c, MAX_ROWS, c));
                (next.first().map_or(MAX_ROWS, |p| p.0), c)
            }
        }
        -4162 => {
            if r == 1 {
                return Ok((r, c));
            }
            if filled(r, c) && filled(r - 1, c) {
                let mut x = r - 1;
                while x > 1 && filled(x - 1, c) {
                    x -= 1;
                }
                (x, c)
            } else {
                let prev = st.content_cells(sheet, Rect::new(1, c, r - 1, c));
                (prev.last().map_or(1, |p| p.0), c)
            }
        }
        -4161 => {
            if c == MAX_COLS {
                return Ok((r, c));
            }
            if filled(r, c) && filled(r, c + 1) {
                let mut x = c + 1;
                while x < MAX_COLS && filled(r, x + 1) {
                    x += 1;
                }
                (r, x)
            } else {
                let next = st.content_cells(sheet, Rect::new(r, c + 1, r, MAX_COLS));
                (r, next.first().map_or(MAX_COLS, |p| p.1))
            }
        }
        -4159 => {
            if c == 1 {
                return Ok((r, c));
            }
            if filled(r, c) && filled(r, c - 1) {
                let mut x = c - 1;
                while x > 1 && filled(r, x - 1) {
                    x -= 1;
                }
                (r, x)
            } else {
                let prev = st.content_cells(sheet, Rect::new(r, 1, r, c - 1));
                (r, prev.last().map_or(1, |p| p.1))
            }
        }
        _ => return Err(err1004("Application-defined or object-defined error")),
    })
}

pub(crate) fn current_region_of(st: &State, sheet: SheetId, start: Rect) -> Rect {
    current_region(st, sheet, start)
}

/// `CurrentRegion`: grow until bordered by empty rows and columns (diagonals count).
fn current_region(st: &State, sheet: SheetId, start: Rect) -> Rect {
    let mut r = start;
    loop {
        let mut grown = r;
        let outer = Rect::new(
            r.r1.saturating_sub(1).max(1),
            r.c1.saturating_sub(1).max(1),
            (r.r2 + 1).min(MAX_ROWS),
            (r.c2 + 1).min(MAX_COLS),
        );
        for (cr, cc) in st.content_cells(sheet, outer) {
            if !r.contains(cr, cc) {
                grown = grown.bounding(&Rect::cell(cr, cc));
            }
        }
        if grown == r {
            return r;
        }
        r = grown;
    }
}

/// Removes `cut` from every piece, splitting a piece into the parts left, right, above and below
/// the cut (in that order, in place), the way the application lists blank and visible cells.
fn subtract(pieces: &mut Vec<Rect>, cut: &Rect) {
    let mut out = Vec::with_capacity(pieces.len() + 4);
    for p in pieces.drain(..) {
        let Some(x) = p.intersect(cut) else {
            out.push(p);
            continue;
        };
        if x.c1 > p.c1 {
            out.push(Rect::new(p.r1, p.c1, p.r2, x.c1 - 1));
        }
        if x.c2 < p.c2 {
            out.push(Rect::new(p.r1, x.c2 + 1, p.r2, p.c2));
        }
        if x.r1 > p.r1 {
            out.push(Rect::new(p.r1, x.c1, x.r1 - 1, x.c2));
        }
        if x.r2 < p.r2 {
            out.push(Rect::new(x.r2 + 1, x.c1, p.r2, x.c2));
        }
    }
    *pieces = out;
}

/// Replaces every match of `what` (wildcards allowed) inside `text`; `None` if nothing matched.
fn replace_part(text: &str, what: &str, repl: &str, match_case: bool) -> Option<String> {
    if what.is_empty() {
        return None;
    }
    let ch: Vec<char> = text.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut hit = false;
    while i < ch.len() {
        let mut best = None;
        for end in (i + 1..=ch.len()).rev() {
            let piece: String = ch[i..end].iter().collect();
            if wildcard_match(what, &piece, true, match_case) {
                best = Some(end);
                break;
            }
        }
        match best {
            Some(end) => {
                out.push_str(repl);
                i = end;
                hit = true;
            }
            None => {
                out.push(ch[i]);
                i += 1;
            }
        }
    }
    hit.then_some(out)
}

/// Merges cells (sorted row-major) into rectangles: runs per row, then identical runs stacked.
fn cells_to_areas(cells: &[(u32, u32)]) -> Vec<Rect> {
    let mut runs: Vec<Rect> = Vec::new();
    for &(r, c) in cells {
        match runs.last_mut() {
            Some(last) if last.r1 == r && last.c2 + 1 == c => last.c2 = c,
            _ => runs.push(Rect::cell(r, c)),
        }
    }
    let mut out: Vec<Rect> = Vec::new();
    for run in runs {
        if let Some(prev) = out.iter_mut().find(|p| p.c1 == run.c1 && p.c2 == run.c2 && p.r2 + 1 == run.r1) {
            prev.r2 = run.r2;
        } else {
            out.push(run);
        }
    }
    out
}

/// Wildcard match (`*`, `?`, `~` escape), as `Find`, `Match` and `CountIf` use.
pub(crate) fn wildcard_match(pattern: &str, text: &str, whole: bool, match_case: bool) -> bool {
    let norm = |s: &str| if match_case { s.to_owned() } else { s.to_lowercase() };
    let p: Vec<char> = norm(pattern).chars().collect();
    let t: Vec<char> = norm(text).chars().collect();
    // Tokens: Some(c) literal, None-star, Any-question.
    #[derive(Clone, Copy, PartialEq)]
    enum Tk {
        Lit(char),
        Star,
        Any,
    }
    let mut toks = Vec::new();
    let mut i = 0;
    while i < p.len() {
        match p[i] {
            '~' if i + 1 < p.len() => {
                toks.push(Tk::Lit(p[i + 1]));
                i += 2;
                continue;
            }
            '*' => toks.push(Tk::Star),
            '?' => toks.push(Tk::Any),
            ch => toks.push(Tk::Lit(ch)),
        }
        i += 1;
    }
    if !whole {
        toks.insert(0, Tk::Star);
        toks.push(Tk::Star);
    }
    // Classic DP over tokens x text.
    let n = t.len();
    let mut dp = vec![false; n + 1];
    dp[0] = true;
    for tk in &toks {
        let mut next = vec![false; n + 1];
        match tk {
            Tk::Star => {
                let mut any = false;
                for j in 0..=n {
                    any |= dp[j];
                    next[j] = any;
                }
            }
            Tk::Any => {
                next[1..(n + 1)].copy_from_slice(&dp[..n]);
            }
            Tk::Lit(ch) => {
                for j in 0..n {
                    next[j + 1] = dp[j] && t[j] == *ch;
                }
            }
        }
        dp = next;
    }
    dp[n]
}

impl RangeObj {
    fn a0(&self) -> Rect {
        self.areas.first().copied().unwrap_or(Rect::cell(1, 1))
    }

    fn with_areas(&self, areas: Vec<Rect>, kind: Kind) -> Value {
        range_value(&self.app, self.sheet, areas, kind)
    }

    fn check_sheet(&self) -> Result<()> {
        with(&self.app, |st| if st.sheet_exists(self.sheet) { Ok(()) } else { Err(RuntimeError::object_required()) })
    }

    fn count(&self) -> u64 {
        self.areas
            .iter()
            .map(|a| match self.kind {
                Kind::Cells => a.area(),
                Kind::Rows => u64::from(a.rows()),
                Kind::Columns => u64::from(a.cols()),
            })
            .sum()
    }

    fn get_value(&self, value2: bool) -> Result<Value> {
        let a = self.a0();
        with(&self.app, |st| {
            if a.is_cell() {
                return Ok(st.read_value(self.sheet, a.r1, a.c1, value2));
            }
            let mut arr = Array::new(vec![(1, i64::from(a.rows())), (1, i64::from(a.cols()))])?;
            for (r, c) in st.book.occupied_cells(self.sheet, a) {
                let v = st.read_value(self.sheet, r, c, value2);
                arr.set(&[i64::from(r - a.r1 + 1), i64::from(c - a.c1 + 1)], v)?;
            }
            Ok(Value::Array(Rc::new(RefCell::new(arr))))
        })
    }

    /// Per-cell text getter over the first area: a scalar for one cell, else a 2-D array.
    fn get_text_grid(&self, f: impl Fn(&State, u32, u32) -> String) -> Result<Value> {
        let a = self.a0();
        with(&self.app, |st| {
            if a.is_cell() {
                return Ok(Value::from(f(st, a.r1, a.c1)));
            }
            let mut arr = Array::new(vec![(1, i64::from(a.rows())), (1, i64::from(a.cols()))])?;
            for v in arr.data.iter_mut() {
                *v = Value::from("");
            }
            for (r, c) in st.book.occupied_cells(self.sheet, a) {
                arr.set(&[i64::from(r - a.r1 + 1), i64::from(c - a.c1 + 1)], Value::from(f(st, r, c)))?;
            }
            Ok(Value::Array(Rc::new(RefCell::new(arr))))
        })
    }

    fn formula_of(st: &State, sheet: SheetId, r: u32, c: u32, r1c1: bool) -> String {
        match st.book.cell_formula(sheet, r, c) {
            Some(f) if r1c1 => a1_to_r1c1(&f, r, c),
            Some(f) => f,
            None => constant_formula_text(&st.book.cell_value(sheet, r, c), &st.book.cell_format(sheet, r, c)),
        }
    }

    /// `Value = v` / `Formula = v` / `FormulaR1C1 = v` over every area.
    fn set_content(&self, value: Value, mode: SetMode) -> Result<()> {
        // A multi-cell Range object given as the value (not its .Value) clears the target.
        let value = match decode_range(&value) {
            Some((_, areas)) if areas.iter().map(Rect::area).sum::<u64>() > 1 => Value::Empty,
            _ => data(&value)?,
        };
        with_mut(&self.app, |st| {
            for a in &self.areas {
                self.set_area(st, *a, &value, mode)?;
            }
            st.saved = false;
            Ok(())
        })
    }

    fn set_area(&self, st: &mut State, a: Rect, value: &Value, mode: SetMode) -> Result<()> {
        let sheet = self.sheet;
        let write = |st: &mut State, r: u32, c: u32, v: &Value| -> Result<()> {
            st.check_cell(sheet, r, c)?;
            match mode {
                SetMode::Value => st.write_value(sheet, r, c, v),
                SetMode::Formula | SetMode::FormulaR1C1 => {
                    let v = data(v)?;
                    match &v {
                        Value::String(s) if s.starts_with('=') && s.len() > 1 => {
                            let f = if mode == SetMode::FormulaR1C1 { r1c1_to_a1(s, r, c) } else { s.to_string() };
                            st.apply_input(sheet, r, c, Input::Formula(f))
                        }
                        other => st.write_value(sheet, r, c, other),
                    }
                }
            }
        };
        if let Value::Array(arr) = value {
            let arr = arr.borrow().clone();
            let (rows, cols, two_d) = match arr.bounds.as_slice() {
                [(lo, hi)] => (1, hi - lo + 1, false),
                [(l1, h1), (l2, h2)] => (h1 - l1 + 1, h2 - l2 + 1, true),
                _ => return Err(RuntimeError::type_mismatch()),
            };
            let limit = if a.area() > LARGE {
                Rect::new(
                    a.r1,
                    a.c1,
                    (a.r1 as i64 + rows.max(1) - 1).min(a.r2 as i64) as u32,
                    (a.c1 as i64 + cols.max(1) - 1).min(a.c2 as i64) as u32,
                )
            } else {
                a
            };
            for r in limit.r1..=limit.r2 {
                for c in limit.c1..=limit.c2 {
                    let i = i64::from(r - a.r1);
                    let j = i64::from(c - a.c1);
                    let v = if two_d {
                        if i < rows && j < cols {
                            arr.get(&[arr.bounds[0].0 + i, arr.bounds[1].0 + j])?
                        } else {
                            Value::Error(2042)
                        }
                    } else if j < cols {
                        arr.get(&[arr.bounds[0].0 + j])?
                    } else {
                        Value::Error(2042)
                    };
                    write(st, r, c, &v)?;
                }
            }
            return Ok(());
        }
        let clears = match value {
            Value::Empty | Value::Null => true,
            Value::String(s) => s.is_empty(),
            _ => false,
        };
        if clears {
            for (r, c) in st.content_cells(sheet, a) {
                st.check_cell(sheet, r, c)?;
                st.book.set_cell_value(sheet, r, c, Value::Empty)?;
            }
            return Ok(());
        }
        let targets = st.write_cells(sheet, a);
        // A formula assigned to a block is filled: relative references move with each cell.
        let formula_text = match (mode, value) {
            (SetMode::Formula, Value::String(s)) if s.starts_with('=') && s.len() > 1 => Some(s.to_string()),
            (SetMode::Value, Value::String(s)) if matches!(parse_typed(s), Input::Formula(_)) => Some(s.to_string()),
            _ => None,
        };
        for (r, c) in targets {
            let text_cell = st.book.cell_format(sheet, r, c).number_format.as_deref() == Some("@");
            match &formula_text {
                Some(f) if !text_cell => {
                    st.check_cell(sheet, r, c)?;
                    let moved = shift_formula(f, i64::from(r - a.r1), i64::from(c - a.c1));
                    st.apply_input(sheet, r, c, Input::Formula(moved))?;
                }
                _ => write(st, r, c, value)?,
            }
        }
        Ok(())
    }

    /// `Item(row, col)` / `Item(n)` / `rng(n)`.
    fn item(&self, args: &Args) -> Result<Value> {
        let a = self.a0();
        let row_arg = args.get(0, "RowIndex").cloned();
        let col_arg = args.get(1, "ColumnIndex").cloned();
        match self.kind {
            Kind::Cells => match (row_arg, col_arg) {
                (Some(r), Some(c)) => {
                    let r = data(&r)?.to_i64()?;
                    let c = column_arg(&data(&c)?)?;
                    let cell = cell_at(i64::from(a.r1) + r - 1, i64::from(a.c1) + c - 1)?;
                    Ok(self.with_areas(vec![cell], Kind::Cells))
                }
                (Some(n), None) => {
                    let n = data(&n)?;
                    if let Value::String(s) = &n {
                        return self.relative_range(s);
                    }
                    let n = n.to_i64()?;
                    if n < 1 {
                        return Err(err1004("Application-defined or object-defined error"));
                    }
                    let w = i64::from(a.cols());
                    let idx = n - 1;
                    let cell = cell_at(i64::from(a.r1) + idx / w, i64::from(a.c1) + idx % w)?;
                    Ok(self.with_areas(vec![cell], Kind::Cells))
                }
                (None, Some(c)) => {
                    let c = column_arg(&data(&c)?)?;
                    let cell = cell_at(i64::from(a.r1), i64::from(a.c1) + c - 1)?;
                    Ok(self.with_areas(vec![cell], Kind::Cells))
                }
                (None, None) => Err(RuntimeError::standard(449)),
            },
            Kind::Rows => {
                let n = data(&row_arg.ok_or_else(|| RuntimeError::standard(449))?)?;
                if let Value::String(s) = &n {
                    let p = parse_reference(s).or_else(|| parse_reference(&format!("{s}:{s}")));
                    let Some(p) = p else { return Err(err1004("Application-defined or object-defined error")) };
                    let rows = p.areas[0];
                    let r1 = i64::from(a.r1) + i64::from(rows.r1) - 1;
                    let r2 = i64::from(a.r1) + i64::from(rows.r2) - 1;
                    cell_at(r2, 1)?;
                    return Ok(self.with_areas(vec![Rect::new(r1 as u32, a.c1, r2 as u32, a.c2)], Kind::Rows));
                }
                let r = i64::from(a.r1) + n.to_i64()? - 1;
                cell_at(r, i64::from(a.c1))?;
                Ok(self.with_areas(vec![Rect::new(r as u32, a.c1, r as u32, a.c2)], Kind::Rows))
            }
            Kind::Columns => {
                let n = data(&row_arg.ok_or_else(|| RuntimeError::standard(449))?)?;
                if let Value::String(s) = &n
                    && s.contains(':')
                {
                    let p = parse_reference(s).ok_or_else(|| err1004("Application-defined or object-defined error"))?;
                    let cols = p.areas[0];
                    let c1 = i64::from(a.c1) + i64::from(cols.c1) - 1;
                    let c2 = i64::from(a.c1) + i64::from(cols.c2) - 1;
                    cell_at(1, c2)?;
                    return Ok(self.with_areas(vec![Rect::new(a.r1, c1 as u32, a.r2, c2 as u32)], Kind::Columns));
                }
                let c = i64::from(a.c1) + column_arg(&n)? - 1;
                cell_at(i64::from(a.r1), c)?;
                Ok(self.with_areas(vec![Rect::new(a.r1, c as u32, a.r2, c as u32)], Kind::Columns))
            }
        }
    }

    /// `rng.Range("B2")`: an address relative to the range's top-left cell.
    fn relative_range(&self, text: &str) -> Result<Value> {
        let a = self.a0();
        let (sheet, areas) = with(&self.app, |st| resolve_address(st, self.sheet, text))?;
        if sheet != self.sheet {
            return Err(err1004("Method 'Range' of object '_Worksheet' failed"));
        }
        if parse_reference(text).is_none() {
            // A defined name refers to fixed cells, not to an offset from this range.
            return Ok(self.with_areas(areas, Kind::Cells));
        }
        let mut out = Vec::new();
        for r in areas {
            let moved = if a.r1 == 1 && a.c1 == 1 {
                Some(r)
            } else {
                // Whole rows/columns stay whole in the other dimension.
                let dr = if r.is_entire_columns() { 0 } else { i64::from(a.r1) - 1 };
                let dc = if r.is_entire_rows() { 0 } else { i64::from(a.c1) - 1 };
                r.offset(dr, dc)
            };
            out.push(moved.ok_or_else(|| err1004("Application-defined or object-defined error"))?);
        }
        Ok(self.with_areas(out, Kind::Cells))
    }

    fn range_method(&self, args: &Args) -> Result<Value> {
        let c1 = args.require(0, "Cell1")?.clone();
        let c2 = args.get(1, "Cell2").cloned();
        let first = match &c1 {
            Value::String(s) => match decode_range(&self.relative_range(s)?) {
                Some(r) => r,
                None => return Err(err1004("Method 'Range' of object '_Worksheet' failed")),
            },
            other => range_arg(&self.app, self.sheet, other)?,
        };
        let Some(c2) = c2 else {
            return Ok(self.with_areas(first.1, Kind::Cells));
        };
        let second = match &c2 {
            Value::String(s) => match decode_range(&self.relative_range(s)?) {
                Some(r) => r,
                None => return Err(err1004("Method 'Range' of object '_Worksheet' failed")),
            },
            other => range_arg(&self.app, self.sheet, other)?,
        };
        if first.0 != second.0 || first.0 != self.sheet {
            return Err(err1004("Method 'Range' of object '_Worksheet' failed"));
        }
        let b = first.1.iter().chain(second.1.iter()).fold(first.1[0], |acc, r| acc.bounding(r));
        Ok(self.with_areas(vec![b], Kind::Cells))
    }

    fn address(&self, args: &Args) -> Result<String> {
        let row_abs = opt_bool(args, 0, "RowAbsolute")?.unwrap_or(true);
        let col_abs = opt_bool(args, 1, "ColumnAbsolute")?.unwrap_or(true);
        let style = opt_i64(args, 2, "ReferenceStyle")?.unwrap_or(1);
        let external = opt_bool(args, 3, "External")?.unwrap_or(false);
        let origin = match args.get(4, "RelativeTo") {
            Some(v) => decode_range(v).map(|(_, a)| (a[0].r1, a[0].c1)).unwrap_or((1, 1)),
            None => (1, 1),
        };
        let parts: Vec<String> = self
            .areas
            .iter()
            .map(|a| {
                if style == -4150 {
                    format_area_r1c1(a, row_abs, col_abs, origin)
                } else {
                    format_area(a, row_abs, col_abs)
                }
            })
            .collect();
        let text = parts.join(",");
        if external {
            let (book, sheet) = with(&self.app, |st| Ok((st.book.name(), st.sheet_name(self.sheet)?)))?;
            let q = quote_sheet(&sheet);
            let prefix = match q.strip_prefix('\'') {
                Some(rest) => format!("'[{book}]{rest}!"),
                None => format!("[{book}]{q}!"),
            };
            return Ok(format!("{prefix}{text}"));
        }
        Ok(text)
    }

    fn offset(&self, args: &Args) -> Result<Value> {
        let dr = opt_i64(args, 0, "RowOffset")?.unwrap_or(0);
        let dc = opt_i64(args, 1, "ColumnOffset")?.unwrap_or(0);
        let mut out = Vec::new();
        for a in &self.areas {
            out.push(a.offset(dr, dc).ok_or_else(|| err1004("Application-defined or object-defined error"))?);
        }
        Ok(self.with_areas(out, self.kind))
    }

    fn resize(&self, args: &Args) -> Result<Value> {
        let a = self.a0();
        let rows = opt_i64(args, 0, "RowSize")?.unwrap_or(i64::from(a.rows()));
        let cols = opt_i64(args, 1, "ColumnSize")?.unwrap_or(i64::from(a.cols()));
        if rows < 1 || cols < 1 {
            return Err(err1004("Application-defined or object-defined error"));
        }
        let r2 = i64::from(a.r1) + rows - 1;
        let c2 = i64::from(a.c1) + cols - 1;
        cell_at(r2, c2)?;
        Ok(self.with_areas(vec![Rect::new(a.r1, a.c1, r2 as u32, c2 as u32)], self.kind))
    }

    fn select(&self) -> Result<()> {
        with_mut(&self.app, |st| {
            if st.active_sheet != self.sheet {
                return Err(err1004("Select method of Range class failed"));
            }
            let a = self.a0();
            let view = st.views.entry(self.sheet).or_default();
            view.selection = self.areas.clone();
            view.active = (a.r1, a.c1);
            Ok(())
        })
    }

    fn activate(&self) -> Result<()> {
        let a = self.a0();
        let inside = with(&self.app, |st| {
            if st.active_sheet != self.sheet {
                return Err(err1004("Activate method of Range class failed"));
            }
            let view = st.view(self.sheet);
            Ok(self.areas.len() == 1 && a.is_cell() && view.selection.iter().any(|s| s.contains(a.r1, a.c1)))
        })?;
        if inside {
            return with_mut(&self.app, |st| {
                st.views.entry(self.sheet).or_default().active = (a.r1, a.c1);
                Ok(())
            });
        }
        self.select()
    }

    fn copy(&self, args: &Args, cut: bool) -> Result<Value> {
        let a = self.a0();
        if self.areas.len() > 1 && cut {
            return Err(err1004("This action won't work on multiple selections."));
        }
        let dest = match args.get(0, "Destination") {
            Some(d) => Some(range_arg(&self.app, self.sheet, d)?),
            None => None,
        };
        with_mut(&self.app, |st| {
            let clip = st.snapshot(self.sheet, a, cut);
            match dest {
                Some((ds, dareas)) => {
                    if cut {
                        st.check_writable(self.sheet)?;
                        st.clear_contents(self.sheet, a)?;
                        st.clear_formats(self.sheet, a)?;
                        crate::notes::clear_comments_in(st, self.sheet, a)?;
                    }
                    let pasted = st.paste(&clip, ds, dareas[0], PasteKind::All, false, false, !cut, -4142)?;
                    if cut {
                        let (from, to) = (st.sheet_name(self.sheet)?, st.sheet_name(ds)?);
                        st.apply_ref_edit(&RefEdit::Move {
                            sheet: from,
                            src: a,
                            dr: i64::from(pasted.r1) - i64::from(a.r1),
                            dc: i64::from(pasted.c1) - i64::from(a.c1),
                            to_sheet: to,
                        })?;
                    }
                }
                None => {
                    st.clipboard = Some(clip);
                }
            }
            Ok(Value::Boolean(true))
        })
    }

    fn paste_special(&self, args: &Args) -> Result<Value> {
        let kind = PasteKind::from_constant(opt_i64(args, 0, "Paste")?.unwrap_or(-4104))?;
        let operation = opt_i64(args, 1, "Operation")?.unwrap_or(-4142);
        let skip = opt_bool(args, 2, "SkipBlanks")?.unwrap_or(false);
        let transpose = opt_bool(args, 3, "Transpose")?.unwrap_or(false);
        let a = self.a0();
        with_mut(&self.app, |st| {
            let clip = st.clipboard.clone().ok_or_else(|| err1004("PasteSpecial method of Range class failed"))?;
            if clip.cut {
                return Err(err1004("PasteSpecial method of Range class failed"));
            }
            st.paste(&clip, self.sheet, a, kind, skip, transpose, true, operation)?;
            Ok(Value::Boolean(true))
        })
    }

    fn insert(&self, args: &Args) -> Result<Value> {
        let shift = opt_i64(args, 0, "Shift")?;
        let origin = opt_i64(args, 1, "CopyOrigin")?.unwrap_or(0);
        let sheet = self.sheet;
        with_mut(&self.app, |st| {
            st.check_writable(sheet)?;
            let clip = st.clipboard.take();
            let mut areas = self.areas.clone();
            areas.sort_by_key(|x| std::cmp::Reverse((x.r1, x.c1)));
            for a in areas {
                // Inserting copied cells: the inserted block takes the clipboard's size.
                let a = match &clip {
                    Some(c) => {
                        let rows = a.r1 + c.src.rows() - 1;
                        let (c1, c2) =
                            if c.src.is_entire_rows() { (1, MAX_COLS) } else { (a.c1, a.c1 + c.src.cols() - 1) };
                        let (r1, r2) = if c.src.is_entire_columns() { (1, MAX_ROWS) } else { (a.r1, rows) };
                        if r2 > MAX_ROWS || c2 > MAX_COLS {
                            return Err(err1004("Insert method of Range class failed"));
                        }
                        Rect::new(r1, c1, r2, c2)
                    }
                    None => a,
                };
                if let Some(c) = &clip
                    && c.cut
                {
                    st.clear_contents(c.sheet, c.src)?;
                    st.clear_formats(c.sheet, c.src)?;
                }
                let name = st.sheet_name(sheet)?;
                let down = if a.is_entire_rows() {
                    st.book.insert_rows(sheet, a.r1, a.rows())?;
                    st.apply_ref_edit(&RefEdit::Rows { sheet: name, at: a.r1, count: a.rows(), insert: true })?;
                    true
                } else if a.is_entire_columns() {
                    st.book.insert_columns(sheet, a.c1, a.cols())?;
                    st.apply_ref_edit(&RefEdit::Cols { sheet: name, at: a.c1, count: a.cols(), insert: true })?;
                    false
                } else {
                    let down = match shift {
                        Some(-4121) => true,
                        Some(-4161) => false,
                        Some(_) => return Err(err1004("Insert method of Range class failed")),
                        None => a.cols() >= a.rows(),
                    };
                    let dir = if down { ShiftDirection::Down } else { ShiftDirection::Right };
                    st.book.insert_cells(sheet, a, dir)?;
                    st.apply_ref_edit(&if down {
                        RefEdit::CellsVert { sheet: name, c1: a.c1, c2: a.c2, at: a.r1, count: a.rows(), insert: true }
                    } else {
                        RefEdit::CellsHorz { sheet: name, r1: a.r1, r2: a.r2, at: a.c1, count: a.cols(), insert: true }
                    })?;
                    down
                };
                match &clip {
                    Some(c) => {
                        st.paste(c, sheet, Rect::cell(a.r1, a.c1), PasteKind::All, false, false, !c.cut, -4142)?;
                    }
                    None => inherit_format(st, sheet, a, down, origin)?,
                }
            }
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    fn delete(&self, args: &Args) -> Result<Value> {
        let shift = opt_i64(args, 0, "Shift")?;
        let sheet = self.sheet;
        with_mut(&self.app, |st| {
            st.check_writable(sheet)?;
            let mut areas = self.areas.clone();
            areas.sort_by_key(|x| std::cmp::Reverse((x.r1, x.c1)));
            for a in areas {
                let name = st.sheet_name(sheet)?;
                if a.is_entire_rows() && a.is_entire_columns() {
                    st.book.delete_rows(sheet, 1, MAX_ROWS)?;
                    st.book.delete_columns(sheet, 1, MAX_COLS)?;
                } else if a.is_entire_rows() {
                    st.book.delete_rows(sheet, a.r1, a.rows())?;
                    st.apply_ref_edit(&RefEdit::Rows { sheet: name, at: a.r1, count: a.rows(), insert: false })?;
                } else if a.is_entire_columns() {
                    st.book.delete_columns(sheet, a.c1, a.cols())?;
                    st.apply_ref_edit(&RefEdit::Cols { sheet: name, at: a.c1, count: a.cols(), insert: false })?;
                } else {
                    let up = match shift {
                        Some(-4162) => true,
                        Some(-4159) => false,
                        Some(_) => return Err(err1004("Delete method of Range class failed")),
                        None => a.cols() >= a.rows(),
                    };
                    st.book.delete_cells(sheet, a, if up { ShiftDirection::Up } else { ShiftDirection::Left })?;
                    st.apply_ref_edit(&if up {
                        RefEdit::CellsVert { sheet: name, c1: a.c1, c2: a.c2, at: a.r1, count: a.rows(), insert: false }
                    } else {
                        RefEdit::CellsHorz { sheet: name, r1: a.r1, r2: a.r2, at: a.c1, count: a.cols(), insert: false }
                    })?;
                }
            }
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    fn clear(&self, contents: bool, formats: bool) -> Result<Value> {
        with_mut(&self.app, |st| {
            if formats {
                st.check_format(self.sheet)?;
            }
            for a in &self.areas {
                if contents {
                    // Every locked cell counts, filled or not.
                    if a.area() <= LARGE {
                        for r in a.r1..=a.r2 {
                            for c in a.c1..=a.c2 {
                                st.check_cell(self.sheet, r, c)?;
                            }
                        }
                    } else {
                        st.check_writable(self.sheet)?;
                    }
                    st.clear_contents(self.sheet, *a)?;
                }
                if formats {
                    st.clear_formats(self.sheet, *a)?;
                }
                if contents && formats {
                    crate::notes::clear_comments_in(st, self.sheet, *a)?;
                    for (r, c) in st.book.occupied_cells(self.sheet, *a) {
                        if st.book.cell_hyperlink(self.sheet, r, c).is_some() {
                            st.book.set_cell_hyperlink(self.sheet, r, c, None)?;
                        }
                    }
                }
            }
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    fn find(&self, args: &Args) -> Result<Value> {
        let what = data(args.require(0, "What")?)?.to_text()?;
        let after = match args.get(1, "After") {
            Some(v) => decode_range(v).map(|(_, a)| (a[0].r1, a[0].c1)),
            None => None,
        };
        // LookIn, LookAt and SearchOrder persist between Find calls.
        let last = with(&self.app, |st| Ok(st.last_find.clone()))?;
        let look_in =
            opt_i64(args, 2, "LookIn")?.map(|v| v as i32).or(last.as_ref().map(|p| p.look_in)).unwrap_or(-4123);
        let look_at = opt_i64(args, 3, "LookAt")?.map(|v| v as i32).or(last.as_ref().map(|p| p.look_at)).unwrap_or(2);
        let by_columns =
            opt_i64(args, 4, "SearchOrder")?.map(|v| v == 2).or(last.as_ref().map(|p| p.by_columns)).unwrap_or(false);
        let previous = opt_i64(args, 5, "SearchDirection")?.unwrap_or(1) == 2;
        let match_case = opt_bool(args, 6, "MatchCase")?.unwrap_or(false);
        let params = crate::state::FindParams { what, look_in, look_at, match_case, by_columns };
        with_mut(&self.app, |st| {
            st.last_find = Some(params.clone());
            Ok(())
        })?;
        self.search(&params, after, previous)
    }

    fn find_next(&self, args: &Args, previous: bool) -> Result<Value> {
        let after = match args.get(0, "After") {
            Some(v) => decode_range(v).map(|(_, a)| (a[0].r1, a[0].c1)),
            None => None,
        };
        let Some(params) = with(&self.app, |st| Ok(st.last_find.clone()))? else {
            return Ok(Value::Object(None));
        };
        self.search(&params, after, previous)
    }

    fn search(&self, p: &crate::state::FindParams, after: Option<(u32, u32)>, previous: bool) -> Result<Value> {
        let a0 = self.a0();
        let found = with(&self.app, |st| {
            let mut cells: Vec<(u32, u32)> = Vec::new();
            for a in &self.areas {
                cells.extend(st.content_cells(self.sheet, *a));
            }
            let key = |&(r, c): &(u32, u32)| if p.by_columns { (c, r) } else { (r, c) };
            cells.sort_by_key(key);
            cells.dedup();
            let after = after.unwrap_or((a0.r1, a0.c1));
            let ak = key(&after);
            let order: Vec<(u32, u32)> = if previous {
                let (before, rest): (Vec<_>, Vec<_>) = cells.iter().partition(|c| key(c) < ak);
                before.into_iter().rev().chain(rest.into_iter().rev()).collect()
            } else {
                let (rest, upto): (Vec<_>, Vec<_>) = cells.iter().partition(|c| key(c) > ak);
                rest.into_iter().chain(upto).collect()
            };
            let whole = p.look_at == 1;
            for (r, c) in order {
                let text = match p.look_in {
                    -4163 => st.text(self.sheet, r, c),
                    -4144 => continue, // xlComments: no comments are stored
                    _ => match st.book.cell_formula(self.sheet, r, c) {
                        Some(f) => f,
                        None => constant_formula_text(
                            &st.book.cell_value(self.sheet, r, c),
                            &st.book.cell_format(self.sheet, r, c),
                        ),
                    },
                };
                if wildcard_match(&p.what, &text, whole, p.match_case) {
                    return Ok(Some((r, c)));
                }
            }
            Ok(None)
        })?;
        Ok(match found {
            Some((r, c)) => self.with_areas(vec![Rect::cell(r, c)], Kind::Cells),
            None => Value::Object(None),
        })
    }

    fn special_cells(&self, args: &Args) -> Result<Value> {
        let ty = opt_i64(args, 0, "Type")?.ok_or_else(|| RuntimeError::standard(449))?;
        let filter = opt_i64(args, 1, "Value")?.unwrap_or(23);
        let found = with(&self.app, |st| {
            let last = st.book.last_cell(self.sheet).unwrap_or(Rect::cell(1, 1));
            let used =
                Rect::new(1, 1, last.r2, last.c2).intersect(&st.used(self.sheet).bounding(&last)).unwrap_or(last);
            // A single cell searches the whole used range.
            let scope: Vec<Rect> = if self.areas.len() == 1 && self.a0().is_cell() {
                vec![used]
            } else {
                self.areas.iter().filter_map(|a| a.intersect(&used)).collect()
            };
            let value_kind = |v: &Value| -> i64 {
                match v {
                    Value::Boolean(_) => 4,
                    Value::Error(_) => 16,
                    Value::String(_) => 2,
                    _ => 1,
                }
            };
            match ty {
                11 => Ok(vec![Rect::cell(last.r2, last.c2)]),
                2 | -4123 => {
                    let mut cells: Vec<(u32, u32)> = Vec::new();
                    for a in &scope {
                        for (r, c) in st.content_cells(self.sheet, *a) {
                            let is_formula = st.book.cell_formula(self.sheet, r, c).is_some();
                            if (ty == -4123) == is_formula
                                && value_kind(&st.book.cell_value(self.sheet, r, c)) & filter != 0
                            {
                                cells.push((r, c));
                            }
                        }
                    }
                    cells.sort_unstable();
                    cells.dedup();
                    Ok(cells_to_areas(&cells))
                }
                4 => {
                    // Blanks: the scope with every non-blank cell carved out.
                    let mut out = Vec::new();
                    for a in &scope {
                        let mut pieces = vec![*a];
                        for (r, c) in st.content_cells(self.sheet, *a) {
                            subtract(&mut pieces, &Rect::cell(r, c));
                        }
                        out.extend(pieces);
                    }
                    Ok(out)
                }
                12 => {
                    // Visible cells: the range with hidden rows and columns carved out.
                    let mut out = Vec::new();
                    for a in &self.areas {
                        let Some(clip) = st.clip(self.sheet, *a) else { continue };
                        let mut pieces = vec![clip];
                        let mut r = clip.r1;
                        while r <= clip.r2 {
                            if st.book.row_hidden(self.sheet, r) {
                                let start = r;
                                while r < clip.r2 && st.book.row_hidden(self.sheet, r + 1) {
                                    r += 1;
                                }
                                subtract(&mut pieces, &Rect::new(start, clip.c1, r, clip.c2));
                            }
                            r += 1;
                        }
                        let mut c = clip.c1;
                        while c <= clip.c2 {
                            if st.book.column_hidden(self.sheet, c) {
                                let start = c;
                                while c < clip.c2 && st.book.column_hidden(self.sheet, c + 1) {
                                    c += 1;
                                }
                                subtract(&mut pieces, &Rect::new(clip.r1, start, clip.r2, c));
                            }
                            c += 1;
                        }
                        out.extend(pieces);
                    }
                    Ok(out)
                }
                -4144 => {
                    let mut cells = Vec::new();
                    for a in &scope {
                        for (r, c) in st.book.occupied_cells(self.sheet, *a) {
                            if st.book.cell_comment(self.sheet, r, c).is_some() {
                                cells.push((r, c));
                            }
                        }
                    }
                    Ok(cells_to_areas(&cells))
                }
                _ => Err(err1004("SpecialCells method of Range class failed")),
            }
        })?;
        if found.is_empty() {
            return Err(err1004("No cells were found."));
        }
        Ok(self.with_areas(found, Kind::Cells))
    }

    fn sort(&self, args: &Args) -> Result<Value> {
        crate::sorting::range_sort(&self.app, self.sheet, self.a0(), args)
    }

    /// `RemoveDuplicates Columns, Header`: later rows equal (case-insensitively) on the key
    /// columns are deleted and the rest move up within the range.
    fn remove_duplicates(&self, args: &Args) -> Result<Value> {
        let a = self.a0();
        let cols: Vec<u32> = match args.get(0, "Columns").map(data).transpose()? {
            Some(Value::Array(arr)) => {
                let items = arr.borrow().data.clone();
                let mut v = Vec::new();
                for x in items {
                    v.push(x.to_i64()? as u32);
                }
                v
            }
            Some(other) => vec![other.to_i64()? as u32],
            None => (1..=a.cols()).collect(),
        };
        let header = opt_i64(args, 1, "Header")?.unwrap_or(2) == 1;
        with_mut(&self.app, |st| {
            st.check_writable(self.sheet)?;
            let Some(a) = st.clip(self.sheet, a) else { return Ok(Value::Empty) };
            let start = if header { a.r1 + 1 } else { a.r1 };
            let mut seen: Vec<Vec<String>> = Vec::new();
            let mut keep: Vec<u32> = Vec::new();
            for r in start..=a.r2 {
                let key: Vec<String> = cols
                    .iter()
                    .map(|&k| {
                        let v = st.book.cell_value(self.sheet, r, a.c1 + k - 1);
                        match v {
                            Value::String(s) => format!("s{}", s.to_lowercase()),
                            Value::Empty => String::new(),
                            other => format!("v{}", other.to_text().unwrap_or_default()),
                        }
                    })
                    .collect();
                if !seen.contains(&key) {
                    seen.push(key);
                    keep.push(r);
                }
            }
            type Cell = (Value, Option<String>, CellFormat);
            let rows: Vec<Vec<Cell>> = keep
                .iter()
                .map(|&r| {
                    (a.c1..=a.c2)
                        .map(|c| {
                            (
                                st.book.cell_value(self.sheet, r, c),
                                st.book.cell_formula(self.sheet, r, c),
                                st.book.cell_format(self.sheet, r, c),
                            )
                        })
                        .collect()
                })
                .collect();
            for (i, r) in (start..=a.r2).enumerate() {
                for (j, c) in (a.c1..=a.c2).enumerate() {
                    match rows.get(i).and_then(|row| row.get(j)) {
                        Some((v, f, fmt)) => {
                            match f {
                                Some(f) => st.book.set_cell_formula(self.sheet, r, c, Some(f.clone()))?,
                                None => st.book.set_cell_value(self.sheet, r, c, v.clone())?,
                            }
                            st.book.set_cell_format(self.sheet, r, c, fmt.clone())?;
                        }
                        None => {
                            st.book.set_cell_value(self.sheet, r, c, Value::Empty)?;
                        }
                    }
                }
            }
            st.saved = false;
            Ok(Value::Empty)
        })
    }

    /// `Replace What, Replacement, LookAt, SearchOrder, MatchCase`: in constants and formulas.
    fn replace(&self, args: &Args) -> Result<Value> {
        let what = data(args.require(0, "What")?)?.to_text()?;
        let repl = data(args.require(1, "Replacement")?)?.to_text()?;
        let last = with(&self.app, |st| Ok(st.last_find.clone()))?;
        let look_at = opt_i64(args, 2, "LookAt")?.or(last.as_ref().map(|p| i64::from(p.look_at))).unwrap_or(2);
        let match_case = opt_bool(args, 4, "MatchCase")?.unwrap_or(false);
        let areas = self.areas.clone();
        with_mut(&self.app, |st| {
            // LookAt is shared with Find and remembered.
            let mut p = st.last_find.clone().unwrap_or(crate::state::FindParams {
                what: what.clone(),
                look_in: -4123,
                look_at: 2,
                match_case: false,
                by_columns: false,
            });
            p.look_at = look_at as i32;
            st.last_find = Some(p);
            for a in &areas {
                for (r, c) in st.content_cells(self.sheet, *a) {
                    let text = Self::formula_of(st, self.sheet, r, c, false);
                    let new = if look_at == 1 {
                        if wildcard_match(&what, &text, true, match_case) { Some(repl.clone()) } else { None }
                    } else {
                        replace_part(&text, &what, &repl, match_case)
                    };
                    if let Some(n) = new
                        && n != text
                    {
                        st.check_cell(self.sheet, r, c)?;
                        let input = crate::state::parse_typed(&n);
                        st.apply_input(self.sheet, r, c, input)?;
                    }
                }
            }
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    /// Whole rows (`true`) or whole columns (`false`) this range spans, for `Hidden`.
    fn hidden_axis(&self) -> Result<bool> {
        let a = self.a0();
        if a.is_entire_rows() || self.kind == Kind::Rows && a.is_entire_rows() {
            Ok(true)
        } else if a.is_entire_columns() {
            Ok(false)
        } else {
            Err(err1004("Unable to set the Hidden property of the Range class"))
        }
    }

    fn row_list(&self) -> Vec<u32> {
        let mut v: Vec<u32> = self.areas.iter().flat_map(|a| a.r1..=a.r2).collect();
        v.sort_unstable();
        v.dedup();
        v
    }

    fn col_list(&self) -> Vec<u32> {
        let mut v: Vec<u32> = self.areas.iter().flat_map(|a| a.c1..=a.c2).collect();
        v.sort_unstable();
        v.dedup();
        v
    }

    fn sub_object(&self, which: Sub) -> Value {
        let fo = FormatObj { app: self.app.clone(), sheet: self.sheet, areas: self.areas.clone(), which };
        Value::Object(Some(Rc::new(fo)))
    }

    fn merge_state(&self) -> Result<Value> {
        let a = self.a0();
        with(&self.app, |st| {
            let merges = st.book.merged_areas(self.sheet);
            if merges.contains(&a) || merges.iter().any(|m| m.contains_rect(&a)) {
                Ok(Value::Boolean(true))
            } else if merges.iter().any(|m| m.intersect(&a).is_some()) {
                Ok(Value::Null)
            } else {
                Ok(Value::Boolean(false))
            }
        })
    }

    fn log(&self, what: &str) -> Result<Value> {
        let addr = self.address(&Args::none())?;
        with_mut(&self.app, |st| {
            let name = st.sheet_name(self.sheet)?;
            st.log.push(format!("{what} {name}!{addr}"));
            Ok(Value::Boolean(true))
        })
    }
}

/// Number formats of the built-in cell styles.
fn style_format(name: &str) -> Option<Option<String>> {
    Some(match name.to_ascii_lowercase().as_str() {
        "normal" => None,
        "currency" => Some("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)".into()),
        "currency [0]" => Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)".into()),
        "comma" => Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)".into()),
        "comma [0]" => Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)".into()),
        "percent" => Some("0%".into()),
        _ => return None,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SetMode {
    Value,
    Formula,
    FormulaR1C1,
}

/// Inserted rows/columns take the format of the neighbor (`CopyOrigin`: 0 above/left, 1 below/right).
fn inherit_format(st: &mut State, sheet: SheetId, a: Rect, down: bool, origin: i64) -> Result<()> {
    let src = if down {
        if origin == 1 {
            (a.r2 < MAX_ROWS).then(|| Rect::new(a.r2 + 1, a.c1, a.r2 + 1, a.c2))
        } else {
            (a.r1 > 1).then(|| Rect::new(a.r1 - 1, a.c1, a.r1 - 1, a.c2))
        }
    } else if origin == 1 {
        (a.c2 < MAX_COLS).then(|| Rect::new(a.r1, a.c2 + 1, a.r2, a.c2 + 1))
    } else {
        (a.c1 > 1).then(|| Rect::new(a.r1, a.c1 - 1, a.r2, a.c1 - 1))
    };
    let Some(src) = src else { return Ok(()) };
    for (r, c) in st.book.occupied_cells(sheet, src) {
        let f = st.book.cell_format(sheet, r, c);
        if f.is_default() {
            continue;
        }
        if down {
            for rr in a.r1..=a.r2.min(a.r1 + 10_000) {
                st.book.set_cell_format(sheet, rr, c, f.clone())?;
            }
        } else {
            for cc in a.c1..=a.c2.min(a.c1 + 10_000) {
                st.book.set_cell_format(sheet, r, cc, f.clone())?;
            }
        }
    }
    Ok(())
}

impl Object for RangeObj {
    fn type_name(&self) -> String {
        "Range".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        if lname != RANGE_REF.to_ascii_lowercase() {
            self.check_sheet()?;
        }
        let a = self.a0();
        match lname.as_str() {
            "" | "_default" => {
                if args.is_empty() {
                    self.get_value(false)
                } else {
                    self.item(args)
                }
            }
            "item" => self.item(args),
            "value" => {
                if let Some(t) = opt_i64(args, 0, "RangeValueDataType")?
                    && t != 10
                {
                    return Err(err1004("Application-defined or object-defined error"));
                }
                self.get_value(false)
            }
            "value2" => self.get_value(true),
            "formula" | "formulalocal" | "formula2" => {
                with(&self.app, |st| {
                    if st.contents_protected(self.sheet) && st.book.cell_format(self.sheet, a.r1, a.c1).formula_hidden {
                        return Err(err1004("Application-defined or object-defined error"));
                    }
                    Ok(())
                })?;
                self.get_text_grid(|st, r, c| Self::formula_of(st, self.sheet, r, c, false))
            }
            "formular1c1" | "formular1c1local" | "formula2r1c1" => {
                self.get_text_grid(|st, r, c| Self::formula_of(st, self.sheet, r, c, true))
            }
            "hasformula" => with(&self.app, |st| {
                let mut any = false;
                let mut all = true;
                for ar in &self.areas {
                    let n = st
                        .book
                        .occupied_cells(self.sheet, *ar)
                        .into_iter()
                        .filter(|&(r, c)| st.book.cell_formula(self.sheet, r, c).is_some())
                        .count() as u64;
                    any |= n > 0;
                    all &= n == ar.area();
                }
                Ok(if all {
                    Value::Boolean(true)
                } else if any {
                    Value::Null
                } else {
                    Value::Boolean(false)
                })
            }),
            "text" => with(&self.app, |st| {
                if a.is_cell() {
                    return Ok(Value::from(st.text(self.sheet, a.r1, a.c1)));
                }
                let Some(clip) = st.clip(self.sheet, a) else { return Ok(Value::from("")) };
                // The first cell's text (CONFORMANCE sheet3 V019: a String, not Null, when cells differ).
                Ok(Value::from(st.text(self.sheet, clip.r1, clip.c1)))
            }),
            "row" => Ok(Value::Long(a.r1 as i32)),
            "column" => Ok(Value::Long(a.c1 as i32)),
            "rows" => {
                if args.is_empty() {
                    Ok(self.with_areas(self.areas.clone(), Kind::Rows))
                } else {
                    RangeObj { app: self.app.clone(), sheet: self.sheet, areas: self.areas.clone(), kind: Kind::Rows }
                        .item(args)
                }
            }
            "columns" => {
                if args.is_empty() {
                    Ok(self.with_areas(self.areas.clone(), Kind::Columns))
                } else {
                    RangeObj {
                        app: self.app.clone(),
                        sheet: self.sheet,
                        areas: self.areas.clone(),
                        kind: Kind::Columns,
                    }
                    .item(args)
                }
            }
            "cells" => {
                if args.is_empty() {
                    Ok(self.with_areas(self.areas.clone(), Kind::Cells))
                } else {
                    RangeObj { app: self.app.clone(), sheet: self.sheet, areas: self.areas.clone(), kind: Kind::Cells }
                        .item(args)
                }
            }
            "count" => {
                let n = self.count();
                i32::try_from(n).map(Value::Long).map_err(|_| RuntimeError::overflow())
            }
            "countlarge" => Ok(Value::LongLong(self.count() as i64)),
            "address" | "addresslocal" => Ok(Value::from(self.address(args)?)),
            "offset" => self.offset(args),
            "resize" => self.resize(args),
            "end" => {
                let dir = opt_i64(args, 0, "Direction")?.ok_or_else(|| RuntimeError::standard(449))?;
                let (r, c) = with(&self.app, |st| end_from(st, self.sheet, a.r1, a.c1, dir))?;
                Ok(self.with_areas(vec![Rect::cell(r, c)], Kind::Cells))
            }
            "currentregion" => {
                let r = with(&self.app, |st| Ok(current_region(st, self.sheet, a)))?;
                Ok(self.with_areas(vec![r], Kind::Cells))
            }
            "entirerow" => {
                Ok(self.with_areas(self.areas.iter().map(|x| Rect::new(x.r1, 1, x.r2, MAX_COLS)).collect(), Kind::Rows))
            }
            "entirecolumn" => {
                Ok(self
                    .with_areas(self.areas.iter().map(|x| Rect::new(1, x.c1, MAX_ROWS, x.c2)).collect(), Kind::Columns))
            }
            "areas" => {
                let ar = AreasObj { app: self.app.clone(), sheet: self.sheet, areas: self.areas.clone() };
                if args.is_empty() { Ok(Value::Object(Some(Rc::new(ar)))) } else { ar.call("Item", args) }
            }
            "range" => self.range_method(args),
            "select" => self.select().map(|_| Value::Boolean(true)),
            "activate" => self.activate().map(|_| Value::Boolean(true)),
            "copy" => self.copy(args, false),
            "cut" => self.copy(args, true),
            "pastespecial" => self.paste_special(args),
            "insert" => self.insert(args),
            "delete" => self.delete(args),
            "clearcontents" => self.clear(true, false),
            "clear" => self.clear(true, true),
            "clearformats" => self.clear(false, true),
            "interior" => Ok(self.sub_object(Sub::Interior)),
            "font" => Ok(self.sub_object(Sub::Font)),
            "borders" => {
                let b = self.sub_object(Sub::Borders);
                if args.is_empty() {
                    Ok(b)
                } else {
                    let idx = opt_i64(args, 0, "Index")?.ok_or_else(|| RuntimeError::standard(449))?;
                    Ok(self.sub_object(Sub::Border(idx as i32)))
                }
            }
            "borderaround" => {
                let style = opt_i64(args, 0, "LineStyle")?.unwrap_or(1) as i32;
                let weight = opt_i64(args, 1, "Weight")?.unwrap_or(2) as i32;
                let color = match args.get(3, "Color") {
                    Some(v) => v.to_i64()? as u32,
                    None => 0,
                };
                for edge in [7, 8, 9, 10] {
                    let line = BorderLine { style, weight, color };
                    with_mut(&self.app, |st| set_edge(st, self.sheet, &self.areas, edge, Some(line)))?;
                }
                Ok(Value::Boolean(true))
            }
            "horizontalalignment" => with(&self.app, |st| {
                Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| halign_constant(f.h_align))))
            }),
            "verticalalignment" => with(&self.app, |st| {
                Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| f.v_align.unwrap_or(-4107))))
            }),
            "numberformat" | "numberformatlocal" => with(&self.app, |st| {
                Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| {
                    f.number_format.clone().unwrap_or_else(|| "General".into())
                })))
            }),
            "wraptext" => with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| f.wrap)))),
            "hidden" => {
                let rows = self.hidden_axis()?;
                with(&self.app, |st| {
                    let list = if rows { self.row_list() } else { self.col_list() };
                    let all = list.into_iter().all(|i| {
                        if rows { st.book.row_hidden(self.sheet, i) } else { st.book.column_hidden(self.sheet, i) }
                    });
                    Ok(Value::Boolean(all))
                })
            }
            "rowheight" | "height" => with(&self.app, |st| {
                let mut acc: Option<f64> = None;
                let mut total = 0.0;
                for r in self.row_list() {
                    let h = if st.book.row_hidden(self.sheet, r) { 0.0 } else { st.book.row_height(self.sheet, r) };
                    total += h;
                    // Mixed heights report the first row's (not Null).
                    if acc.is_none() {
                        acc = Some(h);
                    }
                }
                Ok(if lname == "height" { Value::Double(total) } else { mixed(acc) })
            }),
            "columnwidth" => with(&self.app, |st| {
                let mut acc: Option<f64> = None;
                for c in self.col_list() {
                    let w =
                        if st.book.column_hidden(self.sheet, c) { 0.0 } else { st.book.column_width(self.sheet, c) };
                    if acc.is_none() {
                        acc = Some(w);
                    }
                }
                Ok(mixed(acc))
            }),
            "find" => self.find(args),
            "findnext" => self.find_next(args, false),
            "findprevious" => self.find_next(args, true),
            "specialcells" => self.special_cells(args),
            "sort" => self.sort(args),
            "autofilter" => crate::filter::range_autofilter(&self.app, self.sheet, a, args),
            "removeduplicates" => self.remove_duplicates(args),
            "replace" => self.replace(args),
            "autofill" => {
                let dest = args.require(0, "Destination")?;
                let (_, d) = range_arg(&self.app, self.sheet, dest)?;
                let ty = opt_i64(args, 1, "Type")?.unwrap_or(0);
                with_mut(&self.app, |st| crate::fill::autofill(st, self.sheet, a, d[0], ty))
                    .map(|_| Value::Boolean(true))
            }
            "filldown" | "fillup" | "fillright" | "fillleft" => {
                let dir = match lname.as_str() {
                    "filldown" => 'd',
                    "fillup" => 'u',
                    "fillright" => 'r',
                    _ => 'l',
                };
                with_mut(&self.app, |st| crate::fill::fill_copy(st, self.sheet, a, dir)).map(|_| Value::Boolean(true))
            }
            "dataseries" => {
                let by_columns = match opt_i64(args, 0, "Rowcol")? {
                    Some(n) => n == 2,
                    None => a.rows() > a.cols(),
                };
                let growth = opt_i64(args, 1, "Type")?.unwrap_or(-4132) == 2;
                let step = args.get(3, "Step").map(|v| data(v)?.to_f64()).transpose()?.unwrap_or(1.0);
                with_mut(&self.app, |st| crate::fill::data_series(st, self.sheet, a, by_columns, growth, step))
                    .map(|_| Value::Boolean(true))
            }
            "locked" => with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| !f.unlocked)))),
            "formulahidden" => {
                with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| f.formula_hidden))))
            }
            "indentlevel" => {
                with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| i32::from(f.indent)))))
            }
            "orientation" => with(&self.app, |st| {
                Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| {
                    if f.orientation == 0 { -4128 } else { f.orientation }
                })))
            }),
            "shrinktofit" => {
                with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, |f| f.shrink_to_fit))))
            }
            "style" => with(&self.app, |st| {
                let nf = st.book.cell_format(self.sheet, a.r1, a.c1).number_format;
                let name = ["Currency", "Currency [0]", "Comma", "Comma [0]", "Percent"]
                    .into_iter()
                    .find(|n| style_format(n).flatten() == nf && nf.is_some())
                    .unwrap_or("Normal");
                Ok(Value::from(name))
            }),
            "addcomment" => crate::notes::add_comment(&self.app, self.sheet, a.r1, a.c1, args),
            "comment" => crate::notes::comment_of(&self.app, self.sheet, a.r1, a.c1),
            "clearcomments" => crate::notes::clear_comments(&self.app, self.sheet, &self.areas).map(|_| Value::Empty),
            "hyperlinks" => crate::notes::hyperlinks_value(&self.app, self.sheet, Some(self.areas.clone()), args),
            "clearhyperlinks" => {
                crate::notes::clear_hyperlinks(&self.app, self.sheet, &self.areas).map(|_| Value::Empty)
            }
            "name" => {
                let found = with(&self.app, |st| {
                    let sheet = st.sheet_name(self.sheet)?;
                    let want = format!("{}!{}", quote_sheet(&sheet), self.address(&Args::none()).unwrap_or_default());
                    Ok(st.book.defined_names().into_iter().find(|(_, _, r)| r.eq_ignore_ascii_case(&want)))
                })?;
                match found {
                    Some((n, scope, _)) => Ok(Value::Object(Some(Rc::new(crate::names::NameObj {
                        app: self.app.clone(),
                        name: n,
                        scope,
                    })))),
                    None => Err(err1004("Application-defined or object-defined error")),
                }
            }
            "autofit" => self.log("AutoFit"),
            "merge" => {
                let across = opt_bool(args, 0, "Across")?.unwrap_or(false);
                with_mut(&self.app, |st| {
                    for ar in &self.areas {
                        // Only the upper-left value of each merged block stays.
                        let blocks: Vec<Rect> = if across {
                            (ar.r1..=ar.r2).map(|r| Rect::new(r, ar.c1, r, ar.c2)).collect()
                        } else {
                            vec![*ar]
                        };
                        for b in blocks {
                            for (r, c) in st.content_cells(self.sheet, b) {
                                if (r, c) != (b.r1, b.c1) {
                                    st.book.set_cell_value(self.sheet, r, c, Value::Empty)?;
                                }
                            }
                        }
                        if across {
                            for r in ar.r1..=ar.r2 {
                                st.book.merge_cells(self.sheet, Rect::new(r, ar.c1, r, ar.c2))?;
                            }
                        } else {
                            st.book.merge_cells(self.sheet, *ar)?;
                        }
                    }
                    Ok(())
                })?;
                self.log("Merge")
            }
            "unmerge" => {
                with_mut(&self.app, |st| {
                    for ar in &self.areas {
                        st.book.unmerge_cells(self.sheet, *ar)?;
                    }
                    Ok(())
                })?;
                self.log("UnMerge")
            }
            "mergecells" => self.merge_state(),
            "mergearea" => {
                let m = with(&self.app, |st| {
                    Ok(st.book.merged_areas(self.sheet).into_iter().find(|m| m.contains(a.r1, a.c1)))
                })?;
                Ok(self.with_areas(vec![m.unwrap_or(Rect::cell(a.r1, a.c1))], Kind::Cells))
            }
            "calculate" => with_mut(&self.app, |st| {
                st.book.calculate();
                Ok(Value::Boolean(true))
            }),
            "worksheet" | "parent" => Ok(sheet_value(&self.app, self.sheet)),
            "application" => Ok(crate::objects::app_value(&self.app)),
            "next" => {
                let cell = cell_at(i64::from(a.r1), i64::from(a.c1) + 1)?;
                Ok(self.with_areas(vec![cell], Kind::Cells))
            }
            "previous" => {
                let cell = cell_at(i64::from(a.r1), i64::from(a.c1) - 1)?;
                Ok(self.with_areas(vec![cell], Kind::Cells))
            }
            n if n == RANGE_REF.to_ascii_lowercase() => {
                let areas: Vec<String> =
                    self.areas.iter().map(|r| format!("{},{},{},{}", r.r1, r.c1, r.r2, r.c2)).collect();
                Ok(Value::from(format!("{}|{}", self.sheet, areas.join(";"))))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        self.check_sheet()?;
        match name.to_ascii_lowercase().as_str() {
            "" | "_default" | "value" | "value2" => {
                if name.is_empty() && !args.is_empty() {
                    // rng(2) = x
                    let target = self.item(args)?;
                    if let Value::Object(Some(o)) = target {
                        return o.assign("Value", &Args::none(), value);
                    }
                    return Err(RuntimeError::object_required());
                }
                self.set_content(value, SetMode::Value)
            }
            "formula" | "formulalocal" | "formula2" => self.set_content(value, SetMode::Formula),
            "formular1c1" | "formular1c1local" | "formula2r1c1" => self.set_content(value, SetMode::FormulaR1C1),
            "numberformat" | "numberformatlocal" => {
                let text = data(&value)?.to_text()?;
                with_mut(&self.app, |st| {
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| {
                        f.number_format = if text.eq_ignore_ascii_case("general") { None } else { Some(text.clone()) };
                    })
                })
            }
            "horizontalalignment" => {
                let h = halign_from(data(&value)?.to_i64()?)?;
                with_mut(&self.app, |st| update_format(st, self.sheet, &self.areas, |f, _, _, _| f.h_align = h))
            }
            "verticalalignment" => {
                let n = data(&value)?.to_i64()? as i32;
                if ![-4107, -4108, -4160, -4130, -4117].contains(&n) {
                    return Err(err1004("Unable to set the VerticalAlignment property of the Range class"));
                }
                with_mut(&self.app, |st| {
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| f.v_align = (n != -4107).then_some(n))
                })
            }
            "locked" | "formulahidden" => {
                let b = data(&value)?.to_bool()?;
                let locked = name.eq_ignore_ascii_case("locked");
                with_mut(&self.app, |st| {
                    if st.contents_protected(self.sheet) {
                        let what = if locked { "Locked" } else { "FormulaHidden" };
                        return Err(err1004(&format!("Unable to set the {what} property of the Range class")));
                    }
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| {
                        if locked {
                            f.unlocked = !b;
                        } else {
                            f.formula_hidden = b;
                        }
                    })
                })
            }
            "indentlevel" => {
                let n = data(&value)?.to_i64()?;
                if !(0..=250).contains(&n) {
                    return Err(err1004("Unable to set the IndentLevel property of the Range class"));
                }
                with_mut(&self.app, |st| {
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| {
                        f.indent = n as u8;
                        if n > 0 && matches!(f.h_align, HAlign::General) {
                            f.h_align = HAlign::Left;
                        }
                    })
                })
            }
            "orientation" => {
                let n = data(&value)?.to_i64()? as i32;
                let deg = match n {
                    -4128 => 0,
                    -4171 => 90,
                    -4170 => -90,
                    -4166 => -4166,
                    d if (-90..=90).contains(&d) => d,
                    _ => return Err(err1004("Unable to set the Orientation property of the Range class")),
                };
                with_mut(&self.app, |st| {
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| {
                        // An indented cell stays horizontal.
                        if f.indent == 0 {
                            f.orientation = deg;
                        }
                    })
                })
            }
            "shrinktofit" => {
                let b = data(&value)?.to_bool()?;
                with_mut(&self.app, |st| update_format(st, self.sheet, &self.areas, |f, _, _, _| f.shrink_to_fit = b))
            }
            "style" => {
                let n = data(&value)?.to_text()?;
                let fmt =
                    style_format(&n).ok_or_else(|| err1004("Unable to set the Style property of the Range class"))?;
                with_mut(&self.app, |st| {
                    update_format(st, self.sheet, &self.areas, |f, _, _, _| f.number_format = fmt.clone())
                })
            }
            "name" => {
                let n = data(&value)?.to_text()?;
                let refers = with(&self.app, |st| {
                    let sheet = st.sheet_name(self.sheet)?;
                    Ok(format!("{}!{}", quote_sheet(&sheet), self.address(&Args::none()).unwrap_or_default()))
                })?;
                crate::names::add_name(&self.app, &n, None, &refers).map(|_| ())
            }
            "wraptext" => {
                let w = data(&value)?.to_bool()?;
                with_mut(&self.app, |st| update_format(st, self.sheet, &self.areas, |f, _, _, _| f.wrap = w))
            }
            "hidden" => {
                let rows = self.hidden_axis()?;
                let h = data(&value)?.to_bool()?;
                let list = if rows { self.row_list() } else { self.col_list() };
                with_mut(&self.app, |st| {
                    for i in list {
                        if rows {
                            st.book.set_row_hidden(self.sheet, i, h)?;
                        } else {
                            st.book.set_column_hidden(self.sheet, i, h)?;
                        }
                    }
                    Ok(())
                })
            }
            "rowheight" => {
                let h = data(&value)?.to_f64()?;
                let rows = self.row_list();
                with_mut(&self.app, |st| {
                    for r in rows {
                        if h == 0.0 {
                            st.book.set_row_hidden(self.sheet, r, true)?;
                        } else {
                            st.book.set_row_height(self.sheet, r, h)?;
                            st.book.set_row_hidden(self.sheet, r, false)?;
                        }
                    }
                    Ok(())
                })
            }
            "columnwidth" => {
                let w = data(&value)?.to_f64()?;
                let cols = self.col_list();
                with_mut(&self.app, |st| {
                    for c in cols {
                        if w == 0.0 {
                            st.book.set_column_hidden(self.sheet, c, true)?;
                        } else {
                            st.book.set_column_width(self.sheet, c, w)?;
                            st.book.set_column_hidden(self.sheet, c, false)?;
                        }
                    }
                    Ok(())
                })
            }
            "mergecells" => {
                let m = data(&value)?.to_bool()?;
                self.call(if m { "Merge" } else { "UnMerge" }, &Args::none()).map(|_| ())
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        self.check_sheet()?;
        let mut out = Vec::new();
        for a in &self.areas {
            let clip = with(&self.app, |st| Ok(st.clip(self.sheet, *a)))?;
            let Some(clip) = clip else { continue };
            match self.kind {
                Kind::Cells => {
                    for r in clip.r1..=clip.r2 {
                        for c in clip.c1..=clip.c2 {
                            out.push(self.with_areas(vec![Rect::cell(r, c)], Kind::Cells));
                        }
                    }
                }
                Kind::Rows => {
                    let rows = if a.area() > LARGE { (clip.r1, clip.r2) } else { (a.r1, a.r2) };
                    for r in rows.0..=rows.1 {
                        out.push(self.with_areas(vec![Rect::new(r, a.c1, r, a.c2)], Kind::Rows));
                    }
                }
                Kind::Columns => {
                    let cols = if a.area() > LARGE { (clip.c1, clip.c2) } else { (a.c1, a.c2) };
                    for c in cols.0..=cols.1 {
                        out.push(self.with_areas(vec![Rect::new(a.r1, c, a.r2, c)], Kind::Columns));
                    }
                }
            }
        }
        Ok(out)
    }
}

/// `Range.Areas`.
struct AreasObj {
    app: Shared,
    sheet: SheetId,
    areas: Vec<Rect>,
}

impl Object for AreasObj {
    fn type_name(&self) -> String {
        "Areas".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.areas.len() as i32)),
            "" | "item" => {
                let i = opt_i64(args, 0, "Index")?.ok_or_else(|| RuntimeError::standard(449))?;
                let a =
                    usize::try_from(i - 1).ok().and_then(|i| self.areas.get(i)).ok_or_else(RuntimeError::subscript)?;
                Ok(range_value(&self.app, self.sheet, vec![*a], Kind::Cells))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn items(&self) -> Result<Vec<Value>> {
        Ok(self.areas.iter().map(|a| range_value(&self.app, self.sheet, vec![*a], Kind::Cells)).collect())
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Sub {
    Interior,
    Font,
    Borders,
    Border(i32),
}

/// `Interior`, `Font`, `Borders` and `Border(edge)` of a range.
struct FormatObj {
    app: Shared,
    sheet: SheetId,
    areas: Vec<Rect>,
    which: Sub,
}

/// Cells of `a` whose given side carries edge `edge` (xlEdgeLeft 7, Top 8, Bottom 9, Right 10,
/// InsideVertical 11, InsideHorizontal 12), as (row-range, col-range, side).
fn edge_cells(a: &Rect, edge: i32) -> Option<(Rect, Side)> {
    match edge {
        7 => Some((Rect::new(a.r1, a.c1, a.r2, a.c1), Side::Left)),
        8 => Some((Rect::new(a.r1, a.c1, a.r1, a.c2), Side::Top)),
        9 => Some((Rect::new(a.r2, a.c1, a.r2, a.c2), Side::Bottom)),
        10 => Some((Rect::new(a.r1, a.c2, a.r2, a.c2), Side::Right)),
        11 => (a.c2 > a.c1).then(|| (Rect::new(a.r1, a.c1, a.r2, a.c2 - 1), Side::Right)),
        12 => (a.r2 > a.r1).then(|| (Rect::new(a.r1, a.c1, a.r2 - 1, a.c2), Side::Bottom)),
        5 => Some((*a, Side::DiagDown)),
        6 => Some((*a, Side::DiagUp)),
        _ => None,
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Side {
    Left,
    Top,
    Bottom,
    Right,
    DiagDown,
    DiagUp,
}

fn side_mut(f: &mut CellFormat, s: Side) -> &mut Option<BorderLine> {
    match s {
        Side::Left => &mut f.borders.left,
        Side::Top => &mut f.borders.top,
        Side::Bottom => &mut f.borders.bottom,
        Side::Right => &mut f.borders.right,
        Side::DiagDown => &mut f.borders.diagonal_down,
        Side::DiagUp => &mut f.borders.diagonal_up,
    }
}

/// Reads one side of a cell, also honoring the neighbor's facing side (a shared edge).
fn read_side(st: &State, sheet: SheetId, r: u32, c: u32, s: Side) -> Option<BorderLine> {
    let mut f = st.book.cell_format(sheet, r, c);
    if let Some(b) = *side_mut(&mut f, s) {
        return Some(b);
    }
    let (nr, nc, ns) = match s {
        Side::Left => (Some(r), c.checked_sub(1).filter(|&x| x >= 1), Side::Right),
        Side::Right => (Some(r), (c < MAX_COLS).then_some(c + 1), Side::Left),
        Side::Top => (r.checked_sub(1).filter(|&x| x >= 1), Some(c), Side::Bottom),
        Side::Bottom => ((r < MAX_ROWS).then_some(r + 1), Some(c), Side::Top),
        Side::DiagDown | Side::DiagUp => return None,
    };
    let (Some(nr), Some(nc)) = (nr, nc) else { return None };
    let mut nf = st.book.cell_format(sheet, nr, nc);
    *side_mut(&mut nf, ns)
}

/// Sets (or clears with `None`) edge `edge` over `areas`; `edge` 0 means every edge of every cell.
fn set_edge(st: &mut State, sheet: SheetId, areas: &[Rect], edge: i32, line: Option<BorderLine>) -> Result<()> {
    if edge == 0 {
        return update_format(st, sheet, areas, |f, _, _, _| {
            f.borders.left = line;
            f.borders.top = line;
            f.borders.bottom = line;
            f.borders.right = line;
        });
    }
    for a in areas {
        let Some((cells, side)) = edge_cells(a, edge) else { continue };
        update_format(st, sheet, &[cells], |f, _, _, _| *side_mut(f, side) = line)?;
        // Clearing must also clear the neighbor's facing side.
        if line.is_none() {
            let facing = match side {
                Side::Left => cells.offset(0, -1).map(|x| (x, Side::Right)),
                Side::Right => cells.offset(0, 1).map(|x| (x, Side::Left)),
                Side::Top => cells.offset(-1, 0).map(|x| (x, Side::Bottom)),
                Side::Bottom => cells.offset(1, 0).map(|x| (x, Side::Top)),
                Side::DiagDown | Side::DiagUp => None,
            };
            if let Some((n, ns)) = facing
                && !matches!(edge, 11 | 12)
            {
                update_format(st, sheet, &[n], |f, _, _, _| *side_mut(f, ns) = None)?;
            }
        }
    }
    Ok(())
}

/// Common border line over an edge (or all edges for 0); outer `None` = mixed.
fn read_edge(st: &State, sheet: SheetId, areas: &[Rect], edge: i32) -> Option<Option<BorderLine>> {
    let edges: Vec<i32> = if edge == 0 { vec![7, 8, 9, 10] } else { vec![edge] };
    let mut acc: Option<Option<BorderLine>> = None;
    for a in areas {
        for &e in &edges {
            let Some((cells, side)) = edge_cells(a, e) else { continue };
            let Some(clip) = st.clip(sheet, cells) else { continue };
            for r in clip.r1..=clip.r2 {
                for c in clip.c1..=clip.c2 {
                    let v = read_side(st, sheet, r, c, side);
                    match &acc {
                        Some(x) if *x != v => return None,
                        Some(_) => {}
                        None => acc = Some(v),
                    }
                }
            }
        }
    }
    Some(acc.unwrap_or(None))
}

impl FormatObj {
    fn fold<T: PartialEq + Into<Value>>(&self, f: impl Fn(&CellFormat) -> T) -> Result<Value> {
        with(&self.app, |st| Ok(mixed(fold_format(st, self.sheet, &self.areas, f))))
    }

    fn update(&self, f: impl Fn(&mut CellFormat)) -> Result<()> {
        with_mut(&self.app, |st| update_format(st, self.sheet, &self.areas, |cf, _, _, _| f(cf)))
    }

    fn edge(&self) -> i32 {
        match self.which {
            Sub::Border(e) => e,
            _ => 0,
        }
    }

    fn border_get(&self, prop: &str) -> Result<Value> {
        let edge = self.edge();
        let line = with(&self.app, |st| Ok(read_edge(st, self.sheet, &self.areas, edge)))?;
        let Some(line) = line else { return Ok(Value::Null) };
        Ok(match (prop, line) {
            ("linestyle", None) => Value::Long(-4142),
            ("linestyle", Some(l)) => Value::Long(l.style),
            ("weight", None) => Value::Long(2),
            ("weight", Some(l)) => Value::Long(l.weight),
            ("color", None) => Value::Double(0.0),
            ("color", Some(l)) => Value::Double(f64::from(l.color)),
            ("colorindex", None) => Value::Long(-4142),
            ("colorindex", Some(l)) => Value::Long(nearest_color_index(l.color)),
            _ => Value::Null,
        })
    }

    fn border_set(&self, prop: &str, value: &Value) -> Result<()> {
        let edge = self.edge();
        let n = data(value)?.to_i64()?;
        with_mut(&self.app, |st| {
            let current = read_edge(st, self.sheet, &self.areas, edge).flatten();
            let base = current.unwrap_or_default();
            let line = match prop {
                "linestyle" => {
                    if n == -4142 || n == 0 {
                        None
                    } else if n == -4119 {
                        Some(BorderLine { style: -4119, weight: 4, ..base })
                    } else {
                        Some(BorderLine { style: n as i32, ..base })
                    }
                }
                "weight" => Some(BorderLine { weight: n as i32, ..base }),
                "color" => Some(BorderLine { color: n as u32, ..base }),
                "colorindex" => {
                    let color = if (1..=56).contains(&n) { PALETTE[(n - 1) as usize] } else { 0 };
                    Some(BorderLine { color, ..base })
                }
                _ => return Err(RuntimeError::no_member(prop)),
            };
            set_edge(st, self.sheet, &self.areas, edge, line)
        })
    }
}

impl Object for FormatObj {
    fn type_name(&self) -> String {
        match self.which {
            Sub::Interior => "Interior",
            Sub::Font => "Font",
            Sub::Borders => "Borders",
            Sub::Border(_) => "Border",
        }
        .into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        match (self.which, lname.as_str()) {
            (Sub::Interior, "color") => with(&self.app, |st| {
                let color = |f: &CellFormat| f64::from(f.fill_color.unwrap_or(0xFF_FF_FF));
                Ok(match fold_format(st, self.sheet, &self.areas, color) {
                    Some(c) => Value::Double(c),
                    // Mixed fills report the first cell's.
                    None => {
                        let a = self.areas[0];
                        Value::Double(color(&st.book.cell_format(self.sheet, a.r1, a.c1)))
                    }
                })
            }),
            (Sub::Interior, "colorindex") => self.fold(|f| f.fill_color.map_or(-4142, nearest_color_index)),
            (Sub::Interior, "pattern") => self.fold(|f| if f.fill_color.is_some() { 1 } else { -4142 }),
            (Sub::Interior, "patterncolorindex") => self.fold(|f| if f.fill_color.is_some() { -4105 } else { -4142 }),
            (Sub::Interior, "themecolor" | "tintandshade") => Ok(Value::Long(0)),
            (Sub::Font, "bold") => self.fold(|f| f.bold),
            (Sub::Font, "italic") => self.fold(|f| f.italic),
            (Sub::Font, "underline") => self.fold(|f| {
                if f.double_underline {
                    -4119
                } else if f.underline {
                    2
                } else {
                    -4142
                }
            }),
            (Sub::Font, "strikethrough") => self.fold(|f| f.strikethrough),
            (Sub::Font, "color") => self.fold(|f| f64::from(f.font_color.unwrap_or(0))),
            (Sub::Font, "colorindex") => self.fold(|f| f.font_color.map_or(-4105, nearest_color_index)),
            (Sub::Font, "size") => self.fold(|f| f.font_size.unwrap_or(11.0)),
            (Sub::Font, "name") => self.fold(|f| f.font_name.clone().unwrap_or_else(|| "Calibri".into())),
            (Sub::Font, "fontstyle") => self.fold(|f| {
                match (f.bold, f.italic) {
                    (true, true) => "Bold Italic",
                    (true, false) => "Bold",
                    (false, true) => "Italic",
                    (false, false) => "Regular",
                }
                .to_owned()
            }),
            (Sub::Borders, "" | "item") => {
                let idx = opt_i64(args, 0, "Index")?.ok_or_else(|| RuntimeError::standard(449))?;
                Ok(Value::Object(Some(Rc::new(FormatObj {
                    app: self.app.clone(),
                    sheet: self.sheet,
                    areas: self.areas.clone(),
                    which: Sub::Border(idx as i32),
                }))))
            }
            (Sub::Borders, "count") => Ok(Value::Long(6)),
            (Sub::Borders | Sub::Border(_), "linestyle" | "weight" | "color" | "colorindex") => self.border_get(&lname),
            (_, "parent") => Ok(range_value(&self.app, self.sheet, self.areas.clone(), Kind::Cells)),
            (_, "application") => Ok(crate::objects::app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let lname = name.to_ascii_lowercase();
        let v = data(&value)?;
        match (self.which, lname.as_str()) {
            (Sub::Interior, "color") => {
                let c = v.to_i64()? as u32;
                self.update(|f| f.fill_color = Some(c))
            }
            (Sub::Interior, "colorindex") => {
                let n = v.to_i64()?;
                let color = match n {
                    1..=56 => Some(PALETTE[(n - 1) as usize]),
                    -4142 | -4105 | 0 => None,
                    _ => return Err(err1004("Unable to set the ColorIndex property of the Interior class")),
                };
                self.update(|f| f.fill_color = color)
            }
            (Sub::Interior, "pattern") => {
                let n = v.to_i64()?;
                if n == -4142 {
                    self.update(|f| f.fill_color = None)
                } else {
                    self.update(|f| {
                        if f.fill_color.is_none() {
                            f.fill_color = Some(0xFF_FF_FF);
                        }
                    })
                }
            }
            (Sub::Interior, "patterncolorindex" | "themecolor" | "tintandshade" | "patterncolor") => Ok(()),
            (Sub::Font, "bold") => {
                let b = v.to_bool()?;
                self.update(|f| f.bold = b)
            }
            (Sub::Font, "italic") => {
                let b = v.to_bool()?;
                self.update(|f| f.italic = b)
            }
            (Sub::Font, "underline") => {
                let (u, d) = match &v {
                    Value::Boolean(b) => (*b, false),
                    other => {
                        let n = other.to_i64()?;
                        (!(n == -4142 || n == 0), n == -4119 || n == 5)
                    }
                };
                self.update(|f| {
                    f.underline = u;
                    f.double_underline = d;
                })
            }
            (Sub::Font, "strikethrough") => {
                let b = v.to_bool()?;
                self.update(|f| f.strikethrough = b)
            }
            (Sub::Font, "color") => {
                let c = v.to_i64()? as u32;
                self.update(|f| f.font_color = Some(c))
            }
            (Sub::Font, "colorindex") => {
                let n = v.to_i64()?;
                let color = match n {
                    1..=56 => Some(PALETTE[(n - 1) as usize]),
                    -4105 | -4142 | 0 => None,
                    _ => return Err(err1004("Unable to set the ColorIndex property of the Font class")),
                };
                self.update(|f| f.font_color = color)
            }
            (Sub::Font, "size") => {
                let s = v.to_f64()?;
                if !(1.0..=409.0).contains(&s) {
                    return Err(err1004("Unable to set the Size property of the Font class"));
                }
                self.update(|f| f.font_size = Some(s))
            }
            (Sub::Font, "name") => {
                let n = v.to_text()?;
                self.update(|f| f.font_name = Some(n.clone()))
            }
            (Sub::Font, "fontstyle") => {
                let s = v.to_text()?.to_lowercase();
                let (b, i) = (s.contains("bold"), s.contains("italic"));
                self.update(|f| {
                    f.bold = b;
                    f.italic = i;
                })
            }
            (Sub::Font, "themecolor" | "tintandshade" | "superscript" | "subscript" | "outlinefont" | "shadow") => {
                Ok(())
            }
            (Sub::Borders | Sub::Border(_), "linestyle" | "weight" | "color" | "colorindex") => {
                self.border_set(&lname, &v)
            }
            (Sub::Borders | Sub::Border(_), "themecolor" | "tintandshade") => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        match self.which {
            Sub::Borders => Ok([7, 8, 9, 10, 11, 12]
                .iter()
                .map(|&e| {
                    Value::Object(Some(Rc::new(FormatObj {
                        app: self.app.clone(),
                        sheet: self.sheet,
                        areas: self.areas.clone(),
                        which: Sub::Border(e),
                    }) as Rc<dyn Object>))
                })
                .collect()),
            _ => Err(RuntimeError::no_member("_NewEnum")),
        }
    }
}
