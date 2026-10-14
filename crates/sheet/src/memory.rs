//! A complete in-memory [`Workbook`] with sparse cell storage, for tests and headless runs.

use std::collections::{HashMap, HashSet};

use macroost_lang::{Result, RuntimeError, Value};

use crate::address::{MAX_COLS, MAX_ROWS, Rect};
use crate::workbook::{
    CellFormat, Hyperlink, PasswordHash, SheetId, SheetVisibility, ShiftDirection, Workbook, WorkbookProtection,
};

/// Default row height in points.
pub const DEFAULT_ROW_HEIGHT: f64 = 15.0;
/// Default column width in characters.
pub const DEFAULT_COLUMN_WIDTH: f64 = 8.43;

#[derive(Clone, Debug, Default)]
struct Cell {
    value: Value,
    formula: Option<String>,
    format: CellFormat,
    comment: Option<String>,
    link: Option<Hyperlink>,
}

impl Cell {
    fn is_unused(&self) -> bool {
        self.value.is_empty()
            && self.formula.is_none()
            && self.format.is_default()
            && self.comment.is_none()
            && self.link.is_none()
    }
}

#[derive(Clone, Debug, Default)]
struct Sheet {
    id: SheetId,
    name: String,
    code_name: Option<String>,
    visibility: SheetVisibility,
    protected: bool,
    password: Option<PasswordHash>,
    /// Extent of every cell touched since the last `reset_last_cell`.
    extent: Option<Rect>,
    row_formats: HashMap<u32, CellFormat>,
    col_formats: HashMap<u32, CellFormat>,
    cells: HashMap<(u32, u32), Cell>,
    row_heights: HashMap<u32, f64>,
    hidden_rows: HashSet<u32>,
    filtered_rows: HashSet<u32>,
    col_widths: HashMap<u32, f64>,
    hidden_cols: HashSet<u32>,
    merges: Vec<Rect>,
}

/// Sparse in-memory workbook.
#[derive(Clone, Debug)]
pub struct MemoryWorkbook {
    name: String,
    sheets: Vec<Sheet>,
    next_id: SheetId,
    /// Number of `calculate` calls, for tests.
    pub calculations: u32,
    /// Formula evaluation depth (guards against circular references).
    depth: std::cell::Cell<u32>,
    /// Manual calculation: formula cells keep the value of their last calculation.
    manual: bool,
    protection: Option<WorkbookProtection>,
    /// Defined names: (name, scope, refers-to without `=`).
    names: Vec<(String, Option<SheetId>, String)>,
}

impl Default for MemoryWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

fn err(msg: &str) -> RuntimeError {
    RuntimeError::new(1004, msg)
}

/// Checks a sheet name the way the application does: 1 to 31 characters, none of `[]:*?/\`,
/// not starting or ending with an apostrophe.
pub fn valid_sheet_name(name: &str) -> bool {
    let n = name.chars().count();
    (1..=31).contains(&n)
        && !name.chars().any(|c| matches!(c, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
        && !name.starts_with('\'')
        && !name.ends_with('\'')
}

/// Moves every key of `map` by `f`; keys mapped to `None` are dropped.
fn remap<V>(map: &mut HashMap<(u32, u32), V>, f: impl Fn(u32, u32) -> Option<Option<(u32, u32)>>) {
    let old = std::mem::take(map);
    let mut moved = Vec::new();
    for ((r, c), v) in old {
        match f(r, c) {
            None => {
                map.insert((r, c), v);
            }
            Some(Some(k)) => moved.push((k, v)),
            Some(None) => {}
        }
    }
    for (k, v) in moved {
        map.insert(k, v);
    }
}

fn remap_index<V>(map: &mut HashMap<u32, V>, f: impl Fn(u32) -> Option<u32>) {
    let old = std::mem::take(map);
    for (k, v) in old {
        if let Some(n) = f(k) {
            map.insert(n, v);
        }
    }
}

fn remap_set(set: &mut HashSet<u32>, f: impl Fn(u32) -> Option<u32>) {
    let old = std::mem::take(set);
    set.extend(old.into_iter().filter_map(f));
}

/// Index shift for inserting `count` at `at` with limit `max`: indexes >= at move up; ones pushed
/// past the end vanish.
fn insert_shift(i: u32, at: u32, count: u32, max: u32) -> Option<u32> {
    if i < at {
        Some(i)
    } else {
        let n = u64::from(i) + u64::from(count);
        (n <= u64::from(max)).then_some(n as u32)
    }
}

fn delete_shift(i: u32, at: u32, count: u32) -> Option<u32> {
    if i < at {
        Some(i)
    } else if i < at + count {
        None
    } else {
        Some(i - count)
    }
}

impl MemoryWorkbook {
    /// A workbook named `Book1` with one sheet, `Sheet1`.
    pub fn new() -> Self {
        Self::with_sheets("Book1", &["Sheet1"])
    }

    /// A workbook with the given sheets (in tab order).
    pub fn with_sheets(name: &str, sheets: &[&str]) -> Self {
        let mut wb = Self {
            name: name.to_owned(),
            sheets: Vec::new(),
            next_id: 1,
            calculations: 0,
            depth: std::cell::Cell::new(0),
            manual: false,
            protection: None,
            names: Vec::new(),
        };
        for s in sheets {
            let id = wb.next_id;
            wb.next_id += 1;
            wb.sheets.push(Sheet { id, name: (*s).to_owned(), ..Sheet::default() });
        }
        wb
    }

    fn sheet(&self, id: SheetId) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.id == id)
    }

    fn sheet_mut(&mut self, id: SheetId) -> Result<&mut Sheet> {
        self.sheets.iter_mut().find(|s| s.id == id).ok_or_else(|| err("no such sheet"))
    }

    fn cell_mut(&mut self, id: SheetId, row: u32, col: u32) -> Result<&mut Cell> {
        if row == 0 || col == 0 || row > MAX_ROWS || col > MAX_COLS {
            return Err(err("cell outside the sheet"));
        }
        let s = self.sheet_mut(id)?;
        let cell = Rect::cell(row, col);
        s.extent = Some(s.extent.map_or(cell, |e| e.bounding(&cell)));
        Ok(s.cells.entry((row, col)).or_default())
    }

    fn tidy(&mut self, id: SheetId, row: u32, col: u32) {
        if let Ok(s) = self.sheet_mut(id)
            && s.cells.get(&(row, col)).is_some_and(Cell::is_unused)
        {
            s.cells.remove(&(row, col));
        }
    }

    /// Sets the file name shown by `Workbook.Name`.
    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_owned();
    }

    /// Stores a formula together with its last computed value, as read from a file. Reading the
    /// cell still evaluates the formula; the cached value is returned when the built-in evaluator
    /// does not know a function in it (`#NAME?`), and by [`Self::cached_value`].
    pub fn set_cell_formula_cached(
        &mut self,
        sheet: SheetId,
        row: u32,
        col: u32,
        formula: String,
        cached: Value,
    ) -> Result<()> {
        let c = self.cell_mut(sheet, row, col)?;
        c.formula = Some(formula);
        c.value = cached;
        Ok(())
    }

    /// The stored value without evaluation: a constant, or a formula cell's cached result
    /// (`Empty` once the formula has been replaced through the [`Workbook`] trait).
    pub fn cached_value(&self, sheet: SheetId, row: u32, col: u32) -> Value {
        self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))).map(|c| c.value.clone()).unwrap_or_default()
    }

    /// Rows with an explicit height or hidden flag, sorted.
    pub fn custom_rows(&self, sheet: SheetId) -> Vec<u32> {
        let Some(s) = self.sheet(sheet) else { return Vec::new() };
        let mut v: Vec<u32> = s.row_heights.keys().chain(s.hidden_rows.iter()).copied().collect();
        v.sort_unstable();
        v.dedup();
        v
    }

    /// Columns with an explicit width or hidden flag, sorted.
    pub fn custom_columns(&self, sheet: SheetId) -> Vec<u32> {
        let Some(s) = self.sheet(sheet) else { return Vec::new() };
        let mut v: Vec<u32> = s.col_widths.keys().chain(s.hidden_cols.iter()).copied().collect();
        v.sort_unstable();
        v.dedup();
        v
    }

    /// Evaluates a formula cell now (a constant cell returns its value).
    fn live_value(&self, sheet: SheetId, row: u32, col: u32) -> Value {
        let Some(cell) = self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))) else { return Value::Empty };
        let Some(f) = &cell.formula else { return cell.value.clone() };
        if self.depth.get() > 64 {
            return Value::Error(2023);
        }
        self.depth.set(self.depth.get() + 1);
        let now = {
            let secs = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.as_secs_f64())
                .unwrap_or(0.0);
            macroost_lang::value::UNIX_EPOCH_SERIAL + secs / 86_400.0
        };
        let v = crate::eval::evaluate_cell(self, sheet, row, col, f, now);
        self.depth.set(self.depth.get() - 1);
        // A function the engine does not know: fall back to the value cached in the file.
        if matches!(v, Value::Error(2029)) && !cell.value.is_empty() { cell.value.clone() } else { v }
    }

    /// Stores the current value of a formula cell as its cached value.
    fn cache(&mut self, sheet: SheetId, row: u32, col: u32) {
        let v = self.live_value(sheet, row, col);
        if let Ok(s) = self.sheet_mut(sheet)
            && let Some(c) = s.cells.get_mut(&(row, col))
            && c.formula.is_some()
        {
            c.value = v;
        }
    }

    fn refresh_caches(&mut self) {
        let mut cells = Vec::new();
        for sh in &self.sheets {
            for (k, c) in &sh.cells {
                if c.formula.is_some() {
                    cells.push((sh.id, k.0, k.1));
                }
            }
        }
        let vals: Vec<Value> = cells.iter().map(|&(s, r, c)| self.live_value(s, r, c)).collect();
        for ((s, r, c), v) in cells.into_iter().zip(vals) {
            if let Ok(sh) = self.sheet_mut(s)
                && let Some(cell) = sh.cells.get_mut(&(r, c))
            {
                cell.value = v;
            }
        }
    }

    fn name_taken(&self, name: &str, except: Option<SheetId>) -> bool {
        self.sheets.iter().any(|s| Some(s.id) != except && s.name.to_lowercase() == name.to_lowercase())
    }
}

impl Workbook for MemoryWorkbook {
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }

    fn name(&self) -> String {
        self.name.clone()
    }

    fn sheet_ids(&self) -> Vec<SheetId> {
        self.sheets.iter().map(|s| s.id).collect()
    }

    fn sheet_name(&self, sheet: SheetId) -> Option<String> {
        self.sheet(sheet).map(|s| s.name.clone())
    }

    fn code_name(&self, sheet: SheetId) -> Option<String> {
        self.sheet(sheet).and_then(|s| s.code_name.clone())
    }

    fn set_code_name(&mut self, sheet: SheetId, code_name: &str) {
        if let Some(s) = self.sheets.iter_mut().find(|s| s.id == sheet) {
            s.code_name = Some(code_name.to_owned());
        }
    }

    fn add_sheet(&mut self, name: &str, position: usize) -> Result<SheetId> {
        if !valid_sheet_name(name) || self.name_taken(name, None) {
            return Err(err("That name is already taken or is not a valid sheet name."));
        }
        let id = self.next_id;
        self.next_id += 1;
        let pos = position.min(self.sheets.len());
        self.sheets.insert(pos, Sheet { id, name: name.to_owned(), ..Sheet::default() });
        Ok(id)
    }

    fn delete_sheet(&mut self, sheet: SheetId) -> Result<()> {
        let before = self.sheets.len();
        self.sheets.retain(|s| s.id != sheet);
        if self.sheets.len() == before {
            return Err(err("no such sheet"));
        }
        self.names.retain(|n| n.1 != Some(sheet));
        Ok(())
    }

    fn defined_names(&self) -> Vec<(String, Option<SheetId>, String)> {
        self.names.clone()
    }

    fn set_defined_name(&mut self, name: &str, scope: Option<SheetId>, refers_to: &str) {
        let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to);
        let pos = self.names.iter().position(|n| n.1 == scope && n.0.eq_ignore_ascii_case(name));
        match (pos, refers_to.is_empty()) {
            (Some(i), true) => {
                self.names.remove(i);
            }
            (Some(i), false) => self.names[i].2 = refers_to.to_owned(),
            (None, false) => self.names.push((name.to_owned(), scope, refers_to.to_owned())),
            (None, true) => {}
        }
    }

    fn rename_sheet(&mut self, sheet: SheetId, name: &str) -> Result<()> {
        if !valid_sheet_name(name) || self.name_taken(name, Some(sheet)) {
            return Err(err("That name is already taken or is not a valid sheet name."));
        }
        self.sheet_mut(sheet)?.name = name.to_owned();
        Ok(())
    }

    fn move_sheet(&mut self, sheet: SheetId, position: usize) -> Result<()> {
        let i = self.sheets.iter().position(|s| s.id == sheet).ok_or_else(|| err("no such sheet"))?;
        let s = self.sheets.remove(i);
        let pos = position.min(self.sheets.len());
        self.sheets.insert(pos, s);
        Ok(())
    }

    fn sheet_visibility(&self, sheet: SheetId) -> SheetVisibility {
        self.sheet(sheet).map(|s| s.visibility).unwrap_or_default()
    }

    fn set_sheet_visibility(&mut self, sheet: SheetId, visibility: SheetVisibility) -> Result<()> {
        self.sheet_mut(sheet)?.visibility = visibility;
        Ok(())
    }

    fn is_protected(&self, sheet: SheetId) -> bool {
        self.sheet(sheet).is_some_and(|s| s.protected)
    }

    fn set_protected(&mut self, sheet: SheetId, protected: bool) -> Result<()> {
        self.sheet_mut(sheet)?.protected = protected;
        Ok(())
    }

    /// A formula cell is evaluated on read with the small built-in evaluator (see `eval`).
    fn cell_value(&self, sheet: SheetId, row: u32, col: u32) -> Value {
        let Some(cell) = self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))) else { return Value::Empty };
        if cell.formula.is_none() || self.manual {
            return cell.value.clone();
        }
        self.live_value(sheet, row, col)
    }

    fn set_cell_value(&mut self, sheet: SheetId, row: u32, col: u32, value: Value) -> Result<()> {
        let c = self.cell_mut(sheet, row, col)?;
        c.value = value;
        c.formula = None;
        self.tidy(sheet, row, col);
        Ok(())
    }

    fn cell_formula(&self, sheet: SheetId, row: u32, col: u32) -> Option<String> {
        self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))).and_then(|c| c.formula.clone())
    }

    fn set_cell_formula(&mut self, sheet: SheetId, row: u32, col: u32, formula: Option<String>) -> Result<()> {
        let c = self.cell_mut(sheet, row, col)?;
        c.formula = formula;
        c.value = Value::Empty;
        self.tidy(sheet, row, col);
        if self.manual {
            // A formula is calculated when it is entered, even in manual mode.
            self.cache(sheet, row, col);
        }
        Ok(())
    }

    fn cell_format(&self, sheet: SheetId, row: u32, col: u32) -> CellFormat {
        let Some(s) = self.sheet(sheet) else { return CellFormat::default() };
        if let Some(c) = s.cells.get(&(row, col)) {
            return c.format.clone();
        }
        s.row_formats.get(&row).or_else(|| s.col_formats.get(&col)).cloned().unwrap_or_default()
    }

    fn row_filtered(&self, sheet: SheetId, row: u32) -> bool {
        self.sheet(sheet).is_some_and(|s| s.filtered_rows.contains(&row))
    }

    fn set_row_filtered(&mut self, sheet: SheetId, row: u32, filtered: bool) {
        if let Ok(s) = self.sheet_mut(sheet) {
            if filtered {
                s.filtered_rows.insert(row);
            } else {
                s.filtered_rows.remove(&row);
            }
        }
    }

    fn row_format(&self, sheet: SheetId, row: u32) -> Option<CellFormat> {
        self.sheet(sheet).and_then(|s| s.row_formats.get(&row).cloned())
    }

    fn set_row_format(&mut self, sheet: SheetId, row: u32, format: Option<CellFormat>) {
        if let Ok(s) = self.sheet_mut(sheet) {
            match format.filter(|f| !f.is_default()) {
                Some(f) => s.row_formats.insert(row, f),
                None => s.row_formats.remove(&row),
            };
        }
    }

    fn column_format(&self, sheet: SheetId, col: u32) -> Option<CellFormat> {
        self.sheet(sheet).and_then(|s| s.col_formats.get(&col).cloned())
    }

    fn set_column_format(&mut self, sheet: SheetId, col: u32, format: Option<CellFormat>) {
        if let Ok(s) = self.sheet_mut(sheet) {
            match format.filter(|f| !f.is_default()) {
                Some(f) => s.col_formats.insert(col, f),
                None => s.col_formats.remove(&col),
            };
        }
    }

    fn set_cell_format(&mut self, sheet: SheetId, row: u32, col: u32, format: CellFormat) -> Result<()> {
        self.cell_mut(sheet, row, col)?.format = format;
        self.tidy(sheet, row, col);
        Ok(())
    }

    fn occupied_cells(&self, sheet: SheetId, area: Rect) -> Vec<(u32, u32)> {
        let Some(s) = self.sheet(sheet) else { return Vec::new() };
        let mut out: Vec<(u32, u32)> = if area.area() <= s.cells.len() as u64 {
            let mut v = Vec::new();
            for r in area.r1..=area.r2 {
                for c in area.c1..=area.c2 {
                    if s.cells.contains_key(&(r, c)) {
                        v.push((r, c));
                    }
                }
            }
            v
        } else {
            s.cells.keys().copied().filter(|&(r, c)| area.contains(r, c)).collect()
        };
        out.sort_unstable();
        out
    }

    fn used_range(&self, sheet: SheetId) -> Option<Rect> {
        let s = self.sheet(sheet)?;
        s.cells.keys().fold(None, |acc: Option<Rect>, &(r, c)| {
            let cell = Rect::cell(r, c);
            Some(acc.map_or(cell, |a| a.bounding(&cell)))
        })
    }

    fn insert_rows(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        remap(&mut s.cells, |r, c| (r >= at).then(|| insert_shift(r, at, count, MAX_ROWS).map(|n| (n, c))));
        remap_index(&mut s.row_heights, |r| insert_shift(r, at, count, MAX_ROWS));
        remap_index(&mut s.row_formats, |r| insert_shift(r, at, count, MAX_ROWS));
        remap_set(&mut s.hidden_rows, |r| insert_shift(r, at, count, MAX_ROWS));
        remap_set(&mut s.filtered_rows, |r| insert_shift(r, at, count, MAX_ROWS));
        Ok(())
    }

    fn delete_rows(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        remap(&mut s.cells, |r, c| (r >= at).then(|| delete_shift(r, at, count).map(|n| (n, c))));
        remap_index(&mut s.row_heights, |r| delete_shift(r, at, count));
        remap_index(&mut s.row_formats, |r| delete_shift(r, at, count));
        remap_set(&mut s.hidden_rows, |r| delete_shift(r, at, count));
        remap_set(&mut s.filtered_rows, |r| delete_shift(r, at, count));
        Ok(())
    }

    fn insert_columns(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        remap(&mut s.cells, |r, c| (c >= at).then(|| insert_shift(c, at, count, MAX_COLS).map(|n| (r, n))));
        remap_index(&mut s.col_widths, |c| insert_shift(c, at, count, MAX_COLS));
        remap_index(&mut s.col_formats, |c| insert_shift(c, at, count, MAX_COLS));
        remap_set(&mut s.hidden_cols, |c| insert_shift(c, at, count, MAX_COLS));
        Ok(())
    }

    fn delete_columns(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        remap(&mut s.cells, |r, c| (c >= at).then(|| delete_shift(c, at, count).map(|n| (r, n))));
        remap_index(&mut s.col_widths, |c| delete_shift(c, at, count));
        remap_index(&mut s.col_formats, |c| delete_shift(c, at, count));
        remap_set(&mut s.hidden_cols, |c| delete_shift(c, at, count));
        Ok(())
    }

    fn insert_cells(&mut self, sheet: SheetId, area: Rect, shift: ShiftDirection) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        match shift {
            ShiftDirection::Down | ShiftDirection::Up => {
                let n = area.rows();
                remap(&mut s.cells, |r, c| {
                    (c >= area.c1 && c <= area.c2 && r >= area.r1)
                        .then(|| insert_shift(r, area.r1, n, MAX_ROWS).map(|nr| (nr, c)))
                });
            }
            ShiftDirection::Right | ShiftDirection::Left => {
                let n = area.cols();
                remap(&mut s.cells, |r, c| {
                    (r >= area.r1 && r <= area.r2 && c >= area.c1)
                        .then(|| insert_shift(c, area.c1, n, MAX_COLS).map(|nc| (r, nc)))
                });
            }
        }
        Ok(())
    }

    fn delete_cells(&mut self, sheet: SheetId, area: Rect, shift: ShiftDirection) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        match shift {
            ShiftDirection::Up | ShiftDirection::Down => {
                let n = area.rows();
                remap(&mut s.cells, |r, c| {
                    (c >= area.c1 && c <= area.c2 && r >= area.r1)
                        .then(|| delete_shift(r, area.r1, n).map(|nr| (nr, c)))
                });
            }
            ShiftDirection::Left | ShiftDirection::Right => {
                let n = area.cols();
                remap(&mut s.cells, |r, c| {
                    (r >= area.r1 && r <= area.r2 && c >= area.c1)
                        .then(|| delete_shift(c, area.c1, n).map(|nc| (r, nc)))
                });
            }
        }
        Ok(())
    }

    fn row_hidden(&self, sheet: SheetId, row: u32) -> bool {
        self.sheet(sheet).is_some_and(|s| s.hidden_rows.contains(&row))
    }

    fn set_row_hidden(&mut self, sheet: SheetId, row: u32, hidden: bool) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        if hidden {
            s.hidden_rows.insert(row);
        } else {
            s.hidden_rows.remove(&row);
        }
        Ok(())
    }

    fn row_height(&self, sheet: SheetId, row: u32) -> f64 {
        self.sheet(sheet).and_then(|s| s.row_heights.get(&row).copied()).unwrap_or(DEFAULT_ROW_HEIGHT)
    }

    fn set_row_height(&mut self, sheet: SheetId, row: u32, height: f64) -> Result<()> {
        if !(0.0..=409.0).contains(&height) {
            return Err(err("Unable to set the RowHeight property of the Range class"));
        }
        self.sheet_mut(sheet)?.row_heights.insert(row, height);
        Ok(())
    }

    fn column_hidden(&self, sheet: SheetId, col: u32) -> bool {
        self.sheet(sheet).is_some_and(|s| s.hidden_cols.contains(&col))
    }

    fn set_column_hidden(&mut self, sheet: SheetId, col: u32, hidden: bool) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        if hidden {
            s.hidden_cols.insert(col);
        } else {
            s.hidden_cols.remove(&col);
        }
        Ok(())
    }

    fn column_width(&self, sheet: SheetId, col: u32) -> f64 {
        self.sheet(sheet).and_then(|s| s.col_widths.get(&col).copied()).unwrap_or(DEFAULT_COLUMN_WIDTH)
    }

    fn set_column_width(&mut self, sheet: SheetId, col: u32, width: f64) -> Result<()> {
        if !(0.0..=255.0).contains(&width) {
            return Err(err("Unable to set the ColumnWidth property of the Range class"));
        }
        self.sheet_mut(sheet)?.col_widths.insert(col, width);
        Ok(())
    }

    fn merge_cells(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        s.merges.retain(|m| m.intersect(&area).is_none());
        if !area.is_cell() {
            s.merges.push(area);
        }
        Ok(())
    }

    fn unmerge_cells(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        let s = self.sheet_mut(sheet)?;
        s.merges.retain(|m| m.intersect(&area).is_none());
        Ok(())
    }

    fn merged_areas(&self, sheet: SheetId) -> Vec<Rect> {
        self.sheet(sheet).map(|s| s.merges.clone()).unwrap_or_default()
    }

    fn last_cell(&self, sheet: SheetId) -> Option<Rect> {
        let used = self.used_range(sheet);
        let ext = self.sheet(sheet).and_then(|s| s.extent);
        match (used, ext) {
            (Some(u), Some(e)) => Some(u.bounding(&e)),
            (u, e) => u.or(e),
        }
    }

    fn reset_last_cell(&mut self, sheet: SheetId) {
        let used = self.used_range(sheet);
        if let Ok(s) = self.sheet_mut(sheet) {
            s.extent = used;
        }
    }

    fn sheet_password(&self, sheet: SheetId) -> Option<PasswordHash> {
        self.sheet(sheet).and_then(|s| s.password.clone())
    }

    fn set_sheet_password(&mut self, sheet: SheetId, password: Option<PasswordHash>) {
        if let Ok(s) = self.sheet_mut(sheet) {
            s.password = password;
        }
    }

    fn workbook_protection(&self) -> Option<WorkbookProtection> {
        self.protection.clone()
    }

    fn set_workbook_protection(&mut self, protection: Option<WorkbookProtection>) {
        self.protection = protection;
    }

    fn cell_comment(&self, sheet: SheetId, row: u32, col: u32) -> Option<String> {
        self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))).and_then(|c| c.comment.clone())
    }

    fn set_cell_comment(&mut self, sheet: SheetId, row: u32, col: u32, text: Option<String>) -> Result<()> {
        self.cell_mut(sheet, row, col)?.comment = text;
        self.tidy(sheet, row, col);
        Ok(())
    }

    fn cell_hyperlink(&self, sheet: SheetId, row: u32, col: u32) -> Option<Hyperlink> {
        self.sheet(sheet).and_then(|s| s.cells.get(&(row, col))).and_then(|c| c.link.clone())
    }

    fn set_cell_hyperlink(&mut self, sheet: SheetId, row: u32, col: u32, link: Option<Hyperlink>) -> Result<()> {
        self.cell_mut(sheet, row, col)?.link = link;
        self.tidy(sheet, row, col);
        Ok(())
    }

    fn set_manual_calculation(&mut self, manual: bool) {
        if self.manual && !manual {
            self.manual = false;
        } else if manual && !self.manual {
            self.refresh_caches();
            self.manual = true;
        }
    }

    fn calculate(&mut self) {
        self.calculations += 1;
        if self.manual {
            self.manual = false;
            self.refresh_caches();
            self.manual = true;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn insert_and_delete_rows_shift_cells() {
        let mut wb = MemoryWorkbook::new();
        let id = wb.sheet_ids()[0];
        wb.set_cell_value(id, 2, 1, Value::Double(2.0)).ok();
        wb.set_cell_value(id, MAX_ROWS, 1, Value::Double(9.0)).ok();
        wb.insert_rows(id, 1, 1).ok();
        assert!(matches!(wb.cell_value(id, 3, 1), Value::Double(n) if n == 2.0));
        assert!(wb.cell_value(id, MAX_ROWS, 1).is_empty(), "pushed off the end");
        wb.delete_rows(id, 1, 2).ok();
        assert!(matches!(wb.cell_value(id, 1, 1), Value::Double(n) if n == 2.0));
        assert_eq!(wb.used_range(id), Some(Rect::cell(1, 1)));
    }

    #[test]
    fn names_are_validated() {
        let mut wb = MemoryWorkbook::new();
        assert!(wb.add_sheet("sheet1", 1).is_err());
        assert!(wb.add_sheet("a/b", 1).is_err());
        assert!(wb.add_sheet("Data", 1).is_ok());
        assert_eq!(wb.sheet_ids().len(), 2);
    }
}
