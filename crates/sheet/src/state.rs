//! Application state shared by every object (`Rc<RefCell<State>>`), plus the cell-level rules
//! the object model applies on top of the storage: how assigned values are coerced, how values
//! read back (dates, currency), protection, clipboard snapshots and paste.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use macroost_lang::value::{format_float, parse_number, ymd_to_serial};
use macroost_lang::{Args, Result, RuntimeError, Value};

use crate::address::{MAX_COLS, MAX_ROWS, Rect, column_number};
use crate::format::{error_code, format_value, is_currency_format, is_date_format};
use crate::formula::shift_formula;
use crate::workbook::{CellFormat, SheetId, SheetVisibility, Workbook};

/// Ranges larger than this many cells are never walked cell by cell; writes are clipped to the
/// used range and reads visit only stored cells.
pub(crate) const LARGE: u64 = 1 << 17;

pub(crate) type Shared = Rc<RefCell<State>>;

#[derive(Clone, Debug)]
pub(crate) struct SheetView {
    pub selection: Vec<Rect>,
    pub active: (u32, u32),
}

impl Default for SheetView {
    fn default() -> Self {
        Self { selection: vec![Rect::cell(1, 1)], active: (1, 1) }
    }
}

#[derive(Clone, Debug)]
pub(crate) struct ClipCell {
    pub dr: u32,
    pub dc: u32,
    pub value: Value,
    pub formula: Option<String>,
    pub format: CellFormat,
    pub comment: Option<String>,
}

/// A copied block: a snapshot of the stored cells of `src` on `sheet`.
#[derive(Clone, Debug)]
pub(crate) struct Clipboard {
    pub sheet: SheetId,
    pub src: Rect,
    pub cells: Vec<ClipCell>,
    pub cut: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum PasteKind {
    All,
    Values,
    Formats,
    Formulas,
    ValuesAndNumberFormats,
    AllExceptBorders,
}

impl PasteKind {
    pub fn from_constant(n: i64) -> Result<Self> {
        Ok(match n {
            -4104 => PasteKind::All,
            -4163 => PasteKind::Values,
            -4122 => PasteKind::Formats,
            -4123 => PasteKind::Formulas,
            12 => PasteKind::ValuesAndNumberFormats,
            11 => PasteKind::Formulas, // xlPasteFormulasAndNumberFormats
            7 => PasteKind::AllExceptBorders,
            13 | 8 => PasteKind::All, // xlPasteAllUsingSourceTheme, xlPasteColumnWidths (approx.)
            _ => return Err(err1004("PasteSpecial method of Range class failed")),
        })
    }
}

#[derive(Clone, Debug)]
pub(crate) struct FindParams {
    pub what: String,
    pub look_in: i32,
    pub look_at: i32,
    pub match_case: bool,
    pub by_columns: bool,
}

/// What a sheet's protection covers (`Worksheet.Protect` arguments).
#[derive(Clone, Debug)]
pub(crate) struct Protection {
    pub contents: bool,
    pub drawing: bool,
    pub scenarios: bool,
    /// AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns,
    /// AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows,
    /// AllowSorting, AllowFiltering, AllowUsingPivotTables.
    pub allow: [bool; 11],
}

impl Default for Protection {
    fn default() -> Self {
        Self { contents: true, drawing: true, scenarios: true, allow: [false; 11] }
    }
}

pub(crate) const ALLOW_NAMES: [&str; 11] = [
    "allowformattingcells",
    "allowformattingcolumns",
    "allowformattingrows",
    "allowinsertingcolumns",
    "allowinsertingrows",
    "allowinsertinghyperlinks",
    "allowdeletingcolumns",
    "allowdeletingrows",
    "allowsorting",
    "allowfiltering",
    "allowusingpivottables",
];

/// Settings `Range.Sort` remembers between calls (Order1..3, Orientation, Header, MatchCase).
#[derive(Clone, Debug)]
pub(crate) struct SortDefaults {
    pub order: [i64; 3],
    pub orientation: i64,
    pub header: i64,
    pub match_case: bool,
}

impl Default for SortDefaults {
    fn default() -> Self {
        Self { order: [1, 1, 1], orientation: 1, header: 2, match_case: false }
    }
}

/// The `Worksheet.Sort` object's settings.
#[derive(Clone, Debug)]
pub(crate) struct SheetSort {
    pub fields: Vec<(Rect, bool)>,
    pub range: Option<Rect>,
    pub header: i64,
    pub match_case: bool,
    pub orientation: i64,
}

impl Default for SheetSort {
    fn default() -> Self {
        Self { fields: Vec::new(), range: None, header: 2, match_case: false, orientation: 1 }
    }
}

/// One AutoFilter column condition.
#[derive(Clone, Debug)]
pub(crate) struct FilterCond {
    pub field: u32,
    pub crit1: Option<Value>,
    pub crit2: Option<Value>,
    pub operator: i64,
}

#[derive(Clone, Debug)]
pub(crate) struct AutoFilter {
    pub range: Rect,
    pub conds: Vec<FilterCond>,
}

/// `ActiveWindow` settings remembered without a display.
#[derive(Clone, Debug)]
pub(crate) struct WindowState {
    pub scroll_row: i32,
    pub scroll_col: i32,
    pub freeze: Option<(u32, u32)>,
    pub zoom: Value,
    pub gridlines: bool,
    pub headings: bool,
}

impl Default for WindowState {
    fn default() -> Self {
        Self { scroll_row: 1, scroll_col: 1, freeze: None, zoom: Value::Long(100), gridlines: true, headings: true }
    }
}

/// Everything the object model remembers between calls.
pub struct State {
    pub(crate) book: Box<dyn Workbook>,
    pub(crate) screen_updating: bool,
    pub(crate) display_alerts: bool,
    pub(crate) enable_events: bool,
    pub(crate) calculation: i32,
    pub(crate) status_bar: Value,
    /// Application properties accepted and remembered without further effect.
    pub(crate) settings: HashMap<String, Value>,
    pub(crate) active_sheet: SheetId,
    pub(crate) views: HashMap<SheetId, SheetView>,
    pub(crate) clipboard: Option<Clipboard>,
    pub(crate) protection: HashMap<SheetId, Protection>,
    /// `UserInterfaceOnly` protection; it outlives `Unprotect` for the session.
    pub(crate) ui_only: HashMap<SheetId, bool>,
    /// Per sheet: the settings `Range.Sort` remembers.
    pub(crate) sort_defaults: HashMap<SheetId, SortDefaults>,
    pub(crate) filters: HashMap<SheetId, AutoFilter>,
    pub(crate) windows: HashMap<SheetId, WindowState>,
    pub(crate) sheet_sorts: HashMap<SheetId, SheetSort>,
    pub(crate) visible_comments: std::collections::HashSet<(SheetId, u32, u32)>,
    pub(crate) last_find: Option<FindParams>,
    pub(crate) sheet_counter: u32,
    pub(crate) saved: bool,
    /// Actions with no storage effect (Save, AutoFilter, Merge, ...), in order.
    pub(crate) log: Vec<String>,
}

pub(crate) fn err1004(msg: &str) -> RuntimeError {
    RuntimeError::new(1004, msg)
}

pub(crate) fn busy() -> RuntimeError {
    RuntimeError::new(1004, "the object model is busy (re-entrant call)")
}

pub(crate) fn with<T>(app: &Shared, f: impl FnOnce(&State) -> Result<T>) -> Result<T> {
    let st = app.try_borrow().map_err(|_| busy())?;
    f(&st)
}

pub(crate) fn with_mut<T>(app: &Shared, f: impl FnOnce(&mut State) -> Result<T>) -> Result<T> {
    let mut st = app.try_borrow_mut().map_err(|_| busy())?;
    f(&mut st)
}

/// What an assigned value becomes in a cell.
#[derive(Debug, Clone)]
pub(crate) enum Input {
    Clear,
    Const(Value, Option<&'static str>),
    Formula(String),
}

fn is_number_text(t: &str) -> bool {
    let body = t.strip_prefix(['-', '+']).unwrap_or(t);
    if body.is_empty() {
        return false;
    }
    let (mantissa, exp) = match body.find(['e', 'E']) {
        Some(i) => (&body[..i], Some(&body[i + 1..])),
        None => (body, None),
    };
    if let Some(e) = exp {
        let e = e.strip_prefix(['-', '+']).unwrap_or(e);
        if e.is_empty() || !e.bytes().all(|b| b.is_ascii_digit()) {
            return false;
        }
    }
    let (int, frac) = match mantissa.split_once('.') {
        Some((i, f)) => (i, Some(f)),
        None => (mantissa, None),
    };
    if let Some(f) = frac
        && !f.bytes().all(|b| b.is_ascii_digit())
    {
        return false;
    }
    if int.is_empty() {
        return frac.is_some_and(|f| !f.is_empty());
    }
    if int.contains(',') {
        // Thousands separators must group by three.
        let groups: Vec<&str> = int.split(',').collect();
        let first_ok = groups.first().is_some_and(|g| (1..=3).contains(&g.len()));
        return first_ok
            && groups.iter().all(|g| g.bytes().all(|b| b.is_ascii_digit()))
            && groups.iter().skip(1).all(|g| g.len() == 3);
    }
    int.bytes().all(|b| b.is_ascii_digit())
}

const MONTH_NAMES: [&str; 12] = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

fn month_of(word: &str) -> Option<i64> {
    let w = word.trim().trim_end_matches('.').to_ascii_lowercase();
    if w.len() < 3 {
        return None;
    }
    let full = [
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
    ];
    MONTH_NAMES
        .iter()
        .position(|m| w.starts_with(m) && full.iter().any(|f| f.starts_with(w.as_str())))
        .map(|i| i as i64 + 1)
}

fn current_year() -> i64 {
    let secs =
        std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs_f64()).unwrap_or(0.0);
    macroost_lang::value::serial_to_ymd(macroost_lang::value::UNIX_EPOCH_SERIAL + secs / 86_400.0).0
}

fn full_year(y: i64) -> i64 {
    if y < 100 { y + if y < 30 { 2000 } else { 1900 } } else { y }
}

/// A typed date (without time): serial and the number format the application applies.
fn parse_date_text(t: &str) -> Option<(f64, &'static str)> {
    let nums = |s: &str, sep: char| -> Option<Vec<i64>> {
        let parts: Vec<&str> = s.split(sep).collect();
        if !(2..=3).contains(&parts.len())
            || parts.iter().any(|p| p.is_empty() || !p.bytes().all(|b| b.is_ascii_digit()))
        {
            return None;
        }
        parts.iter().map(|p| p.parse().ok()).collect()
    };
    let ok = |m: i64, d: i64| (1..=12).contains(&m) && (1..=31).contains(&d);
    if let Some(v) = nums(t, '/') {
        if v.len() == 3 {
            let (m, d, y) = (v[0], v[1], full_year(v[2]));
            if ok(m, d) {
                return Some((ymd_to_serial(y, m, d), "m/d/yyyy"));
            }
        } else if (1..=12).contains(&v[0]) && v[1] > 31 {
            // Month/year.
            return Some((ymd_to_serial(full_year(v[1]), v[0], 1), "mmm-yy"));
        } else if ok(v[0], v[1]) {
            // Month/day of the current year.
            return Some((ymd_to_serial(current_year(), v[0], v[1]), "d-mmm"));
        }
    }
    if let Some(v) = nums(t, '-')
        && v.len() == 2
    {
        if (1..=12).contains(&v[0]) && v[1] > 31 {
            return Some((ymd_to_serial(full_year(v[1]), v[0], 1), "mmm-yy"));
        }
        if ok(v[0], v[1]) {
            return Some((ymd_to_serial(current_year(), v[0], v[1]), "d-mmm"));
        }
    }
    if let Some(v) = nums(t, '-')
        && v.len() == 3
    {
        let (y, m, d) = (v[0], v[1], v[2]);
        if y >= 1900 && ok(m, d) {
            return Some((ymd_to_serial(y, m, d), "m/d/yyyy"));
        }
        let (m, d, y) = (v[0], v[1], full_year(v[2]));
        if ok(m, d) {
            return Some((ymd_to_serial(y, m, d), "m/d/yyyy"));
        }
    }
    // Month names: "Jan 5, 2026", "5-Mar-2026", "5-Mar", "Mar-2026", "March 5".
    let words: Vec<&str> = t.split([' ', '-', ',', '/']).filter(|w| !w.is_empty()).collect();
    let num =
        |w: &str| -> Option<i64> { (w.bytes().all(|b| b.is_ascii_digit()) && !w.is_empty()).then(|| w.parse().ok())? };
    match words.as_slice() {
        [a, b, c] => {
            if let (Some(m), Some(d), Some(y)) = (month_of(a), num(b), num(c))
                && (1..=31).contains(&d)
            {
                return Some((ymd_to_serial(full_year(y), m, d), "d-mmm-yy"));
            }
            if let (Some(d), Some(m), Some(y)) = (num(a), month_of(b), num(c))
                && (1..=31).contains(&d)
            {
                return Some((ymd_to_serial(full_year(y), m, d), "d-mmm-yy"));
            }
        }
        [a, b] => {
            if let (Some(d), Some(m)) = (num(a), month_of(b))
                && (1..=31).contains(&d)
            {
                return Some((ymd_to_serial(current_year(), m, d), "d-mmm"));
            }
            if let (Some(m), Some(n)) = (month_of(a), num(b)) {
                if (1..=31).contains(&n) && b.len() <= 2 {
                    return Some((ymd_to_serial(current_year(), m, n), "d-mmm"));
                }
                return Some((ymd_to_serial(full_year(n), m, 1), "mmm-yy"));
            }
        }
        _ => {}
    }
    None
}

fn parse_time_text(t: &str) -> Option<(f64, &'static str)> {
    let upper = t.to_ascii_uppercase();
    let (body, ampm) = if let Some(b) = upper.strip_suffix("AM") {
        (b.trim().to_owned(), Some(false))
    } else if let Some(b) = upper.strip_suffix("PM") {
        (b.trim().to_owned(), Some(true))
    } else {
        (upper.clone(), None)
    };
    let parts: Vec<&str> = body.split(':').collect();
    if !(2..=3).contains(&parts.len()) || parts.iter().any(|p| p.is_empty() || !p.bytes().all(|b| b.is_ascii_digit())) {
        return None;
    }
    let mut h: u32 = parts[0].parse().ok()?;
    let m: u32 = parts[1].parse().ok()?;
    let s: u32 = parts.get(2).map(|p| p.parse().ok()).unwrap_or(Some(0))?;
    if m > 59 || s > 59 {
        return None;
    }
    if let Some(pm) = ampm {
        if h == 0 || h > 12 {
            return None;
        }
        h = match (h, pm) {
            (12, false) => 0,
            (12, true) => 12,
            (h, true) => h + 12,
            (h, false) => h,
        };
    } else if h > 23 {
        return None;
    }
    let fmt = match (ampm.is_some(), parts.len()) {
        (true, 3) => "h:mm:ss AM/PM",
        (true, _) => "h:mm AM/PM",
        (false, 3) => "h:mm:ss",
        (false, _) => "h:mm",
    };
    Some((f64::from(h * 3600 + m * 60 + s) / 86_400.0, fmt))
}

/// Interprets text assigned to a cell the way typing it would: formulas, numbers, percentages,
/// currency, fractions, dates, times, booleans and error literals; a leading apostrophe forces
/// text.
pub(crate) fn parse_typed(s: &str) -> Input {
    if s.is_empty() {
        return Input::Clear;
    }
    if let Some(rest) = s.strip_prefix('\'') {
        return Input::Const(Value::str(rest), None);
    }
    if s.len() > 1 && s.starts_with('=') {
        return Input::Formula(s.to_owned());
    }
    let t = s.trim();
    if t.eq_ignore_ascii_case("true") {
        return Input::Const(Value::Boolean(true), None);
    }
    if t.eq_ignore_ascii_case("false") {
        return Input::Const(Value::Boolean(false), None);
    }
    if let Some(code) = error_code(t) {
        return Input::Const(Value::Error(code), None);
    }
    if let Some(p) = t.strip_suffix('%')
        && is_number_text(p.trim())
        && let Some(n) = parse_number(p)
    {
        let fmt = if p.contains('.') { "0.00%" } else { "0%" };
        return Input::Const(Value::Double(n / 100.0), Some(fmt));
    }
    // (42) is -42.
    if let Some(inner) = t.strip_prefix('(').and_then(|r| r.strip_suffix(')'))
        && is_number_text(inner.trim())
        && let Some(n) = parse_number(inner)
    {
        return Input::Const(Value::Double(-n), None);
    }
    let (neg, unsigned) = match t.strip_prefix('-') {
        Some(r) => (true, r),
        None => (false, t),
    };
    if let Some(c) = unsigned.strip_prefix('$')
        && is_number_text(c)
        && let Some(n) = parse_number(c)
    {
        let fmt = if c.contains('.') { "$#,##0.00_);[Red]($#,##0.00)" } else { "$#,##0_);[Red]($#,##0)" };
        return Input::Const(Value::Double(if neg { -n } else { n }), Some(fmt));
    }
    if is_number_text(t)
        && let Some(n) = parse_number(t)
    {
        let fmt = if t.contains(['e', 'E']) {
            Some("0.00E+00")
        } else if t.contains(',') {
            Some(if t.contains('.') { "#,##0.00" } else { "#,##0" })
        } else {
            None
        };
        return Input::Const(Value::Double(n), fmt);
    }
    // Mixed fraction "1 1/4" / "0 1/2".
    if let Some((whole, frac)) = t.split_once(' ')
        && let Some((a, b)) = frac.split_once('/')
        && !whole.is_empty()
        && whole.bytes().all(|x| x.is_ascii_digit())
        && !a.is_empty()
        && a.bytes().all(|x| x.is_ascii_digit())
        && !b.is_empty()
        && b.bytes().all(|x| x.is_ascii_digit())
    {
        let (w, n, d): (f64, f64, f64) =
            (whole.parse().unwrap_or(0.0), a.parse().unwrap_or(0.0), b.parse().unwrap_or(1.0));
        if d > 0.0 && n < d {
            let fmt = if d >= 10.0 { "# ??/??" } else { "# ?/?" };
            return Input::Const(Value::Double(w + n / d), Some(fmt));
        }
    }
    // Date with a time: "1/2/2026 10:30".
    if let Some((date, time)) = t.split_once(' ')
        && let Some((d, _)) = parse_date_text(date)
        && let Some((tm, _)) = parse_time_text(time.trim())
    {
        return Input::Const(Value::Double(d + tm), Some("m/d/yyyy h:mm"));
    }
    if let Some((d, fmt)) = parse_date_text(t) {
        return Input::Const(Value::Double(d), Some(fmt));
    }
    if let Some((x, fmt)) = parse_time_text(t) {
        return Input::Const(Value::Double(x), Some(fmt));
    }
    Input::Const(Value::str(s), None)
}

/// How a macro value becomes cell content. `text_format` is true for a cell formatted as text (`@`).
pub(crate) fn cell_input(v: &Value, text_format: bool) -> Result<Input> {
    Ok(match v {
        Value::Empty | Value::Null => Input::Clear,
        Value::Boolean(b) => Input::Const(Value::Boolean(*b), None),
        Value::Byte(_)
        | Value::Integer(_)
        | Value::Long(_)
        | Value::LongLong(_)
        | Value::Single(_)
        | Value::Double(_) => Input::Const(Value::Double(v.to_f64()?), None),
        Value::Currency(_) => Input::Const(Value::Double(v.to_f64()?), Some("$#,##0.00_);($#,##0.00)")),
        Value::Date(d) => {
            let fmt = if d.fract() == 0.0 {
                "m/d/yyyy"
            } else if d.trunc() == 0.0 {
                "h:mm:ss AM/PM"
            } else {
                "m/d/yyyy h:mm"
            };
            Input::Const(Value::Double(*d), Some(fmt))
        }
        Value::String(s) => {
            if text_format && !s.is_empty() {
                Input::Const(Value::String(s.clone()), None)
            } else {
                parse_typed(s)
            }
        }
        Value::Error(n) => Input::Const(Value::Error(*n), None),
        Value::Object(None) => return Err(RuntimeError::object_not_set()),
        Value::Object(Some(o)) => return cell_input(&o.default_value()?, text_format),
        Value::Array(a) => {
            let first = a.borrow().data.first().cloned().unwrap_or_default();
            return cell_input(&first, text_format);
        }
    })
}

/// Formula-bar text of a constant (`Range.Formula` on a cell without a formula).
pub(crate) fn constant_formula_text(v: &Value, fmt: &CellFormat) -> String {
    match v {
        Value::Empty => String::new(),
        Value::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_owned(),
        Value::String(s) => s.to_string(),
        Value::Error(n) => crate::format::error_text(*n).to_owned(),
        other => match other.to_f64() {
            Ok(x) => {
                let nf = fmt.number_format.as_deref().unwrap_or("General");
                if is_date_format(nf) {
                    let df = if x.fract() == 0.0 {
                        "m/d/yyyy"
                    } else if x.trunc() == 0.0 {
                        "h:mm:ss AM/PM"
                    } else {
                        "m/d/yyyy h:mm:ss AM/PM"
                    };
                    format_value(&Value::Double(x), df)
                } else {
                    format_float(x, 15)
                }
            }
            Err(_) => String::new(),
        },
    }
}

impl State {
    pub(crate) fn new(mut book: Box<dyn Workbook>) -> Self {
        if book.sheet_ids().is_empty() {
            let _ = book.add_sheet("Sheet1", 0);
        }
        let ids = book.sheet_ids();
        let active = ids
            .iter()
            .copied()
            .find(|&id| book.sheet_visibility(id) == SheetVisibility::Visible)
            .or_else(|| ids.first().copied())
            .unwrap_or(0);
        let counter = ids
            .iter()
            .filter_map(|&id| book.sheet_name(id))
            .filter_map(|n| n.strip_prefix("Sheet").and_then(|d| d.parse::<u32>().ok()))
            .max()
            .unwrap_or(0);
        let mut protection = HashMap::new();
        for &id in &ids {
            if book.is_protected(id) {
                protection.insert(id, Protection::default());
            }
        }
        Self {
            book,
            screen_updating: true,
            display_alerts: true,
            enable_events: true,
            calculation: -4105,
            status_bar: Value::Boolean(false),
            settings: HashMap::new(),
            active_sheet: active,
            views: HashMap::new(),
            clipboard: None,
            protection,
            ui_only: HashMap::new(),
            sort_defaults: HashMap::new(),
            filters: HashMap::new(),
            windows: HashMap::new(),
            sheet_sorts: HashMap::new(),
            visible_comments: std::collections::HashSet::new(),
            last_find: None,
            sheet_counter: counter,
            saved: true,
            log: Vec::new(),
        }
    }

    pub(crate) fn sheet_exists(&self, id: SheetId) -> bool {
        self.book.sheet_ids().contains(&id)
    }

    pub(crate) fn sheet_name(&self, id: SheetId) -> Result<String> {
        self.book.sheet_name(id).ok_or_else(RuntimeError::object_required)
    }

    pub(crate) fn sheet_by_name(&self, name: &str) -> Option<SheetId> {
        let lower = name.to_lowercase();
        self.book
            .sheet_ids()
            .into_iter()
            .find(|&id| self.book.sheet_name(id).is_some_and(|n| n.to_lowercase() == lower))
    }

    /// 1-based tab position.
    pub(crate) fn sheet_index(&self, id: SheetId) -> Option<usize> {
        self.book.sheet_ids().iter().position(|&s| s == id).map(|i| i + 1)
    }

    pub(crate) fn view(&self, id: SheetId) -> SheetView {
        self.views.get(&id).cloned().unwrap_or_default()
    }

    /// Whether macro edits to the sheet's cells are blocked by protection.
    pub(crate) fn contents_protected(&self, sheet: SheetId) -> bool {
        self.book.is_protected(sheet)
            && !self.ui_only.get(&sheet).copied().unwrap_or(false)
            && self.protection.get(&sheet).is_none_or(|p| p.contents)
    }

    pub(crate) fn protected_error(&self, sheet: SheetId) -> RuntimeError {
        let _ = sheet;
        err1004(
            "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.",
        )
    }

    /// Whole-sheet edits (insert, delete, sort...) on a protected sheet fail.
    pub(crate) fn check_writable(&self, sheet: SheetId) -> Result<()> {
        if self.contents_protected(sheet) {
            return Err(self.protected_error(sheet));
        }
        Ok(())
    }

    /// An edit of one cell: locked cells of a protected sheet refuse it.
    pub(crate) fn check_cell(&self, sheet: SheetId, r: u32, c: u32) -> Result<()> {
        if self.contents_protected(sheet) && !self.book.cell_format(sheet, r, c).unlocked {
            return Err(self.protected_error(sheet));
        }
        Ok(())
    }

    /// Format changes on a protected sheet need AllowFormattingCells.
    pub(crate) fn check_format(&self, sheet: SheetId) -> Result<()> {
        if self.contents_protected(sheet) && !self.protection.get(&sheet).is_some_and(|p| p.allow[0]) {
            return Err(err1004("Application-defined or object-defined error"));
        }
        Ok(())
    }

    /// True when the workbook structure is protected (sheets cannot be added, renamed, moved...).
    pub(crate) fn structure_protected(&self) -> bool {
        self.book.workbook_protection().is_some_and(|p| p.structure)
    }

    pub(crate) fn is_blank(&self, sheet: SheetId, r: u32, c: u32) -> bool {
        self.book.cell_value(sheet, r, c).is_empty() && self.book.cell_formula(sheet, r, c).is_none()
    }

    /// Stored cells inside `area` that hold a value or formula.
    pub(crate) fn content_cells(&self, sheet: SheetId, area: Rect) -> Vec<(u32, u32)> {
        self.book.occupied_cells(sheet, area).into_iter().filter(|&(r, c)| !self.is_blank(sheet, r, c)).collect()
    }

    /// `UsedRange` (A1 for an empty sheet).
    pub(crate) fn used(&self, sheet: SheetId) -> Rect {
        self.book.used_range(sheet).unwrap_or(Rect::cell(1, 1))
    }

    /// The part of `area` worth visiting cell by cell.
    pub(crate) fn clip(&self, sheet: SheetId, area: Rect) -> Option<Rect> {
        if area.area() <= LARGE {
            return Some(area);
        }
        // From the area's own start to the end of the used range.
        let u = self.book.used_range(sheet)?;
        let (r2, c2) = (area.r2.min(u.r2), area.c2.min(u.c2));
        (r2 >= area.r1 && c2 >= area.c1).then(|| Rect::new(area.r1, area.c1, r2, c2))
    }

    /// Cells a write over `area` touches: all of them, or for a huge area only the stored ones
    /// (so `Columns("A").Font.Bold = True` never creates a million cells).
    pub(crate) fn write_cells(&self, sheet: SheetId, area: Rect) -> Vec<(u32, u32)> {
        if area.area() <= LARGE {
            let mut v = Vec::with_capacity(area.area() as usize);
            for r in area.r1..=area.r2 {
                for c in area.c1..=area.c2 {
                    v.push((r, c));
                }
            }
            v
        } else {
            self.book.occupied_cells(sheet, area)
        }
    }

    pub(crate) fn number_format(&self, sheet: SheetId, r: u32, c: u32) -> String {
        self.book.cell_format(sheet, r, c).number_format.unwrap_or_else(|| "General".to_owned())
    }

    /// `Range.Value` (or `Value2`) of one cell.
    pub(crate) fn read_value(&self, sheet: SheetId, r: u32, c: u32, value2: bool) -> Value {
        let v = self.book.cell_value(sheet, r, c);
        if value2 {
            return match v {
                Value::Date(d) => Value::Double(d),
                Value::Currency(n) => Value::Double(n as f64 / 10_000.0),
                other => other,
            };
        }
        match v {
            Value::Double(x) => {
                let nf = self.number_format(sheet, r, c);
                if crate::format::has_date_part(&nf) && (0.0..2_958_466.0).contains(&x) {
                    Value::Date(x)
                } else if is_currency_format(&nf) && x.abs() < 9.2e14 {
                    Value::Currency((x * 10_000.0).round() as i64)
                } else {
                    Value::Double(x)
                }
            }
            other => other,
        }
    }

    pub(crate) fn text(&self, sheet: SheetId, r: u32, c: u32) -> String {
        let v = self.book.cell_value(sheet, r, c);
        let nf = self.number_format(sheet, r, c);
        if nf == "@" && !matches!(v, Value::String(_)) {
            return format_value(&v, "General");
        }
        crate::format::format_value_width(&v, &nf, self.book.column_width(sheet, c))
    }

    /// Rewrites every formula and defined name for a structural edit (rows inserted, a sheet
    /// renamed, a block moved...).
    pub(crate) fn apply_ref_edit(&mut self, edit: &crate::formula::RefEdit) -> Result<()> {
        if self.book.adjusts_references() {
            return Ok(());
        }
        for id in self.book.sheet_ids() {
            let Some(name) = self.book.sheet_name(id) else { continue };
            let Some(used) = self.book.used_range(id) else { continue };
            for (r, c) in self.book.occupied_cells(id, used) {
                if let Some(f) = self.book.cell_formula(id, r, c) {
                    let g = crate::formula::adjust_refs(&f, &name, edit);
                    if g != f {
                        let fmt = self.book.cell_format(id, r, c);
                        self.book.set_cell_formula(id, r, c, Some(g))?;
                        self.book.set_cell_format(id, r, c, fmt)?;
                    }
                }
            }
        }
        for (n, scope, refers) in self.book.defined_names() {
            let own = scope.and_then(|s| self.book.sheet_name(s)).unwrap_or_default();
            let g = crate::formula::adjust_refs(&format!("={refers}"), &own, edit);
            let g = g.strip_prefix('=').unwrap_or(&g).to_owned();
            if g != refers {
                self.book.set_defined_name(&n, scope, &g);
            }
        }
        Ok(())
    }

    /// The number format a newly entered formula gives a General cell: a date for `DATE()` and
    /// `TODAY()`, a time for `TIME()`, or the format of the cell the formula starts with.
    fn formula_format(&self, sheet: SheetId, formula: &str) -> Option<String> {
        let fmt_of = |name: Option<&str>, r: Rect| -> Option<String> {
            let target = match name {
                Some(n) => self.sheet_by_name(n)?,
                None => sheet,
            };
            self.book.cell_format(target, r.r1, r.c1).number_format
        };
        crate::eval::auto_format(formula, &fmt_of)
    }

    /// Stores an already-coerced input in one cell.
    pub(crate) fn apply_input(&mut self, sheet: SheetId, r: u32, c: u32, input: Input) -> Result<()> {
        match input {
            Input::Clear => {
                if !self.is_blank(sheet, r, c) {
                    self.book.set_cell_value(sheet, r, c, Value::Empty)?;
                }
                Ok(())
            }
            Input::Formula(f) => {
                let auto = self.formula_format(sheet, &f);
                self.book.set_cell_formula(sheet, r, c, Some(f))?;
                if let Some(fmt) = auto {
                    let mut cf = self.book.cell_format(sheet, r, c);
                    if cf.number_format.is_none() {
                        cf.number_format = Some(fmt);
                        self.book.set_cell_format(sheet, r, c, cf)?;
                    }
                }
                Ok(())
            }
            Input::Const(v, fmt) => {
                let wraps = matches!(&v, Value::String(s) if s.contains('\n'));
                self.book.set_cell_value(sheet, r, c, v)?;
                if fmt.is_some() || wraps {
                    let mut cf = self.book.cell_format(sheet, r, c);
                    let before = cf.clone();
                    if let Some(fmt) = fmt
                        && cf.number_format.is_none()
                    {
                        cf.number_format = Some(fmt.to_owned());
                    }
                    if wraps {
                        cf.wrap = true;
                    }
                    if cf != before {
                        self.book.set_cell_format(sheet, r, c, cf)?;
                    }
                }
                Ok(())
            }
        }
    }

    /// Assigns a macro value to one cell (`Range.Value = v`).
    pub(crate) fn write_value(&mut self, sheet: SheetId, r: u32, c: u32, v: &Value) -> Result<()> {
        let text = self.book.cell_format(sheet, r, c).number_format.as_deref() == Some("@");
        let input = cell_input(v, text)?;
        self.apply_input(sheet, r, c, input)
    }

    pub(crate) fn clear_contents(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        for (r, c) in self.content_cells(sheet, area) {
            self.book.set_cell_value(sheet, r, c, Value::Empty)?;
        }
        Ok(())
    }

    pub(crate) fn clear_formats(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        for (r, c) in self.book.occupied_cells(sheet, area) {
            if !self.book.cell_format(sheet, r, c).is_default() {
                self.book.set_cell_format(sheet, r, c, CellFormat::default())?;
            }
        }
        Ok(())
    }

    pub(crate) fn snapshot(&self, sheet: SheetId, src: Rect, cut: bool) -> Clipboard {
        let cells = self
            .book
            .occupied_cells(sheet, src)
            .into_iter()
            .map(|(r, c)| ClipCell {
                dr: r - src.r1,
                dc: c - src.c1,
                value: self.book.cell_value(sheet, r, c),
                formula: self.book.cell_formula(sheet, r, c),
                format: self.book.cell_format(sheet, r, c),
                comment: self.book.cell_comment(sheet, r, c),
            })
            .collect();
        Clipboard { sheet, src, cells, cut }
    }

    /// Pastes `clip` at `dest` (top-left, or a block that is a multiple of the source size, which
    /// is tiled). Returns the pasted rectangle.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn paste(
        &mut self,
        clip: &Clipboard,
        sheet: SheetId,
        dest: Rect,
        kind: PasteKind,
        skip_blanks: bool,
        transpose: bool,
        shift_formulas: bool,
        operation: i64,
    ) -> Result<Rect> {
        self.check_writable(sheet)?;
        let (sr, sc) = if transpose { (clip.src.cols(), clip.src.rows()) } else { (clip.src.rows(), clip.src.cols()) };
        let (tiles_r, tiles_c) = if dest.rows().is_multiple_of(sr) && dest.cols().is_multiple_of(sc) && !dest.is_cell()
        {
            (dest.rows() / sr, dest.cols() / sc)
        } else {
            (1, 1)
        };
        let total_r = u64::from(sr) * u64::from(tiles_r);
        let total_c = u64::from(sc) * u64::from(tiles_c);
        if u64::from(dest.r1) + total_r - 1 > u64::from(MAX_ROWS)
            || u64::from(dest.c1) + total_c - 1 > u64::from(MAX_COLS)
        {
            return Err(err1004("PasteSpecial method of Range class failed"));
        }
        let pasted = Rect::new(dest.r1, dest.c1, dest.r1 + total_r as u32 - 1, dest.c1 + total_c as u32 - 1);
        for tr in 0..tiles_r {
            for tc in 0..tiles_c {
                let r0 = dest.r1 + tr * sr;
                let c0 = dest.c1 + tc * sc;
                let target = Rect::new(r0, c0, r0 + sr - 1, c0 + sc - 1);
                if !skip_blanks && operation == -4142 {
                    match kind {
                        PasteKind::All | PasteKind::AllExceptBorders => {
                            self.clear_contents(sheet, target)?;
                            self.clear_formats(sheet, target)?;
                        }
                        PasteKind::Values | PasteKind::Formulas | PasteKind::ValuesAndNumberFormats => {
                            self.clear_contents(sheet, target)?;
                        }
                        PasteKind::Formats => self.clear_formats(sheet, target)?,
                    }
                }
                for cell in &clip.cells {
                    let (dr, dc) = if transpose { (cell.dc, cell.dr) } else { (cell.dr, cell.dc) };
                    let (r, c) = (r0 + dr, c0 + dc);
                    let blank = cell.value.is_empty() && cell.formula.is_none();
                    if skip_blanks && blank {
                        continue;
                    }
                    if operation != -4142 && kind != PasteKind::Formats {
                        // Paste with an operation combines numbers with what is there.
                        let src = cell.value.to_f64().ok().filter(|_| cell.value.is_numeric_type() || blank);
                        let cur = self.book.cell_value(sheet, r, c);
                        let dst = if cur.is_empty() {
                            Some(0.0)
                        } else {
                            cur.to_f64().ok().filter(|_| cur.is_numeric_type())
                        };
                        if let (Some(s), Some(d)) = (src.or(Some(0.0)), dst) {
                            let v = match operation {
                                2 => Some(d + s),
                                3 => Some(d - s),
                                4 => Some(d * s),
                                5 => (s != 0.0).then(|| d / s),
                                _ => Some(s),
                            };
                            self.book.set_cell_value(sheet, r, c, v.map_or(Value::Error(2007), Value::Double))?;
                        }
                        continue;
                    }
                    let formula = cell.formula.as_ref().map(|f| {
                        if shift_formulas {
                            let src_r = i64::from(clip.src.r1 + cell.dr);
                            let src_c = i64::from(clip.src.c1 + cell.dc);
                            shift_formula(f, i64::from(r) - src_r, i64::from(c) - src_c)
                        } else {
                            f.clone()
                        }
                    });
                    let write_content = |st: &mut State, with_formula: bool| -> Result<()> {
                        match (&formula, with_formula) {
                            (Some(f), true) => st.book.set_cell_formula(sheet, r, c, Some(f.clone())),
                            _ => st.book.set_cell_value(sheet, r, c, cell.value.clone()),
                        }
                    };
                    match kind {
                        PasteKind::All => {
                            write_content(self, true)?;
                            self.book.set_cell_format(sheet, r, c, cell.format.clone())?;
                            if cell.comment.is_some() {
                                self.book.set_cell_comment(sheet, r, c, cell.comment.clone())?;
                            }
                        }
                        PasteKind::AllExceptBorders => {
                            write_content(self, true)?;
                            let mut f = cell.format.clone();
                            f.borders = self.book.cell_format(sheet, r, c).borders;
                            self.book.set_cell_format(sheet, r, c, f)?;
                        }
                        PasteKind::Values => {
                            if !blank {
                                write_content(self, false)?;
                            }
                        }
                        PasteKind::Formulas => {
                            if !blank {
                                write_content(self, true)?;
                            }
                        }
                        PasteKind::ValuesAndNumberFormats => {
                            if !blank {
                                write_content(self, false)?;
                            }
                            let mut f = self.book.cell_format(sheet, r, c);
                            f.number_format = cell.format.number_format.clone();
                            self.book.set_cell_format(sheet, r, c, f)?;
                        }
                        PasteKind::Formats => {
                            self.book.set_cell_format(sheet, r, c, cell.format.clone())?;
                        }
                    }
                }
            }
        }
        self.saved = false;
        Ok(pasted)
    }
}

// ---- argument helpers ----

pub(crate) fn opt_i64(args: &Args, idx: usize, name: &str) -> Result<Option<i64>> {
    args.get(idx, name).map(|v| v.to_i64()).transpose()
}

pub(crate) fn opt_bool(args: &Args, idx: usize, name: &str) -> Result<Option<bool>> {
    args.get(idx, name).map(|v| v.to_bool()).transpose()
}

pub(crate) fn opt_text(args: &Args, idx: usize, name: &str) -> Result<Option<String>> {
    args.get(idx, name).map(|v| v.to_text()).transpose()
}

/// A column index argument: a number or column letters (`Cells(1, "B")`).
pub(crate) fn column_arg(v: &Value) -> Result<i64> {
    match v {
        Value::String(s) => match column_number(s.trim()) {
            Some(n) => Ok(i64::from(n)),
            None => match parse_number(s) {
                Some(n) => Value::Double(n).to_i64(),
                None => Err(err1004("Application-defined or object-defined error")),
            },
        },
        other => other.to_i64(),
    }
}

/// Unwraps an object argument to its default value (a Range used as data).
pub(crate) fn data(v: &Value) -> Result<Value> {
    v.clone().deref_default()
}
