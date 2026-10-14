//! `Range.AutoFill`, `FillDown`, `FillRight`, `FillUp`, `FillLeft` and `DataSeries`: copies and
//! series the way dragging the fill handle does (numbers by their step, dates by day or month,
//! text with a trailing number counting up, day and month names, quarters, formulas with their
//! relative references moved).

use macroost_lang::{Result, Value};

use crate::address::Rect;
use crate::format::is_date_format;
use crate::formula::shift_formula;
use crate::state::{State, err1004};
use crate::workbook::{CellFormat, SheetId};

const FILL_DEFAULT: i64 = 0;
const FILL_COPY: i64 = 1;
const FILL_SERIES: i64 = 2;
const FILL_FORMATS: i64 = 3;
const FILL_VALUES: i64 = 4;
const FILL_DAYS: i64 = 5;
const FILL_WEEKDAYS: i64 = 6;
const FILL_MONTHS: i64 = 7;
const FILL_YEARS: i64 = 8;
const FILL_LINEAR: i64 = 9;
const FILL_GROWTH: i64 = 10;

const DAYS_SHORT: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const DAYS_LONG: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MONTHS_SHORT: [&str; 12] = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

#[derive(Clone)]
struct Src {
    value: Value,
    formula: Option<String>,
    format: CellFormat,
    /// Position along the fill line (0-based from the first source cell).
    pos: i64,
}

/// Matches text against the built-in lists; (list, index).
fn in_list(s: &str) -> Option<(&'static [&'static str], usize)> {
    let lists: [&'static [&'static str]; 4] = [&DAYS_SHORT, &DAYS_LONG, &MONTHS_SHORT, &MONTHS_LONG];
    for l in lists {
        if let Some(i) = l.iter().position(|x| x.eq_ignore_ascii_case(s)) {
            return Some((l, i));
        }
    }
    None
}

/// Case pattern of `like` applied to `word`.
fn same_case(word: &str, like: &str) -> String {
    if like.chars().all(|c| !c.is_lowercase()) {
        word.to_uppercase()
    } else if like.chars().all(|c| !c.is_uppercase()) {
        word.to_lowercase()
    } else {
        word.to_owned()
    }
}

/// Splits text into (prefix, trailing number) — `Item 1` -> ("Item ", 1).
fn trailing_number(s: &str) -> Option<(String, i64, usize)> {
    let digits: String =
        s.chars().rev().take_while(char::is_ascii_digit).collect::<Vec<_>>().into_iter().rev().collect();
    if digits.is_empty() || digits.len() == s.len() {
        return None;
    }
    let n: i64 = digits.parse().ok()?;
    Some((s[..s.len() - digits.len()].to_owned(), n, digits.len()))
}

fn add_months(serial: f64, months: i64) -> f64 {
    let (y, m, d) = crate::eval::serial_ymd(serial);
    let total = i64::from(m) - 1 + months;
    let (ny, nm) = (y + total.div_euclid(12), total.rem_euclid(12) + 1);
    let last = crate::eval::ymd_serial(ny, nm + 1, 1) - 1.0;
    let (_, _, last_day) = crate::eval::serial_ymd(last);
    crate::eval::ymd_serial(ny, nm, i64::from(d.min(last_day))) + serial.fract()
}

fn is_weekend(serial: f64) -> bool {
    let w = (serial.floor() as i64 - 1).rem_euclid(7); // 0 = Sunday
    w == 0 || w == 6
}

/// The value at position `k` of a series continuing `src` (every source cell numeric).
fn numeric_series(src: &[Src], k: i64, ty: i64, date: bool) -> f64 {
    let xs: Vec<f64> = src.iter().map(|s| s.value.to_f64().unwrap_or(0.0)).collect();
    let first = xs[0];
    if ty == FILL_MONTHS || ty == FILL_YEARS {
        let unit = if ty == FILL_YEARS { 12 } else { 1 };
        let step = if xs.len() >= 2 {
            let (y0, m0, _) = crate::eval::serial_ymd(xs[0]);
            let (y1, m1, _) = crate::eval::serial_ymd(xs[1]);
            ((y1 - y0) * 12 + i64::from(m1) - i64::from(m0)).max(1)
        } else {
            unit
        };
        return add_months(first, step * k);
    }
    if ty == FILL_WEEKDAYS {
        let mut d = first;
        let mut n = 0;
        while n < k {
            d += 1.0;
            if !is_weekend(d) {
                n += 1;
            }
        }
        while n > k {
            d -= 1.0;
            if !is_weekend(d) {
                n -= 1;
            }
        }
        return d;
    }
    if xs.len() == 1 {
        let step = if ty == FILL_SERIES || ty == FILL_DAYS || ty == FILL_LINEAR || (date && ty == FILL_DEFAULT) {
            1.0
        } else {
            0.0
        };
        return first + step * k as f64;
    }
    if ty == FILL_GROWTH {
        let ratio = if xs[0] != 0.0 { xs[1] / xs[0] } else { 1.0 };
        return first * ratio.powi(k as i32);
    }
    // Least-squares line through the source values.
    let n = xs.len() as f64;
    let mean_x = (n - 1.0) / 2.0;
    let mean_y = xs.iter().sum::<f64>() / n;
    let mut num = 0.0;
    let mut den = 0.0;
    for (i, y) in xs.iter().enumerate() {
        let dx = i as f64 - mean_x;
        num += dx * (y - mean_y);
        den += dx * dx;
    }
    let slope = if den == 0.0 { 0.0 } else { num / den };
    let v = mean_y + slope * (k as f64 - mean_x);
    let r: f64 = format!("{v:.12e}").parse().unwrap_or(v);
    if r.abs() < 1e-12 { 0.0 } else { r }
}

/// Value (or formula) for position `k` of a fill line.
fn generate(src: &[Src], k: i64, ty: i64, vertical: bool) -> (Value, Option<String>, CellFormat) {
    let n = src.len() as i64;
    let template = &src[k.rem_euclid(n) as usize];
    let copy_only = ty == FILL_COPY || ty == FILL_FORMATS;
    if let Some(f) = &template.formula {
        let shift = k - template.pos;
        let g = if vertical { shift_formula(f, shift, 0) } else { shift_formula(f, 0, shift) };
        return (Value::Empty, Some(g), template.format.clone());
    }
    let numeric = src.iter().all(|s| s.formula.is_none() && s.value.is_numeric_type());
    if numeric && !copy_only {
        let date = is_date_format(src[0].format.number_format.as_deref().unwrap_or("General"));
        let single_copy = n == 1 && (ty == FILL_DEFAULT && !date || ty == FILL_VALUES);
        if !single_copy {
            return (Value::Double(numeric_series(src, k, ty, date)), None, template.format.clone());
        }
    }
    if !copy_only && let Value::String(s) = &src[0].value {
        // Built-in lists continue from the first cell by the step of the first two.
        if let Some((list, i0)) = in_list(s) {
            let step = match src.get(1).map(|x| &x.value) {
                Some(Value::String(b)) => in_list(b).map_or(1, |(_, i1)| i1 as i64 - i0 as i64),
                _ => 1,
            };
            let j = (i0 as i64 + k * step).rem_euclid(list.len() as i64) as usize;
            return (Value::from(same_case(list[j], s)), None, template.format.clone());
        }
        if let Some((prefix, num, width)) = trailing_number(s) {
            let step = match src.get(1).map(|x| &x.value) {
                Some(Value::String(b)) => {
                    trailing_number(b).filter(|(p, _, _)| *p == prefix).map_or(1, |(_, m, _)| m - num)
                }
                _ => 1,
            };
            let quarter = prefix.eq_ignore_ascii_case("q") && (1..=4).contains(&num);
            let next = num + k * step;
            let next = if quarter { (next - 1).rem_euclid(4) + 1 } else { next };
            let text = if next < 0 { format!("{prefix}{next}") } else { format!("{prefix}{next:0width$}") };
            return (Value::from(text), None, template.format.clone());
        }
    }
    (template.value.clone(), None, template.format.clone())
}

/// AutoFill from `src` over `dest` (which must extend `src` in one direction).
pub(crate) fn autofill(st: &mut State, sheet: SheetId, src: Rect, dest: Rect, ty: i64) -> Result<()> {
    let fail = || err1004("AutoFill method of Range class failed");
    if !dest.contains_rect(&src) || dest == src {
        return Err(fail());
    }
    let vertical = src.c1 == dest.c1 && src.c2 == dest.c2;
    let horizontal = src.r1 == dest.r1 && src.r2 == dest.r2;
    if !vertical && !horizontal {
        return Err(fail());
    }
    st.check_writable(sheet)?;
    let forward = if vertical { src.r1 == dest.r1 } else { src.c1 == dest.c1 };
    if !forward && (if vertical { src.r2 != dest.r2 } else { src.c2 != dest.c2 }) {
        return Err(fail());
    }
    let (lines, len) = if vertical { (src.c1..=src.c2, src.rows()) } else { (src.r1..=src.r2, src.cols()) };
    for line in lines {
        let at = |k: u32| if vertical { (src.r1 + k, line) } else { (line, src.c1 + k) };
        let cells: Vec<Src> = (0..len)
            .map(|k| {
                let (r, c) = at(k);
                Src {
                    value: st.book.cell_value(sheet, r, c),
                    formula: st.book.cell_formula(sheet, r, c),
                    format: st.book.cell_format(sheet, r, c),
                    pos: i64::from(k),
                }
            })
            .collect();
        let targets: Vec<(u32, u32, i64)> = if vertical {
            if forward {
                (src.r2 + 1..=dest.r2).map(|r| (r, line, i64::from(r - src.r1))).collect()
            } else {
                (dest.r1..src.r1).map(|r| (r, line, i64::from(r) - i64::from(src.r1))).collect()
            }
        } else if forward {
            (src.c2 + 1..=dest.c2).map(|c| (line, c, i64::from(c - src.c1))).collect()
        } else {
            (dest.c1..src.c1).map(|c| (line, c, i64::from(c) - i64::from(src.c1))).collect()
        };
        for (r, c, k) in targets {
            let (v, f, fmt) = generate(&cells, k, ty, vertical);
            if ty != FILL_FORMATS {
                match f {
                    Some(f) => st.book.set_cell_formula(sheet, r, c, Some(f))?,
                    None => st.book.set_cell_value(sheet, r, c, v)?,
                }
            }
            if ty != FILL_VALUES {
                st.book.set_cell_format(sheet, r, c, fmt)?;
            }
        }
    }
    st.saved = false;
    Ok(())
}

/// `FillDown` / `FillUp` / `FillRight` / `FillLeft`: the edge row or column copied over the
/// rest of the range.
pub(crate) fn fill_copy(st: &mut State, sheet: SheetId, a: Rect, dir: char) -> Result<()> {
    st.check_writable(sheet)?;
    for k in 0..if matches!(dir, 'd' | 'u') { a.cols() } else { a.rows() } {
        let (sr, sc) = match dir {
            'd' => (a.r1, a.c1 + k),
            'u' => (a.r2, a.c1 + k),
            'r' => (a.r1 + k, a.c1),
            _ => (a.r1 + k, a.c2),
        };
        let v = st.book.cell_value(sheet, sr, sc);
        let f = st.book.cell_formula(sheet, sr, sc);
        let fmt = st.book.cell_format(sheet, sr, sc);
        let targets: Vec<(u32, u32)> = match dir {
            'd' => (a.r1 + 1..=a.r2).map(|r| (r, sc)).collect(),
            'u' => (a.r1..a.r2).map(|r| (r, sc)).collect(),
            'r' => (a.c1 + 1..=a.c2).map(|c| (sr, c)).collect(),
            _ => (a.c1..a.c2).map(|c| (sr, c)).collect(),
        };
        for (r, c) in targets {
            match &f {
                Some(f) => st.book.set_cell_formula(
                    sheet,
                    r,
                    c,
                    Some(shift_formula(f, i64::from(r) - i64::from(sr), i64::from(c) - i64::from(sc))),
                )?,
                None => st.book.set_cell_value(sheet, r, c, v.clone())?,
            }
            st.book.set_cell_format(sheet, r, c, fmt.clone())?;
        }
    }
    st.saved = false;
    Ok(())
}

/// `DataSeries Rowcol, Type, Date, Step, Stop`: from the first cell of each line.
pub(crate) fn data_series(
    st: &mut State,
    sheet: SheetId,
    a: Rect,
    by_columns: bool,
    growth: bool,
    step: f64,
) -> Result<()> {
    st.check_writable(sheet)?;
    let lines = if by_columns { a.c1..=a.c2 } else { a.r1..=a.r2 };
    for line in lines {
        let (r0, c0) = if by_columns { (a.r1, line) } else { (line, a.c1) };
        let Ok(start) = st.book.cell_value(sheet, r0, c0).to_f64() else { continue };
        let fmt = st.book.cell_format(sheet, r0, c0);
        let n = if by_columns { a.rows() } else { a.cols() };
        let mut v = start;
        for k in 1..n {
            v = if growth { v * step } else { v + step };
            let (r, c) = if by_columns { (a.r1 + k, line) } else { (line, a.c1 + k) };
            st.book.set_cell_value(sheet, r, c, Value::Double(v))?;
            st.book.set_cell_format(sheet, r, c, fmt.clone())?;
        }
    }
    st.saved = false;
    Ok(())
}
