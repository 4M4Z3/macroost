//! Sorting: `Range.Sort` (whose Order, Orientation, Header and MatchCase settings persist
//! between calls, as documented) and the `Worksheet.Sort` object (`SortFields`, `SetRange`,
//! `Header`, `Apply`).

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::address::Rect;
use crate::formula::shift_formula;
use crate::range::{decode_range, range_arg};
use crate::state::{Shared, SheetSort, State, data, err1004, opt_bool, opt_i64, with, with_mut};
use crate::workbook::{CellFormat, SheetId};

/// Sort order of two cell values: numbers, text, logicals, errors; blanks always last.
fn compare(a: &Value, b: &Value, desc: bool, match_case: bool) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let rank = |v: &Value| -> u8 {
        match v {
            Value::Empty => 4,
            Value::String(s) if s.is_empty() => 4,
            Value::String(_) => 1,
            Value::Boolean(_) => 2,
            Value::Error(_) => 3,
            _ => 0,
        }
    };
    let (ra, rb) = (rank(a), rank(b));
    if ra == 4 || rb == 4 {
        return ra.cmp(&rb);
    }
    let o = ra.cmp(&rb).then_with(|| match (a, b) {
        (Value::String(x), Value::String(y)) => {
            let l = x.to_lowercase().cmp(&y.to_lowercase());
            if l == Ordering::Equal && match_case {
                // Lower case sorts before upper case.
                y.cmp(x)
            } else {
                l
            }
        }
        (Value::Boolean(x), Value::Boolean(y)) => x.cmp(y),
        _ => {
            let x = a.to_f64().unwrap_or(0.0);
            let y = b.to_f64().unwrap_or(0.0);
            x.partial_cmp(&y).unwrap_or(Ordering::Equal)
        }
    });
    if desc { o.reverse() } else { o }
}

/// Sorts `area` on `sheet`. `keys` are absolute column numbers (or row numbers when
/// `horizontal`) with a descending flag.
pub(crate) fn sort_area(
    st: &mut State,
    sheet: SheetId,
    area: Rect,
    keys: &[(u32, bool)],
    header: bool,
    horizontal: bool,
    match_case: bool,
) -> Result<()> {
    st.check_writable(sheet)?;
    let Some(mut a) = st.clip(sheet, area) else { return Ok(()) };
    if st.book.merged_areas(sheet).iter().any(|m| m.intersect(&a).is_some()) {
        return Err(err1004("To do this, all the merged cells need to be the same size."));
    }
    if header {
        if horizontal {
            if a.c1 == a.c2 {
                return Ok(());
            }
            a.c1 += 1;
        } else {
            if a.r1 == a.r2 {
                return Ok(());
            }
            a.r1 += 1;
        }
    }
    type Cell = (Value, Option<String>, CellFormat);
    // Lines are rows (vertical sort) or columns (horizontal sort).
    let (lines, width) = if horizontal { (a.c1..=a.c2, a.rows()) } else { (a.r1..=a.r2, a.cols()) };
    let mut items: Vec<(u32, Vec<Cell>)> = Vec::new();
    for line in lines {
        let cells: Vec<Cell> = (0..width)
            .map(|k| {
                let (r, c) = if horizontal { (a.r1 + k, line) } else { (line, a.c1 + k) };
                (st.book.cell_value(sheet, r, c), st.book.cell_formula(sheet, r, c), st.book.cell_format(sheet, r, c))
            })
            .collect();
        items.push((line, cells));
    }
    let base = if horizontal { a.r1 } else { a.c1 };
    items.sort_by(|x, y| {
        for &(key, desc) in keys {
            let i = key.saturating_sub(base) as usize;
            let vx = x.1.get(i).map(|t| t.0.clone()).unwrap_or_default();
            let vy = y.1.get(i).map(|t| t.0.clone()).unwrap_or_default();
            let o = compare(&vx, &vy, desc, match_case);
            if o != std::cmp::Ordering::Equal {
                return o;
            }
        }
        x.0.cmp(&y.0)
    });
    let start = if horizontal { a.c1 } else { a.r1 };
    for (i, (src, cells)) in items.into_iter().enumerate() {
        let line = start + i as u32;
        for (k, (v, f, fmt)) in cells.into_iter().enumerate() {
            let (r, c) = if horizontal { (a.r1 + k as u32, line) } else { (line, a.c1 + k as u32) };
            let moved = i64::from(line) - i64::from(src);
            match f {
                Some(f) => {
                    let g = if horizontal { shift_formula(&f, 0, moved) } else { shift_formula(&f, moved, 0) };
                    st.book.set_cell_formula(sheet, r, c, Some(g))?
                }
                None => st.book.set_cell_value(sheet, r, c, v)?,
            }
            st.book.set_cell_format(sheet, r, c, fmt)?;
        }
    }
    st.saved = false;
    Ok(())
}

fn guess_header(st: &State, sheet: SheetId, a: Rect, horizontal: bool) -> bool {
    if horizontal {
        (a.r1..=a.r2).all(|r| matches!(st.book.cell_value(sheet, r, a.c1), Value::String(_)))
    } else {
        (a.c1..=a.c2).all(|c| matches!(st.book.cell_value(sheet, a.r1, c), Value::String(_)))
            && (a.c1..=a.c2).any(|c| !matches!(st.book.cell_value(sheet, a.r1 + 1, c), Value::String(_)))
    }
}

/// `Range.Sort Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase,
/// Orientation, ...`.
pub(crate) fn range_sort(app: &Shared, sheet: SheetId, area: Rect, args: &Args) -> Result<Value> {
    let mut key_args = Vec::new();
    for (i, (k, ki, oi)) in [("Key1", 0, 1), ("Key2", 2, 4), ("Key3", 5, 6)].iter().enumerate() {
        let order = opt_i64(args, *oi, &format!("Order{}", i + 1))?;
        let key = match args.get(*ki, k) {
            Some(kv) => Some(range_arg(app, sheet, kv)?.1[0]),
            None => None,
        };
        key_args.push((key, order));
    }
    let header = opt_i64(args, 7, "Header")?;
    let match_case = opt_bool(args, 9, "MatchCase")?;
    let orientation = opt_i64(args, 10, "Orientation")?;
    with_mut(app, |st| {
        let d = st.sort_defaults.entry(sheet).or_default();
        for (i, (_, o)) in key_args.iter().enumerate() {
            if let Some(o) = o {
                d.order[i] = *o;
            }
        }
        if let Some(h) = header {
            d.header = h;
        }
        if let Some(m) = match_case {
            d.match_case = m;
        }
        if let Some(o) = orientation {
            d.orientation = o;
        }
        let d = d.clone();
        let horizontal = d.orientation == 2;
        let mut keys = Vec::new();
        for (i, (k, _)) in key_args.iter().enumerate() {
            if let Some(k) = k {
                keys.push((if horizontal { k.r1 } else { k.c1 }, d.order[i] == 2));
            }
        }
        let Some(a) = st.clip(sheet, area) else { return Ok(Value::Boolean(true)) };
        if keys.is_empty() {
            keys.push((if horizontal { a.r1 } else { a.c1 }, d.order[0] == 2));
        }
        let has_header = match d.header {
            1 => true,
            0 => guess_header(st, sheet, a, horizontal),
            _ => false,
        };
        sort_area(st, sheet, area, &keys, has_header, horizontal, d.match_case)?;
        Ok(Value::Boolean(true))
    })
}

/// `Worksheet.Sort`.
pub(crate) fn sort_object(app: &Shared, sheet: SheetId) -> Value {
    Value::Object(Some(Rc::new(SortObj { app: app.clone(), sheet })))
}

struct SortObj {
    app: Shared,
    sheet: SheetId,
}

impl SortObj {
    fn get<T>(&self, f: impl FnOnce(&SheetSort) -> T) -> Result<T> {
        with(&self.app, |st| Ok(f(&st.sheet_sorts.get(&self.sheet).cloned().unwrap_or_default())))
    }
    fn set(&self, f: impl FnOnce(&mut SheetSort)) -> Result<()> {
        with_mut(&self.app, |st| {
            f(st.sheet_sorts.entry(self.sheet).or_default());
            Ok(())
        })
    }
}

impl Object for SortObj {
    fn type_name(&self) -> String {
        "Sort".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "sortfields" => {
                Ok(Value::Object(Some(Rc::new(SortFieldsObj { app: self.app.clone(), sheet: self.sheet }))))
            }
            "setrange" => {
                let r = args.require(0, "Rng")?;
                let (_, areas) = decode_range(r).ok_or_else(RuntimeError::type_mismatch)?;
                self.set(|s| s.range = Some(areas[0]))?;
                Ok(Value::Empty)
            }
            "header" => Ok(Value::Long(self.get(|s| s.header)? as i32)),
            "matchcase" => Ok(Value::Boolean(self.get(|s| s.match_case)?)),
            "orientation" => Ok(Value::Long(self.get(|s| s.orientation)? as i32)),
            "apply" => {
                let (fields, range, header, match_case, orientation) =
                    self.get(|s| (s.fields.clone(), s.range, s.header, s.match_case, s.orientation))?;
                let range = range.ok_or_else(|| err1004("The sort reference is not valid."))?;
                let horizontal = orientation == 2;
                with_mut(&self.app, |st| {
                    let keys: Vec<(u32, bool)> =
                        fields.iter().map(|(r, d)| (if horizontal { r.r1 } else { r.c1 }, *d)).collect();
                    let has_header = match header {
                        1 => true,
                        0 => guess_header(st, self.sheet, range, horizontal),
                        _ => false,
                    };
                    sort_area(st, self.sheet, range, &keys, has_header, horizontal, match_case)
                })?;
                Ok(Value::Empty)
            }
            "rng" => match self.get(|s| s.range)? {
                Some(r) => Ok(crate::range::range_value(&self.app, self.sheet, vec![r], crate::range::Kind::Cells)),
                None => Ok(Value::Object(None)),
            },
            "parent" => Ok(crate::objects::sheet_value(&self.app, self.sheet)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let v = data(&value)?;
        match name.to_ascii_lowercase().as_str() {
            "header" => {
                let h = v.to_i64()?;
                self.set(|s| s.header = h)
            }
            "matchcase" => {
                let m = v.to_bool()?;
                self.set(|s| s.match_case = m)
            }
            "orientation" => {
                let o = v.to_i64()?;
                self.set(|s| s.orientation = o)
            }
            "sortmethod" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

struct SortFieldsObj {
    app: Shared,
    sheet: SheetId,
}

impl Object for SortFieldsObj {
    fn type_name(&self) -> String {
        "SortFields".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "clear" => with_mut(&self.app, |st| {
                st.sheet_sorts.entry(self.sheet).or_default().fields.clear();
                Ok(Value::Empty)
            }),
            "add" | "add2" => {
                let k = args.require(0, "Key")?;
                let (_, areas) = decode_range(k).ok_or_else(RuntimeError::type_mismatch)?;
                let order = opt_i64(args, 2, "Order")?.unwrap_or(1);
                with_mut(&self.app, |st| {
                    st.sheet_sorts.entry(self.sheet).or_default().fields.push((areas[0], order == 2));
                    Ok(Value::Empty)
                })
            }
            "count" => with(&self.app, |st| {
                Ok(Value::Long(st.sheet_sorts.get(&self.sheet).map_or(0, |s| s.fields.len()) as i32))
            }),
            "application" => Ok(crate::objects::app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}
