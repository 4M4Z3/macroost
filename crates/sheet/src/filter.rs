//! AutoFilter: `Range.AutoFilter Field, Criteria1, Operator, Criteria2`, the `AutoFilter`
//! object, `AutoFilterMode` / `FilterMode` / `ShowAllData`. Filtering hides the data rows that
//! fail a condition (so `SpecialCells(xlCellTypeVisible)` and `SUBTOTAL(109, ...)` see the rest).

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::address::Rect;
use crate::range::{Kind, range_value, wildcard_match};
use crate::state::{AutoFilter, FilterCond, Shared, State, data, err1004, opt_i64, with, with_mut};
use crate::workbook::SheetId;

const XL_AND: i64 = 1;
const XL_OR: i64 = 2;
const XL_FILTER_VALUES: i64 = 7;

/// Does one criterion (`"East"`, `">15"`, `"<>Bob"`, `"B*"`, `"="`) accept a cell?
fn matches(st: &State, sheet: SheetId, r: u32, c: u32, crit: &str) -> bool {
    let v = st.book.cell_value(sheet, r, c);
    let text = st.text(sheet, r, c);
    let (op, rest) = ["<>", ">=", "<=", "=", ">", "<"]
        .iter()
        .find_map(|op| crit.strip_prefix(op).map(|r| (*op, r)))
        .unwrap_or(("", crit));
    if op.is_empty() || op == "=" || op == "<>" {
        let eq = if rest.is_empty() {
            v.is_empty()
        } else if op.is_empty() {
            // A bare criterion matches text cells (wildcards allowed).
            matches!(v, Value::String(_)) && wildcard_match(rest, &text, true, false)
        } else {
            match (macroost_lang::value::parse_number(rest), &v) {
                (Some(n), x) if x.is_numeric_type() => x.to_f64().is_ok_and(|y| y == n),
                _ => wildcard_match(rest, &text, true, false),
            }
        };
        return if op == "<>" { !eq } else { eq };
    }
    let Some(n) = macroost_lang::value::parse_number(rest) else {
        let o = text.to_lowercase().cmp(&rest.to_lowercase());
        return match op {
            ">" => o.is_gt(),
            "<" => o.is_lt(),
            ">=" => !o.is_lt(),
            _ => !o.is_gt(),
        };
    };
    if !v.is_numeric_type() {
        return false;
    }
    let x = v.to_f64().unwrap_or(0.0);
    match op {
        ">" => x > n,
        "<" => x < n,
        ">=" => x >= n,
        _ => x <= n,
    }
}

fn crit_texts(v: &Value) -> Vec<String> {
    match v {
        Value::Array(a) => a.borrow().data.iter().filter_map(|x| x.to_text().ok()).collect(),
        other => other.to_text().ok().into_iter().collect(),
    }
}

fn row_visible(st: &State, sheet: SheetId, f: &AutoFilter, r: u32) -> bool {
    f.conds.iter().all(|cond| {
        let c = f.range.c1 + cond.field - 1;
        let one = |v: &Option<Value>| -> bool {
            match v {
                None => true,
                Some(v) => crit_texts(v).iter().any(|t| {
                    if cond.operator == XL_FILTER_VALUES {
                        st.text(sheet, r, c).eq_ignore_ascii_case(t)
                    } else {
                        matches(st, sheet, r, c, t)
                    }
                }),
            }
        };
        match (cond.operator, &cond.crit2) {
            (XL_OR, Some(_)) => one(&cond.crit1) || one(&cond.crit2),
            (XL_AND, Some(_)) => one(&cond.crit1) && one(&cond.crit2),
            _ => one(&cond.crit1),
        }
    })
}

/// Shows or hides every data row of the sheet's filter.
fn apply(st: &mut State, sheet: SheetId) -> Result<()> {
    let Some(f) = st.filters.get(&sheet).cloned() else { return Ok(()) };
    for r in f.range.r1 + 1..=f.range.r2 {
        let vis = row_visible(st, sheet, &f, r);
        st.book.set_row_hidden(sheet, r, !vis)?;
        st.book.set_row_filtered(sheet, r, !vis);
    }
    Ok(())
}

/// Turns the sheet's AutoFilter off, showing every row it hid.
pub(crate) fn remove(app: &Shared, sheet: SheetId) -> Result<()> {
    with_mut(app, |st| {
        if let Some(f) = st.filters.remove(&sheet) {
            for r in f.range.r1 + 1..=f.range.r2 {
                st.book.set_row_hidden(sheet, r, false)?;
                st.book.set_row_filtered(sheet, r, false);
            }
        }
        Ok(())
    })
}

pub(crate) fn show_all(app: &Shared, sheet: SheetId) -> Result<()> {
    with_mut(app, |st| {
        let Some(f) = st.filters.get_mut(&sheet) else {
            return Err(err1004("ShowAllData method of Worksheet class failed"));
        };
        f.conds.clear();
        let range = f.range;
        for r in range.r1 + 1..=range.r2 {
            st.book.set_row_hidden(sheet, r, false)?;
            st.book.set_row_filtered(sheet, r, false);
        }
        Ok(())
    })
}

/// `Range.AutoFilter`: with no arguments toggles the filter; with a field sets its condition.
pub(crate) fn range_autofilter(app: &Shared, sheet: SheetId, area: Rect, args: &Args) -> Result<Value> {
    let field = opt_i64(args, 0, "Field")?;
    let crit1 = args.get(1, "Criteria1").map(data).transpose()?;
    let operator = opt_i64(args, 2, "Operator")?.unwrap_or(XL_AND);
    let crit2 = args.get(3, "Criteria2").map(data).transpose()?;
    let exists = with(app, |st| Ok(st.filters.contains_key(&sheet)))?;
    if field.is_none() && exists {
        remove(app, sheet)?;
        return Ok(Value::Boolean(true));
    }
    with_mut(app, |st| {
        let range = match st.filters.get(&sheet) {
            Some(f) => f.range,
            None => {
                let r = if area.is_cell() { crate::range::current_region_of(st, sheet, area) } else { area };
                let r = st.clip(sheet, r).unwrap_or(r);
                st.filters.insert(sheet, AutoFilter { range: r, conds: Vec::new() });
                r
            }
        };
        if let Some(field) = field {
            if field < 1 || field as u32 > range.cols() {
                return Err(err1004("AutoFilter method of Range class failed"));
            }
            let f = st.filters.get_mut(&sheet).ok_or_else(RuntimeError::object_required)?;
            f.conds.retain(|c| c.field != field as u32);
            if crit1.is_some() || crit2.is_some() {
                f.conds.push(FilterCond { field: field as u32, crit1, crit2, operator });
            }
            apply(st, sheet)?;
        }
        Ok(Value::Boolean(true))
    })
}

/// `Worksheet.AutoFilter`: the filter object, or `Nothing` without a filter.
pub(crate) fn autofilter_object(app: &Shared, sheet: SheetId) -> Result<Value> {
    let exists = with(app, |st| Ok(st.filters.contains_key(&sheet)))?;
    Ok(if exists {
        Value::Object(Some(Rc::new(AutoFilterObj { app: app.clone(), sheet })))
    } else {
        Value::Object(None)
    })
}

struct AutoFilterObj {
    app: Shared,
    sheet: SheetId,
}

impl Object for AutoFilterObj {
    fn type_name(&self) -> String {
        "AutoFilter".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let f = with(&self.app, |st| st.filters.get(&self.sheet).cloned().ok_or_else(RuntimeError::object_required))?;
        match name.to_ascii_lowercase().as_str() {
            "range" => Ok(range_value(&self.app, self.sheet, vec![f.range], Kind::Cells)),
            "filtermode" => Ok(Value::Boolean(!f.conds.is_empty())),
            "showalldata" => show_all(&self.app, self.sheet).map(|_| Value::Empty),
            "applyfilter" => with_mut(&self.app, |st| apply(st, self.sheet)).map(|_| Value::Empty),
            "parent" => Ok(crate::objects::sheet_value(&self.app, self.sheet)),
            "application" => Ok(crate::objects::app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}
