//! `WorksheetFunction`, `Application.<Function>` and `Evaluate`: macro values are converted to
//! formula-engine values, the engine's function library runs, and the result comes back as a
//! macro value. Through `WorksheetFunction` an error result raises 1004 ("Unable to get the Sum
//! property of the WorksheetFunction class"); through `Application` it is returned as an error
//! value.

use std::cell::RefCell;
use std::rc::Rc;

use macroost_lang::{Args, Array, Object, Result, RuntimeError, Value};

use crate::eval::{self, Area, Ctx, Grid, V};
use crate::range::{Kind, decode_range, range_value};
use crate::state::{Shared, with};

pub(crate) struct WorksheetFunctionObj {
    pub app: Shared,
}

/// Engine functions that are not members of the `WorksheetFunction` object (the language has its
/// own version of most of them); calling one through the object is error 438.
const NOT_MEMBERS: &[&str] = &[
    "ABS",
    "CONCATENATE",
    "SQRT",
    "MOD",
    "INT",
    "EXP",
    "VALUE",
    "LEN",
    "LEFT",
    "RIGHT",
    "MID",
    "IF",
    "UPPER",
    "LOWER",
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "HOUR",
    "MINUTE",
    "SECOND",
    "TIME",
    "SIGN",
    "TODAY",
    "NOW",
    "RAND",
    "TRUE",
    "FALSE",
    "NOT",
    "ISBLANK",
    "INDIRECT",
    "OFFSET",
    "ROW",
    "COLUMN",
    "ROWS",
    "COLUMNS",
    "ADDRESS",
    "N",
    "T",
    "NA",
    "CHAR",
    "CODE",
    "TRUNC",
    "EXACT",
    "DATEDIF",
];

/// Display spelling of a member for error messages (`vlookup` -> `VLookup`).
fn canonical(name: &str) -> String {
    const SPECIAL: &[&str] = &[
        "VLookup",
        "HLookup",
        "SumIf",
        "SumIfs",
        "CountIf",
        "CountIfs",
        "CountA",
        "CountBlank",
        "AverageIf",
        "AverageIfs",
        "IfError",
        "IfNa",
        "SumProduct",
        "RoundUp",
        "RoundDown",
        "EoMonth",
        "NetworkDays",
        "WeekNum",
        "IsNumber",
        "IsText",
        "IsError",
        "IsNA",
        "IsLogical",
        "IsErr",
        "TextJoin",
    ];
    if let Some(s) = SPECIAL.iter().find(|s| s.eq_ignore_ascii_case(name)) {
        return (*s).to_owned();
    }
    let mut c = name.chars();
    match c.next() {
        Some(f) => f.to_uppercase().chain(c.flat_map(char::to_lowercase)).collect(),
        None => String::new(),
    }
}

fn fail(name: &str) -> RuntimeError {
    RuntimeError::new(1004, format!("Unable to get the {} property of the WorksheetFunction class", canonical(name)))
}

/// A macro argument as an engine value.
pub(crate) fn to_engine(v: &Value) -> Result<V> {
    if let Some((sheet, areas)) = decode_range(v) {
        return Ok(V::Ref(Area { sheet, rects: areas }));
    }
    Ok(match v {
        Value::Error(448) => V::Missing,
        Value::Array(a) => {
            let a = a.borrow();
            let data: Vec<V> = a.data.iter().map(|x| to_engine(x).unwrap_or(V::Err(eval::VALUE))).collect();
            match a.bounds.as_slice() {
                [(lo, hi)] => {
                    // An array of equal-length arrays is a table, one row per inner array.
                    let rows: Vec<Vec<V>> = a
                        .data
                        .iter()
                        .filter_map(|x| match x {
                            Value::Array(inner) if inner.borrow().bounds.len() == 1 => Some(
                                inner
                                    .borrow()
                                    .data
                                    .iter()
                                    .map(|y| to_engine(y).unwrap_or(V::Err(eval::VALUE)))
                                    .collect(),
                            ),
                            _ => None,
                        })
                        .collect();
                    if !rows.is_empty() && rows.len() == a.data.len() && rows.iter().all(|r| r.len() == rows[0].len()) {
                        let cols = rows[0].len();
                        return Ok(V::Arr(Grid { rows: rows.len(), cols, data: rows.into_iter().flatten().collect() }));
                    }
                    V::Arr(Grid { rows: 1, cols: (hi - lo + 1).max(0) as usize, data })
                }
                [(l1, h1), (l2, h2)] => {
                    let (rows, cols) = ((h1 - l1 + 1).max(0) as usize, (h2 - l2 + 1).max(0) as usize);
                    V::Arr(Grid { rows, cols, data })
                }
                _ => return Err(RuntimeError::type_mismatch()),
            }
        }
        Value::Object(Some(o)) => to_engine(&o.default_value()?)?,
        Value::Object(None) => return Err(RuntimeError::object_not_set()),
        other => eval::from_value(other),
    })
}

/// An engine result as a macro value; `Err` for an error result.
fn from_engine(app: &Shared, v: V) -> std::result::Result<Value, i32> {
    Ok(match v {
        V::Err(e) => return Err(e),
        V::Ref(a) => range_value(app, a.sheet, a.rects, Kind::Cells),
        V::Arr(g) => grid_to_array(&g),
        other => eval::to_value(&other),
    })
}

/// A grid as a macro array: one row becomes a 1-D array, otherwise 2-D; both 1-based.
pub(crate) fn grid_to_array(g: &Grid) -> Value {
    let vals: Vec<Value> = g.data.iter().map(eval::to_value).collect();
    let arr = if g.rows == 1 {
        Array::from_vec(1, vals)
    } else {
        match Array::new(vec![(1, g.rows as i64), (1, g.cols as i64)]) {
            Ok(mut a) => {
                for r in 0..g.rows {
                    for c in 0..g.cols {
                        let _ = a.set(&[r as i64 + 1, c as i64 + 1], vals[r * g.cols + c].clone());
                    }
                }
                a
            }
            Err(_) => Array::from_vec(1, vals),
        }
    };
    Value::Array(Rc::new(RefCell::new(arr)))
}

fn engine_args(args: &Args) -> Result<Vec<V>> {
    let mut out = Vec::new();
    for v in &args.positional {
        out.push(to_engine(v)?);
    }
    for (_, v) in &args.named {
        out.push(to_engine(v)?);
    }
    // Trailing skipped arguments do not count.
    while out.last() == Some(&V::Missing) {
        out.pop();
    }
    Ok(out)
}

fn now() -> f64 {
    let secs =
        std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs_f64()).unwrap_or(0.0);
    macroost_lang::value::UNIX_EPOCH_SERIAL + secs / 86_400.0
}

/// Runs engine function `name`: `None` when the engine does not have it.
fn run(app: &Shared, name: &str, args: &Args) -> Option<Result<std::result::Result<Value, i32>>> {
    let vals = match engine_args(args) {
        Ok(v) => v,
        Err(e) => return Some(Err(e)),
    };
    let upper = name.to_ascii_uppercase().replace('_', ".");
    let r = with(app, |st| {
        let ctx = Ctx { book: st.book.as_ref(), sheet: st.active_sheet, cell: None, now: now() };
        Ok(eval::call_checked(&ctx, &upper, vals))
    });
    match r {
        Ok(Some(v)) => Some(Ok(from_engine(app, v))),
        Ok(None) => None,
        Err(e) => Some(Err(e)),
    }
}

/// `WorksheetFunction.<name>`: `None` if it is not a member.
pub(crate) fn worksheet_function(app: &Shared, name: &str, args: &Args) -> Option<Result<Value>> {
    let upper = name.to_ascii_uppercase();
    if NOT_MEMBERS.contains(&upper.as_str()) {
        return None;
    }
    if args.is_empty() && !matches!(upper.as_str(), "PI") {
        return Some(Err(fail(name)));
    }
    Some(match run(app, name, args)? {
        Ok(Ok(v)) => Ok(v),
        Ok(Err(_)) => Err(fail(name)),
        Err(e) if e.number == 13 => Err(fail(name)),
        Err(e) => Err(e),
    })
}

/// `Application.<name>(...)`: the same functions, with error results returned as values.
pub(crate) fn call_function(app: &Shared, name: &str, args: &Args) -> Option<Result<Value>> {
    let upper = name.to_ascii_uppercase();
    if NOT_MEMBERS.contains(&upper.as_str()) {
        return None;
    }
    Some(match run(app, name, args)? {
        Ok(Ok(v)) => Ok(v),
        Ok(Err(e)) => Ok(Value::Error(e)),
        Err(e) => Err(e),
    })
}

/// `Evaluate(text)` / `[text]` on sheet `sheet`.
pub(crate) fn evaluate_text(app: &Shared, sheet: crate::workbook::SheetId, text: &str) -> Result<Value> {
    let v = with(app, |st| {
        let ctx = Ctx { book: st.book.as_ref(), sheet, cell: None, now: now() };
        Ok(eval::eval_text(text, &ctx))
    })?;
    Ok(match v {
        V::Ref(a) => range_value(app, a.sheet, a.rects, Kind::Cells),
        V::Arr(g) => grid_to_array(&g),
        V::Err(e) => Value::Error(e),
        other => eval::to_value(&other),
    })
}

impl Object for WorksheetFunctionObj {
    fn type_name(&self) -> String {
        "WorksheetFunction".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "application" | "parent" => Ok(crate::objects::app_value(&self.app)),
            _ => worksheet_function(&self.app, name, args).unwrap_or_else(|| Err(RuntimeError::no_member(name))),
        }
    }
}
