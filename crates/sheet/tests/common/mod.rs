#![allow(dead_code, clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use macroost_lang::{Args, Host, Result, Value};
use macroost_sheet::{MemoryWorkbook, SheetHost};

pub fn host_with(sheets: &[&str]) -> SheetHost {
    SheetHost::new(Box::new(MemoryWorkbook::with_sheets("Book1.xlsm", sheets)))
}

pub fn host() -> SheetHost {
    host_with(&["Sheet1"])
}

pub fn args(v: Vec<Value>) -> Args {
    Args::positional(v)
}

pub fn named(pos: Vec<Value>, named: Vec<(&str, Value)>) -> Args {
    Args { positional: pos, named: named.into_iter().map(|(n, v)| (n.to_owned(), v)).collect() }
}

pub fn g(h: &mut SheetHost, name: &str, a: Vec<Value>) -> Value {
    h.global(name, &args(a))
        .unwrap_or_else(|| panic!("global {name} not resolved"))
        .unwrap_or_else(|e| panic!("global {name}: {e}"))
}

pub fn k(h: &mut SheetHost, name: &str) -> Value {
    g(h, name, vec![])
}

pub fn try_call(v: &Value, name: &str, a: Args) -> Result<Value> {
    match v {
        Value::Object(Some(o)) => o.call(name, &a),
        other => panic!("not an object: {other:?} calling {name}"),
    }
}

pub fn call(v: &Value, name: &str, a: Vec<Value>) -> Value {
    try_call(v, name, args(a)).unwrap_or_else(|e| panic!("{name}: {e}"))
}

pub fn calln(v: &Value, name: &str, a: Args) -> Value {
    try_call(v, name, a).unwrap_or_else(|e| panic!("{name}: {e}"))
}

pub fn set(v: &Value, name: &str, value: Value) {
    try_set(v, name, value).unwrap_or_else(|e| panic!("set {name}: {e}"))
}

pub fn try_set(v: &Value, name: &str, value: Value) -> Result<()> {
    match v {
        Value::Object(Some(o)) => o.assign(name, &Args::none(), value),
        other => panic!("not an object: {other:?} assigning {name}"),
    }
}

pub fn rng(h: &mut SheetHost, addr: &str) -> Value {
    g(h, "Range", vec![addr.into()])
}

pub fn cells(h: &mut SheetHost, r: i32, c: i32) -> Value {
    g(h, "Cells", vec![r.into(), c.into()])
}

pub fn put(h: &mut SheetHost, addr: &str, v: Value) {
    let r = rng(h, addr);
    set(&r, "Value", v);
}

pub fn val(h: &mut SheetHost, addr: &str) -> Value {
    let r = rng(h, addr);
    call(&r, "Value", vec![])
}

pub fn text(v: &Value) -> String {
    v.to_text().unwrap_or_else(|e| panic!("to_text {v:?}: {e}"))
}

pub fn num(v: &Value) -> f64 {
    v.to_f64().unwrap_or_else(|e| panic!("to_f64 {v:?}: {e}"))
}

pub fn addr(v: &Value) -> String {
    text(&call(v, "Address", vec![]))
}

pub fn is_nothing(v: &Value) -> bool {
    matches!(v, Value::Object(None))
}

/// `Reset` from the oracle cases: clear Sheet1, activate it, select A1.
pub fn reset(h: &mut SheetHost) {
    let ws = g(h, "Worksheets", vec!["Sheet1".into()]);
    let all = call(&ws, "Cells", vec![]);
    call(&all, "Clear", vec![]);
    call(&ws, "Activate", vec![]);
    let a1 = call(&ws, "Range", vec!["A1".into()]);
    call(&a1, "Select", vec![]);
}

pub fn fill_a(h: &mut SheetHost, n: i32) {
    let ws = g(h, "Worksheets", vec!["Sheet1".into()]);
    for i in 1..=n {
        let c = call(&ws, "Cells", vec![i.into(), 1.into()]);
        set(&c, "Value", Value::Long(i));
    }
}
