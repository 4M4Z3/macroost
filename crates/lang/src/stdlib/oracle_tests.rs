//! Expectations recorded from the real host (black-box oracle run, `conformance/cases/lang/
//! expressions.tsv` -> `conformance/expected/lang.json`). Each case is "TypeName:CStr(value)" or
//! "ERR n". Expressions are hand-translated into operator and library calls, since the parser
//! lives elsewhere; literal types follow the language (`1` Integer, `1&` Long, `1.5` Double,
//! `1@` Currency, `#...#` Date).
//!
//! Not translated: C085 `Len(12345)` (a compile error), C098 (the case file's non-ASCII
//! character was lost in transit), C174-C177 (hex literals: lexer behaviour).

use super::*;
use crate::object::{Args, MISSING};
use crate::ops::{BinOp, BinOp::*, binary, negate, not};

struct Fixed;
impl Host for Fixed {
    fn global(&mut self, _name: &str, _args: &Args) -> Option<Result<Value>> {
        None
    }
    fn now(&mut self) -> f64 {
        46309.5
    }
}

fn f(name: &str, args: &[Value]) -> Result<Value> {
    call(name, args, &mut Fixed).unwrap_or_else(|| Err(RuntimeError::new(-99, "not a library function")))
}
fn op(o: BinOp, a: Result<Value>, b: Result<Value>) -> Result<Value> {
    binary(o, &a?, &b?, false)
}
fn int(n: i16) -> Result<Value> {
    Ok(Value::Integer(n))
}
fn lng(n: i32) -> Result<Value> {
    Ok(Value::Long(n))
}
fn dbl(x: f64) -> Result<Value> {
    Ok(Value::Double(x))
}
fn st(s: &str) -> Result<Value> {
    Ok(Value::str(s))
}
fn dt(y: i64, m: i64, d: i64, secs: f64) -> Result<Value> {
    Ok(Value::Date(ymd_to_serial(y, m, d) + secs / 86_400.0))
}
fn tm(secs: f64) -> Result<Value> {
    Ok(Value::Date(secs / 86_400.0))
}
fn fc(name: &str, args: &[Result<Value>]) -> Result<Value> {
    let vals = args.iter().cloned().collect::<Result<Vec<_>>>()?;
    f(name, &vals)
}

/// "TypeName:CStr" as the oracle records it.
fn show(r: Result<Value>) -> String {
    match r {
        Err(e) => format!("ERR {}", e.number),
        Ok(Value::Null) => "Null".into(),
        Ok(v) => {
            let tn = match f("TypeName", std::slice::from_ref(&v)) {
                Ok(Value::String(s)) => s.to_string(),
                other => format!("{other:?}"),
            };
            let text = match &v {
                Value::Array(_) => String::new(),
                other => other.to_text().unwrap_or_else(|e| format!("<err {}>", e.number)),
            };
            format!("{tn}:{text}")
        }
    }
}

fn index(a: Result<Value>, i: usize) -> Result<Value> {
    match a? {
        Value::Array(arr) => arr.borrow().data.get(i).cloned().ok_or_else(RuntimeError::subscript),
        _ => Err(RuntimeError::type_mismatch()),
    }
}

#[test]
fn lang_oracle() {
    let t = Ok(Value::Boolean(true));
    let empty = Ok(Value::Empty);
    let null = Ok(Value::Null);
    let text_cmp = lng(1);
    let d1014 = dt(2026, 10, 14, 0.0);
    let cases: Vec<(&str, Result<Value>, &str)> = vec![
        ("C001", op(Add, int(1), int(1)), "Integer:2"),
        ("C002", op(Add, int(32767), int(1)), "ERR 6"),
        ("C003", op(Add, lng(32767), int(1)), "Long:32768"),
        ("C004", op(Mul, int(300), int(200)), "ERR 6"),
        ("C005", op(Mul, lng(300), int(200)), "Long:60000"),
        ("C006", op(Add, lng(2147483647), int(1)), "ERR 6"),
        ("C007", op(Div, int(7), int(2)), "Double:3.5"),
        ("C008", op(IntDiv, int(7), int(2)), "Integer:3"),
        ("C009", op(IntDiv, int(-7), int(2)), "Integer:-3"),
        ("C010", op(Mod, int(7), int(3)), "Integer:1"),
        ("C011", op(Mod, int(-7), int(3)), "Integer:-1"),
        ("C012", op(Mod, dbl(7.5), int(2)), "Long:0"),
        ("C013", op(IntDiv, dbl(6.5), int(2)), "Long:3"),
        ("C014", op(Pow, int(2), int(10)), "Double:1024"),
        ("C015", negate(&op(Pow, int(2), int(2)).unwrap()), "Double:-4"),
        ("C016", op(Pow, int(-8), op(Div, int(1), int(3))), "ERR 5"),
        ("C017", op(Div, int(1), int(0)), "ERR 11"),
        ("C018", op(Div, int(0), int(0)), "ERR 6"),
        ("C019", op(IntDiv, int(5), int(0)), "ERR 11"),
        ("C020", op(Div, fc("CSng", &[int(1)]), fc("CSng", &[int(3)])), "Single:0.3333333"),
        ("C021", op(Div, fc("CInt", &[int(1)]), fc("CInt", &[int(3)])), "Double:0.333333333333333"),
        ("C022", op(Div, int(1), int(3)), "Double:0.333333333333333"),
        ("C023", op(Add, dbl(0.1), dbl(0.2)), "Double:0.3"),
        ("C024", dbl(1e15), "Double:1E+15"),
        ("C025", dbl(1e16), "Double:1E+16"),
        ("C026", dbl(123456789012345.0), "Double:123456789012345"),
        ("C027", dbl(1234567890123456.0), "Double:1.23456789012346E+15"),
        ("C028", dbl(0.0001), "Double:0.0001"),
        ("C029", dbl(0.00001), "Double:0.00001"),
        ("C030", dbl(0.000001), "Double:0.000001"),
        ("C031", dbl(1e-10), "Double:0.0000000001"),
        ("C032", dbl(-0.5), "Double:-0.5"),
        ("C033", fc("CDbl", &[st("1,234.5")]), "Double:1234.5"),
        ("C034", fc("CDbl", &[st(" 12 ")]), "Double:12"),
        ("C035", fc("CDbl", &[st("&H10")]), "Double:16"),
        ("C036", fc("CDbl", &[st("1e3")]), "Double:1000"),
        ("C037", fc("CDbl", &[st("$5")]), "Double:5"),
        ("C038", fc("CDbl", &[st("abc")]), "ERR 13"),
        ("C039", fc("CDbl", &[st("")]), "ERR 13"),
        ("C040", fc("CInt", &[dbl(2.5)]), "Integer:2"),
        ("C041", fc("CInt", &[dbl(3.5)]), "Integer:4"),
        ("C042", fc("CInt", &[dbl(-2.5)]), "Integer:-2"),
        ("C043", fc("CInt", &[dbl(32767.5)]), "ERR 6"),
        ("C044", fc("CLng", &[dbl(2.5)]), "Long:2"),
        ("C045", fc("CLng", &[st("1.5")]), "Long:2"),
        ("C046", fc("Int", &[dbl(-2.5)]), "Double:-3"),
        ("C047", fc("Fix", &[dbl(-2.5)]), "Double:-2"),
        ("C048", fc("Round", &[dbl(2.5)]), "Double:2"),
        ("C049", fc("Round", &[dbl(3.5)]), "Double:4"),
        ("C050", fc("Round", &[dbl(2.345), int(2)]), "Double:2.35"),
        ("C051", fc("Round", &[dbl(2.355), int(2)]), "Double:2.36"),
        ("C052", fc("Round", &[dbl(-2.5)]), "Double:-2"),
        ("C053", fc("Val", &[st("12abc")]), "Double:12"),
        ("C054", fc("Val", &[st("  1 2 3")]), "Double:123"),
        ("C055", fc("Val", &[st("&HFF")]), "Double:255"),
        ("C056", fc("Val", &[st("1e2x")]), "Double:100"),
        ("C057", fc("Val", &[st("")]), "Double:0"),
        ("C058", op(Add, st("10"), int(5)), "Double:15"),
        ("C059", op(Concat, st("10"), int(5)), "String:105"),
        ("C060", op(Add, st("a"), st("b")), "String:ab"),
        ("C061", op(Add, st("10"), st("5")), "String:105"),
        ("C062", op(Add, int(5), empty.clone()), "Integer:5"),
        ("C063", op(Concat, st("x"), null.clone()), "String:x"),
        ("C064", op(Add, null.clone(), int(1)), "Null"),
        ("C065", op(Add, t.clone(), int(1)), "Integer:0"),
        ("C066", op(And, t.clone(), int(3)), "Integer:3"),
        ("C067", op(And, int(5), int(3)), "Integer:1"),
        ("C068", op(Or, int(5), int(3)), "Integer:7"),
        ("C069", op(Xor, int(5), int(3)), "Integer:6"),
        ("C070", not(&Value::Integer(0)), "Integer:-1"),
        ("C071", not(&Value::Integer(5)), "Integer:-6"),
        ("C072", op(Eq, t.clone(), int(-1)), "Boolean:True"),
        ("C073", op(Lt, st("abc"), st("abd")), "Boolean:True"),
        ("C074", op(Lt, st("a"), st("B")), "Boolean:False"),
        ("C075", op(Lt, st("10"), st("9")), "Boolean:True"),
        ("C076", op(Lt, int(10), st("9")), "Boolean:False"),
        ("C077", op(Eq, empty.clone(), int(0)), "Boolean:True"),
        ("C078", op(Eq, empty.clone(), st("")), "Boolean:True"),
        ("C079", op(Like, st("abc"), st("a*")), "Boolean:True"),
        ("C080", op(Like, st("abc"), st("A*")), "Boolean:False"),
        ("C081", op(Like, st("a1c"), st("a#c")), "Boolean:True"),
        ("C082", op(Like, st("abc"), st("[a-c]bc")), "Boolean:True"),
        ("C083", op(Like, st("xbc"), st("[!a]bc")), "Boolean:True"),
        ("C084", fc("Len", &[st("hello")]), "Long:5"),
        ("C086", fc("Mid", &[st("Hello"), int(2), int(3)]), "String:ell"),
        ("C087", fc("Mid", &[st("Hello"), int(10)]), "String:"),
        ("C088", fc("Left", &[st("Hello"), int(2)]), "String:He"),
        ("C089", fc("Right", &[st("Hello"), int(10)]), "String:Hello"),
        ("C090", fc("InStr", &[st("abcabc"), st("c")]), "Long:3"),
        ("C091", fc("InStr", &[int(4), st("abcabc"), st("c")]), "Long:6"),
        ("C092", fc("InStr", &[int(1), st("abc"), st("B"), text_cmp.clone()]), "Long:2"),
        ("C093", fc("InStr", &[st("abc"), st("")]), "Long:1"),
        ("C094", fc("InStrRev", &[st("abcabc"), st("b")]), "Long:5"),
        ("C095", fc("Replace", &[st("aaa"), st("a"), st("bb")]), "String:bbbbbb"),
        (
            "C096",
            fc("Replace", &[st("Hello"), st("L"), st("x"), Ok(MISSING), Ok(MISSING), text_cmp.clone()]),
            "String:Hexxo",
        ),
        ("C097", fc("Trim", &[st("  a b  ")]), "String:a b"),
        ("C099", fc("StrComp", &[st("a"), st("B")]), "Integer:1"),
        ("C100", fc("StrComp", &[st("a"), st("B"), text_cmp.clone()]), "Integer:-1"),
        ("C101", fc("StrReverse", &[st("abc")]), "String:cba"),
        ("C102", op(Concat, fc("Space", &[int(3)]), st("|")), "String:   |"),
        ("C103", fc("String", &[int(3), st("ab")]), "String:aaa"),
        ("C104", fc("Chr", &[int(65)]), "String:A"),
        ("C105", fc("Asc", &[st("a")]), "Integer:97"),
        ("C106", index(fc("Split", &[st("a,b,,c"), st(",")]), 2), "String:"),
        ("C107", fc("UBound", &[fc("Split", &[st("a,b,,c"), st(",")])]), "Long:3"),
        ("C108", fc("Join", &[fc("Array", &[int(1), int(2), int(3)]), st("-")]), "String:1-2-3"),
        ("C109", fc("UBound", &[fc("Array", &[])]), "Long:-1"),
        ("C110", fc("Hex", &[int(255)]), "String:FF"),
        ("C111", fc("Hex", &[int(-1)]), "String:FFFF"),
        ("C112", fc("Oct", &[int(8)]), "String:10"),
        ("C113", fc("Str", &[int(5)]), "String: 5"),
        ("C114", fc("Str", &[int(-5)]), "String:-5"),
        ("C115", fc("CStr", &[int(5)]), "String:5"),
        ("C116", fc("CStr", &[t.clone()]), "String:True"),
        ("C117", fc("CStr", &[dt(2026, 1, 2, 0.0)]), "String:1/2/2026"),
        ("C118", fc("CStr", &[dt(2026, 1, 2, 15.0 * 3600.0 + 4.0 * 60.0 + 5.0)]), "String:1/2/2026 3:04:05 PM"),
        ("C119", fc("CStr", &[tm(3.0 * 3600.0 + 4.0 * 60.0 + 5.0)]), "String:3:04:05 AM"),
        ("C120", fc("CStr", &[fc("CDate", &[st("2026-10-14")])]), "String:10/14/2026"),
        ("C121", op(Add, d1014.clone(), int(1)), "Date:10/15/2026"),
        ("C122", op(Sub, d1014.clone(), dt(2026, 10, 1, 0.0)), "Double:13"),
        ("C123", fc("CDbl", &[d1014.clone()]), "Double:46309"),
        ("C124", fc("DateSerial", &[int(2026), int(13), int(1)]), "Date:1/1/2027"),
        ("C125", fc("DateSerial", &[int(2026), int(2), int(30)]), "Date:3/2/2026"),
        ("C126", fc("DateAdd", &[st("m"), int(1), dt(2026, 1, 31, 0.0)]), "Date:2/28/2026"),
        ("C127", fc("DateAdd", &[st("d"), int(-1), dt(2024, 3, 1, 0.0)]), "Date:2/29/2024"),
        ("C128", fc("DateDiff", &[st("d"), dt(2026, 1, 1, 0.0), dt(2026, 12, 31, 0.0)]), "Long:364"),
        ("C129", fc("DateDiff", &[st("m"), dt(2026, 1, 31, 0.0), dt(2026, 2, 1, 0.0)]), "Long:1"),
        ("C130", fc("Weekday", &[d1014.clone()]), "Integer:4"),
        (
            "C131",
            op(
                Concat,
                op(
                    Concat,
                    op(Concat, op(Concat, fc("Year", &[d1014.clone()]), st("/")), fc("Month", &[d1014.clone()])),
                    st("/"),
                ),
                fc("Day", &[d1014.clone()]),
            ),
            "String:2026/10/14",
        ),
        ("C132", fc("MonthName", &[int(10)]), "String:October"),
        ("C133", fc("WeekdayName", &[int(4)]), "String:Wednesday"),
        ("C134", fc("Format", &[dbl(1234.5), st("#,##0.00")]), "String:1,234.50"),
        ("C135", fc("Format", &[dbl(1234.5), st("0")]), "String:1235"),
        ("C136", fc("Format", &[dbl(0.256), st("0.0%")]), "String:25.6%"),
        ("C137", fc("Format", &[dbl(-1234.5), st("$#,##0.00")]), "String:-$1,234.50"),
        ("C138", fc("Format", &[dbl(1234.5), st("Currency")]), "String:$1,234.50"),
        ("C139", fc("Format", &[dbl(1234.5), st("Standard")]), "String:1,234.50"),
        ("C140", fc("Format", &[dbl(0.5), st("Percent")]), "String:50.00%"),
        ("C141", fc("Format", &[dbl(1234.5)]), "String:1234.5"),
        ("C142", fc("Format", &[d1014.clone(), st("yyyy-mm-dd")]), "String:2026-10-14"),
        ("C143", fc("Format", &[d1014.clone(), st("mmm d, yyyy")]), "String:Oct 14, 2026"),
        ("C144", fc("Format", &[d1014.clone(), st("dddd")]), "String:Wednesday"),
        (
            "C145",
            fc("Format", &[dt(2026, 10, 14, 15.0 * 3600.0 + 4.0 * 60.0 + 5.0), st("hh:nn AM/PM")]),
            "String:03:04 PM",
        ),
        ("C146", fc("Format", &[d1014.clone(), st("Long Date")]), "String:Wednesday, October 14, 2026"),
        ("C147", fc("Format", &[d1014.clone(), st("Short Date")]), "String:10/14/2026"),
        ("C148", fc("Format", &[int(5), st("00000")]), "String:00005"),
        ("C149", fc("Format", &[t.clone(), st("Yes/No")]), "String:Yes"),
        ("C150", fc("Format", &[st("abc"), st(">")]), "String:ABC"),
        ("C151", fc("FormatNumber", &[dbl(1234.567), int(2)]), "String:1,234.57"),
        ("C152", fc("FormatPercent", &[dbl(0.1234)]), "String:12.34%"),
        ("C153", fc("FormatCurrency", &[int(-5)]), "String:($5.00)"),
        ("C154", fc("IsNumeric", &[st("12.5")]), "Boolean:True"),
        ("C155", fc("IsNumeric", &[st("1,000")]), "Boolean:True"),
        ("C156", fc("IsNumeric", &[st("abc")]), "Boolean:False"),
        ("C157", fc("IsNumeric", &[st("")]), "Boolean:False"),
        ("C158", fc("IsNumeric", &[empty.clone()]), "Boolean:True"),
        ("C159", fc("IsDate", &[st("10/14/2026")]), "Boolean:True"),
        ("C160", fc("IsDate", &[st("hello")]), "Boolean:False"),
        ("C161", fc("IsEmpty", &[empty.clone()]), "Boolean:True"),
        ("C162", fc("VarType", &[int(1)]), "Long:2"),
        ("C163", fc("VarType", &[lng(1)]), "Long:3"),
        ("C164", fc("VarType", &[dbl(1.5)]), "Long:5"),
        ("C165", fc("VarType", &[st("a")]), "Long:8"),
        ("C166", fc("VarType", &[null.clone()]), "Long:1"),
        ("C167", fc("TypeName", &[int(1)]), "String:Integer"),
        ("C168", fc("TypeName", &[lng(1)]), "String:Long"),
        ("C169", fc("TypeName", &[dbl(1.5)]), "String:Double"),
        ("C170", fc("TypeName", &[Ok(Value::Currency(10_000))]), "String:Currency"),
        ("C171", fc("TypeName", &[empty.clone()]), "String:Empty"),
        ("C172", fc("TypeName", &[null.clone()]), "String:Null"),
        ("C173", fc("TypeName", &[fc("Array", &[int(1)])]), "String:Variant()"),
        ("C178", fc("IIf", &[op(Gt, int(1), int(2)), st("a"), st("b")]), "String:b"),
        ("C179", fc("Choose", &[int(2), st("a"), st("b"), st("c")]), "String:b"),
        ("C180", fc("Switch", &[op(Gt, int(1), int(2)), st("a"), op(Gt, int(2), int(1)), st("b")]), "String:b"),
        ("C181", fc("CBool", &[st("True")]), "Boolean:True"),
        ("C182", fc("CBool", &[st("yes")]), "ERR 13"),
        ("C183", fc("CBool", &[dbl(0.1)]), "Boolean:True"),
        ("C184", fc("CByte", &[int(255)]), "Byte:255"),
        ("C185", fc("CByte", &[int(256)]), "ERR 6"),
        ("C186", fc("CCur", &[dbl(1.23456)]), "Currency:1.2346"),
        ("C187", op(Div, fc("CCur", &[int(1)]), int(3)), "Double:0.333333333333333"),
        ("C188", fc("CDate", &[lng(45000)]), "Date:3/15/2023"),
        ("C189", fc("CDate", &[st("3:00 PM")]), "Date:3:00:00 PM"),
        ("C190", fc("Abs", &[int(-3)]), "Integer:3"),
        ("C191", fc("Sgn", &[int(-3)]), "Integer:-1"),
        ("C192", fc("Sqr", &[int(16)]), "Double:4"),
        ("C193", fc("Exp", &[int(1)]), "Double:2.71828182845905"),
        ("C194", fc("Log", &[fc("Exp", &[int(2)])]), "Double:2"),
        ("C195", op(Mul, fc("Atn", &[int(1)]), int(4)), "Double:3.14159265358979"),
    ];
    let total = cases.len();
    let failures: Vec<String> = cases
        .into_iter()
        .filter_map(|(id, got, want)| {
            let got = show(got);
            (got != want).then(|| format!("{id}: got {got:?}, want {want:?}"))
        })
        .collect();
    assert!(failures.is_empty(), "{} of {total} oracle cases differ:\n{}", failures.len(), failures.join("\n"));
    assert_eq!(total, 189);
}
