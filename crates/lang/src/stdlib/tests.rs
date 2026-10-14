use super::*;
use crate::object::{Args, MISSING};

/// A host whose clock is fixed at 2026-10-14 12:00:00 and which records dialogs.
struct TestHost {
    messages: Vec<(String, i32, String)>,
    input: Option<String>,
}

impl Host for TestHost {
    fn global(&mut self, _name: &str, _args: &Args) -> Option<Result<Value>> {
        None
    }
    fn msgbox(&mut self, prompt: &str, buttons: i32, title: &str) -> i32 {
        self.messages.push((prompt.to_owned(), buttons, title.to_owned()));
        6
    }
    fn inputbox(&mut self, _prompt: &str, _title: &str, _default: &str) -> Option<String> {
        self.input.clone()
    }
    fn now(&mut self) -> f64 {
        46309.5
    }
}

fn host() -> TestHost {
    TestHost { messages: Vec::new(), input: Some("typed".into()) }
}

fn run(name: &str, args: &[Value]) -> Result<Value> {
    call(name, args, &mut host()).unwrap_or_else(|| Err(RuntimeError::new(-99, "not a library function")))
}

/// Debug text of the result, or "err N".
fn ev(name: &str, args: &[Value]) -> String {
    match run(name, args) {
        Ok(v) => format!("{v:?}"),
        Err(e) => format!("err {}", e.number),
    }
}

/// The result as plain text (Strings unquoted).
fn txt(name: &str, args: &[Value]) -> String {
    match run(name, args) {
        Ok(Value::String(s)) => s.to_string(),
        Ok(v) => format!("{v:?}"),
        Err(e) => format!("err {}", e.number),
    }
}

fn s(v: &str) -> Value {
    Value::str(v)
}
fn i(n: i16) -> Value {
    Value::Integer(n)
}
fn l(n: i32) -> Value {
    Value::Long(n)
}
fn d(n: f64) -> Value {
    Value::Double(n)
}
fn date(y: i64, m: i64, dd: i64) -> Value {
    Value::Date(ymd_to_serial(y, m, dd))
}
fn arr(items: &[&str]) -> Value {
    array_value(Array::from_vec(0, items.iter().map(|x| s(x)).collect()))
}
fn items(v: &Value) -> Vec<String> {
    match v {
        Value::Array(a) => a.borrow().data.iter().map(|x| x.to_text().unwrap()).collect(),
        _ => panic!("not an array"),
    }
}

#[test]
fn unknown_names() {
    assert!(call("NoSuchFunction", &[], &mut host()).is_none());
    assert!(call("Nz", &[], &mut host()).is_none());
    assert!(call("LEN", &[s("abc")], &mut host()).is_some());
}

#[test]
fn conversions() {
    assert_eq!(ev("CInt", &[d(2.5)]), "Integer(2)");
    assert_eq!(ev("CInt", &[d(3.5)]), "Integer(4)");
    assert_eq!(ev("CInt", &[d(-2.5)]), "Integer(-2)");
    assert_eq!(ev("CInt", &[s("12.7")]), "Integer(13)");
    assert_eq!(ev("CInt", &[d(40000.0)]), "err 6");
    assert_eq!(ev("CInt", &[Value::Null]), "err 94");
    assert_eq!(ev("CInt", &[Value::Boolean(true)]), "Integer(-1)");
    assert_eq!(ev("CLng", &[d(2147483647.4)]), "Long(2147483647)");
    assert_eq!(ev("CLng", &[s("1,234")]), "Long(1234)");
    assert_eq!(ev("CLngLng", &[d(1e12)]), "LongLong(1000000000000)");
    assert_eq!(ev("CByte", &[i(255)]), "Byte(255)");
    assert_eq!(ev("CByte", &[i(256)]), "err 6");
    assert_eq!(ev("CByte", &[i(-1)]), "err 6");
    assert_eq!(ev("CDbl", &[s("1.5e3")]), "Double(1500)");
    assert_eq!(ev("CDbl", &[s("abc")]), "err 13");
    assert_eq!(ev("CSng", &[d(1.5)]), "Single(1.5)");
    assert_eq!(ev("CSng", &[d(1e300)]), "err 6");
    assert_eq!(ev("CCur", &[d(1.23456)]), "Currency(1.2346)");
    assert_eq!(ev("CCur", &[d(0.00005)]), "Currency(0)");
    assert_eq!(ev("CBool", &[i(5)]), "Boolean(true)");
    assert_eq!(ev("CBool", &[s("False")]), "Boolean(false)");
    assert_eq!(ev("CBool", &[s("x")]), "err 13");
    assert_eq!(ev("CStr", &[d(1.5)]), "\"1.5\"");
    assert_eq!(ev("CStr", &[Value::Boolean(true)]), "\"True\"");
    assert_eq!(ev("CStr", &[Value::Null]), "err 94");
    assert_eq!(ev("CStr", &[Value::Empty]), "\"\"");
    assert_eq!(ev("CDate", &[s("10/14/2026")]), "Date(46309)");
    assert_eq!(ev("CDate", &[s("2026-10-14 18:00")]), "Date(46309.75)");
    assert_eq!(ev("CDate", &[d(46309.5)]), "Date(46309.5)");
    assert_eq!(ev("CDate", &[s("not a date")]), "err 13");
    assert_eq!(ev("CDate", &[d(1e9)]), "err 6");
    assert_eq!(ev("CVErr", &[i(2042)]), "Error(2042)");
    assert_eq!(ev("CVar", &[i(1)]), "Integer(1)");
}

#[test]
fn val_str_hex_oct() {
    assert_eq!(ev("Val", &[s("  1615 198th Street N.E.")]), "Double(1615198)");
    assert_eq!(ev("Val", &[s("2457")]), "Double(2457)");
    assert_eq!(ev("Val", &[s(" 2 45 7")]), "Double(2457)");
    assert_eq!(ev("Val", &[s("24 and 57")]), "Double(24)");
    assert_eq!(ev("Val", &[s("&HFFFF")]), "Double(-1)");
    assert_eq!(ev("Val", &[s("&H10")]), "Double(16)");
    assert_eq!(ev("Val", &[s("&O17")]), "Double(15)");
    assert_eq!(ev("Val", &[s("-1.5e2x")]), "Double(-150)");
    assert_eq!(ev("Val", &[s("1,000")]), "Double(1)");
    assert_eq!(ev("Val", &[s("abc")]), "Double(0)");
    assert_eq!(ev("Val", &[s("")]), "Double(0)");
    assert_eq!(ev("Val", &[s(".5")]), "Double(0.5)");
    assert_eq!(ev("Str", &[i(459)]), "\" 459\"");
    assert_eq!(ev("Str", &[i(-459)]), "\"-459\"");
    assert_eq!(ev("Str", &[d(459.001)]), "\" 459.001\"");
    assert_eq!(ev("Str", &[Value::Null]), "Null");
    assert_eq!(ev("Str$", &[Value::Null]), "err 94");
    assert_eq!(ev("Hex", &[i(255)]), "\"FF\"");
    assert_eq!(ev("Hex", &[i(-1)]), "\"FFFF\"");
    assert_eq!(ev("Hex", &[l(-32768)]), "\"FFFF8000\"");
    assert_eq!(ev("Hex", &[l(-70000)]), "\"FFFEEE90\"");
    assert_eq!(ev("Hex", &[Value::LongLong(-1)]), "\"FFFFFFFFFFFFFFFF\"");
    assert_eq!(ev("Hex", &[Value::Empty]), "\"0\"");
    assert_eq!(ev("Hex", &[Value::Null]), "Null");
    assert_eq!(ev("Hex$", &[Value::Null]), "err 94");
    assert_eq!(ev("Hex", &[d(10.5)]), "\"A\"");
    assert_eq!(ev("Oct", &[i(8)]), "\"10\"");
    assert_eq!(ev("Oct", &[i(-1)]), "\"177777\"");
}

#[test]
fn fix_and_int() {
    assert_eq!(ev("Fix", &[d(-2.5)]), "Double(-2)");
    assert_eq!(ev("Int", &[d(-2.5)]), "Double(-3)");
    assert_eq!(ev("Int", &[d(2.9)]), "Double(2)");
    assert_eq!(ev("Fix", &[i(7)]), "Integer(7)");
    assert_eq!(ev("Int", &[s("-8.4")]), "Double(-9)");
    assert_eq!(ev("Int", &[Value::Currency(-25_000)]), "Currency(-3)");
    assert_eq!(ev("Fix", &[Value::Currency(-25_000)]), "Currency(-2)");
    assert_eq!(ev("Int", &[Value::Date(46309.75)]), "Date(46309)");
    assert_eq!(ev("Int", &[Value::Null]), "Null");
    assert_eq!(ev("Int", &[Value::Single(1.5)]), "Single(1)");
}

#[test]
fn string_basics() {
    assert_eq!(ev("Len", &[s("Hello")]), "Long(5)");
    assert_eq!(ev("Len", &[Value::Null]), "Null");
    assert_eq!(ev("Len", &[i(123)]), "Long(3)");
    assert_eq!(ev("Len", &[Value::Empty]), "Long(0)");
    assert_eq!(ev("LenB", &[s("abc")]), "Long(6)");
    assert_eq!(txt("Left", &[s("Hello"), i(2)]), "He");
    assert_eq!(txt("Left", &[s("Hello"), i(20)]), "Hello");
    assert_eq!(ev("Left", &[s("Hello"), i(-1)]), "err 5");
    assert_eq!(ev("Left", &[Value::Null, i(1)]), "Null");
    assert_eq!(ev("Left$", &[Value::Null, i(1)]), "err 94");
    assert_eq!(txt("Right", &[s("Hello"), i(3)]), "llo");
    assert_eq!(txt("Mid", &[s("Hello"), i(2), i(3)]), "ell");
    assert_eq!(txt("Mid", &[s("Hello"), i(2)]), "ello");
    assert_eq!(txt("Mid", &[s("Hello"), i(9)]), "");
    assert_eq!(txt("Mid", &[s("Hello"), i(4), i(100)]), "lo");
    assert_eq!(ev("Mid", &[s("Hello"), i(0)]), "err 5");
    assert_eq!(txt("Mid", &[s("Hello"), i(2), MISSING]), "ello");
    assert_eq!(txt("Trim", &[s("  a b  ")]), "a b");
    assert_eq!(txt("LTrim", &[s("  a  ")]), "a  ");
    assert_eq!(txt("RTrim", &[s("  a  ")]), "  a");
    assert_eq!(txt("Trim", &[s("\ta\t")]), "\ta\t");
    assert_eq!(txt("UCase", &[s("abc")]), "ABC");
    assert_eq!(txt("LCase", &[s("ÀBC")]), "àbc");
    assert_eq!(ev("UCase", &[Value::Null]), "Null");
    assert_eq!(txt("StrReverse", &[s("abc")]), "cba");
    assert_eq!(txt("Space", &[i(3)]), "   ");
    assert_eq!(txt("String", &[i(3), s("xyz")]), "xxx");
    assert_eq!(txt("String", &[i(3), i(65)]), "AAA");
    assert_eq!(txt("String", &[i(2), i(321)]), "AA");
    assert_eq!(ev("String", &[i(2), s("")]), "err 5");
    assert_eq!(ev("Space", &[i(-1)]), "err 5");
}

#[test]
fn searching() {
    assert_eq!(ev("InStr", &[s("abc"), s("c")]), "Long(3)");
    assert_eq!(ev("InStr", &[l(1), s("abc"), s("B"), l(1)]), "Long(2)");
    assert_eq!(ev("InStr", &[l(1), s("abc"), s("B"), l(0)]), "Long(0)");
    assert_eq!(ev("InStr", &[l(1), s("abc"), s("B")]), "Long(0)");
    assert_eq!(ev("InStr", &[l(3), s("abcabc"), s("b")]), "Long(5)");
    assert_eq!(ev("InStr", &[s("abc"), s("")]), "Long(1)");
    assert_eq!(ev("InStr", &[s(""), s("a")]), "Long(0)");
    assert_eq!(ev("InStr", &[l(0), s("abc"), s("a")]), "err 5");
    assert_eq!(ev("InStr", &[Value::Null, s("a")]), "Null");
    assert_eq!(ev("InStr", &[l(9), s("abc"), s("a")]), "Long(0)");
    assert_eq!(ev("InStrRev", &[s("abcabc"), s("b")]), "Long(5)");
    assert_eq!(ev("InStrRev", &[s("abcabc"), s("b"), l(4)]), "Long(2)");
    assert_eq!(ev("InStrRev", &[s("abcabc"), s("B"), l(-1), l(1)]), "Long(5)");
    assert_eq!(ev("InStrRev", &[s("abc"), s("z")]), "Long(0)");
    assert_eq!(ev("InStrRev", &[s("abc"), s("bc"), l(2)]), "Long(0)");
    assert_eq!(txt("Replace", &[s("aXbXc"), s("X"), s("-")]), "a-b-c");
    assert_eq!(txt("Replace", &[s("abcabc"), s("a"), s("x"), l(3)]), "cxbc");
    assert_eq!(txt("Replace", &[s("aaaa"), s("a"), s("b"), l(1), l(2)]), "bbaa");
    assert_eq!(txt("Replace", &[s("AbAb"), s("a"), s("x"), l(1), l(-1), l(1)]), "xbxb");
    assert_eq!(txt("Replace", &[s("AbAb"), s("a"), s("x")]), "AbAb");
    assert_eq!(txt("Replace", &[s("abc"), s(""), s("x")]), "abc");
    assert_eq!(ev("StrComp", &[s("a"), s("B")]), "Integer(1)");
    assert_eq!(ev("StrComp", &[s("a"), s("B"), l(1)]), "Integer(-1)");
    assert_eq!(ev("StrComp", &[s("abc"), s("ABC"), l(1)]), "Integer(0)");
    assert_eq!(ev("StrComp", &[s("a"), Value::Null]), "Null");
    assert_eq!(txt("StrConv", &[s("hello wORLD"), l(3)]), "Hello World");
    assert_eq!(txt("StrConv", &[s("abc"), l(1)]), "ABC");
}

#[test]
fn characters() {
    assert_eq!(txt("Chr", &[i(65)]), "A");
    assert_eq!(txt("Chr", &[i(128)]), "€");
    assert_eq!(txt("Chr", &[i(233)]), "é");
    assert_eq!(ev("Chr", &[i(256)]), "err 5");
    assert_eq!(txt("ChrW", &[l(8364)]), "€");
    assert_eq!(txt("ChrW", &[l(-1)]), "\u{FFFF}");
    assert_eq!(ev("Asc", &[s("A")]), "Integer(65)");
    assert_eq!(ev("Asc", &[s("€")]), "Integer(128)");
    assert_eq!(ev("Asc", &[s("Ω")]), "Integer(63)");
    assert_eq!(ev("Asc", &[s("")]), "err 5");
    assert_eq!(ev("AscW", &[s("€")]), "Integer(8364)");
    assert_eq!(ev("AscW", &[s("\u{FFFF}")]), "Integer(-1)");
}

#[test]
fn split_join_filter() {
    let v = run("Split", &[s("a,b,,c"), s(",")]).unwrap();
    assert_eq!(items(&v), ["a", "b", "", "c"]);
    let v = run("Split", &[s("a b c")]).unwrap();
    assert_eq!(items(&v), ["a", "b", "c"]);
    let v = run("Split", &[s("a,b,c"), s(","), l(2)]).unwrap();
    assert_eq!(items(&v), ["a", "b,c"]);
    let v = run("Split", &[s("aXbxc"), s("x"), l(-1), l(1)]).unwrap();
    assert_eq!(items(&v), ["a", "b", "c"]);
    let v = run("Split", &[s(""), s(",")]).unwrap();
    assert_eq!(ev("UBound", std::slice::from_ref(&v)), "Long(-1)");
    assert_eq!(ev("LBound", &[v]), "Long(0)");
    let v = run("Split", &[s("abc"), s("")]).unwrap();
    assert_eq!(items(&v), ["abc"]);
    assert_eq!(txt("Join", &[arr(&["a", "b", "c"]), s("-")]), "a-b-c");
    assert_eq!(txt("Join", &[arr(&["a", "b"])]), "a b");
    assert_eq!(txt("Join", &[arr(&["a", "b"]), s("")]), "ab");
    assert_eq!(ev("Join", &[s("x")]), "err 13");
    let v = run("Filter", &[arr(&["apple", "banana", "cherry"]), s("an")]).unwrap();
    assert_eq!(items(&v), ["banana"]);
    let v = run("Filter", &[arr(&["apple", "banana", "cherry"]), s("an"), Value::Boolean(false)]).unwrap();
    assert_eq!(items(&v), ["apple", "cherry"]);
    let v = run("Filter", &[arr(&["Apple", "b"]), s("a"), Value::Boolean(true), l(1)]).unwrap();
    assert_eq!(items(&v), ["Apple"]);
}

#[test]
fn formats() {
    assert_eq!(txt("Format", &[d(1234.5), s("#,##0.00")]), "1,234.50");
    assert_eq!(txt("Format", &[d(0.256), s("0.0%")]), "25.6%");
    assert_eq!(txt("Format", &[d(1234.5)]), "1234.5");
    assert_eq!(txt("Format", &[d(5.0), s("0.00")]), "5.00");
    assert_eq!(txt("Format", &[d(1234.5), s("Currency")]), "$1,234.50");
    assert_eq!(txt("Format", &[date(2026, 10, 14), s("dddd, mmmm d, yyyy")]), "Wednesday, October 14, 2026");
    assert_eq!(txt("Format", &[date(2026, 10, 14), s("Short Date")]), "10/14/2026");
    assert_eq!(txt("Format", &[s("abc"), s(">")]), "ABC");
    assert_eq!(ev("Format", &[Value::Null]), "Null");
    assert_eq!(ev("Format$", &[Value::Null]), "err 94");
    assert_eq!(txt("Format", &[Value::Boolean(true)]), "True");
    assert_eq!(txt("Format", &[date(2026, 1, 4), s("ww")]), "2");
    assert_eq!(txt("Format", &[date(2026, 10, 14), s("w"), l(2)]), "3");
    assert_eq!(txt("FormatNumber", &[d(1234.567)]), "1,234.57");
    assert_eq!(txt("FormatNumber", &[d(1234.567), l(1)]), "1,234.6");
    assert_eq!(txt("FormatNumber", &[d(0.5), l(2), l(0)]), ".50");
    assert_eq!(txt("FormatNumber", &[d(-5.0), l(0), l(-2), l(-1)]), "(5)");
    assert_eq!(txt("FormatNumber", &[d(1234.5), l(0), l(-2), l(-2), l(0)]), "1235");
    assert_eq!(ev("FormatNumber", &[s("x")]), "err 13");
    assert_eq!(txt("FormatCurrency", &[d(1234.5)]), "$1,234.50");
    assert_eq!(txt("FormatCurrency", &[d(-1234.5)]), "($1,234.50)");
    assert_eq!(txt("FormatPercent", &[d(0.256)]), "25.60%");
    assert_eq!(txt("FormatPercent", &[d(0.256), l(0)]), "26%");
    assert_eq!(txt("FormatDateTime", &[Value::Date(46309.75)]), "10/14/2026 6:00:00 PM");
    assert_eq!(txt("FormatDateTime", &[Value::Date(46309.75), l(2)]), "10/14/2026");
    assert_eq!(txt("FormatDateTime", &[Value::Date(46309.75), l(3)]), "6:00:00 PM");
    assert_eq!(txt("FormatDateTime", &[Value::Date(46309.75), l(4)]), "18:00");
}

#[test]
fn math_functions() {
    assert_eq!(ev("Abs", &[i(-5)]), "Integer(5)");
    assert_eq!(ev("Abs", &[i(-32768)]), "err 6");
    assert_eq!(ev("Abs", &[d(-2.5)]), "Double(2.5)");
    assert_eq!(ev("Abs", &[Value::Null]), "Null");
    assert_eq!(ev("Abs", &[s("-3")]), "Double(3)");
    assert_eq!(ev("Sgn", &[d(-0.1)]), "Integer(-1)");
    assert_eq!(ev("Sgn", &[i(0)]), "Integer(0)");
    assert_eq!(ev("Sgn", &[l(7)]), "Integer(1)");
    assert_eq!(ev("Sqr", &[i(16)]), "Double(4)");
    assert_eq!(ev("Sqr", &[i(-1)]), "err 5");
    assert_eq!(ev("Exp", &[i(0)]), "Double(1)");
    assert_eq!(ev("Exp", &[i(1000)]), "err 6");
    assert_eq!(ev("Log", &[i(1)]), "Double(0)");
    assert_eq!(ev("Log", &[i(0)]), "err 5");
    assert_eq!(ev("Sin", &[i(0)]), "Double(0)");
    assert_eq!(ev("Cos", &[i(0)]), "Double(1)");
    assert_eq!(ev("Tan", &[i(0)]), "Double(0)");
    let pi = run("Atn", &[i(1)]).unwrap().to_f64().unwrap() * 4.0;
    assert!((pi - std::f64::consts::PI).abs() < 1e-15);
    assert_eq!(ev("Round", &[d(2.5)]), "Double(2)");
    assert_eq!(ev("Round", &[d(3.5)]), "Double(4)");
    assert_eq!(ev("Round", &[d(-2.5)]), "Double(-2)");
    assert_eq!(ev("Round", &[d(2.345), i(2)]), "Double(2.35)");
    assert_eq!(ev("Round", &[d(2.355), i(2)]), "Double(2.36)");
    assert_eq!(ev("Round", &[d(0.125), i(2)]), "Double(0.12)");
    assert_eq!(ev("Round", &[d(1.2345678), i(4)]), "Double(1.2346)");
    assert_eq!(ev("Round", &[i(7)]), "Integer(7)");
    assert_eq!(ev("Round", &[Value::Currency(12_345), i(2)]), "Currency(1.23)");
    assert_eq!(ev("Round", &[Value::Currency(12_250), i(2)]), "Currency(1.22)");
    assert_eq!(ev("Round", &[d(1.5), i(-1)]), "err 5");
    assert_eq!(ev("Round", &[Value::Null]), "Null");
    assert_eq!(ev("Round", &[s("2.5")]), "Double(2)");
}

#[test]
fn random_numbers() {
    // A negative argument reseeds deterministically.
    let a = run("Rnd", &[l(-1)]).unwrap().to_f64().unwrap();
    let b = run("Rnd", &[l(-1)]).unwrap().to_f64().unwrap();
    assert_eq!(a, b);
    assert_eq!(run("Rnd", &[l(0)]).unwrap().to_f64().unwrap(), a);
    let mut seen = Vec::new();
    for _ in 0..100 {
        let x = run("Rnd", &[]).unwrap().to_f64().unwrap();
        assert!((0.0..1.0).contains(&x));
        seen.push(x);
    }
    seen.dedup();
    assert!(seen.len() > 90);
    assert_eq!(ev("Randomize", &[l(42)]), "Empty");
    assert_eq!(ev("Randomize", &[]), "Empty");
}

#[test]
fn clock_functions() {
    assert_eq!(ev("Now", &[]), "Date(46309.5)");
    assert_eq!(ev("Date", &[]), "Date(46309)");
    assert_eq!(ev("Time", &[]), "Date(0.5)");
    assert_eq!(ev("Timer", &[]), "Single(43200)");
    assert_eq!(ev("Date$", &[]), "\"10-14-2026\"");
    assert_eq!(ev("Time$", &[]), "\"12:00:00\"");
    assert_eq!(ev("Now", &[i(1)]), "err 450");
}

#[test]
fn date_parts() {
    let dt = Value::Date(46309.0 + (13.0 * 3600.0 + 45.0 * 60.0 + 30.0) / 86400.0);
    assert_eq!(ev("Year", std::slice::from_ref(&dt)), "Integer(2026)");
    assert_eq!(ev("Month", std::slice::from_ref(&dt)), "Integer(10)");
    assert_eq!(ev("Day", std::slice::from_ref(&dt)), "Integer(14)");
    assert_eq!(ev("Hour", std::slice::from_ref(&dt)), "Integer(13)");
    assert_eq!(ev("Minute", std::slice::from_ref(&dt)), "Integer(45)");
    assert_eq!(ev("Second", std::slice::from_ref(&dt)), "Integer(30)");
    assert_eq!(ev("Year", &[s("2026-10-14")]), "Integer(2026)");
    assert_eq!(ev("Year", &[Value::Null]), "Null");
    assert_eq!(ev("Year", &[s("garbage")]), "err 13");
    assert_eq!(ev("Weekday", &[date(2026, 10, 14)]), "Integer(4)");
    assert_eq!(ev("Weekday", &[date(2026, 10, 14), l(2)]), "Integer(3)");
    assert_eq!(ev("Weekday", &[date(2026, 10, 14), l(9)]), "err 5");
    assert_eq!(txt("WeekdayName", &[l(1)]), "Sunday");
    assert_eq!(txt("WeekdayName", &[l(1), Value::Boolean(true)]), "Sun");
    assert_eq!(txt("WeekdayName", &[l(1), Value::Boolean(false), l(2)]), "Monday");
    assert_eq!(ev("WeekdayName", &[l(8)]), "err 5");
    assert_eq!(txt("MonthName", &[l(10)]), "October");
    assert_eq!(txt("MonthName", &[l(2), Value::Boolean(true)]), "Feb");
    assert_eq!(ev("MonthName", &[l(13)]), "err 5");
}

#[test]
fn date_construction() {
    assert_eq!(ev("DateSerial", &[l(2026), l(13), l(1)]), format!("Date({})", ymd_to_serial(2027, 1, 1)));
    assert_eq!(ev("DateSerial", &[l(2026), l(3), l(0)]), format!("Date({})", ymd_to_serial(2026, 2, 28)));
    assert_eq!(ev("DateSerial", &[l(2026), l(1), l(-1)]), format!("Date({})", ymd_to_serial(2025, 12, 30)));
    assert_eq!(ev("DateSerial", &[l(26), l(1), l(1)]), format!("Date({})", ymd_to_serial(2026, 1, 1)));
    assert_eq!(ev("DateSerial", &[l(95), l(1), l(1)]), format!("Date({})", ymd_to_serial(1995, 1, 1)));
    assert_eq!(ev("DateSerial", &[l(10000), l(1), l(1)]), "err 5");
    assert_eq!(ev("TimeSerial", &[l(18), l(0), l(0)]), "Date(0.75)");
    assert_eq!(ev("TimeSerial", &[l(12), l(-30), l(0)]), format!("Date({})", 11.5 / 24.0));
    assert_eq!(ev("DateValue", &[s("10/14/2026 6:00 PM")]), "Date(46309)");
    assert_eq!(ev("TimeValue", &[s("6:00 PM")]), "Date(0.75)");
    assert_eq!(ev("TimeValue", &[s("10/14/2026 6:00 PM")]), "Date(0.75)");
    assert_eq!(ev("IsDate", &[s("10/14/2026")]), "Boolean(true)");
    assert_eq!(ev("IsDate", &[s("February 30, 2026")]), "Boolean(false)");
    assert_eq!(ev("IsDate", &[date(2026, 1, 1)]), "Boolean(true)");
    assert_eq!(ev("IsDate", &[d(46309.0)]), "Boolean(false)");
    assert_eq!(ev("IsDate", &[s("6:30 PM")]), "Boolean(true)");
}

#[test]
fn date_arithmetic() {
    let base = date(2026, 1, 31);
    let serial = |y, m, dd| format!("Date({})", ymd_to_serial(y, m, dd));
    assert_eq!(ev("DateAdd", &[s("m"), l(1), base.clone()]), serial(2026, 2, 28));
    assert_eq!(ev("DateAdd", &[s("m"), l(-2), base.clone()]), serial(2025, 11, 30));
    assert_eq!(ev("DateAdd", &[s("yyyy"), l(1), date(2024, 2, 29)]), serial(2025, 2, 28));
    assert_eq!(ev("DateAdd", &[s("q"), l(1), base.clone()]), serial(2026, 4, 30));
    assert_eq!(ev("DateAdd", &[s("d"), l(1), base.clone()]), serial(2026, 2, 1));
    assert_eq!(ev("DateAdd", &[s("ww"), l(2), base.clone()]), serial(2026, 2, 14));
    assert_eq!(ev("DateAdd", &[s("h"), l(36), base.clone()]), format!("Date({})", ymd_to_serial(2026, 2, 1) + 0.5));
    assert_eq!(ev("DateAdd", &[s("n"), l(90), Value::Date(0.0)]), format!("Date({})", 1.5 / 24.0));
    assert_eq!(ev("DateAdd", &[s("s"), l(30), Value::Date(0.0)]), format!("Date({})", 30.0 / 86400.0));
    assert_eq!(ev("DateAdd", &[s("x"), l(1), base.clone()]), "err 5");
    assert_eq!(ev("DateAdd", &[s("yyyy"), l(9000), base.clone()]), "err 5");
    assert_eq!(ev("DateDiff", &[s("d"), date(2026, 1, 1), date(2026, 12, 31)]), "Long(364)");
    assert_eq!(ev("DateDiff", &[s("yyyy"), date(2025, 12, 31), date(2026, 1, 1)]), "Long(1)");
    assert_eq!(ev("DateDiff", &[s("m"), date(2026, 1, 31), date(2026, 2, 1)]), "Long(1)");
    assert_eq!(ev("DateDiff", &[s("q"), date(2026, 3, 31), date(2026, 4, 1)]), "Long(1)");
    assert_eq!(ev("DateDiff", &[s("w"), date(2026, 10, 1), date(2026, 10, 14)]), "Long(1)");
    // Oct 10 2026 is a Saturday, Oct 11 a Sunday: one week boundary.
    assert_eq!(ev("DateDiff", &[s("ww"), date(2026, 10, 10), date(2026, 10, 11)]), "Long(1)");
    assert_eq!(ev("DateDiff", &[s("d"), date(2026, 12, 31), date(2026, 1, 1)]), "Long(-364)");
    assert_eq!(ev("DateDiff", &[s("h"), Value::Date(0.99), Value::Date(1.01)]), "Long(1)");
    assert_eq!(ev("DateDiff", &[s("n"), Value::Date(0.0), Value::Date(0.5)]), "Long(720)");
    assert_eq!(ev("DateDiff", &[s("s"), Value::Date(0.0), Value::Date(1.0)]), "Long(86400)");
    assert_eq!(ev("DatePart", &[s("yyyy"), date(2026, 10, 14)]), "Integer(2026)");
    assert_eq!(ev("DatePart", &[s("q"), date(2026, 10, 14)]), "Integer(4)");
    assert_eq!(ev("DatePart", &[s("y"), date(2026, 10, 14)]), "Integer(287)");
    assert_eq!(ev("DatePart", &[s("w"), date(2026, 10, 14)]), "Integer(4)");
    assert_eq!(ev("DatePart", &[s("ww"), date(2026, 1, 4)]), "Integer(2)");
    assert_eq!(ev("DatePart", &[s("ww"), date(2026, 1, 1), l(1), l(3)]), "Integer(52)");
    assert_eq!(ev("DatePart", &[s("h"), Value::Date(46309.75)]), "Integer(18)");
}

#[test]
fn information() {
    assert_eq!(ev("IsNumeric", &[s("12.5")]), "Boolean(true)");
    assert_eq!(ev("IsNumeric", &[s("$1,000")]), "Boolean(true)");
    assert_eq!(ev("IsNumeric", &[s("12a")]), "Boolean(false)");
    assert_eq!(ev("IsNumeric", &[s("")]), "Boolean(false)");
    assert_eq!(ev("IsNumeric", &[Value::Empty]), "Boolean(true)");
    assert_eq!(ev("IsNumeric", &[Value::Null]), "Boolean(false)");
    assert_eq!(ev("IsNumeric", &[date(2026, 1, 1)]), "Boolean(false)");
    assert_eq!(ev("IsNumeric", &[Value::Boolean(true)]), "Boolean(true)");
    assert_eq!(ev("IsEmpty", &[Value::Empty]), "Boolean(true)");
    assert_eq!(ev("IsEmpty", &[s("")]), "Boolean(false)");
    assert_eq!(ev("IsNull", &[Value::Null]), "Boolean(true)");
    assert_eq!(ev("IsNull", &[i(0)]), "Boolean(false)");
    assert_eq!(ev("IsArray", &[arr(&["a"])]), "Boolean(true)");
    assert_eq!(ev("IsArray", &[s("a")]), "Boolean(false)");
    assert_eq!(ev("IsObject", &[Value::Object(None)]), "Boolean(true)");
    assert_eq!(ev("IsObject", &[i(1)]), "Boolean(false)");
    assert_eq!(ev("IsError", &[Value::Error(2042)]), "Boolean(true)");
    assert_eq!(ev("IsError", &[i(1)]), "Boolean(false)");
    assert_eq!(ev("IsMissing", &[MISSING]), "Boolean(true)");
    assert_eq!(ev("IsMissing", &[i(1)]), "Boolean(false)");
    assert_eq!(txt("TypeName", &[i(1)]), "Integer");
    assert_eq!(txt("TypeName", &[s("x")]), "String");
    assert_eq!(txt("TypeName", &[Value::Object(None)]), "Nothing");
    assert_eq!(txt("TypeName", &[Value::Empty]), "Empty");
    assert_eq!(txt("TypeName", &[arr(&["a"])]), "Variant()");
    assert_eq!(ev("VarType", &[s("x")]), "Long(8)");
    assert_eq!(ev("VarType", &[d(1.0)]), "Long(5)");
    assert_eq!(ev("VarType", &[Value::Null]), "Long(1)");
    assert_eq!(ev("VarType", &[Value::Object(None)]), "Long(9)");
    assert_eq!(ev("RGB", &[l(255), l(0), l(0)]), "Long(255)");
    assert_eq!(ev("RGB", &[l(0), l(0), l(255)]), "Long(16711680)");
    assert_eq!(ev("RGB", &[l(300), l(1), l(0)]), "Long(511)");
    assert_eq!(ev("QBColor", &[l(12)]), "Long(255)");
    assert_eq!(ev("QBColor", &[l(16)]), "err 5");
    assert_eq!(txt("Error", &[l(11)]), "Division by zero");
}

#[test]
fn arrays() {
    let a = run("Array", &[i(1), s("two"), d(3.0)]).unwrap();
    assert_eq!(ev("UBound", std::slice::from_ref(&a)), "Long(2)");
    assert_eq!(ev("LBound", std::slice::from_ref(&a)), "Long(0)");
    assert_eq!(ev("UBound", &[a.clone(), l(2)]), "err 9");
    assert_eq!(ev("UBound", &[i(1)]), "err 13");
    assert_eq!(txt("TypeName", &[a]), "Variant()");
    let grid = array_value(Array::new(vec![(1, 3), (0, 4)]).unwrap());
    assert_eq!(ev("UBound", &[grid.clone(), l(1)]), "Long(3)");
    assert_eq!(ev("UBound", &[grid.clone(), l(2)]), "Long(4)");
    assert_eq!(ev("LBound", &[grid.clone(), l(1)]), "Long(1)");
    assert_eq!(ev("LBound", &[grid, l(0)]), "err 9");
    let empty = array_value(Array::default());
    assert_eq!(ev("UBound", &[empty]), "err 9");
    let none = run("Array", &[]).unwrap();
    assert_eq!(ev("UBound", &[none]), "Long(-1)");
}

#[test]
fn interaction() {
    let mut h = host();
    let r = call("MsgBox", &[s("Save?"), l(4), s("Title")], &mut h).unwrap().unwrap();
    assert_eq!(format!("{r:?}"), "Long(6)");
    assert_eq!(h.messages, [("Save?".to_owned(), 4, "Title".to_owned())]);
    let r = call("InputBox", &[s("Name?")], &mut h).unwrap().unwrap();
    assert_eq!(format!("{r:?}"), "\"typed\"");
    h.input = None;
    let r = call("InputBox", &[s("Name?")], &mut h).unwrap().unwrap();
    assert_eq!(format!("{r:?}"), "\"\"");
    assert_eq!(txt("Environ", &[s("PATH")]), "");
    assert_eq!(txt("Environ$", &[s("USERNAME")]), "");
    assert_eq!(txt("IIf", &[Value::Boolean(true), s("yes"), s("no")]), "yes");
    assert_eq!(txt("IIf", &[i(0), s("yes"), s("no")]), "no");
    assert_eq!(txt("IIf", &[Value::Null, s("yes"), s("no")]), "no");
    assert_eq!(txt("Choose", &[i(2), s("a"), s("b"), s("c")]), "b");
    assert_eq!(ev("Choose", &[i(4), s("a"), s("b"), s("c")]), "Null");
    assert_eq!(ev("Choose", &[i(0), s("a")]), "Null");
    assert_eq!(txt("Switch", &[Value::Boolean(false), s("a"), Value::Boolean(true), s("b")]), "b");
    assert_eq!(ev("Switch", &[Value::Boolean(false), s("a")]), "Null");
    assert_eq!(ev("Switch", &[Value::Boolean(false)]), "err 5");
    assert_eq!(ev("DoEvents", &[]), "Integer(0)");
}

#[test]
fn constants() {
    assert_eq!(format!("{:?}", constant("vbCrLf").unwrap()), "\"\\r\\n\"");
    assert_eq!(format!("{:?}", constant("VBTAB").unwrap()), "\"\\t\"");
    assert_eq!(format!("{:?}", constant("vbNullString").unwrap()), "\"\"");
    assert_eq!(format!("{:?}", constant("vbYes").unwrap()), "Long(6)");
    assert_eq!(format!("{:?}", constant("vbOKOnly").unwrap()), "Long(0)");
    assert_eq!(format!("{:?}", constant("vbYesNoCancel").unwrap()), "Long(3)");
    assert_eq!(format!("{:?}", constant("vbBinaryCompare").unwrap()), "Long(0)");
    assert_eq!(format!("{:?}", constant("vbTextCompare").unwrap()), "Long(1)");
    assert_eq!(format!("{:?}", constant("vbSunday").unwrap()), "Long(1)");
    assert_eq!(format!("{:?}", constant("vbSaturday").unwrap()), "Long(7)");
    assert_eq!(format!("{:?}", constant("vbString").unwrap()), "Long(8)");
    assert_eq!(format!("{:?}", constant("vbArray").unwrap()), "Long(8192)");
    assert_eq!(format!("{:?}", constant("vbRed").unwrap()), "Long(255)");
    assert_eq!(format!("{:?}", constant("vbObjectError").unwrap()), "Long(-2147221504)");
    assert_eq!(format!("{:?}", constant("vbUseDefault").unwrap()), "Long(-2)");
    assert_eq!(format!("{:?}", constant("vbCritical").unwrap()), "Long(16)");
    assert!(constant("vbNotAThing").is_none());
    assert!(constant("xlUp").is_none());
}

/// A minimal object with a default value, to check that library functions unwrap it.
struct Cell1(Value);
impl crate::object::Object for Cell1 {
    fn type_name(&self) -> String {
        "Range".into()
    }
    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        if name.is_empty() { Ok(self.0.clone()) } else { Err(RuntimeError::no_member(name)) }
    }
}

#[test]
fn objects_use_their_default_value() {
    let cell = Value::Object(Some(Rc::new(Cell1(s("Hello")))));
    assert_eq!(ev("Len", std::slice::from_ref(&cell)), "Long(5)");
    assert_eq!(txt("UCase", std::slice::from_ref(&cell)), "HELLO");
    assert_eq!(txt("TypeName", std::slice::from_ref(&cell)), "Range");
    assert_eq!(ev("IsObject", std::slice::from_ref(&cell)), "Boolean(true)");
    assert_eq!(ev("VarType", std::slice::from_ref(&cell)), "Long(8)");
    let num = Value::Object(Some(Rc::new(Cell1(d(2.5)))));
    assert_eq!(ev("CInt", std::slice::from_ref(&num)), "Integer(2)");
    let r = crate::ops::binary(crate::ops::BinOp::Add, &num, &d(1.0), false).unwrap();
    assert_eq!(format!("{r:?}"), "Double(3.5)");
}
