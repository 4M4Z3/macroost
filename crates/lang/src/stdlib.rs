//! The standard library ([MS-VBAL] section 6): conversion, string, math, date/time,
//! information, array and interaction functions, and the predefined constants.
//!
//! Functions take runtime values. A skipped optional argument may arrive as
//! [`crate::object::MISSING`] or simply be absent. `Name$` forms are accepted for every function
//! and differ only in raising error 94 where the plain form would return Null.
//!
//! Behaviour not settled by the specification is marked `ORACLE:`.

pub mod datetime;
pub mod format;

use std::cell::{Cell, RefCell};
use std::rc::Rc;

use crate::decimal::Decimal;
use crate::error::{Result, RuntimeError, standard_description};
use crate::object::Host;
use crate::ops::{compare_text, currency_of, fold_unit};
use crate::value::{Array, Value, parse_number, round_half_even, ymd_to_serial};

use format::NumberStyle;

/// Calls a standard library function by name (case-insensitive). `None` = not a library function.
pub fn call(name: &str, args: &[Value], host: &mut dyn Host) -> Option<Result<Value>> {
    let lower = name.to_ascii_lowercase();
    match lower.as_str() {
        "date$" => return Some(date_string(host)),
        "time$" => return Some(time_string(host)),
        _ => {}
    }
    let (base, dollar) = match lower.strip_suffix('$') {
        Some(b) => (b, true),
        None => (lower.as_str(), false),
    };
    let result = dispatch(base, args, host)?;
    Some(match result {
        Ok(Value::Null) if dollar => Err(RuntimeError::standard(94)),
        other => other,
    })
}

fn dispatch(name: &str, args: &[Value], host: &mut dyn Host) -> Option<Result<Value>> {
    Some(match name {
        // Conversion
        "cstr" => conv(args, |v| Ok(Value::str(v.to_text()?))),
        "cint" => conv(args, |v| Ok(Value::Integer(v.to_i16()?))),
        "clng" => conv(args, |v| Ok(Value::Long(v.to_i32()?))),
        "clnglng" => conv(args, |v| Ok(Value::LongLong(v.to_i64()?))),
        // ORACLE: LongPtr is LongLong on 64-bit hosts, which we model.
        "clngptr" => conv(args, |v| Ok(Value::LongLong(v.to_i64()?))),
        "cbyte" => conv(args, |v| u8::try_from(v.to_i64()?).map(Value::Byte).map_err(|_| RuntimeError::overflow())),
        "cdbl" => conv(args, |v| Ok(Value::Double(v.to_f64()?))),
        "cdec" => conv(args, |v| Ok(Value::Decimal(Decimal::from_value(v)?))),
        "csng" => conv(args, |v| {
            let x = v.to_f64()?;
            if x.is_finite() && x.abs() > f64::from(f32::MAX) {
                return Err(RuntimeError::overflow());
            }
            Ok(Value::Single(x as f32))
        }),
        "ccur" => conv(args, |v| Ok(Value::Currency(currency_of(v)?))),
        "cdate" | "cvdate" => {
            let year = current_year(host);
            conv(args, move |v| Ok(Value::Date(datetime::to_date(v, year)?)))
        }
        "cbool" => conv(args, |v| Ok(Value::Boolean(v.to_bool()?))),
        "cvar" => one(args).and_then(|v| v.deref_default()),
        "cverr" => conv(args, |v| Ok(Value::Error(v.to_i32()?))),
        "val" => f_val(args),
        "str" => f_str(args),
        "hex" => f_radix(args, 16),
        "oct" => f_radix(args, 8),
        "fix" => f_int(args, true),
        "int" => f_int(args, false),
        "error" => f_error(args),

        // Strings
        "len" => f_len(args, 1),
        "lenb" => f_len(args, 2),
        "left" => f_left_right(args, true),
        "right" => f_left_right(args, false),
        "mid" => f_mid(args),
        "instr" => f_instr(args),
        "instrrev" => f_instrrev(args),
        "replace" => f_replace(args),
        "trim" => map_str(args, |s| s.trim_matches(' ').to_owned()),
        "ltrim" => map_str(args, |s| s.trim_start_matches(' ').to_owned()),
        "rtrim" => map_str(args, |s| s.trim_end_matches(' ').to_owned()),
        "ucase" => map_str(args, |s| s.to_uppercase()),
        "lcase" => map_str(args, |s| s.to_lowercase()),
        "strreverse" => f_strreverse(args),
        "strcomp" => f_strcomp(args),
        "strconv" => f_strconv(args),
        "space" => f_space(args),
        "string" => f_string(args),
        "chr" => f_chr(args, false),
        "chrw" => f_chr(args, true),
        "asc" => f_asc(args, false),
        "ascw" => f_asc(args, true),
        "split" => f_split(args),
        "join" => f_join(args),
        "filter" => f_filter(args),
        "format" => f_format(args, host),
        "formatnumber" => f_format_number(args, Style::Number),
        "formatcurrency" => f_format_number(args, Style::Currency),
        "formatpercent" => f_format_number(args, Style::Percent),
        "formatdatetime" => f_format_date_time(args, host),

        // Math
        "abs" => f_abs(args),
        "sgn" => f_sgn(args),
        "sqr" => math(args, |x| if x < 0.0 { Err(RuntimeError::invalid_call()) } else { Ok(x.sqrt()) }),
        "exp" => math(args, |x| {
            let r = x.exp();
            if r.is_finite() { Ok(r) } else { Err(RuntimeError::overflow()) }
        }),
        "log" => math(args, |x| if x <= 0.0 { Err(RuntimeError::invalid_call()) } else { Ok(x.ln()) }),
        "sin" => math(args, |x| Ok(x.sin())),
        "cos" => math(args, |x| Ok(x.cos())),
        "tan" => math(args, |x| Ok(x.tan())),
        "atn" => math(args, |x| Ok(x.atan())),
        "round" => f_round(args),
        "rnd" => f_rnd(args),
        "randomize" => f_randomize(args, host),

        // Date and time
        "now" => no_args(args).map(|()| Value::Date(host.now())),
        "date" => no_args(args).map(|()| Value::Date(host.now().trunc())),
        "time" => no_args(args).map(|()| Value::Date(datetime::split(host.now()).1)),
        "timer" => no_args(args).map(|()| Value::Single((datetime::split(host.now()).1 * 86_400.0) as f32)),
        "year" => date_part_fn(args, host, |d| datetime::ymd(d).0),
        "month" => date_part_fn(args, host, |d| i64::from(datetime::ymd(d).1)),
        "day" => date_part_fn(args, host, |d| i64::from(datetime::ymd(d).2)),
        "hour" => date_part_fn(args, host, |d| i64::from(datetime::hms(d).0)),
        "minute" => date_part_fn(args, host, |d| i64::from(datetime::hms(d).1)),
        "second" => date_part_fn(args, host, |d| i64::from(datetime::hms(d).2)),
        "weekday" => f_weekday(args, host),
        "weekdayname" => f_weekday_name(args),
        "monthname" => f_month_name(args),
        "dateserial" => f_date_serial(args),
        "timeserial" => f_time_serial(args),
        "datevalue" => f_date_value(args, host, true),
        "timevalue" => f_date_value(args, host, false),
        "dateadd" => f_date_add(args, host),
        "datediff" => f_date_diff(args, host),
        "datepart" => f_date_part(args, host),
        "isdate" => f_isdate(args, host),

        // Information
        "isnumeric" => one(args).map(|v| Value::Boolean(is_numeric(&v))),
        "isempty" => one(args).map(|v| Value::Boolean(v.is_empty())),
        "isnull" => one(args).map(|v| Value::Boolean(v.is_null())),
        "isarray" => one(args).map(|v| Value::Boolean(matches!(v, Value::Array(_)))),
        "isobject" => one_raw(args).map(|v| Value::Boolean(v.is_object())),
        "iserror" => one(args).map(|v| Value::Boolean(matches!(v, Value::Error(_)))),
        "ismissing" => f_ismissing(args),
        "typename" => one_raw(args).map(|v| Value::str(type_name(&v))),
        "vartype" => f_vartype(args),
        "rgb" => f_rgb(args),
        "qbcolor" => f_qbcolor(args),

        // Arrays
        "array" => Ok(array_value(Array::from_vec(0, args.to_vec()))),
        "ubound" => f_bound(args, false),
        "lbound" => f_bound(args, true),

        // Interaction
        "msgbox" => f_msgbox(args, host),
        "inputbox" => f_inputbox(args, host),
        // Security: macros get no view of the process environment.
        "environ" => max_args(args, 1).map(|()| Value::str("")),
        "iif" => f_iif(args),
        "choose" => f_choose(args),
        "switch" => f_switch(args),
        "doevents" => Ok(Value::Integer(0)),
        _ => return None,
    })
}

/// Standard constants (vbCrLf, vbTab, vbNullString, vbYes, vbOKOnly, ...). `None` = unknown.
pub fn constant(name: &str) -> Option<Value> {
    let s = |t: &str| Some(Value::str(t));
    let n = |x: i32| Some(Value::Long(x));
    match name.to_ascii_lowercase().as_str() {
        // Constants module (6.1.2.2)
        "vbback" => s("\u{8}"),
        "vbcr" => s("\r"),
        "vbcrlf" => s("\r\n"),
        "vbformfeed" => s("\u{c}"),
        "vblf" => s("\n"),
        "vbnewline" => s("\r\n"),
        "vbnullchar" => s("\0"),
        "vbtab" => s("\t"),
        "vbverticaltab" => s("\u{b}"),
        "vbnullstring" => s(""),
        "vbobjecterror" => n(-2_147_221_504),
        // ColorConstants
        "vbblack" => n(0),
        "vbblue" => n(16_711_680),
        "vbcyan" => n(16_776_960),
        "vbgreen" => n(65_280),
        "vbmagenta" => n(16_711_935),
        "vbred" => n(255),
        "vbwhite" => n(16_777_215),
        "vbyellow" => n(65_535),
        // FormShowConstants
        "vbmodal" => n(1),
        "vbmodeless" => n(0),
        // VbAppWinStyle
        "vbhide" => n(0),
        "vbnormalfocus" => n(1),
        "vbminimizedfocus" => n(2),
        "vbmaximizedfocus" => n(3),
        "vbnormalnofocus" => n(4),
        "vbminimizednofocus" => n(6),
        // VbCalendar
        "vbcalgreg" => n(0),
        "vbcalhijri" => n(1),
        // VbCallType
        "vbmethod" => n(1),
        "vbget" => n(2),
        "vblet" => n(4),
        "vbset" => n(8),
        // VbCompareMethod
        "vbbinarycompare" => n(0),
        "vbtextcompare" => n(1),
        "vbdatabasecompare" => n(2),
        "vbusecompareoption" => n(-1),
        // VbDateTimeFormat
        "vbgeneraldate" => n(0),
        "vblongdate" => n(1),
        "vbshortdate" => n(2),
        "vblongtime" => n(3),
        "vbshorttime" => n(4),
        // VbDayOfWeek
        "vbusesystemdayofweek" => n(0),
        "vbsunday" => n(1),
        "vbmonday" => n(2),
        "vbtuesday" => n(3),
        "vbwednesday" => n(4),
        "vbthursday" => n(5),
        "vbfriday" => n(6),
        "vbsaturday" => n(7),
        // VbFileAttribute
        "vbnormal" => n(0),
        "vbreadonly" => n(1),
        "vbhidden" => n(2),
        "vbsystem" => n(4),
        "vbvolume" => n(8),
        "vbdirectory" => n(16),
        "vbarchive" => n(32),
        "vbalias" => n(64),
        // VbFirstWeekOfYear
        "vbusesystem" => n(0),
        "vbfirstjan1" => n(1),
        "vbfirstfourdays" => n(2),
        "vbfirstfullweek" => n(3),
        // VbMsgBoxResult
        "vbok" => n(1),
        "vbcancel" => n(2),
        "vbabort" => n(3),
        "vbretry" => n(4),
        "vbignore" => n(5),
        "vbyes" => n(6),
        "vbno" => n(7),
        // VbMsgBoxStyle
        "vbokonly" => n(0),
        "vbokcancel" => n(1),
        "vbabortretryignore" => n(2),
        "vbyesnocancel" => n(3),
        "vbyesno" => n(4),
        "vbretrycancel" => n(5),
        "vbcritical" => n(16),
        "vbquestion" => n(32),
        "vbexclamation" => n(48),
        "vbinformation" => n(64),
        "vbdefaultbutton1" => n(0),
        "vbdefaultbutton2" => n(256),
        "vbdefaultbutton3" => n(512),
        "vbdefaultbutton4" => n(768),
        "vbapplicationmodal" => n(0),
        "vbsystemmodal" => n(4096),
        "vbmsgboxhelpbutton" => n(16_384),
        "vbmsgboxsetforeground" => n(65_536),
        "vbmsgboxright" => n(524_288),
        "vbmsgboxrtlreading" => n(1_048_576),
        // VbQueryClose
        "vbformcontrolmenu" => n(0),
        "vbformcode" => n(1),
        "vbappwindows" => n(2),
        "vbapptaskmanager" => n(3),
        "vbformmdiform" => n(4),
        // VbStrConv
        "vbuppercase" => n(1),
        "vblowercase" => n(2),
        "vbpropercase" => n(3),
        "vbwide" => n(4),
        "vbnarrow" => n(8),
        "vbkatakana" => n(16),
        "vbhiragana" => n(32),
        "vbunicode" => n(64),
        "vbfromunicode" => n(128),
        // VbTriState
        "vbtrue" => n(-1),
        "vbfalse" => n(0),
        "vbusedefault" => n(-2),
        // VbVarType
        "vbempty" => n(0),
        "vbnull" => n(1),
        "vbinteger" => n(2),
        "vblong" => n(3),
        "vbsingle" => n(4),
        "vbdouble" => n(5),
        "vbcurrency" => n(6),
        "vbdate" => n(7),
        "vbstring" => n(8),
        "vbobject" => n(9),
        "vberror" => n(10),
        "vbboolean" => n(11),
        "vbvariant" => n(12),
        "vbdataobject" => n(13),
        "vbdecimal" => n(14),
        "vbbyte" => n(17),
        "vblonglong" => n(20),
        "vbuserdefinedtype" => n(36),
        "vbarray" => n(8192),
        // VbIMEStatus
        "vbimenoop" | "vbimemodenocontrol" => n(0),
        "vbimeon" | "vbimemodeon" => n(1),
        "vbimeoff" | "vbimemodeoff" => n(2),
        "vbimedisable" | "vbimemodedisable" => n(3),
        "vbimehiragana" | "vbimemodehiragana" => n(4),
        "vbimekatakanadbl" | "vbimemodekatakana" => n(5),
        "vbimekatakanasng" | "vbimemodekatakanahalf" => n(6),
        "vbimealphadbl" | "vbimemodealphafull" => n(7),
        "vbimealphasng" | "vbimemodealpha" => n(8),
        "vbimemodehangulfull" => n(9),
        "vbimemodehangul" => n(10),
        _ => None,
    }
}

// ---------------------------------------------------------------------------------------------
// Argument helpers

/// The argument at `i` as passed (objects not unwrapped); `None` if absent or skipped.
fn raw(args: &[Value], i: usize) -> Option<&Value> {
    args.get(i).filter(|v| !matches!(v, Value::Error(448)))
}

/// The argument at `i` as data (objects unwrapped to their default value).
fn opt(args: &[Value], i: usize) -> Result<Option<Value>> {
    raw(args, i).map(|v| v.clone().deref_default()).transpose()
}

fn req(args: &[Value], i: usize) -> Result<Value> {
    opt(args, i)?.ok_or_else(|| RuntimeError::standard(449))
}

fn max_args(args: &[Value], n: usize) -> Result<()> {
    if args.len() > n { Err(RuntimeError::standard(450)) } else { Ok(()) }
}

fn no_args(args: &[Value]) -> Result<()> {
    max_args(args, 0)
}

fn one(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    req(args, 0)
}

fn one_raw(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    raw(args, 0).cloned().ok_or_else(|| RuntimeError::standard(449))
}

fn not_null(v: Value) -> Result<Value> {
    if v.is_null() { Err(RuntimeError::standard(94)) } else { Ok(v) }
}

/// An optional integer argument; Null is "Invalid use of Null".
fn opt_i64(args: &[Value], i: usize, default: i64) -> Result<i64> {
    match opt(args, i)? {
        None => Ok(default),
        Some(v) => not_null(v)?.to_i64(),
    }
}

fn req_i64(args: &[Value], i: usize) -> Result<i64> {
    not_null(req(args, i)?)?.to_i64()
}

fn req_text(args: &[Value], i: usize) -> Result<String> {
    not_null(req(args, i)?)?.to_text()
}

fn opt_text(args: &[Value], i: usize, default: &str) -> Result<String> {
    match opt(args, i)? {
        None => Ok(default.to_owned()),
        Some(v) => not_null(v)?.to_text(),
    }
}

/// A `Compare` argument: true for text comparison. ORACLE: an omitted argument should follow
/// the module's Option Compare, which the library cannot see; we use binary.
fn compare_mode(args: &[Value], i: usize) -> Result<bool> {
    match opt_i64(args, i, 0)? {
        -1 | 0 => Ok(false),
        n if n > 0 => Ok(true),
        _ => Err(RuntimeError::invalid_call()),
    }
}

/// A string argument where Null passes through as `None`.
fn text_or_null(args: &[Value], i: usize) -> Result<Option<String>> {
    let v = req(args, i)?;
    if v.is_null() { Ok(None) } else { v.to_text().map(Some) }
}

fn current_year(host: &mut dyn Host) -> i64 {
    datetime::ymd(host.now()).0
}

fn array_value(a: Array) -> Value {
    Value::Array(Rc::new(RefCell::new(a)))
}

fn string_array(items: Vec<String>) -> Value {
    array_value(Array::from_vec(0, items.into_iter().map(Value::from).collect()))
}

/// Largest string the library builds in one call (`Space`, `String`), in UTF-16 units.
const MAX_STRING: i64 = 256 * 1024 * 1024;

fn u16s(s: &str) -> Vec<u16> {
    s.encode_utf16().collect()
}

fn from_u16(v: &[u16]) -> String {
    String::from_utf16_lossy(v)
}

fn units_eq(a: &[u16], b: &[u16], text: bool) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(&x, &y)| x == y || (text && fold_unit(x) == fold_unit(y)))
}

/// First index >= `from` where `needle` occurs.
fn find_units(hay: &[u16], needle: &[u16], from: usize, text: bool) -> Option<usize> {
    if needle.len() > hay.len() {
        return None;
    }
    (from..=hay.len() - needle.len()).find(|&i| units_eq(&hay[i..i + needle.len()], needle, text))
}

// ---------------------------------------------------------------------------------------------
// Conversion

/// A one-argument conversion: Null raises error 94 ("Invalid use of Null").
fn conv(args: &[Value], f: impl Fn(&Value) -> Result<Value>) -> Result<Value> {
    let v = not_null(one(args)?)?;
    f(&v)
}

fn f_val(args: &[Value]) -> Result<Value> {
    let s = not_null(one(args)?)?.to_text()?;
    Ok(Value::Double(val(&s)))
}

/// `Val`: the leading number in a string, ignoring blanks, tabs and line feeds.
fn val(s: &str) -> f64 {
    let t: Vec<char> = s.chars().filter(|c| !matches!(c, ' ' | '\t' | '\n' | '\r')).collect();
    let lower: String = t.iter().take(2).collect::<String>().to_ascii_lowercase();
    if lower == "&h" || lower == "&o" {
        let radix = if lower == "&h" { 16 } else { 8 };
        let digits: String = t[2..].iter().take_while(|c| c.is_digit(radix)).collect();
        let Ok(n) = u64::from_str_radix(&digits, radix) else {
            return 0.0;
        };
        // ORACLE: hex/octal literals in text wrap like typed literals: up to 16 bits as Integer,
        // up to 32 bits as Long.
        return if n <= 0xFFFF {
            f64::from(n as u16 as i16)
        } else if n <= 0xFFFF_FFFF {
            f64::from(n as u32 as i32)
        } else {
            n as i64 as f64
        };
    }
    let mut out = String::new();
    let mut i = 0;
    if matches!(t.first(), Some('+' | '-')) {
        out.push(t[0]);
        i = 1;
    }
    let mut seen_point = false;
    let mut seen_digit = false;
    while i < t.len() {
        let c = t[i];
        if c.is_ascii_digit() {
            seen_digit = true;
            out.push(c);
        } else if c == '.' && !seen_point {
            seen_point = true;
            out.push(c);
        } else if matches!(c, 'e' | 'E' | 'd' | 'D') && seen_digit {
            let mut exp = String::from("e");
            let mut j = i + 1;
            if matches!(t.get(j), Some('+' | '-')) {
                exp.push(t[j]);
                j += 1;
            }
            let digits: String = t[j..].iter().take_while(|c| c.is_ascii_digit()).collect();
            if !digits.is_empty() {
                out.push_str(&exp);
                out.push_str(&digits);
            }
            break;
        } else {
            break;
        }
        i += 1;
    }
    if !seen_digit {
        return 0.0;
    }
    out.parse().unwrap_or(0.0)
}

fn f_str(args: &[Value]) -> Result<Value> {
    let v = one(args)?;
    Ok(match v {
        Value::Null => Value::Null,
        Value::Error(n) => Value::str(format!("Error {n}")),
        Value::Date(_) => Value::str(v.to_text()?),
        Value::String(ref s) => {
            let x = parse_number(s).ok_or_else(RuntimeError::type_mismatch)?;
            Value::str(with_sign_space(x, Value::Double(x).to_text()?))
        }
        other => {
            // ORACLE: Boolean follows the spec literally: True (-1) is "True", False " False".
            let x = other.to_f64()?;
            Value::str(with_sign_space(x, other.to_text()?))
        }
    })
}

fn with_sign_space(x: f64, s: String) -> String {
    if x >= 0.0 { format!(" {s}") } else { s }
}

fn f_radix(args: &[Value], radix: u32) -> Result<Value> {
    let v = one(args)?;
    let digits = |n: u64| if radix == 16 { format!("{n:X}") } else { format!("{n:o}") };
    Ok(Value::str(match v {
        Value::Null => return Ok(Value::Null),
        Value::Empty => "0".to_owned(),
        Value::LongLong(n) => digits(n as u64),
        other => {
            let n = other.to_i32()?;
            if n >= 0 {
                digits(n as u64)
            } else if n >= -32_767 {
                digits(u64::from(n as i16 as u16))
            } else {
                digits(u64::from(n as u32))
            }
        }
    }))
}

/// `Fix` (toward zero) and `Int` (toward negative infinity), keeping the value type.
fn f_int(args: &[Value], fix: bool) -> Result<Value> {
    let v = one(args)?;
    let r = |x: f64| if fix { x.trunc() } else { x.floor() };
    Ok(match v {
        Value::Null => Value::Null,
        Value::Integer(_) | Value::Long(_) | Value::LongLong(_) | Value::Byte(_) => v,
        Value::Boolean(b) => Value::Integer(if b { -1 } else { 0 }),
        Value::Empty => Value::Integer(0),
        Value::Single(x) => Value::Single(r(f64::from(x)) as f32),
        Value::Double(x) => Value::Double(r(x)),
        Value::Currency(n) => {
            let q = if fix { n / 10_000 } else { n.div_euclid(10_000) };
            Value::Currency(q * 10_000)
        }
        Value::Decimal(d) => Value::Decimal(d.trunc(!fix)),
        Value::Date(d) => Value::Date(r(d)),
        Value::String(_) => Value::Double(r(v.to_f64()?)),
        other => Value::Integer(other.to_i16()?),
    })
}

fn f_error(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    let n = opt_i64(args, 0, 0)?;
    if n == 0 {
        return Ok(Value::str(""));
    }
    let n = i32::try_from(n).map_err(|_| RuntimeError::overflow())?;
    if !(1..=65_535).contains(&n) {
        return Err(RuntimeError::invalid_call());
    }
    Ok(Value::str(standard_description(n)))
}

// ---------------------------------------------------------------------------------------------
// Strings

fn map_str(args: &[Value], f: impl Fn(&str) -> String) -> Result<Value> {
    max_args(args, 1)?;
    Ok(match text_or_null(args, 0)? {
        None => Value::Null,
        Some(s) => Value::str(f(&s)),
    })
}

fn f_len(args: &[Value], bytes: usize) -> Result<Value> {
    max_args(args, 1)?;
    let v = req(args, 0)?;
    if v.is_null() {
        return Ok(Value::Null);
    }
    let n = v.to_text()?.encode_utf16().count() * bytes;
    Ok(Value::Long(i32::try_from(n).map_err(|_| RuntimeError::overflow())?))
}

fn f_left_right(args: &[Value], left: bool) -> Result<Value> {
    max_args(args, 2)?;
    let s = text_or_null(args, 0)?;
    let n = req_i64(args, 1)?;
    if n < 0 {
        return Err(RuntimeError::invalid_call());
    }
    let Some(s) = s else { return Ok(Value::Null) };
    let u = u16s(&s);
    let n = (n as usize).min(u.len());
    Ok(Value::str(from_u16(if left { &u[..n] } else { &u[u.len() - n..] })))
}

fn f_mid(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let s = text_or_null(args, 0)?;
    let start = req_i64(args, 1)?;
    let len = match opt(args, 2)? {
        None => None,
        Some(v) => Some(not_null(v)?.to_i64()?),
    };
    if start < 1 || len.is_some_and(|l| l < 0) {
        return Err(RuntimeError::invalid_call());
    }
    let Some(s) = s else { return Ok(Value::Null) };
    let u = u16s(&s);
    let from = ((start - 1) as usize).min(u.len());
    let to = match len {
        None => u.len(),
        Some(l) => from.saturating_add(l as usize).min(u.len()),
    };
    Ok(Value::str(from_u16(&u[from..to])))
}

fn f_instr(args: &[Value]) -> Result<Value> {
    max_args(args, 4)?;
    let (start, first) = if args.len() >= 3 {
        let start = req(args, 0)?;
        (not_null(start)?.to_i64()?, 1)
    } else {
        (1, 0)
    };
    if start < 1 {
        return Err(RuntimeError::invalid_call());
    }
    let hay = text_or_null(args, first)?;
    let needle = text_or_null(args, first + 1)?;
    let text = if args.len() >= 4 { compare_mode(args, 3)? } else { false };
    let (Some(hay), Some(needle)) = (hay, needle) else { return Ok(Value::Null) };
    let (h, n) = (u16s(&hay), u16s(&needle));
    let start = start as usize;
    if h.is_empty() || start > h.len() {
        // ORACLE: InStr(start, s, "") with start past the end - we return 0.
        return Ok(Value::Long(0));
    }
    if n.is_empty() {
        return Ok(Value::Long(start as i32));
    }
    Ok(Value::Long(find_units(&h, &n, start - 1, text).map_or(0, |i| i as i32 + 1)))
}

fn f_instrrev(args: &[Value]) -> Result<Value> {
    max_args(args, 4)?;
    let hay = text_or_null(args, 0)?;
    let needle = text_or_null(args, 1)?;
    let start = opt_i64(args, 2, -1)?;
    let text = compare_mode(args, 3)?;
    if start == 0 || start < -1 {
        return Err(RuntimeError::invalid_call());
    }
    let (Some(hay), Some(needle)) = (hay, needle) else { return Ok(Value::Null) };
    let (h, n) = (u16s(&hay), u16s(&needle));
    let start = if start == -1 { h.len() } else { start as usize };
    if h.is_empty() || start > h.len() {
        return Ok(Value::Long(0));
    }
    if n.is_empty() {
        return Ok(Value::Long(start as i32));
    }
    if n.len() > start {
        return Ok(Value::Long(0));
    }
    let found = (0..=start - n.len()).rev().find(|&i| units_eq(&h[i..i + n.len()], &n, text));
    Ok(Value::Long(found.map_or(0, |i| i as i32 + 1)))
}

fn f_replace(args: &[Value]) -> Result<Value> {
    max_args(args, 6)?;
    // ORACLE: Replace(Null, ...) - we return Null.
    let expr = text_or_null(args, 0)?;
    let find = req_text(args, 1)?;
    let repl = req_text(args, 2)?;
    let start = opt_i64(args, 3, 1)?;
    let count = opt_i64(args, 4, -1)?;
    let text = compare_mode(args, 5)?;
    if start < 1 || count < -1 {
        return Err(RuntimeError::invalid_call());
    }
    let Some(expr) = expr else { return Ok(Value::Null) };
    let e = u16s(&expr);
    let start = start as usize;
    if start > e.len() {
        return Ok(Value::str(""));
    }
    // The result starts at `start`: the text before it is dropped.
    let e = &e[start - 1..];
    let f = u16s(&find);
    if f.is_empty() || count == 0 {
        return Ok(Value::str(from_u16(e)));
    }
    let r = u16s(&repl);
    let mut out: Vec<u16> = Vec::with_capacity(e.len());
    let mut i = 0;
    let mut done = 0i64;
    while i < e.len() {
        if (count == -1 || done < count) && e.len() - i >= f.len() && units_eq(&e[i..i + f.len()], &f, text) {
            out.extend_from_slice(&r);
            i += f.len();
            done += 1;
        } else {
            out.push(e[i]);
            i += 1;
        }
    }
    Ok(Value::str(from_u16(&out)))
}

fn f_strreverse(args: &[Value]) -> Result<Value> {
    let s = not_null(one(args)?)?.to_text()?;
    let mut u = u16s(&s);
    u.reverse();
    Ok(Value::str(from_u16(&u)))
}

fn f_strcomp(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let a = text_or_null(args, 0)?;
    let b = text_or_null(args, 1)?;
    let text = compare_mode(args, 2)?;
    let (Some(a), Some(b)) = (a, b) else { return Ok(Value::Null) };
    Ok(Value::Integer(match compare_text(&a, &b, text) {
        std::cmp::Ordering::Less => -1,
        std::cmp::Ordering::Equal => 0,
        std::cmp::Ordering::Greater => 1,
    }))
}

fn f_strconv(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let s = text_or_null(args, 0)?;
    let conv = req_i64(args, 1)?;
    let Some(s) = s else { return Ok(Value::Null) };
    Ok(Value::str(match conv {
        1 => s.to_uppercase(),
        2 => s.to_lowercase(),
        3 => {
            let mut out = String::with_capacity(s.len());
            let mut at_word_start = true;
            for c in s.chars() {
                if at_word_start {
                    out.extend(c.to_uppercase());
                } else {
                    out.extend(c.to_lowercase());
                }
                at_word_start = !c.is_alphanumeric() && c != '\'';
            }
            out
        }
        // ORACLE: wide/narrow/kana and byte-array conversions are not modelled.
        _ => return Err(RuntimeError::invalid_call()),
    }))
}

fn f_space(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    let n = req_i64(args, 0)?;
    if n < 0 {
        return Err(RuntimeError::invalid_call());
    }
    if n > MAX_STRING {
        return Err(RuntimeError::standard(14));
    }
    Ok(Value::str(" ".repeat(n as usize)))
}

fn f_string(args: &[Value]) -> Result<Value> {
    max_args(args, 2)?;
    let n = req_i64(args, 0)?;
    let ch = req(args, 1)?;
    if n < 0 {
        return Err(RuntimeError::invalid_call());
    }
    if n > MAX_STRING {
        return Err(RuntimeError::standard(14));
    }
    let unit: u16 = match ch {
        Value::Null => return Ok(Value::Null),
        Value::String(ref s) => *u16s(s).first().ok_or_else(RuntimeError::invalid_call)?,
        other => {
            // A character code above 255 is taken modulo 256.
            let code = other.to_i64()?.rem_euclid(256);
            u16s(&chr_ansi(code as u8)).first().copied().unwrap_or(0)
        }
    };
    Ok(Value::str(from_u16(&vec![unit; n as usize])))
}

/// Windows-1252 characters 0x80..=0x9F (the rest of 0..=255 maps to the same code point).
const CP1252_HIGH: [u16; 32] = [
    0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D,
    0x017D, 0x008F, 0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A,
    0x0153, 0x009D, 0x017E, 0x0178,
];

/// `Chr` for codes 0..=255 in the ANSI code page. ORACLE: we assume Windows-1252.
fn chr_ansi(code: u8) -> String {
    let unit = if (0x80..=0x9F).contains(&code) { CP1252_HIGH[usize::from(code - 0x80)] } else { u16::from(code) };
    from_u16(&[unit])
}

fn f_chr(args: &[Value], wide: bool) -> Result<Value> {
    let n = not_null(one(args)?)?.to_i64()?;
    if wide {
        if !(-32_768..=65_535).contains(&n) {
            return Err(RuntimeError::invalid_call());
        }
        // ORACLE: a lone surrogate cannot be held in a Rust string; it becomes U+FFFD.
        return Ok(Value::str(from_u16(&[n.rem_euclid(65_536) as u16])));
    }
    match u8::try_from(n) {
        Ok(code) => Ok(Value::str(chr_ansi(code))),
        Err(_) => Err(RuntimeError::invalid_call()),
    }
}

fn f_asc(args: &[Value], wide: bool) -> Result<Value> {
    let s = not_null(one(args)?)?.to_text()?;
    let unit = *u16s(&s).first().ok_or_else(RuntimeError::invalid_call)?;
    if wide {
        return Ok(Value::Integer(unit as i16));
    }
    let code = if unit < 0x80 || (0xA0..=0xFF).contains(&unit) {
        unit
    } else {
        // Characters outside the code page read back as "?".
        CP1252_HIGH.iter().position(|&c| c == unit).map_or(63, |p| 0x80 + p as u16)
    };
    Ok(Value::Integer(code as i16))
}

fn f_split(args: &[Value]) -> Result<Value> {
    max_args(args, 4)?;
    // ORACLE: Split(Null) - we raise "Invalid use of Null".
    let expr = req_text(args, 0)?;
    let delim = opt_text(args, 1, " ")?;
    let limit = opt_i64(args, 2, -1)?;
    let text = compare_mode(args, 3)?;
    if limit < -1 {
        return Err(RuntimeError::invalid_call());
    }
    if expr.is_empty() || limit == 0 {
        return Ok(array_value(Array::new(vec![(0, -1)])?));
    }
    let e = u16s(&expr);
    let d = u16s(&delim);
    if d.is_empty() {
        return Ok(string_array(vec![expr]));
    }
    let mut parts = Vec::new();
    let mut from = 0;
    while limit == -1 || (parts.len() as i64) < limit - 1 {
        match find_units(&e, &d, from, text) {
            Some(i) => {
                parts.push(from_u16(&e[from..i]));
                from = i + d.len();
            }
            None => break,
        }
    }
    parts.push(from_u16(&e[from..]));
    Ok(string_array(parts))
}

/// The elements of a one-dimensional array argument.
fn one_dim(v: &Value) -> Result<Vec<Value>> {
    match v {
        Value::Array(a) => {
            let a = a.borrow();
            if a.bounds.len() > 1 {
                return Err(RuntimeError::invalid_call());
            }
            Ok(a.data.clone())
        }
        _ => Err(RuntimeError::type_mismatch()),
    }
}

fn f_join(args: &[Value]) -> Result<Value> {
    max_args(args, 2)?;
    let items = one_dim(&req(args, 0)?)?;
    let delim = opt_text(args, 1, " ")?;
    let parts = items.iter().map(|v| v.clone().deref_default()?.to_text()).collect::<Result<Vec<_>>>()?;
    Ok(Value::str(parts.join(&delim)))
}

fn f_filter(args: &[Value]) -> Result<Value> {
    max_args(args, 4)?;
    let items = one_dim(&req(args, 0)?)?;
    let needle = u16s(&req_text(args, 1)?);
    let include = match opt(args, 2)? {
        None => true,
        Some(v) => not_null(v)?.to_bool()?,
    };
    let text = compare_mode(args, 3)?;
    let mut out = Vec::new();
    for v in items {
        let s = v.deref_default()?.to_text()?;
        let hit = find_units(&u16s(&s), &needle, 0, text).is_some() || needle.is_empty();
        if hit == include {
            out.push(s);
        }
    }
    if out.is_empty() {
        return Ok(array_value(Array::new(vec![(0, -1)])?));
    }
    Ok(string_array(out))
}

fn first_day_arg(args: &[Value], i: usize) -> Result<u32> {
    match opt_i64(args, i, 1)? {
        // ORACLE: vbUseSystemDayOfWeek is taken as Sunday (en-US).
        0 => Ok(1),
        n @ 1..=7 => Ok(n as u32),
        _ => Err(RuntimeError::invalid_call()),
    }
}

fn first_week_arg(args: &[Value], i: usize) -> Result<u32> {
    match opt_i64(args, i, 1)? {
        0 => Ok(1),
        n @ 1..=3 => Ok(n as u32),
        _ => Err(RuntimeError::invalid_call()),
    }
}

fn f_format(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 4)?;
    let v = req(args, 0)?;
    let fmt = match opt(args, 1)? {
        None | Some(Value::Null) => String::new(),
        Some(f) => f.to_text()?,
    };
    let first_day = first_day_arg(args, 2)?;
    let first_week = first_week_arg(args, 3)?;
    format::format(&v, &fmt, first_day, first_week, current_year(host))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Style {
    Number,
    Currency,
    Percent,
}

/// A VbTriState argument: -2 (vbUseDefault) or absent gives `default`.
fn tristate(args: &[Value], i: usize, default: bool) -> Result<bool> {
    Ok(match opt_i64(args, i, -2)? {
        -2 => default,
        0 => false,
        _ => true,
    })
}

fn f_format_number(args: &[Value], style: Style) -> Result<Value> {
    max_args(args, 5)?;
    // ORACLE: a Null expression - we raise "Invalid use of Null".
    let v = not_null(req(args, 0)?)?;
    let x = match &v {
        Value::String(s) => parse_number(s).ok_or_else(RuntimeError::type_mismatch)?,
        other => other.to_f64()?,
    };
    let decimals = match opt_i64(args, 1, -1)? {
        -1 => 2,
        n if n < -1 => return Err(RuntimeError::invalid_call()),
        n => n.min(99) as usize,
    };
    let st = NumberStyle {
        decimals,
        leading_digit: tristate(args, 2, true)?,
        parens: tristate(args, 3, style == Style::Currency)?,
        group: tristate(args, 4, true)?,
    };
    Ok(Value::str(match style {
        Style::Number => format::format_number_style(x, st, "", ""),
        Style::Currency => format::format_number_style(x, st, "$", ""),
        Style::Percent => format::format_number_style(x * 100.0, st, "", "%"),
    }))
}

fn f_format_date_time(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 2)?;
    let v = not_null(req(args, 0)?)?;
    let d = datetime::to_date(&v, current_year(host))?;
    let named = opt_i64(args, 1, 0)?;
    Ok(Value::str(format::format_date_time(d, i32::try_from(named).unwrap_or(-1))?))
}

// ---------------------------------------------------------------------------------------------
// Math

/// A Double function of one argument; Null passes through.
/// ORACLE: Null in -> Null out for the math functions.
fn math(args: &[Value], f: impl Fn(f64) -> Result<f64>) -> Result<Value> {
    let v = one(args)?;
    if v.is_null() {
        return Ok(Value::Null);
    }
    Ok(Value::Double(f(v.to_f64()?)?))
}

fn f_abs(args: &[Value]) -> Result<Value> {
    let v = one(args)?;
    let of = RuntimeError::overflow;
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(if b { 1 } else { 0 }),
        Value::Byte(_) => v,
        Value::Integer(n) => Value::Integer(n.checked_abs().ok_or_else(of)?),
        Value::Long(n) => Value::Long(n.checked_abs().ok_or_else(of)?),
        Value::LongLong(n) => Value::LongLong(n.checked_abs().ok_or_else(of)?),
        Value::Single(x) => Value::Single(x.abs()),
        Value::Double(x) => Value::Double(x.abs()),
        Value::Currency(n) => Value::Currency(n.checked_abs().ok_or_else(of)?),
        Value::Decimal(d) => Value::Decimal(d.abs()),
        // ORACLE: Abs of a Date - we return the Double.
        Value::Date(d) => Value::Double(d.abs()),
        other => Value::Double(other.to_f64()?.abs()),
    })
}

fn f_sgn(args: &[Value]) -> Result<Value> {
    let v = one(args)?;
    if v.is_null() {
        // ORACLE: Sgn(Null) - we return Null.
        return Ok(Value::Null);
    }
    if let Value::Decimal(d) = v {
        return Ok(Value::Integer(if d.is_negative() { -1 } else { i16::from(!d.is_zero()) }));
    }
    let x = v.to_f64()?;
    Ok(Value::Integer(if x > 0.0 {
        1
    } else if x < 0.0 {
        -1
    } else {
        0
    }))
}

fn f_round(args: &[Value]) -> Result<Value> {
    max_args(args, 2)?;
    let v = req(args, 0)?;
    let digits = opt_i64(args, 1, 0)?;
    if digits < 0 {
        return Err(RuntimeError::invalid_call());
    }
    let digits = digits.min(22) as usize;
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(if b { -1 } else { 0 }),
        Value::Byte(_) | Value::Integer(_) | Value::Long(_) | Value::LongLong(_) => v,
        Value::Single(x) => Value::Single(format::round_f64(f64::from(x), digits) as f32),
        Value::Currency(n) => {
            if digits >= 4 {
                v
            } else {
                let unit = 10i64.pow(4 - digits as u32);
                let q = round_half_even(n as f64 / unit as f64) as i64;
                Value::Currency(q * unit)
            }
        }
        Value::Decimal(d) => Value::Decimal(d.round(digits as u32)),
        // ORACLE: Round of a Date - we return a Double.
        other => Value::Double(format::round_f64(other.to_f64()?, digits)),
    })
}

thread_local! {
    /// (seed, last value) of the 24-bit linear congruential generator behind `Rnd`.
    static RNG: Cell<(u32, f32)> = const { Cell::new((0x50000, 0x50000 as f32 / 16_777_216.0)) };
}

fn rnd_step(seed: u32) -> u32 {
    seed.wrapping_mul(0x43FD_43FD).wrapping_add(0x00C3_9EC3) & 0x00FF_FFFF
}

/// `Rnd`. ORACLE: the generator constants and the seeding from a negative argument follow the
/// language's widely documented 24-bit LCG; confirm sequences against the reference spreadsheet.
fn f_rnd(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    let n = match opt(args, 0)? {
        None => None,
        Some(v) => Some(not_null(v)?.to_f64()?),
    };
    let (seed, last) = RNG.with(Cell::get);
    let next = match n {
        Some(0.0) => return Ok(Value::Single(last)),
        Some(x) if x < 0.0 => {
            let bits = (x as f32).to_bits();
            rnd_step(bits.wrapping_add(bits >> 24) & 0x00FF_FFFF)
        }
        _ => rnd_step(seed),
    };
    let value = next as f32 / 16_777_216.0;
    RNG.with(|r| r.set((next, value)));
    Ok(Value::Single(value))
}

/// `Randomize [seed]`: reseeds `Rnd` from the number (or the timer when omitted).
fn f_randomize(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 1)?;
    let x = match opt(args, 0)? {
        Some(v) => not_null(v)?.to_f64()?,
        None => f64::from((datetime::split(host.now()).1 * 86_400.0) as f32),
    };
    let hi = (x.to_bits() >> 32) as u32;
    let mix = ((hi & 0xFFFF) ^ (hi >> 16)) << 8;
    RNG.with(|r| {
        let (seed, last) = r.get();
        r.set(((seed & 0xFF) | (mix & 0x00FF_FF00), last));
    });
    Ok(Value::Empty)
}

// ---------------------------------------------------------------------------------------------
// Date and time

fn date_string(host: &mut dyn Host) -> Result<Value> {
    let (y, m, d) = datetime::ymd(host.now());
    Ok(Value::str(format!("{m:02}-{d:02}-{y:04}")))
}

fn time_string(host: &mut dyn Host) -> Result<Value> {
    let (h, m, s) = datetime::hms(host.now());
    Ok(Value::str(format!("{h:02}:{m:02}:{s:02}")))
}

/// The date argument at `i`, or `None` for Null.
fn date_arg(args: &[Value], i: usize, host: &mut dyn Host) -> Result<Option<f64>> {
    let v = req(args, i)?;
    if v.is_null() {
        return Ok(None);
    }
    datetime::to_date(&v, current_year(host)).map(Some)
}

fn date_part_fn(args: &[Value], host: &mut dyn Host, f: impl Fn(f64) -> i64) -> Result<Value> {
    max_args(args, 1)?;
    Ok(match date_arg(args, 0, host)? {
        None => Value::Null,
        Some(d) => Value::Integer(f(d) as i16),
    })
}

fn f_weekday(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 2)?;
    let first = first_day_arg(args, 1)?;
    Ok(match date_arg(args, 0, host)? {
        None => Value::Null,
        Some(d) => Value::Integer(datetime::weekday(d, first) as i16),
    })
}

fn f_weekday_name(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let n = req_i64(args, 0)?;
    let abbrev = match opt(args, 1)? {
        None => false,
        Some(v) => not_null(v)?.to_bool()?,
    };
    let first = first_day_arg(args, 2)?;
    if !(1..=7).contains(&n) {
        return Err(RuntimeError::invalid_call());
    }
    let idx = ((n - 1 + i64::from(first) - 1) % 7) as usize;
    let name = datetime::WEEKDAYS[idx];
    Ok(Value::str(if abbrev { &name[..3] } else { name }))
}

fn f_month_name(args: &[Value]) -> Result<Value> {
    max_args(args, 2)?;
    let n = req_i64(args, 0)?;
    let abbrev = match opt(args, 1)? {
        None => false,
        Some(v) => not_null(v)?.to_bool()?,
    };
    if !(1..=12).contains(&n) {
        return Err(RuntimeError::invalid_call());
    }
    let name = datetime::MONTHS[n as usize - 1];
    Ok(Value::str(if abbrev { &name[..3] } else { name }))
}

fn checked_date(d: f64) -> Result<Value> {
    if datetime::in_range(d) { Ok(Value::Date(d)) } else { Err(RuntimeError::invalid_call()) }
}

fn f_date_serial(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let y = req_i64(args, 0)?;
    let m = req_i64(args, 1)?;
    let d = req_i64(args, 2)?;
    let lim = 1_000_000;
    if [y, m, d].iter().any(|v| v.abs() > lim) {
        return Err(RuntimeError::invalid_call());
    }
    let y = datetime::window_year(y, 2);
    checked_date(ymd_to_serial(y, m, d))
}

fn f_time_serial(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let h = req_i64(args, 0)?;
    let m = req_i64(args, 1)?;
    let s = req_i64(args, 2)?;
    let total = h.saturating_mul(3600).saturating_add(m.saturating_mul(60)).saturating_add(s);
    // ORACLE: negative totals wrap to the previous day's time (TimeSerial(-1, 0, 0) = 11 PM).
    checked_date(datetime::from_linear(total as f64 / 86_400.0))
}

fn f_date_value(args: &[Value], host: &mut dyn Host, date_part: bool) -> Result<Value> {
    max_args(args, 1)?;
    Ok(match date_arg(args, 0, host)? {
        None => Value::Null,
        Some(d) => {
            let (day, frac) = datetime::split(d);
            Value::Date(if date_part { day as f64 } else { frac })
        }
    })
}

fn add_months(d: f64, months: i64) -> Result<f64> {
    let (y, m, day) = datetime::ymd(d);
    let total = y * 12 + i64::from(m) - 1 + months;
    let (ny, nm) = (total.div_euclid(12), (total.rem_euclid(12) + 1) as u32);
    if !(100..=9999).contains(&ny) {
        return Err(RuntimeError::invalid_call());
    }
    let nd = day.min(datetime::days_in_month(ny, nm));
    let serial = ymd_to_serial(ny, i64::from(nm), i64::from(nd)) as i64;
    Ok(datetime::join(serial, datetime::split(d).1))
}

/// Whole seconds of a serial on the monotonic time line.
fn seconds(d: f64) -> i64 {
    (datetime::linear(d) * 86_400.0).round() as i64
}

fn from_seconds(s: i64) -> f64 {
    datetime::from_linear(s as f64 / 86_400.0)
}

fn interval_arg(args: &[Value], i: usize) -> Result<String> {
    Ok(req_text(args, i)?.to_ascii_lowercase())
}

fn f_date_add(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 3)?;
    let iv = interval_arg(args, 0)?;
    // ORACLE: a fractional number is rounded (banker's) before use.
    let n = req_i64(args, 1)?;
    let Some(d) = date_arg(args, 2, host)? else { return Ok(Value::Null) };
    let r = match iv.as_str() {
        "yyyy" => add_months(d, n.saturating_mul(12))?,
        "q" => add_months(d, n.saturating_mul(3))?,
        "m" => add_months(d, n)?,
        "y" | "d" | "w" => from_seconds(seconds(d).saturating_add(n.saturating_mul(86_400))),
        "ww" => from_seconds(seconds(d).saturating_add(n.saturating_mul(7 * 86_400))),
        "h" => from_seconds(seconds(d).saturating_add(n.saturating_mul(3600))),
        "n" => from_seconds(seconds(d).saturating_add(n.saturating_mul(60))),
        "s" => from_seconds(seconds(d).saturating_add(n)),
        _ => return Err(RuntimeError::invalid_call()),
    };
    checked_date(r)
}

fn f_date_diff(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 5)?;
    let iv = interval_arg(args, 0)?;
    let a = date_arg(args, 1, host)?;
    let b = date_arg(args, 2, host)?;
    let first_day = first_day_arg(args, 3)?;
    let _first_week = first_week_arg(args, 4)?;
    let (Some(a), Some(b)) = (a, b) else { return Ok(Value::Null) };
    let (ya, ma, _) = datetime::ymd(a);
    let (yb, mb, _) = datetime::ymd(b);
    let day = |d: f64| datetime::linear(d).floor() as i64;
    let (sa, sb) = (seconds(a), seconds(b));
    let r = match iv.as_str() {
        "yyyy" => yb - ya,
        "q" => (yb * 4 + i64::from((mb - 1) / 3)) - (ya * 4 + i64::from((ma - 1) / 3)),
        "m" => (yb * 12 + i64::from(mb)) - (ya * 12 + i64::from(ma)),
        "y" | "d" => day(b) - day(a),
        "w" => (day(b) - day(a)) / 7,
        "ww" => (datetime::week_start(day(b), first_day) - datetime::week_start(day(a), first_day)) / 7,
        "h" => sb.div_euclid(3600) - sa.div_euclid(3600),
        "n" => sb.div_euclid(60) - sa.div_euclid(60),
        "s" => sb - sa,
        _ => return Err(RuntimeError::invalid_call()),
    };
    Ok(Value::Long(i32::try_from(r).map_err(|_| RuntimeError::overflow())?))
}

fn f_date_part(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 4)?;
    let iv = interval_arg(args, 0)?;
    let first_day = first_day_arg(args, 2)?;
    let first_week = first_week_arg(args, 3)?;
    let Some(d) = date_arg(args, 1, host)? else { return Ok(Value::Null) };
    let (y, m, day) = datetime::ymd(d);
    let (h, mi, s) = datetime::hms(d);
    let r: i64 = match iv.as_str() {
        "yyyy" => y,
        "q" => i64::from((m - 1) / 3 + 1),
        "m" => i64::from(m),
        "y" => d.trunc() as i64 - ymd_to_serial(y, 1, 1) as i64 + 1,
        "d" => i64::from(day),
        "w" => i64::from(datetime::weekday(d, first_day)),
        "ww" => datetime::week_of_year(d, first_day, first_week),
        "h" => i64::from(h),
        "n" => i64::from(mi),
        "s" => i64::from(s),
        _ => return Err(RuntimeError::invalid_call()),
    };
    Ok(Value::Integer(r as i16))
}

fn f_isdate(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    let v = one(args)?;
    Ok(Value::Boolean(match v {
        Value::Date(_) => true,
        Value::String(s) => datetime::parse_date(&s, current_year(host)).is_some(),
        _ => false,
    }))
}

// ---------------------------------------------------------------------------------------------
// Information

fn is_numeric(v: &Value) -> bool {
    match v {
        Value::Empty
        | Value::Boolean(_)
        | Value::Byte(_)
        | Value::Integer(_)
        | Value::Long(_)
        | Value::LongLong(_)
        | Value::Single(_)
        | Value::Double(_)
        | Value::Currency(_)
        | Value::Decimal(_) => true,
        Value::String(s) => parse_number(s).is_some(),
        _ => false,
    }
}

fn f_ismissing(args: &[Value]) -> Result<Value> {
    max_args(args, 1)?;
    match args.first() {
        None => Err(RuntimeError::standard(449)),
        Some(v) => Ok(Value::Boolean(matches!(v, Value::Error(448)))),
    }
}

/// `TypeName`. Arrays carry no element type, so every array reports `Variant()` (observed for
/// `Array(1)`). ORACLE: typed arrays (`Dim a() As String`, `Split`) report `String()` on the
/// real host; that needs an element type on [`Array`].
fn type_name(v: &Value) -> String {
    match v {
        Value::Array(a) => format!("{}()", a.borrow().elem_type.clone().unwrap_or_else(|| "Variant".into())),
        other => other.type_name(),
    }
}

fn f_vartype(args: &[Value]) -> Result<Value> {
    let v = one_raw(args)?;
    // VarType returns a Long (observed: VarType(1) is Long 2).
    Ok(Value::Long(match &v {
        Value::Object(Some(o)) => match o.default_value() {
            Ok(d) if !d.is_object() => d.var_type(),
            _ => 9,
        },
        other => other.var_type(),
    }))
}

fn f_rgb(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let mut c = 0i32;
    for (i, shift) in [(0, 0), (1, 8), (2, 16)] {
        let n = req_i64(args, i)?;
        if n < 0 {
            return Err(RuntimeError::invalid_call());
        }
        c |= (n.min(255) as i32) << shift;
    }
    Ok(Value::Long(c))
}

fn f_qbcolor(args: &[Value]) -> Result<Value> {
    const TABLE: [i32; 16] = [
        0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0xFF0000, 0x00FF00,
        0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
    ];
    let n = not_null(one(args)?)?.to_i64()?;
    usize::try_from(n).ok().and_then(|i| TABLE.get(i)).map(|&c| Value::Long(c)).ok_or_else(RuntimeError::invalid_call)
}

// ---------------------------------------------------------------------------------------------
// Arrays

fn f_bound(args: &[Value], lower: bool) -> Result<Value> {
    max_args(args, 2)?;
    let v = req(args, 0)?;
    let dim = opt_i64(args, 1, 1)?;
    let Value::Array(a) = v else { return Err(RuntimeError::type_mismatch()) };
    let a = a.borrow();
    if dim < 1 {
        return Err(RuntimeError::subscript());
    }
    let (lo, hi) = *a.bounds.get(dim as usize - 1).ok_or_else(RuntimeError::subscript)?;
    let n = if lower { lo } else { hi };
    Ok(Value::Long(i32::try_from(n).map_err(|_| RuntimeError::overflow())?))
}

// ---------------------------------------------------------------------------------------------
// Interaction

fn f_msgbox(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 5)?;
    let prompt = req_text(args, 0)?;
    let buttons = opt_i64(args, 1, 0)?;
    let title = opt_text(args, 2, "")?;
    let buttons = i32::try_from(buttons).map_err(|_| RuntimeError::overflow())?;
    Ok(Value::Long(host.msgbox(&prompt, buttons, &title)))
}

fn f_inputbox(args: &[Value], host: &mut dyn Host) -> Result<Value> {
    max_args(args, 7)?;
    let prompt = req_text(args, 0)?;
    let title = opt_text(args, 1, "")?;
    let default = opt_text(args, 2, "")?;
    Ok(Value::str(host.inputbox(&prompt, &title, &default).unwrap_or_default()))
}

fn f_iif(args: &[Value]) -> Result<Value> {
    max_args(args, 3)?;
    let cond = req(args, 0)?;
    let pick = if cond.is_null() { false } else { cond.to_bool()? };
    let i = if pick { 1 } else { 2 };
    Ok(args.get(i).cloned().unwrap_or(Value::Empty))
}

fn f_choose(args: &[Value]) -> Result<Value> {
    let index = not_null(req(args, 0)?)?.to_f64()?;
    // ORACLE: a fractional index is truncated (Choose(1.7, a, b) = a).
    let i = index.trunc();
    if i < 1.0 || i > (args.len() - 1) as f64 {
        return Ok(Value::Null);
    }
    Ok(args.get(i as usize).cloned().unwrap_or(Value::Null))
}

fn f_switch(args: &[Value]) -> Result<Value> {
    if !args.len().is_multiple_of(2) {
        return Err(RuntimeError::invalid_call());
    }
    for pair in args.chunks(2) {
        let cond = pair[0].clone().deref_default()?;
        if not_null(cond)?.to_bool()? {
            return Ok(pair[1].clone());
        }
    }
    Ok(Value::Null)
}

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::panic)]
mod tests;

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::cloned_ref_to_slice_refs)]
mod oracle_tests;
