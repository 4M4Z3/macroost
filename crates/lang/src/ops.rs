//! Operators ([MS-VBAL] 5.6.9): arithmetic, concatenation, comparison, `Like` and the logical
//! (bitwise) operators, with the language's promotion rules.
//!
//! The interpreter passes runtime values only; declared types are not known here. Where the
//! specification makes a result depend on a declared type (for example "if one operand is
//! declared Variant an Integer overflow widens to Long"), this module takes the behaviour of
//! statically typed operands and the choice is marked `ORACLE:`.

use std::cmp::Ordering;

use crate::error::{Result, RuntimeError};
use crate::stdlib::datetime;
use crate::value::{Value, parse_number, round_half_even};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    IntDiv,
    Mod,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Like,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

/// Evaluates a binary operator. `text_compare` is true under `Option Compare Text`.
pub fn binary(op: BinOp, a: &Value, b: &Value, text_compare: bool) -> Result<Value> {
    let a = a.clone().deref_default()?;
    let b = b.clone().deref_default()?;
    match op {
        BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::IntDiv | BinOp::Mod | BinOp::Pow => {
            arithmetic(op, &a, &b)
        }
        BinOp::Concat => concat(&a, &b),
        BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
            let Some(ord) = compare(&a, &b, text_compare)? else {
                return Ok(Value::Null);
            };
            Ok(Value::Boolean(match op {
                BinOp::Eq => ord == Ordering::Equal,
                BinOp::Ne => ord != Ordering::Equal,
                BinOp::Lt => ord == Ordering::Less,
                BinOp::Le => ord != Ordering::Greater,
                BinOp::Gt => ord == Ordering::Greater,
                _ => ord != Ordering::Less,
            }))
        }
        BinOp::Like => {
            if a.is_null() || b.is_null() {
                return Ok(Value::Null);
            }
            check_data(&a)?;
            check_data(&b)?;
            Ok(Value::Boolean(like(&a.to_text()?, &b.to_text()?, text_compare)?))
        }
        BinOp::And | BinOp::Or | BinOp::Xor | BinOp::Eqv | BinOp::Imp => logical(op, &a, &b),
    }
}

/// Unary minus.
pub fn negate(a: &Value) -> Result<Value> {
    let a = a.clone().deref_default()?;
    Ok(match a {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(if b { 1 } else { 0 }),
        Value::Byte(n) => Value::Integer(-i16::from(n)),
        Value::Integer(n) => Value::Integer(n.checked_neg().ok_or_else(RuntimeError::overflow)?),
        Value::Long(n) => Value::Long(n.checked_neg().ok_or_else(RuntimeError::overflow)?),
        Value::LongLong(n) => Value::LongLong(n.checked_neg().ok_or_else(RuntimeError::overflow)?),
        Value::Single(n) => Value::Single(-n),
        Value::Double(n) => Value::Double(-n),
        Value::Currency(n) => Value::Currency(n.checked_neg().ok_or_else(RuntimeError::overflow)?),
        Value::Date(d) => Value::Date(-d),
        Value::String(_) => Value::Double(-a.to_f64()?),
        Value::Object(None) => return Err(RuntimeError::object_not_set()),
        Value::Object(Some(_)) | Value::Error(_) | Value::Array(_) => return Err(RuntimeError::type_mismatch()),
    })
}

/// `Not`: logical on Booleans, bitwise on integers.
pub fn not(a: &Value) -> Result<Value> {
    let a = a.clone().deref_default()?;
    Ok(match a {
        Value::Null => Value::Null,
        Value::Boolean(b) => Value::Boolean(!b),
        Value::Byte(n) => Value::Byte(!n),
        Value::Empty => Value::Integer(-1),
        Value::Integer(n) => Value::Integer(!n),
        Value::Long(n) => Value::Long(!n),
        Value::LongLong(n) => Value::LongLong(!n),
        Value::Object(None) => return Err(RuntimeError::object_not_set()),
        Value::Object(Some(_)) | Value::Error(_) | Value::Array(_) => return Err(RuntimeError::type_mismatch()),
        // Floating, fixed point, Date and String operands are coerced to Long.
        other => Value::Long(!other.to_i32()?),
    })
}

// ---------------------------------------------------------------------------------------------
// Operand classification

/// Effective numeric value types, in promotion order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum Kind {
    Byte,
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
    Date,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Class {
    Null,
    Empty,
    Str,
    Num(Kind),
}

/// Rejects values no simple data operator accepts.
fn check_data(v: &Value) -> Result<()> {
    match v {
        Value::Object(None) => Err(RuntimeError::object_not_set()),
        Value::Object(Some(_)) | Value::Error(_) | Value::Array(_) => Err(RuntimeError::type_mismatch()),
        _ => Ok(()),
    }
}

fn classify(v: &Value) -> Result<Class> {
    check_data(v)?;
    Ok(match v {
        Value::Null => Class::Null,
        Value::Empty => Class::Empty,
        Value::String(_) => Class::Str,
        Value::Boolean(_) | Value::Integer(_) => Class::Num(Kind::Integer),
        Value::Byte(_) => Class::Num(Kind::Byte),
        Value::Long(_) => Class::Num(Kind::Long),
        Value::LongLong(_) => Class::Num(Kind::LongLong),
        Value::Single(_) => Class::Num(Kind::Single),
        Value::Double(_) => Class::Num(Kind::Double),
        Value::Currency(_) => Class::Num(Kind::Currency),
        Value::Date(_) => Class::Num(Kind::Date),
        // Unreachable after check_data; kept total without panicking.
        _ => return Err(RuntimeError::type_mismatch()),
    })
}

/// The standard effective value type of a binary arithmetic operator (5.6.9.3).
fn effective(a: Class, b: Class) -> Kind {
    let kind = |c: Class| match c {
        Class::Num(k) => Some(k),
        Class::Str => Some(Kind::Double),
        Class::Empty | Class::Null => None,
    };
    match (kind(a), kind(b)) {
        (None, None) => Kind::Integer,
        (Some(k), None) | (None, Some(k)) => k,
        (Some(x), Some(y)) => {
            let (lo, hi) = if x <= y { (x, y) } else { (y, x) };
            if hi == Kind::Single && matches!(lo, Kind::Long | Kind::LongLong) { Kind::Double } else { hi }
        }
    }
}

fn int_range(k: Kind) -> (i128, i128) {
    match k {
        Kind::Byte => (0, 255),
        Kind::Integer => (i128::from(i16::MIN), i128::from(i16::MAX)),
        Kind::Long => (i128::from(i32::MIN), i128::from(i32::MAX)),
        _ => (i128::from(i64::MIN), i128::from(i64::MAX)),
    }
}

/// Packs an integer result into the value type `k`, raising Overflow when out of range.
fn int_value(k: Kind, n: i128) -> Result<Value> {
    let (lo, hi) = int_range(k);
    if n < lo || n > hi {
        // ORACLE: with a Variant-declared operand the language widens Integer -> Long -> Double
        // instead of raising; the interpreter does not tell us declared types, so we raise.
        return Err(RuntimeError::overflow());
    }
    Ok(match k {
        Kind::Byte => Value::Byte(n as u8),
        Kind::Integer => Value::Integer(n as i16),
        Kind::Long => Value::Long(n as i32),
        _ => Value::LongLong(n as i64),
    })
}

/// Integer value of an operand already known to fit an integral effective type.
fn int_of(v: &Value) -> Result<i128> {
    Ok(i128::from(v.to_i64()?))
}

/// Currency (scaled by 10,000) of any numeric operand, rounding half to even.
pub(crate) fn currency_of(v: &Value) -> Result<i64> {
    match v {
        Value::Currency(n) => Ok(*n),
        Value::Byte(_) | Value::Integer(_) | Value::Long(_) | Value::LongLong(_) | Value::Boolean(_) | Value::Empty => {
            let n = i128::from(v.to_i64()?) * 10_000;
            i64::try_from(n).map_err(|_| RuntimeError::overflow())
        }
        _ => currency_from_f64(v.to_f64()?),
    }
}

pub(crate) fn currency_from_f64(x: f64) -> Result<i64> {
    let scaled = round_half_even(x * 10_000.0);
    if (-9.223_372_036_854_776e18..9.223_372_036_854_776e18).contains(&scaled) {
        Ok(scaled as i64)
    } else {
        Err(RuntimeError::overflow())
    }
}

/// A floating result, raising Overflow when finite operands produced an infinity or NaN.
fn double_value(x: f64) -> Result<Value> {
    if x.is_finite() { Ok(Value::Double(x)) } else { Err(RuntimeError::overflow()) }
}

fn single_value(x: f64) -> Result<Value> {
    let s = x as f32;
    if s.is_finite() { Ok(Value::Single(s)) } else { Err(RuntimeError::overflow()) }
}

/// A Date result; outside the Date range the result is the Double (5.6.9.3.2, Variant rule).
fn date_value(x: f64) -> Result<Value> {
    if !x.is_finite() {
        return Err(RuntimeError::overflow());
    }
    // ORACLE: statically typed Date operands raise Overflow here instead.
    if (datetime::MIN_DATE..datetime::MAX_DATE_EXCLUSIVE).contains(&x) {
        Ok(Value::Date(x))
    } else {
        Ok(Value::Double(x))
    }
}

/// Numeric value used by arithmetic; Strings must look numeric (Type mismatch otherwise).
fn num_of(v: &Value) -> Result<f64> {
    match v {
        Value::String(s) => parse_number(s).ok_or_else(RuntimeError::type_mismatch),
        other => other.to_f64(),
    }
}

// ---------------------------------------------------------------------------------------------
// Arithmetic

fn arithmetic(op: BinOp, a: &Value, b: &Value) -> Result<Value> {
    let ca = classify(a)?;
    let cb = classify(b)?;
    if ca == Class::Null || cb == Class::Null {
        return Ok(Value::Null);
    }
    match op {
        BinOp::Add => {
            match (ca, cb) {
                (Class::Str, Class::Str) => return Ok(Value::str(format!("{}{}", a.to_text()?, b.to_text()?))),
                // ORACLE: the spec's table makes String + Empty a Double addition; observed
                // behaviour of the language is concatenation with "" which we follow.
                (Class::Str, Class::Empty) => return Ok(a.clone()),
                (Class::Empty, Class::Str) => return Ok(b.clone()),
                _ => {}
            }
            add_sub(effective(ca, cb), a, b, false)
        }
        BinOp::Sub => {
            if ca == Class::Num(Kind::Date) && cb == Class::Num(Kind::Date) {
                return double_value(a.to_f64()? - b.to_f64()?);
            }
            add_sub(effective(ca, cb), a, b, true)
        }
        BinOp::Mul => {
            let k = effective(ca, cb);
            let float_with_currency = (ca == Class::Num(Kind::Currency) || cb == Class::Num(Kind::Currency))
                && [ca, cb]
                    .iter()
                    .any(|c| matches!(c, Class::Str | Class::Num(Kind::Single) | Class::Num(Kind::Double)));
            match k {
                _ if float_with_currency => double_value(num_of(a)? * num_of(b)?),
                Kind::Byte | Kind::Integer | Kind::Long | Kind::LongLong => int_value(k, int_of(a)? * int_of(b)?),
                Kind::Single => single_value(num_of(a)? * num_of(b)?),
                Kind::Currency => {
                    let p = i128::from(currency_of(a)?) * i128::from(currency_of(b)?);
                    let q = div_round_half_even(p, 10_000);
                    Ok(Value::Currency(i64::try_from(q).map_err(|_| RuntimeError::overflow())?))
                }
                // ORACLE: Date * x - spec gives a Double declared type; we return Double.
                Kind::Double | Kind::Date => double_value(num_of(a)? * num_of(b)?),
            }
        }
        BinOp::Div => {
            let x = num_of(a)?;
            let y = num_of(b)?;
            if y == 0.0 {
                let float_left = matches!(a, Value::Single(_) | Value::Double(_) | Value::String(_) | Value::Date(_));
                if x == 0.0 && !(float_left && cb == Class::Empty) {
                    return Err(RuntimeError::overflow());
                }
                return Err(RuntimeError::div_zero());
            }
            let single = effective(ca, cb) == Kind::Single;
            if single { single_value(x / y) } else { double_value(x / y) }
        }
        BinOp::IntDiv | BinOp::Mod => {
            let small = |c: Class| matches!(c, Class::Empty | Class::Num(Kind::Byte | Kind::Integer));
            let k = if ca == Class::Num(Kind::LongLong) || cb == Class::Num(Kind::LongLong) {
                Kind::LongLong
            } else if ca == Class::Num(Kind::Byte) && cb == Class::Num(Kind::Byte) {
                Kind::Byte
            } else if small(ca) && small(cb) {
                Kind::Integer
            } else {
                // ORACLE: the spec's table makes Integer \ Double an Integer operation; common
                // documentation says any floating operand gives Long. We use Long.
                Kind::Long
            };
            let (lo, hi) = int_range(k);
            let x = int_of(a)?;
            let y = int_of(b)?;
            if x < lo || x > hi || y < lo || y > hi {
                return Err(RuntimeError::overflow());
            }
            if y == 0 {
                return Err(RuntimeError::div_zero());
            }
            // Truncating division; the remainder takes the dividend's sign (-7 Mod 3 = -1).
            // ORACLE: the spec text says "absolute value of the difference", which would give 1.
            if op == BinOp::IntDiv { int_value(k, x / y) } else { int_value(k, x % y) }
        }
        _ => {
            // BinOp::Pow
            let x = num_of(a)?;
            let y = num_of(b)?;
            if x == 0.0 && y == 0.0 {
                return Ok(Value::Double(1.0));
            }
            if x == 0.0 && y < 0.0 {
                return Err(RuntimeError::invalid_call());
            }
            if x < 0.0 && y.fract() != 0.0 {
                return Err(RuntimeError::invalid_call());
            }
            let r = x.powf(y);
            if r.is_nan() { Err(RuntimeError::invalid_call()) } else { double_value(r) }
        }
    }
}

fn add_sub(k: Kind, a: &Value, b: &Value, sub: bool) -> Result<Value> {
    let sign = if sub { -1.0 } else { 1.0 };
    match k {
        Kind::Byte | Kind::Integer | Kind::Long | Kind::LongLong => {
            let (x, y) = (int_of(a)?, int_of(b)?);
            int_value(k, if sub { x - y } else { x + y })
        }
        Kind::Single => single_value(num_of(a)? + sign * num_of(b)?),
        Kind::Double => double_value(num_of(a)? + sign * num_of(b)?),
        Kind::Currency => {
            let (x, y) = (currency_of(a)?, currency_of(b)?);
            let r = if sub { x.checked_sub(y) } else { x.checked_add(y) };
            Ok(Value::Currency(r.ok_or_else(RuntimeError::overflow)?))
        }
        Kind::Date => date_value(num_of(a)? + sign * num_of(b)?),
    }
}

fn div_round_half_even(p: i128, d: i128) -> i128 {
    let q = p.div_euclid(d);
    let r = p.rem_euclid(d);
    match (2 * r).cmp(&d) {
        Ordering::Less => q,
        Ordering::Greater => q + 1,
        Ordering::Equal => {
            if q % 2 == 0 {
                q
            } else {
                q + 1
            }
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Concatenation

fn concat(a: &Value, b: &Value) -> Result<Value> {
    if a.is_null() && b.is_null() {
        return Ok(Value::Null);
    }
    let part = |v: &Value| -> Result<String> {
        check_data(v)?;
        if v.is_null() { Ok(String::new()) } else { v.to_text() }
    };
    let mut s = part(a)?;
    s.push_str(&part(b)?);
    Ok(Value::str(s))
}

// ---------------------------------------------------------------------------------------------
// Comparison

/// Compares two data values; `None` when either is Null.
fn compare(a: &Value, b: &Value, text: bool) -> Result<Option<Ordering>> {
    match (a, b) {
        (Value::Null, _) | (_, Value::Null) => return Ok(None),
        (Value::Error(x), Value::Error(y)) => return Ok(Some(x.cmp(y))),
        _ => {}
    }
    let ca = classify(a)?;
    let cb = classify(b)?;
    Ok(Some(match (ca, cb) {
        (Class::Str, Class::Str) => compare_text(&a.to_text()?, &b.to_text()?, text),
        (Class::Str, Class::Empty) => compare_text(&a.to_text()?, "", text),
        (Class::Empty, Class::Str) => compare_text("", &b.to_text()?, text),
        (Class::Empty, Class::Empty) => Ordering::Equal,
        (Class::Num(_), Class::Str) => return compare_num_str(a, b),
        (Class::Str, Class::Num(_)) => return Ok(compare_num_str(b, a)?.map(Ordering::reverse)),
        _ => compare_numbers(a, b, ca, cb)?,
    }))
}

/// Number (left) against String (right). A numeric-looking string compares numerically;
/// otherwise the number is less than the string.
fn compare_num_str(n: &Value, s: &Value) -> Result<Option<Ordering>> {
    let text = s.to_text()?;
    let parsed = match n {
        Value::Boolean(_) => match Value::str(text.as_str()).to_bool() {
            Ok(b) => Some(if b { -1.0 } else { 0.0 }),
            Err(_) => None,
        },
        Value::Date(_) => datetime::parse_date(&text, 2000).or_else(|| parse_number(&text)),
        _ => parse_number(&text),
    };
    // ORACLE: when both operands are Variants the language always orders the number first,
    // even for a numeric-looking string; with a statically typed number it converts the string.
    // We convert when the string looks numeric (the common case of a literal or typed number).
    match parsed {
        Some(y) => Ok(Some(compare_f64(n.to_f64()?, y)?)),
        None => Ok(Some(Ordering::Less)),
    }
}

fn compare_f64(x: f64, y: f64) -> Result<Ordering> {
    x.partial_cmp(&y).ok_or_else(RuntimeError::overflow)
}

fn compare_numbers(a: &Value, b: &Value, ca: Class, cb: Class) -> Result<Ordering> {
    let integral =
        |c: Class| matches!(c, Class::Empty | Class::Num(Kind::Byte | Kind::Integer | Kind::Long | Kind::LongLong));
    if integral(ca) && integral(cb) {
        return Ok(a.to_i64()?.cmp(&b.to_i64()?));
    }
    let has = |k: Kind| ca == Class::Num(k) || cb == Class::Num(k);
    if has(Kind::Date) {
        return compare_f64(a.to_f64()?, b.to_f64()?);
    }
    if has(Kind::Currency) {
        return Ok(currency_of(a)?.cmp(&currency_of(b)?));
    }
    if has(Kind::Single) && !has(Kind::Long) && !has(Kind::LongLong) {
        // Single against Double compares at Single precision (5.6.9.5 table).
        let x = a.to_f64()? as f32;
        let y = b.to_f64()? as f32;
        return x.partial_cmp(&y).ok_or_else(RuntimeError::overflow);
    }
    compare_f64(a.to_f64()?, b.to_f64()?)
}

/// Case folding for text comparison of one UTF-16 unit.
pub(crate) fn fold_unit(u: u16) -> u16 {
    if u < 128 {
        return u8::try_from(u).map(|b| u16::from(b.to_ascii_lowercase())).unwrap_or(u);
    }
    match char::from_u32(u32::from(u)) {
        Some(c) => {
            let mut l = c.to_lowercase();
            match (l.next(), l.next()) {
                (Some(x), None) => u16::try_from(u32::from(x)).unwrap_or(u),
                _ => u,
            }
        }
        None => u,
    }
}

/// Binary comparison orders UTF-16 code units; text comparison folds case first.
/// ORACLE: text mode in the language uses the host's locale collation (punctuation and accent
/// ordering differ from code-unit order); we fold case only.
pub(crate) fn compare_text(a: &str, b: &str, text: bool) -> Ordering {
    let ua = a.encode_utf16();
    let ub = b.encode_utf16();
    if text { ua.map(fold_unit).cmp(ub.map(fold_unit)) } else { ua.cmp(ub) }
}

// ---------------------------------------------------------------------------------------------
// Logical operators

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum LKind {
    Boolean,
    Byte,
    Integer,
    Long,
    LongLong,
}

fn logical_operand(v: &Value) -> Result<Option<(LKind, i64)>> {
    check_data(v)?;
    Ok(Some(match v {
        Value::Null => return Ok(None),
        Value::Boolean(b) => (LKind::Boolean, if *b { -1 } else { 0 }),
        Value::Byte(n) => (LKind::Byte, i64::from(*n)),
        Value::Empty => (LKind::Integer, 0),
        Value::Integer(n) => (LKind::Integer, i64::from(*n)),
        Value::Long(n) => (LKind::Long, i64::from(*n)),
        Value::LongLong(n) => (LKind::LongLong, *n),
        other => (LKind::Long, i64::from(other.to_i32()?)),
    }))
}

fn logical_kind(a: LKind, b: LKind) -> LKind {
    match (a, b) {
        (LKind::Boolean, LKind::Boolean) => LKind::Boolean,
        (LKind::Byte, LKind::Byte) => LKind::Byte,
        _ => {
            let hi = a.max(b);
            if hi <= LKind::Integer { LKind::Integer } else { hi }
        }
    }
}

fn logical_value(k: LKind, n: i64) -> Value {
    match k {
        LKind::Boolean => Value::Boolean(n != 0),
        LKind::Byte => Value::Byte(n as u8),
        LKind::Integer => Value::Integer(n as i16),
        LKind::Long => Value::Long(n as i32),
        LKind::LongLong => Value::LongLong(n),
    }
}

fn logical(op: BinOp, a: &Value, b: &Value) -> Result<Value> {
    let la = logical_operand(a)?;
    let lb = logical_operand(b)?;
    let bits = |x: i64, y: i64| match op {
        BinOp::And => x & y,
        BinOp::Or => x | y,
        BinOp::Xor => x ^ y,
        BinOp::Eqv => !(x ^ y),
        _ => !x | y,
    };
    match (la, lb) {
        (Some((ka, x)), Some((kb, y))) => Ok(logical_value(logical_kind(ka, kb), bits(x, y))),
        (None, None) => Ok(Value::Null),
        (Some((k, x)), None) | (None, Some((k, x))) => {
            let left_known = la.is_some();
            // With a Boolean operand and Null the result is Boolean; Byte stays Byte; other
            // integers give Integer at least.
            let kind = match k {
                LKind::Boolean => LKind::Boolean,
                LKind::Byte => LKind::Byte,
                other => other.max(LKind::Integer),
            };
            let result = match op {
                BinOp::And => (x == 0).then_some(0),
                // ORACLE: the spec's table returns the integral operand for `x Or Null` even
                // when it is 0; documented behaviour is False Or Null = Null. We return Null for 0.
                BinOp::Or => (x != 0).then_some(x),
                BinOp::Xor | BinOp::Eqv => None,
                _ => {
                    // Imp
                    if left_known { (x != -1).then_some(!x) } else { (x != 0).then_some(x) }
                }
            };
            Ok(result.map_or(Value::Null, |n| logical_value(kind, n)))
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Like

enum Pat {
    Char(char),
    Any,
    Digit,
    Star,
    List { negated: bool, ranges: Vec<(char, char)> },
}

fn parse_pattern(pattern: &str) -> Result<Vec<Pat>> {
    let invalid = || RuntimeError::standard(93);
    let chars: Vec<char> = pattern.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        i += 1;
        match c {
            '?' => out.push(Pat::Any),
            '#' => out.push(Pat::Digit),
            '*' => {
                if !matches!(out.last(), Some(Pat::Star)) {
                    out.push(Pat::Star);
                }
            }
            '[' => {
                let mut negated = false;
                if chars.get(i) == Some(&'!') {
                    negated = true;
                    i += 1;
                }
                let mut ranges = Vec::new();
                if chars.get(i) == Some(&'-') {
                    ranges.push(('-', '-'));
                    i += 1;
                }
                loop {
                    let Some(&ch) = chars.get(i) else {
                        return Err(invalid());
                    };
                    i += 1;
                    if ch == ']' {
                        break;
                    }
                    if chars.get(i) == Some(&'-') && chars.get(i + 1).is_some_and(|&n| n != ']') {
                        let end = chars[i + 1];
                        if end < ch {
                            return Err(invalid());
                        }
                        ranges.push((ch, end));
                        i += 2;
                    } else {
                        ranges.push((ch, ch));
                    }
                }
                // ORACLE: an empty list `[]` matches the empty string; we drop the element.
                if !ranges.is_empty() || negated {
                    out.push(Pat::List { negated, ranges });
                }
            }
            other => out.push(Pat::Char(other)),
        }
    }
    Ok(out)
}

fn fold_char(c: char) -> char {
    let mut l = c.to_lowercase();
    match (l.next(), l.next()) {
        (Some(x), None) => x,
        _ => c,
    }
}

fn pat_matches(p: &Pat, c: char, text: bool) -> bool {
    match p {
        Pat::Char(x) => *x == c || (text && fold_char(*x) == fold_char(c)),
        Pat::Any => true,
        Pat::Digit => c.is_ascii_digit(),
        Pat::Star => false,
        Pat::List { negated, ranges } => {
            let hit = ranges.iter().any(|&(lo, hi)| {
                (lo..=hi).contains(&c)
                    || (text && {
                        let f = fold_char(c);
                        let up = c.to_uppercase().next().unwrap_or(c);
                        (fold_char(lo)..=fold_char(hi)).contains(&f) || (lo..=hi).contains(&up)
                    })
            });
            hit != *negated
        }
    }
}

/// `Like` pattern matching (`?`, `*`, `#`, `[list]`, `[!list]`).
pub fn like(text: &str, pattern: &str, text_compare: bool) -> Result<bool> {
    let pats = parse_pattern(pattern)?;
    let t: Vec<char> = text.chars().collect();
    let (mut ti, mut pi) = (0usize, 0usize);
    let mut star: Option<(usize, usize)> = None;
    while ti < t.len() {
        match pats.get(pi) {
            Some(Pat::Star) => {
                star = Some((pi, ti));
                pi += 1;
            }
            Some(p) if pat_matches(p, t[ti], text_compare) => {
                ti += 1;
                pi += 1;
            }
            _ => match star {
                Some((sp, st)) => {
                    pi = sp + 1;
                    ti = st + 1;
                    star = Some((sp, st + 1));
                }
                None => return Ok(false),
            },
        }
    }
    while matches!(pats.get(pi), Some(Pat::Star)) {
        pi += 1;
    }
    Ok(pi == pats.len())
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;
    use BinOp::*;

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
    fn ev(op: BinOp, a: Value, b: Value) -> String {
        match binary(op, &a, &b, false) {
            Ok(v) => format!("{v:?}"),
            Err(e) => format!("err {}", e.number),
        }
    }
    fn evt(op: BinOp, a: Value, b: Value) -> String {
        match binary(op, &a, &b, true) {
            Ok(v) => format!("{v:?}"),
            Err(e) => format!("err {}", e.number),
        }
    }

    #[test]
    fn addition_and_promotion() {
        assert_eq!(ev(Add, i(1), i(2)), "Integer(3)");
        assert_eq!(ev(Add, i(32767), i(1)), "err 6");
        assert_eq!(ev(Add, l(1), i(2)), "Long(3)");
        assert_eq!(ev(Add, Value::Byte(200), Value::Byte(55)), "Byte(255)");
        assert_eq!(ev(Add, Value::Byte(200), Value::Byte(56)), "err 6");
        assert_eq!(ev(Add, Value::Byte(1), Value::Empty), "Byte(1)");
        assert_eq!(ev(Add, Value::Empty, Value::Empty), "Integer(0)");
        assert_eq!(ev(Add, Value::Single(1.5), i(1)), "Single(2.5)");
        assert_eq!(ev(Add, Value::Single(1.5), l(1)), "Double(2.5)");
        assert_eq!(ev(Add, d(1.5), l(1)), "Double(2.5)");
        assert_eq!(ev(Add, Value::LongLong(5), l(1)), "LongLong(6)");
        assert_eq!(ev(Add, s("10"), i(5)), "Double(15)");
        assert_eq!(ev(Add, s("a"), s("b")), "\"ab\"");
        assert_eq!(ev(Add, s("abc"), i(5)), "err 13");
        assert_eq!(ev(Add, Value::Null, i(5)), "Null");
        assert_eq!(ev(Add, Value::Boolean(true), i(5)), "Integer(4)");
        assert_eq!(ev(Add, Value::Currency(15_000), i(1)), "Currency(2.5)");
        assert_eq!(ev(Add, Value::Date(46309.0), i(1)), "Date(46310)");
        assert_eq!(ev(Add, Value::Error(5), i(1)), "err 13");
        assert_eq!(ev(Add, Value::Object(None), i(1)), "err 91");
        assert_eq!(ev(Add, s("x"), Value::Empty), "\"x\"");
    }

    #[test]
    fn subtraction_multiplication() {
        assert_eq!(ev(Sub, Value::Date(46310.0), Value::Date(46309.0)), "Double(1)");
        assert_eq!(ev(Sub, Value::Date(46310.0), i(1)), "Date(46309)");
        assert_eq!(ev(Sub, i(-32768), i(1)), "err 6");
        assert_eq!(ev(Mul, i(300), i(200)), "err 6");
        assert_eq!(ev(Mul, l(300), i(200)), "Long(60000)");
        assert_eq!(ev(Mul, Value::Currency(25_000), Value::Currency(20_000)), "Currency(5)");
        assert_eq!(ev(Mul, Value::Currency(25_000), d(2.0)), "Double(5)");
        assert_eq!(ev(Mul, d(1e308), d(10.0)), "err 6");
    }

    #[test]
    fn division() {
        assert_eq!(ev(Div, i(7), i(2)), "Double(3.5)");
        assert_eq!(ev(Div, Value::Single(1.0), Value::Single(4.0)), "Single(0.25)");
        assert_eq!(ev(Div, i(1), i(0)), "err 11");
        assert_eq!(ev(Div, i(0), i(0)), "err 6");
        assert_eq!(ev(Div, d(0.0), Value::Empty), "err 11");
        assert_eq!(ev(IntDiv, i(7), i(2)), "Integer(3)");
        assert_eq!(ev(IntDiv, i(-7), i(2)), "Integer(-3)");
        assert_eq!(ev(IntDiv, d(7.5), i(2)), "Long(4)");
        assert_eq!(ev(IntDiv, d(6.5), i(2)), "Long(3)");
        assert_eq!(ev(IntDiv, i(1), i(0)), "err 11");
        assert_eq!(ev(IntDiv, Value::LongLong(9), i(2)), "LongLong(4)");
        assert_eq!(ev(IntDiv, Value::Byte(9), Value::Byte(2)), "Byte(4)");
        assert_eq!(ev(IntDiv, i(-32768), i(-1)), "err 6");
        assert_eq!(ev(Mod, d(7.5), i(2)), "Long(0)");
        assert_eq!(ev(Mod, i(-7), i(3)), "Integer(-1)");
        assert_eq!(ev(Mod, i(7), i(-3)), "Integer(1)");
        assert_eq!(ev(Mod, i(10), i(3)), "Integer(1)");
        assert_eq!(ev(Mod, i(5), i(0)), "err 11");
        assert_eq!(ev(Mod, d(1e12), i(7)), "err 6");
        assert_eq!(ev(Mod, Value::Null, i(7)), "Null");
    }

    #[test]
    fn power() {
        assert_eq!(ev(Pow, i(2), i(10)), "Double(1024)");
        assert_eq!(ev(Pow, i(0), i(0)), "Double(1)");
        assert_eq!(ev(Pow, i(0), i(-1)), "err 5");
        assert_eq!(ev(Pow, i(-8), d(1.0 / 3.0)), "err 5");
        assert_eq!(ev(Pow, i(-2), i(3)), "Double(-8)");
        assert_eq!(ev(Pow, d(10.0), i(400)), "err 6");
    }

    #[test]
    fn concatenation() {
        assert_eq!(ev(Concat, s("10"), i(5)), "\"105\"");
        assert_eq!(ev(Concat, Value::Null, s("x")), "\"x\"");
        assert_eq!(ev(Concat, Value::Null, Value::Null), "Null");
        assert_eq!(ev(Concat, Value::Empty, i(1)), "\"1\"");
        assert_eq!(ev(Concat, Value::Boolean(true), s("!")), "\"True!\"");
        assert_eq!(ev(Concat, d(1.5), s("")), "\"1.5\"");
        assert_eq!(ev(Concat, Value::Error(5), s("")), "err 13");
    }

    #[test]
    fn comparisons() {
        assert_eq!(ev(Eq, i(1), d(1.0)), "Boolean(true)");
        assert_eq!(ev(Lt, i(1), l(2)), "Boolean(true)");
        assert_eq!(ev(Gt, s("b"), s("a")), "Boolean(true)");
        assert_eq!(ev(Lt, s("B"), s("a")), "Boolean(true)");
        assert_eq!(evt(Lt, s("B"), s("a")), "Boolean(false)");
        assert_eq!(evt(Eq, s("ABC"), s("abc")), "Boolean(true)");
        assert_eq!(ev(Eq, s("ABC"), s("abc")), "Boolean(false)");
        assert_eq!(ev(Lt, s("ab"), s("abc")), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Null, i(1)), "Null");
        assert_eq!(ev(Eq, Value::Empty, i(0)), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Empty, s("")), "Boolean(true)");
        assert_eq!(ev(Eq, s("10"), i(10)), "Boolean(true)");
        assert_eq!(ev(Gt, s("10"), i(9)), "Boolean(true)");
        assert_eq!(ev(Lt, i(1000), s("abc")), "Boolean(true)");
        assert_eq!(ev(Gt, s("abc"), i(1000)), "Boolean(true)");
        assert_eq!(ev(Lt, Value::Boolean(true), Value::Boolean(false)), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Boolean(true), s("True")), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Boolean(true), i(-1)), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Single(0.1), d(0.1)), "Boolean(true)");
        assert_eq!(ev(Ne, Value::Currency(10_000), i(1)), "Boolean(false)");
        assert_eq!(ev(Ge, Value::Date(2.0), i(2)), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Error(5), Value::Error(5)), "Boolean(true)");
        assert_eq!(ev(Eq, Value::Error(5), i(5)), "err 13");
        assert_eq!(ev(Le, Value::LongLong(i64::MAX), Value::LongLong(i64::MAX)), "Boolean(true)");
        assert_eq!(ev(Eq, d(f64::NAN), d(1.0)), "err 6");
    }

    #[test]
    fn logical_ops() {
        let t = Value::Boolean(true);
        let f = Value::Boolean(false);
        assert_eq!(ev(And, t.clone(), f.clone()), "Boolean(false)");
        assert_eq!(ev(Or, t.clone(), f.clone()), "Boolean(true)");
        assert_eq!(ev(Xor, t.clone(), t.clone()), "Boolean(false)");
        assert_eq!(ev(Eqv, f.clone(), f.clone()), "Boolean(true)");
        assert_eq!(ev(Imp, t.clone(), f.clone()), "Boolean(false)");
        assert_eq!(ev(Imp, f.clone(), t.clone()), "Boolean(true)");
        assert_eq!(ev(And, i(12), i(10)), "Integer(8)");
        assert_eq!(ev(Or, i(12), i(3)), "Integer(15)");
        assert_eq!(ev(Xor, i(12), i(10)), "Integer(6)");
        assert_eq!(ev(And, l(-1), i(255)), "Long(255)");
        assert_eq!(ev(And, d(3.6), i(7)), "Long(4)");
        assert_eq!(ev(And, t.clone(), i(5)), "Integer(5)");
        assert_eq!(ev(And, Value::Byte(6), Value::Byte(3)), "Byte(2)");
        assert_eq!(ev(And, f.clone(), Value::Null), "Boolean(false)");
        assert_eq!(ev(And, t.clone(), Value::Null), "Null");
        assert_eq!(ev(Or, t.clone(), Value::Null), "Boolean(true)");
        assert_eq!(ev(Or, f.clone(), Value::Null), "Null");
        assert_eq!(ev(Xor, t.clone(), Value::Null), "Null");
        assert_eq!(ev(Imp, f.clone(), Value::Null), "Boolean(true)");
        assert_eq!(ev(Imp, t.clone(), Value::Null), "Null");
        assert_eq!(ev(Imp, Value::Null, t.clone()), "Boolean(true)");
        assert_eq!(ev(Imp, Value::Null, f.clone()), "Null");
        assert_eq!(ev(And, Value::Null, Value::Null), "Null");
        assert_eq!(ev(And, s("x"), i(1)), "err 13");
    }

    #[test]
    fn unary() {
        assert_eq!(format!("{:?}", negate(&i(5)).unwrap()), "Integer(-5)");
        assert_eq!(negate(&i(-32768)).unwrap_err().number, 6);
        assert_eq!(format!("{:?}", negate(&Value::Byte(5)).unwrap()), "Integer(-5)");
        assert_eq!(format!("{:?}", negate(&s("2.5")).unwrap()), "Double(-2.5)");
        assert_eq!(format!("{:?}", negate(&Value::Null).unwrap()), "Null");
        assert_eq!(format!("{:?}", negate(&Value::Empty).unwrap()), "Integer(0)");
        assert_eq!(format!("{:?}", not(&Value::Boolean(true)).unwrap()), "Boolean(false)");
        assert_eq!(format!("{:?}", not(&i(0)).unwrap()), "Integer(-1)");
        assert_eq!(format!("{:?}", not(&i(5)).unwrap()), "Integer(-6)");
        assert_eq!(format!("{:?}", not(&Value::Byte(0)).unwrap()), "Byte(255)");
        assert_eq!(format!("{:?}", not(&d(2.0)).unwrap()), "Long(-3)");
        assert_eq!(format!("{:?}", not(&Value::Null).unwrap()), "Null");
    }

    #[test]
    fn like_patterns() {
        assert!(like("aBBBa", "a*a", false).unwrap());
        assert!(like("F", "[A-Z]", false).unwrap());
        assert!(!like("F", "[!A-Z]", false).unwrap());
        assert!(like("a2a", "a#a", false).unwrap());
        assert!(like("aM5b", "a[L-P]#[!c-e]", false).unwrap());
        assert!(like("BAT123khg", "B?T*", false).unwrap());
        assert!(!like("CAT123khg", "B?T*", false).unwrap());
        assert!(!like("abc", "ABC", false).unwrap());
        assert!(like("abc", "ABC", true).unwrap());
        assert!(like("f", "[A-Z]", true).unwrap());
        assert!(like("", "*", false).unwrap());
        assert!(!like("", "?", false).unwrap());
        assert!(like("a-b", "a[-]b", false).unwrap());
        assert!(like("a*b", "a[*]b", false).unwrap());
        assert!(like("abcabc", "*abc", false).unwrap());
        assert!(!like("abcab", "*abc", false).unwrap());
        assert_eq!(like("a", "[a", false).unwrap_err().number, 93);
        assert_eq!(like("a", "[z-a]", false).unwrap_err().number, 93);
        assert_eq!(ev(Like, s("abc"), s("a*")), "Boolean(true)");
        assert_eq!(ev(Like, Value::Null, s("a*")), "Null");
        assert_eq!(ev(Like, i(123), s("1##")), "Boolean(true)");
    }
}
