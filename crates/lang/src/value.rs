//! Values and the language's coercion and operator rules ([MS-VBAL] 5.5 and 5.6.9).
//!
//! Anything not yet confirmed against a running the spreadsheet host is marked `ORACLE:`; the oracle harness
//! settles those by black-box comparison.

use std::cell::RefCell;
use std::rc::Rc;

use crate::decimal::Decimal;
use crate::error::{Result, RuntimeError};
use crate::object::ObjectRef;

/// Days between the date origin (1899-12-30) and 1970-01-01.
pub const UNIX_EPOCH_SERIAL: f64 = 25569.0;

#[derive(Clone, Default)]
pub enum Value {
    #[default]
    Empty,
    Null,
    Boolean(bool),
    Byte(u8),
    Integer(i16),
    Long(i32),
    LongLong(i64),
    Single(f32),
    Double(f64),
    /// Fixed point, scaled by 10,000.
    Currency(i64),
    /// Exact decimal (96-bit mantissa, scale 0..=28); only ever held in a Variant.
    Decimal(Decimal),
    /// Serial date: whole days since 1899-12-30, fraction = time of day.
    Date(f64),
    String(Rc<str>),
    /// `Nothing` is `Object(None)`.
    Object(Option<ObjectRef>),
    /// A `CVErr` value (the reference spreadsheet's #N/A etc. arrive as these).
    Error(i32),
    Array(Rc<RefCell<Array>>),
}

impl std::fmt::Debug for Value {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Value::Empty => write!(f, "Empty"),
            Value::Null => write!(f, "Null"),
            Value::Boolean(b) => write!(f, "Boolean({b})"),
            Value::Byte(n) => write!(f, "Byte({n})"),
            Value::Integer(n) => write!(f, "Integer({n})"),
            Value::Long(n) => write!(f, "Long({n})"),
            Value::LongLong(n) => write!(f, "LongLong({n})"),
            Value::Single(n) => write!(f, "Single({n})"),
            Value::Double(n) => write!(f, "Double({n})"),
            Value::Currency(n) => write!(f, "Currency({})", *n as f64 / 10_000.0),
            Value::Decimal(d) => write!(f, "Decimal({d})"),
            Value::Date(d) => write!(f, "Date({d})"),
            Value::String(s) => write!(f, "{s:?}"),
            Value::Object(None) => write!(f, "Nothing"),
            Value::Object(Some(o)) => write!(f, "Object({})", o.type_name()),
            Value::Error(n) => write!(f, "Error({n})"),
            Value::Array(a) => write!(f, "Array({:?})", a.borrow().bounds),
        }
    }
}

/// A (possibly multi-dimensional) array, stored row-major by first index.
#[derive(Debug, Clone, Default)]
pub struct Array {
    /// (lower, upper) per dimension.
    pub bounds: Vec<(i64, i64)>,
    pub data: Vec<Value>,
    /// Declared with fixed bounds (`Dim a(3)`): ReDim is an error.
    pub fixed: bool,
    /// Declared element type (`Dim a() As Long`): elements are coerced to it. `None` = Variant.
    pub elem_type: Option<String>,
}

/// Largest number of elements one array may hold, so a macro cannot exhaust memory.
pub const MAX_ARRAY_ELEMENTS: i64 = 50_000_000;

impl Array {
    pub fn new(bounds: Vec<(i64, i64)>) -> Result<Self> {
        let mut count: i64 = 1;
        for &(lo, hi) in &bounds {
            let n = hi - lo + 1;
            if n < 0 {
                return Err(RuntimeError::subscript());
            }
            count = count.checked_mul(n).ok_or_else(RuntimeError::overflow)?;
        }
        if count > MAX_ARRAY_ELEMENTS {
            return Err(RuntimeError::standard(7));
        }
        Ok(Self { bounds, data: vec![Value::Empty; count as usize], fixed: false, elem_type: None })
    }

    pub fn from_vec(lower: i64, items: Vec<Value>) -> Self {
        let upper = lower + items.len() as i64 - 1;
        Self { bounds: vec![(lower, upper)], data: items, fixed: false, elem_type: None }
    }

    pub fn offset(&self, idx: &[i64]) -> Result<usize> {
        if idx.len() != self.bounds.len() {
            return Err(RuntimeError::subscript());
        }
        let mut off: i64 = 0;
        for (&i, &(lo, hi)) in idx.iter().zip(&self.bounds) {
            if i < lo || i > hi {
                return Err(RuntimeError::subscript());
            }
            off = off * (hi - lo + 1) + (i - lo);
        }
        Ok(off as usize)
    }

    pub fn get(&self, idx: &[i64]) -> Result<Value> {
        let o = self.offset(idx)?;
        self.data.get(o).cloned().ok_or_else(RuntimeError::subscript)
    }

    pub fn set(&mut self, idx: &[i64], v: Value) -> Result<()> {
        let o = self.offset(idx)?;
        let slot = self.data.get_mut(o).ok_or_else(RuntimeError::subscript)?;
        *slot = v;
        Ok(())
    }

    /// `ReDim Preserve`: only the last dimension may change; contents are kept.
    pub fn redim_preserve(&mut self, bounds: Vec<(i64, i64)>) -> Result<()> {
        if self.bounds.is_empty() {
            *self = Array::new(bounds)?;
            return Ok(());
        }
        if bounds.len() != self.bounds.len() || bounds[..bounds.len() - 1] != self.bounds[..self.bounds.len() - 1] {
            return Err(RuntimeError::subscript());
        }
        let mut next = Array::new(bounds)?;
        let last = self.bounds.len() - 1;
        let (old_lo, old_hi) = self.bounds[last];
        let (new_lo, new_hi) = next.bounds[last];
        let mut idx: Vec<i64> = self.bounds.iter().map(|b| b.0).collect();
        // Walk every index of the old array and copy those still in range.
        if self.data.is_empty() {
            *self = next;
            return Ok(());
        }
        loop {
            if idx[last] >= new_lo && idx[last] <= new_hi {
                let v = self.get(&idx)?;
                next.set(&idx, v)?;
            }
            // Increment the multi-index (last dimension fastest).
            let mut d = last;
            loop {
                idx[d] += 1;
                let hi = if d == last { old_hi } else { self.bounds[d].1 };
                if idx[d] <= hi {
                    break;
                }
                idx[d] = if d == last { old_lo } else { self.bounds[d].0 };
                if d == 0 {
                    *self = next;
                    return Ok(());
                }
                d -= 1;
            }
        }
    }
}

impl From<bool> for Value {
    fn from(b: bool) -> Self {
        Value::Boolean(b)
    }
}
impl From<i32> for Value {
    fn from(n: i32) -> Self {
        Value::Long(n)
    }
}
impl From<f64> for Value {
    fn from(n: f64) -> Self {
        Value::Double(n)
    }
}
impl From<&str> for Value {
    fn from(s: &str) -> Self {
        Value::String(s.into())
    }
}
impl From<String> for Value {
    fn from(s: String) -> Self {
        Value::String(s.into())
    }
}

/// Banker's rounding (round half to even), as used by CInt/CLng and integer coercion.
pub fn round_half_even(x: f64) -> f64 {
    let r = x.round();
    if (x - x.trunc()).abs() == 0.5 { 2.0 * (x / 2.0).round() } else { r }
}

impl Value {
    pub fn str(s: impl Into<Rc<str>>) -> Self {
        Value::String(s.into())
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Value::Empty)
    }
    pub fn is_null(&self) -> bool {
        matches!(self, Value::Null)
    }
    pub fn is_object(&self) -> bool {
        matches!(self, Value::Object(_))
    }

    /// `TypeName`.
    pub fn type_name(&self) -> String {
        match self {
            Value::Empty => "Empty".into(),
            Value::Null => "Null".into(),
            Value::Boolean(_) => "Boolean".into(),
            Value::Byte(_) => "Byte".into(),
            Value::Integer(_) => "Integer".into(),
            Value::Long(_) => "Long".into(),
            Value::LongLong(_) => "LongLong".into(),
            Value::Single(_) => "Single".into(),
            Value::Double(_) => "Double".into(),
            Value::Currency(_) => "Currency".into(),
            Value::Decimal(_) => "Decimal".into(),
            Value::Date(_) => "Date".into(),
            Value::String(_) => "String".into(),
            Value::Object(None) => "Nothing".into(),
            Value::Object(Some(o)) => o.type_name(),
            Value::Error(_) => "Error".into(),
            Value::Array(a) => {
                let elem = a.borrow().data.first().map(|v| v.type_name()).unwrap_or_else(|| "Variant".into());
                format!("{}()", if elem == "Empty" { "Variant".into() } else { elem })
            }
        }
    }

    /// `VarType`.
    pub fn var_type(&self) -> i32 {
        match self {
            Value::Empty => 0,
            Value::Null => 1,
            Value::Integer(_) => 2,
            Value::Long(_) => 3,
            Value::Single(_) => 4,
            Value::Double(_) => 5,
            Value::Currency(_) => 6,
            Value::Date(_) => 7,
            Value::String(_) => 8,
            Value::Object(_) => 9,
            Value::Error(_) => 10,
            Value::Boolean(_) => 11,
            Value::Decimal(_) => 14,
            Value::Byte(_) => 17,
            Value::LongLong(_) => 20,
            Value::Array(_) => 8192 + 12,
        }
    }

    pub fn is_numeric_type(&self) -> bool {
        matches!(
            self,
            Value::Boolean(_)
                | Value::Byte(_)
                | Value::Integer(_)
                | Value::Long(_)
                | Value::LongLong(_)
                | Value::Single(_)
                | Value::Double(_)
                | Value::Currency(_)
                | Value::Decimal(_)
                | Value::Date(_)
                | Value::Empty
        )
    }

    /// Numeric value, converting strings the way `CDbl` does.
    pub fn to_f64(&self) -> Result<f64> {
        Ok(match self {
            Value::Empty => 0.0,
            Value::Boolean(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Value::Byte(n) => f64::from(*n),
            Value::Integer(n) => f64::from(*n),
            Value::Long(n) => f64::from(*n),
            Value::LongLong(n) => *n as f64,
            Value::Single(n) => f64::from(*n),
            Value::Double(n) => *n,
            Value::Currency(n) => *n as f64 / 10_000.0,
            Value::Decimal(d) => d.to_f64(),
            Value::Date(d) => *d,
            Value::String(s) => parse_number(s).ok_or_else(RuntimeError::type_mismatch)?,
            Value::Null => return Err(RuntimeError::standard(94)),
            Value::Object(Some(o)) => return o.default_value()?.to_f64(),
            Value::Object(None) => return Err(RuntimeError::object_not_set()),
            Value::Error(_) | Value::Array(_) => return Err(RuntimeError::type_mismatch()),
        })
    }

    pub fn to_i64(&self) -> Result<i64> {
        match self {
            Value::LongLong(n) => Ok(*n),
            Value::Long(n) => Ok(i64::from(*n)),
            Value::Integer(n) => Ok(i64::from(*n)),
            Value::Byte(n) => Ok(i64::from(*n)),
            Value::Decimal(d) => d.to_i64(),
            _ => {
                let r = round_half_even(self.to_f64()?);
                if (-9.223_372_036_854_776e18..9.223_372_036_854_776e18).contains(&r) {
                    Ok(r as i64)
                } else {
                    Err(RuntimeError::overflow())
                }
            }
        }
    }

    pub fn to_i32(&self) -> Result<i32> {
        i32::try_from(self.to_i64()?).map_err(|_| RuntimeError::overflow())
    }

    pub fn to_i16(&self) -> Result<i16> {
        i16::try_from(self.to_i64()?).map_err(|_| RuntimeError::overflow())
    }

    pub fn to_bool(&self) -> Result<bool> {
        match self {
            Value::Boolean(b) => Ok(*b),
            Value::String(s) => {
                let t = s.trim();
                if t.eq_ignore_ascii_case("true") {
                    Ok(true)
                } else if t.eq_ignore_ascii_case("false") {
                    Ok(false)
                } else {
                    Ok(parse_number(t).ok_or_else(RuntimeError::type_mismatch)? != 0.0)
                }
            }
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    /// `CStr`-style text.
    pub fn to_text(&self) -> Result<String> {
        Ok(match self {
            Value::Empty => String::new(),
            Value::Null => return Err(RuntimeError::standard(94)),
            Value::Boolean(b) => if *b { "True" } else { "False" }.into(),
            Value::Byte(n) => n.to_string(),
            Value::Integer(n) => n.to_string(),
            Value::Long(n) => n.to_string(),
            Value::LongLong(n) => n.to_string(),
            Value::Single(n) => format_float(f64::from(*n), 7),
            Value::Double(n) => format_float(*n, 15),
            Value::Currency(n) => format_currency(*n),
            Value::Decimal(d) => d.to_string(),
            Value::Date(d) => format_date(*d),
            Value::String(s) => s.to_string(),
            Value::Object(Some(o)) => return o.default_value()?.to_text(),
            Value::Object(None) => return Err(RuntimeError::object_not_set()),
            Value::Error(n) => format!("Error {n}"),
            Value::Array(_) => return Err(RuntimeError::type_mismatch()),
        })
    }

    /// Unwraps an object to its default member (e.g. `Range` -> `.Value`) for use as data.
    pub fn deref_default(self) -> Result<Value> {
        match self {
            Value::Object(Some(o)) => o.default_value(),
            v => Ok(v),
        }
    }
}

/// Parses text the way numeric coercion does: spaces, sign, `$`, thousands separators,
/// exponent, `&H`/`&O` prefixes. `None` means "not a number" (type mismatch).
pub fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let lower = t.to_ascii_lowercase();
    if let Some(hex) = lower.strip_prefix("&h") {
        return i64::from_str_radix(hex.trim_end_matches('&'), 16).ok().map(|n| n as f64);
    }
    if let Some(oct) = lower.strip_prefix("&o") {
        return i64::from_str_radix(oct.trim_end_matches('&'), 8).ok().map(|n| n as f64);
    }
    let (neg, body) =
        if let Some(r) = t.strip_prefix('(').and_then(|r| r.strip_suffix(')')) { (true, r.trim()) } else { (false, t) };
    let mut cleaned = String::with_capacity(body.len());
    for c in body.chars() {
        match c {
            ',' | '$' => {}
            c if c.is_whitespace() => {}
            _ => cleaned.push(c),
        }
    }
    let (pct, cleaned) = match cleaned.strip_suffix('%') {
        Some(r) => (true, r.to_owned()),
        None => (false, cleaned),
    };
    if cleaned.is_empty() || !cleaned.chars().any(|c| c.is_ascii_digit()) {
        return None;
    }
    if !cleaned.chars().all(|c| c.is_ascii_digit() || matches!(c, '.' | '-' | '+' | 'e' | 'E' | 'd' | 'D')) {
        return None;
    }
    let normalized = cleaned.replace(['d', 'D'], "e");
    let mut v: f64 = normalized.parse().ok()?;
    if pct {
        v /= 100.0;
    }
    Some(if neg { -v } else { v })
}

/// Shortest text with at most `digits` significant digits; scientific notation outside
/// 1E-15..1E15. ORACLE: confirm the switch-over points against the reference spreadsheet's CStr.
pub fn format_float(x: f64, digits: usize) -> String {
    if x == 0.0 {
        return "0".into();
    }
    if !x.is_finite() {
        return if x.is_nan() {
            "NaN".into()
        } else if x > 0.0 {
            "Infinity".into()
        } else {
            "-Infinity".into()
        };
    }
    let sci = format!("{:.*e}", digits.saturating_sub(1), x);
    let (mantissa, exp) = sci.split_once('e').unwrap_or((&sci, "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let mantissa = trim_zeros(mantissa);
    if exp >= digits as i32 || exp < -(digits as i32) {
        let sign = if exp < 0 { '-' } else { '+' };
        return format!("{mantissa}E{sign}{:02}", exp.abs());
    }
    // Fixed notation from the rounded mantissa.
    let neg = mantissa.starts_with('-');
    let m = mantissa.trim_start_matches('-');
    let digits_only: String = m.chars().filter(char::is_ascii_digit).collect();
    let point = 1 + exp; // position of the decimal point in digits_only
    let mut out = String::new();
    if neg {
        out.push('-');
    }
    if point <= 0 {
        out.push_str("0.");
        out.extend(std::iter::repeat_n('0', (-point) as usize));
        out.push_str(&digits_only);
    } else if point as usize >= digits_only.len() {
        out.push_str(&digits_only);
        out.extend(std::iter::repeat_n('0', point as usize - digits_only.len()));
    } else {
        out.push_str(&digits_only[..point as usize]);
        out.push('.');
        out.push_str(&digits_only[point as usize..]);
    }
    out
}

fn trim_zeros(m: &str) -> String {
    if m.contains('.') { m.trim_end_matches('0').trim_end_matches('.').to_owned() } else { m.to_owned() }
}

/// Exact text of a Currency (scaled by 10,000): trailing fraction zeros dropped (1.2346, 5, -0.5).
pub fn format_currency(n: i64) -> String {
    let sign = if n < 0 { "-" } else { "" };
    let a = n.unsigned_abs();
    let frac = format!("{:04}", a % 10_000);
    let frac = frac.trim_end_matches('0');
    if frac.is_empty() { format!("{sign}{}", a / 10_000) } else { format!("{sign}{}.{frac}", a / 10_000) }
}

/// (year, month, day) for a serial date.
pub fn serial_to_ymd(serial: f64) -> (i64, u32, u32) {
    let days = serial.floor() as i64 - UNIX_EPOCH_SERIAL as i64;
    // Civil-from-days (Howard Hinnant's algorithm).
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// Serial date for (year, month, day); months and days may overflow like DateSerial allows.
pub fn ymd_to_serial(year: i64, month: i64, day: i64) -> f64 {
    let y = year + (month - 1).div_euclid(12);
    let m = (month - 1).rem_euclid(12) + 1;
    let (y2, m2) = if m <= 2 { (y - 1, m + 9) } else { (y, m - 3) };
    let era = y2.div_euclid(400);
    let yoe = y2.rem_euclid(400);
    let doy = (153 * m2 + 2) / 5;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days = era * 146_097 + doe - 719_468;
    (days + UNIX_EPOCH_SERIAL as i64) as f64 + (day - 1) as f64
}

/// (hour, minute, second) of a serial date's time part, rounded to the second.
pub fn serial_to_hms(serial: f64) -> (u32, u32, u32) {
    let frac = (serial - serial.floor()).abs();
    let secs = (frac * 86_400.0).round() as u32 % 86_400;
    (secs / 3600, secs / 60 % 60, secs % 60)
}

/// Default date display (en-US short date + long time). ORACLE: locale.
pub fn format_date(d: f64) -> String {
    let (y, m, day) = serial_to_ymd(d);
    let (h, mi, s) = serial_to_hms(d);
    let time = || {
        let (h12, ampm) = match h {
            0 => (12, "AM"),
            1..=11 => (h, "AM"),
            12 => (12, "PM"),
            _ => (h - 12, "PM"),
        };
        format!("{h12}:{mi:02}:{s:02} {ampm}")
    };
    let has_time = d.fract() != 0.0;
    if d.trunc() == 0.0 && has_time {
        time()
    } else if has_time {
        format!("{m}/{day}/{y} {}", time())
    } else {
        format!("{m}/{day}/{y}")
    }
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;

    #[test]
    fn bankers_rounding() {
        assert_eq!(round_half_even(2.5), 2.0);
        assert_eq!(round_half_even(3.5), 4.0);
        assert_eq!(round_half_even(-2.5), -2.0);
        assert_eq!(round_half_even(2.6), 3.0);
        assert_eq!(Value::Double(32767.5).to_i16(), Err(RuntimeError::overflow()));
        assert_eq!(Value::Double(32766.5).to_i16(), Ok(32766));
    }

    #[test]
    fn numbers_from_text() {
        assert_eq!(parse_number(" 1,234.5 "), Some(1234.5));
        assert_eq!(parse_number("&HFF"), Some(255.0));
        assert_eq!(parse_number("1e3"), Some(1000.0));
        assert_eq!(parse_number("$5"), Some(5.0));
        assert_eq!(parse_number("abc"), None);
        assert_eq!(parse_number(""), None);
    }

    #[test]
    fn float_text() {
        assert_eq!(format_float(0.1 + 0.2, 15), "0.3");
        assert_eq!(format_float(1.0 / 3.0, 15), "0.333333333333333");
        assert_eq!(format_float(1234.0, 15), "1234");
        assert_eq!(format_float(-2.5, 15), "-2.5");
        assert_eq!(format_float(1e20, 15), "1E+20");
        assert_eq!(format_float(123456789012345.0, 15), "123456789012345");
    }

    #[test]
    fn dates_round_trip() {
        assert_eq!(serial_to_ymd(1.0), (1899, 12, 31));
        assert_eq!(ymd_to_serial(2026, 10, 14), 46309.0);
        assert_eq!(serial_to_ymd(46309.0), (2026, 10, 14));
        assert_eq!(ymd_to_serial(2026, 13, 1), ymd_to_serial(2027, 1, 1));
        assert_eq!(format_date(46309.5), "10/14/2026 12:00:00 PM");
        assert_eq!(format_date(0.25), "6:00:00 AM");
    }

    #[test]
    fn redim_preserve_keeps_contents() {
        let mut a = Array::from_vec(0, vec![Value::Long(1), Value::Long(2)]);
        a.redim_preserve(vec![(0, 3)]).unwrap();
        assert_eq!(a.data.len(), 4);
        assert!(matches!(a.get(&[1]).unwrap(), Value::Long(2)));
    }
}
