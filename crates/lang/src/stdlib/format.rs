//! `Format` and its relatives: named formats, user-defined numeric, date/time and string
//! formats, plus decimal rounding shared with `Round`.
//!
//! Output follows en-US regional settings ("." decimal, "," grouping, "$" currency, m/d/yyyy).
//! ORACLE: every locale-dependent piece (separators, currency placement, negative currency
//! style, long date pattern) should be confirmed against the reference spreadsheet on an en-US machine.

use super::datetime::{self, MONTHS, WEEKDAYS};
use crate::error::{Result, RuntimeError};
use crate::value::{Value, format_date, parse_number};

// ---------------------------------------------------------------------------------------------
// Decimal rounding

/// Digits of `|x|` rounded to `decimals` places: (integer digits without leading zeros, exactly
/// `decimals` fraction digits). Rounding works on the shortest decimal representation with
/// `sig` significant digits (15 for Double, 7 for Single), so 2.675 rounds as the decimal 2.675
/// rather than as its binary neighbour 2.67499999.... `half_even` picks banker's rounding
/// (`Round`); otherwise halves round away from zero (`Format`).
pub fn round_decimal(x: f64, decimals: usize, half_even: bool, sig: usize) -> (String, String) {
    let x = x.abs();
    if x == 0.0 || !x.is_finite() {
        return (String::new(), "0".repeat(decimals));
    }
    let sci = format!("{:.*e}", sig.saturating_sub(1), x);
    let (mantissa, exp) = sci.split_once('e').unwrap_or((&sci, "0"));
    let exp: i64 = exp.parse().unwrap_or(0);
    let mut digits: Vec<u8> = mantissa.bytes().filter(u8::is_ascii_digit).map(|b| b - b'0').collect();
    // Value = 0.d1d2d3... * 10^(exp + 1); `point` = digits before the decimal point.
    let mut point = exp + 1;
    if point < 0 {
        let pad = (-point) as usize;
        let mut v = vec![0u8; pad];
        v.extend_from_slice(&digits);
        digits = v;
        point = 0;
    }
    let point = point as usize;
    let keep = point + decimals;
    if digits.len() < keep {
        digits.resize(keep, 0);
    }
    let mut kept: Vec<u8> = digits[..keep].to_vec();
    let rest = &digits[keep..];
    let round_up = match rest.first() {
        None => false,
        Some(&first) => {
            if first > 5 {
                true
            } else if first < 5 {
                false
            } else if !half_even || rest[1..].iter().any(|&d| d != 0) {
                true
            } else {
                kept.last().is_some_and(|d| d % 2 == 1)
            }
        }
    };
    let mut point = point;
    if round_up {
        let mut i = kept.len();
        loop {
            if i == 0 {
                kept.insert(0, 1);
                point += 1;
                break;
            }
            i -= 1;
            if kept[i] == 9 {
                kept[i] = 0;
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    let int: String = kept[..point].iter().map(|d| char::from(b'0' + d)).collect();
    let int = int.trim_start_matches('0').to_owned();
    let frac: String = kept[point..].iter().map(|d| char::from(b'0' + d)).collect();
    (int, frac)
}

/// `Round(x, decimals)`: scale by 10^decimals in binary floating point, round half to even,
/// scale back. Observed on the real host: Round(2.345, 2) = 2.35 (2.345 * 100 is
/// 234.50000000000003), Round(2.355, 2) = 2.36 (235.5 exactly, to even), Round(2.5) = 2,
/// Round(-2.5) = -2. Neither pure decimal nor exact-binary rounding fits all four; this does.
pub fn round_f64(x: f64, decimals: usize) -> f64 {
    let p = 10f64.powi(decimals.min(300) as i32);
    let scaled = x * p;
    if !scaled.is_finite() || scaled.abs() >= 4.5e15 {
        return x;
    }
    crate::value::round_half_even(scaled) / p
}

fn group(int: &str) -> String {
    let mut out = String::with_capacity(int.len() + int.len() / 3);
    let n = int.len();
    for (i, c) in int.chars().enumerate() {
        if i > 0 && (n - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(c);
    }
    out
}

// ---------------------------------------------------------------------------------------------
// FormatNumber / FormatCurrency / FormatPercent / FormatDateTime

/// Options for the FormatNumber family. Tristates: `None` = use default.
#[derive(Debug, Clone, Copy)]
pub struct NumberStyle {
    pub decimals: usize,
    pub leading_digit: bool,
    pub parens: bool,
    pub group: bool,
}

pub fn format_number_style(x: f64, style: NumberStyle, prefix: &str, suffix: &str) -> String {
    let (int, frac) = round_decimal(x, style.decimals, false, 15);
    let zero = int.is_empty() && frac.bytes().all(|b| b == b'0');
    let mut body = String::new();
    body.push_str(prefix);
    if int.is_empty() {
        if style.leading_digit {
            body.push('0');
        }
    } else if style.group {
        body.push_str(&group(&int));
    } else {
        body.push_str(&int);
    }
    if style.decimals > 0 {
        body.push('.');
        body.push_str(&frac);
    }
    body.push_str(suffix);
    if x < 0.0 && !zero { if style.parens { format!("({body})") } else { format!("-{body}") } } else { body }
}

/// `FormatDateTime` named formats: 0 general, 1 long date, 2 short date, 3 long time, 4 short time.
pub fn format_date_time(d: f64, named: i32) -> Result<String> {
    Ok(match named {
        0 => format_date(d),
        1 => format_date_pattern(d, "dddd, mmmm d, yyyy", 1, 1),
        2 => format_date_pattern(d, "m/d/yyyy", 1, 1),
        3 => format_date_pattern(d, "h:mm:ss AM/PM", 1, 1),
        4 => format_date_pattern(d, "hh:mm", 1, 1),
        _ => return Err(RuntimeError::invalid_call()),
    })
}

// ---------------------------------------------------------------------------------------------
// Format

/// `Format(expr, fmt, firstDayOfWeek, firstWeekOfYear)`. Returns Null for a Null expression
/// (unless a numeric format has a fourth, Null section).
pub fn format(v: &Value, fmt: &str, first_day: u32, first_week: u32, current_year: i64) -> Result<Value> {
    let v = v.clone().deref_default()?;
    if let Value::Array(_) = v {
        return Err(RuntimeError::type_mismatch());
    }
    if fmt.is_empty() {
        return Ok(match v {
            Value::Null => Value::Null,
            other => Value::str(other.to_text()?),
        });
    }
    if let Some(r) = named_format(&v, fmt, current_year)? {
        return Ok(r);
    }
    let kind = classify_format(fmt);
    let sections = split_sections(fmt);
    let text = |s: String| Ok(Value::str(s));
    match &v {
        Value::Null => {
            // ORACLE: Format(Null, fmt) - a fourth numeric section is used, otherwise Null.
            if kind == FKind::Numeric && sections.len() >= 4 {
                return text(format_section(&sections[3], 0.0, false));
            }
            Ok(Value::Null)
        }
        Value::String(s) => match kind {
            FKind::Date => match datetime::parse_date(s, current_year).or_else(|| parse_number(s)) {
                Some(d) if datetime::in_range(d) => text(format_date_pattern(d, &sections[0], first_day, first_week)),
                _ => text(s.to_string()),
            },
            FKind::Numeric => match parse_number(s) {
                Some(x) => text(format_numeric(x, &sections)),
                None => text(s.to_string()),
            },
            FKind::Text | FKind::Literal => text(format_string(s, &sections)),
        },
        Value::Date(d) => match kind {
            FKind::Numeric => text(format_numeric(*d, &sections)),
            FKind::Text => text(format_string(&v.to_text()?, &sections)),
            FKind::Date | FKind::Literal => text(format_date_pattern(*d, &sections[0], first_day, first_week)),
        },
        Value::Error(_) => text(v.to_text()?),
        other => {
            // ORACLE: Empty with a numeric format - we format it as 0.
            let x = other.to_f64()?;
            match kind {
                FKind::Date => {
                    if !datetime::in_range(x) {
                        return Err(RuntimeError::overflow());
                    }
                    text(format_date_pattern(x, &sections[0], first_day, first_week))
                }
                FKind::Text => text(format_string(&other.to_text()?, &sections)),
                FKind::Numeric | FKind::Literal => text(format_numeric(x, &sections)),
            }
        }
    }
}

fn named_format(v: &Value, fmt: &str, current_year: i64) -> Result<Option<Value>> {
    let name = fmt.to_ascii_lowercase();
    let numeric = |v: &Value| -> Result<Option<f64>> {
        Ok(match v {
            Value::String(s) => parse_number(s),
            Value::Null => None,
            other => Some(other.to_f64()?),
        })
    };
    let num_pattern = match name.as_str() {
        "fixed" => Some("0.00"),
        "standard" => Some("#,##0.00"),
        "percent" => Some("0.00%"),
        "scientific" => Some("0.00E+00"),
        _ => None,
    };
    if let Some(p) = num_pattern {
        return Ok(Some(match numeric(v)? {
            Some(x) => Value::str(format_numeric(x, &split_sections(p))),
            None => passthrough(v)?,
        }));
    }
    match name.as_str() {
        "general number" => Ok(Some(match numeric(v)? {
            Some(x) => Value::str(Value::Double(x).to_text()?),
            None => passthrough(v)?,
        })),
        "currency" => Ok(Some(match numeric(v)? {
            Some(x) => Value::str(format_number_style(x, CURRENCY_STYLE, "$", "")),
            None => passthrough(v)?,
        })),
        "yes/no" | "true/false" | "on/off" => {
            let (t, f) = match name.as_str() {
                "yes/no" => ("Yes", "No"),
                "true/false" => ("True", "False"),
                _ => ("On", "Off"),
            };
            Ok(Some(match numeric(v)? {
                Some(x) => Value::str(if x != 0.0 { t } else { f }),
                None => passthrough(v)?,
            }))
        }
        "general date" | "long date" | "medium date" | "short date" | "long time" | "medium time" | "short time" => {
            let d = match v {
                Value::Null => return Ok(Some(Value::Null)),
                Value::String(s) => match datetime::parse_date(s, current_year).or_else(|| parse_number(s)) {
                    Some(d) => d,
                    None => return Ok(Some(v.clone())),
                },
                other => other.to_f64()?,
            };
            if !datetime::in_range(d) {
                return Err(RuntimeError::overflow());
            }
            let pattern = match name.as_str() {
                "general date" => return Ok(Some(Value::str(format_date(d)))),
                "long date" => "dddd, mmmm d, yyyy",
                "medium date" => "dd-mmm-yy",
                "short date" => "m/d/yyyy",
                "long time" => "h:mm:ss AM/PM",
                "medium time" => "hh:mm AM/PM",
                _ => "hh:mm",
            };
            Ok(Some(Value::str(format_date_pattern(d, pattern, 1, 1))))
        }
        _ => Ok(None),
    }
}

/// ORACLE: negative currency in parentheses, as in classic en-US settings.
pub const CURRENCY_STYLE: NumberStyle = NumberStyle { decimals: 2, leading_digit: true, parens: true, group: true };

fn passthrough(v: &Value) -> Result<Value> {
    Ok(match v {
        Value::Null => Value::Null,
        other => Value::str(other.to_text()?),
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FKind {
    Numeric,
    Date,
    Text,
    Literal,
}

/// Visits the unquoted, unescaped characters of a format.
fn plain_chars(fmt: &str) -> Vec<char> {
    let mut out = Vec::new();
    let mut chars = fmt.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                chars.next();
            }
            other => out.push(other),
        }
    }
    out
}

fn classify_format(fmt: &str) -> FKind {
    let plain = plain_chars(fmt);
    let lower: String = plain.iter().collect::<String>().to_ascii_lowercase();
    if plain.iter().any(|c| matches!(c, '0' | '#')) {
        return FKind::Numeric;
    }
    if lower.chars().any(|c| matches!(c, 'd' | 'm' | 'y' | 'h' | 'n' | 's' | 'w' | 'q' | 'c'))
        || lower.contains("am/pm")
        || lower.contains("a/p")
        || lower.contains("ttttt")
    {
        return FKind::Date;
    }
    if plain.iter().any(|c| matches!(c, '@' | '&' | '<' | '>' | '!')) {
        return FKind::Text;
    }
    FKind::Literal
}

/// Splits a format at unquoted `;`.
fn split_sections(fmt: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut chars = fmt.chars();
    while let Some(c) = chars.next() {
        let cur = out.last_mut();
        let Some(cur) = cur else { break };
        match c {
            '"' => {
                cur.push('"');
                for q in chars.by_ref() {
                    cur.push(q);
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                cur.push('\\');
                if let Some(n) = chars.next() {
                    cur.push(n);
                }
            }
            ';' => out.push(String::new()),
            other => cur.push(other),
        }
    }
    out
}

// ---------------------------------------------------------------------------------------------
// Numeric formats

#[derive(Debug, Clone, PartialEq)]
enum NTok {
    Digit(bool),
    Point,
    Comma,
    Percent,
    Exp { plus: bool, upper: bool },
    Lit(String),
}

fn parse_numeric(sec: &str) -> Vec<NTok> {
    let mut out = Vec::new();
    let chars: Vec<char> = sec.chars().collect();
    let mut i = 0;
    let mut seen_point = false;
    while i < chars.len() {
        let c = chars[i];
        i += 1;
        match c {
            '"' => {
                let mut s = String::new();
                while i < chars.len() && chars[i] != '"' {
                    s.push(chars[i]);
                    i += 1;
                }
                i += 1;
                out.push(NTok::Lit(s));
            }
            '\\' => {
                if let Some(&n) = chars.get(i) {
                    out.push(NTok::Lit(n.to_string()));
                    i += 1;
                }
            }
            '0' => out.push(NTok::Digit(true)),
            '#' => out.push(NTok::Digit(false)),
            '.' if !seen_point => {
                seen_point = true;
                out.push(NTok::Point);
            }
            ',' => out.push(NTok::Comma),
            '%' => out.push(NTok::Percent),
            'E' | 'e' if matches!(chars.get(i), Some('+' | '-')) => {
                out.push(NTok::Exp { plus: chars.get(i) == Some(&'+'), upper: c == 'E' });
                i += 1;
            }
            other => out.push(NTok::Lit(other.to_string())),
        }
    }
    out
}

/// Picks the section for `x` and formats it.
fn format_numeric(x: f64, sections: &[String]) -> String {
    let nonempty = |i: usize| sections.get(i).filter(|s| !s.is_empty());
    if x < 0.0 {
        if let Some(neg) = nonempty(1) {
            return format_section(neg, -x, false);
        }
        if sections.len() >= 2 && sections[1].is_empty() {
            // An empty negative section uses the positive one, with the sign.
            return format_section(&sections[0], x, true);
        }
    }
    if x == 0.0
        && let Some(zero) = nonempty(2)
    {
        return format_section(zero, 0.0, false);
    }
    format_section(sections.first().map_or("", String::as_str), x, true)
}

/// Formats `x` with one section. With `signed`, a negative value gets a leading "-".
fn format_section(sec: &str, x: f64, signed: bool) -> String {
    let toks = parse_numeric(sec);
    let exp_at = toks.iter().position(|t| matches!(t, NTok::Exp { .. }));
    let mantissa_end = exp_at.unwrap_or(toks.len());
    let point_at = toks[..mantissa_end].iter().position(|t| *t == NTok::Point);
    let int_end = point_at.unwrap_or(mantissa_end);
    let int_digits: Vec<usize> = (0..int_end).filter(|&i| matches!(toks[i], NTok::Digit(_))).collect();
    let dec_digits: Vec<usize> = (int_end..mantissa_end).filter(|&i| matches!(toks[i], NTok::Digit(_))).collect();
    let first_int = int_digits.first().copied();
    let last_int = int_digits.last().copied();
    // Commas between integer placeholders turn on grouping; commas right after the last
    // integer placeholder (before the point) scale by 1000 each.
    let mut grouping = false;
    let mut scale = 0i32;
    for (i, t) in toks[..int_end].iter().enumerate() {
        if *t == NTok::Comma {
            match (first_int, last_int) {
                (Some(f), Some(l)) if i > f && i < l => grouping = true,
                (Some(_), Some(l)) if i > l => scale += 1,
                _ => {}
            }
        }
    }
    let percents = toks.iter().filter(|t| **t == NTok::Percent).count() as i32;
    let mut v = x.abs() * 100f64.powi(percents) / 1000f64.powi(scale);
    let negative = x < 0.0;
    let min_dec = dec_digits.iter().rposition(|&i| toks[i] == NTok::Digit(true)).map_or(0, |p| p + 1);
    let max_dec = dec_digits.len();

    let mut exp_text = String::new();
    if let Some(e) = exp_at {
        let n_int = int_digits.len().max(1);
        let mut exponent = 0i64;
        if v != 0.0 {
            exponent = v.log10().floor() as i64 - (n_int as i64 - 1);
            v /= 10f64.powi(exponent as i32);
            // Rounding may carry into a new digit (9.99 -> 10.0).
            let (int, _) = round_decimal(v, max_dec, false, 15);
            if int.len() > n_int {
                exponent += 1;
                v /= 10.0;
            }
        }
        if let NTok::Exp { plus, upper } = toks[e] {
            exp_text.push(if upper { 'E' } else { 'e' });
            if exponent < 0 {
                exp_text.push('-');
            } else if plus {
                exp_text.push('+');
            }
            let min_exp = toks[e + 1..].iter().filter(|t| **t == NTok::Digit(true)).count();
            exp_text.push_str(&format!("{:0width$}", exponent.abs(), width = min_exp));
        }
    }

    let (int, frac) = round_decimal(v, max_dec, false, 15);
    let mut frac = frac;
    while frac.len() > min_dec && frac.ends_with('0') {
        frac.pop();
    }
    let all_zero = int.is_empty() && frac.bytes().all(|b| b == b'0');
    let int_chars: Vec<char> = int.chars().collect();

    let mut out = String::new();
    // ORACLE: a negative value that rounds to zero is shown without the sign.
    if signed && negative && !all_zero {
        out.push('-');
    }
    let n_int = int_digits.len();
    let mut int_slot = 0usize;
    let mut dec_slot = 0usize;
    let grouped = {
        let min_int = int_digits.iter().filter(|&&i| toks[i] == NTok::Digit(true)).count();
        let mut s = int.clone();
        while s.len() < min_int {
            s.insert(0, '0');
        }
        group(&s)
    };
    for (i, t) in toks[..mantissa_end].iter().enumerate() {
        match t {
            NTok::Digit(zero) if i < int_end => {
                if grouping {
                    if int_slot == 0 {
                        out.push_str(&grouped);
                    }
                } else {
                    // Right-aligned: slot k (from the left) shows digit L - (n - k); the first
                    // slot also takes every extra leading digit.
                    let len = int_chars.len();
                    let from_right = n_int - int_slot; // 1-based position from the right
                    if int_slot == 0 && len > n_int {
                        out.extend(&int_chars[..len - n_int + 1]);
                    } else if from_right <= len {
                        out.push(int_chars[len - from_right]);
                    } else if *zero {
                        out.push('0');
                    }
                }
                int_slot += 1;
            }
            NTok::Digit(_) => {
                if let Some(c) = frac.chars().nth(dec_slot) {
                    out.push(c);
                }
                dec_slot += 1;
            }
            NTok::Point => {
                if n_int == 0 {
                    out.push_str(&int);
                }
                out.push('.');
            }
            NTok::Comma => {
                if !(first_int.is_some() && i < int_end) {
                    out.push(',');
                }
            }
            NTok::Percent => out.push('%'),
            NTok::Lit(s) => out.push_str(s),
            NTok::Exp { .. } => {}
        }
    }
    out.push_str(&exp_text);
    for t in toks.iter().skip(mantissa_end + 1) {
        match t {
            NTok::Lit(s) => out.push_str(s),
            NTok::Percent => out.push('%'),
            _ => {}
        }
    }
    out
}

// ---------------------------------------------------------------------------------------------
// Date formats

#[derive(Debug, Clone, PartialEq)]
enum DTok {
    Lit(String),
    Day(usize),
    Weekday,
    Week,
    Month(usize),
    Minute(usize),
    Quarter,
    DayOfYear,
    Year2,
    Year4,
    Hour(usize),
    Second(usize),
    AmPm(&'static str, &'static str),
    General,
    LongTime,
}

fn parse_date_format(fmt: &str) -> Vec<DTok> {
    let chars: Vec<char> = fmt.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    let rest = |i: usize| -> String { chars[i..].iter().collect::<String>() };
    while i < chars.len() {
        let c = chars[i];
        let lc = c.to_ascii_lowercase();
        let run = chars[i..].iter().take_while(|x| x.to_ascii_lowercase() == lc).count();
        match lc {
            '"' => {
                let mut s = String::new();
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    s.push(chars[i]);
                    i += 1;
                }
                i += 1;
                out.push(DTok::Lit(s));
                continue;
            }
            '\\' => {
                if let Some(&n) = chars.get(i + 1) {
                    out.push(DTok::Lit(n.to_string()));
                }
                i += 2;
                continue;
            }
            'a' => {
                let r = rest(i);
                let lower = r.to_ascii_lowercase();
                if lower.starts_with("am/pm") {
                    out.push(if r.starts_with("AM") { DTok::AmPm("AM", "PM") } else { DTok::AmPm("am", "pm") });
                    i += 5;
                } else if lower.starts_with("a/p") {
                    out.push(if c == 'A' { DTok::AmPm("A", "P") } else { DTok::AmPm("a", "p") });
                    i += 3;
                } else if lower.starts_with("ampm") {
                    out.push(DTok::AmPm("AM", "PM"));
                    i += 4;
                } else {
                    out.push(DTok::Lit(c.to_string()));
                    i += 1;
                }
                continue;
            }
            _ => {}
        }
        let (tok, used) = match lc {
            'd' => (DTok::Day(run.min(6)), run),
            'w' => (if run == 1 { DTok::Weekday } else { DTok::Week }, run.min(2)),
            'm' => (DTok::Month(run.min(4)), run),
            'q' => (DTok::Quarter, 1),
            'y' => match run {
                1 => (DTok::DayOfYear, 1),
                2 => (DTok::Year2, 2),
                3 => (DTok::Year2, 2),
                _ => (DTok::Year4, run),
            },
            'h' => (DTok::Hour(run.min(2)), run.min(2)),
            'n' => (DTok::Minute(run.min(2)), run.min(2)),
            's' => (DTok::Second(run.min(2)), run.min(2)),
            't' if run >= 5 => (DTok::LongTime, run),
            'c' => (DTok::General, 1),
            _ => (DTok::Lit(c.to_string()), 1),
        };
        out.push(tok);
        i += used;
    }
    // `m`/`mm` right after an hour or right before a second means minutes.
    let significant: Vec<usize> = (0..out.len()).filter(|&k| !matches!(out[k], DTok::Lit(_))).collect();
    for (pos, &k) in significant.iter().enumerate() {
        if let DTok::Month(n) = out[k]
            && n <= 2
        {
            let prev_hour = pos > 0 && matches!(out[significant[pos - 1]], DTok::Hour(_));
            let next_second = significant.get(pos + 1).is_some_and(|&j| matches!(out[j], DTok::Second(_)));
            if prev_hour || next_second {
                out[k] = DTok::Minute(n);
            }
        }
    }
    out
}

/// Formats a serial date with a user-defined date/time pattern.
pub fn format_date_pattern(d: f64, fmt: &str, first_day: u32, first_week: u32) -> String {
    let toks = parse_date_format(fmt);
    let twelve = toks.iter().any(|t| matches!(t, DTok::AmPm(..)));
    let (y, m, day) = datetime::ymd(d);
    let (h, mi, s) = datetime::hms(d);
    let mut out = String::new();
    for t in &toks {
        match t {
            DTok::Lit(s) => out.push_str(s),
            DTok::Day(1) => out.push_str(&day.to_string()),
            DTok::Day(2) => out.push_str(&format!("{day:02}")),
            DTok::Day(3) => out.push_str(&WEEKDAYS[datetime::weekday(d, 1) as usize - 1][..3]),
            DTok::Day(4) => out.push_str(WEEKDAYS[datetime::weekday(d, 1) as usize - 1]),
            DTok::Day(5) => out.push_str(&format!("{m}/{day}/{y}")),
            DTok::Day(_) => out.push_str(&format_date_pattern(d, "dddd, mmmm d, yyyy", first_day, first_week)),
            DTok::Weekday => out.push_str(&datetime::weekday(d, first_day).to_string()),
            DTok::Week => out.push_str(&datetime::week_of_year(d, first_day, first_week).to_string()),
            DTok::Month(1) => out.push_str(&m.to_string()),
            DTok::Month(2) => out.push_str(&format!("{m:02}")),
            DTok::Month(3) => out.push_str(&MONTHS[m as usize - 1][..3]),
            DTok::Month(_) => out.push_str(MONTHS[m as usize - 1]),
            DTok::Minute(1) => out.push_str(&mi.to_string()),
            DTok::Minute(_) => out.push_str(&format!("{mi:02}")),
            DTok::Quarter => out.push_str(&((m - 1) / 3 + 1).to_string()),
            DTok::DayOfYear => {
                let jan1 = crate::value::ymd_to_serial(y, 1, 1);
                out.push_str(&((d.trunc() - jan1) as i64 + 1).to_string());
            }
            DTok::Year2 => out.push_str(&format!("{:02}", y.rem_euclid(100))),
            DTok::Year4 => out.push_str(&y.to_string()),
            DTok::Hour(n) => {
                let hh = if twelve {
                    match h % 12 {
                        0 => 12,
                        x => x,
                    }
                } else {
                    h
                };
                if *n == 1 {
                    out.push_str(&hh.to_string());
                } else {
                    out.push_str(&format!("{hh:02}"));
                }
            }
            DTok::Second(1) => out.push_str(&s.to_string()),
            DTok::Second(_) => out.push_str(&format!("{s:02}")),
            DTok::AmPm(am, pm) => out.push_str(if h < 12 { am } else { pm }),
            DTok::General => out.push_str(&format_date(d)),
            DTok::LongTime => out.push_str(&format_date_pattern(d, "h:mm:ss AM/PM", first_day, first_week)),
        }
    }
    out
}

// ---------------------------------------------------------------------------------------------
// String formats

fn format_string(s: &str, sections: &[String]) -> String {
    let sec = if s.is_empty() && sections.len() >= 2 { &sections[1] } else { &sections[0] };
    #[derive(PartialEq)]
    enum STok {
        At,
        Amp,
        Lit(String),
    }
    let mut toks = Vec::new();
    let mut upper = None;
    let mut left_to_right = false;
    let chars: Vec<char> = sec.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        i += 1;
        match c {
            '@' => toks.push(STok::At),
            '&' => toks.push(STok::Amp),
            '<' => upper = Some(false),
            '>' => upper = Some(true),
            '!' => left_to_right = true,
            '"' => {
                let mut lit = String::new();
                while i < chars.len() && chars[i] != '"' {
                    lit.push(chars[i]);
                    i += 1;
                }
                i += 1;
                toks.push(STok::Lit(lit));
            }
            '\\' => {
                if let Some(&n) = chars.get(i) {
                    toks.push(STok::Lit(n.to_string()));
                    i += 1;
                }
            }
            other => toks.push(STok::Lit(other.to_string())),
        }
    }
    let s = match upper {
        Some(true) => s.to_uppercase(),
        Some(false) => s.to_lowercase(),
        None => s.to_owned(),
    };
    let src: Vec<char> = s.chars().collect();
    let slots = toks.iter().filter(|t| matches!(t, STok::At | STok::Amp)).count();
    if slots == 0 {
        let mut out: String = toks.iter().map(|t| if let STok::Lit(l) = t { l.as_str() } else { "" }).collect();
        out.push_str(&s);
        return out;
    }
    // Assign characters to slots: right to left by default, left to right with `!`.
    let mut assigned: Vec<String> = vec![String::new(); slots];
    if left_to_right {
        for (k, slot) in assigned.iter_mut().enumerate() {
            if let Some(c) = src.get(k) {
                slot.push(*c);
            }
        }
        if src.len() > slots
            && let Some(last) = assigned.last_mut()
        {
            last.extend(&src[slots..]);
        }
    } else {
        let extra = src.len().saturating_sub(slots);
        for (k, slot) in assigned.iter_mut().enumerate() {
            let from_right = slots - k;
            if k == 0 && extra > 0 {
                slot.extend(&src[..=extra]);
            } else if from_right <= src.len() {
                slot.push(src[src.len() - from_right]);
            }
        }
    }
    let mut out = String::new();
    let mut k = 0;
    for t in &toks {
        match t {
            STok::Lit(l) => out.push_str(l),
            STok::At | STok::Amp => {
                if assigned[k].is_empty() {
                    if *t == STok::At {
                        out.push(' ');
                    }
                } else {
                    out.push_str(&assigned[k]);
                }
                k += 1;
            }
        }
    }
    out
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;

    fn f(v: Value, fmt: &str) -> String {
        match format(&v, fmt, 1, 1, 2026).unwrap() {
            Value::String(s) => s.to_string(),
            other => format!("{other:?}"),
        }
    }

    #[test]
    fn rounding() {
        assert_eq!(round_decimal(2.675, 2, false, 15), ("2".into(), "68".into()));
        assert_eq!(round_decimal(2.345, 2, true, 15), ("2".into(), "34".into()));
        assert_eq!(round_decimal(0.5, 0, true, 15), ("".into(), "".into()));
        assert_eq!(round_decimal(1.5, 0, true, 15), ("2".into(), "".into()));
        assert_eq!(round_decimal(9.995, 2, false, 15), ("10".into(), "00".into()));
        assert_eq!(round_decimal(0.000123, 5, false, 15), ("".into(), "00012".into()));
        assert_eq!(round_f64(2.345, 2), 2.35);
        assert_eq!(round_f64(2.355, 2), 2.36);
        assert_eq!(round_f64(0.125, 2), 0.12);
        assert_eq!(round_f64(-2.5, 0), -2.0);
        assert_eq!(round_f64(1234.5678, 2), 1234.57);
    }

    #[test]
    fn numeric_formats() {
        assert_eq!(f(Value::Double(1234.5), "#,##0.00"), "1,234.50");
        assert_eq!(f(Value::Double(0.256), "0.0%"), "25.6%");
        assert_eq!(f(Value::Double(1234.5), "0"), "1235");
        assert_eq!(f(Value::Double(2.5), "0"), "3");
        assert_eq!(f(Value::Double(0.5), "0.00"), "0.50");
        assert_eq!(f(Value::Double(1234567.891), "#,##0"), "1,234,568");
        assert_eq!(f(Value::Double(0.25), "0%"), "25%");
        assert_eq!(f(Value::Double(1234.5), "$#,##0.00"), "$1,234.50");
        assert_eq!(f(Value::Double(-1234.5), "$#,##0.00"), "-$1,234.50");
        assert_eq!(f(Value::Double(-5.0), "0;(0)"), "(5)");
        assert_eq!(f(Value::Double(0.0), "0;(0);\"zero\""), "zero");
        assert_eq!(f(Value::Double(5.0), "#.##"), "5.");
        assert_eq!(f(Value::Double(0.5), "#.##"), ".5");
        assert_eq!(f(Value::Double(1.5), "0.0#"), "1.5");
        assert_eq!(f(Value::Double(1.555), "0.0#"), "1.56");
        assert_eq!(f(Value::Long(5), "000"), "005");
        assert_eq!(f(Value::Long(123456789), "000-00-0000"), "123-45-6789");
        assert_eq!(f(Value::Double(1234.5), "0.00E+00"), "1.23E+03");
        assert_eq!(f(Value::Double(0.00012), "0.0E-00"), "1.2E-04");
        assert_eq!(f(Value::Double(1234567.0), "#,##0,"), "1,235");
        assert_eq!(f(Value::str("1234.5"), "#,##0.00"), "1,234.50");
        assert_eq!(f(Value::str("abc"), "0.00"), "abc");
        assert_eq!(f(Value::Double(1.005), "0.00"), "1.01");
        assert_eq!(f(Value::Double(-0.001), "0.00"), "0.00");
        assert_eq!(f(Value::Boolean(true), "0"), "-1");
    }

    #[test]
    fn named_formats() {
        assert_eq!(f(Value::Double(1234.5), "Currency"), "$1,234.50");
        assert_eq!(f(Value::Double(-1234.5), "Currency"), "($1,234.50)");
        assert_eq!(f(Value::Double(1234.5), "Fixed"), "1234.50");
        assert_eq!(f(Value::Double(1234.5), "Standard"), "1,234.50");
        assert_eq!(f(Value::Double(0.256), "Percent"), "25.60%");
        assert_eq!(f(Value::Double(1234.5), "General Number"), "1234.5");
        assert_eq!(f(Value::Double(1234.5), "Scientific"), "1.23E+03");
        assert_eq!(f(Value::Long(1), "Yes/No"), "Yes");
        assert_eq!(f(Value::Long(0), "True/False"), "False");
        assert_eq!(f(Value::Long(2), "On/Off"), "On");
        let dt = Value::Date(46309.75);
        assert_eq!(f(dt.clone(), "Short Date"), "10/14/2026");
        assert_eq!(f(dt.clone(), "Long Date"), "Wednesday, October 14, 2026");
        assert_eq!(f(dt.clone(), "Medium Date"), "14-Oct-26");
        assert_eq!(f(dt.clone(), "Short Time"), "18:00");
        assert_eq!(f(dt.clone(), "Long Time"), "6:00:00 PM");
        assert_eq!(f(dt.clone(), "Medium Time"), "06:00 PM");
        assert_eq!(f(dt, "General Date"), "10/14/2026 6:00:00 PM");
    }

    #[test]
    fn date_formats() {
        let dt = Value::Date(46309.0 + (9.0 * 3600.0 + 5.0 * 60.0 + 7.0) / 86400.0);
        assert_eq!(f(dt.clone(), "yyyy-mm-dd"), "2026-10-14");
        assert_eq!(f(dt.clone(), "m/d/yy"), "10/14/26");
        assert_eq!(f(dt.clone(), "mmm d, yyyy"), "Oct 14, 2026");
        assert_eq!(f(dt.clone(), "mmmm"), "October");
        assert_eq!(f(dt.clone(), "ddd"), "Wed");
        assert_eq!(f(dt.clone(), "dddd"), "Wednesday");
        assert_eq!(f(dt.clone(), "hh:mm:ss"), "09:05:07");
        assert_eq!(f(dt.clone(), "h:mm AM/PM"), "9:05 AM");
        assert_eq!(f(dt.clone(), "h:mm am/pm"), "9:05 am");
        assert_eq!(f(dt.clone(), "hh:nn"), "09:05");
        assert_eq!(f(dt.clone(), "q"), "4");
        assert_eq!(f(dt.clone(), "y"), "287");
        assert_eq!(f(dt.clone(), "w"), "4");
        assert_eq!(f(dt.clone(), "\"Day\" d"), "Day 14");
        assert_eq!(f(dt, "dd\\/mm"), "14/10");
        assert_eq!(f(Value::str("2026-10-14"), "dd mmm yyyy"), "14 Oct 2026");
    }

    #[test]
    fn string_formats() {
        assert_eq!(f(Value::str("hello"), ">"), "HELLO");
        assert_eq!(f(Value::str("HeLLo"), "<"), "hello");
        assert_eq!(f(Value::str("abc"), "@@@@@"), "  abc");
        assert_eq!(f(Value::str("abc"), "!@@@@@"), "abc  ");
        assert_eq!(f(Value::str("5551234"), "@@@-@@@@"), "555-1234");
        assert_eq!(f(Value::str("abc"), "&&&&&"), "abc");
    }

    #[test]
    fn number_styles() {
        let st = NumberStyle { decimals: 2, leading_digit: true, parens: false, group: true };
        assert_eq!(format_number_style(1234.567, st, "", ""), "1,234.57");
        assert_eq!(format_number_style(-0.5, st, "", ""), "-0.50");
        let st2 = NumberStyle { decimals: 0, leading_digit: false, parens: true, group: false };
        assert_eq!(format_number_style(-1234.5, st2, "", ""), "(1235)");
        assert_eq!(format_date_time(46309.75, 4).unwrap(), "18:00");
        assert_eq!(format_date_time(46309.75, 1).unwrap(), "Wednesday, October 14, 2026");
    }
}
