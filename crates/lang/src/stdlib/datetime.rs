//! Date helpers: serial dates, calendar parts, week numbering and parsing of date text.
//!
//! A serial date counts days from 1899-12-30. Before that origin the integer part is the
//! (negative) day and the fraction is still the time of day *added in magnitude*: -1.25 is
//! 1899-12-29 06:00. [`split`] and [`join`] convert between serials and (day, time) pairs,
//! and [`linear`] gives a monotonic day count for arithmetic.

use crate::error::{Result, RuntimeError};
use crate::value::{Value, parse_number, serial_to_ymd, ymd_to_serial};

/// 100-01-01, the first representable date.
pub const MIN_DATE: f64 = -657_434.0;
/// The day after 9999-12-31.
pub const MAX_DATE_EXCLUSIVE: f64 = 2_958_466.0;

pub const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
pub const WEEKDAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/// (day, time fraction in [0, 1)).
pub fn split(d: f64) -> (i64, f64) {
    let day = d.trunc();
    (day as i64, (d - day).abs())
}

/// Inverse of [`split`].
pub fn join(day: i64, frac: f64) -> f64 {
    if day < 0 { day as f64 - frac } else { day as f64 + frac }
}

/// Monotonic day count (time increases with the value even before the origin).
pub fn linear(d: f64) -> f64 {
    let (day, frac) = split(d);
    day as f64 + frac
}

/// Serial for a monotonic day count.
pub fn from_linear(l: f64) -> f64 {
    let day = l.floor();
    join(day as i64, l - day)
}

pub fn in_range(d: f64) -> bool {
    (MIN_DATE..MAX_DATE_EXCLUSIVE).contains(&d)
}

pub fn ymd(d: f64) -> (i64, u32, u32) {
    serial_to_ymd(d.trunc())
}

/// (hour, minute, second), rounded to the nearest second.
pub fn hms(d: f64) -> (u32, u32, u32) {
    let (_, frac) = split(d);
    let secs = ((frac * 86_400.0).round() as u32).min(86_399);
    (secs / 3600, secs / 60 % 60, secs % 60)
}

pub fn is_leap(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

pub fn days_in_month(y: i64, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ => {
            if is_leap(y) {
                29
            } else {
                28
            }
        }
    }
}

/// Day of week of a serial date, 1..=7, counting from `first_day` (1 = Sunday).
pub fn weekday(d: f64, first_day: u32) -> u32 {
    let day = d.trunc() as i64;
    // Serial 1 (1899-12-31) was a Sunday.
    let sunday_based = (day - 1).rem_euclid(7) as u32 + 1;
    let first = if first_day == 0 { 1 } else { first_day };
    (sunday_based + 7 - first) % 7 + 1
}

/// First day (serial) of the week holding `day`.
pub fn week_start(day: i64, first_day: u32) -> i64 {
    day - i64::from(weekday(day as f64, first_day)) + 1
}

/// Week number within the year (`ww`), for `first_week` 1 = week of Jan 1, 2 = first week with
/// four days, 3 = first full week. ORACLE: dates before week 1 count as the last week of the
/// previous year; the language has known quirks near year ends.
pub fn week_of_year(d: f64, first_day: u32, first_week: u32) -> i64 {
    let day = d.trunc() as i64;
    let (y, _, _) = ymd(d);
    let week1 = |year: i64| {
        let jan1 = ymd_to_serial(year, 1, 1) as i64;
        let start = week_start(jan1, first_day);
        match first_week {
            2 => {
                if jan1 - start <= 3 {
                    start
                } else {
                    start + 7
                }
            }
            3 => {
                if start == jan1 {
                    start
                } else {
                    start + 7
                }
            }
            _ => start,
        }
    };
    let mut w1 = week1(y);
    if day < w1 {
        w1 = week1(y - 1);
    }
    (week_start(day, first_day) - w1) / 7 + 1
}

/// Two-digit years: 0..=29 are 2000s, 30..=99 are 1900s.
pub fn window_year(y: i64, digits: usize) -> i64 {
    if digits <= 2 && (0..100).contains(&y) { if y < 30 { 2000 + y } else { 1900 + y } } else { y }
}

/// Month number for a month name or its three-letter (or longer) abbreviation.
pub fn month_from_name(w: &str) -> Option<u32> {
    let w = w.to_ascii_lowercase();
    if w.len() < 3 {
        return None;
    }
    MONTHS.iter().position(|m| m.to_ascii_lowercase().starts_with(&w)).map(|i| i as u32 + 1)
}

fn is_weekday_name(w: &str) -> bool {
    let w = w.to_ascii_lowercase();
    w.len() >= 3 && WEEKDAYS.iter().any(|d| d.to_ascii_lowercase().starts_with(&w))
}

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Num(i64, usize),
    Word(String),
    Sep(char),
}

fn tokenize(s: &str) -> Option<Vec<Tok>> {
    let mut out = Vec::new();
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_ascii_digit() {
            let start = i;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
            let text: String = chars[start..i].iter().collect();
            out.push(Tok::Num(text.parse().ok()?, i - start));
        } else if c.is_alphabetic() {
            let start = i;
            while i < chars.len() && chars[i].is_alphabetic() {
                i += 1;
            }
            out.push(Tok::Word(chars[start..i].iter().collect()));
        } else if matches!(c, '/' | '-' | '.' | ',' | ':' | ' ' | '\t') {
            out.push(Tok::Sep(c));
            i += 1;
        } else {
            return None;
        }
    }
    Some(out)
}

/// Parses date and/or time text the way `CDate` does for en-US settings. `current_year` fills
/// in a missing year ("Oct 14"). Returns a serial date.
/// ORACLE: the language follows the system locale; we accept m/d/y (swapping when the first
/// number cannot be a month), ISO y-m-d, and month names in either order.
pub fn parse_date(s: &str, current_year: i64) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    // Plain numbers are not dates ("1.5", "46309").
    if parse_number(t).is_some() && !t.contains(['/', '-', ':']) {
        return None;
    }
    let toks = tokenize(t)?;
    // Time part: N:N[:N] [AM|PM] or N AM|PM.
    let mut time: Option<f64> = None;
    let mut rest: Vec<Tok> = Vec::new();
    let mut i = 0;
    while i < toks.len() {
        if let Tok::Num(h, _) = toks[i] {
            let colon = |k: usize| matches!(toks.get(k), Some(Tok::Sep(':')));
            let num = |k: usize| match toks.get(k) {
                Some(Tok::Num(n, _)) => Some(*n),
                _ => None,
            };
            let mut j = i;
            let mut parts = vec![h];
            while colon(j + 1) {
                parts.push(num(j + 2)?);
                j += 2;
            }
            // Look past spaces for AM/PM.
            let mut k = j + 1;
            while matches!(toks.get(k), Some(Tok::Sep(' '))) {
                k += 1;
            }
            let ampm = match toks.get(k) {
                Some(Tok::Word(w)) => match w.to_ascii_lowercase().as_str() {
                    "am" | "a" => Some(false),
                    "pm" | "p" => Some(true),
                    _ => None,
                },
                _ => None,
            };
            if parts.len() > 1 || ampm.is_some() {
                if time.is_some() || parts.len() > 3 {
                    return None;
                }
                let (mut hh, mm, ss) =
                    (parts[0], parts.get(1).copied().unwrap_or(0), parts.get(2).copied().unwrap_or(0));
                if let Some(pm) = ampm {
                    if !(1..=12).contains(&hh) {
                        return None;
                    }
                    hh %= 12;
                    if pm {
                        hh += 12;
                    }
                    j = k;
                }
                if hh > 23 || mm > 59 || ss > 59 {
                    return None;
                }
                time = Some((hh * 3600 + mm * 60 + ss) as f64 / 86_400.0);
                i = j + 1;
                continue;
            }
        }
        rest.push(toks[i].clone());
        i += 1;
    }
    let mut nums: Vec<(i64, usize)> = Vec::new();
    let mut month_name: Option<u32> = None;
    for tok in &rest {
        match tok {
            Tok::Num(n, len) => nums.push((*n, *len)),
            Tok::Word(w) => {
                if let Some(m) = month_from_name(w) {
                    if month_name.is_some() {
                        return None;
                    }
                    month_name = Some(m);
                } else if !is_weekday_name(w) {
                    return None;
                }
            }
            Tok::Sep(_) => {}
        }
    }
    let year_like = |(n, len): (i64, usize)| len >= 3 || n > 31;
    let (y, m, d) = match (month_name, nums.as_slice()) {
        (None, []) => {
            return time;
        }
        (Some(m), [a]) => {
            if year_like(*a) {
                (window_year(a.0, a.1), i64::from(m), 1)
            } else {
                (current_year, i64::from(m), a.0)
            }
        }
        (Some(m), [a, b]) => {
            if year_like(*a) {
                (window_year(a.0, a.1), i64::from(m), b.0)
            } else {
                (window_year(b.0, b.1), i64::from(m), a.0)
            }
        }
        (None, [a, b]) => {
            if year_like(*b) {
                (window_year(b.0, b.1), a.0, 1)
            } else if year_like(*a) {
                (window_year(a.0, a.1), b.0, 1)
            } else if a.0 > 12 && b.0 <= 12 {
                (current_year, b.0, a.0)
            } else {
                (current_year, a.0, b.0)
            }
        }
        (None, [a, b, c]) => {
            if year_like(*a) {
                (window_year(a.0, a.1), b.0, c.0)
            } else if a.0 > 12 && b.0 <= 12 {
                (window_year(c.0, c.1), b.0, a.0)
            } else {
                (window_year(c.0, c.1), a.0, b.0)
            }
        }
        _ => return None,
    };
    if !(1..=12).contains(&m) || d < 1 || d > i64::from(days_in_month(y, m as u32)) || !(100..=9999).contains(&y) {
        return None;
    }
    let day = ymd_to_serial(y, m, d) as i64;
    Some(join(day, time.unwrap_or(0.0)))
}

/// Coerces a value to a serial date (`CDate` rules): numbers are serials, strings are parsed.
pub fn to_date(v: &Value, current_year: i64) -> Result<f64> {
    let d = match v {
        Value::Date(d) => *d,
        Value::Null => return Err(RuntimeError::standard(94)),
        Value::String(s) => match parse_date(s, current_year) {
            Some(d) => d,
            // ORACLE: numeric text such as "46309" - we accept it as a serial.
            None => parse_number(s).ok_or_else(RuntimeError::type_mismatch)?,
        },
        Value::Object(Some(o)) => return to_date(&o.default_value()?, current_year),
        other => other.to_f64()?,
    };
    if in_range(d) { Ok(d) } else { Err(RuntimeError::overflow()) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parts() {
        assert_eq!(weekday(46309.0, 1), 4); // 2026-10-14 Wednesday
        assert_eq!(weekday(46309.0, 2), 3);
        assert_eq!(weekday(1.0, 1), 1);
        assert_eq!(weekday(0.0, 1), 7);
        assert_eq!(split(-1.25), (-1, 0.25));
        assert_eq!(join(-1, 0.25), -1.25);
        assert_eq!(linear(-1.25), -0.75);
        assert_eq!(from_linear(-0.75), -1.25);
        assert_eq!(hms(0.75), (18, 0, 0));
        assert_eq!(days_in_month(2024, 2), 29);
        assert_eq!(days_in_month(2100, 2), 28);
        assert_eq!(week_of_year(ymd_to_serial(2026, 1, 1), 1, 1), 1);
        assert_eq!(week_of_year(ymd_to_serial(2026, 1, 4), 1, 1), 2); // Jan 1 2026 is Thursday
        assert_eq!(week_of_year(ymd_to_serial(2026, 12, 31), 1, 1), 53);
    }

    #[test]
    fn parsing() {
        let d = |s: &str| parse_date(s, 2026);
        let serial = ymd_to_serial(2026, 10, 14);
        assert_eq!(d("10/14/2026"), Some(serial));
        assert_eq!(d("2026-10-14"), Some(serial));
        assert_eq!(d("14/10/2026"), Some(serial));
        assert_eq!(d("Oct 14, 2026"), Some(serial));
        assert_eq!(d("14 October 2026"), Some(serial));
        assert_eq!(d("14-Oct-26"), Some(serial));
        assert_eq!(d("Wednesday, October 14, 2026"), Some(serial));
        assert_eq!(d("10/14"), Some(serial));
        assert_eq!(d("10/14/2026 6:30 PM"), Some(serial + 0.770_833_333_333_333_3));
        assert_eq!(d("12:00:00 AM"), Some(0.0));
        assert_eq!(d("13:30"), Some(0.5625));
        assert_eq!(d("2/30/2026"), None);
        assert_eq!(d("hello"), None);
        assert_eq!(d("1.5"), None);
        assert_eq!(d(""), None);
        assert_eq!(d("1/1/30"), Some(ymd_to_serial(1930, 1, 1)));
        assert_eq!(d("1/1/29"), Some(ymd_to_serial(2029, 1, 1)));
    }
}
