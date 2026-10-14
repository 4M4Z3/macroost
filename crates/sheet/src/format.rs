//! Number format codes: rendering a cell's displayed text (`Range.Text`), recognizing date
//! formats (so `Range.Value` returns a Date), cell error text, and the default color palette.
//!
//! The documented format-code language: up to four sections (`pos;neg;zero;text`), conditions
//! (`[>100]`), colors (ignored for text), `0 # ? . , %`, thousands scaling, scientific and
//! engineering notation, fractions (`# ?/?`), quoted and escaped literals, `_x` spacing, `*x`
//! fill (omitted: it depends on the column width), `@`, and date/time codes including elapsed
//! time (`[h]:mm`) and fractional seconds. Numbers are rounded as displayed decimals of their
//! 15-significant-digit value, half away from zero.

use macroost_lang::Value;
use macroost_lang::value::serial_to_hms;

/// Text of a cell error value (`CVErr(xlErrNA)` -> `#N/A`).
pub fn error_text(code: i32) -> &'static str {
    match code {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2015 => "#VALUE!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        2042 => "#N/A",
        2043 => "#GETTING_DATA",
        2045 => "#SPILL!",
        2047 => "#BLOCKED!",
        2050 => "#CALC!",
        _ => "#VALUE!",
    }
}

/// Error code for error text (`#N/A` -> 2042), used when a string assigned to a cell is an error.
pub fn error_code(text: &str) -> Option<i32> {
    let t = text.trim().to_ascii_uppercase();
    Some(match t.as_str() {
        "#NULL!" => 2000,
        "#DIV/0!" => 2007,
        "#VALUE!" => 2015,
        "#REF!" => 2023,
        "#NAME?" => 2029,
        "#NUM!" => 2036,
        "#N/A" => 2042,
        _ => return None,
    })
}

/// Splits a format into sections on `;` outside quotes and escapes.
fn sections(fmt: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut chars = fmt.chars();
    let mut in_quote = false;
    while let Some(c) = chars.next() {
        let cur = out.last_mut();
        let Some(cur) = cur else { break };
        match c {
            '"' => {
                in_quote = !in_quote;
                cur.push(c);
            }
            '\\' if !in_quote => {
                cur.push(c);
                if let Some(n) = chars.next() {
                    cur.push(n);
                }
            }
            ';' if !in_quote => out.push(String::new()),
            _ => cur.push(c),
        }
    }
    out
}

/// The format code with quoted text, escapes and bracketed parts removed (except `[h]`/`[m]`/`[s]`).
fn strip_literals(section: &str) -> String {
    let mut out = String::new();
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for n in chars.by_ref() {
                    if n == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                let mut inner = String::new();
                for n in chars.by_ref() {
                    if n == ']' {
                        break;
                    }
                    inner.push(n);
                }
                let l = inner.to_ascii_lowercase();
                if !l.is_empty() && l.chars().all(|ch| matches!(ch, 'h' | 'm' | 's')) {
                    out.push_str(&l);
                }
            }
            _ => out.push(c),
        }
    }
    out
}

/// True when a number format displays dates or times.
pub fn is_date_format(fmt: &str) -> bool {
    let first = sections(fmt).into_iter().next().unwrap_or_default();
    if first.eq_ignore_ascii_case("general") {
        return false;
    }
    let s = strip_literals(&first).to_ascii_lowercase();
    !s.contains("general") && s.chars().any(|c| matches!(c, 'y' | 'd' | 'm' | 'h' | 's'))
}

/// True when a date/time format shows a calendar date (year, month or day), so that
/// `Range.Value` returns a Date; a time-only format reads back as a Double.
pub fn has_date_part(fmt: &str) -> bool {
    if !is_date_format(fmt) {
        return false;
    }
    let first = sections(fmt).into_iter().next().unwrap_or_default();
    let s = strip_literals(&first).to_ascii_lowercase().replace("am/pm", "").replace("a/p", "");
    if s.contains('y') || s.contains('d') {
        return true;
    }
    // 'm' is a month unless it follows an hour or precedes seconds.
    let codes: Vec<char> = s.chars().filter(|c| matches!(c, 'h' | 'm' | 's')).collect();
    let mut i = 0;
    while i < codes.len() {
        if codes[i] == 'm' {
            let mut j = i;
            while j < codes.len() && codes[j] == 'm' {
                j += 1;
            }
            let after_h = i > 0 && codes[i - 1] == 'h';
            let before_s = codes.get(j) == Some(&'s');
            if !after_h && !before_s {
                return true;
            }
            i = j;
        } else {
            i += 1;
        }
    }
    false
}

/// True when a number format is a currency format (so `Range.Value` returns Currency).
pub fn is_currency_format(fmt: &str) -> bool {
    let first = sections(fmt).into_iter().next().unwrap_or_default();
    let mut in_quote = false;
    let mut chars = first.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => in_quote = !in_quote,
            '\\' if !in_quote => {
                chars.next();
            }
            '_' | '*' if !in_quote => {
                chars.next();
            }
            '$' | '€' | '£' | '¥' if !in_quote => return true,
            '[' if !in_quote && chars.peek() == Some(&'$') => {
                return true;
            }
            _ => {}
        }
    }
    false
}

/// The displayed text of `value` under number format `fmt` (`Range.Text`).
pub fn format_value(value: &Value, fmt: &str) -> String {
    format_value_width(value, fmt, 8.43)
}

/// Like [`format_value`], for a column `width` characters wide: an impossible date fills the
/// cell with `#`.
pub fn format_value_width(value: &Value, fmt: &str, width: f64) -> String {
    match value {
        Value::Empty | Value::Null => String::new(),
        Value::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_owned(),
        Value::Error(code) => error_text(*code).to_owned(),
        Value::String(s) => format_text(s, fmt),
        Value::Object(_) | Value::Array(_) => String::new(),
        other => match other.to_f64() {
            Ok(x) => match format_number(x, fmt) {
                Some(s) => s,
                None => "#".repeat(width.max(1.0) as usize),
            },
            Err(_) => String::new(),
        },
    }
}

fn format_text(s: &str, fmt: &str) -> String {
    let secs = sections(fmt);
    let text_section =
        if secs.len() >= 4 { Some(secs[3].clone()) } else { secs.iter().find(|x| x.contains('@')).cloned() };
    match text_section {
        Some(sec) => {
            let mut out = String::new();
            for tok in tokenize(&sec) {
                match tok {
                    Tok::Lit(l) => out.push_str(&l),
                    Tok::Code('@') => out.push_str(s),
                    Tok::Code(_) => {}
                }
            }
            out
        }
        None => s.to_owned(),
    }
}

/// Rounds a non-negative number to `decimals` places the way the display does: on its
/// 15-significant-digit decimal form, half away from zero. Returns (integer digits, fraction
/// digits); the integer part has no leading zeros (empty for zero).
fn round_decimal(x: f64, decimals: usize) -> (String, String) {
    let x = x.abs();
    if x == 0.0 || !x.is_finite() {
        return (String::new(), "0".repeat(decimals));
    }
    let s = format!("{x:.14e}");
    let (mant, exp) = s.split_once('e').unwrap_or((&s, "0"));
    let exp: i64 = exp.parse().unwrap_or(0);
    let mut digits: Vec<u8> = mant.bytes().filter(u8::is_ascii_digit).map(|b| b - b'0').collect();
    // value = 0.d1d2d3... * 10^point
    let mut point = exp + 1;
    let keep = point + decimals as i64;
    if keep < 0 {
        return (String::new(), "0".repeat(decimals));
    }
    let keep = keep as usize;
    if keep < digits.len() {
        let round_up = digits[keep] >= 5;
        digits.truncate(keep);
        if round_up {
            let mut i = keep;
            loop {
                if i == 0 {
                    digits.insert(0, 1);
                    point += 1;
                    break;
                }
                i -= 1;
                if digits[i] == 9 {
                    digits[i] = 0;
                } else {
                    digits[i] += 1;
                    break;
                }
            }
        }
    }
    let total = (point + decimals as i64).max(0) as usize;
    while digits.len() < total {
        digits.push(0);
    }
    let mut int = String::new();
    let mut frac = String::new();
    for (i, d) in digits.iter().enumerate().take(total) {
        let ch = char::from(b'0' + d);
        if (i as i64) < point {
            int.push(ch);
        } else {
            frac.push(ch);
        }
    }
    if point < 0 {
        frac = "0".repeat((-point) as usize) + &frac;
        frac.truncate(decimals);
    }
    let int = int.trim_start_matches('0').to_owned();
    (int, frac)
}

/// the reference spreadsheet's General display of a number: at most 11 characters, switching to scientific notation
/// (6 significant digits) when fixed notation cannot show the value well.
pub fn general_number(x: f64) -> String {
    if x == 0.0 {
        return "0".into();
    }
    if !x.is_finite() {
        return "#NUM!".into();
    }
    let neg = x < 0.0;
    let a = x.abs();
    let sign = if neg { "-" } else { "" };
    let room = 11 - usize::from(neg);
    // Fixed notation with as many decimals as fit.
    let int_digits = if a >= 1.0 { a.log10().floor() as i64 + 1 } else { 1 };
    if int_digits as usize <= room {
        let decimals = room.saturating_sub(int_digits as usize + 1);
        let (i, f) = round_decimal(a, decimals);
        let f = f.trim_end_matches('0');
        let i = if i.is_empty() { "0".to_owned() } else { i };
        let fixed = if f.is_empty() { i.clone() } else { format!("{i}.{f}") };
        let exact = {
            // Does fixed notation keep every significant digit of the 15-digit value?
            let (fi, ff) = round_decimal(a, 30);
            let full = format!("{fi}{}", ff.trim_end_matches('0'));
            let shown = format!("{i}{f}");
            full.trim_start_matches('0') == shown.trim_start_matches('0')
        };
        let sig_shown = format!("{}{}", i.trim_start_matches('0'), f).trim_start_matches('0').len();
        if fixed.len() <= room && fixed != "0" && (exact || a >= 1.0 || sig_shown >= 6) {
            return format!("{sign}{fixed}");
        }
        if fixed == "0" && a >= 1e-9 && exact {
            return format!("{sign}{fixed}");
        }
    }
    // Scientific: up to 5 decimals.
    let s = format!("{a:.14e}");
    let exp: i32 = s.split_once('e').and_then(|(_, e)| e.parse().ok()).unwrap_or(0);
    let m = a / 10f64.powi(exp);
    let (mi, mf) = round_decimal(m, 5);
    let (mi, exp) = if mi == "10" { ("1".to_owned(), exp + 1) } else { (mi, exp) };
    let mf = mf.trim_end_matches('0');
    let mant = if mf.is_empty() { mi } else { format!("{mi}.{mf}") };
    format!("{sign}{mant}E{}{:02}", if exp < 0 { '-' } else { '+' }, exp.abs())
}

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Lit(String),
    Code(char),
}

/// Splits a section into literal text and code characters.
fn tokenize(section: &str) -> Vec<Tok> {
    let mut out = Vec::new();
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                let mut lit = String::new();
                for n in chars.by_ref() {
                    if n == '"' {
                        break;
                    }
                    lit.push(n);
                }
                out.push(Tok::Lit(lit));
            }
            '\\' => {
                if let Some(n) = chars.next() {
                    out.push(Tok::Lit(n.to_string()));
                }
            }
            '_' => {
                chars.next();
                out.push(Tok::Lit(" ".into()));
            }
            '*' => {
                chars.next();
            }
            '[' => {
                let mut inner = String::new();
                for n in chars.by_ref() {
                    if n == ']' {
                        break;
                    }
                    inner.push(n);
                }
                if let Some(cur) = inner.strip_prefix('$') {
                    // [$€-407] currency symbol
                    let sym = cur.split('-').next().unwrap_or("");
                    out.push(Tok::Lit(sym.to_owned()));
                } else if !inner.is_empty()
                    && inner.chars().all(|ch| matches!(ch.to_ascii_lowercase(), 'h' | 'm' | 's'))
                {
                    out.push(Tok::Code('['));
                    for ch in inner.chars() {
                        out.push(Tok::Code(ch));
                    }
                    out.push(Tok::Code(']'));
                }
            }
            '$' | '-' | '+' | '(' | ')' | ':' | '/' | ' ' | '!' | '^' | '&' | '\'' | '~' | '{' | '}' | '<' | '>'
            | '=' => out.push(Tok::Lit(c.to_string())),
            _ => out.push(Tok::Code(c)),
        }
    }
    out
}

/// A section's `[>100]` style condition.
fn condition(section: &str) -> Option<(String, f64)> {
    let mut chars = section.chars();
    let mut in_quote = false;
    while let Some(c) = chars.next() {
        match c {
            '"' => in_quote = !in_quote,
            '[' if !in_quote => {
                let inner: String = chars.by_ref().take_while(|&c| c != ']').collect();
                let op: String = inner.chars().take_while(|c| matches!(c, '<' | '>' | '=')).collect();
                if !op.is_empty() {
                    let n: f64 = inner[op.len()..].trim().parse().ok()?;
                    return Some((op, n));
                }
            }
            _ => {}
        }
    }
    None
}

fn test_condition(op: &str, n: f64, x: f64) -> bool {
    match op {
        "<" => x < n,
        "<=" | "=<" => x <= n,
        ">" => x > n,
        ">=" | "=>" => x >= n,
        "=" => x == n,
        "<>" => x != n,
        _ => false,
    }
}

/// Formats a number; `None` for an impossible date (shown as `####`).
fn format_number(x: f64, fmt: &str) -> Option<String> {
    if fmt.is_empty() || fmt.eq_ignore_ascii_case("general") {
        return Some(general_number(x));
    }
    let secs = sections(fmt);
    let conds: Vec<Option<(String, f64)>> = secs.iter().map(|s| condition(s)).collect();
    let (section, signless) = if conds.iter().take(2).any(Option::is_some) {
        let pick = if conds[0].as_ref().is_some_and(|(o, n)| test_condition(o, *n, x)) {
            0
        } else if conds.get(1).and_then(Option::as_ref).is_some_and(|(o, n)| test_condition(o, *n, x)) {
            1
        } else if secs.len() >= 3 {
            2
        } else if conds.get(1).is_some_and(Option::is_none) {
            1
        } else {
            0
        };
        (secs[pick].clone(), pick == 1 && conds.get(1).is_some_and(Option::is_none) && x < 0.0)
    } else {
        match secs.len() {
            0 | 1 => (secs.first().cloned().unwrap_or_default(), false),
            2 => {
                if x < 0.0 {
                    (secs[1].clone(), true)
                } else {
                    (secs[0].clone(), false)
                }
            }
            _ => {
                if x < 0.0 {
                    (secs[1].clone(), true)
                } else if x == 0.0 {
                    (secs[2].clone(), false)
                } else {
                    (secs[0].clone(), false)
                }
            }
        }
    };
    let v = if signless { x.abs() } else { x };
    let plain: String = strip_literals(&section);
    if plain.trim().eq_ignore_ascii_case("general") {
        return Some(general_number(v));
    }
    if is_date_format(&section) {
        if !(0.0..2_958_466.0).contains(&v) {
            return None;
        }
        return Some(format_date_section(v, &section));
    }
    Some(format_numeric_section(v, &section))
}

fn is_ph(t: &Tok) -> bool {
    matches!(t, Tok::Code('0' | '#' | '?'))
}

fn format_numeric_section(x: f64, section: &str) -> String {
    let toks = tokenize(section);
    let first = toks.iter().position(is_ph);
    let last = toks.iter().rposition(is_ph);
    let (Some(first), Some(last)) = (first, last) else {
        // No digit placeholders: literals only (a text placeholder shows the number as General).
        let mut out = String::new();
        let mut shown = false;
        for t in &toks {
            match t {
                Tok::Lit(l) => out.push_str(l),
                Tok::Code('@') | Tok::Code('G') | Tok::Code('g') if !shown => {
                    out.push_str(&general_number(x));
                    shown = true;
                }
                Tok::Code('%') => out.push('%'),
                Tok::Code(c) if c.is_ascii_alphabetic() => {}
                Tok::Code(c) => out.push(*c),
            }
        }
        return out;
    };
    // Fractions: a '/' after the placeholders, then a denominator (placeholders or digits).
    if let Some(slash) = toks[first..].iter().position(|t| *t == Tok::Lit("/".into())).map(|p| p + first) {
        let mut end = slash;
        while matches!(toks.get(end + 1), Some(Tok::Code(c)) if matches!(c, '0'..='9' | '#' | '?')) {
            end += 1;
        }
        if end > slash {
            return format_fraction(x, &toks, first, end, slash);
        }
    }
    // Scientific.
    if let Some(e) = toks[first..].iter().position(|t| matches!(t, Tok::Code('E' | 'e'))).map(|p| p + first)
        && matches!(toks.get(e + 1), Some(Tok::Lit(s)) if s == "+" || s == "-")
    {
        let mut end = e + 2;
        while end < toks.len() && is_ph(&toks[end]) {
            end += 1;
        }
        let mant: Vec<Tok> = toks[first..e].to_vec();
        let exp_digits = toks[e + 2..end].iter().filter(|t| is_ph(t)).count().max(1);
        let plus = matches!(&toks[e + 1], Tok::Lit(s) if s == "+");
        let body = render_scientific(x.abs(), &mant, exp_digits, plus);
        let mut out = String::new();
        if x < 0.0 {
            out.push('-');
        }
        push_lits(&mut out, &toks[..first]);
        out.push_str(&body);
        push_lits(&mut out, &toks[end..]);
        return out;
    }
    // Extend over trailing commas (scaling) right after the last placeholder.
    let mut end = last;
    while end + 1 < toks.len() && matches!(toks[end + 1], Tok::Code(',')) {
        end += 1;
    }
    let percent = toks.iter().filter(|t| matches!(t, Tok::Code('%'))).count() as i32;
    let body_toks = &toks[first..=end];
    let trailing_commas = body_toks.iter().rev().take_while(|t| matches!(t, Tok::Code(','))).count();
    let body_toks = &body_toks[..body_toks.len() - trailing_commas];
    let mut value = x * 100f64.powi(percent) / 1000f64.powi(trailing_commas as i32);
    let neg = value < 0.0;
    value = value.abs();
    let dot = body_toks.iter().position(|t| *t == Tok::Code('.'));
    let (int_toks, frac_toks) = match dot {
        Some(d) => (&body_toks[..d], &body_toks[d + 1..]),
        None => (body_toks, &body_toks[..0]),
    };
    let decimals = frac_toks.iter().filter(|t| is_ph(t)).count();
    let (int_digits, frac_digits) = round_decimal(value, decimals);
    let thousands = int_toks.iter().any(|t| matches!(t, Tok::Code(',')));
    let int_text = fill_integer(&int_digits, int_toks, thousands);
    let mut body = int_text;
    if dot.is_some() {
        body.push('.');
        body.push_str(&fill_fraction(&frac_digits, frac_toks));
    }
    let mut out = String::new();
    let nonzero = int_digits.chars().any(|c| c != '0') || frac_digits.chars().any(|c| c != '0');
    if neg && nonzero {
        out.push('-');
    }
    push_lits(&mut out, &toks[..first]);
    out.push_str(&body);
    push_lits(&mut out, &toks[end + 1..]);
    out
}

fn push_lits(out: &mut String, toks: &[Tok]) {
    for t in toks {
        match t {
            Tok::Lit(l) => out.push_str(l),
            Tok::Code('%') => out.push('%'),
            Tok::Code(c) if c.is_ascii_alphabetic() || *c == '@' => {}
            Tok::Code(c) => out.push(*c),
        }
    }
}

/// Integer placeholders filled with digits (right to left); extra digits go at the left.
fn fill_integer(digits: &str, toks: &[Tok], thousands: bool) -> String {
    let phs: Vec<char> = toks
        .iter()
        .filter_map(|t| match t {
            Tok::Code(c @ ('0' | '#' | '?')) => Some(*c),
            _ => None,
        })
        .collect();
    let has_inner_lits = toks.iter().any(|t| matches!(t, Tok::Lit(_)));
    if thousands || !has_inner_lits {
        // Pad to the zero placeholders, then group.
        let zeros = phs.iter().filter(|c| **c == '0').count();
        let qs = phs.iter().filter(|c| **c == '?').count();
        let mut d = digits.to_owned();
        while d.len() < zeros {
            d.insert(0, '0');
        }
        if thousands && d.len() > 3 {
            let n = d.len();
            let mut g = String::new();
            for (i, ch) in d.chars().enumerate() {
                if i > 0 && (n - i).is_multiple_of(3) {
                    g.push(',');
                }
                g.push(ch);
            }
            d = g;
        }
        let width = zeros + qs;
        let mut out = String::new();
        let digit_count = d.chars().filter(char::is_ascii_digit).count();
        if digit_count < width {
            out.push_str(&" ".repeat((width - digit_count).min(qs)));
        }
        // Literals between placeholders (rare with grouping) are appended after the digits.
        out.push_str(&d);
        return out;
    }
    let mut ds: Vec<char> = digits.chars().collect();
    let mut out: Vec<String> = Vec::new();
    let ph_total = phs.len();
    let mut seen = 0;
    for t in toks.iter().rev() {
        match t {
            Tok::Code(c @ ('0' | '#' | '?')) => {
                seen += 1;
                let mut piece = match ds.pop() {
                    Some(d) => d.to_string(),
                    None => match c {
                        '0' => "0".into(),
                        '?' => " ".into(),
                        _ => String::new(),
                    },
                };
                if seen == ph_total && !ds.is_empty() {
                    let rest: String = ds.drain(..).collect();
                    piece = rest + &piece;
                }
                out.push(piece);
            }
            Tok::Lit(l) => out.push(l.clone()),
            Tok::Code(c) => out.push(c.to_string()),
        }
    }
    out.reverse();
    out.concat()
}

/// Fraction placeholders filled left to right; trailing zeros drop for `#`, become spaces for `?`.
fn fill_fraction(digits: &str, toks: &[Tok]) -> String {
    let ds: Vec<char> = digits.chars().collect();
    let phs: Vec<char> = toks
        .iter()
        .filter_map(|t| match t {
            Tok::Code(c @ ('0' | '#' | '?')) => Some(*c),
            _ => None,
        })
        .collect();
    let mut shown: Vec<String> = ds.iter().map(|c| c.to_string()).collect();
    let mut i = shown.len();
    while i > 0 {
        let ph = phs.get(i - 1).copied().unwrap_or('0');
        if shown[i - 1] == "0" && ph != '0' {
            shown[i - 1] = if ph == '?' { " ".into() } else { String::new() };
            i -= 1;
        } else {
            break;
        }
    }
    let mut out = String::new();
    let mut k = 0;
    for t in toks {
        match t {
            Tok::Code('0' | '#' | '?') => {
                out.push_str(shown.get(k).map(String::as_str).unwrap_or(""));
                k += 1;
            }
            Tok::Lit(l) => out.push_str(l),
            Tok::Code('%') => {}
            Tok::Code(c) => out.push(*c),
        }
    }
    out
}

fn render_scientific(value: f64, mant: &[Tok], exp_digits: usize, plus: bool) -> String {
    let dot = mant.iter().position(|t| *t == Tok::Code('.'));
    let (int_toks, frac_toks) = match dot {
        Some(d) => (&mant[..d], &mant[d + 1..]),
        None => (mant, &mant[..0]),
    };
    let int_ph = int_toks.iter().filter(|t| is_ph(t)).count().max(1) as i32;
    let decimals = frac_toks.iter().filter(|t| is_ph(t)).count();
    let mut exp = if value == 0.0 { 0 } else { value.log10().floor() as i32 };
    if int_ph > 1 {
        // Engineering style: the exponent is a multiple of the integer placeholder count.
        exp = exp.div_euclid(int_ph) * int_ph;
    }
    let mut m = if value == 0.0 { 0.0 } else { value / 10f64.powi(exp) };
    let (mut i, mut f) = round_decimal(m, decimals);
    let limit = 10f64.powi(int_ph);
    if i.parse::<f64>().unwrap_or(0.0) >= limit {
        exp += int_ph;
        m = value / 10f64.powi(exp);
        let r = round_decimal(m, decimals);
        i = r.0;
        f = r.1;
    }
    let min_int = int_toks.iter().filter(|t| **t == Tok::Code('0')).count();
    while i.len() < min_int {
        i.insert(0, '0');
    }
    let mut out = i;
    if dot.is_some() {
        out.push('.');
        out.push_str(&fill_fraction(&f, frac_toks));
    }
    let sign = if exp < 0 {
        "-"
    } else if plus {
        "+"
    } else {
        ""
    };
    format!("{out}E{sign}{:0w$}", exp.abs(), w = exp_digits)
}

/// `# ?/?`, `?/?`, `# ?/8` and similar.
fn format_fraction(x: f64, toks: &[Tok], first: usize, last: usize, slash: usize) -> String {
    let neg = x < 0.0;
    let a = x.abs();
    // Integer part: placeholders before the first space-separated group of the numerator.
    let num_start = (first..slash).rev().take_while(|&i| is_ph(&toks[i])).last().unwrap_or(slash);
    let has_int = toks[first..num_start].iter().any(is_ph);
    let num_phs: Vec<char> = toks[num_start..slash]
        .iter()
        .filter_map(|t| match t {
            Tok::Code(c) => Some(*c),
            _ => None,
        })
        .collect();
    let den_toks = &toks[slash + 1..=last];
    let fixed_den: Option<u64> = {
        let d: String = den_toks
            .iter()
            .map(|t| match t {
                Tok::Code(c) => c.to_string(),
                Tok::Lit(l) => l.clone(),
            })
            .collect();
        d.trim().parse().ok()
    };
    let den_phs = den_toks.iter().filter(|t| is_ph(t)).count();
    let (int_part, frac) = if has_int { (a.trunc(), a - a.trunc()) } else { (0.0, a) };
    let (mut num, mut den) = match fixed_den {
        Some(d) => ((frac * d as f64).round() as u64, d),
        None => {
            let max_den = 10u64.pow(den_phs.max(1) as u32) - 1;
            let mut best = (0u64, 1u64, f64::MAX);
            for d in 1..=max_den {
                let n = (frac * d as f64).round();
                let err = (frac - n / d as f64).abs();
                if err < best.2 - 1e-12 {
                    best = (n as u64, d, err);
                }
            }
            (best.0, best.1)
        }
    };
    let mut int_part = int_part;
    if has_int && num == den && den > 0 {
        int_part += 1.0;
        num = 0;
    }
    if !has_int && fixed_den.is_none() && num == 0 {
        den = 1;
    }
    let mut out = String::new();
    if neg && (int_part > 0.0 || num > 0) {
        out.push('-');
    }
    push_lits(&mut out, &toks[..first]);
    if has_int {
        let int_toks = &toks[first..num_start];
        let digits = if int_part > 0.0 { format!("{}", int_part as u64) } else { String::new() };
        out.push_str(&fill_integer(&digits, int_toks, false));
    }
    let pad = |s: String, n: usize, left: bool| -> String {
        if s.len() >= n {
            s
        } else if left {
            " ".repeat(n - s.len()) + &s
        } else {
            let k = n - s.len();
            s + &" ".repeat(k)
        }
    };
    if has_int && num == 0 {
        // Whole number: the fraction part is blank, keeping its width.
        out.push_str(&" ".repeat(num_phs.len() + 1 + den_toks.len()));
    } else {
        out.push_str(&pad(num.to_string(), num_phs.len(), true));
        out.push('/');
        match fixed_den {
            Some(d) => out.push_str(&d.to_string()),
            None => out.push_str(&pad(den.to_string(), den_phs, false)),
        }
    }
    push_lits(&mut out, &toks[last + 1..]);
    out
}

const MONTHS: [&str; 12] = [
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
const DAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

fn format_date_section(serial: f64, section: &str) -> String {
    let toks = tokenize(section);
    // Group consecutive identical code letters.
    let mut groups: Vec<(char, usize)> = Vec::new();
    let mut lits: Vec<Option<String>> = Vec::new();
    let mut i = 0;
    while i < toks.len() {
        match &toks[i] {
            Tok::Lit(l) => {
                groups.push(('\0', 0));
                lits.push(Some(l.clone()));
                i += 1;
            }
            Tok::Code('[') => {
                // Elapsed time: [h], [mm], [ss].
                let mut j = i + 1;
                let mut letters = String::new();
                while let Some(Tok::Code(c)) = toks.get(j) {
                    if *c == ']' {
                        break;
                    }
                    letters.push(c.to_ascii_lowercase());
                    j += 1;
                }
                let kind = letters.chars().next().unwrap_or('h');
                groups.push((kind.to_ascii_uppercase(), letters.len()));
                lits.push(None);
                i = j + 1;
            }
            Tok::Code(c) => {
                let lc = c.to_ascii_lowercase();
                // AM/PM and A/P markers.
                let rest: String = toks[i..]
                    .iter()
                    .take(5)
                    .map(|t| match t {
                        Tok::Code(ch) => *ch,
                        Tok::Lit(l) => l.chars().next().unwrap_or('\0'),
                    })
                    .collect();
                if rest.to_ascii_uppercase().starts_with("AM/PM") {
                    groups.push(('P', 5));
                    lits.push(None);
                    i += 5;
                    continue;
                }
                if rest.to_ascii_uppercase().starts_with("A/P") {
                    groups.push(('p', 3));
                    lits.push(None);
                    i += 3;
                    continue;
                }
                let mut n = 1;
                while i + n < toks.len() && matches!(&toks[i + n], Tok::Code(d) if d.to_ascii_lowercase() == lc) {
                    n += 1;
                }
                groups.push((lc, n));
                lits.push(None);
                i += n;
            }
        }
    }
    let has_ampm = groups.iter().any(|(c, _)| *c == 'P' || *c == 'p');
    // Fractional seconds: `.0` digits right after a seconds code.
    let sec_decimals = groups
        .windows(2)
        .find(|w| (w[0].0 == 's' || w[0].0 == 'S') && w[1].0 == '.')
        .and_then(|_| groups.iter().position(|g| g.0 == '.'))
        .and_then(|p| groups.get(p + 1))
        .filter(|g| g.0 == '0')
        .map_or(0, |g| g.1);
    let total_secs = {
        let scale = 10f64.powi(sec_decimals as i32);
        (serial * 86_400.0 * scale).round() / scale
    };
    let whole_secs = total_secs.floor();
    let frac_secs = total_secs - whole_secs;
    let day_serial = (whole_secs / 86_400.0).floor();
    let (y, mo, d) = crate::eval::serial_ymd(day_serial);
    let (h, mi, s) = if sec_decimals > 0 {
        let t = (whole_secs - day_serial * 86_400.0) as u32;
        (t / 3600, t / 60 % 60, t % 60)
    } else {
        serial_to_hms(serial)
    };
    let weekday = ((day_serial as i64 - 1).rem_euclid(7)) as usize; // serial 1 = Sunday
    let mut out = String::new();
    let mut k = 0;
    while k < groups.len() {
        let (c, n) = groups[k];
        if let Some(Some(l)) = lits.get(k) {
            out.push_str(l);
            k += 1;
            continue;
        }
        // 'm' is minutes right after an hour code or right before a seconds code.
        let prev_code = groups[..k].iter().rev().find(|(pc, _)| *pc != '\0' && *pc != '.').map(|g| g.0);
        let next_code = groups[k + 1..].iter().find(|(nc, _)| *nc != '\0').map(|g| g.0);
        match c {
            'y' => {
                if n <= 2 {
                    out.push_str(&format!("{:02}", y.rem_euclid(100)));
                } else {
                    out.push_str(&format!("{y:04}"));
                }
            }
            'm' if matches!(prev_code, Some('h' | 'H')) || matches!(next_code, Some('s' | 'S')) => {
                if n >= 2 {
                    out.push_str(&format!("{mi:02}"));
                } else {
                    out.push_str(&mi.to_string());
                }
            }
            'm' => {
                let name = MONTHS.get((mo as usize).saturating_sub(1)).copied().unwrap_or("");
                match n {
                    1 => out.push_str(&mo.to_string()),
                    2 => out.push_str(&format!("{mo:02}")),
                    3 => out.push_str(&name[..name.len().min(3)]),
                    4 => out.push_str(name),
                    _ => out.push_str(&name[..1.min(name.len())]),
                }
            }
            'd' => {
                let name = DAYS.get(weekday).copied().unwrap_or("");
                match n {
                    1 => out.push_str(&d.to_string()),
                    2 => out.push_str(&format!("{d:02}")),
                    3 => out.push_str(&name[..name.len().min(3)]),
                    _ => out.push_str(name),
                }
            }
            'h' => {
                let hh = if has_ampm {
                    match h {
                        0 => 12,
                        13..=23 => h - 12,
                        _ => h,
                    }
                } else {
                    h
                };
                if n >= 2 {
                    out.push_str(&format!("{hh:02}"));
                } else {
                    out.push_str(&hh.to_string());
                }
            }
            's' => {
                if n >= 2 {
                    out.push_str(&format!("{s:02}"));
                } else {
                    out.push_str(&s.to_string());
                }
            }
            'H' | 'M' | 'S' => {
                let total = (serial * 86_400.0).round() as i64;
                let v = match c {
                    'H' => total / 3600,
                    'M' => total / 60,
                    _ => total,
                };
                out.push_str(&format!("{v:0w$}", w = n));
            }
            'P' => out.push_str(if h < 12 { "AM" } else { "PM" }),
            'p' => out.push_str(if h < 12 { "A" } else { "P" }),
            '.' => {
                if sec_decimals > 0 && groups.get(k + 1).is_some_and(|g| g.0 == '0') {
                    let f = format!("{:.*}", sec_decimals, frac_secs);
                    out.push_str(f.strip_prefix('0').unwrap_or(&f));
                    k += 2;
                    continue;
                }
                out.push('.');
            }
            '0' => {}
            other => out.push(other),
        }
        k += 1;
    }
    out
}

/// The default 56-color palette (`ColorIndex` 1..=56) as RGB longs.
pub const PALETTE: [u32; 56] = {
    const fn rgb(hex: u32) -> u32 {
        let r = (hex >> 16) & 0xFF;
        let g = (hex >> 8) & 0xFF;
        let b = hex & 0xFF;
        r | (g << 8) | (b << 16)
    }
    [
        rgb(0x000000),
        rgb(0xFFFFFF),
        rgb(0xFF0000),
        rgb(0x00FF00),
        rgb(0x0000FF),
        rgb(0xFFFF00),
        rgb(0xFF00FF),
        rgb(0x00FFFF),
        rgb(0x800000),
        rgb(0x008000),
        rgb(0x000080),
        rgb(0x808000),
        rgb(0x800080),
        rgb(0x008080),
        rgb(0xC0C0C0),
        rgb(0x808080),
        rgb(0x9999FF),
        rgb(0x993366),
        rgb(0xFFFFCC),
        rgb(0xCCFFFF),
        rgb(0x660066),
        rgb(0xFF8080),
        rgb(0x0066CC),
        rgb(0xCCCCFF),
        rgb(0x000080),
        rgb(0xFF00FF),
        rgb(0xFFFF00),
        rgb(0x00FFFF),
        rgb(0x800080),
        rgb(0x800000),
        rgb(0x008080),
        rgb(0x0000FF),
        rgb(0x00CCFF),
        rgb(0xCCFFFF),
        rgb(0xCCFFCC),
        rgb(0xFFFF99),
        rgb(0x99CCFF),
        rgb(0xFF99CC),
        rgb(0xCC99FF),
        rgb(0xFFCC99),
        rgb(0x3366FF),
        rgb(0x33CCCC),
        rgb(0x99CC00),
        rgb(0xFFCC00),
        rgb(0xFF9900),
        rgb(0xFF6600),
        rgb(0x666699),
        rgb(0x969696),
        rgb(0x003366),
        rgb(0x339966),
        rgb(0x003300),
        rgb(0x333300),
        rgb(0x993300),
        rgb(0x993366),
        rgb(0x333399),
        rgb(0x333333),
    ]
};

/// Nearest palette index (1-based) for an RGB long: smallest sum of channel differences, the
/// lower index on a tie.
pub fn nearest_color_index(color: u32) -> i32 {
    let split = |c: u32| (i64::from(c & 0xFF), i64::from((c >> 8) & 0xFF), i64::from((c >> 16) & 0xFF));
    let (r, g, b) = split(color);
    let mut best = (i64::MAX, 1);
    for (i, &p) in PALETTE.iter().enumerate() {
        let (pr, pg, pb) = split(p);
        let d = (r - pr).abs() + (g - pg).abs() + (b - pb).abs();
        if d < best.0 {
            best = (d, i as i32 + 1);
        }
    }
    best.1
}

#[cfg(test)]
mod tests {
    use super::*;

    fn f(x: f64, fmt: &str) -> String {
        format_value(&Value::Double(x), fmt)
    }

    #[test]
    fn numeric_formats() {
        assert_eq!(f(1234.567, "General"), "1234.567");
        assert_eq!(f(1234.567, "0.00"), "1234.57");
        assert_eq!(f(1234.567, "#,##0.00"), "1,234.57");
        assert_eq!(f(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(f(0.125, "0.0%"), "12.5%");
        assert_eq!(f(-5.0, "$#,##0.00"), "-$5.00");
        assert_eq!(f(-5.0, "0.00;(0.00)"), "(5.00)");
        assert_eq!(f(0.0, "0;-0;\"zero\""), "zero");
        assert_eq!(f(7.0, "000"), "007");
        assert_eq!(f(0.5, "#.##"), ".5");
        assert_eq!(f(12345.0, "0.00E+00"), "1.23E+04");
        assert_eq!(f(1e12, "General"), "1E+12");
        assert_eq!(f(1.005, "0.00"), "1.01");
        assert_eq!(f(2.5, "0"), "3");
        assert_eq!(f(123456789.0, "000-00-0000"), "123-45-6789");
        assert_eq!(f(1.5, "# ?/?"), "1 1/2");
        assert_eq!(f(0.333, "# ??/??"), "  1/3 ");
        assert_eq!(f(5.0, "# ?/?"), "5    ");
        assert_eq!(f(12345.678, "##0.0E+0"), "12.3E+3");
        assert_eq!(f(1.5, "[h]:mm"), "36:00");
        assert_eq!(f(1.23456789012345E-05, "General"), "1.23457E-05");
        assert_eq!(f(std::f64::consts::PI, "General"), "3.141592654");
        assert_eq!(f(0.00001234, "General"), "0.00001234");
        assert_eq!(format_value(&Value::str("x"), "\"<\"@\">\""), "<x>");
    }

    #[test]
    fn date_formats() {
        assert!(is_date_format("m/d/yyyy"));
        assert!(is_date_format("hh:mm:ss"));
        assert!(!is_date_format("0.00"));
        assert!(!is_date_format("General"));
        assert!(!is_date_format("\"days\" 0"));
        assert!(has_date_part("m/d/yyyy"));
        assert!(!has_date_part("h:mm AM/PM"));
        assert_eq!(f(46309.0, "m/d/yyyy"), "10/14/2026");
        assert_eq!(f(46309.0, "yyyy-mm-dd"), "2026-10-14");
        assert_eq!(f(46309.75, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(f(46309.0, "dddd, mmmm d"), "Wednesday, October 14");
        assert_eq!(f(46309.0, "mmm-yy"), "Oct-26");
        assert_eq!(f(60.0, "m/d/yyyy"), "2/29/1900");
        assert_eq!(f(1.0, "m/d/yyyy"), "1/1/1900");
    }

    #[test]
    fn palette() {
        assert_eq!(PALETTE[2], 255); // ColorIndex 3 = red
        assert_eq!(nearest_color_index(255), 3);
        assert_eq!(nearest_color_index(0xFFFFFF), 2);
        assert_eq!(nearest_color_index(0xFF8000), 33);
    }
}
