//! Tokens ([MS-VBAL] 3.3). Line continuations (` _` at end of line) are joined, comments and
//! `Rem` lines are dropped, and every logical line ends with a `Newline` token. `:` separates
//! statements on a line.

use crate::error::{Result, RuntimeError};

#[derive(Debug, Clone, PartialEq)]
pub enum Tok {
    Ident(String),
    /// An identifier written with a type character (`name$`, `n%`, `x&`, `d#`, `s!`, `c@`).
    TypedIdent(String, char),
    Integer(i64),
    /// Integer literal with an explicit type suffix: `&` Long, `%` Integer, `^` LongLong.
    TypedInteger(i64, char),
    Float(f64),
    /// Float literal with `!` Single, `#` Double, `@` Currency.
    TypedFloat(f64, char),
    Str(String),
    /// `#...#` date literal, already converted to a serial date.
    Date(f64),
    /// Punctuation and operators: ( ) , . ! : = < > <= >= <> + - * / \ ^ & := ;
    Sym(&'static str),
    Newline,
}

#[derive(Debug, Clone)]
pub struct Token {
    pub tok: Tok,
    pub line: u32,
}

pub fn lex(src: &str) -> Result<Vec<Token>> {
    let mut out = Vec::new();
    let chars: Vec<char> = src.chars().collect();
    let mut i = 0usize;
    let mut line = 1u32;
    let n = chars.len();

    // `Rem` is a comment only where a statement starts.
    let at_statement_start = |out: &Vec<Token>| {
        matches!(out.last().map(|t: &Token| &t.tok), None | Some(Tok::Newline) | Some(Tok::Sym(":")))
    };

    while i < n {
        let c = chars[i];
        match c {
            ' ' | '\t' | '\u{a0}' => i += 1,
            '\r' => i += 1,
            '\n' => {
                push_newline(&mut out, line);
                line += 1;
                i += 1;
            }
            '_' if is_continuation(&chars, i) => {
                // Skip to the end of the physical line; the logical line continues.
                while i < n && chars[i] != '\n' {
                    i += 1;
                }
                i += 1;
                line += 1;
            }
            '\'' => {
                while i < n && chars[i] != '\n' {
                    // A comment may itself be continued with ` _`.
                    if chars[i] == '_' && is_continuation(&chars, i) {
                        while i < n && chars[i] != '\n' {
                            i += 1;
                        }
                        i += 1;
                        line += 1;
                        continue;
                    }
                    i += 1;
                }
            }
            '"' => {
                let mut s = String::new();
                i += 1;
                loop {
                    if i >= n || chars[i] == '\n' {
                        // Unterminated strings run to the end of the line, as the editor allows.
                        break;
                    }
                    if chars[i] == '"' {
                        if i + 1 < n && chars[i + 1] == '"' {
                            s.push('"');
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    s.push(chars[i]);
                    i += 1;
                }
                out.push(Token { tok: Tok::Str(s), line });
            }
            '#' if !prev_is_operand(&out) => {
                // Date literal #...#, or `#` used for file numbers (`Print #1`).
                if let Some(end) = chars[i + 1..].iter().position(|&c| c == '#' || c == '\n') {
                    let body: String = chars[i + 1..i + 1 + end].iter().collect();
                    if chars.get(i + 1 + end) == Some(&'#') {
                        if let Some(serial) = parse_date_literal(&body) {
                            out.push(Token { tok: Tok::Date(serial), line });
                            i += end + 2;
                            continue;
                        }
                    }
                }
                out.push(Token { tok: Tok::Sym("#"), line });
                i += 1;
            }
            '&' if matches!(chars.get(i + 1), Some('h' | 'H' | 'o' | 'O'))
                && chars.get(i + 2).is_some_and(|c| c.is_ascii_alphanumeric()) =>
            {
                let radix = if matches!(chars[i + 1], 'h' | 'H') { 16 } else { 8 };
                let mut j = i + 2;
                while j < n && chars[j].is_digit(radix) {
                    j += 1;
                }
                let digits: String = chars[i + 2..j].iter().collect();
                let raw = u64::from_str_radix(&digits, radix).map_err(|_| syntax(line, "bad number"))?;
                let suffix = chars.get(j).copied().filter(|c| matches!(c, '&' | '%' | '^'));
                if suffix.is_some() {
                    j += 1;
                }
                // Hex literals are Integer when they fit 16 bits (so &HFFFF = -1), else Long.
                let value = match suffix {
                    Some('&') => i64::from(raw as u32 as i32),
                    Some('^') => raw as i64,
                    _ if raw <= 0xFFFF => i64::from(raw as u16 as i16),
                    _ => i64::from(raw as u32 as i32),
                };
                out.push(Token {
                    tok: match suffix {
                        Some(s) => Tok::TypedInteger(value, s),
                        None if raw <= 0xFFFF => Tok::TypedInteger(value, '%'),
                        None => Tok::TypedInteger(value, '&'),
                    },
                    line,
                });
                i = j;
            }
            c if c.is_ascii_digit()
                || (c == '.' && chars.get(i + 1).is_some_and(char::is_ascii_digit) && !prev_is_operand(&out)) =>
            {
                let start = i;
                let mut is_float = false;
                while i < n && chars[i].is_ascii_digit() {
                    i += 1;
                }
                if i < n && chars[i] == '.' && chars.get(i + 1).is_none_or(|c| !c.is_alphabetic()) {
                    is_float = true;
                    i += 1;
                    while i < n && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                }
                if i < n
                    && matches!(chars[i], 'e' | 'E' | 'd' | 'D')
                    && chars.get(i + 1).is_some_and(|c| c.is_ascii_digit() || *c == '+' || *c == '-')
                {
                    is_float = true;
                    i += 2;
                    while i < n && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                }
                let text: String = chars[start..i].iter().collect::<String>().replace(['d', 'D'], "e");
                let suffix = chars.get(i).copied().filter(|c| matches!(c, '%' | '&' | '^' | '!' | '#' | '@'));
                if suffix.is_some() {
                    i += 1;
                }
                let tok = if is_float || matches!(suffix, Some('!' | '#' | '@')) {
                    let v: f64 = text.parse().map_err(|_| syntax(line, "bad number"))?;
                    match suffix {
                        Some(s @ ('!' | '#' | '@')) => Tok::TypedFloat(v, s),
                        _ => Tok::Float(v),
                    }
                } else {
                    match text.parse::<i64>() {
                        Ok(v) => match suffix {
                            Some(s) => Tok::TypedInteger(v, s),
                            None => Tok::Integer(v),
                        },
                        Err(_) => Tok::Float(text.parse().map_err(|_| syntax(line, "bad number"))?),
                    }
                };
                out.push(Token { tok, line });
            }
            c if c.is_alphabetic() || c == '_' => {
                let start = i;
                while i < n && (chars[i].is_alphanumeric() || chars[i] == '_') {
                    i += 1;
                }
                let word: String = chars[start..i].iter().collect();
                if word.eq_ignore_ascii_case("rem") && at_statement_start(&out) {
                    while i < n && chars[i] != '\n' {
                        i += 1;
                    }
                    continue;
                }
                // Type character directly after the name (not `&` followed by an operand).
                if let Some(&t) = chars.get(i) {
                    if matches!(t, '$' | '%' | '!' | '#' | '@')
                        || (t == '&' && !chars.get(i + 1).is_some_and(|c| c.is_alphanumeric() || *c == '"'))
                    {
                        if !(t == '!' && chars.get(i + 1).is_some_and(|c| c.is_alphabetic() || *c == '[')) {
                            i += 1;
                            out.push(Token { tok: Tok::TypedIdent(word, t), line });
                            continue;
                        }
                    }
                }
                out.push(Token { tok: Tok::Ident(word), line });
            }
            '[' => {
                // Bracketed identifier: [A1], [Sheet Name]. the spreadsheet host evaluates [A1] as a range.
                let start = i + 1;
                while i < n && chars[i] != ']' && chars[i] != '\n' {
                    i += 1;
                }
                let word: String = chars[start..i.min(n)].iter().collect();
                i += 1;
                out.push(Token { tok: Tok::Ident(format!("[{word}]")), line });
            }
            _ => {
                let two: String = chars[i..(i + 2).min(n)].iter().collect();
                let sym = match two.as_str() {
                    "<=" | "=<" => Some("<="),
                    ">=" | "=>" => Some(">="),
                    "<>" | "><" => Some("<>"),
                    ":=" => Some(":="),
                    _ => None,
                };
                if let Some(s) = sym {
                    out.push(Token { tok: Tok::Sym(s), line });
                    i += 2;
                    continue;
                }
                let s = match c {
                    '(' => "(",
                    ')' => ")",
                    ',' => ",",
                    '.' => ".",
                    '!' => "!",
                    ':' => ":",
                    '=' => "=",
                    '<' => "<",
                    '>' => ">",
                    '+' => "+",
                    '-' => "-",
                    '*' => "*",
                    '/' => "/",
                    '\\' => "\\",
                    '^' => "^",
                    '&' => "&",
                    ';' => ";",
                    '#' => "#",
                    _ => return Err(syntax(line, &format!("unexpected character {c:?}"))),
                };
                out.push(Token { tok: Tok::Sym(s), line });
                i += 1;
            }
        }
    }
    push_newline(&mut out, line);
    Ok(out)
}

fn push_newline(out: &mut Vec<Token>, line: u32) {
    if !matches!(out.last().map(|t| &t.tok), None | Some(Tok::Newline)) {
        out.push(Token { tok: Tok::Newline, line });
    }
}

/// ` _` followed only by spaces up to the end of the line.
fn is_continuation(chars: &[char], i: usize) -> bool {
    let before_ok = i == 0 || matches!(chars[i - 1], ' ' | '\t');
    if !before_ok {
        return false;
    }
    let mut j = i + 1;
    while j < chars.len() && matches!(chars[j], ' ' | '\t' | '\r') {
        j += 1;
    }
    j >= chars.len() || chars[j] == '\n'
}

/// True when the previous token ends an operand, so `#` / `.` / `!` bind to it.
fn prev_is_operand(out: &[Token]) -> bool {
    matches!(
        out.last().map(|t| &t.tok),
        Some(
            Tok::Ident(_)
                | Tok::TypedIdent(..)
                | Tok::Integer(_)
                | Tok::TypedInteger(..)
                | Tok::Float(_)
                | Tok::TypedFloat(..)
                | Tok::Str(_)
                | Tok::Date(_)
                | Tok::Sym(")")
        )
    )
}

fn syntax(line: u32, what: &str) -> RuntimeError {
    RuntimeError::new(1002, format!("Syntax error on line {line}: {what}"))
}

/// `#10/14/2026#`, `#2026-10-14#`, `#1:30:00 PM#`, `#10/14/2026 13:30#`, `#October 14, 2026#`.
pub fn parse_date_literal(body: &str) -> Option<f64> {
    use crate::value::ymd_to_serial;
    let body = body.trim();
    let (date_part, time_part) = split_date_time(body);
    let mut serial = 0.0;
    if let Some(d) = date_part {
        let parts: Vec<&str> = d.split(['/', '-']).map(str::trim).collect();
        let nums: Option<Vec<i64>> = parts.iter().map(|p| p.parse().ok()).collect();
        serial = match (nums, parts.len()) {
            (Some(n), 3) if n[0] > 31 => ymd_to_serial(n[0], n[1], n[2]),
            (Some(n), 3) => ymd_to_serial(
                if n[2] < 100 { if n[2] < 30 { 2000 + n[2] } else { 1900 + n[2] } } else { n[2] },
                n[0],
                n[1],
            ),
            (Some(n), 2) => ymd_to_serial(2026, n[0], n[1]),
            _ => {
                // Month name forms: "October 14, 2026" / "14 Oct 2026".
                let words: Vec<String> = d.replace(',', " ").split_whitespace().map(str::to_ascii_lowercase).collect();
                let month = words.iter().find_map(|w| month_number(w))?;
                let nums: Vec<i64> = words.iter().filter_map(|w| w.parse().ok()).collect();
                let (day, year) = match nums.as_slice() {
                    [a, b] if *a > 31 => (*b, *a),
                    [a, b] => (*a, *b),
                    _ => return None,
                };
                ymd_to_serial(year, month, day)
            }
        };
    }
    if let Some(t) = time_part {
        serial += parse_time(t)?;
    }
    Some(serial)
}

fn split_date_time(body: &str) -> (Option<&str>, Option<&str>) {
    let words: Vec<&str> = body.split_whitespace().collect();
    match words.iter().position(|w| w.contains(':')) {
        None => (Some(body), None),
        Some(0) => (None, Some(body)),
        Some(t) => (
            Some(body[..body.find(words[t]).unwrap_or(0)].trim()),
            Some(body[body.find(words[t]).unwrap_or(0)..].trim()),
        ),
    }
}

fn parse_time(t: &str) -> Option<f64> {
    let lower = t.to_ascii_lowercase();
    let (clock, pm, am) = if let Some(c) = lower.strip_suffix("pm") {
        (c.trim(), true, false)
    } else if let Some(c) = lower.strip_suffix("am") {
        (c.trim(), false, true)
    } else {
        (lower.trim(), false, false)
    };
    let parts: Vec<i64> = clock.split(':').map(|p| p.trim().parse().ok()).collect::<Option<Vec<_>>>()?;
    let (mut h, m, s) = match parts.as_slice() {
        [h, m] => (*h, *m, 0),
        [h, m, s] => (*h, *m, *s),
        [h] => (*h, 0, 0),
        _ => return None,
    };
    if pm && h < 12 {
        h += 12;
    }
    if am && h == 12 {
        h = 0;
    }
    Some((h * 3600 + m * 60 + s) as f64 / 86_400.0)
}

fn month_number(w: &str) -> Option<i64> {
    const NAMES: [&str; 12] = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    NAMES.iter().position(|m| w.starts_with(m)).map(|p| p as i64 + 1)
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;

    fn toks(s: &str) -> Vec<Tok> {
        lex(s).unwrap().into_iter().map(|t| t.tok).collect()
    }

    #[test]
    fn basics() {
        assert_eq!(
            toks("x = 1 + 2.5 ' comment\nRem whole line\nMsgBox \"a\"\"b\""),
            vec![
                Tok::Ident("x".into()),
                Tok::Sym("="),
                Tok::Integer(1),
                Tok::Sym("+"),
                Tok::Float(2.5),
                Tok::Newline,
                Tok::Ident("MsgBox".into()),
                Tok::Str("a\"b".into()),
                Tok::Newline
            ]
        );
    }

    #[test]
    fn continuation_and_types() {
        assert_eq!(toks("a = b & _\n   c").len(), 6);
        assert_eq!(toks("n& = &HFF&")[0], Tok::TypedIdent("n".into(), '&'));
        assert_eq!(toks("x = &HFFFF")[2], Tok::TypedInteger(-1, '%'));
        assert_eq!(toks("x = 300&")[2], Tok::TypedInteger(300, '&'));
        assert_eq!(toks("s$ = \"\"")[0], Tok::TypedIdent("s".into(), '$'));
        assert!(matches!(toks("a = b&c")[3], Tok::Sym("&")));
    }

    #[test]
    fn dates() {
        assert_eq!(toks("d = #10/14/2026#")[2], Tok::Date(46309.0));
        assert_eq!(parse_date_literal("1:30:00 PM"), Some(13.5 / 24.0));
        assert_eq!(parse_date_literal("2026-10-14"), Some(46309.0));
        assert_eq!(parse_date_literal("October 14, 2026"), Some(46309.0));
    }

    #[test]
    fn bang_and_brackets() {
        let t = toks("x = Sheets!Sheet1\ny = [A1]");
        assert!(t.contains(&Tok::Sym("!")));
        assert!(t.contains(&Tok::Ident("[A1]".into())));
    }
}
