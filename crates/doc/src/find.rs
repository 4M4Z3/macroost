//! `Find` and `Replacement`: plain-text search with the special codes (`^p`, `^t`, `^?`, ...),
//! case and whole-word matching, and Replace one/all with the case of the found text carried
//! over to the replacement when the search ignores case.

use std::cell::RefCell;
use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::state::{FindState, LiveSpan, Shared, chars, data, opt_bool, opt_i64, opt_text, with_mut};

pub(crate) struct FindObj {
    pub app: Shared,
    pub doc: u32,
    /// The range (or selection) the Find belongs to; a successful search moves it.
    pub owner: LiveSpan,
    pub state: Rc<RefCell<FindState>>,
}

pub(crate) struct ReplacementObj {
    pub state: Rc<RefCell<FindState>>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Pat {
    Ch(char),
    Any,
    Digit,
    Letter,
    White,
}

fn special(code: char) -> Option<char> {
    Some(match code {
        'p' | 'P' => '\r',
        't' | 'T' => '\t',
        '^' => '^',
        'l' | 'L' => '\u{b}',
        'm' | 'M' => '\u{c}',
        'n' | 'N' => '\u{e}',
        's' | 'S' => '\u{a0}',
        '~' => '\u{1e}',
        '-' => '\u{1f}',
        _ => return None,
    })
}

/// Reads `^nnn` (a decimal character code) at the start of `rest`.
fn char_code(rest: &[char]) -> Option<(char, usize)> {
    let digits: String = rest.iter().take(3).take_while(|c| c.is_ascii_digit()).collect();
    if digits.is_empty() {
        return None;
    }
    let n: u32 = digits.parse().ok()?;
    Some((char::from_u32(n)?, digits.len()))
}

fn parse_pattern(text: &str) -> Result<Vec<Pat>> {
    let c: Vec<char> = text.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < c.len() {
        if c[i] == '^' && i + 1 < c.len() {
            let code = c[i + 1];
            let p = match code {
                '?' => Pat::Any,
                '#' => Pat::Digit,
                '$' => Pat::Letter,
                'w' | 'W' => Pat::White,
                _ => {
                    if let Some((ch, n)) = char_code(&c[i + 1..]) {
                        out.push(Pat::Ch(ch));
                        i += 1 + n;
                        continue;
                    }
                    Pat::Ch(special(code).ok_or_else(|| {
                        RuntimeError::new(
                            5692,
                            "The Find What text contains a Pattern Match expression which is not valid.",
                        )
                    })?)
                }
            };
            out.push(p);
            i += 2;
            continue;
        }
        out.push(Pat::Ch(c[i]));
        i += 1;
    }
    Ok(out)
}

/// The replacement text with its special codes expanded; `^&` is the found text.
fn expand_replacement(text: &str, found: &str) -> String {
    let c: Vec<char> = text.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < c.len() {
        if c[i] == '^' && i + 1 < c.len() {
            if c[i + 1] == '&' {
                out.push_str(found);
                i += 2;
                continue;
            }
            if let Some((ch, n)) = char_code(&c[i + 1..]) {
                out.push(ch);
                i += 1 + n;
                continue;
            }
            if let Some(ch) = special(c[i + 1]) {
                out.push(ch);
                i += 2;
                continue;
            }
        }
        out.push(c[i]);
        i += 1;
    }
    out
}

fn lower(c: char) -> char {
    c.to_lowercase().next().unwrap_or(c)
}

fn matches_at(hay: &[char], at: usize, pat: &[Pat], match_case: bool) -> bool {
    if at + pat.len() > hay.len() {
        return false;
    }
    pat.iter().zip(&hay[at..]).all(|(p, &h)| match p {
        Pat::Ch(c) => {
            if *c == '\r' {
                h == '\r'
            } else if match_case {
                *c == h
            } else {
                lower(*c) == lower(h)
            }
        }
        Pat::Any => h != '\r' && h != '\u{7}',
        Pat::Digit => h.is_ascii_digit(),
        Pat::Letter => h.is_alphabetic(),
        Pat::White => h == ' ' || h == '\t' || h == '\u{a0}',
    })
}

fn word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

/// The first (or, backward, the last) match lying inside `from..to`.
fn search(hay: &[char], pat: &[Pat], st: &FindState, from: usize, to: usize, forward: bool) -> Option<(usize, usize)> {
    if pat.is_empty() || to < from || to - from < pat.len() {
        return None;
    }
    let ok = |i: usize| {
        if !matches_at(hay, i, pat, st.match_case) {
            return false;
        }
        if st.whole_word {
            let before = i > 0 && word_char(hay[i - 1]);
            let after = i + pat.len() < hay.len() && word_char(hay[i + pat.len()]);
            if before || after {
                return false;
            }
        }
        true
    };
    let last = to - pat.len();
    if forward { (from..=last).find(|i| ok(*i)) } else { (from..=last).rev().find(|i| ok(*i)) }
        .map(|i| (i, i + pat.len()))
}

/// Carries the case of the found text over to the replacement: all capitals stay capitals, a
/// leading capital stays a leading capital.
fn adapt_case(found: &str, replacement: &str) -> String {
    let letters: Vec<char> = found.chars().filter(|c| c.is_alphabetic()).collect();
    if letters.is_empty() {
        return replacement.to_owned();
    }
    if letters.len() >= 2 && letters.iter().all(|c| c.is_uppercase()) {
        return replacement.to_uppercase();
    }
    if found.chars().next().is_some_and(char::is_uppercase) {
        let mut c = replacement.chars();
        return match c.next() {
            Some(f) => f.to_uppercase().chain(c).collect(),
            None => String::new(),
        };
    }
    replacement.to_owned()
}

impl FindObj {
    fn execute(&self, args: &Args) -> Result<Value> {
        {
            let mut st = self.state.borrow_mut();
            if let Some(t) = opt_text(args, 0, "FindText")? {
                st.text = t;
            }
            if let Some(b) = opt_bool(args, 1, "MatchCase")? {
                st.match_case = b;
            }
            if let Some(b) = opt_bool(args, 2, "MatchWholeWord")? {
                st.whole_word = b;
            }
            if let Some(b) = opt_bool(args, 3, "MatchWildcards")? {
                st.wildcards = b;
            }
            if let Some(b) = opt_bool(args, 6, "Forward")? {
                st.forward = b;
            }
            if let Some(w) = opt_i64(args, 7, "Wrap")? {
                st.wrap = w;
            }
            if let Some(b) = opt_bool(args, 8, "Format")? {
                st.format = b;
            }
            if let Some(t) = opt_text(args, 9, "ReplaceWith")? {
                st.replacement = t;
            }
        }
        let replace = opt_i64(args, 10, "Replace")?.unwrap_or(0);
        let fs = self.state.borrow().clone();
        if fs.wildcards {
            return Err(RuntimeError::new(5560, "Pattern-matching (wildcard) searches are not supported yet."));
        }
        let pat = parse_pattern(&fs.text)?;
        if pat.is_empty() {
            // ORACLE: an empty search (no text, no formatting) reports success and moves nothing.
            self.state.borrow_mut().found = true;
            return Ok(Value::Boolean(true));
        }
        let owner = *self.owner.borrow();
        let found = with_mut(&self.app, |stt| {
            let e = stt.entry_mut(self.doc)?;
            let len = e.doc.len();
            let (s, t) = (owner.start.min(len), owner.end.min(len));
            let collapsed = s == t;
            let (from, to) = if collapsed { if fs.forward { (s, len) } else { (0, s) } } else { (s, t) };
            let hay = chars(e.doc.as_ref());
            if replace == 2 {
                let (mut p, mut to, mut any) = (from, to, false);
                while let Some((a, b)) = search(&hay_of(e), &pat, &fs, p, to, true) {
                    any = true;
                    let found_text = crate::state::text(e.doc.as_ref(), a, b);
                    let held_final = b >= e.doc.len();
                    let r = self.replacement_for(&fs, &found_text);
                    let font = e.doc.item(a).map(|i| i.font.clone());
                    let before = e.doc.len();
                    let (ns, ne) = e.replace(a, b, &r, font);
                    let n = ne - ns - usize::from(held_final);
                    let after = e.doc.len();
                    to = (to + after).saturating_sub(before);
                    p = ns + n;
                    if held_final || (a == b) {
                        break;
                    }
                }
                // A non-empty range grows or shrinks to cover the replaced text.
                return Ok(any.then_some(if collapsed { (owner.start, owner.end) } else { (from, to) }));
            }
            let mut hit = search(&hay, &pat, &fs, from, to, fs.forward);
            if hit.is_none() && fs.wrap != 0 {
                // wdFindContinue: past the end of the range to the end of the story, then from
                // the start of the story.
                hit = if fs.forward {
                    search(&hay, &pat, &fs, to, len, true).or_else(|| search(&hay, &pat, &fs, 0, from, true))
                } else {
                    search(&hay, &pat, &fs, 0, from, false).or_else(|| search(&hay, &pat, &fs, to, len, false))
                };
            }
            let Some((a, b)) = hit else { return Ok(None) };
            if replace == 1 {
                let found_text = crate::state::text(e.doc.as_ref(), a, b);
                let r = self.replacement_for(&fs, &found_text);
                let font = e.doc.item(a).map(|i| i.font.clone());
                return Ok(Some(e.replace(a, b, &r, font)));
            }
            Ok(Some((a, b)))
        })?;
        self.state.borrow_mut().found = found.is_some();
        if let Some((a, b)) = found {
            let mut o = self.owner.borrow_mut();
            o.start = a;
            o.end = b;
        }
        Ok(Value::Boolean(found.is_some()))
    }

    fn replacement_for(&self, fs: &FindState, found: &str) -> String {
        let r = expand_replacement(&fs.replacement, found);
        if fs.match_case { r } else { adapt_case(found, &r) }
    }
}

fn hay_of(e: &crate::state::DocEntry) -> Vec<char> {
    chars(e.doc.as_ref())
}

impl Object for FindObj {
    fn type_name(&self) -> String {
        "Find".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let st = || self.state.borrow();
        match name.to_ascii_lowercase().as_str() {
            "execute" => self.execute(args),
            "text" => Ok(Value::from(st().text.as_str())),
            "replacement" => Ok(Value::Object(Some(Rc::new(ReplacementObj { state: self.state.clone() })))),
            "forward" => Ok(Value::Boolean(st().forward)),
            "wrap" => Ok(Value::Long(st().wrap as i32)),
            "matchcase" => Ok(Value::Boolean(st().match_case)),
            "matchwholeword" => Ok(Value::Boolean(st().whole_word)),
            "matchwildcards" => Ok(Value::Boolean(st().wildcards)),
            "matchsoundslike" | "matchallwordforms" | "matchprefix" | "matchsuffix" => Ok(Value::Boolean(false)),
            "format" => Ok(Value::Boolean(st().format)),
            "found" => Ok(Value::Boolean(st().found)),
            "clearformatting" | "clearallfuzzyoptions" => {
                self.state.borrow_mut().format = false;
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let v = data(&value)?;
        let mut st = self.state.borrow_mut();
        match name.to_ascii_lowercase().as_str() {
            "text" => st.text = v.to_text()?,
            "forward" => st.forward = v.to_bool()?,
            "wrap" => st.wrap = v.to_i64()?,
            "matchcase" => st.match_case = v.to_bool()?,
            "matchwholeword" => st.whole_word = v.to_bool()?,
            "matchwildcards" => st.wildcards = v.to_bool()?,
            "format" => st.format = v.to_bool()?,
            "matchsoundslike" | "matchallwordforms" | "matchprefix" | "matchsuffix" | "matchbyte" | "ignorespace"
            | "ignorepunct" => {}
            _ => return Err(RuntimeError::no_member(name)),
        }
        Ok(())
    }
}

impl Object for ReplacementObj {
    fn type_name(&self) -> String {
        "Replacement".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "text" => Ok(Value::from(self.state.borrow().replacement.as_str())),
            "clearformatting" => Ok(Value::Empty),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "text" => {
                self.state.borrow_mut().replacement = data(&value)?.to_text()?;
                Ok(())
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn case_carries_over() {
        assert_eq!(adapt_case("Hello", "bye"), "Bye");
        assert_eq!(adapt_case("HELLO", "bye"), "BYE");
        assert_eq!(adapt_case("hello", "Bye"), "Bye");
        assert_eq!(adapt_case("an", "AN"), "AN");
    }

    #[test]
    fn special_codes() {
        assert_eq!(
            parse_pattern("a^pb^t^^").unwrap(),
            vec![Pat::Ch('a'), Pat::Ch('\r'), Pat::Ch('b'), Pat::Ch('\t'), Pat::Ch('^')]
        );
        assert_eq!(expand_replacement("[^&]^p", "x"), "[x]\r");
        assert_eq!(parse_pattern("^13").unwrap(), vec![Pat::Ch('\r')]);
    }
}
