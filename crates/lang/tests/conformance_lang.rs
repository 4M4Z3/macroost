//! Differential checks against the real host: each suite's `.bas` modules (conformance/cases/<suite>)
//! run through our interpreter and are compared, case by case, with the host's recorded results
//! (conformance/expected/<suite>.json, results recorded from the reference implementation).
#![allow(clippy::unwrap_used, clippy::panic, clippy::expect_used)]

use macroost_lang::ast::ModuleType;
use macroost_lang::{Interpreter, NoHost};

fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

/// (procedure, expected) where expected is the value's text, "ERR n" or "COMPILE".
fn expectations(suite: &str) -> Vec<(String, String)> {
    let text = std::fs::read_to_string(root().join(format!("conformance/expected/{suite}.json"))).unwrap();
    let mut out = Vec::new();
    // Minimal JSON walk: the file is an array of flat result objects.
    for obj in text.split("\"macro\":").skip(1) {
        let id = obj.split('"').nth(1).unwrap_or_default().to_owned();
        let exp = if obj.contains("\"compile_error\"") {
            "COMPILE".to_owned()
        } else if let Some(i) = obj.find("\"number\":") {
            let n: String = obj[i + 9..]
                .chars()
                .skip_while(|c| c.is_whitespace())
                .take_while(|c| c.is_ascii_digit() || *c == '-')
                .collect();
            format!("ERR {n}")
        } else {
            let i = obj.find("\"value\":").unwrap();
            let rest = obj[i + 8..].trim_start();
            if rest.starts_with('"') {
                parse_json_string(rest)
            } else {
                rest.chars().take_while(|c| !matches!(c, ',' | '}' | '\n' | '\r')).collect::<String>().trim().to_owned()
            }
        };
        out.push((id, exp));
    }
    out
}

fn parse_json_string(s: &str) -> String {
    let mut out = String::new();
    let mut chars = s.chars().skip(1);
    while let Some(c) = chars.next() {
        match c {
            '"' => break,
            '\\' => match chars.next() {
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('u') => {
                    let h: String = chars.by_ref().take(4).collect();
                    if let Some(ch) = u32::from_str_radix(&h, 16).ok().and_then(char::from_u32) {
                        out.push(ch);
                    }
                }
                Some(o) => out.push(o),
                None => break,
            },
            c => out.push(c),
        }
    }
    out
}

/// Runs one suite; returns (matches, total) and prints every mismatch.
fn check(suite: &str) -> (usize, usize) {
    let dir = root().join(format!("conformance/cases/{suite}"));
    let notes: std::collections::HashMap<String, String> = std::fs::read_to_string(dir.join("expressions.tsv"))
        .unwrap_or_default()
        .lines()
        .filter_map(|l| l.split_once('\t').map(|(a, b)| (a.to_owned(), b.to_owned())))
        .collect();
    let mut host = NoHost::default();
    let mut it = Interpreter::new(&mut host);
    let mut files: Vec<_> = std::fs::read_dir(&dir).unwrap().flatten().map(|e| e.path()).collect();
    files.sort();
    for p in files.iter().filter(|p| p.extension().is_some_and(|e| e == "bas")) {
        // The case files are Windows-1252; Latin-1 decoding is close enough for them.
        let src: String = std::fs::read(p).unwrap().iter().map(|&b| char::from(b)).collect();
        let name = p.file_stem().unwrap().to_string_lossy().into_owned();
        it.add_module(&name, ModuleType::Standard, &src).unwrap();
    }
    let (mut pass, mut fail) = (0, Vec::new());
    for (id, expected) in expectations(suite) {
        if expected == "COMPILE" {
            continue;
        }
        let got = match it.run(&id, vec![]) {
            Ok(v) => v.to_text().unwrap_or_else(|e| format!("ERR {}", e.number)),
            Err(e) => format!("ERR {}", e.number),
        };
        if got == expected {
            pass += 1;
        } else {
            let short = id.rsplit('.').next().unwrap_or(&id).to_owned();
            let note = notes.get(&short).cloned().unwrap_or_default();
            fail.push(format!("{id} {note:40} host={expected:30} ours={got}"));
        }
    }
    let total = pass + fail.len();
    println!("CONFORMANCE {suite}: {pass}/{total} match");
    for f in &fail {
        println!("  {f}");
    }
    (pass, total)
}

/// Ratchet: a suite's floor only ever goes up. Override with CONFORMANCE_FLOOR_<SUITE>.
fn floor(suite: &str, default: usize) -> usize {
    std::env::var(format!("CONFORMANCE_FLOOR_{}", suite.to_ascii_uppercase()))
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(default)
}

#[test]
fn language_cases_match_the_host() {
    let (pass, _) = check("lang");
    assert!(pass >= floor("lang", 194), "lang: {pass} below floor");
}

#[test]
fn program_cases_match_the_host() {
    let (pass, _) = check("lang2");
    assert!(pass >= floor("lang2", 39), "lang2: {pass} below floor");
}
