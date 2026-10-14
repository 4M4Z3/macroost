//! End to end against the real host: conformance/cases/slides*/*.bas run through the parser, the
//! interpreter and this object model over an in-memory presentation, compared with the host's
//! recorded results (conformance/expected/slides.json, slides_window.json).
//!
//! The oracle runs every case in a fresh copy of the generated presentation (two blank slides),
//! so each case here gets a fresh host too.
#![allow(clippy::unwrap_used, clippy::panic, clippy::expect_used)]

use macroost_lang::Interpreter;
use macroost_lang::ast::ModuleType;
use macroost_slides::{MemoryPresentation, SlidesHost};

fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

/// Minimal reader for the result files: (macro, expected text), where an error is `ERR <n>`.
fn expectations(file: &str) -> Vec<(String, String)> {
    let text = std::fs::read_to_string(root().join(format!("conformance/expected/{file}"))).unwrap();
    let mut out = Vec::new();
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
            if let Some(body) = rest.strip_prefix('"') {
                let mut s = String::new();
                let mut chars = body.chars();
                while let Some(c) = chars.next() {
                    match c {
                        '"' => break,
                        '\\' => match chars.next() {
                            Some('u') => {
                                let hex: String = chars.by_ref().take(4).collect();
                                s.push(char::from_u32(u32::from_str_radix(&hex, 16).unwrap()).unwrap());
                            }
                            Some('r') => s.push('\r'),
                            Some('n') => s.push('\n'),
                            Some('t') => s.push('\t'),
                            Some(other) => s.push(other),
                            None => {}
                        },
                        c => s.push(c),
                    }
                }
                s
            } else {
                rest.chars().take_while(|c| !matches!(c, ',' | '}' | '\n' | '\r')).collect::<String>().trim().to_owned()
            }
        };
        out.push((id, exp));
    }
    out
}

fn run_suite(dir: &str, results: &str, file_name: &str) -> (usize, Vec<String>) {
    let mut modules = Vec::new();
    for p in std::fs::read_dir(root().join(dir)).unwrap().flatten().map(|e| e.path()) {
        if p.extension().is_some_and(|e| e == "bas") {
            let src: String = std::fs::read(&p).unwrap().iter().map(|&b| char::from(b)).collect();
            modules.push((p.file_stem().unwrap().to_string_lossy().into_owned(), src));
        }
    }
    let (mut pass, mut fail) = (0, Vec::new());
    for (id, expected) in expectations(results) {
        let mut host = SlidesHost::new(Box::new(MemoryPresentation::new(file_name, 2)));
        let mut it = Interpreter::new(&mut host);
        for (name, src) in &modules {
            it.add_module(name, ModuleType::Standard, src).unwrap();
        }
        let got = match it.run(&id, vec![]) {
            Ok(v) => v.to_text().unwrap_or_else(|e| format!("ERR {}", e.number)),
            Err(e) => format!("ERR {} {}", e.number, e.description),
        };
        let ok = got == expected || (expected.starts_with("ERR ") && got.starts_with(&format!("{expected} ")));
        if ok {
            pass += 1;
        } else {
            fail.push(format!("{id}: host={expected}  ours={got}"));
        }
    }
    (pass, fail)
}

#[test]
fn presentation_cases_match_the_host_end_to_end() {
    let mut total_pass = 0;
    let mut total = 0;
    for (dir, results, file) in [
        ("conformance/cases/slides", "slides.json", "slides.pptm"),
        ("conformance/cases/slides_window", "slides_window.json", "slides_window.pptm"),
        ("conformance/cases/slides2", "slides2.json", "slides2.pptm"),
    ] {
        let (pass, fail) = run_suite(dir, results, file);
        println!("CONFORMANCE {results} (end to end): {pass}/{} match", pass + fail.len());
        for f in &fail {
            println!("  {f}");
        }
        total_pass += pass;
        total += pass + fail.len();
    }
    println!("CONFORMANCE slides total: {total_pass}/{total}");
    let floor = std::env::var("CONFORMANCE_FLOOR_SLIDES").ok().and_then(|s| s.parse().ok()).unwrap_or(123usize);
    assert!(total_pass >= floor, "{total_pass} below floor {floor}");
}
