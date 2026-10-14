//! End to end against the real host: every `conformance/cases/sheet*/*.bas` suite runs through the
//! parser, the interpreter and this object model over an in-memory workbook, compared with the
//! host's recorded results (`conformance/expected/<suite>.json`). Each suite gets a fresh host, like
//! the oracle's fresh application instance per workbook, and its cases run in recorded order.
#![allow(clippy::unwrap_used, clippy::panic, clippy::expect_used)]

use macroost_lang::Interpreter;
use macroost_lang::ast::ModuleType;
use macroost_sheet::{MemoryWorkbook, SheetHost};

fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

/// A minimal JSON reader (enough for the recorded results).
#[derive(Debug, Clone)]
enum J {
    Null,
    Bool(bool),
    Num(String),
    Str(String),
    Arr(Vec<J>),
    Obj(Vec<(String, J)>),
}

impl J {
    fn get(&self, key: &str) -> Option<&J> {
        match self {
            J::Obj(m) => m.iter().find(|(k, _)| k == key).map(|(_, v)| v),
            _ => None,
        }
    }
    fn text(&self) -> String {
        match self {
            J::Null => String::new(),
            J::Bool(b) => if *b { "True" } else { "False" }.to_owned(),
            J::Num(n) => n.clone(),
            J::Str(s) => s.clone(),
            J::Arr(_) | J::Obj(_) => "<complex>".to_owned(),
        }
    }
}

struct P<'a> {
    s: &'a [char],
    i: usize,
}

impl P<'_> {
    fn ws(&mut self) {
        while self.i < self.s.len() && self.s[self.i].is_whitespace() {
            self.i += 1;
        }
    }
    fn value(&mut self) -> J {
        self.ws();
        match self.s[self.i] {
            '{' => {
                self.i += 1;
                let mut m = Vec::new();
                loop {
                    self.ws();
                    if self.s[self.i] == '}' {
                        self.i += 1;
                        break;
                    }
                    let J::Str(k) = self.value() else { panic!("key") };
                    self.ws();
                    assert_eq!(self.s[self.i], ':');
                    self.i += 1;
                    let v = self.value();
                    m.push((k, v));
                    self.ws();
                    if self.s[self.i] == ',' {
                        self.i += 1;
                    }
                }
                J::Obj(m)
            }
            '[' => {
                self.i += 1;
                let mut v = Vec::new();
                loop {
                    self.ws();
                    if self.s[self.i] == ']' {
                        self.i += 1;
                        break;
                    }
                    v.push(self.value());
                    self.ws();
                    if self.s[self.i] == ',' {
                        self.i += 1;
                    }
                }
                J::Arr(v)
            }
            '"' => {
                self.i += 1;
                let mut out = String::new();
                loop {
                    let c = self.s[self.i];
                    self.i += 1;
                    match c {
                        '"' => break,
                        '\\' => {
                            let e = self.s[self.i];
                            self.i += 1;
                            match e {
                                'n' => out.push('\n'),
                                'r' => out.push('\r'),
                                't' => out.push('\t'),
                                'b' => out.push('\u{8}'),
                                'f' => out.push('\u{c}'),
                                'u' => {
                                    let hex: String = self.s[self.i..self.i + 4].iter().collect();
                                    self.i += 4;
                                    let mut code = u32::from_str_radix(&hex, 16).unwrap();
                                    if (0xD800..0xDC00).contains(&code) && self.s.get(self.i) == Some(&'\\') {
                                        let lo: String = self.s[self.i + 2..self.i + 6].iter().collect();
                                        self.i += 6;
                                        let lo = u32::from_str_radix(&lo, 16).unwrap();
                                        code = 0x10000 + ((code - 0xD800) << 10) + (lo - 0xDC00);
                                    }
                                    out.push(char::from_u32(code).unwrap_or('\u{fffd}'));
                                }
                                other => out.push(other),
                            }
                        }
                        other => out.push(other),
                    }
                }
                J::Str(out)
            }
            't' => {
                self.i += 4;
                J::Bool(true)
            }
            'f' => {
                self.i += 5;
                J::Bool(false)
            }
            'n' => {
                self.i += 4;
                J::Null
            }
            _ => {
                let st = self.i;
                while self.i < self.s.len()
                    && !matches!(self.s[self.i], ',' | '}' | ']')
                    && !self.s[self.i].is_whitespace()
                {
                    self.i += 1;
                }
                J::Num(self.s[st..self.i].iter().collect())
            }
        }
    }
}

fn parse_json(text: &str) -> J {
    let chars: Vec<char> = text.trim_start_matches('\u{feff}').chars().collect();
    P { s: &chars, i: 0 }.value()
}

/// (macro, expected text) in recorded order. Errors are `ERR <number> <description>` (`ERR
/// <number>` for the first suite, which predates descriptions in the comparison).
fn expectations(suite: &str, with_description: bool) -> Vec<(String, String)> {
    let text = std::fs::read_to_string(root().join(format!("conformance/expected/{suite}.json"))).unwrap();
    let J::Arr(items) = parse_json(&text) else { panic!("results must be an array") };
    items
        .iter()
        .map(|r| {
            let id = r.get("macro").map(J::text).unwrap_or_default();
            let exp = if r.get("compile_error").is_some() {
                "COMPILE".to_owned()
            } else if let Some(e) = r.get("runtime_error") {
                let n = e.get("number").map(J::text).unwrap_or_default();
                if with_description {
                    format!("ERR {n} {}", e.get("description").map(J::text).unwrap_or_default())
                } else {
                    format!("ERR {n}")
                }
            } else {
                match r.get("text") {
                    Some(t) => t.text(),
                    None => r.get("value").map(J::text).unwrap_or_default(),
                }
            };
            (id, exp)
        })
        .collect()
}

/// The oracle harness's application state before any case runs.
const INIT: &str = "Sub OracleInit()\n Application.DisplayAlerts = False\n Application.ScreenUpdating = False\n Application.EnableEvents = False\nEnd Sub\n";

fn run_suite(suite: &str, with_description: bool) -> (usize, Vec<String>) {
    let dir = root().join(format!("conformance/cases/{suite}"));
    let mut book = MemoryWorkbook::new();
    book.set_name(&format!("{suite}.xlsm"));
    let mut host = SheetHost::new(Box::new(book));
    let mut it = Interpreter::new(&mut host);
    let mut paths: Vec<_> = std::fs::read_dir(&dir)
        .unwrap()
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|e| e == "bas"))
        .collect();
    paths.sort();
    for p in paths {
        let src: String = std::fs::read(&p).unwrap().iter().map(|&b| char::from(b)).collect();
        let name = p.file_stem().unwrap().to_string_lossy().into_owned();
        it.add_module(&name, ModuleType::Standard, &src).unwrap_or_else(|e| panic!("{name}: {e:?}"));
    }
    it.add_module("OracleInitModule", ModuleType::Standard, INIT).unwrap();
    it.run("OracleInit", vec![]).unwrap();
    let (mut pass, mut fail) = (0, Vec::new());
    for (id, expected) in expectations(suite, with_description) {
        let got = match it.run(&id, vec![]) {
            Ok(v) => v.to_text().unwrap_or_else(|e| format!("ERR {}", e.number)),
            Err(e) if with_description => format!("ERR {} {}", e.number, e.description),
            Err(e) => format!("ERR {}", e.number),
        };
        if got == expected {
            pass += 1;
        } else {
            fail.push(format!("{id}: host={expected:?}\n      ours={got:?}"));
        }
    }
    println!("CONFORMANCE {suite} (end to end): {pass}/{} match", pass + fail.len());
    for f in &fail {
        println!("  {f}");
    }
    (pass, fail)
}

fn check(suite: &str, with_description: bool, default_floor: usize) {
    let (pass, _) = run_suite(suite, with_description);
    let var = format!("CONFORMANCE_FLOOR_{}", suite.to_ascii_uppercase());
    let floor = std::env::var(var).ok().and_then(|s| s.parse().ok()).unwrap_or(default_floor);
    assert!(pass >= floor, "{suite}: {pass} below floor {floor}");
}

#[test]
fn excel_cases_match_the_host_end_to_end() {
    check("sheet1", false, 25);
}

#[test]
fn excel2_cases_match_the_host_end_to_end() {
    check("sheet2", true, 105);
}

#[test]
fn excel3_cases_match_the_host_end_to_end() {
    check("sheet3", true, 102);
}

#[test]
fn excel4_cases_match_the_host_end_to_end() {
    check("sheet4", true, 120);
}

#[test]
fn excel5_cases_match_the_host_end_to_end() {
    check("sheet5", true, 112);
}
