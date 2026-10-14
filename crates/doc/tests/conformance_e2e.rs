//! End to end against the real host: conformance/cases/doc/*.bas run through the parser, the
//! interpreter and this object model over an in-memory document, compared with the host's
//! recorded results (conformance/expected/doc.json). Each case runs in a fresh host, as the oracle
//! reopens the document before every macro.
#![allow(clippy::unwrap_used, clippy::panic, clippy::expect_used)]

use macroost_doc::{DocHost, MemoryDocument};
use macroost_lang::Interpreter;
use macroost_lang::ast::ModuleType;

fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

/// A minimal JSON reader: enough for the oracle's result files.
#[derive(Debug, Clone)]
enum Json {
    Null,
    Bool(bool),
    Num(String),
    Str(String),
    Arr(Vec<Json>),
    Obj(Vec<(String, Json)>),
}

impl Json {
    fn get(&self, key: &str) -> Option<&Json> {
        match self {
            Json::Obj(v) => v.iter().find(|(k, _)| k == key).map(|(_, j)| j),
            _ => None,
        }
    }
    fn text(&self) -> String {
        match self {
            Json::Null => String::new(),
            Json::Bool(b) => if *b { "True" } else { "False" }.to_owned(),
            Json::Num(n) => n.clone(),
            Json::Str(s) => s.clone(),
            _ => "?".to_owned(),
        }
    }
}

struct Parser<'a> {
    s: &'a [u8],
    i: usize,
}

impl Parser<'_> {
    fn ws(&mut self) {
        while self.i < self.s.len() && self.s[self.i].is_ascii_whitespace() {
            self.i += 1;
        }
    }
    fn value(&mut self) -> Json {
        self.ws();
        match self.s[self.i] {
            b'{' => {
                self.i += 1;
                let mut v = Vec::new();
                loop {
                    self.ws();
                    if self.s[self.i] == b'}' {
                        self.i += 1;
                        break;
                    }
                    let Json::Str(k) = self.value() else { panic!("key") };
                    self.ws();
                    assert_eq!(self.s[self.i], b':');
                    self.i += 1;
                    let val = self.value();
                    v.push((k, val));
                    self.ws();
                    if self.s[self.i] == b',' {
                        self.i += 1;
                    }
                }
                Json::Obj(v)
            }
            b'[' => {
                self.i += 1;
                let mut v = Vec::new();
                loop {
                    self.ws();
                    if self.s[self.i] == b']' {
                        self.i += 1;
                        break;
                    }
                    v.push(self.value());
                    self.ws();
                    if self.s[self.i] == b',' {
                        self.i += 1;
                    }
                }
                Json::Arr(v)
            }
            b'"' => {
                self.i += 1;
                let mut out: Vec<u16> = Vec::new();
                let mut bytes = Vec::new();
                loop {
                    let c = self.s[self.i];
                    self.i += 1;
                    match c {
                        b'"' => break,
                        b'\\' => {
                            let e = self.s[self.i];
                            self.i += 1;
                            let ch = match e {
                                b'n' => '\n' as u16,
                                b'r' => '\r' as u16,
                                b't' => '\t' as u16,
                                b'b' => 8,
                                b'f' => 12,
                                b'u' => {
                                    let h = std::str::from_utf8(&self.s[self.i..self.i + 4]).unwrap();
                                    self.i += 4;
                                    u16::from_str_radix(h, 16).unwrap()
                                }
                                other => u16::from(other),
                            };
                            out.extend(String::from_utf8(std::mem::take(&mut bytes)).unwrap().encode_utf16());
                            out.push(ch);
                        }
                        _ => bytes.push(c),
                    }
                }
                out.extend(String::from_utf8(bytes).unwrap().encode_utf16());
                Json::Str(String::from_utf16(&out).unwrap())
            }
            b't' => {
                self.i += 4;
                Json::Bool(true)
            }
            b'f' => {
                self.i += 5;
                Json::Bool(false)
            }
            b'n' => {
                self.i += 4;
                Json::Null
            }
            _ => {
                let start = self.i;
                while self.i < self.s.len() && matches!(self.s[self.i], b'-' | b'+' | b'.' | b'e' | b'E' | b'0'..=b'9')
                {
                    self.i += 1;
                }
                Json::Num(std::str::from_utf8(&self.s[start..self.i]).unwrap().to_owned())
            }
        }
    }
}

fn expectations(suite: &str) -> Vec<(String, String)> {
    let text = std::fs::read_to_string(root().join(format!("conformance/expected/{suite}.json"))).unwrap();
    let text = text.trim_start_matches('\u{feff}');
    let Json::Arr(items) = (Parser { s: text.as_bytes(), i: 0 }).value() else { panic!("array expected") };
    items
        .iter()
        .map(|r| {
            let id = r.get("macro").unwrap().text();
            let id = id.rsplit('.').next().unwrap().to_owned();
            let exp = if r.get("compile_error").is_some() {
                "COMPILE".to_owned()
            } else if let Some(e) = r.get("runtime_error") {
                format!("ERR {}", e.get("number").unwrap().text())
            } else {
                r.get("value").map(Json::text).unwrap_or_default()
            };
            (id, exp)
        })
        .collect()
}

fn show(s: &str) -> String {
    s.replace('\r', "\\r").replace('\u{7}', "\\a")
}

#[test]
fn word_cases_match_the_host_end_to_end() {
    let dir = root().join("conformance/cases/doc");
    let mut modules = Vec::new();
    let mut paths: Vec<_> = std::fs::read_dir(&dir).unwrap().flatten().map(|e| e.path()).collect();
    paths.sort();
    for p in paths {
        if p.extension().is_some_and(|e| e == "bas") {
            let src: String = std::fs::read(&p).unwrap().iter().map(|&b| char::from(b)).collect();
            let name = p.file_stem().unwrap().to_string_lossy().into_owned();
            modules.push((name, src));
        }
    }
    let only = std::env::var("WORD_CASE").ok();
    let (mut pass, mut fail) = (0, Vec::new());
    for (id, expected) in expectations("doc") {
        if only.as_deref().is_some_and(|o| o != id) {
            continue;
        }
        let mut host = DocHost::new(Box::new(MemoryDocument::new("doc.docm")));
        let got = {
            let mut it = Interpreter::new(&mut host);
            for (name, src) in &modules {
                it.add_module(name, ModuleType::Standard, src).unwrap();
            }
            match it.run(&id, vec![]) {
                Ok(v) => v.to_text().unwrap_or_else(|e| format!("ERR {}", e.number)),
                Err(e) => format!("ERR {} {}", e.number, e.description),
            }
        };
        let got_cmp = if expected.starts_with("ERR ") {
            got.split(' ').take(2).collect::<Vec<_>>().join(" ")
        } else {
            got.clone()
        };
        if got_cmp == expected {
            pass += 1;
        } else {
            fail.push(format!("{id}: host={}  ours={}", show(&expected), show(&got)));
        }
    }
    println!("CONFORMANCE word (end to end): {pass}/{} match", pass + fail.len());
    for f in &fail {
        println!("  {f}");
    }
    let floor = std::env::var("CONFORMANCE_FLOOR_WORD").ok().and_then(|s| s.parse().ok()).unwrap_or(280usize);
    assert!(pass >= floor, "{pass} below floor {floor}");
}

#[test]
fn macro_edits_reach_the_storage() {
    let mut host = DocHost::new(Box::new(MemoryDocument::with_paragraphs("t.docm", &["Hello world"])));
    {
        let mut it = Interpreter::new(&mut host);
        let src = "Sub Go()\n  ActiveDocument.Content.Find.Execute FindText:=\"world\", ReplaceWith:=\"there\", Replace:=wdReplaceAll\n  ActiveDocument.Words(1).Bold = True\n  ActiveDocument.Bookmarks.Add \"b\", ActiveDocument.Words(2)\nEnd Sub\n";
        it.add_module("M", ModuleType::Standard, src).unwrap();
        it.run("Go", vec![]).unwrap();
    }
    let doc = host.active_document().unwrap();
    let text: String = (0..doc.len()).map(|p| doc.item(p).unwrap().ch).collect();
    assert_eq!(text, "Hello there\r");
    assert!(doc.item(0).unwrap().font.bold);
    assert!(!doc.item(6).unwrap().font.bold);
    assert_eq!(doc.bookmarks()[0].start, 6);
}
