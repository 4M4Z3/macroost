//! Parses every .bas file in a directory and reports syntax errors (corpus check).
use macroost_lang::{ast::ModuleType, parser::parse_module};
fn main() {
    let dir = std::env::args().nth(1).unwrap_or_else(|| ".".into());
    let (mut ok, mut bad) = (0, 0);
    let mut seen = std::collections::HashSet::new();
    let mut entries: Vec<_> = std::fs::read_dir(&dir).into_iter().flatten().flatten().map(|e| e.path()).collect();
    entries.sort();
    for path in entries {
        let Ok(src) = std::fs::read_to_string(&path) else { continue };
        if !seen.insert(src.clone()) {
            continue;
        }
        match parse_module("M", ModuleType::Standard, &src) {
            Ok(_) => ok += 1,
            Err(e) => {
                bad += 1;
                println!("{}: {}", path.file_name().map(|f| f.to_string_lossy()).unwrap_or_default(), e.description);
            }
        }
    }
    println!("== unique modules parsed: {ok} ok, {bad} failed");
}
