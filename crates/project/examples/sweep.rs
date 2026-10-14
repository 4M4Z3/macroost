//! Reads the macro project of every file named on stdin (one path per line) and prints a summary.
use std::io::BufRead;

fn main() {
    let (mut ok, mut none, mut bad) = (0, 0, 0);
    for line in std::io::stdin().lock().lines().map_while(Result::ok) {
        let path = std::path::Path::new(line.trim());
        match macroost_project::read_file(path) {
            Ok(p) => {
                ok += 1;
                let lines: usize = p.modules.iter().map(|m| m.source.lines().count()).sum();
                let mods: Vec<String> = p
                    .modules
                    .iter()
                    .map(|m| format!("{}({:?},{})", m.name, m.kind, m.source.lines().count()))
                    .collect();
                println!("OK  {} modules, {} lines  {}", p.modules.len(), lines, mods.join(" "));
            }
            Err(macroost_project::Error::NoMacros) => none += 1,
            Err(e) => {
                bad += 1;
                println!("ERR {e}  {}", path.display());
            }
        }
    }
    println!("== {ok} read, {none} without macros, {bad} errors");
}
