//! Writes every module's source from the files named on stdin into a directory (argv[1]).
use std::io::BufRead;
fn main() {
    let out = std::path::PathBuf::from(std::env::args().nth(1).unwrap_or_else(|| "dump".into()));
    let _ = std::fs::create_dir_all(&out);
    for (i, line) in std::io::stdin().lock().lines().map_while(Result::ok).enumerate() {
        if let Ok(p) = macroost_project::read_file(std::path::Path::new(line.trim())) {
            for m in p.modules {
                let _ = std::fs::write(out.join(format!("{i:03}_{}.bas", m.name)), m.source);
            }
        }
    }
}
