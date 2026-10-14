//! `macroost`: list, extract and run the macros in macro-enabled documents.
//!
//! ```text
//! macroost list    <file>                 modules and their procedures
//! macroost extract <file> [out-dir]       write each module's source to <out-dir>/<name>.bas
//! macroost run     <file|dir|.bas> <Macro> [args...]
//! ```

use std::path::{Path, PathBuf};
use std::process::ExitCode;

use macroost_lang::ast::ModuleType;
use macroost_lang::{Interpreter, Value};
use macroost_sheet::{MemoryWorkbook, SheetHost};

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let result = match args.first().map(String::as_str) {
        Some("list") if args.len() >= 2 => list(Path::new(&args[1])),
        Some("extract") if args.len() >= 2 => extract(Path::new(&args[1]), args.get(2).map(PathBuf::from)),
        Some("run") if args.len() >= 3 => run(Path::new(&args[1]), &args[2], &args[3..]),
        Some("sweep") if args.len() >= 2 => sweep(Path::new(&args[1])),
        _ => {
            eprintln!(
                "usage:\n  macroost list <file>\n  macroost extract <file> [out-dir]\n  macroost run <file|dir|file.bas> <Macro> [args...]"
            );
            return ExitCode::from(2);
        }
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("macroost: {e}");
            ExitCode::FAILURE
        }
    }
}

type AnyResult<T> = std::result::Result<T, Box<dyn std::error::Error>>;

/// (module name, kind, source) from an macro-enabled document, a directory of .bas/.cls files, or one file.
fn load(path: &Path) -> AnyResult<Vec<(String, ModuleType, String)>> {
    if path.is_dir() {
        let mut out = Vec::new();
        let mut entries: Vec<PathBuf> = std::fs::read_dir(path)?.flatten().map(|e| e.path()).collect();
        entries.sort();
        for p in entries {
            let ext = p.extension().and_then(|e| e.to_str()).unwrap_or_default().to_ascii_lowercase();
            let kind = match ext.as_str() {
                "bas" => ModuleType::Standard,
                "cls" => ModuleType::Class,
                _ => continue,
            };
            let name = p.file_stem().and_then(|s| s.to_str()).unwrap_or("Module").to_owned();
            out.push((name, kind, read_text(&p)?));
        }
        return Ok(out);
    }
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or_default().to_ascii_lowercase();
    if ext == "bas" || ext == "cls" {
        let name = path.file_stem().and_then(|s| s.to_str()).unwrap_or("Module1").to_owned();
        let kind = if ext == "cls" { ModuleType::Class } else { ModuleType::Standard };
        return Ok(vec![(name, kind, read_text(path)?)]);
    }
    let project = macroost_project::read_file(path)?;
    Ok(project
        .modules
        .into_iter()
        .map(|m| {
            let kind = match m.kind {
                macroost_project::ModuleKind::Class | macroost_project::ModuleKind::Form => ModuleType::Class,
                macroost_project::ModuleKind::Document => ModuleType::Document,
                macroost_project::ModuleKind::Standard => ModuleType::Standard,
            };
            (m.name, kind, m.source)
        })
        .collect())
}

/// Source files exported by the editor are Windows-1252; accept UTF-8 too.
fn read_text(p: &Path) -> AnyResult<String> {
    let bytes = std::fs::read(p)?;
    Ok(match String::from_utf8(bytes) {
        Ok(s) => s,
        Err(e) => e.into_bytes().iter().map(|&b| char::from(b)).collect(),
    })
}

fn list(path: &Path) -> AnyResult<()> {
    for (name, kind, source) in load(path)? {
        let lines = source.lines().filter(|l| !l.trim_start().to_ascii_lowercase().starts_with("attribute ")).count();
        println!("{name} ({kind:?}, {lines} lines)");
        match macroost_lang::parser::parse_module(&name, kind, &source) {
            Ok(m) => {
                for p in &m.procs {
                    let params: Vec<&str> = p.params.iter().map(|x| x.name.as_str()).collect();
                    println!("  {:?} {}({})", p.kind, p.name, params.join(", "));
                }
            }
            Err(e) => println!("  (does not parse: {})", e.description),
        }
    }
    Ok(())
}

fn extract(path: &Path, out: Option<PathBuf>) -> AnyResult<()> {
    let out = out.unwrap_or_else(|| PathBuf::from("macros"));
    std::fs::create_dir_all(&out)?;
    for (name, kind, source) in load(path)? {
        let ext = if kind == ModuleType::Class { "cls" } else { "bas" };
        let file = out.join(format!("{name}.{ext}"));
        std::fs::write(&file, source)?;
        println!("{}", file.display());
    }
    Ok(())
}

/// The workbook a macro runs against: the file's own sheets and data for .xlsx/.xlsm, otherwise
/// an empty book.
fn workbook_for(path: &Path) -> MemoryWorkbook {
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or_default().to_ascii_lowercase();
    if matches!(ext.as_str(), "xlsx" | "xlsm" | "xltm" | "xltx") {
        if let Ok(book) = macroost_xlsx::load_file(path) {
            return book;
        }
    }
    MemoryWorkbook::new()
}

fn run(path: &Path, entry: &str, args: &[String]) -> AnyResult<()> {
    // `--save out.xlsm` writes the workbook after the macro ran (the input file is the template).
    let mut save_to = None;
    let mut values = Vec::new();
    let mut it_args = args.iter();
    while let Some(a) = it_args.next() {
        if a == "--save" {
            save_to = it_args.next().map(PathBuf::from);
        } else {
            values.push(Value::str(a.clone()));
        }
    }
    let modules = load(path)?;
    let mut host = SheetHost::new(Box::new(workbook_for(path)));
    let result = {
        let mut it = Interpreter::new(&mut host);
        for (name, kind, source) in &modules {
            it.add_module(name, *kind, source).map_err(|e| format!("{name}: {e}"))?;
        }
        it.run(entry, values)
    };
    for m in &host.messages {
        println!("[MsgBox] {m}");
    }
    for p in &host.printed {
        println!("{p}");
    }
    if let Some(out) = save_to {
        let book = host.into_workbook().map_err(|_| "the workbook is still in use")?;
        let mem = book.as_any().and_then(|a| a.downcast_ref::<MemoryWorkbook>()).ok_or("unexpected workbook type")?;
        let template = std::fs::read(path).ok();
        let bytes = macroost_xlsx::save(mem, template.as_deref())?;
        std::fs::write(&out, bytes)?;
        println!("saved {}", out.display());
    }
    let result = result?;
    if !result.is_empty() {
        println!("=> {}", result.to_text().unwrap_or_else(|_| format!("{result:?}")));
    }
    Ok(())
}

/// Runs every argument-free procedure of every file listed in `list` (one path per line)
/// against a fresh copy of that workbook's data and tallies the outcomes: a robustness check over real macros.
fn sweep(list: &Path) -> AnyResult<()> {
    let mut outcomes: std::collections::BTreeMap<String, usize> = std::collections::BTreeMap::new();
    let mut seen = std::collections::HashSet::new();
    let (mut files, mut procs) = (0, 0);
    for line in std::fs::read_to_string(list)?.lines() {
        let path = Path::new(line.trim());
        let Ok(modules) = load(path) else { continue };
        // Skip workbooks whose macro code we have already swept.
        let fingerprint: String = modules.iter().map(|m| m.2.as_str()).collect();
        if modules.is_empty() || !seen.insert(fingerprint) {
            continue;
        }
        files += 1;
        let book = workbook_for(path);
        let mut parsed = Vec::new();
        for (name, kind, source) in &modules {
            match macroost_lang::parser::parse_module(name, *kind, source) {
                Ok(m) => parsed.push(m),
                Err(e) => *outcomes.entry(format!("PARSE {}", e.description)).or_default() += 1,
            }
        }
        let entries: Vec<String> = parsed
            .iter()
            .filter(|m| m.kind != ModuleType::Class)
            .flat_map(|m| {
                m.procs
                    .iter()
                    .filter(|p| p.params.is_empty() && matches!(p.kind, macroost_lang::ast::ProcKind::Sub))
                    .map(move |p| format!("{}.{}", m.name, p.name))
            })
            .collect();
        for entry in entries {
            procs += 1;
            let mut host = SheetHost::new(Box::new(book.clone()));
            let mut it = Interpreter::new(&mut host);
            it.limits.max_steps = 2_000_000;
            for m in &parsed {
                let _ = it.add_parsed(m.clone());
            }
            let key = match it.run(&entry, Vec::new()) {
                Ok(_) => "ok".to_owned(),
                Err(e) => format!("ERR {} {}", e.number, e.description),
            };
            if std::env::var_os("SWEEP_VERBOSE").is_some() && key != "ok" && !key.starts_with("ERR 9 ") {
                println!("  {entry}: {key}");
            }
            *outcomes.entry(key).or_default() += 1;
        }
    }
    println!("swept {procs} procedures in {files} distinct macro projects");
    let mut rows: Vec<_> = outcomes.into_iter().collect();
    rows.sort_by_key(|r| std::cmp::Reverse(r.1));
    for (k, n) in rows {
        println!("{n:6}  {k}");
    }
    Ok(())
}
