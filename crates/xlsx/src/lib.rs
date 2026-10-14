//! Load OOXML spreadsheets (`.xlsx` / `.xlsm`, ECMA-376 / ISO/IEC 29500 SpreadsheetML)
//! into a [`MemoryWorkbook`], and save a [`MemoryWorkbook`] back.
//!
//! Loading reads sheet names, order and visibility, cell values (shared, inline and formula
//! strings, numbers, booleans, errors), formulas (shared formulas expanded per cell) with their
//! cached values, the formatting the workbook model holds (number format, font, fill, borders,
//! horizontal alignment, wrap), column widths, row heights, hidden rows and columns, merged
//! cells, sheet protection, sheet code names and defined names. Everything else (charts,
//! drawings, comments, conditional formats, validation, ...) is ignored on load.
//!
//! Saving with the original file as a template keeps every part it does not regenerate (the
//! macro project, drawings, themes, document properties, ...) byte for byte, and splices the
//! regenerated pieces into the original workbook, worksheet and styles parts so features the
//! model cannot hold survive. Without a template a minimal clean workbook is written.
//!
//! All input is treated as hostile: sizes and counts are bounded and nothing panics.

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use std::io::{Cursor, Read};
use std::path::Path;

pub use macroost_sheet::MemoryWorkbook;

mod protection;
mod read;
mod styles;
mod write;
mod xml;

pub use read::{load, load_file, load_named};
pub use write::save;

/// Largest single part accepted (uncompressed).
pub const MAX_PART_SIZE: u64 = 256 * 1024 * 1024;
/// Largest total uncompressed package accepted.
pub const MAX_TOTAL_SIZE: u64 = 1024 * 1024 * 1024;
/// Most cells accepted across the workbook.
pub const MAX_CELLS: u64 = 50_000_000;
/// Most zip entries accepted.
pub const MAX_ENTRIES: usize = 100_000;

#[derive(Debug)]
pub enum Error {
    Io(std::io::Error),
    Zip(String),
    Xml {
        part: String,
        msg: String,
    },
    Missing(String),
    /// A size or count limit was exceeded.
    Limit(String),
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {e}"),
            Error::Zip(e) => write!(f, "not a readable package: {e}"),
            Error::Xml { part, msg } => write!(f, "{part}: {msg}"),
            Error::Missing(p) => write!(f, "missing part {p}"),
            Error::Limit(m) => write!(f, "limit exceeded: {m}"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(e: std::io::Error) -> Self {
        Error::Io(e)
    }
}

impl From<zip::result::ZipError> for Error {
    fn from(e: zip::result::ZipError) -> Self {
        Error::Zip(e.to_string())
    }
}

pub(crate) fn xml_err(part: &str, msg: impl Into<String>) -> Error {
    Error::Xml { part: part.to_owned(), msg: msg.into() }
}

// ---------------------------------------------------------------- package

/// The parts of a zip package, in their original order.
pub(crate) struct Package {
    pub parts: Vec<(String, Vec<u8>)>,
}

impl Package {
    pub fn read(bytes: &[u8]) -> Result<Self, Error> {
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes))?;
        if zip.len() > MAX_ENTRIES {
            return Err(Error::Limit(format!("{} zip entries", zip.len())));
        }
        let mut parts = Vec::new();
        let mut total = 0u64;
        for i in 0..zip.len() {
            let mut f = zip.by_index(i)?;
            if f.is_dir() {
                continue;
            }
            let name = f.name().trim_start_matches('/').replace('\\', "/");
            if f.size() > MAX_PART_SIZE {
                return Err(Error::Limit(format!("{name} is larger than 256 MB")));
            }
            let mut data = Vec::with_capacity(usize::try_from(f.size().min(16 * 1024 * 1024)).unwrap_or(0));
            (&mut f).take(MAX_PART_SIZE + 1).read_to_end(&mut data)?;
            if data.len() as u64 > MAX_PART_SIZE {
                return Err(Error::Limit(format!("{name} is larger than 256 MB")));
            }
            total += data.len() as u64;
            if total > MAX_TOTAL_SIZE {
                return Err(Error::Limit("package larger than 1 GB uncompressed".into()));
            }
            parts.push((name, data));
        }
        Ok(Self { parts })
    }

    /// Part by name (case-insensitive, leading `/` optional).
    pub fn get(&self, name: &str) -> Option<&[u8]> {
        let n = name.trim_start_matches('/');
        self.parts.iter().find(|(p, _)| p.eq_ignore_ascii_case(n)).map(|(_, d)| d.as_slice())
    }

    /// Part as text.
    pub fn text(&self, name: &str) -> Result<Option<String>, Error> {
        match self.get(name) {
            None => Ok(None),
            Some(b) => decode_text(b).map(Some).map_err(|m| xml_err(name, m)),
        }
    }
}

/// Decodes an XML part: UTF-8 (with or without BOM) or UTF-16 with a BOM.
pub(crate) fn decode_text(b: &[u8]) -> Result<String, String> {
    if let Some(rest) = b.strip_prefix(&[0xff, 0xfe]) {
        let units: Vec<u16> = rest.as_chunks::<2>().0.iter().map(|c| u16::from_le_bytes(*c)).collect();
        return Ok(String::from_utf16_lossy(&units));
    }
    if let Some(rest) = b.strip_prefix(&[0xfe, 0xff]) {
        let units: Vec<u16> = rest.as_chunks::<2>().0.iter().map(|c| u16::from_be_bytes(*c)).collect();
        return Ok(String::from_utf16_lossy(&units));
    }
    let b = b.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(b);
    match std::str::from_utf8(b) {
        Ok(s) => Ok(s.to_owned()),
        Err(_) => Ok(String::from_utf8_lossy(b).into_owned()),
    }
}

// ---------------------------------------------------------------- relationships

#[derive(Debug, Clone)]
pub(crate) struct Rel {
    pub id: String,
    pub kind: String,
    pub target: String,
    pub external: bool,
}

impl Rel {
    /// The relationship type's last path segment (`worksheet`, `styles`, `officeDocument`, ...).
    pub fn kind_tail(&self) -> &str {
        self.kind.rsplit('/').next().unwrap_or("")
    }
}

/// The relationships part of `part` (`xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`).
pub(crate) fn rels_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

pub(crate) fn parse_rels(text: &str, part: &str) -> Result<Vec<Rel>, Error> {
    let root = xml::parse_tree(text).map_err(|m| xml_err(part, m))?;
    Ok(root
        .children_named("Relationship")
        .map(|r| Rel {
            id: r.attr("Id").unwrap_or("").to_owned(),
            kind: r.attr("Type").unwrap_or("").to_owned(),
            target: r.attr("Target").unwrap_or("").to_owned(),
            external: r.attr("TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External")),
        })
        .collect())
}

/// Resolves a relationship target against the part that owns the relationship.
pub(crate) fn resolve_target(source_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    let mut segs: Vec<&str> = Vec::new();
    if let Some(abs) = target.strip_prefix('/') {
        for s in abs.split('/') {
            push_seg(&mut segs, s);
        }
    } else {
        if let Some((dir, _)) = source_part.rsplit_once('/') {
            for s in dir.split('/') {
                push_seg(&mut segs, s);
            }
        }
        for s in target.split('/') {
            push_seg(&mut segs, s);
        }
    }
    segs.join("/")
}

fn push_seg<'a>(segs: &mut Vec<&'a str>, s: &'a str) {
    match s {
        "" | "." => {}
        ".." => {
            segs.pop();
        }
        s => segs.push(s),
    }
}

/// A target for `part` relative to the directory of `source_part`.
pub(crate) fn relative_target(source_part: &str, part: &str) -> String {
    let dir = source_part.rsplit_once('/').map_or("", |(d, _)| d);
    if dir.is_empty() {
        return part.to_owned();
    }
    match part.strip_prefix(dir).and_then(|r| r.strip_prefix('/')) {
        Some(r) => r.to_owned(),
        None => format!("/{part}"),
    }
}

/// Checks a package's consistency and returns the problems found: XML parts that do not parse,
/// internal relationships whose target part is missing, parts without a content type, and
/// content-type overrides naming missing parts. An empty list is a consistent package.
pub fn validate(bytes: &[u8]) -> Result<Vec<String>, Error> {
    let pkg = Package::read(bytes)?;
    let mut problems = Vec::new();
    let names: std::collections::HashSet<String> = pkg.parts.iter().map(|(n, _)| n.to_ascii_lowercase()).collect();
    let mut defaults = Vec::new();
    let mut overrides = Vec::new();
    match pkg.text("[Content_Types].xml")? {
        None => problems.push("no [Content_Types].xml".to_owned()),
        Some(t) => match xml::parse_tree(&t) {
            Err(m) => problems.push(format!("[Content_Types].xml: {m}")),
            Ok(root) => {
                for d in root.children_named("Default") {
                    defaults.push(d.attr("Extension").unwrap_or("").to_ascii_lowercase());
                }
                for o in root.children_named("Override") {
                    let p = o.attr("PartName").unwrap_or("").trim_start_matches('/').to_ascii_lowercase();
                    if !names.contains(&p) {
                        problems.push(format!("content type override for missing part {p}"));
                    }
                    overrides.push(p);
                }
            }
        },
    }
    for (name, data) in &pkg.parts {
        let lower = name.to_ascii_lowercase();
        if lower == "[content_types].xml" {
            continue;
        }
        let ext = lower.rsplit_once('.').map_or("", |(_, e)| e);
        if !overrides.contains(&lower) && !defaults.iter().any(|d| d == ext) {
            problems.push(format!("{name} has no content type"));
        }
        if ext == "xml" || ext == "rels" || ext == "vml" {
            let text = match decode_text(data) {
                Ok(t) => t,
                Err(m) => {
                    problems.push(format!("{name}: {m}"));
                    continue;
                }
            };
            if let Err(m) = xml::parse_tree(&text) {
                // Legacy VML is often not well-formed XML; only report OOXML parts.
                if ext != "vml" {
                    problems.push(format!("{name}: {m}"));
                }
                continue;
            }
            if ext == "rels" {
                let source = match lower.strip_suffix(".rels").and_then(|s| s.rsplit_once("_rels/")) {
                    Some((dir, file)) => format!("{dir}{file}"),
                    None => continue,
                };
                for r in parse_rels(&text, name)? {
                    if r.external {
                        continue;
                    }
                    let target = resolve_target(&source, &r.target).to_ascii_lowercase();
                    if !names.contains(&target) {
                        problems.push(format!("{name}: relationship {} targets missing part {target}", r.id));
                    }
                }
            }
        }
    }
    Ok(problems)
}

/// Reads a file into memory, refusing files over [`MAX_TOTAL_SIZE`].
pub(crate) fn read_file(path: &Path) -> Result<Vec<u8>, Error> {
    let f = std::fs::File::open(path)?;
    let mut data = Vec::new();
    f.take(MAX_TOTAL_SIZE + 1).read_to_end(&mut data)?;
    if data.len() as u64 > MAX_TOTAL_SIZE {
        return Err(Error::Limit("file larger than 1 GB".into()));
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_resolve() {
        assert_eq!(resolve_target("xl/workbook.xml", "worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml");
        assert_eq!(resolve_target("xl/workbook.xml", "/xl/styles.xml"), "xl/styles.xml");
        assert_eq!(resolve_target("xl/worksheets/sheet1.xml", "../drawings/d1.xml"), "xl/drawings/d1.xml");
        assert_eq!(resolve_target("", "xl/workbook.xml"), "xl/workbook.xml");
        assert_eq!(relative_target("xl/workbook.xml", "xl/worksheets/sheet9.xml"), "worksheets/sheet9.xml");
        assert_eq!(rels_path("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
    }
}
