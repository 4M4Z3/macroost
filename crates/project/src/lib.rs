//! Read macro projects out of macro-enabled documents, following the published [MS-OVBA]
//! specification. Handles OOXML packages (.xlsm, .xlsb, .docm, .pptm, ...) and legacy compound
//! files (.xls, .doc, .ppt) as well as a bare `vbaProject.bin`. The [`write`] module produces
//! source-only projects and minimal macro-enabled workbooks.

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
#![cfg_attr(test, allow(clippy::unwrap_used, clippy::expect_used, clippy::panic))]

pub mod compression;
pub mod dir;
pub mod write;

pub use write::{
    Reference, docm_with_macros, pptm_with_macros, project_bin, project_bin_with_documents,
    project_bin_with_references, xlsm_with_macros,
};

use std::io::{Cursor, Read};

#[derive(Debug)]
pub enum Error {
    Io(std::io::Error),
    Corrupt(&'static str),
    NoMacros,
    /// The writer was given input it cannot represent.
    Invalid(String),
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Io(e) => write!(f, "{e}"),
            Error::Corrupt(why) => write!(f, "corrupt macro project: {why}"),
            Error::NoMacros => write!(f, "the file contains no macro project"),
            Error::Invalid(why) => write!(f, "cannot write macro project: {why}"),
        }
    }
}

impl std::error::Error for Error {}

impl From<std::io::Error> for Error {
    fn from(e: std::io::Error) -> Self {
        Error::Io(e)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModuleKind {
    /// A standard module (`Module=` in the PROJECT stream).
    Standard,
    /// A class module.
    Class,
    /// A document module such as ThisWorkbook or Sheet1.
    Document,
    /// A UserForm (designer) module.
    Form,
}

#[derive(Debug, Clone)]
pub struct Module {
    pub name: String,
    pub kind: ModuleKind,
    /// Source text, with the `Attribute VB_...` header lines the editor hides.
    pub source: String,
}

#[derive(Debug, Clone)]
pub struct Project {
    pub name: String,
    pub code_page: u16,
    pub modules: Vec<Module>,
}

/// Largest stream we are willing to read, to bound memory on hostile files.
const MAX_STREAM: u64 = 64 * 1024 * 1024;

/// Reads the macro project from the bytes of any supported macro-enabled document.
pub fn read(bytes: &[u8]) -> Result<Project, Error> {
    if bytes.starts_with(b"PK") {
        let bin = project_bin_from_zip(bytes)?;
        read_cfb(&bin)
    } else if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        read_cfb(bytes)
    } else {
        Err(Error::NoMacros)
    }
}

pub fn read_file(path: &std::path::Path) -> Result<Project, Error> {
    read(&std::fs::read(path)?)
}

fn project_bin_from_zip(bytes: &[u8]) -> Result<Vec<u8>, Error> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|_| Error::Corrupt("not a readable zip package"))?;
    let name = zip
        .file_names()
        .find(|n| n.to_ascii_lowercase().ends_with("vbaproject.bin"))
        .map(str::to_owned)
        .ok_or(Error::NoMacros)?;
    let entry = zip.by_name(&name).map_err(|_| Error::Corrupt("vbaProject.bin unreadable"))?;
    let mut out = Vec::new();
    entry.take(MAX_STREAM).read_to_end(&mut out)?;
    Ok(out)
}

fn read_stream<F: Read + std::io::Seek>(cfb: &mut cfb::CompoundFile<F>, path: &str) -> Result<Vec<u8>, Error> {
    let s = cfb.open_stream(path).map_err(|_| Error::Corrupt("missing stream"))?;
    let mut out = Vec::new();
    s.take(MAX_STREAM).read_to_end(&mut out)?;
    Ok(out)
}

/// Finds the storage holding the `dir` stream: `/VBA` in vbaProject.bin, `/_VBA_PROJECT_CUR/VBA`
/// in .xls, `/Macros/VBA` in .doc.
fn find_project_storage<F: Read + std::io::Seek>(cfb: &cfb::CompoundFile<F>) -> Option<String> {
    cfb.walk()
        .filter(|e| e.is_stream() && e.name().eq_ignore_ascii_case("dir"))
        .filter_map(|e| e.path().parent().map(|p| p.to_string_lossy().replace('\\', "/")))
        .find(|p| p.to_ascii_lowercase().ends_with("vba"))
}

fn read_cfb(bytes: &[u8]) -> Result<Project, Error> {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(bytes)).map_err(|_| Error::Corrupt("not a compound file"))?;
    let vba = find_project_storage(&cfb).ok_or(Error::NoMacros)?;
    let project_root = vba.rsplit_once('/').map(|(root, _)| root.to_owned()).unwrap_or_default();

    let dir = compression::decompress(&read_stream(&mut cfb, &format!("{vba}/dir"))?)?;
    let info = dir::parse(&dir)?;

    // The PROJECT stream (plain text) says which modules are classes, documents or forms.
    let kinds = read_stream(&mut cfb, &format!("{project_root}/PROJECT"))
        .map(|b| project_kinds(&dir::decode(&b, info.code_page)))
        .unwrap_or_default();

    let mut modules = Vec::new();
    for m in &info.modules {
        let Ok(stream) = read_stream(&mut cfb, &format!("{vba}/{}", m.stream_name)) else {
            continue;
        };
        let source = stream
            .get(m.text_offset as usize..)
            .map(compression::decompress)
            .transpose()?
            .map(|b| dir::decode(&b, info.code_page))
            .unwrap_or_default();
        let kind = kinds
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(&m.name))
            .map(|(_, k)| *k)
            .unwrap_or(if m.non_procedural { ModuleKind::Class } else { ModuleKind::Standard });
        modules.push(Module { name: m.name.clone(), kind, source });
    }
    Ok(Project { name: info.project_name, code_page: info.code_page, modules })
}

/// `Module=Name`, `Class=Name`, `Document=Name/&H00000000`, `BaseClass=Name` lines (2.3.1).
fn project_kinds(text: &str) -> Vec<(String, ModuleKind)> {
    text.lines()
        .filter_map(|line| {
            let (key, value) = line.split_once('=')?;
            let kind = match key.trim() {
                "Module" => ModuleKind::Standard,
                "Class" => ModuleKind::Class,
                "Document" => ModuleKind::Document,
                "BaseClass" => ModuleKind::Form,
                _ => return None,
            };
            let name = value.split('/').next()?.trim().to_owned();
            Some((name, kind))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn project_stream_kinds() {
        let k = project_kinds(
            "ID=\"{1}\"\r\nDocument=ThisWorkbook/&H00000000\r\nModule=Module1\r\nClass=Thing\r\nBaseClass=UserForm1\r\n",
        );
        assert_eq!(k.len(), 4);
        assert_eq!(k[0], ("ThisWorkbook".to_owned(), ModuleKind::Document));
        assert_eq!(k[3].1, ModuleKind::Form);
    }

    #[test]
    fn non_document_bytes_have_no_macros() {
        assert!(matches!(read(b"hello"), Err(Error::NoMacros)));
    }
}
