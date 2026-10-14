//! macroost's word-processor object model (VBA-compatible).
//!
//! Macros reach a document only through the [`macroost_lang::Object`] and
//! [`macroost_lang::Host`] traits. This crate implements them (Application, Documents,
//! Document, Range, Selection, Paragraphs, Words, Characters, Sentences, Font, ParagraphFormat,
//! Find/Replacement, Tables/Rows/Columns/Cells, Bookmarks, Styles) over the [`Document`] storage
//! trait, which a word-processing application implements. [`MemoryDocument`] is a complete
//! in-memory implementation used by tests and command-line runs.
//!
//! Written from the publicly documented object model and black-box observation of the
//! reference application (conformance/cases/doc, conformance/expected/doc.json).

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
#![cfg_attr(test, allow(clippy::unwrap_used, clippy::expect_used, clippy::panic))]

pub mod constants;
pub mod document;
mod find;
mod host;
mod objects;
mod range;
mod state;
mod table;

pub use document::{Bookmark, CharFormat, Document, Item, Marker, MemoryDocument, ParaFormat};
pub use host::DocHost;
