//! macroost's VBA-compatible object model.
//!
//! Macros reach a spreadsheet only through the [`macroost_lang::Object`] and
//! [`macroost_lang::Host`] traits. This crate implements them (Application, Workbook,
//! Worksheets, Worksheet, Range, Interior, Font, Borders, WorksheetFunction, ...) over the
//! [`Workbook`] storage trait, which a spreadsheet application implements. [`MemoryWorkbook`]
//! is a complete in-memory implementation used by tests and command-line runs.
//!
//! Written from the publicly documented behavior of the object model; anything not yet confirmed
//! against a running application is marked `ORACLE:`.
//!
//! Ranges are descriptors (sheet, areas of (r1, c1, r2, c2), kind), never per-cell allocations,
//! so whole-column and whole-sheet ranges are cheap; operations over huge ranges visit only the
//! stored cells or are clipped to the used range.

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

pub mod address;
pub mod constants;
mod eval;
mod fill;
mod filter;
mod format;
mod formula;
mod functions;
mod host;
pub mod memory;
mod names;
mod notes;
mod objects;
pub mod protect;
mod range;
mod sorting;
mod state;
pub mod workbook;

pub use address::{MAX_COLS, MAX_ROWS, Rect};
pub use format::{format_value, is_date_format};
pub use formula::{RefEdit, a1_to_r1c1, adjust_refs, r1c1_to_a1, shift_formula};
pub use host::SheetHost;
pub use memory::MemoryWorkbook;
pub use protect::{legacy_password_hash, verify_password};
pub use workbook::{
    BorderLine, Borders, CellFormat, HAlign, Hyperlink, PasswordHash, SheetId, SheetVisibility, ShiftDirection,
    Workbook, WorkbookProtection,
};
