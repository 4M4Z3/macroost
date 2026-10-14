//! macroost's VBA-compatible object model.
//!
//! Macros reach a presentation only through the [`macroost_lang::Object`] and
//! [`macroost_lang::Host`] traits. This crate implements them (Application, Presentations,
//! Presentation, Slides, Slide, SlideRange, Shapes, Shape, ShapeRange, TextFrame, TextRange,
//! Font, ParagraphFormat, FillFormat, LineFormat, ColorFormat, Table, Rows, Columns, Cell,
//! PageSetup, SlideShowSettings, Master, CustomLayouts) over the [`Presentation`] storage
//! trait, which a presentation application implements. [`MemoryPresentation`] is a complete
//! in-memory implementation used by tests and command-line runs.
//!
//! Written from the publicly documented object model and black-box observation of the
//! reference host (`conformance/expected/slides*.json`); behavior not yet confirmed is marked `ORACLE:`.

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

pub mod constants;
mod host;
pub mod model;
mod objects;
mod shapes;
mod state;
mod text;

pub use host::SlidesHost;
pub use model::{
    CharFormat, CustomLayout, FillProps, FrameProps, LineProps, MemoryPresentation, PlaceholderSpec, Presentation,
    Shape, ShapeKind, Slide, SlideShow, TableData, TextBody,
};
pub use state::PRESENTATION_ERROR;
