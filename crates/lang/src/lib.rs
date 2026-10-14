//! macroost's language engine: a clean-room interpreter for VBA-compatible macros, written
//! from the published language specification ([MS-VBAL]) and black-box testing.

#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

pub mod ast;
pub mod builtins;
pub mod error;
pub mod interp;
pub mod lexer;
pub mod object;
pub mod ops;
pub mod parser;
pub mod stdlib;
pub mod value;

pub use error::{Result, RuntimeError};
pub use interp::Interpreter;
pub use object::{Args, Host, HostEvent, NoHost, Object, ObjectRef};
pub use value::{Array, Value};
