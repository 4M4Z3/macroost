//! The boundary between the language and an application's object model. The interpreter knows
//! nothing about the spreadsheet host: it calls members by name on [`Object`]s and asks the [`Host`] for
//! global names (`ActiveSheet`, `Range`, `xlUp`, ...). An object model crate implements both.

use std::rc::Rc;

use crate::error::{Result, RuntimeError};
use crate::value::Value;

pub type ObjectRef = Rc<dyn Object>;

/// Arguments of a member call: positional, then `Name:=value` pairs.
#[derive(Debug, Clone, Default)]
pub struct Args {
    pub positional: Vec<Value>,
    pub named: Vec<(String, Value)>,
}

impl Args {
    pub fn none() -> Self {
        Self::default()
    }
    pub fn positional(values: Vec<Value>) -> Self {
        Self { positional: values, named: Vec::new() }
    }
    pub fn is_empty(&self) -> bool {
        self.positional.is_empty() && self.named.is_empty()
    }
    pub fn len(&self) -> usize {
        self.positional.len() + self.named.len()
    }
    /// The argument at `index`, or the one named `name` (case-insensitive). A missing
    /// optional argument is `None`.
    pub fn get(&self, index: usize, name: &str) -> Option<&Value> {
        self.named
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v)| v)
            .or_else(|| self.positional.get(index))
            .filter(|v| !matches!(v, Value::Error(448))) // a skipped argument `f(1, , 3)`
    }
    pub fn require(&self, index: usize, name: &str) -> Result<&Value> {
        self.get(index, name).ok_or_else(|| RuntimeError::standard(449))
    }
}

/// The value standing for a skipped optional argument (`f(1, , 3)`); `IsMissing` is true.
pub const MISSING: Value = Value::Error(448);

pub trait Object {
    /// `TypeName` of the object (`Range`, `Worksheet`, ...).
    fn type_name(&self) -> String;

    /// Property get or method call. `name` is matched case-insensitively; an empty name is the
    /// default member (so `Cells(1, 1)` on a sheet is `call("Cells", [1, 1])` and `rng(2)` is
    /// `call("", [2])`).
    fn call(&self, name: &str, args: &Args) -> Result<Value>;

    /// Property assignment (`obj.Name(args) = value`).
    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        let _ = (args, value);
        Err(RuntimeError::no_member(name))
    }

    /// The default member's value, used when an object appears where data is expected.
    fn default_value(&self) -> Result<Value> {
        self.call("", &Args::none())
    }

    /// The items `For Each` walks.
    fn items(&self) -> Result<Vec<Value>> {
        Err(RuntimeError::no_member("_NewEnum"))
    }

    /// Lets the interpreter recognise its own objects (class instances, collections).
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        None
    }
}

/// What the running application provides to a macro.
pub trait Host {
    /// Resolves a name that is not a variable, constant or procedure of the project:
    /// application globals (`Application`, `ActiveSheet`, `Range`, `Cells`, `Selection`, ...)
    /// and enumeration constants (`xlUp`, `vbYes`, ...). `None` means "not mine".
    fn global(&mut self, name: &str, args: &Args) -> Option<Result<Value>>;

    /// Assignment to an application global (`ActiveCell = 1`, `Range("A1") = 2`).
    fn global_assign(&mut self, name: &str, args: &Args, value: Value) -> Option<Result<()>> {
        let _ = (name, args, value);
        None
    }

    /// `MsgBox`: returns the button pressed (vbOK = 1).
    fn msgbox(&mut self, prompt: &str, buttons: i32, title: &str) -> i32 {
        let _ = (prompt, buttons, title);
        1
    }

    /// `InputBox`: `None` when cancelled.
    fn inputbox(&mut self, prompt: &str, title: &str, default: &str) -> Option<String> {
        let _ = (prompt, title);
        Some(default.to_owned())
    }

    /// `Debug.Print`.
    fn debug_print(&mut self, text: &str) {
        let _ = text;
    }

    /// Events the application raised since the last call (a cell changed, the selection
    /// moved...). The interpreter runs each one's handler, if the project has it, between
    /// statements.
    fn take_events(&mut self) -> Vec<HostEvent> {
        Vec::new()
    }

    /// Current time as a serial date (`Now`). Hosts override for determinism in tests.
    fn now(&mut self) -> f64 {
        let secs =
            std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs_f64()).unwrap_or(0.0);
        crate::value::UNIX_EPOCH_SERIAL + secs / 86_400.0
    }
}

/// An event for a document module's handler, e.g. module `Sheet1`, procedure
/// `Worksheet_Change`, args `[Target]`.
#[derive(Debug, Clone)]
pub struct HostEvent {
    pub module: String,
    pub procedure: String,
    pub args: Vec<Value>,
}

/// A host with no application: plain language only.
#[derive(Debug, Default)]
pub struct NoHost {
    pub printed: Vec<String>,
    pub messages: Vec<String>,
}

impl Host for NoHost {
    fn global(&mut self, _name: &str, _args: &Args) -> Option<Result<Value>> {
        None
    }
    fn msgbox(&mut self, prompt: &str, _buttons: i32, _title: &str) -> i32 {
        self.messages.push(prompt.to_owned());
        1
    }
    fn debug_print(&mut self, text: &str) {
        self.printed.push(text.to_owned());
    }
}

/// Identity comparison for `Is`.
pub fn same_object(a: &Option<ObjectRef>, b: &Option<ObjectRef>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(x), Some(y)) => Rc::ptr_eq(x, y),
        _ => false,
    }
}
