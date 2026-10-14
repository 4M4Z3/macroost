//! [`SheetHost`]: the [`Host`] a macro runs against. It resolves the application's global names
//! and enumeration constants, routes `Range("A1") = x` style assignments, and records `MsgBox`
//! prompts and `Debug.Print` lines for tests.

use std::cell::{Ref, RefCell};
use std::rc::Rc;

use macroost_lang::{Args, Host, Result, RuntimeError, Value};

use crate::constants::constant;
use crate::functions::WorksheetFunctionObj;
use crate::objects::{
    active_cell, active_sheet, app_value, evaluate, global_range, intersect, selection, sheet_member, sheet_value,
    sheets_value, union, workbook_value,
};
use crate::state::{Shared, State, with, with_mut};
use crate::workbook::Workbook;

/// The application host over a workbook.
pub struct SheetHost {
    app: Shared,
    /// `MsgBox` prompts, in order.
    pub messages: Vec<String>,
    /// `Debug.Print` lines, in order.
    pub printed: Vec<String>,
    /// Answers to `InputBox`, consumed in order; when empty the default is returned.
    pub input_replies: Vec<Option<String>>,
    /// Button `MsgBox` reports as pressed (vbOK = 1).
    pub msgbox_reply: i32,
    /// Fixed `Now` (serial date) for deterministic tests; `None` uses the clock.
    pub fixed_now: Option<f64>,
}

impl SheetHost {
    pub fn new(book: Box<dyn Workbook>) -> Self {
        Self {
            app: Rc::new(RefCell::new(State::new(book))),
            messages: Vec::new(),
            printed: Vec::new(),
            input_replies: Vec::new(),
            msgbox_reply: 1,
            fixed_now: None,
        }
    }

    /// The `Application` object.
    pub fn application(&self) -> Value {
        app_value(&self.app)
    }

    /// Read access to the underlying workbook (for tests and the embedding app).
    /// Returns `None` while the object model is in the middle of a call.
    pub fn workbook(&self) -> Option<Ref<'_, dyn Workbook>> {
        let st = self.app.try_borrow().ok()?;
        Some(Ref::map(st, |s| s.book.as_ref()))
    }

    /// Mutable access to the workbook between macro calls.
    pub fn with_workbook_mut<T>(&self, f: impl FnOnce(&mut dyn Workbook) -> T) -> Option<T> {
        let mut st = self.app.try_borrow_mut().ok()?;
        Some(f(st.book.as_mut()))
    }

    /// Actions recorded without a storage effect (Save, AutoFilter, Merge, ...).
    pub fn log(&self) -> Vec<String> {
        self.app.try_borrow().map(|s| s.log.clone()).unwrap_or_default()
    }

    /// Takes the workbook back out of the host.
    pub fn into_workbook(self) -> std::result::Result<Box<dyn Workbook>, Self> {
        match Rc::try_unwrap(self.app) {
            Ok(cell) => Ok(cell.into_inner().book),
            Err(app) => Err(Self { app, ..self }),
        }
    }
}

fn resolve(app: &Shared, name: &str, args: &Args) -> Option<Result<Value>> {
    let lname = name.to_ascii_lowercase();
    let r = match lname.as_str() {
        "application" => Ok(app_value(app)),
        "activesheet" => active_sheet(app).map(|id| sheet_value(app, id)),
        "activeworkbook" | "thisworkbook" => Ok(workbook_value(app)),
        "names" => crate::names::names_member(app, None, args),
        "activewindow" => Ok(crate::objects::window_value(app)),
        "activecell" | "selection" => {
            let base = if lname == "activecell" { active_cell(app) } else { selection(app) };
            // `Selection(2)`: the default member (Item) of the range.
            if args.is_empty() { base } else { base.and_then(|b| call_on(&b, "", args)) }
        }
        "range" => global_range(app, args),
        "cells" | "rows" | "columns" => active_sheet(app).and_then(|id| sheet_member(app, id, name, args)),
        "worksheets" | "sheets" => {
            let coll = sheets_value(app, if lname == "sheets" { "Sheets" } else { "Worksheets" });
            if args.is_empty() { Ok(coll) } else { call_on(&coll, "Item", args) }
        }
        "workbooks" => call_on(&app_value(app), "Workbooks", args),
        "worksheetfunction" => Ok(Value::Object(Some(Rc::new(WorksheetFunctionObj { app: app.clone() })))),
        "intersect" => intersect(app, args),
        "union" => union(app, args),
        "evaluate" => evaluate(app, args),
        "calculate" => call_on(&app_value(app), "Calculate", args),
        _ => {
            if let Some(v) = constant(name) {
                return Some(Ok(Value::Long(v)));
            }
            // A sheet's code name (`Sheet12.Range("A1")`).
            let id = with(app, |st| {
                Ok(st
                    .book
                    .sheet_ids()
                    .into_iter()
                    .find(|id| st.book.code_name(*id).is_some_and(|c| c.eq_ignore_ascii_case(name))))
            })
            .ok()
            .flatten()?;
            Ok(sheet_value(app, id))
        }
    };
    Some(r)
}

fn call_on(v: &Value, name: &str, args: &Args) -> Result<Value> {
    match v {
        Value::Object(Some(o)) => o.call(name, args),
        _ => Err(RuntimeError::object_required()),
    }
}

impl Host for SheetHost {
    fn global(&mut self, name: &str, args: &Args) -> Option<Result<Value>> {
        resolve(&self.app, name, args)
    }

    fn global_assign(&mut self, name: &str, args: &Args, value: Value) -> Option<Result<()>> {
        let lname = name.to_ascii_lowercase();
        match lname.as_str() {
            "range" | "cells" | "activecell" | "selection" | "rows" | "columns" | "evaluate" => {
                let target = resolve(&self.app, name, args)?;
                Some(target.and_then(|t| match t {
                    Value::Object(Some(o)) => o.assign("Value", &Args::none(), value),
                    _ => Err(RuntimeError::object_required()),
                }))
            }
            "worksheets" | "sheets" | "activesheet" | "application" | "activeworkbook" | "thisworkbook" => {
                Some(Err(RuntimeError::new(450, "Wrong number of arguments or invalid property assignment")))
            }
            _ if constant(name).is_some() => {
                Some(Err(RuntimeError::new(450, "Wrong number of arguments or invalid property assignment")))
            }
            _ => None,
        }
    }

    fn msgbox(&mut self, prompt: &str, _buttons: i32, _title: &str) -> i32 {
        self.messages.push(prompt.to_owned());
        self.msgbox_reply
    }

    fn inputbox(&mut self, prompt: &str, _title: &str, default: &str) -> Option<String> {
        self.messages.push(prompt.to_owned());
        if self.input_replies.is_empty() { Some(default.to_owned()) } else { self.input_replies.remove(0) }
    }

    fn debug_print(&mut self, text: &str) {
        self.printed.push(text.to_owned());
    }

    fn now(&mut self) -> f64 {
        match self.fixed_now {
            Some(n) => n,
            None => {
                let secs = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_secs_f64())
                    .unwrap_or(0.0);
                macroost_lang::value::UNIX_EPOCH_SERIAL + secs / 86_400.0
            }
        }
    }
}

/// Control from the embedding application, between macro calls.
impl SheetHost {
    pub fn set_active_sheet(&self, name: &str) -> Result<()> {
        with_mut(&self.app, |st| {
            let id = st.sheet_by_name(name).ok_or_else(RuntimeError::subscript)?;
            st.active_sheet = id;
            Ok(())
        })
    }
}
