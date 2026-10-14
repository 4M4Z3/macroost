//! [`DocHost`]: the [`Host`] a macro runs against. It resolves the application's global names
//! (`ActiveDocument`, `Selection`, `Documents`, `ThisDocument`, `Application`) and the `wd...`
//! constants, and records `MsgBox` prompts and `Debug.Print` lines for tests.

use std::cell::{Ref, RefCell};
use std::rc::Rc;

use macroost_lang::{Args, Host, Result, RuntimeError, Value};

use crate::constants::constant;
use crate::document::Document;
use crate::objects::{app_value, doc_value};
use crate::range::selection_value;
use crate::state::{Shared, State};

/// The application host over one or more documents.
pub struct DocHost {
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

impl DocHost {
    /// A host whose only open document, `doc`, is active and holds the macros (`ThisDocument`).
    pub fn new(doc: Box<dyn Document>) -> Self {
        Self {
            app: Rc::new(RefCell::new(State::new(doc))),
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

    /// Opens another document and makes it active.
    pub fn open(&self, doc: Box<dyn Document>) -> bool {
        self.app.try_borrow_mut().map(|mut st| st.open(doc)).is_ok()
    }

    /// Read access to the active document (for tests and the embedding app). `None` while the
    /// object model is in the middle of a call or when no document is open.
    pub fn active_document(&self) -> Option<Ref<'_, dyn Document>> {
        let st = self.app.try_borrow().ok()?;
        let id = st.active_id().ok()?;
        Ref::filter_map(st, |s| s.docs.iter().find(|d| d.id == id).map(|d| d.doc.as_ref())).ok()
    }

    /// Mutable access to the active document between macro calls.
    pub fn with_active_document_mut<T>(&self, f: impl FnOnce(&mut dyn Document) -> T) -> Option<T> {
        let mut st = self.app.try_borrow_mut().ok()?;
        let id = st.active_id().ok()?;
        let e = st.entry_mut(id).ok()?;
        Some(f(e.doc.as_mut()))
    }

    /// Names of the open documents, oldest first.
    pub fn document_names(&self) -> Vec<String> {
        self.app.try_borrow().map(|st| st.docs.iter().map(|d| d.doc.name()).collect()).unwrap_or_default()
    }

    /// Actions recorded without a storage effect (Save, ...).
    pub fn log(&self) -> Vec<String> {
        self.app.try_borrow().map(|s| s.log.clone()).unwrap_or_default()
    }
}

fn resolve(app: &Shared, name: &str, args: &Args) -> Option<Result<Value>> {
    let lname = name.to_ascii_lowercase();
    let r = match lname.as_str() {
        "application" => Ok(app_value(app)),
        "activedocument" => app
            .try_borrow()
            .map_err(|_| crate::state::busy())
            .and_then(|st| st.active_id())
            .and_then(|id| doc_value(app, id)),
        "thisdocument" => {
            let id = app.try_borrow().map(|st| st.this_document).unwrap_or(0);
            doc_value(app, id)
        }
        "selection" => app
            .try_borrow()
            .map_err(|_| crate::state::busy())
            .and_then(|st| st.active_id())
            .and_then(|id| selection_value(app, id)),
        "documents" => match app_value(app) {
            Value::Object(Some(o)) => o.call("Documents", args),
            _ => Err(RuntimeError::object_required()),
        },
        _ => {
            let v = constant(name)?;
            return Some(Ok(Value::Long(v)));
        }
    };
    Some(r)
}

impl Host for DocHost {
    fn global(&mut self, name: &str, args: &Args) -> Option<Result<Value>> {
        resolve(&self.app, name, args)
    }

    fn global_assign(&mut self, name: &str, args: &Args, value: Value) -> Option<Result<()>> {
        let lname = name.to_ascii_lowercase();
        match lname.as_str() {
            "selection" => {
                let target = resolve(&self.app, name, args)?;
                Some(target.and_then(|t| match t {
                    Value::Object(Some(o)) => o.assign("Text", &Args::none(), value),
                    _ => Err(RuntimeError::object_required()),
                }))
            }
            "activedocument" | "thisdocument" | "documents" | "application" => {
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
