//! [`SlidesHost`]: the [`Host`] a presentation macro runs against. It resolves the global
//! names (`Application`, `ActivePresentation`, `Presentations`) and the `pp*` / `mso*`
//! constants, and records `MsgBox` prompts and `Debug.Print` lines for tests.

use std::cell::{Ref, RefCell};
use std::rc::Rc;

use macroost_lang::{Args, Host, Result, RuntimeError, Value};

use crate::constants::constant;
use crate::model::Presentation;
use crate::objects::{Kind, active_presentation, call_value, obj};
use crate::state::{Shared, State};

pub struct SlidesHost {
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

impl SlidesHost {
    /// A host whose only open presentation (and `ActivePresentation`) is `presentation`.
    pub fn new(presentation: Box<dyn Presentation>) -> Self {
        Self {
            app: Rc::new(RefCell::new(State::new(presentation))),
            messages: Vec::new(),
            printed: Vec::new(),
            input_replies: Vec::new(),
            msgbox_reply: 1,
            fixed_now: None,
        }
    }

    /// The `Application` object.
    pub fn application(&self) -> Value {
        obj(&self.app, Kind::Application)
    }

    /// Number of open presentations.
    pub fn presentation_count(&self) -> usize {
        self.app.try_borrow().map(|s| s.pres.len()).unwrap_or(0)
    }

    /// Read access to open presentation `index` (0-based) between macro calls. `None` while the
    /// object model is in the middle of a call or when there is no such presentation.
    pub fn presentation(&self, index: usize) -> Option<Ref<'_, dyn Presentation>> {
        let st = self.app.try_borrow().ok()?;
        if index >= st.pres.len() {
            return None;
        }
        Some(Ref::map(st, |s| s.pres[index].1.as_ref()))
    }

    /// Mutable access to open presentation `index` (0-based) between macro calls.
    pub fn with_presentation_mut<T>(&self, index: usize, f: impl FnOnce(&mut dyn Presentation) -> T) -> Option<T> {
        let mut st = self.app.try_borrow_mut().ok()?;
        let p = st.pres.get_mut(index)?;
        Some(f(p.1.as_mut()))
    }

    /// Actions recorded without a storage effect (Save, Quit).
    pub fn log(&self) -> Vec<String> {
        self.app.try_borrow().map(|s| s.log.clone()).unwrap_or_default()
    }
}

fn resolve(app: &Shared, name: &str, args: &Args) -> Option<Result<Value>> {
    let lname = name.to_ascii_lowercase();
    let r = match lname.as_str() {
        "application" => Ok(obj(app, Kind::Application)),
        "activepresentation" => active_presentation(app),
        "presentations" => {
            let c = obj(app, Kind::Presentations);
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        _ => return constant(name).map(|v| Ok(Value::Long(v))),
    };
    Some(r)
}

impl Host for SlidesHost {
    fn global(&mut self, name: &str, args: &Args) -> Option<Result<Value>> {
        resolve(&self.app, name, args)
    }

    fn global_assign(&mut self, name: &str, _args: &Args, _value: Value) -> Option<Result<()>> {
        let lname = name.to_ascii_lowercase();
        let ours = matches!(lname.as_str(), "application" | "activepresentation" | "presentations")
            || constant(name).is_some();
        ours.then(|| Err(RuntimeError::new(450, "Wrong number of arguments or invalid property assignment")))
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
