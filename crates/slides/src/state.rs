//! Shared application state behind every object, and small argument/value helpers.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::{Rc, Weak};

use macroost_lang::{Args, Result, RuntimeError, Value};

use crate::model::{Presentation, Shape, Slide};
use crate::objects::{Kind, PresentationObj};

pub(crate) type Shared = Rc<RefCell<State>>;

/// Error number the host raises for an invalid index, name or request (0x80048240).
pub const PRESENTATION_ERROR: i32 = -2_147_188_160;

pub(crate) fn app_err(msg: impl Into<String>) -> RuntimeError {
    RuntimeError::new(PRESENTATION_ERROR, msg)
}

pub(crate) fn busy() -> RuntimeError {
    RuntimeError::new(PRESENTATION_ERROR, "the object model is busy (re-entrant call)")
}

pub(crate) struct State {
    /// Open presentations with their stable keys, in collection order.
    pub pres: Vec<(u32, Box<dyn Presentation>)>,
    pub next_key: u32,
    pub active: Option<u32>,
    /// Counter for `Presentations.Add` names (Presentation1, Presentation2, ...).
    pub untitled: u32,
    /// Actions recorded without a storage effect (Save, Quit, ...).
    pub log: Vec<String>,
    /// Live objects by identity, so `ActivePresentation Is Presentations(1)` holds.
    pub cache: HashMap<Kind, Weak<PresentationObj>>,
}

impl State {
    pub fn new(first: Box<dyn Presentation>) -> Self {
        Self {
            pres: vec![(1, first)],
            next_key: 2,
            active: Some(1),
            untitled: 0,
            log: Vec::new(),
            cache: HashMap::new(),
        }
    }

    pub fn add(&mut self, p: Box<dyn Presentation>) -> u32 {
        let key = self.next_key;
        self.next_key += 1;
        self.pres.push((key, p));
        self.active = Some(key);
        key
    }

    pub fn pres(&self, pk: u32) -> Result<&dyn Presentation> {
        self.pres
            .iter()
            .find(|(k, _)| *k == pk)
            .map(|(_, p)| p.as_ref())
            .ok_or_else(|| app_err("the presentation is closed"))
    }

    pub fn pres_mut(&mut self, pk: u32) -> Result<&mut Box<dyn Presentation>> {
        self.pres
            .iter_mut()
            .find(|(k, _)| *k == pk)
            .map(|(_, p)| p)
            .ok_or_else(|| app_err("the presentation is closed"))
    }

    pub fn slide_index(&self, pk: u32, sid: u32) -> Result<usize> {
        self.pres(pk)?.slides().iter().position(|s| s.id == sid).ok_or_else(|| app_err("the slide was deleted"))
    }

    pub fn slide(&self, pk: u32, sid: u32) -> Result<&Slide> {
        let i = self.slide_index(pk, sid)?;
        self.pres(pk)?.slides().get(i).ok_or_else(|| app_err("the slide was deleted"))
    }

    pub fn slide_mut(&mut self, pk: u32, sid: u32) -> Result<&mut Slide> {
        let i = self.slide_index(pk, sid)?;
        self.pres_mut(pk)?.slides_mut().get_mut(i).ok_or_else(|| app_err("the slide was deleted"))
    }

    pub fn shape(&self, r: &ShapeRef) -> Result<&Shape> {
        self.slide(r.pk, r.sid)?.shape(r.id).ok_or_else(|| app_err("the shape was deleted"))
    }

    pub fn shape_mut(&mut self, r: &ShapeRef) -> Result<&mut Shape> {
        self.slide_mut(r.pk, r.sid)?.shape_mut(r.id).ok_or_else(|| app_err("the shape was deleted"))
    }
}

/// A shape: presentation key, SlideID, shape id.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct ShapeRef {
    pub pk: u32,
    pub sid: u32,
    pub id: u32,
}

pub(crate) fn with<T>(app: &Shared, f: impl FnOnce(&State) -> Result<T>) -> Result<T> {
    let st = app.try_borrow().map_err(|_| busy())?;
    f(&st)
}

pub(crate) fn with_mut<T>(app: &Shared, f: impl FnOnce(&mut State) -> Result<T>) -> Result<T> {
    let mut st = app.try_borrow_mut().map_err(|_| busy())?;
    f(&mut st)
}

// ---- values ----

pub(crate) fn single(x: f32) -> Value {
    Value::Single(x)
}

pub(crate) fn long(n: impl TryInto<i32>) -> Value {
    Value::Long(n.try_into().unwrap_or(i32::MAX))
}

/// `msoTrue` / `msoFalse`.
pub(crate) fn tri(b: bool) -> Value {
    Value::Long(if b { -1 } else { 0 })
}

pub(crate) fn text(s: impl Into<Rc<str>>) -> Value {
    Value::String(s.into())
}

/// A tri-state or Boolean argument: `True`, `msoTrue`, `msoCTrue` are true.
pub(crate) fn to_tri(v: &Value) -> Result<bool> {
    let v = v.clone().deref_default()?;
    match v {
        Value::Boolean(b) => Ok(b),
        other => Ok(other.to_f64()? != 0.0),
    }
}

pub(crate) fn data(v: &Value) -> Result<Value> {
    v.clone().deref_default()
}

pub(crate) fn arg_f64(args: &Args, i: usize, name: &str) -> Result<f64> {
    data(args.require(i, name)?)?.to_f64()
}

pub(crate) fn arg_i64(args: &Args, i: usize, name: &str) -> Result<i64> {
    data(args.require(i, name)?)?.to_i64()
}

pub(crate) fn opt_i64(args: &Args, i: usize, name: &str) -> Result<Option<i64>> {
    args.get(i, name).map(|v| data(v).and_then(|v| v.to_i64())).transpose()
}

pub(crate) fn opt_text(args: &Args, i: usize, name: &str) -> Result<Option<String>> {
    args.get(i, name).map(|v| data(v).and_then(|v| v.to_text())).transpose()
}

pub(crate) fn opt_bool(args: &Args, i: usize, name: &str) -> Result<Option<bool>> {
    args.get(i, name).map(to_tri).transpose()
}

/// The single value assigned to a property.
pub(crate) fn assigned_f64(v: &Value) -> Result<f64> {
    data(v)?.to_f64()
}

/// A 1-based index into `len` items, or the error the host raises.
pub(crate) fn index_in(n: i64, len: usize) -> Result<usize> {
    if n >= 1 && (n as u64) <= len as u64 {
        Ok((n - 1) as usize)
    } else {
        Err(app_err(format!("index {n} is out of range (1 to {len})")))
    }
}

/// Elements of an `Array(...)` argument, or the value itself.
pub(crate) fn list_arg(v: &Value) -> Vec<Value> {
    match v {
        Value::Array(a) => a.borrow().data.clone(),
        other => vec![other.clone()],
    }
}
