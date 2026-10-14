//! Objects the language itself provides: `Err`, `Collection`, and the dictionary that macros
//! create with `CreateObject("Scripting.Dictionary")` (the most common late-bound object).

use std::cell::RefCell;
use std::rc::Rc;

use crate::error::{Result, RuntimeError};
use crate::object::{Args, Object};
use crate::value::{Array, Value};

// ---------------------------------------------------------------- Err

#[derive(Debug, Default)]
pub struct ErrState {
    pub number: i32,
    pub description: String,
    pub source: String,
}

#[derive(Debug, Default)]
pub struct ErrObject {
    pub state: RefCell<ErrState>,
}

impl ErrObject {
    pub fn set(&self, e: &RuntimeError) {
        let mut s = self.state.borrow_mut();
        s.number = e.number;
        s.description = e.description.clone();
        s.source = "VBAProject".into();
    }
    pub fn clear(&self) {
        *self.state.borrow_mut() = ErrState::default();
    }
    pub fn number(&self) -> i32 {
        self.state.borrow().number
    }
}

impl Object for ErrObject {
    fn type_name(&self) -> String {
        "ErrObject".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "" | "number" => Ok(Value::Long(self.state.borrow().number)),
            "description" => Ok(Value::str(self.state.borrow().description.clone())),
            "source" => Ok(Value::str(self.state.borrow().source.clone())),
            "helpfile" => Ok(Value::str("")),
            "helpcontext" | "lastdllerror" => Ok(Value::Long(0)),
            "clear" => {
                self.clear();
                Ok(Value::Empty)
            }
            "raise" => {
                let number = args.require(0, "Number")?.to_i32()?;
                let source = args.get(1, "Source").map(Value::to_text).transpose()?;
                let description = match args.get(2, "Description") {
                    Some(d) => d.to_text()?,
                    None => crate::error::standard_description(number).to_owned(),
                };
                let mut e = RuntimeError::new(number, description);
                if let Some(s) = source {
                    e.description = if e.description.is_empty() { s } else { e.description };
                }
                Err(e)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let mut s = self.state.borrow_mut();
        match name.to_ascii_lowercase().as_str() {
            "" | "number" => s.number = value.to_i32()?,
            "description" => s.description = value.to_text()?,
            "source" => s.source = value.to_text()?,
            _ => return Err(RuntimeError::no_member(name)),
        }
        Ok(())
    }
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

// ---------------------------------------------------------------- Collection

/// `Collection`: 1-based items with optional case-insensitive string keys.
#[derive(Debug, Default)]
pub struct Collection {
    items: RefCell<Vec<(Option<String>, Value)>>,
}

impl Collection {
    pub fn new_ref() -> Value {
        Value::Object(Some(Rc::new(Collection::default())))
    }

    fn index_of(&self, key: &Value) -> Result<usize> {
        let items = self.items.borrow();
        match key {
            Value::String(k) => items
                .iter()
                .position(|(ik, _)| ik.as_deref().is_some_and(|ik| ik.eq_ignore_ascii_case(k)))
                .ok_or_else(RuntimeError::invalid_call),
            other => {
                let i = other.to_i64()?;
                if i >= 1 && (i as usize) <= items.len() {
                    Ok(i as usize - 1)
                } else {
                    Err(RuntimeError::invalid_call())
                }
            }
        }
    }
}

impl Object for Collection {
    fn type_name(&self) -> String {
        "Collection".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.items.borrow().len() as i32)),
            "" | "item" => {
                let i = self.index_of(args.require(0, "Index")?)?;
                Ok(self.items.borrow()[i].1.clone())
            }
            "add" => {
                let item = args.require(0, "Item")?.clone();
                let key = args.get(1, "Key").map(Value::to_text).transpose()?;
                if let Some(k) = &key {
                    if self
                        .items
                        .borrow()
                        .iter()
                        .any(|(ik, _)| ik.as_deref().is_some_and(|ik| ik.eq_ignore_ascii_case(k)))
                    {
                        return Err(RuntimeError::new(
                            457,
                            "This key is already associated with an element of this collection",
                        ));
                    }
                }
                let pos = if let Some(b) = args.get(2, "Before") {
                    Some(self.index_of(b)?)
                } else if let Some(a) = args.get(3, "After") {
                    Some(self.index_of(a)? + 1)
                } else {
                    None
                };
                let mut items = self.items.borrow_mut();
                match pos {
                    Some(p) => {
                        let at = p.min(items.len());
                        items.insert(at, (key, item));
                    }
                    None => items.push((key, item)),
                }
                Ok(Value::Empty)
            }
            "remove" => {
                let i = self.index_of(args.require(0, "Index")?)?;
                self.items.borrow_mut().remove(i);
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn items(&self) -> Result<Vec<Value>> {
        Ok(self.items.borrow().iter().map(|(_, v)| v.clone()).collect())
    }
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

// ---------------------------------------------------------------- Dictionary

/// Key/value map with insertion order. Keys compare as their text under `CompareMode` 1
/// (text, case-insensitive) or exactly under 0 (binary, the default).
#[derive(Debug, Default)]
pub struct Dictionary {
    entries: RefCell<Vec<(Value, Value)>>,
    text_mode: RefCell<bool>,
}

impl Dictionary {
    pub fn new_ref() -> Value {
        Value::Object(Some(Rc::new(Dictionary::default())))
    }

    fn key_text(&self, k: &Value) -> Result<String> {
        let t = match k {
            Value::Object(_) => return Err(RuntimeError::type_mismatch()),
            v => format!("{}:{}", v.var_type().min(8), v.to_text()?),
        };
        Ok(if *self.text_mode.borrow() { t.to_lowercase() } else { t })
    }

    fn find(&self, k: &Value) -> Result<Option<usize>> {
        let kt = self.key_text(k)?;
        for (i, (ek, _)) in self.entries.borrow().iter().enumerate() {
            if self.key_text(ek)? == kt {
                return Ok(Some(i));
            }
        }
        Ok(None)
    }
}

/// Keys of numeric variants compare by value regardless of exact type.
fn normalize_key(v: &Value) -> Value {
    match v {
        Value::Byte(_)
        | Value::Integer(_)
        | Value::Long(_)
        | Value::LongLong(_)
        | Value::Single(_)
        | Value::Double(_)
        | Value::Currency(_) => v.to_f64().map(Value::Double).unwrap_or_else(|_| v.clone()),
        _ => v.clone(),
    }
}

impl Object for Dictionary {
    fn type_name(&self) -> String {
        "Dictionary".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.entries.borrow().len() as i32)),
            "" | "item" => {
                let k = normalize_key(args.require(0, "Key")?);
                match self.find(&k)? {
                    Some(i) => Ok(self.entries.borrow()[i].1.clone()),
                    None => {
                        // Reading a missing key adds it with an Empty value.
                        self.entries.borrow_mut().push((k, Value::Empty));
                        Ok(Value::Empty)
                    }
                }
            }
            "add" => {
                let k = normalize_key(args.require(0, "Key")?);
                if self.find(&k)?.is_some() {
                    return Err(RuntimeError::new(
                        457,
                        "This key is already associated with an element of this collection",
                    ));
                }
                let v = args.get(1, "Item").cloned().unwrap_or_default();
                self.entries.borrow_mut().push((k, v));
                Ok(Value::Empty)
            }
            "exists" => Ok(Value::Boolean(self.find(&normalize_key(args.require(0, "Key")?))?.is_some())),
            "keys" => Ok(Value::Array(Rc::new(RefCell::new(Array::from_vec(
                0,
                self.entries.borrow().iter().map(|(k, _)| k.clone()).collect(),
            ))))),
            "items" => Ok(Value::Array(Rc::new(RefCell::new(Array::from_vec(
                0,
                self.entries.borrow().iter().map(|(_, v)| v.clone()).collect(),
            ))))),
            "remove" => {
                let k = normalize_key(args.require(0, "Key")?);
                match self.find(&k)? {
                    Some(i) => {
                        self.entries.borrow_mut().remove(i);
                        Ok(Value::Empty)
                    }
                    None => Err(RuntimeError::new(32811, "Element not found")),
                }
            }
            "removeall" => {
                self.entries.borrow_mut().clear();
                Ok(Value::Empty)
            }
            "comparemode" => Ok(Value::Long(i32::from(*self.text_mode.borrow()))),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "" | "item" => {
                let k = normalize_key(args.require(0, "Key")?);
                match self.find(&k)? {
                    Some(i) => self.entries.borrow_mut()[i].1 = value,
                    None => self.entries.borrow_mut().push((k, value)),
                }
                Ok(())
            }
            "comparemode" => {
                *self.text_mode.borrow_mut() = value.to_i32()? == 1;
                Ok(())
            }
            "key" => {
                let old = normalize_key(args.require(0, "Key")?);
                let i = self.find(&old)?.ok_or_else(|| RuntimeError::new(32811, "Element not found"))?;
                self.entries.borrow_mut()[i].0 = normalize_key(&value);
                Ok(())
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn items(&self) -> Result<Vec<Value>> {
        // For Each over a dictionary walks its keys.
        Ok(self.entries.borrow().iter().map(|(k, _)| k.clone()).collect())
    }
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;

    #[test]
    fn collection_keys_and_order() {
        let c = Collection::default();
        c.call("Add", &Args::positional(vec![Value::Long(10), "a".into()])).unwrap();
        c.call("Add", &Args::positional(vec![Value::Long(20)])).unwrap();
        c.call("Add", &Args { positional: vec![Value::Long(5)], named: vec![("Before".into(), Value::Long(1))] })
            .unwrap();
        assert!(matches!(c.call("Item", &Args::positional(vec!["A".into()])).unwrap(), Value::Long(10)));
        assert!(matches!(c.call("", &Args::positional(vec![Value::Long(1)])).unwrap(), Value::Long(5)));
        assert!(c.call("Add", &Args::positional(vec![Value::Long(1), "a".into()])).is_err());
        assert_eq!(c.items().unwrap().len(), 3);
    }

    #[test]
    fn dictionary_semantics() {
        let d = Dictionary::default();
        d.call("Add", &Args::positional(vec!["k".into(), Value::Long(1)])).unwrap();
        assert!(matches!(d.call("Exists", &Args::positional(vec!["K".into()])).unwrap(), Value::Boolean(false)));
        d.assign("CompareMode", &Args::none(), Value::Long(1)).unwrap();
        assert!(matches!(d.call("Exists", &Args::positional(vec!["K".into()])).unwrap(), Value::Boolean(true)));
        d.assign("", &Args::positional(vec![Value::Integer(3)]), "x".into()).unwrap();
        assert!(matches!(d.call("Exists", &Args::positional(vec![Value::Double(3.0)])).unwrap(), Value::Boolean(true)));
        assert_eq!(d.items().unwrap().len(), 2);
    }
}
