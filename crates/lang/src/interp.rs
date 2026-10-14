//! The interpreter: runs procedures of parsed modules against a [`Host`].
//!
//! Variables live in shared slots so `ByRef` arguments alias the caller's variable. Statements
//! run block by block; `GoTo`, `Exit` and error handlers travel up the block stack as [`Flow`]
//! values until the block that owns the label or loop handles them.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::ast::*;
use crate::builtins::{Collection, Dictionary, ErrObject};
use crate::error::{Result, RuntimeError};
use crate::object::{Args, Host, MISSING, Object, ObjectRef, same_object};
use crate::ops::{self, BinOp};
use crate::parser::parse_module;
use crate::value::{Array, Value};

/// Declared type of a variable, used to coerce values on assignment.
#[derive(Debug, Clone, PartialEq)]
pub enum Ty {
    Variant,
    Byte,
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
    Boolean,
    String,
    Date,
    /// Any object type (`Object`, `Range`, `Collection`, a class...).
    Object(String),
}

impl Ty {
    pub fn from_name(name: Option<&str>) -> Ty {
        let Some(n) = name else { return Ty::Variant };
        match n.to_ascii_lowercase().as_str() {
            "variant" | "any" => Ty::Variant,
            "byte" => Ty::Byte,
            "integer" => Ty::Integer,
            "long" | "longptr" => Ty::Long,
            "longlong" => Ty::LongLong,
            "single" => Ty::Single,
            "double" => Ty::Double,
            "currency" | "decimal" => Ty::Currency,
            "boolean" => Ty::Boolean,
            "string" => Ty::String,
            "date" => Ty::Date,
            other => Ty::Object(other.to_owned()),
        }
    }

    pub fn default_value(&self) -> Value {
        match self {
            Ty::Variant => Value::Empty,
            Ty::Byte => Value::Byte(0),
            Ty::Integer => Value::Integer(0),
            Ty::Long => Value::Long(0),
            Ty::LongLong => Value::LongLong(0),
            Ty::Single => Value::Single(0.0),
            Ty::Double => Value::Double(0.0),
            Ty::Currency => Value::Currency(0),
            Ty::Boolean => Value::Boolean(false),
            Ty::String => Value::str(""),
            Ty::Date => Value::Date(0.0),
            Ty::Object(_) => Value::Object(None),
        }
    }

    /// Let-coercion of a value into this type.
    pub fn coerce(&self, v: Value) -> Result<Value> {
        let v = match (&v, self) {
            (Value::Object(_), Ty::Object(_)) | (Value::Object(_), Ty::Variant) => return Ok(v),
            (Value::Object(_), _) => v.deref_default()?,
            _ => v,
        };
        Ok(match self {
            Ty::Variant => v,
            Ty::Byte => Value::Byte(u8::try_from(v.to_i64()?).map_err(|_| RuntimeError::overflow())?),
            Ty::Integer => Value::Integer(v.to_i16()?),
            Ty::Long => Value::Long(v.to_i32()?),
            Ty::LongLong => Value::LongLong(v.to_i64()?),
            Ty::Single => Value::Single(v.to_f64()? as f32),
            Ty::Double => Value::Double(v.to_f64()?),
            Ty::Currency => Value::Currency((v.to_f64()? * 10_000.0).round() as i64),
            Ty::Boolean => Value::Boolean(v.to_bool()?),
            Ty::String => match v {
                Value::String(_) => v,
                other => Value::str(other.to_text()?),
            },
            Ty::Date => match v {
                Value::Date(_) => v,
                Value::String(s) => Value::Date(
                    crate::lexer::parse_date_literal(&s)
                        .or_else(|| crate::value::parse_number(&s))
                        .ok_or_else(RuntimeError::type_mismatch)?,
                ),
                other => Value::Date(other.to_f64()?),
            },
            Ty::Object(_) => match v {
                Value::Object(_) | Value::Empty => Value::Object(None),
                _ => return Err(RuntimeError::object_required()),
            },
        })
    }
}

#[derive(Debug, Clone)]
pub struct Var {
    pub value: Value,
    pub ty: Ty,
}

pub type Slot = Rc<RefCell<Var>>;

fn slot(value: Value, ty: Ty) -> Slot {
    Rc::new(RefCell::new(Var { value, ty }))
}

/// A running instance of a class module.
pub struct ClassInstance {
    pub module: usize,
    pub vars: RefCell<HashMap<String, Slot>>,
    pub class_name: String,
}

impl Object for ClassInstance {
    fn type_name(&self) -> String {
        self.class_name.clone()
    }
    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        // Members of class instances are dispatched by the interpreter; reaching here means a
        // host object called back into a class, which is not supported.
        Err(RuntimeError::no_member(name))
    }
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

/// Control flow leaving a statement.
#[derive(Debug, Clone, PartialEq)]
enum Flow {
    Normal,
    Exit(ExitKind),
    GoTo(String),
    GoSub(String),
    /// `Resume Next` / `Resume` from an error handler.
    ResumeNext,
    ResumeSame,
    Return,
    End,
}

struct Frame {
    module: usize,
    locals: HashMap<String, Slot>,
    with_stack: Vec<Value>,
    on_error: OnError,
    in_handler: bool,
    instance: Option<Rc<ClassInstance>>,
    proc_key: String,
    /// Array elements passed ByRef, written back when the call returns.
    writebacks: Vec<(Rc<RefCell<Array>>, Vec<i64>, Slot)>,
}

/// Limits that keep a hostile or buggy macro from hanging or exhausting the application.
#[derive(Debug, Clone)]
pub struct Limits {
    pub max_steps: u64,
    pub max_depth: u32,
}

impl Default for Limits {
    fn default() -> Self {
        Self { max_steps: 200_000_000, max_depth: 1_000 }
    }
}

pub struct Interpreter<'h> {
    host: &'h mut dyn Host,
    pub modules: Vec<Rc<Module>>,
    module_vars: Vec<HashMap<String, Slot>>,
    module_consts: Vec<HashMap<String, Value>>,
    enum_values: HashMap<String, Value>,
    statics: HashMap<String, HashMap<String, Slot>>,
    instances: Vec<std::rc::Weak<ClassInstance>>,
    event_depth: u32,
    pub err: Rc<ErrObject>,
    steps: u64,
    depth: u32,
    pub limits: Limits,
}

fn key(name: &str) -> String {
    name.to_ascii_lowercase()
}

impl<'h> Interpreter<'h> {
    pub fn new(host: &'h mut dyn Host) -> Self {
        Self {
            host,
            modules: Vec::new(),
            module_vars: Vec::new(),
            module_consts: Vec::new(),
            enum_values: HashMap::new(),
            statics: HashMap::new(),
            instances: Vec::new(),
            event_depth: 0,
            err: Rc::new(ErrObject::default()),
            steps: 0,
            depth: 0,
            limits: Limits::default(),
        }
    }

    /// Parses and adds a module. Module-level variables are created immediately.
    pub fn add_module(&mut self, name: &str, kind: ModuleType, source: &str) -> Result<()> {
        let m = parse_module(name, kind, source)?;
        self.add_parsed(m)
    }

    pub fn add_parsed(&mut self, m: Module) -> Result<()> {
        let idx = self.modules.len();
        let base = m.options.base;
        let mut vars = HashMap::new();
        let m = Rc::new(m);
        self.modules.push(m.clone());
        self.module_consts.push(HashMap::new());
        self.module_vars.push(HashMap::new());
        if m.kind != ModuleType::Class {
            for (_, spec) in &m.vars {
                let v = self.new_var(spec, base, idx)?;
                vars.insert(key(&spec.name), v);
            }
        }
        self.module_vars[idx] = vars;
        for (ename, members) in &m.enums {
            let mut next = 0i64;
            for (n, e) in &members.clone() {
                if let Some(e) = e {
                    next = self.const_eval(e, idx)?.to_i64()?;
                }
                let v = Value::Long(next as i32);
                self.enum_values.insert(key(n), v.clone());
                self.enum_values.insert(format!("{}.{}", key(ename), key(n)), v);
                next += 1;
            }
        }
        for (_, n, e) in &m.consts {
            let v = self.const_eval(e, idx)?;
            self.module_consts[idx].insert(key(n), v);
        }
        Ok(())
    }

    fn const_eval(&mut self, e: &Expr, module: usize) -> Result<Value> {
        let mut f = Frame::new(module, None, String::new());
        self.eval(e, &mut f)
    }

    fn new_var(&mut self, spec: &VarSpec, base: i64, module: usize) -> Result<Slot> {
        let ty = Ty::from_name(spec.ty.as_deref());
        let value = match &spec.dims {
            None if spec.new => self.instantiate(spec.ty.as_deref().unwrap_or("Object"))?,
            None => ty.default_value(),
            Some(dims) if dims.is_empty() => {
                Value::Array(Rc::new(RefCell::new(Array { elem_type: elem_name(spec), ..Array::default() })))
            }
            Some(dims) => {
                let mut f = Frame::new(module, None, String::new());
                let bounds = self.eval_bounds(dims, base, &mut f)?;
                let mut a = Array::new(bounds)?;
                let d = ty.default_value();
                for slot in a.data.iter_mut() {
                    *slot = d.clone();
                }
                a.fixed = true;
                a.elem_type = elem_name(spec);
                Value::Array(Rc::new(RefCell::new(a)))
            }
        };
        Ok(slot(value, if spec.dims.is_some() { Ty::Variant } else { ty }))
    }

    fn eval_bounds(&mut self, dims: &[(Option<Expr>, Expr)], base: i64, f: &mut Frame) -> Result<Vec<(i64, i64)>> {
        let mut out = Vec::with_capacity(dims.len());
        for (lo, hi) in dims {
            let lo = match lo {
                Some(e) => self.eval_value(e, f)?.to_i64()?,
                None => base,
            };
            let hi = self.eval_value(hi, f)?.to_i64()?;
            out.push((lo, hi));
        }
        Ok(out)
    }

    /// Finds a procedure by name (`Proc` or `Module.Proc`).
    pub fn find_proc(&self, name: &str) -> Option<(usize, usize)> {
        let (module, proc) = match name.split_once('.') {
            Some((m, p)) => (Some(m), p),
            None => (None, name),
        };
        for (mi, m) in self.modules.iter().enumerate() {
            if module.is_some_and(|mn| !mn.eq_ignore_ascii_case(&m.name)) {
                continue;
            }
            if m.kind == ModuleType::Class {
                continue;
            }
            if let Some(pi) = m.procs.iter().position(|p| {
                p.name.eq_ignore_ascii_case(proc) && p.kind != ProcKind::PropertyLet && p.kind != ProcKind::PropertySet
            }) {
                return Some((mi, pi));
            }
        }
        None
    }

    /// Runs an event handler (`ThisWorkbook.Workbook_Open`, `Sheet1.Worksheet_Change`...) if the
    /// module has it. Returns `Ok(None)` when there is no such handler.
    pub fn fire(&mut self, module: &str, procedure: &str, args: Vec<Value>) -> Result<Option<Value>> {
        let Some(mi) = self.module_index(module) else { return Ok(None) };
        let Some(pi) = self.modules[mi].procs.iter().position(|p| p.name.eq_ignore_ascii_case(procedure)) else {
            return Ok(None);
        };
        let slots = args.into_iter().map(|v| slot(v, Ty::Variant)).collect();
        match self.invoke(mi, pi, slots, Vec::new(), None) {
            Err(e) if e.number == END_SIGNAL => Ok(Some(Value::Empty)),
            r => r.map(Some),
        }
    }

    /// Runs handlers for events the host queued. Handlers that change cells queue more events;
    /// nesting is capped like the application caps recursive event chains.
    fn dispatch_events(&mut self) -> Result<()> {
        if self.event_depth >= 16 {
            return Ok(());
        }
        let events = self.host.take_events();
        for e in events {
            self.event_depth += 1;
            let r = self.fire(&e.module, &e.procedure, e.args);
            self.event_depth -= 1;
            r?;
        }
        Ok(())
    }

    /// Runs a procedure by name with already-evaluated arguments.
    pub fn run(&mut self, name: &str, args: Vec<Value>) -> Result<Value> {
        let (mi, pi) = self
            .find_proc(name)
            .ok_or_else(|| RuntimeError::new(35, format!("Sub or Function not defined: {name}")))?;
        let slots = args.into_iter().map(|v| slot(v, Ty::Variant)).collect();
        match self.invoke(mi, pi, slots, Vec::new(), None) {
            Err(e) if e.number == END_SIGNAL => Ok(Value::Empty),
            r => r,
        }
    }

    // ------------------------------------------------------------ procedure calls

    /// Calls procedure `pi` of module `mi`. `args` are positional slots (aliased for ByRef);
    /// `named` are by-name arguments.
    fn invoke(
        &mut self,
        mi: usize,
        pi: usize,
        args: Vec<Slot>,
        named: Vec<(String, Slot)>,
        instance: Option<Rc<ClassInstance>>,
    ) -> Result<Value> {
        self.depth += 1;
        if self.depth > self.limits.max_depth {
            self.depth -= 1;
            return Err(RuntimeError::standard(28));
        }
        let r = self.invoke_inner(mi, pi, args, named, instance);
        self.depth -= 1;
        r
    }

    fn invoke_inner(
        &mut self,
        mi: usize,
        pi: usize,
        mut args: Vec<Slot>,
        named: Vec<(String, Slot)>,
        instance: Option<Rc<ClassInstance>>,
    ) -> Result<Value> {
        let module = self.modules[mi].clone();
        let proc = &module.procs[pi];
        let proc_key = format!("{}.{}", key(&module.name), key(&proc.name));
        let mut f = Frame::new(mi, instance, proc_key.clone());

        // Bind parameters.
        let mut positional = std::mem::take(&mut args).into_iter();
        for p in &proc.params {
            let k = key(&p.name);
            if p.param_array {
                let rest: Vec<Value> = positional.by_ref().map(|s| s.borrow().value.clone()).collect();
                let arr = Array::from_vec(0, rest);
                f.locals.insert(k, slot(Value::Array(Rc::new(RefCell::new(arr))), Ty::Variant));
                continue;
            }
            let given = named
                .iter()
                .find(|(n, _)| n.eq_ignore_ascii_case(&p.name))
                .map(|(_, s)| s.clone())
                .or_else(|| positional.next());
            let given = given.filter(|s| !matches!(s.borrow().value, Value::Error(448)));
            let ty = if p.is_array { Ty::Variant } else { Ty::from_name(p.ty.as_deref()) };
            let s = match given {
                Some(s) if !p.by_val && s.borrow().ty == ty => s,
                Some(s) if !p.by_val && ty == Ty::Variant => s,
                Some(s) => {
                    let v = s.borrow().value.clone();
                    slot(ty.coerce(v)?, ty)
                }
                None if p.optional => {
                    let v = match &p.default {
                        Some(e) => {
                            let v = self.eval_value(e, &mut f)?;
                            ty.coerce(v)?
                        }
                        None if ty == Ty::Variant => MISSING,
                        None => ty.default_value(),
                    };
                    slot(v, ty)
                }
                None => return Err(RuntimeError::standard(449)),
            };
            f.locals.insert(k, s);
        }
        if positional.next().is_some() {
            return Err(RuntimeError::standard(450));
        }
        // The function's own name is its return variable.
        if matches!(proc.kind, ProcKind::Function | ProcKind::PropertyGet) {
            let ty = Ty::from_name(proc.ret_ty.as_deref());
            f.locals.insert(key(&proc.name), slot(ty.default_value(), ty));
        }
        if proc.is_static {
            let statics = self.statics.entry(proc_key.clone()).or_default();
            for (k, v) in statics.iter() {
                f.locals.entry(k.clone()).or_insert_with(|| v.clone());
            }
        }

        let flow = self.exec_block(&proc.body, &mut f)?;
        if flow == Flow::End {
            return Err(RuntimeError::new(END_SIGNAL, "End"));
        }
        if let Flow::GoTo(l) = flow {
            return Err(RuntimeError::new(1003, format!("Label not defined: {l}")));
        }
        Ok(f.locals.get(&key(&proc.name)).map(|s| s.borrow().value.clone()).unwrap_or_default())
    }

    /// Builds argument slots for a user procedure: variables and array elements are passed by
    /// reference (array elements through a temporary written back after the call).
    fn arg_slots(&mut self, args: &[Arg], f: &mut Frame) -> Result<(Vec<Slot>, Vec<(String, Slot)>)> {
        let mut pos = Vec::new();
        let mut named = Vec::new();
        for a in args {
            let s = match &a.value {
                None => slot(MISSING, Ty::Variant),
                Some(Expr::Ident(n)) => match self.lookup_slot(n, f) {
                    Some(s) => s,
                    None => slot(self.eval(&Expr::Ident(n.clone()), f)?, Ty::Variant),
                },
                Some(Expr::Call(callee, idx)) => match self.array_of(callee, f) {
                    Some(arr) => {
                        let ix =
                            self.eval_args(idx, f)?.positional.iter().map(Value::to_i64).collect::<Result<Vec<_>>>()?;
                        let v = arr.borrow().get(&ix)?;
                        let temp = slot(v, Ty::Variant);
                        f.writebacks.push((arr, ix, temp.clone()));
                        temp
                    }
                    None => slot(self.eval(a.value.as_ref().map_or(&Expr::Lit(Value::Empty), |e| e), f)?, Ty::Variant),
                },
                Some(e) => slot(self.eval(e, f)?, Ty::Variant),
            };
            match &a.name {
                Some(n) => named.push((n.clone(), s)),
                None => pos.push(s),
            }
        }
        Ok((pos, named))
    }

    /// The array held by a variable named by `callee`, if any.
    fn array_of(&self, callee: &Expr, f: &Frame) -> Option<Rc<RefCell<Array>>> {
        let Expr::Ident(n) = callee else { return None };
        match self.lookup_slot(n, f).map(|s| s.borrow().value.clone()) {
            Some(Value::Array(a)) => Some(a),
            _ => None,
        }
    }

    /// Calls a user procedure with unevaluated arguments, honouring ByRef for array elements.
    fn call_user(
        &mut self,
        mi: usize,
        pi: usize,
        args: &[Arg],
        f: &mut Frame,
        instance: Option<Rc<ClassInstance>>,
    ) -> Result<Value> {
        let mark = f.writebacks.len();
        let (pos, named) = self.arg_slots(args, f)?;
        let r = self.invoke(mi, pi, pos, named, instance);
        let pending: Vec<_> = f.writebacks.drain(mark..).collect();
        for (arr, ix, temp) in pending {
            let v = temp.borrow().value.clone();
            arr.borrow_mut().set(&ix, v)?;
        }
        r
    }

    fn eval_args(&mut self, args: &[Arg], f: &mut Frame) -> Result<Args> {
        let mut out = Args::default();
        for a in args {
            let v = match &a.value {
                None => MISSING,
                Some(e) => self.eval(e, f)?,
            };
            match &a.name {
                Some(n) => out.named.push((n.clone(), v)),
                None => out.positional.push(v),
            }
        }
        Ok(out)
    }

    // ------------------------------------------------------------ statements

    fn step(&mut self) -> Result<()> {
        self.steps += 1;
        if self.steps > self.limits.max_steps {
            return Err(RuntimeError::limit("too many steps (possible endless loop)"));
        }
        Ok(())
    }

    fn exec_block(&mut self, stmts: &[Stmt], f: &mut Frame) -> Result<Flow> {
        let mut i = 0usize;
        // Where execution was when an error jumped to a handler in this block.
        let mut resume_at: Option<usize> = None;
        // Return points of GoSub calls made from this block.
        let mut gosub_stack: Vec<usize> = Vec::new();
        while i < stmts.len() {
            self.step()?;
            let mut outcome = self.exec_stmt(&stmts[i], f);
            if outcome.is_ok() {
                if let Err(e) = self.dispatch_events() {
                    outcome = Err(e);
                }
            }
            let flow = match outcome {
                Ok(flow) => flow,
                Err(e) if e.number == END_SIGNAL || e.number == -1 => return Err(e),
                Err(e) => {
                    self.err.set(&e);
                    match f.on_error.clone() {
                        OnError::ResumeNext if !f.in_handler => {
                            i += 1;
                            continue;
                        }
                        OnError::GoTo(label) if !f.in_handler => {
                            if let Some(j) = find_label(stmts, &label) {
                                f.in_handler = true;
                                resume_at = Some(i);
                                i = j + 1;
                                continue;
                            }
                            // The handler is in an enclosing block: let it take the error.
                            return Err(e);
                        }
                        _ => return Err(e),
                    }
                }
            };
            match flow {
                Flow::Normal => i += 1,
                Flow::GoTo(label) => match find_label(stmts, &label) {
                    Some(j) => i = j + 1,
                    None => return Ok(Flow::GoTo(label)),
                },
                Flow::GoSub(label) => match find_label(stmts, &label) {
                    Some(j) => {
                        if gosub_stack.len() > 10_000 {
                            return Err(RuntimeError::standard(28));
                        }
                        gosub_stack.push(i + 1);
                        i = j + 1;
                    }
                    None => return Ok(Flow::GoSub(label)),
                },
                Flow::Return => match gosub_stack.pop() {
                    Some(r) => i = r,
                    None => return Ok(Flow::Return),
                },
                Flow::ResumeNext | Flow::ResumeSame if resume_at.is_some() => {
                    let at = resume_at.take().unwrap_or(0);
                    f.in_handler = false;
                    self.err.clear();
                    i = if flow == Flow::ResumeNext { at + 1 } else { at };
                }
                other => return Ok(other),
            }
        }
        Ok(Flow::Normal)
    }

    fn exec_stmt(&mut self, s: &Stmt, f: &mut Frame) -> Result<Flow> {
        match &s.kind {
            StmtKind::Dim { vars, is_static } => {
                let base = self.modules[f.module].options.base;
                for spec in vars {
                    let k = key(&spec.name);
                    if *is_static {
                        if let Some(existing) = self.statics.get(&f.proc_key).and_then(|m| m.get(&k)) {
                            f.locals.insert(k, existing.clone());
                            continue;
                        }
                        let v = self.new_var_in(spec, base, f)?;
                        self.statics.entry(f.proc_key.clone()).or_default().insert(k.clone(), v.clone());
                        f.locals.insert(k, v);
                    } else {
                        let v = self.new_var_in(spec, base, f)?;
                        f.locals.insert(k, v);
                    }
                }
            }
            StmtKind::ReDim { preserve, items } => {
                let base = self.modules[f.module].options.base;
                for (target, dims) in items {
                    let bounds = self.eval_bounds(dims, base, f)?;
                    let name = match target {
                        Expr::Ident(n) => n.clone(),
                        _ => return Err(RuntimeError::invalid_call()),
                    };
                    let s = match self.lookup_slot(&name, f) {
                        Some(s) => s,
                        None => {
                            let s = slot(Value::Empty, Ty::Variant);
                            f.locals.insert(key(&name), s.clone());
                            s
                        }
                    };
                    let mut var = s.borrow_mut();
                    match &var.value {
                        Value::Array(a) if *preserve => {
                            if a.borrow().fixed {
                                return Err(RuntimeError::standard(10));
                            }
                            let elem = a.borrow().elem_type.clone();
                            let old_len = a.borrow().data.len();
                            a.borrow_mut().redim_preserve(bounds)?;
                            if elem.is_some() {
                                let d = Ty::from_name(elem.as_deref()).default_value();
                                let mut arr = a.borrow_mut();
                                for x in arr.data.iter_mut().filter(|x| x.is_empty()) {
                                    *x = d.clone();
                                }
                                let _ = old_len;
                            }
                        }
                        Value::Array(a) if a.borrow().fixed => return Err(RuntimeError::standard(10)),
                        _ => {
                            let elem = match &var.value {
                                Value::Array(a) => a.borrow().elem_type.clone(),
                                _ => None,
                            };
                            let mut arr = Array::new(bounds)?;
                            if elem.is_some() {
                                let d = Ty::from_name(elem.as_deref()).default_value();
                                arr.data.iter_mut().for_each(|x| *x = d.clone());
                            }
                            arr.elem_type = elem;
                            var.value = Value::Array(Rc::new(RefCell::new(arr)));
                        }
                    }
                }
            }
            StmtKind::Const(list) => {
                for (n, e) in list {
                    let v = self.eval_value(e, f)?;
                    f.locals.insert(key(n), slot(v, Ty::Variant));
                }
            }
            StmtKind::Let { target, value } => {
                // `Range("B1:B3") = Range("A1:A3")`: an application object assigned to an
                // application object's default property is passed as the object; the host
                // decides what that means. Everything else gets the value.
                let host_target = match target {
                    Expr::Call(callee, _) => match callee.as_ref() {
                        Expr::Ident(name) => {
                            self.lookup_slot(name, f).is_none() && self.resolve_proc(name, f).is_none()
                        }
                        Expr::Member(..) => true,
                        _ => false,
                    },
                    Expr::Member(Some(_), _) => true,
                    _ => false,
                };
                let raw = self.eval(value, f)?;
                let v = match raw {
                    Value::Object(Some(ref o)) if host_target && o.as_any().is_none() => raw,
                    Value::Object(Some(ref o))
                        if o.as_any().is_some_and(|a| {
                            a.is::<ClassInstance>() || a.is::<Collection>() || a.is::<Dictionary>() || a.is::<Record>()
                        }) =>
                    {
                        raw
                    }
                    Value::Object(Some(_)) => raw.deref_default()?,
                    other => other,
                };
                self.assign(target, v, false, f)?;
            }
            StmtKind::Set { target, value } => {
                let v = self.eval(value, f)?;
                let v = match v {
                    Value::Object(_) => v,
                    Value::Empty => Value::Object(None),
                    _ => return Err(RuntimeError::object_required()),
                };
                self.assign(target, v, true, f)?;
            }
            StmtKind::Call(e) => {
                self.call_statement(e, f)?;
            }
            StmtKind::If { arms, otherwise } => {
                for (cond, body) in arms {
                    if self.truthy(cond, f)? {
                        return self.exec_block(body, f);
                    }
                }
                if let Some(body) = otherwise {
                    return self.exec_block(body, f);
                }
            }
            StmtKind::Select { subject, cases, otherwise } => {
                let v = self.eval_value(subject, f)?;
                let text = self.modules[f.module].options.compare_text;
                for (clauses, body) in cases {
                    for c in clauses {
                        let hit = match c {
                            CaseClause::Value(e) => {
                                let x = self.eval_value(e, f)?;
                                ops::binary(BinOp::Eq, &v, &x, text)?.to_bool().unwrap_or(false)
                            }
                            CaseClause::Range(a, b) => {
                                let (x, y) = (self.eval_value(a, f)?, self.eval_value(b, f)?);
                                ops::binary(BinOp::Ge, &v, &x, text)?.to_bool().unwrap_or(false)
                                    && ops::binary(BinOp::Le, &v, &y, text)?.to_bool().unwrap_or(false)
                            }
                            CaseClause::Is(op, e) => {
                                let x = self.eval_value(e, f)?;
                                ops::binary(*op, &v, &x, text)?.to_bool().unwrap_or(false)
                            }
                        };
                        if hit {
                            return self.exec_block(body, f);
                        }
                    }
                }
                if let Some(body) = otherwise {
                    return self.exec_block(body, f);
                }
            }
            StmtKind::For { var, from, to, step, body } => {
                let start = self.eval_value(from, f)?;
                let end = self.eval_value(to, f)?;
                let step_v = match step {
                    Some(e) => self.eval_value(e, f)?,
                    None => Value::Integer(1),
                };
                let (endf, stepf) = (end.to_f64()?, step_v.to_f64()?);
                self.assign(var, start, false, f)?;
                loop {
                    let cur = self.eval_value(var, f)?.to_f64()?;
                    if (stepf >= 0.0 && cur > endf) || (stepf < 0.0 && cur < endf) {
                        break;
                    }
                    match self.exec_block(body, f)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    let cur = self.eval_value(var, f)?;
                    let next = ops::binary(BinOp::Add, &cur, &step_v, false)?;
                    self.assign(var, next, false, f)?;
                    self.step()?;
                }
            }
            StmtKind::ForEach { var, collection, body } => {
                let coll = self.eval(collection, f)?;
                let items = match &coll {
                    Value::Array(a) => a.borrow().data.clone(),
                    Value::Object(Some(o)) => o.items()?,
                    Value::Object(None) => return Err(RuntimeError::object_not_set()),
                    _ => return Err(RuntimeError::new(92, "For loop not initialized")),
                };
                for item in items {
                    let is_obj = item.is_object();
                    self.assign(var, item, is_obj, f)?;
                    match self.exec_block(body, f)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    self.step()?;
                }
            }
            StmtKind::Do { pre, post, body } => loop {
                if let Some((until, cond)) = pre {
                    if self.truthy(cond, f)? == *until {
                        break;
                    }
                }
                match self.exec_block(body, f)? {
                    Flow::Normal => {}
                    Flow::Exit(ExitKind::Do) => break,
                    other => return Ok(other),
                }
                if let Some((until, cond)) = post {
                    if self.truthy(cond, f)? == *until {
                        break;
                    }
                }
                self.step()?;
            },
            StmtKind::While { cond, body } => {
                while self.truthy(cond, f)? {
                    match self.exec_block(body, f)? {
                        Flow::Normal => {}
                        other => return Ok(other),
                    }
                    self.step()?;
                }
            }
            StmtKind::With { object, body } => {
                let obj = self.eval(object, f)?;
                f.with_stack.push(obj);
                let r = self.exec_block(body, f);
                f.with_stack.pop();
                return r;
            }
            StmtKind::Exit(k) => return Ok(Flow::Exit(*k)),
            StmtKind::GoTo(l) => return Ok(Flow::GoTo(l.clone())),
            StmtKind::GoSub(l) => return Ok(Flow::GoSub(l.clone())),
            StmtKind::Return => return Ok(Flow::Return),
            StmtKind::Label(_) => {}
            StmtKind::OnError(kind) => {
                match kind {
                    OnError::Reset => {
                        f.in_handler = false;
                        self.err.clear();
                    }
                    OnError::Disable => {
                        f.on_error = OnError::Disable;
                        self.err.clear();
                    }
                    k => {
                        f.on_error = k.clone();
                        self.err.clear();
                    }
                }
                if !matches!(kind, OnError::Reset) {
                    f.in_handler = false;
                }
            }
            StmtKind::Resume(r) => {
                return Ok(match r {
                    Resume::Next => Flow::ResumeNext,
                    Resume::Same => Flow::ResumeSame,
                    Resume::Label(l) => {
                        f.in_handler = false;
                        self.err.clear();
                        Flow::GoTo(l.clone())
                    }
                });
            }
            StmtKind::Erase(items) => {
                for e in items {
                    if let Expr::Ident(n) = e {
                        if let Some(s) = self.lookup_slot(n, f) {
                            let mut var = s.borrow_mut();
                            if let Value::Array(a) = &var.value {
                                let fixed = a.borrow().fixed;
                                if fixed {
                                    let mut a = a.borrow_mut();
                                    let blank = a.data.first().map(reset_like).unwrap_or_default();
                                    for x in a.data.iter_mut() {
                                        *x = blank.clone();
                                    }
                                } else {
                                    var.value = Value::Array(Rc::new(RefCell::new(Array::default())));
                                }
                            }
                        }
                    }
                }
            }
            StmtKind::DebugPrint(items) => {
                let mut out = String::new();
                for e in items.iter().flatten() {
                    let v = self.eval_value(e, f)?;
                    out.push_str(&print_text(&v)?);
                }
                self.host.debug_print(&out);
            }
            StmtKind::End => return Ok(Flow::End),
            StmtKind::Stop => return Err(RuntimeError::limit("Stop statement")),
            StmtKind::Unsupported(what) => {
                return Err(RuntimeError::new(
                    70,
                    format!("Permission denied: '{what}' is not available to macros in this application"),
                ));
            }
        }
        Ok(Flow::Normal)
    }

    fn new_var_in(&mut self, spec: &VarSpec, base: i64, f: &mut Frame) -> Result<Slot> {
        let ty = Ty::from_name(spec.ty.as_deref());
        let value = match &spec.dims {
            None if spec.new => self.instantiate(spec.ty.as_deref().unwrap_or("Object"))?,
            None => {
                if let Ty::Object(tn) = &ty {
                    // A variable of a user Type: give it a record object.
                    if let Some(fields) = self.user_type(tn) {
                        return Ok(slot(self.new_record(tn, &fields)?, ty));
                    }
                }
                ty.default_value()
            }
            Some(dims) if dims.is_empty() => {
                Value::Array(Rc::new(RefCell::new(Array { elem_type: elem_name(spec), ..Array::default() })))
            }
            Some(dims) => {
                let bounds = self.eval_bounds(dims, base, f)?;
                let mut a = Array::new(bounds)?;
                let d = ty.default_value();
                for x in a.data.iter_mut() {
                    *x = d.clone();
                }
                a.fixed = true;
                a.elem_type = elem_name(spec);
                Value::Array(Rc::new(RefCell::new(a)))
            }
        };
        Ok(slot(value, if spec.dims.is_some() { Ty::Variant } else { ty }))
    }

    fn user_type(&self, name: &str) -> Option<Vec<VarSpec>> {
        self.modules
            .iter()
            .flat_map(|m| m.types.iter())
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, f)| f.clone())
    }

    fn new_record(&mut self, name: &str, fields: &[VarSpec]) -> Result<Value> {
        let rec = Record { name: name.to_owned(), fields: RefCell::new(Vec::new()) };
        for spec in fields {
            let ty = Ty::from_name(spec.ty.as_deref());
            rec.fields.borrow_mut().push((key(&spec.name), ty.default_value(), ty));
        }
        Ok(Value::Object(Some(Rc::new(rec))))
    }

    fn truthy(&mut self, e: &Expr, f: &mut Frame) -> Result<bool> {
        match self.eval_value(e, f)? {
            Value::Null => Ok(false),
            v => v.to_bool(),
        }
    }

    // ------------------------------------------------------------ names

    fn lookup_slot(&self, name: &str, f: &Frame) -> Option<Slot> {
        let k = key(name);
        if let Some(s) = f.locals.get(&k) {
            return Some(s.clone());
        }
        if let Some(inst) = &f.instance {
            if let Some(s) = inst.vars.borrow().get(&k) {
                return Some(s.clone());
            }
        }
        if let Some(s) = self.module_vars.get(f.module).and_then(|m| m.get(&k)) {
            return Some(s.clone());
        }
        for (mi, m) in self.modules.iter().enumerate() {
            if mi == f.module || m.kind == ModuleType::Class {
                continue;
            }
            let public =
                m.vars.iter().any(|(vis, spec)| *vis == Visibility::Public && spec.name.eq_ignore_ascii_case(name));
            if public {
                if let Some(s) = self.module_vars[mi].get(&k) {
                    return Some(s.clone());
                }
            }
        }
        None
    }

    fn lookup_const(&self, name: &str, f: &Frame) -> Option<Value> {
        let k = key(name);
        if let Some(v) = self.module_consts.get(f.module).and_then(|m| m.get(&k)) {
            return Some(v.clone());
        }
        for (mi, consts) in self.module_consts.iter().enumerate() {
            if let Some(v) = consts.get(&k) {
                let public = self.modules[mi]
                    .consts
                    .iter()
                    .any(|(vis, n, _)| *vis == Visibility::Public && n.eq_ignore_ascii_case(name));
                if public {
                    return Some(v.clone());
                }
            }
        }
        self.enum_values.get(&k).cloned()
    }

    fn module_index(&self, name: &str) -> Option<usize> {
        self.modules.iter().position(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// Procedure in the current module (or class) first, then any standard module.
    fn resolve_proc(&self, name: &str, f: &Frame) -> Option<(usize, usize)> {
        if let Some(pi) = self.modules[f.module].procs.iter().position(|p| {
            p.name.eq_ignore_ascii_case(name) && !matches!(p.kind, ProcKind::PropertyLet | ProcKind::PropertySet)
        }) {
            return Some((f.module, pi));
        }
        self.find_proc(name)
    }

    // ------------------------------------------------------------ expressions

    /// Evaluates for use as data: objects collapse to their default value.
    fn eval_value(&mut self, e: &Expr, f: &mut Frame) -> Result<Value> {
        let v = self.eval(e, f)?;
        match v {
            Value::Object(Some(ref o))
                if o.as_any().is_some_and(|a| {
                    a.is::<ClassInstance>() || a.is::<Collection>() || a.is::<Dictionary>() || a.is::<Record>()
                }) =>
            {
                Ok(v)
            }
            Value::Object(Some(_)) => v.deref_default(),
            v => Ok(v),
        }
    }

    fn eval(&mut self, e: &Expr, f: &mut Frame) -> Result<Value> {
        match e {
            Expr::Lit(v) => Ok(v.clone()),
            Expr::Paren(inner) => self.eval(inner, f),
            Expr::Ident(name) => self.read_name(name, f),
            Expr::Binary(op, a, b) if matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul) => {
                let x = self.eval_value(a, f)?;
                let y = self.eval_value(b, f)?;
                let text = self.modules.get(f.module).is_some_and(|m| m.options.compare_text);
                match ops::binary(*op, &x, &y, text) {
                    // A Variant operand widens the result (Integer -> Long -> Double) instead of
                    // overflowing; declared numeric variables and literals do overflow.
                    Err(e) if e.number == 6 && (self.is_variant_expr(a, f) || self.is_variant_expr(b, f)) => {
                        ops::binary(*op, &widen(&x), &widen(&y), text)
                            .or_else(|_| ops::binary(*op, &widen(&widen(&x)), &widen(&widen(&y)), text))
                    }
                    r => r,
                }
            }
            Expr::Binary(op, a, b) => {
                let x = self.eval_value(a, f)?;
                let y = self.eval_value(b, f)?;
                if matches!(op, BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge) {
                    self.check_typed_comparison(a, &x, b, &y, f)?;
                }
                ops::binary(*op, &x, &y, self.modules.get(f.module).is_some_and(|m| m.options.compare_text))
            }
            Expr::Neg(a) => {
                let x = self.eval_value(a, f)?;
                ops::negate(&x)
            }
            Expr::Not(a) => {
                let x = self.eval_value(a, f)?;
                ops::not(&x)
            }
            Expr::Is(a, b) => {
                let (x, y) = (self.eval(a, f)?, self.eval(b, f)?);
                match (x, y) {
                    (Value::Object(p), Value::Object(q)) => Ok(Value::Boolean(same_object(&p, &q))),
                    _ => Err(RuntimeError::object_required()),
                }
            }
            Expr::TypeOf(a, tn) => {
                let x = self.eval(a, f)?;
                match x {
                    Value::Object(Some(o)) => {
                        Ok(Value::Boolean(tn.eq_ignore_ascii_case("object") || o.type_name().eq_ignore_ascii_case(tn)))
                    }
                    Value::Object(None) => Ok(Value::Boolean(false)),
                    _ => Err(RuntimeError::object_required()),
                }
            }
            Expr::New(name) => self.instantiate(name),
            Expr::Me => match &f.instance {
                Some(i) => Ok(Value::Object(Some(i.clone() as ObjectRef))),
                None => {
                    let name = self.modules[f.module].name.clone();
                    self.host
                        .global(&name, &Args::none())
                        .unwrap_or(Err(RuntimeError::new(1004, "Me is not available here")))
                }
            },
            Expr::Member(obj, name) => self.member(obj.as_deref(), name, &[], f),
            Expr::Call(callee, args) => self.call_expr(callee, args, f),
            Expr::Bang(obj, name) => {
                let o = match obj.as_ref() {
                    Expr::Member(None, n) if n.is_empty() => self.with_object(f)?,
                    other => self.eval(other, f)?,
                };
                self.call_object(&o, "", &Args::positional(vec![Value::str(name.clone())]))
            }
        }
    }

    /// True when an expression is Variant-typed (so arithmetic on it widens on overflow).
    fn is_variant_expr(&self, e: &Expr, f: &Frame) -> bool {
        match e {
            Expr::Lit(_) => false,
            Expr::Paren(inner) | Expr::Neg(inner) => self.is_variant_expr(inner, f),
            Expr::Binary(_, a, b) => self.is_variant_expr(a, f) || self.is_variant_expr(b, f),
            Expr::Ident(n) => match self.lookup_slot(n, f) {
                Some(s) => s.borrow().ty == Ty::Variant,
                None => self.lookup_const(n, f).is_none(),
            },
            Expr::Call(callee, _) => match callee.as_ref() {
                Expr::Ident(n) => match self.lookup_slot(n, f).map(|s| s.borrow().value.clone()) {
                    Some(Value::Array(a)) => a.borrow().elem_type.is_none(),
                    Some(_) => true,
                    None => match self.resolve_proc(n, f) {
                        Some((mi, pi)) => self.modules[mi].procs[pi]
                            .ret_ty
                            .as_deref()
                            .is_none_or(|t| t.eq_ignore_ascii_case("variant")),
                        // Library conversion functions return a fixed type.
                        None => !is_typed_builtin(n),
                    },
                },
                _ => true,
            },
            _ => true,
        }
    }

    /// A comparison between a statically typed number (or Boolean) and a statically typed
    /// String converts the string to a number, so a non-numeric string is a type mismatch. (With
    /// a Variant operand the values compare instead; see `ops::compare`.)
    fn check_typed_comparison(&self, a: &Expr, x: &Value, b: &Expr, y: &Value, f: &Frame) -> Result<()> {
        let (Some(ka), Some(kb)) = (self.static_kind(a, f), self.static_kind(b, f)) else { return Ok(()) };
        let (num, text) = match (ka, kb) {
            (StaticKind::Num, StaticKind::Str) => (x, y),
            (StaticKind::Str, StaticKind::Num) => (y, x),
            _ => return Ok(()),
        };
        let Value::String(t) = text else { return Ok(()) };
        let convertible = match num {
            Value::Boolean(_) => Value::String(t.clone()).to_bool().is_ok(),
            Value::Null => true,
            _ => crate::value::parse_number(t).is_some(),
        };
        if convertible { Ok(()) } else { Err(RuntimeError::type_mismatch()) }
    }

    /// The declared kind of an expression when it is known statically: literals, typed
    /// variables and constants, comparison and concatenation results. `None` for Variant.
    fn static_kind(&self, e: &Expr, f: &Frame) -> Option<StaticKind> {
        let of_value = |v: &Value| match v {
            Value::String(_) => Some(StaticKind::Str),
            Value::Boolean(_)
            | Value::Byte(_)
            | Value::Integer(_)
            | Value::Long(_)
            | Value::LongLong(_)
            | Value::Single(_)
            | Value::Double(_)
            | Value::Currency(_) => Some(StaticKind::Num),
            _ => None,
        };
        match e {
            Expr::Lit(v) => of_value(v),
            Expr::Paren(inner) => self.static_kind(inner, f),
            Expr::Not(_) => Some(StaticKind::Num),
            Expr::Binary(op, _, _) => match op {
                BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge | BinOp::Like => {
                    Some(StaticKind::Num)
                }
                BinOp::Concat => Some(StaticKind::Str),
                _ => None,
            },
            Expr::Ident(n) => match self.lookup_slot(n, f) {
                Some(s) => match s.borrow().ty {
                    Ty::String => Some(StaticKind::Str),
                    Ty::Variant | Ty::Object(_) | Ty::Date => None,
                    _ => Some(StaticKind::Num),
                },
                None => self.lookup_const(n, f).or_else(|| crate::stdlib::constant(n)).as_ref().and_then(of_value),
            },
            _ => None,
        }
    }

    fn with_object(&self, f: &Frame) -> Result<Value> {
        f.with_stack.last().cloned().ok_or_else(|| RuntimeError::new(1002, "Invalid or unqualified reference"))
    }

    /// `[A1]`, `[Sheet1!B2:C3]`, `[MyName]`: the application's `Evaluate` of the bracket text.
    fn bracket(&mut self, name: &str) -> Result<Value> {
        let inner = name.trim_start_matches('[').trim_end_matches(']').to_owned();
        self.host
            .global("Evaluate", &Args::positional(vec![Value::str(inner)]))
            .unwrap_or_else(|| Err(RuntimeError::new(1004, format!("{name} cannot be evaluated here"))))
    }

    fn read_name(&mut self, name: &str, f: &mut Frame) -> Result<Value> {
        if name.starts_with('[') {
            return self.bracket(name);
        }
        if let Some(s) = self.lookup_slot(name, f) {
            return Ok(s.borrow().value.clone());
        }
        if let Some(v) = self.lookup_const(name, f) {
            return Ok(v);
        }
        let k = key(name);
        if k == "err" {
            return Ok(Value::Object(Some(self.err.clone() as ObjectRef)));
        }
        if k == "erl" {
            // No numbered lines are tracked; the host reports 0 for code without them.
            return Ok(Value::Long(0));
        }
        if let Some((mi, pi)) = self.resolve_proc(name, f) {
            return self.invoke(mi, pi, Vec::new(), Vec::new(), self.instance_for(mi, f));
        }
        if let Some(inst) = f.instance.clone() {
            if let Some(pi) = self.modules[inst.module]
                .procs
                .iter()
                .position(|p| p.name.eq_ignore_ascii_case(name) && p.kind == ProcKind::PropertyGet)
            {
                return self.invoke(inst.module, pi, Vec::new(), Vec::new(), Some(inst));
            }
        }
        if let Some(v) = crate::stdlib::constant(name) {
            return Ok(v);
        }
        if let Some(r) = crate::stdlib::call(name, &[], self.host) {
            return r;
        }
        if let Some(r) = self.host.global(name, &Args::none()) {
            return r;
        }
        if self.modules[f.module].options.explicit {
            return Err(RuntimeError::new(1001, format!("Variable not defined: {name}")));
        }
        // An implicitly declared Variant.
        let s = slot(Value::Empty, Ty::Variant);
        f.locals.insert(k, s);
        Ok(Value::Empty)
    }

    fn instance_for(&self, mi: usize, f: &Frame) -> Option<Rc<ClassInstance>> {
        f.instance.as_ref().filter(|i| i.module == mi).cloned()
    }

    fn instantiate(&mut self, name: &str) -> Result<Value> {
        let short = name.rsplit('.').next().unwrap_or(name);
        match short.to_ascii_lowercase().as_str() {
            "collection" => return Ok(Collection::new_ref()),
            "dictionary" => return Ok(Dictionary::new_ref()),
            _ => {}
        }
        if let Some(mi) =
            self.modules.iter().position(|m| m.kind == ModuleType::Class && m.name.eq_ignore_ascii_case(short))
        {
            let module = self.modules[mi].clone();
            let mut vars = HashMap::new();
            for (_, spec) in &module.vars {
                let v = self.new_var(spec, module.options.base, mi)?;
                vars.insert(key(&spec.name), v);
            }
            let inst = Rc::new(ClassInstance { module: mi, vars: RefCell::new(vars), class_name: module.name.clone() });
            self.instances.retain(|w| w.strong_count() > 0);
            self.instances.push(Rc::downgrade(&inst));
            if let Some(pi) = module.procs.iter().position(|p| p.name.eq_ignore_ascii_case("Class_Initialize")) {
                self.invoke(mi, pi, Vec::new(), Vec::new(), Some(inst.clone()))?;
            }
            return Ok(Value::Object(Some(inst as ObjectRef)));
        }
        if let Some(fields) = self.user_type(short) {
            return self.new_record(short, &fields);
        }
        Err(RuntimeError::new(429, format!("ActiveX component can't create object: {name}")))
    }

    /// `CreateObject(progid)`: only safe, built-in objects are available.
    fn create_object(&mut self, progid: &str) -> Result<Value> {
        match progid.to_ascii_lowercase().as_str() {
            "scripting.dictionary" => Ok(Dictionary::new_ref()),
            "vba.collection" => Ok(Collection::new_ref()),
            _ => Err(RuntimeError::new(
                429,
                format!("ActiveX component can't create object: {progid} (not available to macros here)"),
            )),
        }
    }

    /// `name(args)` / `obj.name(args)` / `expr(args)`.
    fn call_expr(&mut self, callee: &Expr, args: &[Arg], f: &mut Frame) -> Result<Value> {
        match callee {
            Expr::Ident(name) => self.call_name(name, args, f),
            Expr::Member(obj, name) => self.member(obj.as_deref(), name, args, f),
            other => {
                let target = self.eval(other, f)?;
                let a = self.eval_args(args, f)?;
                self.index_or_default(&target, &a)
            }
        }
    }

    fn index_or_default(&mut self, target: &Value, a: &Args) -> Result<Value> {
        match target {
            Value::Array(arr) => {
                let idx = a.positional.iter().map(Value::to_i64).collect::<Result<Vec<_>>>()?;
                arr.borrow().get(&idx)
            }
            Value::Object(_) => self.call_object(target, "", a),
            _ if a.is_empty() => Ok(target.clone()),
            _ => Err(RuntimeError::type_mismatch()),
        }
    }

    fn call_name(&mut self, name: &str, args: &[Arg], f: &mut Frame) -> Result<Value> {
        if let Some(s) = self.lookup_slot(name, f) {
            let v = s.borrow().value.clone();
            // Inside a Function, `F(x)` with F's own name and arguments is recursion.
            let is_self = f.proc_key.rsplit('.').next().is_some_and(|p| p.eq_ignore_ascii_case(name))
                && !matches!(v, Value::Array(_));
            if !is_self {
                let a = self.eval_args(args, f)?;
                return self.index_or_default(&v, &a);
            }
        }
        let k = key(name);
        if let Some((mi, pi)) = self.resolve_proc(name, f) {
            return self.call_user(mi, pi, args, f, self.instance_for(mi, f));
        }
        match k.as_str() {
            "createobject" => {
                let a = self.eval_args(args, f)?;
                let id = a.require(0, "Class")?.to_text()?;
                return self.create_object(&id);
            }
            "getobject" => return Err(RuntimeError::new(429, "GetObject is not available to macros here")),
            "shell" | "environ" | "dir" | "filelen" | "filedatetime" | "getattr" | "curdir" | "freefile" | "eof"
            | "lof" | "loc" | "input" | "kill" | "mkdir" | "rmdir" | "chdir" | "chdrive" | "filecopy" | "sendkeys"
            | "appactivate" | "load" | "unload" | "lock" | "unlock" | "reset" | "setattr" | "savesetting"
            | "getsetting" | "deletesetting" | "getallsettings" => {
                return Err(RuntimeError::new(
                    70,
                    format!("Permission denied: {name} is not available to macros in this application"),
                ));
            }
            "isobject" => {
                let a = self.eval_args(args, f)?;
                return Ok(Value::Boolean(a.positional.first().is_some_and(Value::is_object)));
            }
            "doevents" => return Ok(Value::Long(0)),
            _ => {}
        }
        if let Some(inst) = f.instance.clone() {
            if let Some(pi) = self.modules[inst.module].procs.iter().position(|p| {
                p.name.eq_ignore_ascii_case(name) && !matches!(p.kind, ProcKind::PropertyLet | ProcKind::PropertySet)
            }) {
                return self.call_user(inst.module, pi, args, f, Some(inst));
            }
        }
        let a = self.eval_args(args, f)?;
        if self.modules.iter().flat_map(|m| m.declares.iter()).any(|d| d.eq_ignore_ascii_case(name)) {
            return Err(RuntimeError::new(
                70,
                format!("Permission denied: '{name}' is a native (Declare) call, not available to macros here"),
            ));
        }
        if let Some(r) = crate::stdlib::call(name, &a.positional, self.host) {
            return r;
        }
        if let Some(r) = self.host.global(name, &a) {
            return r;
        }
        Err(RuntimeError::new(35, format!("Sub or Function not defined: {name}")))
    }

    /// `obj.name(args)`; `obj` is `None` inside `With`.
    fn member(&mut self, obj: Option<&Expr>, name: &str, args: &[Arg], f: &mut Frame) -> Result<Value> {
        // Qualified names that are not objects: Module.Proc, Module.Var, VBA.Func, Enum.Member.
        if let Some(Expr::Ident(q)) = obj {
            if self.lookup_slot(q, f).is_none() {
                let qk = key(q);
                if qk == "vba" {
                    let a = self.eval_args(args, f)?;
                    if let Some(r) = crate::stdlib::call(name, &a.positional, self.host) {
                        return r;
                    }
                }
                if let Some(v) = self.enum_values.get(&format!("{qk}.{}", key(name))) {
                    return Ok(v.clone());
                }
                if let Some(mi) = self.module_index(q) {
                    if self.modules[mi].kind != ModuleType::Class {
                        if let Some(pi) = self.modules[mi].procs.iter().position(|p| p.name.eq_ignore_ascii_case(name))
                        {
                            return self.call_user(mi, pi, args, f, None);
                        }
                        if let Some(s) = self.module_vars[mi].get(&key(name)) {
                            let v = s.borrow().value.clone();
                            let a = self.eval_args(args, f)?;
                            return self.index_or_default(&v, &a);
                        }
                        if let Some(v) = self.module_consts[mi].get(&key(name)) {
                            return Ok(v.clone());
                        }
                    }
                }
            }
        }
        let target = match obj {
            None => self.with_object(f)?,
            Some(e) => self.eval(e, f)?,
        };
        // Application.InputBox(Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type):
        // asks the host; Cancel gives False, Type:=1 converts the reply to a number.
        if name.eq_ignore_ascii_case("inputbox") {
            if let Value::Object(Some(o)) = &target {
                if o.type_name().eq_ignore_ascii_case("application") {
                    let a = self.eval_args(args, f)?;
                    let text = |v: Option<&Value>| v.map(Value::to_text).transpose().map(Option::unwrap_or_default);
                    let prompt = text(a.get(0, "Prompt"))?;
                    let title = text(a.get(1, "Title"))?;
                    let default = text(a.get(2, "Default"))?;
                    let kind = a.get(7, "Type").map(Value::to_i64).transpose()?.unwrap_or(2);
                    return match self.host.inputbox(&prompt, &title, &default) {
                        None => Ok(Value::Boolean(false)),
                        Some(reply) if kind & 1 == 1 => crate::value::parse_number(&reply)
                            .map(Value::Double)
                            .ok_or_else(|| RuntimeError::new(1004, "Number is not valid")),
                        Some(reply) if kind & 4 == 4 => Ok(Value::Boolean(Value::str(reply).to_bool()?)),
                        Some(reply) => Ok(Value::str(reply)),
                    };
                }
            }
        }
        // Class instances run their own procedures.
        if let Value::Object(Some(o)) = &target {
            if let Some(inst) = o.as_any().and_then(|a| a.downcast_ref::<ClassInstance>()) {
                let module = self.modules[inst.module].clone();
                if let Some(pi) = module.procs.iter().position(|p| {
                    p.name.eq_ignore_ascii_case(name)
                        && !matches!(p.kind, ProcKind::PropertyLet | ProcKind::PropertySet)
                }) {
                    let inst_rc = self.class_rc(o)?;
                    return self.call_user(inst.module, pi, args, f, Some(inst_rc));
                }
                if let Some(s) = inst.vars.borrow().get(&key(name)) {
                    if module
                        .vars
                        .iter()
                        .any(|(vis, spec)| *vis == Visibility::Public && spec.name.eq_ignore_ascii_case(name))
                    {
                        return Ok(s.borrow().value.clone());
                    }
                }
                return Err(RuntimeError::no_member(name));
            }
        }
        let a = self.eval_args(args, f)?;
        self.call_object(&target, name, &a)
    }

    /// The typed handle of a class instance seen through an `ObjectRef` (same allocation).
    fn class_rc(&self, o: &ObjectRef) -> Result<Rc<ClassInstance>> {
        let addr = Rc::as_ptr(o) as *const ();
        self.instances
            .iter()
            .filter_map(std::rc::Weak::upgrade)
            .find(|i| Rc::as_ptr(i) as *const () == addr)
            .ok_or_else(|| RuntimeError::new(51, "Internal error: unknown class instance"))
    }

    fn call_object(&mut self, target: &Value, name: &str, a: &Args) -> Result<Value> {
        match target {
            Value::Object(Some(o)) => o.call(name, a),
            Value::Object(None) => Err(RuntimeError::object_not_set()),
            Value::Array(_) if name.is_empty() => self.index_or_default(target, a),
            _ => Err(RuntimeError::object_required()),
        }
    }

    fn call_statement(&mut self, e: &Expr, f: &mut Frame) -> Result<()> {
        match e {
            Expr::Ident(name) => {
                if let Some((mi, pi)) = self.resolve_proc(name, f) {
                    self.invoke(mi, pi, Vec::new(), Vec::new(), self.instance_for(mi, f))?;
                    return Ok(());
                }
                self.call_name(name, &[], f)?;
            }
            other => {
                self.eval(other, f)?;
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------ assignment

    fn store(&mut self, s: &Slot, v: Value, is_set: bool) -> Result<()> {
        let ty = s.borrow().ty.clone();
        // `c = 5` without Set on a variable declared as an object type assigns the object's
        // default property (`Dim c As Range: ... c = c + 1`).
        if !is_set && matches!(ty, Ty::Object(_)) {
            let cur = s.borrow().value.clone();
            return match cur {
                Value::Object(Some(o)) => default_assign(&o, &Args::none(), v),
                _ => Err(RuntimeError::object_not_set()),
            };
        }
        let v = if is_set { v } else { ty.coerce(copy_array(v))? };
        if is_set && !matches!(v, Value::Object(_)) {
            return Err(RuntimeError::object_required());
        }
        s.borrow_mut().value = v;
        Ok(())
    }

    fn assign(&mut self, target: &Expr, mut v: Value, is_set: bool, f: &mut Frame) -> Result<()> {
        match target {
            Expr::Ident(name) if name.starts_with('[') => match self.bracket(name)? {
                Value::Object(Some(o)) => default_assign(&o, &Args::none(), v),
                _ => Err(RuntimeError::object_required()),
            },
            Expr::Ident(name) => {
                if let Some(s) = self.lookup_slot(name, f) {
                    return self.store(&s, v, is_set);
                }
                if let Some(inst) = f.instance.clone() {
                    let kind = if is_set { ProcKind::PropertySet } else { ProcKind::PropertyLet };
                    if let Some(pi) = self.modules[inst.module]
                        .procs
                        .iter()
                        .position(|p| p.name.eq_ignore_ascii_case(name) && p.kind == kind)
                    {
                        self.invoke(inst.module, pi, vec![slot(v, Ty::Variant)], Vec::new(), Some(inst))?;
                        return Ok(());
                    }
                }
                if let Some(r) = self.host.global_assign(name, &Args::none(), v.clone()) {
                    return r;
                }
                if let Some(Ok(Value::Object(Some(o)))) = self.host.global(name, &Args::none()) {
                    return default_assign(&o, &Args::none(), v);
                }
                if self.modules[f.module].options.explicit {
                    return Err(RuntimeError::new(1001, format!("Variable not defined: {name}")));
                }
                let s = slot(Value::Empty, Ty::Variant);
                f.locals.insert(key(name), s.clone());
                self.store(&s, v, is_set)
            }
            Expr::Call(callee, args) => {
                let a = self.eval_args(args, f)?;
                match callee.as_ref() {
                    Expr::Ident(name) => {
                        if let Some(s) = self.lookup_slot(name, f) {
                            let cur = s.borrow().value.clone();
                            return match cur {
                                Value::Array(arr) => {
                                    let idx = a.positional.iter().map(Value::to_i64).collect::<Result<Vec<_>>>()?;
                                    let elem = arr.borrow().elem_type.clone();
                                    let v = if is_set {
                                        v
                                    } else {
                                        Ty::from_name(elem.as_deref()).coerce(v.deref_default()?)?
                                    };
                                    arr.borrow_mut().set(&idx, v)
                                }
                                Value::Object(Some(o)) => default_assign(&o, &a, v),
                                Value::Object(None) => Err(RuntimeError::object_not_set()),
                                _ => Err(RuntimeError::type_mismatch()),
                            };
                        }
                        if matches!(key(name).as_str(), "mid" | "mid$") {
                            return self.mid_statement(args, &a, v, f);
                        }
                        if let Some(r) = self.host.global_assign(name, &a, v.clone()) {
                            return r;
                        }
                        match self.host.global(name, &a) {
                            Some(Ok(Value::Object(Some(o)))) => default_assign(&o, &Args::none(), v),
                            Some(Err(e)) => Err(e),
                            _ => Err(RuntimeError::new(35, format!("Sub or Function not defined: {name}"))),
                        }
                    }
                    Expr::Member(obj, name) => {
                        let target = match obj {
                            None => self.with_object(f)?,
                            Some(e) => self.eval(e, f)?,
                        };
                        let Value::Object(Some(o)) = target else {
                            return Err(if matches!(target, Value::Object(None)) {
                                RuntimeError::object_not_set()
                            } else {
                                RuntimeError::object_required()
                            });
                        };
                        // `ws.Cells(1, 2) = v`: the member returns an object whose default
                        // property takes the value. Otherwise it is a parameterised property.
                        match o.call(name, &a) {
                            Ok(Value::Object(Some(inner))) if !is_set => default_assign(&inner, &Args::none(), v),
                            Ok(Value::Array(arr)) if !is_set => {
                                let idx = a.positional.iter().map(Value::to_i64).collect::<Result<Vec<_>>>()?;
                                arr.borrow_mut().set(&idx, v)
                            }
                            _ => o.assign(name, &a, v),
                        }
                    }
                    other => {
                        let t = self.eval(other, f)?;
                        match t {
                            Value::Object(Some(o)) => default_assign(&o, &a, v),
                            _ => Err(RuntimeError::object_required()),
                        }
                    }
                }
            }
            Expr::Member(obj, name) => {
                // An application object passed through (see StmtKind::Let) is only kept for an
                // application object's property; everything else takes its value.
                let host_obj = |v: &Value| matches!(v, Value::Object(Some(o)) if o.as_any().is_none());
                let value_for_lang =
                    |v: Value| -> Result<Value> { if !is_set && host_obj(&v) { v.deref_default() } else { Ok(v) } };
                // Module.Var = v
                if let Some(Expr::Ident(q)) = obj.as_deref() {
                    if self.lookup_slot(q, f).is_none() {
                        if let Some(mi) = self.module_index(q) {
                            if let Some(s) = self.module_vars[mi].get(&key(name)).cloned() {
                                return self.store(&s, value_for_lang(v)?, is_set);
                            }
                        }
                    }
                }
                let target = match obj {
                    None => self.with_object(f)?,
                    Some(e) => self.eval(e, f)?,
                };
                match target {
                    Value::Object(Some(o)) => {
                        if o.as_any().is_some() {
                            v = value_for_lang(v)?;
                        }
                        if let Some(rec) = o.as_any().and_then(|a| a.downcast_ref::<Record>()) {
                            return rec.set_field(name, v);
                        }
                        if let Some(inst) = o.as_any().and_then(|a| a.downcast_ref::<ClassInstance>()) {
                            let module = self.modules[inst.module].clone();
                            let kind = if is_set { ProcKind::PropertySet } else { ProcKind::PropertyLet };
                            if let Some(pi) =
                                module.procs.iter().position(|p| p.name.eq_ignore_ascii_case(name) && p.kind == kind)
                            {
                                let inst_rc = self.class_rc(&o)?;
                                self.invoke(inst.module, pi, vec![slot(v, Ty::Variant)], Vec::new(), Some(inst_rc))?;
                                return Ok(());
                            }
                            let public = module
                                .vars
                                .iter()
                                .any(|(vis, spec)| *vis == Visibility::Public && spec.name.eq_ignore_ascii_case(name));
                            let s = inst.vars.borrow().get(&key(name)).cloned();
                            return match s {
                                Some(s) if public => self.store(&s, v, is_set),
                                _ => Err(RuntimeError::no_member(name)),
                            };
                        }
                        o.assign(name, &Args::none(), v)
                    }
                    Value::Object(None) => Err(RuntimeError::object_not_set()),
                    _ => Err(RuntimeError::object_required()),
                }
            }
            Expr::Bang(obj, name) => {
                let t = self.eval(obj, f)?;
                match t {
                    Value::Object(Some(o)) => o.assign("", &Args::positional(vec![Value::str(name.clone())]), v),
                    _ => Err(RuntimeError::object_required()),
                }
            }
            Expr::Paren(inner) => self.assign(inner, v, is_set, f),
            _ => Err(RuntimeError::new(1002, "Invalid assignment target")),
        }
    }
}

impl Interpreter<'_> {
    /// `Mid(var, start[, length]) = text`: overwrite characters in place.
    fn mid_statement(&mut self, args: &[Arg], a: &Args, v: Value, f: &mut Frame) -> Result<()> {
        let Some(Some(target)) = args.first().map(|x| x.value.as_ref()) else {
            return Err(RuntimeError::standard(449));
        };
        let current = self.eval_value(target, f)?.to_text()?;
        let start = a.require(1, "Start")?.to_i64()?;
        if start < 1 {
            return Err(RuntimeError::invalid_call());
        }
        let repl: Vec<char> = v.to_text()?.chars().collect();
        let mut chars: Vec<char> = current.chars().collect();
        let begin = (start - 1) as usize;
        let room = chars.len().saturating_sub(begin);
        let mut n = repl.len().min(room);
        if let Some(len) = a.get(2, "Length") {
            n = n.min(usize::try_from(len.to_i64()?.max(0)).unwrap_or(0));
        }
        chars[begin..begin + n].copy_from_slice(&repl[..n]);
        self.assign(target, Value::str(chars.into_iter().collect::<String>()), false, f)
    }
}

/// `obj = v` / `obj(args) = v`: the default property, falling back to `Value`.
fn default_assign(o: &ObjectRef, a: &Args, v: Value) -> Result<()> {
    match o.assign("", a, v.clone()) {
        Err(e) if e.number == 438 => o.assign("Value", a, v),
        r => r,
    }
}

/// Marker error number used to unwind on `End`.
const END_SIGNAL: i32 = -2;

impl Frame {
    fn new(module: usize, instance: Option<Rc<ClassInstance>>, proc_key: String) -> Self {
        Self {
            module,
            locals: HashMap::new(),
            with_stack: Vec::new(),
            on_error: OnError::Disable,
            in_handler: false,
            instance,
            proc_key,
            writebacks: Vec::new(),
        }
    }
}

fn find_label(stmts: &[Stmt], label: &str) -> Option<usize> {
    stmts.iter().position(|s| matches!(&s.kind, StmtKind::Label(l) if l == label))
}

fn reset_like(v: &Value) -> Value {
    match v {
        Value::Integer(_) => Value::Integer(0),
        Value::Long(_) => Value::Long(0),
        Value::Double(_) => Value::Double(0.0),
        Value::String(_) => Value::str(""),
        Value::Boolean(_) => Value::Boolean(false),
        Value::Object(_) => Value::Object(None),
        _ => Value::Empty,
    }
}

/// Text written by `Debug.Print` for one value. ORACLE: spacing rules for numbers.
fn print_text(v: &Value) -> Result<String> {
    Ok(match v {
        Value::Null => "Null".into(),
        Value::Error(n) => format!("Error {n}"),
        other => other.to_text()?,
    })
}

/// A value of a user-defined `Type`.
pub struct Record {
    name: String,
    fields: RefCell<Vec<(String, Value, Ty)>>,
}

impl Record {
    fn set_field(&self, name: &str, v: Value) -> Result<()> {
        let mut fields = self.fields.borrow_mut();
        let field = fields
            .iter_mut()
            .find(|(n, _, _)| n.eq_ignore_ascii_case(name))
            .ok_or_else(|| RuntimeError::no_member(name))?;
        field.1 = field.2.coerce(v)?;
        Ok(())
    }
}

impl Object for Record {
    fn type_name(&self) -> String {
        self.name.clone()
    }
    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        self.fields
            .borrow()
            .iter()
            .find(|(n, _, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v, _)| v.clone())
            .ok_or_else(|| RuntimeError::no_member(name))
    }
    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        self.set_field(name, value)
    }
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

/// Arrays are values: assignment copies them.
fn copy_array(v: Value) -> Value {
    match v {
        Value::Array(a) => {
            let mut copy = a.borrow().clone();
            copy.fixed = false;
            Value::Array(Rc::new(RefCell::new(copy)))
        }
        other => other,
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum StaticKind {
    Num,
    Str,
}

/// Element type name of a declared array (`Dim a() As Long` -> "Long"); `None` for Variant.
fn elem_name(spec: &VarSpec) -> Option<String> {
    spec.ty.as_ref().filter(|t| !t.eq_ignore_ascii_case("variant")).map(|t| {
        let mut c = t.chars();
        match c.next() {
            Some(first) => first.to_uppercase().collect::<String>() + &c.as_str().to_ascii_lowercase(),
            None => String::new(),
        }
    })
}

/// One step wider numeric type for Variant overflow promotion.
fn widen(v: &Value) -> Value {
    match v {
        Value::Byte(n) => Value::Integer(i16::from(*n)),
        Value::Integer(n) => Value::Long(i32::from(*n)),
        Value::Long(n) => Value::Double(f64::from(*n)),
        Value::Boolean(b) => Value::Long(if *b { -1 } else { 0 }),
        other => other.clone(),
    }
}

/// Library functions whose result has a fixed (non-Variant) type.
fn is_typed_builtin(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "cint" | "clng" | "clnglng" | "cdbl" | "csng" | "cbyte" | "ccur" | "cbool" | "cstr" | "len" | "asc" | "instr"
    )
}
