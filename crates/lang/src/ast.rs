//! Syntax tree for modules, procedures, statements and expressions.

use crate::ops::BinOp;
use crate::value::Value;

#[derive(Debug, Clone, Default)]
pub struct Options {
    pub explicit: bool,
    pub compare_text: bool,
    /// `Option Base 1`.
    pub base: i64,
    pub private_module: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModuleType {
    Standard,
    Class,
    Document,
}

#[derive(Debug, Clone)]
pub struct Module {
    pub name: String,
    pub kind: ModuleType,
    pub options: Options,
    pub vars: Vec<(Visibility, VarSpec)>,
    pub consts: Vec<(Visibility, String, Expr)>,
    pub enums: Vec<(String, Vec<(String, Option<Expr>)>)>,
    pub types: Vec<(String, Vec<VarSpec>)>,
    pub declares: Vec<String>,
    pub procs: Vec<Proc>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
}

#[derive(Debug, Clone)]
pub struct VarSpec {
    pub name: String,
    /// `None`: scalar. `Some(vec![])`: dynamic array `a()`. Otherwise (lower, upper) per dim.
    pub dims: Option<Vec<(Option<Expr>, Expr)>>,
    pub ty: Option<String>,
    /// `As New Collection`.
    pub new: bool,
    /// `As String * 10`.
    pub fixed_len: Option<Expr>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub by_val: bool,
    pub optional: bool,
    pub default: Option<Expr>,
    pub param_array: bool,
    pub ty: Option<String>,
    pub is_array: bool,
}

#[derive(Debug, Clone)]
pub struct Proc {
    pub name: String,
    pub kind: ProcKind,
    pub visibility: Visibility,
    pub is_static: bool,
    pub params: Vec<Param>,
    pub ret_ty: Option<String>,
    pub body: Vec<Stmt>,
    pub line: u32,
}

#[derive(Debug, Clone)]
pub enum CaseClause {
    Value(Expr),
    Range(Expr, Expr),
    Is(BinOp, Expr),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OnError {
    ResumeNext,
    GoTo(String),
    /// `On Error GoTo 0`.
    Disable,
    /// `On Error GoTo -1`.
    Reset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Resume {
    Same,
    Next,
    Label(String),
}

#[derive(Debug, Clone)]
pub struct Stmt {
    pub kind: StmtKind,
    pub line: u32,
}

#[derive(Debug, Clone)]
pub enum StmtKind {
    Dim {
        vars: Vec<VarSpec>,
        is_static: bool,
    },
    ReDim {
        preserve: bool,
        items: Vec<(Expr, Vec<(Option<Expr>, Expr)>)>,
    },
    Const(Vec<(String, Expr)>),
    Let {
        target: Expr,
        value: Expr,
    },
    Set {
        target: Expr,
        value: Expr,
    },
    /// A procedure or method call used as a statement.
    Call(Expr),
    If {
        arms: Vec<(Expr, Vec<Stmt>)>,
        otherwise: Option<Vec<Stmt>>,
    },
    Select {
        subject: Expr,
        cases: Vec<(Vec<CaseClause>, Vec<Stmt>)>,
        otherwise: Option<Vec<Stmt>>,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    ForEach {
        var: Expr,
        collection: Expr,
        body: Vec<Stmt>,
    },
    /// `Do [While|Until cond] ... Loop [While|Until cond]`; the bool is `until`.
    Do {
        pre: Option<(bool, Expr)>,
        post: Option<(bool, Expr)>,
        body: Vec<Stmt>,
    },
    While {
        cond: Expr,
        body: Vec<Stmt>,
    },
    With {
        object: Expr,
        body: Vec<Stmt>,
    },
    Exit(ExitKind),
    GoTo(String),
    GoSub(String),
    Return,
    Label(String),
    OnError(OnError),
    Resume(Resume),
    Erase(Vec<Expr>),
    /// `Debug.Print` items; `;`/`,` separators are kept as `None` markers for spacing.
    DebugPrint(Vec<Option<Expr>>),
    /// `End` (stop the whole program).
    End,
    Stop,
    /// A statement the engine deliberately does not run (file I/O, native calls...).
    Unsupported(String),
}

#[derive(Debug, Clone)]
pub struct Arg {
    pub name: Option<String>,
    /// `None` for a skipped argument (`f(1, , 3)`).
    pub value: Option<Expr>,
}

#[derive(Debug, Clone)]
pub enum Expr {
    Lit(Value),
    Ident(String),
    /// `obj.name`; `obj` is `None` for a leading-dot reference inside `With`.
    Member(Option<Box<Expr>>, String),
    /// `callee(args)`: a call, an array index, or a default-member call; decided at run time.
    Call(Box<Expr>, Vec<Arg>),
    /// `obj!name` = `obj("name")`.
    Bang(Box<Expr>, String),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    Neg(Box<Expr>),
    Not(Box<Expr>),
    /// A parenthesised expression: passed ByVal even to a ByRef parameter.
    Paren(Box<Expr>),
    /// `obj1 Is obj2`.
    Is(Box<Expr>, Box<Expr>),
    /// `TypeOf x Is Name`.
    TypeOf(Box<Expr>, String),
    New(String),
    Me,
}
