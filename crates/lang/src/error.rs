//! Runtime errors. Numbers follow the language's standard trappable errors so that
//! `Err.Number` and `On Error` handlers behave as macro authors expect.

#[derive(Debug, Clone, PartialEq)]
pub struct RuntimeError {
    pub number: i32,
    pub description: String,
}

impl RuntimeError {
    pub fn new(number: i32, description: impl Into<String>) -> Self {
        Self { number, description: description.into() }
    }
    pub fn standard(number: i32) -> Self {
        Self::new(number, standard_description(number))
    }
    pub fn invalid_call() -> Self {
        Self::standard(5)
    }
    pub fn overflow() -> Self {
        Self::standard(6)
    }
    pub fn subscript() -> Self {
        Self::standard(9)
    }
    pub fn div_zero() -> Self {
        Self::standard(11)
    }
    pub fn type_mismatch() -> Self {
        Self::standard(13)
    }
    pub fn object_not_set() -> Self {
        Self::standard(91)
    }
    pub fn object_required() -> Self {
        Self::standard(424)
    }
    pub fn no_member(name: &str) -> Self {
        let _ = name;
        Self::standard(438)
    }
    /// A limit set by the host (step budget, recursion depth, capability) was hit.
    pub fn limit(what: &str) -> Self {
        Self::new(-1, format!("macro stopped: {what}"))
    }
}

impl std::fmt::Display for RuntimeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "run-time error {}: {}", self.number, self.description)
    }
}

impl std::error::Error for RuntimeError {}

pub fn standard_description(number: i32) -> &'static str {
    match number {
        3 => "Return without GoSub",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        14 => "Out of string space",
        16 => "Expression too complex",
        17 => "Can't perform requested operation",
        18 => "User interrupt occurred",
        20 => "Resume without error",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        48 => "Error in loading DLL",
        51 => "Internal error",
        52 => "Bad file name or number",
        53 => "File not found",
        70 => "Permission denied",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        93 => "Invalid pattern string",
        94 => "Invalid use of Null",
        424 => "Object required",
        429 => "ActiveX component can't create object",
        438 => "Object doesn't support this property or method",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        1004 => "Application-defined or object-defined error",
        _ => "Application-defined or object-defined error",
    }
}

pub type Result<T> = std::result::Result<T, RuntimeError>;
