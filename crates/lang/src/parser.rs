//! Parser: tokens -> [`Module`]. Follows the statement and expression grammar of [MS-VBAL]
//! chapter 5, including the parts that make the language awkward to parse: calls without
//! parentheses, `Foo (x), y`, single-line `If`, labels, `With` leading dots, named arguments.

use crate::ast::*;
use crate::error::{Result, RuntimeError};
use crate::lexer::{Tok, Token, lex};
use crate::ops::BinOp;
use crate::value::Value;

pub fn parse_module(name: &str, kind: ModuleType, source: &str) -> Result<Module> {
    let toks = lex(&preprocess(source))?;
    let mut p = Parser { toks, pos: 0, pending_next: 0 };
    p.module(name, kind)
}

/// Statements outside any procedure (used by tests and the REPL-style `run_snippet`).
pub fn parse_statements(source: &str) -> Result<Vec<Stmt>> {
    let toks = lex(&preprocess(source))?;
    let mut p = Parser { toks, pos: 0, pending_next: 0 };
    let body = p.block(&|_| false)?;
    Ok(body)
}

/// Evaluates `#If` / `#ElseIf` / `#Else` / `#End If` / `#Const` lines. Excluded lines become
/// blank so line numbers stay right. Constants: VBA7, VBA6, Win64, Win32 true; Mac false.
pub fn preprocess(src: &str) -> String {
    let mut consts: Vec<(String, bool)> =
        ["vba7", "vba6", "win64", "win32"].iter().map(|s| (s.to_string(), true)).collect();
    consts.push(("mac".into(), false));
    // Stack of (this branch active, some branch already taken, parent active).
    let mut stack: Vec<(bool, bool, bool)> = Vec::new();
    let active = |stack: &Vec<(bool, bool, bool)>| stack.last().is_none_or(|s| s.0 && s.2);
    let mut out = String::with_capacity(src.len());
    for line in src.split('\n') {
        let t = line.trim_start();
        let lower = t.to_ascii_lowercase();
        let directive = lower.starts_with('#');
        if directive
            && (lower.starts_with("#if ")
                || lower.starts_with("#elseif ")
                || lower.starts_with("#else")
                || lower.starts_with("#end if")
                || lower.starts_with("#endif")
                || lower.starts_with("#const "))
        {
            if let Some(rest) = lower.strip_prefix("#if ") {
                let parent = active(&stack);
                let cond = eval_pp(rest.split(" then").next().unwrap_or(rest), &consts);
                stack.push((cond, cond, parent));
            } else if let Some(rest) = lower.strip_prefix("#elseif ") {
                if let Some(top) = stack.last_mut() {
                    let cond = !top.1 && eval_pp(rest.split(" then").next().unwrap_or(rest), &consts);
                    top.0 = cond;
                    top.1 |= cond;
                }
            } else if lower.starts_with("#else") {
                if let Some(top) = stack.last_mut() {
                    top.0 = !top.1;
                    top.1 = true;
                }
            } else if lower.starts_with("#end") {
                stack.pop();
            } else if let Some(rest) = lower.strip_prefix("#const ") {
                if active(&stack) {
                    if let Some((n, v)) = rest.split_once('=') {
                        let val = eval_pp(v, &consts);
                        consts.push((n.trim().to_owned(), val));
                    }
                }
            }
            out.push('\n');
            continue;
        }
        if active(&stack) {
            out.push_str(line);
        }
        out.push('\n');
    }
    out
}

fn eval_pp(expr: &str, consts: &[(String, bool)]) -> bool {
    let e = expr.trim().to_ascii_lowercase();
    if let Some((a, b)) = e.split_once(" or ") {
        return eval_pp(a, consts) || eval_pp(b, consts);
    }
    if let Some((a, b)) = e.split_once(" and ") {
        return eval_pp(a, consts) && eval_pp(b, consts);
    }
    if let Some(rest) = e.strip_prefix("not ") {
        return !eval_pp(rest, consts);
    }
    let e = e.trim_matches(|c| c == '(' || c == ')' || c == ' ');
    if let Some((a, b)) = e.split_once('=') {
        return eval_pp(a, consts) == eval_pp(b, consts);
    }
    match e {
        "true" | "-1" | "1" => true,
        "false" | "0" => false,
        name => consts.iter().rev().find(|(n, _)| n == name).is_some_and(|(_, v)| *v),
    }
}

struct Parser {
    toks: Vec<Token>,
    pos: usize,
    /// `Next i, j` closes more than one loop.
    pending_next: u32,
}

fn syntax(line: u32, msg: impl Into<String>) -> RuntimeError {
    RuntimeError::new(1002, format!("Syntax error on line {line}: {}", msg.into()))
}

const STATEMENT_END_WORDS: [&str; 1] = ["else"];

impl Parser {
    // ---- token helpers ----
    fn peek(&self) -> &Tok {
        self.toks.get(self.pos).map(|t| &t.tok).unwrap_or(&Tok::Newline)
    }
    fn peek_at(&self, k: usize) -> &Tok {
        self.toks.get(self.pos + k).map(|t| &t.tok).unwrap_or(&Tok::Newline)
    }
    fn line(&self) -> u32 {
        self.toks.get(self.pos).or_else(|| self.toks.last()).map(|t| t.line).unwrap_or(0)
    }
    fn at_end(&self) -> bool {
        self.pos >= self.toks.len()
    }
    fn bump(&mut self) -> Tok {
        let t = self.peek().clone();
        self.pos += 1;
        t
    }
    fn is_kw(&self, word: &str) -> bool {
        matches!(self.peek(), Tok::Ident(w) if w.eq_ignore_ascii_case(word))
    }
    fn is_kw_at(&self, k: usize, word: &str) -> bool {
        matches!(self.peek_at(k), Tok::Ident(w) if w.eq_ignore_ascii_case(word))
    }
    fn eat_kw(&mut self, word: &str) -> bool {
        if self.is_kw(word) {
            self.pos += 1;
            true
        } else {
            false
        }
    }
    fn expect_kw(&mut self, word: &str) -> Result<()> {
        if self.eat_kw(word) {
            Ok(())
        } else {
            Err(syntax(self.line(), format!("expected {word}, found {:?}", self.peek())))
        }
    }
    fn is_sym(&self, s: &str) -> bool {
        matches!(self.peek(), Tok::Sym(x) if *x == s)
    }
    fn eat_sym(&mut self, s: &str) -> bool {
        if self.is_sym(s) {
            self.pos += 1;
            true
        } else {
            false
        }
    }
    fn expect_sym(&mut self, s: &str) -> Result<()> {
        if self.eat_sym(s) {
            Ok(())
        } else {
            Err(syntax(self.line(), format!("expected '{s}', found {:?}", self.peek())))
        }
    }
    fn at_stmt_end(&self) -> bool {
        matches!(self.peek(), Tok::Newline | Tok::Sym(":"))
            || self.at_end()
            || STATEMENT_END_WORDS.iter().any(|w| self.is_kw(w))
    }
    fn skip_to_line_end(&mut self) {
        while !self.at_end() && !matches!(self.peek(), Tok::Newline) {
            self.pos += 1;
        }
    }
    fn skip_separators(&mut self) {
        while matches!(self.peek(), Tok::Newline | Tok::Sym(":")) && !self.at_end() {
            self.pos += 1;
        }
    }
    fn ident(&mut self) -> Result<String> {
        match self.bump() {
            Tok::Ident(w) | Tok::TypedIdent(w, _) => Ok(w),
            t => Err(syntax(self.line(), format!("expected a name, found {t:?}"))),
        }
    }
    /// A name after `.`: keywords are allowed (`.End`, `.Select`, `.Print`).
    fn member_name(&mut self) -> Result<String> {
        self.ident()
    }
    /// Type name after `As`: `Long`, `the spreadsheet host.Range`, `String * 10` handled by caller.
    fn type_name(&mut self) -> Result<String> {
        let mut name = self.ident()?;
        while self.eat_sym(".") {
            name = self.ident()?;
        }
        Ok(name)
    }

    // ---- module ----
    fn module(&mut self, name: &str, kind: ModuleType) -> Result<Module> {
        let mut m = Module {
            name: name.to_owned(),
            kind,
            options: Options::default(),
            vars: Vec::new(),
            consts: Vec::new(),
            enums: Vec::new(),
            types: Vec::new(),
            declares: Vec::new(),
            procs: Vec::new(),
        };
        loop {
            self.skip_separators();
            if self.at_end() {
                break;
            }
            let line = self.line();
            if self.eat_kw("Attribute")
                || self.is_kw("Implements")
                || self.is_kw("Event")
                || self.peek_word_starts_with("Def")
            {
                self.skip_to_line_end();
                continue;
            }
            if self.eat_kw("Option") {
                if self.eat_kw("Explicit") {
                    m.options.explicit = true;
                } else if self.eat_kw("Base") {
                    if let Tok::Integer(n) = self.bump() {
                        m.options.base = n;
                    }
                } else if self.eat_kw("Compare") {
                    m.options.compare_text = self.eat_kw("Text");
                    self.skip_to_line_end();
                } else if self.eat_kw("Private") {
                    self.eat_kw("Module");
                    m.options.private_module = true;
                } else {
                    self.skip_to_line_end();
                }
                continue;
            }
            let mut vis = Visibility::Public;
            let mut is_static = false;
            let mut saw_vis = false;
            loop {
                if self.eat_kw("Public") || self.eat_kw("Global") {
                    vis = Visibility::Public;
                    saw_vis = true;
                } else if self.eat_kw("Private") {
                    vis = Visibility::Private;
                    saw_vis = true;
                } else if self.eat_kw("Friend") {
                    saw_vis = true;
                } else if self.is_kw("Static")
                    && (self.is_kw_at(1, "Sub") || self.is_kw_at(1, "Function") || self.is_kw_at(1, "Property"))
                {
                    self.pos += 1;
                    is_static = true;
                } else {
                    break;
                }
            }
            if self.is_kw("Sub") || self.is_kw("Function") || self.is_kw("Property") {
                let mut proc = self.procedure(vis, is_static)?;
                proc.line = line;
                m.procs.push(proc);
                continue;
            }
            if self.eat_kw("Const") {
                for (n, e) in self.const_list()? {
                    m.consts.push((vis, n, e));
                }
                continue;
            }
            if self.eat_kw("Enum") {
                let ename = self.ident()?;
                let mut members = Vec::new();
                loop {
                    self.skip_separators();
                    if self.is_kw("End") && self.is_kw_at(1, "Enum") {
                        self.pos += 2;
                        break;
                    }
                    if self.at_end() {
                        return Err(syntax(line, "Enum without End Enum"));
                    }
                    let n = self.ident()?;
                    let v = if self.eat_sym("=") { Some(self.expr()?) } else { None };
                    members.push((n, v));
                }
                m.enums.push((ename, members));
                continue;
            }
            if self.eat_kw("Type") {
                let tname = self.ident()?;
                let mut fields = Vec::new();
                loop {
                    self.skip_separators();
                    if self.is_kw("End") && self.is_kw_at(1, "Type") {
                        self.pos += 2;
                        break;
                    }
                    if self.at_end() {
                        return Err(syntax(line, "Type without End Type"));
                    }
                    fields.push(self.var_spec()?);
                }
                m.types.push((tname, fields));
                continue;
            }
            if self.eat_kw("Declare") {
                self.eat_kw("PtrSafe");
                self.eat_kw("Sub");
                self.eat_kw("Function");
                if let Ok(n) = self.ident() {
                    m.declares.push(n);
                }
                self.skip_to_line_end();
                continue;
            }
            if self.eat_kw("Dim") || saw_vis {
                self.eat_kw("WithEvents");
                loop {
                    let spec = self.var_spec()?;
                    m.vars.push((vis, spec));
                    if !self.eat_sym(",") {
                        break;
                    }
                    self.eat_kw("WithEvents");
                }
                continue;
            }
            return Err(syntax(line, format!("unexpected {:?} at module level", self.peek())));
        }
        Ok(m)
    }

    fn peek_word_starts_with(&self, prefix: &str) -> bool {
        matches!(self.peek(), Tok::Ident(w) if w.len() == 6 && w.starts_with(prefix) && w.chars().all(|c| c.is_ascii_alphabetic()))
    }

    fn const_list(&mut self) -> Result<Vec<(String, Expr)>> {
        let mut out = Vec::new();
        loop {
            let name = self.ident()?;
            if self.eat_kw("As") {
                self.type_name()?;
            }
            self.expect_sym("=")?;
            out.push((name, self.expr()?));
            if !self.eat_sym(",") {
                break;
            }
        }
        Ok(out)
    }

    fn var_spec(&mut self) -> Result<VarSpec> {
        let (name, tchar) = match self.bump() {
            Tok::Ident(w) => (w, None),
            Tok::TypedIdent(w, c) => (w, Some(c)),
            t => return Err(syntax(self.line(), format!("expected a variable name, found {t:?}"))),
        };
        let mut spec = VarSpec { name, dims: None, ty: tchar.map(type_for_char), new: false, fixed_len: None };
        if self.eat_sym("(") {
            spec.dims = Some(self.dim_bounds()?);
        }
        if self.eat_kw("As") {
            spec.new = self.eat_kw("New");
            spec.ty = Some(self.type_name()?);
            if self.eat_sym("*") {
                spec.fixed_len = Some(self.primary()?);
            }
        }
        Ok(spec)
    }

    /// Bounds inside `(...)`, the opening paren already consumed.
    fn dim_bounds(&mut self) -> Result<Vec<(Option<Expr>, Expr)>> {
        let mut dims = Vec::new();
        if self.eat_sym(")") {
            return Ok(dims);
        }
        loop {
            let first = self.expr()?;
            if self.eat_kw("To") {
                let upper = self.expr()?;
                dims.push((Some(first), upper));
            } else {
                dims.push((None, first));
            }
            if !self.eat_sym(",") {
                break;
            }
        }
        self.expect_sym(")")?;
        Ok(dims)
    }

    fn procedure(&mut self, visibility: Visibility, is_static: bool) -> Result<Proc> {
        let line = self.line();
        let kind = if self.eat_kw("Sub") {
            ProcKind::Sub
        } else if self.eat_kw("Function") {
            ProcKind::Function
        } else {
            self.expect_kw("Property")?;
            if self.eat_kw("Get") {
                ProcKind::PropertyGet
            } else if self.eat_kw("Let") {
                ProcKind::PropertyLet
            } else {
                self.expect_kw("Set")?;
                ProcKind::PropertySet
            }
        };
        let (name, tchar) = match self.bump() {
            Tok::Ident(w) => (w, None),
            Tok::TypedIdent(w, c) => (w, Some(c)),
            t => return Err(syntax(line, format!("expected a procedure name, found {t:?}"))),
        };
        let mut params = Vec::new();
        if self.eat_sym("(") {
            while !self.eat_sym(")") {
                params.push(self.param()?);
                if !self.eat_sym(",") {
                    self.expect_sym(")")?;
                    break;
                }
            }
        }
        let mut ret_ty = tchar.map(type_for_char);
        if self.eat_kw("As") {
            ret_ty = Some(self.type_name()?);
            if self.eat_sym("(") {
                self.expect_sym(")")?;
            }
        }
        let end_word = match kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            _ => "Property",
        };
        let body = self.block(&|p: &Parser| p.is_kw("End") && p.is_kw_at(1, end_word))?;
        if !(self.eat_kw("End") && self.eat_kw(end_word)) {
            return Err(syntax(line, format!("{name}: missing End {end_word}")));
        }
        Ok(Proc { name, kind, visibility, is_static, params, ret_ty, body, line })
    }

    fn param(&mut self) -> Result<Param> {
        let mut p = Param {
            name: String::new(),
            by_val: false,
            optional: false,
            default: None,
            param_array: false,
            ty: None,
            is_array: false,
        };
        loop {
            if self.eat_kw("Optional") {
                p.optional = true;
            } else if self.eat_kw("ByVal") {
                p.by_val = true;
            } else if self.eat_kw("ByRef") {
                p.by_val = false;
            } else if self.eat_kw("ParamArray") {
                p.param_array = true;
            } else {
                break;
            }
        }
        let (name, tchar) = match self.bump() {
            Tok::Ident(w) => (w, None),
            Tok::TypedIdent(w, c) => (w, Some(c)),
            t => return Err(syntax(self.line(), format!("expected a parameter name, found {t:?}"))),
        };
        p.name = name;
        p.ty = tchar.map(type_for_char);
        if self.eat_sym("(") {
            self.expect_sym(")")?;
            p.is_array = true;
        }
        if self.eat_kw("As") {
            p.ty = Some(self.type_name()?);
        }
        if self.eat_sym("=") {
            p.default = Some(self.expr()?);
        }
        Ok(p)
    }

    // ---- statements ----
    fn block(&mut self, stop: &dyn Fn(&Parser) -> bool) -> Result<Vec<Stmt>> {
        let mut out = Vec::new();
        loop {
            self.skip_separators();
            if self.at_end() || self.pending_next > 0 || stop(self) {
                return Ok(out);
            }
            if let Some(s) = self.statement()? {
                out.push(s);
            }
        }
    }

    fn is_block_end(&self) -> bool {
        self.is_kw("Else")
            || self.is_kw("ElseIf")
            || self.is_kw("Next")
            || self.is_kw("Loop")
            || self.is_kw("Wend")
            || self.is_kw("Case")
            || (self.is_kw("End")
                && ["If", "Select", "With", "Sub", "Function", "Property"].iter().any(|w| self.is_kw_at(1, w)))
    }

    fn statement(&mut self) -> Result<Option<Stmt>> {
        let line = self.line();
        let line_start = self.pos == 0 || matches!(self.toks.get(self.pos - 1).map(|t| &t.tok), Some(Tok::Newline));
        // Labels: `Name:` or a line number at the start of a line.
        if line_start {
            if let (Tok::Ident(w), Tok::Sym(":")) = (self.peek().clone(), self.peek_at(1).clone()) {
                if !is_reserved(&w) {
                    self.pos += 2;
                    return Ok(Some(Stmt { kind: StmtKind::Label(w.to_ascii_lowercase()), line }));
                }
            }
            if let Tok::Integer(n) = self.peek().clone() {
                self.pos += 1;
                self.eat_sym(":");
                return Ok(Some(Stmt { kind: StmtKind::Label(n.to_string()), line }));
            }
        }
        if self.is_block_end() {
            return Err(syntax(line, format!("unexpected {:?}", self.peek())));
        }
        let kind = self.statement_kind()?;
        Ok(kind.map(|kind| Stmt { kind, line }))
    }

    fn statement_kind(&mut self) -> Result<Option<StmtKind>> {
        let line = self.line();
        let word = match self.peek() {
            Tok::Ident(w) => w.to_ascii_lowercase(),
            _ => String::new(),
        };
        match word.as_str() {
            "dim" | "static" => {
                let is_static = word == "static";
                self.pos += 1;
                let mut vars = Vec::new();
                loop {
                    self.eat_kw("WithEvents");
                    vars.push(self.var_spec()?);
                    if !self.eat_sym(",") {
                        break;
                    }
                }
                Ok(Some(StmtKind::Dim { vars, is_static }))
            }
            "redim" => {
                self.pos += 1;
                let preserve = self.eat_kw("Preserve");
                let mut items = Vec::new();
                loop {
                    let mut target = Expr::Ident(self.ident()?);
                    while self.eat_sym(".") {
                        target = Expr::Member(Some(Box::new(target)), self.member_name()?);
                    }
                    self.expect_sym("(")?;
                    let dims = self.dim_bounds()?;
                    if self.eat_kw("As") {
                        self.type_name()?;
                    }
                    items.push((target, dims));
                    if !self.eat_sym(",") {
                        break;
                    }
                }
                Ok(Some(StmtKind::ReDim { preserve, items }))
            }
            "const" => {
                self.pos += 1;
                Ok(Some(StmtKind::Const(self.const_list()?)))
            }
            "set" => {
                self.pos += 1;
                let target = self.postfix_expr()?;
                self.expect_sym("=")?;
                let value = self.expr()?;
                Ok(Some(StmtKind::Set { target, value }))
            }
            "let" => {
                self.pos += 1;
                let target = self.postfix_expr()?;
                self.expect_sym("=")?;
                Ok(Some(StmtKind::Let { target, value: self.expr()? }))
            }
            "call" => {
                self.pos += 1;
                Ok(Some(StmtKind::Call(self.postfix_expr()?)))
            }
            "if" => self.if_statement().map(Some),
            "select" => self.select_statement().map(Some),
            "for" => self.for_statement().map(Some),
            "do" => self.do_statement().map(Some),
            "while" => {
                self.pos += 1;
                let cond = self.expr()?;
                let body = self.block(&|p: &Parser| p.is_kw("Wend"))?;
                self.expect_kw("Wend")?;
                Ok(Some(StmtKind::While { cond, body }))
            }
            "with" => {
                self.pos += 1;
                let object = self.expr()?;
                let body = self.block(&|p: &Parser| p.is_kw("End") && p.is_kw_at(1, "With"))?;
                self.expect_kw("End")?;
                self.expect_kw("With")?;
                Ok(Some(StmtKind::With { object, body }))
            }
            "exit" => {
                self.pos += 1;
                let k = match self.ident()?.to_ascii_lowercase().as_str() {
                    "sub" => ExitKind::Sub,
                    "function" => ExitKind::Function,
                    "property" => ExitKind::Property,
                    "for" => ExitKind::For,
                    "do" => ExitKind::Do,
                    other => return Err(syntax(line, format!("Exit {other}"))),
                };
                Ok(Some(StmtKind::Exit(k)))
            }
            "goto" => {
                self.pos += 1;
                Ok(Some(StmtKind::GoTo(self.label_name()?)))
            }
            "gosub" => {
                self.pos += 1;
                Ok(Some(StmtKind::GoSub(self.label_name()?)))
            }
            "return" => {
                self.pos += 1;
                Ok(Some(StmtKind::Return))
            }
            "on" if self.is_kw_at(1, "Error") => {
                self.pos += 2;
                if self.eat_kw("Resume") {
                    self.expect_kw("Next")?;
                    return Ok(Some(StmtKind::OnError(OnError::ResumeNext)));
                }
                self.expect_kw("GoTo")?;
                let kind = match self.peek().clone() {
                    Tok::Integer(0) => {
                        self.pos += 1;
                        OnError::Disable
                    }
                    Tok::Sym("-") => {
                        self.pos += 2;
                        OnError::Reset
                    }
                    _ => OnError::GoTo(self.label_name()?),
                };
                Ok(Some(StmtKind::OnError(kind)))
            }
            "resume" => {
                self.pos += 1;
                let r = if self.eat_kw("Next") {
                    Resume::Next
                } else if self.at_stmt_end() || matches!(self.peek(), Tok::Integer(0)) {
                    if matches!(self.peek(), Tok::Integer(0)) {
                        self.pos += 1;
                    }
                    Resume::Same
                } else {
                    Resume::Label(self.label_name()?)
                };
                Ok(Some(StmtKind::Resume(r)))
            }
            "erase" => {
                self.pos += 1;
                let mut items = vec![self.postfix_expr()?];
                while self.eat_sym(",") {
                    items.push(self.postfix_expr()?);
                }
                Ok(Some(StmtKind::Erase(items)))
            }
            "end" => {
                self.pos += 1;
                Ok(Some(StmtKind::End))
            }
            "stop" => {
                self.pos += 1;
                Ok(Some(StmtKind::Stop))
            }
            "option" | "attribute" => {
                self.skip_to_line_end();
                Ok(None)
            }
            // Kill, MkDir, SendKeys, Reset and friends parse as ordinary calls: a project may
            // define procedures with those names; the interpreter refuses the built-in ones.
            "open" | "close" | "get" | "put" | "input" | "print" | "write" | "seek" | "width" | "line"
                if self.statement_has_hash() =>
            {
                self.skip_statement();
                Ok(Some(StmtKind::Unsupported(format!("{word} # (file I/O)"))))
            }
            _ => self.expression_statement().map(Some),
        }
    }

    fn statement_has_hash(&self) -> bool {
        let mut k = 0;
        loop {
            match self.peek_at(k) {
                Tok::Newline => return false,
                Tok::Sym("#") => return true,
                Tok::Sym(":") => return false,
                _ => k += 1,
            }
            if self.pos + k >= self.toks.len() {
                return false;
            }
        }
    }

    fn skip_statement(&mut self) {
        while !self.at_end() && !matches!(self.peek(), Tok::Newline | Tok::Sym(":")) {
            self.pos += 1;
        }
    }

    fn label_name(&mut self) -> Result<String> {
        match self.bump() {
            Tok::Ident(w) => Ok(w.to_ascii_lowercase()),
            Tok::Integer(n) => Ok(n.to_string()),
            t => Err(syntax(self.line(), format!("expected a label, found {t:?}"))),
        }
    }

    fn expression_statement(&mut self) -> Result<StmtKind> {
        let line = self.line();
        let target = self.postfix_expr()?;
        // Debug.Print a; b, c
        if let Expr::Member(Some(obj), name) = &target {
            if name.eq_ignore_ascii_case("print")
                && matches!(obj.as_ref(), Expr::Ident(d) if d.eq_ignore_ascii_case("debug"))
            {
                let mut items = Vec::new();
                while !self.at_stmt_end() {
                    if self.eat_sym(";") || self.eat_sym(",") {
                        items.push(None);
                        continue;
                    }
                    items.push(Some(self.expr()?));
                }
                return Ok(StmtKind::DebugPrint(items));
            }
        }
        if self.eat_sym("=") {
            let value = self.expr()?;
            return Ok(StmtKind::Let { target, value });
        }
        if self.at_stmt_end() {
            // `Foo(x)` as a statement: the parentheses belong to the argument, so it is passed
            // by value (only `Call Foo(x)` passes it by reference).
            if let Expr::Call(callee, args) = &target {
                if args.len() == 1 && args[0].name.is_none() {
                    if let Some(v) = &args[0].value {
                        if !matches!(v, Expr::Paren(_)) {
                            let arg = Arg { name: None, value: Some(Expr::Paren(Box::new(v.clone()))) };
                            return Ok(StmtKind::Call(Expr::Call(callee.clone(), vec![arg])));
                        }
                    }
                }
            }
            return Ok(StmtKind::Call(target));
        }
        // `Foo (a) & b, c` / `Foo (a), b`: what parsed as Foo(a) was really a parenthesised
        // first argument.
        if let Expr::Call(callee, args) = &target {
            if args.len() == 1 && args[0].name.is_none() && !self.is_sym("(") && !self.is_sym(".") {
                if let Some(first) = &args[0].value {
                    let first = self.binary_rest(Expr::Paren(Box::new(first.clone())), 0)?;
                    let mut all = vec![Arg { name: None, value: Some(first) }];
                    while self.eat_sym(",") {
                        all.push(self.arg()?);
                    }
                    if !self.at_stmt_end() {
                        return Err(syntax(line, format!("unexpected {:?}", self.peek())));
                    }
                    return Ok(StmtKind::Call(Expr::Call(callee.clone(), all)));
                }
            }
        }
        // Call without parentheses: `MsgBox "x", vbOKOnly` / `rng.Copy Destination:=x`.
        let mut args = vec![self.arg()?];
        while self.eat_sym(",") {
            args.push(self.arg()?);
        }
        if !self.at_stmt_end() {
            return Err(syntax(line, format!("unexpected {:?}", self.peek())));
        }
        Ok(StmtKind::Call(Expr::Call(Box::new(target), args)))
    }

    fn if_statement(&mut self) -> Result<StmtKind> {
        self.expect_kw("If")?;
        let cond = self.expr()?;
        self.expect_kw("Then")?;
        if !matches!(self.peek(), Tok::Newline) && !self.at_end() {
            // Single-line If: statements up to Else / end of line.
            let then = self.single_line_stmts()?;
            let otherwise = if self.eat_kw("Else") { Some(self.single_line_stmts()?) } else { None };
            return Ok(StmtKind::If { arms: vec![(cond, then)], otherwise });
        }
        let mut arms = Vec::new();
        let body =
            self.block(&|p: &Parser| p.is_kw("ElseIf") || p.is_kw("Else") || (p.is_kw("End") && p.is_kw_at(1, "If")))?;
        arms.push((cond, body));
        let mut otherwise = None;
        loop {
            if self.eat_kw("ElseIf") {
                let c = self.expr()?;
                self.expect_kw("Then")?;
                let b = self.block(&|p: &Parser| {
                    p.is_kw("ElseIf") || p.is_kw("Else") || (p.is_kw("End") && p.is_kw_at(1, "If"))
                })?;
                arms.push((c, b));
            } else if self.eat_kw("Else") {
                otherwise = Some(self.block(&|p: &Parser| p.is_kw("End") && p.is_kw_at(1, "If"))?);
            } else {
                self.expect_kw("End")?;
                self.expect_kw("If")?;
                break;
            }
        }
        Ok(StmtKind::If { arms, otherwise })
    }

    fn single_line_stmts(&mut self) -> Result<Vec<Stmt>> {
        let mut out = Vec::new();
        loop {
            let line = self.line();
            if matches!(self.peek(), Tok::Newline) || self.is_kw("Else") || self.at_end() {
                break;
            }
            if self.eat_sym(":") {
                continue;
            }
            // A bare label/number after Then is an implicit GoTo.
            if let Tok::Integer(n) = self.peek().clone() {
                self.pos += 1;
                out.push(Stmt { kind: StmtKind::GoTo(n.to_string()), line });
                continue;
            }
            if let Some(kind) = self.statement_kind()? {
                out.push(Stmt { kind, line });
            }
        }
        Ok(out)
    }

    fn select_statement(&mut self) -> Result<StmtKind> {
        self.expect_kw("Select")?;
        self.expect_kw("Case")?;
        let subject = self.expr()?;
        let mut cases = Vec::new();
        let mut otherwise = None;
        self.skip_separators();
        loop {
            self.skip_separators();
            if self.is_kw("End") && self.is_kw_at(1, "Select") {
                self.pos += 2;
                break;
            }
            self.expect_kw("Case")?;
            let stop = |p: &Parser| p.is_kw("Case") || (p.is_kw("End") && p.is_kw_at(1, "Select"));
            if self.eat_kw("Else") {
                otherwise = Some(self.block(&stop)?);
                continue;
            }
            let mut clauses = Vec::new();
            loop {
                if self.eat_kw("Is") {
                    let op = match self.bump() {
                        Tok::Sym("=") => BinOp::Eq,
                        Tok::Sym("<>") => BinOp::Ne,
                        Tok::Sym("<") => BinOp::Lt,
                        Tok::Sym("<=") => BinOp::Le,
                        Tok::Sym(">") => BinOp::Gt,
                        Tok::Sym(">=") => BinOp::Ge,
                        t => return Err(syntax(self.line(), format!("Case Is {t:?}"))),
                    };
                    clauses.push(CaseClause::Is(op, self.expr()?));
                } else {
                    let a = self.expr()?;
                    if self.eat_kw("To") {
                        clauses.push(CaseClause::Range(a, self.expr()?));
                    } else {
                        clauses.push(CaseClause::Value(a));
                    }
                }
                if !self.eat_sym(",") {
                    break;
                }
            }
            let body = self.block(&stop)?;
            cases.push((clauses, body));
        }
        Ok(StmtKind::Select { subject, cases, otherwise })
    }

    fn for_statement(&mut self) -> Result<StmtKind> {
        self.expect_kw("For")?;
        if self.eat_kw("Each") {
            let var = self.postfix_expr()?;
            self.expect_kw("In")?;
            let collection = self.expr()?;
            let body = self.block(&|p: &Parser| p.is_kw("Next"))?;
            self.close_next()?;
            return Ok(StmtKind::ForEach { var, collection, body });
        }
        let var = self.postfix_expr()?;
        self.expect_sym("=")?;
        let from = self.expr()?;
        self.expect_kw("To")?;
        let to = self.expr()?;
        let step = if self.eat_kw("Step") { Some(self.expr()?) } else { None };
        let body = self.block(&|p: &Parser| p.is_kw("Next"))?;
        self.close_next()?;
        Ok(StmtKind::For { var, from, to, step, body })
    }

    /// Consumes `Next [a[, b...]]`, or one pending close from an earlier `Next a, b`.
    fn close_next(&mut self) -> Result<()> {
        if self.pending_next > 0 {
            self.pending_next -= 1;
            return Ok(());
        }
        self.expect_kw("Next")?;
        if !self.at_stmt_end() {
            self.postfix_expr()?;
            while self.eat_sym(",") {
                self.postfix_expr()?;
                self.pending_next += 1;
            }
        }
        Ok(())
    }

    fn do_statement(&mut self) -> Result<StmtKind> {
        self.expect_kw("Do")?;
        let pre = if self.eat_kw("While") {
            Some((false, self.expr()?))
        } else if self.eat_kw("Until") {
            Some((true, self.expr()?))
        } else {
            None
        };
        let body = self.block(&|p: &Parser| p.is_kw("Loop"))?;
        self.expect_kw("Loop")?;
        let post = if self.eat_kw("While") {
            Some((false, self.expr()?))
        } else if self.eat_kw("Until") {
            Some((true, self.expr()?))
        } else {
            None
        };
        Ok(StmtKind::Do { pre, post, body })
    }

    // ---- expressions ----
    pub fn expr(&mut self) -> Result<Expr> {
        self.binary(0)
    }

    fn binop(&self) -> Option<(BinOp, u8)> {
        let op = match self.peek() {
            Tok::Sym("^") => (BinOp::Pow, 14),
            Tok::Sym("*") => (BinOp::Mul, 12),
            Tok::Sym("/") => (BinOp::Div, 12),
            Tok::Sym("\\") => (BinOp::IntDiv, 11),
            Tok::Ident(w) if w.eq_ignore_ascii_case("mod") => (BinOp::Mod, 10),
            Tok::Sym("+") => (BinOp::Add, 9),
            Tok::Sym("-") => (BinOp::Sub, 9),
            Tok::Sym("&") => (BinOp::Concat, 8),
            Tok::Sym("=") => (BinOp::Eq, 7),
            Tok::Sym("<>") => (BinOp::Ne, 7),
            Tok::Sym("<") => (BinOp::Lt, 7),
            Tok::Sym("<=") => (BinOp::Le, 7),
            Tok::Sym(">") => (BinOp::Gt, 7),
            Tok::Sym(">=") => (BinOp::Ge, 7),
            Tok::Ident(w) if w.eq_ignore_ascii_case("like") => (BinOp::Like, 7),
            // `Is` is handled as its own node; give it comparison precedence.
            Tok::Ident(w) if w.eq_ignore_ascii_case("is") => (BinOp::Eq, 7),
            Tok::Ident(w) if w.eq_ignore_ascii_case("and") => (BinOp::And, 5),
            Tok::Ident(w) if w.eq_ignore_ascii_case("or") => (BinOp::Or, 4),
            Tok::Ident(w) if w.eq_ignore_ascii_case("xor") => (BinOp::Xor, 3),
            Tok::Ident(w) if w.eq_ignore_ascii_case("eqv") => (BinOp::Eqv, 2),
            Tok::Ident(w) if w.eq_ignore_ascii_case("imp") => (BinOp::Imp, 1),
            _ => return None,
        };
        Some(op)
    }

    fn binary(&mut self, min: u8) -> Result<Expr> {
        let lhs = self.unary()?;
        self.binary_rest(lhs, min)
    }

    fn binary_rest(&mut self, mut lhs: Expr, min: u8) -> Result<Expr> {
        while let Some((op, prec)) = self.binop() {
            if prec < min {
                break;
            }
            let is_is = self.is_kw("Is");
            self.pos += 1;
            let rhs = self.binary(prec + 1)?;
            lhs = if is_is {
                Expr::Is(Box::new(lhs), Box::new(rhs))
            } else {
                Expr::Binary(op, Box::new(lhs), Box::new(rhs))
            };
        }
        Ok(lhs)
    }

    fn unary(&mut self) -> Result<Expr> {
        if self.eat_kw("Not") {
            // Not binds looser than comparison: `Not a = b` is `Not (a = b)`.
            return Ok(Expr::Not(Box::new(self.binary(6)?)));
        }
        if self.eat_sym("-") {
            return Ok(Expr::Neg(Box::new(self.binary(13)?)));
        }
        if self.eat_sym("+") {
            return self.binary(13);
        }
        self.postfix_expr()
    }

    fn postfix_expr(&mut self) -> Result<Expr> {
        let mut e = self.primary()?;
        loop {
            if self.is_sym("(") {
                self.pos += 1;
                let args = self.paren_args()?;
                e = Expr::Call(Box::new(e), args);
            } else if self.is_sym(".") {
                self.pos += 1;
                let name = self.member_name()?;
                e = Expr::Member(Some(Box::new(e)), name);
            } else if self.is_sym("!") {
                self.pos += 1;
                let name = self.member_name()?;
                e = Expr::Bang(Box::new(e), name);
            } else {
                return Ok(e);
            }
        }
    }

    fn paren_args(&mut self) -> Result<Vec<Arg>> {
        let mut args = Vec::new();
        if self.eat_sym(")") {
            return Ok(args);
        }
        loop {
            args.push(self.arg()?);
            if self.eat_sym(")") {
                return Ok(args);
            }
            self.expect_sym(",")?;
        }
    }

    fn arg(&mut self) -> Result<Arg> {
        if self.is_sym(",") || self.is_sym(")") || self.at_stmt_end() {
            return Ok(Arg { name: None, value: None });
        }
        if let (Tok::Ident(n) | Tok::TypedIdent(n, _), Tok::Sym(":=")) = (self.peek().clone(), self.peek_at(1).clone())
        {
            self.pos += 2;
            return Ok(Arg { name: Some(n), value: Some(self.expr()?) });
        }
        self.eat_kw("ByVal");
        Ok(Arg { name: None, value: Some(self.expr()?) })
    }

    fn primary(&mut self) -> Result<Expr> {
        let line = self.line();
        let t = self.bump();
        Ok(match t {
            Tok::Integer(n) => Expr::Lit(int_literal(n)),
            Tok::TypedInteger(n, c) => Expr::Lit(match c {
                '%' => Value::Integer(i16::try_from(n).map_err(|_| syntax(line, "Integer literal out of range"))?),
                '&' => Value::Long(i32::try_from(n).map_err(|_| syntax(line, "Long literal out of range"))?),
                '^' => Value::LongLong(n),
                '!' => Value::Single(n as f32),
                '#' => Value::Double(n as f64),
                '@' => Value::Currency(n.saturating_mul(10_000)),
                _ => int_literal(n),
            }),
            Tok::Float(f) => Expr::Lit(Value::Double(f)),
            Tok::TypedFloat(f, c) => Expr::Lit(match c {
                '!' => Value::Single(f as f32),
                '@' => Value::Currency((f * 10_000.0).round() as i64),
                _ => Value::Double(f),
            }),
            Tok::Str(s) => Expr::Lit(Value::str(s)),
            Tok::Date(d) => Expr::Lit(Value::Date(d)),
            Tok::Sym("(") => {
                let e = self.expr()?;
                self.expect_sym(")")?;
                Expr::Paren(Box::new(e))
            }
            Tok::Sym(".") => Expr::Member(None, self.member_name()?),
            Tok::Sym("!") => Expr::Bang(Box::new(Expr::Member(None, String::new())), self.member_name()?),
            Tok::Ident(w) | Tok::TypedIdent(w, _) => match w.to_ascii_lowercase().as_str() {
                "true" => Expr::Lit(Value::Boolean(true)),
                "false" => Expr::Lit(Value::Boolean(false)),
                "nothing" => Expr::Lit(Value::Object(None)),
                "empty" => Expr::Lit(Value::Empty),
                "null" => Expr::Lit(Value::Null),
                "me" => Expr::Me,
                "new" => {
                    let mut name = self.ident()?;
                    while self.eat_sym(".") {
                        name = self.ident()?;
                    }
                    Expr::New(name)
                }
                "typeof" => {
                    let e = self.postfix_expr()?;
                    self.expect_kw("Is")?;
                    Expr::TypeOf(Box::new(e), self.type_name()?)
                }
                "addressof" => Expr::Lit(Value::str(self.ident()?)),
                _ => Expr::Ident(w),
            },
            t => return Err(syntax(line, format!("expected an expression, found {t:?}"))),
        })
    }
}

/// Untyped integer literals are Integer when they fit, else Long, else Double ([MS-VBAL] 3.3.2).
fn int_literal(n: i64) -> Value {
    if let Ok(v) = i16::try_from(n) {
        Value::Integer(v)
    } else if let Ok(v) = i32::try_from(n) {
        Value::Long(v)
    } else {
        Value::Double(n as f64)
    }
}

pub fn type_for_char(c: char) -> String {
    match c {
        '%' => "Integer",
        '&' => "Long",
        '^' => "LongLong",
        '!' => "Single",
        '#' => "Double",
        '@' => "Currency",
        '$' => "String",
        _ => "Variant",
    }
    .to_owned()
}

fn is_reserved(w: &str) -> bool {
    const R: [&str; 24] = [
        "else", "end", "next", "loop", "wend", "case", "if", "then", "for", "do", "while", "with", "dim", "set",
        "call", "exit", "goto", "resume", "on", "select", "redim", "const", "static", "return",
    ];
    R.iter().any(|r| r.eq_ignore_ascii_case(w))
}

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::panic)]
mod tests {
    use super::*;

    fn module(src: &str) -> Module {
        parse_module("M", ModuleType::Standard, src).unwrap()
    }

    #[test]
    fn procedures_and_statements() {
        let m = module(
            "Option Explicit\nPublic total As Long\nSub Test(ByVal n As Long, Optional s As String = \"x\")\n  Dim i As Integer, a(1 To 3) As Long\n  For i = 1 To n Step 2\n    total = total + i\n  Next i\n  If total > 3 Then MsgBox \"big\" Else MsgBox \"small\"\n  Select Case total\n    Case 1, 2 To 4, Is > 10: Debug.Print \"x\"; total\n    Case Else\n  End Select\nEnd Sub\nFunction F() As Double\n  F = 2 ^ -1\nEnd Function\n",
        );
        assert!(m.options.explicit);
        assert_eq!(m.vars.len(), 1);
        assert_eq!(m.procs.len(), 2);
        assert_eq!(m.procs[0].params.len(), 2);
        assert!(m.procs[0].params[1].optional);
        assert_eq!(m.procs[0].body.len(), 4);
    }

    #[test]
    fn calls_without_parens_and_paren_first_arg() {
        let s = parse_statements(
            "MsgBox (\"a\") & \"b\", vbOKOnly\nRange(\"A1\").Copy Destination:=Range(\"B1\")\nFoo (1), 2\n",
        )
        .unwrap();
        assert_eq!(s.len(), 3);
        match &s[0].kind {
            StmtKind::Call(Expr::Call(_, args)) => assert_eq!(args.len(), 2),
            k => panic!("{k:?}"),
        }
        match &s[1].kind {
            StmtKind::Call(Expr::Call(_, args)) => assert_eq!(args[0].name.as_deref(), Some("Destination")),
            k => panic!("{k:?}"),
        }
    }

    #[test]
    fn with_labels_errors() {
        let s = parse_statements("On Error GoTo Handler\nWith ActiveSheet\n  .Range(\"A1\").Value = 5\n  .Cells(1, 2) = .Range(\"A1\") * 2\nEnd With\nExit Sub\nHandler:\nResume Next\n").unwrap();
        assert!(matches!(s[0].kind, StmtKind::OnError(OnError::GoTo(_))));
        assert!(matches!(s[1].kind, StmtKind::With { .. }));
        assert!(matches!(s[3].kind, StmtKind::Label(ref l) if l == "handler"));
    }

    #[test]
    fn precedence() {
        let s = parse_statements("x = -2 ^ 2\ny = Not a = b And c\nz = 1 + 2 * 3 & \"!\"").unwrap();
        match &s[0].kind {
            StmtKind::Let { value: Expr::Neg(inner), .. } => assert!(matches!(**inner, Expr::Binary(BinOp::Pow, ..))),
            k => panic!("{k:?}"),
        }
        match &s[1].kind {
            StmtKind::Let { value: Expr::Binary(BinOp::And, l, _), .. } => assert!(matches!(**l, Expr::Not(_))),
            k => panic!("{k:?}"),
        }
        match &s[2].kind {
            StmtKind::Let { value: Expr::Binary(BinOp::Concat, ..), .. } => {}
            k => panic!("{k:?}"),
        }
    }

    #[test]
    fn conditional_compilation() {
        let m = module(
            "#If VBA7 Then\nPrivate Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\n#Else\nThis is not valid code\n#End If\nSub A()\nEnd Sub\n",
        );
        assert_eq!(m.declares, vec!["GetTickCount".to_string()]);
    }

    #[test]
    fn nested_next_and_do_loops() {
        let s = parse_statements("For i = 1 To 3\nFor j = 1 To 3\nk = k + 1\nNext j, i\nDo While k > 0\nk = k - 1\nLoop\nDo\nk = k + 1\nLoop Until k = 5").unwrap();
        assert_eq!(s.len(), 3);
    }
}
