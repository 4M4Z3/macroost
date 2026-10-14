//! The formula engine: parses formula text into an expression tree and evaluates it over any
//! [`Workbook`]. [`crate::MemoryWorkbook`] uses it so formula cells read back values; the object
//! model uses the same function library for `WorksheetFunction`, `Application.<Function>` and
//! `Evaluate`. A real spreadsheet app may bring its own engine through the Workbook trait.
//!
//! Covers references (A1, ranges, whole rows/columns, sheet-qualified, defined names), array
//! constants, the operators with their documented precedence, implicit intersection, and the
//! common worksheet functions (math, statistics, lookup, text, date, logical, information).

use macroost_lang::Value;
use macroost_lang::value::{format_float, parse_number, serial_to_hms, serial_to_ymd, ymd_to_serial};

use crate::address::{MAX_COLS, MAX_ROWS, Rect, column_letters, parse_reference};
use crate::format::{error_code, format_value};
use crate::range::wildcard_match;
use crate::workbook::{SheetId, Workbook};

pub(crate) const NULL_ERR: i32 = 2000;
pub(crate) const DIV0: i32 = 2007;
pub(crate) const VALUE: i32 = 2015;
pub(crate) const REF: i32 = 2023;
pub(crate) const NAME: i32 = 2029;
pub(crate) const NUM: i32 = 2036;
pub(crate) const NA: i32 = 2042;

/// A reference: one or more rectangles on a sheet.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Area {
    pub sheet: SheetId,
    pub rects: Vec<Rect>,
}

/// A 2-D block of scalar values, row-major.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Grid {
    pub rows: usize,
    pub cols: usize,
    pub data: Vec<V>,
}

impl Grid {
    pub fn get(&self, r: usize, c: usize) -> V {
        self.data.get(r * self.cols + c).cloned().unwrap_or(V::Empty)
    }
}

/// An evaluated value.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum V {
    Empty,
    Num(f64),
    Str(String),
    Bool(bool),
    Err(i32),
    Ref(Area),
    Arr(Grid),
    /// A skipped argument (`IF(A1,,1)`).
    Missing,
}

impl V {
    fn is_err(&self) -> bool {
        matches!(self, V::Err(_))
    }
}

/// Converts a cell value to an engine value.
pub(crate) fn from_value(v: &Value) -> V {
    match v {
        Value::Empty | Value::Null => V::Empty,
        Value::Boolean(b) => V::Bool(*b),
        Value::String(s) => V::Str(s.to_string()),
        Value::Error(n) => V::Err(*n),
        Value::Object(_) | Value::Array(_) => V::Err(VALUE),
        other => other.to_f64().map(V::Num).unwrap_or(V::Err(VALUE)),
    }
}

/// Converts a scalar engine value to a cell value (`Empty` stays `Empty`).
pub(crate) fn to_value(v: &V) -> Value {
    match v {
        V::Empty | V::Missing => Value::Empty,
        V::Num(x) => Value::Double(*x),
        V::Str(s) => Value::from(s.as_str()),
        V::Bool(b) => Value::Boolean(*b),
        V::Err(n) => Value::Error(*n),
        V::Ref(_) | V::Arr(_) => Value::Error(VALUE),
    }
}

/// Where a formula is evaluated.
pub(crate) struct Ctx<'a> {
    pub book: &'a dyn Workbook,
    pub sheet: SheetId,
    /// The formula's own cell, if it lives in one (implicit intersection, `ROW()`).
    pub cell: Option<(u32, u32)>,
    pub now: f64,
}

impl Ctx<'_> {
    fn cell(&self, sheet: SheetId, r: u32, c: u32) -> V {
        from_value(&self.book.cell_value(sheet, r, c))
    }

    fn sheet_by_name(&self, name: &str) -> Option<SheetId> {
        self.book
            .sheet_ids()
            .into_iter()
            .find(|&id| self.book.sheet_name(id).is_some_and(|n| n.eq_ignore_ascii_case(name)))
    }

    fn name_refers(&self, name: &str, scope: SheetId) -> Option<String> {
        let names = self.book.defined_names();
        let find = |s: Option<SheetId>| {
            names.iter().find(|(n, sc, _)| *sc == s && n.eq_ignore_ascii_case(name)).map(|(_, _, r)| r.clone())
        };
        find(Some(scope)).or_else(|| find(None))
    }

    /// Stored (non-blank) cells of a rectangle, sorted row-major.
    fn cells_of(&self, sheet: SheetId, r: &Rect) -> Vec<(u32, u32, V)> {
        let occ = if r.area() <= 4096 {
            let mut v = Vec::new();
            for row in r.r1..=r.r2 {
                for col in r.c1..=r.c2 {
                    v.push((row, col));
                }
            }
            v
        } else {
            self.book.occupied_cells(sheet, *r)
        };
        occ.into_iter()
            .filter_map(|(row, col)| {
                let v = self.cell(sheet, row, col);
                (v != V::Empty).then_some((row, col, v))
            })
            .collect()
    }

    /// A rectangle clipped to the used range (for whole-column references).
    fn clip(&self, sheet: SheetId, r: &Rect) -> Option<Rect> {
        if r.area() <= 4096 {
            return Some(*r);
        }
        let u = self.book.used_range(sheet)?;
        let (r2, c2) = (r.r2.min(u.r2), r.c2.min(u.c2));
        (r2 >= r.r1 && c2 >= r.c1).then(|| Rect::new(r.r1, r.c1, r2, c2))
    }

    /// A reference or array as a grid of values (huge references clipped to the used range).
    pub(crate) fn grid(&self, v: &V) -> Grid {
        match v {
            V::Ref(a) => {
                let r = a.rects[0];
                let Some(c) = self.clip(a.sheet, &r) else { return Grid { rows: 0, cols: 0, data: vec![] } };
                let (rows, cols) = (c.rows() as usize, c.cols() as usize);
                let mut data = vec![V::Empty; rows * cols];
                for (row, col, val) in self.cells_of(a.sheet, &c) {
                    data[(row - c.r1) as usize * cols + (col - c.c1) as usize] = val;
                }
                Grid { rows, cols, data }
            }
            V::Arr(g) => g.clone(),
            other => Grid { rows: 1, cols: 1, data: vec![other.clone()] },
        }
    }

    /// Logical size of a reference or array.
    fn dims(&self, v: &V) -> (usize, usize) {
        match v {
            V::Ref(a) => (a.rects[0].rows() as usize, a.rects[0].cols() as usize),
            V::Arr(g) => (g.rows, g.cols),
            _ => (1, 1),
        }
    }

    /// A value where a single value is wanted: single cells read, larger references use implicit
    /// intersection with the formula's own row or column.
    pub(crate) fn scalar(&self, v: V) -> V {
        match v {
            V::Ref(a) => {
                if a.rects.len() != 1 {
                    return V::Err(VALUE);
                }
                let r = a.rects[0];
                if r.is_cell() {
                    return self.cell(a.sheet, r.r1, r.c1);
                }
                if let Some((row, col)) = self.cell {
                    if r.c1 == r.c2 && (r.r1..=r.r2).contains(&row) {
                        return self.cell(a.sheet, row, r.c1);
                    }
                    if r.r1 == r.r2 && (r.c1..=r.c2).contains(&col) {
                        return self.cell(a.sheet, r.r1, col);
                    }
                }
                V::Err(VALUE)
            }
            V::Arr(g) => g.data.first().cloned().unwrap_or(V::Empty),
            V::Missing => V::Empty,
            other => other,
        }
    }
}

// ---------------------------------------------------------------- parsing

#[derive(Clone, Debug)]
enum E {
    Lit(V),
    Ref(Option<String>, Rect),
    BadRef,
    Name(Option<String>, String),
    Func(String, Vec<E>),
    Neg(Box<E>),
    Pct(Box<E>),
    Bin(&'static str, Box<E>, Box<E>),
    Range(Box<E>, Box<E>),
    Array(Vec<Vec<V>>),
    Missing,
}

struct Parser {
    ch: Vec<char>,
    i: usize,
}

type PResult<T> = std::result::Result<T, ()>;

impl Parser {
    fn ws(&mut self) {
        while self.ch.get(self.i).is_some_and(|c| c.is_whitespace()) {
            self.i += 1;
        }
    }
    fn peek(&mut self) -> Option<char> {
        self.ws();
        self.ch.get(self.i).copied()
    }
    fn eat(&mut self, s: &str) -> bool {
        self.ws();
        let n = s.chars().count();
        if self.ch.len() >= self.i + n && self.ch[self.i..self.i + n].iter().copied().eq(s.chars()) {
            self.i += n;
            true
        } else {
            false
        }
    }

    fn expr(&mut self) -> PResult<E> {
        self.comparison()
    }

    fn comparison(&mut self) -> PResult<E> {
        let mut left = self.concat()?;
        loop {
            let op = ["<=", ">=", "<>", "=", "<", ">"].into_iter().find(|op| self.eat(op));
            let Some(op) = op else { return Ok(left) };
            let right = self.concat()?;
            left = E::Bin(op, Box::new(left), Box::new(right));
        }
    }

    fn concat(&mut self) -> PResult<E> {
        let mut left = self.additive()?;
        while self.eat("&") {
            let right = self.additive()?;
            left = E::Bin("&", Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> PResult<E> {
        let mut left = self.term()?;
        loop {
            let op = if self.eat("+") {
                "+"
            } else if self.eat("-") {
                "-"
            } else {
                return Ok(left);
            };
            let right = self.term()?;
            left = E::Bin(op, Box::new(left), Box::new(right));
        }
    }

    fn term(&mut self) -> PResult<E> {
        let mut left = self.power()?;
        loop {
            let op = if self.eat("*") {
                "*"
            } else if self.eat("/") {
                "/"
            } else {
                return Ok(left);
            };
            let right = self.power()?;
            left = E::Bin(op, Box::new(left), Box::new(right));
        }
    }

    fn power(&mut self) -> PResult<E> {
        let mut left = self.percent()?;
        while self.eat("^") {
            let right = self.percent()?;
            left = E::Bin("^", Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn percent(&mut self) -> PResult<E> {
        let mut v = self.unary()?;
        while self.eat("%") {
            v = E::Pct(Box::new(v));
        }
        Ok(v)
    }

    fn unary(&mut self) -> PResult<E> {
        if self.eat("-") {
            return Ok(E::Neg(Box::new(self.unary()?)));
        }
        if self.eat("+") {
            return self.unary();
        }
        let mut v = self.primary()?;
        // Range operator between references.
        while self.peek() == Some(':') {
            self.i += 1;
            let r = self.primary()?;
            v = E::Range(Box::new(v), Box::new(r));
        }
        Ok(v)
    }

    fn string(&mut self) -> String {
        // At the opening quote.
        self.i += 1;
        let mut s = String::new();
        while let Some(&c) = self.ch.get(self.i) {
            self.i += 1;
            if c == '"' {
                if self.ch.get(self.i) == Some(&'"') {
                    s.push('"');
                    self.i += 1;
                    continue;
                }
                break;
            }
            s.push(c);
        }
        s
    }

    fn number(&mut self) -> f64 {
        let start = self.i;
        while self.ch.get(self.i).is_some_and(|c| c.is_ascii_digit() || *c == '.') {
            self.i += 1;
        }
        if self.ch.get(self.i).is_some_and(|c| *c == 'e' || *c == 'E')
            && self.ch.get(self.i + 1).is_some_and(|c| c.is_ascii_digit() || *c == '+' || *c == '-')
        {
            self.i += 2;
            while self.ch.get(self.i).is_some_and(char::is_ascii_digit) {
                self.i += 1;
            }
        }
        self.ch[start..self.i].iter().collect::<String>().parse().unwrap_or(f64::NAN)
    }

    fn array_const(&mut self) -> PResult<E> {
        self.i += 1; // {
        let mut rows = vec![Vec::new()];
        loop {
            self.ws();
            let neg = self.eat("-");
            let v = match self.peek() {
                Some('"') => V::Str(self.string()),
                Some(c) if c.is_ascii_digit() || c == '.' => {
                    let n = self.number();
                    V::Num(if neg { -n } else { n })
                }
                Some('#') => {
                    let w = self.word_raw();
                    V::Err(error_code(&w).ok_or(())?)
                }
                Some(_) => {
                    let w = self.word_raw().to_ascii_uppercase();
                    match w.as_str() {
                        "TRUE" => V::Bool(true),
                        "FALSE" => V::Bool(false),
                        _ => return Err(()),
                    }
                }
                None => return Err(()),
            };
            rows.last_mut().ok_or(())?.push(v);
            match self.peek() {
                Some(',') => self.i += 1,
                Some(';') => {
                    self.i += 1;
                    rows.push(Vec::new());
                }
                Some('}') => {
                    self.i += 1;
                    break;
                }
                _ => return Err(()),
            }
        }
        let w = rows[0].len();
        if rows.iter().any(|r| r.len() != w) {
            return Err(());
        }
        Ok(E::Array(rows))
    }

    /// Characters of an identifier, reference or error literal.
    fn word_raw(&mut self) -> String {
        self.ws();
        let start = self.i;
        while self
            .ch
            .get(self.i)
            .is_some_and(|c| c.is_alphanumeric() || matches!(c, '_' | '.' | '$' | '#' | '!' | '?' | '/' | '\\'))
        {
            // '!' ends an unquoted sheet name; stop after it so the reference is read separately.
            let c = self.ch[self.i];
            self.i += 1;
            if c == '!' || (c == '/' && start + 1 != self.i && self.ch[start] != '#') {
                if c == '/' {
                    self.i -= 1;
                }
                break;
            }
        }
        self.ch[start..self.i].iter().collect()
    }

    fn primary(&mut self) -> PResult<E> {
        let Some(c) = self.peek() else { return Err(()) };
        match c {
            '(' => {
                self.i += 1;
                let v = self.expr()?;
                if !self.eat(")") {
                    return Err(());
                }
                Ok(v)
            }
            '"' => Ok(E::Lit(V::Str(self.string()))),
            '{' => self.array_const(),
            '#' => {
                // Error literal, possibly followed by a reference it replaced (`#REF!A1`).
                let start = self.i;
                for (text, code) in [
                    ("#NULL!", NULL_ERR),
                    ("#DIV/0!", DIV0),
                    ("#VALUE!", VALUE),
                    ("#REF!", REF),
                    ("#NAME?", NAME),
                    ("#NUM!", NUM),
                    ("#N/A", NA),
                ] {
                    if self.eat_ci(text) {
                        if code == REF {
                            // Swallow a trailing reference.
                            while self.ch.get(self.i).is_some_and(|c| c.is_alphanumeric() || *c == '$') {
                                self.i += 1;
                            }
                            return Ok(E::BadRef);
                        }
                        return Ok(E::Lit(V::Err(code)));
                    }
                }
                self.i = start;
                Err(())
            }
            '\'' => {
                // Quoted sheet name.
                self.i += 1;
                let mut name = String::new();
                loop {
                    let ch = *self.ch.get(self.i).ok_or(())?;
                    self.i += 1;
                    if ch == '\'' {
                        if self.ch.get(self.i) == Some(&'\'') {
                            name.push('\'');
                            self.i += 1;
                            continue;
                        }
                        break;
                    }
                    name.push(ch);
                }
                if self.ch.get(self.i) != Some(&'!') {
                    return Err(());
                }
                self.i += 1;
                self.reference(Some(name))
            }
            c if c.is_ascii_digit() || c == '.' => {
                // A number, or a row range `1:3`.
                let save = self.i;
                let mut j = self.i;
                while self.ch.get(j).is_some_and(char::is_ascii_digit) {
                    j += 1;
                }
                if self.ch.get(j) == Some(&':') {
                    self.i = save;
                    if let Some(e) = self.try_ref(None) {
                        return Ok(e);
                    }
                    self.i = save;
                }
                Ok(E::Lit(V::Num(self.number())))
            }
            _ => {
                let save = self.i;
                let w = self.word_raw();
                if let Some(sheet) = w.strip_suffix('!') {
                    return self.reference(Some(sheet.to_owned()));
                }
                if self.peek() == Some('(') {
                    self.i += 1;
                    let mut args = Vec::new();
                    if !self.eat(")") {
                        loop {
                            match self.peek() {
                                Some(',') | Some(')') => args.push(E::Missing),
                                _ => args.push(self.expr()?),
                            }
                            if self.eat(",") {
                                continue;
                            }
                            if !self.eat(")") {
                                return Err(());
                            }
                            break;
                        }
                    }
                    let name = w.to_ascii_uppercase();
                    let name = name.strip_prefix("_XLFN.").unwrap_or(&name).to_owned();
                    return Ok(E::Func(name, args));
                }
                let up = w.to_ascii_uppercase();
                if up == "TRUE" || up == "FALSE" {
                    return Ok(E::Lit(V::Bool(up == "TRUE")));
                }
                self.i = save;
                if let Some(e) = self.try_ref(None) {
                    return Ok(e);
                }
                self.i = save;
                let w = self.word_raw();
                if w.is_empty() {
                    return Err(());
                }
                Ok(E::Name(None, w))
            }
        }
    }

    fn eat_ci(&mut self, s: &str) -> bool {
        let n = s.chars().count();
        if self.ch.len() >= self.i + n
            && self.ch[self.i..self.i + n].iter().zip(s.chars()).all(|(a, b)| a.eq_ignore_ascii_case(&b))
        {
            self.i += n;
            true
        } else {
            false
        }
    }

    /// After `Sheet!`: a reference, or a name scoped to that sheet.
    fn reference(&mut self, sheet: Option<String>) -> PResult<E> {
        if self.eat_ci("#REF!") {
            return Ok(E::BadRef);
        }
        let save = self.i;
        if let Some(e) = self.try_ref(sheet.clone()) {
            return Ok(e);
        }
        self.i = save;
        let w = self.word_raw();
        if w.is_empty() {
            return Err(());
        }
        Ok(E::Name(sheet, w))
    }

    /// A1 reference at the cursor: cell, cell:cell, col:col or row:row.
    fn try_ref(&mut self, sheet: Option<String>) -> Option<E> {
        let start = self.i;
        let mut j = self.i;
        while self.ch.get(j).is_some_and(|c| c.is_ascii_alphanumeric() || *c == '$' || *c == ':') {
            j += 1;
        }
        // The token must not continue as an identifier or a call.
        if self.ch.get(j).is_some_and(|c| c.is_alphanumeric() || matches!(c, '_' | '.' | '(' | '!')) {
            return None;
        }
        let text: String = self.ch[start..j].iter().collect();
        let mut candidates: Vec<usize> = vec![text.len()];
        // `A1:B2` may be followed by `:C3` handled by the range operator; try the longest
        // prefix that parses, preferring the whole token.
        if let Some(p) = text.find(':') {
            candidates.push(p);
        }
        for len in candidates {
            let t = &text[..len];
            if t.is_empty() {
                continue;
            }
            if let Some(p) = parse_reference(t)
                && p.sheet.is_none()
                && p.areas.len() == 1
            {
                self.i = start + t.chars().count();
                return Some(E::Ref(sheet, p.areas[0]));
            }
        }
        None
    }
}

fn parse(formula: &str) -> PResult<E> {
    let body = formula.trim_start();
    let body = body.strip_prefix('=').unwrap_or(body);
    let mut p = Parser { ch: body.chars().collect(), i: 0 };
    if p.peek().is_none() {
        return Err(());
    }
    let e = p.expr()?;
    if p.peek().is_some() {
        return Err(());
    }
    Ok(e)
}

// ---------------------------------------------------------------- evaluation

/// Evaluates formula text in `ctx`. Unparseable text is `#NAME?`.
pub(crate) fn eval_text(formula: &str, ctx: &Ctx<'_>) -> V {
    match parse(formula) {
        Ok(e) => eval(&e, ctx),
        Err(()) => V::Err(NAME),
    }
}

/// The number format a newly entered formula gives a General cell: dates for `DATE()`,
/// `TODAY()` and date arithmetic, times for `TIME()`, and the format of a reference the formula
/// is (`=A1`, `=(A1)`, `=A1+30`, `=SUM(A1)`). `fmt_of` reads a referenced cell's format.
pub(crate) fn auto_format(formula: &str, fmt_of: &dyn Fn(Option<&str>, Rect) -> Option<String>) -> Option<String> {
    let e = parse(formula).ok()?;
    format_kind(&e, fmt_of)
}

fn format_kind(e: &E, fmt_of: &dyn Fn(Option<&str>, Rect) -> Option<String>) -> Option<String> {
    let is_date = |f: &Option<String>| f.as_deref().is_some_and(crate::format::is_date_format);
    match e {
        E::Ref(sheet, r) => fmt_of(sheet.as_deref(), *r),
        E::Func(name, args) => match name.as_str() {
            "DATE" | "TODAY" | "DATEVALUE" => Some("m/d/yyyy".into()),
            "TIME" | "TIMEVALUE" => Some("h:mm AM/PM".into()),
            "NOW" => Some("m/d/yyyy h:mm".into()),
            "SUM" | "MIN" | "MAX" | "AVERAGE" | "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
                args.first().and_then(|a| format_kind(a, fmt_of))
            }
            _ => None,
        },
        E::Bin(op @ ("+" | "-"), a, b) => {
            let (ka, kb) = (format_kind(a, fmt_of), format_kind(b, fmt_of));
            if *op == "-" && is_date(&ka) && is_date(&kb) {
                return None;
            }
            match (ka, *op) {
                (Some(f), _) => Some(f),
                (None, "+") => kb.filter(|f| crate::format::is_date_format(f)),
                _ => None,
            }
        }
        _ => None,
    }
}

/// A formula cell's value: references read, empty becomes 0.
pub(crate) fn evaluate_cell(book: &dyn Workbook, sheet: SheetId, row: u32, col: u32, formula: &str, now: f64) -> Value {
    let ctx = Ctx { book, sheet, cell: Some((row, col)), now };
    let v = eval_text(formula, &ctx);
    let v = ctx.scalar(v);
    match v {
        V::Empty => Value::Double(0.0),
        other => to_value(&other),
    }
}

fn eval(e: &E, ctx: &Ctx<'_>) -> V {
    match e {
        E::Lit(v) => v.clone(),
        E::Missing => V::Missing,
        E::BadRef => V::Err(REF),
        E::Ref(sheet, r) => {
            let id = match sheet {
                Some(n) => match ctx.sheet_by_name(n) {
                    Some(id) => id,
                    None => return V::Err(REF),
                },
                None => ctx.sheet,
            };
            V::Ref(Area { sheet: id, rects: vec![*r] })
        }
        E::Name(sheet, n) => {
            let scope = match sheet {
                Some(s) => match ctx.sheet_by_name(s) {
                    Some(id) => id,
                    None => return V::Err(REF),
                },
                None => ctx.sheet,
            };
            match ctx.name_refers(n, scope) {
                Some(refers) => {
                    let inner = Ctx { book: ctx.book, sheet: scope, cell: ctx.cell, now: ctx.now };
                    eval_text(&refers, &inner)
                }
                None => V::Err(NAME),
            }
        }
        E::Array(rows) => {
            let r = rows.len();
            let c = rows.first().map_or(0, Vec::len);
            V::Arr(Grid { rows: r, cols: c, data: rows.iter().flatten().cloned().collect() })
        }
        E::Neg(x) => unary(ctx, eval(x, ctx), |n| -n),
        E::Pct(x) => unary(ctx, eval(x, ctx), |n| n / 100.0),
        E::Range(a, b) => {
            let (a, b) = (eval(a, ctx), eval(b, ctx));
            match (a, b) {
                (V::Ref(x), V::Ref(y)) if x.sheet == y.sheet => {
                    let r = x.rects.iter().chain(y.rects.iter()).fold(x.rects[0], |acc, r| acc.bounding(r));
                    V::Ref(Area { sheet: x.sheet, rects: vec![r] })
                }
                (V::Err(n), _) | (_, V::Err(n)) => V::Err(n),
                _ => V::Err(VALUE),
            }
        }
        E::Bin(op, a, b) => {
            let (a, b) = (eval(a, ctx), eval(b, ctx));
            binary(ctx, op, a, b)
        }
        E::Func(name, args) => {
            let vals: Vec<V> = args.iter().map(|a| eval(a, ctx)).collect();
            call(ctx, name, vals)
        }
    }
}

fn unary(ctx: &Ctx<'_>, v: V, f: impl Fn(f64) -> f64) -> V {
    match v {
        V::Arr(g) => V::Arr(Grid {
            rows: g.rows,
            cols: g.cols,
            data: g.data.into_iter().map(|x| num_of(&x).map(|n| V::Num(f(n))).unwrap_or_else(V::Err)).collect(),
        }),
        other => match num_of(&ctx.scalar(other)) {
            Ok(n) => V::Num(f(n)),
            Err(e) => V::Err(e),
        },
    }
}

/// Number coercion of a scalar: empty 0, TRUE 1, numeric text, else `#VALUE!`.
pub(crate) fn num_of(v: &V) -> std::result::Result<f64, i32> {
    match v {
        V::Empty | V::Missing => Ok(0.0),
        V::Num(n) => Ok(*n),
        V::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        V::Str(s) => text_number(s).ok_or(VALUE),
        V::Err(e) => Err(*e),
        V::Ref(_) | V::Arr(_) => Err(VALUE),
    }
}

/// Text a formula accepts as a number: plain numbers, percentages, currency, dates and times.
fn text_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(n) = parse_number(t)
        && !t.starts_with("&")
    {
        return Some(n);
    }
    match crate::state::parse_typed(t) {
        crate::state::Input::Const(Value::Double(x), _) => Some(x),
        _ => None,
    }
}

/// Text coercion of a scalar.
pub(crate) fn text_of(v: &V) -> std::result::Result<String, i32> {
    match v {
        V::Empty | V::Missing => Ok(String::new()),
        V::Num(n) => Ok(num_text(*n)),
        V::Str(s) => Ok(s.clone()),
        V::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_owned()),
        V::Err(e) => Err(*e),
        V::Ref(_) | V::Arr(_) => Err(VALUE),
    }
}

/// A number as formula text (`=A1&""`): up to 15 significant digits.
pub(crate) fn num_text(n: f64) -> String {
    let s = format_float(n, 15);
    if s == "-0" { "0".into() } else { s }
}

fn bool_of(v: &V) -> std::result::Result<bool, i32> {
    match v {
        V::Bool(b) => Ok(*b),
        V::Empty | V::Missing => Ok(false),
        V::Num(n) => Ok(*n != 0.0),
        V::Str(s) => {
            if s.eq_ignore_ascii_case("true") {
                Ok(true)
            } else if s.eq_ignore_ascii_case("false") {
                Ok(false)
            } else {
                Err(VALUE)
            }
        }
        V::Err(e) => Err(*e),
        _ => Err(VALUE),
    }
}

/// Ordering of formula comparisons: numbers < text < logical; text case-insensitive.
pub(crate) fn compare(a: &V, b: &V) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let rank = |v: &V| match v {
        V::Num(_) | V::Empty => 0,
        V::Str(_) => 1,
        V::Bool(_) => 2,
        _ => 3,
    };
    match (a, b) {
        (V::Empty, V::Str(s)) => "".cmp(s.to_lowercase().as_str()),
        (V::Str(s), V::Empty) => s.to_lowercase().as_str().cmp(""),
        (V::Empty, V::Bool(x)) => false.cmp(x),
        (V::Bool(x), V::Empty) => x.cmp(&false),
        _ => {
            let (ra, rb) = (rank(a), rank(b));
            if ra != rb {
                return ra.cmp(&rb);
            }
            match (a, b) {
                (V::Str(x), V::Str(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
                (V::Bool(x), V::Bool(y)) => x.cmp(y),
                _ => {
                    let x = num_of(a).unwrap_or(0.0);
                    let y = num_of(b).unwrap_or(0.0);
                    x.partial_cmp(&y).unwrap_or(Ordering::Equal)
                }
            }
        }
    }
}

fn binary(ctx: &Ctx<'_>, op: &str, a: V, b: V) -> V {
    // Array operands work element by element.
    if matches!(a, V::Arr(_)) || matches!(b, V::Arr(_)) {
        let ga = ctx.grid(&a);
        let gb = ctx.grid(&b);
        let rows = ga.rows.max(gb.rows);
        let cols = ga.cols.max(gb.cols);
        let pick = |g: &Grid, r: usize, c: usize| {
            let rr = if g.rows == 1 { 0 } else { r };
            let cc = if g.cols == 1 { 0 } else { c };
            if rr < g.rows && cc < g.cols { g.get(rr, cc) } else { V::Err(NA) }
        };
        let mut data = Vec::with_capacity(rows * cols);
        for r in 0..rows {
            for c in 0..cols {
                data.push(binary_scalar(op, pick(&ga, r, c), pick(&gb, r, c)));
            }
        }
        return V::Arr(Grid { rows, cols, data });
    }
    binary_scalar(op, ctx.scalar(a), ctx.scalar(b))
}

fn binary_scalar(op: &str, a: V, b: V) -> V {
    let arith = |f: &dyn Fn(f64, f64) -> std::result::Result<f64, i32>| -> V {
        match (num_of(&a), num_of(&b)) {
            (Ok(x), Ok(y)) => match f(x, y) {
                Ok(n) if n.is_finite() => V::Num(n),
                Ok(_) => V::Err(NUM),
                Err(e) => V::Err(e),
            },
            (Err(e), _) | (_, Err(e)) => V::Err(e),
        }
    };
    match op {
        "+" => arith(&|x, y| Ok(x + y)),
        "-" => arith(&|x, y| Ok(x - y)),
        "*" => arith(&|x, y| Ok(x * y)),
        "/" => arith(&|x, y| if y == 0.0 { Err(DIV0) } else { Ok(x / y) }),
        "^" => arith(&|x, y| if x == 0.0 && y == 0.0 { Err(NUM) } else { Ok(x.powf(y)) }),
        "&" => match (text_of(&a), text_of(&b)) {
            (Ok(x), Ok(y)) => V::Str(x + &y),
            (Err(e), _) | (_, Err(e)) => V::Err(e),
        },
        _ => {
            if let V::Err(e) = a {
                return V::Err(e);
            }
            if let V::Err(e) = b {
                return V::Err(e);
            }
            use std::cmp::Ordering::*;
            let o = compare(&a, &b);
            V::Bool(match op {
                "=" => o == Equal,
                "<>" => o != Equal,
                "<" => o == Less,
                ">" => o == Greater,
                "<=" => o != Greater,
                _ => o != Less,
            })
        }
    }
}

// ---------------------------------------------------------------- argument helpers

/// Every value an argument contributes: (value, came from a reference or array).
fn flatten(ctx: &Ctx<'_>, args: &[V]) -> Vec<(V, bool)> {
    let mut out = Vec::new();
    for a in args {
        match a {
            V::Ref(area) => {
                for r in &area.rects {
                    for (_, _, v) in ctx.cells_of(area.sheet, r) {
                        out.push((v, true));
                    }
                }
            }
            V::Arr(g) => out.extend(g.data.iter().cloned().map(|v| (v, true))),
            V::Missing => {}
            other => out.push((other.clone(), false)),
        }
    }
    out
}

/// Numbers for SUM/AVERAGE/MAX...: references and arrays count only numbers; direct arguments
/// are coerced. Errors propagate.
fn numbers(ctx: &Ctx<'_>, args: &[V]) -> std::result::Result<Vec<f64>, i32> {
    let mut out = Vec::new();
    for (v, from_ref) in flatten(ctx, args) {
        match v {
            V::Err(e) => return Err(e),
            V::Num(n) => out.push(n),
            ref x if !from_ref => out.push(num_of(x)?),
            _ => {}
        }
    }
    Ok(out)
}

/// Numbers of the rows a SUBTOTAL/AGGREGATE sees (hidden rows skipped when asked).
fn visible_numbers(
    ctx: &Ctx<'_>,
    args: &[V],
    skip_hidden: bool,
    skip_errors: bool,
) -> std::result::Result<Vec<f64>, i32> {
    let mut out = Vec::new();
    for a in args {
        match a {
            V::Ref(area) => {
                for r in &area.rects {
                    for (row, _, v) in ctx.cells_of(area.sheet, r) {
                        if ctx.book.row_filtered(area.sheet, row) || skip_hidden && ctx.book.row_hidden(area.sheet, row)
                        {
                            continue;
                        }
                        match v {
                            V::Num(n) => out.push(n),
                            V::Err(e) if !skip_errors => return Err(e),
                            _ => {}
                        }
                    }
                }
            }
            other => out.extend(numbers(ctx, std::slice::from_ref(other))?),
        }
    }
    Ok(out)
}

fn arg(args: &[V], i: usize) -> V {
    args.get(i).cloned().unwrap_or(V::Missing)
}

fn sc(ctx: &Ctx<'_>, args: &[V], i: usize) -> V {
    ctx.scalar(arg(args, i))
}

fn n_arg(ctx: &Ctx<'_>, args: &[V], i: usize) -> std::result::Result<f64, i32> {
    num_of(&sc(ctx, args, i))
}

fn opt_n(ctx: &Ctx<'_>, args: &[V], i: usize, default: f64) -> std::result::Result<f64, i32> {
    match arg(args, i) {
        V::Missing => Ok(default),
        _ => n_arg(ctx, args, i),
    }
}

fn t_arg(ctx: &Ctx<'_>, args: &[V], i: usize) -> std::result::Result<String, i32> {
    text_of(&sc(ctx, args, i))
}

fn b_arg(ctx: &Ctx<'_>, args: &[V], i: usize, default: bool) -> std::result::Result<bool, i32> {
    match arg(args, i) {
        V::Missing => Ok(default),
        _ => bool_of(&sc(ctx, args, i)),
    }
}

// ---------------------------------------------------------------- criteria

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Op {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Clone, Debug)]
enum CVal {
    Num(f64),
    Text(String),
    Bool(bool),
    Blank,
    Err(i32),
}

#[derive(Clone, Debug)]
pub(crate) struct Crit {
    op: Op,
    val: CVal,
}

pub(crate) fn parse_crit(v: &V) -> Crit {
    match v {
        V::Str(s) => {
            let (op, rest) = if let Some(r) = s.strip_prefix("<=") {
                (Op::Le, r)
            } else if let Some(r) = s.strip_prefix(">=") {
                (Op::Ge, r)
            } else if let Some(r) = s.strip_prefix("<>") {
                (Op::Ne, r)
            } else if let Some(r) = s.strip_prefix('<') {
                (Op::Lt, r)
            } else if let Some(r) = s.strip_prefix('>') {
                (Op::Gt, r)
            } else if let Some(r) = s.strip_prefix('=') {
                (Op::Eq, r)
            } else {
                (Op::Eq, s.as_str())
            };
            let val = if rest.is_empty() {
                CVal::Blank
            } else if let Some(n) = text_number(rest) {
                CVal::Num(n)
            } else if rest.eq_ignore_ascii_case("true") {
                CVal::Bool(true)
            } else if rest.eq_ignore_ascii_case("false") {
                CVal::Bool(false)
            } else if let Some(code) = error_code(rest) {
                CVal::Err(code)
            } else {
                CVal::Text(rest.to_owned())
            };
            Crit { op, val }
        }
        V::Bool(b) => Crit { op: Op::Eq, val: CVal::Bool(*b) },
        V::Empty | V::Missing => Crit { op: Op::Eq, val: CVal::Num(0.0) },
        V::Err(n) => Crit { op: Op::Eq, val: CVal::Err(*n) },
        V::Num(n) => Crit { op: Op::Eq, val: CVal::Num(*n) },
        V::Ref(_) | V::Arr(_) => Crit { op: Op::Eq, val: CVal::Err(VALUE) },
    }
}

fn cmp_op(op: Op, o: std::cmp::Ordering) -> bool {
    use std::cmp::Ordering::*;
    match op {
        Op::Eq => o == Equal,
        Op::Ne => o != Equal,
        Op::Lt => o == Less,
        Op::Le => o != Greater,
        Op::Gt => o == Greater,
        Op::Ge => o != Less,
    }
}

pub(crate) fn crit_matches(c: &Crit, cell: &V) -> bool {
    if c.op == Op::Ne {
        let eq = Crit { op: Op::Eq, val: c.val.clone() };
        return !crit_matches(&eq, cell);
    }
    let blank = matches!(cell, V::Empty) || matches!(cell, V::Str(s) if s.is_empty());
    match &c.val {
        CVal::Blank => c.op == Op::Eq && blank,
        CVal::Num(n) => {
            let x = match cell {
                V::Num(x) => Some(*x),
                V::Str(s) if c.op == Op::Eq => text_number(s),
                _ => None,
            };
            x.and_then(|x| x.partial_cmp(n)).is_some_and(|o| cmp_op(c.op, o))
        }
        CVal::Text(t) => match cell {
            V::Str(s) => {
                if c.op == Op::Eq {
                    wildcard_match(t, s, true, false)
                } else {
                    cmp_op(c.op, s.to_lowercase().cmp(&t.to_lowercase()))
                }
            }
            V::Empty => c.op == Op::Eq && wildcard_match(t, "", true, false) && t.chars().all(|ch| ch == '*'),
            _ => false,
        },
        CVal::Bool(b) => matches!(cell, V::Bool(x) if x == b) && c.op == Op::Eq,
        CVal::Err(e) => matches!(cell, V::Err(x) if x == e) && c.op == Op::Eq,
    }
}

/// Cells of a criteria range by position (relative row, column) -> value, for the stored ones.
fn positions(ctx: &Ctx<'_>, v: &V) -> std::result::Result<(usize, usize, Vec<((usize, usize), V)>), i32> {
    match v {
        V::Ref(a) => {
            let r = a.rects[0];
            let cells = ctx
                .cells_of(a.sheet, &r)
                .into_iter()
                .map(|(row, col, v)| (((row - r.r1) as usize, (col - r.c1) as usize), v))
                .collect();
            Ok((r.rows() as usize, r.cols() as usize, cells))
        }
        V::Arr(g) => {
            let mut cells = Vec::new();
            for r in 0..g.rows {
                for c in 0..g.cols {
                    let v = g.get(r, c);
                    if v != V::Empty {
                        cells.push(((r, c), v));
                    }
                }
            }
            Ok((g.rows, g.cols, cells))
        }
        _ => Err(VALUE),
    }
}

fn value_at(ctx: &Ctx<'_>, v: &V, r: usize, c: usize) -> V {
    match v {
        V::Ref(a) => {
            let base = a.rects[0];
            let (row, col) = (base.r1 as usize + r, base.c1 as usize + c);
            if row > MAX_ROWS as usize || col > MAX_COLS as usize {
                return V::Empty;
            }
            ctx.cell(a.sheet, row as u32, col as u32)
        }
        V::Arr(g) => {
            if r < g.rows && c < g.cols {
                g.get(r, c)
            } else {
                V::Empty
            }
        }
        _ => V::Empty,
    }
}

/// COUNTIFS/SUMIFS core over (range, criterion) pairs: (sum of `sum_range` values, count, numeric count).
fn multi_criteria(ctx: &Ctx<'_>, pairs: &[(V, V)], sum_range: Option<&V>) -> std::result::Result<(f64, f64, f64), i32> {
    let mut shape: Option<(usize, usize)> = None;
    let mut cand: Vec<(usize, usize)> = Vec::new();
    let mut crits = Vec::new();
    for (rng, crit) in pairs {
        let (rows, cols, cells) = positions(ctx, rng)?;
        match shape {
            Some(s) if s != (rows, cols) => return Err(VALUE),
            _ => shape = Some((rows, cols)),
        }
        cand.extend(cells.into_iter().map(|(p, _)| p));
        crits.push(parse_crit(&ctx.scalar(crit.clone())));
    }
    let (rows, cols) = shape.ok_or(VALUE)?;
    if let Some(s) = sum_range {
        let (_, _, cells) = positions(ctx, s)?;
        cand.extend(cells.into_iter().map(|(p, _)| p).filter(|&(r, c)| r < rows && c < cols));
    }
    cand.sort_unstable();
    cand.dedup();
    let (mut sum, mut count, mut ncount) = (0.0, 0.0, 0.0);
    for &(r, c) in &cand {
        if pairs.iter().zip(&crits).all(|((rng, _), cr)| crit_matches(cr, &value_at(ctx, rng, r, c))) {
            count += 1.0;
            if let Some(s) = sum_range {
                match value_at(ctx, s, r, c) {
                    V::Num(n) => {
                        sum += n;
                        ncount += 1.0;
                    }
                    V::Err(e) => return Err(e),
                    _ => {}
                }
            }
        }
    }
    if crits.iter().all(|c| crit_matches(c, &V::Empty)) {
        count += (rows * cols - cand.len()) as f64;
    }
    Ok((sum, count, ncount))
}

/// The sum range of SUMIF/AVERAGEIF takes the criteria range's shape from its top-left cell.
fn reshape(v: &V, rows: usize, cols: usize) -> V {
    match v {
        V::Ref(a) => {
            let r0 = a.rects[0];
            let r2 = (r0.r1 as usize + rows - 1).min(MAX_ROWS as usize) as u32;
            let c2 = (r0.c1 as usize + cols - 1).min(MAX_COLS as usize) as u32;
            V::Ref(Area { sheet: a.sheet, rects: vec![Rect::new(r0.r1, r0.c1, r2, c2)] })
        }
        other => other.clone(),
    }
}

// ---------------------------------------------------------------- lookups

fn same_class(a: &V, b: &V) -> bool {
    matches!((a, b), (V::Num(_), V::Num(_)) | (V::Str(_), V::Str(_)) | (V::Bool(_), V::Bool(_)))
}

fn lookup_equal(lookup: &V, v: &V) -> bool {
    match (lookup, v) {
        (V::Str(p), V::Str(s)) => wildcard_match(p, s, true, false),
        _ => same_class(lookup, v) && compare(lookup, v) == std::cmp::Ordering::Equal,
    }
}

/// MATCH over a vector: 0 exact, 1 largest <= (ascending), -1 smallest >= (descending).
fn match_in(vec: &[V], lookup: &V, mode: i64) -> Option<usize> {
    if mode == 0 {
        return vec.iter().position(|v| lookup_equal(lookup, v));
    }
    // Binary search the way the application does over sorted data.
    let mut best = None;
    for (i, v) in vec.iter().enumerate() {
        if !same_class(lookup, v) {
            continue;
        }
        let o = compare(v, lookup);
        let ok = if mode > 0 { o != std::cmp::Ordering::Greater } else { o != std::cmp::Ordering::Less };
        if ok {
            best = Some(i);
        } else {
            break;
        }
    }
    best
}

fn vector(g: &Grid) -> Option<Vec<V>> {
    if g.rows == 1 {
        Some((0..g.cols).map(|c| g.get(0, c)).collect())
    } else if g.cols == 1 {
        Some((0..g.rows).map(|r| g.get(r, 0)).collect())
    } else {
        None
    }
}

// ---------------------------------------------------------------- functions

fn round_digits(x: f64, digits: f64, mode: i8) -> f64 {
    let d = digits.trunc().clamp(-15.0, 15.0) as i32;
    let scale = 10f64.powi(d.abs());
    let y = if d >= 0 { x * scale } else { x / scale };
    let y: f64 = format!("{y:.12e}").parse().unwrap_or(y);
    let r = match mode {
        0 => y.abs().round().copysign(y),
        1 => y.abs().ceil().copysign(y),
        _ => y.trunc(),
    };
    if d >= 0 { r / scale } else { r * scale }
}

/// the spreadsheet host serial (1 = 1900-01-01, with the 1900 leap-year quirk) to (y, m, d).
pub(crate) fn serial_ymd(serial: f64) -> (i64, u32, u32) {
    let s = serial.floor();
    if s == 0.0 {
        return (1900, 1, 0);
    }
    if s == 60.0 {
        return (1900, 2, 29);
    }
    if s < 60.0 { serial_to_ymd(s + 1.0) } else { serial_to_ymd(s) }
}

/// (y, m, d) to an the spreadsheet host serial.
pub(crate) fn ymd_serial(y: i64, m: i64, d: i64) -> f64 {
    let y = if (0..1900).contains(&y) { y + 1900 } else { y };
    let s = ymd_to_serial(y, m, d);
    if s < 61.0 { s - 1.0 } else { s }
}

fn weekday_of(serial: f64) -> i64 {
    // Serial 1 (1900-01-01) is a Sunday in the reference spreadsheet's numbering.
    (serial.floor() as i64 - 1).rem_euclid(7) + 1
}

fn grid_value(g: Grid) -> V {
    if g.rows == 1 && g.cols == 1 { g.data.into_iter().next().unwrap_or(V::Empty) } else { V::Arr(g) }
}

fn std_dev(v: &[f64], sample: bool) -> std::result::Result<f64, i32> {
    let n = v.len() as f64;
    if n < if sample { 2.0 } else { 1.0 } {
        return Err(DIV0);
    }
    let mean = v.iter().sum::<f64>() / n;
    let ss: f64 = v.iter().map(|x| (x - mean).powi(2)).sum();
    Ok((ss / if sample { n - 1.0 } else { n }).sqrt())
}

fn proper(s: &str) -> String {
    let mut out = String::new();
    let mut prev_letter = false;
    for c in s.chars() {
        if c.is_alphabetic() {
            if prev_letter {
                out.extend(c.to_lowercase());
            } else {
                out.extend(c.to_uppercase());
            }
            prev_letter = true;
        } else {
            out.push(c);
            prev_letter = false;
        }
    }
    out
}

fn fixed_text(x: f64, decimals: f64, no_commas: bool) -> String {
    let d = decimals.trunc() as i32;
    let r = round_digits(x, f64::from(d), 0);
    let fmt = if d > 0 {
        format!("{}0.{}", if no_commas { "" } else { "#,##" }, "0".repeat(d as usize))
    } else {
        format!("{}0", if no_commas { "" } else { "#,##" })
    };
    format_value(&Value::Double(r), &fmt)
}

/// Calls worksheet function `name` (upper case) with evaluated arguments.
pub(crate) fn call(ctx: &Ctx<'_>, name: &str, args: Vec<V>) -> V {
    match call_inner(ctx, name, &args) {
        Ok(v) => v,
        Err(-1) => V::Err(NAME),
        Err(e) => V::Err(e),
    }
}

/// Like [`call`], but `None` for a function this engine does not implement.
pub(crate) fn call_checked(ctx: &Ctx<'_>, name: &str, args: Vec<V>) -> Option<V> {
    match call_inner(ctx, name, &args) {
        Ok(v) => Some(v),
        Err(-1) => None,
        Err(e) => Some(V::Err(e)),
    }
}

type R = std::result::Result<V, i32>;

#[allow(clippy::too_many_lines)]
fn call_inner(ctx: &Ctx<'_>, name: &str, args: &[V]) -> R {
    let n = args.len();
    // Errors in direct scalar arguments propagate for most functions; handled per function.
    let num1 = |f: &dyn Fn(f64) -> std::result::Result<f64, i32>| -> R {
        let x = n_arg(ctx, args, 0)?;
        let y = f(x)?;
        if y.is_finite() { Ok(V::Num(y)) } else { Err(NUM) }
    };
    match name {
        // ---- math and statistics
        "SUM" => Ok(V::Num(numbers(ctx, args)?.iter().sum())),
        "PRODUCT" => {
            let v = numbers(ctx, args)?;
            Ok(V::Num(if v.is_empty() { 0.0 } else { v.iter().product() }))
        }
        "MAX" => Ok(V::Num(numbers(ctx, args)?.into_iter().reduce(f64::max).unwrap_or(0.0))),
        "MIN" => Ok(V::Num(numbers(ctx, args)?.into_iter().reduce(f64::min).unwrap_or(0.0))),
        "AVERAGE" => {
            let v = numbers(ctx, args)?;
            if v.is_empty() { Err(DIV0) } else { Ok(V::Num(v.iter().sum::<f64>() / v.len() as f64)) }
        }
        "MEDIAN" => {
            let mut v = numbers(ctx, args)?;
            if v.is_empty() {
                return Err(NUM);
            }
            v.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let k = v.len();
            Ok(V::Num(if k % 2 == 1 { v[k / 2] } else { (v[k / 2 - 1] + v[k / 2]) / 2.0 }))
        }
        "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P" | "VAR" | "VAR.S" | "VARP" | "VAR.P" => {
            let v = numbers(ctx, args)?;
            let sample = !name.ends_with('P');
            let s = std_dev(&v, sample)?;
            Ok(V::Num(if name.starts_with("VAR") { s * s } else { s }))
        }
        "COUNT" => {
            let c = flatten(ctx, args)
                .into_iter()
                .filter(|(v, from_ref)| match v {
                    V::Num(_) => true,
                    V::Bool(_) => !from_ref,
                    V::Str(s) => !from_ref && text_number(s).is_some(),
                    _ => false,
                })
                .count();
            Ok(V::Num(c as f64))
        }
        "COUNTA" => {
            let mut c = 0usize;
            for a in args {
                match a {
                    V::Missing => c += 1,
                    other => {
                        c += flatten(ctx, std::slice::from_ref(other)).iter().filter(|(v, _)| *v != V::Empty).count()
                    }
                }
            }
            Ok(V::Num(c as f64))
        }
        "COUNTBLANK" => {
            let V::Ref(a) = arg(args, 0) else { return Err(VALUE) };
            let mut total = 0u64;
            for r in &a.rects {
                let filled =
                    ctx.cells_of(a.sheet, r).iter().filter(|(_, _, v)| !matches!(v, V::Str(s) if s.is_empty())).count();
                total += r.area() - filled as u64;
            }
            Ok(V::Num(total as f64))
        }
        "LARGE" | "SMALL" => {
            let mut v = numbers(ctx, &[arg(args, 0)])?;
            let k = n_arg(ctx, args, 1)?.ceil();
            v.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            if name == "LARGE" {
                v.reverse();
            }
            if k < 1.0 {
                return Err(NUM);
            }
            v.get(k as usize - 1).copied().map(V::Num).ok_or(NUM)
        }
        "RANK" | "RANK.EQ" => {
            let x = n_arg(ctx, args, 0)?;
            let v = numbers(ctx, &[arg(args, 1)])?;
            let asc = opt_n(ctx, args, 2, 0.0)? != 0.0;
            if !v.contains(&x) {
                return Err(NA);
            }
            let better = v.iter().filter(|&&y| if asc { y < x } else { y > x }).count();
            Ok(V::Num((better + 1) as f64))
        }
        "SUMPRODUCT" => {
            let grids: Vec<Grid> = args.iter().map(|a| ctx.grid(a)).collect();
            let Some(first) = grids.first() else { return Err(VALUE) };
            if grids.iter().any(|g| g.rows != first.rows || g.cols != first.cols) {
                return Err(VALUE);
            }
            let mut total = 0.0;
            for i in 0..first.rows * first.cols {
                let mut p = 1.0;
                for g in &grids {
                    match &g.data[i] {
                        V::Num(x) => p *= x,
                        V::Err(e) => return Err(*e),
                        _ => p = 0.0,
                    }
                }
                total += p;
            }
            Ok(V::Num(total))
        }
        "SUBTOTAL" => {
            let f = n_arg(ctx, args, 0)? as i64;
            let skip_hidden = f > 100;
            let v = visible_numbers(ctx, &args[1..], skip_hidden, false)?;
            aggregate(f % 100, &v, ctx, &args[1..], skip_hidden)
        }
        "AGGREGATE" => {
            let f = n_arg(ctx, args, 0)? as i64;
            let opt = n_arg(ctx, args, 1)? as i64;
            let skip_hidden = matches!(opt, 1 | 3 | 5 | 7);
            let skip_err = matches!(opt, 2 | 3 | 6 | 7);
            let v = visible_numbers(ctx, &args[2..], skip_hidden, skip_err)?;
            aggregate(f, &v, ctx, &args[2..], skip_hidden)
        }
        "ABS" => num1(&|x| Ok(x.abs())),
        "SQRT" => num1(&|x| if x < 0.0 { Err(NUM) } else { Ok(x.sqrt()) }),
        "INT" => num1(&|x| Ok(x.floor())),
        "TRUNC" => {
            let x = n_arg(ctx, args, 0)?;
            let d = opt_n(ctx, args, 1, 0.0)?;
            Ok(V::Num(round_digits(x, d, 2)))
        }
        "SIGN" => num1(&|x| {
            Ok(if x > 0.0 {
                1.0
            } else if x < 0.0 {
                -1.0
            } else {
                0.0
            })
        }),
        "EXP" => num1(&|x| Ok(x.exp())),
        "LN" => num1(&|x| if x <= 0.0 { Err(NUM) } else { Ok(x.ln()) }),
        "LOG10" => num1(&|x| if x <= 0.0 { Err(NUM) } else { Ok(x.log10()) }),
        "LOG" => {
            let x = n_arg(ctx, args, 0)?;
            let b = opt_n(ctx, args, 1, 10.0)?;
            if x <= 0.0 || b <= 0.0 || b == 1.0 {
                return Err(if b == 1.0 { DIV0 } else { NUM });
            }
            Ok(V::Num(x.ln() / b.ln()))
        }
        "PI" => Ok(V::Num(std::f64::consts::PI)),
        "POWER" => binary_ok(binary_scalar("^", sc(ctx, args, 0), sc(ctx, args, 1))),
        "MOD" => {
            let x = n_arg(ctx, args, 0)?;
            let y = n_arg(ctx, args, 1)?;
            if y == 0.0 {
                return Err(DIV0);
            }
            Ok(V::Num(x - y * (x / y).floor()))
        }
        "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
            let x = n_arg(ctx, args, 0)?;
            let d = n_arg(ctx, args, 1)?;
            let mode = match name {
                "ROUND" => 0,
                "ROUNDUP" => 1,
                _ => 2,
            };
            Ok(V::Num(round_digits(x, d, mode)))
        }
        "CEILING" | "FLOOR" | "CEILING.MATH" | "FLOOR.MATH" => {
            let x = n_arg(ctx, args, 0)?;
            let s = opt_n(ctx, args, 1, 1.0)?;
            if s == 0.0 {
                return Ok(V::Num(0.0));
            }
            if x > 0.0 && s < 0.0 {
                return Err(NUM);
            }
            let q = x / s;
            let q: f64 = format!("{q:.12e}").parse().unwrap_or(q);
            let r = if name.starts_with("CEILING") { q.ceil() } else { q.floor() };
            Ok(V::Num(r * s))
        }
        "EVEN" | "ODD" => {
            let x = n_arg(ctx, args, 0)?;
            let a = x.abs().ceil();
            let r = if name == "EVEN" {
                if a % 2.0 == 0.0 { a } else { a + 1.0 }
            } else if a % 2.0 == 1.0 {
                a
            } else {
                a + 1.0
            };
            Ok(V::Num(r.copysign(if x == 0.0 { 1.0 } else { x })))
        }
        "FACT" => {
            let x = n_arg(ctx, args, 0)?.trunc();
            if x < 0.0 {
                return Err(NUM);
            }
            Ok(V::Num((1..=x as u64).map(|k| k as f64).product()))
        }
        "COMBIN" => {
            let nn = n_arg(ctx, args, 0)?.trunc();
            let k = n_arg(ctx, args, 1)?.trunc();
            if k < 0.0 || nn < k {
                return Err(NUM);
            }
            let mut r = 1.0;
            for i in 0..k as u64 {
                r = r * (nn - i as f64) / (i as f64 + 1.0);
            }
            Ok(V::Num(r.round()))
        }
        "PMT" => {
            let rate = n_arg(ctx, args, 0)?;
            let nper = n_arg(ctx, args, 1)?;
            let pv = n_arg(ctx, args, 2)?;
            let fv = opt_n(ctx, args, 3, 0.0)?;
            let ty = opt_n(ctx, args, 4, 0.0)?;
            let p = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let f = (1.0 + rate).powf(nper);
                -(rate * (pv * f + fv)) / ((1.0 + rate * ty) * (f - 1.0))
            };
            Ok(V::Num(p))
        }
        "RAND" => Ok(V::Num(0.5)),
        // ---- conditional aggregation
        "SUMIF" => {
            let rng = arg(args, 0);
            let (rows, cols) = ctx.dims(&rng);
            let sum_rng = match arg(args, 2) {
                V::Missing => rng.clone(),
                s => reshape(&s, rows, cols),
            };
            Ok(V::Num(multi_criteria(ctx, &[(rng, arg(args, 1))], Some(&sum_rng))?.0))
        }
        "AVERAGEIF" => {
            let rng = arg(args, 0);
            let (rows, cols) = ctx.dims(&rng);
            let avg = match arg(args, 2) {
                V::Missing => rng.clone(),
                s => reshape(&s, rows, cols),
            };
            let (sum, _, k) = multi_criteria(ctx, &[(rng, arg(args, 1))], Some(&avg))?;
            if k == 0.0 { Err(DIV0) } else { Ok(V::Num(sum / k)) }
        }
        "COUNTIF" => {
            if !matches!(arg(args, 0), V::Ref(_) | V::Arr(_)) {
                return Err(VALUE);
            }
            Ok(V::Num(multi_criteria(ctx, &[(arg(args, 0), arg(args, 1))], None)?.1))
        }
        "MAXIFS" | "MINIFS" => {
            let target = arg(args, 0);
            let rest = args.get(1..).unwrap_or(&[]);
            if rest.is_empty() || rest.len() % 2 != 0 {
                return Err(VALUE);
            }
            let (rows, cols) = ctx.dims(&target);
            let crits: Vec<(V, Crit)> =
                rest.chunks(2).map(|p| (p[0].clone(), parse_crit(&ctx.scalar(p[1].clone())))).collect();
            if crits.iter().any(|(r, _)| ctx.dims(r) != (rows, cols)) {
                return Err(VALUE);
            }
            // Only the stored part of a whole-column reference can match.
            let (lr, lc) = match &target {
                V::Ref(a) => ctx.clip(a.sheet, &a.rects[0]).map_or((0, 0), |c| (c.rows() as usize, c.cols() as usize)),
                _ => (rows, cols),
            };
            let mut best: Option<f64> = None;
            for r in 0..lr {
                for c in 0..lc {
                    if let V::Num(x) = value_at(ctx, &target, r, c)
                        && crits.iter().all(|(rng, cr)| crit_matches(cr, &value_at(ctx, rng, r, c)))
                    {
                        best = Some(match best {
                            Some(b) if name == "MAXIFS" => b.max(x),
                            Some(b) => b.min(x),
                            None => x,
                        });
                    }
                }
            }
            Ok(V::Num(best.unwrap_or(0.0)))
        }
        "COUNTIFS" | "SUMIFS" | "AVERAGEIFS" => {
            let (rest, sum) =
                if name == "COUNTIFS" { (args, None) } else { (args.get(1..).unwrap_or(&[]), args.first()) };
            if rest.is_empty() || rest.len() % 2 != 0 {
                return Err(VALUE);
            }
            let pairs: Vec<(V, V)> = rest.chunks(2).map(|p| (p[0].clone(), p[1].clone())).collect();
            let (s, c, k) = multi_criteria(ctx, &pairs, sum)?;
            match name {
                "COUNTIFS" => Ok(V::Num(c)),
                "SUMIFS" => Ok(V::Num(s)),
                _ => {
                    if k == 0.0 {
                        Err(DIV0)
                    } else {
                        Ok(V::Num(s / k))
                    }
                }
            }
        }
        // ---- logical
        "IF" => {
            if n == 0 {
                return Err(VALUE);
            }
            let c = bool_of(&sc(ctx, args, 0))?;
            let pick = if c { 1 } else { 2 };
            Ok(match arg(args, pick) {
                V::Missing if pick == 1 && n >= 2 => V::Num(0.0),
                V::Missing if n > pick => V::Num(0.0),
                V::Missing => V::Bool(c),
                v => v,
            })
        }
        "IFERROR" => {
            let v = sc(ctx, args, 0);
            Ok(if v.is_err() { arg(args, 1) } else { v })
        }
        "IFNA" => {
            let v = sc(ctx, args, 0);
            Ok(if v == V::Err(NA) { arg(args, 1) } else { v })
        }
        "AND" | "OR" | "XOR" => {
            let mut vals = Vec::new();
            for (v, from_ref) in flatten(ctx, args) {
                match v {
                    V::Err(e) => return Err(e),
                    V::Str(_) if from_ref => {}
                    V::Empty => {}
                    other => vals.push(bool_of(&other)?),
                }
            }
            if vals.is_empty() {
                return Err(VALUE);
            }
            Ok(V::Bool(match name {
                "AND" => vals.iter().all(|b| *b),
                "OR" => vals.iter().any(|b| *b),
                _ => vals.iter().filter(|b| **b).count() % 2 == 1,
            }))
        }
        "NOT" => Ok(V::Bool(!bool_of(&sc(ctx, args, 0))?)),
        "TRUE" => Ok(V::Bool(true)),
        "FALSE" => Ok(V::Bool(false)),
        "CHOOSE" => {
            let i = n_arg(ctx, args, 0)?.trunc();
            if i < 1.0 || i as usize >= n {
                return Err(VALUE);
            }
            Ok(arg(args, i as usize))
        }
        // ---- information
        "ISERROR" => Ok(V::Bool(sc(ctx, args, 0).is_err())),
        "ISERR" => Ok(V::Bool(matches!(sc(ctx, args, 0), V::Err(e) if e != NA))),
        "ISNA" => Ok(V::Bool(sc(ctx, args, 0) == V::Err(NA))),
        "ISNUMBER" => Ok(V::Bool(matches!(sc(ctx, args, 0), V::Num(_)))),
        "ISTEXT" => Ok(V::Bool(matches!(sc(ctx, args, 0), V::Str(_)))),
        "ISNONTEXT" => Ok(V::Bool(!matches!(sc(ctx, args, 0), V::Str(_)))),
        "ISLOGICAL" => Ok(V::Bool(matches!(sc(ctx, args, 0), V::Bool(_)))),
        "ISBLANK" => Ok(V::Bool(matches!(sc(ctx, args, 0), V::Empty))),
        "NA" => Err(NA),
        "N" => Ok(match sc(ctx, args, 0) {
            V::Num(x) => V::Num(x),
            V::Bool(b) => V::Num(if b { 1.0 } else { 0.0 }),
            V::Err(e) => V::Err(e),
            _ => V::Num(0.0),
        }),
        "T" => Ok(match sc(ctx, args, 0) {
            V::Str(s) => V::Str(s),
            V::Err(e) => V::Err(e),
            _ => V::Str(String::new()),
        }),
        // ---- text
        "LEN" => Ok(V::Num(t_arg(ctx, args, 0)?.chars().count() as f64)),
        "LEFT" | "RIGHT" => {
            let s = t_arg(ctx, args, 0)?;
            let k = opt_n(ctx, args, 1, 1.0)?;
            if k < 0.0 {
                return Err(VALUE);
            }
            let chars: Vec<char> = s.chars().collect();
            let k = (k as usize).min(chars.len());
            Ok(V::Str(if name == "LEFT" {
                chars[..k].iter().collect()
            } else {
                chars[chars.len() - k..].iter().collect()
            }))
        }
        "MID" => {
            let s = t_arg(ctx, args, 0)?;
            let start = n_arg(ctx, args, 1)?;
            let k = n_arg(ctx, args, 2)?;
            if start < 1.0 || k < 0.0 {
                return Err(VALUE);
            }
            Ok(V::Str(s.chars().skip(start as usize - 1).take(k as usize).collect()))
        }
        "UPPER" => Ok(V::Str(t_arg(ctx, args, 0)?.to_uppercase())),
        "LOWER" => Ok(V::Str(t_arg(ctx, args, 0)?.to_lowercase())),
        "PROPER" => Ok(V::Str(proper(&t_arg(ctx, args, 0)?))),
        "TRIM" => Ok(V::Str(t_arg(ctx, args, 0)?.split(' ').filter(|p| !p.is_empty()).collect::<Vec<_>>().join(" "))),
        "CLEAN" => Ok(V::Str(t_arg(ctx, args, 0)?.chars().filter(|c| (*c as u32) >= 32).collect())),
        "CONCATENATE" => {
            let mut s = String::new();
            for i in 0..n {
                s.push_str(&t_arg(ctx, args, i)?);
            }
            Ok(V::Str(s))
        }
        "CONCAT" => {
            let mut s = String::new();
            for (v, _) in flatten(ctx, args) {
                s.push_str(&text_of(&v)?);
            }
            Ok(V::Str(s))
        }
        "TEXTJOIN" => {
            let delim = t_arg(ctx, args, 0)?;
            let skip = b_arg(ctx, args, 1, true)?;
            let mut parts = Vec::new();
            for (v, _) in flatten(ctx, args.get(2..).unwrap_or(&[])) {
                let t = text_of(&v)?;
                if !(skip && t.is_empty()) {
                    parts.push(t);
                }
            }
            Ok(V::Str(parts.join(&delim)))
        }
        "REPT" => {
            let s = t_arg(ctx, args, 0)?;
            let k = n_arg(ctx, args, 1)?;
            if k < 0.0 {
                return Err(VALUE);
            }
            Ok(V::Str(s.repeat(k as usize)))
        }
        "EXACT" => Ok(V::Bool(t_arg(ctx, args, 0)? == t_arg(ctx, args, 1)?)),
        "SUBSTITUTE" => {
            let s = t_arg(ctx, args, 0)?;
            let old = t_arg(ctx, args, 1)?;
            let new = t_arg(ctx, args, 2)?;
            if old.is_empty() {
                return Ok(V::Str(s));
            }
            match arg(args, 3) {
                V::Missing => Ok(V::Str(s.replace(&old, &new))),
                _ => {
                    let k = n_arg(ctx, args, 3)?;
                    if k < 1.0 {
                        return Err(VALUE);
                    }
                    let mut out = String::new();
                    let mut rest = s.as_str();
                    let mut seen = 0;
                    while let Some(p) = rest.find(&old) {
                        seen += 1;
                        out.push_str(&rest[..p]);
                        out.push_str(if seen == k as usize { &new } else { &old });
                        rest = &rest[p + old.len()..];
                    }
                    out.push_str(rest);
                    Ok(V::Str(out))
                }
            }
        }
        "REPLACE" => {
            let s: Vec<char> = t_arg(ctx, args, 0)?.chars().collect();
            let start = n_arg(ctx, args, 1)?;
            let k = n_arg(ctx, args, 2)?;
            let new = t_arg(ctx, args, 3)?;
            if start < 1.0 || k < 0.0 {
                return Err(VALUE);
            }
            let a = (start as usize - 1).min(s.len());
            let b = (a + k as usize).min(s.len());
            let mut out: String = s[..a].iter().collect();
            out.push_str(&new);
            out.extend(&s[b..]);
            Ok(V::Str(out))
        }
        "FIND" | "SEARCH" => {
            let needle = t_arg(ctx, args, 0)?;
            let hay = t_arg(ctx, args, 1)?;
            let start = opt_n(ctx, args, 2, 1.0)?;
            let hay_chars: Vec<char> = hay.chars().collect();
            if start < 1.0 || start as usize > hay_chars.len() + 1 {
                return Err(VALUE);
            }
            let from = start as usize - 1;
            for i in from..=hay_chars.len() {
                let rest: String = hay_chars[i..].iter().collect();
                let hit = if name == "FIND" {
                    rest.starts_with(&needle)
                } else {
                    let pat = format!("{needle}*");
                    wildcard_match(&pat, &rest, true, false)
                };
                if hit {
                    return Ok(V::Num((i + 1) as f64));
                }
            }
            Err(VALUE)
        }
        "VALUE" => {
            let v = sc(ctx, args, 0);
            match v {
                V::Num(x) => Ok(V::Num(x)),
                V::Empty => Ok(V::Num(0.0)),
                other => text_number(&text_of(&other)?).map(V::Num).ok_or(VALUE),
            }
        }
        "TEXT" => {
            let v = sc(ctx, args, 0);
            let fmt = t_arg(ctx, args, 1)?;
            match v {
                V::Err(e) => Err(e),
                V::Str(s) => match text_number(&s) {
                    Some(x) => Ok(V::Str(format_value(&Value::Double(x), &fmt))),
                    None => Ok(V::Str(format_value(&Value::from(s.as_str()), &fmt))),
                },
                other => Ok(V::Str(format_value(&to_value(&other), &fmt))),
            }
        }
        "DOLLAR" => {
            let x = n_arg(ctx, args, 0)?;
            let d = opt_n(ctx, args, 1, 2.0)?.trunc() as i32;
            let r = round_digits(x, f64::from(d), 0);
            let body = if d > 0 { format!("#,##0.{}", "0".repeat(d as usize)) } else { "#,##0".to_owned() };
            Ok(V::Str(format_value(&Value::Double(r), &format!("$${body};($${body})").replace("$$", "$"))))
        }
        "FIXED" => {
            let x = n_arg(ctx, args, 0)?;
            let d = opt_n(ctx, args, 1, 2.0)?;
            let nc = b_arg(ctx, args, 2, false)?;
            Ok(V::Str(fixed_text(x, d, nc)))
        }
        "CHAR" => {
            let k = n_arg(ctx, args, 0)?;
            if !(1.0..=255.0).contains(&k) {
                return Err(VALUE);
            }
            Ok(V::Str(char::from(k as u8).to_string()))
        }
        "CODE" => {
            let s = t_arg(ctx, args, 0)?;
            s.chars().next().map(|c| V::Num(f64::from(c as u32))).ok_or(VALUE)
        }
        // ---- dates
        "DATE" => {
            let y = n_arg(ctx, args, 0)?.trunc() as i64;
            let m = n_arg(ctx, args, 1)?.trunc() as i64;
            let d = n_arg(ctx, args, 2)?.trunc() as i64;
            let s = ymd_serial(y, m, d);
            if s < 0.0 { Err(NUM) } else { Ok(V::Num(s)) }
        }
        "TIME" => {
            let h = n_arg(ctx, args, 0)?;
            let m = n_arg(ctx, args, 1)?;
            let s = n_arg(ctx, args, 2)?;
            let t = (h * 3600.0 + m * 60.0 + s).trunc() / 86_400.0;
            Ok(V::Num(t - t.floor()))
        }
        "YEAR" | "MONTH" | "DAY" => {
            let x = n_arg(ctx, args, 0)?;
            if x < 0.0 {
                return Err(NUM);
            }
            let (y, m, d) = serial_ymd(x);
            Ok(V::Num(match name {
                "YEAR" => y as f64,
                "MONTH" => f64::from(m),
                _ => f64::from(d),
            }))
        }
        "HOUR" | "MINUTE" | "SECOND" => {
            let x = n_arg(ctx, args, 0)?;
            let (h, m, s) = serial_to_hms(x);
            Ok(V::Num(f64::from(match name {
                "HOUR" => h,
                "MINUTE" => m,
                _ => s,
            })))
        }
        "WEEKDAY" => {
            let x = n_arg(ctx, args, 0)?;
            let ty = opt_n(ctx, args, 1, 1.0)? as i64;
            let w = weekday_of(x); // 1 = Sunday
            Ok(V::Num(match ty {
                2 => ((w + 5) % 7 + 1) as f64,
                3 => ((w + 5) % 7) as f64,
                _ => w as f64,
            }))
        }
        "WEEKNUM" => {
            let x = n_arg(ctx, args, 0)?;
            let (y, _, _) = serial_ymd(x);
            let jan1 = ymd_serial(y, 1, 1);
            let offset = weekday_of(jan1) - 1;
            Ok(V::Num(((x.floor() - jan1 + offset as f64) / 7.0).floor() + 1.0))
        }
        "EDATE" | "EOMONTH" => {
            let x = n_arg(ctx, args, 0)?;
            let k = n_arg(ctx, args, 1)?.trunc() as i64;
            let (y, m, d) = serial_ymd(x);
            let total = i64::from(m) - 1 + k;
            let (ny, nm) = (y + total.div_euclid(12), total.rem_euclid(12) + 1);
            let last = ymd_serial(ny, nm + 1, 1) - 1.0;
            let (_, _, last_day) = serial_ymd(last);
            Ok(V::Num(if name == "EOMONTH" { last } else { ymd_serial(ny, nm, i64::from(d.min(last_day))) }))
        }
        "DATEDIF" => {
            let a = n_arg(ctx, args, 0)?.floor();
            let b = n_arg(ctx, args, 1)?.floor();
            let unit = t_arg(ctx, args, 2)?.to_ascii_uppercase();
            if a > b {
                return Err(NUM);
            }
            let (y1, m1, d1) = serial_ymd(a);
            let (y2, m2, d2) = serial_ymd(b);
            let months = (y2 - y1) * 12 + i64::from(m2) - i64::from(m1) - i64::from(d2 < d1);
            Ok(V::Num(match unit.as_str() {
                "D" => b - a,
                "M" => months as f64,
                "Y" => (months / 12) as f64,
                "YM" => (months % 12) as f64,
                _ => return Err(NUM),
            }))
        }
        "NETWORKDAYS" => {
            let a = n_arg(ctx, args, 0)?.floor();
            let b = n_arg(ctx, args, 1)?.floor();
            let (lo, hi, sign) = if a <= b { (a, b, 1.0) } else { (b, a, -1.0) };
            let mut k = 0.0;
            let mut d = lo;
            while d <= hi {
                let w = weekday_of(d);
                if w != 1 && w != 7 {
                    k += 1.0;
                }
                d += 1.0;
            }
            Ok(V::Num(k * sign))
        }
        "DAYS" => Ok(V::Num(n_arg(ctx, args, 0)?.floor() - n_arg(ctx, args, 1)?.floor())),
        "TODAY" => Ok(V::Num(ctx.now.floor())),
        "NOW" => Ok(V::Num(ctx.now)),
        // ---- lookup and reference
        "MATCH" => {
            let lookup = sc(ctx, args, 0);
            if lookup.is_err() {
                return Err(NA);
            }
            let mode = opt_n(ctx, args, 2, 1.0)?;
            let g = ctx.grid(&arg(args, 1));
            let v = vector(&g).ok_or(NA)?;
            match_in(
                &v,
                &lookup,
                if mode > 0.0 {
                    1
                } else if mode < 0.0 {
                    -1
                } else {
                    0
                },
            )
            .map(|i| V::Num((i + 1) as f64))
            .ok_or(NA)
        }
        "VLOOKUP" | "HLOOKUP" => {
            let lookup = sc(ctx, args, 0);
            if let V::Err(e) = lookup {
                return Err(e);
            }
            let table = arg(args, 1);
            let idx = n_arg(ctx, args, 2)?.trunc();
            let approx = b_arg(ctx, args, 3, true)?;
            let (rows, cols) = ctx.dims(&table);
            let vertical = name == "VLOOKUP";
            let width = if vertical { cols } else { rows };
            if idx < 1.0 {
                return Err(VALUE);
            }
            if idx as usize > width {
                return Err(REF);
            }
            let g = ctx.grid(&table);
            let keys: Vec<V> = if vertical {
                (0..g.rows).map(|r| g.get(r, 0)).collect()
            } else {
                (0..g.cols).map(|c| g.get(0, c)).collect()
            };
            let i = match_in(&keys, &lookup, if approx { 1 } else { 0 }).ok_or(NA)?;
            let k = idx as usize - 1;
            Ok(if vertical { value_at(ctx, &table, i, k) } else { value_at(ctx, &table, k, i) })
        }
        "INDEX" => {
            let src = arg(args, 0);
            let (rows, cols) = ctx.dims(&src);
            let r = opt_n(ctx, args, 1, 0.0)?.trunc();
            let c = match arg(args, 2) {
                V::Missing => None,
                _ => Some(n_arg(ctx, args, 2)?.trunc()),
            };
            let (r, c) = match c {
                Some(c) => (r, c),
                None if rows == 1 => (1.0, r),
                None => (r, if cols == 1 { 1.0 } else { 0.0 }),
            };
            if r < 0.0 || c < 0.0 || r as usize > rows || c as usize > cols {
                return Err(REF);
            }
            let (r, c) = (r as usize, c as usize);
            match &src {
                V::Ref(a) => {
                    let base = a.rects[0];
                    let rect = match (r, c) {
                        (0, 0) => base,
                        (0, c) => Rect::new(base.r1, base.c1 + c as u32 - 1, base.r2, base.c1 + c as u32 - 1),
                        (r, 0) => Rect::new(base.r1 + r as u32 - 1, base.c1, base.r1 + r as u32 - 1, base.c2),
                        (r, c) => Rect::cell(base.r1 + r as u32 - 1, base.c1 + c as u32 - 1),
                    };
                    Ok(V::Ref(Area { sheet: a.sheet, rects: vec![rect] }))
                }
                other => {
                    let g = ctx.grid(other);
                    Ok(match (r, c) {
                        (0, 0) => V::Arr(g),
                        (0, c) => grid_value(Grid {
                            rows: g.rows,
                            cols: 1,
                            data: (0..g.rows).map(|i| g.get(i, c - 1)).collect(),
                        }),
                        (r, 0) => grid_value(Grid {
                            rows: 1,
                            cols: g.cols,
                            data: (0..g.cols).map(|j| g.get(r - 1, j)).collect(),
                        }),
                        (r, c) => g.get(r - 1, c - 1),
                    })
                }
            }
        }
        "ROW" | "COLUMN" => match arg(args, 0) {
            V::Ref(a) => Ok(V::Num(f64::from(if name == "ROW" { a.rects[0].r1 } else { a.rects[0].c1 }))),
            V::Missing => {
                let (r, c) = ctx.cell.ok_or(VALUE)?;
                Ok(V::Num(f64::from(if name == "ROW" { r } else { c })))
            }
            _ => Err(VALUE),
        },
        "ROWS" | "COLUMNS" => {
            let (r, c) = match arg(args, 0) {
                V::Ref(_) | V::Arr(_) => ctx.dims(&arg(args, 0)),
                V::Err(e) => return Err(e),
                _ => (1, 1),
            };
            Ok(V::Num(if name == "ROWS" { r } else { c } as f64))
        }
        "OFFSET" => {
            let V::Ref(a) = arg(args, 0) else { return Err(VALUE) };
            let base = a.rects[0];
            let dr = n_arg(ctx, args, 1)?.trunc() as i64;
            let dc = n_arg(ctx, args, 2)?.trunc() as i64;
            let h = opt_n(ctx, args, 3, f64::from(base.rows()))?.trunc() as i64;
            let w = opt_n(ctx, args, 4, f64::from(base.cols()))?.trunc() as i64;
            let r1 = i64::from(base.r1) + dr;
            let c1 = i64::from(base.c1) + dc;
            let (r2, c2) = (r1 + h - 1, c1 + w - 1);
            if h < 1 || w < 1 || r1 < 1 || c1 < 1 || r2 > i64::from(MAX_ROWS) || c2 > i64::from(MAX_COLS) {
                return Err(REF);
            }
            Ok(V::Ref(Area { sheet: a.sheet, rects: vec![Rect::new(r1 as u32, c1 as u32, r2 as u32, c2 as u32)] }))
        }
        "INDIRECT" => {
            let t = t_arg(ctx, args, 0)?;
            let p = parse_reference(&t).ok_or(REF)?;
            let sheet = match &p.sheet {
                Some(s) => ctx.sheet_by_name(s).ok_or(REF)?,
                None => ctx.sheet,
            };
            Ok(V::Ref(Area { sheet, rects: p.areas }))
        }
        "ADDRESS" => {
            let r = n_arg(ctx, args, 0)?.trunc() as u32;
            let c = n_arg(ctx, args, 1)?.trunc() as u32;
            let abs = opt_n(ctx, args, 2, 1.0)? as i64;
            if r < 1 || c < 1 || r > MAX_ROWS || c > MAX_COLS {
                return Err(VALUE);
            }
            let (ra, ca) = match abs {
                2 => ("$", ""),
                3 => ("", "$"),
                4 => ("", ""),
                _ => ("$", "$"),
            };
            Ok(V::Str(format!("{ca}{}{ra}{r}", column_letters(c))))
        }
        "TRANSPOSE" => {
            let g = ctx.grid(&arg(args, 0));
            let mut data = Vec::with_capacity(g.rows * g.cols);
            for c in 0..g.cols {
                for r in 0..g.rows {
                    data.push(g.get(r, c));
                }
            }
            Ok(V::Arr(Grid { rows: g.cols, cols: g.rows, data }))
        }
        _ => Err(-1),
    }
}

fn binary_ok(v: V) -> R {
    match v {
        V::Err(e) => Err(e),
        other => Ok(other),
    }
}

fn aggregate(f: i64, v: &[f64], ctx: &Ctx<'_>, refs: &[V], skip_hidden: bool) -> R {
    let k = v.len() as f64;
    Ok(V::Num(match f {
        1 => {
            if v.is_empty() {
                return Err(DIV0);
            }
            v.iter().sum::<f64>() / k
        }
        2 => k,
        3 => {
            let mut c = 0.0;
            for a in refs {
                if let V::Ref(area) = a {
                    for r in &area.rects {
                        for (row, _, _) in ctx.cells_of(area.sheet, r) {
                            if !(ctx.book.row_filtered(area.sheet, row)
                                || skip_hidden && ctx.book.row_hidden(area.sheet, row))
                            {
                                c += 1.0;
                            }
                        }
                    }
                }
            }
            c
        }
        4 => v.iter().copied().reduce(f64::max).unwrap_or(0.0),
        5 => v.iter().copied().reduce(f64::min).unwrap_or(0.0),
        6 => {
            if v.is_empty() {
                0.0
            } else {
                v.iter().product()
            }
        }
        7 => std_dev(v, true)?,
        8 => std_dev(v, false)?,
        9 => v.iter().sum(),
        10 => std_dev(v, true)?.powi(2),
        11 => std_dev(v, false)?.powi(2),
        _ => return Err(VALUE),
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::MemoryWorkbook;

    fn book() -> MemoryWorkbook {
        let mut b = MemoryWorkbook::new();
        let id = b.sheet_ids()[0];
        b.set_cell_value(id, 1, 1, Value::Double(3.0)).ok();
        b.set_cell_value(id, 2, 1, Value::Double(4.0)).ok();
        b.set_cell_value(id, 3, 1, Value::from("x")).ok();
        b
    }

    fn ev(f: &str) -> Value {
        let b = book();
        let id = b.sheet_ids()[0];
        let ctx = Ctx { book: &b, sheet: id, cell: Some((10, 10)), now: 46000.0 };
        to_value(&ctx.scalar(eval_text(f, &ctx)))
    }

    fn num(v: Value) -> f64 {
        v.to_f64().unwrap_or(f64::NAN)
    }

    #[test]
    fn arithmetic_and_refs() {
        assert_eq!(num(ev("=1+2*3")), 7.0);
        assert_eq!(num(ev("=(1+2)*3^2")), 27.0);
        assert_eq!(num(ev("=-2^2")), 4.0);
        assert_eq!(num(ev("=A1*10")), 30.0);
        assert_eq!(num(ev("=SUM(A1:A3)")), 7.0);
        assert_eq!(ev("=A3&\"y\"").to_text().unwrap_or_default(), "xy");
        assert!(matches!(ev("=1/0"), Value::Error(2007)));
        assert!(matches!(ev("=FOO(1)"), Value::Error(2029)));
        assert_eq!(ev("=IF(A1>2,\"big\",\"small\")").to_text().unwrap_or_default(), "big");
        assert_eq!(num(ev("=SUM({1,2;3,4})")), 10.0);
        assert_eq!(num(ev("=VLOOKUP(4,A1:A3,1,FALSE)")), 4.0);
        assert_eq!(num(ev("=7%")), 0.07);
        assert_eq!(num(ev("=DATE(2026,10,14)")), 46309.0);
        assert_eq!(num(ev("=YEAR(46309)")), 2026.0);
    }
}
