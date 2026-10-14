//! Formula text rewriting: moving relative references (copy/fill), and converting between A1 and
//! R1C1 notation. The object model does not evaluate formulas (the workbook's calculation engine
//! does); it only stores text, so these work on references without parsing whole expressions.

use crate::address::{MAX_COLS, MAX_ROWS, column_letters, column_number};

/// A reference coordinate: value and whether it is absolute (`$`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Coord {
    n: u32,
    abs: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RefTok {
    Cell { col: Coord, row: Coord },
    Cols(Coord, Coord),
    Rows(Coord, Coord),
}

enum Seg {
    Text(String),
    Ref(RefTok),
}

fn is_ident(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.'
}

/// Reads `$?letters` at `i`; returns (column, abs, next index).
fn read_col(ch: &[char], i: usize) -> Option<(Coord, usize)> {
    let mut j = i;
    let abs = ch.get(j) == Some(&'$');
    if abs {
        j += 1;
    }
    let start = j;
    while j < ch.len() && ch[j].is_ascii_alphabetic() && j - start < 4 {
        j += 1;
    }
    let letters: String = ch[start..j].iter().collect();
    let n = column_number(&letters)?;
    Some((Coord { n, abs }, j))
}

fn read_row(ch: &[char], i: usize) -> Option<(Coord, usize)> {
    let mut j = i;
    let abs = ch.get(j) == Some(&'$');
    if abs {
        j += 1;
    }
    let start = j;
    while j < ch.len() && ch[j].is_ascii_digit() {
        j += 1;
    }
    if j == start || j - start > 7 {
        return None;
    }
    let digits: String = ch[start..j].iter().collect();
    let n: u32 = digits.parse().ok()?;
    (1..=MAX_ROWS).contains(&n).then_some((Coord { n, abs }, j))
}

/// True when position `j` ends a reference token (not followed by more of an identifier, a call
/// parenthesis or a sheet separator).
fn ends_ref(ch: &[char], j: usize) -> bool {
    match ch.get(j) {
        None => true,
        Some(&c) => !(is_ident(c) || c == '(' || c == '!' || c == '$'),
    }
}

fn try_ref_a1(ch: &[char], i: usize) -> Option<(RefTok, usize)> {
    // Cell: $?col$?row
    if let Some((col, j)) = read_col(ch, i)
        && let Some((row, k)) = read_row(ch, j)
        && ends_ref(ch, k)
    {
        return Some((RefTok::Cell { col, row }, k));
    }
    // Column range: $?col:$?col
    if let Some((a, j)) = read_col(ch, i)
        && ch.get(j) == Some(&':')
        && let Some((b, k)) = read_col(ch, j + 1)
        && ends_ref(ch, k)
        && !ch.get(k).is_some_and(|c| c.is_ascii_digit())
    {
        return Some((RefTok::Cols(a, b), k));
    }
    // Row range: $?row:$?row
    if let Some((a, j)) = read_row(ch, i)
        && ch.get(j) == Some(&':')
        && let Some((b, k)) = read_row(ch, j + 1)
        && ends_ref(ch, k)
    {
        return Some((RefTok::Rows(a, b), k));
    }
    None
}

/// Splits formula text into plain text and A1 references.
fn scan_a1(formula: &str) -> Vec<Seg> {
    let ch: Vec<char> = formula.chars().collect();
    let mut segs = Vec::new();
    let mut text = String::new();
    let mut i = 0;
    while i < ch.len() {
        let c = ch[i];
        if c == '"' || c == '\'' {
            // String literal or quoted sheet name, with doubled-quote escapes.
            text.push(c);
            i += 1;
            while i < ch.len() {
                text.push(ch[i]);
                if ch[i] == c {
                    if ch.get(i + 1) == Some(&c) {
                        text.push(c);
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        let boundary = i == 0 || !is_ident(ch[i - 1]);
        if boundary && (c == '$' || c.is_ascii_alphanumeric()) {
            if let Some((tok, j)) = try_ref_a1(&ch, i) {
                segs.push(Seg::Text(std::mem::take(&mut text)));
                segs.push(Seg::Ref(tok));
                i = j;
                continue;
            }
            // Copy the whole identifier or number so its tail is not mistaken for a reference.
            while i < ch.len() && (is_ident(ch[i]) || ch[i] == '$') {
                text.push(ch[i]);
                i += 1;
            }
            continue;
        }
        text.push(c);
        i += 1;
    }
    segs.push(Seg::Text(text));
    segs
}

fn shift_coord(c: Coord, d: i64, max: u32) -> Option<Coord> {
    if c.abs {
        return Some(c);
    }
    let n = i64::from(c.n) + d;
    (n >= 1 && n <= i64::from(max)).then_some(Coord { n: n as u32, abs: false })
}

fn a1_col(c: Coord) -> String {
    format!("{}{}", if c.abs { "$" } else { "" }, column_letters(c.n))
}

fn a1_row(c: Coord) -> String {
    format!("{}{}", if c.abs { "$" } else { "" }, c.n)
}

fn a1_text(tok: &RefTok) -> String {
    match tok {
        RefTok::Cell { col, row } => format!("{}{}", a1_col(*col), a1_row(*row)),
        RefTok::Cols(a, b) => format!("{}:{}", a1_col(*a), a1_col(*b)),
        RefTok::Rows(a, b) => format!("{}:{}", a1_row(*a), a1_row(*b)),
    }
}

/// Moves the relative references of an A1 formula by (dr, dc), as copying a cell does.
/// References pushed off the sheet become `#REF!`.
pub fn shift_formula(formula: &str, dr: i64, dc: i64) -> String {
    let mut out = String::new();
    for seg in scan_a1(formula) {
        match seg {
            Seg::Text(t) => out.push_str(&t),
            Seg::Ref(tok) => {
                let moved = match tok {
                    RefTok::Cell { col, row } => shift_coord(col, dc, MAX_COLS)
                        .zip(shift_coord(row, dr, MAX_ROWS))
                        .map(|(col, row)| RefTok::Cell { col, row }),
                    RefTok::Cols(a, b) => {
                        shift_coord(a, dc, MAX_COLS).zip(shift_coord(b, dc, MAX_COLS)).map(|(a, b)| RefTok::Cols(a, b))
                    }
                    RefTok::Rows(a, b) => {
                        shift_coord(a, dr, MAX_ROWS).zip(shift_coord(b, dr, MAX_ROWS)).map(|(a, b)| RefTok::Rows(a, b))
                    }
                };
                match moved {
                    Some(t) => out.push_str(&a1_text(&t)),
                    None => out.push_str("#REF!"),
                }
            }
        }
    }
    out
}

fn r1c1_part(letter: char, c: Coord, origin: u32) -> String {
    if c.abs {
        format!("{letter}{}", c.n)
    } else {
        let d = i64::from(c.n) - i64::from(origin);
        if d == 0 { letter.to_string() } else { format!("{letter}[{d}]") }
    }
}

/// Converts an A1 formula to R1C1 notation as seen from cell (row, col).
pub fn a1_to_r1c1(formula: &str, row: u32, col: u32) -> String {
    let mut out = String::new();
    for seg in scan_a1(formula) {
        match seg {
            Seg::Text(t) => out.push_str(&t),
            Seg::Ref(RefTok::Cell { col: c, row: r }) => {
                out.push_str(&r1c1_part('R', r, row));
                out.push_str(&r1c1_part('C', c, col));
            }
            Seg::Ref(RefTok::Cols(a, b)) => {
                let (x, y) = (r1c1_part('C', a, col), r1c1_part('C', b, col));
                if x == y { out.push_str(&x) } else { out.push_str(&format!("{x}:{y}")) }
            }
            Seg::Ref(RefTok::Rows(a, b)) => {
                let (x, y) = (r1c1_part('R', a, row), r1c1_part('R', b, row));
                if x == y { out.push_str(&x) } else { out.push_str(&format!("{x}:{y}")) }
            }
        }
    }
    out
}

/// Reads the number part after `R` or `C`: `[−n]` relative, `n` absolute, or nothing (relative 0).
fn read_r1c1_num(ch: &[char], i: usize, origin: u32, max: u32) -> Option<(Coord, usize)> {
    if ch.get(i) == Some(&'[') {
        let mut j = i + 1;
        let start = j;
        if ch.get(j) == Some(&'-') || ch.get(j) == Some(&'+') {
            j += 1;
        }
        while j < ch.len() && ch[j].is_ascii_digit() {
            j += 1;
        }
        if ch.get(j) != Some(&']') {
            return None;
        }
        let d: i64 = ch[start..j].iter().collect::<String>().parse().ok()?;
        let n = i64::from(origin) + d;
        return (n >= 1 && n <= i64::from(max)).then_some((Coord { n: n as u32, abs: false }, j + 1));
    }
    let mut j = i;
    while j < ch.len() && ch[j].is_ascii_digit() {
        j += 1;
    }
    if j == i {
        return Some((Coord { n: origin, abs: false }, i));
    }
    let n: u32 = ch[i..j].iter().collect::<String>().parse().ok()?;
    (n >= 1 && n <= max).then_some((Coord { n, abs: true }, j))
}

enum R1 {
    Cell(Coord, Coord),
    Row(Coord),
    Col(Coord),
}

fn try_ref_r1c1(ch: &[char], i: usize, row: u32, col: u32) -> Option<(R1, usize)> {
    let up = |k: usize| ch.get(k).map(|c| c.to_ascii_uppercase());
    if up(i) == Some('R') {
        let (r, j) = read_r1c1_num(ch, i + 1, row, MAX_ROWS)?;
        if up(j) == Some('C') {
            let (c, k) = read_r1c1_num(ch, j + 1, col, MAX_COLS)?;
            return ends_ref(ch, k).then_some((R1::Cell(r, c), k));
        }
        return ends_ref(ch, j).then_some((R1::Row(r), j));
    }
    if up(i) == Some('C') {
        let (c, j) = read_r1c1_num(ch, i + 1, col, MAX_COLS)?;
        return ends_ref(ch, j).then_some((R1::Col(c), j));
    }
    None
}

/// Converts an R1C1 formula to A1 notation for the cell (row, col).
pub fn r1c1_to_a1(formula: &str, row: u32, col: u32) -> String {
    let ch: Vec<char> = formula.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < ch.len() {
        let c = ch[i];
        if c == '"' || c == '\'' {
            out.push(c);
            i += 1;
            while i < ch.len() {
                out.push(ch[i]);
                if ch[i] == c {
                    if ch.get(i + 1) == Some(&c) {
                        out.push(c);
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        let boundary = i == 0 || !is_ident(ch[i - 1]);
        if boundary && c.is_ascii_alphabetic() {
            if let Some((tok, j)) = try_ref_r1c1(&ch, i, row, col) {
                // A row-only or column-only reference may be the first half of a range.
                let mut j = j;
                let text = match tok {
                    R1::Cell(r, cc) => format!("{}{}", a1_col(cc), a1_row(r)),
                    R1::Row(a) => {
                        let mut b = a;
                        if ch.get(j) == Some(&':')
                            && let Some((R1::Row(x), k)) = try_ref_r1c1(&ch, j + 1, row, col)
                        {
                            b = x;
                            j = k;
                        }
                        format!("{}:{}", a1_row(a), a1_row(b))
                    }
                    R1::Col(a) => {
                        let mut b = a;
                        if ch.get(j) == Some(&':')
                            && let Some((R1::Col(x), k)) = try_ref_r1c1(&ch, j + 1, row, col)
                        {
                            b = x;
                            j = k;
                        }
                        format!("{}:{}", a1_col(a), a1_col(b))
                    }
                };
                out.push_str(&text);
                i = j;
                continue;
            }
            while i < ch.len() && is_ident(ch[i]) {
                out.push(ch[i]);
                i += 1;
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

/// A structural edit whose effect on formula references [`adjust_refs`] applies, the way the
/// application rewrites formulas when cells move.
#[derive(Clone, Debug)]
pub enum RefEdit {
    /// Whole rows inserted (`insert`) or deleted at `at`.
    Rows {
        sheet: String,
        at: u32,
        count: u32,
        insert: bool,
    },
    /// Whole columns inserted or deleted at `at`.
    Cols {
        sheet: String,
        at: u32,
        count: u32,
        insert: bool,
    },
    /// Cells in columns `c1..=c2` shifted down (insert) or up (delete) from row `at`.
    CellsVert {
        sheet: String,
        c1: u32,
        c2: u32,
        at: u32,
        count: u32,
        insert: bool,
    },
    /// Cells in rows `r1..=r2` shifted right (insert) or left (delete) from column `at`.
    CellsHorz {
        sheet: String,
        r1: u32,
        r2: u32,
        at: u32,
        count: u32,
        insert: bool,
    },
    /// A block cut and pasted: references inside `src` move by (`dr`, `dc`) to `to_sheet`.
    Move {
        sheet: String,
        src: crate::address::Rect,
        dr: i64,
        dc: i64,
        to_sheet: String,
    },
    RenameSheet {
        old: String,
        new: String,
    },
    DeleteSheet {
        name: String,
    },
}

/// One reference in formula text: optional sheet qualifier and one or two corners.
struct FullRef {
    sheet: Option<String>,
    a: RefTok,
    b: Option<RefTok>,
}

enum Item {
    Text(String),
    Ref(FullRef),
}

/// Splits the trailing `Sheet1!` / `'My Sheet'!` qualifier off text before a reference.
fn take_sheet_prefix(text: &mut String) -> Option<String> {
    let t = text.strip_suffix('!')?;
    if let Some(body) = t.strip_suffix('\'') {
        // Quoted: find the opening quote (doubled quotes are escapes).
        let ch: Vec<char> = body.chars().collect();
        let mut i = ch.len();
        while i > 0 {
            i -= 1;
            if ch[i] == '\'' {
                if i > 0 && ch[i - 1] == '\'' {
                    i -= 1;
                    continue;
                }
                let name: String = ch[i + 1..].iter().collect::<String>().replace("''", "'");
                let keep: String = ch[..i].iter().collect();
                *text = keep;
                return Some(name);
            }
        }
        return None;
    }
    let ch: Vec<char> = t.chars().collect();
    let mut i = ch.len();
    while i > 0 && (ch[i - 1].is_alphanumeric() || matches!(ch[i - 1], '_' | '.' | '#')) {
        i -= 1;
    }
    if i == ch.len() {
        return None;
    }
    let name: String = ch[i..].iter().collect();
    *text = ch[..i].iter().collect();
    Some(name)
}

fn items(formula: &str) -> Vec<Item> {
    let segs = scan_a1(formula);
    let mut out: Vec<Item> = Vec::new();
    let mut i = 0;
    while i < segs.len() {
        match &segs[i] {
            Seg::Text(t) => out.push(Item::Text(t.clone())),
            Seg::Ref(tok) => {
                let mut sheet = None;
                if let Some(Item::Text(prev)) = out.last_mut() {
                    sheet = take_sheet_prefix(prev);
                }
                // `A1:B2`: a cell, a lone ':' and another cell.
                let mut b = None;
                if let (Some(Seg::Text(colon)), Some(Seg::Ref(next))) = (segs.get(i + 1), segs.get(i + 2))
                    && colon == ":"
                    && matches!(tok, RefTok::Cell { .. })
                    && matches!(next, RefTok::Cell { .. })
                {
                    b = Some(*next);
                    i += 2;
                }
                out.push(Item::Ref(FullRef { sheet, a: *tok, b }));
            }
        }
        i += 1;
    }
    out
}

/// Row (or column) span edit: insert or delete `count` at `at`. `None` = the reference is gone.
fn span_edit(lo: u32, hi: u32, at: u32, count: u32, insert: bool, max: u32) -> Option<(u32, u32)> {
    if insert {
        let mv = |x: u32| if x >= at { x + count } else { x };
        let (l, h) = (mv(lo), mv(hi));
        return (h <= max).then_some((l, h));
    }
    let end = at + count - 1;
    if lo >= at && hi <= end {
        return None;
    }
    let l = if lo < at {
        lo
    } else if lo <= end {
        at
    } else {
        lo - count
    };
    let h = if hi < at {
        hi
    } else if hi <= end {
        at - 1
    } else {
        hi - count
    };
    Some((l, h))
}

fn ref_rows(r: &FullRef) -> Option<(u32, u32)> {
    match (r.a, r.b) {
        (RefTok::Cell { row, .. }, None) => Some((row.n, row.n)),
        (RefTok::Cell { row: a, .. }, Some(RefTok::Cell { row: b, .. })) => Some((a.n.min(b.n), a.n.max(b.n))),
        (RefTok::Rows(a, b), _) => Some((a.n.min(b.n), a.n.max(b.n))),
        _ => None,
    }
}

fn ref_cols(r: &FullRef) -> Option<(u32, u32)> {
    match (r.a, r.b) {
        (RefTok::Cell { col, .. }, None) => Some((col.n, col.n)),
        (RefTok::Cell { col: a, .. }, Some(RefTok::Cell { col: b, .. })) => Some((a.n.min(b.n), a.n.max(b.n))),
        (RefTok::Cols(a, b), _) => Some((a.n.min(b.n), a.n.max(b.n))),
        _ => None,
    }
}

/// Rewrites the row (`rows`) or column numbers of a reference through `f` (old -> new).
fn remap(r: &mut FullRef, rows: bool, lo: u32, hi: u32, nlo: u32, nhi: u32) {
    let map = |c: &mut Coord| {
        if c.n == lo {
            c.n = nlo;
        } else if c.n == hi {
            c.n = nhi;
        }
    };
    let fix = |t: &mut RefTok| match t {
        RefTok::Cell { row, col } => {
            if rows {
                map(row)
            } else {
                map(col)
            }
        }
        RefTok::Rows(a, b) if rows => {
            map(a);
            map(b)
        }
        RefTok::Cols(a, b) if !rows => {
            map(a);
            map(b)
        }
        _ => {}
    };
    if lo == hi {
        // Single row/column: every coordinate equal to it moves.
        let set = |c: &mut Coord| c.n = nlo;
        match &mut r.a {
            RefTok::Cell { row, col } => set(if rows { row } else { col }),
            RefTok::Rows(a, b) if rows => {
                set(a);
                set(b)
            }
            RefTok::Cols(a, b) if !rows => {
                set(a);
                set(b)
            }
            _ => {}
        }
        if let Some(RefTok::Cell { row, col }) = &mut r.b {
            set(if rows { row } else { col });
        }
        return;
    }
    fix(&mut r.a);
    if let Some(b) = &mut r.b {
        fix(b);
    }
}

fn same_sheet(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}

/// Applies `edit` to one reference; `false` when it no longer exists (`#REF!`).
fn edit_ref(r: &mut FullRef, own: &str, edit: &RefEdit) -> bool {
    let target = r.sheet.clone().unwrap_or_else(|| own.to_owned());
    match edit {
        RefEdit::RenameSheet { old, new } => {
            if r.sheet.as_deref().is_some_and(|s| same_sheet(s, old)) {
                r.sheet = Some(new.clone());
            }
            true
        }
        RefEdit::DeleteSheet { .. } => true,
        RefEdit::Rows { sheet, at, count, insert } | RefEdit::Cols { sheet, at, count, insert } => {
            if !same_sheet(&target, sheet) {
                return true;
            }
            let rows = matches!(edit, RefEdit::Rows { .. });
            let span = if rows { ref_rows(r) } else { ref_cols(r) };
            let Some((lo, hi)) = span else { return true };
            let max = if rows { MAX_ROWS } else { MAX_COLS };
            match span_edit(lo, hi, *at, *count, *insert, max) {
                Some((nlo, nhi)) => {
                    remap(r, rows, lo, hi, nlo, nhi);
                    true
                }
                None => false,
            }
        }
        RefEdit::CellsVert { sheet, c1, c2, at, count, insert } => {
            if !same_sheet(&target, sheet) {
                return true;
            }
            let (Some((lo, hi)), Some((cl, ch))) = (ref_rows(r), ref_cols(r)) else { return true };
            if cl < *c1 || ch > *c2 || hi < *at {
                return true;
            }
            match span_edit(lo, hi, *at, *count, *insert, MAX_ROWS) {
                Some((nlo, nhi)) => {
                    remap(r, true, lo, hi, nlo, nhi);
                    true
                }
                None => false,
            }
        }
        RefEdit::CellsHorz { sheet, r1, r2, at, count, insert } => {
            if !same_sheet(&target, sheet) {
                return true;
            }
            let (Some((lo, hi)), Some((rl, rh))) = (ref_cols(r), ref_rows(r)) else { return true };
            if rl < *r1 || rh > *r2 || hi < *at {
                return true;
            }
            match span_edit(lo, hi, *at, *count, *insert, MAX_COLS) {
                Some((nlo, nhi)) => {
                    remap(r, false, lo, hi, nlo, nhi);
                    true
                }
                None => false,
            }
        }
        RefEdit::Move { sheet, src, dr, dc, to_sheet } => {
            if !same_sheet(&target, sheet) {
                return true;
            }
            let (Some((rl, rh)), Some((cl, ch))) = (ref_rows(r), ref_cols(r)) else { return true };
            if !(src.contains(rl, cl) && src.contains(rh, ch)) {
                return true;
            }
            let mv = |c: &mut Coord, d: i64| c.n = (i64::from(c.n) + d) as u32;
            let shift = |t: &mut RefTok| {
                if let RefTok::Cell { row, col } = t {
                    mv(row, *dr);
                    mv(col, *dc);
                }
            };
            shift(&mut r.a);
            if let Some(b) = &mut r.b {
                shift(b);
            }
            if !same_sheet(to_sheet, sheet) {
                r.sheet = Some(to_sheet.clone());
            }
            true
        }
    }
}

/// Rewrites the references of `formula` (on sheet `own_sheet`) for a structural edit.
pub fn adjust_refs(formula: &str, own_sheet: &str, edit: &RefEdit) -> String {
    let mut out = String::new();
    for item in items(formula) {
        match item {
            Item::Text(t) => out.push_str(&t),
            Item::Ref(mut r) => {
                if let RefEdit::DeleteSheet { name } = edit
                    && r.sheet.as_deref().is_some_and(|s| same_sheet(s, name))
                {
                    out.push_str("#REF!");
                    out.push_str(&a1_text(&r.a));
                    if let Some(b) = r.b {
                        out.push(':');
                        out.push_str(&a1_text(&b));
                    }
                    continue;
                }
                let alive = edit_ref(&mut r, own_sheet, edit);
                if let Some(s) = &r.sheet {
                    out.push_str(&crate::address::quote_sheet(s));
                    out.push('!');
                }
                if !alive {
                    out.push_str("#REF!");
                    continue;
                }
                out.push_str(&a1_text(&r.a));
                if let Some(b) = r.b {
                    out.push(':');
                    out.push_str(&a1_text(&b));
                }
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifting() {
        assert_eq!(shift_formula("=A1*2", 1, 0), "=A2*2");
        assert_eq!(shift_formula("=$A$1+B$2+$C3", 2, 1), "=$A$1+C$2+$C5");
        assert_eq!(shift_formula("=SUM(A1:B2)", 1, 1), "=SUM(B2:C3)");
        assert_eq!(shift_formula("=\"A1\"&A1", 1, 0), "=\"A1\"&A2");
        assert_eq!(shift_formula("=Sheet1!A1", 0, 1), "=Sheet1!B1");
        assert_eq!(shift_formula("=LOG10(A1)", 1, 0), "=LOG10(A2)");
        assert_eq!(shift_formula("=A1", -1, 0), "=#REF!");
        assert_eq!(shift_formula("=SUM(A:A)", 0, 2), "=SUM(C:C)");
        assert_eq!(shift_formula("=SUM(1:2)", 3, 0), "=SUM(4:5)");
    }

    #[test]
    fn structural_edits() {
        let rows = |at, count, insert| RefEdit::Rows { sheet: "Sheet1".into(), at, count, insert };
        assert_eq!(adjust_refs("=SUM(A1:A2)", "Sheet1", &rows(2, 1, true)), "=SUM(A1:A3)");
        assert_eq!(adjust_refs("=$A$2*2", "Sheet1", &rows(2, 1, true)), "=$A$3*2");
        assert_eq!(adjust_refs("=SUM(A1:A5)", "Sheet1", &rows(3, 1, false)), "=SUM(A1:A4)");
        assert_eq!(adjust_refs("=A3", "Sheet1", &rows(3, 1, false)), "=#REF!");
        assert_eq!(adjust_refs("=Other!A3", "Sheet1", &rows(1, 1, true)), "=Other!A3");
        assert_eq!(
            adjust_refs("=Calc!A1*3", "Sheet1", &RefEdit::RenameSheet { old: "Calc".into(), new: "Calc 2".into() }),
            "='Calc 2'!A1*3"
        );
        assert_eq!(adjust_refs("=Gone!A1", "Sheet1", &RefEdit::DeleteSheet { name: "Gone".into() }), "=#REF!A1");
        assert_eq!(adjust_refs("='My Data'!B2", "Sheet1", &rows(1, 1, true)), "='My Data'!B2");
    }

    #[test]
    fn r1c1() {
        assert_eq!(a1_to_r1c1("=A1+$B$2", 2, 2), "=R[-1]C[-1]+R2C2");
        assert_eq!(r1c1_to_a1("=R[-1]C[-1]+R2C2", 2, 2), "=A1+$B$2");
        assert_eq!(r1c1_to_a1("=SUM(RC[-2]:RC[-1])", 5, 3), "=SUM(A5:B5)");
        assert_eq!(r1c1_to_a1("=ROUND(RC1,2)", 3, 3), "=ROUND($A3,2)");
        assert_eq!(r1c1_to_a1("=SUM(C1)", 3, 3), "=SUM($A:$A)");
    }
}
