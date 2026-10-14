//! Cell coordinates, rectangles and A1-style address text.

/// Rows in a sheet.
pub const MAX_ROWS: u32 = 1_048_576;
/// Columns in a sheet (`XFD`).
pub const MAX_COLS: u32 = 16_384;

/// A rectangle of cells, 1-based and inclusive: rows `r1..=r2`, columns `c1..=c2`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Rect {
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
}

impl Rect {
    /// A rectangle from two corners in any order.
    pub fn new(r1: u32, c1: u32, r2: u32, c2: u32) -> Self {
        Self { r1: r1.min(r2), c1: c1.min(c2), r2: r1.max(r2), c2: c1.max(c2) }
    }
    pub fn cell(row: u32, col: u32) -> Self {
        Self { r1: row, c1: col, r2: row, c2: col }
    }
    /// The whole sheet.
    pub fn sheet() -> Self {
        Self { r1: 1, c1: 1, r2: MAX_ROWS, c2: MAX_COLS }
    }
    pub fn rows(&self) -> u32 {
        self.r2 - self.r1 + 1
    }
    pub fn cols(&self) -> u32 {
        self.c2 - self.c1 + 1
    }
    pub fn area(&self) -> u64 {
        u64::from(self.rows()) * u64::from(self.cols())
    }
    pub fn is_cell(&self) -> bool {
        self.r1 == self.r2 && self.c1 == self.c2
    }
    pub fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.r1 && row <= self.r2 && col >= self.c1 && col <= self.c2
    }
    pub fn contains_rect(&self, other: &Rect) -> bool {
        self.contains(other.r1, other.c1) && self.contains(other.r2, other.c2)
    }
    pub fn intersect(&self, other: &Rect) -> Option<Rect> {
        let r1 = self.r1.max(other.r1);
        let c1 = self.c1.max(other.c1);
        let r2 = self.r2.min(other.r2);
        let c2 = self.c2.min(other.c2);
        (r1 <= r2 && c1 <= c2).then_some(Rect { r1, c1, r2, c2 })
    }
    /// Smallest rectangle holding both.
    pub fn bounding(&self, other: &Rect) -> Rect {
        Rect {
            r1: self.r1.min(other.r1),
            c1: self.c1.min(other.c1),
            r2: self.r2.max(other.r2),
            c2: self.c2.max(other.c2),
        }
    }
    /// Spans every column (whole rows).
    pub fn is_entire_rows(&self) -> bool {
        self.c1 == 1 && self.c2 == MAX_COLS
    }
    /// Spans every row (whole columns).
    pub fn is_entire_columns(&self) -> bool {
        self.r1 == 1 && self.r2 == MAX_ROWS
    }
    /// Moved by (dr, dc), or `None` if it would leave the sheet.
    pub fn offset(&self, dr: i64, dc: i64) -> Option<Rect> {
        let r1 = i64::from(self.r1) + dr;
        let r2 = i64::from(self.r2) + dr;
        let c1 = i64::from(self.c1) + dc;
        let c2 = i64::from(self.c2) + dc;
        let ok = r1 >= 1 && c1 >= 1 && r2 <= i64::from(MAX_ROWS) && c2 <= i64::from(MAX_COLS);
        ok.then_some(Rect { r1: r1 as u32, c1: c1 as u32, r2: r2 as u32, c2: c2 as u32 })
    }
}

/// `1 -> "A"`, `28 -> "AB"`, `16384 -> "XFD"`.
pub fn column_letters(col: u32) -> String {
    let mut n = col;
    let mut out = Vec::new();
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        out.push(b'A' + rem);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

/// `"A" -> 1`, `"xfd" -> 16384`; `None` if not letters or beyond the last column.
pub fn column_number(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut n: u32 = 0;
    for b in letters.bytes() {
        if !b.is_ascii_alphabetic() {
            return None;
        }
        n = n * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1);
    }
    (1..=MAX_COLS).contains(&n).then_some(n)
}

/// A parsed reference: optional sheet qualifier and one or more areas.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParsedRef {
    pub sheet: Option<String>,
    pub areas: Vec<Rect>,
}

/// One end of an A1 reference.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Part {
    Cell(u32, u32),
    Col(u32),
    Row(u32),
}

fn parse_part(text: &str) -> Option<Part> {
    let t = text.trim();
    let t = t.strip_prefix('$').unwrap_or(t);
    let letters_end = t.find(|c: char| !c.is_ascii_alphabetic()).unwrap_or(t.len());
    let (letters, rest) = t.split_at(letters_end);
    let rest = rest.strip_prefix('$').unwrap_or(rest);
    if letters.is_empty() {
        // Row only.
        let row: u32 = rest.parse().ok()?;
        return ((1..=MAX_ROWS).contains(&row) && rest.bytes().all(|b| b.is_ascii_digit())).then_some(Part::Row(row));
    }
    let col = column_number(letters)?;
    if rest.is_empty() {
        if t.len() != letters.len() {
            return None; // "A$" is not a column
        }
        return Some(Part::Col(col));
    }
    if !rest.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row: u32 = rest.parse().ok()?;
    (1..=MAX_ROWS).contains(&row).then_some(Part::Cell(row, col))
}

fn parse_area(text: &str) -> Option<Rect> {
    match text.split_once(':') {
        None => match parse_part(text)? {
            Part::Cell(r, c) => Some(Rect::cell(r, c)),
            _ => None,
        },
        Some((a, b)) => match (parse_part(a)?, parse_part(b)?) {
            (Part::Cell(r1, c1), Part::Cell(r2, c2)) => Some(Rect::new(r1, c1, r2, c2)),
            (Part::Col(c1), Part::Col(c2)) => Some(Rect::new(1, c1, MAX_ROWS, c2)),
            (Part::Row(r1), Part::Row(r2)) => Some(Rect::new(r1, 1, r2, MAX_COLS)),
            _ => None,
        },
    }
}

/// Splits on `sep` outside single-quoted sheet names.
fn split_outside_quotes(text: &str, sep: char) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut in_quote = false;
    let mut start = 0;
    for (i, c) in text.char_indices() {
        if c == '\'' {
            in_quote = !in_quote;
        } else if c == sep && !in_quote {
            parts.push(&text[start..i]);
            start = i + c.len_utf8();
        }
    }
    parts.push(&text[start..]);
    parts
}

/// Splits `Sheet1!A1` / `'My Sheet'!A1` into (sheet, reference).
fn split_sheet(text: &str) -> Option<(Option<String>, &str)> {
    let t = text.trim();
    if let Some(rest) = t.strip_prefix('\'') {
        // Quoted name; '' is an escaped quote.
        let mut name = String::new();
        let mut chars = rest.char_indices().peekable();
        while let Some((i, c)) = chars.next() {
            if c == '\'' {
                if matches!(chars.peek(), Some((_, '\''))) {
                    name.push('\'');
                    chars.next();
                    continue;
                }
                let after = &rest[i + 1..];
                let r = after.strip_prefix('!')?;
                return Some((Some(name), r));
            }
            name.push(c);
        }
        return None;
    }
    match t.rfind('!') {
        Some(i) => {
            let name = &t[..i];
            // A workbook qualifier `[Book1]Sheet1` is accepted and dropped.
            let name = match name.find(']') {
                Some(j) if name.starts_with('[') => &name[j + 1..],
                _ => name,
            };
            Some((Some(name.to_owned()), &t[i + 1..]))
        }
        None => Some((None, t)),
    }
}

/// Parses `A1`, `$A$1`, `A1:C10`, `A:A`, `1:3`, `Sheet1!A1`, `'My Sheet'!A1:B2` and
/// comma-separated multiple areas (`A1,B2:C3`). `None` if the text is not a reference.
pub fn parse_reference(text: &str) -> Option<ParsedRef> {
    let mut sheet: Option<String> = None;
    let mut areas = Vec::new();
    for (i, piece) in split_outside_quotes(text, ',').into_iter().enumerate() {
        let (s, r) = split_sheet(piece)?;
        if i == 0 {
            sheet = s;
        } else if let Some(s) = s {
            match &sheet {
                Some(first) if first.eq_ignore_ascii_case(&s) => {}
                None => sheet = Some(s),
                _ => return None,
            }
        }
        areas.push(parse_area(r)?);
    }
    if areas.is_empty() {
        return None;
    }
    Some(ParsedRef { sheet, areas })
}

/// A1 address text of one area, as `Range.Address` produces it.
pub fn format_area(r: &Rect, row_abs: bool, col_abs: bool) -> String {
    let rd = if row_abs { "$" } else { "" };
    let cd = if col_abs { "$" } else { "" };
    if r.is_entire_rows() {
        return format!("{rd}{}:{rd}{}", r.r1, r.r2);
    }
    if r.is_entire_columns() {
        return format!("{cd}{}:{cd}{}", column_letters(r.c1), column_letters(r.c2));
    }
    let a = format!("{cd}{}{rd}{}", column_letters(r.c1), r.r1);
    if r.is_cell() {
        return a;
    }
    format!("{a}:{cd}{}{rd}{}", column_letters(r.c2), r.r2)
}

/// R1C1 address text of one area. Relative parts are measured from `origin` (row, col).
pub fn format_area_r1c1(r: &Rect, row_abs: bool, col_abs: bool, origin: (u32, u32)) -> String {
    let row = |n: u32| -> String {
        if row_abs {
            format!("R{n}")
        } else {
            let d = i64::from(n) - i64::from(origin.0);
            if d == 0 { "R".to_owned() } else { format!("R[{d}]") }
        }
    };
    let col = |n: u32| -> String {
        if col_abs {
            format!("C{n}")
        } else {
            let d = i64::from(n) - i64::from(origin.1);
            if d == 0 { "C".to_owned() } else { format!("C[{d}]") }
        }
    };
    if r.is_entire_rows() {
        return if r.r1 == r.r2 { row(r.r1) } else { format!("{}:{}", row(r.r1), row(r.r2)) };
    }
    if r.is_entire_columns() {
        return if r.c1 == r.c2 { col(r.c1) } else { format!("{}:{}", col(r.c1), col(r.c2)) };
    }
    let a = format!("{}{}", row(r.r1), col(r.c1));
    if r.is_cell() {
        return a;
    }
    format!("{a}:{}{}", row(r.r2), col(r.c2))
}

/// Quotes a sheet name for use in a reference when it needs it.
pub fn quote_sheet(name: &str) -> String {
    let plain = name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
    if plain { name.to_owned() } else { format!("'{}'", name.replace('\'', "''")) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn letters() {
        assert_eq!(column_letters(1), "A");
        assert_eq!(column_letters(26), "Z");
        assert_eq!(column_letters(27), "AA");
        assert_eq!(column_letters(16384), "XFD");
        assert_eq!(column_number("xfd"), Some(16384));
        assert_eq!(column_number("XFE"), None);
        assert_eq!(column_number("A1"), None);
    }

    #[test]
    fn references() {
        let p = parse_reference("$A$1:c10").unwrap_or(ParsedRef { sheet: None, areas: vec![] });
        assert_eq!(p.areas, vec![Rect::new(1, 1, 10, 3)]);
        let p = parse_reference("'My Sheet'!B:B").unwrap_or(ParsedRef { sheet: None, areas: vec![] });
        assert_eq!(p.sheet.as_deref(), Some("My Sheet"));
        assert_eq!(p.areas, vec![Rect::new(1, 2, MAX_ROWS, 2)]);
        assert!(parse_reference("A").is_none());
        assert!(parse_reference("1").is_none());
        assert!(parse_reference("A0").is_none());
        assert!(parse_reference("A1048577").is_none());
        assert_eq!(parse_reference("A1,B2").map(|p| p.areas.len()), Some(2));
        assert_eq!(format_area(&Rect::new(1, 1, MAX_ROWS, 1), true, true), "$A:$A");
        assert_eq!(format_area(&Rect::new(2, 1, 3, MAX_COLS), true, true), "$2:$3");
        assert_eq!(format_area_r1c1(&Rect::new(2, 2, 3, 3), false, false, (1, 1)), "R[1]C[1]:R[2]C[2]");
    }
}
