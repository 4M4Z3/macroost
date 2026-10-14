//! Loading: package -> workbook part -> sheets, shared strings and styles -> [`MemoryWorkbook`].

use std::collections::HashMap;
use std::path::Path;

use macroost_lang::Value;
use macroost_sheet::address::column_number;
use macroost_sheet::{
    MAX_COLS, MAX_ROWS, MemoryWorkbook, Rect, SheetVisibility, Workbook, is_date_format, shift_formula,
};

use crate::styles::{Styles, parse_theme, truthy};
use crate::xml::{self, Event, Reader, attr, decode_xstring, local};
use crate::{Error, MAX_CELLS, Package, Rel, parse_rels, rels_path, resolve_target, xml_err};

/// Serial-date offset between the 1904 and 1900 date systems.
pub(crate) const DATE1904_OFFSET: f64 = 1462.0;
/// Column width the model reports for a column without an explicit width.
const DEFAULT_WIDTH: f64 = 8.43;

/// Loads a workbook from the bytes of an .xlsx/.xlsm file; the workbook is named `Book1.xlsx`.
pub fn load(bytes: &[u8]) -> Result<MemoryWorkbook, Error> {
    load_named(bytes, "Book1.xlsx")
}

/// Loads a workbook from a file; the workbook takes the file's name.
pub fn load_file(path: impl AsRef<Path>) -> Result<MemoryWorkbook, Error> {
    let path = path.as_ref();
    let bytes = crate::read_file(path)?;
    let name = path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_else(|| "Book1.xlsx".into());
    load_named(&bytes, &name)
}

/// Loads a workbook from bytes, giving it the file name `name` (`Workbook.Name`).
pub fn load_named(bytes: &[u8], name: &str) -> Result<MemoryWorkbook, Error> {
    let pkg = Package::read(bytes)?;
    let info = WorkbookInfo::read(&pkg)?;
    let sst = match &info.sst_part {
        Some(p) => match pkg.text(p)? {
            Some(t) => read_sst(&t, p)?.strings,
            None => Vec::new(),
        },
        None => Vec::new(),
    };
    let styles_text = match &info.styles_part {
        Some(p) => pkg.text(p)?,
        None => None,
    };
    let styles = Styles::parse(styles_text.as_deref(), &info.theme);
    let mut book = MemoryWorkbook::with_sheets(name, &[]);
    let mut budget = MAX_CELLS;
    // Original <sheet> index -> model sheet id (worksheets only).
    let mut ids: HashMap<usize, u32> = HashMap::new();
    for (i, entry) in info.sheets.iter().enumerate() {
        if !entry.is_worksheet {
            continue;
        }
        let id = add_sheet_named(&mut book, &entry.name)?;
        ids.insert(i, id);
        book.set_sheet_visibility(id, entry.state).ok();
        let Some(part) = &entry.part else { continue };
        let Some(text) = pkg.text(part)? else { continue };
        let sheet = parse_sheet(&text, part, &sst, &mut budget)?;
        apply_sheet(&mut book, id, &sheet, &styles, info.date1904);
    }
    book.set_workbook_protection(info.protection.clone());
    for d in &info.defined {
        let scope = match d.local_sheet {
            None => None,
            Some(i) => match ids.get(&i) {
                Some(id) => Some(*id),
                None => continue,
            },
        };
        if !d.text.is_empty() {
            book.set_defined_name(&d.name, scope, &d.text);
        }
    }
    Ok(book)
}

fn add_sheet_named(book: &mut MemoryWorkbook, name: &str) -> Result<u32, Error> {
    if let Ok(id) = book.add_sheet(name, usize::MAX) {
        return Ok(id);
    }
    for n in 1..100_000u32 {
        if let Ok(id) = book.add_sheet(&format!("Sheet{n}"), usize::MAX) {
            return Ok(id);
        }
    }
    Err(Error::Limit("too many sheets".into()))
}

fn apply_sheet(book: &mut MemoryWorkbook, id: u32, sheet: &ParsedSheet, styles: &Styles, date1904: bool) {
    for c in &sheet.cells {
        let format = styles.format(c.style);
        let mut value = c.value.clone();
        if date1904
            && let Value::Double(x) = value
            && format.number_format.as_deref().is_some_and(is_date_format)
        {
            value = Value::Double(x + DATE1904_OFFSET);
        }
        match &c.formula {
            Some(f) => {
                book.set_cell_formula_cached(id, c.row, c.col, format!("={f}"), value).ok();
            }
            None if !value.is_empty() => {
                book.set_cell_value(id, c.row, c.col, value).ok();
            }
            None => {}
        }
        if !format.is_default() {
            book.set_cell_format(id, c.row, c.col, format).ok();
        }
    }
    for col in &sheet.cols {
        for n in col.min..=col.max {
            if let Some(w) = col.width
                && (col.custom || (w - DEFAULT_WIDTH).abs() > 0.005)
            {
                book.set_column_width(id, n, w.clamp(0.0, 255.0)).ok();
            }
            if col.hidden {
                book.set_column_hidden(id, n, true).ok();
            }
        }
    }
    for row in &sheet.rows {
        if let Some(h) = row.height
            && (row.custom || (h - macroost_sheet::memory::DEFAULT_ROW_HEIGHT).abs() > 0.005)
        {
            book.set_row_height(id, row.row, h.clamp(0.0, 409.0)).ok();
        }
        if row.hidden {
            book.set_row_hidden(id, row.row, true).ok();
        }
    }
    for m in &sheet.merges {
        book.merge_cells(id, *m).ok();
    }
    if sheet.protected {
        book.set_protected(id, true).ok();
        book.set_sheet_password(id, sheet.password.clone());
    }
    if let Some(cn) = &sheet.code_name {
        book.set_code_name(id, cn);
    }
}

// ---------------------------------------------------------------- workbook part

pub(crate) struct SheetEntry {
    pub name: String,
    pub state: SheetVisibility,
    pub part: Option<String>,
    pub is_worksheet: bool,
    /// The original `<sheet>` element, for sheets the model cannot hold.
    pub raw: String,
    pub sheet_id: Option<u32>,
    pub rid: Option<String>,
}

pub(crate) struct DefinedName {
    pub name: String,
    pub local_sheet: Option<usize>,
    pub text: String,
    /// Attributes other than name/localSheetId, kept on save.
    pub attrs: Vec<(String, String)>,
}

pub(crate) struct WorkbookInfo {
    pub wb_part: String,
    pub wb_xml: String,
    pub wb_rels: Vec<Rel>,
    pub sheets: Vec<SheetEntry>,
    pub defined: Vec<DefinedName>,
    pub date1904: bool,
    pub styles_part: Option<String>,
    pub sst_part: Option<String>,
    pub theme: Vec<u32>,
    pub protection: Option<macroost_sheet::WorkbookProtection>,
}

impl WorkbookInfo {
    pub fn read(pkg: &Package) -> Result<Self, Error> {
        let root_rels = match pkg.text("_rels/.rels")? {
            Some(t) => parse_rels(&t, "_rels/.rels")?,
            None => Vec::new(),
        };
        let wb_part = root_rels
            .iter()
            .find(|r| r.kind_tail() == "officeDocument" && !r.external)
            .map(|r| resolve_target("", &r.target))
            .filter(|p| pkg.get(p).is_some())
            .unwrap_or_else(|| "xl/workbook.xml".to_owned());
        let wb_xml = pkg.text(&wb_part)?.ok_or_else(|| Error::Missing(wb_part.clone()))?;
        let rels_name = rels_path(&wb_part);
        let wb_rels = match pkg.text(&rels_name)? {
            Some(t) => parse_rels(&t, &rels_name)?,
            None => Vec::new(),
        };
        let root = xml::parse_tree(&wb_xml).map_err(|m| xml_err(&wb_part, m))?;
        let find_rel = |id: &str| wb_rels.iter().find(|r| r.id == id);
        let mut sheets = Vec::new();
        if let Some(list) = root.child("sheets") {
            for s in list.children_named("sheet") {
                let rid = s.attrs.iter().find(|(k, _)| local(k) == "id" && k.contains(':')).map(|(_, v)| v.clone());
                let rel = rid.as_deref().and_then(find_rel);
                let state = match s.attr("state") {
                    Some("hidden") => SheetVisibility::Hidden,
                    Some("veryHidden") => SheetVisibility::VeryHidden,
                    _ => SheetVisibility::Visible,
                };
                let is_worksheet = rel.is_none_or(|r| r.kind_tail() == "worksheet");
                sheets.push(SheetEntry {
                    name: s.attr("name").unwrap_or("").to_owned(),
                    state,
                    part: rel.filter(|r| !r.external).map(|r| resolve_target(&wb_part, &r.target)),
                    is_worksheet,
                    raw: s.to_xml(),
                    sheet_id: s.attr("sheetId").and_then(|v| v.trim().parse().ok()),
                    rid,
                });
            }
        }
        let mut defined = Vec::new();
        if let Some(list) = root.child("definedNames") {
            for d in list.children_named("definedName") {
                let text = d.text();
                defined.push(DefinedName {
                    name: d.attr("name").unwrap_or("").to_owned(),
                    local_sheet: d.attr("localSheetId").and_then(|v| v.trim().parse().ok()),
                    text: text.trim().strip_prefix('=').unwrap_or(text.trim()).to_owned(),
                    attrs: d.attrs.iter().filter(|(k, _)| k != "name" && k != "localSheetId").cloned().collect(),
                });
            }
        }
        let date1904 = root.child("workbookPr").and_then(|p| p.attr("date1904")).is_some_and(truthy);
        let part_of = |tail: &str| {
            wb_rels.iter().find(|r| r.kind_tail() == tail && !r.external).map(|r| resolve_target(&wb_part, &r.target))
        };
        let styles_part = part_of("styles");
        let sst_part = part_of("sharedStrings");
        let theme = match part_of("theme") {
            Some(p) => pkg.text(&p)?.map(|t| parse_theme(&t)).unwrap_or_default(),
            None => Vec::new(),
        };
        let protection = root.child("workbookProtection").map(|p| macroost_sheet::WorkbookProtection {
            structure: p.attr("lockStructure").is_some_and(truthy),
            windows: p.attr("lockWindows").is_some_and(truthy),
            password: crate::protection::read_hash(
                &|n: &str| p.attr(n).map(str::to_owned),
                "workbookPassword",
                "workbook",
            ),
        });
        Ok(Self { wb_part, wb_xml, wb_rels, sheets, defined, date1904, styles_part, sst_part, theme, protection })
    }
}

// ---------------------------------------------------------------- shared strings

pub(crate) struct SharedStrings {
    pub strings: Vec<String>,
    /// The original `<si>` element of each string.
    pub raw: Vec<String>,
}

pub(crate) fn read_sst(text: &str, part: &str) -> Result<SharedStrings, Error> {
    let mut r = Reader::new(text);
    let src = r.source();
    let mut strings = Vec::new();
    let mut raw = Vec::new();
    let mut cur: Option<(String, usize)> = None;
    let mut in_t = false;
    let mut rph = 0usize;
    while let Some(ev) = r.next() {
        match ev.map_err(|m| xml_err(part, m))? {
            Event::Start { name, empty, start, end, .. } => match local(name) {
                "si" if empty => {
                    strings.push(String::new());
                    raw.push(src.get(start..end).unwrap_or("").to_owned());
                }
                "si" => cur = Some((String::new(), start)),
                "rPh" if !empty => rph += 1,
                "t" if !empty && rph == 0 => in_t = true,
                _ => {}
            },
            Event::End { name, end, .. } => match local(name) {
                "si" => {
                    if let Some((s, start)) = cur.take() {
                        strings.push(decode_xstring(&s).into_owned());
                        raw.push(src.get(start..end).unwrap_or("").to_owned());
                    }
                }
                "rPh" => rph = rph.saturating_sub(1),
                "t" => in_t = false,
                _ => {}
            },
            Event::Text(t) => {
                if in_t && let Some((s, _)) = cur.as_mut() {
                    s.push_str(&t);
                }
            }
        }
        if strings.len() as u64 > MAX_CELLS {
            return Err(Error::Limit("too many shared strings".into()));
        }
    }
    Ok(SharedStrings { strings, raw })
}

// ---------------------------------------------------------------- worksheets

#[derive(Debug, Clone)]
pub(crate) struct ParsedCell {
    pub row: u32,
    pub col: u32,
    pub value: Value,
    /// Formula text without the leading `=`.
    pub formula: Option<String>,
    pub style: u32,
    /// `ref` of an array formula's master cell.
    pub array_ref: Option<String>,
}

#[derive(Debug, Clone)]
pub(crate) struct ColInfo {
    pub min: u32,
    pub max: u32,
    pub width: Option<f64>,
    pub custom: bool,
    pub hidden: bool,
}

#[derive(Debug, Clone)]
pub(crate) struct RowInfo {
    pub row: u32,
    pub height: Option<f64>,
    pub custom: bool,
    pub hidden: bool,
}

#[derive(Debug, Default)]
pub(crate) struct ParsedSheet {
    pub cells: Vec<ParsedCell>,
    pub cols: Vec<ColInfo>,
    pub rows: Vec<RowInfo>,
    pub merges: Vec<Rect>,
    pub protected: bool,
    pub password: Option<macroost_sheet::PasswordHash>,
    pub code_name: Option<String>,
}

/// `A1` -> (row, col).
pub(crate) fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let s = s.trim();
    let split = s.find(|c: char| c.is_ascii_digit())?;
    let letters = s.get(..split)?.trim_start_matches('$').trim_end_matches('$');
    let digits = s.get(split..)?;
    let col = column_number(&letters.to_ascii_uppercase())?;
    let row: u32 = digits.parse().ok()?;
    ((1..=MAX_ROWS).contains(&row) && (1..=MAX_COLS).contains(&col)).then_some((row, col))
}

/// `A1:C3` (or a single cell) -> rectangle.
pub(crate) fn parse_area_ref(s: &str) -> Option<Rect> {
    match s.split_once(':') {
        Some((a, b)) => {
            let (r1, c1) = parse_cell_ref(a)?;
            let (r2, c2) = parse_cell_ref(b)?;
            Some(Rect::new(r1, c1, r2, c2))
        }
        None => parse_cell_ref(s).map(|(r, c)| Rect::cell(r, c)),
    }
}

/// Converts a stored column width (characters plus cell padding) to the model's characters.
pub(crate) fn width_from_file(w: f64) -> f64 {
    let chars = ((w * 7.0 - 5.0) / 7.0).max(0.0);
    (chars * 100.0).round() / 100.0
}

pub(crate) fn error_code(text: &str) -> i32 {
    match text.trim() {
        "#NULL!" => 2000,
        "#DIV/0!" => 2007,
        "#VALUE!" => 2015,
        "#REF!" => 2023,
        "#NAME?" => 2029,
        "#NUM!" => 2036,
        "#N/A" => 2042,
        "#GETTING_DATA" => 2043,
        "#SPILL!" => 2045,
        "#CALC!" => 2050,
        _ => 2015,
    }
}

/// An ISO 8601 date-time (`t="d"` cells) -> serial date in the 1900 system.
fn iso_to_serial(s: &str) -> Option<f64> {
    let s = s.trim();
    let (date, time) = match s.split_once('T') {
        Some((d, t)) => (d, Some(t)),
        None => (s, None),
    };
    let mut it = date.split('-');
    let y: i64 = it.next()?.parse().ok()?;
    let m: i64 = it.next()?.parse().ok()?;
    let d: i64 = it.next()?.parse().ok()?;
    if !(1..=12).contains(&m) || !(1..=31).contains(&d) {
        return None;
    }
    // Days from civil (proleptic Gregorian), relative to 1899-12-30.
    let (yy, mm) = if m <= 2 { (y - 1, m + 9) } else { (y, m - 3) };
    let era = yy.div_euclid(400);
    let yoe = yy - era * 400;
    let doy = (153 * mm + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days = era * 146_097 + doe - 719_468; // days since 1970-01-01
    let mut serial = (days + 25_569) as f64;
    if serial < 61.0 {
        serial -= 1.0; // the 1900 system counts a nonexistent 1900-02-29
    }
    if let Some(t) = time {
        let t = t.trim_end_matches('Z');
        let t = t.split(['+']).next().unwrap_or(t);
        let mut p = t.split(':');
        let h: f64 = p.next().and_then(|x| x.parse().ok()).unwrap_or(0.0);
        let mi: f64 = p.next().and_then(|x| x.parse().ok()).unwrap_or(0.0);
        let se: f64 = p.next().and_then(|x| x.parse().ok()).unwrap_or(0.0);
        serial += (h * 3600.0 + mi * 60.0 + se) / 86_400.0;
    }
    Some(serial)
}

#[derive(Clone, Copy, PartialEq)]
enum Target {
    None,
    V,
    F,
    IsT,
}

#[derive(Default)]
struct CellState {
    row: u32,
    col: u32,
    valid: bool,
    t: String,
    s: u32,
    v: String,
    f: String,
    has_f: bool,
    f_type: String,
    f_si: Option<String>,
    f_ref: Option<String>,
    inline: String,
}

/// Parses one worksheet part. `budget` counts down the workbook-wide cell limit.
pub(crate) fn parse_sheet(text: &str, part: &str, sst: &[String], budget: &mut u64) -> Result<ParsedSheet, Error> {
    let mut out = ParsedSheet::default();
    let mut r = Reader::new(text);
    let mut in_data = false;
    let mut in_cols = false;
    let mut cur_row = 0u32;
    let mut last_col = 0u32;
    let mut cell: Option<CellState> = None;
    let mut target = Target::None;
    let mut in_is = false;
    let mut rph = 0usize;
    let mut shared: HashMap<String, (u32, u32, String)> = HashMap::new();
    while let Some(ev) = r.next() {
        let ev = ev.map_err(|m| xml_err(part, m))?;
        match ev {
            Event::Start { name, attrs, empty, .. } => {
                let ln = local(name);
                if in_data {
                    match ln {
                        "row" => {
                            cur_row = attr(&attrs, "r").and_then(|v| v.trim().parse().ok()).unwrap_or(cur_row + 1);
                            last_col = 0;
                            if (1..=MAX_ROWS).contains(&cur_row) {
                                let height = attr(&attrs, "ht").and_then(|v| v.trim().parse::<f64>().ok());
                                let custom = attr(&attrs, "customHeight").is_some_and(truthy);
                                let hidden = attr(&attrs, "hidden").is_some_and(truthy);
                                if height.is_some() || hidden {
                                    out.rows.push(RowInfo { row: cur_row, height, custom, hidden });
                                }
                            }
                        }
                        "c" => {
                            let pos = match attr(&attrs, "r") {
                                Some(a) => parse_cell_ref(a),
                                None => {
                                    let c = last_col + 1;
                                    ((1..=MAX_ROWS).contains(&cur_row) && c <= MAX_COLS).then_some((cur_row, c))
                                }
                            };
                            let mut st = CellState {
                                t: attr(&attrs, "t").unwrap_or("n").to_owned(),
                                s: attr(&attrs, "s").and_then(|v| v.trim().parse().ok()).unwrap_or(0),
                                ..CellState::default()
                            };
                            if let Some((rr, cc)) = pos {
                                st.row = rr;
                                st.col = cc;
                                st.valid = true;
                                cur_row = rr;
                                last_col = cc;
                            }
                            if empty {
                                finish_cell(st, sst, &mut shared, &mut out, budget)?;
                            } else {
                                cell = Some(st);
                            }
                        }
                        "v" if !empty => target = Target::V,
                        "f" => {
                            if let Some(c) = cell.as_mut() {
                                c.has_f = true;
                                c.f_type = attr(&attrs, "t").unwrap_or("normal").to_owned();
                                c.f_si = attr(&attrs, "si").map(str::to_owned);
                                c.f_ref = attr(&attrs, "ref").map(str::to_owned);
                            }
                            if !empty {
                                target = Target::F;
                            }
                        }
                        "is" if !empty => in_is = true,
                        "rPh" if !empty => rph += 1,
                        "t" if !empty && in_is && rph == 0 => target = Target::IsT,
                        _ => {}
                    }
                    continue;
                }
                match ln {
                    "sheetData" if !empty => in_data = true,
                    "cols" if !empty => in_cols = true,
                    "col" if in_cols => {
                        let n = |k: &str| attr(&attrs, k).and_then(|v| v.trim().parse::<u32>().ok());
                        let (Some(min), Some(max)) = (n("min"), n("max")) else { continue };
                        let min = min.clamp(1, MAX_COLS);
                        let max = max.clamp(min, MAX_COLS);
                        out.cols.push(ColInfo {
                            min,
                            max,
                            width: attr(&attrs, "width")
                                .and_then(|v| v.trim().parse::<f64>().ok())
                                .map(width_from_file),
                            custom: attr(&attrs, "customWidth").is_some_and(truthy),
                            hidden: attr(&attrs, "hidden").is_some_and(truthy),
                        });
                    }
                    "sheetPr" => out.code_name = attr(&attrs, "codeName").map(str::to_owned),
                    "mergeCell" => {
                        if let Some(rect) = attr(&attrs, "ref").and_then(parse_area_ref)
                            && !rect.is_cell()
                            && out.merges.len() < 1_000_000
                        {
                            out.merges.push(rect);
                        }
                    }
                    "sheetProtection" => {
                        out.protected = attr(&attrs, "sheet").is_some_and(truthy);
                        out.password =
                            crate::protection::read_hash(&|n: &str| attr(&attrs, n).map(str::to_owned), "password", "");
                    }
                    _ => {}
                }
            }
            Event::End { name, .. } => {
                let ln = local(name);
                if in_data {
                    match ln {
                        "sheetData" => in_data = false,
                        "c" => {
                            if let Some(st) = cell.take() {
                                finish_cell(st, sst, &mut shared, &mut out, budget)?;
                            }
                            target = Target::None;
                        }
                        "v" | "f" | "t" => target = Target::None,
                        "is" => in_is = false,
                        "rPh" => rph = rph.saturating_sub(1),
                        _ => {}
                    }
                } else if ln == "cols" {
                    in_cols = false;
                }
            }
            Event::Text(t) => {
                if let Some(c) = cell.as_mut() {
                    match target {
                        Target::V => c.v.push_str(&t),
                        Target::F => c.f.push_str(&t),
                        Target::IsT => c.inline.push_str(&t),
                        Target::None => {}
                    }
                }
            }
        }
    }
    Ok(out)
}

fn finish_cell(
    st: CellState,
    sst: &[String],
    shared: &mut HashMap<String, (u32, u32, String)>,
    out: &mut ParsedSheet,
    budget: &mut u64,
) -> Result<(), Error> {
    if !st.valid {
        return Ok(());
    }
    let v = st.v.trim();
    let value = match st.t.as_str() {
        "s" => v.parse::<usize>().ok().and_then(|i| sst.get(i)).map(|s| Value::from(s.as_str())).unwrap_or_default(),
        "str" => Value::from(decode_xstring(&st.v).as_ref()),
        "inlineStr" => Value::from(decode_xstring(&st.inline).as_ref()),
        "b" => {
            if v.is_empty() {
                Value::Empty
            } else {
                Value::Boolean(v == "1" || v.eq_ignore_ascii_case("true"))
            }
        }
        "e" => {
            if v.is_empty() {
                Value::Empty
            } else {
                Value::Error(error_code(v))
            }
        }
        "d" => iso_to_serial(v).map(Value::Double).unwrap_or_default(),
        _ => match v.parse::<f64>() {
            Ok(x) if x.is_finite() => Value::Double(x),
            _ => Value::Empty,
        },
    };
    let ftext = decode_xstring(&st.f).into_owned();
    let formula = if !st.has_f {
        None
    } else {
        match st.f_type.as_str() {
            "shared" => {
                let si = st.f_si.clone().unwrap_or_default();
                if !ftext.trim().is_empty() {
                    if shared.len() < 1_000_000 {
                        shared.insert(si, (st.row, st.col, ftext.clone()));
                    }
                    Some(ftext)
                } else {
                    shared.get(&si).map(|(mr, mc, text)| {
                        shift_formula(text, i64::from(st.row) - i64::from(*mr), i64::from(st.col) - i64::from(*mc))
                    })
                }
            }
            "dataTable" => None,
            _ => (!ftext.trim().is_empty()).then_some(ftext),
        }
    };
    if value.is_empty() && formula.is_none() && st.s == 0 {
        return Ok(());
    }
    if *budget == 0 {
        return Err(Error::Limit("more than 50 million cells".into()));
    }
    *budget -= 1;
    let array_ref = if st.f_type == "array" { st.f_ref } else { None };
    out.cells.push(ParsedCell { row: st.row, col: st.col, value, formula, style: st.s, array_ref });
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn refs_and_dates() {
        assert_eq!(parse_cell_ref("B3"), Some((3, 2)));
        assert_eq!(parse_cell_ref("XFD1048576"), Some((1_048_576, 16_384)));
        assert_eq!(parse_cell_ref("XFE1"), None);
        assert_eq!(parse_cell_ref("A0"), None);
        assert_eq!(iso_to_serial("2026-10-14"), Some(46_309.0));
        assert_eq!(iso_to_serial("1900-01-01"), Some(1.0));
        assert_eq!(iso_to_serial("1900-03-01"), Some(61.0));
        assert_eq!(width_from_file(9.140625), 8.43);
    }
}
