//! Saving: regenerate the workbook sheet list, defined names, worksheets' cell data, shared
//! strings and styles, splicing them into the template's parts; copy every other part as is.

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Write};

use macroost_lang::Value;
use macroost_sheet::address::column_letters;
use macroost_sheet::{MemoryWorkbook, Rect, SheetVisibility, Workbook, is_date_format};

use crate::read::{DATE1904_OFFSET, SharedStrings, WorkbookInfo, parse_sheet, read_sst};
use crate::styles::{MINIMAL_STYLES, Styles};
use crate::xml::{self, Element, Node, Outline, encode_xstring, escape};
use crate::{Error, Package, Rel, relative_target, rels_path, xml_err};

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_WORKSHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SST: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

const WORKBOOK_ORDER: &[&str] = &[
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
];

const SHEET_ORDER: &[&str] = &[
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

/// A package for a new, empty workbook (the base of a save without a template).
fn minimal_package() -> Package {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                "{XML_DECL}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"{CT_STYLES}\"/></Types>"
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                "{XML_DECL}<Relationships xmlns=\"{NS_PKG_REL}\"><Relationship Id=\"rId1\" Type=\"{NS_REL}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
            )
            .into_bytes(),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(
                "{XML_DECL}<workbook xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\"><bookViews><workbookView/></bookViews><sheets/><calcPr calcId=\"191029\"/></workbook>"
            )
            .into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(
                "{XML_DECL}<Relationships xmlns=\"{NS_PKG_REL}\"><Relationship Id=\"rId1\" Type=\"{NS_REL}/styles\" Target=\"styles.xml\"/></Relationships>"
            )
            .into_bytes(),
        ),
        ("xl/styles.xml".to_owned(), MINIMAL_STYLES.as_bytes().to_vec()),
    ];
    Package { parts }
}

/// Per-cell facts from the template that the model does not keep.
struct OrigCell {
    style: u32,
    array_ref: Option<String>,
    formula: Option<String>,
}

struct SstOut {
    raws: Vec<String>,
    index: HashMap<String, u32>,
    orig: HashMap<String, String>,
    refs: u64,
}

impl SstOut {
    fn new(orig: Option<SharedStrings>) -> Self {
        let mut map = HashMap::new();
        if let Some(o) = orig {
            for (s, r) in o.strings.into_iter().zip(o.raw) {
                map.entry(s).or_insert(r);
            }
        }
        Self { raws: Vec::new(), index: HashMap::new(), orig: map, refs: 0 }
    }

    fn get(&mut self, s: &str) -> u32 {
        self.refs += 1;
        if let Some(&i) = self.index.get(s) {
            return i;
        }
        let raw = match self.orig.get(s) {
            Some(r) => r.clone(),
            None => {
                let sp = if s.starts_with(char::is_whitespace) || s.ends_with(char::is_whitespace) {
                    " xml:space=\"preserve\""
                } else {
                    ""
                };
                format!("<si><t{sp}>{}</t></si>", encode_xstring(s))
            }
        };
        let i = self.raws.len() as u32;
        self.raws.push(raw);
        self.index.insert(s.to_owned(), i);
        i
    }

    fn to_xml(&self, ns: &str) -> String {
        let mut s =
            format!("{XML_DECL}<sst xmlns=\"{ns}\" count=\"{}\" uniqueCount=\"{}\">", self.refs, self.raws.len());
        for r in &self.raws {
            s.push_str(r);
        }
        s.push_str("</sst>");
        s
    }
}

/// Saves `book` as an .xlsx/.xlsm package. With `template` (the bytes of the file the book was
/// loaded from) every part that is not regenerated is copied unchanged; sheets are matched to the
/// template's worksheets by [`macroost_sheet::SheetId`] (the n-th loaded worksheet has id n).
pub fn save(book: &MemoryWorkbook, template: Option<&[u8]>) -> Result<Vec<u8>, Error> {
    let pkg = match template {
        Some(b) => Package::read(b)?,
        None => minimal_package(),
    };
    let info = WorkbookInfo::read(&pkg)?;
    let wb_part = info.wb_part.clone();
    let wb_root = Outline::parse(&info.wb_xml).map_err(|m| xml_err(&wb_part, m))?;
    let wb_prefix = wb_root.root_prefix().map(|p| format!("{p}:")).unwrap_or_default();
    let wb_tree = xml::parse_tree(&info.wb_xml).map_err(|m| xml_err(&wb_part, m))?;
    let ns_main = namespace_of(&wb_tree).unwrap_or(NS_MAIN).to_owned();
    let rel_prefix = wb_tree
        .attrs
        .iter()
        .find(|(k, v)| k.starts_with("xmlns:") && (v == NS_REL || v == NS_REL_STRICT))
        .map(|(k, _)| k.trim_start_matches("xmlns:").to_owned());
    let rel_base = info
        .wb_rels
        .iter()
        .find(|r| r.kind_tail() == "worksheet")
        .and_then(|r| r.kind.rsplit_once('/').map(|(b, _)| b.to_owned()))
        .unwrap_or_else(|| NS_REL.to_owned());

    // Shared strings and styles start from the template's.
    let orig_sst = match &info.sst_part {
        Some(p) => match pkg.text(p)? {
            Some(t) => Some(read_sst(&t, p)?),
            None => None,
        },
        None => None,
    };
    let orig_strings: Vec<String> = orig_sst.as_ref().map(|s| s.strings.clone()).unwrap_or_default();
    let mut sst = SstOut::new(orig_sst);
    let styles_text = match &info.styles_part {
        Some(p) => pkg.text(p)?,
        None => None,
    };
    let mut styles = Styles::parse(styles_text.as_deref(), &info.theme);

    // Model sheet id -> original <sheet> index (worksheets only, in order).
    let orig_ws: Vec<usize> = info.sheets.iter().enumerate().filter(|(_, s)| s.is_worksheet).map(|(i, _)| i).collect();
    let orig_of = |id: u32| -> Option<usize> { orig_ws.get((id as usize).checked_sub(1)?).copied() };

    let mut rels: Vec<Rel> = info.wb_rels.clone();
    let mut new_parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut replaced: HashMap<String, Vec<u8>> = HashMap::new();
    let mut removed: HashSet<String> = HashSet::new();
    let mut add_overrides: Vec<(String, &str)> = Vec::new();
    let mut used_ids: HashSet<u32> = info.sheets.iter().filter_map(|s| s.sheet_id).collect();
    let mut next_rid = 1u32;
    let mut fresh_rid = |rels: &Vec<Rel>| loop {
        let id = format!("rId{next_rid}");
        next_rid += 1;
        if !rels.iter().any(|r| r.id == id) {
            return id;
        }
    };
    let existing: HashSet<String> = pkg.parts.iter().map(|(n, _)| n.to_ascii_lowercase()).collect();
    let mut next_part = 1u32;
    let mut fresh_part = |new_parts: &Vec<(String, Vec<u8>)>| loop {
        let p = format!("xl/worksheets/sheet{next_part}.xml");
        next_part += 1;
        if !existing.contains(&p.to_ascii_lowercase()) && !new_parts.iter().any(|(n, _)| *n == p) {
            return p;
        }
    };

    // ---- worksheets
    let model_ids = book.sheet_ids();
    let kept_orig: HashSet<usize> = model_ids.iter().filter_map(|&id| orig_of(id)).collect();
    // Final sheet list entries: Ok(model id) or Err(original index of a non-worksheet sheet).
    let mut sheet_xml: HashMap<u32, String> = HashMap::new();
    for &id in &model_ids {
        let orig = orig_of(id).and_then(|i| info.sheets.get(i));
        let (part, rid, sheet_id, template_text) = match orig {
            Some(e) if e.part.as_ref().is_some_and(|p| pkg.get(p).is_some()) && e.rid.is_some() => {
                let part = e.part.clone().unwrap_or_default();
                let text = pkg.text(&part)?.unwrap_or_default();
                (part, e.rid.clone().unwrap_or_default(), e.sheet_id, Some(text))
            }
            _ => {
                let part = fresh_part(&new_parts);
                let rid = fresh_rid(&rels);
                rels.push(Rel {
                    id: rid.clone(),
                    kind: format!("{rel_base}/worksheet"),
                    target: relative_target(&wb_part, &part),
                    external: false,
                });
                add_overrides.push((part.clone(), CT_WORKSHEET));
                (part, rid, None, None)
            }
        };
        let sheet_id = match sheet_id {
            Some(s) => s,
            None => {
                let s = (1..).find(|n| !used_ids.contains(n)).unwrap_or(1);
                used_ids.insert(s);
                s
            }
        };
        let xml = write_sheet(
            book,
            id,
            template_text.as_deref(),
            &part,
            &orig_strings,
            &mut styles,
            &mut sst,
            info.date1904,
            &ns_main,
            &rel_prefix,
        )?;
        if template_text.is_some() {
            replaced.insert(part.to_ascii_lowercase(), xml.into_bytes());
        } else {
            new_parts.push((part.clone(), xml.into_bytes()));
        }
        let name = book.sheet_name(id).unwrap_or_default();
        let state = match book.sheet_visibility(id) {
            SheetVisibility::Visible => "",
            SheetVisibility::Hidden => " state=\"hidden\"",
            SheetVisibility::VeryHidden => " state=\"veryHidden\"",
        };
        let rp = rel_prefix.as_deref().unwrap_or("r");
        sheet_xml.insert(
            id,
            format!(
                "<{wb_prefix}sheet name=\"{}\" sheetId=\"{sheet_id}\"{state} {rp}:id=\"{}\"/>",
                escape(&name),
                escape(&rid)
            ),
        );
    }
    // Deleted worksheets: drop their parts and relationships.
    for &i in &orig_ws {
        if kept_orig.contains(&i) {
            continue;
        }
        if let Some(e) = info.sheets.get(i) {
            if let Some(p) = &e.part {
                removed.insert(p.to_ascii_lowercase());
                removed.insert(rels_path(p).to_ascii_lowercase());
            }
            if let Some(rid) = &e.rid {
                rels.retain(|r| &r.id != rid);
            }
        }
    }
    // The calculation chain lists formula cells; a stale one makes readers repair the file.
    if let Some(r) = rels.iter().find(|r| r.kind_tail() == "calcChain") {
        removed.insert(crate::resolve_target(&wb_part, &r.target).to_ascii_lowercase());
    }
    rels.retain(|r| r.kind_tail() != "calcChain");

    // Final order: model sheets, with other sheet kinds after the worksheet they followed.
    let mut order: Vec<Result<u32, usize>> = model_ids.iter().map(|&id| Ok(id)).collect();
    let mut front = 0usize;
    for (i, e) in info.sheets.iter().enumerate() {
        if e.is_worksheet {
            continue;
        }
        let prev_ws = (0..i).rev().find(|&j| info.sheets.get(j).is_some_and(|s| s.is_worksheet));
        let at = match prev_ws {
            None => {
                front += 1;
                front - 1
            }
            Some(j) => {
                let model = model_ids.iter().find(|&&id| orig_of(id) == Some(j));
                match model.and_then(|m| order.iter().position(|o| *o == Ok(*m))) {
                    Some(p) => {
                        // after the worksheet and any earlier non-worksheets placed after it
                        let mut q = p + 1;
                        while matches!(order.get(q), Some(Err(_))) {
                            q += 1;
                        }
                        q
                    }
                    None => order.len(),
                }
            }
        };
        order.insert(at.min(order.len()), Err(i));
    }
    let mut sheets_frag = format!("<{wb_prefix}sheets");
    if rel_prefix.is_none() {
        sheets_frag.push_str(&format!(" xmlns:r=\"{NS_REL}\""));
    }
    sheets_frag.push('>');
    for o in &order {
        match o {
            Ok(id) => sheets_frag.push_str(sheet_xml.get(id).map_or("", String::as_str)),
            Err(i) => sheets_frag.push_str(info.sheets.get(*i).map_or("", |s| s.raw.as_str())),
        }
    }
    sheets_frag.push_str(&format!("</{wb_prefix}sheets>"));
    let final_index = |id: u32| order.iter().position(|o| *o == Ok(id));
    let orig_index_of_final = |pos: usize| -> Option<usize> {
        match order.get(pos)? {
            Ok(id) => orig_of(*id),
            Err(i) => Some(*i),
        }
    };

    // ---- defined names
    let mut names_frag = String::new();
    let mut emitted: HashSet<(String, Option<usize>)> = HashSet::new();
    for (name, scope, text) in book.defined_names() {
        let local = match scope {
            None => None,
            Some(id) => match final_index(id) {
                Some(i) => Some(i),
                None => continue,
            },
        };
        let orig_scope = local.and_then(orig_index_of_final);
        let attrs = info
            .defined
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(&name) && d.local_sheet == orig_scope)
            .map(|d| d.attrs.clone())
            .unwrap_or_default();
        emitted.insert((name.to_ascii_lowercase(), local));
        names_frag.push_str(&format!("<{wb_prefix}definedName name=\"{}\"", escape(&name)));
        if let Some(l) = local {
            names_frag.push_str(&format!(" localSheetId=\"{l}\""));
        }
        for (k, v) in attrs {
            names_frag.push_str(&format!(" {k}=\"{}\"", escape(&v)));
        }
        names_frag.push_str(&format!(">{}</{wb_prefix}definedName>", escape(&text)));
    }
    // Names scoped to sheets the model cannot hold (chart sheets) are kept.
    for d in &info.defined {
        let Some(ls) = d.local_sheet else { continue };
        if info.sheets.get(ls).is_some_and(|s| !s.is_worksheet)
            && let Some(pos) = order.iter().position(|o| *o == Err(ls))
            && !emitted.contains(&(d.name.to_ascii_lowercase(), Some(pos)))
        {
            names_frag
                .push_str(&format!("<{wb_prefix}definedName name=\"{}\" localSheetId=\"{pos}\"", escape(&d.name)));
            for (k, v) in &d.attrs {
                names_frag.push_str(&format!(" {k}=\"{}\"", escape(v)));
            }
            names_frag.push_str(&format!(">{}</{wb_prefix}definedName>", escape(&d.text)));
        }
    }
    let names_frag =
        (!names_frag.is_empty()).then(|| format!("<{wb_prefix}definedNames>{names_frag}</{wb_prefix}definedNames>"));

    // ---- book views: keep the active and first tab in range and visible
    let n_sheets = order.len();
    let first_visible = order
        .iter()
        .position(|o| match o {
            Ok(id) => book.sheet_visibility(*id) == SheetVisibility::Visible,
            Err(i) => info.sheets.get(*i).is_some_and(|s| s.state == SheetVisibility::Visible),
        })
        .unwrap_or(0);
    let book_views = match wb_root.child("bookViews") {
        Some(raw) => match xml::parse_tree(raw) {
            Ok(mut t) => {
                for n in &mut t.children {
                    if let Node::Elem(v) = n {
                        for a in ["activeTab", "firstSheet"] {
                            if let Some(x) = v.attr(a).and_then(|x| x.trim().parse::<usize>().ok()) {
                                let ok = x < n_sheets
                                    && match order.get(x) {
                                        Some(Ok(id)) => book.sheet_visibility(*id) == SheetVisibility::Visible,
                                        _ => true,
                                    };
                                if !ok {
                                    v.set_attr(a, &first_visible.to_string());
                                }
                            }
                        }
                    }
                }
                Some(t.to_xml())
            }
            Err(_) => Some(raw.to_owned()),
        },
        None => Some(format!("<{wb_prefix}bookViews><{wb_prefix}workbookView/></{wb_prefix}bookViews>")),
    };
    // ---- calcPr: ask readers to recalculate, since cached values may be stale after macros.
    let calc = match wb_root.child("calcPr").map(xml::parse_tree) {
        Some(Ok(mut t)) => {
            t.set_attr("fullCalcOnLoad", "1");
            Some(t.to_xml())
        }
        Some(Err(_)) | None => Some(format!("<{wb_prefix}calcPr fullCalcOnLoad=\"1\"/>")),
    };
    // ---- workbookProtection: kept as written unless a macro changed it.
    let wb_protection = book.workbook_protection().map(|prot| match wb_root.child("workbookProtection") {
        Some(raw) if info.protection.as_ref() == Some(&prot) => raw.to_owned(),
        _ => format!(
            "<{wb_prefix}workbookProtection{}{}{}/>",
            crate::protection::hash_attrs(prot.password.as_ref(), "workbookPassword", "workbook"),
            if prot.structure { " lockStructure=\"1\"" } else { "" },
            if prot.windows { " lockWindows=\"1\"" } else { "" },
        ),
    });
    let wb_new = wb_root.rebuild(
        WORKBOOK_ORDER,
        &[
            ("sheets", Some(sheets_frag)),
            ("definedNames", names_frag),
            ("bookViews", book_views),
            ("calcPr", calc),
            ("workbookProtection", wb_protection),
        ],
    );
    replaced.insert(wb_part.to_ascii_lowercase(), wb_new.into_bytes());

    // ---- shared strings and styles
    let sst_part = match &info.sst_part {
        Some(p) => p.clone(),
        None => {
            let p = unused_part(&existing, "xl/sharedStrings.xml");
            let rid = fresh_rid(&rels);
            rels.push(Rel {
                id: rid,
                kind: format!("{rel_base}/sharedStrings"),
                target: relative_target(&wb_part, &p),
                external: false,
            });
            p
        }
    };
    let sst_xml = sst.to_xml(&ns_main).into_bytes();
    if pkg.get(&sst_part).is_some() {
        replaced.insert(sst_part.to_ascii_lowercase(), sst_xml);
    } else {
        new_parts.push((sst_part.clone(), sst_xml));
    }
    add_overrides.push((sst_part, CT_SST));
    let styles_part = match &info.styles_part {
        Some(p) => p.clone(),
        None => {
            let p = unused_part(&existing, "xl/styles.xml");
            let rid = fresh_rid(&rels);
            rels.push(Rel {
                id: rid,
                kind: format!("{rel_base}/styles"),
                target: relative_target(&wb_part, &p),
                external: false,
            });
            p
        }
    };
    let styles_xml = styles.to_xml().into_bytes();
    if pkg.get(&styles_part).is_some() {
        replaced.insert(styles_part.to_ascii_lowercase(), styles_xml);
    } else {
        new_parts.push((styles_part.clone(), styles_xml));
    }
    add_overrides.push((styles_part, CT_STYLES));

    // ---- workbook relationships and content types
    let rels_name = rels_path(&wb_part);
    let rels_xml = write_rels(&rels).into_bytes();
    if pkg.get(&rels_name).is_some() {
        replaced.insert(rels_name.to_ascii_lowercase(), rels_xml);
    } else {
        new_parts.push((rels_name, rels_xml));
    }
    let ct_text = pkg.text("[Content_Types].xml")?.ok_or_else(|| Error::Missing("[Content_Types].xml".into()))?;
    let ct = content_types(&ct_text, &removed, &add_overrides)?;
    replaced.insert("[content_types].xml".into(), ct.into_bytes());

    // ---- write the package
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut write_part = |name: &str, data: &[u8]| -> Result<(), Error> {
        zip.start_file(name, opts)?;
        zip.write_all(data)?;
        Ok(())
    };
    if let Some(ct) = replaced.get("[content_types].xml") {
        write_part("[Content_Types].xml", ct)?;
    }
    for (name, data) in &pkg.parts {
        let key = name.to_ascii_lowercase();
        if key == "[content_types].xml" || removed.contains(&key) {
            continue;
        }
        match replaced.get(&key) {
            Some(new) => write_part(name, new)?,
            None => write_part(name, data)?,
        }
    }
    for (name, data) in &new_parts {
        write_part(name, data)?;
    }
    let out = zip.finish()?;
    Ok(out.into_inner())
}

fn unused_part(existing: &HashSet<String>, want: &str) -> String {
    if !existing.contains(&want.to_ascii_lowercase()) {
        return want.to_owned();
    }
    let (stem, ext) = want.rsplit_once('.').unwrap_or((want, "xml"));
    (2..).map(|n| format!("{stem}{n}.{ext}")).find(|p| !existing.contains(&p.to_ascii_lowercase())).unwrap_or_default()
}

/// The namespace URI of an element's own prefix (or the default namespace).
fn namespace_of(e: &Element) -> Option<&str> {
    let decl = match e.name.split_once(':') {
        Some((p, _)) => format!("xmlns:{p}"),
        None => "xmlns".to_owned(),
    };
    e.attrs.iter().find(|(k, _)| *k == decl).map(|(_, v)| v.as_str())
}

fn write_rels(rels: &[Rel]) -> String {
    let mut s = format!("{XML_DECL}<Relationships xmlns=\"{NS_PKG_REL}\">");
    for r in rels {
        s.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"{}/>",
            escape(&r.id),
            escape(&r.kind),
            escape(&r.target),
            if r.external { " TargetMode=\"External\"" } else { "" }
        ));
    }
    s.push_str("</Relationships>");
    s
}

fn content_types(text: &str, removed: &HashSet<String>, add: &[(String, &str)]) -> Result<String, Error> {
    let mut root = xml::parse_tree(text).map_err(|m| xml_err("[Content_Types].xml", m))?;
    root.children.retain(|n| match n {
        Node::Elem(e) if e.local() == "Override" => {
            let p = e.attr("PartName").unwrap_or("").trim_start_matches('/').to_ascii_lowercase();
            !removed.contains(&p)
        }
        Node::Text(_) => false,
        _ => true,
    });
    let has_default = |root: &Element, ext: &str| {
        root.children_named("Default").any(|d| d.attr("Extension").is_some_and(|x| x.eq_ignore_ascii_case(ext)))
    };
    for (ext, ct) in [("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")]
    {
        if !has_default(&root, ext) {
            let mut d = Element::new("Default");
            d.set_attr("Extension", ext);
            d.set_attr("ContentType", ct);
            root.children.insert(0, Node::Elem(d));
        }
    }
    for (part, ct) in add {
        let pn = format!("/{part}");
        let exists = root.children.iter_mut().find_map(|n| match n {
            Node::Elem(e)
                if e.local() == "Override" && e.attr("PartName").is_some_and(|p| p.eq_ignore_ascii_case(&pn)) =>
            {
                Some(e)
            }
            _ => None,
        });
        match exists {
            Some(e) => e.set_attr("ContentType", ct),
            None => {
                let mut o = Element::new("Override");
                o.set_attr("PartName", &pn);
                o.set_attr("ContentType", ct);
                root.children.push(Node::Elem(o));
            }
        }
    }
    Ok(format!("{XML_DECL}{}", root.to_xml()))
}

/// Stored column width (characters plus padding) for the model's characters.
fn width_to_file(chars: f64) -> f64 {
    ((chars * 7.0 + 5.0) / 7.0 * 256.0).trunc() / 256.0
}

fn area_ref(r: &Rect) -> String {
    let a = format!("{}{}", column_letters(r.c1), r.r1);
    if r.is_cell() { a } else { format!("{a}:{}{}", column_letters(r.c2), r.r2) }
}

fn fmt_num(x: f64) -> String {
    if x == x.trunc() && x.abs() < 1e15 { format!("{}", x as i64) } else { format!("{x}") }
}

fn error_text(code: i32) -> &'static str {
    match code {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2015 => "#VALUE!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        2042 => "#N/A",
        2043 => "#GETTING_DATA",
        2045 => "#SPILL!",
        2050 => "#CALC!",
        _ => "#VALUE!",
    }
}

/// A skeleton for a new worksheet; the regenerated pieces are spliced in like a template's.
fn skeleton(ns_main: &str) -> String {
    format!(
        "{XML_DECL}<worksheet xmlns=\"{ns_main}\" xmlns:r=\"{NS_REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>"
    )
}

#[allow(clippy::too_many_arguments)]
fn write_sheet(
    book: &MemoryWorkbook,
    id: u32,
    template: Option<&str>,
    part: &str,
    orig_strings: &[String],
    styles: &mut Styles,
    sst: &mut SstOut,
    date1904: bool,
    ns_main: &str,
    _rel_prefix: &Option<String>,
) -> Result<String, Error> {
    let skel;
    let text = match template {
        Some(t) => t,
        None => {
            skel = skeleton(ns_main);
            &skel
        }
    };
    let outline = Outline::parse(text).map_err(|m| xml_err(part, m))?;
    let p = outline.root_prefix().map(|p| format!("{p}:")).unwrap_or_default();
    // What the template knew about each cell.
    let mut orig: HashMap<(u32, u32), OrigCell> = HashMap::new();
    let mut orig_protection: Option<String> = None;
    let mut orig_password: Option<macroost_sheet::PasswordHash> = None;
    if template.is_some() {
        let mut budget = crate::MAX_CELLS;
        let parsed = parse_sheet(text, part, orig_strings, &mut budget)?;
        for c in parsed.cells {
            orig.insert((c.row, c.col), OrigCell { style: c.style, array_ref: c.array_ref, formula: c.formula });
        }
        if parsed.protected {
            orig_protection = outline.child("sheetProtection").map(str::to_owned);
            orig_password = parsed.password.clone();
        }
    }

    // sheetData
    let cells = book.occupied_cells(id, Rect::sheet());
    let rows = book.custom_rows(id);
    let mut data = format!("<{p}sheetData>");
    let mut ci = 0usize;
    let mut ri = 0usize;
    loop {
        let next_cell_row = cells.get(ci).map(|c| c.0);
        let next_prop_row = rows.get(ri).copied();
        let r = match (next_cell_row, next_prop_row) {
            (None, None) => break,
            (Some(a), None) => a,
            (None, Some(b)) => b,
            (Some(a), Some(b)) => a.min(b),
        };
        let custom_row = next_prop_row == Some(r);
        if custom_row {
            ri += 1;
        }
        data.push_str(&format!("<{p}row r=\"{r}\""));
        let h = book.row_height(id, r);
        // A row with explicit properties keeps its height even when it equals the default.
        if custom_row || (h - macroost_sheet::memory::DEFAULT_ROW_HEIGHT).abs() > 1e-9 {
            data.push_str(&format!(" ht=\"{}\" customHeight=\"1\"", fmt_num(h)));
        }
        if book.row_hidden(id, r) {
            data.push_str(" hidden=\"1\"");
        }
        let start = ci;
        while cells.get(ci).is_some_and(|c| c.0 == r) {
            ci += 1;
        }
        if start == ci {
            data.push_str("/>");
            continue;
        }
        data.push('>');
        for &(_, c) in cells.get(start..ci).unwrap_or(&[]) {
            write_cell(book, id, r, c, orig.get(&(r, c)), styles, sst, date1904, &p, &mut data);
        }
        data.push_str(&format!("</{p}row>"));
    }
    data.push_str(&format!("</{p}sheetData>"));

    // cols
    let custom = book.custom_columns(id);
    let mut cols = String::new();
    let mut i = 0usize;
    while let Some(&c0) = custom.get(i) {
        let w = book.column_width(id, c0);
        let hid = book.column_hidden(id, c0);
        let mut j = i + 1;
        while let Some(&cj) = custom.get(j) {
            if cj != c0 + (j - i) as u32 || book.column_width(id, cj) != w || book.column_hidden(id, cj) != hid {
                break;
            }
            j += 1;
        }
        let c1 = c0 + (j - i - 1) as u32;
        cols.push_str(&format!(
            "<{p}col min=\"{c0}\" max=\"{c1}\" width=\"{}\" customWidth=\"1\"{}/>",
            fmt_num(width_to_file(w)),
            if hid { " hidden=\"1\"" } else { "" }
        ));
        i = j;
    }
    let cols = (!cols.is_empty()).then(|| format!("<{p}cols>{cols}</{p}cols>"));

    let merges = book.merged_areas(id);
    let merges = (!merges.is_empty()).then(|| {
        let mut s = format!("<{p}mergeCells count=\"{}\">", merges.len());
        for m in &merges {
            s.push_str(&format!("<{p}mergeCell ref=\"{}\"/>", area_ref(m)));
        }
        s.push_str(&format!("</{p}mergeCells>"));
        s
    });
    let password = book.sheet_password(id);
    let protection = book.is_protected(id).then(|| match orig_protection {
        Some(orig) if orig_password == password => orig,
        _ => format!(
            "<{p}sheetProtection{} sheet=\"1\" objects=\"1\" scenarios=\"1\"/>",
            crate::protection::hash_attrs(password.as_ref(), "password", "")
        ),
    });
    let dimension =
        Some(format!("<{p}dimension ref=\"{}\"/>", book.used_range(id).map_or("A1".to_owned(), |r| area_ref(&r))));

    // sheetPr carries the code name.
    let code = book.code_name(id);
    let sheet_pr = match (outline.child("sheetPr"), code) {
        (Some(raw), Some(cn)) => match xml::parse_tree(raw) {
            Ok(mut t) => {
                t.set_attr("codeName", &cn);
                Some(t.to_xml())
            }
            Err(_) => Some(raw.to_owned()),
        },
        (Some(raw), None) => Some(raw.to_owned()),
        (None, Some(cn)) => Some(format!("<{p}sheetPr codeName=\"{}\"/>", escape(&cn))),
        (None, None) => None,
    };
    Ok(outline.rebuild(
        SHEET_ORDER,
        &[
            ("sheetPr", sheet_pr),
            ("dimension", dimension),
            ("cols", cols),
            ("sheetData", Some(data)),
            ("sheetProtection", protection),
            ("mergeCells", merges),
        ],
    ))
}

#[allow(clippy::too_many_arguments)]
fn write_cell(
    book: &MemoryWorkbook,
    id: u32,
    r: u32,
    c: u32,
    orig: Option<&OrigCell>,
    styles: &mut Styles,
    sst: &mut SstOut,
    date1904: bool,
    p: &str,
    out: &mut String,
) {
    let format = book.cell_format(id, r, c);
    let s = styles.index_for(&format, orig.map(|o| o.style));
    let formula = book.cell_formula(id, r, c);
    let value = match &formula {
        // A formula keeps its cached result; without one (a formula set by a macro, or a file
        // saved without values) the built-in evaluator supplies it unless that fails.
        Some(_) => {
            let cached = book.cached_value(id, r, c);
            if cached.is_empty() {
                match book.cell_value(id, r, c) {
                    Value::Error(_) => Value::Empty,
                    v => v,
                }
            } else {
                cached
            }
        }
        None => book.cached_value(id, r, c),
    };
    let is_date = format.number_format.as_deref().is_some_and(is_date_format);
    let num = |x: f64| -> Option<f64> {
        if !x.is_finite() {
            return None;
        }
        Some(if date1904 && is_date { x - DATE1904_OFFSET } else { x })
    };
    // (t attribute, <v> text)
    let (t, v): (&str, Option<String>) = match &value {
        Value::Empty | Value::Null | Value::Object(_) | Value::Array(_) => ("", None),
        Value::Boolean(b) => ("b", Some(if *b { "1" } else { "0" }.to_owned())),
        Value::String(text) => {
            if formula.is_some() {
                ("str", Some(encode_xstring(text)))
            } else {
                ("s", Some(sst.get(text).to_string()))
            }
        }
        Value::Error(code) => ("e", Some(error_text(*code).to_owned())),
        other => match other.to_f64().ok().and_then(num) {
            Some(x) => ("", Some(fmt_num(x))),
            None => ("e", Some("#NUM!".to_owned())),
        },
    };
    out.push_str(&format!("<{p}c r=\"{}{r}\"", column_letters(c)));
    if s != 0 {
        out.push_str(&format!(" s=\"{s}\""));
    }
    if !t.is_empty() {
        out.push_str(&format!(" t=\"{t}\""));
    }
    if formula.is_none() && v.is_none() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    if let Some(f) = &formula {
        let text = f.strip_prefix('=').unwrap_or(f);
        match orig.filter(|o| o.formula.as_deref() == Some(text)).and_then(|o| o.array_ref.as_deref()) {
            Some(rf) => out.push_str(&format!("<{p}f t=\"array\" ref=\"{}\">", escape(rf))),
            None => out.push_str(&format!("<{p}f>")),
        }
        out.push_str(&encode_xstring(text));
        out.push_str(&format!("</{p}f>"));
    }
    if let Some(v) = v {
        out.push_str(&format!("<{p}v>{v}</{p}v>"));
    }
    out.push_str(&format!("</{p}c>"));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn widths_round_trip() {
        for w in [0.0, 1.0, 8.43, 10.71, 25.5, 100.0, 255.0] {
            assert_eq!(crate::read::width_from_file(width_to_file(w)), w, "{w}");
        }
        assert_eq!(width_to_file(8.43), 9.140625);
    }
}
