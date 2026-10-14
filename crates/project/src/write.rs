//! Writes source-only macro projects (MS-OVBA sections 2.2-2.3.4) and minimal macro-enabled
//! OOXML workbooks and documents around them.
//!
//! A source-only project carries no performance cache: `_VBA_PROJECT` has Version 0xFFFF (the
//! value the spec requires writers to use) and every module's MODULEOFFSET is 0, so the host has
//! to compile the project from the source text when it opens the file. That is what the oracle
//! harness needs: the same source text runs in our interpreter and in the reference host.

use std::io::{Cursor, Write};

use crate::{Error, compression};

/// Code page the writer encodes all MBCS text in (Windows-1252, PROJECTCODEPAGE 1252).
const CODE_PAGE: u16 = 1252;
/// CLSID written to the PROJECT stream's `ID=` line. Any GUID is valid for an unprotected
/// project; a fixed one keeps the output deterministic.
/// PROJECTSYSKIND value for 64-bit Windows; the host recompiles from source either way.
const SYS_KIND_WIN64: u32 = 3;
const PROJECT_CLSID: &str = "{917DED54-440B-4FD1-A5C1-74ACF261E600}";

/// Builds a `vbaProject.bin` compound file holding standard (procedural) modules with the
/// given names and source text. `Attribute VB_Name` is added when the source lacks it, and line
/// endings are normalized to CRLF.
pub fn project_bin(modules: &[(String, String)]) -> Result<Vec<u8>, Error> {
    project_bin_with_documents(modules, &[])
}

/// Like [`project_bin`], plus document modules (`ThisDocument`, `ThisWorkbook`, ...) whose
/// source should carry the host's `VB_Base`/`VB_PredeclaredId` attributes. Some hosts refuse a
/// project that lacks the document module of the file it is stored in.
pub fn project_bin_with_documents(
    modules: &[(String, String)],
    documents: &[(String, String)],
) -> Result<Vec<u8>, Error> {
    let all: Vec<(String, String, bool)> = documents
        .iter()
        .map(|(n, s)| (n.clone(), s.clone(), true))
        .chain(modules.iter().map(|(n, s)| (n.clone(), s.clone(), false)))
        .collect();
    build_project(&all, &[])
}

/// A registered type library the project references: (name, libid). The libid has the
/// [MS-OVBA] form `*\G{guid}#major.minor#lcid#path#description`.
pub type Reference<'a> = (&'a str, &'a str);

/// [`project_bin`] with registered library references (REFERENCEREGISTERED records).
pub fn project_bin_with_references(modules: &[(String, String)], refs: &[Reference<'_>]) -> Result<Vec<u8>, Error> {
    let all: Vec<(String, String, bool)> = modules.iter().map(|(n, s)| (n.clone(), s.clone(), false)).collect();
    build_project(&all, refs)
}

/// Builds the compound file for `modules` (name, source, is document module) and `refs`.
fn build_project(modules: &[(String, String, bool)], refs: &[Reference<'_>]) -> Result<Vec<u8>, Error> {
    check_modules(modules)?;
    let mut cfb = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))?;
    put_stream(&mut cfb, "/PROJECT", &project_stream(modules)?)?;
    put_stream(&mut cfb, "/PROJECTwm", &projectwm_stream(modules)?)?;
    cfb.create_storage("/VBA")?;
    put_stream(&mut cfb, "/VBA/_VBA_PROJECT", &version_stream())?;
    put_stream(&mut cfb, "/VBA/dir", &compression::compress(&dir_stream(modules, refs)?))?;
    for (name, source, _) in modules {
        let text = encode(&module_text(name, source))?;
        put_stream(&mut cfb, &format!("/VBA/{name}"), &compression::compress(&text))?;
    }
    cfb.flush()?;
    Ok(cfb.into_inner().into_inner())
}

/// Builds a minimal macro-enabled workbook (.xlsm) with empty worksheets named `sheet_names`
/// (at least one) and a source-only macro project holding `modules`.
pub fn xlsm_with_macros(modules: &[(String, String)], sheet_names: &[&str]) -> Result<Vec<u8>, Error> {
    if sheet_names.is_empty() {
        return Err(Error::Invalid("a workbook needs at least one sheet".into()));
    }
    for s in sheet_names {
        let bad = s.is_empty() || s.chars().count() > 31 || s.chars().any(|c| "[]:*?/\\".contains(c));
        if bad {
            return Err(Error::Invalid(format!("invalid sheet name {s:?}")));
        }
    }
    let bin = project_bin(modules)?;

    let mut types = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>"#,
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>"#,
        r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
        r#"<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>"#,
    ));
    let mut sheets = String::new();
    let mut wb_rels = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    ));
    for (i, name) in sheet_names.iter().enumerate() {
        let n = i + 1;
        types.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
        sheets.push_str(&format!(r#"<sheet name="{}" sheetId="{n}" r:id="rId{n}"/>"#, xml_escape(name)));
        wb_rels.push_str(&format!(
            r#"<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>"#
        ));
    }
    let n = sheet_names.len();
    wb_rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        n + 1
    ));
    wb_rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>"#,
        n + 2
    ));
    wb_rels.push_str("</Relationships>");
    types.push_str("</Types>");

    let root_rels = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
        "</Relationships>",
    );
    let workbook = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\r\n",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<workbookPr codeName="ThisWorkbook"/>"#,
            r#"<bookViews><workbookView/></bookViews>"#,
            "<sheets>{}</sheets></workbook>",
        ),
        sheets
    );
    let worksheet = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        "<sheetData/></worksheet>",
    );
    let styles = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
        r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
        r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
        r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
        r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
        "</styleSheet>",
    );

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut add = |name: &str, data: &[u8]| -> Result<(), Error> {
        zip.start_file(name, opts).map_err(zip_err)?;
        zip.write_all(data)?;
        Ok(())
    };
    add("[Content_Types].xml", types.as_bytes())?;
    add("_rels/.rels", root_rels.as_bytes())?;
    add("xl/workbook.xml", workbook.as_bytes())?;
    add("xl/_rels/workbook.xml.rels", wb_rels.as_bytes())?;
    for i in 1..=n {
        add(&format!("xl/worksheets/sheet{i}.xml"), worksheet.as_bytes())?;
    }
    add("xl/styles.xml", styles.as_bytes())?;
    add("xl/vbaProject.bin", &bin)?;
    let out = zip.finish().map_err(zip_err)?;
    Ok(out.into_inner())
}

/// Builds a minimal macro-enabled word-processing document (.docm) with one paragraph per entry
/// of `paragraphs` (an empty slice gives one empty paragraph) and a source-only macro project
/// holding `modules`. The package carries a small style part (document defaults Calibri 11 pt,
/// a `Normal` paragraph style) so the host's defaults do not depend on the user's template.
pub fn docm_with_macros(modules: &[(String, String)], paragraphs: &[&str]) -> Result<Vec<u8>, Error> {
    let this_document = concat!(
        "Attribute VB_Name = \"ThisDocument\"\n",
        "Attribute VB_Base = \"1Normal.ThisDocument\"\n",
        "Attribute VB_GlobalNameSpace = False\n",
        "Attribute VB_Creatable = False\n",
        "Attribute VB_PredeclaredId = True\n",
        "Attribute VB_Exposed = True\n",
        "Attribute VB_TemplateDerived = True\n",
        "Attribute VB_Customizable = True\n",
    );
    let bin = project_bin_with_documents(modules, &[("ThisDocument".into(), this_document.into())])?;
    let types = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>"#,
        r#"<Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/>"#,
        r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#,
        r#"<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>"#,
        "</Types>",
    );
    let root_rels = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#,
        "</Relationships>",
    );
    let doc_rels = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>"#,
        "</Relationships>",
    );
    let mut body = String::new();
    if paragraphs.is_empty() {
        body.push_str("<w:p/>");
    }
    for p in paragraphs {
        if p.chars().any(|c| c.is_control() && c != '\t') {
            return Err(Error::Invalid(format!("paragraph text {p:?} holds a control character")));
        }
        if p.is_empty() {
            body.push_str("<w:p/>");
        } else {
            body.push_str(&format!(r#"<w:p><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>"#, xml_escape(p)));
        }
    }
    let document = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\r\n",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            "<w:body>{}",
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#,
            "</w:body></w:document>",
        ),
        body
    );
    let styles = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\r\n",
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
        r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>"#,
        r#"<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>"#,
        r#"<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"#,
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>"#,
        r#"<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>"#,
        "</w:styles>",
    );

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut add = |name: &str, data: &[u8]| -> Result<(), Error> {
        zip.start_file(name, opts).map_err(zip_err)?;
        zip.write_all(data)?;
        Ok(())
    };
    add("[Content_Types].xml", types.as_bytes())?;
    add("_rels/.rels", root_rels.as_bytes())?;
    add("word/document.xml", document.as_bytes())?;
    add("word/_rels/document.xml.rels", doc_rels.as_bytes())?;
    add("word/styles.xml", styles.as_bytes())?;
    add("word/vbaProject.bin", &bin)?;
    let out = zip.finish().map_err(zip_err)?;
    Ok(out.into_inner())
}

/// Slide size of the generated presentation: 13.333 x 7.5 in (960 x 540 points), 16:9.
pub const PPTM_SLIDE_CX: i64 = 12_192_000;
pub const PPTM_SLIDE_CY: i64 = 6_858_000;

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_OFFICE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Slide layouts of the generated master: (layout type, name, placeholders as (type, idx)).
/// Layout placeholders without geometry inherit it from the master's title/body placeholders.
const PPTM_LAYOUTS: &[(&str, &str, &[(&str, u32)])] = &[
    ("title", "Title Slide", &[("ctrTitle", 0), ("subTitle", 1)]),
    ("obj", "Title and Content", &[("title", 0), ("obj", 1)]),
    ("tx", "Title and Text", &[("title", 0), ("body", 1)]),
    ("titleOnly", "Title Only", &[("title", 0)]),
    ("blank", "Blank", &[]),
];
/// Index (1-based) of the blank layout in [`PPTM_LAYOUTS`]; generated slides use it.
const PPTM_BLANK_LAYOUT: usize = 5;

fn rels(items: &[(String, String, String)]) -> String {
    let mut s = format!("{XML_DECL}<Relationships xmlns=\"{REL_NS}\">");
    for (id, ty, target) in items {
        s.push_str(&format!("<Relationship Id=\"{id}\" Type=\"{ty}\" Target=\"{target}\"/>"));
    }
    s.push_str("</Relationships>");
    s
}

fn ooxml_rel(id: usize, kind: &str, target: &str) -> (String, String, String) {
    (format!("rId{id}"), format!("{REL_OFFICE}/{kind}"), target.to_owned())
}

/// Text style levels 1..9 with the given size (hundredths of a point) and font.
fn level_styles(sz: u32, font: &str, bullets: bool) -> String {
    let mut s = String::new();
    for lvl in 1..=9u32 {
        let (mar, ind, bullet) = if bullets {
            (
                228_600 + (lvl - 1) * 457_200,
                " indent=\"-228600\"",
                "<a:buFont typeface=\"Arial\"/><a:buChar char=\"&#8226;\"/>",
            )
        } else {
            ((lvl - 1) * 457_200, "", "")
        };
        s.push_str(&format!(
            "<a:lvl{lvl}pPr marL=\"{mar}\"{ind} algn=\"l\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">{bullet}<a:defRPr sz=\"{sz}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"{font}\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{lvl}pPr>"
        ));
    }
    s
}

fn theme_xml() -> String {
    let clr = |name: &str, rgb: &str| format!("<a:{name}><a:srgbClr val=\"{rgb}\"/></a:{name}>");
    let solid = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
    let line = |w: u32| {
        format!(
            "<a:ln w=\"{w}\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">{solid}<a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
        )
    };
    let mut s = format!("{XML_DECL}<a:theme xmlns:a=\"{NS_A}\" name=\"Default Theme\"><a:themeElements>");
    s.push_str("<a:clrScheme name=\"Default\">");
    s.push_str("<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>");
    s.push_str("<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>");
    for (name, rgb) in [
        ("dk2", "44546A"),
        ("lt2", "E7E6E6"),
        ("accent1", "4472C4"),
        ("accent2", "ED7D31"),
        ("accent3", "A5A5A5"),
        ("accent4", "FFC000"),
        ("accent5", "5B9BD5"),
        ("accent6", "70AD47"),
        ("hlink", "0563C1"),
        ("folHlink", "954F72"),
    ] {
        s.push_str(&clr(name, rgb));
    }
    s.push_str("</a:clrScheme><a:fontScheme name=\"Default\">");
    s.push_str(
        "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>",
    );
    s.push_str("<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>");
    s.push_str("</a:fontScheme><a:fmtScheme name=\"Default\">");
    s.push_str(&format!("<a:fillStyleLst>{solid}{solid}{solid}</a:fillStyleLst>"));
    s.push_str(&format!("<a:lnStyleLst>{}{}{}</a:lnStyleLst>", line(6350), line(12700), line(19050)));
    s.push_str("<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>");
    s.push_str(&format!("<a:bgFillStyleLst>{solid}{solid}{solid}</a:bgFillStyleLst>"));
    s.push_str("</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
    s
}

fn placeholder_sp(id: u32, name: &str, ph_type: &str, idx: u32, geometry: Option<(i64, i64, i64, i64)>) -> String {
    let idx_attr = if idx > 0 { format!(" idx=\"{idx}\"") } else { String::new() };
    // Type "obj" is the default and is written by omitting the attribute.
    let type_attr = if ph_type == "obj" { String::new() } else { format!(" type=\"{ph_type}\"") };
    let sp_pr = match geometry {
        Some((x, y, cx, cy)) => format!(
            "<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>"
        ),
        None => "<p:spPr/>".to_owned(),
    };
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph{type_attr}{idx_attr}/></p:nvPr></p:nvSpPr>{sp_pr}<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>"
    )
}

const GROUP_HEAD: &str = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";

fn master_xml() -> String {
    let title = placeholder_sp(2, "Title Placeholder 1", "title", 0, Some((838_200, 365_125, 10_515_600, 1_325_563)));
    let body = placeholder_sp(3, "Text Placeholder 2", "body", 1, Some((838_200, 1_825_625, 10_515_600, 4_351_338)));
    let mut ids = String::new();
    for i in 0..PPTM_LAYOUTS.len() {
        ids.push_str(&format!("<p:sldLayoutId id=\"{}\" r:id=\"rId{}\"/>", 2_147_483_649u64 + i as u64, i + 1));
    }
    let mut s = format!("{XML_DECL}<p:sldMaster xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\">");
    s.push_str("<p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>");
    s.push_str(&format!("<p:spTree>{GROUP_HEAD}{title}{body}</p:spTree></p:cSld>"));
    s.push_str("<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>");
    s.push_str(&format!("<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>"));
    s.push_str(&format!(
        "<p:txStyles><p:titleStyle>{}</p:titleStyle><p:bodyStyle>{}</p:bodyStyle><p:otherStyle>{}</p:otherStyle></p:txStyles>",
        level_styles(4400, "+mj-lt", false),
        level_styles(2800, "+mn-lt", true),
        level_styles(1800, "+mn-lt", false)
    ));
    s.push_str("</p:sldMaster>");
    s
}

fn layout_xml(ty: &str, name: &str, phs: &[(&str, u32)]) -> String {
    let mut shapes = String::new();
    for (i, (ph, idx)) in phs.iter().enumerate() {
        let id = i as u32 + 2;
        let label = match *ph {
            "ctrTitle" | "title" => "Title",
            "subTitle" => "Subtitle",
            _ => "Content Placeholder",
        };
        let geometry = match *ph {
            "ctrTitle" => Some((1_524_000, 1_122_363, 9_144_000, 2_387_600)),
            "subTitle" => Some((1_524_000, 3_602_038, 9_144_000, 1_655_762)),
            _ => None,
        };
        shapes.push_str(&placeholder_sp(id, &format!("{label} {}", id - 1), ph, *idx, geometry));
    }
    format!(
        "{XML_DECL}<p:sldLayout xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\" type=\"{ty}\" preserve=\"1\"><p:cSld name=\"{name}\"><p:spTree>{GROUP_HEAD}{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    )
}

/// Builds a minimal macro-enabled presentation (.pptm): one slide master with a theme and five
/// layouts (title, title and content, title and text, title only, blank), `slide_count` blank
/// slides (0 allowed), 960 x 540 point slides, and a source-only macro project holding `modules`
/// that references the type libraries `refs` (hosts that need a shared library for their
/// constants get it here).
pub fn pptm_with_macros(
    modules: &[(String, String)],
    slide_count: usize,
    refs: &[Reference<'_>],
) -> Result<Vec<u8>, Error> {
    if slide_count > 1000 {
        return Err(Error::Invalid("too many slides".into()));
    }
    let bin = project_bin_with_references(modules, refs)?;
    let ct = |part: &str, ty: &str| format!("<Override PartName=\"{part}\" ContentType=\"{ty}\"/>");
    let pml = "application/vnd.openxmlformats-officedocument.presentationml";
    let mut types = format!(
        "{XML_DECL}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>"
    );
    types.push_str(&ct("/ppt/presentation.xml", "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"));
    types.push_str(&ct("/ppt/slideMasters/slideMaster1.xml", &format!("{pml}.slideMaster+xml")));
    for i in 1..=PPTM_LAYOUTS.len() {
        types.push_str(&ct(&format!("/ppt/slideLayouts/slideLayout{i}.xml"), &format!("{pml}.slideLayout+xml")));
    }
    for i in 1..=slide_count {
        types.push_str(&ct(&format!("/ppt/slides/slide{i}.xml"), &format!("{pml}.slide+xml")));
    }
    types.push_str(&ct("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"));
    types.push_str(&ct("/ppt/presProps.xml", &format!("{pml}.presProps+xml")));
    types.push_str(&ct("/ppt/viewProps.xml", &format!("{pml}.viewProps+xml")));
    types.push_str(&ct("/ppt/tableStyles.xml", &format!("{pml}.tableStyles+xml")));
    types.push_str(&ct("/ppt/vbaProject.bin", "application/vnd.ms-office.vbaProject"));
    types.push_str("</Types>");

    let root_rels = rels(&[ooxml_rel(1, "officeDocument", "ppt/presentation.xml")]);
    let mut pres_rels =
        vec![ooxml_rel(1, "slideMaster", "slideMasters/slideMaster1.xml"), ooxml_rel(2, "theme", "theme/theme1.xml")];
    let mut sld_ids = String::new();
    for i in 1..=slide_count {
        pres_rels.push(ooxml_rel(i + 2, "slide", &format!("slides/slide{i}.xml")));
        sld_ids.push_str(&format!("<p:sldId id=\"{}\" r:id=\"rId{}\"/>", 255 + i, i + 2));
    }
    let n = slide_count + 2;
    pres_rels.push(ooxml_rel(n + 1, "presProps", "presProps.xml"));
    pres_rels.push(ooxml_rel(n + 2, "viewProps", "viewProps.xml"));
    pres_rels.push(ooxml_rel(n + 3, "tableStyles", "tableStyles.xml"));
    pres_rels.push((
        format!("rId{}", n + 4),
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject".to_owned(),
        "vbaProject.bin".to_owned(),
    ));
    let sld_id_lst = if slide_count > 0 { format!("<p:sldIdLst>{sld_ids}</p:sldIdLst>") } else { String::new() };
    let presentation = format!(
        "{XML_DECL}<p:presentation xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\" saveSubsetFonts=\"1\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>{sld_id_lst}<p:sldSz cx=\"{PPTM_SLIDE_CX}\" cy=\"{PPTM_SLIDE_CY}\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/><p:defaultTextStyle>{}</p:defaultTextStyle></p:presentation>",
        level_styles(1800, "+mn-lt", false)
    );
    let mut master_rels: Vec<_> = (1..=PPTM_LAYOUTS.len())
        .map(|i| ooxml_rel(i, "slideLayout", &format!("../slideLayouts/slideLayout{i}.xml")))
        .collect();
    master_rels.push(ooxml_rel(PPTM_LAYOUTS.len() + 1, "theme", "../theme/theme1.xml"));
    let layout_rels = rels(&[ooxml_rel(1, "slideMaster", "../slideMasters/slideMaster1.xml")]);
    let slide_rels =
        rels(&[ooxml_rel(1, "slideLayout", &format!("../slideLayouts/slideLayout{PPTM_BLANK_LAYOUT}.xml"))]);
    let slide = format!(
        "{XML_DECL}<p:sld xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\"><p:cSld><p:spTree>{GROUP_HEAD}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    );
    let pres_props = format!("{XML_DECL}<p:presentationPr xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\"/>");
    let view_props = format!(
        "{XML_DECL}<p:viewPr xmlns:a=\"{NS_A}\" xmlns:r=\"{NS_R}\" xmlns:p=\"{NS_P}\"><p:normalViewPr><p:restoredLeft sz=\"15620\"/><p:restoredTop sz=\"94660\"/></p:normalViewPr><p:gridSpacing cx=\"76200\" cy=\"76200\"/></p:viewPr>"
    );
    let table_styles =
        format!("{XML_DECL}<a:tblStyleLst xmlns:a=\"{NS_A}\" def=\"{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}\"/>");

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut add = |name: &str, data: &[u8]| -> Result<(), Error> {
        zip.start_file(name, opts).map_err(zip_err)?;
        zip.write_all(data)?;
        Ok(())
    };
    add("[Content_Types].xml", types.as_bytes())?;
    add("_rels/.rels", root_rels.as_bytes())?;
    add("ppt/presentation.xml", presentation.as_bytes())?;
    add("ppt/_rels/presentation.xml.rels", rels(&pres_rels).as_bytes())?;
    add("ppt/slideMasters/slideMaster1.xml", master_xml().as_bytes())?;
    add("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels(&master_rels).as_bytes())?;
    for (i, (ty, name, phs)) in PPTM_LAYOUTS.iter().enumerate() {
        add(&format!("ppt/slideLayouts/slideLayout{}.xml", i + 1), layout_xml(ty, name, phs).as_bytes())?;
        add(&format!("ppt/slideLayouts/_rels/slideLayout{}.xml.rels", i + 1), layout_rels.as_bytes())?;
    }
    for i in 1..=slide_count {
        add(&format!("ppt/slides/slide{i}.xml"), slide.as_bytes())?;
        add(&format!("ppt/slides/_rels/slide{i}.xml.rels"), slide_rels.as_bytes())?;
    }
    add("ppt/theme/theme1.xml", theme_xml().as_bytes())?;
    add("ppt/presProps.xml", pres_props.as_bytes())?;
    add("ppt/viewProps.xml", view_props.as_bytes())?;
    add("ppt/tableStyles.xml", table_styles.as_bytes())?;
    add("ppt/vbaProject.bin", &bin)?;
    let out = zip.finish().map_err(zip_err)?;
    Ok(out.into_inner())
}

fn zip_err(e: zip::result::ZipError) -> Error {
    Error::Invalid(format!("zip: {e}"))
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn check_modules(modules: &[(String, String, bool)]) -> Result<(), Error> {
    if modules.is_empty() {
        return Err(Error::Invalid("a macro project needs at least one module".into()));
    }
    if modules.len() > usize::from(u16::MAX) {
        return Err(Error::Invalid("too many modules".into()));
    }
    for (i, (name, _, _)) in modules.iter().enumerate() {
        let mut chars = name.chars();
        let ident = chars.next().is_some_and(|c| c.is_ascii_alphabetic())
            && chars.all(|c| c.is_ascii_alphanumeric() || c == '_')
            && name.len() <= 31;
        if !ident {
            return Err(Error::Invalid(format!("module name {name:?} is not a valid identifier")));
        }
        if modules[..i].iter().any(|(n, _, _)| n.eq_ignore_ascii_case(name)) {
            return Err(Error::Invalid(format!("duplicate module name {name:?}")));
        }
    }
    Ok(())
}

fn put_stream<F: std::io::Read + Write + std::io::Seek>(
    cfb: &mut cfb::CompoundFile<F>,
    path: &str,
    data: &[u8],
) -> Result<(), Error> {
    let mut s = cfb.create_stream(path)?;
    s.write_all(data)?;
    s.flush()?;
    Ok(())
}

/// Text in the project code page; characters it cannot represent are an error rather than
/// silently replaced.
fn encode(text: &str) -> Result<Vec<u8>, Error> {
    let (bytes, _, unmappable) = encoding_rs::WINDOWS_1252.encode(text);
    if unmappable {
        return Err(Error::Invalid("text is not representable in code page 1252".into()));
    }
    Ok(bytes.into_owned())
}

fn utf16le(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// The module text as stored: `Attribute VB_Name` first, CRLF line endings, a final newline.
pub fn module_text(name: &str, source: &str) -> String {
    let normalized = source.replace("\r\n", "\n").replace('\r', "\n");
    let mut out = String::new();
    let has_name =
        normalized.lines().next().is_some_and(|l| l.trim_start().to_ascii_lowercase().starts_with("attribute vb_name"));
    if !has_name {
        out.push_str(&format!("Attribute VB_Name = \"{name}\"\r\n"));
    }
    for line in normalized.lines() {
        out.push_str(line);
        out.push_str("\r\n");
    }
    out
}

/// Data Encryption (section 2.4.3.2) with a fixed seed and fixed ignored bytes.
fn encrypt(seed: u8, data: &[u8]) -> String {
    const VERSION: u8 = 2;
    let proj_key = PROJECT_CLSID.bytes().fold(0u8, u8::wrapping_add);
    let version_enc = seed ^ VERSION;
    let proj_key_enc = seed ^ proj_key;
    let mut out = vec![seed, version_enc, proj_key_enc];
    let (mut unenc1, mut enc1, mut enc2) = (proj_key, proj_key_enc, version_enc);
    let mut push = |byte: u8, out: &mut Vec<u8>| {
        let e = byte ^ enc2.wrapping_add(unenc1);
        out.push(e);
        enc2 = enc1;
        enc1 = e;
        unenc1 = byte;
    };
    let ignored_len = (seed & 6) / 2;
    for _ in 0..ignored_len {
        push(0x07, &mut out);
    }
    for b in (data.len() as u32).to_le_bytes() {
        push(b, &mut out);
    }
    for &b in data {
        push(b, &mut out);
    }
    out.iter().map(|b| format!("{b:02X}")).collect()
}

/// The PROJECT stream (section 2.3.1).
fn project_stream(modules: &[(String, String, bool)]) -> Result<Vec<u8>, Error> {
    let mut t = format!("ID=\"{PROJECT_CLSID}\"\r\n");
    for (name, _, document) in modules {
        if *document {
            t.push_str(&format!("Document={name}/&H00000000\r\n"));
        } else {
            t.push_str(&format!("Module={name}\r\n"));
        }
    }
    t.push_str("Name=\"VBAProject\"\r\n");
    t.push_str("HelpContextID=\"0\"\r\n");
    t.push_str("VersionCompatible32=\"393222000\"\r\n");
    // Unprotected (CMG data 0x00000000), no password (DPB data 0x00), visible (GC data 0xFF).
    t.push_str(&format!("CMG=\"{}\"\r\n", encrypt(0x07, &[0, 0, 0, 0])));
    t.push_str(&format!("DPB=\"{}\"\r\n", encrypt(0x0E, &[0])));
    t.push_str(&format!("GC=\"{}\"\r\n", encrypt(0x15, &[0xFF])));
    t.push_str("\r\n[Host Extender Info]\r\n");
    t.push_str("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
    t.push_str("\r\n[Workspace]\r\n");
    for (name, _, _) in modules {
        t.push_str(&format!("{name}=0, 0, 0, 0, C\r\n"));
    }
    encode(&t)
}

/// The PROJECTwm stream (section 2.3.3): MBCS name, NUL, UTF-16 name, NUL NUL; then 0x0000.
fn projectwm_stream(modules: &[(String, String, bool)]) -> Result<Vec<u8>, Error> {
    let mut out = Vec::new();
    for (name, _, _) in modules {
        out.extend(encode(name)?);
        out.push(0);
        out.extend(utf16le(name));
        out.extend_from_slice(&[0, 0]);
    }
    out.extend_from_slice(&[0, 0]);
    Ok(out)
}

/// The `_VBA_PROJECT` stream (section 2.3.4.1) with no performance cache.
fn version_stream() -> Vec<u8> {
    let mut out = Vec::with_capacity(7);
    out.extend_from_slice(&0x61CCu16.to_le_bytes()); // Reserved1
    out.extend_from_slice(&0xFFFFu16.to_le_bytes()); // Version: MUST be 0xFFFF on write
    out.push(0x00); // Reserved2
    out.extend_from_slice(&0x0000u16.to_le_bytes()); // Reserved3 (undefined)
    out
}

fn rec(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

/// The decompressed `dir` stream (section 2.3.4.2), records in spec order.
fn dir_stream(modules: &[(String, String, bool)], refs: &[Reference<'_>]) -> Result<Vec<u8>, Error> {
    let mut d = Vec::new();
    // PROJECTINFORMATION (2.3.4.2.1).
    rec(&mut d, 0x0001, &SYS_KIND_WIN64.to_le_bytes()); // PROJECTSYSKIND
    rec(&mut d, 0x0002, &0x0409u32.to_le_bytes()); // PROJECTLCID
    rec(&mut d, 0x0014, &0x0409u32.to_le_bytes()); // PROJECTLCIDINVOKE
    rec(&mut d, 0x0003, &CODE_PAGE.to_le_bytes()); // PROJECTCODEPAGE
    rec(&mut d, 0x0004, b"VBAProject"); // PROJECTNAME
    rec(&mut d, 0x0005, b""); // PROJECTDOCSTRING
    rec(&mut d, 0x0040, b""); //   DocStringUnicode (Reserved 0x0040)
    rec(&mut d, 0x0006, b""); // PROJECTHELPFILEPATH
    rec(&mut d, 0x003D, b""); //   HelpFile2 (Reserved 0x003D)
    rec(&mut d, 0x0007, &0u32.to_le_bytes()); // PROJECTHELPCONTEXT
    rec(&mut d, 0x0008, &0u32.to_le_bytes()); // PROJECTLIBFLAGS
    // PROJECTVERSION: Reserved (where a size would be) is 4, yet six bytes follow.
    d.extend_from_slice(&0x0009u16.to_le_bytes());
    d.extend_from_slice(&4u32.to_le_bytes());
    d.extend_from_slice(&1u32.to_le_bytes()); // VersionMajor
    d.extend_from_slice(&0u16.to_le_bytes()); // VersionMinor
    rec(&mut d, 0x000C, b""); // PROJECTCONSTANTS
    rec(&mut d, 0x003C, b""); //   ConstantsUnicode (Reserved 0x003C)
    // PROJECTREFERENCES (2.3.4.2.2): the host supplies its own library implicitly; others are
    // REFERENCENAME + REFERENCEREGISTERED.
    for (name, libid) in refs {
        rec(&mut d, 0x0016, &encode(name)?); // REFERENCENAME
        rec(&mut d, 0x003E, &utf16le(name)); //   NameUnicode (Reserved 0x003E)
        let lib = encode(libid)?;
        let mut body = Vec::with_capacity(lib.len() + 10);
        body.extend_from_slice(&(lib.len() as u32).to_le_bytes()); // SizeOfLibid
        body.extend_from_slice(&lib);
        body.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        body.extend_from_slice(&0u16.to_le_bytes()); // Reserved2
        rec(&mut d, 0x000D, &body); // REFERENCEREGISTERED
    }
    // PROJECTMODULES (2.3.4.2.3).
    // PROJECTMODULES (2.3.4.2.3).
    rec(&mut d, 0x000F, &(modules.len() as u16).to_le_bytes());
    rec(&mut d, 0x0013, &0xFFFFu16.to_le_bytes()); // PROJECTCOOKIE
    for (name, _, document) in modules {
        let mbcs = encode(name)?;
        let wide = utf16le(name);
        rec(&mut d, 0x0019, &mbcs); // MODULENAME
        rec(&mut d, 0x0047, &wide); // MODULENAMEUNICODE
        rec(&mut d, 0x001A, &mbcs); // MODULESTREAMNAME
        rec(&mut d, 0x0032, &wide); //   StreamNameUnicode (Reserved 0x0032)
        rec(&mut d, 0x001C, b""); // MODULEDOCSTRING
        rec(&mut d, 0x0048, b""); //   DocStringUnicode (Reserved 0x0048)
        rec(&mut d, 0x0031, &0u32.to_le_bytes()); // MODULEOFFSET: source starts at 0
        rec(&mut d, 0x001E, &0u32.to_le_bytes()); // MODULEHELPCONTEXT
        rec(&mut d, 0x002C, &0xFFFFu16.to_le_bytes()); // MODULECOOKIE
        // MODULETYPE: procedural (0x0021) or document/class (0x0022).
        rec(&mut d, if *document { 0x0022 } else { 0x0021 }, b"");
        rec(&mut d, 0x002B, b""); // Terminator + Reserved
    }
    rec(&mut d, 0x0010, b""); // dir Terminator + Reserved
    Ok(d)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn mods() -> Vec<(String, String)> {
        vec![
            ("Module1".into(), "Function T1()\n    T1 = 300& * 200\nEnd Function\n".into()),
            ("Helpers".into(), "Attribute VB_Name = \"Helpers\"\r\nSub S()\r\nEnd Sub".into()),
        ]
    }

    #[test]
    fn encryption_matches_the_spec_example() {
        // Section 3 example: same CLSID, seeds 0x07 / 0x0E / 0x15.
        assert_eq!(encrypt(0x07, &[0, 0, 0, 0]), "0705D8E3D8EDDBF1DBF1DBF1DBF1");
        assert_eq!(encrypt(0x0E, &[0]), "0E0CD1ECDFF4E7F5E7F5E7");
        assert_eq!(encrypt(0x15, &[0xFF]), "1517CAF1D6F9D7F9D706");
    }

    #[test]
    fn project_round_trips_through_our_reader() {
        let m = mods();
        let bin = project_bin(&m).unwrap();
        let p = crate::read(&bin).unwrap();
        assert_eq!(p.name, "VBAProject");
        assert_eq!(p.code_page, 1252);
        assert_eq!(p.modules.len(), 2);
        for ((name, src), got) in m.iter().zip(&p.modules) {
            assert_eq!(&got.name, name);
            assert_eq!(got.kind, crate::ModuleKind::Standard);
            assert_eq!(got.source, module_text(name, src));
        }
        assert!(p.modules[0].source.starts_with("Attribute VB_Name = \"Module1\"\r\nFunction T1()\r\n"));
        assert_eq!(p.modules[1].source, "Attribute VB_Name = \"Helpers\"\r\nSub S()\r\nEnd Sub\r\n");
    }

    #[test]
    fn xlsm_round_trips_through_our_reader() {
        let m = mods();
        let x = xlsm_with_macros(&m, &["Sheet1", "Results"]).unwrap();
        let p = crate::read(&x).unwrap();
        assert_eq!(p.modules.len(), 2);
        assert_eq!(p.modules[1].source, module_text("Helpers", &m[1].1));
    }

    #[test]
    fn version_stream_forces_recompile() {
        assert_eq!(version_stream(), [0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn bad_input_is_an_error() {
        assert!(project_bin(&[]).is_err());
        assert!(project_bin(&[("1bad".into(), String::new())]).is_err());
        assert!(project_bin(&[("A".into(), String::new()), ("a".into(), String::new())]).is_err());
        assert!(project_bin(&[("A".into(), "x = \"\u{4e2d}\"".into())]).is_err());
        assert!(xlsm_with_macros(&mods(), &[]).is_err());
        assert!(docm_with_macros(&mods(), &["a\rb"]).is_err());
        assert!(pptm_with_macros(&mods(), 5000, &[]).is_err());
    }

    #[test]
    fn docm_round_trips_through_our_reader() {
        let m = mods();
        let d = docm_with_macros(&m, &["Hello & <world>", ""]).unwrap();
        let p = crate::read(&d).unwrap();
        assert_eq!(p.modules.len(), 3);
        assert_eq!(p.modules[0].name, "ThisDocument");
        assert_eq!(p.modules[0].kind, crate::ModuleKind::Document);
        assert_eq!(p.modules[1].source, module_text("Module1", &m[0].1));
        let mut z = zip::ZipArchive::new(Cursor::new(d)).unwrap();
        let mut doc = String::new();
        std::io::Read::read_to_string(&mut z.by_name("word/document.xml").unwrap(), &mut doc).unwrap();
        assert!(doc.contains("Hello &amp; &lt;world&gt;"));
    }

    #[test]
    fn pptm_round_trips_through_our_reader() {
        let m = mods();
        let x = pptm_with_macros(&m, 3, &[]).unwrap();
        let p = crate::read(&x).unwrap();
        assert_eq!(p.modules.len(), 2);
        assert_eq!(p.modules[0].source, module_text("Module1", &m[0].1));
        let mut z = zip::ZipArchive::new(Cursor::new(x)).unwrap();
        assert!(z.by_name("ppt/slides/slide3.xml").is_ok());
        assert!(z.by_name("ppt/slides/slide4.xml").is_err());
        assert!(pptm_with_macros(&m, 0, &[]).is_ok());
    }
}
