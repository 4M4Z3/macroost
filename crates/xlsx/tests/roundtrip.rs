//! Hand-written packages: load, save (with and without template), reload.
#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use std::io::{Cursor, Read, Write};

use macroost_lang::Value;
use macroost_sheet::{BorderLine, CellFormat, HAlign, MemoryWorkbook, Rect, SheetVisibility, Workbook};
use macroost_xlsx::{load, save};

const NS: &str = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn zip_of(parts: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut z = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (n, d) in parts {
        z.start_file(*n, o).unwrap();
        z.write_all(d).unwrap();
    }
    z.finish().unwrap().into_inner()
}

fn part(bytes: &[u8], name: &str) -> Option<Vec<u8>> {
    let mut z = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut f = z.by_name(name).ok()?;
    let mut v = Vec::new();
    f.read_to_end(&mut v).unwrap();
    Some(v)
}

fn part_text(bytes: &[u8], name: &str) -> String {
    String::from_utf8(part(bytes, name).unwrap_or_else(|| panic!("no part {name}"))).unwrap()
}

fn sample_parts(date1904: bool) -> Vec<(&'static str, Vec<u8>)> {
    let ct = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"#;
    let root_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let wb = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook {NS}><workbookPr codeName="ThisWorkbook"{}/><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden &amp; Co" sheetId="2" state="hidden" r:id="rId2"/><sheet name="VH" sheetId="5" state="veryHidden" r:id="rId3"/></sheets><definedNames><definedName name="Total">Data!$B$1:$B$5</definedName><definedName name="Local" localSheetId="1" hidden="1">'Hidden &amp; Co'!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$E$5</definedName></definedNames><calcPr calcId="191029"/></workbook>"#,
        if date1904 { " date1904=\"1\"" } else { "" }
    );
    let wb_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{REL}/worksheet" Target="/xl/worksheets/sheet3.xml"/><Relationship Id="rId4" Type="{REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="{REL}/styles" Target="styles.xml"/><Relationship Id="rId6" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId7" Type="{REL}/calcChain" Target="calcChain.xml"/><Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>"#
    );
    let sst = r#"<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3"><si><t>Hello</t></si><si><r><rPr><b/></rPr><t>Rich </t></r><r><t xml:space="preserve">text</t></r><rPh sb="0" eb="1"><t>IGNORED</t></rPh></si><si><t>a_x000D_b &amp; &lt;c&gt;</t></si></sst>"#;
    let theme = r#"<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#;
    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts count="3"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><i/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font><font><sz val="11"/><color theme="4"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF0000FF"/></right><top style="dashed"/><bottom style="double"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="1" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" wrapText="1"/></xf><xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font></dxf></dxfs><tableStyles count="0"/></styleSheet>"#;
    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><sheetPr codeName="Sheet12"><tabColor rgb="FFFF0000"/></sheetPr><dimension ref="A1:E11"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><cols><col min="2" max="2" width="20.7109375" customWidth="1"/><col min="4" max="5" width="9.140625" hidden="1"/></cols><sheetData><row r="1" spans="1:5"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#DIV/0!</v></c><c r="E1" t="inlineStr"><is><t>inline</t></is></c></row><row r="2"><c r="A2" s="1"><v>46309</v></c><c r="B2"><f>B1*2</f><v>84</v></c><c r="C2" t="str"><f>"x"&amp;A1</f><v>xHello</v></c><c r="D2" s="2"><v>1234.5</v></c><c r="E2" s="4"><v>46309.75</v></c></row><row r="3" ht="30" customHeight="1"><c r="B3"><f t="shared" ref="B3:B5" si="0">B1+1</f><v>43</v></c><c r="C3" t="s"><v>1</v></c><c r="D3" t="s"><v>2</v></c><c r="E3" s="3"><v>3.14159</v></c></row><row r="4" hidden="1"><c r="B4"><f t="shared" si="0"/><v>85</v></c></row><row r="5"><c r="B5"><f t="shared" si="0"/><v>44</v></c><c r="C5"><f t="array" ref="C5:C6">SUM(B1:B2*1)</f><v>126</v></c></row><row r="7"><c r="A7" s="2"/></row></sheetData><sheetProtection algorithmName="SHA-512" hashValue="abc" saltValue="def" spinCount="100000" sheet="1" objects="1" scenarios="1"/><mergeCells count="1"><mergeCell ref="A10:C11"/></mergeCells><conditionalFormatting sqref="B1:B5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>40</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"a,b"</formula1></dataValidation></dataValidations><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{1}"><x14:dataBar><x14:cfvo type="autoMin"/></x14:dataBar></x14:cfRule><xm:sqref>B1:B5</xm:sqref><xm:f>Data!$A$1</xm:f></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#;
    let sheet1_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#
    );
    // A prefixed main namespace, no r attributes on cells.
    let sheet2 = r#"<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row><x:c><x:v>7</x:v></x:c><x:c t="s"><x:v>0</x:v></x:c></x:row></x:sheetData></x:worksheet>"#;
    let sheet3 = format!(r#"<worksheet {NS}><sheetData/></worksheet>"#);
    let calc =
        r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B2" i="1"/></calcChain>"#;
    let drawing = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;
    vec![
        ("[Content_Types].xml", ct.as_bytes().to_vec()),
        ("_rels/.rels", root_rels.into_bytes()),
        ("xl/workbook.xml", wb.into_bytes()),
        ("xl/_rels/workbook.xml.rels", wb_rels.into_bytes()),
        ("xl/sharedStrings.xml", sst.as_bytes().to_vec()),
        ("xl/theme/theme1.xml", theme.as_bytes().to_vec()),
        ("xl/styles.xml", styles.as_bytes().to_vec()),
        ("xl/worksheets/sheet1.xml", sheet1.as_bytes().to_vec()),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet1_rels.into_bytes()),
        ("xl/worksheets/sheet2.xml", sheet2.as_bytes().to_vec()),
        ("xl/worksheets/sheet3.xml", sheet3.into_bytes()),
        ("xl/calcChain.xml", calc.as_bytes().to_vec()),
        ("xl/drawings/drawing1.xml", drawing.as_bytes().to_vec()),
        ("xl/vbaProject.bin", vec![0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4, 5]),
    ]
}

/// Saves and checks the package is consistent.
fn saved(book: &MemoryWorkbook, template: Option<&[u8]>) -> Vec<u8> {
    let out = save(book, template).unwrap();
    let problems = macroost_xlsx::validate(&out).unwrap();
    assert!(problems.is_empty(), "{problems:?}");
    out
}

fn sample() -> Vec<u8> {
    zip_of(&sample_parts(false))
}

fn s(v: &str) -> Value {
    Value::from(v)
}

fn assert_val(book: &MemoryWorkbook, sheet: u32, r: u32, c: u32, want: Value) {
    let got = book.cached_value(sheet, r, c);
    assert!(same(&got, &want), "cell ({r},{c}): got {got:?}, want {want:?}");
}

fn same(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Double(x), Value::Double(y)) => x == y,
        (Value::String(x), Value::String(y)) => x == y,
        (Value::Boolean(x), Value::Boolean(y)) => x == y,
        (Value::Error(x), Value::Error(y)) => x == y,
        (Value::Empty, Value::Empty) => true,
        _ => false,
    }
}

fn check_sample(book: &MemoryWorkbook) {
    let ids = book.sheet_ids();
    assert_eq!(ids.len(), 3);
    let names: Vec<String> = ids.iter().map(|&i| book.sheet_name(i).unwrap()).collect();
    assert_eq!(names, ["Data", "Hidden & Co", "VH"]);
    assert_eq!(book.sheet_visibility(ids[1]), SheetVisibility::Hidden);
    assert_eq!(book.sheet_visibility(ids[2]), SheetVisibility::VeryHidden);
    let d = ids[0];
    assert_val(book, d, 1, 1, s("Hello"));
    assert_val(book, d, 1, 2, Value::Double(42.0));
    assert_val(book, d, 1, 3, Value::Boolean(true));
    assert_val(book, d, 1, 4, Value::Error(2007));
    assert_val(book, d, 1, 5, s("inline"));
    assert_val(book, d, 2, 1, Value::Double(46309.0));
    assert_eq!(book.cell_format(d, 2, 1).number_format.as_deref(), Some("m/d/yyyy"));
    assert_eq!(book.cell_formula(d, 2, 2).as_deref(), Some("=B1*2"));
    assert_val(book, d, 2, 2, Value::Double(84.0));
    assert_eq!(book.cell_formula(d, 2, 3).as_deref(), Some("=\"x\"&A1"));
    assert_val(book, d, 2, 3, s("xHello"));
    assert_eq!(book.cell_formula(d, 3, 2).as_deref(), Some("=B1+1"));
    assert_eq!(book.cell_formula(d, 4, 2).as_deref(), Some("=B2+1"));
    assert_eq!(book.cell_formula(d, 5, 2).as_deref(), Some("=B3+1"));
    assert_eq!(book.cell_formula(d, 5, 3).as_deref(), Some("=SUM(B1:B2*1)"));
    assert_val(book, d, 3, 3, s("Rich text"));
    assert_val(book, d, 3, 4, s("a\rb & <c>"));
    let f = book.cell_format(d, 2, 4);
    assert_eq!(
        f,
        CellFormat {
            number_format: Some("#,##0.00".into()),
            bold: true,
            italic: true,
            underline: true,
            font_color: Some(255),
            font_size: Some(14.0),
            font_name: Some("Arial".into()),
            fill_color: Some(0x00ffff),
            borders: macroost_sheet::Borders {
                left: Some(BorderLine { style: 1, weight: 2, color: 0 }),
                right: Some(BorderLine { style: 1, weight: -4138, color: 0xff0000 }),
                top: Some(BorderLine { style: -4115, weight: 2, color: 0 }),
                bottom: Some(BorderLine { style: -4119, weight: 4, color: 0 }),
                ..Default::default()
            },
            h_align: HAlign::Center,
            wrap: true,
            ..Default::default()
        }
    );
    // theme color accent1 4472C4 -> RGB long
    let e3 = book.cell_format(d, 3, 5);
    assert_eq!(e3.font_color, Some(0x44 | (0x72 << 8) | (0xC4 << 16)));
    assert_eq!(e3.number_format.as_deref(), Some("0.000"));
    assert_eq!(book.cell_format(d, 2, 5).number_format.as_deref(), Some("m/d/yyyy h:mm"));
    assert!(!book.cell_format(d, 7, 1).is_default(), "style-only cell kept");
    assert!((book.column_width(d, 2) - 20.0).abs() < 0.01, "{}", book.column_width(d, 2));
    assert!(book.column_hidden(d, 4) && book.column_hidden(d, 5) && !book.column_hidden(d, 3));
    assert_eq!(book.row_height(d, 3), 30.0);
    assert!(book.row_hidden(d, 4));
    assert_eq!(book.merged_areas(d), vec![Rect::new(10, 1, 11, 3)]);
    assert!(book.is_protected(d));
    assert!(!book.is_protected(ids[1]));
    assert_eq!(book.code_name(d).as_deref(), Some("Sheet12"));
    assert_val(book, ids[1], 1, 1, Value::Double(7.0));
    assert_val(book, ids[1], 1, 2, s("Hello"));
    let mut names = book.defined_names();
    names.sort_by(|a, b| a.0.cmp(&b.0));
    assert_eq!(
        names,
        vec![
            ("Local".to_owned(), Some(ids[1]), "'Hidden & Co'!$A$1".to_owned()),
            ("Total".to_owned(), None, "Data!$B$1:$B$5".to_owned()),
            ("_xlnm.Print_Area".to_owned(), Some(d), "Data!$A$1:$E$5".to_owned()),
        ]
    );
}

#[test]
fn loads_the_sample() {
    let book = load(&sample()).unwrap();
    check_sample(&book);
    // Formula cells evaluate through the model; unknown functions fall back to the cached value.
    assert!(matches!(book.cell_value(1, 2, 2), Value::Double(x) if x == 84.0));
}

#[test]
fn round_trips_with_template() {
    let src = sample();
    let book = load(&src).unwrap();
    let out = saved(&book, Some(&src));
    let again = load(&out).unwrap();
    check_sample(&again);
    // Untouched parts are copied byte for byte; the macro-enabled content type stays.
    assert_eq!(part(&out, "xl/vbaProject.bin"), part(&src, "xl/vbaProject.bin"));
    assert_eq!(part(&out, "xl/theme/theme1.xml"), part(&src, "xl/theme/theme1.xml"));
    assert!(part(&out, "xl/drawings/drawing1.xml").is_some());
    assert!(part(&out, "xl/calcChain.xml").is_none(), "stale calc chain dropped");
    let ct = part_text(&out, "[Content_Types].xml");
    assert!(ct.contains("sheet.macroEnabled.main+xml"));
    assert!(!ct.contains("calcChain"));
    let rels = part_text(&out, "xl/_rels/workbook.xml.rels");
    assert!(rels.contains("vbaProject") && !rels.contains("calcChain"));
    // Features the model cannot hold survive in the sheet part.
    let s1 = part_text(&out, "xl/worksheets/sheet1.xml");
    for keep in [
        "<conditionalFormatting",
        "<dataValidations",
        "<drawing r:id=\"rId1\"/>",
        "x14:dataBar",
        "tabColor",
        "hashValue=\"abc\"",
        "<f t=\"array\" ref=\"C5:C6\">",
    ] {
        assert!(s1.contains(keep), "lost {keep}");
    }
    // The prefixed sheet keeps its prefix.
    assert!(part_text(&out, "xl/worksheets/sheet2.xml").contains("<x:sheetData>"));
    // Rich text shared string is kept as written.
    assert!(part_text(&out, "xl/sharedStrings.xml").contains("<rPr><b/></rPr>"));
    let st = part_text(&out, "xl/styles.xml");
    assert!(st.contains("<dxfs count=\"1\">"), "dxfs kept for conditional formats");
    let wb = part_text(&out, "xl/workbook.xml");
    assert!(
        wb.contains("codeName=\"ThisWorkbook\"") && wb.contains("hidden=\"1\"") && wb.contains("fullCalcOnLoad=\"1\"")
    );
    // Saving twice is stable.
    let out2 = saved(&again, Some(&out));
    check_sample(&load(&out2).unwrap());
}

#[test]
fn round_trips_without_template() {
    let book = load(&sample()).unwrap();
    let out = saved(&book, None);
    let again = load(&out).unwrap();
    check_sample(&again);
    let ct = part_text(&out, "[Content_Types].xml");
    assert!(ct.contains("sheet.main+xml") && ct.contains("worksheet+xml"));
}

#[test]
fn edits_survive_save() {
    let src = sample();
    let mut book = load(&src).unwrap();
    let ids = book.sheet_ids();
    book.delete_sheet(ids[1]).unwrap();
    let new = book.add_sheet("Added", 0).unwrap();
    book.rename_sheet(ids[0], "Renamed").unwrap();
    book.set_cell_value(new, 2, 3, Value::from("new text")).unwrap();
    book.set_cell_value(new, 1, 1, Value::Double(1.5)).unwrap();
    book.set_cell_formula(new, 1, 2, Some("=A1*2".into())).unwrap();
    let fmt = CellFormat {
        number_format: Some("0.0%".into()),
        italic: true,
        fill_color: Some(0x123456),
        ..CellFormat::default()
    };
    book.set_cell_format(new, 1, 1, fmt.clone()).unwrap();
    book.set_cell_value(ids[0], 1, 2, Value::Double(-0.125)).unwrap();
    book.set_cell_value(ids[0], 1, 1, Value::Empty).unwrap();
    book.set_protected(ids[0], false).unwrap();
    book.set_column_width(new, 3, 30.0).unwrap();
    book.set_row_hidden(new, 9, true).unwrap();
    book.merge_cells(new, Rect::new(5, 1, 6, 2)).unwrap();
    book.set_code_name(new, "Sheet99");
    book.set_defined_name("Fresh", Some(new), "Added!$A$1");
    book.set_sheet_visibility(ids[2], SheetVisibility::Visible).unwrap();

    let out = saved(&book, Some(&src));
    let b = load(&out).unwrap();
    let names: Vec<String> = b.sheet_ids().iter().map(|&i| b.sheet_name(i).unwrap()).collect();
    assert_eq!(names, ["Added", "Renamed", "VH"]);
    let [n, r, v] = [b.sheet_ids()[0], b.sheet_ids()[1], b.sheet_ids()[2]];
    assert_val(&b, n, 2, 3, s("new text"));
    assert_eq!(b.cell_formula(n, 1, 2).as_deref(), Some("=A1*2"));
    assert_val(&b, n, 1, 2, Value::Double(3.0));
    assert_eq!(b.cell_format(n, 1, 1), fmt);
    assert_eq!(b.column_width(n, 3), 30.0);
    assert!(b.row_hidden(n, 9));
    assert_eq!(b.merged_areas(n), vec![Rect::new(5, 1, 6, 2)]);
    assert_eq!(b.code_name(n).as_deref(), Some("Sheet99"));
    assert_val(&b, r, 1, 2, Value::Double(-0.125));
    assert_val(&b, r, 1, 1, Value::Empty);
    assert!(!b.is_protected(r));
    assert_eq!(b.sheet_visibility(v), SheetVisibility::Visible);
    let dn = b.defined_names();
    assert!(dn.contains(&("Fresh".to_owned(), Some(n), "Added!$A$1".to_owned())));
    assert!(!dn.iter().any(|d| d.0 == "Local"), "name scoped to the deleted sheet is gone");
    assert!(dn.contains(&("_xlnm.Print_Area".to_owned(), Some(r), "Data!$A$1:$E$5".to_owned())));
    // The deleted sheet's part and relationship are gone; the new one is registered.
    assert!(part(&out, "xl/worksheets/sheet2.xml").is_none());
    let ct = part_text(&out, "[Content_Types].xml");
    assert!(!ct.contains("/xl/worksheets/sheet2.xml"));
    let rels = part_text(&out, "xl/_rels/workbook.xml.rels");
    assert!(!rels.contains("worksheets/sheet2.xml"));
    let new_part = ["xl/worksheets/sheet4.xml"].into_iter().find(|p| part(&out, p).is_some()).expect("new sheet part");
    assert!(ct.contains(&format!("/{new_part}")));
    // Sheet ids stay unique.
    let wb = part_text(&out, "xl/workbook.xml");
    assert!(wb.contains("name=\"Renamed\" sheetId=\"1\"") && wb.contains("sheetId=\"5\""));
}

#[test]
fn date1904_values_read_in_the_1900_system() {
    let src = zip_of(&sample_parts(true));
    let book = load(&src).unwrap();
    assert_val(&book, 1, 2, 1, Value::Double(46309.0 + 1462.0));
    assert_val(&book, 1, 1, 2, Value::Double(42.0)); // not a date format
    let out = saved(&book, Some(&src));
    assert!(part_text(&out, "xl/worksheets/sheet1.xml").contains("<v>46309</v>"));
    assert_val(&load(&out).unwrap(), 1, 2, 1, Value::Double(46309.0 + 1462.0));
}

#[test]
fn hostile_input_is_an_error_not_a_panic() {
    assert!(load(b"").is_err());
    assert!(load(b"PK\x03\x04garbage").is_err());
    let src = sample();
    assert!(load(&src[..src.len() / 2]).is_err());
    // Broken XML in a sheet.
    let mut parts = sample_parts(false);
    for p in &mut parts {
        if p.0 == "xl/worksheets/sheet1.xml" {
            p.1 = b"<worksheet><sheetData><row><c r=\"A1\"><v>1</v></row>".to_vec();
        }
    }
    let e = load(&zip_of(&parts)).unwrap_err();
    assert!(e.to_string().contains("sheet1.xml"), "{e}");
    // Missing workbook part.
    let parts: Vec<_> = sample_parts(false).into_iter().filter(|p| p.0 != "xl/workbook.xml").collect();
    assert!(matches!(load(&zip_of(&parts)), Err(macroost_xlsx::Error::Missing(_))));
    // Garbage everywhere but the structure: no panic.
    let mut parts = sample_parts(false);
    for p in &mut parts {
        if p.0 == "xl/styles.xml" || p.0 == "xl/sharedStrings.xml" || p.0 == "xl/theme/theme1.xml" {
            p.1 = b"\xff\xfe<<<&#xD800;".to_vec();
        }
    }
    let _ = load(&zip_of(&parts));
    // Out-of-range references and indexes are skipped.
    let mut parts = sample_parts(false);
    for p in &mut parts {
        if p.0 == "xl/worksheets/sheet1.xml" {
            p.1 = br#"<worksheet><sheetData><row r="99999999"><c r="ZZZZ1"><v>1</v></c><c r="A1" t="s" s="999"><v>77</v></c><c r="B1"><f t="shared" si="9"/><v>5</v></c></row></sheetData><mergeCells><mergeCell ref="A1:XFE3"/></mergeCells><cols><col min="0" max="99999" width="-5"/></cols></worksheet>"#.to_vec();
        }
    }
    let book = load(&zip_of(&parts)).unwrap();
    assert_val(&book, 1, 1, 2, Value::Double(5.0));
    assert!(book.cell_formula(1, 1, 2).is_none());
}

#[test]
fn minimal_new_workbook_saves() {
    let mut book = MemoryWorkbook::with_sheets("New.xlsx", &["One", "Two"]);
    book.set_cell_value(1, 1, 1, Value::from(" padded ")).unwrap();
    book.set_cell_value(2, 3, 3, Value::Boolean(false)).unwrap();
    book.set_cell_value(2, 4, 4, Value::Error(2042)).unwrap();
    let out = saved(&book, None);
    let b = load(&out).unwrap();
    assert_eq!(b.sheet_ids().len(), 2);
    assert_val(&b, 1, 1, 1, s(" padded "));
    assert_val(&b, 2, 3, 3, Value::Boolean(false));
    assert_val(&b, 2, 4, 4, Value::Error(2042));
}

/// Verifiers in the form the reference application writes (a sheet protected with "Sample#42",
/// the workbook with "wb"): macros unprotect with the right password only, and a password set by
/// a macro is saved as a verifier.
#[test]
fn passwords_in_the_file_are_verified_and_saved() {
    use macroost_lang::Interpreter;
    use macroost_lang::ast::ModuleType;
    let sheet_hash = r#"algorithmName="SHA-512" hashValue="EjDY5Un+T56uthMPztQAq56G8zM16cWCHxS6ksrAZj15ZkrV5Xk72psQOTwXusp/wmQOIgY2t6uanxayyDQR2w==" saltValue="nwN3NOBvnvK7bIRJnKcguw==" spinCount="100000""#;
    let book_hash = r#"workbookAlgorithmName="SHA-512" workbookHashValue="uUV0LclqRyKkBGpjq/bfP29zCD9oxIaJZDN41fpsCoKiCs7kG44vOmPCk4sF5dVVoquDfcs4Fr+BULf7YSubUw==" workbookSaltValue="9m2qQ6Kr6UbkD9araQPbEg==" workbookSpinCount="100000""#;
    let mut parts = sample_parts(false);
    for (name, data) in parts.iter_mut() {
        let Ok(text) = String::from_utf8(data.clone()) else { continue };
        let text = match *name {
            "xl/worksheets/sheet1.xml" => text
                .replace(r#"algorithmName="SHA-512" hashValue="abc" saltValue="def" spinCount="100000""#, sheet_hash),
            "xl/workbook.xml" => text
                .replace("<bookViews>", &format!("<workbookProtection {book_hash} lockStructure=\"1\"/><bookViews>")),
            _ => continue,
        };
        *data = text.into_bytes();
    }
    let book = load(&zip_of(&parts)).unwrap();
    let data = book.sheet_ids()[0];
    assert!(book.sheet_password(data).is_some());
    assert!(book.workbook_protection().is_some_and(|p| p.structure && p.password.is_some()));
    let mut host = macroost_sheet::SheetHost::new(Box::new(book));
    let mut it = Interpreter::new(&mut host);
    let src = concat!(
        "Function T()\n On Error Resume Next\n",
        " Worksheets(\"Data\").Unprotect\n T = Err.Number\n Err.Clear\n",
        " Worksheets(\"Data\").Unprotect \"sample#42\"\n T = T & \"|\" & Err.Number\n Err.Clear\n",
        " Worksheets(\"Data\").Unprotect (\"Sample#42\")\n",
        " T = T & \"|\" & Err.Number & \"|\" & Worksheets(\"Data\").ProtectContents\n",
        " ActiveWorkbook.Unprotect Password:=\"nope\"\n",
        " T = T & \"|\" & Err.Number & \"|\" & ActiveWorkbook.ProtectStructure\n Err.Clear\n",
        " ActiveWorkbook.Unprotect Password:=\"wb\"\n",
        " T = T & \"|\" & Err.Number & \"|\" & ActiveWorkbook.ProtectStructure\n",
        " Worksheets(\"Data\").Protect Password:=\"again\"\nEnd Function\n"
    );
    it.add_module("M", ModuleType::Standard, src).unwrap();
    let got = it.run("T", vec![]).unwrap().to_text().unwrap();
    assert_eq!(got, "1004|1004|0|False|1004|True|0|False");
    drop(it);
    let book = host.into_workbook().ok().unwrap();
    let mem = book.as_any().and_then(|a| a.downcast_ref::<MemoryWorkbook>()).unwrap();
    let out = saved(mem, Some(&zip_of(&parts)));
    let sheet = part_text(&out, "xl/worksheets/sheet1.xml");
    let want = format!("password=\"{:04X}\"", macroost_sheet::legacy_password_hash("again"));
    assert!(sheet.contains(&want), "{sheet}");
    assert!(!part_text(&out, "xl/workbook.xml").contains("workbookProtection"));
}
