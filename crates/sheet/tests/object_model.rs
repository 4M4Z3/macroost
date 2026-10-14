//! Object model behavior driven through the `Object` / `Host` traits, as the interpreter will.
#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

mod common;
use common::*;
use macroost_lang::{Args, Host, Value};
use macroost_sheet::{MAX_ROWS, Rect};

fn three() -> macroost_sheet::SheetHost {
    host_with(&["Sheet1", "Sheet2", "Sheet3"])
}

fn err_no(r: macroost_lang::Result<Value>) -> i32 {
    match r {
        Ok(v) => panic!("expected an error, got {v:?}"),
        Err(e) => e.number,
    }
}

#[test]
fn constants_resolve() {
    let mut h = host();
    for (n, v) in [
        ("xlUp", -4162),
        ("xlDown", -4121),
        ("xlToLeft", -4159),
        ("xlToRight", -4161),
        ("xlPasteValues", -4163),
        ("xlPasteFormats", -4122),
        ("xlPasteAll", -4104),
        ("xlCalculationManual", -4135),
        ("xlCalculationAutomatic", -4105),
        ("xlNone", -4142),
        ("xlContinuous", 1),
        ("xlThin", 2),
        ("xlMedium", -4138),
        ("xlThick", 4),
        ("xlEdgeLeft", 7),
        ("xlEdgeTop", 8),
        ("xlEdgeBottom", 9),
        ("xlEdgeRight", 10),
        ("xlInsideVertical", 11),
        ("xlInsideHorizontal", 12),
        ("xlShiftUp", -4162),
        ("xlShiftToLeft", -4159),
        ("xlShiftDown", -4121),
        ("xlCenter", -4108),
        ("xlLeft", -4131),
        ("xlRight", -4152),
        ("xlWhole", 1),
        ("xlPart", 2),
        ("xlValues", -4163),
        ("xlFormulas", -4123),
        ("xlCellTypeConstants", 2),
        ("xlCellTypeBlanks", 4),
        ("xlCellTypeLastCell", 11),
        ("xlYes", 1),
        ("xlNo", 2),
        ("xlSheetVisible", -1),
        ("xlSheetHidden", 0),
        ("xlSheetVeryHidden", 2),
    ] {
        assert_eq!(num(&k(&mut h, n)), f64::from(v), "{n}");
    }
    assert!(h.global("NotAThing", &Args::none()).is_none());
}

#[test]
fn values_round_trip_through_default_member() {
    let mut h = host();
    h.global_assign("Range", &args(vec!["A1".into()]), 5.into()).unwrap().unwrap();
    let a1 = rng(&mut h, "A1");
    assert_eq!(num(&a1.clone().deref_default().unwrap()), 5.0);
    assert_eq!(val(&mut h, "A1").type_name(), "Double");
    // ActiveCell = x
    h.global_assign("ActiveCell", &Args::none(), "hi".into()).unwrap().unwrap();
    assert_eq!(text(&val(&mut h, "A1")), "hi");
    // Typed text: numbers, percent, booleans, apostrophe, formulas.
    put(&mut h, "B1", "1,234".into());
    assert_eq!(num(&val(&mut h, "B1")), 1234.0);
    put(&mut h, "B2", "12%".into());
    assert_eq!(num(&val(&mut h, "B2")), 0.12);
    assert_eq!(text(&call(&rng(&mut h, "B2"), "Text", vec![])), "12%");
    put(&mut h, "B3", "true".into());
    assert!(matches!(val(&mut h, "B3"), Value::Boolean(true)));
    put(&mut h, "B4", "'007".into());
    assert_eq!(text(&val(&mut h, "B4")), "007");
    put(&mut h, "B5", "=B1*2".into());
    assert_eq!(num(&val(&mut h, "B5")), 2468.0);
    assert_eq!(text(&call(&rng(&mut h, "B5"), "Formula", vec![])), "=B1*2");
    // Text format keeps text.
    set(&rng(&mut h, "B6"), "NumberFormat", "@".into());
    put(&mut h, "B6", "0012".into());
    assert_eq!(text(&val(&mut h, "B6")), "0012");
    // Currency stays currency.
    put(&mut h, "B7", Value::Currency(12_500));
    assert_eq!(val(&mut h, "B7").type_name(), "Currency");
    assert_eq!(text(&call(&rng(&mut h, "B7"), "Value2", vec![])), "1.25");
}

#[test]
fn two_d_arrays_and_fill() {
    let mut h = host();
    put(&mut h, "A1:C3", 1.into());
    let v = val(&mut h, "A1:C3");
    let Value::Array(a) = &v else { panic!() };
    assert_eq!(a.borrow().bounds, vec![(1, 3), (1, 3)]);
    assert_eq!(num(&a.borrow().get(&[3, 3]).unwrap()), 1.0);
    // Write it back shifted.
    put(&mut h, "E5:G7", v.clone());
    assert_eq!(num(&val(&mut h, "G7")), 1.0);
    // A 1-D array fills across; a larger target gets #N/A.
    let row = Value::Array(std::rc::Rc::new(std::cell::RefCell::new(macroost_lang::Array::from_vec(
        0,
        vec![1.into(), 2.into()],
    ))));
    put(&mut h, "A10:C11", row);
    assert_eq!(num(&val(&mut h, "B11")), 2.0);
    assert!(matches!(val(&mut h, "C10"), Value::Error(2042)));
    assert_eq!(text(&call(&rng(&mut h, "C10"), "Text", vec![])), "#N/A");
    // Value2 of a date is a Double.
    put(&mut h, "D1", Value::Date(46309.0));
    assert_eq!(call(&rng(&mut h, "D1"), "Value2", vec![]).type_name(), "Double");
    assert_eq!(val(&mut h, "D1").type_name(), "Date");
}

#[test]
fn addresses_and_parsing() {
    let mut h = three();
    assert_eq!(addr(&rng(&mut h, "A1:C3")), "$A$1:$C$3");
    assert_eq!(addr(&rng(&mut h, "c3:a1")), "$A$1:$C$3");
    assert_eq!(addr(&rng(&mut h, "$B$2")), "$B$2");
    assert_eq!(addr(&rng(&mut h, "A:A")), "$A:$A");
    assert_eq!(addr(&rng(&mut h, "2:3")), "$2:$3");
    assert_eq!(addr(&rng(&mut h, "A1,B2:C3")), "$A$1,$B$2:$C$3");
    assert_eq!(addr(&g(&mut h, "Cells", vec![])), "$1:$1048576");
    let r = rng(&mut h, "B2:C3");
    let r1c1 = call(&r, "Address", vec![true.into(), true.into(), (-4150).into()]);
    assert_eq!(text(&r1c1), "R2C2:R3C3");
    let ext = calln(&r, "Address", named(vec![], vec![("External", true.into())]));
    assert_eq!(text(&ext), "[Book1.xlsm]Sheet1!$B$2:$C$3");
    // Sheet-qualified global Range.
    let q = rng(&mut h, "Sheet2!B2");
    assert_eq!(text(&call(&call(&q, "Worksheet", vec![]), "Name", vec![])), "Sheet2");
    let q = rng(&mut h, "'Sheet3'!A1:B2");
    assert_eq!(text(&call(&call(&q, "Parent", vec![]), "Name", vec![])), "Sheet3");
    // Range(cell1, cell2) with objects and strings.
    let c1 = cells(&mut h, 2, 2);
    let c2 = cells(&mut h, 4, 5);
    assert_eq!(addr(&g(&mut h, "Range", vec![c1, c2])), "$B$2:$E$4");
    assert_eq!(addr(&g(&mut h, "Range", vec!["A1".into(), "C2".into()])), "$A$1:$C$2");
    // Invalid addresses raise 1004.
    assert_eq!(err_no(h.global("Range", &args(vec!["A0".into()])).unwrap()), 1004);
    assert_eq!(err_no(h.global("Range", &args(vec!["NoSheet!A1".into()])).unwrap()), 1004);
    assert_eq!(err_no(h.global("Cells", &args(vec![0.into(), 1.into()])).unwrap()), 1004);
    // Cells(n) on the sheet is linear across 16384 columns.
    assert_eq!(addr(&g(&mut h, "Cells", vec![16385.into()])), "$A$2");
}

#[test]
fn sheet_qualified_access() {
    let mut h = three();
    let ws2 = g(&mut h, "Worksheets", vec!["Sheet2".into()]);
    set(&call(&ws2, "Range", vec!["B2".into()]), "Value", 99.into());
    assert!(val(&mut h, "B2").is_empty(), "Sheet1 untouched");
    assert_eq!(num(&val(&mut h, "Sheet2!B2")), 99.0);
    assert_eq!(num(&call(&call(&ws2, "Cells", vec![2.into(), 2.into()]), "Value", vec![])), 99.0);
    // A range from another sheet's object is rejected.
    let r = try_call(&ws2, "Range", args(vec!["Sheet1!A1".into()]));
    assert_eq!(err_no(r), 1004);
    // Missing sheet: 9.
    assert_eq!(err_no(h.global("Worksheets", &args(vec!["Nope".into()])).unwrap()), 9);
    assert_eq!(err_no(h.global("Worksheets", &args(vec![4.into()])).unwrap()), 9);
    // Unknown member: 438.
    assert_eq!(err_no(try_call(&ws2, "Frobnicate", Args::none())), 438);
}

#[test]
fn worksheets_collection() {
    let mut h = three();
    let ws = k(&mut h, "Worksheets");
    assert_eq!(num(&call(&ws, "Count", vec![])), 3.0);
    let Value::Object(Some(o)) = &ws else { panic!() };
    let names: Vec<String> = o.items().unwrap().iter().map(|s| text(&call(s, "Name", vec![]))).collect();
    assert_eq!(names, ["Sheet1", "Sheet2", "Sheet3"]);
    // Add after the last sheet; it becomes active and is named Sheet4.
    let last = g(&mut h, "Worksheets", vec![3.into()]);
    let new = calln(&ws, "Add", named(vec![], vec![("After", last)]));
    assert_eq!(text(&call(&new, "Name", vec![])), "Sheet4");
    assert_eq!(num(&call(&new, "Index", vec![])), 4.0);
    assert_eq!(text(&call(&k(&mut h, "ActiveSheet"), "Name", vec![])), "Sheet4");
    // Default Add goes before the active sheet.
    let n5 = call(&ws, "Add", vec![]);
    assert_eq!(num(&call(&n5, "Index", vec![])), 4.0);
    // Rename, with validation.
    set(&n5, "Name", "Data".into());
    assert_eq!(text(&call(&g(&mut h, "Sheets", vec!["data".into()]), "Name", vec![])), "Data");
    assert_eq!(try_set(&n5, "Name", "Sheet1".into()).unwrap_err().number, 1004);
    assert_eq!(try_set(&n5, "Name", "a/b".into()).unwrap_err().number, 1004);
    // Delete the active sheet: activation moves on.
    call(&n5, "Delete", vec![]);
    assert_eq!(num(&call(&ws, "Count", vec![])), 4.0);
    assert_eq!(err_no(try_call(&n5, "Name", Args::none())), 424);
    // Visibility.
    let s2 = g(&mut h, "Worksheets", vec!["Sheet2".into()]);
    set(&s2, "Visible", Value::Long(2));
    assert_eq!(num(&call(&s2, "Visible", vec![])), 2.0);
    assert_eq!(err_no(try_call(&s2, "Activate", Args::none())), 1004);
    set(&s2, "Visible", true.into());
    assert_eq!(num(&call(&s2, "Visible", vec![])), -1.0);
    // Move Sheet1 after Sheet3.
    let s1 = g(&mut h, "Worksheets", vec!["Sheet1".into()]);
    let s3 = g(&mut h, "Worksheets", vec!["Sheet3".into()]);
    calln(&s1, "Move", named(vec![], vec![("After", s3)]));
    assert_eq!(num(&call(&s1, "Index", vec![])), 3.0);
    // Copy a sheet.
    set(&call(&s1, "Range", vec!["A1".into()]), "Value", "copied".into());
    let s2 = g(&mut h, "Worksheets", vec!["Sheet2".into()]);
    calln(&s1, "Copy", named(vec![], vec![("Before", s2)]));
    assert_eq!(text(&val(&mut h, "'Sheet1 (2)'!A1")), "copied");
    // Next / Previous / Parent.
    let first = g(&mut h, "Worksheets", vec![1.into()]);
    assert!(is_nothing(&call(&first, "Previous", vec![])));
    assert_eq!(text(&call(&call(&first, "Parent", vec![]), "Name", vec![])), "Book1.xlsm");
}

#[test]
fn cannot_delete_or_hide_the_last_visible_sheet() {
    let mut h = host();
    let s = k(&mut h, "ActiveSheet");
    assert_eq!(err_no(try_call(&s, "Delete", Args::none())), 1004);
    assert_eq!(try_set(&s, "Visible", false.into()).unwrap_err().number, 1004);
}

#[test]
fn offset_resize_rows_columns() {
    let mut h = host();
    let r = rng(&mut h, "B2:C3");
    assert_eq!(addr(&call(&r, "Offset", vec![1.into(), 1.into()])), "$C$3:$D$4");
    assert_eq!(addr(&call(&r, "Resize", vec![1.into()])), "$B$2:$C$2");
    assert_eq!(addr(&calln(&r, "Resize", named(vec![], vec![("ColumnSize", 4.into())]))), "$B$2:$E$3");
    assert_eq!(err_no(try_call(&r, "Offset", args(vec![(-2).into(), 0.into()]))), 1004);
    assert_eq!(err_no(try_call(&r, "Resize", args(vec![0.into()]))), 1004);
    assert_eq!(addr(&call(&r, "Rows", vec![2.into()])), "$B$3:$C$3");
    assert_eq!(addr(&call(&r, "Columns", vec![2.into()])), "$C$2:$C$3");
    assert_eq!(addr(&call(&r, "Range", vec!["B2".into()])), "$C$3");
    assert_eq!(addr(&call(&r, "Cells", vec![3.into(), 1.into()])), "$B$4");
    assert_eq!(addr(&call(&r, "Item", vec![2.into()])), "$C$2");
    assert_eq!(addr(&call(&r, "", vec![3.into()])), "$B$3");
    assert_eq!(addr(&call(&g(&mut h, "Columns", vec!["B:D".into()]), "Columns", vec![2.into()])), "$C:$C");
    assert_eq!(num(&call(&g(&mut h, "Columns", vec![]), "Count", vec![])), 16384.0);
    assert_eq!(num(&call(&g(&mut h, "Rows", vec!["2:4".into()]), "Count", vec![])), 3.0);
    assert_eq!(num(&call(&rng(&mut h, "A1:C3"), "EntireRow", vec![]).clone().pipe_count()), 3.0);
    assert_eq!(err_no(try_call(&g(&mut h, "Cells", vec![]), "Count", Args::none())), 6);
    assert_eq!(text(&call(&g(&mut h, "Cells", vec![]), "CountLarge", vec![])), "17179869184");
    // Rows of a range iterate as rows.
    let Value::Object(Some(o)) = call(&r, "Rows", vec![]) else { panic!() };
    let items = o.items().unwrap();
    assert_eq!(items.len(), 2);
    assert_eq!(addr(&items[1]), "$B$3:$C$3");
}

trait PipeCount {
    fn pipe_count(self) -> Value;
}
impl PipeCount for Value {
    fn pipe_count(self) -> Value {
        call(&self, "Count", vec![])
    }
}

#[test]
fn end_on_rows() {
    let mut h = host();
    for c in ["B3", "C3", "D3", "G3"] {
        put(&mut h, c, 1.into());
    }
    let right = k(&mut h, "xlToRight");
    let left = k(&mut h, "xlToLeft");
    assert_eq!(addr(&call(&rng(&mut h, "B3"), "End", vec![right.clone()])), "$D$3");
    assert_eq!(addr(&call(&rng(&mut h, "D3"), "End", vec![right.clone()])), "$G$3");
    assert_eq!(addr(&call(&rng(&mut h, "G3"), "End", vec![right])), "$XFD$3");
    assert_eq!(addr(&call(&rng(&mut h, "XFD3"), "End", vec![left.clone()])), "$G$3");
    assert_eq!(addr(&call(&rng(&mut h, "D3"), "End", vec![left.clone()])), "$B$3");
    assert_eq!(addr(&call(&rng(&mut h, "B3"), "End", vec![left])), "$A$3");
    let up = k(&mut h, "xlUp");
    assert_eq!(addr(&call(&rng(&mut h, "A1"), "End", vec![up])), "$A$1");
}

#[test]
fn whole_column_ranges_stay_cheap() {
    let mut h = host();
    put(&mut h, "A1", 1.into());
    put(&mut h, "A500000", 2.into());
    let col = rng(&mut h, "A:A");
    let up = k(&mut h, "xlUp");
    assert_eq!(addr(&call(&cells(&mut h, MAX_ROWS as i32, 1), "End", vec![up])), "$A$500000");
    let wf = k(&mut h, "WorksheetFunction");
    assert_eq!(num(&call(&wf, "Sum", vec![col.clone()])), 3.0);
    assert_eq!(num(&call(&wf, "CountA", vec![col.clone()])), 2.0);
    assert_eq!(num(&call(&wf, "CountBlank", vec![col.clone()])), f64::from(MAX_ROWS) - 2.0);
    assert_eq!(num(&call(&wf, "Match", vec![2.into(), col.clone(), 0.into()])), 500000.0);
    // Formatting a whole column only touches used cells.
    set(&call(&col, "Font", vec![]), "Bold", true.into());
    assert!(matches!(call(&call(&rng(&mut h, "A500000"), "Font", vec![]), "Bold", vec![]), Value::Boolean(true)));
    // Clearing a whole column visits stored cells only.
    call(&col, "ClearContents", vec![]);
    assert!(val(&mut h, "A500000").is_empty());
    // Whole-sheet Value would need 17 billion elements: out of memory (7), not a hang.
    assert_eq!(err_no(try_call(&g(&mut h, "Cells", vec![]), "Value", Args::none())), 7);
    // For Each over a whole column visits the used part only.
    let mut h = host();
    put(&mut h, "A3", 1.into());
    let col = rng(&mut h, "A:A");
    let Value::Object(Some(o)) = &col else { panic!() };
    assert_eq!(o.items().unwrap().len(), 3);
}

#[test]
fn multi_area_union_intersect() {
    let mut h = host();
    let (x, y) = (rng(&mut h, "A1:B2"), rng(&mut h, "D4"));
    let u = g(&mut h, "Union", vec![x, y]);
    assert_eq!(addr(&u), "$A$1:$B$2,$D$4");
    assert_eq!(num(&call(&u, "Count", vec![])), 5.0);
    let areas = call(&u, "Areas", vec![]);
    assert_eq!(num(&call(&areas, "Count", vec![])), 2.0);
    assert_eq!(addr(&call(&areas, "Item", vec![2.into()])), "$D$4");
    let (x, y) = (rng(&mut h, "A1:C3"), rng(&mut h, "B2:D4"));
    let i = g(&mut h, "Intersect", vec![x, y]);
    assert_eq!(addr(&i), "$B$2:$C$3");
    let (x, y) = (rng(&mut h, "A1"), rng(&mut h, "C3"));
    let none = g(&mut h, "Intersect", vec![x, y]);
    assert!(is_nothing(&none));
    // Value assignment goes to every area.
    put(&mut h, "A1,C3", 8.into());
    assert_eq!(num(&val(&mut h, "C3")), 8.0);
    assert!(val(&mut h, "B2").is_empty());
    // Different sheets cannot be combined.
    let mut h3 = three();
    let a = rng(&mut h3, "Sheet2!A1");
    let b = rng(&mut h3, "A1");
    assert_eq!(err_no(h3.global("Union", &args(vec![a, b])).unwrap()), 1004);
}

#[test]
fn copy_paste_variants() {
    let mut h = host();
    put(&mut h, "A1", 1.into());
    put(&mut h, "A2", 2.into());
    set(&rng(&mut h, "A3"), "Formula", "=A1+A2".into());
    set(&call(&rng(&mut h, "A1"), "Font", vec![]), "Bold", true.into());
    // Copy Destination shifts relative references.
    calln(&rng(&mut h, "A1:A3"), "Copy", named(vec![], vec![("Destination", rng(&mut h, "C1"))]));
    assert_eq!(text(&call(&rng(&mut h, "C3"), "Formula", vec![])), "=C1+C2");
    assert_eq!(num(&val(&mut h, "C3")), 3.0);
    assert!(matches!(call(&call(&rng(&mut h, "C1"), "Font", vec![]), "Bold", vec![]), Value::Boolean(true)));
    // Clipboard + PasteSpecial formats only.
    call(&rng(&mut h, "A1"), "Copy", vec![]);
    assert_eq!(num(&call(&k(&mut h, "Application"), "CutCopyMode", vec![])), 1.0);
    let pf = k(&mut h, "xlPasteFormats");
    call(&rng(&mut h, "E1"), "PasteSpecial", vec![pf]);
    assert!(val(&mut h, "E1").is_empty());
    assert!(matches!(call(&call(&rng(&mut h, "E1"), "Font", vec![]), "Bold", vec![]), Value::Boolean(true)));
    // Paste values into a block tiles the source.
    let pv = k(&mut h, "xlPasteValues");
    call(&rng(&mut h, "G1:H2"), "PasteSpecial", vec![pv.clone()]);
    assert_eq!(num(&val(&mut h, "H2")), 1.0);
    // Transpose.
    call(&rng(&mut h, "A1:A2"), "Copy", vec![]);
    calln(&rng(&mut h, "J1"), "PasteSpecial", named(vec![pv], vec![("Transpose", true.into())]));
    assert_eq!(num(&val(&mut h, "K1")), 2.0);
    // Cut + Paste moves.
    call(&rng(&mut h, "A2"), "Cut", vec![]);
    assert_eq!(num(&call(&k(&mut h, "Application"), "CutCopyMode", vec![])), 2.0);
    let ws = k(&mut h, "ActiveSheet");
    calln(&ws, "Paste", named(vec![], vec![("Destination", rng(&mut h, "M5"))]));
    assert_eq!(num(&val(&mut h, "M5")), 2.0);
    assert!(val(&mut h, "A2").is_empty());
    // PasteSpecial with nothing copied fails.
    set(&k(&mut h, "Application"), "CutCopyMode", false.into());
    assert_eq!(err_no(try_call(&rng(&mut h, "A9"), "PasteSpecial", Args::none())), 1004);
}

#[test]
fn insert_and_delete_cells() {
    let mut h = host();
    for (a, v) in [("A1", 1), ("B1", 2), ("C1", 3), ("A2", 4)] {
        put(&mut h, a, v.into());
    }
    let right = k(&mut h, "xlShiftToRight");
    calln(&rng(&mut h, "A1"), "Insert", named(vec![], vec![("Shift", right)]));
    assert!(val(&mut h, "A1").is_empty());
    assert_eq!(num(&val(&mut h, "D1")), 3.0);
    assert_eq!(num(&val(&mut h, "A2")), 4.0, "other rows untouched");
    let left = k(&mut h, "xlShiftToLeft");
    calln(&rng(&mut h, "A1"), "Delete", named(vec![], vec![("Shift", left)]));
    assert_eq!(num(&val(&mut h, "A1")), 1.0);
    // Insert with the default shift on a single cell goes down.
    call(&rng(&mut h, "A1"), "Insert", vec![]);
    assert_eq!(num(&val(&mut h, "A2")), 1.0);
    assert_eq!(num(&val(&mut h, "A3")), 4.0);
    // Entire column insert/delete.
    call(&call(&rng(&mut h, "A1"), "EntireColumn", vec![]), "Insert", vec![]);
    assert_eq!(num(&val(&mut h, "B2")), 1.0);
    call(&g(&mut h, "Columns", vec![1.into()]), "Delete", vec![]);
    assert_eq!(num(&val(&mut h, "A2")), 1.0);
    // Insert copied rows.
    call(&g(&mut h, "Rows", vec![2.into()]), "Copy", vec![]);
    call(&g(&mut h, "Rows", vec![1.into()]), "Insert", vec![]);
    assert_eq!(num(&val(&mut h, "A1")), 1.0);
    assert_eq!(num(&val(&mut h, "A3")), 1.0);
    assert!(matches!(call(&k(&mut h, "Application"), "CutCopyMode", vec![]), Value::Long(0)));
    // Inserted rows inherit the format from above.
    set(&call(&rng(&mut h, "A5"), "Interior", vec![]), "ColorIndex", 6.into());
    call(&g(&mut h, "Rows", vec![6.into()]), "Insert", vec![]);
    assert_eq!(num(&call(&call(&rng(&mut h, "A6"), "Interior", vec![]), "ColorIndex", vec![])), 6.0);
}

#[test]
fn formats() {
    let mut h = host();
    let r = rng(&mut h, "A1:B2");
    let interior = call(&r, "Interior", vec![]);
    assert_eq!(num(&call(&interior, "ColorIndex", vec![])), -4142.0);
    assert_eq!(num(&call(&interior, "Color", vec![])), 16_777_215.0);
    set(&interior, "ColorIndex", 3.into());
    assert_eq!(num(&call(&interior, "Color", vec![])), 255.0);
    assert_eq!(num(&call(&interior, "Pattern", vec![])), 1.0);
    set(&call(&rng(&mut h, "A1"), "Interior", vec![]), "Color", Value::Long(65280));
    // Mixed fills report the first cell's color, not Null (CONFORMANCE sheet3 L003).
    assert_eq!(num(&call(&interior, "Color", vec![])), 65280.0);
    assert!(call(&interior, "ColorIndex", vec![]).is_null(), "mixed color indexes read as Null");
    set(&interior, "ColorIndex", k(&mut h, "xlNone"));
    assert_eq!(num(&call(&interior, "Pattern", vec![])), -4142.0);
    let font = call(&r, "Font", vec![]);
    assert_eq!(num(&call(&font, "Size", vec![])), 11.0);
    set(&font, "Size", 14.into());
    set(&font, "Name", "Arial".into());
    set(&font, "Italic", true.into());
    set(&font, "Color", Value::Long(255));
    assert_eq!(num(&call(&font, "Size", vec![])), 14.0);
    assert_eq!(text(&call(&font, "Name", vec![])), "Arial");
    assert!(matches!(call(&font, "Italic", vec![]), Value::Boolean(true)));
    assert_eq!(num(&call(&font, "ColorIndex", vec![])), 3.0);
    set(&r, "HorizontalAlignment", k(&mut h, "xlCenter"));
    assert_eq!(num(&call(&r, "HorizontalAlignment", vec![])), -4108.0);
    set(&r, "WrapText", true.into());
    assert!(matches!(call(&r, "WrapText", vec![]), Value::Boolean(true)));
    set(&r, "NumberFormat", "0.00".into());
    put(&mut h, "A1", Value::Double(4.56789));
    assert_eq!(text(&call(&rng(&mut h, "A1"), "Text", vec![])), "4.57");
    assert_eq!(text(&call(&r, "NumberFormat", vec![])), "0.00");
    // ClearFormats resets, Clear removes content too.
    call(&r, "ClearFormats", vec![]);
    assert_eq!(text(&call(&r, "NumberFormat", vec![])), "General");
    assert_eq!(num(&val(&mut h, "A1")), 4.56789);
    call(&r, "Clear", vec![]);
    assert!(val(&mut h, "A1").is_empty());
}

#[test]
fn borders() {
    let mut h = host();
    let r = rng(&mut h, "B2:C3");
    let top = call(&r, "Borders", vec![k(&mut h, "xlEdgeTop")]);
    assert_eq!(num(&call(&top, "LineStyle", vec![])), -4142.0);
    set(&top, "LineStyle", k(&mut h, "xlContinuous"));
    set(&top, "Weight", k(&mut h, "xlMedium"));
    assert_eq!(num(&call(&top, "LineStyle", vec![])), 1.0);
    assert_eq!(num(&call(&top, "Weight", vec![])), -4138.0);
    // The shared edge reads from the cell above too.
    let above = call(&rng(&mut h, "B1"), "Borders", vec![k(&mut h, "xlEdgeBottom")]);
    assert_eq!(num(&call(&above, "LineStyle", vec![])), 1.0);
    // Borders collection: every edge.
    let all = call(&rng(&mut h, "E5:F6"), "Borders", vec![]);
    set(&all, "LineStyle", 1.into());
    let inside = call(&rng(&mut h, "E5:F6"), "Borders", vec![k(&mut h, "xlInsideVertical")]);
    assert_eq!(num(&call(&inside, "LineStyle", vec![])), 1.0);
    set(&all, "LineStyle", k(&mut h, "xlNone"));
    assert_eq!(num(&call(&all, "LineStyle", vec![])), -4142.0);
    // BorderAround.
    call(&rng(&mut h, "H1:I2"), "BorderAround", vec![1.into(), 4.into()]);
    let right = call(&rng(&mut h, "H1:I2"), "Borders", vec![k(&mut h, "xlEdgeRight")]);
    assert_eq!(num(&call(&right, "Weight", vec![])), 4.0);
}

#[test]
fn hidden_heights_widths() {
    let mut h = host();
    let rows = call(&rng(&mut h, "A2:A3"), "EntireRow", vec![]);
    set(&rows, "Hidden", true.into());
    assert!(matches!(call(&g(&mut h, "Rows", vec![3.into()]), "Hidden", vec![]), Value::Boolean(true)));
    // Partly hidden rows read False, not Null (CONFORMANCE sheet3 L023).
    assert!(matches!(call(&g(&mut h, "Rows", vec!["1:2".into()]), "Hidden", vec![]), Value::Boolean(false)));
    assert_eq!(try_set(&rng(&mut h, "A1"), "Hidden", true.into()).unwrap_err().number, 1004);
    let col = g(&mut h, "Columns", vec!["C".into()]);
    set(&col, "ColumnWidth", Value::Double(20.0));
    assert_eq!(num(&call(&rng(&mut h, "C5"), "ColumnWidth", vec![])), 20.0);
    set(&col, "Hidden", true.into());
    assert!(matches!(call(&call(&rng(&mut h, "C1"), "EntireColumn", vec![]), "Hidden", vec![]), Value::Boolean(true)));
    set(&rng(&mut h, "A1"), "RowHeight", Value::Double(30.0));
    assert_eq!(num(&call(&rng(&mut h, "B1"), "RowHeight", vec![])), 30.0);
    assert_eq!(num(&call(&rng(&mut h, "A4"), "RowHeight", vec![])), 15.0);
    // SpecialCells(xlCellTypeVisible) skips hidden rows.
    let vis = call(&rng(&mut h, "A1:A4"), "SpecialCells", vec![12.into()]);
    assert_eq!(addr(&vis), "$A$1,$A$4");
}

#[test]
fn find_and_find_next() {
    let mut h = host();
    put(&mut h, "A2", "apple".into());
    put(&mut h, "B5", "Apple pie".into());
    put(&mut h, "A9", "APPLE".into());
    let r = rng(&mut h, "A1:C10");
    let first = calln(&r, "Find", named(vec!["apple".into()], vec![]));
    assert_eq!(addr(&first), "$A$2");
    let second = call(&r, "FindNext", vec![first.clone()]);
    assert_eq!(addr(&second), "$B$5");
    let third = call(&r, "FindNext", vec![second]);
    assert_eq!(addr(&third), "$A$9");
    let wrap = call(&r, "FindNext", vec![third]);
    assert_eq!(addr(&wrap), "$A$2");
    // MatchCase and whole match.
    let mc = calln(&r, "Find", named(vec!["APPLE".into()], vec![("MatchCase", true.into()), ("LookAt", 1.into())]));
    assert_eq!(addr(&mc), "$A$9");
    // Wildcards, column order.
    let wc = calln(&r, "Find", named(vec!["*pie".into()], vec![("LookAt", 1.into())]));
    assert_eq!(addr(&wc), "$B$5");
    let by_col =
        calln(&r, "Find", named(vec!["apple".into()], vec![("SearchOrder", 2.into()), ("After", rng(&mut h, "A2"))]));
    assert_eq!(addr(&by_col), "$A$9");
    // Previous direction from the top wraps to the last match.
    let prev = calln(&r, "Find", named(vec!["apple".into()], vec![("SearchDirection", 2.into())]));
    assert_eq!(addr(&prev), "$A$9");
    // Formulas are searched as text under xlFormulas.
    set(&rng(&mut h, "C1"), "Formula", "=SUM(A1:A2)".into());
    // LookAt persists between Find calls (CONFORMANCE sheet4 D004), so ask for a part match again.
    let f =
        calln(&r, "Find", named(vec!["SUM(".into()], vec![("LookIn", k(&mut h, "xlFormulas")), ("LookAt", 2.into())]));
    assert_eq!(addr(&f), "$C$1");
}

#[test]
fn special_cells() {
    let mut h = host();
    put(&mut h, "A1", 1.into());
    put(&mut h, "B1", "x".into());
    put(&mut h, "A3", 3.into());
    set(&rng(&mut h, "C3"), "Formula", "=A1".into());
    let r = rng(&mut h, "A1:C3");
    assert_eq!(addr(&call(&r, "SpecialCells", vec![k(&mut h, "xlCellTypeConstants")])), "$A$1:$B$1,$A$3");
    assert_eq!(addr(&call(&r, "SpecialCells", vec![k(&mut h, "xlCellTypeConstants"), 1.into()])), "$A$1,$A$3");
    assert_eq!(addr(&call(&r, "SpecialCells", vec![(-4123).into()])), "$C$3");
    assert_eq!(addr(&call(&r, "SpecialCells", vec![k(&mut h, "xlCellTypeBlanks")])), "$C$1:$C$2,$B$2:$B$3,$A$2");
    assert_eq!(addr(&call(&rng(&mut h, "A1"), "SpecialCells", vec![k(&mut h, "xlCellTypeLastCell")])), "$C$3");
    assert_eq!(err_no(try_call(&rng(&mut h, "E1:E2"), "SpecialCells", args(vec![2.into()]))), 1004);
}

#[test]
fn formulas_r1c1_and_fill() {
    let mut h = host();
    put(&mut h, "A1:A3", 2.into());
    set(&rng(&mut h, "B1:B3"), "Formula", "=A1*10".into());
    assert_eq!(text(&call(&rng(&mut h, "B3"), "Formula", vec![])), "=A3*10");
    assert_eq!(num(&val(&mut h, "B3")), 20.0);
    assert_eq!(text(&call(&rng(&mut h, "B3"), "FormulaR1C1", vec![])), "=RC[-1]*10");
    set(&rng(&mut h, "C1:C3"), "FormulaR1C1", "=SUM(RC[-2]:RC[-1])".into());
    assert_eq!(text(&call(&rng(&mut h, "C2"), "Formula", vec![])), "=SUM(A2:B2)");
    assert_eq!(num(&val(&mut h, "C2")), 22.0);
    let f = call(&rng(&mut h, "B1:C1"), "Formula", vec![]);
    let Value::Array(a) = f else { panic!() };
    assert_eq!(text(&a.borrow().get(&[1, 2]).unwrap()), "=SUM(A1:B1)");
    assert!(call(&rng(&mut h, "A1:B1"), "HasFormula", vec![]).is_null());
    // Formula on a constant cell is its text.
    assert_eq!(text(&call(&rng(&mut h, "A1"), "Formula", vec![])), "2");
    assert!(matches!(call(&rng(&mut h, "A1"), "HasFormula", vec![]), Value::Boolean(false)));
}

#[test]
fn worksheet_functions() {
    let mut h = host();
    for (i, (name, qty)) in [("apple", 3), ("pear", 5), ("apple", 7), ("fig", 1)].iter().enumerate() {
        put(&mut h, &format!("A{}", i + 1), (*name).into());
        put(&mut h, &format!("B{}", i + 1), Value::Long(*qty));
    }
    let wf = k(&mut h, "WorksheetFunction");
    let a = rng(&mut h, "A1:A4");
    let b = rng(&mut h, "B1:B4");
    let t = rng(&mut h, "A1:B4");
    assert_eq!(num(&call(&wf, "CountIf", vec![a.clone(), "apple".into()])), 2.0);
    assert_eq!(num(&call(&wf, "CountIf", vec![b.clone(), ">2".into()])), 3.0);
    assert_eq!(num(&call(&wf, "CountIf", vec![a.clone(), "<>apple".into()])), 2.0);
    assert_eq!(num(&call(&wf, "CountIf", vec![a.clone(), "?ear".into()])), 1.0);
    assert_eq!(num(&call(&wf, "SumIf", vec![a.clone(), "apple".into(), b.clone()])), 10.0);
    assert_eq!(num(&call(&wf, "SumIf", vec![b.clone(), ">=5".into()])), 12.0);
    assert_eq!(num(&call(&wf, "SumIfs", vec![b.clone(), a.clone(), "apple".into(), b.clone(), ">3".into()])), 7.0);
    assert_eq!(num(&call(&wf, "CountIfs", vec![a.clone(), "apple".into(), b.clone(), "<5".into()])), 1.0);
    assert_eq!(num(&call(&wf, "VLookup", vec!["pear".into(), t.clone(), 2.into(), false.into()])), 5.0);
    assert_eq!(num(&call(&wf, "VLookup", vec!["FIG".into(), t.clone(), 2.into(), false.into()])), 1.0);
    assert_eq!(err_no(try_call(&wf, "VLookup", args(vec!["kiwi".into(), t.clone(), 2.into(), false.into()]))), 1004);
    assert_eq!(err_no(try_call(&wf, "VLookup", args(vec!["fig".into(), t.clone(), 3.into(), false.into()]))), 1004);
    assert_eq!(num(&call(&wf, "Match", vec!["fig".into(), a.clone(), 0.into()])), 4.0);
    assert_eq!(num(&call(&wf, "Sum", vec![b.clone(), 10.into(), true.into()])), 27.0);
    assert_eq!(num(&call(&wf, "Count", vec![t.clone()])), 4.0);
    assert_eq!(num(&call(&wf, "CountA", vec![t.clone()])), 8.0);
    assert_eq!(num(&call(&wf, "Min", vec![b.clone()])), 1.0);
    assert_eq!(num(&call(&wf, "Max", vec![b.clone(), 9.into()])), 9.0);
    assert_eq!(num(&call(&wf, "Average", vec![b.clone()])), 4.0);
    assert_eq!(err_no(try_call(&wf, "Average", args(vec![a.clone()]))), 1004);
    assert_eq!(num(&call(&wf, "Round", vec![Value::Double(2.675), 2.into()])), 2.68);
    assert_eq!(num(&call(&wf, "Round", vec![Value::Double(-0.5), 0.into()])), -1.0);
    assert_eq!(num(&call(&wf, "Index", vec![t.clone(), 3.into(), 2.into()])), 7.0);
    assert_eq!(text(&call(&wf, "Trim", vec!["  a   b ".into()])), "a b");
    // Approximate match on sorted numbers.
    for (i, v) in [10, 20, 30].iter().enumerate() {
        put(&mut h, &format!("D{}", i + 1), Value::Long(*v));
    }
    assert_eq!(num(&call(&wf, "Match", vec![25.into(), rng(&mut h, "D1:D3")])), 2.0);
    // Application.Match returns an error value instead of raising.
    let app = k(&mut h, "Application");
    assert!(matches!(call(&app, "Match", vec!["kiwi".into(), a, 0.into()]), Value::Error(2042)));
    assert_eq!(num(&call(&app, "Sum", vec![b])), 16.0);
}

#[test]
fn select_activate_and_sheets() {
    let mut h = three();
    let s2 = g(&mut h, "Worksheets", vec!["Sheet2".into()]);
    // Selecting on an inactive sheet fails like the application.
    let other = call(&s2, "Range", vec!["A1".into()]);
    assert_eq!(err_no(try_call(&other, "Select", Args::none())), 1004);
    call(&s2, "Activate", vec![]);
    call(&other, "Select", vec![]);
    assert_eq!(text(&call(&k(&mut h, "ActiveSheet"), "Name", vec![])), "Sheet2");
    call(&rng(&mut h, "B2:D4"), "Select", vec![]);
    call(&rng(&mut h, "C3"), "Activate", vec![]);
    assert_eq!(addr(&k(&mut h, "Selection")), "$B$2:$D$4", "activate inside keeps the selection");
    assert_eq!(addr(&k(&mut h, "ActiveCell")), "$C$3");
    call(&rng(&mut h, "F9"), "Activate", vec![]);
    assert_eq!(addr(&k(&mut h, "Selection")), "$F$9");
    // Unqualified Range follows the active sheet.
    put(&mut h, "A1", "on2".into());
    assert_eq!(text(&val(&mut h, "Sheet2!A1")), "on2");
    // Application.Goto switches sheets.
    let target = rng(&mut h, "Sheet3!C2");
    call(&k(&mut h, "Application"), "Goto", vec![target]);
    assert_eq!(text(&call(&k(&mut h, "ActiveSheet"), "Name", vec![])), "Sheet3");
    assert_eq!(addr(&k(&mut h, "ActiveCell")), "$C$2");
}

#[test]
fn protection_blocks_writes() {
    let mut h = host();
    let ws = k(&mut h, "ActiveSheet");
    call(&ws, "Protect", vec!["pw".into()]);
    assert_eq!(try_set(&rng(&mut h, "A1"), "Value", 1.into()).unwrap_err().number, 1004);
    assert_eq!(err_no(try_call(&ws, "Unprotect", args(vec!["wrong".into()]))), 1004);
    call(&ws, "Unprotect", vec!["pw".into()]);
    put(&mut h, "A1", 1.into());
    // UserInterfaceOnly lets macros write.
    calln(&ws, "Protect", named(vec![], vec![("UserInterfaceOnly", true.into())]));
    put(&mut h, "A2", 2.into());
    assert_eq!(num(&val(&mut h, "A2")), 2.0);
    assert!(matches!(call(&ws, "ProtectContents", vec![]), Value::Boolean(true)));
}

#[test]
fn application_state_and_log() {
    let mut h = host();
    let app = k(&mut h, "Application");
    for p in ["ScreenUpdating", "DisplayAlerts", "EnableEvents"] {
        assert!(matches!(call(&app, p, vec![]), Value::Boolean(true)));
        set(&app, p, false.into());
        assert!(matches!(call(&app, p, vec![]), Value::Boolean(false)));
    }
    set(&app, "Calculation", k(&mut h, "xlCalculationManual"));
    assert_eq!(num(&call(&app, "Calculation", vec![])), -4135.0);
    // An unknown calculation mode is taken as manual without an error (CONFORMANCE sheet5 Y003).
    set(&app, "Calculation", 99.into());
    assert_eq!(num(&call(&app, "Calculation", vec![])), -4135.0);
    set(&app, "StatusBar", "Working...".into());
    assert_eq!(text(&call(&app, "StatusBar", vec![])), "Working...");
    // Handing the status bar back reads as the text "FALSE" (CONFORMANCE sheet5 Y004).
    set(&app, "StatusBar", false.into());
    assert_eq!(text(&call(&app, "StatusBar", vec![])), "FALSE");
    set(&app, "EnableCancelKey", 0.into());
    assert_eq!(num(&call(&app, "EnableCancelKey", vec![])), 0.0);
    assert_eq!(text(&call(&app, "Version", vec![])), "16.0");
    assert_eq!(err_no(try_call(&app, "NoSuchThing", Args::none())), 438);
    let wb = k(&mut h, "ThisWorkbook");
    assert_eq!(text(&call(&wb, "Name", vec![])), "Book1.xlsm");
    call(&wb, "Save", vec![]);
    call(&rng(&mut h, "A1:B2"), "Merge", vec![]);
    assert!(matches!(call(&rng(&mut h, "A1:B2"), "MergeCells", vec![]), Value::Boolean(true)));
    assert!(call(&rng(&mut h, "A1:C1"), "MergeCells", vec![]).is_null());
    call(&rng(&mut h, "A1:B2"), "UnMerge", vec![]);
    call(&rng(&mut h, "A1:D9"), "AutoFilter", vec![]);
    assert!(matches!(call(&k(&mut h, "ActiveSheet"), "AutoFilterMode", vec![]), Value::Boolean(true)));
    assert_eq!(h.log(), ["Workbook.Save", "Merge Sheet1!$A$1:$B$2", "UnMerge Sheet1!$A$1:$B$2"]);
    let books = k(&mut h, "Workbooks");
    assert_eq!(num(&call(&books, "Count", vec![])), 1.0);
    assert_eq!(text(&call(&call(&books, "Item", vec!["Book1".into()]), "Name", vec![])), "Book1.xlsm");
    assert_eq!(err_no(try_call(&books, "Item", args(vec![2.into()]))), 9);
}

#[test]
fn messages_and_debug_print_are_recorded() {
    let mut h = host();
    h.msgbox_reply = 6;
    assert_eq!(h.msgbox("Continue?", 4, "Title"), 6);
    h.debug_print("line 1");
    h.input_replies.push(Some("typed".into()));
    assert_eq!(h.inputbox("Name?", "", "def").as_deref(), Some("typed"));
    assert_eq!(h.inputbox("Again?", "", "def").as_deref(), Some("def"));
    assert_eq!(h.messages, ["Continue?", "Name?", "Again?"]);
    assert_eq!(h.printed, ["line 1"]);
    h.fixed_now = Some(46309.5);
    assert_eq!(h.now(), 46309.5);
}

#[test]
fn sort_rows() {
    let mut h = host();
    put(&mut h, "A1", "Name".into());
    put(&mut h, "B1", "Qty".into());
    for (i, (n, q)) in [("pear", 2), ("apple", 9), ("fig", 5)].iter().enumerate() {
        put(&mut h, &format!("A{}", i + 2), (*n).into());
        put(&mut h, &format!("B{}", i + 2), Value::Long(*q));
    }
    let r = rng(&mut h, "A1:B4");
    calln(&r, "Sort", named(vec![], vec![("Key1", rng(&mut h, "B1")), ("Order1", 2.into()), ("Header", 1.into())]));
    assert_eq!(text(&val(&mut h, "A1")), "Name");
    assert_eq!(text(&val(&mut h, "A2")), "apple");
    assert_eq!(text(&val(&mut h, "A4")), "pear");
    calln(&r, "Sort", named(vec![], vec![("Key1", rng(&mut h, "A1")), ("Header", 1.into())]));
    assert_eq!(text(&val(&mut h, "A3")), "fig");
    assert_eq!(num(&val(&mut h, "B3")), 5.0);
}

#[test]
fn workbook_trait_is_reachable() {
    let mut h = host();
    put(&mut h, "C3", 1.into());
    let used = h.workbook().unwrap().used_range(1);
    assert_eq!(used, Some(Rect::cell(3, 3)));
    h.with_workbook_mut(|wb| wb.set_cell_value(1, 1, 1, Value::from("ext"))).unwrap().unwrap();
    assert_eq!(text(&val(&mut h, "A1")), "ext");
    assert!(h.into_workbook().is_ok());
}

#[test]
fn defined_names_resolve_in_range_and_names() {
    let mut h = three();
    h.with_workbook_mut(|wb| {
        wb.set_defined_name("Total", None, "Sheet2!$B$2:$C$3");
        wb.set_defined_name("Local", Some(1), "$A$5");
        wb.set_defined_name("Local", None, "Sheet3!$D$1");
        wb.set_defined_name("Rate", None, "0.05");
    })
    .unwrap();
    // Global Range("Total") goes to Sheet2.
    let r = rng(&mut h, "Total");
    assert_eq!(text(&call(&r, "Address", vec![])), "$B$2:$C$3");
    assert_eq!(text(&call(&call(&r, "Worksheet", vec![]), "Name", vec![])), "Sheet2");
    // A sheet-scoped name wins over the workbook one from its own sheet.
    let r = rng(&mut h, "Local");
    assert_eq!(text(&call(&r, "Address", vec![])), "$A$5");
    assert_eq!(text(&call(&rng(&mut h, "Sheet2!Total"), "Address", vec![])), "$B$2:$C$3");
    // ws.Range("Total") on its own sheet, and an error on another sheet.
    let ws2 = g(&mut h, "Worksheets", vec!["Sheet2".into()]);
    assert_eq!(text(&call(&call(&ws2, "Range", vec!["Total".into()]), "Address", vec![])), "$B$2:$C$3");
    let ws3 = g(&mut h, "Worksheets", vec!["Sheet3".into()]);
    assert_eq!(err_no(try_call(&ws3, "Range", args(vec!["Total".into()]))), 1004);
    // Unknown names and non-range names raise 1004.
    assert_eq!(err_no(h.global("Range", &args(vec!["Nope".into()])).unwrap()), 1004);
    assert_eq!(err_no(h.global("Range", &args(vec!["Rate".into()])).unwrap()), 1004);
    // Names collection.
    let wb = k(&mut h, "ThisWorkbook");
    let names = call(&wb, "Names", vec![]);
    assert_eq!(num(&call(&names, "Count", vec![])), 4.0);
    let n = call(&wb, "Names", vec!["Total".into()]);
    assert_eq!(text(&call(&n, "RefersTo", vec![])), "=Sheet2!$B$2:$C$3");
    let rr = call(&n, "RefersToRange", vec![]);
    assert_eq!(text(&call(&rr, "Address", vec![])), "$B$2:$C$3");
    calln(&names, "Add", named(vec![], vec![("Name", "Out".into()), ("RefersTo", "=Sheet1!$E$1".into())]));
    put(&mut h, "Out", 7.into());
    assert_eq!(num(&val(&mut h, "E1")), 7.0);
    // Adding from a Range object, then deleting.
    let src = rng(&mut h, "Sheet3!A1:A2");
    calln(&names, "Add", named(vec!["Col".into(), src], vec![]));
    let c = call(&names, "Item", vec!["Col".into()]);
    assert_eq!(text(&call(&c, "RefersTo", vec![])), "=Sheet3!$A$1:$A$2");
    call(&c, "Delete", vec![]);
    assert_eq!(err_no(h.global("Range", &args(vec!["Col".into()])).unwrap()), 1004);
    assert_eq!(err_no(try_call(&names, "Add", args(vec!["A1".into(), "=Sheet1!$A$1".into()]))), 1004);
}

#[test]
fn item_assignment_and_text_of_dates() {
    let mut h = host();
    let r = rng(&mut h, "A1:B2");
    let Value::Object(Some(o)) = &r else { panic!() };
    o.assign("", &args(vec![2.into()]), 5.into()).unwrap();
    assert_eq!(num(&val(&mut h, "B1")), 5.0);
    put(&mut h, "C1", "2026-10-14".into());
    assert_eq!(text(&call(&rng(&mut h, "C1"), "Text", vec![])), "10/14/2026");
    set(&rng(&mut h, "C1"), "NumberFormat", "dddd d mmm yyyy".into());
    assert_eq!(text(&call(&rng(&mut h, "C1"), "Text", vec![])), "Wednesday 14 Oct 2026");
    put(&mut h, "C2", "3:30 PM".into());
    assert_eq!(text(&call(&rng(&mut h, "C2"), "Text", vec![])), "3:30 PM");
    assert_eq!(num(&call(&rng(&mut h, "C2"), "Value2", vec![])), 15.5 / 24.0);
}
