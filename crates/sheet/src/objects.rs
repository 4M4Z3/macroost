//! `Application`, `Window`, `Workbooks`, `Workbook`, `Worksheets`/`Sheets` and `Worksheet`.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::address::Rect;
use crate::formula::RefEdit;
use crate::functions::{WorksheetFunctionObj, call_function};
use crate::range::{Kind, RangeObj, decode_range, range_arg, range_value, resolve_address};
use crate::state::{
    ALLOW_NAMES, PasteKind, Protection, Shared, State, data, err1004, opt_bool, opt_i64, opt_text, with, with_mut,
};
use crate::workbook::{SheetId, SheetVisibility, WorkbookProtection};

/// Hidden member through which objects recognize a Worksheet passed as an argument.
pub(crate) const SHEET_REF: &str = "_macroostSheetRef";

const ONE_VISIBLE: &str = "A workbook must contain at least one visible worksheet. \n\nTo hide, delete, or move the selected sheet(s), you must first insert a new sheet or unhide a sheet that is already hidden.";
const NAME_TAKEN: &str = "That name is already taken. Try a different one.";
const BAD_NAME: &str = "You typed an invalid name for a sheet or chart.";
const STRUCTURE: &str = "Workbook is protected and cannot be changed.";
const BAD_PASSWORD: &str = "The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.";

pub(crate) fn app_value(app: &Shared) -> Value {
    Value::Object(Some(Rc::new(AppObj { app: app.clone() })))
}

pub(crate) fn sheet_value(app: &Shared, id: SheetId) -> Value {
    Value::Object(Some(Rc::new(SheetObj { app: app.clone(), id })))
}

pub(crate) fn workbook_value(app: &Shared) -> Value {
    Value::Object(Some(Rc::new(WorkbookObj { app: app.clone() })))
}

pub(crate) fn sheets_value(app: &Shared, _name: &'static str) -> Value {
    Value::Object(Some(Rc::new(SheetsObj { app: app.clone(), subset: None })))
}

pub(crate) fn window_value(app: &Shared) -> Value {
    Value::Object(Some(Rc::new(WindowObj { app: app.clone() })))
}

pub(crate) fn decode_sheet(v: &Value) -> Option<SheetId> {
    let Value::Object(Some(o)) = v else { return None };
    match o.call(SHEET_REF, &Args::none()) {
        Ok(Value::Long(n)) => u32::try_from(n).ok(),
        _ => None,
    }
}

/// The whole-sheet range, through which sheet-level `Range`, `Cells`, `Rows`, `Columns` work.
fn whole(app: &Shared, id: SheetId) -> RangeObj {
    RangeObj { app: app.clone(), sheet: id, areas: vec![Rect::sheet()], kind: Kind::Cells }
}

pub(crate) fn active_sheet(app: &Shared) -> Result<SheetId> {
    with(
        app,
        |st| if st.sheet_exists(st.active_sheet) { Ok(st.active_sheet) } else { Err(RuntimeError::object_required()) },
    )
}

/// `Range(...)` without an object: a qualified address names its own sheet.
pub(crate) fn global_range(app: &Shared, args: &Args) -> Result<Value> {
    let sheet = active_sheet(app)?;
    if let Some(Value::String(s)) = args.get(0, "Cell1")
        && args.get(1, "Cell2").is_none()
    {
        let (target, areas) = with(app, |st| resolve_address(st, sheet, s))?;
        return Ok(range_value(app, target, areas, Kind::Cells));
    }
    if let Some(first) = args.get(0, "Cell1")
        && let Some((target, _)) = decode_range(first)
    {
        return whole(app, target).call("Range", args);
    }
    whole(app, sheet).call("Range", args)
}

/// Sheet-level member (`Range`, `Cells`, `Rows`, `Columns`) on sheet `id`.
pub(crate) fn sheet_member(app: &Shared, id: SheetId, name: &str, args: &Args) -> Result<Value> {
    whole(app, id).call(name, args)
}

pub(crate) fn active_cell(app: &Shared) -> Result<Value> {
    let sheet = active_sheet(app)?;
    let (r, c) = with(app, |st| Ok(st.view(sheet).active))?;
    Ok(range_value(app, sheet, vec![Rect::cell(r, c)], Kind::Cells))
}

pub(crate) fn selection(app: &Shared) -> Result<Value> {
    let sheet = active_sheet(app)?;
    let sel = with(app, |st| Ok(st.view(sheet).selection))?;
    Ok(range_value(app, sheet, sel, Kind::Cells))
}

fn range_args(app: &Shared, args: &Args) -> Result<Vec<(SheetId, Vec<Rect>)>> {
    let sheet = active_sheet(app)?;
    let mut out = Vec::new();
    for v in &args.positional {
        if matches!(v, Value::Error(448)) {
            continue;
        }
        out.push(range_arg(app, sheet, v)?);
    }
    for (_, v) in &args.named {
        out.push(range_arg(app, sheet, v)?);
    }
    if out.len() < 2 {
        return Err(RuntimeError::standard(449));
    }
    Ok(out)
}

/// `Evaluate("A1:B2")` / `[A1]` / `Evaluate("SUM(A1:A3)")` on the active sheet.
pub(crate) fn evaluate(app: &Shared, args: &Args) -> Result<Value> {
    let text = data(args.require(0, "Name")?)?.to_text()?;
    crate::functions::evaluate_text(app, active_sheet(app)?, &text)
}

/// `Application.Intersect`: `Nothing` when the ranges do not overlap.
pub(crate) fn intersect(app: &Shared, args: &Args) -> Result<Value> {
    let ranges = range_args(app, args)?;
    let sheet = ranges[0].0;
    if ranges.iter().any(|r| r.0 != sheet) {
        return Err(err1004("Method 'Intersect' of object '_Application' failed"));
    }
    let mut acc = ranges[0].1.clone();
    for (_, areas) in &ranges[1..] {
        let mut next = Vec::new();
        for a in &acc {
            for b in areas {
                if let Some(x) = a.intersect(b)
                    && !next.contains(&x)
                {
                    next.push(x);
                }
            }
        }
        acc = next;
    }
    if acc.is_empty() {
        return Ok(Value::Object(None));
    }
    Ok(range_value(app, sheet, acc, Kind::Cells))
}

/// Joins areas that together form one rectangle (`A1` and `A2` become `A1:A2`).
pub(crate) fn merge_adjacent(mut areas: Vec<Rect>) -> Vec<Rect> {
    loop {
        let mut merged = false;
        'outer: for i in 0..areas.len() {
            for j in 0..areas.len() {
                if i == j {
                    continue;
                }
                let (a, b) = (areas[i], areas[j]);
                let joined = if a.contains_rect(&b) {
                    Some(a)
                } else if a.c1 == b.c1 && a.c2 == b.c2 && b.r1 == a.r2 + 1 {
                    Some(Rect::new(a.r1, a.c1, b.r2, a.c2))
                } else if a.r1 == b.r1 && a.r2 == b.r2 && b.c1 == a.c2 + 1 {
                    Some(Rect::new(a.r1, a.c1, a.r2, b.c2))
                } else {
                    None
                };
                if let Some(x) = joined {
                    areas[i] = x;
                    areas.remove(j);
                    merged = true;
                    break 'outer;
                }
            }
        }
        if !merged {
            return areas;
        }
    }
}

/// `Application.Union`: the areas of every argument (same sheet).
pub(crate) fn union(app: &Shared, args: &Args) -> Result<Value> {
    let ranges = range_args(app, args)?;
    let sheet = ranges[0].0;
    if ranges.iter().any(|r| r.0 != sheet) {
        return Err(err1004("Method 'Union' of object '_Application' failed"));
    }
    let mut acc: Vec<Rect> = Vec::new();
    for (_, areas) in ranges {
        acc.extend(areas);
    }
    Ok(range_value(app, sheet, merge_adjacent(acc), Kind::Cells))
}

/// Properties the application accepts and remembers without further effect, with defaults.
fn setting_default(name: &str) -> Option<Value> {
    Some(match name {
        "interactive"
        | "displaystatusbar"
        | "asktoupdatelinks"
        | "printcommunication"
        | "calculatebeforesave"
        | "displayformulabar"
        | "displayscrollbars"
        | "enableanimations"
        | "enableautocomplete"
        | "celldraganddrop"
        | "movesafterreturn"
        | "displaycommentindicator"
        | "showwindowsintaskbar"
        | "displayrecentfiles"
        | "enablelivepreview"
        | "showtooltips"
        | "displaypastoptions"
        | "displayinsertoptions"
        | "autorecover_enabled" => Value::Boolean(true),
        "displaynoteindicator" | "editdirectlyincell" | "extendlist" | "enablesound" | "enablemacroanimations" => {
            Value::Boolean(true)
        }
        "movesafterreturndirection" => Value::Long(-4121),
        "caption" => Value::from("macroost"),
        "iteration" | "displayfullscreen" | "visible_" => Value::Boolean(false),
        "enablecancelkey" => Value::Long(1),
        "cursor" => Value::Long(-4143),
        "maxiterations" => Value::Long(100),
        "maxchange" => Value::Double(0.001),
        "windowstate" => Value::Long(-4137),
        "username" => Value::from("User"),
        "decimalseparator" => Value::from("."),
        "thousandsseparator" => Value::from(","),
        "usesystemseparators" => Value::Boolean(true),
        "sheetsinnewworkbook" => Value::Long(1),
        "standardfont" => Value::from("Calibri"),
        "standardfontsize" => Value::Double(11.0),
        "referencestyle" => Value::Long(1),
        "visible" => Value::Boolean(true),
        _ => return None,
    })
}

/// `Application.International(index)` for an en-US system.
fn international(index: i64) -> Value {
    match index {
        1 => Value::Long(1),
        3 => Value::from("."),
        4 => Value::from(","),
        5 => Value::from(","),
        17 => Value::from("/"),
        18 => Value::from(":"),
        32 => Value::Long(0),
        _ => Value::Boolean(false),
    }
}

struct AppObj {
    app: Shared,
}

impl Object for AppObj {
    fn type_name(&self) -> String {
        "Application".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        let app = &self.app;
        match lname.as_str() {
            "application" => Ok(app_value(app)),
            "parent" => Ok(app_value(app)),
            "name" | "value" => Ok(Value::from("macroost")),
            "version" => Ok(Value::from("16.0")),
            "operatingsystem" => Ok(Value::from("macroost")),
            "screenupdating" => with(app, |st| Ok(Value::Boolean(st.screen_updating))),
            "displayalerts" => with(app, |st| Ok(Value::Boolean(st.display_alerts))),
            "enableevents" => with(app, |st| Ok(Value::Boolean(st.enable_events))),
            "calculation" => with(app, |st| Ok(Value::Long(st.calculation))),
            "statusbar" => with(app, |st| Ok(st.status_bar.clone())),
            "cutcopymode" => with(app, |st| {
                Ok(match &st.clipboard {
                    None => Value::Long(0),
                    Some(c) => Value::Long(if c.cut { 2 } else { 1 }),
                })
            }),
            "international" => Ok(international(opt_i64(args, 0, "Index")?.unwrap_or(0))),
            "activesheet" => Ok(sheet_value(app, active_sheet(app)?)),
            "activeworkbook" | "thisworkbook" => Ok(workbook_value(app)),
            "activewindow" => Ok(window_value(app)),
            "names" => crate::names::names_member(app, None, args),
            "activecell" => active_cell(app),
            "selection" => selection(app),
            "workbooks" => {
                let w = Value::Object(Some(Rc::new(WorkbooksObj { app: app.clone() }) as Rc<dyn Object>));
                if args.is_empty() { Ok(w) } else { WorkbooksObj { app: app.clone() }.call("Item", args) }
            }
            "worksheets" | "sheets" => {
                if args.is_empty() {
                    Ok(sheets_value(app, "Sheets"))
                } else {
                    SheetsObj { app: app.clone(), subset: None }.call("Item", args)
                }
            }
            "range" => global_range(app, args),
            "cells" | "rows" | "columns" => sheet_member(app, active_sheet(app)?, name, args),
            "intersect" => intersect(app, args),
            "union" => union(app, args),
            "evaluate" => evaluate(app, args),
            "worksheetfunction" => Ok(Value::Object(Some(Rc::new(WorksheetFunctionObj { app: app.clone() })))),
            "calculate" | "calculatefull" => with_mut(app, |st| {
                st.book.calculate();
                Ok(Value::Boolean(true))
            }),
            "goto" => {
                let target = args.require(0, "Reference")?;
                let (sheet, areas) = range_arg(app, active_sheet(app)?, target)?;
                with_mut(app, |st| {
                    st.active_sheet = sheet;
                    let v = st.views.entry(sheet).or_default();
                    v.selection = areas.clone();
                    v.active = (areas[0].r1, areas[0].c1);
                    Ok(Value::Boolean(true))
                })
            }
            "wait" => Ok(Value::Boolean(true)),
            "ontime" | "run" | "volatile" | "quit" | "sendkeys" => with_mut(app, |st| {
                st.log.push(format!("Application.{name}"));
                Ok(Value::Boolean(true))
            }),
            "convertformula" => {
                let f = data(args.require(0, "Formula")?)?.to_text()?;
                let from = opt_i64(args, 1, "FromReferenceStyle")?.unwrap_or(1);
                let to = opt_i64(args, 2, "ToReferenceStyle")?.unwrap_or(from);
                let cell = active_cell(app)?;
                let (r, c) = decode_range(&cell).map(|(_, a)| (a[0].r1, a[0].c1)).unwrap_or((1, 1));
                Ok(Value::from(match (from, to) {
                    (-4150, 1) => crate::formula::r1c1_to_a1(&f, r, c),
                    (1, -4150) => crate::formula::a1_to_r1c1(&f, r, c),
                    _ => f,
                }))
            }
            _ => {
                if let Some(v) = with(app, |st| Ok(st.settings.get(&lname).cloned()))? {
                    return Ok(v);
                }
                if let Some(v) = setting_default(&lname) {
                    return Ok(v);
                }
                // Application.Max(...) etc.: worksheet functions returning error values.
                match call_function(app, name, args) {
                    Some(r) => r,
                    None => Err(RuntimeError::no_member(name)),
                }
            }
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let lname = name.to_ascii_lowercase();
        let v = data(&value)?;
        with_mut(&self.app, |st| {
            match lname.as_str() {
                "screenupdating" => st.screen_updating = v.to_bool()?,
                "displayalerts" => st.display_alerts = v.to_bool()?,
                "enableevents" => st.enable_events = v.to_bool()?,
                "calculation" => {
                    let n = match v.to_i32()? {
                        n @ (-4105 | 2) => n,
                        _ => -4135,
                    };
                    st.book.set_manual_calculation(n == -4135);
                    if n == -4105 && st.calculation != -4105 {
                        st.book.calculate();
                    }
                    st.calculation = n;
                }
                "statusbar" => {
                    st.status_bar = match v {
                        Value::Boolean(false) => Value::from("FALSE"),
                        other => Value::from(other.to_text()?),
                    }
                }
                "cutcopymode" => {
                    let keep = match &v {
                        Value::Boolean(b) => *b,
                        other => other.to_i64()? != 0,
                    };
                    if !keep {
                        st.clipboard = None;
                    }
                }
                other if setting_default(other).is_some() => {
                    st.settings.insert(other.to_owned(), v);
                }
                _ => return Err(RuntimeError::no_member(name)),
            }
            Ok(())
        })
    }
}

/// `ActiveWindow`: scroll position, frozen panes, zoom and gridlines of the active sheet,
/// remembered without a display.
struct WindowObj {
    app: Shared,
}

impl Object for WindowObj {
    fn type_name(&self) -> String {
        "Window".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let app = &self.app;
        let sheet = active_sheet(app)?;
        let w = with(app, |st| Ok(st.windows.get(&sheet).cloned().unwrap_or_default()))?;
        match name.to_ascii_lowercase().as_str() {
            "scrollrow" => Ok(Value::Long(w.scroll_row)),
            "scrollcolumn" => Ok(Value::Long(w.scroll_col)),
            "freezepanes" => Ok(Value::Boolean(w.freeze.is_some())),
            "splitrow" => Ok(Value::Long(w.freeze.map_or(0, |f| f.0 as i32))),
            "splitcolumn" => Ok(Value::Long(w.freeze.map_or(0, |f| f.1 as i32))),
            "split" => Ok(Value::Boolean(false)),
            "zoom" => Ok(w.zoom.clone()),
            "displaygridlines" => Ok(Value::Boolean(w.gridlines)),
            "displayheadings" => Ok(Value::Boolean(w.headings)),
            "activesheet" => Ok(sheet_value(app, sheet)),
            "activecell" => active_cell(app),
            "selection" => selection(app),
            "visible" => Ok(Value::Boolean(true)),
            "windowstate" => Ok(Value::Long(-4137)),
            "caption" => with(app, |st| Ok(Value::from(st.book.name()))),
            "application" => Ok(app_value(app)),
            "parent" => Ok(workbook_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        let app = &self.app;
        let sheet = active_sheet(app)?;
        let v = data(&value)?;
        let lname = name.to_ascii_lowercase();
        with_mut(app, |st| {
            let active = st.view(sheet).active;
            let w = st.windows.entry(sheet).or_default();
            match lname.as_str() {
                "scrollrow" => w.scroll_row = v.to_i32()?.max(1),
                "scrollcolumn" => w.scroll_col = v.to_i32()?.max(1),
                "freezepanes" => {
                    w.freeze = if v.to_bool()? { Some((active.0 - 1, active.1 - 1)) } else { None };
                }
                "splitrow" | "splitcolumn" | "split" | "windowstate" | "visible" | "caption" => {}
                "zoom" => w.zoom = v,
                "displaygridlines" => w.gridlines = v.to_bool()?,
                "displayheadings" => w.headings = v.to_bool()?,
                _ => return Err(RuntimeError::no_member(name)),
            }
            Ok(())
        })
    }
}

struct WorkbooksObj {
    app: Shared,
}

impl Object for WorkbooksObj {
    fn type_name(&self) -> String {
        "Workbooks".into()
    }
    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(1)),
            "" | "item" => {
                let idx = data(args.require(0, "Index")?)?;
                let ok = match &idx {
                    Value::String(s) => with(&self.app, |st| {
                        let n = st.book.name().to_lowercase();
                        let s = s.to_lowercase();
                        Ok(n == s || n.rsplit_once('.').is_some_and(|(stem, _)| stem == s))
                    })?,
                    other => other.to_i64()? == 1,
                };
                if ok { Ok(workbook_value(&self.app)) } else { Err(RuntimeError::subscript()) }
            }
            "add" | "open" => Err(err1004("Opening or creating other workbooks is not available to macros here")),
            "application" | "parent" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
    fn items(&self) -> Result<Vec<Value>> {
        Ok(vec![workbook_value(&self.app)])
    }
}

struct WorkbookObj {
    app: Shared,
}

impl WorkbookObj {
    fn protect(&self, args: &Args) -> Result<Value> {
        let password = opt_text(args, 0, "Password")?.unwrap_or_default();
        let structure = opt_bool(args, 1, "Structure")?.unwrap_or(true);
        let windows = opt_bool(args, 2, "Windows")?.unwrap_or(false);
        with_mut(&self.app, |st| {
            if st.book.workbook_protection().is_some() {
                return Ok(Value::Boolean(true));
            }
            st.book.set_workbook_protection(Some(WorkbookProtection {
                structure,
                windows,
                password: crate::protect::new_password_hash(&password),
            }));
            Ok(Value::Boolean(true))
        })
    }

    fn unprotect(&self, args: &Args) -> Result<Value> {
        let password = opt_text(args, 0, "Password")?;
        with_mut(&self.app, |st| {
            let Some(p) = st.book.workbook_protection() else { return Ok(Value::Boolean(true)) };
            if let Some(h) = &p.password
                && !crate::protect::verify_password(h, password.as_deref().unwrap_or(""))
            {
                return Err(err1004(BAD_PASSWORD));
            }
            st.book.set_workbook_protection(None);
            Ok(Value::Boolean(true))
        })
    }
}

impl Object for WorkbookObj {
    fn type_name(&self) -> String {
        "Workbook".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        let app = &self.app;
        match lname.as_str() {
            "name" | "fullname" => with(app, |st| Ok(Value::from(st.book.name()))),
            "path" => Ok(Value::from("")),
            "codename" => Ok(Value::from("ThisWorkbook")),
            "names" => crate::names::names_member(app, None, args),
            "worksheets" | "sheets" => {
                if args.is_empty() {
                    Ok(sheets_value(app, "Sheets"))
                } else {
                    SheetsObj { app: app.clone(), subset: None }.call("Item", args)
                }
            }
            "activesheet" => Ok(sheet_value(app, active_sheet(app)?)),
            "activate" => Ok(Value::Boolean(true)),
            "saved" => with(app, |st| Ok(Value::Boolean(st.saved))),
            "protect" => self.protect(args),
            "unprotect" => self.unprotect(args),
            "protectstructure" => {
                with(app, |st| Ok(Value::Boolean(st.book.workbook_protection().is_some_and(|p| p.structure))))
            }
            "protectwindows" => {
                with(app, |st| Ok(Value::Boolean(st.book.workbook_protection().is_some_and(|p| p.windows))))
            }
            "save" | "saveas" | "savecopyas" | "close" | "refreshall" => {
                let file = opt_text(args, 0, "Filename")?;
                with_mut(app, |st| {
                    match file {
                        Some(f) if lname != "close" => st.log.push(format!("Workbook.{name} {f}")),
                        _ => st.log.push(format!("Workbook.{name}")),
                    }
                    if lname.starts_with("save") {
                        st.saved = true;
                    }
                    Ok(Value::Boolean(true))
                })
            }
            "windows" => Ok(window_value(app)),
            "application" => Ok(app_value(app)),
            "parent" => Ok(app_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "saved" => {
                let b = data(&value)?.to_bool()?;
                with_mut(&self.app, |st| {
                    st.saved = b;
                    Ok(())
                })
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

fn next_sheet_name(st: &mut State) -> String {
    loop {
        st.sheet_counter += 1;
        let n = format!("Sheet{}", st.sheet_counter);
        if st.sheet_by_name(&n).is_none() {
            return n;
        }
    }
}

/// 0-based tab position of a Before/After argument.
fn position_arg(app: &Shared, v: &Value) -> Result<usize> {
    let id = match decode_sheet(v) {
        Some(id) => id,
        None => match data(v)? {
            Value::String(s) => with(app, |st| st.sheet_by_name(&s).ok_or_else(RuntimeError::subscript))?,
            other => {
                let i = other.to_i64()?;
                with(app, |st| {
                    usize::try_from(i - 1)
                        .ok()
                        .and_then(|i| st.book.sheet_ids().get(i).copied())
                        .ok_or_else(RuntimeError::subscript)
                })?
            }
        },
    };
    with(app, |st| st.sheet_index(id).map(|i| i - 1).ok_or_else(RuntimeError::subscript))
}

/// Sheet name rules: the duplicate and invalid-name messages the application shows.
fn check_sheet_name(st: &State, name: &str, except: Option<SheetId>) -> Result<()> {
    if !crate::memory::valid_sheet_name(name) {
        return Err(err1004(BAD_NAME));
    }
    if st.sheet_by_name(name).is_some_and(|id| Some(id) != except) {
        return Err(err1004(NAME_TAKEN));
    }
    Ok(())
}

/// `Worksheets` / `Sheets`, or the group `Sheets(Array(...))`.
struct SheetsObj {
    app: Shared,
    subset: Option<Vec<SheetId>>,
}

impl SheetsObj {
    fn ids(&self, st: &State) -> Vec<SheetId> {
        match &self.subset {
            Some(v) => v.clone(),
            None => st.book.sheet_ids(),
        }
    }

    fn lookup(&self, st: &State, idx: &Value) -> Result<SheetId> {
        let ids = self.ids(st);
        match idx {
            Value::String(s) => st.sheet_by_name(s).filter(|id| ids.contains(id)).ok_or_else(RuntimeError::subscript),
            other => {
                let i = other.to_i64()?;
                usize::try_from(i - 1).ok().and_then(|i| ids.get(i).copied()).ok_or_else(RuntimeError::subscript)
            }
        }
    }
}

impl Object for SheetsObj {
    fn type_name(&self) -> String {
        "Sheets".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let app = &self.app;
        match name.to_ascii_lowercase().as_str() {
            "count" => with(app, |st| Ok(Value::Long(self.ids(st).len() as i32))),
            "" | "item" => {
                let idx = data(args.require(0, "Index")?)?;
                if let Value::Array(a) = &idx {
                    let items = a.borrow().data.clone();
                    let mut ids = Vec::new();
                    for v in items {
                        ids.push(with(app, |st| self.lookup(st, &data(&v)?))?);
                    }
                    return Ok(Value::Object(Some(Rc::new(SheetsObj { app: app.clone(), subset: Some(ids) }))));
                }
                let id = with(app, |st| self.lookup(st, &idx))?;
                Ok(sheet_value(app, id))
            }
            "add" | "add2" => {
                let pos = match (args.get(0, "Before"), args.get(1, "After")) {
                    (Some(b), _) => position_arg(app, b)?,
                    (None, Some(a)) => position_arg(app, a)? + 1,
                    (None, None) => with(app, |st| Ok(st.sheet_index(st.active_sheet).map_or(0, |i| i - 1)))?,
                };
                let count = opt_i64(args, 2, "Count")?.unwrap_or(1).max(1);
                let id = with_mut(app, |st| {
                    if st.structure_protected() {
                        return Err(err1004(STRUCTURE));
                    }
                    let mut first = None;
                    for i in 0..count {
                        let n = next_sheet_name(st);
                        let id = st.book.add_sheet(&n, pos + i as usize)?;
                        first.get_or_insert(id);
                    }
                    let first = first.ok_or_else(RuntimeError::object_required)?;
                    st.active_sheet = first;
                    st.saved = false;
                    Ok(first)
                })?;
                Ok(sheet_value(app, id))
            }
            "visible" => Ok(Value::Boolean(true)),
            "select" => {
                let first = with(app, |st| Ok(self.ids(st).first().copied()))?;
                if let Some(id) = first {
                    with_mut(app, |st| {
                        st.active_sheet = id;
                        Ok(())
                    })?;
                }
                Ok(Value::Boolean(true))
            }
            "application" => Ok(app_value(app)),
            "parent" => Ok(workbook_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "visible" => {
                if self.subset.is_some() {
                    return Err(err1004("Unable to set the Visible property of the Sheets class"));
                }
                let ids = with(&self.app, |st| Ok(st.book.sheet_ids()))?;
                for id in ids {
                    SheetObj { app: self.app.clone(), id }.assign("Visible", &Args::none(), value.clone())?;
                }
                Ok(())
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        let ids = with(&self.app, |st| Ok(self.ids(st)))?;
        Ok(ids.into_iter().map(|id| sheet_value(&self.app, id)).collect())
    }
}

pub(crate) struct SheetObj {
    pub app: Shared,
    pub id: SheetId,
}

impl SheetObj {
    fn activate(&self) -> Result<Value> {
        with_mut(&self.app, |st| {
            // A hidden sheet can be activated (CONFORMANCE sheet2 S016); a very hidden one cannot.
            if st.book.sheet_visibility(self.id) == SheetVisibility::VeryHidden {
                return Err(err1004("Activate method of Worksheet class failed"));
            }
            st.active_sheet = self.id;
            Ok(Value::Boolean(true))
        })
    }

    fn select(&self) -> Result<Value> {
        with_mut(&self.app, |st| {
            if st.book.sheet_visibility(self.id) != SheetVisibility::Visible {
                return Err(err1004("Select method of Worksheet class failed"));
            }
            st.active_sheet = self.id;
            Ok(Value::Boolean(true))
        })
    }

    fn delete(&self) -> Result<Value> {
        with_mut(&self.app, |st| {
            if st.structure_protected() {
                return Err(err1004(STRUCTURE));
            }
            let ids = st.book.sheet_ids();
            let visible = ids.iter().filter(|&&i| st.book.sheet_visibility(i) == SheetVisibility::Visible).count();
            if st.book.sheet_visibility(self.id) == SheetVisibility::Visible && visible <= 1 {
                return Err(err1004(ONE_VISIBLE));
            }
            let pos = ids.iter().position(|&i| i == self.id).unwrap_or(0);
            let name = st.sheet_name(self.id)?;
            st.apply_ref_edit(&RefEdit::DeleteSheet { name: name.clone() })?;
            st.book.delete_sheet(self.id)?;
            st.views.remove(&self.id);
            st.protection.remove(&self.id);
            st.filters.remove(&self.id);
            st.windows.remove(&self.id);
            if st.clipboard.as_ref().is_some_and(|c| c.sheet == self.id) {
                st.clipboard = None;
            }
            if st.active_sheet == self.id {
                let rest = st.book.sheet_ids();
                let vis: Vec<SheetId> =
                    rest.iter().copied().filter(|&i| st.book.sheet_visibility(i) == SheetVisibility::Visible).collect();
                let after = rest.iter().skip(pos).copied().find(|i| vis.contains(i));
                let before = rest.iter().take(pos).rev().copied().find(|i| vis.contains(i));
                st.active_sheet = after.or(before).or_else(|| rest.first().copied()).unwrap_or(0);
            }
            st.log.push(format!("Delete sheet {name}"));
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    fn copy(&self, args: &Args) -> Result<Value> {
        let pos = match (args.get(0, "Before"), args.get(1, "After")) {
            (Some(b), _) => Some(position_arg(&self.app, b)?),
            (None, Some(a)) => Some(position_arg(&self.app, a)? + 1),
            _ => None,
        };
        with_mut(&self.app, |st| {
            let name = st.sheet_name(self.id)?;
            let Some(pos) = pos else {
                st.log.push(format!("Copy sheet {name} to a new workbook"));
                return Ok(Value::Boolean(true));
            };
            if st.structure_protected() {
                return Err(err1004(STRUCTURE));
            }
            let mut n = 2;
            let new_name = loop {
                let base: String = name.chars().take(31 - format!(" ({n})").len()).collect();
                let cand = format!("{base} ({n})");
                if st.sheet_by_name(&cand).is_none() {
                    break cand;
                }
                n += 1;
            };
            let new_id = st.book.add_sheet(&new_name, pos)?;
            if let Some(u) = st.book.used_range(self.id) {
                for (r, c) in st.book.occupied_cells(self.id, u) {
                    match st.book.cell_formula(self.id, r, c) {
                        Some(f) => st.book.set_cell_formula(new_id, r, c, Some(f))?,
                        None => {
                            let v = st.book.cell_value(self.id, r, c);
                            st.book.set_cell_value(new_id, r, c, v)?;
                        }
                    }
                    let f = st.book.cell_format(self.id, r, c);
                    st.book.set_cell_format(new_id, r, c, f)?;
                    if let Some(t) = st.book.cell_comment(self.id, r, c) {
                        st.book.set_cell_comment(new_id, r, c, Some(t))?;
                    }
                    if let Some(l) = st.book.cell_hyperlink(self.id, r, c) {
                        st.book.set_cell_hyperlink(new_id, r, c, Some(l))?;
                    }
                }
            }
            for m in st.book.merged_areas(self.id) {
                st.book.merge_cells(new_id, m)?;
            }
            st.active_sheet = new_id;
            st.saved = false;
            Ok(Value::Boolean(true))
        })
    }

    fn protect(&self, args: &Args) -> Result<Value> {
        let password = opt_text(args, 0, "Password")?.unwrap_or_default();
        let drawing = opt_bool(args, 1, "DrawingObjects")?.unwrap_or(true);
        let contents = opt_bool(args, 2, "Contents")?.unwrap_or(true);
        let scenarios = opt_bool(args, 3, "Scenarios")?.unwrap_or(true);
        let ui_only = opt_bool(args, 4, "UserInterfaceOnly")?;
        let mut allow = [false; 11];
        for (i, n) in ALLOW_NAMES.iter().enumerate() {
            allow[i] = opt_bool(args, 5 + i, n)?.unwrap_or(false);
        }
        with_mut(&self.app, |st| {
            if let Some(u) = ui_only {
                st.ui_only.insert(self.id, u);
            }
            if st.book.is_protected(self.id) {
                // Protecting again keeps the password; the options follow the new call.
                st.protection.insert(self.id, Protection { contents, drawing, scenarios, allow });
                return Ok(Value::Boolean(true));
            }
            st.book.set_protected(self.id, true)?;
            st.book.set_sheet_password(self.id, crate::protect::new_password_hash(&password));
            st.protection.insert(self.id, Protection { contents, drawing, scenarios, allow });
            Ok(Value::Boolean(true))
        })
    }

    fn unprotect(&self, args: &Args) -> Result<Value> {
        let password = opt_text(args, 0, "Password")?;
        with_mut(&self.app, |st| {
            if !st.book.is_protected(self.id) {
                return Ok(Value::Boolean(true));
            }
            if let Some(h) = st.book.sheet_password(self.id)
                && !crate::protect::verify_password(&h, password.as_deref().unwrap_or(""))
            {
                return Err(err1004(BAD_PASSWORD));
            }
            st.book.set_protected(self.id, false)?;
            st.book.set_sheet_password(self.id, None);
            st.protection.remove(&self.id);
            Ok(Value::Boolean(true))
        })
    }

    fn protection_flag(&self, which: &str) -> Result<Value> {
        with(&self.app, |st| {
            if !st.book.is_protected(self.id) {
                return Ok(Value::Boolean(false));
            }
            let p = st.protection.get(&self.id).cloned().unwrap_or_default();
            Ok(Value::Boolean(match which {
                "protectcontents" => p.contents,
                "protectdrawingobjects" => p.drawing,
                _ => p.scenarios,
            }))
        })
    }

    fn rename(&self, n: &str) -> Result<()> {
        with_mut(&self.app, |st| {
            if st.structure_protected() {
                return Err(err1004(STRUCTURE));
            }
            check_sheet_name(st, n, Some(self.id))?;
            let old = st.sheet_name(self.id)?;
            st.book.rename_sheet(self.id, n)?;
            if old != n {
                st.apply_ref_edit(&RefEdit::RenameSheet { old, new: n.to_owned() })?;
            }
            Ok(())
        })
    }
}

impl Object for SheetObj {
    fn type_name(&self) -> String {
        "Worksheet".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        if lname == SHEET_REF.to_ascii_lowercase() {
            return Ok(Value::Long(self.id as i32));
        }
        let app = &self.app;
        with(app, |st| if st.sheet_exists(self.id) { Ok(()) } else { Err(RuntimeError::object_required()) })?;
        match lname.as_str() {
            "name" => with(app, |st| Ok(Value::from(st.sheet_name(self.id)?))),
            "codename" => with(app, |st| Ok(Value::from(st.book.code_name(self.id).unwrap_or_default()))),
            "range" | "cells" | "rows" | "columns" => sheet_member(app, self.id, name, args),
            "usedrange" => {
                let u = with_mut(app, |st| {
                    st.book.reset_last_cell(self.id);
                    Ok(st.used(self.id))
                })?;
                Ok(range_value(app, self.id, vec![u], Kind::Cells))
            }
            "activate" => self.activate(),
            "select" => self.select(),
            "protect" => self.protect(args),
            "unprotect" => self.unprotect(args),
            "protectcontents" | "protectdrawingobjects" | "protectscenarios" => self.protection_flag(&lname),
            "protectionmode" => with(app, |st| Ok(Value::Boolean(st.ui_only.get(&self.id).copied().unwrap_or(false)))),
            "protection" => Ok(Value::Object(Some(Rc::new(ProtectionObj { app: app.clone(), id: self.id })))),
            "visible" => with(app, |st| {
                Ok(Value::Long(match st.book.sheet_visibility(self.id) {
                    SheetVisibility::Visible => -1,
                    SheetVisibility::Hidden => 0,
                    SheetVisibility::VeryHidden => 2,
                }))
            }),
            "index" => with(app, |st| Ok(Value::Long(st.sheet_index(self.id).unwrap_or(0) as i32))),
            "delete" => self.delete(),
            "copy" => self.copy(args),
            "move" => {
                let pos = match (args.get(0, "Before"), args.get(1, "After")) {
                    (Some(b), _) => position_arg(app, b)?,
                    (None, Some(a)) => position_arg(app, a)? + 1,
                    _ => {
                        return with_mut(app, |st| {
                            st.log.push("Move sheet to a new workbook".into());
                            Ok(Value::Boolean(true))
                        });
                    }
                };
                with_mut(app, |st| {
                    if st.structure_protected() {
                        return Err(err1004(STRUCTURE));
                    }
                    let cur = st.sheet_index(self.id).map_or(0, |i| i - 1);
                    let target = if pos > cur { pos - 1 } else { pos };
                    st.book.move_sheet(self.id, target)?;
                    Ok(Value::Boolean(true))
                })
            }
            "paste" => {
                let dest = match args.get(0, "Destination") {
                    Some(d) => range_arg(app, self.id, d)?,
                    None => {
                        let sel = with(app, |st| Ok(st.view(self.id).selection))?;
                        (self.id, sel)
                    }
                };
                with_mut(app, |st| {
                    let clip = st.clipboard.clone().ok_or_else(|| err1004("Paste method of Worksheet class failed"))?;
                    if clip.cut {
                        st.clear_contents(clip.sheet, clip.src)?;
                        st.clear_formats(clip.sheet, clip.src)?;
                        st.clipboard = None;
                    }
                    st.paste(&clip, dest.0, dest.1[0], PasteKind::All, false, false, !clip.cut, -4142)?;
                    Ok(Value::Boolean(true))
                })
            }
            "calculate" => with_mut(app, |st| {
                st.book.calculate();
                Ok(Value::Boolean(true))
            }),
            "evaluate" => {
                let text = data(args.require(0, "Name")?)?.to_text()?;
                crate::functions::evaluate_text(app, self.id, &text)
            }
            "names" => crate::names::names_member(app, Some(self.id), args),
            "comments" => crate::notes::comments_value(app, self.id, args),
            "hyperlinks" => crate::notes::hyperlinks_value(app, self.id, None, args),
            "sort" => Ok(crate::sorting::sort_object(app, self.id)),
            "autofilter" => crate::filter::autofilter_object(app, self.id),
            "autofiltermode" => with(app, |st| Ok(Value::Boolean(st.filters.contains_key(&self.id)))),
            "filtermode" => {
                with(app, |st| Ok(Value::Boolean(st.filters.get(&self.id).is_some_and(|f| !f.conds.is_empty()))))
            }
            "showalldata" => crate::filter::show_all(app, self.id).map(|_| Value::Boolean(true)),
            "next" | "previous" => with(app, |st| {
                let ids = st.book.sheet_ids();
                let i = ids.iter().position(|&x| x == self.id).unwrap_or(0);
                let j = if lname == "next" { i.checked_add(1) } else { i.checked_sub(1) };
                Ok(match j.and_then(|j| ids.get(j)) {
                    Some(&id) => sheet_value(app, id),
                    None => Value::Object(None),
                })
            }),
            "type" => Ok(Value::Long(-4167)),
            "standardwidth" => Ok(Value::Double(crate::memory::DEFAULT_COLUMN_WIDTH)),
            "standardheight" => Ok(Value::Double(crate::memory::DEFAULT_ROW_HEIGHT)),
            "enableselection" => Ok(Value::Long(0)),
            "parent" => Ok(workbook_value(app)),
            "application" => Ok(app_value(app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        let v = data(&value)?;
        match name.to_ascii_lowercase().as_str() {
            "name" => {
                let n = v.to_text()?;
                self.rename(&n)
            }
            "visible" => {
                let vis = match &v {
                    Value::Boolean(true) => SheetVisibility::Visible,
                    Value::Boolean(false) => SheetVisibility::Hidden,
                    other => match other.to_i64()? {
                        -1 | 1 => SheetVisibility::Visible,
                        0 => SheetVisibility::Hidden,
                        2 => SheetVisibility::VeryHidden,
                        _ => return Err(err1004("Unable to set the Visible property of the Worksheet class")),
                    },
                };
                with_mut(&self.app, |st| {
                    if vis != SheetVisibility::Visible {
                        let others = st
                            .book
                            .sheet_ids()
                            .into_iter()
                            .filter(|&i| i != self.id && st.book.sheet_visibility(i) == SheetVisibility::Visible)
                            .collect::<Vec<_>>();
                        if others.is_empty() {
                            return Err(err1004("Unable to set the Visible property of the Worksheet class"));
                        }
                        if st.active_sheet == self.id {
                            st.active_sheet = others[0];
                        }
                    }
                    st.book.set_sheet_visibility(self.id, vis)
                })
            }
            "autofiltermode" => {
                if !v.to_bool()? {
                    crate::filter::remove(&self.app, self.id)?;
                }
                Ok(())
            }
            "enableselection" | "scrollarea" | "enablecalculation" | "displaypagebreaks" => Ok(()),
            "range" | "cells" => {
                // ws.Range("A1") = 5
                let target = sheet_member(&self.app, self.id, name, args)?;
                match target {
                    Value::Object(Some(o)) => o.assign("Value", &Args::none(), v),
                    _ => Err(RuntimeError::object_required()),
                }
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

/// `Worksheet.Protection`: what a protected sheet still allows.
struct ProtectionObj {
    app: Shared,
    id: SheetId,
}

impl Object for ProtectionObj {
    fn type_name(&self) -> String {
        "Protection".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        let lname = name.to_ascii_lowercase();
        let Some(i) = ALLOW_NAMES.iter().position(|n| *n == lname) else {
            return Err(RuntimeError::no_member(name));
        };
        with(&self.app, |st| {
            Ok(Value::Boolean(st.book.is_protected(self.id) && st.protection.get(&self.id).is_some_and(|p| p.allow[i])))
        })
    }
}
