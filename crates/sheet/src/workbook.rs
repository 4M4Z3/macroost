//! The storage a spreadsheet application provides. The object model reads and writes cells,
//! formats and sheet structure only through this trait, so any app (GridCraft first) can run
//! macros by implementing it. [`crate::MemoryWorkbook`] is the reference implementation.
//!
//! Coordinates are 1-based (`row` 1..=1,048,576, `col` 1..=16,384). Sheets are named by a stable
//! [`SheetId`] that survives renames and moves; tab order is `sheet_ids()`.

use macroost_lang::{Result, Value};

use crate::address::Rect;

/// Stable identity of a sheet within a workbook.
pub type SheetId = u32;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum SheetVisibility {
    #[default]
    Visible,
    Hidden,
    /// Hidden and not listed in the app's unhide dialog; only a macro can show it.
    VeryHidden,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum HAlign {
    #[default]
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterAcrossSelection,
    Distributed,
}

/// One border edge of a cell. `style` and `weight` use the object model's constants
/// (`xlContinuous` = 1, `xlDash` = -4115, ...; `xlThin` = 2, `xlMedium` = -4138, ...);
/// `color` is an RGB long (`red + green * 256 + blue * 65536`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct BorderLine {
    pub style: i32,
    pub weight: i32,
    pub color: u32,
}

impl Default for BorderLine {
    fn default() -> Self {
        Self { style: 1, weight: 2, color: 0 }
    }
}

/// Border edges of a cell; `None` is no border.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct Borders {
    pub left: Option<BorderLine>,
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub right: Option<BorderLine>,
    /// Diagonal from top-left to bottom-right (`xlDiagonalDown`).
    pub diagonal_down: Option<BorderLine>,
    /// Diagonal from bottom-left to top-right (`xlDiagonalUp`).
    pub diagonal_up: Option<BorderLine>,
}

/// Basic per-cell formatting. The default value is an unformatted cell.
#[derive(Clone, Debug, PartialEq, Default)]
pub struct CellFormat {
    /// Number format code; `None` is `General`.
    pub number_format: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// RGB long; `None` is automatic (black).
    pub font_color: Option<u32>,
    /// Points; `None` is the workbook default (11).
    pub font_size: Option<f64>,
    /// `None` is the workbook default font.
    pub font_name: Option<String>,
    /// Interior (fill) color as an RGB long; `None` is no fill.
    pub fill_color: Option<u32>,
    pub borders: Borders,
    pub h_align: HAlign,
    pub wrap: bool,
    pub strikethrough: bool,
    /// Double underline (`xlUnderlineStyleDouble`); `underline` is also set.
    pub double_underline: bool,
    /// Vertical alignment constant; `None` is bottom (`xlBottom`).
    pub v_align: Option<i32>,
    pub indent: u8,
    /// Text rotation in degrees (`Orientation`); 0 is horizontal.
    pub orientation: i32,
    pub shrink_to_fit: bool,
    /// Not locked (`Range.Locked = False`): stays editable on a protected sheet.
    pub unlocked: bool,
    /// Formula hidden on a protected sheet (`Range.FormulaHidden`).
    pub formula_hidden: bool,
}

impl CellFormat {
    pub fn is_default(&self) -> bool {
        *self == Self::default()
    }
}

/// A stored password verifier, as a file keeps it (ECMA-376 Part 1, 18.2.29 and 18.3.1.85):
/// the legacy 16-bit hash (`password="CC1A"`) or an iterated hash (`algorithmName`, `hashValue`,
/// `saltValue`, `spinCount`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PasswordHash {
    Legacy(u16),
    Modern { algorithm: String, hash: Vec<u8>, salt: Vec<u8>, spin_count: u32 },
}

/// Workbook-level protection (`Workbook.Protect`).
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct WorkbookProtection {
    pub structure: bool,
    pub windows: bool,
    pub password: Option<PasswordHash>,
}

/// A cell hyperlink.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Hyperlink {
    /// External target (URL, file, mail address); empty for a place in the workbook.
    pub address: String,
    /// Place in the workbook (`Sheet1!A1`).
    pub sub_address: String,
    pub screen_tip: String,
}

/// Which way cells move when a block is inserted or deleted.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ShiftDirection {
    Down,
    Right,
    Up,
    Left,
}

/// Workbook storage. Object-safe: the object model holds a `Box<dyn Workbook>`.
pub trait Workbook {
    /// File name shown by `Workbook.Name` (e.g. `Book1.xlsm`).
    fn name(&self) -> String;

    // ---- sheets ----
    /// Sheets in tab order.
    fn sheet_ids(&self) -> Vec<SheetId>;
    fn sheet_name(&self, sheet: SheetId) -> Option<String>;
    /// The sheet's code name (the name of its document module, e.g. `Sheet12`), if known.
    fn code_name(&self, sheet: SheetId) -> Option<String> {
        let _ = sheet;
        None
    }
    fn set_code_name(&mut self, sheet: SheetId, code_name: &str) {
        let _ = (sheet, code_name);
    }
    /// Defined names as (name, sheet scope or `None` for the workbook, refers-to text without the
    /// leading `=`, e.g. `Sheet1!$A$1:$B$5`).
    fn defined_names(&self) -> Vec<(String, Option<SheetId>, String)> {
        Vec::new()
    }
    /// Adds or replaces a defined name (matched case-insensitively within its scope). An empty
    /// `refers_to` deletes the name.
    fn set_defined_name(&mut self, name: &str, scope: Option<SheetId>, refers_to: &str) {
        let _ = (name, scope, refers_to);
    }
    /// Adds an empty sheet at `position` (0-based tab index, clamped to the end).
    fn add_sheet(&mut self, name: &str, position: usize) -> Result<SheetId>;
    fn delete_sheet(&mut self, sheet: SheetId) -> Result<()>;
    fn rename_sheet(&mut self, sheet: SheetId, name: &str) -> Result<()>;
    /// Moves a sheet to `position` (0-based tab index after removal, clamped).
    fn move_sheet(&mut self, sheet: SheetId, position: usize) -> Result<()>;
    fn sheet_visibility(&self, sheet: SheetId) -> SheetVisibility;
    fn set_sheet_visibility(&mut self, sheet: SheetId, visibility: SheetVisibility) -> Result<()>;
    fn is_protected(&self, sheet: SheetId) -> bool;
    fn set_protected(&mut self, sheet: SheetId, protected: bool) -> Result<()>;

    // ---- cells ----
    /// The cell's value (a formula cell's last computed value). `Empty` for a blank cell.
    fn cell_value(&self, sheet: SheetId, row: u32, col: u32) -> Value;
    /// Stores a constant (clearing any formula). `Empty` clears the content, keeping the format.
    fn set_cell_value(&mut self, sheet: SheetId, row: u32, col: u32, value: Value) -> Result<()>;
    /// Formula text including the leading `=`, or `None` for a constant or blank cell.
    fn cell_formula(&self, sheet: SheetId, row: u32, col: u32) -> Option<String>;
    /// Stores a formula (`Some("=A1+1")`); `None` clears the content.
    fn set_cell_formula(&mut self, sheet: SheetId, row: u32, col: u32, formula: Option<String>) -> Result<()>;
    fn cell_format(&self, sheet: SheetId, row: u32, col: u32) -> CellFormat;
    fn set_cell_format(&mut self, sheet: SheetId, row: u32, col: u32, format: CellFormat) -> Result<()>;
    /// Every stored cell (value, formula or non-default format) inside `area`, sorted by row then
    /// column. This is how the object model walks huge ranges without visiting blank cells.
    fn occupied_cells(&self, sheet: SheetId, area: Rect) -> Vec<(u32, u32)>;
    /// Bounding box of the stored cells, or `None` for an empty sheet.
    fn used_range(&self, sheet: SheetId) -> Option<Rect>;

    // ---- structure ----
    /// Inserts `count` blank rows before `at`, shifting cells and row properties down.
    fn insert_rows(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()>;
    fn delete_rows(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()>;
    fn insert_columns(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()>;
    fn delete_columns(&mut self, sheet: SheetId, at: u32, count: u32) -> Result<()>;
    /// Inserts a blank block the size of `area`, shifting the cells in its columns down
    /// ([`ShiftDirection::Down`]) or the cells in its rows right ([`ShiftDirection::Right`]).
    fn insert_cells(&mut self, sheet: SheetId, area: Rect, shift: ShiftDirection) -> Result<()>;
    /// Deletes the block `area`, shifting cells up ([`ShiftDirection::Up`]) or left
    /// ([`ShiftDirection::Left`]) into the gap.
    fn delete_cells(&mut self, sheet: SheetId, area: Rect, shift: ShiftDirection) -> Result<()>;

    // ---- rows and columns ----
    fn row_hidden(&self, sheet: SheetId, row: u32) -> bool;
    fn set_row_hidden(&mut self, sheet: SheetId, row: u32, hidden: bool) -> Result<()>;
    /// Points.
    fn row_height(&self, sheet: SheetId, row: u32) -> f64;
    fn set_row_height(&mut self, sheet: SheetId, row: u32, height: f64) -> Result<()>;
    fn column_hidden(&self, sheet: SheetId, col: u32) -> bool;
    fn set_column_hidden(&mut self, sheet: SheetId, col: u32, hidden: bool) -> Result<()>;
    /// Characters of the default font.
    fn column_width(&self, sheet: SheetId, col: u32) -> f64;
    fn set_column_width(&mut self, sheet: SheetId, col: u32, width: f64) -> Result<()>;

    // ---- optional ----
    fn merge_cells(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        let _ = (sheet, area);
        Ok(())
    }
    fn unmerge_cells(&mut self, sheet: SheetId, area: Rect) -> Result<()> {
        let _ = (sheet, area);
        Ok(())
    }
    fn merged_areas(&self, sheet: SheetId) -> Vec<Rect> {
        let _ = sheet;
        Vec::new()
    }
    /// True for a row an AutoFilter hid (`SUBTOTAL` skips those even for functions 1-11).
    fn row_filtered(&self, sheet: SheetId, row: u32) -> bool {
        let _ = (sheet, row);
        false
    }
    fn set_row_filtered(&mut self, sheet: SheetId, row: u32, filtered: bool) {
        let _ = (sheet, row, filtered);
    }
    /// Format of a whole row (`Rows(3).Font.Bold = True`), which cells without their own format
    /// show.
    fn row_format(&self, sheet: SheetId, row: u32) -> Option<CellFormat> {
        let _ = (sheet, row);
        None
    }
    fn set_row_format(&mut self, sheet: SheetId, row: u32, format: Option<CellFormat>) {
        let _ = (sheet, row, format);
    }
    /// Format of a whole column.
    fn column_format(&self, sheet: SheetId, col: u32) -> Option<CellFormat> {
        let _ = (sheet, col);
        None
    }
    fn set_column_format(&mut self, sheet: SheetId, col: u32, format: Option<CellFormat>) {
        let _ = (sheet, col, format);
    }
    /// The "last cell" (`SpecialCells(xlCellTypeLastCell)`): like the used range, but it keeps
    /// cells that were cleared until [`Workbook::reset_last_cell`] (the application resets it when
    /// a macro reads `UsedRange`).
    fn last_cell(&self, sheet: SheetId) -> Option<Rect> {
        self.used_range(sheet)
    }
    fn reset_last_cell(&mut self, sheet: SheetId) {
        let _ = sheet;
    }
    /// Manual calculation (`Application.Calculation = xlCalculationManual`): formula cells keep
    /// the value of their last calculation until [`Workbook::calculate`].
    fn set_manual_calculation(&mut self, manual: bool) {
        let _ = manual;
    }
    /// The password verifier of a protected sheet, if it has a password.
    fn sheet_password(&self, sheet: SheetId) -> Option<PasswordHash> {
        let _ = sheet;
        None
    }
    fn set_sheet_password(&mut self, sheet: SheetId, password: Option<PasswordHash>) {
        let _ = (sheet, password);
    }
    /// Workbook-level protection, or `None` when unprotected.
    fn workbook_protection(&self) -> Option<WorkbookProtection> {
        None
    }
    fn set_workbook_protection(&mut self, protection: Option<WorkbookProtection>) {
        let _ = protection;
    }
    /// The cell's comment (note) text.
    fn cell_comment(&self, sheet: SheetId, row: u32, col: u32) -> Option<String> {
        let _ = (sheet, row, col);
        None
    }
    fn set_cell_comment(&mut self, sheet: SheetId, row: u32, col: u32, text: Option<String>) -> Result<()> {
        let _ = (sheet, row, col, text);
        Ok(())
    }
    fn cell_hyperlink(&self, sheet: SheetId, row: u32, col: u32) -> Option<Hyperlink> {
        let _ = (sheet, row, col);
        None
    }
    fn set_cell_hyperlink(&mut self, sheet: SheetId, row: u32, col: u32, link: Option<Hyperlink>) -> Result<()> {
        let _ = (sheet, row, col, link);
        Ok(())
    }
    /// Recalculates formulas (`Application.Calculate`). Storage without a calculation engine
    /// leaves values as they are.
    fn calculate(&mut self) {}

    /// True when the storage itself rewrites formula references and defined names on structural
    /// edits (inserted/deleted rows and columns, moved blocks, renamed sheets), as a spreadsheet
    /// application's own model does. The object model then leaves references alone instead of
    /// adjusting them a second time.
    fn adjusts_references(&self) -> bool {
        false
    }

    /// Lets a host get its concrete storage back after a run (to save it).
    fn as_any(&self) -> Option<&dyn std::any::Any> {
        None
    }
}
