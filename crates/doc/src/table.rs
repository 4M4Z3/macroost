//! `Tables`, `Table`, `Rows`/`Row`, `Columns`/`Column`, `Cells`/`Cell`.
//!
//! A table is a run of positions tagged with its id: cell contents, an end-of-cell mark after
//! each cell and an end-of-row mark after each row. Geometry is derived from those marks on
//! every access, so it always agrees with the text.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::document::{Item, Marker, ParaFormat};
use crate::objects::app_value;
use crate::range::{decode_range, range_value};
use crate::state::{
    DocEntry, Shared, TableGeom, data, no_member_of_collection, req_i64, table_by_id, tables, with, with_mut,
};

fn object(o: impl Object + 'static) -> Value {
    Value::Object(Some(Rc::new(o)))
}

fn index(v: &Value, count: usize) -> Result<usize> {
    let n = data(v)?.to_i64()?;
    if n < 1 || n as u64 > count as u64 {
        return Err(no_member_of_collection());
    }
    Ok(n as usize - 1)
}

fn marker(e: &DocEntry, table: u32, kind: Marker, at: usize) -> Item {
    let mut it = Item::char('\u{7}', e.inherited_font(at));
    it.table = Some(table);
    it.marker = kind;
    it.para = ParaFormat::default();
    it
}

/// `Tables.Add`: the range's text is replaced by a `rows` x `cols` table; a range starting
/// inside a paragraph splits it first.
pub(crate) fn add_table(e: &mut DocEntry, a: usize, b: usize, rows: usize, cols: usize) -> Result<u32> {
    if rows == 0 || cols == 0 || cols > 63 || rows > 32_767 {
        return Err(RuntimeError::new(4198, "Command failed"));
    }
    if e.doc.item(a).is_some_and(|i| i.table.is_some()) {
        return Err(RuntimeError::new(
            4605,
            "This method or property is not available because the object refers to a table.",
        ));
    }
    if a < b {
        e.delete(a, b);
    }
    let mut p = a.min(e.doc.len().saturating_sub(1));
    if p > 0 && e.doc.item(p - 1).is_some_and(|i| !i.ends_paragraph()) {
        p += e.insert_text(p, "\r", None);
    }
    let id = e.doc.new_table_id();
    let mut items = Vec::with_capacity(rows * (cols + 1));
    for _ in 0..rows {
        for _ in 0..cols {
            items.push(marker(e, id, Marker::CellEnd, p));
        }
        items.push(marker(e, id, Marker::RowEnd, p));
    }
    e.insert_items(p, items);
    Ok(id)
}

pub(crate) struct TablesObj {
    pub app: Shared,
    pub doc: u32,
    pub scope: Option<(usize, usize)>,
}

impl TablesObj {
    fn list(&self) -> Result<Vec<TableGeom>> {
        with(&self.app, |st| {
            Ok(tables(st.doc(self.doc)?)
                .into_iter()
                .filter(|t| match self.scope {
                    None => true,
                    Some((s, e)) => t.start < e.max(s + 1) && t.end > s,
                })
                .collect())
        })
    }
}

impl Object for TablesObj {
    fn type_name(&self) -> String {
        "Tables".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.list()?.len() as i32)),
            "" | "item" => {
                let list = self.list()?;
                let t = &list[index(args.require(0, "Index")?, list.len())?];
                Ok(object(TableObj { app: self.app.clone(), doc: self.doc, id: t.id }))
            }
            "add" => {
                let (d, a, b) = decode_range(args.require(0, "Range")?).ok_or_else(RuntimeError::type_mismatch)?;
                if d != self.doc {
                    return Err(RuntimeError::new(4198, "Command failed"));
                }
                let rows = usize::try_from(req_i64(args, 1, "NumRows")?)
                    .map_err(|_| RuntimeError::new(4198, "Command failed"))?;
                let cols = usize::try_from(req_i64(args, 2, "NumColumns")?)
                    .map_err(|_| RuntimeError::new(4198, "Command failed"))?;
                let id = with_mut(&self.app, |st| add_table(st.entry_mut(self.doc)?, a, b, rows, cols))?;
                Ok(object(TableObj { app: self.app.clone(), doc: self.doc, id }))
            }
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(self
            .list()?
            .into_iter()
            .map(|t| object(TableObj { app: self.app.clone(), doc: self.doc, id: t.id }))
            .collect())
    }
}

#[derive(Clone)]
pub(crate) struct TableObj {
    pub app: Shared,
    pub doc: u32,
    pub id: u32,
}

impl TableObj {
    fn geom(&self) -> Result<TableGeom> {
        with(&self.app, |st| table_by_id(st.doc(self.doc)?, self.id))
    }

    fn cell(&self, row: usize, col: usize) -> CellObj {
        CellObj { table: self.clone(), row, col }
    }

    fn cell_span(&self, row: usize, col: usize) -> Result<(usize, usize)> {
        let g = self.geom()?;
        let r = g.rows.get(row).ok_or_else(no_member_of_collection)?;
        r.cells.get(col).copied().ok_or_else(no_member_of_collection)
    }

    fn add_row(&self, before: Option<usize>) -> Result<()> {
        let g = self.geom()?;
        let cols = g.columns();
        let at = match before {
            Some(i) => g.rows.get(i).ok_or_else(no_member_of_collection)?.start,
            None => g.end,
        };
        with_mut(&self.app, |st| {
            let e = st.entry_mut(self.doc)?;
            let mut items: Vec<Item> = (0..cols).map(|_| marker(e, self.id, Marker::CellEnd, at)).collect();
            items.push(marker(e, self.id, Marker::RowEnd, at));
            e.insert_items(at, items);
            Ok(())
        })
    }

    fn add_column(&self, before: Option<usize>) -> Result<()> {
        let g = self.geom()?;
        with_mut(&self.app, |st| {
            let e = st.entry_mut(self.doc)?;
            for r in g.rows.iter().rev() {
                let at = match before {
                    Some(c) => r.cells.get(c).map(|x| x.0).unwrap_or(r.end - 1),
                    None => r.end - 1,
                };
                let it = marker(e, self.id, Marker::CellEnd, at);
                e.insert_items(at, vec![it]);
            }
            Ok(())
        })
    }

    fn delete_row(&self, row: usize) -> Result<()> {
        let g = self.geom()?;
        let r = g.rows.get(row).ok_or_else(no_member_of_collection)?;
        with_mut(&self.app, |st| {
            st.entry_mut(self.doc)?.remove_raw(r.start, r.end);
            Ok(())
        })
    }

    fn delete_column(&self, col: usize) -> Result<()> {
        let g = self.geom()?;
        if col >= g.columns() {
            return Err(no_member_of_collection());
        }
        if g.columns() == 1 {
            return with_mut(&self.app, |st| {
                st.entry_mut(self.doc)?.remove_raw(g.start, g.end);
                Ok(())
            });
        }
        with_mut(&self.app, |st| {
            let e = st.entry_mut(self.doc)?;
            for r in g.rows.iter().rev() {
                if let Some(&(a, b)) = r.cells.get(col) {
                    e.remove_raw(a, b);
                }
            }
            Ok(())
        })
    }
}

impl Object for TableObj {
    fn type_name(&self) -> String {
        "Table".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "rows" => {
                let coll: Rc<dyn Object> = Rc::new(RowsObj { table: self.clone() });
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "columns" => {
                let coll: Rc<dyn Object> = Rc::new(ColumnsObj { table: self.clone() });
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "cell" => {
                let g = self.geom()?;
                let r = req_i64(args, 0, "Row")?;
                let c = req_i64(args, 1, "Column")?;
                let row = usize::try_from(r - 1).map_err(|_| no_member_of_collection())?;
                let col = usize::try_from(c - 1).map_err(|_| no_member_of_collection())?;
                if g.rows.get(row).is_none_or(|x| col >= x.cells.len()) {
                    return Err(no_member_of_collection());
                }
                Ok(object(self.cell(row, col)))
            }
            "range" => {
                let g = self.geom()?;
                range_value(&self.app, self.doc, g.start, g.end)
            }
            "delete" => {
                let g = self.geom()?;
                with_mut(&self.app, |st| {
                    st.entry_mut(self.doc)?.remove_raw(g.start, g.end);
                    Ok(Value::Empty)
                })
            }
            "select" => {
                let g = self.geom()?;
                with_mut(&self.app, |st| {
                    st.active = self.doc;
                    *st.entry_mut(self.doc)?.selection.borrow_mut() = crate::state::Span { start: g.start, end: g.end };
                    Ok(Value::Empty)
                })
            }
            "uniform" => {
                let g = self.geom()?;
                let c = g.columns();
                Ok(Value::Boolean(g.rows.iter().all(|r| r.cells.len() == c)))
            }
            "nestinglevel" => Ok(Value::Long(1)),
            "autofitbehavior" | "autoformat" => Ok(Value::Empty),
            "application" => Ok(app_value(&self.app)),
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            // Presentation only: accepted, nothing to store.
            "style" | "allowautofit" | "preferredwidth" | "preferredwidthtype" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

pub(crate) struct RowsObj {
    table: TableObj,
}

impl Object for RowsObj {
    fn type_name(&self) -> String {
        "Rows".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.table.geom()?.rows.len() as i32)),
            "" | "item" => {
                let n = self.table.geom()?.rows.len();
                let i = index(args.require(0, "Index")?, n)?;
                Ok(object(RowObj { table: self.table.clone(), index: i }))
            }
            "first" => Ok(object(RowObj { table: self.table.clone(), index: 0 })),
            "last" => {
                let n = self.table.geom()?.rows.len();
                Ok(object(RowObj { table: self.table.clone(), index: n.saturating_sub(1) }))
            }
            "add" => {
                let before = match args.get(0, "BeforeRow") {
                    Some(Value::Object(Some(o))) => Some(
                        usize::try_from(o.call("Index", &Args::none())?.to_i64()? - 1)
                            .map_err(|_| no_member_of_collection())?,
                    ),
                    _ => None,
                };
                self.table.add_row(before)?;
                let g = self.table.geom()?;
                let i = before.unwrap_or(g.rows.len().saturating_sub(1));
                Ok(object(RowObj { table: self.table.clone(), index: i }))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        let n = self.table.geom()?.rows.len();
        Ok((0..n).map(|i| object(RowObj { table: self.table.clone(), index: i })).collect())
    }
}

pub(crate) struct RowObj {
    table: TableObj,
    index: usize,
}

impl Object for RowObj {
    fn type_name(&self) -> String {
        "Row".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "cells" => {
                let coll: Rc<dyn Object> = Rc::new(CellsObj::row(&self.table, self.index)?);
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "range" => {
                let g = self.table.geom()?;
                let r = g.rows.get(self.index).ok_or_else(no_member_of_collection)?;
                range_value(&self.table.app, self.table.doc, r.start, r.end)
            }
            "index" => Ok(Value::Long(self.index as i32 + 1)),
            "isfirst" => Ok(Value::Boolean(self.index == 0)),
            "islast" => Ok(Value::Boolean(self.index + 1 == self.table.geom()?.rows.len())),
            "delete" => {
                self.table.delete_row(self.index)?;
                Ok(Value::Empty)
            }
            "select" => {
                let g = self.table.geom()?;
                let r = g.rows.get(self.index).ok_or_else(no_member_of_collection)?;
                let (s, e) = (r.start, r.end);
                with_mut(&self.table.app, |st| {
                    *st.entry_mut(self.table.doc)?.selection.borrow_mut() = crate::state::Span { start: s, end: e };
                    Ok(Value::Empty)
                })
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "height" | "heightrule" | "headingformat" | "allowbreakacrosspages" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

pub(crate) struct ColumnsObj {
    table: TableObj,
}

impl Object for ColumnsObj {
    fn type_name(&self) -> String {
        "Columns".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.table.geom()?.columns() as i32)),
            "" | "item" => {
                let n = self.table.geom()?.columns();
                let i = index(args.require(0, "Index")?, n)?;
                Ok(object(ColumnObj { table: self.table.clone(), index: i }))
            }
            "first" => Ok(object(ColumnObj { table: self.table.clone(), index: 0 })),
            "last" => {
                let n = self.table.geom()?.columns();
                Ok(object(ColumnObj { table: self.table.clone(), index: n.saturating_sub(1) }))
            }
            "add" => {
                let before = match args.get(0, "BeforeColumn") {
                    Some(Value::Object(Some(o))) => Some(
                        usize::try_from(o.call("Index", &Args::none())?.to_i64()? - 1)
                            .map_err(|_| no_member_of_collection())?,
                    ),
                    _ => None,
                };
                self.table.add_column(before)?;
                let n = self.table.geom()?.columns();
                Ok(object(ColumnObj { table: self.table.clone(), index: before.unwrap_or(n.saturating_sub(1)) }))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        let n = self.table.geom()?.columns();
        Ok((0..n).map(|i| object(ColumnObj { table: self.table.clone(), index: i })).collect())
    }
}

pub(crate) struct ColumnObj {
    table: TableObj,
    index: usize,
}

impl Object for ColumnObj {
    fn type_name(&self) -> String {
        "Column".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "cells" => {
                let coll: Rc<dyn Object> = Rc::new(CellsObj::column(&self.table, self.index)?);
                if args.is_empty() { Ok(Value::Object(Some(coll))) } else { coll.call("Item", args) }
            }
            "index" => Ok(Value::Long(self.index as i32 + 1)),
            "isfirst" => Ok(Value::Boolean(self.index == 0)),
            "islast" => Ok(Value::Boolean(self.index + 1 == self.table.geom()?.columns())),
            "delete" => {
                self.table.delete_column(self.index)?;
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "width" | "preferredwidth" | "preferredwidthtype" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

pub(crate) struct CellObj {
    table: TableObj,
    row: usize,
    col: usize,
}

impl Object for CellObj {
    fn type_name(&self) -> String {
        "Cell".into()
    }

    fn call(&self, name: &str, _args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "range" => {
                let (s, e) = self.table.cell_span(self.row, self.col)?;
                range_value(&self.table.app, self.table.doc, s, e)
            }
            "rowindex" => Ok(Value::Long(self.row as i32 + 1)),
            "columnindex" => Ok(Value::Long(self.col as i32 + 1)),
            "select" => {
                let (s, e) = self.table.cell_span(self.row, self.col)?;
                with_mut(&self.table.app, |st| {
                    *st.entry_mut(self.table.doc)?.selection.borrow_mut() = crate::state::Span { start: s, end: e };
                    Ok(Value::Empty)
                })
            }
            "next" => {
                let g = self.table.geom()?;
                let (r, c) = if self.col + 1 < g.rows.get(self.row).map_or(0, |x| x.cells.len()) {
                    (self.row, self.col + 1)
                } else {
                    (self.row + 1, 0)
                };
                if g.rows.get(r).is_some_and(|x| c < x.cells.len()) {
                    Ok(object(CellObj { table: self.table.clone(), row: r, col: c }))
                } else {
                    Ok(Value::Object(None))
                }
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn assign(&self, name: &str, _args: &Args, _value: Value) -> Result<()> {
        match name.to_ascii_lowercase().as_str() {
            "width" | "height" | "verticalalignment" | "preferredwidth" => Ok(()),
            _ => Err(RuntimeError::no_member(name)),
        }
    }
}

/// A list of cells (of a row, a column or a range).
pub(crate) struct CellsObj {
    cells: Vec<CellObj>,
}

impl CellsObj {
    fn row(t: &TableObj, row: usize) -> Result<Self> {
        let g = t.geom()?;
        let r = g.rows.get(row).ok_or_else(no_member_of_collection)?;
        Ok(Self { cells: (0..r.cells.len()).map(|c| t.cell(row, c)).collect() })
    }

    fn column(t: &TableObj, col: usize) -> Result<Self> {
        let g = t.geom()?;
        Ok(Self {
            cells: g
                .rows
                .iter()
                .enumerate()
                .filter(|(_, r)| col < r.cells.len())
                .map(|(i, _)| t.cell(i, col))
                .collect(),
        })
    }

    /// Cells overlapping `start..end`.
    pub(crate) fn in_range(app: &Shared, doc: u32, start: usize, end: usize) -> Result<Self> {
        let geoms = with(app, |st| Ok(tables(st.doc(doc)?)))?;
        let mut cells = Vec::new();
        for g in geoms {
            let t = TableObj { app: app.clone(), doc, id: g.id };
            for (ri, r) in g.rows.iter().enumerate() {
                for (ci, &(a, b)) in r.cells.iter().enumerate() {
                    let hit = if start == end { a <= start && start < b } else { a < end && b > start };
                    if hit {
                        cells.push(t.cell(ri, ci));
                    }
                }
            }
        }
        if cells.is_empty() {
            return Err(RuntimeError::new(
                4605,
                "This method or property is not available because some or all of the object does not refer to a table.",
            ));
        }
        Ok(Self { cells })
    }
}

impl Object for CellsObj {
    fn type_name(&self) -> String {
        "Cells".into()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.cells.len() as i32)),
            "" | "item" => {
                let i = index(args.require(0, "Index")?, self.cells.len())?;
                let c = &self.cells[i];
                Ok(object(CellObj { table: c.table.clone(), row: c.row, col: c.col }))
            }
            _ => Err(RuntimeError::no_member(name)),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        Ok(self.cells.iter().map(|c| object(CellObj { table: c.table.clone(), row: c.row, col: c.col })).collect())
    }
}
