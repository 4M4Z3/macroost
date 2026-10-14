//! `Shapes`, `Shape`, `ShapeRange`, `Placeholders`, `PlaceholderFormat`, `FillFormat`,
//! `LineFormat`, their `ColorFormat`s, and tables (`Table`, `Rows`/`Row`, `Columns`/`Column`,
//! `Cell`).

use macroost_lang::{Args, Result, RuntimeError, Value};

use crate::model::{Shape, ShapeKind, Slide, TableData, TextBody, emu_to_pt, pt_to_emu};
use crate::objects::{Kind, PresentationObj, call_value, obj, read_only};
use crate::state::{
    ShapeRef, Shared, State, app_err, arg_f64, arg_i64, assigned_f64, data, index_in, list_arg, long, opt_i64, single,
    text, to_tri, tri, with, with_mut,
};
use crate::text::{self, Target};

/// Offset of a duplicated shape from its original (12 points right and down).
const DUPLICATE_OFFSET_EMU: i64 = 152_400;
/// `AddTable` without Width / Height: 640 points wide, 29.2-point rows.
const DEFAULT_TABLE_WIDTH_EMU: i64 = 8_128_000;
const DEFAULT_TABLE_ROW_EMU: i64 = 370_840;
/// HRESULT E_FAIL (0x80004005) as the error number macros see.
const E_FAIL: i32 = -2_147_467_259;

/// True when `name` is the default name the host gave a shape of `stem` ("Rectangle 3").
fn is_default_name(name: &str, stem: &str) -> bool {
    name.strip_prefix(stem)
        .and_then(|rest| rest.strip_prefix(' '))
        .is_some_and(|n| !n.is_empty() && n.chars().all(|c| c.is_ascii_digit()))
}

pub(crate) fn call(o: &PresentationObj, n: &str, name: &str, args: &Args) -> Result<Value> {
    let app = &o.app;
    match &o.kind {
        Kind::Shapes(pk, sid) => shapes(app, *pk, *sid, n, args),
        Kind::Placeholders(pk, sid) => placeholders(app, *pk, *sid, n, args),
        Kind::Shape(r) => shape(app, r, n, args),
        Kind::ShapeRange(pk, sid, ids) => shape_range(app, *pk, *sid, ids, n, args),
        Kind::PlaceholderFormat(r) => match n {
            "type" => with(app, |st| match st.shape(r)?.kind {
                ShapeKind::Placeholder(t) => Ok(long(t)),
                _ => Err(app_err("the shape is not a placeholder")),
            }),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::Fill(r) => match n {
            "visible" => with(app, |st| Ok(tri(st.shape(r)?.fill.visible))),
            "forecolor" => Ok(obj(app, Kind::FillColor(r.clone()))),
            "transparency" => Ok(single(0.0)),
            "solid" => with_mut(app, |st| {
                st.shape_mut(r)?.fill.visible = true;
                Ok(Value::Empty)
            }),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::Line(r) => match n {
            "visible" => with(app, |st| Ok(tri(st.shape(r)?.line.visible))),
            "forecolor" => Ok(obj(app, Kind::LineColor(r.clone()))),
            "weight" => with(app, |st| Ok(single(st.shape(r)?.line.weight))),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::FillColor(r) | Kind::LineColor(r) => {
            let fill = matches!(o.kind, Kind::FillColor(_));
            match n {
                "rgb" | "" => with(app, |st| {
                    let s = st.shape(r)?;
                    Ok(long(if fill { s.fill.color } else { s.line.color }))
                }),
                "type" => Ok(long(1)),
                _ => Err(RuntimeError::no_member(name)),
            }
        }
        Kind::Table(r) => table(app, r, n, args),
        Kind::Rows(r) => rows(app, r, n, args),
        Kind::Columns(r) => columns(app, r, n, args),
        Kind::Row(r, i) => match n {
            "height" => with(app, |st| Ok(single(emu_to_pt(table_of(st, r)?.rows.get(*i).copied().unwrap_or(0))))),
            "delete" => edit_table(app, r, |t| {
                if t.rows.len() <= 1 || *i >= t.rows.len() {
                    return Err(app_err("cannot delete the row"));
                }
                t.rows.remove(*i);
                t.cells.remove(*i);
                Ok(())
            }),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::Column(r, i) => match n {
            "width" => with(app, |st| Ok(single(emu_to_pt(table_of(st, r)?.cols.get(*i).copied().unwrap_or(0))))),
            "delete" => edit_table(app, r, |t| {
                if t.cols.len() <= 1 || *i >= t.cols.len() {
                    return Err(app_err("cannot delete the column"));
                }
                t.cols.remove(*i);
                for row in &mut t.cells {
                    row.remove(*i);
                }
                Ok(())
            }),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::Cell(r, row, col) => match n {
            "shape" => Ok(obj(app, Kind::CellShape(r.clone(), *row, *col))),
            "select" => Ok(Value::Empty),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::CellShape(r, row, col) => match n {
            "textframe" | "textframe2" => Ok(obj(app, Kind::TextFrame(Target::Cell(r.clone(), *row, *col)))),
            "hastextframe" => Ok(tri(true)),
            _ => Err(RuntimeError::no_member(name)),
        },
        _ => text::call(o, n, name, args),
    }
}

pub(crate) fn assign(o: &PresentationObj, n: &str, name: &str, args: &Args, value: Value) -> Result<()> {
    let app = &o.app;
    match &o.kind {
        Kind::Shape(r) => shape_assign(app, r, n, value),
        Kind::ShapeRange(pk, sid, ids) => {
            for id in ids {
                shape_assign(app, &ShapeRef { pk: *pk, sid: *sid, id: *id }, n, value.clone())?;
            }
            Ok(())
        }
        Kind::Fill(r) if n == "visible" => {
            let b = to_tri(&value)?;
            with_mut(app, |st| {
                st.shape_mut(r)?.fill.visible = b;
                Ok(())
            })
        }
        Kind::Line(r) if n == "visible" || n == "weight" => {
            let v = data(&value)?;
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                if n == "visible" {
                    s.line.visible = to_tri(&v)?;
                } else {
                    s.line.weight = v.to_f64()? as f32;
                }
                Ok(())
            })
        }
        Kind::FillColor(r) | Kind::LineColor(r) if n == "rgb" || n.is_empty() => {
            let c = data(&value)?.to_i32()?;
            let fill = matches!(o.kind, Kind::FillColor(_));
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                if fill {
                    s.fill.color = c;
                    s.fill.visible = true;
                } else {
                    s.line.color = c;
                    s.line.visible = true;
                }
                Ok(())
            })
        }
        Kind::Row(r, i) if n == "height" => {
            let h = pt_to_emu(assigned_f64(&value)?);
            edit_table(app, r, |t| {
                let slot = t.rows.get_mut(*i).ok_or_else(|| app_err("no such row"))?;
                *slot = h.max(crate::model::MIN_ROW_HEIGHT_EMU);
                Ok(())
            })
            .map(|_| ())
        }
        Kind::Column(r, i) if n == "width" => {
            let w = pt_to_emu(assigned_f64(&value)?);
            edit_table(app, r, |t| {
                let slot = t.cols.get_mut(*i).ok_or_else(|| app_err("no such column"))?;
                *slot = w;
                Ok(())
            })
            .map(|_| ())
        }
        Kind::Shapes(..)
        | Kind::Placeholders(..)
        | Kind::PlaceholderFormat(_)
        | Kind::Fill(_)
        | Kind::Line(_)
        | Kind::FillColor(_)
        | Kind::LineColor(_)
        | Kind::Table(_)
        | Kind::Rows(_)
        | Kind::Columns(_)
        | Kind::Row(..)
        | Kind::Column(..)
        | Kind::Cell(..)
        | Kind::CellShape(..) => Err(read_only(name)),
        _ => text::assign(o, n, name, args, value),
    }
}

pub(crate) fn items(o: &PresentationObj) -> Result<Vec<Value>> {
    let app = &o.app;
    match &o.kind {
        Kind::Shapes(pk, sid) => {
            let ids = with(app, |st| Ok(st.slide(*pk, *sid)?.shapes.iter().map(|s| s.id).collect::<Vec<_>>()))?;
            Ok(ids.into_iter().map(|id| obj(app, Kind::Shape(ShapeRef { pk: *pk, sid: *sid, id }))).collect())
        }
        Kind::Placeholders(pk, sid) => {
            let ids = with(app, |st| Ok(placeholder_ids(st.slide(*pk, *sid)?)))?;
            Ok(ids.into_iter().map(|id| obj(app, Kind::Shape(ShapeRef { pk: *pk, sid: *sid, id }))).collect())
        }
        Kind::ShapeRange(pk, sid, ids) => {
            Ok(ids.iter().map(|id| obj(app, Kind::Shape(ShapeRef { pk: *pk, sid: *sid, id: *id }))).collect())
        }
        Kind::Rows(r) => {
            let n = with(app, |st| Ok(table_of(st, r)?.rows.len()))?;
            Ok((0..n).map(|i| obj(app, Kind::Row(r.clone(), i))).collect())
        }
        Kind::Columns(r) => {
            let n = with(app, |st| Ok(table_of(st, r)?.cols.len()))?;
            Ok((0..n).map(|i| obj(app, Kind::Column(r.clone(), i))).collect())
        }
        _ => text::items(o),
    }
}

// ---- Shapes ----

fn find_shape(slide: &Slide, v: &Value) -> Result<u32> {
    match data(v)? {
        Value::String(s) => slide
            .shapes
            .iter()
            .find(|x| x.name.eq_ignore_ascii_case(&s))
            .map(|x| x.id)
            .ok_or_else(|| app_err(format!("no shape named {s}"))),
        other => {
            let i = index_in(other.to_i64()?, slide.shapes.len())?;
            Ok(slide.shapes[i].id)
        }
    }
}

fn placeholder_ids(slide: &Slide) -> Vec<u32> {
    slide.shapes.iter().filter(|s| matches!(s.kind, ShapeKind::Placeholder(_))).map(|s| s.id).collect()
}

fn geometry(args: &Args, first: usize) -> Result<(i64, i64, i64, i64)> {
    Ok((
        pt_to_emu(arg_f64(args, first, "Left")?),
        pt_to_emu(arg_f64(args, first + 1, "Top")?),
        pt_to_emu(arg_f64(args, first + 2, "Width")?),
        pt_to_emu(arg_f64(args, first + 3, "Height")?),
    ))
}

fn add_shape(app: &Shared, pk: u32, sid: u32, make: impl FnOnce(u32) -> Shape) -> Result<Value> {
    let id = with_mut(app, |st| {
        let slide = st.slide_mut(pk, sid)?;
        let id = slide.take_id();
        slide.shapes.push(make(id));
        st.pres_mut(pk)?.set_saved(false);
        Ok(id)
    })?;
    Ok(obj(app, Kind::Shape(ShapeRef { pk, sid, id })))
}

fn shapes(app: &Shared, pk: u32, sid: u32, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(st.slide(pk, sid)?.shapes.len()))),
        "item" | "" => {
            let v = args.require(0, "Index")?.clone();
            let id = with(app, |st| find_shape(st.slide(pk, sid)?, &v))?;
            Ok(obj(app, Kind::Shape(ShapeRef { pk, sid, id })))
        }
        "addtextbox" | "addlabel" => {
            let orientation = arg_i64(args, 0, "Orientation")? as i32;
            let geo = geometry(args, 1)?;
            let label = n == "addlabel";
            add_shape(app, pk, sid, |id| {
                let mut s = Shape::textbox(id, geo);
                s.frame.orientation = orientation;
                if label {
                    // ORACLE: a label does not wrap.
                    s.frame.word_wrap = false;
                }
                s
            })
        }
        "addshape" => {
            let t = arg_i64(args, 0, "Type")? as i32;
            let geo = geometry(args, 1)?;
            add_shape(app, pk, sid, |id| Shape::autoshape(id, t, geo))
        }
        "addtable" => {
            let rows = arg_i64(args, 0, "NumRows")?;
            let cols = arg_i64(args, 1, "NumColumns")?;
            if !(1..=75).contains(&rows) || !(1..=75).contains(&cols) {
                return Err(app_err("a table has 1 to 75 rows and columns"));
            }
            let pt = |i: usize, name: &str| -> Result<Option<i64>> {
                Ok(args.get(i, name).map(|v| data(v).and_then(|v| v.to_f64())).transpose()?.map(pt_to_emu))
            };
            // Without a size the host makes the table 640 points wide with 29.2-point rows,
            // centered on the slide (ORACLE: the centering).
            let w = pt(4, "Width")?.unwrap_or(DEFAULT_TABLE_WIDTH_EMU);
            let h = pt(5, "Height")?.unwrap_or(DEFAULT_TABLE_ROW_EMU * rows);
            let (sw, sh) = with(app, |st| Ok(st.pres(pk)?.slide_size()))?;
            let x = pt(2, "Left")?.unwrap_or((sw - w) / 2);
            let y = pt(3, "Top")?.unwrap_or((sh - h) / 2);
            add_shape(app, pk, sid, |id| Shape::table(id, rows as usize, cols as usize, (x, y, w, h)))
        }
        "range" => {
            let ids = with(app, |st| {
                let slide = st.slide(pk, sid)?;
                match args.get(0, "Index") {
                    None => Ok(slide.shapes.iter().map(|s| s.id).collect::<Vec<_>>()),
                    Some(v) => list_arg(v).iter().map(|x| find_shape(slide, x)).collect(),
                }
            })?;
            Ok(obj(app, Kind::ShapeRange(pk, sid, ids)))
        }
        "hastitle" => with(app, |st| Ok(tri(title_id(st.slide(pk, sid)?).is_some()))),
        "title" => {
            let id = with(app, |st| title_id(st.slide(pk, sid)?).ok_or_else(|| app_err("the slide has no title")))?;
            Ok(obj(app, Kind::Shape(ShapeRef { pk, sid, id })))
        }
        "placeholders" => {
            let c = obj(app, Kind::Placeholders(pk, sid));
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "selectall" => Ok(Value::Empty),
        "parent" => Ok(obj(app, Kind::Slide(pk, sid))),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn title_id(slide: &Slide) -> Option<u32> {
    slide.shapes.iter().find(|s| matches!(s.kind, ShapeKind::Placeholder(1 | 3 | 5))).map(|s| s.id)
}

fn placeholders(app: &Shared, pk: u32, sid: u32, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(placeholder_ids(st.slide(pk, sid)?).len()))),
        "item" | "" => {
            let i = arg_i64(args, 0, "Index")?;
            let id = with(app, |st| {
                let ids = placeholder_ids(st.slide(pk, sid)?);
                Ok(ids[index_in(i, ids.len())?])
            })?;
            Ok(obj(app, Kind::Shape(ShapeRef { pk, sid, id })))
        }
        _ => Err(RuntimeError::no_member(n)),
    }
}

// ---- Shape ----

fn shape(app: &Shared, r: &ShapeRef, n: &str, args: &Args) -> Result<Value> {
    match n {
        "name" => with(app, |st| Ok(text(st.shape(r)?.name.as_str()))),
        "id" => with(app, |st| Ok(long(st.shape(r)?.id))),
        "type" => with(app, |st| Ok(long(st.shape(r)?.kind.shape_type()))),
        "left" => with(app, |st| Ok(single(emu_to_pt(st.shape(r)?.x)))),
        "top" => with(app, |st| Ok(single(emu_to_pt(st.shape(r)?.y)))),
        "width" => with(app, |st| Ok(single(emu_to_pt(st.shape(r)?.cx)))),
        "height" => with(app, |st| Ok(single(emu_to_pt(st.shape(r)?.cy)))),
        "rotation" => with(app, |st| Ok(single(st.shape(r)?.rotation))),
        "visible" => with(app, |st| Ok(tri(st.shape(r)?.visible))),
        "horizontalflip" => with(app, |st| Ok(tri(st.shape(r)?.flip_h))),
        "verticalflip" => with(app, |st| Ok(tri(st.shape(r)?.flip_v))),
        "autoshapetype" => with(app, |st| {
            Ok(long(match st.shape(r)?.kind {
                ShapeKind::AutoShape(t) => t,
                ShapeKind::Table => -2,
                _ => 1,
            }))
        }),
        "hastextframe" => with(app, |st| Ok(tri(st.shape(r)?.text.is_some()))),
        "hastable" => with(app, |st| Ok(tri(st.shape(r)?.table.is_some()))),
        // Every shape has a TextFrame object, even one without text (a table): only its text
        // is unavailable.
        "textframe" | "textframe2" => {
            with(app, |st| st.shape(r).map(|_| ()))?;
            Ok(obj(app, Kind::TextFrame(Target::Shape(r.clone()))))
        }
        "table" => {
            with(app, |st| table_of(st, r).map(|_| ()).map_err(|_| RuntimeError::new(E_FAIL, "Unspecified error")))?;
            Ok(obj(app, Kind::Table(r.clone())))
        }
        "fill" => Ok(obj(app, Kind::Fill(r.clone()))),
        "line" => Ok(obj(app, Kind::Line(r.clone()))),
        "placeholderformat" => {
            with(app, |st| match st.shape(r)?.kind {
                ShapeKind::Placeholder(_) => Ok(()),
                _ => Err(app_err("the shape is not a placeholder")),
            })?;
            Ok(obj(app, Kind::PlaceholderFormat(r.clone())))
        }
        "zorderposition" => with(app, |st| {
            let s = st.slide(r.pk, r.sid)?;
            s.shapes.iter().position(|x| x.id == r.id).map(|i| long(i + 1)).ok_or_else(|| app_err("deleted"))
        }),
        "zorder" => {
            let cmd = arg_i64(args, 0, "ZOrderCmd")?;
            with_mut(app, |st| {
                let s = st.slide_mut(r.pk, r.sid)?;
                let i = s.shapes.iter().position(|x| x.id == r.id).ok_or_else(|| app_err("deleted"))?;
                let last = s.shapes.len() - 1;
                let to = match cmd {
                    0 => last,
                    1 => 0,
                    2 => (i + 1).min(last),
                    3 => i.saturating_sub(1),
                    _ => return Err(app_err("invalid ZOrder command")),
                };
                let sh = s.shapes.remove(i);
                s.shapes.insert(to, sh);
                Ok(Value::Empty)
            })
        }
        "delete" => with_mut(app, |st| {
            let s = st.slide_mut(r.pk, r.sid)?;
            let i = s.shapes.iter().position(|x| x.id == r.id).ok_or_else(|| app_err("deleted"))?;
            s.shapes.remove(i);
            st.pres_mut(r.pk)?.set_saved(false);
            Ok(Value::Empty)
        }),
        "duplicate" => {
            let id = with_mut(app, |st| {
                let s = st.slide_mut(r.pk, r.sid)?;
                let mut copy = s.shape(r.id).ok_or_else(|| app_err("deleted"))?.clone();
                let id = s.take_id();
                copy.id = id;
                // A default name follows the new id; a name the macro gave is kept.
                if is_default_name(&copy.name, copy.kind.name_stem()) {
                    copy.name = format!("{} {}", copy.kind.name_stem(), id - 1);
                }
                copy.x += DUPLICATE_OFFSET_EMU;
                copy.y += DUPLICATE_OFFSET_EMU;
                s.shapes.push(copy);
                Ok(id)
            })?;
            Ok(obj(app, Kind::ShapeRange(r.pk, r.sid, vec![id])))
        }
        "incrementleft" | "incrementtop" => {
            let d = pt_to_emu(arg_f64(args, 0, "Increment")?);
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                if n == "incrementleft" {
                    s.x += d
                } else {
                    s.y += d
                }
                Ok(Value::Empty)
            })
        }
        "incrementrotation" => {
            let d = arg_f64(args, 0, "Increment")? as f32;
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                s.rotation = (s.rotation + d).rem_euclid(360.0);
                Ok(Value::Empty)
            })
        }
        "flip" => {
            let cmd = arg_i64(args, 0, "FlipCmd")?;
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                if cmd == 0 {
                    s.flip_h = !s.flip_h
                } else {
                    s.flip_v = !s.flip_v
                }
                Ok(Value::Empty)
            })
        }
        "select" => Ok(Value::Empty),
        "parent" => Ok(obj(app, Kind::Slide(r.pk, r.sid))),
        "application" => Ok(obj(app, Kind::Application)),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn shape_assign(app: &Shared, r: &ShapeRef, n: &str, value: Value) -> Result<()> {
    let v = data(&value)?;
    with_mut(app, |st| {
        let s = st.shape_mut(r)?;
        match n {
            "name" => s.name = v.to_text()?,
            "left" => s.x = pt_to_emu(v.to_f64()?),
            "top" => s.y = pt_to_emu(v.to_f64()?),
            "width" => {
                let w = pt_to_emu(v.to_f64()?);
                if let Some(t) = &mut s.table {
                    scale(&mut t.cols, w);
                    s.cx = t.width();
                } else {
                    s.cx = w;
                }
            }
            "height" => {
                let h = pt_to_emu(v.to_f64()?);
                if let Some(t) = &mut s.table {
                    scale(&mut t.rows, h);
                    s.cy = t.height();
                } else {
                    s.cy = h;
                }
            }
            "rotation" => s.rotation = (v.to_f64()? as f32).rem_euclid(360.0),
            "visible" => s.visible = to_tri(&v)?,
            "autoshapetype" => match &mut s.kind {
                ShapeKind::AutoShape(t) => *t = v.to_i32()?,
                _ => return Err(app_err("the shape is not an AutoShape")),
            },
            _ => return Err(read_only(n)),
        }
        Ok(())
    })
}

/// Rescales sizes proportionally so they add up to `total` (each rounded to an EMU).
fn scale(parts: &mut [i64], total: i64) {
    let sum: i64 = parts.iter().sum();
    if sum <= 0 {
        return;
    }
    for p in parts.iter_mut() {
        *p = (*p as f64 * total as f64 / sum as f64).round() as i64;
    }
}

fn shape_range(app: &Shared, pk: u32, sid: u32, ids: &[u32], n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => Ok(long(ids.len())),
        "item" | "" => {
            let v = data(args.require(0, "Index")?)?;
            let id = match v {
                Value::String(s) => with(app, |st| {
                    let slide = st.slide(pk, sid)?;
                    ids.iter()
                        .copied()
                        .find(|id| slide.shape(*id).is_some_and(|x| x.name.eq_ignore_ascii_case(&s)))
                        .ok_or_else(|| app_err(format!("no shape named {s}")))
                })?,
                other => ids[index_in(other.to_i64()?, ids.len())?],
            };
            Ok(obj(app, Kind::Shape(ShapeRef { pk, sid, id })))
        }
        "delete" | "select" | "zorder" => {
            for id in ids {
                shape(app, &ShapeRef { pk, sid, id: *id }, n, args)?;
            }
            Ok(Value::Empty)
        }
        "duplicate" => {
            let mut out = Vec::new();
            for id in ids {
                if let Some(Kind::ShapeRange(_, _, new)) =
                    crate::objects::kind_of(&shape(app, &ShapeRef { pk, sid, id: *id }, n, args)?)
                {
                    out.extend(new);
                }
            }
            Ok(obj(app, Kind::ShapeRange(pk, sid, out)))
        }
        _ => {
            let first = *ids.first().ok_or_else(|| app_err("the range is empty"))?;
            shape(app, &ShapeRef { pk, sid, id: first }, n, args)
        }
    }
}

// ---- tables ----

pub(crate) fn table_of<'a>(st: &'a State, r: &ShapeRef) -> Result<&'a TableData> {
    st.shape(r)?.table.as_ref().ok_or_else(|| app_err("the shape has no table"))
}

/// Edits a table and keeps the shape's size equal to the sum of its rows and columns.
fn edit_table(app: &Shared, r: &ShapeRef, f: impl FnOnce(&mut TableData) -> Result<()>) -> Result<Value> {
    with_mut(app, |st| {
        let s = st.shape_mut(r)?;
        let t = s.table.as_mut().ok_or_else(|| app_err("the shape has no table"))?;
        f(t)?;
        s.cx = t.width();
        s.cy = t.height();
        Ok(Value::Empty)
    })
}

fn table(app: &Shared, r: &ShapeRef, n: &str, args: &Args) -> Result<Value> {
    match n {
        "rows" => {
            let c = obj(app, Kind::Rows(r.clone()));
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "columns" => {
            let c = obj(app, Kind::Columns(r.clone()));
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "cell" => {
            let row = arg_i64(args, 0, "Row")?;
            let col = arg_i64(args, 1, "Column")?;
            let (ri, ci) = with(app, |st| {
                let t = table_of(st, r)?;
                Ok((index_in(row, t.rows.len())?, index_in(col, t.cols.len())?))
            })?;
            Ok(obj(app, Kind::Cell(r.clone(), ri, ci)))
        }
        "parent" => Ok(obj(app, Kind::Shape(r.clone()))),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn rows(app: &Shared, r: &ShapeRef, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(table_of(st, r)?.rows.len()))),
        "item" | "" => {
            let i = arg_i64(args, 0, "Index")?;
            let i = with(app, |st| index_in(i, table_of(st, r)?.rows.len()))?;
            Ok(obj(app, Kind::Row(r.clone(), i)))
        }
        "add" => {
            let before = opt_i64(args, 0, "BeforeRow")?;
            let mut at = 0;
            edit_table(app, r, |t| {
                at = match before {
                    Some(b) => index_in(b, t.rows.len())?,
                    None => t.rows.len(),
                };
                let h = t.rows.last().copied().unwrap_or(crate::model::MIN_ROW_HEIGHT_EMU);
                let fmt = TableData::cell_format(at.max(1));
                t.rows.insert(at, h);
                t.cells.insert(at, (0..t.cols.len()).map(|_| TextBody::new(fmt.clone())).collect());
                Ok(())
            })?;
            Ok(obj(app, Kind::Row(r.clone(), at)))
        }
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn columns(app: &Shared, r: &ShapeRef, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(table_of(st, r)?.cols.len()))),
        "item" | "" => {
            let i = arg_i64(args, 0, "Index")?;
            let i = with(app, |st| index_in(i, table_of(st, r)?.cols.len()))?;
            Ok(obj(app, Kind::Column(r.clone(), i)))
        }
        "add" => {
            let before = opt_i64(args, 0, "BeforeColumn")?;
            let mut at = 0;
            edit_table(app, r, |t| {
                at = match before {
                    Some(b) => index_in(b, t.cols.len())?,
                    None => t.cols.len(),
                };
                // The table keeps its width: every column shrinks to make room.
                let total = t.width();
                let n = t.cols.len() as f64 + 1.0;
                for c in t.cols.iter_mut() {
                    *c = (*c as f64 * (n - 1.0) / n).round() as i64;
                }
                t.cols.insert(at, (total as f64 / n).round() as i64);
                for (ri, row) in t.cells.iter_mut().enumerate() {
                    row.insert(at, TextBody::new(TableData::cell_format(ri)));
                }
                Ok(())
            })?;
            Ok(obj(app, Kind::Column(r.clone(), at)))
        }
        _ => Err(RuntimeError::no_member(n)),
    }
}
