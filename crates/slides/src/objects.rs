//! Every object a macro can hold is a [`PresentationObj`]: a handle (presentation key, SlideID, shape id,
//! text span...) into the shared state, resolved on each access, so objects stay valid while
//! slides and shapes are added, moved and deleted around them. `Application`, `Presentations`,
//! `Presentation`, `Slides`, `Slide`, `SlideRange`, `PageSetup`, `SlideShowSettings`, the master
//! and its layouts live here; shapes and tables in `shapes.rs`, text in `text.rs`.

use std::rc::Rc;

use macroost_lang::{Args, Object, Result, RuntimeError, Value};

use crate::model::{MemoryPresentation, Slide, builtin_layout, emu_to_pt, pt_to_emu};
use crate::state::{
    PRESENTATION_ERROR, ShapeRef, Shared, State, app_err, arg_i64, assigned_f64, data, index_in, list_arg, long,
    single, text, to_tri, tri, with, with_mut,
};
use crate::text::{RangeRef, Target};

/// What a [`PresentationObj`] refers to.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) enum Kind {
    Application,
    Presentations,
    Presentation(u32),
    Slides(u32),
    Slide(u32, u32),
    SlideRange(u32, Vec<u32>),
    Shapes(u32, u32),
    Placeholders(u32, u32),
    Shape(ShapeRef),
    ShapeRange(u32, u32, Vec<u32>),
    PlaceholderFormat(ShapeRef),
    TextFrame(Target),
    TextRange(RangeRef),
    Font(RangeRef),
    ParagraphFormat(RangeRef),
    FontColor(RangeRef),
    Fill(ShapeRef),
    FillColor(ShapeRef),
    Line(ShapeRef),
    LineColor(ShapeRef),
    Table(ShapeRef),
    Rows(ShapeRef),
    Row(ShapeRef, usize),
    Columns(ShapeRef),
    Column(ShapeRef, usize),
    Cell(ShapeRef, usize, usize),
    CellShape(ShapeRef, usize, usize),
    PageSetup(u32),
    SlideShowSettings(u32),
    SlideMaster(u32),
    CustomLayouts(u32),
    CustomLayout(u32, usize),
}

pub struct PresentationObj {
    pub(crate) app: Shared,
    pub(crate) kind: Kind,
}

/// Wraps `kind` as an object value. Presentations, slides and shapes keep their identity while
/// any macro variable holds them (`Is` compares identity).
pub(crate) fn obj(app: &Shared, kind: Kind) -> Value {
    let cached = matches!(kind, Kind::Presentation(_) | Kind::Slide(..) | Kind::Shape(_));
    if cached {
        if let Ok(st) = app.try_borrow()
            && let Some(o) = st.cache.get(&kind).and_then(|w| w.upgrade())
        {
            return Value::Object(Some(o));
        }
    }
    let o = Rc::new(PresentationObj { app: app.clone(), kind: kind.clone() });
    if cached && let Ok(mut st) = app.try_borrow_mut() {
        st.cache.retain(|_, w| w.strong_count() > 0);
        st.cache.insert(kind, Rc::downgrade(&o));
    }
    Value::Object(Some(o))
}

/// The handle behind an object argument, when it is one of ours.
pub(crate) fn kind_of(v: &Value) -> Option<Kind> {
    let Value::Object(Some(o)) = v else { return None };
    o.as_any().and_then(|a| a.downcast_ref::<PresentationObj>()).map(|p| p.kind.clone())
}

pub(crate) fn read_only(name: &str) -> RuntimeError {
    RuntimeError::new(PRESENTATION_ERROR, format!("{name} is read-only"))
}

pub(crate) fn call_value(v: &Value, name: &str, args: &Args) -> Result<Value> {
    match v {
        Value::Object(Some(o)) => o.call(name, args),
        _ => Err(RuntimeError::object_required()),
    }
}

impl Object for PresentationObj {
    fn type_name(&self) -> String {
        match &self.kind {
            Kind::Application => "Application",
            Kind::Presentations => "Presentations",
            Kind::Presentation(_) => "Presentation",
            Kind::Slides(_) => "Slides",
            Kind::Slide(..) => "Slide",
            Kind::SlideRange(..) => "SlideRange",
            Kind::Shapes(..) => "Shapes",
            Kind::Placeholders(..) => "Placeholders",
            Kind::Shape(_) | Kind::CellShape(..) => "Shape",
            Kind::ShapeRange(..) => "ShapeRange",
            Kind::PlaceholderFormat(_) => "PlaceholderFormat",
            Kind::TextFrame(_) => "TextFrame",
            Kind::TextRange(_) => "TextRange",
            Kind::Font(_) => "Font",
            Kind::ParagraphFormat(_) => "ParagraphFormat",
            Kind::FontColor(_) | Kind::FillColor(_) | Kind::LineColor(_) => "ColorFormat",
            Kind::Fill(_) => "FillFormat",
            Kind::Line(_) => "LineFormat",
            Kind::Table(_) => "Table",
            Kind::Rows(_) => "Rows",
            Kind::Row(..) => "Row",
            Kind::Columns(_) => "Columns",
            Kind::Column(..) => "Column",
            Kind::Cell(..) => "Cell",
            Kind::PageSetup(_) => "PageSetup",
            Kind::SlideShowSettings(_) => "SlideShowSettings",
            Kind::SlideMaster(_) => "Master",
            Kind::CustomLayouts(_) => "CustomLayouts",
            Kind::CustomLayout(..) => "CustomLayout",
        }
        .to_owned()
    }

    fn call(&self, name: &str, args: &Args) -> Result<Value> {
        let n = name.to_ascii_lowercase();
        let app = &self.app;
        match &self.kind {
            Kind::Application => application(app, &n, args),
            Kind::Presentations => presentations(app, &n, args),
            Kind::Presentation(pk) => presentation(app, *pk, &n, args),
            Kind::Slides(pk) => slides(app, *pk, &n, args),
            Kind::Slide(pk, sid) => slide(app, *pk, *sid, &n, args),
            Kind::SlideRange(pk, ids) => slide_range(app, *pk, ids, &n, args),
            Kind::PageSetup(pk) => page_setup(app, *pk, &n),
            Kind::SlideShowSettings(pk) => slide_show(app, *pk, &n),
            Kind::SlideMaster(pk) => match n.as_str() {
                "customlayouts" => {
                    let c = obj(app, Kind::CustomLayouts(*pk));
                    if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
                }
                "name" => Ok(text("Default Theme")),
                "parent" => Ok(obj(app, Kind::Presentation(*pk))),
                _ => Err(RuntimeError::no_member(name)),
            },
            Kind::CustomLayouts(pk) => match n.as_str() {
                "count" => with(app, |st| Ok(long(st.pres(*pk)?.layouts().len()))),
                "item" | "" => {
                    let i = arg_i64(args, 0, "Index")?;
                    let i = with(app, |st| index_in(i, st.pres(*pk)?.layouts().len()))?;
                    Ok(obj(app, Kind::CustomLayout(*pk, i)))
                }
                _ => Err(RuntimeError::no_member(name)),
            },
            Kind::CustomLayout(pk, i) => match n.as_str() {
                "name" => with(app, |st| {
                    st.pres(*pk)?.layouts().get(*i).map(|l| text(l.name.as_str())).ok_or_else(|| app_err("no layout"))
                }),
                "parent" => Ok(obj(app, Kind::SlideMaster(*pk))),
                _ => Err(RuntimeError::no_member(name)),
            },
            _ => crate::shapes::call(self, &n, name, args),
        }
    }

    fn assign(&self, name: &str, args: &Args, value: Value) -> Result<()> {
        let n = name.to_ascii_lowercase();
        let app = &self.app;
        match &self.kind {
            Kind::Presentation(pk) => match n.as_str() {
                "saved" => {
                    let b = to_tri(&value)?;
                    with_mut(app, |st| {
                        st.pres_mut(*pk)?.set_saved(b);
                        Ok(())
                    })
                }
                _ => Err(read_only(name)),
            },
            Kind::Slide(pk, sid) => slide_assign(app, *pk, *sid, &n, value),
            Kind::SlideRange(pk, ids) => {
                for sid in ids {
                    slide_assign(app, *pk, *sid, &n, value.clone())?;
                }
                Ok(())
            }
            Kind::PageSetup(pk) => page_setup_assign(app, *pk, &n, value),
            Kind::SlideShowSettings(pk) => slide_show_assign(app, *pk, &n, value),
            Kind::CustomLayout(pk, i) if n == "name" => {
                let s = data(&value)?.to_text()?;
                with_mut(app, |st| {
                    let l = st.pres_mut(*pk)?.layouts_mut().get_mut(*i).ok_or_else(|| app_err("no layout"))?;
                    l.name = s;
                    Ok(())
                })
            }
            Kind::Application
            | Kind::Presentations
            | Kind::Slides(_)
            | Kind::SlideMaster(_)
            | Kind::CustomLayouts(_)
            | Kind::CustomLayout(..) => Err(read_only(name)),
            _ => crate::shapes::assign(self, &n, name, args, value),
        }
    }

    fn default_value(&self) -> Result<Value> {
        match &self.kind {
            Kind::TextRange(_) => self.call("Text", &Args::none()),
            Kind::Presentation(_) | Kind::Slide(..) | Kind::Shape(_) | Kind::CustomLayout(..) => {
                self.call("Name", &Args::none())
            }
            _ => Err(RuntimeError::no_member("(default)")),
        }
    }

    fn items(&self) -> Result<Vec<Value>> {
        let app = &self.app;
        match &self.kind {
            Kind::Presentations => {
                let keys = with(app, |st| Ok(st.pres.iter().map(|(k, _)| *k).collect::<Vec<_>>()))?;
                Ok(keys.into_iter().map(|k| obj(app, Kind::Presentation(k))).collect())
            }
            Kind::Slides(pk) => {
                let ids = with(app, |st| Ok(st.pres(*pk)?.slides().iter().map(|s| s.id).collect::<Vec<_>>()))?;
                Ok(ids.into_iter().map(|id| obj(app, Kind::Slide(*pk, id))).collect())
            }
            Kind::SlideRange(pk, ids) => Ok(ids.iter().map(|id| obj(app, Kind::Slide(*pk, *id))).collect()),
            Kind::CustomLayouts(pk) => {
                let n = with(app, |st| Ok(st.pres(*pk)?.layouts().len()))?;
                Ok((0..n).map(|i| obj(app, Kind::CustomLayout(*pk, i))).collect())
            }
            _ => crate::shapes::items(self),
        }
    }

    fn as_any(&self) -> Option<&dyn std::any::Any> {
        Some(self)
    }
}

// ---- Application / Presentations / Presentation ----

pub(crate) fn active_presentation(app: &Shared) -> Result<Value> {
    let key = with(app, |st| {
        st.active
            .filter(|k| st.pres.iter().any(|(p, _)| p == k))
            .or_else(|| st.pres.first().map(|(k, _)| *k))
            .ok_or_else(|| app_err("Invalid request.  There is no active presentation."))
    })?;
    Ok(obj(app, Kind::Presentation(key)))
}

fn application(app: &Shared, n: &str, args: &Args) -> Result<Value> {
    match n {
        "presentations" => {
            let c = obj(app, Kind::Presentations);
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "activepresentation" => active_presentation(app),
        "version" => Ok(text("16.0")),
        "application" => Ok(obj(app, Kind::Application)),
        "quit" => with_mut(app, |st| {
            st.log.push("Application.Quit".into());
            Ok(Value::Empty)
        }),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn presentations(app: &Shared, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(st.pres.len()))),
        "item" | "" => {
            let v = data(args.require(0, "Index")?)?;
            let key = with(app, |st| match &v {
                Value::String(s) => st
                    .pres
                    .iter()
                    .find(|(_, p)| p.name().eq_ignore_ascii_case(s))
                    .map(|(k, _)| *k)
                    .ok_or_else(|| app_err(format!("no presentation named {s}"))),
                other => {
                    let i = index_in(other.to_i64()?, st.pres.len())?;
                    Ok(st.pres[i].0)
                }
            })?;
            Ok(obj(app, Kind::Presentation(key)))
        }
        "add" => {
            let key = with_mut(app, |st| {
                st.untitled += 1;
                let name = format!("Presentation{}", st.untitled);
                Ok(st.add(Box::new(MemoryPresentation::from_default_template(&name))))
            })?;
            Ok(obj(app, Kind::Presentation(key)))
        }
        "open" | "open2007" => Err(app_err("opening files is not available to macros")),
        "application" | "parent" => Ok(obj(app, Kind::Application)),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn presentation(app: &Shared, pk: u32, n: &str, args: &Args) -> Result<Value> {
    match n {
        "name" | "fullname" => with(app, |st| Ok(text(st.pres(pk)?.name()))),
        "path" => Ok(text("")),
        "slides" => {
            let c = obj(app, Kind::Slides(pk));
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "pagesetup" => Ok(obj(app, Kind::PageSetup(pk))),
        "slideshowsettings" => Ok(obj(app, Kind::SlideShowSettings(pk))),
        "slidemaster" | "designs" => Ok(obj(app, Kind::SlideMaster(pk))),
        "saved" => with(app, |st| Ok(tri(st.pres(pk)?.saved()))),
        "readonly" => Ok(tri(false)),
        "save" | "saveas" | "savecopyas" => with_mut(app, |st| {
            let name = st.pres(pk)?.name();
            st.log.push(format!("Save {name}"));
            st.pres_mut(pk)?.set_saved(true);
            Ok(Value::Empty)
        }),
        "close" => with_mut(app, |st| {
            st.pres(pk)?;
            st.pres.retain(|(k, _)| *k != pk);
            if st.active == Some(pk) {
                st.active = st.pres.last().map(|(k, _)| *k);
            }
            Ok(Value::Empty)
        }),
        "application" => Ok(obj(app, Kind::Application)),
        "parent" => Ok(obj(app, Kind::Presentations)),
        _ => Err(RuntimeError::no_member(n)),
    }
}

// ---- Slides ----

/// SlideID for a new slide: one more than the largest in use.
fn next_slide_id(slides: &[Slide]) -> u32 {
    slides.iter().map(|s| s.id + 1).max().unwrap_or(256).max(256)
}

/// Name for a new slide: `Slide<n>` with n = SlideID - 255, or the next free number.
fn slide_name(slides: &[Slide], id: u32) -> String {
    let mut n = id.saturating_sub(255).max(1);
    loop {
        let name = format!("Slide{n}");
        if !slides.iter().any(|s| s.name.eq_ignore_ascii_case(&name)) {
            return name;
        }
        n += 1;
    }
}

fn find_slide(st: &State, pk: u32, v: &Value) -> Result<u32> {
    let slides = st.pres(pk)?.slides();
    match data(v)? {
        Value::String(s) => slides
            .iter()
            .find(|x| x.name.eq_ignore_ascii_case(&s))
            .map(|x| x.id)
            .ok_or_else(|| app_err(format!("no slide named {s}"))),
        other => {
            let i = index_in(other.to_i64()?, slides.len())?;
            Ok(slides[i].id)
        }
    }
}

fn add_slide(app: &Shared, pk: u32, index: i64, layout: LayoutChoice) -> Result<Value> {
    let id = with_mut(app, |st| {
        let p = st.pres_mut(pk)?;
        let count = p.slides().len();
        if index < 1 || index as u64 > count as u64 + 1 {
            return Err(app_err(format!("index {index} is out of range (1 to {})", count + 1)));
        }
        let lay = match layout {
            LayoutChoice::Type(t) => match p.layouts().iter().find(|l| l.layout_type == t) {
                Some(l) => l.clone(),
                None => {
                    let l = builtin_layout(t);
                    p.layouts_mut().push(l.clone());
                    l
                }
            },
            LayoutChoice::Index(i) => p.layouts().get(i).cloned().ok_or_else(|| app_err("no such layout"))?,
        };
        let id = next_slide_id(p.slides());
        let name = slide_name(p.slides(), id);
        let slide = Slide::new(id, name, &lay);
        p.slides_mut().insert((index - 1) as usize, slide);
        p.set_saved(false);
        Ok(id)
    })?;
    Ok(obj(app, Kind::Slide(pk, id)))
}

enum LayoutChoice {
    Type(i32),
    Index(usize),
}

fn slide_ids(app: &Shared, pk: u32, arg: Option<&Value>) -> Result<Vec<u32>> {
    with(app, |st| match arg {
        None => Ok(st.pres(pk)?.slides().iter().map(|s| s.id).collect()),
        Some(v) => list_arg(v).iter().map(|x| find_slide(st, pk, x)).collect(),
    })
}

fn slides(app: &Shared, pk: u32, n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => with(app, |st| Ok(long(st.pres(pk)?.slides().len()))),
        "item" | "" => {
            let v = args.require(0, "Index")?.clone();
            let id = with(app, |st| find_slide(st, pk, &v))?;
            Ok(obj(app, Kind::Slide(pk, id)))
        }
        "add" => {
            let index = arg_i64(args, 0, "Index")?;
            let layout = arg_i64(args, 1, "Layout")? as i32;
            add_slide(app, pk, index, LayoutChoice::Type(layout))
        }
        "addslide" => {
            let index = arg_i64(args, 0, "Index")?;
            match kind_of(args.require(1, "pCustomLayout")?) {
                Some(Kind::CustomLayout(p, i)) if p == pk => add_slide(app, pk, index, LayoutChoice::Index(i)),
                _ => Err(RuntimeError::type_mismatch()),
            }
        }
        "range" => {
            let ids = slide_ids(app, pk, args.get(0, "Index"))?;
            Ok(obj(app, Kind::SlideRange(pk, ids)))
        }
        "findbyslideid" => {
            let id = arg_i64(args, 0, "SlideID")?;
            let found = with(app, |st| Ok(st.pres(pk)?.slides().iter().any(|s| i64::from(s.id) == id)))?;
            if !found {
                return Err(app_err(format!("no slide with SlideID {id}")));
            }
            Ok(obj(app, Kind::Slide(pk, id as u32)))
        }
        "parent" => Ok(obj(app, Kind::Presentation(pk))),
        _ => Err(RuntimeError::no_member(n)),
    }
}

// ---- Slide / SlideRange ----

fn slide(app: &Shared, pk: u32, sid: u32, n: &str, args: &Args) -> Result<Value> {
    match n {
        "name" => with(app, |st| Ok(text(st.slide(pk, sid)?.name.as_str()))),
        "slideid" => Ok(long(sid)),
        "slideindex" => with(app, |st| Ok(long(st.slide_index(pk, sid)? + 1))),
        "slidenumber" => with(app, |st| {
            let i = st.slide_index(pk, sid)? as i32;
            Ok(long(i + st.pres(pk)?.first_slide_number()))
        }),
        "layout" => with(app, |st| Ok(long(st.slide(pk, sid)?.layout))),
        "customlayout" => {
            let i = with(app, |st| {
                let name = &st.slide(pk, sid)?.layout_name;
                st.pres(pk)?.layouts().iter().position(|l| &l.name == name).ok_or_else(|| app_err("no layout"))
            })?;
            Ok(obj(app, Kind::CustomLayout(pk, i)))
        }
        "shapes" => {
            with(app, |st| st.slide(pk, sid).map(|_| ()))?;
            let c = obj(app, Kind::Shapes(pk, sid));
            if args.is_empty() { Ok(c) } else { call_value(&c, "Item", args) }
        }
        "delete" => with_mut(app, |st| {
            let i = st.slide_index(pk, sid)?;
            let p = st.pres_mut(pk)?;
            p.slides_mut().remove(i);
            p.set_saved(false);
            Ok(Value::Empty)
        }),
        "duplicate" => {
            let id = with_mut(app, |st| {
                let i = st.slide_index(pk, sid)?;
                let p = st.pres_mut(pk)?;
                let id = next_slide_id(p.slides());
                let name = slide_name(p.slides(), id);
                let mut copy = p.slides()[i].clone();
                copy.id = id;
                copy.name = name;
                p.slides_mut().insert(i + 1, copy);
                p.set_saved(false);
                Ok(id)
            })?;
            Ok(obj(app, Kind::SlideRange(pk, vec![id])))
        }
        "moveto" => {
            let to = arg_i64(args, 0, "toPos")?;
            with_mut(app, |st| {
                let i = st.slide_index(pk, sid)?;
                let p = st.pres_mut(pk)?;
                let to = index_in(to, p.slides().len())?;
                let s = p.slides_mut().remove(i);
                p.slides_mut().insert(to, s);
                Ok(Value::Empty)
            })
        }
        "select" => Ok(Value::Empty),
        "parent" => Ok(obj(app, Kind::Presentation(pk))),
        "application" => Ok(obj(app, Kind::Application)),
        _ => Err(RuntimeError::no_member(n)),
    }
}

fn slide_assign(app: &Shared, pk: u32, sid: u32, n: &str, value: Value) -> Result<()> {
    match n {
        "name" => {
            let s = data(&value)?.to_text()?;
            with_mut(app, |st| {
                st.slide_mut(pk, sid)?.name = s;
                Ok(())
            })
        }
        // ORACLE: changing the layout of a slide re-flows its placeholders in the host.
        "layout" => {
            let l = data(&value)?.to_i64()? as i32;
            with_mut(app, |st| {
                st.slide_mut(pk, sid)?.layout = l;
                Ok(())
            })
        }
        _ => Err(read_only(n)),
    }
}

fn slide_range(app: &Shared, pk: u32, ids: &[u32], n: &str, args: &Args) -> Result<Value> {
    match n {
        "count" => Ok(long(ids.len())),
        "item" | "" => {
            let v = data(args.require(0, "Index")?)?;
            let id = match v {
                Value::String(s) => with(app, |st| {
                    ids.iter()
                        .copied()
                        .find(|id| st.slide(pk, *id).is_ok_and(|x| x.name.eq_ignore_ascii_case(&s)))
                        .ok_or_else(|| app_err(format!("no slide named {s}")))
                })?,
                other => ids[index_in(other.to_i64()?, ids.len())?],
            };
            Ok(obj(app, Kind::Slide(pk, id)))
        }
        "delete" => {
            for id in ids {
                slide(app, pk, *id, "delete", args)?;
            }
            Ok(Value::Empty)
        }
        _ => {
            let first = *ids.first().ok_or_else(|| app_err("the range is empty"))?;
            slide(app, pk, first, n, args)
        }
    }
}

// ---- PageSetup / SlideShowSettings ----

fn page_setup(app: &Shared, pk: u32, n: &str) -> Result<Value> {
    with(app, |st| {
        let p = st.pres(pk)?;
        let (cx, cy) = p.slide_size();
        Ok(match n {
            "slidewidth" => single(emu_to_pt(cx)),
            "slideheight" => single(emu_to_pt(cy)),
            "firstslidenumber" => long(p.first_slide_number()),
            "slideorientation" => long(if cx >= cy { 1 } else { 2 }),
            _ => return Err(RuntimeError::no_member(n)),
        })
    })
}

fn page_setup_assign(app: &Shared, pk: u32, n: &str, value: Value) -> Result<()> {
    let x = assigned_f64(&value)?;
    with_mut(app, |st| {
        let p = st.pres_mut(pk)?;
        let (cx, cy) = p.slide_size();
        match n {
            "slidewidth" => p.set_slide_size(pt_to_emu(x), cy),
            "slideheight" => p.set_slide_size(cx, pt_to_emu(x)),
            "firstslidenumber" => p.set_first_slide_number(x as i32),
            _ => return Err(read_only(n)),
        }
        Ok(())
    })
}

fn slide_show(app: &Shared, pk: u32, n: &str) -> Result<Value> {
    with(app, |st| {
        let s = st.pres(pk)?.slide_show();
        Ok(match n {
            "startingslide" => long(s.starting_slide),
            "endingslide" => long(s.ending_slide),
            "showtype" => long(s.show_type),
            "loopuntilstopped" => tri(s.loop_until_stopped),
            "advancemode" => long(s.advance_mode),
            "rangetype" => long(s.range_type),
            "showwithnarration" => tri(s.with_narration),
            "showwithanimation" => tri(s.with_animation),
            "run" => return Err(app_err("a slide show cannot run from a macro here")),
            _ => return Err(RuntimeError::no_member(n)),
        })
    })
}

fn slide_show_assign(app: &Shared, pk: u32, n: &str, value: Value) -> Result<()> {
    let v = data(&value)?;
    with_mut(app, |st| {
        let s = st.pres_mut(pk)?.slide_show_mut();
        match n {
            "startingslide" => s.starting_slide = v.to_i32()?,
            "endingslide" => s.ending_slide = v.to_i32()?,
            "showtype" => s.show_type = v.to_i32()?,
            "loopuntilstopped" => s.loop_until_stopped = to_tri(&v)?,
            "advancemode" => s.advance_mode = v.to_i32()?,
            "rangetype" => s.range_type = v.to_i32()?,
            "showwithnarration" => s.with_narration = to_tri(&v)?,
            "showwithanimation" => s.with_animation = to_tri(&v)?,
            _ => return Err(read_only(n)),
        }
        Ok(())
    })
}
