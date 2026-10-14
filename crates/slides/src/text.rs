//! `TextFrame`, `TextRange`, `Font`, `ParagraphFormat` and the font's `ColorFormat`.
//!
//! A text range is a span of characters of one text body (a shape's or a table cell's).
//! `TextFrame.TextRange` is the live whole-body range: it always covers all the text, so
//! `InsertAfter` and `Text =` on it see the new text. Ranges taken from it (`Paragraphs(2)`,
//! `Characters(1, 3)`, `Replace`'s result) are fixed spans. Paragraphs end with `vbCr` and
//! include it.

use macroost_lang::{Args, Result, RuntimeError, Value};

use crate::model::{CharFormat, FrameProps, TextBody, emu_to_pt, pt_to_emu};
use crate::objects::{Kind, PresentationObj, obj, read_only};
use crate::shapes::table_of;
use crate::state::{
    ShapeRef, Shared, State, app_err, data, long, opt_bool, opt_i64, opt_text, single, text, to_tri, tri, with,
    with_mut,
};

/// HRESULT E_INVALIDARG (0x80070057) as the error number macros see.
const E_INVALIDARG: i32 = -2_147_024_809;

/// Whose text: a shape's, or a table cell's (shape, row, column; 0-based).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) enum Target {
    Shape(ShapeRef),
    Cell(ShapeRef, usize, usize),
}

/// A span of a text body. `whole` ranges follow the text; others are fixed `[start, start +
/// len)`. `count` is what `.Count` reports (the number of paragraphs, runs, characters or
/// words the range was made of).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct RangeRef {
    pub target: Target,
    pub start: usize,
    pub len: usize,
    pub whole: bool,
    pub count: usize,
}

impl RangeRef {
    fn whole(target: Target) -> Self {
        Self { target, start: 0, len: 0, whole: true, count: 1 }
    }
    fn span(target: &Target, start: usize, len: usize, count: usize) -> Self {
        Self { target: target.clone(), start, len, whole: false, count }
    }
}

fn body<'a>(st: &'a State, t: &Target) -> Result<&'a TextBody> {
    match t {
        Target::Shape(r) => st.shape(r)?.text.as_ref().ok_or_else(|| app_err("the shape has no text")),
        Target::Cell(r, row, col) => {
            table_of(st, r)?.cells.get(*row).and_then(|c| c.get(*col)).ok_or_else(|| app_err("no such cell"))
        }
    }
}

/// Edits a text body; a shape set to resize with its text is refitted afterwards.
fn edit<T>(app: &Shared, t: &Target, f: impl FnOnce(&mut TextBody) -> Result<T>) -> Result<T> {
    with_mut(app, |st| {
        let out = match t {
            Target::Shape(r) => {
                let s = st.shape_mut(r)?;
                let out = f(s.text.as_mut().ok_or_else(|| app_err("the shape has no text"))?)?;
                s.fit_text();
                out
            }
            Target::Cell(r, row, col) => {
                let s = st.shape_mut(r)?;
                let t = s.table.as_mut().ok_or_else(|| app_err("the shape has no table"))?;
                let b = t.cells.get_mut(*row).and_then(|c| c.get_mut(*col)).ok_or_else(|| app_err("no such cell"))?;
                f(b)?
            }
        };
        let pk = match t {
            Target::Shape(r) | Target::Cell(r, ..) => r.pk,
        };
        st.pres_mut(pk)?.set_saved(false);
        Ok(out)
    })
}

/// The current `[start, end)` of a range, clipped to the text.
fn bounds(b: &TextBody, r: &RangeRef) -> (usize, usize) {
    if r.whole {
        return (0, b.len());
    }
    let s = r.start.min(b.len());
    (s, (r.start + r.len).min(b.len()))
}

fn frame_props(st: &State, t: &Target) -> Result<FrameProps> {
    match t {
        Target::Shape(r) => Ok(st.shape(r)?.frame.clone()),
        // ORACLE: table cells report the default insets.
        Target::Cell(..) => Ok(FrameProps::new(0, 1)),
    }
}

pub(crate) fn call(o: &PresentationObj, n: &str, name: &str, args: &Args) -> Result<Value> {
    let app = &o.app;
    match &o.kind {
        Kind::TextFrame(t) => text_frame(app, t, n, name),
        Kind::TextRange(r) => text_range(app, r, n, name, args),
        Kind::Font(r) => font(app, r, n, name),
        Kind::FontColor(r) => match n {
            "rgb" | "" => with(app, |st| {
                let b = body(st, &r.target)?;
                let (s, _) = bounds(b, r);
                Ok(long(b.format_at(s).color))
            }),
            "type" => Ok(long(1)),
            _ => Err(RuntimeError::no_member(name)),
        },
        Kind::ParagraphFormat(r) => match n {
            "alignment" => with(app, |st| {
                let b = body(st, &r.target)?;
                let aligns: Vec<i32> = paragraphs_of(b, r).iter().map(|(s, _)| b.format_at(*s).align).collect();
                Ok(long(match aligns.first() {
                    None => b.base.align,
                    Some(a) if aligns.iter().all(|x| x == a) => *a,
                    Some(_) => -2,
                }))
            }),
            _ => Err(RuntimeError::no_member(name)),
        },
        _ => Err(RuntimeError::no_member(name)),
    }
}

pub(crate) fn assign(o: &PresentationObj, n: &str, name: &str, _args: &Args, value: Value) -> Result<()> {
    let app = &o.app;
    let v = data(&value)?;
    match &o.kind {
        Kind::TextFrame(t) => {
            let Target::Shape(r) = t else {
                return Ok(());
            };
            with_mut(app, |st| {
                let s = st.shape_mut(r)?;
                match n {
                    "marginleft" => s.frame.margins[0] = pt_to_emu(v.to_f64()?),
                    "marginright" => s.frame.margins[1] = pt_to_emu(v.to_f64()?),
                    "margintop" => s.frame.margins[2] = pt_to_emu(v.to_f64()?),
                    "marginbottom" => s.frame.margins[3] = pt_to_emu(v.to_f64()?),
                    "wordwrap" => s.frame.word_wrap = to_tri(&v)?,
                    "autosize" => {
                        s.frame.auto_size = v.to_i32()?;
                        s.fit_text();
                    }
                    "verticalanchor" => s.frame.anchor = v.to_i32()?,
                    "orientation" => s.frame.orientation = v.to_i32()?,
                    _ => return Err(read_only(name)),
                }
                Ok(())
            })
        }
        Kind::TextRange(r) => match n {
            "text" | "" => {
                let s = v.to_text()?;
                edit(app, &r.target, |b| {
                    let (start, end) = bounds(b, r);
                    b.replace(start, end - start, &s);
                    Ok(())
                })
            }
            _ => Err(read_only(name)),
        },
        Kind::Font(r) => set_font(app, r, n, name, &v),
        Kind::FontColor(r) if n == "rgb" || n.is_empty() => {
            let c = v.to_i32()?;
            format_chars(app, r, |f| f.color = c)
        }
        Kind::ParagraphFormat(r) if n == "alignment" => {
            let a = v.to_i32()?;
            edit(app, &r.target, |b| {
                let (start, end) = bounds(b, r);
                let paras = paragraphs_of(b, r);
                for (s, e) in &paras {
                    for f in &mut b.fmt[*s..*e] {
                        f.align = a;
                    }
                }
                if start == 0 && end == b.len() {
                    b.base.align = a;
                }
                Ok(())
            })
        }
        _ => Err(read_only(name)),
    }
}

pub(crate) fn items(o: &PresentationObj) -> Result<Vec<Value>> {
    Err(RuntimeError::no_member(&format!("_NewEnum ({})", macroost_lang::Object::type_name(o))))
}

fn text_frame(app: &Shared, t: &Target, n: &str, name: &str) -> Result<Value> {
    match n {
        "textrange" => {
            // A frame without text (a table's) refuses with E_INVALIDARG.
            with(app, |st| body(st, t).map(|_| ()).map_err(|_| RuntimeError::new(E_INVALIDARG, "Invalid argument")))?;
            Ok(obj(app, Kind::TextRange(RangeRef::whole(t.clone()))))
        }
        "hastext" => with(app, |st| Ok(tri(body(st, t).is_ok_and(|b| !b.is_empty())))),
        "deletetext" => edit(app, t, |b| {
            let n = b.len();
            b.replace(0, n, "");
            Ok(Value::Empty)
        }),
        "parent" => Ok(match t {
            Target::Shape(r) => obj(app, Kind::Shape(r.clone())),
            Target::Cell(r, row, col) => obj(app, Kind::CellShape(r.clone(), *row, *col)),
        }),
        _ => with(app, |st| {
            let f = frame_props(st, t)?;
            Ok(match n {
                "marginleft" => single(emu_to_pt(f.margins[0])),
                "marginright" => single(emu_to_pt(f.margins[1])),
                "margintop" => single(emu_to_pt(f.margins[2])),
                "marginbottom" => single(emu_to_pt(f.margins[3])),
                "wordwrap" => tri(f.word_wrap),
                "autosize" => long(f.auto_size),
                "orientation" => long(f.orientation),
                "verticalanchor" => long(f.anchor),
                _ => return Err(RuntimeError::no_member(name)),
            })
        }),
    }
}

// ---- units: paragraphs, runs, words, characters ----

/// Paragraphs overlapping the range (whole paragraphs).
fn paragraphs_of(b: &TextBody, r: &RangeRef) -> Vec<(usize, usize)> {
    let (s, e) = bounds(b, r);
    b.paragraphs().into_iter().filter(|(ps, pe)| *pe > s && (*ps < e || (s == e && *ps == s))).collect()
}

/// Runs (same formatting, not crossing a paragraph end) inside the range.
fn runs_of(b: &TextBody, r: &RangeRef) -> Vec<(usize, usize)> {
    let (s, e) = bounds(b, r);
    let mut out = Vec::new();
    let mut start = s;
    for i in s..e {
        let next_differs = i + 1 < e && (b.chars[i] == '\r' || b.fmt[i + 1] != b.fmt[i]);
        if next_differs || i + 1 == e {
            out.push((start, i + 1));
            start = i + 1;
        }
    }
    out
}

/// Words with their trailing spaces.
fn words_of(b: &TextBody, r: &RangeRef) -> Vec<(usize, usize)> {
    let (s, e) = bounds(b, r);
    let mut out = Vec::new();
    let mut i = s;
    while i < e {
        let start = i;
        if b.chars[i] == '\r' {
            i += 1;
        } else {
            while i < e && !b.chars[i].is_whitespace() {
                i += 1;
            }
            while i < e && b.chars[i] == ' ' {
                i += 1;
            }
            // The paragraph break belongs to the word before it.
            if i < e && b.chars[i] == '\r' {
                i += 1;
            }
            if i == start {
                i += 1;
            }
        }
        out.push((start, i));
    }
    out
}

/// A range made of `units[first - 1 .. first - 1 + n]` (all of them without arguments).
fn unit_range(r: &RangeRef, units: &[(usize, usize)], args: &Args, at: usize) -> Result<RangeRef> {
    let first = opt_i64(args, 0, "Start")?;
    let length = opt_i64(args, 1, "Length")?;
    let (from, take) = match first {
        None => (0, units.len()),
        Some(f) => {
            if f < 1 || f as usize > units.len() {
                return Ok(RangeRef::span(&r.target, at, 0, 0));
            }
            let n = length.unwrap_or(1).max(0) as usize;
            ((f - 1) as usize, n.min(units.len() - (f - 1) as usize))
        }
    };
    let picked = &units[from..from + take];
    Ok(match (picked.first(), picked.last()) {
        (Some(a), Some(z)) => RangeRef::span(&r.target, a.0, z.1 - a.0, take),
        _ => RangeRef::span(&r.target, at, 0, 0),
    })
}

fn find_in(b: &TextBody, r: &RangeRef, what: &str, after: usize, match_case: bool, whole: bool) -> Option<usize> {
    let (s, e) = bounds(b, r);
    let needle: Vec<char> = what.chars().collect();
    if needle.is_empty() {
        return None;
    }
    let fold = |c: char| if match_case { c } else { c.to_lowercase().next().unwrap_or(c) };
    let from = s + after.min(e - s);
    (from..e.saturating_sub(needle.len() - 1)).find(|&i| {
        let hit = needle.iter().enumerate().all(|(k, c)| b.chars.get(i + k).is_some_and(|x| fold(*x) == fold(*c)));
        let edge = |p: Option<&char>| p.is_none_or(|c| !c.is_alphanumeric());
        hit && (!whole || (edge(i.checked_sub(1).and_then(|p| b.chars.get(p))) && edge(b.chars.get(i + needle.len()))))
    })
}

fn text_range(app: &Shared, r: &RangeRef, n: &str, name: &str, args: &Args) -> Result<Value> {
    let range = |rr: RangeRef| obj(app, Kind::TextRange(rr));
    match n {
        "text" | "" => with(app, |st| {
            let b = body(st, &r.target)?;
            let (s, e) = bounds(b, r);
            Ok(text(b.chars[s..e].iter().collect::<String>()))
        }),
        "length" => with(app, |st| {
            let (s, e) = bounds(body(st, &r.target)?, r);
            Ok(long(e - s))
        }),
        "start" => with(app, |st| Ok(long(bounds(body(st, &r.target)?, r).0 + 1))),
        "count" => Ok(long(r.count)),
        "paragraphs" | "lines" | "sentences" | "runs" | "words" | "characters" => {
            let rr = with(app, |st| {
                let b = body(st, &r.target)?;
                let (s, e) = bounds(b, r);
                let units = match n {
                    "runs" => runs_of(b, r),
                    "words" => words_of(b, r),
                    "characters" => (s..e).map(|i| (i, i + 1)).collect(),
                    _ => paragraphs_of(b, r),
                };
                unit_range(r, &units, args, s)
            })?;
            Ok(range(rr))
        }
        "font" => Ok(obj(app, Kind::Font(r.clone()))),
        "paragraphformat" => Ok(obj(app, Kind::ParagraphFormat(r.clone()))),
        "replace" | "find" => {
            let replace = n == "replace";
            let what = data(args.require(0, "FindWhat")?)?.to_text()?;
            let with_text = if replace { Some(data(args.require(1, "ReplaceWhat")?)?.to_text()?) } else { None };
            let k = if replace { 2 } else { 1 };
            let after = opt_i64(args, k, "After")?.unwrap_or(0).max(0) as usize;
            let match_case = opt_bool(args, k + 1, "MatchCase")?.unwrap_or(false);
            let whole_words = opt_bool(args, k + 2, "WholeWords")?.unwrap_or(false);
            let found = with(app, |st| Ok(find_in(body(st, &r.target)?, r, &what, after, match_case, whole_words)))?;
            let Some(at) = found else { return Ok(Value::Object(None)) };
            let len = what.chars().count();
            match with_text {
                None => Ok(range(RangeRef::span(&r.target, at, len, 1))),
                Some(rep) => {
                    let n = edit(app, &r.target, |b| Ok(b.replace(at, len, &rep)))?;
                    Ok(range(RangeRef::span(&r.target, at, n, 1)))
                }
            }
        }
        "insertafter" | "insertbefore" => {
            let s = opt_text(args, 0, "NewText")?.unwrap_or_default();
            let after = n == "insertafter";
            let (at, count) = edit(app, &r.target, |b| {
                let (start, end) = bounds(b, r);
                let at = if after { end } else { start };
                Ok((at, b.replace(at, 0, &s)))
            })?;
            Ok(range(RangeRef::span(&r.target, at, count, 1)))
        }
        "delete" => edit(app, &r.target, |b| {
            let (s, e) = bounds(b, r);
            b.replace(s, e - s, "");
            Ok(Value::Empty)
        }),
        "parent" => Ok(obj(app, Kind::TextFrame(r.target.clone()))),
        "select" => Ok(Value::Empty),
        _ => Err(RuntimeError::no_member(name)),
    }
}

// ---- Font ----

/// The formats of the range's characters (the insertion format when it is empty).
fn formats(b: &TextBody, r: &RangeRef) -> Vec<CharFormat> {
    let (s, e) = bounds(b, r);
    if s == e { vec![b.format_at(s)] } else { b.fmt[s..e].to_vec() }
}

fn tri_of(fs: &[CharFormat], f: impl Fn(&CharFormat) -> bool) -> Value {
    let on = fs.iter().filter(|x| f(x)).count();
    if on == fs.len() {
        tri(true)
    } else if on == 0 {
        tri(false)
    } else {
        long(-2)
    }
}

fn font(app: &Shared, r: &RangeRef, n: &str, name: &str) -> Result<Value> {
    if n == "color" {
        return Ok(obj(app, Kind::FontColor(r.clone())));
    }
    with(app, |st| {
        let fs = formats(body(st, &r.target)?, r);
        let first = fs.first().cloned().ok_or_else(|| app_err("no text"))?;
        Ok(match n {
            "bold" => tri_of(&fs, |f| f.bold),
            "italic" => tri_of(&fs, |f| f.italic),
            "underline" => tri_of(&fs, |f| f.underline),
            // ORACLE: the value a range of mixed sizes or fonts reports.
            "size" => single(first.size),
            "name" => text(first.font),
            _ => return Err(RuntimeError::no_member(name)),
        })
    })
}

/// Applies `f` to the range's characters; a range covering all the text also sets the format
/// new text gets.
fn format_chars(app: &Shared, r: &RangeRef, f: impl Fn(&mut CharFormat)) -> Result<()> {
    edit(app, &r.target, |b| {
        let (s, e) = bounds(b, r);
        for x in &mut b.fmt[s..e] {
            f(x);
        }
        if s == 0 && e == b.len() {
            f(&mut b.base);
        }
        Ok(())
    })
}

fn set_font(app: &Shared, r: &RangeRef, n: &str, name: &str, v: &Value) -> Result<()> {
    match n {
        "bold" => {
            let b = to_tri(v)?;
            format_chars(app, r, |f| f.bold = b)
        }
        "italic" => {
            let b = to_tri(v)?;
            format_chars(app, r, |f| f.italic = b)
        }
        "underline" => {
            let b = to_tri(v)?;
            format_chars(app, r, |f| f.underline = b)
        }
        "size" => {
            let s = v.to_f64()?;
            if !(1.0..=4000.0).contains(&s) {
                return Err(app_err("font size must be between 1 and 4000"));
            }
            format_chars(app, r, |f| f.size = s as f32)
        }
        "name" => {
            let s: std::rc::Rc<str> = v.to_text()?.into();
            format_chars(app, r, |f| f.font = s.clone())
        }
        _ => Err(read_only(name)),
    }
}
