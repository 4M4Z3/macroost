//! The presentation data a macro sees: slides, shapes with geometry, text with character
//! formatting, and tables. Geometry is stored in EMU (914,400 per inch, 12,700 per point), the
//! unit of the file format, so values round the way the reference host rounds them (a column of
//! 200/3 points reads back as 66.66669).
//!
//! Defaults (fonts, sizes, colors, placeholder positions) are those of the presentation that
//! `macroost_project::pptm_with_macros` writes, as observed in the reference host; see
//! `conformance/expected/slides.json`.

use std::rc::Rc;

/// EMU per point.
pub const EMU_PER_PT: f64 = 12_700.0;

/// Points to EMU, rounded to the nearest unit.
pub fn pt_to_emu(pt: f64) -> i64 {
    (pt * EMU_PER_PT).round() as i64
}

/// EMU to points as the host reports them (Single precision).
pub fn emu_to_pt(emu: i64) -> f32 {
    (emu as f64 / EMU_PER_PT) as f32
}

/// `RGB(r, g, b)` as the language packs it (red in the low byte).
pub const fn rgb(r: u8, g: u8, b: u8) -> i32 {
    (r as i32) | ((g as i32) << 8) | ((b as i32) << 16)
}

/// Theme colors of the generated presentation, packed as `RGB` longs.
pub const BLACK: i32 = 0;
pub const WHITE: i32 = rgb(255, 255, 255);
/// accent1 (4472C4): default fill of an AutoShape.
pub const ACCENT1: i32 = rgb(0x44, 0x72, 0xC4);
/// Default outline of an AutoShape (accent1, darkened by the theme's line style).
pub const ACCENT1_LINE: i32 = 5_319_703;

/// Character formatting. `align` is the paragraph alignment of the character's paragraph
/// (`ppAlign*`); a paragraph's alignment is that of its first character.
#[derive(Debug, Clone, PartialEq)]
pub struct CharFormat {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// Points.
    pub size: f32,
    pub color: i32,
    pub font: Rc<str>,
    pub align: i32,
}

impl CharFormat {
    pub fn new(size: f32, font: &str, color: i32, align: i32) -> Self {
        Self { bold: false, italic: false, underline: false, size, color, font: font.into(), align }
    }
}

/// Text of a shape or table cell: characters with one format each. Paragraphs are separated
/// by `'\r'`; `'\n'` is a line break inside a paragraph. `base` formats text typed into an
/// empty body (the end-of-paragraph properties).
#[derive(Debug, Clone, PartialEq)]
pub struct TextBody {
    pub chars: Vec<char>,
    pub fmt: Vec<CharFormat>,
    pub base: CharFormat,
}

impl TextBody {
    pub fn new(base: CharFormat) -> Self {
        Self { chars: Vec::new(), fmt: Vec::new(), base }
    }

    pub fn text(&self) -> String {
        self.chars.iter().collect()
    }

    pub fn len(&self) -> usize {
        self.chars.len()
    }

    pub fn is_empty(&self) -> bool {
        self.chars.is_empty()
    }

    /// The format new text inserted at `pos` takes.
    pub fn format_at(&self, pos: usize) -> CharFormat {
        self.fmt.get(pos).or_else(|| pos.checked_sub(1).and_then(|p| self.fmt.get(p))).unwrap_or(&self.base).clone()
    }

    /// Replaces `chars[start..start + len]` with `text` (line endings normalized: CR LF
    /// becomes one paragraph break). Returns the number of characters inserted.
    pub fn replace(&mut self, start: usize, len: usize, text: &str) -> usize {
        let start = start.min(self.chars.len());
        let end = (start + len).min(self.chars.len());
        // Replaced text keeps the format of its first character; inserted text continues the
        // character before it.
        let fmt = if end == start && start > 0 { self.format_at(start - 1) } else { self.format_at(start) };
        let new: Vec<char> = text.replace("\r\n", "\r").chars().collect();
        let n = new.len();
        self.chars.splice(start..end, new);
        self.fmt.splice(start..end, std::iter::repeat_n(fmt, n));
        n
    }

    /// Paragraph spans `(start, end)`, each including its terminating `'\r'`. Empty text has
    /// no paragraphs.
    pub fn paragraphs(&self) -> Vec<(usize, usize)> {
        let mut out = Vec::new();
        if self.chars.is_empty() {
            return out;
        }
        let mut start = 0;
        for (i, c) in self.chars.iter().enumerate() {
            if *c == '\r' {
                out.push((start, i + 1));
                start = i + 1;
            }
        }
        out.push((start, self.chars.len()));
        out
    }

    /// Visual lines (paragraphs and line breaks), each with its largest font size.
    pub fn line_sizes(&self) -> Vec<f32> {
        if self.chars.is_empty() {
            return vec![self.base.size];
        }
        let mut out = Vec::new();
        let mut cur: Option<f32> = None;
        for (c, f) in self.chars.iter().zip(&self.fmt) {
            if *c == '\r' || *c == '\n' {
                out.push(cur.unwrap_or(f.size).max(f.size));
                cur = None;
            } else {
                cur = Some(cur.map_or(f.size, |s| s.max(f.size)));
            }
        }
        let last = self.fmt.last().map_or(self.base.size, |f| f.size);
        out.push(cur.unwrap_or(last));
        out
    }
}

/// EMU height of one line of text in the theme's body font at `size` points. Measured on the
/// reference host for sizes 8..96, one to four lines, mixed sizes (`conformance/expected/slides.json`
/// T10, T70, T73, T84, T85; `slides2.json` U23, U35).
pub fn line_height_emu(size: f32) -> i64 {
    (f64::from(size) * 15_388.827).round() as i64
}

/// What the top + bottom insets add to an auto-fitted height: slightly more than their sum
/// (7.2 points of insets add 92,333 EMU, 10 points add 128,240; `slides2.json` U24).
pub fn inset_height_emu(insets: i64) -> i64 {
    (insets as f64 * 1.009_765_625).round() as i64
}

/// Height (EMU) of the smallest table row the host allows with 18-point text.
pub const MIN_ROW_HEIGHT_EMU: i64 = 365_760;

/// Text frame properties.
#[derive(Debug, Clone, PartialEq)]
pub struct FrameProps {
    /// Left, right, top, bottom insets in EMU.
    pub margins: [i64; 4],
    pub word_wrap: bool,
    /// `ppAutoSize*`: 0 none, 1 shape to fit text.
    pub auto_size: i32,
    /// `msoAnchor*`.
    pub anchor: i32,
    /// `msoTextOrientation*`.
    pub orientation: i32,
}

impl FrameProps {
    pub fn new(auto_size: i32, anchor: i32) -> Self {
        Self { margins: [91_440, 91_440, 45_720, 45_720], word_wrap: true, auto_size, anchor, orientation: 1 }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FillProps {
    pub visible: bool,
    pub color: i32,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LineProps {
    pub visible: bool,
    pub color: i32,
    /// Points.
    pub weight: f32,
}

/// A table: column widths and row heights in EMU, and one text body per cell.
#[derive(Debug, Clone, PartialEq)]
pub struct TableData {
    pub cols: Vec<i64>,
    pub rows: Vec<i64>,
    pub cells: Vec<Vec<TextBody>>,
}

impl TableData {
    /// Cell text format: the default table style sets the header row bold and white.
    pub fn cell_format(row: usize) -> CharFormat {
        let mut f = CharFormat::new(18.0, "Calibri", if row == 0 { WHITE } else { BLACK }, 1);
        f.bold = row == 0;
        f
    }

    pub fn new(rows: usize, cols: usize, width: i64, height: i64) -> Self {
        let cw = (width as f64 / cols as f64).round() as i64;
        let rh = ((height as f64 / rows as f64).round() as i64).max(MIN_ROW_HEIGHT_EMU);
        let cells = (0..rows).map(|r| (0..cols).map(|_| TextBody::new(Self::cell_format(r))).collect()).collect();
        Self { cols: vec![cw; cols], rows: vec![rh; rows], cells }
    }

    pub fn width(&self) -> i64 {
        self.cols.iter().sum()
    }

    pub fn height(&self) -> i64 {
        self.rows.iter().sum()
    }
}

/// What a shape is, with its `msoShapeType`.
#[derive(Debug, Clone, PartialEq)]
pub enum ShapeKind {
    /// `msoTextBox` (17).
    TextBox,
    /// `msoAutoShape` (1) with its `msoAutoShapeType`.
    AutoShape(i32),
    /// `msoPlaceholder` (14) with its `ppPlaceholder*` type.
    Placeholder(i32),
    /// `msoTable` (19).
    Table,
}

impl ShapeKind {
    pub fn shape_type(&self) -> i32 {
        match self {
            ShapeKind::TextBox => 17,
            ShapeKind::AutoShape(_) => 1,
            ShapeKind::Placeholder(_) => 14,
            ShapeKind::Table => 19,
        }
    }

    /// The name stem the host gives new shapes ("Rectangle 3" = stem + (id - 1)).
    pub fn name_stem(&self) -> &'static str {
        match self {
            ShapeKind::TextBox => "TextBox",
            ShapeKind::Table => "Table",
            ShapeKind::Placeholder(t) => match t {
                1 | 3 | 5 => "Title",
                4 => "Subtitle",
                2 | 6 => "Text Placeholder",
                _ => "Content Placeholder",
            },
            ShapeKind::AutoShape(t) => autoshape_stem(*t),
        }
    }
}

/// Name stems of AutoShapes (observed for the types in oracle case T77; others are marked
/// ORACLE).
pub fn autoshape_stem(t: i32) -> &'static str {
    match t {
        1 => "Rectangle",
        2 => "Parallelogram",
        3 => "Trapezoid",
        4 => "Diamond",
        5 => "Rounded Rectangle",
        6 => "Octagon",
        7 => "Isosceles Triangle",
        8 => "Right Triangle",
        9 => "Oval",
        10 => "Hexagon",
        11 => "Cross",
        12 => "Regular Pentagon",
        13 => "Can",
        14 => "Cube",
        21 => "Heart",
        33 => "Right Arrow",
        34 => "Left Arrow",
        35 => "Up Arrow",
        36 => "Down Arrow",
        92 => "5-Point Star",
        // ORACLE: names of the remaining AutoShape types.
        _ => "Shape",
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    pub id: u32,
    pub name: String,
    pub kind: ShapeKind,
    /// Left, top, width, height in EMU.
    pub x: i64,
    pub y: i64,
    pub cx: i64,
    pub cy: i64,
    pub rotation: f32,
    pub visible: bool,
    pub flip_h: bool,
    pub flip_v: bool,
    /// `None` for shapes without a text frame (tables).
    pub text: Option<TextBody>,
    pub frame: FrameProps,
    pub fill: FillProps,
    pub line: LineProps,
    pub table: Option<TableData>,
}

impl Shape {
    fn base(id: u32, kind: ShapeKind, geo: (i64, i64, i64, i64)) -> Self {
        let name = format!("{} {}", kind.name_stem(), id.saturating_sub(1));
        Self {
            id,
            name,
            kind,
            x: geo.0,
            y: geo.1,
            cx: geo.2,
            cy: geo.3,
            rotation: 0.0,
            visible: true,
            flip_h: false,
            flip_v: false,
            text: None,
            frame: FrameProps::new(0, 1),
            fill: FillProps { visible: false, color: WHITE },
            line: LineProps { visible: false, color: BLACK, weight: 0.75 },
            table: None,
        }
    }

    pub fn textbox(id: u32, geo: (i64, i64, i64, i64)) -> Self {
        let mut s = Self::base(id, ShapeKind::TextBox, geo);
        s.text = Some(TextBody::new(CharFormat::new(18.0, "Calibri", BLACK, 1)));
        s.frame = FrameProps::new(1, 1);
        s
    }

    pub fn autoshape(id: u32, shape_type: i32, geo: (i64, i64, i64, i64)) -> Self {
        let mut s = Self::base(id, ShapeKind::AutoShape(shape_type), geo);
        s.text = Some(TextBody::new(CharFormat::new(18.0, "Calibri", WHITE, 2)));
        s.frame = FrameProps::new(0, 3);
        s.fill = FillProps { visible: true, color: ACCENT1 };
        s.line = LineProps { visible: true, color: ACCENT1_LINE, weight: 1.0 };
        s
    }

    pub fn placeholder(id: u32, spec: &PlaceholderSpec) -> Self {
        let mut s = Self::base(id, ShapeKind::Placeholder(spec.ph_type), spec.geometry);
        s.text = Some(TextBody::new(CharFormat::new(spec.font_size, &spec.font, BLACK, 1)));
        s
    }

    pub fn table(id: u32, rows: usize, cols: usize, geo: (i64, i64, i64, i64)) -> Self {
        let mut s = Self::base(id, ShapeKind::Table, geo);
        let t = TableData::new(rows, cols, geo.2, geo.3);
        s.cx = t.width();
        s.cy = t.height();
        s.table = Some(t);
        s
    }

    /// Applies "resize shape to fit text": the height follows the text, the top stays.
    pub fn fit_text(&mut self) {
        if self.frame.auto_size != 1 {
            return;
        }
        if let Some(t) = &self.text {
            let lines: i64 = t.line_sizes().into_iter().map(line_height_emu).sum();
            self.cy = inset_height_emu(self.frame.margins[2] + self.frame.margins[3]) + lines;
        }
    }
}

/// A placeholder a layout puts on new slides.
#[derive(Debug, Clone, PartialEq)]
pub struct PlaceholderSpec {
    /// `ppPlaceholder*`.
    pub ph_type: i32,
    pub geometry: (i64, i64, i64, i64),
    pub font_size: f32,
    pub font: String,
}

/// A slide layout of the master.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomLayout {
    pub name: String,
    /// The `ppLayout*` value slides built on it report.
    pub layout_type: i32,
    pub placeholders: Vec<PlaceholderSpec>,
}

const TITLE_GEO: (i64, i64, i64, i64) = (838_200, 365_125, 10_515_600, 1_325_563);
const BODY_GEO: (i64, i64, i64, i64) = (838_200, 1_825_625, 10_515_600, 4_351_338);

fn ph(ph_type: i32, geometry: (i64, i64, i64, i64)) -> PlaceholderSpec {
    let (font_size, font) = match ph_type {
        1 | 3 | 5 => (44.0, "Calibri Light"),
        _ => (28.0, "Calibri"),
    };
    PlaceholderSpec { ph_type, geometry, font_size, font: font.to_owned() }
}

/// The layout the host uses (adding it to the master when missing) for `ppLayout*` value
/// `layout`. ORACLE: geometry of layouts the generated presentation does not define.
pub fn builtin_layout(layout: i32) -> CustomLayout {
    let (name, phs): (&str, Vec<PlaceholderSpec>) = match layout {
        1 => (
            "Title Slide",
            vec![
                ph(3, (1_524_000, 1_122_363, 9_144_000, 2_387_600)),
                ph(4, (1_524_000, 3_602_038, 9_144_000, 1_655_762)),
            ],
        ),
        2 => ("Title and Text", vec![ph(1, TITLE_GEO), ph(2, BODY_GEO)]),
        3 => (
            "Title and 2-Column Text",
            vec![
                ph(1, TITLE_GEO),
                ph(2, (838_200, 1_825_625, 5_181_600, 4_351_338)),
                ph(2, (6_172_200, 1_825_625, 5_181_600, 4_351_338)),
            ],
        ),
        11 => ("Title Only", vec![ph(1, TITLE_GEO)]),
        12 => ("Blank", vec![]),
        16 => ("Title and Content", vec![ph(1, TITLE_GEO), ph(7, BODY_GEO)]),
        29 => (
            "Two Content",
            vec![
                ph(1, TITLE_GEO),
                ph(7, (838_200, 1_825_625, 5_181_600, 4_351_338)),
                ph(7, (6_172_200, 1_825_625, 5_181_600, 4_351_338)),
            ],
        ),
        33 => ("Section Header", vec![ph(1, TITLE_GEO), ph(2, BODY_GEO)]),
        34 => ("Comparison", vec![ph(1, TITLE_GEO)]),
        35 => ("Content with Caption", vec![ph(1, TITLE_GEO)]),
        36 => ("Picture with Caption", vec![ph(1, TITLE_GEO)]),
        _ => ("Custom Layout", vec![ph(1, TITLE_GEO), ph(2, BODY_GEO)]),
    };
    CustomLayout { name: name.to_owned(), layout_type: layout, placeholders: phs }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Slide {
    /// `SlideID`: unique in the presentation, 256 and up.
    pub id: u32,
    pub name: String,
    /// `ppLayout*`.
    pub layout: i32,
    /// Name of its custom layout.
    pub layout_name: String,
    /// Shapes in z-order (back to front).
    pub shapes: Vec<Shape>,
    /// Id the next new shape gets; ids are not reused on the slide.
    pub next_shape_id: u32,
}

impl Slide {
    /// A slide built on `layout`, with its placeholders.
    pub fn new(id: u32, name: String, layout: &CustomLayout) -> Self {
        let mut s = Self {
            id,
            name,
            layout: layout.layout_type,
            layout_name: layout.name.clone(),
            shapes: Vec::new(),
            next_shape_id: 2,
        };
        for spec in &layout.placeholders {
            let id = s.take_id();
            s.shapes.push(Shape::placeholder(id, spec));
        }
        s
    }

    pub fn take_id(&mut self) -> u32 {
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        id
    }

    pub fn shape(&self, id: u32) -> Option<&Shape> {
        self.shapes.iter().find(|s| s.id == id)
    }

    pub fn shape_mut(&mut self, id: u32) -> Option<&mut Shape> {
        self.shapes.iter_mut().find(|s| s.id == id)
    }
}

/// Slide-show settings (`Presentation.SlideShowSettings`).
#[derive(Debug, Clone, PartialEq)]
pub struct SlideShow {
    pub starting_slide: i32,
    pub ending_slide: i32,
    /// `ppShowType*`.
    pub show_type: i32,
    pub loop_until_stopped: bool,
    /// `ppSlideShowAdvance*`.
    pub advance_mode: i32,
    /// `ppShowAll` / `ppShowSlideRange` / `ppShowNamedSlideShow`.
    pub range_type: i32,
    pub with_narration: bool,
    pub with_animation: bool,
}

impl Default for SlideShow {
    fn default() -> Self {
        Self {
            starting_slide: 1,
            ending_slide: 1,
            show_type: 1,
            loop_until_stopped: false,
            advance_mode: 2,
            range_type: 1,
            with_narration: true,
            with_animation: true,
        }
    }
}

/// What a presentation application exposes to the object model. The embedding application
/// implements it over its own document; [`MemoryPresentation`] is a complete in-memory one.
pub trait Presentation {
    /// File name (`Presentation.Name`), e.g. `Deck.pptm`.
    fn name(&self) -> String;
    /// Slide size (width, height) in EMU.
    fn slide_size(&self) -> (i64, i64);
    fn set_slide_size(&mut self, cx: i64, cy: i64);
    fn first_slide_number(&self) -> i32;
    fn set_first_slide_number(&mut self, n: i32);
    /// Slides in order.
    fn slides(&self) -> &[Slide];
    fn slides_mut(&mut self) -> &mut Vec<Slide>;
    /// The master's layouts in order.
    fn layouts(&self) -> &[CustomLayout];
    fn layouts_mut(&mut self) -> &mut Vec<CustomLayout>;
    fn slide_show(&self) -> &SlideShow;
    fn slide_show_mut(&mut self) -> &mut SlideShow;
    /// `Saved`: no changes since the last save.
    fn saved(&self) -> bool;
    fn set_saved(&mut self, saved: bool);
}

/// An in-memory presentation.
#[derive(Debug, Clone)]
pub struct MemoryPresentation {
    pub name: String,
    pub size: (i64, i64),
    pub first_number: i32,
    pub slides: Vec<Slide>,
    pub layouts: Vec<CustomLayout>,
    pub show: SlideShow,
    pub saved: bool,
}

impl MemoryPresentation {
    /// The presentation `macroost_project::pptm_with_macros` writes: 960 x 540 points, layouts
    /// Title Slide / Title and Content / Title and Text / Title Only / Blank, and
    /// `slide_count` blank slides (SlideID 256, 257, ...).
    pub fn new(name: &str, slide_count: usize) -> Self {
        let layouts: Vec<CustomLayout> = [1, 16, 2, 11, 12].into_iter().map(builtin_layout).collect();
        let blank = builtin_layout(12);
        let slides = (0..slide_count as u32).map(|i| Slide::new(256 + i, format!("Slide{}", i + 1), &blank)).collect();
        Self {
            name: name.to_owned(),
            size: (12_192_000, 6_858_000),
            first_number: 1,
            slides,
            layouts,
            show: SlideShow::default(),
            saved: true,
        }
    }

    /// A new presentation from the host's default template (`Presentations.Add`): no slides,
    /// 960 x 540 points, eleven layouts.
    pub fn from_default_template(name: &str) -> Self {
        let mut layouts: Vec<CustomLayout> =
            [1, 16, 33, 29, 34, 11, 12, 35, 36].into_iter().map(builtin_layout).collect();
        layouts.push(CustomLayout { name: "Title and Vertical Text".into(), layout_type: 27, placeholders: vec![] });
        layouts.push(CustomLayout { name: "Vertical Title and Text".into(), layout_type: 28, placeholders: vec![] });
        Self { layouts, ..Self::new(name, 0) }
    }
}

impl Presentation for MemoryPresentation {
    fn name(&self) -> String {
        self.name.clone()
    }
    fn slide_size(&self) -> (i64, i64) {
        self.size
    }
    fn set_slide_size(&mut self, cx: i64, cy: i64) {
        self.size = (cx, cy);
    }
    fn first_slide_number(&self) -> i32 {
        self.first_number
    }
    fn set_first_slide_number(&mut self, n: i32) {
        self.first_number = n;
    }
    fn slides(&self) -> &[Slide] {
        &self.slides
    }
    fn slides_mut(&mut self) -> &mut Vec<Slide> {
        &mut self.slides
    }
    fn layouts(&self) -> &[CustomLayout] {
        &self.layouts
    }
    fn layouts_mut(&mut self) -> &mut Vec<CustomLayout> {
        &mut self.layouts
    }
    fn slide_show(&self) -> &SlideShow {
        &self.show
    }
    fn slide_show_mut(&mut self) -> &mut SlideShow {
        &mut self.show
    }
    fn saved(&self) -> bool {
        self.saved
    }
    fn set_saved(&mut self, saved: bool) {
        self.saved = saved;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_height_matches_the_measurements() {
        let mut s = Shape::textbox(2, (0, 0, 2_540_000, 635_000));
        let t = s.text.as_mut().map(|t| t.replace(0, 0, "Hello"));
        assert_eq!(t, Some(5));
        s.fit_text();
        assert_eq!(s.cy, 369_332);
        for (size, h) in [(8.0, 215_444), (32.0, 584_775), (72.0, 1_200_329), (96.0, 1_569_660)] {
            assert_eq!(inset_height_emu(91_440) + line_height_emu(size), h, "{size}");
        }
    }

    #[test]
    fn paragraphs_include_their_break() {
        let mut t = TextBody::new(CharFormat::new(18.0, "Calibri", 0, 1));
        t.replace(0, 0, "One\r\nTwo\rThree");
        assert_eq!(t.paragraphs(), vec![(0, 4), (4, 8), (8, 13)]);
        assert_eq!(t.line_sizes().len(), 3);
    }
}
