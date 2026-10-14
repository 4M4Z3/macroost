//! The storage a word processor provides. The object model reads and edits a document only
//! through the [`Document`] trait, so any application can run macros by implementing it.
//! [`MemoryDocument`] is the reference implementation.
//!
//! A document is one story: a sequence of positions, each holding one [`Item`]. Ordinary
//! characters are items; so is every paragraph mark (`'\r'`, carrying the paragraph's
//! [`ParaFormat`]) and every table structure mark (end of cell, end of row: [`Marker`]). This is
//! the coordinate system macros see: `Range(Start, End)` counts positions, a structure mark
//! occupies one position (its text reads as `Chr(13) & Chr(7)`), and the story always ends with
//! a paragraph mark that cannot be deleted.

/// Character formatting of one position.
#[derive(Clone, Debug, PartialEq)]
pub struct CharFormat {
    pub name: String,
    /// Points, in steps of 0.5.
    pub size: f32,
    pub bold: bool,
    pub italic: bool,
    /// Underline style constant (`wdUnderlineNone` = 0, `wdUnderlineSingle` = 1, ...).
    pub underline: i32,
    pub strike_through: bool,
    pub superscript: bool,
    pub subscript: bool,
    pub all_caps: bool,
    pub small_caps: bool,
    pub hidden: bool,
    /// RGB long (`red + green * 256 + blue * 65536`) or `wdColorAutomatic` (-16777216).
    pub color: i32,
    /// Highlight color index (`wdNoHighlight` = 0).
    pub highlight: i32,
}

impl Default for CharFormat {
    fn default() -> Self {
        Self {
            name: "Calibri".into(),
            size: 11.0,
            bold: false,
            italic: false,
            underline: 0,
            strike_through: false,
            superscript: false,
            subscript: false,
            all_caps: false,
            small_caps: false,
            hidden: false,
            color: COLOR_AUTOMATIC,
            highlight: 0,
        }
    }
}

/// `wdColorAutomatic`.
pub const COLOR_AUTOMATIC: i32 = -16_777_216;

/// Paragraph formatting, stored on the paragraph's terminating mark.
#[derive(Clone, Debug, PartialEq)]
pub struct ParaFormat {
    /// `wdAlignParagraphLeft` = 0, center = 1, right = 2, justify = 3.
    pub alignment: i32,
    /// Points.
    pub left_indent: f32,
    pub right_indent: f32,
    pub first_line_indent: f32,
    pub space_before: f32,
    pub space_after: f32,
    /// Name of the paragraph style.
    pub style: String,
}

impl Default for ParaFormat {
    fn default() -> Self {
        Self {
            alignment: 0,
            left_indent: 0.0,
            right_indent: 0.0,
            first_line_indent: 0.0,
            space_before: 0.0,
            space_after: 0.0,
            style: "Normal".into(),
        }
    }
}

/// Table structure marks.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum Marker {
    #[default]
    None,
    /// End of a cell.
    CellEnd,
    /// End of a row.
    RowEnd,
}

/// One position of the story.
#[derive(Clone, Debug, PartialEq)]
pub struct Item {
    /// The character; `'\r'` for a paragraph mark, `'\u{7}'` for a structure mark.
    pub ch: char,
    pub font: CharFormat,
    /// Meaningful on paragraph marks and structure marks (which end a paragraph too).
    pub para: ParaFormat,
    /// The table this position belongs to (cell content and structure marks).
    pub table: Option<u32>,
    pub marker: Marker,
}

impl Item {
    pub fn char(ch: char, font: CharFormat) -> Self {
        Self { ch, font, para: ParaFormat::default(), table: None, marker: Marker::None }
    }

    /// True for positions that end a paragraph: paragraph marks and structure marks.
    pub fn ends_paragraph(&self) -> bool {
        self.ch == '\r' || self.marker != Marker::None
    }

    /// The text a macro reads for this position.
    pub fn text(&self) -> &'static str {
        match self.marker {
            Marker::None => "",
            _ => "\r\u{7}",
        }
    }
}

/// A named span of the story.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Bookmark {
    pub name: String,
    pub start: usize,
    pub end: usize,
}

/// What the object model needs from a word-processing document.
pub trait Document {
    /// File name shown by `Document.Name`.
    fn name(&self) -> String;
    /// Number of positions, including the final paragraph mark (always at least 1).
    fn len(&self) -> usize;
    fn is_empty(&self) -> bool {
        self.len() == 0
    }
    /// The item at `pos`.
    fn item(&self, pos: usize) -> Option<&Item>;
    /// Replaces positions `start..end` with `items`.
    fn splice(&mut self, start: usize, end: usize, items: Vec<Item>);
    /// Mutable access to one position (formatting changes).
    fn item_mut(&mut self, pos: usize) -> Option<&mut Item>;
    fn bookmarks(&self) -> &[Bookmark];
    fn bookmarks_mut(&mut self) -> &mut Vec<Bookmark>;
    /// Character formatting of the `Normal` style (the document defaults).
    fn default_font(&self) -> CharFormat {
        CharFormat::default()
    }
    /// A fresh identity for a new table.
    fn new_table_id(&mut self) -> u32;
    /// Called when the macro saves the document; the default does nothing.
    fn save(&mut self) {}
}

/// A complete in-memory document.
#[derive(Clone, Debug)]
pub struct MemoryDocument {
    name: String,
    items: Vec<Item>,
    bookmarks: Vec<Bookmark>,
    default_font: CharFormat,
    next_table: u32,
    pub saved_count: usize,
}

impl MemoryDocument {
    /// A document holding one empty paragraph.
    pub fn new(name: &str) -> Self {
        Self::with_paragraphs(name, &[])
    }

    /// A document with one paragraph per entry (an empty slice gives one empty paragraph).
    pub fn with_paragraphs(name: &str, paragraphs: &[&str]) -> Self {
        let font = CharFormat::default();
        let mut items = Vec::new();
        for p in paragraphs {
            items.extend(p.chars().filter(|c| *c != '\r').map(|c| Item::char(c, font.clone())));
            items.push(Item::char('\r', font.clone()));
        }
        if items.is_empty() {
            items.push(Item::char('\r', font.clone()));
        }
        Self { name: name.into(), items, bookmarks: Vec::new(), default_font: font, next_table: 1, saved_count: 0 }
    }

    /// The whole story as plain text (structure marks read as `Chr(13) & Chr(7)`).
    pub fn text(&self) -> String {
        self.items
            .iter()
            .map(|i| if i.marker == Marker::None { i.ch.to_string() } else { i.text().to_owned() })
            .collect()
    }

    pub fn items(&self) -> &[Item] {
        &self.items
    }
}

impl Document for MemoryDocument {
    fn name(&self) -> String {
        self.name.clone()
    }
    fn len(&self) -> usize {
        self.items.len()
    }
    fn item(&self, pos: usize) -> Option<&Item> {
        self.items.get(pos)
    }
    fn splice(&mut self, start: usize, end: usize, items: Vec<Item>) {
        let end = end.min(self.items.len());
        let start = start.min(end);
        self.items.splice(start..end, items);
    }
    fn item_mut(&mut self, pos: usize) -> Option<&mut Item> {
        self.items.get_mut(pos)
    }
    fn bookmarks(&self) -> &[Bookmark] {
        &self.bookmarks
    }
    fn bookmarks_mut(&mut self) -> &mut Vec<Bookmark> {
        &mut self.bookmarks
    }
    fn default_font(&self) -> CharFormat {
        self.default_font.clone()
    }
    fn new_table_id(&mut self) -> u32 {
        self.next_table += 1;
        self.next_table
    }
    fn save(&mut self) {
        self.saved_count += 1;
    }
}
