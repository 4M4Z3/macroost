//! The styles part: decoding cell formats (`cellXfs` -> [`CellFormat`]) and writing them back by
//! appending to the original style tables, so styles the model cannot express survive.

use std::collections::HashMap;

use macroost_sheet::{BorderLine, Borders, CellFormat, HAlign};

use crate::xml::{self, Element, Outline};

/// Text of a built-in number format id (the ids every reader knows without a `numFmt` entry).
pub fn builtin_format(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);\\(\"$\"#,##0\\)",
        6 => "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)",
        7 => "\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)",
        8 => "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

fn builtin_id(code: &str) -> Option<u32> {
    (0..=49).find(|&i| builtin_format(i) == Some(code))
}

/// The legacy 64-color indexed palette (RGB hex), plus system foreground/background.
const PALETTE: [u32; 66] = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000, 0xFFFFFF, 0xFF0000,
    0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080,
    0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080,
    0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99,
    0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699,
    0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333, 0x000000, 0xFFFFFF,
];

/// `0xRRGGBB` -> the object model's RGB long (`r + g*256 + b*65536`).
fn hex_to_long(rgb: u32) -> u32 {
    let r = (rgb >> 16) & 0xff;
    let g = (rgb >> 8) & 0xff;
    let b = rgb & 0xff;
    r | (g << 8) | (b << 16)
}

fn long_to_hex(v: u32) -> String {
    let r = v & 0xff;
    let g = (v >> 8) & 0xff;
    let b = (v >> 16) & 0xff;
    format!("FF{r:02X}{g:02X}{b:02X}")
}

fn parse_hex(s: &str) -> Option<u32> {
    let s = s.trim();
    let s = if s.len() == 8 { s.get(2..)? } else { s };
    if s.len() != 6 {
        return None;
    }
    u32::from_str_radix(s, 16).ok()
}

/// Applies a tint (-1..1) to an `0xRRGGBB` color in HLS space.
fn apply_tint(rgb: u32, tint: f64) -> u32 {
    if tint == 0.0 || !tint.is_finite() {
        return rgb;
    }
    let f = |s: u32| f64::from((rgb >> s) & 0xff) / 255.0;
    let (r, g, b) = (f(16), f(8), f(0));
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let mut l = (max + min) / 2.0;
    let d = max - min;
    let (h, s) = if d == 0.0 {
        (0.0, 0.0)
    } else {
        let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
        let h = if max == r {
            (g - b) / d + if g < b { 6.0 } else { 0.0 }
        } else if max == g {
            (b - r) / d + 2.0
        } else {
            (r - g) / d + 4.0
        };
        (h / 6.0, s)
    };
    let tint = tint.clamp(-1.0, 1.0);
    l = if tint < 0.0 { l * (1.0 + tint) } else { l * (1.0 - tint) + tint };
    let hue = |p: f64, q: f64, mut t: f64| {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    let (r, g, b) = if s == 0.0 {
        (l, l, l)
    } else {
        let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
        let p = 2.0 * l - q;
        (hue(p, q, h + 1.0 / 3.0), hue(p, q, h), hue(p, q, h - 1.0 / 3.0))
    };
    let c = |x: f64| (x.clamp(0.0, 1.0) * 255.0).round() as u32;
    (c(r) << 16) | (c(g) << 8) | c(b)
}

/// Theme colors in the order color references index them (lt1, dk1, lt2, dk2, accents, links).
pub fn parse_theme(text: &str) -> Vec<u32> {
    let Ok(root) = xml::parse_tree(text) else { return Vec::new() };
    let mut scheme = None;
    let mut stack = vec![&root];
    while let Some(e) = stack.pop() {
        if e.local() == "clrScheme" {
            scheme = Some(e);
            break;
        }
        stack.extend(e.elements());
    }
    let Some(scheme) = scheme else { return Vec::new() };
    let get = |n: &str| -> u32 {
        scheme
            .child(n)
            .and_then(|c| c.elements().next())
            .and_then(|c| c.attr("lastClr").or_else(|| c.attr("val")).and_then(parse_hex))
            .unwrap_or(0)
    };
    ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
        .iter()
        .map(|n| get(n))
        .collect()
}

#[derive(Clone, Debug, PartialEq, Default)]
struct FontDec {
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<u32>,
    size: Option<f64>,
    name: Option<String>,
}

/// What a `cellXfs` entry carried in the original file, beyond its decoded format.
pub struct Styles {
    custom_fmts: Vec<(u32, String)>,
    fonts_raw: Vec<String>,
    fonts: Vec<FontDec>,
    fills_raw: Vec<String>,
    fills: Vec<Option<u32>>,
    borders_raw: Vec<String>,
    borders: Vec<Borders>,
    xfs_raw: Vec<String>,
    xfs: Vec<CellFormat>,
    default_font: Option<Element>,
    prefix: String,
    src: String,
    lookup: HashMap<String, u32>,
}

fn key(f: &CellFormat) -> String {
    format!("{f:?}")
}

struct Ctx<'a> {
    theme: &'a [u32],
    palette: Vec<u32>,
}

impl Ctx<'_> {
    /// A color element -> RGB long; `None` for automatic or unresolvable.
    fn color(&self, e: Option<&Element>) -> Option<u32> {
        let e = e?;
        if e.attr("auto").is_some_and(truthy) {
            return None;
        }
        let tint = e.attr("tint").and_then(|t| t.parse::<f64>().ok()).unwrap_or(0.0);
        let rgb = if let Some(rgb) = e.attr("rgb") {
            parse_hex(rgb)?
        } else if let Some(t) = e.attr("theme") {
            *self.theme.get(t.parse::<usize>().ok()?)?
        } else {
            *self.palette.get(e.attr("indexed")?.parse::<usize>().ok()?)?
        };
        Some(hex_to_long(apply_tint(rgb, tint)))
    }
}

pub fn truthy(v: &str) -> bool {
    matches!(v.trim(), "1" | "true" | "on")
}

fn border_style(name: &str) -> Option<(i32, i32)> {
    Some(match name {
        "hair" => (1, 1),
        "thin" => (1, 2),
        "medium" => (1, -4138),
        "thick" => (1, 4),
        "dashed" => (-4115, 2),
        "mediumDashed" => (-4115, -4138),
        "dotted" => (-4118, 2),
        "double" => (-4119, 4),
        "dashDot" => (4, 2),
        "mediumDashDot" => (4, -4138),
        "dashDotDot" => (5, 2),
        "mediumDashDotDot" => (5, -4138),
        "slantDashDot" => (13, -4138),
        _ => return None,
    })
}

fn border_name(style: i32, weight: i32) -> Option<&'static str> {
    let medium = weight == -4138 || weight == 4;
    Some(match style {
        -4142 => return None,
        -4115 => {
            if medium {
                "mediumDashed"
            } else {
                "dashed"
            }
        }
        -4118 => "dotted",
        -4119 => "double",
        4 => {
            if medium {
                "mediumDashDot"
            } else {
                "dashDot"
            }
        }
        5 => {
            if medium {
                "mediumDashDotDot"
            } else {
                "dashDotDot"
            }
        }
        13 => "slantDashDot",
        _ => match weight {
            1 => "hair",
            -4138 => "medium",
            4 => "thick",
            _ => "thin",
        },
    })
}

fn halign(s: &str) -> HAlign {
    match s {
        "left" => HAlign::Left,
        "center" => HAlign::Center,
        "right" => HAlign::Right,
        "fill" => HAlign::Fill,
        "justify" => HAlign::Justify,
        "centerContinuous" => HAlign::CenterAcrossSelection,
        "distributed" => HAlign::Distributed,
        _ => HAlign::General,
    }
}

fn halign_name(a: HAlign) -> Option<&'static str> {
    Some(match a {
        HAlign::General => return None,
        HAlign::Left => "left",
        HAlign::Center => "center",
        HAlign::Right => "right",
        HAlign::Fill => "fill",
        HAlign::Justify => "justify",
        HAlign::CenterAcrossSelection => "centerContinuous",
        HAlign::Distributed => "distributed",
    })
}

fn flag(e: &Element, name: &str) -> bool {
    e.child(name).is_some_and(|b| b.attr("val").is_none_or(truthy))
}

/// A minimal styles part: one default font, the two required fills, one empty border, one xf.
pub const MINIMAL_STYLES: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/></styleSheet>";

impl Styles {
    /// Parses a styles part. A broken part degrades to the minimal styles, never an error.
    pub fn parse(text: Option<&str>, theme: &[u32]) -> Self {
        let text = text.unwrap_or(MINIMAL_STYLES);
        match Self::try_parse(text, theme) {
            Some(s) => s,
            None => Self::try_parse(MINIMAL_STYLES, theme).unwrap_or_else(Self::empty),
        }
    }

    fn empty() -> Self {
        Self {
            custom_fmts: Vec::new(),
            fonts_raw: Vec::new(),
            fonts: Vec::new(),
            fills_raw: Vec::new(),
            fills: Vec::new(),
            borders_raw: Vec::new(),
            borders: Vec::new(),
            xfs_raw: Vec::new(),
            xfs: Vec::new(),
            default_font: None,
            prefix: String::new(),
            src: MINIMAL_STYLES.to_owned(),
            lookup: HashMap::new(),
        }
    }

    fn try_parse(text: &str, theme: &[u32]) -> Option<Self> {
        let root = xml::parse_tree(text).ok()?;
        if root.local() != "styleSheet" {
            return None;
        }
        let prefix = root.name.split_once(':').map(|(p, _)| format!("{p}:")).unwrap_or_default();
        let mut palette = PALETTE.to_vec();
        if let Some(ic) = root.child("colors").and_then(|c| c.child("indexedColors")) {
            for (i, c) in ic.children_named("rgbColor").enumerate() {
                if let (Some(slot), Some(v)) = (palette.get_mut(i), c.attr("rgb").and_then(parse_hex)) {
                    *slot = v;
                }
            }
        }
        let ctx = Ctx { theme, palette };
        let custom_fmts: Vec<(u32, String)> = root
            .child("numFmts")
            .map(|n| {
                n.children_named("numFmt")
                    .filter_map(|f| {
                        Some((f.attr("numFmtId")?.trim().parse().ok()?, f.attr("formatCode").unwrap_or("").to_owned()))
                    })
                    .collect()
            })
            .unwrap_or_default();
        let font_elems: Vec<&Element> =
            root.child("fonts").map(|f| f.children_named("font").collect()).unwrap_or_default();
        let default_font = font_elems.first().map(|e| (*e).clone());
        let def_name =
            default_font.as_ref().and_then(|f| f.child("name")).and_then(|n| n.attr("val")).map(str::to_owned);
        let def_size = default_font.as_ref().and_then(|f| f.child("sz")).and_then(|n| n.attr("val")).map(str::to_owned);
        let def_color = default_font.as_ref().and_then(|f| f.child("color")).map(Element::to_xml);
        let fonts = font_elems
            .iter()
            .map(|f| {
                let name = f.child("name").and_then(|n| n.attr("val")).map(str::to_owned);
                let size_raw = f.child("sz").and_then(|n| n.attr("val")).map(str::to_owned);
                let color_raw = f.child("color").map(Element::to_xml);
                FontDec {
                    bold: flag(f, "b"),
                    italic: flag(f, "i"),
                    underline: f.child("u").is_some_and(|u| u.attr("val") != Some("none")),
                    color: if color_raw.is_none() || color_raw == def_color {
                        None
                    } else {
                        ctx.color(f.child("color"))
                    },
                    size: if size_raw == def_size { None } else { size_raw.and_then(|s| s.trim().parse().ok()) },
                    name: if name == def_name { None } else { name },
                }
            })
            .collect();
        let fill_elems: Vec<&Element> =
            root.child("fills").map(|f| f.children_named("fill").collect()).unwrap_or_default();
        let fills = fill_elems
            .iter()
            .map(|f| {
                let p = f.child("patternFill")?;
                match p.attr("patternType") {
                    None | Some("none") => None,
                    Some(_) => ctx.color(p.child("fgColor")).or_else(|| ctx.color(p.child("bgColor"))),
                }
            })
            .collect();
        let border_elems: Vec<&Element> =
            root.child("borders").map(|f| f.children_named("border").collect()).unwrap_or_default();
        let borders = border_elems
            .iter()
            .map(|b| {
                let side = |names: &[&str]| -> Option<BorderLine> {
                    let e = names.iter().find_map(|n| b.child(n))?;
                    let (style, weight) = border_style(e.attr("style")?)?;
                    Some(BorderLine { style, weight, color: ctx.color(e.child("color")).unwrap_or(0) })
                };
                Borders {
                    left: side(&["left", "start"]),
                    top: side(&["top"]),
                    bottom: side(&["bottom"]),
                    right: side(&["right", "end"]),
                    ..Borders::default()
                }
            })
            .collect::<Vec<_>>();
        let xf_elems: Vec<&Element> =
            root.child("cellXfs").map(|f| f.children_named("xf").collect()).unwrap_or_default();
        let mut s = Self {
            custom_fmts,
            fonts_raw: font_elems.iter().map(|e| e.to_xml()).collect(),
            fonts,
            fills_raw: fill_elems.iter().map(|e| e.to_xml()).collect(),
            fills,
            borders_raw: border_elems.iter().map(|e| e.to_xml()).collect(),
            borders,
            xfs_raw: xf_elems.iter().map(|e| e.to_xml()).collect(),
            xfs: Vec::new(),
            default_font,
            prefix,
            src: text.to_owned(),
            lookup: HashMap::new(),
        };
        s.xfs = xf_elems.iter().map(|x| s.decode_xf(x)).collect();
        for (i, f) in s.xfs.iter().enumerate() {
            s.lookup.entry(key(f)).or_insert(i as u32);
        }
        Some(s)
    }

    fn num_format(&self, id: u32) -> Option<String> {
        let code = match self.custom_fmts.iter().find(|(i, _)| *i == id) {
            Some((_, c)) => c.clone(),
            None => builtin_format(id)?.to_owned(),
        };
        (!code.eq_ignore_ascii_case("general") && !code.is_empty()).then_some(code)
    }

    fn decode_xf(&self, x: &Element) -> CellFormat {
        let id = |n: &str| x.attr(n).and_then(|v| v.trim().parse::<usize>().ok()).unwrap_or(0);
        let font = self.fonts.get(id("fontId")).cloned().unwrap_or_default();
        let align = x.child("alignment");
        CellFormat {
            number_format: self.num_format(id("numFmtId") as u32),
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            font_color: font.color,
            font_size: font.size,
            font_name: font.name,
            fill_color: self.fills.get(id("fillId")).copied().flatten(),
            borders: self.borders.get(id("borderId")).copied().unwrap_or_default(),
            h_align: align.and_then(|a| a.attr("horizontal")).map(halign).unwrap_or_default(),
            wrap: align.and_then(|a| a.attr("wrapText")).is_some_and(truthy),
            unlocked: x.child("protection").and_then(|p| p.attr("locked")).is_some_and(|v| !truthy(v)),
            formula_hidden: x.child("protection").and_then(|p| p.attr("hidden")).is_some_and(truthy),
            ..CellFormat::default()
        }
    }

    /// The decoded format of `cellXfs` entry `s` (out of range: the default format).
    pub fn format(&self, s: u32) -> CellFormat {
        self.xfs.get(s as usize).cloned().unwrap_or_default()
    }

    // ------------------------------------------------------------ writing

    /// The `cellXfs` index for `fmt`, preferring `preferred` (the cell's original style) when it
    /// still decodes to the same format; new fonts/fills/borders/xfs are appended as needed.
    pub fn index_for(&mut self, fmt: &CellFormat, preferred: Option<u32>) -> u32 {
        if let Some(p) = preferred
            && self.xfs.get(p as usize) == Some(fmt)
        {
            return p;
        }
        let k = key(fmt);
        if let Some(&i) = self.lookup.get(&k) {
            return i;
        }
        let num = self.num_id(fmt.number_format.as_deref());
        let font = self.font_id(fmt);
        let fill = self.fill_id(fmt.fill_color);
        let border = self.border_id(&fmt.borders);
        let p = &self.prefix;
        let mut xf = format!(
            "<{p}xf numFmtId=\"{num}\" fontId=\"{font}\" fillId=\"{fill}\" borderId=\"{border}\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\""
        );
        let h = halign_name(fmt.h_align);
        let protect = fmt.unlocked || fmt.formula_hidden;
        if protect {
            xf.push_str(" applyProtection=\"1\"");
        }
        if h.is_some() || fmt.wrap {
            xf.push_str(" applyAlignment=\"1\"");
        }
        if h.is_some() || fmt.wrap || protect {
            xf.push('>');
            if h.is_some() || fmt.wrap {
                xf.push_str(&format!("<{p}alignment"));
                if let Some(h) = h {
                    xf.push_str(&format!(" horizontal=\"{h}\""));
                }
                if fmt.wrap {
                    xf.push_str(" wrapText=\"1\"");
                }
                xf.push_str("/>");
            }
            if protect {
                xf.push_str(&format!("<{p}protection"));
                if fmt.unlocked {
                    xf.push_str(" locked=\"0\"");
                }
                if fmt.formula_hidden {
                    xf.push_str(" hidden=\"1\"");
                }
                xf.push_str("/>");
            }
            xf.push_str(&format!("</{p}xf>"));
        } else {
            xf.push_str("/>");
        }
        let i = self.xfs.len() as u32;
        self.xfs_raw.push(xf);
        self.xfs.push(fmt.clone());
        self.lookup.insert(k, i);
        i
    }

    fn num_id(&mut self, code: Option<&str>) -> u32 {
        let Some(code) = code else { return 0 };
        if code.eq_ignore_ascii_case("general") {
            return 0;
        }
        if let Some((i, _)) = self.custom_fmts.iter().find(|(_, c)| c == code) {
            return *i;
        }
        if let Some(i) = builtin_id(code) {
            return i;
        }
        let id = self.custom_fmts.iter().map(|(i, _)| *i).max().unwrap_or(163).max(163) + 1;
        self.custom_fmts.push((id, code.to_owned()));
        id
    }

    fn font_id(&mut self, fmt: &CellFormat) -> u32 {
        let want = FontDec {
            bold: fmt.bold,
            italic: fmt.italic,
            underline: fmt.underline,
            color: fmt.font_color,
            size: fmt.font_size,
            name: fmt.font_name.clone(),
        };
        if let Some(i) = self.fonts.iter().position(|f| *f == want) {
            return i as u32;
        }
        let p = self.prefix.clone();
        let base = self.default_font.clone().unwrap_or_else(|| Element::new(&format!("{p}font")));
        let mut parts: HashMap<&str, String> = HashMap::new();
        for e in base.elements() {
            if let Some(k) = FONT_ORDER.iter().find(|k| **k == e.local()) {
                parts.insert(k, e.to_xml());
            }
        }
        for (name, on) in [("b", want.bold), ("i", want.italic), ("u", want.underline)] {
            if on {
                parts.insert(name, format!("<{p}{name}/>"));
            } else {
                parts.remove(name);
            }
        }
        if let Some(sz) = want.size {
            parts.insert("sz", format!("<{p}sz val=\"{sz}\"/>"));
        }
        if let Some(c) = want.color {
            parts.insert("color", format!("<{p}color rgb=\"{}\"/>", long_to_hex(c)));
        }
        if let Some(n) = &want.name {
            parts.insert("name", format!("<{p}name val=\"{}\"/>", xml::escape(n)));
            parts.remove("scheme");
            parts.remove("family");
            parts.remove("charset");
        }
        let mut raw = format!("<{p}font>");
        for k in FONT_ORDER {
            if let Some(v) = parts.get(k) {
                raw.push_str(v);
            }
        }
        raw.push_str(&format!("</{p}font>"));
        self.fonts_raw.push(raw);
        self.fonts.push(want);
        (self.fonts.len() - 1) as u32
    }

    fn fill_id(&mut self, color: Option<u32>) -> u32 {
        let Some(c) = color else {
            return self.fills.iter().position(Option::is_none).unwrap_or(0) as u32;
        };
        if let Some(i) = self.fills.iter().position(|f| *f == Some(c)) {
            return i as u32;
        }
        let p = &self.prefix;
        self.fills_raw.push(format!(
            "<{p}fill><{p}patternFill patternType=\"solid\"><{p}fgColor rgb=\"{}\"/><{p}bgColor indexed=\"64\"/></{p}patternFill></{p}fill>",
            long_to_hex(c)
        ));
        self.fills.push(Some(c));
        (self.fills.len() - 1) as u32
    }

    fn border_id(&mut self, b: &Borders) -> u32 {
        if let Some(i) = self.borders.iter().position(|x| x == b) {
            return i as u32;
        }
        let p = &self.prefix;
        let side =
            |name: &str, l: Option<BorderLine>| match l.and_then(|l| border_name(l.style, l.weight).map(|n| (n, l))) {
                Some((n, l)) => {
                    format!("<{p}{name} style=\"{n}\"><{p}color rgb=\"{}\"/></{p}{name}>", long_to_hex(l.color))
                }
                None => format!("<{p}{name}/>"),
            };
        let raw = format!(
            "<{p}border>{}{}{}{}<{p}diagonal/></{p}border>",
            side("left", b.left),
            side("right", b.right),
            side("top", b.top),
            side("bottom", b.bottom)
        );
        self.borders_raw.push(raw);
        // Store the format as the reader will decode it, so later lookups hit.
        let norm = |l: Option<BorderLine>| {
            l.and_then(|l| {
                border_name(l.style, l.weight).and_then(border_style).map(|(style, weight)| BorderLine {
                    style,
                    weight,
                    color: l.color,
                })
            })
        };
        let dec = Borders {
            left: norm(b.left),
            top: norm(b.top),
            bottom: norm(b.bottom),
            right: norm(b.right),
            ..Borders::default()
        };
        if dec != *b {
            // Unrepresentable combination: remember under the requested key too.
            self.borders.push(*b);
        } else {
            self.borders.push(dec);
        }
        (self.borders.len() - 1) as u32
    }

    /// The styles part with the regenerated tables spliced into the original text.
    pub fn to_xml(&self) -> String {
        let p = &self.prefix;
        let block = |name: &str, raws: &[String]| -> String {
            let mut s = format!("<{p}{name} count=\"{}\">", raws.len());
            for r in raws {
                s.push_str(r);
            }
            s.push_str(&format!("</{p}{name}>"));
            s
        };
        let numfmts = if self.custom_fmts.is_empty() {
            None
        } else {
            let raws: Vec<String> = self
                .custom_fmts
                .iter()
                .map(|(i, c)| format!("<{p}numFmt numFmtId=\"{i}\" formatCode=\"{}\"/>", xml::escape(c)))
                .collect();
            Some(block("numFmts", &raws))
        };
        let mut fonts = self.fonts_raw.clone();
        if fonts.is_empty() {
            fonts.push(format!("<{p}font><{p}sz val=\"11\"/><{p}name val=\"Calibri\"/></{p}font>"));
        }
        let mut fills = self.fills_raw.clone();
        while fills.len() < 2 {
            let kind = if fills.is_empty() { "none" } else { "gray125" };
            fills.push(format!("<{p}fill><{p}patternFill patternType=\"{kind}\"/></{p}fill>"));
        }
        let mut borders = self.borders_raw.clone();
        if borders.is_empty() {
            borders.push(format!("<{p}border><{p}left/><{p}right/><{p}top/><{p}bottom/><{p}diagonal/></{p}border>"));
        }
        let mut xfs = self.xfs_raw.clone();
        if xfs.is_empty() {
            xfs.push(format!("<{p}xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"));
        }
        let replace = [
            ("numFmts", numfmts),
            ("fonts", Some(block("fonts", &fonts))),
            ("fills", Some(block("fills", &fills))),
            ("borders", Some(block("borders", &borders))),
            ("cellXfs", Some(block("cellXfs", &xfs))),
        ];
        match Outline::parse(&self.src) {
            Ok(o) => o.rebuild(STYLE_ORDER, &replace),
            Err(_) => {
                let mut s = String::from(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
                );
                for (_, v) in replace.into_iter() {
                    if let Some(v) = v {
                        s.push_str(&v);
                    }
                }
                s.push_str("</styleSheet>");
                s
            }
        }
    }
}

const FONT_ORDER: [&str; 15] = [
    "b",
    "i",
    "strike",
    "condense",
    "extend",
    "outline",
    "shadow",
    "u",
    "vertAlign",
    "sz",
    "color",
    "name",
    "family",
    "charset",
    "scheme",
];

const STYLE_ORDER: &[&str] = &[
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tint_and_colors() {
        assert_eq!(apply_tint(0xFFFFFF, -0.5), 0x808080);
        assert_eq!(apply_tint(0x000000, 0.5), 0x808080);
        assert_eq!(hex_to_long(0xFF0000), 255);
        assert_eq!(long_to_hex(255), "FFFF0000");
    }

    #[test]
    fn appended_formats_round_trip() {
        let mut s = Styles::parse(None, &[]);
        let f = CellFormat {
            number_format: Some("0.000".into()),
            bold: true,
            font_color: Some(255),
            fill_color: Some(0x00ff00),
            borders: Borders { left: Some(BorderLine::default()), ..Borders::default() },
            h_align: HAlign::Center,
            wrap: true,
            ..CellFormat::default()
        };
        let i = s.index_for(&f, None);
        assert_eq!(i, 1);
        assert_eq!(s.index_for(&CellFormat::default(), None), 0);
        let xml = s.to_xml();
        let s2 = Styles::parse(Some(&xml), &[]);
        assert_eq!(s2.format(1), f);
    }
}
