//! A small, non-panicking XML pull tokenizer and tree builder, enough for the parts of an office-suite
//! Open XML package: elements, attributes, text, the five predefined entities, numeric character
//! references, CDATA, comments, processing instructions and a skipped DOCTYPE (no custom
//! entities are ever expanded). Namespace prefixes are kept as written; callers match on local
//! names. Every event carries byte offsets into the source so callers can splice parts.

use std::borrow::Cow;

/// Maximum element nesting accepted.
const MAX_DEPTH: usize = 512;

#[derive(Debug, Clone)]
pub enum Event<'a> {
    Start { name: &'a str, attrs: Vec<(&'a str, Cow<'a, str>)>, empty: bool, start: usize, end: usize },
    End { name: &'a str, start: usize, end: usize },
    Text(Cow<'a, str>),
}

/// The part of a qualified name after the prefix.
pub fn local(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, l)| l)
}

/// The value of attribute `name` (matched by local name; an exact match wins).
pub fn attr<'b>(attrs: &'b [(&str, Cow<'_, str>)], name: &str) -> Option<&'b str> {
    attrs
        .iter()
        .find(|(k, _)| *k == name)
        .or_else(|| attrs.iter().find(|(k, _)| local(k) == name && !k.starts_with("xmlns")))
        .map(|(_, v)| v.as_ref())
}

pub struct Reader<'a> {
    src: &'a str,
    pos: usize,
    stack: Vec<&'a str>,
    done: bool,
}

impl<'a> Reader<'a> {
    pub fn new(src: &'a str) -> Self {
        let src = src.strip_prefix('\u{feff}').unwrap_or(src);
        Self { src, pos: 0, stack: Vec::new(), done: false }
    }

    /// Current nesting depth (number of open elements).
    pub fn depth(&self) -> usize {
        self.stack.len()
    }

    pub fn source(&self) -> &'a str {
        self.src
    }

    fn find_from(&self, pat: &str) -> Option<usize> {
        self.src.get(self.pos..)?.find(pat).map(|i| i + self.pos)
    }

    /// The next event, `None` at the end, or an error message.
    #[allow(clippy::should_implement_trait)]
    pub fn next(&mut self) -> Option<Result<Event<'a>, String>> {
        if self.done {
            return None;
        }
        let r = self.step();
        match r {
            Ok(Some(e)) => Some(Ok(e)),
            Ok(None) => {
                self.done = true;
                if let Some(open) = self.stack.last() {
                    return Some(Err(format!("unexpected end of document inside <{open}>")));
                }
                None
            }
            Err(e) => {
                self.done = true;
                Some(Err(e))
            }
        }
    }

    fn step(&mut self) -> Result<Option<Event<'a>>, String> {
        loop {
            let Some(rest) = self.src.get(self.pos..) else { return Ok(None) };
            if rest.is_empty() {
                return Ok(None);
            }
            if !rest.starts_with('<') {
                let end = rest.find('<').map_or(self.src.len(), |i| i + self.pos);
                let raw = self.src.get(self.pos..end).unwrap_or("");
                self.pos = end;
                return Ok(Some(Event::Text(unescape(raw))));
            }
            if rest.starts_with("<?") {
                let end = self.find_from("?>").ok_or("unterminated processing instruction")?;
                self.pos = end + 2;
                continue;
            }
            if rest.starts_with("<!--") {
                let end = self.find_from("-->").ok_or("unterminated comment")?;
                self.pos = end + 3;
                continue;
            }
            if rest.starts_with("<![CDATA[") {
                let start = self.pos + 9;
                let end = self.find_from("]]>").ok_or("unterminated CDATA section")?;
                let text = self.src.get(start..end).unwrap_or("");
                self.pos = end + 3;
                return Ok(Some(Event::Text(Cow::Borrowed(text))));
            }
            if rest.starts_with("<!") {
                // DOCTYPE or another declaration: skip, including an internal subset.
                let mut depth = 0i32;
                let mut end = None;
                for (i, b) in rest.bytes().enumerate() {
                    match b {
                        b'[' => depth += 1,
                        b']' => depth -= 1,
                        b'>' if depth <= 0 => {
                            end = Some(i);
                            break;
                        }
                        _ => {}
                    }
                }
                let end = end.ok_or("unterminated declaration")?;
                self.pos += end + 1;
                continue;
            }
            let start = self.pos;
            if let Some(after) = rest.strip_prefix("</") {
                let close = after.find('>').ok_or("unterminated end tag")?;
                let name = after.get(..close).unwrap_or("").trim();
                self.pos = start + 2 + close + 1;
                match self.stack.pop() {
                    Some(open) if open == name => {}
                    Some(open) => return Err(format!("</{name}> does not close <{open}>")),
                    None => return Err(format!("unexpected </{name}>")),
                }
                return Ok(Some(Event::End { name, start, end: self.pos }));
            }
            return self.start_tag(start).map(Some);
        }
    }

    fn start_tag(&mut self, start: usize) -> Result<Event<'a>, String> {
        let src = self.src;
        let bytes = src.as_bytes();
        let mut i = start + 1;
        let name_start = i;
        while i < bytes.len() && !matches!(bytes[i], b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>') {
            i += 1;
        }
        let name = src.get(name_start..i).unwrap_or("");
        if name.is_empty() {
            return Err("empty element name".into());
        }
        let mut attrs = Vec::new();
        loop {
            while i < bytes.len() && matches!(bytes[i], b' ' | b'\t' | b'\r' | b'\n') {
                i += 1;
            }
            match bytes.get(i) {
                None => return Err(format!("unterminated <{name}>")),
                Some(b'>') => {
                    i += 1;
                    if self.stack.len() >= MAX_DEPTH {
                        return Err("elements nested too deeply".into());
                    }
                    self.stack.push(name);
                    self.pos = i;
                    return Ok(Event::Start { name, attrs, empty: false, start, end: i });
                }
                Some(b'/') => {
                    if bytes.get(i + 1) != Some(&b'>') {
                        return Err(format!("stray '/' in <{name}>"));
                    }
                    i += 2;
                    self.pos = i;
                    return Ok(Event::Start { name, attrs, empty: true, start, end: i });
                }
                Some(_) => {
                    let an_start = i;
                    while i < bytes.len() && !matches!(bytes[i], b'=' | b' ' | b'\t' | b'\r' | b'\n' | b'>' | b'/') {
                        i += 1;
                    }
                    let an = src.get(an_start..i).unwrap_or("");
                    while i < bytes.len() && matches!(bytes[i], b' ' | b'\t' | b'\r' | b'\n') {
                        i += 1;
                    }
                    if bytes.get(i) != Some(&b'=') {
                        return Err(format!("attribute {an} in <{name}> has no value"));
                    }
                    i += 1;
                    while i < bytes.len() && matches!(bytes[i], b' ' | b'\t' | b'\r' | b'\n') {
                        i += 1;
                    }
                    let q = match bytes.get(i) {
                        Some(&q @ (b'"' | b'\'')) => q,
                        _ => return Err(format!("attribute {an} in <{name}> is not quoted")),
                    };
                    i += 1;
                    let v_start = i;
                    while i < bytes.len() && bytes[i] != q {
                        i += 1;
                    }
                    if i >= bytes.len() {
                        return Err(format!("unterminated attribute {an} in <{name}>"));
                    }
                    let raw = src.get(v_start..i).unwrap_or("");
                    i += 1;
                    attrs.push((an, unescape(raw)));
                }
            }
        }
    }
}

/// Replaces entity and character references. Unknown entities are kept literally.
pub fn unescape(raw: &str) -> Cow<'_, str> {
    if !raw.contains('&') {
        return Cow::Borrowed(raw);
    }
    let mut out = String::with_capacity(raw.len());
    let mut rest = raw;
    while let Some(i) = rest.find('&') {
        out.push_str(rest.get(..i).unwrap_or(""));
        let tail = rest.get(i..).unwrap_or("");
        let semi = tail.get(..tail.len().min(12)).and_then(|t| t.find(';'));
        let mut consumed = 1;
        let mut push = None;
        if let Some(s) = semi {
            let ent = tail.get(1..s).unwrap_or("");
            push = match ent {
                "amp" => Some('&'),
                "lt" => Some('<'),
                "gt" => Some('>'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                _ => {
                    let code = if let Some(h) = ent.strip_prefix("#x").or_else(|| ent.strip_prefix("#X")) {
                        u32::from_str_radix(h, 16).ok()
                    } else if let Some(d) = ent.strip_prefix('#') {
                        d.parse::<u32>().ok()
                    } else {
                        None
                    };
                    code.map(|c| char::from_u32(c).unwrap_or('\u{fffd}'))
                }
            };
            if push.is_some() {
                consumed = s + 1;
            }
        }
        match push {
            Some(c) => out.push(c),
            None => out.push('&'),
        }
        rest = tail.get(consumed..).unwrap_or("");
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// Escapes text for element content or a double-quoted attribute. Characters XML 1.0 cannot
/// carry are dropped.
pub fn escape(s: &str) -> Cow<'_, str> {
    if !s.chars().any(|c| matches!(c, '&' | '<' | '>' | '"') || !valid_xml_char(c)) {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len() + 8);
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            c if valid_xml_char(c) => out.push(c),
            _ => {}
        }
    }
    Cow::Owned(out)
}

pub fn valid_xml_char(c: char) -> bool {
    matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..)
}

/// Decodes the `_xHHHH_` escapes used by spreadsheet string types (ST_Xstring).
pub fn decode_xstring(s: &str) -> Cow<'_, str> {
    if !s.contains("_x") {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len());
    let b = s.as_bytes();
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'_'
            && b.get(i + 1) == Some(&b'x')
            && b.get(i + 6) == Some(&b'_')
            && let Some(hex) = s.get(i + 2..i + 6)
            && hex.bytes().all(|h| h.is_ascii_hexdigit())
            && let Ok(code) = u32::from_str_radix(hex, 16)
        {
            out.push(char::from_u32(code).unwrap_or('\u{fffd}'));
            i += 7;
            continue;
        }
        let ch_len = s.get(i..).and_then(|t| t.chars().next()).map_or(1, char::len_utf8);
        out.push_str(s.get(i..i + ch_len).unwrap_or(""));
        i += ch_len;
    }
    Cow::Owned(out)
}

/// Encodes characters XML cannot carry (and literal `_xHHHH_` look-alikes) as `_xHHHH_`, then
/// escapes for XML.
pub fn encode_xstring(s: &str) -> String {
    let needs = s.chars().any(|c| !valid_xml_char(c)) || s.contains("_x");
    if !needs {
        return escape(s).into_owned();
    }
    let mut tmp = String::with_capacity(s.len() + 8);
    let b = s.as_bytes();
    for (i, c) in s.char_indices() {
        if !valid_xml_char(c) {
            tmp.push_str(&format!("_x{:04X}_", u32::from(c) & 0xffff));
        } else if c == '_'
            && b.get(i + 1) == Some(&b'x')
            && b.get(i + 6) == Some(&b'_')
            && s.get(i + 2..i + 6).is_some_and(|h| h.bytes().all(|x| x.is_ascii_hexdigit()))
        {
            tmp.push_str("_x005F_");
        } else {
            tmp.push(c);
        }
    }
    escape(&tmp).into_owned()
}

// ---------------------------------------------------------------- tree

#[derive(Debug, Clone, Default)]
pub struct Element {
    pub name: String,
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Node>,
}

#[derive(Debug, Clone)]
pub enum Node {
    Elem(Element),
    Text(String),
}

impl Element {
    pub fn new(name: &str) -> Self {
        Self { name: name.to_owned(), ..Self::default() }
    }

    pub fn local(&self) -> &str {
        local(&self.name)
    }

    /// Attribute by local name (an exact match wins).
    pub fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == name)
            .or_else(|| self.attrs.iter().find(|(k, _)| local(k) == name && !k.starts_with("xmlns")))
            .map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, name: &str, value: &str) {
        if let Some(a) = self.attrs.iter_mut().find(|(k, _)| k == name) {
            a.1 = value.to_owned();
        } else {
            self.attrs.push((name.to_owned(), value.to_owned()));
        }
    }

    pub fn elements(&self) -> impl Iterator<Item = &Element> {
        self.children.iter().filter_map(|n| if let Node::Elem(e) = n { Some(e) } else { None })
    }

    pub fn children_named<'s>(&'s self, local_name: &'s str) -> impl Iterator<Item = &'s Element> + 's {
        self.elements().filter(move |e| e.local() == local_name)
    }

    pub fn child(&self, local_name: &str) -> Option<&Element> {
        self.elements().find(|e| e.local() == local_name)
    }

    /// Concatenated text of this element and its descendants.
    pub fn text(&self) -> String {
        let mut s = String::new();
        self.collect_text(&mut s);
        s
    }

    fn collect_text(&self, out: &mut String) {
        for n in &self.children {
            match n {
                Node::Text(t) => out.push_str(t),
                Node::Elem(e) => e.collect_text(out),
            }
        }
    }

    /// Serializes the element (no XML declaration).
    pub fn write(&self, out: &mut String) {
        out.push('<');
        out.push_str(&self.name);
        for (k, v) in &self.attrs {
            out.push(' ');
            out.push_str(k);
            out.push_str("=\"");
            out.push_str(&escape(v));
            out.push('"');
        }
        if self.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for n in &self.children {
            match n {
                Node::Text(t) => out.push_str(&escape(t)),
                Node::Elem(e) => e.write(out),
            }
        }
        out.push_str("</");
        out.push_str(&self.name);
        out.push('>');
    }

    pub fn to_xml(&self) -> String {
        let mut s = String::new();
        self.write(&mut s);
        s
    }
}

/// Parses a whole document (or fragment with one root) into a tree.
pub fn parse_tree(src: &str) -> Result<Element, String> {
    let mut r = Reader::new(src);
    let mut stack: Vec<Element> = Vec::new();
    let mut root: Option<Element> = None;
    while let Some(ev) = r.next() {
        match ev? {
            Event::Start { name, attrs, empty, .. } => {
                let e = Element {
                    name: name.to_owned(),
                    attrs: attrs.into_iter().map(|(k, v)| (k.to_owned(), v.into_owned())).collect(),
                    children: Vec::new(),
                };
                if empty {
                    match stack.last_mut() {
                        Some(p) => p.children.push(Node::Elem(e)),
                        None if root.is_none() => root = Some(e),
                        None => {}
                    }
                } else {
                    stack.push(e);
                }
            }
            Event::End { .. } => {
                if let Some(e) = stack.pop() {
                    match stack.last_mut() {
                        Some(p) => p.children.push(Node::Elem(e)),
                        None if root.is_none() => root = Some(e),
                        None => {}
                    }
                }
            }
            Event::Text(t) => {
                if let Some(p) = stack.last_mut() {
                    p.children.push(Node::Text(t.into_owned()));
                }
            }
        }
    }
    root.ok_or_else(|| "no root element".to_owned())
}

/// The root element's start tag and its top-level children as byte spans of `src`.
pub struct Outline<'a> {
    pub src: &'a str,
    /// Span of the root start tag.
    pub root_start: (usize, usize),
    pub root_name: &'a str,
    pub root_empty: bool,
    /// Span of the root end tag (equal to `root_start` for an empty root).
    pub root_end: (usize, usize),
    /// (local name, span) of each child element of the root.
    pub children: Vec<(&'a str, (usize, usize))>,
}

impl<'a> Outline<'a> {
    pub fn parse(src: &'a str) -> Result<Self, String> {
        let mut r = Reader::new(src);
        let src = r.source();
        let mut root: Option<(&str, (usize, usize), bool)> = None;
        let mut root_end = None;
        let mut children = Vec::new();
        let mut child_start: Option<(&str, usize)> = None;
        while let Some(ev) = r.next() {
            match ev? {
                Event::Start { name, empty, start, end, .. } => {
                    if root.is_none() {
                        root = Some((name, (start, end), empty));
                        if empty {
                            root_end = Some((start, end));
                            break;
                        }
                        continue;
                    }
                    // depth after this start: 1 = root only open before; reader pushed if !empty
                    let d = if empty { r.depth() } else { r.depth() - 1 };
                    if d == 1 {
                        if empty {
                            children.push((local(name), (start, end)));
                        } else {
                            child_start = Some((local(name), start));
                        }
                    }
                }
                Event::End { start, end, .. } => {
                    let d = r.depth();
                    if d == 1 {
                        if let Some((n, s)) = child_start.take() {
                            children.push((n, (s, end)));
                        }
                    } else if d == 0 {
                        root_end = Some((start, end));
                        break;
                    }
                }
                Event::Text(_) => {}
            }
        }
        let (root_name, root_start, root_empty) = root.ok_or("no root element")?;
        let root_end = root_end.ok_or("root element is not closed")?;
        Ok(Self { src, root_start, root_name, root_empty, root_end, children })
    }

    pub fn text(&self, span: (usize, usize)) -> &'a str {
        self.src.get(span.0..span.1).unwrap_or("")
    }

    pub fn child(&self, local_name: &str) -> Option<&'a str> {
        self.children.iter().find(|(n, _)| *n == local_name).map(|(_, s)| self.text(*s))
    }

    /// Rebuilds the document: children named in `replace` are removed, and each `Some` new
    /// fragment is inserted at the position `order` (the schema's child sequence) gives it.
    /// Children not named in `order` keep their place relative to their neighbours.
    pub fn rebuild(&self, order: &[&str], replace: &[(&str, Option<String>)]) -> String {
        let rank = |n: &str| order.iter().position(|o| *o == n);
        let mut kept: Vec<(usize, String)> = Vec::new();
        let mut last = 0usize;
        for (name, span) in &self.children {
            let rk = rank(name).unwrap_or(last);
            last = rk;
            if replace.iter().any(|(r, _)| r == name) {
                continue;
            }
            kept.push((rk, self.text(*span).to_owned()));
        }
        for (name, frag) in replace {
            let Some(frag) = frag else { continue };
            let rk = rank(name).unwrap_or(usize::MAX);
            let at = kept.iter().position(|(k, _)| *k > rk).unwrap_or(kept.len());
            kept.insert(at, (rk, frag.clone()));
        }
        let mut out = String::with_capacity(self.src.len() + 1024);
        out.push_str(self.src.get(..self.root_start.0).unwrap_or(""));
        let start_tag = self.text(self.root_start);
        if self.root_empty {
            out.push_str(start_tag.trim_end_matches('>').trim_end_matches('/').trim_end());
            out.push('>');
        } else {
            out.push_str(start_tag);
        }
        for (_, k) in kept {
            out.push_str(&k);
        }
        out.push_str("</");
        out.push_str(self.root_name);
        out.push('>');
        out.push_str(self.src.get(self.root_end.1..).unwrap_or(""));
        out
    }

    /// The prefix of the root element (`x` for `x:worksheet`), if any.
    pub fn root_prefix(&self) -> Option<&'a str> {
        self.root_name.split_once(':').map(|(p, _)| p)
    }
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;

    #[test]
    fn tokenizes_entities_cdata_and_empty_tags() {
        let t = parse_tree("<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY a 'b'>]><a x='1 &amp; 2'><b/>t&lt;&#65;&#x42;&bogus;<![CDATA[<c>]]><!-- c --></a>")
            .unwrap();
        assert_eq!(t.attr("x"), Some("1 & 2"));
        assert_eq!(t.text(), "t<AB&bogus;<c>");
        assert_eq!(t.elements().count(), 1);
    }

    #[test]
    fn hostile_input_errors_instead_of_panicking() {
        for s in
            ["<a", "<a b>", "<a b='1>", "</a>", "<a></b>", "<a>&#xFFFFFFFFF;</a>", "<", "<!--", "<a/", "\u{feff}<é>"]
        {
            let _ = parse_tree(s);
        }
        let deep = "<a>".repeat(10_000);
        assert!(parse_tree(&deep).is_err());
    }

    #[test]
    fn xstring_round_trip() {
        assert_eq!(decode_xstring("a_x000D_b"), "a\rb");
        let s = "x_x0041_\u{1}";
        assert_eq!(decode_xstring(&encode_xstring(s)), s);
    }

    #[test]
    fn outline_rebuild_inserts_by_order() {
        let src = "<?xml version=\"1.0\"?>\n<w xmlns=\"n\"><a/><c>1</c><zz/><e/></w>";
        let o = Outline::parse(src).unwrap();
        assert_eq!(o.children.len(), 4);
        let out = o.rebuild(
            &["a", "b", "c", "d", "e"],
            &[("b", Some("<b/>".into())), ("c", None), ("d", Some("<d/>".into()))],
        );
        assert_eq!(out, "<?xml version=\"1.0\"?>\n<w xmlns=\"n\"><a/><b/><zz/><d/><e/></w>");
    }
}
