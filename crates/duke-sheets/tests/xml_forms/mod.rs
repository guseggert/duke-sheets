//! Rewriting XML between forms that are supposed to mean the same thing.
//!
//! A parser must not care whether an element was written `<x/>` or
//! `<x></x>`, nor whether its siblings were indented. Both chart parsers
//! are hand-written event loops where that is easy to get wrong, so the
//! rewrites live here and each parser's tests apply them to a fixture
//! covering its own vocabulary.

/// Rewrite every self-closing tag `<x .../>` as `<x ...></x>`,
/// respecting quoted attribute values. Declarations, comments and
/// closing tags are left alone.
pub fn expand_empty_elements(xml: &str) -> String {
    let b = xml.as_bytes();
    let mut out = String::with_capacity(xml.len() + 256);
    let mut i = 0usize;
    while i < b.len() {
        if b[i] != b'<' || i + 1 >= b.len() || matches!(b[i + 1], b'?' | b'!' | b'/') {
            out.push(b[i] as char);
            i += 1;
            continue;
        }
        let start = i;
        let mut j = i + 1;
        let mut quote: Option<u8> = None;
        while j < b.len() {
            match (quote, b[j]) {
                (Some(q), c) if c == q => quote = None,
                (None, c @ (b'"' | b'\'')) => quote = Some(c),
                (None, b'>') => break,
                _ => {}
            }
            j += 1;
        }
        if j >= b.len() {
            out.push_str(&xml[start..]);
            break;
        }
        let tag = &xml[start..=j];
        if let Some(inner) = tag.strip_prefix('<').and_then(|t| t.strip_suffix("/>")) {
            let name_end = inner
                .find(|c: char| c.is_whitespace())
                .unwrap_or(inner.len());
            out.push('<');
            out.push_str(inner.trim_end());
            out.push('>');
            out.push_str("</");
            out.push_str(&inner[..name_end]);
            out.push('>');
        } else {
            out.push_str(tag);
        }
        i = j + 1;
    }
    out
}

/// Remove indentation between elements: whitespace runs that contain a
/// newline and sit between a `>` and a `<`. Content is never touched,
/// so a fixture whose text-bearing elements hold no newline can be
/// collapsed without changing what it says.
pub fn collapse_indentation(xml: &str) -> String {
    let b = xml.as_bytes();
    let mut out = String::with_capacity(xml.len());
    let mut i = 0usize;
    while i < b.len() {
        if b[i] == b'>' {
            let mut j = i + 1;
            while j < b.len() && (b[j] as char).is_whitespace() {
                j += 1;
            }
            if j < b.len() && b[j] == b'<' && xml[i + 1..j].contains('\n') {
                out.push('>');
                i = j;
                continue;
            }
        }
        out.push(b[i] as char);
        i += 1;
    }
    out
}
