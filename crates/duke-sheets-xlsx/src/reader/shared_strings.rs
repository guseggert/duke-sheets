//! Shared string table reader with rich text run support.

use std::io::{BufReader, Read};

use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::{Color, FontVerticalAlign, Underline};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use super::conditional_format::parse_color_element;
use super::decode_excel_escapes;
use crate::error::{XlsxError, XlsxResult};

/// Entry in the shared string table - either plain text or rich text with runs.
#[derive(Debug, Clone)]
pub(crate) enum SharedStringEntry {
    Plain(String),
    Rich(Vec<RichTextRun>),
}

/// Each `<si>` element is either plain text (`<t>`) or rich text (`<r>` runs).
/// Rich text runs preserve per-character formatting (bold, italic, color, etc.)
/// via `<rPr>` elements within each `<r>`.
pub(crate) fn parse_shared_strings<R: Read>(reader: R) -> XlsxResult<Vec<SharedStringEntry>> {
    let mut entries = Vec::new();
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(false);

    let mut buf = Vec::new();

    // State for plain text mode
    let mut in_si = false;
    let mut in_t = false; // <t> directly under <si> (plain text mode)
    let mut plain_text = String::new();

    // State for rich text mode
    let mut in_r = false; // inside <r> element
    let mut in_rpr = false; // inside <rPr> element
    let mut in_run_t = false; // inside <t> within <r>
    let mut has_runs = false; // current <si> has <r> children
    let mut runs: Vec<RichTextRun> = Vec::new();
    let mut run_text = String::new();
    let mut run_font: Option<RunFont> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    plain_text.clear();
                    runs.clear();
                    has_runs = false;
                }
                b"t" if in_si && !in_r => {
                    // Plain text <t> directly under <si>
                    in_t = true;
                }
                b"r" if in_si => {
                    // Rich text run
                    in_r = true;
                    has_runs = true;
                    run_text.clear();
                    run_font = None;
                }
                b"rPr" if in_r => {
                    in_rpr = true;
                    run_font = Some(RunFont::default());
                }
                b"t" if in_r => {
                    // Text within a run
                    in_run_t = true;
                }
                // <rPr> children that use Start+End (rare, but handle defensively)
                name if in_rpr => parse_rpr_element(name, &e, &mut run_font),
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"si" => {
                    if has_runs {
                        // Apply escape decoding to each run's text
                        for run in &mut runs {
                            run.text = decode_excel_escapes(&run.text);
                        }
                        entries.push(SharedStringEntry::Rich(std::mem::take(&mut runs)));
                    } else {
                        let decoded = decode_excel_escapes(&plain_text);
                        entries.push(SharedStringEntry::Plain(decoded));
                    }
                    plain_text.clear();
                    in_si = false;
                }
                b"t" if in_run_t => in_run_t = false,
                b"t" if in_t => in_t = false,
                b"r" if in_r => {
                    // Finish current run
                    let font = run_font
                        .take()
                        .and_then(|f| if f.is_empty() { None } else { Some(f) });
                    runs.push(RichTextRun {
                        text: std::mem::take(&mut run_text),
                        font,
                    });
                    in_r = false;
                }
                b"rPr" => in_rpr = false,
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                let local = e.name().local_name();
                let name = local.as_ref();
                if in_rpr {
                    parse_rpr_element(name, &e, &mut run_font);
                } else if name == b"t" && in_si && !in_r {
                    // Empty <t/> means empty string - valid
                } else if name == b"t" && in_r {
                    // Empty <t/> within a run
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    match e.unescape() {
                        Ok(text) => plain_text.push_str(&text),
                        Err(err) => log::warn!(
                            "Shared string {}: XML unescape failed: {}",
                            entries.len(),
                            err
                        ),
                    }
                } else if in_run_t {
                    match e.unescape() {
                        Ok(text) => run_text.push_str(&text),
                        Err(err) => log::warn!(
                            "Shared string {}: run text unescape failed: {}",
                            entries.len(),
                            err
                        ),
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(entries)
}

/// Parse a single `<rPr>` child element and update the run font.
pub(crate) fn parse_rpr_element(name: &[u8], e: &BytesStart, font: &mut Option<RunFont>) {
    let f = match font.as_mut() {
        Some(f) => f,
        None => return,
    };

    match name {
        b"b" => f.bold = Some(parse_bool_attr(e)),
        b"i" => f.italic = Some(parse_bool_attr(e)),
        b"strike" => f.strikethrough = Some(parse_bool_attr(e)),
        b"u" => f.underline = Some(parse_underline_attr(e)),
        b"sz" => f.size = parse_val_f64(e),
        b"rFont" => f.name = parse_val_string(e),
        b"family" => f.family = parse_val_u8(e),
        b"charset" => f.charset = parse_val_u8(e),
        b"scheme" => f.scheme = parse_val_string(e),
        b"color" => {
            let color = parse_color_element(e);
            if color != Color::Auto {
                f.color = Some(color);
            }
        }
        b"vertAlign" => {
            f.vertical_align = parse_val_string(e).and_then(|s| match s.as_str() {
                "superscript" => Some(FontVerticalAlign::Superscript),
                "subscript" => Some(FontVerticalAlign::Subscript),
                "baseline" => Some(FontVerticalAlign::Baseline),
                _ => None,
            });
        }
        _ => {}
    }
}

fn parse_bool_attr(e: &BytesStart) -> bool {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return !matches!(
                attr.unescape_value().ok().as_deref(),
                Some("0") | Some("false")
            );
        }
    }
    true
}

/// Parse underline style from `<u/>` or `<u val="double"/>`.
fn parse_underline_attr(e: &BytesStart) -> Underline {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return match attr.unescape_value().ok().as_deref() {
                Some("none") => Underline::None,
                Some("double") => Underline::Double,
                Some("singleAccounting") => Underline::SingleAccounting,
                Some("doubleAccounting") => Underline::DoubleAccounting,
                _ => Underline::Single,
            };
        }
    }
    Underline::Single // Bare <u/> means single underline
}

/// Parse `val` attribute as f64.
fn parse_val_f64(e: &BytesStart) -> Option<f64> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return attr.unescape_value().ok().and_then(|s| s.parse().ok());
        }
    }
    None
}

/// Parse `val` attribute as String.
fn parse_val_string(e: &BytesStart) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return attr.unescape_value().ok().map(|s| s.to_string());
        }
    }
    None
}

/// Parse `val` attribute as u8.
fn parse_val_u8(e: &BytesStart) -> Option<u8> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return attr.unescape_value().ok().and_then(|s| s.parse().ok());
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn plain_text_entries() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        assert_eq!(entries.len(), 2);
        assert!(matches!(&entries[0], SharedStringEntry::Plain(s) if s == "Hello"));
        assert!(matches!(&entries[1], SharedStringEntry::Plain(s) if s == "World"));
    }

    #[test]
    fn rich_text_entries() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t xml:space="preserve">Hello </t></r>
    <r>
      <rPr>
        <b/>
        <sz val="14"/>
        <color rgb="FFFF0000"/>
        <rFont val="Arial"/>
        <family val="2"/>
      </rPr>
      <t>World</t>
    </r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        assert_eq!(entries.len(), 1);
        match &entries[0] {
            SharedStringEntry::Rich(runs) => {
                assert_eq!(runs.len(), 2);
                // First run: plain text, no formatting
                assert_eq!(runs[0].text, "Hello ");
                assert!(runs[0].font.is_none());
                // Second run: bold, 14pt, red, Arial
                assert_eq!(runs[1].text, "World");
                let font = runs[1].font.as_ref().unwrap();
                assert_eq!(font.bold, Some(true));
                assert_eq!(font.size, Some(14.0));
                assert_eq!(font.name, Some("Arial".to_string()));
                assert_eq!(font.family, Some(2));
                assert!(matches!(
                    font.color,
                    Some(Color::Rgb { r: 255, g: 0, b: 0 })
                ));
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }

    #[test]
    fn mixed_plain_and_rich() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Plain text</t></si>
  <si>
    <r><rPr><i/></rPr><t>Italic</t></r>
    <r><t> normal</t></r>
  </si>
  <si><t>Another plain</t></si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        assert_eq!(entries.len(), 3);
        assert!(matches!(&entries[0], SharedStringEntry::Plain(s) if s == "Plain text"));
        assert!(matches!(&entries[1], SharedStringEntry::Rich(runs) if runs.len() == 2));
        assert!(matches!(&entries[2], SharedStringEntry::Plain(s) if s == "Another plain"));

        if let SharedStringEntry::Rich(runs) = &entries[1] {
            assert_eq!(runs[0].text, "Italic");
            assert_eq!(runs[0].font.as_ref().unwrap().italic, Some(true));
            assert_eq!(runs[1].text, " normal");
            assert!(runs[1].font.is_none());
        }
    }

    #[test]
    fn underline_and_strikethrough() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><u val="double"/><strike/></rPr>
      <t>styled</t>
    </r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        if let SharedStringEntry::Rich(runs) = &entries[0] {
            let font = runs[0].font.as_ref().unwrap();
            assert_eq!(font.underline, Some(Underline::Double));
            assert_eq!(font.strikethrough, Some(true));
        } else {
            panic!("Expected Rich");
        }
    }

    #[test]
    fn theme_color_in_run() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><color theme="4" tint="-0.25"/><sz val="11"/></rPr>
      <t>themed</t>
    </r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        if let SharedStringEntry::Rich(runs) = &entries[0] {
            let font = runs[0].font.as_ref().unwrap();
            // parse_color_element with no theme palette returns Theme { index, tint }
            assert!(matches!(font.color, Some(Color::Theme { index: 4, .. })));
        } else {
            panic!("Expected Rich");
        }
    }

    #[test]
    fn superscript_and_scheme() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr>
        <vertAlign val="superscript"/>
        <scheme val="minor"/>
      </rPr>
      <t>sup</t>
    </r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        if let SharedStringEntry::Rich(runs) = &entries[0] {
            let font = runs[0].font.as_ref().unwrap();
            assert_eq!(font.vertical_align, Some(FontVerticalAlign::Superscript));
            assert_eq!(font.scheme, Some("minor".to_string()));
        } else {
            panic!("Expected Rich");
        }
    }

    #[test]
    fn empty_sst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        assert!(entries.is_empty());
    }

    #[test]
    fn excel_escape_in_rich_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Hello_x000D_World</t></r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        if let SharedStringEntry::Rich(runs) = &entries[0] {
            assert_eq!(runs[0].text, "Hello\rWorld");
        } else {
            panic!("Expected Rich");
        }
    }

    #[test]
    fn bool_val_false() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr><b val="0"/><i val="1"/></rPr>
      <t>mixed</t>
    </r>
  </si>
</sst>"#;
        let entries = parse_shared_strings(Cursor::new(xml)).unwrap();
        if let SharedStringEntry::Rich(runs) = &entries[0] {
            let font = runs[0].font.as_ref().unwrap();
            assert_eq!(font.bold, Some(false));
            assert_eq!(font.italic, Some(true));
        } else {
            panic!("Expected Rich");
        }
    }
}
