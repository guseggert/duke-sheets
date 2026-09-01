use std::io::{BufReader, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};

pub(crate) use duke_sheets_core::style::ThemePalette;

fn parse_ooxml_hex(hex: &str) -> Option<(u8, u8, u8)> {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 {
        return None;
    }
    let hex = if hex.len() == 8 { &hex[2..] } else { hex };
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some((r, g, b))
}

pub(super) fn read_theme_palette<R: Read>(mut reader: R) -> XlsxResult<(ThemePalette, Vec<u8>)> {
    let mut raw_bytes = Vec::new();
    reader.read_to_end(&mut raw_bytes)?;
    let palette = parse_theme_palette(std::io::Cursor::new(&raw_bytes))?;
    Ok((palette, raw_bytes))
}

pub(super) fn parse_theme_palette<R: Read>(reader: R) -> XlsxResult<ThemePalette> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(true);

    let mut palette = ThemePalette::default();
    let mut buf = Vec::new();
    let mut in_clr_scheme = false;
    let mut current_slot: Option<usize> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"clrScheme" => in_clr_scheme = true,
                b"lt1" if in_clr_scheme => current_slot = Some(0),
                b"dk1" if in_clr_scheme => current_slot = Some(1),
                b"lt2" if in_clr_scheme => current_slot = Some(2),
                b"dk2" if in_clr_scheme => current_slot = Some(3),
                b"accent1" if in_clr_scheme => current_slot = Some(4),
                b"accent2" if in_clr_scheme => current_slot = Some(5),
                b"accent3" if in_clr_scheme => current_slot = Some(6),
                b"accent4" if in_clr_scheme => current_slot = Some(7),
                b"accent5" if in_clr_scheme => current_slot = Some(8),
                b"accent6" if in_clr_scheme => current_slot = Some(9),
                b"hlink" if in_clr_scheme => current_slot = Some(10),
                b"folHlink" if in_clr_scheme => current_slot = Some(11),
                b"srgbClr" if in_clr_scheme => {
                    if let Some(slot) = current_slot {
                        if let Some(rgb) = extract_theme_rgb_from_attrs(&e) {
                            palette.colors[slot] = rgb;
                        }
                    }
                }
                b"sysClr" if in_clr_scheme => {
                    if let Some(slot) = current_slot {
                        if let Some(rgb) = extract_theme_rgb_from_attrs(&e) {
                            palette.colors[slot] = rgb;
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"srgbClr" if in_clr_scheme => {
                    if let Some(slot) = current_slot {
                        if let Some(rgb) = extract_theme_rgb_from_attrs(&e) {
                            palette.colors[slot] = rgb;
                        }
                    }
                }
                b"sysClr" if in_clr_scheme => {
                    if let Some(slot) = current_slot {
                        if let Some(rgb) = extract_theme_rgb_from_attrs(&e) {
                            palette.colors[slot] = rgb;
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"clrScheme" => {
                    in_clr_scheme = false;
                    current_slot = None;
                }
                b"lt1" | b"dk1" | b"lt2" | b"dk2" | b"accent1" | b"accent2" | b"accent3"
                | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                    current_slot = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(palette)
}

fn extract_theme_rgb_from_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<(u8, u8, u8)> {
    let mut val: Option<String> = None;
    let mut last_clr: Option<String> = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"val" => val = attr.unescape_value().ok().map(|s| s.to_string()),
            b"lastClr" => last_clr = attr.unescape_value().ok().map(|s| s.to_string()),
            _ => {}
        }
    }

    if let Some(v) = val {
        if let Some(rgb) = parse_ooxml_hex(&v) {
            return Some(rgb);
        }
    }
    if let Some(v) = last_clr {
        if let Some(rgb) = parse_ooxml_hex(&v) {
            return Some(rgb);
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    #[test]
    fn test_parse_theme_palette_custom_accent() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="101010"/></a:dk1>
      <a:lt1><a:srgbClr val="F0F0F0"/></a:lt1>
      <a:dk2><a:srgbClr val="202020"/></a:dk2>
      <a:lt2><a:srgbClr val="E0E0E0"/></a:lt2>
      <a:accent1><a:srgbClr val="112233"/></a:accent1>
      <a:accent2><a:srgbClr val="445566"/></a:accent2>
      <a:accent3><a:srgbClr val="778899"/></a:accent3>
      <a:accent4><a:srgbClr val="AABBCC"/></a:accent4>
      <a:accent5><a:srgbClr val="DDEEFF"/></a:accent5>
      <a:accent6><a:srgbClr val="334455"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

        let palette = parse_theme_palette(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(palette.colors[4], (0x11, 0x22, 0x33));
        assert_eq!(palette.resolve_theme(4, 0.0), (0x11, 0x22, 0x33));
    }
}
