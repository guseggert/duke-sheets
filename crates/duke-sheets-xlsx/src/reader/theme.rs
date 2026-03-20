use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::style::{Color, FillStyle, Style};

#[derive(Debug, Clone, Copy)]
pub(crate) struct ThemePalette {
    pub(crate) colors: [(u8, u8, u8); 12],
}

impl Default for ThemePalette {
    fn default() -> Self {
        Self {
            colors: [
                (255, 255, 255),
                (0, 0, 0),
                (238, 236, 225),
                (31, 73, 125),
                (79, 129, 189),
                (192, 80, 77),
                (155, 187, 89),
                (128, 100, 162),
                (75, 172, 198),
                (247, 150, 70),
                (0, 0, 255),
                (128, 0, 128),
            ],
        }
    }
}

impl ThemePalette {
    pub(crate) fn resolve_theme_color(&self, index: u8, tint: i8) -> (u8, u8, u8) {
        let base = match index {
            0..=9 => self.colors[index as usize],
            _ => (0, 0, 0),
        };
        Self::apply_tint(base, tint)
    }

    fn apply_tint(color: (u8, u8, u8), tint: i8) -> (u8, u8, u8) {
        let tint_float = tint as f64 / 100.0;

        let apply = |c: u8| -> u8 {
            let c = c as f64;
            let result = if tint_float < 0.0 {
                c * (1.0 + tint_float)
            } else {
                c + (255.0 - c) * tint_float
            };
            result.clamp(0.0, 255.0) as u8
        };

        (apply(color.0), apply(color.1), apply(color.2))
    }

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
}

pub(super) fn read_theme_palette<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    theme_path: Option<&str>,
) -> XlsxResult<(Option<ThemePalette>, Option<Vec<u8>>)> {
    let mut try_paths: Vec<String> = Vec::new();
    if let Some(path) = theme_path {
        try_paths.push(path.to_string());
    }
    if !try_paths.iter().any(|p| p == "xl/theme/theme1.xml") {
        try_paths.push("xl/theme/theme1.xml".to_string());
    }

    for path in try_paths {
        let mut file = match archive.by_name(&path) {
            Ok(f) => f,
            Err(_) => continue,
        };
        // Read raw bytes for roundtrip preservation, then parse palette from them.
        let mut raw_bytes = Vec::new();
        file.read_to_end(&mut raw_bytes)?;
        let palette = parse_theme_palette(std::io::Cursor::new(&raw_bytes))?;
        return Ok((Some(palette), Some(raw_bytes)));
    }

    Ok((None, None))
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
                b"srgbClr" | b"sysClr" if in_clr_scheme => {
                    if let Some(slot) = current_slot {
                        if let Some(rgb) = extract_theme_rgb_from_attrs(&e) {
                            palette.colors[slot] = rgb;
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"srgbClr" | b"sysClr" if in_clr_scheme => {
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
        if let Some(rgb) = ThemePalette::parse_ooxml_hex(&v) {
            return Some(rgb);
        }
    }
    if let Some(v) = last_clr {
        if let Some(rgb) = ThemePalette::parse_ooxml_hex(&v) {
            return Some(rgb);
        }
    }
    None
}

pub(super) fn resolve_style_theme_colors(style: &mut Style, theme: &ThemePalette) {
    style.font.color = resolve_color_theme(style.font.color, theme);

    match &mut style.fill {
        FillStyle::None => {}
        FillStyle::Solid { color } => {
            *color = resolve_color_theme(*color, theme);
        }
        FillStyle::Pattern {
            foreground,
            background,
            ..
        } => {
            *foreground = resolve_color_theme(*foreground, theme);
            *background = resolve_color_theme(*background, theme);
        }
        FillStyle::Gradient { stops, .. } => {
            for stop in stops {
                stop.color = resolve_color_theme(stop.color, theme);
            }
        }
    }

    for edge in [
        &mut style.border.left,
        &mut style.border.right,
        &mut style.border.top,
        &mut style.border.bottom,
        &mut style.border.diagonal,
    ].into_iter().flatten() {
        edge.color = resolve_color_theme(edge.color, theme);
    }
}

pub(super) fn resolve_color_theme(color: Color, theme: &ThemePalette) -> Color {
    match color {
        Color::Theme { index, tint } => {
            let (r, g, b) = theme.resolve_theme_color(index, tint);
            Color::Rgb { r, g, b }
        }
        other => other,
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use quick_xml::events::BytesStart;

    use super::*;
    use crate::reader::conditional_format::parse_color_element;

    #[test]
    fn test_parse_color_element_theme_with_palette_resolves_to_rgb() {
        let mut e = BytesStart::new("color");
        e.push_attribute(("theme", "4"));
        e.push_attribute(("tint", "0.5"));

        let palette = ThemePalette::default();
        assert_eq!(
            parse_color_element(&e, Some(&palette)),
            Color::Rgb {
                r: 167,
                g: 192,
                b: 222
            }
        );
    }

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
        assert_eq!(palette.resolve_theme_color(4, 0), (0x11, 0x22, 0x33));
    }
}
