//! XLSX styles (styles.xml) read/write helpers

use std::collections::HashMap;
use std::io::{BufReader, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, FillStyle, FontStyle,
    HorizontalAlignment, NumberFormat, PatternType, Protection, ReadingOrder, Style, Underline,
    VerticalAlignment,
};
use duke_sheets_core::Workbook;

// === Writing ===

#[derive(Debug)]
pub(crate) struct XlsxStyleTable {
    /// Global, deduplicated styles. Index corresponds to the cellXfs index (xfId).
    styles: Vec<Style>,
    /// Per-worksheet mapping: local worksheet style index -> global xfId.
    sheet_maps: Vec<HashMap<u32, u32>>,
    /// DXF (differential format) styles for conditional formatting.
    /// Key: (sheet_index, rule_index), Value: dxf_id
    dxf_styles: Vec<Style>,
    /// Mapping from (sheet_index, cf_rule_index) to dxf_id
    dxf_map: HashMap<(usize, usize), u32>,
}

#[derive(Debug, Clone, Copy)]
struct ResolvedXfIds {
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    num_fmt_id: u32,
}

impl XlsxStyleTable {
    pub(crate) fn build(workbook: &Workbook) -> Self {
        let mut styles: Vec<Style> = Vec::new();
        let mut style_to_xf: HashMap<Style, u32> = HashMap::new();

        // Index 0 is always default style
        let default = Style::default();
        styles.push(default.clone());
        style_to_xf.insert(default, 0);

        let mut sheet_maps: Vec<HashMap<u32, u32>> = Vec::with_capacity(workbook.sheet_count());

        // DXF styles for conditional formatting
        let mut dxf_styles: Vec<Style> = Vec::new();
        let mut dxf_map: HashMap<(usize, usize), u32> = HashMap::new();
        let mut dxf_style_to_id: HashMap<Style, u32> = HashMap::new();

        for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
            let mut map: HashMap<u32, u32> = HashMap::new();
            map.insert(0, 0);

            for (_row, _col, cell) in sheet.iter_cells() {
                let local_idx = cell.style_index;
                if local_idx == 0 || map.contains_key(&local_idx) {
                    continue;
                }

                let style = sheet
                    .style_by_index(local_idx)
                    .cloned()
                    .unwrap_or_else(Style::default);

                let xf_id = match style_to_xf.get(&style) {
                    Some(&id) => id,
                    None => {
                        let id = styles.len() as u32;
                        styles.push(style.clone());
                        style_to_xf.insert(style, id);
                        id
                    }
                };

                map.insert(local_idx, xf_id);
            }

            sheet_maps.push(map);

            // Collect DXF styles from conditional formatting rules
            for (rule_idx, rule) in sheet.conditional_formats().iter().enumerate() {
                if let Some(ref format) = rule.format {
                    // Check if we already have this DXF style
                    let dxf_id = match dxf_style_to_id.get(format) {
                        Some(&id) => id,
                        None => {
                            let id = dxf_styles.len() as u32;
                            dxf_styles.push(format.clone());
                            dxf_style_to_id.insert(format.clone(), id);
                            id
                        }
                    };
                    dxf_map.insert((sheet_idx, rule_idx), dxf_id);
                }
            }
        }

        Self {
            styles,
            sheet_maps,
            dxf_styles,
            dxf_map,
        }
    }

    pub(crate) fn xf_id_for(&self, sheet_index: usize, local_style_index: u32) -> u32 {
        self.sheet_maps
            .get(sheet_index)
            .and_then(|m| m.get(&local_style_index).copied())
            .unwrap_or(0)
    }

    /// Get the DXF ID for a conditional format rule, if it has a format defined
    pub(crate) fn dxf_id_for(&self, sheet_index: usize, rule_index: usize) -> Option<u32> {
        self.dxf_map.get(&(sheet_index, rule_index)).copied()
    }

    /// Get the DXF styles
    pub(crate) fn dxf_styles(&self) -> &[Style] {
        &self.dxf_styles
    }

    pub(crate) fn to_styles_xml(&self) -> String {
        // Build component tables
        let mut font_ids: HashMap<FontStyle, u32> = HashMap::new();
        let mut fonts: Vec<FontStyle> = Vec::new();

        let default_font = FontStyle::default();
        fonts.push(default_font.clone());
        font_ids.insert(default_font, 0);

        let mut fill_ids: HashMap<FillStyle, u32> = HashMap::new();
        let mut fills: Vec<FillStyle> = Vec::new();
        // Excel requires the first two fills to be: none and gray125
        fills.push(FillStyle::None); // id 0
        fills.push(FillStyle::Pattern {
            pattern: PatternType::Gray125,
            foreground: Color::Auto,
            background: Color::Auto,
        }); // id 1
        fill_ids.insert(FillStyle::None, 0);

        let mut border_ids: HashMap<BorderStyle, u32> = HashMap::new();
        let mut borders: Vec<BorderStyle> = Vec::new();
        let default_border = BorderStyle::default();
        borders.push(default_border.clone());
        border_ids.insert(default_border, 0);

        // Custom number formats
        let mut numfmt_ids: HashMap<String, u32> = HashMap::new();
        let mut numfmts: Vec<(u32, String)> = Vec::new();
        let mut next_numfmt_id: u32 = 164;

        // Resolve component IDs for each style
        let mut resolved: Vec<ResolvedXfIds> = Vec::with_capacity(self.styles.len());

        for style in &self.styles {
            // Font
            let font_id = match font_ids.get(&style.font) {
                Some(&id) => id,
                None => {
                    let id = fonts.len() as u32;
                    fonts.push(style.font.clone());
                    font_ids.insert(style.font.clone(), id);
                    id
                }
            };

            // Fill (normalize gradients to solid first-stop or none)
            let norm_fill = match &style.fill {
                FillStyle::Gradient { stops, .. } => stops
                    .first()
                    .map(|s| FillStyle::Solid { color: s.color })
                    .unwrap_or(FillStyle::None),
                other => other.clone(),
            };

            let fill_id = match norm_fill {
                FillStyle::None => 0,
                other => {
                    if let Some(&id) = fill_ids.get(&other) {
                        id
                    } else {
                        let id = fills.len() as u32;
                        fills.push(other.clone());
                        fill_ids.insert(other, id);
                        id
                    }
                }
            };

            // Border
            let border_id = match border_ids.get(&style.border) {
                Some(&id) => id,
                None => {
                    let id = borders.len() as u32;
                    borders.push(style.border.clone());
                    border_ids.insert(style.border.clone(), id);
                    id
                }
            };

            // Number format
            let num_fmt_id = match &style.number_format {
                NumberFormat::General => 0,
                NumberFormat::BuiltIn(id) => *id,
                NumberFormat::Custom(code) => {
                    if let Some(&id) = numfmt_ids.get(code) {
                        id
                    } else {
                        let id = next_numfmt_id;
                        next_numfmt_id += 1;
                        numfmt_ids.insert(code.clone(), id);
                        numfmts.push((id, code.clone()));
                        id
                    }
                }
            };

            resolved.push(ResolvedXfIds {
                font_id,
                fill_id,
                border_id,
                num_fmt_id,
            });
        }

        // Write XML
        let mut xml = String::new();
        xml.push_str(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        if !numfmts.is_empty() {
            xml.push_str(&format!("\n  <numFmts count=\"{}\">", numfmts.len()));
            for (id, code) in &numfmts {
                xml.push_str(&format!(
                    "\n    <numFmt numFmtId=\"{}\" formatCode=\"{}\"/>",
                    id,
                    escape_xml_attr(code)
                ));
            }
            xml.push_str("\n  </numFmts>");
        }

        // Fonts
        xml.push_str(&format!("\n  <fonts count=\"{}\">", fonts.len()));
        for font in &fonts {
            xml.push_str("\n    ");
            xml.push_str(&write_font(font));
        }
        xml.push_str("\n  </fonts>");

        // Fills
        xml.push_str(&format!("\n  <fills count=\"{}\">", fills.len()));
        for fill in &fills {
            xml.push_str("\n    ");
            xml.push_str(&write_fill(fill));
        }
        xml.push_str("\n  </fills>");

        // Borders
        xml.push_str(&format!("\n  <borders count=\"{}\">", borders.len()));
        for border in &borders {
            xml.push_str("\n    ");
            xml.push_str(&write_border(border));
        }
        xml.push_str("\n  </borders>");

        // cellStyleXfs (required)
        xml.push_str(
            r#"
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>"#,
        );

        // cellXfs
        xml.push_str(&format!("\n  <cellXfs count=\"{}\">", self.styles.len()));
        for (i, ids) in resolved.iter().enumerate() {
            let style = &self.styles[i];
            xml.push_str("\n    ");
            xml.push_str(&write_xf(style, *ids));
        }
        xml.push_str("\n  </cellXfs>");

        // cellStyles (required)
        xml.push_str(
            r#"
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>"#,
        );

        // DXFs (differential formats for conditional formatting)
        if self.dxf_styles.is_empty() {
            xml.push_str("\n  <dxfs count=\"0\"/>");
        } else {
            xml.push_str(&format!("\n  <dxfs count=\"{}\">", self.dxf_styles.len()));
            for dxf_style in &self.dxf_styles {
                xml.push_str("\n    ");
                xml.push_str(&write_dxf(dxf_style));
            }
            xml.push_str("\n  </dxfs>");
        }

        xml.push_str(
            r#"
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>"#,
        );

        xml.push_str("\n</styleSheet>");
        xml
    }
}

fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn write_color(tag: &str, color: &Color) -> String {
    match color {
        Color::Auto => format!("<{tag} indexed=\"64\"/>"),
        Color::Rgb { r, g, b } => format!("<{tag} rgb=\"FF{:02X}{:02X}{:02X}\"/>", r, g, b),
        Color::Argb { a, r, g, b } => {
            format!("<{tag} rgb=\"{:02X}{:02X}{:02X}{:02X}\"/>", a, r, g, b)
        }
        Color::Indexed(i) => format!("<{tag} indexed=\"{}\"/>", i),
        Color::Theme { index, tint } => {
            let tint_f = (*tint as f64) / 100.0;
            if *tint == 0 {
                format!("<{tag} theme=\"{}\"/>", index)
            } else {
                format!("<{tag} theme=\"{}\" tint=\"{}\"/>", index, tint_f)
            }
        }
    }
}

fn write_font(font: &FontStyle) -> String {
    let mut s = String::from("<font>");
    if font.bold {
        s.push_str("<b/>");
    }
    if font.italic {
        s.push_str("<i/>");
    }
    if font.strikethrough {
        s.push_str("<strike/>");
    }
    match font.underline {
        Underline::None => {}
        Underline::Single => s.push_str("<u/>"),
        Underline::Double => s.push_str("<u val=\"double\"/>"),
        Underline::SingleAccounting => s.push_str("<u val=\"singleAccounting\"/>"),
        Underline::DoubleAccounting => s.push_str("<u val=\"doubleAccounting\"/>"),
    }
    s.push_str(&format!("<sz val=\"{}\"/>", font.size));

    if !matches!(font.color, Color::Auto) {
        // Use <color .../>
        // The OOXML tag name is always "color" in <font>
        let mut color = String::from("<color");
        match &font.color {
            Color::Auto => {
                color.push_str(" indexed=\"64\"");
            }
            Color::Rgb { r, g, b } => {
                color.push_str(&format!(" rgb=\"FF{:02X}{:02X}{:02X}\"", r, g, b));
            }
            Color::Argb { a, r, g, b } => {
                color.push_str(&format!(" rgb=\"{:02X}{:02X}{:02X}{:02X}\"", a, r, g, b));
            }
            Color::Indexed(i) => {
                color.push_str(&format!(" indexed=\"{}\"", i));
            }
            Color::Theme { index, tint } => {
                color.push_str(&format!(" theme=\"{}\"", index));
                if *tint != 0 {
                    color.push_str(&format!(" tint=\"{}\"", (*tint as f64) / 100.0));
                }
            }
        }
        color.push_str("/>");
        s.push_str(&color);
    }

    s.push_str(&format!("<name val=\"{}\"/>", escape_xml_attr(&font.name)));
    s.push_str("</font>");
    s
}

fn pattern_type_to_str(p: PatternType) -> &'static str {
    match p {
        PatternType::None => "none",
        PatternType::Solid => "solid",
        PatternType::MediumGray => "mediumGray",
        PatternType::DarkGray => "darkGray",
        PatternType::LightGray => "lightGray",
        PatternType::DarkHorizontal => "darkHorizontal",
        PatternType::DarkVertical => "darkVertical",
        PatternType::DarkDown => "darkDown",
        PatternType::DarkUp => "darkUp",
        PatternType::DarkGrid => "darkGrid",
        PatternType::DarkTrellis => "darkTrellis",
        PatternType::LightHorizontal => "lightHorizontal",
        PatternType::LightVertical => "lightVertical",
        PatternType::LightDown => "lightDown",
        PatternType::LightUp => "lightUp",
        PatternType::LightGrid => "lightGrid",
        PatternType::LightTrellis => "lightTrellis",
        PatternType::Gray125 => "gray125",
        PatternType::Gray0625 => "gray0625",
    }
}

fn write_fill(fill: &FillStyle) -> String {
    match fill {
        FillStyle::None => "<fill><patternFill patternType=\"none\"/></fill>".to_string(),
        FillStyle::Solid { color } => {
            format!(
                "<fill><patternFill patternType=\"solid\">{}<bgColor indexed=\"64\"/></patternFill></fill>",
                write_color("fgColor", color)
            )
        }
        FillStyle::Pattern {
            pattern,
            foreground,
            background,
        } => {
            format!(
                "<fill><patternFill patternType=\"{}\">{}{}{}</patternFill></fill>",
                pattern_type_to_str(*pattern),
                write_color("fgColor", foreground),
                "",
                write_color("bgColor", background)
            )
        }
        FillStyle::Gradient { .. } => {
            // Gradients are currently downgraded when building the table; keep a safe fallback.
            "<fill><patternFill patternType=\"none\"/></fill>".to_string()
        }
    }
}

fn border_style_to_str(s: BorderLineStyle) -> Option<&'static str> {
    match s {
        BorderLineStyle::None => None,
        BorderLineStyle::Thin => Some("thin"),
        BorderLineStyle::Medium => Some("medium"),
        BorderLineStyle::Thick => Some("thick"),
        BorderLineStyle::Dashed => Some("dashed"),
        BorderLineStyle::Dotted => Some("dotted"),
        BorderLineStyle::Double => Some("double"),
        BorderLineStyle::Hair => Some("hair"),
        BorderLineStyle::MediumDashed => Some("mediumDashed"),
        BorderLineStyle::DashDot => Some("dashDot"),
        BorderLineStyle::MediumDashDot => Some("mediumDashDot"),
        BorderLineStyle::DashDotDot => Some("dashDotDot"),
        BorderLineStyle::MediumDashDotDot => Some("mediumDashDotDot"),
        BorderLineStyle::SlantDashDot => Some("slantDashDot"),
    }
}

fn write_border_edge(tag: &str, edge: &Option<BorderEdge>) -> String {
    match edge {
        None => format!("<{tag}/>"),
        Some(e) => {
            let style_attr = border_style_to_str(e.style);
            if style_attr.is_none() {
                return format!("<{tag}/>");
            }
            let mut s = format!("<{tag} style=\"{}\">", style_attr.unwrap());
            // <color .../>
            let mut color = String::from("<color");
            match &e.color {
                Color::Auto => {
                    color.push_str(" indexed=\"64\"");
                }
                Color::Rgb { r, g, b } => {
                    color.push_str(&format!(" rgb=\"FF{:02X}{:02X}{:02X}\"", r, g, b));
                }
                Color::Argb { a, r, g, b } => {
                    color.push_str(&format!(" rgb=\"{:02X}{:02X}{:02X}{:02X}\"", a, r, g, b));
                }
                Color::Indexed(i) => {
                    color.push_str(&format!(" indexed=\"{}\"", i));
                }
                Color::Theme { index, tint } => {
                    color.push_str(&format!(" theme=\"{}\"", index));
                    if *tint != 0 {
                        color.push_str(&format!(" tint=\"{}\"", (*tint as f64) / 100.0));
                    }
                }
            }
            color.push_str("/>");
            s.push_str(&color);
            s.push_str(&format!("</{tag}>",));
            s
        }
    }
}

fn write_border(border: &BorderStyle) -> String {
    let mut attrs = String::new();
    match border.diagonal_direction {
        duke_sheets_core::style::DiagonalDirection::None => {}
        duke_sheets_core::style::DiagonalDirection::Down => attrs.push_str(" diagonalDown=\"1\""),
        duke_sheets_core::style::DiagonalDirection::Up => attrs.push_str(" diagonalUp=\"1\""),
        duke_sheets_core::style::DiagonalDirection::Both => {
            attrs.push_str(" diagonalDown=\"1\" diagonalUp=\"1\"")
        }
    }

    let mut s = format!("<border{}>", attrs);
    s.push_str(&write_border_edge("left", &border.left));
    s.push_str(&write_border_edge("right", &border.right));
    s.push_str(&write_border_edge("top", &border.top));
    s.push_str(&write_border_edge("bottom", &border.bottom));
    s.push_str(&write_border_edge("diagonal", &border.diagonal));
    s.push_str("</border>");
    s
}

fn horiz_to_str(h: HorizontalAlignment) -> &'static str {
    match h {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterContinuous => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    }
}

fn vert_to_str(v: VerticalAlignment) -> &'static str {
    match v {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

fn write_alignment(al: &Alignment) -> String {
    // Only write if any non-default property is set
    let default = Alignment::default();
    if al == &default {
        return String::new();
    }

    let mut s = String::from("<alignment");
    if al.horizontal != default.horizontal {
        s.push_str(&format!(" horizontal=\"{}\"", horiz_to_str(al.horizontal)));
    }
    if al.vertical != default.vertical {
        s.push_str(&format!(" vertical=\"{}\"", vert_to_str(al.vertical)));
    }
    if al.wrap_text {
        s.push_str(" wrapText=\"1\"");
    }
    if al.shrink_to_fit {
        s.push_str(" shrinkToFit=\"1\"");
    }
    if al.indent != 0 {
        s.push_str(&format!(" indent=\"{}\"", al.indent));
    }
    if al.rotation != 0 {
        s.push_str(&format!(" textRotation=\"{}\"", al.rotation));
    }
    match al.reading_order {
        ReadingOrder::ContextDependent => {}
        ReadingOrder::LeftToRight => s.push_str(" readingOrder=\"1\""),
        ReadingOrder::RightToLeft => s.push_str(" readingOrder=\"2\""),
    }
    s.push_str("/>");
    s
}

fn write_protection(p: &Protection) -> String {
    let default = Protection::default();
    if p == &default {
        return String::new();
    }
    let mut s = String::from("<protection");
    if p.locked != default.locked {
        s.push_str(&format!(" locked=\"{}\"", if p.locked { 1 } else { 0 }));
    }
    if p.hidden != default.hidden {
        s.push_str(&format!(" hidden=\"{}\"", if p.hidden { 1 } else { 0 }));
    }
    s.push_str("/>");
    s
}

fn write_xf(style: &Style, ids: ResolvedXfIds) -> String {
    // apply flags
    let mut attrs = String::new();
    if ids.num_fmt_id != 0 {
        attrs.push_str(" applyNumberFormat=\"1\"");
    }
    if style.font != FontStyle::default() {
        attrs.push_str(" applyFont=\"1\"");
    }
    if style.fill != FillStyle::None {
        attrs.push_str(" applyFill=\"1\"");
    }
    if style.border != BorderStyle::default() {
        attrs.push_str(" applyBorder=\"1\"");
    }
    if style.alignment != Alignment::default() {
        attrs.push_str(" applyAlignment=\"1\"");
    }
    if style.protection != Protection::default() {
        attrs.push_str(" applyProtection=\"1\"");
    }

    let mut s = format!(
        "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\"{}",
        ids.num_fmt_id, ids.font_id, ids.fill_id, ids.border_id, attrs
    );

    let alignment_xml = write_alignment(&style.alignment);
    let protection_xml = write_protection(&style.protection);
    if alignment_xml.is_empty() && protection_xml.is_empty() {
        s.push_str("/>");
        return s;
    }

    s.push('>');
    if !alignment_xml.is_empty() {
        s.push_str(&alignment_xml);
    }
    if !protection_xml.is_empty() {
        s.push_str(&protection_xml);
    }
    s.push_str("</xf>");
    s
}

/// Write a DXF (differential format) element for conditional formatting
fn write_dxf(style: &Style) -> String {
    let mut s = String::from("<dxf>");

    // Font (only if non-default)
    if style.font != FontStyle::default() {
        s.push_str(&write_font(&style.font));
    }

    // Fill (only if non-default)
    if style.fill != FillStyle::None {
        s.push_str(&write_fill(&style.fill));
    }

    // Border (only if non-default)
    if style.border != BorderStyle::default() {
        s.push_str(&write_border(&style.border));
    }

    s.push_str("</dxf>");
    s
}

// === Reading ===

/// Result of reading styles.xml, containing both cell styles and DXF styles
#[derive(Debug)]
pub(crate) struct ParsedStyles {
    pub cell_styles: Vec<Style>,
    pub dxf_styles: Vec<Style>,
}

pub(crate) fn read_styles_xml<R: Read>(reader: R) -> XlsxResult<ParsedStyles> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.trim_text(true);

    let mut buf = Vec::new();

    let mut numfmts: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<FontStyle> = Vec::new();
    let mut fills: Vec<FillStyle> = Vec::new();
    let mut borders: Vec<BorderStyle> = Vec::new();
    let mut cell_xfs: Vec<Style> = Vec::new();
    let mut dxf_styles: Vec<Style> = Vec::new();

    // Current objects while parsing
    let mut current_font: Option<FontStyle> = None;
    let mut current_fill_pattern: Option<PatternType> = None;
    let mut current_fill_fg: Color = Color::Auto;
    let mut current_fill_bg: Color = Color::Auto;
    let mut in_fill = false;

    let mut current_border: Option<BorderStyle> = None;
    let mut current_border_edge: Option<&'static str> = None;

    // Current xf
    let mut current_xf: Option<(u32, u32, u32, u32, Alignment, Protection)> = None;
    let mut in_cell_xfs = false;

    // DXF parsing state
    let mut in_dxfs = false;
    let mut in_dxf = false;
    let mut current_dxf: Option<Style> = None;
    let mut dxf_font: Option<FontStyle> = None;
    let mut dxf_fill_pattern: Option<PatternType> = None;
    let mut dxf_fill_fg: Color = Color::Auto;
    let mut dxf_fill_bg: Color = Color::Auto;
    let mut in_dxf_fill = false;
    let mut dxf_border: Option<BorderStyle> = None;
    let mut dxf_border_edge: Option<&'static str> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"numFmts" | b"fonts" | b"fills" | b"borders" => {}

                b"cellXfs" => {
                    in_cell_xfs = true;
                }

                b"dxfs" => {
                    in_dxfs = true;
                }

                b"dxf" if in_dxfs => {
                    in_dxf = true;
                    current_dxf = Some(Style::default());
                    dxf_font = None;
                    dxf_fill_pattern = None;
                    dxf_fill_fg = Color::Auto;
                    dxf_fill_bg = Color::Auto;
                    in_dxf_fill = false;
                    dxf_border = None;
                    dxf_border_edge = None;
                }

                b"font" if in_dxf => {
                    dxf_font = Some(FontStyle::default());
                }

                b"fill" if in_dxf => {
                    in_dxf_fill = true;
                    dxf_fill_pattern = None;
                    dxf_fill_fg = Color::Auto;
                    dxf_fill_bg = Color::Auto;
                }

                b"patternFill" if in_dxf_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"patternType" {
                            if let Ok(v) = attr.unescape_value() {
                                dxf_fill_pattern = str_to_pattern_type(&v);
                            }
                        }
                    }
                }

                b"border" if in_dxf => {
                    let mut b = BorderStyle::default();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"diagonalUp" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Up;
                                }
                            }
                            b"diagonalDown" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Down;
                                }
                            }
                            _ => {}
                        }
                    }
                    dxf_border = Some(b);
                }

                b"font" => {
                    current_font = Some(FontStyle::default());
                }

                b"fill" => {
                    in_fill = true;
                    current_fill_pattern = None;
                    current_fill_fg = Color::Auto;
                    current_fill_bg = Color::Auto;
                }

                b"patternFill" if in_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"patternType" {
                            if let Ok(v) = attr.unescape_value() {
                                current_fill_pattern = str_to_pattern_type(&v);
                            }
                        }
                    }
                }

                b"border" => {
                    let mut b = BorderStyle::default();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"diagonalUp" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Up;
                                }
                            }
                            b"diagonalDown" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Down;
                                }
                            }
                            _ => {}
                        }
                    }
                    current_border = Some(b);
                }

                // Border edges
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                    if current_border.is_some() {
                        current_border_edge = Some(match e.name().as_ref() {
                            b"left" => "left",
                            b"right" => "right",
                            b"top" => "top",
                            b"bottom" => "bottom",
                            _ => "diagonal",
                        });

                        // Parse style attribute
                        if let Some(border) = current_border.as_mut() {
                            let mut style: Option<BorderLineStyle> = None;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"style" {
                                    if let Ok(v) = attr.unescape_value() {
                                        style = str_to_border_style(&v);
                                    }
                                }
                            }
                            // Create edge with default color; color may be overwritten by nested <color>
                            if let Some(st) = style {
                                if st != BorderLineStyle::None {
                                    set_border_edge(
                                        border,
                                        current_border_edge.unwrap(),
                                        Some(BorderEdge {
                                            style: st,
                                            color: Color::Auto,
                                        }),
                                    );
                                }
                            }
                        }
                    }
                }

                b"xf" if in_cell_xfs => {
                    // Parse ids
                    let mut num_fmt_id = 0u32;
                    let mut font_id = 0u32;
                    let mut fill_id = 0u32;
                    let mut border_id = 0u32;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"fontId" => {
                                font_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"fillId" => {
                                fill_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"borderId" => {
                                border_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            _ => {}
                        }
                    }
                    current_xf = Some((
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        Alignment::default(),
                        Protection::default(),
                    ));
                }

                b"alignment" => {
                    if let Some((_n, _f, _fi, _b, align, _p)) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    if let Some(h) = str_to_horizontal(&val) {
                                        align.horizontal = h;
                                    }
                                }
                                b"vertical" => {
                                    if let Some(v) = str_to_vertical(&val) {
                                        align.vertical = v;
                                    }
                                }
                                b"wrapText" => {
                                    align.wrap_text = val.as_ref() == "1";
                                }
                                b"shrinkToFit" => {
                                    align.shrink_to_fit = val.as_ref() == "1";
                                }
                                b"indent" => {
                                    align.indent = val.parse::<u8>().unwrap_or(0);
                                }
                                b"textRotation" => {
                                    align.rotation = val.parse::<i16>().unwrap_or(0);
                                }
                                b"readingOrder" => {
                                    align.reading_order = match val.as_ref() {
                                        "1" => ReadingOrder::LeftToRight,
                                        "2" => ReadingOrder::RightToLeft,
                                        _ => ReadingOrder::ContextDependent,
                                    };
                                }
                                _ => {}
                            }
                        }
                    }
                }

                b"protection" => {
                    if let Some((_n, _f, _fi, _b, _a, prot)) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.as_ref() {
                                b"locked" => prot.locked = val.as_ref() == "1",
                                b"hidden" => prot.hidden = val.as_ref() == "1",
                                _ => {}
                            }
                        }
                    }
                }

                // font sub-elements (handle both regular fonts and DXF fonts)
                b"sz" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.size = v.parse::<f64>().unwrap_or(font.size);
                                }
                            }
                        }
                    }
                }
                b"name" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.name = v.to_string();
                                }
                            }
                        }
                    }
                }
                b"b" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.bold = true;
                    }
                }
                b"i" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.italic = true;
                    }
                }
                b"strike" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.strikethrough = true;
                    }
                }
                b"u" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        let mut underline = Underline::Single;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    underline = str_to_underline(&v);
                                }
                            }
                        }
                        font.underline = underline;
                    }
                }

                b"color" => {
                    // Font color or border color depending on context
                    let color = parse_color_attrs(&e);
                    // Check DXF context first
                    if let Some(font) = dxf_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (dxf_border.as_mut(), dxf_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    } else if let Some(font) = current_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (current_border.as_mut(), current_border_edge)
                    {
                        // Update border edge color if edge exists
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    }
                }

                b"fgColor" => {
                    if in_dxf_fill {
                        dxf_fill_fg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_fg = parse_color_attrs(&e);
                    }
                }
                b"bgColor" => {
                    if in_dxf_fill {
                        dxf_fill_bg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_bg = parse_color_attrs(&e);
                    }
                }

                _ => {}
            },

            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"numFmt" => {
                    let mut id = None;
                    let mut code = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                id = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                            }
                            b"formatCode" => {
                                code = attr.unescape_value().ok().map(|s| s.to_string())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(code)) = (id, code) {
                        numfmts.insert(id, code);
                    }
                }

                // Font empty tags (handle both regular fonts and DXF fonts)
                b"b" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.bold = true;
                    }
                }
                b"i" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.italic = true;
                    }
                }
                b"strike" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.strikethrough = true;
                    }
                }
                b"u" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        let mut underline = Underline::Single;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    underline = str_to_underline(&v);
                                }
                            }
                        }
                        font.underline = underline;
                    }
                }
                b"sz" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.size = v.parse::<f64>().unwrap_or(font.size);
                                }
                            }
                        }
                    }
                }
                b"name" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.name = v.to_string();
                                }
                            }
                        }
                    }
                }
                b"color" => {
                    let color = parse_color_attrs(&e);
                    // Check DXF context first
                    if let Some(font) = dxf_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (dxf_border.as_mut(), dxf_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    } else if let Some(font) = current_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (current_border.as_mut(), current_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    }
                }

                // Fill colors (handle both regular fills and DXF fills)
                b"fgColor" => {
                    if in_dxf_fill {
                        dxf_fill_fg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_fg = parse_color_attrs(&e);
                    }
                }
                b"bgColor" => {
                    if in_dxf_fill {
                        dxf_fill_bg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_bg = parse_color_attrs(&e);
                    }
                }

                // alignment can be self-closing
                b"alignment" => {
                    if let Some((_n, _f, _fi, _b, align, _p)) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    if let Some(h) = str_to_horizontal(&val) {
                                        align.horizontal = h;
                                    }
                                }
                                b"vertical" => {
                                    if let Some(v) = str_to_vertical(&val) {
                                        align.vertical = v;
                                    }
                                }
                                b"wrapText" => {
                                    align.wrap_text = val.as_ref() == "1";
                                }
                                b"shrinkToFit" => {
                                    align.shrink_to_fit = val.as_ref() == "1";
                                }
                                b"indent" => {
                                    align.indent = val.parse::<u8>().unwrap_or(0);
                                }
                                b"textRotation" => {
                                    align.rotation = val.parse::<i16>().unwrap_or(0);
                                }
                                b"readingOrder" => {
                                    align.reading_order = match val.as_ref() {
                                        "1" => ReadingOrder::LeftToRight,
                                        "2" => ReadingOrder::RightToLeft,
                                        _ => ReadingOrder::ContextDependent,
                                    };
                                }
                                _ => {}
                            }
                        }
                    }
                }

                // protection can be self-closing
                b"protection" => {
                    if let Some((_n, _f, _fi, _b, _a, prot)) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.as_ref() {
                                b"locked" => prot.locked = val.as_ref() == "1",
                                b"hidden" => prot.hidden = val.as_ref() == "1",
                                _ => {}
                            }
                        }
                    }
                }

                // xf can be empty (no child elements)
                b"xf" if in_cell_xfs => {
                    // Parse ids
                    let mut num_fmt_id = 0u32;
                    let mut font_id = 0u32;
                    let mut fill_id = 0u32;
                    let mut border_id = 0u32;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"fontId" => {
                                font_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"fillId" => {
                                fill_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"borderId" => {
                                border_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            _ => {}
                        }
                    }
                    // Resolve immediately
                    let style = resolve_style(
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        Alignment::default(),
                        Protection::default(),
                        &numfmts,
                        &fonts,
                        &fills,
                        &borders,
                    );
                    cell_xfs.push(style);
                }

                _ => {}
            },

            Ok(Event::End(e)) => match e.name().as_ref() {
                b"font" => {
                    if in_dxf {
                        // DXF font - apply to current DXF style
                        if let (Some(f), Some(dxf)) = (dxf_font.take(), current_dxf.as_mut()) {
                            dxf.font = f;
                        }
                    } else if let Some(f) = current_font.take() {
                        fonts.push(f);
                    }
                }
                b"fill" => {
                    if in_dxf_fill {
                        // DXF fill - apply to current DXF style
                        let fill = finalize_fill(dxf_fill_pattern, dxf_fill_fg, dxf_fill_bg);
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.fill = fill;
                        }
                        in_dxf_fill = false;
                        dxf_fill_pattern = None;
                    } else if in_fill {
                        let fill =
                            finalize_fill(current_fill_pattern, current_fill_fg, current_fill_bg);
                        fills.push(fill);
                        in_fill = false;
                        current_fill_pattern = None;
                    }
                }
                b"border" => {
                    if in_dxf {
                        // DXF border - apply to current DXF style
                        if let (Some(b), Some(dxf)) = (dxf_border.take(), current_dxf.as_mut()) {
                            dxf.border = b;
                        }
                        dxf_border_edge = None;
                    } else if let Some(b) = current_border.take() {
                        borders.push(b);
                    }
                    current_border_edge = None;
                }
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                    current_border_edge = None;
                    dxf_border_edge = None;
                }
                b"dxf" => {
                    if let Some(dxf) = current_dxf.take() {
                        dxf_styles.push(dxf);
                    }
                    in_dxf = false;
                }
                b"dxfs" => {
                    in_dxfs = false;
                }
                b"xf" => {
                    if let Some((num_fmt_id, font_id, fill_id, border_id, align, prot)) =
                        current_xf.take()
                    {
                        let style = resolve_style(
                            num_fmt_id, font_id, fill_id, border_id, align, prot, &numfmts, &fonts,
                            &fills, &borders,
                        );
                        cell_xfs.push(style);
                    }
                }
                b"cellXfs" => {
                    in_cell_xfs = false;
                }
                _ => {}
            },

            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }

        buf.clear();
    }

    let cell_styles = if cell_xfs.is_empty() {
        vec![Style::default()]
    } else {
        cell_xfs
    };

    Ok(ParsedStyles {
        cell_styles,
        dxf_styles,
    })
}

fn resolve_style(
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Alignment,
    protection: Protection,
    numfmts: &HashMap<u32, String>,
    fonts: &[FontStyle],
    fills: &[FillStyle],
    borders: &[BorderStyle],
) -> Style {
    let mut style = Style::default();
    style.font = fonts.get(font_id as usize).cloned().unwrap_or_default();
    style.fill = fills.get(fill_id as usize).cloned().unwrap_or_default();
    style.border = borders.get(border_id as usize).cloned().unwrap_or_default();
    style.alignment = alignment;
    style.protection = protection;

    style.number_format = if num_fmt_id == 0 {
        NumberFormat::General
    } else if let Some(code) = numfmts.get(&num_fmt_id) {
        NumberFormat::Custom(code.clone())
    } else {
        NumberFormat::BuiltIn(num_fmt_id)
    };

    style
}

fn finalize_fill(pattern: Option<PatternType>, fg: Color, bg: Color) -> FillStyle {
    match pattern.unwrap_or(PatternType::None) {
        PatternType::None => FillStyle::None,
        PatternType::Solid => FillStyle::Solid { color: fg },
        PatternType::Gray125 => FillStyle::None,
        p => FillStyle::Pattern {
            pattern: p,
            foreground: fg,
            background: bg,
        },
    }
}

fn parse_color_attrs(e: &quick_xml::events::BytesStart<'_>) -> Color {
    // Priority: rgb > theme > indexed > auto
    let mut rgb: Option<String> = None;
    let mut theme: Option<u8> = None;
    let mut tint: Option<f64> = None;
    let mut indexed: Option<u8> = None;
    let mut auto = false;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"rgb" => {
                rgb = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"theme" => {
                theme = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"tint" => {
                tint = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<f64>().ok());
            }
            b"indexed" => {
                indexed = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"auto" => {
                auto = attr.unescape_value().ok().as_deref() == Some("1");
            }
            _ => {}
        }
    }

    if let Some(rgb) = rgb {
        let hex = rgb.trim_start_matches('#');
        if hex.len() == 8 {
            if let (Ok(a), Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
                u8::from_str_radix(&hex[6..8], 16),
            ) {
                return Color::Argb { a, r, g, b };
            }
        } else if hex.len() == 6 {
            if let (Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
            ) {
                return Color::Rgb { r, g, b };
            }
        }
    }

    if let Some(index) = theme {
        let tint_i8 = tint.map(|t| (t * 100.0).round() as i8).unwrap_or(0);
        return Color::Theme {
            index,
            tint: tint_i8,
        };
    }

    if let Some(i) = indexed {
        return Color::Indexed(i);
    }

    if auto {
        return Color::Auto;
    }

    Color::Auto
}

fn str_to_pattern_type(s: &str) -> Option<PatternType> {
    Some(match s {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "mediumGray" => PatternType::MediumGray,
        "darkGray" => PatternType::DarkGray,
        "lightGray" => PatternType::LightGray,
        "darkHorizontal" => PatternType::DarkHorizontal,
        "darkVertical" => PatternType::DarkVertical,
        "darkDown" => PatternType::DarkDown,
        "darkUp" => PatternType::DarkUp,
        "darkGrid" => PatternType::DarkGrid,
        "darkTrellis" => PatternType::DarkTrellis,
        "lightHorizontal" => PatternType::LightHorizontal,
        "lightVertical" => PatternType::LightVertical,
        "lightDown" => PatternType::LightDown,
        "lightUp" => PatternType::LightUp,
        "lightGrid" => PatternType::LightGrid,
        "lightTrellis" => PatternType::LightTrellis,
        "gray125" => PatternType::Gray125,
        "gray0625" => PatternType::Gray0625,
        _ => return None,
    })
}

fn str_to_border_style(s: &str) -> Option<BorderLineStyle> {
    Some(match s {
        "thin" => BorderLineStyle::Thin,
        "medium" => BorderLineStyle::Medium,
        "thick" => BorderLineStyle::Thick,
        "dashed" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "double" => BorderLineStyle::Double,
        "hair" => BorderLineStyle::Hair,
        "mediumDashed" => BorderLineStyle::MediumDashed,
        "dashDot" => BorderLineStyle::DashDot,
        "mediumDashDot" => BorderLineStyle::MediumDashDot,
        "dashDotDot" => BorderLineStyle::DashDotDot,
        "mediumDashDotDot" => BorderLineStyle::MediumDashDotDot,
        "slantDashDot" => BorderLineStyle::SlantDashDot,
        _ => return None,
    })
}

fn str_to_horizontal(s: &str) -> Option<HorizontalAlignment> {
    Some(match s {
        "general" => HorizontalAlignment::General,
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "fill" => HorizontalAlignment::Fill,
        "justify" => HorizontalAlignment::Justify,
        "centerContinuous" => HorizontalAlignment::CenterContinuous,
        "distributed" => HorizontalAlignment::Distributed,
        _ => return None,
    })
}

fn str_to_vertical(s: &str) -> Option<VerticalAlignment> {
    Some(match s {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => return None,
    })
}

fn str_to_underline(s: &str) -> Underline {
    match s {
        "double" => Underline::Double,
        "singleAccounting" => Underline::SingleAccounting,
        "doubleAccounting" => Underline::DoubleAccounting,
        _ => Underline::Single,
    }
}

fn get_border_edge<'a>(border: &'a BorderStyle, edge: &str) -> &'a Option<BorderEdge> {
    match edge {
        "left" => &border.left,
        "right" => &border.right,
        "top" => &border.top,
        "bottom" => &border.bottom,
        _ => &border.diagonal,
    }
}

fn set_border_edge(border: &mut BorderStyle, edge: &str, val: Option<BorderEdge>) {
    match edge {
        "left" => border.left = val,
        "right" => border.right = val,
        "top" => border.top = val,
        "bottom" => border.bottom = val,
        _ => border.diagonal = val,
    }
}
