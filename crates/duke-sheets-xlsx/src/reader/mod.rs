//! XLSX reader

use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::{read_styles_xml, register_roundtrip_style_data, ParsedStyles};
use duke_sheets_core::comment::CellComment;
use duke_sheets_core::conditional_format::{
    CfColorValue, CfOperator, CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
    IconSetStyle, TimePeriod,
};
use duke_sheets_core::style::{Color, Style};
use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::{CellAddress, CellError, CellRange, CellValue, SplitPanes, Workbook};

/// Decode Excel's `_xHHHH_` escape sequences in strings.
///
/// Excel uses this format to encode special characters in XML:
/// - `_x000d_` = CR (carriage return)
/// - `_x000a_` = LF (line feed)
/// - `_x0009_` = Tab
/// - `_x005f_` = Underscore (escaped underscore)
fn decode_excel_escapes(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();

    while let Some(c) = chars.next() {
        if c == '_' {
            // Check if this looks like _xHHHH_
            let mut hex_chars = String::new();
            let mut is_escape = false;

            if chars.peek() == Some(&'x') {
                chars.next(); // consume 'x'

                // Try to read 4 hex digits
                for _ in 0..4 {
                    if let Some(&ch) = chars.peek() {
                        if ch.is_ascii_hexdigit() {
                            hex_chars.push(ch);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                }

                // Check for closing underscore
                if hex_chars.len() == 4 && chars.peek() == Some(&'_') {
                    chars.next(); // consume closing '_'
                    if let Ok(code) = u32::from_str_radix(&hex_chars, 16) {
                        if let Some(decoded) = char::from_u32(code) {
                            result.push(decoded);
                            is_escape = true;
                        }
                    }
                }
            }

            if !is_escape {
                // Not a valid escape sequence, output what we consumed
                result.push('_');
                if !hex_chars.is_empty() {
                    result.push('x');
                    result.push_str(&hex_chars);
                }
            }
        } else {
            result.push(c);
        }
    }

    result
}

/// XLSX file reader
pub struct XlsxReader;

/// Parsed workbook properties from workbook.xml
struct WorkbookProps {
    sheets: Vec<(String, String)>,
    date_1904: bool,
    named_ranges: Vec<duke_sheets_core::named_range::NamedRange>,
}

struct WorkbookRels {
    sheet_paths: HashMap<String, String>,
    theme_path: Option<String>,
}

#[derive(Debug, Clone, Copy)]
struct ThemePalette {
    /// [lt1, dk1, lt2, dk2, accent1..accent6, hlink, fol_hlink]
    colors: [(u8, u8, u8); 12],
}

#[derive(Debug, Clone, Copy, Default)]
enum CellFormulaKind {
    #[default]
    Normal,
    Shared,
    Array,
    DataTable,
}

#[derive(Debug, Clone, Default)]
struct CellFormulaState {
    kind: CellFormulaKind,
    shared_index: Option<u32>,
    /// OOXML dataTable formula attrs: input cell refs (`r1` / `r2`).
    data_table_input1_ref: Option<String>,
    data_table_input2_ref: Option<String>,
}

#[derive(Debug, Clone)]
struct SharedFormulaMaster {
    base_cell_ref: String,
    formula: String,
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
    fn resolve_theme_color(&self, index: u8, tint: i8) -> (u8, u8, u8) {
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

impl XlsxReader {
    /// Read a workbook from a file path
    pub fn read_file<P: AsRef<Path>>(path: P) -> XlsxResult<Workbook> {
        let file = File::open(path)?;
        Self::read(file)
    }

    /// Read a workbook from a reader
    pub fn read<R: Read + Seek>(reader: R) -> XlsxResult<Workbook> {
        let mut archive = zip::ZipArchive::new(reader)?;

        // Verify this is an XLSX file
        if archive.by_name("[Content_Types].xml").is_err() {
            return Err(XlsxError::InvalidFormat(
                "Missing [Content_Types].xml".into(),
            ));
        }

        // Read shared strings (if present)
        let shared_strings = Self::read_shared_strings(&mut archive)?;

        // Read styles (if present)
        let mut parsed_styles = Self::read_styles(&mut archive)?;
        let roundtrip_style_data = parsed_styles.roundtrip_data();
        // Read workbook.xml.rels to get sheet/theme paths
        let workbook_rels = Self::read_workbook_rels(&mut archive)?;
        // Read workbook theme (if present) and resolve theme colors in styles
        let theme_palette =
            Self::read_theme_palette(&mut archive, workbook_rels.theme_path.as_deref())?;
        if let Some(theme) = theme_palette {
            for style in &mut parsed_styles.cell_styles {
                Self::resolve_style_theme_colors(style, &theme);
            }
            for style in &mut parsed_styles.dxf_styles {
                Self::resolve_style_theme_colors(style, &theme);
            }
        }
        let cell_styles = parsed_styles.cell_styles;
        let dxf_styles = parsed_styles.dxf_styles;

        // Read workbook.xml to get sheet info, properties, and defined names
        let wb_props = Self::read_workbook_xml(&mut archive)?;

        let sheet_paths = workbook_rels.sheet_paths;

        // Create workbook
        let mut workbook = Workbook::empty();
        workbook.settings_mut().date_1904 = wb_props.date_1904;

        // Add named ranges
        for nr in wb_props.named_ranges {
            workbook.named_ranges_mut().define_or_update(nr);
        }

        let sheet_info = &wb_props.sheets;
        let date_1904 = wb_props.date_1904;

        // Read each worksheet
        for (idx, (name, r_id)) in sheet_info.iter().enumerate() {
            if let Some(path) = sheet_paths.get(r_id) {
                let sheet_idx = workbook.add_worksheet_with_name(name)?;
                workbook
                    .worksheet_mut(sheet_idx)
                    .unwrap()
                    .set_date_1904(date_1904);
                Self::read_worksheet(
                    &mut archive,
                    path,
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                    &shared_strings,
                    &cell_styles,
                    &dxf_styles,
                    theme_palette.as_ref(),
                )?;

                // Read comments for this worksheet (if present)
                Self::read_worksheet_comments(
                    &mut archive,
                    idx,
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                )?;
            }
        }

        // Ensure at least one sheet exists
        if workbook.is_empty() {
            workbook.add_worksheet()?;
        }

        register_roundtrip_style_data(&workbook, roundtrip_style_data);

        Ok(workbook)
    }

    /// Read the shared strings table
    fn read_shared_strings<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> XlsxResult<Vec<String>> {
        let mut strings = Vec::new();

        let file = match archive.by_name("xl/sharedStrings.xml") {
            Ok(f) => f,
            Err(_) => return Ok(strings), // No shared strings is valid
        };

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_string = String::new();
        let mut in_si = false;
        let mut in_t = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_string.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                    b"si" => {
                        // Decode Excel's _xHHHH_ escape sequences
                        let decoded = decode_excel_escapes(&current_string);
                        strings.push(decoded);
                        current_string.clear();
                        in_si = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) if in_t => match e.unescape() {
                    Ok(text) => current_string.push_str(&text),
                    Err(err) => log::warn!(
                        "Shared string {}: XML unescape failed: {}",
                        strings.len(),
                        err
                    ),
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(strings)
    }

    fn read_styles<R: Read + Seek>(archive: &mut zip::ZipArchive<R>) -> XlsxResult<ParsedStyles> {
        let file = match archive.by_name("xl/styles.xml") {
            Ok(f) => f,
            Err(_) => {
                return Ok(ParsedStyles {
                    cell_styles: vec![Style::default()],
                    cell_style_xfs: vec![Style::default()],
                    named_styles: Vec::new(),
                    cell_xf_xf_ids: vec![0],
                    dxf_styles: Vec::new(),
                })
            }
        };
        read_styles_xml(file)
    }

    /// Read workbook.xml to get sheet names, rIds, workbook properties,
    /// and defined names.
    fn read_workbook_xml<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> XlsxResult<WorkbookProps> {
        use duke_sheets_core::named_range::{NameScope, NamedRange};

        let file = archive
            .by_name("xl/workbook.xml")
            .map_err(|_| XlsxError::MissingPart("xl/workbook.xml".into()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut sheets = Vec::new();
        let mut date_1904 = false;
        let mut named_ranges = Vec::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => match e.name().local_name().as_ref() {
                    b"sheet" => {
                        Self::parse_sheet_element(e, &mut sheets);
                    }
                    b"workbookPr" => {
                        Self::parse_workbook_pr(e, &mut date_1904);
                    }
                    _ => {}
                },
                Ok(Event::Start(ref e)) => match e.name().local_name().as_ref() {
                    b"sheet" => {
                        Self::parse_sheet_element(e, &mut sheets);
                    }
                    b"workbookPr" => {
                        Self::parse_workbook_pr(e, &mut date_1904);
                    }
                    b"definedName" => {
                        // Parse attributes
                        let mut dn_name = None;
                        let mut local_sheet_id: Option<usize> = None;
                        let mut comment = None;
                        let mut hidden = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"name" => {
                                    dn_name = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"localSheetId" => {
                                    local_sheet_id =
                                        attr.unescape_value().ok().and_then(|s| s.parse().ok());
                                }
                                b"comment" => {
                                    comment = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"hidden" => {
                                    hidden = attr.unescape_value().ok().map_or(false, |v| {
                                        v.as_ref() == "1" || v.eq_ignore_ascii_case("true")
                                    });
                                }
                                _ => {}
                            }
                        }

                        // Read the text content (the refers_to expression)
                        let mut text_buf = Vec::new();
                        let refers_to = match xml_reader.read_event_into(&mut text_buf) {
                            Ok(Event::Text(t)) => t.unescape().ok().map(|s| s.to_string()),
                            _ => None,
                        };

                        if let (Some(name), Some(refers_to)) = (dn_name, refers_to) {
                            let scope = match local_sheet_id {
                                Some(idx) => NameScope::Sheet(idx),
                                None => NameScope::Workbook,
                            };
                            let mut nr = NamedRange::new(name, refers_to, scope);
                            nr.comment = comment;
                            nr.hidden = hidden;
                            named_ranges.push(nr);
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(WorkbookProps {
            sheets,
            date_1904,
            named_ranges,
        })
    }

    fn parse_sheet_element(
        e: &quick_xml::events::BytesStart<'_>,
        sheets: &mut Vec<(String, String)>,
    ) {
        let mut name = None;
        let mut r_id = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"name" => {
                    name = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"id" => {
                    r_id = attr.unescape_value().ok().map(|s| s.to_string());
                }
                _ => {}
            }
        }

        if let (Some(name), Some(r_id)) = (name, r_id) {
            sheets.push((name, r_id));
        }
    }

    fn parse_workbook_pr(e: &quick_xml::events::BytesStart<'_>, date_1904: &mut bool) {
        for attr in e.attributes().flatten() {
            if attr.key.local_name().as_ref() == b"date1904" {
                if let Ok(val) = attr.unescape_value() {
                    *date_1904 = val.as_ref() == "1" || val.eq_ignore_ascii_case("true");
                }
            }
        }
    }

    /// Read workbook.xml.rels to get sheet file paths
    fn read_workbook_rels<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> XlsxResult<WorkbookRels> {
        let file = archive
            .by_name("xl/_rels/workbook.xml.rels")
            .map_err(|_| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".into()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut rels = HashMap::new();
        let mut theme_path: Option<String> = None;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.name().local_name().as_ref() == b"Relationship" =>
                {
                    let mut id = None;
                    let mut target = None;
                    let mut rel_type = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Id" => {
                                id = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"Target" => {
                                target = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"Type" => {
                                rel_type = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            _ => {}
                        }
                    }

                    // Include worksheet relationships and theme relationship
                    if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                        if rel_type.ends_with("/worksheet") {
                            // Target is relative to xl/ folder
                            let full_path = if target.starts_with('/') {
                                target[1..].to_string()
                            } else {
                                format!("xl/{}", target)
                            };
                            rels.insert(id, full_path);
                        } else if rel_type.ends_with("/theme") {
                            let full_path = if target.starts_with('/') {
                                target[1..].to_string()
                            } else {
                                format!("xl/{}", target)
                            };
                            theme_path = Some(full_path);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(WorkbookRels {
            sheet_paths: rels,
            theme_path,
        })
    }

    fn read_theme_palette<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        theme_path: Option<&str>,
    ) -> XlsxResult<Option<ThemePalette>> {
        let mut try_paths: Vec<String> = Vec::new();
        if let Some(path) = theme_path {
            try_paths.push(path.to_string());
        }
        if !try_paths.iter().any(|p| p == "xl/theme/theme1.xml") {
            try_paths.push("xl/theme/theme1.xml".to_string());
        }

        for path in try_paths {
            let file = match archive.by_name(&path) {
                Ok(f) => f,
                Err(_) => continue,
            };
            let reader = BufReader::new(file);
            let palette = Self::parse_theme_palette(reader)?;
            return Ok(Some(palette));
        }

        Ok(None)
    }

    fn parse_theme_palette<R: Read>(reader: R) -> XlsxResult<ThemePalette> {
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
                            if let Some(rgb) = Self::extract_theme_rgb_from_attrs(&e) {
                                palette.colors[slot] = rgb;
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                    b"srgbClr" | b"sysClr" if in_clr_scheme => {
                        if let Some(slot) = current_slot {
                            if let Some(rgb) = Self::extract_theme_rgb_from_attrs(&e) {
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

    fn resolve_style_theme_colors(style: &mut Style, theme: &ThemePalette) {
        style.font.color = Self::resolve_color_theme(style.font.color, theme);

        match &mut style.fill {
            duke_sheets_core::style::FillStyle::None => {}
            duke_sheets_core::style::FillStyle::Solid { color } => {
                *color = Self::resolve_color_theme(*color, theme);
            }
            duke_sheets_core::style::FillStyle::Pattern {
                foreground,
                background,
                ..
            } => {
                *foreground = Self::resolve_color_theme(*foreground, theme);
                *background = Self::resolve_color_theme(*background, theme);
            }
            duke_sheets_core::style::FillStyle::Gradient { stops, .. } => {
                for stop in stops {
                    stop.color = Self::resolve_color_theme(stop.color, theme);
                }
            }
        }

        for edge in [
            &mut style.border.left,
            &mut style.border.right,
            &mut style.border.top,
            &mut style.border.bottom,
            &mut style.border.diagonal,
        ] {
            if let Some(edge) = edge {
                edge.color = Self::resolve_color_theme(edge.color, theme);
            }
        }
    }

    fn resolve_color_theme(color: Color, theme: &ThemePalette) -> Color {
        match color {
            Color::Theme { index, tint } => {
                let (r, g, b) = theme.resolve_theme_color(index, tint);
                Color::Rgb { r, g, b }
            }
            other => other,
        }
    }

    /// Read a worksheet from the archive
    fn read_worksheet<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        path: &str,
        worksheet: &mut duke_sheets_core::Worksheet,
        shared_strings: &[String],
        cell_styles: &[Style],
        dxf_styles: &[Style],
        theme_palette: Option<&ThemePalette>,
    ) -> XlsxResult<()> {
        let file = archive
            .by_name(path)
            .map_err(|_| XlsxError::MissingPart(path.to_string()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        // Current cell state
        let mut current_cell_ref: Option<String> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_style: Option<u32> = None;
        let mut current_value: Option<String> = None;
        let mut current_formula: Option<String> = None;
        let mut current_formula_state = CellFormulaState::default();
        let mut in_cell = false;
        let mut in_value = false;
        let mut in_formula = false;
        let mut in_inline_str = false;
        let mut in_inline_text = false;
        let mut shared_formula_masters: HashMap<u32, SharedFormulaMaster> = HashMap::new();

        // Data validation state
        let mut in_data_validation = false;
        let mut current_validation: Option<DataValidation> = None;
        let mut in_dv_formula1 = false;
        let mut in_dv_formula2 = false;
        let mut dv_formula1: Option<String> = None;
        let mut dv_formula2: Option<String> = None;

        // Conditional formatting state
        let mut in_cond_formatting = false;
        let mut cf_sqref: Option<String> = None;
        let mut in_cf_rule = false;
        let mut current_cf_rule: Option<ConditionalFormatRule> = None;
        let mut in_cf_formula = false;
        let mut cf_formulas: Vec<String> = Vec::new();
        let mut in_odd_header = false;
        let mut in_odd_footer = false;

        // ColorScale/DataBar/IconSet state
        let mut in_color_scale = false;
        let mut in_data_bar = false;
        let mut in_icon_set = false;
        let mut cf_cfvo_values: Vec<CfValue> = Vec::new();
        let mut cf_colors: Vec<Color> = Vec::new();
        let mut icon_set_style: Option<IconSetStyle> = None;
        let mut icon_set_reverse = false;
        let mut icon_set_show_value = true;
        let mut data_bar_color: Option<Color> = None;
        let mut data_bar_show_value = true;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                    b"sheetView" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"tabSelected" => {
                                    if attr.unescape_value().ok().as_deref() == Some("1") {
                                        worksheet.set_selected(true);
                                    }
                                }
                                b"zoomScale" => {
                                    if let Some(z) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u16>().ok())
                                    {
                                        worksheet.set_zoom_scale(Some(z));
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"selection" => {
                        Self::parse_sheet_selection_attrs(&e, worksheet);
                    }
                    b"pageMargins" => {
                        let mut ps = worksheet.page_setup().clone();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"left" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.left_margin = v;
                                    }
                                }
                                b"right" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.right_margin = v;
                                    }
                                }
                                b"top" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.top_margin = v;
                                    }
                                }
                                b"bottom" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.bottom_margin = v;
                                    }
                                }
                                b"header" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.header_margin = v;
                                    }
                                }
                                b"footer" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok())
                                    {
                                        ps.footer_margin = v;
                                    }
                                }
                                _ => {}
                            }
                        }
                        worksheet.set_page_setup(ps);
                    }
                    b"pageSetup" => {
                        let mut ps = worksheet.page_setup().clone();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"paperSize" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u8>().ok())
                                    {
                                        ps.paper_size = v;
                                    }
                                }
                                b"orientation" => {
                                    if let Some(v) = attr.unescape_value().ok() {
                                        ps.orientation = if v.as_ref() == "landscape" {
                                            duke_sheets_core::PageOrientation::Landscape
                                        } else {
                                            duke_sheets_core::PageOrientation::Portrait
                                        };
                                    }
                                }
                                b"scale" => {
                                    if let Some(v) = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u16>().ok())
                                    {
                                        ps.scale = v;
                                    }
                                }
                                b"fitToWidth" => {
                                    ps.fit_to_width = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u16>().ok());
                                }
                                b"fitToHeight" => {
                                    ps.fit_to_height = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u16>().ok());
                                }
                                _ => {}
                            }
                        }
                        worksheet.set_page_setup(ps);
                    }
                    b"printOptions" => {
                        let mut ps = worksheet.page_setup().clone();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"gridLines" => {
                                    if let Some(v) = attr.unescape_value().ok() {
                                        ps.print_gridlines =
                                            v.as_ref() == "1" || v.as_ref() == "true";
                                    }
                                }
                                b"headings" => {
                                    if let Some(v) = attr.unescape_value().ok() {
                                        ps.print_headings =
                                            v.as_ref() == "1" || v.as_ref() == "true";
                                    }
                                }
                                _ => {}
                            }
                        }
                        worksheet.set_page_setup(ps);
                    }
                    b"oddHeader" => {
                        in_odd_header = true;
                    }
                    b"oddFooter" => {
                        in_odd_footer = true;
                    }
                    b"row" => {
                        // Parse row dimensions: ht, customHeight, hidden
                        let mut row_num: Option<u32> = None;
                        let mut ht: Option<f64> = None;
                        let mut custom_height = false;
                        let mut hidden = false;
                        let mut outline_level: Option<u8> = None;
                        let mut collapsed = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"r" => {
                                    row_num = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u32>().ok());
                                }
                                b"ht" => {
                                    ht = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<f64>().ok());
                                }
                                b"customHeight" => {
                                    custom_height = attr.unescape_value().ok().map_or(false, |s| {
                                        s.as_ref() == "1" || s.as_ref() == "true"
                                    });
                                }
                                b"hidden" => {
                                    hidden = attr.unescape_value().ok().map_or(false, |s| {
                                        s.as_ref() == "1" || s.as_ref() == "true"
                                    });
                                }
                                b"outlineLevel" => {
                                    outline_level = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u8>().ok());
                                }
                                b"collapsed" => {
                                    collapsed = attr.unescape_value().ok().map_or(false, |s| {
                                        s.as_ref() == "1" || s.as_ref() == "true"
                                    });
                                }
                                _ => {}
                            }
                        }
                        if let Some(r) = row_num {
                            let row_idx = r.saturating_sub(1); // 1-based to 0-based
                            if custom_height {
                                if let Some(h) = ht {
                                    worksheet.set_row_height(row_idx, h);
                                }
                            }
                            if hidden {
                                worksheet.set_row_hidden(row_idx, true);
                            }
                            if let Some(level) = outline_level {
                                worksheet.set_row_outline_level(row_idx, level);
                            }
                            if collapsed {
                                worksheet.set_row_collapsed(row_idx, true);
                            }
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        current_cell_ref = None;
                        current_cell_type = None;
                        current_cell_style = None;
                        current_value = None;
                        current_formula = None;
                        current_formula_state = CellFormulaState::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"r" => {
                                    current_cell_ref =
                                        attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"t" => {
                                    current_cell_type =
                                        attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"s" => {
                                    current_cell_style = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse::<u32>().ok());
                                }
                                _ => {}
                            }
                        }
                    }
                    b"v" if in_cell => {
                        in_value = true;
                    }
                    b"f" if in_cell => {
                        current_formula_state = Self::parse_cell_formula_state(&e);
                        in_formula = true;
                    }
                    b"is" if in_cell => {
                        in_inline_str = true;
                    }
                    b"t" if in_inline_str => {
                        in_inline_text = true;
                    }
                    // Data validation parsing
                    b"dataValidation" => {
                        in_data_validation = true;
                        dv_formula1 = None;
                        dv_formula2 = None;
                        current_validation = Some(Self::parse_data_validation_attrs(&e));
                    }
                    b"formula1" if in_data_validation => {
                        in_dv_formula1 = true;
                    }
                    b"formula2" if in_data_validation => {
                        in_dv_formula2 = true;
                    }
                    // Conditional formatting parsing
                    b"conditionalFormatting" => {
                        in_cond_formatting = true;
                        cf_sqref = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"sqref" {
                                cf_sqref = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"cfRule" if in_cond_formatting => {
                        in_cf_rule = true;
                        cf_formulas.clear();
                        cf_cfvo_values.clear();
                        cf_colors.clear();
                        icon_set_style = None;
                        icon_set_reverse = false;
                        icon_set_show_value = true;
                        data_bar_color = None;
                        data_bar_show_value = true;
                        current_cf_rule = Some(Self::parse_cf_rule_attrs(&e, cf_sqref.as_deref()));
                    }
                    b"formula" if in_cf_rule => {
                        in_cf_formula = true;
                    }
                    b"colorScale" if in_cf_rule => {
                        in_color_scale = true;
                    }
                    b"dataBar" if in_cf_rule => {
                        in_data_bar = true;
                        // Parse dataBar attributes
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"showValue" {
                                data_bar_show_value =
                                    attr.unescape_value().ok().map_or(true, |s| s != "0");
                            }
                        }
                    }
                    b"iconSet" if in_cf_rule => {
                        in_icon_set = true;
                        // Parse iconSet attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"iconSet" => {
                                    icon_set_style = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| IconSetStyle::from_xlsx(&s));
                                }
                                b"reverse" => {
                                    icon_set_reverse =
                                        attr.unescape_value().ok().map_or(false, |s| s == "1");
                                }
                                b"showValue" => {
                                    icon_set_show_value =
                                        attr.unescape_value().ok().map_or(true, |s| s != "0");
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    match e.name().local_name().as_ref() {
                        b"c" => {
                            // Process the cell
                            if let Some(ref cell_ref) = current_cell_ref {
                                let resolved_formula = Self::resolve_cell_formula(
                                    cell_ref,
                                    current_formula.as_deref(),
                                    &current_formula_state,
                                    &mut shared_formula_masters,
                                );
                                Self::process_cell(
                                    worksheet,
                                    cell_ref,
                                    current_cell_type.as_deref(),
                                    current_value.as_deref(),
                                    resolved_formula.as_deref(),
                                    current_cell_style,
                                    shared_strings,
                                    cell_styles,
                                )?;
                            }
                            in_cell = false;
                        }
                        b"v" => {
                            in_value = false;
                        }
                        b"f" => {
                            in_formula = false;
                        }
                        b"is" => {
                            in_inline_str = false;
                        }
                        b"t" if in_inline_str => {
                            in_inline_text = false;
                        }
                        // Data validation end events
                        b"dataValidation" => {
                            if let Some(mut validation) = current_validation.take() {
                                // Apply formula values based on validation type
                                Self::apply_validation_formulas(
                                    &mut validation,
                                    dv_formula1.take(),
                                    dv_formula2.take(),
                                );
                                worksheet.add_data_validation(validation);
                            }
                            in_data_validation = false;
                        }
                        b"formula1" if in_data_validation => {
                            in_dv_formula1 = false;
                        }
                        b"formula2" if in_data_validation => {
                            in_dv_formula2 = false;
                        }
                        // Conditional formatting end events
                        b"colorScale" => {
                            // Build ColorScale rule type from collected cfvo and color values
                            if let Some(ref mut rule) = current_cf_rule {
                                if cf_cfvo_values.len() == cf_colors.len()
                                    && !cf_cfvo_values.is_empty()
                                {
                                    let colors: Vec<CfColorValue> = cf_cfvo_values
                                        .iter()
                                        .zip(cf_colors.iter())
                                        .map(|(cfvo, color)| {
                                            CfColorValue::new(
                                                cfvo.value_type,
                                                cfvo.value.clone(),
                                                *color,
                                            )
                                        })
                                        .collect();
                                    rule.rule_type = CfRuleType::ColorScale { colors };
                                }
                            }
                            in_color_scale = false;
                        }
                        b"dataBar" => {
                            // Build DataBar rule type from collected values
                            if let Some(ref mut rule) = current_cf_rule {
                                let min_value =
                                    cf_cfvo_values.first().cloned().unwrap_or_else(CfValue::min);
                                let max_value =
                                    cf_cfvo_values.get(1).cloned().unwrap_or_else(CfValue::max);
                                let color =
                                    data_bar_color.unwrap_or_else(|| Color::rgb(99, 142, 198));
                                rule.rule_type = CfRuleType::DataBar {
                                    min_value,
                                    max_value,
                                    color,
                                    show_value: data_bar_show_value,
                                    gradient: true,
                                    border_color: None,
                                    negative_color: None,
                                };
                            }
                            in_data_bar = false;
                        }
                        b"iconSet" => {
                            // Build IconSet rule type from collected values
                            if let Some(ref mut rule) = current_cf_rule {
                                rule.rule_type = CfRuleType::IconSet {
                                    icon_style: icon_set_style.unwrap_or(IconSetStyle::Arrows3),
                                    values: cf_cfvo_values.clone(),
                                    reverse: icon_set_reverse,
                                    show_value: icon_set_show_value,
                                };
                            }
                            in_icon_set = false;
                        }
                        b"cfRule" => {
                            if let Some(mut rule) = current_cf_rule.take() {
                                Self::apply_cf_formulas(&mut rule, &cf_formulas);
                                // Apply DXF style if present
                                if let Some(dxf_id) = rule.dxf_id {
                                    if let Some(dxf_style) = dxf_styles.get(dxf_id as usize) {
                                        rule.format = Some(dxf_style.clone());
                                    }
                                }
                                worksheet.add_conditional_format(rule);
                            }
                            in_cf_rule = false;
                        }
                        b"conditionalFormatting" => {
                            in_cond_formatting = false;
                            cf_sqref = None;
                        }
                        b"formula" if in_cf_rule => {
                            in_cf_formula = false;
                        }
                        b"oddHeader" => {
                            in_odd_header = false;
                        }
                        b"oddFooter" => {
                            in_odd_footer = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_value {
                        match e.unescape() {
                            Ok(text) => current_value = Some(text.to_string()),
                            Err(err) => log::warn!(
                                "Cell {:?}: value unescape failed: {}",
                                current_cell_ref,
                                err
                            ),
                        }
                    } else if in_formula {
                        match e.unescape() {
                            Ok(text) => current_formula = Some(text.to_string()),
                            Err(err) => log::warn!(
                                "Cell {:?}: formula unescape failed: {}",
                                current_cell_ref,
                                err
                            ),
                        }
                    } else if in_inline_text {
                        match e.unescape() {
                            Ok(text) => {
                                // Inline string - store directly as value
                                current_value = Some(text.to_string());
                                current_cell_type = Some("inlineStr".to_string());
                            }
                            Err(err) => log::warn!(
                                "Cell {:?}: inline string unescape failed: {}",
                                current_cell_ref,
                                err
                            ),
                        }
                    } else if in_dv_formula1 {
                        match e.unescape() {
                            Ok(text) => dv_formula1 = Some(text.to_string()),
                            Err(err) => {
                                log::warn!("Data validation formula1 unescape failed: {}", err)
                            }
                        }
                    } else if in_dv_formula2 {
                        match e.unescape() {
                            Ok(text) => dv_formula2 = Some(text.to_string()),
                            Err(err) => {
                                log::warn!("Data validation formula2 unescape failed: {}", err)
                            }
                        }
                    } else if in_cf_formula {
                        match e.unescape() {
                            Ok(text) => cf_formulas.push(text.to_string()),
                            Err(err) => {
                                log::warn!("Conditional format formula unescape failed: {}", err)
                            }
                        }
                    } else if in_odd_header {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.odd_header = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    } else if in_odd_footer {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.odd_footer = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    match e.name().local_name().as_ref() {
                        b"pageMargins" => {
                            let mut ps = worksheet.page_setup().clone();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"left" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.left_margin = v;
                                        }
                                    }
                                    b"right" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.right_margin = v;
                                        }
                                    }
                                    b"top" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.top_margin = v;
                                        }
                                    }
                                    b"bottom" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.bottom_margin = v;
                                        }
                                    }
                                    b"header" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.header_margin = v;
                                        }
                                    }
                                    b"footer" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok())
                                        {
                                            ps.footer_margin = v;
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"pageSetup" => {
                            let mut ps = worksheet.page_setup().clone();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"paperSize" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u8>().ok())
                                        {
                                            ps.paper_size = v;
                                        }
                                    }
                                    b"orientation" => {
                                        if let Some(v) = attr.unescape_value().ok() {
                                            ps.orientation = if v.as_ref() == "landscape" {
                                                duke_sheets_core::PageOrientation::Landscape
                                            } else {
                                                duke_sheets_core::PageOrientation::Portrait
                                            };
                                        }
                                    }
                                    b"scale" => {
                                        if let Some(v) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok())
                                        {
                                            ps.scale = v;
                                        }
                                    }
                                    b"fitToWidth" => {
                                        ps.fit_to_width = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok());
                                    }
                                    b"fitToHeight" => {
                                        ps.fit_to_height = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok());
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"printOptions" => {
                            let mut ps = worksheet.page_setup().clone();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"gridLines" => {
                                        if let Some(v) = attr.unescape_value().ok() {
                                            ps.print_gridlines =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"headings" => {
                                        if let Some(v) = attr.unescape_value().ok() {
                                            ps.print_headings =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"sheetView" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"tabSelected" => {
                                        if attr.unescape_value().ok().as_deref() == Some("1") {
                                            worksheet.set_selected(true);
                                        }
                                    }
                                    b"zoomScale" => {
                                        if let Some(z) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok())
                                        {
                                            worksheet.set_zoom_scale(Some(z));
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"selection" => {
                            Self::parse_sheet_selection_attrs(&e, worksheet);
                        }
                        b"pane" => {
                            let mut state: Option<String> = None;
                            let mut x_split_raw: Option<f64> = None;
                            let mut y_split_raw: Option<f64> = None;
                            let mut top_left_cell: Option<(u32, u16)> = None;
                            let mut active_pane: Option<String> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"state" => {
                                        state = attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    b"xSplit" => {
                                        x_split_raw = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"ySplit" => {
                                        y_split_raw = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"topLeftCell" => {
                                        if let Some(a1) =
                                            attr.unescape_value().ok().map(|s| s.to_string())
                                        {
                                            if let Ok(addr) = CellAddress::parse(&a1) {
                                                top_left_cell = Some((addr.row, addr.col));
                                            }
                                        }
                                    }
                                    b"activePane" => {
                                        active_pane =
                                            attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }

                            match state.as_deref() {
                                Some("frozen") | Some("frozenSplit") => {
                                    let row = y_split_raw.unwrap_or(0.0).round().max(0.0) as u32;
                                    let col = x_split_raw.unwrap_or(0.0).round().max(0.0) as u16;
                                    worksheet.set_freeze_panes(row, col);
                                }
                                Some("split") => {
                                    worksheet.set_split_panes(Some(SplitPanes {
                                        x_split: x_split_raw.unwrap_or(0.0),
                                        y_split: y_split_raw.unwrap_or(0.0),
                                        top_left: top_left_cell,
                                        active_pane,
                                    }));
                                }
                                _ => {}
                            }
                        }
                        b"f" if in_cell => {
                            // Self-closing formula elements appear for shared formula
                            // follower cells: <f t="shared" si="0"/>
                            current_formula_state = Self::parse_cell_formula_state(&e);
                            in_formula = false;
                        }
                        b"row" => {
                            // Self-closing <row .../> with no cells — may have dimensions
                            let mut row_num: Option<u32> = None;
                            let mut ht: Option<f64> = None;
                            let mut custom_height = false;
                            let mut hidden = false;
                            let mut outline_level: Option<u8> = None;
                            let mut collapsed = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"r" => {
                                        row_num = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    b"ht" => {
                                        ht = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"customHeight" => {
                                        custom_height =
                                            attr.unescape_value().ok().map_or(false, |s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value().ok().map_or(false, |s| {
                                            s.as_ref() == "1" || s.as_ref() == "true"
                                        });
                                    }
                                    b"outlineLevel" => {
                                        outline_level = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u8>().ok());
                                    }
                                    b"collapsed" => {
                                        collapsed = attr.unescape_value().ok().map_or(false, |s| {
                                            s.as_ref() == "1" || s.as_ref() == "true"
                                        });
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(r) = row_num {
                                let row_idx = r.saturating_sub(1);
                                if custom_height {
                                    if let Some(h) = ht {
                                        worksheet.set_row_height(row_idx, h);
                                    }
                                }
                                if hidden {
                                    worksheet.set_row_hidden(row_idx, true);
                                }
                                if let Some(level) = outline_level {
                                    worksheet.set_row_outline_level(row_idx, level);
                                }
                                if collapsed {
                                    worksheet.set_row_collapsed(row_idx, true);
                                }
                            }
                        }
                        b"col" => {
                            // Parse column dimensions: min, max, width, customWidth, hidden
                            let mut col_min: Option<u16> = None;
                            let mut col_max: Option<u16> = None;
                            let mut width: Option<f64> = None;
                            let mut custom_width = false;
                            let mut hidden = false;
                            let mut outline_level: Option<u8> = None;
                            let mut collapsed = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"min" => {
                                        col_min = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok());
                                    }
                                    b"max" => {
                                        col_max = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u16>().ok());
                                    }
                                    b"width" => {
                                        width = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"customWidth" => {
                                        custom_width =
                                            attr.unescape_value().ok().map_or(false, |s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value().ok().map_or(false, |s| {
                                            s.as_ref() == "1" || s.as_ref() == "true"
                                        });
                                    }
                                    b"outlineLevel" => {
                                        outline_level = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u8>().ok());
                                    }
                                    b"collapsed" => {
                                        collapsed = attr.unescape_value().ok().map_or(false, |s| {
                                            s.as_ref() == "1" || s.as_ref() == "true"
                                        });
                                    }
                                    _ => {}
                                }
                            }
                            if let (Some(min), Some(max)) = (col_min, col_max) {
                                // min/max are 1-based in XLSX
                                for col in min..=max {
                                    let col_idx = col.saturating_sub(1); // 0-based
                                    if custom_width {
                                        if let Some(w) = width {
                                            worksheet.set_column_width(col_idx, w);
                                        }
                                    }
                                    if hidden {
                                        worksheet.set_column_hidden(col_idx, true);
                                    }
                                    if let Some(level) = outline_level {
                                        worksheet.set_column_outline_level(col_idx, level);
                                    }
                                    if collapsed {
                                        worksheet.set_column_collapsed(col_idx, true);
                                    }
                                }
                            }
                        }
                        b"c" => {
                            // Empty cell element (may still carry a style)
                            let mut cell_ref: Option<String> = None;
                            let mut cell_type: Option<String> = None;
                            let mut cell_style: Option<u32> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"r" => {
                                        cell_ref =
                                            attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    b"t" => {
                                        cell_type =
                                            attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    b"s" => {
                                        cell_style = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    _ => {}
                                }
                            }

                            if let Some(cell_ref) = cell_ref {
                                Self::process_cell(
                                    worksheet,
                                    &cell_ref,
                                    cell_type.as_deref(),
                                    None,
                                    None,
                                    cell_style,
                                    shared_strings,
                                    cell_styles,
                                )?;
                            }
                        }
                        // Parse cfvo (conditional format value object) elements
                        b"cfvo" if in_color_scale || in_data_bar || in_icon_set => {
                            let mut value_type = CfValueType::Min;
                            let mut value: Option<String> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"type" => {
                                        if let Some(t) = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| CfValueType::from_xlsx(&s))
                                        {
                                            value_type = t;
                                        }
                                    }
                                    b"val" => {
                                        value = attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }

                            cf_cfvo_values.push(CfValue::new(value_type, value));
                        }
                        // Parse color elements for colorScale and dataBar
                        b"color" if in_color_scale || in_data_bar => {
                            let color = Self::parse_color_element(&e, theme_palette);
                            if in_color_scale {
                                cf_colors.push(color);
                            } else if in_data_bar {
                                data_bar_color = Some(color);
                            }
                        }
                        // Merged cells
                        b"mergeCell" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    let ref_str = String::from_utf8_lossy(&attr.value);
                                    match CellRange::parse(&ref_str) {
                                        Ok(range) => {
                                            if let Err(e) = worksheet.merge_cells(&range) {
                                                log::warn!("Skipping merge '{}': {}", ref_str, e);
                                            }
                                        }
                                        Err(e) => {
                                            log::warn!("Invalid merge ref '{}': {}", ref_str, e)
                                        }
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    fn parse_cell_formula_state(e: &quick_xml::events::BytesStart<'_>) -> CellFormulaState {
        let mut state = CellFormulaState::default();

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"t" => {
                    if let Ok(v) = attr.unescape_value() {
                        state.kind = match v.as_ref() {
                            "shared" => CellFormulaKind::Shared,
                            "array" => CellFormulaKind::Array,
                            "dataTable" => CellFormulaKind::DataTable,
                            _ => CellFormulaKind::Normal,
                        };
                    }
                }
                b"si" => {
                    state.shared_index = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok());
                }
                b"r1" => {
                    state.data_table_input1_ref = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"r2" => {
                    state.data_table_input2_ref = attr.unescape_value().ok().map(|s| s.to_string());
                }
                _ => {}
            }
        }

        state
    }

    fn parse_sheet_selection_attrs(
        e: &quick_xml::events::BytesStart<'_>,
        worksheet: &mut duke_sheets_core::Worksheet,
    ) {
        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"activeCell" => {
                    if let Some(cell) = attr.unescape_value().ok().map(|s| s.to_string()) {
                        if let Ok(addr) = CellAddress::parse(&cell) {
                            worksheet.set_selection_active_cell(addr.row, addr.col);
                        }
                    }
                }
                b"sqref" => {
                    if let Some(sqref) = attr.unescape_value().ok().map(|s| s.to_string()) {
                        if let Some(first) = sqref.split_whitespace().next() {
                            if let Ok(range) = CellRange::parse(first) {
                                worksheet.set_selection_range(Some(range));
                            }
                        }
                    }
                }
                _ => {}
            }
        }
    }

    fn resolve_cell_formula(
        cell_ref: &str,
        formula: Option<&str>,
        formula_state: &CellFormulaState,
        shared_formula_masters: &mut HashMap<u32, SharedFormulaMaster>,
    ) -> Option<String> {
        match formula_state.kind {
            CellFormulaKind::Normal | CellFormulaKind::Array => formula.map(|f| f.to_string()),
            CellFormulaKind::DataTable => match formula {
                Some(f) => Some(f.to_string()),
                None => {
                    let arg1 = formula_state.data_table_input1_ref.as_deref().unwrap_or("");
                    let arg2 = formula_state.data_table_input2_ref.as_deref().unwrap_or("");
                    Some(format!("TABLE({},{})", arg1, arg2))
                }
            },
            CellFormulaKind::Shared => {
                let si = formula_state.shared_index?;
                if let Some(f) = formula {
                    // Shared formula master cell
                    shared_formula_masters.insert(
                        si,
                        SharedFormulaMaster {
                            base_cell_ref: cell_ref.to_string(),
                            formula: f.to_string(),
                        },
                    );
                    Some(f.to_string())
                } else {
                    // Shared formula follower cell
                    let master = shared_formula_masters.get(&si)?;
                    Some(Self::translate_shared_formula(
                        &master.formula,
                        &master.base_cell_ref,
                        cell_ref,
                    ))
                }
            }
        }
    }

    fn translate_shared_formula(formula: &str, base_cell_ref: &str, cell_ref: &str) -> String {
        let base = match CellAddress::parse(base_cell_ref) {
            Ok(v) => v,
            Err(_) => return formula.to_string(),
        };
        let target = match CellAddress::parse(cell_ref) {
            Ok(v) => v,
            Err(_) => return formula.to_string(),
        };

        let row_delta = target.row as i32 - base.row as i32;
        let col_delta = target.col as i32 - base.col as i32;

        Self::shift_a1_references(formula, row_delta, col_delta)
    }

    fn shift_a1_references(formula: &str, row_delta: i32, col_delta: i32) -> String {
        let bytes = formula.as_bytes();
        let mut out = String::with_capacity(formula.len());
        let mut i = 0usize;
        let mut in_string = false;

        while i < bytes.len() {
            let ch = bytes[i] as char;

            if ch == '"' {
                in_string = !in_string;
                out.push(ch);
                i += 1;
                continue;
            }

            if !in_string {
                if i > 0 {
                    let prev = bytes[i - 1] as char;
                    if prev.is_ascii_alphanumeric() || prev == '_' || prev == '.' {
                        out.push(ch);
                        i += 1;
                        continue;
                    }
                }
                if let Some((consumed, shifted)) =
                    Self::try_shift_cell_ref(&formula[i..], row_delta, col_delta)
                {
                    out.push_str(&shifted);
                    i += consumed;
                    continue;
                }
            }

            out.push(ch);
            i += 1;
        }

        out
    }

    fn try_shift_cell_ref(s: &str, row_delta: i32, col_delta: i32) -> Option<(usize, String)> {
        let b = s.as_bytes();
        let mut i = 0usize;

        let col_abs = if b.get(i) == Some(&b'$') {
            i += 1;
            true
        } else {
            false
        };

        let col_start = i;
        while let Some(&c) = b.get(i) {
            if (c as char).is_ascii_uppercase() {
                i += 1;
            } else {
                break;
            }
        }
        if i == col_start {
            return None;
        }

        let col_letters = &s[col_start..i];
        let mut col = Self::a1_col_to_index(col_letters)? as i32;

        let row_abs = if b.get(i) == Some(&b'$') {
            i += 1;
            true
        } else {
            false
        };

        let row_start = i;
        while let Some(&c) = b.get(i) {
            if (c as char).is_ascii_digit() {
                i += 1;
            } else {
                break;
            }
        }
        if i == row_start {
            return None;
        }

        let mut row: i32 = s[row_start..i].parse::<i32>().ok()?.saturating_sub(1);

        // Must be token boundary (avoid matching inside names)
        if let Some(&next) = b.get(i) {
            let next = next as char;
            if next.is_ascii_alphanumeric() || next == '_' || next == '.' {
                return None;
            }
        }

        if !col_abs {
            col += col_delta;
        }
        if !row_abs {
            row += row_delta;
        }

        if col < 0 || row < 0 {
            return Some((i, "#REF!".to_string()));
        }

        let mut shifted = String::new();
        if col_abs {
            shifted.push('$');
        }
        shifted.push_str(&Self::a1_index_to_col(col as u16));
        if row_abs {
            shifted.push('$');
        }
        shifted.push_str(&(row as u32 + 1).to_string());

        Some((i, shifted))
    }

    fn a1_col_to_index(col: &str) -> Option<u16> {
        let mut value: u32 = 0;
        for ch in col.chars() {
            if !ch.is_ascii_uppercase() {
                return None;
            }
            value = value
                .saturating_mul(26)
                .saturating_add((ch as u8 - b'A' + 1) as u32);
        }
        if value == 0 {
            None
        } else {
            u16::try_from(value - 1).ok()
        }
    }

    fn a1_index_to_col(mut index: u16) -> String {
        let mut col = String::new();
        loop {
            let rem = (index % 26) as u8;
            col.insert(0, (b'A' + rem) as char);
            if index < 26 {
                break;
            }
            index = index / 26 - 1;
        }
        col
    }

    /// Process a cell and add it to the worksheet
    fn process_cell(
        worksheet: &mut duke_sheets_core::Worksheet,
        cell_ref: &str,
        cell_type: Option<&str>,
        value: Option<&str>,
        formula: Option<&str>,
        style_idx: Option<u32>,
        shared_strings: &[String],
        styles: &[Style],
    ) -> XlsxResult<()> {
        let addr = match CellAddress::parse(cell_ref) {
            Ok(a) => a,
            Err(e) => {
                log::warn!("Skipping cell with invalid reference '{}': {}", cell_ref, e);
                return Ok(());
            }
        };

        // Apply formula or value
        if let Some(f) = formula {
            // Parse cached value (if any) from the <v> element
            let cached = value.and_then(|v| match cell_type {
                Some("b") => Some(CellValue::Boolean(
                    v == "1" || v.eq_ignore_ascii_case("true"),
                )),
                Some("e") => CellError::from_str(v).map(CellValue::Error),
                Some("s") => {
                    let idx: usize = v.parse().ok()?;
                    shared_strings
                        .get(idx)
                        .map(|s| CellValue::String(s.clone().into()))
                }
                Some("str") | Some("inlineStr") => Some(CellValue::String(v.to_string().into())),
                None | Some("n") => v.parse::<f64>().ok().map(CellValue::Number),
                Some(_) => Some(CellValue::String(v.to_string().into())),
            });

            // Ensure formula starts with '='
            let formula_text = if f.starts_with('=') {
                f.to_string()
            } else {
                format!("={}", f)
            };

            if let Err(e) = worksheet.set_cell_value_at(
                addr.row,
                addr.col,
                CellValue::Formula {
                    text: formula_text,
                    cached_value: cached.map(Box::new),
                    array_result: None,
                },
            ) {
                log::warn!("Skipping cell {}: {}", cell_ref, e);
                return Ok(());
            }
        } else if let Some(value) = value {
            // Process value based on type
            let cell_value = match cell_type {
                // Shared string
                Some("s") => match value.parse::<usize>() {
                    Ok(idx) => match shared_strings.get(idx) {
                        Some(s) => CellValue::String(s.clone().into()),
                        None => {
                            log::warn!(
                                "Cell {}: shared string index {} out of bounds (max {}), using #REF!",
                                cell_ref, idx, shared_strings.len()
                            );
                            CellValue::Error(CellError::Ref)
                        }
                    },
                    Err(_) => {
                        log::warn!(
                            "Cell {}: invalid shared string index '{}', using #REF!",
                            cell_ref,
                            value
                        );
                        CellValue::Error(CellError::Ref)
                    }
                },

                // Boolean
                Some("b") => CellValue::Boolean(value == "1" || value.eq_ignore_ascii_case("true")),

                // Error
                Some("e") => CellError::from_str(value)
                    .map(CellValue::Error)
                    .unwrap_or_else(|| CellValue::String(value.to_string().into())),

                // Inline string - decode Excel escape sequences
                Some("inlineStr") => CellValue::String(decode_excel_escapes(value).into()),

                // String (explicit type) - decode Excel escape sequences
                Some("str") => CellValue::String(decode_excel_escapes(value).into()),

                // Number (default type or explicit "n")
                None | Some("n") => match value.parse::<f64>() {
                    Ok(n) => CellValue::Number(n),
                    Err(_) => CellValue::String(value.to_string().into()),
                },

                // Unknown type - treat as string
                Some(_) => CellValue::String(value.to_string().into()),
            };

            if let Err(e) = worksheet.set_cell_value_at(addr.row, addr.col, cell_value) {
                log::warn!("Skipping cell {}: {}", cell_ref, e);
                return Ok(());
            }
        }

        // Apply style (if any)
        if let Some(s) = style_idx {
            if s != 0 {
                match styles.get(s as usize) {
                    Some(style) => {
                        if let Err(e) = worksheet.set_cell_style_at(addr.row, addr.col, style) {
                            log::warn!("Cell {}: failed to apply style: {}", cell_ref, e);
                        }
                    }
                    None => {
                        log::warn!(
                            "Cell {}: style index {} out of bounds (max {}), using default",
                            cell_ref,
                            s,
                            styles.len()
                        );
                    }
                }
            }
        }

        Ok(())
    }

    /// Parse data validation attributes from an element
    fn parse_data_validation_attrs(e: &quick_xml::events::BytesStart) -> DataValidation {
        let mut validation = DataValidation::new();
        let mut dv_type: Option<String> = None;
        let mut operator: Option<String> = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"type" => {
                    dv_type = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"operator" => {
                    operator = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"allowBlank" => {
                    validation.allow_blank = attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"showDropDown" => {
                    // Note: Excel uses showDropDown="1" to HIDE the dropdown (counterintuitive)
                    validation.show_dropdown =
                        attr.unescape_value().ok().map_or(true, |s| s != "1");
                }
                b"showInputMessage" => {
                    validation.show_input_message =
                        attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"showErrorMessage" => {
                    validation.show_error_alert =
                        attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"errorStyle" => {
                    if let Some(style) = attr.unescape_value().ok() {
                        validation.error_style = match style.as_ref() {
                            "warning" => ValidationErrorStyle::Warning,
                            "information" => ValidationErrorStyle::Information,
                            _ => ValidationErrorStyle::Stop,
                        };
                    }
                }
                b"errorTitle" => {
                    validation.error_title = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"error" => {
                    validation.error_message = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"promptTitle" => {
                    validation.input_title = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"prompt" => {
                    validation.input_message = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"sqref" => {
                    if let Some(sqref) = attr.unescape_value().ok() {
                        validation.ranges = Self::parse_sqref(&sqref);
                    }
                }
                _ => {}
            }
        }

        // Set the validation type based on parsed attributes
        let op = operator
            .as_deref()
            .and_then(ValidationOperator::from_xlsx)
            .unwrap_or(ValidationOperator::Between);

        validation.validation_type = match dv_type.as_deref() {
            Some("list") => ValidationType::List {
                source: String::new(),
            },
            Some("whole") => ValidationType::Whole {
                operator: op,
                value1: String::new(),
                value2: None,
            },
            Some("decimal") => ValidationType::Decimal {
                operator: op,
                value1: String::new(),
                value2: None,
            },
            Some("date") => ValidationType::Date {
                operator: op,
                value1: String::new(),
                value2: None,
            },
            Some("time") => ValidationType::Time {
                operator: op,
                value1: String::new(),
                value2: None,
            },
            Some("textLength") => ValidationType::TextLength {
                operator: op,
                value1: String::new(),
                value2: None,
            },
            Some("custom") => ValidationType::Custom {
                formula: String::new(),
            },
            _ => ValidationType::None,
        };

        validation
    }

    /// Apply formula values to a data validation based on its type
    fn apply_validation_formulas(
        validation: &mut DataValidation,
        formula1: Option<String>,
        formula2: Option<String>,
    ) {
        match &mut validation.validation_type {
            ValidationType::List { source } => {
                if let Some(f1) = formula1 {
                    // Remove surrounding quotes if present
                    *source = f1.trim_matches('"').to_string();
                }
            }
            ValidationType::Whole { value1, value2, .. }
            | ValidationType::Decimal { value1, value2, .. }
            | ValidationType::Date { value1, value2, .. }
            | ValidationType::Time { value1, value2, .. }
            | ValidationType::TextLength { value1, value2, .. } => {
                if let Some(f1) = formula1 {
                    *value1 = f1;
                }
                *value2 = formula2;
            }
            ValidationType::Custom { formula } => {
                if let Some(f1) = formula1 {
                    *formula = f1;
                }
            }
            ValidationType::None => {}
        }
    }

    /// Parse a color element from attributes (rgb, theme, tint, etc.)
    fn parse_color_element(
        e: &quick_xml::events::BytesStart,
        theme_palette: Option<&ThemePalette>,
    ) -> Color {
        // Priority: rgb > theme > indexed > auto
        let mut rgb: Option<String> = None;
        let mut theme: Option<u8> = None;
        let mut tint: Option<f64> = None;
        let mut indexed: Option<u8> = None;
        let mut auto = false;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
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

        if let Some(rgb_str) = rgb {
            let hex = rgb_str.trim_start_matches('#');

            if hex.len() == 8 {
                if let (Ok(_a), Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&hex[0..2], 16),
                    u8::from_str_radix(&hex[2..4], 16),
                    u8::from_str_radix(&hex[4..6], 16),
                    u8::from_str_radix(&hex[6..8], 16),
                ) {
                    // Keep worksheet-level CF/data bar color behavior consistent
                    // with existing tests by treating ARGB as RGB and ignoring alpha.
                    return Color::Rgb { r, g, b };
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
            if let Some(theme) = theme_palette {
                let (r, g, b) = theme.resolve_theme_color(index, tint_i8);
                return Color::Rgb { r, g, b };
            }
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

    /// Parse a conditional formatting rule from element attributes
    fn parse_cf_rule_attrs(
        e: &quick_xml::events::BytesStart,
        sqref: Option<&str>,
    ) -> ConditionalFormatRule {
        let mut rule = ConditionalFormatRule::default();
        let mut rule_type: Option<String> = None;
        let mut operator: Option<String> = None;
        let mut text: Option<String> = None;
        let mut rank: Option<u32> = None;
        let mut percent = false;
        let mut bottom = false;
        let mut above_average = true;
        let mut equal_average = false;
        let mut std_dev: Option<u32> = None;
        let mut time_period: Option<String> = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"type" => {
                    rule_type = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"operator" => {
                    operator = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"priority" => {
                    if let Some(p) = attr.unescape_value().ok().and_then(|s| s.parse().ok()) {
                        rule.priority = p;
                    }
                }
                b"stopIfTrue" => {
                    rule.stop_if_true = attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"dxfId" => {
                    rule.dxf_id = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                }
                b"text" => {
                    text = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"rank" => {
                    rank = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                }
                b"percent" => {
                    percent = attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"bottom" => {
                    bottom = attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"aboveAverage" => {
                    above_average = attr.unescape_value().ok().map_or(true, |s| s != "0");
                }
                b"equalAverage" => {
                    equal_average = attr.unescape_value().ok().map_or(false, |s| s == "1");
                }
                b"stdDev" => {
                    std_dev = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                }
                b"timePeriod" => {
                    time_period = attr.unescape_value().ok().map(|s| s.to_string());
                }
                _ => {}
            }
        }

        // Set ranges from sqref
        if let Some(sqref) = sqref {
            rule.ranges = Self::parse_sqref(sqref);
        }

        // Set the rule type based on parsed attributes
        let op = operator
            .as_deref()
            .and_then(CfOperator::from_xlsx)
            .unwrap_or(CfOperator::Equal);

        rule.rule_type = match rule_type.as_deref() {
            Some("cellIs") => CfRuleType::CellIs {
                operator: op,
                formula1: String::new(),
                formula2: None,
            },
            Some("expression") => CfRuleType::Expression {
                formula: String::new(),
            },
            Some("top10") => CfRuleType::Top10 {
                rank: rank.unwrap_or(10),
                percent,
                bottom,
            },
            Some("aboveAverage") => CfRuleType::AboveAverage {
                above: above_average,
                equal_average,
                std_dev,
            },
            Some("containsText") => CfRuleType::ContainsText {
                text: text.unwrap_or_default(),
            },
            Some("beginsWith") => CfRuleType::BeginsWith {
                text: text.unwrap_or_default(),
            },
            Some("endsWith") => CfRuleType::EndsWith {
                text: text.unwrap_or_default(),
            },
            Some("duplicateValues") => CfRuleType::DuplicateValues,
            Some("uniqueValues") => CfRuleType::UniqueValues,
            Some("containsBlanks") => CfRuleType::ContainsBlanks,
            Some("notContainsBlanks") => CfRuleType::NotContainsBlanks,
            Some("containsErrors") => CfRuleType::ContainsErrors,
            Some("notContainsErrors") => CfRuleType::NotContainsErrors,
            Some("timePeriod") => CfRuleType::TimePeriod {
                period: time_period
                    .as_deref()
                    .and_then(TimePeriod::from_xlsx)
                    .unwrap_or(TimePeriod::Today),
            },
            // ColorScale, DataBar, IconSet are handled separately via child elements
            _ => CfRuleType::Expression {
                formula: String::new(),
            },
        };

        rule
    }

    /// Apply formula values to a conditional format rule
    fn apply_cf_formulas(rule: &mut ConditionalFormatRule, formulas: &[String]) {
        match &mut rule.rule_type {
            CfRuleType::CellIs {
                formula1, formula2, ..
            } => {
                if let Some(f1) = formulas.first() {
                    *formula1 = f1.clone();
                }
                *formula2 = formulas.get(1).cloned();
            }
            CfRuleType::Expression { formula } => {
                if let Some(f1) = formulas.first() {
                    *formula = f1.clone();
                }
            }
            _ => {}
        }
    }

    /// Parse a space-separated sqref string into cell ranges
    fn parse_sqref(sqref: &str) -> Vec<CellRange> {
        sqref
            .split_whitespace()
            .filter_map(|s| CellRange::parse(s).ok())
            .collect()
    }

    /// Read comments for a worksheet from the comments XML file
    fn read_worksheet_comments<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        sheet_index: usize,
        worksheet: &mut duke_sheets_core::Worksheet,
    ) -> XlsxResult<()> {
        let visible_map = Self::read_comment_visibility_map(archive, sheet_index)?;

        // Try to read the comments file (may not exist)
        let comments_path = format!("xl/comments{}.xml", sheet_index + 1);
        let file = match archive.by_name(&comments_path) {
            Ok(f) => f,
            Err(_) => return Ok(()), // No comments file is valid
        };

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut authors: Vec<String> = Vec::new();

        // Current comment parsing state
        let mut in_author = false;
        let mut in_comment = false;
        let mut in_text = false;
        let mut in_t = false;
        let mut current_ref: Option<String> = None;
        let mut current_author_id: Option<usize> = None;
        let mut current_text = String::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                    b"author" => {
                        in_author = true;
                    }
                    b"comment" => {
                        in_comment = true;
                        current_ref = None;
                        current_author_id = None;
                        current_text.clear();

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"ref" => {
                                    current_ref = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"authorId" => {
                                    current_author_id =
                                        attr.unescape_value().ok().and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }
                    }
                    b"text" if in_comment => {
                        in_text = true;
                    }
                    b"t" if in_text => {
                        in_t = true;
                    }
                    // Also handle <r> (rich text run) elements
                    b"r" if in_text => {}
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                    b"author" => {
                        in_author = false;
                    }
                    b"comment" => {
                        // Add the comment to the worksheet
                        if let Some(ref cell_ref) = current_ref {
                            match CellAddress::parse(cell_ref) {
                                Ok(addr) => {
                                    let author = current_author_id
                                        .and_then(|id| authors.get(id))
                                        .cloned()
                                        .unwrap_or_default();

                                    let visible = visible_map
                                        .get(&(addr.row, addr.col))
                                        .copied()
                                        .unwrap_or(false);
                                    let comment = CellComment::new(author, current_text.trim())
                                        .with_visible(visible);
                                    worksheet.set_comment_at(addr.row, addr.col, comment);
                                }
                                Err(e) => log::warn!("Skipping comment at '{}': {}", cell_ref, e),
                            }
                        }
                        in_comment = false;
                        current_text.clear();
                    }
                    b"text" => {
                        in_text = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_author {
                        if let Ok(text) = e.unescape() {
                            authors.push(text.to_string());
                        }
                    } else if in_t {
                        if let Ok(text) = e.unescape() {
                            // Append to current text (may have multiple <t> elements in rich text)
                            if !current_text.is_empty() {
                                current_text.push(' ');
                            }
                            current_text.push_str(&text);
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    // Handle self-closing elements
                    if e.name().local_name().as_ref() == b"author" {
                        // Empty author element
                        authors.push(String::new());
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    fn read_comment_visibility_map<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        sheet_index: usize,
    ) -> XlsxResult<HashMap<(u32, u16), bool>> {
        let vml_path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
        let file = match archive.by_name(&vml_path) {
            Ok(f) => f,
            Err(_) => return Ok(HashMap::new()),
        };

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut map: HashMap<(u32, u16), bool> = HashMap::new();

        let mut in_shape = false;
        let mut current_visible = false;
        let mut in_client_data_note = false;
        let mut in_row = false;
        let mut in_col = false;
        let mut current_row: Option<u32> = None;
        let mut current_col: Option<u16> = None;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                    b"shape" => {
                        in_shape = true;
                        current_visible = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"style" {
                                if let Some(style) =
                                    attr.unescape_value().ok().map(|s| s.to_lowercase())
                                {
                                    current_visible = style.contains("visibility:visible");
                                }
                            }
                        }
                    }
                    b"ClientData" if in_shape => {
                        let mut is_note = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"ObjectType" {
                                is_note = attr.unescape_value().ok().as_deref() == Some("Note");
                            }
                        }
                        if is_note {
                            in_client_data_note = true;
                            current_row = None;
                            current_col = None;
                        }
                    }
                    b"Row" if in_client_data_note => {
                        in_row = true;
                    }
                    b"Column" if in_client_data_note => {
                        in_col = true;
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                    b"shape" => {
                        in_shape = false;
                        in_client_data_note = false;
                        in_row = false;
                        in_col = false;
                        current_row = None;
                        current_col = None;
                    }
                    b"ClientData" if in_client_data_note => {
                        if let (Some(r), Some(c)) = (current_row, current_col) {
                            map.insert((r, c), current_visible);
                        }
                        in_client_data_note = false;
                        in_row = false;
                        in_col = false;
                        current_row = None;
                        current_col = None;
                    }
                    b"Row" => in_row = false,
                    b"Column" => in_col = false,
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    if in_row {
                        if let Some(v) = e.unescape().ok().and_then(|s| s.parse::<u32>().ok()) {
                            current_row = Some(v);
                        }
                    } else if in_col {
                        if let Some(v) = e.unescape().ok().and_then(|s| s.parse::<u16>().ok()) {
                            current_col = Some(v);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(map)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::BytesStart;
    use std::io::{Cursor, Write};

    fn build_single_sheet_xlsx(sheet_xml: &str) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(sheet_xml.as_bytes()).unwrap();

            zip.finish().unwrap();
        }
        buf
    }

    #[test]
    fn test_parse_color_element_theme_and_tint() {
        let mut e = BytesStart::new("color");
        e.push_attribute(("theme", "4"));
        e.push_attribute(("tint", "0.5"));

        assert_eq!(
            XlsxReader::parse_color_element(&e, None),
            Color::Theme { index: 4, tint: 50 }
        );
    }

    #[test]
    fn test_parse_color_element_indexed_and_auto() {
        let mut indexed = BytesStart::new("color");
        indexed.push_attribute(("indexed", "12"));
        assert_eq!(
            XlsxReader::parse_color_element(&indexed, None),
            Color::Indexed(12)
        );

        let mut auto = BytesStart::new("color");
        auto.push_attribute(("auto", "1"));
        assert_eq!(XlsxReader::parse_color_element(&auto, None), Color::Auto);
    }

    #[test]
    fn test_parse_color_element_theme_with_palette_resolves_to_rgb() {
        let mut e = BytesStart::new("color");
        e.push_attribute(("theme", "4"));
        e.push_attribute(("tint", "0.5"));

        let palette = ThemePalette::default();
        assert_eq!(
            XlsxReader::parse_color_element(&e, Some(&palette)),
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

        let palette = XlsxReader::parse_theme_palette(Cursor::new(xml.as_bytes())).unwrap();
        assert_eq!(palette.colors[4], (0x11, 0x22, 0x33));
        assert_eq!(palette.resolve_theme_color(4, 0), (0x11, 0x22, 0x33));
    }

    #[test]
    fn test_read_shared_formula_master_and_follower() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>1</v></c>
      <c r="B1" t="n"><v>2</v></c>
      <c r="C1"><f t="shared" si="0">A1+B1</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="n"><v>4</v></c>
      <c r="B2" t="n"><v>5</v></c>
      <c r="C2"><f t="shared" si="0"/><v>9</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_value("C1").unwrap().formula_text(),
            Some("=A1+B1")
        );
        assert_eq!(
            sheet.get_value("C2").unwrap().formula_text(),
            Some("=A2+B2")
        );
    }

    #[test]
    fn test_read_shared_formula_preserves_absolute_and_shifts_ranges() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>1</v></c>
      <c r="B1" t="n"><v>2</v></c>
      <c r="D1"><f t="shared" si="1">SUM($A$1:B1)+LEN("A1")</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="n"><v>4</v></c>
      <c r="B2" t="n"><v>5</v></c>
      <c r="D2"><f t="shared" si="1"/><v>6</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_value("D1").unwrap().formula_text(),
            Some("=SUM($A$1:B1)+LEN(\"A1\")")
        );
        assert_eq!(
            sheet.get_value("D2").unwrap().formula_text(),
            Some("=SUM($A$1:B2)+LEN(\"A1\")")
        );
    }

    #[test]
    fn test_read_array_formula_anchor() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A3">ROW(A1:A3)</f><v>1</v></c>
    </row>
    <row r="2"><c r="A2"><v>2</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_value("A1").unwrap().formula_text(),
            Some("=ROW(A1:A3)")
        );
    }

    #[test]
    fn test_read_datatable_formula_placeholder() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="dataTable" ref="A1:B2" r1="C1" r2="C2"/><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_value("A1").unwrap().formula_text(),
            Some("=TABLE(C1,C2)")
        );
        assert_eq!(sheet.get_value("A1").unwrap().as_number(), Some(42.0));
    }

    #[test]
    fn test_read_sheet_view_selected_and_freeze_panes() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0" tabSelected="1" zoomScale="125">
      <pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="D5" sqref="D5:E6"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert!(sheet.is_selected());
        assert_eq!(
            sheet.freeze_panes().map(|fp| (fp.row, fp.col)),
            Some((3, 2))
        );
        assert_eq!(sheet.zoom_scale(), Some(125));
        assert_eq!(sheet.selection_active_cell(), Some((4, 3)));
        assert_eq!(
            sheet.selection_range().map(|r| r.to_string()),
            Some("D5:E6".to_string())
        );
    }

    #[test]
    fn test_read_sheet_view_split_panes() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0" zoomScale="90">
      <pane xSplit="2000" ySplit="3000" topLeftCell="C4" activePane="bottomRight" state="split"/>
      <selection pane="bottomRight" activeCell="D5" sqref="D5"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let split = sheet.split_panes().expect("split panes should exist");
        assert_eq!(split.x_split, 2000.0);
        assert_eq!(split.y_split, 3000.0);
        assert_eq!(split.top_left, Some((3, 2)));
        assert_eq!(split.active_pane.as_deref(), Some("bottomRight"));
        assert_eq!(sheet.zoom_scale(), Some(90));
        assert_eq!(sheet.selection_active_cell(), Some((4, 3)));
        assert_eq!(
            sheet.selection_range().map(|r| r.to_string()),
            Some("D5".to_string())
        );
    }

    #[test]
    fn test_read_outline_and_collapsed_row_col_attrs() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2" outlineLevel="2" collapsed="1"><c r="A2" t="n"><v>1</v></c></row>
  </sheetData>
  <cols>
    <col min="3" max="3" outlineLevel="3" collapsed="1"/>
  </cols>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(sheet.row_outline_level(1), 2);
        assert!(sheet.is_row_collapsed(1));
        assert_eq!(sheet.column_outline_level(2), 3);
        assert!(sheet.is_column_collapsed(2));
    }

    #[test]
    fn test_read_page_setup_margins_print_and_header_footer() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageMargins left="0.5" right="0.6" top="0.7" bottom="0.8" header="0.2" footer="0.25"/>
  <pageSetup paperSize="9" orientation="landscape" scale="85" fitToWidth="1" fitToHeight="2"/>
  <printOptions gridLines="1" headings="1"/>
  <headerFooter>
    <oddHeader>&amp;LLeft&amp;CCenter</oddHeader>
    <oddFooter>&amp;RPage &amp;P</oddFooter>
  </headerFooter>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let ps = sheet.page_setup();

        assert!((ps.left_margin - 0.5).abs() < 1e-9);
        assert!((ps.right_margin - 0.6).abs() < 1e-9);
        assert!((ps.top_margin - 0.7).abs() < 1e-9);
        assert!((ps.bottom_margin - 0.8).abs() < 1e-9);
        assert!((ps.header_margin - 0.2).abs() < 1e-9);
        assert!((ps.footer_margin - 0.25).abs() < 1e-9);
        assert_eq!(ps.paper_size, 9);
        assert!(matches!(
            ps.orientation,
            duke_sheets_core::PageOrientation::Landscape
        ));
        assert_eq!(ps.scale, 85);
        assert_eq!(ps.fit_to_width, Some(1));
        assert_eq!(ps.fit_to_height, Some(2));
        assert!(ps.print_gridlines);
        assert!(ps.print_headings);
        assert_eq!(ps.odd_header.as_deref(), Some("&LLeft&CCenter"));
        assert_eq!(ps.odd_footer.as_deref(), Some("&RPage &P"));
    }

    #[test]
    fn test_read_comment_visibility_map_from_vml() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:Row>1</x:Row>
      <x:Column>2</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;visibility:hidden">
    <x:ClientData ObjectType="Note">
      <x:Row>3</x:Row>
      <x:Column>4</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let map = XlsxReader::read_comment_visibility_map(&mut archive, 0).unwrap();

        assert_eq!(map.get(&(1, 2)).copied(), Some(true));
        assert_eq!(map.get(&(3, 4)).copied(), Some(false));
    }

    #[test]
    fn test_read_comments_applies_visibility_from_vml() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData></worksheet>"#).unwrap();

            zip.start_file("xl/comments1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>John</author></authors><commentList><comment ref="C2" authorId="0"><text><r><t>Visible note</t></r></text></comment></commentList></comments>"#).unwrap();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:Row>1</x:Row>
      <x:Column>2</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let comment = sheet.comment("C2").unwrap().expect("comment should exist");
        assert_eq!(comment.author, "John");
        assert_eq!(comment.text, "Visible note");
        assert!(comment.visible);
    }

    #[test]
    fn test_decode_excel_escapes_carriage_return() {
        assert_eq!(decode_excel_escapes("hello_x000d_world"), "hello\rworld");
    }

    #[test]
    fn test_decode_excel_escapes_line_feed() {
        assert_eq!(decode_excel_escapes("hello_x000a_world"), "hello\nworld");
    }

    #[test]
    fn test_decode_excel_escapes_tab() {
        assert_eq!(decode_excel_escapes("col1_x0009_col2"), "col1\tcol2");
    }

    #[test]
    fn test_decode_excel_escapes_multiple() {
        assert_eq!(
            decode_excel_escapes("line1_x000d__x000a_line2"),
            "line1\r\nline2"
        );
    }

    #[test]
    fn test_decode_excel_escapes_underscore() {
        // _x005f_ is an escaped underscore
        assert_eq!(decode_excel_escapes("under_x005f_score"), "under_score");
    }

    #[test]
    fn test_decode_excel_escapes_no_escapes() {
        assert_eq!(decode_excel_escapes("plain text"), "plain text");
    }

    #[test]
    fn test_decode_excel_escapes_partial_sequence() {
        // Incomplete sequences should be left as-is
        assert_eq!(decode_excel_escapes("_x00"), "_x00");
        assert_eq!(decode_excel_escapes("_x000"), "_x000");
        assert_eq!(decode_excel_escapes("_x000d"), "_x000d"); // missing trailing _
    }

    #[test]
    fn test_decode_excel_escapes_uppercase() {
        // Should handle uppercase hex digits
        assert_eq!(decode_excel_escapes("_x000D_"), "\r");
        assert_eq!(decode_excel_escapes("_x000A_"), "\n");
    }

    #[test]
    fn test_decode_excel_escapes_real_world() {
        // Real example from the Cardex file
        assert_eq!(
            decode_excel_escapes("D. Potenziani_x000d__x000d_RD1237 Quality Hold"),
            "D. Potenziani\r\rRD1237 Quality Hold"
        );
    }

    #[test]
    fn test_read_empty_xlsx() {
        // Minimal valid XLSX structure
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            // [Content_Types].xml
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#).unwrap();

            // _rels/.rels
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            // xl/workbook.xml
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            // xl/_rels/workbook.xml.rels
            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            // xl/worksheets/sheet1.xml
            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>"#).unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(buf);
        let workbook = XlsxReader::read(cursor).unwrap();

        assert_eq!(workbook.sheet_count(), 1);
        assert_eq!(workbook.worksheet(0).unwrap().name(), "Sheet1");
    }
}
