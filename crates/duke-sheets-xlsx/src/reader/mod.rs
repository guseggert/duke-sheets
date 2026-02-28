//! XLSX reader

use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::{read_styles_xml, register_roundtrip_style_data, ParsedStyles};
use comments::read_worksheet_comments;
use conditional_format::{
    apply_cf_formulas, parse_cf_rule_attrs, parse_color_element, parse_sqref,
};
use data_validation::{apply_validation_formulas, parse_data_validation_attrs};
use duke_sheets_core::conditional_format::{
    CfColorValue, CfRuleType, CfValue, CfValueType, ConditionalFormatRule, IconSetStyle,
};
use duke_sheets_core::style::{Color, Style};
use duke_sheets_core::validation::DataValidation;
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue, Hyperlink, SplitPanes, Workbook,
};
use formulas::{parse_cell_formula_state, resolve_cell_formula, SharedFormulaMaster};
use theme::{read_theme_palette, resolve_style_theme_colors};

mod comments;
mod conditional_format;
mod data_validation;
mod formulas;
mod theme;

pub(crate) use formulas::CellFormulaState;
pub(crate) use theme::ThemePalette;

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

#[derive(Debug, Clone)]
struct SheetRelationship {
    rel_type: String,
    target: String,
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
        let theme_palette = read_theme_palette(&mut archive, workbook_rels.theme_path.as_deref())?;
        if let Some(theme) = theme_palette {
            for style in &mut parsed_styles.cell_styles {
                resolve_style_theme_colors(style, &theme);
            }
            for style in &mut parsed_styles.dxf_styles {
                resolve_style_theme_colors(style, &theme);
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
                let sheet_rels = Self::read_sheet_rels(&mut archive, path)?;
                Self::read_worksheet(
                    &mut archive,
                    path,
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                    &shared_strings,
                    &cell_styles,
                    &dxf_styles,
                    theme_palette.as_ref(),
                    &sheet_rels,
                )?;

                // Read comments for this worksheet (if present)
                let comments_path = format!("xl/comments{}.xml", idx + 1);
                let vml_path = format!("xl/drawings/vmlDrawing{}.vml", idx + 1);
                read_worksheet_comments(
                    &mut archive,
                    &comments_path,
                    Some(&vml_path),
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

    fn read_sheet_rels<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        sheet_path: &str,
    ) -> XlsxResult<HashMap<String, SheetRelationship>> {
        let (base_dir, file_name) = match sheet_path.rsplit_once('/') {
            Some((dir, file)) => (dir, file),
            None => return Ok(HashMap::new()),
        };
        let rels_path = format!("{}/_rels/{}.rels", base_dir, file_name);

        let file = match archive.by_name(&rels_path) {
            Ok(f) => f,
            Err(_) => return Ok(HashMap::new()),
        };

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut rels = HashMap::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.name().local_name().as_ref() == b"Relationship" =>
                {
                    let mut id = None;
                    let mut target = None;
                    let mut rel_type = None;
                    let mut target_mode = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Id" => id = attr.unescape_value().ok().map(|s| s.to_string()),
                            b"Target" => target = attr.unescape_value().ok().map(|s| s.to_string()),
                            b"Type" => rel_type = attr.unescape_value().ok().map(|s| s.to_string()),
                            b"TargetMode" => {
                                target_mode = attr.unescape_value().ok().map(|s| s.to_string())
                            }
                            _ => {}
                        }
                    }

                    if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                        let resolved_target = if target.starts_with('/')
                            || target_mode.as_deref() == Some("External")
                        {
                            target
                        } else {
                            let mut parts: Vec<&str> = base_dir.split('/').collect();
                            for part in target.split('/') {
                                if part == ".." {
                                    parts.pop();
                                } else if part != "." && !part.is_empty() {
                                    parts.push(part);
                                }
                            }
                            parts.join("/")
                        };

                        rels.insert(
                            id,
                            SheetRelationship {
                                rel_type,
                                target: resolved_target,
                            },
                        );
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
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
        sheet_rels: &HashMap<String, SheetRelationship>,
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
                    b"hyperlink" => {
                        Self::parse_hyperlink_element(worksheet, &e, sheet_rels);
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
                        current_formula_state = parse_cell_formula_state(&e);
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
                        current_validation = Some(parse_data_validation_attrs(&e));
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
                        current_cf_rule = Some(parse_cf_rule_attrs(&e, cf_sqref.as_deref()));
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
                                let resolved_formula = resolve_cell_formula(
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
                                apply_validation_formulas(
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
                                apply_cf_formulas(&mut rule, &cf_formulas);
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
                        b"hyperlink" => {
                            Self::parse_hyperlink_element(worksheet, &e, sheet_rels);
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
                            current_formula_state = parse_cell_formula_state(&e);
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
                            let color = parse_color_element(&e, theme_palette);
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

    fn parse_hyperlink_element(
        worksheet: &mut duke_sheets_core::Worksheet,
        e: &quick_xml::events::BytesStart<'_>,
        sheet_rels: &HashMap<String, SheetRelationship>,
    ) {
        let mut cell_ref = None;
        let mut rel_id = None;
        let mut display = None;
        let mut tooltip = None;
        let mut location = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"ref" => cell_ref = attr.unescape_value().ok().map(|s| s.to_string()),
                b"id" => rel_id = attr.unescape_value().ok().map(|s| s.to_string()),
                b"display" => display = attr.unescape_value().ok().map(|s| s.to_string()),
                b"tooltip" => tooltip = attr.unescape_value().ok().map(|s| s.to_string()),
                b"location" => location = attr.unescape_value().ok().map(|s| s.to_string()),
                _ => {}
            }
        }

        let cell_ref = match cell_ref {
            Some(v) => v,
            None => return,
        };

        let cell_a1 = match CellAddress::parse(&cell_ref) {
            Ok(addr) => addr.to_a1_string(),
            Err(_) => match CellRange::parse(&cell_ref) {
                Ok(range) => range.start.to_a1_string(),
                Err(_) => {
                    log::warn!("Invalid hyperlink ref '{}', skipping", cell_ref);
                    return;
                }
            },
        };

        let mut target = String::new();
        if let Some(rel_id) = rel_id {
            if let Some(rel) = sheet_rels.get(&rel_id) {
                if rel.rel_type.ends_with("/hyperlink") {
                    target = rel.target.clone();
                }
            }
        }

        if target.is_empty() {
            if let Some(loc) = &location {
                target = format!("#{}", loc);
            }
        }

        let hyperlink = Hyperlink {
            target,
            display,
            tooltip,
            location,
        };

        if let Err(err) = worksheet.set_hyperlink(&cell_a1, hyperlink) {
            log::warn!("Failed to set hyperlink for {}: {}", cell_a1, err);
        }
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
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    fn build_single_sheet_xlsx(sheet_xml: &str) -> Vec<u8> {
        build_single_sheet_xlsx_with_sheet_rels(sheet_xml, None)
    }

    fn build_single_sheet_xlsx_with_sheet_rels(
        sheet_xml: &str,
        sheet_rels_xml: Option<&str>,
    ) -> Vec<u8> {
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

            if let Some(sheet_rels_xml) = sheet_rels_xml {
                zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                    .unwrap();
                zip.write_all(sheet_rels_xml.as_bytes()).unwrap();
            }

            zip.finish().unwrap();
        }
        buf
    }

    #[test]
    fn test_read_hyperlinks_external_internal_and_tooltip() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="B2" t="s"><v>1</v></c></row>
    <row r="3"><c r="C3" t="s"><v>2</v></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1" display="Example"/>
    <hyperlink ref="B2" location="Sheet2!A1" display="Go to Sheet2"/>
    <hyperlink ref="C3" r:id="rId2" tooltip="Tooltip here"/>
  </hyperlinks>
</worksheet>"#;

        let sheet_rels_xml = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/path" TargetMode="External"/>
</Relationships>"#;

        let bytes = build_single_sheet_xlsx_with_sheet_rels(sheet_xml, Some(sheet_rels_xml));
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let a1 = sheet.hyperlink("A1").expect("A1 hyperlink");
        assert_eq!(a1.target, "https://example.com");
        assert_eq!(a1.display.as_deref(), Some("Example"));
        assert_eq!(a1.tooltip, None);

        let b2 = sheet.hyperlink("B2").expect("B2 hyperlink");
        assert_eq!(b2.target, "#Sheet2!A1");
        assert_eq!(b2.location.as_deref(), Some("Sheet2!A1"));
        assert_eq!(b2.display.as_deref(), Some("Go to Sheet2"));

        let c3 = sheet.hyperlink("C3").expect("C3 hyperlink");
        assert_eq!(c3.target, "https://example.org/path");
        assert_eq!(c3.tooltip.as_deref(), Some("Tooltip here"));
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
