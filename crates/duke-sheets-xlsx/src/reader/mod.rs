//! XLSX reader

use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::{read_styles_xml, ParsedStyles};
use duke_sheets_core::comment::CellComment;
use duke_sheets_core::conditional_format::{
    CfColorValue, CfOperator, CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
    IconSetStyle, TimePeriod,
};
use duke_sheets_core::style::{Color, Style};
use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::{CellAddress, CellError, CellRange, CellValue, Workbook};

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
        let parsed_styles = Self::read_styles(&mut archive)?;
        let cell_styles = parsed_styles.cell_styles;
        let dxf_styles = parsed_styles.dxf_styles;

        // Read workbook.xml to get sheet info
        let sheet_info = Self::read_workbook_xml(&mut archive)?;

        // Read workbook.xml.rels to get sheet paths
        let sheet_paths = Self::read_workbook_rels(&mut archive)?;

        // Create workbook
        let mut workbook = Workbook::empty();

        // Read each worksheet
        for (idx, (name, r_id)) in sheet_info.iter().enumerate() {
            if let Some(path) = sheet_paths.get(r_id) {
                let sheet_idx = workbook.add_worksheet_with_name(name)?;
                Self::read_worksheet(
                    &mut archive,
                    path,
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                    &shared_strings,
                    &cell_styles,
                    &dxf_styles,
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
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut current_string = String::new();
        let mut in_si = false;
        let mut in_t = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_string.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => match e.name().as_ref() {
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
                Ok(Event::Text(e)) if in_t => {
                    if let Ok(text) = e.unescape() {
                        current_string.push_str(&text);
                    }
                }
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
                    dxf_styles: Vec::new(),
                })
            }
        };
        read_styles_xml(file)
    }

    /// Read workbook.xml to get sheet names and rIds
    fn read_workbook_xml<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> XlsxResult<Vec<(String, String)>> {
        let file = archive
            .by_name("xl/workbook.xml")
            .map_err(|_| XlsxError::MissingPart("xl/workbook.xml".into()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut sheets = Vec::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) if e.name().as_ref() == b"sheet" => {
                    let mut name = None;
                    let mut r_id = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"r:id" => {
                                r_id = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            _ => {}
                        }
                    }

                    if let (Some(name), Some(r_id)) = (name, r_id) {
                        sheets.push((name, r_id));
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(sheets)
    }

    /// Read workbook.xml.rels to get sheet file paths
    fn read_workbook_rels<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> XlsxResult<HashMap<String, String>> {
        let file = archive
            .by_name("xl/_rels/workbook.xml.rels")
            .map_err(|_| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".into()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.trim_text(true);

        let mut buf = Vec::new();
        let mut rels = HashMap::new();

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.name().as_ref() == b"Relationship" =>
                {
                    let mut id = None;
                    let mut target = None;
                    let mut rel_type = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
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

                    // Only include worksheet relationships
                    if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                        if rel_type.ends_with("/worksheet") {
                            // Target is relative to xl/ folder
                            let full_path = if target.starts_with('/') {
                                target[1..].to_string()
                            } else {
                                format!("xl/{}", target)
                            };
                            rels.insert(id, full_path);
                        }
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
    ) -> XlsxResult<()> {
        let file = archive
            .by_name(path)
            .map_err(|_| XlsxError::MissingPart(path.to_string()))?;

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.trim_text(true);

        let mut buf = Vec::new();

        // Current cell state
        let mut current_cell_ref: Option<String> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_style: Option<u32> = None;
        let mut current_value: Option<String> = None;
        let mut current_formula: Option<String> = None;
        let mut in_cell = false;
        let mut in_value = false;
        let mut in_formula = false;
        let mut in_inline_str = false;
        let mut in_inline_text = false;

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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"row" => {
                        // Parse row dimensions: ht, customHeight, hidden
                        let mut row_num: Option<u32> = None;
                        let mut ht: Option<f64> = None;
                        let mut custom_height = false;
                        let mut hidden = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
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
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        current_cell_ref = None;
                        current_cell_type = None;
                        current_cell_style = None;
                        current_value = None;
                        current_formula = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
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
                            if attr.key.as_ref() == b"sqref" {
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
                            if attr.key.as_ref() == b"showValue" {
                                data_bar_show_value =
                                    attr.unescape_value().ok().map_or(true, |s| s != "0");
                            }
                        }
                    }
                    b"iconSet" if in_cf_rule => {
                        in_icon_set = true;
                        // Parse iconSet attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
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
                    match e.name().as_ref() {
                        b"c" => {
                            // Process the cell
                            if let Some(ref cell_ref) = current_cell_ref {
                                Self::process_cell(
                                    worksheet,
                                    cell_ref,
                                    current_cell_type.as_deref(),
                                    current_value.as_deref(),
                                    current_formula.as_deref(),
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
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_value {
                        if let Ok(text) = e.unescape() {
                            current_value = Some(text.to_string());
                        }
                    } else if in_formula {
                        if let Ok(text) = e.unescape() {
                            current_formula = Some(text.to_string());
                        }
                    } else if in_inline_text {
                        if let Ok(text) = e.unescape() {
                            // Inline string - store directly as value
                            current_value = Some(text.to_string());
                            current_cell_type = Some("inlineStr".to_string());
                        }
                    } else if in_dv_formula1 {
                        if let Ok(text) = e.unescape() {
                            dv_formula1 = Some(text.to_string());
                        }
                    } else if in_dv_formula2 {
                        if let Ok(text) = e.unescape() {
                            dv_formula2 = Some(text.to_string());
                        }
                    } else if in_cf_formula {
                        if let Ok(text) = e.unescape() {
                            cf_formulas.push(text.to_string());
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    match e.name().as_ref() {
                        b"row" => {
                            // Self-closing <row .../> with no cells — may have dimensions
                            let mut row_num: Option<u32> = None;
                            let mut ht: Option<f64> = None;
                            let mut custom_height = false;
                            let mut hidden = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
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
                            }
                        }
                        b"col" => {
                            // Parse column dimensions: min, max, width, customWidth, hidden
                            let mut col_min: Option<u16> = None;
                            let mut col_max: Option<u16> = None;
                            let mut width: Option<f64> = None;
                            let mut custom_width = false;
                            let mut hidden = false;
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
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
                                }
                            }
                        }
                        b"c" => {
                            // Empty cell element (may still carry a style)
                            let mut cell_ref: Option<String> = None;
                            let mut cell_type: Option<String> = None;
                            let mut cell_style: Option<u32> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
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
                                match attr.key.as_ref() {
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
                            let color = Self::parse_color_element(&e);
                            if in_color_scale {
                                cf_colors.push(color);
                            } else if in_data_bar {
                                data_bar_color = Some(color);
                            }
                        }
                        // Merged cells
                        b"mergeCell" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ref" {
                                    let ref_str = String::from_utf8_lossy(&attr.value);
                                    if let Ok(range) = CellRange::parse(&ref_str) {
                                        let _ = worksheet.merge_cells(&range);
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
        let addr = CellAddress::parse(cell_ref).map_err(|e| {
            XlsxError::Parse(format!("Invalid cell reference '{}': {}", cell_ref, e))
        })?;

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

            worksheet.set_cell_value_at(
                addr.row,
                addr.col,
                CellValue::Formula {
                    text: formula_text,
                    cached_value: cached.map(Box::new),
                    array_result: None,
                },
            )?;
        } else if let Some(value) = value {
            // Process value based on type
            let cell_value = match cell_type {
                // Shared string
                Some("s") => {
                    let idx: usize = value.parse().map_err(|_| {
                        XlsxError::Parse(format!("Invalid shared string index: {}", value))
                    })?;
                    let s = shared_strings.get(idx).ok_or_else(|| {
                        XlsxError::Parse(format!("Shared string index {} out of bounds", idx))
                    })?;
                    CellValue::String(s.clone().into())
                }

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

            worksheet.set_cell_value_at(addr.row, addr.col, cell_value)?;
        }

        // Apply style (if any)
        if let Some(s) = style_idx {
            if s != 0 {
                let style = styles
                    .get(s as usize)
                    .ok_or_else(|| XlsxError::Parse(format!("Style index {} out of bounds", s)))?;
                worksheet.set_cell_style_at(addr.row, addr.col, style)?;
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
            match attr.key.as_ref() {
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
    fn parse_color_element(e: &quick_xml::events::BytesStart) -> Color {
        let mut rgb: Option<String> = None;

        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"rgb" {
                rgb = attr.unescape_value().ok().map(|s| s.to_string());
            }
            // TODO: Handle theme colors and tint
        }

        if let Some(rgb_str) = rgb {
            // Parse ARGB hex string (e.g., "FF638EC6")
            if rgb_str.len() >= 6 {
                // Skip alpha if present
                let hex = if rgb_str.len() == 8 {
                    &rgb_str[2..]
                } else {
                    &rgb_str
                };
                if let (Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&hex[0..2], 16),
                    u8::from_str_radix(&hex[2..4], 16),
                    u8::from_str_radix(&hex[4..6], 16),
                ) {
                    return Color::rgb(r, g, b);
                }
            }
        }

        // Default color if parsing fails
        Color::rgb(0, 0, 0)
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
            match attr.key.as_ref() {
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
        // Try to read the comments file (may not exist)
        let comments_path = format!("xl/comments{}.xml", sheet_index + 1);
        let file = match archive.by_name(&comments_path) {
            Ok(f) => f,
            Err(_) => return Ok(()), // No comments file is valid
        };

        let reader = BufReader::new(file);
        let mut xml_reader = Reader::from_reader(reader);
        xml_reader.trim_text(true);

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
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"author" => {
                        in_author = true;
                    }
                    b"comment" => {
                        in_comment = true;
                        current_ref = None;
                        current_author_id = None;
                        current_text.clear();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
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
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"author" => {
                        in_author = false;
                    }
                    b"comment" => {
                        // Add the comment to the worksheet
                        if let Some(ref cell_ref) = current_ref {
                            if let Ok(addr) = CellAddress::parse(cell_ref) {
                                let author = current_author_id
                                    .and_then(|id| authors.get(id))
                                    .cloned()
                                    .unwrap_or_default();

                                let comment = CellComment::new(author, current_text.trim());
                                worksheet.set_comment_at(addr.row, addr.col, comment);
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
                    if e.name().as_ref() == b"author" {
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
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

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
