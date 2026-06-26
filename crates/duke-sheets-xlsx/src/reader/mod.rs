//! XLSX reader

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::{
    read_styles_xml, register_roundtrip_style_data, register_roundtrip_theme_data, ParsedStyles,
};
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
    CellAddress, CellError, CellRange, CellValue, Hyperlink, PageBreak, SheetSlot, SplitPanes,
    Workbook,
};
use formulas::{
    parse_cell_formula_state, resolve_cell_formula, CellFormulaKind, SharedFormulaMaster,
};
use theme::{read_theme_palette, resolve_style_theme_colors};

mod archive;
pub(crate) mod chart;
pub(crate) mod chart_ex;
mod chartsheet;
mod comments;
mod conditional_format;
mod data_validation;
mod drawing;
mod formulas;
mod pivot;
mod shared_strings;
mod table;
mod theme;
mod workbook;

pub(crate) use archive::archive_by_name;
pub(crate) use formulas::CellFormulaState;
use shared_strings::SharedStringEntry;
pub(crate) use theme::ThemePalette;
use workbook::{read_sheet_rels, read_workbook_rels, read_workbook_xml, SheetRelationship};

/// Resolve a relative path from a drawing's .rels against the drawing's own path.

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

fn read_chart_style_color<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    chart_path: &str,
    chart: &mut duke_sheets_chart::Chart,
) {
    let chart_rels = match workbook::read_sheet_rels(archive, chart_path) {
        Ok(r) => r,
        Err(_) => return,
    };
    for rel in chart_rels.values() {
        if rel.rel_type.ends_with("/chartStyle") {
            if let Ok(mut f) = archive_by_name(archive, &rel.target) {
                let mut bytes = Vec::new();
                if f.read_to_end(&mut bytes).is_ok() {
                    chart.raw_chart_style = Some(bytes);
                }
            }
        } else if rel.rel_type.ends_with("/chartColorStyle") {
            if let Ok(mut f) = archive_by_name(archive, &rel.target) {
                let mut bytes = Vec::new();
                if f.read_to_end(&mut bytes).is_ok() {
                    chart.raw_chart_color_style = Some(bytes);
                }
            }
        }
    }
}

fn read_chart_style_color_for_chart_ex<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    chart_path: &str,
    chart: &mut duke_sheets_chart::ChartEx,
) {
    let chart_rels = match workbook::read_sheet_rels(archive, chart_path) {
        Ok(r) => r,
        Err(_) => return,
    };
    for rel in chart_rels.values() {
        if rel.rel_type.ends_with("/chartStyle") {
            if let Ok(mut f) = archive_by_name(archive, &rel.target) {
                let mut bytes = Vec::new();
                if f.read_to_end(&mut bytes).is_ok() {
                    chart.raw_chart_style = Some(bytes);
                }
            }
        } else if rel.rel_type.ends_with("/chartColorStyle") {
            if let Ok(mut f) = archive_by_name(archive, &rel.target) {
                let mut bytes = Vec::new();
                if f.read_to_end(&mut bytes).is_ok() {
                    chart.raw_chart_color_style = Some(bytes);
                }
            }
        }
    }
}

/// XLSX file reader
pub struct XlsxReader;

/// CFB magic: encrypted XLSX files are CFB envelopes rather than ZIPs,
/// so a leading match here means we should run the bytes through the
/// crypto decrypt path before treating them as a ZIP archive.
fn is_cfb_envelope(bytes: &[u8]) -> bool {
    bytes.len() >= 8 && bytes[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
}

/// Open an encrypted-OOXML CFB envelope, extract the EncryptionInfo and
/// EncryptedPackage streams, and decrypt to the inner ZIP bytes.
///
/// Distinguishes three failure modes via the error message prefix so the
/// top-level dispatcher can decide whether to fall back to the XLS path:
///
///   - `"CFB envelope open failed:"` and `"not an OOXML envelope"` mean
///     "this isn't an encrypted-OOXML container" — fall through to XLS.
///   - Other errors (Encrypted, BadPassword, malformed envelope) indicate
///     the file IS encrypted OOXML and should propagate as-is.
fn decrypt_ooxml_envelope(
    bytes: &[u8],
    password: &str,
    skip_integrity_check: bool,
) -> XlsxResult<Vec<u8>> {
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(std::io::Cursor::new(bytes))
        .map_err(|e| XlsxError::InvalidFormat(format!("CFB envelope open failed: {e}")))?;
    let has_info = cfb.exists("/EncryptionInfo");
    let has_package = cfb.exists("/EncryptedPackage");
    if !has_info && !has_package {
        return Err(XlsxError::InvalidFormat(
            "not an OOXML envelope: CFB has neither /EncryptionInfo nor /EncryptedPackage".into(),
        ));
    }
    let info = cfb.read_stream("/EncryptionInfo").map_err(|e| {
        XlsxError::InvalidFormat(format!(
            "malformed OOXML envelope: read /EncryptionInfo: {e}"
        ))
    })?;
    let package = cfb.read_stream("/EncryptedPackage").map_err(|e| {
        XlsxError::InvalidFormat(format!(
            "malformed OOXML envelope: read /EncryptedPackage: {e}"
        ))
    })?;
    Ok(duke_sheets_crypto::ooxml::decrypt_with_options(
        &info,
        &package,
        password,
        &duke_sheets_crypto::ooxml::DecryptOptions {
            skip_integrity_check,
        },
    )?)
}

impl XlsxReader {
    /// Read a workbook from a file path
    pub fn read_file<P: AsRef<Path>>(path: P) -> XlsxResult<Workbook> {
        let file = File::open(path)?;
        Self::read(file)
    }

    /// Read a workbook from a file path, supplying a password for
    /// encrypted files. When `password` is `None` and
    /// `try_velvet_sweatshop` is true, encrypted files are
    /// transparently retried with the well-known `VelvetSweatshop`
    /// password before reporting them as encrypted.
    pub fn read_file_with_password<P: AsRef<Path>>(
        path: P,
        password: Option<&str>,
        try_velvet_sweatshop: bool,
    ) -> XlsxResult<Workbook> {
        Self::read_file_with_options(path, password, try_velvet_sweatshop, false)
    }

    /// Read a workbook from a file path with full open-options control.
    /// `skip_integrity_check` opts out of the post-decrypt HMAC check;
    /// default-off (false) matches Office.
    pub fn read_file_with_options<P: AsRef<Path>>(
        path: P,
        password: Option<&str>,
        try_velvet_sweatshop: bool,
        skip_integrity_check: bool,
    ) -> XlsxResult<Workbook> {
        let bytes = std::fs::read(path)?;
        Self::read_bytes_with_options(&bytes, password, try_velvet_sweatshop, skip_integrity_check)
    }

    /// Read a workbook from raw bytes with an optional password.
    ///
    /// Encrypted XLSX files are CFB envelopes (not plain ZIPs); when
    /// the leading magic bytes match CFB we delegate to
    /// `duke_sheets_crypto::ooxml::decrypt` and then proceed with the
    /// resulting plaintext ZIP.
    pub fn read_bytes_with_password(
        bytes: &[u8],
        password: Option<&str>,
        try_velvet_sweatshop: bool,
    ) -> XlsxResult<Workbook> {
        Self::read_bytes_with_options(bytes, password, try_velvet_sweatshop, false)
    }

    /// Read a workbook from raw bytes with full open-options control.
    /// `skip_integrity_check` opts out of the post-decrypt HMAC check.
    pub fn read_bytes_with_options(
        bytes: &[u8],
        password: Option<&str>,
        try_velvet_sweatshop: bool,
        skip_integrity_check: bool,
    ) -> XlsxResult<Workbook> {
        if is_cfb_envelope(bytes) {
            let try_pw = match password {
                Some(p) => p,
                None if try_velvet_sweatshop => "VelvetSweatshop",
                None => {
                    return Err(XlsxError::Encrypted(
                        "workbook is encrypted but no password was supplied".into(),
                    ));
                }
            };
            return match decrypt_ooxml_envelope(bytes, try_pw, skip_integrity_check) {
                Ok(decrypted) => Self::read(std::io::Cursor::new(decrypted)),
                Err(XlsxError::BadPassword) if password.is_none() => Err(XlsxError::Encrypted(
                    "workbook is encrypted but no password was supplied".into(),
                )),
                Err(e) => Err(e),
            };
        }
        Self::read(std::io::Cursor::new(bytes))
    }

    /// Read a workbook from a reader
    pub fn read<R: Read + Seek>(reader: R) -> XlsxResult<Workbook> {
        let mut archive = zip::ZipArchive::new(reader)?;

        // Verify this is an XLSX file
        if archive_by_name(&mut archive, "[Content_Types].xml").is_err() {
            return Err(XlsxError::InvalidFormat(
                "Missing [Content_Types].xml".into(),
            ));
        }

        // Read shared strings (if present)
        let shared_strings = shared_strings::read_shared_strings(&mut archive)?;

        // Read styles (if present)
        let mut parsed_styles = Self::read_styles(&mut archive)?;
        let roundtrip_style_data = parsed_styles.roundtrip_data();
        // Read workbook.xml.rels to get sheet/theme paths
        let workbook_rels = read_workbook_rels(&mut archive)?;
        // Read workbook theme (if present) and resolve theme colors in styles
        let (theme_palette, raw_theme_xml) =
            read_theme_palette(&mut archive, workbook_rels.theme_path.as_deref())?;
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
        let wb_props = read_workbook_xml(&mut archive)?;

        let mut pivot_caches = HashMap::new();
        for cache_entry in &wb_props.pivot_caches {
            if let Some(path) = workbook_rels.pivot_cache_paths.get(&cache_entry.r_id) {
                if let Some(cache) =
                    pivot::read_pivot_cache_definition(&mut archive, cache_entry.cache_id, path)?
                {
                    pivot_caches.insert(cache_entry.cache_id, cache);
                }
            }
        }

        let sheet_paths = workbook_rels.sheet_paths;
        let chartsheet_paths = workbook_rels.chartsheet_paths;

        // Create workbook
        let mut workbook = Workbook::empty();
        workbook.settings_mut().date_1904 = wb_props.date_1904;

        // Add named ranges
        for nr in wb_props.named_ranges {
            workbook.named_ranges_mut().define_or_update(nr);
        }

        let sheet_info = &wb_props.sheets;
        let date_1904 = wb_props.date_1904;

        // Read each sheet in tab-bar order (single pass over sheet_info).
        // Worksheets and chartsheets are interleaved in workbook.xml <sheets>;
        // we store them in separate Vecs but record ordering in sheet_order.
        let mut ws_count: usize = 0;
        for (_idx, sheet_entry) in sheet_info.iter().enumerate() {
            if let Some(path) = sheet_paths.get(&sheet_entry.r_id) {
                let sheet_idx = workbook.add_worksheet_with_name_unchecked(&sheet_entry.name);
                ws_count += 1;
                workbook
                    .worksheet_mut(sheet_idx)
                    .unwrap()
                    .set_date_1904(date_1904);
                workbook
                    .worksheet_mut(sheet_idx)
                    .unwrap()
                    .set_visibility(sheet_entry.visibility);
                workbook
                    .sheet_order_mut()
                    .push(SheetSlot::Worksheet(sheet_idx));
                let sheet_rels = read_sheet_rels(&mut archive, path)?;
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

                // Read comments for this worksheet (if present).
                // Resolve paths via sheet .rels relationships; fall back to
                // index-based filenames for files that lack .rels entries.
                let comments_path = sheet_rels
                    .values()
                    .find(|r| r.rel_type.ends_with("/comments"))
                    .map(|r| r.target.clone())
                    .unwrap_or_else(|| format!("xl/comments{}.xml", ws_count));
                let vml_path = sheet_rels
                    .values()
                    .find(|r| r.rel_type.ends_with("/vmlDrawing"))
                    .map(|r| r.target.clone())
                    .unwrap_or_else(|| format!("xl/drawings/vmlDrawing{}.vml", ws_count));
                read_worksheet_comments(
                    &mut archive,
                    &comments_path,
                    Some(&vml_path),
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                )?;

                // Read tables for this worksheet (if present).
                // Each relationship with type ending in "/table" points to
                // an xl/tables/tableN.xml part.
                let mut table_rels: Vec<String> = sheet_rels
                    .values()
                    .filter(|r| r.rel_type.ends_with("/table"))
                    .map(|r| r.target.clone())
                    .collect();
                table_rels.sort();
                for table_path in &table_rels {
                    if let Some(t) = table::read_table(&mut archive, table_path)? {
                        workbook.worksheet_mut(sheet_idx).unwrap().add_table(t);
                    }
                }

                let mut pivot_rels: Vec<String> = sheet_rels
                    .values()
                    .filter(|r| r.rel_type.ends_with("/pivotTable"))
                    .map(|r| r.target.clone())
                    .collect();
                pivot_rels.sort();
                for pivot_path in &pivot_rels {
                    if let Some(pivot) =
                        pivot::read_pivot_table(&mut archive, pivot_path, &pivot_caches)?
                    {
                        workbook
                            .worksheet_mut(sheet_idx)
                            .unwrap()
                            .add_pivot_table(pivot)
                            .map_err(|e| XlsxError::InvalidFormat(e.to_string()))?;
                    }
                }

                // Read charts for this worksheet (if present).
                // Sheet → drawing relationship → drawing XML → chart relationships → chart XML.
                for drawing_rel in sheet_rels
                    .values()
                    .filter(|r| r.rel_type.ends_with("/drawing"))
                {
                    let drawing_path = &drawing_rel.target;
                    let drawing_contents =
                        drawing::read_drawing_contents(&mut archive, drawing_path)?;
                    for raw in drawing_contents.raw_non_chart_anchors {
                        workbook
                            .worksheet_mut(sheet_idx)
                            .unwrap()
                            .raw_drawing_objects
                            .push(raw);
                    }
                    let has_charts = !drawing_contents.chart_refs.is_empty();
                    let has_images = !drawing_contents.images.is_empty();
                    if !has_charts && !has_images {
                        continue;
                    }
                    let drawing_rels = read_sheet_rels(&mut archive, drawing_path)?;
                    if has_images {
                        let mut resolved_images = drawing_contents.images;
                        for image in &mut resolved_images {
                            if let Some(rel) = drawing_rels.get(&image.media_path) {
                                let ext = rel.target.rsplit('.').next().unwrap_or("");
                                if let Some(fmt) =
                                    duke_sheets_chart::ImageFormat::from_extension(ext)
                                {
                                    image.format = fmt;
                                }
                                image.media_path = rel.target.clone();
                                if let Ok(mut f) = archive_by_name(&mut archive, &rel.target) {
                                    let mut buf = Vec::new();
                                    if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
                                        image.data = buf;
                                    }
                                }
                            }
                            if let Some(svg_rel_id) = &image.svg_media_path {
                                if let Some(rel) = drawing_rels.get(svg_rel_id.as_str()) {
                                    image.svg_media_path = Some(rel.target.clone());
                                    if let Ok(mut f) = archive_by_name(&mut archive, &rel.target) {
                                        let mut buf = Vec::new();
                                        if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
                                            image.svg_data = Some(buf);
                                        }
                                    }
                                }
                            }
                        }
                        let ws = workbook.worksheet_mut(sheet_idx).unwrap();
                        for img in resolved_images {
                            ws.add_image(img);
                        }
                    }
                    for chart_ref in drawing_contents.chart_refs {
                        if let Some(dr) = drawing_rels.get(&chart_ref.rel_id) {
                            if chart_ref.is_chart_ex {
                                if let Some(mut cx) = chart_ex::read_chart_ex(
                                    &mut archive,
                                    &dr.target,
                                    chart_ref.anchor,
                                )? {
                                    cx.raw_mc_fallback = chart_ref.raw_mc_fallback;
                                    read_chart_style_color_for_chart_ex(
                                        &mut archive,
                                        &dr.target,
                                        &mut cx,
                                    );
                                    workbook.worksheet_mut(sheet_idx).unwrap().add_chart_ex(cx);
                                }
                            } else if let Some(mut c) =
                                chart::read_chart(&mut archive, &dr.target, chart_ref.anchor)?
                            {
                                read_chart_style_color(&mut archive, &dr.target, &mut c);
                                workbook.worksheet_mut(sheet_idx).unwrap().add_chart(c);
                            }
                        }
                    }
                }
            } else if let Some(cs_path) = chartsheet_paths.get(&sheet_entry.r_id) {
                let mut chart_found = false;
                let drawing_rid = chartsheet::read_chartsheet_drawing_rid(&mut archive, cs_path)?;
                if let Some(rid) = drawing_rid {
                    let cs_rels = read_sheet_rels(&mut archive, cs_path)?;
                    if let Some(drawing_rel) = cs_rels.get(&rid) {
                        let drawing_path = &drawing_rel.target;
                        let drawing_contents =
                            drawing::read_drawing_contents(&mut archive, drawing_path)?;
                        let raw_anchors = drawing_contents.raw_non_chart_anchors;
                        let drawing_rels = read_sheet_rels(&mut archive, drawing_path)?;
                        for chart_ref in drawing_contents.chart_refs {
                            if let Some(dr) = drawing_rels.get(&chart_ref.rel_id) {
                                if chart_ref.is_chart_ex {
                                    // ChartEx in a chartsheet - skip for now (chartsheets
                                    // require a standard Chart). Parse it but don't embed.
                                    continue;
                                }
                                if let Some(mut c) =
                                    chart::read_chart(&mut archive, &dr.target, chart_ref.anchor)?
                                {
                                    read_chart_style_color(&mut archive, &dr.target, &mut c);
                                    let cs_idx = workbook.add_chartsheet_unchecked(
                                        duke_sheets_core::ChartSheet {
                                            name: sheet_entry.name.clone(),
                                            chart: c,
                                            visibility: sheet_entry.visibility,
                                            raw_drawing_objects: raw_anchors.clone(),
                                        },
                                    );
                                    workbook
                                        .sheet_order_mut()
                                        .push(SheetSlot::ChartSheet(cs_idx));
                                    chart_found = true;
                                    break; // chartsheet has exactly one chart
                                }
                            }
                        }
                    }
                }
                if !chart_found {
                    let cs_idx = workbook.add_chartsheet_unchecked(duke_sheets_core::ChartSheet {
                        name: sheet_entry.name.clone(),
                        chart: duke_sheets_chart::Chart::new(
                            duke_sheets_chart::ChartType::Unsupported("missing".into()),
                        ),
                        visibility: sheet_entry.visibility,
                        raw_drawing_objects: Vec::new(),
                    });
                    workbook
                        .sheet_order_mut()
                        .push(SheetSlot::ChartSheet(cs_idx));
                }
            }
        }

        // Apply print area and print titles from named ranges to worksheets.
        Self::apply_print_settings(&mut workbook);

        // Ensure at least one sheet exists
        if workbook.sheet_count() == 0 && workbook.chartsheet_count() == 0 {
            workbook.add_worksheet()?;
        }

        register_roundtrip_style_data(&workbook, roundtrip_style_data);
        if let Some(theme_bytes) = raw_theme_xml {
            register_roundtrip_theme_data(&workbook, theme_bytes);
        }

        Ok(workbook)
    }

    fn read_styles<R: Read + Seek>(archive: &mut zip::ZipArchive<R>) -> XlsxResult<ParsedStyles> {
        let file = match archive_by_name(archive, "xl/styles.xml") {
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

    /// Extract _xlnm.Print_Area and _xlnm.Print_Titles from named ranges
    /// and apply them to the corresponding worksheets.
    fn apply_print_settings(workbook: &mut Workbook) {
        let sheet_count = workbook.sheet_count();
        for sheet_idx in 0..sheet_count {
            let sheet_name = workbook.worksheet(sheet_idx).map(|s| s.name().to_string());
            let sheet_name = match sheet_name {
                Some(n) => n,
                None => continue,
            };

            if let Some(nr) = workbook.named_ranges().get_exact(
                "_xlnm.Print_Area",
                &duke_sheets_core::named_range::NameScope::Sheet(sheet_idx),
            ) {
                if let Some(range) = parse_print_area_formula(&nr.refers_to, &sheet_name) {
                    if let Some(ws) = workbook.worksheet_mut(sheet_idx) {
                        ws.set_print_area(range);
                    }
                }
            }

            if let Some(nr) = workbook.named_ranges().get_exact(
                "_xlnm.Print_Titles",
                &duke_sheets_core::named_range::NameScope::Sheet(sheet_idx),
            ) {
                let (rows, cols) = parse_print_titles_formula(&nr.refers_to, &sheet_name);
                if let Some(ws) = workbook.worksheet_mut(sheet_idx) {
                    if let Some((r1, r2)) = rows {
                        ws.set_repeat_rows(r1, r2);
                    }
                    if let Some((c1, c2)) = cols {
                        ws.set_repeat_cols(c1, c2);
                    }
                }
            }
        }
    }

    /// Read a worksheet from the archive
    #[allow(clippy::too_many_arguments)]
    fn read_worksheet<R: Read + Seek>(
        archive: &mut zip::ZipArchive<R>,
        path: &str,
        worksheet: &mut duke_sheets_core::Worksheet,
        shared_strings: &[SharedStringEntry],
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
        xml_reader.config_mut().trim_text(false);

        let mut buf = Vec::new();

        // Current cell state
        let mut current_cell_ref: Option<String> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_style: Option<u32> = None;
        let mut current_cell_cm: Option<u32> = None;
        let mut current_value: Option<String> = None;
        let mut current_formula: Option<String> = None;
        let mut current_formula_state = CellFormulaState::default();
        let mut in_cell = false;
        let mut in_value = false;
        let mut in_formula = false;
        let mut in_inline_str = false;
        let mut in_inline_text = false;
        // Inline rich text state
        let mut in_inline_r = false;
        let mut in_inline_rpr = false;
        let mut in_inline_run_t = false;
        let mut has_inline_runs = false;
        let mut inline_runs: Vec<duke_sheets_core::RichTextRun> = Vec::new();
        let mut inline_run_text = String::new();
        let mut inline_run_font: Option<duke_sheets_core::RunFont> = None;
        let mut shared_formula_masters: HashMap<u32, SharedFormulaMaster> = HashMap::new();

        // Pending array/dataTable formulas with ref ranges - post-processed after all cells.
        // Each entry: (anchor_cell_ref, ref_range, kind, formula_text, r1, r2)
        #[allow(clippy::type_complexity)]
        let mut pending_array_formulas: Vec<(
            String,
            String,
            CellFormulaKind,
            String,
            Option<String>,
            Option<String>,
        )> = Vec::new();

        // Dynamic array tracking: cm="1" anchors and cm="2" ghost cells.
        // Anchor: (row, col). Ghost: set of (row, col) positions.
        let mut dynamic_array_anchors: Vec<(u32, u16)> = Vec::new();
        let mut dynamic_array_ghosts: HashSet<(u32, u16)> = HashSet::new();

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
        let mut in_even_header = false;
        let mut in_even_footer = false;
        let mut in_first_header = false;
        let mut in_first_footer = false;
        let mut in_row_breaks = false;
        let mut in_col_breaks = false;

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

        // AutoFilter state
        let mut in_auto_filter = false;
        let mut auto_filter_range: Option<CellRange> = None;
        let mut auto_filter_columns: Vec<duke_sheets_core::FilterColumn> = Vec::new();
        let mut current_af_col_id: Option<u32> = None;
        let mut current_af_hidden_button = false;
        let mut current_af_show_button = true;
        let mut current_af_filter_values: Vec<String> = Vec::new();
        let mut current_af_blank = false;
        let mut current_af_custom_and = false;
        let mut current_af_custom_conditions: Vec<duke_sheets_core::CustomFilterCondition> =
            Vec::new();
        let mut current_af_column_filter: Option<duke_sheets_core::ColumnFilter> = None;
        let mut in_af_filters = false;
        let mut in_af_custom_filters = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    match e.name().local_name().as_ref() {
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
                                    b"showFormulas" => {}
                                    b"showGridLines" => {}
                                    b"showZeros" => {}
                                    b"showRuler" => {}
                                    b"showOutlineSymbols" => {}
                                    b"showRowColHeaders" => {}
                                    b"showWhiteSpace" => {}
                                    b"topLeftCell" => {}
                                    b"view" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            match v.as_ref() {
                                                "normal" => {}
                                                "pageBreakPreview" => {}
                                                "pageLayout" => {}
                                                _ => {}
                                            }
                                        }
                                    }
                                    b"windowProtection" => {}
                                    b"workbookViewId" => {}
                                    b"rightToLeft" => {}
                                    b"zoomScaleNormal" => {}
                                    b"zoomScalePageLayoutView" => {}
                                    b"zoomScaleSheetLayoutView" => {}
                                    b"colorId" => {}
                                    _ => {}
                                }
                            }
                        }
                        b"selection" => Self::parse_sheet_selection_attrs(&e, worksheet),
                        b"pane" => Self::parse_pane_attrs(&e, worksheet),
                        b"autoFilter" => {
                            in_auto_filter = true;
                            auto_filter_range = None;
                            auto_filter_columns.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    if let Ok(value) = attr.unescape_value() {
                                        match CellRange::parse(value.as_ref()) {
                                            Ok(range) => auto_filter_range = Some(range),
                                            Err(err) => log::warn!(
                                                "Invalid autoFilter ref '{}': {}",
                                                value,
                                                err
                                            ),
                                        }
                                    }
                                }
                            }
                        }
                        b"filterColumn" if in_auto_filter => {
                            current_af_col_id = None;
                            current_af_hidden_button = false;
                            current_af_show_button = true;
                            current_af_filter_values.clear();
                            current_af_blank = false;
                            current_af_custom_and = false;
                            current_af_custom_conditions.clear();
                            current_af_column_filter = None;
                            in_af_filters = false;
                            in_af_custom_filters = false;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"colId" => {
                                        current_af_col_id = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    b"hiddenButton" => {
                                        current_af_hidden_button =
                                            attr.unescape_value().ok().is_some_and(|s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"showButton" => {
                                        current_af_show_button =
                                            attr.unescape_value().ok().is_none_or(|s| {
                                                !(s.as_ref() == "0" || s.as_ref() == "false")
                                            });
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"filters" if in_auto_filter => {
                            in_af_filters = true;
                            current_af_filter_values.clear();
                            current_af_blank = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"blank" {
                                    current_af_blank = attr
                                        .unescape_value()
                                        .ok()
                                        .is_some_and(|s| s.as_ref() == "1" || s.as_ref() == "true");
                                }
                            }
                        }
                        b"customFilters" if in_auto_filter => {
                            in_af_custom_filters = true;
                            current_af_custom_conditions.clear();
                            current_af_custom_and = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"and" {
                                    current_af_custom_and = attr
                                        .unescape_value()
                                        .ok()
                                        .is_some_and(|s| s.as_ref() == "1" || s.as_ref() == "true");
                                }
                            }
                        }
                        b"hyperlink" => {
                            Self::parse_hyperlink_element(worksheet, &e, sheet_rels);
                        }
                        b"sheetProtection" => {
                            Self::parse_sheet_protection_element(worksheet, &e);
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
                                        if let Ok(v) = attr.unescape_value() {
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
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.print_gridlines =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"headings" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.print_headings =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"headerFooter" => {
                            let mut ps = worksheet.page_setup().clone();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"differentOddEven" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.different_odd_even =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"differentFirst" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.different_first =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"scaleWithDoc" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.scale_with_doc =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"alignWithMargins" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.align_with_margins =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"rowBreaks" => in_row_breaks = true,
                        b"colBreaks" => in_col_breaks = true,
                        b"brk" if in_row_breaks || in_col_breaks => {
                            if let Some(brk) = Self::parse_page_break_attrs(&e) {
                                if in_row_breaks {
                                    let mut breaks = worksheet.row_breaks().to_vec();
                                    breaks.push(brk);
                                    worksheet.set_row_breaks(breaks);
                                } else {
                                    let mut breaks = worksheet.col_breaks().to_vec();
                                    breaks.push(brk);
                                    worksheet.set_col_breaks(breaks);
                                }
                            }
                        }
                        b"oddHeader" => in_odd_header = true,
                        b"oddFooter" => in_odd_footer = true,
                        b"evenHeader" => in_even_header = true,
                        b"evenFooter" => in_even_footer = true,
                        b"firstHeader" => in_first_header = true,
                        b"firstFooter" => in_first_footer = true,
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
                                        custom_height =
                                            attr.unescape_value().ok().is_some_and(|s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value().ok().is_some_and(|s| {
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
                                        collapsed = attr.unescape_value().ok().is_some_and(|s| {
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
                            current_cell_cm = None;
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
                                    b"cm" => {
                                        current_cell_cm = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"v" if in_cell => in_value = true,
                        b"f" if in_cell => {
                            current_formula_state = parse_cell_formula_state(&e);
                            in_formula = true;
                        }
                        b"is" if in_cell => {
                            in_inline_str = true;
                            has_inline_runs = false;
                            inline_runs.clear();
                        }
                        b"r" if in_inline_str => {
                            in_inline_r = true;
                            has_inline_runs = true;
                            inline_run_text.clear();
                            inline_run_font = None;
                        }
                        b"rPr" if in_inline_r => {
                            in_inline_rpr = true;
                            inline_run_font = Some(duke_sheets_core::RunFont::default());
                        }
                        b"t" if in_inline_r && !in_inline_rpr => in_inline_run_t = true,
                        b"t" if in_inline_str && !in_inline_r => in_inline_text = true,
                        // rPr children that use Start+End (rare, but handle defensively)
                        name if in_inline_rpr => {
                            shared_strings::parse_rpr_element(name, &e, &mut inline_run_font);
                        }
                        b"dataValidation" => {
                            in_data_validation = true;
                            dv_formula1 = None;
                            dv_formula2 = None;
                            current_validation = Some(parse_data_validation_attrs(&e));
                        }
                        b"formula1" if in_data_validation => in_dv_formula1 = true,
                        b"formula2" if in_data_validation => in_dv_formula2 = true,
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
                        b"formula" if in_cf_rule => in_cf_formula = true,
                        b"colorScale" if in_cf_rule => in_color_scale = true,
                        b"dataBar" if in_cf_rule => {
                            in_data_bar = true;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"showValue" {
                                    data_bar_show_value =
                                        attr.unescape_value().ok().is_none_or(|s| s != "0");
                                }
                            }
                        }
                        b"iconSet" if in_cf_rule => {
                            in_icon_set = true;
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
                                            attr.unescape_value().ok().is_some_and(|s| s == "1");
                                    }
                                    b"showValue" => {
                                        icon_set_show_value =
                                            attr.unescape_value().ok().is_none_or(|s| s != "0");
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => {
                    match e.name().local_name().as_ref() {
                        b"c" => {
                            // Process the cell
                            if let Some(ref cell_ref) = current_cell_ref {
                                if has_inline_runs {
                                    // Inline rich text - set directly, bypassing process_cell
                                    if let Ok(addr) = CellAddress::parse(cell_ref) {
                                        let runs = std::mem::take(&mut inline_runs);
                                        let resolved_formula = resolve_cell_formula(
                                            cell_ref,
                                            current_formula.as_deref(),
                                            &current_formula_state,
                                            &mut shared_formula_masters,
                                        );
                                        if let Some(f) = resolved_formula {
                                            let formula_text = if f.starts_with('=') {
                                                f
                                            } else {
                                                format!("={}", f)
                                            };
                                            if let Err(e) = worksheet
                                                .set_formula_with_cached_value_at(
                                                    addr.row,
                                                    addr.col,
                                                    &formula_text,
                                                    CellValue::rich_text(runs),
                                                )
                                            {
                                                log::warn!(
                                                    "Skipping rich text formula {}: {}",
                                                    cell_ref,
                                                    e
                                                );
                                            }
                                        } else if let Err(e) = worksheet.set_cell_value_at(
                                            addr.row,
                                            addr.col,
                                            CellValue::rich_text(runs),
                                        ) {
                                            log::warn!(
                                                "Skipping rich text cell {}: {}",
                                                cell_ref,
                                                e
                                            );
                                        }
                                        // Apply style
                                        if let Some(s) = current_cell_style {
                                            if s != 0 {
                                                if let Some(style) = cell_styles.get(s as usize) {
                                                    if let Err(e) = worksheet.set_cell_style_at(
                                                        addr.row, addr.col, style,
                                                    ) {
                                                        log::warn!(
                                                            "Cell {}: failed to apply style: {}",
                                                            cell_ref,
                                                            e
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    has_inline_runs = false;
                                } else {
                                    if let Some(cm) = current_cell_cm {
                                        if let Ok(addr) = CellAddress::parse(cell_ref) {
                                            match cm {
                                                1 => {
                                                    dynamic_array_anchors.push((addr.row, addr.col))
                                                }
                                                2 => {
                                                    dynamic_array_ghosts
                                                        .insert((addr.row, addr.col));
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
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
                            }
                            // Record array/dataTable formulas with ref for post-processing
                            if let Some(ref cell_ref) = current_cell_ref {
                                if let Some(ref array_ref) = current_formula_state.array_ref {
                                    if current_formula_state.kind == CellFormulaKind::Array
                                        || current_formula_state.kind == CellFormulaKind::DataTable
                                    {
                                        let formula_text =
                                            current_formula.clone().unwrap_or_default();
                                        pending_array_formulas.push((
                                            cell_ref.clone(),
                                            array_ref.clone(),
                                            current_formula_state.kind,
                                            formula_text,
                                            current_formula_state.data_table_input1_ref.clone(),
                                            current_formula_state.data_table_input2_ref.clone(),
                                        ));
                                    }
                                }
                            }
                            in_cell = false;
                        }
                        b"v" => in_value = false,
                        b"f" => in_formula = false,
                        b"rPr" if in_inline_rpr => in_inline_rpr = false,
                        b"t" if in_inline_run_t => in_inline_run_t = false,
                        b"r" if in_inline_r => {
                            // Finish current run - mirrors SST parser
                            let font = inline_run_font.take().and_then(|f| {
                                if f.is_empty() {
                                    None
                                } else {
                                    Some(f)
                                }
                            });
                            inline_runs.push(duke_sheets_core::RichTextRun {
                                text: decode_excel_escapes(&std::mem::take(&mut inline_run_text)),
                                font,
                            });
                            in_inline_r = false;
                        }
                        b"is" => in_inline_str = false,
                        b"t" if in_inline_str => in_inline_text = false,
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
                        b"formula1" if in_data_validation => in_dv_formula1 = false,
                        b"formula2" if in_data_validation => in_dv_formula2 = false,
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
                        b"formula" if in_cf_rule => in_cf_formula = false,
                        b"oddHeader" => in_odd_header = false,
                        b"oddFooter" => in_odd_footer = false,
                        b"evenHeader" => in_even_header = false,
                        b"evenFooter" => in_even_footer = false,
                        b"firstHeader" => in_first_header = false,
                        b"firstFooter" => in_first_footer = false,
                        b"rowBreaks" => in_row_breaks = false,
                        b"colBreaks" => in_col_breaks = false,
                        b"filters" if in_auto_filter => {
                            in_af_filters = false;
                            current_af_column_filter =
                                Some(duke_sheets_core::ColumnFilter::Values(
                                    duke_sheets_core::ValueFilter {
                                        values: std::mem::take(&mut current_af_filter_values),
                                        blank: current_af_blank,
                                    },
                                ));
                        }
                        b"customFilters" if in_auto_filter => {
                            in_af_custom_filters = false;
                            current_af_column_filter =
                                Some(duke_sheets_core::ColumnFilter::Custom(
                                    duke_sheets_core::CustomFilters {
                                        and: current_af_custom_and,
                                        conditions: std::mem::take(
                                            &mut current_af_custom_conditions,
                                        ),
                                    },
                                ));
                        }
                        b"filterColumn" if in_auto_filter => {
                            if let Some(col_id) = current_af_col_id {
                                let filter = current_af_column_filter.take().unwrap_or_else(|| {
                                    duke_sheets_core::ColumnFilter::Values(
                                        duke_sheets_core::ValueFilter {
                                            values: Vec::new(),
                                            blank: false,
                                        },
                                    )
                                });
                                auto_filter_columns.push(duke_sheets_core::FilterColumn {
                                    col_id,
                                    hidden_button: current_af_hidden_button,
                                    show_button: current_af_show_button,
                                    filter,
                                });
                            } else {
                                log::warn!("Skipping <filterColumn> without required colId");
                            }
                            current_af_col_id = None;
                            current_af_hidden_button = false;
                            current_af_show_button = true;
                            current_af_filter_values.clear();
                            current_af_blank = false;
                            current_af_custom_and = false;
                            current_af_custom_conditions.clear();
                            current_af_column_filter = None;
                            in_af_filters = false;
                            in_af_custom_filters = false;
                        }
                        b"autoFilter" => {
                            in_auto_filter = false;
                            if let Some(range) = auto_filter_range.take() {
                                worksheet.set_auto_filter(Some(duke_sheets_core::AutoFilter {
                                    range,
                                    filter_columns: std::mem::take(&mut auto_filter_columns),
                                }));
                            } else {
                                log::warn!("Skipping <autoFilter> without valid ref");
                                auto_filter_columns.clear();
                            }
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
                    } else if in_inline_run_t {
                        match e.unescape() {
                            Ok(text) => inline_run_text.push_str(&text),
                            Err(err) => log::warn!(
                                "Cell {:?}: inline run text unescape failed: {}",
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
                    } else if in_even_header {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.even_header = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    } else if in_even_footer {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.even_footer = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    } else if in_first_header {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.first_header = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    } else if in_first_footer {
                        if let Ok(text) = e.unescape() {
                            let mut ps = worksheet.page_setup().clone();
                            ps.first_footer = Some(text.to_string());
                            worksheet.set_page_setup(ps);
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    // Handle rPr children for inline rich text
                    if in_inline_rpr {
                        shared_strings::parse_rpr_element(
                            e.name().local_name().as_ref(),
                            &e,
                            &mut inline_run_font,
                        );
                        buf.clear();
                        continue;
                    }
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
                                        if let Ok(v) = attr.unescape_value() {
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
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.print_gridlines =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    b"headings" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            ps.print_headings =
                                                v.as_ref() == "1" || v.as_ref() == "true";
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            worksheet.set_page_setup(ps);
                        }
                        b"brk" if in_row_breaks || in_col_breaks => {
                            if let Some(brk) = Self::parse_page_break_attrs(&e) {
                                if in_row_breaks {
                                    let mut breaks = worksheet.row_breaks().to_vec();
                                    breaks.push(brk);
                                    worksheet.set_row_breaks(breaks);
                                } else {
                                    let mut breaks = worksheet.col_breaks().to_vec();
                                    breaks.push(brk);
                                    worksheet.set_col_breaks(breaks);
                                }
                            }
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
                                    b"showFormulas" => {}
                                    b"showGridLines" => {}
                                    b"showZeros" => {}
                                    b"showRuler" => {}
                                    b"showOutlineSymbols" => {}
                                    b"showRowColHeaders" => {}
                                    b"showWhiteSpace" => {}
                                    b"topLeftCell" => {}
                                    b"view" => {
                                        if let Ok(v) = attr.unescape_value() {
                                            match v.as_ref() {
                                                "normal" => {}
                                                "pageBreakPreview" => {}
                                                "pageLayout" => {}
                                                _ => {}
                                            }
                                        }
                                    }
                                    b"windowProtection" => {}
                                    b"workbookViewId" => {}
                                    b"rightToLeft" => {}
                                    b"zoomScaleNormal" => {}
                                    b"zoomScalePageLayoutView" => {}
                                    b"zoomScaleSheetLayoutView" => {}
                                    b"colorId" => {}
                                    _ => {}
                                }
                            }
                        }
                        b"selection" => Self::parse_sheet_selection_attrs(&e, worksheet),
                        b"hyperlink" => {
                            Self::parse_hyperlink_element(worksheet, &e, sheet_rels);
                        }
                        b"sheetProtection" => {
                            Self::parse_sheet_protection_element(worksheet, &e);
                        }
                        b"pane" => Self::parse_pane_attrs(&e, worksheet),
                        b"f" if in_cell => {
                            // Self-closing formula elements appear for shared formula
                            // follower cells: <f t="shared" si="0"/>
                            current_formula_state = parse_cell_formula_state(&e);
                            in_formula = false;
                        }
                        b"row" => {
                            // Self-closing <row .../> with no cells - may have dimensions
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
                                            attr.unescape_value().ok().is_some_and(|s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value().ok().is_some_and(|s| {
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
                                        collapsed = attr.unescape_value().ok().is_some_and(|s| {
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
                                            attr.unescape_value().ok().is_some_and(|s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"hidden" => {
                                        hidden = attr.unescape_value().ok().is_some_and(|s| {
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
                                        collapsed = attr.unescape_value().ok().is_some_and(|s| {
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
                        b"autoFilter" => {
                            let mut range: Option<CellRange> = None;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"ref" {
                                    if let Ok(value) = attr.unescape_value() {
                                        match CellRange::parse(value.as_ref()) {
                                            Ok(parsed) => range = Some(parsed),
                                            Err(err) => log::warn!(
                                                "Invalid autoFilter ref '{}': {}",
                                                value,
                                                err
                                            ),
                                        }
                                    }
                                }
                            }

                            if let Some(range) = range {
                                worksheet.set_auto_filter(Some(duke_sheets_core::AutoFilter::new(
                                    range,
                                )));
                            } else {
                                log::warn!("Skipping self-closing <autoFilter> without valid ref");
                            }
                        }
                        b"filterColumn" if in_auto_filter => {
                            let mut col_id: Option<u32> = None;
                            let mut hidden_button = false;
                            let mut show_button = true;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"colId" => {
                                        col_id = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    b"hiddenButton" => {
                                        hidden_button =
                                            attr.unescape_value().ok().is_some_and(|s| {
                                                s.as_ref() == "1" || s.as_ref() == "true"
                                            });
                                    }
                                    b"showButton" => {
                                        show_button = attr.unescape_value().ok().is_none_or(|s| {
                                            !(s.as_ref() == "0" || s.as_ref() == "false")
                                        });
                                    }
                                    _ => {}
                                }
                            }

                            if let Some(col_id) = col_id {
                                auto_filter_columns.push(duke_sheets_core::FilterColumn {
                                    col_id,
                                    hidden_button,
                                    show_button,
                                    filter: duke_sheets_core::ColumnFilter::Values(
                                        duke_sheets_core::ValueFilter {
                                            values: Vec::new(),
                                            blank: false,
                                        },
                                    ),
                                });
                            } else {
                                log::warn!(
                                    "Skipping self-closing <filterColumn> without required colId"
                                );
                            }
                        }
                        b"filters" if in_auto_filter => {
                            current_af_filter_values.clear();
                            current_af_blank = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"blank" {
                                    current_af_blank = attr
                                        .unescape_value()
                                        .ok()
                                        .is_some_and(|s| s.as_ref() == "1" || s.as_ref() == "true");
                                }
                            }
                            in_af_filters = false;
                            current_af_column_filter =
                                Some(duke_sheets_core::ColumnFilter::Values(
                                    duke_sheets_core::ValueFilter {
                                        values: Vec::new(),
                                        blank: current_af_blank,
                                    },
                                ));
                        }
                        b"filter" if in_af_filters => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    if let Ok(value) = attr.unescape_value() {
                                        current_af_filter_values.push(value.to_string());
                                    }
                                }
                            }
                        }
                        b"customFilters" if in_auto_filter => {
                            current_af_custom_conditions.clear();
                            current_af_custom_and = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"and" {
                                    current_af_custom_and = attr
                                        .unescape_value()
                                        .ok()
                                        .is_some_and(|s| s.as_ref() == "1" || s.as_ref() == "true");
                                }
                            }
                            in_af_custom_filters = false;
                            current_af_column_filter =
                                Some(duke_sheets_core::ColumnFilter::Custom(
                                    duke_sheets_core::CustomFilters {
                                        and: current_af_custom_and,
                                        conditions: Vec::new(),
                                    },
                                ));
                        }
                        b"customFilter" if in_af_custom_filters => {
                            let mut op = duke_sheets_core::FilterOperator::Equal;
                            let mut value: Option<String> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"operator" => {
                                        if let Some(parsed) =
                                            attr.unescape_value().ok().and_then(|s| {
                                                duke_sheets_core::FilterOperator::from_ooxml(
                                                    s.as_ref(),
                                                )
                                            })
                                        {
                                            op = parsed;
                                        }
                                    }
                                    b"val" => {
                                        value = attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }

                            if let Some(value) = value {
                                current_af_custom_conditions.push(
                                    duke_sheets_core::CustomFilterCondition {
                                        operator: op,
                                        value,
                                    },
                                );
                            }
                        }
                        b"top10" if in_auto_filter => {
                            let mut top = true;
                            let mut percent = false;
                            let mut val: Option<f64> = None;
                            let mut filter_val: Option<f64> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"top" => {
                                        top = attr.unescape_value().ok().is_none_or(|s| {
                                            !(s.as_ref() == "0" || s.as_ref() == "false")
                                        });
                                    }
                                    b"percent" => {
                                        percent = attr.unescape_value().ok().is_some_and(|s| {
                                            s.as_ref() == "1" || s.as_ref() == "true"
                                        });
                                    }
                                    b"val" => {
                                        val = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"filterVal" => {
                                        filter_val = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    _ => {}
                                }
                            }

                            if let Some(val) = val {
                                current_af_column_filter =
                                    Some(duke_sheets_core::ColumnFilter::Top10(
                                        duke_sheets_core::Top10Filter {
                                            top,
                                            percent,
                                            val,
                                            filter_val,
                                        },
                                    ));
                            } else {
                                log::warn!("Skipping <top10> without required val");
                            }
                        }
                        b"dynamicFilter" if in_auto_filter => {
                            let mut filter_type: Option<
                                duke_sheets_core::auto_filter::DynamicFilterType,
                            > = None;
                            let mut val: Option<f64> = None;
                            let mut max_val: Option<f64> = None;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"type" => {
                                        filter_type = attr.unescape_value().ok().and_then(|s| {
                                                duke_sheets_core::auto_filter::DynamicFilterType::from_ooxml(
                                                    s.as_ref(),
                                                )
                                        });
                                    }
                                    b"val" => {
                                        val = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    b"maxVal" => {
                                        max_val = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<f64>().ok());
                                    }
                                    _ => {}
                                }
                            }

                            if let Some(filter_type) = filter_type {
                                current_af_column_filter =
                                    Some(duke_sheets_core::ColumnFilter::Dynamic(
                                        duke_sheets_core::auto_filter::DynamicFilter {
                                            filter_type,
                                            val,
                                            max_val,
                                        },
                                    ));
                            } else {
                                log::warn!("Skipping <dynamicFilter> without required type");
                            }
                        }
                        b"colorFilter" if in_auto_filter => {
                            let mut dxf_id: Option<u32> = None;
                            let mut cell_color = true;

                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"dxfId" => {
                                        dxf_id = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse::<u32>().ok());
                                    }
                                    b"cellColor" => {
                                        cell_color = attr.unescape_value().ok().is_none_or(|s| {
                                            !(s.as_ref() == "0" || s.as_ref() == "false")
                                        });
                                    }
                                    _ => {}
                                }
                            }

                            current_af_column_filter = Some(duke_sheets_core::ColumnFilter::Color(
                                duke_sheets_core::auto_filter::ColorFilter { dxf_id, cell_color },
                            ));
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

        // Post-process array/dataTable formulas: replicate formula to all
        // cells in the ref range and build array_result on the anchor cell.
        for (anchor_ref, ref_range, kind, formula_text, r1, r2) in pending_array_formulas {
            let anchor = match CellAddress::parse(&anchor_ref) {
                Ok(a) => a,
                Err(_) => continue,
            };
            let range = match CellRange::parse(&ref_range) {
                Ok(r) => r,
                Err(_) => continue,
            };

            let formula_for_cells = match kind {
                CellFormulaKind::Array => {
                    let f = if formula_text.starts_with('=') {
                        formula_text.clone()
                    } else {
                        format!("={}", formula_text)
                    };
                    f
                }
                CellFormulaKind::DataTable => {
                    let arg1 = r1.as_deref().unwrap_or("");
                    let arg2 = r2.as_deref().unwrap_or("");
                    format!("=TABLE({},{})", arg1, arg2)
                }
                _ => continue,
            };

            // Collect cached values from the ref range into a 2D array.
            let num_rows = (range.end.row - range.start.row + 1) as usize;
            let num_cols = (range.end.col - range.start.col + 1) as usize;
            let mut array_result: Vec<Vec<CellValue>> = Vec::with_capacity(num_rows);

            for r in range.start.row..=range.end.row {
                let mut row_values: Vec<CellValue> = Vec::with_capacity(num_cols);
                for c in range.start.col..=range.end.col {
                    // Extract the current cached value from the cell
                    let cached = worksheet.get_value_at(r, c);
                    row_values.push(cached);
                }
                array_result.push(row_values);
            }

            // For array formulas: set array_result on anchor, replicate formula to non-anchor cells.
            // For dataTable: replicate TABLE formula to all cells in range.
            for r in range.start.row..=range.end.row {
                for c in range.start.col..=range.end.col {
                    let is_anchor = r == anchor.row && c == anchor.col;
                    let row_offset = (r - range.start.row) as usize;
                    let col_offset = (c - range.start.col) as usize;
                    let cached = array_result[row_offset][col_offset].clone();

                    let _ = worksheet.set_formula_with_cached_value_at(
                        r,
                        c,
                        &formula_for_cells,
                        cached,
                    );
                    if is_anchor && kind == CellFormulaKind::Array {
                        if let Some(formula) = worksheet.formula_data_at_mut(r, c) {
                            formula.array_result = Some(array_result.clone());
                        }
                    }
                }
            }
        }

        // Post-process dynamic array formulas (cm="1" anchors + cm="2" ghosts).
        // For each anchor, determine the spill rectangle by scanning the ghost set,
        // build array_result from cached values, then create SpillTarget cells.
        for &(anchor_row, anchor_col) in &dynamic_array_anchors {
            // Determine the spill rectangle dimensions by scanning ghosts.
            // The anchor occupies (anchor_row, anchor_col). Ghosts extend right and down.
            let mut num_cols: u16 = 1;
            while dynamic_array_ghosts.contains(&(anchor_row, anchor_col + num_cols)) {
                num_cols += 1;
            }
            let mut num_rows: u32 = 1;
            'row_scan: while dynamic_array_ghosts.contains(&(anchor_row + num_rows, anchor_col)) {
                for c in 1..num_cols {
                    if !dynamic_array_ghosts.contains(&(anchor_row + num_rows, anchor_col + c)) {
                        break 'row_scan;
                    }
                }
                num_rows += 1;
            }

            if !worksheet.has_formula_at(anchor_row, anchor_col) {
                continue;
            }

            let anchor_cached = worksheet.get_value_at(anchor_row, anchor_col);

            let mut array_result: Vec<Vec<CellValue>> = Vec::with_capacity(num_rows as usize);
            for r in 0..num_rows {
                let mut row_values: Vec<CellValue> = Vec::with_capacity(num_cols as usize);
                for c in 0..num_cols {
                    let cell_row = anchor_row + r;
                    let cell_col = anchor_col + c;
                    let val = if r == 0 && c == 0 {
                        anchor_cached.clone()
                    } else {
                        worksheet.get_value_at(cell_row, cell_col)
                    };
                    row_values.push(val);
                }
                array_result.push(row_values);
            }

            if let Some(formula) = worksheet.formula_data_at_mut(anchor_row, anchor_col) {
                formula.array_result = Some(array_result);
            }

            // Replace ghost cells with SpillTarget values.
            for r in 0..num_rows {
                for c in 0..num_cols {
                    if r == 0 && c == 0 {
                        continue;
                    }
                    let cell_row = anchor_row + r;
                    let cell_col = anchor_col + c;
                    let _ = worksheet.set_cell_value_at(
                        cell_row,
                        cell_col,
                        CellValue::SpillTarget {
                            source_row: anchor_row,
                            source_col: anchor_col,
                            offset_row: r,
                            offset_col: c,
                        },
                    );
                }
            }
        }

        Ok(())
    }

    fn parse_sheet_selection_attrs(
        e: &quick_xml::events::BytesStart<'_>,
        worksheet: &mut duke_sheets_core::Worksheet,
    ) {
        let mut pane: Option<String> = None;
        let mut active_cell: Option<String> = None;
        let mut sqref: Option<String> = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"pane" => pane = attr.unescape_value().ok().map(|s| s.to_string()),
                b"activeCell" => {
                    active_cell = attr.unescape_value().ok().map(|s| s.to_string());
                }
                b"sqref" => sqref = attr.unescape_value().ok().map(|s| s.to_string()),
                _ => {}
            }
        }

        worksheet.add_selection(duke_sheets_core::worksheet::Selection {
            pane,
            active_cell,
            sqref,
        });
    }

    /// Parse `<pane>` attributes (frozen / split pane state).
    /// Called from both Event::Empty and Event::Start branches.
    fn parse_pane_attrs(
        e: &quick_xml::events::BytesStart<'_>,
        worksheet: &mut duke_sheets_core::Worksheet,
    ) {
        let mut state: Option<String> = None;
        let mut x_split_raw: Option<f64> = None;
        let mut y_split_raw: Option<f64> = None;
        let mut top_left_cell: Option<(u32, u16)> = None;
        let mut active_pane: Option<String> = None;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"state" => state = attr.unescape_value().ok().map(|s| s.to_string()),
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
                    if let Some(a1) = attr.unescape_value().ok().map(|s| s.to_string()) {
                        if let Ok(addr) = CellAddress::parse(&a1) {
                            top_left_cell = Some((addr.row, addr.col));
                        }
                    }
                }
                b"activePane" => {
                    active_pane = attr.unescape_value().ok().map(|s| s.to_string());
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

    fn parse_page_break_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<PageBreak> {
        let mut id = None;
        let mut min = 0u32;
        let mut max = 0u32;
        let mut man = false;
        let mut pt = false;

        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                b"id" => {
                    id = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok());
                }
                b"min" => {
                    min = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                }
                b"max" => {
                    max = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                }
                b"man" => {
                    man = attr.unescape_value().ok().is_some_and(|v| {
                        v.as_ref() == "1" || v.as_ref().eq_ignore_ascii_case("true")
                    });
                }
                b"pt" => {
                    pt = attr.unescape_value().ok().is_some_and(|v| {
                        v.as_ref() == "1" || v.as_ref().eq_ignore_ascii_case("true")
                    });
                }
                _ => {}
            }
        }

        id.map(|id| PageBreak {
            id,
            min,
            max,
            man,
            pt,
        })
    }

    /// Parse `<sheetProtection sheet="1" password="HHHH" .../>` into
    /// the sheet's `SheetProtection` model. Per ECMA-376 §18.3.1.85,
    /// absent attribute or value `"1"` means the action is **not**
    /// allowed; `"0"` means it **is** allowed. The writer follows the
    /// same convention.
    fn parse_sheet_protection_element(
        worksheet: &mut duke_sheets_core::Worksheet,
        e: &quick_xml::events::BytesStart<'_>,
    ) {
        use duke_sheets_core::worksheet::SheetProtection;

        let mut prot = SheetProtection {
            protected: true,
            ..Default::default()
        };
        let mut explicit_sheet = false;

        // Excel writes `protectAttr="0"` to mean "allowed". Default
        // for our model fields is `false` (= disallowed). When we see
        // `"0"` we set the field to `true` (= allowed).
        let parse_allowed = |val: &str| -> Option<bool> {
            match val {
                "0" | "false" => Some(true),
                "1" | "true" => Some(false),
                _ => None,
            }
        };

        for attr in e.attributes().flatten() {
            let value = match attr.unescape_value() {
                Ok(v) => v.into_owned(),
                Err(_) => continue,
            };
            match attr.key.local_name().as_ref() {
                b"sheet" => {
                    explicit_sheet = true;
                    prot.protected = matches!(value.as_str(), "1" | "true");
                }
                b"password" => {
                    if let Ok(h) = u16::from_str_radix(&value, 16) {
                        prot.password_hash = Some(h);
                    }
                }
                b"formatCells" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.format_cells = v;
                    }
                }
                b"formatColumns" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.format_columns = v;
                    }
                }
                b"formatRows" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.format_rows = v;
                    }
                }
                b"insertColumns" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.insert_columns = v;
                    }
                }
                b"insertRows" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.insert_rows = v;
                    }
                }
                b"insertHyperlinks" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.insert_hyperlinks = v;
                    }
                }
                b"deleteColumns" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.delete_columns = v;
                    }
                }
                b"deleteRows" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.delete_rows = v;
                    }
                }
                b"selectLockedCells" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.select_locked_cells = v;
                    }
                }
                b"selectUnlockedCells" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.select_unlocked_cells = v;
                    }
                }
                b"sort" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.sort = v;
                    }
                }
                b"autoFilter" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.auto_filter = v;
                    }
                }
                b"pivotTables" => {
                    if let Some(v) = parse_allowed(&value) {
                        prot.pivot_tables = v;
                    }
                }
                _ => {}
            }
        }

        if !explicit_sheet {
            // ECMA-376: presence of the element with no `sheet` attr
            // implies sheet=1 / "protected".
            prot.protected = true;
        }

        worksheet.set_protection(Some(prot));
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
                b"location" => {
                    location = attr.unescape_value().ok().map(|s| s.to_string());
                }
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
    #[allow(clippy::too_many_arguments)]
    fn process_cell(
        worksheet: &mut duke_sheets_core::Worksheet,
        cell_ref: &str,
        cell_type: Option<&str>,
        value: Option<&str>,
        formula: Option<&str>,
        style_idx: Option<u32>,
        shared_strings: &[SharedStringEntry],
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
            // Parse cached value (if any) from the <v> element.
            // For str/inlineStr types, an absent or empty <v/> means "".
            let cached = match (value, cell_type) {
                (Some(v), _) => match cell_type {
                    Some("b") => Some(CellValue::Boolean(
                        v == "1" || v.eq_ignore_ascii_case("true"),
                    )),
                    Some("e") => CellError::parse(v).map(CellValue::Error),
                    Some("s") => v.parse::<usize>().ok().and_then(|idx| {
                        shared_strings.get(idx).map(|entry| match entry {
                            SharedStringEntry::Plain(s) => CellValue::String(s.clone().into()),
                            SharedStringEntry::Rich(runs) => CellValue::rich_text(runs.clone()),
                        })
                    }),
                    Some("str") | Some("inlineStr") => {
                        Some(CellValue::String(v.to_string().into()))
                    }
                    None | Some("n") => v.parse::<f64>().ok().map(CellValue::Number),
                    Some(_) => Some(CellValue::String(v.to_string().into())),
                },
                // Empty <v/> with string type → cached value is ""
                (None, Some("str") | Some("inlineStr")) => Some(CellValue::String("".into())),
                (None, _) => None,
            };

            // Ensure formula starts with '='
            let formula_text = if f.starts_with('=') {
                f.to_string()
            } else {
                format!("={}", f)
            };

            if let Err(e) = worksheet.set_formula_with_cached_value_at(
                addr.row,
                addr.col,
                &formula_text,
                cached.unwrap_or(CellValue::Empty),
            ) {
                log::warn!("Skipping formula cell {}: {}", cell_ref, e);
                return Ok(());
            }
        } else if let Some(value) = value {
            let cell_value = match cell_type {
                Some("s") => match value.parse::<usize>() {
                    Ok(idx) => match shared_strings.get(idx) {
                        Some(SharedStringEntry::Plain(s)) => CellValue::String(s.clone().into()),
                        Some(SharedStringEntry::Rich(runs)) => CellValue::rich_text(runs.clone()),
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

                Some("b") => CellValue::Boolean(value == "1" || value.eq_ignore_ascii_case("true")),

                Some("e") => CellError::parse(value)
                    .map(CellValue::Error)
                    .unwrap_or_else(|| CellValue::String(value.to_string().into())),

                Some("inlineStr") => CellValue::String(decode_excel_escapes(value).into()),

                Some("str") => CellValue::String(decode_excel_escapes(value).into()),

                None | Some("n") => match value.parse::<f64>() {
                    Ok(n) => CellValue::Number(n),
                    Err(_) => CellValue::String(value.to_string().into()),
                },

                Some(_) => CellValue::String(value.to_string().into()),
            };

            if let Err(e) = worksheet.set_cell_value_at(addr.row, addr.col, cell_value) {
                log::warn!("Skipping cell {}: {}", cell_ref, e);
                return Ok(());
            }
        } else if matches!(cell_type, Some("str") | Some("inlineStr")) {
            // Empty <v/> or <is><t/></is> with string type → empty string, not Empty
            if let Err(e) =
                worksheet.set_cell_value_at(addr.row, addr.col, CellValue::String("".into()))
            {
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

/// Parse a Print_Area formula like `Sheet1!$A$1:$D$20` or `'Sheet Name'!$A$1:$D$20`
/// into a CellRange. Only handles a single contiguous range (no comma-separated multiple areas).
fn parse_print_area_formula(formula: &str, _sheet_name: &str) -> Option<CellRange> {
    let trimmed = formula.trim().trim_start_matches('=');
    let range_part = trimmed.split('!').next_back()?.trim();
    let first_area = range_part.split(',').next()?.trim();
    let clean = first_area.replace('$', "");
    CellRange::parse(&clean).ok()
}

/// Parse a Print_Titles formula like:
/// - `Sheet1!$1:$5` (repeat rows only)
/// - `Sheet1!$A:$B` (repeat cols only)
/// - `Sheet1!$1:$5,Sheet1!$A:$B` (both)
#[allow(clippy::type_complexity)]
fn parse_print_titles_formula(
    formula: &str,
    _sheet_name: &str,
) -> (Option<(u32, u32)>, Option<(u16, u16)>) {
    let mut rows = None;
    let mut cols = None;

    for part in formula.trim().trim_start_matches('=').split(',') {
        let range_part = match part.split('!').next_back() {
            Some(r) => r.trim(),
            None => continue,
        };
        let clean = range_part.replace('$', "");

        if let Some((start, end)) = clean.split_once(':') {
            let start = start.trim();
            let end = end.trim();
            if start.is_empty() || end.is_empty() {
                continue;
            }

            if start.chars().all(|c| c.is_ascii_digit()) && end.chars().all(|c| c.is_ascii_digit())
            {
                if let (Ok(r1), Ok(r2)) = (start.parse::<u32>(), end.parse::<u32>()) {
                    rows = Some((r1.saturating_sub(1), r2.saturating_sub(1)));
                }
            } else if start.chars().all(|c| c.is_ascii_alphabetic())
                && end.chars().all(|c| c.is_ascii_alphabetic())
            {
                if let (Ok(c1), Ok(c2)) = (
                    CellAddress::letters_to_column(start),
                    CellAddress::letters_to_column(end),
                ) {
                    cols = Some((c1, c2));
                }
            }
        }
    }

    (rows, cols)
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
    fn test_read_pane_non_self_closing_tags() {
        // Some XLSX generators emit <pane ...></pane> instead of <pane ... />.
        // Both forms must be handled identically.
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"></pane>
      <selection pane="bottomRight" activeCell="C4" sqref="C4"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        let freeze = sheet
            .freeze_panes()
            .expect("freeze panes from non-self-closing <pane>");
        assert_eq!(freeze.row, 2);
        assert_eq!(freeze.col, 1);
        assert_eq!(sheet.selection_active_cell(), Some((3, 2)));
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
    fn test_read_header_footer_even_first_and_flags() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0">
    <oddHeader>&amp;COdd</oddHeader>
    <oddFooter>&amp;COdd Footer</oddFooter>
    <evenHeader>&amp;CEven</evenHeader>
    <evenFooter>&amp;CEven Footer</evenFooter>
    <firstHeader>&amp;CFirst</firstHeader>
    <firstFooter>&amp;CFirst Footer</firstFooter>
  </headerFooter>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let ps = sheet.page_setup();

        // Verify all six header/footer strings
        assert_eq!(ps.odd_header.as_deref(), Some("&COdd"));
        assert_eq!(ps.odd_footer.as_deref(), Some("&COdd Footer"));
        assert_eq!(ps.even_header.as_deref(), Some("&CEven"));
        assert_eq!(ps.even_footer.as_deref(), Some("&CEven Footer"));
        assert_eq!(ps.first_header.as_deref(), Some("&CFirst"));
        assert_eq!(ps.first_footer.as_deref(), Some("&CFirst Footer"));

        // Verify flags
        assert!(ps.different_odd_even);
        assert!(ps.different_first);
        assert!(!ps.scale_with_doc);
        assert!(!ps.align_with_margins);
    }

    #[test]
    fn test_read_header_footer_default_flags_when_absent() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <headerFooter>
    <oddHeader>&amp;CTest</oddHeader>
  </headerFooter>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let ps = sheet.page_setup();

        assert_eq!(ps.odd_header.as_deref(), Some("&CTest"));
        assert_eq!(ps.even_header.as_deref(), None);
        assert_eq!(ps.first_header.as_deref(), None);

        // Defaults: differentOddEven=false, differentFirst=false, scaleWithDoc=true, alignWithMargins=true
        assert!(!ps.different_odd_even);
        assert!(!ps.different_first);
        assert!(ps.scale_with_doc);
        assert!(ps.align_with_margins);
    }

    #[test]
    fn test_read_multiple_selections_with_frozen_panes() {
        let sheet_xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="topLeft" activeCell="A1" sqref="A1"/>
      <selection pane="bottomLeft" activeCell="B5" sqref="B5:C6 E2:F3"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let bytes = build_single_sheet_xlsx(sheet_xml);
        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Should have 2 selections
        let sels = sheet.selections();
        assert_eq!(sels.len(), 2, "should have 2 selections, got: {sels:#?}");

        // First selection: topLeft pane
        assert_eq!(sels[0].pane.as_deref(), Some("topLeft"));
        assert_eq!(sels[0].active_cell.as_deref(), Some("A1"));
        assert_eq!(sels[0].sqref.as_deref(), Some("A1"));

        // Second selection: bottomLeft pane with multi-range sqref
        assert_eq!(sels[1].pane.as_deref(), Some("bottomLeft"));
        assert_eq!(sels[1].active_cell.as_deref(), Some("B5"));
        assert_eq!(sels[1].sqref.as_deref(), Some("B5:C6 E2:F3"));

        // Convenience API should return the last selection's active cell
        assert_eq!(sheet.selection_active_cell(), Some((4, 1))); // B5

        // Convenience API: first range from the last selection's sqref
        assert_eq!(
            sheet.selection_range().map(|r| r.to_string()),
            Some("B5:C6".to_string())
        );
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

    #[test]
    fn test_comments_resolved_via_rels_not_index() {
        // Verify that comments are loaded using sheet .rels targets,
        // not by assuming comments{N}.xml naming convention.
        // The comment file here is named "xl/commentsCustom.xml" (non-standard)
        // and is referenced via rId1 in the sheet .rels.
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/commentsCustom.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData></worksheet>"#).unwrap();

            // Sheet rels with non-standard comment filename
            zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../commentsCustom.xml"/></Relationships>"#).unwrap();

            // Comment file with a non-standard name
            zip.start_file("xl/commentsCustom.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Alice</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Custom path comment</t></r></text></comment></commentList></comments>"#).unwrap();

            zip.finish().unwrap();
        }

        let workbook = XlsxReader::read(Cursor::new(buf)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let comment = sheet.comment("A1").unwrap().expect("comment via rels path");
        assert_eq!(comment.author, "Alice");
        assert_eq!(comment.text, "Custom path comment");
    }
}
