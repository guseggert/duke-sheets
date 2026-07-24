//! XLSX reader

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::opc::{
    resolve_internal_target, OpcPackage, PartName, Relationship, RelationshipKind, RelationshipSet,
    XlsxDiagnosticCode, XlsxPackagePolicy,
};
use crate::styles::{
    read_styles_xml, register_roundtrip_style_data, register_roundtrip_theme_data, ParsedStyles,
};
use comments::read_comments_list;
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
use theme::read_theme_palette;

pub(crate) mod chart;
pub(crate) mod chart_ex;
mod chartsheet;
mod comments;
mod conditional_format;
mod data_validation;
mod drawing;
mod form_controls;
mod formulas;
mod shared_strings;
mod table;
mod theme;
mod workbook;

pub(crate) use formulas::CellFormulaState;
use shared_strings::SharedStringEntry;

use workbook::{read_part_rels, read_workbook_rels, read_workbook_xml};

/// Resolve a relative path from a drawing's .rels against the drawing's own path.

/// Decode Excel's `_xHHHH_` escape sequences in strings.
///
/// Excel uses this format to encode special characters in XML:
/// - `_x000d_` = CR (carriage return)
/// - `_x000a_` = LF (line feed)
/// - `_x0009_` = Tab
/// - `_x005f_` = Underscore (escaped underscore)
pub(crate) fn decode_excel_escapes(s: &str) -> String {
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
    package: &mut OpcPackage<R>,
    chart_path: &str,
    chart: &mut duke_sheets_chart::Chart,
) -> XlsxResult<()> {
    let chart_rels = read_part_rels(package, chart_path)?;
    for rel in chart_rels.iter() {
        if rel.kind() == Some(RelationshipKind::ChartStyle) {
            if let Some(mut f) = package.open_related_part(rel)? {
                let mut bytes = Vec::new();
                f.read_to_end(&mut bytes)?;
                chart.raw_chart_style = Some(bytes);
            }
        } else if rel.kind() == Some(RelationshipKind::ChartColorStyle) {
            if let Some(mut f) = package.open_related_part(rel)? {
                let mut bytes = Vec::new();
                f.read_to_end(&mut bytes)?;
                chart.raw_chart_color_style = Some(bytes);
            }
        }
    }
    Ok(())
}

fn read_chart_style_color_for_chart_ex<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    chart_path: &str,
    chart: &mut duke_sheets_chart::ChartEx,
) -> XlsxResult<()> {
    let chart_rels = read_part_rels(package, chart_path)?;
    for rel in chart_rels.iter() {
        if rel.kind() == Some(RelationshipKind::ChartStyle) {
            if let Some(mut f) = package.open_related_part(rel)? {
                let mut bytes = Vec::new();
                f.read_to_end(&mut bytes)?;
                chart.raw_chart_style = Some(bytes);
            }
        } else if rel.kind() == Some(RelationshipKind::ChartColorStyle) {
            if let Some(mut f) = package.open_related_part(rel)? {
                let mut bytes = Vec::new();
                f.read_to_end(&mut bytes)?;
                chart.raw_chart_color_style = Some(bytes);
            }
        }
    }
    Ok(())
}

/// A `<control>` entry resolved against its ctrlProp part and VML
/// shape, waiting to be matched to its drawing-part twin.
struct AssembledControl {
    shape_id: u32,
    object: duke_sheets_core::DrawingObject,
    /// True when neither controlPr nor VML supplied an anchor, so a
    /// matched twin's anchor is the last-resort fallback.
    anchor_defaulted: bool,
    consumed: bool,
}

/// Consume the first unconsumed control matching a twin's shape id.
fn take_control(
    controls: &mut [AssembledControl],
    shape_num: Option<u32>,
) -> Option<(u32, duke_sheets_core::DrawingObject, bool)> {
    let num = shape_num?;
    let control = controls
        .iter_mut()
        .find(|control| !control.consumed && control.shape_id == num)?;
    control.consumed = true;
    Some((
        control.shape_id,
        control.object.clone(),
        control.anchor_defaulted,
    ))
}

/// Build an `EmbeddedImage` from a parsed pic, resolving the blip
/// relationship to media bytes. For group children (`in_group`) the
/// child transform carries rotation/flips, so the payload keeps none.
fn resolve_pic_image<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    drawing_rels: &RelationshipSet,
    pic: drawing::PicShape,
    in_group: bool,
) -> XlsxResult<duke_sheets_chart::EmbeddedImage> {
    let mut image = duke_sheets_chart::EmbeddedImage {
        format: duke_sheets_chart::ImageFormat::Png, // placeholder, resolved from media path
        media_path: pic.blip_rel.unwrap_or_default(),
        svg_media_path: pic.svg_rel,
        width_emu: pic.ext_cx,
        height_emu: pic.ext_cy,
        rotation: if in_group { None } else { pic.rotation },
        flip_h: !in_group && pic.flip_h,
        flip_v: !in_group && pic.flip_v,
        data: Vec::new(),
        svg_data: None,
    };
    let image_rel_id = image.media_path.clone();
    if let Some(rel) = drawing_rels.get(&image_rel_id) {
        let Some(path) = rel.internal_path() else {
            return Ok(image);
        };
        let ext = path.rsplit('.').next().unwrap_or("");
        if let Some(fmt) = duke_sheets_chart::ImageFormat::from_extension(ext) {
            image.format = fmt;
        }
        image.media_path = path.to_string();
        if let Some(mut f) = package.open_related_part(rel)? {
            let mut buf = Vec::new();
            if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
                image.data = buf;
            }
        }
    }
    if let Some(svg_rel_id) = image.svg_media_path.clone() {
        if let Some(rel) = drawing_rels.get(svg_rel_id.as_str()) {
            let Some(path) = rel.internal_path() else {
                return Ok(image);
            };
            image.svg_media_path = Some(path.to_string());
            if let Some(mut f) = package.open_related_part(rel)? {
                let mut buf = Vec::new();
                if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
                    image.svg_data = Some(buf);
                }
            }
        }
    }
    Ok(image)
}

fn shape_from_parsed(parsed: &drawing::ParsedShape) -> duke_sheets_core::Shape {
    let fill = match parsed.fill {
        duke_sheets_chart::drawing_part::ShapeFill::None => duke_sheets_core::ShapeFill::None,
        duke_sheets_chart::drawing_part::ShapeFill::Solid(color) => {
            duke_sheets_core::ShapeFill::Solid(duke_sheets_core::drawing::color_from_drawing_part(
                color,
            ))
        }
    };
    let mut shape = duke_sheets_core::Shape::preset(parsed.geometry.clone());
    shape.fill = fill;
    shape.line = duke_sheets_core::ShapeLine {
        color: parsed
            .line
            .color
            .map(duke_sheets_core::drawing::color_from_drawing_part),
        width_emu: parsed.line.width_emu,
        dash_style: parsed.line.dash_style.clone(),
        no_fill: parsed.line.no_fill,
    };
    shape.text = parsed
        .text
        .as_ref()
        .map(duke_sheets_core::DrawingText::from_drawing_part_text);
    shape.rotation = parsed.xfrm.rotation;
    shape.flip_h = parsed.xfrm.flip_h;
    shape.flip_v = parsed.xfrm.flip_v;
    shape.set_preserved_shape_properties(parsed.raw_shape_properties.clone());
    shape.set_preserved_text_body(parsed.raw_text_body.clone());
    shape
}

/// Convert a parsed group into the model, resolving child images and
/// matching control-twin children to their controls (placed in the
/// group with the twin's child transform).
fn build_group<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    drawing_rels: &RelationshipSet,
    group: drawing::ParsedGroup,
    controls: &mut [AssembledControl],
) -> XlsxResult<duke_sheets_core::Group> {
    use duke_sheets_core::{ChildTransform, DrawingKind, DrawingMeta, GroupChild};
    let mut children = Vec::new();
    for child in group.children {
        match child {
            drawing::ParsedChild::Pic(pic) => {
                let transform = ChildTransform {
                    x_emu: pic.off_x,
                    y_emu: pic.off_y,
                    cx_emu: pic.ext_cx,
                    cy_emu: pic.ext_cy,
                    rotation: pic.rotation.unwrap_or(0),
                    flip_h: pic.flip_h,
                    flip_v: pic.flip_v,
                };
                let meta = DrawingMeta {
                    name: Some(pic.name.clone()),
                    alt_text: pic.descr.clone(),
                    title: pic.title.clone(),
                    hidden: pic.hidden,
                    ..DrawingMeta::default()
                };
                let image = resolve_pic_image(package, drawing_rels, pic, true)?;
                children.push(GroupChild {
                    meta,
                    transform,
                    kind: DrawingKind::Image(image),
                });
            }
            drawing::ParsedChild::Shape(shape) => {
                let meta = DrawingMeta {
                    name: Some(shape.name.clone()),
                    alt_text: shape.descr.clone(),
                    title: shape.title.clone(),
                    hidden: shape.hidden,
                    ..DrawingMeta::default()
                };
                children.push(GroupChild {
                    meta,
                    transform: shape.xfrm.clone(),
                    kind: DrawingKind::Shape(Box::new(shape_from_parsed(&shape))),
                });
            }
            drawing::ParsedChild::Group(inner) => {
                let transform = ChildTransform {
                    x_emu: inner.transform.x_emu,
                    y_emu: inner.transform.y_emu,
                    cx_emu: inner.transform.cx_emu,
                    cy_emu: inner.transform.cy_emu,
                    rotation: inner.transform.rotation,
                    flip_h: inner.transform.flip_h,
                    flip_v: inner.transform.flip_v,
                };
                let meta = DrawingMeta {
                    name: Some(inner.name.clone()),
                    alt_text: inner.descr.clone(),
                    title: inner.title.clone(),
                    hidden: inner.hidden,
                    ..DrawingMeta::default()
                };
                let built = build_group(package, drawing_rels, inner, controls)?;
                children.push(GroupChild {
                    meta,
                    transform,
                    kind: DrawingKind::Group(Box::new(built)),
                });
            }
            drawing::ParsedChild::Twin(twin) => {
                // A control-twin child becomes the matched control,
                // positioned by the twin's child transform.
                if let Some((_, mut object, _)) = take_control(controls, twin.shape_num) {
                    if let Some(name) = twin.name {
                        object.meta.name = Some(name);
                    }
                    if twin.descr.is_some() {
                        object.meta.alt_text = twin.descr;
                    }
                    object.meta.title = twin.title;
                    if let Some(control) = object.kind.as_form_control_mut() {
                        // Foreign files may put a txBody on twins of
                        // caption-less kinds; ignore the stray text.
                        if let (Some(text), Some(caption)) =
                            (twin.text.as_ref(), control.caption_mut())
                        {
                            *caption = form_controls::control_text_from_twin(text);
                        }
                        if control.macro_name.is_none() {
                            control.macro_name = twin
                                .macro_name
                                .as_deref()
                                .map(duke_sheets_vml::decode_macro_formula);
                        }
                    }
                    children.push(GroupChild {
                        meta: object.meta,
                        transform: twin.xfrm,
                        kind: object.kind,
                    });
                }
            }
        }
    }
    Ok(duke_sheets_core::Group {
        transform: group.transform,
        children,
    })
}

/// Scan a raw anchor's bytes for relationship references and capture
/// each referenced relationship with its original id/target plus the
/// target part bytes when internal. References are found by value:
/// any attribute whose value equals a rel id in the drawing's .rels
/// counts (r:id/r:embed/r:link, SmartArt `dgm:relIds` r:dm/r:lo/r:qs/
/// r:cs, VML `o:relid`, arbitrary prefixes).
fn capture_raw_rels<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    source_part: &str,
    bytes: &[u8],
    relationships: &RelationshipSet,
) -> XlsxResult<Vec<duke_sheets_core::RawRel>> {
    let mut ids: Vec<String> = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                for attr in e.attributes().flatten() {
                    if let Ok(value) = attr.unescape_value() {
                        if relationships.get(value.as_ref()).is_some()
                            && !ids.iter().any(|id| id == value.as_ref())
                        {
                            ids.push(value.to_string());
                        }
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let mut rels = Vec::new();
    for id in ids {
        let Some(rel) = relationships.get(&id) else {
            continue;
        };
        let source_name = PartName::from_zip_name(source_part).ok();
        let part = if source_name.as_ref() == rel.internal_part() {
            None
        } else {
            match package.open_related_part(rel)? {
                Some(mut file) => {
                    let mut bytes = Vec::new();
                    file.read_to_end(&mut bytes)?;
                    Some(bytes)
                }
                None => None,
            }
        };
        rels.push(duke_sheets_core::RawRel {
            id,
            rel_type: rel.rel_type.clone(),
            target: rel.raw_target.clone(),
            external: rel.is_external(),
            part,
        });
    }
    Ok(rels)
}

/// Options for opening an XLSX/XLSM workbook.
#[derive(Debug, Clone, Default)]
pub struct XlsxReadOptions {
    /// Password for encrypted workbooks.
    pub password: Option<String>,
    /// Retry encrypted workbooks with Excel's well-known
    /// `VelvetSweatshop` sentinel when no password is supplied,
    /// before reporting them as encrypted.
    pub try_velvet_sweatshop: bool,
    /// Skip the post-decrypt HMAC integrity check; default-off
    /// (false) matches Office.
    pub skip_integrity_check: bool,
}

/// Workbook and package diagnostics returned by report-oriented read APIs.
#[derive(Debug)]
#[non_exhaustive]
pub struct XlsxReadReport {
    /// Parsed workbook.
    pub workbook: Workbook,
    /// Recoverable OPC package problems encountered while traversing it.
    pub diagnostics: Vec<crate::opc::XlsxDiagnostic>,
}

/// XLSX file reader
pub struct XlsxReader;

fn log_package_diagnostics(diagnostics: &[crate::opc::XlsxDiagnostic]) {
    for diagnostic in diagnostics {
        log::warn!("{}", diagnostic.message);
    }
}

fn workbook_resource_path<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    workbook_path: &PartName,
    relationship_path: Option<&str>,
    conventional_target: &str,
    expected_content_type: &str,
) -> XlsxResult<Option<String>> {
    if let Some(path) = relationship_path {
        let part_name = PartName::from_zip_name(path)?;
        if package.part_exists(&part_name) {
            return Ok(Some(path.to_string()));
        }
        package.diagnostics_mut().violation(
            XlsxDiagnosticCode::MissingRelationshipTarget,
            format!("missing Workbook resource {part_name}"),
            Some(workbook_path.as_str()),
            None,
            Some(part_name.as_str()),
        )?;
        return Ok(None);
    }

    if package.policy() == XlsxPackagePolicy::Compatible {
        let path = resolve_internal_target(workbook_path.zip_name(), conventional_target)?;
        let part_name = PartName::from_zip_name(&path)?;
        if package.part_exists(&part_name) {
            package.validate_part_content_type(&part_name, expected_content_type)?;
            package.diagnostics_mut().recovery(
                XlsxDiagnosticCode::CanonicalPartFallback,
                format!(
                    "using conventional Workbook resource {part_name} because its relationship is absent"
                ),
                Some(workbook_path.as_str()),
                None,
                Some(part_name.as_str()),
            );
            return Ok(Some(path));
        }
    }
    Ok(None)
}

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

    /// Read a workbook from a file path with explicit open options
    /// (password, encrypted-workbook handling).
    pub fn read_file_with<P: AsRef<Path>>(
        path: P,
        options: &XlsxReadOptions,
    ) -> XlsxResult<Workbook> {
        let bytes = std::fs::read(path)?;
        Self::read_bytes_with(&bytes, options)
    }

    /// Read a file and return structured OPC package diagnostics.
    pub fn read_file_with_report<P: AsRef<Path>>(
        path: P,
        options: &XlsxReadOptions,
        policy: XlsxPackagePolicy,
    ) -> XlsxResult<XlsxReadReport> {
        let bytes = std::fs::read(path)?;
        Self::read_bytes_with_report(&bytes, options, policy)
    }

    /// Read a workbook from raw bytes with explicit open options.
    ///
    /// Encrypted XLSX files are CFB envelopes (not plain ZIPs); when
    /// the leading magic bytes match CFB we delegate to
    /// `duke_sheets_crypto::ooxml::decrypt` and then proceed with the
    /// resulting plaintext ZIP.
    pub fn read_bytes_with(bytes: &[u8], options: &XlsxReadOptions) -> XlsxResult<Workbook> {
        let report = Self::read_bytes_with_report(bytes, options, XlsxPackagePolicy::Compatible)?;
        log_package_diagnostics(&report.diagnostics);
        Ok(report.workbook)
    }

    /// Read bytes and return structured OPC package diagnostics.
    pub fn read_bytes_with_report(
        bytes: &[u8],
        options: &XlsxReadOptions,
        policy: XlsxPackagePolicy,
    ) -> XlsxResult<XlsxReadReport> {
        let password = options.password.as_deref();
        if is_cfb_envelope(bytes) {
            let try_pw = match password {
                Some(p) => p,
                None if options.try_velvet_sweatshop => "VelvetSweatshop",
                None => {
                    return Err(XlsxError::Encrypted(
                        "workbook is encrypted but no password was supplied".into(),
                    ));
                }
            };
            return match decrypt_ooxml_envelope(bytes, try_pw, options.skip_integrity_check) {
                Ok(decrypted) => Self::read_with_report(std::io::Cursor::new(decrypted), policy),
                Err(XlsxError::BadPassword) if password.is_none() => Err(XlsxError::Encrypted(
                    "workbook is encrypted but no password was supplied".into(),
                )),
                Err(e) => Err(e),
            };
        }
        Self::read_with_report(std::io::Cursor::new(bytes), policy)
    }

    /// Read a workbook from a reader
    pub fn read<R: Read + Seek>(reader: R) -> XlsxResult<Workbook> {
        let report = Self::read_with_report(reader, XlsxPackagePolicy::Compatible)?;
        log_package_diagnostics(&report.diagnostics);
        Ok(report.workbook)
    }

    /// Read a workbook and return structured OPC package diagnostics.
    pub fn read_with_report<R: Read + Seek>(
        reader: R,
        policy: XlsxPackagePolicy,
    ) -> XlsxResult<XlsxReadReport> {
        let mut package = OpcPackage::open(reader, policy)?;
        let workbook_path = package.discover_workbook_part()?;
        let workbook_rels = read_workbook_rels(&mut package, &workbook_path)?;
        let shared_strings_path = workbook_resource_path(
            &mut package,
            &workbook_path,
            workbook_rels.shared_strings_path.as_deref(),
            "sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )?;
        let styles_path = workbook_resource_path(
            &mut package,
            &workbook_path,
            workbook_rels.styles_path.as_deref(),
            "styles.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        )?;
        let theme_path = workbook_resource_path(
            &mut package,
            &workbook_path,
            workbook_rels.theme_path.as_deref(),
            "theme/theme1.xml",
            "application/vnd.openxmlformats-officedocument.theme+xml",
        )?;
        let shared_strings = match shared_strings_path.as_deref() {
            Some(path) => shared_strings::parse_shared_strings(package.open_zip_name(path)?)?,
            None => Vec::new(),
        };
        let parsed_styles = match styles_path.as_deref() {
            Some(path) => read_styles_xml(package.open_zip_name(path)?)?,
            None => Self::default_styles(),
        };
        let roundtrip_style_data = parsed_styles.roundtrip_data();
        let (theme_palette, raw_theme_xml) = match theme_path.as_deref() {
            Some(path) => {
                let (palette, bytes) = read_theme_palette(package.open_zip_name(path)?)?;
                (Some(palette), Some(bytes))
            }
            None => (None, None),
        };
        let cell_styles = parsed_styles.cell_styles;
        let dxf_styles = parsed_styles.dxf_styles;

        let policy = package.policy();
        let wb_props = read_workbook_xml(package.open_part(&workbook_path)?, policy)?;

        let sheet_paths = workbook_rels.sheet_paths;
        let chartsheet_paths = workbook_rels.chartsheet_paths;
        let unmodeled_sheet_rels = workbook_rels.unmodeled_sheet_rels;

        // Create workbook
        let mut workbook = Workbook::empty();
        workbook.settings_mut().date_1904 = wb_props.date_1904;
        workbook.set_workbook_protection(wb_props.workbook_protection);

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
                let sheet_rels = read_part_rels(&mut package, path)?;
                let worksheet_file = package.open_zip_name(path)?;
                let pending_controls =
                    form_controls::dedupe_pending_controls(Self::read_worksheet(
                    worksheet_file,
                    workbook.worksheet_mut(sheet_idx).unwrap(),
                    &shared_strings,
                    &cell_styles,
                    &dxf_styles,
                    &sheet_rels,
                )?);

                // Read comments for this worksheet (if present).
                // Resolve paths via sheet .rels relationships; fall back to
                // index-based filenames for files that lack .rels entries.
                let comments_rel = sheet_rels
                    .iter()
                    .find(|rel| rel.kind() == Some(RelationshipKind::Comments));
                let comments_path = comments_rel
                    .and_then(|rel| rel.internal_path().map(str::to_string))
                    .or_else(|| {
                        (package.policy() == XlsxPackagePolicy::Compatible)
                            .then(|| format!("xl/comments{}.xml", ws_count))
                    });
                let vml_rel = sheet_rels
                    .iter()
                    .find(|rel| rel.kind() == Some(RelationshipKind::VmlDrawing));
                let vml_path = vml_rel
                    .and_then(|rel| rel.internal_path().map(str::to_string))
                    .or_else(|| {
                        (package.policy() == XlsxPackagePolicy::Compatible)
                            .then(|| format!("xl/drawings/vmlDrawing{}.vml", ws_count))
                    });

                // Assemble the sheet's drawing objects: the drawing
                // part's document order (with control twins replaced
                // by their controls), unmatched controls, and comments
                // spliced by the legacy VML shape order.
                let vml_bytes = match vml_rel {
                    Some(rel) => match package.open_related_part(rel)? {
                        Some(mut file) => {
                            let mut bytes = Vec::new();
                            file.read_to_end(&mut bytes)?;
                            Some(bytes)
                        }
                        None => None,
                    },
                    None => match vml_path.as_deref() {
                        Some(path) => match package.open_zip_name(path) {
                            Ok(mut file) => {
                                let mut bytes = Vec::new();
                                file.read_to_end(&mut bytes)?;
                                Some(bytes)
                            }
                            Err(_) => None,
                        },
                        None => None,
                    },
                };
                let comments = match comments_rel {
                    Some(rel) => match package.open_related_part(rel)? {
                        Some(file) => read_comments_list(file)?,
                        None => Vec::new(),
                    },
                    None => match comments_path.as_deref() {
                        Some(path) => match package.open_zip_name(path) {
                            Ok(file) => read_comments_list(file)?,
                            Err(_) => Vec::new(),
                        },
                        None => Vec::new(),
                    },
                };
                let objects = Self::merge_sheet_drawings(
                    &mut package,
                    &sheet_rels,
                    &pending_controls,
                    vml_bytes.as_deref(),
                    comments,
                )?;
                workbook
                    .worksheet_mut(sheet_idx)
                    .unwrap()
                    .drawings_mut()
                    .extend(objects);

                // Read tables for this worksheet (if present).
                // Each relationship with type ending in "/table" points to
                // an xl/tables/tableN.xml part.
                let mut table_rels: Vec<(&Relationship, &str)> = sheet_rels
                    .iter()
                    .filter_map(|rel| {
                        (rel.kind() == Some(RelationshipKind::Table))
                            .then(|| rel.internal_path().map(|path| (rel, path)))
                            .flatten()
                    })
                    .collect();
                table_rels.sort_by_key(|(_, path)| *path);
                for (rel, _) in table_rels {
                    let Some(file) = package.open_related_part(rel)? else {
                        continue;
                    };
                    if let Some(t) = table::parse_table(file)? {
                        workbook.worksheet_mut(sheet_idx).unwrap().add_table(t);
                    }
                }
            } else if let Some(cs_path) = chartsheet_paths.get(&sheet_entry.r_id) {
                let mut chart_found = false;
                let drawing_rid =
                    chartsheet::read_chartsheet_drawing_rid(package.open_zip_name(cs_path)?)?;
                if let Some(rid) = drawing_rid {
                    let cs_rels = read_part_rels(&mut package, cs_path)?;
                    if let Some(drawing_rel) = cs_rels.get(&rid) {
                        if let Some(drawing_path) = drawing_rel.internal_path() {
                            let entries = match package.open_related_part(drawing_rel)? {
                                Some(file) => drawing::parse_drawing_entries(file, drawing_path)?,
                                None => Vec::new(),
                            };
                            // Chartsheets model only the chart; every other
                            // anchor is preserved as raw bytes.
                            let mut raw_anchors: Vec<Vec<u8>> = Vec::new();
                            let mut chart_refs: Vec<drawing::DrawingChartRef> = Vec::new();
                            for entry in entries {
                                match entry.kind {
                                    drawing::DrawingEntryKind::Chart(chart_ref) => {
                                        chart_refs.push(chart_ref)
                                    }
                                    _ => raw_anchors.push(entry.bytes),
                                }
                            }
                            // Capture the relationships the raw anchors
                            // reference, minus the modeled chart's.
                            let drawing_rels = read_part_rels(&mut package, drawing_path)?;
                            let chart_ids: std::collections::HashSet<&str> =
                                chart_refs.iter().map(|c| c.rel_id.as_str()).collect();
                            let mut captured_ids: std::collections::HashSet<String> =
                                std::collections::HashSet::new();
                            let mut raw_rels: Vec<duke_sheets_core::RawRel> = Vec::new();
                            for anchor in &raw_anchors {
                                for rel in capture_raw_rels(
                                    &mut package,
                                    drawing_path,
                                    anchor,
                                    &drawing_rels,
                                )? {
                                    if chart_ids.contains(rel.id.as_str())
                                        || !captured_ids.insert(rel.id.clone())
                                    {
                                        continue;
                                    }
                                    raw_rels.push(rel);
                                }
                            }
                            for chart_ref in chart_refs {
                                if let Some(dr) = drawing_rels.get(&chart_ref.rel_id) {
                                    if chart_ref.is_chart_ex {
                                        // ChartEx in a chartsheet - skip for now (chartsheets
                                        // require a standard Chart). Parse it but don't embed.
                                        continue;
                                    }
                                    let Some(chart_path) = dr.internal_path() else {
                                        continue;
                                    };
                                    let Some(file) = package.open_related_part(dr)? else {
                                        continue;
                                    };
                                    let mut c = chart::parse_chart(file)?;
                                    read_chart_style_color(&mut package, chart_path, &mut c)?;
                                    let cs_idx = workbook.add_chartsheet_unchecked(
                                        duke_sheets_core::ChartSheet {
                                            name: sheet_entry.name.clone(),
                                            chart: c,
                                            visibility: sheet_entry.visibility,
                                            raw_drawing_objects: raw_anchors.clone(),
                                            raw_drawing_rels: raw_rels.clone(),
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
                        raw_drawing_rels: Vec::new(),
                    });
                    workbook
                        .sheet_order_mut()
                        .push(SheetSlot::ChartSheet(cs_idx));
                }
            } else if !unmodeled_sheet_rels.contains(&sheet_entry.r_id)
                && package.policy() == XlsxPackagePolicy::Strict
            {
                return Err(XlsxError::InvalidFormat(format!(
                    "sheet {} has no resolvable worksheet or chartsheet relationship",
                    sheet_entry.name
                )));
            }
        }

        // Apply print area and print titles from named ranges to worksheets.
        Self::apply_print_settings(&mut workbook);

        // Compatible preserves the historical at-least-one-sheet
        // recovery. Strict does not invent a worksheet when all source
        // sheets were valid but unmodeled dialog or macro sheets.
        if workbook.sheet_count() == 0 && workbook.chartsheet_count() == 0 {
            if package.policy() == XlsxPackagePolicy::Strict {
                if unmodeled_sheet_rels.is_empty() {
                    return Err(XlsxError::InvalidFormat(
                        "Workbook contains no readable sheets".into(),
                    ));
                }
            } else {
                workbook.add_worksheet()?;
            }
        }

        register_roundtrip_style_data(&workbook, roundtrip_style_data);
        if let Some(theme) = theme_palette {
            workbook.set_theme_palette(theme);
        }
        if let Some(theme_bytes) = raw_theme_xml {
            register_roundtrip_theme_data(&workbook, theme_bytes);
        }

        let diagnostics = package.into_diagnostics();
        Ok(XlsxReadReport {
            workbook,
            diagnostics,
        })
    }

    fn default_styles() -> ParsedStyles {
        ParsedStyles {
            cell_styles: vec![Style::default()],
            cell_style_xfs: vec![Style::default()],
            named_styles: Vec::new(),
            cell_xf_xf_ids: vec![0],
            dxf_styles: Vec::new(),
        }
    }

    /// Assemble a worksheet's drawing list: drawing-part entries in
    /// document order (control twins replaced by their matched
    /// controls), unmatched controls appended in `<controls>` block
    /// order, and comments spliced by the legacy VML shape sequence.
    fn merge_sheet_drawings<R: Read + Seek>(
        package: &mut OpcPackage<R>,
        sheet_rels: &RelationshipSet,
        pending_controls: &[form_controls::PendingControl],
        vml_bytes: Option<&[u8]>,
        comments: Vec<(u32, u16, duke_sheets_core::comment::CellComment)>,
    ) -> XlsxResult<Vec<duke_sheets_core::DrawingObject>> {
        use duke_sheets_core::DrawingObject;

        let vml_shapes = vml_bytes
            .map(duke_sheets_vml::parse_vml_shapes)
            .unwrap_or_default();
        let vml_controls: HashMap<u32, &duke_sheets_vml::VmlControl> = vml_shapes
            .iter()
            .filter_map(|shape| match &shape.kind {
                duke_sheets_vml::VmlShapeKind::Control(control) => Some((shape.shape_num, control)),
                _ => None,
            })
            .collect();

        // Resolve form controls in <controls> block order: state from
        // ctrlProps, caption from VML, anchor precedence controlPr >
        // VML x:Anchor > (later) twin anchor.
        let mut controls: Vec<AssembledControl> = Vec::new();
        for pending in pending_controls {
            let Some(rel) = sheet_rels.get(&pending.rid) else {
                continue;
            };
            if rel.kind() != Some(RelationshipKind::ControlProperties) {
                continue;
            }
            let Some(mut f) = package.open_related_part(rel)? else {
                continue;
            };
            let mut bytes = Vec::new();
            f.read_to_end(&mut bytes)?;
            let Some(pr) = form_controls::parse_ctrl_prop(&bytes) else {
                continue;
            };
            let vml = vml_controls.get(&pending.shape_id).copied();
            if let Some(object) = form_controls::assemble_with_vml(pending, &pr, vml) {
                let anchor_defaulted =
                    pending.anchor.is_none() && vml.and_then(|shape| shape.anchor_px).is_none();
                controls.push(AssembledControl {
                    shape_id: pending.shape_id,
                    object,
                    anchor_defaulted,
                    consumed: false,
                });
            }
        }

        // Older or partially-authored workbooks can carry Forms shapes
        // only in VML. Surface any shape not represented by a ctrlProps
        // entry, including unknown legacy Forms controls. Pict/ActiveX and
        // Note shapes return None and remain outside the form-control model.
        let mut represented: std::collections::HashSet<u32> =
            controls.iter().map(|control| control.shape_id).collect();
        for shape in &vml_shapes {
            let duke_sheets_vml::VmlShapeKind::Control(vml) = &shape.kind else {
                continue;
            };
            if represented.contains(&shape.shape_num) {
                continue;
            }
            if let Some(object) = vml.to_drawing_object() {
                represented.insert(shape.shape_num);
                controls.push(AssembledControl {
                    shape_id: shape.shape_num,
                    object,
                    anchor_defaulted: vml.anchor_px.is_none(),
                    consumed: false,
                });
            }
        }

        // Drawing part entries in document order.
        let mut natives: Vec<(DrawingObject, Option<u32>)> = Vec::new();
        let mut drawing_targets: Vec<(&Relationship, &str)> = sheet_rels
            .iter()
            .filter(|rel| rel.kind() == Some(RelationshipKind::Drawing))
            .filter_map(|rel| rel.internal_path().map(|path| (rel, path)))
            .collect();
        // Numeric-aware sort so rId2 precedes rId10.
        drawing_targets.sort_by_key(|(rel, _)| {
            (
                rel.id
                    .strip_prefix("rId")
                    .and_then(|n| n.parse::<u64>().ok())
                    .unwrap_or(u64::MAX),
                rel.id.clone(),
            )
        });
        for (drawing_rel, drawing_path) in &drawing_targets {
            let Some(file) = package.open_related_part(drawing_rel)? else {
                continue;
            };
            let entries = drawing::parse_drawing_entries(file, drawing_path)?;
            if entries.is_empty() {
                continue;
            }
            let drawing_rels = read_part_rels(package, drawing_path)?;
            for entry in entries {
                match entry.kind {
                    drawing::DrawingEntryKind::Image(pic) => {
                        let pic = *pic;
                        // Preserve the cNvPr name verbatim (even when
                        // empty) so the writer re-emits name=""
                        // byte-identically.
                        let name = pic.name.clone();
                        let descr = pic.descr.clone();
                        let title = pic.title.clone();
                        let hidden = pic.hidden;
                        let image = resolve_pic_image(package, &drawing_rels, pic, false)?;
                        let mut object = DrawingObject::image(image).with_anchor(entry.anchor);
                        object.meta.name = Some(name);
                        object.meta.alt_text = descr;
                        object.meta.title = title;
                        object.meta.hidden = hidden;
                        object.meta.locked = entry.locked;
                        object.meta.printable = entry.printable;
                        natives.push((object, None));
                    }
                    drawing::DrawingEntryKind::Shape(shape) => {
                        let mut object = DrawingObject::shape(shape_from_parsed(&shape))
                            .with_anchor(entry.anchor);
                        object.meta.name = Some(shape.name.clone());
                        object.meta.alt_text = shape.descr.clone();
                        object.meta.title = shape.title.clone();
                        object.meta.hidden = shape.hidden;
                        object.meta.locked = entry.locked;
                        object.meta.printable = entry.printable;
                        natives.push((object, None));
                    }
                    drawing::DrawingEntryKind::Chart(chart_ref) => {
                        let Some(dr) = drawing_rels.get(&chart_ref.rel_id) else {
                            continue;
                        };
                        let Some(chart_path) = dr.internal_path() else {
                            continue;
                        };
                        let Some(file) = package.open_related_part(dr)? else {
                            continue;
                        };
                        if chart_ref.is_chart_ex {
                            let mut cx = chart_ex::parse_chart_ex(file)?;
                            cx.raw_mc_fallback = chart_ref.raw_mc_fallback;
                            read_chart_style_color_for_chart_ex(package, chart_path, &mut cx)?;
                            let mut object =
                                DrawingObject::chart_ex(cx).with_anchor(chart_ref.anchor);
                            object.meta.name = chart_ref.name;
                            object.meta.alt_text = chart_ref.descr;
                            object.meta.title = chart_ref.title;
                            object.meta.hidden = chart_ref.hidden;
                            object.meta.locked = entry.locked;
                            object.meta.printable = entry.printable;
                            natives.push((object, None));
                        } else {
                            let mut c = chart::parse_chart(file)?;
                            read_chart_style_color(package, chart_path, &mut c)?;
                            let mut object = DrawingObject::chart(c).with_anchor(chart_ref.anchor);
                            object.meta.name = chart_ref.name;
                            object.meta.alt_text = chart_ref.descr;
                            object.meta.title = chart_ref.title;
                            object.meta.hidden = chart_ref.hidden;
                            object.meta.locked = entry.locked;
                            object.meta.printable = entry.printable;
                            natives.push((object, None));
                        }
                    }
                    drawing::DrawingEntryKind::Group(group) => {
                        let name = group.name.clone();
                        let descr = group.descr.clone();
                        let title = group.title.clone();
                        let hidden = group.hidden;
                        let built = build_group(package, &drawing_rels, group, &mut controls)?;
                        let mut object = DrawingObject::group(built).with_anchor(entry.anchor);
                        object.meta.name = Some(name);
                        object.meta.alt_text = descr;
                        object.meta.title = title;
                        object.meta.hidden = hidden;
                        object.meta.locked = entry.locked;
                        object.meta.printable = entry.printable;
                        natives.push((object, None));
                    }
                    drawing::DrawingEntryKind::ControlTwin(twin) => {
                        // The twin is a placeholder for the matched
                        // legacy control; unmatched twins are dropped.
                        if let Some((shape_id, mut object, anchor_defaulted)) =
                            take_control(&mut controls, twin.shape_num)
                        {
                            if anchor_defaulted {
                                object.anchor = entry.anchor;
                            }
                            if let Some(name) = twin.name {
                                object.meta.name = Some(name);
                            }
                            if twin.descr.is_some() {
                                object.meta.alt_text = twin.descr;
                            }
                            object.meta.title = twin.title;
                            if let Some(control) = object.kind.as_form_control_mut() {
                                // Foreign files may put a txBody on
                                // twins of caption-less kinds; ignore
                                // the stray text.
                                if let (Some(text), Some(caption)) =
                                    (twin.text.as_ref(), control.caption_mut())
                                {
                                    *caption = form_controls::control_text_from_twin(text);
                                }
                                if control.macro_name.is_none() {
                                    control.macro_name = twin
                                        .macro_name
                                        .as_deref()
                                        .map(duke_sheets_vml::decode_macro_formula);
                                }
                            }
                            natives.push((object, Some(shape_id)));
                        }
                    }
                    drawing::DrawingEntryKind::Raw => {
                        let rels =
                            capture_raw_rels(package, drawing_path, &entry.bytes, &drawing_rels)?;
                        let object = DrawingObject::raw(duke_sheets_core::RawDrawing {
                            bytes: entry.bytes,
                            rels,
                        })
                        .with_anchor(entry.anchor);
                        natives.push((object, None));
                    }
                }
            }
        }

        // Unmatched controls (no twin; legacy files) append after all
        // native entries, in <controls> block order.
        for control in controls {
            if !control.consumed {
                natives.push((control.object, Some(control.shape_id)));
            }
        }

        Ok(duke_sheets_vml::splice_comments(
            natives,
            comments,
            &vml_shapes,
        ))
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

    /// Read a worksheet XML stream.
    #[allow(clippy::too_many_arguments)]
    fn read_worksheet<R: Read>(
        reader: R,
        worksheet: &mut duke_sheets_core::Worksheet,
        shared_strings: &[SharedStringEntry],
        cell_styles: &[Style],
        dxf_styles: &[Style],
        sheet_rels: &RelationshipSet,
    ) -> XlsxResult<Vec<form_controls::PendingControl>> {
        let mut xml_reader = Reader::from_reader(BufReader::new(reader));
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

        // Form control state (<controls> block after legacyDrawing).
        let mut pending_controls: Vec<form_controls::PendingControl> = Vec::new();
        let mut in_controls = false;
        let mut current_control: Option<form_controls::PendingControl> = None;
        let mut in_control_anchor = false;
        let mut control_anchor_move = false;
        let mut control_anchor_size = false;
        let mut control_anchor_in_from = true;
        let mut control_anchor_vals = [[0i64; 4]; 2];
        // (marker 0=from/1=to, field 0=col 1=colOff 2=row 3=rowOff)
        let mut control_anchor_field: Option<(usize, usize)> = None;
        let mut control_anchor_text = String::new();

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

        // Protected range state
        let mut current_protected_range: Option<duke_sheets_core::ProtectedRange> = None;
        let mut in_protected_range_security_descriptor = false;

        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    match e.name().local_name().as_ref() {
                        b"controls" => in_controls = true,
                        b"control" if in_controls => {
                            let mut pending = form_controls::PendingControl::new();
                            for attr in e.attributes().flatten() {
                                let value = String::from_utf8_lossy(&attr.value).into_owned();
                                match attr.key.local_name().as_ref() {
                                    b"shapeId" => pending.shape_id = value.parse().unwrap_or(0),
                                    b"name" => pending.name = Some(value),
                                    b"id" => pending.rid = value,
                                    _ => {}
                                }
                            }
                            current_control = Some(pending);
                        }
                        b"controlPr" if current_control.is_some() => {
                            let pending = current_control.as_mut().unwrap();
                            for attr in e.attributes().flatten() {
                                let truthy = matches!(&*attr.value, b"1" | b"true");
                                match attr.key.local_name().as_ref() {
                                    b"locked" => pending.locked = truthy,
                                    b"print" => pending.printable = truthy,
                                    b"altText" => {
                                        pending.alt_text =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned())
                                    }
                                    b"macro" => {
                                        pending.macro_name = Some(
                                            duke_sheets_vml::decode_macro_formula(
                                                &String::from_utf8_lossy(&attr.value),
                                            ),
                                        )
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"anchor" if current_control.is_some() => {
                            in_control_anchor = true;
                            control_anchor_move = false;
                            control_anchor_size = false;
                            control_anchor_vals = [[0i64; 4]; 2];
                            for attr in e.attributes().flatten() {
                                let truthy = matches!(&*attr.value, b"1" | b"true");
                                match attr.key.local_name().as_ref() {
                                    b"moveWithCells" => control_anchor_move = truthy,
                                    b"sizeWithCells" => control_anchor_size = truthy,
                                    _ => {}
                                }
                            }
                        }
                        b"from" if in_control_anchor => control_anchor_in_from = true,
                        b"to" if in_control_anchor => control_anchor_in_from = false,
                        b"col" | b"colOff" | b"row" | b"rowOff" if in_control_anchor => {
                            let field = match e.name().local_name().as_ref() {
                                b"col" => 0,
                                b"colOff" => 1,
                                b"row" => 2,
                                _ => 3,
                            };
                            let marker = if control_anchor_in_from { 0 } else { 1 };
                            control_anchor_field = Some((marker, field));
                            control_anchor_text.clear();
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
                        b"protectedRange" => {
                            current_protected_range = Self::parse_protected_range_element(&e);
                        }
                        b"securityDescriptor" if current_protected_range.is_some() => {
                            in_protected_range_security_descriptor = true;
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
                        b"controls" => in_controls = false,
                        b"control" if current_control.is_some() => {
                            let pending = current_control.take().unwrap();
                            if !pending.rid.is_empty() {
                                pending_controls.push(pending);
                            }
                        }
                        b"anchor" if in_control_anchor => {
                            in_control_anchor = false;
                            if let Some(pending) = current_control.as_mut() {
                                pending.anchor = Some(form_controls::anchor_from_markers(
                                    control_anchor_vals[0],
                                    control_anchor_vals[1],
                                    control_anchor_move,
                                    control_anchor_size,
                                ));
                            }
                        }
                        b"col" | b"colOff" | b"row" | b"rowOff"
                            if control_anchor_field.is_some() =>
                        {
                            let (marker, field) = control_anchor_field.take().unwrap();
                            control_anchor_vals[marker][field] =
                                control_anchor_text.trim().parse().unwrap_or(0);
                        }
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
                        b"securityDescriptor" if in_protected_range_security_descriptor => {
                            in_protected_range_security_descriptor = false;
                        }
                        b"protectedRange" => {
                            if let Some(protected_range) = current_protected_range.take() {
                                worksheet.add_protected_range(protected_range);
                            }
                            in_protected_range_security_descriptor = false;
                        }
                        _ => {}
                    }
                }
                Ok(Event::Text(e)) => {
                    if control_anchor_field.is_some() {
                        if let Ok(text) = e.unescape() {
                            control_anchor_text.push_str(&text);
                        }
                    }
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
                    } else if in_protected_range_security_descriptor {
                        if let (Some(ref mut protected_range), Ok(text)) =
                            (&mut current_protected_range, e.unescape())
                        {
                            protected_range.security_descriptor = Some(text.to_string());
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
                        b"protectedRange" => {
                            if let Some(protected_range) = Self::parse_protected_range_element(&e) {
                                worksheet.add_protected_range(protected_range);
                            }
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
                            let color = parse_color_element(&e);
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

        Ok(pending_controls)
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
    /// value `"1"` means the action is **not** allowed; `"0"` means it
    /// **is** allowed. The writer follows the same convention. The
    /// selectLockedCells/selectUnlockedCells attributes default to false
    /// in OOXML, so absence means selection is allowed.
    fn parse_sheet_protection_element(
        worksheet: &mut duke_sheets_core::Worksheet,
        e: &quick_xml::events::BytesStart<'_>,
    ) {
        use duke_sheets_core::worksheet::SheetProtection;

        let mut prot = SheetProtection {
            protected: true,
            select_locked_cells: true,
            select_unlocked_cells: true,
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

    fn parse_protected_range_element(
        e: &quick_xml::events::BytesStart<'_>,
    ) -> Option<duke_sheets_core::ProtectedRange> {
        let mut name = None;
        let mut ranges = Vec::new();
        let mut password_hash = None;
        let mut security_descriptor = None;

        for attr in e.attributes().flatten() {
            let Ok(value) = attr.unescape_value() else {
                continue;
            };
            match attr.key.local_name().as_ref() {
                b"name" => name = Some(value.to_string()),
                b"sqref" => {
                    for piece in value.split_whitespace() {
                        match CellRange::parse(piece) {
                            Ok(range) => ranges.push(range),
                            Err(err) => log::warn!(
                                "Invalid protectedRange sqref piece '{}': {}",
                                piece,
                                err
                            ),
                        }
                    }
                }
                b"password" => {
                    if let Ok(h) = u16::from_str_radix(value.as_ref(), 16) {
                        password_hash = Some(h);
                    }
                }
                b"securityDescriptor" => security_descriptor = Some(value.to_string()),
                _ => {}
            }
        }

        let name = name?;
        Some(duke_sheets_core::ProtectedRange {
            name,
            ranges,
            password_hash,
            security_descriptor,
        })
    }

    fn parse_hyperlink_element(
        worksheet: &mut duke_sheets_core::Worksheet,
        e: &quick_xml::events::BytesStart<'_>,
        sheet_rels: &RelationshipSet,
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
                if rel.kind() == Some(RelationshipKind::Hyperlink) {
                    target = rel.target().to_string();
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
        assert_eq!(comment.plain_text(), "Custom path comment");
    }
}
