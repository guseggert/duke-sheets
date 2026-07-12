use std::io::{Cursor, Read, Seek};

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::{CellMarker, DrawingAnchor, EditAs};
use duke_sheets_core::{ChildTransform, GroupTransform};

/// A chart reference discovered in a drawing XML, paired with its anchor position.
pub(crate) struct DrawingChartRef {
    /// The relationship id (e.g. "rId1") pointing to the chart part.
    pub(crate) rel_id: String,
    /// The two-cell anchor positioning the chart in the worksheet.
    pub(crate) anchor: DrawingAnchor,
    /// Whether this references a ChartEx part (`cx:chart`) rather than a standard chart.
    pub(crate) is_chart_ex: bool,
    /// Raw `mc:Fallback` inner XML bytes for roundtrip (chartEx only).
    pub(crate) raw_mc_fallback: Option<Vec<u8>>,
}

/// One top-level entry of a drawing part, in document order.
/// Document order carries z-order (back to front).
pub(crate) struct DrawingEntry {
    /// Captured anchor XML (the whole `<xdr:*Anchor>` element),
    /// preserved for raw passthrough.
    pub(crate) bytes: Vec<u8>,
    /// Parsed wrapper anchor.
    pub(crate) anchor: DrawingAnchor,
    /// `clientData/@fLocksWithSheet` (missing = true).
    pub(crate) locked: bool,
    /// `clientData/@fPrintsWithSheet` (missing = true).
    pub(crate) printable: bool,
    /// Classified payload.
    pub(crate) kind: DrawingEntryKind,
}

pub(crate) enum DrawingEntryKind {
    /// A graphicFrame chart or chartEx reference.
    Chart(DrawingChartRef),
    /// An `<xdr:pic>` picture.
    Image(Box<PicShape>),
    /// An `<xdr:sp>` with an `a14:compatExt` spid: the drawing twin
    /// of a legacy form control.
    ControlTwin(TwinShape),
    /// An `<xdr:grpSp>` whose content is fully modelable.
    Group(ParsedGroup),
    /// Anything else, preserved via `bytes`.
    Raw,
}

/// Parsed `<xdr:pic>` content. `blip_rel`/`svg_rel` hold relationship
/// ids until the caller resolves them against the drawing rels.
#[derive(Debug, Default)]
pub(crate) struct PicShape {
    pub(crate) name: String,
    pub(crate) descr: Option<String>,
    pub(crate) blip_rel: Option<String>,
    pub(crate) svg_rel: Option<String>,
    /// spPr/a:xfrm placement; off is meaningful for group children.
    pub(crate) off_x: i64,
    pub(crate) off_y: i64,
    pub(crate) ext_cx: i64,
    pub(crate) ext_cy: i64,
    pub(crate) rotation: Option<i32>,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
}

/// Parsed control-twin `<xdr:sp>`.
#[derive(Debug, Default)]
pub(crate) struct TwinShape {
    /// `a14:compatExt/@spid`, e.g. "_x0000_s1026".
    pub(crate) spid: String,
    /// The twin's numeric shape id, parsed from `spid`.
    pub(crate) shape_num: Option<u32>,
    /// spPr/a:xfrm placement (meaningful for group children).
    pub(crate) xfrm: ChildTransform,
}

/// Parsed `<xdr:grpSp>` content.
#[derive(Debug, Default)]
pub(crate) struct ParsedGroup {
    pub(crate) name: String,
    pub(crate) descr: Option<String>,
    /// grpSpPr/a:xfrm with child-space mapping.
    pub(crate) transform: GroupTransform,
    pub(crate) children: Vec<ParsedChild>,
}

#[derive(Debug)]
pub(crate) enum ParsedChild {
    Pic(PicShape),
    Group(ParsedGroup),
    Twin(TwinShape),
}

const URI_CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

/// Parse a SpreadsheetML drawing part into its top-level entries in
/// document order. Each anchor is owned by exactly one entry.
///
/// `mc:AlternateContent` is handled both at wsDr level (a Choice
/// wrapping whole anchors; the first parseable Choice wins, else the
/// Fallback) and inside an anchor (a Choice wrapping the chartEx
/// graphicFrame or the a14 control-twin shape).
pub(crate) fn read_drawing_entries<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<Vec<DrawingEntry>> {
    let mut file = match archive_by_name(archive, drawing_path) {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    parse_wsdr_fragment(&bytes)
}

/// Parse a sequence of anchors / wsDr-level `mc:AlternateContent`
/// elements. Used for the whole part and, recursively, for the inner
/// content of an `mc:Choice`/`mc:Fallback`.
fn parse_wsdr_fragment(bytes: &[u8]) -> XlsxResult<Vec<DrawingEntry>> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let mut entries = Vec::new();
    // Depth within elements we do not slice (wsDr itself).
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    let captured = capture_element(&mut reader, e)?;
                    entries.push(parse_anchor(&captured).unwrap_or_else(|| DrawingEntry {
                        bytes: captured,
                        anchor: DrawingAnchor::default(),
                        locked: true,
                        printable: true,
                        kind: DrawingEntryKind::Raw,
                    }));
                }
                b"AlternateContent" => {
                    entries.extend(parse_wsdr_alternate(&mut reader)?);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(entries)
}

/// Resolve a wsDr-level `mc:AlternateContent`: the first Choice whose
/// content yields entries wins; otherwise the Fallback content is
/// used. Never both.
fn parse_wsdr_alternate<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> XlsxResult<Vec<DrawingEntry>> {
    let mut buf = Vec::new();
    let mut chosen: Option<Vec<DrawingEntry>> = None;
    let mut fallback_bytes: Option<Vec<u8>> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"Choice" => {
                    let inner = capture_inner(reader)?;
                    if chosen.is_none() {
                        let parsed = parse_wsdr_fragment(&inner)?;
                        if !parsed.is_empty() {
                            chosen = Some(parsed);
                        }
                    }
                }
                b"Fallback" => {
                    fallback_bytes = Some(capture_inner(reader)?);
                }
                _ => skip_element(reader)?,
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"AlternateContent" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    if let Some(entries) = chosen {
        return Ok(entries);
    }
    match fallback_bytes {
        Some(bytes) => parse_wsdr_fragment(&bytes),
        None => Ok(Vec::new()),
    }
}

fn new_reader(bytes: &[u8]) -> Reader<&[u8]> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    reader
}

/// Capture a whole element (start tag already consumed and passed in)
/// as serialized XML bytes, including the wrapper tags.
fn capture_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart<'_>,
) -> XlsxResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Start(start.to_owned()))
        .map_err(std::io::Error::other)?;
    capture_until_end(reader, &mut w, 1)?;
    Ok(w.into_inner().into_inner())
}

/// Capture an element's inner content (start tag already consumed),
/// excluding the wrapper tags.
fn capture_inner<R: std::io::BufRead>(reader: &mut Reader<R>) -> XlsxResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    capture_until_end_inner(reader, &mut w)?;
    Ok(w.into_inner().into_inner())
}

fn capture_until_end<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    w: &mut Writer<Cursor<Vec<u8>>>,
    mut depth: u32,
) -> XlsxResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                depth += 1;
                w.write_event(Event::Start(e.into_owned()))
                    .map_err(std::io::Error::other)?;
            }
            Ok(Event::End(e)) => {
                depth -= 1;
                w.write_event(Event::End(e.into_owned()))
                    .map_err(std::io::Error::other)?;
                if depth == 0 {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => {
                return Err(XlsxError::InvalidFormat(
                    "unterminated drawing anchor".into(),
                ))
            }
            Ok(event) => {
                w.write_event(event.into_owned())
                    .map_err(std::io::Error::other)?;
            }
            Err(e) => return Err(XlsxError::Xml(e)),
        }
        buf.clear();
    }
}

/// Like [`capture_until_end`] but stops before writing the final End
/// event (inner content only; the wrapper End is consumed).
fn capture_until_end_inner<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    w: &mut Writer<Cursor<Vec<u8>>>,
) -> XlsxResult<()> {
    let mut depth: u32 = 1;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                depth += 1;
                w.write_event(Event::Start(e.into_owned()))
                    .map_err(std::io::Error::other)?;
            }
            Ok(Event::End(e)) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
                w.write_event(Event::End(e.into_owned()))
                    .map_err(std::io::Error::other)?;
            }
            Ok(Event::Eof) => {
                return Err(XlsxError::InvalidFormat(
                    "unterminated mc element in drawing".into(),
                ))
            }
            Ok(event) => {
                w.write_event(event.into_owned())
                    .map_err(std::io::Error::other)?;
            }
            Err(e) => return Err(XlsxError::Xml(e)),
        }
        buf.clear();
    }
}

/// Skip an element (start tag already consumed) through its End.
fn skip_element<R: std::io::BufRead>(reader: &mut Reader<R>) -> XlsxResult<()> {
    let mut depth: u32 = 1;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

fn attr_string(e: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == name {
            return attr.unescape_value().ok().map(|s| s.to_string());
        }
    }
    None
}

fn attr_i64(e: &BytesStart<'_>, name: &[u8]) -> Option<i64> {
    attr_string(e, name).and_then(|s| s.parse().ok())
}

fn attr_bool_default_true(e: &BytesStart<'_>, name: &[u8]) -> bool {
    match attr_string(e, name) {
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
        None => true,
    }
}

/// The content classification of an anchor's object element.
enum AnchorContent {
    None,
    Chart {
        rel_id: String,
        is_chart_ex: bool,
        raw_mc_fallback: Option<Vec<u8>>,
    },
    Pic(PicShape),
    Twin(TwinShape),
    Group(ParsedGroup),
    /// Present but unmodeled content: keep the whole anchor raw.
    Other,
}

/// Second pass: parse one captured anchor element into a classified
/// [`DrawingEntry`]. Returns `None` if the bytes are unparseable.
fn parse_anchor(bytes: &[u8]) -> Option<DrawingEntry> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();

    // Anchor start tag.
    let (variant, edit_as) = loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(ref e) => {
                let edit_as = attr_string(e, b"editAs").and_then(|s| match s.as_str() {
                    "twoCell" => Some(EditAs::TwoCell),
                    "oneCell" => Some(EditAs::OneCell),
                    "absolute" => Some(EditAs::Absolute),
                    _ => None,
                });
                break (e.local_name().as_ref().to_vec(), edit_as);
            }
            Event::Eof => return None,
            _ => {}
        }
    };
    buf.clear();

    let mut from = CellMarker::default();
    let mut to = CellMarker::default();
    let mut pos_x = 0i64;
    let mut pos_y = 0i64;
    let mut ext_cx = 0i64;
    let mut ext_cy = 0i64;
    let mut locked = true;
    let mut printable = true;
    let mut content = AnchorContent::None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"from" => from = parse_marker(&mut reader)?,
                    b"to" => to = parse_marker(&mut reader)?,
                    b"ext" => {
                        ext_cx = attr_i64(e, b"cx").unwrap_or(0);
                        ext_cy = attr_i64(e, b"cy").unwrap_or(0);
                        skip_element(&mut reader).ok()?;
                    }
                    b"pos" => {
                        pos_x = attr_i64(e, b"x").unwrap_or(0);
                        pos_y = attr_i64(e, b"y").unwrap_or(0);
                        skip_element(&mut reader).ok()?;
                    }
                    b"clientData" => {
                        locked = attr_bool_default_true(e, b"fLocksWithSheet");
                        printable = attr_bool_default_true(e, b"fPrintsWithSheet");
                        skip_element(&mut reader).ok()?;
                    }
                    _ => {
                        let parsed = parse_anchor_content(&mut reader, e).ok()?;
                        set_content(&mut content, parsed);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"ext" => {
                    ext_cx = attr_i64(e, b"cx").unwrap_or(0);
                    ext_cy = attr_i64(e, b"cy").unwrap_or(0);
                }
                b"pos" => {
                    pos_x = attr_i64(e, b"x").unwrap_or(0);
                    pos_y = attr_i64(e, b"y").unwrap_or(0);
                }
                b"clientData" => {
                    locked = attr_bool_default_true(e, b"fLocksWithSheet");
                    printable = attr_bool_default_true(e, b"fPrintsWithSheet");
                }
                _ => {}
            },
            Ok(Event::End(_)) | Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    let anchor = match variant.as_slice() {
        b"oneCellAnchor" => DrawingAnchor::OneCell {
            from: from.clone(),
            width_emu: ext_cx,
            height_emu: ext_cy,
        },
        b"absoluteAnchor" => DrawingAnchor::Absolute {
            x_emu: pos_x,
            y_emu: pos_y,
            width_emu: ext_cx,
            height_emu: ext_cy,
        },
        _ => DrawingAnchor::TwoCell {
            from: from.clone(),
            to: to.clone(),
            edit_as,
        },
    };

    let kind = match content {
        AnchorContent::Chart {
            rel_id,
            is_chart_ex,
            raw_mc_fallback,
        } => DrawingEntryKind::Chart(DrawingChartRef {
            rel_id,
            // Chart anchors normalize to two-cell markers (legacy
            // behavior kept for chartsheet/absolute variants).
            anchor: DrawingAnchor::TwoCell {
                from,
                to,
                edit_as: None,
            },
            is_chart_ex,
            raw_mc_fallback,
        }),
        AnchorContent::Pic(pic) if pic.blip_rel.is_some() => {
            DrawingEntryKind::Image(Box::new(pic))
        }
        AnchorContent::Twin(twin) => DrawingEntryKind::ControlTwin(twin),
        AnchorContent::Group(group) => DrawingEntryKind::Group(group),
        _ => DrawingEntryKind::Raw,
    };

    Some(DrawingEntry {
        bytes: bytes.to_vec(),
        anchor,
        locked,
        printable,
        kind,
    })
}

/// Merge a newly parsed content element into the anchor's content
/// slot. A second object element degrades the anchor to raw.
fn set_content(slot: &mut AnchorContent, parsed: AnchorContent) {
    match (&slot, parsed) {
        (_, AnchorContent::None) => {}
        (AnchorContent::None, parsed) => *slot = parsed,
        _ => *slot = AnchorContent::Other,
    }
}

/// Parse one anchor content element (start tag consumed).
fn parse_anchor_content<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart<'_>,
) -> XlsxResult<AnchorContent> {
    let name = start.local_name().as_ref().to_vec();
    match name.as_slice() {
        b"graphicFrame" => Ok(match parse_graphic_frame(reader)? {
            Some((rel_id, is_chart_ex)) => AnchorContent::Chart {
                rel_id,
                is_chart_ex,
                raw_mc_fallback: None,
            },
            None => AnchorContent::Other,
        }),
        b"pic" => {
            let pic = parse_pic(reader)?;
            if pic.blip_rel.is_some() {
                Ok(AnchorContent::Pic(pic))
            } else {
                Ok(AnchorContent::Other)
            }
        }
        b"sp" => Ok(match parse_sp(reader)? {
            Some(twin) => AnchorContent::Twin(twin),
            None => AnchorContent::Other,
        }),
        b"grpSp" => Ok(match parse_group(reader)? {
            Some(group) => AnchorContent::Group(group),
            None => AnchorContent::Other,
        }),
        b"AlternateContent" => parse_anchor_alternate(reader),
        _ => {
            skip_element(reader)?;
            Ok(AnchorContent::Other)
        }
    }
}

/// `mc:AlternateContent` inside an anchor: a Choice wrapping the
/// chartEx graphicFrame or the a14 control-twin `xdr:sp`. The first
/// Choice with recognized content wins; the Fallback content is
/// preserved only for chartEx round-trip.
fn parse_anchor_alternate<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> XlsxResult<AnchorContent> {
    let mut buf = Vec::new();
    let mut chosen = AnchorContent::None;
    let mut fallback: Option<Vec<u8>> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"Choice" => {
                    let inner = capture_inner(reader)?;
                    if matches!(chosen, AnchorContent::None) {
                        chosen = parse_choice_content(&inner)?;
                    }
                }
                b"Fallback" => {
                    fallback = Some(capture_inner(reader)?);
                }
                _ => skip_element(reader)?,
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"AlternateContent" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    if let AnchorContent::Chart {
        rel_id,
        is_chart_ex,
        ..
    } = chosen
    {
        return Ok(AnchorContent::Chart {
            rel_id,
            is_chart_ex,
            raw_mc_fallback: fallback,
        });
    }
    match chosen {
        AnchorContent::None => Ok(AnchorContent::Other),
        other => Ok(other),
    }
}

/// Parse the inner content of an in-anchor `mc:Choice`.
fn parse_choice_content(bytes: &[u8]) -> XlsxResult<AnchorContent> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let e = e.to_owned();
                return parse_anchor_content(&mut reader, &e);
            }
            Ok(Event::Eof) => return Ok(AnchorContent::None),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

/// Parse `<xdr:graphicFrame>` content (start consumed); returns the
/// chart relationship id and chartEx-ness when a chart URI is found.
fn parse_graphic_frame<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> XlsxResult<Option<(String, bool)>> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut graphic_data_uri: Option<String> = None;
    let mut found: Option<(String, bool)> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                if e.local_name().as_ref() == b"graphicData" {
                    graphic_data_uri = attr_string(e, b"uri");
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"chart" {
                    let is_chart_ex = graphic_data_uri.as_deref() == Some(URI_CHART_EX);
                    if let Some(id) = attr_string(e, b"id") {
                        found = Some((id, is_chart_ex));
                    }
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(found);
                }
            }
            Ok(Event::Eof) => return Ok(found),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

/// Parse `<xdr:pic>` content (start consumed).
fn parse_pic<R: std::io::BufRead>(reader: &mut Reader<R>) -> XlsxResult<PicShape> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut pic = PicShape::default();
    let mut in_sp_pr = false;
    let mut in_xfrm = false;

    let handle = |e: &BytesStart<'_>, pic: &mut PicShape, in_sp_pr: bool, in_xfrm: bool| {
        match e.local_name().as_ref() {
            b"cNvPr" => {
                if let Some(name) = attr_string(e, b"name") {
                    pic.name = name;
                }
                if let Some(descr) = attr_string(e, b"descr") {
                    pic.descr = Some(descr);
                }
            }
            b"blip" => {
                if let Some(embed) = attr_string(e, b"embed") {
                    pic.blip_rel = Some(embed);
                }
            }
            b"svgBlip" => {
                pic.svg_rel = attr_string(e, b"embed");
            }
            b"xfrm" if in_sp_pr => {
                pic.rotation = attr_i64(e, b"rot").map(|v| v as i32);
                pic.flip_h = matches!(attr_string(e, b"flipH").as_deref(), Some("1") | Some("true"));
                pic.flip_v = matches!(attr_string(e, b"flipV").as_deref(), Some("1") | Some("true"));
            }
            b"off" if in_xfrm => {
                pic.off_x = attr_i64(e, b"x").unwrap_or(0);
                pic.off_y = attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if in_xfrm => {
                pic.ext_cx = attr_i64(e, b"cx").unwrap_or(0);
                pic.ext_cy = attr_i64(e, b"cy").unwrap_or(0);
            }
            _ => {}
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                match e.local_name().as_ref() {
                    b"spPr" => in_sp_pr = true,
                    b"xfrm" if in_sp_pr => in_xfrm = true,
                    _ => {}
                }
                handle(e, &mut pic, in_sp_pr, in_xfrm);
            }
            Ok(Event::Empty(ref e)) => handle(e, &mut pic, in_sp_pr, in_xfrm),
            Ok(Event::End(ref e)) => {
                depth -= 1;
                match e.local_name().as_ref() {
                    b"spPr" => in_sp_pr = false,
                    b"xfrm" => in_xfrm = false,
                    _ => {}
                }
                if depth == 0 {
                    return Ok(pic);
                }
            }
            Ok(Event::Eof) => return Ok(pic),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

/// Parse `<xdr:sp>` content (start consumed). Returns a twin when an
/// `a14:compatExt` spid is present, `None` otherwise (plain shape).
fn parse_sp<R: std::io::BufRead>(reader: &mut Reader<R>) -> XlsxResult<Option<TwinShape>> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut twin = TwinShape::default();
    let mut has_spid = false;
    let mut in_sp_pr = false;
    let mut in_xfrm = false;

    let handle = |e: &BytesStart<'_>,
                      twin: &mut TwinShape,
                      has_spid: &mut bool,
                      in_sp_pr: bool,
                      in_xfrm: bool| {
        match e.local_name().as_ref() {
            b"xfrm" if in_sp_pr => {
                twin.xfrm.rotation = attr_i64(e, b"rot").map(|v| v as i32).unwrap_or(0);
                twin.xfrm.flip_h =
                    matches!(attr_string(e, b"flipH").as_deref(), Some("1") | Some("true"));
                twin.xfrm.flip_v =
                    matches!(attr_string(e, b"flipV").as_deref(), Some("1") | Some("true"));
            }
            b"off" if in_xfrm => {
                twin.xfrm.x_emu = attr_i64(e, b"x").unwrap_or(0);
                twin.xfrm.y_emu = attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if in_xfrm => {
                twin.xfrm.cx_emu = attr_i64(e, b"cx").unwrap_or(0);
                twin.xfrm.cy_emu = attr_i64(e, b"cy").unwrap_or(0);
            }
            b"compatExt" => {
                if let Some(spid) = attr_string(e, b"spid") {
                    twin.shape_num = spid.rsplit(['s', 'S']).next().and_then(|n| n.parse().ok());
                    twin.spid = spid;
                    *has_spid = true;
                }
            }
            _ => {}
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                match e.local_name().as_ref() {
                    b"spPr" => in_sp_pr = true,
                    b"xfrm" if in_sp_pr => in_xfrm = true,
                    _ => {}
                }
                handle(e, &mut twin, &mut has_spid, in_sp_pr, in_xfrm);
            }
            Ok(Event::Empty(ref e)) => handle(e, &mut twin, &mut has_spid, in_sp_pr, in_xfrm),
            Ok(Event::End(ref e)) => {
                depth -= 1;
                match e.local_name().as_ref() {
                    b"spPr" => in_sp_pr = false,
                    b"xfrm" => in_xfrm = false,
                    _ => {}
                }
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(has_spid.then_some(twin))
}

/// Parse `<xdr:grpSp>` content (start consumed). Returns `None` when
/// any child is not modelable (the caller keeps the anchor raw).
fn parse_group<R: std::io::BufRead>(reader: &mut Reader<R>) -> XlsxResult<Option<ParsedGroup>> {
    let mut buf = Vec::new();
    let mut group = ParsedGroup::default();
    let mut modelable = true;
    let mut in_nv = false;
    let mut in_grp_sp_pr = false;
    let mut in_xfrm = false;

    let handle_xfrm_child = |e: &BytesStart<'_>, transform: &mut GroupTransform| {
        match e.local_name().as_ref() {
            b"off" => {
                transform.x_emu = attr_i64(e, b"x").unwrap_or(0);
                transform.y_emu = attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" => {
                transform.cx_emu = attr_i64(e, b"cx").unwrap_or(0);
                transform.cy_emu = attr_i64(e, b"cy").unwrap_or(0);
            }
            b"chOff" => {
                transform.child_x_emu = attr_i64(e, b"x").unwrap_or(0);
                transform.child_y_emu = attr_i64(e, b"y").unwrap_or(0);
            }
            b"chExt" => {
                transform.child_cx_emu = attr_i64(e, b"cx").unwrap_or(0);
                transform.child_cy_emu = attr_i64(e, b"cy").unwrap_or(0);
            }
            _ => {}
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"nvGrpSpPr" => in_nv = true,
                    b"grpSpPr" => in_grp_sp_pr = true,
                    b"xfrm" if in_grp_sp_pr => {
                        in_xfrm = true;
                        parse_group_xfrm_attrs(e, &mut group.transform);
                    }
                    b"off" | b"ext" | b"chOff" | b"chExt" if in_xfrm => {
                        handle_xfrm_child(e, &mut group.transform);
                    }
                    b"cNvPr" if in_nv => parse_group_cnvpr(e, &mut group),
                    b"pic" if !in_nv && !in_grp_sp_pr => match parse_pic(reader)? {
                        pic if pic.blip_rel.is_some() => {
                            group.children.push(ParsedChild::Pic(pic))
                        }
                        _ => modelable = false,
                    },
                    b"sp" if !in_nv && !in_grp_sp_pr => match parse_sp(reader)? {
                        Some(twin) => group.children.push(ParsedChild::Twin(twin)),
                        None => modelable = false,
                    },
                    b"grpSp" if !in_nv && !in_grp_sp_pr => match parse_group(reader)? {
                        Some(inner) => group.children.push(ParsedChild::Group(inner)),
                        None => modelable = false,
                    },
                    b"AlternateContent" if !in_nv && !in_grp_sp_pr => {
                        match parse_anchor_alternate(reader)? {
                            AnchorContent::Twin(twin) => {
                                group.children.push(ParsedChild::Twin(twin))
                            }
                            AnchorContent::Pic(pic) if pic.blip_rel.is_some() => {
                                group.children.push(ParsedChild::Pic(pic))
                            }
                            AnchorContent::Group(inner) => {
                                group.children.push(ParsedChild::Group(inner))
                            }
                            _ => modelable = false,
                        }
                    }
                    _ => {
                        // grpSpPr/nvGrpSpPr subtrees are handled via
                        // the flags; anything else is unmodeled.
                        if !in_grp_sp_pr && !in_nv {
                            modelable = false;
                            skip_element(reader)?;
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                    b"cNvPr" if in_nv => parse_group_cnvpr(e, &mut group),
                    b"xfrm" if in_grp_sp_pr => {
                        parse_group_xfrm_attrs(e, &mut group.transform)
                    }
                    b"off" | b"ext" | b"chOff" | b"chExt" if in_xfrm => {
                        handle_xfrm_child(e, &mut group.transform);
                    }
                    b"pic" | b"sp" | b"grpSp" | b"cxnSp" | b"graphicFrame" | b"contentPart"
                        if !in_nv && !in_grp_sp_pr =>
                    {
                        modelable = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"nvGrpSpPr" => in_nv = false,
                b"grpSpPr" => in_grp_sp_pr = false,
                b"xfrm" => in_xfrm = false,
                b"grpSp" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(modelable.then_some(group))
}

fn parse_group_cnvpr(e: &BytesStart<'_>, group: &mut ParsedGroup) {
    if let Some(name) = attr_string(e, b"name") {
        group.name = name;
    }
    if let Some(descr) = attr_string(e, b"descr") {
        group.descr = Some(descr);
    }
}

fn parse_group_xfrm_attrs(e: &BytesStart<'_>, transform: &mut GroupTransform) {
    transform.rotation = attr_i64(e, b"rot").map(|v| v as i32).unwrap_or(0);
    transform.flip_h = matches!(attr_string(e, b"flipH").as_deref(), Some("1") | Some("true"));
    transform.flip_v = matches!(attr_string(e, b"flipV").as_deref(), Some("1") | Some("true"));
}

/// Parse an `xdr:from`/`xdr:to` marker element (start consumed).
fn parse_marker<R: std::io::BufRead>(reader: &mut Reader<R>) -> Option<CellMarker> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut marker = CellMarker::default();
    let mut field: Option<&'static str> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                field = match e.local_name().as_ref() {
                    b"col" => Some("col"),
                    b"colOff" => Some("colOff"),
                    b"row" => Some("row"),
                    b"rowOff" => Some("rowOff"),
                    _ => None,
                };
            }
            Ok(Event::Text(ref t)) => {
                if let Ok(text) = t.unescape() {
                    let text = text.trim();
                    match field {
                        Some("col") => marker.col = text.parse().unwrap_or(0),
                        Some("colOff") => marker.col_offset_emu = text.parse().unwrap_or(0),
                        Some("row") => marker.row = text.parse().unwrap_or(0),
                        Some("rowOff") => marker.row_offset_emu = text.parse().unwrap_or(0),
                        _ => {}
                    }
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                field = None;
                if depth == 0 {
                    return Some(marker);
                }
            }
            Ok(Event::Eof) => return Some(marker),
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

/// Backward-compatible view returning only chart refs.
#[cfg(test)]
pub(crate) fn read_drawing_chart_refs<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<Vec<DrawingChartRef>> {
    Ok(read_drawing_entries(archive, drawing_path)?
        .into_iter()
        .filter_map(|entry| match entry.kind {
            DrawingEntryKind::Chart(chart_ref) => Some(chart_ref),
            _ => None,
        })
        .collect())
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;

    fn zip_with_entry(path: &str, xml: &str) -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(path, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        let cursor = zip.finish().unwrap();
        zip::ZipArchive::new(cursor).unwrap()
    }

    #[test]
    fn test_parse_drawing_with_chart_ref() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>100</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>200</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col>
      <xdr:colOff>300</xdr:colOff>
      <xdr:row>20</xdr:row>
      <xdr:rowOff>400</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].rel_id, "rId1");
        assert!(!refs[0].is_chart_ex);
        if let DrawingAnchor::TwoCell { from, to, .. } = &refs[0].anchor {
            assert_eq!(from.col, 1);
            assert_eq!(from.col_offset_emu, 100);
            assert_eq!(from.row, 2);
            assert_eq!(from.row_offset_emu, 200);
            assert_eq!(to.col, 10);
            assert_eq!(to.col_offset_emu, 300);
            assert_eq!(to.row, 20);
            assert_eq!(to.row_offset_emu, 400);
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_parse_drawing_with_one_cell_anchor() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>100</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>200</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="5000000" cy="3000000"/>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].rel_id, "rId2");
        assert!(!refs[0].is_chart_ex);
        if let DrawingAnchor::TwoCell { from, to, .. } = &refs[0].anchor {
            assert_eq!(from.col, 1);
            assert_eq!(from.col_offset_emu, 100);
            assert_eq!(from.row, 2);
            assert_eq!(from.row_offset_emu, 200);
            assert_eq!(to.col, 0);
            assert_eq!(to.row, 0);
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_drawing_without_chart_is_ignored() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert!(refs.is_empty());
    }

    #[test]
    fn test_drawing_missing_file() {
        let mut archive = zip_with_entry("other.xml", "<dummy/>");
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert!(refs.is_empty());
    }

    #[test]
    fn test_parse_drawing_with_absolute_anchor() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/>
    <xdr:ext cx="9144000" cy="6858000"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].rel_id, "rId1");
        assert!(!refs[0].is_chart_ex);
        // absoluteAnchor defaults all anchor values to zero
        if let DrawingAnchor::TwoCell { from, to, .. } = &refs[0].anchor {
            assert_eq!(from.col, 0);
            assert_eq!(from.row, 0);
            assert_eq!(to.col, 0);
            assert_eq!(to.row, 0);
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_parse_drawing_with_chart_ex_ref() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1">
        <xdr:graphicFrame>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
              <cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rId3"/>
            </a:graphicData>
          </a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback>
        <xdr:sp><xdr:txBody><a:p><a:r><a:t>Fallback text</a:t></a:r></a:p></xdr:txBody></xdr:sp>
      </mc:Fallback>
    </mc:AlternateContent>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let refs = read_drawing_chart_refs(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].rel_id, "rId3");
        assert!(refs[0].is_chart_ex);
        // Fallback inner content is captured without the wrapper tags.
        let fallback = refs[0].raw_mc_fallback.as_deref().expect("fallback");
        let fallback = std::str::from_utf8(fallback).unwrap();
        assert!(fallback.starts_with("<xdr:sp>"), "{fallback}");
        assert!(!fallback.contains("mc:Fallback"), "{fallback}");
    }

    #[test]
    fn test_parse_drawing_with_pic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>200</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>300</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>400</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1" descr="A test image"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm rot="5400000" flipH="1">
          <a:off x="0" y="0"/>
          <a:ext cx="1000000" cy="2000000"/>
        </a:xfrm>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        // Single ownership: exactly one entry, classified as an image.
        assert_eq!(entries.len(), 1);
        let entry = &entries[0];
        let DrawingEntryKind::Image(pic) = &entry.kind else {
            panic!("expected image entry");
        };
        assert_eq!(pic.name, "Picture 1");
        assert_eq!(pic.descr.as_deref(), Some("A test image"));
        assert_eq!(pic.blip_rel.as_deref(), Some("rId1"));
        assert_eq!(pic.ext_cx, 1000000);
        assert_eq!(pic.ext_cy, 2000000);
        assert_eq!(pic.rotation, Some(5400000));
        assert!(pic.flip_h);
        assert!(!pic.flip_v);
        assert!(pic.svg_rel.is_none());
        assert!(entry.locked);
        assert!(entry.printable);

        if let DrawingAnchor::TwoCell { from, to, .. } = &entry.anchor {
            assert_eq!(from.col, 1);
            assert_eq!(from.col_offset_emu, 100);
            assert_eq!(from.row, 2);
            assert_eq!(from.row_offset_emu, 200);
            assert_eq!(to.col, 5);
            assert_eq!(to.col_offset_emu, 300);
            assert_eq!(to.row, 10);
            assert_eq!(to.row_offset_emu, 400);
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_parse_drawing_with_pic_and_svg() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="SVG Pic" descr="Has SVG"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2">
          <a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
            <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                          r:embed="rId3"/>
          </a:ext></a:extLst>
        </a:blip>
      </xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Image(pic) = &entries[0].kind else {
            panic!("expected image entry");
        };
        assert_eq!(pic.name, "SVG Pic");
        assert_eq!(pic.blip_rel.as_deref(), Some("rId2"));
        assert_eq!(pic.svg_rel.as_deref(), Some("rId3"));
        assert_eq!(pic.ext_cx, 500000);
        assert_eq!(pic.ext_cy, 500000);
    }

    #[test]
    fn test_parse_control_twin_and_client_data() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
      <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="3" name="Check Box 1" hidden="1">
              <a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1026"/></a:ext></a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto">
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </xdr:spPr>
        </xdr:sp>
        <xdr:clientData fLocksWithSheet="0" fPrintsWithSheet="0"/>
      </xdr:twoCellAnchor>
    </mc:Choice>
    <mc:Fallback/>
  </mc:AlternateContent>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let entry = &entries[0];
        let DrawingEntryKind::ControlTwin(twin) = &entry.kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.spid, "_x0000_s1026");
        assert_eq!(twin.shape_num, Some(1026));
        assert!(!entry.locked);
        assert!(!entry.printable);
        if let DrawingAnchor::TwoCell { edit_as, .. } = &entry.anchor {
            assert_eq!(edit_as, &Some(EditAs::OneCell));
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_parse_control_twin_inside_anchor_alternate_content() {
        // The other Excel emission shape: a bare anchor whose content
        // is mc:AlternateContent wrapping the a14 twin sp.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="3" name="Check Box 7" hidden="1">
              <a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1031"/></a:ext></a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></xdr:spPr>
        </xdr:sp>
      </mc:Choice>
      <mc:Fallback/>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::ControlTwin(twin) = &entries[0].kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.shape_num, Some(1031));
    }

    #[test]
    fn test_parse_group_of_pics() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
      <xdr:grpSpPr>
        <a:xfrm>
          <a:off x="609600" y="190500"/><a:ext cx="1219200" cy="190500"/>
          <a:chOff x="0" y="0"/><a:chExt cx="1219200" cy="190500"/>
        </a:xfrm>
      </xdr:grpSpPr>
      <xdr:pic>
        <xdr:nvPicPr><xdr:cNvPr id="5" name="Left"/><xdr:cNvPicPr/></xdr:nvPicPr>
        <xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></xdr:blipFill>
        <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm></xdr:spPr>
      </xdr:pic>
      <xdr:pic>
        <xdr:nvPicPr><xdr:cNvPr id="6" name="Right"/><xdr:cNvPicPr/></xdr:nvPicPr>
        <xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/></xdr:blipFill>
        <xdr:spPr><a:xfrm><a:off x="609600" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm></xdr:spPr>
      </xdr:pic>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Group(group) = &entries[0].kind else {
            panic!("expected group entry");
        };
        assert_eq!(group.name, "Group 1");
        assert_eq!(group.transform.x_emu, 609600);
        assert_eq!(group.transform.child_cx_emu, 1219200);
        assert_eq!(group.children.len(), 2);
        let ParsedChild::Pic(right) = &group.children[1] else {
            panic!("expected pic child");
        };
        assert_eq!(right.name, "Right");
        assert_eq!(right.off_x, 609600);
        assert_eq!(right.blip_rel.as_deref(), Some("rId2"));
    }

    #[test]
    fn test_group_with_unmodeled_child_degrades_to_raw() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
      <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/><a:chOff x="0" y="0"/><a:chExt cx="1" cy="1"/></a:xfrm></xdr:grpSpPr>
      <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>
        <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr>
        <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>hi</a:t></a:r></a:p></xdr:txBody>
      </xdr:sp>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert_eq!(entries.len(), 1);
        assert!(matches!(entries[0].kind, DrawingEntryKind::Raw));
        let raw = std::str::from_utf8(&entries[0].bytes).unwrap();
        assert!(raw.contains("TextBox 1"), "raw bytes keep the group: {raw}");
    }
}
