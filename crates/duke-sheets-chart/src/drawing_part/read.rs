//! Drawing-part reader: parse a `wsDr` XML fragment into classified
//! top-level entries in document order.

use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;

use super::{
    ShapeFill, ShapeLine, TwinColor, TwinHorizontalAlignment, TwinRunFont, TwinText, TwinTextRun,
    TwinUnderline, TwinVerticalAlignment,
};
use crate::error::{ChartParseError, ChartParseResult};
use crate::{CellMarker, ChildTransform, DrawingAnchor, EditAs, GroupTransform};

/// A chart reference discovered in a drawing XML, paired with its anchor position.
pub struct DrawingChartRef {
    /// The relationship id (e.g. "rId1") pointing to the chart part.
    pub rel_id: String,
    /// The anchor positioning the chart in the worksheet (variant
    /// preserved: TwoCell keeps editAs, OneCell/Absolute keep extents).
    pub anchor: DrawingAnchor,
    /// Whether this references a ChartEx part (`cx:chart`) rather than a standard chart.
    pub is_chart_ex: bool,
    /// Raw `mc:Fallback` inner XML bytes for roundtrip (chartEx only).
    pub raw_mc_fallback: Option<Vec<u8>>,
    pub name: Option<String>,
    pub descr: Option<String>,
    pub title: Option<String>,
    /// The frame's `cNvPr/@hidden` (absent = false).
    pub hidden: bool,
}

/// One top-level entry of a drawing part, in document order.
/// Document order carries z-order (back to front).
pub struct DrawingEntry {
    /// Captured anchor XML (the whole `<xdr:*Anchor>` element),
    /// preserved for raw passthrough.
    pub bytes: Vec<u8>,
    /// Parsed wrapper anchor.
    pub anchor: DrawingAnchor,
    /// `clientData/@fLocksWithSheet` (missing = true).
    pub locked: bool,
    /// `clientData/@fPrintsWithSheet` (missing = true).
    pub printable: bool,
    /// Classified payload.
    pub kind: DrawingEntryKind,
}

pub enum DrawingEntryKind {
    /// A graphicFrame chart or chartEx reference.
    Chart(DrawingChartRef),
    /// An `<xdr:pic>` picture.
    Image(Box<PicShape>),
    /// An ordinary `<xdr:sp>` worksheet shape.
    Shape(Box<ParsedShape>),
    /// The drawing twin of a legacy form control: an `<xdr:sp>` with
    /// an `a14:compatExt` spid (XLSX flavor) or an `<xdr:graphicFrame>`
    /// with a `com14:compatSp` spid (XLSB flavor).
    ControlTwin(TwinShape),
    /// An `<xdr:grpSp>` whose content is fully modelable.
    Group(ParsedGroup),
    /// Anything else, preserved via `bytes`.
    Raw,
}

/// Parsed `<xdr:pic>` content. `blip_rel`/`svg_rel` hold relationship
/// ids until the caller resolves them against the drawing rels.
#[derive(Debug, Default)]
pub struct PicShape {
    pub name: String,
    pub descr: Option<String>,
    pub title: Option<String>,
    /// `cNvPr/@hidden` (absent = false).
    pub hidden: bool,
    pub blip_rel: Option<String>,
    pub svg_rel: Option<String>,
    /// spPr/a:xfrm placement; off is meaningful for group children.
    pub off_x: i64,
    pub off_y: i64,
    pub ext_cx: i64,
    pub ext_cy: i64,
    pub rotation: Option<i32>,
    pub flip_h: bool,
    pub flip_v: bool,
}

/// Parsed control-twin shape (either flavor).
#[derive(Debug, Default)]
pub struct TwinShape {
    /// `a14:compatExt/@spid` or `com14:compatSp/@spid`, e.g. "_x0000_s1026".
    pub spid: String,
    /// The twin's numeric shape id, parsed from `spid`.
    pub shape_num: Option<u32>,
    /// `cNvPr/@name`, the control's display name.
    pub name: Option<String>,
    pub descr: Option<String>,
    pub title: Option<String>,
    pub macro_name: Option<String>,
    pub text: Option<TwinText>,
    /// Placement (meaningful for group children).
    pub xfrm: ChildTransform,
}

/// Parsed ordinary `<xdr:sp>` content.
#[derive(Debug, Default)]
pub struct ParsedShape {
    pub name: String,
    pub descr: Option<String>,
    pub title: Option<String>,
    pub hidden: bool,
    pub xfrm: ChildTransform,
    pub geometry: String,
    pub fill: ShapeFill,
    pub line: ShapeLine,
    pub text: Option<TwinText>,
    /// Unmodeled direct children of `spPr`, serialized as XML fragments.
    pub raw_shape_properties: Option<Vec<u8>>,
    /// Complete parsed `txBody`, retained for untouched round-trip.
    pub raw_text_body: Option<Vec<u8>>,
}

/// Parsed `<xdr:grpSp>` content.
#[derive(Debug, Default)]
pub struct ParsedGroup {
    pub name: String,
    pub descr: Option<String>,
    pub title: Option<String>,
    /// `cNvPr/@hidden` (absent = false).
    pub hidden: bool,
    /// grpSpPr/a:xfrm with child-space mapping.
    pub transform: GroupTransform,
    pub children: Vec<ParsedChild>,
}

#[derive(Debug)]
pub enum ParsedChild {
    Pic(PicShape),
    Shape(ParsedShape),
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
/// graphicFrame or a control-twin shape of either flavor).
pub fn parse_drawing_part(bytes: &[u8]) -> ChartParseResult<Vec<DrawingEntry>> {
    parse_wsdr_fragment(bytes, &[])
}

/// `xmlns` / `xmlns:*` declarations in scope from sliced-away wrapper
/// elements (mc:AlternateContent, mc:Choice, mc:Fallback), as raw
/// (attribute key, escaped value) pairs.
type NsDecls = Vec<(Vec<u8>, Vec<u8>)>;

fn collect_ns_decls(e: &BytesStart<'_>) -> NsDecls {
    e.attributes()
        .flatten()
        .filter(|attr| {
            attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:")
        })
        .map(|attr| (attr.key.as_ref().to_vec(), attr.value.into_owned()))
        .collect()
}

/// Merge wrapper declarations, the inner element's shadowing the
/// outer's.
fn merge_ns_decls(outer: &[(Vec<u8>, Vec<u8>)], inner: NsDecls) -> NsDecls {
    let mut merged = outer.to_vec();
    for (key, value) in inner {
        match merged.iter_mut().find(|(k, _)| *k == key) {
            Some(slot) => slot.1 = value,
            None => merged.push((key, value)),
        }
    }
    merged
}

/// Parse a sequence of anchors / wsDr-level `mc:AlternateContent`
/// elements. Used for the whole part and, recursively, for the inner
/// content of an `mc:Choice`/`mc:Fallback`; `ambient` carries the
/// xmlns declarations of the sliced-away wrappers so captured anchor
/// bytes stay namespace-complete.
fn parse_wsdr_fragment(
    bytes: &[u8],
    ambient: &[(Vec<u8>, Vec<u8>)],
) -> ChartParseResult<Vec<DrawingEntry>> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let mut entries = Vec::new();
    // Depth within elements we do not slice (wsDr itself).
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    let captured = capture_element(&mut reader, e, ambient)?;
                    entries.push(parse_anchor(&captured).unwrap_or_else(|| DrawingEntry {
                        bytes: captured,
                        anchor: DrawingAnchor::default(),
                        locked: true,
                        printable: true,
                        kind: DrawingEntryKind::Raw,
                    }));
                }
                b"AlternateContent" => {
                    let scope = merge_ns_decls(ambient, collect_ns_decls(e));
                    entries.extend(parse_wsdr_alternate(&mut reader, &scope)?);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
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
    ambient: &[(Vec<u8>, Vec<u8>)],
) -> ChartParseResult<Vec<DrawingEntry>> {
    let mut buf = Vec::new();
    let mut chosen: Option<Vec<DrawingEntry>> = None;
    let mut fallback: Option<(Vec<u8>, NsDecls)> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"Choice" => {
                    let scope = merge_ns_decls(ambient, collect_ns_decls(e));
                    let inner = capture_inner(reader)?;
                    if chosen.is_none() {
                        let parsed = parse_wsdr_fragment(&inner, &scope)?;
                        if !parsed.is_empty() {
                            chosen = Some(parsed);
                        }
                    }
                }
                b"Fallback" => {
                    let scope = merge_ns_decls(ambient, collect_ns_decls(e));
                    fallback = Some((capture_inner(reader)?, scope));
                }
                _ => skip_element(reader)?,
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"AlternateContent" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    if let Some(entries) = chosen {
        return Ok(entries);
    }
    match fallback {
        Some((bytes, scope)) => parse_wsdr_fragment(&bytes, &scope),
        None => Ok(Vec::new()),
    }
}

fn new_reader(bytes: &[u8]) -> Reader<&[u8]> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader
}

/// Capture a whole element (start tag already consumed and passed in)
/// as serialized XML bytes, including the wrapper tags. Ambient xmlns
/// declarations from sliced-away wrappers are injected into the root
/// start tag unless it already declares the same prefix.
fn capture_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart<'_>,
    ambient: &[(Vec<u8>, Vec<u8>)],
) -> ChartParseResult<Vec<u8>> {
    let mut root = start.to_owned();
    for (key, value) in ambient {
        let declared = start
            .attributes()
            .flatten()
            .any(|attr| attr.key.as_ref() == key.as_slice());
        if !declared {
            root.push_attribute((key.as_slice(), value.as_slice()));
        }
    }
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Start(root))
        .map_err(std::io::Error::other)?;
    capture_until_end(reader, &mut w, 1)?;
    Ok(w.into_inner().into_inner())
}

/// Capture an element's inner content (start tag already consumed),
/// excluding the wrapper tags.
fn capture_inner<R: std::io::BufRead>(reader: &mut Reader<R>) -> ChartParseResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    capture_until_end_inner(reader, &mut w)?;
    Ok(w.into_inner().into_inner())
}

fn capture_until_end<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    w: &mut Writer<Cursor<Vec<u8>>>,
    mut depth: u32,
) -> ChartParseResult<()> {
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
                return Err(ChartParseError::Parse("unterminated drawing anchor".into()))
            }
            Ok(event) => {
                w.write_event(event.into_owned())
                    .map_err(std::io::Error::other)?;
            }
            Err(e) => return Err(ChartParseError::Xml(e)),
        }
        buf.clear();
    }
}

/// Like [`capture_until_end`] but stops before writing the final End
/// event (inner content only; the wrapper End is consumed).
fn capture_until_end_inner<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    w: &mut Writer<Cursor<Vec<u8>>>,
) -> ChartParseResult<()> {
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
                return Err(ChartParseError::Parse(
                    "unterminated mc element in drawing".into(),
                ))
            }
            Ok(event) => {
                w.write_event(event.into_owned())
                    .map_err(std::io::Error::other)?;
            }
            Err(e) => return Err(ChartParseError::Xml(e)),
        }
        buf.clear();
    }
}

/// Skip an element (start tag already consumed) through its End.
fn skip_element<R: std::io::BufRead>(reader: &mut Reader<R>) -> ChartParseResult<()> {
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
            Err(e) => return Err(ChartParseError::Xml(e)),
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

fn attr_bool_default_false(e: &BytesStart<'_>, name: &[u8]) -> bool {
    matches!(
        attr_string(e, name).as_deref(),
        Some("1") | Some("true") | Some("TRUE") | Some("True")
    )
}

/// The content classification of an anchor's object element.
enum AnchorContent {
    None,
    Chart {
        rel_id: String,
        is_chart_ex: bool,
        raw_mc_fallback: Option<Vec<u8>>,
        name: Option<String>,
        descr: Option<String>,
        title: Option<String>,
        hidden: bool,
    },
    Pic(PicShape),
    Shape(ParsedShape),
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
            name,
            descr,
            title,
            hidden,
        } => DrawingEntryKind::Chart(DrawingChartRef {
            rel_id,
            anchor: anchor.clone(),
            is_chart_ex,
            raw_mc_fallback,
            name,
            descr,
            title,
            hidden,
        }),
        AnchorContent::Pic(pic) if pic.blip_rel.is_some() => DrawingEntryKind::Image(Box::new(pic)),
        AnchorContent::Twin(twin) => DrawingEntryKind::ControlTwin(twin),
        AnchorContent::Shape(shape) => DrawingEntryKind::Shape(Box::new(shape)),
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
) -> ChartParseResult<AnchorContent> {
    let name = start.local_name().as_ref().to_vec();
    match name.as_slice() {
        b"graphicFrame" => Ok(match parse_graphic_frame(reader)? {
            FrameContent::Chart {
                rel_id,
                is_chart_ex,
                name,
                descr,
                title,
                hidden,
            } => AnchorContent::Chart {
                rel_id,
                is_chart_ex,
                raw_mc_fallback: None,
                name,
                descr,
                title,
                hidden,
            },
            FrameContent::Twin(twin) => AnchorContent::Twin(twin),
            FrameContent::None => AnchorContent::Other,
        }),
        b"pic" => {
            let pic = parse_pic(reader)?;
            if pic.blip_rel.is_some() {
                Ok(AnchorContent::Pic(pic))
            } else {
                Ok(AnchorContent::Other)
            }
        }
        b"sp" => Ok(match parse_sp(reader, start)? {
            ParsedSp::Twin(twin) => AnchorContent::Twin(twin),
            ParsedSp::Shape(shape) => AnchorContent::Shape(shape),
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
/// chartEx graphicFrame or a control-twin shape (either flavor). The
/// first Choice with recognized content wins; the Fallback content is
/// preserved only for chartEx round-trip.
fn parse_anchor_alternate<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> ChartParseResult<AnchorContent> {
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
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    if let AnchorContent::Chart {
        rel_id,
        is_chart_ex,
        name,
        descr,
        title,
        hidden,
        ..
    } = chosen
    {
        return Ok(AnchorContent::Chart {
            rel_id,
            is_chart_ex,
            raw_mc_fallback: fallback,
            name,
            descr,
            title,
            hidden,
        });
    }
    match chosen {
        AnchorContent::None => Ok(AnchorContent::Other),
        other => Ok(other),
    }
}

/// Parse the inner content of an in-anchor `mc:Choice`.
fn parse_choice_content(bytes: &[u8]) -> ChartParseResult<AnchorContent> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let e = e.to_owned();
                return parse_anchor_content(&mut reader, &e);
            }
            Ok(Event::Eof) => return Ok(AnchorContent::None),
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

/// Content of an `<xdr:graphicFrame>`.
enum FrameContent {
    Chart {
        rel_id: String,
        is_chart_ex: bool,
        name: Option<String>,
        descr: Option<String>,
        title: Option<String>,
        hidden: bool,
    },
    Twin(TwinShape),
    None,
}

/// Parse `<xdr:graphicFrame>` content (start consumed): a chart
/// reference (`c:chart`/`cx:chart` inside graphicData) or an XLSB
/// control twin (`com14:compatSp`).
fn parse_graphic_frame<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> ChartParseResult<FrameContent> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut graphic_data_uri: Option<String> = None;
    let mut found: Option<(String, bool)> = None;
    let mut twin = TwinShape::default();
    let mut has_spid = false;
    let mut hidden = false;
    let mut in_xfrm = false;

    let handle = |e: &BytesStart<'_>,
                  twin: &mut TwinShape,
                  has_spid: &mut bool,
                  hidden: &mut bool,
                  found: &mut Option<(String, bool)>,
                  graphic_data_uri: &Option<String>,
                  in_xfrm: bool| {
        match e.local_name().as_ref() {
            b"chart" => {
                let is_chart_ex = graphic_data_uri.as_deref() == Some(URI_CHART_EX);
                if let Some(id) = attr_string(e, b"id") {
                    *found = Some((id, is_chart_ex));
                }
            }
            b"compatSp" => {
                if let Some(spid) = attr_string(e, b"spid") {
                    twin.shape_num = spid.rsplit(['s', 'S']).next().and_then(|n| n.parse().ok());
                    twin.spid = spid;
                    *has_spid = true;
                }
            }
            b"cNvPr" => {
                if let Some(name) = attr_string(e, b"name") {
                    twin.name = Some(name);
                }
                twin.descr = attr_string(e, b"descr");
                twin.title = attr_string(e, b"title");
                *hidden = attr_bool_default_false(e, b"hidden");
            }
            b"off" if in_xfrm => {
                twin.xfrm.x_emu = attr_i64(e, b"x").unwrap_or(0);
                twin.xfrm.y_emu = attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if in_xfrm => {
                twin.xfrm.cx_emu = attr_i64(e, b"cx").unwrap_or(0);
                twin.xfrm.cy_emu = attr_i64(e, b"cy").unwrap_or(0);
            }
            _ => {}
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                match e.local_name().as_ref() {
                    b"graphicData" => graphic_data_uri = attr_string(e, b"uri"),
                    b"xfrm" => in_xfrm = true,
                    _ => {}
                }
                handle(
                    e,
                    &mut twin,
                    &mut has_spid,
                    &mut hidden,
                    &mut found,
                    &graphic_data_uri,
                    in_xfrm,
                );
            }
            Ok(Event::Empty(ref e)) => {
                handle(
                    e,
                    &mut twin,
                    &mut has_spid,
                    &mut hidden,
                    &mut found,
                    &graphic_data_uri,
                    in_xfrm,
                );
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"xfrm" {
                    in_xfrm = false;
                }
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    if has_spid {
        return Ok(FrameContent::Twin(twin));
    }
    Ok(match found {
        Some((rel_id, is_chart_ex)) => FrameContent::Chart {
            rel_id,
            is_chart_ex,
            name: twin.name,
            descr: twin.descr,
            title: twin.title,
            hidden,
        },
        None => FrameContent::None,
    })
}

/// Parse `<xdr:pic>` content (start consumed).
fn parse_pic<R: std::io::BufRead>(reader: &mut Reader<R>) -> ChartParseResult<PicShape> {
    let mut buf = Vec::new();
    let mut depth: u32 = 1;
    let mut pic = PicShape::default();
    let mut in_sp_pr = false;
    let mut in_xfrm = false;

    let handle = |e: &BytesStart<'_>, pic: &mut PicShape, in_sp_pr: bool, in_xfrm: bool| match e
        .local_name()
        .as_ref()
    {
            b"cNvPr" => {
                if let Some(name) = attr_string(e, b"name") {
                    pic.name = name;
                }
                if let Some(descr) = attr_string(e, b"descr") {
                    pic.descr = Some(descr);
                }
            pic.title = attr_string(e, b"title");
                pic.hidden = attr_bool_default_false(e, b"hidden");
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
            pic.flip_h = matches!(
                attr_string(e, b"flipH").as_deref(),
                Some("1") | Some("true")
            );
            pic.flip_v = matches!(
                attr_string(e, b"flipV").as_deref(),
                Some("1") | Some("true")
            );
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
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
}

fn parse_twin_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> ChartParseResult<TwinText> {
    let mut text = TwinText::default();
    let mut buf = Vec::new();
    let mut depth = 1u32;
    let mut in_run = false;
    let mut in_text = false;
    let mut run_text = String::new();
    let mut run_font: Option<TwinRunFont> = None;

    let apply = |e: &BytesStart<'_>,
                 text: &mut TwinText,
                 run_font: &mut Option<TwinRunFont>,
                 in_run: bool| {
        match e.local_name().as_ref() {
            b"bodyPr" => {
                text.vertical_alignment =
                    attr_string(e, b"anchor").and_then(|value| match value.as_str() {
                        "t" => Some(TwinVerticalAlignment::Top),
                        "ctr" => Some(TwinVerticalAlignment::Center),
                        "b" => Some(TwinVerticalAlignment::Bottom),
                        "just" => Some(TwinVerticalAlignment::Justify),
                        "dist" => Some(TwinVerticalAlignment::Distributed),
                        _ => None,
                    });
            }
            b"pPr" => {
                text.horizontal_alignment =
                    attr_string(e, b"algn").and_then(|value| match value.as_str() {
                        "l" => Some(TwinHorizontalAlignment::Left),
                        "ctr" => Some(TwinHorizontalAlignment::Center),
                        "r" => Some(TwinHorizontalAlignment::Right),
                        "just" => Some(TwinHorizontalAlignment::Justify),
                        "dist" => Some(TwinHorizontalAlignment::Distributed),
                        _ => None,
                    });
            }
            b"rPr" if in_run => {
                let font = run_font.get_or_insert_with(TwinRunFont::default);
                font.size = attr_i64(e, b"sz").map(|size| size as f64 / 100.0);
                font.bold = attr_string(e, b"b")
                    .map(|value| value == "1" || value.eq_ignore_ascii_case("true"));
                font.italic = attr_string(e, b"i")
                    .map(|value| value == "1" || value.eq_ignore_ascii_case("true"));
                font.underline = attr_string(e, b"u").and_then(|value| match value.as_str() {
                    "sng" => Some(TwinUnderline::Single),
                    "dbl" => Some(TwinUnderline::Double),
                    _ => None,
                });
                font.strikethrough = attr_string(e, b"strike").map(|value| value != "noStrike");
                font.baseline = attr_i64(e, b"baseline").map(|value| value as i32);
            }
            b"latin" if in_run => {
                run_font.get_or_insert_with(TwinRunFont::default).name =
                    attr_string(e, b"typeface");
            }
            b"srgbClr" if in_run => {
                if let Some(value) = attr_string(e, b"val") {
                    let value = value.trim_start_matches('#');
                    if value.len() == 6 {
                        if let (Ok(r), Ok(g), Ok(b)) = (
                            u8::from_str_radix(&value[0..2], 16),
                            u8::from_str_radix(&value[2..4], 16),
                            u8::from_str_radix(&value[4..6], 16),
                        ) {
                            run_font.get_or_insert_with(TwinRunFont::default).color =
                                Some(TwinColor::Rgb { r, g, b });
                        }
                    }
                }
            }
            b"schemeClr" if in_run => {
                let index = attr_string(e, b"val").and_then(|value| match value.as_str() {
                    "lt1" => Some(0),
                    "dk1" => Some(1),
                    "lt2" => Some(2),
                    "dk2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    _ => None,
                });
                if let Some(index) = index {
                    run_font.get_or_insert_with(TwinRunFont::default).color =
                        Some(TwinColor::Theme { index, tint: 0 });
                }
            }
            b"lumMod" if in_run => {
                if let Some(value) = attr_i64(e, b"val") {
                    if let Some(TwinColor::Theme { index, .. }) =
                        run_font.as_ref().and_then(|font| font.color)
                    {
                        let tint = ((value - 100_000) / 1_000).clamp(-100, 0) as i8;
                        run_font.get_or_insert_with(TwinRunFont::default).color =
                            Some(TwinColor::Theme { index, tint });
                    }
                }
            }
            b"lumOff" if in_run => {
                if let Some(value) = attr_i64(e, b"val") {
                    if let Some(TwinColor::Theme { index, .. }) =
                        run_font.as_ref().and_then(|font| font.color)
                    {
                        let tint = (value / 1_000).clamp(0, 100) as i8;
                        run_font.get_or_insert_with(TwinRunFont::default).color =
                            Some(TwinColor::Theme { index, tint });
                    }
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
                    b"r" => {
                        in_run = true;
                        run_text.clear();
                        run_font = None;
                    }
                    b"t" if in_run => in_text = true,
                    b"br" if in_run => run_text.push('\n'),
                    _ => {}
                }
                apply(e, &mut text, &mut run_font, in_run);
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"br" && in_run {
                    run_text.push('\n');
                }
                apply(e, &mut text, &mut run_font, in_run);
            }
            Ok(Event::Text(ref value)) if in_text => {
                run_text.push_str(
                    &value
                        .unescape()
                        .map(|value| value.into_owned())
                        .unwrap_or_else(|_| String::from_utf8_lossy(value.as_ref()).into_owned()),
                );
            }
            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"t" => in_text = false,
                    b"r" if in_run => {
                        text.runs.push(TwinTextRun {
                            text: std::mem::take(&mut run_text),
                            font: run_font.take().filter(|font| {
                                font.name.is_some()
                                    || font.size.is_some()
                                    || font.color.is_some()
                                    || font.bold.is_some()
                                    || font.italic.is_some()
                                    || font.underline.is_some()
                                    || font.strikethrough.is_some()
                                    || font.baseline.is_some()
                            }),
                        });
                        in_run = false;
                    }
                    _ => {}
                }
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

enum ParsedSp {
    Twin(TwinShape),
    Shape(ParsedShape),
}

/// Parse one complete `<xdr:sp>`. Capturing first lets the shape
/// parser retain unmodeled direct child fragments while the caller's
/// stream advances exactly once.
fn parse_sp<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart<'_>,
) -> ChartParseResult<ParsedSp> {
    let bytes = capture_element(reader, start, &[])?;
    parse_sp_xml(&bytes)
}

fn parse_sp_xml(bytes: &[u8]) -> ChartParseResult<ParsedSp> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let root = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => break e.into_owned(),
            Ok(Event::Eof) => return Err(ChartParseError::Parse("empty xdr:sp".into())),
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => buf.clear(),
        }
    };

    let mut twin = TwinShape {
        macro_name: attr_string(&root, b"macro").filter(|value| !value.is_empty()),
        ..TwinShape::default()
    };
    let mut shape = ParsedShape {
        geometry: "rect".to_string(),
        ..ParsedShape::default()
    };
    let mut has_spid = false;
    let mut depth = 1u32;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if depth == 1 && e.local_name().as_ref() == b"spPr" => {
                let fragment = capture_element(&mut reader, e, &[])?;
                parse_shape_properties(&fragment, &mut shape)?;
            }
            Ok(Event::Start(ref e)) if depth == 1 && e.local_name().as_ref() == b"txBody" => {
                let fragment = capture_element(&mut reader, e, &[])?;
                let (text, raw) = parse_shape_text_body(&fragment)?;
                shape.text = Some(text.clone());
                shape.raw_text_body = raw;
                twin.text = Some(text);
            }
            Ok(Event::Start(ref e)) => {
                depth += 1;
                parse_sp_metadata(e, &mut twin, &mut shape, &mut has_spid);
            }
            Ok(Event::Empty(ref e)) => {
                parse_sp_metadata(e, &mut twin, &mut shape, &mut has_spid);
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
    }

    if has_spid {
        twin.xfrm = shape.xfrm;
        Ok(ParsedSp::Twin(twin))
    } else {
        Ok(ParsedSp::Shape(shape))
    }
}

fn parse_sp_metadata(
    e: &BytesStart<'_>,
    twin: &mut TwinShape,
    shape: &mut ParsedShape,
    has_spid: &mut bool,
) {
    match e.local_name().as_ref() {
        b"cNvPr" => {
            if let Some(name) = attr_string(e, b"name") {
                twin.name = Some(name.clone());
                shape.name = name;
            }
            twin.descr = attr_string(e, b"descr");
            twin.title = attr_string(e, b"title");
            shape.descr = twin.descr.clone();
            shape.title = twin.title.clone();
            shape.hidden = attr_bool_default_false(e, b"hidden");
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
}

fn parse_shape_properties(bytes: &[u8], shape: &mut ParsedShape) -> ChartParseResult<()> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    // Consume spPr itself.
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => break,
            Ok(Event::Eof) => return Ok(()),
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => buf.clear(),
        }
    }

    let mut raw = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"xfrm" => {
                    let fragment = capture_element(&mut reader, e, &[])?;
                    shape.xfrm = parse_shape_xfrm(&fragment)?;
                }
                b"prstGeom" => {
                    shape.geometry = attr_string(e, b"prst").unwrap_or_else(|| "rect".into());
                    raw.extend_from_slice(&capture_element(&mut reader, e, &[])?);
                }
                b"solidFill" => {
                    let fragment = capture_element(&mut reader, e, &[])?;
                    if let Some(color) = parse_drawing_color(&fragment) {
                        shape.fill = ShapeFill::Solid(color);
                    }
                    raw.extend_from_slice(&fragment);
                }
                b"noFill" => {
                    shape.fill = ShapeFill::None;
                    skip_element(&mut reader)?;
                }
                b"ln" => {
                    let fragment = capture_element(&mut reader, e, &[])?;
                    shape.line = parse_shape_line(&fragment)?;
                    raw.extend_from_slice(&fragment);
                }
                _ => raw.extend_from_slice(&capture_element(&mut reader, e, &[])?),
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"xfrm" => shape.xfrm = parse_xfrm_attrs(e),
                b"prstGeom" => {
                    shape.geometry = attr_string(e, b"prst").unwrap_or_else(|| "rect".into());
                    write_empty_fragment(&mut raw, e)?;
                }
                b"noFill" => shape.fill = ShapeFill::None,
                b"solidFill" => write_empty_fragment(&mut raw, e)?,
                b"ln" => {
                    shape.line.width_emu = attr_i64(e, b"w");
                    write_empty_fragment(&mut raw, e)?;
                }
                _ => write_empty_fragment(&mut raw, e)?,
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"spPr" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
    }
    shape.raw_shape_properties = (!raw.is_empty()).then_some(raw);
    Ok(())
}

fn parse_xfrm_attrs(e: &BytesStart<'_>) -> ChildTransform {
    ChildTransform {
        rotation: attr_i64(e, b"rot").unwrap_or(0) as i32,
        flip_h: attr_bool_default_false(e, b"flipH"),
        flip_v: attr_bool_default_false(e, b"flipV"),
        ..ChildTransform::default()
    }
}

fn parse_shape_xfrm(bytes: &[u8]) -> ChartParseResult<ChildTransform> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let mut transform = ChildTransform::default();
    let mut in_root = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"xfrm" => {
                    transform = parse_xfrm_attrs(e);
                    in_root = true;
                }
                b"off" if in_root => {
                    transform.x_emu = attr_i64(e, b"x").unwrap_or(0);
                    transform.y_emu = attr_i64(e, b"y").unwrap_or(0);
                }
                b"ext" if in_root => {
                    transform.cx_emu = attr_i64(e, b"cx").unwrap_or(0);
                    transform.cy_emu = attr_i64(e, b"cy").unwrap_or(0);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(transform)
}

fn parse_shape_line(bytes: &[u8]) -> ChartParseResult<ShapeLine> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let mut line = ShapeLine::default();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"ln" => line.width_emu = attr_i64(e, b"w"),
                b"solidFill" => {
                    let fragment = capture_element(&mut reader, e, &[])?;
                    line.color = parse_drawing_color(&fragment);
                }
                b"noFill" => {
                    line.no_fill = true;
                    skip_element(&mut reader)?;
                }
                b"prstDash" => {
                    line.dash_style = attr_string(e, b"val");
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"ln" => line.width_emu = attr_i64(e, b"w"),
                b"noFill" => line.no_fill = true,
                b"prstDash" => line.dash_style = attr_string(e, b"val"),
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(line)
}

fn parse_drawing_color(bytes: &[u8]) -> Option<TwinColor> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let mut color = None;
    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"srgbClr" => {
                    let value = attr_string(e, b"val")?;
                    let value = value.trim_start_matches('#');
                    if value.len() == 6 {
                        color = Some(TwinColor::Rgb {
                            r: u8::from_str_radix(&value[0..2], 16).ok()?,
                            g: u8::from_str_radix(&value[2..4], 16).ok()?,
                            b: u8::from_str_radix(&value[4..6], 16).ok()?,
                        });
                    }
                }
                b"schemeClr" => {
                    let index = scheme_color_index(&attr_string(e, b"val")?)?;
                    color = Some(TwinColor::Theme { index, tint: 0 });
                }
                b"lumMod" => {
                    if let (Some(value), Some(TwinColor::Theme { index, .. })) =
                        (attr_i64(e, b"val"), color)
                    {
                        color = Some(TwinColor::Theme {
                            index,
                            tint: ((value - 100_000) / 1_000).clamp(-100, 0) as i8,
                        });
                    }
                }
                b"lumOff" => {
                    if let (Some(value), Some(TwinColor::Theme { index, .. })) =
                        (attr_i64(e, b"val"), color)
                    {
                        color = Some(TwinColor::Theme {
                            index,
                            tint: (value / 1_000).clamp(0, 100) as i8,
                        });
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    color
}

fn scheme_color_index(value: &str) -> Option<u8> {
    match value {
        "lt1" => Some(0),
        "dk1" => Some(1),
        "lt2" => Some(2),
        "dk2" => Some(3),
        "accent1" => Some(4),
        "accent2" => Some(5),
        "accent3" => Some(6),
        "accent4" => Some(7),
        "accent5" => Some(8),
        "accent6" => Some(9),
        _ => None,
    }
}

fn parse_shape_text_body(bytes: &[u8]) -> ChartParseResult<(TwinText, Option<Vec<u8>>)> {
    let mut reader = new_reader(bytes);
    let mut buf = Vec::new();
    let text = loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txBody" => {
                break parse_twin_text(&mut reader)?;
            }
            Ok(Event::Eof) => break TwinText::default(),
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => buf.clear(),
        }
    };
    Ok((text, Some(bytes.to_vec())))
}

fn write_empty_fragment(out: &mut Vec<u8>, e: &BytesStart<'_>) -> ChartParseResult<()> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer
        .write_event(Event::Empty(e.to_owned()))
        .map_err(std::io::Error::other)?;
    out.extend_from_slice(&writer.into_inner().into_inner());
    Ok(())
}

/// Parse `<xdr:grpSp>` content (start consumed). Returns `None` when
/// any child is not modelable (the caller keeps the anchor raw).
fn parse_group<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> ChartParseResult<Option<ParsedGroup>> {
    let mut buf = Vec::new();
    let mut group = ParsedGroup::default();
    let mut modelable = true;
    let mut in_nv = false;
    let mut in_grp_sp_pr = false;
    let mut in_xfrm = false;

    let handle_xfrm_child =
        |e: &BytesStart<'_>, transform: &mut GroupTransform| match e.local_name().as_ref() {
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
                        pic if pic.blip_rel.is_some() => group.children.push(ParsedChild::Pic(pic)),
                        _ => modelable = false,
                    },
                    b"sp" if !in_nv && !in_grp_sp_pr => match parse_sp(reader, e)? {
                        ParsedSp::Twin(twin) => group.children.push(ParsedChild::Twin(twin)),
                        ParsedSp::Shape(shape) => group.children.push(ParsedChild::Shape(shape)),
                    },
                    b"graphicFrame" if !in_nv && !in_grp_sp_pr => {
                        match parse_graphic_frame(reader)? {
                            FrameContent::Twin(twin) => {
                                group.children.push(ParsedChild::Twin(twin))
                            }
                            _ => modelable = false,
                        }
                    }
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
                            AnchorContent::Shape(shape) => {
                                group.children.push(ParsedChild::Shape(shape))
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
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                    b"cNvPr" if in_nv => parse_group_cnvpr(e, &mut group),
                b"xfrm" if in_grp_sp_pr => parse_group_xfrm_attrs(e, &mut group.transform),
                    b"off" | b"ext" | b"chOff" | b"chExt" if in_xfrm => {
                        handle_xfrm_child(e, &mut group.transform);
                    }
                    b"pic" | b"sp" | b"grpSp" | b"cxnSp" | b"graphicFrame" | b"contentPart"
                        if !in_nv && !in_grp_sp_pr =>
                    {
                        modelable = false;
                    }
                    _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"nvGrpSpPr" => in_nv = false,
                b"grpSpPr" => in_grp_sp_pr = false,
                b"xfrm" => in_xfrm = false,
                b"grpSp" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
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
    group.title = attr_string(e, b"title");
    group.hidden = attr_bool_default_false(e, b"hidden");
}

fn parse_group_xfrm_attrs(e: &BytesStart<'_>, transform: &mut GroupTransform) {
    transform.rotation = attr_i64(e, b"rot").map(|v| v as i32).unwrap_or(0);
    transform.flip_h = matches!(
        attr_string(e, b"flipH").as_deref(),
        Some("1") | Some("true")
    );
    transform.flip_v = matches!(
        attr_string(e, b"flipV").as_deref(),
        Some("1") | Some("true")
    );
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

#[cfg(test)]
mod tests {
    use super::*;

    /// The XLSB control-twin flavor (com14:compatSp graphicFrame)
    /// parses as a twin, both bare and inside AlternateContent.
    #[test]
    fn parse_compat_sp_frame_twin() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1025" name="Check Box 1"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/compatibility"><com14:compatSp xmlns:com14="http://schemas.microsoft.com/office/drawing/2010/compatibility" spid="_x0000_s1025"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::ControlTwin(twin) = &entries[0].kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.spid, "_x0000_s1025");
        assert_eq!(twin.shape_num, Some(1025));
        assert_eq!(twin.name.as_deref(), Some("Check Box 1"));
    }

    /// A plain chart graphicFrame still parses as a chart ref.
    #[test]
    fn parse_chart_frame_still_chart() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Chart(chart_ref) = &entries[0].kind else {
            panic!("expected chart entry");
        };
        assert_eq!(chart_ref.rel_id, "rId1");
        assert!(!chart_ref.is_chart_ex);
    }

    /// The XLSX sp-flavor twin also carries the cNvPr name.
    #[test]
    fn parse_compat_ext_sp_twin_name() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Check Box 7" hidden="1"><a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1031"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::ControlTwin(twin) = &entries[0].kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.shape_num, Some(1031));
        assert_eq!(twin.name.as_deref(), Some("Check Box 7"));
    }

    /// An anchor captured from inside a wsDr-level mc:Choice must
    /// inherit the xmlns declarations carried by the
    /// mc:AlternateContent / mc:Choice wrappers, or the captured raw
    /// bytes reference undeclared prefixes and are not well formed
    /// when re-spliced on write.
    #[test]
    fn wsdr_alternate_choice_ns_decls_are_injected_into_captured_anchor() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"><mc:Choice xmlns:cx9="http://schemas.microsoft.com/office/drawing/2016/9/9/chartex" Requires="am3d"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="3D Model 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:mdl3d><am3d:spPr/></am3d:mdl3d></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        assert!(matches!(entries[0].kind, DrawingEntryKind::Raw));
        let bytes = std::str::from_utf8(&entries[0].bytes).unwrap();
        assert!(
            bytes.starts_with("<xdr:twoCellAnchor"),
            "anchor root captured: {bytes}"
        );
        assert!(
            bytes.contains(
                r#"xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d""#
            ),
            "AlternateContent xmlns:am3d injected into the captured root: {bytes}"
        );
        assert!(
            bytes.contains(
                r#"xmlns:cx9="http://schemas.microsoft.com/office/drawing/2016/9/9/chartex""#
            ),
            "Choice xmlns:cx9 injected into the captured root: {bytes}"
        );
        // The anchor's payload is intact.
        assert!(bytes.contains("<am3d:mdl3d>"), "{bytes}");
    }

    /// A root element that already declares a prefix must not have it
    /// re-injected (the duplicate attribute would be malformed).
    #[test]
    fn wsdr_alternate_ns_injection_skips_already_declared_prefixes() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:am3d="urn:outer"><mc:Choice Requires="am3d"><xdr:twoCellAnchor xmlns:am3d="urn:inner"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="c"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><am3d:x/></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        let bytes = std::str::from_utf8(&entries[0].bytes).unwrap();
        assert_eq!(
            bytes.matches("xmlns:am3d").count(),
            1,
            "root's own declaration wins, no duplicate: {bytes}"
        );
        assert!(bytes.contains(r#"xmlns:am3d="urn:inner""#), "{bytes}");
    }

    /// compatSp twins inside a group parse as twin children.
    #[test]
    fn parse_group_with_compat_sp_frame_child() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></xdr:grpSpPr><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1027" name="Check Box 3"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/compatibility"><com14:compatSp xmlns:com14="http://schemas.microsoft.com/office/drawing/2010/compatibility" spid="_x0000_s1027"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let entries = parse_drawing_part(xml.as_bytes()).unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Group(group) = &entries[0].kind else {
            panic!("expected group entry");
        };
        assert_eq!(group.children.len(), 1);
        let ParsedChild::Twin(twin) = &group.children[0] else {
            panic!("expected twin child");
        };
        assert_eq!(twin.shape_num, Some(1027));
        assert_eq!(twin.xfrm.x_emu, 10);
        assert_eq!(twin.xfrm.cy_emu, 40);
    }
}
