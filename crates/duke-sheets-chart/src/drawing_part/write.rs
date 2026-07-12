//! Drawing-part writer: emit a `wsDr` part (and its rels) from a
//! caller-built object list in z-order.
//!
//! The caller resolves everything model-specific up front (control
//! display names, unsupported-chart filtering, comment skipping);
//! this module owns the XML shape, relationship-id planning, and the
//! control-twin flavor.

use std::collections::HashSet;
use std::io::{Cursor, Write};

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use crate::{
    marker_position_emu, CellMarker, ChildTransform, DefaultDrawingMetrics, DrawingAnchor,
    DrawingMetrics, EditAs, EmbeddedImage, GroupTransform,
};

use super::{
    ShapeFill, ShapeLine, TwinColor, TwinHorizontalAlignment, TwinText, TwinUnderline,
    TwinVerticalAlignment,
};
use super::{RT_CHART, RT_CHART_EX, RT_IMAGE};

const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const NS_MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CX1: &str = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const NS_A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const NS_COM14: &str = "http://schemas.microsoft.com/office/drawing/2010/compatibility";
const NS_DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

type XmlWriter = Writer<Cursor<Vec<u8>>>;
type WriteResult<T> = std::io::Result<T>;

/// Control-twin markup flavor. XLSX emits `a14:compatExt` inside an
/// `<xdr:sp>`; XLSB emits `com14:compatSp` inside an
/// `<xdr:graphicFrame>`. The reader accepts both regardless.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TwinStyle {
    /// `<xdr:sp>` with `a14:compatExt` (XLSX).
    CompatExtSp,
    /// `<xdr:graphicFrame>` with `com14:compatSp` (XLSB).
    CompatSpFrame,
}

/// One drawing-list object prepared for emission.
pub struct PartObject<'a> {
    pub name: Option<&'a str>,
    pub alt_text: Option<&'a str>,
    pub title: Option<&'a str>,
    pub macro_name: Option<String>,
    pub control_text: Option<TwinText>,
    pub locked: bool,
    pub printable: bool,
    /// Emits `cNvPr/@hidden="1"` on the native shape (images, charts,
    /// groups). Control twins ignore it: their markup is structural.
    pub hidden: bool,
    pub anchor: &'a DrawingAnchor,
    pub kind: PartKind<'a>,
}

/// One group child prepared for emission.
pub struct PartChild<'a> {
    pub name: Option<&'a str>,
    pub alt_text: Option<&'a str>,
    pub title: Option<&'a str>,
    pub macro_name: Option<String>,
    pub control_text: Option<TwinText>,
    pub hidden: bool,
    pub transform: &'a ChildTransform,
    pub kind: PartKind<'a>,
}

/// One modeled basic shape prepared by an XLSX/XLSB adapter.
pub struct PartShape<'a> {
    pub geometry: &'a str,
    pub fill: ShapeFill,
    pub line: ShapeLine,
    pub text: Option<TwinText>,
    pub rotation: i32,
    pub flip_h: bool,
    pub flip_v: bool,
    pub raw_shape_properties: Option<&'a [u8]>,
    pub raw_text_body: Option<&'a [u8]>,
    pub raw_geometry_unchanged: bool,
    pub raw_fill_unchanged: bool,
    pub raw_line_unchanged: bool,
    pub raw_text_unchanged: bool,
}

/// Kind payload of a [`PartObject`] / [`PartChild`]. Comments and
/// unsupported charts have no drawing-part presence; the caller omits
/// them.
pub enum PartKind<'a> {
    Image(&'a EmbeddedImage),
    Shape(PartShape<'a>),
    Chart,
    ChartEx {
        fallback: Option<&'a [u8]>,
    },
    /// A form control's placeholder twin; `name` is the resolved
    /// display name.
    Control {
        name: String,
    },
    Group {
        transform: &'a GroupTransform,
        children: Vec<PartChild<'a>>,
    },
    Raw {
        bytes: &'a [u8],
        rels: Vec<PartRel<'a>>,
    },
}

/// A relationship preserved on a raw entry.
pub struct PartRel<'a> {
    pub id: &'a str,
    pub rel_type: &'a str,
    pub target: &'a str,
    pub external: bool,
}

/// One relationship of a sheet's drawing part.
pub struct PlannedRel {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub external: bool,
}

/// Relationship-id assignments for a sheet's drawing part. Writer
/// generated ids never collide with ids preserved from raw entries.
pub struct DrawingPlan {
    /// rIds for supported charts, in drawing-list order.
    pub chart_rids: Vec<String>,
    /// rIds for chartEx charts, in drawing-list order.
    pub chartex_rids: Vec<String>,
    /// rIds for images (including group children), depth-first in
    /// drawing-list order.
    pub image_rids: Vec<String>,
    /// All relationships to write in the drawing's .rels part.
    pub rels: Vec<PlannedRel>,
}

/// Assign relationship ids for the sheet's drawing part. Raw entries
/// keep their original ids; generated ids skip over them.
pub fn plan_drawing_rels(
    objects: &[PartObject<'_>],
    chart_globals: &[usize],
    chartex_globals: &[usize],
    image_parts: &[(usize, &'static str)],
) -> DrawingPlan {
    let mut taken_ids: HashSet<String> = HashSet::new();
    let mut taken_nums: HashSet<usize> = HashSet::new();
    for object in objects {
        if let PartKind::Raw { rels, .. } = &object.kind {
            for rel in rels {
                if taken_ids.insert(rel.id.to_string()) {
                    if let Some(num) = rel.id.strip_prefix("rId").and_then(|n| n.parse().ok()) {
                        taken_nums.insert(num);
                    }
                }
            }
        }
    }

    let mut next = 1usize;
    let mut alloc = move || {
        while taken_nums.contains(&next) {
            next += 1;
        }
        taken_nums.insert(next);
        format!("rId{next}")
    };

    let mut plan = DrawingPlan {
        chart_rids: Vec::new(),
        chartex_rids: Vec::new(),
        image_rids: Vec::new(),
        rels: Vec::new(),
    };
    let mut chart_i = 0usize;
    let mut chartex_i = 0usize;
    let mut image_i = 0usize;
    let mut seen_raw: HashSet<String> = HashSet::new();

    fn plan_image(
        plan: &mut DrawingPlan,
        alloc: &mut impl FnMut() -> String,
        image_parts: &[(usize, &'static str)],
        image_i: &mut usize,
    ) {
        let rid = alloc();
        let (global_num, ext) = image_parts[*image_i];
        *image_i += 1;
        plan.rels.push(PlannedRel {
            id: rid.clone(),
            rel_type: RT_IMAGE.to_string(),
            target: format!("../media/image{global_num}.{ext}"),
            external: false,
        });
        plan.image_rids.push(rid);
    }

    fn plan_children(
        children: &[PartChild<'_>],
        plan: &mut DrawingPlan,
        alloc: &mut impl FnMut() -> String,
        image_parts: &[(usize, &'static str)],
        image_i: &mut usize,
    ) {
        for child in children {
            match &child.kind {
                PartKind::Image(_) => plan_image(plan, alloc, image_parts, image_i),
                PartKind::Group { children, .. } => {
                    plan_children(children, plan, alloc, image_parts, image_i)
                }
                _ => {}
            }
        }
    }

    for object in objects {
        match &object.kind {
            PartKind::Chart => {
                let rid = alloc();
                plan.rels.push(PlannedRel {
                    id: rid.clone(),
                    rel_type: RT_CHART.to_string(),
                    target: format!("../charts/chart{}.xml", chart_globals[chart_i]),
                    external: false,
                });
                chart_i += 1;
                plan.chart_rids.push(rid);
            }
            PartKind::ChartEx { .. } => {
                let rid = alloc();
                plan.rels.push(PlannedRel {
                    id: rid.clone(),
                    rel_type: RT_CHART_EX.to_string(),
                    target: format!("../charts/chartEx{}.xml", chartex_globals[chartex_i]),
                    external: false,
                });
                chartex_i += 1;
                plan.chartex_rids.push(rid);
            }
            PartKind::Image(_) => plan_image(&mut plan, &mut alloc, image_parts, &mut image_i),
            PartKind::Group { children, .. } => {
                plan_children(children, &mut plan, &mut alloc, image_parts, &mut image_i)
            }
            PartKind::Raw { rels, .. } => {
                for rel in rels {
                    if seen_raw.insert(rel.id.to_string()) {
                        plan.rels.push(PlannedRel {
                            id: rel.id.to_string(),
                            rel_type: rel.rel_type.to_string(),
                            target: rel.target.to_string(),
                            external: rel.external,
                        });
                    }
                }
            }
            PartKind::Control { .. } | PartKind::Shape(_) => {}
        }
    }
    plan
}

/// Walk state for emitting a drawing part in drawing-list order.
struct EmitCtx<'a> {
    plan: &'a DrawingPlan,
    twin_style: TwinStyle,
    chart_i: usize,
    chartex_i: usize,
    image_i: usize,
    /// cNvPr id space, unique across the part, sequential in emission
    /// order (base 2 by convention).
    cnv_id: usize,
    /// 1-based sequence of chart graphic frames, for default names.
    frame_seq: usize,
    /// Placed-control ordinal (drives twin shape ids).
    control_ordinal: usize,
    shape_base: usize,
}

impl EmitCtx<'_> {
    fn next_cnv_id(&mut self) -> usize {
        let id = self.cnv_id;
        self.cnv_id += 1;
        id
    }
}

/// Serialize one sheet's drawing part: every object in list order.
/// Form controls emit their placeholder twins in the chosen flavor;
/// raw entries pass through verbatim. `shape_base` is the numeric
/// spid of the first control twin (`_x0000_s{shape_base}`).
pub fn write_drawing_part(
    objects: &[PartObject<'_>],
    plan: &DrawingPlan,
    twin_style: TwinStyle,
    shape_base: usize,
) -> WriteResult<Vec<u8>> {
    write_drawing_part_with_metrics(
        objects,
        plan,
        twin_style,
        shape_base,
        &DefaultDrawingMetrics,
    )
}

/// Serialize one sheet's drawing part with worksheet metrics available
/// for control twins and other marker conversions. Native OneCell and
/// Absolute wrappers remain in their original exact forms.
pub fn write_drawing_part_with_metrics(
    objects: &[PartObject<'_>],
    plan: &DrawingPlan,
    twin_style: TwinStyle,
    shape_base: usize,
    metrics: &dyn DrawingMetrics,
) -> WriteResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut tag = BytesStart::new("xdr:wsDr");
    tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
    tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
    tag.push_attribute(("xmlns:r", NS_DOC_RELS));
    tag.push_attribute(("xmlns:mc", NS_MC));
    tag.push_attribute(("xmlns:a14", NS_A14));
    w.write_event(Event::Start(tag))?;

    let mut ctx = EmitCtx {
        plan,
        twin_style,
        chart_i: 0,
        chartex_i: 0,
        image_i: 0,
        cnv_id: 2,
        frame_seq: 0,
        control_ordinal: 0,
        shape_base,
    };

    for object in objects {
        match &object.kind {
            PartKind::Chart => {
                ctx.frame_seq += 1;
                let rid = plan.chart_rids[ctx.chart_i].clone();
                ctx.chart_i += 1;
                let cnv_id = ctx.next_cnv_id();
                let default_name = format!("Chart {}", ctx.frame_seq);
                let name = object.name.unwrap_or(&default_name);
                write_anchor_wrapper(&mut w, object, |w| {
                    write_chart_frame(
                        w,
                        &rid,
                        cnv_id,
                        name,
                        object.alt_text,
                        object.title,
                        object.hidden,
                    )
                })?;
            }
            PartKind::ChartEx { fallback } => {
                ctx.frame_seq += 1;
                let rid = plan.chartex_rids[ctx.chartex_i].clone();
                ctx.chartex_i += 1;
                let cnv_id = ctx.next_cnv_id();
                let default_name = format!("Chart {}", ctx.frame_seq);
                let name = object.name.unwrap_or(&default_name);
                write_anchor_wrapper(&mut w, object, |w| {
                    write_chartex_frame(
                        w,
                        &rid,
                        cnv_id,
                        name,
                        object.alt_text,
                        object.title,
                        object.hidden,
                        *fallback,
                    )
                })?;
            }
            PartKind::Image(image) => {
                let rid = plan.image_rids[ctx.image_i].clone();
                ctx.image_i += 1;
                let cnv_id = ctx.next_cnv_id();
                let xfrm = PicXfrm {
                    off_x: 0,
                    off_y: 0,
                    cx: image.width_emu,
                    cy: image.height_emu,
                    rot: image.rotation,
                    flip_h: image.flip_h,
                    flip_v: image.flip_v,
                };
                write_anchor_wrapper(&mut w, object, |w| {
                    write_picture_element(
                        w,
                        object.name,
                        object.alt_text,
                        object.title,
                        object.hidden,
                        &rid,
                        cnv_id,
                        &xfrm,
                    )
                })?;
            }
            PartKind::Shape(shape) => {
                let cnv_id = ctx.next_cnv_id();
                let placement = placement_from_anchor(object.anchor, shape, metrics);
                write_anchor_wrapper(&mut w, object, |w| {
                    write_shape_element(
                        w,
                        object.name,
                        object.alt_text,
                        object.title,
                        object.hidden,
                        cnv_id,
                        shape,
                        &placement,
                    )
                })?;
            }
            PartKind::Group {
                transform,
                children,
            } => {
                let placement = Placement {
                    x: transform.x_emu,
                    y: transform.y_emu,
                    cx: transform.cx_emu,
                    cy: transform.cy_emu,
                    rot: transform.rotation,
                    flip_h: transform.flip_h,
                    flip_v: transform.flip_v,
                };
                write_anchor_wrapper(&mut w, object, |w| {
                    write_group_shape(
                        w,
                        object.name,
                        object.alt_text,
                        object.title,
                        object.hidden,
                        transform,
                        children,
                        &placement,
                        &mut ctx,
                    )
                })?;
            }
            PartKind::Control { name } => {
                let shape_id = ctx.shape_base + ctx.control_ordinal;
                ctx.control_ordinal += 1;
                let cnv_id = ctx.next_cnv_id();
                write_control_twin_anchor(
                    &mut w,
                    object,
                    name,
                    shape_id,
                    cnv_id,
                    twin_style,
                    metrics,
                )?;
            }
            PartKind::Raw { bytes, .. } => {
                w.get_mut().write_all(bytes)?;
            }
        }
    }

    w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
    Ok(w.into_inner().into_inner())
}

/// Group placement in its parent space (anchor space for top-level
/// groups, the parent group's child space for nested ones).
struct Placement {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    rot: i32,
    flip_h: bool,
    flip_v: bool,
}

fn placement_from_anchor(
    anchor: &DrawingAnchor,
    shape: &PartShape<'_>,
    metrics: &dyn DrawingMetrics,
) -> Placement {
    let marker_xy = |marker: &CellMarker| {
        let (x, y) = marker_position_emu(marker, metrics);
        (
            x.clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64,
            y.clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64,
        )
    };
    let (x, y, cx, cy) = match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            let (x1, y1) = marker_xy(from);
            let (x2, y2) = marker_xy(to);
            (
                x1,
                y1,
                x2.saturating_sub(x1).max(0),
                y2.saturating_sub(y1).max(0),
            )
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let (x, y) = marker_xy(from);
            (x, y, (*width_emu).max(0), (*height_emu).max(0))
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => (*x_emu, *y_emu, (*width_emu).max(0), (*height_emu).max(0)),
    };
    Placement {
        x,
        y,
        cx,
        cy,
        rot: shape.rotation,
        flip_h: shape.flip_h,
        flip_v: shape.flip_v,
    }
}

#[allow(clippy::too_many_arguments)]
fn write_shape_element(
    w: &mut XmlWriter,
    name: Option<&str>,
    alt_text: Option<&str>,
    title: Option<&str>,
    hidden: bool,
    cnv_id: usize,
    shape: &PartShape<'_>,
    placement: &Placement,
) -> WriteResult<()> {
    let mut sp = BytesStart::new("xdr:sp");
    sp.push_attribute(("macro", ""));
    sp.push_attribute(("textlink", ""));
    w.write_event(Event::Start(sp))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvSpPr")))?;
    let id = cnv_id.to_string();
    let default_name = format!("Shape {}", cnv_id.saturating_sub(1));
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", id.as_str()));
    cnv_pr.push_attribute(("name", name.unwrap_or(&default_name)));
    if let Some(alt_text) = alt_text {
        cnv_pr.push_attribute(("descr", alt_text));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    if hidden {
        cnv_pr.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvSpPr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvSpPr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
    let mut xfrm = BytesStart::new("a:xfrm");
    if placement.rot != 0 {
        xfrm.push_attribute(("rot", placement.rot.to_string().as_str()));
    }
    if placement.flip_h {
        xfrm.push_attribute(("flipH", "1"));
    }
    if placement.flip_v {
        xfrm.push_attribute(("flipV", "1"));
    }
    w.write_event(Event::Start(xfrm))?;
    write_point(w, "a:off", placement.x, placement.y)?;
    write_extent(w, "a:ext", placement.cx, placement.cy)?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;

    let raw = shape.raw_shape_properties.unwrap_or_default();
    const RAW_GEOMETRY: &[&str] = &["prstGeom", "custGeom"];
    let raw_geometry = shape.raw_geometry_unchanged && raw_has_any(raw, RAW_GEOMETRY);
    if raw_geometry {
        write_raw_shape_fragments(w, raw, RAW_GEOMETRY)?;
    } else {
        let mut geometry = BytesStart::new("a:prstGeom");
        geometry.push_attribute(("prst", shape.geometry));
        w.write_event(Event::Start(geometry))?;
        w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    }

    const RAW_FILLS: &[&str] = &["solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
    let raw_fill = shape.raw_fill_unchanged && raw_has_any(raw, RAW_FILLS);
    if raw_fill {
        write_raw_shape_fragments(w, raw, RAW_FILLS)?;
    } else {
        match shape.fill {
            ShapeFill::None => w.write_event(Event::Empty(BytesStart::new("a:noFill")))?,
            ShapeFill::Solid(color) => {
                w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                write_drawing_color(w, color)?;
                w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }
        }
    }

    if shape.raw_line_unchanged && raw_has_any(raw, &["ln"]) {
        write_raw_shape_fragments(w, raw, &["ln"])?;
    } else {
        let mut line = BytesStart::new("a:ln");
        let line_width = shape.line.width_emu.map(|width| width.max(0).to_string());
        if let Some(width) = &line_width {
            line.push_attribute(("w", width.as_str()));
        }
        w.write_event(Event::Start(line))?;
        if shape.line.no_fill {
            w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
        } else if let Some(color) = shape.line.color {
            w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            write_drawing_color(w, color)?;
            w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        }
        if let Some(dash) = &shape.line.dash_style {
            let mut dash_tag = BytesStart::new("a:prstDash");
            dash_tag.push_attribute(("val", dash.as_str()));
            w.write_event(Event::Empty(dash_tag))?;
        }
        w.write_event(Event::End(BytesEnd::new("a:ln")))?;
    }

    write_remaining_raw_shape_fragments(w, raw)?;
    w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

    if let Some(text) = &shape.text {
        if shape.raw_text_unchanged {
            if let Some(raw_text) = shape.raw_text_body {
                w.get_mut().write_all(raw_text)?;
            } else {
                write_drawing_text(w, text, None)?;
            }
        } else {
            write_drawing_text(w, text, None)?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("xdr:sp")))?;
    Ok(())
}

fn raw_has_any(raw: &[u8], names: &[&str]) -> bool {
    let mut reader = quick_xml::Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if names
                    .iter()
                    .any(|name| e.local_name().as_ref() == name.as_bytes())
                {
                    return true;
                }
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buf.clear();
    }
}

fn write_raw_shape_fragments(w: &mut XmlWriter, raw: &[u8], names: &[&str]) -> WriteResult<()> {
    write_selected_raw_fragments(w, raw, |name| names.iter().any(|value| *value == name))
}

fn write_remaining_raw_shape_fragments(w: &mut XmlWriter, raw: &[u8]) -> WriteResult<()> {
    const CONFLICTING: &[&str] = &[
        "xfrm",
        "prstGeom",
        "custGeom",
        "noFill",
        "solidFill",
        "gradFill",
        "blipFill",
        "pattFill",
        "grpFill",
        "ln",
    ];
    write_selected_raw_fragments(w, raw, |name| !CONFLICTING.contains(&name))
}

fn write_selected_raw_fragments(
    w: &mut XmlWriter,
    raw: &[u8],
    mut include: impl FnMut(&str) -> bool,
) -> WriteResult<()> {
    let mut reader = quick_xml::Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                let mut fragment = Writer::new(Cursor::new(Vec::new()));
                fragment.write_event(Event::Start(e.to_owned()))?;
                let mut depth = 1u32;
                let mut inner = Vec::new();
                while depth > 0 {
                    match reader.read_event_into(&mut inner) {
                        Ok(Event::Start(e)) => {
                            depth += 1;
                            fragment.write_event(Event::Start(e.into_owned()))?;
                        }
                        Ok(Event::End(e)) => {
                            depth -= 1;
                            fragment.write_event(Event::End(e.into_owned()))?;
                        }
                        Ok(Event::Eof) => break,
                        Ok(event) => fragment.write_event(event.into_owned())?,
                        Err(error) => return Err(std::io::Error::other(error)),
                    }
                    inner.clear();
                }
                if include(&name) {
                    w.get_mut().write_all(&fragment.into_inner().into_inner())?;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                if include(&name) {
                    w.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Ok(Event::Text(ref text)) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                w.write_event(Event::Text(text.to_owned()))?;
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(std::io::Error::other(error)),
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_group_shape(
    w: &mut XmlWriter,
    name: Option<&str>,
    alt_text: Option<&str>,
    title: Option<&str>,
    hidden: bool,
    transform: &GroupTransform,
    children: &[PartChild<'_>],
    placement: &Placement,
    ctx: &mut EmitCtx<'_>,
) -> WriteResult<()> {
    let cnv_id = ctx.next_cnv_id();
    w.write_event(Event::Start(BytesStart::new("xdr:grpSp")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGrpSpPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name.unwrap_or("")));
    if let Some(descr) = alt_text {
        cnv_pr.push_attribute(("descr", descr));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    if hidden {
        cnv_pr.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvGrpSpPr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGrpSpPr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:grpSpPr")))?;
    let mut xfrm = BytesStart::new("a:xfrm");
    if placement.rot != 0 {
        xfrm.push_attribute(("rot", placement.rot.to_string().as_str()));
    }
    if placement.flip_h {
        xfrm.push_attribute(("flipH", "1"));
    }
    if placement.flip_v {
        xfrm.push_attribute(("flipV", "1"));
    }
    w.write_event(Event::Start(xfrm))?;
    write_point(w, "a:off", placement.x, placement.y)?;
    write_extent(w, "a:ext", placement.cx, placement.cy)?;
    write_point(w, "a:chOff", transform.child_x_emu, transform.child_y_emu)?;
    write_extent(w, "a:chExt", transform.child_cx_emu, transform.child_cy_emu)?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:grpSpPr")))?;

    for child in children {
        write_group_child(w, child, ctx)?;
    }

    w.write_event(Event::End(BytesEnd::new("xdr:grpSp")))?;
    Ok(())
}

fn write_group_child(
    w: &mut XmlWriter,
    child: &PartChild<'_>,
    ctx: &mut EmitCtx<'_>,
) -> WriteResult<()> {
    match &child.kind {
        PartKind::Image(_) => {
            let rid = ctx.plan.image_rids[ctx.image_i].clone();
            ctx.image_i += 1;
            let cnv_id = ctx.next_cnv_id();
            let t = child.transform;
            let xfrm = PicXfrm {
                off_x: t.x_emu,
                off_y: t.y_emu,
                cx: t.cx_emu,
                cy: t.cy_emu,
                rot: (t.rotation != 0).then_some(t.rotation),
                flip_h: t.flip_h,
                flip_v: t.flip_v,
            };
            write_picture_element(
                w,
                child.name,
                child.alt_text,
                child.title,
                child.hidden,
                &rid,
                cnv_id,
                &xfrm,
            )?;
        }
        PartKind::Group {
            transform,
            children,
        } => {
            let t = child.transform;
            let placement = Placement {
                x: t.x_emu,
                y: t.y_emu,
                cx: t.cx_emu,
                cy: t.cy_emu,
                rot: t.rotation,
                flip_h: t.flip_h,
                flip_v: t.flip_v,
            };
            write_group_shape(
                w,
                child.name,
                child.alt_text,
                child.title,
                child.hidden,
                transform,
                children,
                &placement,
                ctx,
            )?;
        }
        PartKind::Control { name } => {
            let shape_id = ctx.shape_base + ctx.control_ordinal;
            ctx.control_ordinal += 1;
            let cnv_id = ctx.next_cnv_id();
            let twin_style = ctx.twin_style;
            write_mc_a14_choice(w, |w| match twin_style {
                TwinStyle::CompatExtSp => write_control_twin_sp(
                    w,
                    name,
                    child.alt_text,
                    child.title,
                    child.macro_name.as_deref(),
                    child.control_text.as_ref(),
                    shape_id,
                    cnv_id,
                    Some(child.transform),
                ),
                TwinStyle::CompatSpFrame => write_control_twin_frame(
                    w,
                    name,
                    child.alt_text,
                    child.title,
                    shape_id,
                    Some(child.transform),
                ),
            })?;
        }
        PartKind::Shape(shape) => {
            let cnv_id = ctx.next_cnv_id();
            let t = child.transform;
            let placement = Placement {
                x: t.x_emu,
                y: t.y_emu,
                cx: t.cx_emu,
                cy: t.cy_emu,
                rot: if shape.rotation != 0 {
                    shape.rotation
                } else {
                    t.rotation
                },
                flip_h: shape.flip_h || t.flip_h,
                flip_v: shape.flip_v || t.flip_v,
            };
            write_shape_element(
                w,
                child.name,
                child.alt_text,
                child.title,
                child.hidden,
                cnv_id,
                shape,
                &placement,
            )?;
        }
        PartKind::Raw { bytes, .. } => {
            w.get_mut().write_all(bytes)?;
        }
        // No group representation for these kinds.
        PartKind::Chart | PartKind::ChartEx { .. } => {}
    }
    Ok(())
}

/// Emit the anchor element matching the object's `DrawingAnchor`
/// variant around `content`, with clientData carrying the object's
/// locked/printable flags.
fn write_anchor_wrapper(
    w: &mut XmlWriter,
    object: &PartObject<'_>,
    content: impl FnOnce(&mut XmlWriter) -> WriteResult<()>,
) -> WriteResult<()> {
    match object.anchor {
        DrawingAnchor::TwoCell { from, to, edit_as } => {
            let mut tag = BytesStart::new("xdr:twoCellAnchor");
            if let Some(ea) = edit_as {
                let s = match ea {
                    EditAs::TwoCell => "twoCell",
                    EditAs::OneCell => "oneCell",
                    EditAs::Absolute => "absolute",
                };
                tag.push_attribute(("editAs", s));
            }
            w.write_event(Event::Start(tag))?;
            w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
            write_cell_marker(w, from)?;
            w.write_event(Event::End(BytesEnd::new("xdr:from")))?;
            w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
            write_cell_marker(w, to)?;
            w.write_event(Event::End(BytesEnd::new("xdr:to")))?;
            content(w)?;
            write_client_data(w, object.locked, object.printable)?;
            w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            w.write_event(Event::Start(BytesStart::new("xdr:oneCellAnchor")))?;
            w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
            write_cell_marker(w, from)?;
            w.write_event(Event::End(BytesEnd::new("xdr:from")))?;
            write_extent(w, "xdr:ext", *width_emu, *height_emu)?;
            content(w)?;
            write_client_data(w, object.locked, object.printable)?;
            w.write_event(Event::End(BytesEnd::new("xdr:oneCellAnchor")))?;
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;
            write_point(w, "xdr:pos", *x_emu, *y_emu)?;
            write_extent(w, "xdr:ext", *width_emu, *height_emu)?;
            content(w)?;
            write_client_data(w, object.locked, object.printable)?;
            w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
        }
    }
    Ok(())
}

/// `<xdr:clientData/>` when both flags hold their defaults, else the
/// explicit fLocksWithSheet/fPrintsWithSheet attributes.
fn write_client_data(w: &mut XmlWriter, locked: bool, printable: bool) -> WriteResult<()> {
    let mut tag = BytesStart::new("xdr:clientData");
    if !locked {
        tag.push_attribute(("fLocksWithSheet", "0"));
    }
    if !printable {
        tag.push_attribute(("fPrintsWithSheet", "0"));
    }
    w.write_event(Event::Empty(tag))?;
    Ok(())
}

fn write_point(w: &mut XmlWriter, tag: &str, x: i64, y: i64) -> WriteResult<()> {
    let x_s = x.to_string();
    let y_s = y.to_string();
    w.create_element(tag)
        .with_attribute(("x", x_s.as_str()))
        .with_attribute(("y", y_s.as_str()))
        .write_empty()?;
    Ok(())
}

fn write_extent(w: &mut XmlWriter, tag: &str, cx: i64, cy: i64) -> WriteResult<()> {
    let cx_s = cx.to_string();
    let cy_s = cy.to_string();
    w.create_element(tag)
        .with_attribute(("cx", cx_s.as_str()))
        .with_attribute(("cy", cy_s.as_str()))
        .write_empty()?;
    Ok(())
}

/// Picture placement inside its container (anchor or group child
/// space).
struct PicXfrm {
    off_x: i64,
    off_y: i64,
    cx: i64,
    cy: i64,
    rot: Option<i32>,
    flip_h: bool,
    flip_v: bool,
}

/// Emit just the `<xdr:pic>` element (without the surrounding anchor
/// wrapper).
fn write_picture_element(
    w: &mut XmlWriter,
    name: Option<&str>,
    alt_text: Option<&str>,
    title: Option<&str>,
    hidden: bool,
    rid: &str,
    cnv_id: usize,
    xfrm: &PicXfrm,
) -> WriteResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:pic")))?;

    // <xdr:nvPicPr> non-visual picture properties.
    w.write_event(Event::Start(BytesStart::new("xdr:nvPicPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name.unwrap_or("")));
    if let Some(desc) = alt_text {
        cnv_pr.push_attribute(("descr", desc));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    if hidden {
        cnv_pr.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Start(BytesStart::new("xdr:cNvPicPr")))?;
    let mut pic_locks = BytesStart::new("a:picLocks");
    pic_locks.push_attribute(("noChangeAspect", "1"));
    w.write_event(Event::Empty(pic_locks))?;
    w.write_event(Event::End(BytesEnd::new("xdr:cNvPicPr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvPicPr")))?;

    // <xdr:blipFill> blip reference to the image part.
    w.write_event(Event::Start(BytesStart::new("xdr:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("xmlns:r", NS_DOC_RELS));
    blip.push_attribute(("r:embed", rid));
    w.write_event(Event::Empty(blip))?;
    w.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    w.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    w.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:blipFill")))?;

    // <xdr:spPr> shape properties: xfrm with container placement.
    w.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
    let mut xfrm_tag = BytesStart::new("a:xfrm");
    if let Some(rot) = xfrm.rot {
        xfrm_tag.push_attribute(("rot", rot.to_string().as_str()));
    }
    if xfrm.flip_h {
        xfrm_tag.push_attribute(("flipH", "1"));
    }
    if xfrm.flip_v {
        xfrm_tag.push_attribute(("flipV", "1"));
    }
    w.write_event(Event::Start(xfrm_tag))?;
    write_point(w, "a:off", xfrm.off_x, xfrm.off_y)?;
    write_extent(w, "a:ext", xfrm.cx, xfrm.cy)?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut prst = BytesStart::new("a:prstGeom");
    prst.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(prst))?;
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:pic")))?;
    Ok(())
}

/// `<mc:AlternateContent><mc:Choice Requires="a14">content</mc:Choice>
/// <mc:Fallback/></mc:AlternateContent>`.
fn write_mc_a14_choice(
    w: &mut XmlWriter,
    content: impl FnOnce(&mut XmlWriter) -> WriteResult<()>,
) -> WriteResult<()> {
    let mut mc = BytesStart::new("mc:AlternateContent");
    mc.push_attribute(("xmlns:mc", NS_MC));
    w.write_event(Event::Start(mc))?;
    let mut choice = BytesStart::new("mc:Choice");
    choice.push_attribute(("xmlns:a14", NS_A14));
    choice.push_attribute(("Requires", "a14"));
    w.write_event(Event::Start(choice))?;
    content(w)?;
    w.write_event(Event::End(BytesEnd::new("mc:Choice")))?;
    w.write_event(Event::Empty(BytesStart::new("mc:Fallback")))?;
    w.write_event(Event::End(BytesEnd::new("mc:AlternateContent")))?;
    Ok(())
}

/// The legacy-object twin of a form control, as a whole
/// AlternateContent-wrapped anchor (per MS-ODRAWXML Legacy Object
/// Wrapper). The anchor markers mirror the legacy part's conversion
/// of the control's wrapper anchor.
fn write_control_twin_anchor(
    w: &mut XmlWriter,
    object: &PartObject<'_>,
    name: &str,
    shape_id: usize,
    cnv_id: usize,
    twin_style: TwinStyle,
    metrics: &dyn DrawingMetrics,
) -> WriteResult<()> {
    let (from, to) = anchor_cell_markers(object.anchor, metrics);
    let edit_as = twin_edit_as(object.anchor);
    write_mc_a14_choice(w, |w| {
        let mut tag = BytesStart::new("xdr:twoCellAnchor");
        if let Some(ea) = edit_as {
            tag.push_attribute(("editAs", ea));
        }
        w.write_event(Event::Start(tag))?;
        w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
        write_cell_marker(w, &from)?;
        w.write_event(Event::End(BytesEnd::new("xdr:from")))?;
        w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
        write_cell_marker(w, &to)?;
        w.write_event(Event::End(BytesEnd::new("xdr:to")))?;
        match twin_style {
            TwinStyle::CompatExtSp => write_control_twin_sp(
                w,
                name,
                object.alt_text,
                object.title,
                object.macro_name.as_deref(),
                object.control_text.as_ref(),
                shape_id,
                cnv_id,
                None,
            )?,
            TwinStyle::CompatSpFrame => {
                write_control_twin_frame(w, name, object.alt_text, object.title, shape_id, None)?
            }
        }
        let mut cd = BytesStart::new("xdr:clientData");
        cd.push_attribute(("fLocksWithSheet", if object.locked { "1" } else { "0" }));
        cd.push_attribute(("fPrintsWithSheet", if object.printable { "1" } else { "0" }));
        w.write_event(Event::Empty(cd))?;
        w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
        Ok(())
    })
}

/// editAs derived from the anchor's moveWithCells/sizeWithCells
/// semantics, matching the legacy control conversion.
fn twin_edit_as(anchor: &DrawingAnchor) -> Option<&'static str> {
    let (move_wc, size_wc) = match anchor {
        DrawingAnchor::TwoCell { edit_as, .. } => {
            match edit_as.clone().unwrap_or(EditAs::TwoCell) {
                EditAs::TwoCell => (true, true),
                EditAs::OneCell => (true, false),
                EditAs::Absolute => (false, false),
            }
        }
        DrawingAnchor::OneCell { .. } => (true, false),
        DrawingAnchor::Absolute { .. } => (false, false),
    };
    match (move_wc, size_wc) {
        (true, true) => None,
        (true, false) => Some("oneCell"),
        _ => Some("absolute"),
    }
}

/// Resolve any anchor variant to concrete from/to cell markers using
/// the worksheet metrics supplied by the format adapter.
fn anchor_cell_markers(
    anchor: &DrawingAnchor,
    metrics: &dyn DrawingMetrics,
) -> (CellMarker, CellMarker) {
    match anchor.to_two_cell_with_metrics(metrics) {
        DrawingAnchor::TwoCell { from, to, .. } => (from, to),
        _ => unreachable!("to_two_cell_with_metrics always returns TwoCell"),
    }
}

/// The twin `<xdr:sp>` element (XLSX flavor). `xfrm` is zero for
/// anchored twins and carries the child transform for twins inside
/// groups.
fn write_drawing_text(
    w: &mut XmlWriter,
    text: &TwinText,
    raw_children: Option<&[u8]>,
) -> WriteResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:txBody")))?;
    let mut body_pr = BytesStart::new("a:bodyPr");
    if let Some(vertical) = text.vertical_alignment {
        let anchor = match vertical {
            TwinVerticalAlignment::Top => "t",
            TwinVerticalAlignment::Center => "ctr",
            TwinVerticalAlignment::Bottom => "b",
            TwinVerticalAlignment::Justify => "just",
            TwinVerticalAlignment::Distributed => "dist",
        };
        body_pr.push_attribute(("anchor", anchor));
    }
    w.write_event(Event::Empty(body_pr))?;
    w.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
    w.write_event(Event::Start(BytesStart::new("a:p")))?;
    if let Some(horizontal) = text.horizontal_alignment {
        let alignment = match horizontal {
            TwinHorizontalAlignment::Left => "l",
            TwinHorizontalAlignment::Center => "ctr",
            TwinHorizontalAlignment::Right => "r",
            TwinHorizontalAlignment::Justify => "just",
            TwinHorizontalAlignment::Distributed => "dist",
        };
        let mut p_pr = BytesStart::new("a:pPr");
        p_pr.push_attribute(("algn", alignment));
        w.write_event(Event::Empty(p_pr))?;
    }

    for run in &text.runs {
        w.write_event(Event::Start(BytesStart::new("a:r")))?;
        if let Some(font) = &run.font {
            let mut r_pr = BytesStart::new("a:rPr");
            r_pr.push_attribute(("lang", "en-US"));
            let size = font
                .size
                .map(|size| (size * 100.0).round().clamp(1.0, i32::MAX as f64) as i32)
                .map(|size| size.to_string());
            if let Some(size) = &size {
                r_pr.push_attribute(("sz", size.as_str()));
            }
            if let Some(bold) = font.bold {
                r_pr.push_attribute(("b", if bold { "1" } else { "0" }));
            }
            if let Some(italic) = font.italic {
                r_pr.push_attribute(("i", if italic { "1" } else { "0" }));
            }
            if let Some(underline) = font.underline {
                let value = match underline {
                    TwinUnderline::Single => "sng",
                    TwinUnderline::Double => "dbl",
                    TwinUnderline::SingleAccounting => "sng",
                    TwinUnderline::DoubleAccounting => "dbl",
                };
                r_pr.push_attribute(("u", value));
            }
            if let Some(strikethrough) = font.strikethrough {
                r_pr.push_attribute((
                    "strike",
                    if strikethrough {
                        "sngStrike"
                    } else {
                        "noStrike"
                    },
                ));
            }
            let baseline = font.baseline.map(|baseline| baseline.to_string());
            if let Some(baseline) = &baseline {
                r_pr.push_attribute(("baseline", baseline.as_str()));
            }
            w.write_event(Event::Start(r_pr))?;
            if let Some(color) = font.color {
                w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                write_drawing_color(w, color)?;
                w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }
            if let Some(name) = &font.name {
                let mut latin = BytesStart::new("a:latin");
                latin.push_attribute(("typeface", name.as_str()));
                w.write_event(Event::Empty(latin))?;
            }
            w.write_event(Event::End(BytesEnd::new("a:rPr")))?;
        }
        let mut t = BytesStart::new("a:t");
        if run.text.chars().next().is_some_and(char::is_whitespace)
            || run
                .text
                .chars()
                .next_back()
                .is_some_and(char::is_whitespace)
        {
            t.push_attribute(("xml:space", "preserve"));
        }
        w.write_event(Event::Start(t))?;
        w.write_event(Event::Text(BytesText::new(&run.text)))?;
        w.write_event(Event::End(BytesEnd::new("a:t")))?;
        w.write_event(Event::End(BytesEnd::new("a:r")))?;
    }
    let mut end_pr = BytesStart::new("a:endParaRPr");
    end_pr.push_attribute(("lang", "en-US"));
    w.write_event(Event::Empty(end_pr))?;
    w.write_event(Event::End(BytesEnd::new("a:p")))?;
    if let Some(raw_children) = raw_children {
        w.get_mut().write_all(raw_children)?;
    }
    w.write_event(Event::End(BytesEnd::new("xdr:txBody")))?;
    Ok(())
}

fn write_drawing_color(w: &mut XmlWriter, color: TwinColor) -> WriteResult<()> {
    match color {
        TwinColor::Rgb { r, g, b } => {
            let value = format!("{r:02X}{g:02X}{b:02X}");
            let mut color = BytesStart::new("a:srgbClr");
            color.push_attribute(("val", value.as_str()));
            w.write_event(Event::Empty(color))?;
        }
        TwinColor::Theme { index, tint } => {
            let scheme = match index {
                0 => "lt1",
                1 => "dk1",
                2 => "lt2",
                3 => "dk2",
                4 => "accent1",
                5 => "accent2",
                6 => "accent3",
                7 => "accent4",
                8 => "accent5",
                9 => "accent6",
                _ => "dk1",
            };
            let mut color = BytesStart::new("a:schemeClr");
            color.push_attribute(("val", scheme));
            if tint == 0 {
                w.write_event(Event::Empty(color))?;
            } else {
                w.write_event(Event::Start(color))?;
                let tint = i32::from(tint).clamp(-100, 100);
                let lum_mod = if tint > 0 {
                    100_000 - tint * 1_000
                } else {
                    100_000 + tint * 1_000
                };
                let mut lum_mod_tag = BytesStart::new("a:lumMod");
                let lum_mod = lum_mod.to_string();
                lum_mod_tag.push_attribute(("val", lum_mod.as_str()));
                w.write_event(Event::Empty(lum_mod_tag))?;
                if tint > 0 {
                    let mut lum_off_tag = BytesStart::new("a:lumOff");
                    let lum_off = (tint * 1_000).to_string();
                    lum_off_tag.push_attribute(("val", lum_off.as_str()));
                    w.write_event(Event::Empty(lum_off_tag))?;
                }
                w.write_event(Event::End(BytesEnd::new("a:schemeClr")))?;
            }
        }
    }
    Ok(())
}

fn write_control_twin_sp(
    w: &mut XmlWriter,
    name: &str,
    alt_text: Option<&str>,
    title: Option<&str>,
    macro_name: Option<&str>,
    text: Option<&TwinText>,
    shape_id: usize,
    cnv_id: usize,
    xfrm: Option<&ChildTransform>,
) -> WriteResult<()> {
    let mut sp = BytesStart::new("xdr:sp");
    sp.push_attribute(("macro", macro_name.unwrap_or("")));
    sp.push_attribute(("textlink", ""));
    w.write_event(Event::Start(sp))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvSpPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name));
    cnv_pr.push_attribute(("hidden", "1"));
    if let Some(alt_text) = alt_text {
        cnv_pr.push_attribute(("descr", alt_text));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    w.write_event(Event::Start(cnv_pr))?;
    w.write_event(Event::Start(BytesStart::new("a:extLst")))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("uri", "{63B3BB69-23CF-44E3-9099-C40C66FF867C}"));
    w.write_event(Event::Start(ext))?;
    let spid = format!("_x0000_s{shape_id}");
    let mut compat = BytesStart::new("a14:compatExt");
    compat.push_attribute(("spid", spid.as_str()));
    w.write_event(Event::Empty(compat))?;
    w.write_event(Event::End(BytesEnd::new("a:ext")))?;
    w.write_event(Event::End(BytesEnd::new("a:extLst")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:cNvPr")))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvSpPr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvSpPr")))?;

    let mut sp_pr = BytesStart::new("xdr:spPr");
    sp_pr.push_attribute(("bwMode", "auto"));
    w.write_event(Event::Start(sp_pr))?;
    let mut xfrm_tag = BytesStart::new("a:xfrm");
    if let Some(t) = xfrm {
        if t.rotation != 0 {
            xfrm_tag.push_attribute(("rot", t.rotation.to_string().as_str()));
        }
        if t.flip_h {
            xfrm_tag.push_attribute(("flipH", "1"));
        }
        if t.flip_v {
            xfrm_tag.push_attribute(("flipV", "1"));
        }
    }
    w.write_event(Event::Start(xfrm_tag))?;
    let (x, y, cx, cy) = match xfrm {
        Some(t) => (t.x_emu, t.y_emu, t.cx_emu, t.cy_emu),
        None => (0, 0, 0, 0),
    };
    write_point(w, "a:off", x, y)?;
    write_extent(w, "a:ext", cx, cy)?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut prst = BytesStart::new("a:prstGeom");
    prst.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(prst))?;
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    // Fixed hidden fill/line markup per Excel's twin shape.
    w.get_mut().write_all(
        concat!(
            "<a:noFill/><a:ln><a:noFill/></a:ln>",
            "<a:extLst>",
            "<a:ext uri=\"{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}\">",
            "<a14:hiddenFill><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a14:hiddenFill>",
            "</a:ext>",
            "<a:ext uri=\"{91240B29-F687-4F45-9708-019B960494DF}\">",
            "<a14:hiddenLine w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>",
            "<a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a14:hiddenLine>",
            "</a:ext>",
            "</a:extLst>"
        )
        .as_bytes(),
    )?;
    w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

    if let Some(text) = text {
        write_drawing_text(w, text, None)?;
    }

    w.write_event(Event::End(BytesEnd::new("xdr:sp")))?;
    Ok(())
}

/// The twin `<xdr:graphicFrame>` element (XLSB flavor, per Excel's
/// own emission): a com14:compatSp graphicData whose spid points at
/// the VML shape. Excel reuses the numeric shape id as the cNvPr id.
fn write_control_twin_frame(
    w: &mut XmlWriter,
    name: &str,
    alt_text: Option<&str>,
    title: Option<&str>,
    shape_id: usize,
    xfrm: Option<&ChildTransform>,
) -> WriteResult<()> {
    let mut gf = BytesStart::new("xdr:graphicFrame");
    gf.push_attribute(("macro", ""));
    w.write_event(Event::Start(gf))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id_s = shape_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name));
    if let Some(alt_text) = alt_text {
        cnv_pr.push_attribute(("descr", alt_text));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Start(BytesStart::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::Empty(BytesStart::new("a:graphicFrameLocks")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

    let mut xfrm_tag = BytesStart::new("xdr:xfrm");
    if let Some(t) = xfrm {
        if t.rotation != 0 {
            xfrm_tag.push_attribute(("rot", t.rotation.to_string().as_str()));
        }
        if t.flip_h {
            xfrm_tag.push_attribute(("flipH", "1"));
        }
        if t.flip_v {
            xfrm_tag.push_attribute(("flipV", "1"));
        }
    }
    w.write_event(Event::Start(xfrm_tag))?;
    let (x, y, cx, cy) = match xfrm {
        Some(t) => (t.x_emu, t.y_emu, t.cx_emu, t.cy_emu),
        None => (0, 0, 0, 0),
    };
    write_point(w, "a:off", x, y)?;
    write_extent(w, "a:ext", cx, cy)?;
    w.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

    w.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", NS_COM14));
    w.write_event(Event::Start(gd))?;
    let spid = format!("_x0000_s{shape_id}");
    let mut compat = BytesStart::new("com14:compatSp");
    compat.push_attribute(("xmlns:com14", NS_COM14));
    compat.push_attribute(("spid", spid.as_str()));
    w.write_event(Event::Empty(compat))?;
    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;
    Ok(())
}

fn write_chart_frame(
    w: &mut XmlWriter,
    rid: &str,
    cnv_id: usize,
    name: &str,
    alt_text: Option<&str>,
    title: Option<&str>,
    hidden: bool,
) -> WriteResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:graphicFrame")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name));
    if let Some(alt_text) = alt_text {
        cnv_pr.push_attribute(("descr", alt_text));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    if hidden {
        cnv_pr.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:xfrm")))?;
    write_point(w, "a:off", 0, 0)?;
    write_extent(w, "a:ext", 9525000, 6096000)?;
    w.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

    w.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", NS_CHART));
    w.write_event(Event::Start(gd))?;

    let mut chart_el = BytesStart::new("c:chart");
    chart_el.push_attribute(("xmlns:c", NS_CHART));
    chart_el.push_attribute(("r:id", rid));
    w.write_event(Event::Empty(chart_el))?;

    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;
    Ok(())
}

fn write_chartex_frame(
    w: &mut XmlWriter,
    rid: &str,
    cnv_id: usize,
    name: &str,
    alt_text: Option<&str>,
    title: Option<&str>,
    hidden: bool,
    raw_mc_fallback: Option<&[u8]>,
) -> WriteResult<()> {
    let mut mc_tag = BytesStart::new("mc:AlternateContent");
    mc_tag.push_attribute(("xmlns:mc", NS_MC));
    w.write_event(Event::Start(mc_tag))?;

    let mut choice_tag = BytesStart::new("mc:Choice");
    choice_tag.push_attribute(("xmlns:cx1", NS_CX1));
    choice_tag.push_attribute(("Requires", "cx1"));
    w.write_event(Event::Start(choice_tag))?;

    let mut gf_tag = BytesStart::new("xdr:graphicFrame");
    gf_tag.push_attribute(("macro", ""));
    w.write_event(Event::Start(gf_tag))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name));
    if let Some(alt_text) = alt_text {
        cnv_pr.push_attribute(("descr", alt_text));
    }
    if let Some(title) = title {
        cnv_pr.push_attribute(("title", title));
    }
    if hidden {
        cnv_pr.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:xfrm")))?;
    write_point(w, "a:off", 0, 0)?;
    write_extent(w, "a:ext", 0, 0)?;
    w.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

    w.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", NS_CX));
    w.write_event(Event::Start(gd))?;

    let mut cx_chart = BytesStart::new("cx:chart");
    cx_chart.push_attribute(("xmlns:cx", NS_CX));
    cx_chart.push_attribute(("xmlns:r", NS_DOC_RELS));
    cx_chart.push_attribute(("r:id", rid));
    w.write_event(Event::Empty(cx_chart))?;

    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;

    w.write_event(Event::End(BytesEnd::new("mc:Choice")))?;

    w.write_event(Event::Start(BytesStart::new("mc:Fallback")))?;
    if let Some(raw) = raw_mc_fallback {
        w.get_mut().write_all(raw)?;
    }
    w.write_event(Event::End(BytesEnd::new("mc:Fallback")))?;

    w.write_event(Event::End(BytesEnd::new("mc:AlternateContent")))?;
    Ok(())
}

/// Serialize a chartsheet drawing part: one absolute-anchored chart
/// frame (rId1) followed by preserved raw anchors.
pub fn write_chartsheet_drawing_part(raw_drawing_objects: &[Vec<u8>]) -> WriteResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut tag = BytesStart::new("xdr:wsDr");
    tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
    tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
    tag.push_attribute(("xmlns:r", NS_DOC_RELS));
    w.write_event(Event::Start(tag))?;

    w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;
    write_point(&mut w, "xdr:pos", 0, 0)?;
    write_extent(&mut w, "xdr:ext", 9144000, 6858000)?;
    write_chart_frame(&mut w, "rId1", 2, "Chart 1", None, None, false)?;
    w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;

    for raw in raw_drawing_objects {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
    Ok(w.into_inner().into_inner())
}

fn write_cell_marker(w: &mut XmlWriter, marker: &CellMarker) -> WriteResult<()> {
    let col_s = marker.col.to_string();
    w.create_element("xdr:col")
        .write_text_content(BytesText::new(&col_s))?;
    let col_off_s = marker.col_offset_emu.to_string();
    w.create_element("xdr:colOff")
        .write_text_content(BytesText::new(&col_off_s))?;
    let row_s = marker.row.to_string();
    w.create_element("xdr:row")
        .write_text_content(BytesText::new(&row_s))?;
    let row_off_s = marker.row_offset_emu.to_string();
    w.create_element("xdr:rowOff")
        .write_text_content(BytesText::new(&row_off_s))?;
    Ok(())
}

/// Serialize a relationships part from planned rels.
pub fn write_rels_part(rels: &[PlannedRel]) -> WriteResult<Vec<u8>> {
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut tag = BytesStart::new("Relationships");
    tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
    w.write_event(Event::Start(tag))?;

    for rel in rels {
        let mut el = w
            .create_element("Relationship")
            .with_attribute(("Id", rel.id.as_str()))
            .with_attribute(("Type", rel.rel_type.as_str()))
            .with_attribute(("Target", rel.target.as_str()));
        if rel.external {
            el = el.with_attribute(("TargetMode", "External"));
        }
        el.write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("Relationships")))?;
    Ok(w.into_inner().into_inner())
}
