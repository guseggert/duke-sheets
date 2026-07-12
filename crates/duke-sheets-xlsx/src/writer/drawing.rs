use std::collections::HashSet;
use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_chart::{CellMarker, Chart, DrawingAnchor, EmbeddedImage};
use duke_sheets_core::{
    ChildTransform, DrawingKind, DrawingMeta, DrawingObject, Group, GroupChild, RawRel, Worksheet,
};
use duke_sheets_vml::anchor_cell_markers;

use super::{write_xml_part, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS, RT_CHART};

const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const NS_MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CX1: &str = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const NS_A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const RT_CHART_EX: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
/// OOXML relationship type for embedded image parts (`xl/media/*`).
pub(super) const RT_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// One relationship of a sheet's drawing part.
pub(super) struct PlannedRel {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub external: bool,
}

/// Relationship-id assignments for a sheet's drawing part. Writer
/// generated ids never collide with ids preserved from raw entries.
pub(super) struct DrawingPlan {
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

fn raw_rels_of(object: &DrawingObject) -> Option<&[RawRel]> {
    match &object.kind {
        DrawingKind::Raw(raw) => Some(&raw.rels),
        _ => None,
    }
}

/// Assign relationship ids for the sheet's drawing part. Raw entries
/// keep their original ids; generated ids skip over them.
pub(super) fn plan_drawing_rels(
    sheet: &Worksheet,
    chart_globals: &[usize],
    chartex_globals: &[usize],
    image_parts: &[(usize, &'static str)],
) -> DrawingPlan {
    let mut taken_ids: HashSet<String> = HashSet::new();
    let mut taken_nums: HashSet<usize> = HashSet::new();
    for object in sheet.drawings() {
        for rel in raw_rels_of(object).unwrap_or(&[]) {
            if taken_ids.insert(rel.id.clone()) {
                if let Some(num) = rel.id.strip_prefix("rId").and_then(|n| n.parse().ok()) {
                    taken_nums.insert(num);
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

    fn plan_group(
        group: &Group,
        plan: &mut DrawingPlan,
        alloc: &mut impl FnMut() -> String,
        image_parts: &[(usize, &'static str)],
        image_i: &mut usize,
    ) {
        for child in &group.children {
            match &child.kind {
                DrawingKind::Image(_) => plan_image(plan, alloc, image_parts, image_i),
                DrawingKind::Group(inner) => {
                    plan_group(inner, plan, alloc, image_parts, image_i)
                }
                _ => {}
            }
        }
    }

    for object in sheet.drawings() {
        match &object.kind {
            DrawingKind::Chart(chart) => {
                if is_unsupported(chart) {
                    continue;
                }
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
            DrawingKind::ChartEx(_) => {
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
            DrawingKind::Image(_) => plan_image(&mut plan, &mut alloc, image_parts, &mut image_i),
            DrawingKind::Group(group) => {
                plan_group(group, &mut plan, &mut alloc, image_parts, &mut image_i)
            }
            DrawingKind::Raw(raw) => {
                for rel in &raw.rels {
                    if seen_raw.insert(rel.id.clone()) {
                        plan.rels.push(PlannedRel {
                            id: rel.id.clone(),
                            rel_type: rel.rel_type.clone(),
                            target: rel.target.clone(),
                            external: rel.external,
                        });
                    }
                }
            }
            DrawingKind::FormControl(_) | DrawingKind::Comment { .. } => {}
        }
    }
    plan
}

pub(super) fn is_unsupported(chart: &Chart) -> bool {
    matches!(chart.chart_type, duke_sheets_chart::ChartType::Unsupported(_))
}

/// Walk state for emitting a drawing part in drawing-list order.
struct EmitCtx<'a> {
    plan: &'a DrawingPlan,
    chart_i: usize,
    chartex_i: usize,
    image_i: usize,
    /// cNvPr id space, unique across the part, sequential in emission
    /// order (base 2 by convention).
    cnv_id: usize,
    /// 1-based sequence of chart graphic frames, for default names.
    frame_seq: usize,
    /// Placed-control ordinal (drives twin shape ids and names).
    control_ordinal: usize,
    shape_base: usize,
    comment_count: usize,
}

impl EmitCtx<'_> {
    fn next_cnv_id(&mut self) -> usize {
        let id = self.cnv_id;
        self.cnv_id += 1;
        id
    }
}

/// Write one sheet's drawing part: every drawing object in list
/// order. Comments have no native-part presence; form controls emit
/// their a14 placeholder twins; raw entries pass through verbatim.
pub(super) fn write_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    sheet: &Worksheet,
    sheet_index: usize,
    plan: &DrawingPlan,
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    let comment_count = sheet.comment_count();
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("xdr:wsDr");
        tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
        tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        w.write_event(Event::Start(tag))?;

        let mut ctx = EmitCtx {
            plan,
            chart_i: 0,
            chartex_i: 0,
            image_i: 0,
            cnv_id: 2,
            frame_seq: 0,
            control_ordinal: 0,
            shape_base: (sheet_index + 1) * 1024 + 1 + comment_count,
            comment_count,
        };

        for object in sheet.drawings() {
            match &object.kind {
                DrawingKind::Comment { .. } => {}
                DrawingKind::Chart(chart) => {
                    if is_unsupported(chart) {
                        continue;
                    }
                    ctx.frame_seq += 1;
                    let rid = plan.chart_rids[ctx.chart_i].clone();
                    ctx.chart_i += 1;
                    let cnv_id = ctx.next_cnv_id();
                    let name = format!("Chart {}", ctx.frame_seq);
                    write_anchor_wrapper(w, object, |w| {
                        write_chart_frame(w, &rid, cnv_id, &name)
                    })?;
                }
                DrawingKind::ChartEx(chart) => {
                    ctx.frame_seq += 1;
                    let rid = plan.chartex_rids[ctx.chartex_i].clone();
                    ctx.chartex_i += 1;
                    let cnv_id = ctx.next_cnv_id();
                    let name = format!("Chart {}", ctx.frame_seq);
                    let fallback = chart.raw_mc_fallback.clone();
                    write_anchor_wrapper(w, object, |w| {
                        write_chartex_frame(w, &rid, cnv_id, &name, fallback.as_deref())
                    })?;
                }
                DrawingKind::Image(image) => {
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
                    write_anchor_wrapper(w, object, |w| {
                        write_picture_element(w, &object.meta, image, &rid, cnv_id, &xfrm)
                    })?;
                }
                DrawingKind::Group(group) => {
                    let transform = &group.transform;
                    let placement = Placement {
                        x: transform.x_emu,
                        y: transform.y_emu,
                        cx: transform.cx_emu,
                        cy: transform.cy_emu,
                        rot: transform.rotation,
                        flip_h: transform.flip_h,
                        flip_v: transform.flip_v,
                    };
                    // Split borrow: the closure needs &mut ctx while
                    // the wrapper borrows the object.
                    let meta = &object.meta;
                    let anchor_object = object;
                    write_anchor_wrapper(w, anchor_object, |w| {
                        write_group_shape(w, meta, group, &placement, &mut ctx)
                    })?;
                }
                DrawingKind::FormControl(control) => {
                    let shape_id = ctx.shape_base + ctx.control_ordinal;
                    let name = object.meta.name.clone().unwrap_or_else(|| {
                        super::form_controls::default_control_name(
                            &control.kind,
                            ctx.comment_count + ctx.control_ordinal + 1,
                        )
                    });
                    ctx.control_ordinal += 1;
                    let cnv_id = ctx.next_cnv_id();
                    write_control_twin_anchor(w, object, &name, shape_id, cnv_id)?;
                }
                DrawingKind::Raw(raw) => {
                    w.get_mut().write_all(&raw.bytes)?;
                }
            }
        }

        w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
        Ok(())
    })
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

fn write_group_shape(
    w: &mut XmlWriter,
    meta: &DrawingMeta,
    group: &Group,
    placement: &Placement,
    ctx: &mut EmitCtx<'_>,
) -> XlsxResult<()> {
    let cnv_id = ctx.next_cnv_id();
    w.write_event(Event::Start(BytesStart::new("xdr:grpSp")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGrpSpPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", meta.name.as_deref().unwrap_or("")));
    if let Some(descr) = meta.alt_text.as_deref() {
        cnv_pr.push_attribute(("descr", descr));
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
    write_point(w, "a:chOff", group.transform.child_x_emu, group.transform.child_y_emu)?;
    write_extent(
        w,
        "a:chExt",
        group.transform.child_cx_emu,
        group.transform.child_cy_emu,
    )?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:grpSpPr")))?;

    for child in &group.children {
        write_group_child(w, child, ctx)?;
    }

    w.write_event(Event::End(BytesEnd::new("xdr:grpSp")))?;
    Ok(())
}

fn write_group_child(
    w: &mut XmlWriter,
    child: &GroupChild,
    ctx: &mut EmitCtx<'_>,
) -> XlsxResult<()> {
    match &child.kind {
        DrawingKind::Image(image) => {
            let rid = ctx.plan.image_rids[ctx.image_i].clone();
            ctx.image_i += 1;
            let cnv_id = ctx.next_cnv_id();
            let t = &child.transform;
            let xfrm = PicXfrm {
                off_x: t.x_emu,
                off_y: t.y_emu,
                cx: t.cx_emu,
                cy: t.cy_emu,
                rot: (t.rotation != 0).then_some(t.rotation),
                flip_h: t.flip_h,
                flip_v: t.flip_v,
            };
            write_picture_element(w, &child.meta, image, &rid, cnv_id, &xfrm)?;
        }
        DrawingKind::Group(inner) => {
            let t = &child.transform;
            let placement = Placement {
                x: t.x_emu,
                y: t.y_emu,
                cx: t.cx_emu,
                cy: t.cy_emu,
                rot: t.rotation,
                flip_h: t.flip_h,
                flip_v: t.flip_v,
            };
            write_group_shape(w, &child.meta, inner, &placement, ctx)?;
        }
        DrawingKind::FormControl(control) => {
            let shape_id = ctx.shape_base + ctx.control_ordinal;
            let name = child.meta.name.clone().unwrap_or_else(|| {
                super::form_controls::default_control_name(
                    &control.kind,
                    ctx.comment_count + ctx.control_ordinal + 1,
                )
            });
            ctx.control_ordinal += 1;
            let cnv_id = ctx.next_cnv_id();
            write_mc_a14_choice(w, |w| {
                write_control_twin_sp(w, &name, shape_id, cnv_id, Some(&child.transform))
            })?;
        }
        DrawingKind::Raw(raw) => {
            w.get_mut().write_all(&raw.bytes)?;
        }
        // No group representation for these kinds.
        DrawingKind::Chart(_) | DrawingKind::ChartEx(_) | DrawingKind::Comment { .. } => {}
    }
    Ok(())
}

/// Emit the anchor element matching the object's `DrawingAnchor`
/// variant around `content`, with clientData carrying the object's
/// locked/printable flags.
fn write_anchor_wrapper(
    w: &mut XmlWriter,
    object: &DrawingObject,
    content: impl FnOnce(&mut XmlWriter) -> XlsxResult<()>,
) -> XlsxResult<()> {
    match &object.anchor {
        DrawingAnchor::TwoCell { from, to, edit_as } => {
            let mut tag = BytesStart::new("xdr:twoCellAnchor");
            if let Some(ea) = edit_as {
                let s = match ea {
                    duke_sheets_chart::EditAs::TwoCell => "twoCell",
                    duke_sheets_chart::EditAs::OneCell => "oneCell",
                    duke_sheets_chart::EditAs::Absolute => "absolute",
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
            write_client_data(w, object.meta.locked, object.meta.printable)?;
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
            write_client_data(w, object.meta.locked, object.meta.printable)?;
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
            write_client_data(w, object.meta.locked, object.meta.printable)?;
            w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
        }
    }
    Ok(())
}

/// `<xdr:clientData/>` when both flags hold their defaults, else the
/// explicit fLocksWithSheet/fPrintsWithSheet attributes.
fn write_client_data(w: &mut XmlWriter, locked: bool, printable: bool) -> XlsxResult<()> {
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

fn write_point(w: &mut XmlWriter, tag: &str, x: i64, y: i64) -> XlsxResult<()> {
    let x_s = x.to_string();
    let y_s = y.to_string();
    w.create_element(tag)
        .with_attribute(("x", x_s.as_str()))
        .with_attribute(("y", y_s.as_str()))
        .write_empty()?;
    Ok(())
}

fn write_extent(w: &mut XmlWriter, tag: &str, cx: i64, cy: i64) -> XlsxResult<()> {
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
    meta: &DrawingMeta,
    _image: &EmbeddedImage,
    rid: &str,
    cnv_id: usize,
    xfrm: &PicXfrm,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:pic")))?;

    // <xdr:nvPicPr> non-visual picture properties.
    w.write_event(Event::Start(BytesStart::new("xdr:nvPicPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", meta.name.as_deref().unwrap_or("")));
    if let Some(desc) = meta.alt_text.as_deref() {
        cnv_pr.push_attribute(("descr", desc));
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
    content: impl FnOnce(&mut XmlWriter) -> XlsxResult<()>,
) -> XlsxResult<()> {
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

/// The a14 legacy-object twin of a form control, as a whole
/// AlternateContent-wrapped anchor (per MS-ODRAWXML Legacy Object
/// Wrapper). The anchor markers mirror the `<controls>` block
/// conversion of the control's wrapper anchor.
fn write_control_twin_anchor(
    w: &mut XmlWriter,
    object: &DrawingObject,
    name: &str,
    shape_id: usize,
    cnv_id: usize,
) -> XlsxResult<()> {
    let (from, to) = anchor_cell_markers(&object.anchor);
    let edit_as = twin_edit_as(&object.anchor);
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
        write_control_twin_sp(w, name, shape_id, cnv_id, None)?;
        let mut cd = BytesStart::new("xdr:clientData");
        cd.push_attribute((
            "fLocksWithSheet",
            if object.meta.locked { "1" } else { "0" },
        ));
        cd.push_attribute((
            "fPrintsWithSheet",
            if object.meta.printable { "1" } else { "0" },
        ));
        w.write_event(Event::Empty(cd))?;
        w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
        Ok(())
    })
}

/// editAs derived from the anchor's moveWithCells/sizeWithCells
/// semantics, matching the `<controls>` block conversion.
fn twin_edit_as(anchor: &DrawingAnchor) -> Option<&'static str> {
    let (move_wc, size_wc) = match anchor {
        DrawingAnchor::TwoCell { edit_as, .. } => {
            match edit_as.clone().unwrap_or(duke_sheets_chart::EditAs::TwoCell) {
                duke_sheets_chart::EditAs::TwoCell => (true, true),
                duke_sheets_chart::EditAs::OneCell => (true, false),
                duke_sheets_chart::EditAs::Absolute => (false, false),
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

/// The twin `<xdr:sp>` element. `xfrm` is zero for anchored twins and
/// carries the child transform for twins inside groups.
fn write_control_twin_sp(
    w: &mut XmlWriter,
    name: &str,
    shape_id: usize,
    cnv_id: usize,
    xfrm: Option<&ChildTransform>,
) -> XlsxResult<()> {
    let mut sp = BytesStart::new("xdr:sp");
    sp.push_attribute(("macro", ""));
    sp.push_attribute(("textlink", ""));
    w.write_event(Event::Start(sp))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvSpPr")))?;
    let cnv_id_s = cnv_id.to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id_s.as_str()));
    cnv_pr.push_attribute(("name", name));
    cnv_pr.push_attribute(("hidden", "1"));
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

    w.write_event(Event::End(BytesEnd::new("xdr:sp")))?;
    Ok(())
}

fn write_chart_frame(w: &mut XmlWriter, rid: &str, cnv_id: usize, name: &str) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:graphicFrame")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id_s = cnv_id.to_string();
    w.create_element("xdr:cNvPr")
        .with_attribute(("id", cnv_id_s.as_str()))
        .with_attribute(("name", name))
        .write_empty()?;
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
    raw_mc_fallback: Option<&[u8]>,
) -> XlsxResult<()> {
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
    w.create_element("xdr:cNvPr")
        .with_attribute(("id", cnv_id_s.as_str()))
        .with_attribute(("name", name))
        .write_empty()?;
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

pub(super) fn write_chartsheet_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    _chart: &Chart,
    raw_drawing_objects: &[Vec<u8>],
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("xdr:wsDr");
        tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
        tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        w.write_event(Event::Start(tag))?;

        write_absolute_anchor(w, "rId1", 0)?;

        for raw in raw_drawing_objects {
            w.get_mut().write_all(raw)?;
        }

        w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
        Ok(())
    })
}

fn write_absolute_anchor(w: &mut XmlWriter, rid: &str, chart_idx: usize) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;

    write_point(w, "xdr:pos", 0, 0)?;
    write_extent(w, "xdr:ext", 9144000, 6858000)?;

    write_chart_frame(w, rid, chart_idx + 2, &format!("Chart {}", chart_idx + 1))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
    Ok(())
}

fn write_cell_marker(w: &mut XmlWriter, marker: &CellMarker) -> XlsxResult<()> {
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

/// Write a drawing part's .rels from its plan.
pub(super) fn write_drawing_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    drawing_num: usize,
    rels: &[PlannedRel],
) -> XlsxResult<()> {
    let path = format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_num);
    write_xml_part(zip, &path, |w| {
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
        Ok(())
    })
}

/// Chartsheet drawing rels: at most one chart.
pub(super) fn write_chartsheet_drawing_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    drawing_num: usize,
    chart_num: Option<usize>,
) -> XlsxResult<()> {
    let rels: Vec<PlannedRel> = chart_num
        .map(|cn| PlannedRel {
            id: "rId1".to_string(),
            rel_type: RT_CHART.to_string(),
            target: format!("../charts/chart{}.xml", cn),
            external: false,
        })
        .into_iter()
        .collect();
    write_drawing_rels(zip, drawing_num, &rels)
}

/// Map an `ImageFormat` to the file extension used in `xl/media/`.
pub(super) fn image_format_extension(fmt: duke_sheets_chart::ImageFormat) -> &'static str {
    use duke_sheets_chart::ImageFormat;
    match fmt {
        ImageFormat::Png => "png",
        ImageFormat::Jpeg => "jpeg",
        ImageFormat::Gif => "gif",
        ImageFormat::Bmp => "bmp",
        ImageFormat::Tiff => "tiff",
        ImageFormat::Emf => "emf",
        ImageFormat::Wmf => "wmf",
        ImageFormat::Svg => "svg",
    }
}

/// The IANA MIME type for an image format in `[Content_Types].xml`.
pub(super) fn image_format_mime(fmt: duke_sheets_chart::ImageFormat) -> &'static str {
    use duke_sheets_chart::ImageFormat;
    match fmt {
        ImageFormat::Png => "image/png",
        ImageFormat::Jpeg => "image/jpeg",
        ImageFormat::Gif => "image/gif",
        ImageFormat::Bmp => "image/bmp",
        ImageFormat::Tiff => "image/tiff",
        ImageFormat::Emf => "image/x-emf",
        ImageFormat::Wmf => "image/x-wmf",
        ImageFormat::Svg => "image/svg+xml",
    }
}

/// Write the raw image bytes as a part `xl/media/imageN.<ext>` inside
/// the zip archive.
pub(super) fn write_media_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    global_num: usize,
    ext: &str,
    bytes: &[u8],
) -> XlsxResult<()> {
    let path = format!("xl/media/image{global_num}.{ext}");
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(bytes)?;
    Ok(())
}
