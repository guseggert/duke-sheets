use std::io::{Seek, Write};

use duke_sheets_chart::drawing_part::write::{
    self as part_write, PartChild, PartKind, PartObject, PartRel, PartShape, TwinStyle,
};
use duke_sheets_chart::drawing_part::{
    ShapeFill as PartShapeFill, ShapeLine as PartShapeLine, TwinText,
};
use duke_sheets_chart::Chart;
use duke_sheets_core::{
    DrawingKind, DrawingMeta, FormControl, Shape, ShapeFill, ShapeGeometry, Worksheet,
};

fn control_twin_text(control: &FormControl) -> Option<TwinText> {
    control.caption().map(|text| text.to_drawing_part_text())
}

fn shape_part(shape: &Shape) -> PartShape<'_> {
    let geometry = match &shape.geometry {
        ShapeGeometry::Preset(name) => name.as_str(),
    };
    let fill = match shape.fill {
        ShapeFill::None => PartShapeFill::None,
        ShapeFill::Solid(color) => duke_sheets_core::drawing::color_to_drawing_part(color)
            .map(PartShapeFill::Solid)
            .unwrap_or(PartShapeFill::None),
    };
    PartShape {
        geometry,
        fill,
        line: PartShapeLine {
            color: shape
                .line
                .color
                .and_then(duke_sheets_core::drawing::color_to_drawing_part),
            width_emu: shape.line.width_emu,
            dash_style: shape.line.dash_style.clone(),
            no_fill: shape.line.no_fill,
        },
        text: shape.text.as_ref().map(|text| text.to_drawing_part_text()),
        rotation: shape.rotation,
        flip_h: shape.flip_h,
        flip_v: shape.flip_v,
        raw_shape_properties: shape.raw_shape_properties.as_deref(),
        raw_text_body: shape.raw_text_body.as_deref(),
        raw_geometry_unchanged: shape.preserved_geometry_unchanged(),
        raw_fill_unchanged: shape.preserved_fill_unchanged(),
        raw_line_unchanged: shape.preserved_line_unchanged(),
        raw_text_unchanged: shape.preserved_text_unchanged(),
    }
}

use super::XlsxResult;

pub(super) use duke_sheets_chart::drawing_part::write::DrawingPlan;
pub(super) use duke_sheets_chart::drawing_part::{image_format_extension, image_format_mime};

pub(super) fn is_unsupported(chart: &Chart) -> bool {
    matches!(
        chart.chart_type,
        duke_sheets_chart::ChartType::Unsupported(_)
    )
}

/// Convert one drawing node into its part-emission form. Comments and
/// unsupported charts have no drawing-part presence. `control_ordinal`
/// advances across the whole tree in emission (depth-first) order and
/// drives default control names.
fn convert_kind<'a>(
    kind: &'a DrawingKind,
    meta: &'a DrawingMeta,
    comment_count: usize,
    control_ordinal: &mut usize,
) -> Option<PartKind<'a>> {
    match kind {
        DrawingKind::Comment { .. } => None,
        DrawingKind::Chart(chart) => {
            if is_unsupported(chart) {
                None
            } else {
                Some(PartKind::Chart)
            }
        }
        DrawingKind::ChartEx(chart) => Some(PartKind::ChartEx {
            fallback: chart.raw_mc_fallback.as_deref(),
        }),
        DrawingKind::Image(image) => Some(PartKind::Image(image)),
        DrawingKind::Shape(shape) => Some(PartKind::Shape(shape_part(shape))),
        DrawingKind::FormControl(control) => {
            let name = meta.name.clone().unwrap_or_else(|| {
                super::form_controls::default_control_name(
                    &control.kind,
                    comment_count + *control_ordinal + 1,
                )
            });
            *control_ordinal += 1;
            Some(PartKind::Control { name })
        }
        DrawingKind::Group(group) => {
            let children = group
                .children
                .iter()
                .filter_map(|child| {
                    convert_kind(&child.kind, &child.meta, comment_count, control_ordinal).map(
                        |kind| PartChild {
                            name: child.meta.name.as_deref(),
                            alt_text: child.meta.alt_text.as_deref(),
                            title: child.meta.title.as_deref(),
                            macro_name: child
                                .kind
                                .as_form_control()
                                .and_then(|control| control.macro_name.as_deref())
                                .map(duke_sheets_vml::encode_macro_formula),
                            control_text: child.kind.as_form_control().and_then(control_twin_text),
                            hidden: child.meta.hidden,
                            transform: &child.transform,
                            kind,
                        },
                    )
                })
                .collect();
            Some(PartKind::Group {
                transform: &group.transform,
                children,
            })
        }
        DrawingKind::Raw(raw) => Some(PartKind::Raw {
            bytes: &raw.bytes,
            rels: raw
                .rels
                .iter()
                .map(|rel| PartRel {
                    id: &rel.id,
                    rel_type: &rel.rel_type,
                    target: &rel.target,
                    external: rel.external,
                })
                .collect(),
        }),
    }
}

/// The sheet's drawing list in part-emission form.
fn part_objects(sheet: &Worksheet) -> Vec<PartObject<'_>> {
    let comment_count = sheet.comment_count();
    let mut control_ordinal = 0usize;
    sheet
        .drawings()
        .iter()
        .filter_map(|object| {
            convert_kind(
                &object.kind,
                &object.meta,
                comment_count,
                &mut control_ordinal,
            )
            .map(|kind| PartObject {
                name: object.meta.name.as_deref(),
                alt_text: object.meta.alt_text.as_deref(),
                title: object.meta.title.as_deref(),
                macro_name: object
                    .kind
                    .as_form_control()
                    .and_then(|control| control.macro_name.as_deref())
                    .map(duke_sheets_vml::encode_macro_formula),
                control_text: object.kind.as_form_control().and_then(control_twin_text),
                locked: object.meta.locked,
                printable: object.meta.printable,
                hidden: object.meta.hidden,
                anchor: &object.anchor,
                kind,
            })
        })
        .collect()
}

/// Assign relationship ids for the sheet's drawing part. Raw entries
/// keep their original ids; generated ids skip over them.
pub(super) fn plan_drawing_rels(
    sheet: &Worksheet,
    chart_globals: &[usize],
    chartex_globals: &[usize],
    image_parts: &[(usize, &'static str)],
) -> DrawingPlan {
    part_write::plan_drawing_rels(
        &part_objects(sheet),
        chart_globals,
        chartex_globals,
        image_parts,
    )
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
    let shape_base = (sheet_index + 1) * 1024 + 1 + sheet.comment_count();
    let bytes = part_write::write_drawing_part_with_metrics(
        &part_objects(sheet),
        plan,
        TwinStyle::CompatExtSp,
        shape_base,
        sheet,
    )?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
}

/// A chartsheet's drawing emission plan: anchor fragments in order,
/// the relationships captured for them, and a chart rel id allocated
/// around every id the anchors keep.
pub(super) struct ChartsheetDrawingPlan {
    pub anchors: Vec<Vec<u8>>,
    pub raw_rels: Vec<duke_sheets_core::RawRel>,
    pub chart_rid: String,
}

pub(super) fn plan_chartsheet_drawing(
    raw_drawing_objects: &[Vec<u8>],
    raw_drawing_rels: &[duke_sheets_core::RawRel],
) -> ChartsheetDrawingPlan {
    let anchors: Vec<Vec<u8>> = raw_drawing_objects.to_vec();
    let raw_rels: Vec<duke_sheets_core::RawRel> = raw_drawing_rels.to_vec();
    let mut taken: std::collections::HashSet<usize> = raw_rels
        .iter()
        .filter_map(|rel| rel.id.strip_prefix("rId").and_then(|n| n.parse().ok()))
        .collect();
    for anchor in &anchors {
        taken.extend(part_write::quoted_rel_id_nums(anchor));
    }
    let mut next = 1usize;
    while taken.contains(&next) {
        next += 1;
    }
    ChartsheetDrawingPlan {
        anchors,
        raw_rels,
        chart_rid: format!("rId{next}"),
    }
}

pub(super) fn write_chartsheet_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    plan: &ChartsheetDrawingPlan,
    has_chart: bool,
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    let bytes = part_write::write_chartsheet_drawing_part(
        has_chart.then_some(plan.chart_rid.as_str()),
        &plan.anchors,
    )?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
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
