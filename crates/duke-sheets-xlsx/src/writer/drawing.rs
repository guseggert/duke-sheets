use std::io::{Seek, Write};

use duke_sheets_chart::drawing_part::write::{
    self as part_write, PartChild, PartKind, PartObject, PartRel, TwinStyle,
};
use duke_sheets_chart::Chart;
use duke_sheets_core::{DrawingKind, DrawingMeta, Worksheet};

use super::{XlsxResult, RT_CHART};

pub(super) use duke_sheets_chart::drawing_part::write::{DrawingPlan, PlannedRel};
pub(super) use duke_sheets_chart::drawing_part::{image_format_extension, image_format_mime};

pub(super) fn is_unsupported(chart: &Chart) -> bool {
    matches!(chart.chart_type, duke_sheets_chart::ChartType::Unsupported(_))
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
                locked: object.meta.locked,
                printable: object.meta.printable,
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
    let bytes = part_write::write_drawing_part(
        &part_objects(sheet),
        plan,
        TwinStyle::CompatExtSp,
        shape_base,
    )?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
}

/// Write a drawing part's .rels from its plan.
pub(super) fn write_drawing_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    drawing_num: usize,
    rels: &[PlannedRel],
) -> XlsxResult<()> {
    let path = format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_num);
    let bytes = part_write::write_rels_part(rels)?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
}

pub(super) fn write_chartsheet_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    _chart: &Chart,
    raw_drawing_objects: &[Vec<u8>],
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    let bytes = part_write::write_chartsheet_drawing_part(raw_drawing_objects)?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
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
