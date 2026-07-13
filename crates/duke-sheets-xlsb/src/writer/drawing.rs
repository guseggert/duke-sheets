//! Model-driven drawing emission: drawing part XML (shared codec,
//! com14:compatSp twins), drawing rels, chart parts, media parts,
//! and their content types.

use std::collections::HashSet;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use duke_sheets_chart::drawing_part::image_format_extension;
use duke_sheets_chart::drawing_part::write::{
    self as part_write, PartChild, PartKind, PartObject, PartRel, PartShape, TwinStyle,
};
use duke_sheets_chart::drawing_part::{ShapeFill as PartShapeFill, ShapeLine as PartShapeLine};
use duke_sheets_core::{
    DrawingKind, DrawingMeta, RawRel, Shape, ShapeFill, ShapeGeometry, Worksheet,
};

use crate::error::XlsbResult;

const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_CHART_EX: &str = "application/vnd.ms-office.chartex+xml";
const CT_CHART_STYLE: &str = "application/vnd.ms-office.chartstyle+xml";
const CT_CHART_COLOR_STYLE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const RT_CHART_STYLE: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const RT_CHART_COLOR_STYLE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

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

pub(crate) fn is_unsupported(chart: &duke_sheets_chart::Chart) -> bool {
    matches!(
        chart.chart_type,
        duke_sheets_chart::ChartType::Unsupported(_)
    )
}

/// Whether a worksheet needs a drawing part: any drawing object with
/// a native-part presence. Comments live only in the legacy VML part;
/// unsupported charts are skipped by the writer.
pub(crate) fn sheet_has_drawing_content(sheet: &Worksheet) -> bool {
    sheet.drawings().iter().any(|object| match &object.kind {
        DrawingKind::Comment { .. } => false,
        DrawingKind::Chart(chart) => !is_unsupported(chart),
        _ => true,
    })
}

/// Every image payload the drawing part will reference, depth-first
/// in drawing-list order (group children included). Feeds media part
/// numbering and content types; must match the emission walk.
pub(crate) fn sheet_image_payloads(sheet: &Worksheet) -> Vec<&duke_sheets_chart::EmbeddedImage> {
    fn walk<'a>(kind: &'a DrawingKind, out: &mut Vec<&'a duke_sheets_chart::EmbeddedImage>) {
        match kind {
            DrawingKind::Image(image) => out.push(image),
            DrawingKind::Group(group) => {
                for child in &group.children {
                    walk(&child.kind, out);
                }
            }
            _ => {}
        }
    }
    let mut out = Vec::new();
    for object in sheet.drawings() {
        walk(&object.kind, &mut out);
    }
    out
}

/// Supported charts on the sheet, in drawing-list order.
pub(crate) fn sheet_charts(sheet: &Worksheet) -> Vec<&duke_sheets_chart::Chart> {
    sheet
        .charts()
        .filter(|drawn| !is_unsupported(drawn.payload))
        .map(|drawn| drawn.payload)
        .collect()
}

/// Raw relationships preserved on a worksheet's raw drawing entries
/// (including raw group children), in list order, deduplicated by
/// target. Conflicting reuses of one relationship id across fragments
/// get distinct ids at plan time, so every distinct target's part
/// must be collected.
pub(crate) fn sheet_raw_rels(sheet: &Worksheet) -> Vec<&RawRel> {
    fn collect<'a>(kind: &'a DrawingKind, seen: &mut HashSet<&'a str>, out: &mut Vec<&'a RawRel>) {
        match kind {
            DrawingKind::Raw(raw) => {
                for rel in &raw.rels {
                    if seen.insert(rel.target.as_str()) {
                        out.push(rel);
                    }
                }
            }
            DrawingKind::Group(group) => {
                for child in &group.children {
                    collect(&child.kind, seen, out);
                }
            }
            _ => {}
        }
    }
    let mut seen = HashSet::new();
    let mut out = Vec::new();
    for object in sheet.drawings() {
        collect(&object.kind, &mut seen, &mut out);
    }
    out
}

/// Resolve a relationship target against a base directory (e.g.
/// "xl/drawings"): `../media/image5.png` -> `xl/media/image5.png`.
pub(crate) fn resolve_rel_target(base_dir: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for part in target.split('/') {
        if part == ".." {
            parts.pop();
        } else if part != "." && !part.is_empty() {
            parts.push(part);
        }
    }
    parts.join("/")
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
                duke_sheets_vml::default_control_name(
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
                            macro_name: None,
                            control_text: None,
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
                macro_name: None,
                control_text: None,
                locked: object.meta.locked,
                printable: object.meta.printable,
                hidden: object.meta.hidden,
                anchor: &object.anchor,
                kind,
            })
        })
        .collect()
}

pub(crate) struct DrawingWriteResult {
    pub drawing_path: Option<String>,
    /// `(part_name, content_type)` overrides for `[Content_Types].xml`.
    pub content_type_overrides: Vec<(String, String)>,
    /// Media extensions used by written media parts (Defaults).
    pub media_default_exts: Vec<String>,
}

/// Per-sheet global part numbering, assigned by the caller.
pub(crate) struct DrawingNumbering {
    pub drawing_num: usize,
    /// Global number of this sheet's first chart part.
    pub chart_start: usize,
    /// Global number of this sheet's first chartEx part.
    pub chartex_start: usize,
    /// Global number of this sheet's first media part.
    pub image_start: usize,
}

/// Write every drawing-related part for one worksheet: the drawing
/// XML (list order, com14:compatSp control twins), its rels, chart /
/// chartEx parts, media parts, and raw-preserved parts.
pub(crate) fn write_drawing_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    ws: &Worksheet,
    sheet_index: usize,
    numbering: &DrawingNumbering,
    written_media: &mut HashSet<String>,
) -> XlsbResult<DrawingWriteResult> {
    let mut overrides: Vec<(String, String)> = Vec::new();
    let mut media_exts: Vec<String> = Vec::new();

    let charts = sheet_charts(ws);
    let charts_ex: Vec<_> = ws.charts_ex().map(|drawn| drawn.payload).collect();
    let image_payloads = sheet_image_payloads(ws);

    let chart_globals: Vec<usize> = (0..charts.len())
        .map(|j| numbering.chart_start + j)
        .collect();
    let chartex_globals: Vec<usize> = (0..charts_ex.len())
        .map(|j| numbering.chartex_start + j)
        .collect();
    let image_parts: Vec<(usize, &'static str)> = image_payloads
        .iter()
        .enumerate()
        .map(|(j, img)| {
            (
                numbering.image_start + j,
                image_format_extension(img.format),
            )
        })
        .collect();

    let objects = part_objects(ws);
    let plan =
        part_write::plan_drawing_rels(&objects, &chart_globals, &chartex_globals, &image_parts);
    let shape_base = (sheet_index + 1) * 1024 + 1 + ws.comment_count();
    let drawing_bytes = part_write::write_drawing_part_with_metrics(
        &objects,
        &plan,
        TwinStyle::CompatSpFrame,
        shape_base,
        ws,
    )?;

    let drawing_path = format!("xl/drawings/drawing{}.xml", numbering.drawing_num);
    zip.start_file(&drawing_path, *options)?;
    zip.write_all(&drawing_bytes)?;
    overrides.push((format!("/{}", drawing_path), CT_DRAWING.to_string()));

    let rels_bytes = part_write::write_rels_part(&plan.rels)?;
    zip.start_file(
        format!(
            "xl/drawings/_rels/drawing{}.xml.rels",
            numbering.drawing_num
        ),
        *options,
    )?;
    zip.write_all(&rels_bytes)?;

    // Chart parts (XML content, same as XLSX).
    for (chart, &gn) in charts.iter().zip(&chart_globals) {
        let path = format!("xl/charts/chart{}.xml", gn);
        let bytes = duke_sheets_chart::write::chart_part_bytes(chart)?;
        zip.start_file(&path, *options)?;
        zip.write_all(&bytes)?;
        overrides.push((format!("/{}", path), CT_CHART.to_string()));
        write_chart_style_color_parts(zip, options, chart, gn, &mut overrides)?;
    }
    for (chart_ex, &gn) in charts_ex.iter().zip(&chartex_globals) {
        let path = format!("xl/charts/chartEx{}.xml", gn);
        let bytes = duke_sheets_chart::write::chart_ex_part_bytes(chart_ex)?;
        zip.start_file(&path, *options)?;
        zip.write_all(&bytes)?;
        overrides.push((format!("/{}", path), CT_CHART_EX.to_string()));
        write_chart_ex_style_color_parts(zip, options, chart_ex, gn, &mut overrides)?;
    }

    // Media parts (xl/media/imageN.<ext>).
    for (img, &(gn, ext)) in image_payloads.iter().zip(&image_parts) {
        let path = format!("xl/media/image{gn}.{ext}");
        if written_media.insert(path.clone()) {
            zip.start_file(&path, *options)?;
            zip.write_all(&img.data)?;
        }
        media_exts.push(ext.to_string());
    }

    // Raw-preserved parts at their original paths.
    for rel in sheet_raw_rels(ws) {
        let Some(part) = rel.part.as_deref() else {
            continue;
        };
        if rel.external {
            continue;
        }
        let path = resolve_rel_target("xl/drawings", &rel.target);
        if !written_media.insert(path.clone()) {
            continue;
        }
        zip.start_file(&path, *options)?;
        zip.write_all(part)?;
        if let Some(ct) = content_type_for_path(&path) {
            overrides.push((format!("/{}", path), ct.to_string()));
        } else if path.starts_with("xl/media/") {
            if let Some(ext) = path.rsplit('.').next() {
                media_exts.push(ext.to_ascii_lowercase());
            }
        }
    }

    Ok(DrawingWriteResult {
        drawing_path: Some(drawing_path),
        content_type_overrides: overrides,
        media_default_exts: media_exts,
    })
}

fn write_chart_style_color_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    chart: &duke_sheets_chart::Chart,
    chart_num: usize,
    overrides: &mut Vec<(String, String)>,
) -> XlsbResult<()> {
    let has_style = chart.raw_chart_style.is_some();
    let has_color = chart.raw_chart_color_style.is_some();
    if !has_style && !has_color {
        return Ok(());
    }
    let mut rel_id = 1u32;
    let mut rels_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    if let Some(ref bytes) = chart.raw_chart_style {
        let style_path = format!("xl/charts/style{}.xml", chart_num);
        zip.start_file(&style_path, *options)?;
        zip.write_all(bytes)?;
        overrides.push((format!("/{}", style_path), CT_CHART_STYLE.to_string()));
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="style{}.xml"/>"#,
            rel_id, RT_CHART_STYLE, chart_num
        ));
        rel_id += 1;
    }
    if let Some(ref bytes) = chart.raw_chart_color_style {
        let color_path = format!("xl/charts/colors{}.xml", chart_num);
        zip.start_file(&color_path, *options)?;
        zip.write_all(bytes)?;
        overrides.push((format!("/{}", color_path), CT_CHART_COLOR_STYLE.to_string()));
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="colors{}.xml"/>"#,
            rel_id, RT_CHART_COLOR_STYLE, chart_num
        ));
    }
    rels_xml.push_str("</Relationships>");
    let rels_path = format!("xl/charts/_rels/chart{}.xml.rels", chart_num);
    zip.start_file(&rels_path, *options)?;
    zip.write_all(rels_xml.as_bytes())?;
    Ok(())
}

fn write_chart_ex_style_color_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    chart_ex: &duke_sheets_chart::ChartEx,
    chart_ex_num: usize,
    overrides: &mut Vec<(String, String)>,
) -> XlsbResult<()> {
    let has_style = chart_ex.raw_chart_style.is_some();
    let has_color = chart_ex.raw_chart_color_style.is_some();
    if !has_style && !has_color {
        return Ok(());
    }
    let mut rel_id = 1u32;
    let mut rels_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    if let Some(ref bytes) = chart_ex.raw_chart_style {
        let style_path = format!("xl/charts/style{}.xml", chart_ex_num);
        zip.start_file(&style_path, *options)?;
        zip.write_all(bytes)?;
        overrides.push((format!("/{}", style_path), CT_CHART_STYLE.to_string()));
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="style{}.xml"/>"#,
            rel_id, RT_CHART_STYLE, chart_ex_num
        ));
        rel_id += 1;
    }
    if let Some(ref bytes) = chart_ex.raw_chart_color_style {
        let color_path = format!("xl/charts/colors{}.xml", chart_ex_num);
        zip.start_file(&color_path, *options)?;
        zip.write_all(bytes)?;
        overrides.push((format!("/{}", color_path), CT_CHART_COLOR_STYLE.to_string()));
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="colors{}.xml"/>"#,
            rel_id, RT_CHART_COLOR_STYLE, chart_ex_num
        ));
    }
    rels_xml.push_str("</Relationships>");
    let rels_path = format!("xl/charts/_rels/chartEx{}.xml.rels", chart_ex_num);
    zip.start_file(&rels_path, *options)?;
    zip.write_all(rels_xml.as_bytes())?;
    Ok(())
}

/// Content type inferred from a raw-preserved part path.
fn content_type_for_path(path: &str) -> Option<&'static str> {
    if path.contains("/drawings/") && path.ends_with(".xml") && !path.contains("/_rels/") {
        Some(CT_DRAWING)
    } else if path.contains("/charts/chartEx") && path.ends_with(".xml") {
        Some(CT_CHART_EX)
    } else if path.contains("/charts/chart") && path.ends_with(".xml") {
        Some(CT_CHART)
    } else if path.contains("/charts/style") && path.ends_with(".xml") {
        Some(CT_CHART_STYLE)
    } else if path.contains("/charts/colors") && path.ends_with(".xml") {
        Some(CT_CHART_COLOR_STYLE)
    } else {
        None
    }
}
