//! Assemble a worksheet's drawing objects from the drawing part XML
//! (shared codec), the legacy VML part (control state), and the
//! chart parts.

use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};

use duke_sheets_chart::drawing_part::read::{
    parse_drawing_part, DrawingEntryKind, ParsedChild, ParsedGroup, ParsedShape, PicShape,
};
use duke_sheets_core::{
    ChildTransform, DrawingKind, DrawingMeta, DrawingObject, Group, GroupChild, RawDrawing, RawRel,
};
use quick_xml::events::Event;

use crate::error::XlsbResult;

use super::SheetRel;

/// A VML control shape waiting to be matched to its drawing-part
/// twin. XLSB has no ctrlProps part or worksheet controls block:
/// state comes from the VML ClientData alone, the twin carries the
/// name and z-position.
struct AssembledControl {
    shape_id: u32,
    object: DrawingObject,
    /// True when the VML shape carried no x:Anchor, so a matched
    /// twin's anchor is the fallback.
    anchor_defaulted: bool,
    consumed: bool,
}

/// Consume the first unconsumed control matching a twin's shape id.
fn take_control(
    controls: &mut [AssembledControl],
    shape_num: Option<u32>,
) -> Option<(u32, DrawingObject, bool)> {
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

/// Assemble the sheet's native drawing list: drawing-part entries in
/// document order (control twins replaced by their matched VML
/// controls), then unmatched controls in VML order. Each object is
/// paired with its legacy shape id when it wraps a control, for
/// comment splicing.
pub(crate) fn merge_sheet_drawings<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
    sheet_rels: &HashMap<String, SheetRel>,
    vml_shapes: &[duke_sheets_vml::VmlShape],
) -> XlsbResult<Vec<(DrawingObject, Option<u32>)>> {
    let mut controls: Vec<AssembledControl> = Vec::new();
    for shape in vml_shapes {
        let duke_sheets_vml::VmlShapeKind::Control(control) = &shape.kind else {
            continue;
        };
        if let Some(object) = control.to_drawing_object() {
            controls.push(AssembledControl {
                shape_id: shape.shape_num,
                object,
                anchor_defaulted: control.anchor_px.is_none(),
                consumed: false,
            });
        }
    }

    let mut natives: Vec<(DrawingObject, Option<u32>)> = Vec::new();
    let mut drawing_targets: Vec<(String, String)> = sheet_rels
        .iter()
        .filter(|(_, r)| r.rel_type.ends_with("/drawing"))
        .map(|(id, r)| (id.clone(), r.target.clone()))
        .collect();
    // Numeric-aware sort so rId2 precedes rId10.
    drawing_targets.sort_by_key(|(id, _)| {
        (
            id.strip_prefix("rId")
                .and_then(|n| n.parse::<u64>().ok())
                .unwrap_or(u64::MAX),
            id.clone(),
        )
    });

    for (_, target) in &drawing_targets {
        let drawing_path = super::resolve_rel_path(sheet_path, target);
        let Some(bytes) = read_zip_entry(archive, &drawing_path) else {
            continue;
        };
        let entries = match parse_drawing_part(&bytes) {
            Ok(entries) => entries,
            Err(e) => {
                log::warn!("failed to parse drawing part {drawing_path}: {e}");
                continue;
            }
        };
        if entries.is_empty() {
            continue;
        }
        let drawing_rels = super::read_sheet_rels(archive, &drawing_path)?;

        for entry in entries {
            match entry.kind {
                DrawingEntryKind::Image(pic) => {
                    let pic = *pic;
                    let name = pic.name.clone();
                    let descr = pic.descr.clone();
                    let title = pic.title.clone();
                    let hidden = pic.hidden;
                    let image =
                        resolve_pic_image(archive, &drawing_path, &drawing_rels, pic, false);
                    let mut object = DrawingObject::image(image).with_anchor(entry.anchor);
                    object.meta.name = Some(name);
                    object.meta.alt_text = descr;
                    object.meta.title = title;
                    object.meta.hidden = hidden;
                    object.meta.locked = entry.locked;
                    object.meta.printable = entry.printable;
                    natives.push((object, None));
                }
                DrawingEntryKind::Shape(shape) => {
                    let mut object =
                        DrawingObject::shape(shape_from_parsed(&shape)).with_anchor(entry.anchor);
                    object.meta.name = Some(shape.name.clone());
                    object.meta.alt_text = shape.descr.clone();
                    object.meta.title = shape.title.clone();
                    object.meta.hidden = shape.hidden;
                    object.meta.locked = entry.locked;
                    object.meta.printable = entry.printable;
                    natives.push((object, None));
                }
                DrawingEntryKind::Chart(chart_ref) => {
                    let Some(rel) = drawing_rels.get(&chart_ref.rel_id) else {
                        continue;
                    };
                    let chart_path = super::resolve_rel_path(&drawing_path, &rel.target);
                    let Some(chart_bytes) = read_zip_entry(archive, &chart_path) else {
                        continue;
                    };
                    if chart_ref.is_chart_ex {
                        if let Ok(mut cx) =
                            duke_sheets_chart::parse::parse_chart_ex_xml(Cursor::new(&chart_bytes))
                        {
                            cx.raw_mc_fallback = chart_ref.raw_mc_fallback;
                            let (style, color) = read_chart_style_color(archive, &chart_path);
                            cx.raw_chart_style = style;
                            cx.raw_chart_color_style = color;
                            let mut object =
                                DrawingObject::chart_ex(cx).with_anchor(chart_ref.anchor);
                            object.meta.name = chart_ref.name;
                            object.meta.alt_text = chart_ref.descr;
                            object.meta.title = chart_ref.title;
                            object.meta.hidden = chart_ref.hidden;
                            object.meta.locked = entry.locked;
                            object.meta.printable = entry.printable;
                            natives.push((object, None));
                        }
                    } else if let Ok(mut c) =
                        duke_sheets_chart::parse::parse_chart_xml(Cursor::new(&chart_bytes))
                    {
                        let (style, color) = read_chart_style_color(archive, &chart_path);
                        c.raw_chart_style = style;
                        c.raw_chart_color_style = color;
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
                DrawingEntryKind::Group(group) => {
                    let name = group.name.clone();
                    let descr = group.descr.clone();
                    let title = group.title.clone();
                    let hidden = group.hidden;
                    let built =
                        build_group(archive, &drawing_path, &drawing_rels, group, &mut controls);
                    let mut object = DrawingObject::group(built).with_anchor(entry.anchor);
                    object.meta.name = Some(name);
                    object.meta.alt_text = descr;
                    object.meta.title = title;
                    object.meta.hidden = hidden;
                    object.meta.locked = entry.locked;
                    object.meta.printable = entry.printable;
                    natives.push((object, None));
                }
                DrawingEntryKind::ControlTwin(twin) => {
                    // The twin is a placeholder for the matched VML
                    // control; unmatched twins are dropped.
                    if let Some((shape_id, mut object, anchor_defaulted)) =
                        take_control(&mut controls, twin.shape_num)
                    {
                        if let Some(name) = twin.name {
                            object.meta.name = Some(name);
                        }
                        if twin.descr.is_some() {
                            object.meta.alt_text = twin.descr;
                        }
                        object.meta.title = twin.title;
                        if anchor_defaulted {
                            object.anchor = entry.anchor;
                        }
                        natives.push((object, Some(shape_id)));
                    }
                }
                DrawingEntryKind::Raw => {
                    let rels =
                        capture_raw_rels(archive, &drawing_path, &entry.bytes, &drawing_rels);
                    let object = DrawingObject::raw(RawDrawing {
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
    // native entries, in VML order.
    for control in controls {
        if !control.consumed {
            natives.push((control.object, Some(control.shape_id)));
        }
    }

    Ok(natives)
}

fn read_zip_entry<R: Read + Seek>(archive: &mut zip::ZipArchive<R>, path: &str) -> Option<Vec<u8>> {
    let mut file = archive.by_name(path).ok()?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).ok()?;
    Some(bytes)
}

/// Capture the chartStyle / chartColorStyle parts referenced by a
/// chart part's rels, for round-trip.
fn read_chart_style_color<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    chart_path: &str,
) -> (Option<Vec<u8>>, Option<Vec<u8>>) {
    let Ok(chart_rels) = super::read_sheet_rels(archive, chart_path) else {
        return (None, None);
    };
    let mut style = None;
    let mut color = None;
    for rel in chart_rels.values() {
        let path = super::resolve_rel_path(chart_path, &rel.target);
        if rel.rel_type.ends_with("/chartStyle") {
            style = read_zip_entry(archive, &path);
        } else if rel.rel_type.ends_with("/chartColorStyle") {
            color = read_zip_entry(archive, &path);
        }
    }
    (style, color)
}

/// Build an `EmbeddedImage` from a parsed pic, resolving the blip
/// relationship to media bytes. For group children (`in_group`) the
/// child transform carries rotation/flips, so the payload keeps none.
fn resolve_pic_image<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
    drawing_rels: &HashMap<String, SheetRel>,
    pic: PicShape,
    in_group: bool,
) -> duke_sheets_chart::EmbeddedImage {
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
    if let Some(rel) = drawing_rels.get(&image.media_path) {
        let path = super::resolve_rel_path(drawing_path, &rel.target);
        let ext = path.rsplit('.').next().unwrap_or("");
        if let Some(fmt) = duke_sheets_chart::ImageFormat::from_extension(ext) {
            image.format = fmt;
        }
        image.media_path = path.clone();
        if let Some(bytes) = read_zip_entry(archive, &path) {
            image.data = bytes;
        }
    }
    if let Some(svg_rel_id) = &image.svg_media_path {
        if let Some(rel) = drawing_rels.get(svg_rel_id.as_str()) {
            let path = super::resolve_rel_path(drawing_path, &rel.target);
            image.svg_media_path = Some(path.clone());
            if let Some(bytes) = read_zip_entry(archive, &path) {
                image.svg_data = Some(bytes);
            }
        }
    }
    image
}

fn shape_from_parsed(parsed: &ParsedShape) -> duke_sheets_core::Shape {
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
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
    drawing_rels: &HashMap<String, SheetRel>,
    group: ParsedGroup,
    controls: &mut [AssembledControl],
) -> Group {
    let mut children = Vec::new();
    for child in group.children {
        match child {
            ParsedChild::Pic(pic) => {
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
                let image = resolve_pic_image(archive, drawing_path, drawing_rels, pic, true);
                children.push(GroupChild {
                    meta,
                    transform,
                    kind: DrawingKind::Image(image),
                });
            }
            ParsedChild::Shape(shape) => {
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
            ParsedChild::Group(inner) => {
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
                let built = build_group(archive, drawing_path, drawing_rels, inner, controls);
                children.push(GroupChild {
                    meta,
                    transform,
                    kind: DrawingKind::Group(Box::new(built)),
                });
            }
            ParsedChild::Twin(twin) => {
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
                    children.push(GroupChild {
                        meta: object.meta,
                        transform: twin.xfrm,
                        kind: object.kind,
                    });
                }
            }
        }
    }
    Group {
        transform: group.transform,
        children,
    }
}

/// Scan a raw anchor's bytes for relationship references (r:id,
/// r:embed, r:link) and capture each referenced relationship with its
/// original id/target plus the target part bytes when internal.
fn capture_raw_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
    bytes: &[u8],
    drawing_rels: &HashMap<String, SheetRel>,
) -> Vec<RawRel> {
    let mut ids: Vec<String> = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                for attr in e.attributes().flatten() {
                    if matches!(attr.key.as_ref(), b"r:id" | b"r:embed" | b"r:link") {
                        if let Ok(value) = attr.unescape_value() {
                            let value = value.to_string();
                            if !ids.contains(&value) {
                                ids.push(value);
                            }
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
        let Some(rel) = drawing_rels.get(&id) else {
            continue;
        };
        let part = if rel.external {
            None
        } else {
            let path = super::resolve_rel_path(drawing_path, &rel.target);
            read_zip_entry(archive, &path)
        };
        rels.push(RawRel {
            id,
            rel_type: rel.rel_type.clone(),
            target: rel.target.clone(),
            external: rel.external,
            part,
        });
    }
    rels
}
