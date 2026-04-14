use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::drawing_bundle::{is_drawing_bundle, DrawingBundle};
use crate::error::XlsbResult;

pub(crate) struct DrawingWriteResult {
    pub drawing_path: Option<String>,
    pub content_type_overrides: Vec<(String, String)>,
}

pub(crate) fn write_drawing_parts<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    raw_drawing_objects: &[Vec<u8>],
    sheet_index: usize,
) -> XlsbResult<DrawingWriteResult> {
    if raw_drawing_objects.is_empty() {
        return Ok(DrawingWriteResult {
            drawing_path: None,
            content_type_overrides: Vec::new(),
        });
    }

    let first = &raw_drawing_objects[0];
    if is_drawing_bundle(first) {
        write_from_bundle(zip, options, first)
    } else {
        write_from_anchors(zip, options, raw_drawing_objects, sheet_index)
    }
}

fn write_from_bundle<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    encoded: &[u8],
) -> XlsbResult<DrawingWriteResult> {
    let bundle = match DrawingBundle::decode(encoded) {
        Some(b) => b,
        None => {
            return Ok(DrawingWriteResult {
                drawing_path: None,
                content_type_overrides: Vec::new(),
            })
        }
    };

    let mut drawing_path = None;
    let mut overrides = Vec::new();

    for (path, _) in &bundle.entries {
        if path.ends_with(".rels") {
            continue;
        }

        let ct = content_type_for_path(path);
        if let Some(ct) = ct {
            overrides.push((format!("/{}", path), ct.to_string()));
        }

        if drawing_path.is_none() && path.contains("/drawings/") && path.ends_with(".xml") {
            drawing_path = Some(path.clone());
        }
    }

    for (path, data) in &bundle.entries {
        zip.start_file(path, *options)?;
        zip.write_all(data)?;
    }

    Ok(DrawingWriteResult {
        drawing_path,
        content_type_overrides: overrides,
    })
}

fn write_from_anchors<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    anchors: &[Vec<u8>],
    sheet_index: usize,
) -> XlsbResult<DrawingWriteResult> {
    let drawing_num = sheet_index + 1;
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);

    zip.start_file(&path, *options)?;
    zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")?;
    zip.write_all(
        b"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
    )?;
    for anchor in anchors {
        zip.write_all(anchor)?;
    }
    zip.write_all(b"</xdr:wsDr>")?;

    let ct = "application/vnd.openxmlformats-officedocument.drawing+xml".to_string();
    let overrides = vec![(format!("/{}", path), ct)];

    Ok(DrawingWriteResult {
        drawing_path: Some(path),
        content_type_overrides: overrides,
    })
}

fn content_type_for_path(path: &str) -> Option<&'static str> {
    if path.contains("/drawings/") && path.ends_with(".xml") && !path.contains("/_rels/") {
        Some("application/vnd.openxmlformats-officedocument.drawing+xml")
    } else if path.contains("/charts/chart") && path.ends_with(".xml") && !path.contains("Ex") {
        Some("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
    } else if path.contains("/charts/chartEx") && path.ends_with(".xml") {
        Some("application/vnd.ms-office.chartEx+xml")
    } else if path.contains("/charts/style") && path.ends_with(".xml") {
        Some("application/vnd.ms-office.chartstyle+xml")
    } else if path.contains("/charts/colors") && path.ends_with(".xml") {
        Some("application/vnd.ms-office.chartcolorstyle+xml")
    } else {
        None
    }
}
