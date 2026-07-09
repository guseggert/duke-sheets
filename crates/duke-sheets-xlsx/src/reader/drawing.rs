use std::io::{BufReader, Cursor, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

/// A chart reference discovered in a drawing XML, paired with its anchor position.
pub(crate) struct DrawingChartRef {
    /// The relationship id (e.g. "rId1") pointing to the chart part.
    pub(crate) rel_id: String,
    /// The two-cell anchor positioning the chart in the worksheet.
    pub(crate) anchor: DrawingAnchor,
    /// Whether this references a ChartEx part (`cx:chart`) rather than a standard chart.
    pub(crate) is_chart_ex: bool,
    /// Raw `mc:Fallback` XML bytes for roundtrip (chartEx only).
    pub(crate) raw_mc_fallback: Option<Vec<u8>>,
}

/// Chart refs, image refs, and raw non-chart drawing anchors from a drawing XML.
pub(crate) struct DrawingContents {
    pub chart_refs: Vec<DrawingChartRef>,
    pub images: Vec<EmbeddedImage>,
    pub raw_non_chart_anchors: Vec<Vec<u8>>,
}

const URI_STANDARD_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const URI_CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

/// Parse a SpreadsheetML drawing and return chart refs plus raw non-chart anchors.
///
/// For each anchor (twoCellAnchor, oneCellAnchor, absoluteAnchor):
/// - If it contains a graphicFrame with a chart URI → extract a `DrawingChartRef`
/// - Otherwise → capture the entire anchor element as raw XML bytes
pub(crate) fn read_drawing_contents<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<DrawingContents> {
    let file = match archive_by_name(archive, drawing_path) {
        Ok(f) => f,
        Err(_) => {
            return Ok(DrawingContents {
                chart_refs: Vec::new(),
                images: Vec::new(),
                raw_non_chart_anchors: Vec::new(),
            })
        }
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut chart_refs = Vec::new();
    let mut images: Vec<EmbeddedImage> = Vec::new();
    let mut raw_non_chart_anchors: Vec<Vec<u8>> = Vec::new();

    let mut in_two_cell_anchor = false;
    let mut in_one_cell_anchor = false;
    let mut in_absolute_anchor = false;
    let mut from = CellMarker::default();
    let mut to = CellMarker::default();
    let mut in_from = false;
    let mut in_to = false;
    let mut in_col = false;
    let mut in_col_off = false;
    let mut in_row = false;
    let mut in_row_off = false;
    let mut in_graphic_data = false;
    let mut graphic_data_uri: Option<String> = None;
    let mut chart_rel_id: Option<String> = None;
    let mut is_chart_ex = false;
    let mut capture: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut fallback_capture: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut fallback_depth: u32 = 0;
    let mut raw_mc_fallback: Option<Vec<u8>> = None;
    // Set when the captured anchor contains an a14:compatExt marker:
    // it is the drawing twin of a legacy form control (parsed from
    // ctrlProps/VML instead) and must not be kept as a raw anchor.
    let mut saw_compat_ext = false;

    // Image (pic) parsing state
    let mut in_pic = false;
    let mut _in_grp_sp = false;
    let mut pic_id: u32 = 0;
    let mut pic_name = String::new();
    let mut pic_descr: Option<String> = None;
    let mut blip_rel_id: Option<String> = None;
    let mut svg_blip_rel_id: Option<String> = None;
    let mut pic_width_emu: i64 = 0;
    let mut pic_height_emu: i64 = 0;
    let mut pic_rotation: Option<i32> = None;
    let mut pic_flip_h = false;
    let mut pic_flip_v = false;
    let mut in_sp_pr = false;
    let mut sp_pr_depth: u32 = 0;

    // Anchor-variant-specific state. For OneCell, the anchor's
    // <xdr:ext> carries the picture extent (width × height in EMU).
    // For Absolute, the anchor's <xdr:pos> carries x/y and
    // <xdr:ext> carries width × height. The xfrm-level ext inside
    // spPr is captured separately in pic_width_emu/pic_height_emu.
    let mut anchor_ext_cx: i64 = 0;
    let mut anchor_ext_cy: i64 = 0;
    let mut anchor_pos_x: i64 = 0;
    let mut anchor_pos_y: i64 = 0;
    let mut twocell_edit_as: Option<duke_sheets_chart::EditAs> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let in_any_anchor = in_two_cell_anchor || in_one_cell_anchor || in_absolute_anchor;

                // Capture mc:Fallback content
                if let Some(ref mut w) = fallback_capture {
                    let _ = w.write_event(Event::Start(e.clone().into_owned()));
                    fallback_depth += 1;
                } else if in_any_anchor {
                    if let Some(ref mut w) = capture {
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                    }
                }

                match e.name().local_name().as_ref() {
                    b"twoCellAnchor" => {
                        in_two_cell_anchor = true;
                        from = CellMarker::default();
                        to = CellMarker::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        twocell_edit_as = None;
                        // Parse the optional editAs="..." attribute.
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"editAs" {
                                if let Ok(s) = attr.unescape_value() {
                                    twocell_edit_as = match s.as_ref() {
                                        "twoCell" => Some(duke_sheets_chart::EditAs::TwoCell),
                                        "oneCell" => Some(duke_sheets_chart::EditAs::OneCell),
                                        "absolute" => Some(duke_sheets_chart::EditAs::Absolute),
                                        _ => None,
                                    };
                                }
                            }
                        }
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"oneCellAnchor" => {
                        in_one_cell_anchor = true;
                        from = CellMarker::default();
                        to = CellMarker::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        anchor_ext_cx = 0;
                        anchor_ext_cy = 0;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"absoluteAnchor" => {
                        in_absolute_anchor = true;
                        from = CellMarker::default();
                        to = CellMarker::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        anchor_ext_cx = 0;
                        anchor_ext_cy = 0;
                        anchor_pos_x = 0;
                        anchor_pos_y = 0;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"from" if in_two_cell_anchor || in_one_cell_anchor => in_from = true,
                    b"to" if in_two_cell_anchor => in_to = true,
                    b"col" if in_from || in_to => in_col = true,
                    b"colOff" if in_from || in_to => in_col_off = true,
                    b"row" if in_from || in_to => in_row = true,
                    b"rowOff" if in_from || in_to => in_row_off = true,
                    b"graphicData" if in_any_anchor => {
                        in_graphic_data = true;
                        graphic_data_uri = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"uri" {
                                graphic_data_uri =
                                    attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"Fallback" if in_any_anchor && fallback_capture.is_none() && is_chart_ex => {
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        fallback_depth = 1;
                        fallback_capture = Some(w);
                    }
                    b"pic" if in_any_anchor => {
                        in_pic = true;
                        pic_id = 0;
                        pic_name = String::new();
                        pic_descr = None;
                        blip_rel_id = None;
                        svg_blip_rel_id = None;
                        pic_width_emu = 0;
                        pic_height_emu = 0;
                        pic_rotation = None;
                        pic_flip_h = false;
                        pic_flip_v = false;
                        in_sp_pr = false;
                        sp_pr_depth = 0;
                    }
                    b"grpSp" if in_any_anchor => _in_grp_sp = true,
                    b"cNvPr" if in_pic => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"id" => {
                                    pic_id = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"name" => {
                                    pic_name = attr
                                        .unescape_value()
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();
                                }
                                b"descr" => {
                                    pic_descr = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                _ => {}
                            }
                        }
                    }
                    b"blip" if in_pic => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"embed" {
                                blip_rel_id = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"spPr" if in_pic => {
                        in_sp_pr = true;
                        sp_pr_depth = 1;
                    }
                    b"xfrm" if in_sp_pr => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"rot" => {
                                    pic_rotation =
                                        attr.unescape_value().ok().and_then(|s| s.parse().ok());
                                }
                                b"flipH" => {
                                    pic_flip_h = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s == "true")
                                        .unwrap_or(false);
                                }
                                b"flipV" => {
                                    pic_flip_v = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s == "true")
                                        .unwrap_or(false);
                                }
                                _ => {}
                            }
                        }
                    }
                    _ if in_sp_pr => sp_pr_depth += 1,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(ref mut w) = fallback_capture {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                } else if let Some(ref mut w) = capture {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                }
                if capture.is_some() && e.name().local_name().as_ref() == b"compatExt" {
                    saw_compat_ext = true;
                }
                let local = e.name().local_name();
                if in_graphic_data && local.as_ref() == b"chart" {
                    let uri = graphic_data_uri.as_deref().unwrap_or("");
                    match uri {
                        URI_CHART_EX => is_chart_ex = true,
                        URI_STANDARD_CHART | _ => is_chart_ex = uri == URI_CHART_EX,
                    }
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"id" {
                            chart_rel_id = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                    }
                }
                // Anchor-level <xdr:ext> and <xdr:pos> live OUTSIDE
                // the <xdr:pic>/<xdr:spPr> nesting. They carry the
                // picture extent (OneCell / Absolute) and origin
                // (Absolute only). Both are self-closing tags.
                let in_any_anchor = in_two_cell_anchor || in_one_cell_anchor || in_absolute_anchor;
                if in_any_anchor && !in_pic {
                    match local.as_ref() {
                        b"ext" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"cx" => {
                                        anchor_ext_cx = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    b"cy" => {
                                        anchor_ext_cy = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"pos" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"x" => {
                                        anchor_pos_x = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    b"y" => {
                                        anchor_pos_y = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if in_pic {
                    match local.as_ref() {
                        b"cNvPr" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"id" => {
                                        pic_id = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    b"name" => {
                                        pic_name = attr
                                            .unescape_value()
                                            .map(|s| s.to_string())
                                            .unwrap_or_default();
                                    }
                                    b"descr" => {
                                        pic_descr =
                                            attr.unescape_value().ok().map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"blip" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"embed" {
                                    blip_rel_id = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                            }
                        }
                        b"svgBlip" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"embed" {
                                    svg_blip_rel_id =
                                        attr.unescape_value().ok().map(|s| s.to_string());
                                }
                            }
                        }
                        b"ext" if in_sp_pr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"cx" => {
                                        pic_width_emu = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    b"cy" => {
                                        pic_height_emu = attr
                                            .unescape_value()
                                            .ok()
                                            .and_then(|s| s.parse().ok())
                                            .unwrap_or(0);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        b"xfrm" if in_sp_pr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"rot" => {
                                        pic_rotation =
                                            attr.unescape_value().ok().and_then(|s| s.parse().ok());
                                    }
                                    b"flipH" => {
                                        pic_flip_h = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|s| s == "1" || s == "true")
                                            .unwrap_or(false);
                                    }
                                    b"flipV" => {
                                        pic_flip_v = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|s| s == "1" || s == "true")
                                            .unwrap_or(false);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Some(ref mut w) = fallback_capture {
                    let _ = w.write_event(Event::Text(e.clone().into_owned()));
                } else if let Some(ref mut w) = capture {
                    let _ = w.write_event(Event::Text(e.clone().into_owned()));
                }
                if let Ok(text) = e.unescape() {
                    let text = text.trim();
                    if in_col {
                        if in_from {
                            from.col = text.parse().unwrap_or(0);
                        } else if in_to {
                            to.col = text.parse().unwrap_or(0);
                        }
                    } else if in_col_off {
                        if in_from {
                            from.col_offset_emu = text.parse().unwrap_or(0);
                        } else if in_to {
                            to.col_offset_emu = text.parse().unwrap_or(0);
                        }
                    } else if in_row {
                        if in_from {
                            from.row = text.parse().unwrap_or(0);
                        } else if in_to {
                            to.row = text.parse().unwrap_or(0);
                        }
                    } else if in_row_off {
                        if in_from {
                            from.row_offset_emu = text.parse().unwrap_or(0);
                        } else if in_to {
                            to.row_offset_emu = text.parse().unwrap_or(0);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if let Some(ref mut w) = fallback_capture {
                    fallback_depth -= 1;
                    let _ = w.write_event(Event::End(e.clone().into_owned()));
                    if fallback_depth == 0 {
                        if let Some(w) = fallback_capture.take() {
                            raw_mc_fallback = Some(w.into_inner().into_inner());
                        }
                    }
                } else if let Some(ref mut w) = capture {
                    let _ = w.write_event(Event::End(e.clone().into_owned()));
                }
                match e.name().local_name().as_ref() {
                    b"pic" => {
                        if in_pic {
                            if let Some(rel_id) = blip_rel_id.take() {
                                // Build the appropriate DrawingAnchor
                                // variant based on which anchor element
                                // wraps this <xdr:pic>.
                                let anchor = if in_one_cell_anchor {
                                    DrawingAnchor::OneCell {
                                        from: from.clone(),
                                        width_emu: anchor_ext_cx,
                                        height_emu: anchor_ext_cy,
                                    }
                                } else if in_absolute_anchor {
                                    DrawingAnchor::Absolute {
                                        x_emu: anchor_pos_x,
                                        y_emu: anchor_pos_y,
                                        width_emu: anchor_ext_cx,
                                        height_emu: anchor_ext_cy,
                                    }
                                } else {
                                    DrawingAnchor::TwoCell {
                                        from: from.clone(),
                                        to: to.clone(),
                                        edit_as: twocell_edit_as.clone(),
                                    }
                                };
                                images.push(EmbeddedImage {
                                    id: pic_id,
                                    name: std::mem::take(&mut pic_name),
                                    description: pic_descr.take(),
                                    anchor,
                                    format: ImageFormat::Png, // placeholder, resolved later from media path
                                    media_path: rel_id,
                                    svg_media_path: svg_blip_rel_id.take(),
                                    width_emu: pic_width_emu,
                                    height_emu: pic_height_emu,
                                    rotation: pic_rotation,
                                    flip_h: pic_flip_h,
                                    flip_v: pic_flip_v,
                                    data: Vec::new(), // populated later from archive
                                    svg_data: None,   // populated later from archive
                                });
                            }
                            in_pic = false;
                        }
                    }
                    b"grpSp" => _in_grp_sp = false,
                    b"spPr" if in_pic && in_sp_pr => {
                        in_sp_pr = false;
                        sp_pr_depth = 0;
                    }
                    _ if in_sp_pr && in_pic => {
                        sp_pr_depth = sp_pr_depth.saturating_sub(1);
                    }
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        if let Some(rel_id) = chart_rel_id.take() {
                            chart_refs.push(DrawingChartRef {
                                rel_id,
                                anchor: DrawingAnchor::TwoCell {
                                    from: from.clone(),
                                    to: to.clone(),
                                    edit_as: None,
                                },
                                is_chart_ex,
                                raw_mc_fallback: raw_mc_fallback.take(),
                            });
                        } else if let Some(w) = capture.take() {
                            if !saw_compat_ext {
                                raw_non_chart_anchors.push(w.into_inner().into_inner());
                            }
                        }
                        capture = None;
                        saw_compat_ext = false;
                        in_two_cell_anchor = false;
                        in_one_cell_anchor = false;
                        in_absolute_anchor = false;
                        in_from = false;
                        in_to = false;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        in_pic = false;
                        _in_grp_sp = false;
                    }
                    b"from" => in_from = false,
                    b"to" => in_to = false,
                    b"col" => in_col = false,
                    b"colOff" => in_col_off = false,
                    b"row" => in_row = false,
                    b"rowOff" => in_row_off = false,
                    b"graphicData" => in_graphic_data = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {
                if let Some(ref mut _w) = fallback_capture {
                    // Forward other events (CData, PI, etc.) into fallback capture
                    // We can't easily clone arbitrary events, so just skip.
                } else if let Some(ref mut _w) = capture {
                    // same
                }
            }
        }
        buf.clear();
    }

    Ok(DrawingContents {
        chart_refs,
        images,
        raw_non_chart_anchors,
    })
}

/// Backward-compatible wrapper that returns only chart refs.
#[cfg(test)]
pub(crate) fn read_drawing_chart_refs<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<Vec<DrawingChartRef>> {
    Ok(read_drawing_contents(archive, drawing_path)?.chart_refs)
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
        assert!(refs[0].raw_mc_fallback.is_some());
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
        let contents = read_drawing_contents(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(contents.images.len(), 1);
        assert!(contents.chart_refs.is_empty());
        assert_eq!(contents.raw_non_chart_anchors.len(), 1);

        let img = &contents.images[0];
        assert_eq!(img.id, 2);
        assert_eq!(img.name, "Picture 1");
        assert_eq!(img.description, Some("A test image".to_string()));
        assert_eq!(img.media_path, "rId1");
        assert_eq!(img.width_emu, 1000000);
        assert_eq!(img.height_emu, 2000000);
        assert_eq!(img.rotation, Some(5400000));
        assert!(img.flip_h);
        assert!(!img.flip_v);
        assert!(img.svg_media_path.is_none());

        if let DrawingAnchor::TwoCell { from, to, .. } = &img.anchor {
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
        let contents = read_drawing_contents(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(contents.images.len(), 1);
        assert!(contents.chart_refs.is_empty());

        let img = &contents.images[0];
        assert_eq!(img.id, 3);
        assert_eq!(img.name, "SVG Pic");
        assert_eq!(img.media_path, "rId2");
        assert_eq!(img.svg_media_path, Some("rId3".to_string()));
        assert_eq!(img.width_emu, 500000);
        assert_eq!(img.height_emu, 500000);
    }
}
