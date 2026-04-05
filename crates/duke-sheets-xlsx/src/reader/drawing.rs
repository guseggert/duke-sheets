use std::io::{BufReader, Cursor, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::ChartAnchor;

/// A chart reference discovered in a drawing XML, paired with its anchor position.
pub(crate) struct DrawingChartRef {
    /// The relationship id (e.g. "rId1") pointing to the chart part.
    pub(crate) rel_id: String,
    /// The two-cell anchor positioning the chart in the worksheet.
    pub(crate) anchor: ChartAnchor,
    /// Whether this references a ChartEx part (`cx:chart`) rather than a standard chart.
    pub(crate) is_chart_ex: bool,
    /// Raw `mc:Fallback` XML bytes for roundtrip (chartEx only).
    pub(crate) raw_mc_fallback: Option<Vec<u8>>,
}

/// Chart refs plus raw non-chart drawing anchors from a drawing XML.
pub(crate) struct DrawingContents {
    pub chart_refs: Vec<DrawingChartRef>,
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
    let file = match archive.by_name(drawing_path) {
        Ok(f) => f,
        Err(_) => {
            return Ok(DrawingContents {
                chart_refs: Vec::new(),
                raw_non_chart_anchors: Vec::new(),
            })
        }
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut chart_refs = Vec::new();
    let mut raw_non_chart_anchors: Vec<Vec<u8>> = Vec::new();

    let mut in_two_cell_anchor = false;
    let mut in_one_cell_anchor = false;
    let mut in_absolute_anchor = false;
    let mut anchor = ChartAnchor::default();
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
                        anchor = ChartAnchor::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"oneCellAnchor" => {
                        in_one_cell_anchor = true;
                        anchor = ChartAnchor::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"absoluteAnchor" => {
                        in_absolute_anchor = true;
                        anchor = ChartAnchor::default();
                        chart_rel_id = None;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        capture = Some(w);
                    }
                    b"from" if in_two_cell_anchor || in_one_cell_anchor => {
                        in_from = true;
                    }
                    b"to" if in_two_cell_anchor => {
                        in_to = true;
                    }
                    b"col" if in_from || in_to => {
                        in_col = true;
                    }
                    b"colOff" if in_from || in_to => {
                        in_col_off = true;
                    }
                    b"row" if in_from || in_to => {
                        in_row = true;
                    }
                    b"rowOff" if in_from || in_to => {
                        in_row_off = true;
                    }
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
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(ref mut w) = fallback_capture {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                } else if let Some(ref mut w) = capture {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                }
                if in_graphic_data && e.name().local_name().as_ref() == b"chart" {
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
                            anchor.from_col = text.parse().unwrap_or(0);
                        } else if in_to {
                            anchor.to_col = text.parse().unwrap_or(0);
                        }
                    } else if in_col_off {
                        if in_from {
                            anchor.from_col_offset = text.parse().unwrap_or(0);
                        } else if in_to {
                            anchor.to_col_offset = text.parse().unwrap_or(0);
                        }
                    } else if in_row {
                        if in_from {
                            anchor.from_row = text.parse().unwrap_or(0);
                        } else if in_to {
                            anchor.to_row = text.parse().unwrap_or(0);
                        }
                    } else if in_row_off {
                        if in_from {
                            anchor.from_row_offset = text.parse().unwrap_or(0);
                        } else if in_to {
                            anchor.to_row_offset = text.parse().unwrap_or(0);
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
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        if let Some(rel_id) = chart_rel_id.take() {
                            chart_refs.push(DrawingChartRef {
                                rel_id,
                                anchor: anchor.clone(),
                                is_chart_ex,
                                raw_mc_fallback: raw_mc_fallback.take(),
                            });
                        } else if let Some(w) = capture.take() {
                            raw_non_chart_anchors.push(w.into_inner().into_inner());
                        }
                        capture = None;
                        in_two_cell_anchor = false;
                        in_one_cell_anchor = false;
                        in_absolute_anchor = false;
                        in_from = false;
                        in_to = false;
                        is_chart_ex = false;
                        raw_mc_fallback = None;
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
        assert_eq!(refs[0].anchor.from_col, 1);
        assert_eq!(refs[0].anchor.from_col_offset, 100);
        assert_eq!(refs[0].anchor.from_row, 2);
        assert_eq!(refs[0].anchor.from_row_offset, 200);
        assert_eq!(refs[0].anchor.to_col, 10);
        assert_eq!(refs[0].anchor.to_col_offset, 300);
        assert_eq!(refs[0].anchor.to_row, 20);
        assert_eq!(refs[0].anchor.to_row_offset, 400);
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
        assert_eq!(refs[0].anchor.from_col, 1);
        assert_eq!(refs[0].anchor.from_col_offset, 100);
        assert_eq!(refs[0].anchor.from_row, 2);
        assert_eq!(refs[0].anchor.from_row_offset, 200);
        assert_eq!(refs[0].anchor.to_col, 0);
        assert_eq!(refs[0].anchor.to_row, 0);
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
        assert_eq!(refs[0].anchor.from_col, 0);
        assert_eq!(refs[0].anchor.from_row, 0);
        assert_eq!(refs[0].anchor.to_col, 0);
        assert_eq!(refs[0].anchor.to_row, 0);
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
}
