use std::io::{Read, Seek};

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};

pub(crate) use duke_sheets_chart::drawing_part::read::{
    DrawingChartRef, DrawingEntry, DrawingEntryKind, ParsedChild, ParsedGroup, ParsedShape,
    PicShape,
};

/// Parse a SpreadsheetML drawing part into its top-level entries in
/// document order (see [`duke_sheets_chart::drawing_part::read`]).
/// A missing part yields an empty list.
pub(crate) fn read_drawing_entries<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<Vec<DrawingEntry>> {
    let mut file = match archive_by_name(archive, drawing_path) {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    duke_sheets_chart::drawing_part::read::parse_drawing_part(&bytes)
        .map_err(|e| XlsxError::InvalidFormat(format!("drawing part: {e}")))
}

/// Backward-compatible view returning only chart refs.
#[cfg(test)]
pub(crate) fn read_drawing_chart_refs<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    drawing_path: &str,
) -> XlsxResult<Vec<DrawingChartRef>> {
    Ok(read_drawing_entries(archive, drawing_path)?
        .into_iter()
        .filter_map(|entry| match entry.kind {
            DrawingEntryKind::Chart(chart_ref) => Some(chart_ref),
            _ => None,
        })
        .collect())
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;
    use duke_sheets_chart::{DrawingAnchor, EditAs};

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
        // Fallback inner content is captured without the wrapper tags.
        let fallback = refs[0].raw_mc_fallback.as_deref().expect("fallback");
        let fallback = std::str::from_utf8(fallback).unwrap();
        assert!(fallback.starts_with("<xdr:sp>"), "{fallback}");
        assert!(!fallback.contains("mc:Fallback"), "{fallback}");
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
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        // Single ownership: exactly one entry, classified as an image.
        assert_eq!(entries.len(), 1);
        let entry = &entries[0];
        let DrawingEntryKind::Image(pic) = &entry.kind else {
            panic!("expected image entry");
        };
        assert_eq!(pic.name, "Picture 1");
        assert_eq!(pic.descr.as_deref(), Some("A test image"));
        assert_eq!(pic.blip_rel.as_deref(), Some("rId1"));
        assert_eq!(pic.ext_cx, 1000000);
        assert_eq!(pic.ext_cy, 2000000);
        assert_eq!(pic.rotation, Some(5400000));
        assert!(pic.flip_h);
        assert!(!pic.flip_v);
        assert!(pic.svg_rel.is_none());
        assert!(entry.locked);
        assert!(entry.printable);

        if let DrawingAnchor::TwoCell { from, to, .. } = &entry.anchor {
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
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Image(pic) = &entries[0].kind else {
            panic!("expected image entry");
        };
        assert_eq!(pic.name, "SVG Pic");
        assert_eq!(pic.blip_rel.as_deref(), Some("rId2"));
        assert_eq!(pic.svg_rel.as_deref(), Some("rId3"));
        assert_eq!(pic.ext_cx, 500000);
        assert_eq!(pic.ext_cy, 500000);
    }

    #[test]
    fn test_parse_control_twin_and_client_data() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
      <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="3" name="Check Box 1" hidden="1">
              <a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1026"/></a:ext></a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto">
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </xdr:spPr>
        </xdr:sp>
        <xdr:clientData fLocksWithSheet="0" fPrintsWithSheet="0"/>
      </xdr:twoCellAnchor>
    </mc:Choice>
    <mc:Fallback/>
  </mc:AlternateContent>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let entry = &entries[0];
        let DrawingEntryKind::ControlTwin(twin) = &entry.kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.spid, "_x0000_s1026");
        assert_eq!(twin.shape_num, Some(1026));
        assert!(!entry.locked);
        assert!(!entry.printable);
        if let DrawingAnchor::TwoCell { edit_as, .. } = &entry.anchor {
            assert_eq!(edit_as, &Some(EditAs::OneCell));
        } else {
            panic!("expected TwoCell anchor");
        }
    }

    #[test]
    fn test_parse_control_twin_inside_anchor_alternate_content() {
        // The other Excel emission shape: a bare anchor whose content
        // is mc:AlternateContent wrapping the a14 twin sp.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="3" name="Check Box 7" hidden="1">
              <a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1031"/></a:ext></a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></xdr:spPr>
        </xdr:sp>
      </mc:Choice>
      <mc:Fallback/>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::ControlTwin(twin) = &entries[0].kind else {
            panic!("expected control twin entry");
        };
        assert_eq!(twin.shape_num, Some(1031));
    }

    #[test]
    fn test_parse_group_of_pics() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
      <xdr:grpSpPr>
        <a:xfrm>
          <a:off x="609600" y="190500"/><a:ext cx="1219200" cy="190500"/>
          <a:chOff x="0" y="0"/><a:chExt cx="1219200" cy="190500"/>
        </a:xfrm>
      </xdr:grpSpPr>
      <xdr:pic>
        <xdr:nvPicPr><xdr:cNvPr id="5" name="Left"/><xdr:cNvPicPr/></xdr:nvPicPr>
        <xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></xdr:blipFill>
        <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm></xdr:spPr>
      </xdr:pic>
      <xdr:pic>
        <xdr:nvPicPr><xdr:cNvPr id="6" name="Right"/><xdr:cNvPicPr/></xdr:nvPicPr>
        <xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/></xdr:blipFill>
        <xdr:spPr><a:xfrm><a:off x="609600" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm></xdr:spPr>
      </xdr:pic>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();

        assert_eq!(entries.len(), 1);
        let DrawingEntryKind::Group(group) = &entries[0].kind else {
            panic!("expected group entry");
        };
        assert_eq!(group.name, "Group 1");
        assert_eq!(group.transform.x_emu, 609600);
        assert_eq!(group.transform.child_cx_emu, 1219200);
        assert_eq!(group.children.len(), 2);
        let ParsedChild::Pic(right) = &group.children[1] else {
            panic!("expected pic child");
        };
        assert_eq!(right.name, "Right");
        assert_eq!(right.off_x, 609600);
        assert_eq!(right.blip_rel.as_deref(), Some("rId2"));
    }

    #[test]
    fn test_group_with_unmodeled_child_degrades_to_raw() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
      <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/><a:chOff x="0" y="0"/><a:chExt cx="1" cy="1"/></a:xfrm></xdr:grpSpPr>
      <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>
        <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr>
        <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>hi</a:t></a:r></a:p></xdr:txBody>
      </xdr:sp>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut archive = zip_with_entry("xl/drawings/drawing1.xml", xml);
        let entries = read_drawing_entries(&mut archive, "xl/drawings/drawing1.xml").unwrap();
        assert_eq!(entries.len(), 1);
        assert!(matches!(entries[0].kind, DrawingEntryKind::Raw));
        let raw = std::str::from_utf8(&entries[0].bytes).unwrap();
        assert!(raw.contains("TextBox 1"), "raw bytes keep the group: {raw}");
    }
}
