use std::io::{Read, Seek};

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::{Chart, DrawingAnchor};

pub(crate) fn read_chart<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    chart_path: &str,
    anchor: DrawingAnchor,
) -> XlsxResult<Option<Chart>> {
    let file = match archive.by_name(chart_path) {
        Ok(f) => f,
        Err(_) => return Ok(None),
    };

    let mut chart = duke_sheets_chart::parse::parse_chart_xml(file).map_err(|e| {
        XlsxError::Xml(quick_xml::Error::from(std::io::Error::new(
            std::io::ErrorKind::Other,
            e.to_string(),
        )))
    })?;
    chart.anchor = anchor;
    Ok(Some(chart))
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;
    use crate::reader::XlsxReader;
    use duke_sheets_chart::{ChartType, DataReference, DrawingAnchor, LegendPosition};

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
    fn test_parse_bar_chart() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Sales Chart</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Category</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:scaling><c:min val="0"/><c:max val="100"/></c:scaling>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

        let mut archive = zip_with_entry("xl/charts/chart1.xml", xml);
        let chart = read_chart(
            &mut archive,
            "xl/charts/chart1.xml",
            DrawingAnchor::default(),
        )
        .unwrap()
        .unwrap();

        assert_eq!(chart.chart_type, ChartType::ColumnClustered);
        assert_eq!(chart.title.as_deref(), Some("Sales Chart"));
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].name.as_deref(), Some("Sheet1!$B$1"));
        match &chart.series[0].values {
            DataReference::Formula(f) => assert_eq!(f, "Sheet1!$B$2:$B$4"),
            other => panic!("expected Formula, got {:?}", other),
        }
        match chart.series[0].categories.as_ref().unwrap() {
            DataReference::Formula(f) => assert_eq!(f, "Sheet1!$A$2:$A$4"),
            other => panic!("expected Formula, got {:?}", other),
        }
        assert_eq!(
            chart.category_axis.as_ref().unwrap().title.as_deref(),
            Some("Category")
        );
        let val_ax = chart.value_axis.as_ref().unwrap();
        assert_eq!(val_ax.title.as_deref(), Some("Value"));
        assert_eq!(val_ax.minimum, Some(0.0));
        assert_eq!(val_ax.maximum, Some(100.0));
        assert_eq!(
            chart.legend.as_ref().unwrap().position,
            LegendPosition::Bottom
        );
    }

    #[test]
    fn test_full_xlsx_with_chart_roundtrip() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData><drawing r:id="rId1"/></worksheet>"#).unwrap();

            zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/drawings/drawing1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#).unwrap();

            zip.start_file("xl/drawings/_rels/drawing1.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/charts/chart1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue by Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser><c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:min val="0"/><c:max val="50000"/></c:scaling></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>"#).unwrap();

            zip.finish().unwrap();
        }

        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(sheet.chart_count(), 1);
        let chart = &sheet.charts()[0];

        assert_eq!(chart.chart_type, ChartType::ColumnClustered);
        assert_eq!(chart.title.as_deref(), Some("Revenue by Quarter"));
        assert_eq!(chart.series.len(), 2);

        if let DrawingAnchor::TwoCell { from, to, .. } = &chart.anchor {
            assert_eq!(from.col, 2);
            assert_eq!(from.row, 3);
            assert_eq!(to.col, 12);
            assert_eq!(to.row, 18);
        } else {
            panic!("expected TwoCell anchor");
        }

        let val_ax = chart.value_axis.as_ref().unwrap();
        assert_eq!(val_ax.minimum, Some(0.0));
        assert_eq!(val_ax.maximum, Some(50000.0));

        assert_eq!(
            chart.legend.as_ref().unwrap().position,
            LegendPosition::Right
        );
    }

    #[test]
    fn test_xlsx_without_charts() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#).unwrap();
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();
            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();
            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#).unwrap();

            zip.finish().unwrap();
        }

        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(sheet.chart_count(), 0);
    }
}
