use std::io::{Cursor, Seek, Write};

use duke_sheets_chart::{Chart, PivotChartSource};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use super::XlsxResult;

pub(super) fn write_chart_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    chart: &Chart,
    chart_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/charts/chart{}.xml", chart_num);
    let mut bytes = duke_sheets_chart::write::chart_part_bytes(chart)?;
    if let Some(source) = &chart.pivot_source {
        insert_pivot_source(&mut bytes, source)?;
    }
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
}

fn insert_pivot_source(bytes: &mut Vec<u8>, source: &PivotChartSource) -> XlsxResult<()> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Start(BytesStart::new("c:pivotSource")))?;
    writer
        .create_element("c:name")
        .write_text_content(BytesText::new(&source.name))?;
    let format_id = source.format_id.to_string();
    writer
        .create_element("c:fmtId")
        .with_attribute(("val", format_id.as_str()))
        .write_empty()?;
    writer.write_event(Event::End(BytesEnd::new("c:pivotSource")))?;

    let marker = b"<c:chart>";
    let position = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or_else(|| {
            crate::error::XlsxError::InvalidFormat(
                "generated chart has no c:chart element".to_string(),
            )
        })?;
    bytes.splice(position..position, writer.into_inner().into_inner());
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::io::{Cursor, Read, Write};

    use duke_sheets_chart::{
        Axis, AxisPosition, AxisType, Chart, ChartColor, ChartLine, ChartLines,
        ChartShapeProperties, ChartType, DataLabels, DataReference, DataSeries, PivotChartSource,
        UpDownBars,
    };

    use super::write_chart_part;
    use crate::reader::chart::read_chart;

    #[test]
    fn test_extlst_roundtrip() {
        // Build a chart with raw extension data at multiple levels
        let mut chart = Chart::new(ChartType::ColumnClustered);
        let mut ser = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"));
        ser.raw_ext =
            Some(b"<c:extLst><c:ext uri=\"{ser-ext}\"><serData/></c:ext></c:extLst>".to_vec());
        chart.series.push(ser);

        let mut exts = HashMap::new();
        exts.insert(
            "chartType".to_string(),
            b"<c:extLst><c:ext uri=\"{ct-ext}\"><ctData/></c:ext></c:extLst>".to_vec(),
        );
        exts.insert(
            "plotArea".to_string(),
            b"<c:extLst><c:ext uri=\"{pa-ext}\"><paData/></c:ext></c:extLst>".to_vec(),
        );
        exts.insert(
            "chart".to_string(),
            b"<c:extLst><c:ext uri=\"{ch-ext}\"><chData/></c:ext></c:extLst>".to_vec(),
        );
        exts.insert(
            "chartSpace".to_string(),
            b"<c:extLst><c:ext uri=\"{cs-ext}\"><csData/></c:ext></c:extLst>".to_vec(),
        );
        chart.raw_extensions = exts;

        // Write to a zip
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, &chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Read back
        let reparsed = read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap();

        // Verify series extLst survived
        let ser_ext = reparsed.series[0]
            .raw_ext
            .as_ref()
            .expect("series extLst lost");
        let ser_str = std::str::from_utf8(ser_ext).unwrap();
        assert!(
            ser_str.contains("ser-ext"),
            "series ext content lost: {}",
            ser_str
        );

        // Verify chart-level extensions survived
        let ct = reparsed
            .raw_extensions
            .get("chartType")
            .expect("chartType extLst lost");
        assert!(std::str::from_utf8(ct).unwrap().contains("ct-ext"));

        let pa = reparsed
            .raw_extensions
            .get("plotArea")
            .expect("plotArea extLst lost");
        assert!(std::str::from_utf8(pa).unwrap().contains("pa-ext"));

        let ch = reparsed
            .raw_extensions
            .get("chart")
            .expect("chart extLst lost");
        assert!(std::str::from_utf8(ch).unwrap().contains("ch-ext"));

        let cs = reparsed
            .raw_extensions
            .get("chartSpace")
            .expect("chartSpace extLst lost");
        assert!(std::str::from_utf8(cs).unwrap().contains("cs-ext"));
    }

    #[test]
    fn test_pivot_source_roundtrip() {
        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.pivot_source = Some(PivotChartSource::new("SalesPivot").with_format_id(4));
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));

        let written = chart_xml_after_write(&chart);
        assert!(written.contains(
            "<c:pivotSource><c:name>SalesPivot</c:name><c:fmtId val=\"4\"/></c:pivotSource>"
        ));
        let pivot_pos = written.find("<c:pivotSource>").unwrap();
        let chart_pos = written.find("<c:chart>").unwrap();
        assert!(pivot_pos < chart_pos);

        let reparsed = roundtrip_chart(&chart);
        assert_eq!(reparsed.pivot_source, chart.pivot_source);
    }

    #[test]
    fn test_date_ax_roundtrip() {
        let mut chart = Chart::new(ChartType::Line);
        let ser = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
        chart.series.push(ser);
        let mut cat_ax = Axis::new();
        cat_ax.axis_type = AxisType::Date;
        chart.category_axis = Some(cat_ax);
        chart.value_axis = Some(Axis::new());

        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, &chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let reparsed = read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap();

        let cat = reparsed.category_axis.unwrap();
        assert_eq!(cat.axis_type, AxisType::Date);
        let val = reparsed.value_axis.unwrap();
        assert_eq!(val.axis_type, AxisType::Value);
    }

    #[test]
    fn test_ser_ax_roundtrip() {
        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.is_3d = true;
        let ser = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"));
        chart.series.push(ser);
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        let mut ser_ax = Axis::new();
        ser_ax.axis_type = AxisType::Series;
        ser_ax.delete = Some(false);
        chart.series_axis = Some(ser_ax);

        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, &chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let reparsed = read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap();

        let ser = reparsed.series_axis.unwrap();
        assert_eq!(ser.axis_type, AxisType::Series);
        assert_eq!(ser.delete, Some(false));
    }

    #[test]
    fn test_cat_ax_default_roundtrip() {
        let mut chart = Chart::new(ChartType::ColumnClustered);
        let ser = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"));
        chart.series.push(ser);
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());

        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, &chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let reparsed = read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap();

        let cat = reparsed.category_axis.unwrap();
        assert_eq!(cat.axis_type, AxisType::Category);
    }

    fn roundtrip_chart(chart: &Chart) -> Chart {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap()
    }

    fn read_chart_from_xml(xml: &str) -> Chart {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        zip_writer
            .start_file(
                "xl/charts/chart1.xml",
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
        zip_writer.write_all(xml.as_bytes()).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        read_chart(&mut archive, "xl/charts/chart1.xml")
        .unwrap()
        .unwrap()
    }

    fn chart_xml_after_write(chart: &Chart) -> String {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut file = archive.by_name("xl/charts/chart1.xml").unwrap();
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        xml
    }

    #[test]
    fn test_imported_chart_colors_survive_write() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:spPr>
            <a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill>
            <a:ln w="0"><a:noFill/></a:ln>
          </c:spPr>
          <c:dPt>
            <c:idx val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/>
            <c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="2"/>
            <c:spPr><a:solidFill><a:srgbClr val="9BBB59"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="3"/>
            <c:spPr><a:solidFill><a:srgbClr val="8064A2"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$1:$B$4</c:f></c:numRef></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9360"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>"#;

        let chart = read_chart_from_xml(xml);
        let written = chart_xml_after_write(&chart);

        for expected in ["4F81BD", "C0504D", "9BBB59", "8064A2", "FFFFFF", "D9D9D9"] {
            assert!(
                written.contains(&format!("srgbClr val=\"{expected}\"")),
                "missing {expected} in {written}"
            );
        }
    }

    #[test]
    fn test_roundtrip_drop_lines() {
        let mut chart = Chart::new(ChartType::Line);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        chart.drop_lines = Some(ChartLines {
            shape_properties: Some(ChartShapeProperties {
                solid_fill: None,
                no_fill: false,
                line: Some(ChartLine {
                    width: Some(12700),
                    solid_fill: Some(ChartColor {
                        hex: "FF0000".into(),
                    }),
                    no_fill: false,
                    dash_style: Some("dash".into()),
                }),
            }),
        });

        let reparsed = roundtrip_chart(&chart);
        let dl = reparsed.drop_lines.expect("drop_lines lost");
        let sp = dl.shape_properties.expect("drop_lines spPr lost");
        let ln = sp.line.expect("drop_lines line lost");
        assert_eq!(ln.width, Some(12700));
        assert_eq!(ln.solid_fill.as_ref().unwrap().hex, "FF0000");
        assert_eq!(ln.dash_style.as_deref(), Some("dash"));
    }

    #[test]
    fn test_roundtrip_high_low_lines() {
        let mut chart = Chart::new(ChartType::Stock);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        chart.high_low_lines = Some(ChartLines {
            shape_properties: None,
        });

        let reparsed = roundtrip_chart(&chart);
        let hl = reparsed.high_low_lines.expect("high_low_lines lost");
        assert!(hl.shape_properties.is_none());
    }

    #[test]
    fn test_roundtrip_series_lines() {
        let mut chart = Chart::new(ChartType::BarStacked);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        let group1 = duke_sheets_chart::ChartTypeGroup {
            chart_type: ChartType::BarStacked,
            is_3d: false,
            series: vec![DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"))],
            data_labels: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            wireframe: None,
            drop_lines: None,
            high_low_lines: None,
            series_lines: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: Some(ChartColor {
                        hex: "00FF00".into(),
                    }),
                    no_fill: false,
                    line: None,
                }),
            }),
            up_down_bars: None,
            axis_ids: vec![1, 2],
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            raw_ext: None,
        };
        let group2 = duke_sheets_chart::ChartTypeGroup {
            chart_type: ChartType::Line,
            is_3d: false,
            series: vec![DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$3"))],
            data_labels: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            wireframe: None,
            drop_lines: None,
            high_low_lines: None,
            series_lines: None,
            up_down_bars: None,
            axis_ids: vec![1, 2],
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            raw_ext: None,
        };
        chart.type_groups = vec![group1, group2];
        chart.axes = vec![
            duke_sheets_chart::ChartAxis {
                id: 1,
                cross_id: 2,
                axis: Axis::new(),
            },
            duke_sheets_chart::ChartAxis {
                id: 2,
                cross_id: 1,
                axis: {
                    let mut a = Axis::new();
                    a.axis_type = duke_sheets_chart::AxisType::Value;
                    a
                },
            },
        ];

        let reparsed = roundtrip_chart(&chart);
        assert!(reparsed.type_groups.len() >= 2);
        let sl = reparsed.type_groups[0]
            .series_lines
            .as_ref()
            .expect("series_lines lost");
        let sp = sl.shape_properties.as_ref().expect("serLines spPr lost");
        assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "00FF00");
    }

    #[test]
    fn test_roundtrip_up_down_bars() {
        let mut chart = Chart::new(ChartType::Stock);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        chart.up_down_bars = Some(UpDownBars {
            gap_width: Some(150),
            up_bars: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: Some(ChartColor {
                        hex: "00FF00".into(),
                    }),
                    no_fill: false,
                    line: None,
                }),
            }),
            down_bars: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: Some(ChartColor {
                        hex: "FF0000".into(),
                    }),
                    no_fill: false,
                    line: None,
                }),
            }),
        });

        let reparsed = roundtrip_chart(&chart);
        let udb = reparsed.up_down_bars.expect("up_down_bars lost");
        assert_eq!(udb.gap_width, Some(150));
        let ub = udb.up_bars.expect("up_bars lost");
        assert_eq!(
            ub.shape_properties
                .as_ref()
                .unwrap()
                .solid_fill
                .as_ref()
                .unwrap()
                .hex,
            "00FF00"
        );
        let db = udb.down_bars.expect("down_bars lost");
        assert_eq!(
            db.shape_properties
                .as_ref()
                .unwrap()
                .solid_fill
                .as_ref()
                .unwrap()
                .hex,
            "FF0000"
        );
    }

    #[test]
    fn test_roundtrip_leader_lines() {
        let mut chart = Chart::new(ChartType::Pie);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.data_labels = Some(DataLabels {
            show_value: Some(true),
            leader_lines: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: None,
                    no_fill: false,
                    line: Some(ChartLine {
                        width: Some(9525),
                        solid_fill: Some(ChartColor {
                            hex: "808080".into(),
                        }),
                        no_fill: false,
                        dash_style: None,
                    }),
                }),
            }),
            ..DataLabels::default()
        });

        let reparsed = roundtrip_chart(&chart);
        let dl = reparsed.data_labels.expect("data_labels lost");
        let ll = dl.leader_lines.expect("leader_lines lost");
        let sp = ll.shape_properties.expect("leader_lines spPr lost");
        let ln = sp.line.expect("leader_lines line lost");
        assert_eq!(ln.width, Some(9525));
        assert_eq!(ln.solid_fill.as_ref().unwrap().hex, "808080");
    }

    /// A bar chart's value axis sits at the bottom. The writer used to
    /// treat that as "unset" because Bottom is the enum's default, and
    /// moved the axis to the left.
    #[test]
    fn an_explicit_bottom_value_axis_stays_at_the_bottom() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart><c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="111"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
        let chart = read_chart_from_xml(xml);
        let value_axis = chart.value_axis.as_ref().expect("value axis");
        assert_eq!(value_axis.position, Some(AxisPosition::Bottom));
        assert_eq!(
            chart.category_axis.as_ref().expect("category axis").position,
            Some(AxisPosition::Left),
            "the category axis keeps its own explicit position too"
        );

        let written = chart_xml_after_write(&chart);
        let val_ax = &written[written.find("<c:valAx>").expect("valAx")..];
        assert!(
            val_ax.contains(r#"<c:axPos val="b"/>"#),
            "the value axis must stay at the bottom: {val_ax}"
        );
    }

    /// c:plotVisOnly defaults to true, so an explicit val="0" - plot the
    /// data in hidden cells too - has to be written out, or the chart
    /// quietly loses those points.
    #[test]
    fn plot_visible_only_false_survives_a_round_trip() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart><c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="111"/></c:valAx></c:plotArea><c:plotVisOnly val="0"/></c:chart></c:chartSpace>"#;
        let chart = read_chart_from_xml(xml);
        assert_eq!(chart.plot_visible_only, Some(false));

        let written = chart_xml_after_write(&chart);
        assert!(
            written.contains(r#"<c:plotVisOnly val="0"/>"#),
            "the explicit false must be written: {written}"
        );
        assert_eq!(read_chart_from_xml(&written).plot_visible_only, Some(false));
    }

    /// Excel treats an axis with no c:delete as deleted: it adds
    /// <c:delete val="1"/> and discards the axis formatting when it opens
    /// such a chart. Reading the omission as "unspecified" and writing
    /// val="0" therefore made a hidden axis visible on rewrite.
    #[test]
    fn an_axis_without_delete_is_read_as_deleted() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart><c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="111"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
        let chart = read_chart_from_xml(xml);
        assert_eq!(
            chart.category_axis.as_ref().expect("category axis").delete,
            Some(true),
            "an omitted c:delete means the axis is deleted"
        );

        // Written back explicitly, so Excel keeps reading it as deleted
        // instead of depending on the omission.
        let written = chart_xml_after_write(&chart);
        assert!(
            written.contains(r#"<c:delete val="1"/>"#),
            "the deleted state must be written explicitly: {written}"
        );
        assert!(
            !written.contains(r#"<c:delete val="0"/>"#),
            "no axis may be revived as visible: {written}"
        );

        let reread = read_chart_from_xml(&written);
        assert_eq!(
            reread.category_axis.as_ref().expect("category axis").delete,
            Some(true),
            "the deleted state must survive the round trip"
        );
    }

    /// An axis the caller builds has no delete state, and must come out
    /// visible rather than inheriting the omission's meaning.
    #[test]
    fn a_model_built_axis_is_written_as_not_deleted() {
        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart
            .series
            .push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        assert_eq!(chart.category_axis.as_ref().unwrap().delete, None);

        let written = chart_xml_after_write(&chart);
        assert!(
            written.contains(r#"<c:delete val="0"/>"#),
            "a built axis must be written visible: {written}"
        );
        assert!(!written.contains(r#"<c:delete val="1"/>"#));
    }
}
