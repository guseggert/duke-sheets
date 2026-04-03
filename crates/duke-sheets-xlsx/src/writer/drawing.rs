use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_chart::{Chart, ChartAnchor};

use super::{write_xml_part, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS, RT_CHART};

const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

pub(super) fn write_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    charts: &[&Chart],
    raw_drawing_objects: &[Vec<u8>],
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("xdr:wsDr");
        tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
        tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        w.write_event(Event::Start(tag))?;

        for (i, chart) in charts.iter().enumerate() {
            let rid = format!("rId{}", i + 1);
            write_two_cell_anchor(w, &chart.anchor, &rid, i)?;
        }

        for raw in raw_drawing_objects {
            w.get_mut().write_all(raw)?;
        }

        w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
        Ok(())
    })
}

pub(super) fn write_chartsheet_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    _chart: &Chart,
    raw_drawing_objects: &[Vec<u8>],
    drawing_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/drawings/drawing{}.xml", drawing_num);
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("xdr:wsDr");
        tag.push_attribute(("xmlns:xdr", NS_SPREADSHEET_DRAWING));
        tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        w.write_event(Event::Start(tag))?;

        write_absolute_anchor(w, "rId1", 0)?;

        for raw in raw_drawing_objects {
            w.get_mut().write_all(raw)?;
        }

        w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
        Ok(())
    })
}

fn write_absolute_anchor(
    w: &mut XmlWriter,
    rid: &str,
    chart_idx: usize,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;

    w.create_element("xdr:pos")
        .with_attribute(("x", "0"))
        .with_attribute(("y", "0"))
        .write_empty()?;
    w.create_element("xdr:ext")
        .with_attribute(("cx", "9144000"))
        .with_attribute(("cy", "6858000"))
        .write_empty()?;

    w.write_event(Event::Start(BytesStart::new("xdr:graphicFrame")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id = (chart_idx + 2).to_string();
    let name = format!("Chart {}", chart_idx + 1);
    w.create_element("xdr:cNvPr")
        .with_attribute(("id", cnv_id.as_str()))
        .with_attribute(("name", name.as_str()))
        .write_empty()?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:xfrm")))?;
    w.create_element("a:off")
        .with_attribute(("x", "0"))
        .with_attribute(("y", "0"))
        .write_empty()?;
    w.create_element("a:ext")
        .with_attribute(("cx", "0"))
        .with_attribute(("cy", "0"))
        .write_empty()?;
    w.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

    w.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", NS_CHART));
    w.write_event(Event::Start(gd))?;

    let mut chart_el = BytesStart::new("c:chart");
    chart_el.push_attribute(("xmlns:c", NS_CHART));
    chart_el.push_attribute(("r:id", rid));
    w.write_event(Event::Empty(chart_el))?;

    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
    Ok(())
}

fn write_two_cell_anchor(
    w: &mut XmlWriter,
    anchor: &ChartAnchor,
    rid: &str,
    chart_idx: usize,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("xdr:twoCellAnchor")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
    write_anchor_point(
        w,
        anchor.from_col,
        anchor.from_col_offset,
        anchor.from_row,
        anchor.from_row_offset,
    )?;
    w.write_event(Event::End(BytesEnd::new("xdr:from")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
    write_anchor_point(
        w,
        anchor.to_col,
        anchor.to_col_offset,
        anchor.to_row,
        anchor.to_row_offset,
    )?;
    w.write_event(Event::End(BytesEnd::new("xdr:to")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:graphicFrame")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id = (chart_idx + 2).to_string();
    let name = format!("Chart {}", chart_idx + 1);
    w.create_element("xdr:cNvPr")
        .with_attribute(("id", cnv_id.as_str()))
        .with_attribute(("name", name.as_str()))
        .write_empty()?;
    w.write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:xfrm")))?;
    w.create_element("a:off")
        .with_attribute(("x", "0"))
        .with_attribute(("y", "0"))
        .write_empty()?;
    w.create_element("a:ext")
        .with_attribute(("cx", "9525000"))
        .with_attribute(("cy", "6096000"))
        .write_empty()?;
    w.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

    w.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", NS_CHART));
    w.write_event(Event::Start(gd))?;

    let mut chart_el = BytesStart::new("c:chart");
    chart_el.push_attribute(("xmlns:c", NS_CHART));
    chart_el.push_attribute(("r:id", rid));
    w.write_event(Event::Empty(chart_el))?;

    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;
    w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
    Ok(())
}

fn write_anchor_point(
    w: &mut XmlWriter,
    col: u16,
    col_off: i64,
    row: u32,
    row_off: i64,
) -> XlsxResult<()> {
    let col_s = col.to_string();
    w.create_element("xdr:col")
        .write_text_content(BytesText::new(&col_s))?;
    let col_off_s = col_off.to_string();
    w.create_element("xdr:colOff")
        .write_text_content(BytesText::new(&col_off_s))?;
    let row_s = row.to_string();
    w.create_element("xdr:row")
        .write_text_content(BytesText::new(&row_s))?;
    let row_off_s = row_off.to_string();
    w.create_element("xdr:rowOff")
        .write_text_content(BytesText::new(&row_off_s))?;
    Ok(())
}

pub(super) fn write_drawing_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    drawing_num: usize,
    chart_nums: &[usize],
) -> XlsxResult<()> {
    let path = format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_num);
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("Relationships");
        tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
        w.write_event(Event::Start(tag))?;

        for (i, &chart_num) in chart_nums.iter().enumerate() {
            let rid = format!("rId{}", i + 1);
            let target = format!("../charts/chart{}.xml", chart_num);
            w.create_element("Relationship")
                .with_attribute(("Id", rid.as_str()))
                .with_attribute(("Type", RT_CHART))
                .with_attribute(("Target", target.as_str()))
                .write_empty()?;
        }

        w.write_event(Event::End(BytesEnd::new("Relationships")))?;
        Ok(())
    })
}
