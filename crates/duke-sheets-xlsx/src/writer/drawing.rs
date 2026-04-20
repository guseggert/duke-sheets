use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_chart::{CellMarker, Chart, ChartEx, DrawingAnchor};

use super::{write_xml_part, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS, RT_CHART};

const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const NS_MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CX1: &str = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const RT_CHART_EX: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

pub(super) fn write_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    charts: &[&Chart],
    charts_ex: &[&ChartEx],
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

        let chartex_rid_start = charts.len() + 1;
        for (i, cx) in charts_ex.iter().enumerate() {
            let rid = format!("rId{}", chartex_rid_start + i);
            let obj_idx = charts.len() + i;
            write_chartex_two_cell_anchor(
                w,
                &cx.anchor,
                &rid,
                obj_idx,
                cx.raw_mc_fallback.as_deref(),
            )?;
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

fn write_absolute_anchor(w: &mut XmlWriter, rid: &str, chart_idx: usize) -> XlsxResult<()> {
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
    anchor: &DrawingAnchor,
    rid: &str,
    chart_idx: usize,
) -> XlsxResult<()> {
    let (from, to) = match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => (from.clone(), to.clone()),
        _ => (CellMarker::default(), CellMarker::default()),
    };
    w.write_event(Event::Start(BytesStart::new("xdr:twoCellAnchor")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
    write_cell_marker(w, &from)?;
    w.write_event(Event::End(BytesEnd::new("xdr:from")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
    write_cell_marker(w, &to)?;
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

fn write_chartex_two_cell_anchor(
    w: &mut XmlWriter,
    anchor: &DrawingAnchor,
    rid: &str,
    obj_idx: usize,
    raw_mc_fallback: Option<&[u8]>,
) -> XlsxResult<()> {
    let (from, to) = match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => (from.clone(), to.clone()),
        _ => (CellMarker::default(), CellMarker::default()),
    };
    w.write_event(Event::Start(BytesStart::new("xdr:twoCellAnchor")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
    write_cell_marker(w, &from)?;
    w.write_event(Event::End(BytesEnd::new("xdr:from")))?;

    w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
    write_cell_marker(w, &to)?;
    w.write_event(Event::End(BytesEnd::new("xdr:to")))?;

    let mut mc_tag = BytesStart::new("mc:AlternateContent");
    mc_tag.push_attribute(("xmlns:mc", NS_MC));
    w.write_event(Event::Start(mc_tag))?;

    let mut choice_tag = BytesStart::new("mc:Choice");
    choice_tag.push_attribute(("xmlns:cx1", NS_CX1));
    choice_tag.push_attribute(("Requires", "cx1"));
    w.write_event(Event::Start(choice_tag))?;

    let mut gf_tag = BytesStart::new("xdr:graphicFrame");
    gf_tag.push_attribute(("macro", ""));
    w.write_event(Event::Start(gf_tag))?;

    w.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
    let cnv_id = (obj_idx + 2).to_string();
    let name = format!("Chart {}", obj_idx + 1);
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
    gd.push_attribute(("uri", NS_CX));
    w.write_event(Event::Start(gd))?;

    let mut cx_chart = BytesStart::new("cx:chart");
    cx_chart.push_attribute(("xmlns:cx", NS_CX));
    cx_chart.push_attribute(("xmlns:r", NS_DOC_RELS));
    cx_chart.push_attribute(("r:id", rid));
    w.write_event(Event::Empty(cx_chart))?;

    w.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    w.write_event(Event::End(BytesEnd::new("a:graphic")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;

    w.write_event(Event::End(BytesEnd::new("mc:Choice")))?;

    w.write_event(Event::Start(BytesStart::new("mc:Fallback")))?;
    if let Some(raw) = raw_mc_fallback {
        w.get_mut().write_all(raw)?;
    }
    w.write_event(Event::End(BytesEnd::new("mc:Fallback")))?;

    w.write_event(Event::End(BytesEnd::new("mc:AlternateContent")))?;

    w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
    Ok(())
}

fn write_cell_marker(w: &mut XmlWriter, marker: &CellMarker) -> XlsxResult<()> {
    let col_s = marker.col.to_string();
    w.create_element("xdr:col")
        .write_text_content(BytesText::new(&col_s))?;
    let col_off_s = marker.col_offset_emu.to_string();
    w.create_element("xdr:colOff")
        .write_text_content(BytesText::new(&col_off_s))?;
    let row_s = marker.row.to_string();
    w.create_element("xdr:row")
        .write_text_content(BytesText::new(&row_s))?;
    let row_off_s = marker.row_offset_emu.to_string();
    w.create_element("xdr:rowOff")
        .write_text_content(BytesText::new(&row_off_s))?;
    Ok(())
}

pub(super) fn write_drawing_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    drawing_num: usize,
    chart_nums: &[usize],
    chart_ex_nums: &[usize],
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

        let cx_rid_start = chart_nums.len() + 1;
        for (i, &cx_num) in chart_ex_nums.iter().enumerate() {
            let rid = format!("rId{}", cx_rid_start + i);
            let target = format!("../charts/chartEx{}.xml", cx_num);
            w.create_element("Relationship")
                .with_attribute(("Id", rid.as_str()))
                .with_attribute(("Type", RT_CHART_EX))
                .with_attribute(("Target", target.as_str()))
                .write_empty()?;
        }

        w.write_event(Event::End(BytesEnd::new("Relationships")))?;
        Ok(())
    })
}
