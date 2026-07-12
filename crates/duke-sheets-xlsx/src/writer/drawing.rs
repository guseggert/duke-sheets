use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_chart::{CellMarker, Chart, ChartEx, DrawingAnchor, EmbeddedImage};
use duke_sheets_core::Drawn;

use super::{write_xml_part, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS, RT_CHART};

const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const NS_MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CX1: &str = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const RT_CHART_EX: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
/// OOXML relationship type for embedded image parts (`xl/media/*`).
pub(super) const RT_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// One image referenced by a sheet's drawing, paired with the global
/// part number used for its filename in `xl/media/imageN.<ext>`. The
/// pair lets the drawing writer and the rels writer agree on which
/// rId points at which media part.
pub(super) struct DrawingImage<'a> {
    pub image: &'a Drawn<'a, EmbeddedImage>,
    pub global_num: usize,
}

pub(super) fn write_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    charts: &[&Drawn<'_, Chart>],
    charts_ex: &[&Drawn<'_, ChartEx>],
    images: &[DrawingImage<'_>],
    raw_drawing_objects: &[&[u8]],
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
            write_two_cell_anchor(w, &chart.object.anchor, &rid, i)?;
        }

        let chartex_rid_start = charts.len() + 1;
        for (i, cx) in charts_ex.iter().enumerate() {
            let rid = format!("rId{}", chartex_rid_start + i);
            let obj_idx = charts.len() + i;
            write_chartex_two_cell_anchor(
                w,
                &cx.object.anchor,
                &rid,
                obj_idx,
                cx.payload.raw_mc_fallback.as_deref(),
            )?;
        }

        // Pictures: rId numbering continues after charts + chartEx.
        let pic_rid_start = charts.len() + charts_ex.len() + 1;
        for (i, drawing_image) in images.iter().enumerate() {
            let rid = format!("rId{}", pic_rid_start + i);
            // Shape ID space follows the chart-frame convention used
            // elsewhere in this file: the first non-DOM object is id=2,
            // then incrementing.
            let shape_idx = charts.len() + charts_ex.len() + i;
            write_picture_anchor(w, drawing_image.image, &rid, shape_idx)?;
        }

        for raw in raw_drawing_objects {
            w.get_mut().write_all(raw)?;
        }

        w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
        Ok(())
    })
}

/// Emit an `<xdr:twoCellAnchor>` / `<xdr:oneCellAnchor>` /
/// `<xdr:absoluteAnchor>` wrapper around an `<xdr:pic>` element,
/// dispatching on the wrapper object's `DrawingAnchor` variant.
fn write_picture_anchor(
    w: &mut XmlWriter,
    image: &Drawn<'_, EmbeddedImage>,
    rid: &str,
    shape_idx: usize,
) -> XlsxResult<()> {
    match &image.object.anchor {
        DrawingAnchor::TwoCell { from, to, edit_as } => {
            let mut tag = BytesStart::new("xdr:twoCellAnchor");
            if let Some(ea) = edit_as {
                let s = match ea {
                    duke_sheets_chart::EditAs::TwoCell => "twoCell",
                    duke_sheets_chart::EditAs::OneCell => "oneCell",
                    duke_sheets_chart::EditAs::Absolute => "absolute",
                };
                tag.push_attribute(("editAs", s));
            }
            w.write_event(Event::Start(tag))?;
            w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
            write_cell_marker(w, from)?;
            w.write_event(Event::End(BytesEnd::new("xdr:from")))?;
            w.write_event(Event::Start(BytesStart::new("xdr:to")))?;
            write_cell_marker(w, to)?;
            w.write_event(Event::End(BytesEnd::new("xdr:to")))?;
            write_picture_element(w, image, rid, shape_idx)?;
            w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
            w.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            w.write_event(Event::Start(BytesStart::new("xdr:oneCellAnchor")))?;
            w.write_event(Event::Start(BytesStart::new("xdr:from")))?;
            write_cell_marker(w, from)?;
            w.write_event(Event::End(BytesEnd::new("xdr:from")))?;
            let cx_s = width_emu.to_string();
            let cy_s = height_emu.to_string();
            w.create_element("xdr:ext")
                .with_attribute(("cx", cx_s.as_str()))
                .with_attribute(("cy", cy_s.as_str()))
                .write_empty()?;
            write_picture_element(w, image, rid, shape_idx)?;
            w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
            w.write_event(Event::End(BytesEnd::new("xdr:oneCellAnchor")))?;
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;
            let x_s = x_emu.to_string();
            let y_s = y_emu.to_string();
            w.create_element("xdr:pos")
                .with_attribute(("x", x_s.as_str()))
                .with_attribute(("y", y_s.as_str()))
                .write_empty()?;
            let cx_s = width_emu.to_string();
            let cy_s = height_emu.to_string();
            w.create_element("xdr:ext")
                .with_attribute(("cx", cx_s.as_str()))
                .with_attribute(("cy", cy_s.as_str()))
                .write_empty()?;
            write_picture_element(w, image, rid, shape_idx)?;
            w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
            w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
        }
    }
    Ok(())
}

/// Emit just the `<xdr:pic>` element (without the surrounding anchor
/// wrapper). Used by all three anchor variants.
fn write_picture_element(
    w: &mut XmlWriter,
    drawn: &Drawn<'_, EmbeddedImage>,
    rid: &str,
    shape_idx: usize,
) -> XlsxResult<()> {
    let image = drawn.payload;
    w.write_event(Event::Start(BytesStart::new("xdr:pic")))?;

    // <xdr:nvPicPr> non-visual picture properties.
    w.write_event(Event::Start(BytesStart::new("xdr:nvPicPr")))?;
    let cnv_id = (shape_idx + 2).to_string();
    let mut cnv_pr = BytesStart::new("xdr:cNvPr");
    cnv_pr.push_attribute(("id", cnv_id.as_str()));
    cnv_pr.push_attribute(("name", drawn.object.meta.name.as_deref().unwrap_or("")));
    if let Some(desc) = drawn.object.meta.alt_text.as_deref() {
        cnv_pr.push_attribute(("descr", desc));
    }
    w.write_event(Event::Empty(cnv_pr))?;
    w.write_event(Event::Start(BytesStart::new("xdr:cNvPicPr")))?;
    let mut pic_locks = BytesStart::new("a:picLocks");
    pic_locks.push_attribute(("noChangeAspect", "1"));
    w.write_event(Event::Empty(pic_locks))?;
    w.write_event(Event::End(BytesEnd::new("xdr:cNvPicPr")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:nvPicPr")))?;

    // <xdr:blipFill> blip reference to the image part.
    w.write_event(Event::Start(BytesStart::new("xdr:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("xmlns:r", NS_DOC_RELS));
    blip.push_attribute(("r:embed", rid));
    w.write_event(Event::Empty(blip))?;
    w.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    w.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    w.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:blipFill")))?;

    // <xdr:spPr> shape properties: xfrm with image-supplied geometry.
    w.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
    let mut xfrm = BytesStart::new("a:xfrm");
    if let Some(rot) = image.rotation {
        xfrm.push_attribute(("rot", rot.to_string().as_str()));
    }
    if image.flip_h {
        xfrm.push_attribute(("flipH", "1"));
    }
    if image.flip_v {
        xfrm.push_attribute(("flipV", "1"));
    }
    w.write_event(Event::Start(xfrm))?;
    w.create_element("a:off")
        .with_attribute(("x", "0"))
        .with_attribute(("y", "0"))
        .write_empty()?;
    let cx_s = image.width_emu.to_string();
    let cy_s = image.height_emu.to_string();
    w.create_element("a:ext")
        .with_attribute(("cx", cx_s.as_str()))
        .with_attribute(("cy", cy_s.as_str()))
        .write_empty()?;
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut prst = BytesStart::new("a:prstGeom");
    prst.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(prst))?;
    w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

    w.write_event(Event::End(BytesEnd::new("xdr:pic")))?;
    Ok(())
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
    image_parts: &[(usize, &'static str)],
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

        let pic_rid_start = chart_nums.len() + chart_ex_nums.len() + 1;
        for (i, (global_num, ext)) in image_parts.iter().enumerate() {
            let rid = format!("rId{}", pic_rid_start + i);
            let target = format!("../media/image{global_num}.{ext}");
            w.create_element("Relationship")
                .with_attribute(("Id", rid.as_str()))
                .with_attribute(("Type", RT_IMAGE))
                .with_attribute(("Target", target.as_str()))
                .write_empty()?;
        }

        w.write_event(Event::End(BytesEnd::new("Relationships")))?;
        Ok(())
    })
}

/// Map an `ImageFormat` to the file extension used in `xl/media/`.
pub(super) fn image_format_extension(fmt: duke_sheets_chart::ImageFormat) -> &'static str {
    use duke_sheets_chart::ImageFormat;
    match fmt {
        ImageFormat::Png => "png",
        ImageFormat::Jpeg => "jpeg",
        ImageFormat::Gif => "gif",
        ImageFormat::Bmp => "bmp",
        ImageFormat::Tiff => "tiff",
        ImageFormat::Emf => "emf",
        ImageFormat::Wmf => "wmf",
        ImageFormat::Svg => "svg",
    }
}

/// Write the raw image bytes as a part `xl/media/imageN.<ext>` inside
/// the zip archive.
pub(super) fn write_media_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    global_num: usize,
    ext: &str,
    bytes: &[u8],
) -> XlsxResult<()> {
    let path = format!("xl/media/image{global_num}.{ext}");
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(bytes)?;
    Ok(())
}
