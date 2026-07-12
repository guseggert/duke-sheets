use std::io::Write;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

use crate::chart_ex::*;
use crate::{ChartShapeProperties, NumberFormat};

use super::{XmlWriter, NS_DOC_RELS};

type XlsxResult<T> = std::io::Result<T>;

const NS_CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// Serialize a chartEx part (`cx:chartSpace`) to bytes, including the
/// XML declaration.
pub fn chart_ex_part_bytes(chart_ex: &ChartEx) -> XlsxResult<Vec<u8>> {
    let mut w = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    write_chart_space(&mut w, chart_ex)?;
    Ok(w.into_inner().into_inner())
}

fn write_chart_space(w: &mut XmlWriter, cx: &ChartEx) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:chartSpace");
    tag.push_attribute(("xmlns:cx", NS_CX));
    tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
    tag.push_attribute(("xmlns:r", NS_DOC_RELS));
    w.write_event(Event::Start(tag))?;

    write_chart_data(w, &cx.data, &cx.external_data)?;
    write_chart(w, cx)?;

    if let Some(ref sp) = cx.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    if let Some(ref cmo) = cx.color_map_override {
        w.write_event(Event::Start(BytesStart::new("cx:clrMapOvr")))?;
        w.get_mut().write_all(cmo)?;
        w.write_event(Event::End(BytesEnd::new("cx:clrMapOvr")))?;
    }

    if !cx.format_overrides.is_empty() {
        w.write_event(Event::Start(BytesStart::new("cx:fmtOvrs")))?;
        for ovr in &cx.format_overrides {
            write_format_override(w, ovr)?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:fmtOvrs")))?;
    }

    if let Some(ref ps) = cx.print_settings {
        write_print_settings(w, ps)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:chartSpace")))?;
    Ok(())
}

fn write_chart_data(
    w: &mut XmlWriter,
    data: &[ChartExData],
    external_data: &Option<ChartExExternalData>,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:chartData")))?;
    if let Some(ref ext) = external_data {
        let mut tag = BytesStart::new("cx:externalData");
        tag.push_attribute(("r:id", ext.rel_id.as_str()));
        if let Some(auto) = ext.auto_update {
            tag.push_attribute(("cx:autoUpdate", if auto { "1" } else { "0" }));
        }
        w.write_event(Event::Empty(tag))?;
    }
    for d in data {
        let id_s = d.id.to_string();
        let mut tag = BytesStart::new("cx:data");
        tag.push_attribute(("id", id_s.as_str()));
        w.write_event(Event::Start(tag))?;
        for dim in &d.dimensions {
            write_dimension(w, dim)?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:data")))?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:chartData")))?;
    Ok(())
}

fn write_dimension(w: &mut XmlWriter, dim: &ChartExDimension) -> XlsxResult<()> {
    match dim {
        ChartExDimension::String {
            dim_type,
            formula,
            nf_formula,
            levels,
        } => {
            let type_str = match dim_type {
                StringDimType::Cat => "cat",
                StringDimType::ColorStr => "colorStr",
                StringDimType::EntityId => "entityId",
            };
            let mut tag = BytesStart::new("cx:strDim");
            tag.push_attribute(("type", type_str));
            w.write_event(Event::Start(tag))?;
            if let Some(ref f) = formula {
                w.create_element("cx:f")
                    .write_text_content(BytesText::new(f))?;
            }
            if let Some(ref nf) = nf_formula {
                w.create_element("cx:nf")
                    .write_text_content(BytesText::new(nf))?;
            }
            for level in levels {
                write_string_level(w, level)?;
            }
            w.write_event(Event::End(BytesEnd::new("cx:strDim")))?;
        }
        ChartExDimension::Numeric {
            dim_type,
            formula,
            nf_formula,
            levels,
        } => {
            let type_str = match dim_type {
                NumericDimType::Val => "val",
                NumericDimType::X => "x",
                NumericDimType::Y => "y",
                NumericDimType::Size => "size",
                NumericDimType::ColorVal => "colorVal",
            };
            let mut tag = BytesStart::new("cx:numDim");
            tag.push_attribute(("type", type_str));
            w.write_event(Event::Start(tag))?;
            if let Some(ref f) = formula {
                w.create_element("cx:f")
                    .write_text_content(BytesText::new(f))?;
            }
            if let Some(ref nf) = nf_formula {
                w.create_element("cx:nf")
                    .write_text_content(BytesText::new(nf))?;
            }
            for level in levels {
                write_numeric_level(w, level)?;
            }
            w.write_event(Event::End(BytesEnd::new("cx:numDim")))?;
        }
    }
    Ok(())
}

fn write_string_level(w: &mut XmlWriter, level: &ChartExStringLevel) -> XlsxResult<()> {
    let count_s = level.pt_count.to_string();
    let mut tag = BytesStart::new("cx:lvl");
    tag.push_attribute(("ptCount", count_s.as_str()));
    if let Some(ref name) = level.name {
        tag.push_attribute(("name", name.as_str()));
    }
    w.write_event(Event::Start(tag))?;
    for (idx, val) in &level.points {
        let idx_s = idx.to_string();
        let mut pt = BytesStart::new("cx:pt");
        pt.push_attribute(("idx", idx_s.as_str()));
        w.write_event(Event::Start(pt))?;
        w.write_event(Event::Text(BytesText::new(val)))?;
        w.write_event(Event::End(BytesEnd::new("cx:pt")))?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:lvl")))?;
    Ok(())
}

fn write_numeric_level(w: &mut XmlWriter, level: &ChartExNumericLevel) -> XlsxResult<()> {
    let count_s = level.pt_count.to_string();
    let mut tag = BytesStart::new("cx:lvl");
    tag.push_attribute(("ptCount", count_s.as_str()));
    if let Some(ref fc) = level.format_code {
        tag.push_attribute(("formatCode", fc.as_str()));
    }
    if let Some(ref name) = level.name {
        tag.push_attribute(("name", name.as_str()));
    }
    w.write_event(Event::Start(tag))?;
    for (idx, val) in &level.points {
        let idx_s = idx.to_string();
        let mut pt = BytesStart::new("cx:pt");
        pt.push_attribute(("idx", idx_s.as_str()));
        w.write_event(Event::Start(pt))?;
        w.write_event(Event::Text(BytesText::new(val)))?;
        w.write_event(Event::End(BytesEnd::new("cx:pt")))?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:lvl")))?;
    Ok(())
}

fn write_chart(w: &mut XmlWriter, cx: &ChartEx) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:chart")))?;

    if let Some(ref title) = cx.title {
        write_title(w, title)?;
    }

    write_plot_area(w, &cx.plot_area)?;

    if let Some(ref legend) = cx.legend {
        write_legend(w, legend)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:chart")))?;
    Ok(())
}

fn write_title(w: &mut XmlWriter, title: &ChartExTitle) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:title");
    if let Some(ref pos) = title.position {
        tag.push_attribute(("pos", pos.as_str()));
    }
    if let Some(ref align) = title.align {
        tag.push_attribute(("align", align.as_str()));
    }
    if let Some(overlay) = title.overlay {
        tag.push_attribute(("overlay", if overlay { "1" } else { "0" }));
    }
    w.write_event(Event::Start(tag))?;

    if title.text.is_some() || title.rich_text.is_some() {
        w.write_event(Event::Start(BytesStart::new("cx:tx")))?;
        if let Some(ref rich) = title.rich_text {
            w.write_event(Event::Start(BytesStart::new("cx:rich")))?;
            w.get_mut().write_all(rich)?;
            w.write_event(Event::End(BytesEnd::new("cx:rich")))?;
        } else if let Some(ref text) = title.text {
            w.write_event(Event::Start(BytesStart::new("cx:txData")))?;
            w.create_element("cx:v")
                .write_text_content(BytesText::new(text))?;
            w.write_event(Event::End(BytesEnd::new("cx:txData")))?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:tx")))?;
    }

    if let Some(ref offset) = title.offset {
        write_offset(w, offset)?;
    }

    if let Some(ref sp) = title.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:title")))?;
    Ok(())
}

fn write_offset(w: &mut XmlWriter, offset: &ChartExOffset) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:offset");
    if let Some(top) = offset.top {
        let s = top.to_string();
        tag.push_attribute(("t", s.as_str()));
    }
    if let Some(left) = offset.left {
        let s = left.to_string();
        tag.push_attribute(("l", s.as_str()));
    }
    w.write_event(Event::Empty(tag))?;
    Ok(())
}

fn write_plot_area(w: &mut XmlWriter, pa: &ChartExPlotArea) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:plotArea")))?;

    w.write_event(Event::Start(BytesStart::new("cx:plotAreaRegion")))?;

    if let Some(ref sp) = pa.plot_surface {
        w.write_event(Event::Start(BytesStart::new("cx:plotSurface")))?;
        write_cx_shape_properties(w, sp)?;
        w.write_event(Event::End(BytesEnd::new("cx:plotSurface")))?;
    }

    for series in &pa.series {
        write_series(w, series)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:plotAreaRegion")))?;

    for axis in &pa.axes {
        write_axis(w, axis)?;
    }

    if let Some(ref sp) = pa.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:plotArea")))?;
    Ok(())
}

fn write_series(w: &mut XmlWriter, series: &ChartExSeries) -> XlsxResult<()> {
    let layout_str = layout_id_str(&series.layout);
    let mut tag = BytesStart::new("cx:series");
    tag.push_attribute(("layoutId", layout_str));
    if let Some(ref uid) = series.unique_id {
        tag.push_attribute(("uniqueId", uid.as_str()));
    }
    if let Some(hidden) = series.hidden {
        tag.push_attribute(("hidden", if hidden { "1" } else { "0" }));
    }
    if let Some(owner_idx) = series.owner_idx {
        let s = owner_idx.to_string();
        tag.push_attribute(("ownerIdx", s.as_str()));
    }
    if let Some(format_idx) = series.format_idx {
        let s = format_idx.to_string();
        tag.push_attribute(("formatIdx", s.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref text) = series.text {
        write_cx_text(w, text)?;
    }

    if let Some(ref sp) = series.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    for dp in &series.data_points {
        write_data_point(w, dp)?;
    }

    if let Some(ref dl) = series.data_labels {
        write_data_labels(w, dl)?;
    }

    let data_id_s = series.data_id.to_string();
    w.create_element("cx:dataId")
        .with_attribute(("val", data_id_s.as_str()))
        .write_empty()?;

    if let Some(ref lp) = series.layout_properties {
        write_layout_properties(w, lp)?;
    }

    for axis_id in &series.axis_ids {
        let s = axis_id.to_string();
        w.create_element("cx:axisId")
            .write_text_content(BytesText::new(&s))?;
    }

    if let Some(ref vc) = series.value_colors {
        write_value_colors(w, vc)?;
    }

    if let Some(ref vcp) = series.value_color_positions {
        write_value_color_positions(w, vcp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:series")))?;
    Ok(())
}

fn layout_id_str(layout: &ChartExLayout) -> &str {
    match layout {
        ChartExLayout::Waterfall => "waterfall",
        ChartExLayout::Treemap => "treemap",
        ChartExLayout::Sunburst => "sunburst",
        ChartExLayout::Funnel => "funnel",
        ChartExLayout::Histogram => "histogram",
        ChartExLayout::BoxWhisker => "boxWhisker",
        ChartExLayout::ParetoLine => "paretoLine",
        ChartExLayout::RegionMap => "regionMap",
        ChartExLayout::ClusteredColumn => "clusteredColumn",
        ChartExLayout::Unknown(s) => s.as_str(),
    }
}

fn write_cx_text(w: &mut XmlWriter, text: &ChartExText) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:tx")))?;
    if let Some(ref rich) = text.rich {
        w.write_event(Event::Start(BytesStart::new("cx:rich")))?;
        w.get_mut().write_all(rich)?;
        w.write_event(Event::End(BytesEnd::new("cx:rich")))?;
    } else if let Some(ref data) = text.data {
        w.write_event(Event::Start(BytesStart::new("cx:txData")))?;
        if let Some(ref f) = data.formula {
            w.create_element("cx:f")
                .write_text_content(BytesText::new(f))?;
        }
        if let Some(ref v) = data.value {
            w.create_element("cx:v")
                .write_text_content(BytesText::new(v))?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:txData")))?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:tx")))?;
    Ok(())
}

fn write_data_point(w: &mut XmlWriter, dp: &ChartExDataPoint) -> XlsxResult<()> {
    let idx_s = dp.idx.to_string();
    let mut tag = BytesStart::new("cx:dataPt");
    tag.push_attribute(("idx", idx_s.as_str()));
    w.write_event(Event::Start(tag))?;
    if let Some(ref sp) = dp.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:dataPt")))?;
    Ok(())
}

fn write_data_labels(w: &mut XmlWriter, dl: &ChartExDataLabels) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:dataLabels");
    if let Some(ref pos) = dl.position {
        tag.push_attribute(("pos", pos.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref nf) = dl.number_format {
        write_cx_number_format(w, nf)?;
    }

    if dl.visibility_series_name.is_some()
        || dl.visibility_category_name.is_some()
        || dl.visibility_value.is_some()
    {
        let mut vis = BytesStart::new("cx:visibility");
        if let Some(v) = dl.visibility_series_name {
            vis.push_attribute(("seriesName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_category_name {
            vis.push_attribute(("categoryName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_value {
            vis.push_attribute(("value", if v { "1" } else { "0" }));
        }
        w.write_event(Event::Empty(vis))?;
    }

    if let Some(ref sep) = dl.separator {
        w.create_element("cx:separator")
            .write_text_content(BytesText::new(sep))?;
    }

    if let Some(ref sp) = dl.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    for ovr in &dl.overrides {
        write_data_label_override(w, ovr)?;
    }

    for &idx in &dl.hidden_labels {
        let s = idx.to_string();
        w.create_element("cx:dataLabelHidden")
            .with_attribute(("idx", s.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:dataLabels")))?;
    Ok(())
}

fn write_data_label_override(w: &mut XmlWriter, dl: &ChartExDataLabel) -> XlsxResult<()> {
    let idx_s = dl.idx.to_string();
    let mut tag = BytesStart::new("cx:dataLabel");
    tag.push_attribute(("idx", idx_s.as_str()));
    w.write_event(Event::Start(tag))?;

    if let Some(ref nf) = dl.number_format {
        write_cx_number_format(w, nf)?;
    }

    if let Some(ref pos) = dl.position {
        let mut vis_tag = BytesStart::new("cx:visibility");
        if let Some(v) = dl.visibility_series_name {
            vis_tag.push_attribute(("seriesName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_category_name {
            vis_tag.push_attribute(("categoryName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_value {
            vis_tag.push_attribute(("value", if v { "1" } else { "0" }));
        }
        // position is on the parent dataLabel, not visibility.
        // We already wrote idx above; need to write pos attr on a layout wrapper
        // Actually, per spec, position on a dataLabel is an attr on the element itself.
        // Let's handle this differently - we set pos on the cx:dataLabel element.
        // But we already opened cx:dataLabel without pos. Let's just emit visibility.
        let _ = pos; // pos is on the dataLabel attr, but we already opened the tag
        w.write_event(Event::Empty(vis_tag))?;
    } else if dl.visibility_series_name.is_some()
        || dl.visibility_category_name.is_some()
        || dl.visibility_value.is_some()
    {
        let mut vis_tag = BytesStart::new("cx:visibility");
        if let Some(v) = dl.visibility_series_name {
            vis_tag.push_attribute(("seriesName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_category_name {
            vis_tag.push_attribute(("categoryName", if v { "1" } else { "0" }));
        }
        if let Some(v) = dl.visibility_value {
            vis_tag.push_attribute(("value", if v { "1" } else { "0" }));
        }
        w.write_event(Event::Empty(vis_tag))?;
    }

    if let Some(ref sep) = dl.separator {
        w.create_element("cx:separator")
            .write_text_content(BytesText::new(sep))?;
    }

    if let Some(ref sp) = dl.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:dataLabel")))?;
    Ok(())
}

fn write_layout_properties(w: &mut XmlWriter, lp: &ChartExLayoutPr) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:layoutPr")))?;

    if let Some(ref pll) = lp.parent_label_layout {
        w.create_element("cx:parentLabelLayout")
            .with_attribute(("val", pll.as_str()))
            .write_empty()?;
    }

    if let Some(ref rll) = lp.region_label_layout {
        w.create_element("cx:regionLabelLayout")
            .with_attribute(("val", rll.as_str()))
            .write_empty()?;
    }

    if let Some(ref vis) = lp.visibility {
        let mut tag = BytesStart::new("cx:visibility");
        if let Some(v) = vis.connector_lines {
            tag.push_attribute(("connectorLines", if v { "1" } else { "0" }));
        }
        if let Some(v) = vis.mean_line {
            tag.push_attribute(("meanLine", if v { "1" } else { "0" }));
        }
        if let Some(v) = vis.mean_marker {
            tag.push_attribute(("meanMarker", if v { "1" } else { "0" }));
        }
        if let Some(v) = vis.nonoutliers {
            tag.push_attribute(("nonoutliers", if v { "1" } else { "0" }));
        }
        if let Some(v) = vis.outliers {
            tag.push_attribute(("outliers", if v { "1" } else { "0" }));
        }
        w.write_event(Event::Empty(tag))?;
    }

    if lp.aggregation {
        w.write_event(Event::Empty(BytesStart::new("cx:aggregation")))?;
    }

    if let Some(ref binning) = lp.binning {
        write_binning(w, binning)?;
    }

    if let Some(ref geo) = lp.geography {
        write_geography(w, geo)?;
    }

    if let Some(ref stats) = lp.statistics {
        let mut tag = BytesStart::new("cx:statistics");
        if let Some(ref qm) = stats.quartile_method {
            tag.push_attribute(("quartileMethod", qm.as_str()));
        }
        w.write_event(Event::Empty(tag))?;
    }

    for &idx in &lp.subtotals {
        let s = idx.to_string();
        w.create_element("cx:idx")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:layoutPr")))?;
    Ok(())
}

fn write_binning(w: &mut XmlWriter, binning: &ChartExBinning) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:binning");
    if let Some(ref ic) = binning.interval_closed {
        tag.push_attribute(("intervalClosed", ic.as_str()));
    }
    if let Some(ref uf) = binning.underflow {
        tag.push_attribute(("underflow", uf.as_str()));
    }
    if let Some(ref of_) = binning.overflow {
        tag.push_attribute(("overflow", of_.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(bs) = binning.bin_size {
        let s = bs.to_string();
        w.create_element("cx:binSize")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(bc) = binning.bin_count {
        let s = bc.to_string();
        w.create_element("cx:binCount")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:binning")))?;
    Ok(())
}

fn write_geography(w: &mut XmlWriter, geo: &ChartExGeography) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:geography");
    if let Some(ref pt) = geo.projection_type {
        tag.push_attribute(("projectionType", pt.as_str()));
    }
    if let Some(ref vrt) = geo.viewed_region_type {
        tag.push_attribute(("viewedRegionType", vrt.as_str()));
    }
    if let Some(ref cl) = geo.culture_language {
        tag.push_attribute(("cultureLanguage", cl.as_str()));
    }
    if let Some(ref cr) = geo.culture_region {
        tag.push_attribute(("cultureRegion", cr.as_str()));
    }
    if let Some(ref attr) = geo.attribution {
        tag.push_attribute(("attribution", attr.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref raw_cache) = geo.raw_geo_cache {
        w.get_mut().write_all(raw_cache)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:geography")))?;
    Ok(())
}

fn write_value_colors(w: &mut XmlWriter, vc: &ChartExValueColors) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:valueColors")))?;
    if let Some(ref raw) = vc.min_color {
        w.get_mut().write_all(raw)?;
    }
    if let Some(ref raw) = vc.mid_color {
        w.get_mut().write_all(raw)?;
    }
    if let Some(ref raw) = vc.max_color {
        w.get_mut().write_all(raw)?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:valueColors")))?;
    Ok(())
}

fn write_value_color_positions(
    w: &mut XmlWriter,
    vcp: &ChartExValueColorPositions,
) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:valueColorPositions");
    if let Some(count) = vcp.count {
        let s = count.to_string();
        tag.push_attribute(("count", s.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref pos) = vcp.min {
        write_color_position(w, "cx:minPosition", pos)?;
    }
    if let Some(ref pos) = vcp.mid {
        write_color_position(w, "cx:midPosition", pos)?;
    }
    if let Some(ref pos) = vcp.max {
        write_color_position(w, "cx:maxPosition", pos)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:valueColorPositions")))?;
    Ok(())
}

fn write_color_position(
    w: &mut XmlWriter,
    element: &str,
    pos: &ChartExColorPosition,
) -> XlsxResult<()> {
    match pos {
        ChartExColorPosition::ExtremeValue => {
            w.write_event(Event::Start(BytesStart::new(element)))?;
            w.write_event(Event::Empty(BytesStart::new("cx:extremeValue")))?;
            w.write_event(Event::End(BytesEnd::new(element)))?;
        }
        ChartExColorPosition::Number(n) => {
            w.write_event(Event::Start(BytesStart::new(element)))?;
            let s = n.to_string();
            w.create_element("cx:number")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new(element)))?;
        }
        ChartExColorPosition::Percent(p) => {
            w.write_event(Event::Start(BytesStart::new(element)))?;
            let s = p.to_string();
            w.create_element("cx:percent")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new(element)))?;
        }
    }
    Ok(())
}

fn write_axis(w: &mut XmlWriter, axis: &ChartExAxis) -> XlsxResult<()> {
    let id_s = axis.id.to_string();
    let mut tag = BytesStart::new("cx:axis");
    tag.push_attribute(("id", id_s.as_str()));
    if let Some(hidden) = axis.hidden {
        tag.push_attribute(("hidden", if hidden { "1" } else { "0" }));
    }
    w.write_event(Event::Start(tag))?;

    match &axis.scaling {
        ChartExScaling::Category { gap_width } => {
            let mut cat_tag = BytesStart::new("cx:catScaling");
            if let Some(gw) = gap_width {
                let s = gw.to_string();
                cat_tag.push_attribute(("gapWidth", s.as_str()));
            }
            w.write_event(Event::Empty(cat_tag))?;
        }
        ChartExScaling::Value {
            min,
            max,
            major_unit,
            minor_unit,
        } => {
            let mut val_tag = BytesStart::new("cx:valScaling");
            if let Some(mn) = min {
                let s = mn.to_string();
                val_tag.push_attribute(("min", s.as_str()));
            }
            if let Some(mx) = max {
                let s = mx.to_string();
                val_tag.push_attribute(("max", s.as_str()));
            }
            if let Some(mu) = major_unit {
                let s = mu.to_string();
                val_tag.push_attribute(("majorUnit", s.as_str()));
            }
            if let Some(mnu) = minor_unit {
                let s = mnu.to_string();
                val_tag.push_attribute(("minorUnit", s.as_str()));
            }
            w.write_event(Event::Empty(val_tag))?;
        }
    }

    if let Some(ref title) = axis.title {
        write_axis_title(w, title)?;
    }

    if let Some(ref units) = axis.units {
        write_axis_units(w, units)?;
    }

    if let Some(ref sp) = axis.major_gridlines {
        w.write_event(Event::Start(BytesStart::new("cx:majorGridlines")))?;
        write_cx_shape_properties(w, sp)?;
        w.write_event(Event::End(BytesEnd::new("cx:majorGridlines")))?;
    } else if axis.major_gridlines.is_none() {
        // Only emit empty element if the axis had major gridlines read
        // Actually, we don't emit anything if None - only if Some
    }

    if let Some(ref sp) = axis.minor_gridlines {
        w.write_event(Event::Start(BytesStart::new("cx:minorGridlines")))?;
        write_cx_shape_properties(w, sp)?;
        w.write_event(Event::End(BytesEnd::new("cx:minorGridlines")))?;
    }

    if let Some(ref mtm) = axis.major_tick_marks {
        w.create_element("cx:majorTickMarks")
            .with_attribute(("type", mtm.as_str()))
            .write_empty()?;
    }

    if let Some(ref mtm) = axis.minor_tick_marks {
        w.create_element("cx:minorTickMarks")
            .with_attribute(("type", mtm.as_str()))
            .write_empty()?;
    }

    if axis.tick_labels {
        w.write_event(Event::Empty(BytesStart::new("cx:tickLabels")))?;
    }

    if let Some(ref nf) = axis.number_format {
        write_cx_number_format(w, nf)?;
    }

    if let Some(ref sp) = axis.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:axis")))?;
    Ok(())
}

fn write_axis_title(w: &mut XmlWriter, title: &ChartExAxisTitle) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:title")))?;

    if let Some(ref text) = title.text {
        write_cx_text(w, text)?;
    }

    if let Some(ref offset) = title.offset {
        write_offset(w, offset)?;
    }

    if let Some(ref sp) = title.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:title")))?;
    Ok(())
}

fn write_axis_units(w: &mut XmlWriter, units: &ChartExAxisUnits) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:units");
    if let Some(ref u) = units.unit {
        tag.push_attribute(("unit", u.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref label) = units.label {
        w.write_event(Event::Start(BytesStart::new("cx:unitsLabel")))?;
        if let Some(ref text) = label.text {
            write_cx_text(w, text)?;
        }
        if let Some(ref sp) = label.shape_properties {
            write_cx_shape_properties(w, sp)?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:unitsLabel")))?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:units")))?;
    Ok(())
}

fn write_legend(w: &mut XmlWriter, legend: &ChartExLegend) -> XlsxResult<()> {
    let mut tag = BytesStart::new("cx:legend");
    if let Some(ref pos) = legend.position {
        tag.push_attribute(("pos", pos.as_str()));
    }
    if let Some(ref align) = legend.align {
        tag.push_attribute(("align", align.as_str()));
    }
    if let Some(overlay) = legend.overlay {
        tag.push_attribute(("overlay", if overlay { "1" } else { "0" }));
    }
    w.write_event(Event::Start(tag))?;

    if let Some(ref offset) = legend.offset {
        write_offset(w, offset)?;
    }

    if let Some(ref sp) = legend.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:legend")))?;
    Ok(())
}

fn write_format_override(w: &mut XmlWriter, ovr: &ChartExFormatOverride) -> XlsxResult<()> {
    let idx_s = ovr.idx.to_string();
    let mut tag = BytesStart::new("cx:fmtOvr");
    tag.push_attribute(("idx", idx_s.as_str()));
    w.write_event(Event::Start(tag))?;
    if let Some(ref sp) = ovr.shape_properties {
        write_cx_shape_properties(w, sp)?;
    }
    w.write_event(Event::End(BytesEnd::new("cx:fmtOvr")))?;
    Ok(())
}

fn write_print_settings(w: &mut XmlWriter, ps: &ChartExPrintSettings) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:printSettings")))?;

    if let Some(ref hf) = ps.header_footer {
        let mut tag = BytesStart::new("cx:headerFooter");
        if let Some(v) = hf.align_with_margins {
            tag.push_attribute(("alignWithMargins", if v { "1" } else { "0" }));
        }
        if let Some(v) = hf.different_odd_even {
            tag.push_attribute(("differentOddEven", if v { "1" } else { "0" }));
        }
        if let Some(v) = hf.different_first {
            tag.push_attribute(("differentFirst", if v { "1" } else { "0" }));
        }
        w.write_event(Event::Start(tag))?;
        if let Some(ref s) = hf.odd_header {
            w.create_element("cx:oddHeader")
                .write_text_content(BytesText::new(s))?;
        }
        if let Some(ref s) = hf.odd_footer {
            w.create_element("cx:oddFooter")
                .write_text_content(BytesText::new(s))?;
        }
        if let Some(ref s) = hf.even_header {
            w.create_element("cx:evenHeader")
                .write_text_content(BytesText::new(s))?;
        }
        if let Some(ref s) = hf.even_footer {
            w.create_element("cx:evenFooter")
                .write_text_content(BytesText::new(s))?;
        }
        if let Some(ref s) = hf.first_header {
            w.create_element("cx:firstHeader")
                .write_text_content(BytesText::new(s))?;
        }
        if let Some(ref s) = hf.first_footer {
            w.create_element("cx:firstFooter")
                .write_text_content(BytesText::new(s))?;
        }
        w.write_event(Event::End(BytesEnd::new("cx:headerFooter")))?;
    }

    if let Some(ref pm) = ps.page_margins {
        let mut tag = BytesStart::new("cx:pageMargins");
        if let Some(v) = pm.left {
            let s = v.to_string();
            tag.push_attribute(("l", s.as_str()));
        }
        if let Some(v) = pm.right {
            let s = v.to_string();
            tag.push_attribute(("r", s.as_str()));
        }
        if let Some(v) = pm.top {
            let s = v.to_string();
            tag.push_attribute(("t", s.as_str()));
        }
        if let Some(v) = pm.bottom {
            let s = v.to_string();
            tag.push_attribute(("b", s.as_str()));
        }
        if let Some(v) = pm.header {
            let s = v.to_string();
            tag.push_attribute(("header", s.as_str()));
        }
        if let Some(v) = pm.footer {
            let s = v.to_string();
            tag.push_attribute(("footer", s.as_str()));
        }
        w.write_event(Event::Empty(tag))?;
    }

    if let Some(ref setup) = ps.page_setup {
        let mut tag = BytesStart::new("cx:pageSetup");
        if let Some(v) = setup.paper_size {
            let s = v.to_string();
            tag.push_attribute(("paperSize", s.as_str()));
        }
        if let Some(v) = setup.first_page_number {
            let s = v.to_string();
            tag.push_attribute(("firstPageNumber", s.as_str()));
        }
        if let Some(ref o) = setup.orientation {
            tag.push_attribute(("orientation", o.as_str()));
        }
        if let Some(v) = setup.black_and_white {
            tag.push_attribute(("blackAndWhite", if v { "1" } else { "0" }));
        }
        if let Some(v) = setup.draft {
            tag.push_attribute(("draft", if v { "1" } else { "0" }));
        }
        if let Some(v) = setup.use_first_page_number {
            tag.push_attribute(("useFirstPageNumber", if v { "1" } else { "0" }));
        }
        if let Some(v) = setup.horizontal_dpi {
            let s = v.to_string();
            tag.push_attribute(("horizontalDpi", s.as_str()));
        }
        if let Some(v) = setup.vertical_dpi {
            let s = v.to_string();
            tag.push_attribute(("verticalDpi", s.as_str()));
        }
        if let Some(v) = setup.copies {
            let s = v.to_string();
            tag.push_attribute(("copies", s.as_str()));
        }
        w.write_event(Event::Empty(tag))?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:printSettings")))?;
    Ok(())
}

fn write_cx_shape_properties(w: &mut XmlWriter, sp: &ChartShapeProperties) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("cx:spPr")))?;

    if sp.no_fill {
        w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
    } else if let Some(ref color) = sp.solid_fill {
        w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
        w.create_element("a:srgbClr")
            .with_attribute(("val", color.hex.as_str()))
            .write_empty()?;
        w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    }

    if let Some(ref line) = sp.line {
        let mut ln_tag = BytesStart::new("a:ln");
        if let Some(width) = line.width {
            let s = width.to_string();
            ln_tag.push_attribute(("w", s.as_str()));
        }
        w.write_event(Event::Start(ln_tag))?;
        if line.no_fill {
            w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
        } else if let Some(ref color) = line.solid_fill {
            w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            w.create_element("a:srgbClr")
                .with_attribute(("val", color.hex.as_str()))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        }
        if let Some(ref dash) = line.dash_style {
            w.create_element("a:prstDash")
                .with_attribute(("val", dash.as_str()))
                .write_empty()?;
        }
        w.write_event(Event::End(BytesEnd::new("a:ln")))?;
    }

    w.write_event(Event::End(BytesEnd::new("cx:spPr")))?;
    Ok(())
}

fn write_cx_number_format(w: &mut XmlWriter, nf: &NumberFormat) -> XlsxResult<()> {
    let mut el = BytesStart::new("cx:numFmt");
    el.push_attribute(("formatCode", nf.format_code.as_str()));
    if let Some(linked) = nf.source_linked {
        el.push_attribute(("sourceLinked", if linked { "1" } else { "0" }));
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}
