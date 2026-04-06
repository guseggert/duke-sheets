use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_chart::{
    Axis, AxisCrosses, AxisPosition, AxisType, Chart, ChartAxis, ChartColor, ChartDataTable,
    ChartLine, ChartLines, ChartShapeProperties, ChartType, ChartTypeGroup, CrossBetween,
    DataLabelPosition, DataLabels, DataPoint, DataReference, DataSeries, DisplayBlanksAs,
    ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType, Layout, Legend, LegendPosition,
    Marker, MarkerSymbol, NumberFormat, TickLabelPosition, TickMark, Trendline, TrendlineType,
    UpDownBars, View3D,
};

use super::{write_xml_part, XlsxResult, XmlWriter, NS_DOC_RELS};

const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

pub(super) fn write_chart_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    chart: &Chart,
    chart_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/charts/chart{}.xml", chart_num);
    write_xml_part(zip, &path, |w| write_chart_space(w, chart))
}

fn write_chart_space(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    let mut tag = BytesStart::new("c:chartSpace");
    tag.push_attribute(("xmlns:c", NS_CHART));
    tag.push_attribute(("xmlns:a", NS_DRAWING_MAIN));
    tag.push_attribute(("xmlns:r", NS_DOC_RELS));
    w.write_event(Event::Start(tag))?;

    write_bool_element(w, "c:roundedCorners", chart.rounded_corners)?;

    w.write_event(Event::Start(BytesStart::new("c:chart")))?;

    if let Some(ref title) = chart.title {
        write_title(w, title)?;
    }
    write_bool_element(w, "c:autoTitleDeleted", chart.auto_title_deleted)?;

    if let Some(ref v) = chart.view_3d {
        write_view_3d(w, v)?;
    }

    write_plot_area(w, chart)?;

    if let Some(ref legend) = chart.legend {
        write_legend(w, legend)?;
    }

    if let Some(true) = chart.plot_visible_only {
        w.create_element("c:plotVisOnly")
            .with_attribute(("val", "1"))
            .write_empty()?;
    }
    write_bool_element(w, "c:showDLblsOverMax", chart.show_dlbls_over_max)?;

    if let Some(ref dba) = chart.display_blanks_as {
        let val = match dba {
            DisplayBlanksAs::Gap => "gap",
            DisplayBlanksAs::Span => "span",
            DisplayBlanksAs::Zero => "zero",
        };
        w.create_element("c:dispBlanksAs")
            .with_attribute(("val", val))
            .write_empty()?;
    }

    if let Some(raw) = chart.raw_extensions.get("chart") {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:chart")))?;
    if let Some(ref sp) = chart.shape_properties {
        write_shape_properties(w, sp)?;
    }
    if let Some(raw) = chart.raw_extensions.get("chartSpace") {
        w.get_mut().write_all(raw)?;
    }
    w.write_event(Event::End(BytesEnd::new("c:chartSpace")))?;
    Ok(())
}

fn write_title(w: &mut XmlWriter, text: &str) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:title")))?;
    w.write_event(Event::Start(BytesStart::new("c:tx")))?;
    w.write_event(Event::Start(BytesStart::new("c:rich")))?;
    w.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    w.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
    w.write_event(Event::Start(BytesStart::new("a:p")))?;
    w.write_event(Event::Start(BytesStart::new("a:r")))?;
    w.create_element("a:t")
        .write_text_content(BytesText::new(text))?;
    w.write_event(Event::End(BytesEnd::new("a:r")))?;
    w.write_event(Event::End(BytesEnd::new("a:p")))?;
    w.write_event(Event::End(BytesEnd::new("c:rich")))?;
    w.write_event(Event::End(BytesEnd::new("c:tx")))?;
    w.write_event(Event::End(BytesEnd::new("c:title")))?;
    Ok(())
}

fn write_plot_area(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:plotArea")))?;

    write_layout(w, &chart.layout)?;

    if chart.type_groups.len() >= 2 {
        for group in &chart.type_groups {
            write_combo_chart_type_group(w, group)?;
        }
        write_combo_axes(w, &chart.axes)?;
    } else if chart.type_groups.len() == 1 {
        let group = &chart.type_groups[0];
        let legacy = Chart {
            chart_type: group.chart_type.clone(),
            title: chart.title.clone(),
            series: group.series.clone(),
            category_axis: chart.category_axis.clone(),
            value_axis: chart.value_axis.clone(),
            series_axis: chart.series_axis.clone(),
            legend: chart.legend.clone(),
            data_labels: group.data_labels.clone(),
            view_3d: chart.view_3d.clone(),
            data_table: chart.data_table.clone(),
            display_blanks_as: chart.display_blanks_as.clone(),
            plot_visible_only: chart.plot_visible_only,
            layout: chart.layout.clone(),
            shape_properties: chart.shape_properties.clone(),
            vary_colors: group.vary_colors,
            gap_width: group.gap_width,
            overlap: group.overlap,
            raw_extensions: chart.raw_extensions.clone(),
            is_3d: group.is_3d,
            first_slice_angle: group.first_slice_angle,
            hole_size: group.hole_size,
            bubble_scale: group.bubble_scale,
            show_negative_bubbles: group.show_negative_bubbles,
            radar_style: group.radar_style.clone(),
            auto_title_deleted: chart.auto_title_deleted,
            rounded_corners: chart.rounded_corners,
            show_dlbls_over_max: chart.show_dlbls_over_max,
            wireframe: group.wireframe,
            drop_lines: group.drop_lines.clone(),
            high_low_lines: group.high_low_lines.clone(),
            up_down_bars: group.up_down_bars.clone(),
            series_lines: group.series_lines.clone(),
            anchor: chart.anchor.clone(),
            type_groups: Vec::new(),
            axes: Vec::new(),
            raw_chart_style: None,
            raw_chart_color_style: None,
            show_marker: chart.show_marker,
            of_pie_type: group.of_pie_type,
            split_type: group.split_type,
            split_pos: group.split_pos,
            second_pie_size: group.second_pie_size,
            bar_shape: group.bar_shape,
            floor: group.floor.clone(),
            side_wall: group.side_wall.clone(),
            back_wall: group.back_wall.clone(),
            text_properties: chart.text_properties.clone(),
        };
        write_chart_type_group(w, &legacy)?;
        write_axes(w, &legacy)?;
    } else {
        write_chart_type_group(w, chart)?;
        write_axes(w, chart)?;
    }

    if let Some(ref dt) = chart.data_table {
        write_data_table(w, dt)?;
    }

    if let Some(raw) = chart.raw_extensions.get("plotArea") {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:plotArea")))?;
    Ok(())
}

fn write_layout(w: &mut XmlWriter, layout: &Option<Layout>) -> XlsxResult<()> {
    match layout {
        Some(layout) => {
            if let Some(ref ml) = layout.manual_layout {
                w.write_event(Event::Start(BytesStart::new("c:layout")))?;
                w.write_event(Event::Start(BytesStart::new("c:manualLayout")))?;

                if let Some(x) = ml.x {
                    let s = x.to_string();
                    w.create_element("c:x")
                        .with_attribute(("val", s.as_str()))
                        .write_empty()?;
                }
                if let Some(y) = ml.y {
                    let s = y.to_string();
                    w.create_element("c:y")
                        .with_attribute(("val", s.as_str()))
                        .write_empty()?;
                }
                if let Some(w_val) = ml.width {
                    let s = w_val.to_string();
                    w.create_element("c:w")
                        .with_attribute(("val", s.as_str()))
                        .write_empty()?;
                }
                if let Some(h) = ml.height {
                    let s = h.to_string();
                    w.create_element("c:h")
                        .with_attribute(("val", s.as_str()))
                        .write_empty()?;
                }

                w.write_event(Event::End(BytesEnd::new("c:manualLayout")))?;
                w.write_event(Event::End(BytesEnd::new("c:layout")))?;
            } else {
                w.write_event(Event::Empty(BytesStart::new("c:layout")))?;
            }
        }
        None => {
            w.write_event(Event::Empty(BytesStart::new("c:layout")))?;
        }
    }
    Ok(())
}

fn write_chart_type_group(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    let element = match chart_element_name(&chart.chart_type, chart.is_3d) {
        Some(name) => name,
        None => return Ok(()),
    };
    w.write_event(Event::Start(BytesStart::new(element)))?;

    write_chart_type_props(w, chart)?;

    if let Some(v) = chart.vary_colors {
        w.create_element("c:varyColors")
            .with_attribute(("val", if v { "1" } else { "0" }))
            .write_empty()?;
    }

    let use_xy = uses_xy_data(&chart.chart_type);
    for (i, series) in chart.series.iter().enumerate() {
        write_series(w, series, i, use_xy)?;
    }

    if let Some(ref dl) = chart.data_labels {
        write_data_labels(w, dl)?;
    }

    if let Some(val) = chart.gap_width {
        let s = val.to_string();
        w.create_element("c:gapWidth")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = chart.overlap {
        let s = val.to_string();
        w.create_element("c:overlap")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }


    write_chart_lines_for_legacy(w, chart)?;

    if needs_axes(&chart.chart_type) {
        w.create_element("c:axId")
            .with_attribute(("val", "1"))
            .write_empty()?;
        w.create_element("c:axId")
            .with_attribute(("val", "2"))
            .write_empty()?;
        if matches!(chart.chart_type, ChartType::Surface) {
            w.create_element("c:axId")
                .with_attribute(("val", "3"))
                .write_empty()?;
        }
    }

    if let Some(val) = chart.first_slice_angle {
        let s = val.to_string();
        w.create_element("c:firstSliceAng")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = chart.hole_size {
        let s = val.to_string();
        w.create_element("c:holeSize")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = chart.bubble_scale {
        let s = val.to_string();
        w.create_element("c:bubbleScale")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    write_bool_element(w, "c:showNegBubbles", chart.show_negative_bubbles)?;
    write_bool_element(w, "c:wireframe", chart.wireframe)?;

    if let Some(raw) = chart.raw_extensions.get("chartType") {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new(element)))?;
    Ok(())
}

fn write_combo_chart_type_group(w: &mut XmlWriter, group: &ChartTypeGroup) -> XlsxResult<()> {
    let element = match chart_element_name(&group.chart_type, group.is_3d) {
        Some(name) => name,
        None => return Ok(()),
    };
    w.write_event(Event::Start(BytesStart::new(element)))?;

    write_chart_type_props_for_group(w, &group.chart_type, group.radar_style.as_deref())?;

    if let Some(v) = group.vary_colors {
        w.create_element("c:varyColors")
            .with_attribute(("val", if v { "1" } else { "0" }))
            .write_empty()?;
    }

    let use_xy = uses_xy_data(&group.chart_type);
    for (i, series) in group.series.iter().enumerate() {
        write_series(w, series, i, use_xy)?;
    }

    if let Some(ref dl) = group.data_labels {
        write_data_labels(w, dl)?;
    }

    if let Some(val) = group.gap_width {
        let s = val.to_string();
        w.create_element("c:gapWidth")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = group.overlap {
        let s = val.to_string();
        w.create_element("c:overlap")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }


    write_chart_lines_for_group(w, group)?;

    for ax_id in &group.axis_ids {
        let s = ax_id.to_string();
        w.create_element("c:axId")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(val) = group.first_slice_angle {
        let s = val.to_string();
        w.create_element("c:firstSliceAng")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = group.hole_size {
        let s = val.to_string();
        w.create_element("c:holeSize")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = group.bubble_scale {
        let s = val.to_string();
        w.create_element("c:bubbleScale")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    write_bool_element(w, "c:showNegBubbles", group.show_negative_bubbles)?;
    write_bool_element(w, "c:wireframe", group.wireframe)?;

    if let Some(ref raw) = group.raw_ext {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new(element)))?;
    Ok(())
}

fn write_chart_type_props_for_group(
    w: &mut XmlWriter,
    chart_type: &ChartType,
    radar_style: Option<&str>,
) -> XlsxResult<()> {
    match chart_type {
        ChartType::ColumnClustered => {
            w.create_element("c:barDir").with_attribute(("val", "col")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "clustered")).write_empty()?;
        }
        ChartType::ColumnStacked => {
            w.create_element("c:barDir").with_attribute(("val", "col")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "stacked")).write_empty()?;
        }
        ChartType::ColumnPercentStacked => {
            w.create_element("c:barDir").with_attribute(("val", "col")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "percentStacked")).write_empty()?;
        }
        ChartType::BarClustered => {
            w.create_element("c:barDir").with_attribute(("val", "bar")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "clustered")).write_empty()?;
        }
        ChartType::BarStacked => {
            w.create_element("c:barDir").with_attribute(("val", "bar")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "stacked")).write_empty()?;
        }
        ChartType::BarPercentStacked => {
            w.create_element("c:barDir").with_attribute(("val", "bar")).write_empty()?;
            w.create_element("c:grouping").with_attribute(("val", "percentStacked")).write_empty()?;
        }
        ChartType::Line => {
            w.create_element("c:grouping").with_attribute(("val", "standard")).write_empty()?;
        }
        ChartType::LineStacked => {
            w.create_element("c:grouping").with_attribute(("val", "stacked")).write_empty()?;
        }
        ChartType::Area => {
            w.create_element("c:grouping").with_attribute(("val", "standard")).write_empty()?;
        }
        ChartType::AreaStacked => {
            w.create_element("c:grouping").with_attribute(("val", "stacked")).write_empty()?;
        }
        ChartType::AreaPercentStacked => {
            w.create_element("c:grouping").with_attribute(("val", "percentStacked")).write_empty()?;
        }
        ChartType::ScatterMarkers => {
            w.create_element("c:scatterStyle").with_attribute(("val", "marker")).write_empty()?;
        }
        ChartType::ScatterSmooth => {
            w.create_element("c:scatterStyle").with_attribute(("val", "smoothMarker")).write_empty()?;
        }
        ChartType::ScatterLines => {
            w.create_element("c:scatterStyle").with_attribute(("val", "lineMarker")).write_empty()?;
        }
        ChartType::Radar => {
            let style = radar_style.unwrap_or("marker");
            w.create_element("c:radarStyle").with_attribute(("val", style)).write_empty()?;
        }
        ChartType::Pie
        | ChartType::PieExploded
        | ChartType::Doughnut
        | ChartType::Bubble
        | ChartType::Stock
        | ChartType::Surface => {}
        ChartType::Unsupported(_) => {}
    }
    Ok(())
}

fn write_combo_axes(w: &mut XmlWriter, axes: &[ChartAxis]) -> XlsxResult<()> {
    let mut val_axis_count = 0;
    for chart_axis in axes {
        let tag = match chart_axis.axis.axis_type {
            AxisType::Category => "c:catAx",
            AxisType::Date => "c:dateAx",
            AxisType::Value => "c:valAx",
            AxisType::Series => "c:serAx",
        };
        let default_pos = match chart_axis.axis.axis_type {
            AxisType::Value => {
                val_axis_count += 1;
                if val_axis_count == 1 { "l" } else { "r" }
            }
            _ => "b",
        };
        let mut axis = chart_axis.axis.clone();
        if axis.position == AxisPosition::default() {
            axis.position = match default_pos {
                "l" => AxisPosition::Left,
                "r" => AxisPosition::Right,
                "t" => AxisPosition::Top,
                _ => AxisPosition::Bottom,
            };
        }
        let axis_opt = Some(axis);
        if matches!(chart_axis.axis.axis_type, AxisType::Value) {
            write_val_ax(w, chart_axis.id, default_pos, chart_axis.cross_id, &axis_opt)?;
        } else {
            write_cat_ax(w, tag, chart_axis.id, chart_axis.cross_id, &axis_opt)?;
        }
    }
    Ok(())
}

fn chart_element_name(ct: &ChartType, is_3d: bool) -> Option<&'static str> {
    match ct {
        ChartType::ColumnClustered
        | ChartType::ColumnStacked
        | ChartType::ColumnPercentStacked
        | ChartType::BarClustered
        | ChartType::BarStacked
        | ChartType::BarPercentStacked => Some(if is_3d { "c:bar3DChart" } else { "c:barChart" }),
        ChartType::Line | ChartType::LineStacked => {
            Some(if is_3d { "c:line3DChart" } else { "c:lineChart" })
        }
        ChartType::Pie | ChartType::PieExploded => Some(if is_3d { "c:pie3DChart" } else { "c:pieChart" }),
        ChartType::Doughnut => Some("c:doughnutChart"),
        ChartType::Area | ChartType::AreaStacked | ChartType::AreaPercentStacked => {
            Some(if is_3d { "c:area3DChart" } else { "c:areaChart" })
        }
        ChartType::ScatterMarkers | ChartType::ScatterSmooth | ChartType::ScatterLines => {
            Some("c:scatterChart")
        }
        ChartType::Bubble => Some("c:bubbleChart"),
        ChartType::Radar => Some("c:radarChart"),
        ChartType::Stock => Some("c:stockChart"),
        ChartType::Surface => Some(if is_3d { "c:surface3DChart" } else { "c:surfaceChart" }),
        ChartType::Unsupported(_) => None,
    }
}

fn write_chart_type_props(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    match &chart.chart_type {
        ChartType::ColumnClustered => {
            w.create_element("c:barDir")
                .with_attribute(("val", "col"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "clustered"))
                .write_empty()?;
        }
        ChartType::ColumnStacked => {
            w.create_element("c:barDir")
                .with_attribute(("val", "col"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "stacked"))
                .write_empty()?;
        }
        ChartType::ColumnPercentStacked => {
            w.create_element("c:barDir")
                .with_attribute(("val", "col"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "percentStacked"))
                .write_empty()?;
        }
        ChartType::BarClustered => {
            w.create_element("c:barDir")
                .with_attribute(("val", "bar"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "clustered"))
                .write_empty()?;
        }
        ChartType::BarStacked => {
            w.create_element("c:barDir")
                .with_attribute(("val", "bar"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "stacked"))
                .write_empty()?;
        }
        ChartType::BarPercentStacked => {
            w.create_element("c:barDir")
                .with_attribute(("val", "bar"))
                .write_empty()?;
            w.create_element("c:grouping")
                .with_attribute(("val", "percentStacked"))
                .write_empty()?;
        }
        ChartType::Line => {
            w.create_element("c:grouping")
                .with_attribute(("val", "standard"))
                .write_empty()?;
        }
        ChartType::LineStacked => {
            w.create_element("c:grouping")
                .with_attribute(("val", "stacked"))
                .write_empty()?;
        }
        ChartType::Area => {
            w.create_element("c:grouping")
                .with_attribute(("val", "standard"))
                .write_empty()?;
        }
        ChartType::AreaStacked => {
            w.create_element("c:grouping")
                .with_attribute(("val", "stacked"))
                .write_empty()?;
        }
        ChartType::AreaPercentStacked => {
            w.create_element("c:grouping")
                .with_attribute(("val", "percentStacked"))
                .write_empty()?;
        }
        ChartType::ScatterMarkers => {
            w.create_element("c:scatterStyle")
                .with_attribute(("val", "marker"))
                .write_empty()?;
        }
        ChartType::ScatterSmooth => {
            w.create_element("c:scatterStyle")
                .with_attribute(("val", "smoothMarker"))
                .write_empty()?;
        }
        ChartType::ScatterLines => {
            w.create_element("c:scatterStyle")
                .with_attribute(("val", "lineMarker"))
                .write_empty()?;
        }
        ChartType::Radar => {
            let style = chart.radar_style.as_deref().unwrap_or("marker");
            w.create_element("c:radarStyle")
                .with_attribute(("val", style))
                .write_empty()?;
        }
        ChartType::Pie
        | ChartType::PieExploded
        | ChartType::Doughnut
        | ChartType::Bubble
        | ChartType::Stock
        | ChartType::Surface => {}
        ChartType::Unsupported(_) => {}
    }
    Ok(())
}

fn write_series(
    w: &mut XmlWriter,
    series: &DataSeries,
    idx: usize,
    use_xy: bool,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:ser")))?;

    let idx_str = idx.to_string();
    w.create_element("c:idx")
        .with_attribute(("val", idx_str.as_str()))
        .write_empty()?;
    w.create_element("c:order")
        .with_attribute(("val", idx_str.as_str()))
        .write_empty()?;

    if let Some(ref name) = series.name {
        write_series_tx(w, name)?;
    }

    if let Some(ref sp) = series.shape_properties {
        write_shape_properties(w, sp)?;
    }

    if let Some(ref marker) = series.marker {
        write_marker(w, marker)?;
    }

    if let Some(val) = series.explosion {
        let s = val.to_string();
        w.create_element("c:explosion")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    write_bool_element(w, "c:invertIfNegative", series.invert_if_negative)?;

    for dp in &series.data_points {
        write_data_point(w, dp)?;
    }

    if let Some(ref dl) = series.data_labels {
        write_data_labels(w, dl)?;
    }

    if let Some(ref tl) = series.trendline {
        write_trendline(w, tl)?;
    }

    if let Some(ref eb) = series.error_bars {
        write_error_bars(w, eb)?;
    }

    if let Some(ref cats) = series.categories {
        let tag = if use_xy { "c:xVal" } else { "c:cat" };
        write_data_ref(w, tag, cats)?;
    }

    let val_tag = if use_xy { "c:yVal" } else { "c:val" };
    write_data_ref(w, val_tag, &series.values)?;

    if let Some(smooth) = series.smooth {
        w.create_element("c:smooth")
            .with_attribute(("val", if smooth { "1" } else { "0" }))
            .write_empty()?;
    }

    if let Some(ref raw) = series.raw_ext {
        w.get_mut().write_all(raw)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:ser")))?;
    Ok(())
}

fn write_series_tx(w: &mut XmlWriter, name: &str) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:tx")))?;
    if name.contains('!') {
        w.write_event(Event::Start(BytesStart::new("c:strRef")))?;
        w.create_element("c:f")
            .write_text_content(BytesText::new(name))?;
        w.write_event(Event::End(BytesEnd::new("c:strRef")))?;
    } else {
        w.create_element("c:v")
            .write_text_content(BytesText::new(name))?;
    }
    w.write_event(Event::End(BytesEnd::new("c:tx")))?;
    Ok(())
}

fn write_data_ref(w: &mut XmlWriter, outer_tag: &str, data: &DataReference) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new(outer_tag)))?;

    match data {
        DataReference::Formula(f) => {
            let is_value = outer_tag == "c:val" || outer_tag == "c:yVal";
            let ref_tag = if is_value { "c:numRef" } else { "c:strRef" };
            w.write_event(Event::Start(BytesStart::new(ref_tag)))?;
            w.create_element("c:f")
                .write_text_content(BytesText::new(f))?;
            w.write_event(Event::End(BytesEnd::new(ref_tag)))?;
        }
        DataReference::Numbers(nums) => {
            w.write_event(Event::Start(BytesStart::new("c:numRef")))?;
            w.write_event(Event::Start(BytesStart::new("c:numCache")))?;
            let count = nums.len().to_string();
            w.create_element("c:ptCount")
                .with_attribute(("val", count.as_str()))
                .write_empty()?;
            for (i, v) in nums.iter().enumerate() {
                let idx = i.to_string();
                let val = v.to_string();
                let mut pt = BytesStart::new("c:pt");
                pt.push_attribute(("idx", idx.as_str()));
                w.write_event(Event::Start(pt))?;
                w.create_element("c:v")
                    .write_text_content(BytesText::new(&val))?;
                w.write_event(Event::End(BytesEnd::new("c:pt")))?;
            }
            w.write_event(Event::End(BytesEnd::new("c:numCache")))?;
            w.write_event(Event::End(BytesEnd::new("c:numRef")))?;
        }
        DataReference::Strings(strs) => {
            w.write_event(Event::Start(BytesStart::new("c:strRef")))?;
            w.write_event(Event::Start(BytesStart::new("c:strCache")))?;
            let count = strs.len().to_string();
            w.create_element("c:ptCount")
                .with_attribute(("val", count.as_str()))
                .write_empty()?;
            for (i, s) in strs.iter().enumerate() {
                let idx = i.to_string();
                let mut pt = BytesStart::new("c:pt");
                pt.push_attribute(("idx", idx.as_str()));
                w.write_event(Event::Start(pt))?;
                w.create_element("c:v")
                    .write_text_content(BytesText::new(s))?;
                w.write_event(Event::End(BytesEnd::new("c:pt")))?;
            }
            w.write_event(Event::End(BytesEnd::new("c:strCache")))?;
            w.write_event(Event::End(BytesEnd::new("c:strRef")))?;
        }
    }

    w.write_event(Event::End(BytesEnd::new(outer_tag)))?;
    Ok(())
}

fn write_data_labels(w: &mut XmlWriter, dl: &DataLabels) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:dLbls")))?;

    if let Some(ref nf) = dl.number_format {
        write_number_format(w, nf)?;
    }

    if let Some(ref pos) = dl.position {
        w.create_element("c:dLblPos")
            .with_attribute(("val", data_label_position_val(pos)))
            .write_empty()?;
    }

    write_bool_element(w, "c:showLegendKey", dl.show_legend_key)?;
    write_bool_element(w, "c:showVal", dl.show_value)?;
    write_bool_element(w, "c:showCatName", dl.show_category_name)?;
    write_bool_element(w, "c:showSerName", dl.show_series_name)?;
    write_bool_element(w, "c:showPercent", dl.show_percent)?;
    write_bool_element(w, "c:showBubbleSize", dl.show_bubble_size)?;
    write_bool_element(w, "c:showLeaderLines", dl.show_leader_lines)?;

    if let Some(ref sep) = dl.separator {
        w.create_element("c:separator")
            .write_text_content(BytesText::new(sep))?;
    }


    if let Some(ref ll) = dl.leader_lines {
        write_chart_lines_element(w, "c:leaderLines", ll)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:dLbls")))?;
    Ok(())
}

fn write_data_point(w: &mut XmlWriter, dp: &DataPoint) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:dPt")))?;

    let idx = dp.index.to_string();
    w.create_element("c:idx")
        .with_attribute(("val", idx.as_str()))
        .write_empty()?;

    if let Some(ref m) = dp.marker {
        write_marker(w, m)?;
    }

    if let Some(val) = dp.explosion {
        let s = val.to_string();
        w.create_element("c:explosion")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("c:dPt")))?;
    Ok(())
}

fn write_marker(w: &mut XmlWriter, marker: &Marker) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:marker")))?;

    if let Some(ref sym) = marker.symbol {
        w.create_element("c:symbol")
            .with_attribute(("val", marker_symbol_val(sym)))
            .write_empty()?;
    }

    if let Some(size) = marker.size {
        let s = size.to_string();
        w.create_element("c:size")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("c:marker")))?;
    Ok(())
}

fn write_trendline(w: &mut XmlWriter, tl: &Trendline) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:trendline")))?;

    if let Some(ref name) = tl.name {
        w.create_element("c:name")
            .write_text_content(BytesText::new(name))?;
    }

    let ttype = match tl.trendline_type {
        TrendlineType::Linear => "linear",
        TrendlineType::Exponential => "exp",
        TrendlineType::Logarithmic => "log",
        TrendlineType::MovingAverage => "movingAvg",
        TrendlineType::Polynomial => "poly",
        TrendlineType::Power => "power",
    };
    w.create_element("c:trendlineType")
        .with_attribute(("val", ttype))
        .write_empty()?;

    if let Some(val) = tl.order {
        let s = val.to_string();
        w.create_element("c:order")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(val) = tl.period {
        let s = val.to_string();
        w.create_element("c:period")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(val) = tl.forward {
        let s = val.to_string();
        w.create_element("c:forward")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(val) = tl.backward {
        let s = val.to_string();
        w.create_element("c:backward")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(val) = tl.intercept {
        let s = val.to_string();
        w.create_element("c:intercept")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    write_bool_element(w, "c:dispRSqr", tl.display_r_squared)?;
    write_bool_element(w, "c:dispEq", tl.display_equation)?;

    w.write_event(Event::End(BytesEnd::new("c:trendline")))?;
    Ok(())
}

fn write_error_bars(w: &mut XmlWriter, eb: &ErrorBars) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:errBars")))?;

    let dir = match eb.direction {
        ErrorBarDirection::X => "x",
        ErrorBarDirection::Y => "y",
    };
    w.create_element("c:errDir")
        .with_attribute(("val", dir))
        .write_empty()?;

    let bar_type = match eb.bar_type {
        ErrorBarType::Both => "both",
        ErrorBarType::Minus => "minus",
        ErrorBarType::Plus => "plus",
    };
    w.create_element("c:errBarType")
        .with_attribute(("val", bar_type))
        .write_empty()?;

    let val_type = match eb.value_type {
        ErrorValueType::Custom => "cust",
        ErrorValueType::FixedValue => "fixedVal",
        ErrorValueType::Percentage => "percentage",
        ErrorValueType::StandardDeviation => "stdDev",
        ErrorValueType::StandardError => "stdErr",
    };
    w.create_element("c:errValType")
        .with_attribute(("val", val_type))
        .write_empty()?;

    if let Some(val) = eb.value {
        let s = val.to_string();
        w.create_element("c:val")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    write_bool_element(w, "c:noEndCap", eb.no_end_cap)?;

    w.write_event(Event::End(BytesEnd::new("c:errBars")))?;
    Ok(())
}

fn write_number_format(w: &mut XmlWriter, nf: &NumberFormat) -> XlsxResult<()> {
    let mut el = BytesStart::new("c:numFmt");
    el.push_attribute(("formatCode", nf.format_code.as_str()));
    let linked = nf.source_linked.unwrap_or(false);
    el.push_attribute(("sourceLinked", if linked { "1" } else { "0" }));
    w.write_event(Event::Empty(el))?;
    Ok(())
}

fn write_shape_properties(w: &mut XmlWriter, sp: &ChartShapeProperties) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:spPr")))?;

    if sp.no_fill {
        w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
    } else if let Some(ref color) = sp.solid_fill {
        write_solid_fill(w, color)?;
    }

    if let Some(ref line) = sp.line {
        write_chart_line(w, line)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:spPr")))?;
    Ok(())
}

fn write_solid_fill(w: &mut XmlWriter, color: &ChartColor) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
    w.create_element("a:srgbClr")
        .with_attribute(("val", color.hex.as_str()))
        .write_empty()?;
    w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    Ok(())
}

fn write_chart_line(w: &mut XmlWriter, line: &ChartLine) -> XlsxResult<()> {
    let mut tag = BytesStart::new("a:ln");
    if let Some(width) = line.width {
        let s = width.to_string();
        tag.push_attribute(("w", s.as_str()));
    }
    w.write_event(Event::Start(tag))?;

    if line.no_fill {
        w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
    } else if let Some(ref color) = line.solid_fill {
        write_solid_fill(w, color)?;
    }

    if let Some(ref dash) = line.dash_style {
        w.create_element("a:prstDash")
            .with_attribute(("val", dash.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("a:ln")))?;
    Ok(())
}

fn write_chart_lines_element(w: &mut XmlWriter, tag: &str, cl: &ChartLines) -> XlsxResult<()> {
    if let Some(ref sp) = cl.shape_properties {
        w.write_event(Event::Start(BytesStart::new(tag)))?;
        write_shape_properties(w, sp)?;
        w.write_event(Event::End(BytesEnd::new(tag)))?;
    } else {
        w.write_event(Event::Empty(BytesStart::new(tag)))?;
    }
    Ok(())
}

fn write_up_down_bars(w: &mut XmlWriter, udb: &UpDownBars) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:upDownBars")))?;

    if let Some(val) = udb.gap_width {
        let s = val.to_string();
        w.create_element("c:gapWidth")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }

    if let Some(ref ub) = udb.up_bars {
        write_chart_lines_element(w, "c:upBars", ub)?;
    }

    if let Some(ref db) = udb.down_bars {
        write_chart_lines_element(w, "c:downBars", db)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:upDownBars")))?;
    Ok(())
}

fn write_chart_lines_for_legacy(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    if let Some(ref sl) = chart.series_lines {
        write_chart_lines_element(w, "c:serLines", sl)?;
    }
    if let Some(ref dl) = chart.drop_lines {
        write_chart_lines_element(w, "c:dropLines", dl)?;
    }
    if let Some(ref hl) = chart.high_low_lines {
        write_chart_lines_element(w, "c:hiLowLines", hl)?;
    }
    if let Some(ref udb) = chart.up_down_bars {
        write_up_down_bars(w, udb)?;
    }
    Ok(())
}

fn write_chart_lines_for_group(w: &mut XmlWriter, group: &ChartTypeGroup) -> XlsxResult<()> {
    if let Some(ref sl) = group.series_lines {
        write_chart_lines_element(w, "c:serLines", sl)?;
    }
    if let Some(ref dl) = group.drop_lines {
        write_chart_lines_element(w, "c:dropLines", dl)?;
    }
    if let Some(ref hl) = group.high_low_lines {
        write_chart_lines_element(w, "c:hiLowLines", hl)?;
    }
    if let Some(ref udb) = group.up_down_bars {
        write_up_down_bars(w, udb)?;
    }
    Ok(())
}

fn write_view_3d(w: &mut XmlWriter, v: &View3D) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:view3D")))?;

    if let Some(val) = v.rotate_x {
        let s = val.to_string();
        w.create_element("c:rotX")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = v.rotate_y {
        let s = val.to_string();
        w.create_element("c:rotY")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = v.depth_percent {
        let s = val.to_string();
        w.create_element("c:depthPercent")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = v.height_percent {
        let s = val.to_string();
        w.create_element("c:hPercent")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = v.perspective {
        let s = val.to_string();
        w.create_element("c:perspective")
            .with_attribute(("val", s.as_str()))
            .write_empty()?;
    }
    if let Some(val) = v.right_angle_axes {
        w.create_element("c:rAngAx")
            .with_attribute(("val", if val { "1" } else { "0" }))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("c:view3D")))?;
    Ok(())
}

fn write_data_table(w: &mut XmlWriter, dt: &ChartDataTable) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:dTable")))?;

    write_bool_element(w, "c:showHorzBorder", dt.show_horizontal_border)?;
    write_bool_element(w, "c:showVertBorder", dt.show_vertical_border)?;
    write_bool_element(w, "c:showOutline", dt.show_outline)?;
    write_bool_element(w, "c:showKeys", dt.show_keys)?;

    w.write_event(Event::End(BytesEnd::new("c:dTable")))?;
    Ok(())
}

fn write_axes(w: &mut XmlWriter, chart: &Chart) -> XlsxResult<()> {
    if !needs_axes(&chart.chart_type) {
        return Ok(());
    }

    if uses_two_value_axes(&chart.chart_type) {
        write_val_ax(w, 1, "b", 2, &chart.category_axis)?;
        write_val_ax(w, 2, "l", 1, &chart.value_axis)?;
    } else {
        let cat_tag = match chart.category_axis.as_ref().map(|a| a.axis_type) {
            Some(AxisType::Date) => "c:dateAx",
            _ => "c:catAx",
        };
        write_cat_ax(w, cat_tag, 1, 2, &chart.category_axis)?;
        write_val_ax(w, 2, "l", 1, &chart.value_axis)?;
    }

    if chart.series_axis.is_some() || matches!(chart.chart_type, ChartType::Surface) {
        let ser_ax = chart.series_axis.clone().unwrap_or_else(|| Axis::new());
        write_cat_ax(w, "c:serAx", 3, 1, &Some(ser_ax))?;
    }

    Ok(())
}

fn write_cat_ax(
    w: &mut XmlWriter,
    tag: &str,
    ax_id: u32,
    cross_ax: u32,
    axis: &Option<Axis>,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new(tag)))?;

    let id = ax_id.to_string();
    w.create_element("c:axId")
        .with_attribute(("val", id.as_str()))
        .write_empty()?;

    write_scaling(w, axis)?;

    let delete_val = axis.as_ref().and_then(|a| a.delete).unwrap_or(false);
    w.create_element("c:delete")
        .with_attribute(("val", if delete_val { "1" } else { "0" }))
        .write_empty()?;

    let pos = axis
        .as_ref()
        .map(|a| axis_position_val(&a.position))
        .unwrap_or("b");
    w.create_element("c:axPos")
        .with_attribute(("val", pos))
        .write_empty()?;

    if let Some(ref ax) = axis {
        if ax.major_gridlines {
            w.write_event(Event::Empty(BytesStart::new("c:majorGridlines")))?;
        }
        if ax.minor_gridlines {
            w.write_event(Event::Empty(BytesStart::new("c:minorGridlines")))?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref t) = ax.title {
            write_title(w, t)?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref nf) = ax.number_format {
            write_number_format(w, nf)?;
        }
        if let Some(ref tm) = ax.major_tick_mark {
            w.create_element("c:majorTickMark")
                .with_attribute(("val", tick_mark_val(tm)))
                .write_empty()?;
        }
        if let Some(ref tm) = ax.minor_tick_mark {
            w.create_element("c:minorTickMark")
                .with_attribute(("val", tick_mark_val(tm)))
                .write_empty()?;
        }
        if let Some(ref lp) = ax.label_position {
            w.create_element("c:tickLblPos")
                .with_attribute(("val", tick_label_position_val(lp)))
                .write_empty()?;
        }
    }

    let cross = cross_ax.to_string();
    w.create_element("c:crossAx")
        .with_attribute(("val", cross.as_str()))
        .write_empty()?;

    if let Some(ref ax) = axis {
        if let Some(ref c) = ax.crosses {
            w.create_element("c:crosses")
                .with_attribute(("val", axis_crosses_val(c)))
                .write_empty()?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref cb) = ax.cross_between {
            let val = match cb {
                CrossBetween::Between => "between",
                CrossBetween::MidCat => "midCat",
            };
            w.create_element("c:crossBetween")
                .with_attribute(("val", val))
                .write_empty()?;
        }
        if let Some(val) = ax.major_unit {
            let s = val.to_string();
            w.create_element("c:majorUnit")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
        }
        if let Some(val) = ax.minor_unit {
            let s = val.to_string();
            w.create_element("c:minorUnit")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref sp) = ax.shape_properties {
            write_shape_properties(w, sp)?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref raw) = ax.raw_ext {
            w.get_mut().write_all(raw)?;
        }
    }

    w.write_event(Event::End(BytesEnd::new(tag)))?;
    Ok(())
}

fn write_val_ax(
    w: &mut XmlWriter,
    ax_id: u32,
    default_pos: &str,
    cross_ax: u32,
    axis: &Option<Axis>,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:valAx")))?;

    let id = ax_id.to_string();
    w.create_element("c:axId")
        .with_attribute(("val", id.as_str()))
        .write_empty()?;

    write_scaling(w, axis)?;

    let delete_val = axis.as_ref().and_then(|a| a.delete).unwrap_or(false);
    w.create_element("c:delete")
        .with_attribute(("val", if delete_val { "1" } else { "0" }))
        .write_empty()?;

    let pos = axis
        .as_ref()
        .map(|a| axis_position_val(&a.position))
        .unwrap_or(default_pos);
    w.create_element("c:axPos")
        .with_attribute(("val", pos))
        .write_empty()?;

    if let Some(ref ax) = axis {
        if ax.major_gridlines {
            w.write_event(Event::Empty(BytesStart::new("c:majorGridlines")))?;
        }
        if ax.minor_gridlines {
            w.write_event(Event::Empty(BytesStart::new("c:minorGridlines")))?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref t) = ax.title {
            write_title(w, t)?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref nf) = ax.number_format {
            write_number_format(w, nf)?;
        }
        if let Some(ref tm) = ax.major_tick_mark {
            w.create_element("c:majorTickMark")
                .with_attribute(("val", tick_mark_val(tm)))
                .write_empty()?;
        }
        if let Some(ref tm) = ax.minor_tick_mark {
            w.create_element("c:minorTickMark")
                .with_attribute(("val", tick_mark_val(tm)))
                .write_empty()?;
        }
        if let Some(ref lp) = ax.label_position {
            w.create_element("c:tickLblPos")
                .with_attribute(("val", tick_label_position_val(lp)))
                .write_empty()?;
        }
    }

    let cross = cross_ax.to_string();
    w.create_element("c:crossAx")
        .with_attribute(("val", cross.as_str()))
        .write_empty()?;

    if let Some(ref ax) = axis {
        if let Some(ref c) = ax.crosses {
            w.create_element("c:crosses")
                .with_attribute(("val", axis_crosses_val(c)))
                .write_empty()?;
        }
        if let Some(ref cb) = ax.cross_between {
            let val = match cb {
                CrossBetween::Between => "between",
                CrossBetween::MidCat => "midCat",
            };
            w.create_element("c:crossBetween")
                .with_attribute(("val", val))
                .write_empty()?;
        }
        if let Some(val) = ax.major_unit {
            let s = val.to_string();
            w.create_element("c:majorUnit")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
        }
        if let Some(val) = ax.minor_unit {
            let s = val.to_string();
            w.create_element("c:minorUnit")
                .with_attribute(("val", s.as_str()))
                .write_empty()?;
        }
    }


    if let Some(ref ax) = axis {
        if let Some(ref sp) = ax.shape_properties {
            write_shape_properties(w, sp)?;
        }
    }

    if let Some(ref ax) = axis {
        if let Some(ref raw) = ax.raw_ext {
            w.get_mut().write_all(raw)?;
        }
    }

    w.write_event(Event::End(BytesEnd::new("c:valAx")))?;
    Ok(())
}

fn write_scaling(w: &mut XmlWriter, axis: &Option<Axis>) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:scaling")))?;

    w.create_element("c:orientation")
        .with_attribute(("val", "minMax"))
        .write_empty()?;

    if let Some(ref ax) = axis {
        if let Some(min) = ax.minimum {
            let v = min.to_string();
            w.create_element("c:min")
                .with_attribute(("val", v.as_str()))
                .write_empty()?;
        }
        if let Some(max) = ax.maximum {
            let v = max.to_string();
            w.create_element("c:max")
                .with_attribute(("val", v.as_str()))
                .write_empty()?;
        }
    }

    w.write_event(Event::End(BytesEnd::new("c:scaling")))?;
    Ok(())
}

fn write_legend(w: &mut XmlWriter, legend: &Legend) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("c:legend")))?;

    let pos = match legend.position {
        LegendPosition::Right => "r",
        LegendPosition::Top => "t",
        LegendPosition::Bottom => "b",
        LegendPosition::Left => "l",
        LegendPosition::TopRight => "tr",
    };
    w.create_element("c:legendPos")
        .with_attribute(("val", pos))
        .write_empty()?;

    if legend.overlay {
        w.create_element("c:overlay")
            .with_attribute(("val", "1"))
            .write_empty()?;
    }

    if let Some(ref sp) = legend.shape_properties {
        write_shape_properties(w, sp)?;
    }

    w.write_event(Event::End(BytesEnd::new("c:legend")))?;
    Ok(())
}

fn write_bool_element(w: &mut XmlWriter, tag: &str, val: Option<bool>) -> XlsxResult<()> {
    if let Some(v) = val {
        w.create_element(tag)
            .with_attribute(("val", if v { "1" } else { "0" }))
            .write_empty()?;
    }
    Ok(())
}

fn axis_position_val(pos: &AxisPosition) -> &'static str {
    match pos {
        AxisPosition::Bottom => "b",
        AxisPosition::Top => "t",
        AxisPosition::Left => "l",
        AxisPosition::Right => "r",
    }
}

fn tick_mark_val(tm: &TickMark) -> &'static str {
    match tm {
        TickMark::Cross => "cross",
        TickMark::Inside => "in",
        TickMark::None => "none",
        TickMark::Outside => "out",
    }
}

fn tick_label_position_val(pos: &TickLabelPosition) -> &'static str {
    match pos {
        TickLabelPosition::High => "high",
        TickLabelPosition::Low => "low",
        TickLabelPosition::NextTo => "nextTo",
        TickLabelPosition::None => "none",
    }
}

fn axis_crosses_val(c: &AxisCrosses) -> &'static str {
    match c {
        AxisCrosses::AutoZero => "autoZero",
        AxisCrosses::Min => "min",
        AxisCrosses::Max => "max",
    }
}

fn data_label_position_val(pos: &DataLabelPosition) -> &'static str {
    match pos {
        DataLabelPosition::BestFit => "bestFit",
        DataLabelPosition::Bottom => "b",
        DataLabelPosition::Center => "ctr",
        DataLabelPosition::InsideBase => "inBase",
        DataLabelPosition::InsideEnd => "inEnd",
        DataLabelPosition::Left => "l",
        DataLabelPosition::OutsideEnd => "outEnd",
        DataLabelPosition::Right => "r",
        DataLabelPosition::Top => "t",
    }
}

fn marker_symbol_val(sym: &MarkerSymbol) -> &'static str {
    match sym {
        MarkerSymbol::Circle => "circle",
        MarkerSymbol::Dash => "dash",
        MarkerSymbol::Diamond => "diamond",
        MarkerSymbol::Dot => "dot",
        MarkerSymbol::None => "none",
        MarkerSymbol::Picture => "picture",
        MarkerSymbol::Plus => "plus",
        MarkerSymbol::Square => "square",
        MarkerSymbol::Star => "star",
        MarkerSymbol::Triangle => "triangle",
        MarkerSymbol::X => "x",
        MarkerSymbol::Auto => "auto",
    }
}

fn needs_axes(ct: &ChartType) -> bool {
    !matches!(
        ct,
        ChartType::Pie | ChartType::PieExploded | ChartType::Doughnut
    )
}

fn uses_xy_data(ct: &ChartType) -> bool {
    matches!(
        ct,
        ChartType::ScatterMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterLines
            | ChartType::Bubble
    )
}

fn uses_two_value_axes(ct: &ChartType) -> bool {
    matches!(
        ct,
        ChartType::ScatterMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterLines
            | ChartType::Bubble
    )
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::io::Cursor;

    use duke_sheets_chart::{
        Axis, AxisType, Chart, DrawingAnchor, ChartColor, ChartLine, ChartLines,
        ChartShapeProperties, ChartType, DataLabels, DataReference, DataSeries, UpDownBars,
    };

    use super::write_chart_part;
    use crate::reader::chart::read_chart;

    #[test]
    fn test_extlst_roundtrip() {
        // Build a chart with raw extension data at multiple levels
        let mut chart = Chart::new(ChartType::ColumnClustered);
        let mut ser = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"));
        ser.raw_ext = Some(b"<c:extLst><c:ext uri=\"{ser-ext}\"><serData/></c:ext></c:extLst>".to_vec());
        chart.series.push(ser);

        let mut exts = HashMap::new();
        exts.insert("chartType".to_string(), b"<c:extLst><c:ext uri=\"{ct-ext}\"><ctData/></c:ext></c:extLst>".to_vec());
        exts.insert("plotArea".to_string(), b"<c:extLst><c:ext uri=\"{pa-ext}\"><paData/></c:ext></c:extLst>".to_vec());
        exts.insert("chart".to_string(), b"<c:extLst><c:ext uri=\"{ch-ext}\"><chData/></c:ext></c:extLst>".to_vec());
        exts.insert("chartSpace".to_string(), b"<c:extLst><c:ext uri=\"{cs-ext}\"><csData/></c:ext></c:extLst>".to_vec());
        chart.raw_extensions = exts;

        // Write to a zip
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip_writer = zip::ZipWriter::new(cursor);
        write_chart_part(&mut zip_writer, &chart, 1).unwrap();
        let cursor = zip_writer.finish().unwrap();
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Read back
        let reparsed = read_chart(
            &mut archive,
            "xl/charts/chart1.xml",
            DrawingAnchor::default(),
        )
        .unwrap()
        .unwrap();

        // Verify series extLst survived
        let ser_ext = reparsed.series[0].raw_ext.as_ref().expect("series extLst lost");
        let ser_str = std::str::from_utf8(ser_ext).unwrap();
        assert!(ser_str.contains("ser-ext"), "series ext content lost: {}", ser_str);

        // Verify chart-level extensions survived
        let ct = reparsed.raw_extensions.get("chartType").expect("chartType extLst lost");
        assert!(std::str::from_utf8(ct).unwrap().contains("ct-ext"));

        let pa = reparsed.raw_extensions.get("plotArea").expect("plotArea extLst lost");
        assert!(std::str::from_utf8(pa).unwrap().contains("pa-ext"));

        let ch = reparsed.raw_extensions.get("chart").expect("chart extLst lost");
        assert!(std::str::from_utf8(ch).unwrap().contains("ch-ext"));

        let cs = reparsed.raw_extensions.get("chartSpace").expect("chartSpace extLst lost");
        assert!(std::str::from_utf8(cs).unwrap().contains("cs-ext"));
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

        let reparsed = read_chart(
            &mut archive,
            "xl/charts/chart1.xml",
            DrawingAnchor::default(),
        )
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

        let reparsed = read_chart(
            &mut archive,
            "xl/charts/chart1.xml",
            DrawingAnchor::default(),
        )
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

        let reparsed = read_chart(
            &mut archive,
            "xl/charts/chart1.xml",
            DrawingAnchor::default(),
        )
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
        read_chart(&mut archive, "xl/charts/chart1.xml", DrawingAnchor::default())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn test_roundtrip_drop_lines() {
        let mut chart = Chart::new(ChartType::Line);
        chart.series.push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        chart.drop_lines = Some(ChartLines {
            shape_properties: Some(ChartShapeProperties {
                solid_fill: None,
                no_fill: false,
                line: Some(ChartLine {
                    width: Some(12700),
                    solid_fill: Some(ChartColor { hex: "FF0000".into() }),
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
        chart.series.push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
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
        chart.series.push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3")));
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
                    solid_fill: Some(ChartColor { hex: "00FF00".into() }),
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
            duke_sheets_chart::ChartAxis { id: 1, cross_id: 2, axis: Axis::new() },
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
        let sl = reparsed.type_groups[0].series_lines.as_ref().expect("series_lines lost");
        let sp = sl.shape_properties.as_ref().expect("serLines spPr lost");
        assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "00FF00");
    }

    #[test]
    fn test_roundtrip_up_down_bars() {
        let mut chart = Chart::new(ChartType::Stock);
        chart.series.push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.category_axis = Some(Axis::new());
        chart.value_axis = Some(Axis::new());
        chart.up_down_bars = Some(UpDownBars {
            gap_width: Some(150),
            up_bars: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: Some(ChartColor { hex: "00FF00".into() }),
                    no_fill: false,
                    line: None,
                }),
            }),
            down_bars: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: Some(ChartColor { hex: "FF0000".into() }),
                    no_fill: false,
                    line: None,
                }),
            }),
        });

        let reparsed = roundtrip_chart(&chart);
        let udb = reparsed.up_down_bars.expect("up_down_bars lost");
        assert_eq!(udb.gap_width, Some(150));
        let ub = udb.up_bars.expect("up_bars lost");
        assert_eq!(ub.shape_properties.as_ref().unwrap().solid_fill.as_ref().unwrap().hex, "00FF00");
        let db = udb.down_bars.expect("down_bars lost");
        assert_eq!(db.shape_properties.as_ref().unwrap().solid_fill.as_ref().unwrap().hex, "FF0000");
    }

    #[test]
    fn test_roundtrip_leader_lines() {
        let mut chart = Chart::new(ChartType::Pie);
        chart.series.push(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.data_labels = Some(DataLabels {
            show_value: Some(true),
            leader_lines: Some(ChartLines {
                shape_properties: Some(ChartShapeProperties {
                    solid_fill: None,
                    no_fill: false,
                    line: Some(ChartLine {
                        width: Some(9525),
                        solid_fill: Some(ChartColor { hex: "808080".into() }),
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
}
