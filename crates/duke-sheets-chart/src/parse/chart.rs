use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::error::{ChartParseError, ChartParseResult};
use crate::{
    Axis, AxisCrosses, AxisPosition, AxisType, Chart, ChartAxis, ChartColor, ChartDataTable,
    ChartLine, ChartLines, ChartShapeProperties, ChartType, ChartTypeGroup, CrossBetween,
    DataLabelPosition, DataLabels, DataPoint, DataReference, DataSeries, DisplayBlanksAs,
    ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType, Layout, Legend, LegendPosition,
    ManualLayout, Marker, MarkerSymbol, NumberFormat, PivotChartSource, TickLabelPosition,
    TickMark, Trendline, TrendlineType, UpDownBars, View3D,
};

/// Parsed chart data before anchor assignment.
struct ParsedChart {
    chart_type: ChartType,
    title: Option<String>,
    series: Vec<DataSeries>,
    category_axis: Option<Axis>,
    value_axis: Option<Axis>,
    series_axis: Option<Axis>,
    legend: Option<Legend>,
    data_labels: Option<DataLabels>,
    view_3d: Option<View3D>,
    data_table: Option<ChartDataTable>,
    display_blanks_as: Option<DisplayBlanksAs>,
    plot_visible_only: Option<bool>,
    layout: Option<Layout>,
    shape_properties: Option<ChartShapeProperties>,
    vary_colors: Option<bool>,
    gap_width: Option<u32>,
    overlap: Option<i32>,
    raw_extensions: HashMap<String, Vec<u8>>,
    is_3d: bool,
    first_slice_angle: Option<u32>,
    hole_size: Option<u32>,
    bubble_scale: Option<u32>,
    show_negative_bubbles: Option<bool>,
    radar_style: Option<String>,
    auto_title_deleted: Option<bool>,
    rounded_corners: Option<bool>,
    pivot_source: Option<PivotChartSource>,
    show_dlbls_over_max: Option<bool>,
    wireframe: Option<bool>,
    drop_lines: Option<ChartLines>,
    high_low_lines: Option<ChartLines>,
    up_down_bars: Option<UpDownBars>,
    series_lines: Option<ChartLines>,
    type_groups: Vec<ChartTypeGroup>,
    axes: Vec<ChartAxis>,
}

/// Parse chart XML from a reader and return a `Chart`.
pub fn parse_chart_xml<R: Read>(reader: R) -> ChartParseResult<Chart> {
    let buf_reader = BufReader::new(reader);
    let mut xml_reader = Reader::from_reader(buf_reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let parsed = parse_chart_xml_inner(&mut xml_reader, &mut buf)?;

    let mut chart = Chart::new(parsed.chart_type);
    chart.title = parsed.title;
    chart.series = parsed.series;
    chart.category_axis = parsed.category_axis;
    chart.value_axis = parsed.value_axis;
    chart.series_axis = parsed.series_axis;
    chart.legend = parsed.legend;
    chart.data_labels = parsed.data_labels;
    chart.view_3d = parsed.view_3d;
    chart.data_table = parsed.data_table;
    chart.display_blanks_as = parsed.display_blanks_as;
    chart.plot_visible_only = parsed.plot_visible_only;
    chart.layout = parsed.layout;
    chart.shape_properties = parsed.shape_properties;
    chart.vary_colors = parsed.vary_colors;
    chart.gap_width = parsed.gap_width;
    chart.overlap = parsed.overlap;
    chart.raw_extensions = parsed.raw_extensions;
    chart.is_3d = parsed.is_3d;
    chart.first_slice_angle = parsed.first_slice_angle;
    chart.hole_size = parsed.hole_size;
    chart.bubble_scale = parsed.bubble_scale;
    chart.show_negative_bubbles = parsed.show_negative_bubbles;
    chart.radar_style = parsed.radar_style;
    chart.auto_title_deleted = parsed.auto_title_deleted;
    chart.rounded_corners = parsed.rounded_corners;
    chart.pivot_source = parsed.pivot_source;
    chart.show_dlbls_over_max = parsed.show_dlbls_over_max;
    chart.wireframe = parsed.wireframe;
    chart.drop_lines = parsed.drop_lines;
    chart.high_low_lines = parsed.high_low_lines;
    chart.up_down_bars = parsed.up_down_bars;
    chart.series_lines = parsed.series_lines;
    chart.type_groups = parsed.type_groups;
    chart.axes = parsed.axes;

    Ok(chart)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SpPrContext {
    None,
    Series,
    DataPoint,
    CatAxis,
    ValAxis,
    MajorGridlines,
    MinorGridlines,
    ChartSpace,
    Legend,
    DropLines,
    HiLowLines,
    SerLines,
    UpBars,
    DownBars,
    LeaderLines,
}

fn parse_chart_xml_inner<R: Read>(
    xml_reader: &mut Reader<BufReader<R>>,
    buf: &mut Vec<u8>,
) -> ChartParseResult<ParsedChart> {
    let mut result = ParsedChart {
        chart_type: ChartType::Unsupported("unknown".into()),
        title: None,
        series: Vec::new(),
        category_axis: None,
        value_axis: None,
        series_axis: None,
        legend: None,
        data_labels: None,
        view_3d: None,
        data_table: None,
        display_blanks_as: None,
        plot_visible_only: None,
        layout: None,
        shape_properties: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        raw_extensions: HashMap::new(),
        is_3d: false,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        auto_title_deleted: None,
        rounded_corners: None,
        pivot_source: None,
        show_dlbls_over_max: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        up_down_bars: None,
        series_lines: None,
        type_groups: Vec::new(),
        axes: Vec::new(),
    };

    // Nesting tracking
    let mut in_chart = false;
    let mut in_plot_area = false;
    let mut in_chart_type_element = false;
    let mut chart_type_tag: Option<String> = None;
    let mut bar_dir: Option<String> = None;
    let mut grouping: Option<String> = None;
    let mut scatter_style: Option<String> = None;
    let mut vary_colors: Option<bool> = None;
    let mut gap_width: Option<u32> = None;
    let mut overlap: Option<i32> = None;
    let mut first_slice_angle: Option<u32> = None;
    let mut hole_size: Option<u32> = None;
    let mut bubble_scale: Option<u32> = None;
    let mut show_negative_bubbles: Option<bool> = None;
    let mut radar_style: Option<String> = None;
    let mut wireframe: Option<bool> = None;
    let mut in_chart_space = false;
    let mut in_pivot_source = false;
    let mut in_pivot_source_name = false;
    let mut pivot_source_name = String::new();
    let mut pivot_source_format_id: Option<u32> = None;

    // Per-group accumulators for combo chart support
    let mut group_series: Vec<DataSeries> = Vec::new();
    let mut group_data_labels: Option<DataLabels> = None;
    let mut group_axis_ids: Vec<u32> = Vec::new();
    let mut group_raw_ext: Option<Vec<u8>> = None;
    let mut group_drop_lines: Option<ChartLines> = None;
    let mut group_high_low_lines: Option<ChartLines> = None;
    let mut group_series_lines: Option<ChartLines> = None;
    let mut group_up_down_bars: Option<UpDownBars> = None;

    // Chart lines / up-down bars state
    let mut in_drop_lines = false;
    let mut in_hi_low_lines = false;
    let mut in_ser_lines = false;
    let mut in_up_down_bars = false;
    let mut in_up_bars = false;
    let mut in_down_bars = false;
    let mut in_leader_lines = false;
    let mut up_down_bars_gap_width: Option<u32> = None;
    let mut up_bars_sp: Option<ChartShapeProperties> = None;
    let mut down_bars_sp: Option<ChartShapeProperties> = None;
    let mut had_up_bars = false;
    let mut had_down_bars = false;

    // Title parsing state
    let mut in_chart_title = false;
    let mut in_title_tx = false;
    let mut in_title_rich = false;
    let mut in_title_p = false;
    let mut in_title_r = false;
    let mut in_title_t = false;
    let mut title_text = String::new();
    let mut in_title_str_ref = false;
    let mut in_title_str_ref_f = false;
    // Track depth to distinguish chart-level title from axis titles
    let mut title_depth = 0u32;

    // Series parsing state
    let mut in_ser = false;
    let mut in_ser_tx = false;
    let mut in_ser_tx_str_ref = false;
    let mut in_ser_tx_str_ref_f = false;
    let mut in_ser_tx_v = false;
    let mut ser_name: Option<String> = None;
    let mut in_ser_val = false;
    let mut in_ser_yval = false;
    let mut in_ser_val_num_ref = false;
    let mut in_ser_val_num_ref_f = false;
    let mut ser_val_formula: Option<String> = None;
    let mut in_ser_val_num_cache = false;
    let mut in_ser_val_pt = false;
    let mut in_ser_val_pt_v = false;
    let mut ser_val_cache: Vec<f64> = Vec::new();
    let mut in_ser_cat = false;
    let mut in_ser_xval = false;
    let mut in_ser_cat_ref = false;
    let mut in_ser_cat_ref_f = false;
    let mut ser_cat_formula: Option<String> = None;

    // Series extras
    let mut ser_smooth: Option<bool> = None;
    let mut ser_explosion: Option<u32> = None;
    let mut ser_data_labels: Option<DataLabels> = None;
    let mut ser_marker: Option<Marker> = None;
    let mut ser_trendline: Option<Trendline> = None;
    let mut ser_error_bars: Option<ErrorBars> = None;
    let mut ser_shape_properties: Option<ChartShapeProperties> = None;
    let mut ser_raw_ext: Option<Vec<u8>> = None;
    let mut ser_invert_if_negative: Option<bool> = None;

    // Data labels state
    let mut in_dlbls = false;
    let mut in_dlbls_separator = false;
    let mut dlbls = DataLabels::default();

    // Data point state
    let mut in_dpt = false;
    let mut dpt_index: u32 = 0;
    let mut dpt_explosion: Option<u32> = None;
    let mut dpt_marker: Option<Marker> = None;
    let mut dpt_shape_properties: Option<ChartShapeProperties> = None;
    let mut ser_data_points: Vec<DataPoint> = Vec::new();

    // Marker state
    let mut in_marker = false;
    let mut marker_symbol: Option<MarkerSymbol> = None;
    let mut marker_size: Option<u8> = None;

    // Trendline state
    let mut in_trendline = false;
    let mut in_trendline_name = false;
    let mut trendline_type: Option<TrendlineType> = None;
    let mut trendline_name: Option<String> = None;
    let mut trendline_order: Option<u32> = None;
    let mut trendline_period: Option<u32> = None;
    let mut trendline_forward: Option<f64> = None;
    let mut trendline_backward: Option<f64> = None;
    let mut trendline_intercept: Option<f64> = None;
    let mut trendline_disp_r_sqr: Option<bool> = None;
    let mut trendline_disp_eq: Option<bool> = None;

    // Error bars state
    let mut in_err_bars = false;
    let mut err_dir: Option<ErrorBarDirection> = None;
    let mut err_bar_type: Option<ErrorBarType> = None;
    let mut err_val_type: Option<ErrorValueType> = None;
    let mut err_val: Option<f64> = None;
    let mut err_no_end_cap: Option<bool> = None;

    // Axis parsing state
    let mut in_cat_ax = false;
    let mut in_val_ax = false;
    let mut in_ser_ax = false;
    let mut is_date_ax = false;
    let mut in_ax_title = false;
    let mut in_ax_title_tx = false;
    let mut in_ax_title_rich = false;
    let mut in_ax_title_p = false;
    let mut in_ax_title_r = false;
    let mut in_ax_title_t = false;
    let mut ax_title_text = String::new();
    let mut in_ax_scaling = false;
    let mut ax_min: Option<f64> = None;
    let mut ax_max: Option<f64> = None;
    let mut ax_number_format: Option<NumberFormat> = None;
    let mut ax_major_gridlines = false;
    let mut ax_minor_gridlines = false;
    let mut in_ax_major_gridlines = false;
    let mut in_ax_minor_gridlines = false;
    let mut ax_major_gridlines_shape_properties: Option<ChartShapeProperties> = None;
    let mut ax_minor_gridlines_shape_properties: Option<ChartShapeProperties> = None;
    let mut ax_major_tick_mark: Option<TickMark> = None;
    let mut ax_minor_tick_mark: Option<TickMark> = None;
    let mut ax_label_position: Option<TickLabelPosition> = None;
    let mut ax_delete: Option<bool> = None;
    let mut ax_crosses: Option<AxisCrosses> = None;
    let mut ax_cross_between: Option<CrossBetween> = None;
    let mut ax_position: Option<AxisPosition> = None;
    let mut ax_major_unit: Option<f64> = None;
    let mut ax_minor_unit: Option<f64> = None;
    let mut ax_shape_properties: Option<ChartShapeProperties> = None;
    let mut ax_raw_ext: Option<Vec<u8>> = None;
    let mut ax_id: Option<u32> = None;
    let mut ax_cross_id: Option<u32> = None;

    // Legend parsing state
    let mut in_legend = false;
    let mut legend_pos: Option<LegendPosition> = None;
    let mut legend_overlay: Option<bool> = None;
    let mut legend_shape_properties: Option<ChartShapeProperties> = None;

    // View 3D state
    let mut in_view_3d = false;
    let mut view_3d = View3D::default();

    // Layout state
    let mut in_layout = false;
    let mut in_manual_layout = false;
    let mut had_manual_layout = false;
    let mut manual_layout = ManualLayout::default();

    // Shape properties state
    let mut in_sp_pr = false;
    let mut sp_pr_depth = 0u32;
    let mut sp_solid_fill: Option<ChartColor> = None;
    let mut sp_no_fill = false;
    let mut sp_line: Option<ChartLine> = None;
    let mut in_sp_ln = false;
    let mut sp_ln_width: Option<i64> = None;
    let mut sp_ln_solid_fill: Option<ChartColor> = None;
    let mut sp_ln_no_fill = false;
    let mut sp_ln_dash: Option<String> = None;
    let mut sp_pr_context: SpPrContext = SpPrContext::None;

    // Data table state
    let mut in_d_table = false;
    let mut d_table = ChartDataTable::default();

    loop {
        match xml_reader.read_event_into(buf) {
            Ok(Event::Start(e)) => {
                let local = e.name().local_name();
                let tag = local.as_ref();
                match tag {
                    b"chartSpace" => in_chart_space = true,
                    b"pivotSource" if in_chart_space && !in_chart => {
                        in_pivot_source = true;
                        pivot_source_name.clear();
                        pivot_source_format_id = None;
                    }
                    b"name" if in_pivot_source => in_pivot_source_name = true,
                    b"chart" if !in_chart => in_chart = true,
                    b"plotArea" if in_chart => in_plot_area = true,
                    b"title"
                        if in_chart
                            && !in_plot_area
                            && !in_chart_title
                            && !in_cat_ax
                            && !in_val_ax
                            && !in_ser_ax =>
                    {
                        in_chart_title = true;
                        title_depth = 1;
                        title_text.clear();
                    }
                    b"tx" if in_chart_title && title_depth == 1 => in_title_tx = true,
                    b"rich" if in_title_tx => in_title_rich = true,
                    b"p" if in_title_rich => in_title_p = true,
                    b"r" if in_title_p => in_title_r = true,
                    b"t" if in_title_r => in_title_t = true,
                    b"strRef" if in_title_tx => in_title_str_ref = true,
                    b"f" if in_title_str_ref => in_title_str_ref_f = true,
                    // View 3D
                    b"view3D" if in_chart && !in_plot_area => {
                        in_view_3d = true;
                        view_3d = View3D::default();
                    }
                    // Chart type elements in plotArea
                    b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
                    | b"pie3DChart" | b"doughnutChart" | b"areaChart" | b"area3DChart"
                    | b"scatterChart" | b"bubbleChart" | b"radarChart" | b"stockChart"
                    | b"surfaceChart" | b"surface3DChart" | b"ofPieChart"
                        if in_plot_area && !in_chart_type_element =>
                    {
                        in_chart_type_element = true;
                        let tag_str = std::str::from_utf8(tag).unwrap_or("unknown");
                        chart_type_tag = Some(tag_str.to_string());
                        bar_dir = None;
                        grouping = None;
                        scatter_style = None;
                        group_series.clear();
                        group_data_labels = None;
                        group_axis_ids.clear();
                        group_raw_ext = None;
                    }
                    // Layout
                    b"layout" if in_plot_area && !in_chart_type_element => {
                        in_layout = true;
                        had_manual_layout = false;
                    }
                    b"manualLayout" if in_layout => {
                        in_manual_layout = true;
                        had_manual_layout = true;
                        manual_layout = ManualLayout::default();
                    }
                    // Data table
                    b"dTable" if in_plot_area && !in_chart_type_element => {
                        in_d_table = true;
                        d_table = ChartDataTable::default();
                    }
                    // Data labels (chart-level or series-level)
                    b"dLbls" if in_chart_type_element && !in_dlbls => {
                        in_dlbls = true;
                        dlbls = DataLabels::default();
                    }
                    b"separator" if in_dlbls => in_dlbls_separator = true,
                    b"numFmt" if in_dlbls => {
                        dlbls.number_format = Some(parse_num_fmt(&e));
                    }
                    // Chart lines and up-down bars
                    b"dropLines" if in_chart_type_element && !in_ser => {
                        in_drop_lines = true;
                    }
                    b"hiLowLines" if in_chart_type_element && !in_ser => {
                        in_hi_low_lines = true;
                    }
                    b"serLines" if in_chart_type_element && !in_ser => {
                        in_ser_lines = true;
                    }
                    b"upDownBars" if in_chart_type_element && !in_ser => {
                        in_up_down_bars = true;
                        up_down_bars_gap_width = None;
                        up_bars_sp = None;
                        down_bars_sp = None;
                        had_up_bars = false;
                        had_down_bars = false;
                    }
                    b"upBars" if in_up_down_bars => {
                        in_up_bars = true;
                        had_up_bars = true;
                    }
                    b"downBars" if in_up_down_bars => {
                        in_down_bars = true;
                        had_down_bars = true;
                    }
                    b"leaderLines" if in_dlbls => in_leader_lines = true,
                    b"ser" if in_chart_type_element => {
                        in_ser = true;
                        ser_name = None;
                        ser_val_formula = None;
                        ser_val_cache.clear();
                        ser_cat_formula = None;
                        ser_smooth = None;
                        ser_explosion = None;
                        ser_data_labels = None;
                        ser_marker = None;
                        ser_trendline = None;
                        ser_error_bars = None;
                        ser_data_points.clear();
                        ser_shape_properties = None;
                        ser_raw_ext = None;
                        ser_invert_if_negative = None;
                    }
                    b"tx" if in_ser => in_ser_tx = true,
                    b"strRef" if in_ser_tx => in_ser_tx_str_ref = true,
                    b"f" if in_ser_tx_str_ref => in_ser_tx_str_ref_f = true,
                    b"v" if in_ser_tx && !in_ser_tx_str_ref => in_ser_tx_v = true,
                    // Data points
                    b"dPt" if in_ser => {
                        in_dpt = true;
                        dpt_index = 0;
                        dpt_explosion = None;
                        dpt_marker = None;
                        dpt_shape_properties = None;
                    }
                    // Trendline
                    b"trendline" if in_ser => {
                        in_trendline = true;
                        trendline_type = None;
                        trendline_name = None;
                        trendline_order = None;
                        trendline_period = None;
                        trendline_forward = None;
                        trendline_backward = None;
                        trendline_intercept = None;
                        trendline_disp_r_sqr = None;
                        trendline_disp_eq = None;
                    }
                    b"name" if in_trendline => in_trendline_name = true,
                    // Error bars
                    b"errBars" if in_ser => {
                        in_err_bars = true;
                        err_dir = None;
                        err_bar_type = None;
                        err_val_type = None;
                        err_val = None;
                        err_no_end_cap = None;
                    }
                    // Marker (series or data point level)
                    b"marker" if in_ser && !in_dlbls && !in_trendline && !in_err_bars => {
                        in_marker = true;
                        marker_symbol = None;
                        marker_size = None;
                    }
                    b"val" if in_err_bars => err_val = get_val_f64(&e),
                    b"val" if in_ser && !in_err_bars => in_ser_val = true,
                    b"yVal" if in_ser => in_ser_yval = true,
                    b"numRef" if in_ser_val || in_ser_yval => in_ser_val_num_ref = true,
                    b"f" if in_ser_val_num_ref => in_ser_val_num_ref_f = true,
                    b"numCache" if in_ser_val_num_ref => in_ser_val_num_cache = true,
                    b"pt" if in_ser_val_num_cache => in_ser_val_pt = true,
                    b"v" if in_ser_val_pt => in_ser_val_pt_v = true,
                    b"cat" if in_ser => in_ser_cat = true,
                    b"xVal" if in_ser => in_ser_xval = true,
                    b"strRef" | b"numRef" if in_ser_cat || in_ser_xval => {
                        in_ser_cat_ref = true;
                    }
                    b"f" if in_ser_cat_ref => in_ser_cat_ref_f = true,
                    // Axis elements
                    b"catAx" | b"dateAx" if in_plot_area => {
                        if tag == b"catAx" {
                        } else {
                        }
                        in_cat_ax = true;
                        is_date_ax = tag == b"dateAx";
                        ax_title_text.clear();
                        ax_min = None;
                        ax_max = None;
                        ax_number_format = None;
                        ax_major_gridlines = false;
                        ax_minor_gridlines = false;
                        in_ax_major_gridlines = false;
                        in_ax_minor_gridlines = false;
                        ax_major_gridlines_shape_properties = None;
                        ax_minor_gridlines_shape_properties = None;
                        ax_major_tick_mark = None;
                        ax_minor_tick_mark = None;
                        ax_label_position = None;
                        ax_delete = None;
                        ax_crosses = None;
                        ax_cross_between = None;
                        ax_position = None;
                        ax_major_unit = None;
                        ax_minor_unit = None;
                        ax_shape_properties = None;
                        ax_raw_ext = None;
                        ax_id = None;
                        ax_cross_id = None;
                    }
                    b"valAx" if in_plot_area => {
                        in_val_ax = true;
                        ax_title_text.clear();
                        ax_min = None;
                        ax_max = None;
                        ax_number_format = None;
                        ax_major_gridlines = false;
                        ax_minor_gridlines = false;
                        in_ax_major_gridlines = false;
                        in_ax_minor_gridlines = false;
                        ax_major_gridlines_shape_properties = None;
                        ax_minor_gridlines_shape_properties = None;
                        ax_major_tick_mark = None;
                        ax_minor_tick_mark = None;
                        ax_label_position = None;
                        ax_delete = None;
                        ax_crosses = None;
                        ax_cross_between = None;
                        ax_position = None;
                        ax_major_unit = None;
                        ax_minor_unit = None;
                        ax_shape_properties = None;
                        ax_raw_ext = None;
                        ax_id = None;
                        ax_cross_id = None;
                    }
                    b"serAx" if in_plot_area => {
                        in_ser_ax = true;
                        ax_title_text.clear();
                        ax_min = None;
                        ax_max = None;
                        ax_number_format = None;
                        ax_major_gridlines = false;
                        ax_minor_gridlines = false;
                        in_ax_major_gridlines = false;
                        in_ax_minor_gridlines = false;
                        ax_major_gridlines_shape_properties = None;
                        ax_minor_gridlines_shape_properties = None;
                        ax_major_tick_mark = None;
                        ax_minor_tick_mark = None;
                        ax_label_position = None;
                        ax_delete = None;
                        ax_crosses = None;
                        ax_cross_between = None;
                        ax_position = None;
                        ax_major_unit = None;
                        ax_minor_unit = None;
                        ax_shape_properties = None;
                        ax_raw_ext = None;
                        ax_id = None;
                        ax_cross_id = None;
                    }
                    b"title" if in_cat_ax || in_val_ax || in_ser_ax => in_ax_title = true,
                    b"tx" if in_ax_title => in_ax_title_tx = true,
                    b"rich" if in_ax_title_tx => in_ax_title_rich = true,
                    b"p" if in_ax_title_rich => in_ax_title_p = true,
                    b"r" if in_ax_title_p => in_ax_title_r = true,
                    b"t" if in_ax_title_r => in_ax_title_t = true,
                    b"scaling" if in_cat_ax || in_val_ax || in_ser_ax => {
                        in_ax_scaling = true;
                    }
                    b"majorGridlines" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_major_gridlines = true;
                        in_ax_major_gridlines = true;
                    }
                    b"minorGridlines" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_minor_gridlines = true;
                        in_ax_minor_gridlines = true;
                    }
                    b"numFmt"
                        if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title && !in_dlbls =>
                    {
                        ax_number_format = Some(parse_num_fmt(&e));
                    }
                    // Legend
                    b"legend" if in_chart && !in_plot_area => {
                        in_legend = true;
                        legend_pos = None;
                    }
                    b"spPr" if !in_sp_pr => {
                        in_sp_pr = true;
                        sp_pr_depth = 1;
                        sp_solid_fill = None;
                        sp_no_fill = false;
                        sp_line = None;
                        in_sp_ln = false;
                        sp_ln_width = None;
                        sp_ln_solid_fill = None;
                        sp_ln_no_fill = false;
                        sp_ln_dash = None;
                        if in_drop_lines {
                            sp_pr_context = SpPrContext::DropLines;
                        } else if in_hi_low_lines {
                            sp_pr_context = SpPrContext::HiLowLines;
                        } else if in_ser_lines {
                            sp_pr_context = SpPrContext::SerLines;
                        } else if in_up_bars {
                            sp_pr_context = SpPrContext::UpBars;
                        } else if in_down_bars {
                            sp_pr_context = SpPrContext::DownBars;
                        } else if in_leader_lines {
                            sp_pr_context = SpPrContext::LeaderLines;
                        } else if in_dpt && !in_marker && !in_dlbls {
                            sp_pr_context = SpPrContext::DataPoint;
                        } else if in_ser
                            && !in_dpt
                            && !in_trendline
                            && !in_err_bars
                            && !in_marker
                            && !in_dlbls
                        {
                            sp_pr_context = SpPrContext::Series;
                        } else if in_ax_major_gridlines {
                            sp_pr_context = SpPrContext::MajorGridlines;
                        } else if in_ax_minor_gridlines {
                            sp_pr_context = SpPrContext::MinorGridlines;
                        } else if in_cat_ax || in_ser_ax {
                            sp_pr_context = SpPrContext::CatAxis;
                        } else if in_val_ax {
                            sp_pr_context = SpPrContext::ValAxis;
                        } else if in_legend {
                            sp_pr_context = SpPrContext::Legend;
                        } else if in_chart_space && !in_chart && !in_plot_area {
                            sp_pr_context = SpPrContext::ChartSpace;
                        } else {
                            sp_pr_context = SpPrContext::None;
                        }
                    }
                    b"ln" if in_sp_pr => {
                        sp_pr_depth += 1;
                        in_sp_ln = true;
                        sp_ln_width = None;
                        sp_ln_solid_fill = None;
                        sp_ln_no_fill = false;
                        sp_ln_dash = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"w" {
                                sp_ln_width = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse::<i64>().ok());
                            }
                        }
                    }
                    b"solidFill" if in_sp_pr && !in_sp_ln => sp_pr_depth += 1,
                    b"solidFill" if in_sp_ln => sp_pr_depth += 1,
                    b"extLst" => {
                        if let Some(raw) = capture_extlst(xml_reader, &e)? {
                            if in_pivot_source {
                                // PivotSource extensions are not modeled yet.
                            } else if in_ser {
                                ser_raw_ext = Some(raw);
                            } else if in_cat_ax || in_val_ax || in_ser_ax {
                                ax_raw_ext = Some(raw);
                            } else if in_chart_type_element {
                                group_raw_ext = Some(raw);
                            } else if in_chart && !in_plot_area {
                                result.raw_extensions.insert("chart".into(), raw);
                            } else if in_plot_area {
                                result.raw_extensions.insert("plotArea".into(), raw);
                            } else {
                                result.raw_extensions.insert("chartSpace".into(), raw);
                            }
                        }
                    }
                    _ => {
                        if in_chart_title {
                            title_depth += 1;
                        }
                        if in_sp_pr {
                            sp_pr_depth += 1;
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let local = e.name().local_name();
                let tag = local.as_ref();
                match tag {
                    b"barDir" if in_chart_type_element && !in_ser => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                bar_dir = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"grouping" if in_chart_type_element && !in_ser => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                grouping = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"scatterStyle" if in_chart_type_element && !in_ser => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                scatter_style = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"radarStyle" if in_chart_type_element && !in_ser => {
                        radar_style = get_val_attr(&e);
                    }
                    b"firstSliceAng" if in_chart_type_element && !in_ser => {
                        first_slice_angle = get_val_u32(&e);
                    }
                    b"holeSize" if in_chart_type_element && !in_ser => {
                        hole_size = get_val_u32(&e);
                    }
                    b"bubbleScale" if in_chart_type_element && !in_ser => {
                        bubble_scale = get_val_u32(&e);
                    }
                    b"showNegBubbles" if in_chart_type_element && !in_ser => {
                        show_negative_bubbles = get_val_bool(&e);
                    }
                    b"wireframe" if in_chart_type_element && !in_ser => {
                        wireframe = get_val_bool(&e);
                    }
                    b"legendPos" if in_legend => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(val) = attr.unescape_value() {
                                    legend_pos = Some(match val.as_ref() {
                                        "b" => LegendPosition::Bottom,
                                        "t" => LegendPosition::Top,
                                        "l" => LegendPosition::Left,
                                        "r" => LegendPosition::Right,
                                        "tr" => LegendPosition::TopRight,
                                        _ => LegendPosition::Right,
                                    });
                                }
                            }
                        }
                    }
                    b"min" if in_ax_scaling => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                ax_min = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse::<f64>().ok());
                            }
                        }
                    }
                    b"max" if in_ax_scaling => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                ax_max = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse::<f64>().ok());
                            }
                        }
                    }
                    // Data label children
                    b"showLegendKey" if in_dlbls => {
                        dlbls.show_legend_key = get_val_bool(&e);
                    }
                    b"showVal" if in_dlbls => dlbls.show_value = get_val_bool(&e),
                    b"showCatName" if in_dlbls => {
                        dlbls.show_category_name = get_val_bool(&e);
                    }
                    b"showSerName" if in_dlbls => {
                        dlbls.show_series_name = get_val_bool(&e);
                    }
                    b"showPercent" if in_dlbls => dlbls.show_percent = get_val_bool(&e),
                    b"showBubbleSize" if in_dlbls => {
                        dlbls.show_bubble_size = get_val_bool(&e);
                    }
                    b"dLblPos" if in_dlbls => {
                        dlbls.position =
                            get_val_attr(&e).and_then(|s| parse_data_label_position(&s));
                    }
                    b"numFmt" if in_dlbls => {
                        dlbls.number_format = Some(parse_num_fmt(&e));
                    }
                    b"showLeaderLines" if in_dlbls => {
                        dlbls.show_leader_lines = get_val_bool(&e);
                    }
                    // Data point children
                    b"idx" if in_dpt => dpt_index = get_val_u32(&e).unwrap_or(0),
                    b"explosion" if in_dpt => dpt_explosion = get_val_u32(&e),
                    b"explosion" if in_ser && !in_dpt => ser_explosion = get_val_u32(&e),
                    // Marker children
                    b"symbol" if in_marker => {
                        marker_symbol = get_val_attr(&e).and_then(|s| parse_marker_symbol(&s));
                    }
                    b"size" if in_marker => marker_size = get_val_u8(&e),
                    // Trendline children
                    b"trendlineType" if in_trendline => {
                        trendline_type = get_val_attr(&e).and_then(|s| parse_trendline_type(&s));
                    }
                    b"order" if in_trendline => trendline_order = get_val_u32(&e),
                    b"period" if in_trendline => trendline_period = get_val_u32(&e),
                    b"forward" if in_trendline => trendline_forward = get_val_f64(&e),
                    b"backward" if in_trendline => trendline_backward = get_val_f64(&e),
                    b"intercept" if in_trendline => trendline_intercept = get_val_f64(&e),
                    b"dispRSqr" if in_trendline => {
                        trendline_disp_r_sqr = get_val_bool(&e);
                    }
                    b"dispEq" if in_trendline => trendline_disp_eq = get_val_bool(&e),
                    // Error bar children
                    b"errDir" if in_err_bars => {
                        err_dir = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "x" => Some(ErrorBarDirection::X),
                            "y" => Some(ErrorBarDirection::Y),
                            _ => None,
                        });
                    }
                    b"errBarType" if in_err_bars => {
                        err_bar_type = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "both" => Some(ErrorBarType::Both),
                            "minus" => Some(ErrorBarType::Minus),
                            "plus" => Some(ErrorBarType::Plus),
                            _ => None,
                        });
                    }
                    b"errValType" if in_err_bars => {
                        err_val_type = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "cust" => Some(ErrorValueType::Custom),
                            "fixedVal" => Some(ErrorValueType::FixedValue),
                            "percentage" => Some(ErrorValueType::Percentage),
                            "stdDev" => Some(ErrorValueType::StandardDeviation),
                            "stdErr" => Some(ErrorValueType::StandardError),
                            _ => None,
                        });
                    }
                    b"val" if in_err_bars => err_val = get_val_f64(&e),
                    b"noEndCap" if in_err_bars => err_no_end_cap = get_val_bool(&e),
                    // Series smooth
                    b"smooth" if in_ser => ser_smooth = get_val_bool(&e),
                    b"invertIfNegative" if in_ser => {
                        ser_invert_if_negative = get_val_bool(&e);
                    }
                    // Axis enhancements
                    b"numFmt"
                        if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title && !in_dlbls =>
                    {
                        ax_number_format = Some(parse_num_fmt(&e));
                    }
                    b"majorGridlines" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_major_gridlines = true;
                    }
                    b"minorGridlines" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_minor_gridlines = true;
                    }
                    b"majorTickMark" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_major_tick_mark = get_val_attr(&e).and_then(|s| parse_tick_mark(&s));
                    }
                    b"minorTickMark" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_minor_tick_mark = get_val_attr(&e).and_then(|s| parse_tick_mark(&s));
                    }
                    b"tickLblPos" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_label_position =
                            get_val_attr(&e).and_then(|s| parse_tick_label_position(&s));
                    }
                    b"delete" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_delete = get_val_bool(&e);
                    }
                    b"crosses" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_crosses = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "autoZero" => Some(AxisCrosses::AutoZero),
                            "min" => Some(AxisCrosses::Min),
                            "max" => Some(AxisCrosses::Max),
                            _ => None,
                        });
                    }
                    b"crossBetween" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_cross_between = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "between" => Some(CrossBetween::Between),
                            "midCat" => Some(CrossBetween::MidCat),
                            _ => None,
                        });
                    }
                    // View 3D children
                    b"rotX" if in_view_3d => view_3d.rotate_x = get_val_i32(&e),
                    b"rotY" if in_view_3d => view_3d.rotate_y = get_val_i32(&e),
                    b"depthPercent" if in_view_3d => {
                        view_3d.depth_percent = get_val_u32(&e);
                    }
                    b"hPercent" if in_view_3d => view_3d.height_percent = get_val_u32(&e),
                    b"perspective" if in_view_3d => view_3d.perspective = get_val_u32(&e),
                    b"rAngAx" if in_view_3d => {
                        view_3d.right_angle_axes = get_val_bool(&e);
                    }
                    // Chart-level config
                    b"plotVisOnly" if in_chart && !in_plot_area => {
                        result.plot_visible_only = get_val_bool(&e);
                    }
                    b"autoTitleDeleted" if in_chart && !in_plot_area => {
                        result.auto_title_deleted = get_val_bool(&e);
                    }
                    b"showDLblsOverMax" if in_chart && !in_plot_area => {
                        result.show_dlbls_over_max = get_val_bool(&e);
                    }
                    b"roundedCorners" if in_chart_space && !in_chart => {
                        result.rounded_corners = get_val_bool(&e);
                    }
                    b"fmtId" if in_pivot_source => {
                        pivot_source_format_id = get_val_u32(&e);
                    }
                    b"dispBlanksAs" if in_chart && !in_plot_area => {
                        result.display_blanks_as =
                            get_val_attr(&e).and_then(|s| match s.as_str() {
                                "gap" => Some(DisplayBlanksAs::Gap),
                                "span" => Some(DisplayBlanksAs::Span),
                                "zero" => Some(DisplayBlanksAs::Zero),
                                _ => None,
                            });
                    }
                    // Manual layout children
                    b"x" if in_manual_layout => manual_layout.x = get_val_f64(&e),
                    b"y" if in_manual_layout => manual_layout.y = get_val_f64(&e),
                    b"w" if in_manual_layout => manual_layout.width = get_val_f64(&e),
                    b"h" if in_manual_layout => manual_layout.height = get_val_f64(&e),
                    // Data table children
                    b"showHorzBorder" if in_d_table => {
                        d_table.show_horizontal_border = get_val_bool(&e);
                    }
                    b"showVertBorder" if in_d_table => {
                        d_table.show_vertical_border = get_val_bool(&e);
                    }
                    b"showOutline" if in_d_table => {
                        d_table.show_outline = get_val_bool(&e);
                    }
                    b"showKeys" if in_d_table => d_table.show_keys = get_val_bool(&e),
                    b"srgbClr" if in_sp_pr && !in_sp_ln => {
                        if let Some(hex) = get_val_attr(&e) {
                            sp_solid_fill = Some(ChartColor { hex });
                        }
                    }
                    b"srgbClr" if in_sp_ln => {
                        if let Some(hex) = get_val_attr(&e) {
                            sp_ln_solid_fill = Some(ChartColor { hex });
                        }
                    }
                    b"noFill" if in_sp_pr && !in_sp_ln => sp_no_fill = true,
                    b"noFill" if in_sp_ln => sp_ln_no_fill = true,
                    b"prstDash" if in_sp_ln => sp_ln_dash = get_val_attr(&e),
                    b"axPos" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_position = get_val_attr(&e).and_then(|s| match s.as_str() {
                            "b" => Some(AxisPosition::Bottom),
                            "t" => Some(AxisPosition::Top),
                            "l" => Some(AxisPosition::Left),
                            "r" => Some(AxisPosition::Right),
                            _ => None,
                        });
                    }
                    b"majorUnit" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_major_unit = get_val_f64(&e);
                    }
                    b"minorUnit" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_minor_unit = get_val_f64(&e);
                    }
                    b"overlay" if in_legend => legend_overlay = get_val_bool(&e),
                    b"varyColors" if in_chart_type_element && !in_ser => {
                        vary_colors = get_val_bool(&e);
                    }
                    b"gapWidth" if in_up_down_bars => {
                        up_down_bars_gap_width = get_val_u32(&e);
                    }
                    b"gapWidth" if in_chart_type_element && !in_ser => {
                        gap_width = get_val_u32(&e);
                    }
                    b"overlap" if in_chart_type_element && !in_ser => {
                        overlap = get_val_i32(&e).map(|v| v as i32);
                    }
                    b"axId" if in_chart_type_element && !in_ser => {
                        if let Some(id) = get_val_u32(&e) {
                            group_axis_ids.push(id);
                        }
                    }
                    b"axId" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_id = get_val_u32(&e);
                    }
                    b"crossAx" if (in_cat_ax || in_val_ax || in_ser_ax) && !in_ax_title => {
                        ax_cross_id = get_val_u32(&e);
                    }
                    b"dropLines" if in_chart_type_element && !in_ser => {
                        group_drop_lines = Some(ChartLines::default());
                    }
                    b"hiLowLines" if in_chart_type_element && !in_ser => {
                        group_high_low_lines = Some(ChartLines::default());
                    }
                    b"serLines" if in_chart_type_element && !in_ser => {
                        group_series_lines = Some(ChartLines::default());
                    }
                    b"leaderLines" if in_dlbls => {
                        dlbls.leader_lines = Some(ChartLines::default());
                    }
                    b"upBars" if in_up_down_bars => had_up_bars = true,
                    b"downBars" if in_up_down_bars => had_down_bars = true,
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if let Ok(text) = e.unescape() {
                    let text_str = text.as_ref();
                    if in_title_t {
                        title_text.push_str(text_str);
                    } else if in_title_str_ref_f {
                        title_text.push_str(text_str);
                    } else if in_ser_tx_str_ref_f {
                        ser_name = Some(text_str.to_string());
                    } else if in_ser_tx_v {
                        ser_name = Some(text_str.to_string());
                    } else if in_ser_val_num_ref_f {
                        ser_val_formula = Some(text_str.to_string());
                    } else if in_ser_val_pt_v {
                        if let Ok(v) = text_str.parse::<f64>() {
                            ser_val_cache.push(v);
                        }
                    } else if in_ser_cat_ref_f {
                        ser_cat_formula = Some(text_str.to_string());
                    } else if in_ax_title_t {
                        ax_title_text.push_str(text_str);
                    } else if in_trendline_name {
                        trendline_name = Some(text_str.to_string());
                    } else if in_dlbls_separator {
                        dlbls.separator = Some(text_str.to_string());
                    } else if in_pivot_source_name {
                        pivot_source_name.push_str(text_str);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let local = e.name().local_name();
                let tag = local.as_ref();
                match tag {
                    b"chart" => in_chart = false,
                    b"name" if in_pivot_source_name => in_pivot_source_name = false,
                    b"pivotSource" if in_pivot_source => {
                        if !pivot_source_name.is_empty() {
                            result.pivot_source = Some(PivotChartSource {
                                name: pivot_source_name.clone(),
                                format_id: pivot_source_format_id.unwrap_or(0),
                            });
                        }
                        in_pivot_source = false;
                    }
                    b"plotArea" => in_plot_area = false,
                    b"view3D" if in_view_3d => {
                        result.view_3d = Some(view_3d.clone());
                        in_view_3d = false;
                    }
                    b"layout" if in_layout => {
                        if had_manual_layout {
                            result.layout = Some(Layout {
                                manual_layout: Some(manual_layout.clone()),
                            });
                        }
                        had_manual_layout = false;
                        manual_layout = ManualLayout::default();
                        in_layout = false;
                    }
                    b"manualLayout" if in_manual_layout => in_manual_layout = false,
                    b"dTable" if in_d_table => {
                        result.data_table = Some(d_table.clone());
                        in_d_table = false;
                    }
                    b"title" if in_ax_title => in_ax_title = false,
                    b"title" if in_chart_title => {
                        title_depth = title_depth.saturating_sub(1);
                        if title_depth == 0 {
                            if !title_text.is_empty() {
                                result.title = Some(title_text.clone());
                            }
                            in_chart_title = false;
                        }
                    }
                    b"tx" if in_title_tx => in_title_tx = false,
                    b"rich" if in_title_rich => in_title_rich = false,
                    b"p" if in_title_p && in_title_rich => in_title_p = false,
                    b"r" if in_title_r && in_title_p => in_title_r = false,
                    b"t" if in_title_t => in_title_t = false,
                    b"strRef" if in_title_str_ref => in_title_str_ref = false,
                    b"f" if in_title_str_ref_f => in_title_str_ref_f = false,
                    b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
                    | b"pie3DChart" | b"doughnutChart" | b"areaChart" | b"area3DChart"
                    | b"scatterChart" | b"bubbleChart" | b"radarChart" | b"stockChart"
                    | b"surfaceChart" | b"surface3DChart" | b"ofPieChart"
                        if in_chart_type_element =>
                    {
                        let ct = resolve_chart_type(
                            chart_type_tag.as_deref(),
                            bar_dir.as_deref(),
                            grouping.as_deref(),
                            scatter_style.as_deref(),
                        );
                        let is_3d = chart_type_tag
                            .as_deref()
                            .map_or(false, |t| t.contains("3D"));
                        let group = ChartTypeGroup {
                            chart_type: ct,
                            is_3d,
                            series: std::mem::take(&mut group_series),
                            data_labels: group_data_labels.take(),
                            vary_colors: vary_colors.take(),
                            gap_width: gap_width.take(),
                            overlap: overlap.take(),
                            first_slice_angle: first_slice_angle.take(),
                            hole_size: hole_size.take(),
                            bubble_scale: bubble_scale.take(),
                            show_negative_bubbles: show_negative_bubbles.take(),
                            radar_style: radar_style.take(),
                            wireframe: wireframe.take(),
                            drop_lines: group_drop_lines.take(),
                            high_low_lines: group_high_low_lines.take(),
                            series_lines: group_series_lines.take(),
                            up_down_bars: group_up_down_bars.take(),
                            axis_ids: std::mem::take(&mut group_axis_ids),
                            raw_ext: group_raw_ext.take(),
                            of_pie_type: None,
                            split_type: None,
                            split_pos: None,
                            second_pie_size: None,
                            bar_shape: None,
                            floor: None,
                            side_wall: None,
                            back_wall: None,
                        };
                        result.type_groups.push(group);
                        in_chart_type_element = false;
                    }
                    b"dropLines" if in_drop_lines => {
                        if group_drop_lines.is_none() {
                            group_drop_lines = Some(ChartLines::default());
                        }
                        in_drop_lines = false;
                    }
                    b"hiLowLines" if in_hi_low_lines => {
                        if group_high_low_lines.is_none() {
                            group_high_low_lines = Some(ChartLines::default());
                        }
                        in_hi_low_lines = false;
                    }
                    b"serLines" if in_ser_lines => {
                        if group_series_lines.is_none() {
                            group_series_lines = Some(ChartLines::default());
                        }
                        in_ser_lines = false;
                    }
                    b"upBars" if in_up_bars => in_up_bars = false,
                    b"downBars" if in_down_bars => in_down_bars = false,
                    b"upDownBars" if in_up_down_bars => {
                        let up = up_bars_sp
                            .take()
                            .map(|sp| ChartLines {
                                shape_properties: Some(sp),
                            })
                            .or(if had_up_bars {
                                Some(ChartLines::default())
                            } else {
                                None
                            });
                        let down = down_bars_sp
                            .take()
                            .map(|sp| ChartLines {
                                shape_properties: Some(sp),
                            })
                            .or(if had_down_bars {
                                Some(ChartLines::default())
                            } else {
                                None
                            });
                        group_up_down_bars = Some(UpDownBars {
                            gap_width: up_down_bars_gap_width.take(),
                            up_bars: up,
                            down_bars: down,
                        });
                        had_up_bars = false;
                        had_down_bars = false;
                        in_up_down_bars = false;
                    }
                    b"leaderLines" if in_leader_lines => {
                        if dlbls.leader_lines.is_none() {
                            dlbls.leader_lines = Some(ChartLines::default());
                        }
                        in_leader_lines = false;
                    }
                    // Data labels
                    b"separator" if in_dlbls_separator => in_dlbls_separator = false,
                    b"dLbls" if in_dlbls => {
                        if in_ser {
                            ser_data_labels = Some(dlbls.clone());
                        } else {
                            group_data_labels = Some(dlbls.clone());
                        }
                        in_dlbls = false;
                    }
                    // Marker
                    b"marker" if in_marker => {
                        let m = Marker {
                            symbol: marker_symbol.take(),
                            size: marker_size.take(),
                        };
                        if in_dpt {
                            dpt_marker = Some(m);
                        } else {
                            ser_marker = Some(m);
                        }
                        in_marker = false;
                    }
                    // Data point
                    b"dPt" if in_dpt => {
                        ser_data_points.push(DataPoint {
                            index: dpt_index,
                            marker: dpt_marker.take(),
                            explosion: dpt_explosion.take(),
                            shape_properties: dpt_shape_properties.take(),
                        });
                        in_dpt = false;
                    }
                    // Trendline
                    b"name" if in_trendline_name => in_trendline_name = false,
                    b"trendline" if in_trendline => {
                        if let Some(tt) = trendline_type.take() {
                            ser_trendline = Some(Trendline {
                                trendline_type: tt,
                                name: trendline_name.take(),
                                order: trendline_order.take(),
                                period: trendline_period.take(),
                                forward: trendline_forward.take(),
                                backward: trendline_backward.take(),
                                intercept: trendline_intercept.take(),
                                label: None,
                                display_r_squared: trendline_disp_r_sqr.take(),
                                display_equation: trendline_disp_eq.take(),
                            });
                        }
                        in_trendline = false;
                    }
                    // Error bars
                    b"errBars" if in_err_bars => {
                        ser_error_bars = Some(ErrorBars {
                            direction: err_dir.unwrap_or(ErrorBarDirection::Y),
                            bar_type: err_bar_type.unwrap_or(ErrorBarType::Both),
                            value_type: err_val_type.unwrap_or(ErrorValueType::FixedValue),
                            value: err_val.take(),
                            no_end_cap: err_no_end_cap.take(),
                            plus: None,
                            minus: None,
                        });
                        in_err_bars = false;
                    }
                    b"ser" if in_ser => {
                        let values = if let Some(ref f) = ser_val_formula {
                            DataReference::formula(f)
                        } else if !ser_val_cache.is_empty() {
                            DataReference::numbers(ser_val_cache.clone())
                        } else {
                            DataReference::numbers(Vec::new())
                        };

                        let mut ds = DataSeries::new(values);
                        if let Some(ref name) = ser_name {
                            ds = ds.with_name(name);
                        }
                        if let Some(ref f) = ser_cat_formula {
                            ds = ds.with_categories(DataReference::formula(f));
                        }
                        ds.data_labels = ser_data_labels.take();
                        ds.trendline = ser_trendline.take();
                        ds.error_bars = ser_error_bars.take();
                        ds.marker = ser_marker.take();
                        ds.data_points = std::mem::take(&mut ser_data_points);
                        ds.smooth = ser_smooth.take();
                        ds.explosion = ser_explosion.take();
                        ds.shape_properties = ser_shape_properties.take();
                        ds.raw_ext = ser_raw_ext.take();
                        ds.invert_if_negative = ser_invert_if_negative.take();
                        group_series.push(ds);

                        in_ser = false;
                        ser_name = None;
                        ser_val_formula = None;
                        ser_val_cache.clear();
                        ser_cat_formula = None;
                    }
                    b"tx" if in_ser_tx => in_ser_tx = false,
                    b"strRef" if in_ser_tx_str_ref => in_ser_tx_str_ref = false,
                    b"f" if in_ser_tx_str_ref_f => in_ser_tx_str_ref_f = false,
                    b"v" if in_ser_tx_v => in_ser_tx_v = false,
                    b"val" if in_ser_val => {
                        in_ser_val = false;
                        in_ser_val_num_ref = false;
                    }
                    b"yVal" if in_ser_yval => {
                        in_ser_yval = false;
                        in_ser_val_num_ref = false;
                    }
                    b"numRef" if in_ser_val_num_ref => in_ser_val_num_ref = false,
                    b"f" if in_ser_val_num_ref_f => in_ser_val_num_ref_f = false,
                    b"numCache" if in_ser_val_num_cache => in_ser_val_num_cache = false,
                    b"pt" if in_ser_val_pt => in_ser_val_pt = false,
                    b"v" if in_ser_val_pt_v => in_ser_val_pt_v = false,
                    b"cat" if in_ser_cat => {
                        in_ser_cat = false;
                        in_ser_cat_ref = false;
                    }
                    b"xVal" if in_ser_xval => {
                        in_ser_xval = false;
                        in_ser_cat_ref = false;
                    }
                    b"strRef" | b"numRef" if in_ser_cat_ref => in_ser_cat_ref = false,
                    b"f" if in_ser_cat_ref_f => in_ser_cat_ref_f = false,
                    b"catAx" | b"dateAx" if in_cat_ax => {
                        let mut axis = Axis::new();
                        if !ax_title_text.is_empty() {
                            axis = axis.with_title(&ax_title_text);
                        }
                        if let (Some(min), Some(max)) = (ax_min, ax_max) {
                            axis = axis.with_bounds(min, max);
                        } else {
                            axis.minimum = ax_min;
                            axis.maximum = ax_max;
                        }
                        axis.number_format = ax_number_format.take();
                        axis.major_gridlines = ax_major_gridlines;
                        axis.minor_gridlines = ax_minor_gridlines;
                        axis.major_gridlines_shape_properties =
                            ax_major_gridlines_shape_properties.take();
                        axis.minor_gridlines_shape_properties =
                            ax_minor_gridlines_shape_properties.take();
                        axis.major_tick_mark = ax_major_tick_mark.take();
                        axis.minor_tick_mark = ax_minor_tick_mark.take();
                        axis.label_position = ax_label_position.take();
                        axis.delete = ax_delete.take();
                        axis.crosses = ax_crosses.take();
                        axis.cross_between = ax_cross_between.take();
                        if let Some(pos) = ax_position.take() {
                            axis.position = pos;
                        }
                        axis.major_unit = ax_major_unit.take();
                        axis.minor_unit = ax_minor_unit.take();
                        axis.shape_properties = ax_shape_properties.take();
                        axis.raw_ext = ax_raw_ext.take();
                        if is_date_ax {
                            axis.axis_type = AxisType::Date;
                        }
                        result.category_axis = Some(axis.clone());
                        if let Some(id) = ax_id.take() {
                            result.axes.push(ChartAxis {
                                id,
                                cross_id: ax_cross_id.take().unwrap_or(0),
                                axis: axis,
                            });
                        }
                        in_cat_ax = false;
                        ax_title_text.clear();
                    }
                    b"valAx" if in_val_ax => {
                        let mut axis = Axis::new();
                        if !ax_title_text.is_empty() {
                            axis = axis.with_title(&ax_title_text);
                        }
                        if let (Some(min), Some(max)) = (ax_min, ax_max) {
                            axis = axis.with_bounds(min, max);
                        } else {
                            axis.minimum = ax_min;
                            axis.maximum = ax_max;
                        }
                        axis.number_format = ax_number_format.take();
                        axis.major_gridlines = ax_major_gridlines;
                        axis.minor_gridlines = ax_minor_gridlines;
                        axis.major_gridlines_shape_properties =
                            ax_major_gridlines_shape_properties.take();
                        axis.minor_gridlines_shape_properties =
                            ax_minor_gridlines_shape_properties.take();
                        axis.major_tick_mark = ax_major_tick_mark.take();
                        axis.minor_tick_mark = ax_minor_tick_mark.take();
                        axis.label_position = ax_label_position.take();
                        axis.delete = ax_delete.take();
                        axis.crosses = ax_crosses.take();
                        axis.cross_between = ax_cross_between.take();
                        if let Some(pos) = ax_position.take() {
                            axis.position = pos;
                        }
                        axis.major_unit = ax_major_unit.take();
                        axis.minor_unit = ax_minor_unit.take();
                        axis.shape_properties = ax_shape_properties.take();
                        axis.raw_ext = ax_raw_ext.take();
                        axis.axis_type = AxisType::Value;
                        result.value_axis = Some(axis.clone());
                        if let Some(id) = ax_id.take() {
                            result.axes.push(ChartAxis {
                                id,
                                cross_id: ax_cross_id.take().unwrap_or(0),
                                axis: axis,
                            });
                        }
                        in_val_ax = false;
                        ax_title_text.clear();
                    }
                    b"serAx" if in_ser_ax => {
                        let mut axis = Axis::new();
                        axis.axis_type = AxisType::Series;
                        if !ax_title_text.is_empty() {
                            axis = axis.with_title(&ax_title_text);
                        }
                        if let (Some(min), Some(max)) = (ax_min, ax_max) {
                            axis = axis.with_bounds(min, max);
                        } else {
                            axis.minimum = ax_min;
                            axis.maximum = ax_max;
                        }
                        axis.number_format = ax_number_format.take();
                        axis.major_gridlines = ax_major_gridlines;
                        axis.minor_gridlines = ax_minor_gridlines;
                        axis.major_gridlines_shape_properties =
                            ax_major_gridlines_shape_properties.take();
                        axis.minor_gridlines_shape_properties =
                            ax_minor_gridlines_shape_properties.take();
                        axis.major_tick_mark = ax_major_tick_mark.take();
                        axis.minor_tick_mark = ax_minor_tick_mark.take();
                        axis.label_position = ax_label_position.take();
                        axis.delete = ax_delete.take();
                        axis.crosses = ax_crosses.take();
                        axis.cross_between = ax_cross_between.take();
                        if let Some(pos) = ax_position.take() {
                            axis.position = pos;
                        }
                        axis.major_unit = ax_major_unit.take();
                        axis.minor_unit = ax_minor_unit.take();
                        axis.shape_properties = ax_shape_properties.take();
                        axis.raw_ext = ax_raw_ext.take();
                        result.series_axis = Some(axis.clone());
                        if let Some(id) = ax_id.take() {
                            result.axes.push(ChartAxis {
                                id,
                                cross_id: ax_cross_id.take().unwrap_or(0),
                                axis: axis,
                            });
                        }
                        in_ser_ax = false;
                        ax_title_text.clear();
                    }
                    b"tx" if in_ax_title_tx => in_ax_title_tx = false,
                    b"rich" if in_ax_title_rich => in_ax_title_rich = false,
                    b"p" if in_ax_title_p && in_ax_title_rich => in_ax_title_p = false,
                    b"r" if in_ax_title_r => in_ax_title_r = false,
                    b"t" if in_ax_title_t => in_ax_title_t = false,
                    b"scaling" if in_ax_scaling => in_ax_scaling = false,
                    b"majorGridlines" if in_ax_major_gridlines => in_ax_major_gridlines = false,
                    b"minorGridlines" if in_ax_minor_gridlines => in_ax_minor_gridlines = false,
                    b"legend" if in_legend => {
                        let mut leg = Legend::new(legend_pos.unwrap_or(LegendPosition::Right));
                        if let Some(true) = legend_overlay {
                            leg.overlay = true;
                        }
                        leg.shape_properties = legend_shape_properties.take();
                        result.legend = Some(leg);
                        in_legend = false;
                    }
                    b"ln" if in_sp_ln => {
                        sp_line = Some(ChartLine {
                            width: sp_ln_width.take(),
                            solid_fill: sp_ln_solid_fill.take(),
                            no_fill: sp_ln_no_fill,
                            dash_style: sp_ln_dash.take(),
                        });
                        in_sp_ln = false;
                        sp_ln_no_fill = false;
                        sp_pr_depth = sp_pr_depth.saturating_sub(1);
                    }
                    b"spPr" if in_sp_pr => {
                        sp_pr_depth = sp_pr_depth.saturating_sub(1);
                        if sp_pr_depth == 0 {
                            let props = ChartShapeProperties {
                                solid_fill: sp_solid_fill.take(),
                                no_fill: sp_no_fill,
                                line: sp_line.take(),
                            };
                            let has_content =
                                props.solid_fill.is_some() || props.no_fill || props.line.is_some();
                            if has_content {
                                match sp_pr_context {
                                    SpPrContext::Series => {
                                        ser_shape_properties = Some(props);
                                    }
                                    SpPrContext::DataPoint => {
                                        dpt_shape_properties = Some(props);
                                    }
                                    SpPrContext::CatAxis => {
                                        ax_shape_properties = Some(props);
                                    }
                                    SpPrContext::ValAxis => {
                                        ax_shape_properties = Some(props);
                                    }
                                    SpPrContext::MajorGridlines => {
                                        ax_major_gridlines = true;
                                        ax_major_gridlines_shape_properties = Some(props);
                                    }
                                    SpPrContext::MinorGridlines => {
                                        ax_minor_gridlines = true;
                                        ax_minor_gridlines_shape_properties = Some(props);
                                    }
                                    SpPrContext::ChartSpace => {
                                        result.shape_properties = Some(props);
                                    }
                                    SpPrContext::Legend => {
                                        legend_shape_properties = Some(props);
                                    }
                                    SpPrContext::None => {}
                                    SpPrContext::DropLines => {
                                        group_drop_lines = Some(ChartLines {
                                            shape_properties: Some(props),
                                        });
                                    }
                                    SpPrContext::HiLowLines => {
                                        group_high_low_lines = Some(ChartLines {
                                            shape_properties: Some(props),
                                        });
                                    }
                                    SpPrContext::SerLines => {
                                        group_series_lines = Some(ChartLines {
                                            shape_properties: Some(props),
                                        });
                                    }
                                    SpPrContext::UpBars => up_bars_sp = Some(props),
                                    SpPrContext::DownBars => down_bars_sp = Some(props),
                                    SpPrContext::LeaderLines => {
                                        dlbls.leader_lines = Some(ChartLines {
                                            shape_properties: Some(props),
                                        });
                                    }
                                }
                            }
                            in_sp_pr = false;
                            sp_no_fill = false;
                            sp_pr_context = SpPrContext::None;
                        }
                    }
                    _ => {
                        if in_chart_title {
                            title_depth = title_depth.saturating_sub(1);
                            if title_depth == 0 {
                                if !title_text.is_empty() {
                                    result.title = Some(title_text.clone());
                                }
                                in_chart_title = false;
                            }
                        }
                        if in_sp_pr {
                            sp_pr_depth = sp_pr_depth.saturating_sub(1);
                            if sp_pr_depth == 0 {
                                in_sp_pr = false;
                                sp_no_fill = false;
                                sp_pr_context = SpPrContext::None;
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ChartParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    // Populate legacy fields from type_groups for backward compatibility
    if let Some(first) = result.type_groups.first() {
        result.chart_type = first.chart_type.clone();
        result.is_3d = first.is_3d;
        result.series = first.series.clone();
        result.data_labels = first.data_labels.clone();
        result.vary_colors = first.vary_colors;
        result.gap_width = first.gap_width;
        result.overlap = first.overlap;
        result.first_slice_angle = first.first_slice_angle;
        result.hole_size = first.hole_size;
        result.bubble_scale = first.bubble_scale;
        result.show_negative_bubbles = first.show_negative_bubbles;
        result.radar_style = first.radar_style.clone();
        result.wireframe = first.wireframe;
        result.drop_lines = first.drop_lines.clone();
        result.high_low_lines = first.high_low_lines.clone();
        result.up_down_bars = first.up_down_bars.clone();
        result.series_lines = first.series_lines.clone();
    }

    // Detect PieExploded: shares the same XML element (pieChart) as Pie
    // but has explosion attributes on series.
    if result.chart_type == ChartType::Pie {
        let has_explosion = result.series.iter().any(|s| s.explosion.is_some());
        if has_explosion {
            result.chart_type = ChartType::PieExploded;
        }
    }

    // Re-populate legacy axis fields from result.axes using first group's axis_ids.
    // Only needed for combo charts (2+ groups) where multiple value axes exist
    // and the parse-loop's last-wins behavior gives the wrong legacy value_axis.
    if result.type_groups.len() >= 2 {
        if let Some(first) = result.type_groups.first() {
            let axis_ids = &first.axis_ids;
            result.category_axis = None;
            result.value_axis = None;
            for ax in &result.axes {
                if axis_ids.contains(&ax.id) {
                    match ax.axis.axis_type {
                        AxisType::Category | AxisType::Date => {
                            if result.category_axis.is_none() {
                                result.category_axis = Some(ax.axis.clone());
                            }
                        }
                        AxisType::Value => {
                            if result.value_axis.is_none() {
                                result.value_axis = Some(ax.axis.clone());
                            }
                        }
                        AxisType::Series => {
                            if result.series_axis.is_none() {
                                result.series_axis = Some(ax.axis.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    // Only keep type_groups/axes in combo mode (2+ groups)
    if result.type_groups.len() < 2 {
        if let Some(first) = result.type_groups.first_mut() {
            if let Some(raw) = first.raw_ext.take() {
                result.raw_extensions.insert("chartType".into(), raw);
            }
        }
        result.type_groups.clear();
        result.axes.clear();
    }

    Ok(result)
}

/// Capture an entire `<c:extLst>...</c:extLst>` element as raw XML bytes.
/// The Start event for `extLst` has already been read; we write it plus all
/// inner events until the matching End into a buffer.
fn capture_extlst<R: Read>(
    xml_reader: &mut Reader<BufReader<R>>,
    start_event: &quick_xml::events::BytesStart,
) -> ChartParseResult<Option<Vec<u8>>> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Start(start_event.to_owned()))?;

    let mut depth: u32 = 1;
    let mut read_buf = Vec::new();
    loop {
        read_buf.clear();
        match xml_reader.read_event_into(&mut read_buf) {
            Ok(ref ev @ Event::Start(_)) => {
                depth += 1;
                writer.write_event(ev.clone())?;
            }
            Ok(ref ev @ Event::End(_)) => {
                depth -= 1;
                writer.write_event(ev.clone())?;
                if depth == 0 {
                    return Ok(Some(writer.into_inner().into_inner()));
                }
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(ref ev) => writer.write_event(ev.clone())?,
            Err(e) => return Err(ChartParseError::Xml(e)),
        }
    }
}

fn resolve_chart_type(
    tag: Option<&str>,
    bar_dir: Option<&str>,
    grouping: Option<&str>,
    scatter_style: Option<&str>,
) -> ChartType {
    match tag {
        Some("barChart") | Some("bar3DChart") => {
            let is_bar = bar_dir == Some("bar");
            match grouping.unwrap_or("clustered") {
                "stacked" => {
                    if is_bar {
                        ChartType::BarStacked
                    } else {
                        ChartType::ColumnStacked
                    }
                }
                "percentStacked" => {
                    if is_bar {
                        ChartType::BarPercentStacked
                    } else {
                        ChartType::ColumnPercentStacked
                    }
                }
                _ => {
                    if is_bar {
                        ChartType::BarClustered
                    } else {
                        ChartType::ColumnClustered
                    }
                }
            }
        }
        Some("lineChart") | Some("line3DChart") => match grouping.unwrap_or("standard") {
            "stacked" | "percentStacked" => ChartType::LineStacked,
            _ => ChartType::Line,
        },
        Some("pieChart") | Some("pie3DChart") => ChartType::Pie,
        Some("doughnutChart") => ChartType::Doughnut,
        Some("areaChart") | Some("area3DChart") => match grouping.unwrap_or("standard") {
            "stacked" => ChartType::AreaStacked,
            "percentStacked" => ChartType::AreaPercentStacked,
            _ => ChartType::Area,
        },
        Some("scatterChart") => match scatter_style.unwrap_or("lineMarker") {
            "smoothMarker" | "smooth" => ChartType::ScatterSmooth,
            "line" | "lineMarker" => ChartType::ScatterLines,
            _ => ChartType::ScatterMarkers,
        },
        Some("bubbleChart") => ChartType::Bubble,
        Some("radarChart") => ChartType::Radar,
        Some("stockChart") => ChartType::Stock,
        Some("surfaceChart") | Some("surface3DChart") => ChartType::Surface,
        Some("ofPieChart") => ChartType::Unsupported("c:ofPieChart".into()),
        Some(other) => ChartType::Unsupported(format!("c:{other}")),
        None => ChartType::Unsupported("unknown".into()),
    }
}

fn get_val_attr(e: &quick_xml::events::BytesStart) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
            return attr.unescape_value().ok().map(|s| s.to_string());
        }
    }
    None
}

fn get_val_bool(e: &quick_xml::events::BytesStart) -> Option<bool> {
    get_val_attr(e).map(|s| s == "1" || s == "true")
}

fn get_val_f64(e: &quick_xml::events::BytesStart) -> Option<f64> {
    get_val_attr(e).and_then(|s| s.parse::<f64>().ok())
}

fn get_val_u32(e: &quick_xml::events::BytesStart) -> Option<u32> {
    get_val_attr(e).and_then(|s| s.parse::<u32>().ok())
}

fn get_val_i32(e: &quick_xml::events::BytesStart) -> Option<i32> {
    get_val_attr(e).and_then(|s| s.parse::<i32>().ok())
}

fn get_val_u8(e: &quick_xml::events::BytesStart) -> Option<u8> {
    get_val_attr(e).and_then(|s| s.parse::<u8>().ok())
}

fn parse_num_fmt(e: &quick_xml::events::BytesStart) -> NumberFormat {
    let mut nf = NumberFormat::default();
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"formatCode" => {
                nf.format_code = attr.unescape_value().unwrap_or_default().to_string();
            }
            b"sourceLinked" => {
                nf.source_linked = attr
                    .unescape_value()
                    .ok()
                    .map(|s| s.as_ref() == "1" || s.as_ref() == "true");
            }
            _ => {}
        }
    }
    nf
}

fn parse_data_label_position(s: &str) -> Option<DataLabelPosition> {
    match s {
        "bestFit" => Some(DataLabelPosition::BestFit),
        "b" => Some(DataLabelPosition::Bottom),
        "ctr" => Some(DataLabelPosition::Center),
        "inBase" => Some(DataLabelPosition::InsideBase),
        "inEnd" => Some(DataLabelPosition::InsideEnd),
        "l" => Some(DataLabelPosition::Left),
        "outEnd" => Some(DataLabelPosition::OutsideEnd),
        "r" => Some(DataLabelPosition::Right),
        "t" => Some(DataLabelPosition::Top),
        _ => None,
    }
}

fn parse_marker_symbol(s: &str) -> Option<MarkerSymbol> {
    match s {
        "circle" => Some(MarkerSymbol::Circle),
        "dash" => Some(MarkerSymbol::Dash),
        "diamond" => Some(MarkerSymbol::Diamond),
        "dot" => Some(MarkerSymbol::Dot),
        "none" => Some(MarkerSymbol::None),
        "picture" => Some(MarkerSymbol::Picture),
        "plus" => Some(MarkerSymbol::Plus),
        "square" => Some(MarkerSymbol::Square),
        "star" => Some(MarkerSymbol::Star),
        "triangle" => Some(MarkerSymbol::Triangle),
        "x" => Some(MarkerSymbol::X),
        "auto" => Some(MarkerSymbol::Auto),
        _ => None,
    }
}

fn parse_trendline_type(s: &str) -> Option<TrendlineType> {
    match s {
        "linear" => Some(TrendlineType::Linear),
        "exp" => Some(TrendlineType::Exponential),
        "log" => Some(TrendlineType::Logarithmic),
        "movingAvg" => Some(TrendlineType::MovingAverage),
        "poly" => Some(TrendlineType::Polynomial),
        "power" => Some(TrendlineType::Power),
        _ => None,
    }
}

fn parse_tick_mark(s: &str) -> Option<TickMark> {
    match s {
        "cross" => Some(TickMark::Cross),
        "in" => Some(TickMark::Inside),
        "none" => Some(TickMark::None),
        "out" => Some(TickMark::Outside),
        _ => None,
    }
}

fn parse_tick_label_position(s: &str) -> Option<TickLabelPosition> {
    match s {
        "high" => Some(TickLabelPosition::High),
        "low" => Some(TickLabelPosition::Low),
        "nextTo" => Some(TickLabelPosition::NextTo),
        "none" => Some(TickLabelPosition::None),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    fn parse_chart_xml_str(xml: &str) -> Chart {
        parse_chart_xml(Cursor::new(xml.as_bytes())).unwrap()
    }

    #[test]
    fn test_parse_pivot_source() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:pivotSource>
    <c:name>SalesPivot</c:name>
    <c:fmtId val="4"/>
  </c:pivotSource>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml_str(xml);
        assert_eq!(
            chart.pivot_source,
            Some(PivotChartSource {
                name: "SalesPivot".into(),
                format_id: 4,
            })
        );
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

        let chart = parse_chart_xml_str(xml);
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
    fn test_parse_pie_chart_with_cache_data() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Slice Data</c:v></c:tx>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>30</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml_str(xml);
        assert_eq!(chart.chart_type, ChartType::Pie);
        assert!(chart.title.is_none());
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].name.as_deref(), Some("Slice Data"));
        match &chart.series[0].values {
            DataReference::Numbers(nums) => assert_eq!(nums, &[10.0, 20.0, 30.0]),
            other => panic!("expected Numbers, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_scatter_chart() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="smoothMarker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:xVal><c:numRef><c:f>Sheet1!$A$1:$A$5</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$5</c:f></c:numRef></c:yVal>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml_str(xml);
        assert_eq!(chart.chart_type, ChartType::ScatterSmooth);
        assert_eq!(chart.series.len(), 1);
        match chart.series[0].categories.as_ref().unwrap() {
            DataReference::Formula(f) => assert_eq!(f, "Sheet1!$A$1:$A$5"),
            other => panic!("expected Formula, got {:?}", other),
        }
        match &chart.series[0].values {
            DataReference::Formula(f) => assert_eq!(f, "Sheet1!$B$1:$B$5"),
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_chart_type_mapping() {
        assert_eq!(
            resolve_chart_type(Some("barChart"), Some("bar"), Some("stacked"), None),
            ChartType::BarStacked
        );
        assert_eq!(
            resolve_chart_type(Some("barChart"), Some("col"), Some("percentStacked"), None),
            ChartType::ColumnPercentStacked
        );
        assert_eq!(
            resolve_chart_type(Some("lineChart"), None, Some("stacked"), None),
            ChartType::LineStacked
        );
        assert_eq!(
            resolve_chart_type(Some("doughnutChart"), None, None, None),
            ChartType::Doughnut
        );
        assert_eq!(
            resolve_chart_type(Some("areaChart"), None, Some("percentStacked"), None),
            ChartType::AreaPercentStacked
        );
        assert_eq!(
            resolve_chart_type(Some("surfaceChart"), None, None, None),
            ChartType::Surface
        );
        assert_eq!(
            resolve_chart_type(Some("ofPieChart"), None, None, None),
            ChartType::Unsupported("c:ofPieChart".into())
        );
    }

    #[test]
    fn test_parse_data_labels_chart_level() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:dLbls>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="1"/>
          <c:showPercent val="0"/>
          <c:showLegendKey val="1"/>
          <c:showBubbleSize val="0"/>
          <c:dLblPos val="outEnd"/>
          <c:separator>,</c:separator>
          <c:numFmt formatCode="0.00" sourceLinked="0"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let dl = parsed.data_labels.unwrap();
        assert_eq!(dl.show_value, Some(true));
        assert_eq!(dl.show_category_name, Some(false));
        assert_eq!(dl.show_series_name, Some(true));
        assert_eq!(dl.show_percent, Some(false));
        assert_eq!(dl.show_legend_key, Some(true));
        assert_eq!(dl.show_bubble_size, Some(false));
        assert_eq!(dl.position, Some(DataLabelPosition::OutsideEnd));
        assert_eq!(dl.separator.as_deref(), Some(","));
        let nf = dl.number_format.unwrap();
        assert_eq!(nf.format_code, "0.00");
        assert_eq!(nf.source_linked, Some(false));
    }

    #[test]
    fn test_parse_data_labels_series_level() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:dLbls>
            <c:showPercent val="1"/>
            <c:dLblPos val="ctr"/>
          </c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert!(parsed.data_labels.is_none());
        assert_eq!(parsed.series.len(), 1);
        let sdl = parsed.series[0].data_labels.as_ref().unwrap();
        assert_eq!(sdl.show_percent, Some(true));
        assert_eq!(sdl.position, Some(DataLabelPosition::Center));
    }

    #[test]
    fn test_parse_data_points() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:dPt>
            <c:idx val="0"/>
            <c:explosion val="25"/>
          </c:dPt>
          <c:dPt>
            <c:idx val="2"/>
            <c:marker>
              <c:symbol val="diamond"/>
              <c:size val="8"/>
            </c:marker>
          </c:dPt>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let pts = &parsed.series[0].data_points;
        assert_eq!(pts.len(), 2);
        assert_eq!(pts[0].index, 0);
        assert_eq!(pts[0].explosion, Some(25));
        assert!(pts[0].marker.is_none());
        assert_eq!(pts[1].index, 2);
        assert!(pts[1].explosion.is_none());
        let m = pts[1].marker.as_ref().unwrap();
        assert_eq!(m.symbol, Some(MarkerSymbol::Diamond));
        assert_eq!(m.size, Some(8));
    }

    #[test]
    fn test_parse_trendline() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$5</c:f></c:numRef></c:yVal>
          <c:trendline>
            <c:trendlineType val="poly"/>
            <c:name>My Trend</c:name>
            <c:order val="3"/>
            <c:forward val="2.5"/>
            <c:backward val="1.0"/>
            <c:intercept val="0.5"/>
            <c:dispRSqr val="1"/>
            <c:dispEq val="0"/>
          </c:trendline>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let t = parsed.series[0].trendline.as_ref().unwrap();
        assert_eq!(t.trendline_type, TrendlineType::Polynomial);
        assert_eq!(t.name.as_deref(), Some("My Trend"));
        assert_eq!(t.order, Some(3));
        assert_eq!(t.forward, Some(2.5));
        assert_eq!(t.backward, Some(1.0));
        assert_eq!(t.intercept, Some(0.5));
        assert_eq!(t.display_r_squared, Some(true));
        assert_eq!(t.display_equation, Some(false));
    }

    #[test]
    fn test_parse_error_bars() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="percentage"/>
            <c:val val="10.0"/>
            <c:noEndCap val="1"/>
          </c:errBars>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let eb = parsed.series[0].error_bars.as_ref().unwrap();
        assert_eq!(eb.direction, ErrorBarDirection::Y);
        assert_eq!(eb.bar_type, ErrorBarType::Both);
        assert_eq!(eb.value_type, ErrorValueType::Percentage);
        assert_eq!(eb.value, Some(10.0));
        assert_eq!(eb.no_end_cap, Some(true));
    }

    #[test]
    fn test_parse_marker_on_series() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$5</c:f></c:numRef></c:val>
          <c:marker>
            <c:symbol val="triangle"/>
            <c:size val="10"/>
          </c:marker>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let m = parsed.series[0].marker.as_ref().unwrap();
        assert_eq!(m.symbol, Some(MarkerSymbol::Triangle));
        assert_eq!(m.size, Some(10));
    }

    #[test]
    fn test_parse_series_smooth_and_explosion() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$5</c:f></c:numRef></c:val>
          <c:smooth val="1"/>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.series[0].smooth, Some(true));
    }

    #[test]
    fn test_parse_series_explosion() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:explosion val="15"/>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.series[0].explosion, Some(15));
    }

    #[test]
    fn test_parse_axis_enhancements() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:delete val="0"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorGridlines/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:delete val="1"/>
        <c:majorTickMark val="cross"/>
        <c:minorTickMark val="in"/>
        <c:tickLblPos val="low"/>
        <c:crosses val="max"/>
        <c:crossBetween val="midCat"/>
        <c:numFmt formatCode="0.00" sourceLinked="0"/>
        <c:majorGridlines/>
        <c:minorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);

        let cat = parsed.category_axis.unwrap();
        assert_eq!(cat.delete, Some(false));
        assert_eq!(cat.major_tick_mark, Some(TickMark::Outside));
        assert_eq!(cat.minor_tick_mark, Some(TickMark::None));
        assert_eq!(cat.label_position, Some(TickLabelPosition::NextTo));
        assert_eq!(cat.crosses, Some(AxisCrosses::AutoZero));
        assert_eq!(cat.cross_between, Some(CrossBetween::Between));
        let cnf = cat.number_format.unwrap();
        assert_eq!(cnf.format_code, "General");
        assert_eq!(cnf.source_linked, Some(true));
        assert!(cat.major_gridlines);
        assert!(!cat.minor_gridlines);

        let val = parsed.value_axis.unwrap();
        assert_eq!(val.delete, Some(true));
        assert_eq!(val.major_tick_mark, Some(TickMark::Cross));
        assert_eq!(val.minor_tick_mark, Some(TickMark::Inside));
        assert_eq!(val.label_position, Some(TickLabelPosition::Low));
        assert_eq!(val.crosses, Some(AxisCrosses::Max));
        assert_eq!(val.cross_between, Some(CrossBetween::MidCat));
        let vnf = val.number_format.unwrap();
        assert_eq!(vnf.format_code, "0.00");
        assert_eq!(vnf.source_linked, Some(false));
        assert!(val.major_gridlines);
        assert!(val.minor_gridlines);
    }

    #[test]
    fn test_parse_view_3d() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:view3D>
      <c:rotX val="15"/>
      <c:rotY val="20"/>
      <c:depthPercent val="100"/>
      <c:hPercent val="150"/>
      <c:perspective val="30"/>
      <c:rAngAx val="1"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let v = parsed.view_3d.unwrap();
        assert_eq!(v.rotate_x, Some(15));
        assert_eq!(v.rotate_y, Some(20));
        assert_eq!(v.depth_percent, Some(100));
        assert_eq!(v.height_percent, Some(150));
        assert_eq!(v.perspective, Some(30));
        assert_eq!(v.right_angle_axes, Some(true));
    }

    #[test]
    fn test_parse_plot_visible_only_and_display_blanks() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.plot_visible_only, Some(true));
        assert_eq!(parsed.display_blanks_as, Some(DisplayBlanksAs::Gap));
    }

    #[test]
    fn test_parse_manual_layout() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout>
        <c:manualLayout>
          <c:x val="0.1"/>
          <c:y val="0.2"/>
          <c:w val="0.8"/>
          <c:h val="0.6"/>
        </c:manualLayout>
      </c:layout>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let layout = parsed.layout.unwrap();
        let ml = layout.manual_layout.unwrap();
        assert_eq!(ml.x, Some(0.1));
        assert_eq!(ml.y, Some(0.2));
        assert_eq!(ml.width, Some(0.8));
        assert_eq!(ml.height, Some(0.6));
    }

    #[test]
    fn test_parse_data_table() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:dTable>
        <c:showHorzBorder val="1"/>
        <c:showVertBorder val="0"/>
        <c:showOutline val="1"/>
        <c:showKeys val="1"/>
      </c:dTable>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let dt = parsed.data_table.unwrap();
        assert_eq!(dt.show_horizontal_border, Some(true));
        assert_eq!(dt.show_vertical_border, Some(false));
        assert_eq!(dt.show_outline, Some(true));
        assert_eq!(dt.show_keys, Some(true));
    }

    #[test]
    fn test_parse_trendline_types() {
        assert_eq!(parse_trendline_type("linear"), Some(TrendlineType::Linear));
        assert_eq!(
            parse_trendline_type("exp"),
            Some(TrendlineType::Exponential)
        );
        assert_eq!(
            parse_trendline_type("log"),
            Some(TrendlineType::Logarithmic)
        );
        assert_eq!(
            parse_trendline_type("movingAvg"),
            Some(TrendlineType::MovingAverage)
        );
        assert_eq!(
            parse_trendline_type("poly"),
            Some(TrendlineType::Polynomial)
        );
        assert_eq!(parse_trendline_type("power"), Some(TrendlineType::Power));
        assert_eq!(parse_trendline_type("unknown"), None);
    }

    #[test]
    fn test_parse_marker_symbols() {
        assert_eq!(parse_marker_symbol("circle"), Some(MarkerSymbol::Circle));
        assert_eq!(parse_marker_symbol("diamond"), Some(MarkerSymbol::Diamond));
        assert_eq!(parse_marker_symbol("none"), Some(MarkerSymbol::None));
        assert_eq!(parse_marker_symbol("square"), Some(MarkerSymbol::Square));
        assert_eq!(
            parse_marker_symbol("triangle"),
            Some(MarkerSymbol::Triangle)
        );
        assert_eq!(parse_marker_symbol("x"), Some(MarkerSymbol::X));
        assert_eq!(parse_marker_symbol("auto"), Some(MarkerSymbol::Auto));
        assert_eq!(parse_marker_symbol("bogus"), None);
    }

    #[test]
    fn test_parse_data_label_positions() {
        assert_eq!(
            parse_data_label_position("bestFit"),
            Some(DataLabelPosition::BestFit)
        );
        assert_eq!(
            parse_data_label_position("ctr"),
            Some(DataLabelPosition::Center)
        );
        assert_eq!(
            parse_data_label_position("outEnd"),
            Some(DataLabelPosition::OutsideEnd)
        );
        assert_eq!(
            parse_data_label_position("inEnd"),
            Some(DataLabelPosition::InsideEnd)
        );
        assert_eq!(
            parse_data_label_position("inBase"),
            Some(DataLabelPosition::InsideBase)
        );
        assert_eq!(parse_data_label_position("t"), Some(DataLabelPosition::Top));
        assert_eq!(
            parse_data_label_position("b"),
            Some(DataLabelPosition::Bottom)
        );
        assert_eq!(
            parse_data_label_position("l"),
            Some(DataLabelPosition::Left)
        );
        assert_eq!(
            parse_data_label_position("r"),
            Some(DataLabelPosition::Right)
        );
        assert_eq!(parse_data_label_position("unknown"), None);
    }

    #[test]
    fn test_parse_tick_marks_and_positions() {
        assert_eq!(parse_tick_mark("cross"), Some(TickMark::Cross));
        assert_eq!(parse_tick_mark("in"), Some(TickMark::Inside));
        assert_eq!(parse_tick_mark("none"), Some(TickMark::None));
        assert_eq!(parse_tick_mark("out"), Some(TickMark::Outside));
        assert_eq!(parse_tick_mark("bad"), None);

        assert_eq!(
            parse_tick_label_position("high"),
            Some(TickLabelPosition::High)
        );
        assert_eq!(
            parse_tick_label_position("low"),
            Some(TickLabelPosition::Low)
        );
        assert_eq!(
            parse_tick_label_position("nextTo"),
            Some(TickLabelPosition::NextTo)
        );
        assert_eq!(
            parse_tick_label_position("none"),
            Some(TickLabelPosition::None)
        );
        assert_eq!(parse_tick_label_position("bad"), None);
    }

    #[test]
    fn test_parse_moving_average_trendline() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$10</c:f></c:numRef></c:val>
          <c:trendline>
            <c:trendlineType val="movingAvg"/>
            <c:period val="3"/>
          </c:trendline>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let t = parsed.series[0].trendline.as_ref().unwrap();
        assert_eq!(t.trendline_type, TrendlineType::MovingAverage);
        assert_eq!(t.period, Some(3));
        assert!(t.order.is_none());
    }

    #[test]
    fn test_parse_error_bars_custom_type() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:errBars>
            <c:errDir val="x"/>
            <c:errBarType val="plus"/>
            <c:errValType val="stdErr"/>
            <c:noEndCap val="0"/>
          </c:errBars>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let eb = parsed.series[0].error_bars.as_ref().unwrap();
        assert_eq!(eb.direction, ErrorBarDirection::X);
        assert_eq!(eb.bar_type, ErrorBarType::Plus);
        assert_eq!(eb.value_type, ErrorValueType::StandardError);
        assert!(eb.value.is_none());
        assert_eq!(eb.no_end_cap, Some(false));
    }

    #[test]
    fn test_parse_display_blanks_as_zero() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.display_blanks_as, Some(DisplayBlanksAs::Zero));
    }

    #[test]
    fn test_parse_gridlines_as_start_end_element() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:majorGridlines>
          <c:spPr/>
        </c:majorGridlines>
        <c:minorGridlines>
          <c:spPr/>
        </c:minorGridlines>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let val = parsed.value_axis.unwrap();
        assert!(val.major_gridlines);
        assert!(val.minor_gridlines);
    }

    #[test]
    fn test_parse_shape_properties_on_series() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:spPr>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            <a:ln w="25400">
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
              <a:prstDash val="dash"/>
            </a:ln>
          </c:spPr>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let sp = parsed.series[0].shape_properties.as_ref().unwrap();
        assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "FF0000");
        assert!(!sp.no_fill);
        let line = sp.line.as_ref().unwrap();
        assert_eq!(line.width, Some(25400));
        assert_eq!(line.solid_fill.as_ref().unwrap().hex, "000000");
        assert_eq!(line.dash_style.as_deref(), Some("dash"));
    }

    #[test]
    fn test_parse_shape_properties_no_fill() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:spPr>
            <a:noFill/>
          </c:spPr>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let sp = parsed.series[0].shape_properties.as_ref().unwrap();
        assert!(sp.no_fill);
        assert!(sp.solid_fill.is_none());
    }

    #[test]
    fn test_parse_shape_properties_on_axis() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:spPr>
          <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
        </c:spPr>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let sp = parsed
            .value_axis
            .as_ref()
            .unwrap()
            .shape_properties
            .as_ref()
            .unwrap();
        assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "0000FF");
    }

    #[test]
    fn test_parse_ax_pos() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:axPos val="b"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:axPos val="l"/>
        <c:majorUnit val="5.0"/>
        <c:minorUnit val="1.0"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let cat = parsed.category_axis.as_ref().unwrap();
        assert_eq!(cat.position, AxisPosition::Bottom);
        let val = parsed.value_axis.as_ref().unwrap();
        assert_eq!(val.position, AxisPosition::Left);
        assert_eq!(val.major_unit, Some(5.0));
        assert_eq!(val.minor_unit, Some(1.0));
    }

    #[test]
    fn test_parse_legend_overlay() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="1"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let legend = parsed.legend.as_ref().unwrap();
        assert_eq!(legend.position, LegendPosition::Right);
        assert!(legend.overlay);
    }

    #[test]
    fn test_parse_vary_colors_gap_width_overlap() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/>
        <c:overlap val="-25"/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.vary_colors, Some(true));
        assert_eq!(parsed.gap_width, Some(150));
        assert_eq!(parsed.overlap, Some(-25));
    }

    #[test]
    fn test_extlst_capture() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:extLst>
            <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}">
              <c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{00000000-1111-2222-3333-444444444444}"/>
            </c:ext>
          </c:extLst>
        </c:ser>
        <c:extLst>
          <c:ext uri="{chart-type-ext}"><dummy/></c:ext>
        </c:extLst>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:extLst>
          <c:ext uri="{cat-ax-ext}"><axData/></c:ext>
        </c:extLst>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:extLst>
          <c:ext uri="{val-ax-ext}"><axData2/></c:ext>
        </c:extLst>
      </c:valAx>
      <c:extLst>
        <c:ext uri="{plot-area-ext}"><plotData/></c:ext>
      </c:extLst>
    </c:plotArea>
    <c:extLst>
      <c:ext uri="{chart-ext}"><chartData/></c:ext>
    </c:extLst>
  </c:chart>
  <c:extLst>
    <c:ext uri="{chart-space-ext}"><spaceData/></c:ext>
  </c:extLst>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);

        // Verify extensions were captured at each level
        let ser_ext = parsed.series[0]
            .raw_ext
            .as_ref()
            .expect("series extLst missing");
        let ser_ext_str = std::str::from_utf8(ser_ext).unwrap();
        assert!(
            ser_ext_str.contains("00000000-1111-2222-3333-444444444444"),
            "series extLst content not preserved: {}",
            ser_ext_str
        );

        let ct_ext = parsed
            .raw_extensions
            .get("chartType")
            .expect("chartType extLst missing");
        assert!(std::str::from_utf8(ct_ext)
            .unwrap()
            .contains("chart-type-ext"));

        let cat_ax = parsed.category_axis.as_ref().unwrap();
        let cat_ext = cat_ax.raw_ext.as_ref().expect("catAx extLst missing");
        assert!(std::str::from_utf8(cat_ext).unwrap().contains("cat-ax-ext"));

        let val_ax = parsed.value_axis.as_ref().unwrap();
        let val_ext = val_ax.raw_ext.as_ref().expect("valAx extLst missing");
        assert!(std::str::from_utf8(val_ext).unwrap().contains("val-ax-ext"));

        let pa_ext = parsed
            .raw_extensions
            .get("plotArea")
            .expect("plotArea extLst missing");
        assert!(std::str::from_utf8(pa_ext)
            .unwrap()
            .contains("plot-area-ext"));

        let ch_ext = parsed
            .raw_extensions
            .get("chart")
            .expect("chart extLst missing");
        assert!(std::str::from_utf8(ch_ext).unwrap().contains("chart-ext"));

        let cs_ext = parsed
            .raw_extensions
            .get("chartSpace")
            .expect("chartSpace extLst missing");
        assert!(std::str::from_utf8(cs_ext)
            .unwrap()
            .contains("chart-space-ext"));
    }

    #[test]
    fn test_parse_3d_bar_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::ColumnClustered);
        assert!(parsed.is_3d);
    }

    #[test]
    fn test_parse_3d_line_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:line3DChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:line3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Line);
        assert!(parsed.is_3d);
    }

    #[test]
    fn test_parse_3d_pie_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pie3DChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Pie);
        assert!(parsed.is_3d);
    }

    #[test]
    fn test_parse_3d_surface_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:surface3DChart>
        <c:wireframe val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:surface3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Surface);
        assert!(parsed.is_3d);
        assert_eq!(parsed.wireframe, Some(true));
    }

    #[test]
    fn test_parse_2d_chart_not_3d() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert!(!parsed.is_3d);
    }

    #[test]
    fn test_parse_first_slice_angle_and_hole_size() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:doughnutChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="90"/>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Doughnut);
        assert_eq!(parsed.first_slice_angle, Some(90));
        assert_eq!(parsed.hole_size, Some(50));
    }

    #[test]
    fn test_parse_bubble_scale_and_neg_bubbles() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:bubbleChart>
        <c:ser>
          <c:idx val="0"/>
          <c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f></c:numRef></c:yVal>
        </c:ser>
        <c:bubbleScale val="75"/>
        <c:showNegBubbles val="0"/>
      </c:bubbleChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Bubble);
        assert_eq!(parsed.bubble_scale, Some(75));
        assert_eq!(parsed.show_negative_bubbles, Some(false));
    }

    #[test]
    fn test_parse_radar_style() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:radarChart>
        <c:radarStyle val="filled"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:radarChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::Radar);
        assert_eq!(parsed.radar_style.as_deref(), Some("filled"));
    }

    #[test]
    fn test_parse_auto_title_deleted() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.auto_title_deleted, Some(true));
    }

    #[test]
    fn test_parse_rounded_corners() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.rounded_corners, Some(true));
    }

    #[test]
    fn test_parse_show_leader_lines() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:dLbls>
            <c:showVal val="1"/>
            <c:showLeaderLines val="1"/>
          </c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let dl = parsed.series[0].data_labels.as_ref().unwrap();
        assert_eq!(dl.show_leader_lines, Some(true));
    }

    #[test]
    fn test_parse_invert_if_negative() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:invertIfNegative val="1"/>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.series[0].invert_if_negative, Some(true));
    }

    #[test]
    fn test_parse_show_dlbls_over_max() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:showDLblsOverMax val="1"/>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.show_dlbls_over_max, Some(true));
    }

    #[test]
    fn test_parse_area_3d_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:area3DChart>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:area3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        assert_eq!(parsed.chart_type, ChartType::AreaStacked);
        assert!(parsed.is_3d);
    }

    #[test]
    fn test_parse_date_ax_sets_axis_type() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:dateAx>
        <c:axId val="1"/>
        <c:delete val="0"/>
      </c:dateAx>
      <c:valAx>
        <c:axId val="2"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let cat = parsed.category_axis.unwrap();
        assert_eq!(cat.axis_type, AxisType::Date);
        let val = parsed.value_axis.unwrap();
        assert_eq!(val.axis_type, AxisType::Value);
    }

    #[test]
    fn test_parse_cat_ax_defaults_to_category() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let cat = parsed.category_axis.unwrap();
        assert_eq!(cat.axis_type, AxisType::Category);
    }

    #[test]
    fn test_parse_ser_ax() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:bar3DChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
      <c:serAx>
        <c:axId val="3"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Series</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:delete val="0"/>
      </c:serAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let parsed = parse_chart_xml_str(xml);
        let ser = parsed.series_axis.unwrap();
        assert_eq!(ser.axis_type, AxisType::Series);
        assert_eq!(ser.title.as_deref(), Some("Series"));
        assert_eq!(ser.delete, Some(false));
    }
}
