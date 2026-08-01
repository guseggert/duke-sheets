use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::error::{ChartParseError, ChartParseResult};
use crate::{
    Axis, AxisCrosses, AxisPosition, AxisType, Chart, ChartAxis, ChartColor, ChartDataTable,
    ChartLine, ChartLines, ChartShapeProperties, ChartType, ChartTypeGroup, CrossBetween,
    DataLabelPosition, DataLabels, DataPoint, DataReference, DataSeries, DisplayBlanksAs,
    ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType, Layout, Legend, LegendPosition,
    ManualLayout, Marker, MarkerSymbol, NumberFormat, TickLabelPosition, TickMark, Trendline,
    TrendlineType, UpDownBars, View3D,
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
    // Whitespace inside a text-bearing element is part of the value: a
    // category can be named " Q1 " and a label separator can be " | ".
    // The handlers below only take text while one of them is open, so
    // the indentation between elements is ignored without trimming it
    // away.
    xml_reader.config_mut().trim_text(false);

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

/// Append to a text target. A single element's content can arrive as
/// more than one event, so assigning would keep only the last piece.
fn push_text(slot: &mut Option<String>, t: &str) {
    match slot {
        Some(existing) => existing.push_str(t),
        None => *slot = Some(t.to_string()),
    }
}

/// Streaming state for one standard chart part.
///
/// The document is deep and its elements are context sensitive, so
/// parsing is a state machine rather than a recursive descent. One
/// value holds it so each event kind has a single handler, which is
/// what stops an element being handled differently depending on
/// whether it was written self-closing.
struct ChartParser {
    result: ParsedChart,
    in_chart: bool,
    in_plot_area: bool,
    in_chart_type_element: bool,
    chart_type_tag: Option<String>,
    bar_dir: Option<String>,
    grouping: Option<String>,
    scatter_style: Option<String>,
    vary_colors: Option<bool>,
    gap_width: Option<u32>,
    overlap: Option<i32>,
    first_slice_angle: Option<u32>,
    hole_size: Option<u32>,
    bubble_scale: Option<u32>,
    show_negative_bubbles: Option<bool>,
    radar_style: Option<String>,
    wireframe: Option<bool>,
    in_chart_space: bool,
    group_series: Vec<DataSeries>,
    group_data_labels: Option<DataLabels>,
    group_axis_ids: Vec<u32>,
    group_raw_ext: Option<Vec<u8>>,
    group_drop_lines: Option<ChartLines>,
    group_high_low_lines: Option<ChartLines>,
    group_series_lines: Option<ChartLines>,
    group_up_down_bars: Option<UpDownBars>,
    in_drop_lines: bool,
    in_hi_low_lines: bool,
    in_ser_lines: bool,
    in_up_down_bars: bool,
    in_up_bars: bool,
    in_down_bars: bool,
    in_leader_lines: bool,
    up_down_bars_gap_width: Option<u32>,
    up_bars_sp: Option<ChartShapeProperties>,
    down_bars_sp: Option<ChartShapeProperties>,
    had_up_bars: bool,
    had_down_bars: bool,
    in_chart_title: bool,
    in_title_tx: bool,
    in_title_rich: bool,
    in_title_p: bool,
    in_title_r: bool,
    in_title_t: bool,
    title_text: String,
    in_title_str_ref: bool,
    in_title_str_ref_f: bool,
    title_depth: u32,
    in_ser: bool,
    in_ser_tx: bool,
    in_ser_tx_str_ref: bool,
    in_ser_tx_str_ref_f: bool,
    in_ser_tx_v: bool,
    ser_name: Option<String>,
    in_ser_val: bool,
    in_ser_yval: bool,
    in_ser_val_num_ref: bool,
    in_ser_val_num_ref_f: bool,
    ser_val_formula: Option<String>,
    in_ser_val_num_cache: bool,
    in_ser_val_pt: bool,
    in_ser_val_pt_v: bool,
    ser_val_cache: Vec<f64>,
    in_ser_cat: bool,
    in_ser_xval: bool,
    in_ser_cat_ref: bool,
    in_ser_cat_ref_f: bool,
    ser_cat_formula: Option<String>,
    ser_smooth: Option<bool>,
    ser_explosion: Option<u32>,
    ser_data_labels: Option<DataLabels>,
    ser_marker: Option<Marker>,
    ser_trendline: Option<Trendline>,
    ser_error_bars: Option<ErrorBars>,
    ser_shape_properties: Option<ChartShapeProperties>,
    ser_raw_ext: Option<Vec<u8>>,
    ser_invert_if_negative: Option<bool>,
    in_dlbls: bool,
    in_dlbls_separator: bool,
    dlbls: DataLabels,
    in_dpt: bool,
    dpt_index: u32,
    dpt_explosion: Option<u32>,
    dpt_marker: Option<Marker>,
    dpt_shape_properties: Option<ChartShapeProperties>,
    ser_data_points: Vec<DataPoint>,
    in_marker: bool,
    marker_symbol: Option<MarkerSymbol>,
    marker_size: Option<u8>,
    in_trendline: bool,
    in_trendline_name: bool,
    trendline_type: Option<TrendlineType>,
    trendline_name: Option<String>,
    trendline_order: Option<u32>,
    trendline_period: Option<u32>,
    trendline_forward: Option<f64>,
    trendline_backward: Option<f64>,
    trendline_intercept: Option<f64>,
    trendline_disp_r_sqr: Option<bool>,
    trendline_disp_eq: Option<bool>,
    in_err_bars: bool,
    err_dir: Option<ErrorBarDirection>,
    err_bar_type: Option<ErrorBarType>,
    err_val_type: Option<ErrorValueType>,
    err_val: Option<f64>,
    err_no_end_cap: Option<bool>,
    in_cat_ax: bool,
    in_val_ax: bool,
    in_ser_ax: bool,
    is_date_ax: bool,
    in_ax_title: bool,
    in_ax_title_tx: bool,
    in_ax_title_rich: bool,
    in_ax_title_p: bool,
    in_ax_title_r: bool,
    in_ax_title_t: bool,
    ax_title_text: String,
    in_ax_scaling: bool,
    ax_min: Option<f64>,
    ax_max: Option<f64>,
    ax_number_format: Option<NumberFormat>,
    ax_major_gridlines: bool,
    ax_minor_gridlines: bool,
    in_ax_major_gridlines: bool,
    in_ax_minor_gridlines: bool,
    ax_major_gridlines_shape_properties: Option<ChartShapeProperties>,
    ax_minor_gridlines_shape_properties: Option<ChartShapeProperties>,
    ax_major_tick_mark: Option<TickMark>,
    ax_minor_tick_mark: Option<TickMark>,
    ax_label_position: Option<TickLabelPosition>,
    ax_delete: Option<bool>,
    ax_crosses: Option<AxisCrosses>,
    ax_cross_between: Option<CrossBetween>,
    ax_position: Option<AxisPosition>,
    ax_major_unit: Option<f64>,
    ax_minor_unit: Option<f64>,
    ax_shape_properties: Option<ChartShapeProperties>,
    ax_raw_ext: Option<Vec<u8>>,
    ax_id: Option<u32>,
    ax_cross_id: Option<u32>,
    in_legend: bool,
    legend_pos: Option<LegendPosition>,
    legend_overlay: Option<bool>,
    legend_shape_properties: Option<ChartShapeProperties>,
    in_view_3d: bool,
    view_3d: View3D,
    in_layout: bool,
    in_manual_layout: bool,
    had_manual_layout: bool,
    manual_layout: ManualLayout,
    ext_writer: Option<Writer<Cursor<Vec<u8>>>>,
    ext_depth: u32,
    ext_dest: ExtLstDest,
    in_sp_pr: bool,
    sp_pr_depth: u32,
    sp_solid_fill: Option<ChartColor>,
    sp_no_fill: bool,
    sp_line: Option<ChartLine>,
    in_sp_ln: bool,
    sp_ln_width: Option<i64>,
    sp_ln_solid_fill: Option<ChartColor>,
    sp_ln_no_fill: bool,
    sp_ln_dash: Option<String>,
    sp_pr_context: SpPrContext,
    in_d_table: bool,
    d_table: ChartDataTable,
}

impl ChartParser {
    fn new() -> Self {
        Self {
            result: ParsedChart {
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
                show_dlbls_over_max: None,
                wireframe: None,
                drop_lines: None,
                high_low_lines: None,
                up_down_bars: None,
                series_lines: None,
                type_groups: Vec::new(),
                axes: Vec::new(),
            },
            in_chart: false,
            in_plot_area: false,
            in_chart_type_element: false,
            chart_type_tag: None,
            bar_dir: None,
            grouping: None,
            scatter_style: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            wireframe: None,
            in_chart_space: false,
            group_series: Vec::new(),
            group_data_labels: None,
            group_axis_ids: Vec::new(),
            group_raw_ext: None,
            group_drop_lines: None,
            group_high_low_lines: None,
            group_series_lines: None,
            group_up_down_bars: None,
            in_drop_lines: false,
            in_hi_low_lines: false,
            in_ser_lines: false,
            in_up_down_bars: false,
            in_up_bars: false,
            in_down_bars: false,
            in_leader_lines: false,
            up_down_bars_gap_width: None,
            up_bars_sp: None,
            down_bars_sp: None,
            had_up_bars: false,
            had_down_bars: false,
            in_chart_title: false,
            in_title_tx: false,
            in_title_rich: false,
            in_title_p: false,
            in_title_r: false,
            in_title_t: false,
            title_text: String::new(),
            in_title_str_ref: false,
            in_title_str_ref_f: false,
            title_depth: 0u32,
            in_ser: false,
            in_ser_tx: false,
            in_ser_tx_str_ref: false,
            in_ser_tx_str_ref_f: false,
            in_ser_tx_v: false,
            ser_name: None,
            in_ser_val: false,
            in_ser_yval: false,
            in_ser_val_num_ref: false,
            in_ser_val_num_ref_f: false,
            ser_val_formula: None,
            in_ser_val_num_cache: false,
            in_ser_val_pt: false,
            in_ser_val_pt_v: false,
            ser_val_cache: Vec::new(),
            in_ser_cat: false,
            in_ser_xval: false,
            in_ser_cat_ref: false,
            in_ser_cat_ref_f: false,
            ser_cat_formula: None,
            ser_smooth: None,
            ser_explosion: None,
            ser_data_labels: None,
            ser_marker: None,
            ser_trendline: None,
            ser_error_bars: None,
            ser_shape_properties: None,
            ser_raw_ext: None,
            ser_invert_if_negative: None,
            in_dlbls: false,
            in_dlbls_separator: false,
            dlbls: DataLabels::default(),
            in_dpt: false,
            dpt_index: 0,
            dpt_explosion: None,
            dpt_marker: None,
            dpt_shape_properties: None,
            ser_data_points: Vec::new(),
            in_marker: false,
            marker_symbol: None,
            marker_size: None,
            in_trendline: false,
            in_trendline_name: false,
            trendline_type: None,
            trendline_name: None,
            trendline_order: None,
            trendline_period: None,
            trendline_forward: None,
            trendline_backward: None,
            trendline_intercept: None,
            trendline_disp_r_sqr: None,
            trendline_disp_eq: None,
            in_err_bars: false,
            err_dir: None,
            err_bar_type: None,
            err_val_type: None,
            err_val: None,
            err_no_end_cap: None,
            in_cat_ax: false,
            in_val_ax: false,
            in_ser_ax: false,
            is_date_ax: false,
            in_ax_title: false,
            in_ax_title_tx: false,
            in_ax_title_rich: false,
            in_ax_title_p: false,
            in_ax_title_r: false,
            in_ax_title_t: false,
            ax_title_text: String::new(),
            in_ax_scaling: false,
            ax_min: None,
            ax_max: None,
            ax_number_format: None,
            ax_major_gridlines: false,
            ax_minor_gridlines: false,
            in_ax_major_gridlines: false,
            in_ax_minor_gridlines: false,
            ax_major_gridlines_shape_properties: None,
            ax_minor_gridlines_shape_properties: None,
            ax_major_tick_mark: None,
            ax_minor_tick_mark: None,
            ax_label_position: None,
            ax_delete: None,
            ax_crosses: None,
            ax_cross_between: None,
            ax_position: None,
            ax_major_unit: None,
            ax_minor_unit: None,
            ax_shape_properties: None,
            ax_raw_ext: None,
            ax_id: None,
            ax_cross_id: None,
            in_legend: false,
            legend_pos: None,
            legend_overlay: None,
            legend_shape_properties: None,
            in_view_3d: false,
            view_3d: View3D::default(),
            in_layout: false,
            in_manual_layout: false,
            had_manual_layout: false,
            manual_layout: ManualLayout::default(),
            ext_writer: None,
            ext_depth: 0u32,
            ext_dest: ExtLstDest::ChartSpace,
            in_sp_pr: false,
            sp_pr_depth: 0u32,
            sp_solid_fill: None,
            sp_no_fill: false,
            sp_line: None,
            in_sp_ln: false,
            sp_ln_width: None,
            sp_ln_solid_fill: None,
            sp_ln_no_fill: false,
            sp_ln_dash: None,
            sp_pr_context: SpPrContext::None,
            in_d_table: false,
            d_table: ChartDataTable::default(),
        }
    }

    /// True while an extLst is being kept as bytes.
    fn capturing(&self) -> bool {
        self.ext_writer.is_some()
    }

    /// Forward an event into the open capture, closing it at depth zero.
    fn capture(&mut self, ev: &Event) {
    let done = match ev {
        Event::Start(_) => {
            self.ext_depth += 1;
            false
        }
        Event::End(_) => {
            self.ext_depth -= 1;
            self.ext_depth == 0
        }
        Event::Eof => true,
        _ => false,
    };
    if let Some(w) = self.ext_writer.as_mut() {
        if !matches!(ev, Event::Eof) {
            let _ = w.write_event(ev.clone());
        }
    }
    if done {
        if let Some(w) = self.ext_writer.take() {
            let raw = w.into_inner().into_inner();
            match self.ext_dest {
                ExtLstDest::Series => self.ser_raw_ext = Some(raw),
                ExtLstDest::Axis => self.ax_raw_ext = Some(raw),
                ExtLstDest::TypeGroup => self.group_raw_ext = Some(raw),
                ExtLstDest::Chart => {
                    self.result.raw_extensions.insert("chart".into(), raw);
                }
                ExtLstDest::PlotArea => {
                    self.result.raw_extensions.insert("plotArea".into(), raw);
                }
                ExtLstDest::ChartSpace => {
                    self.result.raw_extensions.insert("chartSpace".into(), raw);
                }
            }
        }
    }
    }

    fn on_start(&mut self, e: &BytesStart) {
        let local = e.name().local_name();
        let tag = local.as_ref();

        // Inside a chart title or a spPr, every element start deepens the
        // nesting and its end unwinds it, including the end synthesized
        // for a self-closing element. Applied after the match, not
        // before, so a guard still sees the depth its own element sits
        // at, and tracked here rather than in individual arms so that
        // adding an arm cannot unbalance it.
        let deepen_title = self.in_chart_title;
        let deepen_sp_pr = self.in_sp_pr;

        match tag {
        b"chartSpace" => self.in_chart_space = true,
        b"chart" if !self.in_chart => self.in_chart = true,
        b"plotArea" if self.in_chart => self.in_plot_area = true,
        b"title"
            if self.in_chart
                && !self.in_plot_area
                && !self.in_chart_title
                && !self.in_cat_ax
                && !self.in_val_ax
                && !self.in_ser_ax =>
        {
            self.in_chart_title = true;
            self.title_depth = 1;
            self.title_text.clear();
        }
        b"tx" if self.in_chart_title && self.title_depth == 1 => self.in_title_tx = true,
        b"rich" if self.in_title_tx => self.in_title_rich = true,
        b"p" if self.in_title_rich => self.in_title_p = true,
        b"r" if self.in_title_p => self.in_title_r = true,
        b"t" if self.in_title_r => self.in_title_t = true,
        b"strRef" if self.in_title_tx => self.in_title_str_ref = true,
        b"f" if self.in_title_str_ref => self.in_title_str_ref_f = true,
        // View 3D
        b"view3D" if self.in_chart && !self.in_plot_area => {
            self.in_view_3d = true;
            self.view_3d = View3D::default();
        }
        // Chart type elements in plotArea
        b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
        | b"pie3DChart" | b"doughnutChart" | b"areaChart" | b"area3DChart"
        | b"scatterChart" | b"bubbleChart" | b"radarChart" | b"stockChart"
        | b"surfaceChart" | b"surface3DChart" | b"ofPieChart"
            if self.in_plot_area && !self.in_chart_type_element =>
        {
            self.in_chart_type_element = true;
            let tag_str = std::str::from_utf8(tag).unwrap_or("unknown");
            self.chart_type_tag = Some(tag_str.to_string());
            self.bar_dir = None;
            self.grouping = None;
            self.scatter_style = None;
            self.group_series.clear();
            self.group_data_labels = None;
            self.group_axis_ids.clear();
            self.group_raw_ext = None;
        }
        // Layout
        b"layout" if self.in_plot_area && !self.in_chart_type_element => {
            self.in_layout = true;
            self.had_manual_layout = false;
        }
        b"manualLayout" if self.in_layout => {
            self.in_manual_layout = true;
            self.had_manual_layout = true;
            self.manual_layout = ManualLayout::default();
        }
        // Data table
        b"dTable" if self.in_plot_area && !self.in_chart_type_element => {
            self.in_d_table = true;
            self.d_table = ChartDataTable::default();
        }
        // Data labels (chart-level or series-level)
        b"dLbls" if self.in_chart_type_element && !self.in_dlbls => {
            self.in_dlbls = true;
            self.dlbls = DataLabels::default();
        }
        b"separator" if self.in_dlbls => self.in_dlbls_separator = true,
        b"numFmt" if self.in_dlbls => {
            self.dlbls.number_format = Some(parse_num_fmt(e));
        }
        // Chart lines and up-down bars
        b"dropLines" if self.in_chart_type_element && !self.in_ser => {
            self.in_drop_lines = true;
        }
        b"hiLowLines" if self.in_chart_type_element && !self.in_ser => {
            self.in_hi_low_lines = true;
        }
        b"serLines" if self.in_chart_type_element && !self.in_ser => {
            self.in_ser_lines = true;
        }
        b"upDownBars" if self.in_chart_type_element && !self.in_ser => {
            self.in_up_down_bars = true;
            self.up_down_bars_gap_width = None;
            self.up_bars_sp = None;
            self.down_bars_sp = None;
            self.had_up_bars = false;
            self.had_down_bars = false;
        }
        b"upBars" if self.in_up_down_bars => {
            self.in_up_bars = true;
            self.had_up_bars = true;
        }
        b"downBars" if self.in_up_down_bars => {
            self.in_down_bars = true;
            self.had_down_bars = true;
        }
        b"leaderLines" if self.in_dlbls => self.in_leader_lines = true,
        b"ser" if self.in_chart_type_element => {
            self.in_ser = true;
            self.ser_name = None;
            self.ser_val_formula = None;
            self.ser_val_cache.clear();
            self.ser_cat_formula = None;
            self.ser_smooth = None;
            self.ser_explosion = None;
            self.ser_data_labels = None;
            self.ser_marker = None;
            self.ser_trendline = None;
            self.ser_error_bars = None;
            self.ser_data_points.clear();
            self.ser_shape_properties = None;
            self.ser_raw_ext = None;
            self.ser_invert_if_negative = None;
        }
        b"tx" if self.in_ser => self.in_ser_tx = true,
        b"strRef" if self.in_ser_tx => self.in_ser_tx_str_ref = true,
        b"f" if self.in_ser_tx_str_ref => self.in_ser_tx_str_ref_f = true,
        b"v" if self.in_ser_tx && !self.in_ser_tx_str_ref => self.in_ser_tx_v = true,
        // Data points
        b"dPt" if self.in_ser => {
            self.in_dpt = true;
            self.dpt_index = 0;
            self.dpt_explosion = None;
            self.dpt_marker = None;
            self.dpt_shape_properties = None;
        }
        // Trendline
        b"trendline" if self.in_ser => {
            self.in_trendline = true;
            self.trendline_type = None;
            self.trendline_name = None;
            self.trendline_order = None;
            self.trendline_period = None;
            self.trendline_forward = None;
            self.trendline_backward = None;
            self.trendline_intercept = None;
            self.trendline_disp_r_sqr = None;
            self.trendline_disp_eq = None;
        }
        b"name" if self.in_trendline => self.in_trendline_name = true,
        // Error bars
        b"errBars" if self.in_ser => {
            self.in_err_bars = true;
            self.err_dir = None;
            self.err_bar_type = None;
            self.err_val_type = None;
            self.err_val = None;
            self.err_no_end_cap = None;
        }
        // Marker (series or data point level)
        b"marker" if self.in_ser && !self.in_dlbls && !self.in_trendline && !self.in_err_bars => {
            self.in_marker = true;
            self.marker_symbol = None;
            self.marker_size = None;
        }
        b"val" if self.in_err_bars => self.err_val = get_val_f64(e),
        b"val" if self.in_ser && !self.in_err_bars => self.in_ser_val = true,
        b"yVal" if self.in_ser => self.in_ser_yval = true,
        b"numRef" if self.in_ser_val || self.in_ser_yval => self.in_ser_val_num_ref = true,
        b"f" if self.in_ser_val_num_ref => self.in_ser_val_num_ref_f = true,
        b"numCache" if self.in_ser_val_num_ref => self.in_ser_val_num_cache = true,
        b"pt" if self.in_ser_val_num_cache => self.in_ser_val_pt = true,
        b"v" if self.in_ser_val_pt => self.in_ser_val_pt_v = true,
        b"cat" if self.in_ser => self.in_ser_cat = true,
        b"xVal" if self.in_ser => self.in_ser_xval = true,
        b"strRef" | b"numRef" if self.in_ser_cat || self.in_ser_xval => {
            self.in_ser_cat_ref = true;
        }
        b"f" if self.in_ser_cat_ref => self.in_ser_cat_ref_f = true,
        // Axis elements
        b"catAx" | b"dateAx" if self.in_plot_area => {
            if tag == b"catAx" {
            } else {
            }
            self.in_cat_ax = true;
            self.is_date_ax = tag == b"dateAx";
            self.ax_title_text.clear();
            self.ax_min = None;
            self.ax_max = None;
            self.ax_number_format = None;
            self.ax_major_gridlines = false;
            self.ax_minor_gridlines = false;
            self.in_ax_major_gridlines = false;
            self.in_ax_minor_gridlines = false;
            self.ax_major_gridlines_shape_properties = None;
            self.ax_minor_gridlines_shape_properties = None;
            self.ax_major_tick_mark = None;
            self.ax_minor_tick_mark = None;
            self.ax_label_position = None;
            self.ax_delete = None;
            self.ax_crosses = None;
            self.ax_cross_between = None;
            self.ax_position = None;
            self.ax_major_unit = None;
            self.ax_minor_unit = None;
            self.ax_shape_properties = None;
            self.ax_raw_ext = None;
            self.ax_id = None;
            self.ax_cross_id = None;
        }
        b"valAx" if self.in_plot_area => {
            self.in_val_ax = true;
            self.ax_title_text.clear();
            self.ax_min = None;
            self.ax_max = None;
            self.ax_number_format = None;
            self.ax_major_gridlines = false;
            self.ax_minor_gridlines = false;
            self.in_ax_major_gridlines = false;
            self.in_ax_minor_gridlines = false;
            self.ax_major_gridlines_shape_properties = None;
            self.ax_minor_gridlines_shape_properties = None;
            self.ax_major_tick_mark = None;
            self.ax_minor_tick_mark = None;
            self.ax_label_position = None;
            self.ax_delete = None;
            self.ax_crosses = None;
            self.ax_cross_between = None;
            self.ax_position = None;
            self.ax_major_unit = None;
            self.ax_minor_unit = None;
            self.ax_shape_properties = None;
            self.ax_raw_ext = None;
            self.ax_id = None;
            self.ax_cross_id = None;
        }
        b"serAx" if self.in_plot_area => {
            self.in_ser_ax = true;
            self.ax_title_text.clear();
            self.ax_min = None;
            self.ax_max = None;
            self.ax_number_format = None;
            self.ax_major_gridlines = false;
            self.ax_minor_gridlines = false;
            self.in_ax_major_gridlines = false;
            self.in_ax_minor_gridlines = false;
            self.ax_major_gridlines_shape_properties = None;
            self.ax_minor_gridlines_shape_properties = None;
            self.ax_major_tick_mark = None;
            self.ax_minor_tick_mark = None;
            self.ax_label_position = None;
            self.ax_delete = None;
            self.ax_crosses = None;
            self.ax_cross_between = None;
            self.ax_position = None;
            self.ax_major_unit = None;
            self.ax_minor_unit = None;
            self.ax_shape_properties = None;
            self.ax_raw_ext = None;
            self.ax_id = None;
            self.ax_cross_id = None;
        }
        b"title" if self.in_cat_ax || self.in_val_ax || self.in_ser_ax => self.in_ax_title = true,
        b"tx" if self.in_ax_title => self.in_ax_title_tx = true,
        b"rich" if self.in_ax_title_tx => self.in_ax_title_rich = true,
        b"p" if self.in_ax_title_rich => self.in_ax_title_p = true,
        b"r" if self.in_ax_title_p => self.in_ax_title_r = true,
        b"t" if self.in_ax_title_r => self.in_ax_title_t = true,
        b"scaling" if self.in_cat_ax || self.in_val_ax || self.in_ser_ax => {
            self.in_ax_scaling = true;
        }
        b"majorGridlines" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_major_gridlines = true;
            self.in_ax_major_gridlines = true;
        }
        b"minorGridlines" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_minor_gridlines = true;
            self.in_ax_minor_gridlines = true;
        }
        b"numFmt"
            if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title && !self.in_dlbls =>
        {
            self.ax_number_format = Some(parse_num_fmt(e));
        }
        // Legend
        b"legend" if self.in_chart && !self.in_plot_area => {
            self.in_legend = true;
            self.legend_pos = None;
        }
        b"spPr" if !self.in_sp_pr => {
            self.in_sp_pr = true;
            self.sp_pr_depth = 1;
            self.sp_solid_fill = None;
            self.sp_no_fill = false;
            self.sp_line = None;
            self.in_sp_ln = false;
            self.sp_ln_width = None;
            self.sp_ln_solid_fill = None;
            self.sp_ln_no_fill = false;
            self.sp_ln_dash = None;
            if self.in_drop_lines {
                self.sp_pr_context = SpPrContext::DropLines;
            } else if self.in_hi_low_lines {
                self.sp_pr_context = SpPrContext::HiLowLines;
            } else if self.in_ser_lines {
                self.sp_pr_context = SpPrContext::SerLines;
            } else if self.in_up_bars {
                self.sp_pr_context = SpPrContext::UpBars;
            } else if self.in_down_bars {
                self.sp_pr_context = SpPrContext::DownBars;
            } else if self.in_leader_lines {
                self.sp_pr_context = SpPrContext::LeaderLines;
            } else if self.in_dpt && !self.in_marker && !self.in_dlbls {
                self.sp_pr_context = SpPrContext::DataPoint;
            } else if self.in_ser
                && !self.in_dpt
                && !self.in_trendline
                && !self.in_err_bars
                && !self.in_marker
                && !self.in_dlbls
            {
                self.sp_pr_context = SpPrContext::Series;
            } else if self.in_ax_major_gridlines {
                self.sp_pr_context = SpPrContext::MajorGridlines;
            } else if self.in_ax_minor_gridlines {
                self.sp_pr_context = SpPrContext::MinorGridlines;
            } else if self.in_cat_ax || self.in_ser_ax {
                self.sp_pr_context = SpPrContext::CatAxis;
            } else if self.in_val_ax {
                self.sp_pr_context = SpPrContext::ValAxis;
            } else if self.in_legend {
                self.sp_pr_context = SpPrContext::Legend;
            } else if self.in_chart_space && !self.in_chart && !self.in_plot_area {
                self.sp_pr_context = SpPrContext::ChartSpace;
            } else {
                self.sp_pr_context = SpPrContext::None;
            }
        }
        b"ln" if self.in_sp_pr => {
            self.in_sp_ln = true;
            self.sp_ln_width = None;
            self.sp_ln_solid_fill = None;
            self.sp_ln_no_fill = false;
            self.sp_ln_dash = None;
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"w" {
                    self.sp_ln_width = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<i64>().ok());
                }
            }
        }
        b"extLst" => {
            self.ext_dest = if self.in_ser {
                ExtLstDest::Series
            } else if self.in_cat_ax || self.in_val_ax || self.in_ser_ax {
                ExtLstDest::Axis
            } else if self.in_chart_type_element {
                ExtLstDest::TypeGroup
            } else if self.in_chart && !self.in_plot_area {
                ExtLstDest::Chart
            } else if self.in_plot_area {
                ExtLstDest::PlotArea
            } else {
                ExtLstDest::ChartSpace
            };
            let mut w = Writer::new(Cursor::new(Vec::new()));
            let _ = w.write_event(Event::Start(e.to_owned()));
            self.ext_depth = 1;
            self.ext_writer = Some(w);
        }
        // Handling merged from the former separate arm for empty
        // elements, which a self-closing element no longer reaches.
        b"barDir" if self.in_chart_type_element && !self.in_ser => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    self.bar_dir = attr.unescape_value().ok().map(|s| s.to_string());
                }
            }
        }
        b"grouping" if self.in_chart_type_element && !self.in_ser => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    self.grouping = attr.unescape_value().ok().map(|s| s.to_string());
                }
            }
        }
        b"scatterStyle" if self.in_chart_type_element && !self.in_ser => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    self.scatter_style = attr.unescape_value().ok().map(|s| s.to_string());
                }
            }
        }
        b"radarStyle" if self.in_chart_type_element && !self.in_ser => {
            self.radar_style = get_val_attr(e);
        }
        b"firstSliceAng" if self.in_chart_type_element && !self.in_ser => {
            self.first_slice_angle = get_val_u32(e);
        }
        b"holeSize" if self.in_chart_type_element && !self.in_ser => {
            self.hole_size = get_val_u32(e);
        }
        b"bubbleScale" if self.in_chart_type_element && !self.in_ser => {
            self.bubble_scale = get_val_u32(e);
        }
        b"showNegBubbles" if self.in_chart_type_element && !self.in_ser => {
            self.show_negative_bubbles = get_val_bool(e);
        }
        b"wireframe" if self.in_chart_type_element && !self.in_ser => {
            self.wireframe = get_val_bool(e);
        }
        b"legendPos" if self.in_legend => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    if let Ok(val) = attr.unescape_value() {
                        self.legend_pos = Some(match val.as_ref() {
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
        b"min" if self.in_ax_scaling => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    self.ax_min = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<f64>().ok());
                }
            }
        }
        b"max" if self.in_ax_scaling => {
            for attr in e.attributes().flatten() {
                if attr.key.local_name().as_ref() == b"val" {
                    self.ax_max = attr
                        .unescape_value()
                        .ok()
                        .and_then(|s| s.parse::<f64>().ok());
                }
            }
        }
        // Data label children
        b"showLegendKey" if self.in_dlbls => {
            self.dlbls.show_legend_key = get_val_bool(e);
        }
        b"showVal" if self.in_dlbls => self.dlbls.show_value = get_val_bool(e),
        b"showCatName" if self.in_dlbls => {
            self.dlbls.show_category_name = get_val_bool(e);
        }
        b"showSerName" if self.in_dlbls => {
            self.dlbls.show_series_name = get_val_bool(e);
        }
        b"showPercent" if self.in_dlbls => self.dlbls.show_percent = get_val_bool(e),
        b"showBubbleSize" if self.in_dlbls => {
            self.dlbls.show_bubble_size = get_val_bool(e);
        }
        b"dLblPos" if self.in_dlbls => {
            self.dlbls.position =
                get_val_attr(e).and_then(|s| parse_data_label_position(&s));
        }
        b"showLeaderLines" if self.in_dlbls => {
            self.dlbls.show_leader_lines = get_val_bool(e);
        }
        // Data point children
        b"idx" if self.in_dpt => self.dpt_index = get_val_u32(e).unwrap_or(0),
        b"explosion" if self.in_dpt => self.dpt_explosion = get_val_u32(e),
        b"explosion" if self.in_ser && !self.in_dpt => self.ser_explosion = get_val_u32(e),
        // Marker children
        b"symbol" if self.in_marker => {
            self.marker_symbol = get_val_attr(e).and_then(|s| parse_marker_symbol(&s));
        }
        b"size" if self.in_marker => self.marker_size = get_val_u8(e),
        // Trendline children
        b"trendlineType" if self.in_trendline => {
            self.trendline_type = get_val_attr(e).and_then(|s| parse_trendline_type(&s));
        }
        b"order" if self.in_trendline => self.trendline_order = get_val_u32(e),
        b"period" if self.in_trendline => self.trendline_period = get_val_u32(e),
        b"forward" if self.in_trendline => self.trendline_forward = get_val_f64(e),
        b"backward" if self.in_trendline => self.trendline_backward = get_val_f64(e),
        b"intercept" if self.in_trendline => self.trendline_intercept = get_val_f64(e),
        b"dispRSqr" if self.in_trendline => {
            self.trendline_disp_r_sqr = get_val_bool(e);
        }
        b"dispEq" if self.in_trendline => self.trendline_disp_eq = get_val_bool(e),
        // Error bar children
        b"errDir" if self.in_err_bars => {
            self.err_dir = get_val_attr(e).and_then(|s| match s.as_str() {
                "x" => Some(ErrorBarDirection::X),
                "y" => Some(ErrorBarDirection::Y),
                _ => None,
            });
        }
        b"errBarType" if self.in_err_bars => {
            self.err_bar_type = get_val_attr(e).and_then(|s| match s.as_str() {
                "both" => Some(ErrorBarType::Both),
                "minus" => Some(ErrorBarType::Minus),
                "plus" => Some(ErrorBarType::Plus),
                _ => None,
            });
        }
        b"errValType" if self.in_err_bars => {
            self.err_val_type = get_val_attr(e).and_then(|s| match s.as_str() {
                "cust" => Some(ErrorValueType::Custom),
                "fixedVal" => Some(ErrorValueType::FixedValue),
                "percentage" => Some(ErrorValueType::Percentage),
                "stdDev" => Some(ErrorValueType::StandardDeviation),
                "stdErr" => Some(ErrorValueType::StandardError),
                _ => None,
            });
        }
        b"noEndCap" if self.in_err_bars => self.err_no_end_cap = get_val_bool(e),
        // Series smooth
        b"smooth" if self.in_ser => self.ser_smooth = get_val_bool(e),
        b"invertIfNegative" if self.in_ser => {
            self.ser_invert_if_negative = get_val_bool(e);
        }
        // Axis enhancements
        b"majorTickMark" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_major_tick_mark = get_val_attr(e).and_then(|s| parse_tick_mark(&s));
        }
        b"minorTickMark" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_minor_tick_mark = get_val_attr(e).and_then(|s| parse_tick_mark(&s));
        }
        b"tickLblPos" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_label_position =
                get_val_attr(e).and_then(|s| parse_tick_label_position(&s));
        }
        b"delete" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_delete = get_val_bool(e);
        }
        b"crosses" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_crosses = get_val_attr(e).and_then(|s| match s.as_str() {
                "autoZero" => Some(AxisCrosses::AutoZero),
                "min" => Some(AxisCrosses::Min),
                "max" => Some(AxisCrosses::Max),
                _ => None,
            });
        }
        b"crossBetween" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_cross_between = get_val_attr(e).and_then(|s| match s.as_str() {
                "between" => Some(CrossBetween::Between),
                "midCat" => Some(CrossBetween::MidCat),
                _ => None,
            });
        }
        // View 3D children
        b"rotX" if self.in_view_3d => self.view_3d.rotate_x = get_val_i32(e),
        b"rotY" if self.in_view_3d => self.view_3d.rotate_y = get_val_i32(e),
        b"depthPercent" if self.in_view_3d => {
            self.view_3d.depth_percent = get_val_u32(e);
        }
        b"hPercent" if self.in_view_3d => self.view_3d.height_percent = get_val_u32(e),
        b"perspective" if self.in_view_3d => self.view_3d.perspective = get_val_u32(e),
        b"rAngAx" if self.in_view_3d => {
            self.view_3d.right_angle_axes = get_val_bool(e);
        }
        // Chart-level config
        b"plotVisOnly" if self.in_chart && !self.in_plot_area => {
            self.result.plot_visible_only = get_val_bool(e);
        }
        b"autoTitleDeleted" if self.in_chart && !self.in_plot_area => {
            self.result.auto_title_deleted = get_val_bool(e);
        }
        b"showDLblsOverMax" if self.in_chart && !self.in_plot_area => {
            self.result.show_dlbls_over_max = get_val_bool(e);
        }
        b"roundedCorners" if self.in_chart_space && !self.in_chart => {
            self.result.rounded_corners = get_val_bool(e);
        }
        b"dispBlanksAs" if self.in_chart && !self.in_plot_area => {
            self.result.display_blanks_as =
                get_val_attr(e).and_then(|s| match s.as_str() {
                    "gap" => Some(DisplayBlanksAs::Gap),
                    "span" => Some(DisplayBlanksAs::Span),
                    "zero" => Some(DisplayBlanksAs::Zero),
                    _ => None,
                });
        }
        // Manual layout children
        b"x" if self.in_manual_layout => self.manual_layout.x = get_val_f64(e),
        b"y" if self.in_manual_layout => self.manual_layout.y = get_val_f64(e),
        b"w" if self.in_manual_layout => self.manual_layout.width = get_val_f64(e),
        b"h" if self.in_manual_layout => self.manual_layout.height = get_val_f64(e),
        // Data table children
        b"showHorzBorder" if self.in_d_table => {
            self.d_table.show_horizontal_border = get_val_bool(e);
        }
        b"showVertBorder" if self.in_d_table => {
            self.d_table.show_vertical_border = get_val_bool(e);
        }
        b"showOutline" if self.in_d_table => {
            self.d_table.show_outline = get_val_bool(e);
        }
        b"showKeys" if self.in_d_table => self.d_table.show_keys = get_val_bool(e),
        b"srgbClr" if self.in_sp_pr && !self.in_sp_ln => {
            if let Some(hex) = get_val_attr(e) {
                self.sp_solid_fill = Some(ChartColor { hex });
            }
        }
        b"srgbClr" if self.in_sp_ln => {
            if let Some(hex) = get_val_attr(e) {
                self.sp_ln_solid_fill = Some(ChartColor { hex });
            }
        }
        b"noFill" if self.in_sp_pr && !self.in_sp_ln => self.sp_no_fill = true,
        b"noFill" if self.in_sp_ln => self.sp_ln_no_fill = true,
        b"prstDash" if self.in_sp_ln => self.sp_ln_dash = get_val_attr(e),
        b"axPos" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_position = get_val_attr(e).and_then(|s| match s.as_str() {
                "b" => Some(AxisPosition::Bottom),
                "t" => Some(AxisPosition::Top),
                "l" => Some(AxisPosition::Left),
                "r" => Some(AxisPosition::Right),
                _ => None,
            });
        }
        b"majorUnit" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_major_unit = get_val_f64(e);
        }
        b"minorUnit" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_minor_unit = get_val_f64(e);
        }
        b"overlay" if self.in_legend => self.legend_overlay = get_val_bool(e),
        b"varyColors" if self.in_chart_type_element && !self.in_ser => {
            self.vary_colors = get_val_bool(e);
        }
        b"gapWidth" if self.in_up_down_bars => {
            self.up_down_bars_gap_width = get_val_u32(e);
        }
        b"gapWidth" if self.in_chart_type_element && !self.in_ser => {
            self.gap_width = get_val_u32(e);
        }
        b"overlap" if self.in_chart_type_element && !self.in_ser => {
            self.overlap = get_val_i32(e).map(|v| v as i32);
        }
        b"axId" if self.in_chart_type_element && !self.in_ser => {
            if let Some(id) = get_val_u32(e) {
                self.group_axis_ids.push(id);
            }
        }
        b"axId" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_id = get_val_u32(e);
        }
        b"crossAx" if (self.in_cat_ax || self.in_val_ax || self.in_ser_ax) && !self.in_ax_title => {
            self.ax_cross_id = get_val_u32(e);
        }
        _ => {}
        }

        if deepen_title {
            self.title_depth += 1;
        }
        if deepen_sp_pr {
            self.sp_pr_depth += 1;
        }
    }

    fn on_text(&mut self, e: &BytesText) {
    if let Ok(text) = e.unescape() {
        let text_str = text.as_ref();
        if self.in_title_t {
            self.title_text.push_str(text_str);
        } else if self.in_title_str_ref_f {
            self.title_text.push_str(text_str);
        } else if self.in_ser_tx_str_ref_f {
            push_text(&mut self.ser_name, text_str);
        } else if self.in_ser_tx_v {
            push_text(&mut self.ser_name, text_str);
        } else if self.in_ser_val_num_ref_f {
            push_text(&mut self.ser_val_formula, text_str);
        } else if self.in_ser_val_pt_v {
            if let Ok(v) = text_str.parse::<f64>() {
                self.ser_val_cache.push(v);
            }
        } else if self.in_ser_cat_ref_f {
            push_text(&mut self.ser_cat_formula, text_str);
        } else if self.in_ax_title_t {
            self.ax_title_text.push_str(text_str);
        } else if self.in_trendline_name {
            push_text(&mut self.trendline_name, text_str);
        } else if self.in_dlbls_separator {
            push_text(&mut self.dlbls.separator, text_str);
        }
    }
    }

    /// Close the chart title, keeping whatever text it accumulated.
    fn finish_title(&mut self) {
        if !self.title_text.is_empty() {
            self.result.title = Some(self.title_text.clone());
        }
        self.in_chart_title = false;
    }

    /// Close a cx:spPr, attaching what it described to whatever it
    /// belongs to.
    fn finish_shape_properties(&mut self) {
            let props = ChartShapeProperties {
                solid_fill: self.sp_solid_fill.take(),
                no_fill: self.sp_no_fill,
                line: self.sp_line.take(),
            };
            let has_content =
                props.solid_fill.is_some() || props.no_fill || props.line.is_some();
            if has_content {
                match self.sp_pr_context {
                    SpPrContext::Series => {
                        self.ser_shape_properties = Some(props);
                    }
                    SpPrContext::DataPoint => {
                        self.dpt_shape_properties = Some(props);
                    }
                    SpPrContext::CatAxis => {
                        self.ax_shape_properties = Some(props);
                    }
                    SpPrContext::ValAxis => {
                        self.ax_shape_properties = Some(props);
                    }
                    SpPrContext::MajorGridlines => {
                        self.ax_major_gridlines = true;
                        self.ax_major_gridlines_shape_properties = Some(props);
                    }
                    SpPrContext::MinorGridlines => {
                        self.ax_minor_gridlines = true;
                        self.ax_minor_gridlines_shape_properties = Some(props);
                    }
                    SpPrContext::ChartSpace => {
                        self.result.shape_properties = Some(props);
                    }
                    SpPrContext::Legend => {
                        self.legend_shape_properties = Some(props);
                    }
                    SpPrContext::None => {}
                    SpPrContext::DropLines => {
                        self.group_drop_lines = Some(ChartLines {
                            shape_properties: Some(props),
                        });
                    }
                    SpPrContext::HiLowLines => {
                        self.group_high_low_lines = Some(ChartLines {
                            shape_properties: Some(props),
                        });
                    }
                    SpPrContext::SerLines => {
                        self.group_series_lines = Some(ChartLines {
                            shape_properties: Some(props),
                        });
                    }
                    SpPrContext::UpBars => self.up_bars_sp = Some(props),
                    SpPrContext::DownBars => self.down_bars_sp = Some(props),
                    SpPrContext::LeaderLines => {
                        self.dlbls.leader_lines = Some(ChartLines {
                            shape_properties: Some(props),
                        });
                    }
                }
            }
            self.in_sp_pr = false;
            self.sp_no_fill = false;
            self.sp_pr_context = SpPrContext::None;
    }

    fn on_end(&mut self, e: &BytesEnd) {
        let local = e.name().local_name();
        let tag = local.as_ref();

        // The mirror of on_start's bookkeeping: one unwind per element
        // end, wherever the element is handled, so that an arm existing
        // for a tag cannot leave a region open.
        if self.in_chart_title {
            self.title_depth = self.title_depth.saturating_sub(1);
            if self.title_depth == 0 {
                self.finish_title();
            }
        }
        if self.in_sp_pr {
            self.sp_pr_depth = self.sp_pr_depth.saturating_sub(1);
            if self.sp_pr_depth == 0 {
                self.finish_shape_properties();
            }
        }

        match tag {
        b"chart" => self.in_chart = false,
        b"plotArea" => self.in_plot_area = false,
        b"view3D" if self.in_view_3d => {
            self.result.view_3d = Some(self.view_3d.clone());
            self.in_view_3d = false;
        }
        b"layout" if self.in_layout => {
            if self.had_manual_layout {
                self.result.layout = Some(Layout {
                    manual_layout: Some(self.manual_layout.clone()),
                });
            }
            self.had_manual_layout = false;
            self.manual_layout = ManualLayout::default();
            self.in_layout = false;
        }
        b"manualLayout" if self.in_manual_layout => self.in_manual_layout = false,
        b"dTable" if self.in_d_table => {
            self.result.data_table = Some(self.d_table.clone());
            self.in_d_table = false;
        }
        b"title" if self.in_ax_title => self.in_ax_title = false,
        b"tx" if self.in_title_tx => self.in_title_tx = false,
        b"rich" if self.in_title_rich => self.in_title_rich = false,
        b"p" if self.in_title_p && self.in_title_rich => self.in_title_p = false,
        b"r" if self.in_title_r && self.in_title_p => self.in_title_r = false,
        b"t" if self.in_title_t => self.in_title_t = false,
        b"strRef" if self.in_title_str_ref => self.in_title_str_ref = false,
        b"f" if self.in_title_str_ref_f => self.in_title_str_ref_f = false,
        b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
        | b"pie3DChart" | b"doughnutChart" | b"areaChart" | b"area3DChart"
        | b"scatterChart" | b"bubbleChart" | b"radarChart" | b"stockChart"
        | b"surfaceChart" | b"surface3DChart" | b"ofPieChart"
            if self.in_chart_type_element =>
        {
            let ct = resolve_chart_type(
                self.chart_type_tag.as_deref(),
                self.bar_dir.as_deref(),
                self.grouping.as_deref(),
                self.scatter_style.as_deref(),
            );
            let is_3d = self.chart_type_tag
                .as_deref()
                .map_or(false, |t| t.contains("3D"));
            let group = ChartTypeGroup {
                chart_type: ct,
                is_3d,
                series: std::mem::take(&mut self.group_series),
                data_labels: self.group_data_labels.take(),
                vary_colors: self.vary_colors.take(),
                gap_width: self.gap_width.take(),
                overlap: self.overlap.take(),
                first_slice_angle: self.first_slice_angle.take(),
                hole_size: self.hole_size.take(),
                bubble_scale: self.bubble_scale.take(),
                show_negative_bubbles: self.show_negative_bubbles.take(),
                radar_style: self.radar_style.take(),
                wireframe: self.wireframe.take(),
                drop_lines: self.group_drop_lines.take(),
                high_low_lines: self.group_high_low_lines.take(),
                series_lines: self.group_series_lines.take(),
                up_down_bars: self.group_up_down_bars.take(),
                axis_ids: std::mem::take(&mut self.group_axis_ids),
                raw_ext: self.group_raw_ext.take(),
                of_pie_type: None,
                split_type: None,
                split_pos: None,
                second_pie_size: None,
                bar_shape: None,
                floor: None,
                side_wall: None,
                back_wall: None,
            };
            self.result.type_groups.push(group);
            self.in_chart_type_element = false;
        }
        b"dropLines" if self.in_drop_lines => {
            if self.group_drop_lines.is_none() {
                self.group_drop_lines = Some(ChartLines::default());
            }
            self.in_drop_lines = false;
        }
        b"hiLowLines" if self.in_hi_low_lines => {
            if self.group_high_low_lines.is_none() {
                self.group_high_low_lines = Some(ChartLines::default());
            }
            self.in_hi_low_lines = false;
        }
        b"serLines" if self.in_ser_lines => {
            if self.group_series_lines.is_none() {
                self.group_series_lines = Some(ChartLines::default());
            }
            self.in_ser_lines = false;
        }
        b"upBars" if self.in_up_bars => self.in_up_bars = false,
        b"downBars" if self.in_down_bars => self.in_down_bars = false,
        b"upDownBars" if self.in_up_down_bars => {
            let up = self.up_bars_sp
                .take()
                .map(|sp| ChartLines {
                    shape_properties: Some(sp),
                })
                .or(if self.had_up_bars {
                    Some(ChartLines::default())
                } else {
                    None
                });
            let down = self.down_bars_sp
                .take()
                .map(|sp| ChartLines {
                    shape_properties: Some(sp),
                })
                .or(if self.had_down_bars {
                    Some(ChartLines::default())
                } else {
                    None
                });
            self.group_up_down_bars = Some(UpDownBars {
                gap_width: self.up_down_bars_gap_width.take(),
                up_bars: up,
                down_bars: down,
            });
            self.had_up_bars = false;
            self.had_down_bars = false;
            self.in_up_down_bars = false;
        }
        b"leaderLines" if self.in_leader_lines => {
            if self.dlbls.leader_lines.is_none() {
                self.dlbls.leader_lines = Some(ChartLines::default());
            }
            self.in_leader_lines = false;
        }
        // Data labels
        b"separator" if self.in_dlbls_separator => self.in_dlbls_separator = false,
        b"dLbls" if self.in_dlbls => {
            if self.in_ser {
                self.ser_data_labels = Some(self.dlbls.clone());
            } else {
                self.group_data_labels = Some(self.dlbls.clone());
            }
            self.in_dlbls = false;
        }
        // Marker
        b"marker" if self.in_marker => {
            let m = Marker {
                symbol: self.marker_symbol.take(),
                size: self.marker_size.take(),
            };
            if self.in_dpt {
                self.dpt_marker = Some(m);
            } else {
                self.ser_marker = Some(m);
            }
            self.in_marker = false;
        }
        // Data point
        b"dPt" if self.in_dpt => {
            self.ser_data_points.push(DataPoint {
                index: self.dpt_index,
                marker: self.dpt_marker.take(),
                explosion: self.dpt_explosion.take(),
                shape_properties: self.dpt_shape_properties.take(),
            });
            self.in_dpt = false;
        }
        // Trendline
        b"name" if self.in_trendline_name => self.in_trendline_name = false,
        b"trendline" if self.in_trendline => {
            if let Some(tt) = self.trendline_type.take() {
                self.ser_trendline = Some(Trendline {
                    trendline_type: tt,
                    name: self.trendline_name.take(),
                    order: self.trendline_order.take(),
                    period: self.trendline_period.take(),
                    forward: self.trendline_forward.take(),
                    backward: self.trendline_backward.take(),
                    intercept: self.trendline_intercept.take(),
                    label: None,
                    display_r_squared: self.trendline_disp_r_sqr.take(),
                    display_equation: self.trendline_disp_eq.take(),
                });
            }
            self.in_trendline = false;
        }
        // Error bars
        b"errBars" if self.in_err_bars => {
            self.ser_error_bars = Some(ErrorBars {
                direction: self.err_dir.unwrap_or(ErrorBarDirection::Y),
                bar_type: self.err_bar_type.unwrap_or(ErrorBarType::Both),
                value_type: self.err_val_type.unwrap_or(ErrorValueType::FixedValue),
                value: self.err_val.take(),
                no_end_cap: self.err_no_end_cap.take(),
                plus: None,
                minus: None,
            });
            self.in_err_bars = false;
        }
        b"ser" if self.in_ser => {
            let values = if let Some(ref f) = self.ser_val_formula {
                DataReference::formula(f)
            } else if !self.ser_val_cache.is_empty() {
                DataReference::numbers(self.ser_val_cache.clone())
            } else {
                DataReference::numbers(Vec::new())
            };

            let mut ds = DataSeries::new(values);
            if let Some(ref name) = self.ser_name {
                ds = ds.with_name(name);
            }
            if let Some(ref f) = self.ser_cat_formula {
                ds = ds.with_categories(DataReference::formula(f));
            }
            ds.data_labels = self.ser_data_labels.take();
            ds.trendline = self.ser_trendline.take();
            ds.error_bars = self.ser_error_bars.take();
            ds.marker = self.ser_marker.take();
            ds.data_points = std::mem::take(&mut self.ser_data_points);
            ds.smooth = self.ser_smooth.take();
            ds.explosion = self.ser_explosion.take();
            ds.shape_properties = self.ser_shape_properties.take();
            ds.raw_ext = self.ser_raw_ext.take();
            ds.invert_if_negative = self.ser_invert_if_negative.take();
            self.group_series.push(ds);

            self.in_ser = false;
            self.ser_name = None;
            self.ser_val_formula = None;
            self.ser_val_cache.clear();
            self.ser_cat_formula = None;
        }
        b"tx" if self.in_ser_tx => self.in_ser_tx = false,
        b"strRef" if self.in_ser_tx_str_ref => self.in_ser_tx_str_ref = false,
        b"f" if self.in_ser_tx_str_ref_f => self.in_ser_tx_str_ref_f = false,
        b"v" if self.in_ser_tx_v => self.in_ser_tx_v = false,
        b"val" if self.in_ser_val => {
            self.in_ser_val = false;
            self.in_ser_val_num_ref = false;
        }
        b"yVal" if self.in_ser_yval => {
            self.in_ser_yval = false;
            self.in_ser_val_num_ref = false;
        }
        b"numRef" if self.in_ser_val_num_ref => self.in_ser_val_num_ref = false,
        b"f" if self.in_ser_val_num_ref_f => self.in_ser_val_num_ref_f = false,
        b"numCache" if self.in_ser_val_num_cache => self.in_ser_val_num_cache = false,
        b"pt" if self.in_ser_val_pt => self.in_ser_val_pt = false,
        b"v" if self.in_ser_val_pt_v => self.in_ser_val_pt_v = false,
        b"cat" if self.in_ser_cat => {
            self.in_ser_cat = false;
            self.in_ser_cat_ref = false;
        }
        b"xVal" if self.in_ser_xval => {
            self.in_ser_xval = false;
            self.in_ser_cat_ref = false;
        }
        b"strRef" | b"numRef" if self.in_ser_cat_ref => self.in_ser_cat_ref = false,
        b"f" if self.in_ser_cat_ref_f => self.in_ser_cat_ref_f = false,
        b"catAx" | b"dateAx" if self.in_cat_ax => {
            let mut axis = Axis::new();
            if !self.ax_title_text.is_empty() {
                axis = axis.with_title(&self.ax_title_text);
            }
            if let (Some(min), Some(max)) = (self.ax_min, self.ax_max) {
                axis = axis.with_bounds(min, max);
            } else {
                axis.minimum = self.ax_min;
                axis.maximum = self.ax_max;
            }
            axis.number_format = self.ax_number_format.take();
            axis.major_gridlines = self.ax_major_gridlines;
            axis.minor_gridlines = self.ax_minor_gridlines;
            axis.major_gridlines_shape_properties =
                self.ax_major_gridlines_shape_properties.take();
            axis.minor_gridlines_shape_properties =
                self.ax_minor_gridlines_shape_properties.take();
            axis.major_tick_mark = self.ax_major_tick_mark.take();
            axis.minor_tick_mark = self.ax_minor_tick_mark.take();
            axis.label_position = self.ax_label_position.take();
            // An omitted c:delete is not "unspecified": Excel treats the
            // axis as deleted. Recording that keeps a rewrite from
            // silently making a hidden axis visible again.
            axis.delete = Some(self.ax_delete.take().unwrap_or(true));
            axis.crosses = self.ax_crosses.take();
            axis.cross_between = self.ax_cross_between.take();
            axis.position = self.ax_position.take();
            axis.major_unit = self.ax_major_unit.take();
            axis.minor_unit = self.ax_minor_unit.take();
            axis.shape_properties = self.ax_shape_properties.take();
            axis.raw_ext = self.ax_raw_ext.take();
            if self.is_date_ax {
                axis.axis_type = AxisType::Date;
            }
            self.result.category_axis = Some(axis.clone());
            if let Some(id) = self.ax_id.take() {
                self.result.axes.push(ChartAxis {
                    id,
                    cross_id: self.ax_cross_id.take().unwrap_or(0),
                    axis: axis,
                });
            }
            self.in_cat_ax = false;
            self.ax_title_text.clear();
        }
        b"valAx" if self.in_val_ax => {
            let mut axis = Axis::new();
            if !self.ax_title_text.is_empty() {
                axis = axis.with_title(&self.ax_title_text);
            }
            if let (Some(min), Some(max)) = (self.ax_min, self.ax_max) {
                axis = axis.with_bounds(min, max);
            } else {
                axis.minimum = self.ax_min;
                axis.maximum = self.ax_max;
            }
            axis.number_format = self.ax_number_format.take();
            axis.major_gridlines = self.ax_major_gridlines;
            axis.minor_gridlines = self.ax_minor_gridlines;
            axis.major_gridlines_shape_properties =
                self.ax_major_gridlines_shape_properties.take();
            axis.minor_gridlines_shape_properties =
                self.ax_minor_gridlines_shape_properties.take();
            axis.major_tick_mark = self.ax_major_tick_mark.take();
            axis.minor_tick_mark = self.ax_minor_tick_mark.take();
            axis.label_position = self.ax_label_position.take();
            // An omitted c:delete is not "unspecified": Excel treats the
            // axis as deleted. Recording that keeps a rewrite from
            // silently making a hidden axis visible again.
            axis.delete = Some(self.ax_delete.take().unwrap_or(true));
            axis.crosses = self.ax_crosses.take();
            axis.cross_between = self.ax_cross_between.take();
            axis.position = self.ax_position.take();
            axis.major_unit = self.ax_major_unit.take();
            axis.minor_unit = self.ax_minor_unit.take();
            axis.shape_properties = self.ax_shape_properties.take();
            axis.raw_ext = self.ax_raw_ext.take();
            axis.axis_type = AxisType::Value;
            self.result.value_axis = Some(axis.clone());
            if let Some(id) = self.ax_id.take() {
                self.result.axes.push(ChartAxis {
                    id,
                    cross_id: self.ax_cross_id.take().unwrap_or(0),
                    axis: axis,
                });
            }
            self.in_val_ax = false;
            self.ax_title_text.clear();
        }
        b"serAx" if self.in_ser_ax => {
            let mut axis = Axis::new();
            axis.axis_type = AxisType::Series;
            if !self.ax_title_text.is_empty() {
                axis = axis.with_title(&self.ax_title_text);
            }
            if let (Some(min), Some(max)) = (self.ax_min, self.ax_max) {
                axis = axis.with_bounds(min, max);
            } else {
                axis.minimum = self.ax_min;
                axis.maximum = self.ax_max;
            }
            axis.number_format = self.ax_number_format.take();
            axis.major_gridlines = self.ax_major_gridlines;
            axis.minor_gridlines = self.ax_minor_gridlines;
            axis.major_gridlines_shape_properties =
                self.ax_major_gridlines_shape_properties.take();
            axis.minor_gridlines_shape_properties =
                self.ax_minor_gridlines_shape_properties.take();
            axis.major_tick_mark = self.ax_major_tick_mark.take();
            axis.minor_tick_mark = self.ax_minor_tick_mark.take();
            axis.label_position = self.ax_label_position.take();
            // An omitted c:delete is not "unspecified": Excel treats the
            // axis as deleted. Recording that keeps a rewrite from
            // silently making a hidden axis visible again.
            axis.delete = Some(self.ax_delete.take().unwrap_or(true));
            axis.crosses = self.ax_crosses.take();
            axis.cross_between = self.ax_cross_between.take();
            axis.position = self.ax_position.take();
            axis.major_unit = self.ax_major_unit.take();
            axis.minor_unit = self.ax_minor_unit.take();
            axis.shape_properties = self.ax_shape_properties.take();
            axis.raw_ext = self.ax_raw_ext.take();
            self.result.series_axis = Some(axis.clone());
            if let Some(id) = self.ax_id.take() {
                self.result.axes.push(ChartAxis {
                    id,
                    cross_id: self.ax_cross_id.take().unwrap_or(0),
                    axis: axis,
                });
            }
            self.in_ser_ax = false;
            self.ax_title_text.clear();
        }
        b"tx" if self.in_ax_title_tx => self.in_ax_title_tx = false,
        b"rich" if self.in_ax_title_rich => self.in_ax_title_rich = false,
        b"p" if self.in_ax_title_p && self.in_ax_title_rich => self.in_ax_title_p = false,
        b"r" if self.in_ax_title_r => self.in_ax_title_r = false,
        b"t" if self.in_ax_title_t => self.in_ax_title_t = false,
        b"scaling" if self.in_ax_scaling => self.in_ax_scaling = false,
        b"majorGridlines" if self.in_ax_major_gridlines => self.in_ax_major_gridlines = false,
        b"minorGridlines" if self.in_ax_minor_gridlines => self.in_ax_minor_gridlines = false,
        b"legend" if self.in_legend => {
            let mut leg = Legend::new(self.legend_pos.unwrap_or(LegendPosition::Right));
            if let Some(true) = self.legend_overlay {
                leg.overlay = true;
            }
            leg.shape_properties = self.legend_shape_properties.take();
            self.result.legend = Some(leg);
            self.in_legend = false;
        }
        b"ln" if self.in_sp_ln => {
            self.sp_line = Some(ChartLine {
                width: self.sp_ln_width.take(),
                solid_fill: self.sp_ln_solid_fill.take(),
                no_fill: self.sp_ln_no_fill,
                dash_style: self.sp_ln_dash.take(),
            });
            self.in_sp_ln = false;
            self.sp_ln_no_fill = false;
        }
        _ => {}
        }
    }

}

fn parse_chart_xml_inner<R: Read>(
    xml_reader: &mut Reader<BufReader<R>>,
    buf: &mut Vec<u8>,
) -> ChartParseResult<ParsedChart> {
    let mut parser = ChartParser::new();

    // `<x/>` and `<x></x>` are the same document, so a self-closing
    // element is split into the start and end its expanded form would
    // produce and handled by one code path; keeping two let them drift.
    // An extLst being kept as bytes is exempt, so that what it replays
    // matches its source.
    let mut pending_end: Option<Event<'static>> = None;

    loop {
        let event = match pending_end.take() {
            Some(ev) => ev,
            None => {
                let ev = match xml_reader.read_event_into(buf) {
                    Ok(ev) => ev.into_owned(),
                    Err(e) => return Err(ChartParseError::Xml(e)),
                };
                buf.clear();
                ev
            }
        };
        if parser.capturing() {
            parser.capture(&event);
            continue;
        }
        let event = match event {
            Event::Empty(e) => {
                pending_end = Some(Event::End(BytesEnd::new(
                    String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                )));
                Event::Start(e)
            }
            other => other,
        };
        match event {
            Event::Start(ref e) => parser.on_start(e),
            Event::Text(ref e) => parser.on_text(e),
            Event::End(ref e) => parser.on_end(e),
            Event::Eof => break,
            _ => {}
        }
    }

    // Populate legacy fields from type_groups for backward compatibility
    if let Some(first) = parser.result.type_groups.first() {
        parser.result.chart_type = first.chart_type.clone();
        parser.result.is_3d = first.is_3d;
        parser.result.series = first.series.clone();
        parser.result.data_labels = first.data_labels.clone();
        parser.result.vary_colors = first.vary_colors;
        parser.result.gap_width = first.gap_width;
        parser.result.overlap = first.overlap;
        parser.result.first_slice_angle = first.first_slice_angle;
        parser.result.hole_size = first.hole_size;
        parser.result.bubble_scale = first.bubble_scale;
        parser.result.show_negative_bubbles = first.show_negative_bubbles;
        parser.result.radar_style = first.radar_style.clone();
        parser.result.wireframe = first.wireframe;
        parser.result.drop_lines = first.drop_lines.clone();
        parser.result.high_low_lines = first.high_low_lines.clone();
        parser.result.up_down_bars = first.up_down_bars.clone();
        parser.result.series_lines = first.series_lines.clone();
    }

    // Detect PieExploded: shares the same XML element (pieChart) as Pie
    // but has explosion attributes on series.
    if parser.result.chart_type == ChartType::Pie {
        let has_explosion = parser.result.series.iter().any(|s| s.explosion.is_some());
        if has_explosion {
            parser.result.chart_type = ChartType::PieExploded;
        }
    }

    // Re-populate legacy axis fields from result.axes using first group's axis_ids.
    // Only needed for combo charts (2+ groups) where multiple value axes exist
    // and the parse-loop's last-wins behavior gives the wrong legacy value_axis.
    if parser.result.type_groups.len() >= 2 {
        if let Some(first) = parser.result.type_groups.first() {
            let axis_ids = &first.axis_ids;
            parser.result.category_axis = None;
            parser.result.value_axis = None;
            for ax in &parser.result.axes {
                if axis_ids.contains(&ax.id) {
                    match ax.axis.axis_type {
                        AxisType::Category | AxisType::Date => {
                            if parser.result.category_axis.is_none() {
                                parser.result.category_axis = Some(ax.axis.clone());
                            }
                        }
                        AxisType::Value => {
                            if parser.result.value_axis.is_none() {
                                parser.result.value_axis = Some(ax.axis.clone());
                            }
                        }
                        AxisType::Series => {
                            if parser.result.series_axis.is_none() {
                                parser.result.series_axis = Some(ax.axis.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    // Only keep type_groups/axes in combo mode (2+ groups)
    if parser.result.type_groups.len() < 2 {
        if let Some(first) = parser.result.type_groups.first_mut() {
            if let Some(raw) = first.raw_ext.take() {
                parser.result.raw_extensions.insert("chartType".into(), raw);
            }
        }
        parser.result.type_groups.clear();
        parser.result.axes.clear();
    }

    Ok(parser.result)
}

/// Where a captured `c:extLst` belongs, decided when it opens.
#[derive(Debug, Clone, Copy)]
enum ExtLstDest {
    Series,
    Axis,
    TypeGroup,
    Chart,
    PlotArea,
    ChartSpace,
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
        assert_eq!(cat.position, Some(AxisPosition::Bottom));
        let val = parsed.value_axis.as_ref().unwrap();
        assert_eq!(val.position, Some(AxisPosition::Left));
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
