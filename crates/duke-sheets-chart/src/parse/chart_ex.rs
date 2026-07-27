use std::io::{BufReader, Cursor, Read};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::chart_ex::*;
use crate::error::{ChartParseError, ChartParseResult};
use crate::{ChartColor, ChartLine, ChartShapeProperties, NumberFormat};

/// Parse chart-ex XML from a reader and return a `ChartEx`.
///
/// Placement comes from the wrapping drawing object's anchor, not
/// the chart XML.
pub fn parse_chart_ex_xml<R: Read>(reader: R) -> ChartParseResult<ChartEx> {
    let buf_reader = BufReader::new(reader);
    let mut xml_reader = Reader::from_reader(buf_reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let parsed = parse_chart_ex_xml_inner(&mut xml_reader, &mut buf)?;

    Ok(ChartEx {
        title: parsed.title,
        data: parsed.data,
        plot_area: parsed.plot_area,
        legend: parsed.legend,
        shape_properties: parsed.shape_properties,
        text_properties: None,
        color_map_override: None,
        format_overrides: Vec::new(),
        print_settings: parsed.print_settings,
        raw_chart_style: None,
        raw_chart_color_style: None,
        raw_extensions: parsed.raw_extensions,
        raw_mc_fallback: None,
        version: parsed.version,
        feature_list: parsed.feature_list,
        fallback_img: parsed.fallback_img,
        external_data: parsed.external_data,
        extensions: None,
    })
}

struct ParsedChartEx {
    title: Option<ChartExTitle>,
    data: Vec<ChartExData>,
    plot_area: ChartExPlotArea,
    legend: Option<ChartExLegend>,
    shape_properties: Option<ChartShapeProperties>,
    print_settings: Option<ChartExPrintSettings>,
    raw_extensions: std::collections::HashMap<String, Vec<u8>>,
    external_data: Option<ChartExExternalData>,
    version: Option<String>,
    feature_list: Option<String>,
    fallback_img: Option<String>,
}

    #[derive(Clone, Copy, PartialEq, Eq)]
    enum SpCtx {
        None,
        ChartSpace,
        Series,
        DataPt,
        Legend,
        Axis,
        AxisTitle,
        Title,
        MajorGrid,
        MinorGrid,
        DataLabels,
        DataLabel,
        PlotSurface,
    }

/// Streaming state for one chartEx part.
///
/// The document is deep and its elements are context sensitive, so
/// parsing is a state machine rather than a recursive descent: the
/// fields below are what a start tag, a text node or an end tag needs
/// to know about where it sits. Holding them on one value keeps the
/// three event handlers to a single definition each, which is what
/// stops the handling of an element drifting between them.
struct ChartExParser {
    result: ParsedChartEx,
    in_chart_space: bool,
    in_chart_data: bool,
    in_chart: bool,
    in_plot_area: bool,
    in_plot_area_region: bool,
    in_data: bool,
    data_id: u32,
    data_dims: Vec<ChartExDimension>,
    in_str_dim: bool,
    in_num_dim: bool,
    dim_str_type: StringDimType,
    dim_num_type: NumericDimType,
    dim_formula: Option<String>,
    dim_nf: Option<String>,
    in_dim_f: bool,
    in_dim_nf: bool,
    dim_levels: Vec<ChartExStringLevel>,
    dim_num_levels: Vec<ChartExNumericLevel>,
    in_lvl: bool,
    lvl_pt_count: u32,
    lvl_name: Option<String>,
    lvl_format_code: Option<String>,
    lvl_str_points: Vec<(u32, String)>,
    lvl_num_points: Vec<(u32, String)>,
    in_lvl_pt: bool,
    lvl_pt_idx: u32,
    in_lvl_pt_text: bool,
    in_title: bool,
    title_pos: Option<String>,
    title_align: Option<String>,
    title_overlay: Option<bool>,
    in_title_tx: bool,
    in_title_tx_data: bool,
    in_title_tx_data_v: bool,
    in_title_tx_data_f: bool,
    title_text: Option<String>,
    title_sp: Option<ChartShapeProperties>,
    in_series: bool,
    ser_layout: ChartExLayout,
    ser_unique_id: Option<String>,
    ser_hidden: Option<bool>,
    ser_owner_idx: Option<u32>,
    ser_format_idx: Option<u32>,
    ser_text: Option<ChartExText>,
    ser_data_id: u32,
    ser_data_labels: Option<ChartExDataLabels>,
    ser_data_points: Vec<ChartExDataPoint>,
    ser_layout_pr: Option<ChartExLayoutPr>,
    ser_axis_ids: Vec<u32>,
    ser_value_colors: Option<ChartExValueColors>,
    ser_value_color_positions: Option<ChartExValueColorPositions>,
    ser_shape_properties: Option<ChartShapeProperties>,
    in_ser_tx: bool,
    in_ser_tx_data: bool,
    in_ser_tx_data_v: bool,
    in_ser_tx_data_f: bool,
    ser_tx_value: Option<String>,
    ser_tx_formula: Option<String>,
    in_data_id: bool,
    in_data_labels: bool,
    dlbl_pos: Option<String>,
    dlbl_vis_series: Option<bool>,
    dlbl_vis_cat: Option<bool>,
    dlbl_vis_val: Option<bool>,
    dlbl_num_fmt: Option<NumberFormat>,
    dlbl_separator: Option<String>,
    dlbl_sp: Option<ChartShapeProperties>,
    in_dlbl_separator: bool,
    dl_overrides: Vec<ChartExDataLabel>,
    dl_hidden: Vec<u32>,
    in_data_label: bool,
    lbl_idx: u32,
    lbl_pos: Option<String>,
    lbl_vis_series: Option<bool>,
    lbl_vis_cat: Option<bool>,
    lbl_vis_val: Option<bool>,
    lbl_num_fmt: Option<NumberFormat>,
    lbl_separator: Option<String>,
    lbl_sp: Option<ChartShapeProperties>,
    in_lbl_separator: bool,
    in_data_pt: bool,
    dpt_idx: u32,
    dpt_sp: Option<ChartShapeProperties>,
    in_layout_pr: bool,
    layout_pr: ChartExLayoutPr,
    in_subtotals: bool,
    in_binning: bool,
    in_geography: bool,
    geo_cache_writer: Option<Writer<Cursor<Vec<u8>>>>,
    geo_cache_depth: u32,
    in_statistics: bool,
    in_layout_visibility: bool,
    in_axis_id: bool,
    in_value_colors: bool,
    vc_capturing_tag: Option<String>,
    vc_writer: Option<Writer<Cursor<Vec<u8>>>>,
    vc_depth: u32,
    in_value_color_positions: bool,
    vcp: ChartExValueColorPositions,
    in_vcp_min: bool,
    in_vcp_mid: bool,
    in_vcp_max: bool,
    in_axis: bool,
    ax_id: u32,
    ax_hidden: Option<bool>,
    ax_scaling: ChartExScaling,
    ax_title: Option<ChartExAxisTitle>,
    ax_units: Option<ChartExAxisUnits>,
    ax_major_gridlines: Option<ChartExGridlines>,
    ax_minor_gridlines: Option<ChartExGridlines>,
    ax_major_tick_marks: Option<String>,
    ax_minor_tick_marks: Option<String>,
    ax_tick_labels: bool,
    ax_num_fmt: Option<NumberFormat>,
    ax_shape_properties: Option<ChartShapeProperties>,
    in_cat_scaling: bool,
    in_val_scaling: bool,
    in_ax_title: bool,
    in_ax_title_tx: bool,
    in_ax_title_tx_data: bool,
    in_ax_title_tx_data_v: bool,
    ax_title_text: Option<String>,
    ax_title_sp: Option<ChartShapeProperties>,
    in_units: bool,
    units_unit: Option<String>,
    in_major_gridlines: bool,
    in_minor_gridlines: bool,
    in_legend: bool,
    legend_pos: Option<String>,
    legend_align: Option<String>,
    legend_overlay: Option<bool>,
    legend_sp: Option<ChartShapeProperties>,
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
    sp_ctx: SpCtx,
    in_print_settings: bool,
    print_settings: ChartExPrintSettings,
    in_ps_header_footer: bool,
    ps_hf: ChartExHeaderFooter,
    in_ps_hf_odd_header: bool,
    in_ps_hf_odd_footer: bool,
    in_ps_hf_even_header: bool,
    in_ps_hf_even_footer: bool,
    in_ps_hf_first_header: bool,
    in_ps_hf_first_footer: bool,
    skip_depth: u32,
    skipping: bool,
    in_plot_surface: bool,
}

impl ChartExParser {
    fn new() -> Self {
        Self {
            result: ParsedChartEx {
                title: None,
                data: Vec::new(),
                plot_area: ChartExPlotArea::default(),
                legend: None,
                shape_properties: None,
                print_settings: None,
                raw_extensions: std::collections::HashMap::new(),
                external_data: None,
                version: None,
                feature_list: None,
                fallback_img: None,
            },
            in_chart_space: false,
            in_chart_data: false,
            in_chart: false,
            in_plot_area: false,
            in_plot_area_region: false,
            in_data: false,
            data_id: 0,
            data_dims: Vec::new(),
            in_str_dim: false,
            in_num_dim: false,
            dim_str_type: StringDimType::Cat,
            dim_num_type: NumericDimType::Val,
            dim_formula: None,
            dim_nf: None,
            in_dim_f: false,
            in_dim_nf: false,
            dim_levels: Vec::new(),
            dim_num_levels: Vec::new(),
            in_lvl: false,
            lvl_pt_count: 0,
            lvl_name: None,
            lvl_format_code: None,
            lvl_str_points: Vec::new(),
            lvl_num_points: Vec::new(),
            in_lvl_pt: false,
            lvl_pt_idx: 0,
            in_lvl_pt_text: false,
            in_title: false,
            title_pos: None,
            title_align: None,
            title_overlay: None,
            in_title_tx: false,
            in_title_tx_data: false,
            in_title_tx_data_v: false,
            in_title_tx_data_f: false,
            title_text: None,
            title_sp: None,
            in_series: false,
            ser_layout: ChartExLayout::Unknown("unknown".into()),
            ser_unique_id: None,
            ser_hidden: None,
            ser_owner_idx: None,
            ser_format_idx: None,
            ser_text: None,
            ser_data_id: 0,
            ser_data_labels: None,
            ser_data_points: Vec::new(),
            ser_layout_pr: None,
            ser_axis_ids: Vec::new(),
            ser_value_colors: None,
            ser_value_color_positions: None,
            ser_shape_properties: None,
            in_ser_tx: false,
            in_ser_tx_data: false,
            in_ser_tx_data_v: false,
            in_ser_tx_data_f: false,
            ser_tx_value: None,
            ser_tx_formula: None,
            in_data_id: false,
            in_data_labels: false,
            dlbl_pos: None,
            dlbl_vis_series: None,
            dlbl_vis_cat: None,
            dlbl_vis_val: None,
            dlbl_num_fmt: None,
            dlbl_separator: None,
            dlbl_sp: None,
            in_dlbl_separator: false,
            dl_overrides: Vec::new(),
            dl_hidden: Vec::new(),
            in_data_label: false,
            lbl_idx: 0,
            lbl_pos: None,
            lbl_vis_series: None,
            lbl_vis_cat: None,
            lbl_vis_val: None,
            lbl_num_fmt: None,
            lbl_separator: None,
            lbl_sp: None,
            in_lbl_separator: false,
            in_data_pt: false,
            dpt_idx: 0,
            dpt_sp: None,
            in_layout_pr: false,
            layout_pr: ChartExLayoutPr::default(),
            in_subtotals: false,
            in_binning: false,
            in_geography: false,
            geo_cache_writer: None,
            geo_cache_depth: 0,
            in_statistics: false,
            in_layout_visibility: false,
            in_axis_id: false,
            in_value_colors: false,
            vc_capturing_tag: None,
            vc_writer: None,
            vc_depth: 0,
            in_value_color_positions: false,
            vcp: ChartExValueColorPositions::default(),
            in_vcp_min: false,
            in_vcp_mid: false,
            in_vcp_max: false,
            in_axis: false,
            ax_id: 0,
            ax_hidden: None,
            ax_scaling: ChartExScaling::Category { gap_width: None },
            ax_title: None,
            ax_units: None,
            ax_major_gridlines: None,
            ax_minor_gridlines: None,
            ax_major_tick_marks: None,
            ax_minor_tick_marks: None,
            ax_tick_labels: false,
            ax_num_fmt: None,
            ax_shape_properties: None,
            in_cat_scaling: false,
            in_val_scaling: false,
            in_ax_title: false,
            in_ax_title_tx: false,
            in_ax_title_tx_data: false,
            in_ax_title_tx_data_v: false,
            ax_title_text: None,
            ax_title_sp: None,
            in_units: false,
            units_unit: None,
            in_major_gridlines: false,
            in_minor_gridlines: false,
            in_legend: false,
            legend_pos: None,
            legend_align: None,
            legend_overlay: None,
            legend_sp: None,
            in_sp_pr: false,
            sp_pr_depth: 0,
            sp_solid_fill: None,
            sp_no_fill: false,
            sp_line: None,
            in_sp_ln: false,
            sp_ln_width: None,
            sp_ln_solid_fill: None,
            sp_ln_no_fill: false,
            sp_ln_dash: None,
            sp_ctx: SpCtx::None,
            in_print_settings: false,
            print_settings: ChartExPrintSettings::default(),
            in_ps_header_footer: false,
            ps_hf: ChartExHeaderFooter::default(),
            in_ps_hf_odd_header: false,
            in_ps_hf_odd_footer: false,
            in_ps_hf_even_header: false,
            in_ps_hf_even_footer: false,
            in_ps_hf_first_header: false,
            in_ps_hf_first_footer: false,
            skip_depth: 0,
            skipping: false,
            in_plot_surface: false,
        }
    }

    /// A raw-capture region reproduces source bytes, so while one is
    /// open the driver must not split a self-closing element.
    fn in_raw_capture(&self) -> bool {
        self.geo_cache_writer.is_some() || self.vc_writer.is_some()
    }

    fn on_start(&mut self, e: &BytesStart) {
        // geoCache raw capture
        if let Some(ref mut w) = self.geo_cache_writer {
            let _ = w.write_event(Event::Start(e.clone().into_owned()));
            self.geo_cache_depth += 1;
            return;
        }
        // valueColors raw capture
        if let Some(ref mut w) = self.vc_writer {
            let _ = w.write_event(Event::Start(e.clone().into_owned()));
            self.vc_depth += 1;
            return;
        }
        // skip unmodeled
        if self.skipping {
            self.skip_depth += 1;
            return;
        }

        let local = e.name().local_name();
        let tag = local.as_ref();

        // Inside a cx:spPr every element start deepens the
        // nesting and its end unwinds it, including the end
        // synthesized for a self-closing element. Tracked here
        // rather than in individual arms so that adding an arm
        // cannot unbalance it.
        if self.in_sp_pr {
            self.sp_pr_depth += 1;
        }

        match tag {
            b"chartSpace" => {
                self.in_chart_space = true;
                for attr in e.attributes().flatten() {
                    let value = attr.unescape_value().ok().map(|v| v.to_string());
                    match attr.key.local_name().as_ref() {
                        b"version" => self.result.version = value,
                        b"featureList" => self.result.feature_list = value,
                        b"fallbackImg" => self.result.fallback_img = value,
                        _ => {}
                    }
                }
            }
            b"chartData" if self.in_chart_space => self.in_chart_data = true,
            b"data" if self.in_chart_data => {
                self.in_data = true;
                self.data_id = 0;
                self.data_dims.clear();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"id" {
                        self.data_id = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"strDim" if self.in_data => {
                self.in_str_dim = true;
                self.dim_formula = None;
                self.dim_nf = None;
                self.dim_levels.clear();
                self.dim_str_type = StringDimType::Cat;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        if let Ok(v) = attr.unescape_value() {
                            self.dim_str_type = match v.as_ref() {
                                "colorStr" => StringDimType::ColorStr,
                                "entityId" => StringDimType::EntityId,
                                _ => StringDimType::Cat,
                            };
                        }
                    }
                }
            }
            b"numDim" if self.in_data => {
                self.in_num_dim = true;
                self.dim_formula = None;
                self.dim_nf = None;
                self.dim_num_levels.clear();
                self.dim_num_type = NumericDimType::Val;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        if let Ok(v) = attr.unescape_value() {
                            self.dim_num_type = match v.as_ref() {
                                "x" => NumericDimType::X,
                                "y" => NumericDimType::Y,
                                "size" => NumericDimType::Size,
                                "colorVal" => NumericDimType::ColorVal,
                                _ => NumericDimType::Val,
                            };
                        }
                    }
                }
            }
            b"f" if (self.in_str_dim || self.in_num_dim) && !self.in_lvl => self.in_dim_f = true,
            b"nf" if (self.in_str_dim || self.in_num_dim) && !self.in_lvl => self.in_dim_nf = true,
            b"lvl" if self.in_str_dim || self.in_num_dim => {
                self.in_lvl = true;
                self.lvl_pt_count = 0;
                self.lvl_name = None;
                self.lvl_format_code = None;
                self.lvl_str_points.clear();
                self.lvl_num_points.clear();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"ptCount" => {
                            self.lvl_pt_count = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"name" => {
                            self.lvl_name = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        b"formatCode" => {
                            self.lvl_format_code =
                                attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        _ => {}
                    }
                }
            }
            b"pt" if self.in_lvl => {
                self.in_lvl_pt = true;
                self.lvl_pt_idx = 0;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"idx" {
                        self.lvl_pt_idx = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
                self.in_lvl_pt_text = true;
            }
            b"chart" if self.in_chart_space && !self.in_chart => self.in_chart = true,
            b"title" if self.in_chart && !self.in_plot_area && !self.in_axis && !self.in_title => {
                self.in_title = true;
                self.title_pos = None;
                self.title_align = None;
                self.title_overlay = None;
                self.title_text = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"pos" => {
                            self.title_pos = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"align" => {
                            self.title_align = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"overlay" => {
                            self.title_overlay = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"tx" if self.in_title && !self.in_series => self.in_title_tx = true,
            b"txData" if self.in_title_tx && !self.in_series => self.in_title_tx_data = true,
            b"v" if self.in_title_tx_data && !self.in_series => self.in_title_tx_data_v = true,
            b"f" if self.in_title_tx_data && !self.in_series => self.in_title_tx_data_f = true,
            b"plotArea" if self.in_chart => self.in_plot_area = true,
            b"plotAreaRegion" if self.in_plot_area => self.in_plot_area_region = true,
            b"plotSurface" if self.in_plot_area && !self.in_plot_area_region => {
                self.in_plot_surface = true;
            }
            b"series" if self.in_plot_area_region => {
                self.in_series = true;
                self.ser_layout = ChartExLayout::Unknown("unknown".into());
                self.ser_unique_id = None;
                self.ser_hidden = None;
                self.ser_owner_idx = None;
                self.ser_format_idx = None;
                self.ser_text = None;
                self.ser_data_id = 0;
                self.ser_data_labels = None;
                self.ser_data_points.clear();
                self.ser_layout_pr = None;
                self.ser_axis_ids.clear();
                self.ser_value_colors = None;
                self.ser_value_color_positions = None;
                self.ser_shape_properties = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"layoutId" => {
                            if let Ok(v) = attr.unescape_value() {
                                self.ser_layout = parse_layout_id(&v);
                            }
                        }
                        b"uniqueId" => {
                            self.ser_unique_id =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"hidden" => {
                            self.ser_hidden = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        b"ownerIdx" => {
                            self.ser_owner_idx =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"formatIdx" => {
                            self.ser_format_idx =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        _ => {}
                    }
                }
            }
            b"tx" if self.in_series && !self.in_data_labels => {
                self.in_ser_tx = true;
                self.ser_tx_value = None;
                self.ser_tx_formula = None;
            }
            b"txData" if self.in_ser_tx => self.in_ser_tx_data = true,
            b"v" if self.in_ser_tx_data => self.in_ser_tx_data_v = true,
            b"f" if self.in_ser_tx_data => self.in_ser_tx_data_f = true,
            b"dataId" if self.in_series && !self.in_data_labels => {
                self.in_data_id = true;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"val" {
                        self.ser_data_id = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"dataLabels" if self.in_series => {
                self.in_data_labels = true;
                self.dl_overrides = Vec::new();
                self.dl_hidden = Vec::new();
                self.dlbl_pos = None;
                self.dlbl_vis_series = None;
                self.dlbl_vis_cat = None;
                self.dlbl_vis_val = None;
                self.dlbl_num_fmt = None;
                self.dlbl_separator = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"pos" {
                        self.dlbl_pos = attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"dataLabel" if self.in_data_labels && !self.in_data_label => {
                self.in_data_label = true;
                self.lbl_idx = 0;
                self.lbl_pos = None;
                self.lbl_vis_series = None;
                self.lbl_vis_cat = None;
                self.lbl_vis_val = None;
                self.lbl_num_fmt = None;
                self.lbl_separator = None;
                self.lbl_sp = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"idx" => {
                            self.lbl_idx = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"pos" => {
                            self.lbl_pos = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        _ => {}
                    }
                }
            }
            b"dataLabelHidden" if self.in_data_labels => {
                if let Some(idx) =
                    get_val_attr_named(e, b"idx").and_then(|s| s.parse().ok())
                {
                    self.dl_hidden.push(idx);
                }
            }
            b"separator" if self.in_data_label => self.in_lbl_separator = true,
            b"separator" if self.in_data_labels => self.in_dlbl_separator = true,
            b"dataPt" if self.in_series && !self.in_data_labels => {
                self.in_data_pt = true;
                self.dpt_idx = 0;
                self.dpt_sp = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"idx" {
                        self.dpt_idx = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"layoutPr" if self.in_series => {
                self.in_layout_pr = true;
                self.layout_pr = ChartExLayoutPr::default();
            }
            b"subtotals" if self.in_layout_pr => {
                self.in_subtotals = true;
                self.layout_pr.subtotals.get_or_insert_with(Vec::new);
            }
            b"idx" if self.in_subtotals => push_subtotal_idx(&mut self.layout_pr, e),
            b"binning" if self.in_layout_pr => {
                self.in_binning = true;
                let mut binning = ChartExBinning::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"intervalClosed" => {
                            binning.interval_closed =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"underflow" => {
                            binning.underflow =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"overflow" => {
                            binning.overflow =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        _ => {}
                    }
                }
                self.layout_pr.binning = Some(binning);
            }
            b"binSize" if self.in_binning => {
                if let (Some(v), Some(b)) = (get_val_f64(e), self.layout_pr.binning.as_mut()) {
                    b.bin_size = Some(v);
                }
            }
            b"binCount" if self.in_binning => {
                if let (Some(v), Some(b)) = (
                    get_val_attr(e).and_then(|s| s.parse().ok()),
                    self.layout_pr.binning.as_mut(),
                ) {
                    b.bin_count = Some(v);
                }
            }
            b"geography" if self.in_layout_pr => {
                self.in_geography = true;
                let mut geo = ChartExGeography::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"projectionType" => {
                            geo.projection_type =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"viewedRegionType" => {
                            geo.viewed_region_type =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"cultureLanguage" => {
                            geo.culture_language =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"cultureRegion" => {
                            geo.culture_region =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"attribution" => {
                            geo.attribution =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        _ => {}
                    }
                }
                self.layout_pr.geography = Some(geo);
            }
            b"geoCache" if self.in_geography => {
                let mut w = Writer::new(Cursor::new(Vec::new()));
                let _ = w.write_event(Event::Start(e.clone().into_owned()));
                self.geo_cache_depth = 1;
                self.geo_cache_writer = Some(w);
            }
            b"statistics" if self.in_layout_pr => {
                self.in_statistics = true;
                let mut stats = ChartExStatistics::default();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"quartileMethod" {
                        stats.quartile_method =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
                self.layout_pr.statistics = Some(stats);
            }
            b"visibility" if self.in_layout_pr && !self.in_data_labels => {
                self.in_layout_visibility = true;
                let mut vis = ChartExSeriesVisibility::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"connectorLines" => vis.connector_lines = parse_bool_attr(&attr),
                        b"meanLine" => vis.mean_line = parse_bool_attr(&attr),
                        b"meanMarker" => vis.mean_marker = parse_bool_attr(&attr),
                        b"nonoutliers" => vis.nonoutliers = parse_bool_attr(&attr),
                        b"outliers" => vis.outliers = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
                self.layout_pr.visibility = Some(vis);
            }
            b"axisId" if self.in_series && !self.in_axis => self.in_axis_id = true,
            b"valueColors" if self.in_series => {
                self.in_value_colors = true;
                self.ser_value_colors = Some(ChartExValueColors::default());
            }
            b"minColor" | b"midColor" | b"maxColor"
                if self.in_value_colors && self.vc_writer.is_none() =>
            {
                let tag_name = std::str::from_utf8(tag).unwrap_or("").to_string();
                self.vc_capturing_tag = Some(tag_name);
                let mut w = Writer::new(Cursor::new(Vec::new()));
                let _ = w.write_event(Event::Start(e.clone().into_owned()));
                self.vc_depth = 1;
                self.vc_writer = Some(w);
            }
            b"valueColorPositions" if self.in_series => {
                self.in_value_color_positions = true;
                self.vcp = ChartExValueColorPositions::default();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"count" {
                        self.vcp.count = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
            }
            b"min" if self.in_value_color_positions => self.in_vcp_min = true,
            b"mid" if self.in_value_color_positions => self.in_vcp_mid = true,
            b"max" if self.in_value_color_positions => self.in_vcp_max = true,
            b"axis" if self.in_plot_area && !self.in_plot_area_region => {
                self.in_axis = true;
                self.ax_id = 0;
                self.ax_hidden = None;
                self.ax_scaling = ChartExScaling::Category { gap_width: None };
                self.ax_title = None;
                self.ax_units = None;
                self.ax_major_gridlines = None;
                self.ax_minor_gridlines = None;
                self.ax_major_tick_marks = None;
                self.ax_minor_tick_marks = None;
                self.ax_tick_labels = false;
                self.ax_num_fmt = None;
                self.ax_shape_properties = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"id" => {
                            self.ax_id = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0)
                        }
                        b"hidden" => {
                            self.ax_hidden = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"catScaling" if self.in_axis => {
                self.in_cat_scaling = true;
                let mut gw = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"gapWidth" {
                        gw = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
                self.ax_scaling = ChartExScaling::Category { gap_width: gw };
            }
            b"valScaling" if self.in_axis => {
                self.in_val_scaling = true;
                let mut min = None;
                let mut max = None;
                let mut major = None;
                let mut minor = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"min" => {
                            min = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"max" => {
                            max = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"majorUnit" => {
                            major = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"minorUnit" => {
                            minor = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        _ => {}
                    }
                }
                self.ax_scaling = ChartExScaling::Value {
                    min,
                    max,
                    major_unit: major,
                    minor_unit: minor,
                };
            }
            b"title" if self.in_axis => {
                self.in_ax_title = true;
                self.ax_title_text = None;
            }
            b"tx" if self.in_ax_title => self.in_ax_title_tx = true,
            b"txData" if self.in_ax_title_tx => self.in_ax_title_tx_data = true,
            b"v" if self.in_ax_title_tx_data => self.in_ax_title_tx_data_v = true,
            b"units" if self.in_axis => {
                self.in_units = true;
                self.units_unit = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"unit" {
                        self.units_unit = attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"majorGridlines" if self.in_axis => {
                self.in_major_gridlines = true;
                self.ax_major_gridlines = Some(ChartExGridlines::default());
            }
            b"minorGridlines" if self.in_axis => {
                self.in_minor_gridlines = true;
                self.ax_minor_gridlines = Some(ChartExGridlines::default());
            }
            b"legend" if self.in_chart && !self.in_plot_area => {
                self.in_legend = true;
                self.legend_pos = None;
                self.legend_align = None;
                self.legend_overlay = None;
                self.legend_sp = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"pos" => {
                            self.legend_pos = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"align" => {
                            self.legend_align = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"overlay" => {
                            self.legend_overlay = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"printSettings" if self.in_chart_space => {
                self.in_print_settings = true;
                self.print_settings = ChartExPrintSettings::default();
            }
            b"headerFooter" if self.in_print_settings => {
                self.in_ps_header_footer = true;
                self.ps_hf = ChartExHeaderFooter::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"alignWithMargins" => {
                            self.ps_hf.align_with_margins = parse_bool_attr(&attr)
                        }
                        b"differentOddEven" => {
                            self.ps_hf.different_odd_even = parse_bool_attr(&attr)
                        }
                        b"differentFirst" => self.ps_hf.different_first = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"oddHeader" if self.in_ps_header_footer => self.in_ps_hf_odd_header = true,
            b"oddFooter" if self.in_ps_header_footer => self.in_ps_hf_odd_footer = true,
            b"evenHeader" if self.in_ps_header_footer => self.in_ps_hf_even_header = true,
            b"evenFooter" if self.in_ps_header_footer => self.in_ps_hf_even_footer = true,
            b"firstHeader" if self.in_ps_header_footer => self.in_ps_hf_first_header = true,
            b"firstFooter" if self.in_ps_header_footer => self.in_ps_hf_first_footer = true,
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
                if self.in_data_pt {
                    self.sp_ctx = SpCtx::DataPt;
                } else if self.in_data_label {
                    self.sp_ctx = SpCtx::DataLabel;
                } else if self.in_data_labels {
                    self.sp_ctx = SpCtx::DataLabels;
                } else if self.in_series {
                    self.sp_ctx = SpCtx::Series;
                } else if self.in_major_gridlines {
                    self.sp_ctx = SpCtx::MajorGrid;
                } else if self.in_minor_gridlines {
                    self.sp_ctx = SpCtx::MinorGrid;
                } else if self.in_ax_title {
                    self.sp_ctx = SpCtx::AxisTitle;
                } else if self.in_axis {
                    self.sp_ctx = SpCtx::Axis;
                } else if self.in_title {
                    self.sp_ctx = SpCtx::Title;
                } else if self.in_legend {
                    self.sp_ctx = SpCtx::Legend;
                } else if self.in_plot_surface {
                    self.sp_ctx = SpCtx::PlotSurface;
                } else if self.in_chart_space && !self.in_chart {
                    self.sp_ctx = SpCtx::ChartSpace;
                } else {
                    self.sp_ctx = SpCtx::None;
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
                        self.sp_ln_width =
                            attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
            }
            b"txPr" => {
                self.skipping = true;
                self.skip_depth = 1;
            }
            b"rich" => {
                self.skipping = true;
                self.skip_depth = 1;
            }
            b"clrMapOvr" => {
                self.skipping = true;
                self.skip_depth = 1;
            }
            b"fmtOvrs" => {
                self.skipping = true;
                self.skip_depth = 1;
            }
            b"extLst" => {
                self.skipping = true;
                self.skip_depth = 1;
            }
            // Handling merged from the former separate arm for
            // empty elements, which a self-closing element no
            // longer reaches.
            b"externalData" if self.in_chart_data => {
                let mut rel_id = String::new();
                let mut auto_update = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"id" => {
                            rel_id = attr
                                .unescape_value()
                                .map(|s| s.to_string())
                                .unwrap_or_default();
                        }
                        b"autoUpdate" => {
                            auto_update = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s.as_ref() == "1" || s.as_ref() == "true");
                        }
                        _ => {}
                    }
                }
                self.result.external_data = Some(ChartExExternalData {
                    rel_id,
                    auto_update,
                });
            }
            b"visibility" if self.in_data_label => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"seriesName" => self.lbl_vis_series = parse_bool_attr(&attr),
                        b"categoryName" => self.lbl_vis_cat = parse_bool_attr(&attr),
                        b"value" => self.lbl_vis_val = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"visibility" if self.in_data_labels => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"seriesName" => self.dlbl_vis_series = parse_bool_attr(&attr),
                        b"categoryName" => self.dlbl_vis_cat = parse_bool_attr(&attr),
                        b"value" => self.dlbl_vis_val = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"numFmt" if self.in_data_label => self.lbl_num_fmt = Some(parse_num_fmt(e)),
            b"numFmt" if self.in_data_labels => self.dlbl_num_fmt = Some(parse_num_fmt(e)),
            b"numFmt" if self.in_axis && !self.in_data_labels => {
                self.ax_num_fmt = Some(parse_num_fmt(e));
            }
            b"parentLabelLayout" if self.in_layout_pr => {
                self.layout_pr.parent_label_layout = get_val_attr(e);
            }
            b"regionLabelLayout" if self.in_layout_pr => {
                self.layout_pr.region_label_layout = get_val_attr(e);
            }
            b"aggregation" if self.in_layout_pr => self.layout_pr.aggregation = true,
            b"majorTickMarks" if self.in_axis => {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        self.ax_major_tick_marks =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"minorTickMarks" if self.in_axis => {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        self.ax_minor_tick_marks =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"tickLabels" if self.in_axis => self.ax_tick_labels = true,
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
            b"pageMargins" if self.in_print_settings => {
                let mut pm = ChartExPageMargins::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"l" => {
                            pm.left =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"r" => {
                            pm.right =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"t" => {
                            pm.top = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"b" => {
                            pm.bottom =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"header" => {
                            pm.header =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"footer" => {
                            pm.footer =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        _ => {}
                    }
                }
                self.print_settings.page_margins = Some(pm);
            }
            b"pageSetup" if self.in_print_settings => {
                let mut ps = ChartExPageSetup::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"paperSize" => {
                            ps.paper_size =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"firstPageNumber" => {
                            ps.first_page_number =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"orientation" => {
                            ps.orientation =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"blackAndWhite" => ps.black_and_white = parse_bool_attr(&attr),
                        b"draft" => ps.draft = parse_bool_attr(&attr),
                        b"useFirstPageNumber" => {
                            ps.use_first_page_number = parse_bool_attr(&attr)
                        }
                        b"horizontalDpi" => {
                            ps.horizontal_dpi =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"verticalDpi" => {
                            ps.vertical_dpi =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"copies" => {
                            ps.copies =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        _ => {}
                    }
                }
                self.print_settings.page_setup = Some(ps);
            }
            b"extremeValue" if self.in_vcp_min => {
                self.vcp.min = Some(ChartExColorPosition::ExtremeValue)
            }
            b"extremeValue" if self.in_vcp_mid => {
                self.vcp.mid = Some(ChartExColorPosition::ExtremeValue)
            }
            b"extremeValue" if self.in_vcp_max => {
                self.vcp.max = Some(ChartExColorPosition::ExtremeValue)
            }
            b"number" if self.in_vcp_min => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.min = Some(ChartExColorPosition::Number(v));
                }
            }
            b"number" if self.in_vcp_mid => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.mid = Some(ChartExColorPosition::Number(v));
                }
            }
            b"number" if self.in_vcp_max => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.max = Some(ChartExColorPosition::Number(v));
                }
            }
            b"percent" if self.in_vcp_min => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.min = Some(ChartExColorPosition::Percent(v));
                }
            }
            b"percent" if self.in_vcp_mid => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.mid = Some(ChartExColorPosition::Percent(v));
                }
            }
            b"percent" if self.in_vcp_max => {
                if let Some(v) = get_val_f64(e) {
                    self.vcp.max = Some(ChartExColorPosition::Percent(v));
                }
            }
            _ => {}
        }
    }

    fn on_empty(&mut self, e: &BytesStart) {
        // Only reached inside a raw-capture region; everywhere
        // else the read site split the element into start + end.
        if let Some(ref mut w) = self.geo_cache_writer {
            let _ = w.write_event(Event::Empty(e.clone().into_owned()));
        } else if let Some(ref mut w) = self.vc_writer {
            let _ = w.write_event(Event::Empty(e.clone().into_owned()));
        }
    }

    fn on_text(&mut self, e: &BytesText) {
        if let Some(ref mut w) = self.geo_cache_writer {
            let _ = w.write_event(Event::Text(e.clone().into_owned()));
            return;
        }
        if let Some(ref mut w) = self.vc_writer {
            let _ = w.write_event(Event::Text(e.clone().into_owned()));
            return;
        }
        if self.skipping {
            return;
        }
        if let Ok(text) = e.unescape() {
            let t = text.as_ref();
            if self.in_dim_f {
                self.dim_formula = Some(t.to_string());
            } else if self.in_dim_nf {
                self.dim_nf = Some(t.to_string());
            } else if self.in_lvl_pt_text && self.in_lvl_pt {
                if self.in_str_dim {
                    self.lvl_str_points.push((self.lvl_pt_idx, t.to_string()));
                } else if self.in_num_dim {
                    self.lvl_num_points.push((self.lvl_pt_idx, t.to_string()));
                }
            } else if self.in_title_tx_data_v {
                self.title_text = Some(t.to_string());
            } else if self.in_title_tx_data_f {
                self.title_text = Some(t.to_string());
            } else if self.in_ser_tx_data_v {
                self.ser_tx_value = Some(t.to_string());
            } else if self.in_ser_tx_data_f {
                self.ser_tx_formula = Some(t.to_string());
            } else if self.in_axis_id {
                if let Ok(id) = t.parse::<u32>() {
                    self.ser_axis_ids.push(id);
                }
            } else if self.in_ax_title_tx_data_v {
                self.ax_title_text = Some(t.to_string());
            } else if self.in_lbl_separator {
                self.lbl_separator = Some(t.to_string());
            } else if self.in_dlbl_separator {
                self.dlbl_separator = Some(t.to_string());
            } else if self.in_ps_hf_odd_header {
                self.ps_hf.odd_header = Some(t.to_string());
            } else if self.in_ps_hf_odd_footer {
                self.ps_hf.odd_footer = Some(t.to_string());
            } else if self.in_ps_hf_even_header {
                self.ps_hf.even_header = Some(t.to_string());
            } else if self.in_ps_hf_even_footer {
                self.ps_hf.even_footer = Some(t.to_string());
            } else if self.in_ps_hf_first_header {
                self.ps_hf.first_header = Some(t.to_string());
            } else if self.in_ps_hf_first_footer {
                self.ps_hf.first_footer = Some(t.to_string());
            }
        }
    }

    fn on_end(&mut self, e: &BytesEnd) {
        // geoCache raw capture
        if let Some(ref mut w) = self.geo_cache_writer {
            self.geo_cache_depth -= 1;
            let _ = w.write_event(Event::End(e.clone().into_owned()));
            if self.geo_cache_depth == 0 {
                if let Some(w) = self.geo_cache_writer.take() {
                    if let Some(ref mut geo) = self.layout_pr.geography {
                        geo.raw_geo_cache = Some(w.into_inner().into_inner());
                    }
                }
            }
            return;
        }
        // valueColor capture
        if let Some(ref mut w) = self.vc_writer {
            self.vc_depth -= 1;
            let _ = w.write_event(Event::End(e.clone().into_owned()));
            if self.vc_depth == 0 {
                if let Some(w) = self.vc_writer.take() {
                    let bytes = w.into_inner().into_inner();
                    if let Some(ref mut vc) = self.ser_value_colors {
                        match self.vc_capturing_tag.as_deref() {
                            Some("minColor") => vc.min_color = Some(bytes),
                            Some("midColor") => vc.mid_color = Some(bytes),
                            Some("maxColor") => vc.max_color = Some(bytes),
                            _ => {}
                        }
                    }
                    self.vc_capturing_tag = None;
                }
            }
            return;
        }
        if self.skipping {
            self.skip_depth -= 1;
            if self.skip_depth == 0 {
                self.skipping = false;
                // The element that opened this skip region was
                // counted by the spPr depth bookkeeping; its
                // subtree ends here. It cannot close the spPr
                // itself, since the spPr's own start put the
                // depth at 1 before the opener raised it.
                if self.in_sp_pr {
                    self.sp_pr_depth = self.sp_pr_depth.saturating_sub(1);
                }
            }
            return;
        }

        let local = e.name().local_name();
        let tag = local.as_ref();
        match tag {
            b"chartSpace" => self.in_chart_space = false,
            b"chartData" => self.in_chart_data = false,
            b"data" if self.in_data => {
                self.result.data.push(ChartExData {
                    id: self.data_id,
                    dimensions: std::mem::take(&mut self.data_dims),
                    extensions: None,
                });
                self.in_data = false;
            }
            b"strDim" if self.in_str_dim => {
                self.data_dims.push(ChartExDimension::String {
                    dim_type: self.dim_str_type.clone(),
                    formula: self.dim_formula.take(),
                    nf_formula: self.dim_nf.take(),
                    levels: std::mem::take(&mut self.dim_levels),
                });
                self.in_str_dim = false;
            }
            b"numDim" if self.in_num_dim => {
                self.data_dims.push(ChartExDimension::Numeric {
                    dim_type: self.dim_num_type.clone(),
                    formula: self.dim_formula.take(),
                    nf_formula: self.dim_nf.take(),
                    levels: std::mem::take(&mut self.dim_num_levels),
                });
                self.in_num_dim = false;
            }
            b"f" if self.in_dim_f => self.in_dim_f = false,
            b"nf" if self.in_dim_nf => self.in_dim_nf = false,
            b"lvl" if self.in_lvl => {
                if self.in_str_dim {
                    self.dim_levels.push(ChartExStringLevel {
                        pt_count: self.lvl_pt_count,
                        name: self.lvl_name.take(),
                        points: std::mem::take(&mut self.lvl_str_points),
                    });
                } else if self.in_num_dim {
                    self.dim_num_levels.push(ChartExNumericLevel {
                        pt_count: self.lvl_pt_count,
                        format_code: self.lvl_format_code.take(),
                        name: self.lvl_name.take(),
                        points: std::mem::take(&mut self.lvl_num_points),
                    });
                }
                self.in_lvl = false;
            }
            b"pt" if self.in_lvl_pt => {
                self.in_lvl_pt = false;
                self.in_lvl_pt_text = false;
            }
            b"chart" if self.in_chart => self.in_chart = false,
            b"title" if self.in_ax_title => {
                let title = ChartExAxisTitle {
                    text: self.ax_title_text.take().map(|t| ChartExText {
                        data: Some(ChartExTextData {
                            formula: None,
                            value: Some(t),
                        }),
                        rich: None,
                    }),
                    offset: None,
                    shape_properties: self.ax_title_sp.take(),
                    text_properties: None,
                    extensions: None,
                };
                self.ax_title = Some(title);
                self.in_ax_title = false;
                self.in_ax_title_tx = false;
                self.in_ax_title_tx_data = false;
                self.in_ax_title_tx_data_v = false;
            }
            b"title" if self.in_title => {
                self.result.title = Some(ChartExTitle {
                    text: self.title_text.take(),
                    rich_text: None,
                    position: self.title_pos.take(),
                    align: self.title_align.take(),
                    overlay: self.title_overlay.take(),
                    offset: None,
                    shape_properties: self.title_sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.in_title = false;
                self.in_title_tx = false;
                self.in_title_tx_data = false;
            }
            b"tx" if self.in_title_tx && !self.in_series => self.in_title_tx = false,
            b"txData" if self.in_title_tx_data && !self.in_series => self.in_title_tx_data = false,
            b"v" if self.in_title_tx_data_v && !self.in_series => self.in_title_tx_data_v = false,
            b"f" if self.in_title_tx_data_f && !self.in_series => self.in_title_tx_data_f = false,
            b"plotArea" => self.in_plot_area = false,
            b"plotAreaRegion" => self.in_plot_area_region = false,
            b"plotSurface" if self.in_plot_surface => self.in_plot_surface = false,
            b"series" if self.in_series => {
                let series = ChartExSeries {
                    layout: self.ser_layout.clone(),
                    unique_id: self.ser_unique_id.take(),
                    hidden: self.ser_hidden.take(),
                    owner_idx: self.ser_owner_idx.take(),
                    format_idx: self.ser_format_idx.take(),
                    text: self.ser_text.take(),
                    data_id: self.ser_data_id,
                    data_labels: self.ser_data_labels.take(),
                    data_points: std::mem::take(&mut self.ser_data_points),
                    layout_properties: self.ser_layout_pr.take(),
                    axis_ids: std::mem::take(&mut self.ser_axis_ids),
                    value_colors: self.ser_value_colors.take(),
                    value_color_positions: self.ser_value_color_positions.take(),
                    shape_properties: self.ser_shape_properties.take(),
                    extensions: None,
                };
                self.result.plot_area.series.push(series);
                self.in_series = false;
            }
            b"tx" if self.in_ser_tx => {
                self.ser_text = Some(ChartExText {
                    data: Some(ChartExTextData {
                        formula: self.ser_tx_formula.take(),
                        value: self.ser_tx_value.take(),
                    }),
                    rich: None,
                });
                self.in_ser_tx = false;
                self.in_ser_tx_data = false;
            }
            b"txData" if self.in_ser_tx_data => self.in_ser_tx_data = false,
            b"v" if self.in_ser_tx_data_v => self.in_ser_tx_data_v = false,
            b"f" if self.in_ser_tx_data_f => self.in_ser_tx_data_f = false,
            b"dataId" if self.in_data_id => self.in_data_id = false,
            b"dataLabels" if self.in_data_labels => {
                self.ser_data_labels = Some(ChartExDataLabels {
                    position: self.dlbl_pos.take(),
                    visibility_series_name: self.dlbl_vis_series.take(),
                    visibility_category_name: self.dlbl_vis_cat.take(),
                    visibility_value: self.dlbl_vis_val.take(),
                    number_format: self.dlbl_num_fmt.take(),
                    separator: self.dlbl_separator.take(),
                    shape_properties: self.dlbl_sp.take(),
                    overrides: std::mem::take(&mut self.dl_overrides),
                    hidden_labels: std::mem::take(&mut self.dl_hidden),
                    text_properties: None,
                    extensions: None,
                });
                self.in_data_labels = false;
            }
            b"dataLabel" if self.in_data_label => {
                self.dl_overrides.push(ChartExDataLabel {
                    idx: self.lbl_idx,
                    position: self.lbl_pos.take(),
                    visibility_series_name: self.lbl_vis_series.take(),
                    visibility_category_name: self.lbl_vis_cat.take(),
                    visibility_value: self.lbl_vis_val.take(),
                    number_format: self.lbl_num_fmt.take(),
                    separator: self.lbl_separator.take(),
                    shape_properties: self.lbl_sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.in_data_label = false;
            }
            b"separator" if self.in_lbl_separator => self.in_lbl_separator = false,
            b"separator" if self.in_dlbl_separator => self.in_dlbl_separator = false,
            b"dataPt" if self.in_data_pt => {
                self.ser_data_points.push(ChartExDataPoint {
                    idx: self.dpt_idx,
                    shape_properties: self.dpt_sp.take(),
                    extensions: None,
                });
                self.in_data_pt = false;
            }
            b"layoutPr" if self.in_layout_pr => {
                self.ser_layout_pr = Some(std::mem::take(&mut self.layout_pr));
                self.in_layout_pr = false;
            }
            b"subtotals" if self.in_subtotals => self.in_subtotals = false,
            b"binning" if self.in_binning => self.in_binning = false,
            b"geography" if self.in_geography => self.in_geography = false,
            b"statistics" if self.in_statistics => self.in_statistics = false,
            b"visibility" if self.in_layout_visibility => self.in_layout_visibility = false,
            b"axisId" if self.in_axis_id => self.in_axis_id = false,
            b"valueColors" if self.in_value_colors => self.in_value_colors = false,
            b"valueColorPositions" if self.in_value_color_positions => {
                self.ser_value_color_positions = Some(self.vcp.clone());
                self.in_value_color_positions = false;
                self.in_vcp_min = false;
                self.in_vcp_mid = false;
                self.in_vcp_max = false;
            }
            b"min" if self.in_vcp_min => self.in_vcp_min = false,
            b"mid" if self.in_vcp_mid => self.in_vcp_mid = false,
            b"max" if self.in_vcp_max => self.in_vcp_max = false,
            b"axis" if self.in_axis => {
                self.result.plot_area.axes.push(ChartExAxis {
                    id: self.ax_id,
                    hidden: self.ax_hidden.take(),
                    scaling: self.ax_scaling.clone(),
                    title: self.ax_title.take(),
                    units: self.ax_units.take(),
                    major_gridlines: self.ax_major_gridlines.take(),
                    minor_gridlines: self.ax_minor_gridlines.take(),
                    major_tick_marks: self.ax_major_tick_marks.take(),
                    minor_tick_marks: self.ax_minor_tick_marks.take(),
                    tick_labels: self.ax_tick_labels,
                    number_format: self.ax_num_fmt.take(),
                    shape_properties: self.ax_shape_properties.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.in_axis = false;
                self.in_cat_scaling = false;
                self.in_val_scaling = false;
            }
            b"catScaling" if self.in_cat_scaling => self.in_cat_scaling = false,
            b"valScaling" if self.in_val_scaling => self.in_val_scaling = false,
            b"tx" if self.in_ax_title_tx => self.in_ax_title_tx = false,
            b"txData" if self.in_ax_title_tx_data => self.in_ax_title_tx_data = false,
            b"v" if self.in_ax_title_tx_data_v => self.in_ax_title_tx_data_v = false,
            b"units" if self.in_units => {
                self.ax_units = Some(ChartExAxisUnits {
                    unit: self.units_unit.take(),
                    label: None,
                    extensions: None,
                });
                self.in_units = false;
            }
            b"majorGridlines" if self.in_major_gridlines => self.in_major_gridlines = false,
            b"minorGridlines" if self.in_minor_gridlines => self.in_minor_gridlines = false,
            b"legend" if self.in_legend => {
                self.result.legend = Some(ChartExLegend {
                    position: self.legend_pos.take(),
                    align: self.legend_align.take(),
                    overlay: self.legend_overlay.take(),
                    offset: None,
                    shape_properties: self.legend_sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.in_legend = false;
            }
            b"printSettings" if self.in_print_settings => {
                self.result.print_settings = Some(self.print_settings.clone());
                self.in_print_settings = false;
            }
            b"headerFooter" if self.in_ps_header_footer => {
                self.print_settings.header_footer = Some(self.ps_hf.clone());
                self.in_ps_header_footer = false;
            }
            b"oddHeader" => self.in_ps_hf_odd_header = false,
            b"oddFooter" => self.in_ps_hf_odd_footer = false,
            b"evenHeader" => self.in_ps_hf_even_header = false,
            b"evenFooter" => self.in_ps_hf_even_footer = false,
            b"firstHeader" => self.in_ps_hf_first_header = false,
            b"firstFooter" => self.in_ps_hf_first_footer = false,
            b"ln" if self.in_sp_ln => {
                self.sp_line = Some(ChartLine {
                    width: self.sp_ln_width.take(),
                    solid_fill: self.sp_ln_solid_fill.take(),
                    no_fill: self.sp_ln_no_fill,
                    dash_style: self.sp_ln_dash.take(),
                });
                self.in_sp_ln = false;
                self.sp_ln_no_fill = false;
                self.sp_pr_depth = self.sp_pr_depth.saturating_sub(1);
            }
            b"spPr" if self.in_sp_pr => {
                self.sp_pr_depth = self.sp_pr_depth.saturating_sub(1);
                if self.sp_pr_depth == 0 {
                    let props = ChartShapeProperties {
                        solid_fill: self.sp_solid_fill.take(),
                        no_fill: self.sp_no_fill,
                        line: self.sp_line.take(),
                    };
                    let has =
                        props.solid_fill.is_some() || props.no_fill || props.line.is_some();
                    match self.sp_ctx {
                        SpCtx::ChartSpace => {
                            if has {
                                self.result.shape_properties = Some(props);
                            }
                        }
                        SpCtx::Series => {
                            if has {
                                self.ser_shape_properties = Some(props);
                            }
                        }
                        SpCtx::DataPt => {
                            self.dpt_sp = if has { Some(props) } else { None };
                        }
                        SpCtx::DataLabels => {
                            if has {
                                self.dlbl_sp = Some(props);
                            }
                        }
                        SpCtx::DataLabel => {
                            if has {
                                self.lbl_sp = Some(props);
                            }
                        }
                        SpCtx::AxisTitle => {
                            if has {
                                self.ax_title_sp = Some(props);
                            }
                        }
                        SpCtx::Title => {
                            if has {
                                self.title_sp = Some(props);
                            }
                        }
                        SpCtx::Legend => {
                            if has {
                                self.legend_sp = Some(props);
                            }
                        }
                        SpCtx::Axis => {
                            if has {
                                self.ax_shape_properties = Some(props);
                            }
                        }
                        SpCtx::MajorGrid => {
                            self.ax_major_gridlines
                                .get_or_insert_with(ChartExGridlines::default)
                                .shape_properties = Some(props);
                        }
                        SpCtx::MinorGrid => {
                            self.ax_minor_gridlines
                                .get_or_insert_with(ChartExGridlines::default)
                                .shape_properties = Some(props);
                        }
                        SpCtx::PlotSurface => {
                            if has {
                                self.result.plot_area.plot_surface = Some(props);
                            }
                        }
                        SpCtx::None => {}
                    }
                    self.in_sp_pr = false;
                    self.sp_no_fill = false;
                    self.sp_ctx = SpCtx::None;
                }
            }
            _ => {
                if self.in_sp_pr {
                    self.sp_pr_depth = self.sp_pr_depth.saturating_sub(1);
                    if self.sp_pr_depth == 0 {
                        self.in_sp_pr = false;
                        self.sp_no_fill = false;
                        self.sp_ctx = SpCtx::None;
                    }
                }
            }
        }
    }

    fn finish(self) -> ParsedChartEx {
        self.result
    }
}

fn parse_chart_ex_xml_inner<R: Read>(
    xml_reader: &mut Reader<BufReader<R>>,
    buf: &mut Vec<u8>,
) -> ChartParseResult<ParsedChartEx> {
    let mut parser = ChartExParser::new();
    // `<x/>` and `<x></x>` are the same document, so a self-closing
    // element is split into the start and end its expanded form would
    // produce and handled by one code path; keeping two let them drift.
    // Elements inside a raw-capture region are exempt so the captured
    // bytes match the source (a self-closing capture *opener* is still
    // split, since the capture starts after it).
    let mut pending_end: Option<Event<'static>> = None;


    loop {
        let event: Event<'static> = match pending_end.take() {
            Some(end) => end,
            None => {
                let read = match xml_reader.read_event_into(buf) {
                    Ok(ev) => ev.into_owned(),
                    Err(e) => return Err(ChartParseError::Xml(e)),
                };
                buf.clear();
                match read {
                    Event::Empty(e) if !parser.in_raw_capture() => {
                        pending_end = Some(Event::End(BytesEnd::new(
                            String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        )));
                        Event::Start(e)
                    }
                    other => other,
                }
            }
        };

        match event {
            Event::Start(ref e) => parser.on_start(e),
            Event::Empty(ref e) => parser.on_empty(e),
            Event::Text(ref e) => parser.on_text(e),
            Event::End(ref e) => parser.on_end(e),
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(parser.finish())
}

fn parse_layout_id(s: &str) -> ChartExLayout {
    match s {
        "waterfall" => ChartExLayout::Waterfall,
        "treemap" => ChartExLayout::Treemap,
        "sunburst" => ChartExLayout::Sunburst,
        "funnel" => ChartExLayout::Funnel,
        "histogram" => ChartExLayout::Histogram,
        "boxWhisker" => ChartExLayout::BoxWhisker,
        "paretoLine" => ChartExLayout::ParetoLine,
        "regionMap" => ChartExLayout::RegionMap,
        "clusteredColumn" => ChartExLayout::ClusteredColumn,
        other => ChartExLayout::Unknown(other.to_string()),
    }
}

fn get_val_attr(e: &quick_xml::events::BytesStart) -> Option<String> {
    get_val_attr_named(e, b"val")
}

fn get_val_attr_named(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == name {
            return attr.unescape_value().ok().map(|s| s.to_string());
        }
    }
    None
}

fn get_val_f64(e: &quick_xml::events::BytesStart) -> Option<f64> {
    get_val_attr(e).and_then(|s| s.parse().ok())
}

fn parse_bool_attr(attr: &quick_xml::events::attributes::Attribute) -> Option<bool> {
    attr.unescape_value()
        .ok()
        .map(|s| s == "1" || s.as_ref() == "true")
}

/// Record one `cx:idx` of a `cx:subtotals` list. The index rides the `val`
/// attribute ([MS-ODRAWXML] §5.22 `CT_SubtotalIndex`) of a normally
/// self-closing element, not a text node.
fn push_subtotal_idx(layout_pr: &mut ChartExLayoutPr, e: &quick_xml::events::BytesStart) {
    if let Some(idx) = get_val_attr(e).and_then(|s| s.parse::<u32>().ok()) {
        layout_pr.subtotals.get_or_insert_with(Vec::new).push(idx);
    }
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

#[cfg(test)]
mod subtotal_tests {
    use super::*;

    /// Wrap `layout_pr` in the smallest chartEx document that parses.
    fn doc_with_layout_pr(layout_pr: &str) -> Vec<u8> {
        format!(
            r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$B$1</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/>{layout_pr}</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
        )
        .into_bytes()
    }

    fn subtotals_of(layout_pr: &str) -> Option<Vec<u32>> {
        let doc = doc_with_layout_pr(layout_pr);
        let cx = parse_chart_ex_xml(&doc[..]).expect("parse");
        cx.plot_area.series[0]
            .layout_properties
            .as_ref()
            .expect("layoutPr")
            .subtotals
            .clone()
    }

    /// `cx:idx` carries its value in the `val` attribute on a self-closing
    /// element ([MS-ODRAWXML] CT_SubtotalIndex), not in a text node.
    #[test]
    fn subtotal_indices_parse_from_val_attribute() {
        assert_eq!(
            subtotals_of(r#"<cx:layoutPr><cx:subtotals><cx:idx val="0"/><cx:idx val="2"/></cx:subtotals></cx:layoutPr>"#),
            Some(vec![0, 2])
        );
    }

    /// The wrapper is optional (`minOccurs="0"`) but real waterfall charts
    /// carry it empty, so present-but-empty must be distinguishable.
    #[test]
    fn empty_subtotals_element_parses_as_present_and_empty() {
        assert_eq!(
            subtotals_of(r#"<cx:layoutPr><cx:subtotals/></cx:layoutPr>"#),
            Some(Vec::new())
        );
    }

    #[test]
    fn absent_subtotals_element_parses_as_none() {
        assert_eq!(subtotals_of(r#"<cx:layoutPr></cx:layoutPr>"#), None);
    }

    /// Non-self-closing form must parse identically.
    #[test]
    fn subtotal_indices_parse_from_expanded_form() {
        assert_eq!(
            subtotals_of(r#"<cx:layoutPr><cx:subtotals><cx:idx val="3"></cx:idx></cx:subtotals></cx:layoutPr>"#),
            Some(vec![3])
        );
    }
}
