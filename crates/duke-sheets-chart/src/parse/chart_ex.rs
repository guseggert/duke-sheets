use std::io::{BufReader, Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::chart_ex::*;
use crate::error::{ChartParseError, ChartParseResult};
use crate::{ChartColor, ChartLine, ChartShapeProperties, DrawingAnchor, NumberFormat};

/// Parse chart-ex XML from a reader and return a `ChartEx`.
///
/// The anchor will be set to `DrawingAnchor::default()`; callers should
/// override it from the drawing XML.
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
        anchor: DrawingAnchor::default(),
        shape_properties: parsed.shape_properties,
        text_properties: None,
        color_map_override: None,
        format_overrides: Vec::new(),
        print_settings: parsed.print_settings,
        raw_chart_style: None,
        raw_chart_color_style: None,
        raw_extensions: parsed.raw_extensions,
        raw_mc_fallback: None,
        version: None,
        feature_list: None,
        fallback_img: None,
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
}

fn parse_chart_ex_xml_inner<R: Read>(
    xml_reader: &mut Reader<BufReader<R>>,
    buf: &mut Vec<u8>,
) -> ChartParseResult<ParsedChartEx> {
    let mut result = ParsedChartEx {
        title: None,
        data: Vec::new(),
        plot_area: ChartExPlotArea::default(),
        legend: None,
        shape_properties: None,
        print_settings: None,
        raw_extensions: std::collections::HashMap::new(),
        external_data: None,
    };

    let mut in_chart_space = false;
    let mut in_chart_data = false;
    let mut in_chart = false;
    let mut in_plot_area = false;
    let mut in_plot_area_region = false;

    // cx:data state
    let mut in_data = false;
    let mut data_id: u32 = 0;
    let mut data_dims: Vec<ChartExDimension> = Vec::new();

    // dimension state
    let mut in_str_dim = false;
    let mut in_num_dim = false;
    let mut dim_str_type = StringDimType::Cat;
    let mut dim_num_type = NumericDimType::Val;
    let mut dim_formula: Option<String> = None;
    let mut dim_nf: Option<String> = None;
    let mut in_dim_f = false;
    let mut in_dim_nf = false;
    let mut dim_levels: Vec<ChartExStringLevel> = Vec::new();
    let mut dim_num_levels: Vec<ChartExNumericLevel> = Vec::new();
    let mut in_lvl = false;
    let mut lvl_pt_count: u32 = 0;
    let mut lvl_name: Option<String> = None;
    let mut lvl_format_code: Option<String> = None;
    let mut lvl_str_points: Vec<(u32, String)> = Vec::new();
    let mut lvl_num_points: Vec<(u32, String)> = Vec::new();
    let mut in_lvl_pt = false;
    let mut lvl_pt_idx: u32 = 0;
    let mut in_lvl_pt_text = false;

    // Title state
    let mut in_title = false;
    let mut title_pos: Option<String> = None;
    let mut title_align: Option<String> = None;
    let mut title_overlay: Option<bool> = None;
    let mut in_title_tx = false;
    let mut in_title_tx_data = false;
    let mut in_title_tx_data_v = false;
    let mut in_title_tx_data_f = false;
    let mut title_text: Option<String> = None;

    // Series state
    let mut in_series = false;
    let mut ser_layout = ChartExLayout::Unknown("unknown".into());
    let mut ser_unique_id: Option<String> = None;
    let mut ser_hidden: Option<bool> = None;
    let mut ser_owner_idx: Option<u32> = None;
    let mut ser_format_idx: Option<u32> = None;
    let mut ser_text: Option<ChartExText> = None;
    let mut ser_data_id: u32 = 0;
    let mut ser_data_labels: Option<ChartExDataLabels> = None;
    let mut ser_data_points: Vec<ChartExDataPoint> = Vec::new();
    let mut ser_layout_pr: Option<ChartExLayoutPr> = None;
    let mut ser_axis_ids: Vec<u32> = Vec::new();
    let mut ser_value_colors: Option<ChartExValueColors> = None;
    let mut ser_value_color_positions: Option<ChartExValueColorPositions> = None;
    let mut ser_shape_properties: Option<ChartShapeProperties> = None;

    // Series tx state
    let mut in_ser_tx = false;
    let mut in_ser_tx_data = false;
    let mut in_ser_tx_data_v = false;
    let mut in_ser_tx_data_f = false;
    let mut ser_tx_value: Option<String> = None;
    let mut ser_tx_formula: Option<String> = None;

    // dataId state
    let mut in_data_id = false;

    // dataLabels state
    let mut in_data_labels = false;
    let mut dlbl_pos: Option<String> = None;
    let mut dlbl_vis_series: Option<bool> = None;
    let mut dlbl_vis_cat: Option<bool> = None;
    let mut dlbl_vis_val: Option<bool> = None;
    let mut dlbl_num_fmt: Option<NumberFormat> = None;
    let mut dlbl_separator: Option<String> = None;
    let mut in_dlbl_separator = false;

    // dataPt state
    let mut in_data_pt = false;
    let mut dpt_idx: u32 = 0;
    let mut dpt_sp: Option<ChartShapeProperties> = None;

    // layoutPr state
    let mut in_layout_pr = false;
    let mut layout_pr = ChartExLayoutPr::default();
    let mut in_subtotals = false;
    let mut in_subtotal_idx = false;
    let mut in_binning = false;
    let mut in_geography = false;
    let mut _in_geo_cache = false;
    let mut geo_cache_writer: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut geo_cache_depth: u32 = 0;
    let mut in_statistics = false;
    let mut in_layout_visibility = false;

    // axisId (text content)
    let mut in_axis_id = false;

    // valueColors state
    let mut in_value_colors = false;
    let mut vc_capturing_tag: Option<String> = None;
    let mut vc_writer: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut vc_depth: u32 = 0;

    // valueColorPositions state
    let mut in_value_color_positions = false;
    let mut vcp = ChartExValueColorPositions::default();
    let mut in_vcp_min = false;
    let mut in_vcp_mid = false;
    let mut in_vcp_max = false;

    // binSize / binCount text
    let mut in_bin_size = false;
    let mut in_bin_count = false;

    // Axis state
    let mut in_axis = false;
    let mut ax_id: u32 = 0;
    let mut ax_hidden: Option<bool> = None;
    let mut ax_scaling: ChartExScaling = ChartExScaling::Category { gap_width: None };
    let mut ax_title: Option<ChartExAxisTitle> = None;
    let mut ax_units: Option<ChartExAxisUnits> = None;
    let mut ax_major_gridlines: Option<ChartShapeProperties> = None;
    let mut ax_minor_gridlines: Option<ChartShapeProperties> = None;
    let mut ax_major_tick_marks: Option<String> = None;
    let mut ax_minor_tick_marks: Option<String> = None;
    let mut ax_tick_labels = false;
    let mut ax_num_fmt: Option<NumberFormat> = None;
    let mut ax_shape_properties: Option<ChartShapeProperties> = None;

    let mut in_cat_scaling = false;
    let mut in_val_scaling = false;
    let mut in_ax_title = false;
    let mut in_ax_title_tx = false;
    let mut in_ax_title_tx_data = false;
    let mut in_ax_title_tx_data_v = false;
    let mut ax_title_text: Option<String> = None;
    let mut in_units = false;
    let mut units_unit: Option<String> = None;
    let mut in_major_gridlines = false;
    let mut in_minor_gridlines = false;

    // Legend state
    let mut in_legend = false;
    let mut legend_pos: Option<String> = None;
    let mut legend_align: Option<String> = None;
    let mut legend_overlay: Option<bool> = None;
    let mut legend_sp: Option<ChartShapeProperties> = None;

    // spPr state
    let mut in_sp_pr = false;
    let mut sp_pr_depth: u32 = 0;
    let mut sp_solid_fill: Option<ChartColor> = None;
    let mut sp_no_fill = false;
    let mut sp_line: Option<ChartLine> = None;
    let mut in_sp_ln = false;
    let mut sp_ln_width: Option<i64> = None;
    let mut sp_ln_solid_fill: Option<ChartColor> = None;
    let mut sp_ln_no_fill = false;
    let mut sp_ln_dash: Option<String> = None;

    #[derive(Clone, Copy, PartialEq, Eq)]
    enum SpCtx {
        None,
        ChartSpace,
        Series,
        DataPt,
        Legend,
        Axis,
        MajorGrid,
        MinorGrid,
        DataLabels,
        PlotSurface,
    }
    let mut sp_ctx = SpCtx::None;

    // print settings state
    let mut in_print_settings = false;
    let mut print_settings = ChartExPrintSettings::default();
    let mut in_ps_header_footer = false;
    let mut ps_hf = ChartExHeaderFooter::default();
    let mut in_ps_hf_odd_header = false;
    let mut in_ps_hf_odd_footer = false;
    let mut in_ps_hf_even_header = false;
    let mut in_ps_hf_even_footer = false;
    let mut in_ps_hf_first_header = false;
    let mut in_ps_hf_first_footer = false;

    // Skip state for unmodeled elements (txPr, etc.)
    let mut skip_depth: u32 = 0;
    let mut skipping = false;

    let mut in_plot_surface = false;

    loop {
        match xml_reader.read_event_into(buf) {
            Ok(Event::Start(ref e)) => {
                // geoCache raw capture
                if let Some(ref mut w) = geo_cache_writer {
                    let _ = w.write_event(Event::Start(e.clone().into_owned()));
                    geo_cache_depth += 1;
                    buf.clear();
                    continue;
                }
                // valueColors raw capture
                if let Some(ref mut w) = vc_writer {
                    let _ = w.write_event(Event::Start(e.clone().into_owned()));
                    vc_depth += 1;
                    buf.clear();
                    continue;
                }
                // skip unmodeled
                if skipping {
                    skip_depth += 1;
                    buf.clear();
                    continue;
                }

                let local = e.name().local_name();
                let tag = local.as_ref();

                match tag {
                    b"chartSpace" => in_chart_space = true,
                    b"chartData" if in_chart_space => in_chart_data = true,
                    b"data" if in_chart_data => {
                        in_data = true;
                        data_id = 0;
                        data_dims.clear();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"id" {
                                data_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                        }
                    }
                    b"strDim" if in_data => {
                        in_str_dim = true;
                        dim_formula = None;
                        dim_nf = None;
                        dim_levels.clear();
                        dim_str_type = StringDimType::Cat;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"type" {
                                if let Ok(v) = attr.unescape_value() {
                                    dim_str_type = match v.as_ref() {
                                        "colorStr" => StringDimType::ColorStr,
                                        "entityId" => StringDimType::EntityId,
                                        _ => StringDimType::Cat,
                                    };
                                }
                            }
                        }
                    }
                    b"numDim" if in_data => {
                        in_num_dim = true;
                        dim_formula = None;
                        dim_nf = None;
                        dim_num_levels.clear();
                        dim_num_type = NumericDimType::Val;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"type" {
                                if let Ok(v) = attr.unescape_value() {
                                    dim_num_type = match v.as_ref() {
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
                    b"f" if (in_str_dim || in_num_dim) && !in_lvl => in_dim_f = true,
                    b"nf" if (in_str_dim || in_num_dim) && !in_lvl => in_dim_nf = true,
                    b"lvl" if in_str_dim || in_num_dim => {
                        in_lvl = true;
                        lvl_pt_count = 0;
                        lvl_name = None;
                        lvl_format_code = None;
                        lvl_str_points.clear();
                        lvl_num_points.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"ptCount" => {
                                    lvl_pt_count = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"name" => {
                                    lvl_name = attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                b"formatCode" => {
                                    lvl_format_code =
                                        attr.unescape_value().ok().map(|s| s.to_string());
                                }
                                _ => {}
                            }
                        }
                    }
                    b"pt" if in_lvl => {
                        in_lvl_pt = true;
                        lvl_pt_idx = 0;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"idx" {
                                lvl_pt_idx = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                        }
                        in_lvl_pt_text = true;
                    }
                    b"chart" if in_chart_space && !in_chart => in_chart = true,
                    b"title" if in_chart && !in_plot_area && !in_axis && !in_title => {
                        in_title = true;
                        title_pos = None;
                        title_align = None;
                        title_overlay = None;
                        title_text = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"pos" => {
                                    title_pos = attr.unescape_value().ok().map(|s| s.to_string())
                                }
                                b"align" => {
                                    title_align = attr.unescape_value().ok().map(|s| s.to_string())
                                }
                                b"overlay" => {
                                    title_overlay = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s.as_ref() == "true")
                                }
                                _ => {}
                            }
                        }
                    }
                    b"tx" if in_title && !in_series => in_title_tx = true,
                    b"txData" if in_title_tx && !in_series => in_title_tx_data = true,
                    b"v" if in_title_tx_data && !in_series => in_title_tx_data_v = true,
                    b"f" if in_title_tx_data && !in_series => in_title_tx_data_f = true,
                    b"plotArea" if in_chart => in_plot_area = true,
                    b"plotAreaRegion" if in_plot_area => in_plot_area_region = true,
                    b"plotSurface" if in_plot_area && !in_plot_area_region => {
                        in_plot_surface = true;
                    }
                    b"series" if in_plot_area_region => {
                        in_series = true;
                        ser_layout = ChartExLayout::Unknown("unknown".into());
                        ser_unique_id = None;
                        ser_hidden = None;
                        ser_owner_idx = None;
                        ser_format_idx = None;
                        ser_text = None;
                        ser_data_id = 0;
                        ser_data_labels = None;
                        ser_data_points.clear();
                        ser_layout_pr = None;
                        ser_axis_ids.clear();
                        ser_value_colors = None;
                        ser_value_color_positions = None;
                        ser_shape_properties = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"layoutId" => {
                                    if let Ok(v) = attr.unescape_value() {
                                        ser_layout = parse_layout_id(&v);
                                    }
                                }
                                b"uniqueId" => {
                                    ser_unique_id =
                                        attr.unescape_value().ok().map(|s| s.to_string())
                                }
                                b"hidden" => {
                                    ser_hidden = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s.as_ref() == "true")
                                }
                                b"ownerIdx" => {
                                    ser_owner_idx =
                                        attr.unescape_value().ok().and_then(|s| s.parse().ok())
                                }
                                b"formatIdx" => {
                                    ser_format_idx =
                                        attr.unescape_value().ok().and_then(|s| s.parse().ok())
                                }
                                _ => {}
                            }
                        }
                    }
                    b"tx" if in_series && !in_data_labels => {
                        in_ser_tx = true;
                        ser_tx_value = None;
                        ser_tx_formula = None;
                    }
                    b"txData" if in_ser_tx => in_ser_tx_data = true,
                    b"v" if in_ser_tx_data => in_ser_tx_data_v = true,
                    b"f" if in_ser_tx_data => in_ser_tx_data_f = true,
                    b"dataId" if in_series && !in_data_labels => {
                        in_data_id = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                ser_data_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                        }
                    }
                    b"dataLabels" if in_series => {
                        in_data_labels = true;
                        dlbl_pos = None;
                        dlbl_vis_series = None;
                        dlbl_vis_cat = None;
                        dlbl_vis_val = None;
                        dlbl_num_fmt = None;
                        dlbl_separator = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"pos" {
                                dlbl_pos = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"separator" if in_data_labels => in_dlbl_separator = true,
                    b"dataPt" if in_series && !in_data_labels => {
                        in_data_pt = true;
                        dpt_idx = 0;
                        dpt_sp = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"idx" {
                                dpt_idx = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                        }
                    }
                    b"layoutPr" if in_series => {
                        in_layout_pr = true;
                        layout_pr = ChartExLayoutPr::default();
                    }
                    b"subtotals" if in_layout_pr => in_subtotals = true,
                    b"idx" if in_subtotals => in_subtotal_idx = true,
                    b"binning" if in_layout_pr => {
                        in_binning = true;
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
                        layout_pr.binning = Some(binning);
                    }
                    b"binSize" if in_binning => in_bin_size = true,
                    b"binCount" if in_binning => in_bin_count = true,
                    b"geography" if in_layout_pr => {
                        in_geography = true;
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
                        layout_pr.geography = Some(geo);
                    }
                    b"geoCache" if in_geography => {
                        _in_geo_cache = true;
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        geo_cache_depth = 1;
                        geo_cache_writer = Some(w);
                    }
                    b"statistics" if in_layout_pr => {
                        in_statistics = true;
                        let mut stats = ChartExStatistics::default();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"quartileMethod" {
                                stats.quartile_method =
                                    attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                        layout_pr.statistics = Some(stats);
                    }
                    b"visibility" if in_layout_pr && !in_data_labels => {
                        in_layout_visibility = true;
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
                        layout_pr.visibility = Some(vis);
                    }
                    b"axisId" if in_series && !in_axis => in_axis_id = true,
                    b"valueColors" if in_series => {
                        in_value_colors = true;
                        ser_value_colors = Some(ChartExValueColors::default());
                    }
                    b"minColor" | b"midColor" | b"maxColor"
                        if in_value_colors && vc_writer.is_none() =>
                    {
                        let tag_name = std::str::from_utf8(tag).unwrap_or("").to_string();
                        vc_capturing_tag = Some(tag_name);
                        let mut w = Writer::new(Cursor::new(Vec::new()));
                        let _ = w.write_event(Event::Start(e.clone().into_owned()));
                        vc_depth = 1;
                        vc_writer = Some(w);
                    }
                    b"valueColorPositions" if in_series => {
                        in_value_color_positions = true;
                        vcp = ChartExValueColorPositions::default();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"count" {
                                vcp.count = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                        }
                    }
                    b"min" if in_value_color_positions => in_vcp_min = true,
                    b"mid" if in_value_color_positions => in_vcp_mid = true,
                    b"max" if in_value_color_positions => in_vcp_max = true,
                    b"axis" if in_plot_area && !in_plot_area_region => {
                        in_axis = true;
                        ax_id = 0;
                        ax_hidden = None;
                        ax_scaling = ChartExScaling::Category { gap_width: None };
                        ax_title = None;
                        ax_units = None;
                        ax_major_gridlines = None;
                        ax_minor_gridlines = None;
                        ax_major_tick_marks = None;
                        ax_minor_tick_marks = None;
                        ax_tick_labels = false;
                        ax_num_fmt = None;
                        ax_shape_properties = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"id" => {
                                    ax_id = attr
                                        .unescape_value()
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0)
                                }
                                b"hidden" => {
                                    ax_hidden = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s.as_ref() == "true")
                                }
                                _ => {}
                            }
                        }
                    }
                    b"catScaling" if in_axis => {
                        in_cat_scaling = true;
                        let mut gw = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"gapWidth" {
                                gw = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                        }
                        ax_scaling = ChartExScaling::Category { gap_width: gw };
                    }
                    b"valScaling" if in_axis => {
                        in_val_scaling = true;
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
                        ax_scaling = ChartExScaling::Value {
                            min,
                            max,
                            major_unit: major,
                            minor_unit: minor,
                        };
                    }
                    b"title" if in_axis => {
                        in_ax_title = true;
                        ax_title_text = None;
                    }
                    b"tx" if in_ax_title => in_ax_title_tx = true,
                    b"txData" if in_ax_title_tx => in_ax_title_tx_data = true,
                    b"v" if in_ax_title_tx_data => in_ax_title_tx_data_v = true,
                    b"units" if in_axis => {
                        in_units = true;
                        units_unit = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"unit" {
                                units_unit = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"majorGridlines" if in_axis => in_major_gridlines = true,
                    b"minorGridlines" if in_axis => in_minor_gridlines = true,
                    b"legend" if in_chart && !in_plot_area => {
                        in_legend = true;
                        legend_pos = None;
                        legend_align = None;
                        legend_overlay = None;
                        legend_sp = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"pos" => {
                                    legend_pos = attr.unescape_value().ok().map(|s| s.to_string())
                                }
                                b"align" => {
                                    legend_align = attr.unescape_value().ok().map(|s| s.to_string())
                                }
                                b"overlay" => {
                                    legend_overlay = attr
                                        .unescape_value()
                                        .ok()
                                        .map(|s| s == "1" || s.as_ref() == "true")
                                }
                                _ => {}
                            }
                        }
                    }
                    b"printSettings" if in_chart_space => {
                        in_print_settings = true;
                        print_settings = ChartExPrintSettings::default();
                    }
                    b"headerFooter" if in_print_settings => {
                        in_ps_header_footer = true;
                        ps_hf = ChartExHeaderFooter::default();
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"alignWithMargins" => {
                                    ps_hf.align_with_margins = parse_bool_attr(&attr)
                                }
                                b"differentOddEven" => {
                                    ps_hf.different_odd_even = parse_bool_attr(&attr)
                                }
                                b"differentFirst" => ps_hf.different_first = parse_bool_attr(&attr),
                                _ => {}
                            }
                        }
                    }
                    b"oddHeader" if in_ps_header_footer => in_ps_hf_odd_header = true,
                    b"oddFooter" if in_ps_header_footer => in_ps_hf_odd_footer = true,
                    b"evenHeader" if in_ps_header_footer => in_ps_hf_even_header = true,
                    b"evenFooter" if in_ps_header_footer => in_ps_hf_even_footer = true,
                    b"firstHeader" if in_ps_header_footer => in_ps_hf_first_header = true,
                    b"firstFooter" if in_ps_header_footer => in_ps_hf_first_footer = true,
                    b"spPr" if !in_sp_pr && !skipping => {
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
                        if in_data_pt {
                            sp_ctx = SpCtx::DataPt;
                        } else if in_data_labels {
                            sp_ctx = SpCtx::DataLabels;
                        } else if in_series {
                            sp_ctx = SpCtx::Series;
                        } else if in_major_gridlines {
                            sp_ctx = SpCtx::MajorGrid;
                        } else if in_minor_gridlines {
                            sp_ctx = SpCtx::MinorGrid;
                        } else if in_axis {
                            sp_ctx = SpCtx::Axis;
                        } else if in_legend {
                            sp_ctx = SpCtx::Legend;
                        } else if in_plot_surface {
                            sp_ctx = SpCtx::PlotSurface;
                        } else if in_chart_space && !in_chart {
                            sp_ctx = SpCtx::ChartSpace;
                        } else {
                            sp_ctx = SpCtx::None;
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
                                sp_ln_width =
                                    attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                        }
                    }
                    b"solidFill" if in_sp_pr => sp_pr_depth += 1,
                    b"txPr" if !skipping => {
                        skipping = true;
                        skip_depth = 1;
                    }
                    b"rich" if !skipping => {
                        skipping = true;
                        skip_depth = 1;
                    }
                    b"clrMapOvr" if !skipping => {
                        skipping = true;
                        skip_depth = 1;
                    }
                    b"fmtOvrs" if !skipping => {
                        skipping = true;
                        skip_depth = 1;
                    }
                    b"extLst" if !skipping => {
                        skipping = true;
                        skip_depth = 1;
                    }
                    _ => {
                        if in_sp_pr {
                            sp_pr_depth += 1;
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if let Some(ref mut w) = geo_cache_writer {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                    buf.clear();
                    continue;
                }
                if let Some(ref mut w) = vc_writer {
                    let _ = w.write_event(Event::Empty(e.clone().into_owned()));
                    buf.clear();
                    continue;
                }
                if skipping {
                    buf.clear();
                    continue;
                }

                let local = e.name().local_name();
                let tag = local.as_ref();
                match tag {
                    b"externalData" if in_chart_data => {
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
                        result.external_data = Some(ChartExExternalData {
                            rel_id,
                            auto_update,
                        });
                    }
                    b"dataId" if in_series && !in_data_labels => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                ser_data_id = attr
                                    .unescape_value()
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                        }
                    }
                    b"visibility" if in_data_labels => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"seriesName" => dlbl_vis_series = parse_bool_attr(&attr),
                                b"categoryName" => dlbl_vis_cat = parse_bool_attr(&attr),
                                b"value" => dlbl_vis_val = parse_bool_attr(&attr),
                                _ => {}
                            }
                        }
                    }
                    b"numFmt" if in_data_labels => dlbl_num_fmt = Some(parse_num_fmt(e)),
                    b"numFmt" if in_axis && !in_data_labels => {
                        ax_num_fmt = Some(parse_num_fmt(e));
                    }
                    b"visibility" if in_layout_pr && !in_data_labels => {
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
                        layout_pr.visibility = Some(vis);
                    }
                    b"parentLabelLayout" if in_layout_pr => {
                        layout_pr.parent_label_layout = get_val_attr(e);
                    }
                    b"regionLabelLayout" if in_layout_pr => {
                        layout_pr.region_label_layout = get_val_attr(e);
                    }
                    b"aggregation" if in_layout_pr => layout_pr.aggregation = true,
                    b"statistics" if in_layout_pr => {
                        let mut stats = ChartExStatistics::default();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"quartileMethod" {
                                stats.quartile_method =
                                    attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                        layout_pr.statistics = Some(stats);
                    }
                    b"catScaling" if in_axis => {
                        let mut gw = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"gapWidth" {
                                gw = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                        }
                        ax_scaling = ChartExScaling::Category { gap_width: gw };
                    }
                    b"valScaling" if in_axis => {
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
                        ax_scaling = ChartExScaling::Value {
                            min,
                            max,
                            major_unit: major,
                            minor_unit: minor,
                        };
                    }
                    b"majorTickMarks" if in_axis => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"type" {
                                ax_major_tick_marks =
                                    attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"minorTickMarks" if in_axis => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"type" {
                                ax_minor_tick_marks =
                                    attr.unescape_value().ok().map(|s| s.to_string());
                            }
                        }
                    }
                    b"tickLabels" if in_axis => ax_tick_labels = true,
                    b"majorGridlines" if in_axis => {
                        ax_major_gridlines = Some(ChartShapeProperties::default());
                    }
                    b"minorGridlines" if in_axis => {
                        ax_minor_gridlines = Some(ChartShapeProperties::default());
                    }
                    b"srgbClr" if in_sp_pr && !in_sp_ln => {
                        if let Some(hex) = get_val_attr(e) {
                            sp_solid_fill = Some(ChartColor { hex });
                        }
                    }
                    b"srgbClr" if in_sp_ln => {
                        if let Some(hex) = get_val_attr(e) {
                            sp_ln_solid_fill = Some(ChartColor { hex });
                        }
                    }
                    b"noFill" if in_sp_pr && !in_sp_ln => sp_no_fill = true,
                    b"noFill" if in_sp_ln => sp_ln_no_fill = true,
                    b"prstDash" if in_sp_ln => sp_ln_dash = get_val_attr(e),
                    b"pageMargins" if in_print_settings => {
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
                        print_settings.page_margins = Some(pm);
                    }
                    b"pageSetup" if in_print_settings => {
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
                        print_settings.page_setup = Some(ps);
                    }
                    b"extremeValue" if in_vcp_min => {
                        vcp.min = Some(ChartExColorPosition::ExtremeValue)
                    }
                    b"extremeValue" if in_vcp_mid => {
                        vcp.mid = Some(ChartExColorPosition::ExtremeValue)
                    }
                    b"extremeValue" if in_vcp_max => {
                        vcp.max = Some(ChartExColorPosition::ExtremeValue)
                    }
                    b"number" if in_vcp_min => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.min = Some(ChartExColorPosition::Number(v));
                        }
                    }
                    b"number" if in_vcp_mid => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.mid = Some(ChartExColorPosition::Number(v));
                        }
                    }
                    b"number" if in_vcp_max => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.max = Some(ChartExColorPosition::Number(v));
                        }
                    }
                    b"percent" if in_vcp_min => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.min = Some(ChartExColorPosition::Percent(v));
                        }
                    }
                    b"percent" if in_vcp_mid => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.mid = Some(ChartExColorPosition::Percent(v));
                        }
                    }
                    b"percent" if in_vcp_max => {
                        if let Some(v) = get_val_f64(e) {
                            vcp.max = Some(ChartExColorPosition::Percent(v));
                        }
                    }
                    b"binSize" if in_binning => {
                        if let Some(ref mut b) = layout_pr.binning {
                            b.bin_size = get_val_f64(e);
                        }
                    }
                    b"binCount" if in_binning => {
                        if let Some(ref mut b) = layout_pr.binning {
                            b.bin_count = get_val_attr(e).and_then(|s| s.parse().ok());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Some(ref mut w) = geo_cache_writer {
                    let _ = w.write_event(Event::Text(e.clone().into_owned()));
                    buf.clear();
                    continue;
                }
                if let Some(ref mut w) = vc_writer {
                    let _ = w.write_event(Event::Text(e.clone().into_owned()));
                    buf.clear();
                    continue;
                }
                if skipping {
                    buf.clear();
                    continue;
                }
                if let Ok(text) = e.unescape() {
                    let t = text.as_ref();
                    if in_dim_f {
                        dim_formula = Some(t.to_string());
                    } else if in_dim_nf {
                        dim_nf = Some(t.to_string());
                    } else if in_lvl_pt_text && in_lvl_pt {
                        if in_str_dim {
                            lvl_str_points.push((lvl_pt_idx, t.to_string()));
                        } else if in_num_dim {
                            lvl_num_points.push((lvl_pt_idx, t.to_string()));
                        }
                    } else if in_title_tx_data_v {
                        title_text = Some(t.to_string());
                    } else if in_title_tx_data_f {
                        title_text = Some(t.to_string());
                    } else if in_ser_tx_data_v {
                        ser_tx_value = Some(t.to_string());
                    } else if in_ser_tx_data_f {
                        ser_tx_formula = Some(t.to_string());
                    } else if in_axis_id {
                        if let Ok(id) = t.parse::<u32>() {
                            ser_axis_ids.push(id);
                        }
                    } else if in_ax_title_tx_data_v {
                        ax_title_text = Some(t.to_string());
                    } else if in_subtotal_idx {
                        if let Ok(idx) = t.parse::<u32>() {
                            layout_pr.subtotals.push(idx);
                        }
                    } else if in_bin_size {
                        if let Some(ref mut b) = layout_pr.binning {
                            b.bin_size = t.parse().ok();
                        }
                    } else if in_bin_count {
                        if let Some(ref mut b) = layout_pr.binning {
                            b.bin_count = t.parse().ok();
                        }
                    } else if in_dlbl_separator {
                        dlbl_separator = Some(t.to_string());
                    } else if in_ps_hf_odd_header {
                        ps_hf.odd_header = Some(t.to_string());
                    } else if in_ps_hf_odd_footer {
                        ps_hf.odd_footer = Some(t.to_string());
                    } else if in_ps_hf_even_header {
                        ps_hf.even_header = Some(t.to_string());
                    } else if in_ps_hf_even_footer {
                        ps_hf.even_footer = Some(t.to_string());
                    } else if in_ps_hf_first_header {
                        ps_hf.first_header = Some(t.to_string());
                    } else if in_ps_hf_first_footer {
                        ps_hf.first_footer = Some(t.to_string());
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                // geoCache raw capture
                if let Some(ref mut w) = geo_cache_writer {
                    geo_cache_depth -= 1;
                    let _ = w.write_event(Event::End(e.clone().into_owned()));
                    if geo_cache_depth == 0 {
                        if let Some(w) = geo_cache_writer.take() {
                            if let Some(ref mut geo) = layout_pr.geography {
                                geo.raw_geo_cache = Some(w.into_inner().into_inner());
                            }
                        }
                        _in_geo_cache = false;
                    }
                    buf.clear();
                    continue;
                }
                // valueColor capture
                if let Some(ref mut w) = vc_writer {
                    vc_depth -= 1;
                    let _ = w.write_event(Event::End(e.clone().into_owned()));
                    if vc_depth == 0 {
                        if let Some(w) = vc_writer.take() {
                            let bytes = w.into_inner().into_inner();
                            if let Some(ref mut vc) = ser_value_colors {
                                match vc_capturing_tag.as_deref() {
                                    Some("minColor") => vc.min_color = Some(bytes),
                                    Some("midColor") => vc.mid_color = Some(bytes),
                                    Some("maxColor") => vc.max_color = Some(bytes),
                                    _ => {}
                                }
                            }
                            vc_capturing_tag = None;
                        }
                    }
                    buf.clear();
                    continue;
                }
                if skipping {
                    skip_depth -= 1;
                    if skip_depth == 0 {
                        skipping = false;
                    }
                    buf.clear();
                    continue;
                }

                let local = e.name().local_name();
                let tag = local.as_ref();
                match tag {
                    b"chartSpace" => in_chart_space = false,
                    b"chartData" => in_chart_data = false,
                    b"data" if in_data => {
                        result.data.push(ChartExData {
                            id: data_id,
                            dimensions: std::mem::take(&mut data_dims),
                            extensions: None,
                        });
                        in_data = false;
                    }
                    b"strDim" if in_str_dim => {
                        data_dims.push(ChartExDimension::String {
                            dim_type: dim_str_type.clone(),
                            formula: dim_formula.take(),
                            nf_formula: dim_nf.take(),
                            levels: std::mem::take(&mut dim_levels),
                        });
                        in_str_dim = false;
                    }
                    b"numDim" if in_num_dim => {
                        data_dims.push(ChartExDimension::Numeric {
                            dim_type: dim_num_type.clone(),
                            formula: dim_formula.take(),
                            nf_formula: dim_nf.take(),
                            levels: std::mem::take(&mut dim_num_levels),
                        });
                        in_num_dim = false;
                    }
                    b"f" if in_dim_f => in_dim_f = false,
                    b"nf" if in_dim_nf => in_dim_nf = false,
                    b"lvl" if in_lvl => {
                        if in_str_dim {
                            dim_levels.push(ChartExStringLevel {
                                pt_count: lvl_pt_count,
                                name: lvl_name.take(),
                                points: std::mem::take(&mut lvl_str_points),
                            });
                        } else if in_num_dim {
                            dim_num_levels.push(ChartExNumericLevel {
                                pt_count: lvl_pt_count,
                                format_code: lvl_format_code.take(),
                                name: lvl_name.take(),
                                points: std::mem::take(&mut lvl_num_points),
                            });
                        }
                        in_lvl = false;
                    }
                    b"pt" if in_lvl_pt => {
                        in_lvl_pt = false;
                        in_lvl_pt_text = false;
                    }
                    b"chart" if in_chart => in_chart = false,
                    b"title" if in_ax_title => {
                        let title = ChartExAxisTitle {
                            text: ax_title_text.take().map(|t| ChartExText {
                                data: Some(ChartExTextData {
                                    formula: None,
                                    value: Some(t),
                                }),
                                rich: None,
                            }),
                            offset: None,
                            shape_properties: None,
                            text_properties: None,
                            extensions: None,
                        };
                        ax_title = Some(title);
                        in_ax_title = false;
                        in_ax_title_tx = false;
                        in_ax_title_tx_data = false;
                        in_ax_title_tx_data_v = false;
                    }
                    b"title" if in_title => {
                        result.title = Some(ChartExTitle {
                            text: title_text.take(),
                            rich_text: None,
                            position: title_pos.take(),
                            align: title_align.take(),
                            overlay: title_overlay.take(),
                            offset: None,
                            shape_properties: None,
                            text_properties: None,
                            extensions: None,
                        });
                        in_title = false;
                        in_title_tx = false;
                        in_title_tx_data = false;
                    }
                    b"tx" if in_title_tx && !in_series => in_title_tx = false,
                    b"txData" if in_title_tx_data && !in_series => in_title_tx_data = false,
                    b"v" if in_title_tx_data_v && !in_series => in_title_tx_data_v = false,
                    b"f" if in_title_tx_data_f && !in_series => in_title_tx_data_f = false,
                    b"plotArea" => in_plot_area = false,
                    b"plotAreaRegion" => in_plot_area_region = false,
                    b"plotSurface" if in_plot_surface => in_plot_surface = false,
                    b"series" if in_series => {
                        let series = ChartExSeries {
                            layout: ser_layout.clone(),
                            unique_id: ser_unique_id.take(),
                            hidden: ser_hidden.take(),
                            owner_idx: ser_owner_idx.take(),
                            format_idx: ser_format_idx.take(),
                            text: ser_text.take(),
                            data_id: ser_data_id,
                            data_labels: ser_data_labels.take(),
                            data_points: std::mem::take(&mut ser_data_points),
                            layout_properties: ser_layout_pr.take(),
                            axis_ids: std::mem::take(&mut ser_axis_ids),
                            value_colors: ser_value_colors.take(),
                            value_color_positions: ser_value_color_positions.take(),
                            shape_properties: ser_shape_properties.take(),
                            extensions: None,
                        };
                        result.plot_area.series.push(series);
                        in_series = false;
                    }
                    b"tx" if in_ser_tx => {
                        ser_text = Some(ChartExText {
                            data: Some(ChartExTextData {
                                formula: ser_tx_formula.take(),
                                value: ser_tx_value.take(),
                            }),
                            rich: None,
                        });
                        in_ser_tx = false;
                        in_ser_tx_data = false;
                    }
                    b"txData" if in_ser_tx_data => in_ser_tx_data = false,
                    b"v" if in_ser_tx_data_v => in_ser_tx_data_v = false,
                    b"f" if in_ser_tx_data_f => in_ser_tx_data_f = false,
                    b"dataId" if in_data_id => in_data_id = false,
                    b"dataLabels" if in_data_labels => {
                        ser_data_labels = Some(ChartExDataLabels {
                            position: dlbl_pos.take(),
                            visibility_series_name: dlbl_vis_series.take(),
                            visibility_category_name: dlbl_vis_cat.take(),
                            visibility_value: dlbl_vis_val.take(),
                            number_format: dlbl_num_fmt.take(),
                            separator: dlbl_separator.take(),
                            shape_properties: None,
                            text_properties: None,
                            overrides: Vec::new(),
                            hidden_labels: Vec::new(),
                            extensions: None,
                        });
                        in_data_labels = false;
                    }
                    b"separator" if in_dlbl_separator => in_dlbl_separator = false,
                    b"dataPt" if in_data_pt => {
                        ser_data_points.push(ChartExDataPoint {
                            idx: dpt_idx,
                            shape_properties: dpt_sp.take(),
                            extensions: None,
                        });
                        in_data_pt = false;
                    }
                    b"layoutPr" if in_layout_pr => {
                        ser_layout_pr = Some(std::mem::take(&mut layout_pr));
                        in_layout_pr = false;
                    }
                    b"subtotals" if in_subtotals => in_subtotals = false,
                    b"idx" if in_subtotal_idx => in_subtotal_idx = false,
                    b"binning" if in_binning => {
                        in_binning = false;
                        in_bin_size = false;
                        in_bin_count = false;
                    }
                    b"binSize" if in_bin_size => in_bin_size = false,
                    b"binCount" if in_bin_count => in_bin_count = false,
                    b"geography" if in_geography => in_geography = false,
                    b"statistics" if in_statistics => in_statistics = false,
                    b"visibility" if in_layout_visibility => in_layout_visibility = false,
                    b"axisId" if in_axis_id => in_axis_id = false,
                    b"valueColors" if in_value_colors => in_value_colors = false,
                    b"valueColorPositions" if in_value_color_positions => {
                        ser_value_color_positions = Some(vcp.clone());
                        in_value_color_positions = false;
                        in_vcp_min = false;
                        in_vcp_mid = false;
                        in_vcp_max = false;
                    }
                    b"min" if in_vcp_min => in_vcp_min = false,
                    b"mid" if in_vcp_mid => in_vcp_mid = false,
                    b"max" if in_vcp_max => in_vcp_max = false,
                    b"axis" if in_axis => {
                        result.plot_area.axes.push(ChartExAxis {
                            id: ax_id,
                            hidden: ax_hidden.take(),
                            scaling: ax_scaling.clone(),
                            title: ax_title.take(),
                            units: ax_units.take(),
                            major_gridlines: ax_major_gridlines.take(),
                            minor_gridlines: ax_minor_gridlines.take(),
                            major_tick_marks: ax_major_tick_marks.take(),
                            minor_tick_marks: ax_minor_tick_marks.take(),
                            tick_labels: ax_tick_labels,
                            number_format: ax_num_fmt.take(),
                            shape_properties: ax_shape_properties.take(),
                            text_properties: None,
                            extensions: None,
                        });
                        in_axis = false;
                        in_cat_scaling = false;
                        in_val_scaling = false;
                    }
                    b"catScaling" if in_cat_scaling => in_cat_scaling = false,
                    b"valScaling" if in_val_scaling => in_val_scaling = false,
                    b"tx" if in_ax_title_tx => in_ax_title_tx = false,
                    b"txData" if in_ax_title_tx_data => in_ax_title_tx_data = false,
                    b"v" if in_ax_title_tx_data_v => in_ax_title_tx_data_v = false,
                    b"units" if in_units => {
                        ax_units = Some(ChartExAxisUnits {
                            unit: units_unit.take(),
                            label: None,
                            extensions: None,
                        });
                        in_units = false;
                    }
                    b"majorGridlines" if in_major_gridlines => in_major_gridlines = false,
                    b"minorGridlines" if in_minor_gridlines => in_minor_gridlines = false,
                    b"legend" if in_legend => {
                        result.legend = Some(ChartExLegend {
                            position: legend_pos.take(),
                            align: legend_align.take(),
                            overlay: legend_overlay.take(),
                            offset: None,
                            shape_properties: legend_sp.take(),
                            text_properties: None,
                            extensions: None,
                        });
                        in_legend = false;
                    }
                    b"printSettings" if in_print_settings => {
                        result.print_settings = Some(print_settings.clone());
                        in_print_settings = false;
                    }
                    b"headerFooter" if in_ps_header_footer => {
                        print_settings.header_footer = Some(ps_hf.clone());
                        in_ps_header_footer = false;
                    }
                    b"oddHeader" => in_ps_hf_odd_header = false,
                    b"oddFooter" => in_ps_hf_odd_footer = false,
                    b"evenHeader" => in_ps_hf_even_header = false,
                    b"evenFooter" => in_ps_hf_even_footer = false,
                    b"firstHeader" => in_ps_hf_first_header = false,
                    b"firstFooter" => in_ps_hf_first_footer = false,
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
                            let has =
                                props.solid_fill.is_some() || props.no_fill || props.line.is_some();
                            match sp_ctx {
                                SpCtx::ChartSpace => {
                                    if has {
                                        result.shape_properties = Some(props);
                                    }
                                }
                                SpCtx::Series => {
                                    if has {
                                        ser_shape_properties = Some(props);
                                    }
                                }
                                SpCtx::DataPt => {
                                    dpt_sp = if has { Some(props) } else { None };
                                }
                                SpCtx::DataLabels => {
                                    // stored on data labels completion
                                }
                                SpCtx::Legend => {
                                    if has {
                                        legend_sp = Some(props);
                                    }
                                }
                                SpCtx::Axis => {
                                    if has {
                                        ax_shape_properties = Some(props);
                                    }
                                }
                                SpCtx::MajorGrid => {
                                    ax_major_gridlines = if has {
                                        Some(props)
                                    } else {
                                        Some(ChartShapeProperties::default())
                                    };
                                }
                                SpCtx::MinorGrid => {
                                    ax_minor_gridlines = if has {
                                        Some(props)
                                    } else {
                                        Some(ChartShapeProperties::default())
                                    };
                                }
                                SpCtx::PlotSurface => {
                                    if has {
                                        result.plot_area.plot_surface = Some(props);
                                    }
                                }
                                SpCtx::None => {}
                            }
                            in_sp_pr = false;
                            sp_no_fill = false;
                            sp_ctx = SpCtx::None;
                        }
                    }
                    _ => {
                        if in_sp_pr {
                            sp_pr_depth = sp_pr_depth.saturating_sub(1);
                            if sp_pr_depth == 0 {
                                in_sp_pr = false;
                                sp_no_fill = false;
                                sp_ctx = SpCtx::None;
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

    Ok(result)
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
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"val" {
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
