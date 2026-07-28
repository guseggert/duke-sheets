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

#[derive(Debug, Default)]
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

/// Where a captured subtree's bytes belong.
#[derive(Debug, Clone, Copy)]
enum CaptureDest {
    /// `cx:geoCache` of the series' `cx:geography`.
    GeoCache,
    /// One colour slot of the series' `cx:valueColors`.
    ValueColor(ValueColorSlot),
}

#[derive(Debug, Clone, Copy)]
enum ValueColorSlot {
    Min,
    Mid,
    Max,
}

/// A subtree the parser does not read element by element.
///
/// Both kinds behave identically as far as the event loop is concerned -
/// swallow everything until the matching end - so they share one
/// mechanism. Keeping them separate is what let a captured subtree drop
/// comments and CDATA that a skipped one never had to care about, and
/// let a capture forget the depth bookkeeping a skip did.
enum Opaque {
    /// Dropped: an element we do not model.
    Skip { depth: u32 },
    /// Kept verbatim, opener included, for replay on write.
    Capture {
        dest: CaptureDest,
        depth: u32,
        writer: Writer<Cursor<Vec<u8>>>,
    },
}

/// What a `cx:spPr` subtree currently belongs to. The element itself is
/// identical wherever it appears, so its owner has to be tracked.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum SpCtx {
    #[default]
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
/// Where in the document tree the cursor sits.
#[derive(Debug, Default)]
struct DocPos {
    chart_space: bool,
    chart_data: bool,
    chart: bool,
    plot_area: bool,
    plot_area_region: bool,
    plot_surface: bool,
}

/// One `cx:data` block and the dimension or level being read.
#[derive(Debug)]
struct DataState {
    open: bool,
    id: u32,
    dims: Vec<ChartExDimension>,
    in_str_dim: bool,
    in_num_dim: bool,
    str_type: StringDimType,
    num_type: NumericDimType,
    formula: Option<String>,
    nf: Option<String>,
    in_f: bool,
    in_nf: bool,
    str_levels: Vec<ChartExStringLevel>,
    num_levels: Vec<ChartExNumericLevel>,
    in_lvl: bool,
    lvl_pt_count: u32,
    lvl_name: Option<String>,
    lvl_format_code: Option<String>,
    lvl_str_points: Vec<(u32, String)>,
    lvl_num_points: Vec<(u32, String)>,
    in_lvl_pt: bool,
    lvl_pt_idx: u32,
    in_lvl_pt_text: bool,
}

impl Default for DataState {
    fn default() -> Self {
        Self {
            open: false,
            id: 0,
            dims: Vec::new(),
            in_str_dim: false,
            in_num_dim: false,
            str_type: StringDimType::Cat,
            num_type: NumericDimType::Val,
            formula: None,
            nf: None,
            in_f: false,
            in_nf: false,
            str_levels: Vec::new(),
            num_levels: Vec::new(),
            in_lvl: false,
            lvl_pt_count: 0,
            lvl_name: None,
            lvl_format_code: None,
            lvl_str_points: Vec::new(),
            lvl_num_points: Vec::new(),
            in_lvl_pt: false,
            lvl_pt_idx: 0,
            in_lvl_pt_text: false,
        }
    }
}

/// The chart-level `cx:title`.
#[derive(Debug, Default)]
struct TitleState {
    open: bool,
    pos: Option<String>,
    align: Option<String>,
    overlay: Option<bool>,
    in_tx: bool,
    in_tx_data: bool,
    in_tx_data_v: bool,
    in_tx_data_f: bool,
    text: Option<String>,
    sp: Option<ChartShapeProperties>,
    offset: Option<ChartExOffset>,
}

/// The `cx:series` being read.
#[derive(Debug)]
struct SeriesState {
    open: bool,
    layout: ChartExLayout,
    unique_id: Option<String>,
    hidden: Option<bool>,
    owner_idx: Option<u32>,
    format_idx: Option<u32>,
    text: Option<ChartExText>,
    data_id: u32,
    data_labels: Option<ChartExDataLabels>,
    data_points: Vec<ChartExDataPoint>,
    layout_pr: Option<ChartExLayoutPr>,
    axis_ids: Vec<u32>,
    value_colors: Option<ChartExValueColors>,
    value_color_positions: Option<ChartExValueColorPositions>,
    shape_properties: Option<ChartShapeProperties>,
    in_tx: bool,
    in_tx_data: bool,
    in_tx_data_v: bool,
    in_tx_data_f: bool,
    tx_value: Option<String>,
    tx_formula: Option<String>,
    in_data_id: bool,
    in_axis_id: bool,
    in_value_colors: bool,
}

impl Default for SeriesState {
    fn default() -> Self {
        Self {
            open: false,
            layout: ChartExLayout::Unknown("unknown".into()),
            unique_id: None,
            hidden: None,
            owner_idx: None,
            format_idx: None,
            text: None,
            data_id: 0,
            data_labels: None,
            data_points: Vec::new(),
            layout_pr: None,
            axis_ids: Vec::new(),
            value_colors: None,
            value_color_positions: None,
            shape_properties: None,
            in_tx: false,
            in_tx_data: false,
            in_tx_data_v: false,
            in_tx_data_f: false,
            tx_value: None,
            tx_formula: None,
            in_data_id: false,
            in_axis_id: false,
            in_value_colors: false,
        }
    }
}

/// Series-level `cx:dataLabels`.
#[derive(Debug, Default)]
struct DataLabelsState {
    open: bool,
    pos: Option<String>,
    vis_series: Option<bool>,
    vis_cat: Option<bool>,
    vis_val: Option<bool>,
    num_fmt: Option<NumberFormat>,
    separator: Option<String>,
    sp: Option<ChartShapeProperties>,
    in_separator: bool,
    overrides: Vec<ChartExDataLabel>,
    hidden: Vec<u32>,
}

/// A per-point `cx:dataLabel` override.
#[derive(Debug, Default)]
struct LabelOverrideState {
    open: bool,
    idx: u32,
    pos: Option<String>,
    vis_series: Option<bool>,
    vis_cat: Option<bool>,
    vis_val: Option<bool>,
    num_fmt: Option<NumberFormat>,
    separator: Option<String>,
    sp: Option<ChartShapeProperties>,
    in_separator: bool,
}

/// A `cx:dataPt`.
#[derive(Debug, Default)]
struct DataPointState {
    open: bool,
    idx: u32,
    sp: Option<ChartShapeProperties>,
}

/// `cx:layoutPr` and the layout-specific child being read.
#[derive(Debug, Default)]
struct LayoutPrState {
    open: bool,
    pr: ChartExLayoutPr,
    in_subtotals: bool,
    in_binning: bool,
    in_geography: bool,
    in_statistics: bool,
    in_visibility: bool,
}

/// `cx:valueColorPositions`.
#[derive(Debug, Default)]
struct ColorPositionsState {
    open: bool,
    value: ChartExValueColorPositions,
    in_min: bool,
    in_mid: bool,
    in_max: bool,
}

/// The `cx:axis` being read, including its title and units.
#[derive(Debug)]
struct AxisState {
    open: bool,
    id: u32,
    hidden: Option<bool>,
    scaling: ChartExScaling,
    title: Option<ChartExAxisTitle>,
    units: Option<ChartExAxisUnits>,
    major_gridlines: Option<ChartExGridlines>,
    minor_gridlines: Option<ChartExGridlines>,
    major_tick_marks: Option<String>,
    minor_tick_marks: Option<String>,
    tick_labels: bool,
    num_fmt: Option<NumberFormat>,
    shape_properties: Option<ChartShapeProperties>,
    in_cat_scaling: bool,
    in_val_scaling: bool,
    in_title: bool,
    in_title_tx: bool,
    in_title_tx_data: bool,
    in_title_tx_data_v: bool,
    title_text: Option<String>,
    title_sp: Option<ChartShapeProperties>,
    title_offset: Option<ChartExOffset>,
    in_units: bool,
    units_unit: Option<String>,
    in_major_gridlines: bool,
    in_minor_gridlines: bool,
}

impl Default for AxisState {
    fn default() -> Self {
        Self {
            open: false,
            id: 0,
            hidden: None,
            scaling: ChartExScaling::Category { gap_width: None },
            title: None,
            units: None,
            major_gridlines: None,
            minor_gridlines: None,
            major_tick_marks: None,
            minor_tick_marks: None,
            tick_labels: false,
            num_fmt: None,
            shape_properties: None,
            in_cat_scaling: false,
            in_val_scaling: false,
            in_title: false,
            in_title_tx: false,
            in_title_tx_data: false,
            in_title_tx_data_v: false,
            title_text: None,
            title_sp: None,
            title_offset: None,
            in_units: false,
            units_unit: None,
            in_major_gridlines: false,
            in_minor_gridlines: false,
        }
    }
}

/// `cx:legend`.
#[derive(Debug, Default)]
struct LegendState {
    open: bool,
    pos: Option<String>,
    align: Option<String>,
    overlay: Option<bool>,
    sp: Option<ChartShapeProperties>,
    offset: Option<ChartExOffset>,
}

/// The `cx:spPr` subtree being read and what it belongs to.
#[derive(Debug, Default)]
struct SpPrState {
    open: bool,
    depth: u32,
    solid_fill: Option<ChartColor>,
    no_fill: bool,
    line: Option<ChartLine>,
    in_ln: bool,
    ln_width: Option<i64>,
    ln_solid_fill: Option<ChartColor>,
    ln_no_fill: bool,
    ln_dash: Option<String>,
    ctx: SpCtx,
}

/// `cx:printSettings`.
#[derive(Debug, Default)]
struct PrintSettingsState {
    open: bool,
    settings: ChartExPrintSettings,
    in_header_footer: bool,
    hf: ChartExHeaderFooter,
    in_hf_odd_header: bool,
    in_hf_odd_footer: bool,
    in_hf_even_header: bool,
    in_hf_even_footer: bool,
    in_hf_first_header: bool,
    in_hf_first_footer: bool,
}

/// Streaming state for one chartEx part.
///
/// The document is deep and its elements are context sensitive, so
/// parsing is a state machine rather than a recursive descent. State is
/// grouped by the element it belongs to, so a handler names the region
/// it is working in.
#[derive(Default)]
struct ChartExParser {
    result: ParsedChartEx,
    pos: DocPos,
    data: DataState,
    title: TitleState,
    series: SeriesState,
    labels: DataLabelsState,
    label: LabelOverrideState,
    data_pt: DataPointState,
    layout: LayoutPrState,
    color_pos: ColorPositionsState,
    axis: AxisState,
    legend: LegendState,
    sp: SpPrState,
    print: PrintSettingsState,
    /// The open opaque subtree, if any.
    opaque: Option<Opaque>,
}

impl ChartExParser {
    /// Begin dropping a subtree we do not model.
    fn begin_skip(&mut self) {
        self.opaque = Some(Opaque::Skip { depth: 1 });
    }

    /// Begin keeping a subtree verbatim, opener included.
    fn begin_capture(&mut self, dest: CaptureDest, opener: &BytesStart) {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        let _ = writer.write_event(Event::Start(opener.borrow()));
        self.opaque = Some(Opaque::Capture {
            dest,
            depth: 1,
            writer,
        });
    }

    /// Offer an event to the open opaque subtree. Returns whether it was
    /// swallowed, in which case no handler sees it.
    fn consume_opaque(&mut self, event: &Event) -> bool {
        let Some(region) = self.opaque.as_mut() else {
            return false;
        };
        let depth = match region {
            Opaque::Skip { depth } => depth,
            Opaque::Capture { depth, writer, .. } => {
                // Every event kind, so a capture replays its source
                // exactly - comments and CDATA included.
                let _ = writer.write_event(event.clone());
                depth
            }
        };
        let closed = match event {
            Event::Start(_) => {
                *depth += 1;
                false
            }
            Event::End(_) => {
                *depth -= 1;
                *depth == 0
            }
            _ => false,
        };
        if closed {
            self.end_opaque();
        }
        true
    }

    /// Close the opaque subtree and store what it captured.
    fn end_opaque(&mut self) {
        // The element that opened the subtree was counted by the spPr
        // depth bookkeeping; its subtree ends here. It cannot close the
        // spPr itself, since the spPr's own start put the depth at 1
        // before the opener raised it.
        if self.sp.open {
            self.sp.depth = self.sp.depth.saturating_sub(1);
        }
        let Some(Opaque::Capture { dest, writer, .. }) = self.opaque.take() else {
            return;
        };
        let bytes = writer.into_inner().into_inner();
        match dest {
            CaptureDest::GeoCache => {
                if let Some(geo) = self.layout.pr.geography.as_mut() {
                    geo.raw_geo_cache = Some(bytes);
                }
            }
            CaptureDest::ValueColor(slot) => {
                if let Some(vc) = self.series.value_colors.as_mut() {
                    match slot {
                        ValueColorSlot::Min => vc.min_color = Some(bytes),
                        ValueColorSlot::Mid => vc.mid_color = Some(bytes),
                        ValueColorSlot::Max => vc.max_color = Some(bytes),
                    }
                }
            }
        }
    }

    /// Handle one event from the reader.
    fn handle(&mut self, event: &Event) {
        if let Some(end) = self.dispatch(event) {
            // `<x/>` is `<x></x>`, so a self-closing element is a start
            // followed by an end; the synthesized end cannot itself be
            // self-closing, so this cannot recur.
            self.dispatch(&Event::End(end));
        }
    }

    /// Route one event, returning the end to synthesize when the event
    /// was a self-closing element.
    fn dispatch(&mut self, event: &Event) -> Option<BytesEnd<'static>> {
        if self.consume_opaque(event) {
            return None;
        }
        match event {
            Event::Start(e) => self.on_start(e),
            Event::Empty(e) => {
                let end = BytesEnd::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                self.on_start(e);
                return Some(end);
            }
            Event::Text(e) => self.on_text(e),
            Event::End(e) => self.on_end(e),
            _ => {}
        }
        None
    }

    fn on_start(&mut self, e: &BytesStart) {
        let local = e.name().local_name();
        let tag = local.as_ref();

        // Inside a cx:spPr every element start deepens the
        // nesting and its end unwinds it, including the end
        // synthesized for a self-closing element. Tracked here
        // rather than in individual arms so that adding an arm
        // cannot unbalance it.
        if self.sp.open {
            self.sp.depth += 1;
        }

        match tag {
            b"chartSpace" => {
                self.pos.chart_space = true;
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
            b"chartData" if self.pos.chart_space => self.pos.chart_data = true,
            b"data" if self.pos.chart_data => {
                self.data.open = true;
                self.data.id = 0;
                self.data.dims.clear();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"id" {
                        self.data.id = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"strDim" if self.data.open => {
                self.data.in_str_dim = true;
                self.data.formula = None;
                self.data.nf = None;
                self.data.str_levels.clear();
                self.data.str_type = StringDimType::Cat;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        if let Ok(v) = attr.unescape_value() {
                            self.data.str_type = match v.as_ref() {
                                "colorStr" => StringDimType::ColorStr,
                                "entityId" => StringDimType::EntityId,
                                _ => StringDimType::Cat,
                            };
                        }
                    }
                }
            }
            b"numDim" if self.data.open => {
                self.data.in_num_dim = true;
                self.data.formula = None;
                self.data.nf = None;
                self.data.num_levels.clear();
                self.data.num_type = NumericDimType::Val;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        if let Ok(v) = attr.unescape_value() {
                            self.data.num_type = match v.as_ref() {
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
            b"f" if (self.data.in_str_dim || self.data.in_num_dim) && !self.data.in_lvl => self.data.in_f = true,
            b"nf" if (self.data.in_str_dim || self.data.in_num_dim) && !self.data.in_lvl => self.data.in_nf = true,
            b"lvl" if self.data.in_str_dim || self.data.in_num_dim => {
                self.data.in_lvl = true;
                self.data.lvl_pt_count = 0;
                self.data.lvl_name = None;
                self.data.lvl_format_code = None;
                self.data.lvl_str_points.clear();
                self.data.lvl_num_points.clear();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"ptCount" => {
                            self.data.lvl_pt_count = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"name" => {
                            self.data.lvl_name = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        b"formatCode" => {
                            self.data.lvl_format_code =
                                attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        _ => {}
                    }
                }
            }
            b"pt" if self.data.in_lvl => {
                self.data.in_lvl_pt = true;
                self.data.lvl_pt_idx = 0;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"idx" {
                        self.data.lvl_pt_idx = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
                self.data.in_lvl_pt_text = true;
            }
            b"chart" if self.pos.chart_space && !self.pos.chart => self.pos.chart = true,
            b"title" if self.pos.chart && !self.pos.plot_area && !self.axis.open && !self.title.open => {
                self.title.open = true;
                self.title.pos = None;
                self.title.align = None;
                self.title.overlay = None;
                self.title.text = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"pos" => {
                            self.title.pos = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"align" => {
                            self.title.align = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"overlay" => {
                            self.title.overlay = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"tx" if self.title.open && !self.series.open => self.title.in_tx = true,
            b"txData" if self.title.in_tx && !self.series.open => self.title.in_tx_data = true,
            b"v" if self.title.in_tx_data && !self.series.open => self.title.in_tx_data_v = true,
            b"f" if self.title.in_tx_data && !self.series.open => self.title.in_tx_data_f = true,
            b"plotArea" if self.pos.chart => self.pos.plot_area = true,
            b"plotAreaRegion" if self.pos.plot_area => self.pos.plot_area_region = true,
            b"plotSurface" if self.pos.plot_area && !self.pos.plot_area_region => {
                self.pos.plot_surface = true;
            }
            b"series" if self.pos.plot_area_region => {
                self.series.open = true;
                self.series.layout = ChartExLayout::Unknown("unknown".into());
                self.series.unique_id = None;
                self.series.hidden = None;
                self.series.owner_idx = None;
                self.series.format_idx = None;
                self.series.text = None;
                self.series.data_id = 0;
                self.series.data_labels = None;
                self.series.data_points.clear();
                self.series.layout_pr = None;
                self.series.axis_ids.clear();
                self.series.value_colors = None;
                self.series.value_color_positions = None;
                self.series.shape_properties = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"layoutId" => {
                            if let Ok(v) = attr.unescape_value() {
                                self.series.layout = parse_layout_id(&v);
                            }
                        }
                        b"uniqueId" => {
                            self.series.unique_id =
                                attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"hidden" => {
                            self.series.hidden = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        b"ownerIdx" => {
                            self.series.owner_idx =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        b"formatIdx" => {
                            self.series.format_idx =
                                attr.unescape_value().ok().and_then(|s| s.parse().ok())
                        }
                        _ => {}
                    }
                }
            }
            b"tx" if self.series.open && !self.labels.open => {
                self.series.in_tx = true;
                self.series.tx_value = None;
                self.series.tx_formula = None;
            }
            b"txData" if self.series.in_tx => self.series.in_tx_data = true,
            b"v" if self.series.in_tx_data => self.series.in_tx_data_v = true,
            b"f" if self.series.in_tx_data => self.series.in_tx_data_f = true,
            b"dataId" if self.series.open && !self.labels.open => {
                self.series.in_data_id = true;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"val" {
                        self.series.data_id = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"dataLabels" if self.series.open => {
                self.labels.open = true;
                self.labels.overrides = Vec::new();
                self.labels.hidden = Vec::new();
                self.labels.pos = None;
                self.labels.vis_series = None;
                self.labels.vis_cat = None;
                self.labels.vis_val = None;
                self.labels.num_fmt = None;
                self.labels.separator = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"pos" {
                        self.labels.pos = attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"dataLabel" if self.labels.open && !self.label.open => {
                self.label.open = true;
                self.label.idx = 0;
                self.label.pos = None;
                self.label.vis_series = None;
                self.label.vis_cat = None;
                self.label.vis_val = None;
                self.label.num_fmt = None;
                self.label.separator = None;
                self.label.sp = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"idx" => {
                            self.label.idx = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0);
                        }
                        b"pos" => {
                            self.label.pos = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        _ => {}
                    }
                }
            }
            b"dataLabelHidden" if self.labels.open => {
                if let Some(idx) =
                    get_val_attr_named(e, b"idx").and_then(|s| s.parse().ok())
                {
                    self.labels.hidden.push(idx);
                }
            }
            b"separator" if self.label.open => self.label.in_separator = true,
            b"separator" if self.labels.open => self.labels.in_separator = true,
            b"dataPt" if self.series.open && !self.labels.open => {
                self.data_pt.open = true;
                self.data_pt.idx = 0;
                self.data_pt.sp = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"idx" {
                        self.data_pt.idx = attr
                            .unescape_value()
                            .ok()
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                }
            }
            b"layoutPr" if self.series.open => {
                self.layout.open = true;
                self.layout.pr = ChartExLayoutPr::default();
            }
            b"subtotals" if self.layout.open => {
                self.layout.in_subtotals = true;
                self.layout.pr.subtotals.get_or_insert_with(Vec::new);
            }
            b"idx" if self.layout.in_subtotals => push_subtotal_idx(&mut self.layout.pr, e),
            b"binning" if self.layout.open => {
                self.layout.in_binning = true;
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
                self.layout.pr.binning = Some(binning);
            }
            b"binSize" if self.layout.in_binning => {
                if let (Some(v), Some(b)) = (get_val_f64(e), self.layout.pr.binning.as_mut()) {
                    b.bin_size = Some(v);
                }
            }
            b"binCount" if self.layout.in_binning => {
                if let (Some(v), Some(b)) = (
                    get_val_attr(e).and_then(|s| s.parse().ok()),
                    self.layout.pr.binning.as_mut(),
                ) {
                    b.bin_count = Some(v);
                }
            }
            b"geography" if self.layout.open => {
                self.layout.in_geography = true;
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
                self.layout.pr.geography = Some(geo);
            }
            b"geoCache" if self.layout.in_geography => {
                self.begin_capture(CaptureDest::GeoCache, e);
            }
            b"statistics" if self.layout.open => {
                self.layout.in_statistics = true;
                let mut stats = ChartExStatistics::default();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"quartileMethod" {
                        stats.quartile_method =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
                self.layout.pr.statistics = Some(stats);
            }
            b"visibility" if self.layout.open && !self.labels.open => {
                self.layout.in_visibility = true;
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
                self.layout.pr.visibility = Some(vis);
            }
            b"axisId" if self.series.open && !self.axis.open => self.series.in_axis_id = true,
            b"valueColors" if self.series.open => {
                self.series.in_value_colors = true;
                self.series.value_colors = Some(ChartExValueColors::default());
            }
            b"minColor" | b"midColor" | b"maxColor" if self.series.in_value_colors => {
                let slot = match tag {
                    b"minColor" => ValueColorSlot::Min,
                    b"midColor" => ValueColorSlot::Mid,
                    _ => ValueColorSlot::Max,
                };
                self.begin_capture(CaptureDest::ValueColor(slot), e);
            }
            b"valueColorPositions" if self.series.open => {
                self.color_pos.open = true;
                self.color_pos.value = ChartExValueColorPositions::default();
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"count" {
                        self.color_pos.value.count = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
            }
            b"min" if self.color_pos.open => self.color_pos.in_min = true,
            b"mid" if self.color_pos.open => self.color_pos.in_mid = true,
            b"max" if self.color_pos.open => self.color_pos.in_max = true,
            b"axis" if self.pos.plot_area && !self.pos.plot_area_region => {
                self.axis.open = true;
                self.axis.id = 0;
                self.axis.hidden = None;
                self.axis.scaling = ChartExScaling::Category { gap_width: None };
                self.axis.title = None;
                self.axis.units = None;
                self.axis.major_gridlines = None;
                self.axis.minor_gridlines = None;
                self.axis.major_tick_marks = None;
                self.axis.minor_tick_marks = None;
                self.axis.tick_labels = false;
                self.axis.num_fmt = None;
                self.axis.shape_properties = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"id" => {
                            self.axis.id = attr
                                .unescape_value()
                                .ok()
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(0)
                        }
                        b"hidden" => {
                            self.axis.hidden = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"catScaling" if self.axis.open => {
                self.axis.in_cat_scaling = true;
                let mut gw = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"gapWidth" {
                        gw = attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
                self.axis.scaling = ChartExScaling::Category { gap_width: gw };
            }
            b"valScaling" if self.axis.open => {
                self.axis.in_val_scaling = true;
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
                self.axis.scaling = ChartExScaling::Value {
                    min,
                    max,
                    major_unit: major,
                    minor_unit: minor,
                };
            }
            b"title" if self.axis.open => {
                self.axis.in_title = true;
                self.axis.title_text = None;
            }
            b"tx" if self.axis.in_title => self.axis.in_title_tx = true,
            b"txData" if self.axis.in_title_tx => self.axis.in_title_tx_data = true,
            b"v" if self.axis.in_title_tx_data => self.axis.in_title_tx_data_v = true,
            b"units" if self.axis.open => {
                self.axis.in_units = true;
                self.axis.units_unit = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"unit" {
                        self.axis.units_unit = attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"majorGridlines" if self.axis.open => {
                self.axis.in_major_gridlines = true;
                self.axis.major_gridlines = Some(ChartExGridlines::default());
            }
            b"minorGridlines" if self.axis.open => {
                self.axis.in_minor_gridlines = true;
                self.axis.minor_gridlines = Some(ChartExGridlines::default());
            }
            b"legend" if self.pos.chart && !self.pos.plot_area => {
                self.legend.open = true;
                self.legend.pos = None;
                self.legend.align = None;
                self.legend.overlay = None;
                self.legend.sp = None;
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"pos" => {
                            self.legend.pos = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"align" => {
                            self.legend.align = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        b"overlay" => {
                            self.legend.overlay = attr
                                .unescape_value()
                                .ok()
                                .map(|s| s == "1" || s.as_ref() == "true")
                        }
                        _ => {}
                    }
                }
            }
            b"printSettings" if self.pos.chart_space => {
                self.print.open = true;
                self.print.settings = ChartExPrintSettings::default();
            }
            b"headerFooter" if self.print.open => {
                self.print.in_header_footer = true;
                self.print.hf = ChartExHeaderFooter::default();
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"alignWithMargins" => {
                            self.print.hf.align_with_margins = parse_bool_attr(&attr)
                        }
                        b"differentOddEven" => {
                            self.print.hf.different_odd_even = parse_bool_attr(&attr)
                        }
                        b"differentFirst" => self.print.hf.different_first = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"oddHeader" if self.print.in_header_footer => self.print.in_hf_odd_header = true,
            b"oddFooter" if self.print.in_header_footer => self.print.in_hf_odd_footer = true,
            b"evenHeader" if self.print.in_header_footer => self.print.in_hf_even_header = true,
            b"evenFooter" if self.print.in_header_footer => self.print.in_hf_even_footer = true,
            b"firstHeader" if self.print.in_header_footer => self.print.in_hf_first_header = true,
            b"firstFooter" if self.print.in_header_footer => self.print.in_hf_first_footer = true,
            b"spPr" if !self.sp.open => {
                self.sp.open = true;
                self.sp.depth = 1;
                self.sp.solid_fill = None;
                self.sp.no_fill = false;
                self.sp.line = None;
                self.sp.in_ln = false;
                self.sp.ln_width = None;
                self.sp.ln_solid_fill = None;
                self.sp.ln_no_fill = false;
                self.sp.ln_dash = None;
                if self.data_pt.open {
                    self.sp.ctx = SpCtx::DataPt;
                } else if self.label.open {
                    self.sp.ctx = SpCtx::DataLabel;
                } else if self.labels.open {
                    self.sp.ctx = SpCtx::DataLabels;
                } else if self.series.open {
                    self.sp.ctx = SpCtx::Series;
                } else if self.axis.in_major_gridlines {
                    self.sp.ctx = SpCtx::MajorGrid;
                } else if self.axis.in_minor_gridlines {
                    self.sp.ctx = SpCtx::MinorGrid;
                } else if self.axis.in_title {
                    self.sp.ctx = SpCtx::AxisTitle;
                } else if self.axis.open {
                    self.sp.ctx = SpCtx::Axis;
                } else if self.title.open {
                    self.sp.ctx = SpCtx::Title;
                } else if self.legend.open {
                    self.sp.ctx = SpCtx::Legend;
                } else if self.pos.plot_surface {
                    self.sp.ctx = SpCtx::PlotSurface;
                } else if self.pos.chart_space && !self.pos.chart {
                    self.sp.ctx = SpCtx::ChartSpace;
                } else {
                    self.sp.ctx = SpCtx::None;
                }
            }
            b"ln" if self.sp.open => {
                self.sp.in_ln = true;
                self.sp.ln_width = None;
                self.sp.ln_solid_fill = None;
                self.sp.ln_no_fill = false;
                self.sp.ln_dash = None;
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"w" {
                        self.sp.ln_width =
                            attr.unescape_value().ok().and_then(|s| s.parse().ok());
                    }
                }
            }
            b"offset" => {
                let mut offset = ChartExOffset::default();
                for attr in e.attributes().flatten() {
                    let v = attr.unescape_value().ok().and_then(|s| s.parse::<f64>().ok());
                    match attr.key.local_name().as_ref() {
                        b"t" => offset.top = v,
                        b"l" => offset.left = v,
                        _ => {}
                    }
                }
                if self.axis.in_title {
                    self.axis.title_offset = Some(offset);
                } else if self.title.open {
                    self.title.offset = Some(offset);
                } else if self.legend.open {
                    self.legend.offset = Some(offset);
                }
            }
            b"txPr" | b"rich" | b"clrMapOvr" | b"fmtOvrs" | b"extLst" => self.begin_skip(),
            // Handling merged from the former separate arm for
            // empty elements, which a self-closing element no
            // longer reaches.
            b"externalData" if self.pos.chart_data => {
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
            b"visibility" if self.label.open => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"seriesName" => self.label.vis_series = parse_bool_attr(&attr),
                        b"categoryName" => self.label.vis_cat = parse_bool_attr(&attr),
                        b"value" => self.label.vis_val = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"visibility" if self.labels.open => {
                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"seriesName" => self.labels.vis_series = parse_bool_attr(&attr),
                        b"categoryName" => self.labels.vis_cat = parse_bool_attr(&attr),
                        b"value" => self.labels.vis_val = parse_bool_attr(&attr),
                        _ => {}
                    }
                }
            }
            b"numFmt" if self.label.open => self.label.num_fmt = Some(parse_num_fmt(e)),
            b"numFmt" if self.labels.open => self.labels.num_fmt = Some(parse_num_fmt(e)),
            b"numFmt" if self.axis.open && !self.labels.open => {
                self.axis.num_fmt = Some(parse_num_fmt(e));
            }
            b"parentLabelLayout" if self.layout.open => {
                self.layout.pr.parent_label_layout = get_val_attr(e);
            }
            b"regionLabelLayout" if self.layout.open => {
                self.layout.pr.region_label_layout = get_val_attr(e);
            }
            b"aggregation" if self.layout.open => self.layout.pr.aggregation = true,
            b"majorTickMarks" if self.axis.open => {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        self.axis.major_tick_marks =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"minorTickMarks" if self.axis.open => {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"type" {
                        self.axis.minor_tick_marks =
                            attr.unescape_value().ok().map(|s| s.to_string());
                    }
                }
            }
            b"tickLabels" if self.axis.open => self.axis.tick_labels = true,
            b"srgbClr" if self.sp.open && !self.sp.in_ln => {
                if let Some(hex) = get_val_attr(e) {
                    self.sp.solid_fill = Some(ChartColor { hex });
                }
            }
            b"srgbClr" if self.sp.in_ln => {
                if let Some(hex) = get_val_attr(e) {
                    self.sp.ln_solid_fill = Some(ChartColor { hex });
                }
            }
            b"noFill" if self.sp.open && !self.sp.in_ln => self.sp.no_fill = true,
            b"noFill" if self.sp.in_ln => self.sp.ln_no_fill = true,
            b"prstDash" if self.sp.in_ln => self.sp.ln_dash = get_val_attr(e),
            b"pageMargins" if self.print.open => {
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
                self.print.settings.page_margins = Some(pm);
            }
            b"pageSetup" if self.print.open => {
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
                self.print.settings.page_setup = Some(ps);
            }
            b"extremeValue" if self.color_pos.in_min => {
                self.color_pos.value.min = Some(ChartExColorPosition::ExtremeValue)
            }
            b"extremeValue" if self.color_pos.in_mid => {
                self.color_pos.value.mid = Some(ChartExColorPosition::ExtremeValue)
            }
            b"extremeValue" if self.color_pos.in_max => {
                self.color_pos.value.max = Some(ChartExColorPosition::ExtremeValue)
            }
            b"number" if self.color_pos.in_min => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.min = Some(ChartExColorPosition::Number(v));
                }
            }
            b"number" if self.color_pos.in_mid => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.mid = Some(ChartExColorPosition::Number(v));
                }
            }
            b"number" if self.color_pos.in_max => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.max = Some(ChartExColorPosition::Number(v));
                }
            }
            b"percent" if self.color_pos.in_min => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.min = Some(ChartExColorPosition::Percent(v));
                }
            }
            b"percent" if self.color_pos.in_mid => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.mid = Some(ChartExColorPosition::Percent(v));
                }
            }
            b"percent" if self.color_pos.in_max => {
                if let Some(v) = get_val_f64(e) {
                    self.color_pos.value.max = Some(ChartExColorPosition::Percent(v));
                }
            }
            _ => {}
        }
    }

    fn on_text(&mut self, e: &BytesText) {
        if let Ok(text) = e.unescape() {
            let t = text.as_ref();
            if self.data.in_f {
                self.data.formula = Some(t.to_string());
            } else if self.data.in_nf {
                self.data.nf = Some(t.to_string());
            } else if self.data.in_lvl_pt_text && self.data.in_lvl_pt {
                if self.data.in_str_dim {
                    self.data.lvl_str_points.push((self.data.lvl_pt_idx, t.to_string()));
                } else if self.data.in_num_dim {
                    self.data.lvl_num_points.push((self.data.lvl_pt_idx, t.to_string()));
                }
            } else if self.title.in_tx_data_v {
                self.title.text = Some(t.to_string());
            } else if self.title.in_tx_data_f {
                self.title.text = Some(t.to_string());
            } else if self.series.in_tx_data_v {
                self.series.tx_value = Some(t.to_string());
            } else if self.series.in_tx_data_f {
                self.series.tx_formula = Some(t.to_string());
            } else if self.series.in_axis_id {
                if let Ok(id) = t.parse::<u32>() {
                    self.series.axis_ids.push(id);
                }
            } else if self.axis.in_title_tx_data_v {
                self.axis.title_text = Some(t.to_string());
            } else if self.label.in_separator {
                self.label.separator = Some(t.to_string());
            } else if self.labels.in_separator {
                self.labels.separator = Some(t.to_string());
            } else if self.print.in_hf_odd_header {
                self.print.hf.odd_header = Some(t.to_string());
            } else if self.print.in_hf_odd_footer {
                self.print.hf.odd_footer = Some(t.to_string());
            } else if self.print.in_hf_even_header {
                self.print.hf.even_header = Some(t.to_string());
            } else if self.print.in_hf_even_footer {
                self.print.hf.even_footer = Some(t.to_string());
            } else if self.print.in_hf_first_header {
                self.print.hf.first_header = Some(t.to_string());
            } else if self.print.in_hf_first_footer {
                self.print.hf.first_footer = Some(t.to_string());
            }
        }
    }

    fn on_end(&mut self, e: &BytesEnd) {
        let local = e.name().local_name();
        let tag = local.as_ref();
        match tag {
            b"chartSpace" => self.pos.chart_space = false,
            b"chartData" => self.pos.chart_data = false,
            b"data" if self.data.open => {
                self.result.data.push(ChartExData {
                    id: self.data.id,
                    dimensions: std::mem::take(&mut self.data.dims),
                    extensions: None,
                });
                self.data.open = false;
            }
            b"strDim" if self.data.in_str_dim => {
                self.data.dims.push(ChartExDimension::String {
                    dim_type: self.data.str_type.clone(),
                    formula: self.data.formula.take(),
                    nf_formula: self.data.nf.take(),
                    levels: std::mem::take(&mut self.data.str_levels),
                });
                self.data.in_str_dim = false;
            }
            b"numDim" if self.data.in_num_dim => {
                self.data.dims.push(ChartExDimension::Numeric {
                    dim_type: self.data.num_type.clone(),
                    formula: self.data.formula.take(),
                    nf_formula: self.data.nf.take(),
                    levels: std::mem::take(&mut self.data.num_levels),
                });
                self.data.in_num_dim = false;
            }
            b"f" if self.data.in_f => self.data.in_f = false,
            b"nf" if self.data.in_nf => self.data.in_nf = false,
            b"lvl" if self.data.in_lvl => {
                if self.data.in_str_dim {
                    self.data.str_levels.push(ChartExStringLevel {
                        pt_count: self.data.lvl_pt_count,
                        name: self.data.lvl_name.take(),
                        points: std::mem::take(&mut self.data.lvl_str_points),
                    });
                } else if self.data.in_num_dim {
                    self.data.num_levels.push(ChartExNumericLevel {
                        pt_count: self.data.lvl_pt_count,
                        format_code: self.data.lvl_format_code.take(),
                        name: self.data.lvl_name.take(),
                        points: std::mem::take(&mut self.data.lvl_num_points),
                    });
                }
                self.data.in_lvl = false;
            }
            b"pt" if self.data.in_lvl_pt => {
                self.data.in_lvl_pt = false;
                self.data.in_lvl_pt_text = false;
            }
            b"chart" if self.pos.chart => self.pos.chart = false,
            b"title" if self.axis.in_title => {
                let title = ChartExAxisTitle {
                    text: self.axis.title_text.take().map(|t| ChartExText {
                        data: Some(ChartExTextData {
                            formula: None,
                            value: Some(t),
                        }),
                        rich: None,
                    }),
                    offset: self.axis.title_offset.take(),
                    shape_properties: self.axis.title_sp.take(),
                    text_properties: None,
                    extensions: None,
                };
                self.axis.title = Some(title);
                self.axis.in_title = false;
                self.axis.in_title_tx = false;
                self.axis.in_title_tx_data = false;
                self.axis.in_title_tx_data_v = false;
            }
            b"title" if self.title.open => {
                self.result.title = Some(ChartExTitle {
                    text: self.title.text.take(),
                    rich_text: None,
                    position: self.title.pos.take(),
                    align: self.title.align.take(),
                    overlay: self.title.overlay.take(),
                    offset: self.title.offset.take(),
                    shape_properties: self.title.sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.title.open = false;
                self.title.in_tx = false;
                self.title.in_tx_data = false;
            }
            b"tx" if self.title.in_tx && !self.series.open => self.title.in_tx = false,
            b"txData" if self.title.in_tx_data && !self.series.open => self.title.in_tx_data = false,
            b"v" if self.title.in_tx_data_v && !self.series.open => self.title.in_tx_data_v = false,
            b"f" if self.title.in_tx_data_f && !self.series.open => self.title.in_tx_data_f = false,
            b"plotArea" => self.pos.plot_area = false,
            b"plotAreaRegion" => self.pos.plot_area_region = false,
            b"plotSurface" if self.pos.plot_surface => self.pos.plot_surface = false,
            b"series" if self.series.open => {
                let series = ChartExSeries {
                    layout: self.series.layout.clone(),
                    unique_id: self.series.unique_id.take(),
                    hidden: self.series.hidden.take(),
                    owner_idx: self.series.owner_idx.take(),
                    format_idx: self.series.format_idx.take(),
                    text: self.series.text.take(),
                    data_id: self.series.data_id,
                    data_labels: self.series.data_labels.take(),
                    data_points: std::mem::take(&mut self.series.data_points),
                    layout_properties: self.series.layout_pr.take(),
                    axis_ids: std::mem::take(&mut self.series.axis_ids),
                    value_colors: self.series.value_colors.take(),
                    value_color_positions: self.series.value_color_positions.take(),
                    shape_properties: self.series.shape_properties.take(),
                    extensions: None,
                };
                self.result.plot_area.series.push(series);
                self.series.open = false;
            }
            b"tx" if self.series.in_tx => {
                self.series.text = Some(ChartExText {
                    data: Some(ChartExTextData {
                        formula: self.series.tx_formula.take(),
                        value: self.series.tx_value.take(),
                    }),
                    rich: None,
                });
                self.series.in_tx = false;
                self.series.in_tx_data = false;
            }
            b"txData" if self.series.in_tx_data => self.series.in_tx_data = false,
            b"v" if self.series.in_tx_data_v => self.series.in_tx_data_v = false,
            b"f" if self.series.in_tx_data_f => self.series.in_tx_data_f = false,
            b"dataId" if self.series.in_data_id => self.series.in_data_id = false,
            b"dataLabels" if self.labels.open => {
                self.series.data_labels = Some(ChartExDataLabels {
                    position: self.labels.pos.take(),
                    visibility_series_name: self.labels.vis_series.take(),
                    visibility_category_name: self.labels.vis_cat.take(),
                    visibility_value: self.labels.vis_val.take(),
                    number_format: self.labels.num_fmt.take(),
                    separator: self.labels.separator.take(),
                    shape_properties: self.labels.sp.take(),
                    overrides: std::mem::take(&mut self.labels.overrides),
                    hidden_labels: std::mem::take(&mut self.labels.hidden),
                    text_properties: None,
                    extensions: None,
                });
                self.labels.open = false;
            }
            b"dataLabel" if self.label.open => {
                self.labels.overrides.push(ChartExDataLabel {
                    idx: self.label.idx,
                    position: self.label.pos.take(),
                    visibility_series_name: self.label.vis_series.take(),
                    visibility_category_name: self.label.vis_cat.take(),
                    visibility_value: self.label.vis_val.take(),
                    number_format: self.label.num_fmt.take(),
                    separator: self.label.separator.take(),
                    shape_properties: self.label.sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.label.open = false;
            }
            b"separator" if self.label.in_separator => self.label.in_separator = false,
            b"separator" if self.labels.in_separator => self.labels.in_separator = false,
            b"dataPt" if self.data_pt.open => {
                self.series.data_points.push(ChartExDataPoint {
                    idx: self.data_pt.idx,
                    shape_properties: self.data_pt.sp.take(),
                    extensions: None,
                });
                self.data_pt.open = false;
            }
            b"layoutPr" if self.layout.open => {
                self.series.layout_pr = Some(std::mem::take(&mut self.layout.pr));
                self.layout.open = false;
            }
            b"subtotals" if self.layout.in_subtotals => self.layout.in_subtotals = false,
            b"binning" if self.layout.in_binning => self.layout.in_binning = false,
            b"geography" if self.layout.in_geography => self.layout.in_geography = false,
            b"statistics" if self.layout.in_statistics => self.layout.in_statistics = false,
            b"visibility" if self.layout.in_visibility => self.layout.in_visibility = false,
            b"axisId" if self.series.in_axis_id => self.series.in_axis_id = false,
            b"valueColors" if self.series.in_value_colors => self.series.in_value_colors = false,
            b"valueColorPositions" if self.color_pos.open => {
                self.series.value_color_positions = Some(self.color_pos.value.clone());
                self.color_pos.open = false;
                self.color_pos.in_min = false;
                self.color_pos.in_mid = false;
                self.color_pos.in_max = false;
            }
            b"min" if self.color_pos.in_min => self.color_pos.in_min = false,
            b"mid" if self.color_pos.in_mid => self.color_pos.in_mid = false,
            b"max" if self.color_pos.in_max => self.color_pos.in_max = false,
            b"axis" if self.axis.open => {
                self.result.plot_area.axes.push(ChartExAxis {
                    id: self.axis.id,
                    hidden: self.axis.hidden.take(),
                    scaling: self.axis.scaling.clone(),
                    title: self.axis.title.take(),
                    units: self.axis.units.take(),
                    major_gridlines: self.axis.major_gridlines.take(),
                    minor_gridlines: self.axis.minor_gridlines.take(),
                    major_tick_marks: self.axis.major_tick_marks.take(),
                    minor_tick_marks: self.axis.minor_tick_marks.take(),
                    tick_labels: self.axis.tick_labels,
                    number_format: self.axis.num_fmt.take(),
                    shape_properties: self.axis.shape_properties.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.axis.open = false;
                self.axis.in_cat_scaling = false;
                self.axis.in_val_scaling = false;
            }
            b"catScaling" if self.axis.in_cat_scaling => self.axis.in_cat_scaling = false,
            b"valScaling" if self.axis.in_val_scaling => self.axis.in_val_scaling = false,
            b"tx" if self.axis.in_title_tx => self.axis.in_title_tx = false,
            b"txData" if self.axis.in_title_tx_data => self.axis.in_title_tx_data = false,
            b"v" if self.axis.in_title_tx_data_v => self.axis.in_title_tx_data_v = false,
            b"units" if self.axis.in_units => {
                self.axis.units = Some(ChartExAxisUnits {
                    unit: self.axis.units_unit.take(),
                    label: None,
                    extensions: None,
                });
                self.axis.in_units = false;
            }
            b"majorGridlines" if self.axis.in_major_gridlines => self.axis.in_major_gridlines = false,
            b"minorGridlines" if self.axis.in_minor_gridlines => self.axis.in_minor_gridlines = false,
            b"legend" if self.legend.open => {
                self.result.legend = Some(ChartExLegend {
                    position: self.legend.pos.take(),
                    align: self.legend.align.take(),
                    overlay: self.legend.overlay.take(),
                    offset: self.legend.offset.take(),
                    shape_properties: self.legend.sp.take(),
                    text_properties: None,
                    extensions: None,
                });
                self.legend.open = false;
            }
            b"printSettings" if self.print.open => {
                self.result.print_settings = Some(self.print.settings.clone());
                self.print.open = false;
            }
            b"headerFooter" if self.print.in_header_footer => {
                self.print.settings.header_footer = Some(self.print.hf.clone());
                self.print.in_header_footer = false;
            }
            b"oddHeader" => self.print.in_hf_odd_header = false,
            b"oddFooter" => self.print.in_hf_odd_footer = false,
            b"evenHeader" => self.print.in_hf_even_header = false,
            b"evenFooter" => self.print.in_hf_even_footer = false,
            b"firstHeader" => self.print.in_hf_first_header = false,
            b"firstFooter" => self.print.in_hf_first_footer = false,
            b"ln" if self.sp.in_ln => {
                self.sp.line = Some(ChartLine {
                    width: self.sp.ln_width.take(),
                    solid_fill: self.sp.ln_solid_fill.take(),
                    no_fill: self.sp.ln_no_fill,
                    dash_style: self.sp.ln_dash.take(),
                });
                self.sp.in_ln = false;
                self.sp.ln_no_fill = false;
                self.sp.depth = self.sp.depth.saturating_sub(1);
            }
            b"spPr" if self.sp.open => {
                self.sp.depth = self.sp.depth.saturating_sub(1);
                if self.sp.depth == 0 {
                    let props = ChartShapeProperties {
                        solid_fill: self.sp.solid_fill.take(),
                        no_fill: self.sp.no_fill,
                        line: self.sp.line.take(),
                    };
                    let has =
                        props.solid_fill.is_some() || props.no_fill || props.line.is_some();
                    match self.sp.ctx {
                        SpCtx::ChartSpace => {
                            if has {
                                self.result.shape_properties = Some(props);
                            }
                        }
                        SpCtx::Series => {
                            if has {
                                self.series.shape_properties = Some(props);
                            }
                        }
                        SpCtx::DataPt => {
                            self.data_pt.sp = if has { Some(props) } else { None };
                        }
                        SpCtx::DataLabels => {
                            if has {
                                self.labels.sp = Some(props);
                            }
                        }
                        SpCtx::DataLabel => {
                            if has {
                                self.label.sp = Some(props);
                            }
                        }
                        SpCtx::AxisTitle => {
                            if has {
                                self.axis.title_sp = Some(props);
                            }
                        }
                        SpCtx::Title => {
                            if has {
                                self.title.sp = Some(props);
                            }
                        }
                        SpCtx::Legend => {
                            if has {
                                self.legend.sp = Some(props);
                            }
                        }
                        SpCtx::Axis => {
                            if has {
                                self.axis.shape_properties = Some(props);
                            }
                        }
                        SpCtx::MajorGrid => {
                            self.axis.major_gridlines
                                .get_or_insert_with(ChartExGridlines::default)
                                .shape_properties = Some(props);
                        }
                        SpCtx::MinorGrid => {
                            self.axis.minor_gridlines
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
                    self.sp.open = false;
                    self.sp.no_fill = false;
                    self.sp.ctx = SpCtx::None;
                }
            }
            _ => {
                if self.sp.open {
                    self.sp.depth = self.sp.depth.saturating_sub(1);
                    if self.sp.depth == 0 {
                        self.sp.open = false;
                        self.sp.no_fill = false;
                        self.sp.ctx = SpCtx::None;
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
    let mut parser = ChartExParser::default();

    loop {
        buf.clear();
        let event = match xml_reader.read_event_into(buf) {
            Ok(ev) => ev,
            Err(e) => return Err(ChartParseError::Xml(e)),
        };
        if matches!(event, Event::Eof) {
            break;
        }
        parser.handle(&event);
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
