//! Model types for ChartEx (Office 2016+ extended charts).
//!
//! ChartEx charts use the `cx:chartSpace` namespace
//! (`http://schemas.microsoft.com/office/drawing/2014/chartex`) and support
//! chart types that don't fit the original `c:chartSpace` model: waterfall,
//! treemap, sunburst, funnel, histogram, box & whisker, pareto, region map,
//! and clustered column.

use std::collections::HashMap;

use crate::formatting::{ChartShapeProperties, NumberFormat};
use crate::text_properties::TextProperties;

/// Series layout type (the `layoutId` attribute on `cx:series`).
#[derive(Debug, Clone, PartialEq)]
pub enum ChartExLayout {
    /// Waterfall chart
    Waterfall,
    /// Treemap chart
    Treemap,
    /// Sunburst chart
    Sunburst,
    /// Funnel chart
    Funnel,
    /// Histogram chart
    Histogram,
    /// Box & whisker chart
    BoxWhisker,
    /// Pareto line chart
    ParetoLine,
    /// Region map (geography) chart
    RegionMap,
    /// Clustered column chart
    ClusteredColumn,
    /// Unrecognized layout type (forward-compatible)
    Unknown(String),
}

/// Top-level ChartEx type, analogous to `Chart` for standard charts.
///
/// Represents the entire `cx:chartSpace` element.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartEx {
    /// Version attribute on cx:chartSpace
    pub version: Option<String>,
    /// Feature list attribute on cx:chartSpace
    pub feature_list: Option<String>,
    /// Fallback image relationship id
    pub fallback_img: Option<String>,
    /// Chart title (`cx:chart > cx:title`)
    pub title: Option<ChartExTitle>,
    /// Shared data blocks (`cx:chartData > cx:data`)
    pub data: Vec<ChartExData>,
    /// External data source link (`cx:chartData > cx:externalData`)
    pub external_data: Option<ChartExExternalData>,
    /// Plot area containing series and axes (`cx:chart > cx:plotArea`)
    pub plot_area: ChartExPlotArea,
    /// Legend (`cx:chart > cx:legend`)
    pub legend: Option<ChartExLegend>,
    /// Chart-level shape properties (`cx:chartSpace > cx:spPr`)
    pub shape_properties: Option<ChartShapeProperties>,
    /// Chart-level text properties (`cx:chartSpace > cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Color map override (`cx:chartSpace > cx:clrMapOvr`, raw XML bytes)
    pub color_map_override: Option<Vec<u8>>,
    /// Format overrides (`cx:chartSpace > cx:fmtOvrs > cx:fmtOvr`)
    pub format_overrides: Vec<ChartExFormatOverride>,
    /// Print settings (`cx:chartSpace > cx:printSettings`)
    pub print_settings: Option<ChartExPrintSettings>,
    /// Raw chart style XML (preserved for roundtrip)
    pub raw_chart_style: Option<Vec<u8>>,
    /// Raw chart color style XML (preserved for roundtrip)
    pub raw_chart_color_style: Option<Vec<u8>>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
    /// Raw extension elements keyed by namespace (preserved for roundtrip)
    pub raw_extensions: HashMap<String, Vec<u8>>,
    /// Raw `mc:Fallback` XML for drawing roundtrip
    pub raw_mc_fallback: Option<Vec<u8>>,
}

/// Chart title (`cx:chart > cx:title`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExTitle {
    /// Literal text from `cx:txData > cx:v`, or formula from `cx:f`
    pub text: Option<String>,
    /// Rich text (`cx:rich`, raw DrawingML XML bytes)
    pub rich_text: Option<Vec<u8>>,
    /// Position attribute (`t`, `b`, `l`, `r`)
    pub position: Option<String>,
    /// Alignment attribute (`min`, `ctr`, `max`)
    pub align: Option<String>,
    /// Overlay attribute
    pub overlay: Option<bool>,
    /// Title offset
    pub offset: Option<ChartExOffset>,
    /// Shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Plot area containing series and axes (`cx:chart > cx:plotArea`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExPlotArea {
    /// Plot surface shape properties (`cx:plotSurface > cx:spPr`)
    pub plot_surface: Option<ChartShapeProperties>,
    /// Series in the plot area region (`cx:plotAreaRegion > cx:series`)
    pub series: Vec<ChartExSeries>,
    /// Axes (`cx:axis` elements)
    pub axes: Vec<ChartExAxis>,
    /// Plot area shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// A shared data block (`cx:chartData > cx:data`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExData {
    /// Data block id (referenced by series via `cx:dataId`)
    pub id: u32,
    /// Dimensions (string or numeric) within this data block
    pub dimensions: Vec<ChartExDimension>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// External data source reference (`cx:externalData`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExExternalData {
    /// Relationship id pointing to the external data source (required)
    pub rel_id: String,
    /// Whether to automatically update from the external source
    pub auto_update: Option<bool>,
}

/// A data dimension - either string-typed or numeric-typed.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartExDimension {
    /// String dimension (`cx:strDim`)
    String {
        /// Dimension type: `cat`, `colorStr`, `entityId`
        dim_type: StringDimType,
        /// Formula reference (`cx:f`)
        formula: Option<String>,
        /// Display formula (`cx:nf`)
        nf_formula: Option<String>,
        /// Cached literal string levels
        levels: Vec<ChartExStringLevel>,
    },
    /// Numeric dimension (`cx:numDim`)
    Numeric {
        /// Dimension type: `val`, `x`, `y`, `size`, `colorVal`
        dim_type: NumericDimType,
        /// Formula reference (`cx:f`)
        formula: Option<String>,
        /// Display formula (`cx:nf`)
        nf_formula: Option<String>,
        /// Cached literal numeric levels
        levels: Vec<ChartExNumericLevel>,
    },
}

/// String dimension type attribute values.
#[derive(Debug, Clone, PartialEq)]
pub enum StringDimType {
    /// Category dimension
    Cat,
    /// Color string dimension
    ColorStr,
    /// Entity ID dimension (region maps)
    EntityId,
}

/// Numeric dimension type attribute values.
#[derive(Debug, Clone, PartialEq)]
pub enum NumericDimType {
    /// Value dimension
    Val,
    /// X coordinate dimension
    X,
    /// Y coordinate dimension
    Y,
    /// Size dimension
    Size,
    /// Color value dimension
    ColorVal,
}

/// Cached string level data within a string dimension.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExStringLevel {
    /// Number of points in this level
    pub pt_count: u32,
    /// Level name
    pub name: Option<String>,
    /// Cached data points as `(index, value)` pairs
    pub points: Vec<(u32, String)>,
}

/// Cached numeric level data within a numeric dimension.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExNumericLevel {
    /// Number of points in this level
    pub pt_count: u32,
    /// Number format code for this level
    pub format_code: Option<String>,
    /// Level name
    pub name: Option<String>,
    /// Cached data points as `(index, value_as_string)` pairs
    pub points: Vec<(u32, String)>,
}

/// A chart series (`cx:series` inside `cx:plotAreaRegion`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExSeries {
    /// Layout type (`layoutId` attribute, required)
    pub layout: ChartExLayout,
    /// Unique series identifier (`uniqueId` attribute)
    pub unique_id: Option<String>,
    /// Whether the series is hidden
    pub hidden: Option<bool>,
    /// Owner index for grouped series
    pub owner_idx: Option<u32>,
    /// Format index
    pub format_idx: Option<u32>,
    /// Series name (`cx:tx`)
    pub text: Option<ChartExText>,
    /// Reference to a `ChartExData` block (`cx:dataId val`)
    pub data_id: u32,
    /// Data labels for this series
    pub data_labels: Option<ChartExDataLabels>,
    /// Per-point formatting overrides
    pub data_points: Vec<ChartExDataPoint>,
    /// Layout-specific properties (`cx:layoutPr`)
    pub layout_properties: Option<ChartExLayoutPr>,
    /// Axis ID references (`cx:axisId`)
    pub axis_ids: Vec<u32>,
    /// Value-based color mapping
    pub value_colors: Option<ChartExValueColors>,
    /// Value color position stops
    pub value_color_positions: Option<ChartExValueColorPositions>,
    /// Series shape properties (`cx:spPr`)
    pub shape_properties: Option<ChartShapeProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Series text element (`cx:tx`), a choice of `txData` or `rich`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExText {
    /// Text data with formula/literal (`cx:txData`)
    pub data: Option<ChartExTextData>,
    /// Rich text (`cx:rich`, raw DrawingML XML bytes)
    pub rich: Option<Vec<u8>>,
}

/// Text data element (`cx:txData`) containing a formula or literal value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExTextData {
    /// Formula reference (`cx:f`)
    pub formula: Option<String>,
    /// Literal text value (`cx:v`)
    pub value: Option<String>,
}

/// Value-based color mapping for series (`cx:valueColors`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExValueColors {
    /// Minimum color (solid fill, raw DrawingML XML bytes)
    pub min_color: Option<Vec<u8>>,
    /// Midpoint color (solid fill, raw DrawingML XML bytes)
    pub mid_color: Option<Vec<u8>>,
    /// Maximum color (solid fill, raw DrawingML XML bytes)
    pub max_color: Option<Vec<u8>>,
}

/// Color position stops for value coloring (`cx:valueColorPositions`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExValueColorPositions {
    /// Number of position stops
    pub count: Option<u32>,
    /// Minimum position
    pub min: Option<ChartExColorPosition>,
    /// Midpoint position
    pub mid: Option<ChartExColorPosition>,
    /// Maximum position
    pub max: Option<ChartExColorPosition>,
}

/// A color position stop - extreme value, absolute number, or percentage.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartExColorPosition {
    /// Extreme (min or max) data value
    ExtremeValue,
    /// Absolute numeric value
    Number(f64),
    /// Percentage value
    Percent(f64),
}

/// Layout-specific properties (`cx:layoutPr`) that vary by chart type.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExLayoutPr {
    /// Parent label layout for treemap/sunburst (`none`, `banner`, `overlapping`)
    pub parent_label_layout: Option<String>,
    /// Region label layout for regionMap (`none`, `bestFitOnly`, `showAll`)
    pub region_label_layout: Option<String>,
    /// Box & whisker visibility toggles
    pub visibility: Option<ChartExSeriesVisibility>,
    /// Funnel aggregation (presence-only empty element)
    pub aggregation: bool,
    /// Histogram binning configuration
    pub binning: Option<ChartExBinning>,
    /// Region map geography settings
    pub geography: Option<ChartExGeography>,
    /// Box & whisker statistics settings
    pub statistics: Option<ChartExStatistics>,
    /// Waterfall subtotal bar indices (`cx:subtotals`).
    ///
    /// `None` is an absent `cx:subtotals` element; `Some(vec![])` is a
    /// present but empty one, which is what Excel writes for a waterfall
    /// with no subtotal bars. The wrapper is `minOccurs="0"` in
    /// [MS-ODRAWXML] §5.22 `CT_SeriesLayoutProperties`, so the two are
    /// distinct documents and round-trip separately.
    pub subtotals: Option<Vec<u32>>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Box & whisker visibility toggles (`cx:visibility` inside `cx:layoutPr`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExSeriesVisibility {
    /// Show connector lines
    pub connector_lines: Option<bool>,
    /// Show mean line
    pub mean_line: Option<bool>,
    /// Show mean marker
    pub mean_marker: Option<bool>,
    /// Show non-outlier points
    pub nonoutliers: Option<bool>,
    /// Show outlier points
    pub outliers: Option<bool>,
}

/// Histogram binning configuration (`cx:binning`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExBinning {
    /// Interval closed side (`l` or `r`)
    pub interval_closed: Option<String>,
    /// Underflow value (`auto` or a number as string)
    pub underflow: Option<String>,
    /// Overflow value (`auto` or a number as string)
    pub overflow: Option<String>,
    /// Bin size (`cx:binSize`)
    pub bin_size: Option<f64>,
    /// Bin count (`cx:binCount`)
    pub bin_count: Option<u32>,
}

/// Geography settings for region map charts (`cx:geography`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExGeography {
    /// Map projection type (`mercator`, `miller`, `robinson`, `albers`)
    pub projection_type: Option<String>,
    /// Viewed region type (from `GeoMappingLevel` enum)
    pub viewed_region_type: Option<String>,
    /// Culture language code
    pub culture_language: Option<String>,
    /// Culture region code
    pub culture_region: Option<String>,
    /// Attribution text
    pub attribution: Option<String>,
    /// Entire `cx:geoCache` element as raw XML bytes (contains `cx:binary`
    /// base64 blob and/or `cx:clear` structured geo query results)
    pub raw_geo_cache: Option<Vec<u8>>,
}

/// Box & whisker statistics settings (`cx:statistics`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExStatistics {
    /// Quartile calculation method (`inclusive` or `exclusive`)
    pub quartile_method: Option<String>,
}

/// An axis in a ChartEx chart (`cx:axis`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExAxis {
    /// Axis id (required)
    pub id: u32,
    /// Whether the axis is hidden
    pub hidden: Option<bool>,
    /// Scaling type and parameters (category or value)
    pub scaling: ChartExScaling,
    /// Axis title (`cx:title`)
    pub title: Option<ChartExAxisTitle>,
    /// Axis units (`cx:units`)
    pub units: Option<ChartExAxisUnits>,
    /// Major gridlines (`cx:majorGridlines`)
    pub major_gridlines: Option<ChartExGridlines>,
    /// Minor gridlines (`cx:minorGridlines`)
    pub minor_gridlines: Option<ChartExGridlines>,
    /// Major tick mark type (`in`, `out`, `cross`, `none`)
    pub major_tick_marks: Option<String>,
    /// Minor tick mark type
    pub minor_tick_marks: Option<String>,
    /// Whether tick labels are present (`cx:tickLabels`)
    pub tick_labels: bool,
    /// Number format for axis labels (`cx:numFmt`)
    pub number_format: Option<NumberFormat>,
    /// Axis shape properties (`cx:spPr`)
    pub shape_properties: Option<ChartShapeProperties>,
    /// Axis text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Gridlines on a ChartEx axis (`cx:majorGridlines` / `cx:minorGridlines`).
///
/// Presence of the element is what turns gridlines on; the formatting
/// override is separate and optional, so `CT_Gridlines`
/// ([MS-ODRAWXML] §5.22) is modelled as its own type rather than as a
/// bare `Option<ChartShapeProperties>`. That keeps a bare
/// `<cx:majorGridlines/>` - the form Excel writes - distinct from one
/// carrying a `cx:spPr`, and stops the writer inventing an empty
/// `cx:spPr` for gridlines that never had one.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExGridlines {
    /// Formatting override (`cx:spPr`), absent for plain gridlines.
    pub shape_properties: Option<ChartShapeProperties>,
}

/// Axis scaling - either category-based or value-based.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartExScaling {
    /// Category scaling (`cx:catScaling`)
    Category {
        /// Gap width between bars
        gap_width: Option<f64>,
    },
    /// Value scaling (`cx:valScaling`)
    Value {
        /// Minimum axis value
        min: Option<f64>,
        /// Maximum axis value
        max: Option<f64>,
        /// Major unit spacing
        major_unit: Option<f64>,
        /// Minor unit spacing
        minor_unit: Option<f64>,
    },
}

/// Axis title (`cx:axis > cx:title`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExAxisTitle {
    /// Title text
    pub text: Option<ChartExText>,
    /// Title offset
    pub offset: Option<ChartExOffset>,
    /// Shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Axis units (`cx:units`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExAxisUnits {
    /// Unit type (`hundreds`, `thousands`, ..., `percentage`)
    pub unit: Option<String>,
    /// Units label
    pub label: Option<ChartExAxisUnitsLabel>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Axis units label (`cx:unitsLabel`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExAxisUnitsLabel {
    /// Label text
    pub text: Option<ChartExText>,
    /// Shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Offset for positioning elements (`cx:offset`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExOffset {
    /// Top offset
    pub top: Option<f64>,
    /// Left offset
    pub left: Option<f64>,
}

/// Chart legend (`cx:chart > cx:legend`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExLegend {
    /// Position (`l`, `t`, `r`, `b`)
    pub position: Option<String>,
    /// Alignment (`min`, `ctr`, `max`)
    pub align: Option<String>,
    /// Whether the legend overlays the plot area
    pub overlay: Option<bool>,
    /// Legend offset
    pub offset: Option<ChartExOffset>,
    /// Shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Data labels for a series (`cx:dataLabels`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExDataLabels {
    /// Label position (from `DataLabelPos` enum)
    pub position: Option<String>,
    /// Show series name in labels
    pub visibility_series_name: Option<bool>,
    /// Show category name in labels
    pub visibility_category_name: Option<bool>,
    /// Show value in labels
    pub visibility_value: Option<bool>,
    /// Number format for label values
    pub number_format: Option<NumberFormat>,
    /// Separator between label parts
    pub separator: Option<String>,
    /// Shape properties for labels
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Per-point data label overrides
    pub overrides: Vec<ChartExDataLabel>,
    /// Indices of hidden labels (`cx:dataLabelHidden`)
    pub hidden_labels: Vec<u32>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Per-point data label override (`cx:dataLabel`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExDataLabel {
    /// Point index
    pub idx: u32,
    /// Label position override
    pub position: Option<String>,
    /// Show series name
    pub visibility_series_name: Option<bool>,
    /// Show category name
    pub visibility_category_name: Option<bool>,
    /// Show value
    pub visibility_value: Option<bool>,
    /// Number format override
    pub number_format: Option<NumberFormat>,
    /// Separator override
    pub separator: Option<String>,
    /// Shape properties override
    pub shape_properties: Option<ChartShapeProperties>,
    /// Text properties override (`cx:txPr`, raw XML bytes)
    pub text_properties: Option<TextProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Per-point formatting override (`cx:dataPt`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExDataPoint {
    /// Point index
    pub idx: u32,
    /// Shape properties for this data point
    pub shape_properties: Option<ChartShapeProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Format override entry (`cx:fmtOvr`).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartExFormatOverride {
    /// Override index
    pub idx: u32,
    /// Shape properties for this override
    pub shape_properties: Option<ChartShapeProperties>,
    /// Extension list (`cx:extLst`, raw XML bytes)
    pub extensions: Option<Vec<u8>>,
}

/// Print settings for the chart (`cx:printSettings`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExPrintSettings {
    /// Header and footer settings
    pub header_footer: Option<ChartExHeaderFooter>,
    /// Page margins
    pub page_margins: Option<ChartExPageMargins>,
    /// Page setup (paper size, orientation, etc.)
    pub page_setup: Option<ChartExPageSetup>,
}

/// Header and footer settings for print (`cx:headerFooter`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExHeaderFooter {
    /// Align headers/footers with page margins
    pub align_with_margins: Option<bool>,
    /// Different odd/even page headers/footers
    pub different_odd_even: Option<bool>,
    /// Different first page header/footer
    pub different_first: Option<bool>,
    /// Odd page header text
    pub odd_header: Option<String>,
    /// Odd page footer text
    pub odd_footer: Option<String>,
    /// Even page header text
    pub even_header: Option<String>,
    /// Even page footer text
    pub even_footer: Option<String>,
    /// First page header text
    pub first_header: Option<String>,
    /// First page footer text
    pub first_footer: Option<String>,
}

/// Page margins for print (`cx:pageMargins`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExPageMargins {
    /// Left margin
    pub left: Option<f64>,
    /// Right margin
    pub right: Option<f64>,
    /// Top margin
    pub top: Option<f64>,
    /// Bottom margin
    pub bottom: Option<f64>,
    /// Header margin
    pub header: Option<f64>,
    /// Footer margin
    pub footer: Option<f64>,
}

/// Page setup for print (`cx:pageSetup`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartExPageSetup {
    /// Paper size code
    pub paper_size: Option<u32>,
    /// First page number
    pub first_page_number: Option<u32>,
    /// Page orientation (`default`, `portrait`, `landscape`)
    pub orientation: Option<String>,
    /// Print in black and white
    pub black_and_white: Option<bool>,
    /// Draft quality printing
    pub draft: Option<bool>,
    /// Use the first page number setting
    pub use_first_page_number: Option<bool>,
    /// Horizontal DPI
    pub horizontal_dpi: Option<u32>,
    /// Vertical DPI
    pub vertical_dpi: Option<u32>,
    /// Number of copies
    pub copies: Option<u32>,
}
