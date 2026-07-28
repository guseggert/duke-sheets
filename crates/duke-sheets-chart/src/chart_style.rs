//! The chart style and chart colour style parts.
//!
//! Excel will not open a workbook whose chartEx part lacks these, and it
//! validates the style part rather than repairing it: every required
//! entry must be present, in schema order, each with its own reference
//! children ([MS-ODRAWXML] §5.15). Modelling the entries as named,
//! non-optional fields makes a part that Excel would reject unable to
//! exist in the first place.
//!
//! What sits *inside* an entry - `a:spPr`, `a:defRPr`, `a:bodyPr` - is
//! DrawingML, a vocabulary far larger than anything a chart style needs
//! to reason about, and is kept as the bytes it was read as. That keeps
//! a part read from a file byte-faithful through a round trip while
//! still letting the parts a caller cares about be named and edited.

/// A reference into the theme's style matrix (`cs:lnRef`, `cs:fillRef`,
/// `cs:effectRef`).
#[derive(Debug, Clone, PartialEq)]
pub struct StyleReference {
    /// Index into the matrix. `0` means "none".
    pub idx: u32,
    /// Colour override, as the complete child element.
    pub color: Option<Vec<u8>>,
}

impl StyleReference {
    /// A reference to matrix entry `idx` with no colour override.
    pub fn new(idx: u32) -> Self {
        Self { idx, color: None }
    }
}

impl Default for StyleReference {
    fn default() -> Self {
        Self::new(0)
    }
}

/// Which of the theme's fonts an entry uses (`cs:fontRef`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum FontCollection {
    /// The theme's major (heading) font.
    Major,
    /// The theme's minor (body) font.
    #[default]
    Minor,
    /// No theme font.
    None,
}

impl FontCollection {
    /// The `idx` attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Major => "major",
            Self::Minor => "minor",
            Self::None => "none",
        }
    }

    /// Parse an `idx` attribute value.
    pub fn from_str_value(value: &str) -> Option<Self> {
        match value {
            "major" => Some(Self::Major),
            "minor" => Some(Self::Minor),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

/// A font reference with its colour (`cs:fontRef`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct FontReference {
    /// Which theme font.
    pub collection: FontCollection,
    /// Colour, as the complete child element.
    pub color: Option<Vec<u8>>,
}

/// One entry of a chart style: how a single chart element is drawn.
///
/// `CT_StyleEntry` requires all four references; the rest is optional.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct StyleEntry {
    /// `cs:lnRef`.
    pub line_reference: StyleReference,
    /// `cs:lineWidthScale`.
    pub line_width_scale: Option<f64>,
    /// `cs:fillRef`.
    pub fill_reference: StyleReference,
    /// `cs:effectRef`.
    pub effect_reference: StyleReference,
    /// `cs:fontRef`.
    pub font_reference: FontReference,
    /// `a:spPr`, the complete element, verbatim.
    pub shape_properties: Option<Vec<u8>>,
    /// `a:defRPr`, the complete element, verbatim.
    pub default_run_properties: Option<Vec<u8>>,
    /// `a:bodyPr`, the complete element, verbatim.
    pub body_properties: Option<Vec<u8>>,
    /// `cs:extLst`, the complete element, verbatim.
    pub extensions: Option<Vec<u8>>,
    /// `mods`: which parts of the entry a consumer may override.
    pub mods: Option<String>,
}

impl StyleEntry {
    /// An entry that defers entirely to the theme.
    pub fn theme_default() -> Self {
        Self::default()
    }
}

/// Marker shape and size (`cs:dataPointMarkerLayout`), which is a marker
/// layout rather than a style entry.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct MarkerLayout {
    /// `symbol`.
    pub symbol: Option<String>,
    /// `size`, in points.
    pub size: Option<u32>,
}

/// A chart style part (`xl/charts/styleN.xml`).
///
/// The fields are the `CT_ChartStyle` sequence, in order. The two Excel
/// treats as optional are `Option`; the rest are required, so a value of
/// this type is always a part Excel will accept.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartStyle {
    /// Which entry of Excel's chart-style gallery this is.
    ///
    /// `CT_ChartStyle` marks the attribute optional, but Excel refuses a
    /// style part without it, so it is not optional here.
    pub id: u32,
    /// `cs:axisTitle`.
    pub axis_title: StyleEntry,
    /// `cs:categoryAxis`.
    pub category_axis: StyleEntry,
    /// `cs:chartArea`.
    pub chart_area: StyleEntry,
    /// `cs:dataLabel`.
    pub data_label: StyleEntry,
    /// `cs:dataLabelCallout`.
    pub data_label_callout: Option<StyleEntry>,
    /// `cs:dataPoint`.
    pub data_point: StyleEntry,
    /// `cs:dataPoint3D`.
    pub data_point_3d: StyleEntry,
    /// `cs:dataPointLine`.
    pub data_point_line: StyleEntry,
    /// `cs:dataPointMarker`.
    pub data_point_marker: StyleEntry,
    /// `cs:dataPointMarkerLayout`.
    pub data_point_marker_layout: Option<MarkerLayout>,
    /// `cs:dataPointWireframe`.
    pub data_point_wireframe: StyleEntry,
    /// `cs:dataTable`.
    pub data_table: StyleEntry,
    /// `cs:downBar`.
    pub down_bar: StyleEntry,
    /// `cs:dropLine`.
    pub drop_line: StyleEntry,
    /// `cs:errorBar`.
    pub error_bar: StyleEntry,
    /// `cs:floor`.
    pub floor: StyleEntry,
    /// `cs:gridlineMajor`.
    pub gridline_major: StyleEntry,
    /// `cs:gridlineMinor`.
    pub gridline_minor: StyleEntry,
    /// `cs:hiLoLine`.
    pub hi_lo_line: StyleEntry,
    /// `cs:leaderLine`.
    pub leader_line: StyleEntry,
    /// `cs:legend`.
    pub legend: StyleEntry,
    /// `cs:plotArea`.
    pub plot_area: StyleEntry,
    /// `cs:plotArea3D`.
    pub plot_area_3d: StyleEntry,
    /// `cs:seriesAxis`.
    pub series_axis: StyleEntry,
    /// `cs:seriesLine`.
    pub series_line: StyleEntry,
    /// `cs:title`.
    pub title: StyleEntry,
    /// `cs:trendline`.
    pub trendline: StyleEntry,
    /// `cs:trendlineLabel`.
    pub trendline_label: StyleEntry,
    /// `cs:upBar`.
    pub up_bar: StyleEntry,
    /// `cs:valueAxis`.
    pub value_axis: StyleEntry,
    /// `cs:wall`.
    pub wall: StyleEntry,
    /// `cs:extLst`, the complete element, verbatim.
    pub extensions: Option<Vec<u8>>,
}

/// How a colour style walks its palette (`meth`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum ColorMethod {
    /// Take the next colour for each series, wrapping round.
    #[default]
    Cycle,
    /// Shade one colour across the series.
    WithinLinear,
    /// Shade one colour across the points of a series.
    AcrossLinear,
    /// As `WithinLinear`, reversed.
    WithinLinearReversed,
    /// As `AcrossLinear`, reversed.
    AcrossLinearReversed,
    /// A method this crate does not model, kept as written.
    Other(String),
}

impl ColorMethod {
    /// The `meth` attribute value.
    pub fn as_str(&self) -> &str {
        match self {
            Self::Cycle => "cycle",
            Self::WithinLinear => "withinLinear",
            Self::AcrossLinear => "acrossLinear",
            Self::WithinLinearReversed => "withinLinearReversed",
            Self::AcrossLinearReversed => "acrossLinearReversed",
            Self::Other(value) => value,
        }
    }

    /// Parse a `meth` attribute value.
    pub fn from_str_value(value: &str) -> Self {
        match value {
            "cycle" => Self::Cycle,
            "withinLinear" => Self::WithinLinear,
            "acrossLinear" => Self::AcrossLinear,
            "withinLinearReversed" => Self::WithinLinearReversed,
            "acrossLinearReversed" => Self::AcrossLinearReversed,
            other => Self::Other(other.to_string()),
        }
    }
}

/// A chart colour style part (`xl/charts/colorsN.xml`).
///
/// `CT_ColorStyle` requires `meth` and at least one colour.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartColorStyle {
    /// How the palette is walked.
    pub method: ColorMethod,
    /// Which entry of Excel's colour gallery this is.
    pub id: Option<u32>,
    /// The palette, each colour the complete element, verbatim.
    pub colors: Vec<Vec<u8>>,
    /// `cs:variation` elements, each complete, verbatim.
    pub variations: Vec<Vec<u8>>,
}

/// A chart style part, as read or as built.
///
/// A part read from a file that this crate cannot model is kept as its
/// bytes rather than rejected or coerced, so reading stays permissive
/// and a round trip stays faithful. Anything built through the typed
/// model is a part Excel will accept by construction.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartStylePart {
    /// Modelled.
    Typed(Box<ChartStyle>),
    /// Kept as read, because it did not conform.
    Raw(Vec<u8>),
}

/// A chart colour style part, as read or as built. See
/// [`ChartStylePart`].
#[derive(Debug, Clone, PartialEq)]
pub enum ChartColorStylePart {
    /// Modelled.
    Typed(ChartColorStyle),
    /// Kept as read, because it did not conform.
    Raw(Vec<u8>),
}

/// Hairline in the "15% tint of text 1" that Excel uses for chart
/// borders, gridlines and axis lines.
const SUBTLE_LINE: &str = concat!(
    r#"<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>"#,
    r#"<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>"#,
    r#"</a:schemeClr></a:solidFill><a:round/></a:ln>"#
);

/// Text 1 at 65% luminance, Excel's colour for secondary chart text.
const MUTED_TEXT: &str = concat!(
    r#"<a:schemeClr val="tx1"><a:lumMod val="65000"/>"#,
    r#"<a:lumOff val="35000"/></a:schemeClr>"#
);

/// A series colour taken from the colour style part: `styleClr` picks the
/// next colour in the cycle and `phClr` resolves to it.
const CYCLE_COLOR: &str = r#"<cs:styleClr val="auto"/>"#;
const CYCLE_FILL: &str = r#"<a:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:spPr>"#;

/// Style id emitted on the generated default.
///
/// The id selects which entry of Excel's gallery is shown as current; it
/// does not affect how the chart renders, since the entries carry the
/// formatting. Excel accepts any value but requires the attribute.
const DEFAULT_STYLE_ID: u32 = 201;

fn entry_with_line(line: &str) -> StyleEntry {
    StyleEntry {
        shape_properties: Some(format!("<a:spPr>{line}</a:spPr>").into_bytes()),
        ..StyleEntry::default()
    }
}

fn text_entry(size: u32) -> StyleEntry {
    StyleEntry {
        font_reference: FontReference {
            collection: FontCollection::Minor,
            color: Some(MUTED_TEXT.as_bytes().to_vec()),
        },
        default_run_properties: Some(format!(r#"<a:defRPr sz="{size}"/>"#).into_bytes()),
        ..StyleEntry::default()
    }
}

fn cycle_entry(shape_properties: &str) -> StyleEntry {
    StyleEntry {
        fill_reference: StyleReference {
            idx: 0,
            color: Some(CYCLE_COLOR.as_bytes().to_vec()),
        },
        shape_properties: Some(shape_properties.as_bytes().to_vec()),
        ..StyleEntry::default()
    }
}

impl Default for ChartStyle {
    /// The style a chart gets when none was read from a file.
    ///
    /// Modelled on what Excel emits: the entries that visibly matter
    /// carry its formatting, and the rest defer to the theme.
    fn default() -> Self {
        let subtle_text = |size| StyleEntry {
            shape_properties: Some(format!("<a:spPr>{SUBTLE_LINE}</a:spPr>").into_bytes()),
            ..text_entry(size)
        };
        Self {
            id: DEFAULT_STYLE_ID,
            axis_title: text_entry(900),
            category_axis: subtle_text(900),
            chart_area: StyleEntry {
                mods: Some("allowNoFillOverride allowNoLineOverride".into()),
                shape_properties: Some(
                    format!(
                        r#"<a:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>{SUBTLE_LINE}</a:spPr>"#
                    )
                    .into_bytes(),
                ),
                default_run_properties: Some(br#"<a:defRPr sz="1000"/>"#.to_vec()),
                ..StyleEntry::default()
            },
            data_label: text_entry(900),
            data_label_callout: Some(subtle_text(900)),
            data_point: cycle_entry(CYCLE_FILL),
            data_point_3d: cycle_entry(CYCLE_FILL),
            data_point_line: cycle_entry(concat!(
                r#"<a:spPr><a:ln w="28575" cap="rnd"><a:solidFill>"#,
                r#"<a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></a:spPr>"#
            )),
            data_point_marker: cycle_entry(concat!(
                r#"<a:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
                r#"<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:spPr>"#
            )),
            data_point_marker_layout: Some(MarkerLayout {
                symbol: Some("circle".into()),
                size: Some(5),
            }),
            data_point_wireframe: cycle_entry(concat!(
                r#"<a:spPr><a:ln w="9525"><a:solidFill>"#,
                r#"<a:schemeClr val="phClr"/></a:solidFill></a:ln></a:spPr>"#
            )),
            data_table: subtle_text(900),
            down_bar: StyleEntry::theme_default(),
            drop_line: entry_with_line(SUBTLE_LINE),
            error_bar: entry_with_line(SUBTLE_LINE),
            floor: StyleEntry::theme_default(),
            gridline_major: entry_with_line(SUBTLE_LINE),
            gridline_minor: entry_with_line(SUBTLE_LINE),
            hi_lo_line: entry_with_line(SUBTLE_LINE),
            leader_line: entry_with_line(SUBTLE_LINE),
            legend: text_entry(900),
            plot_area: StyleEntry {
                mods: Some("allowNoFillOverride allowNoLineOverride".into()),
                ..StyleEntry::default()
            },
            plot_area_3d: StyleEntry {
                mods: Some("allowNoFillOverride allowNoLineOverride".into()),
                ..StyleEntry::default()
            },
            series_axis: subtle_text(900),
            series_line: entry_with_line(SUBTLE_LINE),
            title: text_entry(1400),
            trendline: entry_with_line(SUBTLE_LINE),
            trendline_label: text_entry(900),
            up_bar: StyleEntry::theme_default(),
            value_axis: text_entry(900),
            wall: StyleEntry::theme_default(),
            extensions: None,
        }
    }
}

impl Default for ChartColorStyle {
    /// The "colourful" palette Excel applies to a new chart: cycle
    /// through the six theme accents, then luminance variations of them.
    fn default() -> Self {
        Self {
            method: ColorMethod::Cycle,
            id: Some(10),
            colors: (1..=6)
                .map(|i| format!(r#"<a:schemeClr val="accent{i}"/>"#).into_bytes())
                .collect(),
            variations: [
                "",
                r#"<a:lumMod val="60000"/>"#,
                r#"<a:lumMod val="80000"/><a:lumOff val="20000"/>"#,
                r#"<a:lumMod val="80000"/>"#,
                r#"<a:lumMod val="60000"/><a:lumOff val="40000"/>"#,
                r#"<a:lumMod val="50000"/>"#,
                r#"<a:lumMod val="70000"/><a:lumOff val="30000"/>"#,
                r#"<a:lumMod val="70000"/>"#,
                r#"<a:lumMod val="50000"/><a:lumOff val="50000"/>"#,
            ]
            .iter()
            .map(|inner| {
                if inner.is_empty() {
                    b"<cs:variation/>".to_vec()
                } else {
                    format!("<cs:variation>{inner}</cs:variation>").into_bytes()
                }
            })
            .collect(),
        }
    }
}

/// Every entry of a style, paired with the element name it belongs to,
/// in `CT_ChartStyle` order. Bindings surface entries by name rather
/// than as thirty-one fields, and this keeps that mapping in one place.
pub fn entries_by_name(style: &ChartStyle) -> Vec<(&'static str, &StyleEntry)> {
    let mut out = vec![
        ("axisTitle", &style.axis_title),
        ("categoryAxis", &style.category_axis),
        ("chartArea", &style.chart_area),
        ("dataLabel", &style.data_label),
    ];
    if let Some(ref callout) = style.data_label_callout {
        out.push(("dataLabelCallout", callout));
    }
    out.extend([
        ("dataPoint", &style.data_point),
        ("dataPoint3D", &style.data_point_3d),
        ("dataPointLine", &style.data_point_line),
        ("dataPointMarker", &style.data_point_marker),
        ("dataPointWireframe", &style.data_point_wireframe),
        ("dataTable", &style.data_table),
        ("downBar", &style.down_bar),
        ("dropLine", &style.drop_line),
        ("errorBar", &style.error_bar),
        ("floor", &style.floor),
        ("gridlineMajor", &style.gridline_major),
        ("gridlineMinor", &style.gridline_minor),
        ("hiLoLine", &style.hi_lo_line),
        ("leaderLine", &style.leader_line),
        ("legend", &style.legend),
        ("plotArea", &style.plot_area),
        ("plotArea3D", &style.plot_area_3d),
        ("seriesAxis", &style.series_axis),
        ("seriesLine", &style.series_line),
        ("title", &style.title),
        ("trendline", &style.trendline),
        ("trendlineLabel", &style.trendline_label),
        ("upBar", &style.up_bar),
        ("valueAxis", &style.value_axis),
        ("wall", &style.wall),
    ]);
    out
}
