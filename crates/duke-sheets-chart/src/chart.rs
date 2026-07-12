//! Chart types

use std::collections::HashMap;

use crate::axis::Axis;
use crate::config::{ChartDataTable, DisplayBlanksAs, Layout, View3D};
use crate::data_labels::DataLabels;
use crate::formatting::ChartShapeProperties;
use crate::legend::Legend;
use crate::series::DataSeries;
use crate::text_properties::TextProperties;

/// Chart types
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartType {
    // Column/Bar
    ColumnClustered,
    ColumnStacked,
    ColumnPercentStacked,
    BarClustered,
    BarStacked,
    BarPercentStacked,

    // Line
    Line,
    LineStacked,

    // Pie
    Pie,
    PieExploded,
    Doughnut,

    // Area
    Area,
    AreaStacked,
    AreaPercentStacked,

    // Scatter
    ScatterMarkers,
    ScatterSmooth,
    ScatterLines,

    // Other
    Bubble,
    Radar,
    Stock,
    Surface,

    /// Imported chart type not yet mapped to a known variant.
    /// Contains the original OOXML element tag (e.g. "c:surface3DChart").
    Unsupported(String),
}

/// Chart line overlay (drop lines, high-low lines, series lines, leader lines).
/// All share the same structure: optional shape properties for formatting.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartLines {
    pub shape_properties: Option<ChartShapeProperties>,
}

/// Up-down bars (stock charts).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct UpDownBars {
    pub gap_width: Option<u32>,
    pub up_bars: Option<ChartLines>,
    pub down_bars: Option<ChartLines>,
}

/// Represents one chart type block (e.g. barChart, lineChart) within a plotArea.
/// Used for combo charts where multiple chart types share axes.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartTypeGroup {
    pub chart_type: ChartType,
    pub is_3d: bool,
    pub series: Vec<DataSeries>,
    pub data_labels: Option<DataLabels>,
    pub vary_colors: Option<bool>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub wireframe: Option<bool>,
    pub drop_lines: Option<ChartLines>,
    pub high_low_lines: Option<ChartLines>,
    pub series_lines: Option<ChartLines>,
    pub up_down_bars: Option<UpDownBars>,
    /// axId values this group references (e.g. [1, 2] or [1, 3])
    pub of_pie_type: Option<OfPieType>,
    pub split_type: Option<SplitType>,
    pub split_pos: Option<f64>,
    pub second_pie_size: Option<u32>,
    pub bar_shape: Option<BarShape>,
    pub floor: Option<Surface>,
    pub side_wall: Option<Surface>,
    pub back_wall: Option<Surface>,
    pub axis_ids: Vec<u32>,
    #[doc(hidden)]
    pub raw_ext: Option<Vec<u8>>,
}

/// Represents one axis in the plotArea with its ID and cross-reference.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartAxis {
    pub id: u32,
    pub cross_id: u32,
    pub axis: Axis,
}

/// Chart definition
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    /// Chart type
    pub chart_type: ChartType,
    /// Chart title
    pub title: Option<String>,
    /// Data series
    pub series: Vec<DataSeries>,
    /// Category axis (X)
    pub category_axis: Option<Axis>,
    /// Value axis (Y)
    pub value_axis: Option<Axis>,
    /// Series axis (Z, for 3D charts)
    pub series_axis: Option<Axis>,
    /// Legend
    pub legend: Option<Legend>,
    pub data_labels: Option<DataLabels>,
    pub view_3d: Option<View3D>,
    pub data_table: Option<ChartDataTable>,
    pub display_blanks_as: Option<DisplayBlanksAs>,
    pub plot_visible_only: Option<bool>,
    pub layout: Option<Layout>,
    pub shape_properties: Option<ChartShapeProperties>,
    pub vary_colors: Option<bool>,
    pub is_3d: bool,
    pub first_slice_angle: Option<u32>,
    pub hole_size: Option<u32>,
    pub bubble_scale: Option<u32>,
    pub show_negative_bubbles: Option<bool>,
    pub radar_style: Option<String>,
    pub auto_title_deleted: Option<bool>,
    pub rounded_corners: Option<bool>,
    pub show_dlbls_over_max: Option<bool>,
    pub wireframe: Option<bool>,
    pub drop_lines: Option<ChartLines>,
    pub high_low_lines: Option<ChartLines>,
    pub up_down_bars: Option<UpDownBars>,
    pub series_lines: Option<ChartLines>,
    pub gap_width: Option<u32>,
    pub overlap: Option<i32>,
    pub show_marker: Option<bool>,
    pub of_pie_type: Option<OfPieType>,
    pub split_type: Option<SplitType>,
    pub split_pos: Option<f64>,
    pub second_pie_size: Option<u32>,
    pub bar_shape: Option<BarShape>,
    pub floor: Option<Surface>,
    pub side_wall: Option<Surface>,
    pub back_wall: Option<Surface>,
    pub text_properties: Option<TextProperties>,
    /// Raw XML fragments for extension lists we don't parse but need to preserve.
    /// Key is the context path (e.g. "chartSpace", "chart", "plotArea", "chartType").
    #[doc(hidden)]
    pub raw_extensions: HashMap<String, Vec<u8>>,
    /// Raw chart style XML (Office 2013+), preserved for roundtrip.
    #[doc(hidden)]
    pub raw_chart_style: Option<Vec<u8>>,
    /// Raw chart color style XML (Office 2013+), preserved for roundtrip.
    #[doc(hidden)]
    pub raw_chart_color_style: Option<Vec<u8>>,
    /// Chart type groups for combo charts (multiple chart types in one plotArea).
    /// When non-empty, the writer uses these instead of the legacy single-group fields.
    pub type_groups: Vec<ChartTypeGroup>,
    /// Axes for combo charts, each with its ID and cross-reference.
    pub axes: Vec<ChartAxis>,
}

impl Chart {
    /// Create a new chart
    pub fn new(chart_type: ChartType) -> Self {
        Self {
            chart_type,
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
            show_marker: None,
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            text_properties: None,
            raw_extensions: HashMap::new(),
            raw_chart_style: None,
            raw_chart_color_style: None,
            type_groups: Vec::new(),
            axes: Vec::new(),
        }
    }

    /// Set chart title
    pub fn with_title<S: Into<String>>(mut self, title: S) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Add a data series
    pub fn add_series(&mut self, series: DataSeries) {
        self.series.push(series);
    }
}

/// A cell position marker for drawing anchors.
///
/// Cell positions use zero-based column and row indices.
/// Offsets are in EMU (English Metric Units, 1 inch = 914400 EMU)
/// and represent sub-cell positioning within the cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CellMarker {
    /// Column (zero-based)
    pub col: u16,
    /// Offset within column (EMU)
    pub col_offset_emu: i64,
    /// Row (zero-based)
    pub row: u32,
    /// Offset within row (EMU)
    pub row_offset_emu: i64,
}

/// How the drawing should be resized when cells are resized.
#[derive(Debug, Clone, PartialEq)]
pub enum EditAs {
    TwoCell,
    OneCell,
    Absolute,
}

/// Drawing anchor position in a worksheet.
#[derive(Debug, Clone, PartialEq)]
pub enum DrawingAnchor {
    TwoCell {
        from: CellMarker,
        to: CellMarker,
        edit_as: Option<EditAs>,
    },
    OneCell {
        from: CellMarker,
        width_emu: i64,
        height_emu: i64,
    },
    Absolute {
        x_emu: i64,
        y_emu: i64,
        width_emu: i64,
        height_emu: i64,
    },
}

impl Default for DrawingAnchor {
    fn default() -> Self {
        DrawingAnchor::TwoCell {
            from: CellMarker::default(),
            to: CellMarker::default(),
            edit_as: None,
        }
    }
}

impl DrawingAnchor {
    /// The canonical two-cell form of any anchor. OneCell and
    /// Absolute anchors are flattened to from/to markers at Excel's
    /// default cell metrics (609,600 EMU per column, 190,500 EMU per
    /// row), with `edit_as` preserving the original sizing behavior.
    /// Markers clamp to the grid limits, and euclidean division keeps
    /// offsets in `0..metric`, so off-grid negative coordinates
    /// normalize onto the grid origin. TwoCell anchors are returned
    /// unchanged.
    pub fn to_two_cell(&self) -> DrawingAnchor {
        self.to_two_cell_with_metrics(&crate::DefaultDrawingMetrics)
    }

    /// The canonical two-cell form of this anchor using the worksheet's
    /// actual row heights and column widths. TwoCell anchors are returned
    /// unchanged; OneCell and Absolute anchors retain their sizing behavior
    /// in `edit_as` while their endpoint markers are resolved through
    /// `metrics`.
    pub fn to_two_cell_with_metrics(
        &self,
        metrics: &(impl crate::DrawingMetrics + ?Sized),
    ) -> DrawingAnchor {
        match self {
            DrawingAnchor::TwoCell { .. } => self.clone(),
            DrawingAnchor::OneCell {
                from,
                width_emu,
                height_emu,
            } => {
                let (x, y) = crate::marker_position_emu(from, metrics);
                DrawingAnchor::TwoCell {
                    from: from.clone(),
                    to: crate::marker_at_emu(
                        x + i128::from((*width_emu).max(0)),
                        y + i128::from((*height_emu).max(0)),
                        metrics,
                    ),
                    edit_as: Some(EditAs::OneCell),
                }
            }
            DrawingAnchor::Absolute {
                x_emu,
                y_emu,
                width_emu,
                height_emu,
            } => DrawingAnchor::TwoCell {
                from: crate::marker_at_emu(i128::from(*x_emu), i128::from(*y_emu), metrics),
                to: crate::marker_at_emu(
                    i128::from(*x_emu) + i128::from((*width_emu).max(0)),
                    i128::from(*y_emu) + i128::from((*height_emu).max(0)),
                    metrics,
                ),
                edit_as: Some(EditAs::Absolute),
            },
        }
    }
}

/// Image format for embedded images.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Emf,
    Wmf,
    Tiff,
    Svg,
}

impl ImageFormat {
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_lowercase().as_str() {
            "png" => Some(Self::Png),
            "jpg" | "jpeg" => Some(Self::Jpeg),
            "gif" => Some(Self::Gif),
            "bmp" => Some(Self::Bmp),
            "emf" => Some(Self::Emf),
            "wmf" => Some(Self::Wmf),
            "tif" | "tiff" => Some(Self::Tiff),
            "svg" => Some(Self::Svg),
            _ => None,
        }
    }

    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
            Self::Bmp => "bmp",
            Self::Emf => "emf",
            Self::Wmf => "wmf",
            Self::Tiff => "tiff",
            Self::Svg => "svg",
        }
    }
}

/// An image embedded in a worksheet drawing.
///
/// Placement, shape name, and description (alt text) live on the
/// wrapping drawing object in `duke-sheets-core`.
#[derive(Debug, Clone, PartialEq)]
pub struct EmbeddedImage {
    pub format: ImageFormat,
    pub media_path: String,
    pub svg_media_path: Option<String>,
    pub width_emu: i64,
    pub height_emu: i64,
    pub rotation: Option<i32>,
    pub flip_h: bool,
    pub flip_v: bool,
    /// Raw image bytes (PNG, JPEG, etc.), loaded from the archive at parse time.
    #[doc(hidden)]
    pub data: Vec<u8>,
    /// SVG image bytes when the image has an SVG variant.
    #[doc(hidden)]
    pub svg_data: Option<Vec<u8>>,
}

impl EmbeddedImage {
    /// Raw image bytes (PNG, JPEG, EMF, etc.).
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// SVG image bytes, if the image has an SVG variant.
    pub fn svg_data(&self) -> Option<&[u8]> {
        self.svg_data.as_deref()
    }
}

/// 3D chart surface (floor, side wall, back wall)
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Surface {
    pub thickness: Option<u32>,
    pub shape_properties: Option<ChartShapeProperties>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfPieType {
    Pie,
    Bar,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SplitType {
    Auto,
    Custom,
    Percent,
    Position,
    Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarShape {
    Box,
    Cone,
    ConeToMax,
    Cylinder,
    Pyramid,
    PyramidToMax,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn two_cell_anchor_is_returned_unchanged() {
        let anchor = DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 100,
                row: 2,
                row_offset_emu: 200,
            },
            to: CellMarker {
                col: 3,
                col_offset_emu: 300,
                row: 4,
                row_offset_emu: 400,
            },
            edit_as: Some(EditAs::OneCell),
        };
        assert_eq!(anchor.to_two_cell(), anchor);
    }

    #[test]
    fn one_cell_anchor_flattens_at_default_metrics() {
        let anchor = DrawingAnchor::OneCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 9_600,
                row: 1,
                row_offset_emu: 500,
            },
            width_emu: 1_219_200, // exactly two columns
            height_emu: 381_000,  // exactly two rows
        };
        match anchor.to_two_cell() {
            DrawingAnchor::TwoCell { from, to, edit_as } => {
                assert_eq!((from.col, from.row), (2, 1));
                assert_eq!((to.col, to.col_offset_emu), (4, 9_600));
                assert_eq!((to.row, to.row_offset_emu), (3, 500));
                assert_eq!(edit_as, Some(EditAs::OneCell));
            }
            other => panic!("expected TwoCell, got {other:?}"),
        }
    }

    #[test]
    fn absolute_anchor_flattens_at_default_metrics() {
        let anchor = DrawingAnchor::Absolute {
            x_emu: 1_219_300,
            y_emu: 190_600,
            width_emu: 609_600,
            height_emu: 190_500,
        };
        match anchor.to_two_cell() {
            DrawingAnchor::TwoCell { from, to, edit_as } => {
                assert_eq!((from.col, from.col_offset_emu), (2, 100));
                assert_eq!((from.row, from.row_offset_emu), (1, 100));
                assert_eq!((to.col, to.col_offset_emu), (3, 100));
                assert_eq!((to.row, to.row_offset_emu), (2, 100));
                assert_eq!(edit_as, Some(EditAs::Absolute));
            }
            other => panic!("expected TwoCell, got {other:?}"),
        }
    }

    #[test]
    fn negative_absolute_coordinates_normalize_euclidean() {
        // Off-grid negative coordinates from permissively-read files
        // normalize with euclidean division: the marker clamps to the
        // grid origin and the offset stays in 0..metric.
        let anchor = DrawingAnchor::Absolute {
            x_emu: -1_000,
            y_emu: -1_000,
            width_emu: 0,
            height_emu: 0,
        };
        match anchor.to_two_cell() {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!((from.col, from.col_offset_emu), (0, 608_600));
                assert_eq!((from.row, from.row_offset_emu), (0, 189_500));
                assert_eq!(to, from);
            }
            other => panic!("expected TwoCell, got {other:?}"),
        }
    }

    #[test]
    fn extreme_anchor_values_clamp_without_panicking() {
        // Offsets near i64::MAX must not overflow the flattening
        // arithmetic; markers clamp to the grid limits.
        let hostile = DrawingAnchor::OneCell {
            from: CellMarker {
                col: u16::MAX,
                col_offset_emu: i64::MAX,
                row: u32::MAX,
                row_offset_emu: i64::MAX,
            },
            width_emu: i64::MAX,
            height_emu: i64::MAX,
        };
        match hostile.to_two_cell() {
            DrawingAnchor::TwoCell { to, .. } => {
                assert_eq!(to.col, u16::MAX);
                assert_eq!(to.row, u32::MAX);
                assert!((0..609_600).contains(&to.col_offset_emu));
                assert!((0..190_500).contains(&to.row_offset_emu));
            }
            other => panic!("expected TwoCell, got {other:?}"),
        }

        let hostile_absolute = DrawingAnchor::Absolute {
            x_emu: i64::MIN,
            y_emu: i64::MIN,
            width_emu: i64::MAX,
            height_emu: i64::MAX,
        };
        match hostile_absolute.to_two_cell() {
            DrawingAnchor::TwoCell { from, .. } => {
                assert_eq!((from.col, from.row), (0, 0));
            }
            other => panic!("expected TwoCell, got {other:?}"),
        }
    }
}
