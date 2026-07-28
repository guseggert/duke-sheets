//! # duke-sheets-chart
//!
//! Chart support for duke-sheets.

#[cfg(feature = "parse")]
pub mod error;
#[cfg(feature = "parse")]
pub mod parse;

#[cfg(any(feature = "parse", feature = "write"))]
pub mod drawing_part;
#[cfg(feature = "write")]
pub mod write;

mod axis;
mod chart;
pub mod chart_ex;
pub mod raw_rel;
mod config;
mod data_labels;
mod drawing;
mod error_bars;
mod formatting;
mod legend;
mod marker;
mod series;
mod text_properties;
mod trendline;
mod types;

pub use axis::{Axis, AxisCrosses, AxisPosition, AxisType, CrossBetween, TickLabelPosition, TickMark};
pub use chart::{
    BarShape, CellMarker, Chart, ChartAxis, ChartLines, ChartType, ChartTypeGroup,
    DrawingAnchor, EditAs, EmbeddedImage, ImageFormat, OfPieType, SplitType, Surface, UpDownBars,
};
pub use config::{ChartDataTable, DisplayBlanksAs, Layout, ManualLayout, View3D};
pub use drawing::{
    column_width_to_emu, marker_at_emu, marker_position_emu, row_height_to_emu, ChildTransform,
    DefaultDrawingMetrics, DrawingMetrics, GroupTransform, EMU_PER_PIXEL, EMU_PER_POINT,
};
pub use data_labels::{DataLabel, DataLabelPosition, DataLabels, DataPoint};
pub use error_bars::{ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType};
pub use formatting::{ChartColor, ChartLine, ChartShapeProperties, NumberFormat, PictureOptions};
pub use legend::{Legend, LegendEntry, LegendPosition};
pub use marker::{Marker, MarkerSymbol};
pub use series::{DataReference, DataSeries};
pub use text_properties::{TextAnchor, TextProperties, TextVertical, TextWrap};
pub use trendline::{Trendline, TrendlineLabel, TrendlineType};
pub use raw_rel::RawRel;
pub use chart_ex::{
    ChartEx, ChartExAxis, ChartExAxisTitle, ChartExAxisUnits, ChartExAxisUnitsLabel,
    ChartExBinning, ChartExColorPosition, ChartExData, ChartExDataLabel, ChartExDataLabels,
    ChartExDataPoint, ChartExDimension, ChartExExternalData, ChartExFormatOverride,
    ChartExGeography, ChartExGridlines, ChartExHeaderFooter, ChartExLayout, ChartExLayoutPr,
    ChartExLegend,
    ChartExNumericLevel, ChartExOffset, ChartExPageMargins, ChartExPageSetup, ChartExPlotArea,
    ChartExPrintSettings, ChartExScaling, ChartExSeries, ChartExSeriesVisibility,
    ChartExStatistics, ChartExStringLevel, ChartExText, ChartExTextData, ChartExTitle,
    ChartExValueColorPositions, ChartExValueColors, NumericDimType, StringDimType,
};
