//! Shared SpreadsheetML drawing-part (`xl/drawings/drawingN.xml`)
//! codec used by the XLSX and XLSB crates.
//!
//! The reader ([`parse_drawing_part`], feature `parse`) classifies
//! each top-level anchor in document order; document order carries
//! z-order. The writer (feature `write`) emits a drawing part from a
//! caller-built object list, with the control-twin flavor selected by
//! [`write::TwinStyle`]: XLSX uses `a14:compatExt` `<xdr:sp>` twins,
//! XLSB uses `com14:compatSp` `<xdr:graphicFrame>` twins. The reader
//! recognizes both flavors regardless of the container format.

#[cfg(feature = "parse")]
pub mod read;
#[cfg(feature = "write")]
pub mod write;

/// DrawingML text carried by an XLSX legacy-control twin.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TwinText {
    pub runs: Vec<TwinTextRun>,
    pub horizontal_alignment: Option<TwinHorizontalAlignment>,
    pub vertical_alignment: Option<TwinVerticalAlignment>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TwinTextRun {
    pub text: String,
    pub font: Option<TwinRunFont>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TwinRunFont {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<TwinColor>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<TwinUnderline>,
    pub strikethrough: Option<bool>,
    /// DrawingML baseline percentage in thousandths of a percent.
    pub baseline: Option<i32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TwinColor {
    Rgb { r: u8, g: u8, b: u8 },
    Theme { index: u8, tint: i8 },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TwinUnderline {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TwinHorizontalAlignment {
    Left,
    Center,
    Right,
    Justify,
    Distributed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TwinVerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Fill representation shared by the XLSX/XLSB DrawingML adapters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShapeFill {
    #[default]
    None,
    Solid(TwinColor),
}

/// Line representation shared by the XLSX/XLSB DrawingML adapters.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ShapeLine {
    pub color: Option<TwinColor>,
    pub width_emu: Option<i64>,
    pub dash_style: Option<String>,
    pub no_fill: bool,
}

/// OOXML relationship type for chart parts.
pub const RT_CHART: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
/// OOXML relationship type for chartEx parts.
pub const RT_CHART_EX: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
/// OOXML relationship type for embedded image parts (`xl/media/*`).
pub const RT_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// Map an `ImageFormat` to the file extension used in `xl/media/`.
pub fn image_format_extension(fmt: crate::ImageFormat) -> &'static str {
    use crate::ImageFormat;
    match fmt {
        ImageFormat::Png => "png",
        ImageFormat::Jpeg => "jpeg",
        ImageFormat::Gif => "gif",
        ImageFormat::Bmp => "bmp",
        ImageFormat::Tiff => "tiff",
        ImageFormat::Emf => "emf",
        ImageFormat::Wmf => "wmf",
        ImageFormat::Svg => "svg",
    }
}

/// The IANA MIME type for an image format in `[Content_Types].xml`.
pub fn image_format_mime(fmt: crate::ImageFormat) -> &'static str {
    use crate::ImageFormat;
    match fmt {
        ImageFormat::Png => "image/png",
        ImageFormat::Jpeg => "image/jpeg",
        ImageFormat::Gif => "image/gif",
        ImageFormat::Bmp => "image/bmp",
        ImageFormat::Tiff => "image/tiff",
        ImageFormat::Emf => "image/x-emf",
        ImageFormat::Wmf => "image/x-wmf",
        ImageFormat::Svg => "image/svg+xml",
    }
}
