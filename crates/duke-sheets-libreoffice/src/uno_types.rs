//! UNO application-level types for LibreOffice spreadsheet automation.
//!
//! This module defines Rust representations of UNO structs and enum constants
//! used by the spreadsheet API. Each struct has a `to_uno()` method that converts
//! it into a `UnoValue::Struct(...)` suitable for sending over URP.
//!
//! These types live in the higher-level crate (not in libreoffice-urp) because
//! they are application-specific, not protocol-level.

use libreoffice_urp::types::UnoValue;

// ============================================================================
// UNO Structs
// ============================================================================

/// `com.sun.star.table.CellAddress` — identifies a single cell by sheet, column, row.
///
/// Wire format: Struct(Short, Long, Long)
#[derive(Debug, Clone, Copy)]
pub struct CellAddress {
    pub sheet: i16,
    pub column: i32,
    pub row: i32,
}

impl CellAddress {
    pub fn new(sheet: i16, column: i32, row: i32) -> Self {
        Self { sheet, column, row }
    }

    /// Convert to the URP wire representation.
    pub fn to_uno(&self) -> UnoValue {
        UnoValue::Struct(vec![
            UnoValue::Short(self.sheet),
            UnoValue::Long(self.column),
            UnoValue::Long(self.row),
        ])
    }
}

/// `com.sun.star.table.BorderLine2` — describes a border line on a cell.
///
/// Wire format: Struct(Long color, Short innerLineWidth, Short outerLineWidth,
///                     Short lineDistance, Short lineStyle, Long lineWidth)
///
/// Note: The UNO IDL has 6 fields. For our purposes, we set Color, LineStyle,
/// and LineWidth, leaving InnerLineWidth/OuterLineWidth/LineDistance as 0.
#[derive(Debug, Clone, Copy)]
pub struct BorderLine2 {
    pub color: i32,
    pub line_style: i16,
    pub line_width: i32,
}

impl BorderLine2 {
    pub fn new(color: i32, line_style: i16, line_width: i32) -> Self {
        Self {
            color,
            line_style,
            line_width,
        }
    }

    /// Create a "no border" value.
    pub fn none() -> Self {
        Self {
            color: 0,
            line_style: border_line_style::NONE,
            line_width: 0,
        }
    }

    /// Convert to the URP wire representation.
    ///
    /// Fields in order: Color, InnerLineWidth, OuterLineWidth, LineDistance, LineStyle, LineWidth
    pub fn to_uno(&self) -> UnoValue {
        UnoValue::Struct(vec![
            UnoValue::Long(self.color),       // Color
            UnoValue::Short(0),               // InnerLineWidth
            UnoValue::Short(0),               // OuterLineWidth
            UnoValue::Short(0),               // LineDistance
            UnoValue::Short(self.line_style), // LineStyle
            UnoValue::Long(self.line_width),  // LineWidth
        ])
    }
}

/// `com.sun.star.lang.Locale` — specifies a language/country locale.
///
/// Wire format: Struct(String language, String country, String variant)
///
/// An empty Locale (all empty strings) is the default used for number format queries.
#[derive(Debug, Clone)]
pub struct Locale {
    pub language: String,
    pub country: String,
    pub variant: String,
}

impl Locale {
    /// Create an empty (default) locale.
    pub fn empty() -> Self {
        Self {
            language: String::new(),
            country: String::new(),
            variant: String::new(),
        }
    }

    /// Create a locale with language and country.
    pub fn new(language: &str, country: &str) -> Self {
        Self {
            language: language.to_string(),
            country: country.to_string(),
            variant: String::new(),
        }
    }

    /// Convert to the URP wire representation.
    pub fn to_uno(&self) -> UnoValue {
        UnoValue::Struct(vec![
            UnoValue::String(self.language.clone()),
            UnoValue::String(self.country.clone()),
            UnoValue::String(self.variant.clone()),
        ])
    }
}

// ============================================================================
// UNO Enum Constants
// ============================================================================

/// `com.sun.star.awt.FontWeight` — font weight constants.
///
/// These are float values in UNO but sent as property values (typically as floats).
/// The Python code uses raw numeric value 150 for BOLD.
pub mod font_weight {
    /// Normal weight.
    pub const NORMAL: f32 = 100.0;
    /// Bold weight.
    pub const BOLD: f32 = 150.0;
}

/// `com.sun.star.awt.FontSlant` — font slant (posture) constants.
///
/// These are enum ordinal values sent as integers.
pub mod font_slant {
    /// No slant (upright).
    pub const NONE: i16 = 0;
    /// Italic.
    pub const ITALIC: i16 = 2;
}

/// `com.sun.star.awt.FontUnderline` — underline style constants.
pub mod font_underline {
    pub const NONE: i16 = 0;
    pub const SINGLE: i16 = 1;
    pub const DOUBLE: i16 = 2;
    pub const SINGLE_ACCOUNTING: i16 = 3;
    pub const DOUBLE_ACCOUNTING: i16 = 4;

    /// Look up underline style by name (matching Python UNDERLINE_STYLES map).
    pub fn from_name(name: &str) -> i16 {
        match name {
            "none" => NONE,
            "single" => SINGLE,
            "double" => DOUBLE,
            "singleAccounting" => SINGLE_ACCOUNTING,
            "doubleAccounting" => DOUBLE_ACCOUNTING,
            _ => SINGLE,
        }
    }
}

/// `com.sun.star.awt.FontStrikeout` — strikethrough constants.
pub mod font_strikeout {
    pub const NONE: i16 = 0;
    pub const SINGLE: i16 = 1;
}

/// `com.sun.star.table.CellHoriJustify` — horizontal alignment constants.
///
/// IDL enum ordinals: STANDARD=0, LEFT=1, CENTER=2, RIGHT=3, BLOCK=4, REPEAT=5.
///
/// The Python fixture framework maps additional names (fill, justify,
/// center_continuous, distributed) to integer values 4-7 that LO accepts
/// via setPropertyValue. We match the Python mapping for XLSX compatibility.
pub mod hori_justify {
    pub const STANDARD: i32 = 0;
    pub const LEFT: i32 = 1;
    pub const CENTER: i32 = 2;
    pub const RIGHT: i32 = 3;
    pub const FILL: i32 = 4;
    pub const JUSTIFY: i32 = 5;
    pub const CENTER_CONTINUOUS: i32 = 6;
    pub const DISTRIBUTED: i32 = 7;

    /// Look up horizontal alignment by name (matching Python HORIZONTAL_ALIGN map).
    pub fn from_name(name: &str) -> i32 {
        match name {
            "general" => STANDARD,
            "left" => LEFT,
            "center" => CENTER,
            "right" => RIGHT,
            "fill" => FILL,
            "justify" => JUSTIFY,
            "center_continuous" => CENTER_CONTINUOUS,
            "distributed" => DISTRIBUTED,
            _ => STANDARD,
        }
    }
}

/// `com.sun.star.table.CellVertJustify` — vertical alignment constants.
///
/// IDL enum ordinals: STANDARD=0, TOP=1, CENTER=2, BOTTOM=3.
///
/// The Python fixture framework maps additional names (justify, distributed)
/// to integer values 4-5 that LO accepts. We match the Python mapping.
pub mod vert_justify {
    pub const STANDARD: i32 = 0;
    pub const TOP: i32 = 1;
    pub const CENTER: i32 = 2;
    pub const BOTTOM: i32 = 3;
    pub const JUSTIFY: i32 = 4;
    pub const DISTRIBUTED: i32 = 5;

    /// Look up vertical alignment by name (matching Python VERTICAL_ALIGN map).
    pub fn from_name(name: &str) -> i32 {
        match name {
            "standard" => STANDARD,
            "top" => TOP,
            "center" => CENTER,
            "bottom" => BOTTOM,
            "justify" => JUSTIFY,
            "distributed" => DISTRIBUTED,
            _ => STANDARD,
        }
    }
}

/// `com.sun.star.table.BorderLineStyle` — border line style constants.
///
/// IDL constants group (NOT a sequential enum):
///   SOLID=0, DOTTED=1, DASHED=2, DOUBLE=3, DOUBLE_THIN=15, NONE=0x7FFF
///
/// The Python fixture framework uses a simplified mapping where the
/// `LineStyle` and `LineWidth` fields of `BorderLine2` together determine
/// the visual appearance. We match the Python framework behavior exactly
/// so that generated XLSX files are functionally equivalent.
pub mod border_line_style {
    // Actual IDL constants (com.sun.star.table.BorderLineStyle)
    pub const NONE: i16 = 0; // Used in from_name("none") with width=0
    pub const SOLID: i16 = 0;
    pub const DOTTED: i16 = 1;
    pub const DASHED: i16 = 2;
    pub const DOUBLE: i16 = 3;
    pub const DOUBLE_THIN: i16 = 15;
    pub const IDL_NONE: i16 = 0x7FFF_u16 as i16; // Actual IDL sentinel for "no border"

    /// Line widths corresponding to named styles (in 1/100 mm).
    pub const WIDTH_THIN: i32 = 50;
    pub const WIDTH_MEDIUM: i32 = 100;
    pub const WIDTH_THICK: i32 = 150;

    /// Look up border line style by name, matching the Python framework's
    /// `BORDER_STYLES` / `_apply_border` logic.
    ///
    /// The Python code maps style names → integer LineStyle values (0-6)
    /// and sets LineWidth separately. Returns (LineStyle, LineWidth) pair.
    pub fn from_name(name: &str) -> (i16, i32) {
        match name {
            "none" => (0, 0),
            "thin" => (SOLID, WIDTH_THIN),
            "medium" => (SOLID, WIDTH_MEDIUM),
            "thick" => (SOLID, WIDTH_THICK),
            "dashed" => (DASHED, WIDTH_THIN),
            "dotted" => (DOTTED, WIDTH_THIN),
            "double" => (DOUBLE, WIDTH_THIN),
            _ => (SOLID, WIDTH_THIN),
        }
    }
}

/// `com.sun.star.sheet.ConditionOperator` — conditional format operator constants.
///
/// These are UNO enum values (sent as i32 over URP).
pub mod condition_operator {
    pub const NONE: i32 = 0;
    pub const EQUAL: i32 = 1;
    pub const NOT_EQUAL: i32 = 2;
    pub const GREATER: i32 = 3;
    pub const GREATER_EQUAL: i32 = 4;
    pub const LESS: i32 = 5;
    pub const LESS_EQUAL: i32 = 6;
    pub const BETWEEN: i32 = 7;
    pub const NOT_BETWEEN: i32 = 8;
    pub const FORMULA: i32 = 9;

    /// Look up condition operator by name (matching Python CONDITION_OPERATORS map).
    pub fn from_name(name: &str) -> i32 {
        match name {
            "none" => NONE,
            "equal" => EQUAL,
            "not_equal" => NOT_EQUAL,
            "greater_than" => GREATER,
            "greater_or_equal" => GREATER_EQUAL,
            "less_than" => LESS,
            "less_or_equal" => LESS_EQUAL,
            "between" => BETWEEN,
            "not_between" => NOT_BETWEEN,
            "formula" => FORMULA,
            _ => NONE,
        }
    }
}

/// `com.sun.star.sheet.ValidationType` — data validation type constants.
pub mod validation_type {
    pub const ANY: i32 = 0;
    pub const LIST: i32 = 6;
    pub const WHOLE: i32 = 1;
    pub const DECIMAL: i32 = 2;
    pub const DATE: i32 = 3;
    pub const TIME: i32 = 4;
    pub const TEXT_LEN: i32 = 5;
    pub const CUSTOM: i32 = 7;

    /// Look up validation type by name (matching Python VALIDATION_TYPES map).
    pub fn from_name(name: &str) -> i32 {
        match name {
            "any" => ANY,
            "list" => LIST,
            "whole" => WHOLE,
            "decimal" => DECIMAL,
            "date" => DATE,
            "time" => TIME,
            "text_length" => TEXT_LEN,
            "custom" => CUSTOM,
            _ => ANY,
        }
    }
}

/// `com.sun.star.sheet.ValidationAlertStyle` — error alert style constants.
pub mod validation_alert_style {
    pub const STOP: i32 = 0;
    pub const WARNING: i32 = 1;
    pub const INFO: i32 = 2;
    pub const MACRO: i32 = 3;

    /// Look up alert style by name.
    pub fn from_name(name: &str) -> i32 {
        match name {
            "stop" => STOP,
            "warning" => WARNING,
            "info" => INFO,
            "macro" => MACRO,
            _ => STOP,
        }
    }
}

/// `com.sun.star.awt.Gradient` — describes a gradient fill.
///
/// Wire format: Struct(Enum style, Long startColor, Long endColor, Short angle,
///                     Short border, Short xOffset, Short yOffset,
///                     Short startIntensity, Short endIntensity, Short stepCount)
#[derive(Debug, Clone)]
pub struct GradientSpec {
    /// Gradient style: LINEAR=0, AXIAL=1, RADIAL=2, ELLIPTICAL=3, SQUARE=4, RECT=5
    pub style: i32,
    /// Start color as 0xRRGGBB integer
    pub start_color: i32,
    /// End color as 0xRRGGBB integer
    pub end_color: i32,
    /// Angle in 1/10 degree (0-3600), e.g. 900 = 90°
    pub angle: i16,
    /// Border percentage (0-100)
    pub border: i16,
    /// X offset for radial/elliptical (0-100)
    pub x_offset: i16,
    /// Y offset for radial/elliptical (0-100)
    pub y_offset: i16,
    /// Start intensity (0-100, typically 100)
    pub start_intensity: i16,
    /// End intensity (0-100, typically 100)
    pub end_intensity: i16,
    /// Step count (0 = automatic)
    pub step_count: i16,
}

impl GradientSpec {
    /// Create a simple linear gradient from start_color to end_color at given angle.
    pub fn linear(start_color: i32, end_color: i32, angle_degrees: f64) -> Self {
        Self {
            style: drawing_fill_style::GRADIENT_LINEAR,
            start_color,
            end_color,
            angle: (angle_degrees * 10.0) as i16,
            border: 0,
            x_offset: 0,
            y_offset: 0,
            start_intensity: 100,
            end_intensity: 100,
            step_count: 0,
        }
    }

    /// Convert to the URP wire representation.
    pub fn to_uno(&self) -> UnoValue {
        UnoValue::Struct(vec![
            UnoValue::Enum(self.style),            // Style
            UnoValue::Long(self.start_color),      // StartColor
            UnoValue::Long(self.end_color),        // EndColor
            UnoValue::Short(self.angle),           // Angle
            UnoValue::Short(self.border),          // Border
            UnoValue::Short(self.x_offset),        // XOffset
            UnoValue::Short(self.y_offset),        // YOffset
            UnoValue::Short(self.start_intensity), // StartIntensity
            UnoValue::Short(self.end_intensity),   // EndIntensity
            UnoValue::Short(self.step_count),      // StepCount
        ])
    }
}

/// `com.sun.star.drawing.FillStyle` — enum for fill type on cell backgrounds.
pub mod drawing_fill_style {
    pub const NONE: i32 = 0;
    pub const SOLID: i32 = 1;
    pub const GRADIENT: i32 = 2;
    // BITMAP = 3, HATCH = 4

    /// Gradient sub-styles (for awt.Gradient.Style enum)
    pub const GRADIENT_LINEAR: i32 = 0;
    pub const GRADIENT_AXIAL: i32 = 1;
    pub const GRADIENT_RADIAL: i32 = 2;
}

/// Fully-qualified UNO type names for gradient-related structs.
pub mod gradient_type_names {
    pub const AWT_GRADIENT: &str = "com.sun.star.awt.Gradient";
}

// ============================================================================
// UNO type name constants (struct/service names used in URP calls)
// ============================================================================

/// Fully-qualified UNO type names for structs used in this crate.
pub mod struct_type_names {
    pub const CELL_ADDRESS: &str = "com.sun.star.table.CellAddress";
    pub const BORDER_LINE2: &str = "com.sun.star.table.BorderLine2";
    pub const LOCALE: &str = "com.sun.star.lang.Locale";
    pub const CELL_STYLE: &str = "com.sun.star.style.CellStyle";
}

// ============================================================================
// StyleSpec — Rust port of Python's StyleSpec dataclass
// ============================================================================

/// Specification for cell/range styling, mirroring the Python `StyleSpec` dataclass.
///
/// All fields are optional. Only non-default values are applied to the cell.
#[derive(Debug, Clone, Default)]
pub struct StyleSpec {
    // Font
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<String>,
    pub strikethrough: bool,
    pub font_color: Option<i32>,
    pub font_size: Option<f32>,
    pub font_name: Option<String>,
    /// "superscript" or "subscript" — sets CharEscapement/CharEscapementHeight
    pub font_vertical_align: Option<String>,

    // Fill
    pub fill_color: Option<i32>,
    /// Gradient fill specification (overrides fill_color when set)
    pub fill_gradient: Option<GradientSpec>,

    // Alignment
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub rotation: i32,
    pub indent: i32,

    // Border (all sides)
    pub border_style: Option<String>,
    pub border_color: Option<i32>,

    // Individual borders: (style, color)
    pub left_border: Option<(String, i32)>,
    pub right_border: Option<(String, i32)>,
    pub top_border: Option<(String, i32)>,
    pub bottom_border: Option<(String, i32)>,

    // Number format
    pub number_format: Option<String>,
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_address_to_uno() {
        let addr = CellAddress::new(0, 3, 5);
        let value = addr.to_uno();
        match value {
            UnoValue::Struct(fields) => {
                assert_eq!(fields.len(), 3);
                assert_eq!(fields[0], UnoValue::Short(0));
                assert_eq!(fields[1], UnoValue::Long(3));
                assert_eq!(fields[2], UnoValue::Long(5));
            }
            _ => panic!("Expected Struct"),
        }
    }

    #[test]
    fn test_border_line2_to_uno() {
        let border = BorderLine2::new(0xFF0000, 1, 50);
        let value = border.to_uno();
        match value {
            UnoValue::Struct(fields) => {
                assert_eq!(fields.len(), 6);
                assert_eq!(fields[0], UnoValue::Long(0xFF0000_u32 as i32)); // Color
                assert_eq!(fields[1], UnoValue::Short(0)); // InnerLineWidth
                assert_eq!(fields[2], UnoValue::Short(0)); // OuterLineWidth
                assert_eq!(fields[3], UnoValue::Short(0)); // LineDistance
                assert_eq!(fields[4], UnoValue::Short(1)); // LineStyle
                assert_eq!(fields[5], UnoValue::Long(50)); // LineWidth
            }
            _ => panic!("Expected Struct"),
        }
    }

    #[test]
    fn test_border_line2_none() {
        let border = BorderLine2::none();
        let value = border.to_uno();
        match value {
            UnoValue::Struct(fields) => {
                assert_eq!(fields[0], UnoValue::Long(0));
                assert_eq!(fields[4], UnoValue::Short(0));
                assert_eq!(fields[5], UnoValue::Long(0));
            }
            _ => panic!("Expected Struct"),
        }
    }

    #[test]
    fn test_locale_empty() {
        let locale = Locale::empty();
        let value = locale.to_uno();
        match value {
            UnoValue::Struct(fields) => {
                assert_eq!(fields.len(), 3);
                assert_eq!(fields[0], UnoValue::String(String::new()));
                assert_eq!(fields[1], UnoValue::String(String::new()));
                assert_eq!(fields[2], UnoValue::String(String::new()));
            }
            _ => panic!("Expected Struct"),
        }
    }

    #[test]
    fn test_locale_with_values() {
        let locale = Locale::new("en", "US");
        let value = locale.to_uno();
        match value {
            UnoValue::Struct(fields) => {
                assert_eq!(fields[0], UnoValue::String("en".to_string()));
                assert_eq!(fields[1], UnoValue::String("US".to_string()));
                assert_eq!(fields[2], UnoValue::String(String::new()));
            }
            _ => panic!("Expected Struct"),
        }
    }

    #[test]
    fn test_hori_justify_from_name() {
        assert_eq!(hori_justify::from_name("left"), hori_justify::LEFT);
        assert_eq!(hori_justify::from_name("center"), hori_justify::CENTER);
        assert_eq!(hori_justify::from_name("right"), hori_justify::RIGHT);
        assert_eq!(hori_justify::from_name("general"), hori_justify::STANDARD);
        assert_eq!(hori_justify::from_name("unknown"), hori_justify::STANDARD);
    }

    #[test]
    fn test_vert_justify_from_name() {
        assert_eq!(vert_justify::from_name("top"), vert_justify::TOP);
        assert_eq!(vert_justify::from_name("center"), vert_justify::CENTER);
        assert_eq!(vert_justify::from_name("bottom"), vert_justify::BOTTOM);
        assert_eq!(vert_justify::from_name("unknown"), vert_justify::STANDARD);
    }

    #[test]
    fn test_border_style_from_name() {
        let (style, width) = border_line_style::from_name("thin");
        assert_eq!(style, border_line_style::SOLID);
        assert_eq!(width, 50);

        let (style, width) = border_line_style::from_name("medium");
        assert_eq!(style, border_line_style::SOLID);
        assert_eq!(width, 100);

        let (style, width) = border_line_style::from_name("thick");
        assert_eq!(style, border_line_style::SOLID);
        assert_eq!(width, 150);

        let (style, _width) = border_line_style::from_name("dashed");
        assert_eq!(style, border_line_style::DASHED);

        let (style, _width) = border_line_style::from_name("dotted");
        assert_eq!(style, border_line_style::DOTTED);

        let (style, _width) = border_line_style::from_name("double");
        assert_eq!(style, border_line_style::DOUBLE);

        let (style, width) = border_line_style::from_name("none");
        assert_eq!(style, 0);
        assert_eq!(width, 0);
    }

    #[test]
    fn test_underline_from_name() {
        assert_eq!(font_underline::from_name("single"), font_underline::SINGLE);
        assert_eq!(font_underline::from_name("double"), font_underline::DOUBLE);
        assert_eq!(
            font_underline::from_name("singleAccounting"),
            font_underline::SINGLE_ACCOUNTING
        );
        assert_eq!(
            font_underline::from_name("doubleAccounting"),
            font_underline::DOUBLE_ACCOUNTING
        );
    }

    #[test]
    fn test_condition_operator_from_name() {
        assert_eq!(
            condition_operator::from_name("greater_than"),
            condition_operator::GREATER
        );
        assert_eq!(
            condition_operator::from_name("less_than"),
            condition_operator::LESS
        );
        assert_eq!(
            condition_operator::from_name("between"),
            condition_operator::BETWEEN
        );
    }

    #[test]
    fn test_validation_type_from_name() {
        assert_eq!(validation_type::from_name("list"), validation_type::LIST);
        assert_eq!(validation_type::from_name("whole"), validation_type::WHOLE);
        assert_eq!(
            validation_type::from_name("custom"),
            validation_type::CUSTOM
        );
    }

    #[test]
    fn test_validation_alert_style_from_name() {
        assert_eq!(
            validation_alert_style::from_name("stop"),
            validation_alert_style::STOP
        );
        assert_eq!(
            validation_alert_style::from_name("warning"),
            validation_alert_style::WARNING
        );
        assert_eq!(
            validation_alert_style::from_name("info"),
            validation_alert_style::INFO
        );
    }

    #[test]
    fn test_style_spec_default() {
        let spec = StyleSpec::default();
        assert!(!spec.bold);
        assert!(!spec.italic);
        assert!(spec.underline.is_none());
        assert!(!spec.strikethrough);
        assert!(spec.font_color.is_none());
        assert!(spec.fill_color.is_none());
        assert!(spec.horizontal.is_none());
        assert!(spec.vertical.is_none());
        assert!(!spec.wrap_text);
        assert_eq!(spec.rotation, 0);
        assert_eq!(spec.indent, 0);
        assert!(spec.number_format.is_none());
    }
}
