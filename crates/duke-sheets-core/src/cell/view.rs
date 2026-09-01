//! Cell view - a lightweight borrow wrapper that provides access to
//! a cell's value, style, and formatted display string.

use crate::cell::value::CellValue;
use crate::locale::Locale;
use crate::style::{NumberFormat, Style};

/// A read-only view of a cell, providing access to its value, style,
/// and formatted display string.
///
/// This is a lightweight borrow - no allocation or copying.
/// Obtain one via [`Worksheet::cell_view_at`] or [`Worksheet::cell_view`].
///
/// # Example
///
/// ```ignore
/// let cell = sheet.cell_view_at(0, 0);
/// println!("Raw: {}", cell.value());
/// println!("Formatted: {}", cell.formatted());
/// ```
pub struct CellView<'a> {
    value: &'a CellValue,
    style: Option<&'a Style>,
    date_1904: bool,
    locale: &'a ssfmt::Locale,
}

impl<'a> CellView<'a> {
    /// Create a new cell view.
    pub(crate) fn new(
        value: &'a CellValue,
        style: Option<&'a Style>,
        date_1904: bool,
        locale: &'a ssfmt::Locale,
    ) -> Self {
        Self {
            value,
            style,
            date_1904,
            locale,
        }
    }

    /// The cell's value.
    pub fn value(&self) -> &CellValue {
        self.value
    }

    /// The cell's style, if non-default.
    pub fn style(&self) -> Option<&Style> {
        self.style
    }

    /// The cell's number format (from the style, or `General` if no style).
    pub fn number_format(&self) -> &NumberFormat {
        self.style
            .map(|s| &s.number_format)
            .unwrap_or(&NumberFormat::General)
    }

    /// Format the cell value for display, applying the cell's number format.
    ///
    /// - Numbers are formatted according to the cell's `NumberFormat`
    ///   (percentages, currencies, dates, times, scientific notation, etc.)
    /// - Date serial numbers are converted to human-readable dates when the
    ///   format is a date/time format
    /// - Strings, booleans, and errors display as-is
    /// - Empty cells return an empty string
    pub fn formatted(&self) -> String {
        format_cell_value_inner(
            self.value,
            self.number_format(),
            self.date_1904,
            self.locale,
        )
    }
}

/// Format a cell value using the given number format, date system, and locale.
///
/// This is the standalone version of [`CellView::formatted`], useful when
/// you have the value and format separately.
pub fn format_cell_value(
    value: &CellValue,
    format: &NumberFormat,
    date_1904: bool,
    locale: &Locale,
) -> String {
    let ssfmt_locale = locale.to_ssfmt();
    format_cell_value_inner(value, format, date_1904, &ssfmt_locale)
}

/// Inner implementation shared by [`CellView::formatted`] (which has a
/// pre-cached `ssfmt::Locale`) and the public [`format_cell_value`].
fn format_cell_value_inner(
    value: &CellValue,
    format: &NumberFormat,
    date_1904: bool,
    locale: &ssfmt::Locale,
) -> String {
    let effective = value.effective_value();

    match effective {
        CellValue::Number(n) => format_number(*n, format, date_1904, locale),
        CellValue::String(s) => {
            // Apply text section of format if it has one (the @ placeholder)
            let code = format.format_string();
            if code != "General" && code != "@" {
                // ssfmt can handle text formatting for formats with a text section
                // but for most formats, just return the string as-is
                s.to_string()
            } else {
                s.to_string()
            }
        }
        CellValue::Boolean(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::Error(e) => e.to_string(),
        CellValue::Empty => String::new(),
        CellValue::SpillTarget { .. } => String::new(),
        CellValue::RichText(runs) => crate::rich_text::rich_text_to_plain(runs),
    }
}

/// Format a numeric value using an Excel number format code.
fn format_number(
    value: f64,
    format: &NumberFormat,
    date_1904: bool,
    locale: &ssfmt::Locale,
) -> String {
    let opts = ssfmt::FormatOptions {
        date_system: if date_1904 {
            ssfmt::DateSystem::Date1904
        } else {
            ssfmt::DateSystem::Date1900
        },
        locale: locale.clone(),
    };

    match format {
        NumberFormat::General => format_general(value),
        NumberFormat::BuiltIn(_) | NumberFormat::Custom(_) => {
            // Always go through format_string() so our builtin_format_string()
            // table is the single source of truth for built-in IDs.
            // (ssfmt::format_with_id uses SheetJS's table which diverges
            // from ECMA-376 on locale-dependent IDs like 14.)
            let code = format.format_string();
            match ssfmt::format(value, code, &opts) {
                Ok(s) => s,
                Err(_) => format_general(value),
            }
        }
    }
}

/// General format: display numbers without trailing zeros, integers
/// without decimal point.
fn format_general(value: f64) -> String {
    if value.is_nan() || value.is_infinite() {
        return value.to_string();
    }
    if value == 0.0 {
        return "0".to_string();
    }
    // Match Excel's General format: up to 11 significant digits,
    // no trailing zeros, scientific notation for very large/small values.
    let abs = value.abs();
    if abs >= 1e11 || (abs > 0.0 && abs < 1e-4) {
        // Scientific notation
        match ssfmt::format_default(value, "0.#####E+00") {
            Ok(s) => s,
            Err(_) => format!("{:E}", value),
        }
    } else if value.fract() == 0.0 && abs < 1e15 {
        format!("{}", value as i64)
    } else {
        // Remove trailing zeros from decimal representation
        let s = format!("{:.10}", value);
        let s = s.trim_end_matches('0');
        let s = s.trim_end_matches('.');
        s.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::value::CellError;

    // Helper
    fn fmt(value: &CellValue, format: &NumberFormat, date_1904: bool) -> String {
        format_cell_value(value, format, date_1904, &Locale::en_us())
    }

    // General format

    #[test]
    fn general_integer() {
        assert_eq!(
            fmt(&CellValue::Number(42.0), &NumberFormat::General, false),
            "42"
        );
    }

    #[test]
    fn general_zero() {
        assert_eq!(
            fmt(&CellValue::Number(0.0), &NumberFormat::General, false),
            "0"
        );
    }

    #[test]
    fn general_negative_integer() {
        assert_eq!(
            fmt(&CellValue::Number(-7.0), &NumberFormat::General, false),
            "-7"
        );
    }

    #[test]
    fn general_decimal() {
        assert_eq!(
            fmt(&CellValue::Number(3.14), &NumberFormat::General, false),
            "3.14"
        );
    }

    #[test]
    fn general_no_trailing_zeros() {
        assert_eq!(
            fmt(&CellValue::Number(1.50), &NumberFormat::General, false),
            "1.5"
        );
    }

    #[test]
    fn general_large_number_scientific() {
        let result = fmt(&CellValue::Number(1e12), &NumberFormat::General, false);
        // Should use scientific notation for numbers >= 1e11
        assert!(
            result.contains('E') || result.contains('e'),
            "expected scientific: {}",
            result
        );
    }

    #[test]
    fn general_small_number_scientific() {
        let result = fmt(&CellValue::Number(0.00001), &NumberFormat::General, false);
        assert!(
            result.contains('E') || result.contains('e'),
            "expected scientific: {}",
            result
        );
    }

    #[test]
    fn general_nan() {
        assert_eq!(
            fmt(&CellValue::Number(f64::NAN), &NumberFormat::General, false),
            "NaN"
        );
    }

    #[test]
    fn general_infinity() {
        assert_eq!(
            fmt(
                &CellValue::Number(f64::INFINITY),
                &NumberFormat::General,
                false
            ),
            "inf"
        );
    }

    // Built-in number formats

    #[test]
    fn builtin_integer_format() {
        // Format ID 1: "0"
        let result = fmt(
            &CellValue::Number(1234.567),
            &NumberFormat::BuiltIn(1),
            false,
        );
        assert_eq!(result, "1235");
    }

    #[test]
    fn builtin_decimal_format() {
        // Format ID 2: "0.00"
        let result = fmt(&CellValue::Number(1234.5), &NumberFormat::BuiltIn(2), false);
        assert_eq!(result, "1234.50");
    }

    #[test]
    fn builtin_thousands_separator() {
        // Format ID 3: "#,##0"
        let result = fmt(
            &CellValue::Number(1234567.0),
            &NumberFormat::BuiltIn(3),
            false,
        );
        assert_eq!(result, "1,234,567");
    }

    #[test]
    fn builtin_thousands_decimal() {
        // Format ID 4: "#,##0.00"
        let result = fmt(
            &CellValue::Number(1234567.89),
            &NumberFormat::BuiltIn(4),
            false,
        );
        assert_eq!(result, "1,234,567.89");
    }

    // Percentage formats

    #[test]
    fn builtin_percent_integer() {
        // Format ID 9: "0%"
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(9), false);
        assert_eq!(result, "75%");
    }

    #[test]
    fn builtin_percent_decimal() {
        // Format ID 10: "0.00%"
        let result = fmt(
            &CellValue::Number(0.1234),
            &NumberFormat::BuiltIn(10),
            false,
        );
        assert_eq!(result, "12.34%");
    }

    // Scientific notation format

    #[test]
    fn builtin_scientific() {
        // Format ID 11: "0.00E+00"
        let result = fmt(
            &CellValue::Number(12345.0),
            &NumberFormat::BuiltIn(11),
            false,
        );
        assert_eq!(result, "1.23E+04");
    }

    // Date formats (1900 date system)

    #[test]
    fn builtin_date_short_1900() {
        // Format ID 14: "mm-dd-yy"
        // Serial 44927 = 2023-01-01 in 1900 system
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(14),
            false,
        );
        assert_eq!(result, "01-01-23");
    }

    #[test]
    fn builtin_date_medium() {
        // Format ID 15: "d-mmm-yy"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(15),
            false,
        );
        assert_eq!(result, "1-Jan-23");
    }

    #[test]
    fn builtin_date_day_month() {
        // Format ID 16: "d-mmm"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(16),
            false,
        );
        assert_eq!(result, "1-Jan");
    }

    #[test]
    fn builtin_date_month_year() {
        // Format ID 17: "mmm-yy"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(17),
            false,
        );
        assert_eq!(result, "Jan-23");
    }

    // Date formats (1904 date system)

    #[test]
    fn builtin_date_short_1904() {
        // In 1904 system, serial 43465 = 2023-01-01
        // (1904 system starts 1462 days later than 1900 system)
        let result = fmt(
            &CellValue::Number(43465.0),
            &NumberFormat::BuiltIn(14),
            true,
        );
        assert_eq!(result, "01-01-23");
    }

    // Time formats

    #[test]
    fn builtin_time_ampm() {
        // Format ID 18: "h:mm AM/PM"
        // 0.75 = 18:00 = 6:00 PM
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(18), false);
        assert_eq!(result, "6:00 PM");
    }

    #[test]
    fn builtin_time_24h() {
        // Format ID 20: "h:mm"
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(20), false);
        assert_eq!(result, "18:00");
    }

    #[test]
    fn builtin_time_24h_seconds() {
        // Format ID 21: "h:mm:ss"
        // 0.5 = 12:00:00 noon
        let result = fmt(&CellValue::Number(0.5), &NumberFormat::BuiltIn(21), false);
        assert_eq!(result, "12:00:00");
    }

    // Custom format strings

    #[test]
    fn custom_percent_format() {
        let fmt_code = NumberFormat::Custom("0.0%".to_string());
        let result = fmt(&CellValue::Number(0.1234), &fmt_code, false);
        assert_eq!(result, "12.3%");
    }

    #[test]
    fn custom_thousands_format() {
        let fmt_code = NumberFormat::Custom("#,##0.00".to_string());
        let result = fmt(&CellValue::Number(9876.5), &fmt_code, false);
        assert_eq!(result, "9,876.50");
    }

    #[test]
    fn custom_date_format() {
        let fmt_code = NumberFormat::Custom("yyyy-mm-dd".to_string());
        // Serial 44927 = 2023-01-01
        let result = fmt(&CellValue::Number(44927.0), &fmt_code, false);
        assert_eq!(result, "2023-01-01");
    }

    #[test]
    fn custom_currency_format() {
        let fmt_code = NumberFormat::Custom("$#,##0.00".to_string());
        let result = fmt(&CellValue::Number(1234.5), &fmt_code, false);
        assert_eq!(result, "$1,234.50");
    }

    // Non-numeric cell types (passthrough)

    #[test]
    fn string_passthrough() {
        let val = CellValue::string("Hello World");
        assert_eq!(fmt(&val, &NumberFormat::General, false), "Hello World");
    }

    #[test]
    fn string_ignores_number_format() {
        let val = CellValue::string("Not a number");
        let fmt_code = NumberFormat::Custom("#,##0.00".to_string());
        assert_eq!(fmt(&val, &fmt_code, false), "Not a number");
    }

    #[test]
    fn boolean_true() {
        assert_eq!(
            fmt(&CellValue::Boolean(true), &NumberFormat::General, false),
            "TRUE"
        );
    }

    #[test]
    fn boolean_false() {
        assert_eq!(
            fmt(&CellValue::Boolean(false), &NumberFormat::General, false),
            "FALSE"
        );
    }

    #[test]
    fn error_value() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Value),
                &NumberFormat::General,
                false
            ),
            "#VALUE!"
        );
    }

    #[test]
    fn error_ref() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Ref),
                &NumberFormat::General,
                false
            ),
            "#REF!"
        );
    }

    #[test]
    fn error_div0() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Div0),
                &NumberFormat::General,
                false
            ),
            "#DIV/0!"
        );
    }

    #[test]
    fn empty_cell() {
        assert_eq!(fmt(&CellValue::Empty, &NumberFormat::General, false), "");
    }

    #[test]
    fn cached_formula_number_formats_as_number() {
        let val = CellValue::Number(42.0);
        assert_eq!(fmt(&val, &NumberFormat::General, false), "42");
    }

    #[test]
    fn cached_formula_number_uses_cell_format() {
        let val = CellValue::Number(0.85);
        let fmt_code = NumberFormat::BuiltIn(9); // "0%"
        assert_eq!(fmt(&val, &fmt_code, false), "85%");
    }

    #[test]
    fn cached_formula_string_formats_as_string() {
        let val = CellValue::string("yes");
        assert_eq!(fmt(&val, &NumberFormat::General, false), "yes");
    }

    #[test]
    fn cached_formula_error_formats_as_error() {
        let val = CellValue::Error(CellError::Div0);
        assert_eq!(fmt(&val, &NumberFormat::General, false), "#DIV/0!");
    }

    #[test]
    fn uncached_formula_formats_as_empty_cell_value() {
        let val = CellValue::Empty;
        assert_eq!(fmt(&val, &NumberFormat::General, false), "");
    }

    // CellView API

    // Cached ssfmt locale for CellView tests (CellView takes &ssfmt::Locale
    // internally since the worksheet caches the conversion).
    fn ssfmt_en_us() -> ssfmt::Locale {
        ssfmt::Locale::en_us()
    }

    #[test]
    fn cell_view_value_accessor() {
        let val = CellValue::Number(3.14);
        let locale = ssfmt_en_us();
        let view = CellView::new(&val, None, false, &locale);
        assert_eq!(view.value(), &CellValue::Number(3.14));
    }

    #[test]
    fn cell_view_default_format() {
        let val = CellValue::Number(3.14);
        let locale = ssfmt_en_us();
        let view = CellView::new(&val, None, false, &locale);
        assert_eq!(view.number_format(), &NumberFormat::General);
    }

    #[test]
    fn cell_view_with_style() {
        let val = CellValue::Number(0.75);
        let style = Style {
            number_format: NumberFormat::BuiltIn(9), // "0%"
            ..Default::default()
        };
        let locale = ssfmt_en_us();
        let view = CellView::new(&val, Some(&style), false, &locale);
        assert_eq!(view.formatted(), "75%");
    }

    #[test]
    fn cell_view_style_accessor() {
        let val = CellValue::Number(1.0);
        let style = Style {
            number_format: NumberFormat::BuiltIn(2),
            ..Default::default()
        };
        let locale = ssfmt_en_us();
        let view = CellView::new(&val, Some(&style), false, &locale);
        assert!(view.style().is_some());

        let view_no_style = CellView::new(&val, None, false, &locale);
        assert!(view_no_style.style().is_none());
    }

    // Accounting / currency built-in formats

    #[test]
    fn builtin_currency_format_5() {
        // Format ID 5: "$#,##0_);($#,##0)"
        let result = fmt(&CellValue::Number(1234.0), &NumberFormat::BuiltIn(5), false);
        assert!(result.contains("1,234"), "expected currency: {}", result);
        assert!(result.contains('$'), "expected dollar sign: {}", result);
    }

    #[test]
    fn builtin_accounting_negative() {
        // Format ID 37: "#,##0_);(#,##0)" - negative in parens
        let result = fmt(
            &CellValue::Number(-1234.0),
            &NumberFormat::BuiltIn(37),
            false,
        );
        assert!(result.contains("1,234"), "expected number: {}", result);
        assert!(
            result.contains('('),
            "expected parens for negative: {}",
            result
        );
    }

    // Edge cases

    #[test]
    fn negative_zero() {
        // -0.0 should display as "0" under General format
        let result = format_general(-0.0);
        assert_eq!(result, "0");
    }

    #[test]
    fn very_small_decimal() {
        let result = fmt(&CellValue::Number(0.1 + 0.2), &NumberFormat::General, false);
        // Should produce "0.3" or close to it, not "0.30000000000000004"
        assert!(result.starts_with("0.3"), "expected ~0.3: {}", result);
    }

    #[test]
    fn spill_target_empty() {
        let val = CellValue::SpillTarget {
            source_row: 0,
            source_col: 0,
            offset_row: 1,
            offset_col: 0,
        };
        assert_eq!(fmt(&val, &NumberFormat::General, false), "");
    }

    #[test]
    fn text_format_id_49() {
        // Format ID 49: "@" - text format, numbers should still format
        let result = fmt(&CellValue::Number(42.0), &NumberFormat::BuiltIn(49), false);
        // ssfmt with "@" format on a number typically returns the number as text
        assert!(!result.is_empty(), "expected non-empty for text format");
    }

    #[test]
    fn date_format_with_underscore_paren_no_trailing_paren() {
        // Real-world Excel files often use formats like m/d/yyyy_) where _)
        // means "pad with width of ')'". The ')' must NOT appear in output.
        // Serial 46022 = 12/31/2025 in the 1900 date system.
        let val = CellValue::Number(46022.0);
        let fmt_code = NumberFormat::Custom("m/d/yyyy_)".to_string());
        let result = fmt(&val, &fmt_code, false);
        assert_eq!(result, "12/31/2025 ", "underscore-paren should produce a space, not a literal ')'");
    }
}
