//! Number format types

/// Number format for cell display
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub enum NumberFormat {
    /// General format (default)
    #[default]
    General,

    /// Built-in format by ID
    BuiltIn(u32),

    /// Custom format string
    Custom(String),
}

impl NumberFormat {
    /// General format
    pub const GENERAL: Self = NumberFormat::General;

    // Built-in format IDs
    /// 0 - General
    pub const ID_GENERAL: u32 = 0;
    /// 1 - 0
    pub const ID_NUMBER_INT: u32 = 1;
    /// 2 - 0.00
    pub const ID_NUMBER_DEC2: u32 = 2;
    /// 3 - #,##0
    pub const ID_NUMBER_SEP: u32 = 3;
    /// 4 - #,##0.00
    pub const ID_NUMBER_SEP_DEC2: u32 = 4;
    /// 9 - 0%
    pub const ID_PERCENT_INT: u32 = 9;
    /// 10 - 0.00%
    pub const ID_PERCENT_DEC2: u32 = 10;
    /// 11 - 0.00E+00
    pub const ID_SCIENTIFIC: u32 = 11;
    /// 12 - # ?/?
    pub const ID_FRACTION: u32 = 12;
    /// 13 - # ??/??
    pub const ID_FRACTION2: u32 = 13;
    /// 14 - mm-dd-yy
    pub const ID_DATE_SHORT: u32 = 14;
    /// 15 - d-mmm-yy
    pub const ID_DATE_MEDIUM: u32 = 15;
    /// 16 - d-mmm
    pub const ID_DATE_DAY_MONTH: u32 = 16;
    /// 17 - mmm-yy
    pub const ID_DATE_MONTH_YEAR: u32 = 17;
    /// 18 - h:mm AM/PM
    pub const ID_TIME_AMPM: u32 = 18;
    /// 19 - h:mm:ss AM/PM
    pub const ID_TIME_AMPM_SEC: u32 = 19;
    /// 20 - h:mm
    pub const ID_TIME_24H: u32 = 20;
    /// 21 - h:mm:ss
    pub const ID_TIME_24H_SEC: u32 = 21;
    /// 22 - m/d/yy h:mm
    pub const ID_DATETIME: u32 = 22;
    /// 37 - #,##0 ;(#,##0)
    pub const ID_ACCOUNTING_INT: u32 = 37;
    /// 38 - #,##0 ;[Red](#,##0)
    pub const ID_ACCOUNTING_INT_RED: u32 = 38;
    /// 39 - #,##0.00;(#,##0.00)
    pub const ID_ACCOUNTING_DEC2: u32 = 39;
    /// 40 - #,##0.00;[Red](#,##0.00)
    pub const ID_ACCOUNTING_DEC2_RED: u32 = 40;
    /// 49 - @
    pub const ID_TEXT: u32 = 49;

    /// Create a number format from a format string
    pub fn from_string<S: Into<String>>(format: S) -> Self {
        NumberFormat::Custom(format.into())
    }

    /// Create a built-in format by ID
    pub fn from_id(id: u32) -> Self {
        NumberFormat::BuiltIn(id)
    }

    /// Integer format (0)
    pub fn integer() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_INT)
    }

    /// Decimal format (0.00)
    pub fn decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_DEC2)
    }

    /// Number with thousands separator (#,##0)
    pub fn thousands() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_SEP)
    }

    /// Number with thousands separator and decimals (#,##0.00)
    pub fn thousands_decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_SEP_DEC2)
    }

    /// Percentage (0%)
    pub fn percent() -> Self {
        NumberFormat::BuiltIn(Self::ID_PERCENT_INT)
    }

    /// Percentage with decimals (0.00%)
    pub fn percent_decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_PERCENT_DEC2)
    }

    /// Scientific notation (0.00E+00)
    pub fn scientific() -> Self {
        NumberFormat::BuiltIn(Self::ID_SCIENTIFIC)
    }

    /// Short date (mm-dd-yy)
    pub fn date_short() -> Self {
        NumberFormat::BuiltIn(Self::ID_DATE_SHORT)
    }

    /// Time with AM/PM (h:mm AM/PM)
    pub fn time_ampm() -> Self {
        NumberFormat::BuiltIn(Self::ID_TIME_AMPM)
    }

    /// Date and time (m/d/yy h:mm)
    pub fn datetime() -> Self {
        NumberFormat::BuiltIn(Self::ID_DATETIME)
    }

    /// Text format (@)
    pub fn text() -> Self {
        NumberFormat::BuiltIn(Self::ID_TEXT)
    }

    /// Get the format string
    pub fn format_string(&self) -> &str {
        match self {
            NumberFormat::General => "General",
            NumberFormat::BuiltIn(id) => Self::builtin_format_string(*id),
            NumberFormat::Custom(s) => s,
        }
    }

    /// Get built-in format string by ID
    ///
    /// Reference: ECMA-376 Section 18.8.30 (numFmt), plus the implicit
    /// built-in formats defined by the spec and Microsoft documentation.
    fn builtin_format_string(id: u32) -> &'static str {
        match id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            5 => "$#,##0_);($#,##0)",
            6 => "$#,##0_);[Red]($#,##0)",
            7 => "$#,##0.00_);($#,##0.00)",
            8 => "$#,##0.00_);[Red]($#,##0.00)",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "mm-dd-yy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yy h:mm",
            37 => "#,##0_);(#,##0)",
            38 => "#,##0_);[Red](#,##0)",
            39 => "#,##0.00_);(#,##0.00)",
            40 => "#,##0.00_);[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mm:ss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => "General",
        }
    }

    /// Check if this is a date/time format
    ///
    /// For built-in formats, checks the known ID ranges.
    /// For custom formats, uses token-aware parsing that skips quoted
    /// literal text and escaped characters to avoid false positives.
    pub fn is_date_format(&self) -> bool {
        match self {
            NumberFormat::BuiltIn(id) => matches!(id, 14..=22 | 45..=47),
            NumberFormat::Custom(s) => Self::custom_is_date_format(s),
            NumberFormat::General => false,
        }
    }

    /// Token-aware check for date/time format codes in a custom format string.
    /// Skips quoted literals ("..."), escaped characters (\x), and bracketed
    /// sections ([...]) to avoid false positives on strings like `0 "items"`.
    fn custom_is_date_format(s: &str) -> bool {
        let bytes = s.as_bytes();
        let mut i = 0;
        while i < bytes.len() {
            match bytes[i] {
                // Skip quoted literal text
                b'"' => {
                    i += 1;
                    while i < bytes.len() && bytes[i] != b'"' {
                        i += 1;
                    }
                    i += 1; // skip closing quote
                }
                // Skip escaped character
                b'\\' => {
                    i += 2;
                }
                // Skip bracketed sections (colors, conditions, locale, elapsed time)
                // but check for [h], [m], [s] which ARE date/time tokens
                b'[' => {
                    let start = i + 1;
                    i += 1;
                    while i < bytes.len() && bytes[i] != b']' {
                        i += 1;
                    }
                    let bracket_len = i - start;
                    if bracket_len == 1 {
                        let ch = bytes[start];
                        if ch == b'h' || ch == b'm' || ch == b's' {
                            return true;
                        }
                    }
                    i += 1; // skip closing bracket
                }
                // Skip underscore + next char (alignment spacer)
                b'_' => {
                    i += 2;
                }
                // Skip asterisk + next char (fill repeat)
                b'*' => {
                    i += 2;
                }
                // Date/time tokens
                b'y' | b'Y' | b'd' | b'D' => return true,
                // m/M is date (month) - could be minutes after h, but either way it's date/time
                b'm' | b'M' => return true,
                // h/H and s/S are time tokens
                b'h' | b'H' | b's' | b'S' => return true,
                // AM/PM indicator
                b'A' | b'a' => {
                    // Check for AM/PM or A/P
                    let remaining = &s[i..];
                    let lower = remaining.to_ascii_lowercase();
                    if lower.starts_with("am/pm") || lower.starts_with("a/p") {
                        return true;
                    }
                    i += 1;
                }
                _ => {
                    i += 1;
                }
            }
        }
        false
    }
}
