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
    fn builtin_format_string(id: u32) -> &'static str {
        match id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
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
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            49 => "@",
            _ => "General",
        }
    }

    /// Check if this is a date/time format
    pub fn is_date_format(&self) -> bool {
        match self {
            NumberFormat::BuiltIn(id) => matches!(id, 14..=22),
            NumberFormat::Custom(s) => {
                // Simple heuristic: contains date/time placeholders but not literal text
                let lower = s.to_lowercase();
                (lower.contains('y')
                    || lower.contains('m')
                    || lower.contains('d')
                    || lower.contains('h')
                    || lower.contains('s'))
                    && !lower.contains('"')
            }
            NumberFormat::General => false,
        }
    }
}
