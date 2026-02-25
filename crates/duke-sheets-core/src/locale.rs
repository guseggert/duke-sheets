//! Locale settings for cell display formatting.
//!
//! Controls decimal separators, thousands separators, currency symbols,
//! and month/day names when rendering cell values via [`CellView::formatted()`]
//! or [`format_cell_value()`].
//!
//! [`CellView::formatted()`]: crate::CellView::formatted
//! [`format_cell_value()`]: crate::format_cell_value

/// Locale settings for number and date formatting.
///
/// Set on a worksheet via [`Worksheet::set_locale`] to control how
/// built-in format IDs render.  Custom format strings with `[$-XXXX]`
/// locale prefixes override these settings per-cell.
///
/// # Example
///
/// ```
/// use duke_sheets_core::Locale;
///
/// // Use the built-in German locale
/// let locale = Locale::de_de();
/// assert_eq!(locale.decimal_separator, ',');
/// assert_eq!(locale.thousands_separator, '.');
/// ```
///
/// [`Worksheet::set_locale`]: crate::Worksheet::set_locale
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Locale {
    /// Decimal point character (e.g., '.' for en-US, ',' for de-DE).
    pub decimal_separator: char,
    /// Thousands grouping character (e.g., ',' for en-US, '.' for de-DE).
    pub thousands_separator: char,
    /// Currency symbol (e.g., "$", "€", "£").
    pub currency_symbol: String,
    /// AM string for 12-hour time (e.g., "AM").
    pub am_string: String,
    /// PM string for 12-hour time (e.g., "PM").
    pub pm_string: String,
    /// Abbreviated month names (Jan, Feb, ..., Dec).
    pub month_names_short: [String; 12],
    /// Full month names (January, February, ..., December).
    pub month_names_full: [String; 12],
    /// Abbreviated day names starting from Sunday (Sun, Mon, ..., Sat).
    pub day_names_short: [String; 7],
    /// Full day names starting from Sunday (Sunday, Monday, ..., Saturday).
    pub day_names_full: [String; 7],
}

impl Default for Locale {
    fn default() -> Self {
        Self::en_us()
    }
}

impl Locale {
    /// US English locale.
    pub fn en_us() -> Self {
        Self {
            decimal_separator: '.',
            thousands_separator: ',',
            currency_symbol: "$".into(),
            am_string: "AM".into(),
            pm_string: "PM".into(),
            month_names_short: arr_s([
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ]),
            month_names_full: arr_s([
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ]),
            day_names_short: arr_s_7(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]),
            day_names_full: arr_s_7([
                "Sunday",
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
            ]),
        }
    }

    /// German locale (de-DE).
    pub fn de_de() -> Self {
        Self {
            decimal_separator: ',',
            thousands_separator: '.',
            currency_symbol: "€".into(),
            am_string: "AM".into(),
            pm_string: "PM".into(),
            month_names_short: arr_s([
                "Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez",
            ]),
            month_names_full: arr_s([
                "Januar",
                "Februar",
                "März",
                "April",
                "Mai",
                "Juni",
                "Juli",
                "August",
                "September",
                "Oktober",
                "November",
                "Dezember",
            ]),
            day_names_short: arr_s_7(["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"]),
            day_names_full: arr_s_7([
                "Sonntag",
                "Montag",
                "Dienstag",
                "Mittwoch",
                "Donnerstag",
                "Freitag",
                "Samstag",
            ]),
        }
    }

    /// French locale (fr-FR).
    pub fn fr_fr() -> Self {
        Self {
            decimal_separator: ',',
            // French uses narrow no-break space (U+202F) as grouping separator
            // but regular space is more practical for plain text output.
            thousands_separator: ' ',
            currency_symbol: "€".into(),
            am_string: "AM".into(),
            pm_string: "PM".into(),
            month_names_short: arr_s([
                "janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.",
                "nov.", "déc.",
            ]),
            month_names_full: arr_s([
                "janvier",
                "février",
                "mars",
                "avril",
                "mai",
                "juin",
                "juillet",
                "août",
                "septembre",
                "octobre",
                "novembre",
                "décembre",
            ]),
            day_names_short: arr_s_7(["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."]),
            day_names_full: arr_s_7([
                "dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi",
            ]),
        }
    }

    /// British English locale (en-GB).
    pub fn en_gb() -> Self {
        Self {
            currency_symbol: "£".into(),
            ..Self::en_us()
        }
    }

    /// Japanese locale (ja-JP).
    pub fn ja_jp() -> Self {
        Self {
            decimal_separator: '.',
            thousands_separator: ',',
            currency_symbol: "¥".into(),
            am_string: "午前".into(),
            pm_string: "午後".into(),
            month_names_short: arr_s([
                "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月",
                "12月",
            ]),
            month_names_full: arr_s([
                "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月",
                "12月",
            ]),
            day_names_short: arr_s_7(["日", "月", "火", "水", "木", "金", "土"]),
            day_names_full: arr_s_7([
                "日曜日",
                "月曜日",
                "火曜日",
                "水曜日",
                "木曜日",
                "金曜日",
                "土曜日",
            ]),
        }
    }

    /// Convert to ssfmt's Locale type (internal use only).
    pub(crate) fn to_ssfmt(&self) -> ssfmt::Locale {
        ssfmt::Locale {
            decimal_separator: self.decimal_separator,
            thousands_separator: self.thousands_separator,
            // ssfmt uses &'static str — we leak the string to satisfy the
            // lifetime.  This is fine: locales are long-lived and few in
            // number (typically one per workbook).
            currency_symbol: leak_str(&self.currency_symbol),
            am_string: leak_str(&self.am_string),
            pm_string: leak_str(&self.pm_string),
            month_names_short: leak_arr_12(&self.month_names_short),
            month_names_full: leak_arr_12(&self.month_names_full),
            day_names_short: leak_arr_7(&self.day_names_short),
            day_names_full: leak_arr_7(&self.day_names_full),
        }
    }
}

// ---- Helpers ----

fn arr_s<const N: usize>(src: [&str; N]) -> [String; N] {
    src.map(String::from)
}

fn arr_s_7(src: [&str; 7]) -> [String; 7] {
    src.map(String::from)
}

fn leak_str(s: &str) -> &'static str {
    // Check common values to avoid unnecessary leaks.
    match s {
        "$" => "$",
        "€" => "€",
        "£" => "£",
        "¥" => "¥",
        "AM" => "AM",
        "PM" => "PM",
        _ => Box::leak(s.to_string().into_boxed_str()),
    }
}

fn leak_arr_12(src: &[String; 12]) -> [&'static str; 12] {
    let mut out: [&'static str; 12] = [""; 12];
    for (i, s) in src.iter().enumerate() {
        out[i] = leak_str(s);
    }
    out
}

fn leak_arr_7(src: &[String; 7]) -> [&'static str; 7] {
    let mut out: [&'static str; 7] = [""; 7];
    for (i, s) in src.iter().enumerate() {
        out[i] = leak_str(s);
    }
    out
}
