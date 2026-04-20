use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};

pub(super) fn parse_data_validation_attrs(e: &quick_xml::events::BytesStart) -> DataValidation {
    let mut validation = DataValidation::new();
    let mut dv_type: Option<String> = None;
    let mut operator: Option<String> = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"type" => dv_type = attr.unescape_value().ok().map(|s| s.to_string()),
            b"operator" => operator = attr.unescape_value().ok().map(|s| s.to_string()),
            b"allowBlank" => {
                validation.allow_blank = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"showDropDown" => {
                validation.show_dropdown = attr.unescape_value().ok().is_none_or(|s| s != "1");
            }
            b"showInputMessage" => {
                validation.show_input_message =
                    attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"showErrorMessage" => {
                validation.show_error_alert = attr.unescape_value().ok().is_some_and(|s| s == "1");
            }
            b"errorStyle" => {
                if let Ok(style) = attr.unescape_value() {
                    validation.error_style = match style.as_ref() {
                        "warning" => ValidationErrorStyle::Warning,
                        "information" => ValidationErrorStyle::Information,
                        _ => ValidationErrorStyle::Stop,
                    };
                }
            }
            b"errorTitle" => {
                validation.error_title = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"error" => {
                validation.error_message = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"promptTitle" => {
                validation.input_title = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"prompt" => {
                validation.input_message = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"sqref" => {
                if let Ok(sqref) = attr.unescape_value() {
                    validation.ranges = super::parse_sqref(&sqref);
                }
            }
            _ => {}
        }
    }

    let op = operator
        .as_deref()
        .and_then(ValidationOperator::from_xlsx)
        .unwrap_or(ValidationOperator::Between);

    validation.validation_type = match dv_type.as_deref() {
        Some("list") => ValidationType::List {
            source: String::new(),
        },
        Some("whole") => ValidationType::Whole {
            operator: op,
            value1: String::new(),
            value2: None,
        },
        Some("decimal") => ValidationType::Decimal {
            operator: op,
            value1: String::new(),
            value2: None,
        },
        Some("date") => ValidationType::Date {
            operator: op,
            value1: String::new(),
            value2: None,
        },
        Some("time") => ValidationType::Time {
            operator: op,
            value1: String::new(),
            value2: None,
        },
        Some("textLength") => ValidationType::TextLength {
            operator: op,
            value1: String::new(),
            value2: None,
        },
        Some("custom") => ValidationType::Custom {
            formula: String::new(),
        },
        _ => ValidationType::None,
    };

    validation
}

pub(super) fn apply_validation_formulas(
    validation: &mut DataValidation,
    formula1: Option<String>,
    formula2: Option<String>,
) {
    match &mut validation.validation_type {
        ValidationType::List { source } => {
            if let Some(f1) = formula1 {
                *source = f1.trim_matches('"').to_string();
            }
        }
        ValidationType::Whole { value1, value2, .. }
        | ValidationType::Decimal { value1, value2, .. }
        | ValidationType::Date { value1, value2, .. }
        | ValidationType::Time { value1, value2, .. }
        | ValidationType::TextLength { value1, value2, .. } => {
            if let Some(f1) = formula1 {
                *value1 = f1;
            }
            *value2 = formula2;
        }
        ValidationType::Custom { formula } => {
            if let Some(f1) = formula1 {
                *formula = f1;
            }
        }
        ValidationType::None => {}
    }
}
