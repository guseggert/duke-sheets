use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use super::{XlsxResult, XmlWriter};

pub(super) fn write_data_validations(
    w: &mut XmlWriter,
    sheet: &duke_sheets_core::Worksheet,
) -> XlsxResult<()> {
    use duke_sheets_core::validation::ValidationType;

    let validations = sheet.data_validations();
    if validations.is_empty() {
        return Ok(());
    }

    let count = validations.len().to_string();
    let mut dv_tag = BytesStart::new("dataValidations");
    dv_tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(dv_tag))?;

    for validation in validations {
        if validation.ranges.is_empty() {
            continue;
        }

        let sqref: String = validation
            .ranges
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");

        let mut tag = BytesStart::new("dataValidation");

        match &validation.validation_type {
            ValidationType::None => {}
            _ => {
                tag.push_attribute(("type", validation.validation_type.xlsx_type()));
            }
        }

        match &validation.validation_type {
            ValidationType::Whole { operator, .. }
            | ValidationType::Decimal { operator, .. }
            | ValidationType::Date { operator, .. }
            | ValidationType::Time { operator, .. }
            | ValidationType::TextLength { operator, .. } => {
                tag.push_attribute(("operator", operator.xlsx_operator()));
            }
            _ => {}
        }

        if validation.allow_blank {
            tag.push_attribute(("allowBlank", "1"));
        }
        if !validation.show_dropdown {
            tag.push_attribute(("showDropDown", "1"));
        }
        if validation.show_input_message {
            tag.push_attribute(("showInputMessage", "1"));
        }
        if validation.show_error_alert {
            tag.push_attribute(("showErrorMessage", "1"));
        }

        match validation.error_style {
            duke_sheets_core::ValidationErrorStyle::Stop => {}
            duke_sheets_core::ValidationErrorStyle::Warning => {
                tag.push_attribute(("errorStyle", "warning"));
            }
            duke_sheets_core::ValidationErrorStyle::Information => {
                tag.push_attribute(("errorStyle", "information"));
            }
        }

        if let Some(ref t) = validation.error_title {
            tag.push_attribute(("errorTitle", t.as_str()));
        }
        if let Some(ref m) = validation.error_message {
            tag.push_attribute(("error", m.as_str()));
        }
        if let Some(ref t) = validation.input_title {
            tag.push_attribute(("promptTitle", t.as_str()));
        }
        if let Some(ref m) = validation.input_message {
            tag.push_attribute(("prompt", m.as_str()));
        }

        tag.push_attribute(("sqref", sqref.as_str()));
        w.write_event(Event::Start(tag))?;

        match &validation.validation_type {
            ValidationType::List { source } => {
                let formula = if let Some(stripped) = source.strip_prefix('=') {
                    stripped.to_string()
                } else if source.contains('!')
                    || source
                        .chars()
                        .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':')
                {
                    source.clone()
                } else {
                    format!("\"{}\"", source)
                };
                w.create_element("formula1")
                    .write_text_content(BytesText::new(&formula))?;
            }
            ValidationType::Whole { value1, value2, .. }
            | ValidationType::Decimal { value1, value2, .. }
            | ValidationType::Date { value1, value2, .. }
            | ValidationType::Time { value1, value2, .. }
            | ValidationType::TextLength { value1, value2, .. } => {
                w.create_element("formula1")
                    .write_text_content(BytesText::new(value1))?;
                if let Some(v2) = value2 {
                    w.create_element("formula2")
                        .write_text_content(BytesText::new(v2))?;
                }
            }
            ValidationType::Custom { formula } => {
                let f = if let Some(stripped) = formula.strip_prefix('=') {
                    stripped
                } else {
                    formula
                };
                w.create_element("formula1")
                    .write_text_content(BytesText::new(f))?;
            }
            ValidationType::None => {}
        }

        w.write_event(Event::End(BytesEnd::new("dataValidation")))?;
    }

    w.write_event(Event::End(BytesEnd::new("dataValidations")))?;
    Ok(())
}
