//! XLSX writer

use std::fs::File;
use std::io::{Seek, Write};
use std::path::Path;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::XlsxStyleTable;
use duke_sheets_core::{CellAddress, Workbook};

/// XLSX file writer
pub struct XlsxWriter;

impl XlsxWriter {
    /// Write a workbook to a file path
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsxResult<()> {
        let file = File::create(path)?;
        Self::write(workbook, file)
    }

    /// Write a workbook to a writer
    pub fn write<W: Write + Seek>(workbook: &Workbook, writer: W) -> XlsxResult<()> {
        let mut zip = zip::ZipWriter::new(writer);

        // Build a workbook-wide style table.
        let style_table = XlsxStyleTable::build(workbook);

        // Determine which sheets have comments
        let sheets_with_comments: Vec<usize> = workbook
            .worksheets()
            .enumerate()
            .filter(|(_, sheet)| sheet.comment_count() > 0)
            .map(|(i, _)| i)
            .collect();

        // Write [Content_Types].xml
        Self::write_content_types(&mut zip, workbook, &sheets_with_comments)?;

        // Write _rels/.rels
        Self::write_root_rels(&mut zip)?;

        // Write xl/workbook.xml
        Self::write_workbook_xml(&mut zip, workbook)?;

        // Write xl/_rels/workbook.xml.rels
        Self::write_workbook_rels(&mut zip, workbook)?;

        // Write xl/styles.xml
        Self::write_styles_xml(&mut zip, &style_table)?;

        // Write worksheets and their relationships
        for (i, sheet) in workbook.worksheets().enumerate() {
            Self::write_worksheet(&mut zip, workbook, i, &style_table)?;

            // Write worksheet relationships if sheet has comments
            if sheet.comment_count() > 0 {
                Self::write_worksheet_rels(&mut zip, i)?;
                Self::write_comments(&mut zip, workbook, i)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    fn write_content_types<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheets_with_comments: &[usize],
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("[Content_Types].xml", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
        );

        // Add an override for each worksheet
        for i in 0..workbook.sheet_count() {
            content.push_str(&format!(
                r#"
    <Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i + 1
            ));
        }

        // Add content type for comments files
        for &i in sheets_with_comments {
            content.push_str(&format!(
                r#"
    <Override PartName="/xl/comments{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"#,
                i + 1
            ));
        }

        content.push_str("\n</Types>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_root_rels<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("_rels/.rels", options)?;

        let content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/workbook.xml", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>"#,
        );

        for (i, sheet) in workbook.worksheets().enumerate() {
            content.push_str(&format!(
                r#"
        <sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                sheet.name(),
                i + 1,
                i + 1
            ));
        }

        content.push_str(
            r#"
    </sheets>
</workbook>"#,
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/_rels/workbook.xml.rels", options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        for i in 0..workbook.sheet_count() {
            content.push_str(&format!(
                r#"
    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i + 1,
                i + 1
            ));
        }

        // Styles relationship
        let styles_rid = workbook.sheet_count() + 1;
        content.push_str(&format!(
            r#"
    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
            styles_rid
        ));

        content.push_str(
            r#"
 </Relationships>"#,
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_styles_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/styles.xml", options)?;
        let xml = style_table.to_styles_xml();
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn write_worksheet<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        index: usize,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(&format!("xl/worksheets/sheet{}.xml", index + 1), options)?;

        let sheet = workbook
            .worksheet(index)
            .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>"#,
        );

        // Write cell data (sparse, row-major)
        let mut current_row: Option<u32> = None;
        for (row, col, cell) in sheet.iter_cells() {
            if current_row != Some(row) {
                // Close previous row
                if current_row.is_some() {
                    content.push_str("\n        </row>");
                }
                // Open new row
                content.push_str(&format!("\n        <row r=\"{}\">", row + 1));
                current_row = Some(row);
            }

            let addr = duke_sheets_core::CellAddress::new(row, col);
            let cell_ref = addr.to_a1_string();

            let xf_id = style_table.xf_id_for(index, cell.style_index);
            let style_attr = if xf_id != 0 {
                format!(" s=\"{}\"", xf_id)
            } else {
                String::new()
            };

            match &cell.value {
                duke_sheets_core::CellValue::Number(n) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{}><v>{}</v></c>",
                        cell_ref, style_attr, n
                    ));
                }
                duke_sheets_core::CellValue::String(s) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{} t=\"inlineStr\"><is><t>{}</t></is></c>",
                        cell_ref,
                        style_attr,
                        Self::escape_xml(s.as_str())
                    ));
                }
                duke_sheets_core::CellValue::Boolean(b) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{} t=\"b\"><v>{}</v></c>",
                        cell_ref,
                        style_attr,
                        if *b { 1 } else { 0 }
                    ));
                }
                duke_sheets_core::CellValue::Formula { text, .. } => {
                    let formula_text = if text.starts_with('=') {
                        &text[1..]
                    } else {
                        text.as_str()
                    };
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{}><f>{}</f></c>",
                        cell_ref,
                        style_attr,
                        Self::escape_xml(formula_text)
                    ));
                }
                duke_sheets_core::CellValue::Error(e) => {
                    content.push_str(&format!(
                        "\n            <c r=\"{}\"{} t=\"e\"><v>{}</v></c>",
                        cell_ref,
                        style_attr,
                        Self::escape_xml(e.as_str())
                    ));
                }
                duke_sheets_core::CellValue::Empty => {
                    // Preserve style-only cells
                    if xf_id != 0 {
                        content.push_str(&format!(
                            "\n            <c r=\"{}\"{} />",
                            cell_ref, style_attr
                        ));
                    }
                }
                duke_sheets_core::CellValue::SpillTarget { .. } => {
                    // SpillTarget cells are not written to the file - they are computed
                    // at runtime from the source formula's array result.
                    // In Excel's file format, dynamic array formulas use a special
                    // mechanism, but for simplicity we skip spill targets during write.
                }
            }
        }

        if current_row.is_some() {
            content.push_str("\n        </row>");
        }

        content.push_str("\n    </sheetData>");

        // Write merged cells (if any)
        let merged_regions = sheet.merged_regions();
        if !merged_regions.is_empty() {
            content.push_str(&format!(
                "\n    <mergeCells count=\"{}\">",
                merged_regions.len()
            ));
            for range in merged_regions {
                content.push_str(&format!("\n        <mergeCell ref=\"{}\"/>", range));
            }
            content.push_str("\n    </mergeCells>");
        }

        // Write conditional formatting (if any)
        Self::write_conditional_formatting(&mut content, sheet, index, style_table);

        // Write data validations (if any)
        Self::write_data_validations(&mut content, sheet);

        content.push_str("\n</worksheet>");

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    fn write_conditional_formatting(
        content: &mut String,
        sheet: &duke_sheets_core::Worksheet,
        sheet_index: usize,
        style_table: &XlsxStyleTable,
    ) {
        use duke_sheets_core::conditional_format::CfRuleType;

        let rules = sheet.conditional_formats();
        if rules.is_empty() {
            return;
        }

        // Group rules by their range sets for the <conditionalFormatting> element
        // For simplicity, we output one <conditionalFormatting> per rule
        for (rule_idx, rule) in rules.iter().enumerate() {
            if rule.ranges.is_empty() {
                continue;
            }

            // Build sqref from ranges
            let sqref: String = rule
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            content.push_str(&format!(
                "\n    <conditionalFormatting sqref=\"{}\">",
                sqref
            ));

            // Build the cfRule element
            let rule_type = rule.rule_type.xlsx_type();
            // Get dxf_id from style table (if rule has format) or from rule itself (if loaded from file)
            let dxf_id = style_table
                .dxf_id_for(sheet_index, rule_idx)
                .or(rule.dxf_id);
            let dxf_attr = dxf_id.map_or(String::new(), |id| format!(" dxfId=\"{}\"", id));
            let priority_val = rule.priority.max(1);
            let stop_if_true = if rule.stop_if_true {
                " stopIfTrue=\"1\""
            } else {
                ""
            };

            match &rule.rule_type {
                CfRuleType::CellIs {
                    operator,
                    formula1,
                    formula2,
                } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" operator=\"{}\" priority=\"{}\"{}{}>\n            <formula>{}</formula>",
                        rule_type,
                        operator.xlsx_operator(),
                        priority_val,
                        dxf_attr,
                        stop_if_true,
                        Self::escape_xml(formula1)
                    ));
                    if let Some(f2) = formula2 {
                        content.push_str(&format!(
                            "\n            <formula>{}</formula>",
                            Self::escape_xml(f2)
                        ));
                    }
                    content.push_str("\n        </cfRule>");
                }

                CfRuleType::Expression { formula } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}>\n            <formula>{}</formula>\n        </cfRule>",
                        rule_type,
                        priority_val,
                        dxf_attr,
                        stop_if_true,
                        Self::escape_xml(formula)
                    ));
                }

                CfRuleType::ColorScale { colors } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <colorScale>",
                        rule_type, priority_val, stop_if_true
                    ));
                    for cv in colors {
                        let val_attr = cv
                            .value
                            .as_ref()
                            .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                        content.push_str(&format!(
                            "\n                <cfvo type=\"{}\"{} />",
                            cv.value_type.xlsx_type(),
                            val_attr
                        ));
                    }
                    for cv in colors {
                        content.push_str(&format!(
                            "\n                <color rgb=\"{}\" />",
                            cv.color.to_argb_hex()
                        ));
                    }
                    content.push_str("\n            </colorScale>\n        </cfRule>");
                }

                CfRuleType::DataBar {
                    min_value,
                    max_value,
                    color,
                    show_value,
                    ..
                } => {
                    let show_val_attr = if *show_value { "" } else { " showValue=\"0\"" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <dataBar{}>",
                        rule_type, priority_val, stop_if_true, show_val_attr
                    ));

                    // cfvo for min
                    let min_val_attr = min_value
                        .value
                        .as_ref()
                        .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                    content.push_str(&format!(
                        "\n                <cfvo type=\"{}\"{} />",
                        min_value.value_type.xlsx_type(),
                        min_val_attr
                    ));

                    // cfvo for max
                    let max_val_attr = max_value
                        .value
                        .as_ref()
                        .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                    content.push_str(&format!(
                        "\n                <cfvo type=\"{}\"{} />",
                        max_value.value_type.xlsx_type(),
                        max_val_attr
                    ));

                    content.push_str(&format!(
                        "\n                <color rgb=\"{}\" />",
                        color.to_argb_hex()
                    ));
                    content.push_str("\n            </dataBar>\n        </cfRule>");
                }

                CfRuleType::IconSet {
                    icon_style,
                    values,
                    reverse,
                    show_value,
                } => {
                    let reverse_attr = if *reverse { " reverse=\"1\"" } else { "" };
                    let show_val_attr = if *show_value { "" } else { " showValue=\"0\"" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}>\n            <iconSet iconSet=\"{}\"{}{}>\n",
                        rule_type, priority_val, stop_if_true, icon_style.xlsx_name(), reverse_attr, show_val_attr
                    ));
                    for val in values {
                        let val_attr = val
                            .value
                            .as_ref()
                            .map_or(String::new(), |v| format!(" val=\"{}\"", v));
                        content.push_str(&format!(
                            "                <cfvo type=\"{}\"{} />\n",
                            val.value_type.xlsx_type(),
                            val_attr
                        ));
                    }
                    content.push_str("            </iconSet>\n        </cfRule>");
                }

                CfRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => {
                    let percent_attr = if *percent { " percent=\"1\"" } else { "" };
                    let bottom_attr = if *bottom { " bottom=\"1\"" } else { "" };
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" rank=\"{}\"{}{}{}{}/>",
                        rule_type,
                        priority_val,
                        rank,
                        percent_attr,
                        bottom_attr,
                        dxf_attr,
                        stop_if_true
                    ));
                }

                CfRuleType::AboveAverage {
                    above,
                    equal_average,
                    std_dev,
                } => {
                    let above_attr = if !*above { " aboveAverage=\"0\"" } else { "" };
                    let equal_attr = if *equal_average {
                        " equalAverage=\"1\""
                    } else {
                        ""
                    };
                    let std_dev_attr =
                        std_dev.map_or(String::new(), |s| format!(" stdDev=\"{}\"", s));
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}{}{}{}/>",
                        rule_type,
                        priority_val,
                        above_attr,
                        equal_attr,
                        std_dev_attr,
                        dxf_attr,
                        stop_if_true
                    ));
                }

                CfRuleType::ContainsText { text } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>NOT(ISERROR(SEARCH(\"{}\",{})))</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        Self::escape_xml(text), sqref.split(' ').next().unwrap_or("A1")
                    ));
                }

                CfRuleType::BeginsWith { text } => {
                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>LEFT({},{})=\"{}\"</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        first_cell, text.len(), Self::escape_xml(text)
                    ));
                }

                CfRuleType::EndsWith { text } => {
                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" text=\"{}\"{}{}>\n            <formula>RIGHT({},{})=\"{}\"</formula>\n        </cfRule>",
                        rule_type, priority_val, Self::escape_xml(text), dxf_attr, stop_if_true,
                        first_cell, text.len(), Self::escape_xml(text)
                    ));
                }

                CfRuleType::DuplicateValues
                | CfRuleType::UniqueValues
                | CfRuleType::ContainsBlanks
                | CfRuleType::NotContainsBlanks
                | CfRuleType::ContainsErrors
                | CfRuleType::NotContainsErrors => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\"{}{}/>",
                        rule_type, priority_val, dxf_attr, stop_if_true
                    ));
                }

                CfRuleType::TimePeriod { period } => {
                    content.push_str(&format!(
                        "\n        <cfRule type=\"{}\" priority=\"{}\" timePeriod=\"{}\"{}{}/>",
                        rule_type,
                        priority_val,
                        period.xlsx_period(),
                        dxf_attr,
                        stop_if_true
                    ));
                }
            }

            content.push_str("\n    </conditionalFormatting>");
        }
    }

    fn write_data_validations(content: &mut String, sheet: &duke_sheets_core::Worksheet) {
        use duke_sheets_core::validation::ValidationType;

        let validations = sheet.data_validations();
        if validations.is_empty() {
            return;
        }

        content.push_str(&format!(
            "\n    <dataValidations count=\"{}\">",
            validations.len()
        ));

        for validation in validations {
            if validation.ranges.is_empty() {
                continue;
            }

            // Build sqref from ranges
            let sqref: String = validation
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            let type_attr = match &validation.validation_type {
                ValidationType::None => String::new(),
                _ => format!(" type=\"{}\"", validation.validation_type.xlsx_type()),
            };

            let operator_attr = match &validation.validation_type {
                ValidationType::Whole { operator, .. }
                | ValidationType::Decimal { operator, .. }
                | ValidationType::Date { operator, .. }
                | ValidationType::Time { operator, .. }
                | ValidationType::TextLength { operator, .. } => {
                    format!(" operator=\"{}\"", operator.xlsx_operator())
                }
                _ => String::new(),
            };

            let allow_blank = if validation.allow_blank {
                " allowBlank=\"1\""
            } else {
                ""
            };
            let show_dropdown = if !validation.show_dropdown {
                " showDropDown=\"1\""
            } else {
                ""
            };
            let show_input = if validation.show_input_message {
                " showInputMessage=\"1\""
            } else {
                ""
            };
            let show_error = if validation.show_error_alert {
                " showErrorMessage=\"1\""
            } else {
                ""
            };

            let error_style = match validation.error_style {
                duke_sheets_core::ValidationErrorStyle::Stop => "",
                duke_sheets_core::ValidationErrorStyle::Warning => " errorStyle=\"warning\"",
                duke_sheets_core::ValidationErrorStyle::Information => {
                    " errorStyle=\"information\""
                }
            };

            let error_title = validation.error_title.as_ref().map_or(String::new(), |t| {
                format!(" errorTitle=\"{}\"", Self::escape_xml(t))
            });
            let error_msg = validation
                .error_message
                .as_ref()
                .map_or(String::new(), |m| {
                    format!(" error=\"{}\"", Self::escape_xml(m))
                });
            let prompt_title = validation.input_title.as_ref().map_or(String::new(), |t| {
                format!(" promptTitle=\"{}\"", Self::escape_xml(t))
            });
            let prompt_msg = validation
                .input_message
                .as_ref()
                .map_or(String::new(), |m| {
                    format!(" prompt=\"{}\"", Self::escape_xml(m))
                });

            content.push_str(&format!(
                "\n        <dataValidation{}{}{}{}{}{}{}{}{}{}{} sqref=\"{}\">",
                type_attr,
                operator_attr,
                allow_blank,
                show_dropdown,
                show_input,
                show_error,
                error_style,
                error_title,
                error_msg,
                prompt_title,
                prompt_msg,
                sqref
            ));

            // Write formulas based on validation type
            match &validation.validation_type {
                ValidationType::List { source } => {
                    // List source: either a range or comma-separated values
                    let formula = if source.starts_with('=') {
                        source[1..].to_string()
                    } else if source.contains('!')
                        || source
                            .chars()
                            .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':')
                    {
                        source.clone()
                    } else {
                        // Inline list - wrap in quotes
                        format!("\"{}\"", source)
                    };
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(&formula)
                    ));
                }
                ValidationType::Whole { value1, value2, .. }
                | ValidationType::Decimal { value1, value2, .. }
                | ValidationType::Date { value1, value2, .. }
                | ValidationType::Time { value1, value2, .. }
                | ValidationType::TextLength { value1, value2, .. } => {
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(value1)
                    ));
                    if let Some(v2) = value2 {
                        content.push_str(&format!(
                            "\n            <formula2>{}</formula2>",
                            Self::escape_xml(v2)
                        ));
                    }
                }
                ValidationType::Custom { formula } => {
                    let formula = if formula.starts_with('=') {
                        &formula[1..]
                    } else {
                        formula
                    };
                    content.push_str(&format!(
                        "\n            <formula1>{}</formula1>",
                        Self::escape_xml(formula)
                    ));
                }
                ValidationType::None => {}
            }

            content.push_str("\n        </dataValidation>");
        }

        content.push_str("\n    </dataValidations>");
    }

    fn escape_xml(s: &str) -> String {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Write worksheet relationships file (for comments, drawings, etc.)
    fn write_worksheet_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(
            &format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_index + 1),
            options,
        )?;

        let content = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments{}.xml"/>
</Relationships>"#,
            sheet_index + 1
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }

    /// Write comments file for a worksheet
    fn write_comments<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let sheet = workbook
            .worksheet(sheet_index)
            .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

        if sheet.comment_count() == 0 {
            return Ok(());
        }

        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(&format!("xl/comments{}.xml", sheet_index + 1), options)?;

        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>"#,
        );

        // Write authors
        let authors = sheet.comment_authors();
        for author in authors {
            content.push_str(&format!(
                "\n        <author>{}</author>",
                Self::escape_xml(author)
            ));
        }
        // Add empty author if no authors defined (for comments without author)
        if authors.is_empty() {
            content.push_str("\n        <author></author>");
        }

        content.push_str(
            r#"
    </authors>
    <commentList>"#,
        );

        // Collect and sort comments by cell position for consistent output
        let mut comments: Vec<_> = sheet.comments().collect();
        comments.sort_by_key(|((row, col), _)| (*row, *col));

        // Build author index map
        let author_index: std::collections::HashMap<&str, usize> = authors
            .iter()
            .enumerate()
            .map(|(i, a)| (a.as_str(), i))
            .collect();

        // Write comments
        for ((row, col), comment) in comments {
            let cell_ref = CellAddress::new(row, col).to_a1_string();
            let author_id = if comment.author.is_empty() {
                if authors.is_empty() {
                    0
                } else {
                    0 // Fallback to first author
                }
            } else {
                author_index
                    .get(comment.author.as_str())
                    .copied()
                    .unwrap_or(0)
            };

            content.push_str(&format!(
                r#"
        <comment ref="{}" authorId="{}">
            <text>
                <r>
                    <t>{}</t>
                </r>
            </text>
        </comment>"#,
                cell_ref,
                author_id,
                Self::escape_xml(&comment.text)
            ));
        }

        content.push_str(
            r#"
    </commentList>
</comments>"#,
        );

        zip.write_all(content.as_bytes())?;
        Ok(())
    }
}
