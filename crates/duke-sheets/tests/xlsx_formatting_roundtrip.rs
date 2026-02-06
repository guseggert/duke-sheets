//! End-to-end tests for XLSX data validation, conditional formatting, and comments roundtrip

use duke_sheets::prelude::*;
use std::io::Cursor;

/// Test data validation list roundtrip
#[test]
fn test_roundtrip_data_validation_list() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some data
    sheet.set_cell_value("A1", "Select:").unwrap();

    // Add list validation to B1:B10
    let validation = DataValidation::list("Yes,No,Maybe")
        .with_range(CellRange::parse("B1:B10").unwrap())
        .with_input_message("Choose", "Select a value from the list")
        .with_error_message("Error", "Invalid selection");
    sheet.add_data_validation(validation);

    // Verify we have the validation
    assert_eq!(sheet.data_validation_count(), 1);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify validation was read back
    assert_eq!(
        sheet2.data_validation_count(),
        1,
        "Should have 1 data validation"
    );

    let dv = sheet2.data_validation_at(0, 1); // B1
    assert!(dv.is_some(), "B1 should have data validation");

    if let Some(dv) = dv {
        match &dv.validation_type {
            ValidationType::List { source } => {
                assert!(source.contains("Yes"), "List should contain 'Yes'");
            }
            _ => panic!("Expected List validation type"),
        }
    }
}

/// Test data validation whole number roundtrip
#[test]
fn test_roundtrip_data_validation_number() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add number validation
    let validation = DataValidation::whole_number_between(ValidationOperator::Between, "1", "100")
        .with_range(CellRange::parse("A1:A5").unwrap())
        .with_allow_blank(false);
    sheet.add_data_validation(validation);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.data_validation_count(), 1);

    let dv = sheet2.data_validation_at(0, 0); // A1
    assert!(dv.is_some(), "A1 should have data validation");
}

/// Test conditional formatting cell is rule roundtrip
#[test]
fn test_roundtrip_conditional_format_cell_is() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some test data
    for i in 0..5 {
        sheet
            .set_cell_value_at(i, 0, (i as f64 + 1.0) * 25.0)
            .unwrap();
    }

    // Add conditional format: highlight cells > 50
    let rule = ConditionalFormatRule::cell_is_greater_than("50")
        .with_range(CellRange::parse("A1:A5").unwrap())
        .with_priority(1);
    sheet.add_conditional_format(rule);

    // Verify we have the rule
    assert_eq!(sheet.conditional_format_count(), 1);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify conditional format was read back
    assert_eq!(
        sheet2.conditional_format_count(),
        1,
        "Should have 1 conditional format"
    );

    let rules = sheet2.conditional_formats_at(0, 0); // A1
    assert!(!rules.is_empty(), "A1 should have conditional formatting");

    if let Some(rule) = rules.first() {
        match &rule.rule_type {
            CfRuleType::CellIs {
                operator, formula1, ..
            } => {
                assert_eq!(*operator, CfOperator::GreaterThan);
                assert_eq!(formula1, "50");
            }
            _ => panic!("Expected CellIs rule type"),
        }
    }
}

/// Test conditional formatting expression rule roundtrip
#[test]
fn test_roundtrip_conditional_format_expression() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add conditional format with custom formula
    let rule = ConditionalFormatRule::expression("MOD(A1,2)=0")
        .with_range(CellRange::parse("A1:A10").unwrap())
        .with_priority(1);
    sheet.add_conditional_format(rule);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.conditional_format_count(), 1);

    let rules = sheet2.conditional_formats_at(0, 0);
    assert!(!rules.is_empty());

    if let Some(rule) = rules.first() {
        match &rule.rule_type {
            CfRuleType::Expression { formula } => {
                assert!(formula.contains("MOD"), "Formula should contain MOD");
            }
            _ => panic!("Expected Expression rule type"),
        }
    }
}

/// Test multiple validations and conditional formats
#[test]
fn test_roundtrip_multiple_rules() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add multiple data validations
    sheet.add_data_validation(
        DataValidation::list("A,B,C").with_range(CellRange::parse("A1:A5").unwrap()),
    );
    sheet.add_data_validation(
        DataValidation::whole_number(ValidationOperator::GreaterThan, "0")
            .with_range(CellRange::parse("B1:B5").unwrap()),
    );

    // Add multiple conditional formats (use two CellIs rules for simplicity)
    sheet.add_conditional_format(
        ConditionalFormatRule::cell_is_greater_than("100")
            .with_range(CellRange::parse("C1:C10").unwrap()),
    );
    sheet.add_conditional_format(
        ConditionalFormatRule::cell_is_less_than("0")
            .with_range(CellRange::parse("D1:D10").unwrap()),
    );

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify counts
    assert_eq!(
        sheet2.data_validation_count(),
        2,
        "Should have 2 data validations"
    );
    assert_eq!(
        sheet2.conditional_format_count(),
        2,
        "Should have 2 conditional formats"
    );
}

/// Test cell comments roundtrip
#[test]
fn test_roundtrip_cell_comments() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some cell data
    sheet.set_cell_value("A1", "Data with comment").unwrap();
    sheet.set_cell_value("B2", 42.0).unwrap();

    // Add comments to cells
    sheet
        .set_comment("A1", CellComment::new("John Doe", "This is a note"))
        .unwrap();
    sheet
        .set_comment("B2", CellComment::new("Jane Smith", "Review this value"))
        .unwrap();

    // Verify we have comments
    assert_eq!(sheet.comment_count(), 2);
    assert_eq!(sheet.comment_authors().len(), 2);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify comments were read back
    assert_eq!(sheet2.comment_count(), 2, "Should have 2 comments");

    // Check A1 comment
    let comment_a1 = sheet2.comment("A1").unwrap();
    assert!(comment_a1.is_some(), "A1 should have a comment");
    if let Some(c) = comment_a1 {
        assert_eq!(c.author, "John Doe");
        assert_eq!(c.text, "This is a note");
    }

    // Check B2 comment
    let comment_b2 = sheet2.comment("B2").unwrap();
    assert!(comment_b2.is_some(), "B2 should have a comment");
    if let Some(c) = comment_b2 {
        assert_eq!(c.author, "Jane Smith");
        assert_eq!(c.text, "Review this value");
    }
}

/// Test cell comments without author
#[test]
fn test_roundtrip_cell_comments_no_author() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add comment without author
    sheet
        .set_comment("A1", CellComment::text_only("Anonymous note"))
        .unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify comment was read back
    assert_eq!(sheet2.comment_count(), 1, "Should have 1 comment");

    let comment = sheet2.comment("A1").unwrap();
    assert!(comment.is_some(), "A1 should have a comment");
    if let Some(c) = comment {
        assert!(c.author.is_empty(), "Author should be empty");
        assert_eq!(c.text, "Anonymous note");
    }
}

/// Test multiple sheets with comments
#[test]
fn test_roundtrip_comments_multiple_sheets() {
    let mut wb = Workbook::new();

    // Add comments to first sheet
    let sheet1 = wb.worksheet_mut(0).unwrap();
    sheet1.set_cell_value("A1", "Sheet1 data").unwrap();
    sheet1
        .set_comment("A1", CellComment::new("Author1", "Comment on sheet 1"))
        .unwrap();

    // Add second sheet with comments
    wb.add_worksheet_with_name("Sheet2").unwrap();
    let sheet2 = wb.worksheet_mut(1).unwrap();
    sheet2.set_cell_value("B2", "Sheet2 data").unwrap();
    sheet2
        .set_comment("B2", CellComment::new("Author2", "Comment on sheet 2"))
        .unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Verify sheet 1 comments
    let sheet1 = wb2.worksheet(0).unwrap();
    assert_eq!(sheet1.comment_count(), 1);
    let c1 = sheet1.comment("A1").unwrap().unwrap();
    assert_eq!(c1.author, "Author1");
    assert_eq!(c1.text, "Comment on sheet 1");

    // Verify sheet 2 comments
    let sheet2 = wb2.worksheet(1).unwrap();
    assert_eq!(sheet2.comment_count(), 1);
    let c2 = sheet2.comment("B2").unwrap().unwrap();
    assert_eq!(c2.author, "Author2");
    assert_eq!(c2.text, "Comment on sheet 2");
}

/// Test color scale conditional formatting roundtrip
#[test]
fn test_roundtrip_color_scale() {
    use duke_sheets::prelude::Color;

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some test data
    for i in 0..10 {
        sheet
            .set_cell_value_at(i, 0, (i as f64 + 1.0) * 10.0)
            .unwrap();
    }

    // Add a 2-color scale (red to green)
    let rule = ConditionalFormatRule::color_scale_2(
        Color::rgb(255, 0, 0), // Red for min
        Color::rgb(0, 255, 0), // Green for max
    )
    .with_range(CellRange::parse("A1:A10").unwrap());
    sheet.add_conditional_format(rule);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.conditional_format_count(), 1);

    let rules = sheet2.conditional_formats_at(0, 0);
    assert!(!rules.is_empty(), "A1 should have conditional formatting");

    if let Some(rule) = rules.first() {
        match &rule.rule_type {
            CfRuleType::ColorScale { colors } => {
                assert_eq!(colors.len(), 2, "Should have 2 colors");
                // Check min color is red
                assert_eq!(colors[0].color, Color::rgb(255, 0, 0));
                // Check max color is green
                assert_eq!(colors[1].color, Color::rgb(0, 255, 0));
            }
            _ => panic!("Expected ColorScale rule type, got {:?}", rule.rule_type),
        }
    }
}

/// Test data bar conditional formatting roundtrip
#[test]
fn test_roundtrip_data_bar() {
    use duke_sheets::prelude::Color;

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some test data
    for i in 0..5 {
        sheet
            .set_cell_value_at(i, 0, (i as f64 + 1.0) * 20.0)
            .unwrap();
    }

    // Add a data bar (blue)
    let rule = ConditionalFormatRule::data_bar(Color::rgb(99, 142, 198))
        .with_range(CellRange::parse("A1:A5").unwrap());
    sheet.add_conditional_format(rule);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.conditional_format_count(), 1);

    let rules = sheet2.conditional_formats_at(0, 0);
    assert!(!rules.is_empty(), "A1 should have conditional formatting");

    if let Some(rule) = rules.first() {
        match &rule.rule_type {
            CfRuleType::DataBar { color, .. } => {
                assert_eq!(*color, Color::rgb(99, 142, 198));
            }
            _ => panic!("Expected DataBar rule type, got {:?}", rule.rule_type),
        }
    }
}

/// Test icon set conditional formatting roundtrip
#[test]
fn test_roundtrip_icon_set() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some test data
    for i in 0..5 {
        sheet
            .set_cell_value_at(i, 0, (i as f64 + 1.0) * 20.0)
            .unwrap();
    }

    // Add a traffic light icon set
    let rule = ConditionalFormatRule::icon_set(IconSetStyle::TrafficLights3)
        .with_range(CellRange::parse("A1:A5").unwrap());
    sheet.add_conditional_format(rule);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.conditional_format_count(), 1);

    let rules = sheet2.conditional_formats_at(0, 0);
    assert!(!rules.is_empty(), "A1 should have conditional formatting");

    if let Some(rule) = rules.first() {
        match &rule.rule_type {
            CfRuleType::IconSet {
                icon_style, values, ..
            } => {
                assert_eq!(*icon_style, IconSetStyle::TrafficLights3);
                assert_eq!(values.len(), 3, "Should have 3 threshold values");
            }
            _ => panic!("Expected IconSet rule type, got {:?}", rule.rule_type),
        }
    }
}

/// Test DXF (differential format) roundtrip with conditional formatting styles
#[test]
fn test_roundtrip_dxf_styles() {
    use duke_sheets::prelude::{Color, FillStyle, Style};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add some test data
    for i in 0..5 {
        sheet
            .set_cell_value_at(i, 0, (i as f64 + 1.0) * 30.0)
            .unwrap();
    }

    // Create a style for the conditional format
    let highlight_style = Style::new()
        .fill_color(Color::rgb(255, 199, 206)) // Light red fill
        .font_color(Color::rgb(156, 0, 6)) // Dark red text
        .bold(true);

    // Add conditional format with style: highlight cells > 100
    let rule = ConditionalFormatRule::cell_is_greater_than("100")
        .with_range(CellRange::parse("A1:A5").unwrap())
        .with_format(highlight_style.clone());
    sheet.add_conditional_format(rule);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.conditional_format_count(), 1);

    let rules = sheet2.conditional_formats_at(0, 0);
    assert!(!rules.is_empty(), "A1 should have conditional formatting");

    if let Some(rule) = rules.first() {
        // Check that the rule has a format (DXF style was preserved)
        assert!(
            rule.format.is_some(),
            "Rule should have a format/style from DXF"
        );

        if let Some(ref format) = rule.format {
            // Check that the fill color was preserved (may be Rgb or Argb)
            match &format.fill {
                FillStyle::Solid { color } => {
                    // Colors may be stored as Rgb or Argb, so check the actual values
                    let (r, g, b) = match color {
                        Color::Rgb { r, g, b } => (*r, *g, *b),
                        Color::Argb { r, g, b, .. } => (*r, *g, *b),
                        _ => panic!("Expected Rgb or Argb color"),
                    };
                    assert_eq!((r, g, b), (255, 199, 206), "Fill color should be light red");
                }
                _ => panic!("Expected Solid fill style"),
            }

            // Check that font is bold
            assert!(format.font.bold, "Font should be bold");

            // Check font color (may be Rgb or Argb)
            let (r, g, b) = match &format.font.color {
                Color::Rgb { r, g, b } => (*r, *g, *b),
                Color::Argb { r, g, b, .. } => (*r, *g, *b),
                _ => panic!("Expected Rgb or Argb font color"),
            };
            assert_eq!((r, g, b), (156, 0, 6), "Font color should be dark red");
        }
    }
}

/// Test multiple rules with different DXF styles
#[test]
fn test_roundtrip_multiple_dxf_styles() {
    use duke_sheets::prelude::{Color, FillStyle, Style};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add test data
    for i in 0..10 {
        sheet.set_cell_value_at(i, 0, (i as f64) * 10.0).unwrap();
    }

    // Rule 1: Green for values >= 70
    let green_style = Style::new()
        .fill_color(Color::rgb(198, 239, 206)) // Light green
        .font_color(Color::rgb(0, 97, 0)); // Dark green
    let rule1 = ConditionalFormatRule::cell_is_greater_than("69")
        .with_range(CellRange::parse("A1:A10").unwrap())
        .with_format(green_style)
        .with_priority(1);
    sheet.add_conditional_format(rule1);

    // Rule 2: Red for values < 30
    let red_style = Style::new()
        .fill_color(Color::rgb(255, 199, 206)) // Light red
        .font_color(Color::rgb(156, 0, 6)); // Dark red
    let rule2 = ConditionalFormatRule::cell_is_less_than("30")
        .with_range(CellRange::parse("A1:A10").unwrap())
        .with_format(red_style)
        .with_priority(2);
    sheet.add_conditional_format(rule2);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(
        sheet2.conditional_format_count(),
        2,
        "Should have 2 conditional format rules"
    );

    // Check that both rules have formats
    let rules = sheet2.conditional_formats();
    for (i, rule) in rules.iter().enumerate() {
        assert!(rule.format.is_some(), "Rule {} should have a format", i);
    }
}
