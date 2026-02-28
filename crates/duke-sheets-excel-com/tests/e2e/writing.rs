//! Writer E2E tests: duke-sheets writes XLSX → Excel opens + re-saves → duke-sheets reads back.
//!
//! Each test builds a `Workbook` in memory, writes it with `XlsxWriter`,
//! pushes it to the Windows VM where real Excel opens it (asserting no
//! repair), re-saves to normalise the XML, pulls it back, and reads it
//! with `XlsxReader` to verify styles/values survived the round trip.

use crate::roundtrip_through_excel;
use duke_sheets_core::style::{
    BorderLineStyle, BorderStyle, FillStyle, HorizontalAlignment, NumberFormat, Underline,
    VerticalAlignment,
};
use duke_sheets_core::{
    CellRange, Color, ConditionalFormatRule, DataValidation, Style, ValidationOperator, Workbook,
};

// =========================================================================
// Font tests
// =========================================================================

#[test]
fn test_write_font_bold() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Bold").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().bold(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Font should be bold");
}

#[test]
fn test_write_font_italic() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Italic").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().italic(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.italic, "Font should be italic");
}

#[test]
fn test_write_font_size() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Big").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_size(20.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        (style.font.size - 20.0).abs() < 0.5,
        "Expected font size ~20, got {}",
        style.font.size
    );
}

#[test]
fn test_write_font_name() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Courier").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_name("Courier New"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.name, "Courier New");
}

#[test]
fn test_write_font_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Red text").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_color(Color::rgb(255, 0, 0)))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g < 50 && b < 50,
        "Expected red font, got ({r}, {g}, {b})"
    );
}

#[test]
fn test_write_font_strikethrough() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Strike").unwrap();
    let mut style = Style::new();
    style.font.strikethrough = true;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.font.strikethrough, "Font should be strikethrough");
}

#[test]
fn test_write_font_underline() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Underline").unwrap();
    let mut style = Style::new();
    style.font.underline = Underline::Single;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.font.underline, Underline::Single);
}

#[test]
fn test_write_font_combination() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Combo").unwrap();
    let mut style = Style::new()
        .bold(true)
        .italic(true)
        .font_size(14.0)
        .font_color(Color::rgb(0, 0, 255));
    style.font.underline = Underline::Single;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.font.bold, "Should be bold");
    assert!(st.font.italic, "Should be italic");
    assert_eq!(st.font.underline, Underline::Single);
    let (r, g, b) = st.font.color.to_rgb();
    assert!(b > 200 && r < 50, "Should be blue, got ({r}, {g}, {b})");
    assert!(
        (st.font.size - 14.0).abs() < 0.5,
        "Expected size ~14, got {}",
        st.font.size
    );
}

// =========================================================================
// Fill tests
// =========================================================================

#[test]
fn test_write_solid_fill() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Red fill").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().fill_color(Color::rgb(255, 0, 0)))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g < 50 && b < 50,
                "Expected red fill, got ({r}, {g}, {b})"
            );
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }
}

#[test]
fn test_write_fill_with_font_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "White on Blue").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new()
                .fill_color(Color::rgb(0, 0, 255))
                .font_color(Color::rgb(255, 255, 255)),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");

    match &style.fill {
        FillStyle::Solid { color } => {
            let (_, _, b) = color.to_rgb();
            assert!(b > 200, "Expected blue fill");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g > 200 && b > 200,
        "Expected white font, got ({r}, {g}, {b})"
    );
}

// =========================================================================
// Border tests
// =========================================================================

#[test]
fn test_write_border_thin_all() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Bordered").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::all(BorderLineStyle::Thin, Color::Auto),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "left border");
    assert!(style.border.right.is_some(), "right border");
    assert!(style.border.top.is_some(), "top border");
    assert!(style.border.bottom.is_some(), "bottom border");
}

#[test]
fn test_write_border_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Blue border").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::all(BorderLineStyle::Thin, Color::rgb(0, 0, 255)),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let edge = style.border.left.as_ref().expect("left border");
    let (_, _, b) = edge.color.to_rgb();
    assert!(b > 200, "Expected blue border color, got b={b}");
}

#[test]
fn test_write_border_individual() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Left only").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::new().with_left(BorderLineStyle::Thin, Color::Auto),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "left border should be set");
    assert!(style.border.right.is_none(), "right border should be empty");
    assert!(style.border.top.is_none(), "top border should be empty");
    assert!(
        style.border.bottom.is_none(),
        "bottom border should be empty"
    );
}

// =========================================================================
// Alignment tests
// =========================================================================

#[test]
fn test_write_alignment_horizontal_center() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Centered").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().horizontal_alignment(HorizontalAlignment::Center),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Center);
}

#[test]
fn test_write_alignment_horizontal_right() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Right").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().horizontal_alignment(HorizontalAlignment::Right),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Right);
}

#[test]
fn test_write_alignment_vertical_top() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Top").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().vertical_alignment(VerticalAlignment::Top),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.vertical, VerticalAlignment::Top);
}

#[test]
fn test_write_alignment_wrap() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Wrap\nText").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().wrap_text(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.wrap_text, "wrap_text should be true");
}

#[test]
fn test_write_alignment_shrink() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Shrink").unwrap();
    let mut style = Style::new();
    style.alignment.shrink_to_fit = true;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.alignment.shrink_to_fit, "shrink_to_fit should be true");
}

#[test]
fn test_write_alignment_rotation() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Rotated").unwrap();
    let mut style = Style::new();
    style.alignment.rotation = 45;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.alignment.rotation, 45);
}

#[test]
fn test_write_alignment_indent() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Indented").unwrap();
    let mut style = Style::new();
    style.alignment.indent = 2;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.alignment.indent, 2);
}

// =========================================================================
// Number format tests
// =========================================================================

#[test]
fn test_write_number_format_percent() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 0.1234).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("0.00%"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(fmt.contains('%'), "Expected percentage format, got '{fmt}'");
}

#[test]
fn test_write_number_format_date() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 45000.0).unwrap(); // a date serial
    sheet
        .set_cell_style(
            "A1",
            &Style {
                number_format: NumberFormat::BuiltIn(14), // m/d/yyyy
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_ne!(
        style.number_format,
        NumberFormat::General,
        "Should have a date format"
    );
}

#[test]
fn test_write_number_format_currency() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 1234.56).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("#,##0.00"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("#,##0"),
        "Expected thousands format, got '{fmt}'"
    );
}

#[test]
fn test_write_number_format_custom() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 3.14159).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("0.000"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("0.000"),
        "Expected '0.000' format, got '{fmt}'"
    );
}

// =========================================================================
// Dimension tests
// =========================================================================

#[test]
fn test_write_row_height() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Tall row").unwrap();
    sheet.set_row_height(0, 40.0);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let height = s.row_height(0);
    assert!(
        (height - 40.0).abs() < 1.0,
        "Expected row height ~40, got {height}"
    );
}

#[test]
fn test_write_column_width() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Wide column").unwrap();
    sheet.set_column_width(0, 25.0);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let width = s.column_width(0);
    // Excel may adjust width slightly during re-save
    assert!(
        (width - 25.0).abs() < 2.0,
        "Expected column width ~25, got {width}"
    );
}

#[test]
fn test_write_hidden_row() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Visible").unwrap();
    sheet.set_cell_value("A2", "Hidden").unwrap();
    sheet.set_row_hidden(1, true);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    assert!(!s.is_row_hidden(0), "Row 0 should be visible");
    assert!(s.is_row_hidden(1), "Row 1 should be hidden");
}

// =========================================================================
// Merged cells
// =========================================================================

#[test]
fn test_write_merged_cells() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Merged region").unwrap();
    sheet
        .merge_cells(&CellRange::parse("A1:C1").unwrap())
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let regions = s.merged_regions();
    assert!(
        !regions.is_empty(),
        "Should have at least one merged region"
    );
    let expected = CellRange::parse("A1:C1").unwrap();
    assert!(
        regions.iter().any(|r| *r == expected),
        "Should contain A1:C1, got {regions:?}"
    );
}

// =========================================================================
// Conditional formatting
// =========================================================================

#[test]
fn test_write_conditional_format() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 60.0).unwrap();
    sheet.set_cell_value("A3", 90.0).unwrap();

    let rule = ConditionalFormatRule::cell_is_greater_than("50")
        .with_range(CellRange::parse("A1:A3").unwrap())
        .with_format(Style::new().fill_color(Color::rgb(0, 255, 0)));
    sheet.add_conditional_format(rule);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let cfs = s.conditional_formats();
    assert!(
        !cfs.is_empty(),
        "Should have at least one conditional format rule"
    );
}

// =========================================================================
// Data validation
// =========================================================================

#[test]
fn test_write_data_validation_list() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Apple").unwrap();

    let dv = DataValidation::list("Apple,Banana,Cherry")
        .with_range(CellRange::parse("A1").unwrap())
        .with_input_message("Pick a fruit", "Choose from the list")
        .with_error_message("Invalid", "Must be a fruit");
    sheet.add_data_validation(dv);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let dvs = s.data_validations();
    assert!(!dvs.is_empty(), "Should have at least one data validation");
}

#[test]
fn test_write_data_validation_number() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 50.0).unwrap();

    let dv = DataValidation::whole_number(ValidationOperator::Between, "1")
        .with_range(CellRange::parse("A1").unwrap());
    sheet.add_data_validation(dv);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let dvs = s.data_validations();
    assert!(!dvs.is_empty(), "Should have at least one data validation");
}

// =========================================================================
// Page setup / header-footer tests
// =========================================================================

#[test]
fn test_write_header_footer_even_first_and_flags() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Page setup test").unwrap();

    let mut ps = sheet.page_setup().clone();
    ps.odd_header = Some("&COdd Header".to_string());
    ps.odd_footer = Some("&COdd Footer".to_string());
    ps.even_header = Some("&CEven Header".to_string());
    ps.even_footer = Some("&CEven Footer".to_string());
    ps.first_header = Some("&CFirst Page".to_string());
    ps.first_footer = Some("&CFirst Footer".to_string());
    ps.different_odd_even = true;
    ps.different_first = true;
    ps.scale_with_doc = false;
    ps.align_with_margins = false;
    sheet.set_page_setup(ps);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let ps2 = s.page_setup();

    // Verify all header/footer strings survived Excel roundtrip
    assert_eq!(ps2.odd_header.as_deref(), Some("&COdd Header"), "odd_header");
    assert_eq!(ps2.odd_footer.as_deref(), Some("&COdd Footer"), "odd_footer");
    assert_eq!(ps2.even_header.as_deref(), Some("&CEven Header"), "even_header");
    assert_eq!(ps2.even_footer.as_deref(), Some("&CEven Footer"), "even_footer");
    assert_eq!(ps2.first_header.as_deref(), Some("&CFirst Page"), "first_header");
    assert_eq!(ps2.first_footer.as_deref(), Some("&CFirst Footer"), "first_footer");

    // Verify flags
    assert!(ps2.different_odd_even, "different_odd_even");
    assert!(ps2.different_first, "different_first");
    assert!(!ps2.scale_with_doc, "scale_with_doc should be false");
    assert!(!ps2.align_with_margins, "align_with_margins should be false");
}

#[test]
fn test_read_header_footer_from_excel() {
    let bridge = crate::excel_bridge();
    let fixture = crate::temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        crate::ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Set headers/footers via PageSetup properties (left/center/right)
        wb.set_page_setup_property("CenterHeader", serde_json::Value::from("Center Header"))
            .expect("set CenterHeader");
        wb.set_page_setup_property("LeftFooter", serde_json::Value::from("Left Footer"))
            .expect("set LeftFooter");
        wb.set_page_setup_property("RightFooter", serde_json::Value::from("Right Footer"))
            .expect("set RightFooter");

        // Enable different first page flag
        wb.set_page_setup_property(
            "DifferentFirstPageHeaderFooter",
            serde_json::Value::from(true),
        )
        .expect("set DifferentFirstPageHeaderFooter");

        // Disable scale with doc
        wb.set_page_setup_property("ScaleWithDocHeaderFooter", serde_json::Value::from(false))
            .expect("set ScaleWithDocHeaderFooter");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    crate::pull_file_from_vm(&fixture);
    let workbook =
        duke_sheets_xlsx::XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let ps = sheet.page_setup();

    // Odd header: Excel wraps CenterHeader in &C code
    let odd_header = ps.odd_header.as_deref().unwrap_or("");
    assert!(
        odd_header.contains("Center Header"),
        "odd_header should contain 'Center Header', got: {odd_header:?}"
    );

    // Odd footer: left and right sections combined
    let odd_footer = ps.odd_footer.as_deref().unwrap_or("");
    assert!(
        odd_footer.contains("Left Footer"),
        "odd_footer should contain 'Left Footer', got: {odd_footer:?}"
    );
    assert!(
        odd_footer.contains("Right Footer"),
        "odd_footer should contain 'Right Footer', got: {odd_footer:?}"
    );

    // Flags
    assert!(ps.different_first, "different_first should be true");
    assert!(!ps.scale_with_doc, "scale_with_doc should be false");

    crate::cleanup_fixture(&fixture);
}
