//! Writer E2E tests: duke-sheets writes XLSX → Excel opens + re-saves → duke-sheets reads back.
//!
//! Each test builds a `Workbook` in memory, writes it with `XlsxWriter`,
//! pushes it to the Windows VM where real Excel opens it (asserting no
//! repair), re-saves to normalise the XML, pulls it back, and reads it
//! with `XlsxReader` to verify styles/values survived the round trip.

use crate::roundtrip_through_excel;
use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::{
    BorderLineStyle, BorderStyle, FillStyle, HorizontalAlignment, NumberFormat, Protection,
    Underline, VerticalAlignment,
};
use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::{
    CellAddress, CellRange, CellValue, Color, ConditionalFormatRule, DataValidation, Hyperlink,
    Style, ValidationOperator, Workbook,
};

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).unwrap(),
        CellAddress::parse(end).unwrap(),
    )
}

// Font tests

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
    let (r, g, b) = style.font.color.to_rgb().unwrap();
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
    let (r, g, b) = st.font.color.to_rgb().unwrap();
    assert!(b > 200 && r < 50, "Should be blue, got ({r}, {g}, {b})");
    assert!(
        (st.font.size - 14.0).abs() < 0.5,
        "Expected size ~14, got {}",
        st.font.size
    );
}

// Fill tests

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
            let (r, g, b) = color.to_rgb().unwrap();
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
            let (_, _, b) = color.to_rgb().unwrap();
            assert!(b > 200, "Expected blue fill");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }
    let (r, g, b) = style.font.color.to_rgb().unwrap();
    assert!(
        r > 200 && g > 200 && b > 200,
        "Expected white font, got ({r}, {g}, {b})"
    );
}

// Border tests

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
    let (_, _, b) = edge.color.to_rgb().unwrap();
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

// Alignment tests

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

// Number format tests

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

// Dimension tests

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

// Merged cells

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

// Conditional formatting

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

// Data validation

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

// Page setup / header-footer tests

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
    assert_eq!(
        ps2.odd_header.as_deref(),
        Some("&COdd Header"),
        "odd_header"
    );
    assert_eq!(
        ps2.odd_footer.as_deref(),
        Some("&COdd Footer"),
        "odd_footer"
    );
    assert_eq!(
        ps2.even_header.as_deref(),
        Some("&CEven Header"),
        "even_header"
    );
    assert_eq!(
        ps2.even_footer.as_deref(),
        Some("&CEven Footer"),
        "even_footer"
    );
    assert_eq!(
        ps2.first_header.as_deref(),
        Some("&CFirst Page"),
        "first_header"
    );
    assert_eq!(
        ps2.first_footer.as_deref(),
        Some("&CFirst Footer"),
        "first_footer"
    );

    // Verify flags
    assert!(ps2.different_odd_even, "different_odd_even");
    assert!(ps2.different_first, "different_first");
    assert!(!ps2.scale_with_doc, "scale_with_doc should be false");
    assert!(
        !ps2.align_with_margins,
        "align_with_margins should be false"
    );
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
    let workbook = duke_sheets_xlsx::XlsxReader::read_file(&fixture.host_path).expect("read");
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

// Feature parity tests (XLSX writer): verify Excel preserves the
// feature itself, not just cell values.

#[test]
fn test_write_external_url_hyperlink() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "click").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com/path".into(),
            display: Some("click".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("click"));
    let hl = s
        .hyperlink("A1")
        .expect("hyperlink must survive XLSX Excel round-trip");
    assert!(
        hl.target.contains("example.com"),
        "hyperlink target lost: {:?}",
        hl.target
    );
}

#[test]
fn test_write_internal_hyperlink() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Main").unwrap();
    wb.add_worksheet_with_name("Other").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "internal").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "#Other!B5".into(),
            display: Some("internal".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet_by_name("Main").unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("internal"));
    let hl = s
        .hyperlink("A1")
        .expect("internal hyperlink must survive XLSX Excel round-trip");
    let combined = format!("{}|{}", hl.target, hl.location.as_deref().unwrap_or(""));
    assert!(
        combined.contains("Other") && combined.contains("B5"),
        "internal target lost: target={:?} location={:?}",
        hl.target,
        hl.location
    );
}

#[test]
fn test_write_cross_sheet_formula() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 20.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 30.0)
        .unwrap();
    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=Data!A1").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(10.0))
        .unwrap();
    calc.set_cell_formula("B2", "=SUM(Data!A1:A3)").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(60.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 10.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(10), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 60.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(60), got {other:?}"),
    }
}

#[test]
fn test_write_autofilter() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Name").unwrap();
    ws.set_cell_value("B1", "Value").unwrap();
    ws.set_cell_value("A2", "alpha").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "beta").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    let mut af = AutoFilter::new(range("A1", "B3"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 1.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive XLSX Excel round-trip");
    assert_eq!(af.range.start, CellAddress::parse("A1").unwrap());
}

#[test]
fn test_write_rich_text_runs() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let red_big = RunFont {
        size: Some(16.0),
        color: Some(Color::Indexed(2)),
        ..Default::default()
    };
    ws.set_cell_value_at(
        0,
        0,
        CellValue::rich_text(vec![
            RichTextRun {
                text: "plain ".into(),
                font: None,
            },
            RichTextRun {
                text: "bold ".into(),
                font: Some(bold),
            },
            RichTextRun {
                text: "italic ".into(),
                font: Some(italic),
            },
            RichTextRun {
                text: "loud".into(),
                font: Some(red_big),
            },
        ]),
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let value = s.get_value_at(0, 0);
    assert_eq!(format!("{value}"), "plain bold italic loud");
    let runs = match &value {
        CellValue::RichText(runs) => runs,
        other => panic!("expected RichText after Excel round-trip, got {other:?}"),
    };
    assert!(
        runs.len() >= 4,
        "expected ≥4 runs, got {}: {runs:?}",
        runs.len()
    );
    let has_bold = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.bold == Some(true)));
    let has_italic = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.italic == Some(true)));
    let has_big = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if matches!(f.size, Some(s) if s >= 14.0)));
    assert!(has_bold, "bold run lost: {runs:?}");
    assert!(has_italic, "italic run lost: {runs:?}");
    assert!(has_big, "size-≥14 run lost: {runs:?}");
}

#[test]
fn test_write_named_range_formula() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 15.0)
        .unwrap();
    wb.define_name("Numbers", "Data!$A$1:$A$3").unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let names: Vec<&str> = result
        .named_ranges()
        .iter()
        .map(|nr| nr.name.as_str())
        .collect();
    assert!(
        names.iter().any(|n| n.contains("Numbers")),
        "Numbers name must survive XLSX Excel round-trip; got {names:?}"
    );
    let s = result.worksheet_by_name("Calc").unwrap();
    let v = s.get_value_at(0, 1);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 30.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(30), got {other:?}"),
    }
}

#[test]
fn test_write_print_area() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_cell_value("B2", 100.0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "B2"));
    ws.set_page_setup(ps);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let print_area = s
        .page_setup()
        .print_area
        .as_ref()
        .expect("print_area must survive XLSX Excel round-trip");
    assert_eq!(print_area.start, CellAddress::parse("A1").unwrap());
    assert_eq!(print_area.end, CellAddress::parse("B2").unwrap());
}

#[test]
fn test_write_sheet_visibility() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").unwrap();
    wb.add_worksheet_with_name("Hidden").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "p")
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);

    let result = roundtrip_through_excel(&wb);
    let hidden = result.worksheet_by_name("Hidden").unwrap();
    assert!(
        hidden.visibility() != SheetVisibility::Visible,
        "Hidden sheet must remain non-visible; got {:?}",
        hidden.visibility()
    );
}

#[test]
fn test_write_freeze_panes() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_freeze_panes(1, 1);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let freeze = s
        .freeze_panes()
        .expect("freeze panes must survive XLSX Excel round-trip");
    assert_eq!((freeze.row, freeze.col), (1, 1));
}

#[test]
fn test_write_sheet_protection() {
    use duke_sheets_core::worksheet::SheetProtection;
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "locked").unwrap();
    ws.set_protection(Some(SheetProtection {
        protected: true,
        ..Default::default()
    }));

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let prot = s
        .protection()
        .expect("sheet protection must survive XLSX Excel round-trip");
    assert!(prot.protected, "protected flag lost");
}

#[test]
fn excel_preserves_workbook_protection_and_protected_ranges() {
    use duke_sheets_core::{ProtectedRange, WorkbookProtection};

    let mut wb = Workbook::new();
    wb.set_workbook_protection(Some(WorkbookProtection {
        structure: true,
        windows: true,
        password_hash: Some(0xCAFE),
    }));

    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "editable").unwrap();
    ws.set_protected_ranges(vec![ProtectedRange {
        name: "Editable".to_string(),
        ranges: vec![range("A1", "B2"), range("D4", "D5")],
        password_hash: Some(0xCAFE),
        security_descriptor: None,
    }]);

    let result = roundtrip_through_excel(&wb);
    let protection = result
        .workbook_protection()
        .expect("workbook protection must survive XLSX Excel round-trip");
    assert!(protection.structure, "workbook structure protection lost");
    assert!(protection.windows, "workbook window protection lost");
    assert!(
        protection.password_hash.unwrap_or_default() != 0,
        "workbook password hash lost"
    );

    let ranges = result.worksheet(0).unwrap().protected_ranges();
    assert_eq!(ranges.len(), 1, "protected range count changed");
    assert_eq!(ranges[0].name, "Editable");
    assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
    assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
    assert!(
        ranges[0].password_hash.unwrap_or_default() != 0,
        "protected range password hash lost"
    );
}

#[test]
fn test_write_cell_protection_hidden_formula() {
    // Hidden formula bit must survive Excel re-save.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "formula").unwrap();
    ws.set_cell_formula("B1", "=A1").unwrap();
    ws.set_formula_result(0, 1, CellValue::Empty).unwrap();
    let mut hidden = Style::new();
    hidden.protection = Protection {
        locked: true,
        hidden: true,
    };
    ws.set_cell_style("B1", &hidden).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let b1 = s
        .cell_style_at(0, 1)
        .expect("B1 should have non-default style after round-trip");
    assert!(
        b1.protection.hidden,
        "B1 hidden flag must survive round-trip, got hidden={}",
        b1.protection.hidden
    );
}

#[test]
fn test_write_cell_protection_unlocked() {
    // User explicitly unlocks a cell. Per ECMA-376 §18.8.32, Excel's
    // effective default is `locked=true`, so the writer must emit
    // `<protection locked="0"/>` for the unlock intent to survive
    // through Excel re-save. This is a regression test against an
    // earlier writer bug where `Protection::default()` matched our
    // Rust `derive(Default)` (= locked=false), causing the writer to
    // skip emission and silently lose the user's unlock intent.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "default-locked").unwrap();
    ws.set_cell_value("B1", "explicitly-unlocked").unwrap();
    let mut unlocked = Style::new();
    unlocked.protection = Protection {
        locked: false,
        hidden: false,
    };
    ws.set_cell_style("B1", &unlocked).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let b1 = s
        .cell_style_at(0, 1)
        .expect("B1 must have non-default style for unlocked state");
    assert!(
        !b1.protection.locked,
        "B1 unlock intent lost in round-trip: locked={}",
        b1.protection.locked
    );
}

#[test]
fn excel_can_evaluate_intersection_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Intersection of A1:B3 with B2:C3 is the cells at B2, B3 = 5+8 = 13.
    ws.set_cell_formula("E1", "=SUM(A1:B3 B2:C3)").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!(
            (n - 13.0).abs() < 1e-9,
            "intersection sum drifted: E1 = {n}"
        ),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:B3") && formula.contains("B2:C3"),
        "intersection ranges lost from formula: {formula:?}"
    );
}

#[test]
fn excel_can_evaluate_union_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Union of A1:A2 and C2:C3 = {1, 4, 6, 9} = 20. Double parens are
    // required so SUM treats the comma as union, not arg separator.
    ws.set_cell_formula("E1", "=SUM((A1:A2,C2:C3))").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 20.0).abs() < 1e-9, "union sum drifted: E1 = {n}"),
        other => panic!("E1 expected Number(20), got {other:?}"),
    }
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:A2") && formula.contains("C2:C3"),
        "union ranges lost from formula: {formula:?}"
    );
}

/// A valid 68-byte 1x1 transparent PNG used as a deterministic image
/// payload for Excel parity tests.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

/// Excel must accept an XLSX with a `<xdr:oneCellAnchor>` wrapped
/// picture and preserve the OneCell anchor through SaveAs.
#[test]
fn excel_can_read_xlsx_onecell_picture_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_500_000,
            height_emu: 800_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::OneCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            width_emu: 1_500_000,
            height_emu: 800_000,
        },
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "OneCell picture must survive Excel");
    let img = &images[0];
    match &img.object.unwrap().anchor {
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            assert_eq!(from.col, 2);
            assert_eq!(from.row, 3);
            assert_eq!(*width_emu, 1_500_000);
            assert_eq!(*height_emu, 800_000);
        }
        other => panic!("expected OneCell anchor after Excel round-trip, got {other:?}"),
    }
}

/// Excel must accept an XLSX with a `<xdr:absoluteAnchor>` wrapped
/// picture and preserve the Absolute anchor through SaveAs.
#[test]
fn excel_can_read_xlsx_absolute_picture_we_emit() {
    use duke_sheets_chart::{DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 900_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::Absolute {
            x_emu: 2_500_000,
            y_emu: 1_200_000,
            width_emu: 1_000_000,
            height_emu: 900_000,
        },
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "Absolute picture must survive Excel");
    let img = &images[0];
    match &img.object.unwrap().anchor {
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            assert_eq!(*x_emu, 2_500_000);
            assert_eq!(*y_emu, 1_200_000);
            assert_eq!(*width_emu, 1_000_000);
            assert_eq!(*height_emu, 900_000);
        }
        other => panic!("expected Absolute anchor after Excel round-trip, got {other:?}"),
    }
}

/// Excel must accept the `<xdr:pic>` + `xl/media/imageN.png` emit
/// (no `Repaired` warning) and round-trip the PNG bytes verbatim
/// through SaveAs.
#[test]
fn excel_can_read_xlsx_png_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 2_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 2,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 5,
                col_offset_emu: 0,
                row: 10,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "image must survive Excel re-save");
    let img = &images[0];
    assert_eq!(img.payload.format, ImageFormat::Png);
    assert_eq!(
        img.payload.data, TEST_PNG_1X1,
        "PNG bytes must round-trip through Excel verbatim"
    );
}

/// The drawing-list z-order (image below a form control below an
/// image) survives Excel's re-save: the control's position among
/// native shapes rides its a14 placeholder twin, which Excel keeps
/// in the drawing part's document order.
#[test]
fn excel_preserves_xlsx_drawing_z_order_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingKind, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Middle".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
        .with_anchor(two_cell(1, 1, 3, 3)),
    ).unwrap();
    ws.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4))).unwrap();

    let result = roundtrip_through_excel(&wb);
    let sheet = result.worksheet(0).unwrap();
    let tags: Vec<&str> = sheet
        .drawings()
        .iter()
        .map(|object| match &object.kind {
            DrawingKind::Image(_) => "image",
            DrawingKind::FormControl(_) => "control",
            other => panic!("unexpected drawing kind after Excel round-trip: {other:?}"),
        })
        .collect();
    assert_eq!(
        tags,
        vec!["image", "control", "image"],
        "z-order must survive Excel re-save"
    );
    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images[0].object.unwrap().meta.name.as_deref(), Some("Below"));
    assert_eq!(images[1].object.unwrap().meta.name.as_deref(), Some("Above"));
    assert_eq!(
        sheet
            .form_controls()
            .next()
            .unwrap()
            .payload
            .caption_text()
            .as_deref(),
        Some("Middle")
    );
}

/// One of every Forms control kind survives the Excel XLSX
/// round-trip: worksheet <controls> block, ctrlProps parts, VML
/// shapes (captions), and the a14 drawing-part twins all re-read
/// intact; the Repaired check inside the roundtrip helper proves
/// Excel accepts the twin markup we emit.
#[test]
fn excel_can_read_xlsx_form_controls_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

    let anchor = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        ws.set_cell_value_at(i as u32, 7, *item).expect("list item");
    }
    let kinds: Vec<FormControlKind> = vec![
        FormControlKind::Button {
            caption: "Run Report".into(),
        },
        FormControlKind::Checkbox {
            caption: "Enable audit".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: true,
        },
        FormControlKind::Checkbox {
            caption: "Tri state".into(),
            state: CheckState::Mixed,
            cell_link: None,
            no_3d: true,
        },
        FormControlKind::OptionButton {
            caption: "Opt A".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$3".to_string()),
            first_in_group: false,
            no_3d: true,
        },
        FormControlKind::OptionButton {
            caption: "Opt B".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            first_in_group: false,
            no_3d: true,
        },
        FormControlKind::Label {
            caption: "Status label".into(),
        },
        FormControlKind::GroupBox {
            caption: "Choices".into(),
            no_3d: true,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$5".to_string()),
            selection: ListSelection::Single,
            selected: vec![3],
            no_3d: true,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![1, 3],
            no_3d: true,
        },
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$4".to_string()),
            selected: Some(2),
            lines: 6,
            no_3d: true,
        },
        FormControlKind::Scrollbar {
            value: 40,
            min: 5,
            max: 95,
            increment: 2,
            page: 10,
            horizontal: false,
            cell_link: Some("$D$6".to_string()),
        },
        FormControlKind::Spinner {
            value: 12,
            min: 0,
            max: 30,
            increment: 3,
            cell_link: Some("$D$7".to_string()),
        },
    ];
    let count = kinds.len();
    let expected = kinds.clone();
    for (i, kind) in kinds.into_iter().enumerate() {
        let row = 1 + 2 * i as u32;
        ws.add_form_control(FormControl::new(kind), anchor(1, row, 3, row + 1)).unwrap();
    }
    assert_eq!(wb.sync_form_control_links(), 6);

    let result = roundtrip_through_excel(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.get_value("D2").unwrap(), CellValue::Boolean(true));
    assert_eq!(sheet.get_value("D3").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.get_value("D4").unwrap(), CellValue::Number(3.0));
    assert_eq!(sheet.get_value("D5").unwrap(), CellValue::Number(4.0));
    assert_eq!(sheet.get_value("D6").unwrap(), CellValue::Number(40.0));
    assert_eq!(sheet.get_value("D7").unwrap(), CellValue::Number(12.0));
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), count, "every control survives Excel");
    for (i, control) in controls.iter().enumerate() {
        let mut want = expected[i].clone();
        if let FormControlKind::OptionButton { first_in_group, .. } = &mut want {
            // Writer recomputes grouping; first radio heads the group.
            *first_in_group = i == 3;
        }
        assert_eq!(
            control.payload.kind, want,
            "control {i} kind mismatch after Excel"
        );
    }
}

#[test]
fn excel_preserves_xlsx_custom_metric_control_anchor_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "metric anchor".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        }),
        DrawingAnchor::OneCell {
            from: CellMarker::default(),
            width_emu: 609_600,
            height_emu: 190_500,
        },
    ).unwrap();

    let result = roundtrip_through_excel(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    match &drawn.object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 0));
            assert_eq!((from.row, from.row_offset_emu), (0, 0));
            assert_eq!((to.col, to.col_offset_emu), (0, 609_600));
            assert_eq!((to.row, to.row_offset_emu), (0, 190_500));
        }
        other => panic!("expected Excel-resaved TwoCell control anchor, got {other:?}"),
    }
}

#[test]
fn excel_preserves_xlsx_control_visual_metadata_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::style::Underline;
    use duke_sheets_core::{
        CheckState, ControlText, DrawingObject, FormControl, FormControlKind, HorizontalAlignment,
        VerticalAlignment,
    };

    let text = ControlText {
        runs: vec![
            RichTextRun::with_font(
                "Red ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(9.0),
                    color: Some(Color::rgb(255, 0, 0)),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Blue",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    color: Some(Color::rgb(0, 0, 255)),
                    italic: Some(true),
                    underline: Some(Underline::Single),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Right),
        vertical_alignment: Some(VerticalAlignment::Bottom),
    };
    let control = FormControl::new(FormControlKind::Checkbox {
        caption: text,
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
    .with_macro_name("RunProbe");
    let mut object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 0,
            row: 1,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 4,
            col_offset_emu: 0,
            row: 3,
            row_offset_emu: 0,
        },
        edit_as: None,
    });
    object.meta.name = Some("Visual Probe".into());
    object.meta.alt_text = Some("Visual probe alternative".into());
    object.meta.title = Some("Visual probe title".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Visual Probe"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("Visual probe alternative")
    );
    assert_eq!(
        drawn.object.unwrap().meta.title.as_deref(),
        Some("Visual probe title")
    );
    assert_eq!(drawn.payload.caption_text().as_deref(), Some("Red Blue"));
    assert_eq!(drawn.payload.macro_name.as_deref(), Some("RunProbe"));
    let caption = drawn.payload.caption().unwrap();
    assert_eq!(
        caption.horizontal_alignment,
        Some(HorizontalAlignment::Right)
    );
    assert_eq!(caption.vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(caption.runs.len(), 2);
    let red = caption.runs[0].font.as_ref().unwrap();
    assert_eq!(red.name.as_deref(), Some("Segoe UI"));
    assert_eq!(red.size, Some(9.0));
    assert_eq!(red.color, Some(Color::rgb(255, 0, 0)));
    assert_eq!(red.bold, Some(true));
    let blue = caption.runs[1].font.as_ref().unwrap();
    assert_eq!(blue.name.as_deref(), Some("Arial"));
    assert_eq!(blue.size, Some(12.0));
    assert_eq!(blue.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(blue.italic, Some(true));
    assert_eq!(blue.underline, Some(Underline::Single));
}

/// Drawing-object hidden flags survive Excel's XLSX re-save: a
/// hidden image rides its `cNvPr@hidden="1"`, a hidden form control
/// rides the VML shape's `visibility:hidden` style, and the visible
/// siblings stay visible.
#[test]
fn excel_preserves_hidden_drawing_flags_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };
    let checkbox = |caption: &str| {
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: caption.into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Shown").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        png("Ghost")
            .with_anchor(two_cell(2, 2, 4, 4))
            .with_hidden(true),
    ).unwrap();
    ws.add_drawing(checkbox("Visible box").with_anchor(two_cell(4, 4, 6, 6))).unwrap();
    ws.add_drawing(
        checkbox("Cloaked box")
            .with_anchor(two_cell(6, 6, 8, 8))
            .with_hidden(true),
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let sheet = result.worksheet(0).unwrap();

    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images.len(), 2, "both images survive Excel re-save");
    let image_hidden = |name: &str| {
        images
            .iter()
            .find(|i| i.object.unwrap().meta.name.as_deref() == Some(name))
            .unwrap_or_else(|| panic!("image {name:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(!image_hidden("Shown"), "visible image must stay visible");
    assert!(
        image_hidden("Ghost"),
        "hidden image must survive Excel re-save with hidden intact"
    );

    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 2, "both controls survive Excel re-save");
    let control_hidden = |caption: &str| {
        controls
            .iter()
            .find(|c| c.payload.caption_text().as_deref() == Some(caption))
            .unwrap_or_else(|| panic!("control {caption:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(
        !control_hidden("Visible box"),
        "visible control must stay visible"
    );
    assert!(
        control_hidden("Cloaked box"),
        "hidden control must survive Excel re-save with hidden intact"
    );
}

/// A cell comment survives Excel's XLSX re-save: text, author, cell,
/// and popup visibility (visible notes ride the VML shape's
/// `visibility:visible` style, hidden ones the default).
#[test]
fn excel_can_read_xlsx_comment_we_emit() {
    use duke_sheets_core::{CellComment, DrawingObject};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("B2", "reviewed").unwrap();
    ws.set_comment_at(1, 1, CellComment::new("Reviewer", "Check this figure")).unwrap();
    ws.add_drawing(
        DrawingObject::comment(3, 3, CellComment::new("Reviewer", "Always shown"))
            .with_hidden(false),
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let sheet = result.worksheet(0).unwrap();

    let hidden = sheet.comment_at(1, 1).expect("comment survives at B2");
    assert!(
        hidden.plain_text().contains("Check this figure"),
        "comment text lost: {:?}",
        hidden.plain_text()
    );
    assert!(
        hidden.author.contains("Reviewer"),
        "comment author lost: {:?}",
        hidden.author
    );
    assert!(sheet.comment_at(0, 0).is_none(), "comment cell moved");

    let visibility: Vec<(u32, u16, bool)> = sheet
        .comments_drawn()
        .map(|cr| (cr.row, cr.col, cr.object.meta.hidden))
        .collect();
    assert!(
        visibility.contains(&(1, 1, true)),
        "B2 note must stay hover-only: {visibility:?}"
    );
    assert!(
        visibility.contains(&(3, 3, false)),
        "D4 note must stay visible: {visibility:?}"
    );
}

#[test]
fn excel_preserves_xlsx_basic_shape_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{
        DrawingObject, DrawingText, Shape, ShapeFill, ShapeGeometry, ShapeLine,
    };

    let text = DrawingText {
        runs: vec![
            RichTextRun::with_font(
                "Bold ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(10.0),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Italic",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    italic: Some(true),
                    color: Some(Color::rgb(0, 0, 255)),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Center),
        vertical_alignment: Some(VerticalAlignment::Center),
    };
    let shape = Shape::rectangle()
        .with_fill(ShapeFill::Solid(Color::rgb(255, 0, 0)))
        .with_line(ShapeLine {
            color: Some(Color::rgb(0, 0, 255)),
            width_emu: Some(25_400),
            dash_style: Some("dash".into()),
            no_fill: false,
        })
        .with_text(text)
        .with_rotation(900_000)
        .with_flip_h(true);
    let mut object = DrawingObject::shape(shape).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            row: 2,
            ..CellMarker::default()
        },
        to: CellMarker {
            col: 5,
            row: 8,
            ..CellMarker::default()
        },
        edit_as: None,
    });
    object.meta.name = Some("Status panel".into());
    object.meta.alt_text = Some("red status rectangle".into());
    object.meta.title = Some("Status".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel(&workbook);
    let drawn = result.worksheet(0).unwrap().shapes().next().expect("shape");
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Status panel"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("red status rectangle")
    );
    assert_eq!(drawn.object.unwrap().meta.title.as_deref(), Some("Status"));
    assert_eq!(drawn.payload.geometry, ShapeGeometry::Preset("rect".into()));
    assert_eq!(drawn.payload.fill, ShapeFill::Solid(Color::rgb(255, 0, 0)));
    assert_eq!(drawn.payload.line.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(drawn.payload.line.width_emu, Some(25_400));
    assert_eq!(drawn.payload.line.dash_style.as_deref(), Some("dash"));
    assert_eq!(drawn.payload.rotation, 900_000);
    assert!(drawn.payload.flip_h);
    let text = drawn.payload.text.as_ref().expect("shape text");
    assert_eq!(text.plain_text(), "Bold Italic");
    assert_eq!(text.horizontal_alignment, Some(HorizontalAlignment::Center));
    assert_eq!(text.vertical_alignment, Some(VerticalAlignment::Center));
    assert_eq!(
        text.runs[0].font.as_ref().unwrap().name.as_deref(),
        Some("Segoe UI")
    );
    assert_eq!(text.runs[0].font.as_ref().unwrap().bold, Some(true));
    assert_eq!(text.runs[0].font.as_ref().unwrap().size, Some(10.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().name.as_deref(),
        Some("Arial")
    );
    assert_eq!(text.runs[1].font.as_ref().unwrap().italic, Some(true));
    assert_eq!(text.runs[1].font.as_ref().unwrap().size, Some(12.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().color,
        Some(Color::rgb(0, 0, 255))
    );
}

/// Unmodeled ClientData children we replay on a modeled control kind
/// survive a real Excel XLSX round-trip: the Repaired check inside the
/// helper proves our raw emission does not corrupt the legacy VML
/// part, and the re-read model proves Excel re-saved `x:Accel` rather
/// than discarding it. Pinned Excel normalization: `x:Disabled` on a
/// worksheet checkbox is accepted without repair but dropped from
/// Excel's own re-save, so its survival cannot be asserted here (the
/// in-process round-trip in `unknown_controls_anchors` covers our
/// side of it).
// features: Form control unmodeled ClientData passthrough
#[test]
fn excel_preserves_unmodeled_client_data_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    let mut control = FormControl::new(FormControlKind::Checkbox {
        caption: "Audit".into(),
        state: CheckState::Checked,
        cell_link: None,
        no_3d: true,
    });
    control.raw_client_data = vec![b"<x:Disabled/>".to_vec(), b"<x:Accel>65</x:Accel>".to_vec()];
    ws.add_form_control(
        control,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel(&wb);
    let control = result
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("checkbox survives");
    assert!(matches!(
        control.payload.kind,
        duke_sheets_core::FormControlKind::Checkbox { .. }
    ));
    let raws: Vec<String> = control
        .payload
        .raw_client_data
        .iter()
        .map(|raw| String::from_utf8_lossy(raw).into_owned())
        .collect();
    assert!(
        raws.iter()
            .any(|raw| raw.contains("Accel") && raw.contains("65")),
        "x:Accel value survives Excel's re-save: {raws:?}"
    );
}

/// Excel parity for rich comment text in XLSX: the `<r>/<rPr>` runs
/// we emit in the comments part must survive Excel's re-save with
/// their run boundary and bold flag.
// features: Rich text in comments
#[test]
fn excel_preserves_rich_comment_runs_we_emit() {
    use duke_sheets_core::rich_text::{RichTextRun, RunFont};
    use duke_sheets_core::{CellComment, DrawingText};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.set_comment_at(
        0,
        0,
        CellComment {
            author: "Reviewer".to_string(),
            text: DrawingText {
                runs: vec![
                    RichTextRun {
                        text: "Bold lead".to_string(),
                        font: Some(RunFont {
                            bold: Some(true),
                            ..RunFont::default()
                        }),
                    },
                    RichTextRun {
                        text: " then plain".to_string(),
                        font: None,
                    },
                ],
                ..DrawingText::default()
            },
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let comment = result
        .worksheet(0)
        .unwrap()
        .comment_at(0, 0)
        .expect("comment must survive Excel re-save");
    assert_eq!(comment.plain_text(), "Bold lead then plain");
    let bold_run = comment
        .text
        .runs
        .iter()
        .find(|run| run.text.contains("Bold lead"))
        .expect("bold run boundary survives Excel re-save");
    assert_eq!(
        bold_run.font.as_ref().and_then(|font| font.bold),
        Some(true),
        "bold formatting lost: {:?}",
        comment.text.runs
    );
}

/// Excel parity for unmodeled `formControlPr` attribute passthrough
/// on a modeled kind: our emit must not trip the Repaired dialog and
/// the control must stay modeled. Attribute survival itself is
/// Excel's call; this is a spec-compliance smoke check for the raw
/// emission path.
// features: Form control unmodeled ctrlProps passthrough
#[test]
fn excel_accepts_unmodeled_ctrl_props_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    let mut control = FormControl::new(FormControlKind::Checkbox {
        caption: "Audit".into(),
        state: CheckState::Checked,
        cell_link: None,
        no_3d: true,
    });
    control.raw_properties = vec![("customFlag".to_string(), "kept".to_string())];
    ws.add_form_control(
        control,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let control = result
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("checkbox survives");
    let FormControlKind::Checkbox { state, .. } = &control.payload.kind else {
        panic!("checkbox must stay modeled, got {:?}", control.payload.kind);
    };
    assert_eq!(*state, CheckState::Checked);
}

/// A chartEx built through the model - not read from a file - must open
/// in Excel and keep its waterfall series and subtotal bars through
/// Excel's own re-save.
///
/// Excel validates a chartEx's chart style and chart colour style
/// siblings and refuses the whole workbook when either is missing or
/// when the style part is not a complete `CT_ChartStyle`. Charts built
/// through the model have no raw parts to replay, so the writer
/// generates them; before it did, every such workbook failed to open.
/// Nothing here depends on a corpus file.
// features: ChartEx: Waterfall
#[test]
fn excel_opens_a_model_built_waterfall_chart_ex() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::DrawingObject;

    // `cx:series` must not be self-closing: the parser only collects a
    // series on its End event.
    let chart_ex = duke_sheets_chart::parse::parse_chart_ex_xml(
        &br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall" uniqueId="{1D8F9C4E-1C1B-4A5F-9C6B-2E7A0F3B5D11}"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Series1</cx:v></cx:txData></cx:tx><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="0"/><cx:idx val="2"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>"#[..],
    )
    .expect("parse chartEx");

    assert_eq!(
        chart_ex.plot_area.series[0]
            .layout_properties
            .as_ref()
            .and_then(|l| l.subtotals.clone()),
        Some(vec![0, 2]),
        "fixture must carry subtotals for the survival assertion to mean anything"
    );

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, (cat, val)) in [("a", 1.0), ("b", 2.0), ("c", 3.0)].iter().enumerate() {
        let row = row + 1;
        ws.set_cell_value(&format!("A{row}"), *cat).unwrap();
        ws.set_cell_value(&format!("B{row}"), *val).unwrap();
    }
    ws.add_drawing(
        DrawingObject::chart_ex(chart_ex).with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 0,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 10,
                col_offset_emu: 0,
                row: 15,
                row_offset_emu: 0,
            },
            edit_as: None,
        }),
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);

    let chart = result
        .worksheet(0)
        .unwrap()
        .charts_ex()
        .next()
        .expect("chartEx must survive Excel's re-save")
        .payload;
    let series = chart
        .plot_area
        .series
        .first()
        .expect("waterfall series survives");
    assert_eq!(series.layout, duke_sheets_chart::ChartExLayout::Waterfall);
    assert_eq!(
        series
            .layout_properties
            .as_ref()
            .and_then(|l| l.subtotals.clone()),
        Some(vec![0, 2]),
        "subtotal bars must survive Excel's re-save"
    );
}
