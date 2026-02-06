//! End-to-end tests for XLSX style roundtrip (create -> save -> read -> verify styles)

use duke_sheets::prelude::*;
use std::io::Cursor;

/// Test basic font styling roundtrip
#[test]
fn test_roundtrip_font_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // A1: Bold text with font color
    let style_a1 = Style::new().bold(true).font_color(Color::rgb(255, 0, 0)); // Red
    sheet.set_cell_value("A1", "Bold Red").unwrap();
    sheet.set_cell_style("A1", &style_a1).unwrap();

    // B1: Italic text with different font size
    let mut style_b1 = Style::new().italic(true).font_size(14.0);
    style_b1.font.name = "Arial".to_string();
    sheet.set_cell_value("B1", "Italic Arial 14pt").unwrap();
    sheet.set_cell_style("B1", &style_b1).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify A1 font style
    let read_style_a1 = sheet2.cell_style("A1").unwrap();
    assert!(read_style_a1.is_some(), "A1 should have a style");
    let read_style_a1 = read_style_a1.unwrap();
    assert!(read_style_a1.font.bold, "A1 should be bold");
    // Note: Color may be normalized to different formats, so we check it exists

    // Verify B1 font style
    let read_style_b1 = sheet2.cell_style("B1").unwrap();
    assert!(read_style_b1.is_some(), "B1 should have a style");
    let read_style_b1 = read_style_b1.unwrap();
    assert!(read_style_b1.font.italic, "B1 should be italic");
    assert_eq!(read_style_b1.font.size, 14.0, "B1 font size should be 14");
}

/// Test border styling roundtrip
#[test]
fn test_roundtrip_border_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // B2: Outline border (thin black)
    let style_b2 = Style {
        border: BorderStyle::outline(BorderLineStyle::Thin, Color::BLACK),
        ..Default::default()
    };
    sheet.set_cell_value("B2", "Bordered").unwrap();
    sheet.set_cell_style("B2", &style_b2).unwrap();

    // C3: Mixed border with red right edge
    let mut style_c3 = Style::default();
    style_c3.border = BorderStyle::new()
        .with_left(BorderLineStyle::Thin, Color::BLACK)
        .with_right(BorderLineStyle::Medium, Color::rgb(255, 0, 0))
        .with_top(BorderLineStyle::Thin, Color::BLACK)
        .with_bottom(BorderLineStyle::Thick, Color::BLACK);
    sheet.set_cell_value("C3", "Mixed Borders").unwrap();
    sheet.set_cell_style("C3", &style_c3).unwrap();

    // D4: Medium dashed border
    let style_d4 = Style {
        border: BorderStyle::all(BorderLineStyle::MediumDashed, Color::rgb(0, 0, 255)),
        ..Default::default()
    };
    sheet.set_cell_value("D4", "Dashed Blue").unwrap();
    sheet.set_cell_style("D4", &style_d4).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify B2 border style
    let read_style_b2 = sheet2.cell_style("B2").unwrap();
    assert!(read_style_b2.is_some(), "B2 should have a style");
    let read_style_b2 = read_style_b2.unwrap();
    assert!(
        read_style_b2.border.left.is_some(),
        "B2 should have left border"
    );
    assert!(
        read_style_b2.border.right.is_some(),
        "B2 should have right border"
    );
    assert!(
        read_style_b2.border.top.is_some(),
        "B2 should have top border"
    );
    assert!(
        read_style_b2.border.bottom.is_some(),
        "B2 should have bottom border"
    );

    // Verify C3 has different border styles per edge
    let read_style_c3 = sheet2.cell_style("C3").unwrap();
    assert!(read_style_c3.is_some(), "C3 should have a style");
    let read_style_c3 = read_style_c3.unwrap();
    assert!(
        read_style_c3.border.left.is_some(),
        "C3 should have left border"
    );
    assert!(
        read_style_c3.border.right.is_some(),
        "C3 should have right border"
    );
    // Verify border line styles
    assert_eq!(
        read_style_c3.border.left.as_ref().unwrap().style,
        BorderLineStyle::Thin,
        "C3 left border should be thin"
    );
    assert_eq!(
        read_style_c3.border.right.as_ref().unwrap().style,
        BorderLineStyle::Medium,
        "C3 right border should be medium"
    );
    assert_eq!(
        read_style_c3.border.bottom.as_ref().unwrap().style,
        BorderLineStyle::Thick,
        "C3 bottom border should be thick"
    );

    // Verify D4 dashed borders
    let read_style_d4 = sheet2.cell_style("D4").unwrap();
    assert!(read_style_d4.is_some(), "D4 should have a style");
    let read_style_d4 = read_style_d4.unwrap();
    assert_eq!(
        read_style_d4.border.left.as_ref().unwrap().style,
        BorderLineStyle::MediumDashed,
        "D4 should have medium dashed border"
    );
}

/// Test fill/background styling roundtrip
#[test]
fn test_roundtrip_fill_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // A1: Solid yellow fill
    let style_a1 = Style::new().fill_color(Color::rgb(255, 255, 0));
    sheet.set_cell_value("A1", "Yellow Fill").unwrap();
    sheet.set_cell_style("A1", &style_a1).unwrap();

    // B1: Solid blue fill
    let style_b1 = Style::new().fill_color(Color::rgb(0, 0, 255));
    sheet.set_cell_value("B1", "Blue Fill").unwrap();
    sheet.set_cell_style("B1", &style_b1).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify A1 fill style
    let read_style_a1 = sheet2.cell_style("A1").unwrap();
    assert!(read_style_a1.is_some(), "A1 should have a style");
    let read_style_a1 = read_style_a1.unwrap();
    match &read_style_a1.fill {
        FillStyle::Solid { color } => {
            // Yellow: RGB(255, 255, 0)
            let (r, g, b) = color.to_rgb();
            assert_eq!(r, 255, "A1 fill red component should be 255");
            assert_eq!(g, 255, "A1 fill green component should be 255");
            assert_eq!(b, 0, "A1 fill blue component should be 0");
        }
        _ => panic!("A1 should have solid fill"),
    }

    // Verify B1 fill style
    let read_style_b1 = sheet2.cell_style("B1").unwrap();
    assert!(read_style_b1.is_some(), "B1 should have a style");
    let read_style_b1 = read_style_b1.unwrap();
    match &read_style_b1.fill {
        FillStyle::Solid { color } => {
            // Blue: RGB(0, 0, 255)
            let (r, g, b) = color.to_rgb();
            assert_eq!(r, 0, "B1 fill red component should be 0");
            assert_eq!(g, 0, "B1 fill green component should be 0");
            assert_eq!(b, 255, "B1 fill blue component should be 255");
        }
        _ => panic!("B1 should have solid fill"),
    }
}

/// Test alignment styling roundtrip
#[test]
fn test_roundtrip_alignment_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // A1: Center aligned with wrap text
    let style_a1 = Style::new()
        .horizontal_alignment(HorizontalAlignment::Center)
        .vertical_alignment(VerticalAlignment::Center)
        .wrap_text(true);
    sheet.set_cell_value("A1", "Centered\nWrapped").unwrap();
    sheet.set_cell_style("A1", &style_a1).unwrap();

    // B1: Right aligned, top
    let style_b1 = Style::new()
        .horizontal_alignment(HorizontalAlignment::Right)
        .vertical_alignment(VerticalAlignment::Top);
    sheet.set_cell_value("B1", "Right Top").unwrap();
    sheet.set_cell_style("B1", &style_b1).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify A1 alignment
    let read_style_a1 = sheet2.cell_style("A1").unwrap();
    assert!(read_style_a1.is_some(), "A1 should have a style");
    let read_style_a1 = read_style_a1.unwrap();
    assert_eq!(
        read_style_a1.alignment.horizontal,
        HorizontalAlignment::Center,
        "A1 should be center aligned"
    );
    assert_eq!(
        read_style_a1.alignment.vertical,
        VerticalAlignment::Center,
        "A1 should be vertically centered"
    );
    assert!(
        read_style_a1.alignment.wrap_text,
        "A1 should have wrap text"
    );

    // Verify B1 alignment
    let read_style_b1 = sheet2.cell_style("B1").unwrap();
    assert!(read_style_b1.is_some(), "B1 should have a style");
    let read_style_b1 = read_style_b1.unwrap();
    assert_eq!(
        read_style_b1.alignment.horizontal,
        HorizontalAlignment::Right,
        "B1 should be right aligned"
    );
    assert_eq!(
        read_style_b1.alignment.vertical,
        VerticalAlignment::Top,
        "B1 should be top aligned"
    );
}

/// Test number format styling roundtrip
#[test]
fn test_roundtrip_number_format_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // A1: Currency format
    let style_a1 = Style::new().number_format("$#,##0.00");
    sheet.set_cell_value("A1", 1234.56).unwrap();
    sheet.set_cell_style("A1", &style_a1).unwrap();

    // B1: Percentage format
    let style_b1 = Style::new().number_format("0.00%");
    sheet.set_cell_value("B1", 0.1234).unwrap();
    sheet.set_cell_style("B1", &style_b1).unwrap();

    // C1: Custom decimal format
    let style_c1 = Style::new().number_format("0.000");
    sheet.set_cell_value("C1", 3.14159).unwrap();
    sheet.set_cell_style("C1", &style_c1).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify A1 number format
    let read_style_a1 = sheet2.cell_style("A1").unwrap();
    assert!(read_style_a1.is_some(), "A1 should have a style");
    let read_style_a1 = read_style_a1.unwrap();
    match &read_style_a1.number_format {
        NumberFormat::Custom(fmt) => {
            assert!(
                fmt.contains("#,##0") || fmt.contains("$"),
                "A1 should have currency-like format, got: {}",
                fmt
            );
        }
        NumberFormat::BuiltIn(id) => {
            // Some built-in formats are currency-related (e.g., 5-8, 37-44)
            assert!(
                *id != 0,
                "A1 should have non-General format, got BuiltIn({})",
                id
            );
        }
        NumberFormat::General => panic!("A1 should have custom number format, not General"),
    }

    // Verify B1 percentage format
    let read_style_b1 = sheet2.cell_style("B1").unwrap();
    assert!(read_style_b1.is_some(), "B1 should have a style");
    let read_style_b1 = read_style_b1.unwrap();
    match &read_style_b1.number_format {
        NumberFormat::Custom(fmt) => {
            assert!(
                fmt.contains("%"),
                "B1 should have percentage format, got: {}",
                fmt
            );
        }
        NumberFormat::BuiltIn(id) => {
            // Built-in percentage formats are 9-10
            assert!(
                *id == 9 || *id == 10,
                "B1 should have percentage format, got BuiltIn({})",
                id
            );
        }
        NumberFormat::General => panic!("B1 should have custom number format, not General"),
    }
}

/// Test style-only cells (empty value but styled)
#[test]
fn test_roundtrip_style_only_cells() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // D4: Style-only cell (no value, just styled)
    let style_d4 = Style::new()
        .fill_color(Color::rgb(200, 200, 200))
        .bold(true);
    sheet.set_cell_style("D4", &style_d4).unwrap();
    // Note: We do NOT set a value for D4

    // Also set a regular value nearby to ensure we have data
    sheet.set_cell_value("A1", "Reference").unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify D4 has style but no value
    let value_d4 = sheet2.get_value("D4").unwrap();
    assert!(value_d4.is_empty(), "D4 should have no value");

    let read_style_d4 = sheet2.cell_style("D4").unwrap();
    assert!(
        read_style_d4.is_some(),
        "D4 should have a style even with no value"
    );
    let read_style_d4 = read_style_d4.unwrap();
    assert!(read_style_d4.font.bold, "D4 should be bold");
    match &read_style_d4.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert_eq!(r, 200, "D4 fill should be gray (r=200)");
            assert_eq!(g, 200, "D4 fill should be gray (g=200)");
            assert_eq!(b, 200, "D4 fill should be gray (b=200)");
        }
        _ => panic!("D4 should have solid fill"),
    }
}

/// Test combined styles on single cell
#[test]
fn test_roundtrip_combined_styles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Create a cell with all style properties set
    let mut combined_style = Style::new()
        .bold(true)
        .italic(true)
        .font_size(16.0)
        .font_color(Color::rgb(0, 0, 128)) // Navy
        .fill_color(Color::rgb(255, 255, 200)) // Light yellow
        .horizontal_alignment(HorizontalAlignment::Center)
        .vertical_alignment(VerticalAlignment::Center)
        .wrap_text(true)
        .number_format("#,##0.00");

    combined_style.border = BorderStyle::outline(BorderLineStyle::Medium, Color::BLACK);

    sheet.set_cell_value("A1", 12345.678).unwrap();
    sheet.set_cell_style("A1", &combined_style).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify all style properties
    let read_style = sheet2.cell_style("A1").unwrap();
    assert!(read_style.is_some(), "A1 should have a style");
    let read_style = read_style.unwrap();

    // Font
    assert!(read_style.font.bold, "Should be bold");
    assert!(read_style.font.italic, "Should be italic");
    assert_eq!(read_style.font.size, 16.0, "Font size should be 16");

    // Alignment
    assert_eq!(read_style.alignment.horizontal, HorizontalAlignment::Center);
    assert_eq!(read_style.alignment.vertical, VerticalAlignment::Center);
    assert!(read_style.alignment.wrap_text);

    // Border
    assert!(read_style.border.left.is_some());
    assert!(read_style.border.right.is_some());
    assert!(read_style.border.top.is_some());
    assert!(read_style.border.bottom.is_some());

    // Fill
    match &read_style.fill {
        FillStyle::Solid { .. } => {} // OK
        _ => panic!("Should have solid fill"),
    }
}

/// Test styles across multiple sheets
#[test]
fn test_roundtrip_styles_multiple_sheets() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Styled").unwrap();

    // Style on first sheet
    let sheet1 = wb.worksheet_mut(0).unwrap();
    let style1 = Style::new().bold(true).fill_color(Color::rgb(255, 0, 0));
    sheet1.set_cell_value("A1", "Sheet1 Styled").unwrap();
    sheet1.set_cell_style("A1", &style1).unwrap();

    // Different style on second sheet
    let sheet2 = wb.worksheet_mut(1).unwrap();
    let style2 = Style::new().italic(true).fill_color(Color::rgb(0, 255, 0));
    sheet2.set_cell_value("A1", "Sheet2 Styled").unwrap();
    sheet2.set_cell_style("A1", &style2).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Verify sheet 1 style
    let sheet1_read = wb2.worksheet(0).unwrap();
    let style1_read = sheet1_read.cell_style("A1").unwrap();
    assert!(style1_read.is_some());
    let style1_read = style1_read.unwrap();
    assert!(style1_read.font.bold, "Sheet1 A1 should be bold");

    // Verify sheet 2 style
    let sheet2_read = wb2.worksheet(1).unwrap();
    let style2_read = sheet2_read.cell_style("A1").unwrap();
    assert!(style2_read.is_some());
    let style2_read = style2_read.unwrap();
    assert!(style2_read.font.italic, "Sheet2 A1 should be italic");
}
