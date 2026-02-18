//! End-to-end tests for XLSX roundtrip (create -> save -> read -> verify)

use duke_sheets::prelude::*;
use std::io::Cursor;

/// Test basic roundtrip with numeric values
#[test]
fn test_roundtrip_numbers() {
    // Create a workbook with numeric data
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 42.0).unwrap();
    sheet.set_cell_value("B1", 3.14159).unwrap();
    sheet.set_cell_value("C1", -100.5).unwrap();
    sheet.set_cell_value("A2", 0.0).unwrap();
    sheet.set_cell_value("B2", 1e10).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify
    assert_eq!(sheet2.get_value("A1").unwrap().as_number(), Some(42.0));
    assert!((sheet2.get_value("B1").unwrap().as_number().unwrap() - 3.14159).abs() < 1e-10);
    assert_eq!(sheet2.get_value("C1").unwrap().as_number(), Some(-100.5));
    assert_eq!(sheet2.get_value("A2").unwrap().as_number(), Some(0.0));
    assert_eq!(sheet2.get_value("B2").unwrap().as_number(), Some(1e10));
}

/// Test basic roundtrip with string values
#[test]
fn test_roundtrip_strings() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", "Hello, World!").unwrap();
    sheet.set_cell_value("B1", "").unwrap(); // Empty string
    sheet.set_cell_value("C1", "Special: <>&\"'").unwrap(); // XML entities
    sheet.set_cell_value("A2", "Multi\nLine").unwrap();
    sheet.set_cell_value("B2", "Unicode: \u{1F600}").unwrap(); // Emoji

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify
    assert_eq!(
        sheet2.get_value("A1").unwrap().as_string(),
        Some("Hello, World!")
    );
    // Note: empty string cells might become Empty in roundtrip
    assert_eq!(
        sheet2.get_value("C1").unwrap().as_string(),
        Some("Special: <>&\"'")
    );
    assert_eq!(
        sheet2.get_value("A2").unwrap().as_string(),
        Some("Multi\nLine")
    );
    assert_eq!(
        sheet2.get_value("B2").unwrap().as_string(),
        Some("Unicode: \u{1F600}")
    );
}

/// Test roundtrip with boolean values
#[test]
fn test_roundtrip_booleans() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", true).unwrap();
    sheet.set_cell_value("B1", false).unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify
    assert_eq!(sheet2.get_value("A1").unwrap().as_bool(), Some(true));
    assert_eq!(sheet2.get_value("B1").unwrap().as_bool(), Some(false));
}

/// Test roundtrip with formulas
#[test]
fn test_roundtrip_formulas() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 20.0).unwrap();
    sheet.set_cell_formula("A3", "=SUM(A1:A2)").unwrap();
    sheet.set_cell_formula("B1", "=A1*2").unwrap();
    sheet
        .set_cell_formula("C1", "=IF(A1>5,\"Yes\",\"No\")")
        .unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify formulas are preserved
    assert!(sheet2.get_value("A3").unwrap().is_formula());
    assert_eq!(
        sheet2.get_value("A3").unwrap().formula_text(),
        Some("=SUM(A1:A2)")
    );
    assert_eq!(
        sheet2.get_value("B1").unwrap().formula_text(),
        Some("=A1*2")
    );
    assert_eq!(
        sheet2.get_value("C1").unwrap().formula_text(),
        Some("=IF(A1>5,\"Yes\",\"No\")")
    );
}

/// Test roundtrip with multiple worksheets
#[test]
fn test_roundtrip_multiple_sheets() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.add_worksheet_with_name("Summary").unwrap();

    // Populate first sheet
    let sheet1 = wb.worksheet_mut(0).unwrap();
    sheet1.set_cell_value("A1", "Sheet 1 Data").unwrap();

    // Populate second sheet
    let sheet2 = wb.worksheet_mut(1).unwrap();
    sheet2.set_cell_value("A1", "Data Sheet").unwrap();
    sheet2.set_cell_value("B1", 100.0).unwrap();

    // Populate third sheet
    let sheet3 = wb.worksheet_mut(2).unwrap();
    sheet3.set_cell_value("A1", "Summary").unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Verify structure
    assert_eq!(wb2.sheet_count(), 3);
    assert_eq!(wb2.worksheet(0).unwrap().name(), "Sheet1");
    assert_eq!(wb2.worksheet(1).unwrap().name(), "Data");
    assert_eq!(wb2.worksheet(2).unwrap().name(), "Summary");

    // Verify content
    assert_eq!(
        wb2.worksheet(0)
            .unwrap()
            .get_value("A1")
            .unwrap()
            .as_string(),
        Some("Sheet 1 Data")
    );
    assert_eq!(
        wb2.worksheet(1)
            .unwrap()
            .get_value("A1")
            .unwrap()
            .as_string(),
        Some("Data Sheet")
    );
    assert_eq!(
        wb2.worksheet(1)
            .unwrap()
            .get_value("B1")
            .unwrap()
            .as_number(),
        Some(100.0)
    );
    assert_eq!(
        wb2.worksheet(2)
            .unwrap()
            .get_value("A1")
            .unwrap()
            .as_string(),
        Some("Summary")
    );
}

/// Test roundtrip with mixed cell types
#[test]
fn test_roundtrip_mixed_types() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Row 1: Headers (strings)
    sheet.set_cell_value("A1", "Name").unwrap();
    sheet.set_cell_value("B1", "Value").unwrap();
    sheet.set_cell_value("C1", "Active").unwrap();

    // Row 2: Mixed data
    sheet.set_cell_value("A2", "Item 1").unwrap();
    sheet.set_cell_value("B2", 42.5).unwrap();
    sheet.set_cell_value("C2", true).unwrap();

    // Row 3: Formula
    sheet.set_cell_formula("B3", "=SUM(B2:B2)").unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify all types
    assert_eq!(sheet2.get_value("A1").unwrap().as_string(), Some("Name"));
    assert_eq!(sheet2.get_value("B2").unwrap().as_number(), Some(42.5));
    assert_eq!(sheet2.get_value("C2").unwrap().as_bool(), Some(true));
    assert!(sheet2.get_value("B3").unwrap().is_formula());
}

/// Test roundtrip with large row/column indices
#[test]
fn test_roundtrip_large_indices() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Set values at various positions
    sheet.set_cell_value_at(0, 0, "A1").unwrap(); // A1
    sheet.set_cell_value_at(100, 25, "Z101").unwrap(); // Z101
    sheet.set_cell_value_at(999, 51, "AZ1000").unwrap(); // AZ1000
    sheet.set_cell_value_at(9999, 701, "ZZ10000").unwrap(); // ZZ10000

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify
    assert_eq!(sheet2.get_value_at(0, 0).as_string(), Some("A1"));
    assert_eq!(sheet2.get_value_at(100, 25).as_string(), Some("Z101"));
    assert_eq!(sheet2.get_value_at(999, 51).as_string(), Some("AZ1000"));
    assert_eq!(sheet2.get_value_at(9999, 701).as_string(), Some("ZZ10000"));
}

/// Test roundtrip preserves empty cells
#[test]
fn test_roundtrip_sparse_data() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Create sparse data
    sheet.set_cell_value("A1", "Start").unwrap();
    sheet.set_cell_value("Z50", "Middle").unwrap();
    sheet.set_cell_value("A100", "End").unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify set cells
    assert_eq!(sheet2.get_value("A1").unwrap().as_string(), Some("Start"));
    assert_eq!(sheet2.get_value("Z50").unwrap().as_string(), Some("Middle"));
    assert_eq!(sheet2.get_value("A100").unwrap().as_string(), Some("End"));

    // Verify empty cells remain empty
    assert!(sheet2.get_value("B1").unwrap().is_empty());
    assert!(sheet2.get_value("A2").unwrap().is_empty());
}

/// Test empty workbook roundtrip
#[test]
fn test_roundtrip_empty_workbook() {
    let wb = Workbook::new();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Should have at least one sheet
    assert!(wb2.sheet_count() >= 1);
}

/// Test roundtrip with special sheet names
#[test]
fn test_roundtrip_special_sheet_names() {
    let mut wb = Workbook::empty();
    wb.add_worksheet_with_name("Data 2024").unwrap();
    wb.add_worksheet_with_name("Q1 Report").unwrap();
    wb.add_worksheet_with_name("Sales-Summary").unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Verify sheet names
    assert_eq!(wb2.sheet_count(), 3);
    assert_eq!(wb2.worksheet(0).unwrap().name(), "Data 2024");
    assert_eq!(wb2.worksheet(1).unwrap().name(), "Q1 Report");
    assert_eq!(wb2.worksheet(2).unwrap().name(), "Sales-Summary");
}

/// Test row heights and column widths roundtrip
#[test]
fn test_roundtrip_row_heights_column_widths() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", "Tall row").unwrap();
    sheet.set_row_height(0, 30.0);
    sheet.set_row_height(2, 50.0);
    sheet.set_column_width(0, 20.0);
    sheet.set_column_width(2, 5.0);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert!(
        (sheet2.row_height(0) - 30.0).abs() < 0.1,
        "Row 0 height should be ~30, got {}",
        sheet2.row_height(0)
    );
    assert!(
        (sheet2.row_height(2) - 50.0).abs() < 0.1,
        "Row 2 height should be ~50, got {}",
        sheet2.row_height(2)
    );
    assert!(
        (sheet2.column_width(0) - 20.0).abs() < 0.1,
        "Column A width should be ~20, got {}",
        sheet2.column_width(0)
    );
    assert!(
        (sheet2.column_width(2) - 5.0).abs() < 0.1,
        "Column C width should be ~5, got {}",
        sheet2.column_width(2)
    );
}

/// Test hidden rows/columns roundtrip
#[test]
fn test_roundtrip_hidden_rows_columns() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", "Visible").unwrap();
    sheet.set_cell_value("A2", "Hidden row").unwrap();
    sheet.set_row_hidden(1, true);
    sheet.set_column_hidden(1, true);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert!(!sheet2.is_row_hidden(0), "Row 0 should not be hidden");
    assert!(sheet2.is_row_hidden(1), "Row 1 should be hidden");
    assert!(!sheet2.is_column_hidden(0), "Col A should not be hidden");
    assert!(sheet2.is_column_hidden(1), "Col B should be hidden");
}
