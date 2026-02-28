//! End-to-end tests for XLSX roundtrip (create -> save -> read -> verify)

use duke_sheets::prelude::*;
use duke_sheets_core::style::Underline;
use duke_sheets_core::{PageOrientation, SplitPanes};
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

/// Test row/column outline metadata roundtrip
#[test]
fn test_roundtrip_outline_metadata() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A2", "Grouped row").unwrap();
    sheet.set_row_outline_level(1, 2);
    sheet.set_row_collapsed(1, true);
    sheet.set_column_outline_level(2, 3);
    sheet.set_column_collapsed(2, true);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.row_outline_level(1), 2);
    assert!(sheet2.is_row_collapsed(1));
    assert_eq!(sheet2.column_outline_level(2), 3);
    assert!(sheet2.is_column_collapsed(2));
}

/// Test split pane + sheet selection metadata roundtrip
#[test]
fn test_roundtrip_split_panes_and_selection() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_split_panes(Some(SplitPanes {
        x_split: 2000.0,
        y_split: 3000.0,
        top_left: Some((3, 2)),
        active_pane: Some("bottomRight".to_string()),
    }));
    sheet.set_zoom_scale(Some(90));
    sheet.set_selection_active_cell(4, 3); // D5
    sheet.set_selection_range(Some(CellRange::parse("D5:E6").unwrap()));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let split = sheet2.split_panes().expect("split panes should roundtrip");
    assert_eq!(split.x_split, 2000.0);
    assert_eq!(split.y_split, 3000.0);
    assert_eq!(split.top_left, Some((3, 2)));
    assert_eq!(split.active_pane.as_deref(), Some("bottomRight"));
    assert_eq!(sheet2.zoom_scale(), Some(90));
    assert_eq!(sheet2.selection_active_cell(), Some((4, 3)));
    assert_eq!(
        sheet2.selection_range().map(|r| r.to_string()),
        Some("D5:E6".to_string())
    );
}

/// Test page setup and header/footer roundtrip
#[test]
fn test_roundtrip_page_setup_and_header_footer() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut ps = sheet.page_setup().clone();
    ps.paper_size = 9;
    ps.orientation = PageOrientation::Landscape;
    ps.scale = 85;
    ps.fit_to_width = Some(1);
    ps.fit_to_height = Some(2);
    ps.left_margin = 0.5;
    ps.right_margin = 0.6;
    ps.top_margin = 0.7;
    ps.bottom_margin = 0.8;
    ps.header_margin = 0.2;
    ps.footer_margin = 0.25;
    ps.print_gridlines = true;
    ps.print_headings = true;
    ps.odd_header = Some("&LLeft&CCenter".to_string());
    ps.odd_footer = Some("&RPage &P".to_string());
    sheet.set_page_setup(ps);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    let ps2 = sheet2.page_setup();

    assert_eq!(ps2.paper_size, 9);
    assert!(matches!(ps2.orientation, PageOrientation::Landscape));
    assert_eq!(ps2.scale, 85);
    assert_eq!(ps2.fit_to_width, Some(1));
    assert_eq!(ps2.fit_to_height, Some(2));
    assert!((ps2.left_margin - 0.5).abs() < 1e-9);
    assert!((ps2.right_margin - 0.6).abs() < 1e-9);
    assert!((ps2.top_margin - 0.7).abs() < 1e-9);
    assert!((ps2.bottom_margin - 0.8).abs() < 1e-9);
    assert!((ps2.header_margin - 0.2).abs() < 1e-9);
    assert!((ps2.footer_margin - 0.25).abs() < 1e-9);
    assert!(ps2.print_gridlines);
    assert!(ps2.print_headings);
    assert_eq!(ps2.odd_header.as_deref(), Some("&LLeft&CCenter"));
    assert_eq!(ps2.odd_footer.as_deref(), Some("&RPage &P"));
}

// --- Formula cached value roundtrip tests ---

/// Test roundtrip of formula with numeric cached value
#[test]
fn test_roundtrip_formula_cached_number() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_value_at(
            0,
            0,
            CellValue::Formula {
                text: "=1+2".to_string(),
                cached_value: Some(Box::new(CellValue::Number(3.0))),
                array_result: None,
            },
        )
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(val.formula_text(), Some("=1+2"));
    match val.effective_value() {
        CellValue::Number(n) => assert!((n - 3.0).abs() < 1e-10),
        other => panic!("Expected Number(3.0), got {:?}", other),
    }
}

/// Test roundtrip of formula with string cached value
#[test]
fn test_roundtrip_formula_cached_string() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_value_at(
            0,
            0,
            CellValue::Formula {
                text: "=CONCAT(\"hello\")".to_string(),
                cached_value: Some(Box::new(CellValue::string("hello"))),
                array_result: None,
            },
        )
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(val.formula_text(), Some("=CONCAT(\"hello\")"));
    match val.effective_value() {
        CellValue::String(s) => assert_eq!(s.as_str(), "hello"),
        other => panic!("Expected String(\"hello\"), got {:?}", other),
    }
}

/// Test roundtrip of formula with boolean cached value
#[test]
fn test_roundtrip_formula_cached_boolean() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_value_at(
            0,
            0,
            CellValue::Formula {
                text: "=TRUE()".to_string(),
                cached_value: Some(Box::new(CellValue::Boolean(true))),
                array_result: None,
            },
        )
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(val.formula_text(), Some("=TRUE()"));
    match val.effective_value() {
        CellValue::Boolean(b) => assert!(*b),
        other => panic!("Expected Boolean(true), got {:?}", other),
    }
}

/// Test roundtrip of formula with error cached value
#[test]
fn test_roundtrip_formula_cached_error() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_value_at(
            0,
            0,
            CellValue::Formula {
                text: "=1/0".to_string(),
                cached_value: Some(Box::new(CellValue::Error(CellError::Div0))),
                array_result: None,
            },
        )
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(val.formula_text(), Some("=1/0"));
    match val.effective_value() {
        CellValue::Error(e) => assert_eq!(e.as_str(), "#DIV/0!"),
        other => panic!("Expected Error(Div0), got {:?}", other),
    }
}

/// Test roundtrip of formula with no cached value (regression)
#[test]
fn test_roundtrip_formula_no_cached_value() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_value_at(
            0,
            0,
            CellValue::Formula {
                text: "=SUM(B1:B10)".to_string(),
                cached_value: None,
                array_result: None,
            },
        )
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(val.formula_text(), Some("=SUM(B1:B10)"));
    // No cached value expected — effective_value is the formula itself
    assert!(val.is_formula());
    match val {
        CellValue::Formula { cached_value, .. } => assert!(cached_value.is_none()),
        _ => panic!("Expected Formula variant"),
    }
}

/// Test roundtrip of named ranges (definedNames)
#[test]
fn test_roundtrip_named_ranges() {
    use duke_sheets::named_range::{NameScope, NamedRange};

    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 100.0)
        .unwrap();

    // Add workbook-scoped named range
    wb.named_ranges_mut()
        .define_or_update(NamedRange::workbook_scope("TaxRate", "Sheet1!$A$1"));

    // Add sheet-scoped named range
    wb.named_ranges_mut()
        .define_or_update(NamedRange::sheet_scope("LocalName", "Sheet1!$B$1:$B$10", 0));

    // Add hidden named range with comment
    let mut hidden_nr = NamedRange::workbook_scope("_xlnm.Print_Area", "Sheet1!$A$1:$D$20");
    hidden_nr.hidden = true;
    hidden_nr.comment = Some("Print area".to_string());
    wb.named_ranges_mut().define_or_update(hidden_nr);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Verify named ranges survive roundtrip
    let nr = wb2.named_ranges();
    assert_eq!(nr.len(), 3);

    let tax = nr.get("TaxRate", 0).unwrap();
    assert_eq!(tax.refers_to, "Sheet1!$A$1");
    assert_eq!(tax.scope, NameScope::Workbook);

    let local = nr.get("LocalName", 0).unwrap();
    assert_eq!(local.refers_to, "Sheet1!$B$1:$B$10");
    assert!(matches!(local.scope, NameScope::Sheet(0)));

    let print = nr.get("_xlnm.Print_Area", 0).unwrap();
    assert_eq!(print.refers_to, "Sheet1!$A$1:$D$20");
    assert!(print.hidden);
    assert_eq!(print.comment.as_deref(), Some("Print area"));
}

/// Test that worksheet XML elements are in spec-canonical order (ISO 29500-1 §18.3.1.99)
#[test]
fn test_worksheet_element_ordering() {
    use std::io::{Cursor, Read};
    use std::str::FromStr;

    // Create a workbook with various features to ensure all elements are present
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Add cell data
    sheet.set_cell_value("A1", "Header").unwrap();
    sheet.set_cell_value("A2", 42.0).unwrap();
    sheet.set_cell_value("B1", "Value").unwrap();
    sheet.set_cell_value("B2", 3.14).unwrap();

    // Add merged cells
    let range = CellRange::from_str("C1:D1").unwrap();
    sheet.merge_cells(&range).unwrap();

    // Set print options
    let mut ps = sheet.page_setup().clone();
    ps.print_gridlines = true;
    sheet.set_page_setup(ps);

    // Set custom margins
    let mut ps = sheet.page_setup().clone();
    ps.left_margin = 1.0;
    ps.right_margin = 1.0;
    ps.top_margin = 1.0;
    ps.bottom_margin = 1.0;
    sheet.set_page_setup(ps);

    // Set zoom to ensure sheetViews is emitted
    sheet.set_zoom_scale(Some(110));

    // Set header
    let mut ps = sheet.page_setup().clone();
    ps.odd_header = Some("Test Header".to_string());
    sheet.set_page_setup(ps);

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Extract raw sheet XML from ZIP and verify element ordering
    let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
    let mut sheet_xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut sheet_xml)
        .unwrap();

    // Verify spec-canonical element ordering: elements that appear must be
    // in this order relative to each other (some may be absent).
    let canonical_order = [
        "<dimension",
        "<sheetViews",
        "<sheetFormatPr",
        "<sheetData",
        "<mergeCells",
        "<printOptions",
        "<pageMargins",
        "<pageSetup",
        "<headerFooter",
    ];
    let positions: Vec<(&str, usize)> = canonical_order
        .iter()
        .filter_map(|tag| sheet_xml.find(tag).map(|pos| (*tag, pos)))
        .collect();
    // At least dimension, sheetViews, sheetFormatPr, sheetData, mergeCells,
    // printOptions, pageMargins, headerFooter should be present
    assert!(
        positions.len() >= 7,
        "expected at least 7 elements, found {}: {:?}",
        positions.len(),
        positions,
    );
    for window in positions.windows(2) {
        assert!(
            window[0].1 < window[1].1,
            "element {} (pos {}) should appear before {} (pos {})",
            window[0].0,
            window[0].1,
            window[1].0,
            window[1].1,
        );
    }

    // Also verify roundtrip data preservation
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    assert_eq!(sheet2.get_value("A1").unwrap().as_string(), Some("Header"));
    assert_eq!(sheet2.get_value("A2").unwrap().as_number(), Some(42.0));
    let merged = sheet2.merged_regions();
    assert!(!merged.is_empty(), "Merged cells should be preserved");
    let ps2 = sheet2.page_setup();
    assert!(ps2.print_gridlines, "Print gridlines should be preserved");
    assert!((ps2.left_margin - 1.0).abs() < 1e-9, "Left margin should be preserved");
    assert!(ps2.odd_header.is_some(), "Header should be preserved");
}

/// Test roundtrip of rich text (inline string with per-run formatting)
#[test]
fn test_roundtrip_rich_text() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Cell with mixed formatting: plain + bold + italic+colored
    let runs = vec![
        RichTextRun::plain("Hello "),
        RichTextRun::with_font("bold", RunFont {
            bold: Some(true),
            ..Default::default()
        }),
        RichTextRun::with_font(" world", RunFont {
            italic: Some(true),
            color: Some(Color::rgb(255, 0, 0)),
            ..Default::default()
        }),
    ];
    sheet
        .set_cell_value_at(0, 0, CellValue::RichText(runs.clone()))
        .unwrap();

    // Cell with a single plain run (no formatting)
    let plain_runs = vec![RichTextRun::plain("Just plain text")];
    sheet
        .set_cell_value_at(1, 0, CellValue::RichText(plain_runs.clone()))
        .unwrap();

    // Cell with many font properties
    let fancy_runs = vec![
        RichTextRun::with_font("fancy", RunFont {
            bold: Some(true),
            italic: Some(true),
            size: Some(14.0),
            name: Some("Arial".to_string()),
            underline: Some(Underline::Single),
            strikethrough: Some(true),
            color: Some(Color::rgb(0, 128, 255)),
            ..Default::default()
        }),
    ];
    sheet
        .set_cell_value_at(2, 0, CellValue::RichText(fancy_runs.clone()))
        .unwrap();

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Verify cell A1: mixed formatting
    match sheet2.get_value_at(0, 0) {
        CellValue::RichText(read_runs) => {
            assert_eq!(read_runs.len(), 3, "expected 3 runs");
            assert_eq!(read_runs[0].text, "Hello ");
            assert!(read_runs[0].font.is_none(), "first run should have no font");
            assert_eq!(read_runs[1].text, "bold");
            assert_eq!(read_runs[1].font.as_ref().unwrap().bold, Some(true));
            assert_eq!(read_runs[2].text, " world");
            let f2 = read_runs[2].font.as_ref().unwrap();
            assert_eq!(f2.italic, Some(true));
            // Color roundtrips as Rgb (parse_color_element drops alpha from ARGB)
            assert_eq!(f2.color, Some(Color::rgb(255, 0, 0)));
        }
        other => panic!("expected RichText, got {:?}", other),
    }

    // Verify cell A2: plain run
    match sheet2.get_value_at(1, 0) {
        CellValue::RichText(read_runs) => {
            assert_eq!(read_runs.len(), 1);
            assert_eq!(read_runs[0].text, "Just plain text");
            assert!(read_runs[0].font.is_none());
        }
        other => panic!("expected RichText, got {:?}", other),
    }

    // Verify cell A3: fancy run
    match sheet2.get_value_at(2, 0) {
        CellValue::RichText(read_runs) => {
            assert_eq!(read_runs.len(), 1);
            assert_eq!(read_runs[0].text, "fancy");
            let f = read_runs[0].font.as_ref().unwrap();
            assert_eq!(f.bold, Some(true));
            assert_eq!(f.italic, Some(true));
            assert_eq!(f.size, Some(14.0));
            assert_eq!(f.name.as_deref(), Some("Arial"));
            assert_eq!(f.underline, Some(Underline::Single));
            assert_eq!(f.strikethrough, Some(true));
            assert_eq!(f.color, Some(Color::rgb(0, 128, 255)));
        }
        other => panic!("expected RichText, got {:?}", other),
    }
}
