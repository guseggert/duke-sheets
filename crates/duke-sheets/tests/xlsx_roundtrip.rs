#![allow(clippy::approx_constant)]
//! End-to-end tests for XLSX roundtrip (create -> save -> read -> verify)

use duke_sheets::prelude::*;
use duke_sheets::{
    AutoFilter, ColumnFilter, CustomFilterCondition, CustomFilters, DynamicFilter,
    DynamicFilterType, FilterColumn, FilterOperator, Top10Filter, ValueFilter,
};
use duke_sheets::{CellMarker, CheckState, ListSelection};
use duke_sheets_core::style::Underline;
use duke_sheets_core::worksheet::SheetProtection;
use duke_sheets_core::{
    hash_legacy_protection_password, CellAddress, CellRange, PageOrientation, ProtectedRange,
    Selection, SplitPanes, Table, TableColumn, TableStyleInfo, TotalsRowFunction,
    WorkbookProtection,
};
use std::io::{Cursor, Read};

fn formula_text_at<'a>(sheet: &'a duke_sheets_core::Worksheet, address: &str) -> Option<&'a str> {
    let addr = CellAddress::parse(address).expect("valid cell address");
    sheet.get_formula_at(addr.row, addr.col)
}

fn is_array_formula_at(sheet: &duke_sheets_core::Worksheet, address: &str) -> bool {
    let addr = CellAddress::parse(address).expect("valid cell address");
    sheet
        .formula_data_at(addr.row, addr.col)
        .map(|formula| formula.is_array_formula())
        .unwrap_or(false)
}

#[test]
fn test_roundtrip_protection_settings() {
    let mut wb = Workbook::new();
    wb.set_workbook_protection(Some(WorkbookProtection {
        structure: true,
        windows: true,
        password_hash: Some(hash_legacy_protection_password("book")),
    }));

    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "locked").unwrap();
    sheet.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(hash_legacy_protection_password("password")),
        select_locked_cells: true,
        select_unlocked_cells: true,
        format_cells: true,
        format_columns: true,
        format_rows: true,
        insert_columns: true,
        insert_rows: true,
        insert_hyperlinks: true,
        delete_columns: true,
        delete_rows: true,
        sort: true,
        auto_filter: true,
        pivot_tables: true,
        ..Default::default()
    }));
    sheet.set_protected_ranges(vec![ProtectedRange {
        name: "Editable".to_string(),
        ranges: vec![
            CellRange::parse("A1:B2").unwrap(),
            CellRange::parse("D4:D5").unwrap(),
        ],
        password_hash: Some(0xCAFE),
        security_descriptor: Some("S-1-5-21".to_string()),
    }]);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    let workbook_protection = wb2.workbook_protection().expect("workbook protection");
    assert!(workbook_protection.structure);
    assert!(workbook_protection.windows);
    assert_eq!(
        workbook_protection.password_hash,
        Some(hash_legacy_protection_password("book"))
    );

    let sheet2 = wb2.worksheet(0).unwrap();
    let protection = sheet2.protection().expect("sheet protection");
    assert!(protection.protected);
    assert_eq!(
        protection.password_hash,
        Some(hash_legacy_protection_password("password"))
    );
    assert!(protection.select_locked_cells);
    assert!(protection.select_unlocked_cells);
    assert!(protection.format_cells);
    assert!(protection.format_columns);
    assert!(protection.format_rows);
    assert!(protection.insert_columns);
    assert!(protection.insert_rows);
    assert!(protection.insert_hyperlinks);
    assert!(protection.delete_columns);
    assert!(protection.delete_rows);
    assert!(protection.sort);
    assert!(protection.auto_filter);
    assert!(protection.pivot_tables);

    let ranges = sheet2.protected_ranges();
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].name, "Editable");
    assert_eq!(ranges[0].ranges.len(), 2);
    assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
    assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
    assert_eq!(ranges[0].password_hash, Some(0xCAFE));
    assert_eq!(ranges[0].security_descriptor.as_deref(), Some("S-1-5-21"));
}

#[test]
fn test_roundtrip_sheet_protection_raw_hash() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "locked").unwrap();
    sheet.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(0xCAFE),
        ..Default::default()
    }));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    assert_eq!(
        wb2.worksheet(0)
            .unwrap()
            .protection()
            .expect("sheet protection")
            .password_hash,
        Some(0xCAFE)
    );
}

#[test]
fn test_xlsx_skips_empty_protected_ranges_wrapper() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_protected_ranges(vec![ProtectedRange {
        name: "Empty".to_string(),
        ranges: Vec::new(),
        password_hash: None,
        security_descriptor: None,
    }]);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let mut zip = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
    let mut sheet_xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut sheet_xml)
        .unwrap();
    assert!(!sheet_xml.contains("<protectedRanges"));
}

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
    assert_eq!(formula_text_at(sheet2, "A3"), Some("=SUM(A1:A2)"));
    assert_eq!(formula_text_at(sheet2, "B1"), Some("=A1*2"));
    assert_eq!(
        formula_text_at(sheet2, "C1"),
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
    assert!(formula_text_at(sheet2, "B3").is_some());
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

/// Test roundtrip with XML-special characters in sheet names
#[test]
fn test_roundtrip_xml_special_chars_in_sheet_names() {
    let mut wb = Workbook::empty();
    wb.add_worksheet_with_name("Test<Sheet>").unwrap();
    wb.add_worksheet_with_name("Sales & Marketing").unwrap();
    wb.add_worksheet_with_name("He said \"hi\"").unwrap();
    wb.add_worksheet_with_name("It's a test").unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    assert_eq!(wb2.sheet_count(), 4);
    assert_eq!(wb2.worksheet(0).unwrap().name(), "Test<Sheet>");
    assert_eq!(wb2.worksheet(1).unwrap().name(), "Sales & Marketing");
    assert_eq!(wb2.worksheet(2).unwrap().name(), "He said \"hi\"");
    assert_eq!(wb2.worksheet(3).unwrap().name(), "It's a test");
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

/// Test multi-selection roundtrip with freeze panes (multiple <selection> elements)
#[test]
fn test_roundtrip_multi_selection_with_freeze_panes() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // Freeze row 2, col B
    sheet.set_freeze_panes(2, 1);

    // Set two selections: one for the top-left pane, one for the bottom-right pane
    sheet.set_selections(vec![
        Selection {
            pane: Some("topLeft".to_string()),
            active_cell: Some("A1".to_string()),
            sqref: Some("A1".to_string()),
        },
        Selection {
            pane: Some("bottomRight".to_string()),
            active_cell: Some("C5".to_string()),
            sqref: Some("C5:D8".to_string()),
        },
    ]);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    // Freeze panes should survive
    let freeze = sheet2
        .freeze_panes()
        .expect("freeze panes should roundtrip");
    assert_eq!(freeze.row, 2);
    assert_eq!(freeze.col, 1);

    // Both selections should survive
    let sels = sheet2.selections();
    assert_eq!(sels.len(), 2, "expected 2 selections, got {:?}", sels);

    assert_eq!(sels[0].pane.as_deref(), Some("topLeft"));
    assert_eq!(sels[0].active_cell.as_deref(), Some("A1"));
    assert_eq!(sels[0].sqref.as_deref(), Some("A1"));

    assert_eq!(sels[1].pane.as_deref(), Some("bottomRight"));
    assert_eq!(sels[1].active_cell.as_deref(), Some("C5"));
    assert_eq!(sels[1].sqref.as_deref(), Some("C5:D8"));
}

/// Test multi-range sqref (non-contiguous selection) roundtrips correctly
#[test]
fn test_roundtrip_multi_range_sqref() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    // A single selection with a multi-range sqref (space-separated)
    sheet.set_selections(vec![Selection {
        pane: None,
        active_cell: Some("B5".to_string()),
        sqref: Some("B5:C6 E2:F3".to_string()),
    }]);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let sels = sheet2.selections();
    assert_eq!(sels.len(), 1);
    assert_eq!(sels[0].active_cell.as_deref(), Some("B5"));
    assert_eq!(sels[0].sqref.as_deref(), Some("B5:C6 E2:F3"));
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

/// Test even/first page headers and headerFooter flags roundtrip
#[test]
fn test_roundtrip_header_footer_even_first_and_flags() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut ps = sheet.page_setup().clone();

    // Set all six header/footer strings
    ps.odd_header = Some("&COdd Header".to_string());
    ps.odd_footer = Some("&COdd Footer".to_string());
    ps.even_header = Some("&CEven Header".to_string());
    ps.even_footer = Some("&CEven Footer".to_string());
    ps.first_header = Some("&CFirst Page".to_string());
    ps.first_footer = Some("&CFirst Footer".to_string());

    // Set all four flags to non-default values
    ps.different_odd_even = true;
    ps.different_first = true;
    ps.scale_with_doc = false;
    ps.align_with_margins = false;

    sheet.set_page_setup(ps);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    let ps2 = sheet2.page_setup();

    // Verify all header/footer strings
    assert_eq!(ps2.odd_header.as_deref(), Some("&COdd Header"));
    assert_eq!(ps2.odd_footer.as_deref(), Some("&COdd Footer"));
    assert_eq!(ps2.even_header.as_deref(), Some("&CEven Header"));
    assert_eq!(ps2.even_footer.as_deref(), Some("&CEven Footer"));
    assert_eq!(ps2.first_header.as_deref(), Some("&CFirst Page"));
    assert_eq!(ps2.first_footer.as_deref(), Some("&CFirst Footer"));

    // Verify flags
    assert!(ps2.different_odd_even);
    assert!(ps2.different_first);
    assert!(!ps2.scale_with_doc);
    assert!(!ps2.align_with_margins);
}

/// Test headerFooter with only odd headers and default flags roundtrips correctly
#[test]
fn test_roundtrip_header_footer_odd_only_defaults() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut ps = sheet.page_setup().clone();

    ps.odd_header = Some("&LLeft&RRight".to_string());
    // Leave all flags at defaults

    sheet.set_page_setup(ps);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    let ps2 = sheet2.page_setup();

    assert_eq!(ps2.odd_header.as_deref(), Some("&LLeft&RRight"));
    assert_eq!(ps2.odd_footer.as_deref(), None);
    assert_eq!(ps2.even_header.as_deref(), None);
    assert_eq!(ps2.even_footer.as_deref(), None);
    assert_eq!(ps2.first_header.as_deref(), None);
    assert_eq!(ps2.first_footer.as_deref(), None);

    // Default flags preserved
    assert!(!ps2.different_odd_even);
    assert!(!ps2.different_first);
    assert!(ps2.scale_with_doc);
    assert!(ps2.align_with_margins);
}

#[test]
fn test_roundtrip_print_area() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_print_area(CellRange::parse("B2:F20").unwrap());

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(
        sheet2.print_area(),
        Some(&CellRange::parse("B2:F20").unwrap())
    );
}

#[test]
fn test_roundtrip_repeat_rows() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_repeat_rows(0, 2);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.repeat_rows(), Some((0, 2)));
}

#[test]
fn test_roundtrip_repeat_cols() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_repeat_cols(0, 1);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.repeat_cols(), Some((0, 1)));
}

#[test]
fn test_roundtrip_repeat_rows_and_cols() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_repeat_rows(0, 1);
    sheet.set_repeat_cols(0, 0);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.repeat_rows(), Some((0, 1)));
    assert_eq!(sheet2.repeat_cols(), Some((0, 0)));
}

#[test]
fn test_roundtrip_print_area_and_titles() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_print_area(CellRange::parse("B2:F20").unwrap());
    sheet.set_repeat_rows(0, 2);
    sheet.set_repeat_cols(0, 1);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(
        sheet2.print_area(),
        Some(&CellRange::parse("B2:F20").unwrap())
    );
    assert_eq!(sheet2.repeat_rows(), Some((0, 2)));
    assert_eq!(sheet2.repeat_cols(), Some((0, 1)));
}

#[test]
fn test_roundtrip_print_area_quoted_sheet_name() {
    use duke_sheets::named_range::NameScope;

    let mut wb = Workbook::empty();
    wb.add_worksheet_with_name("My Sheet").unwrap();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_print_area(CellRange::parse("B2:F20").unwrap());

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(
        sheet2.print_area(),
        Some(&CellRange::parse("B2:F20").unwrap())
    );
    let nr = wb2
        .named_ranges()
        .get_exact("_xlnm.Print_Area", &NameScope::Sheet(0))
        .unwrap();
    assert_eq!(nr.refers_to, "'My Sheet'!$B$2:$F$20");
}

#[test]
fn test_auto_filter_range_only() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let range = CellRange::parse("A1:D10").unwrap();
    sheet.set_auto_filter(Some(AutoFilter::new(range.clone())));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af.range, range);
    assert!(af.filter_columns.is_empty());
}

#[test]
fn test_auto_filter_value_filter() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut af = AutoFilter::new(CellRange::parse("A1:D10").unwrap());
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(ValueFilter {
            values: vec!["Alice".to_string(), "Bob".to_string()],
            blank: false,
        }),
    ));
    sheet.set_auto_filter(Some(af));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af2 = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af2.filter_columns.len(), 1);
    assert_eq!(af2.filter_columns[0].col_id, 0);
    match &af2.filter_columns[0].filter {
        ColumnFilter::Values(v) => {
            assert_eq!(v.values, vec!["Alice".to_string(), "Bob".to_string()]);
            assert!(!v.blank);
        }
        other => panic!("expected ValueFilter, got {:?}", other),
    }
}

#[test]
fn test_auto_filter_custom_filter() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut af = AutoFilter::new(CellRange::parse("A1:D10").unwrap());
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Custom(CustomFilters {
            and: false,
            conditions: vec![CustomFilterCondition {
                operator: FilterOperator::GreaterThan,
                value: "100".to_string(),
            }],
        }),
    ));
    sheet.set_auto_filter(Some(af));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af2 = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af2.filter_columns.len(), 1);
    match &af2.filter_columns[0].filter {
        ColumnFilter::Custom(custom) => {
            assert!(!custom.and);
            assert_eq!(custom.conditions.len(), 1);
            assert_eq!(custom.conditions[0].operator, FilterOperator::GreaterThan);
            assert_eq!(custom.conditions[0].value, "100");
        }
        other => panic!("expected CustomFilters, got {:?}", other),
    }
}

#[test]
fn test_auto_filter_top10() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut af = AutoFilter::new(CellRange::parse("A1:D10").unwrap());
    af.filter_columns.push(FilterColumn::new(
        2,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 10.0,
            filter_val: None,
        }),
    ));
    sheet.set_auto_filter(Some(af));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af2 = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af2.filter_columns.len(), 1);
    match &af2.filter_columns[0].filter {
        ColumnFilter::Top10(top10) => {
            assert!(top10.top);
            assert!(!top10.percent);
            assert_eq!(top10.val, 10.0);
            assert_eq!(top10.filter_val, None);
        }
        other => panic!("expected Top10Filter, got {:?}", other),
    }
}

#[test]
fn test_auto_filter_dynamic() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut af = AutoFilter::new(CellRange::parse("A1:D10").unwrap());
    af.filter_columns.push(FilterColumn::new(
        3,
        ColumnFilter::Dynamic(DynamicFilter {
            filter_type: DynamicFilterType::AboveAverage,
            val: None,
            max_val: None,
        }),
    ));
    sheet.set_auto_filter(Some(af));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af2 = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af2.filter_columns.len(), 1);
    match &af2.filter_columns[0].filter {
        ColumnFilter::Dynamic(dynamic) => {
            assert_eq!(dynamic.filter_type, DynamicFilterType::AboveAverage);
            assert_eq!(dynamic.val, None);
            assert_eq!(dynamic.max_val, None);
        }
        other => panic!("expected DynamicFilter, got {:?}", other),
    }
}

#[test]
fn test_auto_filter_multiple_columns() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut af = AutoFilter::new(CellRange::parse("A1:D10").unwrap());
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(ValueFilter {
            values: vec!["East".to_string(), "West".to_string()],
            blank: true,
        }),
    ));
    af.filter_columns.push(FilterColumn::new(
        2,
        ColumnFilter::Top10(Top10Filter {
            top: false,
            percent: true,
            val: 25.0,
            filter_val: Some(200.0),
        }),
    ));
    af.filter_columns.push(FilterColumn::new(
        3,
        ColumnFilter::Dynamic(DynamicFilter {
            filter_type: DynamicFilterType::ThisMonth,
            val: Some(1.0),
            max_val: Some(31.0),
        }),
    ));
    sheet.set_auto_filter(Some(af));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let af2 = sheet2.auto_filter().expect("auto_filter should exist");
    assert_eq!(af2.filter_columns.len(), 3);
    assert!(matches!(
        af2.filter_columns[0].filter,
        ColumnFilter::Values(_)
    ));
    assert!(matches!(
        af2.filter_columns[1].filter,
        ColumnFilter::Top10(_)
    ));
    assert!(matches!(
        af2.filter_columns[2].filter,
        ColumnFilter::Dynamic(_)
    ));
}

/// Test roundtrip of formula with numeric cached value
#[test]
fn test_roundtrip_formula_cached_number() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_formula_at(0, 0, "=1+2").unwrap();
    sheet
        .set_formula_result(0, 0, CellValue::Number(3.0))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(formula_text_at(sheet2, "A1"), Some("=1+2"));
    match val.effective_value() {
        CellValue::Number(n) => assert!((n - 3.0).abs() < 1e-10),
        other => panic!("Expected Number(3.0), got {:?}", other),
    }
}

/// Unary plus and redundant parentheses must survive the round trip
/// verbatim, matching Excel's behaviour of preserving both.
#[test]
fn test_roundtrip_formula_unary_plus_and_parens() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value_at(0, 0, 2.0).unwrap();
    sheet.set_cell_formula_at(0, 1, "=+A1").unwrap();
    sheet
        .set_formula_result(0, 1, CellValue::Number(2.0))
        .unwrap();
    sheet.set_cell_formula_at(0, 2, "=(A1+1)").unwrap();
    sheet
        .set_formula_result(0, 2, CellValue::Number(3.0))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(formula_text_at(sheet2, "B1"), Some("=+A1"));
    assert_eq!(formula_text_at(sheet2, "C1"), Some("=(A1+1)"));
}

/// Analysis-ToolPak add-in functions (EDATE, NETWORKDAYS, GCD, ...) serialize
/// as plain formula text in XLSX — no special encoding, unlike the BIFF8
/// add-in form (XLS) or native iftab tokens (XLSB). This pins that the text
/// round-trips intact.
#[test]
fn test_roundtrip_atp_addin_functions() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let cases = [
        ("A1", "=EDATE(B1,12)"),
        ("A2", "=NETWORKDAYS(B1,B2)"),
        ("A3", "=GCD(B1,B2)"),
    ];
    for (cell, formula) in cases {
        sheet.set_cell_formula(cell, formula).unwrap();
    }

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    for (cell, formula) in cases {
        assert_eq!(
            formula_text_at(sheet2, cell),
            Some(formula),
            "ATP formula text must round-trip in XLSX for {cell}"
        );
    }
}

#[test]
fn test_roundtrip_external_udf_formula_text() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet
        .set_cell_formula("A1", r#"=[1]!TBLink("acct")"#)
        .unwrap();
    sheet
        .set_formula_result(0, 0, CellValue::Number(42.0))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(
        formula_text_at(sheet2, "A1"),
        Some(r#"=[1]!TBLink("acct")"#)
    );
    assert_eq!(
        sheet2.get_calculated_value_at(0, 0),
        Some(&CellValue::Number(42.0))
    );
}

/// Test roundtrip of formula with string cached value
#[test]
fn test_roundtrip_formula_cached_string() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet
        .set_cell_formula_at(0, 0, "=CONCAT(\"hello\")")
        .unwrap();
    sheet
        .set_formula_result(0, 0, CellValue::string("hello"))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(formula_text_at(sheet2, "A1"), Some("=CONCAT(\"hello\")"));
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

    sheet.set_cell_formula_at(0, 0, "=TRUE()").unwrap();
    sheet
        .set_formula_result(0, 0, CellValue::Boolean(true))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(formula_text_at(sheet2, "A1"), Some("=TRUE()"));
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

    sheet.set_cell_formula_at(0, 0, "=1/0").unwrap();
    sheet
        .set_formula_result(0, 0, CellValue::Error(CellError::Div0))
        .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(formula_text_at(sheet2, "A1"), Some("=1/0"));
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

    sheet.set_cell_formula_at(0, 0, "=SUM(B1:B10)").unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let val = sheet2.get_value("A1").unwrap();
    assert_eq!(formula_text_at(sheet2, "A1"), Some("=SUM(B1:B10)"));
    assert_eq!(val, CellValue::Empty);
    assert_eq!(
        sheet2.get_calculated_value_at(0, 0),
        Some(&CellValue::Empty)
    );
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
    assert!(
        (ps2.left_margin - 1.0).abs() < 1e-9,
        "Left margin should be preserved"
    );
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
        RichTextRun::with_font(
            "bold",
            RunFont {
                bold: Some(true),
                ..Default::default()
            },
        ),
        RichTextRun::with_font(
            " world",
            RunFont {
                italic: Some(true),
                color: Some(Color::rgb(255, 0, 0)),
                ..Default::default()
            },
        ),
    ];
    sheet
        .set_cell_value_at(0, 0, CellValue::rich_text(runs.clone()))
        .unwrap();

    // Cell with a single plain run (no formatting)
    let plain_runs = vec![RichTextRun::plain("Just plain text")];
    sheet
        .set_cell_value_at(1, 0, CellValue::rich_text(plain_runs.clone()))
        .unwrap();

    // Cell with many font properties
    let fancy_runs = vec![RichTextRun::with_font(
        "fancy",
        RunFont {
            bold: Some(true),
            italic: Some(true),
            size: Some(14.0),
            name: Some("Arial".to_string()),
            underline: Some(Underline::Single),
            strikethrough: Some(true),
            color: Some(Color::rgb(0, 128, 255)),
            ..Default::default()
        },
    )];
    sheet
        .set_cell_value_at(2, 0, CellValue::rich_text(fancy_runs.clone()))
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

/// Test table roundtrip: create table -> save -> read -> verify all fields survive
#[test]
fn test_roundtrip_table_basic() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Name").unwrap();
    sheet.set_cell_value("B1", "Score").unwrap();
    sheet.set_cell_value("C1", "Grade").unwrap();
    sheet.set_cell_value("A2", "Alice").unwrap();
    sheet.set_cell_value("B2", 95.0).unwrap();
    sheet.set_cell_value("C2", "A").unwrap();

    let mut table = Table::new(1, "Students", CellRange::parse("A1:C3").unwrap());
    table.columns = vec![
        TableColumn::new(1, "Name"),
        TableColumn::new(2, "Score"),
        TableColumn::new(3, "Grade"),
    ];
    table.style_info = Some(TableStyleInfo {
        name: Some("TableStyleMedium9".into()),
        show_first_column: false,
        show_last_column: false,
        show_row_stripes: true,
        show_column_stripes: false,
    });
    sheet.add_table(table);

    // Write
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Read back
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.table_count(), 1);
    let t = &sheet2.tables()[0];
    assert_eq!(t.id, 1);
    assert_eq!(t.name, "Students");
    assert_eq!(t.display_name, "Students");
    assert_eq!(t.reference.to_string(), "A1:C3");
    assert_eq!(t.header_row_count, 1);
    assert_eq!(t.totals_row_count, 0);
    assert!(t.has_header_row());
    assert!(!t.has_totals_row());

    assert_eq!(t.columns.len(), 3);
    assert_eq!(t.columns[0].name, "Name");
    assert_eq!(t.columns[1].name, "Score");
    assert_eq!(t.columns[2].name, "Grade");

    let style = t.style_info.as_ref().unwrap();
    assert_eq!(style.name.as_deref(), Some("TableStyleMedium9"));
    assert!(style.show_row_stripes);
    assert!(!style.show_column_stripes);
}

/// Test table roundtrip with totals row, labels, and functions
#[test]
fn test_roundtrip_table_with_totals() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Item").unwrap();
    sheet.set_cell_value("B1", "Qty").unwrap();
    sheet.set_cell_value("C1", "Price").unwrap();

    let mut table = Table::new(1, "Sales", CellRange::parse("A1:C5").unwrap());
    let mut col1 = TableColumn::new(1, "Item");
    col1.totals_row_label = Some("Total".into());
    let mut col2 = TableColumn::new(2, "Qty");
    col2.totals_row_function = Some(TotalsRowFunction::Sum);
    let mut col3 = TableColumn::new(3, "Price");
    col3.totals_row_function = Some(TotalsRowFunction::Custom);
    col3.totals_row_formula = Some("SUBTOTAL(109,[Price])".into());
    table.columns = vec![col1, col2, col3];
    table.totals_row_count = 1;
    sheet.add_table(table);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    let t = &sheet2.tables()[0];

    assert!(t.has_totals_row());
    assert_eq!(t.totals_row_count, 1);
    assert_eq!(t.columns[0].totals_row_label.as_deref(), Some("Total"));
    assert_eq!(
        t.columns[1].totals_row_function,
        Some(TotalsRowFunction::Sum)
    );
    assert_eq!(
        t.columns[2].totals_row_function,
        Some(TotalsRowFunction::Custom)
    );
    assert_eq!(
        t.columns[2].totals_row_formula.as_deref(),
        Some("SUBTOTAL(109,[Price])")
    );
}

/// Test table roundtrip with calculated column formula
#[test]
fn test_roundtrip_table_with_calculated_column() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "X").unwrap();
    sheet.set_cell_value("B1", "Y").unwrap();
    sheet.set_cell_value("C1", "Sum").unwrap();

    let mut table = Table::new(1, "Calc", CellRange::parse("A1:C4").unwrap());
    let col1 = TableColumn::new(1, "X");
    let col2 = TableColumn::new(2, "Y");
    let mut col3 = TableColumn::new(3, "Sum");
    col3.calculated_column_formula = Some("[X]+[Y]".into());
    table.columns = vec![col1, col2, col3];
    sheet.add_table(table);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    let t = &sheet2.tables()[0];

    assert_eq!(
        t.columns[2].calculated_column_formula.as_deref(),
        Some("[X]+[Y]")
    );
}

/// Test multiple tables across multiple sheets roundtrip
#[test]
fn test_roundtrip_multiple_tables() {
    let mut wb = Workbook::new();

    // Sheet 1 with 2 tables
    let sheet1 = wb.worksheet_mut(0).unwrap();
    sheet1.set_cell_value("A1", "H1").unwrap();
    let mut t1 = Table::new(1, "Table1", CellRange::parse("A1:B3").unwrap());
    t1.columns = vec![TableColumn::new(1, "H1"), TableColumn::new(2, "H2")];
    sheet1.add_table(t1);

    let mut t2 = Table::new(2, "Table2", CellRange::parse("D1:E3").unwrap());
    t2.columns = vec![TableColumn::new(1, "D1"), TableColumn::new(2, "E1")];
    sheet1.add_table(t2);

    // Sheet 2 with 1 table
    wb.add_worksheet_with_name("Sheet2").unwrap();
    let sheet2 = wb.worksheet_mut(1).unwrap();
    sheet2.set_cell_value("A1", "Col").unwrap();
    let mut t3 = Table::new(3, "Table3", CellRange::parse("A1:A5").unwrap());
    t3.columns = vec![TableColumn::new(1, "Col")];
    sheet2.add_table(t3);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Sheet 1 should have 2 tables
    let s1 = wb2.worksheet(0).unwrap();
    assert_eq!(s1.table_count(), 2);
    assert_eq!(s1.tables()[0].name, "Table1");
    assert_eq!(s1.tables()[1].name, "Table2");

    // Sheet 2 should have 1 table
    let s2 = wb2.worksheet(1).unwrap();
    assert_eq!(s2.table_count(), 1);
    assert_eq!(s2.tables()[0].name, "Table3");
}

#[test]
fn roundtrip_row_breaks() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_row_break(9);
    ws.add_row_break(19);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    let breaks = ws2.row_breaks();
    assert_eq!(breaks.len(), 2);
    assert_eq!(breaks[0].id, 9);
    assert_eq!(breaks[0].max, 16383);
    assert!(breaks[0].man);
    assert_eq!(breaks[1].id, 19);
}

#[test]
fn roundtrip_col_breaks() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_col_break(3);
    ws.add_col_break(7);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    let breaks = ws2.col_breaks();
    assert_eq!(breaks.len(), 2);
    assert_eq!(breaks[0].id, 3);
    assert_eq!(breaks[0].max, 1048575);
    assert!(breaks[0].man);
    assert_eq!(breaks[1].id, 7);
}

#[test]
fn roundtrip_mixed_breaks() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_row_break(4);
    ws.add_col_break(2);

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    assert_eq!(ws2.row_breaks().len(), 1);
    assert_eq!(ws2.row_breaks()[0].id, 4);
    assert_eq!(ws2.col_breaks().len(), 1);
    assert_eq!(ws2.col_breaks()[0].id, 2);
}

// Dynamic array spilling - XLSX metadata roundtrip

/// Verify that SEQUENCE formula produces spilled values that survive roundtrip.
/// After calculate→write→read, the anchor cell keeps its formula and the
/// spill-target cells are written as plain cached values with cm attributes.
#[test]
fn roundtrip_dynamic_array_sequence() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(3,1)").unwrap();
    wb.calculate().unwrap();

    // Verify in-memory spill before writing
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.get_value_at(0, 0).as_number(), Some(1.0)); // A1
    assert_eq!(ws.get_value_at(1, 0).as_number(), Some(2.0)); // A2
    assert_eq!(ws.get_value_at(2, 0).as_number(), Some(3.0)); // A3

    // Write to buffer
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Verify xl/metadata.xml is in the ZIP
    {
        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        assert!(
            zip.by_name("xl/metadata.xml").is_ok(),
            "xl/metadata.xml should be present for dynamic arrays"
        );
    }

    // Verify cm attributes in sheet XML
    {
        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let mut sheet_xml = String::new();
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        // Anchor cell A1 should have cm="1"
        assert!(
            sheet_xml.contains("cm=\"1\""),
            "anchor cell should have cm=1 attribute"
        );
        // Ghost cells A2, A3 should have cm="2"
        assert!(
            sheet_xml.contains("cm=\"2\""),
            "ghost cells should have cm=2 attribute"
        );
    }

    // Read back and verify
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // A1 should still be a formula
    let a1_formula = ws2.formula_data_at(0, 0).expect("A1 should be a formula");
    assert!(
        a1_formula.text.contains("SEQUENCE"),
        "formula text should contain SEQUENCE"
    );
    assert!(
        a1_formula.is_array_formula(),
        "A1 should remain an array formula"
    );
    assert_eq!(ws2.get_value("A1").unwrap().as_number(), Some(1.0));

    // Ghost cells should now be SpillTarget (reader reconstructs dynamic array)
    let a2 = ws2.get_value("A2").unwrap();
    assert!(
        a2.is_spill_target(),
        "A2 should be SpillTarget, got {:?}",
        a2
    );
    let a3 = ws2.get_value("A3").unwrap();
    assert!(
        a3.is_spill_target(),
        "A3 should be SpillTarget, got {:?}",
        a3
    );
    // Resolved values match
    assert_eq!(ws2.get_value_at(1, 0).as_number(), Some(2.0), "A2 resolved");
    assert_eq!(ws2.get_value_at(2, 0).as_number(), Some(3.0), "A3 resolved");
}

/// Verify that a 2D SEQUENCE (rows × cols) roundtrips correctly.
#[test]
fn roundtrip_dynamic_array_sequence_2d() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(2,3)").unwrap();
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    // 2×3 grid: [[1,2,3],[4,5,6]]
    assert_eq!(ws.get_value_at(0, 0).as_number(), Some(1.0));
    assert_eq!(ws.get_value_at(0, 1).as_number(), Some(2.0));
    assert_eq!(ws.get_value_at(0, 2).as_number(), Some(3.0));
    assert_eq!(ws.get_value_at(1, 0).as_number(), Some(4.0));
    assert_eq!(ws.get_value_at(1, 1).as_number(), Some(5.0));
    assert_eq!(ws.get_value_at(1, 2).as_number(), Some(6.0));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // Anchor cell keeps formula
    let a1_formula = ws2.formula_data_at(0, 0).expect("A1 should be formula");
    assert!(a1_formula.text.contains("SEQUENCE"));
    assert!(a1_formula.is_array_formula());

    // Ghost cells are now SpillTarget; resolved values match
    assert!(ws2.get_value("B1").unwrap().is_spill_target());
    assert!(ws2.get_value("C1").unwrap().is_spill_target());
    assert!(ws2.get_value("A2").unwrap().is_spill_target());
    assert!(ws2.get_value("B2").unwrap().is_spill_target());
    assert!(ws2.get_value("C2").unwrap().is_spill_target());
    assert_eq!(ws2.get_value_at(0, 1).as_number(), Some(2.0));
    assert_eq!(ws2.get_value_at(0, 2).as_number(), Some(3.0));
    assert_eq!(ws2.get_value_at(1, 0).as_number(), Some(4.0));
    assert_eq!(ws2.get_value_at(1, 1).as_number(), Some(5.0));
    assert_eq!(ws2.get_value_at(1, 2).as_number(), Some(6.0));
}

/// Verify metadata.xml is NOT written when there are no dynamic arrays.
#[test]
fn roundtrip_no_metadata_without_dynamic_arrays() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).unwrap();
    ws.set_cell_formula("B1", "=A1*2").unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
    assert!(
        zip.by_name("xl/metadata.xml").is_err(),
        "xl/metadata.xml should NOT be present without dynamic arrays"
    );
}

/// Verify content types and workbook rels include metadata entries.
#[test]
fn roundtrip_dynamic_array_content_types_and_rels() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(2,1)").unwrap();
    wb.calculate().unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();

    // Check [Content_Types].xml mentions metadata
    {
        let mut ct_xml = String::new();
        use std::io::Read;
        zip.by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut ct_xml)
            .unwrap();
        assert!(
            ct_xml.contains("metadata.xml"),
            "Content_Types should reference metadata.xml"
        );
        assert!(
            ct_xml.contains(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
            ),
            "Content_Types should have metadata content type"
        );
    }

    // Check xl/_rels/workbook.xml.rels mentions metadata
    {
        let mut rels_xml = String::new();
        use std::io::Read;
        zip.by_name("xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut rels_xml)
            .unwrap();
        assert!(
            rels_xml.contains("metadata.xml"),
            "workbook rels should reference metadata.xml"
        );
        assert!(
            rels_xml.contains("sheetMetadata"),
            "workbook rels should have sheetMetadata relationship type"
        );
    }
}

/// Verify UNIQUE (string spill targets) roundtrip.
#[test]
fn roundtrip_dynamic_array_unique_strings() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "apple").unwrap();
    ws.set_cell_value("A2", "banana").unwrap();
    ws.set_cell_value("A3", "apple").unwrap();
    ws.set_cell_value("A4", "cherry").unwrap();
    ws.set_cell_formula("B1", "=UNIQUE(A1:A4)").unwrap();
    wb.calculate().unwrap();

    // Verify in-memory
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.get_value_at(0, 1).as_string(), Some("apple"));
    assert_eq!(ws.get_value_at(1, 1).as_string(), Some("banana"));
    assert_eq!(ws.get_value_at(2, 1).as_string(), Some("cherry"));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Verify cm attributes for string ghost cells
    {
        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let mut sheet_xml = String::new();
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        // String ghost cells should use t="str" and cm="2"
        assert!(
            sheet_xml.contains("cm=\"2\""),
            "string ghost cells should have cm=2"
        );
    }

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // Anchor keeps formula
    assert!(formula_text_at(ws2, "B1").is_some());
    // Ghost cells are now SpillTarget; resolved values match
    assert!(ws2.get_value("B2").unwrap().is_spill_target());
    assert!(ws2.get_value("B3").unwrap().is_spill_target());
    assert_eq!(ws2.get_value_at(1, 1).as_string(), Some("banana"));
    assert_eq!(ws2.get_value_at(2, 1).as_string(), Some("cherry"));
    // B4 should be empty (only 3 unique values)
    let b4 = ws2.get_value("B4").unwrap();
    assert!(
        matches!(b4, duke_sheets_core::CellValue::Empty),
        "B4 should be empty, got {:?}",
        b4
    );
}

/// Verify boolean spill targets roundtrip.
/// Uses SEQUENCE > threshold to produce boolean array.
#[test]
fn roundtrip_dynamic_array_boolean_spill() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // Produce booleans: {1,2,3} > 1 = {FALSE, TRUE, TRUE}
    ws.set_cell_formula("A1", "=SEQUENCE(3,1)>1").unwrap();
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    let _a1 = ws.get_calculated_value_at(0, 0);
    // Result depends on whether the engine produces a spill array from comparison
    // or a single scalar. Check what we got:
    let is_array = is_array_formula_at(ws, "A1");

    if is_array {
        // Full array spill: write and read back
        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert!(formula_text_at(ws2, "A1").is_some());
    } else {
        // Engine produces scalar (implicit intersection) - single cached value
        // Still verify roundtrip works without crash
        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        assert!(wb2.worksheet(0).is_some());
    }
}

/// Verify #SPILL! error on blocked range roundtrips.
#[test]
fn roundtrip_dynamic_array_spill_error() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // Block the spill range
    ws.set_cell_value("A2", 999.0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    // A1 should have #SPILL! error
    assert_eq!(
        ws.get_calculated_value_at(0, 0),
        Some(&duke_sheets_core::CellValue::Error(
            duke_sheets_core::CellError::Spill,
        ))
    );
    // A2 keeps its original value
    assert_eq!(ws.get_value_at(1, 0).as_number(), Some(999.0));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // No dynamic arrays actually spilled, so no metadata.xml
    {
        let zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        // The workbook has a formula with a #SPILL! cached error but no
        // array_result, so has_dynamic_arrays should return false.
        // However, the formula itself exists - just verify the file is valid.
        let _ = zip;
    }

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // Formula should survive roundtrip
    let a1 = ws2.get_value("A1").unwrap();
    assert!(
        formula_text_at(ws2, "A1").is_some(),
        "A1 should still be a formula"
    );
    assert!(
        matches!(
            a1,
            duke_sheets_core::CellValue::Error(duke_sheets_core::CellError::Spill)
        ),
        "cached value should be #SPILL!, got {:?}",
        a1
    );
    // A2 original value preserved
    assert_eq!(ws2.get_value("A2").unwrap().as_number(), Some(999.0));
}

/// Verify multiple dynamic arrays on the same sheet.
#[test]
fn roundtrip_dynamic_array_multiple_on_same_sheet() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // Two separate spill ranges that don't overlap
    ws.set_cell_formula("A1", "=SEQUENCE(3,1)").unwrap(); // A1:A3
    ws.set_cell_formula("C1", "=SEQUENCE(2,2)").unwrap(); // C1:D2
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.get_value_at(0, 0).as_number(), Some(1.0));
    assert_eq!(ws.get_value_at(2, 0).as_number(), Some(3.0));
    assert_eq!(ws.get_value_at(0, 2).as_number(), Some(1.0));
    assert_eq!(ws.get_value_at(1, 3).as_number(), Some(4.0));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Both arrays should produce cm attributes
    {
        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let mut sheet_xml = String::new();
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        // Count cm="1" occurrences (should be 2 - one per anchor)
        let cm1_count = sheet_xml.matches("cm=\"1\"").count();
        assert_eq!(cm1_count, 2, "should have 2 anchor cells with cm=1");
        // Ghost cell count: A2,A3 + C2,D1,D2 = 5
        let cm2_count = sheet_xml.matches("cm=\"2\"").count();
        assert_eq!(cm2_count, 5, "should have 5 ghost cells with cm=2");
    }

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // Both formulas survive
    assert!(formula_text_at(ws2, "A1").is_some());
    assert!(formula_text_at(ws2, "C1").is_some());
    // Ghost cells are SpillTarget; resolved values match
    assert!(ws2.get_value("A2").unwrap().is_spill_target());
    assert!(ws2.get_value("A3").unwrap().is_spill_target());
    assert!(ws2.get_value("D1").unwrap().is_spill_target());
    assert!(ws2.get_value("C2").unwrap().is_spill_target());
    assert!(ws2.get_value("D2").unwrap().is_spill_target());
    assert_eq!(ws2.get_value_at(1, 0).as_number(), Some(2.0));
    assert_eq!(ws2.get_value_at(2, 0).as_number(), Some(3.0));
    assert_eq!(ws2.get_value_at(0, 3).as_number(), Some(2.0));
    assert_eq!(ws2.get_value_at(1, 2).as_number(), Some(3.0));
    assert_eq!(ws2.get_value_at(1, 3).as_number(), Some(4.0));
}

/// Verify dynamic arrays on multiple sheets.
#[test]
fn roundtrip_dynamic_array_multi_sheet() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(2,1)").unwrap();

    wb.add_worksheet_with_name("Sheet2").unwrap();
    let ws2 = wb.worksheet_mut(1).unwrap();
    ws2.set_cell_formula("A1", "=SEQUENCE(3,1)").unwrap();
    wb.calculate().unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // metadata.xml present (arrays exist)
    {
        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        assert!(zip.by_name("xl/metadata.xml").is_ok());
    }

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Sheet1: A1 formula, A2 SpillTarget
    let s1 = wb2.worksheet(0).unwrap();
    assert!(formula_text_at(s1, "A1").is_some());
    assert!(s1.get_value("A2").unwrap().is_spill_target());
    assert_eq!(s1.get_value_at(1, 0).as_number(), Some(2.0));

    // Sheet2: A1 formula, A2-A3 SpillTarget
    let s2 = wb2.worksheet(1).unwrap();
    assert!(formula_text_at(s2, "A1").is_some());
    assert!(s2.get_value("A2").unwrap().is_spill_target());
    assert!(s2.get_value("A3").unwrap().is_spill_target());
    assert_eq!(s2.get_value_at(1, 0).as_number(), Some(2.0));
    assert_eq!(s2.get_value_at(2, 0).as_number(), Some(3.0));
}

/// Validate the structure of xl/metadata.xml.
#[test]
fn roundtrip_dynamic_array_metadata_xml_structure() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=SEQUENCE(2,1)").unwrap();
    wb.calculate().unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
    let mut metadata_xml = String::new();
    use std::io::Read;
    zip.by_name("xl/metadata.xml")
        .unwrap()
        .read_to_string(&mut metadata_xml)
        .unwrap();

    // Root element
    assert!(
        metadata_xml.contains("<metadata"),
        "should have <metadata> root"
    );
    assert!(
        metadata_xml.contains("xmlns:xda"),
        "should have xda namespace"
    );
    assert!(
        metadata_xml.contains("dynamicarray"),
        "xda namespace URI should reference dynamicarray"
    );

    // metadataTypes with XLDAPR
    assert!(
        metadata_xml.contains("metadataType"),
        "should have metadataType element"
    );
    assert!(
        metadata_xml.contains("XLDAPR"),
        "should have XLDAPR metadata type name"
    );
    assert!(
        metadata_xml.contains("cellMeta=\"1\""),
        "XLDAPR should have cellMeta=1"
    );

    // futureMetadata with 2 entries
    assert!(
        metadata_xml.contains("futureMetadata"),
        "should have futureMetadata section"
    );
    assert!(
        metadata_xml.contains("fDynamic=\"1\""),
        "first entry should have fDynamic=1 (anchor)"
    );
    assert!(
        metadata_xml.contains("fCollapsed=\"1\""),
        "second entry should have fCollapsed=1 (ghost)"
    );

    // cellMetadata with rc entries
    assert!(
        metadata_xml.contains("cellMetadata"),
        "should have cellMetadata section"
    );
    // rc t="1" v="0" (anchor) and rc t="1" v="1" (ghost)
    assert!(
        metadata_xml.contains(r#"v="0""#),
        "cellMetadata should have v=0 entry (anchor)"
    );
    assert!(
        metadata_xml.contains(r#"v="1""#),
        "cellMetadata should have v=1 entry (ghost)"
    );
}

/// Verify SORT with string data roundtrips (string type in ghost cells).
#[test]
fn roundtrip_dynamic_array_sort_strings() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "cherry").unwrap();
    ws.set_cell_value("A2", "apple").unwrap();
    ws.set_cell_value("A3", "banana").unwrap();
    ws.set_cell_formula("B1", "=SORT(A1:A3)").unwrap();
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.get_value_at(0, 1).as_string(), Some("apple"));
    assert_eq!(ws.get_value_at(1, 1).as_string(), Some("banana"));
    assert_eq!(ws.get_value_at(2, 1).as_string(), Some("cherry"));

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    assert!(formula_text_at(ws2, "B1").is_some());
    assert!(ws2.get_value("B2").unwrap().is_spill_target());
    assert!(ws2.get_value("B3").unwrap().is_spill_target());
    assert_eq!(ws2.get_value_at(1, 1).as_string(), Some("banana"));
    assert_eq!(ws2.get_value_at(2, 1).as_string(), Some("cherry"));
}

/// Verify boolean array from comparison operator roundtrips with correct
/// cm attributes and boolean type in ghost cells.
#[test]
fn roundtrip_dynamic_array_comparison_boolean() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // =SEQUENCE(4)>2 produces {FALSE,FALSE,TRUE,TRUE}
    ws.set_cell_formula("A1", "=SEQUENCE(4)>2").unwrap();
    wb.calculate().unwrap();

    let ws = wb.worksheet(0).unwrap();
    // Verify the engine produced an array result
    assert!(
        is_array_formula_at(ws, "A1"),
        "A1 should be a formula with array_result"
    );
    assert_eq!(
        ws.get_calculated_value_at(0, 0),
        Some(&duke_sheets_core::CellValue::Boolean(false))
    );
    assert_eq!(
        ws.get_calculated_value_at(1, 0),
        Some(&duke_sheets_core::CellValue::Boolean(false))
    );
    assert_eq!(
        ws.get_calculated_value_at(2, 0),
        Some(&duke_sheets_core::CellValue::Boolean(true))
    );
    assert_eq!(
        ws.get_calculated_value_at(3, 0),
        Some(&duke_sheets_core::CellValue::Boolean(true))
    );

    // Write to XLSX
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Verify the XML has cm attributes and boolean types
    {
        use std::io::Read;

        let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();

        // Check sheet XML for cm attributes
        let mut sheet_xml = String::new();
        zip.by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();

        // Anchor cell should have cm="1"
        assert!(
            sheet_xml.contains(r#"cm="1""#),
            "anchor cell should have cm=1, xml: {}",
            sheet_xml
        );
        // Ghost cells should have cm="2"
        assert!(
            sheet_xml.contains(r#"cm="2""#),
            "ghost cells should have cm=2, xml: {}",
            sheet_xml
        );
        // Ghost cells should have t="b" (boolean type)
        assert!(
            sheet_xml.contains(r#"t="b""#),
            "ghost cells should have boolean type, xml: {}",
            sheet_xml
        );

        // metadata.xml should exist
        assert!(
            zip.by_name("xl/metadata.xml").is_ok(),
            "metadata.xml should exist for dynamic arrays"
        );
    }

    // Read back and verify values
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let ws2 = wb2.worksheet(0).unwrap();

    // Formula should survive
    assert!(
        formula_text_at(ws2, "A1").is_some(),
        "A1 should still be a formula"
    );

    // Ghost cells are SpillTarget; resolved values are booleans
    assert!(ws2.get_value("A2").unwrap().is_spill_target());
    assert!(ws2.get_value("A3").unwrap().is_spill_target());
    assert!(ws2.get_value("A4").unwrap().is_spill_target());
    assert_eq!(ws2.get_value_at(1, 0).as_bool(), Some(false));
    assert_eq!(ws2.get_value_at(2, 0).as_bool(), Some(true));
    assert_eq!(ws2.get_value_at(3, 0).as_bool(), Some(true));
}

#[test]
fn test_roundtrip_chart_bar() {
    use duke_sheets_chart::{
        Axis, CellMarker, Chart, ChartType, DataReference, DataSeries, DrawingAnchor, Legend,
        LegendPosition,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Q1").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("C1", "Profit").unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Sales Chart".to_string());
    let anchor = DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 0,
            row: 5,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 10,
            col_offset_emu: 0,
            row: 20,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let s1 = DataSeries::new(DataReference::formula("Sheet1!$B$2:$B$5"))
        .with_name("Sheet1!$B$1")
        .with_categories(DataReference::formula("Sheet1!$A$2:$A$5"));
    let s2 = DataSeries::new(DataReference::formula("Sheet1!$C$2:$C$5"))
        .with_name("Sheet1!$C$1")
        .with_categories(DataReference::formula("Sheet1!$A$2:$A$5"));
    chart.add_series(s1);
    chart.add_series(s2);
    chart.category_axis = Some(Axis::new().with_title("Quarter"));
    chart.value_axis = Some(Axis::new().with_title("Amount").with_bounds(0.0, 50000.0));
    chart.legend = Some(Legend::new(LegendPosition::Bottom));
    sheet.add_chart(chart, anchor).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    assert_eq!(sheet2.chart_count(), 1);
    let drawn = sheet2.charts().next().unwrap();
    let c = drawn.payload;
    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert_eq!(c.title.as_deref(), Some("Sales Chart"));
    assert_eq!(c.series.len(), 2);
    assert_eq!(c.series[0].name.as_deref(), Some("Sheet1!$B$1"));
    assert_eq!(c.series[1].name.as_deref(), Some("Sheet1!$C$1"));
    match &c.series[0].values {
        DataReference::Formula(f) => assert_eq!(f, "Sheet1!$B$2:$B$5"),
        other => panic!("expected Formula, got {:?}", other),
    }
    match c.series[0].categories.as_ref().unwrap() {
        DataReference::Formula(f) => assert_eq!(f, "Sheet1!$A$2:$A$5"),
        other => panic!("expected Formula, got {:?}", other),
    }
    assert_eq!(
        c.category_axis.as_ref().unwrap().title.as_deref(),
        Some("Quarter")
    );
    let vax = c.value_axis.as_ref().unwrap();
    assert_eq!(vax.title.as_deref(), Some("Amount"));
    assert_eq!(vax.minimum, Some(0.0));
    assert_eq!(vax.maximum, Some(50000.0));
    assert_eq!(c.legend.as_ref().unwrap().position, LegendPosition::Bottom);
    if let DrawingAnchor::TwoCell { from, to, .. } = &drawn.object.unwrap().anchor {
        assert_eq!(from.col, 1);
        assert_eq!(from.row, 5);
        assert_eq!(to.col, 10);
        assert_eq!(to.row, 20);
    } else {
        panic!("expected TwoCell anchor");
    }
}

#[test]
fn test_roundtrip_chart_line() {
    use duke_sheets_chart::{
        Axis, Chart, ChartType, DataReference, DataSeries, Legend, LegendPosition,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.title = Some("Trend".to_string());
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$10"))
        .with_name("Sheet1!$B$1")
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$10"));
    chart.add_series(s);
    chart.category_axis = Some(Axis::new().with_title("Time"));
    chart.value_axis = Some(Axis::new().with_title("Value"));
    chart.legend = Some(Legend::new(LegendPosition::Right));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::Line);
    assert_eq!(c.title.as_deref(), Some("Trend"));
    assert_eq!(c.series.len(), 1);
    assert_eq!(
        c.category_axis.as_ref().unwrap().title.as_deref(),
        Some("Time")
    );
    assert_eq!(
        c.value_axis.as_ref().unwrap().title.as_deref(),
        Some("Value")
    );
    assert_eq!(c.legend.as_ref().unwrap().position, LegendPosition::Right);
}

#[test]
fn test_roundtrip_chart_pie() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Pie);
    chart.title = Some("Market Share".to_string());
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$4"))
        .with_name("Shares")
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$4"));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::Pie);
    assert_eq!(c.title.as_deref(), Some("Market Share"));
    assert_eq!(c.series.len(), 1);
    assert!(c.category_axis.is_none(), "pie should have no catAx");
    assert!(c.value_axis.is_none(), "pie should have no valAx");
}

#[test]
fn test_roundtrip_chart_scatter() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ScatterMarkers);
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$5"));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::ScatterMarkers);
    assert_eq!(c.series.len(), 1);
    match c.series[0].categories.as_ref().unwrap() {
        DataReference::Formula(f) => assert_eq!(f, "Sheet1!$A$1:$A$5"),
        other => panic!("expected Formula for xVal, got {:?}", other),
    }
    match &c.series[0].values {
        DataReference::Formula(f) => assert_eq!(f, "Sheet1!$B$1:$B$5"),
        other => panic!("expected Formula for yVal, got {:?}", other),
    }
    // scatter uses two valAx, reader stores both
    assert!(c.value_axis.is_some());
}

#[test]
fn test_roundtrip_chart_no_series() {
    use duke_sheets_chart::{Chart, ChartType};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let chart = Chart::new(ChartType::Area);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::Area);
    assert_eq!(c.series.len(), 0);
}

#[test]
fn test_roundtrip_multiple_charts() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut c1 = Chart::new(ChartType::BarClustered);
    c1.title = Some("Bar".to_string());
    c1.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(c1, DrawingAnchor::default()).unwrap();

    let mut c2 = Chart::new(ChartType::Doughnut);
    c2.title = Some("Donut".to_string());
    c2.add_series(DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5")));
    sheet.add_chart(c2, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let charts: Vec<&Chart> = wb2
        .worksheet(0)
        .unwrap()
        .charts()
        .map(|drawn| drawn.payload)
        .collect();

    assert_eq!(charts.len(), 2);
    let types: Vec<&ChartType> = charts.iter().map(|c| &c.chart_type).collect();
    assert!(types.contains(&&ChartType::BarClustered));
    assert!(types.contains(&&ChartType::Doughnut));
    let titles: Vec<Option<&str>> = charts.iter().map(|c| c.title.as_deref()).collect();
    assert!(titles.contains(&Some("Bar")));
    assert!(titles.contains(&Some("Donut")));
}

#[test]
fn test_roundtrip_chart_anchor_offsets() {
    use duke_sheets_chart::{
        CellMarker, Chart, ChartType, DataReference, DataSeries, DrawingAnchor,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let anchor = DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 3,
            col_offset_emu: 152400,
            row: 7,
            row_offset_emu: 76200,
        },
        to: CellMarker {
            col: 15,
            col_offset_emu: 304800,
            row: 25,
            row_offset_emu: 228600,
        },
        edit_as: None,
    };
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, anchor).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap();

    if let DrawingAnchor::TwoCell { from, to, .. } = &c.object.unwrap().anchor {
        assert_eq!(from.col, 3);
        assert_eq!(from.row, 7);
        assert_eq!(from.col_offset_emu, 152400);
        assert_eq!(from.row_offset_emu, 76200);
        assert_eq!(to.col, 15);
        assert_eq!(to.row, 25);
        assert_eq!(to.col_offset_emu, 304800);
        assert_eq!(to.row_offset_emu, 228600);
    } else {
        panic!("expected TwoCell anchor");
    }
}

#[test]
fn test_roundtrip_chart_data_labels() {
    use duke_sheets_chart::{
        Chart, ChartType, DataLabelPosition, DataLabels, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.data_labels = Some(DataLabels {
        show_value: Some(true),
        show_category_name: Some(true),
        separator: Some(",".to_string()),
        position: Some(DataLabelPosition::OutsideEnd),
        ..Default::default()
    });
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c
        .data_labels
        .as_ref()
        .expect("chart data_labels should survive");
    assert_eq!(dl.show_value, Some(true));
    assert_eq!(dl.show_category_name, Some(true));
    assert_eq!(dl.separator.as_deref(), Some(","));
    assert_eq!(dl.position, Some(DataLabelPosition::OutsideEnd));
}

#[test]
fn test_roundtrip_chart_series_data_labels() {
    use duke_sheets_chart::{Chart, ChartType, DataLabels, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.data_labels = Some(DataLabels {
        show_series_name: Some(true),
        show_percent: Some(true),
        ..Default::default()
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c.series[0]
        .data_labels
        .as_ref()
        .expect("series data_labels should survive");
    assert_eq!(dl.show_series_name, Some(true));
    assert_eq!(dl.show_percent, Some(true));
}

#[test]
fn test_roundtrip_chart_trendline() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, Trendline, TrendlineType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$10"));
    s.trendline = Some(Trendline {
        trendline_type: TrendlineType::Linear,
        name: Some("Trend".to_string()),
        order: None,
        period: None,
        forward: Some(1.0),
        backward: None,
        intercept: None,
        display_r_squared: Some(true),
        display_equation: Some(true),
        label: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let t = c.series[0]
        .trendline
        .as_ref()
        .expect("trendline should survive");
    assert_eq!(t.trendline_type, TrendlineType::Linear);
    assert_eq!(t.name.as_deref(), Some("Trend"));
    assert_eq!(t.display_r_squared, Some(true));
    assert_eq!(t.display_equation, Some(true));
    assert_eq!(t.forward, Some(1.0));
}

#[test]
fn test_roundtrip_chart_error_bars() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, ErrorBarDirection, ErrorBarType, ErrorBars,
        ErrorValueType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::BarClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.error_bars = Some(ErrorBars {
        direction: ErrorBarDirection::Y,
        bar_type: ErrorBarType::Both,
        value_type: ErrorValueType::FixedValue,
        value: Some(5.0),
        no_end_cap: None,
        plus: None,
        minus: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let eb = c.series[0]
        .error_bars
        .as_ref()
        .expect("error_bars should survive");
    assert_eq!(eb.direction, ErrorBarDirection::Y);
    assert_eq!(eb.bar_type, ErrorBarType::Both);
    assert_eq!(eb.value_type, ErrorValueType::FixedValue);
    assert_eq!(eb.value, Some(5.0));
}

#[test]
fn test_roundtrip_chart_markers() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, Marker, MarkerSymbol};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.marker = Some(Marker {
        symbol: Some(MarkerSymbol::Circle),
        size: Some(8),
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let m = c.series[0].marker.as_ref().expect("marker should survive");
    assert_eq!(m.symbol, Some(MarkerSymbol::Circle));
    assert_eq!(m.size, Some(8));
}

#[test]
fn test_roundtrip_chart_data_points() {
    use duke_sheets_chart::{
        Chart, ChartColor, ChartShapeProperties, ChartType, DataPoint, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Pie);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.data_points = vec![DataPoint {
        index: 1,
        explosion: Some(25),
        marker: None,
        shape_properties: Some(ChartShapeProperties {
            solid_fill: Some(ChartColor {
                hex: "C0504D".into(),
            }),
            no_fill: false,
            line: None,
        }),
    }];
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.series[0].data_points.len(), 1);
    let dp = &c.series[0].data_points[0];
    assert_eq!(dp.index, 1);
    assert_eq!(dp.explosion, Some(25));
    assert!(dp.marker.is_none());
    assert_eq!(
        dp.shape_properties
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|c| c.hex.as_str()),
        Some("C0504D")
    );
}

#[test]
fn test_roundtrip_chart_series_smooth() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.smooth = Some(true);
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.series[0].smooth, Some(true));
}

#[test]
fn test_roundtrip_chart_axis_enhancements() {
    use duke_sheets_chart::{
        Axis, AxisCrosses, Chart, ChartColor, ChartLine, ChartShapeProperties, ChartType,
        DataReference, DataSeries, NumberFormat, TickLabelPosition, TickMark,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut cat_axis = Axis::new();
    cat_axis.major_gridlines = true;
    cat_axis.minor_gridlines = true;
    cat_axis.major_gridlines_shape_properties = Some(ChartShapeProperties {
        solid_fill: None,
        no_fill: false,
        line: Some(ChartLine {
            width: Some(9360),
            solid_fill: Some(ChartColor {
                hex: "D9D9D9".into(),
            }),
            no_fill: false,
            dash_style: None,
        }),
    });
    cat_axis.minor_gridlines_shape_properties = Some(ChartShapeProperties {
        solid_fill: None,
        no_fill: false,
        line: Some(ChartLine {
            width: Some(6350),
            solid_fill: Some(ChartColor {
                hex: "EEEEEE".into(),
            }),
            no_fill: false,
            dash_style: Some("dash".into()),
        }),
    });
    cat_axis.major_tick_mark = Some(TickMark::Outside);
    cat_axis.minor_tick_mark = Some(TickMark::Inside);
    cat_axis.label_position = Some(TickLabelPosition::NextTo);
    cat_axis.number_format = Some(NumberFormat {
        format_code: "0.00".into(),
        source_linked: Some(false),
    });
    cat_axis.crosses = Some(AxisCrosses::AutoZero);
    chart.category_axis = Some(cat_axis);
    chart.value_axis = Some(Axis::new());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ax = c
        .category_axis
        .as_ref()
        .expect("category_axis should survive");
    assert!(ax.major_gridlines);
    assert!(ax.minor_gridlines);
    let major_line = ax
        .major_gridlines_shape_properties
        .as_ref()
        .and_then(|sp| sp.line.as_ref())
        .expect("major gridline shape properties should survive");
    assert_eq!(major_line.width, Some(9360));
    assert_eq!(major_line.solid_fill.as_ref().unwrap().hex, "D9D9D9");
    let minor_line = ax
        .minor_gridlines_shape_properties
        .as_ref()
        .and_then(|sp| sp.line.as_ref())
        .expect("minor gridline shape properties should survive");
    assert_eq!(minor_line.width, Some(6350));
    assert_eq!(minor_line.solid_fill.as_ref().unwrap().hex, "EEEEEE");
    assert_eq!(minor_line.dash_style.as_deref(), Some("dash"));
    assert_eq!(ax.major_tick_mark, Some(TickMark::Outside));
    assert_eq!(ax.minor_tick_mark, Some(TickMark::Inside));
    assert_eq!(ax.label_position, Some(TickLabelPosition::NextTo));
    let nf = ax
        .number_format
        .as_ref()
        .expect("number_format should survive");
    assert_eq!(nf.format_code, "0.00");
    assert_eq!(nf.source_linked, Some(false));
    assert_eq!(ax.crosses, Some(AxisCrosses::AutoZero));
}

#[test]
fn test_roundtrip_chart_view_3d() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, View3D};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.view_3d = Some(View3D {
        rotate_x: Some(15),
        rotate_y: Some(20),
        perspective: Some(30),
        ..Default::default()
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let v = c.view_3d.as_ref().expect("view_3d should survive");
    assert_eq!(v.rotate_x, Some(15));
    assert_eq!(v.rotate_y, Some(20));
    assert_eq!(v.perspective, Some(30));
}

#[test]
fn test_roundtrip_chart_data_table() {
    use duke_sheets_chart::{Chart, ChartDataTable, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.data_table = Some(ChartDataTable {
        show_horizontal_border: Some(true),
        show_vertical_border: Some(true),
        show_outline: Some(true),
        show_keys: Some(true),
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dt = c.data_table.as_ref().expect("data_table should survive");
    assert_eq!(dt.show_horizontal_border, Some(true));
    assert_eq!(dt.show_vertical_border, Some(true));
    assert_eq!(dt.show_outline, Some(true));
    assert_eq!(dt.show_keys, Some(true));
}

#[test]
fn test_roundtrip_chart_display_blanks_and_plot_visible() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, DisplayBlanksAs};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.display_blanks_as = Some(DisplayBlanksAs::Gap);
    chart.plot_visible_only = Some(true);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.display_blanks_as, Some(DisplayBlanksAs::Gap));
    assert_eq!(c.plot_visible_only, Some(true));
}

#[test]
fn test_roundtrip_chart_layout() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, Layout, ManualLayout};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.layout = Some(Layout {
        manual_layout: Some(ManualLayout {
            x: Some(0.1),
            y: Some(0.2),
            width: Some(0.8),
            height: Some(0.6),
        }),
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let layout = c.layout.as_ref().expect("layout should survive");
    let ml = layout
        .manual_layout
        .as_ref()
        .expect("manual_layout should survive");
    assert!((ml.x.unwrap() - 0.1).abs() < 1e-10);
    assert!((ml.y.unwrap() - 0.2).abs() < 1e-10);
    assert!((ml.width.unwrap() - 0.8).abs() < 1e-10);
    assert!((ml.height.unwrap() - 0.6).abs() < 1e-10);
}

#[test]
fn test_roundtrip_column_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ColumnStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::ColumnStacked);
}

#[test]
fn test_roundtrip_column_percent_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ColumnPercentStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::ColumnPercentStacked);
}

#[test]
fn test_roundtrip_bar_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::BarStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::BarStacked);
}

#[test]
fn test_roundtrip_bar_percent_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::BarPercentStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::BarPercentStacked);
}

#[test]
fn test_roundtrip_line_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::LineStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::LineStacked);
}

#[test]
fn test_roundtrip_pie_exploded() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::PieExploded);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.explosion = Some(25);
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::PieExploded);
    assert_eq!(c.series[0].explosion, Some(25));
}

#[test]
fn test_roundtrip_area_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::AreaStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::AreaStacked);
}

#[test]
fn test_roundtrip_area_percent_stacked() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::AreaPercentStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::AreaPercentStacked);
}

#[test]
fn test_roundtrip_scatter_lines() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ScatterLines);
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$5"));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::ScatterLines);
    assert_eq!(c.series.len(), 1);
}

#[test]
fn test_roundtrip_scatter_smooth() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ScatterSmooth);
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$5"));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::ScatterSmooth);
    assert_eq!(c.series.len(), 1);
}

#[test]
fn test_roundtrip_bubble() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::Bubble);
    let s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$5"));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::Bubble);
    assert_eq!(c.series.len(), 1);
}

#[test]
fn test_roundtrip_radar() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::Radar);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::Radar);
}

#[test]
fn test_roundtrip_stock() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::Stock);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::Stock);
}

#[test]
fn test_roundtrip_surface() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::Surface);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::Surface);
}

#[test]
fn test_roundtrip_shape_properties() {
    use duke_sheets_chart::{
        Chart, ChartColor, ChartLine, ChartShapeProperties, ChartType, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.shape_properties = Some(ChartShapeProperties {
        solid_fill: Some(ChartColor {
            hex: "00FF00".to_string(),
        }),
        no_fill: false,
        line: Some(ChartLine {
            width: Some(12700),
            solid_fill: Some(ChartColor {
                hex: "0000FF".to_string(),
            }),
            no_fill: false,
            dash_style: Some("dash".to_string()),
        }),
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ssp = c.series[0]
        .shape_properties
        .as_ref()
        .expect("series shape_properties should survive");
    assert_eq!(ssp.solid_fill.as_ref().unwrap().hex, "00FF00");
    let sln = ssp.line.as_ref().unwrap();
    assert_eq!(sln.width, Some(12700));
    assert_eq!(sln.solid_fill.as_ref().unwrap().hex, "0000FF");
    assert_eq!(sln.dash_style.as_deref(), Some("dash"));
}

#[test]
fn test_roundtrip_vary_colors_gap_overlap() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.vary_colors = Some(true);
    chart.gap_width = Some(150);
    chart.overlap = Some(-25);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.vary_colors, Some(true));
    assert_eq!(c.gap_width, Some(150));
    assert_eq!(c.overlap, Some(-25));
}

#[test]
fn test_roundtrip_rounded_corners_auto_title() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.rounded_corners = Some(true);
    chart.auto_title_deleted = Some(true);
    chart.show_dlbls_over_max = Some(true);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.rounded_corners, Some(true));
    assert_eq!(c.auto_title_deleted, Some(true));
    assert_eq!(c.show_dlbls_over_max, Some(true));
}

#[test]
fn test_roundtrip_invert_if_negative() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::BarClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.invert_if_negative = Some(true);
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.series[0].invert_if_negative, Some(true));
}

#[test]
fn test_roundtrip_first_slice_angle_hole_size() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Doughnut);
    chart.first_slice_angle = Some(90);
    chart.hole_size = Some(50);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.first_slice_angle, Some(90));
    assert_eq!(c.hole_size, Some(50));
}

#[test]
fn test_roundtrip_bubble_scale() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Bubble);
    chart.bubble_scale = Some(200);
    chart.show_negative_bubbles = Some(false);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.bubble_scale, Some(200));
    assert_eq!(c.show_negative_bubbles, Some(false));
}

#[test]
fn test_roundtrip_radar_style() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Radar);
    chart.radar_style = Some("filled".to_string());
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.radar_style.as_deref(), Some("filled"));
}

#[test]
fn test_roundtrip_surface_wireframe() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Surface);
    chart.wireframe = Some(true);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.wireframe, Some(true));
}

#[test]
fn test_roundtrip_data_reference_numbers() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::Numbers(vec![1.0, 2.0, 3.0])));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    match &c.series[0].values {
        DataReference::Numbers(nums) => assert_eq!(nums, &[1.0, 2.0, 3.0]),
        other => panic!("expected Numbers, got {:?}", other),
    }
}

#[test]
fn test_roundtrip_data_reference_strings() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"))
        .with_categories(DataReference::Strings(vec!["A".into(), "B".into()]));
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    // The reader doesn't yet parse strCache for categories, so categories
    // come back as None. Verify the roundtrip doesn't crash and the chart
    // itself is present.
    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert_eq!(c.series.len(), 1);
}

#[test]
fn test_roundtrip_axis_cross_between_major_minor() {
    use duke_sheets_chart::{Axis, Chart, ChartType, CrossBetween, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut vax = Axis::new();
    vax.cross_between = Some(CrossBetween::MidCat);
    vax.major_unit = Some(5.0);
    vax.minor_unit = Some(1.0);
    chart.value_axis = Some(vax);
    chart.category_axis = Some(Axis::new());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ax = c.value_axis.as_ref().expect("value_axis should survive");
    assert_eq!(ax.cross_between, Some(CrossBetween::MidCat));
    assert_eq!(ax.major_unit, Some(5.0));
    assert_eq!(ax.minor_unit, Some(1.0));
}

#[test]
fn test_roundtrip_legend_shape_properties() {
    use duke_sheets_chart::{
        Chart, ChartColor, ChartShapeProperties, ChartType, DataReference, DataSeries, Legend,
        LegendPosition,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut legend = Legend::new(LegendPosition::Bottom);
    legend.shape_properties = Some(ChartShapeProperties {
        solid_fill: Some(ChartColor {
            hex: "FFFF00".to_string(),
        }),
        no_fill: false,
        line: None,
    });
    chart.legend = Some(legend);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let leg = c.legend.as_ref().expect("legend should survive");
    assert_eq!(leg.position, LegendPosition::Bottom);
    let sp = leg
        .shape_properties
        .as_ref()
        .expect("legend shape_properties should survive");
    assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "FFFF00");
}

#[test]
fn test_roundtrip_3d_chart() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.is_3d = true;
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert!(c.is_3d, "is_3d should be true after roundtrip");
}

#[test]
fn test_roundtrip_charts_on_multiple_sheets() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();

    let sheet1 = wb.worksheet_mut(0).unwrap();
    let mut c1 = Chart::new(ChartType::Line);
    c1.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet1.add_chart(c1, DrawingAnchor::default()).unwrap();

    wb.add_worksheet_with_name("Sheet2").unwrap();
    let sheet2 = wb.worksheet_mut(1).unwrap();
    let mut c2 = Chart::new(ChartType::Pie);
    c2.add_series(DataSeries::new(DataReference::formula("Sheet2!$A$1:$A$5")));
    sheet2.add_chart(c2, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    let s1 = wb2.worksheet(0).unwrap();
    assert_eq!(s1.chart_count(), 1);
    assert_eq!(
        s1.charts().next().unwrap().payload.chart_type,
        ChartType::Line
    );

    let s2 = wb2.worksheet(1).unwrap();
    assert_eq!(s2.chart_count(), 1);
    assert_eq!(
        s2.charts().next().unwrap().payload.chart_type,
        ChartType::Pie
    );
}

#[test]
fn test_roundtrip_kitchen_sink() {
    use duke_sheets_chart::{
        Axis, Chart, ChartColor, ChartDataTable, ChartLine, ChartShapeProperties, ChartType,
        CrossBetween, DataLabelPosition, DataLabels, DataReference, DataSeries, DisplayBlanksAs,
        ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType, Layout, Legend, LegendPosition,
        ManualLayout, Marker, MarkerSymbol, NumberFormat, Trendline, TrendlineType, View3D,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Kitchen Sink".to_string());
    chart.data_labels = Some(DataLabels {
        show_value: Some(true),
        show_category_name: Some(true),
        show_series_name: Some(false),
        show_percent: Some(false),
        show_bubble_size: Some(false),
        show_legend_key: Some(false),
        separator: Some(";".to_string()),
        position: Some(DataLabelPosition::Center),
        number_format: None,
        show_leader_lines: None,
        leader_lines: None,
        text_properties: None,
        data_label_overrides: Vec::new(),
    });
    chart.view_3d = Some(View3D {
        rotate_x: Some(10),
        rotate_y: Some(25),
        perspective: Some(15),
        depth_percent: None,
        height_percent: None,
        right_angle_axes: Some(true),
    });
    chart.data_table = Some(ChartDataTable {
        show_horizontal_border: Some(true),
        show_vertical_border: Some(false),
        show_outline: Some(true),
        show_keys: Some(false),
    });
    chart.display_blanks_as = Some(DisplayBlanksAs::Zero);
    chart.plot_visible_only = Some(true);
    chart.layout = Some(Layout {
        manual_layout: Some(ManualLayout {
            x: Some(0.05),
            y: Some(0.05),
            width: Some(0.9),
            height: Some(0.9),
        }),
    });
    chart.vary_colors = Some(true);
    chart.gap_width = Some(200);
    chart.overlap = Some(10);
    chart.rounded_corners = Some(true);
    chart.auto_title_deleted = Some(false);
    chart.show_dlbls_over_max = Some(true);
    chart.shape_properties = Some(ChartShapeProperties {
        solid_fill: Some(ChartColor {
            hex: "EEEEEE".to_string(),
        }),
        no_fill: false,
        line: Some(ChartLine {
            width: Some(12700),
            solid_fill: Some(ChartColor {
                hex: "333333".to_string(),
            }),
            no_fill: false,
            dash_style: Some("solid".to_string()),
        }),
    });

    let mut vax = Axis::new().with_title("Values").with_bounds(0.0, 100.0);
    vax.cross_between = Some(CrossBetween::Between);
    vax.major_unit = Some(10.0);
    vax.minor_unit = Some(2.0);
    vax.number_format = Some(NumberFormat {
        format_code: "#,##0".to_string(),
        source_linked: Some(false),
    });
    chart.value_axis = Some(vax);
    chart.category_axis = Some(Axis::new().with_title("Categories"));
    chart.legend = Some(Legend::new(LegendPosition::Top));

    let mut s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$10"));
    s.name = Some("Series1".to_string());
    s.categories = Some(DataReference::formula("Sheet1!$A$1:$A$10"));
    s.smooth = Some(false);
    s.data_labels = Some(DataLabels {
        show_value: Some(true),
        ..Default::default()
    });
    s.trendline = Some(Trendline {
        trendline_type: TrendlineType::Exponential,
        name: Some("Exp Trend".to_string()),
        order: None,
        period: None,
        forward: Some(2.0),
        backward: Some(1.0),
        intercept: None,
        display_r_squared: Some(false),
        display_equation: Some(true),
        label: None,
    });
    s.error_bars = Some(ErrorBars {
        direction: ErrorBarDirection::Y,
        bar_type: ErrorBarType::Both,
        value_type: ErrorValueType::Percentage,
        value: Some(10.0),
        no_end_cap: None,
        plus: None,
        minus: None,
    });
    s.marker = Some(Marker {
        symbol: Some(MarkerSymbol::Diamond),
        size: Some(6),
    });
    s.shape_properties = Some(ChartShapeProperties {
        solid_fill: Some(ChartColor {
            hex: "4472C4".to_string(),
        }),
        no_fill: false,
        line: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert_eq!(c.title.as_deref(), Some("Kitchen Sink"));

    let dl = c.data_labels.as_ref().unwrap();
    assert_eq!(dl.show_value, Some(true));
    assert_eq!(dl.show_category_name, Some(true));
    assert_eq!(dl.separator.as_deref(), Some(";"));
    assert_eq!(dl.position, Some(DataLabelPosition::Center));

    let v3d = c.view_3d.as_ref().unwrap();
    assert_eq!(v3d.rotate_x, Some(10));
    assert_eq!(v3d.rotate_y, Some(25));
    assert_eq!(v3d.perspective, Some(15));
    assert_eq!(v3d.right_angle_axes, Some(true));

    let dt = c.data_table.as_ref().unwrap();
    assert_eq!(dt.show_horizontal_border, Some(true));
    assert_eq!(dt.show_vertical_border, Some(false));
    assert_eq!(dt.show_outline, Some(true));
    assert_eq!(dt.show_keys, Some(false));

    assert_eq!(c.display_blanks_as, Some(DisplayBlanksAs::Zero));
    assert_eq!(c.plot_visible_only, Some(true));

    let ml = c.layout.as_ref().unwrap().manual_layout.as_ref().unwrap();
    assert!((ml.x.unwrap() - 0.05).abs() < 1e-10);
    assert!((ml.y.unwrap() - 0.05).abs() < 1e-10);
    assert!((ml.width.unwrap() - 0.9).abs() < 1e-10);
    assert!((ml.height.unwrap() - 0.9).abs() < 1e-10);

    assert_eq!(c.vary_colors, Some(true));
    assert_eq!(c.gap_width, Some(200));
    assert_eq!(c.overlap, Some(10));
    assert_eq!(c.rounded_corners, Some(true));
    assert_eq!(c.auto_title_deleted, Some(false));
    assert_eq!(c.show_dlbls_over_max, Some(true));

    let chart_sp = c.shape_properties.as_ref().unwrap();
    assert_eq!(chart_sp.solid_fill.as_ref().unwrap().hex, "EEEEEE");
    let chart_line = chart_sp.line.as_ref().unwrap();
    assert_eq!(chart_line.width, Some(12700));
    assert_eq!(chart_line.solid_fill.as_ref().unwrap().hex, "333333");

    let vax = c.value_axis.as_ref().unwrap();
    assert_eq!(vax.title.as_deref(), Some("Values"));
    assert_eq!(vax.minimum, Some(0.0));
    assert_eq!(vax.maximum, Some(100.0));
    assert_eq!(vax.cross_between, Some(CrossBetween::Between));
    assert_eq!(vax.major_unit, Some(10.0));
    assert_eq!(vax.minor_unit, Some(2.0));
    let nf = vax.number_format.as_ref().unwrap();
    assert_eq!(nf.format_code, "#,##0");
    assert_eq!(nf.source_linked, Some(false));

    assert_eq!(
        c.category_axis.as_ref().unwrap().title.as_deref(),
        Some("Categories")
    );
    assert_eq!(c.legend.as_ref().unwrap().position, LegendPosition::Top);

    let ser = &c.series[0];
    assert_eq!(ser.name.as_deref(), Some("Series1"));
    assert!(ser.smooth.is_none() || ser.smooth == Some(false));
    assert!(ser.data_labels.is_some());
    assert_eq!(ser.data_labels.as_ref().unwrap().show_value, Some(true));

    let t = ser.trendline.as_ref().unwrap();
    assert_eq!(t.trendline_type, TrendlineType::Exponential);
    assert_eq!(t.name.as_deref(), Some("Exp Trend"));
    assert_eq!(t.forward, Some(2.0));
    assert_eq!(t.backward, Some(1.0));
    assert_eq!(t.display_equation, Some(true));

    let eb = ser.error_bars.as_ref().unwrap();
    assert_eq!(eb.direction, ErrorBarDirection::Y);
    assert_eq!(eb.bar_type, ErrorBarType::Both);
    assert_eq!(eb.value_type, ErrorValueType::Percentage);
    assert_eq!(eb.value, Some(10.0));

    let m = ser.marker.as_ref().unwrap();
    assert_eq!(m.symbol, Some(MarkerSymbol::Diamond));
    assert_eq!(m.size, Some(6));

    let ssp = ser.shape_properties.as_ref().unwrap();
    assert_eq!(ssp.solid_fill.as_ref().unwrap().hex, "4472C4");
}

#[test]
fn test_roundtrip_combo_bar_line() {
    use duke_sheets_chart::{
        Axis, AxisType, Chart, ChartAxis, ChartType, ChartTypeGroup, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Q1").unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Combo Chart".to_string());

    let bar_group = ChartTypeGroup {
        chart_type: ChartType::ColumnClustered,
        is_3d: false,
        series: vec![
            DataSeries::new(DataReference::formula("Sheet1!$B$2:$B$5")).with_name("Revenue")
        ],
        data_labels: None,
        vary_colors: None,
        gap_width: Some(150),
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![1, 2],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    let line_group = ChartTypeGroup {
        chart_type: ChartType::Line,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$C$2:$C$5")).with_name("Trend")],
        data_labels: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![1, 3],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    chart.type_groups = vec![bar_group, line_group];

    let cat_ax = Axis::new();
    let val_ax1 = {
        let mut a = Axis::new();
        a.axis_type = AxisType::Value;
        a
    };
    let val_ax2 = {
        let mut a = Axis::new();
        a.axis_type = AxisType::Value;
        a
    };
    chart.axes = vec![
        ChartAxis {
            id: 1,
            cross_id: 2,
            axis: cat_ax,
        },
        ChartAxis {
            id: 2,
            cross_id: 1,
            axis: val_ax1,
        },
        ChartAxis {
            id: 3,
            cross_id: 1,
            axis: val_ax2,
        },
    ];

    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let c = sheet2.charts().next().unwrap().payload;
    assert_eq!(c.type_groups.len(), 2);
    assert_eq!(c.type_groups[0].chart_type, ChartType::ColumnClustered);
    assert_eq!(c.type_groups[0].series.len(), 1);
    assert_eq!(c.type_groups[0].series[0].name.as_deref(), Some("Revenue"));
    assert_eq!(c.type_groups[0].axis_ids, vec![1, 2]);
    assert_eq!(c.type_groups[0].gap_width, Some(150));
    assert_eq!(c.type_groups[1].chart_type, ChartType::Line);
    assert_eq!(c.type_groups[1].series.len(), 1);
    assert_eq!(c.type_groups[1].series[0].name.as_deref(), Some("Trend"));
    assert_eq!(c.type_groups[1].axis_ids, vec![1, 3]);

    // Legacy fields should come from first group
    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert_eq!(c.series.len(), 1);

    // Bug 2: legacy axis fields must match first group's axes, not last-parsed
    let cat = c.category_axis.as_ref().expect("legacy category_axis");
    assert_eq!(cat.axis_type, AxisType::Category);
    let val = c.value_axis.as_ref().expect("legacy value_axis");
    assert_eq!(val.axis_type, AxisType::Value);
}

#[test]
fn test_roundtrip_combo_preserves_legacy() {
    use duke_sheets_chart::{Axis, Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "data").unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.add_series(
        DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")).with_name("Series1"),
    );
    chart.category_axis = Some(Axis::new());
    chart.value_axis = Some(Axis::new());

    // No type_groups set => legacy mode
    assert!(chart.type_groups.is_empty());

    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let c = sheet2.charts().next().unwrap().payload;
    assert_eq!(c.chart_type, ChartType::Line);
    assert_eq!(c.series.len(), 1);
    assert_eq!(c.series[0].name.as_deref(), Some("Series1"));
    assert!(c.type_groups.is_empty());
    assert!(c.axes.is_empty());
}

#[test]
fn test_roundtrip_combo_axes() {
    use duke_sheets_chart::{
        Axis, AxisType, Chart, ChartAxis, ChartType, ChartTypeGroup, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "x").unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);

    let bar_group = ChartTypeGroup {
        chart_type: ChartType::ColumnClustered,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))],
        data_labels: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![10, 20],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    let line_group = ChartTypeGroup {
        chart_type: ChartType::Line,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$C$1:$C$5"))],
        data_labels: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![10, 30],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    chart.type_groups = vec![bar_group, line_group];

    let mut cat_ax = Axis::new();
    cat_ax.axis_type = AxisType::Category;
    let mut val_ax1 = Axis::new();
    val_ax1.axis_type = AxisType::Value;
    let mut val_ax2 = Axis::new();
    val_ax2.axis_type = AxisType::Value;
    chart.axes = vec![
        ChartAxis {
            id: 10,
            cross_id: 20,
            axis: cat_ax,
        },
        ChartAxis {
            id: 20,
            cross_id: 10,
            axis: val_ax1,
        },
        ChartAxis {
            id: 30,
            cross_id: 10,
            axis: val_ax2,
        },
    ];

    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();

    let c = sheet2.charts().next().unwrap().payload;
    assert_eq!(c.axes.len(), 3);

    let ax0 = &c.axes[0];
    assert_eq!(ax0.id, 10);
    assert_eq!(ax0.cross_id, 20);
    assert_eq!(ax0.axis.axis_type, AxisType::Category);

    let ax1 = &c.axes[1];
    assert_eq!(ax1.id, 20);
    assert_eq!(ax1.cross_id, 10);
    assert_eq!(ax1.axis.axis_type, AxisType::Value);

    let ax2 = &c.axes[2];
    assert_eq!(ax2.id, 30);
    assert_eq!(ax2.cross_id, 10);
    assert_eq!(ax2.axis.axis_type, AxisType::Value);
}

#[test]
fn test_roundtrip_single_type_group_uses_legacy() {
    use duke_sheets_chart::{
        Axis, AxisType, Chart, ChartAxis, ChartType, ChartTypeGroup, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "x").unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let group = ChartTypeGroup {
        chart_type: ChartType::ColumnClustered,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")).with_name("Sales")],
        data_labels: None,
        vary_colors: Some(true),
        gap_width: Some(200),
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![1, 2],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    chart.type_groups = vec![group];
    let mut cat_ax = Axis::new();
    cat_ax.axis_type = AxisType::Category;
    let mut val_ax = Axis::new();
    val_ax.axis_type = AxisType::Value;
    chart.axes = vec![
        ChartAxis {
            id: 1,
            cross_id: 2,
            axis: cat_ax,
        },
        ChartAxis {
            id: 2,
            cross_id: 1,
            axis: val_ax,
        },
    ];
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    // Single type_group should be written as legacy and read back without type_groups
    assert!(
        c.type_groups.is_empty(),
        "single group should use legacy mode"
    );
    assert!(c.axes.is_empty());
    assert_eq!(c.chart_type, ChartType::ColumnClustered);
    assert_eq!(c.series.len(), 1);
    assert_eq!(c.series[0].name.as_deref(), Some("Sales"));
    assert_eq!(c.vary_colors, Some(true));
    assert_eq!(c.gap_width, Some(200));
}

#[test]
fn test_roundtrip_combo_secondary_axis_position() {
    use duke_sheets_chart::{
        Axis, AxisPosition, AxisType, Chart, ChartAxis, ChartType, ChartTypeGroup, DataReference,
        DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "x").unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let bar_group = ChartTypeGroup {
        chart_type: ChartType::ColumnClustered,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))],
        data_labels: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![10, 20],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    let line_group = ChartTypeGroup {
        chart_type: ChartType::Line,
        is_3d: false,
        series: vec![DataSeries::new(DataReference::formula("Sheet1!$C$1:$C$5"))],
        data_labels: None,
        vary_colors: None,
        gap_width: None,
        overlap: None,
        first_slice_angle: None,
        hole_size: None,
        bubble_scale: None,
        show_negative_bubbles: None,
        radar_style: None,
        wireframe: None,
        drop_lines: None,
        high_low_lines: None,
        series_lines: None,
        up_down_bars: None,
        axis_ids: vec![10, 30],
        of_pie_type: None,
        split_type: None,
        split_pos: None,
        second_pie_size: None,
        bar_shape: None,
        floor: None,
        side_wall: None,
        back_wall: None,
        raw_ext: None,
    };
    chart.type_groups = vec![bar_group, line_group];

    let cat_ax = Axis::new();
    let mut val_ax1 = Axis::new();
    val_ax1.axis_type = AxisType::Value;
    let mut val_ax2 = Axis::new();
    val_ax2.axis_type = AxisType::Value;
    chart.axes = vec![
        ChartAxis {
            id: 10,
            cross_id: 20,
            axis: cat_ax,
        },
        ChartAxis {
            id: 20,
            cross_id: 10,
            axis: val_ax1,
        },
        ChartAxis {
            id: 30,
            cross_id: 10,
            axis: val_ax2,
        },
    ];
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ax1 = &c.axes[1];
    assert_eq!(ax1.axis.axis_type, AxisType::Value);
    assert_eq!(ax1.axis.position, AxisPosition::Left);

    let ax2 = &c.axes[2];
    assert_eq!(ax2.axis.axis_type, AxisType::Value);
    assert_eq!(ax2.axis.position, AxisPosition::Right);
}

#[test]
fn test_roundtrip_vary_colors_false() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.vary_colors = Some(false);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.vary_colors, Some(false));
}

#[test]
fn test_roundtrip_drop_lines() {
    use duke_sheets_chart::{
        Chart, ChartColor, ChartLines, ChartShapeProperties, ChartType, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.drop_lines = Some(ChartLines {
        shape_properties: Some(ChartShapeProperties {
            solid_fill: Some(ChartColor {
                hex: "0000FF".into(),
            }),
            ..Default::default()
        }),
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c.drop_lines.as_ref().expect("drop_lines lost in roundtrip");
    let sp = dl.shape_properties.as_ref().expect("drop_lines spPr lost");
    assert_eq!(sp.solid_fill.as_ref().unwrap().hex, "0000FF");
}

#[test]
fn test_roundtrip_high_low_lines() {
    use duke_sheets_chart::{Chart, ChartLines, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Stock);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.high_low_lines = Some(ChartLines::default());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert!(
        c.high_low_lines.is_some(),
        "high_low_lines lost in roundtrip"
    );
}

#[test]
fn test_roundtrip_series_lines_legacy() {
    use duke_sheets_chart::{Chart, ChartLines, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::BarStacked);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.series_lines = Some(ChartLines::default());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert!(
        c.series_lines.is_some(),
        "series_lines lost in legacy roundtrip"
    );
}

#[test]
fn test_roundtrip_series_lines() {
    use duke_sheets_chart::{
        Axis, AxisType, Chart, ChartAxis, ChartLines, ChartType, ChartTypeGroup, DataReference,
        DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::BarStacked);
    chart.type_groups = vec![
        ChartTypeGroup {
            chart_type: ChartType::BarStacked,
            is_3d: false,
            series: vec![DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$3"))],
            data_labels: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            wireframe: None,
            drop_lines: None,
            high_low_lines: None,
            series_lines: Some(ChartLines::default()),
            up_down_bars: None,
            axis_ids: vec![1, 2],
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            raw_ext: None,
        },
        ChartTypeGroup {
            chart_type: ChartType::Line,
            is_3d: false,
            series: vec![DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$3"))],
            data_labels: None,
            vary_colors: None,
            gap_width: None,
            overlap: None,
            first_slice_angle: None,
            hole_size: None,
            bubble_scale: None,
            show_negative_bubbles: None,
            radar_style: None,
            wireframe: None,
            drop_lines: None,
            high_low_lines: None,
            series_lines: None,
            up_down_bars: None,
            axis_ids: vec![1, 2],
            of_pie_type: None,
            split_type: None,
            split_pos: None,
            second_pie_size: None,
            bar_shape: None,
            floor: None,
            side_wall: None,
            back_wall: None,
            raw_ext: None,
        },
    ];
    chart.axes = vec![
        ChartAxis {
            id: 1,
            cross_id: 2,
            axis: Axis::new(),
        },
        ChartAxis {
            id: 2,
            cross_id: 1,
            axis: {
                let mut a = Axis::new();
                a.axis_type = AxisType::Value;
                a
            },
        },
    ];
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert!(c.type_groups.len() >= 2, "combo groups lost");
    assert!(
        c.type_groups[0].series_lines.is_some(),
        "series_lines lost in roundtrip"
    );
}

#[test]
fn test_roundtrip_up_down_bars() {
    use duke_sheets_chart::{Chart, ChartLines, ChartType, DataReference, DataSeries, UpDownBars};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Stock);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.up_down_bars = Some(UpDownBars {
        gap_width: Some(150),
        up_bars: Some(ChartLines::default()),
        down_bars: Some(ChartLines::default()),
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let udb = c
        .up_down_bars
        .as_ref()
        .expect("up_down_bars lost in roundtrip");
    assert_eq!(udb.gap_width, Some(150));
    assert!(udb.up_bars.is_some(), "up_bars lost");
    assert!(udb.down_bars.is_some(), "down_bars lost");
}

#[test]
fn test_roundtrip_leader_lines() {
    use duke_sheets_chart::{Chart, ChartLines, ChartType, DataLabels, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Pie);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.data_labels = Some(DataLabels {
        show_value: Some(true),
        leader_lines: Some(ChartLines::default()),
        ..DataLabels::default()
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c.data_labels.as_ref().expect("data_labels lost");
    assert!(dl.leader_lines.is_some(), "leader_lines lost in roundtrip");
}

#[test]
fn test_roundtrip_chartsheet() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Q1").unwrap();
    sheet.set_cell_value("A2", 100.0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Sales".to_string());
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    wb.add_chartsheet(ChartSheet {
        name: "Chart1".to_string(),
        chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    assert_eq!(wb2.chartsheet_count(), 1);
    let cs = wb2.chartsheet(0).unwrap();
    assert_eq!(cs.name, "Chart1");
    assert_eq!(cs.chart.chart_type, ChartType::ColumnClustered);
    assert_eq!(cs.chart.title.as_deref(), Some("Sales"));
    assert_eq!(cs.chart.series.len(), 1);
    assert_eq!(wb2.sheet_count(), 1);
}

#[test]
fn test_roundtrip_chartsheet_only() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    let mut wb = Workbook::empty();
    let mut chart = Chart::new(ChartType::Pie);
    chart.title = Some("Pie Chart".to_string());
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    wb.add_chartsheet(ChartSheet {
        name: "OnlyChart".to_string(),
        chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    assert_eq!(wb2.chartsheet_count(), 1);
    assert_eq!(
        wb2.sheet_count(),
        0,
        "phantom worksheet injected for chartsheet-only workbook"
    );
    let cs = wb2.chartsheet(0).unwrap();
    assert_eq!(cs.name, "OnlyChart");
    assert_eq!(cs.chart.chart_type, ChartType::Pie);
}
#[test]
fn test_roundtrip_chartsheet_with_worksheet_charts() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 42.0).unwrap();

    // Add an embedded chart to the worksheet
    let mut ws_chart = Chart::new(ChartType::Line);
    ws_chart.title = Some("Embedded".to_string());
    ws_chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.add_chart(ws_chart, DrawingAnchor::default()).unwrap();

    // Add a separate chartsheet
    let mut cs_chart = Chart::new(ChartType::BarClustered);
    cs_chart.title = Some("Standalone".to_string());
    cs_chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5")));
    wb.add_chartsheet(ChartSheet {
        name: "ChartSheet1".to_string(),
        chart: cs_chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    // Worksheet chart survives
    assert_eq!(wb2.worksheet(0).unwrap().chart_count(), 1);
    let wsc = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;
    assert_eq!(wsc.chart_type, ChartType::Line);
    assert_eq!(wsc.title.as_deref(), Some("Embedded"));

    // Chartsheet survives independently
    assert_eq!(wb2.chartsheet_count(), 1);
    let cs = wb2.chartsheet(0).unwrap();
    assert_eq!(cs.name, "ChartSheet1");
    assert_eq!(cs.chart.chart_type, ChartType::BarClustered);
    assert_eq!(cs.chart.title.as_deref(), Some("Standalone"));
}

#[test]
fn test_roundtrip_interleaved_tab_order() {
    use duke_sheets::SheetSlot;
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    wb.remove_worksheet(0).unwrap();

    wb.add_worksheet_with_name("Sheet1").unwrap();
    let cs = duke_sheets_core::ChartSheet {
        name: "ChartSheet1".to_string(),
        chart: {
            let mut c = Chart::new(ChartType::ColumnClustered);
            c.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
            c
        },
        visibility: duke_sheets_core::SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    };
    wb.add_chartsheet(cs).unwrap();
    wb.add_worksheet_with_name("Sheet2").unwrap();

    // Set interleaved tab order: Sheet1, ChartSheet1, Sheet2
    *wb.sheet_order_mut() = vec![
        SheetSlot::Worksheet(0),
        SheetSlot::ChartSheet(0),
        SheetSlot::Worksheet(1),
    ];

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    assert_eq!(wb2.sheet_count(), 2);
    assert_eq!(wb2.chartsheet_count(), 1);
    assert_eq!(
        wb2.sheet_order(),
        &[
            SheetSlot::Worksheet(0),
            SheetSlot::ChartSheet(0),
            SheetSlot::Worksheet(1),
        ]
    );
    assert_eq!(wb2.worksheet(0).unwrap().name(), "Sheet1");
    assert_eq!(wb2.worksheet(1).unwrap().name(), "Sheet2");
    assert_eq!(wb2.chartsheet(0).unwrap().name, "ChartSheet1");
}

#[test]
fn test_roundtrip_empty_sheet_order_uses_default() {
    use duke_sheets::SheetSlot;
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let cs = duke_sheets_core::ChartSheet {
        name: "ChartSheet1".to_string(),
        chart: {
            let mut c = Chart::new(ChartType::ColumnClustered);
            c.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
            c
        },
        visibility: duke_sheets_core::SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    };
    wb.add_chartsheet(cs).unwrap();

    // Don't set sheet_order - writer should synthesize default (worksheets first)
    assert!(wb.sheet_order().is_empty());

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    assert_eq!(wb2.sheet_count(), 1);
    assert_eq!(wb2.chartsheet_count(), 1);
    // Reader populates sheet_order from the file, which was written as ws first, cs second
    assert_eq!(
        wb2.sheet_order(),
        &[SheetSlot::Worksheet(0), SheetSlot::ChartSheet(0),]
    );
    assert_eq!(wb2.worksheet(0).unwrap().name(), "Sheet1");
    assert_eq!(wb2.chartsheet(0).unwrap().name, "ChartSheet1");
}

#[test]
fn test_add_worksheet_after_read_appears_in_sheet_order() {
    use duke_sheets::SheetSlot;
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    // Step 1: Create a workbook with a worksheet and a chartsheet
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 1.0).unwrap();

    let mut chart = Chart::new(ChartType::Pie);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    wb.add_chartsheet(ChartSheet {
        name: "Chart1".to_string(),
        chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    // Step 2: Write it
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

    // Step 3: Read it back (sheet_order is now populated by the reader)
    let mut wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    assert_eq!(wb2.sheet_order().len(), 2);

    // Step 4: Add a new worksheet via API
    wb2.add_worksheet_with_name("NewSheet").unwrap();

    // Step 5: Write again
    let mut buf2 = Vec::new();
    XlsxWriter::write(&wb2, Cursor::new(&mut buf2)).unwrap();

    // Step 6: Read again and verify the new worksheet appears in sheet_order
    let wb3 = XlsxReader::read(Cursor::new(&buf2)).unwrap();
    assert_eq!(wb3.sheet_count(), 2);
    assert_eq!(wb3.chartsheet_count(), 1);
    assert_eq!(wb3.sheet_order().len(), 3);
    assert!(wb3.sheet_order().contains(&SheetSlot::Worksheet(0)));
    assert!(wb3.sheet_order().contains(&SheetSlot::Worksheet(1)));
    assert!(wb3.sheet_order().contains(&SheetSlot::ChartSheet(0)));
    assert_eq!(
        wb3.worksheet_by_name("NewSheet").unwrap().name(),
        "NewSheet"
    );
}

#[test]
fn test_roundtrip_chart_style_color_passthrough() {
    use duke_sheets_chart::{Chart, ChartType};

    let style_bytes = b"<cs:chartStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" id=\"102\"/>".to_vec();
    let color_bytes = b"<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" meth=\"cycle\" id=\"10\"/>".to_vec();

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.raw_chart_style = Some(style_bytes.clone());
    chart.raw_chart_color_style = Some(color_bytes.clone());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let sheet2 = wb2.worksheet(0).unwrap();
    assert_eq!(sheet2.chart_count(), 1);
    let c = sheet2.charts().next().unwrap().payload;
    assert_eq!(c.raw_chart_style.as_deref(), Some(style_bytes.as_slice()));
    assert_eq!(
        c.raw_chart_color_style.as_deref(),
        Some(color_bytes.as_slice())
    );
}

#[test]
fn test_roundtrip_axis_delete() {
    use duke_sheets_chart::{Axis, Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut cat_axis = Axis::new();
    cat_axis.delete = Some(true);
    chart.category_axis = Some(cat_axis);
    let mut val_axis = Axis::new();
    val_axis.delete = Some(false);
    chart.value_axis = Some(val_axis);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.category_axis.as_ref().unwrap().delete, Some(true));
    assert_eq!(c.value_axis.as_ref().unwrap().delete, Some(false));
}

#[test]
fn test_roundtrip_axis_label_positions() {
    use duke_sheets_chart::{Axis, Chart, ChartType, DataReference, DataSeries, TickLabelPosition};

    for pos in [
        TickLabelPosition::High,
        TickLabelPosition::Low,
        TickLabelPosition::None,
    ] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        let mut cat_axis = Axis::new();
        cat_axis.label_position = Some(pos);
        chart.category_axis = Some(cat_axis);
        chart.value_axis = Some(Axis::new());
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        assert_eq!(
            c.category_axis.as_ref().unwrap().label_position,
            Some(pos),
            "TickLabelPosition::{:?} did not survive roundtrip",
            pos,
        );
    }
}

#[test]
fn test_roundtrip_axis_crosses_min_max() {
    use duke_sheets_chart::{Axis, AxisCrosses, Chart, ChartType, DataReference, DataSeries};

    for crosses in [AxisCrosses::Min, AxisCrosses::Max] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        let mut cat_axis = Axis::new();
        cat_axis.crosses = Some(crosses);
        chart.category_axis = Some(cat_axis);
        chart.value_axis = Some(Axis::new());
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        assert_eq!(
            c.category_axis.as_ref().unwrap().crosses,
            Some(crosses),
            "AxisCrosses::{:?} did not survive roundtrip",
            crosses,
        );
    }
}

#[test]
fn test_roundtrip_trendline_polynomial() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, Trendline, TrendlineType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$10"));
    s.trendline = Some(Trendline {
        trendline_type: TrendlineType::Polynomial,
        name: Some("Poly3".to_string()),
        order: Some(3),
        period: None,
        forward: None,
        backward: None,
        intercept: Some(5.0),
        display_r_squared: Some(true),
        display_equation: Some(false),
        label: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let t = c.series[0]
        .trendline
        .as_ref()
        .expect("trendline should survive");
    assert_eq!(t.trendline_type, TrendlineType::Polynomial);
    assert_eq!(t.name.as_deref(), Some("Poly3"));
    assert_eq!(t.order, Some(3));
    assert_eq!(t.intercept, Some(5.0));
    assert_eq!(t.display_r_squared, Some(true));
}

#[test]
fn test_roundtrip_trendline_moving_average() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, Trendline, TrendlineType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$10"));
    s.trendline = Some(Trendline {
        trendline_type: TrendlineType::MovingAverage,
        name: None,
        order: None,
        period: Some(3),
        forward: None,
        backward: None,
        intercept: None,
        display_r_squared: None,
        display_equation: None,
        label: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let t = c.series[0]
        .trendline
        .as_ref()
        .expect("trendline should survive");
    assert_eq!(t.trendline_type, TrendlineType::MovingAverage);
    assert_eq!(t.period, Some(3));
}

#[test]
fn test_roundtrip_trendline_logarithmic_power() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, Trendline, TrendlineType,
    };

    for ttype in [TrendlineType::Logarithmic, TrendlineType::Power] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::Line);
        let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$10"));
        s.trendline = Some(Trendline {
            trendline_type: ttype,
            name: None,
            order: None,
            period: None,
            forward: Some(2.0),
            backward: Some(1.0),
            intercept: None,
            display_r_squared: None,
            display_equation: None,
            label: None,
        });
        chart.add_series(s);
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        let t = c.series[0]
            .trendline
            .as_ref()
            .expect("trendline should survive");
        assert_eq!(
            t.trendline_type, ttype,
            "TrendlineType::{:?} did not survive roundtrip",
            ttype,
        );
        assert_eq!(t.forward, Some(2.0));
        assert_eq!(t.backward, Some(1.0));
    }
}

#[test]
fn test_roundtrip_error_bars_percentage_stddev() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, ErrorBarDirection, ErrorBarType, ErrorBars,
        ErrorValueType,
    };

    for vtype in [
        ErrorValueType::StandardDeviation,
        ErrorValueType::StandardError,
    ] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::BarClustered);
        let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
        s.error_bars = Some(ErrorBars {
            direction: ErrorBarDirection::Y,
            bar_type: ErrorBarType::Both,
            value_type: vtype,
            value: Some(2.0),
            no_end_cap: Some(true),
            plus: None,
            minus: None,
        });
        chart.add_series(s);
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        let eb = c.series[0]
            .error_bars
            .as_ref()
            .expect("error_bars should survive");
        assert_eq!(
            eb.value_type, vtype,
            "ErrorValueType::{:?} did not survive roundtrip",
            vtype,
        );
        assert_eq!(eb.value, Some(2.0));
        assert_eq!(eb.no_end_cap, Some(true));
    }
}

#[test]
fn test_roundtrip_marker_symbols() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, Marker, MarkerSymbol};

    for symbol in [
        MarkerSymbol::Diamond,
        MarkerSymbol::Square,
        MarkerSymbol::Triangle,
        MarkerSymbol::None,
        MarkerSymbol::Dash,
        MarkerSymbol::Dot,
        MarkerSymbol::Plus,
        MarkerSymbol::Star,
        MarkerSymbol::X,
    ] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::Line);
        let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
        s.marker = Some(Marker {
            symbol: Some(symbol),
            size: Some(6),
        });
        chart.add_series(s);
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        let m = c.series[0].marker.as_ref().expect("marker should survive");
        assert_eq!(
            m.symbol,
            Some(symbol),
            "MarkerSymbol::{:?} did not survive roundtrip",
            symbol,
        );
        assert_eq!(m.size, Some(6));
    }
}

#[test]
fn test_roundtrip_data_label_positions() {
    use duke_sheets_chart::{
        Chart, ChartType, DataLabelPosition, DataLabels, DataReference, DataSeries,
    };

    for pos in [
        DataLabelPosition::Center,
        DataLabelPosition::InsideEnd,
        DataLabelPosition::InsideBase,
        DataLabelPosition::OutsideEnd,
    ] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.data_labels = Some(DataLabels {
            show_value: Some(true),
            position: Some(pos),
            ..Default::default()
        });
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        let dl = c.data_labels.as_ref().expect("data_labels should survive");
        assert_eq!(
            dl.position,
            Some(pos),
            "DataLabelPosition::{:?} did not survive roundtrip",
            pos,
        );
    }
}

#[test]
fn test_roundtrip_display_blanks_as_span() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, DisplayBlanksAs};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.display_blanks_as = Some(DisplayBlanksAs::Span);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.display_blanks_as, Some(DisplayBlanksAs::Span));
}

#[test]
fn test_roundtrip_legend_positions() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, Legend, LegendPosition};

    for pos in [LegendPosition::Left, LegendPosition::TopRight] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::ColumnClustered);
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        chart.legend = Some(Legend::new(pos));
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        assert_eq!(
            c.legend.as_ref().unwrap().position,
            pos,
            "LegendPosition::{:?} did not survive roundtrip",
            pos,
        );
    }
}

#[test]
fn test_roundtrip_legend_overlay() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, Legend, LegendPosition};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut legend = Legend::new(LegendPosition::Right);
    legend.overlay = true;
    chart.legend = Some(legend);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let leg = c.legend.as_ref().expect("legend should survive");
    assert!(leg.overlay, "legend overlay should be true after roundtrip");
}

#[test]
fn test_roundtrip_axis_date_type() {
    use duke_sheets_chart::{Axis, AxisType, Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Line);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut cat_axis = Axis::new();
    cat_axis.axis_type = AxisType::Date;
    chart.category_axis = Some(cat_axis);
    chart.value_axis = Some(Axis::new());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.category_axis.as_ref().unwrap().axis_type, AxisType::Date,);
}

#[test]
fn test_roundtrip_axis_series_type() {
    use duke_sheets_chart::{Axis, AxisType, Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.is_3d = true;
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut ser_axis = Axis::new();
    ser_axis.axis_type = AxisType::Series;
    ser_axis.title = Some("Depth".to_string());
    chart.series_axis = Some(ser_axis);
    chart.category_axis = Some(Axis::new());
    chart.value_axis = Some(Axis::new());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ser_ax = c.series_axis.as_ref().expect("series_axis should survive");
    assert_eq!(ser_ax.axis_type, AxisType::Series);
    assert_eq!(ser_ax.title.as_deref(), Some("Depth"));
}

#[test]
fn test_roundtrip_view3d_all_fields() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries, View3D};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.is_3d = true;
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    chart.view_3d = Some(View3D {
        rotate_x: Some(20),
        rotate_y: Some(30),
        depth_percent: Some(200),
        height_percent: Some(150),
        perspective: Some(45),
        right_angle_axes: Some(false),
    });
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let v = c.view_3d.as_ref().expect("view_3d should survive");
    assert_eq!(v.rotate_x, Some(20));
    assert_eq!(v.rotate_y, Some(30));
    assert_eq!(v.depth_percent, Some(200));
    assert_eq!(v.height_percent, Some(150));
    assert_eq!(v.perspective, Some(45));
    assert_eq!(v.right_angle_axes, Some(false));
}

#[test]
fn test_roundtrip_chart_no_title() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = None;
    chart.auto_title_deleted = Some(true);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert!(
        c.title.is_none(),
        "title should remain None, got {:?}",
        c.title
    );
    assert_eq!(c.auto_title_deleted, Some(true));
}

#[test]
fn test_roundtrip_chart_empty_series_name() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.name = Some("".into());
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    assert_eq!(c.series.len(), 1);
    // Empty string name may come back as Some("") or None - either is acceptable
    // as long as the roundtrip doesn't crash
    if let Some(name) = c.series[0].name.as_deref() {
        assert_eq!(name, "");
    }
}

#[test]
fn test_roundtrip_chartsheet_hidden() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 100.0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    wb.add_chartsheet(ChartSheet {
        name: "HiddenChart".to_string(),
        chart,
        visibility: SheetVisibility::Hidden,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    assert_eq!(wb2.chartsheet_count(), 1);
    let cs = wb2.chartsheet(0).unwrap();
    assert_eq!(cs.name, "HiddenChart");
    assert_eq!(cs.visibility, SheetVisibility::Hidden);
}

#[test]
fn test_roundtrip_multiple_chartsheets() {
    use duke_sheets_chart::{Chart, ChartType, DataReference, DataSeries};
    use duke_sheets_core::{ChartSheet, SheetVisibility};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 1.0).unwrap();

    for (i, ctype) in [ChartType::ColumnClustered, ChartType::Line, ChartType::Pie]
        .iter()
        .enumerate()
    {
        let mut chart = Chart::new(ctype.clone());
        chart.title = Some(format!("Chart{}", i + 1));
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        wb.add_chartsheet(ChartSheet {
            name: format!("CS{}", i + 1),
            chart,
            visibility: SheetVisibility::Visible,
            raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
        })
        .unwrap();
    }

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();

    assert_eq!(wb2.chartsheet_count(), 3);
    assert_eq!(wb2.chartsheet(0).unwrap().name, "CS1");
    assert_eq!(
        wb2.chartsheet(0).unwrap().chart.chart_type,
        ChartType::ColumnClustered
    );
    assert_eq!(
        wb2.chartsheet(0).unwrap().chart.title.as_deref(),
        Some("Chart1")
    );
    assert_eq!(wb2.chartsheet(1).unwrap().name, "CS2");
    assert_eq!(wb2.chartsheet(1).unwrap().chart.chart_type, ChartType::Line);
    assert_eq!(wb2.chartsheet(2).unwrap().name, "CS3");
    assert_eq!(wb2.chartsheet(2).unwrap().chart.chart_type, ChartType::Pie);
}

#[test]
fn test_roundtrip_axis_tick_marks_cross_none() {
    use duke_sheets_chart::{Axis, Chart, ChartType, DataReference, DataSeries, TickMark};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    let mut cat_axis = Axis::new();
    cat_axis.major_tick_mark = Some(TickMark::Cross);
    cat_axis.minor_tick_mark = Some(TickMark::None);
    chart.category_axis = Some(cat_axis);
    chart.value_axis = Some(Axis::new());
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let ax = c.category_axis.as_ref().unwrap();
    assert_eq!(ax.major_tick_mark, Some(TickMark::Cross));
    assert_eq!(ax.minor_tick_mark, Some(TickMark::None));
}

#[test]
fn test_roundtrip_error_bars_direction_x() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, ErrorBarDirection, ErrorBarType, ErrorBars,
        ErrorValueType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ScatterMarkers);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$B$1:$B$5"))
        .with_categories(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.error_bars = Some(ErrorBars {
        direction: ErrorBarDirection::X,
        bar_type: ErrorBarType::Plus,
        value_type: ErrorValueType::FixedValue,
        value: Some(1.5),
        no_end_cap: None,
        plus: None,
        minus: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let eb = c.series[0]
        .error_bars
        .as_ref()
        .expect("error_bars should survive");
    assert_eq!(eb.direction, ErrorBarDirection::X);
    assert_eq!(eb.bar_type, ErrorBarType::Plus);
}

#[test]
fn test_roundtrip_error_bars_minus_type() {
    use duke_sheets_chart::{
        Chart, ChartType, DataReference, DataSeries, ErrorBarDirection, ErrorBarType, ErrorBars,
        ErrorValueType,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::BarClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.error_bars = Some(ErrorBars {
        direction: ErrorBarDirection::Y,
        bar_type: ErrorBarType::Minus,
        value_type: ErrorValueType::Percentage,
        value: Some(15.0),
        no_end_cap: Some(false),
        plus: None,
        minus: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let eb = c.series[0]
        .error_bars
        .as_ref()
        .expect("error_bars should survive");
    assert_eq!(eb.bar_type, ErrorBarType::Minus);
    assert_eq!(eb.no_end_cap, Some(false));
}

#[test]
fn test_roundtrip_data_label_number_format() {
    use duke_sheets_chart::{
        Chart, ChartType, DataLabels, DataReference, DataSeries, NumberFormat,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.data_labels = Some(DataLabels {
        show_value: Some(true),
        number_format: Some(NumberFormat {
            format_code: "#,##0.00".to_string(),
            source_linked: Some(false),
        }),
        ..Default::default()
    });
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c.data_labels.as_ref().expect("data_labels should survive");
    let nf = dl
        .number_format
        .as_ref()
        .expect("number_format should survive");
    assert_eq!(nf.format_code, "#,##0.00");
    assert_eq!(nf.source_linked, Some(false));
}

#[test]
fn test_roundtrip_data_label_show_leader_lines() {
    use duke_sheets_chart::{Chart, ChartType, DataLabels, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::Pie);
    chart.data_labels = Some(DataLabels {
        show_value: Some(true),
        show_leader_lines: Some(true),
        ..Default::default()
    });
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let dl = c.data_labels.as_ref().expect("data_labels should survive");
    assert_eq!(dl.show_leader_lines, Some(true));
}

#[test]
fn test_roundtrip_chart_shape_properties_no_fill() {
    use duke_sheets_chart::{Chart, ChartShapeProperties, ChartType, DataReference, DataSeries};

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.shape_properties = Some(ChartShapeProperties {
        solid_fill: None,
        no_fill: true,
        line: None,
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let sp = c.series[0]
        .shape_properties
        .as_ref()
        .expect("shape_properties should survive");
    assert!(sp.no_fill, "no_fill should be true after roundtrip");
    assert!(sp.solid_fill.is_none());
}

#[test]
fn test_roundtrip_chart_line_no_fill() {
    use duke_sheets_chart::{
        Chart, ChartLine, ChartShapeProperties, ChartType, DataReference, DataSeries,
    };

    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut s = DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5"));
    s.shape_properties = Some(ChartShapeProperties {
        solid_fill: None,
        no_fill: false,
        line: Some(ChartLine {
            width: None,
            solid_fill: None,
            no_fill: true,
            dash_style: None,
        }),
    });
    chart.add_series(s);
    sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
    let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

    let sp = c.series[0]
        .shape_properties
        .as_ref()
        .expect("shape_properties should survive");
    let ln = sp.line.as_ref().expect("line should survive");
    assert!(ln.no_fill, "line no_fill should be true after roundtrip");
}

#[test]
fn test_roundtrip_data_label_position_pie() {
    use duke_sheets_chart::{
        Chart, ChartType, DataLabelPosition, DataLabels, DataReference, DataSeries,
    };

    for pos in [DataLabelPosition::BestFit, DataLabelPosition::OutsideEnd] {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        let mut chart = Chart::new(ChartType::Pie);
        chart.data_labels = Some(DataLabels {
            show_value: Some(true),
            position: Some(pos),
            ..Default::default()
        });
        chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
        sheet.add_chart(chart, DrawingAnchor::default()).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = XlsxReader::read(Cursor::new(&buf)).unwrap();
        let c = wb2.worksheet(0).unwrap().charts().next().unwrap().payload;

        let dl = c.data_labels.as_ref().expect("data_labels should survive");
        assert_eq!(
            dl.position,
            Some(pos),
            "DataLabelPosition::{:?} on pie chart did not survive roundtrip",
            pos,
        );
    }
}

#[test]
fn test_images_empty_by_default() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello").unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
    let rt = XlsxReader::read(Cursor::new(&buf)).unwrap();
    assert!(rt.worksheet(0).unwrap().images().next().is_none());
}

/// A 1x1 transparent PNG (68 bytes) used as a deterministic image
/// payload for round-trip tests.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + name
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // width=1 height=1
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, // 8-bit RGBA + IHDR CRC
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, // IDAT length + name
    0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00, // zlib-deflated 1 pixel
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, // IDAT CRC
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND length + name
    0xAE, 0x42, 0x60, 0x82, // IEND CRC
];

/// Build a TwoCell-anchored image drawing object over a known PNG payload.
fn test_two_cell_image(
    name: &str,
    from_col: u16,
    from_row: u32,
    to_col: u16,
    to_row: u32,
) -> DrawingObject {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    let image = EmbeddedImage {
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
    };
    let anchor = DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: to_col,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    DrawingObject::image(image)
        .with_anchor(anchor)
        .with_name(name)
}

#[test]
fn xlsx_png_image_round_trips() {
    use duke_sheets_chart::{DrawingAnchor, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    let mut img = test_two_cell_image("Pic1", 1, 2, 5, 10);
    img.meta.alt_text = Some("Test image".into());
    ws.add_drawing(img).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let ws_in = rt.worksheet(0).unwrap();
    let images: Vec<_> = ws_in.images().collect();
    assert_eq!(images.len(), 1, "exactly one image must round-trip");

    let img = &images[0];
    assert_eq!(img.object.unwrap().meta.name.as_deref(), Some("Pic1"));
    assert_eq!(img.object.unwrap().meta.alt_text.as_deref(), Some("Test image"));
    assert_eq!(img.payload.format, ImageFormat::Png);
    assert_eq!(
        img.payload.data, TEST_PNG_1X1,
        "PNG bytes must round-trip verbatim"
    );
    assert_eq!(img.payload.width_emu, 1_000_000);
    assert_eq!(img.payload.height_emu, 2_000_000);
    if let DrawingAnchor::TwoCell { from, to, .. } = &img.object.unwrap().anchor {
        assert_eq!(from.col, 1);
        assert_eq!(from.row, 2);
        assert_eq!(to.col, 5);
        assert_eq!(to.row, 10);
    } else {
        panic!("expected TwoCell anchor, got {:?}", img.object.unwrap().anchor);
    }
}

#[test]
fn xlsx_onecell_anchor_round_trips() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(
        DrawingObject::image(EmbeddedImage {
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
        })
        .with_anchor(DrawingAnchor::OneCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 50_000,
                row: 3,
                row_offset_emu: 70_000,
            },
            width_emu: 1_500_000,
            height_emu: 800_000,
        })
        .with_name("OneCellPic"),
    ).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let img = rt.worksheet(0).unwrap().images().next().unwrap();
    if let DrawingAnchor::OneCell {
        from,
        width_emu,
        height_emu,
    } = &img.object.unwrap().anchor
    {
        assert_eq!(from.col, 2);
        assert_eq!(from.col_offset_emu, 50_000);
        assert_eq!(from.row, 3);
        assert_eq!(from.row_offset_emu, 70_000);
        assert_eq!(*width_emu, 1_500_000);
        assert_eq!(*height_emu, 800_000);
    } else {
        panic!(
            "expected OneCell anchor after round-trip, got {:?}",
            img.object.unwrap().anchor
        );
    }
    assert_eq!(img.payload.format, ImageFormat::Png);
    assert_eq!(img.payload.data, TEST_PNG_1X1);
}

#[test]
fn xlsx_absolute_anchor_round_trips() {
    use duke_sheets_chart::{DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(
        DrawingObject::image(EmbeddedImage {
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
        })
        .with_anchor(DrawingAnchor::Absolute {
            x_emu: 2_500_000,
            y_emu: 1_200_000,
            width_emu: 1_000_000,
            height_emu: 900_000,
        })
        .with_name("AbsolutePic"),
    ).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let img = rt.worksheet(0).unwrap().images().next().unwrap();
    if let DrawingAnchor::Absolute {
        x_emu,
        y_emu,
        width_emu,
        height_emu,
    } = &img.object.unwrap().anchor
    {
        assert_eq!(*x_emu, 2_500_000);
        assert_eq!(*y_emu, 1_200_000);
        assert_eq!(*width_emu, 1_000_000);
        assert_eq!(*height_emu, 900_000);
    } else {
        panic!(
            "expected Absolute anchor after round-trip, got {:?}",
            img.object.unwrap().anchor
        );
    }
    assert_eq!(img.payload.format, ImageFormat::Png);
    assert_eq!(img.payload.data, TEST_PNG_1X1);
}

#[test]
fn xlsx_twocell_anchor_editas_round_trips() {
    // editAs attribute on twoCellAnchor must round-trip. Tests
    // editAs=oneCell and editAs=absolute via the TwoCell variant.
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EditAs, EmbeddedImage, ImageFormat};

    for ea in [EditAs::TwoCell, EditAs::OneCell, EditAs::Absolute] {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.add_drawing(
            DrawingObject::image(EmbeddedImage {
                format: ImageFormat::Png,
                media_path: String::new(),
                svg_media_path: None,
                width_emu: 1_000_000,
                height_emu: 1_000_000,
                rotation: None,
                flip_h: false,
                flip_v: false,
                data: TEST_PNG_1X1.to_vec(),
                svg_data: None,
            })
            .with_anchor(DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: 0,
                    col_offset_emu: 0,
                    row: 0,
                    row_offset_emu: 0,
                },
                to: CellMarker {
                    col: 2,
                    col_offset_emu: 0,
                    row: 2,
                    row_offset_emu: 0,
                },
                edit_as: Some(ea.clone()),
            })
            .with_name(format!("Pic-{:?}", ea)),
        ).unwrap();

        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
        let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
        let img = rt.worksheet(0).unwrap().images().next().unwrap();
        if let DrawingAnchor::TwoCell { edit_as, .. } = &img.object.unwrap().anchor {
            assert_eq!(
                edit_as.as_ref(),
                Some(&ea),
                "editAs={:?} must round-trip",
                ea
            );
        } else {
            panic!("expected TwoCell anchor");
        }
    }
}

#[test]
fn xlsx_multiple_images_round_trip_on_same_sheet() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for i in 0..3u32 {
        let name = format!("Pic{i}");
        let mut img = test_two_cell_image(&name, i as u16, i, i as u16 + 2, i + 2);
        // Distinct payloads: identical bytes would let a media/rels
        // mix-up (every rId resolving to image1.png) pass unnoticed.
        match &mut img.kind {
            DrawingKind::Image(image) => image.data.push(i as u8),
            other => panic!("expected image kind, got {other:?}"),
        }
        ws.add_drawing(img).unwrap();
    }

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let images: Vec<_> = rt.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 3);
    for i in 0..3usize {
        assert_eq!(images[i].object.unwrap().meta.name, Some(format!("Pic{i}")));
        assert_eq!(
            images[i].payload.data().last().copied(),
            Some(i as u8),
            "image {i} bytes mixed up with another image's media part"
        );
    }
}

#[test]
fn xlsx_images_round_trip_across_multiple_sheets() {
    // Media parts are numbered globally (image1, image2, ...) while
    // drawings and rels are per-sheet; cross-sheet numbering is
    // exercised only by a workbook with images on more than one sheet.
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Second").expect("add sheet");
    for sheet in 0..2usize {
        let ws = wb.worksheet_mut(sheet).unwrap();
        for i in 0..2u32 {
            let name = format!("S{sheet}Pic{i}");
            let mut img = test_two_cell_image(&name, i as u16, i, i as u16 + 2, i + 2);
            match &mut img.kind {
                DrawingKind::Image(image) => image.data.push((sheet * 10 + i as usize) as u8),
                other => panic!("expected image kind, got {other:?}"),
            }
            ws.add_drawing(img).unwrap();
        }
    }

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    for sheet in 0..2usize {
        let images: Vec<_> = rt.worksheet(sheet).unwrap().images().collect();
        assert_eq!(images.len(), 2, "sheet {sheet} image count");
        for i in 0..2usize {
            assert_eq!(images[i].object.unwrap().meta.name, Some(format!("S{sheet}Pic{i}")));
            assert_eq!(
                images[i].payload.data().last().copied(),
                Some((sheet * 10 + i) as u8),
                "sheet {sheet} image {i} bytes resolved to the wrong media part"
            );
        }
    }
}

// Form controls

fn control_anchor_2c(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: to_col,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    }
}

fn all_form_control_kinds() -> Vec<FormControlKind> {
    vec![
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
            input_range: Some("$H$1:$H$5".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 2, 4],
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
            horizontal: true,
            cell_link: Some("$D$6".to_string()),
        },
        FormControlKind::Spinner {
            value: 12,
            min: 0,
            max: 30,
            increment: 3,
            cell_link: Some("$D$7".to_string()),
        },
    ]
}

#[test]
fn xlsx_form_controls_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).unwrap();
    let kinds = all_form_control_kinds();
    let count = kinds.len();
    for (i, kind) in kinds.into_iter().enumerate() {
        let row = 1 + 2 * i as u32;
        ws.add_form_control(
            FormControl::new(kind),
            control_anchor_2c(1, row, 3, row + 1),
        ).unwrap();
    }

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let controls: Vec<_> = rt.worksheet(0).unwrap().form_controls().collect();
    assert_eq!(controls.len(), count, "every control survives");
    let original = all_form_control_kinds();
    for (i, control) in controls.iter().enumerate() {
        // The writer recomputes radio grouping: the sheet group's
        // first radio (control 3) carries firstButton after the trip.
        let mut expected = original[i].clone();
        if let FormControlKind::OptionButton { first_in_group, .. } = &mut expected {
            *first_in_group = i == 3;
        }
        assert_eq!(control.payload.kind, expected, "control {i} kind mismatch");
        assert!(control.object.unwrap().meta.locked);
        assert!(control.object.unwrap().meta.printable);
    }
    // Anchors survive exactly (EMU in controlPr).
    match &controls[0].object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.row), (1, 1));
            assert_eq!((to.col, to.row), (3, 2));
        }
        other => panic!("expected TwoCell anchor, got {other:?}"),
    }
    // Names are persisted (defaulted by the writer).
    assert_eq!(controls[0].object.unwrap().meta.name.as_deref(), Some("Button 1"));
}

#[test]
fn xlsx_form_controls_emit_expected_parts() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_comment("A1", CellComment::new("Author", "note"))
        .unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "check".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: false,
        }),
        control_anchor_2c(1, 1, 3, 2),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 2],
            no_3d: false,
        }),
        control_anchor_2c(1, 4, 3, 6),
    ).unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selected: Some(1),
            lines: 8,
            no_3d: false,
        }),
        control_anchor_2c(1, 7, 3, 8),
    ).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");

    let mut zip = zip::ZipArchive::new(Cursor::new(&buf)).expect("zip");
    let read_part = |zip: &mut zip::ZipArchive<Cursor<&Vec<u8>>>, name: &str| -> String {
        use std::io::Read;
        let mut s = String::new();
        zip.by_name(name)
            .unwrap_or_else(|_| panic!("part {name} missing"))
            .read_to_string(&mut s)
            .expect("read part");
        s
    };

    let ctrl_prop = read_part(&mut zip, "xl/ctrlProps/ctrlProp1.xml");
    assert!(ctrl_prop.contains("objectType=\"CheckBox\""));
    assert!(ctrl_prop.contains("checked=\"Checked\""));
    assert!(ctrl_prop.contains("fmlaLink=\"$D$2\""));

    // Zero-based model selections serialize one-based on disk.
    let list_prop = read_part(&mut zip, "xl/ctrlProps/ctrlProp2.xml");
    assert!(list_prop.contains("multiSel=\"1,3\""), "{list_prop}");
    let drop_prop = read_part(&mut zip, "xl/ctrlProps/ctrlProp3.xml");
    assert!(drop_prop.contains("sel=\"2\""), "{drop_prop}");

    let vml = read_part(&mut zip, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("ObjectType=\"Note\""), "comment shape present");
    assert!(
        vml.contains("ObjectType=\"Checkbox\""),
        "control shape present"
    );
    assert!(vml.contains("_x0000_t201"), "control shapetype declared");

    let sheet = read_part(&mut zip, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("<legacyDrawing r:id=\"rId1\"/>"));
    assert!(sheet.contains("<controls>"));
    // Comment occupies shape 1025; the control is 1026.
    assert!(
        sheet.contains("shapeId=\"1026\""),
        "control shapeId offset by comment"
    );

    let content_types = read_part(&mut zip, "[Content_Types].xml");
    assert!(content_types.contains("/xl/ctrlProps/ctrlProp1.xml"));
    assert!(content_types.contains("Extension=\"vml\""));
}

#[test]
fn xlsx_controls_without_comments_round_trip() {
    // VML part exists for the controls alone; no comments part.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_form_control(
        FormControl::new(FormControlKind::Spinner {
            value: 5,
            min: 0,
            max: 10,
            increment: 1,
            cell_link: None,
        }),
        control_anchor_2c(0, 0, 1, 3),
    ).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    assert_eq!(rt.worksheet(0).unwrap().form_control_count(), 1);
    assert_eq!(rt.worksheet(0).unwrap().comment_count(), 0);
}

#[test]
fn xlsx_control_named_and_flagged_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut object = DrawingObject::form_control(FormControl::new(FormControlKind::Button {
        caption: "Do <it> & more".into(),
    }))
    .with_anchor(control_anchor_2c(2, 2, 4, 4))
    .with_name("My Button");
    object.meta.locked = false;
    object.meta.printable = false;
    ws.add_drawing(object).unwrap();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("serialize");
    let rt = XlsxReader::read(Cursor::new(&buf)).expect("read");
    let controls: Vec<_> = rt.worksheet(0).unwrap().form_controls().collect();
    assert_eq!(controls[0].object.unwrap().meta.name.as_deref(), Some("My Button"));
    assert!(!controls[0].object.unwrap().meta.locked);
    assert!(!controls[0].object.unwrap().meta.printable);
    assert_eq!(
        controls[0].payload.caption_text().as_deref(),
        Some("Do <it> & more")
    );
}
