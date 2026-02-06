//! Tests for reading data types from PyUNO fixtures.
//!
//! Fixture: `data_types.xlsx`
//! - Numbers (integers, decimals, negative, scientific)
//! - Strings (ASCII, Unicode, escaped characters)
//! - Booleans (TRUE/FALSE)
//! - Formulas (with cached values)
//! - Errors (#DIV/0!, #VALUE!, #REF!, etc.)

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_data_types_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open data_types.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_number_values() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Test integer - B2 should have 42
    if let Some(cell) = sheet.cell_at(1, 1) {
        // B2
        match &cell.value {
            duke_sheets_core::CellValue::Number(n) => {
                assert!((*n - 42.0).abs() < 0.001, "Expected 42, got {}", n);
            }
            other => panic!("Expected Number, got {:?}", other),
        }
    }
}

#[test]
fn test_string_values() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for "Hello" string somewhere in the sheet
    let mut found_hello = false;
    for row in 0..20 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                if let duke_sheets_core::CellValue::String(s) = &cell.value {
                    if s.as_ref().contains("Hello") {
                        found_hello = true;
                    }
                }
            }
        }
    }
    assert!(
        found_hello,
        "Should find 'Hello' string somewhere in the sheet"
    );
}

#[test]
fn test_boolean_values() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for boolean values in the sheet
    let mut found_true = false;
    let mut found_false = false;

    for row in 0..30 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                if let duke_sheets_core::CellValue::Boolean(b) = &cell.value {
                    if *b {
                        found_true = true;
                    } else {
                        found_false = true;
                    }
                }
            }
        }
    }

    // At least one of TRUE/FALSE should be found
    assert!(
        found_true || found_false,
        "Should find at least one boolean value"
    );
}

#[test]
fn test_formula_values() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for cells with formulas
    let mut found_formula = false;

    for row in 0..30 {
        for col in 0..5 {
            if sheet.get_formula_at(row, col).is_some() {
                found_formula = true;
                break;
            }
        }
        if found_formula {
            break;
        }
    }

    assert!(found_formula, "Should find at least one formula");
}

#[test]
fn test_error_values() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_types.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Look for error values
    let mut found_error = false;

    for row in 0..40 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                if let duke_sheets_core::CellValue::Error(_) = &cell.value {
                    found_error = true;
                    break;
                }
            }
        }
        if found_error {
            break;
        }
    }

    assert!(found_error, "Should find at least one error value");
}
