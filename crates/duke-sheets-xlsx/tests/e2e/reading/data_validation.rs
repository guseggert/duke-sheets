//! Tests for reading data validation from PyUNO fixtures.
//!
//! Fixture: `data_validation.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_data_validation_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_validation.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open data_validation.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_data_validations_present() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_validation.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let validations = sheet.data_validations();
    assert!(
        !validations.is_empty(),
        "Should have at least one data validation"
    );
}

#[test]
fn test_list_validation() {
    skip_if_no_fixtures!();

    let path = fixture_path("data_validation.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let validations = sheet.data_validations();

    let has_list = validations.iter().any(|v| {
        matches!(
            &v.validation_type,
            duke_sheets_core::ValidationType::List { .. }
        )
    });

    assert!(has_list, "Should have at least one list validation");
}
