//! Tests for reading conditional formatting from PyUNO fixtures.
//!
//! Fixture: `conditional_format.xlsx`

use crate::{fixture_path, skip_if_no_fixtures};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_conditional_format_file_opens() {
    skip_if_no_fixtures!();

    let path = fixture_path("conditional_format.xlsx");
    let result = XlsxReader::read_file(&path);

    assert!(
        result.is_ok(),
        "Failed to open conditional_format.xlsx: {:?}",
        result.err()
    );
}

#[test]
fn test_conditional_formats_present() {
    skip_if_no_fixtures!();

    let path = fixture_path("conditional_format.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let cf_rules = sheet.conditional_formats();
    assert!(
        !cf_rules.is_empty(),
        "Should have at least one conditional format rule"
    );
}

#[test]
fn test_cell_is_condition() {
    skip_if_no_fixtures!();

    let path = fixture_path("conditional_format.xlsx");
    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let cf_rules = sheet.conditional_formats();

    let has_cell_is = cf_rules
        .iter()
        .any(|r| matches!(&r.rule_type, duke_sheets_core::CfRuleType::CellIs { .. }));

    assert!(
        has_cell_is,
        "Should have at least one CellIs conditional format"
    );
}
