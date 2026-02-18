//! Tests for reading data validation from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_list_validation() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Pick color").await.unwrap();
        wb.add_data_validation(
            0, "B1", "list", "",
            "\"Red,Green,Blue\"", "",
            true, true,
            Some("Colors"), Some("Pick a color"),
            None, None, "stop",
        )
        .await
        .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let validations = sheet.data_validations();
    assert!(!validations.is_empty(), "Should have at least one data validation");

    cleanup_fixture(&path);
}

#[test]
fn test_whole_number_validation() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Enter 1-100").await.unwrap();
        wb.add_data_validation(
            0, "B1", "whole", "between",
            "1", "100",
            true, false,
            None, None, None, None, "stop",
        )
        .await
        .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let validations = sheet.data_validations();
    assert!(!validations.is_empty(), "Should have at least one data validation");

    cleanup_fixture(&path);
}

#[test]
fn test_validation_with_messages() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "With messages").await.unwrap();
        wb.add_data_validation(
            0, "B1", "whole", "greaterThan",
            "0", "",
            false, false,
            Some("Positive Numbers"), Some("Please enter a positive integer"),
            Some("Invalid Input"), Some("Value must be greater than 0"),
            "stop",
        )
        .await
        .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let validations = sheet.data_validations();
    assert!(!validations.is_empty(), "Should have at least one data validation");

    cleanup_fixture(&path);
}

#[test]
fn test_text_length_validation() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Max 10 chars").await.unwrap();
        wb.add_data_validation(
            0, "B1", "textLength", "lessThan",
            "10", "",
            true, false,
            None, None, None, None, "warning",
        )
        .await
        .unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let validations = sheet.data_validations();
    assert!(!validations.is_empty(), "Should have at least one data validation");

    cleanup_fixture(&path);
}
