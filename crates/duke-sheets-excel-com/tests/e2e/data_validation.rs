//! Tests for reading data validation from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_list_validation() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Pick color").expect("set value");
        // xlValidateList=3, xlValidAlertStop=1, operator=None (skipped for list)
        wb.add_data_validation("B1", 3, 1, None, "\"Red,Green,Blue\"", None)
            .expect("add validation");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let validations = sheet.data_validations();
    assert!(
        !validations.is_empty(),
        "Should have at least one data validation"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_whole_number_validation() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Enter 1-100").expect("set value");
        // xlValidateWholeNumber=1, xlValidAlertStop=1, xlBetween=1
        wb.add_data_validation("B1", 1, 1, Some(1), "1", Some("100"))
            .expect("add validation");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let validations = sheet.data_validations();
    assert!(
        !validations.is_empty(),
        "Should have at least one data validation"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_validation_with_messages() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "With messages").expect("set value");
        // xlValidateWholeNumber=1, xlValidAlertStop=1, xlGreater=5
        wb.add_data_validation("B1", 1, 1, Some(5), "0", None)
            .expect("add validation");
        wb.set_validation_input("B1", "Positive Numbers", "Please enter a positive integer")
            .expect("set input");
        wb.set_validation_error("B1", "Invalid Input", "Value must be greater than 0")
            .expect("set error");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let validations = sheet.data_validations();
    assert!(
        !validations.is_empty(),
        "Should have at least one data validation"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_text_length_validation() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Max 10 chars").expect("set value");
        // xlValidateTextLength=6, xlValidAlertWarning=2, xlLess=6
        wb.add_data_validation("B1", 6, 2, Some(6), "10", None)
            .expect("add validation");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let validations = sheet.data_validations();
    assert!(
        !validations.is_empty(),
        "Should have at least one data validation"
    );

    cleanup_fixture(&fixture);
}
