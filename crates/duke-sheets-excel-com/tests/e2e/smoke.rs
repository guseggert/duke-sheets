//! Smoke test: full round-trip through real Excel.
//!
//! Creates a workbook via Excel COM, populates cells with various data types
//! and formulas, saves as XLSX inside the VM, pulls the file to the host
//! via WinRM, reads it back with duke-sheets, and asserts everything matches.

use crate::{cleanup_fixture, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

/// Full round-trip smoke test: numbers, strings, booleans, formulas.
///
/// 1. Connect to real Excel via COM bridge
/// 2. Create workbook, set cells with various data types
/// 3. Verify Excel computed the right formula results
/// 4. Save as XLSX to C:\temp inside the VM
/// 5. Pull the file to the Linux host via WinRM
/// 6. Read back with duke-sheets XlsxReader
/// 7. Assert all values match
#[test]
fn smoke_test_round_trip() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    // -- Phase 1: Build the spreadsheet in Excel --
    {
        let excel = bridge.lock().unwrap();

        // Ensure C:\temp exists inside the VM
        crate::ensure_vm_temp_dir();

        let wb = excel.create_workbook().expect("create workbook");

        // Numbers
        wb.set_cell_value("A1", 42.0).expect("set A1");
        wb.set_cell_value("A2", 3.14159).expect("set A2");
        wb.set_cell_value("A3", -100.0).expect("set A3");
        wb.set_cell_value("A4", 0.0).expect("set A4");

        // Strings
        wb.set_cell_value("B1", "Hello").expect("set B1");
        wb.set_cell_value("B2", "World with spaces")
            .expect("set B2");
        wb.set_cell_value("B3", "").expect("set B3");

        // Booleans
        wb.set_cell_value("C1", true).expect("set C1");
        wb.set_cell_value("C2", false).expect("set C2");

        // Formulas
        wb.set_cell_formula("D1", "=A1+A2").expect("set D1");
        wb.set_cell_formula("D2", "=SUM(A1:A4)").expect("set D2");
        wb.set_cell_formula("D3", "=A1*2").expect("set D3");
        wb.set_cell_formula("D4", "=IF(C1,\"yes\",\"no\")")
            .expect("set D4");

        // Force recalculate
        excel.recalculate().expect("recalculate");

        // Verify Excel computed the right values before saving
        let d1 = wb.get_cell_value("D1").expect("get D1");
        assert_eq!(
            d1.as_f64().expect("D1 should be number"),
            42.0 + 3.14159,
            "D1 = A1+A2"
        );

        let d2 = wb.get_cell_value("D2").expect("get D2");
        assert!(
            (d2.as_f64().expect("D2 should be number") - (42.0 + 3.14159 - 100.0)).abs() < 0.001,
            "D2 = SUM(A1:A4)"
        );

        let d3 = wb.get_cell_value("D3").expect("get D3");
        assert_eq!(d3.as_f64().expect("D3 should be number"), 84.0, "D3 = A1*2");

        let d4 = wb.get_cell_value("D4").expect("get D4");
        assert_eq!(
            d4.as_str().expect("D4 should be string"),
            "yes",
            "D4 = IF(C1, yes, no)"
        );

        // Save inside the VM
        wb.save(&fixture.vm_path).expect("save workbook");
        wb.close().expect("close workbook");
    }

    // -- Phase 2: Pull file from VM and read back --
    pull_file_from_vm(&fixture);

    let workbook = XlsxReader::read_file(&fixture.host_path)
        .unwrap_or_else(|e| panic!("Failed to read {}: {e}", fixture.host_path.display()));
    let sheet = workbook.worksheet(0).expect("should have a worksheet");

    // Numbers
    assert_number(&sheet, 0, 0, 42.0, "A1");
    assert_number(&sheet, 1, 0, 3.14159, "A2");
    assert_number(&sheet, 2, 0, -100.0, "A3");
    assert_number(&sheet, 3, 0, 0.0, "A4");

    // Strings
    assert_string(&sheet, 0, 1, "Hello", "B1");
    assert_string(&sheet, 1, 1, "World with spaces", "B2");
    // B3 (empty string) — Excel may not write a cell for "" at all
    let b3 = sheet.cell_at(2, 1);
    assert!(
        b3.is_none()
            || matches!(
                &b3.unwrap().value,
                duke_sheets_core::CellValue::String(s) if s.as_ref().is_empty()
            )
            || matches!(&b3.unwrap().value, duke_sheets_core::CellValue::Empty),
        "B3 should be empty or absent"
    );

    // Booleans
    assert_bool(&sheet, 0, 2, true, "C1");
    assert_bool(&sheet, 1, 2, false, "C2");

    // Formulas — check the cached computed values
    assert_number(&sheet, 0, 3, 42.0 + 3.14159, "D1");
    assert_number(&sheet, 1, 3, 42.0 + 3.14159 - 100.0, "D2");
    assert_number(&sheet, 2, 3, 84.0, "D3");
    assert_formula_string(&sheet, 3, 3, "yes", "D4");

    cleanup_fixture(&fixture);
}

// ---- Assertion helpers ----

fn assert_number(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: f64,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!(
                (*n - expected).abs() < 0.001,
                "{label}: expected {expected}, got {n}"
            );
        }
        // Formulas store their cached value — could be Formula variant with cached Number
        duke_sheets_core::CellValue::Formula { cached_value, .. } => {
            if let Some(cached) = cached_value {
                match cached.as_ref() {
                    duke_sheets_core::CellValue::Number(n) => {
                        assert!(
                            (*n - expected).abs() < 0.001,
                            "{label}: expected {expected}, got {n} (cached)"
                        );
                    }
                    other => panic!("{label}: expected Number in formula cache, got {other:?}"),
                }
            } else {
                panic!("{label}: formula has no cached value");
            }
        }
        other => panic!("{label}: expected Number, got {other:?}"),
    }
}

fn assert_string(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: &str,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_ref(), expected, "{label}");
        }
        other => panic!("{label}: expected String, got {other:?}"),
    }
}

fn assert_bool(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: bool,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Boolean(b) => {
            assert_eq!(*b, expected, "{label}");
        }
        other => panic!("{label}: expected Boolean, got {other:?}"),
    }
}

fn assert_formula_string(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: &str,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Formula { cached_value, .. } => {
            if let Some(cached) = cached_value {
                match cached.as_ref() {
                    duke_sheets_core::CellValue::String(s) => {
                        assert_eq!(s.as_ref(), expected, "{label}");
                    }
                    other => panic!("{label}: expected String in formula cache, got {other:?}"),
                }
            } else {
                panic!("{label}: formula has no cached value");
            }
        }
        duke_sheets_core::CellValue::String(s) => {
            // Some formulas may be inlined as strings
            assert_eq!(s.as_ref(), expected, "{label}");
        }
        other => panic!("{label}: expected Formula or String, got {other:?}"),
    }
}
