//! Round-trip tests for user-defined NAME records and tName ptg
//! emission in the XLS writer.
//!
//! Important caveat: the XLS reader currently does not populate
//! `Workbook::named_ranges()` from NAME records (it stores them in
//! its formula context for decompilation only). These tests therefore
//! verify the formula text round-trip, not that the named range
//! survives via `wb.named_ranges()`.

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn user_named_range_in_cell_formula_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Data").expect("rename");
    wb.add_worksheet_with_name("Out").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A2", 2.0)
        .expect("A2");
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A3", 3.0)
        .expect("A3");
    wb.define_name("MyData", "Data!$A$1:$A$3").expect("define");

    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("B1", "=SUM(MyData)")
        .expect("formula");
    wb.worksheet_mut(1)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(6.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet_by_name("Out")
        .unwrap()
        .get_formula_at(0, 1)
        .expect("formula must round-trip via formula path");
    assert!(formula.to_uppercase().contains("SUM"), "got {formula:?}");
    assert!(formula.contains("MyData"), "got {formula:?}");
}

#[test]
fn workbook_scoped_constant_name_round_trips_in_formula() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 100.0)
        .expect("A1");
    wb.define_name("TaxRate", "0.05").expect("define");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("B1", "=A1*TaxRate")
        .expect("formula");
    wb.worksheet_mut(0)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(5.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 1)
        .expect("formula must round-trip");
    assert!(formula.contains("TaxRate"), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
}

#[test]
fn sheet_scoped_name_round_trips_in_formula() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "First").expect("rename");
    wb.add_worksheet_with_name("Second").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .expect("A1");
    wb.define_name_for_sheet("Local", "First!$A$1", 0)
        .expect("define");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("B1", "=Local")
        .expect("formula");
    wb.worksheet_mut(0)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(10.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet_by_name("First")
        .unwrap()
        .get_formula_at(0, 1)
        .expect("formula must round-trip");
    assert!(formula.contains("Local"), "got {formula:?}");
}

#[test]
fn multiple_names_round_trip() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .expect("A1");
    wb.define_name("Alpha", "1").expect("alpha");
    wb.define_name("Beta", "2").expect("beta");
    wb.define_name("Gamma", "3").expect("gamma");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("B1", "=Alpha+Beta+Gamma")
        .expect("formula");
    wb.worksheet_mut(0)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(6.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let formula = parsed
        .worksheet(0)
        .unwrap()
        .get_formula_at(0, 1)
        .expect("formula must round-trip");
    assert!(formula.contains("Alpha"), "got {formula:?}");
    assert!(formula.contains("Beta"), "got {formula:?}");
    assert!(formula.contains("Gamma"), "got {formula:?}");
}

/// LibreOffice must evaluate user-defined NAME references via tName
/// ptg. The XLS reader doesn't repopulate `wb.named_ranges()` so
/// in-tree round-trip tests can only verify the formula decompiles
/// — this LO test verifies the formula actually computes the right
/// value when evaluated.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_evaluate_named_range_formulas_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    // Sheet 0 = "Calc" so the bridge's sheet-0-only get_cell_value
    // sees our formula cells.
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").expect("rename");
    wb.add_worksheet_with_name("Data").expect("add");

    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .expect("Data!A1");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 10.0)
        .expect("Data!A2");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 15.0)
        .expect("Data!A3");

    wb.define_name("Numbers", "Data!$A$1:$A$3").expect("name");
    wb.define_name("TaxRate", "0.1").expect("constant");

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").expect("B1");
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .expect("cache B1");
    calc.set_cell_formula("B2", "=B1*TaxRate").expect("B2");
    calc.set_formula_result(1, 1, CellValue::Number(3.0))
        .expect("cache B2");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_named_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(f64, f64), String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let b1 = wb
            .get_cell_value("B1")
            .await
            .map_err(|e| format!("B1: {e}"))?;
        let b2 = wb
            .get_cell_value("B2")
            .await
            .map_err(|e| format!("B2: {e}"))?;
        Ok((b1, b2))
    });
    let _ = std::fs::remove_file(&path);
    let (b1, b2) = outcome.expect("LO must evaluate named-range formulas");
    assert!((b1 - 30.0).abs() < 1e-9, "B1 = {b1} (expected 30)");
    assert!((b2 - 3.0).abs() < 1e-9, "B2 = {b2} (expected 3)");
}
