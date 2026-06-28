//! Round-trip tests for cross-sheet references in XLS formulas:
//! SUPBOOK + EXTERNSHEET globals + tRef3D / tArea3D ptg emission.

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn cross_sheet_cell_ref_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Source").expect("rename");
    wb.add_worksheet_with_name("Dest").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 42.0)
        .expect("Source A1");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("B1", "=Source!A1")
        .expect("Dest B1 formula");
    wb.worksheet_mut(1)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(42.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let dest = parsed.worksheet_by_name("Dest").unwrap();
    let formula = dest
        .get_formula_at(0, 1)
        .expect("formula must round-trip via formula path");
    assert!(formula.contains("Source"), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
}

#[test]
fn cross_sheet_range_ref_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Data").expect("rename");
    wb.add_worksheet_with_name("Summary").expect("add");

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
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("B1", "=SUM(Data!A1:A3)")
        .expect("formula");
    wb.worksheet_mut(1)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(6.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let summary = parsed.worksheet_by_name("Summary").unwrap();
    let formula = summary
        .get_formula_at(0, 1)
        .expect("range ref must round-trip via formula path");
    assert!(formula.to_uppercase().contains("SUM"), "got {formula:?}");
    assert!(formula.contains("Data"), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
    assert!(formula.contains("A3"), "got {formula:?}");
}

#[test]
fn quoted_sheet_name_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "My Data").expect("rename");
    wb.add_worksheet_with_name("Out").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 99.0)
        .expect("A1");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("B1", "='My Data'!A1")
        .expect("formula");
    wb.worksheet_mut(1)
        .unwrap()
        .set_formula_result(0, 1, CellValue::Number(99.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let out = parsed.worksheet_by_name("Out").unwrap();
    let formula = out
        .get_formula_at(0, 1)
        .expect("quoted sheet name formula must round-trip");
    // The decompiler quotes sheet names that contain spaces; in the
    // raw round-trip we tolerate either the quoted or unquoted shape
    // as long as the sheet name and address are present.
    assert!(formula.contains("My Data"), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
}

#[test]
fn cross_sheet_absolute_ref_round_trips() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Constants").expect("rename");
    wb.add_worksheet_with_name("Calc").expect("add");

    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("B5", 7.0)
        .expect("B5");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "=Constants!$B$5")
        .expect("formula");
    wb.worksheet_mut(1)
        .unwrap()
        .set_formula_result(0, 0, CellValue::Number(7.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let calc = parsed.worksheet_by_name("Calc").unwrap();
    let formula = calc.get_formula_at(0, 0).expect("absolute 3D ref");
    assert!(formula.contains("Constants"), "got {formula:?}");
    assert!(
        formula.contains("$B$5") || formula.contains("B5"),
        "got {formula:?}"
    );
}

/// LibreOffice must evaluate cross-sheet formulas we emit. Exercises
/// the SUPBOOK/EXTERNSHEET global pair plus tRef3D and tArea3D ptg
/// shapes. The reader's tRef3D handler used to silently accept ptg
/// 0x5C (PTG_REF_ERR_3D) for V-class refs, which decompiled clean
/// but rendered as `#REF!` in any external tool — this test
/// reproduces the failure mode end-to-end.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_evaluate_cross_sheet_formulas_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    // Calc must be sheet 0 so the bridge's `get_cell_value("B1")` (which
    // is hardcoded to sheet 0) reads the formula cells we care about.
    // Data sits on sheet 1 and is referenced by the cross-sheet ptgs.
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").expect("rename");
    wb.add_worksheet_with_name("Data").expect("add");

    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .expect("Data!A1");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 20.0)
        .expect("Data!A2");
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 30.0)
        .expect("Data!A3");

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=Data!A1")
        .expect("3D cell ref");
    calc.set_formula_result(0, 1, CellValue::Number(10.0))
        .expect("cache B1");
    calc.set_cell_formula("B2", "=SUM(Data!A1:A3)")
        .expect("3D area sum");
    calc.set_formula_result(1, 1, CellValue::Number(60.0))
        .expect("cache B2");
    calc.set_cell_formula("B3", "=Data!$A$1+Data!$A$3")
        .expect("absolute mix");
    calc.set_formula_result(2, 1, CellValue::Number(40.0))
        .expect("cache B3");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_xsheet_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(f64, f64, f64), String> = rt.block_on(async {
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
            .map_err(|e| format!("Calc!B1: {e}"))?;
        let b2 = wb
            .get_cell_value("B2")
            .await
            .map_err(|e| format!("Calc!B2: {e}"))?;
        let b3 = wb
            .get_cell_value("B3")
            .await
            .map_err(|e| format!("Calc!B3: {e}"))?;
        Ok((b1, b2, b3))
    });
    let _ = std::fs::remove_file(&path);
    let (b1, b2, b3) = outcome.expect("LO must evaluate cross-sheet formulas");
    assert!((b1 - 10.0).abs() < 1e-9, "Calc!B1 = {b1} (expected 10)");
    assert!((b2 - 60.0).abs() < 1e-9, "Calc!B2 = {b2} (expected 60)");
    assert!((b3 - 40.0).abs() < 1e-9, "Calc!B3 = {b3} (expected 40)");
}
