//! Round-trip tests for cross-sheet references in XLS formulas:
//! SUPBOOK + EXTERNSHEET globals + tRef3D / tArea3D ptg emission.

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

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
