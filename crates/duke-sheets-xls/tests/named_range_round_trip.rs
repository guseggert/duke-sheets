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
