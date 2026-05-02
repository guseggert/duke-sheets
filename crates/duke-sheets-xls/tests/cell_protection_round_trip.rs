//! Round-trip tests for the XLS writer's per-cell protection bits
//! (XF.fLocked and XF.fHidden, MS-XLS §2.4.353).

use std::io::Cursor;

use duke_sheets_core::style::{Protection, Style};
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn unlocked_cell_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    let mut style = Style::new();
    style.protection = Protection {
        locked: false,
        hidden: false,
    };
    ws.set_cell_style("A1", &style).expect("style");

    let parsed = write_then_read(&wb);
    let s = parsed
        .worksheet(0)
        .unwrap()
        .cell_style("A1")
        .expect("ok")
        .cloned()
        .unwrap_or_default();
    assert!(!s.protection.locked, "fLocked should round-trip false");
}

#[test]
fn formula_hidden_cell_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=1+1").expect("formula");
    ws.set_formula_result(0, 0, duke_sheets_core::CellValue::Number(2.0))
        .expect("cached");
    let mut style = Style::new();
    style.protection = Protection {
        locked: true,
        hidden: true,
    };
    ws.set_cell_style("A1", &style).expect("style");

    let parsed = write_then_read(&wb);
    let s = parsed
        .worksheet(0)
        .unwrap()
        .cell_style("A1")
        .expect("ok")
        .cloned()
        .unwrap_or_default();
    assert!(s.protection.locked);
    assert!(s.protection.hidden, "fHidden should round-trip");
}

#[test]
fn explicitly_locked_state_round_trips() {
    // `Style::default()` (and therefore `Style::new()`) sets
    // `protection.locked = false` because the derived Default for
    // Protection zeros both fields - the Excel-style "locked by
    // default" lives in `Protection::new()`, which the Style helper
    // chain doesn't go through. Use that explicitly here.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    let mut style = Style::new().bold(true);
    style.protection = Protection::new();
    ws.set_cell_style("A1", &style).expect("style");

    let parsed = write_then_read(&wb);
    let s = parsed
        .worksheet(0)
        .unwrap()
        .cell_style("A1")
        .expect("ok")
        .cloned()
        .unwrap_or_default();
    assert!(
        s.protection.locked,
        "explicit Protection::new() must round-trip"
    );
    assert!(!s.protection.hidden);
    assert!(
        s.font.bold,
        "bold must still round-trip alongside protection"
    );
}

#[test]
fn unlocked_and_hidden_combination_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=42").expect("formula");
    ws.set_formula_result(0, 0, duke_sheets_core::CellValue::Number(42.0))
        .expect("cached");
    let mut style = Style::new();
    style.protection = Protection {
        locked: false,
        hidden: true,
    };
    ws.set_cell_style("A1", &style).expect("style");

    let parsed = write_then_read(&wb);
    let s = parsed
        .worksheet(0)
        .unwrap()
        .cell_style("A1")
        .expect("ok")
        .cloned()
        .unwrap_or_default();
    assert!(!s.protection.locked);
    assert!(s.protection.hidden);
}
