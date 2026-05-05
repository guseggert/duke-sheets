//! Round-trip tests for the XLS writer's per-cell protection bits
//! (XF.fLocked and XF.fHidden, MS-XLS §2.4.353).

use std::io::Cursor;

use duke_sheets_core::style::{Protection, Style};
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

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

/// LibreOffice must accept XF.fLocked / fHidden bits.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_read_cell_protection_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "default").expect("A1");
    ws.set_cell_value("B1", "unlocked").expect("B1");
    ws.set_cell_formula("C1", "=2+2").expect("C1");
    ws.set_formula_result(0, 2, duke_sheets_core::CellValue::Number(4.0))
        .expect("cache C1");

    let mut unlocked = Style::new();
    unlocked.protection = Protection {
        locked: false,
        hidden: false,
    };
    ws.set_cell_style("B1", &unlocked).expect("style B1");

    let mut hidden = Style::new();
    hidden.protection = Protection {
        locked: true,
        hidden: true,
    };
    ws.set_cell_style("C1", &hidden).expect("style C1");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_cellprot_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<(String, String), String> = rt.block_on(async {
        let mut bridge = duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect(
            "127.0.0.1",
            2002,
        )
        .await
        .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("A1: {e}"))?;
        let b1 = wb
            .get_cell_string("B1")
            .await
            .map_err(|e| format!("B1: {e}"))?;
        Ok((a1, b1))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b1) = outcome.expect("LO must open cell-protection workbook");
    assert_eq!(a1, "default");
    assert_eq!(b1, "unlocked");
}
