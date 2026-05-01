//! Round-trip tests for the XLS writer's PROTECT (0x0012) and
//! PASSWORD (0x0013) records.

use std::io::Cursor;

use duke_sheets_core::worksheet::SheetProtection;
use duke_sheets_core::Workbook;
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn unprotected_sheet_emits_no_protect_record() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    assert!(parsed.worksheet(0).unwrap().protection().is_none());
}

#[test]
fn protected_sheet_round_trips_flag() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_protection(Some(SheetProtection {
        protected: true,
        ..Default::default()
    }));

    let parsed = write_then_read(&wb);
    let protection = parsed
        .worksheet(0)
        .unwrap()
        .protection()
        .expect("protection present after round-trip")
        .clone();
    assert!(protection.protected);
}

#[test]
fn protected_sheet_with_password_hash_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(0xCAFE),
        ..Default::default()
    }));

    let parsed = write_then_read(&wb);
    let protection = parsed
        .worksheet(0)
        .unwrap()
        .protection()
        .expect("protection present")
        .clone();
    assert!(protection.protected);
    assert_eq!(protection.password_hash, Some(0xCAFE));
}

#[test]
fn protection_persists_per_sheet() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Locked").expect("rename");
    wb.add_worksheet_with_name("Open").expect("add");
    wb.worksheet_mut(0)
        .unwrap()
        .set_protection(Some(SheetProtection {
            protected: true,
            ..Default::default()
        }));

    let parsed = write_then_read(&wb);
    assert!(parsed
        .worksheet_by_name("Locked")
        .unwrap()
        .protection()
        .map(|p| p.protected)
        .unwrap_or(false));
    assert!(parsed
        .worksheet_by_name("Open")
        .unwrap()
        .protection()
        .is_none());
}
