//! Round-trip tests for the XLS writer's PROTECT (0x0012) and
//! PASSWORD (0x0013) records.

use std::io::Cursor;

use duke_sheets_core::worksheet::SheetProtection;
use duke_sheets_core::{
    hash_legacy_protection_password, CellRange, ProtectedRange, Workbook, WorkbookProtection,
};
use duke_sheets_xls::{XlsReader, XlsWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

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
    ws.set_protection(Some(SheetProtection::protected().with_password("password")));

    let parsed = write_then_read(&wb);
    let protection = parsed
        .worksheet(0)
        .unwrap()
        .protection()
        .expect("protection present")
        .clone();
    assert!(protection.protected);
    assert_eq!(
        protection.password_hash,
        Some(hash_legacy_protection_password("password"))
    );
}

#[test]
fn protected_sheet_with_raw_password_hash_round_trips() {
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
    assert_eq!(protection.password_hash, Some(0xCAFE));
}

#[test]
fn workbook_protection_round_trips() {
    let mut wb = Workbook::new();
    wb.set_workbook_protection(Some(WorkbookProtection {
        structure: true,
        windows: true,
        password_hash: Some(0xCAFE),
    }));
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", 1.0)
        .expect("A1");

    let parsed = write_then_read(&wb);
    let protection = parsed
        .workbook_protection()
        .expect("workbook protection present");
    assert!(protection.structure);
    assert!(protection.windows);
    assert_eq!(protection.password_hash, Some(0xCAFE));
}

#[test]
fn protected_ranges_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "editable").expect("A1");
    ws.set_protected_ranges(vec![
        ProtectedRange::new(
            "MainEdit",
            vec![
                CellRange::parse("A1:B2").unwrap(),
                CellRange::parse("D4:D5").unwrap(),
            ],
        )
        .with_password("range"),
        ProtectedRange {
            name: "RawHash".to_string(),
            ranges: vec![CellRange::parse("F1:F3").unwrap()],
            password_hash: Some(0xCAFE),
            security_descriptor: None,
        },
    ]);

    let parsed = write_then_read(&wb);
    let ranges = parsed.worksheet(0).unwrap().protected_ranges();
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0].name, "MainEdit");
    assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
    assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
    assert_eq!(
        ranges[0].password_hash,
        Some(hash_legacy_protection_password("range"))
    );
    assert_eq!(ranges[1].name, "RawHash");
    assert_eq!(ranges[1].ranges[0].to_string(), "F1:F3");
    assert_eq!(ranges[1].password_hash, Some(0xCAFE));
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

/// LibreOffice must accept PROTECT + PASSWORD records.
#[test]
fn lo_can_read_sheet_protection_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "locked").expect("A1");
    ws.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(0xCAFE),
        ..Default::default()
    }));

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_protect_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<String, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb.get_cell_string("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    assert_eq!(outcome.expect("LO must open protected workbook"), "locked");
}
