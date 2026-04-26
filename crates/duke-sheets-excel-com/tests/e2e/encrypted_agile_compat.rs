//! Cross-tool compatibility test: real Excel must open files we encrypt
//! via [`duke_sheets_xlsx::XlsxWriter::write_to_bytes_encrypted`].
//!
//! The complementary "Excel writes, we read" coverage already lives in
//! `crypto_fixtures::generate_xlsx_agile_fixture_via_excel` plus the
//! `encrypted_agile` reader tests. This module validates the inverse
//! direction: bytes we produce open cleanly in Excel without "Repaired"
//! warnings or read-only fallback.

use std::path::PathBuf;

use crate::{ensure_vm_temp_dir, excel_bridge, push_file_to_vm, TempFixture};
use duke_sheets_core::Workbook;
use duke_sheets_xlsx::{EncryptionProfile, XlsxWriter};

const PASSWORD: &str = "duke-excel-compat-pw";
const HOST_DIR: &str = "/tmp/duke-sheets-excel";

fn fixture_path(name: &str) -> TempFixture {
    let _ = std::fs::create_dir_all(HOST_DIR);
    TempFixture {
        host_path: PathBuf::from(format!("{HOST_DIR}/{name}")),
        vm_path: format!(r"C:\temp\{name}"),
        name: name.to_string(),
    }
}

fn build_wb() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "duke compat round-trip").unwrap();
    ws.set_cell_value("B1", 9876.5).unwrap();
    wb
}

fn write_encrypted_to(path: &std::path::Path, profile: &EncryptionProfile) {
    let wb = build_wb();
    XlsxWriter::write_file_encrypted(&wb, path, PASSWORD, profile).expect("encrypted write");
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn excel_can_read_aes256_default_profile() {
    let fixture = fixture_path("duke_agile_aes256_default.xlsx");
    let _ = std::fs::remove_file(&fixture.host_path);
    write_encrypted_to(&fixture.host_path, &EncryptionProfile::agile_default());
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let wb = excel
        .open_workbook_with_password(&fixture.vm_path, PASSWORD)
        .expect("Excel must open AES-256 Agile file we wrote");

    let wb_name = wb.name().expect("workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel reported the file as 'Repaired': {wb_name}"
    );
    let read_only = wb.is_read_only().expect("read-only flag");
    assert!(!read_only, "Excel opened the file as read-only");

    let a1 = wb.get_cell_value("A1").expect("read A1");
    let b1 = wb.get_cell_value("B1").expect("read B1");

    wb.close().expect("close");
    let _ = std::fs::remove_file(&fixture.host_path);

    assert_eq!(
        a1.as_str(),
        Some("duke compat round-trip"),
        "Excel read of A1 must match"
    );
    let b1_num = b1.as_f64().expect("B1 numeric");
    assert!(
        (b1_num - 9876.5).abs() < 1e-6,
        "Excel read of B1 must match"
    );
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn excel_can_read_aes128_profile() {
    let fixture = fixture_path("duke_agile_aes128.xlsx");
    let _ = std::fs::remove_file(&fixture.host_path);
    write_encrypted_to(
        &fixture.host_path,
        &EncryptionProfile::Agile {
            key_bits: 128,
            spin_count: 100_000,
        },
    );
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let wb = excel
        .open_workbook_with_password(&fixture.vm_path, PASSWORD)
        .expect("Excel must open AES-128 Agile file we wrote");

    let wb_name = wb.name().expect("workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel reported the file as 'Repaired': {wb_name}"
    );

    let a1 = wb.get_cell_value("A1").expect("read A1");
    wb.close().expect("close");
    let _ = std::fs::remove_file(&fixture.host_path);

    assert_eq!(a1.as_str(), Some("duke compat round-trip"));
}
