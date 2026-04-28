//! Cross-tool compatibility test: real Excel must open OOXML
//! Standard-encryption files we produce.

use std::path::PathBuf;

use crate::{ensure_vm_temp_dir, excel_bridge, push_file_to_vm, TempFixture};
use duke_sheets_core::Workbook;
use duke_sheets_xlsx::{EncryptionProfile, XlsxWriter};

const PASSWORD: &str = "duke-excel-standard-pw";
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
    ws.set_cell_value("A1", "duke standard compat round-trip")
        .unwrap();
    ws.set_cell_value("B1", 1234.5).unwrap();
    wb
}

fn write_encrypted_to(path: &std::path::Path, profile: &EncryptionProfile) {
    let wb = build_wb();
    XlsxWriter::write_file_encrypted(&wb, path, PASSWORD, profile).expect("encrypted write");
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn excel_can_read_standard_aes128_default() {
    let fixture = fixture_path("duke_standard_aes128.xlsx");
    let _ = std::fs::remove_file(&fixture.host_path);
    write_encrypted_to(&fixture.host_path, &EncryptionProfile::standard_default());
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let wb = excel
        .open_workbook_with_password(&fixture.vm_path, PASSWORD)
        .expect("Excel must open AES-128 Standard file we wrote");

    let wb_name = wb.name().expect("workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel reported the file as 'Repaired': {wb_name}"
    );
    let read_only = wb.is_read_only().expect("read-only flag");
    assert!(!read_only);

    let a1 = wb.get_cell_value("A1").expect("read A1");
    let b1 = wb.get_cell_value("B1").expect("read B1");
    wb.close().expect("close");
    let _ = std::fs::remove_file(&fixture.host_path);

    assert_eq!(a1.as_str(), Some("duke standard compat round-trip"));
    let b1_num = b1.as_f64().expect("B1 numeric");
    assert!((b1_num - 1234.5).abs() < 1e-6);
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn excel_can_read_standard_aes256() {
    let fixture = fixture_path("duke_standard_aes256.xlsx");
    let _ = std::fs::remove_file(&fixture.host_path);
    write_encrypted_to(
        &fixture.host_path,
        &EncryptionProfile::Standard { key_bits: 256 },
    );
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let wb = excel
        .open_workbook_with_password(&fixture.vm_path, PASSWORD)
        .expect("Excel must open AES-256 Standard file we wrote");

    let wb_name = wb.name().expect("workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel reported the file as 'Repaired': {wb_name}"
    );

    let a1 = wb.get_cell_value("A1").expect("read A1");
    wb.close().expect("close");
    let _ = std::fs::remove_file(&fixture.host_path);
    assert_eq!(a1.as_str(), Some("duke standard compat round-trip"));
}
