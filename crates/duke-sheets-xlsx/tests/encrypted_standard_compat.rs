use std::net::TcpStream;
use std::path::PathBuf;
use std::time::Duration;

use duke_sheets_core::Workbook;
use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use duke_sheets_xlsx::{EncryptionProfile, XlsxWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";
const PASSWORD: &str = "standard-compat-pw";

fn lo_available() -> bool {
    TcpStream::connect_timeout(&"127.0.0.1:2002".parse().unwrap(), Duration::from_secs(2))
        .is_ok()
}

fn build_wb() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "standard compat round-trip").unwrap();
    ws.set_cell_value("B1", 4242.5).unwrap();
    wb
}

fn write_encrypted_to_shared(profile: &EncryptionProfile) -> PathBuf {
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");
    let pid = std::process::id();
    let path = PathBuf::from(format!("{SHARED_DIR}/duke_standard_compat_{pid}.xlsx"));
    let _ = std::fs::remove_file(&path);
    let wb = build_wb();
    XlsxWriter::write_file_encrypted(&wb, &path, PASSWORD, profile).expect("encrypted write");
    path
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn lo_can_read_standard_aes128_default_profile() {
    if !lo_available() {
        eprintln!("SKIP: LibreOffice URP not reachable on 127.0.0.1:2002");
        return;
    }
    let path = write_encrypted_to_shared(&EncryptionProfile::standard_default());
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let result = rt.block_on(async {
        let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook_with_password(path.to_str().unwrap(), PASSWORD)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        let b1 = wb
            .get_cell_value("B1")
            .await
            .map_err(|e| format!("read B1: {e}"))?;
        Ok::<(String, f64), String>((a1, b1))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b1) = result.expect("LO must open AES-128 Standard file we wrote");
    assert_eq!(a1, "standard compat round-trip");
    assert!((b1 - 4242.5).abs() < 1e-6);
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn lo_can_read_standard_aes256_profile() {
    if !lo_available() {
        eprintln!("SKIP: LibreOffice URP not reachable on 127.0.0.1:2002");
        return;
    }
    let path = write_encrypted_to_shared(&EncryptionProfile::Standard { key_bits: 256 });
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let result = rt.block_on(async {
        let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook_with_password(path.to_str().unwrap(), PASSWORD)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        Ok::<String, String>(a1)
    });
    let _ = std::fs::remove_file(&path);
    let a1 = result.expect("LO must open AES-256 Standard file we wrote");
    assert_eq!(a1, "standard compat round-trip");
}
