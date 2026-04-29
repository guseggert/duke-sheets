//! Cross-tool compatibility tests for files we encrypt with Agile.
//!
//! The library's own reader can decrypt what the library writes (see
//! `encrypted_agile_write.rs`). These tests verify external readers
//! can also open the files — proof that we're emitting spec-compliant
//! Office output, not just a private serialization round-trip.
//!
//! All tests are `#[ignore]`-gated because they need a running
//! LibreOffice URP daemon on `127.0.0.1:2002`. The recommended way
//! to run them is via `mise run test:crypto-compat` which auto-starts
//! the LO container (and the Excel COM bridge) and tears them down.
//! Direct invocation is also supported once LO is up:
//!
//! ```sh
//! cargo test -p duke-sheets-xlsx --test encrypted_agile_compat -- \
//!   --ignored --nocapture
//! ```
//!
//! If LibreOffice isn't reachable when these tests are explicitly
//! invoked with `--ignored`, they panic (rather than silently passing)
//! so a missing backend can't be mistaken for a clean run.

use std::net::TcpStream;
use std::path::PathBuf;
use std::time::Duration;

use duke_sheets_core::Workbook;
use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use duke_sheets_xlsx::{EncryptionProfile, XlsxWriter};

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";
const PASSWORD: &str = "compat-test-pw";

fn require_lo() {
    if TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        Duration::from_secs(2),
    )
    .is_err()
    {
        panic!(
            "LibreOffice URP not reachable on 127.0.0.1:2002. \
             Start it with `mise run urp:start` or run the suite via \
             `mise run test:crypto-compat`."
        );
    }
}

fn build_wb() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "compat round-trip").unwrap();
    ws.set_cell_value("B1", 12345.5).unwrap();
    wb
}

/// Encrypt a workbook to a path under the shared Docker volume so LO
/// (running in a container) can read it back.
fn write_encrypted_to_shared(profile: &EncryptionProfile) -> PathBuf {
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");
    let pid = std::process::id();
    let path = PathBuf::from(format!("{SHARED_DIR}/duke_agile_compat_{pid}.xlsx"));
    let _ = std::fs::remove_file(&path);
    let wb = build_wb();
    XlsxWriter::write_file_encrypted(&wb, &path, PASSWORD, profile)
        .expect("encrypted write must succeed");
    path
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn lo_can_read_aes256_default_profile() {
    require_lo();
    let path = write_encrypted_to_shared(&EncryptionProfile::agile_default());

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

    let (a1, b1) = result.expect("LO must open AES-256 Agile file we wrote");
    assert_eq!(a1, "compat round-trip", "LO read of A1 must match");
    assert!((b1 - 12345.5).abs() < 1e-6, "LO read of B1 must match");
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn lo_can_read_aes128_profile() {
    require_lo();
    let path = write_encrypted_to_shared(&EncryptionProfile::Agile {
        key_bits: 128,
        spin_count: 100_000,
    });

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();

    let result = rt.block_on(async {
        let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002).await
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

    let a1 = result.expect("LO must open AES-128 Agile file we wrote");
    assert_eq!(a1, "compat round-trip");
}
