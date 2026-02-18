//! Common utilities for XLS E2E tests.
//!
//! Mirrors the XLSX E2E common module — reuses the same LO bridge singleton.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::OnceLock;

use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use tokio::runtime::Runtime;
use tokio::sync::Mutex;

const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

static RUNTIME: OnceLock<Runtime> = OnceLock::new();
static LO_BRIDGE: OnceLock<Mutex<LibreOfficeBridge>> = OnceLock::new();
static FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

pub fn lo_available() -> bool {
    std::net::TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        std::time::Duration::from_secs(2),
    )
    .is_ok()
}

pub fn runtime() -> &'static Runtime {
    RUNTIME.get_or_init(|| {
        tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .expect("Failed to create tokio runtime")
    })
}

pub async fn lo_bridge() -> Option<&'static Mutex<LibreOfficeBridge>> {
    if !lo_available() {
        return None;
    }
    let _ = std::fs::create_dir_all(SHARED_DIR);
    if LO_BRIDGE.get().is_none() {
        let bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .expect("Failed to connect to LibreOffice on localhost:2002");
        let _ = LO_BRIDGE.set(Mutex::new(bridge));
    }
    LO_BRIDGE.get()
}

pub fn temp_fixture_path() -> PathBuf {
    let n = FILE_COUNTER.fetch_add(1, Ordering::SeqCst);
    let pid = std::process::id();
    PathBuf::from(format!("{SHARED_DIR}/test_xls_{pid}_{n}.xls"))
}

pub fn cleanup_fixture(path: &PathBuf) {
    let _ = std::fs::remove_file(path);
}

#[macro_export]
macro_rules! skip_if_no_lo {
    () => {
        if !$crate::lo_available() {
            eprintln!("SKIP: LibreOffice not available on localhost:2002");
            return;
        }
    };
}
