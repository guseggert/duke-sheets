//! Common utilities for E2E tests.
//!
//! Provides a global LibreOffice connection singleton (mutex-protected) and a
//! shared tokio runtime. Each test creates fixtures on-demand: acquires the
//! lock, builds the spreadsheet it needs, saves to a temp file under the shared
//! Docker volume, reads it back with `XlsxReader`, asserts, and cleans up.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::OnceLock;

use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use tokio::runtime::Runtime;
use tokio::sync::Mutex;

/// Shared Docker volume path — accessible from both the host and the
/// LibreOffice Docker container.
const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

/// A single tokio runtime shared across all tests. This ensures the TCP
/// connection to LibreOffice outlives any individual test.
static RUNTIME: OnceLock<Runtime> = OnceLock::new();

/// Global LO bridge, initialized once and shared across all tests.
/// Protected by a tokio Mutex so only one test talks to LO at a time.
static LO_BRIDGE: OnceLock<Mutex<LibreOfficeBridge>> = OnceLock::new();

/// Counter for generating unique temp file names.
static FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

/// Check if LibreOffice URP is reachable on localhost:2002.
pub fn lo_available() -> bool {
    std::net::TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        std::time::Duration::from_secs(2),
    )
    .is_ok()
}

/// Get the shared tokio runtime.
pub fn runtime() -> &'static Runtime {
    RUNTIME.get_or_init(|| {
        tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .expect("Failed to create tokio runtime")
    })
}

/// Get the global LibreOffice bridge, connecting on first call.
///
/// Returns `None` if LibreOffice is not available (test should skip).
/// Must be called from within the shared `runtime()`.
pub async fn lo_bridge() -> Option<&'static Mutex<LibreOfficeBridge>> {
    if !lo_available() {
        return None;
    }

    // Ensure the shared directory exists
    let _ = std::fs::create_dir_all(SHARED_DIR);

    // Initialize if needed
    if LO_BRIDGE.get().is_none() {
        let bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .expect("Failed to connect to LibreOffice on localhost:2002");
        let _ = LO_BRIDGE.set(Mutex::new(bridge));
    }

    LO_BRIDGE.get()
}

/// Generate a unique temp file path under the shared Docker volume.
///
/// The file is placed in `/tmp/duke-sheets-urp/` so that LibreOffice (inside
/// Docker) can write to it and the host test process can read it back.
pub fn temp_fixture_path() -> PathBuf {
    let n = FILE_COUNTER.fetch_add(1, Ordering::SeqCst);
    let pid = std::process::id();
    PathBuf::from(format!("{SHARED_DIR}/test_{pid}_{n}.xlsx"))
}

/// Clean up a temp fixture file. Ignores errors.
pub fn cleanup_fixture(path: &PathBuf) {
    let _ = std::fs::remove_file(path);
}

/// Skip the test if LibreOffice is not available.
#[macro_export]
macro_rules! skip_if_no_lo {
    () => {
        if !$crate::lo_available() {
            eprintln!(
                "SKIP: LibreOffice not available on localhost:2002. \
                 Start with: docker run --rm -d -p 2002:2002 -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp duke-sheets-pyuno \
                 bash -c 'soffice --headless --accept=\"socket,host=0.0.0.0,port=2002;urp;StarOffice.ComponentContext\" & sleep infinity'"
            );
            return;
        }
    };
}
