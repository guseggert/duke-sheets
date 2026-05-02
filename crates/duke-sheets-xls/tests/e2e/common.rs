//! Common utilities for XLS E2E tests.
//!
//! The LibreOffice URP container is auto-started on first call to
//! `lo_bridge()` (or `skip_if_no_lo!()`) via
//! `duke_sheets_test_harness::lo::ensure_lo()`. There is no silent-skip
//! path: if the container can't be brought up the test panics.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::OnceLock;

use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use duke_sheets_test_harness::lo::SHARED_DIR;
use tokio::runtime::Runtime;
use tokio::sync::Mutex;

static RUNTIME: OnceLock<Runtime> = OnceLock::new();
static LO_BRIDGE: OnceLock<Mutex<LibreOfficeBridge>> = OnceLock::new();
static FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

/// Re-export of the auto-start primitive so the `skip_if_no_lo!()`
/// macro below can reach it via `$crate::ensure_lo()`.
pub fn ensure_lo() {
    duke_sheets_test_harness::lo::ensure_lo();
}

/// Vestigial port probe. Returns true once `ensure_lo()` has run.
/// Kept for backward compatibility with existing test code.
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

/// Get the global LibreOffice bridge, auto-starting the container on
/// first call. Panics if the container cannot be brought up.
///
/// Returns `Option` for backward compatibility with existing call sites
/// (`.unwrap()` / `.expect()`). The value is always `Some` once
/// `ensure_lo()` has run; failure modes panic from `ensure_lo()` itself.
pub async fn lo_bridge() -> Option<&'static Mutex<LibreOfficeBridge>> {
    ensure_lo();
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

/// Triggers `ensure_lo()`. Name retained for backward compatibility;
/// the macro no longer silently skips when LO is unavailable.
#[macro_export]
macro_rules! skip_if_no_lo {
    () => {
        $crate::ensure_lo();
    };
}
