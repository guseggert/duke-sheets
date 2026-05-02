//! Test harness for duke-sheets E2E tests.
//!
//! Auto-starts the LibreOffice URP container (`lo`) and the Excel COM
//! bridge VM (`excel`) on demand from test code. Idempotent and process
//! singleton, so the first call in a test process pays the startup cost
//! and subsequent calls are no-ops.
//!
//! Containers/VMs are deliberately **not** stopped on test exit; both are
//! expensive to boot and leaving them warm makes incremental test runs
//! cheap. Use `mise run vm:stop` and `docker rm -f duke-lo-test` (or
//! whichever container name was started) to tear down explicitly.
//!
//! All functions panic with a descriptive message on failure rather than
//! returning errors. Tests that need a backend should call `ensure_lo()`
//! / `ensure_excel_bridge()` from setup; if the backend can't be made
//! ready the test should fail loud, not silently skip.

pub mod excel;
pub mod lo;
