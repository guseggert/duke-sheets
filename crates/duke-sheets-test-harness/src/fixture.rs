//! On-demand test fixture generation.
//!
//! `ensure_via_cargo_test()` is a small shim around `std::process::Command`
//! that spawns `cargo test ...` to produce a missing fixture file under
//! `data/`. The parity tests in `duke-sheets` use this so that the first
//! run on a new machine generates the prerequisite Excel COM fixture
//! transparently; subsequent runs see the file and skip the spawn.
//!
//! The mise task wrappers that used to encode this fall-through workflow
//! (`test:formula-parity`, `test:xlsb-parity`, `test:chart-parity`,
//! `generate:*-parity`, `test:corpus-charts`) have been deleted in favor
//! of this in-test logic. To run a parity test now:
//!
//! ```bash
//! mise run test:lo -- --test formula_parity -- --ignored
//! mise run test:lo -- --test xlsb_parity -- --ignored
//! mise run test:lo -- --test chart_parity -- --ignored
//! ```
//!
//! The first call generates whatever is missing; subsequent calls reuse
//! the cached file under `data/`.

use std::path::PathBuf;
use std::process::Command;

/// Ensure a fixture file exists at `repo_root/relative_path`. If it
/// doesn't, runs `cargo test <cargo_args>` to generate it; on success,
/// asserts that the file now exists and returns its path.
///
/// `cargo_args` should include the appropriate `-p`, `--test`, test name,
/// `--`, `--ignored`, `--nocapture` etc. The function adds nothing to
/// the args; it just runs them. `RUST_TEST_THREADS=1` is set in the
/// child env to match the rest of the harness.
///
/// Panics on spawn failure, on non-zero exit, or if the generation
/// completes but the file is still missing.
pub fn ensure_via_cargo_test(relative_path: &str, cargo_args: &[&str]) -> PathBuf {
    let path = repo_root().join(relative_path);
    if path.exists() {
        return path;
    }
    eprintln!(
        "[duke-sheets-test-harness] {} missing; generating via: cargo test {}",
        relative_path,
        cargo_args.join(" "),
    );
    let status = Command::new("cargo")
        .arg("test")
        .args(cargo_args)
        .env("RUST_TEST_THREADS", "1")
        .status()
        .unwrap_or_else(|e| panic!("failed to spawn `cargo test {}`: {e}", cargo_args.join(" "),));
    if !status.success() {
        panic!(
            "fixture generation failed: `cargo test {}` exited with {:?}",
            cargo_args.join(" "),
            status.code(),
        );
    }
    if !path.exists() {
        panic!(
            "fixture generation completed but {} is still missing",
            path.display(),
        );
    }
    path
}

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .map(PathBuf::from)
        .expect("CARGO_MANIFEST_DIR has no grandparent (expected crates/<name>)")
}
