//! Common utilities for E2E tests.

use std::path::{Path, PathBuf};
use std::process::Command;

/// Get the path to a PyUNO-generated fixture file.
///
/// This looks for fixtures in `tests/fixtures/pyuno/output/`.
/// If fixtures haven't been generated yet, the tests using this will be skipped.
///
/// # Example
///
/// ```rust,ignore
/// let path = fixture_path("data_types.xlsx");
/// // Returns: tests/fixtures/pyuno/output/data_types.xlsx
/// ```
pub fn fixture_path(filename: &str) -> PathBuf {
    // Navigate from the crate directory to the workspace root, then to fixtures
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    let workspace_root = Path::new(manifest_dir)
        .parent()
        .expect("crate should be in crates/")
        .parent()
        .expect("crates/ should be in workspace root");

    workspace_root
        .join("tests")
        .join("fixtures")
        .join("pyuno")
        .join("output")
        .join(filename)
}

/// Check if PyUNO fixtures have been generated.
///
/// Returns true if the output directory exists and contains files.
pub fn fixtures_available() -> bool {
    let output_dir = fixture_path("");
    output_dir.exists()
        && output_dir
            .read_dir()
            .map(|d| d.count() > 0)
            .unwrap_or(false)
}

/// Verify an XLSX file using PyUNO (requires Docker).
///
/// This function runs the PyUNO verifier in a Docker container to verify
/// that an XLSX file written by Rust is valid and contains the expected content.
///
/// # Arguments
///
/// * `xlsx_path` - Path to the XLSX file to verify
/// * `spec_json` - JSON string containing verification assertions
///
/// # Returns
///
/// * `Ok(())` if verification passes
/// * `Err(String)` with error message if verification fails
///
/// # Example
///
/// ```rust,ignore
/// let spec = r#"{
///     "cells": {
///         "A1": {"value": "Hello", "type": "string"}
///     }
/// }"#;
///
/// verify_with_pyuno("/tmp/test.xlsx", spec)?;
/// ```
#[allow(dead_code)]
pub fn verify_with_pyuno(xlsx_path: &Path, spec_json: &str) -> Result<(), String> {
    // Check if Docker is available
    let docker_check = Command::new("docker")
        .args(["info"])
        .output()
        .map_err(|e| format!("Docker not available: {}", e))?;

    if !docker_check.status.success() {
        return Err("Docker daemon not running".to_string());
    }

    // Create a temp file for the spec
    let spec_path = std::env::temp_dir().join("verify_spec.json");
    std::fs::write(&spec_path, spec_json)
        .map_err(|e| format!("Failed to write spec file: {}", e))?;

    // Get absolute paths
    let xlsx_abs = xlsx_path
        .canonicalize()
        .map_err(|e| format!("XLSX file not found: {}", e))?;
    let spec_abs = spec_path
        .canonicalize()
        .map_err(|e| format!("Spec file not found: {}", e))?;

    // Run verification in Docker
    let output = Command::new("docker")
        .args([
            "run",
            "--rm",
            "-v",
            &format!("{}:/input/file.xlsx:ro", xlsx_abs.display()),
            "-v",
            &format!("{}:/input/spec.json:ro", spec_abs.display()),
            "duke-sheets-pyuno",
            "/app/run.sh",
            "--verify",
            "/input/file.xlsx",
            "/input/spec.json",
        ])
        .output()
        .map_err(|e| format!("Failed to run Docker: {}", e))?;

    // Clean up spec file
    let _ = std::fs::remove_file(&spec_path);

    if output.status.success() {
        Ok(())
    } else {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        Err(format!(
            "Verification failed:\nstdout: {}\nstderr: {}",
            stdout, stderr
        ))
    }
}

/// Macro to skip a test if fixtures are not available.
///
/// This is useful for tests that depend on PyUNO-generated fixtures,
/// which may not exist if Docker hasn't been run.
///
/// # Example
///
/// ```rust,ignore
/// #[test]
/// fn test_data_types() {
///     skip_if_no_fixtures!();
///     let path = fixture_path("data_types.xlsx");
///     // ... rest of test
/// }
/// ```
#[macro_export]
macro_rules! skip_if_no_fixtures {
    () => {
        if !$crate::fixtures_available() {
            eprintln!(
                "SKIP: PyUNO fixtures not available. Run `mise run fixtures:generate` to generate them."
            );
            return;
        }
    };
}

/// Macro to mark a test as requiring Docker for write verification.
///
/// # Example
///
/// ```rust,ignore
/// #[test]
/// #[ignore = "Requires Docker for PyUNO verification"]
/// fn test_write_basic() {
///     requires_docker!();
///     // ... rest of test
/// }
/// ```
#[macro_export]
macro_rules! requires_docker {
    () => {
        if std::process::Command::new("docker")
            .args(["info"])
            .output()
            .map(|o| !o.status.success())
            .unwrap_or(true)
        {
            eprintln!("SKIP: Docker not available for PyUNO verification");
            return;
        }
    };
}
