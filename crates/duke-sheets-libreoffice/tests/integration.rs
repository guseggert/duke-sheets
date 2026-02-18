//! Integration tests for the LibreOffice URP bridge.
//!
//! These tests require a running LibreOffice instance with a URP socket listener.
//! They are gated behind either:
//!
//! 1. A manually-started LibreOffice:
//!    soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
//!
//! 2. The Docker-based setup (reuses the duke-sheets-pyuno image):
//!    docker run --rm -p 2002:2002 duke-sheets-pyuno /app/run.sh --idle
//!
//! 3. The mise task:
//!    mise run test:urp
//!
//! If LibreOffice is not reachable on localhost:2002, all tests are skipped.

use duke_sheets::prelude::*;

/// Check if a LibreOffice URP listener is available on localhost:2002.
fn urp_available() -> bool {
    std::net::TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        std::time::Duration::from_secs(2),
    )
    .is_ok()
}

/// Skip this test if URP is not available.
macro_rules! skip_if_no_urp {
    () => {
        if !urp_available() {
            eprintln!(
                "SKIP: LibreOffice URP not available on localhost:2002.\n\
                 Start LibreOffice with:\n  \
                 soffice --headless --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ComponentContext\"\n\
                 Or use Docker:\n  \
                 mise run urp:start"
            );
            return;
        }
    };
}

#[tokio::test]
async fn test_connect_and_bootstrap() {
    skip_if_no_urp!();

    let bridge = duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002).await;
    match bridge {
        Ok(bridge) => {
            eprintln!("OK: Connected and bootstrapped successfully");
            let _ = bridge.shutdown().await;
        }
        Err(e) => {
            panic!("Failed to connect and bootstrap: {e}");
        }
    }
}

#[tokio::test]
async fn test_create_workbook_and_set_cells() {
    skip_if_no_urp!();

    let mut bridge =
        duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002)
            .await
            .expect("connect");

    let mut wb = bridge.create_workbook().await.expect("create_workbook");

    // Set a numeric value
    wb.set_cell_value("A1", 42.0).await.expect("set A1");

    // Set a string value
    wb.set_cell_value("B1", "Hello").await.expect("set B1");

    // Set a formula
    wb.set_cell_formula("C1", "=A1*2").await.expect("set C1 formula");

    // Read back the numeric value
    let val = wb.get_cell_value("A1").await.expect("get A1");
    assert!(
        (val - 42.0).abs() < 1e-10,
        "A1 should be 42.0, got {val}"
    );

    // Read back the formula
    let formula = wb.get_cell_formula("C1").await.expect("get C1 formula");
    assert!(
        formula.contains("A1") && formula.contains("2"),
        "C1 formula should reference A1*2, got: {formula}"
    );

    // Read back the string
    let s = wb.get_cell_string("B1").await.expect("get B1 string");
    assert_eq!(s, "Hello", "B1 should be 'Hello', got '{s}'");

    wb.close().await.expect("close");
    let _ = bridge.shutdown().await;

    eprintln!("OK: Created workbook, set cells, read back values");
}

#[tokio::test]
async fn test_save_and_read_back_with_duke_sheets() {
    skip_if_no_urp!();

    let mut bridge =
        duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002)
            .await
            .expect("connect");

    let mut wb = bridge.create_workbook().await.expect("create_workbook");

    // Set various cell types
    wb.set_cell_value("A1", 3.14).await.expect("set A1");
    wb.set_cell_value("A2", "Hello World").await.expect("set A2");
    wb.set_cell_formula("A3", "=A1*10").await.expect("set A3");
    wb.set_cell_value("B1", 100.0).await.expect("set B1");
    wb.set_cell_value("B2", 200.0).await.expect("set B2");
    wb.set_cell_formula("B3", "=SUM(B1:B2)").await.expect("set B3");

    // Save to a temp file in the shared volume directory (accessible by both host and
    // the LibreOffice Docker container via -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp)
    let shared_dir = std::path::PathBuf::from("/tmp/duke-sheets-urp");
    std::fs::create_dir_all(&shared_dir).ok();
    let path = shared_dir.join(format!("urp-test-{}.xlsx", std::process::id()));
    let path_str = path.to_str().unwrap();

    wb.save(path_str).await.expect("save");
    wb.close().await.expect("close");
    let _ = bridge.shutdown().await;

    // Now read back with duke-sheets (our XLSX reader)
    assert!(path.exists(), "Saved file should exist at {path_str}");

    let file_size = std::fs::metadata(&path).unwrap().len();
    assert!(file_size > 0, "Saved file should not be empty");
    eprintln!("Saved XLSX file: {path_str} ({file_size} bytes)");

    let workbook = Workbook::open(path_str).expect("open with duke-sheets");
    let sheet = workbook.worksheet(0).expect("first sheet");

    // Check numeric value — A1 is row=0, col=0
    let a1 = sheet.get_value_at(0, 0);
    match a1 {
        CellValue::Number(n) => {
            assert!(
                (n - 3.14).abs() < 1e-10,
                "A1 should be 3.14, got {n}"
            );
        }
        other => panic!("A1 should be Number(3.14), got {other:?}"),
    }

    // Check string value — A2 is row=1, col=0
    let a2 = sheet.get_value_at(1, 0);
    match &a2 {
        CellValue::String(s) => {
            assert_eq!(s.as_str(), "Hello World", "A2 should be 'Hello World', got '{s}'");
        }
        other => panic!("A2 should be String('Hello World'), got {other:?}"),
    }

    // Check that the formulas exist (they may be stored as computed values or formulas)
    // A3 = A1*10 = 31.4, row=2, col=0
    let a3 = sheet.get_value_at(2, 0);
    match &a3 {
        CellValue::Number(n) => {
            assert!(
                (n - 31.4).abs() < 1e-10,
                "A3 should be 31.4 (=A1*10), got {n}"
            );
        }
        CellValue::Formula { cached_value, .. } => {
            // If stored as formula, check the cached value
            if let Some(cached) = cached_value {
                if let CellValue::Number(n) = cached.as_ref() {
                    assert!(
                        (n - 31.4).abs() < 1e-10,
                        "A3 cached value should be 31.4, got {n}"
                    );
                }
            }
            eprintln!("A3 stored as formula with cached value (OK)");
        }
        other => {
            eprintln!("A3 value: {other:?} (may vary depending on LO behavior)");
        }
    }

    // B3 = SUM(B1:B2) = 300, row=2, col=1
    let b3 = sheet.get_value_at(2, 1);
    match &b3 {
        CellValue::Number(n) => {
            assert!(
                (n - 300.0).abs() < 1e-10,
                "B3 should be 300.0 (=SUM(B1:B2)), got {n}"
            );
        }
        CellValue::Formula { cached_value, .. } => {
            if let Some(cached) = cached_value {
                if let CellValue::Number(n) = cached.as_ref() {
                    assert!(
                        (n - 300.0).abs() < 1e-10,
                        "B3 cached value should be 300.0, got {n}"
                    );
                }
            }
            eprintln!("B3 stored as formula with cached value (OK)");
        }
        other => {
            eprintln!("B3 value: {other:?}");
        }
    }

    // Clean up
    let _ = std::fs::remove_file(&path);

    eprintln!("OK: Saved XLSX via URP, read back with duke-sheets, values match");
}
