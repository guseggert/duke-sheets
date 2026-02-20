//! E2E tests for BIFF8 formula token decompilation.
//!
//! Each test creates a workbook in LibreOffice with specific formulas,
//! saves as .xls, then reads back and verifies the decompiled formula text.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::CellValue;
use duke_sheets_xls::XlsReader;

/// Helper to extract formula text from a cell value.
fn formula_text(val: &CellValue) -> &str {
    match val {
        CellValue::Formula { text, .. } => text.as_str(),
        other => panic!("Expected Formula, got {:?}", other),
    }
}

#[test]
fn test_xls_formula_simple_arithmetic() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("B1", 20.0).await.unwrap();
        // C1 = A1+B1
        wb.set_cell_formula("C1", "=A1+B1").await.unwrap();
        // D1 = A1*B1
        wb.set_cell_formula("D1", "=A1*B1").await.unwrap();
        // E1 = A1-B1
        wb.set_cell_formula("E1", "=A1-B1").await.unwrap();
        // F1 = A1/B1
        wb.set_cell_formula("F1", "=A1/B1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let c1 = sheet.get_value_at(0, 2);
    assert_eq!(formula_text(&c1), "=A1+B1", "C1 formula");

    let d1 = sheet.get_value_at(0, 3);
    assert_eq!(formula_text(&d1), "=A1*B1", "D1 formula");

    let e1 = sheet.get_value_at(0, 4);
    assert_eq!(formula_text(&e1), "=A1-B1", "E1 formula");

    let f1 = sheet.get_value_at(0, 5);
    assert_eq!(formula_text(&f1), "=A1/B1", "F1 formula");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_functions() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Set up some data
        for i in 1..=5 {
            wb.set_cell_value(&format!("A{i}"), i as f64)
                .await
                .unwrap();
        }
        // B1 = SUM(A1:A5)
        wb.set_cell_formula("B1", "=SUM(A1:A5)").await.unwrap();
        // B2 = AVERAGE(A1:A5)
        wb.set_cell_formula("B2", "=AVERAGE(A1:A5)").await.unwrap();
        // B3 = MAX(A1:A5)
        wb.set_cell_formula("B3", "=MAX(A1:A5)").await.unwrap();
        // B4 = MIN(A1:A5)
        wb.set_cell_formula("B4", "=MIN(A1:A5)").await.unwrap();
        // B5 = COUNT(A1:A5)
        wb.set_cell_formula("B5", "=COUNT(A1:A5)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=SUM(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(1, 1)), "=AVERAGE(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(2, 1)), "=MAX(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(3, 1)), "=MIN(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(4, 1)), "=COUNT(A1:A5)");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_if_and_comparison() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    // LO uses semicolons as argument separators (locale-dependent), but
    // BIFF8 token streams are locale-independent — the decompiler always
    // produces commas.
    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        // B1 = IF(A1>5;1;0)
        wb.set_cell_formula("B1", "=IF(A1>5;1;0)").await.unwrap();
        // C1 = IF(A1<>0;"yes";"no")
        wb.set_cell_formula("C1", "=IF(A1<>0;\"yes\";\"no\")")
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=IF(A1>5,1,0)");
    assert_eq!(
        formula_text(&sheet.get_value_at(0, 2)),
        "=IF(A1<>0,\"yes\",\"no\")"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_string_concat() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "hello").await.unwrap();
        // B1 = CONCATENATE(A1;" world") — use function form for locale safety
        wb.set_cell_formula("B1", "=CONCATENATE(A1;\" world\")")
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(
        formula_text(&sheet.get_value_at(0, 1)),
        "=CONCATENATE(A1,\" world\")"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_nested_functions() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "hello").await.unwrap();
        // B1 = LEN(A1)
        wb.set_cell_formula("B1", "=LEN(A1)").await.unwrap();
        // C1 = LEFT(A1;3) — semicolon for LO locale
        wb.set_cell_formula("C1", "=LEFT(A1;3)").await.unwrap();
        // D1 = MID(A1;2;3) — semicolons for LO locale
        wb.set_cell_formula("D1", "=MID(A1;2;3)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=LEN(A1)");
    assert_eq!(formula_text(&sheet.get_value_at(0, 2)), "=LEFT(A1,3)");
    assert_eq!(formula_text(&sheet.get_value_at(0, 3)), "=MID(A1,2,3)");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_constants_and_unary() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // A1 = 100 (integer constant in formula)
        wb.set_cell_formula("A1", "=100").await.unwrap();
        // B1 = 3.14 (float constant)
        wb.set_cell_formula("B1", "=3.14").await.unwrap();
        // C1 = TRUE
        wb.set_cell_formula("C1", "=TRUE()").await.unwrap();
        // D1 = -A1 (unary minus)
        wb.set_cell_formula("D1", "=-A1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 0)), "=100");
    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=3.14");
    // TRUE() in BIFF8 is stored as tFuncV for TRUE (idx 34) with 0 args
    let c1_val = sheet.get_value_at(0, 2);
    let c1_text = formula_text(&c1_val);
    assert!(
        c1_text == "=TRUE()" || c1_text == "=TRUE",
        "C1 should be =TRUE() or =TRUE, got {c1_text}"
    );
    assert_eq!(formula_text(&sheet.get_value_at(0, 3)), "=-A1");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_parentheses() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 2.0).await.unwrap();
        wb.set_cell_value("B1", 3.0).await.unwrap();
        wb.set_cell_value("C1", 4.0).await.unwrap();
        // D1 = (A1+B1)*C1
        wb.set_cell_formula("D1", "=(A1+B1)*C1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(
        formula_text(&sheet.get_value_at(0, 3)),
        "=(A1+B1)*C1"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_cached_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("A2", 20.0).await.unwrap();
        // B1 = SUM(A1:A2) — cached numeric result should be 30
        wb.set_cell_formula("B1", "=SUM(A1:A2)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let val = sheet.get_value_at(0, 1);
    match &val {
        CellValue::Formula {
            text,
            cached_value,
            ..
        } => {
            assert_eq!(text, "=SUM(A1:A2)", "formula text");
            match cached_value.as_deref() {
                Some(CellValue::Number(n)) => {
                    assert!(
                        (*n - 30.0).abs() < f64::EPSILON,
                        "cached value should be 30.0, got {n}"
                    );
                }
                other => panic!("Expected cached Number(30.0), got {:?}", other),
            }
        }
        other => panic!("Expected Formula, got {:?}", other),
    }

    cleanup_fixture(&path);
}

// ── Phase 2: Cross-sheet references and defined names ───────────────

#[test]
fn test_xls_formula_cross_sheet_ref() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Add a second sheet
        wb.add_sheet("Data").await.unwrap();
        // Put a value on Sheet2 (index 1), cell A1
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, duke_sheets_libreoffice::CellValue::Number(42.0))
            .await
            .unwrap();
        // On Sheet1, B1 = Data.A1 (cross-sheet reference)
        // LO uses dot notation for sheet refs: Data.A1
        wb.set_cell_formula("B1", "=Data.A1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let b1 = sheet.get_value_at(0, 1);
    let text = formula_text(&b1);
    // The decompiler should produce "=Data!A1" (Excel-style ! separator)
    assert_eq!(text, "=Data!A1", "cross-sheet ref formula");

    // Verify cached value is 42
    match &b1 {
        CellValue::Formula { cached_value, .. } => {
            match cached_value.as_deref() {
                Some(CellValue::Number(n)) => {
                    assert!(
                        (*n - 42.0).abs() < f64::EPSILON,
                        "cached value should be 42.0, got {n}"
                    );
                }
                other => panic!("Expected cached Number(42.0), got {:?}", other),
            }
        }
        _ => unreachable!(),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_cross_sheet_quoted_name() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Rename Sheet1 to "My Sheet" (contains a space → needs quoting)
        wb.set_sheet_name(0, "My Sheet").await.unwrap();
        wb.add_sheet("Other").await.unwrap();
        // Put 99 on Other.A1
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, duke_sheets_libreoffice::CellValue::Number(99.0))
            .await
            .unwrap();
        // On "My Sheet", B1 = Other.A1
        wb.set_cell_formula("B1", "=Other.A1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    // First sheet is "My Sheet"
    assert_eq!(workbook.worksheet(0).unwrap().name(), "My Sheet");
    let sheet = workbook.worksheet(0).unwrap();

    let b1 = sheet.get_value_at(0, 1);
    let text = formula_text(&b1);
    assert_eq!(text, "=Other!A1", "cross-sheet ref to plain-named sheet");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_named_range() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Put values in A1:A5
        for i in 1..=5 {
            wb.set_cell_value(&format!("A{i}"), i as f64)
                .await
                .unwrap();
        }
        // Define a named range "MyData" = $Sheet1.$A$1:$A$5
        // LO named range content uses absolute $-notation with dot separator
        wb.add_named_range("MyData", "$Sheet1.$A$1:$A$5", 0, 0, 0)
            .await
            .unwrap();
        // B1 = SUM(MyData)
        wb.set_cell_formula("B1", "=SUM(MyData)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let b1 = sheet.get_value_at(0, 1);
    let text = formula_text(&b1);
    assert_eq!(text, "=SUM(MyData)", "named range in formula");

    // Cached value should be 15 (1+2+3+4+5)
    match &b1 {
        CellValue::Formula { cached_value, .. } => {
            match cached_value.as_deref() {
                Some(CellValue::Number(n)) => {
                    assert!(
                        (*n - 15.0).abs() < f64::EPSILON,
                        "cached value should be 15.0, got {n}"
                    );
                }
                other => panic!("Expected cached Number(15.0), got {:?}", other),
            }
        }
        _ => unreachable!(),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_named_range_in_expression() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        // Define "TaxRate" = 0.15 (a named constant)
        // LO syntax: named range content is a formula string
        wb.add_named_range("TaxRate", "0.15", 0, 0, 0)
            .await
            .unwrap();
        // B1 = A1*TaxRate
        wb.set_cell_formula("B1", "=A1*TaxRate").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let b1 = sheet.get_value_at(0, 1);
    let text = formula_text(&b1);
    assert_eq!(text, "=A1*TaxRate", "named range in expression");

    // Cached value should be 1.5 (10 * 0.15)
    match &b1 {
        CellValue::Formula { cached_value, .. } => {
            match cached_value.as_deref() {
                Some(CellValue::Number(n)) => {
                    assert!(
                        (*n - 1.5).abs() < f64::EPSILON,
                        "cached value should be 1.5, got {n}"
                    );
                }
                other => panic!("Expected cached Number(1.5), got {:?}", other),
            }
        }
        _ => unreachable!(),
    }

    cleanup_fixture(&path);
}

/// Phase 3: Shared formulas — when a column has the same formula pattern,
/// Excel stores a single SHAREDFMLA record and each cell uses tExp + tRefN.
/// LO does this automatically when adjacent cells have the same pattern.
#[test]
fn test_xls_formula_shared_formula() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Set up A1:A5 with values
        for i in 1..=5 {
            wb.set_cell_value(&format!("A{i}"), i as f64 * 10.0)
                .await
                .unwrap();
        }
        // B1:B5 = A1*2, A2*2, ... — LO will generate a shared formula
        for i in 1..=5 {
            wb.set_cell_formula(&format!("B{i}"), &format!("=A{i}*2"))
                .await
                .unwrap();
        }
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let wb = XlsReader::read_file(&path).unwrap();
    let ws = wb.worksheet(0).unwrap();

    // All 5 cells should have correct formula text with adjusted refs
    for i in 1u32..=5 {
        let val = ws.get_value_at(i - 1, 1);
        let text = formula_text(&val);
        let expected = format!("=A{}*2", i);
        assert_eq!(text, expected, "shared formula at B{}", i);

        // Cached value should be i*10*2
        match &val {
            CellValue::Formula { cached_value, .. } => {
                let expected_val = i as f64 * 10.0 * 2.0;
                match cached_value.as_deref() {
                    Some(CellValue::Number(n)) => {
                        assert!(
                            (*n - expected_val).abs() < f64::EPSILON,
                            "B{}: expected cached {expected_val}, got {n}",
                            i
                        );
                    }
                    other => panic!("B{}: expected cached Number, got {:?}", i, other),
                }
            }
            _ => unreachable!(),
        }
    }

    cleanup_fixture(&path);
}
