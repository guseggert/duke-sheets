//! Round-trip tests for the XLS writer's FORMULA emission
//! (slice 5a: literals + binary ops + cell/range refs). Function
//! calls and named ranges are not yet supported and round-trip as
//! their cached value (the FORMULA record falls back to a static
//! cell record when the AST contains an unsupported construct).

use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

#[test]
fn arithmetic_formula_round_trips_with_text_intact() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 10.0).expect("set A1");
    ws.set_cell_value("B1", 20.0).expect("set B1");
    ws.set_cell_formula("C1", "=A1+B1").expect("set formula");
    ws.set_formula_result(0, 2, CellValue::Number(30.0))
        .expect("set cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 2).expect("C1 has a formula");
    assert!(formula.contains("A1") && formula.contains("B1"));
    assert!(formula.contains('+'));
    let cached = sheet.get_value("C1").expect("C1 cached");
    assert_eq!(cached.as_number(), Some(30.0));
}

#[test]
fn comparison_formula_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).expect("set A1");
    ws.set_cell_formula("B1", "=A1>3").expect("set formula");
    ws.set_formula_result(0, 1, CellValue::Boolean(true))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("formula present");
    assert!(formula.contains('>'));
    assert!(formula.contains("A1"));
    let cached = sheet.get_value("B1").expect("cached");
    assert_eq!(cached.as_bool(), Some(true));
}

#[test]
fn literal_only_formula_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=42").expect("set formula");
    ws.set_formula_result(0, 0, CellValue::Number(42.0))
        .expect("cached");
    ws.set_cell_formula("A2", "=\"hello\"").expect("set formula");
    ws.set_formula_result(1, 0, CellValue::String("hello".into()))
        .expect("cached");
    ws.set_cell_formula("A3", "=TRUE").expect("set formula");
    ws.set_formula_result(2, 0, CellValue::Boolean(true))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(sheet.get_formula_at(0, 0).is_some());
    assert_eq!(sheet.get_value("A1").unwrap().as_number(), Some(42.0));
    assert!(sheet.get_formula_at(1, 0).is_some());
    assert!(sheet.get_formula_at(2, 0).is_some());
}

#[test]
fn unary_minus_formula_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 10.0).expect("set A1");
    ws.set_cell_formula("B1", "=-A1").expect("set formula");
    ws.set_formula_result(0, 1, CellValue::Number(-10.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("formula present");
    assert!(formula.contains('-'));
    assert!(formula.contains("A1"));
    assert_eq!(
        sheet.get_value("B1").expect("cached").as_number(),
        Some(-10.0)
    );
}

#[test]
fn percent_operator_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 50.0).expect("A1");
    ws.set_cell_formula("B1", "=A1%").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(0.5))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet
        .get_formula_at(0, 1)
        .expect("percent operator must round-trip via formula path");
    assert!(formula.contains('%'), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
}

#[test]
fn concat_operator_round_trips_via_formula_path() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello").expect("A1");
    ws.set_cell_formula("B1", "=A1&\" world\"").expect("formula");
    ws.set_formula_result(0, 1, CellValue::String("hello world".into()))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet
        .get_formula_at(0, 1)
        .expect("concat operator must round-trip via formula path");
    assert!(formula.contains('&'), "got {formula:?}");
    assert!(formula.contains("A1"), "got {formula:?}");
}

#[test]
fn absolute_cell_reference_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 100.0).expect("set A1");
    ws.set_cell_formula("B1", "=$A$1*2").expect("set formula");
    ws.set_formula_result(0, 1, CellValue::Number(200.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("formula present");
    assert!(
        formula.contains("$A$1"),
        "absolute ref must round-trip with $ markers; got {formula:?}"
    );
}

#[test]
fn range_reference_in_formula_round_trips() {
    // Slice 5a emits tArea but no SUM-style functions, so use a binary
    // operator on what the parser turns into a range value: cell minus
    // cell still uses tRef twice. To exercise tArea we'd need a
    // function call \u2014 deferred. This test instead verifies a "range
    // operator" expression round-trips, even if Excel-level semantics
    // differ.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_cell_value("A2", 2.0).expect("A2");
    ws.set_cell_value("A3", 3.0).expect("A3");
    ws.set_cell_formula("B1", "=A1-A3").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(-2.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("formula present");
    assert!(formula.contains("A1") && formula.contains("A3"));
}

#[test]
fn sum_function_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 10.0).expect("A1");
    ws.set_cell_value("A2", 20.0).expect("A2");
    ws.set_cell_value("A3", 30.0).expect("A3");
    ws.set_cell_formula("B1", "=SUM(A1:A3)").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(60.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("SUM round-trips");
    assert!(formula.to_uppercase().contains("SUM"), "got {formula:?}");
    assert!(formula.contains("A1") && formula.contains("A3"));
    assert_eq!(
        sheet.get_value("B1").expect("cached").as_number(),
        Some(60.0)
    );
}

#[test]
fn if_function_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 5.0).expect("A1");
    ws.set_cell_formula("B1", "=IF(A1>0,1,0)").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(1.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("IF round-trips");
    assert!(formula.to_uppercase().contains("IF"), "got {formula:?}");
}

#[test]
fn nested_function_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 9.0).expect("A1");
    ws.set_cell_value("A2", 16.0).expect("A2");
    ws.set_cell_formula("B1", "=SUM(SQRT(A1),SQRT(A2))")
        .expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(7.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet.get_formula_at(0, 1).expect("nested round-trips");
    assert!(formula.to_uppercase().contains("SUM"));
    assert!(formula.to_uppercase().contains("SQRT"));
}

#[test]
fn case_insensitive_function_lookup() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 4.0).expect("A1");
    ws.set_cell_formula("B1", "=sum(A1)").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(4.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let formula = sheet
        .get_formula_at(0, 1)
        .expect("case-insensitive sum round-trips");
    assert!(formula.to_uppercase().contains("SUM"), "got {formula:?}");
}

#[test]
fn unknown_function_falls_back_to_static_value() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=ABSOLUTELY_NOT_A_REAL_FUNCTION(1)")
        .expect("formula");
    ws.set_formula_result(0, 0, CellValue::Number(99.0))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    assert!(
        sheet.get_formula_at(0, 0).is_none(),
        "unknown function falls back to static value"
    );
    assert_eq!(
        sheet.get_value("A1").expect("cached").as_number(),
        Some(99.0)
    );
}

#[test]
fn cached_string_result_includes_string_followup_record() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_formula("A1", "=\"foo\"&\"bar\"").expect("formula");
    ws.set_formula_result(0, 0, CellValue::String("foobar".into()))
        .expect("cached");

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    let cached = sheet.get_value("A1").expect("cached");
    assert_eq!(cached.as_string().as_deref(), Some("foobar"));
}

#[test]
fn many_formulas_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let count = 50u32;
    for row in 0..count {
        let addr = format!("A{}", row + 1);
        ws.set_cell_value(&addr, row as f64).expect("set");
    }
    for row in 0..count {
        let addr = format!("B{}", row + 1);
        let formula = format!("=A{}*2", row + 1);
        ws.set_cell_formula(&addr, &formula).expect("set formula");
        ws.set_formula_result(row, 1, CellValue::Number(row as f64 * 2.0))
            .expect("cached");
    }

    let parsed = write_then_read(&wb);
    let sheet = parsed.worksheet(0).unwrap();
    for row in 0..count {
        let addr = format!("B{}", row + 1);
        assert!(sheet.get_formula_at(row, 1).is_some(), "formula at row {row}");
        let cached = sheet.get_value(&addr).expect("cached");
        assert_eq!(cached.as_number(), Some(row as f64 * 2.0));
    }
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_evaluate_sum_function_we_emit() {
    use std::net::TcpStream;
    use std::time::Duration;

    if TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        Duration::from_secs(2),
    )
    .is_err()
    {
        panic!("LO URP not reachable; start with `mise run urp:start`");
    }

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    ws.set_cell_value("A2", 2.0).expect("A2");
    ws.set_cell_value("A3", 3.0).expect("A3");
    ws.set_cell_value("A4", 4.0).expect("A4");
    ws.set_cell_value("A5", 5.0).expect("A5");
    ws.set_cell_formula("B1", "=SUM(A1:A5)").expect("formula");
    ws.set_formula_result(0, 1, CellValue::Number(15.0))
        .expect("cached");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_sumfn_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<f64, String> = rt.block_on(async {
        let mut bridge = duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect(
            "127.0.0.1",
            2002,
        )
        .await
        .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb.get_cell_value("B1").await.map_err(|e| format!("B1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let b1 = outcome.expect("LO must compute SUM");
    assert!((b1 - 15.0).abs() < 1e-9, "LO computed B1 = {b1}");
}

#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_evaluate_arithmetic_formula_we_emit() {
    use std::net::TcpStream;
    use std::time::Duration;

    if TcpStream::connect_timeout(
        &"127.0.0.1:2002".parse().unwrap(),
        Duration::from_secs(2),
    )
    .is_err()
    {
        panic!("LO URP not reachable; start with `mise run urp:start`");
    }

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 7.0).expect("A1");
    ws.set_cell_value("B1", 6.0).expect("B1");
    ws.set_cell_formula("C1", "=A1*B1").expect("C1 formula");
    ws.set_formula_result(0, 2, CellValue::Number(42.0))
        .expect("cached");
    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");

    std::fs::create_dir_all("/tmp/duke-sheets-urp").expect("shared dir");
    let pid = std::process::id();
    let path = format!("/tmp/duke-sheets-urp/duke_formula_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<f64, String> = rt.block_on(async {
        let mut bridge = duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect(
            "127.0.0.1",
            2002,
        )
        .await
        .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb.get_cell_value("C1").await.map_err(|e| format!("C1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let c1 = outcome.expect("LO must read C1");
    assert!((c1 - 42.0).abs() < 1e-9, "LO computed C1 = {c1}");
}
