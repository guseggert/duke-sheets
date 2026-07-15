//! Excel COM parity tests for the XLSB (BIFF12) writer.
//!
//! Each test builds a workbook in memory, writes it via `XlsbWriter`,
//! pushes to the Windows VM, opens in real Excel (asserting no
//! `Repaired` warning), re-saves to a second `.xlsb`, pulls it back,
//! and reads it with `XlsbReader`. The `Repaired` check is the
//! canonical signal for writer bugs - Excel auto-recovers from
//! malformations our permissive reader silently tolerates.
//!
//! All tests are batched into one process to amortise the per-test
//! VM round-trip cost (~15-25s of warm-VM time per test).

use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
use duke_sheets_core::conditional_format::{CfOperator, CfRuleType, ConditionalFormatRule};
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::Color;
use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo};
use duke_sheets_core::validation::{DataValidation, ValidationOperator, ValidationType};
use duke_sheets_core::worksheet::{PageOrientation, SheetProtection, SheetVisibility};
use duke_sheets_core::{CellAddress, CellRange, CellValue, Hyperlink, Workbook};

use crate::{
    atp_all_formulas, cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm,
    roundtrip_through_excel_xlsb, roundtrip_through_excel_xlsb_bytes, temp_fixture_xlsb,
    xlsb_formula_ptg_streams_for_compare,
};

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).unwrap(),
        CellAddress::parse(end).unwrap(),
    )
}

/// Formulas exercising the XLSB formula compiler's operand-class threading,
/// volatile prefix, fixed-arity allow-list, IF/CHOOSE short-circuit, and
/// reference-class function tokens. Verified byte-for-byte against Excel's
/// native XLSB emission.
const XLSB_FORMULA_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=ABS(A1)", 2.0),           // value-class ref arg; PtgFunc
    ("B2", "=SUM(A1,A2)", 5.0),        // R-class cell refs in aggregator
    ("B3", "=SUM(A1:A3)", 9.0),        // single-arg SUM → PtgAttrSum, R-area
    ("B4", "=VLOOKUP(A1,A1:A3,1)", 2.0), // not on allow-list → PtgFuncVar
    ("B5", "=NOW()", 45000.0),         // volatile prefix
    ("B6", "=IF(A1>0,1,2)", 1.0),      // PtgAttrIf 3-arg
    ("B7", "=IF(A1>0,A1)", 2.0),       // PtgAttrIf 2-arg
    ("B8", "=CHOOSE(A1,10,20)", 20.0), // PtgAttrChoose
    ("B9", "=SUM(IF(A1>0,A1,A2))", 2.0), // nested IF R-class in SUM
    ("B10", "=SUM(OFFSET(A1,0,0))", 2.0), // OFFSET R-class + volatile
    ("B11", "=INDEX(A1:A3,1)", 2.0),   // INDEX arg0 R-class, V token
    ("B12", "=+A1", 2.0),              // PtgUplus
    ("B13", "=(A1+A2)*2", 10.0),       // PtgParen
    ("B14", "=((A1))", 2.0),           // nested PtgParen
    ("B15", "=SUM({1,2,3})", 6.0),     // array constant: PtgArray(A) + rgcb
    ("B16", "=SUM({1,2;3,4})", 10.0),  // 2x2 array constant
    ("B17", "=COUNTA({\"ab\",\"cde\"})", 2.0), // SerAr string elements (u16 cch)
    ("B18", "=COUNT({1,TRUE,3})", 2.0), // SerAr bool element (1 byte, no pad)
    ("B19", "=COUNT({1,#N/A,3})", 2.0), // SerAr error element (1 byte + 3 reserved)
    ("B20", "=SUM(-A1)", -2.0),        // unary operand class under R-forced arg
    ("B21", "=-A1+A2", 1.0),           // unary minus on a ref at value position
];

fn xlsb_formula_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 2.0).unwrap();
    ws.set_cell_value("A2", 3.0).unwrap();
    ws.set_cell_value("A3", 4.0).unwrap();
    for (cell, formula, expected) in XLSB_FORMULA_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_xlsb_formula_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xlsb();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 2.0).unwrap();
        wb.set_cell_value("A2", 3.0).unwrap();
        wb.set_cell_value("A3", 4.0).unwrap();
        for (cell, formula, _) in XLSB_FORMULA_FORMULAS {
            wb.set_cell_formula(cell, formula).unwrap();
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 50).unwrap();
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_byte_parity_for_xlsb_formulas_we_emit() {
    let wb = xlsb_formula_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(&wb);
    let writer_ptgs = xlsb_formula_ptg_streams_for_compare(&writer_bytes);
    assert_eq!(
        writer_ptgs.len(),
        XLSB_FORMULA_FORMULAS.len(),
        "formula-stream extraction came back short; the parity comparison below would be vacuous"
    );
    let resave_ptgs = xlsb_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLSB formula token streams on re-save"
    );
    let authored_ptgs = xlsb_formula_ptg_streams_for_compare(&excel_authored_xlsb_formula_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLSB formula token streams differ from Excel-authored output"
    );
}

/// Analysis-ToolPak functions in XLSB. Unlike XLS (BIFF8), which serializes
/// these via the add-in PtgNameX + EXTERNNAME mechanism, XLSB (Excel 2007+)
/// emits them as NATIVE PtgFunc/PtgFuncVar carrying the real Ftab index —
/// EDATE/EOMONTH are fixed-arity (PtgFunc 0x41), the rest variable-arity
/// (PtgFuncVar 0x42). Verified byte-for-byte against Excel-authored output.
const XLSB_ATP_FORMULAS: &[(&str, &str, f64)] = &[
    ("B1", "=EDATE(A1,12)", 0.0),       // iftab=449, fixed 2-arg → PtgFunc
    ("B2", "=EOMONTH(A1,1)", 0.0),      // iftab=450, fixed 2-arg → PtgFunc
    ("B3", "=GCD(A1,A2)", 0.0),         // iftab=473, variable → PtgFuncVar
    ("B4", "=NETWORKDAYS(A1,A2)", 0.0), // iftab=472, variable → PtgFuncVar
    ("B5", "=WORKDAY(A1,5)", 0.0),      // iftab=471, variable → PtgFuncVar
];

fn xlsb_atp_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 43831.0).unwrap();
    ws.set_cell_value("A2", 43862.0).unwrap();
    for (cell, formula, expected) in XLSB_ATP_FORMULAS {
        ws.set_cell_formula(cell, formula).unwrap();
        let addr = CellAddress::parse(cell).unwrap();
        ws.set_formula_result(addr.row, addr.col, CellValue::Number(*expected))
            .unwrap();
    }
    wb
}

fn excel_authored_xlsb_atp_bytes() -> Vec<u8> {
    let fixture = temp_fixture_xlsb();
    ensure_vm_temp_dir();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        wb.set_cell_value("A1", 43831.0).unwrap();
        wb.set_cell_value("A2", 43862.0).unwrap();
        for (cell, formula, _) in XLSB_ATP_FORMULAS {
            wb.set_cell_formula(cell, formula).unwrap();
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 50).unwrap();
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    cleanup_fixture(&fixture);
    bytes
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_byte_parity_for_xlsb_atp_functions_we_emit() {
    let wb = xlsb_atp_workbook();
    let (_result, writer_bytes, excel_bytes) = roundtrip_through_excel_xlsb_bytes(&wb);
    let writer_ptgs = xlsb_formula_ptg_streams_for_compare(&writer_bytes);
    let resave_ptgs = xlsb_formula_ptg_streams_for_compare(&excel_bytes);
    assert_eq!(
        writer_ptgs, resave_ptgs,
        "Excel canonicalized our XLSB ATP token streams on re-save"
    );
    assert_eq!(
        writer_ptgs.len(),
        XLSB_ATP_FORMULAS.len(),
        "expected one formula stream per ATP formula"
    );
    let authored_ptgs = xlsb_formula_ptg_streams_for_compare(&excel_authored_xlsb_atp_bytes());
    assert_eq!(
        writer_ptgs, authored_ptgs,
        "our XLSB ATP token streams differ from Excel-authored output"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_external_udf_xlsb_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 7.0).unwrap();
    ws.set_cell_formula("B1", r#"=[1]!TBLink("acct",A1)"#)
        .unwrap();
    ws.set_formula_result(0, 1, CellValue::Number(42.0))
        .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let ws = result.worksheet(0).unwrap();
    let formula = ws
        .get_formula_at(0, 1)
        .expect("external UDF formula must survive Excel re-save");
    assert!(
        formula.to_ascii_uppercase().contains("TBLINK"),
        "external UDF formula lost through Excel: {formula:?}"
    );
}

/// Comprehensive: every Analysis-ToolPak function (Ftab 384..=476, minus
/// RANDBETWEEN) in one workbook, byte-compared against Excel's native XLSB
/// emission. Pins PtgFunc-vs-PtgFuncVar and R-class arguments across the whole
/// range, not just a hand-picked sample.
///
/// Any formula Excel rejects on entry is a hard failure (a wrong `min_args`
/// in our metadata yields an invalid minimal call), asserted via
/// `rejected.is_empty()` below. For XLSB there is no SUPBOOK/EXTERNNAME, so
/// token-stream parity
/// against Excel's authoring is sufficient (a malformed token stream would not
/// match Excel's bytes).
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_byte_parity_for_all_xlsb_atp_functions_we_emit() {
    use duke_sheets_xlsb::XlsbWriter;
    use std::io::Cursor;

    let formulas = atp_all_formulas();
    // Lock the coverage count: the add-in range is Ftab 384..=476 (93 fns).
    assert_eq!(formulas.len(), 93, "must exercise all 93 ATP functions");

    // Author each formula in Excel; keep the set Excel accepts + its bytes.
    let fixture = temp_fixture_xlsb();
    ensure_vm_temp_dir();
    let mut rejected = Vec::new();
    let mut accepted: Vec<(String, String)> = Vec::new();
    {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create Excel workbook");
        for i in 1..=10u32 {
            wb.set_cell_value(&format!("A{i}"), 30000.0 + i as f64)
                .unwrap();
        }
        for (cell, formula) in &formulas {
            match wb.set_cell_formula(cell, formula) {
                Ok(_) => accepted.push((cell.clone(), formula.clone())),
                Err(e) => rejected.push(format!("{formula} ({e})")),
            }
        }
        excel.recalculate().unwrap();
        wb.save_as(&fixture.vm_path, 50).unwrap();
        wb.close().unwrap();
    }
    pull_file_from_vm(&fixture);
    let excel_bytes = std::fs::read(&fixture.host_path).unwrap();
    cleanup_fixture(&fixture);

    // Build our writer's bytes from the SAME accepted formulas (VM-free).
    let mut wb = Workbook::new();
    {
        let ws = wb.worksheet_mut(0).unwrap();
        for i in 1..=10u32 {
            ws.set_cell_value_at(i - 1, 0, 30000.0 + i as f64).unwrap();
        }
        for (cell, formula) in &accepted {
            ws.set_cell_formula(cell, formula).unwrap();
        }
    }
    let mut our_buf = Vec::new();
    XlsbWriter::write(&wb, Cursor::new(&mut our_buf)).expect("our write");

    let ours = xlsb_formula_ptg_streams_for_compare(&our_buf);
    let excel = xlsb_formula_ptg_streams_for_compare(&excel_bytes);
    // The full range is expected to be valid — any rejection is a real
    // metadata bug (wrong min_args) to investigate, not a function to skip.
    assert!(
        rejected.is_empty(),
        "Excel rejected {} ATP call(s): {:?}",
        rejected.len(),
        rejected
    );
    assert_eq!(
        accepted.len(),
        formulas.len(),
        "all ATP formulas should be authored"
    );
    assert_eq!(
        ours.len(),
        accepted.len(),
        "our writer must emit a formula stream for every accepted ATP call"
    );

    let mut diffs = Vec::new();
    for (i, (w, a)) in ours.iter().zip(excel.iter()).enumerate() {
        if w.rgce != a.rgce {
            let fname = accepted.get(i).map(|(_, f)| f.as_str()).unwrap_or("?");
            diffs.push(format!(
                "{fname}\n      ours ={:02X?}\n      excel={:02X?}",
                w.rgce, a.rgce
            ));
        }
    }
    assert!(
        diffs.is_empty() && ours == excel,
        "XLSB ATP token-stream mismatches ({} of {}):\n{}",
        diffs.len(),
        accepted.len(),
        diffs.join("\n")
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_hyperlinks_we_emit() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Other").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "go").unwrap();
    ws.set_cell_value("A2", "elsewhere").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com".into(),
            display: Some("Click here".into()),
            tooltip: Some("Visit the example site".into()),
            location: None,
        },
    )
    .unwrap();
    ws.set_hyperlink(
        "A2",
        Hyperlink {
            target: String::new(),
            display: Some("Elsewhere".into()),
            tooltip: None,
            location: Some("Other!B5".into()),
        },
    )
    .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("go"));
    assert_eq!(s.get_value_at(1, 0).as_string(), Some("elsewhere"));

    let hl_external = s
        .hyperlink("A1")
        .expect("external hyperlink on A1 lost after round-trip");
    assert!(
        hl_external.target.contains("example.com"),
        "external target mangled: {:?}",
        hl_external.target
    );
    assert_eq!(
        hl_external.display.as_deref(),
        Some("Click here"),
        "external display text lost after round-trip"
    );
    assert_eq!(
        hl_external.tooltip.as_deref(),
        Some("Visit the example site"),
        "external tooltip lost after round-trip"
    );

    let hl_internal = s
        .hyperlink("A2")
        .expect("internal hyperlink on A2 lost after round-trip");
    let internal_addr = hl_internal
        .location
        .as_deref()
        .or(Some(hl_internal.target.as_str()))
        .unwrap_or("");
    assert!(
        internal_addr.contains("Other") && internal_addr.contains("B5"),
        "internal hyperlink address mangled: target={:?} location={:?}",
        hl_internal.target,
        hl_internal.location
    );
    assert_eq!(
        hl_internal.display.as_deref(),
        Some("Elsewhere"),
        "internal display text lost after round-trip"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_mailto_hyperlink_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "email").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "mailto:foo@example.com?subject=hi".into(),
            display: Some("Email me".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let hl = s
        .hyperlink("A1")
        .expect("mailto hyperlink lost after round-trip");
    assert!(
        hl.target.starts_with("mailto:"),
        "target dropped mailto: prefix: {:?}",
        hl.target
    );
    assert!(
        hl.target.contains("foo@example.com"),
        "mailto address mangled: {:?}",
        hl.target
    );
    assert_eq!(
        hl.display.as_deref(),
        Some("Email me"),
        "mailto display text lost"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_cross_sheet_formulas_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();

    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 20.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 30.0)
        .unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=Data!A1").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(10.0))
        .unwrap();
    calc.set_cell_formula("B2", "=SUM(Data!A1:A3)").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(60.0))
        .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 10.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(10), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 60.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(60), got {other:?}"),
    }

    // Verify the formula TEXT survived too — otherwise Excel could
    // have inlined the cross-sheet reference (e.g. rewritten "=Data!A1"
    // as "=10") and the cached value would still match but the
    // formula structure would be lost.
    let f1 = s
        .get_formula_at(0, 1)
        .expect("B1 must still be a formula after Excel re-save");
    assert!(
        f1.contains("Data") && f1.contains("A1"),
        "cross-sheet ref lost from B1 formula: {f1:?}"
    );
    let f2 = s
        .get_formula_at(1, 1)
        .expect("B2 must still be a formula after Excel re-save");
    assert!(
        f2.contains("SUM") && f2.contains("Data") && f2.contains("A1:A3"),
        "cross-sheet SUM lost from B2 formula: {f2:?}"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_autofilter_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Name").unwrap();
    ws.set_cell_value("B1", "Value").unwrap();
    ws.set_cell_value("A2", "alpha").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "beta").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let mut af = AutoFilter::new(range("A1", "B3"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 1.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("Name"));
    assert_eq!(s.get_value_at(0, 1).as_string(), Some("Value"));
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    assert_eq!(
        af.range.start,
        CellAddress::parse("A1").unwrap(),
        "autofilter range start lost"
    );
    assert_eq!(
        af.range.end,
        CellAddress::parse("B3").unwrap(),
        "autofilter range end lost"
    );

    let top10_col = af
        .filter_columns
        .iter()
        .find(|fc| fc.col_id == 1)
        .expect("Top10 filter column on column B must survive");
    match &top10_col.filter {
        ColumnFilter::Top10(t) => {
            assert!(t.top, "Top10.top flag lost (was true)");
            assert!(!t.percent, "Top10.percent flag flipped");
            assert!((t.val - 1.0).abs() < 1e-9, "Top10.val drifted: {}", t.val);
        }
        other => panic!("expected Top10 filter on column B, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_dynamic_filter_we_emit() {
    use duke_sheets_core::auto_filter::{DynamicFilter, DynamicFilterType};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 10.0).unwrap();
    ws.set_cell_value("A3", 50.0).unwrap();
    ws.set_cell_value("A4", 90.0).unwrap();
    ws.set_cell_value("A5", 30.0).unwrap();

    let mut af = AutoFilter::new(range("A1", "A5"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Dynamic(DynamicFilter {
            filter_type: DynamicFilterType::AboveAverage,
            val: None,
            max_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    let col0 = af
        .filter_columns
        .iter()
        .find(|fc| fc.col_id == 0)
        .expect("dynamic filter column on A lost");
    match &col0.filter {
        ColumnFilter::Dynamic(d) => {
            assert_eq!(
                d.filter_type,
                DynamicFilterType::AboveAverage,
                "dynamic filter type drifted: {:?}",
                d.filter_type
            );
        }
        other => panic!("expected Dynamic filter on A, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_color_filter_we_emit() {
    use duke_sheets_core::auto_filter::ColorFilter;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 10.0).unwrap();
    ws.set_cell_value("A3", 20.0).unwrap();

    let mut af = AutoFilter::new(range("A1", "A3"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Color(ColorFilter {
            dxf_id: Some(0),
            cell_color: true,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    let col0 = af
        .filter_columns
        .iter()
        .find(|fc| fc.col_id == 0)
        .expect("color filter column on A lost");
    match &col0.filter {
        ColumnFilter::Color(c) => {
            assert!(c.cell_color, "cell_color flag flipped");
            // dxf_id may get rewritten by Excel; just ensure something
            // is present so we know the record survived as a color
            // filter (not a discarded one).
            assert!(c.dxf_id.is_some(), "dxf_id stripped after round-trip");
        }
        other => panic!("expected Color filter on A, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_custom_filter_we_emit() {
    use duke_sheets_core::auto_filter::{CustomFilterCondition, CustomFilters, FilterOperator};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 3.0).unwrap();
    ws.set_cell_value("A3", 7.0).unwrap();
    ws.set_cell_value("A4", 12.0).unwrap();

    let mut af = AutoFilter::new(range("A1", "A4"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Custom(CustomFilters {
            and: true,
            conditions: vec![
                CustomFilterCondition {
                    operator: FilterOperator::GreaterThan,
                    value: "5".to_string(),
                },
                CustomFilterCondition {
                    operator: FilterOperator::LessThan,
                    value: "10".to_string(),
                },
            ],
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    let col0 = af
        .filter_columns
        .iter()
        .find(|fc| fc.col_id == 0)
        .expect("custom filter column on A lost");
    match &col0.filter {
        ColumnFilter::Custom(cf) => {
            assert!(cf.and, "AND flag flipped to OR");
            assert_eq!(cf.conditions.len(), 2, "conditions lost");
            assert_eq!(cf.conditions[0].operator, FilterOperator::GreaterThan);
            assert_eq!(cf.conditions[0].value, "5", "first condition value");
            assert_eq!(cf.conditions[1].operator, FilterOperator::LessThan);
            assert_eq!(cf.conditions[1].value, "10", "second condition value");
        }
        other => panic!("expected Custom filter on A, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_discrete_value_filter_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Color").unwrap();
    ws.set_cell_value("A2", "red").unwrap();
    ws.set_cell_value("A3", "green").unwrap();
    ws.set_cell_value("A4", "blue").unwrap();

    let mut af = AutoFilter::new(range("A1", "A4"));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(duke_sheets_core::auto_filter::ValueFilter {
            values: vec!["red".into(), "blue".into()],
            blank: false,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive Excel round-trip");
    let col0 = af
        .filter_columns
        .iter()
        .find(|fc| fc.col_id == 0)
        .expect("discrete-values filter column on A lost");
    match &col0.filter {
        ColumnFilter::Values(v) => {
            let mut sorted = v.values.clone();
            sorted.sort();
            assert_eq!(
                sorted,
                vec!["blue".to_string(), "red".to_string()],
                "discrete filter values mangled: {:?}",
                v.values
            );
        }
        other => panic!("expected Values filter on A, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_data_validations_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "list").unwrap();
    ws.set_cell_value("B1", 50.0).unwrap();
    ws.set_cell_value("C1", 7.0).unwrap();

    let mut v_list = DataValidation::list("Red,Green,Blue");
    v_list.ranges = vec![range("A1", "A1")];
    ws.add_data_validation(v_list);

    let mut v_whole = DataValidation::whole_number_between(ValidationOperator::Between, "1", "100");
    v_whole.ranges = vec![range("B1", "B1")];
    ws.add_data_validation(v_whole);

    let mut v_custom = DataValidation::new();
    v_custom.validation_type = ValidationType::Custom {
        formula: "=ISNUMBER(C1)".into(),
    };
    v_custom.ranges = vec![range("C1", "C1")];
    ws.add_data_validation(v_custom);

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("list"));
    let validations = s.data_validations();
    assert_eq!(
        validations.len(),
        3,
        "expected 3 data validations to survive, got {}",
        validations.len()
    );

    let v_a1 = validations
        .iter()
        .find(|v| {
            v.ranges
                .iter()
                .any(|r| r.start == CellAddress::parse("A1").unwrap())
        })
        .expect("List validation on A1 lost");
    match &v_a1.validation_type {
        ValidationType::List { source } => {
            assert_eq!(source, "Red,Green,Blue", "List source mangled: {source:?}")
        }
        other => panic!("expected List validation on A1, got {other:?}"),
    }

    let v_b1 = validations
        .iter()
        .find(|v| {
            v.ranges
                .iter()
                .any(|r| r.start == CellAddress::parse("B1").unwrap())
        })
        .expect("Whole validation on B1 lost");
    match &v_b1.validation_type {
        ValidationType::Whole {
            operator,
            value1,
            value2,
        } => {
            assert_eq!(*operator, ValidationOperator::Between);
            assert_eq!(value1, "1", "Whole lower bound mangled: {value1:?}");
            assert_eq!(
                value2.as_deref(),
                Some("100"),
                "Whole upper bound mangled: {value2:?}"
            );
        }
        other => panic!("expected Whole validation on B1, got {other:?}"),
    }

    let v_c1 = validations
        .iter()
        .find(|v| {
            v.ranges
                .iter()
                .any(|r| r.start == CellAddress::parse("C1").unwrap())
        })
        .expect("Custom validation on C1 lost");
    match &v_c1.validation_type {
        ValidationType::Custom { formula } => assert!(
            formula.contains("ISNUMBER"),
            "Custom formula mangled: {formula:?}"
        ),
        other => panic!("expected Custom validation on C1, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_conditional_formats_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 150.0).unwrap();
    ws.set_cell_value("A2", 50.0).unwrap();
    ws.set_cell_value("B1", 75.0).unwrap();

    let mut r1 = ConditionalFormatRule::cell_is_greater_than("100");
    r1.ranges = vec![range("A1", "A2")];
    ws.add_conditional_format(r1);

    let mut r2 = ConditionalFormatRule::cell_is_between("10", "100");
    r2.ranges = vec![range("B1", "B1")];
    ws.add_conditional_format(r2);

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 0);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 150.0).abs() < 1e-9),
        other => panic!("A1 expected Number(150), got {other:?}"),
    }
    let rules = s.conditional_formats();
    assert_eq!(rules.len(), 2, "expected 2 CF rules, got {}", rules.len());

    // Find each rule by its operator since Excel may rewrite ranges
    // (e.g. consolidating overlapping rules or reordering by priority).
    let r_gt = rules
        .iter()
        .find(|r| {
            matches!(
                &r.rule_type,
                CfRuleType::CellIs {
                    operator: CfOperator::GreaterThan,
                    ..
                }
            )
        })
        .expect("CellIs(GreaterThan) rule lost");
    match &r_gt.rule_type {
        CfRuleType::CellIs { formula1, .. } => assert!(
            formula1.contains("100"),
            "greater-than threshold mangled: {formula1:?}"
        ),
        _ => unreachable!(),
    }

    let r_between = rules
        .iter()
        .find(|r| {
            matches!(
                &r.rule_type,
                CfRuleType::CellIs {
                    operator: CfOperator::Between,
                    ..
                }
            )
        })
        .expect("CellIs(Between) rule lost");
    match &r_between.rule_type {
        CfRuleType::CellIs {
            formula1, formula2, ..
        } => {
            assert!(
                formula1.contains("10"),
                "between lower bound mangled: {formula1:?}"
            );
            assert!(
                formula2
                    .as_deref()
                    .map(|s| s.contains("100"))
                    .unwrap_or(false),
                "between upper bound mangled: {formula2:?}"
            );
        }
        _ => unreachable!(),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_advanced_conditional_formats_we_emit() {
    use duke_sheets_core::conditional_format::{CfValueType, IconSetStyle};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for row in 0..5 {
        ws.set_cell_value_at(row, 0, (row + 1) as f64 * 10.0)
            .unwrap();
        ws.set_cell_value_at(row, 1, (row + 1) as f64 * 20.0)
            .unwrap();
        ws.set_cell_value_at(row, 2, (row + 1) as f64 * 30.0)
            .unwrap();
    }
    let mut scale = ConditionalFormatRule::color_scale_3(
        Color::rgb(255, 0, 0),
        Color::rgb(255, 255, 0),
        Color::rgb(0, 255, 0),
    );
    scale.ranges = vec![range("A1", "A5")];
    ws.add_conditional_format(scale);
    let mut bar = ConditionalFormatRule::data_bar(Color::rgb(99, 142, 198));
    bar.ranges = vec![range("B1", "B5")];
    ws.add_conditional_format(bar);
    let mut icons = ConditionalFormatRule::icon_set(IconSetStyle::Arrows3);
    icons.ranges = vec![range("C1", "C5")];
    ws.add_conditional_format(icons);

    let result = roundtrip_through_excel_xlsb(&wb);
    let rules = result.worksheet(0).unwrap().conditional_formats();
    assert_eq!(rules.len(), 3, "advanced CF rules lost");
    let scale = rules
        .iter()
        .find_map(|rule| match &rule.rule_type {
            CfRuleType::ColorScale { colors } => Some(colors),
            _ => None,
        })
        .expect("3-color scale lost");
    assert_eq!(scale.len(), 3);
    assert_eq!(scale[0].value_type, CfValueType::Min);
    assert_eq!(scale[1].value_type, CfValueType::Percentile);
    assert_eq!(scale[1].value.as_deref(), Some("50"));
    assert_eq!(scale[2].value_type, CfValueType::Max);
    assert_eq!(scale[0].color, Color::rgb(255, 0, 0));
    assert_eq!(scale[1].color, Color::rgb(255, 255, 0));
    assert_eq!(scale[2].color, Color::rgb(0, 255, 0));

    let data_bar = rules
        .iter()
        .find_map(|rule| match &rule.rule_type {
            CfRuleType::DataBar {
                min_value,
                max_value,
                color,
                show_value,
                ..
            } => Some((min_value, max_value, color, show_value)),
            _ => None,
        })
        .expect("data bar lost");
    assert_eq!(data_bar.0.value_type, CfValueType::Min);
    assert_eq!(data_bar.1.value_type, CfValueType::Max);
    assert_eq!(*data_bar.2, Color::rgb(99, 142, 198));
    assert!(*data_bar.3);

    let icon_set = rules
        .iter()
        .find_map(|rule| match &rule.rule_type {
            CfRuleType::IconSet {
                icon_style,
                values,
                reverse,
                show_value,
            } => Some((icon_style, values, reverse, show_value)),
            _ => None,
        })
        .expect("3 Arrows icon set lost");
    assert_eq!(*icon_set.0, IconSetStyle::Arrows3);
    assert_eq!(icon_set.1.len(), 3);
    assert!(icon_set
        .1
        .iter()
        .all(|value| value.value_type == CfValueType::Percent));
    assert_eq!(
        icon_set
            .1
            .iter()
            .map(|value| value.value.as_deref())
            .collect::<Vec<_>>(),
        vec![Some("0"), Some("33"), Some("67")]
    );
    assert!(!icon_set.2);
    assert!(*icon_set.3);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_validation_messages_we_emit() {
    use duke_sheets_core::validation::ValidationErrorStyle;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, value) in ["Red", "Green", "Blue"].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 5, *value).unwrap();
    }
    let validation = DataValidation::list("=$F$1:$F$3")
        .with_range(range("D1", "D3"))
        .with_input_message("Pick", "Choose a listed color")
        .with_error_message("Invalid", "Use the dropdown")
        .with_error_style(ValidationErrorStyle::Warning);
    ws.add_data_validation(validation);

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();
    let validation = sheet
        .data_validations()
        .iter()
        .find(|validation| {
            validation
                .ranges
                .iter()
                .any(|cell_range| cell_range.start == CellAddress::parse("D1").unwrap())
        })
        .expect("range-list validation lost");
    match &validation.validation_type {
        ValidationType::List { source } => assert!(source.contains("$F$1:$F$3")),
        other => panic!("expected range list, got {other:?}"),
    }
    assert_eq!(validation.input_title.as_deref(), Some("Pick"));
    assert_eq!(
        validation.input_message.as_deref(),
        Some("Choose a listed color")
    );
    assert_eq!(validation.error_title.as_deref(), Some("Invalid"));
    assert_eq!(
        validation.error_message.as_deref(),
        Some("Use the dropdown")
    );
    assert_eq!(validation.error_style, ValidationErrorStyle::Warning);
    assert!(validation.show_input_message);
    assert!(validation.show_error_alert);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_split_panes_we_emit() {
    use duke_sheets_core::worksheet::SplitPanes;

    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_split_panes(Some(SplitPanes {
            x_split: 1500.0,
            y_split: 3000.0,
            top_left: Some((4, 3)),
            active_pane: Some("bottomRight".into()),
        }));

    let result = roundtrip_through_excel_xlsb(&wb);
    let split = result
        .worksheet(0)
        .unwrap()
        .split_panes()
        .expect("split pane lost");
    assert!((split.x_split - 1500.0).abs() < 1.0);
    assert!((split.y_split - 3000.0).abs() < 1.0);
    assert_eq!(split.top_left, Some((4, 3)));
    assert_eq!(split.active_pane.as_deref(), Some("bottomRight"));
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_sheet_permission_flags_we_emit() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_protection(Some(SheetProtection {
            protected: true,
            password_hash: Some(0xCAFE),
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: true,
            format_columns: true,
            sort: true,
            auto_filter: true,
            ..Default::default()
        }));

    let result = roundtrip_through_excel_xlsb(&wb);
    let protection = result
        .worksheet(0)
        .unwrap()
        .protection()
        .expect("sheet protection lost");
    assert!(protection.protected);
    assert_eq!(protection.password_hash, Some(0xCAFE));
    assert!(protection.select_locked_cells);
    assert!(protection.select_unlocked_cells);
    assert!(protection.format_cells);
    assert!(protection.format_columns);
    assert!(protection.sort);
    assert!(protection.auto_filter);
    assert!(!protection.format_rows);
    assert!(!protection.insert_rows);
    assert!(!protection.delete_columns);
    assert!(!protection.pivot_tables);
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_rich_text_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let red_big = RunFont {
        size: Some(16.0),
        color: Some(Color::Indexed(2)),
        ..Default::default()
    };
    ws.set_cell_value_at(
        0,
        0,
        CellValue::rich_text(vec![
            RichTextRun {
                text: "plain ".into(),
                font: None,
            },
            RichTextRun {
                text: "bold ".into(),
                font: Some(bold),
            },
            RichTextRun {
                text: "italic ".into(),
                font: Some(italic),
            },
            RichTextRun {
                text: "loud".into(),
                font: Some(red_big),
            },
        ]),
    )
    .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let value = s.get_value_at(0, 0);

    // Concatenated text must round-trip exactly.
    assert_eq!(format!("{value}"), "plain bold italic loud");

    // The cell must come back as RichText with at least the bold,
    // italic, and size+color runs distinguishable from the plain run.
    // If Excel collapsed the runs into a single span the formatting
    // is lost - the writer's RichText emission would be effectively
    // ineffective even though the text reads back correctly.
    let runs = match &value {
        CellValue::RichText(runs) => runs,
        other => panic!("expected RichText after Excel round-trip, got {other:?}"),
    };
    assert!(
        runs.len() >= 4,
        "expected at least 4 runs (plain/bold/italic/loud), got {}: {runs:?}",
        runs.len()
    );

    let has_bold = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.bold == Some(true)));
    let has_italic = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.italic == Some(true)));
    let has_big = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if matches!(f.size, Some(s) if s >= 14.0)));
    assert!(
        has_bold,
        "expected at least one run with bold=true: {runs:?}"
    );
    assert!(
        has_italic,
        "expected at least one run with italic=true: {runs:?}"
    );
    assert!(has_big, "expected at least one run with size>=14: {runs:?}");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_named_range_formulas_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 15.0)
        .unwrap();

    wb.define_name("Numbers", "Data!$A$1:$A$3").unwrap();
    wb.define_name("TaxRate", "0.1").unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .unwrap();
    calc.set_cell_formula("B2", "=B1*TaxRate").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(3.0))
        .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 30.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(30), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 3.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(3), got {other:?}"),
    }

    // Verify the named range names appear in the formula text — without
    // this, Excel could have inlined the named references and the test
    // would silently pass even if the named ranges were dropped.
    let f1 = s.get_formula_at(0, 1).expect("B1 still a formula");
    assert!(
        f1.contains("Numbers"),
        "named range Numbers lost from B1 formula: {f1:?}"
    );
    let f2 = s.get_formula_at(1, 1).expect("B2 still a formula");
    assert!(
        f2.contains("TaxRate"),
        "named range TaxRate lost from B2 formula: {f2:?}"
    );

    // Both named ranges must still be defined in the resulting
    // workbook (not just survive as formula text but no actual
    // definition).
    let names: Vec<&str> = result
        .named_ranges()
        .iter()
        .map(|n| n.name.as_str())
        .collect();
    assert!(
        names.contains(&"Numbers"),
        "named range 'Numbers' not in workbook after re-save: {names:?}"
    );
    assert!(
        names.contains(&"TaxRate"),
        "named range 'TaxRate' not in workbook after re-save: {names:?}"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_named_range_comment_we_emit() {
    use duke_sheets_core::named_range::{NameScope, NamedRange};

    let mut wb = Workbook::new();
    let nr = NamedRange::new("MyTax", "0.07", NameScope::Workbook).with_comment("Sales tax rate");
    wb.named_ranges_mut().define_or_update(nr);

    let result = roundtrip_through_excel_xlsb(&wb);
    let got = result
        .named_ranges()
        .iter()
        .find(|n| n.name == "MyTax")
        .expect("named range lost in Excel round-trip");
    assert_eq!(
        got.comment.as_deref(),
        Some("Sales tax rate"),
        "named range comment dropped by Excel: {:?}",
        got.comment
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_intersection_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    // 3x3 grid: A1=1 B1=2 C1=3 / A2=4 B2=5 C2=6 / A3=7 B3=8 C3=9.
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Intersection of A1:B3 with B2:C3 is the cells at B2, B3.
    // SUM = 5 + 8 = 13.
    ws.set_cell_formula("E1", "=SUM(A1:B3 B2:C3)").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!(
            (n - 13.0).abs() < 1e-9,
            "intersection sum drifted: E1 = {n}"
        ),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
    // Verify the intersection operator survives in the formula text —
    // Excel could otherwise have rewritten the formula to a value or
    // a different expression that happens to produce 13.
    let formula = s
        .get_formula_at(0, 4)
        .expect("E1 must still be a formula after Excel re-save");
    assert!(
        formula.contains("A1:B3") && formula.contains("B2:C3"),
        "intersection ranges lost from formula: {formula:?}"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_union_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Union of A1:A2 and C2:C3: {A1=1, A2=4, C2=6, C3=9} = 20.
    // Double parens are required so SUM treats the comma as union
    // rather than its argument separator.
    ws.set_cell_formula("E1", "=SUM((A1:A2,C2:C3))").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 20.0).abs() < 1e-9, "union sum drifted: E1 = {n}"),
        other => panic!("E1 expected Number(20), got {other:?}"),
    }
    // Verify the union operator survives. The formula should still
    // reference both A1:A2 and C2:C3 — otherwise Excel could have
    // inlined or rewritten the union.
    let formula = s
        .get_formula_at(0, 4)
        .expect("E1 must still be a formula after Excel re-save");
    assert!(
        formula.contains("A1:A2") && formula.contains("C2:C3"),
        "union ranges lost from formula: {formula:?}"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_print_names_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_cell_value("B2", 100.0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "B2"));
    ps.repeat_rows = Some((0, 0));
    ws.set_page_setup(ps);

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("header"));
    let print_area = s
        .page_setup()
        .print_area
        .as_ref()
        .expect("print_area must survive Excel round-trip");
    assert_eq!(print_area.start, CellAddress::parse("A1").unwrap());
    assert_eq!(print_area.end, CellAddress::parse("B2").unwrap());

    // repeat_rows = (0, 0) means "repeat row 1 at top of each printed
    // page". This was set on the input but previously not asserted —
    // a writer regression that dropped repeat_rows would have gone
    // unnoticed.
    let repeat_rows = s
        .page_setup()
        .repeat_rows
        .expect("repeat_rows must survive Excel round-trip");
    assert_eq!(
        repeat_rows,
        (0, 0),
        "repeat_rows lost or mangled: {:?}",
        repeat_rows
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_visual_state_we_emit() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").unwrap();
    wb.add_worksheet_with_name("Hidden").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "p")
        .unwrap();
    wb.worksheet_mut(0).unwrap().set_freeze_panes(1, 1);
    wb.worksheet_mut(0).unwrap().set_zoom_scale(Some(125));
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);
    wb.set_active_sheet(0).unwrap();

    let mut ps = wb.worksheet(0).unwrap().page_setup().clone();
    ps.orientation = PageOrientation::Landscape;
    ps.left_margin = 0.5;
    ps.odd_header = Some("Hdr".into());
    ps.print_gridlines = true;
    wb.worksheet_mut(0).unwrap().set_page_setup(ps);

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet_by_name("Public").unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("p"));

    let freeze = s.freeze_panes().expect("freeze panes must survive");
    assert_eq!(
        (freeze.row, freeze.col),
        (1, 1),
        "freeze panes lost after round-trip"
    );

    assert_eq!(
        s.zoom_scale(),
        Some(125),
        "zoom scale lost after round-trip"
    );

    let hidden = result.worksheet_by_name("Hidden").unwrap();
    assert!(
        hidden.visibility() != SheetVisibility::Visible,
        "Hidden sheet must remain non-visible after round-trip; got {:?}",
        hidden.visibility()
    );

    let ps = s.page_setup();
    assert_eq!(
        ps.orientation,
        PageOrientation::Landscape,
        "landscape orientation lost"
    );
    assert!(
        ps.print_gridlines,
        "print_gridlines flag lost after round-trip"
    );
    let header = ps
        .odd_header
        .as_deref()
        .expect("odd_header text lost after round-trip");
    assert!(
        header.contains("Hdr"),
        "odd_header content mangled: {header:?}"
    );
    // left_margin was set on input but previously not asserted.
    // Excel may round slightly (inches), so allow ±0.05 inch.
    assert!(
        (ps.left_margin - 0.5).abs() < 0.05,
        "left_margin lost or mangled: {} (expected ~0.5)",
        ps.left_margin
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_protection_state_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "locked").unwrap();
    ws.set_protection(Some(SheetProtection {
        protected: true,
        password_hash: Some(0xCAFE),
        ..Default::default()
    }));

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("locked"));
    let prot = s
        .protection()
        .expect("sheet protection must survive Excel round-trip");
    assert!(prot.protected, "protected flag lost after round-trip");
    // Excel may rewrite the hash with its own algorithm; require a
    // non-zero hash but don't require byte-identity. The pre-fix
    // writer was emitting no BrtSheetProtection at all so any
    // password presence proves the record is making it through.
    let hash = prot
        .password_hash
        .expect("password_hash lost after round-trip");
    assert_ne!(
        hash, 0,
        "password_hash dropped to 0 (pre-fix writer behaviour)"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_workbook_protection_and_protected_ranges_we_emit() {
    use duke_sheets_core::{ProtectedRange, WorkbookProtection};

    let mut wb = Workbook::new();
    wb.set_workbook_protection(Some(WorkbookProtection {
        structure: true,
        windows: true,
        password_hash: Some(0xCAFE),
    }));

    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "editable").unwrap();
    ws.set_protected_ranges(vec![ProtectedRange {
        name: "Editable".to_string(),
        ranges: vec![range("A1", "B2"), range("D4", "D5")],
        password_hash: Some(0xCAFE),
        security_descriptor: None,
    }]);

    let result = roundtrip_through_excel_xlsb(&wb);
    let protection = result
        .workbook_protection()
        .expect("workbook protection must survive Excel round-trip");
    assert!(protection.structure, "workbook structure protection lost");
    // Excel opens fLockWindow in XLSB, but current Excel re-saves the
    // workbook with only fLockStructure set. The in-process XLSB round-trip
    // pins byte-level preservation of the window bit; this parity test pins
    // the portion Excel itself preserves.
    assert!(
        protection.password_hash.unwrap_or_default() != 0,
        "workbook password hash lost"
    );

    let ranges = result.worksheet(0).unwrap().protected_ranges();
    assert_eq!(ranges.len(), 1, "protected range count changed");
    assert_eq!(ranges[0].name, "Editable");
    assert_eq!(ranges[0].ranges[0].to_string(), "A1:B2");
    assert_eq!(ranges[0].ranges[1].to_string(), "D4:D5");
    assert!(
        ranges[0].password_hash.unwrap_or_default() != 0,
        "protected range password hash lost"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_dimensions_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "tall").unwrap();
    ws.set_cell_value("B1", "wide").unwrap();
    ws.set_row_height(0, 36.0);
    ws.set_column_width(1, 25.0);
    // Non-default sheet-wide defaults: 22pt row height, 12 char col
    // width. These travel via BrtWsFmtInfo and should survive Excel.
    ws.set_default_row_height(22.0);
    ws.set_default_column_width(12.0);

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("tall"));
    assert_eq!(s.get_value_at(0, 1).as_string(), Some("wide"));
    // Row 0 should retain a non-default height. Excel may round or
    // adjust slightly, so allow ±1 pt slack.
    let h = s.row_height(0);
    assert!((h - 36.0).abs() < 1.0, "row 0 height = {h} (expected ~36)",);
    // Column B (index 1) should retain a non-default width.
    let w = s.column_width(1);
    assert!(w > 20.0 && w < 30.0, "col B width = {w} (expected ~25)",);
    // BrtWsFmtInfo carries miyDefRwHeight; the fUnsynced flag tells
    // Excel the row height is user-set, otherwise Excel resets to 15pt
    // on save. Verified end-to-end here.
    let dh = s.default_row_height();
    assert!(
        (dh - 22.0).abs() < 1.0,
        "default row height = {dh} (expected ~22)"
    );
    // Note: cchDefColWidth in BrtWsFmtInfo does NOT carry custom
    // default column widths through Excel — Excel rewrites it to the
    // base font width (8) regardless of input. Per-column custom
    // widths via BrtColInfo are how Excel persists those. Don't
    // assert default_column_width survival here.
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_table_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Product").unwrap();
    ws.set_cell_value("B1", "Price").unwrap();
    ws.set_cell_value("A2", "Widget").unwrap();
    ws.set_cell_value("B2", 9.99).unwrap();
    ws.set_cell_value("A3", "Gadget").unwrap();
    ws.set_cell_value("B3", 19.99).unwrap();
    ws.add_table(Table {
        id: 1,
        name: "Products".to_string(),
        display_name: "Products".to_string(),
        reference: range("A1", "B3"),
        columns: vec![
            TableColumn {
                id: 1,
                name: "Product".to_string(),
                totals_row_function: None,
                totals_row_formula: None,
                totals_row_label: None,
                calculated_column_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Price".to_string(),
                totals_row_function: None,
                totals_row_formula: None,
                totals_row_label: None,
                calculated_column_formula: None,
            },
        ],
        style_info: Some(TableStyleInfo {
            name: Some("TableStyleMedium2".to_string()),
            show_first_column: true,
            show_last_column: true,
            show_row_stripes: true,
            show_column_stripes: true,
        }),
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: true,
    });

    let result = roundtrip_through_excel_xlsb(&wb);
    let s = result.worksheet(0).unwrap();
    let tables = s.tables();
    assert_eq!(tables.len(), 1, "table must survive Excel round-trip");
    assert_eq!(tables[0].name, "Products", "table name lost");
    assert_eq!(tables[0].columns.len(), 2, "column count lost");
    assert_eq!(
        tables[0].columns[0].name, "Product",
        "first column name lost"
    );
    assert_eq!(
        tables[0].columns[1].name, "Price",
        "second column name lost"
    );
    assert_eq!(
        tables[0].reference,
        range("A1", "B3"),
        "table reference lost"
    );

    // style_info and header_row_count were set on the input but
    // previously not asserted. A writer regression that dropped the
    // style metadata or header row count would have gone unnoticed.
    let style = tables[0]
        .style_info
        .as_ref()
        .expect("style_info must survive Excel round-trip");
    assert!(
        style
            .name
            .as_deref()
            .map(|s| s.contains("TableStyleMedium"))
            .unwrap_or(false),
        "table style name lost: {:?}",
        style.name
    );
    assert!(
        style.show_row_stripes,
        "show_row_stripes flag lost after round-trip"
    );
    assert!(
        style.show_column_stripes,
        "show_column_stripes flag lost after round-trip"
    );
    assert!(
        style.show_first_column,
        "show_first_column flag lost after round-trip"
    );
    assert!(
        style.show_last_column,
        "show_last_column flag lost after round-trip"
    );
    assert_eq!(
        tables[0].header_row_count, 1,
        "header_row_count must survive Excel round-trip"
    );
}

/// Form controls survive the Excel XLSB round-trip. In XLSB the
/// controls live entirely in the legacy VML part referenced by
/// BrtLegacyDrawing; the Repaired check also pins the corrected
/// record id (0x0227).
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsb_form_controls_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::table::{Table, TableColumn};
    use duke_sheets_core::CellRange;
    use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

    let anchor = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value_at(0, 0, 42.0).expect("A1");
    for (i, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        ws.set_cell_value_at(i as u32, 7, *item).expect("list item");
    }
    ws.set_cell_value_at(0, 9, "Name").expect("J1");
    ws.set_cell_value_at(1, 9, "Alice").expect("J2");
    ws.add_table(Table {
        id: 1,
        name: "People".to_string(),
        display_name: "People".to_string(),
        reference: CellRange::parse("J1:J2").unwrap(),
        columns: vec![TableColumn {
            id: 1,
            name: "Name".to_string(),
            totals_row_function: None,
            totals_row_formula: None,
            totals_row_label: None,
            calculated_column_formula: None,
        }],
        style_info: None,
        header_row_count: 1,
        totals_row_count: 0,
        totals_row_shown: false,
    });

    let kinds: Vec<FormControlKind> = vec![
        FormControlKind::Button {
            caption: "Run Report".into(),
        },
        FormControlKind::Checkbox {
            caption: "Enable audit".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: true,
        },
        FormControlKind::OptionButton {
            caption: "Opt A".into(),
            state: CheckState::Checked,
            cell_link: None,
            first_in_group: false,
            no_3d: true,
        },
        FormControlKind::Label {
            caption: "Status".into(),
        },
        FormControlKind::GroupBox {
            caption: "Choices".into(),
            no_3d: true,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![1, 3],
            no_3d: true,
        },
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$4".to_string()),
            selected: Some(2),
            lines: 6,
            no_3d: true,
        },
        FormControlKind::Scrollbar {
            value: 40,
            min: 5,
            max: 95,
            increment: 2,
            page: 10,
            horizontal: false,
            cell_link: Some("$D$6".to_string()),
        },
        FormControlKind::Spinner {
            value: 12,
            min: 0,
            max: 30,
            increment: 3,
            cell_link: None,
        },
    ];
    let count = kinds.len();
    let expected = kinds.clone();
    for (i, kind) in kinds.into_iter().enumerate() {
        let row = 1 + 2 * i as u32;
        ws.add_form_control(FormControl::new(kind), anchor(1, row, 3, row + 1)).unwrap();
    }
    assert_eq!(wb.sync_form_control_links(), 3);

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.get_value("D2").unwrap(), CellValue::Boolean(true));
    assert_eq!(sheet.get_value("D4").unwrap(), CellValue::Number(3.0));
    assert_eq!(sheet.get_value("D6").unwrap(), CellValue::Number(40.0));
    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), count, "every control survives Excel");
    for (i, control) in controls.iter().enumerate() {
        let mut want = expected[i].clone();
        if let FormControlKind::OptionButton { first_in_group, .. } = &mut want {
            *first_in_group = true;
        }
        assert_eq!(
            control.payload.kind, want,
            "control {i} kind mismatch after Excel"
        );
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_custom_metric_control_anchor_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{CheckState, FormControl, FormControlKind};

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "metric anchor".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        }),
        DrawingAnchor::OneCell {
            from: CellMarker::default(),
            width_emu: 609_600,
            height_emu: 190_500,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    match &drawn.object.unwrap().anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 0));
            assert_eq!((from.row, from.row_offset_emu), (0, 0));
            assert_eq!((to.col, to.col_offset_emu), (0, 609_600));
            assert_eq!((to.row, to.row_offset_emu), (0, 190_500));
        }
        other => panic!("expected Excel-resaved TwoCell control anchor, got {other:?}"),
    }
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_control_visual_metadata_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::style::{HorizontalAlignment, Underline, VerticalAlignment};
    use duke_sheets_core::{CheckState, ControlText, DrawingObject, FormControl, FormControlKind};

    let text = ControlText {
        runs: vec![
            RichTextRun::with_font(
                "Red ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(9.0),
                    color: Some(Color::rgb(255, 0, 0)),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Blue",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    color: Some(Color::rgb(0, 0, 255)),
                    italic: Some(true),
                    underline: Some(Underline::Single),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Right),
        vertical_alignment: Some(VerticalAlignment::Bottom),
    };
    let control = FormControl::new(FormControlKind::Checkbox {
        caption: text,
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    })
    .with_macro_name("RunProbe");
    let mut object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 0,
            row: 1,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 4,
            col_offset_emu: 0,
            row: 3,
            row_offset_emu: 0,
        },
        edit_as: None,
    });
    object.meta.name = Some("Visual Probe".into());
    object.meta.alt_text = Some("Visual probe alternative".into());
    object.meta.title = Some("Visual probe title".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel_xlsb(&workbook);
    let drawn = result.worksheet(0).unwrap().form_controls().next().unwrap();
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Visual Probe"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("Visual probe alternative")
    );
    assert_eq!(
        drawn.object.unwrap().meta.title.as_deref(),
        Some("Visual probe title")
    );
    assert_eq!(drawn.payload.caption_text().as_deref(), Some("Red Blue"));
    assert_eq!(drawn.payload.macro_name.as_deref(), Some("RunProbe"));
    let caption = drawn.payload.caption().unwrap();
    assert_eq!(
        caption.horizontal_alignment,
        Some(HorizontalAlignment::Right)
    );
    assert_eq!(caption.vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(caption.runs.len(), 2);
    let red = caption.runs[0].font.as_ref().unwrap();
    assert_eq!(red.name.as_deref(), Some("Segoe UI"));
    assert_eq!(red.size, Some(9.0));
    assert_eq!(red.color, Some(Color::rgb(255, 0, 0)));
    assert_eq!(red.bold, Some(true));
    let blue = caption.runs[1].font.as_ref().unwrap();
    assert_eq!(blue.name.as_deref(), Some("Arial"));
    assert_eq!(blue.size, Some(12.0));
    assert_eq!(blue.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(blue.italic, Some(true));
    assert_eq!(blue.underline, Some(Underline::Single));
}

const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

/// A cell comment survives the Excel XLSB round-trip: the comments
/// part must use the MS-XLSB record ids (the old 0x0278-based emit
/// made Excel refuse the file outright).
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsb_comment_we_emit() {
    use duke_sheets_core::CellComment;

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("B2", "has note").unwrap();
    ws.set_comment_at(1, 1, CellComment::new("Reviewer", "Check this figure")).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();
    let comment = sheet.comment_at(1, 1).expect("comment survives at B2");
    assert!(
        comment.text.contains("Check this figure"),
        "comment text lost: {:?}",
        comment.text
    );
    assert!(
        comment.author.contains("Reviewer"),
        "comment author lost: {:?}",
        comment.author
    );
    assert!(sheet.comment_at(0, 0).is_none(), "comment cell moved");
}

/// A PNG picture survives the Excel XLSB round-trip with its bytes
/// verbatim (Excel re-packages media parts without re-encoding).
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsb_png_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(
        EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 2_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        },
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 2,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 5,
                col_offset_emu: 0,
                row: 10,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 1, "image must survive Excel re-save");
    let img = &images[0];
    assert_eq!(img.payload.format, ImageFormat::Png);
    assert_eq!(
        img.payload.data, TEST_PNG_1X1,
        "PNG bytes must round-trip through Excel verbatim"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_one_cell_and_absolute_images_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::DrawingObject;

    let image = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_200_000,
            height_emu: 700_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(
        image("OneCell").with_anchor(DrawingAnchor::OneCell {
            from: CellMarker {
                col: 1,
                row: 2,
                col_offset_emu: 95_250,
                row_offset_emu: 47_625,
            },
            width_emu: 1_200_000,
            height_emu: 700_000,
        }),
    ).unwrap();
    ws.add_drawing(
        image("Absolute").with_anchor(DrawingAnchor::Absolute {
            x_emu: 2_000_000,
            y_emu: 1_000_000,
            width_emu: 900_000,
            height_emu: 500_000,
        }),
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let images: Vec<_> = result.worksheet(0).unwrap().images().collect();
    assert_eq!(images.len(), 2);
    assert_eq!(
        images[0].object.unwrap().anchor,
        DrawingAnchor::OneCell {
            from: CellMarker {
                col: 1,
                row: 2,
                col_offset_emu: 95_250,
                row_offset_emu: 47_625,
            },
            width_emu: 1_200_000,
            height_emu: 700_000,
        }
    );
    assert_eq!(
        images[1].object.unwrap().anchor,
        DrawingAnchor::Absolute {
            x_emu: 2_000_000,
            y_emu: 1_000_000,
            width_emu: 900_000,
            height_emu: 500_000,
        }
    );
}

/// A model-authored chart survives the Excel XLSB round-trip (the
/// old writer never emitted model charts at all, leaving a dangling
/// BrtDrawing pointer).
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsb_chart_we_emit() {
    use duke_sheets_chart::{
        CellMarker, Chart, ChartType, DataReference, DataSeries, DrawingAnchor,
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (i, v) in [3.0, 1.0, 4.0, 1.0, 5.0].iter().enumerate() {
        ws.set_cell_value_at(i as u32, 0, *v).unwrap();
    }
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Sales".to_string());
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$A$1:$A$5")));
    ws.add_chart(
        chart,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 2,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 10,
                col_offset_emu: 0,
                row: 17,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();
    assert_eq!(sheet.chart_count(), 1, "chart must survive Excel re-save");
    let chart = sheet.charts().next().unwrap().payload;
    assert_eq!(chart.chart_type, ChartType::ColumnClustered);
    assert_eq!(chart.title.as_deref(), Some("Sales"));
    assert_eq!(chart.series.len(), 1);
}

/// The drawing-list z-order (image below a form control below an
/// image) survives Excel's XLSB re-save: the control's position among
/// native shapes rides its com14:compatSp placeholder twin, which
/// Excel keeps in the drawing part's document order.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_drawing_z_order_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingKind, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Below").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Middle".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
        .with_anchor(two_cell(1, 1, 3, 3)),
    ).unwrap();
    ws.add_drawing(png("Above").with_anchor(two_cell(2, 2, 4, 4))).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();
    let tags: Vec<&str> = sheet
        .drawings()
        .iter()
        .map(|object| match &object.kind {
            DrawingKind::Image(_) => "image",
            DrawingKind::FormControl(_) => "control",
            other => panic!("unexpected drawing kind after Excel round-trip: {other:?}"),
        })
        .collect();
    assert_eq!(
        tags,
        vec!["image", "control", "image"],
        "z-order must survive Excel XLSB re-save"
    );
    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images[0].object.unwrap().meta.name.as_deref(), Some("Below"));
    assert_eq!(images[1].object.unwrap().meta.name.as_deref(), Some("Above"));
    assert_eq!(
        sheet
            .form_controls()
            .next()
            .unwrap()
            .payload
            .caption_text()
            .as_deref(),
        Some("Middle")
    );
}

/// Drawing-object hidden flags survive Excel's XLSB re-save: a
/// hidden image rides its `cNvPr@hidden="1"` in the shared drawing
/// part, a hidden form control rides the VML shape's
/// `visibility:hidden` style, and the visible siblings stay visible.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_hidden_drawing_flags_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};
    use duke_sheets_core::{CheckState, DrawingObject, FormControl, FormControlKind};

    let two_cell = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };
    let png = |name: &str| {
        DrawingObject::image(EmbeddedImage {
            format: ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 300_000,
            height_emu: 300_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_name(name)
    };
    let checkbox = |caption: &str| {
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: caption.into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }))
    };

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_drawing(png("Shown").with_anchor(two_cell(0, 0, 2, 2))).unwrap();
    ws.add_drawing(
        png("Ghost")
            .with_anchor(two_cell(2, 2, 4, 4))
            .with_hidden(true),
    ).unwrap();
    ws.add_drawing(checkbox("Visible box").with_anchor(two_cell(4, 4, 6, 6))).unwrap();
    ws.add_drawing(
        checkbox("Cloaked box")
            .with_anchor(two_cell(6, 6, 8, 8))
            .with_hidden(true),
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let sheet = result.worksheet(0).unwrap();

    let images: Vec<_> = sheet.images().collect();
    assert_eq!(images.len(), 2, "both images survive Excel re-save");
    let image_hidden = |name: &str| {
        images
            .iter()
            .find(|i| i.object.unwrap().meta.name.as_deref() == Some(name))
            .unwrap_or_else(|| panic!("image {name:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(!image_hidden("Shown"), "visible image must stay visible");
    assert!(
        image_hidden("Ghost"),
        "hidden image must survive Excel re-save with hidden intact"
    );

    let controls: Vec<_> = sheet.form_controls().collect();
    assert_eq!(controls.len(), 2, "both controls survive Excel re-save");
    let control_hidden = |caption: &str| {
        controls
            .iter()
            .find(|c| c.payload.caption_text().as_deref() == Some(caption))
            .unwrap_or_else(|| panic!("control {caption:?} lost in Excel re-save"))
            .meta
            .hidden
    };
    assert!(
        !control_hidden("Visible box"),
        "visible control must stay visible"
    );
    assert!(
        control_hidden("Cloaked box"),
        "hidden control must survive Excel re-save with hidden intact"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_xlsb_basic_shape_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::style::{HorizontalAlignment, VerticalAlignment};
    use duke_sheets_core::{
        DrawingObject, DrawingText, Shape, ShapeFill, ShapeGeometry, ShapeLine,
    };

    let text = DrawingText {
        runs: vec![
            RichTextRun::with_font(
                "Bold ",
                RunFont {
                    name: Some("Segoe UI".into()),
                    size: Some(10.0),
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
            RichTextRun::with_font(
                "Italic",
                RunFont {
                    name: Some("Arial".into()),
                    size: Some(12.0),
                    italic: Some(true),
                    color: Some(Color::rgb(0, 0, 255)),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: Some(HorizontalAlignment::Center),
        vertical_alignment: Some(VerticalAlignment::Center),
    };
    let shape = Shape::rectangle()
        .with_fill(ShapeFill::Solid(Color::rgb(255, 0, 0)))
        .with_line(ShapeLine {
            color: Some(Color::rgb(0, 0, 255)),
            width_emu: Some(25_400),
            dash_style: Some("dash".into()),
            no_fill: false,
        })
        .with_text(text)
        .with_rotation(900_000)
        .with_flip_h(true);
    let mut object = DrawingObject::shape(shape).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            row: 2,
            ..CellMarker::default()
        },
        to: CellMarker {
            col: 5,
            row: 8,
            ..CellMarker::default()
        },
        edit_as: None,
    });
    object.meta.name = Some("Status panel".into());
    object.meta.alt_text = Some("red status rectangle".into());
    object.meta.title = Some("Status".into());
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let result = roundtrip_through_excel_xlsb(&workbook);
    let drawn = result.worksheet(0).unwrap().shapes().next().expect("shape");
    assert_eq!(drawn.object.unwrap().meta.name.as_deref(), Some("Status panel"));
    assert_eq!(
        drawn.object.unwrap().meta.alt_text.as_deref(),
        Some("red status rectangle")
    );
    assert_eq!(drawn.object.unwrap().meta.title.as_deref(), Some("Status"));
    assert_eq!(drawn.payload.geometry, ShapeGeometry::Preset("rect".into()));
    assert_eq!(drawn.payload.fill, ShapeFill::Solid(Color::rgb(255, 0, 0)));
    assert_eq!(drawn.payload.line.color, Some(Color::rgb(0, 0, 255)));
    assert_eq!(drawn.payload.line.width_emu, Some(25_400));
    assert_eq!(drawn.payload.line.dash_style.as_deref(), Some("dash"));
    assert_eq!(drawn.payload.rotation, 900_000);
    assert!(drawn.payload.flip_h);
    let text = drawn.payload.text.as_ref().expect("shape text");
    assert_eq!(text.plain_text(), "Bold Italic");
    assert_eq!(text.horizontal_alignment, Some(HorizontalAlignment::Center));
    assert_eq!(text.vertical_alignment, Some(VerticalAlignment::Center));
    assert_eq!(
        text.runs[0].font.as_ref().unwrap().name.as_deref(),
        Some("Segoe UI")
    );
    assert_eq!(text.runs[0].font.as_ref().unwrap().bold, Some(true));
    assert_eq!(text.runs[0].font.as_ref().unwrap().size, Some(10.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().name.as_deref(),
        Some("Arial")
    );
    assert_eq!(text.runs[1].font.as_ref().unwrap().italic, Some(true));
    assert_eq!(text.runs[1].font.as_ref().unwrap().size, Some(12.0));
    assert_eq!(
        text.runs[1].font.as_ref().unwrap().color,
        Some(Color::rgb(0, 0, 255))
    );
}

/// Unmodeled ClientData children we replay on a modeled control kind
/// survive a real Excel XLSB round-trip. Dialog-button semantics
/// (`x:Default`, `x:Cancel`) ride a worksheet button; the Repaired
/// check inside the helper proves the raw emission is spec-clean and
/// the re-read model proves Excel re-saved the elements.
// features: Form control unmodeled ClientData passthrough
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_preserves_unmodeled_client_data_we_emit_xlsb() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor};
    use duke_sheets_core::{FormControl, FormControlKind};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 1.0).expect("A1");
    let mut control = FormControl::new(FormControlKind::Button {
        caption: "OK".into(),
    });
    control.raw_client_data = vec![b"<x:Default/>".to_vec(), b"<x:Cancel/>".to_vec()];
    ws.add_form_control(
        control,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 1,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    ).unwrap();

    let result = roundtrip_through_excel_xlsb(&wb);
    let control = result
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("button survives");
    assert!(matches!(
        control.payload.kind,
        duke_sheets_core::FormControlKind::Button { .. }
    ));
    let raws: Vec<String> = control
        .payload
        .raw_client_data
        .iter()
        .map(|raw| String::from_utf8_lossy(raw).into_owned())
        .collect();
    assert!(
        raws.iter().any(|raw| raw.contains("Default")),
        "x:Default survives Excel's re-save: {raws:?}"
    );
    assert!(
        raws.iter().any(|raw| raw.contains("Cancel")),
        "x:Cancel survives Excel's re-save: {raws:?}"
    );
}
