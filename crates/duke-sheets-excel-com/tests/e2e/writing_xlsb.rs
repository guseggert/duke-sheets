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
use duke_sheets_core::{
    CellAddress, CellRange, CellValue, Hyperlink, PivotAggregate, PivotTable, Workbook,
};

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

fn xlsb_basic_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Region").unwrap();
    ws.set_cell_value("B1", "Revenue").unwrap();
    ws.set_cell_value("A2", "East").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "West").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();

    let pivot = PivotTable::builder("BasicPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    ws.add_pivot_table(pivot).unwrap();
    wb
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_opens_xlsb_with_native_pivot_table() {
    let (_result, _writer_bytes, excel_bytes) =
        roundtrip_through_excel_xlsb_bytes(&xlsb_basic_pivot_workbook());

    assert!(zip_has_entry(&excel_bytes, "xl/pivotTables/pivotTable1.bin"));
    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotCache/pivotCacheDefinition1.bin"
    ));
    assert!(zip_has_entry(
        &excel_bytes,
        "xl/pivotCache/pivotCacheRecords1.bin"
    ));
}

fn zip_has_entry(bytes: &[u8], name: &str) -> bool {
    let reader = std::io::Cursor::new(bytes);
    let Ok(mut archive) = zip::ZipArchive::new(reader) else {
        return false;
    };
    let found = archive.by_name(name).is_ok();
    found
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
    ws.set_formula_result(0, 1, CellValue::Number(42.0)).unwrap();

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
    assert_eq!(accepted.len(), formulas.len(), "all ATP formulas should be authored");
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
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
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
    assert_eq!(
        tables[0].header_row_count, 1,
        "header_row_count must survive Excel round-trip"
    );
}
