//! E2E tests for freeze panes and multi-selection roundtrip via Excel COM.
//!
//! Test 1 (writer): Write XLSX with freeze panes + multi-selection using our
//! writer → open in Excel → verify COM properties → re-save → read back.
//!
//! Test 2 (reader): Create workbook in Excel via COM with freeze panes and
//! a selection → save → read with our reader → verify data model.

use duke_sheets_core::Selection;
use duke_sheets_excel_com::ChainStep;
use excel_com_protocol::ResponseData;

/// Helper: extract a serde_json::Value from a bridge get/invoke result.
fn extract_value(data: Option<ResponseData>) -> serde_json::Value {
    match data {
        Some(ResponseData::Value { value }) => value,
        other => panic!("expected Value, got {other:?}"),
    }
}

// =========================================================================
// Writer E2E: our XLSX → Excel → verify COM → re-save → read back
// =========================================================================

/// Write a workbook with freeze panes and two selections (topLeft + bottomRight),
/// open in Excel, verify Excel sees freeze panes and the active selection,
/// re-save, and confirm our reader recovers both selections.
#[test]
fn test_write_freeze_panes_and_multi_selection() {
    use duke_sheets_xlsx::{XlsxReader, XlsxWriter};
    use std::io::Cursor;

    // 1. Build workbook with our API
    let mut wb = duke_sheets_core::Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Header A").unwrap();
    sheet.set_cell_value("B1", "Header B").unwrap();
    sheet.set_cell_value("A2", 100.0).unwrap();
    sheet.set_cell_value("B2", 200.0).unwrap();

    // Freeze at row 2, col 1 (first unfrozen row = 2, first unfrozen col = 1)
    // → rows 0-1 frozen, col 0 frozen
    sheet.set_freeze_panes(2, 1);

    // Two selections: one per pane
    sheet.set_selections(vec![
        Selection {
            pane: Some("topLeft".to_string()),
            active_cell: Some("A1".to_string()),
            sqref: Some("A1".to_string()),
        },
        Selection {
            pane: Some("bottomRight".to_string()),
            active_cell: Some("C5".to_string()),
            sqref: Some("C5:D8".to_string()),
        },
    ]);

    // 2. Write XLSX with our writer
    let input = crate::temp_fixture();
    let output = crate::temp_fixture();

    let mut buf = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut buf)).expect("XlsxWriter::write");
    std::fs::write(&input.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", input.host_path.display()));

    // 3. Push to VM and open in Excel
    crate::ensure_vm_temp_dir();
    crate::push_file_to_vm(&input);

    let bridge = crate::excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&input.vm_path)
        .expect("Excel should open our file");

    // Assert no repair (Excel silently repairs bad XML)
    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the file! Name: {wb_name}"
    );
    let read_only = opened.is_read_only().expect("ReadOnly");
    assert!(!read_only, "Excel opened as read-only (possible repair)");

    // 4. Verify Excel sees freeze panes via COM
    let aw = || vec![ChainStep::Property("ActiveWindow".to_string())];

    let freeze = extract_value(excel.get(0, aw(), "FreezePanes").expect("get FreezePanes"));
    assert_eq!(freeze.as_bool(), Some(true), "FreezePanes should be true");

    let split_row = extract_value(excel.get(0, aw(), "SplitRow").expect("get SplitRow"));
    assert_eq!(
        split_row.as_f64().map(|v| v as i64),
        Some(2),
        "SplitRow should be 2 (2 frozen rows)"
    );

    let split_col = extract_value(excel.get(0, aw(), "SplitColumn").expect("get SplitColumn"));
    assert_eq!(
        split_col.as_f64().map(|v| v as i64),
        Some(1),
        "SplitColumn should be 1 (1 frozen column)"
    );

    // Pane count: row+col freeze → 4 panes
    let pane_count = extract_value(
        excel
            .get(
                0,
                vec![
                    ChainStep::Property("ActiveWindow".to_string()),
                    ChainStep::Property("Panes".to_string()),
                ],
                "Count",
            )
            .expect("get Panes.Count"),
    );
    assert_eq!(
        pane_count.as_f64().map(|v| v as i64),
        Some(4),
        "Should have 4 panes (row + col freeze)"
    );

    // Active pane selection should be C5:D8 (bottomRight pane)
    let sel_addr = extract_value(
        excel
            .get(
                0,
                vec![
                    ChainStep::Property("ActiveWindow".to_string()),
                    ChainStep::Property("Selection".to_string()),
                ],
                "Address",
            )
            .expect("get Selection.Address"),
    );
    // Excel returns absolute addresses ($C$5:$D$8), strip $ for comparison
    let sel_norm = sel_addr.as_str().unwrap_or("").replace('$', "");
    assert!(
        sel_norm.contains("C5") && sel_norm.contains("D8"),
        "Active selection should be C5:D8, got: {sel_norm}"
    );

    // Active cell should be C5
    let active_cell = extract_value(
        excel
            .get(
                0,
                vec![
                    ChainStep::Property("ActiveWindow".to_string()),
                    ChainStep::Property("ActiveCell".to_string()),
                ],
                "Address",
            )
            .expect("get ActiveCell.Address"),
    );
    let ac_norm = active_cell.as_str().unwrap_or("").replace('$', "");
    assert!(
        ac_norm.contains("C5"),
        "ActiveCell should be C5, got: {ac_norm}"
    );

    // 5. Re-save (Excel normalises the XML)
    opened.save(&output.vm_path).expect("Excel save");
    opened.close().expect("close workbook");
    drop(excel);

    // 6. Read back with our reader and verify selections survive
    crate::pull_file_from_vm(&output);
    let result = XlsxReader::read_file(&output.host_path).expect("XlsxReader::read_file");
    let s = result.worksheet(0).expect("worksheet 0");

    // Freeze panes preserved
    let freeze = s
        .freeze_panes()
        .expect("freeze panes should survive roundtrip");
    assert_eq!(freeze.row, 2, "freeze row");
    assert_eq!(freeze.col, 1, "freeze col");

    // Selections preserved (Excel should keep all pane selections)
    let sels = s.selections();
    assert!(
        !sels.is_empty(),
        "selections should survive Excel roundtrip"
    );

    // The bottomRight selection should be present with C5:D8
    let br_sel = sels
        .iter()
        .find(|s| s.pane.as_deref() == Some("bottomRight"));
    assert!(
        br_sel.is_some(),
        "bottomRight selection should survive, got: {sels:?}"
    );
    let br = br_sel.unwrap();
    assert!(
        br.sqref.as_deref().unwrap_or("").contains("C5"),
        "bottomRight sqref should contain C5, got: {:?}",
        br.sqref
    );

    crate::cleanup_fixture(&input);
    crate::cleanup_fixture(&output);
}

// =========================================================================
// Reader E2E: create freeze panes + selection in Excel → read with our code
// =========================================================================

/// Create a workbook in Excel with freeze panes and a selection, save it,
/// and verify our reader correctly parses the freeze pane state and
/// selection metadata.
#[test]
fn test_read_freeze_panes_and_selection_from_excel() {
    let bridge = crate::excel_bridge();
    let fixture = crate::temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        crate::ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Add data
        wb.set_cell_value("A1", "Frozen header A").unwrap();
        wb.set_cell_value("B1", "Frozen header B").unwrap();
        wb.set_cell_value("C1", "Frozen header C").unwrap();
        wb.set_cell_value("A2", 10.0).unwrap();
        wb.set_cell_value("B2", 20.0).unwrap();
        wb.set_cell_value("C3", "Data").unwrap();

        // Select cell C3 — freeze boundary will be at C3
        // (rows 1-2 frozen, columns A-B frozen)
        excel
            .invoke(
                0,
                vec![
                    ChainStep::Property("ActiveSheet".to_string()),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("C3")),
                ],
                "Select",
                vec![],
            )
            .expect("select C3");

        // Freeze panes at the current selection
        excel
            .set(
                0,
                vec![ChainStep::Property("ActiveWindow".to_string())],
                "FreezePanes",
                serde_json::Value::from(true),
            )
            .expect("set FreezePanes");

        // Now select a range in the active (bottom-right) pane
        excel
            .invoke(
                0,
                vec![
                    ChainStep::Property("ActiveSheet".to_string()),
                    ChainStep::Indexed("Range".to_string(), serde_json::Value::from("D5:F8")),
                ],
                "Select",
                vec![],
            )
            .expect("select D5:F8");

        // Verify freeze is active before saving
        let freeze_check = extract_value(
            excel
                .get(
                    0,
                    vec![ChainStep::Property("ActiveWindow".to_string())],
                    "FreezePanes",
                )
                .expect("check FreezePanes"),
        );
        assert_eq!(
            freeze_check.as_bool(),
            Some(true),
            "FreezePanes should be true before save"
        );

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    // Read back with our reader
    crate::pull_file_from_vm(&fixture);
    let workbook = duke_sheets_xlsx::XlsxReader::read_file(&fixture.host_path).expect("XlsxReader");
    let sheet = workbook.worksheet(0).expect("worksheet 0");

    // Verify freeze panes
    let freeze = sheet
        .freeze_panes()
        .expect("freeze panes should be present");
    // Froze at C3 → 2 rows frozen (first unfrozen row = 2, 0-based),
    // 2 columns frozen (first unfrozen col = 2, 0-based)
    assert_eq!(freeze.row, 2, "freeze row (first unfrozen row, 0-based)");
    assert_eq!(freeze.col, 2, "freeze col (first unfrozen col, 0-based)");

    // Verify selections
    let sels = sheet.selections();
    assert!(
        !sels.is_empty(),
        "should have at least one selection from Excel"
    );

    // The active pane (bottomRight) should have our D5:F8 selection
    let br_sel = sels
        .iter()
        .find(|s| s.pane.as_deref() == Some("bottomRight"));
    assert!(
        br_sel.is_some(),
        "should have bottomRight pane selection, got: {sels:?}"
    );
    let br = br_sel.unwrap();
    assert_eq!(
        br.active_cell.as_deref(),
        Some("D5"),
        "active cell should be D5"
    );
    assert!(
        br.sqref.as_deref().unwrap_or("").contains("D5"),
        "sqref should contain D5, got: {:?}",
        br.sqref
    );
    assert!(
        br.sqref.as_deref().unwrap_or("").contains("F8"),
        "sqref should contain F8, got: {:?}",
        br.sqref
    );

    // Verify cell data survived
    crate::assert_string(&sheet, 0, 0, "Frozen header A", "A1");
    crate::assert_number(&sheet, 1, 1, 20.0, "B2");

    crate::cleanup_fixture(&fixture);
}
