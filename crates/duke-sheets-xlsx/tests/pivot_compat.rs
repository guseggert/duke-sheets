//! Cross-tool compatibility tests for pivot tables written by duke-sheets.

use std::path::PathBuf;

use duke_sheets_core::{CellRange, PivotAggregate, PivotGrouping, PivotManualGroup, Workbook};
use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use duke_sheets_test_harness::lo::{ensure_lo, SHARED_DIR};
use duke_sheets_xlsx::XlsxWriter;

fn build_manual_grouped_pivot_wb() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "Central").unwrap();
    sheet.set_cell_value("B4", 5.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("ManualGroupedRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn write_pivot_to_shared() -> PathBuf {
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");
    let pid = std::process::id();
    let path = PathBuf::from(format!("{SHARED_DIR}/duke_pivot_manual_group_{pid}.xlsx"));
    let _ = std::fs::remove_file(&path);
    XlsxWriter::write_file(&build_manual_grouped_pivot_wb(), &path)
        .expect("manual grouped pivot write must succeed");
    path
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn lo_can_open_manual_grouped_pivot() {
    ensure_lo();
    let path = write_pivot_to_shared();

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();

    let result = rt.block_on(async {
        let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook(path.to_str().unwrap())
            .await
            .map_err(|e| format!("open: {e}"))?;
        let region = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        wb.close().await.map_err(|e| format!("close: {e}"))?;
        Ok::<String, String>(region)
    });

    let _ = std::fs::remove_file(&path);
    assert_eq!(result.expect("LO must open manual grouped pivot"), "Region");
}
