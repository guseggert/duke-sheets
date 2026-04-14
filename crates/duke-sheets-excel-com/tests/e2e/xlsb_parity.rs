use std::path::PathBuf;

use crate::{ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, push_file_to_vm, TempFixture};

const XLSX_FIXTURE_PATH: &str = "data/formula-parity.xlsx";
const XLSB_REPO_PATH: &str = "data/formula-parity.xlsb";
const VM_INPUT: &str = r"C:\temp\formula_parity_src.xlsx";
const VM_OUTPUT: &str = r"C:\temp\formula_parity.xlsb";

const CHART_XLSX_PATH: &str = "data/chart-parity.xlsx";
const CHART_XLSB_PATH: &str = "data/chart-parity.xlsb";
const CHART_VM_INPUT: &str = r"C:\temp\chart_parity_src.xlsx";
const CHART_VM_OUTPUT: &str = r"C:\temp\chart_parity.xlsb";

#[test]
fn generate_xlsb_parity_from_xlsx() {
    let xlsx_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(XLSX_FIXTURE_PATH);
    if !xlsx_path.exists() {
        panic!("{XLSX_FIXTURE_PATH} not found — run `mise run generate:formula-parity` first");
    }

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    ensure_vm_temp_dir();

    let input_fixture = TempFixture {
        host_path: xlsx_path.clone(),
        vm_path: VM_INPUT.to_string(),
        name: "formula_parity_src.xlsx".to_string(),
    };
    push_file_to_vm(&input_fixture);

    let wb = excel
        .open_workbook(VM_INPUT)
        .expect("open formula-parity.xlsx");

    // FileFormat=50 = xlExcel12 (.xlsb)
    wb.save_as(VM_OUTPUT, 50).expect("SaveAs XLSB");
    wb.close().expect("close");

    let output_fixture = TempFixture {
        host_path: PathBuf::from("/tmp/duke-sheets-excel/formula_parity.xlsb"),
        vm_path: VM_OUTPUT.to_string(),
        name: "formula_parity.xlsb".to_string(),
    };
    let _ = std::fs::create_dir_all("/tmp/duke-sheets-excel");
    pull_file_from_vm(&output_fixture);

    let dest = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(XLSB_REPO_PATH);
    std::fs::copy(&output_fixture.host_path, &dest).expect("copy to data/formula-parity.xlsb");
    println!("Generated {}", dest.display());
}

#[test]
fn generate_xlsb_chart_parity_from_xlsx() {
    let xlsx_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(CHART_XLSX_PATH);
    if !xlsx_path.exists() {
        panic!("{CHART_XLSX_PATH} not found — run `mise run generate:chart-parity` first");
    }

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    ensure_vm_temp_dir();

    let input_fixture = TempFixture {
        host_path: xlsx_path.clone(),
        vm_path: CHART_VM_INPUT.to_string(),
        name: "chart_parity_src.xlsx".to_string(),
    };
    push_file_to_vm(&input_fixture);

    let wb = excel
        .open_workbook(CHART_VM_INPUT)
        .expect("open chart-parity.xlsx");

    wb.save_as(CHART_VM_OUTPUT, 50).expect("SaveAs XLSB");
    wb.close().expect("close");

    let output_fixture = TempFixture {
        host_path: PathBuf::from("/tmp/duke-sheets-excel/chart_parity.xlsb"),
        vm_path: CHART_VM_OUTPUT.to_string(),
        name: "chart_parity.xlsb".to_string(),
    };
    let _ = std::fs::create_dir_all("/tmp/duke-sheets-excel");
    pull_file_from_vm(&output_fixture);

    let dest = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(CHART_XLSB_PATH);
    std::fs::copy(&output_fixture.host_path, &dest).expect("copy to data/chart-parity.xlsb");
    println!("Generated {}", dest.display());
}
