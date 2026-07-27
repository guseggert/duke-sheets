//! XLSB parity fixtures: Excel converts the XLSX parity workbooks to
//! `.xlsb` via SaveAs, giving ground-truth BIFF12 to read back.
//!
//! Both fixtures chain to their XLSX source, which generates itself if
//! absent, so a clean checkout needs no manual fixture step.

use std::path::PathBuf;
use std::sync::OnceLock;

use crate::chart_parity::chart_parity_fixture;
use crate::formula_parity::formula_parity_fixture;
use crate::{ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, push_file_to_vm, TempFixture};

/// `xlExcel12` - the `.xlsb` SaveAs FileFormat.
const XL_EXCEL12: i32 = 50;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

/// Convert `source_xlsx` to XLSB through Excel and copy the result to
/// `repo_relative`, returning its path.
fn convert_to_xlsb(source_xlsx: &PathBuf, stem: &str, repo_relative: &str) -> PathBuf {
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    ensure_vm_temp_dir();

    let vm_input = format!(r"C:\temp\{stem}_src.xlsx");
    let vm_output = format!(r"C:\temp\{stem}.xlsb");

    push_file_to_vm(&TempFixture {
        host_path: source_xlsx.clone(),
        vm_path: vm_input.clone(),
        name: format!("{stem}_src.xlsx"),
    });

    let wb = excel
        .open_workbook(&vm_input)
        .unwrap_or_else(|e| panic!("open {}: {e}", source_xlsx.display()));
    wb.save_as(&vm_output, XL_EXCEL12).expect("SaveAs XLSB");
    wb.close().expect("close");

    let pulled = TempFixture {
        host_path: PathBuf::from(format!("/tmp/duke-sheets-excel/{stem}.xlsb")),
        vm_path: vm_output,
        name: format!("{stem}.xlsb"),
    };
    let _ = std::fs::create_dir_all("/tmp/duke-sheets-excel");
    pull_file_from_vm(&pulled);

    let dest = repo_path(repo_relative);
    if let Some(parent) = dest.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    std::fs::copy(&pulled.host_path, &dest)
        .unwrap_or_else(|e| panic!("copy to {}: {e}", dest.display()));
    dest
}

/// Path to `data/formula-parity.xlsb`, generating it (and its XLSX source)
/// if absent.
pub fn formula_parity_xlsb() -> PathBuf {
    static PATH: OnceLock<PathBuf> = OnceLock::new();
    PATH.get_or_init(|| {
        let dest = repo_path("data/formula-parity.xlsb");
        if dest.exists() {
            return dest;
        }
        convert_to_xlsb(&formula_parity_fixture(), "formula_parity", "data/formula-parity.xlsb")
    })
    .clone()
}

/// Path to `data/chart-parity.xlsb`, generating it (and its XLSX source)
/// if absent.
pub fn chart_parity_xlsb() -> PathBuf {
    static PATH: OnceLock<PathBuf> = OnceLock::new();
    PATH.get_or_init(|| {
        let dest = repo_path("data/chart-parity.xlsb");
        if dest.exists() {
            return dest;
        }
        convert_to_xlsb(&chart_parity_fixture(), "chart_parity", "data/chart-parity.xlsb")
    })
    .clone()
}

#[test]
fn xlsb_parity_fixtures_generate() {
    assert!(formula_parity_xlsb().exists());
    assert!(chart_parity_xlsb().exists());
}
