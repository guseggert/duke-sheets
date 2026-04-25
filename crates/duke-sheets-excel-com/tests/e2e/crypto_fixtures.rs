//! Generate password-protected Office fixtures via real Excel.
//!
//! LibreOffice's XLS-with-password output uses the legacy Office 97/2000
//! RC4 (MD5 KDF, FilePass `vMajor=1, vMinor=1`); to validate our
//! `rc4_cryptoapi` decrypt path we need a fixture from a writer that
//! emits the modern CryptoAPI variant (`vMajor=2-4, vMinor=2`, SHA-1 KDF).
//! Excel does so when `Workbook.SetPasswordEncryptionOptions` is set to
//! a CryptoAPI provider before SaveAs.
//!
//! Output lands at `crates/duke-sheets-crypto/tests/fixtures/`, gitignored.
//!
//! Run with: `cargo test -p duke-sheets-excel-com --test e2e
//!   crypto_fixtures -- --ignored --nocapture --test-threads=1`
//! (or via `mise run crypto:fixtures:excel`).

use std::path::PathBuf;

use crate::{ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, TempFixture};

const FIXTURE_PASSWORD: &str = "duke-test-pw";
const FIXTURE_A1: &str = "hello crypto";
const FIXTURE_B1: f64 = 42.0;

/// Excel `XlFileFormat` constants.
const XL_EXCEL8: i32 = 56;
const XL_OPEN_XML_WORKBOOK: i32 = 51;

/// Crypto fixtures land alongside the LO-generated ones.
fn fixtures_dir() -> PathBuf {
    let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    p.push("../duke-sheets-crypto/tests/fixtures");
    p
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn generate_xls_rc4_cryptoapi_fixture_via_excel() {
    let bridge = excel_bridge();
    ensure_vm_temp_dir();

    let fixture = TempFixture {
        host_path: PathBuf::from("/tmp/duke-sheets-excel/xls_rc4_cryptoapi_excel.xls"),
        vm_path: r"C:\temp\xls_rc4_cryptoapi_excel.xls".to_string(),
        name: "xls_rc4_cryptoapi_excel.xls".to_string(),
    };

    {
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", FIXTURE_A1).expect("set A1");
        wb.set_cell_value("B1", FIXTURE_B1).expect("set B1");

        // Switch from Excel's default legacy MD5 RC4 to RC4 CryptoAPI
        // (SHA-1 KDF, vMajor=2-4, vMinor=2 in the FilePass record).
        wb.set_password_encryption_options(
            "Microsoft Enhanced Cryptographic Provider v1.0",
            "RC4",
            128,
            false,
        )
        .expect("set encryption options");

        wb.save_with_password(&fixture.vm_path, XL_EXCEL8, FIXTURE_PASSWORD)
            .expect("save with password");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);

    // Stash a copy in the canonical fixtures directory for downstream
    // crypto tests to consume.
    let dest = fixtures_dir();
    std::fs::create_dir_all(&dest).expect("create fixtures dir");
    let dest_path = dest.join("xls_rc4_cryptoapi_excel.xls");
    std::fs::copy(&fixture.host_path, &dest_path).expect("copy into fixtures");

    let bytes = std::fs::read(&dest_path).expect("read");
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "Excel-encrypted .xls must be a CFB envelope"
    );
    eprintln!(
        "wrote {} ({} bytes); password={FIXTURE_PASSWORD:?}",
        dest_path.display(),
        bytes.len()
    );
}

#[test]
#[ignore = "requires running Excel COM bridge on localhost:9876 (mise run vm:start)"]
fn generate_xlsx_agile_fixture_via_excel() {
    let bridge = excel_bridge();
    ensure_vm_temp_dir();

    let fixture = TempFixture {
        host_path: PathBuf::from("/tmp/duke-sheets-excel/xlsx_agile_excel.xlsx"),
        vm_path: r"C:\temp\xlsx_agile_excel.xlsx".to_string(),
        name: "xlsx_agile_excel.xlsx".to_string(),
    };

    {
        let excel = bridge.lock().unwrap();
        let wb = excel.create_workbook().expect("create workbook");
        wb.set_cell_value("A1", FIXTURE_A1).expect("set A1");
        wb.set_cell_value("B1", FIXTURE_B1).expect("set B1");

        // Excel's default password encryption for .xlsx is Agile (since
        // Office 2010). Skipping SetPasswordEncryptionOptions selects
        // the default, producing an Agile-encrypted CFB envelope.
        wb.save_with_password(&fixture.vm_path, XL_OPEN_XML_WORKBOOK, FIXTURE_PASSWORD)
            .expect("save with password");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);

    let dest = fixtures_dir();
    std::fs::create_dir_all(&dest).expect("create fixtures dir");
    let dest_path = dest.join("xlsx_agile_excel.xlsx");
    std::fs::copy(&fixture.host_path, &dest_path).expect("copy into fixtures");

    let bytes = std::fs::read(&dest_path).expect("read");
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "Excel-encrypted .xlsx must be a CFB envelope"
    );
    eprintln!(
        "wrote {} ({} bytes); password={FIXTURE_PASSWORD:?}",
        dest_path.display(),
        bytes.len()
    );
}
