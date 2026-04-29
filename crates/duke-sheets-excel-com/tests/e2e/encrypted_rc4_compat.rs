//! Cross-tool compatibility test: real Excel must open XLS FilePass
//! files we encrypt. The plaintext source is the gitignored LO fixture
//! `xls_rc4_cryptoapi.plain.xls`; for each variant we encrypt its
//! `/Workbook` stream, wrap it in a fresh CFB envelope, push to the VM,
//! and ask Excel to open it with the password.

use std::io::Cursor;
use std::path::PathBuf;

use crate::{ensure_vm_temp_dir, excel_bridge, push_file_to_vm, TempFixture};
use duke_sheets_crypto::xls::{encrypt_workbook_stream, XlsEncryptionVariant};
use duke_sheets_xls::cfb::{CompoundFile, CompoundFileBuilder};

const FIXTURE_PASSWORD: &str = "duke-test-pw";
const FIXTURE_NAME: &str = "xls_rc4_cryptoapi.plain.xls";
const HOST_DIR: &str = "/tmp/duke-sheets-excel";

fn fixture_path_in_repo() -> PathBuf {
    let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    p.push("..");
    p.push("duke-sheets-crypto");
    p.push("tests/fixtures");
    p.push(FIXTURE_NAME);
    p.canonicalize().unwrap_or_else(|_| {
        panic!(
            "{FIXTURE_NAME} not present in tests/fixtures. \
             Regenerate with `mise run crypto:fixtures` (requires LO container)."
        )
    })
}

fn encrypt_fixture_workbook(variant: XlsEncryptionVariant) -> Vec<u8> {
    let fixture = fixture_path_in_repo();
    let bytes = std::fs::read(&fixture).expect("read fixture bytes");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open plaintext CFB");
    let plaintext = cfb.read_stream("/Workbook").expect("read /Workbook");
    let encrypted = encrypt_workbook_stream(&plaintext, FIXTURE_PASSWORD, variant)
        .expect("encrypt_workbook_stream succeeds");
    let mut builder = CompoundFileBuilder::new();
    builder
        .add_stream("/Workbook", encrypted)
        .expect("add encrypted /Workbook");
    builder.build().expect("build CFB envelope")
}

fn fixture_for(name: &str) -> TempFixture {
    let _ = std::fs::create_dir_all(HOST_DIR);
    TempFixture {
        host_path: PathBuf::from(format!("{HOST_DIR}/{name}")),
        vm_path: format!(r"C:\temp\{name}"),
        name: name.to_string(),
    }
}

fn excel_can_read(variant: XlsEncryptionVariant, name: &str) {
    let fixture = fixture_for(name);
    let _ = std::fs::remove_file(&fixture.host_path);
    std::fs::write(&fixture.host_path, encrypt_fixture_workbook(variant))
        .expect("write CFB to host dir");
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let wb = excel
        .open_workbook_with_password(&fixture.vm_path, FIXTURE_PASSWORD)
        .unwrap_or_else(|e| panic!("Excel must open file we encrypted with {variant:?}: {e}"));

    let wb_name = wb.name().expect("workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel reported the file as 'Repaired': {wb_name}"
    );

    let a1 = wb.get_cell_value("A1").expect("read A1");
    let b1 = wb.get_cell_value("B1").expect("read B1");
    wb.close().expect("close workbook");
    let _ = std::fs::remove_file(&fixture.host_path);

    assert_eq!(a1.as_str(), Some("hello crypto"), "A1 value");
    let b1_num = b1.as_f64().expect("B1 numeric");
    assert!((b1_num - 42.0).abs() < 1e-9, "B1 = {b1_num}");
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876 + crypto fixture"]
fn excel_can_read_rc4_cryptoapi_128() {
    excel_can_read(
        XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 },
        "duke_xls_rc4cryptoapi128.xls",
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876 + crypto fixture"]
fn excel_can_read_rc4_legacy() {
    excel_can_read(XlsEncryptionVariant::Rc4Legacy, "duke_xls_rc4legacy.xls");
}

// XOR Obfuscation cross-tool compat is intentionally not exercised
// here: Excel's COM API doesn't expose XOR write for `.xls` (modern
// Excel ships only RC4 variants for password-protect-on-save). Without
// a real Excel-emitted XOR fixture we can't disambiguate "our XOR
// cipher walk has a bug" from "modern Excel's XOR reader has its own
// quirks". Round-trip via our own reader is covered by the
// encrypted_rc4_write tests in duke-sheets-xls.
