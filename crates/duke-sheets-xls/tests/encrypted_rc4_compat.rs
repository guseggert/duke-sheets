//! LibreOffice cross-tool compatibility for the XLS FilePass encrypt
//! path. Each test builds a plaintext workbook in memory, encrypts its
//! `/Workbook` stream with one of our FilePass variants, wraps the result
//! in a fresh CFB envelope, writes the file into the URP-shared directory,
//! and asks LibreOffice to open it with the password and read back the
//! original cell values.
//!
//! The LO container is auto-started by the test harness, so these run
//! unconditionally and fail loud rather than silent-skipping.

use std::io::Cursor;
use std::path::PathBuf;

use duke_sheets_core::Workbook;
use duke_sheets_crypto::xls::{encrypt_workbook_stream, XlsEncryptionVariant};
use duke_sheets_libreoffice::bridge::LibreOfficeBridge;
use duke_sheets_test_harness::lo::{ensure_lo, SHARED_DIR};
use duke_sheets_xls::cfb::{CompoundFile, CompoundFileBuilder};
use duke_sheets_xls::XlsWriter;

const FIXTURE_PASSWORD: &str = "duke-test-pw";
const FIXTURE_CELL_A1: &str = "hello crypto";
const FIXTURE_CELL_B1: f64 = 42.0;

fn require_lo() {
    ensure_lo();
}

/// The plaintext workbook this suite encrypts, built in memory.
///
/// The subject under test is the FilePass wrapper, not the BIFF8 body, so
/// the plaintext only needs to be a valid workbook carrying the two cells
/// the assertions check. Building it here keeps the suite hermetic: there
/// is no fixture file to generate, stage, or go stale.
fn plaintext_workbook_bytes() -> Vec<u8> {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).expect("default worksheet");
    sheet
        .set_cell_value("A1", FIXTURE_CELL_A1)
        .expect("set A1");
    sheet
        .set_cell_value("B1", FIXTURE_CELL_B1)
        .expect("set B1");
    XlsWriter::write_to_bytes(&wb).expect("serialize plaintext workbook")
}

fn encrypt_fixture_workbook(variant: XlsEncryptionVariant) -> Vec<u8> {
    let bytes = plaintext_workbook_bytes();
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

fn write_encrypted_to_shared(variant: XlsEncryptionVariant, suffix: &str) -> PathBuf {
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");
    let pid = std::process::id();
    let path = PathBuf::from(format!("{SHARED_DIR}/duke_xls_compat_{suffix}_{pid}.xls"));
    let _ = std::fs::remove_file(&path);
    let bytes = encrypt_fixture_workbook(variant);
    std::fs::write(&path, &bytes).expect("write CFB to shared dir");
    path
}

fn lo_can_read(variant: XlsEncryptionVariant, suffix: &str) {
    require_lo();
    let path = write_encrypted_to_shared(variant, suffix);
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let result = rt.block_on(async {
        let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
            .await
            .map_err(|e| format!("connect: {e}"))?;
        let mut wb = bridge
            .open_workbook_with_password(path.to_str().unwrap(), FIXTURE_PASSWORD)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb
            .get_cell_string("A1")
            .await
            .map_err(|e| format!("read A1: {e}"))?;
        let b1 = wb
            .get_cell_value("B1")
            .await
            .map_err(|e| format!("read B1: {e}"))?;
        Ok::<(String, f64), String>((a1, b1))
    });
    let _ = std::fs::remove_file(&path);
    let (a1, b1) = result.unwrap_or_else(|e| {
        panic!("LO must open file we encrypted with {variant:?}: {e}")
    });
    assert_eq!(a1, FIXTURE_CELL_A1);
    assert!((b1 - FIXTURE_CELL_B1).abs() < 1e-9, "B1 = {b1}");
}

#[test]
fn lo_can_read_rc4_cryptoapi_128() {
    lo_can_read(
        XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 },
        "rc4cryptoapi128",
    );
}

#[test]
fn lo_can_read_rc4_legacy() {
    lo_can_read(XlsEncryptionVariant::Rc4Legacy, "rc4legacy");
}

// XOR Obfuscation cross-tool compat is intentionally not exercised
// here: modern Excel's COM API doesn't expose an XOR write path for
// `.xls`, and LibreOffice's `MS Excel 97` filter emits legacy RC4
// (MD5) by default with no XOR option. Without a known-good reference
// XOR fixture we can't disambiguate "our cipher walk is wrong" from
// "real tools have idiosyncratic XOR readers". The XOR write path is
// covered by the encrypted_rc4_write::round_trip_via_xls_reader_xor
// round-trip (encrypt → our reader's decrypt) and exercised by
// xls_encrypt::round_trip_xor_obfuscation; that verifies symmetry but
// does not certify cross-tool interop.
