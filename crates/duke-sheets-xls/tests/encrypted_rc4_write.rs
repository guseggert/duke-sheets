//! End-to-end round-trip tests for the XLS FilePass encrypt path.
//!
//! Each test pulls a known plaintext `/Workbook` byte stream from the
//! LibreOffice-produced `xls_rc4_cryptoapi.plain.xls` fixture, encrypts
//! it with `xls::encrypt_workbook_stream` for one of the three FilePass
//! variants, wraps the result in a fresh CFB envelope, and reads the
//! result back through `XlsReader::read_with` to confirm the
//! original cell values are recovered.

use std::io::Cursor;
use std::path::PathBuf;

use duke_sheets_crypto::xls::{encrypt_workbook_stream, XlsEncryptionVariant};
use duke_sheets_xls::cfb::{CompoundFile, CompoundFileBuilder};
use duke_sheets_xls::{XlsError, XlsReadOptions, XlsReader};

const FIXTURE_PASSWORD: &str = "duke-test-pw";
const FIXTURE_NAME: &str = "xls_rc4_cryptoapi.plain.xls";

fn fixture_path() -> Option<PathBuf> {
    let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    p.push("..");
    p.push("duke-sheets-crypto");
    p.push("tests/fixtures");
    p.push(FIXTURE_NAME);
    p.canonicalize().ok()
}

fn skip_if_missing() -> Option<PathBuf> {
    let Some(p) = fixture_path() else {
        eprintln!("SKIP: {FIXTURE_NAME} not present; run `mise run crypto:fixtures`");
        return None;
    };
    Some(p)
}

fn extract_workbook_stream(fixture: &PathBuf) -> Vec<u8> {
    let bytes = std::fs::read(fixture).expect("read fixture bytes");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("open plaintext CFB");
    cfb.read_stream("/Workbook").expect("read /Workbook stream")
}

fn wrap_in_cfb(workbook_bytes: Vec<u8>) -> Vec<u8> {
    let mut builder = CompoundFileBuilder::new();
    builder
        .add_stream("/Workbook", workbook_bytes)
        .expect("add /Workbook stream");
    builder.build().expect("build CFB envelope")
}

fn run_round_trip(variant: XlsEncryptionVariant) {
    let Some(fixture) = skip_if_missing() else {
        return;
    };
    let plaintext_workbook = extract_workbook_stream(&fixture);

    let encrypted_workbook =
        encrypt_workbook_stream(&plaintext_workbook, FIXTURE_PASSWORD, variant)
            .expect("encrypt_workbook_stream succeeds on LO plaintext fixture");
    let cfb_bytes = wrap_in_cfb(encrypted_workbook);

    let workbook =
        XlsReader::read_with(Cursor::new(&cfb_bytes), &XlsReadOptions { password: Some(FIXTURE_PASSWORD.to_string()), ..Default::default() })
            .expect("XlsReader::read_with recovers our encrypted output");

    let sheet = workbook.worksheet(0).expect("sheet 0 exists");
    let a1 = sheet.get_value("A1").expect("A1 must exist");
    let b1 = sheet.get_value("B1").expect("B1 must exist");
    assert_eq!(a1.as_string().as_deref(), Some("hello crypto"), "A1 value");
    assert_eq!(b1.as_number(), Some(42.0), "B1 value");
}

#[test]
fn round_trip_via_xls_reader_xor_obfuscation() {
    run_round_trip(XlsEncryptionVariant::Xor);
}

#[test]
fn round_trip_via_xls_reader_rc4_legacy() {
    run_round_trip(XlsEncryptionVariant::Rc4Legacy);
}

#[test]
fn round_trip_via_xls_reader_rc4_cryptoapi_128() {
    run_round_trip(XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 });
}

#[test]
fn round_trip_via_xls_reader_rc4_cryptoapi_40() {
    run_round_trip(XlsEncryptionVariant::Rc4CryptoApi { key_bits: 40 });
}

#[test]
fn round_trip_wrong_password_yields_bad_password() {
    let Some(fixture) = skip_if_missing() else {
        return;
    };
    let plaintext_workbook = extract_workbook_stream(&fixture);
    let encrypted_workbook = encrypt_workbook_stream(
        &plaintext_workbook,
        FIXTURE_PASSWORD,
        XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 },
    )
    .expect("encrypt succeeds");
    let cfb_bytes = wrap_in_cfb(encrypted_workbook);

    let err = XlsReader::read_with(Cursor::new(&cfb_bytes), &XlsReadOptions { password: Some("wrong".to_string()), ..Default::default() })
        .expect_err("wrong password must fail");
    assert!(matches!(err, XlsError::BadPassword), "got {err:?}");
}

#[test]
fn round_trip_no_password_reports_encrypted() {
    let Some(fixture) = skip_if_missing() else {
        return;
    };
    let plaintext_workbook = extract_workbook_stream(&fixture);
    let encrypted_workbook = encrypt_workbook_stream(
        &plaintext_workbook,
        FIXTURE_PASSWORD,
        XlsEncryptionVariant::Rc4Legacy,
    )
    .expect("encrypt succeeds");
    let cfb_bytes = wrap_in_cfb(encrypted_workbook);

    let err = XlsReader::read(Cursor::new(&cfb_bytes)).expect_err("must require password");
    assert!(matches!(err, XlsError::Encrypted(_)), "got {err:?}");
}
