//! End-to-end tests for the OOXML Agile *write* path.
//!
//! These are pure round-trip checks (write → read with the same
//! library); cross-tool compatibility tests (LO, Excel COM) live in
//! `#[ignore]`-gated modules so the default `cargo test` doesn't need
//! either backend.

use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xlsx::{EncryptionProfile, XlsxError, XlsxReadOptions, XlsxReader, XlsxWriter};

const PASSWORD: &str = "round-trip-pw";

fn build_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello agile write").unwrap();
    ws.set_cell_value("B1", 1234.5).unwrap();
    ws.set_cell_value("C1", true).unwrap();
    wb
}

fn assert_workbook_contents(wb: &Workbook) {
    let ws = wb.worksheet(0).expect("sheet 0");
    assert_eq!(
        ws.get_value("A1").unwrap().as_string().as_deref(),
        Some("hello agile write")
    );
    assert_eq!(ws.get_value("B1").unwrap().as_number(), Some(1234.5));
    assert_eq!(ws.get_value("C1").unwrap(), CellValue::Boolean(true));
}

#[test]
fn write_then_read_agile_round_trips_workbook_contents() {
    let wb = build_workbook();
    let bytes =
        XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &EncryptionProfile::agile_default())
            .expect("encrypted write must succeed");
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "encrypted output must be a CFB envelope"
    );

    let opened = XlsxReader::read_bytes_with(&bytes, &XlsxReadOptions { password: Some(PASSWORD.to_string()), ..Default::default() })
        .expect("decrypt with correct password must succeed");
    assert_workbook_contents(&opened);
}

#[test]
fn write_then_read_agile_with_wrong_password_yields_bad_password() {
    let wb = build_workbook();
    let bytes =
        XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &EncryptionProfile::agile_default())
            .unwrap();
    let err = XlsxReader::read_bytes_with(&bytes, &XlsxReadOptions { password: Some("not-the-pw".to_string()), ..Default::default() })
        .expect_err("must reject wrong password");
    assert!(
        matches!(err, XlsxError::BadPassword),
        "expected BadPassword, got {err:?}"
    );
}

#[test]
fn write_then_read_agile_aes128_round_trip() {
    let wb = build_workbook();
    let profile = EncryptionProfile::Agile {
        key_bits: 128,
        spin_count: 100_000,
    };
    let bytes = XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &profile).unwrap();
    let opened =
        XlsxReader::read_bytes_with(&bytes, &XlsxReadOptions { password: Some(PASSWORD.to_string()), ..Default::default() }).expect("decrypt ok");
    assert_workbook_contents(&opened);
}

#[test]
fn write_to_file_encrypted_writes_cfb_envelope() {
    let wb = build_workbook();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("encrypted.xlsx");
    XlsxWriter::write_file_encrypted(&wb, &path, PASSWORD, &EncryptionProfile::agile_default())
        .expect("write to file must succeed");

    let bytes = std::fs::read(&path).unwrap();
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "file at path must be a CFB envelope"
    );
    let opened =
        XlsxReader::read_file_with(&path, &XlsxReadOptions { password: Some(PASSWORD.to_string()), ..Default::default() }).expect("decrypt file");
    assert_workbook_contents(&opened);
}

#[test]
fn write_with_low_spincount_speeds_up_kdf() {
    // spinCount=10 still has to run the KDF and produce a valid wrap;
    // this is just a sanity check that the parameter is honored end to
    // end, not a benchmark.
    let wb = build_workbook();
    let profile = EncryptionProfile::Agile {
        key_bits: 256,
        spin_count: 10,
    };
    let bytes = XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &profile).unwrap();
    let opened =
        XlsxReader::read_bytes_with(&bytes, &XlsxReadOptions { password: Some(PASSWORD.to_string()), ..Default::default() }).expect("decrypt ok");
    assert_workbook_contents(&opened);
}
