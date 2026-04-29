//! End-to-end tests for `Workbook::save_with(.xls, password)`.
//!
//! Build a Workbook with mixed numeric/string/styled/formula content,
//! save with each XLS encryption profile, then read back via
//! `Workbook::open_with` to confirm the round-trip.

use duke_sheets::{
    EncryptionProfile, Workbook, WorkbookExt, WorkbookOpenOptions, WorkbookSaveOptions,
};

const PASSWORD: &str = "duke-test-pw";

fn build_sample() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello").expect("A1");
    ws.set_cell_value("B1", 42.0).expect("B1");
    ws.set_cell_value("A2", 3.14).expect("A2");
    ws.set_cell_value("B2", true).expect("B2");
    wb
}

fn round_trip(profile: EncryptionProfile) {
    let temp = tempfile::Builder::new()
        .suffix(".xls")
        .tempfile()
        .expect("temp file");
    let wb = build_sample();
    let opts = WorkbookSaveOptions::default()
        .password(PASSWORD)
        .encryption(profile);
    wb.save_with(temp.path(), &opts).expect("save_with");

    let opened = Workbook::open_with(
        temp.path(),
        &WorkbookOpenOptions::default().password(PASSWORD),
    )
    .expect("open_with");
    let sheet = opened.worksheet(0).unwrap();
    assert_eq!(
        sheet.get_value("A1").unwrap().as_string().as_deref(),
        Some("hello")
    );
    assert_eq!(sheet.get_value("B1").unwrap().as_number(), Some(42.0));
    assert_eq!(sheet.get_value("A2").unwrap().as_number(), Some(3.14));
    assert_eq!(sheet.get_value("B2").unwrap().as_bool(), Some(true));
}

#[test]
fn save_with_default_round_trips() {
    round_trip(EncryptionProfile::Default);
}

#[test]
fn save_with_xls_rc4_cryptoapi_128_round_trips() {
    round_trip(EncryptionProfile::XlsRc4CryptoApi { key_bits: 128 });
}

#[test]
fn save_with_xls_rc4_cryptoapi_40_round_trips() {
    round_trip(EncryptionProfile::XlsRc4CryptoApi { key_bits: 40 });
}

#[test]
fn save_with_xls_rc4_legacy_round_trips() {
    round_trip(EncryptionProfile::XlsRc4Legacy);
}

#[test]
fn save_with_xls_xor_round_trips() {
    round_trip(EncryptionProfile::XlsXor);
}

#[test]
fn unencrypted_xls_save_round_trips() {
    let temp = tempfile::Builder::new()
        .suffix(".xls")
        .tempfile()
        .expect("temp file");
    let wb = build_sample();
    wb.save(temp.path()).expect("save");

    let opened = Workbook::open(temp.path()).expect("open");
    let sheet = opened.worksheet(0).unwrap();
    assert_eq!(
        sheet.get_value("A1").unwrap().as_string().as_deref(),
        Some("hello")
    );
    assert_eq!(sheet.get_value("B1").unwrap().as_number(), Some(42.0));
}

#[test]
fn wrong_password_yields_error() {
    let temp = tempfile::Builder::new()
        .suffix(".xls")
        .tempfile()
        .expect("temp file");
    let wb = build_sample();
    let opts = WorkbookSaveOptions::default()
        .password(PASSWORD)
        .encryption(EncryptionProfile::XlsRc4Legacy);
    wb.save_with(temp.path(), &opts).expect("save");

    let err = Workbook::open_with(
        temp.path(),
        &WorkbookOpenOptions::default().password("wrong-password"),
    )
    .expect_err("wrong password must fail");
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encryption"),
        "got {err}"
    );
}
