//! Top-level `Workbook::save_with` round-trip for OOXML Agile.
//!
//! This is the user-facing API for writing password-protected files;
//! these tests exercise the dispatch from [`WorkbookSaveOptions`] down
//! through to the XLSX writer.

use duke_sheets::prelude::*;
use duke_sheets::{
    EncryptionProfile, WorkbookExt, WorkbookOpenOptions, WorkbookSaveOptions, XlsxError,
};

const PASSWORD: &str = "save-with-pw";

fn build_wb() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "round-trip via save_with").unwrap();
    ws.set_cell_value("B1", 7.5).unwrap();
    wb
}

#[test]
fn save_with_password_default_profile_round_trips() {
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("rt.xlsx");

    let opts = WorkbookSaveOptions::new().password(PASSWORD);
    wb.save_with(&path, &opts).expect("encrypted save");

    let bytes = std::fs::read(&path).unwrap();
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "default profile must produce CFB envelope"
    );

    let opened = Workbook::open_with(&path, &WorkbookOpenOptions::new().password(PASSWORD))
        .expect("open with password");
    let ws = opened.worksheet(0).unwrap();
    assert_eq!(
        ws.get_value("A1").unwrap().as_string().as_deref(),
        Some("round-trip via save_with")
    );
    assert_eq!(ws.get_value("B1").unwrap().as_number(), Some(7.5));
}

#[test]
fn save_with_explicit_agile_profile_round_trips() {
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("agile.xlsx");

    let opts =
        WorkbookSaveOptions::new()
            .password(PASSWORD)
            .encryption(EncryptionProfile::OoxmlAgile {
                key_bits: 128,
                spin_count: 50,
            });
    wb.save_with(&path, &opts).expect("save");

    let opened =
        Workbook::open_with(&path, &WorkbookOpenOptions::new().password(PASSWORD)).expect("open");
    assert_eq!(
        opened
            .worksheet(0)
            .unwrap()
            .get_value("A1")
            .unwrap()
            .as_string()
            .as_deref(),
        Some("round-trip via save_with")
    );
}

#[test]
fn save_with_no_password_writes_plain_file() {
    // Without a password, save_with must produce a plain ZIP-format
    // .xlsx (no encryption, regardless of `encryption` field).
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("plain.xlsx");
    wb.save_with(&path, &WorkbookSaveOptions::default())
        .expect("plain save");
    let bytes = std::fs::read(&path).unwrap();
    assert_eq!(
        &bytes[0..4],
        &[0x50, 0x4B, 0x03, 0x04],
        "no-password save must produce a plain ZIP"
    );
}

#[test]
fn save_with_wrong_password_on_open_yields_bad_password() {
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("rt.xlsx");
    wb.save_with(&path, &WorkbookSaveOptions::new().password(PASSWORD))
        .unwrap();

    let result = Workbook::open_with(&path, &WorkbookOpenOptions::new().password("nope"));
    let err = result.expect_err("wrong password must fail");
    // Top-level Error wraps the format-specific BadPassword as a string;
    // we just assert the error mentions "password" since it's
    // reformatted through Error::other.
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encrypted"),
        "expected password-related error message, got: {msg}"
    );

    // Sanity: the underlying XLSX layer surfaces BadPassword, even if
    // the duke-sheets error layer flattens it.
    let bytes = std::fs::read(&path).unwrap();
    let err =
        duke_sheets::XlsxReader::read_bytes_with(&bytes, &duke_sheets::XlsxReadOptions { password: Some("nope".to_string()), ..Default::default() }).unwrap_err();
    assert!(matches!(err, XlsxError::BadPassword));
}

/// Read the file's CFB envelope, find the byte offset where
/// `/EncryptedPackage` content lives, and flip a byte there so the
/// CFB structure + EncryptionInfo XML stay parseable but the HMAC
/// fails.
fn tamper_encrypted_package_byte(path: &std::path::Path) {
    let bytes = std::fs::read(path).unwrap();
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(std::io::Cursor::new(&bytes)).unwrap();
    let pkg = cfb.read_stream("/EncryptedPackage").unwrap();
    let needle: [u8; 16] = pkg[16..32].try_into().unwrap();
    let pos = bytes
        .windows(needle.len())
        .position(|w| w == needle)
        .expect("encrypted package bytes must appear in file");
    let mut new_bytes = bytes;
    new_bytes[pos + 4] ^= 0xFF;
    std::fs::write(path, new_bytes).unwrap();
}

#[test]
fn open_with_tampered_envelope_default_rejects_via_integrity_check() {
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tampered.xlsx");
    wb.save_with(&path, &WorkbookSaveOptions::new().password(PASSWORD))
        .unwrap();
    tamper_encrypted_package_byte(&path);

    let bytes = std::fs::read(&path).unwrap();
    let xlsx_err = duke_sheets::XlsxReader::read_bytes_with(&bytes, &duke_sheets::XlsxReadOptions { password: Some(PASSWORD.to_string()), ..Default::default() })
        .expect_err("tampered envelope must fail integrity check");
    assert!(
        matches!(xlsx_err, XlsxError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {xlsx_err:?}"
    );
}

#[test]
fn open_with_tampered_envelope_and_skip_integrity_check_bypasses_check() {
    let wb = build_wb();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tampered.xlsx");
    wb.save_with(&path, &WorkbookSaveOptions::new().password(PASSWORD))
        .unwrap();
    tamper_encrypted_package_byte(&path);

    let opts = WorkbookOpenOptions::new()
        .password(PASSWORD)
        .skip_integrity_check();
    // Skip-flag suppresses IntegrityCheckFailed. The 16-byte AES block
    // containing the flipped byte is garbled, so ZIP parsing of the
    // decrypted package may fail — we only assert that whatever error
    // surfaces, it is NOT the integrity check.
    let result = Workbook::open_with(&path, &opts);
    if let Err(e) = &result {
        let msg = e.to_string().to_lowercase();
        assert!(
            !msg.contains("integrity check"),
            "skip_integrity_check must suppress IntegrityCheckFailed but got: {msg}"
        );
    }
}
