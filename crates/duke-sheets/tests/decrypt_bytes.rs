//! `duke_sheets::decrypt_bytes`: password-protected file bytes in, plain
//! file bytes out, for every encryption profile the writers produce.

use duke_sheets::prelude::*;
use duke_sheets::{
    decrypt_bytes, EncryptionProfile, WorkbookExt, WorkbookOpenOptions, WorkbookSaveOptions,
};

// XOR obfuscation caps passwords at 15 characters.
const PASSWORD: &str = "decrypt-pw";
// Enough rows that the package spans several 4096-byte Agile segments
// and the XLS workbook stream spans several 1024-byte RC4 blocks.
const ROWS: u32 = 300;

fn build_wb() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for row in 1..=ROWS {
        ws.set_cell_value(&format!("A{row}"), format!("Item {row}").as_str())
            .unwrap();
        ws.set_cell_value(&format!("B{row}"), f64::from(row) * 1.5)
            .unwrap();
    }
    wb
}

fn profiles() -> Vec<(&'static str, EncryptionProfile)> {
    vec![
        (
            "xlsx",
            EncryptionProfile::OoxmlAgile {
                key_bits: 256,
                spin_count: 1000,
            },
        ),
        ("xlsx", EncryptionProfile::OoxmlStandard { key_bits: 128 }),
        ("xls", EncryptionProfile::XlsRc4CryptoApi { key_bits: 128 }),
        ("xls", EncryptionProfile::XlsRc4Legacy),
        ("xls", EncryptionProfile::XlsXor),
    ]
}

fn save_encrypted(extension: &str, profile: &EncryptionProfile, password: &str) -> Vec<u8> {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join(format!("encrypted.{extension}"));
    let opts = WorkbookSaveOptions::new()
        .password(password)
        .encryption(profile.clone());
    build_wb().save_with(&path, &opts).expect("encrypted save");
    std::fs::read(&path).unwrap()
}

/// The decrypted bytes must open with no password and no sentinel retry.
fn assert_plain_workbook(bytes: &[u8], context: &str) {
    let opened = Workbook::from_bytes_with(bytes, &WorkbookOpenOptions::new().strict_password())
        .unwrap_or_else(|e| panic!("{context}: decrypted bytes must open as plain: {e}"));
    let ws = opened.worksheet(0).unwrap();
    assert_eq!(
        ws.get_value("A1").unwrap().as_string().as_deref(),
        Some("Item 1"),
        "{context}"
    );
    assert_eq!(
        ws.get_value(&format!("A{ROWS}"))
            .unwrap()
            .as_string()
            .as_deref(),
        Some(format!("Item {ROWS}").as_str()),
        "{context}"
    );
    assert_eq!(
        ws.get_value(&format!("B{ROWS}")).unwrap().as_number(),
        Some(f64::from(ROWS) * 1.5),
        "{context}"
    );
}

#[test]
fn decrypts_every_profile_with_its_password() {
    for (extension, profile) in profiles() {
        let context = format!("{profile:?}");
        let encrypted = save_encrypted(extension, &profile, PASSWORD);

        let decrypted = decrypt_bytes(&encrypted, &WorkbookOpenOptions::new().password(PASSWORD))
            .unwrap_or_else(|e| panic!("{context}: {e}"))
            .unwrap_or_else(|| panic!("{context}: file must be reported as encrypted"));

        assert_plain_workbook(&decrypted, &context);
    }
}

#[test]
fn rejects_a_wrong_password_for_every_profile() {
    for (extension, profile) in profiles() {
        let encrypted = save_encrypted(extension, &profile, PASSWORD);

        let err = decrypt_bytes(&encrypted, &WorkbookOpenOptions::new().password("wrong-pw"))
            .expect_err("wrong password must fail");

        assert!(
            err.to_string().contains("Incorrect password"),
            "{profile:?}: unexpected error: {err}"
        );
    }
}

#[test]
fn decrypts_sentinel_password_files_without_a_password() {
    for (extension, profile) in profiles() {
        let context = format!("{profile:?}");
        let encrypted = save_encrypted(extension, &profile, "VelvetSweatshop");

        let decrypted = decrypt_bytes(&encrypted, &WorkbookOpenOptions::new())
            .unwrap_or_else(|e| panic!("{context}: {e}"))
            .unwrap_or_else(|| panic!("{context}: file must be reported as encrypted"));

        assert_plain_workbook(&decrypted, &context);
    }
}

#[test]
fn reports_encrypted_when_no_password_is_supplied() {
    for (extension, profile) in profiles() {
        let encrypted = save_encrypted(extension, &profile, PASSWORD);

        let err = decrypt_bytes(&encrypted, &WorkbookOpenOptions::new())
            .expect_err("a user password is required");

        assert!(
            err.to_string().contains("no password was supplied"),
            "{profile:?}: unexpected error: {err}"
        );
    }
}

#[test]
fn returns_none_for_unencrypted_workbooks() {
    let dir = tempfile::tempdir().unwrap();
    for extension in ["xlsx", "xls"] {
        let path = dir.path().join(format!("plain.{extension}"));
        build_wb().save(&path).unwrap();
        let bytes = std::fs::read(&path).unwrap();

        let result = decrypt_bytes(&bytes, &WorkbookOpenOptions::new().password(PASSWORD)).unwrap();

        assert!(result.is_none(), "{extension}");
    }
}

#[test]
fn xls_output_differs_from_input_only_inside_the_workbook_stream() {
    let encrypted = save_encrypted(
        "xls",
        &EncryptionProfile::XlsRc4CryptoApi { key_bits: 128 },
        PASSWORD,
    );

    let decrypted = decrypt_bytes(&encrypted, &WorkbookOpenOptions::new().password(PASSWORD))
        .unwrap()
        .unwrap();

    assert_eq!(decrypted.len(), encrypted.len());
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(std::io::Cursor::new(&encrypted)).unwrap();
    let stream_len = cfb.read_stream("/Workbook").unwrap().len();
    let changed = decrypted
        .iter()
        .zip(&encrypted)
        .filter(|(a, b)| a != b)
        .count();
    assert!(changed > 0 && changed <= stream_len);
}
