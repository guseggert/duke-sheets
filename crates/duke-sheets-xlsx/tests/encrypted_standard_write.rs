use duke_sheets_core::{CellValue, Workbook};
use duke_sheets_xlsx::{EncryptionProfile, XlsxError, XlsxReader, XlsxWriter};

const PASSWORD: &str = "standard-rt-pw";

fn build_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "hello standard write").unwrap();
    ws.set_cell_value("B1", 9876.5).unwrap();
    ws.set_cell_value("C1", true).unwrap();
    wb
}

fn assert_workbook_contents(wb: &Workbook) {
    let ws = wb.worksheet(0).expect("sheet 0");
    assert_eq!(
        ws.get_value("A1").unwrap().as_string().as_deref(),
        Some("hello standard write")
    );
    assert_eq!(ws.get_value("B1").unwrap().as_number(), Some(9876.5));
    assert_eq!(ws.get_value("C1").unwrap(), CellValue::Boolean(true));
}

#[test]
fn write_then_read_standard_round_trips_workbook_contents() {
    let wb = build_workbook();
    let bytes =
        XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &EncryptionProfile::standard_default())
            .expect("encrypted write");
    assert_eq!(
        &bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "Standard envelope must be CFB"
    );
    let opened =
        XlsxReader::read_bytes_with_password(&bytes, Some(PASSWORD), false).expect("decrypt");
    assert_workbook_contents(&opened);
}

#[test]
fn write_then_read_standard_with_wrong_password_yields_bad_password() {
    let wb = build_workbook();
    let bytes =
        XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &EncryptionProfile::standard_default())
            .unwrap();
    let err = XlsxReader::read_bytes_with_password(&bytes, Some("wrong"), false)
        .expect_err("wrong password");
    assert!(matches!(err, XlsxError::BadPassword));
}

#[test]
fn write_then_read_standard_aes256_round_trip() {
    let wb = build_workbook();
    let profile = EncryptionProfile::Standard { key_bits: 256 };
    let bytes = XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &profile).unwrap();
    let opened =
        XlsxReader::read_bytes_with_password(&bytes, Some(PASSWORD), false).expect("decrypt");
    assert_workbook_contents(&opened);
}

#[test]
fn write_then_read_standard_aes192_round_trip() {
    let wb = build_workbook();
    let profile = EncryptionProfile::Standard { key_bits: 192 };
    let bytes = XlsxWriter::write_to_bytes_encrypted(&wb, PASSWORD, &profile).unwrap();
    let opened =
        XlsxReader::read_bytes_with_password(&bytes, Some(PASSWORD), false).expect("decrypt");
    assert_workbook_contents(&opened);
}
