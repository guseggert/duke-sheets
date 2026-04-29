//! End-to-end round-trip tests for `xls::encrypt_workbook_stream`.
//!
//! Each test builds a synthetic plaintext Workbook stream containing a
//! globals BOF, a BoundSheet8 record pointing at a worksheet stream,
//! an EOF, then a worksheet BOF/EOF pair. The stream is encrypted with
//! each XLS encryption variant and decrypted back through
//! `xls::decrypt_workbook_stream`.

use duke_sheets_crypto::error::CryptoError;
use duke_sheets_crypto::xls::{
    decrypt_workbook_stream, encrypt_workbook_stream, XlsEncryptionVariant,
};

const PASSWORD: &str = "duke-test-pw";
const FILEPASS_RECORD_TYPE: u16 = 0x002F;
const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const BOUND_SHEET_8: u16 = 0x0085;

fn put_u16(buf: &mut Vec<u8>, v: u16) {
    buf.extend_from_slice(&v.to_le_bytes());
}

fn put_u32(buf: &mut Vec<u8>, v: u32) {
    buf.extend_from_slice(&v.to_le_bytes());
}

fn put_record(buf: &mut Vec<u8>, record_type: u16, body: &[u8]) {
    put_u16(buf, record_type);
    put_u16(buf, body.len() as u16);
    buf.extend_from_slice(body);
}

/// Build a plaintext Workbook stream with one BoundSheet8 record
/// pointing at a worksheet whose body is just a NUMBER record.
fn build_plaintext_stream() -> Vec<u8> {
    let mut globals = Vec::new();
    put_record(&mut globals, BOF, &[0x06, 0x00, 0x05, 0x00, 0, 0, 0, 0]);
    let bound_sheet_offset_field_position = globals.len() + 4;
    put_record(
        &mut globals,
        BOUND_SHEET_8,
        &[0, 0, 0, 0, 0, 0, 6, b'S', b'h', b'e', b'e', b't', b'1'],
    );
    put_record(&mut globals, EOF, &[]);

    let sheet_offset = globals.len() as u32;
    let lbplypos_bytes = sheet_offset.to_le_bytes();
    globals[bound_sheet_offset_field_position..bound_sheet_offset_field_position + 4]
        .copy_from_slice(&lbplypos_bytes);

    let mut stream = globals;
    put_record(&mut stream, BOF, &[0x06, 0x00, 0x10, 0x00, 0, 0, 0, 0]);
    let mut number_body = Vec::new();
    put_u16(&mut number_body, 0); // row
    put_u16(&mut number_body, 0); // col
    put_u16(&mut number_body, 0); // xf
    number_body.extend_from_slice(&42.0f64.to_le_bytes());
    put_record(&mut stream, 0x0203, &number_body);
    put_record(&mut stream, EOF, &[]);
    stream
}

fn read_lbplypos(stream: &[u8]) -> u32 {
    let mut cursor = 0;
    while cursor + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
        if record_type == BOUND_SHEET_8 && size >= 4 {
            return u32::from_le_bytes([
                stream[cursor + 4],
                stream[cursor + 5],
                stream[cursor + 6],
                stream[cursor + 7],
            ]);
        }
        cursor += 4 + size;
    }
    panic!("BoundSheet8 record not found in stream");
}

fn find_filepass_record_size(stream: &[u8]) -> usize {
    let mut cursor = 0;
    while cursor + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
        if record_type == FILEPASS_RECORD_TYPE {
            return 4 + size;
        }
        cursor += 4 + size;
    }
    panic!("FilePass record not found in stream");
}

/// Reconstruct a plaintext-equivalent stream from a decrypted stream
/// by deleting the (post-decrypt-neutered) FilePass record and
/// un-bumping every `BoundSheet8.lbPlyPos` value by the FilePass record
/// size. The result should byte-equal the input that was originally
/// passed to `encrypt_workbook_stream`.
fn strip_neutered_filepass(stream: &[u8], filepass_pos: usize) -> Vec<u8> {
    let filepass_size =
        u16::from_le_bytes([stream[filepass_pos + 2], stream[filepass_pos + 3]]) as usize;
    let total_filepass_len = 4 + filepass_size;
    let mut out = Vec::with_capacity(stream.len() - total_filepass_len);
    out.extend_from_slice(&stream[..filepass_pos]);
    out.extend_from_slice(&stream[filepass_pos + total_filepass_len..]);

    let delta = total_filepass_len as u32;
    let mut cursor = 0;
    while cursor + 4 <= out.len() {
        let record_type = u16::from_le_bytes([out[cursor], out[cursor + 1]]);
        let size = u16::from_le_bytes([out[cursor + 2], out[cursor + 3]]) as usize;
        if record_type == BOUND_SHEET_8 && size >= 4 {
            let prev = u32::from_le_bytes([
                out[cursor + 4],
                out[cursor + 5],
                out[cursor + 6],
                out[cursor + 7],
            ]);
            let next = prev - delta;
            out[cursor + 4..cursor + 8].copy_from_slice(&next.to_le_bytes());
        }
        cursor += 4 + size;
    }
    out
}

fn run_round_trip(variant: XlsEncryptionVariant) {
    let plain = build_plaintext_stream();
    let original_lbplypos = read_lbplypos(&plain);

    let encrypted = encrypt_workbook_stream(&plain, PASSWORD, variant)
        .expect("encrypt_workbook_stream should succeed");

    let filepass_size = find_filepass_record_size(&encrypted);
    assert_eq!(
        encrypted.len(),
        plain.len() + filepass_size,
        "ciphertext length must equal plaintext length plus FilePass record size",
    );

    let bumped = read_lbplypos(&encrypted);
    assert_eq!(
        bumped,
        original_lbplypos + filepass_size as u32,
        "BoundSheet8.lbPlyPos must be bumped by FilePass record size",
    );

    let decrypted = decrypt_workbook_stream(&encrypted, PASSWORD)
        .expect("decrypt_workbook_stream should succeed");

    let bof_size = u16::from_le_bytes([decrypted[2], decrypted[3]]) as usize;
    let filepass_pos = 4 + bof_size;
    let stripped = strip_neutered_filepass(&decrypted, filepass_pos);
    assert_eq!(
        stripped, plain,
        "round-trip after stripping the neutered FilePass must recover original plaintext",
    );
}

#[test]
fn round_trip_xor_obfuscation() {
    run_round_trip(XlsEncryptionVariant::Xor);
}

#[test]
fn round_trip_rc4_legacy() {
    run_round_trip(XlsEncryptionVariant::Rc4Legacy);
}

#[test]
fn round_trip_rc4_cryptoapi_128() {
    run_round_trip(XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 });
}

#[test]
fn round_trip_rc4_cryptoapi_40() {
    run_round_trip(XlsEncryptionVariant::Rc4CryptoApi { key_bits: 40 });
}

#[test]
fn wrong_password_yields_bad_password_for_rc4_legacy() {
    let plain = build_plaintext_stream();
    let encrypted = encrypt_workbook_stream(&plain, PASSWORD, XlsEncryptionVariant::Rc4Legacy)
        .expect("encrypt_workbook_stream should succeed");
    let err = decrypt_workbook_stream(&encrypted, "wrong-password").unwrap_err();
    assert!(matches!(err, CryptoError::BadPassword), "got {err:?}");
}

#[test]
fn wrong_password_yields_bad_password_for_rc4_cryptoapi() {
    let plain = build_plaintext_stream();
    let encrypted = encrypt_workbook_stream(
        &plain,
        PASSWORD,
        XlsEncryptionVariant::Rc4CryptoApi { key_bits: 128 },
    )
    .expect("encrypt_workbook_stream should succeed");
    let err = decrypt_workbook_stream(&encrypted, "wrong-password").unwrap_err();
    assert!(matches!(err, CryptoError::BadPassword), "got {err:?}");
}

#[test]
fn rejects_stream_without_bof() {
    let mut bad = Vec::new();
    put_record(&mut bad, EOF, &[]);
    let err = encrypt_workbook_stream(&bad, PASSWORD, XlsEncryptionVariant::Rc4Legacy)
        .expect_err("must reject non-BOF first record");
    assert!(matches!(err, CryptoError::InvalidFormat(_)), "got {err:?}");
}

#[test]
fn rejects_stream_with_existing_filepass() {
    let mut bad = Vec::new();
    put_record(&mut bad, BOF, &[0x06, 0x00, 0x05, 0x00, 0, 0, 0, 0]);
    let mut filepass_body = Vec::new();
    put_u16(&mut filepass_body, 0);
    put_u32(&mut filepass_body, 0);
    put_record(&mut bad, FILEPASS_RECORD_TYPE, &filepass_body);
    put_record(&mut bad, EOF, &[]);
    let err = encrypt_workbook_stream(&bad, PASSWORD, XlsEncryptionVariant::Rc4Legacy)
        .expect_err("must reject pre-existing FilePass");
    assert!(matches!(err, CryptoError::InvalidFormat(_)), "got {err:?}");
}

#[test]
fn xor_rejects_overlong_password() {
    let plain = build_plaintext_stream();
    let err = encrypt_workbook_stream(&plain, "this_is_sixteen_", XlsEncryptionVariant::Xor)
        .expect_err("XOR must reject 16-char passwords");
    assert!(matches!(err, CryptoError::InvalidFormat(_)), "got {err:?}");
}
