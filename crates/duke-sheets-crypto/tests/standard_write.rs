use duke_sheets_crypto::ooxml::standard::{encrypt, StandardWriteOptions};
use duke_sheets_crypto::ooxml::{decrypt as decrypt_ooxml, detect_variant, OoxmlVariant};
use duke_sheets_crypto::CryptoError;
use duke_sheets_xls::cfb::CompoundFileBuilder;

fn sample_zip() -> Vec<u8> {
    let mut buf = Vec::new();
    {
        use std::io::Write as _;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let opts = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("hello.txt", opts).unwrap();
        zip.write_all(b"hello, standard world").unwrap();
        zip.finish().unwrap();
    }
    buf
}

fn assemble_envelope(
    parts: &duke_sheets_crypto::ooxml::standard::StandardEnvelopeParts,
) -> Vec<u8> {
    let mut b = CompoundFileBuilder::new();
    b.add_stream("/EncryptionInfo", parts.encryption_info.clone())
        .unwrap();
    b.add_stream("/EncryptedPackage", parts.encrypted_package.clone())
        .unwrap();
    b.build().unwrap()
}

fn extract_streams(envelope: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(std::io::Cursor::new(envelope))
        .expect("envelope is a valid CFB");
    let info = cfb.read_stream("/EncryptionInfo").unwrap();
    let pkg = cfb.read_stream("/EncryptedPackage").unwrap();
    (info, pkg)
}

#[test]
fn standard_encrypt_emits_standard_header() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "pw", &StandardWriteOptions::default()).expect("encrypt");
    assert_eq!(
        detect_variant(&parts.encryption_info).expect("detect"),
        OoxmlVariant::Standard,
        "EncryptionInfo header should identify as Standard"
    );
}

#[test]
fn standard_encrypt_decrypt_round_trip_yields_original_bytes() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "pw", &StandardWriteOptions::default()).expect("encrypt");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt");
    assert_eq!(decrypted, plain);
}

#[test]
fn standard_encrypt_decrypt_with_wrong_password_yields_bad_password() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "real", &StandardWriteOptions::default()).expect("encrypt");
    let err = decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "wrong")
        .expect_err("must reject wrong password");
    assert!(matches!(err, CryptoError::BadPassword));
}

#[test]
fn standard_encrypt_handles_misaligned_plaintext() {
    let plain = b"abc".to_vec();
    let parts = encrypt(&plain, "pw", &StandardWriteOptions::default()).expect("encrypt");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt");
    assert_eq!(decrypted, plain);
}

#[test]
fn standard_encrypt_handles_empty_plaintext() {
    let plain: Vec<u8> = Vec::new();
    let parts = encrypt(&plain, "pw", &StandardWriteOptions::default()).expect("encrypt");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt");
    assert!(decrypted.is_empty());
}

#[test]
fn standard_encrypt_uses_random_salt() {
    let plain = sample_zip();
    let p1 = encrypt(&plain, "pw", &StandardWriteOptions::default()).unwrap();
    let p2 = encrypt(&plain, "pw", &StandardWriteOptions::default()).unwrap();
    assert_ne!(
        p1.encryption_info, p2.encryption_info,
        "EncryptionInfo must differ across encrypt runs (random salt + verifier)"
    );
}

#[test]
fn standard_encrypt_aes128_round_trip() {
    let plain = sample_zip();
    let opts = StandardWriteOptions {
        key_bits: 128,
        ..StandardWriteOptions::default()
    };
    let parts = encrypt(&plain, "pw", &opts).unwrap();
    let decrypted = decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").unwrap();
    assert_eq!(decrypted, plain);
}

#[test]
fn standard_encrypt_aes192_round_trip() {
    let plain = sample_zip();
    let opts = StandardWriteOptions {
        key_bits: 192,
        ..StandardWriteOptions::default()
    };
    let parts = encrypt(&plain, "pw", &opts).unwrap();
    let decrypted = decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").unwrap();
    assert_eq!(decrypted, plain);
}

#[test]
fn standard_assembled_envelope_round_trips_via_cfb_reader() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "pw", &StandardWriteOptions::default()).unwrap();
    let envelope = assemble_envelope(&parts);
    assert_eq!(
        &envelope[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
    );
    let (info, pkg) = extract_streams(&envelope);
    assert_eq!(detect_variant(&info).unwrap(), OoxmlVariant::Standard);
    let decrypted = decrypt_ooxml(&info, &pkg, "pw").unwrap();
    assert_eq!(decrypted, plain);
}
