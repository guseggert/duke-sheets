//! Round-trip tests for OOXML Agile encryption write path.
//!
//! These tests cover the bytes-in/parts-out write API in
//! `duke_sheets_crypto::ooxml::agile::encrypt`. The CFB envelope is
//! assembled here using the in-house `duke-sheets-xls::cfb` writer so
//! tests can drive the same CFB-pack pipeline xlsx will use in
//! production. Cross-tool compatibility tests (LibreOffice, Excel COM)
//! live in dedicated `#[ignore]`-gated modules.

use duke_sheets_crypto::ooxml::agile::{
    encrypt, AgileEnvelopeParts, AgileWriteOptions, AGILE_REQUIRED_STORAGES,
};
use duke_sheets_crypto::ooxml::{decrypt as decrypt_ooxml, detect_variant, OoxmlVariant};
use duke_sheets_xls::cfb::CompoundFileBuilder;

fn assemble_envelope(parts: &AgileEnvelopeParts) -> Vec<u8> {
    let mut b = CompoundFileBuilder::new();
    for storage in AGILE_REQUIRED_STORAGES {
        b.add_storage(storage).expect("add storage");
    }
    for (path, bytes) in &parts.dataspaces {
        b.add_stream(path, bytes.clone())
            .expect("add dataspace stream");
    }
    b.add_stream("/EncryptionInfo", parts.encryption_info.clone())
        .expect("add EncryptionInfo");
    b.add_stream("/EncryptedPackage", parts.encrypted_package.clone())
        .expect("add EncryptedPackage");
    b.build().expect("CFB build")
}

/// A tiny but well-formed inner ZIP. The crypto layer is bytes-in
/// bytes-out — the ZIP only needs to be non-trivial enough to surface
/// off-by-one bugs in segment/padding code.
fn sample_zip() -> Vec<u8> {
    let mut buf = Vec::new();
    {
        use std::io::Write as _;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let opts = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("hello.txt", opts).unwrap();
        zip.write_all(b"hello, agile world").unwrap();
        zip.finish().unwrap();
    }
    buf
}

fn extract_streams(envelope: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let cfb = duke_sheets_xls::cfb::CompoundFile::open(std::io::Cursor::new(envelope))
        .expect("envelope is a valid CFB");
    assert!(cfb.exists("/EncryptionInfo"));
    assert!(cfb.exists("/EncryptedPackage"));
    let info = cfb.read_stream("/EncryptionInfo").unwrap();
    let pkg = cfb.read_stream("/EncryptedPackage").unwrap();
    (info, pkg)
}

#[test]
fn agile_encrypt_emits_cfb_envelope() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "test-pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let envelope = assemble_envelope(&parts);
    assert!(envelope.len() >= 8, "envelope too short");
    assert_eq!(
        &envelope[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "expected CFB magic at start of encrypted envelope"
    );
}

#[test]
fn agile_encrypt_envelope_has_agile_header() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "test-pw", &AgileWriteOptions::default()).expect("encrypt ok");
    assert_eq!(
        detect_variant(&parts.encryption_info).expect("detect variant"),
        OoxmlVariant::Agile,
        "EncryptionInfo header should identify as Agile"
    );
}

#[test]
fn agile_encrypt_decrypt_round_trip_yields_original_bytes() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "test-pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let decrypted = decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "test-pw")
        .expect("decrypt ok");
    assert_eq!(
        decrypted, plain,
        "round-trip must reproduce the original plaintext bytes exactly"
    );
}

#[test]
fn agile_encrypt_decrypt_with_wrong_password_yields_bad_password() {
    let plain = sample_zip();
    let parts = encrypt(&plain, "real-pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let err = decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "wrong-pw")
        .expect_err("must reject wrong password");
    assert!(
        matches!(err, duke_sheets_crypto::CryptoError::BadPassword),
        "expected BadPassword, got {err:?}"
    );
}

#[test]
fn agile_encrypt_handles_misaligned_plaintext() {
    // Plaintext whose length is NOT a multiple of 16 should still encrypt
    // cleanly (internally padded) and decrypt back to the original
    // (padding stripped via the totalSize prefix).
    let plain = b"abc".to_vec();
    let parts = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}

#[test]
fn agile_encrypt_handles_empty_plaintext() {
    let plain: Vec<u8> = Vec::new();
    let parts = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert!(decrypted.is_empty());
}

#[test]
fn agile_encrypt_handles_multi_segment_plaintext() {
    // > 4096 bytes forces multiple AES-CBC segments with per-segment IVs.
    let plain: Vec<u8> = (0..20_000u32)
        .map(|i| (i.wrapping_mul(0x9E3779B1) >> 24) as u8)
        .collect();
    let parts = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}

#[test]
fn agile_encrypt_uses_random_salts() {
    // Two encryptions of identical plaintext + password must produce
    // different ciphertext (random salts/IVs/keys).
    let plain = sample_zip();
    let p1 = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt 1");
    let p2 = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt 2");
    assert_ne!(
        p1.encrypted_package, p2.encrypted_package,
        "encrypted package must differ across encryption runs"
    );
    let env1 = assemble_envelope(&p1);
    let env2 = assemble_envelope(&p2);
    assert_ne!(env1, env2, "full envelopes must differ across runs");
}

#[test]
fn agile_encrypt_aes128_round_trip() {
    let plain = sample_zip();
    let opts = AgileWriteOptions {
        key_bits: 128,
        ..AgileWriteOptions::default()
    };
    let parts = encrypt(&plain, "pw", &opts).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}

#[test]
fn agile_encrypt_aes192_round_trip() {
    let plain = sample_zip();
    let opts = AgileWriteOptions {
        key_bits: 192,
        ..AgileWriteOptions::default()
    };
    let parts = encrypt(&plain, "pw", &opts).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}

#[test]
fn agile_encrypt_low_spincount_round_trip() {
    // Use a low spin count to keep the test fast; the KDF still has to
    // run and produce a valid intermediate key.
    let plain = sample_zip();
    let opts = AgileWriteOptions {
        spin_count: 100,
        ..AgileWriteOptions::default()
    };
    let parts = encrypt(&plain, "pw", &opts).expect("encrypt ok");
    let decrypted =
        decrypt_ooxml(&parts.encryption_info, &parts.encrypted_package, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}

#[test]
fn agile_assembled_envelope_round_trips_via_cfb_reader() {
    // Sanity: the assembled envelope is parseable by our CFB reader,
    // its EncryptionInfo decodes to Agile, and decrypt of the recovered
    // streams yields the original plaintext.
    let plain = sample_zip();
    let parts = encrypt(&plain, "pw", &AgileWriteOptions::default()).expect("encrypt ok");
    let envelope = assemble_envelope(&parts);
    let (info, pkg) = extract_streams(&envelope);
    assert_eq!(detect_variant(&info).unwrap(), OoxmlVariant::Agile);
    let decrypted = decrypt_ooxml(&info, &pkg, "pw").expect("decrypt ok");
    assert_eq!(decrypted, plain);
}
