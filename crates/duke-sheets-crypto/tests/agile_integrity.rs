use duke_sheets_crypto::ooxml::agile::{
    decrypt_with_options as agile_decrypt_with_options, encrypt, AgileReadOptions,
    AgileWriteOptions,
};
use duke_sheets_crypto::ooxml::{decrypt, decrypt_with_options, DecryptOptions};
use duke_sheets_crypto::CryptoError;

fn build_parts(plain: &[u8], password: &str) -> (Vec<u8>, Vec<u8>) {
    let parts = encrypt(plain, password, &AgileWriteOptions::default()).expect("encrypt");
    (parts.encryption_info, parts.encrypted_package)
}

// Plaintext long enough that there's ciphertext past byte 40 to flip,
// but short enough to keep the KDF cheap. ~256 bytes -> 264 enc bytes.
fn long_plain() -> Vec<u8> {
    (0..256u32).map(|i| (i & 0xFF) as u8).collect()
}

#[test]
fn agile_decrypt_default_verifies_hmac_and_accepts_intact_package() {
    let plain = b"hello agile integrity".to_vec();
    let (info, pkg) = build_parts(&plain, "pw");
    let out = decrypt(&info, &pkg, "pw").expect("intact package must decrypt");
    assert_eq!(out, plain);
}

#[test]
fn agile_decrypt_default_rejects_tampered_package() {
    let plain = long_plain();
    let (info, mut pkg) = build_parts(&plain, "pw");
    pkg[40] ^= 0xFF;
    let err = decrypt(&info, &pkg, "pw").expect_err("tampered package must error");
    assert!(
        matches!(err, CryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}

#[test]
fn agile_decrypt_with_skip_integrity_check_accepts_tampered_package() {
    let plain = long_plain();
    let (info, mut pkg) = build_parts(&plain, "pw");
    pkg[40] ^= 0xFF;
    let out = decrypt_with_options(
        &info,
        &pkg,
        "pw",
        &DecryptOptions {
            skip_integrity_check: true,
        },
    )
    .expect("skip_integrity_check must let tampered packages through");
    assert_eq!(out.len(), plain.len());
}

#[test]
fn agile_decrypt_module_level_with_options_also_supports_skip_flag() {
    let plain = long_plain();
    let (info, mut pkg) = build_parts(&plain, "pw");
    pkg[40] ^= 0xFF;
    let err = agile_decrypt_with_options(
        &info,
        &pkg,
        "pw",
        &AgileReadOptions {
            skip_integrity_check: false,
        },
    )
    .expect_err("explicit skip=false must still verify");
    assert!(matches!(err, CryptoError::IntegrityCheckFailed));

    let out = agile_decrypt_with_options(
        &info,
        &pkg,
        "pw",
        &AgileReadOptions {
            skip_integrity_check: true,
        },
    )
    .expect("skip=true must accept tampered package");
    assert_eq!(out.len(), plain.len());
}

#[test]
fn agile_decrypt_wrong_password_still_errors_as_bad_password_not_integrity() {
    let plain = b"x".to_vec();
    let (info, pkg) = build_parts(&plain, "pw");
    let err = decrypt(&info, &pkg, "wrong-pw").expect_err("wrong password must error");
    assert!(
        matches!(err, CryptoError::BadPassword),
        "wrong password must surface BadPassword, not IntegrityCheckFailed: {err:?}"
    );
}

#[test]
fn agile_decrypt_default_rejects_tamper_in_size_prefix() {
    let plain = long_plain();
    let (info, mut pkg) = build_parts(&plain, "pw");
    pkg[0] = pkg[0].wrapping_add(1);
    let err = decrypt(&info, &pkg, "pw").expect_err("size-prefix tamper must error");
    assert!(
        matches!(err, CryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed for size-prefix tamper, got {err:?}"
    );
}
