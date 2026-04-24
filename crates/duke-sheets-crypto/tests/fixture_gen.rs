//! Generate password-protected workbook fixtures for crypto tests.
//!
//! All tests in this file are `#[ignore]`-gated: they require a running
//! LibreOffice URP daemon on `127.0.0.1:2002` and produce files under
//! `tests/fixtures/`. Run with:
//!
//! ```sh
//! # Start LibreOffice (see crates/duke-sheets-xlsx/tests/e2e/common.rs
//! # for the Docker image we use in CI, or soffice --accept=... locally):
//! docker run --rm -d -p 2002:2002 -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp \
//!   duke-sheets-pyuno bash -c 'soffice --headless \
//!     --accept="socket,host=0.0.0.0,port=2002;urp;" & sleep infinity'
//!
//! cargo test -p duke-sheets-crypto --test fixture_gen -- --ignored --nocapture
//! ```
//!
//! Generated files land in `crates/duke-sheets-crypto/tests/fixtures/`.
//! That directory is gitignored so fixtures don't bloat the repo; the
//! unit/integration tests in the crypto phases depend on regenerating
//! them locally (or downloading from a canonical source - TBD).
//!
//! ## Fixture matrix
//!
//! For each encryption variant below we emit:
//!   - `{variant}.xlsx` or `.xls` - the encrypted file.
//!   - `{variant}.plain.xlsx` or `.xls` - an unencrypted reference with the
//!     same cell values, for sanity-check round-trips.
//!
//! The plaintext content is the same for every variant: two cells with
//! well-known values that Phase 1+ tests assert against.
//!
//! | Variant | File |
//! |---|---|
//! | OOXML Agile (LO default: AES-256 + SHA-512) | `agile_aes256.xlsx` |
//! | XLS RC4 CryptoAPI (LO default: 128-bit SHA-1 KDF) | `xls_rc4_cryptoapi.xls` |
//!
//! LO's UI only exposes its format defaults for password-encrypted saves;
//! the other variants in our support matrix (Standard, Binary RC4,
//! Legacy XLS RC4, XOR) need a different generation path (Office itself
//! via the COM bridge, or bespoke header construction). Those fixtures
//! land when those phases are ready to consume them.

use std::net::TcpStream;
use std::path::{Path, PathBuf};
use std::time::Duration;

use duke_sheets_libreoffice::bridge::LibreOfficeBridge;

/// Canonical password used for every encrypted fixture this test emits.
/// Kept short and ASCII to keep KDF cost down on the LO side; the crypto
/// path doesn't change with password length, and Phase 1+ tests that
/// need wrong-password cases can re-encrypt ad hoc.
pub const FIXTURE_PASSWORD: &str = "duke-test-pw";

/// Cell values written to every plaintext fixture (and to the plaintext
/// round-trip of every encrypted fixture). Phase 1+ tests assert these
/// show up after decryption.
pub const FIXTURE_CELL_A1: &str = "hello crypto";
pub const FIXTURE_CELL_B1: f64 = 42.0;

/// Directory where fixtures land. Gitignored.
fn fixtures_dir() -> PathBuf {
    let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    p.push("tests/fixtures");
    p
}

/// Path shared with the LibreOffice Docker container. When LO runs in a
/// container we must write to a bind-mounted directory so the host can
/// read the output back. Matches the pattern in
/// `crates/duke-sheets-xlsx/tests/e2e/common.rs`.
const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

fn lo_available() -> bool {
    TcpStream::connect_timeout(&"127.0.0.1:2002".parse().unwrap(), Duration::from_secs(2))
        .is_ok()
}

/// Copy a file produced by LibreOffice (inside the shared Docker volume)
/// into our gitignored fixtures directory.
fn move_into_fixtures(shared_path: &Path, target_name: &str) -> std::io::Result<PathBuf> {
    let dir = fixtures_dir();
    std::fs::create_dir_all(&dir)?;
    let target = dir.join(target_name);
    std::fs::copy(shared_path, &target)?;
    let _ = std::fs::remove_file(shared_path); // best effort
    Ok(target)
}

/// Variant selector for [`build_and_save`].
enum SaveVariant<'a> {
    Xlsx,
    XlsxWithPassword(&'a str),
    Xls,
    XlsWithPassword(&'a str),
}

/// Build a minimal workbook with the canonical cell values and save it to
/// the given absolute path via LibreOffice. Async closures over
/// short-lived `Workbook<'_>` borrows play poorly with rustc's lifetime
/// inference, so the save variant is selected by an enum rather than a
/// closure.
async fn build_and_save(path: &str, variant: SaveVariant<'_>) -> Result<(), String> {
    let mut bridge = LibreOfficeBridge::connect("127.0.0.1", 2002)
        .await
        .map_err(|e| format!("connect: {e}"))?;
    let mut wb = bridge
        .create_workbook()
        .await
        .map_err(|e| format!("create_workbook: {e}"))?;
    wb.set_cell_value("A1", FIXTURE_CELL_A1)
        .await
        .map_err(|e| format!("set A1: {e}"))?;
    wb.set_cell_value("B1", FIXTURE_CELL_B1)
        .await
        .map_err(|e| format!("set B1: {e}"))?;
    match variant {
        SaveVariant::Xlsx => wb.save(path).await.map_err(|e| format!("save: {e}"))?,
        SaveVariant::XlsxWithPassword(pw) => wb
            .save_with_password_xlsx(path, pw)
            .await
            .map_err(|e| format!("save_with_password_xlsx: {e}"))?,
        SaveVariant::Xls => wb
            .save_as_xls(path)
            .await
            .map_err(|e| format!("save_as_xls: {e}"))?,
        SaveVariant::XlsWithPassword(pw) => wb
            .save_with_password_xls(path, pw)
            .await
            .map_err(|e| format!("save_with_password_xls: {e}"))?,
    };
    // Don't call bridge.shutdown(): the container-hosted LibreOffice daemon
    // is shared across tests and should keep running. Dropping the bridge
    // closes this test's TCP connection without killing the daemon.
    drop(bridge);
    Ok(())
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn generate_ooxml_agile_aes256_fixture() {
    if !lo_available() {
        eprintln!("SKIP: LibreOffice URP not reachable on 127.0.0.1:2002");
        return;
    }
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");

    let shared_enc = PathBuf::from(format!("{SHARED_DIR}/agile_aes256.xlsx"));
    let shared_plain = PathBuf::from(format!("{SHARED_DIR}/agile_aes256.plain.xlsx"));

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();

    let enc_path = shared_enc.display().to_string();
    let plain_path = shared_plain.display().to_string();

    rt.block_on(build_and_save(
        &enc_path,
        SaveVariant::XlsxWithPassword(FIXTURE_PASSWORD),
    ))
    .unwrap();
    rt.block_on(build_and_save(&plain_path, SaveVariant::Xlsx))
        .unwrap();

    let enc = move_into_fixtures(&shared_enc, "agile_aes256.xlsx").expect("move enc");
    let plain = move_into_fixtures(&shared_plain, "agile_aes256.plain.xlsx").expect("move plain");

    // Sanity: encrypted file must start with CFB magic (it's a CFB-wrapped
    // encrypted package, not a plain ZIP).
    let enc_bytes = std::fs::read(&enc).expect("read enc");
    assert_eq!(
        &enc_bytes[0..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
        "encrypted .xlsx should be a CFB envelope (D0CF11E0...), got {:02X?}",
        &enc_bytes[0..8]
    );

    // Sanity: plaintext must start with ZIP magic.
    let plain_bytes = std::fs::read(&plain).expect("read plain");
    assert_eq!(
        &plain_bytes[0..4],
        &[0x50, 0x4B, 0x03, 0x04],
        "plain .xlsx should start with ZIP magic, got {:02X?}",
        &plain_bytes[0..4]
    );

    eprintln!("wrote:");
    eprintln!("  {}", enc.display());
    eprintln!("  {}", plain.display());
}

#[test]
#[ignore = "requires running LibreOffice on 127.0.0.1:2002"]
fn generate_xls_rc4_cryptoapi_fixture() {
    if !lo_available() {
        eprintln!("SKIP: LibreOffice URP not reachable on 127.0.0.1:2002");
        return;
    }
    std::fs::create_dir_all(SHARED_DIR).expect("create shared dir");

    let shared_enc = PathBuf::from(format!("{SHARED_DIR}/xls_rc4_cryptoapi.xls"));
    let shared_plain = PathBuf::from(format!("{SHARED_DIR}/xls_rc4_cryptoapi.plain.xls"));

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();

    let enc_path = shared_enc.display().to_string();
    let plain_path = shared_plain.display().to_string();

    rt.block_on(build_and_save(
        &enc_path,
        SaveVariant::XlsWithPassword(FIXTURE_PASSWORD),
    ))
    .unwrap();
    rt.block_on(build_and_save(&plain_path, SaveVariant::Xls))
        .unwrap();

    let enc = move_into_fixtures(&shared_enc, "xls_rc4_cryptoapi.xls").expect("move enc");
    let plain = move_into_fixtures(&shared_plain, "xls_rc4_cryptoapi.plain.xls").expect("move plain");

    // Both are CFB containers (XLS is always CFB, encrypted or not).
    for p in [&enc, &plain] {
        let bytes = std::fs::read(p).expect("read");
        assert_eq!(
            &bytes[0..8],
            &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
            "XLS file should start with CFB magic: {}",
            p.display()
        );
    }

    eprintln!("wrote:");
    eprintln!("  {}", enc.display());
    eprintln!("  {}", plain.display());
}
