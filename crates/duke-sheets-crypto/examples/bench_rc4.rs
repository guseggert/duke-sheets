//! Micro-bench: how fast is XLS RC4 decryption?
//!
//! Run with: `cargo run --release -p duke-sheets-crypto --example bench_rc4`.
//! Requires the fixture at `tests/fixtures/xls_rc4_cryptoapi.xls` (run
//! `mise run crypto:fixtures` first).
//!
//! Measures: per-block KDF time for both the legacy MD5 and CryptoAPI
//! SHA-1 paths, plus full-stream decrypt throughput on the real
//! fixture and on a synthetic 1 MB stream (steady-state RC4 + per-block
//! KDF, isolated from CFB plumbing).

use std::time::Instant;

use duke_sheets_crypto::xls::{rc4_cryptoapi, rc4_legacy};

const FIXTURE_PASSWORD: &str = "duke-test-pw";

fn main() {
    let mut path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests/fixtures/xls_rc4_cryptoapi.xls");
    let data = match std::fs::read(&path) {
        Ok(d) => d,
        Err(_) => {
            eprintln!(
                "SKIP: {} not present - run `mise run crypto:fixtures`",
                path.display()
            );
            return;
        }
    };
    eprintln!("loaded {} bytes from {}", data.len(), path.display());

    // KDF cost in isolation - the per-1024-byte-block re-keying cost.
    let salt = [0u8; 16];
    let iters = 10_000;

    let t = Instant::now();
    for i in 0..iters {
        std::hint::black_box(rc4_legacy::make_key(FIXTURE_PASSWORD, &salt, i));
    }
    eprintln!(
        "legacy RC4 MD5 KDF: {:?} per block-key derivation",
        t.elapsed() / iters as u32
    );

    let t = Instant::now();
    for i in 0..iters {
        std::hint::black_box(rc4_cryptoapi::make_key(FIXTURE_PASSWORD, &salt, 128, i));
    }
    eprintln!(
        "RC4 CryptoAPI SHA-1 KDF (128-bit): {:?} per block-key derivation",
        t.elapsed() / iters as u32
    );

    let workbook_stream = extract_workbook_stream(&data).expect("parse CFB");

    // Real fixture decrypt.
    let iters = 1000;
    let t = Instant::now();
    for _ in 0..iters {
        let decrypted =
            duke_sheets_crypto::xls::decrypt_workbook_stream(&workbook_stream, FIXTURE_PASSWORD)
                .expect("decrypt");
        std::hint::black_box(decrypted);
    }
    let per = t.elapsed() / iters as u32;
    let mbps = (workbook_stream.len() as f64 / 1_000_000.0) / per.as_secs_f64();
    eprintln!(
        "decrypt real {}-byte stream: {:?} per iter ({:.1} MB/s)",
        workbook_stream.len(),
        per,
        mbps
    );

    // Synthetic 1 MB stream to extract steady-state per-byte throughput
    // independent of the per-call KDF setup. We pre-compute a verifier
    // hash so verify_password passes.
    let synthetic = vec![0u8; 1_000_000];
    let params = synthetic_legacy_params(FIXTURE_PASSWORD, salt);
    let iters = 100;
    let t = Instant::now();
    for _ in 0..iters {
        let decrypted = rc4_legacy::decrypt_workbook_stream(&synthetic, FIXTURE_PASSWORD, &params)
            .expect("decrypt synthetic");
        std::hint::black_box(decrypted);
    }
    let per = t.elapsed() / iters as u32;
    let mbps = (synthetic.len() as f64 / 1_000_000.0) / per.as_secs_f64();
    eprintln!(
        "decrypt synthetic 1MB stream: {:?} per iter ({:.0} MB/s steady-state)",
        per, mbps
    );
}

fn synthetic_legacy_params(password: &str, salt: [u8; 16]) -> rc4_legacy::Rc4LegacyParams {
    use md5::{Digest as _, Md5};
    use rc4::{KeyInit, Rc4, StreamCipher};

    let key = rc4_legacy::make_key(password, &salt, 0);
    let mut rc4 = Rc4::new_from_slice(&key).unwrap();
    let mut verifier = [0u8; 16];
    rc4.apply_keystream(&mut verifier);
    let mut md = Md5::new();
    md.update(verifier);
    let h: [u8; 16] = md.finalize().into();
    let mut h_enc = h;
    rc4.apply_keystream(&mut h_enc);
    rc4_legacy::Rc4LegacyParams {
        salt,
        encrypted_verifier: [0u8; 16],
        encrypted_verifier_hash: h_enc,
    }
}

fn extract_workbook_stream(data: &[u8]) -> Result<Vec<u8>, String> {
    if data[0..8] != [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        return Err("not CFB".into());
    }
    let sector_shift = u16::from_le_bytes([data[30], data[31]]);
    let sector_size = 1usize << sector_shift;
    let first_dir_sect = u32::from_le_bytes([data[48], data[49], data[50], data[51]]) as usize;
    let dir_off = (first_dir_sect + 1) * sector_size;
    let dir = &data[dir_off..dir_off + sector_size];

    let mut wb_size = 0usize;
    let mut root_start = 0usize;
    for i in 0..dir.len() / 128 {
        let e = i * 128;
        let nlen = u16::from_le_bytes([dir[e + 64], dir[e + 65]]) as usize;
        if nlen == 0 {
            continue;
        }
        let name_bytes = &dir[e..e + nlen.saturating_sub(2)];
        let name_u16: Vec<u16> = name_bytes
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        let name = String::from_utf16_lossy(&name_u16);
        let start =
            u32::from_le_bytes([dir[e + 116], dir[e + 117], dir[e + 118], dir[e + 119]]) as usize;
        let sz =
            u32::from_le_bytes([dir[e + 120], dir[e + 121], dir[e + 122], dir[e + 123]]) as usize;
        if name == "Workbook" {
            wb_size = sz;
        } else if name == "Root Entry" {
            root_start = start;
        }
    }

    // Fixture's Workbook is in the mini-stream (< 4096 B); take it
    // straight from the root entry's start sector.
    let mini_off = (root_start + 1) * sector_size;
    Ok(data[mini_off..mini_off + wb_size].to_vec())
}
