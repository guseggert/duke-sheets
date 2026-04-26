//! OS-level random byte generation for salts, IVs, and key material.
//!
//! All write paths in this crate must source ciphertext-sensitive bytes
//! from the operating system (`OsRng` via `getrandom`), never from a
//! seedable PRNG. This module is the single funnel so reviewers can
//! audit randomness usage in one place.

use crate::error::{CryptoError, CryptoResult};

/// Fill `buf` with cryptographically-secure random bytes from the OS.
///
/// Returns [`CryptoError::Io`] if the platform RNG fails. On WASM
/// targets, this requires the consumer to enable `getrandom`'s `js`
/// feature; the `bindings/wasm` crate does so via its own `Cargo.toml`.
pub(crate) fn fill_random(buf: &mut [u8]) -> CryptoResult<()> {
    getrandom::getrandom(buf).map_err(|e| {
        CryptoError::Io(std::io::Error::new(
            std::io::ErrorKind::Other,
            format!("OS random source failed: {e}"),
        ))
    })
}

/// Return `n` random bytes as a fresh `Vec<u8>`.
pub(crate) fn random_bytes(n: usize) -> CryptoResult<Vec<u8>> {
    let mut buf = vec![0u8; n];
    fill_random(&mut buf)?;
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn random_bytes_returns_requested_length() {
        let v = random_bytes(32).unwrap();
        assert_eq!(v.len(), 32);
    }

    #[test]
    fn random_bytes_changes_across_calls() {
        let a = random_bytes(32).unwrap();
        let b = random_bytes(32).unwrap();
        assert_ne!(
            a, b,
            "two 32-byte draws should differ with overwhelming probability"
        );
    }
}
