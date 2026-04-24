//! Error types for Office document encryption/decryption.

use thiserror::Error;

/// Result type for operations in this crate.
pub type CryptoResult<T> = std::result::Result<T, CryptoError>;

/// Errors raised by encryption / decryption primitives.
///
/// This type is format-neutral: the XLS and XLSX readers map from their
/// format-specific error types into these variants when an encryption step
/// fails, and then re-wrap the `CryptoError` in their own top-level error.
#[derive(Debug, Error)]
pub enum CryptoError {
    /// A password was required but not supplied, or the supplied password
    /// was rejected by the file's verifier.
    ///
    /// Distinct from `MissingPassword` so that interactive callers can
    /// retry with a different value.
    #[error("incorrect password")]
    BadPassword,

    /// The file is encrypted but no password was supplied.
    #[error("workbook is encrypted but no password was supplied")]
    MissingPassword,

    /// The file uses an encryption variant this crate does not yet
    /// implement (e.g. certificate-based key encryptor, non-password
    /// key wrap, or a variant not in the supported matrix).
    #[error("unsupported encryption variant: {0}")]
    UnsupportedVariant(String),

    /// The encryption header is malformed or the ciphertext failed an
    /// integrity check (HMAC mismatch, truncated stream, invalid field
    /// size, etc).
    #[error("invalid encrypted format: {0}")]
    InvalidFormat(String),

    /// An I/O error occurred while reading or writing encrypted data.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}
