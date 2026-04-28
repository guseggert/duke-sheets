//! XLS error types

use thiserror::Error;

/// Result type for XLS operations
pub type XlsResult<T> = std::result::Result<T, XlsError>;

/// Errors that can occur during XLS reading/writing
#[derive(Debug, Error)]
pub enum XlsError {
    /// IO error (also covers CFB errors which use std::io::Error)
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// Invalid file format
    #[error("Invalid XLS format: {0}")]
    InvalidFormat(String),

    /// Unsupported version
    #[error("Unsupported XLS version: {0}")]
    UnsupportedVersion(String),

    /// File is encrypted and no password was supplied, or decryption is not
    /// supported for the detected variant.
    ///
    /// Distinct from [`XlsError::BadPassword`]: this means "decryption was
    /// not attempted" (no password, or variant we can't handle), while
    /// `BadPassword` means "decryption was attempted and the verifier
    /// rejected the password".
    #[error("Encrypted XLS file: {0}")]
    Encrypted(String),

    /// A password was supplied but the file's password verifier rejected it.
    ///
    /// Callers can catch this specifically to prompt for a new password
    /// without confusing it with the "encrypted but no password given"
    /// case.
    #[error("Incorrect password for encrypted XLS file")]
    BadPassword,

    /// The file is encrypted with a variant this crate does not yet
    /// implement (e.g. certificate-based key encryptor).
    #[error("Unsupported XLS encryption: {0}")]
    UnsupportedEncryption(String),

    /// The password decrypted successfully but the data integrity HMAC
    /// did not match. Encrypted package was modified after encryption.
    #[error("Encrypted XLS file failed integrity check (HMAC mismatch)")]
    IntegrityCheckFailed,

    /// Parse error
    #[error("Parse error: {0}")]
    Parse(String),

    /// Core error
    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}

impl From<duke_sheets_crypto::CryptoError> for XlsError {
    fn from(err: duke_sheets_crypto::CryptoError) -> Self {
        use duke_sheets_crypto::CryptoError;
        match err {
            CryptoError::BadPassword => XlsError::BadPassword,
            CryptoError::MissingPassword => {
                XlsError::Encrypted("workbook is encrypted but no password was supplied".into())
            }
            CryptoError::UnsupportedVariant(s) => XlsError::UnsupportedEncryption(s),
            CryptoError::InvalidFormat(s) => XlsError::InvalidFormat(format!("crypto: {s}")),
            CryptoError::IntegrityCheckFailed => XlsError::IntegrityCheckFailed,
            CryptoError::Io(e) => XlsError::Io(e),
        }
    }
}
