//! XLSX error types

use thiserror::Error;

/// Result type for XLSX operations
pub type XlsxResult<T> = std::result::Result<T, XlsxError>;

/// Errors that can occur during XLSX reading/writing
#[derive(Debug, Error)]
pub enum XlsxError {
    /// IO error
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP error
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML error
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Invalid file format
    #[error("Invalid XLSX format: {0}")]
    InvalidFormat(String),

    /// Missing required part
    #[error("Missing required part: {0}")]
    MissingPart(String),

    /// Workbook is encrypted and no password was supplied, or decryption
    /// is not supported for the detected variant.
    ///
    /// Distinct from [`XlsxError::BadPassword`]: this means "decryption
    /// was not attempted", while `BadPassword` means "decryption was
    /// attempted and the verifier rejected the password".
    #[error("Encrypted XLSX file: {0}")]
    Encrypted(String),

    /// A password was supplied but the file's password verifier rejected it.
    #[error("Incorrect password for encrypted XLSX file")]
    BadPassword,

    /// The file is encrypted with a variant this crate does not yet
    /// implement (e.g. certificate-based key encryptor).
    #[error("Unsupported XLSX encryption: {0}")]
    UnsupportedEncryption(String),

    /// The password decrypted successfully but the data integrity HMAC
    /// did not match. Encrypted package was modified after encryption.
    #[error("Encrypted XLSX file failed integrity check (HMAC mismatch)")]
    IntegrityCheckFailed,

    /// Parse error
    #[error("Parse error: {0}")]
    Parse(String),

    /// Core error
    #[error("Core error: {0}")]
    Core(#[from] duke_sheets_core::Error),
}

impl From<duke_sheets_crypto::CryptoError> for XlsxError {
    fn from(err: duke_sheets_crypto::CryptoError) -> Self {
        use duke_sheets_crypto::CryptoError;
        match err {
            CryptoError::BadPassword => XlsxError::BadPassword,
            CryptoError::MissingPassword => {
                XlsxError::Encrypted("workbook is encrypted but no password was supplied".into())
            }
            CryptoError::UnsupportedVariant(s) => XlsxError::UnsupportedEncryption(s),
            CryptoError::InvalidFormat(s) => XlsxError::InvalidFormat(format!("crypto: {s}")),
            CryptoError::IntegrityCheckFailed => XlsxError::IntegrityCheckFailed,
            CryptoError::Io(e) => XlsxError::Io(e),
        }
    }
}
