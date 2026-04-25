//! # duke-sheets-crypto
//!
//! Office document encryption and decryption primitives for duke-sheets.
//!
//! This crate handles the wire-level crypto for password-protected Excel
//! files across both the binary XLS format and OOXML (XLSX). It knows about
//! bytes-in, bytes-out decryption and encryption — not about spreadsheet
//! semantics. The XLS and XLSX crates call into this crate when they
//! encounter an encrypted workbook.
//!
//! Supported variants (as they land):
//!
//! | Variant | Read | Write |
//! |---|---|---|
//! | OOXML Agile (AES-CBC + HMAC) | planned | planned |
//! | OOXML Standard (AES-ECB) | planned | planned |
//! | OOXML Binary RC4 CryptoAPI | planned | planned |
//! | XLS RC4 CryptoAPI (SHA-1 KDF) | **implemented** | planned |
//! | XLS Legacy RC4 (MD5 KDF) | planned | planned |
//! | XLS XOR Obfuscation | planned | planned |

pub mod error;
pub mod ooxml;
pub(crate) mod password;
pub mod xls;

pub use error::{CryptoError, CryptoResult};
