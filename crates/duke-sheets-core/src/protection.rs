//! Legacy workbook and worksheet protection helpers.
//!
//! This module models Excel's workbook/sheet/protected-range protection.
//! It is not file encryption: passwords are stored as Excel's legacy 16-bit
//! verifier so readers and writers can round-trip existing files.

use crate::CellRange;

/// Workbook-level structure/window protection.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookProtection {
    /// Protect workbook structure (sheet add/delete/rename/reorder).
    pub structure: bool,
    /// Protect workbook windows.
    pub windows: bool,
    /// Excel legacy 16-bit password verifier.
    pub password_hash: Option<u16>,
}

impl WorkbookProtection {
    /// Create workbook protection with structure protection enabled.
    pub fn protected() -> Self {
        Self {
            structure: true,
            ..Default::default()
        }
    }

    /// Set the password from plaintext input, storing only the legacy verifier.
    pub fn with_password(mut self, password: &str) -> Self {
        self.password_hash = Some(hash_legacy_protection_password(password));
        self
    }

    /// Set a precomputed legacy password verifier.
    pub fn with_password_hash(mut self, password_hash: u16) -> Self {
        self.password_hash = Some(password_hash);
        self
    }
}

/// A protected editable range within a worksheet.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProtectedRange {
    /// Display name for the protected range.
    pub name: String,
    /// Cell ranges covered by this entry.
    pub ranges: Vec<CellRange>,
    /// Excel legacy 16-bit password verifier.
    pub password_hash: Option<u16>,
    /// Optional security descriptor text where the format exposes it.
    pub security_descriptor: Option<String>,
}

impl ProtectedRange {
    /// Create a protected range with a name and covered ranges.
    pub fn new(name: impl Into<String>, ranges: Vec<CellRange>) -> Self {
        Self {
            name: name.into(),
            ranges,
            ..Default::default()
        }
    }

    /// Set the password from plaintext input, storing only the legacy verifier.
    pub fn with_password(mut self, password: &str) -> Self {
        self.password_hash = Some(hash_legacy_protection_password(password));
        self
    }

    /// Set a precomputed legacy password verifier.
    pub fn with_password_hash(mut self, password_hash: u16) -> Self {
        self.password_hash = Some(password_hash);
        self
    }
}

/// Compute Excel's legacy 16-bit protection password verifier.
///
/// This is the verifier used by OOXML sheet/workbook/range protection and
/// BIFF PASSWORD records, not a cryptographic password hash.
pub fn hash_legacy_protection_password(password: &str) -> u16 {
    let mut hash = 0u16;
    for (idx, code_unit) in password.encode_utf16().enumerate() {
        let shift = (idx + 1) as u32;
        let mut value = (code_unit as u32) << shift;
        let rotated = value >> 15;
        value &= 0x7FFF;
        hash ^= (value | rotated) as u16;
    }
    hash ^= password.encode_utf16().count() as u16;
    hash ^ 0xCE4B
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn legacy_protection_password_hash_vectors() {
        assert_eq!(hash_legacy_protection_password(""), 0xCE4B);
        assert_eq!(hash_legacy_protection_password("password"), 0x83AF);
        assert_eq!(hash_legacy_protection_password("secret"), 0xDAA7);
        assert_eq!(hash_legacy_protection_password("duke"), 0xCA5B);
    }
}
