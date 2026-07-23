//! # duke-sheets-xlsx
//!
//! XLSX (Office Open XML) reader and writer for duke-sheets.

pub mod error;
pub mod reader;
pub mod writer;

mod opc;
mod styles;

pub use error::{XlsxError, XlsxResult};
pub use opc::{XlsxDiagnostic, XlsxDiagnosticCode, XlsxDiagnosticSeverity, XlsxPackagePolicy};
pub use reader::{XlsxReadOptions, XlsxReadReport, XlsxReader};
pub use writer::{EncryptionProfile, XlsxWriter};
