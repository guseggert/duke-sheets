//! # duke-sheets-xlsx
//!
//! XLSX (Office Open XML) reader and writer for duke-sheets.

pub mod error;
pub mod reader;
pub mod writer;

mod styles;

pub use error::{XlsxError, XlsxResult};
pub use reader::XlsxReader;
pub use writer::XlsxWriter;
