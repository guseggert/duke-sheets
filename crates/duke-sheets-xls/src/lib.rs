//! # duke-sheets-xls
//!
//! XLS (BIFF8) reader and writer for duke-sheets.
//!
//! This crate handles the legacy Excel binary format (.xls).

pub mod biff;
pub mod reader;
pub mod writer;
pub mod error;

pub use error::{XlsError, XlsResult};
