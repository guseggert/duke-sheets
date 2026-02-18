//! # duke-sheets-xls
//!
//! XLS (BIFF8) reader for duke-sheets.
//!
//! This crate handles the legacy Excel binary format (.xls) used by
//! Excel 97, 2000, 2002, and 2003.
//!
//! # Example
//!
//! ```rust,no_run
//! use duke_sheets_xls::XlsReader;
//!
//! let workbook = XlsReader::read_file("input.xls").unwrap();
//! let sheet = workbook.worksheet(0).unwrap();
//! println!("{:?}", sheet.get_value("A1"));
//! ```

pub mod biff;
pub mod error;
pub mod reader;
pub mod styles;
pub mod writer;

pub use error::{XlsError, XlsResult};
pub use reader::XlsReader;
