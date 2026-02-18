//! Native Linux client library for Excel COM automation via a WINE bridge process.
//!
//! This crate spawns a Windows `.exe` under WINE that automates Excel through COM,
//! communicating over JSON-over-stdio. It provides an ergonomic Rust API for
//! creating/opening workbooks, reading/writing cells and formulas, recalculating,
//! and saving files.
//!
//! # Architecture
//!
//! ```text
//! Your Rust code (native Linux)
//!     └── ExcelBridge (this crate)
//!           └── spawns: wine excel-com-bridge.exe
//!                 └── COM: Excel.Application
//! ```
//!
//! # Example
//!
//! ```rust,no_run
//! use duke_sheets_excel_com::{ExcelBridge, ExcelBridgeConfig};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let bridge = ExcelBridge::start(ExcelBridgeConfig::default())?;
//!     let wb = bridge.create_workbook()?;
//!     wb.set_cell_value("A1", "Hello")?;
//!     wb.set_cell_value("B1", 42.0)?;
//!     wb.set_cell_formula("C1", "=B1*2")?;
//!     bridge.recalculate()?;
//!     let val = wb.get_cell_value("C1")?;
//!     println!("C1 = {val}");
//!     wb.save("output.xlsx")?;
//!     bridge.shutdown()?;
//!     Ok(())
//! }
//! ```

mod bridge;
mod workbook;

pub use bridge::{ExcelBridge, ExcelBridgeConfig};
pub use excel_com_protocol::{CellValue, SheetRef};
pub use workbook::Workbook;
