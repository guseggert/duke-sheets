//! LibreOffice automation bridge for duke-sheets.
//!
//! This crate provides a high-level Rust API for controlling LibreOffice via
//! the UNO Remote Protocol (URP). It can create/open spreadsheets, manipulate
//! cells, and save files — all without Python or any other scripting language.
//!
//! # Architecture
//!
//! ```text
//! Your Rust code (native Linux)
//!     └── LibreOfficeBridge (this crate)
//!           └── UrpConnection (libreoffice-urp crate)
//!                 └── TCP socket to LibreOffice
//! ```
//!
//! # Example
//!
//! ```rust,no_run
//! use duke_sheets_libreoffice::{LibreOfficeBridge, LibreOfficeConfig};
//!
//! # async fn example() -> duke_sheets_libreoffice::error::Result<()> {
//! // Start LibreOffice automatically
//! let mut bridge = LibreOfficeBridge::start(LibreOfficeConfig::default()).await?;
//!
//! // Create a new spreadsheet
//! let mut wb = bridge.create_workbook().await?;
//! wb.set_cell_value("A1", "Hello").await?;
//! wb.set_cell_value("B1", 42.0).await?;
//! wb.set_cell_formula("C1", "=B1*2").await?;
//!
//! // Save as XLSX
//! wb.save("/tmp/output.xlsx").await?;
//! wb.close().await?;
//!
//! bridge.shutdown().await?;
//! # Ok(())
//! # }
//! ```

pub mod bridge;
pub mod error;
pub mod uno_types;
pub mod workbook;

pub use bridge::{LibreOfficeBridge, LibreOfficeConfig};
pub use error::BridgeError;
pub use uno_types::{GradientSpec, StyleSpec};
pub use workbook::{CellValue, Workbook};
