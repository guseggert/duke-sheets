//! Excel COM automation via a bridge server running in a Windows VM.
//!
//! This crate connects to a C# bridge server over TCP that provides generic
//! COM proxy operations. The server runs inside a QEMU/KVM Windows VM with
//! a real Excel installation, enabling high-fidelity parity testing.
//!
//! # Architecture
//!
//! ```text
//! Linux host (this crate)
//!     └── TCP (NDJSON) ──► Windows VM (QEMU/KVM)
//!                              └── ExcelBridgeServer.exe (C#)
//!                                    └── COM: Excel.Application
//! ```
//!
//! The protocol is a generic COM proxy with just 5 commands: `Init`, `Get`,
//! `Set`, `Invoke`, `Release`, and `Shutdown`. All Excel-specific knowledge
//! lives in this client crate — the server never needs modification when new
//! Excel features are added.
//!
//! # Example
//!
//! ```rust,no_run
//! use duke_sheets_excel_com::{ExcelBridge, ExcelBridgeConfig};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     // Connect to bridge server running in the VM
//!     let bridge = ExcelBridge::connect(ExcelBridgeConfig::default())?;
//!
//!     let wb = bridge.create_workbook()?;
//!     wb.set_cell_value("A1", "Hello")?;
//!     wb.set_cell_value("B1", 42.0)?;
//!     wb.set_cell_formula("C1", "=B1*2")?;
//!     bridge.recalculate()?;
//!     let val = wb.get_cell_value("C1")?;
//!     println!("C1 = {val}");
//!
//!     // Save via shared SMB mount (QEMU user networking)
//!     wb.save(r"\\10.0.2.4\qemu\output.xlsx")?;
//!     bridge.shutdown()?;
//!     Ok(())
//! }
//! ```

mod bridge;
mod workbook;

pub use bridge::{BridgeError, ExcelBridge, ExcelBridgeConfig};
pub use excel_com_protocol::{CellValue, ChainStep, SheetRef};
pub use workbook::Workbook;
