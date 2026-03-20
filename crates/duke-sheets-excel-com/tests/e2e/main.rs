//! End-to-end tests for duke-sheets-excel-com.
//!
//! Each test connects to a running Excel COM bridge server in a QEMU/KVM
//! Windows VM, builds a spreadsheet via real Excel, saves to the VM,
//! pulls the file via WinRM, reads it back with `XlsxReader`, and asserts.
//!
//! ## Requirements
//!
//! A Windows VM with the bridge server running on `localhost:9876`:
//!
//! ```bash
//! bash tools/vm/qemu-start.sh
//! ```
//!
//! Tests will fail if the bridge is not available.

mod alignment;
mod border_styles;
mod comments;
mod common;
mod conditional_format;
mod data_types;
mod data_validation;
mod dimensions;
mod fill_styles;
mod font_styles;
mod formula_parity;
mod merged_cells;
mod number_formats;
mod rich_text;
mod roundtrip;
mod selections;
mod smoke;
mod writing;
mod xls_reader;

pub use common::*;
