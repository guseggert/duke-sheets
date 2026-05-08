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
mod chart_parity;
mod comments;
mod common;
mod conditional_format;
mod crypto_fixtures;
mod data_types;
mod data_validation;
mod dimensions;
mod encrypted_agile_compat;
mod encrypted_rc4_compat;
mod encrypted_standard_compat;
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
mod writing_xls;
mod writing_xlsb;
mod xls_reader;
mod xlsb_parity;

pub use common::*;
