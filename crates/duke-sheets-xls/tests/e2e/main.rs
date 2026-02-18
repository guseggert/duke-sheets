//! E2E tests for XLS reader: create .xls fixtures via LibreOffice URP bridge,
//! read them back with XlsReader, and assert correctness.

mod common;
mod reading;

// Re-export common utilities for use in submodules
pub use common::*;
