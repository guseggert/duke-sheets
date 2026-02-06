//! # duke-sheets-ffi
//!
//! C FFI bindings for duke-sheets.
//!
//! This crate provides a C-compatible API for using duke-sheets from other languages.

mod handles;
mod error;
mod workbook;
mod worksheet;
mod cell;

pub use error::*;
pub use handles::Handle;
