//! # duke-sheets-ffi
//!
//! C FFI bindings for duke-sheets.
//!
//! This crate provides a C-compatible API for using duke-sheets from other languages.

mod cell;
mod error;
mod handles;
mod workbook;
mod worksheet;

pub use error::*;
pub use handles::Handle;
