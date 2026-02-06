//! Cell-related types and utilities
//!
//! This module contains:
//! - [`CellValue`] - The value stored in a cell
//! - [`CellAddress`] - A cell's location (e.g., "A1")
//! - [`CellRange`] - A range of cells (e.g., "A1:B10")
//! - [`CellData`] - Complete cell data including value and style

mod address;
mod storage;
mod value;

pub use address::{CellAddress, CellRange};
pub use storage::{CellData, CellStorage, SpillInfo, StorageMode};
pub use value::{CellError, CellValue, SharedString, StringPool};
