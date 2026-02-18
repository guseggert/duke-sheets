//! Workbook handle — ergonomic API for working with an Excel workbook via the bridge.

use excel_com_protocol::{CellValue, SheetRef};

use crate::bridge::{linux_to_wine_path, BridgeError, ExcelBridge};

/// A handle to an open workbook in the Excel COM bridge.
///
/// Operations on this workbook are forwarded to the bridge process.
/// By default, operations target the first worksheet (index 0).
pub struct Workbook<'a> {
    bridge: &'a ExcelBridge,
    handle: u64,
    /// The active sheet for shorthand methods. Defaults to index 0.
    active_sheet: SheetRef,
}

impl<'a> Workbook<'a> {
    pub(crate) fn new(bridge: &'a ExcelBridge, handle: u64) -> Self {
        Self {
            bridge,
            handle,
            active_sheet: SheetRef::Index(0),
        }
    }

    /// Get the internal handle ID.
    pub fn handle(&self) -> u64 {
        self.handle
    }

    /// Set the active sheet for shorthand methods (by 0-based index).
    pub fn set_active_sheet_index(&mut self, index: u32) {
        self.active_sheet = SheetRef::Index(index);
    }

    /// Set the active sheet for shorthand methods (by name).
    pub fn set_active_sheet_name(&mut self, name: impl Into<String>) {
        self.active_sheet = SheetRef::Name(name.into());
    }

    // -- Shorthand methods that use the active sheet --

    /// Set a cell's value on the active sheet.
    ///
    /// Accepts anything that converts to CellValue:
    /// - `&str` / `String` -> String value
    /// - `f64`, `i32`, etc. -> Number value
    /// - `bool` -> Boolean value
    pub fn set_cell_value(
        &self,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_value(self.handle, self.active_sheet.clone(), cell, value.into())
    }

    /// Set a cell's formula on the active sheet (e.g., "=SUM(A1:A10)").
    pub fn set_cell_formula(&self, cell: &str, formula: &str) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_formula(self.handle, self.active_sheet.clone(), cell, formula)
    }

    /// Get a cell's computed value from the active sheet.
    pub fn get_cell_value(&self, cell: &str) -> Result<CellValue, BridgeError> {
        self.bridge
            .get_cell_value(self.handle, self.active_sheet.clone(), cell)
    }

    /// Get a cell's formula from the active sheet (empty string if no formula).
    pub fn get_cell_formula(&self, cell: &str) -> Result<String, BridgeError> {
        self.bridge
            .get_cell_formula(self.handle, self.active_sheet.clone(), cell)
    }

    // -- Sheet-specific methods --

    /// Set a cell value on a specific sheet (by index).
    pub fn set_cell_value_on_sheet(
        &self,
        sheet: SheetRef,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<(), BridgeError> {
        self.bridge
            .set_cell_value(self.handle, sheet, cell, value.into())
    }

    /// Get a cell value from a specific sheet.
    pub fn get_cell_value_on_sheet(
        &self,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<CellValue, BridgeError> {
        self.bridge.get_cell_value(self.handle, sheet, cell)
    }

    // -- File operations --

    /// Save the workbook to a file path.
    ///
    /// Accepts a Linux path — it will be automatically converted to a WINE path.
    /// Format is inferred from the extension (.xlsx, .xls, .csv).
    pub fn save(&self, path: &str) -> Result<(), BridgeError> {
        let wine_path = linux_to_wine_path(std::path::Path::new(path));
        self.bridge.save_workbook(self.handle, &wine_path)
    }

    /// Save the workbook using a raw Windows/WINE path (no conversion).
    pub fn save_raw_path(&self, wine_path: &str) -> Result<(), BridgeError> {
        self.bridge.save_workbook(self.handle, wine_path)
    }

    /// Close the workbook without saving.
    pub fn close(self) -> Result<(), BridgeError> {
        self.bridge.close_workbook(self.handle)
    }
}

// From<T> for CellValue impls live in excel-com-protocol crate
