//! Workbook handle — ergonomic API for working with an Excel workbook via the bridge.

use excel_com_protocol::{CellValue, SheetRef};

use crate::bridge::{BridgeError, ExcelBridge};

/// A handle to an open workbook in the Excel COM bridge.
///
/// Operations on this workbook are forwarded to the bridge server.
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
    /// Accepts anything that converts to `CellValue`:
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

    /// Set a cell value on a specific sheet.
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

    /// Save the workbook to a Windows file path.
    ///
    /// The path must be a Windows path visible to the VM. For files shared
    /// via QEMU SMB, use a UNC path like `\\10.0.2.4\qemu\output.xlsx`.
    ///
    /// Format is inferred from extension: `.xlsx` = 51, `.xls` = -4143, `.csv` = 6.
    pub fn save(&self, windows_path: &str) -> Result<(), BridgeError> {
        let format = infer_save_format(windows_path);
        self.bridge.save_workbook(self.handle, windows_path, format)
    }

    /// Save the workbook with an explicit Excel file format constant.
    pub fn save_as(&self, windows_path: &str, format: i32) -> Result<(), BridgeError> {
        self.bridge.save_workbook(self.handle, windows_path, format)
    }

    /// Close the workbook without saving.
    pub fn close(self) -> Result<(), BridgeError> {
        self.bridge.close_workbook(self.handle)
    }
}

/// Infer the Excel file format constant from a file extension.
///
/// - `.xlsx` -> 51 (xlOpenXMLWorkbook)
/// - `.xls`  -> -4143 (xlWorkbookNormal)
/// - `.csv`  -> 6 (xlCSV)
/// - other   -> 51 (default to xlsx)
fn infer_save_format(path: &str) -> i32 {
    let lower = path.to_lowercase();
    if lower.ends_with(".xlsx") {
        51
    } else if lower.ends_with(".xls") {
        -4143
    } else if lower.ends_with(".csv") {
        6
    } else {
        51
    }
}
