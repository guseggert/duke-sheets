//! Excel-specific COM automation layer built on top of the generic IDispatch wrapper.

#![cfg(windows)]

use std::collections::HashMap;

use windows::Win32::System::Variant::VARIANT;

use excel_com_protocol::{CellError, CellValue, SheetRef};

use crate::dispatch::{
    variant_bool, variant_empty, variant_f64, variant_get_bool, variant_get_f64,
    variant_get_string, variant_i32, variant_is_empty, variant_is_error, variant_str,
    DispatchObject,
};

/// Manages an Excel.Application COM instance and its open workbooks.
pub struct ExcelApp {
    app: DispatchObject,
    workbooks_collection: DispatchObject,
    /// Map from our handle IDs to workbook dispatch objects.
    workbooks: HashMap<u64, DispatchObject>,
    next_handle: u64,
}

impl ExcelApp {
    /// Create a new Excel.Application instance via COM.
    pub fn new() -> Result<Self, String> {
        let app = DispatchObject::create_from_progid("Excel.Application")?;

        // Disable UI elements for automation
        app.set_property("Visible", variant_bool(false))?;
        app.set_property("DisplayAlerts", variant_bool(false))?;
        app.set_property("ScreenUpdating", variant_bool(false))?;

        let workbooks_collection = app.get_child("Workbooks")?;

        Ok(Self {
            app,
            workbooks_collection,
            workbooks: HashMap::new(),
            next_handle: 1,
        })
    }

    /// Create a new empty workbook. Returns the handle ID.
    pub fn create_workbook(&mut self) -> Result<u64, String> {
        let wb = self.workbooks_collection.invoke_child("Add", &[])?;
        let handle = self.next_handle;
        self.next_handle += 1;
        self.workbooks.insert(handle, wb);
        Ok(handle)
    }

    /// Open a workbook from a file path. Returns the handle ID.
    pub fn open_workbook(&mut self, path: &str) -> Result<u64, String> {
        let wb = self
            .workbooks_collection
            .invoke_child("Open", &[variant_str(path)])?;
        let handle = self.next_handle;
        self.next_handle += 1;
        self.workbooks.insert(handle, wb);
        Ok(handle)
    }

    /// Get a worksheet from a workbook.
    fn get_sheet(&self, wb_handle: u64, sheet: &SheetRef) -> Result<DispatchObject, String> {
        let wb = self
            .workbooks
            .get(&wb_handle)
            .ok_or_else(|| format!("Unknown workbook handle: {wb_handle}"))?;

        let sheets = wb.get_child("Worksheets")?;
        match sheet {
            SheetRef::Index(idx) => {
                // Excel worksheets are 1-based, our protocol uses 0-based
                let excel_index = (*idx as i32) + 1;
                sheets.get_indexed("Item", &variant_i32(excel_index))
            }
            SheetRef::Name(name) => sheets.get_indexed("Item", &variant_str(name)),
        }
    }

    /// Get a Range object for a cell reference.
    fn get_range(
        &self,
        wb_handle: u64,
        sheet: &SheetRef,
        cell_ref: &str,
    ) -> Result<DispatchObject, String> {
        let ws = self.get_sheet(wb_handle, sheet)?;
        ws.get_indexed("Range", &variant_str(cell_ref))
    }

    /// Set a cell's value.
    pub fn set_cell_value(
        &self,
        wb_handle: u64,
        sheet: &SheetRef,
        cell_ref: &str,
        value: &CellValue,
    ) -> Result<(), String> {
        let range = self.get_range(wb_handle, sheet, cell_ref)?;
        let variant = cell_value_to_variant(value);
        range.set_property("Value", variant)
    }

    /// Set a cell's formula.
    pub fn set_cell_formula(
        &self,
        wb_handle: u64,
        sheet: &SheetRef,
        cell_ref: &str,
        formula: &str,
    ) -> Result<(), String> {
        let range = self.get_range(wb_handle, sheet, cell_ref)?;
        range.set_property("Formula", variant_str(formula))
    }

    /// Get a cell's computed value.
    pub fn get_cell_value(
        &self,
        wb_handle: u64,
        sheet: &SheetRef,
        cell_ref: &str,
    ) -> Result<CellValue, String> {
        let range = self.get_range(wb_handle, sheet, cell_ref)?;
        let variant = range.get_property("Value")?;
        Ok(variant_to_cell_value(&variant))
    }

    /// Get a cell's formula (empty string if none).
    pub fn get_cell_formula(
        &self,
        wb_handle: u64,
        sheet: &SheetRef,
        cell_ref: &str,
    ) -> Result<String, String> {
        let range = self.get_range(wb_handle, sheet, cell_ref)?;
        let variant = range.get_property("Formula")?;
        match variant_get_string(&variant) {
            Some(s) => Ok(s),
            None => Ok(String::new()),
        }
    }

    /// Force a full recalculation.
    pub fn recalculate(&self) -> Result<(), String> {
        self.app.invoke_method("Calculate", &[])?;
        Ok(())
    }

    /// Save a workbook to a file path.
    pub fn save_workbook(&self, wb_handle: u64, path: &str) -> Result<(), String> {
        let wb = self
            .workbooks
            .get(&wb_handle)
            .ok_or_else(|| format!("Unknown workbook handle: {wb_handle}"))?;

        // Determine file format from extension
        // xlOpenXMLWorkbook = 51, xlWorkbookNormal (xls) = -4143, xlCSV = 6
        let format: i32 = if path.ends_with(".xlsx") {
            51
        } else if path.ends_with(".xls") {
            -4143
        } else if path.ends_with(".csv") {
            6
        } else {
            51 // default to xlsx
        };

        wb.invoke_method("SaveAs", &[variant_str(path), variant_i32(format)])?;
        Ok(())
    }

    /// Close a workbook without saving.
    pub fn close_workbook(&mut self, wb_handle: u64) -> Result<(), String> {
        let wb = self
            .workbooks
            .remove(&wb_handle)
            .ok_or_else(|| format!("Unknown workbook handle: {wb_handle}"))?;
        wb.invoke_method("Close", &[variant_bool(false)])?;
        Ok(())
    }

    /// Shut down: close all workbooks and quit Excel.
    pub fn shutdown(mut self) -> Result<(), String> {
        let handles: Vec<u64> = self.workbooks.keys().copied().collect();
        for h in handles {
            let _ = self.close_workbook(h);
        }
        self.app.invoke_method("Quit", &[])?;
        Ok(())
    }
}

/// Convert our protocol CellValue to a COM VARIANT.
fn cell_value_to_variant(value: &CellValue) -> VARIANT {
    match value {
        CellValue::Null => variant_empty(),
        CellValue::Bool(b) => variant_bool(*b),
        CellValue::Number(n) => variant_f64(*n),
        CellValue::String(s) => variant_str(s),
        CellValue::Error(_) => variant_empty(), // Can't set error values
    }
}

/// Convert a COM VARIANT to our protocol CellValue.
fn variant_to_cell_value(variant: &VARIANT) -> CellValue {
    if variant_is_empty(variant) {
        CellValue::Null
    } else if let Some(b) = variant_get_bool(variant) {
        CellValue::Bool(b)
    } else if let Some(n) = variant_get_f64(variant) {
        CellValue::Number(n)
    } else if let Some(s) = variant_get_string(variant) {
        CellValue::String(s)
    } else if variant_is_error(variant) {
        CellValue::Error(CellError {
            code: format!("#ERR(VT_ERROR)"),
        })
    } else {
        CellValue::Null
    }
}
