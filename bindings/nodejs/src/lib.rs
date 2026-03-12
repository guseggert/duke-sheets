//! Node.js/TypeScript bindings for duke-sheets
//!
//! This module provides NAPI-RS-based native Node.js bindings for the duke-sheets
//! library, allowing JavaScript/TypeScript code to read, write, and manipulate
//! Excel files with native performance.

use napi::bindgen_prelude::*;
use napi_derive::napi;

use std::io::Cursor;
use std::path::PathBuf;
use std::sync::{Arc, RwLock};

use duke_sheets::{
    CalculationOptions, CalculationStats as CoreCalculationStats, WorkbookCalculationExt,
    WorkbookExt,
};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue as CoreCellValue, Workbook as CoreWorkbook,
};

fn to_napi_err(e: impl std::fmt::Display) -> napi::Error {
    napi::Error::from_reason(e.to_string())
}

/// Catch Rust panics at the FFI boundary and convert them to napi::Error.
///
/// Without this, a panic (e.g. integer overflow, index out of bounds) would
/// unwind across the FFI boundary into Node.js, which is undefined behavior
/// and kills the process. With this wrapper, panics become JS exceptions that
/// callers can catch normally.
pub(crate) fn catch_panic<T>(f: impl FnOnce() -> napi::Result<T>) -> napi::Result<T> {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(f)) {
        Ok(result) => result,
        Err(payload) => {
            let msg = if let Some(s) = payload.downcast_ref::<&str>() {
                s.to_string()
            } else if let Some(s) = payload.downcast_ref::<String>() {
                s.clone()
            } else {
                "unknown error".to_string()
            };
            Err(napi::Error::from_reason(format!("Internal error: {}", msg)))
        }
    }
}


mod types;
pub use types::*;
mod workbook_read;
mod worksheet_read;

fn cell_error_to_string(e: &CellError) -> &'static str {
    match e {
        CellError::Div0 => "#DIV/0!",
        CellError::Na => "#N/A",
        CellError::Name => "#NAME?",
        CellError::Null => "#NULL!",
        CellError::Num => "#NUM!",
        CellError::Ref => "#REF!",
        CellError::Value => "#VALUE!",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
    }
}

/// Represents a cell value in a spreadsheet.
///
/// Cell values can be one of several types:
/// - Empty (null/undefined)
/// - Number
/// - Text (string)
/// - Boolean
/// - Error (like "#DIV/0!")
/// - Formula cached results are exposed as regular cell values; formula text lives on Worksheet accessors
#[napi]
pub struct CellValue {
    inner: CoreCellValue,
}

#[napi]
impl CellValue {
    /// Check if the cell is empty
    #[napi(getter)]
    pub fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    /// Check if the cell contains a number
    #[napi(getter)]
    pub fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    /// Check if the cell contains text
    #[napi(getter)]
    pub fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    /// Check if the cell contains a boolean
    #[napi(getter)]
    pub fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    /// Check if the cell contains an error
    #[napi(getter)]
    pub fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    /// Get the value as a number, or null if not a number
    #[napi]
    pub fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Get the value as text, or null if not text
    #[napi]
    pub fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    /// Get the value as a boolean, or null if not a boolean
    #[napi]
    pub fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the error string, or null if not an error
    #[napi]
    pub fn as_error(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e).to_string()),
            _ => None,
        }
    }

    /// Convert to a JavaScript native value (number, string, boolean, or null)
    #[napi(js_name = "toJs")]
    pub fn to_js(&self) -> Either4<f64, String, bool, Null> {
        match &self.inner {
            CoreCellValue::Empty => Either4::D(Null),
            CoreCellValue::Number(n) => Either4::A(*n),
            CoreCellValue::String(s) => Either4::B(s.to_string()),
            CoreCellValue::Boolean(b) => Either4::C(*b),
            CoreCellValue::Error(e) => Either4::B(cell_error_to_string(e).to_string()),
            _ => Either4::D(Null),
        }
    }

    /// Get string representation of the cell value
    #[napi(js_name = "toString")]
    pub fn to_string_js(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => String::new(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            _ => String::new(),
        }
    }
}

/// Statistics from calculating a workbook.
#[napi]
pub struct CalculationStats {
    inner: CoreCalculationStats,
}

#[napi]
impl CalculationStats {
    /// Number of formulas found
    #[napi(getter)]
    pub fn formula_count(&self) -> u32 {
        self.inner.formula_count as u32
    }

    /// Number of cells calculated
    #[napi(getter)]
    pub fn cells_calculated(&self) -> u32 {
        self.inner.cells_calculated as u32
    }

    /// Number of errors encountered
    #[napi(getter)]
    pub fn errors(&self) -> u32 {
        self.inner.errors as u32
    }

    /// Number of circular references detected
    #[napi(getter)]
    pub fn circular_references(&self) -> u32 {
        self.inner.circular_references as u32
    }

    /// Number of volatile cells (e.g., NOW(), RAND())
    #[napi(getter)]
    pub fn volatile_cells(&self) -> u32 {
        self.inner.volatile_cells as u32
    }

    /// Whether iterative calculation converged
    #[napi(getter)]
    pub fn converged(&self) -> bool {
        self.inner.converged
    }

    /// Number of iterations performed
    #[napi(getter)]
    pub fn iterations(&self) -> u32 {
        self.inner.iterations as u32
    }
}

/// The used range of a worksheet, describing the bounding box of all cells
/// that contain data.
#[napi(object)]
pub struct UsedRange {
    pub min_row: u32,
    pub min_col: u32,
    pub max_row: u32,
    pub max_col: u32,
}

/// A worksheet within a workbook.
///
/// Worksheets contain cells organized in rows and columns. Each cell can
/// contain a value (number, text, boolean) or a formula.
#[napi]
pub struct Worksheet {
    workbook: Arc<RwLock<CoreWorkbook>>,
    sheet_index: usize,
}

#[napi]
impl Worksheet {
    /// Get the worksheet name
    #[napi(getter)]
    pub fn name(&self) -> Result<String> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            wb.worksheet(self.sheet_index)
                .map(|ws| ws.name().to_string())
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))
        })
    }

    /// Set a cell value by address (e.g., "A1", "B2")
    ///
    /// The value can be:
    /// - `null` or `undefined` (clears the cell)
    /// - `number` (numeric value)
    /// - `string` (text)
    /// - `boolean`
    #[napi]
    pub fn set_cell(
        &self,
        address: String,
        value: Option<Either3<f64, String, bool>>,
    ) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let cell_value = if let Some(value) = value {
                match value {
                    Either3::A(n) => CoreCellValue::Number(n),
                    Either3::B(s) => CoreCellValue::string(s),
                    Either3::C(b) => CoreCellValue::Boolean(b),
                }
            } else {
                CoreCellValue::Empty
            };

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            ws.set_cell_value_at(addr.row, addr.col, cell_value)
                .map_err(to_napi_err)
        })
    }

    /// Set a formula in a cell
    ///
    /// @param address - Cell address (e.g., "A1")
    /// @param formula - Formula string (e.g., "=SUM(A1:A10)")
    #[napi]
    pub fn set_formula(&self, address: String, formula: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_cell_formula(&address, &formula).map_err(to_napi_err)
        })
    }

    /// Get the raw cell value (not calculated)
    #[napi]
    pub fn get_cell(&self, address: String) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            let value = ws.get_value_at(addr.row, addr.col);
            Ok(CellValue { inner: value })
        })
    }

    /// Get the raw cell value by row/col (0-based).
    #[napi]
    pub fn get_cell_at(&self, row: u32, col: u32) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let value = ws.get_value_at(row, col as u16);
            Ok(CellValue { inner: value })
        })
    }

    /// Get the calculated value of a cell
    ///
    /// For formulas, this returns the computed result.
    /// For regular values, returns the value itself.
    #[napi]
    pub fn get_calculated_value(&self, address: String) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let addr = CellAddress::parse(&address)
                .map_err(|e| napi::Error::from_reason(format!("Invalid cell address: {}", e)))?;

            let value = ws
                .get_calculated_value_at(addr.row, addr.col)
                .cloned()
                .unwrap_or(CoreCellValue::Empty);

            Ok(CellValue { inner: value })
        })
    }

    /// Get the calculated value of a cell by row/col (0-based).
    #[napi]
    pub fn get_calculated_value_at(&self, row: u32, col: u32) -> Result<CellValue> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let value = ws
                .get_calculated_value_at(row, col as u16)
                .cloned()
                .unwrap_or(CoreCellValue::Empty);

            Ok(CellValue { inner: value })
        })
    }

    /// Get the used range as `{ minRow, minCol, maxRow, maxCol }` or null
    /// if the worksheet is empty.
    #[napi(getter)]
    pub fn used_range(&self) -> Result<Option<UsedRange>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.used_range().map(|r| UsedRange {
                min_row: r.start.row,
                min_col: r.start.col as u32,
                max_row: r.end.row,
                max_col: r.end.col as u32,
            }))
        })
    }

    /// Set the height of a row in points
    #[napi]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_row_height(row, height);
            Ok(())
        })
    }

    /// Set the width of a column in character units
    #[napi]
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            ws.set_column_width(col as u16, width);
            Ok(())
        })
    }

    /// Get the row height in points, or null if not explicitly set
    #[napi]
    pub fn get_row_height(&self, row: u32) -> Result<Option<f64>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.custom_row_heights().get(&row).copied())
        })
    }

    /// Get the column width in character units, or null if not explicitly set
    #[napi]
    pub fn get_column_width(&self, col: u32) -> Result<Option<f64>> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            Ok(ws.custom_column_widths().get(&(col as u16)).copied())
        })
    }

    /// Merge cells in a range (e.g., "A1:C3")
    #[napi]
    pub fn merge_cells(&self, range_str: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let range = CellRange::parse(&range_str)
                .map_err(|e| napi::Error::from_reason(format!("Invalid range: {}", e)))?;
            ws.merge_cells(&range).map_err(to_napi_err)
        })
    }

    /// Unmerge cells in a range
    #[napi]
    pub fn unmerge_cells(&self, range_str: String) -> Result<bool> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            let ws = wb
                .worksheet_mut(self.sheet_index)
                .ok_or_else(|| napi::Error::from_reason("Worksheet no longer exists"))?;

            let range = CellRange::parse(&range_str)
                .map_err(|e| napi::Error::from_reason(format!("Invalid range: {}", e)))?;
            Ok(ws.unmerge_cells(&range))
        })
    }
}

/// A workbook containing one or more worksheets.
///
/// This is the main entry point for working with spreadsheet files.
///
/// @example
/// ```typescript
/// const wb = new Workbook();
/// const sheet = wb.getSheet(0);
/// sheet.setCell("A1", 10);
/// sheet.setCell("A2", 20);
/// sheet.setFormula("A3", "=A1+A2");
/// wb.calculate();
/// console.log(sheet.getCalculatedValue("A3").asNumber()); // 30
/// ```
#[napi]
pub struct Workbook {
    inner: Arc<RwLock<CoreWorkbook>>,
}

#[napi]
impl Workbook {
    /// Create a new empty workbook with one worksheet
    #[napi(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(CoreWorkbook::new())),
        }
    }

    /// Open a workbook from a file
    ///
    /// Supported formats:
    /// - `.xlsx` (Excel 2007+)
    /// - `.xls` (Legacy Excel)
    /// - `.csv` (Comma-separated values)
    ///
    /// @param path - Path to the file
    #[napi(factory)]
    pub fn open(path: String) -> Result<Self> {
        catch_panic(|| {
            let path = PathBuf::from(path);
            let wb = CoreWorkbook::open(&path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to open file: {}", e)))?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Load a workbook from bytes (Buffer/Uint8Array), auto-detecting the format.
    ///
    /// Supports XLSX and XLS formats. The format is detected from magic bytes.
    ///
    /// @param data - The file content as a Buffer
    #[napi(factory)]
    pub fn from_bytes(data: Buffer) -> Result<Self> {
        catch_panic(|| {
            use duke_sheets::WorkbookExt;
            let wb = duke_sheets_core::Workbook::from_bytes(data.as_ref())
                .map_err(|e| napi::Error::from_reason(format!("Failed to read file: {}", e)))?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }

    /// Load a workbook from a CSV string
    ///
    /// @param csv - The CSV content as a string
    #[napi(factory)]
    pub fn from_csv_string(csv: String) -> Result<Self> {
        catch_panic(|| {
            let reader = Cursor::new(csv.into_bytes());
            let ws = duke_sheets_csv::CsvReader::read(
                reader,
                &duke_sheets_csv::CsvReadOptions::default(),
            )
            .map_err(|e| napi::Error::from_reason(format!("Failed to read CSV: {}", e)))?;

            let mut wb = CoreWorkbook::empty();
            wb.add_existing_worksheet(ws).map_err(to_napi_err)?;

            Ok(Self {
                inner: Arc::new(RwLock::new(wb)),
            })
        })
    }


    /// Save the workbook to a file
    ///
    /// The format is determined by the file extension:
    /// - `.xlsx` for Excel format
    /// - `.csv` for CSV format (first sheet only)
    ///
    /// @param path - Path to save to
    #[napi]
    pub fn save(&self, path: String) -> Result<()> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let path = PathBuf::from(path);
            wb.save(&path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to save: {}", e)))
        })
    }

    /// Save the workbook as a CSV string (first sheet only)
    #[napi]
    pub fn save_csv_string(&self) -> Result<String> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let ws = wb
                .worksheet(0)
                .ok_or_else(|| napi::Error::from_reason("No worksheets to save"))?;

            let mut buf = Vec::new();
            duke_sheets_csv::CsvWriter::write(
                ws,
                &mut buf,
                &duke_sheets_csv::CsvWriteOptions::default(),
            )
            .map_err(|e| napi::Error::from_reason(format!("Failed to write CSV: {}", e)))?;

            String::from_utf8(buf)
                .map_err(|e| napi::Error::from_reason(format!("Invalid UTF-8: {}", e)))
        })
    }

    /// Get the number of worksheets
    #[napi(getter)]
    pub fn sheet_count(&self) -> Result<u32> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            let count = u32::try_from(wb.sheet_count()).map_err(to_napi_err)?;
            Ok(count)
        })
    }

    /// Get a list of all worksheet names
    #[napi(getter)]
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok((0..wb.sheet_count())
                .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
                .collect())
        })
    }

    /// Get a worksheet by index (number) or name (string)
    ///
    /// @param indexOrName - Either a zero-based index or a sheet name
    /// @throws Error if index out of range or name not found
    #[napi]
    pub fn get_sheet(&self, index_or_name: Either<u32, String>) -> Result<Worksheet> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;

            let sheet_index = match index_or_name {
                Either::A(idx) => {
                    let idx = idx as usize;
                    if idx >= wb.sheet_count() {
                        return Err(napi::Error::from_reason(format!(
                            "Sheet index {} out of range (0..{})",
                            idx,
                            wb.sheet_count()
                        )));
                    }
                    idx
                }
                Either::B(name) => wb.sheet_index(&name).ok_or_else(|| {
                    napi::Error::from_reason(format!("Sheet '{}' not found", name))
                })?,
            };

            drop(wb);

            Ok(Worksheet {
                workbook: Arc::clone(&self.inner),
                sheet_index,
            })
        })
    }

    /// Add a new worksheet with the given name
    ///
    /// @param name - Name for the new worksheet
    /// @returns Index of the new worksheet
    #[napi]
    pub fn add_sheet(&self, name: String) -> Result<u32> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.add_worksheet_with_name(&name)
                .map(|idx| idx as u32)
                .map_err(to_napi_err)
        })
    }

    /// Remove a worksheet by index
    ///
    /// @param index - Zero-based index of the worksheet to remove
    #[napi]
    pub fn remove_sheet(&self, index: u32) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.remove_worksheet(index as usize)
                .map(|_| ())
                .map_err(to_napi_err)
        })
    }

    /// Calculate all formulas in the workbook
    ///
    /// @returns Statistics about the calculation
    #[napi]
    pub fn calculate(&self) -> Result<CalculationStats> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            let stats = wb.calculate().map_err(to_napi_err)?;
            Ok(CalculationStats { inner: stats })
        })
    }

    /// Calculate with custom options for iterative calculation
    ///
    /// @param iterative - Enable iterative calculation for circular references
    /// @param maxIterations - Maximum iterations (default 100)
    /// @param maxChange - Convergence threshold (default 0.001)
    #[napi]
    pub fn calculate_with_options(
        &self,
        iterative: Option<bool>,
        max_iterations: Option<u32>,
        max_change: Option<f64>,
    ) -> Result<CalculationStats> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            let options = CalculationOptions {
                iterative: iterative.unwrap_or(false),
                max_iterations: max_iterations.unwrap_or(100),
                max_change: max_change.unwrap_or(0.001),
                ..Default::default()
            };
            let stats = wb.calculate_with_options(&options).map_err(to_napi_err)?;
            Ok(CalculationStats { inner: stats })
        })
    }

    /// Define a named range
    ///
    /// @param name - Name for the range (e.g., "TaxRate")
    /// @param refersTo - What the name refers to (e.g., "Sheet1!$A$1" or "0.05")
    #[napi]
    pub fn define_name(&self, name: String, refers_to: String) -> Result<()> {
        catch_panic(|| {
            let mut wb = self.inner.write().map_err(to_napi_err)?;
            wb.define_name(&name, &refers_to).map_err(to_napi_err)
        })
    }

    /// Get a named range definition
    ///
    /// @param name - Name to look up
    /// @returns The refers_to string, or null if not found
    #[napi]
    pub fn get_named_range(&self, name: String) -> Result<Option<String>> {
        catch_panic(|| {
            let wb = self.inner.read().map_err(to_napi_err)?;
            Ok(wb.get_named_range(&name, 0).map(|nr| nr.refers_to.clone()))
        })
    }
}

pub struct OpenTask {
    path: PathBuf,
}

impl Task for OpenTask {
    type Output = CoreWorkbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            CoreWorkbook::open(&self.path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to open file: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook {
            inner: Arc::new(RwLock::new(output)),
        })
    }
}

pub struct OpenBytesTask {
    data: Vec<u8>,
}

impl Task for OpenBytesTask {
    type Output = CoreWorkbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            use duke_sheets::WorkbookExt;
            CoreWorkbook::from_bytes(&self.data)
                .map_err(|e| napi::Error::from_reason(format!("Failed to read file: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook {
            inner: Arc::new(RwLock::new(output)),
        })
    }
}

pub struct SaveTask {
    workbook: Arc<RwLock<CoreWorkbook>>,
    path: PathBuf,
}

impl Task for SaveTask {
    type Output = ();
    type JsValue = ();

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            let wb = self.workbook.read().map_err(to_napi_err)?;
            wb.save(&self.path)
                .map_err(|e| napi::Error::from_reason(format!("Failed to save: {}", e)))
        })
    }

    fn resolve(&mut self, _env: Env, _output: Self::Output) -> Result<Self::JsValue> {
        Ok(())
    }
}

pub struct CalculateTask {
    workbook: Arc<RwLock<CoreWorkbook>>,
    options: Option<CalculationOptions>,
}

impl Task for CalculateTask {
    type Output = CoreCalculationStats;
    type JsValue = CalculationStats;

    fn compute(&mut self) -> Result<Self::Output> {
        catch_panic(|| {
            let mut wb = self.workbook.write().map_err(to_napi_err)?;
            if let Some(opts) = &self.options {
                wb.calculate_with_options(opts).map_err(to_napi_err)
            } else {
                wb.calculate().map_err(to_napi_err)
            }
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(CalculationStats { inner: output })
    }
}

/// Open a workbook from a file asynchronously (non-blocking).
///
/// Runs file I/O and parsing on the libuv thread pool so the
/// Node.js event loop is not blocked.
///
/// @param path - Path to the file
/// @returns Promise<Workbook>
#[napi]
pub fn open_async(path: String) -> AsyncTask<OpenTask> {
    AsyncTask::new(OpenTask {
        path: PathBuf::from(path),
    })
}

/// Load a workbook from bytes asynchronously (non-blocking).
///
/// Auto-detects format (XLSX or XLS) from magic bytes.
///
/// @param data - The file content as a Buffer
/// @returns Promise<Workbook>
#[napi]
pub fn from_bytes_async(data: Buffer) -> AsyncTask<OpenBytesTask> {
    AsyncTask::new(OpenBytesTask {
        data: data.to_vec(),
    })
}

#[napi]
impl Workbook {
    /// Save the workbook to a file asynchronously (non-blocking).
    ///
    /// @param path - Path to save to
    /// @returns Promise<void>
    #[napi]
    pub fn save_async(&self, path: String) -> AsyncTask<SaveTask> {
        AsyncTask::new(SaveTask {
            workbook: Arc::clone(&self.inner),
            path: PathBuf::from(path),
        })
    }

    /// Calculate all formulas asynchronously (non-blocking).
    ///
    /// @returns Promise<CalculationStats>
    #[napi]
    pub fn calculate_async(&self) -> AsyncTask<CalculateTask> {
        AsyncTask::new(CalculateTask {
            workbook: Arc::clone(&self.inner),
            options: None,
        })
    }

    /// Calculate with custom options asynchronously (non-blocking).
    ///
    /// @param iterative - Enable iterative calculation for circular references
    /// @param maxIterations - Maximum iterations (default 100)
    /// @param maxChange - Convergence threshold (default 0.001)
    /// @returns Promise<CalculationStats>
    #[napi]
    pub fn calculate_with_options_async(
        &self,
        iterative: Option<bool>,
        max_iterations: Option<u32>,
        max_change: Option<f64>,
    ) -> Result<AsyncTask<CalculateTask>> {
        Ok(AsyncTask::new(CalculateTask {
            workbook: Arc::clone(&self.inner),
            options: Some(CalculationOptions {
                iterative: iterative.unwrap_or(false),
                max_iterations: max_iterations.unwrap_or(100),
                max_change: max_change.unwrap_or(0.001),
                ..Default::default()
            }),
        }))
    }
}
