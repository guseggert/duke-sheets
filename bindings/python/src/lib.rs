//! Python bindings for duke-sheets
//!
//! This module provides PyO3-based Python bindings for the duke-sheets library,
//! allowing Python code to read, write, and manipulate Excel files.

use pyo3::exceptions::{PyIOError, PyIndexError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;

use std::path::PathBuf;
use std::sync::{Arc, RwLock};

use duke_sheets::prelude::*;
use duke_sheets_core::{CellError, CellValue as CoreCellValue};

// =============================================================================
// Error Conversion
// =============================================================================

fn to_py_err(e: impl std::fmt::Display) -> PyErr {
    PyRuntimeError::new_err(e.to_string())
}

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

// =============================================================================
// CellValue - Python wrapper for cell values
// =============================================================================

/// Represents a cell value in a spreadsheet.
///
/// Cell values can be one of several types:
/// - Empty (None)
/// - Number (float)
/// - Text (str)
/// - Boolean (bool)
/// - Error (str like "#DIV/0!")
/// - Formula (has formula text and calculated result)
#[pyclass(name = "CellValue")]
#[derive(Clone)]
pub struct PyCellValue {
    inner: CoreCellValue,
}

#[pymethods]
impl PyCellValue {
    /// Check if the cell is empty
    #[getter]
    fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    /// Check if the cell contains a number
    #[getter]
    fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    /// Check if the cell contains text
    #[getter]
    fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    /// Check if the cell contains a boolean
    #[getter]
    fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    /// Check if the cell contains an error
    #[getter]
    fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    /// Check if the cell contains a formula
    #[getter]
    fn is_formula(&self) -> bool {
        matches!(self.inner, CoreCellValue::Formula { .. })
    }

    /// Get the value as a number, or None if not a number
    fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Get the value as text, or None if not text
    fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    /// Get the value as a boolean, or None if not a boolean
    fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    /// Get the error string, or None if not an error
    fn as_error(&self) -> Option<&'static str> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e)),
            _ => None,
        }
    }

    /// Get the formula text, or None if not a formula
    fn formula_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Formula { text, .. } => Some(text.clone()),
            _ => None,
        }
    }

    /// Convert to a Python object (None, float, str, bool)
    fn to_python(&self, py: Python<'_>) -> PyObject {
        match &self.inner {
            CoreCellValue::Empty => py.None(),
            CoreCellValue::Number(n) => n.into_py(py),
            CoreCellValue::String(s) => s.to_string().into_py(py),
            CoreCellValue::Boolean(b) => b.into_py(py),
            CoreCellValue::Error(e) => cell_error_to_string(e).into_py(py),
            CoreCellValue::Formula { text, .. } => text.into_py(py),
            CoreCellValue::SpillTarget { .. } => py.None(), // Spill targets appear empty
        }
    }

    fn __repr__(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => "CellValue(Empty)".to_string(),
            CoreCellValue::Number(n) => format!("CellValue(Number({}))", n),
            CoreCellValue::String(s) => format!("CellValue(Text({:?}))", s.to_string()),
            CoreCellValue::Boolean(b) => format!("CellValue(Boolean({}))", b),
            CoreCellValue::Error(e) => format!("CellValue(Error({}))", cell_error_to_string(e)),
            CoreCellValue::Formula { text, .. } => format!("CellValue(Formula({:?}))", text),
            CoreCellValue::SpillTarget { .. } => "CellValue(SpillTarget)".to_string(),
        }
    }

    fn __str__(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => "".to_string(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            CoreCellValue::Formula { text, .. } => text.clone(),
            CoreCellValue::SpillTarget { .. } => "".to_string(),
        }
    }
}

// =============================================================================
// CalculationStats - Statistics from workbook calculation
// =============================================================================

/// Statistics from calculating a workbook.
#[pyclass(name = "CalculationStats")]
#[derive(Clone)]
pub struct PyCalculationStats {
    /// Number of formulas found
    #[pyo3(get)]
    pub formula_count: usize,
    /// Number of cells calculated
    #[pyo3(get)]
    pub cells_calculated: usize,
    /// Number of errors encountered
    #[pyo3(get)]
    pub errors: usize,
    /// Number of circular references detected
    #[pyo3(get)]
    pub circular_references: usize,
    /// Number of volatile cells (e.g., NOW(), RAND())
    #[pyo3(get)]
    pub volatile_cells: usize,
    /// Whether iterative calculation converged
    #[pyo3(get)]
    pub converged: bool,
    /// Number of iterations performed
    #[pyo3(get)]
    pub iterations: usize,
}

#[pymethods]
impl PyCalculationStats {
    fn __repr__(&self) -> String {
        format!(
            "CalculationStats(formulas={}, calculated={}, errors={}, circular={}, converged={})",
            self.formula_count,
            self.cells_calculated,
            self.errors,
            self.circular_references,
            self.converged
        )
    }
}

impl From<CalculationStats> for PyCalculationStats {
    fn from(stats: CalculationStats) -> Self {
        Self {
            formula_count: stats.formula_count,
            cells_calculated: stats.cells_calculated,
            errors: stats.errors,
            circular_references: stats.circular_references,
            volatile_cells: stats.volatile_cells,
            converged: stats.converged,
            iterations: stats.iterations as usize,
        }
    }
}

// =============================================================================
// Worksheet - Python wrapper
// =============================================================================

/// A worksheet within a workbook.
///
/// Worksheets contain cells organized in rows and columns. Each cell can
/// contain a value (number, text, boolean) or a formula.
#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    workbook: Arc<RwLock<Workbook>>,
    sheet_index: usize,
}

#[pymethods]
impl PyWorksheet {
    /// Get the worksheet name
    #[getter]
    fn name(&self) -> PyResult<String> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))
    }

    /// Set a cell value by address (e.g., "A1", "B2")
    ///
    /// The value can be:
    /// - None (clears the cell)
    /// - int or float (number)
    /// - str (text)
    /// - bool (boolean)
    #[pyo3(signature = (address, value))]
    fn set_cell(&self, address: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let cell_value = python_to_cell_value(value)?;

        // Parse address
        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_py_err)
    }

    /// Set a formula in a cell
    ///
    /// Args:
    ///     address: Cell address (e.g., "A1")
    ///     formula: Formula string (e.g., "=SUM(A1:A10)")
    #[pyo3(signature = (address, formula))]
    fn set_formula(&self, address: &str, formula: &str) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_cell_formula(address, formula).map_err(to_py_err)
    }

    /// Get the raw cell value (not calculated)
    #[pyo3(signature = (address))]
    fn get_cell(&self, address: &str) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        let value = ws.get_value_at(addr.row, addr.col);

        Ok(PyCellValue { inner: value })
    }

    /// Get the calculated value of a cell
    ///
    /// For formulas, this returns the computed result.
    /// For regular values, returns the value itself.
    #[pyo3(signature = (address))]
    fn get_calculated_value(&self, address: &str) -> PyResult<PyCellValue> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| PyValueError::new_err(format!("Invalid cell address: {}", e)))?;

        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);

        Ok(PyCellValue { inner: value })
    }

    /// Get the used range as (min_row, min_col, max_row, max_col)
    ///
    /// Returns None if the worksheet is empty.
    #[getter]
    fn used_range(&self) -> PyResult<Option<(u32, u16, u32, u16)>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws
            .used_range()
            .map(|r| (r.start.row, r.start.col, r.end.row, r.end.col)))
    }

    /// Set the height of a row in points
    #[pyo3(signature = (row, height))]
    fn set_row_height(&self, row: u32, height: f64) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_row_height(row, height);
        Ok(())
    }

    /// Set the width of a column in character units
    #[pyo3(signature = (col, width))]
    fn set_column_width(&self, col: u16, width: f64) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        ws.set_column_width(col, width);
        Ok(())
    }

    /// Merge cells in a range
    ///
    /// Args:
    ///     range_str: Range to merge (e.g., "A1:C3")
    #[pyo3(signature = (range_str))]
    fn merge_cells(&self, range_str: &str) -> PyResult<()> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let range = duke_sheets_core::CellRange::parse(range_str)
            .map_err(|e| PyValueError::new_err(format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_py_err)
    }

    /// Unmerge cells in a range
    #[pyo3(signature = (range_str))]
    fn unmerge_cells(&self, range_str: &str) -> PyResult<bool> {
        let mut wb = self.workbook.write().map_err(to_py_err)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        let range = duke_sheets_core::CellRange::parse(range_str)
            .map_err(|e| PyValueError::new_err(format!("Invalid range: {}", e)))?;
        Ok(ws.unmerge_cells(&range))
    }

    /// Get the row height in points
    #[pyo3(signature = (row))]
    fn get_row_height(&self, row: u32) -> PyResult<f64> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws.row_height(row))
    }

    /// Get the column width in character units
    #[pyo3(signature = (col))]
    fn get_column_width(&self, col: u16) -> PyResult<f64> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

        Ok(ws.column_width(col))
    }

    fn __repr__(&self) -> PyResult<String> {
        let name = self.name()?;
        Ok(format!("Worksheet({:?})", name))
    }
}

// =============================================================================
// Workbook - Python wrapper
// =============================================================================

/// A workbook containing one or more worksheets.
///
/// This is the main entry point for working with spreadsheet files.
///
/// Example:
///     >>> wb = Workbook()
///     >>> sheet = wb.get_sheet(0)
///     >>> sheet.set_cell("A1", 10)
///     >>> sheet.set_cell("A2", 20)
///     >>> sheet.set_formula("A3", "=A1+A2")
///     >>> wb.calculate()
///     >>> sheet.get_calculated_value("A3").as_number()
///     30.0
#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    inner: Arc<RwLock<Workbook>>,
}

#[pymethods]
impl PyWorkbook {
    /// Create a new empty workbook with one worksheet
    #[new]
    fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(Workbook::new())),
        }
    }

    /// Open a workbook from a file
    ///
    /// Supported formats:
    /// - .xlsx (Excel 2007+)
    /// - .csv (Comma-separated values)
    ///
    /// Args:
    ///     path: Path to the file
    ///
    /// Returns:
    ///     Workbook instance
    #[staticmethod]
    #[pyo3(signature = (path))]
    fn open(path: &str) -> PyResult<Self> {
        use duke_sheets::WorkbookExt;
        let path = PathBuf::from(path);

        let wb = Workbook::open(&path).map_err(|e| PyIOError::new_err(e.to_string()))?;

        Ok(Self {
            inner: Arc::new(RwLock::new(wb)),
        })
    }

    /// Save the workbook to a file
    ///
    /// The format is determined by the file extension:
    /// - .xlsx for Excel format
    /// - .csv for CSV format (first sheet only)
    ///
    /// Args:
    ///     path: Path to save to
    #[pyo3(signature = (path))]
    fn save(&self, path: &str) -> PyResult<()> {
        use duke_sheets::WorkbookExt;
        let wb = self.inner.read().map_err(to_py_err)?;
        let path = PathBuf::from(path);

        wb.save(&path)
            .map_err(|e| PyIOError::new_err(e.to_string()))
    }

    /// Get the number of worksheets
    #[getter]
    fn sheet_count(&self) -> PyResult<usize> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.sheet_count())
    }

    /// Get a list of all worksheet names
    #[getter]
    fn sheet_names(&self) -> PyResult<Vec<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    /// Get a worksheet by index or name
    ///
    /// Args:
    ///     index_or_name: Either an integer index or a string name
    ///
    /// Returns:
    ///     Worksheet instance
    ///
    /// Raises:
    ///     IndexError: If the index is out of range or name not found
    #[pyo3(signature = (index_or_name))]
    fn get_sheet(&self, index_or_name: &Bound<'_, PyAny>) -> PyResult<PyWorksheet> {
        let wb = self.inner.read().map_err(to_py_err)?;

        let sheet_index = if let Ok(idx) = index_or_name.extract::<usize>() {
            if idx >= wb.sheet_count() {
                return Err(PyIndexError::new_err(format!(
                    "Sheet index {} out of range (0..{})",
                    idx,
                    wb.sheet_count()
                )));
            }
            idx
        } else if let Ok(name) = index_or_name.extract::<String>() {
            wb.sheet_index(&name)
                .ok_or_else(|| PyIndexError::new_err(format!("Sheet '{}' not found", name)))?
        } else {
            return Err(PyValueError::new_err(
                "Expected int or str for sheet index/name",
            ));
        };

        drop(wb); // Release the lock

        Ok(PyWorksheet {
            workbook: Arc::clone(&self.inner),
            sheet_index,
        })
    }

    /// Add a new worksheet with the given name
    ///
    /// Args:
    ///     name: Name for the new worksheet
    ///
    /// Returns:
    ///     Index of the new worksheet
    #[pyo3(signature = (name))]
    fn add_sheet(&self, name: &str) -> PyResult<usize> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.add_worksheet_with_name(name).map_err(to_py_err)
    }

    /// Remove a worksheet by index
    ///
    /// Args:
    ///     index: Index of the worksheet to remove
    #[pyo3(signature = (index))]
    fn remove_sheet(&self, index: usize) -> PyResult<()> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.remove_worksheet(index).map(|_| ()).map_err(to_py_err)
    }

    /// Calculate all formulas in the workbook
    ///
    /// Returns:
    ///     CalculationStats with information about the calculation
    fn calculate(&self) -> PyResult<PyCalculationStats> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        let stats = wb.calculate().map_err(to_py_err)?;
        Ok(stats.into())
    }

    /// Calculate with custom options for iterative calculation
    ///
    /// Args:
    ///     iterative: Enable iterative calculation for circular references
    ///     max_iterations: Maximum number of iterations (default 100)
    ///     max_change: Convergence threshold (default 0.001)
    ///
    /// Returns:
    ///     CalculationStats with information about the calculation
    #[pyo3(signature = (iterative = false, max_iterations = 100, max_change = 0.001))]
    fn calculate_with_options(
        &self,
        iterative: bool,
        max_iterations: u32,
        max_change: f64,
    ) -> PyResult<PyCalculationStats> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        let options = CalculationOptions {
            iterative,
            max_iterations,
            max_change,
            ..Default::default()
        };
        let stats = wb.calculate_with_options(&options).map_err(to_py_err)?;
        Ok(stats.into())
    }

    /// Define a named range
    ///
    /// Args:
    ///     name: Name for the range (e.g., "TaxRate")
    ///     refers_to: What the name refers to (e.g., "Sheet1!$A$1" or "0.05")
    #[pyo3(signature = (name, refers_to))]
    fn define_name(&self, name: &str, refers_to: &str) -> PyResult<()> {
        let mut wb = self.inner.write().map_err(to_py_err)?;
        wb.define_name(name, refers_to).map_err(to_py_err)
    }

    /// Get a named range definition
    ///
    /// Args:
    ///     name: Name to look up
    ///
    /// Returns:
    ///     The refers_to string, or None if not found
    #[pyo3(signature = (name))]
    fn get_named_range(&self, name: &str) -> PyResult<Option<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.get_named_range(name, 0).map(|nr| nr.refers_to.clone()))
    }

    fn __repr__(&self) -> PyResult<String> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(format!("Workbook(sheets={})", wb.sheet_count()))
    }
}

// =============================================================================
// Helper functions
// =============================================================================

/// Convert a Python value to a CellValue
fn python_to_cell_value(value: &Bound<'_, PyAny>) -> PyResult<CoreCellValue> {
    if value.is_none() {
        Ok(CoreCellValue::Empty)
    } else if let Ok(b) = value.extract::<bool>() {
        Ok(CoreCellValue::Boolean(b))
    } else if let Ok(n) = value.extract::<f64>() {
        Ok(CoreCellValue::Number(n))
    } else if let Ok(n) = value.extract::<i64>() {
        Ok(CoreCellValue::Number(n as f64))
    } else if let Ok(s) = value.extract::<String>() {
        Ok(CoreCellValue::string(s))
    } else {
        Err(PyValueError::new_err(
            "Cell value must be None, bool, int, float, or str",
        ))
    }
}

// =============================================================================
// Module definition
// =============================================================================

/// duke_sheets - High-performance Excel file library for Python
///
/// This module provides fast, memory-efficient access to Excel files (.xlsx)
/// and CSV files, with full formula calculation support.
///
/// Example:
///     >>> import duke_sheets
///     >>> wb = duke_sheets.Workbook()
///     >>> sheet = wb.get_sheet(0)
///     >>> sheet.set_cell("A1", 10)
///     >>> sheet.set_formula("A2", "=A1*2")
///     >>> wb.calculate()
///     >>> print(sheet.get_calculated_value("A2").as_number())
///     20.0
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCellValue>()?;
    m.add_class::<PyCalculationStats>()?;
    Ok(())
}
