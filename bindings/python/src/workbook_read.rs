use std::io::Cursor;
use std::sync::{Arc, RwLock};

use duke_sheets_core::{named_range::NameScope, Workbook};
use pyo3::exceptions::{PyIndexError, PyValueError};
use pyo3::prelude::*;

use crate::{to_py_err, PyNamedRange, PyWorkbook, PyWorkbookSettings};

#[pymethods]
impl PyWorkbook {

    #[staticmethod]
    fn from_csv_string(csv: &str) -> PyResult<Self> {
        let reader = Cursor::new(csv.as_bytes());
        let ws =
            duke_sheets_csv::CsvReader::read(reader, &duke_sheets_csv::CsvReadOptions::default())
                .map_err(|e| PyValueError::new_err(format!("Failed to read CSV: {}", e)))?;

        let mut wb = Workbook::empty();
        wb.add_existing_worksheet(ws).map_err(to_py_err)?;

        Ok(Self {
            inner: Arc::new(RwLock::new(wb)),
        })
    }

    /// Load a workbook from bytes, auto-detecting the format.
    ///
    /// Supports XLSX and XLS formats. The format is detected from magic bytes.
    #[staticmethod]
    fn from_bytes(data: &[u8]) -> PyResult<Self> {
        use duke_sheets::WorkbookExt;
        let wb = Workbook::from_bytes(data)
            .map_err(|e| PyValueError::new_err(format!("Failed to read file: {}", e)))?;

        Ok(Self {
            inner: Arc::new(RwLock::new(wb)),
        })
    }

    fn save_csv_string(&self) -> PyResult<String> {
        let wb = self.inner.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(0)
            .ok_or_else(|| PyIndexError::new_err("No worksheets to save"))?;

        let mut buf = Vec::new();
        duke_sheets_csv::CsvWriter::write(
            ws,
            &mut buf,
            &duke_sheets_csv::CsvWriteOptions::default(),
        )
        .map_err(|e| PyValueError::new_err(format!("Failed to write CSV: {}", e)))?;

        String::from_utf8(buf).map_err(|e| PyValueError::new_err(format!("Invalid UTF-8: {}", e)))
    }

    #[getter]
    fn is_empty(&self) -> PyResult<bool> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.is_empty())
    }

    #[getter]
    fn active_sheet(&self) -> PyResult<u32> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.active_sheet() as u32)
    }

    fn sheet_index(&self, name: String) -> PyResult<Option<u32>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.sheet_index(&name).map(|i| i as u32))
    }

    #[getter]
    fn settings(&self) -> PyResult<PyWorkbookSettings> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(PyWorkbookSettings::from(wb.settings()))
    }

    #[getter]
    fn named_ranges(&self) -> PyResult<Vec<PyNamedRange>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb
            .named_ranges()
            .iter()
            .map(|nr| PyNamedRange {
                name: nr.name.clone(),
                scope: match &nr.scope {
                    NameScope::Workbook => "workbook".into(),
                    NameScope::Sheet(_) => "sheet".into(),
                },
                sheet_index: match &nr.scope {
                    NameScope::Workbook => None,
                    NameScope::Sheet(idx) => Some(*idx as u32),
                },
                refers_to: nr.refers_to.clone(),
                comment: nr.comment.clone(),
                hidden: nr.hidden,
            })
            .collect())
    }
}
