use std::io::Cursor;
use std::sync::{Arc, RwLock};

use duke_sheets_core::{named_range::NameScope, Workbook};
use pyo3::exceptions::{PyIndexError, PyValueError};
use pyo3::prelude::*;

use crate::{
    to_py_err, PyChartSheet, PyNamedRange, PySheetSlot, PyWorkbook, PyWorkbookProtection,
    PyWorkbookSettings,
};

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

    /// Save the workbook as a CSV string (first sheet only, with
    /// form-control state synchronized into linked cells in the output).
    fn save_csv_string(&self) -> PyResult<String> {
        let wb = self.inner.read().map_err(to_py_err)?;
        let snapshot = wb.synchronized_for_save();
        let wb = snapshot.as_ref().unwrap_or_else(|| &*wb);
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
    fn workbook_protection(&self) -> PyResult<Option<PyWorkbookProtection>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb
            .workbook_protection()
            .map(|protection| PyWorkbookProtection::from(&protection)))
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

    /// The workbook theme's 12 clrScheme colors as ``RRGGBB`` hex, in
    /// theme-index order (background 1, text 1, background 2, text 2,
    /// accent 1-6, hyperlink, followed hyperlink). The Office default
    /// palette when the file carries no theme.
    #[getter]
    fn theme_palette(&self) -> PyResult<Vec<String>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb
            .theme_palette()
            .colors
            .iter()
            .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}"))
            .collect())
    }

    /// Resolve a :class:`Color` to display RGB (``RRGGBB`` hex)
    /// against this workbook's theme palette. ``auto`` has no fixed
    /// RGB and resolves to ``None``.
    fn resolve_color(&self, color: &crate::PyColor) -> PyResult<Option<String>> {
        let core = color.to_core()?;
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb
            .resolve_color(&core)
            .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}")))
    }
}

#[pymethods]
impl PyWorkbook {
    #[getter]
    fn chartsheets(&self) -> PyResult<Vec<PyChartSheet>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.chartsheets().iter().map(PyChartSheet::from).collect())
    }

    #[getter]
    fn chartsheet_count(&self) -> PyResult<u32> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.chartsheet_count() as u32)
    }

    #[getter]
    fn sheet_order(&self) -> PyResult<Vec<PySheetSlot>> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.sheet_order().iter().map(PySheetSlot::from).collect())
    }

    #[getter]
    fn total_sheet_count(&self) -> PyResult<u32> {
        let wb = self.inner.read().map_err(to_py_err)?;
        Ok(wb.total_sheet_count() as u32)
    }
}
