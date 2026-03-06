use pyo3::exceptions::PyIndexError;
use pyo3::prelude::*;

use crate::{
    to_py_err, PyAutoFilter, PyColor, PyComment, PyCommentEntry, PyConditionalFormatRule,
    PyDataValidation, PyFormulaCell, PyFreezePanes, PyHyperlink, PyHyperlinkEntry, PyPageBreak,
    PyPageSetup, PySelection, PySheetProtection, PySpillSource, PySplitPanes, PyStyle, PyTable,
    PyWorksheet,
};

#[pymethods]
impl PyWorksheet {
    #[getter]
    fn visibility(&self) -> PyResult<String> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(match ws.visibility() {
            duke_sheets_core::SheetVisibility::Visible => "visible".to_string(),
            duke_sheets_core::SheetVisibility::Hidden => "hidden".to_string(),
            duke_sheets_core::SheetVisibility::VeryHidden => "veryHidden".to_string(),
        })
    }

    #[getter]
    fn is_selected(&self) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_selected())
    }

    #[getter]
    fn zoom_scale(&self) -> PyResult<Option<u32>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.zoom_scale().map(|z| z as u32))
    }

    #[getter]
    fn tab_color(&self) -> PyResult<Option<PyColor>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.tab_color().map(|c| PyColor::from(&c)))
    }

    #[getter(is_empty)]
    fn ws_is_empty(&self) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_empty())
    }

    #[getter]
    fn cell_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.cell_count() as u32)
    }

    #[getter]
    fn selections(&self) -> PyResult<Vec<PySelection>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.selections().iter().map(PySelection::from).collect())
    }

    fn get_cell_style(&self, address: String) -> PyResult<Option<PyStyle>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let style = ws.cell_style(&address).map_err(to_py_err)?;
        Ok(style.map(PyStyle::from))
    }

    fn get_cell_style_at(&self, row: u32, col: u32) -> PyResult<Option<PyStyle>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.cell_style_at(row, col as u16).map(PyStyle::from))
    }

    fn get_formatted_value(&self, address: String) -> PyResult<String> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        ws.formatted_value(&address).map_err(to_py_err)
    }

    fn get_formatted_value_at(&self, row: u32, col: u32) -> PyResult<String> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.formatted_value_at(row, col as u16))
    }

    #[getter]
    fn default_row_height(&self) -> PyResult<f64> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.default_row_height())
    }

    #[getter]
    fn default_column_width(&self) -> PyResult<f64> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.default_column_width())
    }

    fn is_row_hidden(&self, row: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_row_hidden(row))
    }

    fn is_column_hidden(&self, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_column_hidden(col as u16))
    }

    fn get_row_outline_level(&self, row: u32) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.row_outline_level(row) as u32)
    }

    fn get_column_outline_level(&self, col: u32) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.column_outline_level(col as u16) as u32)
    }

    fn is_row_collapsed(&self, row: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_row_collapsed(row))
    }

    fn is_column_collapsed(&self, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_column_collapsed(col as u16))
    }

    #[getter]
    fn freeze_panes(&self) -> PyResult<Option<PyFreezePanes>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.freeze_panes().map(PyFreezePanes::from))
    }

    #[getter]
    fn split_panes(&self) -> PyResult<Option<PySplitPanes>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.split_panes().map(PySplitPanes::from))
    }

    fn get_hyperlink(&self, address: String) -> PyResult<Option<PyHyperlink>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.hyperlink(&address).map(PyHyperlink::from))
    }

    #[getter]
    fn hyperlink_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.hyperlink_count() as u32)
    }

    #[getter]
    fn hyperlinks(&self) -> PyResult<Vec<PyHyperlinkEntry>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .hyperlinks()
            .iter()
            .map(|(addr, hl)| PyHyperlinkEntry {
                address: addr.to_string(),
                hyperlink: PyHyperlink::from(hl),
            })
            .collect())
    }

    fn get_comment(&self, address: String) -> PyResult<Option<PyComment>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        let comment = ws.comment(&address).map_err(to_py_err)?;
        Ok(comment.map(PyComment::from))
    }

    fn get_comment_at(&self, row: u32, col: u32) -> PyResult<Option<PyComment>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.comment_at(row, col as u16).map(PyComment::from))
    }

    fn has_comment(&self, address: String) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        ws.has_comment(&address).map_err(to_py_err)
    }

    fn has_comment_at(&self, row: u32, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.has_comment_at(row, col as u16))
    }

    #[getter]
    fn comment_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.comment_count() as u32)
    }

    #[getter]
    fn comments(&self) -> PyResult<Vec<PyCommentEntry>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .comments()
            .map(|((row, col), c)| PyCommentEntry {
                row,
                col: col as u32,
                comment: PyComment::from(c),
            })
            .collect())
    }

    #[getter]
    fn comment_authors(&self) -> PyResult<Vec<String>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.comment_authors().to_vec())
    }

    #[getter]
    fn tables(&self) -> PyResult<Vec<PyTable>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.tables().iter().map(PyTable::from).collect())
    }

    fn get_table_by_name(&self, name: String) -> PyResult<Option<PyTable>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.table_by_name(&name).map(PyTable::from))
    }

    #[getter]
    fn table_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.table_count() as u32)
    }

    #[getter]
    fn data_validations(&self) -> PyResult<Vec<PyDataValidation>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .data_validations()
            .iter()
            .map(PyDataValidation::from)
            .collect())
    }

    #[getter]
    fn data_validation_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.data_validation_count() as u32)
    }

    #[getter]
    fn conditional_formats(&self) -> PyResult<Vec<PyConditionalFormatRule>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .conditional_formats()
            .iter()
            .map(PyConditionalFormatRule::from)
            .collect())
    }

    #[getter]
    fn conditional_format_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.conditional_format_count() as u32)
    }

    #[getter]
    fn auto_filter(&self) -> PyResult<Option<PyAutoFilter>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.auto_filter().map(PyAutoFilter::from))
    }

    #[getter]
    fn protection(&self) -> PyResult<Option<PySheetProtection>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.protection().map(PySheetProtection::from))
    }

    #[getter]
    fn page_setup(&self) -> PyResult<PyPageSetup> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(PyPageSetup::from(ws.page_setup()))
    }

    #[getter]
    fn print_area(&self) -> PyResult<Option<String>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.print_area().map(|r| r.to_string()))
    }

    #[getter]
    fn repeat_rows(&self) -> PyResult<Option<Vec<u32>>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.repeat_rows().map(|(s, e)| vec![s, e]))
    }

    #[getter]
    fn repeat_cols(&self) -> PyResult<Option<Vec<u32>>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.repeat_cols().map(|(s, e)| vec![s as u32, e as u32]))
    }

    #[getter]
    fn row_breaks(&self) -> PyResult<Vec<PyPageBreak>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.row_breaks().iter().map(PyPageBreak::from).collect())
    }

    #[getter]
    fn col_breaks(&self) -> PyResult<Vec<PyPageBreak>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.col_breaks().iter().map(PyPageBreak::from).collect())
    }

    fn get_formula_at(&self, row: u32, col: u32) -> PyResult<Option<String>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.get_formula_at(row, col as u16).map(|s| s.to_string()))
    }

    #[getter]
    fn formula_cells(&self) -> PyResult<Vec<PyFormulaCell>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .formula_cells()
            .map(|(row, col, formula)| PyFormulaCell {
                row,
                col: col as u32,
                formula: formula.to_string(),
            })
            .collect())
    }

    fn is_spill_target(&self, row: u32, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_spill_target(row, col as u16))
    }

    fn is_spill_source(&self, row: u32, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_spill_source(row, col as u16))
    }

    fn get_spill_source(&self, row: u32, col: u32) -> PyResult<Option<PySpillSource>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .get_spill_source(row, col as u16)
            .map(|(r, c)| PySpillSource {
                row: r,
                col: c as u32,
            }))
    }

    #[getter]
    fn date_1904(&self) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.date_1904())
    }

    #[getter]
    fn merged_regions(&self) -> PyResult<Vec<String>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.merged_regions().iter().map(|r| r.to_string()).collect())
    }
}
