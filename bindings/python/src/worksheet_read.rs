use std::collections::BTreeSet;
use std::sync::{Arc, RwLock};

use duke_sheets::Workbook;
use pyo3::exceptions::PyIndexError;
use pyo3::prelude::*;

use crate::{
    image_sizing_to_python, to_py_err, PyAutoFilter, PyCalculationImage, PyChart, PyColor,
    PyComment, PyCommentEntry, PyConditionalFormatRule, PyDataValidation, PyFormulaCell,
    PyFreezePanes, PyHyperlink, PyHyperlinkEntry, PyMergeSpan, PyMergedRegion, PyPageBreak,
    PyPageSetup, PyRow, PyRowCell, PySelection, PySheetProtection, PySpillSource, PySplitPanes,
    PyStyle, PyTable, PyWorksheet,
};

const ROW_ITER_BATCH_SIZE: u32 = 1000;

fn sparse_rows_batch(
    workbook: &Arc<RwLock<Workbook>>,
    sheet_index: usize,
    start_row: u32,
    max_rows: u32,
    use_formatted_values: bool,
    use_calculated_values: bool,
    include_styles: bool,
    include_merge_info: bool,
    include_hyperlinks: bool,
    include_comments: bool,
    include_formulas: bool,
    include_images: bool,
    skip_empty_values: bool,
    skip_blank_values: bool,
) -> PyResult<Vec<PyRow>> {
    let wb = workbook.read().map_err(to_py_err)?;
    let ws = wb
        .worksheet(sheet_index)
        .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;

    let end_row = ws
        .used_range()
        .map(|r| r.end.row)
        .unwrap_or(0)
        .min(start_row.saturating_add(max_rows).saturating_sub(1));

    if start_row > end_row {
        return Ok(vec![]);
    }

    let mut coords: BTreeSet<(u32, u16)> = ws
        .populated_cells_in_range(start_row, end_row)
        .into_iter()
        .collect();

    if include_styles {
        for (&(row, col), data) in ws.cells_map_in_range(start_row, end_row) {
            if data.style_index != 0 {
                coords.insert((row, col));
            }
        }
    }

    if include_merge_info {
        for region in ws.merged_regions() {
            for row in region.start.row.max(start_row)..=region.end.row.min(end_row) {
                for col in region.start.col..=region.end.col {
                    coords.insert((row, col));
                }
            }
        }
    }

    if include_hyperlinks {
        for (addr, _) in ws.hyperlinks() {
            if addr.row >= start_row && addr.row <= end_row {
                coords.insert((addr.row, addr.col));
            }
        }
    }

    if include_comments {
        for ((row, col), _) in ws.comments() {
            if row >= start_row && row <= end_row {
                coords.insert((row, col));
            }
        }
    }

    if include_formulas {
        for (row, col, _) in ws.formula_cells() {
            if row >= start_row && row <= end_row {
                coords.insert((row, col));
            }
        }
    }

    let mut rows = Vec::new();
    let mut current_row = None;
    let mut current_cells = Vec::new();

    for (row, col) in &coords {
        let (row, col) = (*row, *col);

        if current_row != Some(row) {
            if let Some(prev_row) = current_row {
                if !current_cells.is_empty() {
                    rows.push(PyRow {
                        index: prev_row,
                        cells: std::mem::take(&mut current_cells),
                    });
                }
            }
            current_row = Some(row);
        }

        let value = if use_formatted_values {
            ws.formatted_value_at(row, col)
        } else if use_calculated_values {
            ws.get_calculated_value_at(row, col)
                .map(|value| value.to_string())
                .unwrap_or_default()
        } else {
            ws.get_value_at(row, col).to_string()
        };

        if skip_blank_values || skip_empty_values {
            let raw = ws.get_value_at(row, col);
            if skip_blank_values && raw.is_blank() {
                continue;
            }
            if skip_empty_values && raw.is_empty() {
                continue;
            }
        }

        let style = if include_styles {
            ws.cell_style_at(row, col).map(PyStyle::from)
        } else {
            None
        };

        let merge_span = if include_merge_info {
            ws.get_merge_span(row, col)
                .map(|(row_span, col_span)| PyMergeSpan {
                    row_span,
                    col_span: col_span as u32,
                })
        } else {
            None
        };

        let is_merged_secondary = if include_merge_info {
            let is_secondary = ws.is_merged_secondary(row, col);
            if is_secondary {
                Some(true)
            } else {
                None
            }
        } else {
            None
        };

        let hyperlink = if include_hyperlinks {
            ws.hyperlink_at(row, col).map(PyHyperlink::from)
        } else {
            None
        };

        let comment = if include_comments {
            ws.comment_at(row, col).map(PyComment::from)
        } else {
            None
        };

        let formula = if include_formulas {
            ws.get_formula_at(row, col)
                .map(|formula| formula.to_string())
        } else {
            None
        };

        let image = if include_images {
            ws.get_image_at(row, col).map(|info| PyCalculationImage {
                source: info.source,
                alt_text: info.alt_text,
                sizing: image_sizing_to_python(info.sizing).to_string(),
                width: info.width,
                height: info.height,
            })
        } else {
            None
        };

        current_cells.push(PyRowCell {
            col: col as u32,
            value,
            style,
            merge_span,
            is_merged_secondary,
            hyperlink,
            comment,
            formula,
            image,
        });
    }

    if let Some(last_row) = current_row {
        if !current_cells.is_empty() {
            rows.push(PyRow {
                index: last_row,
                cells: current_cells,
            });
        }
    }

    Ok(rows)
}

fn worksheet_max_row(workbook: &Arc<RwLock<Workbook>>, sheet_index: usize) -> PyResult<u32> {
    let wb = workbook.read().map_err(to_py_err)?;
    let ws = wb
        .worksheet(sheet_index)
        .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
    Ok(ws.used_range().map(|range| range.end.row).unwrap_or(0))
}

#[pyclass(name = "RowIterator")]
pub struct PyRowIterator {
    workbook: Arc<RwLock<Workbook>>,
    sheet_index: usize,
    use_formatted_values: bool,
    use_calculated_values: bool,
    include_styles: bool,
    include_merge_info: bool,
    include_hyperlinks: bool,
    include_comments: bool,
    include_formulas: bool,
    include_images: bool,
    skip_empty_values: bool,
    skip_blank_values: bool,
    next_row: u32,
    max_row: u32,
    buffer: Vec<PyRow>,
    cursor: usize,
    done: bool,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> PyResult<Option<PyRow>> {
        while self.cursor >= self.buffer.len() {
            if self.done || self.next_row > self.max_row {
                self.done = true;
                return Ok(None);
            }

            let batch_size = ROW_ITER_BATCH_SIZE
                .min(self.max_row.saturating_sub(self.next_row).saturating_add(1));
            self.buffer = sparse_rows_batch(
                &self.workbook,
                self.sheet_index,
                self.next_row,
                batch_size,
                self.use_formatted_values,
                self.use_calculated_values,
                self.include_styles,
                self.include_merge_info,
                self.include_hyperlinks,
                self.include_comments,
                self.include_formulas,
                self.include_images,
                self.skip_empty_values,
                self.skip_blank_values,
            )?;
            self.cursor = 0;

            if self.buffer.is_empty() {
                self.next_row = self.next_row.saturating_add(batch_size);
                continue;
            }

            self.next_row = self
                .buffer
                .last()
                .map(|row| row.index.saturating_add(1))
                .unwrap_or(self.next_row.saturating_add(batch_size));
        }

        let row = self.buffer[self.cursor].clone();
        self.cursor += 1;
        Ok(Some(row))
    }
}

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

    #[pyo3(signature = (
        start_row,
        max_rows,
        *,
        use_formatted_values=false,
        use_calculated_values=false,
        include_styles=None,
        include_merge_info=None,
        include_hyperlinks=None,
        include_comments=None,
        include_formulas=None,
        include_images=None,
        skip_empty_values=false,
        skip_blank_values=false
    ))]
    fn get_rows_batch(
        &self,
        start_row: u32,
        max_rows: u32,
        use_formatted_values: bool,
        use_calculated_values: bool,
        include_styles: Option<bool>,
        include_merge_info: Option<bool>,
        include_hyperlinks: Option<bool>,
        include_comments: Option<bool>,
        include_formulas: Option<bool>,
        include_images: Option<bool>,
        skip_empty_values: bool,
        skip_blank_values: bool,
    ) -> PyResult<Vec<PyRow>> {
        sparse_rows_batch(
            &self.workbook,
            self.sheet_index,
            start_row,
            max_rows,
            use_formatted_values,
            use_calculated_values,
            include_styles.unwrap_or(false),
            include_merge_info.unwrap_or(false),
            include_hyperlinks.unwrap_or(false),
            include_comments.unwrap_or(false),
            include_formulas.unwrap_or(false),
            include_images.unwrap_or(false),
            skip_empty_values,
            skip_blank_values,
        )
    }

    #[pyo3(signature = (
        *,
        use_formatted_values=false,
        use_calculated_values=false,
        include_styles=None,
        include_merge_info=None,
        include_hyperlinks=None,
        include_comments=None,
        include_formulas=None,
        include_images=None,
        skip_empty_values=false,
        skip_blank_values=false
    ))]
    fn iterate_rows(
        &self,
        use_formatted_values: bool,
        use_calculated_values: bool,
        include_styles: Option<bool>,
        include_merge_info: Option<bool>,
        include_hyperlinks: Option<bool>,
        include_comments: Option<bool>,
        include_formulas: Option<bool>,
        include_images: Option<bool>,
        skip_empty_values: bool,
        skip_blank_values: bool,
    ) -> PyResult<PyRowIterator> {
        Ok(PyRowIterator {
            workbook: Arc::clone(&self.workbook),
            sheet_index: self.sheet_index,
            use_formatted_values,
            use_calculated_values,
            include_styles: include_styles.unwrap_or(false),
            include_merge_info: include_merge_info.unwrap_or(false),
            include_hyperlinks: include_hyperlinks.unwrap_or(false),
            include_comments: include_comments.unwrap_or(false),
            include_formulas: include_formulas.unwrap_or(false),
            include_images: include_images.unwrap_or(false),
            skip_empty_values,
            skip_blank_values,
            next_row: 0,
            max_row: worksheet_max_row(&self.workbook, self.sheet_index)?,
            buffer: Vec::new(),
            cursor: 0,
            done: false,
        })
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

    /// Get the hyperlink on a cell by row/col (0-based), or None if none.
    fn get_hyperlink_at(&self, row: u32, col: u32) -> PyResult<Option<PyHyperlink>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.hyperlink_at(row, col as u16).map(PyHyperlink::from))
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

    /// Get the number of formula cells in this worksheet.
    #[getter]
    fn formula_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.formula_cells().count() as u32)
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
    fn merged_regions(&self) -> PyResult<Vec<PyMergedRegion>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .merged_regions()
            .iter()
            .map(|r| PyMergedRegion {
                start_row: r.start.row,
                start_col: r.start.col as u32,
                end_row: r.end.row,
                end_col: r.end.col as u32,
                range: r.to_string(),
            })
            .collect())
    }
    /// Get the merge span for a cell if it is the top-left origin of a merged region.
    ///
    /// Returns a MergeSpan with row_span/col_span if the cell is a merge origin, or None.
    fn get_merge_span(&self, row: u32, col: u32) -> PyResult<Option<PyMergeSpan>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws
            .get_merge_span(row, col as u16)
            .map(|(rs, cs)| PyMergeSpan {
                row_span: rs,
                col_span: cs as u32,
            }))
    }

    /// Whether a cell is a non-origin member of a merged region.
    fn is_merged_secondary(&self, row: u32, col: u32) -> PyResult<bool> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.is_merged_secondary(row, col as u16))
    }
    #[getter]
    fn charts(&self) -> PyResult<Vec<PyChart>> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.charts().iter().map(PyChart::from).collect())
    }

    #[getter]
    fn chart_count(&self) -> PyResult<u32> {
        let wb = self.workbook.read().map_err(to_py_err)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| PyIndexError::new_err("Worksheet no longer exists"))?;
        Ok(ws.chart_count() as u32)
    }
}
