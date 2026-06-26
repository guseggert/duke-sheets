use std::collections::BTreeSet;

use wasm_bindgen::prelude::*;

use crate::{
    to_js_error, to_js_value,
    types::{
        WasmAutoFilter, WasmChart, WasmChartEx, WasmColor, WasmComment, WasmCommentEntry,
        WasmConditionalFormatRule, WasmDataValidation, WasmEmbeddedImage, WasmFormulaCell,
        WasmFreezePanes, WasmHyperlink, WasmHyperlinkEntry, WasmImageInfo, WasmMergeSpan,
        WasmMergedRegion, WasmPageBreak, WasmPageSetup, WasmProtectedRange, WasmRow, WasmRowCell,
        WasmRowsOptions, WasmSelection, WasmSheetProtection, WasmSpillSource, WasmSplitPanes,
        WasmStyle, WasmTable,
    },
    Worksheet,
};

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter)]
    pub fn visibility(&self) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(match ws.visibility() {
            duke_sheets_core::SheetVisibility::Visible => "visible".to_string(),
            duke_sheets_core::SheetVisibility::Hidden => "hidden".to_string(),
            duke_sheets_core::SheetVisibility::VeryHidden => "veryHidden".to_string(),
        })
    }

    #[wasm_bindgen(getter, js_name = isSelected)]
    pub fn is_selected(&self) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_selected())
    }

    #[wasm_bindgen(getter, js_name = zoomScale)]
    pub fn zoom_scale(&self) -> Result<Option<u32>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.zoom_scale().map(|z| z as u32))
    }

    #[wasm_bindgen(getter, js_name = tabColor)]
    pub fn tab_color(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.tab_color() {
            Some(c) => to_js_value(&WasmColor::from(&c)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = isEmpty)]
    pub fn ws_is_empty(&self) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_empty())
    }

    #[wasm_bindgen(getter, js_name = cellCount)]
    pub fn cell_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.cell_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = selections)]
    pub fn selections(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let selections: Vec<WasmSelection> =
            ws.selections().iter().map(WasmSelection::from).collect();
        to_js_value(&selections)
    }

    #[wasm_bindgen(js_name = getCellStyle)]
    pub fn get_cell_style(&self, address: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let style = ws.cell_style(address).map_err(to_js_error)?;
        match style {
            Some(s) => to_js_value(&WasmStyle::from(s)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getCellStyleAt)]
    pub fn get_cell_style_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.cell_style_at(row, col as u16) {
            Some(s) => to_js_value(&WasmStyle::from(s)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getFormattedValue)]
    pub fn get_formatted_value(&self, address: &str) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.formatted_value(address).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getFormattedValueAt)]
    pub fn get_formatted_value_at(&self, row: u32, col: u32) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.formatted_value_at(row, col as u16))
    }

    #[wasm_bindgen(skip_typescript, js_name = getRowsBatch)]
    pub fn get_rows_batch(
        &self,
        start_row: u32,
        max_rows: u32,
        options: JsValue,
    ) -> Result<JsValue, JsValue> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsValue::from(JsError::new("Worksheet no longer exists")))?;

        let options = if options.is_null() || options.is_undefined() {
            None
        } else {
            Some(
                serde_wasm_bindgen::from_value::<WasmRowsOptions>(options)
                    .map_err(|e| JsValue::from(to_js_error(e)))?,
            )
        };

        let use_formatted = options
            .as_ref()
            .and_then(|o| o.use_formatted_values)
            .unwrap_or(false);
        let use_calculated = options
            .as_ref()
            .and_then(|o| o.use_calculated_values)
            .unwrap_or(false);
        let inc_styles = options
            .as_ref()
            .and_then(|o| o.include_styles)
            .unwrap_or(false);
        let inc_merge = options
            .as_ref()
            .and_then(|o| o.include_merge_info)
            .unwrap_or(false);
        let inc_hyperlinks = options
            .as_ref()
            .and_then(|o| o.include_hyperlinks)
            .unwrap_or(false);
        let inc_comments = options
            .as_ref()
            .and_then(|o| o.include_comments)
            .unwrap_or(false);
        let inc_formulas = options
            .as_ref()
            .and_then(|o| o.include_formulas)
            .unwrap_or(false);
        let inc_images = options
            .as_ref()
            .and_then(|o| o.include_images)
            .unwrap_or(false);
        let skip_empty = options
            .as_ref()
            .and_then(|o| o.skip_empty_values)
            .unwrap_or(false);
        let skip_blank = options
            .as_ref()
            .and_then(|o| o.skip_blank_values)
            .unwrap_or(false);

        let max_row = ws.used_range().map(|r| r.end.row).unwrap_or(0);
        let end_row = max_row.min(start_row.saturating_add(max_rows).saturating_sub(1));

        if start_row > end_row {
            let rows: Vec<WasmRow> = Vec::new();
            return to_js_value(&rows).map_err(JsValue::from);
        }

        let mut coords: BTreeSet<(u32, u16)> = ws
            .populated_cells_in_range(start_row, end_row)
            .into_iter()
            .collect();

        if inc_styles {
            for (&(row, col), data) in ws.cells_map_in_range(start_row, end_row) {
                if data.style_index != 0 {
                    coords.insert((row, col));
                }
            }
        }
        if inc_merge {
            for region in ws.merged_regions() {
                for row in region.start.row.max(start_row)..=region.end.row.min(end_row) {
                    for col in region.start.col..=region.end.col {
                        coords.insert((row, col));
                    }
                }
            }
        }
        if inc_hyperlinks {
            for (addr, _) in ws.hyperlinks() {
                if addr.row >= start_row && addr.row <= end_row {
                    coords.insert((addr.row, addr.col));
                }
            }
        }
        if inc_comments {
            for ((row, col), _) in ws.comments() {
                if row >= start_row && row <= end_row {
                    coords.insert((row, col));
                }
            }
        }
        if inc_formulas {
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
                        rows.push(WasmRow {
                            index: prev_row,
                            cells: std::mem::take(&mut current_cells),
                        });
                    }
                }
                current_row = Some(row);
            }

            let value = if use_formatted {
                ws.formatted_value_at(row, col)
            } else if use_calculated {
                ws.get_calculated_value_at(row, col)
                    .map(|v| v.to_string())
                    .unwrap_or_default()
            } else {
                ws.get_value_at(row, col).to_string()
            };

            if skip_blank || skip_empty {
                let raw = ws.get_value_at(row, col);
                if skip_blank && raw.is_blank() {
                    continue;
                }
                if skip_empty && raw.is_empty() {
                    continue;
                }
            }

            let style = if inc_styles {
                ws.cell_style_at(row, col).map(WasmStyle::from)
            } else {
                None
            };

            let merge_span = if inc_merge {
                ws.get_merge_span(row, col)
                    .map(|(row_span, col_span)| WasmMergeSpan {
                        row_span,
                        col_span: col_span as u32,
                    })
            } else {
                None
            };

            let is_merged_secondary = if inc_merge {
                let value = ws.is_merged_secondary(row, col);
                if value {
                    Some(true)
                } else {
                    None
                }
            } else {
                None
            };

            let hyperlink = if inc_hyperlinks {
                ws.hyperlink_at(row, col).map(WasmHyperlink::from)
            } else {
                None
            };

            let comment = if inc_comments {
                ws.comment_at(row, col).map(WasmComment::from)
            } else {
                None
            };

            let formula = if inc_formulas {
                ws.get_formula_at(row, col)
                    .map(|formula| formula.to_string())
            } else {
                None
            };

            let image = if inc_images {
                ws.get_image_at(row, col).map(WasmImageInfo::from)
            } else {
                None
            };

            current_cells.push(WasmRowCell {
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
                rows.push(WasmRow {
                    index: last_row,
                    cells: current_cells,
                });
            }
        }

        to_js_value(&rows).map_err(JsValue::from)
    }

    #[wasm_bindgen(getter, js_name = defaultRowHeight)]
    pub fn default_row_height(&self) -> Result<f64, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.default_row_height())
    }

    #[wasm_bindgen(getter, js_name = defaultColumnWidth)]
    pub fn default_column_width(&self) -> Result<f64, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.default_column_width())
    }

    #[wasm_bindgen(js_name = isRowHidden)]
    pub fn is_row_hidden(&self, row: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_row_hidden(row))
    }

    #[wasm_bindgen(js_name = isColumnHidden)]
    pub fn is_column_hidden(&self, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_column_hidden(col as u16))
    }

    #[wasm_bindgen(js_name = getRowOutlineLevel)]
    pub fn get_row_outline_level(&self, row: u32) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.row_outline_level(row) as u32)
    }

    #[wasm_bindgen(js_name = getColumnOutlineLevel)]
    pub fn get_column_outline_level(&self, col: u32) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.column_outline_level(col as u16) as u32)
    }

    #[wasm_bindgen(js_name = isRowCollapsed)]
    pub fn is_row_collapsed(&self, row: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_row_collapsed(row))
    }

    #[wasm_bindgen(js_name = isColumnCollapsed)]
    pub fn is_column_collapsed(&self, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_column_collapsed(col as u16))
    }

    #[wasm_bindgen(getter, js_name = freezePanes)]
    pub fn freeze_panes(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.freeze_panes() {
            Some(v) => to_js_value(&WasmFreezePanes::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = splitPanes)]
    pub fn split_panes(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.split_panes() {
            Some(v) => to_js_value(&WasmSplitPanes::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getHyperlink)]
    pub fn get_hyperlink(&self, address: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.hyperlink(address) {
            Some(v) => to_js_value(&WasmHyperlink::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getHyperlinkAt)]
    pub fn get_hyperlink_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.hyperlink_at(row, col as u16) {
            Some(v) => to_js_value(&WasmHyperlink::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = hyperlinkCount)]
    pub fn hyperlink_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.hyperlink_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = hyperlinks)]
    pub fn hyperlinks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let hyperlinks: Vec<WasmHyperlinkEntry> = ws
            .hyperlinks()
            .iter()
            .map(|(addr, hl)| WasmHyperlinkEntry {
                address: addr.to_string(),
                hyperlink: WasmHyperlink::from(hl),
            })
            .collect();
        to_js_value(&hyperlinks)
    }

    #[wasm_bindgen(js_name = getComment)]
    pub fn get_comment(&self, address: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let comment = ws.comment(address).map_err(to_js_error)?;
        match comment {
            Some(v) => to_js_value(&WasmComment::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getCommentAt)]
    pub fn get_comment_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.comment_at(row, col as u16) {
            Some(v) => to_js_value(&WasmComment::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = hasComment)]
    pub fn has_comment(&self, address: &str) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.has_comment(address).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = hasCommentAt)]
    pub fn has_comment_at(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.has_comment_at(row, col as u16))
    }

    #[wasm_bindgen(getter, js_name = commentCount)]
    pub fn comment_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.comment_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = comments)]
    pub fn comments(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let comments: Vec<WasmCommentEntry> = ws
            .comments()
            .map(|((row, col), c)| WasmCommentEntry {
                row,
                col: col as u32,
                comment: WasmComment::from(c),
            })
            .collect();
        to_js_value(&comments)
    }

    #[wasm_bindgen(getter, js_name = commentAuthors)]
    pub fn comment_authors(&self) -> Result<Vec<String>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.comment_authors().to_vec())
    }

    #[wasm_bindgen(getter, js_name = tables)]
    pub fn tables(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let tables: Vec<WasmTable> = ws.tables().iter().map(WasmTable::from).collect();
        to_js_value(&tables)
    }

    #[wasm_bindgen(js_name = getTableByName)]
    pub fn get_table_by_name(&self, name: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.table_by_name(name) {
            Some(v) => to_js_value(&WasmTable::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = tableCount)]
    pub fn table_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.table_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = dataValidations)]
    pub fn data_validations(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let rules: Vec<WasmDataValidation> = ws
            .data_validations()
            .iter()
            .map(WasmDataValidation::from)
            .collect();
        to_js_value(&rules)
    }

    #[wasm_bindgen(getter, js_name = dataValidationCount)]
    pub fn data_validation_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.data_validation_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = conditionalFormats)]
    pub fn conditional_formats(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let rules: Vec<WasmConditionalFormatRule> = ws
            .conditional_formats()
            .iter()
            .map(WasmConditionalFormatRule::from)
            .collect();
        to_js_value(&rules)
    }

    #[wasm_bindgen(getter, js_name = conditionalFormatCount)]
    pub fn conditional_format_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.conditional_format_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = autoFilter)]
    pub fn auto_filter(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.auto_filter() {
            Some(v) => to_js_value(&WasmAutoFilter::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = protection)]
    pub fn protection(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.protection() {
            Some(v) => to_js_value(&WasmSheetProtection::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = protectedRanges)]
    pub fn protected_ranges(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let ranges: Vec<WasmProtectedRange> = ws
            .protected_ranges()
            .iter()
            .map(WasmProtectedRange::from)
            .collect();
        to_js_value(&ranges)
    }

    #[wasm_bindgen(getter, js_name = pageSetup)]
    pub fn page_setup(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        to_js_value(&WasmPageSetup::from(ws.page_setup()))
    }

    #[wasm_bindgen(getter, js_name = printArea)]
    pub fn print_area(&self) -> Result<Option<String>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.print_area().map(|r| r.to_string()))
    }

    #[wasm_bindgen(getter, js_name = repeatRows)]
    pub fn repeat_rows(&self) -> Result<Option<Vec<u32>>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.repeat_rows().map(|(s, e)| vec![s, e]))
    }

    #[wasm_bindgen(getter, js_name = repeatCols)]
    pub fn repeat_cols(&self) -> Result<Option<Vec<u32>>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.repeat_cols().map(|(s, e)| vec![s as u32, e as u32]))
    }

    #[wasm_bindgen(getter, js_name = rowBreaks)]
    pub fn row_breaks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let breaks: Vec<WasmPageBreak> = ws.row_breaks().iter().map(WasmPageBreak::from).collect();
        to_js_value(&breaks)
    }

    #[wasm_bindgen(getter, js_name = colBreaks)]
    pub fn col_breaks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let breaks: Vec<WasmPageBreak> = ws.col_breaks().iter().map(WasmPageBreak::from).collect();
        to_js_value(&breaks)
    }

    #[wasm_bindgen(js_name = getFormulaAt)]
    pub fn get_formula_at(&self, row: u32, col: u32) -> Result<Option<String>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.get_formula_at(row, col as u16).map(|s| s.to_string()))
    }

    #[wasm_bindgen(getter, js_name = formulaCells)]
    pub fn formula_cells(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let formulas: Vec<WasmFormulaCell> = ws
            .formula_cells()
            .map(|(row, col, formula)| WasmFormulaCell {
                row,
                col: col as u32,
                formula: formula.to_string(),
            })
            .collect();
        to_js_value(&formulas)
    }

    /// Get the number of formula cells in this worksheet.
    #[wasm_bindgen(getter, js_name = formulaCount)]
    pub fn formula_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.formula_cells().count() as u32)
    }

    #[wasm_bindgen(js_name = isSpillTarget)]
    pub fn is_spill_target(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_spill_target(row, col as u16))
    }

    #[wasm_bindgen(js_name = isSpillSource)]
    pub fn is_spill_source(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_spill_source(row, col as u16))
    }

    #[wasm_bindgen(js_name = getSpillSource)]
    pub fn get_spill_source(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.get_spill_source(row, col as u16) {
            Some((r, c)) => to_js_value(&WasmSpillSource {
                row: r,
                col: c as u32,
            }),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = date1904)]
    pub fn date_1904(&self) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.date_1904())
    }

    #[wasm_bindgen(getter, js_name = mergedRegions)]
    pub fn merged_regions(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let regions: Vec<WasmMergedRegion> = ws
            .merged_regions()
            .iter()
            .map(|r| WasmMergedRegion {
                start_row: r.start.row,
                start_col: r.start.col as u32,
                end_row: r.end.row,
                end_col: r.end.col as u32,
                range: r.to_string(),
            })
            .collect();
        to_js_value(&regions)
    }

    #[wasm_bindgen(js_name = getMergeSpan)]
    pub fn get_merge_span(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.get_merge_span(row, col as u16) {
            Some((rs, cs)) => to_js_value(&WasmMergeSpan {
                row_span: rs,
                col_span: cs as u32,
            }),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = isMergedSecondary)]
    pub fn is_merged_secondary(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_merged_secondary(row, col as u16))
    }

    #[wasm_bindgen(getter, js_name = charts)]
    pub fn charts(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let charts: Vec<WasmChart> = ws.charts().iter().map(WasmChart::from).collect();
        to_js_value(&charts)
    }

    #[wasm_bindgen(getter, js_name = chartCount)]
    pub fn chart_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.chart_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = chartsEx)]
    pub fn charts_ex(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let charts: Vec<WasmChartEx> = ws.charts_ex().iter().map(WasmChartEx::from).collect();
        to_js_value(&charts)
    }

    #[wasm_bindgen(getter, js_name = chartExCount)]
    pub fn chart_ex_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.chart_ex_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = images)]
    pub fn images(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let images: Vec<WasmEmbeddedImage> =
            ws.images().iter().map(WasmEmbeddedImage::from).collect();
        to_js_value(&images)
    }

    #[wasm_bindgen(getter, js_name = imageCount)]
    pub fn image_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.image_count() as u32)
    }
}
