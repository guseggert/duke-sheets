use wasm_bindgen::prelude::*;

use crate::{
    to_js_error, to_js_value,
    types::{
        WasmAutoFilter, WasmColor, WasmComment, WasmCommentEntry, WasmConditionalFormatRule,
        WasmDataValidation, WasmFormulaCell, WasmFreezePanes, WasmHyperlink, WasmHyperlinkEntry,
        WasmPageBreak, WasmPageSetup, WasmSelection, WasmSheetProtection, WasmSpillSource,
        WasmSplitPanes, WasmStyle, WasmTable,
    },
    Worksheet,
};

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter, js_name = isVisible)]
    pub fn is_visible(&self) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_visible())
    }

    #[wasm_bindgen(getter, js_name = isSelected)]
    pub fn is_selected(&self) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_selected())
    }

    #[wasm_bindgen(getter, js_name = zoomScale)]
    pub fn zoom_scale(&self) -> Result<Option<u32>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.zoom_scale().map(|z| z as u32))
    }

    #[wasm_bindgen(getter, js_name = tabColor)]
    pub fn tab_color(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_empty())
    }

    #[wasm_bindgen(getter, js_name = cellCount)]
    pub fn cell_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.cell_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = selections)]
    pub fn selections(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let selections: Vec<WasmSelection> =
            ws.selections().iter().map(WasmSelection::from).collect();
        to_js_value(&selections)
    }

    #[wasm_bindgen(js_name = getCellStyle)]
    pub fn get_cell_style(&self, address: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.formatted_value(address).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getFormattedValueAt)]
    pub fn get_formatted_value_at(&self, row: u32, col: u32) -> Result<String, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.formatted_value_at(row, col as u16))
    }

    #[wasm_bindgen(getter, js_name = defaultRowHeight)]
    pub fn default_row_height(&self) -> Result<f64, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.default_row_height())
    }

    #[wasm_bindgen(getter, js_name = defaultColumnWidth)]
    pub fn default_column_width(&self) -> Result<f64, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.default_column_width())
    }

    #[wasm_bindgen(js_name = isRowHidden)]
    pub fn is_row_hidden(&self, row: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_row_hidden(row))
    }

    #[wasm_bindgen(js_name = isColumnHidden)]
    pub fn is_column_hidden(&self, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_column_hidden(col as u16))
    }

    #[wasm_bindgen(js_name = getRowOutlineLevel)]
    pub fn get_row_outline_level(&self, row: u32) -> Result<u32, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.row_outline_level(row) as u32)
    }

    #[wasm_bindgen(js_name = getColumnOutlineLevel)]
    pub fn get_column_outline_level(&self, col: u32) -> Result<u32, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.column_outline_level(col as u16) as u32)
    }

    #[wasm_bindgen(js_name = isRowCollapsed)]
    pub fn is_row_collapsed(&self, row: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_row_collapsed(row))
    }

    #[wasm_bindgen(js_name = isColumnCollapsed)]
    pub fn is_column_collapsed(&self, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_column_collapsed(col as u16))
    }

    #[wasm_bindgen(getter, js_name = freezePanes)]
    pub fn freeze_panes(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.hyperlink(address) {
            Some(v) => to_js_value(&WasmHyperlink::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = hyperlinkCount)]
    pub fn hyperlink_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.hyperlink_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = hyperlinks)]
    pub fn hyperlinks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.has_comment(address).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = hasCommentAt)]
    pub fn has_comment_at(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.has_comment_at(row, col as u16))
    }

    #[wasm_bindgen(getter, js_name = commentCount)]
    pub fn comment_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.comment_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = comments)]
    pub fn comments(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.comment_authors().to_vec())
    }

    #[wasm_bindgen(getter, js_name = tables)]
    pub fn tables(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let tables: Vec<WasmTable> = ws.tables().iter().map(WasmTable::from).collect();
        to_js_value(&tables)
    }

    #[wasm_bindgen(js_name = getTableByName)]
    pub fn get_table_by_name(&self, name: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.table_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = dataValidations)]
    pub fn data_validations(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.data_validation_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = conditionalFormats)]
    pub fn conditional_formats(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.conditional_format_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = autoFilter)]
    pub fn auto_filter(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.protection() {
            Some(v) => to_js_value(&WasmSheetProtection::from(v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = pageSetup)]
    pub fn page_setup(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        to_js_value(&WasmPageSetup::from(ws.page_setup()))
    }

    #[wasm_bindgen(getter, js_name = printArea)]
    pub fn print_area(&self) -> Result<Option<String>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.print_area().map(|r| r.to_string()))
    }

    #[wasm_bindgen(getter, js_name = repeatRows)]
    pub fn repeat_rows(&self) -> Result<Option<Vec<u32>>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.repeat_rows().map(|(s, e)| vec![s, e]))
    }

    #[wasm_bindgen(getter, js_name = repeatCols)]
    pub fn repeat_cols(&self) -> Result<Option<Vec<u32>>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.repeat_cols().map(|(s, e)| vec![s as u32, e as u32]))
    }

    #[wasm_bindgen(getter, js_name = rowBreaks)]
    pub fn row_breaks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let breaks: Vec<WasmPageBreak> = ws.row_breaks().iter().map(WasmPageBreak::from).collect();
        to_js_value(&breaks)
    }

    #[wasm_bindgen(getter, js_name = colBreaks)]
    pub fn col_breaks(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let breaks: Vec<WasmPageBreak> = ws.col_breaks().iter().map(WasmPageBreak::from).collect();
        to_js_value(&breaks)
    }

    #[wasm_bindgen(js_name = getFormulaAt)]
    pub fn get_formula_at(&self, row: u32, col: u32) -> Result<Option<String>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.get_formula_at(row, col as u16).map(|s| s.to_string()))
    }

    #[wasm_bindgen(getter, js_name = formulaCells)]
    pub fn formula_cells(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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

    #[wasm_bindgen(js_name = isSpillTarget)]
    pub fn is_spill_target(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_spill_target(row, col as u16))
    }

    #[wasm_bindgen(js_name = isSpillSource)]
    pub fn is_spill_source(&self, row: u32, col: u32) -> Result<bool, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.is_spill_source(row, col as u16))
    }

    #[wasm_bindgen(js_name = getSpillSource)]
    pub fn get_spill_source(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
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
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.date_1904())
    }

    #[wasm_bindgen(getter, js_name = mergedRegions)]
    pub fn merged_regions(&self) -> Result<Vec<String>, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.merged_regions().iter().map(|r| r.to_string()).collect())
    }
}
