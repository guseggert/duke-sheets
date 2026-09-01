use wasm_bindgen::prelude::*;

use crate::{
    to_js_value, types::WasmChartSheet, types::WasmNamedRange, types::WasmSheetSlot,
    types::WasmWorkbookProtection, types::WasmWorkbookSettings, Workbook,
};

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(getter, js_name = isEmpty)]
    pub fn is_empty(&self) -> Result<bool, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.is_empty())
    }

    #[wasm_bindgen(getter, js_name = activeSheet)]
    pub fn active_sheet(&self) -> Result<u32, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.active_sheet() as u32)
    }

    #[wasm_bindgen(js_name = sheetIndex)]
    pub fn sheet_index(&self, name: &str) -> Result<Option<u32>, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.sheet_index(name).map(|i| i as u32))
    }

    #[wasm_bindgen(getter, js_name = settings)]
    pub fn settings(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        to_js_value(&WasmWorkbookSettings::from(wb.settings()))
    }

    #[wasm_bindgen(getter, js_name = workbookProtection)]
    pub fn workbook_protection(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        match wb.workbook_protection() {
            Some(v) => to_js_value(&WasmWorkbookProtection::from(&v)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = namedRanges)]
    pub fn named_ranges(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        let ranges: Vec<WasmNamedRange> =
            wb.named_ranges().iter().map(WasmNamedRange::from).collect();
        to_js_value(&ranges)
    }

    /// The workbook theme's 12 clrScheme colors as `RRGGBB` hex, in
    /// theme-index order (background 1, text 1, background 2, text 2,
    /// accent 1-6, hyperlink, followed hyperlink). The Office default
    /// palette when the file carries no theme.
    #[wasm_bindgen(getter, js_name = themePalette)]
    pub fn theme_palette(&self) -> Vec<String> {
        let wb = self.inner.borrow();
        wb.theme_palette()
            .colors
            .iter()
            .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}"))
            .collect()
    }

    /// Resolve a drawing color to display RGB (`RRGGBB` hex) against
    /// this workbook's theme palette. `auto` has no fixed RGB and
    /// resolves to `null`.
    #[wasm_bindgen(js_name = resolveColor, skip_typescript)]
    pub fn resolve_color(&self, color: JsValue) -> Result<Option<String>, JsError> {
        let color: crate::drawings::WasmDrawingColor = serde_wasm_bindgen::from_value(color)
            .map_err(|error| JsError::new(&format!("invalid color: {error}")))?;
        let wb = self.inner.borrow();
        Ok(wb
            .resolve_color(&color.into())
            .map(|(r, g, b)| format!("{r:02X}{g:02X}{b:02X}")))
    }
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(getter, js_name = chartsheets)]
    pub fn chartsheets(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        let sheets: Vec<WasmChartSheet> =
            wb.chartsheets().iter().map(WasmChartSheet::from).collect();
        to_js_value(&sheets)
    }

    #[wasm_bindgen(getter, js_name = chartsheetCount)]
    pub fn chartsheet_count(&self) -> Result<u32, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.chartsheet_count() as u32)
    }

    #[wasm_bindgen(getter, js_name = sheetOrder)]
    pub fn sheet_order(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        let slots: Vec<WasmSheetSlot> = wb.sheet_order().iter().map(WasmSheetSlot::from).collect();
        to_js_value(&slots)
    }

    #[wasm_bindgen(getter, js_name = totalSheetCount)]
    pub fn total_sheet_count(&self) -> Result<u32, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.total_sheet_count() as u32)
    }
}
