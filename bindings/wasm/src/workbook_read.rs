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
