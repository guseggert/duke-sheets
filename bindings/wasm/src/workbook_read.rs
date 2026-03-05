use wasm_bindgen::prelude::*;

use crate::{
    to_js_error, to_js_value, types::WasmNamedRange, types::WasmWorkbookSettings, Workbook,
};

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(getter, js_name = isEmpty)]
    pub fn is_empty(&self) -> Result<bool, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        Ok(wb.is_empty())
    }

    #[wasm_bindgen(getter, js_name = activeSheet)]
    pub fn active_sheet(&self) -> Result<u32, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        Ok(wb.active_sheet() as u32)
    }

    #[wasm_bindgen(js_name = sheetIndex)]
    pub fn sheet_index(&self, name: &str) -> Result<Option<u32>, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        Ok(wb.sheet_index(name).map(|i| i as u32))
    }

    #[wasm_bindgen(getter, js_name = settings)]
    pub fn settings(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        to_js_value(&WasmWorkbookSettings::from(wb.settings()))
    }

    #[wasm_bindgen(getter, js_name = namedRanges)]
    pub fn named_ranges(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        let ranges: Vec<WasmNamedRange> =
            wb.named_ranges().iter().map(WasmNamedRange::from).collect();
        to_js_value(&ranges)
    }
}
