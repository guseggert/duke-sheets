use wasm_bindgen::prelude::*;

use crate::{
    to_js_value,
    types::{WasmFormControl, WasmFormControlInput},
    Worksheet,
};

#[wasm_bindgen]
impl Worksheet {
    /// All form controls in worksheet order.
    #[wasm_bindgen(getter, skip_typescript)]
    pub fn form_controls(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let controls: Vec<WasmFormControl> =
            ws.form_controls().iter().map(WasmFormControl::from).collect();
        to_js_value(&controls)
    }

    /// Number of form controls in the worksheet.
    #[wasm_bindgen(getter)]
    pub fn form_control_count(&self) -> Result<u32, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        u32::try_from(ws.form_control_count())
            .map_err(|_| JsError::new("form control count exceeds u32"))
    }

    /// Append a form control and return its zero-based index.
    #[wasm_bindgen(skip_typescript)]
    pub fn add_form_control(&self, value: JsValue) -> Result<u32, JsError> {
        let input: WasmFormControlInput = serde_wasm_bindgen::from_value(value)
            .map_err(|err| JsError::new(&format!("invalid form control: {err}")))?;
        let control = duke_sheets_core::FormControl::try_from(input).map_err(|err| JsError::new(&err))?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let index = ws
            .try_add_form_control(control)
            .map_err(|err| JsError::new(&err.to_string()))?;
        u32::try_from(index).map_err(|_| JsError::new("form control index exceeds u32"))
    }

    /// Replace a form control by zero-based index.
    #[wasm_bindgen(skip_typescript)]
    pub fn set_form_control(&self, index: u32, value: JsValue) -> Result<(), JsError> {
        let input: WasmFormControlInput = serde_wasm_bindgen::from_value(value)
            .map_err(|err| JsError::new(&format!("invalid form control: {err}")))?;
        let control = duke_sheets_core::FormControl::try_from(input).map_err(|err| JsError::new(&err))?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_form_control(index as usize, control)
            .map_err(|err| JsError::new(&err.to_string()))
    }

    /// Remove a form control by zero-based index.
    #[wasm_bindgen(skip_typescript)]
    pub fn remove_form_control(&self, index: u32) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.remove_form_control(index as usize)
            .map(|_| ())
            .map_err(|err| JsError::new(&err.to_string()))
    }
}
