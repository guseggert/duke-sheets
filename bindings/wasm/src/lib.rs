//! WebAssembly bindings for duke-sheets
//!
//! This module provides wasm-bindgen-based WebAssembly bindings for the duke-sheets library,
//! allowing JavaScript/TypeScript code to create and manipulate spreadsheets in the browser.

use std::sync::{Arc, RwLock};
use wasm_bindgen::prelude::*;

use duke_sheets_core::{
    CellError, CellRange, CellValue as CoreCellValue, Workbook as CoreWorkbook,
};

// =============================================================================
// Error Conversion
// =============================================================================

fn to_js_error(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

fn cell_error_to_string(e: &CellError) -> &'static str {
    match e {
        CellError::Div0 => "#DIV/0!",
        CellError::Na => "#N/A",
        CellError::Name => "#NAME?",
        CellError::Null => "#NULL!",
        CellError::Num => "#NUM!",
        CellError::Ref => "#REF!",
        CellError::Value => "#VALUE!",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
    }
}

// =============================================================================
// CellValue - JavaScript wrapper for cell values
// =============================================================================

/// Represents a cell value in a spreadsheet.
#[wasm_bindgen]
pub struct CellValue {
    inner: CoreCellValue,
}

#[wasm_bindgen]
impl CellValue {
    #[wasm_bindgen(getter)]
    pub fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    #[wasm_bindgen(getter)]
    pub fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_formula(&self) -> bool {
        matches!(self.inner, CoreCellValue::Formula { .. })
    }

    #[wasm_bindgen(js_name = asNumber)]
    pub fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asText)]
    pub fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asBoolean)]
    pub fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asError)]
    pub fn as_error(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e).to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = formulaText)]
    pub fn formula_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Formula { text, .. } => Some(text.clone()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = toJs)]
    pub fn to_js(&self) -> JsValue {
        match &self.inner {
            CoreCellValue::Empty => JsValue::NULL,
            CoreCellValue::Number(n) => JsValue::from_f64(*n),
            CoreCellValue::String(s) => JsValue::from_str(&s.to_string()),
            CoreCellValue::Boolean(b) => JsValue::from_bool(*b),
            CoreCellValue::Error(e) => JsValue::from_str(cell_error_to_string(e)),
            CoreCellValue::Formula {
                cached_value: Some(v),
                ..
            } => {
                // Return the cached value
                match v.as_ref() {
                    CoreCellValue::Number(n) => JsValue::from_f64(*n),
                    CoreCellValue::String(s) => JsValue::from_str(&s.to_string()),
                    CoreCellValue::Boolean(b) => JsValue::from_bool(*b),
                    CoreCellValue::Error(e) => JsValue::from_str(cell_error_to_string(e)),
                    _ => JsValue::NULL,
                }
            }
            CoreCellValue::Formula { text, .. } => JsValue::from_str(text),
            CoreCellValue::SpillTarget { .. } => JsValue::NULL,
        }
    }

    #[wasm_bindgen(js_name = toString)]
    pub fn to_string_js(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => String::new(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            CoreCellValue::Formula { text, .. } => text.clone(),
            CoreCellValue::SpillTarget { .. } => String::new(),
        }
    }
}

// =============================================================================
// Worksheet - JavaScript wrapper
// =============================================================================

#[wasm_bindgen]
pub struct Worksheet {
    workbook: Arc<RwLock<CoreWorkbook>>,
    sheet_index: usize,
}

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> Result<String, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))
    }

    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(&self, address: &str, value: JsValue) -> Result<(), JsError> {
        let mut wb = self.workbook.write().map_err(to_js_error)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        let cell_value = js_to_cell_value(value)?;
        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;

        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setFormula)]
    pub fn set_formula(&self, address: &str, formula: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.write().map_err(to_js_error)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        ws.set_cell_formula(address, formula).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getCell)]
    pub fn get_cell(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;

        let value = ws.get_value_at(addr.row, addr.col);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = getCalculatedValue)]
    pub fn get_calculated_value(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        let addr = duke_sheets_core::CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;

        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);

        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = usedRange)]
    pub fn used_range(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.read().map_err(to_js_error)?;
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        match ws.used_range() {
            Some(range) => {
                let arr = js_sys::Array::new();
                arr.push(&JsValue::from(range.start.row));
                arr.push(&JsValue::from(range.start.col));
                arr.push(&JsValue::from(range.end.row));
                arr.push(&JsValue::from(range.end.col));
                Ok(arr.into())
            }
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = mergeCells)]
    pub fn merge_cells(&self, range_str: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.write().map_err(to_js_error)?;
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_js_error)
    }
}

// =============================================================================
// Workbook - JavaScript wrapper
// =============================================================================

#[wasm_bindgen]
pub struct Workbook {
    inner: Arc<RwLock<CoreWorkbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(RwLock::new(CoreWorkbook::new())),
        }
    }

    #[wasm_bindgen(getter, js_name = sheetCount)]
    pub fn sheet_count(&self) -> Result<usize, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        Ok(wb.sheet_count())
    }

    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Result<Vec<String>, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    #[wasm_bindgen(js_name = getSheet)]
    pub fn get_sheet(&self, index: usize) -> Result<Worksheet, JsError> {
        let wb = self.inner.read().map_err(to_js_error)?;
        if index >= wb.sheet_count() {
            return Err(JsError::new(&format!("Sheet index {} out of range", index)));
        }
        drop(wb);

        Ok(Worksheet {
            workbook: Arc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&self, name: &str) -> Result<usize, JsError> {
        let mut wb = self.inner.write().map_err(to_js_error)?;
        wb.add_worksheet_with_name(name).map_err(to_js_error)
    }

    /// Calculate all formulas in the workbook
    pub fn calculate(&self) -> Result<JsValue, JsError> {
        use duke_sheets_formula::evaluator::{evaluate, EvaluationContext};
        use duke_sheets_formula::parser::parse_formula;

        let mut wb = self.inner.write().map_err(to_js_error)?;
        let mut cells_calculated = 0;
        let mut errors = 0;

        // Collect all formula cells
        let sheet_count = wb.sheet_count();
        let mut formulas: Vec<(usize, u32, u16, String)> = Vec::new();

        for sheet_idx in 0..sheet_count {
            if let Some(ws) = wb.worksheet(sheet_idx) {
                if let Some(range) = ws.used_range() {
                    for row in range.start.row..=range.end.row {
                        for col in range.start.col..=range.end.col {
                            let val = ws.get_value_at(row, col);
                            if let CoreCellValue::Formula { text, .. } = val {
                                formulas.push((sheet_idx, row, col, text));
                            }
                        }
                    }
                }
            }
        }

        // Evaluate each formula
        for (sheet_idx, row, col, formula_text) in formulas {
            cells_calculated += 1;

            let result = parse_formula(&formula_text)
                .map_err(|e| format!("{:?}", e))
                .and_then(|ast| {
                    let ctx = EvaluationContext::new(Some(&*wb), sheet_idx, row, col);
                    evaluate(&ast, &ctx).map_err(|e| format!("{:?}", e))
                });

            if let Some(ws) = wb.worksheet_mut(sheet_idx) {
                let cached: CoreCellValue = match result {
                    Ok(val) => val.into(),
                    Err(_) => {
                        errors += 1;
                        CoreCellValue::Error(CellError::Calc)
                    }
                };

                let current = ws.get_value_at(row, col);
                if let CoreCellValue::Formula { text, .. } = current {
                    let _ = ws.set_cell_value_at(
                        row,
                        col,
                        CoreCellValue::Formula {
                            text,
                            cached_value: Some(Box::new(cached)),
                            array_result: None,
                        },
                    );
                }
            }
        }

        let stats = js_sys::Object::new();
        js_sys::Reflect::set(
            &stats,
            &"cellsCalculated".into(),
            &JsValue::from(cells_calculated),
        )
        .ok();
        js_sys::Reflect::set(&stats, &"errors".into(), &JsValue::from(errors)).ok();
        Ok(stats.into())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

// =============================================================================
// Helper functions
// =============================================================================

fn js_to_cell_value(value: JsValue) -> Result<CoreCellValue, JsError> {
    if value.is_null() || value.is_undefined() {
        Ok(CoreCellValue::Empty)
    } else if let Some(b) = value.as_bool() {
        Ok(CoreCellValue::Boolean(b))
    } else if let Some(n) = value.as_f64() {
        Ok(CoreCellValue::Number(n))
    } else if let Some(s) = value.as_string() {
        Ok(CoreCellValue::string(s))
    } else {
        Err(JsError::new(
            "Cell value must be null, boolean, number, or string",
        ))
    }
}

#[wasm_bindgen(start)]
pub fn init() {}
