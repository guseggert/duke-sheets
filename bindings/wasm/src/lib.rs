use std::cell::RefCell;
use std::io::Cursor;
use std::rc::Rc;
use std::sync::Arc;

use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;

use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue as CoreCellValue, Workbook as CoreWorkbook,
};
use duke_sheets_xlsx::XlsxWriter;

mod types;
mod workbook_read;
mod worksheet_read;

pub use types::*;

#[wasm_bindgen(typescript_custom_section)]
const ROW_ITERATOR_TYPES: &str = r#"
declare global {
  interface SymbolConstructor {
    readonly dispose: unique symbol;
  }
}

export interface JsRowCell {
  col: number;
  value: string;
  style?: any;
  mergeSpan?: JsMergeSpan;
  isMergedSecondary?: boolean;
  hyperlink?: any;
  comment?: any;
  formula?: string;
  image?: any;
}

export interface JsRow {
  index: number;
  cells: Array<JsRowCell>;
}

export interface JsMergeSpan {
  rowSpan: number;
  colSpan: number;
}

export interface JsRowsOptions {
  useFormattedValues?: boolean;
  useCalculatedValues?: boolean;
  includeStyles?: boolean;
  includeMergeInfo?: boolean;
  includeHyperlinks?: boolean;
  includeComments?: boolean;
  includeFormulas?: boolean;
  includeImages?: boolean;
  skipEmptyValues?: boolean;
  skipBlankValues?: boolean;
}

export class RowIterator implements IterableIterator<JsRow> {
  constructor(ws: Worksheet, opts?: JsRowsOptions, maxRow?: number);
  [Symbol.iterator](): RowIterator;
  next(): IteratorResult<JsRow>;
}

export interface Worksheet {
  iterateRows(opts?: JsRowsOptions): RowIterator;
}
"#;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsCalculationOptions {
    iterative: Option<bool>,
    max_iterations: Option<u32>,
    max_change: Option<f64>,
    force_full_calculation: Option<bool>,
    calculate_volatile: Option<bool>,
    sheets: Option<Vec<usize>>,
    max_threads: Option<usize>,
}

/// Wrapper to make `js_sys::Function` implement Send + Sync.
/// SAFETY: WASM is single-threaded, so Send + Sync are no-ops.
struct SendSyncFunction(js_sys::Function);
unsafe impl Send for SendSyncFunction {}
unsafe impl Sync for SendSyncFunction {}

impl SendSyncFunction {
    fn call1(&self, this: &JsValue, arg: &JsValue) -> Result<JsValue, JsValue> {
        self.0.call1(this, arg)
    }

    fn call3(
        &self,
        this: &JsValue,
        arg1: &JsValue,
        arg2: &JsValue,
        arg3: &JsValue,
    ) -> Result<JsValue, JsValue> {
        self.0.call3(this, arg1, arg2, arg3)
    }
}

pub(crate) fn to_js_error(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

pub(crate) fn to_js_value<T: Serialize>(value: &T) -> Result<JsValue, JsError> {
    serde_wasm_bindgen::to_value(value).map_err(to_js_error)
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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmCalculationStats {
    formula_count: u32,
    cells_calculated: u32,
    errors: u32,
    circular_references: u32,
    volatile_cells: u32,
    converged: bool,
    iterations: u32,
}

impl From<&duke_sheets::CalculationStats> for WasmCalculationStats {
    fn from(stats: &duke_sheets::CalculationStats) -> Self {
        Self {
            formula_count: stats.formula_count as u32,
            cells_calculated: stats.cells_calculated as u32,
            errors: stats.errors as u32,
            circular_references: stats.circular_references as u32,
            volatile_cells: stats.volatile_cells as u32,
            converged: stats.converged,
            iterations: stats.iterations as u32,
        }
    }
}

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

    #[wasm_bindgen(js_name = toJs)]
    pub fn to_js(&self) -> JsValue {
        match &self.inner {
            CoreCellValue::Empty => JsValue::NULL,
            CoreCellValue::Number(n) => JsValue::from_f64(*n),
            CoreCellValue::String(s) => JsValue::from_str(&s.to_string()),
            CoreCellValue::Boolean(b) => JsValue::from_bool(*b),
            CoreCellValue::Error(e) => JsValue::from_str(cell_error_to_string(e)),
            CoreCellValue::RichText(runs) => {
                let text: String = runs.iter().map(|r| r.text.as_str()).collect();
                JsValue::from_str(&text)
            }
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
            CoreCellValue::RichText(runs) => runs.iter().map(|r| r.text.as_str()).collect(),
            CoreCellValue::SpillTarget { .. } => String::new(),
        }
    }
}

#[wasm_bindgen]
pub struct Worksheet {
    workbook: Rc<RefCell<CoreWorkbook>>,
    sheet_index: usize,
}

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))
    }

    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(&self, address: &str, value: JsValue) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let cell_value = js_to_cell_value(value)?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setFormula)]
    pub fn set_formula(&self, address: &str, formula: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_cell_formula(address, formula).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getCell)]
    pub fn get_cell(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        Ok(CellValue {
            inner: ws.get_value_at(addr.row, addr.col),
        })
    }

    #[wasm_bindgen(js_name = getCellAt)]
    pub fn get_cell_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(CellValue {
            inner: ws.get_value_at(row, col as u16),
        })
    }

    #[wasm_bindgen(js_name = getCalculatedValue)]
    pub fn get_calculated_value(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = getCalculatedValueAt)]
    pub fn get_calculated_value_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let value = ws
            .get_calculated_value_at(row, col as u16)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = usedRange)]
    pub fn used_range(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
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

    #[wasm_bindgen(js_name = setRowHeight)]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_row_height(row, height);
        Ok(())
    }

    #[wasm_bindgen(js_name = setColumnWidth)]
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_column_width(col as u16, width);
        Ok(())
    }

    #[wasm_bindgen(js_name = getRowHeight)]
    pub fn get_row_height(&self, row: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_row_heights().get(&row).copied())
    }

    #[wasm_bindgen(js_name = getColumnWidth)]
    pub fn get_column_width(&self, col: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_column_widths().get(&(col as u16)).copied())
    }

    #[wasm_bindgen(js_name = mergeCells)]
    pub fn merge_cells(&self, range_str: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = unmergeCells)]
    pub fn unmerge_cells(&self, range_str: &str) -> Result<bool, JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        Ok(ws.unmerge_cells(&range))
    }

    #[wasm_bindgen(js_name = getImageAt)]
    pub fn get_image_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        match ws.get_image_at(row, col as u16) {
            Some(info) => to_js_value(&WasmImageInfo::from(info)),
            None => Ok(JsValue::NULL),
        }
    }
}

#[wasm_bindgen]
pub struct Workbook {
    pub(crate) inner: Rc<RefCell<CoreWorkbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Rc::new(RefCell::new(CoreWorkbook::new())),
        }
    }

    #[wasm_bindgen(js_name = loadCsvString)]
    pub fn load_csv_string(csv: &str) -> Result<Workbook, JsError> {
        let reader = Cursor::new(csv.as_bytes());
        let ws =
            duke_sheets_csv::CsvReader::read(reader, &duke_sheets_csv::CsvReadOptions::default())
                .map_err(to_js_error)?;
        let mut wb = CoreWorkbook::empty();
        wb.add_existing_worksheet(ws).map_err(to_js_error)?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    #[wasm_bindgen(js_name = fromCsvString)]
    pub fn from_csv_string(csv: &str) -> Result<Workbook, JsError> {
        Self::load_csv_string(csv)
    }

    /// Load a workbook from bytes, auto-detecting the format (XLSX or XLS).
    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8]) -> Result<Workbook, JsError> {
        use duke_sheets::WorkbookExt;
        let wb = duke_sheets_core::Workbook::from_bytes(data)
            .map_err(|e| JsError::new(&format!("Failed to read file: {}", e)))?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    #[wasm_bindgen(js_name = saveXlsxBytes)]
    pub fn save_xlsx_bytes(&self) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).map_err(to_js_error)?;
        Ok(buf)
    }

    #[wasm_bindgen(js_name = saveCsvString)]
    pub fn save_csv_string(&self) -> Result<String, JsError> {
        let wb = self.inner.borrow();
        let ws = wb
            .worksheet(0)
            .ok_or_else(|| JsError::new("No worksheets to save"))?;
        let mut buf = Vec::new();
        duke_sheets_csv::CsvWriter::write(
            ws,
            &mut buf,
            &duke_sheets_csv::CsvWriteOptions::default(),
        )
        .map_err(to_js_error)?;
        String::from_utf8(buf).map_err(to_js_error)
    }

    #[wasm_bindgen(getter, js_name = sheetCount)]
    pub fn sheet_count(&self) -> Result<usize, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.sheet_count())
    }

    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Result<Vec<String>, JsError> {
        let wb = self.inner.borrow();
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    #[wasm_bindgen(js_name = getSheet)]
    pub fn get_sheet(&self, index: usize) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        if index >= wb.sheet_count() {
            return Err(JsError::new(&format!("Sheet index {} out of range", index)));
        }
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = getSheetByName)]
    pub fn get_sheet_by_name(&self, name: &str) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        let index = wb
            .sheet_index(name)
            .ok_or_else(|| JsError::new(&format!("Sheet '{}' not found", name)))?;
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&self, name: &str) -> Result<usize, JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.add_worksheet_with_name(name).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = removeSheet)]
    pub fn remove_sheet(&self, index: usize) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.remove_worksheet(index).map(|_| ()).map_err(to_js_error)
    }

    pub fn calculate(&self, options: Option<JsValue>) -> Result<JsValue, JsError> {
        let mut wb = self.inner.borrow_mut();
        let options = if let Some(options) = options {
            // Extract JS callback functions before serde deserialization (which consumes the JsValue).
            // Serde can't deserialize JS functions, so we pull them out via Reflect::get.
            let web_service_js_fn =
                js_sys::Reflect::get(&options, &JsValue::from_str("webServiceFn"))
                    .ok()
                    .and_then(|v| v.dyn_into::<js_sys::Function>().ok());
            let rtd_js_fn = js_sys::Reflect::get(&options, &JsValue::from_str("rtdFn"))
                .ok()
                .and_then(|v| v.dyn_into::<js_sys::Function>().ok());

            let js_opts: JsCalculationOptions =
                serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;

            // Build web_service_fn from callback
            let web_service_fn = web_service_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(move |url: &str| -> Option<String> {
                    let result = wrapper
                        .call1(&JsValue::NULL, &JsValue::from_str(url))
                        .ok()?;
                    result.as_string()
                }) as Arc<dyn Fn(&str) -> Option<String> + Send + Sync>
            });

            // Build rtd_fn from callback
            let rtd_fn = rtd_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(
                    move |prog_id: &str, server: &str, topics: &[String]| -> Option<String> {
                        let topics_arr = js_sys::Array::new();
                        for t in topics {
                            topics_arr.push(&JsValue::from_str(t));
                        }
                        let result = wrapper
                            .call3(
                                &JsValue::NULL,
                                &JsValue::from_str(prog_id),
                                &JsValue::from_str(server),
                                &topics_arr.into(),
                            )
                            .ok()?;
                        result.as_string()
                    },
                )
                    as Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>
            });

            CalculationOptions {
                iterative: js_opts.iterative.unwrap_or(false),
                max_iterations: js_opts.max_iterations.unwrap_or(100),
                max_change: js_opts.max_change.unwrap_or(0.001),
                force_full_calculation: js_opts.force_full_calculation.unwrap_or(true),
                calculate_volatile: js_opts.calculate_volatile.unwrap_or(true),
                sheets: js_opts.sheets.unwrap_or_default(),
                max_threads: js_opts.max_threads,
                web_service_fn,
                rtd_fn,
            }
        } else {
            CalculationOptions::default()
        };
        let stats = wb.calculate_with_options(&options).map_err(to_js_error)?;
        to_js_value(&WasmCalculationStats::from(&stats))
    }

    #[wasm_bindgen(js_name = defineName)]
    pub fn define_name(&self, name: &str, refers_to: &str) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.define_name(name, refers_to).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getNamedRange)]
    pub fn get_named_range(&self, name: &str) -> Result<Option<String>, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.get_named_range(name, 0).map(|nr| nr.refers_to.clone()))
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

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
