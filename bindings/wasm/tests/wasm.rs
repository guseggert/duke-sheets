//! WASM binding tests
//!
//! Run with: wasm-pack test --node

#![cfg(target_arch = "wasm32")]

use wasm_bindgen::JsValue;
use wasm_bindgen_test::*;

use duke_sheets_wasm::*;

wasm_bindgen_test_configure!(run_in_browser);

// =============================================================================
// Workbook Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_workbook_new() {
    let wb = Workbook::new();
    assert_eq!(wb.sheet_count().unwrap(), 1);
}

#[wasm_bindgen_test]
fn test_workbook_sheet_names() {
    let wb = Workbook::new();
    let names = wb.sheet_names().unwrap();
    assert_eq!(names.len(), 1);
    assert_eq!(names[0], "Sheet1");
}

#[wasm_bindgen_test]
fn test_workbook_add_sheet() {
    let wb = Workbook::new();
    let idx = wb.add_sheet("NewSheet").unwrap();

    assert_eq!(idx, 1);
    assert_eq!(wb.sheet_count().unwrap(), 2);

    let names = wb.sheet_names().unwrap();
    assert!(names.contains(&"NewSheet".to_string()));
}

#[wasm_bindgen_test]
fn test_workbook_remove_sheet() {
    let wb = Workbook::new();
    wb.add_sheet("ToRemove").unwrap();
    assert_eq!(wb.sheet_count().unwrap(), 2);

    wb.remove_sheet(1).unwrap();
    assert_eq!(wb.sheet_count().unwrap(), 1);
}

#[wasm_bindgen_test]
fn test_workbook_get_sheet_by_index() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    assert_eq!(sheet.name().unwrap(), "Sheet1");
}

#[wasm_bindgen_test]
fn test_workbook_get_sheet_by_name() {
    let wb = Workbook::new();
    wb.add_sheet("MySheet").unwrap();

    let sheet = wb.get_sheet_by_name("MySheet").unwrap();
    assert_eq!(sheet.name().unwrap(), "MySheet");
}

#[wasm_bindgen_test]
fn test_workbook_invalid_sheet_index() {
    let wb = Workbook::new();
    let result = wb.get_sheet(999);
    assert!(result.is_err());
}

// =============================================================================
// Worksheet Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_worksheet_set_get_number() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.0)).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_number);
    assert_eq!(value.as_number(), Some(42.0));
}

#[wasm_bindgen_test]
fn test_worksheet_set_get_text() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("Hello")).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_text);
    assert_eq!(value.as_text(), Some("Hello".to_string()));
}

#[wasm_bindgen_test]
fn test_worksheet_set_get_boolean() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_bool(true)).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_boolean);
    assert_eq!(value.as_boolean(), Some(true));
}

#[wasm_bindgen_test]
fn test_worksheet_set_null_clears() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.0)).unwrap();
    sheet.set_cell("A1", JsValue::NULL).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_empty);
}

#[wasm_bindgen_test]
fn test_worksheet_get_empty_cell() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let value = sheet.get_cell("Z99").unwrap();
    assert!(value.is_empty);
}

#[wasm_bindgen_test]
fn test_worksheet_used_range_empty() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let range = sheet.used_range().unwrap();
    assert!(range.is_null());
}

#[wasm_bindgen_test]
fn test_worksheet_used_range_with_data() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("B2", JsValue::from_f64(1.0)).unwrap();
    sheet.set_cell("D4", JsValue::from_f64(2.0)).unwrap();

    let range = sheet.used_range().unwrap();
    assert!(!range.is_null());
}

// =============================================================================
// Formula Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_formula_simple() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_formula("A1", "=1+1").unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_formula);
}

#[wasm_bindgen_test]
fn test_formula_cell_reference() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(10.0)).unwrap();
    sheet.set_cell("A2", JsValue::from_f64(20.0)).unwrap();
    sheet.set_formula("A3", "=A1+A2").unwrap();

    wb.calculate().unwrap();

    let value = sheet.get_calculated_value("A3").unwrap();
    assert_eq!(value.as_number(), Some(30.0));
}

#[wasm_bindgen_test]
fn test_formula_sum() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(1.0)).unwrap();
    sheet.set_cell("A2", JsValue::from_f64(2.0)).unwrap();
    sheet.set_cell("A3", JsValue::from_f64(3.0)).unwrap();
    sheet.set_formula("A4", "=SUM(A1:A3)").unwrap();

    wb.calculate().unwrap();

    let value = sheet.get_calculated_value("A4").unwrap();
    assert_eq!(value.as_number(), Some(6.0));
}

#[wasm_bindgen_test]
fn test_formula_nested() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(5.0)).unwrap();
    sheet.set_formula("A2", "=A1*2").unwrap(); // 10
    sheet.set_formula("A3", "=A2+A1").unwrap(); // 15

    wb.calculate().unwrap();

    assert_eq!(
        sheet.get_calculated_value("A2").unwrap().as_number(),
        Some(10.0)
    );
    assert_eq!(
        sheet.get_calculated_value("A3").unwrap().as_number(),
        Some(15.0)
    );
}

// =============================================================================
// Calculation Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_calculation_stats() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_formula("A1", "=1+1").unwrap();
    sheet.set_formula("A2", "=2+2").unwrap();

    let stats = wb.calculate().unwrap();

    assert_eq!(stats.formula_count, 2);
    assert!(stats.cells_calculated >= 2);
    assert_eq!(stats.errors, 0);
}

#[wasm_bindgen_test]
fn test_calculation_with_options() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_formula("A1", "=1+1").unwrap();

    let stats = wb.calculate_with_options(false, 100, 0.001).unwrap();

    assert_eq!(stats.formula_count, 1);
}

// =============================================================================
// Named Range Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_named_range_constant() {
    let wb = Workbook::new();

    wb.define_name("TaxRate", "0.1").unwrap();

    let result = wb.get_named_range("TaxRate").unwrap();
    assert_eq!(result, Some("0.1".to_string()));
}

#[wasm_bindgen_test]
fn test_named_range_undefined() {
    let wb = Workbook::new();

    let result = wb.get_named_range("NotDefined").unwrap();
    assert_eq!(result, None);
}

// =============================================================================
// CellValue Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_cell_value_to_js_number() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.5)).unwrap();
    let value = sheet.get_cell("A1").unwrap();

    let js = value.to_js();
    assert_eq!(js.as_f64(), Some(42.5));
}

#[wasm_bindgen_test]
fn test_cell_value_to_js_string() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("Hello")).unwrap();
    let value = sheet.get_cell("A1").unwrap();

    let js = value.to_js();
    assert_eq!(js.as_string(), Some("Hello".to_string()));
}

#[wasm_bindgen_test]
fn test_cell_value_to_js_boolean() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_bool(true)).unwrap();
    let value = sheet.get_cell("A1").unwrap();

    let js = value.to_js();
    assert_eq!(js.as_bool(), Some(true));
}

#[wasm_bindgen_test]
fn test_cell_value_to_js_null() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let value = sheet.get_cell("Z99").unwrap();

    let js = value.to_js();
    assert!(js.is_null());
}

#[wasm_bindgen_test]
fn test_cell_value_to_string() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.0)).unwrap();
    let value = sheet.get_cell("A1").unwrap();

    assert_eq!(value.to_string_js(), "42");
}

// =============================================================================
// CSV Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_csv_roundtrip() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(1.0)).unwrap();
    sheet.set_cell("B1", JsValue::from_f64(2.0)).unwrap();
    sheet.set_cell("A2", JsValue::from_f64(3.0)).unwrap();
    sheet.set_cell("B2", JsValue::from_f64(4.0)).unwrap();

    let csv = wb.save_csv_string().unwrap();
    assert!(csv.contains("1"));
    assert!(csv.contains("2"));

    let wb2 = Workbook::load_csv_string(&csv).unwrap();
    let sheet2 = wb2.get_sheet(0).unwrap();

    assert_eq!(sheet2.get_cell("A1").unwrap().as_number(), Some(1.0));
}

// =============================================================================
// Row/Column Dimension Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_row_height() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_row_height(0, 30.0).unwrap();
    // Setting succeeds (no getter to verify in WASM yet)
}

#[wasm_bindgen_test]
fn test_column_width() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_column_width(0, 15.0).unwrap();
    // Setting succeeds
}

// =============================================================================
// Merge Cell Tests
// =============================================================================

#[wasm_bindgen_test]
fn test_merge_cells() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("Merged")).unwrap();
    sheet.merge_cells("A1:C3").unwrap();
    // Merging succeeds
}

#[wasm_bindgen_test]
fn test_unmerge_cells() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.merge_cells("A1:C3").unwrap();
    sheet.unmerge_cells("A1:C3").unwrap();
    // Both operations succeed
}
