//! WASM binding tests
//!
//! Run with: wasm-pack test --node

#![cfg(target_arch = "wasm32")]

use js_sys::{Array, Function, Object, Reflect};
use wasm_bindgen::JsValue;
use wasm_bindgen_test::*;

use duke_sheets_wasm::*;

// Tests run in Node.js via wasm-pack test --node

// Helper to get a numeric field from a JsValue object
fn get_f64_field(obj: &JsValue, key: &str) -> f64 {
    Reflect::get(obj, &JsValue::from_str(key))
        .unwrap()
        .as_f64()
        .unwrap()
}

fn get_string_field(obj: &JsValue, key: &str) -> String {
    Reflect::get(obj, &JsValue::from_str(key))
        .unwrap()
        .as_string()
        .unwrap()
}

fn get_bool_field(obj: &JsValue, key: &str) -> bool {
    Reflect::get(obj, &JsValue::from_str(key))
        .unwrap()
        .as_bool()
        .unwrap()
}

// Workbook Tests

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

// Worksheet Tests

#[wasm_bindgen_test]
fn test_worksheet_set_get_number() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.0)).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_number());
    assert_eq!(value.as_number(), Some(42.0));
}

#[wasm_bindgen_test]
fn test_worksheet_set_get_text() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("Hello")).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_text());
    assert_eq!(value.as_text(), Some("Hello".to_string()));
}

#[wasm_bindgen_test]
fn test_worksheet_set_get_boolean() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_bool(true)).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_boolean());
    assert_eq!(value.as_boolean(), Some(true));
}

#[wasm_bindgen_test]
fn test_worksheet_set_null_clears() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(42.0)).unwrap();
    sheet.set_cell("A1", JsValue::NULL).unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_empty());
}

#[wasm_bindgen_test]
fn test_worksheet_get_empty_cell() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let value = sheet.get_cell("Z99").unwrap();
    assert!(value.is_empty());
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

#[wasm_bindgen_test]
fn test_worksheet_set_cell_style() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let font_color = make_options(&[("hex", JsValue::from_str("FFFFFF"))]);
    let font = make_options(&[
        ("name", JsValue::from_str("Aptos Display")),
        ("size", JsValue::from_f64(14.0)),
        ("bold", JsValue::TRUE),
        ("color", font_color),
    ]);
    let fill_color = make_options(&[("hex", JsValue::from_str("1F4E79"))]);
    let fill = make_options(&[
        ("fillType", JsValue::from_str("solid")),
        ("color", fill_color),
    ]);
    let style = make_options(&[("font", font), ("fill", fill)]);

    sheet.set_cell_style("A1", style).unwrap();

    let style = sheet.get_cell_style("A1").unwrap();
    let font = Reflect::get(&style, &JsValue::from_str("font")).unwrap();
    let fill = Reflect::get(&style, &JsValue::from_str("fill")).unwrap();
    let font_color = Reflect::get(&font, &JsValue::from_str("color")).unwrap();
    let fill_color = Reflect::get(&fill, &JsValue::from_str("color")).unwrap();

    assert_eq!(get_string_field(&font, "name"), "Aptos Display");
    assert_eq!(get_f64_field(&font, "size"), 14.0);
    assert!(get_bool_field(&font, "bold"));
    assert_eq!(get_string_field(&font_color, "hex"), "FFFFFF");
    assert_eq!(get_string_field(&fill, "fillType"), "solid");
    assert_eq!(get_string_field(&fill_color, "hex"), "1F4E79");
}

#[wasm_bindgen_test]
fn test_worksheet_set_range_style() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let font = make_options(&[("italic", JsValue::TRUE)]);
    let style = make_options(&[("font", font)]);
    sheet.set_range_style("C1:D2", style).unwrap();

    for address in ["C1", "D1", "C2", "D2"] {
        let style = sheet.get_cell_style(address).unwrap();
        let font = Reflect::get(&style, &JsValue::from_str("font")).unwrap();
        assert!(get_bool_field(&font, "italic"));
    }
}

// Formula Tests

#[wasm_bindgen_test]
fn test_formula_simple() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_formula("A1", "=1+1").unwrap();

    let value = sheet.get_cell("A1").unwrap();
    assert!(value.is_empty());
    assert_eq!(
        sheet.get_formula_at(0, 0).unwrap(),
        Some("=1+1".to_string())
    );

    wb.calculate(None).unwrap();
    let calculated = sheet.get_calculated_value("A1").unwrap();
    assert_eq!(calculated.as_number(), Some(2.0));
}

#[wasm_bindgen_test]
fn test_formula_cell_reference() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(10.0)).unwrap();
    sheet.set_cell("A2", JsValue::from_f64(20.0)).unwrap();
    sheet.set_formula("A3", "=A1+A2").unwrap();

    wb.calculate(None).unwrap();

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

    wb.calculate(None).unwrap();

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

    wb.calculate(None).unwrap();

    assert_eq!(
        sheet.get_calculated_value("A2").unwrap().as_number(),
        Some(10.0)
    );
    assert_eq!(
        sheet.get_calculated_value("A3").unwrap().as_number(),
        Some(15.0)
    );
}

// Calculation Tests

#[wasm_bindgen_test]
fn test_calculation_stats() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_formula("A1", "=1+1").unwrap();
    sheet.set_formula("A2", "=2+2").unwrap();

    let stats = wb.calculate(None).unwrap();

    assert_eq!(get_f64_field(&stats, "formulaCount") as u32, 2);
    assert!(get_f64_field(&stats, "cellsCalculated") as u32 >= 2);
    assert_eq!(get_f64_field(&stats, "errors") as u32, 0);
}

/// Helper to build a JS options object from key-value pairs
fn make_options(entries: &[(&str, JsValue)]) -> JsValue {
    let obj = Object::new();
    for (key, val) in entries {
        Reflect::set(&obj, &JsValue::from_str(key), val).unwrap();
    }
    obj.into()
}

fn control_anchor() -> JsValue {
    make_options(&[
        ("fromCol", JsValue::from_f64(1.0)),
        ("fromRow", JsValue::from_f64(1.0)),
        ("fromColOffset", JsValue::from_f64(0.0)),
        ("fromRowOffset", JsValue::from_f64(0.0)),
        ("toCol", JsValue::from_f64(3.0)),
        ("toRow", JsValue::from_f64(2.0)),
        ("toColOffset", JsValue::from_f64(0.0)),
        ("toRowOffset", JsValue::from_f64(0.0)),
        ("editAs", JsValue::from_str("twoCell")),
    ])
}

fn form_control(kind: JsValue) -> JsValue {
    make_options(&[("anchor", control_anchor()), ("kind", kind)])
}

#[wasm_bindgen_test]
fn test_form_control_mutations_and_roundtrip() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    // Zero-based: first and third items.
    let selected = Array::new();
    selected.push(&JsValue::from_f64(0.0));
    selected.push(&JsValue::from_f64(2.0));
    let kinds = vec![
        make_options(&[("kind", JsValue::from_str("button")), ("caption", JsValue::from_str("Run"))]),
        make_options(&[("kind", JsValue::from_str("checkbox")), ("caption", JsValue::from_str("Check")), ("state", JsValue::from_str("checked")), ("no3D", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("optionButton")), ("caption", JsValue::from_str("Option")), ("state", JsValue::from_str("unchecked")), ("no3D", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("label")), ("caption", JsValue::from_str("Label"))]),
        make_options(&[("kind", JsValue::from_str("groupBox")), ("caption", JsValue::from_str("Group")), ("no3D", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("listBox")), ("inputRange", JsValue::from_str("$A$1:$A$3")), ("selection", JsValue::from_str("multi")), ("selected", selected.into()), ("no3D", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("dropdown")), ("inputRange", JsValue::from_str("$A$1:$A$3")), ("selected", JsValue::from_f64(2.0)), ("lines", JsValue::from_f64(8.0)), ("no3D", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("scrollbar")), ("value", JsValue::from_f64(5.0)), ("min", JsValue::from_f64(0.0)), ("max", JsValue::from_f64(10.0)), ("increment", JsValue::from_f64(1.0)), ("page", JsValue::from_f64(2.0)), ("horizontal", JsValue::FALSE)]),
        make_options(&[("kind", JsValue::from_str("spinner")), ("value", JsValue::from_f64(2.0)), ("min", JsValue::from_f64(0.0)), ("max", JsValue::from_f64(10.0)), ("increment", JsValue::from_f64(1.0))]),
    ];
    for kind in kinds {
        sheet.add_form_control(form_control(kind)).unwrap();
    }
    assert_eq!(sheet.form_control_count().unwrap(), 9);
    let controls = Array::from(&sheet.form_controls().unwrap());
    assert_eq!(controls.length(), 9);

    sheet
        .set_form_control(
            0,
            form_control(make_options(&[("kind", JsValue::from_str("label")), ("caption", JsValue::from_str("Replaced"))])),
        )
        .unwrap();
    sheet.remove_form_control(0).unwrap();
    assert_eq!(sheet.form_control_count().unwrap(), 8);

    for bytes in [wb.save_xlsx_bytes().unwrap(), wb.save_xlsb_bytes().unwrap()] {
        let reopened = Workbook::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.get_sheet(0).unwrap().form_control_count().unwrap(), 8);
    }
    let xls = wb
        .save_xls_bytes_encrypted("password", None, None)
        .unwrap();
    let reopened = Workbook::from_bytes_with_password(&xls, "password", None).unwrap();
    assert_eq!(reopened.get_sheet(0).unwrap().form_control_count().unwrap(), 8);
}

/// The form-control API must be reachable through its documented
/// camelCase JS names. Rust method calls bypass JS property lookup,
/// so a missing `js_name` attribute is only caught here.
#[wasm_bindgen_test]
fn test_form_control_js_api_names() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_form_control(form_control(make_options(&[
            ("kind", JsValue::from_str("button")),
            ("caption", JsValue::from_str("Run")),
        ])))
        .unwrap();

    let sheet_js = JsValue::from(sheet);
    let count = Reflect::get(&sheet_js, &"formControlCount".into()).unwrap();
    assert_eq!(count.as_f64(), Some(1.0), "formControlCount getter");
    let controls = Reflect::get(&sheet_js, &"formControls".into()).unwrap();
    assert!(Array::is_array(&controls), "formControls getter");
    let kind = Reflect::get(&Array::from(&controls).get(0), &"kind".into()).unwrap();
    let kind_tag = Reflect::get(&kind, &"kind".into()).unwrap();
    assert_eq!(kind_tag.as_string().as_deref(), Some("button"));
    for name in ["addFormControl", "setFormControl", "removeFormControl"] {
        let method = Reflect::get(&sheet_js, &(*name).into()).unwrap();
        assert!(method.is_function(), "{name} missing from the JS API");
    }
}

#[wasm_bindgen_test]
fn test_calculation_with_options() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_formula("A1", "=1+1").unwrap();

    let opts = make_options(&[
        ("iterative", JsValue::from(false)),
        ("maxIterations", JsValue::from(100)),
        ("maxChange", JsValue::from(0.001)),
    ]);
    let stats = wb.calculate(Some(opts)).unwrap();
    assert_eq!(get_f64_field(&stats, "formulaCount") as u32, 1);
    assert_eq!(
        sheet
            .get_calculated_value("A1")
            .unwrap()
            .as_number()
            .unwrap(),
        2.0
    );
}

#[wasm_bindgen_test]
fn test_calculation_with_empty_options() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_formula("A1", "=1+1").unwrap();

    let opts = make_options(&[]);
    let stats = wb.calculate(Some(opts)).unwrap();
    assert_eq!(get_f64_field(&stats, "formulaCount") as u32, 1);
    assert_eq!(
        sheet
            .get_calculated_value("A1")
            .unwrap()
            .as_number()
            .unwrap(),
        2.0
    );
}

#[wasm_bindgen_test]
fn test_calculation_with_max_threads() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_formula("A1", "=1+1").unwrap();

    let opts = make_options(&[("maxThreads", JsValue::from(1))]);
    let stats = wb.calculate(Some(opts)).unwrap();
    assert_eq!(get_f64_field(&stats, "formulaCount") as u32, 1);
    assert_eq!(
        sheet
            .get_calculated_value("A1")
            .unwrap()
            .as_number()
            .unwrap(),
        2.0
    );
}

#[wasm_bindgen_test]
fn test_calculation_image_metadata() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .set_formula(
            "A1",
            r#"=IMAGE("https://example.com/logo.png","Logo",3,48,96)"#,
        )
        .unwrap();
    wb.calculate(None).unwrap();
    assert_eq!(
        sheet.get_calculated_value("A1").unwrap().as_text().unwrap(),
        "Logo"
    );
    let image = sheet.get_image_at(0, 0).unwrap();
    assert!(!image.is_null());
    assert_eq!(
        get_string_field(&image, "source"),
        "https://example.com/logo.png"
    );
    assert_eq!(get_string_field(&image, "altText"), "Logo");
    assert_eq!(get_f64_field(&image, "sizing") as u32, 3);
    assert_eq!(get_f64_field(&image, "width"), 96.0);
    assert_eq!(get_f64_field(&image, "height"), 48.0);
    // Non-image cell returns null
    let no_image = sheet.get_image_at(1, 0).unwrap();
    assert!(no_image.is_null());
}

#[wasm_bindgen_test]
fn test_worksheet_formula_count() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    assert_eq!(sheet.formula_count().unwrap(), 0);
    sheet.set_formula("A1", "=1+1").unwrap();
    sheet.set_formula("B1", "=2+2").unwrap();
    sheet.set_cell("C1", JsValue::from_f64(42.0)).unwrap();
    assert_eq!(sheet.formula_count().unwrap(), 2);
}

// Named Range Tests

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

// CellValue Tests

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

// CSV Tests

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

    // CSV reader parses values as text; verify the text content roundtrips correctly
    let val = sheet2.get_cell("A1").unwrap();
    assert!(val.as_number() == Some(1.0) || val.as_text() == Some("1".to_string()));
}

// Row/Column Dimension Tests

#[wasm_bindgen_test]
fn test_row_height() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_row_height(0, 30.0).unwrap();
    assert_eq!(sheet.get_row_height(0).unwrap(), Some(30.0));
}

#[wasm_bindgen_test]
fn test_column_width() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_column_width(0, 15.0).unwrap();
    assert_eq!(sheet.get_column_width(0).unwrap(), Some(15.0));
}

// Merge Cell Tests

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

#[wasm_bindgen_test]
fn test_save_xlsx_bytes() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from(10.0)).unwrap();
    sheet.set_cell("A2", JsValue::from(20.0)).unwrap();
    sheet.set_formula("A3", "=A1+A2").unwrap();
    wb.calculate(None).unwrap();

    let bytes = wb.save_xlsx_bytes().unwrap();
    assert!(!bytes.is_empty(), "XLSX bytes should not be empty");

    // Roundtrip: load back and verify
    let wb2 = Workbook::from_bytes(&bytes).unwrap();
    let sheet2 = wb2.get_sheet(0).unwrap();
    let val = sheet2.get_cell("A1").unwrap();
    assert_eq!(val.as_number(), Some(10.0));
}

#[wasm_bindgen_test]
fn test_populated_feature_reads_through_binding() {
    // Comments, autofilters, and data validations are read-only in the
    // binding; build the fixture in-test with the core crates (path
    // deps of this package), write XLSX bytes, and read it back through
    // the binding's DTO conversions.
    use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, ValueFilter};
    use duke_sheets_core::comment::CellComment;
    use duke_sheets_core::{CellAddress, CellRange};
    use duke_sheets_core::validation::DataValidation;

    let mut core_wb = duke_sheets_core::Workbook::new();
    let ws = core_wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 1.0).unwrap();
    ws.set_comment(
        "A1",
        CellComment {
            author: "Tester".to_string(),
            text: "fixture comment".to_string(),
            visible: false,
        },
    )
    .unwrap();
    let mut af = AutoFilter::new(CellRange::new(
        CellAddress::parse("A1").unwrap(),
        CellAddress::parse("A2").unwrap(),
    ));
    af.filter_columns.push(FilterColumn::new(
        0,
        ColumnFilter::Values(ValueFilter {
            values: vec!["1".to_string()],
            blank: false,
        }),
    ));
    ws.set_auto_filter(Some(af));
    let mut dv = DataValidation::list("Red,Green,Blue");
    dv.ranges = vec![CellRange::parse("C1:C5").unwrap()];
    ws.add_data_validation(dv);

    let mut bytes = std::io::Cursor::new(Vec::new());
    duke_sheets_xlsx::XlsxWriter::write(&core_wb, &mut bytes).unwrap();

    let wb = Workbook::from_bytes(&bytes.into_inner()).unwrap();
    let sheet = wb.get_sheet(0).unwrap();

    assert_eq!(sheet.comment_count().unwrap(), 1);
    let comment = sheet.get_comment("A1").unwrap();
    assert_eq!(get_string_field(&comment, "author"), "Tester");
    assert_eq!(get_string_field(&comment, "text"), "fixture comment");

    let af = sheet.auto_filter().unwrap();
    assert!(!af.is_null(), "autofilter lost through the binding");
    assert_eq!(get_string_field(&af, "range"), "A1:A2");

    let dvs = sheet.data_validations().unwrap();
    let dvs_arr = js_sys::Array::from(&dvs);
    assert_eq!(dvs_arr.length(), 1);
    let dv0 = dvs_arr.get(0);
    assert_eq!(get_string_field(&dv0, "validationType"), "list");
    assert_eq!(get_string_field(&dv0, "listSource"), "Red,Green,Blue");
}

#[wasm_bindgen_test]
fn test_save_xlsb_bytes_roundtrip() {
    // XLSB carries binary formula token streams (BIFF12), so formula
    // text survival exercises the compiler and reader end to end.
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from(1.0)).unwrap();
    sheet.set_cell("A2", JsValue::from(2.0)).unwrap();
    sheet.set_cell("A3", JsValue::from(3.0)).unwrap();
    sheet.set_cell("B1", JsValue::from("label")).unwrap();
    sheet.set_formula("C1", "=SUM(A1:A3)").unwrap();
    sheet.set_formula("C2", "=IF(A1>0,A2,A3)").unwrap();
    wb.calculate(None).unwrap();

    let bytes = wb.save_xlsb_bytes().unwrap();
    assert!(!bytes.is_empty(), "XLSB bytes should not be empty");

    let wb2 = Workbook::from_bytes(&bytes).unwrap();
    let sheet2 = wb2.get_sheet(0).unwrap();
    assert_eq!(sheet2.get_cell("A1").unwrap().as_number(), Some(1.0));
    assert_eq!(
        sheet2.get_cell("B1").unwrap().as_text().as_deref(),
        Some("label")
    );
    assert_eq!(
        sheet2.get_formula_at(0, 2).unwrap().as_deref(),
        Some("=SUM(A1:A3)"),
        "XLSB formula text must survive the byte round-trip"
    );
    assert_eq!(
        sheet2.get_formula_at(1, 2).unwrap().as_deref(),
        Some("=IF(A1>0,A2,A3)")
    );
    assert_eq!(sheet2.get_cell("C1").unwrap().as_number(), Some(6.0));
}

#[wasm_bindgen_test]
fn test_now_and_today_formulas() {
    // NOW() and TODAY() use Local::now() which requires chrono's wasmbind feature
    // on wasm32 targets, otherwise SystemTime::now() panics.
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_formula("A1", "=NOW()").unwrap();
    sheet.set_formula("A2", "=TODAY()").unwrap();
    wb.calculate(None).unwrap();

    let now_val = sheet.get_calculated_value("A1").unwrap();
    let today_val = sheet.get_calculated_value("A2").unwrap();

    // NOW() returns a serial number > 0 (date + time fraction)
    let now_num = now_val.as_number().expect("NOW() should return a number");
    assert!(
        now_num > 0.0,
        "NOW() serial should be positive, got {}",
        now_num
    );

    // TODAY() returns an integer serial number > 0
    let today_num = today_val
        .as_number()
        .expect("TODAY() should return a number");
    assert!(
        today_num > 0.0,
        "TODAY() serial should be positive, got {}",
        today_num
    );

    // TODAY() should be the integer part of NOW()
    assert_eq!(today_num.floor(), today_num, "TODAY() should be an integer");
    assert_eq!(
        now_num.floor(),
        today_num,
        "NOW() date part should equal TODAY()"
    );
}

// JS callback function tests

#[wasm_bindgen_test]
fn test_calculation_with_web_service_fn() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .set_formula("A1", r#"=WEBSERVICE("https://example.com/api")"#)
        .unwrap();

    // JS function: (url) => "response:" + url
    let js_fn = Function::new_with_args("url", "return 'response:' + url");

    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("webServiceFn"), &js_fn).unwrap();

    wb.calculate(Some(opts.into())).unwrap();

    assert_eq!(
        sheet.get_calculated_value("A1").unwrap().as_text().unwrap(),
        "response:https://example.com/api"
    );
}

#[wasm_bindgen_test]
fn test_calculation_with_rtd_fn() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .set_formula("A1", r#"=RTD("prog","srv","topic1")"#)
        .unwrap();

    // JS function: (progId, server, topics) => progId + ":" + server + ":" + topics.join(",")
    let js_fn = Function::new_with_args(
        "progId, server, topics",
        "return progId + ':' + server + ':' + topics.join(',')",
    );

    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("rtdFn"), &js_fn).unwrap();

    wb.calculate(Some(opts.into())).unwrap();

    assert_eq!(
        sheet.get_calculated_value("A1").unwrap().as_text().unwrap(),
        "prog:srv:topic1"
    );
}

#[wasm_bindgen_test]
fn test_calculation_with_external_fn() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from_f64(7.0)).unwrap();
    sheet
        .set_formula("B1", r#"=[1]!TBLink("acct",A1)"#)
        .unwrap();

    let js_fn = Function::new_with_args(
        "book, name, args",
        "return book + ':' + name + ':' + args.join(',')",
    );

    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("externalFn"), &js_fn).unwrap();

    wb.calculate(Some(opts.into())).unwrap();

    assert_eq!(
        sheet.get_calculated_value("B1").unwrap().as_text().unwrap(),
        "1:TBLink:acct,7"
    );
}

#[wasm_bindgen_test]
fn test_web_service_fn_returning_null() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .set_formula("A1", r#"=WEBSERVICE("https://example.com/data")"#)
        .unwrap();

    // Callback returns null => should produce #N/A
    let js_fn = Function::new_with_args("url", "return null");

    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("webServiceFn"), &js_fn).unwrap();

    wb.calculate(Some(opts.into())).unwrap();

    let val = sheet.get_calculated_value("A1").unwrap();
    assert!(val.is_error());
    assert_eq!(val.as_error().unwrap(), "#N/A");
}

// Sparse Row Iteration Tests

#[wasm_bindgen_test]
fn test_get_rows_batch_basic() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(10.0)).unwrap();
    sheet.set_cell("C1", JsValue::from_str("hello")).unwrap();
    sheet.set_cell("A3", JsValue::from_f64(42.0)).unwrap();
    sheet.set_cell("B5", JsValue::from_bool(true)).unwrap();

    let result = sheet.get_rows_batch(0, 1000, JsValue::UNDEFINED).unwrap();
    let rows: Vec<JsValue> = js_sys::Array::from(&result).iter().collect();

    assert_eq!(rows.len(), 3); // rows 0, 2, 4

    // Row 0 should have 2 cells (A1 and C1)
    let row0 = &rows[0];
    assert_eq!(get_f64_field(row0, "index") as u32, 0);
    let cells0: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(row0, &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(cells0.len(), 2);
    assert_eq!(get_f64_field(&cells0[0], "col") as u32, 0);
    assert_eq!(get_string_field(&cells0[0], "value"), "10");
    assert_eq!(get_f64_field(&cells0[1], "col") as u32, 2);
    assert_eq!(get_string_field(&cells0[1], "value"), "hello");

    // Row 2 (index 2) has A3
    let row1 = &rows[1];
    assert_eq!(get_f64_field(row1, "index") as u32, 2);

    // Row 4 (index 4) has B5
    let row2 = &rows[2];
    assert_eq!(get_f64_field(row2, "index") as u32, 4);
}

#[wasm_bindgen_test]
fn test_get_rows_batch_empty_sheet() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let result = sheet.get_rows_batch(0, 1000, JsValue::UNDEFINED).unwrap();
    let rows: Vec<JsValue> = js_sys::Array::from(&result).iter().collect();
    assert_eq!(rows.len(), 0);
}

#[wasm_bindgen_test]
fn test_get_rows_batch_range() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("first")).unwrap();
    sheet.set_cell("A100", JsValue::from_str("last")).unwrap();

    // Only row 0 in range 0..49
    let batch1 = sheet.get_rows_batch(0, 50, JsValue::UNDEFINED).unwrap();
    let rows1: Vec<JsValue> = js_sys::Array::from(&batch1).iter().collect();
    assert_eq!(rows1.len(), 1);
    assert_eq!(get_f64_field(&rows1[0], "index") as u32, 0);

    // Only row 99 in range 50..149
    let batch2 = sheet.get_rows_batch(50, 100, JsValue::UNDEFINED).unwrap();
    let rows2: Vec<JsValue> = js_sys::Array::from(&batch2).iter().collect();
    assert_eq!(rows2.len(), 1);
    assert_eq!(get_f64_field(&rows2[0], "index") as u32, 99);

    // No data beyond row 99
    let batch3 = sheet.get_rows_batch(100, 100, JsValue::UNDEFINED).unwrap();
    let rows3: Vec<JsValue> = js_sys::Array::from(&batch3).iter().collect();
    assert_eq!(rows3.len(), 0);
}

#[wasm_bindgen_test]
fn test_get_rows_batch_calculated() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(10.0)).unwrap();
    sheet.set_cell("A2", JsValue::from_f64(20.0)).unwrap();
    sheet.set_formula("A3", "=A1+A2").unwrap();
    wb.calculate(None).unwrap();

    let opts = make_options(&[("useCalculatedValues", JsValue::from(true))]);
    let result = sheet.get_rows_batch(0, 1000, opts).unwrap();
    let rows: Vec<JsValue> = js_sys::Array::from(&result).iter().collect();

    // Row 2 (A3) should have the calculated value
    let row2 = &rows[2];
    let cells: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(row2, &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(get_string_field(&cells[0], "value"), "30");
}

#[wasm_bindgen_test]
fn test_get_rows_batch_skip_empty() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_str("merged")).unwrap();
    sheet.merge_cells("A1:C1").unwrap();
    sheet.set_cell("A2", JsValue::from_str("")).unwrap();

    let opts = make_options(&[
        ("includeMergeInfo", JsValue::from(true)),
        ("skipEmptyValues", JsValue::from(true)),
    ]);
    let result = sheet.get_rows_batch(0, 10, opts).unwrap();
    let rows: Vec<JsValue> = js_sys::Array::from(&result).iter().collect();

    assert_eq!(rows.len(), 2);
    assert_eq!(get_f64_field(&rows[0], "index") as u32, 0);
    let merged_cells: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(&rows[0], &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(merged_cells.len(), 1);
    assert_eq!(get_f64_field(&merged_cells[0], "col") as u32, 0);
    assert_eq!(get_string_field(&merged_cells[0], "value"), "merged");

    assert_eq!(get_f64_field(&rows[1], "index") as u32, 1);
    let empty_string_cells: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(&rows[1], &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(empty_string_cells.len(), 1);
    assert_eq!(get_f64_field(&empty_string_cells[0], "col") as u32, 0);
    assert_eq!(get_string_field(&empty_string_cells[0], "value"), "");
}

#[wasm_bindgen_test]
fn test_get_rows_batch_skip_blank() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    sheet.set_cell("A1", JsValue::from_f64(10.0)).unwrap();
    sheet.set_cell("B1", JsValue::from_str("")).unwrap();
    sheet.set_cell("A2", JsValue::from_str("merged")).unwrap();
    sheet.merge_cells("A2:C2").unwrap();

    let opts = make_options(&[
        ("includeMergeInfo", JsValue::from(true)),
        ("skipBlankValues", JsValue::from(true)),
    ]);
    let result = sheet.get_rows_batch(0, 10, opts).unwrap();
    let rows: Vec<JsValue> = js_sys::Array::from(&result).iter().collect();

    assert_eq!(rows.len(), 2);
    assert_eq!(get_f64_field(&rows[0], "index") as u32, 0);
    let first_row_cells: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(&rows[0], &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(first_row_cells.len(), 1);
    assert_eq!(get_f64_field(&first_row_cells[0], "col") as u32, 0);
    assert_eq!(get_string_field(&first_row_cells[0], "value"), "10");

    assert_eq!(get_f64_field(&rows[1], "index") as u32, 1);
    let merged_row_cells: Vec<JsValue> =
        js_sys::Array::from(&Reflect::get(&rows[1], &JsValue::from_str("cells")).unwrap())
            .iter()
            .collect();
    assert_eq!(merged_row_cells.len(), 1);
    assert_eq!(get_f64_field(&merged_row_cells[0], "col") as u32, 0);
    assert_eq!(get_string_field(&merged_row_cells[0], "value"), "merged");
}

// Password-protected save/open round-trips

const PASSWORD: &str = "duke-test-pw";

fn encrypted_round_trip_xlsx(profile: Option<&str>, key_bits: Option<u32>) {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from_str("hello")).unwrap();
    sheet.set_cell("B1", JsValue::from_f64(42.0)).unwrap();

    let bytes = wb
        .save_xlsx_bytes_encrypted(PASSWORD, profile.map(String::from), key_bits, None)
        .expect("save_xlsx_bytes_encrypted");

    let opened = Workbook::from_bytes_with_password(&bytes, PASSWORD, None)
        .expect("from_bytes_with_password");
    let sheet = opened.get_sheet(0).unwrap();
    assert_eq!(sheet.get_cell("A1").unwrap().as_text().as_deref(), Some("hello"));
    assert_eq!(sheet.get_cell("B1").unwrap().as_number(), Some(42.0));
}

fn encrypted_round_trip_xls(profile: Option<&str>, key_bits: Option<u32>) {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from_str("hello")).unwrap();
    sheet.set_cell("B1", JsValue::from_f64(42.0)).unwrap();

    let bytes = wb
        .save_xls_bytes_encrypted(PASSWORD, profile.map(String::from), key_bits)
        .expect("save_xls_bytes_encrypted");

    let opened = Workbook::from_bytes_with_password(&bytes, PASSWORD, None)
        .expect("from_bytes_with_password");
    let sheet = opened.get_sheet(0).unwrap();
    assert_eq!(sheet.get_cell("A1").unwrap().as_text().as_deref(), Some("hello"));
    assert_eq!(sheet.get_cell("B1").unwrap().as_number(), Some(42.0));
}

#[wasm_bindgen_test]
fn test_save_xlsx_encrypted_default_round_trips() {
    encrypted_round_trip_xlsx(None, None);
}

#[wasm_bindgen_test]
fn test_save_xlsx_encrypted_agile_256_round_trips() {
    encrypted_round_trip_xlsx(Some("agile"), Some(256));
}

#[wasm_bindgen_test]
fn test_save_xlsx_encrypted_standard_round_trips() {
    encrypted_round_trip_xlsx(Some("standard"), None);
}

#[wasm_bindgen_test]
fn test_save_xls_encrypted_default_round_trips() {
    encrypted_round_trip_xls(None, None);
}

#[wasm_bindgen_test]
fn test_save_xls_encrypted_rc4_cryptoapi_40_round_trips() {
    encrypted_round_trip_xls(Some("rc4-cryptoapi"), Some(40));
}

#[wasm_bindgen_test]
fn test_save_xls_encrypted_rc4_legacy_round_trips() {
    encrypted_round_trip_xls(Some("rc4-legacy"), None);
}

#[wasm_bindgen_test]
fn test_save_xls_encrypted_xor_round_trips() {
    encrypted_round_trip_xls(Some("xor"), None);
}

#[wasm_bindgen_test]
fn test_wrong_password_rejects() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet.set_cell("A1", JsValue::from_f64(1.0)).unwrap();
    let bytes = wb
        .save_xlsx_bytes_encrypted(PASSWORD, None, None, None)
        .expect("save");
    let err = Workbook::from_bytes_with_password(&bytes, "wrong", None);
    assert!(err.is_err(), "wrong password must reject");
}

#[wasm_bindgen_test]
fn test_unknown_xlsx_profile_rejects() {
    let wb = Workbook::new();
    let res = wb.save_xlsx_bytes_encrypted(PASSWORD, Some("not-a-thing".into()), None, None);
    assert!(res.is_err(), "unknown profile must reject");
}
