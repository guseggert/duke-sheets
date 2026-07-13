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

#[wasm_bindgen_test]
fn test_workbook_protection() {
    let wb = Workbook::new();
    let protection = make_options(&[
        ("structure", JsValue::TRUE),
        ("windows", JsValue::TRUE),
        ("password", JsValue::from_str("password")),
    ]);
    wb.set_workbook_protection(protection).unwrap();

    let protection = wb.workbook_protection().unwrap();
    assert!(get_bool_field(&protection, "structure"));
    assert!(get_bool_field(&protection, "windows"));
    assert_eq!(get_f64_field(&protection, "passwordHash") as u16, 0x83af);

    wb.set_workbook_protection(JsValue::NULL).unwrap();
    assert!(wb.workbook_protection().unwrap().is_null());
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

#[wasm_bindgen_test]
fn test_worksheet_protection_and_protected_ranges() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let protection = make_options(&[
        ("protected", JsValue::TRUE),
        ("password", JsValue::from_str("password")),
        ("selectLockedCells", JsValue::TRUE),
        ("selectUnlockedCells", JsValue::TRUE),
        ("formatCells", JsValue::TRUE),
        ("sort", JsValue::TRUE),
    ]);
    sheet.set_protection(protection).unwrap();

    let protection = sheet.protection().unwrap();
    assert!(get_bool_field(&protection, "protected"));
    assert_eq!(get_f64_field(&protection, "passwordHash") as u16, 0x83af);
    assert!(get_bool_field(&protection, "selectLockedCells"));
    assert!(get_bool_field(&protection, "selectUnlockedCells"));
    assert!(get_bool_field(&protection, "formatCells"));
    assert!(get_bool_field(&protection, "sort"));

    let ranges = make_array(&[JsValue::from_str("A1:B2"), JsValue::from_str("D4:D5")]);
    let protected_range = make_options(&[
        ("name", JsValue::from_str("Editable")),
        ("ranges", ranges),
        ("passwordHash", JsValue::from_f64(0xcafe as f64)),
        ("securityDescriptor", JsValue::from_str("S-1-5-21")),
    ]);
    sheet
        .set_protected_ranges(make_array(&[protected_range]))
        .unwrap();

    let protected_ranges = Array::from(&sheet.protected_ranges().unwrap());
    assert_eq!(protected_ranges.length(), 1);
    let first = protected_ranges.get(0);
    assert_eq!(get_string_field(&first, "name"), "Editable");
    assert_eq!(get_f64_field(&first, "passwordHash") as u16, 0xcafe);
    assert_eq!(get_string_field(&first, "securityDescriptor"), "S-1-5-21");
    let first_ranges = Array::from(&Reflect::get(&first, &JsValue::from_str("ranges")).unwrap());
    assert_eq!(first_ranges.length(), 2);
    assert_eq!(first_ranges.get(0).as_string().unwrap(), "A1:B2");
    assert_eq!(first_ranges.get(1).as_string().unwrap(), "D4:D5");

    sheet.set_protection(JsValue::NULL).unwrap();
    assert!(sheet.protection().unwrap().is_null());
}

#[wasm_bindgen_test]
fn test_sheet_protection_defaults_allow_selection() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    let protection = make_options(&[("password", JsValue::from_str("password"))]);
    sheet.set_protection(protection).unwrap();

    let protection = sheet.protection().unwrap();
    assert!(get_bool_field(&protection, "protected"));
    assert_eq!(get_f64_field(&protection, "passwordHash") as u16, 0x83af);
    assert!(get_bool_field(&protection, "selectLockedCells"));
    assert!(get_bool_field(&protection, "selectUnlockedCells"));
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

fn make_array(values: &[JsValue]) -> JsValue {
    let array = Array::new();
    for value in values {
        array.push(value);
    }
    array.into()
}

fn drawing_marker(col: u32, row: u32) -> JsValue {
    make_options(&[
        ("col", JsValue::from_f64(col as f64)),
        ("row", JsValue::from_f64(row as f64)),
    ])
}

fn drawing_anchor(from_col: u32, from_row: u32, to_col: u32, to_row: u32) -> JsValue {
    make_options(&[
        ("type", JsValue::from_str("twoCell")),
        ("from", drawing_marker(from_col, from_row)),
        ("to", drawing_marker(to_col, to_row)),
        ("editAs", JsValue::from_str("twoCell")),
    ])
}

fn child_transform(x: u32, y: u32) -> JsValue {
    make_options(&[
        ("xEmu", JsValue::from_f64(x as f64)),
        ("yEmu", JsValue::from_f64(y as f64)),
        ("cxEmu", JsValue::from_f64(100_000.0)),
        ("cyEmu", JsValue::from_f64(50_000.0)),
    ])
}

fn drawing_text(text: &str) -> JsValue {
    let run = make_options(&[("text", JsValue::from_str(text))]);
    make_options(&[("runs", make_array(&[run]))])
}

fn shape_drawing(name: &str, anchor: JsValue) -> JsValue {
    let shape = make_options(&[("geometry", JsValue::from_str("rect"))]);
    make_options(&[
        ("name", JsValue::from_str(name)),
        ("anchor", anchor),
        ("kind", JsValue::from_str("shape")),
        ("shape", shape),
    ])
}

fn form_control_drawing(
    name: &str,
    anchor: JsValue,
    control_kind: JsValue,
) -> JsValue {
    let form_control = make_options(&[("kind", control_kind)]);
    make_options(&[
        ("name", JsValue::from_str(name)),
        ("anchor", anchor),
        ("kind", JsValue::from_str("formControl")),
        ("formControl", form_control),
    ])
}

fn drawing_path(values: &[u32]) -> JsValue {
    make_array(
        &values
            .iter()
            .map(|value| JsValue::from_f64(*value as f64))
            .collect::<Vec<_>>(),
    )
}

fn assert_drawing_path(drawing: &JsValue, expected: &[f64]) {
    let path = Array::from(&Reflect::get(drawing, &"drawingPath".into()).unwrap());
    let actual: Vec<f64> = path.iter().map(|value| value.as_f64().unwrap()).collect();
    assert_eq!(actual, expected);
}

#[wasm_bindgen_test]
fn test_drawing_paths_order_and_nested_groups() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_drawing(shape_drawing(
            "back",
            drawing_anchor(0, 0, 1, 1),
        ))
        .unwrap();

    let label = make_options(&[
        ("transform", child_transform(0, 0)),
        ("kind", JsValue::from_str("formControl")),
        (
            "formControl",
            make_options(&[(
                "kind",
                make_options(&[
                    ("kind", JsValue::from_str("label")),
                    ("caption", drawing_text("nested")),
                ]),
            )]),
        ),
    ]);
    let nested_shape = make_options(&[
        ("transform", child_transform(10, 20)),
        ("kind", JsValue::from_str("shape")),
        (
            "shape",
            make_options(&[("geometry", JsValue::from_str("ellipse"))]),
        ),
    ]);
    let nested_group = make_options(&[
        ("transform", child_transform(100, 200)),
        ("kind", JsValue::from_str("group")),
        (
            "group",
            make_options(&[("children", make_array(&[nested_shape]))]),
        ),
    ]);
    let group = make_options(&[
        ("name", JsValue::from_str("front group")),
        ("anchor", drawing_anchor(1, 1, 5, 6)),
        ("kind", JsValue::from_str("group")),
        (
            "group",
            make_options(&[("children", make_array(&[label, nested_group]))]),
        ),
    ]);
    sheet.add_drawing(group).unwrap();

    let drawings = Array::from(&sheet.drawings().unwrap());
    assert_eq!(drawings.length(), 2);
    assert_eq!(get_string_field(&drawings.get(0), "name"), "back");
    assert_eq!(get_string_field(&drawings.get(1), "name"), "front group");
    assert_drawing_path(&drawings.get(0), &[0.0]);
    assert_drawing_path(&drawings.get(1), &[1.0]);
    let group = Reflect::get(&drawings.get(1), &"group".into()).unwrap();
    let children = Array::from(&Reflect::get(&group, &"children".into()).unwrap());
    assert_drawing_path(&children.get(0), &[1.0, 0.0]);
    let nested = Reflect::get(&children.get(1), &"group".into()).unwrap();
    let grandchildren = Array::from(&Reflect::get(&nested, &"children".into()).unwrap());
    assert_drawing_path(&grandchildren.get(0), &[1.0, 1.0, 0.0]);

    let controls = Array::from(&sheet.form_controls().unwrap());
    assert_eq!(controls.length(), 1);
    assert_drawing_path(&controls.get(0), &[1.0, 0.0]);
}

#[wasm_bindgen_test]
fn test_drawing_image_bytes_are_lazy() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    let bytes = js_sys::Uint8Array::from(&[137, 80, 78, 71][..]);
    let image = make_options(&[
        ("format", JsValue::from_str("png")),
        ("widthEmu", JsValue::from_f64(200_000.0)),
        ("heightEmu", JsValue::from_f64(100_000.0)),
        ("data", bytes.into()),
    ]);
    let drawing = make_options(&[
        ("anchor", drawing_anchor(0, 0, 2, 2)),
        ("kind", JsValue::from_str("image")),
        ("image", image),
    ]);
    sheet.add_drawing(drawing).unwrap();

    let images = Array::from(&sheet.images().unwrap());
    let image = Reflect::get(&images.get(0), &"image".into()).unwrap();
    assert!(Reflect::get(&image, &"data".into()).unwrap().is_undefined());
    assert_eq!(
        sheet
            .drawing_image_data(drawing_path(&[0]))
            .unwrap(),
        [137, 80, 78, 71]
    );
    assert!(sheet
        .drawing_svg_data(drawing_path(&[0]))
        .unwrap()
        .is_none());
}

#[wasm_bindgen_test]
fn test_drawing_mutation_and_raw_rejection() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_drawing(shape_drawing("one", drawing_anchor(0, 0, 1, 1)))
        .unwrap();
    sheet
        .add_drawing(shape_drawing("three", drawing_anchor(2, 0, 3, 1)))
        .unwrap();
    sheet
        .insert_drawing(
            1,
            shape_drawing("two", drawing_anchor(1, 0, 2, 1)),
        )
        .unwrap();
    sheet.move_drawing(2, 0).unwrap();

    let child = make_options(&[
        ("transform", child_transform(0, 0)),
        ("kind", JsValue::from_str("shape")),
        (
            "shape",
            make_options(&[("geometry", JsValue::from_str("rect"))]),
        ),
    ]);
    let group = make_options(&[
        ("anchor", drawing_anchor(3, 0, 5, 2)),
        ("kind", JsValue::from_str("group")),
        (
            "group",
            make_options(&[("children", make_array(&[child]))]),
        ),
    ]);
    sheet.add_drawing(group).unwrap();
    let replacement = make_options(&[
        ("name", JsValue::from_str("nested label")),
        ("transform", child_transform(5, 5)),
        ("kind", JsValue::from_str("formControl")),
        (
            "formControl",
            make_options(&[(
                "kind",
                make_options(&[
                    ("kind", JsValue::from_str("label")),
                    ("caption", drawing_text("replacement")),
                ]),
            )]),
        ),
    ]);
    sheet
        .set_drawing(drawing_path(&[3, 0]), replacement)
        .unwrap();
    assert_eq!(sheet.form_control_count().unwrap(), 1);
    sheet.remove_drawing(drawing_path(&[3, 0])).unwrap();
    assert_eq!(sheet.form_control_count().unwrap(), 0);

    let raw = make_options(&[
        ("anchor", drawing_anchor(0, 0, 1, 1)),
        ("kind", JsValue::from_str("raw")),
        ("raw", make_options(&[])),
    ]);
    assert!(sheet.add_drawing(raw).is_err());
}

fn comment_drawing(anchor: JsValue, row: u32, col: u32, text: &str) -> JsValue {
    let comment = make_options(&[
        ("row", JsValue::from_f64(row as f64)),
        ("col", JsValue::from_f64(col as f64)),
        ("author", JsValue::from_str("author")),
        ("text", JsValue::from_str(text)),
    ]);
    make_options(&[
        ("anchor", anchor),
        ("kind", JsValue::from_str("comment")),
        ("comment", comment),
    ])
}

fn form_control_kind(control: &JsValue) -> JsValue {
    Reflect::get(
        &Reflect::get(control, &"formControl".into()).unwrap(),
        &"kind".into(),
    )
    .unwrap()
}

fn byte_array_to_vec(value: &JsValue) -> Vec<u8> {
    Array::from(value)
        .iter()
        .map(|item| item.as_f64().unwrap() as u8)
        .collect()
}

#[wasm_bindgen_test]
fn test_unknown_control_passthrough_survives_set_drawing_and_save() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    let raw_properties = make_array(&[
        make_array(&[JsValue::from_str("customFlag"), JsValue::from_str("kept")]),
        make_array(&[JsValue::from_str("val"), JsValue::from_str("17")]),
        make_array(&[JsValue::from_str("fmlaLink"), JsValue::from_str("$A$1")]),
    ]);
    let raw_client_data = make_array(&[make_array(
        &b"<x:Val>17</x:Val>"
            .iter()
            .map(|byte| JsValue::from_f64(f64::from(*byte)))
            .collect::<Vec<_>>(),
    )]);
    sheet
        .add_drawing(form_control_drawing(
            "Legacy editor",
            drawing_anchor(1, 2, 3, 4),
            make_options(&[
                ("kind", JsValue::from_str("unknown")),
                ("objectType", JsValue::from_str("EditBox")),
                ("caption", drawing_text("Unsupported editor")),
                ("rawProperties", raw_properties),
                ("rawClientData", raw_client_data),
            ]),
        ))
        .unwrap();
    let bytes = wb.save_xlsx_bytes().unwrap();

    let first = Workbook::from_bytes(&bytes).unwrap();
    let sheet = first.get_sheet(0).unwrap();
    let control = Array::from(&sheet.form_controls().unwrap()).get(0);
    let kind = form_control_kind(&control);
    assert_eq!(get_string_field(&kind, "kind"), "unknown");
    assert_eq!(get_string_field(&kind, "objectType"), "EditBox");
    let properties = Array::from(&Reflect::get(&kind, &"rawProperties".into()).unwrap());
    let property_count = properties.length();
    assert!(
        property_count >= 3,
        "expected raw properties to survive the read, got {property_count}"
    );
    let client_data = Array::from(&Reflect::get(&kind, &"rawClientData".into()).unwrap());
    assert!(client_data.length() >= 1);

    // Identity rewrite: the read snapshot keeps its passthrough data.
    sheet.set_drawing(drawing_path(&[0]), control).unwrap();
    let bytes = first.save_xlsx_bytes().unwrap();

    let second = Workbook::from_bytes(&bytes).unwrap();
    let sheet = second.get_sheet(0).unwrap();
    let control = Array::from(&sheet.form_controls().unwrap()).get(0);
    let kind = form_control_kind(&control);
    assert_eq!(get_string_field(&kind, "objectType"), "EditBox");
    let properties = Array::from(&Reflect::get(&kind, &"rawProperties".into()).unwrap());
    assert_eq!(properties.length(), property_count);
    let has_custom_flag = properties.iter().any(|pair| {
        let pair = Array::from(&pair);
        pair.get(0).as_string().as_deref() == Some("customFlag")
            && pair.get(1).as_string().as_deref() == Some("kept")
    });
    assert!(has_custom_flag, "customFlag raw property must survive rewrite");
    let client_data = Array::from(&Reflect::get(&kind, &"rawClientData".into()).unwrap());
    let has_val_fragment = client_data.iter().any(|fragment| {
        String::from_utf8_lossy(&byte_array_to_vec(&fragment)).contains("<x:Val>17</x:Val>")
    });
    assert!(has_val_fragment, "raw ClientData fragment must survive rewrite");
}

#[wasm_bindgen_test]
fn test_uint8array_payloads_survive_add_and_set_drawing() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();

    // Image svgData as Uint8Array.
    let svg = b"<svg xmlns='http://www.w3.org/2000/svg'/>";
    let image = make_options(&[
        ("format", JsValue::from_str("png")),
        ("widthEmu", JsValue::from_f64(200_000.0)),
        ("heightEmu", JsValue::from_f64(100_000.0)),
        (
            "data",
            js_sys::Uint8Array::from(&[137u8, 80, 78, 71][..]).into(),
        ),
        ("svgData", js_sys::Uint8Array::from(&svg[..]).into()),
    ]);
    sheet
        .add_drawing(make_options(&[
            ("anchor", drawing_anchor(0, 0, 2, 2)),
            ("kind", JsValue::from_str("image")),
            ("image", image),
        ]))
        .unwrap();
    assert_eq!(
        sheet
            .drawing_svg_data(drawing_path(&[0]))
            .unwrap()
            .as_deref(),
        Some(&svg[..])
    );

    // Unknown-control rawClientData elements and rawObj as Uint8Array,
    // exercised through both addDrawing and setDrawing (both pass the
    // payload through serde's tagged-enum buffering).
    let fragment = b"<x:Val>17</x:Val>";
    let obj_body = [0x15u8, 0x00, 0x12, 0x00];
    let unknown_control = |name: &str| {
        form_control_drawing(
            name,
            drawing_anchor(1, 2, 3, 4),
            make_options(&[
                ("kind", JsValue::from_str("unknown")),
                ("objectType", JsValue::from_str("EditBox")),
                ("caption", drawing_text("editor")),
                (
                    "rawClientData",
                    make_array(&[js_sys::Uint8Array::from(&fragment[..]).into()]),
                ),
                ("rawObj", js_sys::Uint8Array::from(&obj_body[..]).into()),
            ]),
        )
    };
    sheet.add_drawing(unknown_control("added")).unwrap();
    sheet
        .set_drawing(drawing_path(&[1]), unknown_control("replaced"))
        .unwrap();

    let control = Array::from(&sheet.form_controls().unwrap()).get(0);
    assert_eq!(get_string_field(&control, "name"), "replaced");
    let kind = form_control_kind(&control);
    let client_data = Array::from(&Reflect::get(&kind, &"rawClientData".into()).unwrap());
    assert_eq!(byte_array_to_vec(&client_data.get(0)), fragment.to_vec());
    let raw_obj = Reflect::get(&kind, &"rawObj".into()).unwrap();
    assert_eq!(byte_array_to_vec(&raw_obj), obj_body.to_vec());
}

#[wasm_bindgen_test]
fn test_set_drawing_rejects_comment_group_children() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    let child = make_options(&[
        ("transform", child_transform(0, 0)),
        ("kind", JsValue::from_str("shape")),
        (
            "shape",
            make_options(&[("geometry", JsValue::from_str("rect"))]),
        ),
    ]);
    let group = make_options(&[
        ("anchor", drawing_anchor(0, 0, 2, 2)),
        ("kind", JsValue::from_str("group")),
        (
            "group",
            make_options(&[("children", make_array(&[child]))]),
        ),
    ]);
    sheet.add_drawing(group).unwrap();
    let comment_child = make_options(&[
        ("transform", child_transform(0, 0)),
        ("kind", JsValue::from_str("comment")),
        (
            "comment",
            make_options(&[
                ("row", JsValue::from_f64(0.0)),
                ("col", JsValue::from_f64(0.0)),
                ("author", JsValue::from_str("a")),
                ("text", JsValue::from_str("nested")),
            ]),
        ),
    ]);
    assert!(sheet
        .set_drawing(drawing_path(&[0, 0]), comment_child)
        .is_err());
}

#[wasm_bindgen_test]
fn test_set_drawing_rejects_duplicate_comment_cell() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_drawing(comment_drawing(drawing_anchor(2, 0, 4, 2), 0, 0, "first"))
        .unwrap();
    sheet
        .add_drawing(comment_drawing(drawing_anchor(2, 4, 4, 6), 4, 0, "second"))
        .unwrap();

    // Replacing the second comment with one for the first comment's
    // cell must be rejected.
    assert!(sheet
        .set_drawing(
            drawing_path(&[1]),
            comment_drawing(drawing_anchor(2, 4, 4, 6), 0, 0, "duplicate"),
        )
        .is_err());

    // Replacing a comment in place (same cell) stays allowed.
    sheet
        .set_drawing(
            drawing_path(&[0]),
            comment_drawing(drawing_anchor(2, 0, 4, 2), 0, 0, "updated"),
        )
        .unwrap();
    assert_eq!(Array::from(&sheet.drawings().unwrap()).length(), 2);
}

#[wasm_bindgen_test]
fn test_form_control_radio_semantics_and_linked_cell_sync() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_drawing(form_control_drawing(
            "group",
            drawing_anchor(0, 0, 4, 6),
            make_options(&[
                ("kind", JsValue::from_str("groupBox")),
                ("caption", drawing_text("Choose")),
            ]),
        ))
        .unwrap();
    for (name, state, row) in [("one", "checked", 1), ("two", "unchecked", 3)] {
        sheet
            .add_drawing(form_control_drawing(
                name,
                drawing_anchor(1, row, 2, row + 1),
                make_options(&[
                    ("kind", JsValue::from_str("optionButton")),
                    ("caption", drawing_text(name)),
                    ("state", JsValue::from_str(state)),
                    ("cellLink", JsValue::from_str("$D$2")),
                ]),
            ))
            .unwrap();
    }

    let result = sheet
        .set_form_control_check_state(drawing_path(&[2]), "checked")
        .unwrap();
    assert_eq!(get_f64_field(&result, "controlsChanged"), 2.0);
    assert_eq!(get_f64_field(&result, "linkedCellsChanged"), 1.0);
    assert_eq!(sheet.get_cell("D2").unwrap().as_number(), Some(2.0));

    let controls = Array::from(&sheet.form_controls().unwrap());
    let first = Reflect::get(&controls.get(1), &"formControl".into()).unwrap();
    let second = Reflect::get(&controls.get(2), &"formControl".into()).unwrap();
    assert_eq!(
        get_string_field(&Reflect::get(&first, &"kind".into()).unwrap(), "state"),
        "unchecked"
    );
    assert_eq!(
        get_string_field(&Reflect::get(&second, &"kind".into()).unwrap(), "state"),
        "checked"
    );
}

#[wasm_bindgen_test]
fn test_drawing_rich_caption_and_shared_metadata() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    let font = make_options(&[("bold", JsValue::TRUE), ("size", JsValue::from_f64(14.0))]);
    let runs = make_array(&[
        make_options(&[("text", JsValue::from_str("Run "))]),
        make_options(&[("text", JsValue::from_str("now")), ("font", font)]),
    ]);
    let caption = make_options(&[
        ("runs", runs),
        ("horizontalAlignment", JsValue::from_str("center")),
    ]);
    let form_control = make_options(&[
        (
            "kind",
            make_options(&[
                ("kind", JsValue::from_str("button")),
                ("caption", caption),
            ]),
        ),
        ("macroName", JsValue::from_str("RunReport")),
    ]);
    let drawing = make_options(&[
        ("name", JsValue::from_str("Run button")),
        ("hidden", JsValue::TRUE),
        ("locked", JsValue::FALSE),
        ("printable", JsValue::FALSE),
        ("altText", JsValue::from_str("Runs the report")),
        ("title", JsValue::from_str("Report action")),
        ("anchor", drawing_anchor(0, 0, 2, 1)),
        ("kind", JsValue::from_str("formControl")),
        ("formControl", form_control),
    ]);
    sheet.add_drawing(drawing).unwrap();

    let drawing = Array::from(&sheet.drawings().unwrap()).get(0);
    assert_eq!(get_string_field(&drawing, "name"), "Run button");
    assert!(get_bool_field(&drawing, "hidden"));
    assert!(!get_bool_field(&drawing, "locked"));
    assert!(!get_bool_field(&drawing, "printable"));
    assert_eq!(get_string_field(&drawing, "altText"), "Runs the report");
    assert_eq!(get_string_field(&drawing, "title"), "Report action");
    let control = Reflect::get(&drawing, &"formControl".into()).unwrap();
    assert_eq!(get_string_field(&control, "macroName"), "RunReport");
    let kind = Reflect::get(&control, &"kind".into()).unwrap();
    let caption = Reflect::get(&kind, &"caption".into()).unwrap();
    let runs = Array::from(&Reflect::get(&caption, &"runs".into()).unwrap());
    assert_eq!(runs.length(), 2);
    assert_eq!(get_string_field(&runs.get(1), "text"), "now");
}

/// CamelCase names are checked through JS reflection because direct Rust calls
/// do not exercise wasm-bindgen's exported property names.
#[wasm_bindgen_test]
fn test_drawing_js_api_names() {
    let wb = Workbook::new();
    let sheet = wb.get_sheet(0).unwrap();
    sheet
        .add_drawing(shape_drawing("shape", drawing_anchor(0, 0, 1, 1)))
        .unwrap();

    let sheet_js = JsValue::from(sheet);
    for name in ["drawings", "formControls", "images", "charts", "chartsEx"] {
        let value = Reflect::get(&sheet_js, &name.into()).unwrap();
        assert!(Array::is_array(&value), "{name} getter missing");
    }
    for name in [
        "addDrawing",
        "insertDrawing",
        "setDrawing",
        "removeDrawing",
        "moveDrawing",
        "drawingImageData",
        "drawingSvgData",
        "setFormControlCheckState",
    ] {
        let method = Reflect::get(&sheet_js, &name.into()).unwrap();
        assert!(method.is_function(), "{name} missing from the JS API");
    }
    for removed in ["addFormControl", "setFormControl", "removeFormControl"] {
        assert!(
            Reflect::get(&sheet_js, &removed.into())
                .unwrap()
                .is_undefined(),
            "{removed} must not remain in the JS API"
        );
    }

    let workbook_js = JsValue::from(wb);
    for name in ["syncFormControls", "syncFormControlsFromLinkedCells"] {
        assert!(
            Reflect::get(&workbook_js, &name.into())
                .unwrap()
                .is_function(),
            "{name} missing from the JS API"
        );
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
    use duke_sheets_core::validation::DataValidation;
    use duke_sheets_core::{CellAddress, CellRange};

    let mut core_wb = duke_sheets_core::Workbook::new();
    let ws = core_wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Score").unwrap();
    ws.set_cell_value("A2", 1.0).unwrap();
    ws.set_comment(
        "A1",
        CellComment::new("Tester", "fixture comment"),
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
    assert_eq!(
        sheet.get_cell("A1").unwrap().as_text().as_deref(),
        Some("hello")
    );
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
    assert_eq!(
        sheet.get_cell("A1").unwrap().as_text().as_deref(),
        Some("hello")
    );
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
