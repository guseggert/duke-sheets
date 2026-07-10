//! Fuzz target for the VML control-shape parser.
//!
//! Feeds arbitrary bytes to `parse_vml_controls` (x:ClientData fields,
//! MultiSel index lists, x:Anchor px tuples, textbox caption HTML) and
//! converts each parsed shape to a model FormControl, mirroring the
//! XLSB reader's legacy-drawing path.

#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    for shape in duke_sheets_vml::parse_vml_controls(data) {
        let _ = shape.to_form_control();
    }
});
