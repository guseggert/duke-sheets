//! Fuzz target for XLSB write → read roundtrip.
//!
//! Builds a workbook programmatically from `Arbitrary` input, writes it
//! to XLSB bytes, then reads it back. Any panic in write or read-back
//! indicates a serialization/deserialization bug. Formulas lean on the
//! shapes the BIFF12 compiler special-cases (IF/CHOOSE/SUM attrs, array
//! constants, unions, names).

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

use duke_sheets_core::{CellValue, Workbook};

#[path = "common/form_controls.rs"]
mod form_controls;
use form_controls::FuzzFormControl;

#[derive(Arbitrary, Debug)]
struct FuzzWorkbook {
    cells: Vec<FuzzCell>,
    named_range: Option<bool>,
    controls: Vec<FuzzFormControl>,
}

#[derive(Arbitrary, Debug)]
struct FuzzCell {
    row: u16,
    col: u8,
    value: FuzzValue,
}

#[derive(Arbitrary, Debug)]
enum FuzzValue {
    Number(f64),
    Bool(bool),
    Str(SmallString),
    Formula(SmallFormula),
}

#[derive(Debug)]
struct SmallString(String);

impl<'a> Arbitrary<'a> for SmallString {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let len = u.int_in_range(0..=40)?;
        let mut s = String::with_capacity(len);
        for _ in 0..len {
            let c: u8 = u.arbitrary()?;
            if (32..128).contains(&c) {
                s.push(c as char);
            } else {
                s.push('x');
            }
        }
        Ok(SmallString(s))
    }
}

#[derive(Debug)]
struct SmallFormula(String);

impl<'a> Arbitrary<'a> for SmallFormula {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let kind: u8 = u.int_in_range(0..=8)?;
        let cell = |u: &mut Unstructured<'a>| -> arbitrary::Result<String> {
            let c = (b'A' + u.int_in_range(0..=25)? as u8) as char;
            let r = u.int_in_range(1..=200)?;
            Ok(format!("{c}{r}"))
        };
        let f = match kind {
            0 => {
                let n: f64 = u.arbitrary()?;
                if n.is_finite() {
                    format!("={n}")
                } else {
                    "=0".into()
                }
            }
            1 => format!("={}", cell(u)?),
            2 => format!("=SUM({}:{})", cell(u)?, cell(u)?),
            3 => format!("=IF({}>0,1,{})", cell(u)?, cell(u)?),
            4 => format!("=CHOOSE(2,{},{},3)", cell(u)?, cell(u)?),
            5 => {
                let a = u.int_in_range(0..=99)?;
                let b = u.int_in_range(0..=99)?;
                format!("=SUM({{{a},{b};1,2}})")
            }
            6 => format!("=COUNT(({},{}))", cell(u)?, cell(u)?),
            7 => format!("=-{}%", cell(u)?),
            _ => format!("=MyName+{}", cell(u)?),
        };
        Ok(SmallFormula(f))
    }
}

fn addr(col: u8, row: u16) -> String {
    let col = col % 26;
    format!("{}{}", (b'A' + col) as char, row as u32 % 5000 + 1)
}

fuzz_target!(|data: &[u8]| {
    let fwb = match FuzzWorkbook::arbitrary(&mut Unstructured::new(data)) {
        Ok(wb) => wb,
        Err(_) => return,
    };

    let mut workbook = Workbook::new();
    if fwb.named_range.is_some() {
        workbook
            .named_ranges_mut()
            .define_or_update(duke_sheets_core::named_range::NamedRange::new(
                "MyName",
                "Sheet1!$A$1:$A$3",
                duke_sheets_core::named_range::NameScope::Workbook,
            ));
    }
    let sheet = workbook.worksheet_mut(0).unwrap();
    for cell in fwb.cells.iter().take(60) {
        let a = addr(cell.col, cell.row);
        match &cell.value {
            FuzzValue::Number(n) => {
                if n.is_finite() {
                    let _ = sheet.set_cell_value(&a, *n);
                }
            }
            FuzzValue::Bool(b) => {
                let _ = sheet.set_cell_value(&a, CellValue::Boolean(*b));
            }
            FuzzValue::Str(s) => {
                let _ = sheet.set_cell_value(&a, s.0.as_str());
            }
            FuzzValue::Formula(f) => {
                let _ = sheet.set_cell_formula(&a, &f.0);
                if let Ok(parsed) = duke_sheets_core::CellAddress::parse(&a) {
                    let _ = sheet.set_formula_result(parsed.row, parsed.col, CellValue::Number(1.0));
                }
            }
        }
    }

    for control in fwb.controls.iter().take(6) {
        sheet.add_form_control(control.to_control(), control.anchor());
    }
    let control_count = sheet.form_control_count();

    let mut output = Cursor::new(Vec::new());
    if duke_sheets_xlsb::XlsbWriter::write(&workbook, &mut output).is_err() {
        return;
    }
    let written = output.into_inner();
    // Read-back must not panic; our own output must also be readable.
    let rt = duke_sheets_xlsb::XlsbReader::read(Cursor::new(&written))
        .expect("our own XLSB output must read back");
    assert_eq!(
        rt.worksheet(0).unwrap().form_control_count(),
        control_count,
        "form control count mismatch"
    );
});
