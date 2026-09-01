//! Fuzz target for XLS write -> read roundtrip.
//!
//! Builds a small workbook from `Arbitrary` input, writes it through the
//! BIFF8 writer, then reads it back. This targets the CFB envelope,
//! formula compiler, comments, autofilters, and data-validation writer
//! paths changed in the XLS parity work.

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, ValueFilter};
use duke_sheets_core::comment::CellComment;
use duke_sheets_core::validation::DataValidation;
use duke_sheets_core::{CellAddress, CellRange, CellValue, Workbook};
use duke_sheets_xls::{XlsReader, XlsWriter};

#[path = "common/form_controls.rs"]
mod form_controls;
use form_controls::FuzzFormControl;

#[derive(Arbitrary, Debug)]
struct FuzzWorkbook {
    cells: Vec<FuzzCell>,
    comment: Option<SmallString>,
    list_validation: Option<SmallString>,
    autofilter: bool,
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
    Empty,
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
            if (32..128).contains(&c) && c != b'\0' {
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
        let cell = |u: &mut Unstructured<'a>| -> arbitrary::Result<String> {
            let c = (b'A' + u.int_in_range(0..=15)? as u8) as char;
            let r = u.int_in_range(1..=100)?;
            Ok(format!("{c}{r}"))
        };
        let kind = u.int_in_range(0..=9)?;
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
            4 => format!("=CHOOSE(1,{},{})", cell(u)?, cell(u)?),
            5 => format!("=-{}", cell(u)?),
            6 => format!("=+{}", cell(u)?),
            7 => format!("=COUNT(({},{}))", cell(u)?, cell(u)?),
            8 => {
                let a = u.int_in_range(0..=99)?;
                let b = u.int_in_range(0..=99)?;
                format!("=SUM({{{a},{b};1,2}})")
            }
            _ => format!("={}:{}", cell(u)?, cell(u)?),
        };
        Ok(SmallFormula(f))
    }
}

fn addr(col: u8, row: u16) -> String {
    let col = col % 16;
    format!("{}{}", (b'A' + col) as char, row as u32 % 5000 + 1)
}

fuzz_target!(|data: &[u8]| {
    let fwb = match FuzzWorkbook::arbitrary(&mut Unstructured::new(data)) {
        Ok(wb) => wb,
        Err(_) => return,
    };

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Header").ok();
    sheet.set_cell_value("A2", 1.0).ok();
    sheet.set_cell_value("A3", 2.0).ok();

    for cell in fwb.cells.iter().take(60) {
        let a = addr(cell.col, cell.row);
        match &cell.value {
            FuzzValue::Empty => {}
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
                if let Ok(parsed) = CellAddress::parse(&a) {
                    let _ = sheet.set_formula_result(parsed.row, parsed.col, CellValue::Number(1.0));
                }
            }
        }
    }

    if let Some(text) = &fwb.comment {
        let _ = sheet.set_comment("B1", CellComment::new("fuzz", text.0.as_str()));
    }

    if fwb.autofilter {
        let mut af = AutoFilter::new(CellRange::from_indices(0, 0, 2, 0));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Values(ValueFilter {
                values: vec!["1".to_string()],
                blank: false,
            }),
        ));
        sheet.set_auto_filter(Some(af));
    }

    if let Some(list) = &fwb.list_validation {
        let mut dv = DataValidation::list(if list.0.is_empty() { "x" } else { &list.0 });
        dv.ranges = vec![CellRange::from_indices(0, 2, 4, 2)];
        sheet.add_data_validation(dv);
    }

    for control in fwb.controls.iter().take(6) {
        sheet.add_form_control(control.to_control(), control.anchor());
    }
    let control_count = sheet.form_control_count();

    let bytes = match XlsWriter::write_to_bytes(&workbook) {
        Ok(bytes) => bytes,
        Err(_) => return,
    };
    let rt = XlsReader::read(Cursor::new(bytes)).expect("our own XLS output must read back");
    assert_eq!(
        rt.worksheet(0).unwrap().form_control_count(),
        control_count,
        "form control count mismatch"
    );
});
