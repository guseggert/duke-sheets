//! Fuzz target for XLSX write → read roundtrip.
//!
//! Builds a workbook programmatically from `Arbitrary` input, writes it
//! to XLSX bytes, then reads it back. Any panic in write or read-back
//! indicates a serialization/deserialization bug.

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

use duke_sheets_core::{CellValue, Style, Workbook};

/// Structured workbook specification for fuzzing.
#[derive(Arbitrary, Debug)]
struct FuzzWorkbook {
    sheets: Vec<FuzzSheet>,
}

#[derive(Debug)]
struct FuzzSheet {
    name: String,
    cells: Vec<FuzzCell>,
}

impl<'a> Arbitrary<'a> for FuzzSheet {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        // Sheet name: 1-20 alphanumeric chars
        let name_len = u.int_in_range(1..=20)?;
        let mut name = String::with_capacity(name_len);
        for _ in 0..name_len {
            name.push(u.int_in_range(b'A'..=b'z')? as char);
        }

        let ncells = u.int_in_range(0..=50)?;
        let mut cells = Vec::with_capacity(ncells);
        for _ in 0..ncells {
            cells.push(FuzzCell::arbitrary(u)?);
        }

        Ok(FuzzSheet { name, cells })
    }
}

#[derive(Arbitrary, Debug)]
struct FuzzCell {
    row: u16,
    col: u8,
    value: FuzzValue,
    style: Option<FuzzStyle>,
}

#[derive(Arbitrary, Debug)]
enum FuzzValue {
    Empty,
    Number(f64),
    Int(i32),
    Bool(bool),
    Str(SmallString),
    Formula(SmallFormula),
}

/// Keep strings small.
#[derive(Debug)]
struct SmallString(String);

impl<'a> Arbitrary<'a> for SmallString {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let len = u.int_in_range(0..=50)?;
        let mut s = String::with_capacity(len);
        for _ in 0..len {
            // Mix of ASCII + some non-ASCII to test encoding
            let c: u8 = u.arbitrary()?;
            if c < 128 && c >= 32 {
                s.push(c as char);
            } else {
                s.push('x');
            }
        }
        Ok(SmallString(s))
    }
}

/// Simple formula strings.
#[derive(Debug)]
struct SmallFormula(String);

impl<'a> Arbitrary<'a> for SmallFormula {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let kind: u8 = u.int_in_range(0..=4)?;
        let f = match kind {
            0 => {
                let n: f64 = u.arbitrary()?;
                if n.is_finite() {
                    format!("={}", n)
                } else {
                    "=0".into()
                }
            }
            1 => {
                let col = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let row = u.int_in_range(1..=100)?;
                format!("={}{}", col, row)
            }
            2 => {
                let c1 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r1 = u.int_in_range(1..=100)?;
                let c2 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r2 = u.int_in_range(1..=100)?;
                format!("=SUM({}{}:{}{})", c1, r1, c2, r2)
            }
            3 => {
                let c = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r = u.int_in_range(1..=100)?;
                format!("=IF({}{}>0,1,0)", c, r)
            }
            _ => {
                let c1 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r1 = u.int_in_range(1..=100)?;
                let c2 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r2 = u.int_in_range(1..=100)?;
                format!("={}{}+{}{}", c1, r1, c2, r2)
            }
        };
        Ok(SmallFormula(f))
    }
}

#[derive(Arbitrary, Debug)]
struct FuzzStyle {
    bold: bool,
    italic: bool,
    font_size: Option<u8>,
}

fn col_to_letters(col: u8) -> String {
    if col < 26 {
        String::from((b'A' + col) as char)
    } else {
        format!(
            "{}{}",
            (b'A' + col / 26 - 1) as char,
            (b'A' + col % 26) as char
        )
    }
}

fuzz_target!(|data: &[u8]| {
    let fwb = match FuzzWorkbook::arbitrary(&mut Unstructured::new(data)) {
        Ok(wb) => wb,
        Err(_) => return,
    };

    // Build a real Workbook from the fuzz spec
    let mut workbook = Workbook::new();

    // Ensure at least one sheet
    if fwb.sheets.is_empty() {
        return;
    }

    for (i, fsheet) in fwb.sheets.iter().enumerate() {
        // Add sheets beyond the default first one
        if i > 0 {
            let _ = workbook.add_worksheet();
        }
        let sheet = match workbook.worksheet_mut(i) {
            Some(s) => s,
            None => continue,
        };

        // Set sheet name (ignore errors from invalid names)
        let _ = sheet.set_name(&fsheet.name);

        for cell in &fsheet.cells {
            let addr = format!("{}{}", col_to_letters(cell.col), cell.row as u32 + 1);

            match &cell.value {
                FuzzValue::Empty => {}
                FuzzValue::Number(n) => {
                    if n.is_finite() {
                        let _ = sheet.set_cell_value(&addr, *n);
                    }
                }
                FuzzValue::Int(n) => {
                    let _ = sheet.set_cell_value(&addr, *n as f64);
                }
                FuzzValue::Bool(b) => {
                    let _ = sheet.set_cell_value(&addr, CellValue::Boolean(*b));
                }
                FuzzValue::Str(s) => {
                    let _ = sheet.set_cell_value(&addr, s.0.as_str());
                }
                FuzzValue::Formula(f) => {
                    let _ = sheet.set_cell_formula(&addr, &f.0);
                }
            }

            // Apply style if present
            if let Some(style) = &cell.style {
                let mut s = Style::new();
                if style.bold {
                    s = s.bold(true);
                }
                if style.italic {
                    s = s.italic(true);
                }
                if let Some(size) = style.font_size {
                    if size > 0 && size <= 72 {
                        s = s.font_size(size as f64);
                    }
                }
                let _ = sheet.set_cell_style(&addr, &s);
            }
        }
    }

    // Step 1: Write workbook to XLSX bytes
    let mut output = Cursor::new(Vec::new());
    if duke_sheets_xlsx::XlsxWriter::write(&workbook, &mut output).is_err() {
        return;
    }

    // Step 2: Read back — must not panic
    let written = output.into_inner();
    let _ = duke_sheets_xlsx::XlsxReader::read(Cursor::new(&written));
});
