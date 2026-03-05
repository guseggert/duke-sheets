//! Fuzz target for the CSV reader.
//!
//! Generates structured CSV content and reader options via `Arbitrary`
//! to exercise type detection, delimiter handling, quoting edge cases,
//! and worksheet population with controlled inputs.

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

/// Structured CSV input: options + content generated together.
#[derive(Arbitrary, Debug)]
struct FuzzCsv {
    /// Delimiter byte (the fuzzer picks interesting values)
    delimiter: u8,
    /// Whether to treat first row as header
    has_header: bool,
    /// Whether to auto-detect types
    auto_detect: bool,
    /// Row data
    rows: Vec<FuzzRow>,
}

#[derive(Debug)]
struct FuzzRow(Vec<String>);

impl<'a> Arbitrary<'a> for FuzzRow {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let ncols = u.int_in_range(1..=20)?;
        let mut cells = Vec::with_capacity(ncols);
        for _ in 0..ncols {
            cells.push(arbitrary_cell_value(u)?);
        }
        Ok(FuzzRow(cells))
    }
}

/// Generate cell values that exercise type detection: numbers, booleans,
/// dates, empty strings, quoted strings with embedded delimiters/newlines.
fn arbitrary_cell_value(u: &mut Unstructured) -> arbitrary::Result<String> {
    let kind: u8 = u.int_in_range(0..=7)?;
    match kind {
        0 => {
            // Integer
            let n: i32 = u.arbitrary()?;
            Ok(n.to_string())
        }
        1 => {
            // Float
            let f: f64 = u.arbitrary()?;
            if f.is_finite() {
                Ok(format!("{}", f))
            } else {
                Ok("NaN".into())
            }
        }
        2 => {
            // Boolean
            Ok(if u.arbitrary()? { "true" } else { "false" }.into())
        }
        3 => {
            // Empty
            Ok(String::new())
        }
        4 => {
            // Plain string (alphanumeric)
            let len = u.int_in_range(0..=30)?;
            let mut s = String::new();
            for _ in 0..len {
                s.push(u.int_in_range(b'a'..=b'z')? as char);
            }
            Ok(s)
        }
        5 => {
            // String with special chars (needs quoting)
            let len = u.int_in_range(1..=20)?;
            let mut s = String::new();
            for _ in 0..len {
                let c: u8 = u.arbitrary()?;
                // Include chars that stress CSV parsing: commas, quotes, newlines
                s.push(match c % 10 {
                    0 => ',',
                    1 => '"',
                    2 => '\n',
                    3 => '\r',
                    4 => '\t',
                    5 => ' ',
                    _ => (b'A' + (c % 26)) as char,
                });
            }
            Ok(s)
        }
        6 => {
            // Date-like string
            let y = u.int_in_range(1900..=2100)?;
            let m = u.int_in_range(1..=12)?;
            let d = u.int_in_range(1..=28)?;
            Ok(format!("{}-{:02}-{:02}", y, m, d))
        }
        _ => {
            // yes/no (exercises boolean detection)
            Ok(if u.arbitrary()? { "yes" } else { "no" }.into())
        }
    }
}

impl FuzzCsv {
    /// Render into CSV bytes, quoting fields that contain the delimiter,
    /// quotes, or newlines.
    fn render(&self) -> Vec<u8> {
        let delim = self.delimiter as char;
        let mut out = String::new();
        for row in &self.rows {
            for (i, cell) in row.0.iter().enumerate() {
                if i > 0 {
                    out.push(delim);
                }
                // Quote if needed
                if cell.contains(delim)
                    || cell.contains('"')
                    || cell.contains('\n')
                    || cell.contains('\r')
                {
                    out.push('"');
                    out.push_str(&cell.replace('"', "\"\""));
                    out.push('"');
                } else {
                    out.push_str(cell);
                }
            }
            out.push('\n');
        }
        out.into_bytes()
    }
}

fuzz_target!(|data: &[u8]| {
    // Strategy 1: Structured generation
    if let Ok(fuzz_csv) = FuzzCsv::arbitrary(&mut Unstructured::new(data)) {
        let csv_bytes = fuzz_csv.render();
        let mut opts = duke_sheets_csv::CsvReadOptions::default();
        opts.delimiter = fuzz_csv.delimiter;
        opts.has_header = fuzz_csv.has_header;
        opts.auto_detect_types = fuzz_csv.auto_detect;
        let _ = duke_sheets_csv::CsvReader::read(Cursor::new(&csv_bytes), &opts);
    }

    // Strategy 2: Raw bytes with default options (catches encoding edge cases)
    let _ = duke_sheets_csv::CsvReader::read(
        Cursor::new(data),
        &duke_sheets_csv::CsvReadOptions::default(),
    );
});
