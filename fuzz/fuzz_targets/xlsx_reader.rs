//! Fuzz target for the XLSX reader.
//!
//! Feeds arbitrary bytes as a ZIP/XLSX file to `XlsxReader::read()`.
//! Exercises: ZIP parsing, XML parsing, shared strings, styles, themes,
//! conditional formatting, data validation, comments, formulas, tables.

#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    // XlsxReader::read accepts any Read + Seek; Cursor<&[u8]> qualifies.
    let _ = duke_sheets_xlsx::XlsxReader::read(Cursor::new(data));
});
