//! Fuzz target for the XLSB (BIFF12) reader.
//!
//! Feeds arbitrary bytes as a ZIP container to `XlsbReader::read()`.
//! Exercises: ZIP part discovery, BIFF12 record framing (varint type +
//! size), workbook/worksheet/SST/styles record parsing, formula token
//! parsing incl. array-constant rgcb, autofilters, data validations,
//! hyperlinks, header/footer strings.

#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    let _ = duke_sheets_xlsb::XlsbReader::read(Cursor::new(data));
});
