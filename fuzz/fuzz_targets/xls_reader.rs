//! Fuzz target for the XLS (BIFF8) reader.
//!
//! Feeds arbitrary bytes as a CFB/OLE2 container to `XlsReader::read()`.
//! Exercises: CFB container parsing, BIFF8 record framing, CONTINUE
//! record merging, SST with encoding changes, style records, cell
//! records, formula token parsing, comments, conditional formatting,
//! data validation, hyperlinks.

#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    // XlsReader::read accepts any Read + Seek; Cursor<&[u8]> qualifies.
    let _ = duke_sheets_xls::XlsReader::read(Cursor::new(data));
});
