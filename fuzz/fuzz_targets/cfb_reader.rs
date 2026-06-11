//! Fuzz target for the OLE2/CFB container reader used by XLS.
//!
//! This is narrower and faster than `xls_reader`: it reaches CFB header,
//! FAT/DIFAT, directory, and stream-chain parsing directly. The DIFAT
//! count OOM fixed in `read_difat` is caught here without needing a
//! valid Workbook stream.

#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    let _ = duke_sheets_xls::cfb::CompoundFile::open(Cursor::new(data));
});
