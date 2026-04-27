//! Walk a CFB file and print directory entries with sizes. Useful for
//! diffing our encrypted output against Office or LibreOffice-produced
//! files when debugging compatibility issues.
//!
//! Usage:
//!   cargo run -p duke-sheets-crypto --example inspect_cfb -- <path>

use duke_sheets_xls::cfb::CompoundFile;

fn main() {
    let path = std::env::args()
        .nth(1)
        .expect("usage: inspect_cfb <path-to-cfb>");
    let bytes = std::fs::read(&path).expect("read file");
    let cfb = CompoundFile::open(std::io::Cursor::new(&bytes)).expect("parse CFB");

    println!("file: {path} ({} bytes)", bytes.len());
    println!();
    for entry in cfb.directory_entries() {
        let kind = match entry.object_type {
            5 => "root",
            1 => "stor",
            2 => "strm",
            _ => "????",
        };
        let raw_name: Vec<u8> = entry.name.bytes().collect();
        println!(
            "{kind:>4}  size={:>10}  name={:?}  raw_bytes={:02x?}",
            entry.stream_size, entry.name, raw_name
        );
    }
}
