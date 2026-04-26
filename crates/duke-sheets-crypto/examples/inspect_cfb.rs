//! Walk a CFB file and print all streams + sizes. Useful for diffing
//! our encrypted output against Office or LibreOffice-produced files.
//!
//! Usage:
//!   cargo run -p duke-sheets-crypto --example inspect_cfb -- <path>

fn main() {
    let path = std::env::args()
        .nth(1)
        .expect("usage: inspect_cfb <path-to-cfb>");
    let file = std::fs::File::open(&path).expect("open file");
    let comp = cfb::CompoundFile::open(file).expect("parse CFB");

    println!("CFB version: {:?}", comp.version());
    println!();
    for entry in comp.walk() {
        let path = entry.path().display().to_string();
        let kind = if entry.is_storage() {
            "storage"
        } else {
            "stream"
        };
        let size = entry.len();
        println!("{kind:8} {size:>10}  {path}");
    }
}
