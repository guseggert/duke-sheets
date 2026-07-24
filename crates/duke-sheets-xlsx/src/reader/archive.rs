//! ZIP archive lookup helpers that tolerate non-standard entry name spellings.
//!
//! OOXML packages are ZIP archives, and the ZIP spec requires forward-slash
//! path separators. Some third-party XLSX writers leak the Windows convention
//! and emit entry names like `xl\_rels\workbook.xml.rels`. Those files open
//! fine in Excel because Excel normalizes, so we do the same rather than
//! rejecting them.

use std::io::{Read, Seek};

/// Look up a ZIP entry by its OPC-equivalent name.
///
/// On success returns a `ZipFile` for the requested entry. On failure returns
/// `ZipError::FileNotFound` (or the underlying error) as if the fallback had
/// not been attempted, so callers' error messages stay canonical.
pub(crate) fn archive_by_name<'a, R: Read + Seek>(
    archive: &'a mut zip::ZipArchive<R>,
    path: &str,
) -> zip::result::ZipResult<zip::read::ZipFile<'a>> {
    // Collect the actual name to open first, while only holding an immutable
    // borrow of the archive. Then take the mutable borrow for by_name exactly
    // once. Doing it this way keeps the borrow checker happy without relying
    // on a match-and-retry structure.
    let exact_exists = archive.file_names().any(|n| n == path);
    if exact_exists {
        return archive.by_name(path);
    }
    let equivalent = archive
        .file_names()
        .find(|name| zip_name_eq(name, path))
        .map(str::to_string);
    if let Some(equivalent) = equivalent {
        return archive.by_name(&equivalent);
    }
    // Fall through: let the exact-name lookup produce the idiomatic error.
    archive.by_name(path)
}

fn zip_name_eq(left: &str, right: &str) -> bool {
    left.bytes()
        .map(normalize_zip_name_byte)
        .eq(right.bytes().map(normalize_zip_name_byte))
}

fn normalize_zip_name_byte(byte: u8) -> u8 {
    if byte == b'\\' {
        b'/'
    } else {
        byte.to_ascii_lowercase()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    fn zip_with_entries(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = zip::write::SimpleFileOptions::default();
            for (name, data) in entries {
                zip.start_file(*name, opts).unwrap();
                zip.write_all(data).unwrap();
            }
            zip.finish().unwrap();
        }
        bytes
    }

    #[test]
    fn archive_by_name_finds_exact_forward_slash() {
        let bytes = zip_with_entries(&[("xl/_rels/workbook.xml.rels", b"rels")]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut buf = String::new();
        archive_by_name(&mut archive, "xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut buf)
            .unwrap();
        assert_eq!(buf, "rels");
    }

    /// Buggy third-party XLSX writers leak Windows path separators into ZIP
    /// entry names, so `xl\_rels\workbook.xml.rels` appears even though ZIP
    /// spec mandates forward slashes. We fall back to that spelling.
    #[test]
    fn archive_by_name_falls_back_to_backslashes() {
        let bytes = zip_with_entries(&[("xl\\_rels\\workbook.xml.rels", b"rels")]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut buf = String::new();
        archive_by_name(&mut archive, "xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut buf)
            .unwrap();
        assert_eq!(buf, "rels");
    }

    /// A canonical forward-slash entry still wins over a backslash variant if
    /// both are present (which shouldn't really happen, but the preference is
    /// stable rather than order-dependent on archive iteration).
    #[test]
    fn archive_by_name_prefers_forward_slash_when_both_present() {
        let bytes = zip_with_entries(&[
            ("xl\\_rels\\workbook.xml.rels", b"backslash"),
            ("xl/_rels/workbook.xml.rels", b"forward"),
        ]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut buf = String::new();
        archive_by_name(&mut archive, "xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut buf)
            .unwrap();
        assert_eq!(buf, "forward");
    }

    #[test]
    fn archive_by_name_uses_opc_case_insensitive_equivalence() {
        let bytes = zip_with_entries(&[("XL/Workbook.xml", b"wb")]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut buf = String::new();
        archive_by_name(&mut archive, "xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut buf)
            .unwrap();
        assert_eq!(buf, "wb");
    }

    #[test]
    fn archive_by_name_returns_file_not_found_for_missing_entries() {
        let bytes = zip_with_entries(&[("xl/workbook.xml", b"wb")]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let outcome = match archive_by_name(&mut archive, "xl/sharedStrings.xml") {
            Ok(_) => "unexpected_ok",
            Err(zip::result::ZipError::FileNotFound) => "file_not_found",
            Err(_) => "other_error",
        };
        assert_eq!(outcome, "file_not_found");
    }
}
