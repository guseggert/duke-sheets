use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::Worksheet;

/// Write the sheet's comments part with the MS-XLSB 2.4.33 record
/// sequence (BrtBeginComments 0x0274 .. BrtEndComments 0x0275).
/// Excel refuses the off-spec 0x0278-based ids our writer used to
/// emit; the reader still accepts both.
pub(crate) fn write_comments<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    index: usize,
    ws: &Worksheet,
) -> XlsbResult<()> {
    let mut comments: Vec<((u32, u16), &duke_sheets_core::comment::CellComment)> =
        ws.comments().collect();
    if comments.is_empty() {
        return Ok(());
    }
    comments.sort_by_key(|((row, col), _)| (*row, *col));

    let authors = ws.comment_authors();

    let path = format!("xl/comments{}.bin", index + 1);
    zip.start_file(&path, *options)?;

    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(records::BRT_BEGIN_COMMENTS, &[])?;

    rw.write_record(records::BRT_BEGIN_COMMENT_AUTHORS, &[])?;
    for author in &authors {
        rw.write_record(records::BRT_COMMENT_AUTHOR, &encode_wide_str(author))?;
    }
    rw.write_record(records::BRT_END_COMMENT_AUTHORS, &[])?;

    rw.write_record(records::BRT_BEGIN_COMMENT_LIST, &[])?;
    for (seq, ((row, col), comment)) in comments.iter().enumerate() {
        let author_id = authors
            .iter()
            .position(|a| a == &comment.author)
            .unwrap_or(0) as u32;

        // iauthor + UncheckedRfX (rwFirst rwLast colFirst colLast) +
        // a stable 16-byte GUID (Excel emits a random one).
        let mut payload = Vec::with_capacity(36);
        payload.extend_from_slice(&author_id.to_le_bytes());
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&(*col as u32).to_le_bytes());
        payload.extend_from_slice(&(*col as u32).to_le_bytes());
        payload.extend_from_slice(&comment_guid(*row, *col, seq as u32));
        rw.write_record(records::BRT_BEGIN_COMMENT, &payload)?;

        let mut text_payload = Vec::new();
        text_payload.push(0x00);
        text_payload.extend_from_slice(&encode_wide_str(&comment.text));
        rw.write_record(records::BRT_COMMENT_TEXT, &text_payload)?;

        rw.write_record(records::BRT_END_COMMENT, &[])?;
    }
    rw.write_record(records::BRT_END_COMMENT_LIST, &[])?;

    rw.write_record(records::BRT_END_COMMENTS, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

/// Deterministic per-comment GUID (any stable value is acceptable;
/// version/variant bits set for a well-formed v4 layout).
fn comment_guid(row: u32, col: u16, seq: u32) -> [u8; 16] {
    let mut guid = [0u8; 16];
    guid[0..4].copy_from_slice(&row.to_le_bytes());
    guid[4..6].copy_from_slice(&col.to_le_bytes());
    guid[6] = 0x40; // version 4
    guid[7] = 0x5D;
    guid[8] = 0x80; // RFC 4122 variant
    guid[9] = 0x75;
    guid[10..14].copy_from_slice(&seq.to_le_bytes());
    guid[14] = 0xD5;
    guid[15] = 0x0B;
    guid
}
