use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::Worksheet;

pub(crate) fn write_comments<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    index: usize,
    ws: &Worksheet,
) -> XlsbResult<()> {
    let comments: Vec<((u32, u16), &duke_sheets_core::comment::CellComment)> =
        ws.comments().collect();
    if comments.is_empty() {
        return Ok(());
    }

    let mut authors: Vec<String> = Vec::new();
    for (_, comment) in &comments {
        if !authors.contains(&comment.author) {
            authors.push(comment.author.clone());
        }
    }

    let path = format!("xl/comments{}.bin", index + 1);
    zip.start_file(&path, *options)?;

    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(records::BRT_BEGIN_COMMENT_AUTHORS, &[])?;
    for author in &authors {
        rw.write_record(records::BRT_COMMENT_AUTHOR, &encode_wide_str(author))?;
    }
    rw.write_record(records::BRT_END_COMMENT_AUTHORS, &[])?;

    rw.write_record(records::BRT_BEGIN_COMMENT_LIST, &[])?;
    for ((row, col), comment) in &comments {
        let author_id = authors
            .iter()
            .position(|a| a == &comment.author)
            .unwrap_or(0) as u32;

        let mut comment_payload = Vec::new();
        comment_payload.extend_from_slice(&author_id.to_le_bytes());
        comment_payload.extend_from_slice(&row.to_le_bytes());
        comment_payload.extend_from_slice(&(*col as u32).to_le_bytes());
        rw.write_record(records::BRT_BEGIN_COMMENT, &comment_payload)?;

        let mut text_payload = Vec::new();
        text_payload.push(0x00);
        text_payload.extend_from_slice(&encode_wide_str(&comment.text));
        rw.write_record(records::BRT_COMMENT_TEXT, &text_payload)?;

        rw.write_record(records::BRT_END_COMMENT, &[])?;
    }
    rw.write_record(records::BRT_END_COMMENT_LIST, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}
