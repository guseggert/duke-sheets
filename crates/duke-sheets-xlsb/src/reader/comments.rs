use std::io::{BufReader, Read, Seek};

use duke_sheets_core::comment::CellComment;
use duke_sheets_core::CellAddress;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::biff12::records;
use crate::biff12::{parser, RecordIter};
use crate::error::XlsbResult;

/// Read a comments part into `(row, col, comment)` entries, in part
/// order. Accepts both the MS-XLSB record ids (0x0274-based) and the
/// off-spec 0x0278-based ids our old writer emitted; an XML comments
/// part is also accepted as a fallback.
pub(crate) fn read_comments_list<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    comments_path: &str,
) -> XlsbResult<Vec<(u32, u16, CellComment)>> {
    if comments_path.ends_with(".bin") {
        if archive.by_name(comments_path).is_ok() {
            let file = archive.by_name(comments_path).unwrap();
            return read_comments_bin(file);
        }
        let xml_path = comments_path.replace(".bin", ".xml");
        if archive.by_name(&xml_path).is_ok() {
            let file = archive.by_name(&xml_path).unwrap();
            return read_comments_xml(file);
        }
    } else if comments_path.ends_with(".xml") {
        if archive.by_name(comments_path).is_ok() {
            let file = archive.by_name(comments_path).unwrap();
            return read_comments_xml(file);
        }
    } else {
        let bin_path = format!("{}.bin", comments_path);
        if archive.by_name(&bin_path).is_ok() {
            let file = archive.by_name(&bin_path).unwrap();
            return read_comments_bin(file);
        }
    }
    Ok(Vec::new())
}

/// Record-id dialect of a comments part, keyed off the first record:
/// spec parts open with BrtBeginComments (0x0274), our legacy parts
/// opened with the off-spec BrtBeginCommentAuthors (0x0278).
#[derive(Clone, Copy, PartialEq)]
enum CommentDialect {
    Spec,
    Legacy,
}

fn read_comments_bin<R: Read>(reader: R) -> XlsbResult<Vec<(u32, u16, CellComment)>> {
    let mut iter = RecordIter::new(reader);
    let mut buf = Vec::with_capacity(1024);
    let mut authors: Vec<String> = Vec::new();
    let mut comments = Vec::new();

    let mut dialect: Option<CommentDialect> = None;
    let mut current_row: Option<u32> = None;
    let mut current_col: Option<u16> = None;
    let mut current_author_id: Option<u32> = None;

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => break,
        };
        // Latch on the first comment record: BrtACBegin/BrtACEnd
        // future-record wrappers (and their wrapped version payloads)
        // may precede BrtBeginComments and must not decide the
        // dialect.
        if dialect.is_none()
            && matches!(typ, records::BRT_AC_BEGIN | records::BRT_AC_END)
        {
            continue;
        }
        let dialect = *dialect.get_or_insert(if typ == records::BRT_BEGIN_COMMENTS {
            CommentDialect::Spec
        } else {
            CommentDialect::Legacy
        });
        let (author_id, begin_id, end_id, text_id, end_list_id) = match dialect {
            CommentDialect::Spec => (
                records::BRT_COMMENT_AUTHOR,
                records::BRT_BEGIN_COMMENT,
                records::BRT_END_COMMENT,
                records::BRT_COMMENT_TEXT,
                records::BRT_END_COMMENT_LIST,
            ),
            CommentDialect::Legacy => (
                records::BRT_LEGACY_COMMENT_AUTHOR,
                records::BRT_LEGACY_BEGIN_COMMENT,
                records::BRT_LEGACY_END_COMMENT,
                records::BRT_LEGACY_COMMENT_TEXT,
                records::BRT_LEGACY_END_COMMENT_LIST,
            ),
        };
        match typ {
            t if t == author_id => {
                if let Ok((s, _)) = parser::wide_str(&buf, 0) {
                    authors.push(s);
                }
            }
            t if t == begin_id => {
                match dialect {
                    // iauthor u32 + UncheckedRfX (rwFirst, rwLast,
                    // colFirst, colLast) + 16-byte GUID.
                    CommentDialect::Spec if len >= 16 => {
                        current_author_id = Some(parser::read_u32(&buf, 0));
                        current_row = Some(parser::read_u32(&buf, 4));
                        current_col = Some(parser::read_u32(&buf, 12) as u16);
                    }
                    // iauthor u32 + row u32 + col u32.
                    CommentDialect::Legacy if len >= 12 => {
                        current_author_id = Some(parser::read_u32(&buf, 0));
                        current_row = Some(parser::read_u32(&buf, 4));
                        current_col = Some(parser::read_u32(&buf, 8) as u16);
                    }
                    _ => {}
                }
            }
            t if t == text_id => {
                if let (Some(row), Some(col)) = (current_row, current_col) {
                    let text = parse_comment_text(&buf[..len]);
                    let author = current_author_id
                        .and_then(|id| authors.get(id as usize))
                        .cloned()
                        .unwrap_or_default();
                    comments.push((row, col, CellComment::new(author, text)));
                }
            }
            t if t == end_id => {
                current_row = None;
                current_col = None;
                current_author_id = None;
            }
            t if t == end_list_id => break,
            // BrtACBegin/uid/BrtACEnd future-record wrappers and any
            // other unknown records are skipped transparently.
            _ => {}
        }
    }

    Ok(comments)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff12::{encode_wide_str, RecordWriter};

    /// A spec-id comments part prefixed by a BrtACBegin/BrtACEnd
    /// future-record wrapper must still latch the Spec dialect: the
    /// wrapper ids are not comment records and must not decide the
    /// dialect.
    #[test]
    fn dialect_latch_skips_ac_wrapper_records() {
        let mut part = Vec::new();
        let mut rw = RecordWriter::new(&mut part);
        // BrtACBegin wrapper around the whole part.
        rw.write_record(records::BRT_AC_BEGIN, &[0x01, 0x00, 0x00, 0x00])
            .unwrap();
        rw.write_record(records::BRT_BEGIN_COMMENTS, &[]).unwrap();
        rw.write_record(records::BRT_BEGIN_COMMENT_AUTHORS, &[])
            .unwrap();
        rw.write_record(records::BRT_COMMENT_AUTHOR, &encode_wide_str("probe"))
            .unwrap();
        rw.write_record(records::BRT_END_COMMENT_AUTHORS, &[])
            .unwrap();
        rw.write_record(records::BRT_BEGIN_COMMENT_LIST, &[]).unwrap();
        let mut begin = Vec::new();
        begin.extend_from_slice(&0u32.to_le_bytes()); // iauthor
        begin.extend_from_slice(&5u32.to_le_bytes()); // rwFirst
        begin.extend_from_slice(&5u32.to_le_bytes()); // rwLast
        begin.extend_from_slice(&3u32.to_le_bytes()); // colFirst
        begin.extend_from_slice(&3u32.to_le_bytes()); // colLast
        begin.extend_from_slice(&[0u8; 16]); // guid
        rw.write_record(records::BRT_BEGIN_COMMENT, &begin).unwrap();
        let mut text = vec![0x00];
        text.extend_from_slice(&encode_wide_str("a note"));
        rw.write_record(records::BRT_COMMENT_TEXT, &text).unwrap();
        rw.write_record(records::BRT_END_COMMENT, &[]).unwrap();
        rw.write_record(records::BRT_END_COMMENT_LIST, &[]).unwrap();
        rw.write_record(records::BRT_END_COMMENTS, &[]).unwrap();
        rw.write_record(records::BRT_AC_END, &[]).unwrap();
        drop(rw);

        let comments = read_comments_bin(std::io::Cursor::new(part)).unwrap();
        assert_eq!(comments.len(), 1, "spec dialect latched past the AC wrapper");
        let (row, col, comment) = &comments[0];
        assert_eq!((*row, *col), (5, 3));
        assert_eq!(comment.author, "probe");
        assert_eq!(comment.plain_text(), "a note");
    }
}

fn parse_comment_text(data: &[u8]) -> String {
    // RichStr: flags(1 byte) + XLWideString
    if data.len() < 5 {
        return String::new();
    }
    match parser::wide_str(data, 1) {
        Ok((s, _)) => s,
        Err(_) => String::new(),
    }
}

fn read_comments_xml<R: Read>(reader: R) -> XlsbResult<Vec<(u32, u16, CellComment)>> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut authors: Vec<String> = Vec::new();
    let mut comments = Vec::new();
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_text = false;
    let mut in_t = false;
    let mut current_ref: Option<String> = None;
    let mut current_author_id: Option<usize> = None;
    let mut current_text = String::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"author" => in_author = true,
                b"comment" => {
                    in_comment = true;
                    current_ref = None;
                    current_author_id = None;
                    current_text.clear();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"ref" => {
                                current_ref = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"authorId" => {
                                current_author_id =
                                    attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                            _ => {}
                        }
                    }
                }
                b"text" if in_comment => in_text = true,
                b"t" if in_text => in_t = true,
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"author" => in_author = false,
                b"comment" => {
                    if let Some(ref cell_ref) = current_ref {
                        if let Ok(addr) = CellAddress::parse(cell_ref) {
                            let author = current_author_id
                                .and_then(|id| authors.get(id))
                                .cloned()
                                .unwrap_or_default();
                            let comment = CellComment::new(author, current_text.trim());
                            comments.push((addr.row, addr.col, comment));
                        }
                    }
                    in_comment = false;
                    current_text.clear();
                }
                b"text" => in_text = false,
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_author {
                    if let Ok(text) = e.unescape() {
                        authors.push(text.to_string());
                    }
                } else if in_t {
                    if let Ok(text) = e.unescape() {
                        if !current_text.is_empty() {
                            current_text.push(' ');
                        }
                        current_text.push_str(&text);
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                if e.name().local_name().as_ref() == b"author" {
                    authors.push(String::new());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                log::warn!("XML error in comments: {e}");
                break;
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}
