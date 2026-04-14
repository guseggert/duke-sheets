use std::io::{BufReader, Read, Seek};

use duke_sheets_core::comment::CellComment;
use duke_sheets_core::{CellAddress, Worksheet};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::biff12::records;
use crate::biff12::{parser, RecordIter};
use crate::error::XlsbResult;

pub(crate) fn read_comments<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    comments_path: &str,
    ws: &mut Worksheet,
) -> XlsbResult<()> {
    if comments_path.ends_with(".bin") {
        if archive.by_name(comments_path).is_ok() {
            let file = archive.by_name(comments_path).unwrap();
            return read_comments_bin(file, ws);
        }
        let xml_path = comments_path.replace(".bin", ".xml");
        if archive.by_name(&xml_path).is_ok() {
            let file = archive.by_name(&xml_path).unwrap();
            return read_comments_xml(file, ws);
        }
    } else if comments_path.ends_with(".xml") {
        if archive.by_name(comments_path).is_ok() {
            let file = archive.by_name(comments_path).unwrap();
            return read_comments_xml(file, ws);
        }
    } else {
        let bin_path = format!("{}.bin", comments_path);
        if archive.by_name(&bin_path).is_ok() {
            let file = archive.by_name(&bin_path).unwrap();
            return read_comments_bin(file, ws);
        }
    }
    Ok(())
}

fn read_comments_bin<R: Read>(reader: R, ws: &mut Worksheet) -> XlsbResult<()> {
    let mut iter = RecordIter::new(reader);
    let mut buf = Vec::with_capacity(1024);
    let mut authors: Vec<String> = Vec::new();

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => return Ok(()),
        };
        match typ {
            records::BRT_COMMENT_AUTHOR => {
                if let Ok((s, _)) = parser::wide_str(&buf, 0) {
                    authors.push(s);
                }
            }
            records::BRT_END_COMMENT_AUTHORS => break,
            records::BRT_BEGIN_COMMENT_AUTHORS => {}
            _ => {
                if len == 0 && typ == 0 {
                    return Ok(());
                }
            }
        }
    }

    let mut current_row: Option<u32> = None;
    let mut current_col: Option<u16> = None;
    let mut current_author_id: Option<u32> = None;

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => break,
        };
        match typ {
            records::BRT_BEGIN_COMMENT => {
                if len >= 12 {
                    current_author_id = Some(parser::read_u32(&buf, 0));
                    current_row = Some(parser::read_u32(&buf, 4));
                    current_col = Some(parser::read_u32(&buf, 8) as u16);
                }
            }
            records::BRT_COMMENT_TEXT => {
                if let (Some(row), Some(col)) = (current_row, current_col) {
                    let text = parse_comment_text(&buf[..len]);
                    let author = current_author_id
                        .and_then(|id| authors.get(id as usize))
                        .cloned()
                        .unwrap_or_default();
                    let comment = CellComment::new(author, text);
                    ws.set_comment_at(row, col, comment);
                }
            }
            records::BRT_END_COMMENT => {
                current_row = None;
                current_col = None;
                current_author_id = None;
            }
            records::BRT_END_COMMENT_LIST => break,
            _ => {}
        }
    }

    Ok(())
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

fn read_comments_xml<R: Read>(reader: R, ws: &mut Worksheet) -> XlsbResult<()> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut authors: Vec<String> = Vec::new();
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
                            ws.set_comment_at(addr.row, addr.col, comment);
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

    Ok(())
}
