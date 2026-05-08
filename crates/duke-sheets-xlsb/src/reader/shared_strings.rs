use std::io::{Read, Seek};

use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::FontStyle;

use crate::biff12::parser;
use crate::biff12::records;
use crate::biff12::RecordIter;
use crate::error::XlsbResult;

#[derive(Debug, Clone)]
pub(crate) enum SharedStringEntry {
    Plain(String),
    Rich(Vec<RichTextRun>),
}

pub(crate) fn read_shared_strings<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    fonts: &[FontStyle],
) -> XlsbResult<Vec<SharedStringEntry>> {
    let file = match archive.by_name("xl/sharedStrings.bin") {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);
    let mut entries = Vec::new();

    let len = iter.skip_to(records::BRT_BEGIN_SST, &[], &mut buf)?;
    let count = if len >= 8 {
        parser::read_u32(&buf, 4) as usize
    } else {
        0
    };

    for _ in 0..count {
        let len = iter.skip_to(
            records::BRT_SS_ITEM,
            &[(records::BRT_BEGIN_FRT, Some(records::BRT_END_FRT))],
            &mut buf,
        )?;
        if len > 1 {
            let flags = buf[0];
            let (text, str_consumed) = parser::wide_str(&buf, 1)?;

            if flags & 0x01 != 0 {
                let runs_offset = 1 + str_consumed;
                if runs_offset < len {
                    let runs = parse_rich_runs(&text, &buf[runs_offset..len], fonts);
                    entries.push(SharedStringEntry::Rich(runs));
                } else {
                    entries.push(SharedStringEntry::Plain(text));
                }
            } else {
                entries.push(SharedStringEntry::Plain(text));
            }
        } else {
            entries.push(SharedStringEntry::Plain(String::new()));
        }
    }
    Ok(entries)
}

fn parse_rich_runs(text: &str, data: &[u8], fonts: &[FontStyle]) -> Vec<RichTextRun> {
    if data.len() < 4 {
        return vec![RichTextRun::plain(text.to_string())];
    }

    let c_runs = parser::read_u32(data, 0) as usize;
    // Per [MS-XLSB] §2.5.157 each StrRun is u16 ich + u16 ifnt = 4 bytes.
    if c_runs == 0 || data.len() < 4 + c_runs * 4 {
        return vec![RichTextRun::plain(text.to_string())];
    }

    let mut markers: Vec<(u32, u32)> = Vec::with_capacity(c_runs);
    for i in 0..c_runs {
        let off = 4 + i * 4;
        let ich = parser::read_u16(data, off) as u32;
        let ifnt = parser::read_u16(data, off + 2) as u32;
        markers.push((ich, ifnt));
    }

    let utf16: Vec<u16> = text.encode_utf16().collect();
    let total_cu = utf16.len() as u32;

    let mut runs = Vec::with_capacity(c_runs);

    for (i, &(ich, ifnt)) in markers.iter().enumerate() {
        let start = ich.min(total_cu) as usize;
        let end = if i + 1 < markers.len() {
            markers[i + 1].0.min(total_cu) as usize
        } else {
            total_cu as usize
        };

        if start >= end {
            continue;
        }

        let run_text = String::from_utf16_lossy(&utf16[start..end]);
        let run_font = fonts
            .get(ifnt as usize)
            .map(font_style_to_run_font)
            .and_then(|f| if f.is_empty() { None } else { Some(f) });

        runs.push(RichTextRun {
            text: run_text,
            font: run_font,
        });
    }

    if runs.is_empty() {
        runs.push(RichTextRun::plain(text.to_string()));
    }

    runs
}

fn font_style_to_run_font(f: &FontStyle) -> RunFont {
    let d = FontStyle::default();
    RunFont {
        bold: if f.bold != d.bold { Some(f.bold) } else { None },
        italic: if f.italic != d.italic {
            Some(f.italic)
        } else {
            None
        },
        size: if f.size != d.size { Some(f.size) } else { None },
        color: if f.color != d.color {
            Some(f.color.clone())
        } else {
            None
        },
        name: if f.name != d.name {
            Some(f.name.clone())
        } else {
            None
        },
        underline: if f.underline != d.underline {
            Some(f.underline)
        } else {
            None
        },
        strikethrough: if f.strikethrough != d.strikethrough {
            Some(f.strikethrough)
        } else {
            None
        },
        vertical_align: if f.vertical_align != d.vertical_align {
            Some(f.vertical_align)
        } else {
            None
        },
        family: f.family,
        charset: f.charset,
        scheme: f.scheme.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    use crate::biff12::build_record;

    fn utf16le(s: &str) -> Vec<u8> {
        s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect()
    }

    fn xlwide(s: &str) -> Vec<u8> {
        let encoded = utf16le(s);
        let char_count = (encoded.len() / 2) as u32;
        let mut out = Vec::new();
        out.extend_from_slice(&char_count.to_le_bytes());
        out.extend_from_slice(&encoded);
        out
    }

    fn make_sst_zip(sst_data: &[u8]) -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("xl/sharedStrings.bin", opts).unwrap();
            zip.write_all(sst_data).unwrap();
            zip.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(buf)).unwrap()
    }

    fn build_sst_plain(strings: &[&str]) -> Vec<u8> {
        let count = strings.len() as u32;
        let mut header = Vec::new();
        header.extend_from_slice(&count.to_le_bytes());
        header.extend_from_slice(&count.to_le_bytes());
        let mut data = build_record(records::BRT_BEGIN_SST, &header);
        for s in strings {
            let mut item = vec![0u8]; // flags=0 (plain)
            item.extend_from_slice(&xlwide(s));
            data.extend_from_slice(&build_record(records::BRT_SS_ITEM, &item));
        }
        data
    }

    fn build_sst_rich(text: &str, runs: &[(u16, u16)]) -> Vec<u8> {
        let count = 1u32;
        let mut header = Vec::new();
        header.extend_from_slice(&count.to_le_bytes());
        header.extend_from_slice(&count.to_le_bytes());
        let mut data = build_record(records::BRT_BEGIN_SST, &header);

        let mut item = vec![0x01u8]; // flags=1 (fRichStr)
        item.extend_from_slice(&xlwide(text));
        item.extend_from_slice(&(runs.len() as u32).to_le_bytes());
        for &(ich, ifnt) in runs {
            item.extend_from_slice(&ich.to_le_bytes());
            item.extend_from_slice(&ifnt.to_le_bytes());
        }
        data.extend_from_slice(&build_record(records::BRT_SS_ITEM, &item));
        data
    }

    #[test]
    fn read_plain_entries() {
        let sst_data = build_sst_plain(&["Hello", "World"]);
        let mut archive = make_sst_zip(&sst_data);
        let entries = read_shared_strings(&mut archive, &[]).unwrap();
        assert_eq!(entries.len(), 2);
        assert!(matches!(&entries[0], SharedStringEntry::Plain(s) if s == "Hello"));
        assert!(matches!(&entries[1], SharedStringEntry::Plain(s) if s == "World"));
    }

    #[test]
    fn read_rich_text_with_bold_font() {
        let bold_font = FontStyle {
            bold: true,
            name: "Arial".to_string(),
            size: 14.0,
            ..FontStyle::default()
        };
        let fonts = vec![FontStyle::default(), bold_font];

        let sst_data = build_sst_rich("Hello World", &[(0, 0), (6, 1)]);
        let mut archive = make_sst_zip(&sst_data);
        let entries = read_shared_strings(&mut archive, &fonts).unwrap();
        assert_eq!(entries.len(), 1);

        match &entries[0] {
            SharedStringEntry::Rich(runs) => {
                assert_eq!(runs.len(), 2);
                assert_eq!(runs[0].text, "Hello ");
                assert!(runs[0].font.is_none());
                assert_eq!(runs[1].text, "World");
                let f = runs[1].font.as_ref().unwrap();
                assert_eq!(f.bold, Some(true));
                assert_eq!(f.name, Some("Arial".to_string()));
                assert_eq!(f.size, Some(14.0));
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }

    #[test]
    fn read_rich_text_single_run() {
        let bold_font = FontStyle {
            bold: true,
            ..FontStyle::default()
        };
        let fonts = vec![bold_font];

        let sst_data = build_sst_rich("Bold", &[(0, 0)]);
        let mut archive = make_sst_zip(&sst_data);
        let entries = read_shared_strings(&mut archive, &fonts).unwrap();

        match &entries[0] {
            SharedStringEntry::Rich(runs) => {
                assert_eq!(runs.len(), 1);
                assert_eq!(runs[0].text, "Bold");
                assert_eq!(runs[0].font.as_ref().unwrap().bold, Some(true));
            }
            other => panic!("Expected Rich, got {:?}", other),
        }
    }

    #[test]
    fn read_plain_flag_zero() {
        let sst_data = build_sst_plain(&["plain"]);
        let mut archive = make_sst_zip(&sst_data);
        let entries = read_shared_strings(&mut archive, &[]).unwrap();
        assert!(matches!(&entries[0], SharedStringEntry::Plain(s) if s == "plain"));
    }
}
