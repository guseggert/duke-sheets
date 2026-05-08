use std::collections::HashMap;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::{Color, FontStyle, FontVerticalAlign, Underline};
use duke_sheets_core::{CellValue, Workbook};

pub(crate) enum SstEntry {
    Plain(String),
    Rich { text: String, runs: Vec<SstRichRun> },
}

pub(crate) struct SstRichRun {
    /// Character index into the parent XLString. Bounded to u16 by
    /// the on-disk format ([MS-XLSB] §2.5.157 StrRun).
    pub ich: u16,
    pub font_idx: Option<usize>,
}

pub(crate) struct SstMap {
    entries: Vec<SstEntry>,
    plain_index: HashMap<String, u32>,
    rich_keys: HashMap<u64, u32>,
    pub(crate) rich_fonts: Vec<FontStyle>,
    _font_dedup: HashMap<u64, usize>,
}

impl SstMap {
    pub fn get_plain(&self, s: &str) -> Option<u32> {
        self.plain_index.get(s).copied()
    }

    pub fn get_rich(&self, runs: &[RichTextRun]) -> Option<u32> {
        let key = hash_rich_text(runs);
        self.rich_keys.get(&key).copied()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

fn hash_rich_text(runs: &[RichTextRun]) -> u64 {
    use std::hash::{Hash, Hasher};
    let mut h = std::collections::hash_map::DefaultHasher::new();
    runs.len().hash(&mut h);
    for run in runs {
        run.text.hash(&mut h);
        match &run.font {
            None => 0u8.hash(&mut h),
            Some(f) => {
                1u8.hash(&mut h);
                f.bold.hash(&mut h);
                f.italic.hash(&mut h);
                f.strikethrough.hash(&mut h);
                f.size.map(|v| v.to_bits()).hash(&mut h);
                f.name.hash(&mut h);
                f.underline.hash(&mut h);
                f.vertical_align.hash(&mut h);
                f.family.hash(&mut h);
                f.charset.hash(&mut h);
                f.scheme.hash(&mut h);
                match &f.color {
                    None => 0u8.hash(&mut h),
                    Some(c) => {
                        1u8.hash(&mut h);
                        c.hash(&mut h);
                    }
                }
            }
        }
    }
    h.finish()
}

fn font_style_hash(f: &FontStyle) -> u64 {
    use std::hash::{Hash, Hasher};
    let mut h = std::collections::hash_map::DefaultHasher::new();
    f.hash(&mut h);
    h.finish()
}

fn run_font_to_font_style(rf: &RunFont) -> FontStyle {
    FontStyle {
        bold: rf.bold.unwrap_or(false),
        italic: rf.italic.unwrap_or(false),
        size: rf.size.unwrap_or(11.0),
        color: rf.color.clone().unwrap_or(Color::Auto),
        name: rf.name.clone().unwrap_or_else(|| "Calibri".to_string()),
        underline: rf.underline.unwrap_or(Underline::None),
        strikethrough: rf.strikethrough.unwrap_or(false),
        vertical_align: rf.vertical_align.unwrap_or(FontVerticalAlign::Baseline),
        family: rf.family,
        charset: rf.charset,
        scheme: rf.scheme.clone(),
    }
}

pub(crate) fn build_sst(workbook: &Workbook) -> SstMap {
    let mut entries = Vec::new();
    let mut plain_index = HashMap::new();
    let mut rich_keys: HashMap<u64, u32> = HashMap::new();
    let mut rich_fonts: Vec<FontStyle> = Vec::new();
    let mut font_dedup: HashMap<u64, usize> = HashMap::new();

    for sheet in workbook.worksheets() {
        for (_row, _col, cell) in sheet.iter_cells() {
            match &cell.value {
                CellValue::String(s) => {
                    let s_str = s.as_str();
                    if !plain_index.contains_key(s_str) {
                        let idx = entries.len() as u32;
                        plain_index.insert(s_str.to_owned(), idx);
                        entries.push(SstEntry::Plain(s_str.to_owned()));
                    }
                }
                CellValue::RichText(runs) => {
                    let key = hash_rich_text(runs);
                    if rich_keys.contains_key(&key) {
                        continue;
                    }

                    let mut sst_runs = Vec::with_capacity(runs.len());
                    let mut current_cu: u32 = 0;
                    let mut full_text = String::new();

                    for run in runs.iter() {
                        let font_idx = match &run.font {
                            None => None,
                            Some(rf) => {
                                let fs = run_font_to_font_style(rf);
                                let fk = font_style_hash(&fs);
                                let idx = if let Some(&existing) = font_dedup.get(&fk) {
                                    existing
                                } else {
                                    let idx = rich_fonts.len();
                                    rich_fonts.push(fs);
                                    font_dedup.insert(fk, idx);
                                    idx
                                };
                                Some(idx)
                            }
                        };
                        // Per [MS-XLSB] §2.5.157 ich is u16. The parent
                        // XLString is itself capped at 32,767 chars, so
                        // legitimate input always fits; clamp loudly on
                        // pathological overflow rather than silently
                        // truncating.
                        let ich_u16 = u16::try_from(current_cu).unwrap_or_else(|_| {
                            log::warn!(
                                "rich text run ich {} exceeds u16; clamping to u16::MAX",
                                current_cu
                            );
                            u16::MAX
                        });
                        sst_runs.push(SstRichRun {
                            ich: ich_u16,
                            font_idx,
                        });
                        current_cu += run.text.encode_utf16().count() as u32;
                        full_text.push_str(&run.text);
                    }

                    let idx = entries.len() as u32;
                    rich_keys.insert(key, idx);
                    entries.push(SstEntry::Rich {
                        text: full_text,
                        runs: sst_runs,
                    });
                }
                _ => continue,
            }
        }
    }

    SstMap {
        entries,
        plain_index,
        rich_keys,
        rich_fonts,
        _font_dedup: font_dedup,
    }
}

pub(crate) fn write_sst<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    sst: &SstMap,
    rich_font_ids: &[u16],
) -> XlsbResult<()> {
    zip.start_file("xl/sharedStrings.bin", *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    let count = sst.entries.len() as u32;
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&count.to_le_bytes());
    payload.extend_from_slice(&count.to_le_bytes());
    rw.write_record(records::BRT_BEGIN_SST, &payload)?;

    for entry in &sst.entries {
        match entry {
            SstEntry::Plain(s) => {
                let mut item = Vec::new();
                item.push(0u8);
                item.extend_from_slice(&encode_wide_str(s));
                rw.write_record(records::BRT_SS_ITEM, &item)?;
            }
            SstEntry::Rich { text, runs } => {
                // RichStr per [MS-XLSB] §2.5.139:
                //   1 byte flags (bit 0 = fRichStr)
                //   XLWideString text
                //   u32 dwSizeStrRun
                //   StrRun × dwSizeStrRun, where each StrRun is
                //   u16 ich + u16 ifnt (4 bytes total) per §2.5.157.
                let mut item = Vec::new();
                item.push(0x01u8); // fRichStr
                item.extend_from_slice(&encode_wide_str(text));
                item.extend_from_slice(&(runs.len() as u32).to_le_bytes());
                for run in runs {
                    item.extend_from_slice(&run.ich.to_le_bytes());
                    let ifnt: u16 = match run.font_idx {
                        None => 0,
                        Some(idx) => rich_font_ids.get(idx).copied().unwrap_or(0),
                    };
                    item.extend_from_slice(&ifnt.to_le_bytes());
                }
                rw.write_record(records::BRT_SS_ITEM, &item)?;
            }
        }
    }

    rw.write_record(0x00A0, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}
