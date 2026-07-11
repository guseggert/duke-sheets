use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use duke_sheets_formula::decompile::{
    ExternName, ExternSheetEntry, FormulaContext, NameRecord, SupBook, BUILTIN_NAMES,
};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::biff12::parser;
use crate::biff12::records;
use crate::biff12::token_parser;
use crate::biff12::RecordIter;
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::WorkbookProtection;

#[derive(Debug)]
pub(crate) struct SheetEntry {
    pub name: String,
    pub path: String,
    pub visibility: u32,
}

pub(crate) struct PrintSetting {
    pub sheet_idx: u32,
    pub refers_to: String,
}

pub(crate) struct WorkbookProps {
    pub sheets: Vec<SheetEntry>,
    pub date_1904: bool,
    pub active_sheet: usize,
    pub workbook_protection: Option<WorkbookProtection>,
    pub formula_ctx: FormulaContext,
    pub named_ranges: Vec<(String, u32, String, bool, Option<String>)>,
    pub print_areas: Vec<PrintSetting>,
    pub print_titles: Vec<PrintSetting>,
}

pub(crate) struct WorkbookRelationships {
    pub targets: HashMap<String, String>,
    pub theme_path: Option<String>,
}

pub(crate) fn read_relationships<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsbResult<WorkbookRelationships> {
    let file = match archive.by_name("xl/_rels/workbook.bin.rels") {
        Ok(f) => f,
        Err(_) => {
            return Ok(WorkbookRelationships {
                targets: HashMap::new(),
                theme_path: None,
            })
        }
    };
    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut targets = HashMap::new();
    let mut theme_path = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = String::new();
                let mut target = String::new();
                let mut rel_type = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Target" => target = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Type" => rel_type = String::from_utf8_lossy(&attr.value).into_owned(),
                        _ => {}
                    }
                }
                if rel_type.ends_with("/theme") {
                    let path = if target.starts_with("xl/") || target.starts_with("/xl/") {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("xl/{}", target)
                    };
                    theme_path = Some(path);
                }
                if !id.is_empty() && !target.is_empty() {
                    targets.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsbError::Parse(format!("XML error in .rels: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(WorkbookRelationships {
        targets,
        theme_path,
    })
}

pub(crate) fn read_workbook<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    relationships: &WorkbookRelationships,
) -> XlsbResult<WorkbookProps> {
    let file = archive
        .by_name("xl/workbook.bin")
        .map_err(|_| XlsbError::InvalidFormat("missing xl/workbook.bin".into()))?;
    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);
    let mut sheets = Vec::new();
    let mut date_1904 = false;
    let mut active_sheet: usize = 0;
    let mut workbook_protection = None;
    let mut extern_sheet_entries = Vec::new();
    let mut names = Vec::new();
    let mut supbooks: Vec<SupBook> = Vec::new();
    let mut extern_names: Vec<ExternName> = Vec::new();
    let mut in_sup_book = false;
    let mut current_sup_book: Option<SupBook> = None;
    let mut last_sup_book_idx: Option<u16> = None;

    let mut name_hidden_flags: Vec<bool> = Vec::new();

    let mut past_bundle_shs = false;
    loop {
        let record = iter.next_record(&mut buf);
        let (typ, len) = match record {
            Ok(r) => r,
            Err(e) => {
                if past_bundle_shs {
                    break;
                }
                return Err(e);
            }
        };
        match typ {
            records::BRT_WB_PROP => {
                if len >= 1 {
                    date_1904 = (buf[0] & 0x01) != 0;
                }
            }
            records::BRT_BOOK_PROTECTION => {
                if len >= 6 {
                    let password_hash = parser::read_u16(&buf, 0);
                    let flags = parser::read_u16(&buf, 4);
                    let protection = WorkbookProtection {
                        structure: (flags & 0x0001) != 0,
                        windows: (flags & 0x0002) != 0,
                        password_hash: if password_hash != 0 {
                            Some(password_hash)
                        } else {
                            None
                        },
                    };
                    if protection.structure
                        || protection.windows
                        || protection.password_hash.is_some()
                    {
                        workbook_protection = Some(protection);
                    }
                }
            }
            0x009E => {
                if len >= 28 {
                    active_sheet = parser::read_u32(&buf, 20) as usize;
                }
            }
            records::BRT_BUNDLE_SH => {
                if len < 12 {
                    continue;
                }
                let visibility = parser::read_u32(&buf, 0);
                let rel_char_count = parser::read_u32(&buf, 8) as usize;
                let rel_byte_len = rel_char_count * 2;
                let rel_start = 12;
                let rel_end = rel_start + rel_byte_len;
                if rel_end > len {
                    continue;
                }

                let (rel_id, _, _) = encoding_rs::UTF_16LE.decode(&buf[rel_start..rel_end]);

                let (sheet_name, _) = parser::wide_str(&buf, rel_end)?;

                let target = relationships
                    .targets
                    .get(rel_id.as_ref())
                    .map(|t| format!("xl/{}", t))
                    .unwrap_or_default();

                sheets.push(SheetEntry {
                    name: sheet_name,
                    path: target,
                    visibility,
                });
            }
            records::BRT_END_BUNDLE_SHS => past_bundle_shs = true,
            records::BRT_EXTERN_SHEET => {
                parse_extern_sheet(&buf[..len], &mut extern_sheet_entries);
            }
            records::BRT_NAME => {
                let hidden = len >= 4 && (parser::read_u32(&buf, 0) & 0x01) != 0;
                if let Some(nr) = parse_name_record(&buf[..len]) {
                    names.push(nr);
                    name_hidden_flags.push(hidden);
                }
            }
            records::BRT_BEGIN_SUP_BOOK => {
                in_sup_book = true;
                current_sup_book = None;
            }
            records::BRT_SUP_SELF => {
                if in_sup_book {
                    current_sup_book = Some(SupBook::SelfRef { sheet_count: 0 });
                } else {
                    last_sup_book_idx = Some(supbooks.len() as u16);
                    supbooks.push(SupBook::SelfRef { sheet_count: 0 });
                }
            }
            records::BRT_SUP_ADDIN => {
                if in_sup_book {
                    current_sup_book = Some(SupBook::AddIn);
                } else {
                    last_sup_book_idx = Some(supbooks.len() as u16);
                    supbooks.push(SupBook::AddIn);
                }
            }
            records::BRT_PLACEHOLDER_NAME => {
                let supbook_idx = if in_sup_book && current_sup_book.is_some() {
                    Some(supbooks.len() as u16)
                } else {
                    last_sup_book_idx
                };
                if let Some(supbook_idx) = supbook_idx {
                    let name = parser::wide_str(&buf, 0)
                        .map(|(s, _)| s)
                        .unwrap_or_default();
                    if !name.is_empty() {
                        extern_names.push(ExternName { supbook_idx, name });
                    }
                }
            }
            records::BRT_SUP_BOOK_SRC => {
                if in_sup_book && len >= 4 {
                    let path = parser::wide_str(&buf, 0)
                        .map(|(s, _)| s)
                        .unwrap_or_default();
                    current_sup_book = Some(SupBook::External {
                        path,
                        sheets: Vec::new(),
                    });
                }
            }
            records::BRT_END_SUP_BOOK => {
                if let Some(sb) = current_sup_book.take() {
                    last_sup_book_idx = Some(supbooks.len() as u16);
                    supbooks.push(sb);
                }
                in_sup_book = false;
            }
            0x0084 => break,
            _ => {}
        }
    }

    let sheet_names: Vec<String> = sheets.iter().map(|s| s.name.clone()).collect();
    let sheet_count = sheet_names.len() as u16;

    if supbooks.is_empty() {
        supbooks.push(SupBook::SelfRef { sheet_count });
    }
    for sb in &mut supbooks {
        if let SupBook::SelfRef { sheet_count: sc } = sb {
            *sc = sheet_count;
        }
    }

    let formula_ctx = FormulaContext {
        sheet_names,
        extern_sheet: extern_sheet_entries,
        supbooks,
        names,
        extern_names,
        extern_name_index_base: 0,
        base_cell: None,
    };

    let mut named_ranges = Vec::new();
    let mut print_areas = Vec::new();
    let mut print_titles = Vec::new();
    for (i, nr) in formula_ctx.names.iter().enumerate() {
        if nr.name.starts_with("_xlfn.") || nr.formula_body.is_empty() {
            continue;
        }
        let tokens = token_parser::parse_tokens(&nr.formula_body);
        if tokens.is_empty() {
            continue;
        }
        let refers_to =
            duke_sheets_formula::decompile::decompiler::decompile(&tokens, &formula_ctx);
        if refers_to.is_empty() {
            continue;
        }

        if nr.name == "Print_Area" || nr.name == "_xlnm.Print_Area" {
            print_areas.push(PrintSetting {
                sheet_idx: nr.sheet_idx,
                refers_to,
            });
            continue;
        }
        if nr.name == "Print_Titles" || nr.name == "_xlnm.Print_Titles" {
            print_titles.push(PrintSetting {
                sheet_idx: nr.sheet_idx,
                refers_to,
            });
            continue;
        }
        if nr.is_builtin {
            continue;
        }

        let hidden = name_hidden_flags.get(i).copied().unwrap_or(false);
        named_ranges.push((
            nr.name.clone(),
            nr.sheet_idx,
            refers_to,
            hidden,
            nr.comment.clone(),
        ));
    }

    Ok(WorkbookProps {
        sheets,
        date_1904,
        active_sheet,
        workbook_protection,
        formula_ctx,
        named_ranges,
        print_areas,
        print_titles,
    })
}

fn parse_extern_sheet(data: &[u8], entries: &mut Vec<ExternSheetEntry>) {
    if data.len() < 4 {
        return;
    }
    let count = parser::read_u32(data, 0) as usize;
    let mut pos = 4;
    for _ in 0..count {
        // Each entry: supBookIdx(u32) + firstSheet(u32) + lastSheet(u32) = 12 bytes
        if pos + 12 > data.len() {
            break;
        }
        let sup_book_idx = parser::read_u32(data, pos) as u16;
        let first_sheet = parser::read_u32(data, pos + 4) as u16;
        let last_sheet = parser::read_u32(data, pos + 8) as u16;
        pos += 12;
        entries.push(ExternSheetEntry {
            sup_book_idx,
            first_sheet,
            last_sheet,
        });
    }
}

fn parse_name_record(data: &[u8]) -> Option<NameRecord> {
    // BrtName layout: flags(u32) + chKey(u8) + itab(u32) + name(XLWideString) + formula
    if data.len() < 13 {
        return None;
    }

    let flags = parser::read_u32(data, 0);
    let is_builtin = (flags & 0x20) != 0;
    let itab = parser::read_u32(data, 5);

    let mut pos = 9;
    let (raw_name, consumed) = parser::wide_str(data, pos).ok()?;
    pos += consumed;

    let name = if is_builtin && raw_name.len() == 1 {
        let idx = raw_name.as_bytes()[0] as usize;
        BUILTIN_NAMES
            .get(idx)
            .map(|s| s.to_string())
            .unwrap_or(raw_name)
    } else {
        raw_name
    };

    let mut formula_body = Vec::new();
    if pos + 4 <= data.len() {
        let cce = parser::read_u32(data, pos) as usize;
        pos += 4;
        if pos + cce <= data.len() {
            formula_body.extend_from_slice(&data[pos..pos + cce]);
            pos += cce;
        }
        if pos + 4 <= data.len() {
            let cb = parser::read_u32(data, pos) as usize;
            pos += 4;
            if cb > 0 && pos + cb <= data.len() {
                formula_body.extend_from_slice(&data[pos..pos + cb]);
                pos += cb;
            }
        }
    }

    // Trailing strings per [MS-XLSB] §2.4.668: comment, customMenu,
    // description, help, statusBar — each XLNullableWideString.
    let comment = read_nullable_wide_str_at(data, &mut pos);

    Some(NameRecord {
        name,
        sheet_idx: itab,
        is_builtin,
        formula_body,
        comment,
    })
}

/// Pull an XLNullableWideString starting at *pos. NULL marker
/// (0xFFFFFFFF) consumes 4 bytes and returns None; otherwise advance
/// past the encoded XLWideString and return the string.
fn read_nullable_wide_str_at(data: &[u8], pos: &mut usize) -> Option<String> {
    if *pos + 4 > data.len() {
        return None;
    }
    let marker = parser::read_u32(data, *pos);
    if marker == 0xFFFFFFFF {
        *pos += 4;
        return None;
    }
    let (s, consumed) = parser::wide_str(data, *pos).ok()?;
    *pos += consumed;
    if s.is_empty() {
        None
    } else {
        Some(s)
    }
}
