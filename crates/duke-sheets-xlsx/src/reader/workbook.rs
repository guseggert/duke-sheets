//! Workbook-level XLSX reader helpers (workbook.xml, workbook.xml.rels, sheet .rels).

use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};

/// Parsed workbook properties from workbook.xml
pub(super) struct WorkbookProps {
    pub(super) sheets: Vec<(String, String)>,
    pub(super) date_1904: bool,
    pub(super) named_ranges: Vec<duke_sheets_core::named_range::NamedRange>,
}

pub(super) struct WorkbookRels {
    pub(super) sheet_paths: HashMap<String, String>,
    pub(super) theme_path: Option<String>,
}

#[derive(Debug, Clone)]
pub(super) struct SheetRelationship {
    pub(super) rel_type: String,
    pub(super) target: String,
}

/// Read workbook.xml to get sheet names, rIds, workbook properties,
/// and defined names.
pub(super) fn read_workbook_xml<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsxResult<WorkbookProps> {
    use duke_sheets_core::named_range::{NameScope, NamedRange};

    let file = archive
        .by_name("xl/workbook.xml")
        .map_err(|_| XlsxError::MissingPart("xl/workbook.xml".into()))?;

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut date_1904 = false;
    let mut named_ranges = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => {
                    parse_sheet_element(e, &mut sheets);
                }
                b"workbookPr" => {
                    parse_workbook_pr(e, &mut date_1904);
                }
                _ => {}
            },
            Ok(Event::Start(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => {
                    parse_sheet_element(e, &mut sheets);
                }
                b"workbookPr" => {
                    parse_workbook_pr(e, &mut date_1904);
                }
                b"definedName" => {
                    // Parse attributes
                    let mut dn_name = None;
                    let mut local_sheet_id: Option<usize> = None;
                    let mut comment = None;
                    let mut hidden = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"name" => {
                                dn_name = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"localSheetId" => {
                                local_sheet_id =
                                    attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                            b"comment" => {
                                comment = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"hidden" => {
                                hidden = attr.unescape_value().ok().map_or(false, |v| {
                                    v.as_ref() == "1" || v.eq_ignore_ascii_case("true")
                                });
                            }
                            _ => {}
                        }
                    }

                    // Read the text content (the refers_to expression)
                    let mut text_buf = Vec::new();
                    let refers_to = match xml_reader.read_event_into(&mut text_buf) {
                        Ok(Event::Text(t)) => t.unescape().ok().map(|s| s.to_string()),
                        _ => None,
                    };

                    if let (Some(name), Some(refers_to)) = (dn_name, refers_to) {
                        let scope = match local_sheet_id {
                            Some(idx) => NameScope::Sheet(idx),
                            None => NameScope::Workbook,
                        };
                        let mut nr = NamedRange::new(name, refers_to, scope);
                        nr.comment = comment;
                        nr.hidden = hidden;
                        named_ranges.push(nr);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(WorkbookProps {
        sheets,
        date_1904,
        named_ranges,
    })
}

fn parse_sheet_element(e: &quick_xml::events::BytesStart<'_>, sheets: &mut Vec<(String, String)>) {
    let mut name = None;
    let mut r_id = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"name" => {
                name = attr.unescape_value().ok().map(|s| s.to_string());
            }
            b"id" => {
                r_id = attr.unescape_value().ok().map(|s| s.to_string());
            }
            _ => {}
        }
    }

    if let (Some(name), Some(r_id)) = (name, r_id) {
        sheets.push((name, r_id));
    }
}

fn parse_workbook_pr(e: &quick_xml::events::BytesStart<'_>, date_1904: &mut bool) {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"date1904" {
            if let Ok(val) = attr.unescape_value() {
                *date_1904 = val.as_ref() == "1" || val.eq_ignore_ascii_case("true");
            }
        }
    }
}

/// Read workbook.xml.rels to get sheet file paths and theme path.
pub(super) fn read_workbook_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsxResult<WorkbookRels> {
    let file = archive
        .by_name("xl/_rels/workbook.xml.rels")
        .map_err(|_| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".into()))?;

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = HashMap::new();
    let mut theme_path: Option<String> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.name().local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut rel_type = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"Id" => {
                            id = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        b"Target" => {
                            target = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        b"Type" => {
                            rel_type = attr.unescape_value().ok().map(|s| s.to_string());
                        }
                        _ => {}
                    }
                }

                // Include worksheet relationships and theme relationship
                if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                    if rel_type.ends_with("/worksheet") {
                        // Target is relative to xl/ folder
                        let full_path = if target.starts_with('/') {
                            target[1..].to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        rels.insert(id, full_path);
                    } else if rel_type.ends_with("/theme") {
                        let full_path = if target.starts_with('/') {
                            target[1..].to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        theme_path = Some(full_path);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(WorkbookRels {
        sheet_paths: rels,
        theme_path,
    })
}

/// Read per-sheet .rels to get hyperlinks, comments, tables, etc.
pub(super) fn read_sheet_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
) -> XlsxResult<HashMap<String, SheetRelationship>> {
    let (base_dir, file_name) = match sheet_path.rsplit_once('/') {
        Some((dir, file)) => (dir, file),
        None => return Ok(HashMap::new()),
    };
    let rels_path = format!("{}/_rels/{}.rels", base_dir, file_name);

    let file = match archive.by_name(&rels_path) {
        Ok(f) => f,
        Err(_) => return Ok(HashMap::new()),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = HashMap::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.name().local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut rel_type = None;
                let mut target_mode = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"Id" => id = attr.unescape_value().ok().map(|s| s.to_string()),
                        b"Target" => target = attr.unescape_value().ok().map(|s| s.to_string()),
                        b"Type" => rel_type = attr.unescape_value().ok().map(|s| s.to_string()),
                        b"TargetMode" => {
                            target_mode = attr.unescape_value().ok().map(|s| s.to_string())
                        }
                        _ => {}
                    }
                }

                if let (Some(id), Some(target), Some(rel_type)) = (id, target, rel_type) {
                    let resolved_target =
                        if target.starts_with('/') || target_mode.as_deref() == Some("External") {
                            target
                        } else {
                            let mut parts: Vec<&str> = base_dir.split('/').collect();
                            for part in target.split('/') {
                                if part == ".." {
                                    parts.pop();
                                } else if part != "." && !part.is_empty() {
                                    parts.push(part);
                                }
                            }
                            parts.join("/")
                        };

                    rels.insert(
                        id,
                        SheetRelationship {
                            rel_type,
                            target: resolved_target,
                        },
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}
