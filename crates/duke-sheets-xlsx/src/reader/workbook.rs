//! Workbook-level XLSX reader helpers (workbook.xml, workbook.xml.rels, sheet .rels).

use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use crate::opc::{read_relationships, Relationship};
use duke_sheets_core::{SheetVisibility, WorkbookProtection};

/// Parsed workbook properties from workbook.xml
pub(super) struct WorkbookProps {
    pub(super) sheets: Vec<SheetEntry>,
    pub(super) date_1904: bool,
    pub(super) workbook_protection: Option<WorkbookProtection>,
    pub(super) named_ranges: Vec<duke_sheets_core::named_range::NamedRange>,
}

pub(super) struct WorkbookRels {
    pub(super) sheet_paths: HashMap<String, String>,
    pub(super) chartsheet_paths: HashMap<String, String>,
    pub(super) theme_path: Option<String>,
}

pub(super) type PartRelationship = Relationship;

pub(super) struct SheetEntry {
    pub(super) name: String,
    pub(super) r_id: String,
    pub(super) visibility: SheetVisibility,
}

/// Read workbook.xml to get sheet names, rIds, workbook properties,
/// and defined names.
pub(super) fn read_workbook_xml<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsxResult<WorkbookProps> {
    use duke_sheets_core::named_range::{NameScope, NamedRange};

    let file = archive_by_name(archive, "xl/workbook.xml")
        .map_err(|_| XlsxError::MissingPart("xl/workbook.xml".into()))?;

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut date_1904 = false;
    let mut workbook_protection = None;
    let mut named_ranges = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => parse_sheet_element(e, &mut sheets),
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"workbookProtection" => {
                    workbook_protection = parse_workbook_protection(e);
                }
                _ => {}
            },
            Ok(Event::Start(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => parse_sheet_element(e, &mut sheets),
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"workbookProtection" => {
                    workbook_protection = parse_workbook_protection(e);
                }
                b"definedName" => {
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
                                hidden = attr.unescape_value().ok().is_some_and(|v| {
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
        workbook_protection,
        named_ranges,
    })
}

fn parse_sheet_element(e: &quick_xml::events::BytesStart<'_>, sheets: &mut Vec<SheetEntry>) {
    let mut name = None;
    let mut r_id = None;
    let mut visibility = SheetVisibility::Visible;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"name" => name = attr.unescape_value().ok().map(|s| s.to_string()),
            b"id" => r_id = attr.unescape_value().ok().map(|s| s.to_string()),
            b"state" => {
                if let Ok(val) = attr.unescape_value() {
                    visibility = match val.as_ref() {
                        "hidden" => SheetVisibility::Hidden,
                        "veryHidden" => SheetVisibility::VeryHidden,
                        _ => SheetVisibility::Visible,
                    };
                }
            }
            _ => {}
        }
    }

    if let (Some(name), Some(r_id)) = (name, r_id) {
        sheets.push(SheetEntry {
            name,
            r_id,
            visibility,
        });
    }
}

fn parse_workbook_pr(e: &quick_xml::events::BytesStart<'_>, date_1904: &mut bool) {
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"date1904" => {
                if let Ok(val) = attr.unescape_value() {
                    *date_1904 = val.as_ref() == "1" || val.eq_ignore_ascii_case("true");
                }
            }
            _ => {}
        }
    }
}

fn parse_workbook_protection(e: &quick_xml::events::BytesStart<'_>) -> Option<WorkbookProtection> {
    let mut protection = WorkbookProtection::default();

    for attr in e.attributes().flatten() {
        let Ok(value) = attr.unescape_value() else {
            continue;
        };
        match attr.key.local_name().as_ref() {
            b"lockStructure" => {
                protection.structure =
                    value.as_ref() == "1" || value.as_ref().eq_ignore_ascii_case("true");
            }
            b"lockWindows" => {
                protection.windows =
                    value.as_ref() == "1" || value.as_ref().eq_ignore_ascii_case("true");
            }
            b"workbookPassword" => {
                if let Ok(h) = u16::from_str_radix(value.as_ref(), 16) {
                    protection.password_hash = Some(h);
                }
            }
            _ => {}
        }
    }

    if protection.structure || protection.windows || protection.password_hash.is_some() {
        Some(protection)
    } else {
        None
    }
}

/// Read workbook.xml.rels to get sheet file paths and theme path.
pub(super) fn read_workbook_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsxResult<WorkbookRels> {
    let relationships = read_relationships(archive, "xl/workbook.xml", true)?;
    let mut rels = HashMap::new();
    let mut chartsheet_rels = HashMap::new();
    let mut theme_path: Option<String> = None;
    for (id, relationship) in relationships {
        let Some(path) = relationship.internal_path() else {
            continue;
        };
        if relationship.rel_type.ends_with("/worksheet") {
            rels.insert(id, path.to_string());
        } else if relationship.rel_type.ends_with("/chartsheet") {
            chartsheet_rels.insert(id, path.to_string());
        } else if relationship.rel_type.ends_with("/theme") {
            theme_path = Some(path.to_string());
        }
    }

    Ok(WorkbookRels {
        sheet_paths: rels,
        chartsheet_paths: chartsheet_rels,
        theme_path,
    })
}

/// Read relationships owned by any package part.
pub(super) fn read_part_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    part_path: &str,
) -> XlsxResult<HashMap<String, PartRelationship>> {
    read_relationships(archive, part_path, false)
}
