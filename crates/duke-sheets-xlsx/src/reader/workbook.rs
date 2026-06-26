//! Workbook-level XLSX reader helpers (workbook.xml, workbook.xml.rels, sheet .rels).

use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::archive_by_name;
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::{SheetVisibility, WorkbookConnection, WorkbookConnectionKind};

/// Parsed workbook properties from workbook.xml
pub(super) struct WorkbookProps {
    pub(super) sheets: Vec<SheetEntry>,
    pub(super) date_1904: bool,
    pub(super) named_ranges: Vec<duke_sheets_core::named_range::NamedRange>,
    pub(super) pivot_caches: Vec<PivotCacheEntry>,
}

pub(super) struct WorkbookRels {
    pub(super) sheet_paths: HashMap<String, String>,
    pub(super) chartsheet_paths: HashMap<String, String>,
    pub(super) theme_path: Option<String>,
    pub(super) pivot_cache_paths: HashMap<String, String>,
    pub(super) connections_path: Option<String>,
}

#[derive(Debug, Clone)]
pub(super) struct SheetRelationship {
    pub(super) rel_type: String,
    pub(super) target: String,
}

pub(super) struct SheetEntry {
    pub(super) name: String,
    pub(super) r_id: String,
    pub(super) visibility: SheetVisibility,
}

pub(super) struct PivotCacheEntry {
    pub(super) cache_id: u32,
    pub(super) r_id: String,
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
    let mut named_ranges = Vec::new();
    let mut pivot_caches = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => parse_sheet_element(e, &mut sheets),
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"pivotCache" => parse_pivot_cache_element(e, &mut pivot_caches),
                _ => {}
            },
            Ok(Event::Start(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => parse_sheet_element(e, &mut sheets),
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"pivotCache" => parse_pivot_cache_element(e, &mut pivot_caches),
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
        named_ranges,
        pivot_caches,
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

fn parse_pivot_cache_element(
    e: &quick_xml::events::BytesStart<'_>,
    pivot_caches: &mut Vec<PivotCacheEntry>,
) {
    let mut cache_id = None;
    let mut r_id = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"cacheId" => {
                cache_id = attr.unescape_value().ok().and_then(|s| s.parse().ok());
            }
            b"id" => r_id = attr.unescape_value().ok().map(|s| s.to_string()),
            _ => {}
        }
    }

    if let (Some(cache_id), Some(r_id)) = (cache_id, r_id) {
        pivot_caches.push(PivotCacheEntry { cache_id, r_id });
    }
}

/// Read workbook.xml.rels to get sheet file paths and theme path.
pub(super) fn read_workbook_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> XlsxResult<WorkbookRels> {
    let file = archive_by_name(archive, "xl/_rels/workbook.xml.rels")
        .map_err(|_| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".into()))?;

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = HashMap::new();
    let mut chartsheet_rels = HashMap::new();
    let mut theme_path: Option<String> = None;
    let mut pivot_cache_paths = HashMap::new();
    let mut connections_path = None;
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
                        b"Id" => id = attr.unescape_value().ok().map(|s| s.to_string()),
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
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        rels.insert(id, full_path);
                    } else if rel_type.ends_with("/chartsheet") {
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        chartsheet_rels.insert(id, full_path);
                    } else if rel_type.ends_with("/theme") {
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        theme_path = Some(full_path);
                    } else if rel_type.ends_with("/pivotCacheDefinition") {
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        pivot_cache_paths.insert(id, full_path);
                    } else if rel_type.ends_with("/connections") {
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("xl/{}", target)
                        };
                        connections_path = Some(full_path);
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
        chartsheet_paths: chartsheet_rels,
        theme_path,
        pivot_cache_paths,
        connections_path,
    })
}

pub(super) fn read_workbook_connections<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: Option<&str>,
) -> XlsxResult<Vec<WorkbookConnection>> {
    let Some(path) = path else {
        return Ok(Vec::new());
    };
    let file = match archive_by_name(archive, path) {
        Ok(file) => file,
        Err(_) => return Ok(Vec::new()),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut connections = Vec::new();
    let mut current: Option<ParsedConnection> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"connection" => current = parse_connection_attrs(&e),
                b"dbPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = parse_db_pr(&e);
                    }
                }
                b"olapPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_olap_pr(&e));
                    }
                }
                b"webPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_web_pr(&e));
                    }
                }
                b"textPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_text_pr(&e));
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"connection" => {
                    if let Some(connection) = parse_connection_attrs(&e).and_then(|c| c.build()) {
                        connections.push(connection);
                    }
                }
                b"dbPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = parse_db_pr(&e);
                    }
                }
                b"olapPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_olap_pr(&e));
                    }
                }
                b"webPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_web_pr(&e));
                    }
                }
                b"textPr" => {
                    if let Some(connection) = &mut current {
                        connection.kind = Some(parse_text_pr(&e));
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().local_name().as_ref() == b"connection" {
                    if let Some(connection) = current.take().and_then(|c| c.build()) {
                        connections.push(connection);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(connections)
}

struct ParsedConnection {
    id: u32,
    name: String,
    refreshed_version: u8,
    refresh_on_load: bool,
    background: bool,
    save_data: bool,
    kind: Option<WorkbookConnectionKind>,
}

impl ParsedConnection {
    fn build(self) -> Option<WorkbookConnection> {
        Some(WorkbookConnection {
            id: self.id,
            name: self.name,
            kind: self.kind?,
            refreshed_version: self.refreshed_version,
            refresh_on_load: self.refresh_on_load,
            background: self.background,
            save_data: self.save_data,
        })
    }
}

fn parse_connection_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<ParsedConnection> {
    Some(ParsedConnection {
        id: attr_u32(e, b"id")?,
        name: attr_string(e, b"name").unwrap_or_default(),
        refreshed_version: attr_u32(e, b"refreshedVersion")
            .and_then(|value| u8::try_from(value).ok())
            .unwrap_or(0),
        refresh_on_load: attr_bool(e, b"refreshOnLoad").unwrap_or(false),
        background: attr_bool(e, b"background").unwrap_or(false),
        save_data: attr_bool(e, b"saveData").unwrap_or(false),
        kind: None,
    })
}

fn parse_db_pr(e: &quick_xml::events::BytesStart<'_>) -> Option<WorkbookConnectionKind> {
    Some(WorkbookConnectionKind::Database {
        connection: attr_string(e, b"connection")?,
        command: attr_string(e, b"command"),
        command_type: attr_u32(e, b"commandType"),
    })
}

fn parse_olap_pr(e: &quick_xml::events::BytesStart<'_>) -> WorkbookConnectionKind {
    WorkbookConnectionKind::Olap {
        local: attr_bool(e, b"local").unwrap_or(false),
        local_connection: attr_string(e, b"localConnection"),
        local_refresh: attr_bool(e, b"localRefresh").unwrap_or(true),
        send_locale: attr_bool(e, b"sendLocale").unwrap_or(false),
        row_drill_count: attr_u32(e, b"rowDrillCount"),
    }
}

fn parse_web_pr(e: &quick_xml::events::BytesStart<'_>) -> WorkbookConnectionKind {
    WorkbookConnectionKind::Web {
        url: attr_string(e, b"url"),
        xml: attr_bool(e, b"xml").unwrap_or(false),
        source_data: attr_bool(e, b"sourceData").unwrap_or(false),
        html_tables: attr_bool(e, b"htmlTables").unwrap_or(false),
        html_format: attr_string(e, b"htmlFormat"),
        post: attr_string(e, b"post"),
        edit_page: attr_string(e, b"editPage"),
    }
}

fn parse_text_pr(e: &quick_xml::events::BytesStart<'_>) -> WorkbookConnectionKind {
    WorkbookConnectionKind::Text {
        source_file: attr_string(e, b"sourceFile"),
        delimiter: attr_string(e, b"delimiter"),
        first_row: attr_u32(e, b"firstRow").unwrap_or(1),
        delimited: attr_bool(e, b"delimited").unwrap_or(true),
        decimal: attr_string(e, b"decimal"),
        thousands: attr_string(e, b"thousands"),
    }
}

fn attr_string(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    e.attributes().flatten().find_map(|attr| {
        (attr.key.local_name().as_ref() == key)
            .then(|| attr.unescape_value().ok().map(|s| s.to_string()))
            .flatten()
    })
}

fn attr_u32(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<u32> {
    attr_string(e, key).and_then(|value| value.parse().ok())
}

fn attr_bool(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<bool> {
    attr_string(e, key).map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
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

    let file = match archive_by_name(archive, &rels_path) {
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
