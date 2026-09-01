//! Workbook-level XLSX reader helpers (workbook.xml, workbook.xml.rels, sheet .rels).

use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read, Seek};
use std::sync::Arc;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::{NsReader, Reader};
use quick_xml::Writer;

use crate::error::{XlsxError, XlsxResult};
use crate::opc::{OpcPackage, PartName, RelationshipKind, RelationshipSet, RelationshipSource};
use crate::XlsxPackagePolicy;
use duke_sheets_core::{
    SheetVisibility, WorkbookConnection, WorkbookConnectionCredentials, WorkbookConnectionKind,
    WorkbookConnectionParameter, WorkbookConnectionParameterType, WorkbookConnectionParameterValue,
    WorkbookExtension, WorkbookProtection,
};

/// Parsed workbook properties from workbook.xml
pub(super) struct WorkbookProps {
    pub(super) sheets: Vec<SheetEntry>,
    pub(super) date_1904: bool,
    pub(super) workbook_protection: Option<WorkbookProtection>,
    pub(super) named_ranges: Vec<duke_sheets_core::named_range::NamedRange>,
    pub(super) pivot_caches: Vec<PivotCacheEntry>,
    pub(super) workbook_extensions: Vec<WorkbookExtension>,
}

pub(super) struct WorkbookRels {
    pub(super) sheet_paths: HashMap<String, String>,
    pub(super) chartsheet_paths: HashMap<String, String>,
    pub(super) theme_path: Option<String>,
    pub(super) styles_path: Option<String>,
    pub(super) shared_strings_path: Option<String>,
    pub(super) pivot_cache_paths: HashMap<String, String>,
    pub(super) connections_path: Option<String>,
    pub(super) extension_parts: Vec<WorkbookExtensionRelationship>,
    /// Relationship ids of valid but unmodeled sheet kinds
    /// (dialog/macro sheets); their sheet entries are skipped.
    pub(super) unmodeled_sheet_rels: std::collections::HashSet<String>,
}
pub(super) struct WorkbookExtensionRelationship {
    pub(super) r_id: String,
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
pub(super) fn read_workbook_xml<R: Read>(
    reader: R,
    policy: XlsxPackagePolicy,
) -> XlsxResult<WorkbookProps> {
    use duke_sheets_core::named_range::{NameScope, NamedRange};

    let mut xml_reader = NsReader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    // Compatible mode keeps the historical namespace-agnostic parse.
    if policy == XlsxPackagePolicy::Strict {
        loop {
            match xml_reader.read_resolved_event_into(&mut buf) {
                Ok((ResolveResult::Bound(namespace), Event::Start(element)))
                    if element.name().local_name().as_ref() == b"workbook"
                        && matches!(
                            namespace.as_ref(),
                            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                | b"http://purl.oclc.org/ooxml/spreadsheetml/main"
                        ) =>
                {
                    break;
                }
                Ok((_, Event::Decl(_) | Event::Comment(_) | Event::PI(_))) => {}
                Ok((_, Event::Text(text)))
                    if text.unescape().is_ok_and(|value| value.trim().is_empty()) => {}
                Ok((_, Event::Eof)) => {
                    return Err(XlsxError::InvalidFormat(
                        "Workbook part has no workbook root element".into(),
                    ));
                }
                Ok(_) => {
                    return Err(XlsxError::InvalidFormat(
                        "Workbook part has an invalid root element or namespace".into(),
                    ));
                }
                Err(error) => return Err(XlsxError::Xml(error)),
            }
            buf.clear();
        }
    }
    buf.clear();
    let mut sheets = Vec::new();
    let mut date_1904 = false;
    let mut workbook_protection = None;
    let mut named_ranges = Vec::new();
    let mut pivot_caches = Vec::new();
    let mut workbook_extensions = Vec::new();
    let mut in_ext_list = false;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => {
                    if !parse_sheet_element(e, &mut sheets) && policy == XlsxPackagePolicy::Strict {
                        return Err(XlsxError::InvalidFormat(
                            "Workbook sheet is missing name or relationship id".into(),
                        ));
                    }
                }
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"pivotCache" => parse_pivot_cache_element(e, &mut pivot_caches),
                b"ext" if in_ext_list => {
                    workbook_extensions.push(empty_workbook_extension(e)?);
                }
                b"workbookProtection" => {
                    workbook_protection = parse_workbook_protection(e);
                }
                _ => {}
            },
            Ok(Event::Start(ref e)) => match e.name().local_name().as_ref() {
                b"sheet" => {
                    if !parse_sheet_element(e, &mut sheets) && policy == XlsxPackagePolicy::Strict {
                        return Err(XlsxError::InvalidFormat(
                            "Workbook sheet is missing name or relationship id".into(),
                        ));
                    }
                }
                b"workbookPr" => parse_workbook_pr(e, &mut date_1904),
                b"pivotCache" => parse_pivot_cache_element(e, &mut pivot_caches),
                b"extLst" => in_ext_list = true,
                b"ext" if in_ext_list => {
                    let extension =
                        read_workbook_extension(&mut xml_reader, e.clone().into_owned(), &mut buf)?;
                    workbook_extensions.push(extension);
                }
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
            Ok(Event::End(ref e)) if e.name().local_name().as_ref() == b"extLst" => {
                in_ext_list = false;
            }
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
        pivot_caches,
        workbook_extensions,
    })
}

fn parse_sheet_element(
    e: &quick_xml::events::BytesStart<'_>,
    sheets: &mut Vec<SheetEntry>,
) -> bool {
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
        true
    } else {
        false
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

fn read_workbook_extension<B: std::io::BufRead>(
    xml_reader: &mut NsReader<B>,
    start: BytesStart<'static>,
    buf: &mut Vec<u8>,
) -> XlsxResult<WorkbookExtension> {
    let uri = attr_string(&start, b"uri").unwrap_or_default();
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Start(start))?;

    let mut depth = 1usize;
    loop {
        buf.clear();
        let event = xml_reader.read_event_into(buf)?;
        match event {
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::End(e) => {
                writer.write_event(Event::End(e.into_owned()))?;
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    break;
                }
            }
            Event::Empty(e) => writer.write_event(Event::Empty(e.into_owned()))?,
            Event::Text(e) => writer.write_event(Event::Text(e.into_owned()))?,
            Event::CData(e) => writer.write_event(Event::CData(e.into_owned()))?,
            Event::Comment(e) => writer.write_event(Event::Comment(e.into_owned()))?,
            Event::Decl(e) => writer.write_event(Event::Decl(e.into_owned()))?,
            Event::PI(e) => writer.write_event(Event::PI(e.into_owned()))?,
            Event::DocType(e) => writer.write_event(Event::DocType(e.into_owned()))?,
            Event::Eof => {
                return Err(XlsxError::InvalidFormat(
                    "unexpected EOF while reading workbook extension".into(),
                ))
            }
        }
    }

    Ok(WorkbookExtension {
        uri,
        payload: writer.into_inner().into_inner(),
    })
}

fn empty_workbook_extension(ext: &BytesStart<'_>) -> XlsxResult<WorkbookExtension> {
    let uri = attr_string(ext, b"uri").unwrap_or_default();
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Empty(ext.clone().into_owned()))?;
    Ok(WorkbookExtension {
        uri,
        payload: writer.into_inner().into_inner(),
    })
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
    package: &mut OpcPackage<R>,
    workbook_path: &PartName,
) -> XlsxResult<WorkbookRels> {
    let source = RelationshipSource::Part(workbook_path.clone());
    let relationships = package.relationships(&source, true)?;
    let mut rels = HashMap::new();
    let mut chartsheet_rels = HashMap::new();
    let mut theme_path = None;
    let mut styles_path = None;
    let mut shared_strings_path = None;
    let mut pivot_cache_paths = HashMap::new();
    let mut connections_path = None;
    let mut extension_parts = Vec::new();
    let mut unmodeled_sheet_rels = std::collections::HashSet::new();

    for relationship in relationships.iter() {
        let rel_type = relationship.rel_type.as_str();
        if rel_type.ends_with("/pivotCacheDefinition")
            || rel_type.ends_with("/connections")
            || is_workbook_extension_relationship(rel_type)
        {
            let Some(path) = relationship.internal_path() else {
                continue;
            };
            if package.open_related_part(relationship)?.is_none() {
                continue;
            }
            if rel_type.ends_with("/pivotCacheDefinition") {
                pivot_cache_paths.insert(relationship.id.clone(), path.to_string());
            } else if rel_type.ends_with("/connections") {
                connections_path = Some(path.to_string());
            } else {
                extension_parts.push(WorkbookExtensionRelationship {
                    r_id: relationship.id.clone(),
                    rel_type: relationship.rel_type.clone(),
                    target: path.to_string(),
                });
            }
            continue;
        }

        let Some(kind) = relationship.kind() else {
            continue;
        };
        // Valid OOXML sheet kinds this library does not model yet; a
        // capability limitation, never a conformance violation.
        if let Some(label) = kind.unmodeled_sheet_label() {
            if package.open_related_part(relationship)?.is_none() {
                continue;
            }
            package.diagnostics_mut().warning(
                crate::opc::XlsxDiagnosticCode::UnsupportedSheetType,
                format!("{label} sheets are not supported by the workbook model"),
                Some(workbook_path.as_str()),
                Some(&relationship.id),
                Some(&relationship.raw_target),
            );
            unmodeled_sheet_rels.insert(relationship.id.clone());
            continue;
        }
        if !matches!(
            kind,
            RelationshipKind::Worksheet
                | RelationshipKind::Chartsheet
                | RelationshipKind::Theme
                | RelationshipKind::Styles
                | RelationshipKind::SharedStrings
        ) {
            continue;
        }
        let Some(path) = relationship.internal_path() else {
            continue;
        };
        if package.open_related_part(relationship)?.is_none() {
            continue;
        }
        if kind == RelationshipKind::Worksheet {
            rels.insert(relationship.id.clone(), path.to_string());
        } else if kind == RelationshipKind::Chartsheet {
            chartsheet_rels.insert(relationship.id.clone(), path.to_string());
        } else if kind == RelationshipKind::Theme {
            theme_path = Some(path.to_string());
        } else if kind == RelationshipKind::Styles {
            styles_path = Some(path.to_string());
        } else if kind == RelationshipKind::SharedStrings {
            shared_strings_path = Some(path.to_string());
        }
    }

    Ok(WorkbookRels {
        sheet_paths: rels,
        chartsheet_paths: chartsheet_rels,
        theme_path,
        styles_path,
        shared_strings_path,
        pivot_cache_paths,
        connections_path,
        extension_parts,
        unmodeled_sheet_rels,
    })
}

fn is_workbook_extension_relationship(rel_type: &str) -> bool {
    rel_type.ends_with("/slicerCache") || rel_type.ends_with("/timelineCache")
}

pub(super) fn read_workbook_connections<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    path: Option<&str>,
) -> XlsxResult<Vec<WorkbookConnection>> {
    let Some(path) = path else {
        return Ok(Vec::new());
    };
    let file = match package.open_zip_name(path) {
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
                        connection.apply_db_pr(&e);
                    }
                }
                b"olapPr" => {
                    if let Some(connection) = &mut current {
                        connection.apply_olap_pr(&e);
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
                b"parameter" => {
                    if let Some(connection) = &mut current {
                        connection.parameters.push(parse_connection_parameter(&e));
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
                        connection.apply_db_pr(&e);
                    }
                }
                b"olapPr" => {
                    if let Some(connection) = &mut current {
                        connection.apply_olap_pr(&e);
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
                b"parameter" => {
                    if let Some(connection) = &mut current {
                        connection.parameters.push(parse_connection_parameter(&e));
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
    source_file: Option<String>,
    odc_file: Option<String>,
    description: Option<String>,
    connection_type: Option<u32>,
    refreshed_version: u8,
    min_refreshable_version: u8,
    keep_alive: bool,
    interval: u32,
    reconnection_method: u32,
    refresh_on_load: bool,
    background: bool,
    save_data: bool,
    save_password: bool,
    new_connection: bool,
    deleted: bool,
    only_use_connection_file: bool,
    credentials: Option<WorkbookConnectionCredentials>,
    single_sign_on_id: Option<String>,
    parameters: Vec<WorkbookConnectionParameter>,
    kind: Option<WorkbookConnectionKind>,
    db_props: Option<ConnectionDbProps>,
}

impl ParsedConnection {
    fn apply_db_pr(&mut self, e: &quick_xml::events::BytesStart<'_>) {
        let Some(db_props) = parse_db_pr(e) else {
            return;
        };
        self.kind = Some(match self.kind.take() {
            Some(WorkbookConnectionKind::Olap {
                local,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
                ..
            }) => db_props.olap_kind(
                local,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
            ),
            _ => db_props.database_kind(),
        });
        self.db_props = Some(db_props);
    }

    fn apply_olap_pr(&mut self, e: &quick_xml::events::BytesStart<'_>) {
        self.kind = Some(parse_olap_pr(e, self.db_props.as_ref()));
    }

    fn build(self) -> Option<WorkbookConnection> {
        Some(WorkbookConnection {
            id: self.id,
            name: self.name,
            source_file: self.source_file,
            odc_file: self.odc_file,
            description: self.description,
            connection_type: self.connection_type,
            kind: self.kind?,
            refreshed_version: self.refreshed_version,
            min_refreshable_version: self.min_refreshable_version,
            keep_alive: self.keep_alive,
            interval: self.interval,
            reconnection_method: self.reconnection_method,
            refresh_on_load: self.refresh_on_load,
            background: self.background,
            save_data: self.save_data,
            save_password: self.save_password,
            new_connection: self.new_connection,
            deleted: self.deleted,
            only_use_connection_file: self.only_use_connection_file,
            credentials: self.credentials,
            single_sign_on_id: self.single_sign_on_id,
            parameters: self.parameters,
        })
    }
}

fn parse_connection_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<ParsedConnection> {
    Some(ParsedConnection {
        id: attr_u32(e, b"id")?,
        name: attr_string(e, b"name").unwrap_or_default(),
        source_file: attr_string(e, b"sourceFile"),
        odc_file: attr_string(e, b"odcFile"),
        description: attr_string(e, b"description"),
        connection_type: attr_u32(e, b"type"),
        refreshed_version: attr_u32(e, b"refreshedVersion")
            .and_then(|value| u8::try_from(value).ok())
            .unwrap_or(0),
        min_refreshable_version: attr_u32(e, b"minRefreshableVersion")
            .and_then(|value| u8::try_from(value).ok())
            .unwrap_or(0),
        keep_alive: attr_bool(e, b"keepAlive").unwrap_or(false),
        interval: attr_u32(e, b"interval").unwrap_or(0),
        reconnection_method: attr_u32(e, b"reconnectionMethod").unwrap_or(1),
        refresh_on_load: attr_bool(e, b"refreshOnLoad").unwrap_or(false),
        background: attr_bool(e, b"background").unwrap_or(false),
        save_data: attr_bool(e, b"saveData").unwrap_or(false),
        save_password: attr_bool(e, b"savePassword").unwrap_or(false),
        new_connection: attr_bool(e, b"new").unwrap_or(false),
        deleted: attr_bool(e, b"deleted").unwrap_or(false),
        only_use_connection_file: attr_bool(e, b"onlyUseConnectionFile").unwrap_or(false),
        credentials: attr_string(e, b"credentials").and_then(|value| match value.as_str() {
            "integrated" => Some(WorkbookConnectionCredentials::Integrated),
            "none" => Some(WorkbookConnectionCredentials::None),
            "stored" => Some(WorkbookConnectionCredentials::Stored),
            "prompt" => Some(WorkbookConnectionCredentials::Prompt),
            _ => None,
        }),
        single_sign_on_id: attr_string(e, b"singleSignOnId"),
        parameters: Vec::new(),
        kind: None,
        db_props: None,
    })
}

#[derive(Clone)]
struct ConnectionDbProps {
    connection: String,
    command: Option<String>,
    command_type: Option<u32>,
}

impl ConnectionDbProps {
    fn database_kind(&self) -> WorkbookConnectionKind {
        WorkbookConnectionKind::Database {
            connection: self.connection.clone(),
            command: self.command.clone(),
            command_type: self.command_type,
        }
    }

    fn olap_kind(
        &self,
        local: bool,
        local_connection: Option<String>,
        local_refresh: bool,
        send_locale: bool,
        row_drill_count: Option<u32>,
    ) -> WorkbookConnectionKind {
        WorkbookConnectionKind::Olap {
            connection: Some(self.connection.clone()),
            command: self.command.clone(),
            command_type: self.command_type,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        }
    }
}

fn parse_db_pr(e: &quick_xml::events::BytesStart<'_>) -> Option<ConnectionDbProps> {
    Some(ConnectionDbProps {
        connection: attr_string(e, b"connection")?,
        command: attr_string(e, b"command"),
        command_type: attr_u32(e, b"commandType"),
    })
}

fn parse_olap_pr(
    e: &quick_xml::events::BytesStart<'_>,
    db_props: Option<&ConnectionDbProps>,
) -> WorkbookConnectionKind {
    let local = attr_bool(e, b"local").unwrap_or(false);
    let local_connection = attr_string(e, b"localConnection");
    let local_refresh = attr_bool(e, b"localRefresh").unwrap_or(true);
    let send_locale = attr_bool(e, b"sendLocale").unwrap_or(false);
    let row_drill_count = attr_u32(e, b"rowDrillCount");
    if let Some(db_props) = db_props {
        db_props.olap_kind(
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        )
    } else {
        WorkbookConnectionKind::Olap {
            connection: None,
            command: None,
            command_type: None,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        }
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

fn parse_connection_parameter(
    e: &quick_xml::events::BytesStart<'_>,
) -> WorkbookConnectionParameter {
    WorkbookConnectionParameter {
        name: attr_string(e, b"name"),
        sql_type: attr_i32(e, b"sqlType").unwrap_or(0),
        parameter_type: attr_string(e, b"parameterType")
            .and_then(|value| match value.as_str() {
                "value" => Some(WorkbookConnectionParameterType::Value),
                "cell" => Some(WorkbookConnectionParameterType::Cell),
                "prompt" => Some(WorkbookConnectionParameterType::Prompt),
                _ => None,
            })
            .unwrap_or(WorkbookConnectionParameterType::Prompt),
        refresh_on_change: attr_bool(e, b"refreshOnChange").unwrap_or(false),
        prompt: attr_string(e, b"prompt"),
        value: parse_connection_parameter_value(e),
    }
}

fn parse_connection_parameter_value(
    e: &quick_xml::events::BytesStart<'_>,
) -> WorkbookConnectionParameterValue {
    if let Some(value) = attr_bool(e, b"boolean") {
        WorkbookConnectionParameterValue::Boolean(value)
    } else if let Some(value) = attr_f64(e, b"double") {
        WorkbookConnectionParameterValue::Double(value)
    } else if let Some(value) = attr_i32(e, b"integer") {
        WorkbookConnectionParameterValue::Integer(value)
    } else if let Some(value) = attr_string(e, b"string") {
        WorkbookConnectionParameterValue::String(value)
    } else if let Some(value) = attr_string(e, b"cell") {
        WorkbookConnectionParameterValue::Cell(value)
    } else {
        WorkbookConnectionParameterValue::None
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

fn attr_i32(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<i32> {
    attr_string(e, key).and_then(|value| value.parse().ok())
}

fn attr_f64(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<f64> {
    attr_string(e, key).and_then(|value| value.parse().ok())
}

fn attr_bool(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<bool> {
    attr_string(e, key).map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
}

/// Read relationships owned by any package part.
pub(super) fn read_part_rels<R: Read + Seek>(
    package: &mut OpcPackage<R>,
    part_path: &str,
) -> XlsxResult<Arc<RelationshipSet>> {
    let source = RelationshipSource::Part(PartName::from_zip_name(part_path)?);
    package.relationships(&source, false)
}
