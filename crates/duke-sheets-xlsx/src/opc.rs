use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};
use crate::reader::archive_by_name;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum RelationshipTarget {
    Internal(String),
    External(String),
    UnresolvedInternal,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Relationship {
    pub(crate) rel_type: String,
    pub(crate) raw_target: String,
    pub(crate) target: RelationshipTarget,
}

impl Relationship {
    pub(crate) fn internal_path(&self) -> Option<&str> {
        match &self.target {
            RelationshipTarget::Internal(path) => Some(path),
            RelationshipTarget::External(_) | RelationshipTarget::UnresolvedInternal => None,
        }
    }

    pub(crate) fn target(&self) -> &str {
        match &self.target {
            RelationshipTarget::Internal(path) | RelationshipTarget::External(path) => path,
            RelationshipTarget::UnresolvedInternal => &self.raw_target,
        }
    }

    pub(crate) fn is_external(&self) -> bool {
        matches!(self.target, RelationshipTarget::External(_))
    }
}

pub(crate) fn relationships_part_path(source_part: &str) -> String {
    let source_part = source_part.trim_start_matches('/');
    if source_part.is_empty() {
        return "_rels/.rels".to_string();
    }
    match source_part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{source_part}.rels"),
    }
}

fn decode_percent_encoded_unreserved(value: &str) -> String {
    fn hex_value(byte: u8) -> Option<u8> {
        match byte {
            b'0'..=b'9' => Some(byte - b'0'),
            b'a'..=b'f' => Some(byte - b'a' + 10),
            b'A'..=b'F' => Some(byte - b'A' + 10),
            _ => None,
        }
    }

    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let escaped = if bytes[index] == b'%' && index + 2 < bytes.len() {
            hex_value(bytes[index + 1])
                .zip(hex_value(bytes[index + 2]))
                .map(|(high, low)| high * 16 + low)
        } else {
            None
        };
        if let Some(byte) = escaped {
            if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~') {
                decoded.push(byte);
                index += 3;
                continue;
            }
        }
        decoded.push(bytes[index]);
        index += 1;
    }
    match String::from_utf8(decoded) {
        Ok(decoded) => decoded,
        Err(_) => value.to_string(),
    }
}

/// Resolve an internal OPC relationship target to its ZIP entry name.
pub(crate) fn resolve_internal_target(source_part: &str, target: &str) -> XlsxResult<String> {
    let normalized_target = target.replace('\\', "/");
    let target_path = decode_percent_encoded_unreserved(
        normalized_target
            .split(['?', '#'])
            .next()
            .unwrap_or_default(),
    );
    if target_path.is_empty() {
        let source_part = source_part.trim_start_matches('/');
        if !source_part.is_empty() {
            return Ok(source_part.to_string());
        }
    }
    let absolute = target_path.starts_with('/');
    let mut parts: Vec<&str> = if absolute {
        Vec::new()
    } else {
        let source_part = source_part.trim_start_matches('/');
        source_part
            .rsplit_once('/')
            .map(|(dir, _)| dir.split('/').filter(|part| !part.is_empty()).collect())
            .unwrap_or_default()
    };

    for part in target_path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(XlsxError::InvalidFormat(format!(
                        "relationship target escapes package root: source={source_part}, target={target}"
                    )));
                }
            }
            _ => parts.push(part),
        }
    }

    if parts.is_empty() {
        return Err(XlsxError::InvalidFormat(format!(
            "relationship target does not name a package part: source={source_part}, target={target}"
        )));
    }
    Ok(parts.join("/"))
}

pub(crate) fn read_relationships<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    source_part: &str,
    required: bool,
) -> XlsxResult<HashMap<String, Relationship>> {
    let rels_path = relationships_part_path(source_part);
    let file = match archive_by_name(archive, &rels_path) {
        Ok(file) => file,
        Err(_) if !required => return Ok(HashMap::new()),
        Err(_) => return Err(XlsxError::MissingPart(rels_path)),
    };
    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut relationships = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if element.name().local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut rel_type = None;
                let mut target_mode = None;
                for attr in element.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"Id" => id = attr.unescape_value().ok().map(|value| value.to_string()),
                        b"Target" => {
                            target = attr.unescape_value().ok().map(|value| value.to_string())
                        }
                        b"Type" => {
                            rel_type = attr.unescape_value().ok().map(|value| value.to_string())
                        }
                        b"TargetMode" => {
                            target_mode = attr.unescape_value().ok().map(|value| value.to_string())
                        }
                        _ => {}
                    }
                }

                if let (Some(id), Some(raw_target), Some(rel_type)) = (id, target, rel_type) {
                    let target = if target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
                    {
                        RelationshipTarget::External(raw_target.clone())
                    } else {
                        let path = match resolve_internal_target(source_part, &raw_target) {
                            Ok(path) => path,
                            Err(error) => {
                                log::warn!(
                                    "cannot resolve internal relationship: source={source_part}, id={id}, target={raw_target}: {error}"
                                );
                                relationships.insert(
                                    id,
                                    Relationship {
                                        rel_type,
                                        raw_target,
                                        target: RelationshipTarget::UnresolvedInternal,
                                    },
                                );
                                continue;
                            }
                        };
                        RelationshipTarget::Internal(path)
                    };
                    relationships.insert(
                        id,
                        Relationship {
                            rel_type,
                            raw_target,
                            target,
                        },
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(XlsxError::Xml(error)),
            _ => {}
        }
        buf.clear();
    }

    Ok(relationships)
}

pub(crate) fn open_relationship_part<'a, R: Read + Seek>(
    archive: &'a mut zip::ZipArchive<R>,
    source_part: &str,
    relationship_id: &str,
    relationship: &Relationship,
) -> Option<zip::read::ZipFile<'a>> {
    let path = relationship.internal_path()?;
    match archive_by_name(archive, path) {
        Ok(file) => Some(file),
        Err(error) => {
            log::warn!(
                "missing internal relationship target: source={source_part}, id={relationship_id}, target={}, resolved={path}: {error}",
                relationship.raw_target
            );
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    #[test]
    fn relationship_part_paths_follow_the_owner_part() {
        assert_eq!(relationships_part_path(""), "_rels/.rels");
        assert_eq!(
            relationships_part_path("xl/workbook.xml"),
            "xl/_rels/workbook.xml.rels"
        );
        assert_eq!(relationships_part_path("root.xml"), "_rels/root.xml.rels");
    }

    #[test]
    fn internal_targets_resolve_against_the_owner_part() {
        assert_eq!(
            resolve_internal_target("xl/worksheets/sheet1.xml", "../drawings/./drawing1.xml")
                .unwrap(),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_internal_target("xl/worksheets/sheet1.xml", "/xl/drawings/drawing1.xml")
                .unwrap(),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_internal_target("workbook.xml", "worksheets/sheet1.xml").unwrap(),
            "worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_internal_target("xl/workbook.xml", "#fragment").unwrap(),
            "xl/workbook.xml"
        );
        assert_eq!(
            resolve_internal_target("xl/drawings/drawing1.xml", "../charts/chart%31.xml").unwrap(),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn internal_targets_cannot_escape_the_package_root() {
        assert!(resolve_internal_target("xl/workbook.xml", "../../outside.xml").is_err());
        assert!(resolve_internal_target("xl/workbook.xml", "/../outside.xml").is_err());
        assert!(
            resolve_internal_target("xl/workbook.xml", "%2e%2e/%2e%2e/outside.xml").is_err()
        );
    }

    #[test]
    fn relationship_parser_preserves_raw_and_external_targets() {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            zip.start_file(
                "xl/worksheets/_rels/sheet1.xml.rels",
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="drawing" Target="/xl/drawings/drawing1.xml"/><Relationship Id="rId2" Type="hyperlink" Target="https://example.com/a/../b" TargetMode="External"/><Relationship Id="rId3" Type="raw" Target="../../../outside.xml"/></Relationships>"#).unwrap();
            zip.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let rels = read_relationships(&mut archive, "xl/worksheets/sheet1.xml", true).unwrap();

        assert_eq!(rels["rId1"].raw_target, "/xl/drawings/drawing1.xml");
        assert_eq!(
            rels["rId1"].internal_path(),
            Some("xl/drawings/drawing1.xml")
        );
        assert_eq!(rels["rId2"].target(), "https://example.com/a/../b");
        assert!(rels["rId2"].is_external());
        assert_eq!(rels["rId3"].raw_target, "../../../outside.xml");
        assert_eq!(rels["rId3"].internal_path(), None);
        assert!(!rels["rId3"].is_external());
    }
}
