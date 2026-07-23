mod content_types;
mod diagnostics;
mod package;
mod part_name;
mod relationships;

use std::io::{Read, Seek};

#[cfg(test)]
use std::collections::HashMap;
#[cfg(test)]
use std::io::BufReader;

#[cfg(test)]
use crate::error::XlsxError;
use crate::error::XlsxResult;
#[cfg(test)]
use crate::reader::archive_by_name;

#[cfg(test)]
pub(crate) use diagnostics::DiagnosticSink;
pub use diagnostics::{
    XlsxDiagnostic, XlsxDiagnosticCode, XlsxDiagnosticSeverity, XlsxPackagePolicy,
};
pub(crate) use package::OpcPackage;
pub(crate) use part_name::PartName;
#[cfg(test)]
pub(crate) use relationships::RelationshipSet;
pub(crate) use relationships::{Relationship, RelationshipSource};

#[cfg(test)]
pub(crate) fn relationships_part_path(source_part: &str) -> String {
    if source_part.is_empty() {
        return "_rels/.rels".to_string();
    }
    PartName::from_zip_name(source_part)
        .and_then(|part_name| part_name.relationships_part())
        .map(|part_name| part_name.zip_name().to_string())
        .unwrap_or_else(|_| {
            let (directory, file) = source_part.rsplit_once('/').unwrap_or(("", source_part));
            if directory.is_empty() {
                format!("_rels/{file}.rels")
            } else {
                format!("{directory}/_rels/{file}.rels")
            }
        })
}

/// Resolve an internal OPC relationship target to its ZIP entry name.
pub(crate) fn resolve_internal_target(source_part: &str, target: &str) -> XlsxResult<String> {
    let source = if source_part.is_empty() {
        None
    } else {
        Some(PartName::from_zip_name(source_part)?)
    };
    part_name::resolve_internal_target(source.as_ref(), target)
        .map(|part_name| part_name.zip_name().to_string())
}

#[cfg(test)]
pub(crate) fn read_relationships<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    source_part: &str,
    required: bool,
) -> XlsxResult<HashMap<String, Relationship>> {
    let source = if source_part.is_empty() {
        RelationshipSource::Package
    } else {
        RelationshipSource::Part(PartName::from_zip_name(source_part)?)
    };
    let rels_path = source.relationships_part()?;
    let file = match archive_by_name(archive, rels_path.zip_name()) {
        Ok(file) => file,
        Err(_) if !required => return Ok(HashMap::new()),
        Err(_) => return Err(XlsxError::MissingPart(rels_path.zip_name().to_string())),
    };
    let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
    let relationships = RelationshipSet::parse(BufReader::new(file), &source, &mut diagnostics)?;
    for diagnostic in diagnostics.into_diagnostics() {
        log::warn!("{}", diagnostic.message);
    }
    Ok(relationships
        .iter()
        .cloned()
        .map(|relationship| (relationship.id.clone(), relationship))
        .collect())
}

pub(crate) fn open_relationship_part<'a, R: Read + Seek>(
    package: &'a mut OpcPackage<R>,
    source_part: &str,
    _relationship_id: &str,
    relationship: &Relationship,
) -> XlsxResult<Option<zip::read::ZipFile<'a>>> {
    let source = RelationshipSource::Part(PartName::from_zip_name(source_part)?);
    package.open_related_part(&source, relationship)
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
        assert!(resolve_internal_target("xl/workbook.xml", "%2e%2e/%2e%2e/outside.xml").is_err());
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
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:drawing" Target="/xl/drawings/drawing1.xml"/><Relationship Id="rId2" Type="urn:hyperlink" Target="https://example.com/a/../b" TargetMode="External"/><Relationship Id="rId3" Type="urn:raw" Target="../../../outside.xml"/></Relationships>"#).unwrap();
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
