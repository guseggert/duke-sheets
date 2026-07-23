use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek};

use super::content_types::ContentTypes;
use super::diagnostics::{DiagnosticSink, XlsxDiagnostic, XlsxDiagnosticCode, XlsxPackagePolicy};
use super::part_name::PartName;
use super::relationships::{RelationshipSet, RelationshipSource};
use crate::error::{XlsxError, XlsxResult};

pub(crate) const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub(crate) const CT_TEMPLATE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
pub(crate) const CT_MACRO_WORKBOOK: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
pub(crate) const CT_MACRO_TEMPLATE: &str =
    "application/vnd.ms-excel.template.macroEnabled.main+xml";
pub(crate) const WORKBOOK_CONTENT_TYPES: &[&str] = &[
    CT_WORKBOOK,
    CT_TEMPLATE,
    CT_MACRO_WORKBOOK,
    CT_MACRO_TEMPLATE,
];

const RT_OFFICE_DOCUMENT_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const RT_OFFICE_DOCUMENT_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

pub(crate) struct OpcPackage<R: Read + Seek> {
    archive: zip::ZipArchive<R>,
    entries: HashMap<PartName, usize>,
    content_types: ContentTypes,
    relationships: HashMap<RelationshipSource, RelationshipSet>,
    validated_content_types: HashSet<PartName>,
    diagnostics: DiagnosticSink,
}

impl<R: Read + Seek> OpcPackage<R> {
    pub(crate) fn open(reader: R, policy: XlsxPackagePolicy) -> XlsxResult<Self> {
        let archive = zip::ZipArchive::new(reader)?;
        Self::from_archive(archive, policy)
    }

    pub(crate) fn from_archive(
        mut archive: zip::ZipArchive<R>,
        policy: XlsxPackagePolicy,
    ) -> XlsxResult<Self> {
        let mut diagnostics = DiagnosticSink::new(policy);
        let mut entries = HashMap::new();
        let mut ordered_names = Vec::new();
        let mut content_types_index = None;

        for (index, raw_name) in archive.file_names().map(str::to_string).enumerate() {
            if raw_name.ends_with('/') {
                continue;
            }
            if raw_name == "[Content_Types].xml" {
                if content_types_index.is_some() {
                    diagnostics.violation(
                        XlsxDiagnosticCode::EquivalentPartName,
                        "package contains more than one [Content_Types].xml stream",
                        None,
                        None,
                        Some("/[Content_Types].xml"),
                    )?;
                } else {
                    content_types_index = Some(index);
                }
                continue;
            }
            let part_name = match PartName::from_zip_name_with_policy(
                &raw_name,
                policy == XlsxPackagePolicy::Compatible,
            ) {
                Ok(part_name) => part_name,
                Err(error) => {
                    diagnostics.violation(
                        XlsxDiagnosticCode::InvalidPartName,
                        error.to_string(),
                        None,
                        None,
                        Some(&raw_name),
                    )?;
                    continue;
                }
            };
            if entries.contains_key(&part_name) {
                diagnostics.violation(
                    XlsxDiagnosticCode::EquivalentPartName,
                    format!("equivalent OPC part names include {part_name}"),
                    None,
                    None,
                    Some(part_name.as_str()),
                )?;
                continue;
            }
            for existing in &ordered_names {
                if part_name.is_derivable_from(existing) || existing.is_derivable_from(&part_name) {
                    diagnostics.violation(
                        XlsxDiagnosticCode::DerivablePartName,
                        format!("derivable OPC part names: {existing} and {part_name}"),
                        None,
                        None,
                        Some(part_name.as_str()),
                    )?;
                }
            }
            ordered_names.push(part_name.clone());
            entries.insert(part_name, index);
        }

        let content_types_index = content_types_index
            .ok_or_else(|| XlsxError::MissingPart("[Content_Types].xml".to_string()))?;
        let mut content_types_xml = Vec::new();
        archive
            .by_index(content_types_index)?
            .read_to_end(&mut content_types_xml)?;
        let content_types = ContentTypes::parse(Cursor::new(content_types_xml), &mut diagnostics)?;

        Ok(Self {
            archive,
            entries,
            content_types,
            relationships: HashMap::new(),
            validated_content_types: HashSet::new(),
            diagnostics,
        })
    }

    pub(crate) fn policy(&self) -> XlsxPackagePolicy {
        self.diagnostics.policy()
    }

    /// Locate the single SpreadsheetML Workbook part through package relationships.
    ///
    /// ECMA-376 Part 1 §§12.2 and 12.3.23 require the Workbook part to be
    /// the internal target of the package's officeDocument relationship.
    pub(crate) fn discover_workbook_part(&mut self) -> XlsxResult<PartName> {
        let source = RelationshipSource::Package;
        let relationships_part = source.relationships_part()?;
        let has_package_relationships = self.part_exists(&relationships_part);

        if has_package_relationships {
            let relationships = self.relationships(&source, true)?;
            let office_relationships: Vec<_> = relationships
                .by_type(&[RT_OFFICE_DOCUMENT_TRANSITIONAL, RT_OFFICE_DOCUMENT_STRICT])
                .cloned()
                .collect();
            if office_relationships.is_empty() {
                self.diagnostics.violation(
                    XlsxDiagnosticCode::MissingOfficeDocumentRelationship,
                    "package relationships contain no officeDocument relationship",
                    None,
                    None,
                    None,
                )?;
                return self.discover_workbook_fallback();
            }
            if office_relationships.len() > 1 {
                self.diagnostics.violation(
                    XlsxDiagnosticCode::AmbiguousOfficeDocumentRelationship,
                    "package contains more than one officeDocument relationship",
                    None,
                    None,
                    None,
                )?;
            }

            let mut targets = Vec::new();
            for relationship in office_relationships {
                let Some(part_name) = relationship.internal_part().cloned() else {
                    return Err(XlsxError::InvalidFormat(format!(
                        "officeDocument relationship {} is not a resolvable internal target",
                        relationship.id
                    )));
                };
                if !self.part_exists(&part_name) {
                    return Err(XlsxError::MissingPart(part_name.zip_name().to_string()));
                }
                if !targets.iter().any(|target| target == &part_name) {
                    targets.push(part_name);
                }
            }
            if targets.len() != 1 {
                return Err(XlsxError::InvalidFormat(
                    "package has multiple distinct officeDocument targets".into(),
                ));
            }
            let workbook = targets.remove(0);
            self.validate_workbook_content_type(&workbook)?;
            return Ok(workbook);
        }

        self.diagnostics.violation(
            XlsxDiagnosticCode::MissingPackageRelationships,
            "package is missing /_rels/.rels",
            None,
            None,
            None,
        )?;

        self.discover_workbook_fallback()
    }

    fn discover_workbook_fallback(&mut self) -> XlsxResult<PartName> {
        let candidates: Vec<_> = self
            .entries
            .keys()
            .filter(|part_name| {
                self.content_types
                    .content_type_for(part_name)
                    .is_some_and(|content_type| {
                        WORKBOOK_CONTENT_TYPES
                            .iter()
                            .any(|accepted| content_type.eq_ignore_ascii_case(accepted))
                    })
            })
            .cloned()
            .collect();
        if candidates.len() == 1 {
            self.diagnostics.recovery(
                XlsxDiagnosticCode::CanonicalPartFallback,
                format!(
                    "located Workbook part {} through its content type",
                    candidates[0]
                ),
                None,
                None,
                Some(candidates[0].as_str()),
            );
            return Ok(candidates[0].clone());
        }
        if candidates.len() > 1 {
            return Err(XlsxError::InvalidFormat(
                "content types identify multiple possible Workbook parts".into(),
            ));
        }

        let canonical = PartName::new("/xl/workbook.xml")?;
        if self.part_exists(&canonical) {
            self.validate_workbook_content_type(&canonical)?;
            self.diagnostics.recovery(
                XlsxDiagnosticCode::CanonicalPartFallback,
                "using conventional /xl/workbook.xml because package relationships are absent",
                None,
                None,
                Some(canonical.as_str()),
            );
            return Ok(canonical);
        }
        Err(XlsxError::MissingPart("xl/workbook.xml".into()))
    }

    pub(crate) fn part_exists(&self, part_name: &PartName) -> bool {
        self.entries.contains_key(part_name)
    }

    pub(crate) fn open_part(&mut self, part_name: &PartName) -> XlsxResult<zip::read::ZipFile<'_>> {
        self.validate_content_type_presence(part_name)?;
        let index = self
            .entries
            .get(part_name)
            .copied()
            .ok_or_else(|| XlsxError::MissingPart(part_name.zip_name().to_string()))?;
        Ok(self.archive.by_index(index)?)
    }

    pub(crate) fn relationships(
        &mut self,
        source: &RelationshipSource,
        required: bool,
    ) -> XlsxResult<RelationshipSet> {
        if let Some(relationships) = self.relationships.get(source) {
            return Ok(relationships.clone());
        }
        let relationships_part = source.relationships_part()?;
        if !self.part_exists(&relationships_part) {
            if required {
                self.diagnostics.violation(
                    XlsxDiagnosticCode::MissingRelationshipsPart,
                    format!(
                        "missing relationships part {} for {}",
                        relationships_part,
                        source.display_name()
                    ),
                    source.part_name().map(PartName::as_str),
                    None,
                    None,
                )?;
            }
            let relationships = RelationshipSet::default();
            self.relationships
                .insert(source.clone(), relationships.clone());
            return Ok(relationships);
        }

        let mut bytes = Vec::new();
        self.open_part(&relationships_part)?
            .read_to_end(&mut bytes)?;
        let relationships =
            RelationshipSet::parse(Cursor::new(bytes), source, &mut self.diagnostics)?;
        self.validate_relationship_set(source, &relationships)?;
        self.relationships
            .insert(source.clone(), relationships.clone());
        Ok(relationships)
    }

    pub(crate) fn open_related_part(
        &mut self,
        source: &RelationshipSource,
        relationship: &super::relationships::Relationship,
    ) -> XlsxResult<Option<zip::read::ZipFile<'_>>> {
        let Some(part_name) = relationship.internal_part().cloned() else {
            return Ok(None);
        };
        if !self.part_exists(&part_name) {
            self.diagnostics.violation(
                XlsxDiagnosticCode::MissingRelationshipTarget,
                format!(
                    "missing relationship target {part_name} from {}",
                    source.display_name()
                ),
                source.part_name().map(PartName::as_str),
                Some(&relationship.id),
                Some(&relationship.raw_target),
            )?;
            return Ok(None);
        }
        self.validate_relationship_content_type(&part_name, &relationship.rel_type)?;
        self.open_part(&part_name).map(Some)
    }

    pub(crate) fn diagnostics_mut(&mut self) -> &mut DiagnosticSink {
        &mut self.diagnostics
    }

    pub(crate) fn validate_part_content_type(
        &mut self,
        part_name: &PartName,
        expected: &str,
    ) -> XlsxResult<()> {
        self.validate_content_type_presence(part_name)?;
        let actual = self.content_types.content_type_for(part_name);
        if actual.is_some_and(|actual| actual.eq_ignore_ascii_case(expected)) {
            return Ok(());
        }
        self.diagnostics.violation(
            XlsxDiagnosticCode::PartContentTypeMismatch,
            format!(
                "part {part_name} has unexpected content type {}",
                actual.unwrap_or("<missing>")
            ),
            Some(part_name.as_str()),
            None,
            None,
        )
    }

    pub(crate) fn archive_mut(&mut self) -> &mut zip::ZipArchive<R> {
        &mut self.archive
    }

    pub(crate) fn into_diagnostics(self) -> Vec<XlsxDiagnostic> {
        self.diagnostics.into_diagnostics()
    }

    fn validate_workbook_content_type(&mut self, workbook: &PartName) -> XlsxResult<()> {
        let workbook_parts: Vec<_> = self
            .entries
            .keys()
            .filter(|part_name| {
                self.content_types
                    .content_type_for(part_name)
                    .is_some_and(|content_type| {
                        WORKBOOK_CONTENT_TYPES
                            .iter()
                            .any(|accepted| content_type.eq_ignore_ascii_case(accepted))
                    })
            })
            .cloned()
            .collect();
        if workbook_parts.len() > 1 {
            self.diagnostics.violation(
                XlsxDiagnosticCode::AmbiguousOfficeDocumentRelationship,
                "content types identify more than one Workbook part",
                Some(workbook.as_str()),
                None,
                None,
            )?;
        }
        let content_type = self.content_types.content_type_for(workbook);
        if content_type.is_some_and(|content_type| {
            WORKBOOK_CONTENT_TYPES
                .iter()
                .any(|accepted| content_type.eq_ignore_ascii_case(accepted))
        }) {
            return Ok(());
        }
        self.diagnostics.violation(
            XlsxDiagnosticCode::WorkbookContentTypeMismatch,
            format!(
                "Workbook part {workbook} has unsupported content type {}",
                content_type.unwrap_or("<missing>")
            ),
            Some(workbook.as_str()),
            None,
            None,
        )
    }

    fn validate_content_type_presence(&mut self, part_name: &PartName) -> XlsxResult<()> {
        if is_relationship_part(part_name) || self.validated_content_types.contains(part_name) {
            return Ok(());
        }
        if self.content_types.content_type_for(part_name).is_none() {
            self.diagnostics.violation(
                XlsxDiagnosticCode::MissingContentType,
                format!("part {part_name} has no effective content type"),
                Some(part_name.as_str()),
                None,
                None,
            )?;
        }
        self.validated_content_types.insert(part_name.clone());
        Ok(())
    }

    fn validate_relationship_content_type(
        &mut self,
        part_name: &PartName,
        rel_type: &str,
    ) -> XlsxResult<()> {
        let Some(expected) = expected_content_type(rel_type) else {
            return Ok(());
        };
        let actual = self.content_types.content_type_for(part_name);
        let matches = match expected {
            ContentTypeExpectation::Exact(expected) => {
                actual.is_some_and(|actual| actual.eq_ignore_ascii_case(expected))
            }
            ContentTypeExpectation::Prefix(prefix) => {
                actual.is_some_and(|actual| actual.to_ascii_lowercase().starts_with(prefix))
            }
        };
        if matches {
            return Ok(());
        }
        self.diagnostics.violation(
            XlsxDiagnosticCode::PartContentTypeMismatch,
            format!(
                "relationship target {part_name} has unexpected content type {}",
                actual.unwrap_or("<missing>")
            ),
            Some(part_name.as_str()),
            None,
            None,
        )
    }

    fn validate_relationship_set(
        &mut self,
        source: &RelationshipSource,
        relationships: &RelationshipSet,
    ) -> XlsxResult<()> {
        for relationship in relationships.iter() {
            let Some(part_name) = relationship.internal_part() else {
                continue;
            };
            if is_relationship_part(part_name) {
                self.diagnostics.violation(
                    XlsxDiagnosticCode::MalformedRelationship,
                    format!(
                        "relationship {} from {} targets a Relationships part",
                        relationship.id,
                        source.display_name()
                    ),
                    source.part_name().map(PartName::as_str),
                    Some(&relationship.id),
                    Some(&relationship.raw_target),
                )?;
                continue;
            }
            if !self.part_exists(part_name) {
                self.diagnostics.violation(
                    XlsxDiagnosticCode::MissingRelationshipTarget,
                    format!(
                        "missing relationship target {part_name} from {}",
                        source.display_name()
                    ),
                    source.part_name().map(PartName::as_str),
                    Some(&relationship.id),
                    Some(&relationship.raw_target),
                )?;
                continue;
            }
            self.validate_content_type_presence(part_name)?;
            self.validate_relationship_content_type(part_name, &relationship.rel_type)?;
        }
        Ok(())
    }
}

enum ContentTypeExpectation {
    Exact(&'static str),
    Prefix(&'static str),
}

fn expected_content_type(rel_type: &str) -> Option<ContentTypeExpectation> {
    let exact = |content_type| Some(ContentTypeExpectation::Exact(content_type));
    if office_relationship_type_is(rel_type, "worksheet") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    } else if office_relationship_type_is(rel_type, "chartsheet") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")
    } else if office_relationship_type_is(rel_type, "styles") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    } else if office_relationship_type_is(rel_type, "sharedStrings") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
    } else if office_relationship_type_is(rel_type, "theme") {
        exact("application/vnd.openxmlformats-officedocument.theme+xml")
    } else if office_relationship_type_is(rel_type, "drawing") {
        exact("application/vnd.openxmlformats-officedocument.drawing+xml")
    } else if office_relationship_type_is(rel_type, "chart") {
        exact("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
    } else if rel_type == "http://schemas.microsoft.com/office/2014/relationships/chartEx" {
        exact("application/vnd.ms-office.chartex+xml")
    } else if rel_type == "http://schemas.microsoft.com/office/2011/relationships/chartStyle" {
        exact("application/vnd.ms-office.chartstyle+xml")
    } else if rel_type == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" {
        exact("application/vnd.ms-office.chartcolorstyle+xml")
    } else if office_relationship_type_is(rel_type, "comments") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
    } else if office_relationship_type_is(rel_type, "vmlDrawing") {
        exact("application/vnd.openxmlformats-officedocument.vmlDrawing")
    } else if office_relationship_type_is(rel_type, "table") {
        exact("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    } else if office_relationship_type_is(rel_type, "ctrlProp") {
        exact("application/vnd.ms-excel.controlproperties+xml")
    } else if office_relationship_type_is(rel_type, "image") {
        Some(ContentTypeExpectation::Prefix("image/"))
    } else {
        None
    }
}

fn office_relationship_type_is(rel_type: &str, name: &str) -> bool {
    rel_type
        == format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{name}")
        || rel_type == format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{name}")
}

fn is_relationship_part(part_name: &PartName) -> bool {
    let segments: Vec<_> = part_name.as_str().split('/').collect();
    segments.len() >= 3
        && segments[segments.len() - 2].eq_ignore_ascii_case("_rels")
        && segments[segments.len() - 1]
            .to_ascii_lowercase()
            .ends_with(".rels")
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn package(
        entries: &[(&str, &[u8])],
        policy: XlsxPackagePolicy,
    ) -> XlsxResult<OpcPackage<Cursor<Vec<u8>>>> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for (name, bytes) in entries {
            zip.start_file(*name, zip::write::SimpleFileOptions::default())?;
            zip.write_all(bytes)?;
        }
        OpcPackage::open(zip.finish()?, policy)
    }

    #[test]
    fn package_indexes_parts_case_insensitively() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let mut package = package(
            &[
                ("[Content_Types].xml", content_types),
                ("XL/Workbook.xml", b"workbook"),
            ],
            XlsxPackagePolicy::Compatible,
        )
        .unwrap();
        let part_name = PartName::new("/xl/workbook.xml").unwrap();
        assert!(package.part_exists(&part_name));
        let mut contents = String::new();
        package
            .open_part(&part_name)
            .unwrap()
            .read_to_string(&mut contents)
            .unwrap();
        assert_eq!(contents, "workbook");
    }

    #[test]
    fn relationships_are_cached_by_source() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:type" Target="book.xml"/></Relationships>"#;
        let mut package = package(
            &[
                ("[Content_Types].xml", content_types),
                ("_rels/.rels", rels),
                ("book.xml", b"book"),
            ],
            XlsxPackagePolicy::Compatible,
        )
        .unwrap();
        let first = package
            .relationships(&RelationshipSource::Package, true)
            .unwrap();
        let second = package
            .relationships(&RelationshipSource::Package, true)
            .unwrap();
        assert_eq!(first, second);
        assert_eq!(first.len(), 1);
    }
}
