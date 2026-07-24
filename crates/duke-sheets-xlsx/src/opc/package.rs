use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

use super::content_types::ContentTypes;
use super::diagnostics::{DiagnosticSink, XlsxDiagnostic, XlsxDiagnosticCode, XlsxPackagePolicy};
use super::part_name::PartName;
use super::relationship_kind::{
    ContentTypeExpectation, RelationshipKind, CT_MACRO_TEMPLATE, CT_MACRO_WORKBOOK, CT_TEMPLATE,
    CT_WORKBOOK,
};
use super::relationships::{RelationshipSet, RelationshipSource};
use crate::error::{XlsxError, XlsxResult};

pub(crate) const WORKBOOK_CONTENT_TYPES: &[&str] = &[
    CT_WORKBOOK,
    CT_TEMPLATE,
    CT_MACRO_WORKBOOK,
    CT_MACRO_TEMPLATE,
];

pub(crate) struct OpcPackage<R: Read + Seek> {
    archive: zip::ZipArchive<R>,
    entries: HashMap<PartName, PackageEntry>,
    content_types: ContentTypes,
    /// False when `[Content_Types].xml` was unreadable and Compatible
    /// mode degraded to an empty map; content-type checks are skipped.
    content_types_usable: bool,
    relationships: HashMap<RelationshipSource, Arc<RelationshipSet>>,
    validated_content_types: HashSet<PartName>,
    diagnostics: DiagnosticSink,
}

struct PackageEntry {
    variants: Vec<EntryVariant>,
}

struct EntryVariant {
    index: usize,
    name: String,
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
        let mut entries: HashMap<PartName, PackageEntry> = HashMap::new();
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
            if let Some(existing) = entries.get_mut(&part_name) {
                diagnostics.violation(
                    XlsxDiagnosticCode::EquivalentPartName,
                    format!("equivalent OPC part names include {part_name}"),
                    None,
                    None,
                    Some(part_name.as_str()),
                )?;
                existing.variants.push(EntryVariant {
                    index,
                    name: raw_name,
                });
                continue;
            }
            ordered_names.push(part_name.clone());
            entries.insert(
                part_name,
                PackageEntry {
                    variants: vec![EntryVariant {
                        index,
                        name: raw_name,
                    }],
                },
            );
        }

        let lowercased: HashSet<String> = ordered_names
            .iter()
            .map(|name| name.as_str().to_ascii_lowercase())
            .collect();
        for part_name in &ordered_names {
            let lower = part_name.as_str().to_ascii_lowercase();
            for (offset, _) in lower.match_indices('/').skip(1) {
                if lowercased.contains(&lower[..offset]) {
                    diagnostics.violation(
                        XlsxDiagnosticCode::DerivablePartName,
                        format!(
                            "derivable OPC part names: {} and {part_name}",
                            &lower[..offset]
                        ),
                        None,
                        None,
                        Some(part_name.as_str()),
                    )?;
                }
            }
        }

        let content_types_index = content_types_index
            .ok_or_else(|| XlsxError::MissingPart("[Content_Types].xml".to_string()))?;
        let mut content_types_xml = Vec::new();
        archive
            .by_index(content_types_index)?
            .read_to_end(&mut content_types_xml)?;
        let (content_types, content_types_usable) =
            match ContentTypes::parse(Cursor::new(content_types_xml), &mut diagnostics) {
                Ok(content_types) => (content_types, true),
                Err(error) if policy == XlsxPackagePolicy::Strict => return Err(error),
                Err(error) => {
                    diagnostics.warning(
                        XlsxDiagnosticCode::MalformedContentType,
                        format!("ignoring unreadable [Content_Types].xml: {error}"),
                        Some("/[Content_Types].xml"),
                        None,
                        None,
                    );
                    (ContentTypes::default(), false)
                }
            };

        Ok(Self {
            archive,
            entries,
            content_types,
            content_types_usable,
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
                .iter()
                .filter(|relationship| {
                    relationship.kind() == Some(RelationshipKind::OfficeDocument)
                })
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

            let mut targets: Vec<PartName> = Vec::new();
            for relationship in office_relationships {
                let Some(part_name) = relationship.internal_part().cloned() else {
                    // External mode is diagnosed here; unresolved internal
                    // targets were already diagnosed while parsing.
                    if relationship.is_external() {
                        self.diagnostics.violation(
                            XlsxDiagnosticCode::MalformedRelationship,
                            format!(
                                "officeDocument relationship {} must target an internal part",
                                relationship.id
                            ),
                            None,
                            Some(&relationship.id),
                            Some(&relationship.raw_target),
                        )?;
                    }
                    continue;
                };
                // Missing targets were diagnosed while parsing the set.
                if self.part_exists(&part_name) && !targets.contains(&part_name) {
                    targets.push(part_name);
                }
            }
            if targets.is_empty() {
                return self.discover_workbook_fallback();
            }
            let canonical = PartName::new("/xl/workbook.xml")?;
            let preferred = targets
                .iter()
                .position(|target| target == &canonical && self.has_workbook_content_type(target))
                .or_else(|| {
                    targets
                        .iter()
                        .position(|target| self.has_workbook_content_type(target))
                })
                .or_else(|| targets.iter().position(|target| target == &canonical))
                .unwrap_or(0);
            let workbook = targets.remove(preferred);
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
        let mut candidates: Vec<_> = self
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
        candidates.sort_by(|left, right| {
            left.as_str()
                .to_ascii_lowercase()
                .cmp(&right.as_str().to_ascii_lowercase())
        });
        if candidates.len() > 1 {
            self.diagnostics.violation(
                XlsxDiagnosticCode::AmbiguousOfficeDocumentRelationship,
                "content types identify multiple possible Workbook parts",
                None,
                None,
                None,
            )?;
        }
        if !candidates.is_empty() {
            let canonical = PartName::new("/xl/workbook.xml")?;
            let chosen = if candidates.contains(&canonical) {
                canonical
            } else {
                candidates.remove(0)
            };
            self.diagnostics.recovery(
                XlsxDiagnosticCode::CanonicalPartFallback,
                format!("located Workbook part {chosen} through its content type"),
                None,
                None,
                Some(chosen.as_str()),
            );
            return Ok(chosen);
        }

        let canonical = PartName::new("/xl/workbook.xml")?;
        if self.part_exists(&canonical) {
            self.validate_workbook_content_type(&canonical)?;
            self.diagnostics.recovery(
                XlsxDiagnosticCode::CanonicalPartFallback,
                "using conventional /xl/workbook.xml because it was not discoverable through package relationships",
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
        let entry = self
            .entries
            .get(part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.zip_name().to_string()))?;
        let backslashed = part_name.zip_name().replace('/', "\\");
        let index = entry
            .variants
            .iter()
            .find(|variant| variant.name == part_name.zip_name())
            .or_else(|| {
                entry
                    .variants
                    .iter()
                    .find(|variant| variant.name == backslashed)
            })
            .or_else(|| entry.variants.first())
            .map(|variant| variant.index)
            .ok_or_else(|| XlsxError::MissingPart(part_name.zip_name().to_string()))?;
        Ok(self.archive.by_index(index)?)
    }

    pub(crate) fn open_zip_name(&mut self, path: &str) -> XlsxResult<zip::read::ZipFile<'_>> {
        self.open_part(&PartName::from_zip_name(path)?)
    }

    pub(crate) fn relationships(
        &mut self,
        source: &RelationshipSource,
        required: bool,
    ) -> XlsxResult<Arc<RelationshipSet>> {
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
            let relationships = Arc::new(RelationshipSet::default());
            self.relationships
                .insert(source.clone(), relationships.clone());
            return Ok(relationships);
        }

        let mut bytes = Vec::new();
        self.open_part(&relationships_part)?
            .read_to_end(&mut bytes)?;
        let relationships =
            match RelationshipSet::parse(Cursor::new(bytes), source, &mut self.diagnostics) {
                Ok(relationships) => relationships,
                Err(error) if self.policy() == XlsxPackagePolicy::Compatible => {
                    self.diagnostics.warning(
                        XlsxDiagnosticCode::MalformedRelationship,
                        format!(
                            "ignoring unreadable relationships for {}: {error}",
                            source.display_name()
                        ),
                        source.part_name().map(PartName::as_str),
                        None,
                        None,
                    );
                    RelationshipSet::default()
                }
                Err(error) => return Err(error),
            };
        self.validate_relationship_set(source, &relationships)?;
        let relationships = Arc::new(relationships);
        self.relationships
            .insert(source.clone(), relationships.clone());
        Ok(relationships)
    }

    /// Open a relationship's internal target. Missing or mistyped targets
    /// were already diagnosed when the relationship set was parsed.
    pub(crate) fn open_related_part(
        &mut self,
        relationship: &super::relationships::Relationship,
    ) -> XlsxResult<Option<zip::read::ZipFile<'_>>> {
        let Some(part_name) = relationship.internal_part().cloned() else {
            return Ok(None);
        };
        if !self.part_exists(&part_name) {
            return Ok(None);
        }
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
        if !self.content_types_usable {
            return Ok(());
        }
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

    pub(crate) fn into_diagnostics(self) -> Vec<XlsxDiagnostic> {
        self.diagnostics.into_diagnostics()
    }

    fn validate_workbook_content_type(&mut self, workbook: &PartName) -> XlsxResult<()> {
        if !self.content_types_usable {
            return Ok(());
        }
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

    fn has_workbook_content_type(&self, part_name: &PartName) -> bool {
        self.content_types_usable
            && self
                .content_types
                .content_type_for(part_name)
                .is_some_and(|content_type| {
                    WORKBOOK_CONTENT_TYPES
                        .iter()
                        .any(|accepted| content_type.eq_ignore_ascii_case(accepted))
                })
    }

    fn validate_content_type_presence(&mut self, part_name: &PartName) -> XlsxResult<()> {
        if !self.content_types_usable
            || is_relationship_part(part_name)
            || self.validated_content_types.contains(part_name)
        {
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
        if !self.content_types_usable {
            return Ok(());
        }
        let Some(expected) =
            RelationshipKind::from_uri(rel_type).and_then(RelationshipKind::content_type)
        else {
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
    fn package_accepts_backslash_entry_names_in_compatible_mode() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let mut package = package(
            &[
                ("[Content_Types].xml", content_types),
                ("xl\\workbook.xml", b"workbook"),
            ],
            XlsxPackagePolicy::Compatible,
        )
        .unwrap();
        let mut contents = String::new();
        package
            .open_zip_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut contents)
            .unwrap();
        assert_eq!(contents, "workbook");
    }

    #[test]
    fn package_prefers_exact_entry_name_among_equivalents() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let mut package = package(
            &[
                ("[Content_Types].xml", content_types),
                ("xl\\workbook.xml", b"backslash"),
                ("xl/workbook.xml", b"exact"),
            ],
            XlsxPackagePolicy::Compatible,
        )
        .unwrap();
        let mut contents = String::new();
        package
            .open_zip_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut contents)
            .unwrap();
        assert_eq!(contents, "exact");
    }

    #[test]
    fn package_prefers_exact_backslash_before_case_only_equivalent() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let mut package = package(
            &[
                ("[Content_Types].xml", content_types),
                ("xl\\workbook.xml", b"backslash"),
                ("XL/workbook.xml", b"case-only"),
            ],
            XlsxPackagePolicy::Compatible,
        )
        .unwrap();
        let mut contents = String::new();
        package
            .open_zip_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut contents)
            .unwrap();
        assert_eq!(contents, "backslash");
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
