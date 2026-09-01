use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, Event};

use crate::error::{XlsxError, XlsxResult};
use crate::opc::PartName;

#[derive(Debug, Default)]
pub(super) struct OpcManifest {
    defaults: BTreeMap<String, String>,
    parts: HashMap<PartName, ManifestPart>,
    relationships: HashMap<ManifestSource, Vec<ManifestRelationship>>,
    relationship_parts: HashSet<PartName>,
}

/// Owner of a relationships part: the package itself or one part.
/// `PartName` compares case-insensitively, so this must keep the
/// original spelling to derive a correctly cased `.rels` name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum ManifestSource {
    Package,
    Part(PartName),
}

impl ManifestSource {
    fn relationships_part(&self) -> XlsxResult<PartName> {
        match self {
            Self::Package => PartName::new("/_rels/.rels"),
            Self::Part(part_name) => part_name.relationships_part(),
        }
    }

    fn display_name(&self) -> &str {
        match self {
            Self::Package => "/",
            Self::Part(part_name) => part_name.as_str(),
        }
    }
}

#[derive(Debug)]
struct ManifestPart {
    content_type: String,
    raw: bool,
}

#[derive(Debug)]
struct ManifestRelationship {
    id: String,
    rel_type: String,
    target: String,
    external: bool,
}

impl OpcManifest {
    pub(super) fn new() -> XlsxResult<Self> {
        let mut manifest = Self::default();
        manifest.register_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        )?;
        manifest.register_default("xml", "application/xml")?;
        Ok(manifest)
    }

    pub(super) fn register_default(
        &mut self,
        extension: &str,
        content_type: &str,
    ) -> XlsxResult<()> {
        let extension = extension.to_ascii_lowercase();
        if let Some(existing) = self.defaults.get(&extension) {
            if existing != content_type {
                return Err(XlsxError::InvalidFormat(format!(
                    "conflicting content types for .{extension}: {existing} and {content_type}"
                )));
            }
            return Ok(());
        }
        self.defaults.insert(extension, content_type.to_string());
        Ok(())
    }

    pub(super) fn register_part(&mut self, zip_path: &str, content_type: &str) -> XlsxResult<()> {
        let part_name = PartName::from_zip_name(zip_path)?;
        if let Some(existing) = self.parts.get(&part_name) {
            if existing.raw {
                return Err(XlsxError::InvalidFormat(format!(
                    "emitted part {part_name} collides with a raw preserved part"
                )));
            }
            if existing.content_type != content_type {
                return Err(XlsxError::InvalidFormat(format!(
                    "conflicting content types for {part_name}: {} and {content_type}",
                    existing.content_type
                )));
            }
            return Ok(());
        }
        self.parts.insert(
            part_name,
            ManifestPart {
                content_type: content_type.to_string(),
                raw: false,
            },
        );
        Ok(())
    }

    pub(super) fn register_raw_part(
        &mut self,
        zip_path: &str,
        content_type: &str,
    ) -> XlsxResult<()> {
        let part_name = PartName::from_zip_name(zip_path)?;
        if self.parts.contains_key(&part_name) || self.relationship_parts.contains(&part_name) {
            return Err(XlsxError::InvalidFormat(format!(
                "raw part {part_name} collides with an emitted package part"
            )));
        }
        self.parts.insert(
            part_name,
            ManifestPart {
                content_type: content_type.to_string(),
                raw: true,
            },
        );
        Ok(())
    }

    pub(super) fn generated_zip_paths(&self) -> impl Iterator<Item = &str> {
        self.parts
            .iter()
            .filter(|(_, part)| !part.raw)
            .map(|(name, _)| name.zip_name())
            .chain(self.relationship_parts.iter().map(PartName::zip_name))
    }

    pub(super) fn write_content_types<W: Write + Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
    ) -> XlsxResult<()> {
        super::write_xml_part(zip, "[Content_Types].xml", |writer| {
            let mut root = BytesStart::new("Types");
            root.push_attribute((
                "xmlns",
                "http://schemas.openxmlformats.org/package/2006/content-types",
            ));
            writer.write_event(Event::Start(root))?;

            for (extension, content_type) in &self.defaults {
                writer
                    .create_element("Default")
                    .with_attribute(("Extension", extension.as_str()))
                    .with_attribute(("ContentType", content_type.as_str()))
                    .write_empty()?;
            }
            let mut parts: Vec<_> = self.parts.iter().collect();
            parts.sort_by(|(left, _), (right, _)| {
                left.as_str()
                    .to_ascii_lowercase()
                    .cmp(&right.as_str().to_ascii_lowercase())
            });
            for (part_name, part) in parts {
                let content_type = &part.content_type;
                let extension = part_name
                    .zip_name()
                    .rsplit('/')
                    .next()
                    .unwrap_or_default()
                    .rsplit_once('.')
                    .map(|(_, extension)| extension.to_ascii_lowercase());
                let covered_by_default = extension
                    .as_ref()
                    .and_then(|extension| self.defaults.get(extension))
                    .is_some_and(|default| default == content_type);
                if covered_by_default {
                    continue;
                }
                writer
                    .create_element("Override")
                    .with_attribute(("PartName", part_name.as_str()))
                    .with_attribute(("ContentType", content_type.as_str()))
                    .write_empty()?;
            }
            writer.write_event(Event::End(BytesEnd::new("Types")))?;
            Ok(())
        })
    }

    pub(super) fn register_relationship(
        &mut self,
        source_zip_path: Option<&str>,
        id: &str,
        rel_type: &str,
        target: &str,
        external: bool,
    ) -> XlsxResult<()> {
        self.register(source_zip_path, id, rel_type, target, external, false)
    }

    /// Register a relationship round-tripped from the source package.
    /// A preserved internal target whose part we never captured is
    /// dropped with a warning: emitting a relationship to a part that
    /// is not in the package would make the output non-conforming
    /// (ECMA-376 Part 2 §9.3.1), and we cannot invent the bytes.
    pub(super) fn register_preserved_relationship(
        &mut self,
        source_zip_path: Option<&str>,
        id: &str,
        rel_type: &str,
        target: &str,
        external: bool,
    ) -> XlsxResult<()> {
        self.register(source_zip_path, id, rel_type, target, external, true)
    }

    fn register(
        &mut self,
        source_zip_path: Option<&str>,
        id: &str,
        rel_type: &str,
        target: &str,
        external: bool,
        preserved: bool,
    ) -> XlsxResult<()> {
        let source = match source_zip_path {
            Some(path) => {
                let source_part = PartName::from_zip_name(path)?;
                if !self.parts.contains_key(&source_part) {
                    return Err(XlsxError::InvalidFormat(format!(
                        "relationship source {source_part} is not an emitted part"
                    )));
                }
                ManifestSource::Part(source_part)
            }
            None => ManifestSource::Package,
        };
        let relationships_part = source.relationships_part()?;
        if self
            .parts
            .get(&relationships_part)
            .is_some_and(|part| part.raw)
        {
            return Err(XlsxError::InvalidFormat(format!(
                "relationships part {relationships_part} collides with a raw preserved part"
            )));
        }
        if !external {
            let source_path = source_zip_path.unwrap_or("");
            let resolved = match crate::opc::resolve_internal_target(source_path, target) {
                Ok(resolved) => Some(resolved),
                Err(error) if preserved => {
                    log::warn!(
                        "dropping preserved relationship {id} from {}: {error}",
                        source.display_name()
                    );
                    return Ok(());
                }
                Err(error) => return Err(error),
            };
            let target_part = resolved
                .as_deref()
                .map(PartName::from_zip_name)
                .transpose()?;
            if target_part
                .as_ref()
                .is_none_or(|part| !self.parts.contains_key(part))
            {
                if preserved {
                    log::warn!(
                        "dropping preserved relationship {id} from {}: target part {target} is not in the package",
                        source.display_name()
                    );
                    return Ok(());
                }
                return Err(XlsxError::InvalidFormat(format!(
                    "relationship {id} from {} targets unregistered part {}",
                    source.display_name(),
                    target_part.map_or_else(|| target.to_string(), |part| part.to_string())
                )));
            }
        }
        self.relationship_parts.insert(relationships_part);
        let display_name = source.display_name().to_string();
        let relationships = self.relationships.entry(source).or_default();
        if relationships
            .iter()
            .any(|relationship| relationship.id == id)
        {
            return Err(XlsxError::InvalidFormat(format!(
                "duplicate relationship id {id} for {display_name}"
            )));
        }
        relationships.push(ManifestRelationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            external,
        });
        Ok(())
    }

    pub(super) fn write_relationships<W: Write + Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
    ) -> XlsxResult<()> {
        let mut sources: Vec<_> = self.relationships.keys().collect();
        sources.sort_by_key(|source| source.display_name().to_ascii_lowercase());
        for source in sources {
            let relationships = &self.relationships[source];
            let path = source.relationships_part()?.zip_name().to_string();
            super::write_xml_part(zip, &path, |writer| {
                let mut root = BytesStart::new("Relationships");
                root.push_attribute((
                    "xmlns",
                    "http://schemas.openxmlformats.org/package/2006/relationships",
                ));
                writer.write_event(Event::Start(root))?;
                for relationship in relationships {
                    let mut element = BytesStart::new("Relationship");
                    element.push_attribute(("Id", relationship.id.as_str()));
                    element.push_attribute(("Type", relationship.rel_type.as_str()));
                    element.push_attribute(("Target", relationship.target.as_str()));
                    if relationship.external {
                        element.push_attribute(("TargetMode", "External"));
                    }
                    writer.write_event(Event::Empty(element))?;
                }
                writer.write_event(Event::End(BytesEnd::new("Relationships")))?;
                Ok(())
            })?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::opc::RelationshipKind;
    use std::io::{Cursor, Read};

    #[test]
    fn manifest_uses_defaults_and_overrides() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest.register_default("png", "image/png").unwrap();
        manifest
            .register_part("xl/media/image1.png", "image/png")
            .unwrap();
        manifest
            .register_part("xl/workbook.xml", "workbook/type")
            .unwrap();
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        manifest.write_content_types(&mut zip).unwrap();
        let mut archive = zip::ZipArchive::new(zip.finish().unwrap()).unwrap();
        let mut xml = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("Extension=\"png\""));
        assert!(xml.contains("PartName=\"/xl/workbook.xml\""));
        assert!(!xml.contains("PartName=\"/xl/media/image1.png\""));
    }

    #[test]
    fn manifest_treats_case_equivalent_part_names_as_identical() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("XL/Workbook.xml", "workbook/type")
            .unwrap();
        manifest
            .register_part("xl/workbook.xml", "workbook/type")
            .unwrap();
        assert!(manifest
            .register_part("xl/workbook.xml", "different/type")
            .is_err());
    }

    #[test]
    fn raw_parts_cannot_claim_generated_paths() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("xl/charts/chart1.xml", "chart/type")
            .unwrap();
        assert!(manifest
            .register_raw_part("XL/charts/chart1.xml", "chart/type")
            .is_err());
    }

    #[test]
    fn generated_parts_cannot_claim_raw_paths() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_raw_part("xl/metadata.xml", "metadata/type")
            .unwrap();
        assert!(manifest
            .register_part("xl/metadata.xml", "metadata/type")
            .is_err());
    }

    #[test]
    fn raw_parts_cannot_claim_relationship_paths() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("xl/workbook.xml", "workbook/type")
            .unwrap();
        manifest
            .register_raw_part(
                "xl/_rels/workbook.xml.rels",
                "application/vnd.openxmlformats-package.relationships+xml",
            )
            .unwrap();
        assert!(manifest
            .register_relationship(
                Some("xl/workbook.xml"),
                "rId1",
                RelationshipKind::Styles.uri(),
                "styles.xml",
                false,
            )
            .is_err());
    }

    #[test]
    fn extensionless_parts_use_overrides() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("xl/vendor.v1/blob", "application/octet-stream")
            .unwrap();
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        manifest.write_content_types(&mut zip).unwrap();
        let mut archive = zip::ZipArchive::new(zip.finish().unwrap()).unwrap();
        let mut xml = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("PartName=\"/xl/vendor.v1/blob\""));
        assert!(!xml.contains("Extension=\"v1/blob\""));
    }

    #[test]
    fn manifest_serializes_registered_relationships() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("xl/workbook.xml", "workbook/type")
            .unwrap();
        manifest
            .register_part("xl/worksheets/sheet1.xml", "worksheet/type")
            .unwrap();
        manifest
            .register_relationship(
                None,
                "rId1",
                RelationshipKind::OfficeDocument.uri(),
                "xl/workbook.xml",
                false,
            )
            .unwrap();
        manifest
            .register_relationship(
                Some("xl/workbook.xml"),
                "rId1",
                RelationshipKind::Worksheet.uri(),
                "worksheets/sheet1.xml",
                false,
            )
            .unwrap();

        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        manifest.write_relationships(&mut zip).unwrap();
        let mut archive = zip::ZipArchive::new(zip.finish().unwrap()).unwrap();
        let mut root = String::new();
        archive
            .by_name("_rels/.rels")
            .unwrap()
            .read_to_string(&mut root)
            .unwrap();
        assert!(root.contains(RelationshipKind::OfficeDocument.uri()));
        let mut workbook = String::new();
        archive
            .by_name("xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut workbook)
            .unwrap();
        assert!(workbook.contains("Target=\"worksheets/sheet1.xml\""));
    }

    /// A case-insensitive reader cannot catch a mis-cased `.rels` name,
    /// so assert the ZIP entry spelling directly.
    #[test]
    fn relationship_part_names_preserve_owner_casing() {
        let mut manifest = OpcManifest::new().unwrap();
        manifest
            .register_part("xl/charts/chartEx1.xml", "chartex/type")
            .unwrap();
        manifest
            .register_part("xl/charts/style1.xml", "chartstyle/type")
            .unwrap();
        manifest
            .register_relationship(
                Some("xl/charts/chartEx1.xml"),
                "rId1",
                RelationshipKind::ChartStyle.uri(),
                "style1.xml",
                false,
            )
            .unwrap();

        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        manifest.write_relationships(&mut zip).unwrap();
        let archive = zip::ZipArchive::new(zip.finish().unwrap()).unwrap();
        let names: Vec<_> = archive.file_names().collect();
        assert_eq!(names, vec!["xl/charts/_rels/chartEx1.xml.rels"]);
    }
}
