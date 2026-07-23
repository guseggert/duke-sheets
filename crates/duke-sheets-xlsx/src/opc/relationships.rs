use std::collections::HashMap;
use std::io::{BufRead, Cursor};

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::diagnostics::{DiagnosticSink, XlsxDiagnosticCode};
use super::part_name::{resolve_internal_target_with_policy, PartName};
use crate::error::{XlsxError, XlsxResult};

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) enum RelationshipSource {
    Package,
    Part(PartName),
}

impl RelationshipSource {
    pub(crate) fn part_name(&self) -> Option<&PartName> {
        match self {
            Self::Package => None,
            Self::Part(part_name) => Some(part_name),
        }
    }

    pub(crate) fn relationships_part(&self) -> XlsxResult<PartName> {
        match self {
            Self::Package => PartName::new("/_rels/.rels"),
            Self::Part(part_name) => part_name.relationships_part(),
        }
    }

    pub(crate) fn display_name(&self) -> &str {
        self.part_name().map(PartName::as_str).unwrap_or("/")
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum RelationshipTarget {
    Internal(PartName),
    External(String),
    UnresolvedInternal,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Relationship {
    pub(crate) id: String,
    pub(crate) rel_type: String,
    pub(crate) raw_target: String,
    pub(crate) target: RelationshipTarget,
}

impl Relationship {
    pub(crate) fn internal_part(&self) -> Option<&PartName> {
        match &self.target {
            RelationshipTarget::Internal(part_name) => Some(part_name),
            RelationshipTarget::External(_) | RelationshipTarget::UnresolvedInternal => None,
        }
    }

    pub(crate) fn target(&self) -> &str {
        match &self.target {
            RelationshipTarget::Internal(part_name) => part_name.zip_name(),
            RelationshipTarget::External(target) => target,
            RelationshipTarget::UnresolvedInternal => &self.raw_target,
        }
    }

    pub(crate) fn internal_path(&self) -> Option<&str> {
        self.internal_part().map(PartName::zip_name)
    }

    pub(crate) fn is_external(&self) -> bool {
        matches!(self.target, RelationshipTarget::External(_))
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct RelationshipSet {
    relationships: Vec<Relationship>,
    by_id: HashMap<String, usize>,
}

impl RelationshipSet {
    pub(crate) fn parse<R: BufRead>(
        mut reader: R,
        source: &RelationshipSource,
        diagnostics: &mut DiagnosticSink,
    ) -> XlsxResult<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        validate_well_formed_xml(&bytes)?;
        if diagnostics.policy() == super::diagnostics::XlsxPackagePolicy::Strict {
            validate_relationships_structure(&bytes)?;
        }
        let mut xml_reader = NsReader::from_reader(Cursor::new(bytes));
        xml_reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut relationships = Self::default();
        let mut root_seen = false;

        loop {
            match xml_reader.read_resolved_event_into(&mut buf) {
                Ok((namespace, Event::Empty(element))) | Ok((namespace, Event::Start(element)))
                    if element.name().local_name().as_ref() == b"Relationships" =>
                {
                    root_seen = true;
                    if !namespace_is(
                        &namespace,
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                    ) {
                        diagnostics.violation(
                            XlsxDiagnosticCode::MalformedRelationship,
                            format!(
                                "relationships root for {} has the wrong namespace",
                                source.display_name()
                            ),
                            source.part_name().map(PartName::as_str),
                            None,
                            None,
                        )?;
                    }
                }
                Ok((namespace, Event::Empty(element))) | Ok((namespace, Event::Start(element)))
                    if element.name().local_name().as_ref() == b"Relationship" =>
                {
                    if !namespace_is(
                        &namespace,
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                    ) {
                        diagnostics.violation(
                            XlsxDiagnosticCode::MalformedRelationship,
                            format!(
                                "relationship in {} has the wrong namespace",
                                source.display_name()
                            ),
                            source.part_name().map(PartName::as_str),
                            None,
                            None,
                        )?;
                        buf.clear();
                        continue;
                    }
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
                                target_mode =
                                    attr.unescape_value().ok().map(|value| value.to_string())
                            }
                            _ => {}
                        }
                    }

                    let (Some(id), Some(raw_target), Some(rel_type)) = (id, target, rel_type)
                    else {
                        diagnostics.violation(
                            XlsxDiagnosticCode::MalformedRelationship,
                            format!(
                                "relationship in {} is missing Id, Type, or Target",
                                source.display_name()
                            ),
                            source.part_name().map(PartName::as_str),
                            None,
                            None,
                        )?;
                        buf.clear();
                        continue;
                    };
                    if raw_target.is_empty()
                        || (diagnostics.policy() == super::diagnostics::XlsxPackagePolicy::Strict
                            && !valid_relationship_type(&rel_type))
                    {
                        diagnostics.violation(
                            XlsxDiagnosticCode::MalformedRelationship,
                            format!("relationship {id} has invalid Id, Type, or Target"),
                            source.part_name().map(PartName::as_str),
                            Some(&id),
                            Some(&raw_target),
                        )?;
                        buf.clear();
                        continue;
                    }

                    if relationships.by_id.contains_key(&id) {
                        diagnostics.violation(
                            XlsxDiagnosticCode::DuplicateRelationshipId,
                            format!(
                                "duplicate relationship id {id} in {}",
                                source.display_name()
                            ),
                            source.part_name().map(PartName::as_str),
                            Some(&id),
                            Some(&raw_target),
                        )?;
                        buf.clear();
                        continue;
                    }

                    // TargetMode values are matched case-insensitively for
                    // compatibility; non-canonical casing is a violation.
                    let mode = target_mode.as_deref();
                    if mode.is_some_and(|mode| {
                        (mode.eq_ignore_ascii_case("Internal") && mode != "Internal")
                            || (mode.eq_ignore_ascii_case("External") && mode != "External")
                    }) {
                        diagnostics.violation(
                            XlsxDiagnosticCode::UnknownTargetMode,
                            format!(
                                "non-canonical relationship TargetMode {}",
                                mode.unwrap_or_default()
                            ),
                            source.part_name().map(PartName::as_str),
                            Some(&id),
                            Some(&raw_target),
                        )?;
                    }
                    let parsed_target = match mode {
                        None => Self::resolve_internal(source, &id, &raw_target, diagnostics)?,
                        Some(mode) if mode.eq_ignore_ascii_case("Internal") => {
                            Self::resolve_internal(source, &id, &raw_target, diagnostics)?
                        }
                        Some(mode) if mode.eq_ignore_ascii_case("External") => {
                            RelationshipTarget::External(raw_target.clone())
                        }
                        Some(mode) => {
                            diagnostics.violation(
                                XlsxDiagnosticCode::UnknownTargetMode,
                                format!("unknown relationship TargetMode {mode}"),
                                source.part_name().map(PartName::as_str),
                                Some(&id),
                                Some(&raw_target),
                            )?;
                            RelationshipTarget::UnresolvedInternal
                        }
                    };
                    relationships
                        .by_id
                        .insert(id.clone(), relationships.relationships.len());
                    relationships.relationships.push(Relationship {
                        id,
                        rel_type,
                        raw_target,
                        target: parsed_target,
                    });
                }
                Ok((_, Event::Eof)) => break,
                Err(error) => return Err(XlsxError::Xml(error)),
                _ => {}
            }
            buf.clear();
        }
        if !root_seen {
            diagnostics.violation(
                XlsxDiagnosticCode::MalformedRelationship,
                format!(
                    "relationships part for {} has no Relationships root",
                    source.display_name()
                ),
                source.part_name().map(PartName::as_str),
                None,
                None,
            )?;
        }
        Ok(relationships)
    }

    fn resolve_internal(
        source: &RelationshipSource,
        id: &str,
        raw_target: &str,
        diagnostics: &mut DiagnosticSink,
    ) -> XlsxResult<RelationshipTarget> {
        match resolve_internal_target_with_policy(
            source.part_name(),
            raw_target,
            diagnostics.policy() == super::diagnostics::XlsxPackagePolicy::Compatible,
        ) {
            Ok(part_name) => Ok(RelationshipTarget::Internal(part_name)),
            Err(error) => {
                diagnostics.violation(
                    XlsxDiagnosticCode::UnresolvedRelationshipTarget,
                    error.to_string(),
                    source.part_name().map(PartName::as_str),
                    Some(id),
                    Some(raw_target),
                )?;
                Ok(RelationshipTarget::UnresolvedInternal)
            }
        }
    }

    #[cfg(test)]
    pub(crate) fn get(&self, id: &str) -> Option<&Relationship> {
        self.by_id
            .get(id)
            .and_then(|index| self.relationships.get(*index))
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.relationships.iter()
    }

    pub(crate) fn by_type<'a>(
        &'a self,
        rel_types: &'a [&'a str],
    ) -> impl Iterator<Item = &'a Relationship> + 'a {
        self.relationships.iter().filter(move |relationship| {
            rel_types
                .iter()
                .any(|rel_type| relationship.rel_type == *rel_type)
        })
    }

    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.relationships.len()
    }
}

fn namespace_is(resolution: &ResolveResult<'_>, namespace: &str) -> bool {
    matches!(resolution, ResolveResult::Bound(actual) if actual.as_ref() == namespace.as_bytes())
}

fn validate_well_formed_xml(bytes: &[u8]) -> XlsxResult<()> {
    let mut reader = quick_xml::Reader::from_reader(Cursor::new(bytes));
    let mut buf = Vec::new();
    let mut open_elements: Vec<Vec<u8>> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => open_elements.push(element.name().as_ref().to_vec()),
            Ok(Event::End(element)) => {
                let expected = open_elements.pop().ok_or_else(|| {
                    XlsxError::InvalidFormat("unexpected closing XML element".into())
                })?;
                if expected != element.name().as_ref() {
                    return Err(XlsxError::InvalidFormat(
                        "mismatched closing XML element".into(),
                    ));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(XlsxError::Xml(error)),
            _ => {}
        }
        buf.clear();
    }
    if !open_elements.is_empty() {
        return Err(XlsxError::InvalidFormat(
            "unclosed XML element at end of stream".into(),
        ));
    }
    Ok(())
}

fn validate_relationships_structure(bytes: &[u8]) -> XlsxResult<()> {
    const NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    let mut reader = NsReader::from_reader(Cursor::new(bytes));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        match reader.read_resolved_event_into(&mut buf) {
            Ok((namespace, Event::Start(element))) => {
                if root_closed
                    || (depth == 0
                        && (element.name().local_name().as_ref() != b"Relationships"
                            || !namespace_is(&namespace, NS)
                            || !only_namespace_attributes(&element)))
                    || (depth == 1
                        && (element.name().local_name().as_ref() != b"Relationship"
                            || !namespace_is(&namespace, NS)
                            || !valid_relationship_attributes(&element)))
                    || depth > 1
                {
                    return Err(XlsxError::InvalidFormat(
                        "invalid Relationships part structure".into(),
                    ));
                }
                depth += 1;
            }
            Ok((namespace, Event::Empty(element))) => {
                if depth != 1
                    || element.name().local_name().as_ref() != b"Relationship"
                    || !namespace_is(&namespace, NS)
                    || !valid_relationship_attributes(&element)
                {
                    return Err(XlsxError::InvalidFormat(
                        "invalid Relationships part structure".into(),
                    ));
                }
            }
            Ok((_, Event::End(_))) => {
                if depth == 0 {
                    return Err(XlsxError::InvalidFormat(
                        "invalid Relationships part structure".into(),
                    ));
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            }
            Ok((_, Event::Text(text))) => {
                if !text.unescape().is_ok_and(|value| value.trim().is_empty()) {
                    return Err(XlsxError::InvalidFormat(
                        "invalid text in Relationships part".into(),
                    ));
                }
            }
            Ok((_, Event::Decl(_) | Event::Comment(_) | Event::PI(_))) => {}
            Ok((_, Event::Eof)) => break,
            Ok(_) => {
                return Err(XlsxError::InvalidFormat(
                    "invalid Relationships part structure".into(),
                ));
            }
            Err(error) => return Err(XlsxError::Xml(error)),
        }
        buf.clear();
    }
    if !root_closed {
        return Err(XlsxError::InvalidFormat(
            "Relationships part has no complete Relationships root".into(),
        ));
    }
    Ok(())
}

fn only_namespace_attributes(element: &quick_xml::events::BytesStart<'_>) -> bool {
    element
        .attributes()
        .all(|attr| attr.is_ok_and(|attr| attr.key.as_ref().starts_with(b"xmlns")))
}

fn valid_relationship_attributes(element: &quick_xml::events::BytesStart<'_>) -> bool {
    let mut id = false;
    let mut rel_type = false;
    let mut target = false;
    let mut target_mode = false;
    for attr in element.attributes() {
        let Ok(attr) = attr else {
            return false;
        };
        if attr.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        match attr.key.as_ref() {
            b"Id" if !id => id = true,
            b"Type" if !rel_type => rel_type = true,
            b"Target" if !target => target = true,
            b"TargetMode" if !target_mode => target_mode = true,
            _ => return false,
        }
    }
    id && rel_type && target
}

fn valid_relationship_type(value: &str) -> bool {
    value.split_once(':').is_some_and(|(scheme, rest)| {
        let mut characters = scheme.chars();
        characters
            .next()
            .is_some_and(|first| first.is_ascii_alphabetic())
            && characters.all(|character| {
                character.is_ascii_alphanumeric() || matches!(character, '+' | '-' | '.')
            })
            && !rest.is_empty()
    }) && !value.chars().any(char::is_whitespace)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::opc::diagnostics::XlsxPackagePolicy;

    #[test]
    fn relationships_preserve_document_order() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="urn:second" Target="b.xml"/><Relationship Id="rId1" Type="urn:first" Target="a.xml"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
        let source = RelationshipSource::Part(PartName::new("/xl/workbook.xml").unwrap());
        let relationships =
            RelationshipSet::parse(Cursor::new(xml), &source, &mut diagnostics).unwrap();
        assert_eq!(
            relationships
                .iter()
                .map(|relationship| relationship.id.as_str())
                .collect::<Vec<_>>(),
            vec!["rId2", "rId1"]
        );
    }

    #[test]
    fn duplicate_ids_keep_the_first_relationship_in_compatible_mode() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:first" Target="a.xml"/><Relationship Id="rId1" Type="urn:second" Target="b.xml"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
        let source = RelationshipSource::Package;
        let relationships =
            RelationshipSet::parse(Cursor::new(xml), &source, &mut diagnostics).unwrap();
        assert_eq!(relationships.len(), 1);
        assert_eq!(relationships.get("rId1").unwrap().rel_type, "urn:first");
        assert_eq!(diagnostics.into_diagnostics().len(), 1);
    }

    #[test]
    fn duplicate_ids_fail_in_strict_mode() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:first" Target="a.xml"/><Relationship Id="rId1" Type="urn:second" Target="b.xml"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Strict);
        assert!(RelationshipSet::parse(
            Cursor::new(xml),
            &RelationshipSource::Package,
            &mut diagnostics
        )
        .is_err());
    }

    #[test]
    fn compatible_mode_accepts_noncanonical_external_target_mode() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:link" Target="https://example.com" TargetMode="external"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
        let relationships = RelationshipSet::parse(
            Cursor::new(xml),
            &RelationshipSource::Package,
            &mut diagnostics,
        )
        .unwrap();
        assert!(relationships.get("rId1").unwrap().is_external());
        assert_eq!(diagnostics.into_diagnostics().len(), 1);
    }

    #[test]
    fn strict_mode_rejects_noncanonical_target_mode() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:link" Target="https://example.com" TargetMode="external"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Strict);
        assert!(RelationshipSet::parse(
            Cursor::new(xml),
            &RelationshipSource::Package,
            &mut diagnostics,
        )
        .is_err());
    }

    #[test]
    fn strict_mode_rejects_wrong_relationship_namespace() {
        let xml = br#"<Relationships xmlns="urn:not-opc"><Relationship Id="rId1" Type="urn:type" Target="a.xml"/></Relationships>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Strict);
        assert!(RelationshipSet::parse(
            Cursor::new(xml),
            &RelationshipSource::Package,
            &mut diagnostics
        )
        .is_err());
    }
}
