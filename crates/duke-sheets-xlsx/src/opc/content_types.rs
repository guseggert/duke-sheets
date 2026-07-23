use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read};

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::diagnostics::{DiagnosticSink, XlsxDiagnosticCode, XlsxPackagePolicy};
use super::part_name::PartName;
use crate::error::{XlsxError, XlsxResult};

#[derive(Debug, Clone, Default)]
pub(crate) struct ContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<PartName, String>,
}

impl ContentTypes {
    pub(crate) fn parse<R: Read>(
        mut reader: R,
        diagnostics: &mut DiagnosticSink,
    ) -> XlsxResult<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        validate_well_formed_xml(&bytes)?;
        if diagnostics.policy() == XlsxPackagePolicy::Strict {
            validate_content_types_structure(&bytes)?;
        }
        let mut xml_reader = NsReader::from_reader(BufReader::new(Cursor::new(bytes)));
        xml_reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut content_types = Self::default();
        let mut root_seen = false;

        loop {
            match xml_reader.read_resolved_event_into(&mut buf) {
                Ok((namespace, Event::Empty(element))) | Ok((namespace, Event::Start(element))) => {
                    if element.name().local_name().as_ref() == b"Types" {
                        root_seen = true;
                        if !namespace_is(
                            &namespace,
                            "http://schemas.openxmlformats.org/package/2006/content-types",
                        ) {
                            diagnostics.violation(
                                XlsxDiagnosticCode::MalformedContentType,
                                "content types root has the wrong namespace",
                                Some("/[Content_Types].xml"),
                                None,
                                None,
                            )?;
                        }
                        buf.clear();
                        continue;
                    }
                    if !namespace_is(
                        &namespace,
                        "http://schemas.openxmlformats.org/package/2006/content-types",
                    ) {
                        if matches!(
                            element.name().local_name().as_ref(),
                            b"Default" | b"Override"
                        ) {
                            diagnostics.violation(
                                XlsxDiagnosticCode::MalformedContentType,
                                "content type child has the wrong namespace",
                                Some("/[Content_Types].xml"),
                                None,
                                None,
                            )?;
                        }
                        buf.clear();
                        continue;
                    }
                    match element.name().local_name().as_ref() {
                        b"Default" => {
                            let mut extension = None;
                            let mut content_type = None;
                            for attr in element.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"Extension" => {
                                        extension = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|value| value.to_ascii_lowercase())
                                    }
                                    b"ContentType" => {
                                        content_type = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|value| value.to_string())
                                    }
                                    _ => {}
                                }
                            }
                            let (Some(extension), Some(content_type)) = (extension, content_type)
                            else {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::MalformedContentType,
                                    "content type Default is missing Extension or ContentType",
                                    Some("/[Content_Types].xml"),
                                    None,
                                    None,
                                )?;
                                buf.clear();
                                continue;
                            };
                            if extension.is_empty()
                                || extension.contains(['.', '/', '\\'])
                                || extension.chars().any(char::is_whitespace)
                                || !valid_media_type(&content_type)
                            {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::MalformedContentType,
                                    format!("invalid content type Default for .{extension}"),
                                    Some("/[Content_Types].xml"),
                                    None,
                                    None,
                                )?;
                                buf.clear();
                                continue;
                            }
                            if let Some(existing) = content_types.defaults.get(&extension) {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::DuplicateContentType,
                                    format!(
                                        "duplicate content type Default for .{extension}: {existing} and {content_type}"
                                    ),
                                    Some("/[Content_Types].xml"),
                                    None,
                                    None,
                                )?;
                            } else {
                                content_types.defaults.insert(extension, content_type);
                            }
                        }
                        b"Override" => {
                            let mut part_name = None;
                            let mut content_type = None;
                            for attr in element.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"PartName" => {
                                        part_name = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|value| value.to_string())
                                    }
                                    b"ContentType" => {
                                        content_type = attr
                                            .unescape_value()
                                            .ok()
                                            .map(|value| value.to_string())
                                    }
                                    _ => {}
                                }
                            }
                            let (Some(raw_part_name), Some(content_type)) =
                                (part_name, content_type)
                            else {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::MalformedContentType,
                                    "content type Override is missing PartName or ContentType",
                                    Some("/[Content_Types].xml"),
                                    None,
                                    None,
                                )?;
                                buf.clear();
                                continue;
                            };
                            if !valid_media_type(&content_type) {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::MalformedContentType,
                                    format!("invalid media type for {raw_part_name}"),
                                    Some("/[Content_Types].xml"),
                                    None,
                                    Some(&raw_part_name),
                                )?;
                                buf.clear();
                                continue;
                            }
                            let part_name = match PartName::new(raw_part_name.clone()) {
                                Ok(part_name) => part_name,
                                Err(error) => {
                                    diagnostics.violation(
                                        XlsxDiagnosticCode::InvalidPartName,
                                        error.to_string(),
                                        Some("/[Content_Types].xml"),
                                        None,
                                        Some(&raw_part_name),
                                    )?;
                                    buf.clear();
                                    continue;
                                }
                            };
                            if let Some(existing) = content_types.overrides.get(&part_name) {
                                diagnostics.violation(
                                    XlsxDiagnosticCode::DuplicateContentType,
                                    format!(
                                        "duplicate content type Override for {part_name}: {existing} and {content_type}"
                                    ),
                                    Some("/[Content_Types].xml"),
                                    None,
                                    Some(part_name.as_str()),
                                )?;
                            } else {
                                content_types.overrides.insert(part_name, content_type);
                            }
                        }
                        _ => {}
                    }
                }
                Ok((_, Event::Eof)) => break,
                Err(error) => return Err(XlsxError::Xml(error)),
                _ => {}
            }
            buf.clear();
        }
        if !root_seen {
            diagnostics.violation(
                XlsxDiagnosticCode::MalformedContentType,
                "content types stream has no Types root",
                Some("/[Content_Types].xml"),
                None,
                None,
            )?;
        }
        Ok(content_types)
    }

    pub(crate) fn content_type_for(&self, part_name: &PartName) -> Option<&str> {
        self.overrides
            .get(part_name)
            .map(String::as_str)
            .or_else(|| {
                part_name
                    .extension()
                    .and_then(|extension| self.defaults.get(&extension.to_ascii_lowercase()))
                    .map(String::as_str)
            })
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

fn validate_content_types_structure(bytes: &[u8]) -> XlsxResult<()> {
    const NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
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
                        && (element.name().local_name().as_ref() != b"Types"
                            || !namespace_is(&namespace, NS)))
                    || (depth == 1
                        && (!matches!(
                            element.name().local_name().as_ref(),
                            b"Default" | b"Override"
                        ) || !namespace_is(&namespace, NS)
                            || !valid_content_type_attributes(&element)))
                    || depth > 1
                {
                    return Err(XlsxError::InvalidFormat(
                        "invalid [Content_Types].xml structure".into(),
                    ));
                }
                depth += 1;
            }
            Ok((namespace, Event::Empty(element))) => {
                if depth != 1
                    || !matches!(
                        element.name().local_name().as_ref(),
                        b"Default" | b"Override"
                    )
                    || !namespace_is(&namespace, NS)
                    || !valid_content_type_attributes(&element)
                {
                    return Err(XlsxError::InvalidFormat(
                        "invalid [Content_Types].xml structure".into(),
                    ));
                }
            }
            Ok((_, Event::End(_))) => {
                if depth == 0 {
                    return Err(XlsxError::InvalidFormat(
                        "invalid [Content_Types].xml structure".into(),
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
                        "invalid text in [Content_Types].xml".into(),
                    ));
                }
            }
            Ok((_, Event::Decl(_) | Event::Comment(_) | Event::PI(_))) => {}
            Ok((_, Event::Eof)) => break,
            Ok(_) => {
                return Err(XlsxError::InvalidFormat(
                    "invalid [Content_Types].xml structure".into(),
                ));
            }
            Err(error) => return Err(XlsxError::Xml(error)),
        }
        buf.clear();
    }
    if !root_closed {
        return Err(XlsxError::InvalidFormat(
            "[Content_Types].xml has no complete Types root".into(),
        ));
    }
    Ok(())
}

fn valid_content_type_attributes(element: &quick_xml::events::BytesStart<'_>) -> bool {
    let default = element.name().local_name().as_ref() == b"Default";
    let mut first = false;
    let mut content_type = false;
    for attr in element.attributes() {
        let Ok(attr) = attr else {
            return false;
        };
        if attr.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        match attr.key.as_ref() {
            b"Extension" if default && !first => first = true,
            b"PartName" if !default && !first => first = true,
            b"ContentType" if !content_type => content_type = true,
            _ => return false,
        }
    }
    first && content_type
}

fn valid_media_type(value: &str) -> bool {
    if value != value.trim() {
        return false;
    }
    let mut sections = value.split(';');
    let Some(essence) = sections.next() else {
        return false;
    };
    let Some((kind, subtype)) = essence.split_once('/') else {
        return false;
    };
    if subtype.contains('/') || !valid_media_token(kind) || !valid_media_token(subtype) {
        return false;
    }
    sections.all(|parameter| {
        let parameter = parameter.trim();
        parameter
            .split_once('=')
            .is_some_and(|(name, value)| valid_media_token(name) && !value.trim().is_empty())
    })
}

fn valid_media_token(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(
                    byte,
                    b'!' | b'#'
                        | b'$'
                        | b'%'
                        | b'&'
                        | b'\''
                        | b'*'
                        | b'+'
                        | b'-'
                        | b'.'
                        | b'^'
                        | b'_'
                        | b'`'
                        | b'|'
                        | b'~'
                )
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn override_takes_precedence_over_default() {
        let xml = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="workbook/type"/></Types>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
        let content_types = ContentTypes::parse(xml.as_slice(), &mut diagnostics).unwrap();
        assert_eq!(
            content_types.content_type_for(&PartName::new("/xl/workbook.xml").unwrap()),
            Some("workbook/type")
        );
        assert_eq!(
            content_types.content_type_for(&PartName::new("/xl/other.xml").unwrap()),
            Some("application/xml")
        );
    }

    #[test]
    fn duplicate_defaults_are_deterministic_in_compatible_mode() {
        let xml = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="XML" ContentType="first/type"/><Default Extension="xml" ContentType="second/type"/></Types>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Compatible);
        let content_types = ContentTypes::parse(xml.as_slice(), &mut diagnostics).unwrap();
        assert_eq!(
            content_types.content_type_for(&PartName::new("/part.xml").unwrap()),
            Some("first/type")
        );
        assert_eq!(diagnostics.into_diagnostics().len(), 1);
    }

    #[test]
    fn duplicate_defaults_fail_in_strict_mode() {
        let xml = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="first/type"/><Default Extension="xml" ContentType="second/type"/></Types>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Strict);
        assert!(ContentTypes::parse(xml.as_slice(), &mut diagnostics).is_err());
    }

    #[test]
    fn strict_mode_rejects_wrong_content_types_namespace() {
        let xml = br#"<Types xmlns="urn:not-opc"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let mut diagnostics = DiagnosticSink::new(XlsxPackagePolicy::Strict);
        assert!(ContentTypes::parse(xml.as_slice(), &mut diagnostics).is_err());
    }
}
