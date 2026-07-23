use std::hash::{Hash, Hasher};

use crate::error::{XlsxError, XlsxResult};

/// Canonical abstract OPC part name, including its package-root slash.
#[derive(Debug, Clone)]
pub(crate) struct PartName(String);

impl PartName {
    pub(crate) fn new(value: impl Into<String>) -> XlsxResult<Self> {
        let value = value.into();
        validate_part_name(&value)?;
        Ok(Self(value))
    }

    pub(crate) fn from_zip_name(value: &str) -> XlsxResult<Self> {
        Self::from_zip_name_with_policy(value, true)
    }

    pub(crate) fn from_zip_name_with_policy(value: &str, compatible: bool) -> XlsxResult<Self> {
        if !compatible && (value.starts_with('/') || value.contains('\\')) {
            return Err(XlsxError::InvalidFormat(format!(
                "invalid OPC ZIP item name: {value}"
            )));
        }
        let normalized = value.replace('\\', "/");
        Self::new(format!("/{}", normalized.trim_start_matches('/')))
    }

    pub(crate) fn as_str(&self) -> &str {
        &self.0
    }

    pub(crate) fn zip_name(&self) -> &str {
        self.0.trim_start_matches('/')
    }

    pub(crate) fn extension(&self) -> Option<&str> {
        self.0
            .rsplit('/')
            .next()
            .and_then(|file| file.rsplit_once('.'))
            .map(|(_, extension)| extension)
    }

    pub(crate) fn parent(&self) -> Option<&str> {
        self.0.rsplit_once('/').map(|(parent, _)| parent)
    }

    pub(crate) fn relationships_part(&self) -> XlsxResult<Self> {
        let (parent, file) = self
            .0
            .rsplit_once('/')
            .ok_or_else(|| XlsxError::InvalidFormat(format!("invalid part name: {}", self.0)))?;
        let path = if parent.is_empty() {
            format!("/_rels/{file}.rels")
        } else {
            format!("{parent}/_rels/{file}.rels")
        };
        Self::new(path)
    }
}

impl PartialEq for PartName {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(&other.0)
    }
}

impl Eq for PartName {}

impl Hash for PartName {
    fn hash<H: Hasher>(&self, state: &mut H) {
        for byte in self.0.bytes() {
            state.write_u8(byte.to_ascii_lowercase());
        }
    }
}

impl std::fmt::Display for PartName {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

fn validate_part_name(value: &str) -> XlsxResult<()> {
    if value == "/" || !value.starts_with('/') || value.contains(['?', '#', '\\']) {
        return Err(XlsxError::InvalidFormat(format!(
            "invalid OPC part name: {value}"
        )));
    }

    for segment in value[1..].split('/') {
        if segment.is_empty() || matches!(segment, "." | "..") || segment.ends_with('.') {
            return Err(XlsxError::InvalidFormat(format!(
                "invalid OPC part name segment in {value}"
            )));
        }
        if segment.chars().any(|character| {
            character.is_control()
                || character == ' '
                || matches!(
                    character,
                    '[' | ']' | '<' | '>' | '"' | '{' | '}' | '|' | '^' | '`'
                )
                || private_or_noncharacter(character)
        }) {
            return Err(XlsxError::InvalidFormat(format!(
                "invalid character in OPC part name: {value}"
            )));
        }
        validate_percent_encoding(segment, value)?;
    }
    Ok(())
}

fn private_or_noncharacter(character: char) -> bool {
    let value = character as u32;
    matches!(value, 0xE000..=0xF8FF | 0xF0000..=0xFFFFD | 0x100000..=0x10FFFD)
        || matches!(value, 0xFDD0..=0xFDEF)
        || value & 0xFFFF == 0xFFFE
        || value & 0xFFFF == 0xFFFF
}

fn validate_percent_encoding(segment: &str, part_name: &str) -> XlsxResult<()> {
    let bytes = segment.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Err(XlsxError::InvalidFormat(format!(
                "invalid percent encoding in OPC part name: {part_name}"
            )));
        }
        let high = hex_value(bytes[index + 1]);
        let low = hex_value(bytes[index + 2]);
        let Some(byte) = high.zip(low).map(|(high, low)| high * 16 + low) else {
            return Err(XlsxError::InvalidFormat(format!(
                "invalid percent encoding in OPC part name: {part_name}"
            )));
        };
        if byte == b'/'
            || byte == b'\\'
            || byte.is_ascii_alphanumeric()
            || matches!(byte, b'-' | b'.' | b'_' | b'~')
        {
            return Err(XlsxError::InvalidFormat(format!(
                "forbidden percent encoding in OPC part name: {part_name}"
            )));
        }
        index += 3;
    }
    Ok(())
}

pub(crate) fn resolve_internal_target(
    source: Option<&PartName>,
    target: &str,
) -> XlsxResult<PartName> {
    resolve_internal_target_with_policy(source, target, true)
}

pub(crate) fn resolve_internal_target_with_policy(
    source: Option<&PartName>,
    target: &str,
    compatible: bool,
) -> XlsxResult<PartName> {
    if !compatible && target.contains('\\') {
        return Err(XlsxError::InvalidFormat(format!(
            "internal relationship target contains a backslash: {target}"
        )));
    }
    let normalized_target = target.replace('\\', "/");
    let target_path = decode_percent_encoded_unreserved(
        normalized_target
            .split(['?', '#'])
            .next()
            .unwrap_or_default(),
    );
    if !target_path.starts_with('/')
        && target_path
            .split('/')
            .next()
            .is_some_and(|segment| segment.contains(':'))
    {
        return Err(XlsxError::InvalidFormat(format!(
            "internal relationship target is an absolute IRI: {target}"
        )));
    }
    if target_path.starts_with("//") {
        return Err(XlsxError::InvalidFormat(format!(
            "internal relationship target has an authority: {target}"
        )));
    }
    if target_path.is_empty() {
        return source.cloned().ok_or_else(|| {
            XlsxError::InvalidFormat("package relationship target does not name a part".into())
        });
    }

    let absolute = target_path.starts_with('/');
    let mut parts: Vec<&str> = if absolute {
        Vec::new()
    } else {
        source
            .and_then(PartName::parent)
            .map(|parent| {
                parent
                    .trim_start_matches('/')
                    .split('/')
                    .filter(|part| !part.is_empty())
                    .collect()
            })
            .unwrap_or_default()
    };

    for part in target_path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(XlsxError::InvalidFormat(format!(
                        "relationship target escapes package root: target={target}"
                    )));
                }
            }
            _ => parts.push(part),
        }
    }

    if parts.is_empty() {
        return Err(XlsxError::InvalidFormat(format!(
            "relationship target does not name a package part: target={target}"
        )));
    }
    PartName::new(format!("/{}", parts.join("/")))
}

fn decode_percent_encoded_unreserved(value: &str) -> String {
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

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn part_names_are_ascii_case_insensitive() {
        assert_eq!(
            PartName::new("/XL/Workbook.xml").unwrap(),
            PartName::new("/xl/workbook.xml").unwrap()
        );
    }

    #[test]
    fn invalid_part_names_are_rejected() {
        for value in [
            "xl/workbook.xml",
            "/",
            "/xl//workbook.xml",
            "/xl/../workbook.xml",
            "/xl/workbook.",
            "/xl/workbook%31.xml",
            "/xl/a%2fb.xml",
        ] {
            assert!(PartName::new(value).is_err(), "{value}");
        }
    }

    #[test]
    fn extension_comes_only_from_the_final_segment() {
        assert_eq!(
            PartName::new("/xl/media/image.png").unwrap().extension(),
            Some("png")
        );
        assert_eq!(
            PartName::new("/xl/media.v2/blob").unwrap().extension(),
            None
        );
    }

    #[test]
    fn relationship_part_name_is_derived_from_its_owner() {
        assert_eq!(
            PartName::new("/xl/workbook.xml")
                .unwrap()
                .relationships_part()
                .unwrap()
                .as_str(),
            "/xl/_rels/workbook.xml.rels"
        );
        assert_eq!(
            PartName::new("/workbook.xml")
                .unwrap()
                .relationships_part()
                .unwrap()
                .as_str(),
            "/_rels/workbook.xml.rels"
        );
    }

    #[test]
    fn relationship_targets_resolve_to_part_names() {
        let source = PartName::new("/xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(
            resolve_internal_target(Some(&source), "../drawings/./drawing1.xml")
                .unwrap()
                .as_str(),
            "/xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_internal_target(Some(&source), "/xl/drawings/drawing%31.xml")
                .unwrap()
                .as_str(),
            "/xl/drawings/drawing1.xml"
        );
        assert!(resolve_internal_target(Some(&source), "//host/drawing.xml").is_err());
    }
}
