mod content_types;
mod diagnostics;
mod package;
mod part_name;
mod relationships;
mod xml;

use crate::error::XlsxResult;

pub use diagnostics::{
    XlsxDiagnostic, XlsxDiagnosticCode, XlsxDiagnosticSeverity, XlsxPackagePolicy,
};
pub(crate) use package::OpcPackage;
pub(crate) use part_name::PartName;
pub(crate) use relationships::{Relationship, RelationshipSource};

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
mod tests {
    use super::*;

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
}
