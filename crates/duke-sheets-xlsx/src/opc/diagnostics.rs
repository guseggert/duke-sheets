use crate::error::{XlsxError, XlsxResult};

/// Controls how strictly traversed OPC package structures are validated.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum XlsxPackagePolicy {
    /// Recover from common package defects when the intended part is unambiguous.
    #[default]
    Compatible,
    /// Reject OPC violations encountered while traversing the workbook graph.
    /// Valid but unmodeled features (for example dialog sheets) are reported
    /// as warnings, not errors.
    Strict,
}

/// Severity of a package diagnostic emitted while reading an XLSX file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum XlsxDiagnosticSeverity {
    /// A package defect or unsupported construct was observed; reading
    /// continued, possibly without the affected data.
    Warning,
    /// A compatibility fallback was used to continue reading the workbook.
    Recovery,
}

/// Stable category for a package diagnostic.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum XlsxDiagnosticCode {
    /// A ZIP entry or target breaks the OPC part-name rules.
    InvalidPartName,
    /// Two entries differ only by ASCII case, so they name one part.
    EquivalentPartName,
    /// One part name is a prefix of another, which OPC forbids.
    DerivablePartName,
    /// A part has no `Default` or `Override` content type.
    MissingContentType,
    /// An extension or part name is declared more than once.
    DuplicateContentType,
    /// `[Content_Types].xml` is unreadable or does not match its schema.
    MalformedContentType,
    /// A part that owns relationships has no `.rels` part.
    MissingRelationshipsPart,
    /// A `Relationship` element is malformed or wrongly namespaced.
    MalformedRelationship,
    /// Two relationships from one source share an `Id`.
    DuplicateRelationshipId,
    /// `TargetMode` is neither `Internal` nor `External`, or is miscased.
    UnknownTargetMode,
    /// An internal target could not be resolved to a part name.
    UnresolvedRelationshipTarget,
    /// An internal target resolved but names no part in the package.
    MissingRelationshipTarget,
    /// The package has no `/_rels/.rels`.
    MissingPackageRelationships,
    /// Package relationships contain no `officeDocument` relationship.
    MissingOfficeDocumentRelationship,
    /// More than one candidate Workbook part was found.
    AmbiguousOfficeDocumentRelationship,
    /// The Workbook part's content type is not a SpreadsheetML workbook.
    WorkbookContentTypeMismatch,
    /// A part's content type does not match its relationship type.
    PartContentTypeMismatch,
    /// A part was located by convention because relationships were unusable.
    CanonicalPartFallback,
    /// A valid sheet kind this library does not model, such as a dialog
    /// or macro sheet, was skipped.
    UnsupportedSheetType,
}

/// A recoverable OPC package problem observed while reading an XLSX file.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct XlsxDiagnostic {
    /// Stable category, suitable for matching.
    pub code: XlsxDiagnosticCode,
    /// Whether data was skipped or a fallback was applied.
    pub severity: XlsxDiagnosticSeverity,
    /// Part that owns the problem, as an OPC part name.
    pub source_part: Option<String>,
    /// Relationship id, when the problem is relationship-scoped.
    pub relationship_id: Option<String>,
    /// Relationship target or part name the problem refers to.
    pub target: Option<String>,
    /// Human-readable explanation; not stable across versions.
    pub message: String,
}

#[derive(Debug)]
pub(crate) struct DiagnosticSink {
    policy: XlsxPackagePolicy,
    diagnostics: Vec<XlsxDiagnostic>,
}

impl DiagnosticSink {
    pub(crate) fn new(policy: XlsxPackagePolicy) -> Self {
        Self {
            policy,
            diagnostics: Vec::new(),
        }
    }

    pub(crate) fn policy(&self) -> XlsxPackagePolicy {
        self.policy
    }

    pub(crate) fn warning(
        &mut self,
        code: XlsxDiagnosticCode,
        message: impl Into<String>,
        source_part: Option<&str>,
        relationship_id: Option<&str>,
        target: Option<&str>,
    ) {
        self.push(
            code,
            XlsxDiagnosticSeverity::Warning,
            message,
            source_part,
            relationship_id,
            target,
        );
    }

    pub(crate) fn recovery(
        &mut self,
        code: XlsxDiagnosticCode,
        message: impl Into<String>,
        source_part: Option<&str>,
        relationship_id: Option<&str>,
        target: Option<&str>,
    ) {
        self.push(
            code,
            XlsxDiagnosticSeverity::Recovery,
            message,
            source_part,
            relationship_id,
            target,
        );
    }

    pub(crate) fn violation(
        &mut self,
        code: XlsxDiagnosticCode,
        message: impl Into<String>,
        source_part: Option<&str>,
        relationship_id: Option<&str>,
        target: Option<&str>,
    ) -> XlsxResult<()> {
        let message = message.into();
        if self.policy == XlsxPackagePolicy::Strict {
            return Err(XlsxError::InvalidFormat(message));
        }
        self.warning(code, message, source_part, relationship_id, target);
        Ok(())
    }

    pub(crate) fn into_diagnostics(self) -> Vec<XlsxDiagnostic> {
        self.diagnostics
    }

    fn push(
        &mut self,
        code: XlsxDiagnosticCode,
        severity: XlsxDiagnosticSeverity,
        message: impl Into<String>,
        source_part: Option<&str>,
        relationship_id: Option<&str>,
        target: Option<&str>,
    ) {
        self.diagnostics.push(XlsxDiagnostic {
            code,
            severity,
            source_part: source_part.map(str::to_string),
            relationship_id: relationship_id.map(str::to_string),
            target: target.map(str::to_string),
            message: message.into(),
        });
    }
}
