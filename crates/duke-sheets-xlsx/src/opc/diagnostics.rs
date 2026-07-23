use crate::error::{XlsxError, XlsxResult};

/// Controls how strictly traversed OPC package structures are validated.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum XlsxPackagePolicy {
    /// Recover from common package defects when the intended part is unambiguous.
    #[default]
    Compatible,
    /// Reject OPC violations encountered while traversing the workbook graph.
    Strict,
}

/// Severity of a package diagnostic emitted while reading an XLSX file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsxDiagnosticSeverity {
    /// The package is malformed, but the affected optional data was skipped.
    Warning,
    /// A compatibility fallback was used to continue reading the workbook.
    Recovery,
}

/// Stable category for a package diagnostic.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsxDiagnosticCode {
    InvalidPartName,
    EquivalentPartName,
    DerivablePartName,
    MissingContentType,
    DuplicateContentType,
    MalformedContentType,
    MissingRelationshipsPart,
    MalformedRelationship,
    DuplicateRelationshipId,
    UnknownTargetMode,
    UnresolvedRelationshipTarget,
    MissingRelationshipTarget,
    MissingPackageRelationships,
    MissingOfficeDocumentRelationship,
    AmbiguousOfficeDocumentRelationship,
    WorkbookContentTypeMismatch,
    PartContentTypeMismatch,
    CanonicalPartFallback,
    UnsupportedSheetType,
}

/// A recoverable OPC package problem observed while reading an XLSX file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsxDiagnostic {
    pub code: XlsxDiagnosticCode,
    pub severity: XlsxDiagnosticSeverity,
    pub source_part: Option<String>,
    pub relationship_id: Option<String>,
    pub target: Option<String>,
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
