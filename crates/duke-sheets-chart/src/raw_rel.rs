//! A relationship captured from a part kept verbatim.

/// A captured relationship referenced by a a part kept verbatim.
#[derive(Debug, Clone, PartialEq)]
pub struct RawRel {
    /// Relationship id as it appears in the raw bytes (e.g. "rId7").
    pub id: String,
    /// Relationship type URI.
    pub rel_type: String,
    /// Relationship target.
    pub target: String,
    /// True when the target is external (TargetMode="External").
    pub external: bool,
    /// Bytes of the target part, when internal and captured.
    #[doc(hidden)]
    pub part: Option<Vec<u8>>,
    /// The target part's own relationships, captured with their targets.
    ///
    /// A preserved part is not always self-contained: a diagram's data
    /// part references its images, and a chart part references its
    /// style and colour parts. Writing the part back without these
    /// leaves those references dangling, and for a chartEx Excel
    /// refuses the workbook outright.
    #[doc(hidden)]
    pub part_rels: Vec<RawRel>,
}
