#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ContentTypeExpectation {
    Exact(&'static str),
    Prefix(&'static str),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RelationshipKind {
    OfficeDocument,
    Worksheet,
    Chartsheet,
    Dialogsheet,
    MacroSheet,
    IntlMacroSheet,
    Styles,
    SharedStrings,
    Theme,
    Comments,
    VmlDrawing,
    Hyperlink,
    Table,
    ControlProperties,
    SheetMetadata,
    Drawing,
    Chart,
    ChartEx,
    ChartStyle,
    ChartColorStyle,
    Image,
    DiagramData,
    DiagramLayout,
    DiagramQuickStyle,
    DiagramColors,
}

impl RelationshipKind {
    const ALL: [Self; 25] = [
        Self::OfficeDocument,
        Self::Worksheet,
        Self::Chartsheet,
        Self::Dialogsheet,
        Self::MacroSheet,
        Self::IntlMacroSheet,
        Self::Styles,
        Self::SharedStrings,
        Self::Theme,
        Self::Comments,
        Self::VmlDrawing,
        Self::Hyperlink,
        Self::Table,
        Self::ControlProperties,
        Self::SheetMetadata,
        Self::Drawing,
        Self::Chart,
        Self::ChartEx,
        Self::ChartStyle,
        Self::ChartColorStyle,
        Self::Image,
        Self::DiagramData,
        Self::DiagramLayout,
        Self::DiagramQuickStyle,
        Self::DiagramColors,
    ];

    pub(crate) fn from_uri(uri: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|kind| uri == kind.uri() || kind.strict_uri().is_some_and(|strict| uri == strict))
    }

    pub(crate) const fn uri(self) -> &'static str {
        match self {
            Self::OfficeDocument => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            Self::Worksheet => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            Self::Chartsheet => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
            Self::Dialogsheet => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
            Self::MacroSheet => "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
            Self::IntlMacroSheet => "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
            Self::Styles => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            Self::SharedStrings => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            Self::Theme => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            Self::Comments => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            Self::VmlDrawing => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            Self::Hyperlink => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            Self::Table => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            Self::ControlProperties => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
            Self::SheetMetadata => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
            Self::Drawing => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            Self::Chart => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            Self::ChartEx => "http://schemas.microsoft.com/office/2014/relationships/chartEx",
            Self::ChartStyle => "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
            Self::ChartColorStyle => "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
            Self::Image => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Self::DiagramData => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
            Self::DiagramLayout => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
            Self::DiagramQuickStyle => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
            Self::DiagramColors => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
        }
    }

    pub(crate) const fn strict_uri(self) -> Option<&'static str> {
        match self {
            Self::OfficeDocument => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
            }
            Self::Worksheet => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet")
            }
            Self::Chartsheet => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet")
            }
            Self::Dialogsheet => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet")
            }
            Self::Styles => Some("http://purl.oclc.org/ooxml/officeDocument/relationships/styles"),
            Self::SharedStrings => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings")
            }
            Self::Theme => Some("http://purl.oclc.org/ooxml/officeDocument/relationships/theme"),
            Self::Comments => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/comments")
            }
            Self::VmlDrawing => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing")
            }
            Self::Hyperlink => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink")
            }
            Self::Table => Some("http://purl.oclc.org/ooxml/officeDocument/relationships/table"),
            Self::ControlProperties => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/ctrlProp")
            }
            Self::SheetMetadata => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata")
            }
            Self::Drawing => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/drawing")
            }
            Self::Chart => Some("http://purl.oclc.org/ooxml/officeDocument/relationships/chart"),
            Self::Image => Some("http://purl.oclc.org/ooxml/officeDocument/relationships/image"),
            Self::DiagramData => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData")
            }
            Self::DiagramLayout => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout")
            }
            Self::DiagramQuickStyle => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle")
            }
            Self::DiagramColors => {
                Some("http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors")
            }
            Self::MacroSheet
            | Self::IntlMacroSheet
            | Self::ChartEx
            | Self::ChartStyle
            | Self::ChartColorStyle => None,
        }
    }

    pub(crate) const fn requires_internal_target(self) -> bool {
        !matches!(self, Self::Hyperlink)
    }

    pub(crate) const fn unmodeled_sheet_label(self) -> Option<&'static str> {
        match self {
            Self::Dialogsheet => Some("Dialog"),
            Self::MacroSheet => Some("Macro"),
            Self::IntlMacroSheet => Some("International macro"),
            _ => None,
        }
    }

    pub(crate) const fn content_type(self) -> Option<ContentTypeExpectation> {
        use ContentTypeExpectation::{Exact, Prefix};
        match self {
            Self::Worksheet => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )),
            Self::Chartsheet => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
            )),
            Self::Dialogsheet => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
            )),
            Self::MacroSheet => Some(Exact("application/vnd.ms-excel.macrosheet+xml")),
            Self::IntlMacroSheet => Some(Exact("application/vnd.ms-excel.intlmacrosheet+xml")),
            Self::Styles => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            )),
            Self::SharedStrings => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            )),
            Self::Theme => Some(Exact(
                "application/vnd.openxmlformats-officedocument.theme+xml",
            )),
            Self::Comments => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            )),
            Self::VmlDrawing => Some(Exact(
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
            )),
            Self::Table => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            )),
            Self::ControlProperties => {
                Some(Exact("application/vnd.ms-excel.controlproperties+xml"))
            }
            Self::SheetMetadata => Some(Exact(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml",
            )),
            Self::Drawing => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawing+xml",
            )),
            Self::Chart => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            )),
            Self::ChartEx => Some(Exact("application/vnd.ms-office.chartex+xml")),
            Self::ChartStyle => Some(Exact("application/vnd.ms-office.chartstyle+xml")),
            Self::ChartColorStyle => Some(Exact("application/vnd.ms-office.chartcolorstyle+xml")),
            Self::Image => Some(Prefix("image/")),
            Self::DiagramData => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
            )),
            Self::DiagramLayout => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
            )),
            Self::DiagramQuickStyle => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
            )),
            Self::DiagramColors => Some(Exact(
                "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
            )),
            Self::OfficeDocument | Self::Hyperlink => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn recognizes_transitional_strict_and_extension_relationships() {
        assert_eq!(
            RelationshipKind::from_uri(RelationshipKind::Worksheet.uri()),
            Some(RelationshipKind::Worksheet)
        );
        assert_eq!(
            RelationshipKind::from_uri(RelationshipKind::Worksheet.strict_uri().unwrap()),
            Some(RelationshipKind::Worksheet)
        );
        assert_eq!(
            RelationshipKind::from_uri(RelationshipKind::ChartEx.uri()),
            Some(RelationshipKind::ChartEx)
        );
    }
}
