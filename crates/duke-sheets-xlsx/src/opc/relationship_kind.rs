#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ContentTypeExpectation {
    Exact(&'static str),
    Prefix(&'static str),
}

pub(crate) const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub(crate) const CT_TEMPLATE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
pub(crate) const CT_MACRO_WORKBOOK: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
pub(crate) const CT_MACRO_TEMPLATE: &str =
    "application/vnd.ms-excel.template.macroEnabled.main+xml";
pub(crate) const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
pub(crate) const CT_CHARTSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
pub(crate) const CT_DIALOGSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";
pub(crate) const CT_MACROSHEET: &str = "application/vnd.ms-excel.macrosheet+xml";
pub(crate) const CT_INTL_MACROSHEET: &str = "application/vnd.ms-excel.intlmacrosheet+xml";
pub(crate) const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
pub(crate) const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
pub(crate) const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
pub(crate) const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
pub(crate) const CT_VML_DRAWING: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
pub(crate) const CT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
pub(crate) const CT_CONTROL_PROPERTIES: &str = "application/vnd.ms-excel.controlproperties+xml";
pub(crate) const CT_SHEET_METADATA: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml";
pub(crate) const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(crate) const CT_CHART_EX: &str = "application/vnd.ms-office.chartex+xml";
pub(crate) const CT_CHART_STYLE: &str = "application/vnd.ms-office.chartstyle+xml";
pub(crate) const CT_CHART_COLOR_STYLE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub(crate) const CT_DIAGRAM_DATA: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
pub(crate) const CT_DIAGRAM_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
pub(crate) const CT_DIAGRAM_STYLE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
pub(crate) const CT_DIAGRAM_COLORS: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";

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
            Self::Worksheet => Some(Exact(CT_WORKSHEET)),
            Self::Chartsheet => Some(Exact(CT_CHARTSHEET)),
            Self::Dialogsheet => Some(Exact(CT_DIALOGSHEET)),
            Self::MacroSheet => Some(Exact(CT_MACROSHEET)),
            Self::IntlMacroSheet => Some(Exact(CT_INTL_MACROSHEET)),
            Self::Styles => Some(Exact(CT_STYLES)),
            Self::SharedStrings => Some(Exact(CT_SHARED_STRINGS)),
            Self::Theme => Some(Exact(CT_THEME)),
            Self::Comments => Some(Exact(CT_COMMENTS)),
            Self::VmlDrawing => Some(Exact(CT_VML_DRAWING)),
            Self::Table => Some(Exact(CT_TABLE)),
            Self::ControlProperties => Some(Exact(CT_CONTROL_PROPERTIES)),
            Self::SheetMetadata => Some(Exact(CT_SHEET_METADATA)),
            Self::Drawing => Some(Exact(CT_DRAWING)),
            Self::Chart => Some(Exact(CT_CHART)),
            Self::ChartEx => Some(Exact(CT_CHART_EX)),
            Self::ChartStyle => Some(Exact(CT_CHART_STYLE)),
            Self::ChartColorStyle => Some(Exact(CT_CHART_COLOR_STYLE)),
            Self::Image => Some(Prefix("image/")),
            Self::DiagramData => Some(Exact(CT_DIAGRAM_DATA)),
            Self::DiagramLayout => Some(Exact(CT_DIAGRAM_LAYOUT)),
            Self::DiagramQuickStyle => Some(Exact(CT_DIAGRAM_STYLE)),
            Self::DiagramColors => Some(Exact(CT_DIAGRAM_COLORS)),
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
