use std::io::Read;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::ChartEx;

pub(crate) fn parse_chart_ex<R: Read>(reader: R) -> XlsxResult<ChartEx> {
    duke_sheets_chart::parse::parse_chart_ex_xml(reader).map_err(|e| {
        XlsxError::Xml(quick_xml::Error::from(std::io::Error::new(
            std::io::ErrorKind::Other,
            e.to_string(),
        )))
    })
}
