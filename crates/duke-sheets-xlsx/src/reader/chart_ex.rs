use std::io::{Read, Seek};

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_chart::{ChartEx, DrawingAnchor};

pub(crate) fn read_chart_ex<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    chart_path: &str,
    anchor: DrawingAnchor,
) -> XlsxResult<Option<ChartEx>> {
    let file = match archive.by_name(chart_path) {
        Ok(f) => f,
        Err(_) => return Ok(None),
    };

    let mut cx = duke_sheets_chart::parse::parse_chart_ex_xml(file).map_err(|e| {
        XlsxError::Xml(quick_xml::Error::from(std::io::Error::new(
            std::io::ErrorKind::Other,
            e.to_string(),
        )))
    })?;
    cx.anchor = anchor;
    Ok(Some(cx))
}
