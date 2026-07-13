use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::XlsbResult;
use duke_sheets_core::Worksheet;

pub(crate) use duke_sheets_vml::sheet_controls;

/// Write the sheet's legacy VML drawing part carrying comment Note
/// shapes and form control shapes, in drawing-list order (the shared
/// VML sequence carries their relative z-order). Returns whether a
/// part was written.
pub(crate) fn write_legacy_vml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    sheet_index: usize,
    ws: &Worksheet,
) -> XlsbResult<bool> {
    for control in &sheet_controls(ws) {
        control.payload.validate()?;
    }
    let Some(xml) = duke_sheets_vml::build_legacy_vml(ws, sheet_index) else {
        return Ok(false);
    };
    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    zip.start_file(&path, *options)?;
    zip.write_all(xml.as_bytes())?;
    Ok(true)
}
