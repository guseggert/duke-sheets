use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::{XlsxError, XlsxResult};

/// Parse a chartsheet XML part and return the drawing relationship id.
///
/// A chartsheet has the structure:
/// ```xml
/// <chartsheet ...>
///   <sheetViews>...</sheetViews>
///   <drawing r:id="rId1"/>
/// </chartsheet>
/// ```
pub(super) fn read_chartsheet_drawing_rid<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> XlsxResult<Option<String>> {
    let file = match archive.by_name(path) {
        Ok(f) => f,
        Err(_) => return Ok(None),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().local_name().as_ref() == b"drawing" =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == b"id" {
                        return Ok(attr.unescape_value().ok().map(|s| s.to_string()));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}
