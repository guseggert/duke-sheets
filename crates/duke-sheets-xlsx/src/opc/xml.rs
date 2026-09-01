use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;

use crate::error::{XlsxError, XlsxResult};

pub(super) fn namespace_is(resolution: &ResolveResult<'_>, namespace: &str) -> bool {
    matches!(resolution, ResolveResult::Bound(actual) if actual.as_ref() == namespace.as_bytes())
}

pub(super) fn validate_well_formed(bytes: &[u8]) -> XlsxResult<()> {
    let mut reader = quick_xml::Reader::from_reader(Cursor::new(bytes));
    let mut buf = Vec::new();
    let mut open_elements: Vec<Vec<u8>> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => open_elements.push(element.name().as_ref().to_vec()),
            Ok(Event::End(element)) => {
                let expected = open_elements.pop().ok_or_else(|| {
                    XlsxError::InvalidFormat("unexpected closing XML element".into())
                })?;
                if expected != element.name().as_ref() {
                    return Err(XlsxError::InvalidFormat(
                        "mismatched closing XML element".into(),
                    ));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(XlsxError::Xml(error)),
            _ => {}
        }
        buf.clear();
    }
    if !open_elements.is_empty() {
        return Err(XlsxError::InvalidFormat(
            "unclosed XML element at end of stream".into(),
        ));
    }
    Ok(())
}
