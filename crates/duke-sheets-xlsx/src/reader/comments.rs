#[cfg(test)]
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::shared_strings::parse_rpr_element;
use super::{archive_by_name, decode_excel_escapes};
use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::comment::CellComment;
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::{CellAddress, DrawingText};

/// Parse a comments part into `(row, col, comment)` tuples in
/// document order. Placement (anchor, visibility, z-position) is
/// resolved by the caller against the legacy VML part.
pub(crate) fn read_comments_list<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    comments_path: &str,
) -> XlsxResult<Vec<(u32, u16, CellComment)>> {
    let mut comments = Vec::new();
    let file = match archive_by_name(archive, comments_path) {
        Ok(f) => f,
        Err(_) => return Ok(comments),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    // Keep whitespace: rich runs carry significant leading/trailing
    // spaces.
    xml_reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut authors: Vec<String> = Vec::new();

    let mut in_author = false;
    let mut author_text = String::new();
    let mut in_comment = false;
    let mut in_text = false;
    let mut in_t = false;
    let mut current_ref: Option<String> = None;
    let mut current_author_id: Option<usize> = None;
    let mut plain_text = String::new();

    // Rich-run state, mirroring the shared-strings CT_Rst parser.
    let mut in_r = false;
    let mut in_rpr = false;
    let mut in_run_t = false;
    let mut has_runs = false;
    let mut runs: Vec<RichTextRun> = Vec::new();
    let mut run_text = String::new();
    let mut run_font: Option<RunFont> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"author" => {
                    in_author = true;
                    author_text.clear();
                }
                b"comment" => {
                    in_comment = true;
                    current_ref = None;
                    current_author_id = None;
                    plain_text.clear();
                    runs.clear();
                    has_runs = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"ref" => {
                                current_ref = attr.unescape_value().ok().map(|s| s.to_string());
                            }
                            b"authorId" => {
                                current_author_id =
                                    attr.unescape_value().ok().and_then(|s| s.parse().ok());
                            }
                            _ => {}
                        }
                    }
                }
                b"text" if in_comment => in_text = true,
                b"r" if in_text => {
                    in_r = true;
                    has_runs = true;
                    run_text.clear();
                    run_font = None;
                }
                b"rPr" if in_r => {
                    in_rpr = true;
                    run_font = Some(RunFont::default());
                }
                b"t" if in_r => in_run_t = true,
                b"t" if in_text => in_t = true,
                name if in_rpr => parse_rpr_element(name, &e, &mut run_font),
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"author" => {
                    authors.push(std::mem::take(&mut author_text).trim().to_string());
                    in_author = false;
                }
                b"comment" => {
                    if let Some(ref cell_ref) = current_ref {
                        match CellAddress::parse(cell_ref) {
                            Ok(addr) => {
                                let author = current_author_id
                                    .and_then(|id| authors.get(id))
                                    .cloned()
                                    .unwrap_or_default();
                                let text = if has_runs {
                                    let mut runs = std::mem::take(&mut runs);
                                    for run in &mut runs {
                                        run.text = decode_excel_escapes(&run.text);
                                    }
                                    DrawingText {
                                        runs,
                                        ..DrawingText::default()
                                    }
                                } else {
                                    DrawingText::plain(decode_excel_escapes(&plain_text))
                                };
                                comments.push((addr.row, addr.col, CellComment { author, text }));
                            }
                            Err(e) => log::warn!("Skipping comment at '{}': {}", cell_ref, e),
                        }
                    }
                    in_comment = false;
                    plain_text.clear();
                }
                b"text" => in_text = false,
                b"r" if in_r => {
                    let font = run_font
                        .take()
                        .and_then(|font| if font.is_empty() { None } else { Some(font) });
                    runs.push(RichTextRun {
                        text: std::mem::take(&mut run_text),
                        font,
                    });
                    in_r = false;
                }
                b"rPr" => in_rpr = false,
                b"t" if in_run_t => in_run_t = false,
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_author {
                    if let Ok(text) = e.unescape() {
                        author_text.push_str(&text);
                    }
                } else if in_run_t {
                    if let Ok(text) = e.unescape() {
                        run_text.push_str(&text);
                    }
                } else if in_t {
                    if let Ok(text) = e.unescape() {
                        plain_text.push_str(&text);
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let local = e.name().local_name();
                let name = local.as_ref();
                if in_rpr {
                    parse_rpr_element(name, &e, &mut run_font);
                } else if name == b"author" {
                    authors.push(String::new());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

#[cfg(test)]
pub(crate) fn read_comment_visibility_map<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    vml_path: Option<&str>,
) -> XlsxResult<HashMap<(u32, u16), bool>> {
    let Some(vml_path) = vml_path else {
        return Ok(HashMap::new());
    };

    let file = match archive_by_name(archive, vml_path) {
        Ok(f) => f,
        Err(_) => return Ok(HashMap::new()),
    };

    let reader = BufReader::new(file);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut map: HashMap<(u32, u16), bool> = HashMap::new();

    let mut in_shape = false;
    let mut current_visible = false;
    let mut in_client_data_note = false;
    let mut in_row = false;
    let mut in_col = false;
    let mut current_row: Option<u32> = None;
    let mut current_col: Option<u16> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"shape" => {
                    in_shape = true;
                    current_visible = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"style" {
                            if let Some(style) =
                                attr.unescape_value().ok().map(|s| s.to_lowercase())
                            {
                                // Tolerate whitespace variations:
                                // "visibility:visible", "visibility: visible", etc.
                                let normalized: String =
                                    style.chars().filter(|c| !c.is_whitespace()).collect();
                                current_visible = normalized.contains("visibility:visible");
                            }
                        }
                    }
                }
                b"ClientData" if in_shape => {
                    let mut is_note = false;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"ObjectType" {
                            is_note = attr.unescape_value().ok().as_deref() == Some("Note");
                        }
                    }
                    if is_note {
                        in_client_data_note = true;
                        current_row = None;
                        current_col = None;
                    }
                }
                b"Row" if in_client_data_note => in_row = true,
                b"Column" if in_client_data_note => in_col = true,
                b"Visible" if in_client_data_note => {
                    // <x:Visible/> element explicitly marks the note as visible
                    current_visible = true;
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"shape" => {
                    in_shape = false;
                    in_client_data_note = false;
                    in_row = false;
                    in_col = false;
                    current_row = None;
                    current_col = None;
                }
                b"ClientData" if in_client_data_note => {
                    if let (Some(r), Some(c)) = (current_row, current_col) {
                        map.insert((r, c), current_visible);
                    }
                    in_client_data_note = false;
                    in_row = false;
                    in_col = false;
                    current_row = None;
                    current_col = None;
                }
                b"Row" => in_row = false,
                b"Column" => in_col = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_row {
                    if let Some(v) = e.unescape().ok().and_then(|s| s.parse::<u32>().ok()) {
                        current_row = Some(v);
                    }
                } else if in_col {
                    if let Some(v) = e.unescape().ok().and_then(|s| s.parse::<u16>().ok()) {
                        current_col = Some(v);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                // The VML visibility map is auxiliary metadata. Excel embeds
                // legacy HTML (e.g. unclosed `<br>`) inside `<v:textbox>`
                // content which is well-formed HTML but malformed XML.
                // Rather than failing the whole XLSX read over markup we
                // don't care about, log a warning and return whatever entries
                // we already extracted from the well-formed prefix.
                log::warn!(
                    "VML visibility map parse error at {vml_path}: {e} (returning partial map)"
                );
                break;
            }
            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"Visible" if in_client_data_note => {
                    // <x:Visible/> (self-closing) explicitly marks the note as visible
                    current_visible = true;
                }
                _ => {}
            },
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;
    use std::io::Write;

    use super::*;
    use crate::reader::XlsxReader;

    #[test]
    fn test_read_comment_visibility_map_from_vml() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:Row>1</x:Row>
      <x:Column>2</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;visibility:hidden">
    <x:ClientData ObjectType="Note">
      <x:Row>3</x:Row>
      <x:Column>4</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let map =
            read_comment_visibility_map(&mut archive, Some("xl/drawings/vmlDrawing1.vml")).unwrap();

        assert_eq!(map.get(&(1, 2)).copied(), Some(true));
        assert_eq!(map.get(&(3, 4)).copied(), Some(false));
    }

    #[test]
    fn test_read_comments_applies_visibility_from_vml() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData></worksheet>"#).unwrap();

            zip.start_file("xl/comments1.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>John</author></authors><commentList><comment ref="C2" authorId="0"><text><r><t>Visible note</t></r></text></comment></commentList></comments>"#).unwrap();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:Row>1</x:Row>
      <x:Column>2</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        let comment = sheet.comment("C2").unwrap().expect("comment should exist");
        assert_eq!(comment.author, "John");
        assert_eq!(comment.plain_text(), "Visible note");
        assert_eq!(sheet.comment_visible(1, 2), Some(true));
    }

    #[test]
    fn test_vml_visible_element_marks_comment_visible() {
        // Test that <x:Visible/> element (not just style attribute) marks visibility
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute">
    <x:ClientData ObjectType="Note">
      <x:Visible/>
      <x:Row>0</x:Row>
      <x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute">
    <x:ClientData ObjectType="Note">
      <x:Row>1</x:Row>
      <x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let map =
            read_comment_visibility_map(&mut archive, Some("xl/drawings/vmlDrawing1.vml")).unwrap();

        // Shape with <x:Visible/> should be visible
        assert_eq!(map.get(&(0, 0)).copied(), Some(true));
        // Shape without <x:Visible/> or style should default to hidden
        assert_eq!(map.get(&(1, 0)).copied(), Some(false));
    }

    #[test]
    fn test_vml_style_visibility_with_whitespace() {
        // Test that "visibility: visible" (with space) is tolerated
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute; visibility: visible">
    <x:ClientData ObjectType="Note">
      <x:Row>0</x:Row>
      <x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let map =
            read_comment_visibility_map(&mut archive, Some("xl/drawings/vmlDrawing1.vml")).unwrap();

        // Space between visibility: and visible should still be recognized
        assert_eq!(map.get(&(0, 0)).copied(), Some(true));
    }

    /// Excel sometimes embeds HTML-style markup with void `<br>` tags inside
    /// VML `<v:textbox>` content. That's well-formed HTML but malformed XML:
    /// quick-xml in strict mode rejects it with "expected `</br>`".
    /// Since VML visibility is auxiliary metadata for comments, a parse
    /// failure in unrelated markup must not propagate up and fail the whole
    /// XLSX read. Any visibility entries collected before the malformed
    /// content should still be returned.
    #[test]
    fn malformed_br_in_vml_textbox_yields_partial_visibility() {
        let mut bytes = Vec::new();
        {
            let cursor = Cursor::new(&mut bytes);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default();

            zip.start_file("xl/drawings/vmlDrawing1.vml", options)
                .unwrap();
            zip.write_all(
                br##"<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Note">
      <x:Row>5</x:Row>
      <x:Column>7</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t201" style="position:absolute">
    <v:textbox>
      <div><font>line one<br>
         </font><font>line two</font></div>
    </v:textbox>
    <x:ClientData ObjectType="Button"/>
  </v:shape>
</xml>"##,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let map = read_comment_visibility_map(&mut archive, Some("xl/drawings/vmlDrawing1.vml"))
            .expect("malformed VML must not fail the read");

        // The first shape's visibility was captured before the parser hit the
        // malformed <br>; it should still be in the map.
        assert_eq!(map.get(&(5, 7)).copied(), Some(true));
    }
}
