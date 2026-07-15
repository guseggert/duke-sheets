use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use duke_sheets_core::{CellAddress, Workbook};

use super::{write_xml_part, XlsxError, XlsxResult, NS_SPREADSHEET};

pub(super) fn write_vml_drawing<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    sheet_index: usize,
) -> XlsxResult<()> {
    let sheet = workbook
        .worksheet(sheet_index)
        .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

    duke_sheets_vml::validate_sheet_raw_client_data(sheet).map_err(XlsxError::InvalidFormat)?;
    let Some(xml) =
        duke_sheets_vml::build_legacy_vml(sheet, sheet_index, &workbook.theme_palette())
    else {
        return Ok(());
    };
    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options)?;
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

pub(super) fn write_comments<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    sheet_index: usize,
) -> XlsxResult<()> {
    let sheet = workbook
        .worksheet(sheet_index)
        .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

    if sheet.comment_count() == 0 {
        return Ok(());
    }

    let path = format!("xl/comments{}.xml", sheet_index + 1);
    write_xml_part(zip, &path, |w| {
        let mut tag = BytesStart::new("comments");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        w.write_event(Event::Start(tag))?;

        w.write_event(Event::Start(BytesStart::new("authors")))?;
        let authors = sheet.comment_authors();
        for author in &authors {
            w.create_element("author")
                .write_text_content(BytesText::new(author))?;
        }
        if authors.is_empty() {
            w.create_element("author")
                .write_text_content(BytesText::new(""))?;
        }
        w.write_event(Event::End(BytesEnd::new("authors")))?;

        w.write_event(Event::Start(BytesStart::new("commentList")))?;

        let mut comments: Vec<_> = sheet.comments().collect();
        comments.sort_by_key(|((row, col), _)| (*row, *col));

        let author_index: std::collections::HashMap<&str, usize> = authors
            .iter()
            .enumerate()
            .map(|(i, a)| (a.as_str(), i))
            .collect();

        for ((row, col), comment) in comments {
            let cell_ref = CellAddress::new(row, col).to_a1_string();
            let author_id = author_index
                .get(comment.author.as_str())
                .copied()
                .unwrap_or(0);
            let aid = author_id.to_string();

            let mut c_tag = BytesStart::new("comment");
            c_tag.push_attribute(("ref", cell_ref.as_str()));
            c_tag.push_attribute(("authorId", aid.as_str()));
            w.write_event(Event::Start(c_tag))?;

            w.write_event(Event::Start(BytesStart::new("text")))?;
            w.write_event(Event::Start(BytesStart::new("r")))?;
            w.create_element("t")
                .write_text_content(BytesText::new(&comment.text))?;
            w.write_event(Event::End(BytesEnd::new("r")))?;
            w.write_event(Event::End(BytesEnd::new("text")))?;

            w.write_event(Event::End(BytesEnd::new("comment")))?;
        }

        w.write_event(Event::End(BytesEnd::new("commentList")))?;
        w.write_event(Event::End(BytesEnd::new("comments")))?;
        Ok(())
    })
}
