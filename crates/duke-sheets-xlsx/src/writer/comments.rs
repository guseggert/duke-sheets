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

    if sheet.comment_count() == 0 && sheet.form_control_count() == 0 {
        return Ok(());
    }

    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options)?;

    let sheet_idx = sheet_index + 1;
    // (row, col, popup-visible); visibility lives on the wrapping
    // drawing object (`!hidden`).
    let mut comments: Vec<(u32, u16, bool)> = sheet
        .comments_drawn()
        .map(|cr| (cr.row, cr.col, !cr.object.meta.hidden))
        .collect();
    comments.sort_by_key(|(row, col, _)| (*row, *col));

    let mut xml = String::new();
    xml.push_str("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n");
    xml.push_str(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
    xml.push_str(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n");
    xml.push_str(" <o:shapelayout v:ext=\"edit\">\n");
    xml.push_str(&format!(
        "  <o:idmap v:ext=\"edit\" data=\"{}\"/>\n",
        sheet_idx
    ));
    xml.push_str(" </o:shapelayout>\n");
    if !comments.is_empty() {
        xml.push_str(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n");
        xml.push_str("  path=\"m,l,21600r21600,l21600,xe\">\n");
        xml.push_str("  <v:stroke joinstyle=\"miter\"/>\n");
        xml.push_str("  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n");
        xml.push_str(" </v:shapetype>\n");
    }
    if sheet.form_control_count() > 0 {
        xml.push_str(duke_sheets_vml::CONTROL_SHAPETYPE);
    }

    for (shape_index, &(row, col, visible)) in comments.iter().enumerate() {
        let row_above = row.saturating_sub(1);
        let shape_id = sheet_idx * 1024 + 1 + shape_index;
        let z_index = shape_index + 1;
        let left = (u32::from(col) + 1) * 64;
        let top = row * 15;
        let visibility = if visible { "visible" } else { "hidden" };

        xml.push_str(&format!(
            " <v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\"\n",
            shape_id
        ));
        xml.push_str(&format!(
            "  style='position:absolute;margin-left:{}pt;margin-top:{}pt;width:96pt;height:55.5pt;z-index:{};visibility:{}'\n",
            left, top, z_index, visibility
        ));
        xml.push_str("  fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n");
        xml.push_str("  <v:fill color2=\"#ffffe1\"/>\n");
        xml.push_str("  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n");
        xml.push_str("  <v:path o:connecttype=\"none\"/>\n");
        xml.push_str("  <v:textbox style='mso-direction-alt:auto'>\n");
        xml.push_str("   <div style='text-align:left'></div>\n");
        xml.push_str("  </v:textbox>\n");
        xml.push_str("  <x:ClientData ObjectType=\"Note\">\n");
        xml.push_str("   <x:MoveWithCells/>\n");
        xml.push_str("   <x:SizeWithCells/>\n");
        xml.push_str(&format!(
            "   <x:Anchor>{}, 15, {}, 10, {}, 15, {}, 4</x:Anchor>\n",
            col + 1,
            row_above,
            col + 3,
            row + 3
        ));
        xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
        xml.push_str(&format!("   <x:Row>{}</x:Row>\n", row));
        xml.push_str(&format!("   <x:Column>{}</x:Column>\n", col));
        xml.push_str("  </x:ClientData>\n");
        xml.push_str(" </v:shape>\n");
    }

    // Form control shapes follow the comment shapes; their shape ids
    // continue the same per-sheet 1024 block and must match the
    // worksheet <control shapeId> values.
    let controls: Vec<_> = sheet.form_controls().collect();
    if !controls.is_empty() {
        let heads = super::form_controls::radio_head_flags(sheet);
        let shape_base = sheet_idx * 1024 + 1 + comments.len();
        for (j, drawn) in controls.iter().enumerate() {
            duke_sheets_vml::write_control_shape(
                &mut xml,
                shape_base + j,
                comments.len() + j + 1,
                &drawn.object.meta,
                &drawn.object.anchor,
                drawn.payload,
                heads[j],
            );
        }
    }

    xml.push_str("</xml>");
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
            let author_id = if comment.author.is_empty() {
                0
            } else {
                author_index
                    .get(comment.author.as_str())
                    .copied()
                    .unwrap_or(0)
            };
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
