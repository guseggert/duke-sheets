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

    let controls = super::form_controls::sheet_controls(sheet);
    if sheet.comment_count() == 0 && controls.is_empty() {
        return Ok(());
    }

    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options)?;

    let sheet_idx = sheet_index + 1;
    // Comment shape ids are assigned in (row, col) order; emission
    // order (and z-index) follows the drawing list.
    let mut comment_cells: Vec<(u32, u16)> = sheet
        .comments_drawn()
        .map(|cr| (cr.row, cr.col))
        .collect();
    comment_cells.sort();
    let comment_count = comment_cells.len();
    let comment_id = |row: u32, col: u16| -> usize {
        let index = comment_cells
            .iter()
            .position(|&(r, c)| (r, c) == (row, col))
            .unwrap_or(0);
        sheet_idx * 1024 + 1 + index
    };

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
    if comment_count > 0 {
        xml.push_str(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n");
        xml.push_str("  path=\"m,l,21600r21600,l21600,xe\">\n");
        xml.push_str("  <v:stroke joinstyle=\"miter\"/>\n");
        xml.push_str("  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n");
        xml.push_str(" </v:shapetype>\n");
    }
    if !controls.is_empty() {
        xml.push_str(duke_sheets_vml::CONTROL_SHAPETYPE);
    }

    // Shapes in drawing-list order (the Comment/FormControl
    // subsequence, descending into groups), z-index 1-based within
    // this sequence. Control shape ids follow the comments in the
    // per-sheet 1024 block, in placed order, matching the worksheet
    // <control shapeId> values and the drawing-part twins.
    let heads = super::form_controls::radio_head_flags(sheet);
    let control_base = sheet_idx * 1024 + 1 + comment_count;
    let mut z_index = 0usize;
    let mut ordinal = 0usize;

    fn walk_controls(
        kind: &duke_sheets_core::DrawingKind,
        xml: &mut String,
        controls: &[super::form_controls::SheetControl<'_>],
        heads: &[bool],
        control_base: usize,
        z_index: &mut usize,
        ordinal: &mut usize,
    ) {
        match kind {
            duke_sheets_core::DrawingKind::FormControl(_) => {
                let control = &controls[*ordinal];
                *z_index += 1;
                duke_sheets_vml::write_control_shape(
                    xml,
                    control_base + *ordinal,
                    *z_index,
                    control.meta,
                    &control.anchor,
                    control.payload,
                    heads[*ordinal],
                );
                *ordinal += 1;
            }
            duke_sheets_core::DrawingKind::Group(group) => {
                for child in &group.children {
                    walk_controls(
                        &child.kind,
                        xml,
                        controls,
                        heads,
                        control_base,
                        z_index,
                        ordinal,
                    );
                }
            }
            _ => {}
        }
    }

    for object in sheet.drawings() {
        match &object.kind {
            duke_sheets_core::DrawingKind::Comment { row, col, .. } => {
                z_index += 1;
                write_note_shape(
                    &mut xml,
                    comment_id(*row, *col),
                    z_index,
                    *row,
                    *col,
                    &object.anchor,
                    !object.meta.hidden,
                );
            }
            kind => walk_controls(
                kind,
                &mut xml,
                &controls,
                &heads,
                control_base,
                &mut z_index,
                &mut ordinal,
            ),
        }
    }

    xml.push_str("</xml>");
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

/// One comment Note shape. The `x:Anchor` (and the style box) derive
/// from the wrapper anchor instead of being re-synthesized from the
/// cell position.
fn write_note_shape(
    xml: &mut String,
    shape_id: usize,
    z_index: usize,
    row: u32,
    col: u16,
    anchor: &duke_sheets_chart::DrawingAnchor,
    visible: bool,
) {
    let a = duke_sheets_vml::anchor_to_px(anchor);
    let left = a[0] * duke_sheets_vml::DEFAULT_COL_PX + a[1];
    let top = a[2] * duke_sheets_vml::DEFAULT_ROW_PX + a[3];
    let width = (a[4] * duke_sheets_vml::DEFAULT_COL_PX + a[5]) - left;
    let height = (a[6] * duke_sheets_vml::DEFAULT_ROW_PX + a[7]) - top;
    let visibility = if visible { "visible" } else { "hidden" };

    xml.push_str(&format!(
        " <v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\"\n",
        shape_id
    ));
    xml.push_str(&format!(
        "  style='position:absolute;margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{};visibility:{}'\n",
        duke_sheets_vml::px_to_pt_string(left),
        duke_sheets_vml::px_to_pt_string(top),
        duke_sheets_vml::px_to_pt_string(width.max(0)),
        duke_sheets_vml::px_to_pt_string(height.max(0)),
        z_index,
        visibility
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
        "   <x:Anchor>{}, {}, {}, {}, {}, {}, {}, {}</x:Anchor>\n",
        a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]
    ));
    xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
    xml.push_str(&format!("   <x:Row>{}</x:Row>\n", row));
    xml.push_str(&format!("   <x:Column>{}</x:Column>\n", col));
    xml.push_str("  </x:ClientData>\n");
    xml.push_str(" </v:shape>\n");
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
