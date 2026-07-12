use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::XlsbResult;
use duke_sheets_core::{DrawingKind, DrawingMeta, FormControl, Worksheet};

/// One control in the sheet's emission sequence: every form control
/// in the drawing tree, in placed (depth-first) order. This order
/// drives shape ids, VML shapes, and the drawing-part twins.
pub(crate) struct SheetControl<'a> {
    pub payload: &'a FormControl,
    pub meta: &'a DrawingMeta,
    /// Top-level controls keep their wrapper anchor; group children
    /// get an absolute anchor from their resolved on-sheet rectangle.
    pub anchor: duke_sheets_chart::DrawingAnchor,
}

/// The sheet's control sequence in placed (depth-first) order.
pub(crate) fn sheet_controls(sheet: &Worksheet) -> Vec<SheetControl<'_>> {
    sheet
        .placed_form_controls()
        .into_iter()
        .map(|placed| {
            let meta = sheet
                .drawing_at_path(&placed.path)
                .map(|node| node.meta)
                .expect("placed control path is valid");
            let anchor = if let [index] = placed.path.as_slice() {
                sheet.drawings()[*index].anchor.clone()
            } else {
                let (x1, y1, x2, y2) = placed.rect_emu;
                let clamp = |v: i128| v.clamp(0, i64::MAX as i128) as i64;
                duke_sheets_chart::DrawingAnchor::Absolute {
                    x_emu: clamp(x1),
                    y_emu: clamp(y1),
                    width_emu: clamp((x2 - x1).max(0)),
                    height_emu: clamp((y2 - y1).max(0)),
                }
            };
            SheetControl {
                payload: placed.control,
                meta,
                anchor,
            }
        })
        .collect()
}

/// Per-control radio-group-head flags, aligned with the placed
/// (depth-first) control order.
pub(crate) fn radio_head_flags(sheet: &Worksheet) -> Vec<bool> {
    let placed = sheet.placed_form_controls();
    let mut flags = vec![false; placed.len()];
    for group in duke_sheets_core::radio_groups(&placed) {
        if let Some(&head) = group.first() {
            flags[head] = true;
        }
    }
    flags
}

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
    let controls = sheet_controls(ws);
    if ws.comment_count() == 0 && controls.is_empty() {
        return Ok(false);
    }

    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    zip.start_file(&path, *options)?;

    let sheet_idx = sheet_index + 1;
    // Comment shape ids are assigned in (row, col) order; emission
    // order (and z-index) follows the drawing list.
    let mut comment_cells: Vec<(u32, u16)> =
        ws.comments_drawn().map(|cr| (cr.row, cr.col)).collect();
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
    // per-sheet 1024 block, in placed order, matching the
    // drawing-part twins.
    let heads = radio_head_flags(ws);
    let control_base = sheet_idx * 1024 + 1 + comment_count;
    let mut z_index = 0usize;
    let mut ordinal = 0usize;

    fn walk_controls(
        kind: &DrawingKind,
        xml: &mut String,
        controls: &[SheetControl<'_>],
        heads: &[bool],
        control_base: usize,
        z_index: &mut usize,
        ordinal: &mut usize,
    ) {
        match kind {
            DrawingKind::FormControl(_) => {
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
            DrawingKind::Group(group) => {
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

    for object in ws.drawings() {
        match &object.kind {
            DrawingKind::Comment { row, col, .. } => {
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
    Ok(true)
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
