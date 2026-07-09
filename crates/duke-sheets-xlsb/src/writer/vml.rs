use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::XlsbResult;
use duke_sheets_core::Worksheet;

/// Write the sheet's legacy VML drawing part carrying comment
/// shapes and/or form control shapes. Returns whether a part was
/// written.
pub(crate) fn write_legacy_vml<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    sheet_index: usize,
    ws: &Worksheet,
) -> XlsbResult<bool> {
    let mut comments: Vec<_> = ws.comments().collect();
    let controls = ws.form_controls();
    if comments.is_empty() && controls.is_empty() {
        return Ok(false);
    }
    comments.sort_by_key(|((row, col), _)| (*row, *col));

    let path = format!("xl/drawings/vmlDrawing{}.vml", sheet_index + 1);
    zip.start_file(&path, *options)?;

    let sheet_idx = sheet_index + 1;

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
    if !controls.is_empty() {
        xml.push_str(duke_sheets_vml::CONTROL_SHAPETYPE);
    }

    for (shape_index, ((row, col), comment)) in comments.iter().enumerate() {
        let row = *row;
        let col = *col;
        let row_above = row.saturating_sub(1);
        let shape_id = sheet_idx * 1024 + 1 + shape_index;
        let z_index = shape_index + 1;
        let left = (u32::from(col) + 1) * 64;
        let top = row * 15;
        let visibility = if comment.visible { "visible" } else { "hidden" };

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

    // Form control shapes follow the comment shapes in the same
    // per-sheet shape id block.
    if !controls.is_empty() {
        let mut head_flags = vec![false; controls.len()];
        for group in duke_sheets_core::radio_groups(controls) {
            if let Some(&head) = group.first() {
                head_flags[head] = true;
            }
        }
        let shape_base = sheet_idx * 1024 + 1 + comments.len();
        for (j, control) in controls.iter().enumerate() {
            duke_sheets_vml::write_control_shape(
                &mut xml,
                shape_base + j,
                comments.len() + j + 1,
                control,
                head_flags[j],
            );
        }
    }

    xml.push_str("</xml>");
    zip.write_all(xml.as_bytes())?;
    Ok(true)
}
