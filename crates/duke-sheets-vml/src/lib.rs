//! Legacy VML drawing codec for Forms controls.
//!
//! Both XLSX and XLSB carry form controls in a legacy VML drawing
//! part (`xl/drawings/vmlDrawingN.vml`): each control is a
//! `<v:shape>` of shapetype 201 whose `<x:ClientData>` holds the
//! control state and whose `<v:textbox>` holds the caption. XLSX
//! additionally duplicates most state into `ctrlProps` parts; XLSB
//! stores controls in the VML alone.
//!
//! The emitted shape markup mirrors Excel's own output per control
//! kind (styles, fills, lock elements, ClientData element order),
//! pinned against Excel-authored files. Element semantics follow
//! ECMA-376 Part 4 (VML spreadsheet drawing extensions), with one
//! empirically-pinned quirk: `x:MoveWithCells` / `x:SizeWithCells`
//! are *negations* — the element is present when the shape does NOT
//! move / size with cells (cross-referenced against the same file's
//! `controlPr@moveWithCells/sizeWithCells` attributes).

use duke_sheets_chart::{
    marker_position_emu, CellMarker, DrawingAnchor, DrawingMetrics, EditAs,
};
use duke_sheets_core::style::Underline;
use duke_sheets_core::{
    CheckState, Color, ControlText, DrawingMeta, DrawingObject, FormControl, FormControlKind,
    HorizontalAlignment, ListSelection, RichTextRun, RunFont, VerticalAlignment,
};
use quick_xml::events::Event;
use quick_xml::Reader;

/// EMU per screen pixel (96 dpi).
pub const EMU_PER_PX: i64 = 9525;
/// Default column width in pixels (Excel default metrics).
pub const DEFAULT_COL_PX: i64 = 64;
/// Default row height in pixels.
pub const DEFAULT_ROW_PX: i64 = 20;

/// The shapetype declaration required once per VML part containing
/// form controls (`o:spt="201"`, the host-control shape type).
pub const CONTROL_SHAPETYPE: &str = concat!(
    " <v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\"\n",
    "  path=\"m,l,21600r21600,l21600,xe\">\n",
    "  <v:stroke joinstyle=\"miter\"/>\n",
    "  <v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/>\n",
    "  <o:lock v:ext=\"edit\" shapetype=\"t\"/>\n",
    " </v:shapetype>\n",
);

/// Excel-style default shape name ("Check Box 3"). `seq` is the
/// 1-based per-sheet drawing object number.
pub fn default_control_name(kind: &FormControlKind, seq: usize) -> String {
    let base = match kind {
        FormControlKind::Button { .. } => "Button",
        FormControlKind::Checkbox { .. } => "Check Box",
        FormControlKind::OptionButton { .. } => "Option Button",
        FormControlKind::Label { .. } => "Label",
        FormControlKind::GroupBox { .. } => "Group Box",
        FormControlKind::ListBox { .. } => "List Box",
        FormControlKind::Dropdown { .. } => "Drop Down",
        FormControlKind::Scrollbar { .. } => "Scroll Bar",
        FormControlKind::Spinner { .. } => "Spinner",
        FormControlKind::Unknown { object_type, .. } => object_type.as_str(),
    };
    format!("{base} {seq}")
}

/// The VML `x:ClientData/@ObjectType` name for a control kind.
pub fn vml_object_type(kind: &FormControlKind) -> &str {
    match kind {
        FormControlKind::Button { .. } => "Button",
        FormControlKind::Checkbox { .. } => "Checkbox",
        FormControlKind::OptionButton { .. } => "Radio",
        FormControlKind::Label { .. } => "Label",
        FormControlKind::GroupBox { .. } => "GBox",
        FormControlKind::ListBox { .. } => "List",
        FormControlKind::Dropdown { .. } => "Drop",
        FormControlKind::Scrollbar { .. } => "Scroll",
        FormControlKind::Spinner { .. } => "Spin",
        FormControlKind::Unknown { object_type, .. } => object_type,
    }
}

/// Encode a model macro name for the legacy control formula carriers.
pub fn encode_macro_formula(macro_name: &str) -> String {
    if macro_name.starts_with('[') {
        macro_name.to_string()
    } else {
        format!("[0]!{macro_name}")
    }
}

/// Decode the current-workbook prefix used by legacy control formulas.
pub fn decode_macro_formula(formula: &str) -> String {
    formula
        .strip_prefix("[0]!")
        .unwrap_or(formula)
        .to_string()
}

/// The 8-value VML `x:Anchor` tuple
/// (`[colL, dxL_px, rowT, dyT_px, colR, dxR_px, rowB, dyB_px]`)
/// resolved with worksheet metrics.
pub fn anchor_to_px_with_metrics(
    anchor: &DrawingAnchor,
    metrics: &(impl DrawingMetrics + ?Sized),
) -> [i64; 8] {
    let (from, to) = anchor_cell_markers_with_metrics(anchor, metrics);
    [
        from.col as i64,
        from.col_offset_emu / EMU_PER_PX,
        from.row as i64,
        from.row_offset_emu / EMU_PER_PX,
        to.col as i64,
        to.col_offset_emu / EMU_PER_PX,
        to.row as i64,
        to.row_offset_emu / EMU_PER_PX,
    ]
}

/// Resolve any anchor variant to concrete from/to markers using worksheet
/// row heights and column widths.
pub fn anchor_cell_markers_with_metrics(
    anchor: &DrawingAnchor,
    metrics: &(impl DrawingMetrics + ?Sized),
) -> (CellMarker, CellMarker) {
    match anchor.to_two_cell_with_metrics(metrics) {
        DrawingAnchor::TwoCell { from, to, .. } => (from, to),
        _ => unreachable!("to_two_cell_with_metrics always returns TwoCell"),
    }
}

/// Rebuild a two-cell drawing anchor from an `x:Anchor` px tuple and
/// the (already un-negated) move/size flags. Move+size (Excel's
/// default) maps to `edit_as: None`, its canonical model form.
pub fn px_to_anchor(a: &[i64; 8], move_with_cells: bool, size_with_cells: bool) -> DrawingAnchor {
    let edit_as = match (move_with_cells, size_with_cells) {
        (true, true) => None,
        (true, false) => Some(EditAs::OneCell),
        _ => Some(EditAs::Absolute),
    };
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: a[0].clamp(0, u16::MAX as i64) as u16,
            col_offset_emu: a[1].saturating_mul(EMU_PER_PX),
            row: a[2].clamp(0, u32::MAX as i64) as u32,
            row_offset_emu: a[3].saturating_mul(EMU_PER_PX),
        },
        to: CellMarker {
            col: a[4].clamp(0, u16::MAX as i64) as u16,
            col_offset_emu: a[5].saturating_mul(EMU_PER_PX),
            row: a[6].clamp(0, u32::MAX as i64) as u32,
            row_offset_emu: a[7].saturating_mul(EMU_PER_PX),
        },
        edit_as,
    }
}

/// The negated `x:MoveWithCells` / `x:SizeWithCells` emission for an
/// anchor's editAs semantics: `(emit_move_with_cells_element,
/// emit_size_with_cells_element)`.
fn negated_move_size(anchor: &DrawingAnchor) -> (bool, bool) {
    let edit_as = match anchor {
        DrawingAnchor::TwoCell { edit_as, .. } => edit_as.clone().unwrap_or(EditAs::TwoCell),
        DrawingAnchor::OneCell { .. } => EditAs::OneCell,
        DrawingAnchor::Absolute { .. } => EditAs::Absolute,
    };
    match edit_as {
        EditAs::TwoCell => (false, false),
        EditAs::OneCell => (false, true),
        EditAs::Absolute => (true, true),
    }
}

/// A pixel length as a pt string for VML style attributes (96 dpi ->
/// 72 dpi, trailing zeros trimmed).
pub fn px_to_pt_string(px: i64) -> String {
    let pt = px as f64 * 0.75;
    let s = format!("{pt:.2}");
    let s = s.trim_end_matches('0').trim_end_matches('.');
    s.to_string()
}

fn fmt_pt(px: i64) -> String {
    px_to_pt_string(px)
}

fn xml_escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn xml_escape_attr(text: &str) -> String {
    xml_escape(text).replace('"', "&quot;")
}

fn horizontal_alignment_name(alignment: HorizontalAlignment) -> &'static str {
    match alignment {
        HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => "Center",
        HorizontalAlignment::Right => "Right",
        HorizontalAlignment::Justify => "Justify",
        HorizontalAlignment::Distributed => "Distributed",
        HorizontalAlignment::General | HorizontalAlignment::Left | HorizontalAlignment::Fill => {
            "Left"
        }
    }
}

fn vertical_alignment_name(alignment: VerticalAlignment) -> &'static str {
    match alignment {
        VerticalAlignment::Top => "Top",
        VerticalAlignment::Center => "Center",
        VerticalAlignment::Bottom => "Bottom",
        VerticalAlignment::Justify => "Justify",
        VerticalAlignment::Distributed => "Distributed",
    }
}

fn parse_horizontal_alignment(value: &str) -> Option<HorizontalAlignment> {
    match value.trim().to_ascii_lowercase().as_str() {
        "left" => Some(HorizontalAlignment::Left),
        "center" => Some(HorizontalAlignment::Center),
        "right" => Some(HorizontalAlignment::Right),
        "justify" => Some(HorizontalAlignment::Justify),
        "distributed" => Some(HorizontalAlignment::Distributed),
        _ => None,
    }
}

fn parse_vertical_alignment(value: &str) -> Option<VerticalAlignment> {
    match value.trim().to_ascii_lowercase().as_str() {
        "top" => Some(VerticalAlignment::Top),
        "center" => Some(VerticalAlignment::Center),
        "bottom" => Some(VerticalAlignment::Bottom),
        "justify" => Some(VerticalAlignment::Justify),
        "distributed" => Some(VerticalAlignment::Distributed),
        _ => None,
    }
}

fn normalize_vml_defaults(control: &mut VmlControl) {
    let (horizontal, vertical, size) = match control.object_type.as_str() {
        "Button" => (HorizontalAlignment::Center, VerticalAlignment::Center, 11.0),
        "Checkbox" | "Radio" => (HorizontalAlignment::Left, VerticalAlignment::Center, 8.0),
        _ => (HorizontalAlignment::Left, VerticalAlignment::Top, 8.0),
    };
    if control.text.horizontal_alignment == Some(horizontal) {
        control.text.horizontal_alignment = None;
    }
    if control.text.vertical_alignment == Some(vertical) {
        control.text.vertical_alignment = None;
    }
    for run in &mut control.text.runs {
        let is_default = run.font.as_ref().is_some_and(|font| {
            font.name.as_deref() == Some("Segoe UI")
                && font.size == Some(size)
                && font.color.is_none()
                && font.bold.is_none()
                && font.italic.is_none()
                && font.underline.is_none()
                && font.strikethrough.is_none()
                && font.vertical_align.is_none()
                && font.family.is_none()
                && font.charset.is_none()
                && font.scheme.is_none()
        });
        if is_default {
            run.font = None;
        }
    }
}

fn default_text_alignment(kind: &FormControlKind) -> (HorizontalAlignment, VerticalAlignment) {
    match kind {
        FormControlKind::Button { .. } => (HorizontalAlignment::Center, VerticalAlignment::Center),
        FormControlKind::Checkbox { .. } | FormControlKind::OptionButton { .. } => {
            (HorizontalAlignment::Left, VerticalAlignment::Center)
        }
        _ => (HorizontalAlignment::Left, VerticalAlignment::Top),
    }
}

fn write_vml_run(xml: &mut String, text: &str, font: Option<&RunFont>, default_size: u16) {
    let name = font
        .and_then(|font| font.name.as_deref())
        .unwrap_or("Segoe UI");
    let size = font
        .and_then(|font| font.size)
        .map(|size| (size * 20.0).round().clamp(1.0, u16::MAX as f64) as u16)
        .unwrap_or(default_size);
    let color = match font.and_then(|font| font.color) {
        Some(Color::Auto) | None => "auto".to_string(),
        Some(color) => {
            let (r, g, b) = color.to_rgb();
            format!("#{r:02X}{g:02X}{b:02X}")
        }
    };
    xml.push_str(&format!(
        "<font face=\"{}\" size=\"{size}\" color=\"{color}\">",
        xml_escape_attr(name)
    ));
    let bold = font.and_then(|font| font.bold).unwrap_or(false);
    let italic = font.and_then(|font| font.italic).unwrap_or(false);
    let underline = font
        .and_then(|font| font.underline)
        .is_some_and(|underline| underline != Underline::None);
    if bold {
        xml.push_str("<b>");
    }
    if italic {
        xml.push_str("<i>");
    }
    if underline {
        xml.push_str("<u>");
    }
    xml.push_str(&xml_escape(text));
    if underline {
        xml.push_str("</u>");
    }
    if italic {
        xml.push_str("</i>");
    }
    if bold {
        xml.push_str("</b>");
    }
    xml.push_str("</font>");
}

fn caption_lines(text: &ControlText) -> Vec<Vec<(String, Option<RunFont>)>> {
    let mut lines = vec![Vec::new()];
    for run in &text.runs {
        let parts: Vec<&str> = run.text.split('\n').collect();
        for (index, part) in parts.iter().enumerate() {
            let part = part.strip_suffix('\r').unwrap_or(part);
            lines
                .last_mut()
                .expect("caption has a line")
                .push((part.to_string(), run.font.clone()));
            if index + 1 < parts.len() {
                lines.push(Vec::new());
            }
        }
    }
    lines
}

/// Append one control `<v:shape>` to a VML part body using worksheet
/// row and column metrics. `shape_id` is the numeric part of
/// `_x0000_s{id}` (must match the worksheet `control/@shapeId` in
/// XLSX); `first_button` is the recomputed radio-group-head flag for
/// option buttons. `meta` and `anchor` come from the control's
/// wrapping drawing object.
#[allow(clippy::too_many_arguments)]
pub fn write_control_shape_with_metrics(
    xml: &mut String,
    shape_id: usize,
    z_index: usize,
    meta: &DrawingMeta,
    anchor: &DrawingAnchor,
    control: &FormControl,
    first_button: bool,
    metrics: &(impl DrawingMetrics + ?Sized),
) {
    use FormControlKind as K;
    let kind = &control.kind;
    let a = anchor_to_px_with_metrics(anchor, metrics);
    let (from, to) = anchor_cell_markers_with_metrics(anchor, metrics);
    let (left_emu, top_emu) = marker_position_emu(&from, metrics);
    let (right_emu, bottom_emu) = marker_position_emu(&to, metrics);
    let left = (left_emu / i128::from(EMU_PER_PX)) as i64;
    let top = (top_emu / i128::from(EMU_PER_PX)) as i64;
    let width = ((right_emu - left_emu).max(0) / i128::from(EMU_PER_PX)) as i64;
    let height = ((bottom_emu - top_emu).max(0) / i128::from(EMU_PER_PX)) as i64;

    let wrap_tight = matches!(
        kind,
        K::Button { .. }
            | K::Checkbox { .. }
            | K::OptionButton { .. }
            | K::GroupBox { .. }
            | K::Label { .. }
            | K::Unknown { .. }
    );

    xml.push_str(&format!(
        " <v:shape id=\"_x0000_s{shape_id}\" type=\"#_x0000_t201\""
    ));
    if let Some(alt_text) = &meta.alt_text {
        xml.push_str(&format!(" alt=\"{}\"", xml_escape_attr(alt_text)));
    }
    xml.push_str(" style='position:absolute;\n");
    // Excel writes visible control shapes with no visibility token;
    // only hidden ones carry it.
    xml.push_str(&format!(
        "  margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{z_index}{}{}'\n",
        fmt_pt(left),
        fmt_pt(top),
        fmt_pt(width.max(0)),
        fmt_pt(height.max(0)),
        if meta.hidden {
            ";visibility:hidden"
        } else {
            ""
        },
        if wrap_tight {
            ";\n  mso-wrap-style:tight"
        } else {
            ""
        },
    ));
    match kind {
        K::Button { .. } => {
            xml.push_str("  fillcolor=\"buttonFace [67]\" o:insetmode=\"auto\">\n");
            xml.push_str("  <v:fill color2=\"buttonFace [67]\" o:detectmouseclick=\"t\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Checkbox { .. } | K::OptionButton { .. } => {
            xml.push_str("  filled=\"f\" fillcolor=\"windowText [64]\" stroked=\"f\" strokecolor=\"window [65]\"\n");
            xml.push_str("  strokeweight=\"3e-5mm\" o:insetmode=\"auto\">\n");
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Dropdown { .. } => {
            xml.push_str(
                "  filled=\"f\" fillcolor=\"windowText [64]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
        K::GroupBox { .. } => {
            xml.push_str(
                "  fillcolor=\"window [65]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::Label { .. } | K::Unknown { .. } => {
            xml.push_str(
                "  filled=\"f\" fillcolor=\"windowText [64]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\"/>\n");
        }
        K::ListBox { .. } => {
            xml.push_str(
                "  fillcolor=\"window [65]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\">\n",
            );
            xml.push_str("  <v:fill color2=\"window [65]\"/>\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
        K::Scrollbar { .. } | K::Spinner { .. } => {
            xml.push_str("  o:insetmode=\"auto\">\n");
            xml.push_str("  <o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>\n");
        }
    }

    // Caption textbox for captioned kinds.
    if let Some(caption) = control.caption() {
        let (default_horizontal, _) = default_text_alignment(kind);
        let align =
            horizontal_alignment_name(caption.horizontal_alignment.unwrap_or(default_horizontal))
                .to_ascii_lowercase();
        let size = if matches!(kind, K::Button { .. }) {
            220
        } else {
            160
        };
        xml.push_str("  <v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">\n");
        for line in caption_lines(caption) {
            xml.push_str(&format!("   <div style='text-align:{align}'>"));
            if line.is_empty() {
                write_vml_run(xml, "", None, size);
            }
            for (text, font) in line {
                write_vml_run(xml, &text, font.as_ref(), size);
            }
            xml.push_str("</div>\n");
        }
        xml.push_str("  </v:textbox>\n");
    }

    // ClientData.
    xml.push_str(&format!(
        "  <x:ClientData ObjectType=\"{}\">\n",
        vml_object_type(kind)
    ));
    if !meta.locked {
        xml.push_str("   <x:Locked>False</x:Locked>\n");
    }
    let (no_move, no_size) = negated_move_size(anchor);
    if no_move {
        xml.push_str("   <x:MoveWithCells/>\n");
    }
    if no_size {
        xml.push_str("   <x:SizeWithCells/>\n");
    }
    xml.push_str(&format!(
        "   <x:Anchor>\n    {}, {}, {}, {}, {}, {}, {}, {}</x:Anchor>\n",
        a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]
    ));
    if !meta.printable {
        xml.push_str("   <x:PrintObject>False</x:PrintObject>\n");
    }
    if let Some(macro_name) = &control.macro_name {
        xml.push_str(&format!(
            "   <x:FmlaMacro>{}</x:FmlaMacro>\n",
            xml_escape(&encode_macro_formula(macro_name))
        ));
    }

    if let Some(caption) = control.caption() {
        let (default_horizontal, default_vertical) = default_text_alignment(kind);
        let horizontal = caption.horizontal_alignment.unwrap_or(default_horizontal);
        let vertical = caption.vertical_alignment.unwrap_or(default_vertical);
        xml.push_str(&format!(
            "   <x:TextHAlign>{}</x:TextHAlign>\n",
            horizontal_alignment_name(horizontal)
        ));
        xml.push_str(&format!(
            "   <x:TextVAlign>{}</x:TextVAlign>\n",
            vertical_alignment_name(vertical)
        ));
    }

    let push_checked = |xml: &mut String, state: &CheckState| {
        let v = match state {
            CheckState::Unchecked => 0,
            CheckState::Checked => 1,
            CheckState::Mixed => 2,
        };
        if v != 0 {
            xml.push_str(&format!("   <x:Checked>{v}</x:Checked>\n"));
        }
    };
    let push_link = |xml: &mut String, link: &Option<String>| {
        if let Some(link) = link {
            xml.push_str(&format!(
                "   <x:FmlaLink>{}</x:FmlaLink>\n",
                xml_escape(link)
            ));
        }
    };

    match kind {
        K::Button { .. } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
        }
        K::Checkbox {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            xml.push_str("   <x:AutoLine>False</x:AutoLine>\n");
            push_checked(xml, state);
            push_link(xml, cell_link);
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
        }
        K::OptionButton {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            xml.push_str("   <x:AutoLine>False</x:AutoLine>\n");
            push_checked(xml, state);
            push_link(xml, cell_link);
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
            if first_button {
                xml.push_str("   <x:FirstButton/>\n");
            }
        }
        K::Label { .. } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
        }
        K::GroupBox { no_3d, .. } => {
            if *no_3d {
                xml.push_str("   <x:NoThreeD/>\n");
            }
        }
        K::ListBox {
            input_range,
            cell_link,
            selection,
            selected,
            no_3d,
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            push_link(xml, cell_link);
            xml.push_str("   <x:Val>0</x:Val>\n   <x:Min>0</x:Min>\n   <x:Max>0</x:Max>\n");
            xml.push_str("   <x:Inc>1</x:Inc>\n   <x:Page>6</x:Page>\n   <x:Dx>22</x:Dx>\n");
            if let Some(range) = input_range {
                xml.push_str(&format!(
                    "   <x:FmlaRange>{}</x:FmlaRange>\n",
                    xml_escape(range)
                ));
            }
            if let Some(&first) = selected.first() {
                xml.push_str(&format!("   <x:Sel>{}</x:Sel>\n", u32::from(first) + 1));
            }
            if *no_3d {
                xml.push_str("   <x:NoThreeD2/>\n");
            }
            let sel_type = match selection {
                ListSelection::Single => "Single",
                ListSelection::Multi => "Multi",
                ListSelection::Extend => "Extended",
            };
            xml.push_str(&format!("   <x:SelType>{sel_type}</x:SelType>\n"));
            xml.push_str("   <x:LCT>Normal</x:LCT>\n");
            if !matches!(selection, ListSelection::Single) && !selected.is_empty() {
                let list = selected
                    .iter()
                    .map(|&v| (u32::from(v) + 1).to_string())
                    .collect::<Vec<_>>()
                    .join(",");
                xml.push_str(&format!("   <x:MultiSel>{list}</x:MultiSel>\n"));
            }
        }
        K::Dropdown {
            input_range,
            cell_link,
            selected,
            lines,
            no_3d,
        } => {
            xml.push_str("   <x:AutoFill>False</x:AutoFill>\n");
            push_link(xml, cell_link);
            xml.push_str("   <x:Val>0</x:Val>\n   <x:Min>0</x:Min>\n   <x:Max>0</x:Max>\n");
            xml.push_str("   <x:Inc>1</x:Inc>\n   <x:Page>10</x:Page>\n   <x:Dx>22</x:Dx>\n");
            if let Some(range) = input_range {
                xml.push_str(&format!(
                    "   <x:FmlaRange>{}</x:FmlaRange>\n",
                    xml_escape(range)
                ));
            }
            if let Some(sel) = selected {
                xml.push_str(&format!("   <x:Sel>{}</x:Sel>\n", u32::from(*sel) + 1));
            }
            if *no_3d {
                xml.push_str("   <x:NoThreeD2/>\n");
            }
            xml.push_str("   <x:SelType>Single</x:SelType>\n");
            xml.push_str("   <x:LCT>Normal</x:LCT>\n");
            xml.push_str("   <x:DropStyle>Combo</x:DropStyle>\n");
            xml.push_str(&format!("   <x:DropLines>{lines}</x:DropLines>\n"));
        }
        K::Scrollbar {
            value,
            min,
            max,
            increment,
            page,
            horizontal,
            cell_link,
        } => {
            push_link(xml, cell_link);
            xml.push_str(&format!("   <x:Val>{value}</x:Val>\n"));
            xml.push_str(&format!("   <x:Min>{min}</x:Min>\n"));
            xml.push_str(&format!("   <x:Max>{max}</x:Max>\n"));
            xml.push_str(&format!("   <x:Inc>{increment}</x:Inc>\n"));
            xml.push_str(&format!("   <x:Page>{page}</x:Page>\n"));
            if *horizontal {
                xml.push_str("   <x:Horiz/>\n");
            }
            xml.push_str("   <x:Dx>22</x:Dx>\n");
        }
        K::Spinner {
            value,
            min,
            max,
            increment,
            cell_link,
        } => {
            push_link(xml, cell_link);
            xml.push_str(&format!("   <x:Val>{value}</x:Val>\n"));
            xml.push_str(&format!("   <x:Min>{min}</x:Min>\n"));
            xml.push_str(&format!("   <x:Max>{max}</x:Max>\n"));
            xml.push_str(&format!("   <x:Inc>{increment}</x:Inc>\n"));
            xml.push_str("   <x:Page>10</x:Page>\n   <x:Dx>22</x:Dx>\n");
        }
        K::Unknown { .. } => {}
    }
    let object_type = vml_object_type(kind);
    for child in &control.raw_client_data {
        // A raw child whose name collides with a modeled emission for
        // this kind would double-represent it; the modeled value wins.
        if raw_child_local_name(child)
            .is_some_and(|name| is_modeled_client_data_child(object_type, &name))
        {
            continue;
        }
        xml.push_str("   ");
        xml.push_str(&String::from_utf8_lossy(child));
        xml.push('\n');
    }
    xml.push_str("  </x:ClientData>\n");
    xml.push_str(" </v:shape>\n");
}

/// One control in a sheet's emission sequence: every form control in
/// the drawing tree, in [`Worksheet::placed_form_controls`]
/// (depth-first) order. This order drives shape ids, VML shapes, the
/// drawing-part twins, and (in XLSX) ctrlProp part numbering and
/// `<controls>` entries.
///
/// [`Worksheet::placed_form_controls`]: duke_sheets_core::Worksheet::placed_form_controls
pub struct SheetControl<'a> {
    pub payload: &'a FormControl,
    pub meta: &'a DrawingMeta,
    /// Top-level controls keep their wrapper anchor; group children
    /// get an absolute anchor from their resolved on-sheet rectangle.
    pub anchor: DrawingAnchor,
}

/// The sheet's control sequence in placed (depth-first) order.
pub fn sheet_controls(sheet: &duke_sheets_core::Worksheet) -> Vec<SheetControl<'_>> {
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
                let rect = placed.rect_emu;
                DrawingAnchor::Absolute {
                    x_emu: rect.x_emu.max(0),
                    y_emu: rect.y_emu.max(0),
                    width_emu: rect.width_emu.max(0),
                    height_emu: rect.height_emu.max(0),
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
/// (depth-first) control order, derived from the spatial grouping in
/// [`duke_sheets_core::radio_groups`].
pub fn radio_head_flags(sheet: &duke_sheets_core::Worksheet) -> Vec<bool> {
    let placed = sheet.placed_form_controls();
    let mut flags = vec![false; placed.len()];
    for group in duke_sheets_core::radio_groups(&placed) {
        if let Some(&head) = group.first() {
            flags[head] = true;
        }
    }
    flags
}

/// Build the sheet's legacy VML drawing part body carrying comment
/// Note shapes and form control shapes, in drawing-list order (the
/// shared VML sequence carries their relative z-order). Comment shape
/// ids are assigned in (row, col) order and control shape ids follow
/// the comments in the per-sheet 1024 block, in placed order,
/// matching the drawing-part twins (and, in XLSX, the worksheet
/// `<control shapeId>` values). Returns `None` when the sheet has no
/// comments and no controls. `sheet_index` is zero-based.
pub fn build_legacy_vml(
    sheet: &duke_sheets_core::Worksheet,
    sheet_index: usize,
) -> Option<String> {
    use duke_sheets_core::DrawingKind;

    let controls = sheet_controls(sheet);
    if sheet.comment_count() == 0 && controls.is_empty() {
        return None;
    }

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
        xml.push_str(CONTROL_SHAPETYPE);
    }

    let heads = radio_head_flags(sheet);
    let control_base = sheet_idx * 1024 + 1 + comment_count;
    let mut z_index = 0usize;
    let mut ordinal = 0usize;

    fn walk_controls(
        kind: &duke_sheets_core::DrawingKind,
        xml: &mut String,
        controls: &[SheetControl<'_>],
        heads: &[bool],
        control_base: usize,
        z_index: &mut usize,
        ordinal: &mut usize,
        metrics: &duke_sheets_core::Worksheet,
    ) {
        match kind {
            duke_sheets_core::DrawingKind::FormControl(_) => {
                let control = &controls[*ordinal];
                *z_index += 1;
                write_control_shape_with_metrics(
                    xml,
                    control_base + *ordinal,
                    *z_index,
                    control.meta,
                    &control.anchor,
                    control.payload,
                    heads[*ordinal],
                    metrics,
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
                        metrics,
                    );
                }
            }
            _ => {}
        }
    }

    for object in sheet.drawings() {
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
                    sheet,
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
                sheet,
            ),
        }
    }

    xml.push_str("</xml>");
    Some(xml)
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
    anchor: &DrawingAnchor,
    visible: bool,
    metrics: &duke_sheets_core::Worksheet,
) {
    let a = anchor_to_px_with_metrics(anchor, metrics);
    let (from, to) = anchor_cell_markers_with_metrics(anchor, metrics);
    let (left_emu, top_emu) = marker_position_emu(&from, metrics);
    let (right_emu, bottom_emu) = marker_position_emu(&to, metrics);
    let left = (left_emu / i128::from(EMU_PER_PX)) as i64;
    let top = (top_emu / i128::from(EMU_PER_PX)) as i64;
    let width = ((right_emu - left_emu).max(0) / i128::from(EMU_PER_PX)) as i64;
    let height = ((bottom_emu - top_emu).max(0) / i128::from(EMU_PER_PX)) as i64;
    let visibility = if visible { "visible" } else { "hidden" };

    xml.push_str(&format!(
        " <v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\"\n",
        shape_id
    ));
    xml.push_str(&format!(
        "  style='position:absolute;margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{};visibility:{}'\n",
        px_to_pt_string(left),
        px_to_pt_string(top),
        px_to_pt_string(width.max(0)),
        px_to_pt_string(height.max(0)),
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
    // Excel keys the shown-state on the x:Visible element; the style
    // visibility alone is not enough.
    if visible {
        xml.push_str("   <x:Visible/>\n");
    }
    xml.push_str(&format!("   <x:Row>{}</x:Row>\n", row));
    xml.push_str(&format!("   <x:Column>{}</x:Column>\n", col));
    xml.push_str("  </x:ClientData>\n");
    xml.push_str(" </v:shape>\n");
}

/// Splice comments into an assembled native drawing list by the
/// legacy VML shape order: a comment goes immediately after the
/// nearest control shape preceding its Note in the VML sequence;
/// comments before any control go before the first control's object;
/// comments with no usable VML position (or on control-less sheets)
/// append at the end. `natives` pairs each object with its VML shape
/// number when it is (or wraps) a legacy control.
pub fn splice_comments(
    natives: Vec<(DrawingObject, Option<u32>)>,
    comments: Vec<(u32, u16, duke_sheets_core::comment::CellComment)>,
    vml_shapes: &[VmlShape],
) -> Vec<DrawingObject> {
    use std::collections::{HashMap, HashSet};

    let native_ctrl_ids: HashSet<u32> = natives.iter().filter_map(|(_, id)| *id).collect();

    // Comments part content, deduped by cell (first wins).
    let mut remaining: Vec<Option<(u32, u16, duke_sheets_core::comment::CellComment)>> = {
        let mut seen = HashSet::new();
        comments
            .into_iter()
            .filter(|(row, col, _)| seen.insert((*row, *col)))
            .map(Some)
            .collect()
    };
    let mut take_comment =
        |row: u32, col: u16| -> Option<(u32, u16, duke_sheets_core::comment::CellComment)> {
            remaining
                .iter_mut()
                .find(|slot| matches!(slot, Some((r, c, _)) if (*r, *c) == (row, col)))
                .and_then(Option::take)
        };

    let mut before_first: Vec<DrawingObject> = Vec::new();
    let mut after_ctrl: HashMap<u32, Vec<DrawingObject>> = HashMap::new();
    let mut last_ctrl: Option<u32> = None;
    for shape in vml_shapes {
        match &shape.kind {
            VmlShapeKind::Control(_) => {
                if native_ctrl_ids.contains(&shape.shape_num) {
                    last_ctrl = Some(shape.shape_num);
                }
            }
            VmlShapeKind::Note(note) => {
                let Some((row, col, comment)) = take_comment(note.row, note.col) else {
                    continue;
                };
                let mut object = DrawingObject::comment(row, col, comment);
                if let Some(a) = note.anchor_px {
                    let mut anchor = px_to_anchor(&a, true, true);
                    if let DrawingAnchor::TwoCell { edit_as, .. } = &mut anchor {
                        // Comments carry no editAs semantics.
                        *edit_as = None;
                    }
                    object.anchor = anchor;
                }
                object.meta.hidden = !note.visible;
                match last_ctrl {
                    Some(id) => after_ctrl.entry(id).or_default().push(object),
                    None => before_first.push(object),
                }
            }
        }
    }
    // Comments without a VML note keep default placement, at the end.
    let mut at_end: Vec<DrawingObject> = remaining
        .into_iter()
        .flatten()
        .map(|(row, col, comment)| DrawingObject::comment(row, col, comment))
        .collect();

    let first_ctrl_pos = natives.iter().position(|(_, id)| id.is_some());
    let mut result = Vec::new();
    for (i, (object, ctrl_id)) in natives.into_iter().enumerate() {
        if Some(i) == first_ctrl_pos {
            result.append(&mut before_first);
        }
        result.push(object);
        if let Some(id) = ctrl_id {
            if let Some(mut list) = after_ctrl.remove(&id) {
                result.append(&mut list);
            }
        }
    }
    // No control objects: "before first" degrades to append-at-end.
    result.append(&mut before_first);
    result.append(&mut at_end);
    result
}

/// One `<v:shape>`'s control-relevant contents, parsed from a VML
/// part.
#[derive(Debug, Clone, Default)]
pub struct VmlControl {
    /// Numeric shape id (the `N` of `_x0000_sN`).
    pub shape_num: u32,
    /// `x:ClientData/@ObjectType`.
    pub object_type: String,
    /// Caption text, rich runs, and alignment from the shape's textbox.
    pub text: ControlText,
    /// Macro assignment from `x:FmlaMacro`.
    pub macro_name: Option<String>,
    /// Alternative text from the VML shape's `alt` attribute.
    pub alt_text: Option<String>,
    /// `x:Anchor` values (col/px offsets).
    pub anchor_px: Option<[i64; 8]>,
    /// Un-negated flags (true = moves/sizes with cells).
    pub move_with_cells: bool,
    pub size_with_cells: bool,
    /// `x:Locked` (defaults true).
    pub locked: bool,
    /// `x:PrintObject` (defaults true).
    pub print_object: bool,
    /// Style attribute carries `visibility:hidden` (absent = shown).
    pub hidden: bool,
    /// `x:Checked` value (0/1/2).
    pub checked: u16,
    pub fmla_link: Option<String>,
    pub fmla_range: Option<String>,
    /// One-based `x:Sel` selection index (0 = no selection).
    pub sel: u16,
    /// `x:SelType` text (Single/Multi/Extended).
    pub sel_type: String,
    /// `x:MultiSel` one-based indices.
    pub multi_sel: Vec<u16>,
    /// `x:LCT` behavior class text (empty = unspecified).
    pub lct: String,
    pub drop_lines: u16,
    pub val: u16,
    pub min: u16,
    pub max: u16,
    pub inc: u16,
    pub page: u16,
    pub horiz: bool,
    /// `x:NoThreeD` or `x:NoThreeD2` present.
    pub no_3d: bool,
    /// `x:FirstButton` present.
    pub first_button: bool,
    /// `x:UIObj` marker (ECMA-376 Part 4 §14.4.2.62): the shape is
    /// Excel-managed UI chrome (autofilter/DV dropdown button), not a
    /// user object. OOXML twin of the BIFF ftCmo fUIObj bit.
    pub ui_obj: bool,
    /// Unmodeled immediate children of `x:ClientData`.
    pub raw_client_data: Vec<Vec<u8>>,
}

impl VmlControl {
    fn new() -> Self {
        VmlControl {
            locked: true,
            print_object: true,
            move_with_cells: true,
            size_with_cells: true,
            drop_lines: 8,
            inc: 1,
            page: 10,
            ..Default::default()
        }
    }

    /// Convert to a model [`DrawingObject`] wrapping a
    /// [`FormControl`]. Returns `None` for non-control shapes
    /// (comments) and unsupported object types.
    pub fn to_drawing_object(&self) -> Option<DrawingObject> {
        // Auxiliary UI shapes are regenerated by Excel from feature
        // state (autofilter, data validation); they are not user
        // content. Mirrors the XLS reader's fUIObj skip.
        if self.ui_obj {
            return None;
        }
        let caption = || self.text.clone();
        let state = match self.checked {
            2 => CheckState::Mixed,
            0 => CheckState::Unchecked,
            _ => CheckState::Checked,
        };
        let kind = match self.object_type.as_str() {
            "Button" => FormControlKind::Button { caption: caption() },
            "Checkbox" => FormControlKind::Checkbox {
                caption: caption(),
                state,
                cell_link: self.fmla_link.clone(),
                no_3d: self.no_3d,
            },
            "Radio" => FormControlKind::OptionButton {
                caption: caption(),
                state: if state == CheckState::Mixed {
                    CheckState::Checked
                } else {
                    state
                },
                cell_link: self.fmla_link.clone(),
                first_in_group: self.first_button,
                no_3d: self.no_3d,
            },
            "Label" => FormControlKind::Label { caption: caption() },
            "GBox" => FormControlKind::GroupBox {
                caption: caption(),
                no_3d: self.no_3d,
            },
            "List" => {
                let selection = match self.sel_type.as_str() {
                    "Multi" => ListSelection::Multi,
                    "Extended" => ListSelection::Extend,
                    _ => ListSelection::Single,
                };
                // File values are one-based (0 = none); the model is
                // zero-based.
                let selected = if selection == ListSelection::Single {
                    if self.sel > 0 {
                        vec![self.sel - 1]
                    } else {
                        Vec::new()
                    }
                } else if !self.multi_sel.is_empty() {
                    self.multi_sel
                        .iter()
                        .filter(|&&v| v > 0)
                        .map(|&v| v - 1)
                        .collect()
                } else if self.sel > 0 {
                    vec![self.sel - 1]
                } else {
                    Vec::new()
                };
                FormControlKind::ListBox {
                    input_range: self.fmla_range.clone(),
                    cell_link: self.fmla_link.clone(),
                    selection,
                    selected,
                    no_3d: self.no_3d,
                }
            }
            "Drop" => {
                // Auxiliary UI dropdowns (autofilter, data validation)
                // carry a non-Normal behavior class.
                if !self.lct.is_empty() && self.lct != "Normal" {
                    return None;
                }
                FormControlKind::Dropdown {
                    input_range: self.fmla_range.clone(),
                    cell_link: self.fmla_link.clone(),
                    selected: if self.sel > 0 {
                        Some(self.sel - 1)
                    } else {
                        None
                    },
                    lines: self.drop_lines,
                    no_3d: self.no_3d,
                }
            }
            "Scroll" => FormControlKind::Scrollbar {
                value: self.val,
                min: self.min,
                max: self.max,
                increment: self.inc,
                page: self.page,
                horizontal: self.horiz,
                cell_link: self.fmla_link.clone(),
            },
            "Spin" => FormControlKind::Spinner {
                value: self.val,
                min: self.min,
                max: self.max,
                increment: self.inc,
                cell_link: self.fmla_link.clone(),
            },
            object_type if object_type.eq_ignore_ascii_case("Note") => return None,
            object_type if duke_sheets_core::form_control::is_activex_object_type(object_type) => {
                return None
            }
            object_type => FormControlKind::Unknown {
                object_type: object_type.to_string(),
                legacy_object_type: None,
                caption: caption(),
                raw_properties: Vec::new(),
                raw_obj: None,
            },
        };

        let anchor = self
            .anchor_px
            .map(|a| px_to_anchor(&a, self.move_with_cells, self.size_with_cells))
            .unwrap_or_default();
        let mut control = FormControl::new(kind);
        control.macro_name = self.macro_name.clone();
        control.raw_client_data = self.raw_client_data.clone();
        let mut object = DrawingObject::form_control(control).with_anchor(anchor);
        object.meta.locked = self.locked;
        object.meta.printable = self.print_object;
        object.meta.hidden = self.hidden;
        object.meta.alt_text = self.alt_text.clone();
        Some(object)
    }
}

/// One `<v:shape>` from a VML part, in document order. Legacy Note
/// (comment) shapes and control shapes share this sequence, which
/// carries their relative z-order.
#[derive(Debug, Clone)]
pub struct VmlShape {
    /// Numeric shape id (the `N` of `_x0000_sN`).
    pub shape_num: u32,
    /// Kind-specific payload.
    pub kind: VmlShapeKind,
}

/// Payload of a [`VmlShape`].
#[derive(Debug, Clone)]
pub enum VmlShapeKind {
    /// A form-control (or other non-Note) shape.
    Control(VmlControl),
    /// A cell-comment Note shape.
    Note(VmlNote),
}

/// A comment Note shape's placement data.
#[derive(Debug, Clone)]
pub struct VmlNote {
    /// Anchored cell row (`x:Row`).
    pub row: u32,
    /// Anchored cell column (`x:Column`).
    pub col: u16,
    /// `x:Anchor` values (col/px offsets).
    pub anchor_px: Option<[i64; 8]>,
    /// Popup visibility (style `visibility:visible` or `x:Visible`).
    pub visible: bool,
}

/// `x:*` boolean element semantics (ST_TrueFalseBlank): present with
/// empty/`t`/`true`/`1` text means true; `f`/`false`/`0` means false.
fn blank_true(text: &str) -> bool {
    !matches!(
        text.trim().to_ascii_lowercase().as_str(),
        "f" | "false" | "0"
    )
}

/// Collapse line-wrap whitespace: runs containing a newline become a
/// single space mid-string and vanish at the ends (they are markup
/// indentation, not caption content). Plain space runs are
/// deliberate and preserved.
fn normalize_caption(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut run = String::new();
    let flush = |out: &mut String, run: &mut String| {
        if run.is_empty() {
            return;
        }
        if run.contains('\n') || run.contains('\r') {
            if !out.is_empty() {
                out.push(' ');
            }
        } else {
            out.push_str(run);
        }
        run.clear();
    };
    for ch in text.chars() {
        if ch.is_whitespace() {
            run.push(ch);
        } else {
            flush(&mut out, &mut run);
            out.push(ch);
        }
    }
    // Trailing run: drop when it contains a newline, keep deliberate
    // trailing spaces.
    if !run.is_empty() && !run.contains('\n') && !run.contains('\r') {
        out.push_str(&run);
    }
    out
}

/// Parse every control shape out of a VML drawing part.
///
/// Permissive: parse errors (Excel embeds unclosed HTML like `<br>`
/// inside textboxes) terminate the walk but shapes completed before
/// the error are returned.
pub fn parse_vml_controls(bytes: &[u8]) -> Vec<VmlControl> {
    parse_raw_shapes(bytes)
        .into_iter()
        .map(|shape| shape.control)
        .collect()
}

/// Parse the full ordered shape sequence out of a VML drawing part:
/// control shapes and comment Note shapes, in document order.
///
/// Same permissiveness as [`parse_vml_controls`]. Note shapes lacking
/// `x:Row`/`x:Column` are dropped (they cannot be joined to a cell).
pub fn parse_vml_shapes(bytes: &[u8]) -> Vec<VmlShape> {
    parse_raw_shapes(bytes)
        .into_iter()
        .filter_map(|shape| {
            let shape_num = shape.control.shape_num;
            let kind = if shape.control.object_type == "Note" {
                let (row, col) = (shape.row?, shape.col?);
                VmlShapeKind::Note(VmlNote {
                    row,
                    col,
                    anchor_px: shape.control.anchor_px,
                    visible: shape.visible,
                })
            } else {
                VmlShapeKind::Control(shape.control)
            };
            Some(VmlShape { shape_num, kind })
        })
        .collect()
}

/// A parsed `<v:shape>` before Note/control classification.
struct RawShape {
    control: VmlControl,
    row: Option<u32>,
    col: Option<u16>,
    visible: bool,
    /// Zero-based `<v:shape>` position in the part, pairing this shape
    /// with its raw ClientData captures. Shape ids are not usable as
    /// the key: third-party parts duplicate or free-form them.
    ordinal: usize,
}

fn parse_raw_shapes(bytes: &[u8]) -> Vec<RawShape> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;

    let mut out = Vec::new();
    let mut buf = Vec::new();

    let mut current: Option<RawShape> = None;
    let mut shape_ordinal = 0usize;
    let mut in_client_data = false;
    let mut in_textbox = false;
    let mut in_caption_div = false;
    let mut caption_line_runs: Vec<RichTextRun> = Vec::new();
    let mut caption_lines: Vec<Vec<RichTextRun>> = Vec::new();
    let mut current_font: Option<RunFont> = None;
    let mut current_font_text = String::new();
    let mut element_text: Option<(String, String)> = None; // (name, text)

    let flush_font =
        |runs: &mut Vec<RichTextRun>, font: &mut Option<RunFont>, text: &mut String| {
            if font.is_none() && text.is_empty() {
                return;
            }
            let normalized = normalize_caption(text);
            if !normalized.is_empty() || !text.is_empty() {
                runs.push(RichTextRun {
                    text: normalized,
                    font: font.take().filter(|font| !font.is_empty()),
                });
            } else {
                *font = None;
            }
            text.clear();
        };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"shape" => {
                        let mut shape = RawShape {
                            control: VmlControl::new(),
                            row: None,
                            col: None,
                            visible: false,
                            ordinal: shape_ordinal,
                        };
                        shape_ordinal += 1;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"id" => {
                                    let id = String::from_utf8_lossy(&attr.value);
                                    if let Some(num) = id.rsplit(['s', 'S']).next() {
                                        shape.control.shape_num = num.parse().unwrap_or(0);
                                    }
                                }
                                b"style" => {
                                    let style = String::from_utf8_lossy(&attr.value).to_lowercase();
                                    let normalized: String =
                                        style.chars().filter(|c| !c.is_whitespace()).collect();
                                    shape.visible = normalized.contains("visibility:visible");
                                    shape.control.hidden = normalized.contains("visibility:hidden");
                                }
                                b"alt" => {
                                    shape.control.alt_text = Some(
                                        attr.unescape_value()
                                            .map(|value| value.into_owned())
                                            .unwrap_or_else(|_| {
                                                String::from_utf8_lossy(&attr.value).into_owned()
                                            }),
                                    );
                                }
                                _ => {}
                            }
                        }
                        caption_line_runs.clear();
                        caption_lines.clear();
                        current_font = None;
                        current_font_text.clear();
                        in_client_data = false;
                        in_textbox = false;
                        in_caption_div = false;
                        current = Some(shape);
                    }
                    b"textbox" => in_textbox = true,
                    b"div" if in_textbox => {
                        in_caption_div = true;
                        caption_line_runs.clear();
                        current_font = None;
                        current_font_text.clear();
                        if let Some(shape) = current.as_mut() {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() != b"style" {
                                    continue;
                                }
                                let style = String::from_utf8_lossy(&attr.value);
                                for declaration in style.split(';') {
                                    let Some((name, value)) = declaration.split_once(':') else {
                                        continue;
                                    };
                                    if name.trim().eq_ignore_ascii_case("text-align") {
                                        if let Some(alignment) = parse_horizontal_alignment(value) {
                                            shape.control.text.horizontal_alignment =
                                                Some(alignment);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"font" if in_caption_div => {
                        flush_font(
                            &mut caption_line_runs,
                            &mut current_font,
                            &mut current_font_text,
                        );
                        let mut font = RunFont::default();
                        for attr in e.attributes().flatten() {
                            let value = attr
                                .unescape_value()
                                .map(|value| value.into_owned())
                                .unwrap_or_else(|_| {
                                    String::from_utf8_lossy(&attr.value).into_owned()
                                });
                            match attr.key.local_name().as_ref() {
                                b"face" => font.name = Some(value),
                                b"size" => {
                                    font.size = value.parse::<f64>().ok().map(|size| size / 20.0)
                                }
                                b"color" if !value.eq_ignore_ascii_case("auto") => {
                                    font.color = Color::from_hex(&value)
                                }
                                _ => {}
                            }
                        }
                        current_font = Some(font);
                    }
                    b"b" if in_caption_div => {
                        current_font.get_or_insert_with(RunFont::default).bold = Some(true);
                    }
                    b"i" if in_caption_div => {
                        current_font.get_or_insert_with(RunFont::default).italic = Some(true);
                    }
                    b"u" if in_caption_div => {
                        current_font.get_or_insert_with(RunFont::default).underline =
                            Some(Underline::Single);
                    }
                    b"ClientData" => {
                        in_client_data = true;
                        if let Some(shape) = current.as_mut() {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ObjectType" {
                                    shape.control.object_type =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                }
                            }
                        }
                    }
                    other if in_client_data => {
                        let elem = String::from_utf8_lossy(other).into_owned();
                        element_text = Some((elem, String::new()));
                    }
                    _ => {}
                }
                // Empty elements never receive an End event; commit
                // presence-only ClientData flags immediately.
                if in_client_data {
                    if let Some(shape) = current.as_mut() {
                        match name.as_slice() {
                            b"MoveWithCells" => shape.control.move_with_cells = false,
                            b"SizeWithCells" => shape.control.size_with_cells = false,
                            b"NoThreeD" | b"NoThreeD2" => shape.control.no_3d = true,
                            b"FirstButton" => shape.control.first_button = true,
                            b"UIObj" => shape.control.ui_obj = true,
                            b"Horiz" => shape.control.horiz = true,
                            b"Visible" => shape.visible = true,
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Text(t)) => {
                let text = t
                    .unescape()
                    .map(|c| c.into_owned())
                    .unwrap_or_else(|_| String::from_utf8_lossy(t.as_ref()).into_owned());
                if in_textbox && in_caption_div {
                    if current_font.is_some() {
                        current_font_text.push_str(&text);
                    } else {
                        let normalized = normalize_caption(&text);
                        if !normalized.is_empty() {
                            caption_line_runs.push(RichTextRun::plain(normalized));
                        }
                    }
                } else if let Some((_, buf_text)) = element_text.as_mut() {
                    buf_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name().as_ref().to_vec();
                match name.as_slice() {
                    b"shape" => {
                        if let Some(mut shape) = current.take() {
                            let mut runs = Vec::new();
                            let line_count = caption_lines.len();
                            for (index, mut line) in
                                std::mem::take(&mut caption_lines).into_iter().enumerate()
                            {
                                runs.append(&mut line);
                                if index + 1 < line_count {
                                    runs.push(RichTextRun::plain("\n"));
                                }
                            }
                            shape.control.text.runs = runs;
                            normalize_vml_defaults(&mut shape.control);
                            if !shape.control.object_type.is_empty() {
                                out.push(shape);
                            }
                        }
                    }
                    b"textbox" => in_textbox = false,
                    b"div" => {
                        if in_caption_div {
                            flush_font(
                                &mut caption_line_runs,
                                &mut current_font,
                                &mut current_font_text,
                            );
                            caption_lines.push(std::mem::take(&mut caption_line_runs));
                        }
                        in_caption_div = false;
                    }
                    b"font" if in_caption_div => flush_font(
                        &mut caption_line_runs,
                        &mut current_font,
                        &mut current_font_text,
                    ),
                    b"ClientData" => in_client_data = false,
                    _ => {
                        if let Some((elem, text)) = element_text.take() {
                            if elem.as_bytes() == name.as_slice() {
                                if let Some(shape) = current.as_mut() {
                                    match elem.as_str() {
                                        "Row" => shape.row = text.trim().parse().ok(),
                                        "Column" => shape.col = text.trim().parse().ok(),
                                        "Visible" => {
                                            shape.visible = blank_true(&text);
                                        }
                                        _ => {
                                            apply_client_data_text(&mut shape.control, &elem, &text)
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break, // permissive: keep completed shapes
            _ => {}
        }
        buf.clear();
    }
    let raw_client_data = parse_raw_client_data(bytes);
    for shape in &mut out {
        if let Some(children) = raw_client_data.get(&shape.ordinal) {
            shape.control.raw_client_data = children.clone();
        }
    }
    out
}

/// Validate every control's raw `ClientData` children on a sheet as
/// balanced, well-formed XML element fragments. Writers call this
/// before building the legacy VML part: one malformed fragment would
/// corrupt the whole part and take every comment and control with it.
pub fn validate_sheet_raw_client_data(
    sheet: &duke_sheets_core::Worksheet,
) -> Result<(), String> {
    for control in &sheet_controls(sheet) {
        for child in &control.payload.raw_client_data {
            validate_raw_client_data_fragment(child)?;
        }
    }
    Ok(())
}

/// A character permitted by XML 1.0 §2.2 `Char`.
fn valid_xml_char(c: char) -> bool {
    matches!(c, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&c)
        || ('\u{E000}'..='\u{FFFD}').contains(&c)
        || c >= '\u{10000}'
}

/// Check content bytes are UTF-8 made of XML-permitted characters.
/// quick-xml is a non-validating parser and never checks `Char`
/// validity; a control character or noncharacter in an emitted part
/// makes conforming consumers reject the whole part.
fn scan_xml_chars(bytes: &[u8]) -> Result<(), String> {
    let text =
        std::str::from_utf8(bytes).map_err(|_| "content is not valid UTF-8".to_string())?;
    match text.chars().find(|&c| !valid_xml_char(c)) {
        Some(c) => Err(format!(
            "character U+{:04X} is not permitted in XML",
            c as u32
        )),
        None => Ok(()),
    }
}

fn validate_raw_client_data_fragment(fragment: &[u8]) -> Result<(), String> {
    let display = || String::from_utf8_lossy(fragment).into_owned();
    let mut reader = Reader::from_reader(fragment);
    // quick-xml can enforce the comment `--` rules itself.
    reader.config_mut().check_comments = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(event) => event,
            Err(error) => {
                return Err(format!(
                    "raw ClientData fragment is not well-formed XML ({error}): {}",
                    display()
                ));
            }
        };
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                if e.local_name().as_ref().is_empty() {
                    return Err(format!(
                        "raw ClientData fragment has an element without a name: {}",
                        display()
                    ));
                }
                for attr in e.attributes() {
                    match attr {
                        Err(error) => {
                            return Err(format!(
                                "raw ClientData fragment has a malformed attribute ({error}): {}",
                                display()
                            ));
                        }
                        // Undefined entities, invalid character
                        // references, characters outside XML `Char`,
                        // or a literal `<` in a value make the part
                        // ill-formed; quick-xml checks none of these.
                        Ok(attr) => {
                            if attr.value.contains(&b'<') {
                                return Err(format!(
                                    "raw ClientData fragment has a literal '<' in an attribute value: {}",
                                    display()
                                ));
                            }
                            match attr.unescape_value() {
                                Err(error) => {
                                    return Err(format!(
                                        "raw ClientData fragment has an invalid attribute value ({error}): {}",
                                        display()
                                    ));
                                }
                                Ok(value) => {
                                    if let Some(c) =
                                        value.chars().find(|&c| !valid_xml_char(c))
                                    {
                                        return Err(format!(
                                            "raw ClientData fragment attribute contains U+{:04X}, not permitted in XML: {}",
                                            c as u32,
                                            display()
                                        ));
                                    }
                                }
                            }
                        }
                    }
                }
                if depth == 0 {
                    roots += 1;
                    // Exactly one root: the duplicate-name guard and
                    // the modeled/raw split inspect fragments as one
                    // element each.
                    if roots > 1 {
                        return Err(format!(
                            "raw ClientData fragment must contain exactly one element: {}",
                            display()
                        ));
                    }
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::End(_) => {
                if depth == 0 {
                    return Err(format!(
                        "raw ClientData fragment has an unopened end tag: {}",
                        display()
                    ));
                }
                depth -= 1;
            }
            Event::Text(text) => {
                // XML `S` is space/tab/CR/LF only (is_ascii_whitespace
                // would also admit form feed).
                if depth == 0
                    && !text
                        .iter()
                        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return Err(format!(
                        "raw ClientData fragment must contain exactly one element: {}",
                        display()
                    ));
                }
                // A literal ]]> is forbidden in character data
                // (XML 1.0 section 2.4).
                if text.windows(3).any(|window| window == b"]]>") {
                    return Err(format!(
                        "raw ClientData fragment contains a literal ]]> in text: {}",
                        display()
                    ));
                }
                // Undefined entities, invalid character references,
                // or resolved characters outside XML `Char` make the
                // part ill-formed.
                match text.unescape() {
                    Err(error) => {
                        return Err(format!(
                            "raw ClientData fragment has invalid text content ({error}): {}",
                            display()
                        ));
                    }
                    Ok(resolved) => {
                        if let Some(c) = resolved.chars().find(|&c| !valid_xml_char(c)) {
                            return Err(format!(
                                "raw ClientData fragment text contains U+{:04X}, not permitted in XML: {}",
                                c as u32,
                                display()
                            ));
                        }
                    }
                }
            }
            Event::Comment(content) => {
                if depth == 0 {
                    return Err(format!(
                        "raw ClientData fragment must start with its element, found a comment: {}",
                        display()
                    ));
                }
                if let Err(error) = scan_xml_chars(content) {
                    return Err(format!(
                        "raw ClientData fragment comment: {error}: {}",
                        display()
                    ));
                }
            }
            Event::CData(content) => {
                if depth == 0 {
                    return Err(format!(
                        "raw ClientData fragment must start with its element, found CDATA: {}",
                        display()
                    ));
                }
                if let Err(error) = scan_xml_chars(content) {
                    return Err(format!(
                        "raw ClientData fragment CDATA: {error}: {}",
                        display()
                    ));
                }
            }
            // An XML declaration or DOCTYPE is malformed anywhere
            // inside a part.
            Event::Decl(_) | Event::DocType(_) => {
                return Err(format!(
                    "raw ClientData fragment must not contain an XML declaration or DOCTYPE: {}",
                    display()
                ));
            }
            Event::PI(pi) => {
                if pi.target().eq_ignore_ascii_case(b"xml") {
                    return Err(format!(
                        "raw ClientData fragment must not contain an XML declaration or DOCTYPE: {}",
                        display()
                    ));
                }
                if depth == 0 {
                    return Err(format!(
                        "raw ClientData fragment must start with its element, found a processing instruction: {}",
                        display()
                    ));
                }
                if let Err(error) = scan_xml_chars(pi) {
                    return Err(format!(
                        "raw ClientData fragment processing instruction: {error}: {}",
                        display()
                    ));
                }
            }
            Event::Eof => break,
        }
        buf.clear();
    }
    if depth != 0 {
        return Err(format!(
            "raw ClientData fragment has an unclosed element: {}",
            display()
        ));
    }
    if roots == 0 {
        return Err(format!(
            "raw ClientData fragment must contain an element: {}",
            display()
        ));
    }
    Ok(())
}

/// Local element name of a raw `ClientData` child fragment: the tag
/// name of its first element, prefix stripped.
fn raw_child_local_name(fragment: &[u8]) -> Option<Vec<u8>> {
    let start = fragment.iter().position(|&b| b == b'<')?;
    let name: Vec<u8> = fragment[start + 1..]
        .iter()
        .take_while(|&&b| !b.is_ascii_whitespace() && b != b'>' && b != b'/')
        .copied()
        .collect();
    if name.is_empty() {
        return None;
    }
    let local = match name.iter().rposition(|&b| b == b':') {
        Some(colon) => name[colon + 1..].to_vec(),
        None => name,
    };
    (!local.is_empty()).then_some(local)
}

/// Whether a `ClientData` child is modeled for the given VML
/// ObjectType: parsed into a model field for that kind, or emitted by
/// the writer for that kind. Everything else is captured into
/// `FormControl::raw_client_data` on read and replayed verbatim on
/// write; the writer also drops raw children whose name is modeled so
/// the part never carries a double representation.
fn is_modeled_client_data_child(object_type: &str, name: &[u8]) -> bool {
    // Wrapper-level children: parsed for every kind and re-emitted
    // from the drawing object's meta/anchor/caption/macro.
    if matches!(
        name,
        b"Anchor"
            | b"Locked"
            | b"PrintObject"
            | b"MoveWithCells"
            | b"SizeWithCells"
            | b"FmlaMacro"
            | b"TextHAlign"
            | b"TextVAlign"
            | b"Row"
            | b"Column"
            | b"Visible"
            | b"UIObj"
    ) {
        return true;
    }
    match object_type {
        "Button" | "Label" => matches!(name, b"AutoFill"),
        "Checkbox" => matches!(
            name,
            b"AutoFill" | b"AutoLine" | b"Checked" | b"FmlaLink" | b"NoThreeD" | b"NoThreeD2"
        ),
        "Radio" => matches!(
            name,
            b"AutoFill"
                | b"AutoLine"
                | b"Checked"
                | b"FmlaLink"
                | b"NoThreeD"
                | b"NoThreeD2"
                | b"FirstButton"
        ),
        "GBox" => matches!(name, b"NoThreeD" | b"NoThreeD2"),
        "List" => matches!(
            name,
            b"AutoFill"
                | b"FmlaLink"
                | b"Val"
                | b"Min"
                | b"Max"
                | b"Inc"
                | b"Page"
                | b"Dx"
                | b"FmlaRange"
                | b"Sel"
                | b"NoThreeD"
                | b"NoThreeD2"
                | b"SelType"
                | b"LCT"
                | b"MultiSel"
        ),
        "Drop" => matches!(
            name,
            b"AutoFill"
                | b"FmlaLink"
                | b"Val"
                | b"Min"
                | b"Max"
                | b"Inc"
                | b"Page"
                | b"Dx"
                | b"FmlaRange"
                | b"Sel"
                | b"NoThreeD"
                | b"NoThreeD2"
                | b"SelType"
                | b"LCT"
                | b"DropStyle"
                | b"DropLines"
                // A Drop's selection is Sel-driven; a stray MultiSel
                // is a stale hint, normalized away like on List.
                | b"MultiSel"
        ),
        "Scroll" => matches!(
            name,
            b"FmlaLink" | b"Val" | b"Min" | b"Max" | b"Inc" | b"Page" | b"Horiz" | b"Dx"
        ),
        "Spin" => matches!(
            name,
            b"FmlaLink" | b"Val" | b"Min" | b"Max" | b"Inc" | b"Page" | b"Dx"
        ),
        // Unknown kinds keep every kind-specific child raw.
        _ => false,
    }
}

/// Capture complete immediate `ClientData` children that the modeled
/// parser does not understand, keyed by zero-based `<v:shape>`
/// position (ids are not unique in third-party parts). Byte slicing
/// keeps prefixes, attributes, nested content, and whitespace intact
/// for passthrough; fragments that fail validation (unclosed children
/// terminated by an ancestor's end tag) are dropped so a readable
/// file stays writable.
fn parse_raw_client_data(bytes: &[u8]) -> std::collections::HashMap<usize, Vec<Vec<u8>>> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut out: std::collections::HashMap<usize, Vec<Vec<u8>>> =
        std::collections::HashMap::new();
    let mut buf = Vec::new();
    let mut shape_ordinal: Option<usize> = None;
    let mut next_ordinal = 0usize;
    let mut in_client_data = false;
    let mut object_type = String::new();
    let mut capture: Option<(usize, u32)> = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let event = match reader.read_event_into(&mut buf) {
            Ok(event) => event,
            Err(_) => break,
        };
        let event_end = reader.buffer_position() as usize;

        if let Some((start, depth)) = capture.as_mut() {
            match &event {
                // A legitimate ClientData child never spans a shape
                // boundary or its ancestors' end tags. Seeing one
                // means the child was left unclosed: drop the capture
                // (a readable file must stay writable) and process
                // the boundary exactly as the main scan does, keeping
                // the shape ordinals of the two scans in lockstep.
                Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"shape" => {
                    capture = None;
                    shape_ordinal = Some(next_ordinal);
                    next_ordinal += 1;
                    in_client_data = false;
                    buf.clear();
                    continue;
                }
                Event::End(e)
                    if matches!(e.local_name().as_ref(), b"ClientData" | b"shape") =>
                {
                    capture = None;
                    in_client_data = false;
                    buf.clear();
                    continue;
                }
                Event::Start(_) => *depth += 1,
                Event::End(_) => *depth = depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            if *depth == 0 {
                let fragment = &bytes[*start..event_end];
                // Inner malformations (mismatched nested tags, lazy
                // attribute errors) must not enter the model either.
                if validate_raw_client_data_fragment(fragment).is_ok() {
                    if let Some(ordinal) = shape_ordinal {
                        out.entry(ordinal).or_default().push(fragment.to_vec());
                    }
                }
                capture = None;
            }
            buf.clear();
            continue;
        }

        match &event {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"shape" => {
                shape_ordinal = Some(next_ordinal);
                next_ordinal += 1;
                in_client_data = false;
            }
            // The shape's lifetime ends here even when its ClientData
            // was left unclosed; without this, presentation markup
            // between shapes would be captured as ClientData children.
            Event::End(e) if e.local_name().as_ref() == b"shape" => {
                shape_ordinal = None;
                in_client_data = false;
            }
            Event::Start(e) if e.local_name().as_ref() == b"ClientData" => {
                in_client_data = true;
                object_type.clear();
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"ObjectType" {
                        object_type = String::from_utf8_lossy(&attr.value).into_owned();
                    }
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"ClientData" => {
                in_client_data = false;
            }
            Event::Start(e)
                if in_client_data
                    && !is_modeled_client_data_child(&object_type, e.local_name().as_ref()) =>
            {
                capture = Some((event_start, 1));
            }
            Event::Empty(e)
                if in_client_data
                    && !is_modeled_client_data_child(&object_type, e.local_name().as_ref()) =>
            {
                // quick-xml parses attributes lazily, so an Empty
                // event can still carry malformed attributes; validate
                // like Start-rooted captures so it cannot poison a
                // later save.
                let fragment = &bytes[event_start..event_end];
                if validate_raw_client_data_fragment(fragment).is_ok() {
                    if let Some(ordinal) = shape_ordinal {
                        out.entry(ordinal).or_default().push(fragment.to_vec());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

/// Apply a text-bearing ClientData element to the control being
/// built.
fn apply_client_data_text(ctrl: &mut VmlControl, elem: &str, text: &str) {
    let text = text.trim();
    let num = || text.parse::<u16>().unwrap_or(0);
    match elem {
        "Anchor" => {
            let vals: Vec<i64> = text
                .split(',')
                .filter_map(|p| p.trim().parse().ok())
                .collect();
            if vals.len() == 8 {
                let mut a = [0i64; 8];
                a.copy_from_slice(&vals);
                ctrl.anchor_px = Some(a);
            }
        }
        "Locked" => ctrl.locked = blank_true(text),
        "PrintObject" => ctrl.print_object = blank_true(text),
        "MoveWithCells" => ctrl.move_with_cells = !blank_true(text),
        "SizeWithCells" => ctrl.size_with_cells = !blank_true(text),
        "Checked" => ctrl.checked = num(),
        "FmlaLink" => ctrl.fmla_link = Some(text.to_string()),
        "FmlaRange" => ctrl.fmla_range = Some(text.to_string()),
        "FmlaMacro" => ctrl.macro_name = Some(decode_macro_formula(text)),
        "TextHAlign" => {
            if let Some(alignment) = parse_horizontal_alignment(text) {
                ctrl.text.horizontal_alignment = Some(alignment);
            }
        }
        "TextVAlign" => {
            if let Some(alignment) = parse_vertical_alignment(text) {
                ctrl.text.vertical_alignment = Some(alignment);
            }
        }
        "Sel" => ctrl.sel = num(),
        "SelType" => ctrl.sel_type = text.to_string(),
        "LCT" => ctrl.lct = text.to_string(),
        "DropLines" => ctrl.drop_lines = num(),
        "Val" => ctrl.val = num(),
        "Min" => ctrl.min = num(),
        "Max" => ctrl.max = num(),
        "Inc" => ctrl.inc = num(),
        "Page" => ctrl.page = num(),
        "MultiSel" => {
            // Excel's emit order is not stable ("3, 1"); selection is
            // a set, so normalize to sorted indices.
            ctrl.multi_sel = text
                .split(|c: char| !c.is_ascii_digit())
                .filter(|s| !s.is_empty())
                .filter_map(|s| s.parse().ok())
                .collect();
            ctrl.multi_sel.sort_unstable();
            ctrl.multi_sel.dedup();
        }
        "Horiz" => ctrl.horiz = blank_true(text),
        "NoThreeD" | "NoThreeD2" => ctrl.no_3d = blank_true(text),
        "FirstButton" => ctrl.first_button = blank_true(text),
        "UIObj" => ctrl.ui_obj = blank_true(text),
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_chart::DefaultDrawingMetrics;
    use duke_sheets_core::DrawingKind;

    fn checkbox(cell_link: Option<&str>) -> DrawingObject {
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Enable audit".into(),
            state: CheckState::Checked,
            cell_link: cell_link.map(str::to_string),
            no_3d: true,
        }))
        .with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 0,
                col_offset_emu: 13 * EMU_PER_PX,
                row: 3,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 2,
                col_offset_emu: 18 * EMU_PER_PX,
                row: 4,
                row_offset_emu: 4 * EMU_PER_PX,
            },
            edit_as: Some(EditAs::OneCell),
        })
    }

    fn shape_xml(object: &DrawingObject, shape_id: usize, z_index: usize, first: bool) -> String {
        let mut xml = String::new();
        let control = object.kind.as_form_control().expect("form control");
        write_control_shape_with_metrics(
            &mut xml,
            shape_id,
            z_index,
            &object.meta,
            &object.anchor,
            control,
            first,
            &DefaultDrawingMetrics,
        );
        xml
    }

    fn wrap(xml_body: &str) -> String {
        format!(
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n{xml_body}</xml>"
        )
    }

    #[test]
    fn checkbox_shape_round_trips() {
        let object = checkbox(Some("$D$2"));
        let xml = shape_xml(&object, 1026, 2, false);
        let full = wrap(&xml);

        let parsed = parse_vml_controls(full.as_bytes());
        assert_eq!(parsed.len(), 1);
        let vml = &parsed[0];
        assert_eq!(vml.shape_num, 1026);
        assert_eq!(vml.object_type, "Checkbox");
        assert_eq!(vml.text.plain_text(), "Enable audit");
        assert_eq!(vml.anchor_px, Some([0, 13, 3, 0, 2, 18, 4, 4]));
        assert!(vml.move_with_cells);
        assert!(!vml.size_with_cells, "OneCell editAs negates sizing");
        assert_eq!(vml.checked, 1);
        assert_eq!(vml.fmla_link.as_deref(), Some("$D$2"));
        assert!(vml.no_3d);

        let back = vml.to_drawing_object().expect("control");
        assert_eq!(back.kind, object.kind);
        assert!(back.meta.locked);
        assert!(back.meta.printable);
    }

    #[test]
    fn excel_authored_checkbox_parses() {
        // Verbatim shape from an Excel-authored vmlDrawing part
        // (line-wrapped caption, quirky whitespace).
        let body = r##" <v:shape id="_x0000_s1026" type="#_x0000_t201" style='position:absolute;
  margin-left:9.75pt;margin-top:45pt;width:99.75pt;height:18pt;z-index:2;
  mso-wrap-style:tight' filled="f" fillcolor="windowText [64]" stroked="f"
  strokecolor="window [65]" strokeweight="3e-5mm" o:insetmode="auto">
  <v:fill color2="window [65]"/>
  <v:path shadowok="t" strokeok="t" fillok="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:left'><font face="Segoe UI" size="160" color="auto">Enable
   audit</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:SizeWithCells/>
   <x:Anchor>
    0, 13, 3, 0, 2, 18, 4, 4</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:TextVAlign>Center</x:TextVAlign>
   <x:Checked>1</x:Checked>
   <x:FmlaLink>$D$2</x:FmlaLink>
   <x:NoThreeD/>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        let vml = &parsed[0];
        assert_eq!(
            vml.text.plain_text(),
            "Enable audit",
            "wrapped caption normalizes"
        );
        assert_eq!(vml.checked, 1);
        assert!(!vml.size_with_cells);
        assert!(vml.move_with_cells);
        assert_eq!(vml.fmla_link.as_deref(), Some("$D$2"));
        let object = vml.to_drawing_object().expect("control");
        match &object.kind {
            DrawingKind::FormControl(control) => match &control.kind {
                FormControlKind::Checkbox { state, .. } => {
                    assert_eq!(state, &CheckState::Checked)
                }
                other => panic!("expected Checkbox, got {other:?}"),
            },
            other => panic!("expected form control, got {other:?}"),
        }
    }

    #[test]
    fn all_kinds_round_trip_through_vml() {
        let kinds: Vec<FormControlKind> = vec![
            FormControlKind::Button {
                caption: "Run".into(),
            },
            FormControlKind::OptionButton {
                caption: "Opt".into(),
                state: CheckState::Checked,
                cell_link: Some("$D$3".to_string()),
                first_in_group: true,
                no_3d: false,
            },
            FormControlKind::Label {
                caption: "Info <&>".into(),
            },
            FormControlKind::GroupBox {
                caption: "Frame".into(),
                no_3d: true,
            },
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$5".to_string()),
                cell_link: None,
                selection: ListSelection::Multi,
                selected: vec![0, 2, 4],
                no_3d: true,
            },
            FormControlKind::Dropdown {
                input_range: Some("$H$1:$H$4".to_string()),
                cell_link: Some("$D$4".to_string()),
                selected: Some(2),
                lines: 6,
                no_3d: true,
            },
            FormControlKind::Scrollbar {
                value: 40,
                min: 5,
                max: 95,
                increment: 2,
                page: 10,
                horizontal: true,
                cell_link: Some("$D$6".to_string()),
            },
            FormControlKind::Spinner {
                value: 12,
                min: 0,
                max: 30,
                increment: 3,
                cell_link: Some("$D$7".to_string()),
            },
        ];
        let mut xml = String::new();
        let objects: Vec<DrawingObject> = kinds
            .into_iter()
            .map(|kind| {
                DrawingObject::form_control(FormControl::new(kind)).with_anchor(
                    DrawingAnchor::TwoCell {
                        from: CellMarker {
                            col: 1,
                            col_offset_emu: 0,
                            row: 1,
                            row_offset_emu: 0,
                        },
                        to: CellMarker {
                            col: 3,
                            col_offset_emu: 0,
                            row: 3,
                            row_offset_emu: 0,
                        },
                        edit_as: None,
                    },
                )
            })
            .collect();
        for (i, object) in objects.iter().enumerate() {
            let control = object.kind.as_form_control().unwrap();
            let first = matches!(control.kind, FormControlKind::OptionButton { .. });
            xml.push_str(&shape_xml(object, 1025 + i, i + 1, first));
        }

        // Zero-based model selections serialize one-based on disk.
        assert!(xml.contains("<x:MultiSel>1,3,5</x:MultiSel>"), "{xml}");
        assert!(xml.contains("<x:Sel>3</x:Sel>"), "{xml}");

        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        assert_eq!(parsed.len(), objects.len());
        for (vml, original) in parsed.iter().zip(&objects) {
            let back = vml.to_drawing_object().expect("control");
            assert_eq!(back.kind, original.kind, "kind mismatch for {vml:?}");
        }
    }

    #[test]
    fn unlocked_unprintable_flags_round_trip() {
        let mut object = checkbox(None);
        object.meta.locked = false;
        object.meta.printable = false;
        let xml = shape_xml(&object, 1025, 1, false);
        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        let back = parsed[0].to_drawing_object().expect("control");
        assert!(!back.meta.locked);
        assert!(!back.meta.printable);
    }

    #[test]
    fn multiline_caption_round_trips_as_multiple_divs() {
        let mut object = checkbox(None);
        if let DrawingKind::FormControl(control) = &mut object.kind {
            if let FormControlKind::Checkbox { caption, .. } = &mut control.kind {
                *caption = "Line one\nLine two\n".into();
            }
        }
        let xml = shape_xml(&object, 1025, 1, false);
        assert_eq!(xml.matches("<div ").count(), 3);

        let parsed = parse_vml_controls(wrap(&xml).as_bytes());
        assert_eq!(parsed[0].text.plain_text(), "Line one\nLine two\n");
    }

    /// A shown comment's Note ClientData must carry `<x:Visible/>`:
    /// Excel keys the shown-state on the element (the style
    /// `visibility` alone is not honored). Hidden comments omit it.
    #[test]
    fn note_shape_emits_visible_element_only_when_shown() {
        use duke_sheets_core::comment::CellComment;

        let mut sheet = duke_sheets_core::Worksheet::new("Sheet1");
        sheet.add_drawing(DrawingObject::comment(
            0,
            0,
            CellComment::new("a", "hidden note"),
        ));
        sheet.add_drawing(
            DrawingObject::comment(5, 2, CellComment::new("a", "shown note")).with_hidden(false),
        );

        let xml = build_legacy_vml(&sheet, 0).expect("vml part");
        let shapes: Vec<&str> = xml.split("<v:shape ").skip(1).collect();
        assert_eq!(shapes.len(), 2);
        let hidden = shapes
            .iter()
            .find(|s| s.contains("<x:Row>0</x:Row>"))
            .expect("hidden note shape");
        let shown = shapes
            .iter()
            .find(|s| s.contains("<x:Row>5</x:Row>"))
            .expect("shown note shape");
        assert!(
            shown.contains("<x:Visible/>"),
            "shown note carries x:Visible: {shown}"
        );
        assert!(shown.contains("visibility:visible"));
        assert!(
            !hidden.contains("<x:Visible/>"),
            "hidden note omits x:Visible: {hidden}"
        );
        assert!(hidden.contains("visibility:hidden"));
    }

    #[test]
    fn note_shapes_are_not_controls() {
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t202">
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:Row>0</x:Row>
   <x:Column>0</x:Column>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert!(parsed[0].to_drawing_object().is_none());
    }

    #[test]
    fn malformed_textbox_html_keeps_completed_shapes() {
        // Unclosed <br> in a later shape must not lose earlier ones.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <v:textbox><div>broken<br></div></v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert!(!parsed.is_empty(), "first shape must survive");
        assert_eq!(parsed[0].shape_num, 1025);
        assert_eq!(parsed[0].checked, 1);
    }

    #[test]
    fn parse_vml_shapes_keeps_note_and_control_order() {
        let body = r##" <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;visibility:visible'>
  <x:ClientData ObjectType="Note">
   <x:Anchor>3, 15, 1, 10, 5, 15, 5, 4</x:Anchor>
   <x:Row>2</x:Row>
   <x:Column>2</x:Column>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1027" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 4, 0, 1, 0, 5, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
"##;
        let shapes = parse_vml_shapes(wrap(body).as_bytes());
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].shape_num, 1026);
        assert!(matches!(shapes[0].kind, VmlShapeKind::Control(_)));
        match &shapes[1].kind {
            VmlShapeKind::Note(note) => {
                assert_eq!((note.row, note.col), (2, 2));
                assert_eq!(note.anchor_px, Some([3, 15, 1, 10, 5, 15, 5, 4]));
                assert!(note.visible);
            }
            other => panic!("expected note, got {other:?}"),
        }
        assert_eq!(shapes[2].shape_num, 1027);

        // The control-only view stays unchanged (Notes included, as
        // before, filtered downstream by to_drawing_object).
        assert_eq!(parse_vml_controls(wrap(body).as_bytes()).len(), 3);
    }

    #[test]
    fn autofilter_lct_dropdowns_are_skipped() {
        let mut vml = VmlControl::new();
        vml.object_type = "Drop".to_string();
        vml.lct = "AutoFilter".to_string();
        assert!(vml.to_drawing_object().is_none());
    }

    #[test]
    fn uiobj_shapes_are_skipped() {
        // Old-Excel autofilter dropdown: x:UIObj + PrintObject=False,
        // no x:LCT. Must not surface as a user Dropdown control.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Drop">
   <x:SizeWithCells/>
   <x:Anchor>0, 0, 0, 0, 1, 16, 1, 0</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:DropStyle>Combo</x:DropStyle>
   <x:DropLines>8</x:DropLines>
   <x:Sel>0</x:Sel>
   <x:NoThreeD/>
   <x:PrintObject>False</x:PrintObject>
   <x:UIObj/>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert!(
            parsed[0].to_drawing_object().is_none(),
            "UIObj auxiliary shape must not become a control"
        );
    }

    #[test]
    fn uiobj_skip_is_object_type_independent() {
        // The marker, not the object type, identifies auxiliary UI.
        let vml_body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:UIObj>True</x:UIObj>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(vml_body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert!(parsed[0].to_drawing_object().is_none());
    }

    #[test]
    fn modeled_kinds_capture_unmodeled_client_data() {
        // ECMA-376 Part 4 ClientData children we do not model
        // (x:Disabled §14.4.2.19, x:Accel §14.4.2.1) must survive on
        // modeled control kinds, while modeled elements stay out of
        // the raw set (no double representation).
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:Checked>1</x:Checked>
   <x:FmlaLink>$A$1</x:FmlaLink>
   <x:Disabled/>
   <x:Accel>65</x:Accel>
   <x:NoThreeD/>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        let object = parsed[0].to_drawing_object().expect("control");
        let control = object.kind.as_form_control().expect("form control");
        assert!(matches!(
            control.kind,
            FormControlKind::Checkbox { .. }
        ));
        let raws: Vec<String> = control
            .raw_client_data
            .iter()
            .map(|raw| String::from_utf8_lossy(raw).into_owned())
            .collect();
        assert_eq!(
            raws,
            vec!["<x:Disabled/>".to_string(), "<x:Accel>65</x:Accel>".to_string()],
            "exactly the unmodeled children, in document order"
        );
    }

    #[test]
    fn raw_client_data_stays_with_its_shape_on_duplicate_ids() {
        // Third-party generators reuse or free-form the shape id; raw
        // capture must attach by document position, not parsed id.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Disabled/>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
   <x:Accel>65</x:Accel>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].raw_client_data, vec![b"<x:Disabled/>".to_vec()]);
        assert_eq!(
            parsed[1].raw_client_data,
            vec![b"<x:Accel>65</x:Accel>".to_vec()]
        );
    }

    #[test]
    fn unclosed_client_data_child_is_dropped_without_contamination() {
        // An unmodeled child left unclosed makes </x:ClientData> the
        // capture terminator. The malformed fragment must be dropped
        // (never fail a later save) and must not leak the ClientData
        // state into the next shape's presentation elements.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Broken>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <v:fill color="red"/>
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
   <x:Accel>65</x:Accel>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 2);
        assert_eq!(
            parsed[0].raw_client_data,
            Vec::<Vec<u8>>::new(),
            "malformed fragment must not enter the model"
        );
        assert_eq!(
            parsed[1].raw_client_data,
            vec![b"<x:Accel>65</x:Accel>".to_vec()],
            "next shape must not inherit ClientData state"
        );
    }

    #[test]
    fn empty_child_with_malformed_attribute_is_dropped_at_capture() {
        // quick-xml parses attributes lazily, so <x:Weird foo=bar/>
        // arrives as a valid Empty event. Capturing it unvalidated
        // would make a readable file unsaveable at write validation.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Weird foo=bar/>
   <x:Accel>65</x:Accel>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert_eq!(
            parsed[0].raw_client_data,
            vec![b"<x:Accel>65</x:Accel>".to_vec()],
            "malformed Empty child must be dropped like malformed Start children"
        );
    }

    #[test]
    fn shape_boundary_inside_capture_keeps_ordinals_in_sync() {
        // A captured child containing a <v:shape/> is well-formed XML,
        // but the main scan treats every shape start as a new shape.
        // The raw scan must count it too, or every later capture
        // attaches to the wrong control.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Weird><v:shape/></x:Weird>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
   <x:AAA>1</x:AAA>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1027" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 4, 0, 1, 0, 5, 0</x:Anchor>
   <x:BBB>2</x:BBB>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        let raws_of = |num: u32| -> Vec<Vec<u8>> {
            parsed
                .iter()
                .find(|control| control.shape_num == num)
                .expect("shape present")
                .raw_client_data
                .clone()
        };
        assert_eq!(raws_of(1026), vec![b"<x:AAA>1</x:AAA>".to_vec()]);
        assert_eq!(raws_of(1027), vec![b"<x:BBB>2</x:BBB>".to_vec()]);
    }

    #[test]
    fn deeply_unclosed_child_does_not_swallow_later_shapes() {
        // Several unclosed children mean the capture's depth counter
        // never reaches zero; without a boundary abort it would run to
        // </xml> and swallow every later shape's raws.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:A><x:B><x:C>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
   <x:AAA>1</x:AAA>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 2);
        assert_eq!(
            parsed[0].raw_client_data,
            Vec::<Vec<u8>>::new(),
            "unclosed children must be dropped"
        );
        assert_eq!(
            parsed[1].raw_client_data,
            vec![b"<x:AAA>1</x:AAA>".to_vec()],
            "the next shape's raws must survive"
        );
    }

    #[test]
    fn unclosed_client_data_ended_by_shape_does_not_capture_between_shapes() {
        // When </x:ClientData> is missing and the shape is closed by
        // </v:shape>, the raw scan must end the ClientData scope there
        // too; otherwise presentation markup between shapes would be
        // captured and re-emitted inside x:ClientData.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
 </v:shape>
 <v:fill color="red"/>
 <v:oval><x:Foo>1</x:Foo></v:oval>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
   <x:Accel>65</x:Accel>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 2);
        assert_eq!(
            parsed[0].raw_client_data,
            Vec::<Vec<u8>>::new(),
            "content between shapes must not become ClientData"
        );
        assert_eq!(
            parsed[1].raw_client_data,
            vec![b"<x:Accel>65</x:Accel>".to_vec()]
        );
    }

    #[test]
    fn floating_client_data_between_shapes_is_not_captured() {
        // A stray x:ClientData outside any shape applies to nothing in
        // the main scan; the raw scan must not attach its children to
        // the previous shape.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
 <x:ClientData ObjectType="Checkbox"><x:Foo>1</x:Foo></x:ClientData>
 <v:shape id="_x0000_s1026" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 2, 0, 1, 0, 3, 0</x:Anchor>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].raw_client_data, Vec::<Vec<u8>>::new());
        assert_eq!(parsed[1].raw_client_data, Vec::<Vec<u8>>::new());
    }

    #[test]
    fn multisel_on_dropdown_is_normalized_away() {
        // A Drop's selection is Sel-driven; a stray x:MultiSel is a
        // stale hint that must not be replayed next to the modeled
        // selection.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Drop">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Sel>2</x:Sel>
   <x:MultiSel>1,3</x:MultiSel>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].raw_client_data, Vec::<Vec<u8>>::new());
    }

    #[test]
    fn raw_client_data_write_skips_children_colliding_with_modeled_emit() {
        let mut object = checkbox(Some("$B$2"));
        {
            let control = match &mut object.kind {
                DrawingKind::FormControl(control) => control,
                other => panic!("expected form control, got {other:?}"),
            };
            control.raw_client_data = vec![
                b"<x:Disabled/>".to_vec(),
                b"<x:Accel>65</x:Accel>".to_vec(),
                // Collides with the checkbox's modeled emission; must
                // be dropped so the part carries a single x:Checked.
                b"<x:Checked>0</x:Checked>".to_vec(),
            ];
        }
        let xml = shape_xml(&object, 1025, 1, true);
        assert_eq!(xml.matches("<x:Disabled/>").count(), 1);
        assert_eq!(xml.matches("<x:Accel>65</x:Accel>").count(), 1);
        assert_eq!(
            xml.matches("<x:Checked>").count(),
            1,
            "raw duplicate of a modeled element must be dropped: {xml}"
        );
        assert!(xml.contains("<x:Checked>1</x:Checked>"));
    }

    #[test]
    fn explicit_uiobj_false_keeps_the_control() {
        // ST_TrueFalseBlank: "False" text negates the marker.
        let body = r##" <v:shape id="_x0000_s1025" type="#_x0000_t201">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:Checked>1</x:Checked>
   <x:UIObj>False</x:UIObj>
  </x:ClientData>
 </v:shape>
"##;
        let parsed = parse_vml_controls(wrap(body).as_bytes());
        assert_eq!(parsed.len(), 1);
        let object = parsed[0].to_drawing_object().expect("control");
        assert!(matches!(
            object.kind.as_form_control().unwrap().kind,
            FormControlKind::Checkbox { .. }
        ));
    }

    #[test]
    fn hostile_anchor_values_saturate_without_panicking() {
        let anchor = px_to_anchor(
            &[
                i64::MAX,
                i64::MAX,
                i64::MAX,
                i64::MIN,
                i64::MAX,
                i64::MIN,
                i64::MAX,
                i64::MAX,
            ],
            true,
            false,
        );
        match anchor {
            DrawingAnchor::TwoCell { from, to, .. } => {
                assert_eq!(from.col, u16::MAX);
                assert_eq!(from.row, u32::MAX);
                assert_eq!(from.col_offset_emu, i64::MAX);
                assert_eq!(from.row_offset_emu, i64::MIN);
                assert_eq!(to.col_offset_emu, i64::MIN);
                assert_eq!(to.row_offset_emu, i64::MAX);
            }
            other => panic!("expected TwoCell anchor, got {other:?}"),
        }
    }

    #[test]
    fn anchor_cell_markers_clamp_extreme_absolute_anchor() {
        let anchor = DrawingAnchor::Absolute {
            x_emu: i64::MAX,
            y_emu: i64::MAX,
            width_emu: i64::MAX,
            height_emu: i64::MAX,
        };
        let (from, to) = anchor_cell_markers_with_metrics(&anchor, &DefaultDrawingMetrics);
        assert_eq!(from.col, u16::MAX);
        assert_eq!(from.row, u32::MAX);
        assert_eq!(to.col, u16::MAX);
        assert_eq!(to.row, u32::MAX);
    }
}
