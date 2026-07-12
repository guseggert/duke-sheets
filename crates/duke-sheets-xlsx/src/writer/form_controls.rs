//! Form control emission: `ctrlProps` parts and the worksheet
//! `<controls>` block (ECMA-376 CT_Control/CT_ControlPr plus the
//! [MS-XLSX] 2009/9 `formControlPr` extension). The control shapes
//! themselves live in the legacy VML drawing part, emitted alongside
//! comment shapes by `comments::write_vml_drawing`.
//!
//! Markup mirrors Excel's own output: `objectType` first with the
//! remaining attributes alphabetical, per-kind `controlPr`
//! auto-flags, and per-kind `formControlPr` defaults - pinned
//! against Excel-authored files.

use std::io::{Seek, Write};

use duke_sheets_core::{
    CheckState, DrawingMeta, FormControl, FormControlKind, ListSelection, Worksheet,
};
use duke_sheets_vml::anchor_cell_markers;

use super::{XlsxError, XlsxResult};

/// The `formControlPr/@objectType` name (differs from the VML
/// ObjectType for checkboxes: `CheckBox` vs `Checkbox`).
pub(super) fn ctrl_prop_object_type(kind: &FormControlKind) -> &'static str {
    match kind {
        FormControlKind::Button { .. } => "Button",
        FormControlKind::Checkbox { .. } => "CheckBox",
        FormControlKind::OptionButton { .. } => "Radio",
        FormControlKind::Label { .. } => "Label",
        FormControlKind::GroupBox { .. } => "GBox",
        FormControlKind::ListBox { .. } => "List",
        FormControlKind::Dropdown { .. } => "Drop",
        FormControlKind::Scrollbar { .. } => "Scroll",
        FormControlKind::Spinner { .. } => "Spin",
    }
}

/// Excel-style default shape name ("Check Box 3"). `seq` is the
/// 1-based per-sheet drawing object number.
pub(super) fn default_control_name(kind: &FormControlKind, seq: usize) -> String {
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
    };
    format!("{base} {seq}")
}

fn escape_attr(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Write one `xl/ctrlProps/ctrlProp{num}.xml` part.
pub(super) fn write_ctrl_prop_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    num: usize,
    control: &FormControl,
    first_button: bool,
) -> XlsxResult<()> {
    let path = format!("xl/ctrlProps/ctrlProp{num}.xml");
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options)
        .map_err(XlsxError::from)?;

    let mut attrs: Vec<(&str, String)> = Vec::new();
    let push_checked = |attrs: &mut Vec<(&str, String)>, state: &CheckState| {
        match state {
            CheckState::Unchecked => {}
            CheckState::Checked => attrs.push(("checked", "Checked".to_string())),
            CheckState::Mixed => attrs.push(("checked", "Mixed".to_string())),
        }
    };
    let push_link = |attrs: &mut Vec<(&str, String)>, link: &Option<String>| {
        if let Some(link) = link {
            attrs.push(("fmlaLink", link.clone()));
        }
    };
    let push_range = |attrs: &mut Vec<(&str, String)>, range: &Option<String>| {
        if let Some(range) = range {
            attrs.push(("fmlaRange", range.clone()));
        }
    };

    match &control.kind {
        FormControlKind::Button { .. } => {
            attrs.push(("lockText", "1".to_string()));
        }
        FormControlKind::Checkbox {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            push_checked(&mut attrs, state);
            push_link(&mut attrs, cell_link);
            attrs.push(("lockText", "1".to_string()));
            if *no_3d {
                attrs.push(("noThreeD", "1".to_string()));
            }
        }
        FormControlKind::OptionButton {
            state,
            cell_link,
            no_3d,
            ..
        } => {
            push_checked(&mut attrs, state);
            if first_button {
                attrs.push(("firstButton", "1".to_string()));
            }
            push_link(&mut attrs, cell_link);
            attrs.push(("lockText", "1".to_string()));
            if *no_3d {
                attrs.push(("noThreeD", "1".to_string()));
            }
        }
        FormControlKind::Label { .. } => {
            attrs.push(("lockText", "1".to_string()));
        }
        FormControlKind::GroupBox { no_3d, .. } => {
            if *no_3d {
                attrs.push(("noThreeD", "1".to_string()));
            }
        }
        FormControlKind::ListBox {
            input_range,
            cell_link,
            selection,
            selected,
            no_3d,
        } => {
            attrs.push(("dx", "22".to_string()));
            push_link(&mut attrs, cell_link);
            push_range(&mut attrs, input_range);
            // Model indices are zero-based; the attributes are
            // one-based (0 = none).
            if !selected.is_empty() && !matches!(selection, ListSelection::Single) {
                let list = selected
                    .iter()
                    .map(|&v| (u32::from(v) + 1).to_string())
                    .collect::<Vec<_>>()
                    .join(",");
                attrs.push(("multiSel", list));
            }
            if *no_3d {
                attrs.push(("noThreeD", "1".to_string()));
            }
            if matches!(selection, ListSelection::Single) {
                if let Some(&first) = selected.first() {
                    attrs.push(("sel", (u32::from(first) + 1).to_string()));
                }
            }
            match selection {
                ListSelection::Single => {}
                ListSelection::Multi => attrs.push(("seltype", "multi".to_string())),
                ListSelection::Extend => attrs.push(("seltype", "extended".to_string())),
            }
            attrs.push(("val", "0".to_string()));
        }
        FormControlKind::Dropdown {
            input_range,
            cell_link,
            selected,
            lines,
            no_3d,
        } => {
            if *lines != 8 {
                attrs.push(("dropLines", lines.to_string()));
            }
            attrs.push(("dropStyle", "combo".to_string()));
            attrs.push(("dx", "22".to_string()));
            push_link(&mut attrs, cell_link);
            push_range(&mut attrs, input_range);
            if *no_3d {
                attrs.push(("noThreeD", "1".to_string()));
            }
            if let Some(sel) = selected {
                attrs.push(("sel", (u32::from(*sel) + 1).to_string()));
            }
            attrs.push(("val", "0".to_string()));
        }
        FormControlKind::Scrollbar {
            value,
            min,
            max,
            increment,
            page,
            horizontal,
            cell_link,
        } => {
            attrs.push(("dx", "22".to_string()));
            push_link(&mut attrs, cell_link);
            if *horizontal {
                attrs.push(("horiz", "1".to_string()));
            }
            if *increment != 1 {
                attrs.push(("inc", increment.to_string()));
            }
            attrs.push(("max", max.to_string()));
            if *min != 0 {
                attrs.push(("min", min.to_string()));
            }
            attrs.push(("page", page.to_string()));
            attrs.push(("val", value.to_string()));
        }
        FormControlKind::Spinner {
            value,
            min,
            max,
            increment,
            cell_link,
        } => {
            attrs.push(("dx", "22".to_string()));
            push_link(&mut attrs, cell_link);
            if *increment != 1 {
                attrs.push(("inc", increment.to_string()));
            }
            attrs.push(("max", max.to_string()));
            if *min != 0 {
                attrs.push(("min", min.to_string()));
            }
            attrs.push(("page", "10".to_string()));
            attrs.push(("val", value.to_string()));
        }
    }

    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    xml.push_str(&format!(
        "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"{}\"",
        ctrl_prop_object_type(&control.kind)
    ));
    for (name, value) in &attrs {
        xml.push_str(&format!(" {name}=\"{}\"", escape_attr(value)));
    }
    xml.push_str("/>");
    zip.write_all(xml.as_bytes()).map_err(XlsxError::from)?;
    Ok(())
}

/// Build the worksheet `<mc:AlternateContent><mc:Choice
/// Requires="x14"><controls>...` block referencing the sheet's
/// controls. `entries` pairs each control with its shape id, rel id,
/// display name, and radio-group-head flag.
pub(super) fn controls_block(entries: &[ControlEntry<'_>]) -> String {
    const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let mut xml = String::new();
    xml.push_str(&format!(
        "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"><mc:Choice Requires=\"x14\"><controls>"
    ));
    for entry in entries {
        let control = entry.control;
        xml.push_str(&format!(
            "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"><mc:Choice Requires=\"x14\">"
        ));
        xml.push_str(&format!(
            "<control shapeId=\"{}\" r:id=\"{}\" name=\"{}\">",
            entry.shape_id,
            entry.rid,
            escape_attr(&entry.name)
        ));
        xml.push_str("<controlPr defaultSize=\"0\"");
        if !entry.meta.locked {
            xml.push_str(" locked=\"0\"");
        }
        if !entry.meta.printable {
            xml.push_str(" print=\"0\"");
        }
        // Per-kind auto flags, mirroring Excel's emit.
        match &control.kind {
            FormControlKind::Button { .. }
            | FormControlKind::Dropdown { .. }
            | FormControlKind::Label { .. } => {
                xml.push_str(" autoFill=\"0\" autoPict=\"0\"");
            }
            FormControlKind::Checkbox { .. } | FormControlKind::OptionButton { .. } => {
                xml.push_str(" autoFill=\"0\" autoLine=\"0\" autoPict=\"0\"");
            }
            FormControlKind::GroupBox { .. }
            | FormControlKind::ListBox { .. }
            | FormControlKind::Scrollbar { .. }
            | FormControlKind::Spinner { .. } => {
                xml.push_str(" autoPict=\"0\"");
            }
        }
        xml.push('>');

        // CT_ObjectAnchor: moveWithCells / sizeWithCells default
        // false; EMU offsets.
        let (move_wc, size_wc) = match &entry.anchor {
            duke_sheets_chart::DrawingAnchor::TwoCell { edit_as, .. } => {
                match edit_as.clone().unwrap_or(duke_sheets_chart::EditAs::TwoCell) {
                    duke_sheets_chart::EditAs::TwoCell => (true, true),
                    duke_sheets_chart::EditAs::OneCell => (true, false),
                    duke_sheets_chart::EditAs::Absolute => (false, false),
                }
            }
            duke_sheets_chart::DrawingAnchor::OneCell { .. } => (true, false),
            duke_sheets_chart::DrawingAnchor::Absolute { .. } => (false, false),
        };
        xml.push_str("<anchor");
        if move_wc {
            xml.push_str(" moveWithCells=\"1\"");
        }
        if size_wc {
            xml.push_str(" sizeWithCells=\"1\"");
        }
        xml.push('>');
        let (from, to) = anchor_cell_markers(&entry.anchor);
        for (tag, marker) in [("from", &from), ("to", &to)] {
            xml.push_str(&format!(
                "<{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></{tag}>",
                marker.col, marker.col_offset_emu, marker.row, marker.row_offset_emu
            ));
        }
        xml.push_str("</anchor></controlPr></control></mc:Choice><mc:Fallback>");
        xml.push_str(&format!(
            "<control shapeId=\"{}\" r:id=\"{}\" name=\"{}\"/>",
            entry.shape_id,
            entry.rid,
            escape_attr(&entry.name)
        ));
        xml.push_str("</mc:Fallback></mc:AlternateContent>");
    }
    xml.push_str("</controls></mc:Choice></mc:AlternateContent>");
    xml
}

/// One control's identifiers for the worksheet block. `meta` and
/// `anchor` come from the control's wrapping drawing object (or, for
/// controls nested in groups, its resolved placement).
pub(super) struct ControlEntry<'a> {
    pub control: &'a FormControl,
    pub meta: &'a DrawingMeta,
    pub anchor: duke_sheets_chart::DrawingAnchor,
    pub shape_id: usize,
    pub rid: String,
    pub name: String,
}

/// One control in the sheet's emission sequence: every form control
/// in the drawing tree, in [`Worksheet::placed_form_controls`]
/// (depth-first) order. This order drives shape ids, ctrlProp part
/// numbering, `<controls>` entries, drawing twins, and VML shapes.
pub(super) struct SheetControl<'a> {
    pub payload: &'a FormControl,
    pub meta: &'a DrawingMeta,
    /// Top-level controls keep their wrapper anchor; group children
    /// get an absolute anchor from their resolved on-sheet rectangle.
    pub anchor: duke_sheets_chart::DrawingAnchor,
}

/// The sheet's control sequence in placed (depth-first) order.
pub(super) fn sheet_controls(sheet: &Worksheet) -> Vec<SheetControl<'_>> {
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
/// (depth-first) control order, derived from the spatial grouping in
/// [`duke_sheets_core::radio_groups`].
pub(super) fn radio_head_flags(sheet: &Worksheet) -> Vec<bool> {
    let placed = sheet.placed_form_controls();
    let mut flags = vec![false; placed.len()];
    for group in duke_sheets_core::radio_groups(&placed) {
        if let Some(&head) = group.first() {
            flags[head] = true;
        }
    }
    flags
}
