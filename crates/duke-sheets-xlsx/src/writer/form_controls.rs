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
    CheckState, DrawingMeta, FormControl, FormControlKind, HorizontalAlignment, ListSelection,
    VerticalAlignment,
};
use duke_sheets_vml::anchor_cell_markers_with_metrics;

use super::{XlsxError, XlsxResult};

/// The `formControlPr/@objectType` name (differs from the VML
/// ObjectType for checkboxes: `CheckBox` vs `Checkbox`).
pub(super) fn ctrl_prop_object_type(kind: &FormControlKind) -> &str {
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
        FormControlKind::Unknown { object_type, .. } => object_type,
    }
}

pub(super) use duke_sheets_vml::default_control_name;

fn escape_attr(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn horizontal_alignment_value(alignment: HorizontalAlignment) -> &'static str {
    match alignment {
        HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::Distributed => "distributed",
        HorizontalAlignment::General | HorizontalAlignment::Left | HorizontalAlignment::Fill => {
            "left"
        }
    }
}

fn vertical_alignment_value(alignment: VerticalAlignment) -> &'static str {
    match alignment {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

/// Write one `xl/ctrlProps/ctrlProp{num}.xml` part.
pub(super) fn write_ctrl_prop_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    num: usize,
    control: &FormControl,
    first_button: bool,
) -> XlsxResult<()> {
    control.validate()?;
    let path = format!("xl/ctrlProps/ctrlProp{num}.xml");
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options).map_err(XlsxError::from)?;

    let mut attrs: Vec<(&str, String)> = Vec::new();
    let push_checked = |attrs: &mut Vec<(&str, String)>, state: &CheckState| match state {
            CheckState::Unchecked => {}
            CheckState::Checked => attrs.push(("checked", "Checked".to_string())),
            CheckState::Mixed => attrs.push(("checked", "Mixed".to_string())),
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
        FormControlKind::Unknown { raw_properties, .. } => {
            for (name, value) in raw_properties {
                if name != "objectType" && name != "xmlns" && !name.starts_with("xmlns:") {
                    attrs.push((name.as_str(), value.clone()));
                }
            }
        }
    }
    if let Some(caption) = control.caption() {
        if let Some(alignment) = caption.horizontal_alignment {
            attrs.push((
                "textHAlign",
                horizontal_alignment_value(alignment).to_string(),
            ));
        }
        if let Some(alignment) = caption.vertical_alignment {
            attrs.push((
                "textVAlign",
                vertical_alignment_value(alignment).to_string(),
            ));
        }
    }
    if let Some(macro_name) = &control.macro_name {
        attrs.push((
            "macro",
            duke_sheets_vml::encode_macro_formula(macro_name),
        ));
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
pub(super) fn controls_block(
    entries: &[ControlEntry<'_>],
    metrics: &(impl duke_sheets_chart::DrawingMetrics + ?Sized),
) -> String {
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
        if let Some(macro_name) = &control.macro_name {
            xml.push_str(&format!(
                " macro=\"{}\"",
                escape_attr(&duke_sheets_vml::encode_macro_formula(macro_name))
            ));
        }
        if let Some(alt_text) = &entry.meta.alt_text {
            xml.push_str(&format!(" altText=\"{}\"", escape_attr(alt_text)));
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
            | FormControlKind::Spinner { .. }
            | FormControlKind::Unknown { .. } => {
                xml.push_str(" autoPict=\"0\"");
            }
        }
        xml.push('>');

        // CT_ObjectAnchor: moveWithCells / sizeWithCells default
        // false; EMU offsets.
        let (move_wc, size_wc) = match &entry.anchor {
            duke_sheets_chart::DrawingAnchor::TwoCell { edit_as, .. } => {
                match edit_as
                    .clone()
                    .unwrap_or(duke_sheets_chart::EditAs::TwoCell)
                {
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
        let (from, to) = anchor_cell_markers_with_metrics(&entry.anchor, metrics);
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

pub(super) use duke_sheets_vml::{radio_head_flags, sheet_controls};
