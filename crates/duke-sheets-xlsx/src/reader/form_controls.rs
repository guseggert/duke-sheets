//! Form control reading: worksheet `<control>` entries, `ctrlProps`
//! part parsing ([MS-XLSX] `formControlPr`), and assembly into model
//! [`DrawingObject`]s wrapping [`FormControl`]s, with captions
//! sourced from the legacy VML drawing part.

use duke_sheets_chart::{CellMarker, DrawingAnchor, EditAs};
use duke_sheets_core::{
    CheckState, DrawingObject, FormControl, FormControlKind, ListSelection,
};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashSet;

/// A `<control>` element collected during worksheet parsing, before
/// its ctrlProp part and VML caption are resolved.
#[derive(Debug, Clone, Default)]
pub(super) struct PendingControl {
    /// `control/@shapeId` - matches the VML shape `_x0000_s{N}`.
    pub shape_id: u32,
    /// `control/@name`.
    pub name: Option<String>,
    /// `control/@r:id` - the ctrlProp relationship.
    pub rid: String,
    /// Anchor from `controlPr/anchor` (EMU offsets).
    pub anchor: Option<DrawingAnchor>,
    /// `controlPr/@locked` (default true).
    pub locked: bool,
    /// `controlPr/@print` (default true).
    pub printable: bool,
}

impl PendingControl {
    pub(super) fn new() -> Self {
        PendingControl {
            locked: true,
            printable: true,
            ..Default::default()
        }
    }
}

/// Markup-compatibility fallbacks may repeat the same `<control>` in
/// expanded rather than self-closing form. Keep the first entry for
/// each ctrlProp relationship so serialization style cannot duplicate
/// model controls.
pub(super) fn dedupe_pending_controls(
    controls: Vec<PendingControl>,
) -> Vec<PendingControl> {
    let mut seen = HashSet::new();
    controls
        .into_iter()
        .filter(|control| seen.insert(control.rid.clone()))
        .collect()
}

/// Build the anchor from parsed `<from>`/`<to>` marker values plus
/// the CT_ObjectAnchor move/size attributes. Move+size (Excel's
/// default) maps to `edit_as: None`, its canonical model form.
pub(super) fn anchor_from_markers(
    from: [i64; 4],
    to: [i64; 4],
    move_with_cells: bool,
    size_with_cells: bool,
) -> DrawingAnchor {
    let edit_as = match (move_with_cells, size_with_cells) {
        (true, true) => None,
        (true, false) => Some(EditAs::OneCell),
        _ => Some(EditAs::Absolute),
    };
    let marker = |v: [i64; 4]| CellMarker {
        col: v[0].clamp(0, u16::MAX as i64) as u16,
        col_offset_emu: v[1],
        row: v[2].clamp(0, u32::MAX as i64) as u32,
        row_offset_emu: v[3],
    };
    DrawingAnchor::TwoCell {
        from: marker(from),
        to: marker(to),
        edit_as,
    }
}

/// Parsed `formControlPr` attributes.
#[derive(Debug, Clone, Default)]
pub(super) struct CtrlProp {
    pub object_type: String,
    pub checked: u16,
    pub fmla_link: Option<String>,
    pub fmla_range: Option<String>,
    pub sel: u16,
    pub multi_sel: Vec<u16>,
    pub sel_type: String,
    pub drop_lines: u16,
    pub inc: u16,
    pub min: u16,
    pub max: u16,
    pub page: u16,
    pub val: u16,
    pub horiz: bool,
    pub first_button: bool,
    pub no_3d: bool,
}

/// Parse a `xl/ctrlProps/ctrlPropN.xml` part. Returns `None` when
/// the part has no recognizable `formControlPr` element.
pub(super) fn parse_ctrl_prop(bytes: &[u8]) -> Option<CtrlProp> {
    let mut reader = Reader::from_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == b"formControlPr" =>
            {
                let mut pr = CtrlProp {
                    drop_lines: 8,
                    inc: 1,
                    page: 10,
                    ..Default::default()
                };
                for attr in e.attributes().flatten() {
                    let value = String::from_utf8_lossy(&attr.value).into_owned();
                    let num = value.parse::<u16>().ok();
                    let truthy = value == "1" || value.eq_ignore_ascii_case("true");
                    match attr.key.local_name().as_ref() {
                        b"objectType" => pr.object_type = value,
                        b"checked" => {
                            pr.checked = match value.as_str() {
                                "Checked" => 1,
                                "Mixed" => 2,
                                _ => 0,
                            }
                        }
                        b"fmlaLink" => pr.fmla_link = Some(value),
                        b"fmlaRange" => pr.fmla_range = Some(value),
                        b"sel" => pr.sel = num.unwrap_or(0),
                        b"multiSel" => {
                            // Excel's emit order is not stable
                            // ("3, 1"); normalize to sorted indices.
                            pr.multi_sel = value
                                .split(|c: char| !c.is_ascii_digit())
                                .filter(|s| !s.is_empty())
                                .filter_map(|s| s.parse().ok())
                                .collect();
                            pr.multi_sel.sort_unstable();
                            pr.multi_sel.dedup();
                        }
                        b"seltype" => pr.sel_type = value,
                        b"dropLines" => {
                            if let Some(num) = num {
                                pr.drop_lines = num;
                            }
                        }
                        b"inc" => {
                            if let Some(num) = num {
                                pr.inc = num;
                            }
                        }
                        b"min" => pr.min = num.unwrap_or(0),
                        b"max" => pr.max = num.unwrap_or(0),
                        b"page" => {
                            if let Some(num) = num {
                                pr.page = num;
                            }
                        }
                        b"val" => pr.val = num.unwrap_or(0),
                        b"horiz" => pr.horiz = truthy,
                        b"firstButton" => pr.first_button = truthy,
                        b"noThreeD" | b"noThreeD2" => pr.no_3d = pr.no_3d || truthy,
                        _ => {}
                    }
                }
                return Some(pr);
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

/// Assemble a model control from the worksheet entry, its ctrlProp
/// attributes, and the caption extracted from the VML shape.
pub(super) fn assemble(
    pending: &PendingControl,
    pr: &CtrlProp,
    caption: String,
) -> Option<DrawingObject> {
    let state = match pr.checked {
        2 => CheckState::Mixed,
        0 => CheckState::Unchecked,
        _ => CheckState::Checked,
    };
    let kind = match pr.object_type.as_str() {
        "Button" => FormControlKind::Button { caption },
        "CheckBox" => FormControlKind::Checkbox {
            caption,
            state,
            cell_link: pr.fmla_link.clone(),
            no_3d: pr.no_3d,
        },
        "Radio" => FormControlKind::OptionButton {
            caption,
            state: if state == CheckState::Mixed {
                CheckState::Checked
            } else {
                state
            },
            cell_link: pr.fmla_link.clone(),
            first_in_group: pr.first_button,
            no_3d: pr.no_3d,
        },
        "Label" => FormControlKind::Label { caption },
        "GBox" => FormControlKind::GroupBox {
            caption,
            no_3d: pr.no_3d,
        },
        "List" => {
            let selection = match pr.sel_type.as_str() {
                "multi" => ListSelection::Multi,
                "extended" => ListSelection::Extend,
                _ => ListSelection::Single,
            };
            // Attribute values are one-based (0 = none); the model is
            // zero-based.
            let selected = if selection == ListSelection::Single {
                if pr.sel > 0 {
                    vec![pr.sel - 1]
                } else {
                    Vec::new()
                }
            } else if !pr.multi_sel.is_empty() {
                pr.multi_sel
                    .iter()
                    .filter(|&&v| v > 0)
                    .map(|&v| v - 1)
                    .collect()
            } else if pr.sel > 0 {
                vec![pr.sel - 1]
            } else {
                Vec::new()
            };
            FormControlKind::ListBox {
                input_range: pr.fmla_range.clone(),
                cell_link: pr.fmla_link.clone(),
                selection,
                selected,
                no_3d: pr.no_3d,
            }
        }
        "Drop" => FormControlKind::Dropdown {
            input_range: pr.fmla_range.clone(),
            cell_link: pr.fmla_link.clone(),
            selected: if pr.sel > 0 { Some(pr.sel - 1) } else { None },
            lines: pr.drop_lines,
            no_3d: pr.no_3d,
        },
        "Scroll" => FormControlKind::Scrollbar {
            value: pr.val,
            min: pr.min,
            max: pr.max,
            increment: pr.inc,
            page: pr.page,
            horizontal: pr.horiz,
            cell_link: pr.fmla_link.clone(),
        },
        "Spin" => FormControlKind::Spinner {
            value: pr.val,
            min: pr.min,
            max: pr.max,
            increment: pr.inc,
            cell_link: pr.fmla_link.clone(),
        },
        // EditBox / Dialog are dialog-sheet controls; unknown types
        // are skipped.
        _ => return None,
    };

    let mut object = DrawingObject::form_control(FormControl::new(kind))
        .with_anchor(pending.anchor.clone().unwrap_or_default());
    object.meta.name = pending.name.clone();
    object.meta.locked = pending.locked;
    object.meta.printable = pending.printable;
    Some(object)
}

pub(super) fn assemble_with_vml(
    pending: &PendingControl,
    pr: &CtrlProp,
    vml: Option<&duke_sheets_vml::VmlControl>,
) -> Option<DrawingObject> {
    let caption = vml.map(|shape| shape.caption.clone()).unwrap_or_default();
    let mut object = assemble(pending, pr, caption)?;
    if pending.anchor.is_none() {
        if let Some(vml_object) = vml.and_then(|shape| shape.to_drawing_object()) {
            object.anchor = vml_object.anchor;
        }
    }
    Some(object)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_excel_checkbox_ctrl_prop() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked" fmlaLink="$D$2" lockText="1" noThreeD="1"/>"#;
        let pr = parse_ctrl_prop(xml).expect("parse");
        assert_eq!(pr.object_type, "CheckBox");
        assert_eq!(pr.checked, 1);
        assert_eq!(pr.fmla_link.as_deref(), Some("$D$2"));
        assert!(pr.no_3d);
    }

    #[test]
    fn parses_excel_dropdown_ctrl_prop() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Drop" dropLines="6" dropStyle="combo" dx="22" fmlaLink="$D$4" fmlaRange="$H$1:$H$4" noThreeD="1" sel="2" val="0"/>"#;
        let pr = parse_ctrl_prop(xml).expect("parse");
        assert_eq!(pr.object_type, "Drop");
        assert_eq!(pr.drop_lines, 6);
        assert_eq!(pr.sel, 2);
        assert_eq!(pr.fmla_range.as_deref(), Some("$H$1:$H$4"));

        // The one-based sel attribute becomes a zero-based model index.
        let object = assemble(&PendingControl::new(), &pr, String::new()).expect("assemble");
        let control = object.kind.as_form_control().expect("form control");
        match &control.kind {
            FormControlKind::Dropdown { selected, .. } => assert_eq!(*selected, Some(1)),
            other => panic!("expected Dropdown, got {other:?}"),
        }
    }

    #[test]
    fn parses_excel_scroll_ctrl_prop() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Scroll" dx="22" fmlaLink="$D$6" inc="2" max="95" min="5" page="10" val="40"/>"#;
        let pr = parse_ctrl_prop(xml).expect("parse");
        let object = assemble(&PendingControl::new(), &pr, String::new()).expect("assemble");
        let control = object.kind.as_form_control().expect("form control");
        match &control.kind {
            FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                ..
            } => {
                assert_eq!(
                    (*value, *min, *max, *increment, *page),
                    (40, 5, 95, 2, 10)
                );
            }
            other => panic!("expected Scrollbar, got {other:?}"),
        }
    }

    #[test]
    fn duplicate_fallback_controls_are_deduplicated_by_relationship() {
        let mut first = PendingControl::new();
        first.rid = "rId3".to_string();
        first.shape_id = 1025;
        let mut fallback = first.clone();
        fallback.shape_id = 9999;
        let mut second = PendingControl::new();
        second.rid = "rId4".to_string();

        let deduped = dedupe_pending_controls(vec![first, fallback, second]);
        assert_eq!(deduped.len(), 2);
        assert_eq!(deduped[0].shape_id, 1025);
        assert_eq!(deduped[1].rid, "rId4");
    }

    #[test]
    fn missing_control_pr_anchor_falls_back_to_vml_anchor() {
        let mut pending = PendingControl::new();
        pending.rid = "rId3".to_string();
        let pr = CtrlProp {
            object_type: "CheckBox".to_string(),
            ..Default::default()
        };
        let vml = duke_sheets_vml::VmlControl {
            object_type: "Checkbox".to_string(),
            anchor_px: Some([1, 2, 3, 4, 5, 6, 7, 8]),
            move_with_cells: true,
            size_with_cells: false,
            ..Default::default()
        };
        let object = assemble_with_vml(&pending, &pr, Some(&vml)).unwrap();
        match object.anchor {
            DrawingAnchor::TwoCell { from, to, edit_as } => {
                assert_eq!((from.col, from.row), (1, 3));
                assert_eq!((to.col, to.row), (5, 7));
                assert_eq!(edit_as, Some(EditAs::OneCell));
            }
            other => panic!("expected TwoCell anchor, got {other:?}"),
        }
    }
}
