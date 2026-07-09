//! Form control reading: worksheet `<control>` entries, `ctrlProps`
//! part parsing ([MS-XLSX] `formControlPr`), and assembly into model
//! [`FormControl`]s with captions sourced from the legacy VML
//! drawing part.

use duke_sheets_chart::{CellMarker, DrawingAnchor, EditAs};
use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};
use quick_xml::events::Event;
use quick_xml::Reader;

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

/// Build the anchor from parsed `<from>`/`<to>` marker values plus
/// the CT_ObjectAnchor move/size attributes.
pub(super) fn anchor_from_markers(
    from: [i64; 4],
    to: [i64; 4],
    move_with_cells: bool,
    size_with_cells: bool,
) -> DrawingAnchor {
    let edit_as = match (move_with_cells, size_with_cells) {
        (true, true) => Some(EditAs::TwoCell),
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
                    let num = value.parse::<u16>().unwrap_or(0);
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
                        b"sel" => pr.sel = num,
                        b"multiSel" => {
                            pr.multi_sel = value
                                .split(|c: char| !c.is_ascii_digit())
                                .filter(|s| !s.is_empty())
                                .filter_map(|s| s.parse().ok())
                                .collect();
                        }
                        b"seltype" => pr.sel_type = value,
                        b"dropLines" => pr.drop_lines = num,
                        b"inc" => pr.inc = num,
                        b"min" => pr.min = num,
                        b"max" => pr.max = num,
                        b"page" => pr.page = num,
                        b"val" => pr.val = num,
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
) -> Option<FormControl> {
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
            let selected = if selection == ListSelection::Single {
                if pr.sel > 0 {
                    vec![pr.sel]
                } else {
                    Vec::new()
                }
            } else if !pr.multi_sel.is_empty() {
                pr.multi_sel.clone()
            } else if pr.sel > 0 {
                vec![pr.sel]
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
            selected: if pr.sel > 0 { Some(pr.sel) } else { None },
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

    let mut control = FormControl::with_anchor(kind, pending.anchor.clone().unwrap_or_default());
    control.name = pending.name.clone();
    control.locked = pending.locked;
    control.printable = pending.printable;
    Some(control)
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
    }

    #[test]
    fn parses_excel_scroll_ctrl_prop() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Scroll" dx="22" fmlaLink="$D$6" inc="2" max="95" min="5" page="10" val="40"/>"#;
        let pr = parse_ctrl_prop(xml).expect("parse");
        let control = assemble(&PendingControl::new(), &pr, String::new()).expect("assemble");
        match control.kind {
            FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                ..
            } => {
                assert_eq!((value, min, max, increment, page), (40, 5, 95, 2, 10));
            }
            other => panic!("expected Scrollbar, got {other:?}"),
        }
    }
}
