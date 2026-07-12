//! Form controls (legacy Forms toolbar controls)
//!
//! Worksheet-embedded interactive controls: buttons, checkboxes,
//! option (radio) buttons, labels, group boxes, list boxes,
//! dropdowns (combo boxes), scrollbars, and spinners.
//!
//! A form control is a drawing object: its placement, shape name,
//! and locked/printable/hidden flags live on the wrapping
//! [`crate::DrawingObject`] in the worksheet's z-ordered drawing
//! list. [`FormControl`] carries the control-specific state.
//!
//! Cell links and input ranges are stored as A1-style formula text
//! with an optional sheet prefix (e.g. `"$A$1"` or `"Sheet2!$B$2:$B$9"`),
//! matching the convention used by data validation formulas.
//!
//! List box and dropdown selection indices are zero-based, like every
//! other index in this library. The file formats store them one-based
//! (0 = no selection); codecs convert at the read/write boundary. Note
//! that a *linked cell* still receives Excel's one-based value at
//! runtime.
//!
//! [`crate::Workbook::sync_form_control_links`] projects control state
//! into linked cells using Excel's runtime conventions. The high-level
//! save APIs serialize a synchronized snapshot without changing the caller's
//! workbook. Synchronization replaces disagreeing cell values and formulas in
//! the output; a linked cell that already holds the control's value is left
//! untouched, so a formula driving the control survives, as in Excel.
//! Low-level format writers are immutable and require an explicit call
//! before writing.
//!
//! Calculation binds the two live, as in Excel: control state is projected
//! into constant linked cells before evaluation (so formulas see it), and
//! recalculated formulas in linked cells drive their controls afterwards via
//! [`crate::Workbook::sync_form_controls_from_linked_cells`].
//!
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{
//!     CheckState, DrawingObject, FormControl, FormControlKind, Workbook,
//! };
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! sheet.try_add_drawing(DrawingObject::form_control(FormControl::new(
//!     FormControlKind::Checkbox {
//!         caption: "Enable feature".to_string(),
//!         state: CheckState::Checked,
//!         cell_link: Some("$D$2".to_string()),
//!         no_3d: false,
//!     },
//! ))).unwrap();
//!
//! assert_eq!(sheet.form_control_count(), 1);
//! ```

use crate::drawing::{DrawingPath, RectEmu};
use crate::{Error, Result};

/// State of a checkbox or option button.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CheckState {
    /// Not checked.
    #[default]
    Unchecked,
    /// Checked.
    Checked,
    /// Mixed (indeterminate). Only valid for checkboxes, not option
    /// buttons (MS-XLS 2.5.141 FtCblsData.fChecked).
    Mixed,
}

/// Selection behavior of a list box (MS-XLS 2.5.147 FtLbsData.wListSelType).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ListSelection {
    /// One item at a time.
    #[default]
    Single,
    /// Any combination of items (simple multi-select).
    Multi,
    /// Contiguous runs of items (extended multi-select).
    Extend,
}

/// Kind-specific properties of a form control.
#[derive(Debug, Clone, PartialEq)]
pub enum FormControlKind {
    /// Push button. Macro assignment is not modeled.
    Button {
        /// Text shown on the button face.
        caption: String,
    },
    /// Checkbox with an optional linked cell.
    Checkbox {
        /// Text shown beside the box.
        caption: String,
        /// Checked / unchecked / mixed.
        state: CheckState,
        /// Cell receiving TRUE/FALSE (A1-style, optional sheet prefix).
        cell_link: Option<String>,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// Option (radio) button with an optional linked cell.
    OptionButton {
        /// Text shown beside the button.
        caption: String,
        /// Checked or unchecked ([`CheckState::Mixed`] is invalid here).
        state: CheckState,
        /// Cell receiving the one-based index of the selected option
        /// in the group (A1-style, optional sheet prefix).
        cell_link: Option<String>,
        /// Whether this is the first button of its group. Read-side
        /// information; writers recompute groups from the innermost
        /// enclosing group box (or the sheet-level group) and mark the
        /// first radio in insertion order.
        first_in_group: bool,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// Static text label.
    Label {
        /// Label text.
        caption: String,
    },
    /// Group box framing a set of option buttons.
    GroupBox {
        /// Text shown in the frame's top edge.
        caption: String,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// List box populated from a cell range.
    ListBox {
        /// Range providing the list items (A1-style, optional sheet prefix).
        input_range: Option<String>,
        /// Cell receiving the one-based index of the selected item
        /// (Excel's linked-cell convention).
        cell_link: Option<String>,
        /// Selection behavior.
        selection: ListSelection,
        /// Zero-based indices of selected items, sorted ascending. At
        /// most one entry for [`ListSelection::Single`]; any number
        /// for multi/extend.
        selected: Vec<u16>,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// Dropdown (combo box) populated from a cell range.
    Dropdown {
        /// Range providing the list items (A1-style, optional sheet prefix).
        input_range: Option<String>,
        /// Cell receiving the one-based index of the selected item
        /// (Excel's linked-cell convention).
        cell_link: Option<String>,
        /// Zero-based index of the selected item; `None` = no selection.
        selected: Option<u16>,
        /// Number of lines shown when the list is dropped down.
        lines: u16,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// Scrollbar with an optional linked cell.
    Scrollbar {
        /// Current value.
        value: u16,
        /// Minimum value (Excel UI allows 0-30000).
        min: u16,
        /// Maximum value (Excel UI allows 0-30000).
        max: u16,
        /// Step per arrow click ("incremental change").
        increment: u16,
        /// Step per trough click ("page change").
        page: u16,
        /// Horizontal orientation (vertical when false).
        horizontal: bool,
        /// Cell receiving the current value.
        cell_link: Option<String>,
    },
    /// Spinner (spin button) with an optional linked cell.
    Spinner {
        /// Current value.
        value: u16,
        /// Minimum value (Excel UI allows 0-30000).
        min: u16,
        /// Maximum value (Excel UI allows 0-30000).
        max: u16,
        /// Step per arrow click.
        increment: u16,
        /// Cell receiving the current value.
        cell_link: Option<String>,
    },
}

impl FormControlKind {
    /// The control's caption text, for kinds that have one.
    pub fn caption(&self) -> Option<&str> {
        match self {
            FormControlKind::Button { caption }
            | FormControlKind::Checkbox { caption, .. }
            | FormControlKind::OptionButton { caption, .. }
            | FormControlKind::Label { caption }
            | FormControlKind::GroupBox { caption, .. } => Some(caption),
            _ => None,
        }
    }

    /// The control's cell link formula, for kinds that support one.
    pub fn cell_link(&self) -> Option<&str> {
        match self {
            FormControlKind::Checkbox { cell_link, .. }
            | FormControlKind::OptionButton { cell_link, .. }
            | FormControlKind::ListBox { cell_link, .. }
            | FormControlKind::Dropdown { cell_link, .. }
            | FormControlKind::Scrollbar { cell_link, .. }
            | FormControlKind::Spinner { cell_link, .. } => cell_link.as_deref(),
            _ => None,
        }
    }

    /// Validate kind-specific invariants shared by all bindings and
    /// writers.
    pub fn validate(&self) -> Result<()> {
        let nonempty_formula = |label: &str, value: &Option<String>| -> Result<()> {
            if value.as_ref().is_some_and(|formula| formula.trim().is_empty()) {
                return Err(Error::other(format!("{label} cannot be empty")));
            }
            Ok(())
        };
        match self {
            FormControlKind::Button { .. } | FormControlKind::Label { .. } => {}
            FormControlKind::Checkbox { cell_link, .. } => {
                nonempty_formula("cell link", cell_link)?;
            }
            FormControlKind::OptionButton {
                state, cell_link, ..
            } => {
                if *state == CheckState::Mixed {
                    return Err(Error::other("option buttons cannot use the mixed state"));
                }
                nonempty_formula("cell link", cell_link)?;
            }
            FormControlKind::GroupBox { .. } => {}
            FormControlKind::ListBox {
                input_range,
                cell_link,
                selection,
                selected,
                ..
            } => {
                nonempty_formula("input range", input_range)?;
                nonempty_formula("cell link", cell_link)?;
                if *selection == ListSelection::Single && selected.len() > 1 {
                    return Err(Error::other(
                        "single-select list boxes can select at most one item",
                    ));
                }
                if selected.contains(&u16::MAX) {
                    return Err(Error::other(format!(
                        "list selection index {} exceeds the maximum of {}",
                        u16::MAX,
                        u16::MAX - 1
                    )));
                }
                if selected.windows(2).any(|pair| pair[0] >= pair[1]) {
                    return Err(Error::other(
                        "list selection indices must be sorted and unique",
                    ));
                }
            }
            FormControlKind::Dropdown {
                input_range,
                cell_link,
                selected,
                lines,
                ..
            } => {
                nonempty_formula("input range", input_range)?;
                nonempty_formula("cell link", cell_link)?;
                if selected == &Some(u16::MAX) {
                    return Err(Error::other(format!(
                        "dropdown selection index {} exceeds the maximum of {}",
                        u16::MAX,
                        u16::MAX - 1
                    )));
                }
                if *lines == 0 {
                    return Err(Error::other("dropdown lines must be greater than zero"));
                }
            }
            FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                cell_link,
                ..
            } => {
                validate_numeric_control(*value, *min, *max, *increment)?;
                if *page == 0 {
                    return Err(Error::other("scrollbar page must be greater than zero"));
                }
                nonempty_formula("cell link", cell_link)?;
            }
            FormControlKind::Spinner {
                value,
                min,
                max,
                increment,
                cell_link,
            } => {
                validate_numeric_control(*value, *min, *max, *increment)?;
                nonempty_formula("cell link", cell_link)?;
            }
        }
        Ok(())
    }
}

fn validate_numeric_control(value: u16, min: u16, max: u16, increment: u16) -> Result<()> {
    if min > max || value < min || value > max {
        return Err(Error::other(format!(
            "control requires min <= value <= max, got {min} <= {value} <= {max}"
        )));
    }
    if increment == 0 {
        return Err(Error::other("control increment must be greater than zero"));
    }
    Ok(())
}

/// A worksheet form control (Forms toolbar object).
///
/// Placement, shape name, and protection/print/hidden flags live on
/// the wrapping [`crate::DrawingObject`].
#[derive(Debug, Clone, PartialEq)]
pub struct FormControl {
    /// Kind-specific properties.
    pub kind: FormControlKind,
}

impl FormControl {
    /// Create a control of the given kind.
    pub fn new(kind: FormControlKind) -> Self {
        Self { kind }
    }

    /// The control's caption text, for kinds that have one.
    pub fn caption(&self) -> Option<&str> {
        self.kind.caption()
    }

    /// The control's cell link formula, for kinds that support one.
    pub fn cell_link(&self) -> Option<&str> {
        self.kind.cell_link()
    }

    /// Validate kind-specific properties.
    pub fn validate(&self) -> Result<()> {
        self.kind.validate()
    }
}

/// A form control located in the worksheet's drawing tree: its path,
/// its resolved on-sheet rectangle, and the control payload.
///
/// Produced by [`crate::Worksheet::placed_form_controls`] in
/// depth-first order (top-level objects in z-order, group children
/// within their group).
#[derive(Debug)]
pub struct PlacedControl<'a> {
    /// Path to the control in the drawing tree.
    pub path: DrawingPath,
    /// Absolute EMU rectangle at Excel's default cell metrics.
    pub rect_emu: RectEmu,
    /// The control payload.
    pub control: &'a FormControl,
}

/// Partition a sheet's option buttons into radio groups, mirroring
/// Excel's grouping semantics: each radio belongs to the innermost
/// group box whose rectangle contains the radio's center point, and
/// radios outside every box form the sheet-level group.
///
/// Takes the sheet's controls as produced by
/// [`crate::Worksheet::placed_form_controls`] and returns groups of
/// indices into that slice, in traversal order (both across groups
/// and within each group). Non-radio controls never appear in the
/// result.
///
/// Containment is evaluated in EMU at Excel's default cell metrics
/// (609,600 EMU per column, 190,500 EMU per row), matching the
/// quantisation used by the binary anchor encodings; custom row and
/// column sizes that visually move a radio across a box edge are not
/// accounted for.
pub fn radio_groups(controls: &[PlacedControl<'_>]) -> Vec<Vec<usize>> {
    let box_rects: Vec<RectEmu> = controls
        .iter()
        .filter(|placed| matches!(placed.control.kind, FormControlKind::GroupBox { .. }))
        .map(|placed| placed.rect_emu)
        .collect();

    let containing_box = |rect: RectEmu| -> Option<usize> {
        let (x1, y1, x2, y2) = rect;
        let (cx, cy) = (x1.saturating_add(x2) / 2, y1.saturating_add(y2) / 2);
        box_rects
            .iter()
            .enumerate()
            .filter(|(_, (bx1, by1, bx2, by2))| {
                (*bx1..=*bx2).contains(&cx) && (*by1..=*by2).contains(&cy)
            })
            .min_by_key(|(_, (bx1, by1, bx2, by2))| {
                (bx2 - bx1).max(0).saturating_mul((by2 - by1).max(0))
            })
            .map(|(i, _)| i)
    };

    let mut groups: Vec<(Option<usize>, Vec<usize>)> = Vec::new();
    for (idx, placed) in controls.iter().enumerate() {
        if !matches!(placed.control.kind, FormControlKind::OptionButton { .. }) {
            continue;
        }
        let key = containing_box(placed.rect_emu);
        match groups.iter_mut().find(|(k, _)| *k == key) {
            Some((_, members)) => members.push(idx),
            None => groups.push((key, vec![idx])),
        }
    }
    groups.into_iter().map(|(_, members)| members).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::drawing::DrawingObject;
    use duke_sheets_chart::{CellMarker, DrawingAnchor};

    fn two_cell(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: from_col,
                col_offset_emu: 0,
                row: from_row,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: to_col,
                col_offset_emu: 0,
                row: to_row,
                row_offset_emu: 0,
            },
            edit_as: None,
        }
    }

    fn radio(anchor: DrawingAnchor) -> DrawingObject {
        DrawingObject::form_control(FormControl::new(FormControlKind::OptionButton {
            caption: String::new(),
            state: CheckState::Unchecked,
            cell_link: None,
            first_in_group: false,
            no_3d: false,
        }))
        .with_anchor(anchor)
    }

    fn group_box(caption: &str, anchor: DrawingAnchor) -> DrawingObject {
        DrawingObject::form_control(FormControl::new(FormControlKind::GroupBox {
            caption: caption.to_string(),
            no_3d: false,
        }))
        .with_anchor(anchor)
    }

    fn placed(objects: &[DrawingObject]) -> Vec<PlacedControl<'_>> {
        let mut sheet = crate::Worksheet::new("t");
        for object in objects {
            sheet.add_drawing(object.clone());
        }
        // Re-borrow from the slice we were given: rebuild placements
        // through a worksheet to exercise the real traversal.
        objects
            .iter()
            .enumerate()
            .filter_map(|(i, object)| {
                object.kind.as_form_control().map(|control| PlacedControl {
                    path: vec![i],
                    rect_emu: crate::drawing::anchor_rect_emu(&object.anchor),
                    control,
                })
            })
            .collect()
    }

    #[test]
    fn new_control_defaults() {
        let ctrl = FormControl::new(FormControlKind::Button {
            caption: "Run".to_string(),
        });
        assert_eq!(ctrl.caption(), Some("Run"));
        assert_eq!(ctrl.cell_link(), None);
    }

    #[test]
    fn cell_link_accessor() {
        let ctrl = FormControl::new(FormControlKind::Checkbox {
            caption: "On".to_string(),
            state: CheckState::Checked,
            cell_link: Some("$A$1".to_string()),
            no_3d: true,
        });
        assert_eq!(ctrl.cell_link(), Some("$A$1"));
        assert_eq!(ctrl.caption(), Some("On"));
    }

    #[test]
    fn radio_groups_by_enclosing_box() {
        let objects = vec![
            group_box("A", two_cell(0, 0, 2, 6)),
            group_box("B", two_cell(4, 0, 6, 6)),
            radio(two_cell(1, 1, 2, 2)), // in A
            radio(two_cell(5, 1, 6, 2)), // in B
            radio(two_cell(1, 3, 2, 4)), // in A
            radio(two_cell(8, 1, 9, 2)), // loose
        ];
        let controls = placed(&objects);
        let groups = radio_groups(&controls);
        assert_eq!(groups, vec![vec![2, 4], vec![3], vec![5]]);
    }

    #[test]
    fn radio_groups_nested_boxes_pick_innermost() {
        let objects = vec![
            group_box("outer", two_cell(0, 0, 10, 20)),
            group_box("inner", two_cell(1, 1, 4, 6)),
            radio(two_cell(2, 2, 3, 3)),   // inner box
            radio(two_cell(6, 10, 7, 11)), // outer box only
        ];
        let controls = placed(&objects);
        let groups = radio_groups(&controls);
        assert_eq!(groups, vec![vec![2], vec![3]]);
    }

    #[test]
    fn radio_groups_handles_extreme_anchor_offsets() {
        let objects = vec![
            group_box(
                "huge",
                DrawingAnchor::Absolute {
                    x_emu: i64::MIN,
                    y_emu: i64::MIN,
                    width_emu: i64::MAX,
                    height_emu: i64::MAX,
                },
            ),
            radio(DrawingAnchor::Absolute {
                x_emu: i64::MAX,
                y_emu: i64::MAX,
                width_emu: i64::MAX,
                height_emu: i64::MAX,
            }),
        ];
        let controls = placed(&objects);
        assert_eq!(radio_groups(&controls), vec![vec![1]]);
    }

    #[test]
    fn validates_kind_invariants() {
        let invalid_radio = FormControl::new(FormControlKind::OptionButton {
            caption: "radio".to_string(),
            state: CheckState::Mixed,
            cell_link: None,
            first_in_group: false,
            no_3d: false,
        });
        assert!(invalid_radio
            .validate()
            .unwrap_err()
            .to_string()
            .contains("mixed"));

        let invalid_list = FormControl::new(FormControlKind::ListBox {
            input_range: None,
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![2, 1],
            no_3d: false,
        });
        assert!(invalid_list
            .validate()
            .unwrap_err()
            .to_string()
            .contains("sorted and unique"));

        // Selection indices are zero-based: 0 selects the first item.
        let zero_selected = FormControl::new(FormControlKind::ListBox {
            input_range: None,
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 2],
            no_3d: false,
        });
        zero_selected.validate().unwrap();

        let overflowing_list = FormControl::new(FormControlKind::ListBox {
            input_range: None,
            cell_link: None,
            selection: ListSelection::Single,
            selected: vec![u16::MAX],
            no_3d: false,
        });
        assert!(overflowing_list
            .validate()
            .unwrap_err()
            .to_string()
            .contains("maximum"));

        let zero_dropdown = FormControl::new(FormControlKind::Dropdown {
            input_range: None,
            cell_link: None,
            selected: Some(0),
            lines: 8,
            no_3d: false,
        });
        zero_dropdown.validate().unwrap();

        let overflowing_dropdown = FormControl::new(FormControlKind::Dropdown {
            input_range: None,
            cell_link: None,
            selected: Some(u16::MAX),
            lines: 8,
            no_3d: false,
        });
        assert!(overflowing_dropdown
            .validate()
            .unwrap_err()
            .to_string()
            .contains("maximum"));

        let multi_selected_single = FormControl::new(FormControlKind::ListBox {
            input_range: None,
            cell_link: None,
            selection: ListSelection::Single,
            selected: vec![0, 1],
            no_3d: false,
        });
        assert!(multi_selected_single
            .validate()
            .unwrap_err()
            .to_string()
            .contains("at most one item"));

        let zero_lines = FormControl::new(FormControlKind::Dropdown {
            input_range: None,
            cell_link: None,
            selected: None,
            lines: 0,
            no_3d: false,
        });
        assert!(zero_lines
            .validate()
            .unwrap_err()
            .to_string()
            .contains("lines"));

        let blank_cell_link = FormControl::new(FormControlKind::Checkbox {
            caption: "blank".to_string(),
            state: CheckState::Checked,
            cell_link: Some("   ".to_string()),
            no_3d: false,
        });
        assert!(blank_cell_link
            .validate()
            .unwrap_err()
            .to_string()
            .contains("cell link cannot be empty"));
    }

    #[test]
    fn validates_numeric_control_invariants() {
        let scrollbar = |value, min, max, increment, page| {
            FormControl::new(FormControlKind::Scrollbar {
                value,
                min,
                max,
                increment,
                page,
                horizontal: false,
                cell_link: None,
            })
        };
        scrollbar(5, 0, 10, 1, 2).validate().unwrap();
        assert!(scrollbar(5, 6, 10, 1, 2)
            .validate()
            .unwrap_err()
            .to_string()
            .contains("min <= value <= max"));
        assert!(scrollbar(11, 0, 10, 1, 2)
            .validate()
            .unwrap_err()
            .to_string()
            .contains("min <= value <= max"));
        assert!(scrollbar(5, 10, 0, 1, 2)
            .validate()
            .unwrap_err()
            .to_string()
            .contains("min <= value <= max"));
        assert!(scrollbar(5, 0, 10, 0, 2)
            .validate()
            .unwrap_err()
            .to_string()
            .contains("increment"));
        assert!(scrollbar(5, 0, 10, 1, 0)
            .validate()
            .unwrap_err()
            .to_string()
            .contains("page"));

        let spinner = FormControl::new(FormControlKind::Spinner {
            value: 5,
            min: 0,
            max: 10,
            increment: 0,
            cell_link: None,
        });
        assert!(spinner
            .validate()
            .unwrap_err()
            .to_string()
            .contains("increment"));
    }

    #[test]
    fn worksheet_drawing_mutations_validate_and_shift_indices() {
        let mut worksheet = crate::Worksheet::new("Sheet1");
        let first = worksheet
            .try_add_drawing(DrawingObject::form_control(FormControl::new(
                FormControlKind::Button {
                    caption: "one".to_string(),
                },
            )))
            .unwrap();
        let second = worksheet
            .try_add_drawing(DrawingObject::form_control(FormControl::new(
                FormControlKind::Label {
                    caption: "two".to_string(),
                },
            )))
            .unwrap();
        assert_eq!((first, second), (0, 1));
        assert_eq!(worksheet.form_control_count(), 2);

        let removed = worksheet.remove_drawing(0).unwrap();
        assert_eq!(
            removed.kind.as_form_control().unwrap().caption(),
            Some("one")
        );
        let remaining: Vec<_> = worksheet
            .form_controls()
            .map(|drawn| drawn.payload.caption().unwrap().to_string())
            .collect();
        assert_eq!(remaining, vec!["two"]);
        assert!(worksheet.remove_drawing(1).is_err());
    }
}
