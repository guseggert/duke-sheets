//! Form controls (legacy Forms toolbar controls)
//!
//! Worksheet-embedded interactive controls: buttons, checkboxes,
//! option (radio) buttons, labels, group boxes, list boxes,
//! dropdowns (combo boxes), scrollbars, and spinners.
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
//! ## Example
//!
//! ```rust
//! use duke_sheets_core::{Workbook, FormControl, FormControlKind, CheckState};
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//!
//! sheet.add_form_control(FormControl::new(FormControlKind::Checkbox {
//!     caption: "Enable feature".to_string(),
//!     state: CheckState::Checked,
//!     cell_link: Some("$D$2".to_string()),
//!     no_3d: false,
//! }));
//!
//! assert_eq!(sheet.form_control_count(), 1);
//! ```

use duke_sheets_chart::DrawingAnchor;

use crate::{Error, Result, MAX_COLS, MAX_ROWS};

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
#[derive(Debug, Clone, PartialEq)]
pub struct FormControl {
    /// Shape name (e.g. "Check Box 1"). `None` lets the writer pick one.
    pub name: Option<String>,
    /// Placement of the control on the sheet.
    pub anchor: DrawingAnchor,
    /// Kind-specific properties.
    pub kind: FormControlKind,
    /// Whether the control is locked when the sheet is protected.
    pub locked: bool,
    /// Whether the control is included when the sheet is printed.
    pub printable: bool,
}

impl FormControl {
    /// Create a control of the given kind with a default anchor.
    pub fn new(kind: FormControlKind) -> Self {
        Self {
            name: None,
            anchor: DrawingAnchor::default(),
            kind,
            locked: true,
            printable: true,
        }
    }

    /// Create a control of the given kind at the given anchor.
    pub fn with_anchor(kind: FormControlKind, anchor: DrawingAnchor) -> Self {
        Self {
            anchor,
            ..Self::new(kind)
        }
    }

    /// Set the shape name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The control's caption text, for kinds that have one.
    pub fn caption(&self) -> Option<&str> {
        self.kind.caption()
    }

    /// The control's cell link formula, for kinds that support one.
    pub fn cell_link(&self) -> Option<&str> {
        self.kind.cell_link()
    }

    /// Validate the anchor and kind-specific properties.
    pub fn validate(&self) -> Result<()> {
        validate_anchor(&self.anchor)?;
        self.kind.validate()
    }
}

fn validate_anchor(anchor: &DrawingAnchor) -> Result<()> {
    let validate_marker = |marker: &duke_sheets_chart::CellMarker| -> Result<()> {
        if marker.row >= MAX_ROWS {
            return Err(Error::RowOutOfBounds(marker.row, MAX_ROWS - 1));
        }
        if marker.col >= MAX_COLS {
            return Err(Error::ColumnOutOfBounds(marker.col, MAX_COLS - 1));
        }
        Ok(())
    };
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            validate_marker(from)?;
            validate_marker(to)?;
            if (to.col, to.col_offset_emu) < (from.col, from.col_offset_emu)
                || (to.row, to.row_offset_emu) < (from.row, from.row_offset_emu)
            {
                return Err(Error::other("form control anchor endpoints are reversed"));
            }
        }
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            validate_marker(from)?;
            if *width_emu < 0 || *height_emu < 0 {
                return Err(Error::other(
                    "form control anchor dimensions cannot be negative",
                ));
            }
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            if *x_emu < 0 || *y_emu < 0 || *width_emu < 0 || *height_emu < 0 {
                return Err(Error::other(
                    "absolute form control anchor values cannot be negative",
                ));
            }
        }
    }
    Ok(())
}

/// Partition a sheet's option buttons into radio groups, mirroring
/// Excel's grouping semantics: each radio belongs to the innermost
/// group box whose rectangle contains the radio's center point, and
/// radios outside every box form the sheet-level group.
///
/// Returns groups of indices into `controls`, in insertion order
/// (both across groups and within each group). Non-radio controls
/// never appear in the result.
///
/// Containment is evaluated in EMU at Excel's default cell metrics
/// (609,600 EMU per column, 190,500 EMU per row), matching the
/// quantisation used by the binary anchor encodings; custom row and
/// column sizes that visually move a radio across a box edge are not
/// accounted for.
pub fn radio_groups(controls: &[FormControl]) -> Vec<Vec<usize>> {
    let box_rects: Vec<(i128, i128, i128, i128)> = controls
        .iter()
        .filter(|c| matches!(c.kind, FormControlKind::GroupBox { .. }))
        .map(|c| anchor_rect_emu(&c.anchor))
        .collect();

    let containing_box = |anchor: &DrawingAnchor| -> Option<usize> {
        let (x1, y1, x2, y2) = anchor_rect_emu(anchor);
        let (cx, cy) = (
            x1.saturating_add(x2) / 2,
            y1.saturating_add(y2) / 2,
        );
        box_rects
            .iter()
            .enumerate()
            .filter(|(_, (bx1, by1, bx2, by2))| {
                (*bx1..=*bx2).contains(&cx) && (*by1..=*by2).contains(&cy)
            })
            .min_by_key(|(_, (bx1, by1, bx2, by2))| {
                (bx2 - bx1)
                    .max(0)
                    .saturating_mul((by2 - by1).max(0))
            })
            .map(|(i, _)| i)
    };

    let mut groups: Vec<(Option<usize>, Vec<usize>)> = Vec::new();
    for (idx, control) in controls.iter().enumerate() {
        if !matches!(control.kind, FormControlKind::OptionButton { .. }) {
            continue;
        }
        let key = containing_box(&control.anchor);
        match groups.iter_mut().find(|(k, _)| *k == key) {
            Some((_, members)) => members.push(idx),
            None => groups.push((key, vec![idx])),
        }
    }
    groups.into_iter().map(|(_, members)| members).collect()
}

/// Absolute EMU rectangle (x1, y1, x2, y2) for a drawing anchor at
/// Excel's default cell metrics.
fn anchor_rect_emu(anchor: &DrawingAnchor) -> (i128, i128, i128, i128) {
    const COL_EMU: i128 = 609_600;
    const ROW_EMU: i128 = 190_500;
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => (
            from.col as i128 * COL_EMU + from.col_offset_emu as i128,
            from.row as i128 * ROW_EMU + from.row_offset_emu as i128,
            to.col as i128 * COL_EMU + to.col_offset_emu as i128,
            to.row as i128 * ROW_EMU + to.row_offset_emu as i128,
        ),
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let x1 = from.col as i128 * COL_EMU + from.col_offset_emu as i128;
            let y1 = from.row as i128 * ROW_EMU + from.row_offset_emu as i128;
            (
                x1,
                y1,
                x1 + (*width_emu).max(0) as i128,
                y1 + (*height_emu).max(0) as i128,
            )
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => (
            *x_emu as i128,
            *y_emu as i128,
            *x_emu as i128 + (*width_emu).max(0) as i128,
            *y_emu as i128 + (*height_emu).max(0) as i128,
        ),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_chart::CellMarker;

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

    fn radio(anchor: DrawingAnchor) -> FormControl {
        FormControl::with_anchor(
            FormControlKind::OptionButton {
                caption: String::new(),
                state: CheckState::Unchecked,
                cell_link: None,
                first_in_group: false,
                no_3d: false,
            },
            anchor,
        )
    }

    #[test]
    fn new_control_defaults() {
        let ctrl = FormControl::new(FormControlKind::Button {
            caption: "Run".to_string(),
        });
        assert!(ctrl.name.is_none());
        assert!(ctrl.locked);
        assert!(ctrl.printable);
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
    fn with_name_builder() {
        let ctrl = FormControl::new(FormControlKind::Label {
            caption: "Info".to_string(),
        })
        .with_name("Label 7");
        assert_eq!(ctrl.name.as_deref(), Some("Label 7"));
    }

    #[test]
    fn radio_groups_by_enclosing_box() {
        let controls = vec![
            FormControl::with_anchor(
                FormControlKind::GroupBox {
                    caption: "A".to_string(),
                    no_3d: false,
                },
                two_cell(0, 0, 2, 6),
            ),
            FormControl::with_anchor(
                FormControlKind::GroupBox {
                    caption: "B".to_string(),
                    no_3d: false,
                },
                two_cell(4, 0, 6, 6),
            ),
            radio(two_cell(1, 1, 2, 2)), // in A
            radio(two_cell(5, 1, 6, 2)), // in B
            radio(two_cell(1, 3, 2, 4)), // in A
            radio(two_cell(8, 1, 9, 2)), // loose
        ];
        let groups = radio_groups(&controls);
        assert_eq!(groups, vec![vec![2, 4], vec![3], vec![5]]);
    }

    #[test]
    fn radio_groups_nested_boxes_pick_innermost() {
        let controls = vec![
            FormControl::with_anchor(
                FormControlKind::GroupBox {
                    caption: "outer".to_string(),
                    no_3d: false,
                },
                two_cell(0, 0, 10, 20),
            ),
            FormControl::with_anchor(
                FormControlKind::GroupBox {
                    caption: "inner".to_string(),
                    no_3d: false,
                },
                two_cell(1, 1, 4, 6),
            ),
            radio(two_cell(2, 2, 3, 3)), // inner box
            radio(two_cell(6, 10, 7, 11)), // outer box only
        ];
        let groups = radio_groups(&controls);
        assert_eq!(groups, vec![vec![2], vec![3]]);
    }

    #[test]
    fn radio_groups_handles_extreme_anchor_offsets() {
        let controls = vec![
            FormControl::with_anchor(
                FormControlKind::GroupBox {
                    caption: "huge".to_string(),
                    no_3d: false,
                },
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
        assert_eq!(radio_groups(&controls), vec![vec![1]]);
    }

    #[test]
    fn validates_kind_and_anchor_invariants() {
        let invalid_radio = FormControl::new(FormControlKind::OptionButton {
            caption: "radio".to_string(),
            state: CheckState::Mixed,
            cell_link: None,
            first_in_group: false,
            no_3d: false,
        });
        assert!(invalid_radio.validate().unwrap_err().to_string().contains("mixed"));

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

        let invalid_anchor = FormControl::with_anchor(
            FormControlKind::Button {
                caption: "button".to_string(),
            },
            DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: 4,
                    col_offset_emu: 0,
                    row: 4,
                    row_offset_emu: 0,
                },
                to: CellMarker {
                    col: 2,
                    col_offset_emu: 0,
                    row: 2,
                    row_offset_emu: 0,
                },
                edit_as: None,
            },
        );
        assert!(invalid_anchor
            .validate()
            .unwrap_err()
            .to_string()
            .contains("reversed"));
    }

    #[test]
    fn worksheet_indexed_mutations_validate_and_shift_indices() {
        let mut worksheet = crate::Worksheet::new("Sheet1");
        let first = worksheet
            .try_add_form_control(FormControl::new(FormControlKind::Button {
                caption: "one".to_string(),
            }))
            .unwrap();
        let second = worksheet
            .try_add_form_control(FormControl::new(FormControlKind::Label {
                caption: "two".to_string(),
            }))
            .unwrap();
        assert_eq!((first, second), (0, 1));
        assert_eq!(worksheet.form_control_count(), 2);

        worksheet
            .set_form_control(
                0,
                FormControl::new(FormControlKind::Label {
                    caption: "replaced".to_string(),
                }),
            )
            .unwrap();
        assert_eq!(worksheet.form_control(0).unwrap().caption(), Some("replaced"));

        let removed = worksheet.remove_form_control(0).unwrap();
        assert_eq!(removed.caption(), Some("replaced"));
        assert_eq!(worksheet.form_control(0).unwrap().caption(), Some("two"));
        assert!(worksheet.remove_form_control(1).is_err());
    }
}
