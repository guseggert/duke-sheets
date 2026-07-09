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
        /// information; the XLS writer chains all option buttons on a
        /// sheet into a single group in insertion order and recomputes
        /// this flag.
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
        /// Cell receiving the one-based index of the selected item.
        cell_link: Option<String>,
        /// Selection behavior.
        selection: ListSelection,
        /// One-based indices of selected items. At most one entry for
        /// [`ListSelection::Single`]; any number for multi/extend.
        selected: Vec<u16>,
        /// Render without 3D shading.
        no_3d: bool,
    },
    /// Dropdown (combo box) populated from a cell range.
    Dropdown {
        /// Range providing the list items (A1-style, optional sheet prefix).
        input_range: Option<String>,
        /// Cell receiving the one-based index of the selected item.
        cell_link: Option<String>,
        /// One-based index of the selected item; `None` = no selection.
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
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
