//! Shared `Arbitrary` form-control specs for the roundtrip fuzz
//! targets. Controls are added through the permissive
//! `add_form_control`, so writer validation and clean-error paths get
//! exercised alongside the happy path; writers may reject a workbook
//! with an error, but must never panic, and anything they write must
//! read back.

use arbitrary::{Arbitrary, Unstructured};
use duke_sheets_chart::{CellMarker, DrawingAnchor};
use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection};

#[derive(Debug)]
pub struct FuzzFormControl {
    kind: FormControlKind,
    from_col: u16,
    from_row: u32,
    col_span: u16,
    row_span: u32,
}

impl<'a> Arbitrary<'a> for FuzzFormControl {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let caption = |u: &mut Unstructured<'a>| -> arbitrary::Result<String> {
            let len = u.int_in_range(0..=24)?;
            let mut s = String::with_capacity(len);
            for _ in 0..len {
                let c: u8 = u.arbitrary()?;
                s.push(if (32..127).contains(&c) {
                    c as char
                } else {
                    'x'
                });
            }
            Ok(s)
        };
        let cell_link = |u: &mut Unstructured<'a>| -> arbitrary::Result<Option<String>> {
            Ok(if u.arbitrary()? {
                Some(format!("$D${}", u.int_in_range(1..=40u32)?))
            } else {
                None
            })
        };
        // Input range plus its item count, so selections can be kept
        // mostly in range.
        let input_range = |u: &mut Unstructured<'a>| -> arbitrary::Result<(Option<String>, u16)> {
            Ok(if u.arbitrary()? {
                let items = u.int_in_range(1..=8u16)?;
                (Some(format!("$H$1:$H${items}")), items)
            } else {
                (None, 0)
            })
        };
        let state = |u: &mut Unstructured<'a>| -> arbitrary::Result<CheckState> {
            Ok(match u.int_in_range(0..=2u8)? {
                0 => CheckState::Unchecked,
                1 => CheckState::Checked,
                _ => CheckState::Mixed,
            })
        };
        // Zero-based selection index, occasionally past the item count
        // to drive writer rejection paths.
        let selected_index = |u: &mut Unstructured<'a>, items: u16| -> arbitrary::Result<u16> {
            u.int_in_range(0..=items.saturating_add(1))
        };

        let kind = match u.int_in_range(0..=8u8)? {
            0 => FormControlKind::Button {
                caption: caption(u)?,
            },
            1 => FormControlKind::Checkbox {
                caption: caption(u)?,
                state: state(u)?,
                cell_link: cell_link(u)?,
                no_3d: u.arbitrary()?,
            },
            2 => FormControlKind::OptionButton {
                caption: caption(u)?,
                state: state(u)?,
                cell_link: cell_link(u)?,
                first_in_group: u.arbitrary()?,
                no_3d: u.arbitrary()?,
            },
            3 => FormControlKind::Label {
                caption: caption(u)?,
            },
            4 => FormControlKind::GroupBox {
                caption: caption(u)?,
                no_3d: u.arbitrary()?,
            },
            5 => {
                let (range, items) = input_range(u)?;
                let selection = match u.int_in_range(0..=2u8)? {
                    0 => ListSelection::Single,
                    1 => ListSelection::Multi,
                    _ => ListSelection::Extend,
                };
                let n = u.int_in_range(0..=3usize)?;
                let mut selected = Vec::with_capacity(n);
                for _ in 0..n {
                    selected.push(selected_index(u, items)?);
                }
                selected.sort_unstable();
                selected.dedup();
                FormControlKind::ListBox {
                    input_range: range,
                    cell_link: cell_link(u)?,
                    selection,
                    selected,
                    no_3d: u.arbitrary()?,
                }
            }
            6 => {
                let (range, items) = input_range(u)?;
                FormControlKind::Dropdown {
                    input_range: range,
                    cell_link: cell_link(u)?,
                    selected: if u.arbitrary()? {
                        Some(selected_index(u, items)?)
                    } else {
                        None
                    },
                    lines: u.int_in_range(0..=20)?,
                    no_3d: u.arbitrary()?,
                }
            }
            7 => FormControlKind::Scrollbar {
                value: u.int_in_range(0..=30_000)?,
                min: u.int_in_range(0..=30_000)?,
                max: u.int_in_range(0..=30_000)?,
                increment: u.int_in_range(0..=100)?,
                page: u.int_in_range(0..=100)?,
                horizontal: u.arbitrary()?,
                cell_link: cell_link(u)?,
            },
            _ => FormControlKind::Spinner {
                value: u.int_in_range(0..=30_000)?,
                min: u.int_in_range(0..=30_000)?,
                max: u.int_in_range(0..=30_000)?,
                increment: u.int_in_range(0..=100)?,
                cell_link: cell_link(u)?,
            },
        };

        Ok(FuzzFormControl {
            kind,
            from_col: u.int_in_range(0..=20)?,
            from_row: u.int_in_range(0..=200)?,
            col_span: u.int_in_range(1..=6)?,
            row_span: u.int_in_range(1..=8)?,
        })
    }
}

impl FuzzFormControl {
    /// Materialize as a model control payload.
    pub fn to_control(&self) -> FormControl {
        FormControl::new(self.kind.clone())
    }

    /// The spec's anchor.
    pub fn anchor(&self) -> DrawingAnchor {
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: self.from_col,
                col_offset_emu: 0,
                row: self.from_row,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: self.from_col + self.col_span,
                col_offset_emu: 0,
                row: self.from_row + self.row_span,
                row_offset_emu: 0,
            },
            edit_as: None,
        }
    }
}
