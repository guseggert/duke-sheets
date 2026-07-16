//! Workbook type - the main document structure

use std::any::Any;
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicU64, Ordering};

use crate::cell::{CellAddress, CellError, CellValue};
use crate::drawing::DrawingPath;
use crate::error::{Error, Result};
use crate::form_control::{
    radio_groups, CheckState, FormControlInteractionResult, FormControlKind, ListSelection,
};
use crate::named_range::{NameScope, NamedRange, NamedRangeCollection};
use crate::protection::WorkbookProtection;
use crate::worksheet::{SheetVisibility, Worksheet};
use crate::MAX_SHEET_NAME_LEN;
use duke_sheets_chart::Chart;

static NEXT_WORKBOOK_NONCE: AtomicU64 = AtomicU64::new(1);

/// A slot in the workbook tab bar, referencing either a worksheet or a chartsheet by index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetSlot {
    Worksheet(usize),
    ChartSheet(usize),
}

/// A workbook (spreadsheet document)
///
/// A workbook contains one or more worksheets and global settings.
#[derive(Debug)]
pub struct Workbook {
    /// Worksheets in the workbook
    worksheets: Vec<Worksheet>,
    /// Chart sheets in the workbook
    chartsheets: Vec<ChartSheet>,
    /// Tab-bar ordering of worksheets and chartsheets.
    /// Empty means "use default order: all worksheets first, then all chartsheets".
    sheet_order: Vec<SheetSlot>,
    /// Workbook settings
    settings: WorkbookSettings,
    /// Workbook structure/window protection.
    workbook_protection: Option<WorkbookProtection>,
    /// Active sheet index
    active_sheet: usize,
    /// Named ranges (defined names)
    named_ranges: NamedRangeCollection,
    /// Theme color scheme parsed from the file's theme part, when any.
    theme_palette: Option<crate::style::ThemePalette>,
    /// Opaque calculation cache, populated and consumed by the calculation engine.
    /// Stored as type-erased `Box<dyn Any>` so the core crate needs no dependency
    /// on `duke-sheets-formula`.
    calc_cache: Option<Box<dyn Any + Send + Sync>>,
    /// Structural generation counter - incremented when sheets are added, removed,
    /// reordered, or renamed. The calculation engine uses this to detect stale caches.
    structural_generation: u64,
    /// Unique identity for roundtrip state lookup (not persisted).
    nonce: u64,
}

impl Workbook {
    /// Create a new empty workbook with one worksheet
    pub fn new() -> Self {
        let mut wb = Self {
            worksheets: Vec::new(),
            chartsheets: Vec::new(),
            sheet_order: Vec::new(),
            settings: WorkbookSettings::default(),
            workbook_protection: None,
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
            theme_palette: None,
            calc_cache: None,
            structural_generation: 0,
            nonce: NEXT_WORKBOOK_NONCE.fetch_add(1, Ordering::Relaxed),
        };
        wb.add_worksheet_with_name("Sheet1").unwrap();
        wb
    }

    /// Create an empty workbook with no worksheets
    pub fn empty() -> Self {
        Self {
            worksheets: Vec::new(),
            chartsheets: Vec::new(),
            sheet_order: Vec::new(),
            settings: WorkbookSettings::default(),
            workbook_protection: None,
            active_sheet: 0,
            named_ranges: NamedRangeCollection::new(),
            theme_palette: None,
            calc_cache: None,
            structural_generation: 0,
            nonce: NEXT_WORKBOOK_NONCE.fetch_add(1, Ordering::Relaxed),
        }
    }

    /// Get the number of worksheets
    pub fn sheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get the tab-bar order as a slice of SheetSlot entries.
    /// Empty means the default order (all worksheets first, then all chartsheets).
    pub fn sheet_order(&self) -> &[SheetSlot] {
        &self.sheet_order
    }

    /// Get a mutable reference to the tab-bar order.
    pub fn sheet_order_mut(&mut self) -> &mut Vec<SheetSlot> {
        &mut self.sheet_order
    }

    /// Total number of tabs (worksheets + chartsheets).
    pub fn total_sheet_count(&self) -> usize {
        self.worksheets.len() + self.chartsheets.len()
    }

    /// Check if the workbook has no worksheets and no chartsheets
    pub fn is_empty(&self) -> bool {
        self.worksheets.is_empty() && self.chartsheets.is_empty()
    }

    /// Get a worksheet by index
    pub fn worksheet(&self, index: usize) -> Option<&Worksheet> {
        self.worksheets.get(index)
    }

    /// Get a mutable worksheet by index
    pub fn worksheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.worksheets.get_mut(index)
    }

    /// Get a worksheet by name
    pub fn worksheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.worksheets.iter().find(|ws| ws.name() == name)
    }

    /// Get a mutable worksheet by name
    pub fn worksheet_by_name_mut(&mut self, name: &str) -> Option<&mut Worksheet> {
        self.worksheets.iter_mut().find(|ws| ws.name() == name)
    }

    /// Get the index of a worksheet by name
    pub fn sheet_index(&self, name: &str) -> Option<usize> {
        self.worksheets.iter().position(|ws| ws.name() == name)
    }

    /// Iterate over all worksheets
    pub fn worksheets(&self) -> impl Iterator<Item = &Worksheet> {
        self.worksheets.iter()
    }

    /// Iterate over all worksheets mutably
    pub fn worksheets_mut(&mut self) -> impl Iterator<Item = &mut Worksheet> {
        self.worksheets.iter_mut()
    }

    /// Apply an interactive checkbox/option-button state change and
    /// immediately synchronize the linked cells of the affected
    /// controls only (the target, plus its radio-group siblings for
    /// option buttons). Checkboxes and radio groups elsewhere that
    /// link to the same cell(s) are then driven from the new cell
    /// value, as in Excel, where every control sharing a linked cell
    /// follows it. Unrelated controls' linked cells are left alone,
    /// matching Excel, which never rewrites unrelated links on
    /// interaction; use [`Self::sync_form_control_links`] for a full
    /// projection.
    pub fn set_form_control_check_state(
        &mut self,
        sheet_index: usize,
        path: &[usize],
        state: CheckState,
    ) -> Result<FormControlInteractionResult> {
        let sheet = self.worksheets.get_mut(sheet_index).ok_or_else(|| {
            Error::other(format!("worksheet index {sheet_index} out of bounds"))
        })?;
        let (controls_changed, affected) = sheet.set_form_control_check_state(path, state)?;
        let (linked_cells_changed, touched_cells) =
            self.sync_form_control_links_scoped(false, Some((sheet_index, &affected)));
        let reconciled =
            self.reconcile_shared_link_controls(&touched_cells, sheet_index, &affected);
        Ok(FormControlInteractionResult {
            controls_changed: controls_changed + reconciled,
            linked_cells_changed,
        })
    }

    /// Synchronize form-control state into each control's linked cell.
    ///
    /// This mirrors Excel's runtime behavior: checkboxes write booleans
    /// (`#N/A` for mixed), single-select lists and dropdowns write a
    /// one-based item index, option-button groups write their one-based
    /// selected index, and scrollbars/spinners write their numeric value.
    /// Fresh unchecked/no-selection/multi-select links preserve a blank cell;
    /// an existing linked value is reset to `FALSE` or `0`. A linked cell that
    /// already holds the control's value is left untouched, so a formula
    /// driving the control survives, as in Excel; a disagreeing formula is
    /// replaced with the control's constant. Malformed, external-workbook,
    /// and unknown-sheet links are left unchanged. If multiple controls
    /// target one cell, the last control in worksheet order wins.
    ///
    /// Returns the number of distinct linked cells whose value changed.
    pub fn sync_form_control_links(&mut self) -> usize {
        self.sync_form_control_links_impl(false)
    }

    /// Calculation-entry variant of [`Self::sync_form_control_links`]:
    /// formula-holding linked cells are left for the engine to evaluate,
    /// after which [`Self::sync_form_controls_from_linked_cells`] lets the
    /// recalculated formulas drive the controls, mirroring Excel's live
    /// cell-to-control direction.
    #[doc(hidden)]
    pub fn sync_form_control_links_for_calculation(&mut self) -> usize {
        self.sync_form_control_links_impl(true)
    }

    fn sync_form_control_links_impl(&mut self, skip_formula_cells: bool) -> usize {
        self.sync_form_control_links_scoped(skip_formula_cells, None)
            .0
    }

    /// `scope` restricts which controls project into their linked
    /// cells: `(sheet index, participating drawing paths)`. Radio
    /// group values are still computed from the whole sheet so a
    /// scoped member sees its group's true selection.
    ///
    /// Returns the number of cells whose value changed plus every
    /// projected `(sheet, row, col)` target, including targets whose
    /// value already agreed.
    fn sync_form_control_links_scoped(
        &mut self,
        skip_formula_cells: bool,
        scope: Option<(usize, &[crate::DrawingPath])>,
    ) -> (usize, Vec<(usize, u32, u16)>) {
        let mut updates: BTreeMap<(usize, u32, u16), CellValue> = BTreeMap::new();

        for source_sheet in 0..self.worksheets.len() {
            if let Some((scope_sheet, _)) = scope {
                if source_sheet != scope_sheet {
                    continue;
                }
            }
            let controls = self.worksheets[source_sheet].form_controls().collect::<Vec<_>>();
            let mut radio_values = vec![None; controls.len()];
            for group in radio_groups(&controls) {
                let value = group
                    .iter()
                    .position(|&index| {
                        matches!(
                            controls[index].payload.kind,
                            FormControlKind::OptionButton {
                                state: CheckState::Checked,
                                ..
                            }
                        )
                    })
                    .map(|index| CellValue::Number(index as f64 + 1.0))
                    .unwrap_or(CellValue::Empty);
                for index in group {
                    radio_values[index] = Some(value.clone());
                }
            }

            for (index, placed) in controls.iter().enumerate() {
                if let Some((_, paths)) = scope {
                    if !paths.iter().any(|path| *path == placed.path) {
                        continue;
                    }
                }
                let (link, value) = match &placed.payload.kind {
                    FormControlKind::Checkbox {
                        state, cell_link, ..
                    } => {
                        let value = match state {
                            CheckState::Unchecked => CellValue::Boolean(false),
                            CheckState::Checked => CellValue::Boolean(true),
                            CheckState::Mixed => CellValue::Error(CellError::Na),
                        };
                        (cell_link.as_deref(), value)
                    }
                    FormControlKind::ListBox {
                        cell_link,
                        selection,
                        selected,
                        ..
                    } => {
                        let value = if *selection == ListSelection::Single {
                            selected
                                .first()
                                .map(|index| CellValue::Number(f64::from(*index) + 1.0))
                                .unwrap_or(CellValue::Empty)
                        } else {
                            CellValue::Empty
                        };
                        (cell_link.as_deref(), value)
                    }
                    FormControlKind::Dropdown {
                        cell_link,
                        selected,
                        ..
                    } => (
                        cell_link.as_deref(),
                        selected
                            .map(|index| CellValue::Number(f64::from(index) + 1.0))
                            .unwrap_or(CellValue::Empty),
                    ),
                    FormControlKind::Scrollbar {
                        cell_link, value, ..
                    }
                    | FormControlKind::Spinner {
                        cell_link, value, ..
                    } => (cell_link.as_deref(), CellValue::Number(f64::from(*value))),
                    FormControlKind::OptionButton { cell_link, .. } => (
                        cell_link.as_deref(),
                        radio_values[index].clone().unwrap_or(CellValue::Empty),
                    ),
                    FormControlKind::Button { .. }
                    | FormControlKind::Label { .. }
                    | FormControlKind::GroupBox { .. }
                    | FormControlKind::Unknown { .. } => continue,
                };

                if let Some((sheet, address)) =
                    link.and_then(|link| self.resolve_control_link(source_sheet, link))
                {
                    let has_formula =
                        self.worksheets[sheet].has_formula_at(address.row, address.col);
                    if skip_formula_cells && has_formula {
                        continue;
                    }
                    let existing = self.worksheets[sheet].get_value_at(address.row, address.col);
                    let fresh_blank = existing == CellValue::Empty && !has_formula;
                    let value = if matches!(
                        &placed.payload.kind,
                        FormControlKind::Checkbox {
                            state: CheckState::Unchecked,
                            ..
                        }
                    ) && fresh_blank
                    {
                        CellValue::Empty
                    } else if value == CellValue::Empty && !fresh_blank {
                        CellValue::Number(0.0)
                    } else {
                        value
                    };
                    updates.insert((sheet, address.row, address.col), value);
                }
            }
        }

        let touched: Vec<(usize, u32, u16)> = updates.keys().copied().collect();
        let mut updated = 0;
        for ((sheet, row, col), value) in updates {
            // Excel only writes the linked cell when the control state
            // disagrees with it, so an agreeing driving formula survives.
            if self.worksheets[sheet].get_value_at(row, col) == value {
                continue;
            }
            if self.worksheets[sheet]
                .set_cell_value_at(row, col, value)
                .is_ok()
            {
                updated += 1;
            }
        }
        (updated, touched)
    }

    /// Drive checkboxes and radio groups whose links resolve to one of
    /// `cells` from the cell's current value, skipping the interaction's
    /// own controls (`exclude` on `exclude_sheet`). Without this, a
    /// control sharing the target's linked cell keeps its stale state
    /// and the save-time full sync projects that state back over the
    /// user's change (last control in worksheet order wins).
    ///
    /// Returns the number of controls whose state changed.
    fn reconcile_shared_link_controls(
        &mut self,
        cells: &[(usize, u32, u16)],
        exclude_sheet: usize,
        exclude: &[DrawingPath],
    ) -> usize {
        if cells.is_empty() {
            return 0;
        }
        let mut planned: Vec<(usize, DrawingPath, CheckState)> = Vec::new();
        for source_sheet in 0..self.worksheets.len() {
            let controls = self.worksheets[source_sheet].form_controls().collect::<Vec<_>>();
            let excluded =
                |path: &DrawingPath| source_sheet == exclude_sheet && exclude.contains(path);
            let shared_cell_value = |link: Option<&str>| -> Option<CellValue> {
                let (sheet, address) = self.resolve_control_link(source_sheet, link?)?;
                cells
                    .contains(&(sheet, address.row, address.col))
                    .then(|| self.worksheets[sheet].get_value_at(address.row, address.col))
            };

            for group in radio_groups(&controls) {
                if group.iter().any(|&index| excluded(&controls[index].path)) {
                    continue;
                }
                let Some(value) = group
                    .iter()
                    .find_map(|&index| shared_cell_value(controls[index].payload.cell_link()))
                else {
                    continue;
                };
                let Some(selected) = shared_link_radio_index(&value, group.len()) else {
                    continue;
                };
                for (position, &index) in group.iter().enumerate() {
                    let state = if position + 1 == selected {
                        CheckState::Checked
                    } else {
                        CheckState::Unchecked
                    };
                    planned.push((source_sheet, controls[index].path.clone(), state));
                }
            }

            for placed in &controls {
                if !matches!(placed.payload.kind, FormControlKind::Checkbox { .. })
                    || excluded(&placed.path)
                {
                    continue;
                }
                let Some(value) = shared_cell_value(placed.payload.cell_link()) else {
                    continue;
                };
                // The same truthiness Excel applies when a cell drives a
                // checkbox; Empty means an unchecked fresh link.
                let state = match value {
                    CellValue::Boolean(true) => CheckState::Checked,
                    CellValue::Boolean(false) | CellValue::Empty => CheckState::Unchecked,
                    CellValue::Number(n) if n == 0.0 => CheckState::Unchecked,
                    CellValue::Number(_) | CellValue::String(_) | CellValue::RichText(_) => {
                        CheckState::Checked
                    }
                    CellValue::Error(CellError::Na) => CheckState::Mixed,
                    _ => continue,
                };
                planned.push((source_sheet, placed.path.clone(), state));
            }
        }

        let mut changed = 0;
        for (sheet, path, state) in planned {
            let Some(control) = self.worksheets[sheet].form_control_at_path_mut(&path) else {
                continue;
            };
            match &mut control.kind {
                FormControlKind::Checkbox { state: current, .. }
                | FormControlKind::OptionButton { state: current, .. } => {
                    if std::mem::replace(current, state) != state {
                        changed += 1;
                    }
                }
                _ => {}
            }
        }
        changed
    }

    /// Drive form-control state from linked cells that hold formulas,
    /// mirroring Excel's cell-to-control direction.
    ///
    /// A formula can only persist in a linked cell when the control was not
    /// changed after it (Excel replaces the formula with a constant on
    /// interaction), so for formula-holding links the cell's cached value
    /// wins: checkboxes treat zero as unchecked, any other number, text, or
    /// boolean truth as checked, and `#N/A` as mixed; scrollbars and spinners
    /// clamp the number into their min/max range; single-select lists,
    /// dropdowns, and option groups treat the number as a one-based index,
    /// clamping past-the-end values to the last item and `<= 0` as no
    /// selection. Non-formula links, other error values, and unevaluated
    /// formulas leave the control unchanged.
    ///
    /// Returns the number of controls whose state changed.
    pub fn sync_form_controls_from_linked_cells(&mut self) -> usize {
        enum Driven {
            Check(CheckState),
            ListSelection(Vec<u16>),
            DropSelection(Option<u16>),
            Value(u16),
        }

        let driving_value = |workbook: &Self, source_sheet: usize, link: &str| {
            let (sheet, address) = workbook.resolve_control_link(source_sheet, link)?;
            if !workbook.worksheets[sheet].has_formula_at(address.row, address.col) {
                return None;
            }
            Some(workbook.worksheets[sheet].get_value_at(address.row, address.col))
        };
        let as_number = |value: &CellValue| match value {
            CellValue::Number(n) => Some(*n),
            CellValue::Boolean(b) => Some(f64::from(*b as u8)),
            _ => None,
        };
        let as_index = |value: &CellValue, count: Option<u16>| -> Option<u16> {
            let index = as_number(value)?.trunc();
            if index <= 0.0 {
                return Some(0);
            }
            let index = index.min(f64::from(u16::MAX)) as u16;
            Some(match count {
                Some(count) if count > 0 => index.min(count),
                _ => index,
            })
        };

        let mut planned: Vec<(usize, DrawingPath, Driven)> = Vec::new();
        for source_sheet in 0..self.worksheets.len() {
            let controls = self.worksheets[source_sheet].form_controls().collect::<Vec<_>>();

            for group in radio_groups(&controls) {
                // Excel persists the group's link on the first radio.
                let Some(value) = group.iter().find_map(|&index| {
                    controls[index]
                        .payload
                        .cell_link()
                        .and_then(|link| driving_value(self, source_sheet, link))
                }) else {
                    continue;
                };
                let Some(selected) = as_index(&value, u16::try_from(group.len()).ok()) else {
                    continue;
                };
                for (position, &index) in group.iter().enumerate() {
                    let state = if position + 1 == selected as usize {
                        CheckState::Checked
                    } else {
                        CheckState::Unchecked
                    };
                    planned.push((source_sheet, controls[index].path.clone(), Driven::Check(state)));
                }
            }

            for placed in &controls {
                match &placed.payload.kind {
                    FormControlKind::Checkbox { cell_link, .. } => {
                        let Some(value) = cell_link
                            .as_deref()
                            .and_then(|link| driving_value(self, source_sheet, link))
                        else {
                            continue;
                        };
                        let state = match value {
                            CellValue::Boolean(true) => CheckState::Checked,
                            CellValue::Boolean(false) => CheckState::Unchecked,
                            CellValue::Number(n) if n == 0.0 => CheckState::Unchecked,
                            CellValue::Number(_) => CheckState::Checked,
                            CellValue::String(_) | CellValue::RichText(_) => CheckState::Checked,
                            CellValue::Error(CellError::Na) => CheckState::Mixed,
                            _ => continue,
                        };
                        planned.push((source_sheet, placed.path.clone(), Driven::Check(state)));
                    }
                    FormControlKind::ListBox {
                        cell_link,
                        selection: ListSelection::Single,
                        input_range,
                        ..
                    } => {
                        let Some(value) = cell_link
                            .as_deref()
                            .and_then(|link| driving_value(self, source_sheet, link))
                        else {
                            continue;
                        };
                        let Some(selected) = as_index(&value, input_range_rows(input_range))
                        else {
                            continue;
                        };
                        let selected = if selected == 0 {
                            Vec::new()
                        } else {
                            vec![selected - 1]
                        };
                        planned.push((source_sheet, placed.path.clone(), Driven::ListSelection(selected)));
                    }
                    FormControlKind::Dropdown {
                        cell_link,
                        input_range,
                        ..
                    } => {
                        let Some(value) = cell_link
                            .as_deref()
                            .and_then(|link| driving_value(self, source_sheet, link))
                        else {
                            continue;
                        };
                        let Some(selected) = as_index(&value, input_range_rows(input_range))
                        else {
                            continue;
                        };
                        let selected = (selected > 0).then(|| selected - 1);
                        planned.push((source_sheet, placed.path.clone(), Driven::DropSelection(selected)));
                    }
                    FormControlKind::Scrollbar {
                        cell_link,
                        min,
                        max,
                        ..
                    }
                    | FormControlKind::Spinner {
                        cell_link,
                        min,
                        max,
                        ..
                    } => {
                        let Some(value) = cell_link
                            .as_deref()
                            .and_then(|link| driving_value(self, source_sheet, link))
                        else {
                            continue;
                        };
                        let (Some(number), true) = (as_number(&value), min <= max) else {
                            continue;
                        };
                        let clamped = number
                            .trunc()
                            .clamp(f64::from(*min), f64::from(*max)) as u16;
                        planned.push((source_sheet, placed.path.clone(), Driven::Value(clamped)));
                    }
                    _ => {}
                }
            }
        }

        let mut changed = 0;
        for (sheet, path, driven) in planned {
            let Some(control) = self.worksheets[sheet].form_control_at_path_mut(&path) else {
                continue;
            };
            let did_change = match (&mut control.kind, driven) {
                (
                    FormControlKind::Checkbox { state, .. }
                    | FormControlKind::OptionButton { state, .. },
                    Driven::Check(new_state),
                ) => std::mem::replace(state, new_state) != new_state,
                (FormControlKind::ListBox { selected, .. }, Driven::ListSelection(new)) => {
                    std::mem::replace(selected, new.clone()) != new
                }
                (FormControlKind::Dropdown { selected, .. }, Driven::DropSelection(new)) => {
                    std::mem::replace(selected, new) != new
                }
                (
                    FormControlKind::Scrollbar { value, .. }
                    | FormControlKind::Spinner { value, .. },
                    Driven::Value(new),
                ) => std::mem::replace(value, new) != new,
                _ => false,
            };
            if did_change {
                changed += 1;
            }
        }
        changed
    }

    /// Create a serialization snapshot with form-control linked cells
    /// synchronized. Calculation caches are omitted; persisted workbook state
    /// and the roundtrip nonce are retained.
    #[doc(hidden)]
    pub fn synchronized_for_save(&self) -> Option<Self> {
        if !self.worksheets.iter().any(|sheet| {
            sheet
                .form_controls().collect::<Vec<_>>()
                .iter()
                .any(|placed| placed.payload.cell_link().is_some())
        }) {
            return None;
        }
        let mut snapshot = Self {
            worksheets: self.worksheets.clone(),
            chartsheets: self.chartsheets.clone(),
            sheet_order: self.sheet_order.clone(),
            settings: self.settings.clone(),
            active_sheet: self.active_sheet,
            named_ranges: self.named_ranges.clone(),
            theme_palette: self.theme_palette,
            workbook_protection: self.workbook_protection.clone(),
            calc_cache: None,
            structural_generation: self.structural_generation,
            nonce: self.nonce,
        };
        snapshot.sync_form_control_links();
        Some(snapshot)
    }

    fn resolve_control_link(
        &self,
        source_sheet: usize,
        link: &str,
    ) -> Option<(usize, CellAddress)> {
        // Names may refer to other names; Excel resolves the chain.
        self.resolve_control_link_depth(source_sheet, link, 4)
    }

    fn resolve_control_link_depth(
        &self,
        source_sheet: usize,
        link: &str,
        depth: u8,
    ) -> Option<(usize, CellAddress)> {
        let depth = depth.checked_sub(1)?;
        let link = link.trim().strip_prefix('=').unwrap_or(link.trim());
        let (sheet, address) = match link.rsplit_once('!') {
            Some((sheet, address)) => {
                let sheet = parse_link_sheet_name(sheet)?;
                let folded = sheet.to_lowercase();
                let index = self
                    .worksheets
                    .iter()
                    .position(|worksheet| worksheet.name().to_lowercase() == folded)?;
                (index, address)
            }
            None => (source_sheet, link),
        };
        if let Ok(address) = CellAddress::parse(address) {
            return Some((sheet, address));
        }
        // Excel also accepts a defined name as a cell link.
        let name = self.named_ranges.get(address, sheet)?;
        self.resolve_control_link_depth(sheet, name.expression(), depth)
    }

    /// Structural generation counter - incremented when sheets are added,
    /// removed, reordered, or renamed.
    pub fn structural_generation(&self) -> u64 {
        self.structural_generation
    }

    /// Unique nonce for roundtrip state lookup.
    pub fn nonce(&self) -> u64 {
        self.nonce
    }

    /// Take the calculation cache (moves it out of the workbook).
    pub fn take_calc_cache(&mut self) -> Option<Box<dyn Any + Send + Sync>> {
        self.calc_cache.take()
    }

    /// Store a calculation cache on the workbook.
    pub fn set_calc_cache(&mut self, cache: Box<dyn Any + Send + Sync>) {
        self.calc_cache = Some(cache);
    }

    /// Add a new worksheet with default name
    pub fn add_worksheet(&mut self) -> Result<usize> {
        let name = self.generate_sheet_name();
        self.add_worksheet_with_name(&name)
    }

    /// Add a new worksheet with specified name
    pub fn add_worksheet_with_name(&mut self, name: &str) -> Result<usize> {
        self.validate_sheet_name(name)?;

        let index = self.worksheets.len();
        let worksheet = Worksheet::new(name);
        self.worksheets.push(worksheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Add a new worksheet with the exact name from a file, skipping validation.
    ///
    /// Real-world files can have empty names or names >31 chars. Readers must
    /// preserve them as-is; validation only applies to user-facing APIs.
    pub fn add_worksheet_with_name_unchecked(&mut self, name: &str) -> usize {
        let index = self.worksheets.len();
        let worksheet = Worksheet::new(name);
        self.worksheets.push(worksheet);
        self.structural_generation += 1;
        index
    }

    /// Insert a worksheet at a specific index
    pub fn insert_worksheet(&mut self, index: usize, name: &str) -> Result<()> {
        if index > self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        self.validate_sheet_name(name)?;

        let worksheet = Worksheet::new(name);
        self.worksheets.insert(index, worksheet);

        // Update sheet_order: increment indices >= the insertion point
        if !self.sheet_order.is_empty() {
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx >= index {
                        *idx += 1;
                    }
                }
            }
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }

        // Adjust active sheet index if needed
        if self.active_sheet >= index && !self.worksheets.is_empty() {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }
        self.structural_generation += 1;
        Ok(())
    }

    /// Add an existing worksheet to the workbook
    pub fn add_existing_worksheet(&mut self, worksheet: Worksheet) -> Result<usize> {
        self.validate_sheet_name(worksheet.name())?;
        let index = self.worksheets.len();
        self.worksheets.push(worksheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::Worksheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Remove a worksheet by index
    pub fn remove_worksheet(&mut self, index: usize) -> Result<Worksheet> {
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        let worksheet = self.worksheets.remove(index);

        // Update sheet_order: remove the entry and decrement indices above it
        if !self.sheet_order.is_empty() {
            self.sheet_order
                .retain(|slot| *slot != SheetSlot::Worksheet(index));
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx > index {
                        *idx -= 1;
                    }
                }
            }
        }

        // Adjust active sheet index
        if !self.worksheets.is_empty() {
            if self.active_sheet >= self.worksheets.len() {
                self.active_sheet = self.worksheets.len() - 1;
            }
        } else {
            self.active_sheet = 0;
        }
        self.structural_generation += 1;
        Ok(worksheet)
    }

    /// Move a worksheet to a new position
    pub fn move_worksheet(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(from, self.worksheets.len()));
        }
        if to >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(to, self.worksheets.len()));
        }

        let worksheet = self.worksheets.remove(from);
        self.worksheets.insert(to, worksheet);

        // Update sheet_order indices to reflect the move
        if !self.sheet_order.is_empty() {
            for slot in &mut self.sheet_order {
                if let SheetSlot::Worksheet(ref mut idx) = slot {
                    if *idx == from {
                        *idx = to;
                    } else if from < to && *idx > from && *idx <= to {
                        *idx -= 1;
                    } else if from > to && *idx >= to && *idx < from {
                        *idx += 1;
                    }
                }
            }
        }

        // Adjust active sheet if needed
        if self.active_sheet == from {
            self.active_sheet = to;
        } else if from < self.active_sheet && to >= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_sub(1);
        } else if from > self.active_sheet && to <= self.active_sheet {
            self.active_sheet = self.active_sheet.saturating_add(1);
        }
        self.structural_generation += 1;
        Ok(())
    }

    /// Rename a worksheet
    pub fn rename_worksheet(&mut self, index: usize, new_name: &str) -> Result<()> {
        // Check index first
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }

        // Validate the new name (excluding current sheet from duplicate check)
        self.validate_sheet_name_excluding(new_name, Some(index))?;

        self.worksheets[index].set_name(new_name);
        self.structural_generation += 1;
        Ok(())
    }

    /// Get the active sheet index
    pub fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    /// Set the active sheet index
    pub fn set_active_sheet(&mut self, index: usize) -> Result<()> {
        if index >= self.worksheets.len() {
            return Err(Error::SheetOutOfBounds(index, self.worksheets.len()));
        }
        self.active_sheet = index;
        Ok(())
    }

    /// Get workbook settings
    pub fn settings(&self) -> &WorkbookSettings {
        &self.settings
    }

    /// Get mutable workbook settings
    pub fn settings_mut(&mut self) -> &mut WorkbookSettings {
        &mut self.settings
    }

    /// The workbook's theme color scheme: the file's `clrScheme` when
    /// one was read, otherwise the default Office palette.
    pub fn theme_palette(&self) -> crate::style::ThemePalette {
        self.theme_palette.unwrap_or_default()
    }

    /// Record the theme color scheme parsed from a file's theme part.
    pub fn set_theme_palette(&mut self, palette: crate::style::ThemePalette) {
        self.theme_palette = Some(palette);
    }

    /// Resolve a color to display RGB against this workbook's theme
    /// palette. [`Color::Auto`] resolves to `None`.
    ///
    /// [`Color::Auto`]: crate::style::Color::Auto
    pub fn resolve_color(&self, color: &crate::style::Color) -> Option<(u8, u8, u8)> {
        self.theme_palette().resolve(color)
    }

    /// Get workbook protection settings, honoring legacy `WorkbookSettings`
    /// aliases for structure/password while preserving the new windows flag.
    pub fn workbook_protection(&self) -> Option<WorkbookProtection> {
        if let Some(protection) = &self.workbook_protection {
            let mut protection = protection.clone();
            protection.structure = self.settings.protected;
            protection.password_hash = self.settings.password_hash;
            Some(protection)
        } else {
            if self.settings.protected || self.settings.password_hash.is_some() {
                Some(WorkbookProtection {
                    structure: self.settings.protected,
                    windows: false,
                    password_hash: self.settings.password_hash,
                })
            } else {
                None
            }
        }
    }

    /// Set workbook protection settings.
    ///
    /// This keeps the legacy `WorkbookSettings::protected` and
    /// `WorkbookSettings::password_hash` fields synchronized as aliases for
    /// structure protection.
    pub fn set_workbook_protection(&mut self, protection: Option<WorkbookProtection>) {
        if let Some(ref protection) = protection {
            self.settings.protected = protection.structure;
            self.settings.password_hash = protection.password_hash;
        } else {
            self.settings.protected = false;
            self.settings.password_hash = None;
        }
        self.workbook_protection = protection;
    }

    /// Define a new workbook-scoped named range
    ///
    /// # Example
    /// ```
    /// use duke_sheets_core::Workbook;
    ///
    /// let mut wb = Workbook::new();
    /// wb.define_name("TaxRate", "Sheet1!$B$1").unwrap();
    /// ```
    pub fn define_name(&mut self, name: &str, refers_to: &str) -> Result<()> {
        self.define_name_with_scope(name, refers_to, NameScope::Workbook)
    }

    /// Define a named range with a specific scope
    pub fn define_name_with_scope(
        &mut self,
        name: &str,
        refers_to: &str,
        scope: NameScope,
    ) -> Result<()> {
        let range = NamedRange::new(name, refers_to, scope);
        self.named_ranges
            .define(range)
            .map_err(Error::InvalidName)?;
        self.structural_generation += 1;
        Ok(())
    }

    /// Define a sheet-scoped named range
    pub fn define_name_for_sheet(
        &mut self,
        name: &str,
        refers_to: &str,
        sheet_index: usize,
    ) -> Result<()> {
        self.define_name_with_scope(name, refers_to, NameScope::Sheet(sheet_index))
    }

    /// Get a named range by name, following Excel's scoping rules
    ///
    /// Looks for sheet-scoped name first (for the given sheet), then workbook-scoped.
    pub fn get_named_range(&self, name: &str, current_sheet: usize) -> Option<&NamedRange> {
        self.named_ranges.get(name, current_sheet)
    }

    /// Remove a workbook-scoped named range
    pub fn remove_name(&mut self, name: &str) -> Option<NamedRange> {
        let result = self.named_ranges.remove(name, &NameScope::Workbook);
        if result.is_some() {
            self.structural_generation += 1;
        }
        result
    }

    /// Remove a sheet-scoped named range
    pub fn remove_name_from_sheet(&mut self, name: &str, sheet_index: usize) -> Option<NamedRange> {
        let result = self
            .named_ranges
            .remove(name, &NameScope::Sheet(sheet_index));
        if result.is_some() {
            self.structural_generation += 1;
        }
        result
    }

    /// Get the named range collection (read-only)
    pub fn named_ranges(&self) -> &NamedRangeCollection {
        &self.named_ranges
    }

    /// Get the named range collection (mutable)
    pub fn named_ranges_mut(&mut self) -> &mut NamedRangeCollection {
        &mut self.named_ranges
    }

    /// Validate a sheet name
    fn validate_sheet_name(&self, name: &str) -> Result<()> {
        self.validate_sheet_name_excluding(name, None)
    }

    /// Validate a sheet name, optionally excluding a sheet from duplicate check
    fn validate_sheet_name_excluding(
        &self,
        name: &str,
        exclude_index: Option<usize>,
    ) -> Result<()> {
        // Check length
        if name.is_empty() {
            return Err(Error::InvalidSheetName("Sheet name cannot be empty".into()));
        }
        if name.len() > MAX_SHEET_NAME_LEN {
            return Err(Error::InvalidSheetName(format!(
                "Sheet name too long (max {} characters)",
                MAX_SHEET_NAME_LEN
            )));
        }

        // Check for invalid characters
        const INVALID_CHARS: &[char] = &[':', '\\', '/', '?', '*', '[', ']'];
        for c in INVALID_CHARS {
            if name.contains(*c) {
                return Err(Error::InvalidSheetName(format!(
                    "Sheet name cannot contain '{}'",
                    c
                )));
            }
        }

        // Check for duplicate names against worksheets (case-insensitive)
        let name_lower = name.to_lowercase();
        for (i, ws) in self.worksheets.iter().enumerate() {
            if Some(i) != exclude_index && ws.name().to_lowercase() == name_lower {
                return Err(Error::DuplicateSheetName(name.into()));
            }
        }

        // Check for duplicate names against chartsheets (case-insensitive)
        for cs in &self.chartsheets {
            if cs.name.to_lowercase() == name_lower {
                return Err(Error::DuplicateSheetName(name.into()));
            }
        }

        Ok(())
    }

    /// Generate a unique sheet name
    fn generate_sheet_name(&self) -> String {
        let mut n = self.worksheets.len() + 1;
        loop {
            let name = format!("Sheet{}", n);
            if self.validate_sheet_name(&name).is_ok() {
                return name;
            }
            n += 1;
        }
    }

    /// Add a chart sheet to the workbook.
    ///
    /// Returns an error if the name is empty or duplicates an existing
    /// worksheet or chartsheet name.
    pub fn add_chartsheet(&mut self, sheet: ChartSheet) -> Result<usize> {
        self.validate_sheet_name(&sheet.name)?;
        let index = self.chartsheets.len();
        self.chartsheets.push(sheet);
        if !self.sheet_order.is_empty() {
            self.sheet_order.push(SheetSlot::ChartSheet(index));
        }
        self.structural_generation += 1;
        Ok(index)
    }

    /// Add a chart sheet without name validation (for the reader).
    ///
    /// Real-world files can have names that violate validation rules.
    /// Readers must preserve them as-is.
    pub fn add_chartsheet_unchecked(&mut self, sheet: ChartSheet) -> usize {
        let index = self.chartsheets.len();
        self.chartsheets.push(sheet);
        self.structural_generation += 1;
        index
    }

    /// Get all chart sheets.
    pub fn chartsheets(&self) -> &[ChartSheet] {
        &self.chartsheets
    }

    /// Get a chart sheet by index.
    pub fn chartsheet(&self, index: usize) -> Option<&ChartSheet> {
        self.chartsheets.get(index)
    }

    /// Get the number of chart sheets.
    pub fn chartsheet_count(&self) -> usize {
        self.chartsheets.len()
    }
}

/// One-based radio selection driven by a shared linked cell: numbers
/// and booleans coerce like Excel (truncate; `<= 0` or blank means no
/// selection; past-the-end clamps to the last member). `None` leaves
/// the group unchanged.
fn shared_link_radio_index(value: &CellValue, count: usize) -> Option<usize> {
    let number = match value {
        CellValue::Number(n) => *n,
        CellValue::Boolean(b) => f64::from(*b as u8),
        CellValue::Empty => 0.0,
        _ => return None,
    };
    let index = number.trunc();
    if index <= 0.0 {
        return Some(0);
    }
    Some((index as usize).min(count))
}

/// Row count of a list control's input range, when it parses as a
/// same-workbook range (optionally sheet-qualified).
fn input_range_rows(input_range: &Option<String>) -> Option<u16> {
    let range = input_range.as_deref()?.trim();
    let range = range.strip_prefix('=').unwrap_or(range);
    let range = match range.rsplit_once('!') {
        Some((_, address)) => address,
        None => range,
    };
    let range = crate::CellRange::parse(range).ok()?;
    u16::try_from(range.end.row.checked_sub(range.start.row)? + 1).ok()
}

fn parse_link_sheet_name(value: &str) -> Option<String> {
    let value = value.trim();
    if value.contains(['[', ']']) {
        return None;
    }
    if let Some(quoted) = value.strip_prefix('\'') {
        let quoted = quoted.strip_suffix('\'')?;
        let mut name = String::with_capacity(quoted.len());
        let mut chars = quoted.chars().peekable();
        while let Some(ch) = chars.next() {
            if ch == '\'' {
                if chars.next() != Some('\'') {
                    return None;
                }
                name.push('\'');
            } else {
                name.push(ch);
            }
        }
        (!name.is_empty()).then_some(name)
    } else if value.is_empty() || value.contains('\'') {
        None
    } else {
        Some(value.to_string())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// Workbook-level settings
#[derive(Debug, Clone)]
pub struct WorkbookSettings {
    /// Date system: false = 1900 (Windows), true = 1904 (Mac)
    pub date_1904: bool,
    /// Workbook is protected
    pub protected: bool,
    /// Password hash for protection (if protected)
    pub password_hash: Option<u16>,
    /// Calculate formulas on open
    pub calc_on_open: bool,
    /// Default theme name
    pub theme: Option<String>,
}

impl Default for WorkbookSettings {
    fn default() -> Self {
        Self {
            date_1904: false,
            protected: false,
            password_hash: None,
            calc_on_open: true,
            theme: None,
        }
    }
}

/// A chart sheet - a sheet that contains only a chart, no cell data.
#[derive(Debug, Clone)]
pub struct ChartSheet {
    /// Sheet name (as it appears on the tab)
    pub name: String,
    /// The chart displayed in this sheet
    pub chart: Chart,
    /// Sheet visibility
    pub visibility: SheetVisibility,
    /// Raw XML fragments for non-chart drawing anchors, preserved for roundtrip.
    #[doc(hidden)]
    pub raw_drawing_objects: Vec<Vec<u8>>,
    /// Relationships referenced by `raw_drawing_objects`, captured so
    /// rewrites keep them resolvable.
    #[doc(hidden)]
    pub raw_drawing_rels: Vec<crate::drawing::RawRel>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        CellError, CellValue, CheckState, DrawingObject, FormControl, FormControlKind,
        ListSelection,
    };

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet1");
    }

    #[test]
    fn workbook_settings_remain_protection_aliases() {
        let mut wb = Workbook::new();
        wb.set_workbook_protection(Some(WorkbookProtection {
            structure: true,
            windows: true,
            password_hash: Some(0x1111),
        }));

        wb.settings_mut().protected = false;
        wb.settings_mut().password_hash = Some(0x2222);

        let protection = wb.workbook_protection().unwrap();
        assert!(!protection.structure);
        assert!(protection.windows);
        assert_eq!(protection.password_hash, Some(0x2222));
    }

    #[test]
    fn test_add_worksheets() {
        let mut wb = Workbook::new();

        let idx = wb.add_worksheet().unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_count(), 2);

        let idx = wb.add_worksheet_with_name("Data").unwrap();
        assert_eq!(idx, 2);
        assert_eq!(wb.worksheet(2).unwrap().name(), "Data");
    }

    #[test]
    fn test_duplicate_name() {
        let mut wb = Workbook::new();

        // Case-insensitive duplicate check
        assert!(wb.add_worksheet_with_name("SHEET1").is_err());
        assert!(wb.add_worksheet_with_name("sheet1").is_err());
    }

    #[test]
    fn test_invalid_sheet_name() {
        let mut wb = Workbook::new();

        assert!(wb.add_worksheet_with_name("").is_err());
        assert!(wb.add_worksheet_with_name("Sheet/1").is_err());
        assert!(wb.add_worksheet_with_name("Sheet:1").is_err());
        assert!(wb.add_worksheet_with_name("Sheet[1]").is_err());

        // Too long
        let long_name = "A".repeat(MAX_SHEET_NAME_LEN + 1);
        assert!(wb.add_worksheet_with_name(&long_name).is_err());
    }

    #[test]
    fn test_move_worksheet() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("A").unwrap();
        wb.add_worksheet_with_name("B").unwrap();
        wb.add_worksheet_with_name("C").unwrap();

        // Move C to position 1
        wb.move_worksheet(3, 1).unwrap();

        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet1");
        assert_eq!(wb.worksheet(1).unwrap().name(), "C");
        assert_eq!(wb.worksheet(2).unwrap().name(), "A");
        assert_eq!(wb.worksheet(3).unwrap().name(), "B");
    }

    #[test]
    fn test_worksheet_by_name() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();

        assert!(wb.worksheet_by_name("Data").is_some());
        assert!(wb.worksheet_by_name("NonExistent").is_none());
    }

    #[test]
    fn test_insert_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        // wb starts with Sheet1 at worksheets[0], sheet_order is empty
        // Manually populate sheet_order to simulate a read workbook:
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        // Now insert at position 0
        wb.insert_worksheet(0, "Inserted").unwrap();
        // Expected: sheet_order should have [Worksheet(0), Worksheet(1)]
        // The original Worksheet(0) became Worksheet(1) because it shifted right
        // The new sheet at index 0 was pushed to the end of sheet_order
        assert_eq!(wb.sheet_order().len(), 2);
        assert!(wb.sheet_order().contains(&SheetSlot::Worksheet(0)));
        assert!(wb.sheet_order().contains(&SheetSlot::Worksheet(1)));
        // Verify the original slot was renumbered:
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(1)); // was 0, now 1
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(0)); // newly inserted, pushed to end
    }

    #[test]
    fn test_remove_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        // worksheets: [Sheet1(0), Sheet2(1)]
        // Manually set sheet_order:
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(1));
        // Remove worksheet 0 (Sheet1)
        wb.remove_worksheet(0).unwrap();
        // Expected: sheet_order should have [Worksheet(0)] - the old index 1 became 0
        assert_eq!(wb.sheet_order().len(), 1);
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(0));
        assert_eq!(wb.worksheet(0).unwrap().name(), "Sheet2");
    }

    #[test]
    fn test_move_worksheet_updates_sheet_order() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        wb.add_worksheet_with_name("Sheet3").unwrap();
        // worksheets: [Sheet1(0), Sheet2(1), Sheet3(2)]
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(1));
        wb.sheet_order_mut().push(SheetSlot::Worksheet(2));
        // Move worksheet from index 0 to index 2
        wb.move_worksheet(0, 2).unwrap();
        // After move: worksheets = [Sheet2(0), Sheet3(1), Sheet1(2)]
        // sheet_order indices should be remapped:
        // Old 0→new 2, old 1→new 0, old 2→new 1
        assert_eq!(wb.sheet_order()[0], SheetSlot::Worksheet(2)); // was 0, moved to 2
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(0)); // was 1, shifted to 0
        assert_eq!(wb.sheet_order()[2], SheetSlot::Worksheet(1)); // was 2, shifted to 1
    }

    #[test]
    fn test_add_worksheet_pushes_to_nonempty_sheet_order() {
        let mut wb = Workbook::new();
        // Simulate reading: manually populate sheet_order
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        // Now add a new worksheet via the API
        wb.add_worksheet_with_name("Sheet2").unwrap();
        // Since sheet_order was non-empty, the new sheet should have been appended
        assert_eq!(wb.sheet_order().len(), 2);
        assert_eq!(wb.sheet_order()[1], SheetSlot::Worksheet(1));
    }

    #[test]
    fn test_add_chartsheet_pushes_to_nonempty_sheet_order() {
        use duke_sheets_chart::ChartType;

        let mut wb = Workbook::new();
        wb.sheet_order_mut().push(SheetSlot::Worksheet(0));
        let cs = ChartSheet {
            name: "Chart1".into(),
            chart: Chart::new(ChartType::Pie),
            visibility: SheetVisibility::Visible,
            raw_drawing_objects: Vec::new(),
            raw_drawing_rels: Vec::new(),
        };
        wb.add_chartsheet(cs).unwrap();
        assert_eq!(wb.sheet_order().len(), 2);
        assert_eq!(wb.sheet_order()[1], SheetSlot::ChartSheet(0));
    }

    #[test]
    fn sync_form_control_links_matches_excel_semantics() {
        let mut wb = Workbook::new();
        wb.rename_worksheet(0, "Controls").unwrap();
        wb.add_worksheet_with_name("Linked Data").unwrap();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_formula("A8", "=FALSE").unwrap();
        ws.set_cell_formula("A11", "=0").unwrap();
        ws.set_formula_with_cached_value_at(11, 0, "=1=2", CellValue::Boolean(false))
            .unwrap();
        ws.set_cell_value("A3", 99.0).unwrap();
        ws.set_cell_value("A10", true).unwrap();

        let kinds = [
            FormControlKind::Checkbox {
                caption: "mixed".into(),
                state: CheckState::Mixed,
                cell_link: Some("$A$1".into()),
                no_3d: false,
            },
            FormControlKind::ListBox {
                input_range: None,
                cell_link: Some("$A$2".into()),
                selection: ListSelection::Single,
                selected: vec![2],
                no_3d: false,
            },
            FormControlKind::ListBox {
                input_range: None,
                cell_link: Some("$A$3".into()),
                selection: ListSelection::Multi,
                selected: vec![0, 2, 3],
                no_3d: false,
            },
            FormControlKind::Dropdown {
                input_range: None,
                cell_link: Some("$A$4".into()),
                selected: None,
                lines: 8,
                no_3d: false,
            },
            FormControlKind::Scrollbar {
                value: 55,
                min: 0,
                max: 100,
                increment: 1,
                page: 10,
                horizontal: false,
                cell_link: Some("$A$5".into()),
            },
            FormControlKind::Spinner {
                value: 18,
                min: 0,
                max: 30,
                increment: 1,
                cell_link: Some("'Linked Data'!$A$1".into()),
            },
            FormControlKind::OptionButton {
                caption: "one".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$7".into()),
                first_in_group: false,
                no_3d: false,
            },
            FormControlKind::OptionButton {
                caption: "two".into(),
                state: CheckState::Checked,
                cell_link: Some("$A$7".into()),
                first_in_group: false,
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "formula overwrite".into(),
                state: CheckState::Checked,
                cell_link: Some("$A$8".into()),
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "fresh unchecked".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$9".into()),
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "changed unchecked".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$10".into()),
                no_3d: false,
            },
            FormControlKind::Dropdown {
                input_range: None,
                cell_link: Some("$A$11".into()),
                selected: None,
                lines: 8,
                no_3d: false,
            },
            FormControlKind::Checkbox {
                caption: "agreeing formula".into(),
                state: CheckState::Unchecked,
                cell_link: Some("$A$12".into()),
                no_3d: false,
            },
        ];
        for kind in kinds {
            ws.add_drawing(DrawingObject::form_control(FormControl::new(kind))).unwrap();
        }

        assert_eq!(wb.sync_form_control_links(), 9);
        let ws = wb.worksheet(0).unwrap();
        assert_eq!(ws.get_value("A1").unwrap(), CellValue::Error(CellError::Na));
        assert_eq!(ws.get_value("A2").unwrap(), CellValue::Number(3.0));
        assert_eq!(ws.get_value("A3").unwrap(), CellValue::Number(0.0));
        assert_eq!(ws.get_value("A4").unwrap(), CellValue::Empty);
        assert_eq!(ws.get_value("A5").unwrap(), CellValue::Number(55.0));
        assert_eq!(ws.get_value("A7").unwrap(), CellValue::Number(2.0));
        assert_eq!(ws.get_value("A8").unwrap(), CellValue::Boolean(true));
        assert_eq!(ws.get_value("A9").unwrap(), CellValue::Empty);
        assert_eq!(ws.get_value("A10").unwrap(), CellValue::Boolean(false));
        assert_eq!(ws.get_value("A11").unwrap(), CellValue::Number(0.0));
        assert!(!ws.has_formula_at(10, 0));
        assert!(!ws.has_formula_at(7, 0));
        // A formula whose cached value agrees with the control state
        // survives, matching Excel's drive-the-control pattern.
        assert_eq!(ws.get_value("A12").unwrap(), CellValue::Boolean(false));
        assert!(ws.has_formula_at(11, 0));
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("A1").unwrap(),
            CellValue::Number(18.0)
        );
    }

    #[test]
    fn sync_form_control_links_resolves_quoted_sheets_and_skips_unsupported_links() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("It's A! Sheet").unwrap();
        let ws = wb.worksheet_mut(0).unwrap();
        for link in [
            "='It''s A! Sheet'!$B$2",
            "SUM(A1:A3)",
            "[Other.xlsx]Sheet1!$A$1",
            "Missing!$A$1",
            "=''!$A$1",
            "='Bob's'!$A$1",
        ] {
            ws.add_drawing(DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
                caption: link.into(),
                state: CheckState::Checked,
                cell_link: Some(link.into()),
                no_3d: false,
            }))).unwrap();
        }

        assert_eq!(wb.sync_form_control_links(), 1);
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("B2").unwrap(),
            CellValue::Boolean(true)
        );
    }

    #[test]
    fn sync_form_control_links_resolves_defined_names() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Target", "Data!$B$2"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Alias", "=Target"));
        wb.named_ranges_mut().define_or_update(NamedRange::new(
            "Local",
            "$C$3",
            NameScope::Sheet(0),
        ));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("Wide", "Data!$A$1:$A$4"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("LoopA", "LoopB"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("LoopB", "LoopA"));

        let ws = wb.worksheet_mut(0).unwrap();
        for (link, value) in [
            ("Target", 11),
            ("Alias", 22),
            ("Local", 33),
            ("Wide", 44),
            ("LoopA", 55),
            ("NoSuchName", 66),
        ] {
            ws.add_drawing(DrawingObject::form_control(FormControl::new(FormControlKind::Spinner {
                value,
                min: 0,
                max: 100,
                increment: 1,
                cell_link: Some(link.into()),
            }))).unwrap();
        }

        assert_eq!(wb.sync_form_control_links(), 2);
        assert_eq!(
            wb.worksheet(1).unwrap().get_value("B2").unwrap(),
            CellValue::Number(22.0),
            "workbook-scoped name resolves; the aliased control targets it too"
        );
        assert_eq!(
            wb.worksheet(0).unwrap().get_value("C3").unwrap(),
            CellValue::Number(33.0),
            "sheet-scoped name resolves against the control's sheet"
        );
    }

    #[test]
    fn sync_form_controls_from_linked_cells_matches_excel_semantics() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let formula = [
            ("A1", CellValue::Number(5.0)),
            ("A2", CellValue::string("text")),
            ("A3", CellValue::Error(CellError::Na)),
            ("A4", CellValue::Error(CellError::Div0)),
            ("A5", CellValue::Number(150.0)),
            ("A6", CellValue::Number(99.0)),
            ("A7", CellValue::Number(0.0)),
            ("A8", CellValue::Number(2.0)),
        ];
        for (address, cached) in formula {
            let parsed = crate::CellAddress::parse(address).unwrap();
            ws.set_formula_with_cached_value_at(parsed.row, parsed.col, "=X1", cached)
                .unwrap();
        }
        // Constant, not a formula: must not drive the control.
        ws.set_cell_value("A9", 0.0).unwrap();

        let checkbox = |link: &str, state| FormControlKind::Checkbox {
            caption: link.into(),
            state,
            cell_link: Some(link.into()),
            no_3d: false,
        };
        let kinds = [
            checkbox("$A$1", CheckState::Unchecked),
            checkbox("$A$2", CheckState::Unchecked),
            checkbox("$A$3", CheckState::Unchecked),
            checkbox("$A$4", CheckState::Checked),
            FormControlKind::Scrollbar {
                value: 40,
                min: 5,
                max: 95,
                increment: 1,
                page: 10,
                horizontal: false,
                cell_link: Some("$A$5".into()),
            },
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$4".into()),
                cell_link: Some("$A$6".into()),
                selection: ListSelection::Single,
                selected: vec![0],
                no_3d: false,
            },
            FormControlKind::Dropdown {
                input_range: Some("$H$1:$H$4".into()),
                cell_link: Some("$A$7".into()),
                selected: Some(1),
                lines: 8,
                no_3d: false,
            },
            FormControlKind::OptionButton {
                caption: "one".into(),
                state: CheckState::Checked,
                cell_link: Some("$A$8".into()),
                first_in_group: false,
                no_3d: false,
            },
            FormControlKind::OptionButton {
                caption: "two".into(),
                state: CheckState::Unchecked,
                cell_link: None,
                first_in_group: false,
                no_3d: false,
            },
            checkbox("$A$9", CheckState::Checked),
        ];
        for kind in kinds {
            ws.add_drawing(DrawingObject::form_control(FormControl::new(kind))).unwrap();
        }

        assert_eq!(wb.sync_form_controls_from_linked_cells(), 8);
        let controls: Vec<_> = wb.worksheet(0).unwrap().form_controls().map(|d| d.payload).collect();
        let state = |index: usize| match &controls[index].kind {
            FormControlKind::Checkbox { state, .. }
            | FormControlKind::OptionButton { state, .. } => *state,
            other => panic!("expected stateful control, got {other:?}"),
        };
        assert_eq!(state(0), CheckState::Checked, "nonzero number checks");
        assert_eq!(state(1), CheckState::Checked, "text is truthy");
        assert_eq!(state(2), CheckState::Mixed, "#N/A is mixed");
        assert_eq!(state(3), CheckState::Checked, "other errors leave state");
        match &controls[4].kind {
            FormControlKind::Scrollbar { value, .. } => {
                assert_eq!(*value, 95, "value clamps to max");
            }
            other => panic!("expected Scrollbar, got {other:?}"),
        }
        match &controls[5].kind {
            FormControlKind::ListBox { selected, .. } => {
                assert_eq!(selected, &vec![3], "out-of-range clamps to last item");
            }
            other => panic!("expected ListBox, got {other:?}"),
        }
        match &controls[6].kind {
            FormControlKind::Dropdown { selected, .. } => {
                assert_eq!(*selected, None, "zero deselects");
            }
            other => panic!("expected Dropdown, got {other:?}"),
        }
        assert_eq!(state(7), CheckState::Unchecked, "group index 2 moves selection");
        assert_eq!(state(8), CheckState::Checked, "second radio becomes checked");
        assert_eq!(state(9), CheckState::Checked, "constant cells do not drive");
    }

    #[test]
    fn sync_form_control_links_uses_control_order_for_duplicate_targets() {
        let mut wb = Workbook::new();
        assert!(wb.synchronized_for_save().is_none());
        let ws = wb.worksheet_mut(0).unwrap();
        ws.add_drawing(DrawingObject::form_control(FormControl::new(FormControlKind::Spinner {
            value: 7,
            min: 0,
            max: 10,
            increment: 1,
            cell_link: Some("$A$1".into()),
        }))).unwrap();
        ws.add_drawing(DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "later".into(),
            state: CheckState::Checked,
            cell_link: Some("$A$1".into()),
            no_3d: false,
        }))).unwrap();

        assert_eq!(wb.sync_form_control_links(), 1);
        assert_eq!(
            wb.worksheet(0).unwrap().get_value("A1").unwrap(),
            CellValue::Boolean(true)
        );
    }
}
