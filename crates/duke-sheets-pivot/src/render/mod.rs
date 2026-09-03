use crate::aggregate::*;
use crate::compile::*;
use crate::prelude::*;
use crate::show_as::*;
use crate::snapshot::*;

mod writeback;
pub(crate) use writeback::*;

#[derive(Debug, Clone)]
pub(crate) struct RenderedPivot {
    pub(crate) cells: Vec<Vec<CellValue>>,
    pub(crate) range: CellRange,
    pub(crate) source_rows: usize,
    pub(crate) column_number_formats: Vec<Option<String>>,
    pub(crate) cell_number_formats: Vec<Vec<Option<String>>>,
    pub(crate) data_start_row: usize,
    pub(crate) row_outline_levels: Vec<u8>,
    pub(crate) column_outline_levels: Vec<u8>,
    pub(crate) row_hidden: Vec<bool>,
    pub(crate) column_hidden: Vec<bool>,
    pub(crate) row_collapsed: Vec<bool>,
    pub(crate) column_collapsed: Vec<bool>,
    pub(crate) row_page_break_offsets: Vec<u32>,
    pub(crate) merged_ranges: Vec<CellRange>,
}

impl RenderedPivot {
    pub(crate) fn cell_count(&self) -> usize {
        self.cells.iter().map(Vec::len).sum()
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RenderedCells {
    cells: Vec<Vec<CellValue>>,
    row_measure_indexes: Vec<Option<usize>>,
}

impl RenderedCells {
    pub(crate) fn new() -> Self {
        Self {
            cells: Vec::new(),
            row_measure_indexes: Vec::new(),
        }
    }

    pub(crate) fn from_cells(cells: Vec<Vec<CellValue>>) -> Self {
        let row_measure_indexes = vec![None; cells.len()];
        Self {
            cells,
            row_measure_indexes,
        }
    }

    pub(crate) fn push_row(&mut self, row: Vec<CellValue>) {
        self.cells.push(row);
        self.row_measure_indexes.push(None);
    }

    pub(crate) fn push_measure_row(&mut self, row: Vec<CellValue>, measure_index: usize) {
        self.cells.push(row);
        self.row_measure_indexes.push(Some(measure_index));
    }

    pub(crate) fn sync_unmeasured_rows(&mut self) {
        self.row_measure_indexes.resize(self.cells.len(), None);
    }

    pub(crate) fn prepend_unmeasured_rows(&mut self, count: usize) {
        if count == 0 {
            return;
        }
        let mut indexes = vec![None; count];
        indexes.append(&mut self.row_measure_indexes);
        self.row_measure_indexes = indexes;
    }
}

pub(crate) fn render_pivot(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Result<RenderedPivot> {
    let mut row_page_break_offsets = pivot_row_page_break_offsets(pivot, plan, aggregation);
    let mut rendered_cells = if values_on_rows(pivot, plan) {
        match (
            compact_row_layout(pivot, plan),
            plan.column_indexes.is_empty(),
        ) {
            (true, true) => render_compact_values_on_rows_without_column_fields(
                pivot,
                snapshot,
                plan,
                aggregation,
            ),
            (true, false) => {
                render_compact_values_on_rows_with_column_fields(pivot, snapshot, plan, aggregation)
            }
            (false, true) => {
                render_values_on_rows_without_column_fields(pivot, snapshot, plan, aggregation)
            }
            (false, false) => {
                render_values_on_rows_with_column_fields(pivot, snapshot, plan, aggregation)
            }
        }
    } else {
        RenderedCells::from_cells(
            match (
                compact_row_layout(pivot, plan),
                plan.column_indexes.is_empty(),
            ) {
                (true, true) => {
                    render_compact_without_column_fields(pivot, snapshot, plan, aggregation)
                }
                (true, false) => {
                    render_compact_with_column_fields(pivot, snapshot, plan, aggregation)
                }
                (false, true) => render_without_column_fields(pivot, snapshot, plan, aggregation),
                (false, false) => render_with_column_fields(pivot, snapshot, plan, aggregation),
            },
        )
    };
    let page_field_row_count = page_field_row_count(pivot, plan);
    if page_field_row_count > 0 {
        for offset in &mut row_page_break_offsets {
            *offset += page_field_row_count;
        }
    }
    prepend_page_fields(&mut rendered_cells.cells, pivot, snapshot, plan);
    rendered_cells.prepend_unmeasured_rows(page_field_row_count as usize);

    let width = rendered_cells
        .cells
        .iter()
        .map(Vec::len)
        .max()
        .unwrap_or(1)
        .max(1);
    for row in &mut rendered_cells.cells {
        row.resize(width, CellValue::Empty);
    }
    if rendered_cells.cells.is_empty() {
        rendered_cells.push_row(vec![CellValue::Empty; width]);
    }
    rendered_cells.sync_unmeasured_rows();
    let mut column_number_formats = pivot_column_number_formats(pivot, plan, aggregation);
    column_number_formats.resize(width, None);
    let cell_number_formats = pivot_cell_number_formats(pivot, plan, &rendered_cells);
    let mut row_outline_levels = pivot_row_outline_levels(pivot, plan, aggregation);
    row_outline_levels.truncate(rendered_cells.cells.len());
    row_outline_levels.resize(rendered_cells.cells.len(), 0);
    let mut column_outline_levels = pivot_column_outline_levels(pivot, plan, aggregation);
    column_outline_levels.truncate(width);
    column_outline_levels.resize(width, 0);
    let mut row_hidden = pivot_row_hidden_flags(pivot, plan, aggregation);
    row_hidden.truncate(rendered_cells.cells.len());
    row_hidden.resize(rendered_cells.cells.len(), false);
    let mut column_hidden = pivot_column_hidden_flags(pivot, plan, aggregation);
    column_hidden.truncate(width);
    column_hidden.resize(width, false);
    let mut row_collapsed = pivot_row_collapsed_flags(pivot, plan, aggregation);
    row_collapsed.truncate(rendered_cells.cells.len());
    row_collapsed.resize(rendered_cells.cells.len(), false);
    let mut column_collapsed = pivot_column_collapsed_flags(pivot, plan, aggregation);
    column_collapsed.truncate(width);
    column_collapsed.resize(width, false);
    let data_start_row = pivot_data_start_row(pivot, plan);

    let range = output_range(pivot.target, rendered_cells.cells.len(), width)?;
    let merged_ranges = pivot_item_label_merged_ranges(pivot, plan, &rendered_cells.cells);
    Ok(RenderedPivot {
        cells: rendered_cells.cells,
        range,
        source_rows: snapshot.row_count,
        column_number_formats,
        cell_number_formats,
        data_start_row,
        row_outline_levels,
        column_outline_levels,
        row_hidden,
        column_hidden,
        row_collapsed,
        column_collapsed,
        row_page_break_offsets,
        merged_ranges,
    })
}

pub(crate) fn compact_row_layout(pivot: &PivotTable, plan: &CompiledPivotPlan) -> bool {
    matches!(pivot.layout.kind, PivotLayoutKind::Compact) && plan.row_indexes.len() > 1
}

pub(crate) fn values_on_rows(pivot: &PivotTable, plan: &CompiledPivotPlan) -> bool {
    matches!(pivot.layout.values_axis, PivotValuesAxis::Rows) && plan.measures.len() > 1
}

pub(crate) fn pivot_item_label_merged_ranges(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    cells: &[Vec<CellValue>],
) -> Vec<CellRange> {
    if !pivot.layout.merge_item_labels || compact_row_layout(pivot, plan) {
        return Vec::new();
    }

    let label_width = plan.row_indexes.len();
    if label_width < 2 {
        return Vec::new();
    }

    let mut ranges = Vec::new();
    let mut row = pivot_data_start_row(pivot, plan);
    while row < cells.len() {
        for col in 0..(label_width - 1) {
            if !cells[row][col].is_empty() {
                let end_row = merged_item_label_end_row(cells, row, col, label_width);
                if end_row > row {
                    ranges.push(CellRange::from_indices(
                        pivot.target.row + row as u32,
                        pivot.target.col + col as u16,
                        pivot.target.row + end_row as u32,
                        pivot.target.col + col as u16,
                    ));
                }
            }
        }
        row += 1;
    }
    ranges
}

pub(crate) fn merged_item_label_end_row(
    cells: &[Vec<CellValue>],
    start_row: usize,
    col: usize,
    label_width: usize,
) -> usize {
    let mut row = start_row + 1;
    while row < cells.len()
        && cells[row][col].is_empty()
        && row_has_deeper_item_label(&cells[row], col + 1, label_width)
    {
        row += 1;
    }
    row - 1
}

pub(crate) fn row_has_deeper_item_label(
    row: &[CellValue],
    start_col: usize,
    label_width: usize,
) -> bool {
    row.iter()
        .take(label_width)
        .skip(start_col)
        .any(|cell| !cell.is_empty())
}

pub(crate) fn pivot_data_start_row(pivot: &PivotTable, plan: &CompiledPivotPlan) -> usize {
    (page_field_row_count(pivot, plan) + body_header_row_count(pivot)) as usize
}

pub(crate) fn pivot_row_outline_levels(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<u8> {
    let mut levels = Vec::new();
    if pivot.layout.show_field_headers {
        levels.push(0);
    }

    let compact = compact_row_layout(pivot, plan);
    let rows_per_body_item = rows_per_pivot_body_item(pivot, plan);
    let outlines_enabled = pivot.layout.show_expand_collapse && plan.row_indexes.len() > 1;

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        if compact {
            for position in compact_group_header_positions(row_key, previous_row_key) {
                push_outline_levels(
                    &mut levels,
                    1,
                    pivot_outline_level(outlines_enabled, position),
                );
            }
        } else {
            for position in row_group_start_positions(row_key, previous_row_key) {
                if row_subtotal_at_top(pivot, plan, position) {
                    push_outline_levels(
                        &mut levels,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        pivot_outline_level(outlines_enabled, position),
                    );
                }
            }
        }

        push_outline_levels(
            &mut levels,
            rows_per_body_item,
            pivot_outline_level(outlines_enabled, row_key.len().saturating_sub(1)),
        );

        let next_row_key = aggregation.row_order.get(row_index + 1);
        for position in row_group_end_positions(row_key, next_row_key) {
            if compact {
                if is_row_subtotal_position(row_key, position)
                    && row_subtotal_enabled(plan, position)
                {
                    push_outline_levels(
                        &mut levels,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        pivot_outline_level(outlines_enabled, position),
                    );
                }
            } else if !row_subtotal_at_top(pivot, plan, position) {
                push_outline_levels(
                    &mut levels,
                    rows_per_body_item
                        * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                    pivot_outline_level(outlines_enabled, position),
                );
            }
            if row_field_inserts_blank_row(plan, position) {
                levels.push(0);
            }
        }
    }

    if pivot.layout.show_row_grand_totals {
        push_outline_levels(&mut levels, rows_per_body_item, 0);
    }

    let page_rows = page_field_row_count(pivot, plan) as usize;
    if page_rows > 0 {
        let mut prefixed = vec![0; page_rows];
        prefixed.extend(levels);
        levels = prefixed;
    }
    levels
}

pub(crate) fn pivot_row_collapsed_flags(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<bool> {
    let mut collapsed = Vec::new();
    if pivot.layout.show_field_headers {
        collapsed.push(false);
    }

    let compact = compact_row_layout(pivot, plan);
    let rows_per_body_item = rows_per_pivot_body_item(pivot, plan);
    let collapse_enabled = pivot.layout.show_expand_collapse && plan.row_indexes.len() > 1;

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        if compact {
            for position in compact_group_header_positions(row_key, previous_row_key) {
                collapsed.push(row_item_collapsed(
                    plan,
                    row_key,
                    position,
                    collapse_enabled,
                ));
            }
        } else {
            for position in row_group_start_positions(row_key, previous_row_key) {
                if row_subtotal_at_top(pivot, plan, position) {
                    push_collapsed_flags(
                        &mut collapsed,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        row_item_collapsed(plan, row_key, position, collapse_enabled),
                    );
                }
            }
        }

        push_collapsed_flags(&mut collapsed, rows_per_body_item, false);

        let next_row_key = aggregation.row_order.get(row_index + 1);
        for position in row_group_end_positions(row_key, next_row_key) {
            if compact {
                if is_row_subtotal_position(row_key, position)
                    && row_subtotal_enabled(plan, position)
                {
                    push_collapsed_flags(
                        &mut collapsed,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        row_item_collapsed(plan, row_key, position, collapse_enabled),
                    );
                }
            } else if !row_subtotal_at_top(pivot, plan, position) {
                push_collapsed_flags(
                    &mut collapsed,
                    rows_per_body_item
                        * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                    row_item_collapsed(plan, row_key, position, collapse_enabled),
                );
            }
            if row_field_inserts_blank_row(plan, position) {
                collapsed.push(false);
            }
        }
    }

    if pivot.layout.show_row_grand_totals {
        push_collapsed_flags(&mut collapsed, rows_per_body_item, false);
    }

    let page_rows = page_field_row_count(pivot, plan) as usize;
    if page_rows > 0 {
        let mut prefixed = vec![false; page_rows];
        prefixed.extend(collapsed);
        collapsed = prefixed;
    }
    collapsed
}

pub(crate) fn pivot_row_hidden_flags(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<bool> {
    let mut hidden = Vec::new();
    if pivot.layout.show_field_headers {
        hidden.push(false);
    }

    let compact = compact_row_layout(pivot, plan);
    let rows_per_body_item = rows_per_pivot_body_item(pivot, plan);
    let collapse_enabled = pivot.layout.show_expand_collapse && plan.row_indexes.len() > 1;

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let collapsed_position = first_row_collapsed_position(plan, row_key, collapse_enabled);
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        if compact {
            for position in compact_group_header_positions(row_key, previous_row_key) {
                hidden.push(position_hidden_by_collapsed(collapsed_position, position));
            }
        } else {
            for position in row_group_start_positions(row_key, previous_row_key) {
                if row_subtotal_at_top(pivot, plan, position) {
                    push_hidden_flags(
                        &mut hidden,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        position_hidden_by_collapsed(collapsed_position, position),
                    );
                }
            }
        }

        push_hidden_flags(
            &mut hidden,
            rows_per_body_item,
            collapsed_position.is_some(),
        );

        let next_row_key = aggregation.row_order.get(row_index + 1);
        for position in row_group_end_positions(row_key, next_row_key) {
            if compact {
                if is_row_subtotal_position(row_key, position)
                    && row_subtotal_enabled(plan, position)
                {
                    push_hidden_flags(
                        &mut hidden,
                        rows_per_body_item
                            * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                        position_hidden_by_collapsed(collapsed_position, position),
                    );
                }
            } else if !row_subtotal_at_top(pivot, plan, position) {
                push_hidden_flags(
                    &mut hidden,
                    rows_per_body_item
                        * row_subtotal_rendered_count(plan, aggregation, row_key, position),
                    position_hidden_by_collapsed(collapsed_position, position),
                );
            }
            if row_field_inserts_blank_row(plan, position) {
                hidden.push(position_hidden_by_collapsed(collapsed_position, position));
            }
        }
    }

    if pivot.layout.show_row_grand_totals {
        push_hidden_flags(&mut hidden, rows_per_body_item, false);
    }

    let page_rows = page_field_row_count(pivot, plan) as usize;
    if page_rows > 0 {
        let mut prefixed = vec![false; page_rows];
        prefixed.extend(hidden);
        hidden = prefixed;
    }
    hidden
}

pub(crate) fn rows_per_pivot_body_item(pivot: &PivotTable, plan: &CompiledPivotPlan) -> usize {
    if values_on_rows(pivot, plan) {
        plan.measures.len()
    } else {
        1
    }
}

pub(crate) fn push_outline_levels(levels: &mut Vec<u8>, count: usize, level: u8) {
    levels.extend(std::iter::repeat(level).take(count));
}

pub(crate) fn push_collapsed_flags(flags: &mut Vec<bool>, count: usize, collapsed: bool) {
    flags.extend(std::iter::repeat(collapsed).take(count));
}

pub(crate) fn push_hidden_flags(flags: &mut Vec<bool>, count: usize, hidden: bool) {
    flags.extend(std::iter::repeat(hidden).take(count));
}

pub(crate) fn pivot_outline_level(enabled: bool, depth: usize) -> u8 {
    if enabled && depth > 0 {
        depth.min(7) as u8
    } else {
        0
    }
}

pub(crate) fn pivot_column_outline_levels(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<u8> {
    let label_width = pivot_label_column_width(pivot, plan);
    let mut levels = vec![0; label_width];
    if plan.column_indexes.len() < 2 {
        return levels;
    }

    let outlines_enabled = pivot.layout.show_expand_collapse;
    let repetitions = if values_on_rows(pivot, plan) {
        1
    } else {
        plan.measures.len()
    };
    for slot in column_render_slots(pivot, plan, aggregation) {
        let level = match slot {
            ColumnRenderSlot::Leaf(column_key) => {
                pivot_outline_level(outlines_enabled, column_key.len().saturating_sub(1))
            }
            ColumnRenderSlot::Subtotal { prefix, .. } => {
                pivot_outline_level(outlines_enabled, prefix.len().saturating_sub(1))
            }
            ColumnRenderSlot::GrandTotal => 0,
        };
        push_outline_levels(&mut levels, repetitions, level);
    }
    levels
}

pub(crate) fn pivot_column_collapsed_flags(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<bool> {
    let label_width = pivot_label_column_width(pivot, plan);
    let mut collapsed = vec![false; label_width];
    if plan.column_indexes.len() < 2 {
        return collapsed;
    }

    let collapse_enabled = pivot.layout.show_expand_collapse;
    let repetitions = if values_on_rows(pivot, plan) {
        1
    } else {
        plan.measures.len()
    };
    for slot in column_render_slots(pivot, plan, aggregation) {
        let flag = match &slot {
            ColumnRenderSlot::Leaf(_) | ColumnRenderSlot::GrandTotal => false,
            ColumnRenderSlot::Subtotal { prefix, .. } => column_prefix_collapsed(
                plan,
                prefix,
                prefix.len().saturating_sub(1),
                collapse_enabled,
            ),
        };
        push_collapsed_flags(&mut collapsed, repetitions, flag);
    }
    collapsed
}

pub(crate) fn pivot_column_hidden_flags(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<bool> {
    let label_width = pivot_label_column_width(pivot, plan);
    let mut hidden = vec![false; label_width];
    if plan.column_indexes.len() < 2 {
        return hidden;
    }

    let collapse_enabled = pivot.layout.show_expand_collapse;
    let repetitions = if values_on_rows(pivot, plan) {
        1
    } else {
        plan.measures.len()
    };
    for slot in column_render_slots(pivot, plan, aggregation) {
        let flag = match &slot {
            ColumnRenderSlot::Leaf(column_key) => {
                first_column_collapsed_position(plan, column_key, collapse_enabled).is_some()
            }
            ColumnRenderSlot::Subtotal { prefix, .. } => {
                let position = prefix.len().saturating_sub(1);
                position_hidden_by_collapsed(
                    first_column_prefix_collapsed_position(plan, prefix, collapse_enabled),
                    position,
                )
            }
            ColumnRenderSlot::GrandTotal => false,
        };
        push_hidden_flags(&mut hidden, repetitions, flag);
    }
    hidden
}

pub(crate) fn pivot_label_column_width(pivot: &PivotTable, plan: &CompiledPivotPlan) -> usize {
    if values_on_rows(pivot, plan) {
        if compact_row_layout(pivot, plan) {
            2
        } else {
            plan.row_indexes.len() + 1
        }
    } else if compact_row_layout(pivot, plan) {
        1
    } else {
        plan.row_indexes.len()
    }
}

pub(crate) fn page_field_row_count(pivot: &PivotTable, plan: &CompiledPivotPlan) -> u32 {
    if plan.page_fields.is_empty() {
        0
    } else {
        page_field_display_row_count(pivot, plan) as u32 + 1
    }
}

pub(crate) fn page_field_display_row_count(pivot: &PivotTable, plan: &CompiledPivotPlan) -> usize {
    let count = plan.page_fields.len();
    if count == 0 {
        return 0;
    }

    let wrap = pivot.layout.page_wrap as usize;
    if wrap == 0 {
        count
    } else if pivot.layout.page_over_then_down {
        (count + wrap - 1) / wrap
    } else {
        wrap.min(count)
    }
}

pub(crate) fn body_header_row_count(pivot: &PivotTable) -> u32 {
    u32::from(pivot.layout.show_field_headers)
}

pub(crate) fn pivot_row_page_break_offsets(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<u32> {
    if !plan.row_fields.iter().any(|field| field.insert_page_break) {
        return Vec::new();
    }

    let compact = compact_row_layout(pivot, plan);
    let rows_per_body_item = if values_on_rows(pivot, plan) {
        plan.measures.len() as u32
    } else {
        1
    };
    let mut offsets = Vec::new();
    let mut row_offset = body_header_row_count(pivot);
    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        if !compact {
            for position in row_group_start_positions(row_key, previous_row_key) {
                if row_subtotal_at_top(pivot, plan, position) {
                    row_offset += rows_per_body_item
                        * row_subtotal_rendered_count(plan, aggregation, row_key, position) as u32;
                }
            }
        }
        if compact {
            row_offset += compact_group_header_positions(row_key, previous_row_key).len() as u32;
        }
        row_offset += rows_per_body_item;

        let next_row_key = aggregation.row_order.get(row_index + 1);
        for position in row_group_end_positions(row_key, next_row_key) {
            if !row_subtotal_at_top(pivot, plan, position) {
                row_offset += rows_per_body_item
                    * row_subtotal_rendered_count(plan, aggregation, row_key, position) as u32;
            }
            if row_field_inserts_blank_row(plan, position) {
                row_offset += 1;
            }
            if row_field_inserts_page_break(plan, position) {
                offsets.push(row_offset - 1);
            }
        }
    }

    offsets
}

pub(crate) fn row_subtotal_rendered_count(
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    position: usize,
) -> usize {
    if !is_row_subtotal_position(row_key, position) || !row_subtotal_enabled(plan, position) {
        return 0;
    }
    let subtotal_count = row_subtotals_for_position(plan, position).len();
    if !plan.column_indexes.is_empty() {
        return subtotal_count;
    }

    if aggregation
        .row_subtotal_states(&row_key[..=position])
        .is_some()
    {
        subtotal_count
    } else {
        0
    }
}

pub(crate) fn pivot_column_number_formats(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Option<String>> {
    if values_on_rows(pivot, plan) {
        let label_width = if compact_row_layout(pivot, plan) {
            2
        } else {
            plan.row_indexes.len() + 1
        };
        let data_width = if plan.column_indexes.is_empty() {
            1
        } else {
            column_render_slots(pivot, plan, aggregation).len()
        };
        return vec![None; label_width + data_width];
    }

    let label_width = if compact_row_layout(pivot, plan) {
        1
    } else {
        plan.row_indexes.len()
    };
    let mut formats = vec![None; label_width];
    let measure_formats = plan
        .measures
        .iter()
        .map(|measure| measure.number_format.clone())
        .collect::<Vec<_>>();
    let repetitions = if plan.column_indexes.is_empty() {
        1
    } else {
        column_render_slots(pivot, plan, aggregation).len()
    };
    for _ in 0..repetitions {
        formats.extend(measure_formats.iter().cloned());
    }
    formats
}

pub(crate) fn pivot_cell_number_formats(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    rendered: &RenderedCells,
) -> Vec<Vec<Option<String>>> {
    if !values_on_rows(pivot, plan) {
        return Vec::new();
    }

    let data_start_col = if compact_row_layout(pivot, plan) {
        2
    } else {
        plan.row_indexes.len() + 1
    };
    rendered
        .cells
        .iter()
        .zip(rendered.row_measure_indexes.iter())
        .map(|(row, measure_index)| {
            let Some(format) = measure_index
                .and_then(|index| plan.measures.get(index))
                .and_then(|measure| measure.number_format.clone())
            else {
                return vec![None; row.len()];
            };
            let mut formats = vec![None; row.len()];
            for cell_format in formats.iter_mut().skip(data_start_col) {
                *cell_format = Some(format.clone());
            }
            formats
        })
        .collect()
}

pub(crate) fn render_without_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let mut header = row_field_header_cells(plan);
    header.extend(
        plan.measures
            .iter()
            .map(|measure| CellValue::string(measure.caption())),
    );
    if pivot.layout.show_field_headers {
        cells.push(header);
    }

    let empty_column_key = Vec::new();
    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let emitted_top_subtotal = append_row_top_subtotals_without_column_fields(
            &mut cells,
            pivot,
            snapshot,
            plan,
            aggregation,
            row_key,
            previous_row_key,
            &empty_column_key,
        );
        let label_previous_row_key = if emitted_top_subtotal {
            Some(row_key)
        } else {
            previous_row_key
        };
        let mut row = row_label_cells(pivot, snapshot, plan, row_key, label_previous_row_key);
        let key = GroupKey {
            rows: row_key.clone(),
            columns: empty_column_key.clone(),
        };
        let context = ShowAsContext {
            snapshot,
            plan,
            aggregation,
            row_key: Some(row_key),
            column_key: Some(&empty_column_key),
        };
        row.extend(finalize_states_with_context(
            aggregation.groups.get(&key),
            &plan.measures,
            aggregation.row_total_states(row_key),
            aggregation.column_total_states(&empty_column_key),
            aggregation.grand_total_states(),
            &context,
        ));
        cells.push(row);

        let next_row_key = aggregation.row_order.get(row_index + 1);
        append_row_subtotals_without_column_fields(
            &mut cells,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &empty_column_key,
            plan.row_indexes.len() + plan.measures.len(),
            pivot,
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = grand_total_label_row(plan.row_indexes.len(), &grand_total_caption(pivot));
        let context = ShowAsContext {
            snapshot,
            plan,
            aggregation,
            row_key: None,
            column_key: None,
        };
        row.extend(finalize_state_slice_with_context(
            aggregation.grand_total_states(),
            &plan.measures,
            aggregation.grand_total_states(),
            aggregation.grand_total_states(),
            aggregation.grand_total_states(),
            &context,
        ));
        cells.push(row);
    }

    cells
}

pub(crate) fn render_compact_without_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let mut header = vec![CellValue::string("Row Labels")];
    header.extend(
        plan.measures
            .iter()
            .map(|measure| CellValue::string(measure.caption())),
    );
    if pivot.layout.show_field_headers {
        cells.push(header);
    }

    let empty_column_key = Vec::new();
    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        append_compact_group_headers(
            &mut cells,
            snapshot,
            plan,
            row_key,
            previous_row_key,
            plan.measures.len(),
        );

        let mut row = vec![compact_leaf_label_cell(snapshot, plan, row_key)];
        let key = GroupKey {
            rows: row_key.clone(),
            columns: empty_column_key.clone(),
        };
        let context = ShowAsContext {
            snapshot,
            plan,
            aggregation,
            row_key: Some(row_key),
            column_key: Some(&empty_column_key),
        };
        row.extend(finalize_states_with_context(
            aggregation.groups.get(&key),
            &plan.measures,
            aggregation.row_total_states(row_key),
            aggregation.column_total_states(&empty_column_key),
            aggregation.grand_total_states(),
            &context,
        ));
        cells.push(row);

        let next_row_key = aggregation.row_order.get(row_index + 1);
        append_compact_row_subtotals_without_column_fields(
            &mut cells,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &empty_column_key,
            1 + plan.measures.len(),
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = vec![CellValue::string(grand_total_caption(pivot))];
        let context = ShowAsContext {
            snapshot,
            plan,
            aggregation,
            row_key: None,
            column_key: None,
        };
        row.extend(finalize_state_slice_with_context(
            aggregation.grand_total_states(),
            &plan.measures,
            aggregation.grand_total_states(),
            aggregation.grand_total_states(),
            aggregation.grand_total_states(),
            &context,
        ));
        cells.push(row);
    }

    cells
}

pub(crate) fn render_with_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let column_slots = column_render_slots(pivot, plan, aggregation);
    let mut header = row_field_header_cells(plan);

    for slot in &column_slots {
        for measure in &plan.measures {
            let caption = match slot {
                ColumnRenderSlot::GrandTotal => grand_total_measure_caption(
                    &grand_total_caption(pivot),
                    measure,
                    plan.measures.len(),
                ),
                _ => measure_column_caption(
                    &column_slot_label(snapshot, plan, slot),
                    measure,
                    plan.measures.len(),
                ),
            };
            header.push(CellValue::string(caption));
        }
    }
    if pivot.layout.show_field_headers {
        cells.push(header);
    }

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let emitted_top_subtotal = append_row_top_subtotals_with_column_fields(
            &mut cells,
            pivot,
            snapshot,
            plan,
            aggregation,
            row_key,
            previous_row_key,
            &column_slots,
        );
        let label_previous_row_key = if emitted_top_subtotal {
            Some(row_key)
        } else {
            previous_row_key
        };
        let mut row = row_label_cells(pivot, snapshot, plan, row_key, label_previous_row_key);
        for slot in &column_slots {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: column_context_key(slot),
            };
            row.extend(finalize_states_with_context_and_aggregate(
                leaf_row_slot_states(aggregation, row_key, slot),
                &plan.measures,
                aggregation.row_total_states(row_key),
                column_slot_total(aggregation, slot),
                aggregation.grand_total_states(),
                &context,
                column_slot_aggregate_override(slot),
            ));
        }
        cells.push(row);

        let next_row_key = aggregation.row_order.get(row_index + 1);
        append_row_subtotals_with_column_fields(
            &mut cells,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &column_slots,
            plan.row_indexes.len() + column_slots.len() * plan.measures.len(),
            pivot,
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = grand_total_label_row(plan.row_indexes.len(), &grand_total_caption(pivot));
        for slot in &column_slots {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: column_context_key(slot),
            };
            row.extend(finalize_states_with_context_and_aggregate(
                grand_row_slot_states(aggregation, slot),
                &plan.measures,
                Some(aggregation.grand_total_states()),
                column_slot_total(aggregation, slot),
                aggregation.grand_total_states(),
                &context,
                column_slot_aggregate_override(slot),
            ));
        }
        cells.push(row);
    }

    cells
}

pub(crate) fn render_compact_with_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let column_slots = column_render_slots(pivot, plan, aggregation);
    let mut header = vec![CellValue::string("Row Labels")];

    for slot in &column_slots {
        for measure in &plan.measures {
            let caption = match slot {
                ColumnRenderSlot::GrandTotal => grand_total_measure_caption(
                    &grand_total_caption(pivot),
                    measure,
                    plan.measures.len(),
                ),
                _ => measure_column_caption(
                    &column_slot_label(snapshot, plan, slot),
                    measure,
                    plan.measures.len(),
                ),
            };
            header.push(CellValue::string(caption));
        }
    }
    if pivot.layout.show_field_headers {
        cells.push(header);
    }

    let data_width = column_slots.len() * plan.measures.len();
    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        append_compact_group_headers(
            &mut cells,
            snapshot,
            plan,
            row_key,
            previous_row_key,
            data_width,
        );

        let mut row = vec![compact_leaf_label_cell(snapshot, plan, row_key)];
        for slot in &column_slots {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: column_context_key(slot),
            };
            row.extend(finalize_states_with_context_and_aggregate(
                leaf_row_slot_states(aggregation, row_key, slot),
                &plan.measures,
                aggregation.row_total_states(row_key),
                column_slot_total(aggregation, slot),
                aggregation.grand_total_states(),
                &context,
                column_slot_aggregate_override(slot),
            ));
        }
        cells.push(row);

        let next_row_key = aggregation.row_order.get(row_index + 1);
        append_compact_row_subtotals_with_column_fields(
            &mut cells,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &column_slots,
            1 + data_width,
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = vec![CellValue::string(grand_total_caption(pivot))];
        for slot in &column_slots {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: column_context_key(slot),
            };
            row.extend(finalize_states_with_context_and_aggregate(
                grand_row_slot_states(aggregation, slot),
                &plan.measures,
                Some(aggregation.grand_total_states()),
                column_slot_total(aggregation, slot),
                aggregation.grand_total_states(),
                &context,
                column_slot_aggregate_override(slot),
            ));
        }
        cells.push(row);
    }

    cells
}

pub(crate) fn render_values_on_rows_without_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> RenderedCells {
    let mut rendered = RenderedCells::new();
    let empty_column_key = Vec::new();
    let label_width = plan.row_indexes.len() + 1;
    let mut header = values_on_rows_header(pivot, plan);
    header.push(CellValue::string(grand_total_caption(pivot)));
    if pivot.layout.show_field_headers {
        rendered.push_row(header);
    }

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let emitted_top_subtotal = rendered
            .append_values_on_rows_row_top_subtotals_without_column_fields(
                pivot,
                snapshot,
                plan,
                aggregation,
                row_key,
                previous_row_key,
                &empty_column_key,
            );
        let label_previous_row_key = if emitted_top_subtotal {
            Some(row_key)
        } else {
            previous_row_key
        };
        let key = GroupKey {
            rows: row_key.clone(),
            columns: empty_column_key.clone(),
        };
        for measure_index in 0..plan.measures.len() {
            let mut row = values_on_rows_label_cells(
                pivot,
                snapshot,
                plan,
                row_key,
                label_previous_row_key,
                measure_index,
            );
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: Some(&empty_column_key),
            };
            row.push(finalize_measure_from_states(
                aggregation.groups.get(&key),
                &plan.measures,
                aggregation.row_total_states(row_key),
                aggregation.column_total_states(&empty_column_key),
                aggregation.grand_total_states(),
                &context,
                measure_index,
                None,
            ));
            rendered.push_measure_row(row, measure_index);
        }

        let next_row_key = aggregation.row_order.get(row_index + 1);
        rendered.append_values_on_rows_row_subtotals_without_column_fields(
            pivot,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &empty_column_key,
            label_width + 1,
        );
    }

    if pivot.layout.show_row_grand_totals {
        for measure_index in 0..plan.measures.len() {
            let mut row = values_on_rows_grand_total_label_row(pivot, plan, measure_index);
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: None,
            };
            row.push(finalize_measure_from_state_slice(
                aggregation.grand_total_states(),
                &plan.measures,
                aggregation.grand_total_states(),
                aggregation.grand_total_states(),
                aggregation.grand_total_states(),
                &context,
                measure_index,
                None,
            ));
            rendered.push_measure_row(row, measure_index);
        }
    }

    rendered
}

pub(crate) fn render_compact_values_on_rows_without_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> RenderedCells {
    let mut rendered = RenderedCells::new();
    let empty_column_key = Vec::new();
    let mut header = vec![
        CellValue::string("Row Labels"),
        CellValue::string(values_caption(pivot)),
        CellValue::string(grand_total_caption(pivot)),
    ];
    if pivot.layout.show_field_headers {
        rendered.push_row(std::mem::take(&mut header));
    }

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        append_compact_group_headers(
            &mut rendered.cells,
            snapshot,
            plan,
            row_key,
            previous_row_key,
            2,
        );
        rendered.sync_unmeasured_rows();

        let key = GroupKey {
            rows: row_key.clone(),
            columns: empty_column_key.clone(),
        };
        for measure_index in 0..plan.measures.len() {
            let mut row = vec![
                compact_leaf_label_cell(snapshot, plan, row_key),
                CellValue::string(plan.measures[measure_index].caption()),
            ];
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: Some(&empty_column_key),
            };
            row.push(finalize_measure_from_states(
                aggregation.groups.get(&key),
                &plan.measures,
                aggregation.row_total_states(row_key),
                aggregation.column_total_states(&empty_column_key),
                aggregation.grand_total_states(),
                &context,
                measure_index,
                None,
            ));
            rendered.push_measure_row(row, measure_index);
        }

        let next_row_key = aggregation.row_order.get(row_index + 1);
        rendered.append_compact_values_on_rows_row_subtotals_without_column_fields(
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &empty_column_key,
            3,
        );
    }

    if pivot.layout.show_row_grand_totals {
        for measure_index in 0..plan.measures.len() {
            let mut row = vec![
                CellValue::string(grand_total_caption(pivot)),
                CellValue::string(plan.measures[measure_index].caption()),
            ];
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: None,
            };
            row.push(finalize_measure_from_state_slice(
                aggregation.grand_total_states(),
                &plan.measures,
                aggregation.grand_total_states(),
                aggregation.grand_total_states(),
                aggregation.grand_total_states(),
                &context,
                measure_index,
                None,
            ));
            rendered.push_measure_row(row, measure_index);
        }
    }

    rendered
}

pub(crate) fn render_values_on_rows_with_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> RenderedCells {
    let mut rendered = RenderedCells::new();
    let column_slots = column_render_slots(pivot, plan, aggregation);
    let label_width = plan.row_indexes.len() + 1;
    let mut header = values_on_rows_header(pivot, plan);
    for slot in &column_slots {
        header.push(CellValue::string(column_slot_label(snapshot, plan, slot)));
    }
    if pivot.layout.show_field_headers {
        rendered.push_row(header);
    }

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let emitted_top_subtotal = rendered
            .append_values_on_rows_row_top_subtotals_with_column_fields(
                pivot,
                snapshot,
                plan,
                aggregation,
                row_key,
                previous_row_key,
                &column_slots,
            );
        let label_previous_row_key = if emitted_top_subtotal {
            Some(row_key)
        } else {
            previous_row_key
        };
        for measure_index in 0..plan.measures.len() {
            let mut row = values_on_rows_label_cells(
                pivot,
                snapshot,
                plan,
                row_key,
                label_previous_row_key,
                measure_index,
            );
            for slot in &column_slots {
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: Some(row_key),
                    column_key: column_context_key(slot),
                };
                row.push(finalize_measure_from_states(
                    leaf_row_slot_states(aggregation, row_key, slot),
                    &plan.measures,
                    aggregation.row_total_states(row_key),
                    column_slot_total(aggregation, slot),
                    aggregation.grand_total_states(),
                    &context,
                    measure_index,
                    column_slot_aggregate_override(slot),
                ));
            }
            rendered.push_measure_row(row, measure_index);
        }

        let next_row_key = aggregation.row_order.get(row_index + 1);
        rendered.append_values_on_rows_row_subtotals_with_column_fields(
            pivot,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &column_slots,
            label_width + column_slots.len(),
        );
    }

    if pivot.layout.show_row_grand_totals {
        for measure_index in 0..plan.measures.len() {
            let mut row = values_on_rows_grand_total_label_row(pivot, plan, measure_index);
            for slot in &column_slots {
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: None,
                    column_key: column_context_key(slot),
                };
                row.push(finalize_measure_from_states(
                    grand_row_slot_states(aggregation, slot),
                    &plan.measures,
                    Some(aggregation.grand_total_states()),
                    column_slot_total(aggregation, slot),
                    aggregation.grand_total_states(),
                    &context,
                    measure_index,
                    column_slot_aggregate_override(slot),
                ));
            }
            rendered.push_measure_row(row, measure_index);
        }
    }

    rendered
}

pub(crate) fn render_compact_values_on_rows_with_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> RenderedCells {
    let mut rendered = RenderedCells::new();
    let column_slots = column_render_slots(pivot, plan, aggregation);
    let mut header = vec![
        CellValue::string("Row Labels"),
        CellValue::string(values_caption(pivot)),
    ];
    for slot in &column_slots {
        header.push(CellValue::string(column_slot_label(snapshot, plan, slot)));
    }
    if pivot.layout.show_field_headers {
        rendered.push_row(header);
    }

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        append_compact_group_headers(
            &mut rendered.cells,
            snapshot,
            plan,
            row_key,
            previous_row_key,
            1 + column_slots.len(),
        );
        rendered.sync_unmeasured_rows();

        for measure_index in 0..plan.measures.len() {
            let mut row = vec![
                compact_leaf_label_cell(snapshot, plan, row_key),
                CellValue::string(plan.measures[measure_index].caption()),
            ];
            for slot in &column_slots {
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: Some(row_key),
                    column_key: column_context_key(slot),
                };
                row.push(finalize_measure_from_states(
                    leaf_row_slot_states(aggregation, row_key, slot),
                    &plan.measures,
                    aggregation.row_total_states(row_key),
                    column_slot_total(aggregation, slot),
                    aggregation.grand_total_states(),
                    &context,
                    measure_index,
                    column_slot_aggregate_override(slot),
                ));
            }
            rendered.push_measure_row(row, measure_index);
        }

        let next_row_key = aggregation.row_order.get(row_index + 1);
        rendered.append_compact_values_on_rows_row_subtotals_with_column_fields(
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &column_slots,
            2 + column_slots.len(),
        );
    }

    if pivot.layout.show_row_grand_totals {
        for measure_index in 0..plan.measures.len() {
            let mut row = vec![
                CellValue::string(grand_total_caption(pivot)),
                CellValue::string(plan.measures[measure_index].caption()),
            ];
            for slot in &column_slots {
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: None,
                    column_key: column_context_key(slot),
                };
                row.push(finalize_measure_from_states(
                    grand_row_slot_states(aggregation, slot),
                    &plan.measures,
                    Some(aggregation.grand_total_states()),
                    column_slot_total(aggregation, slot),
                    aggregation.grand_total_states(),
                    &context,
                    measure_index,
                    column_slot_aggregate_override(slot),
                ));
            }
            rendered.push_measure_row(row, measure_index);
        }
    }

    rendered
}

pub(crate) fn values_on_rows_header(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
) -> Vec<CellValue> {
    let mut header = row_field_header_cells(plan);
    header.insert(
        values_axis_row_position(pivot, plan),
        CellValue::string(values_caption(pivot)),
    );
    header
}

pub(crate) fn row_field_header_cells(plan: &CompiledPivotPlan) -> Vec<CellValue> {
    plan.row_fields
        .iter()
        .map(|field| CellValue::string(pivot_field_caption(field)))
        .collect()
}

pub(crate) fn pivot_field_caption(field: &PivotField) -> &str {
    field.caption.as_deref().unwrap_or(&field.field.name)
}

pub(crate) fn values_on_rows_label_cells(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
    measure_index: usize,
) -> Vec<CellValue> {
    let mut row = row_label_cells(pivot, snapshot, plan, row_key, previous_row_key);
    row.insert(
        values_axis_row_position(pivot, plan),
        CellValue::string(plan.measures[measure_index].caption()),
    );
    row
}

pub(crate) fn values_on_rows_grand_total_label_row(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    measure_index: usize,
) -> Vec<CellValue> {
    let label_width = plan.row_indexes.len() + 1;
    let mut row = vec![CellValue::Empty; label_width];
    let values_position = values_axis_row_position(pivot, plan);
    if let Some(cell) = row.get_mut(values_position) {
        *cell = CellValue::string(plan.measures[measure_index].caption());
    }
    if label_width > 1 {
        let total_position = if values_position == 0 { 1 } else { 0 };
        row[total_position] = CellValue::string(grand_total_caption(pivot));
    }
    row
}

pub(crate) fn values_axis_row_position(pivot: &PivotTable, plan: &CompiledPivotPlan) -> usize {
    pivot
        .layout
        .values_axis_position
        .map(|position| position as usize)
        .unwrap_or(plan.row_indexes.len())
        .min(plan.row_indexes.len())
}

pub(crate) fn values_on_rows_subtotal_label_cells(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    prefix: &[u32],
    measure_index: usize,
    subtotal: PivotSubtotal,
) -> Vec<CellValue> {
    let mut row = row_subtotal_label_cells(snapshot, plan, prefix, subtotal);
    row.insert(
        values_axis_row_position(pivot, plan),
        CellValue::string(plan.measures[measure_index].caption()),
    );
    row
}

impl RenderedCells {
    fn append_values_on_rows_row_top_subtotals_without_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        previous_row_key: Option<&Vec<u32>>,
        empty_column_key: &Vec<u32>,
    ) -> bool {
        let mut emitted = false;
        for position in row_group_start_positions(row_key, previous_row_key) {
            if !row_subtotal_at_top(pivot, plan, position) {
                continue;
            }

            let prefix = row_key[..=position].to_vec();
            if let Some(states) = aggregation.row_subtotal_states(&prefix) {
                self.append_values_on_rows_subtotal_rows_without_column_fields(
                    pivot,
                    snapshot,
                    plan,
                    aggregation,
                    &prefix,
                    states,
                    position,
                    empty_column_key,
                );
                emitted = true;
            }
        }
        emitted
    }

    fn append_values_on_rows_row_top_subtotals_with_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        previous_row_key: Option<&Vec<u32>>,
        column_slots: &[ColumnRenderSlot],
    ) -> bool {
        let mut emitted = false;
        for position in row_group_start_positions(row_key, previous_row_key) {
            if !row_subtotal_at_top(pivot, plan, position) {
                continue;
            }

            let prefix = row_key[..=position].to_vec();
            self.append_values_on_rows_subtotal_rows_with_column_fields(
                pivot,
                snapshot,
                plan,
                aggregation,
                &prefix,
                position,
                column_slots,
            );
            emitted = true;
        }
        emitted
    }

    fn append_values_on_rows_row_subtotals_without_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        next_row_key: Option<&Vec<u32>>,
        empty_column_key: &Vec<u32>,
        row_width: usize,
    ) {
        for position in row_group_end_positions(row_key, next_row_key) {
            if is_row_subtotal_position(row_key, position)
                && row_subtotal_enabled(plan, position)
                && !row_subtotal_at_top(pivot, plan, position)
            {
                let prefix = row_key[..=position].to_vec();
                if let Some(states) = aggregation.row_subtotal_states(&prefix) {
                    self.append_values_on_rows_subtotal_rows_without_column_fields(
                        pivot,
                        snapshot,
                        plan,
                        aggregation,
                        &prefix,
                        states,
                        position,
                        empty_column_key,
                    );
                }
            }
            append_blank_row_after_ended_row_field(&mut self.cells, plan, position, row_width);
            self.sync_unmeasured_rows();
        }
    }

    fn append_values_on_rows_row_subtotals_with_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        next_row_key: Option<&Vec<u32>>,
        column_slots: &[ColumnRenderSlot],
        row_width: usize,
    ) {
        for position in row_group_end_positions(row_key, next_row_key) {
            if is_row_subtotal_position(row_key, position)
                && row_subtotal_enabled(plan, position)
                && !row_subtotal_at_top(pivot, plan, position)
            {
                let prefix = row_key[..=position].to_vec();
                self.append_values_on_rows_subtotal_rows_with_column_fields(
                    pivot,
                    snapshot,
                    plan,
                    aggregation,
                    &prefix,
                    position,
                    column_slots,
                );
            }
            append_blank_row_after_ended_row_field(&mut self.cells, plan, position, row_width);
            self.sync_unmeasured_rows();
        }
    }

    fn append_values_on_rows_subtotal_rows_without_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        prefix: &[u32],
        states: &[AggregateState],
        position: usize,
        empty_column_key: &Vec<u32>,
    ) {
        for subtotal in row_subtotals_for_position(plan, position) {
            let aggregate_override = subtotal_aggregate_for_field(subtotal);
            for measure_index in 0..plan.measures.len() {
                let mut row = values_on_rows_subtotal_label_cells(
                    pivot,
                    snapshot,
                    plan,
                    prefix,
                    measure_index,
                    subtotal,
                );
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: None,
                    column_key: Some(empty_column_key),
                };
                row.push(finalize_measure_from_state_slice(
                    states,
                    &plan.measures,
                    states,
                    aggregation
                        .column_total_states(empty_column_key)
                        .map(Vec::as_slice)
                        .unwrap_or(&[]),
                    aggregation.grand_total_states(),
                    &context,
                    measure_index,
                    aggregate_override,
                ));
                self.push_measure_row(row, measure_index);
            }
        }
    }

    fn append_values_on_rows_subtotal_rows_with_column_fields(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        prefix: &[u32],
        position: usize,
        column_slots: &[ColumnRenderSlot],
    ) {
        let row_total = aggregation.row_subtotal_states(prefix);
        for subtotal in row_subtotals_for_position(plan, position) {
            let row_aggregate_override = subtotal_aggregate_for_field(subtotal);
            for measure_index in 0..plan.measures.len() {
                let mut row = values_on_rows_subtotal_label_cells(
                    pivot,
                    snapshot,
                    plan,
                    prefix,
                    measure_index,
                    subtotal,
                );
                for slot in column_slots {
                    let context = ShowAsContext {
                        snapshot,
                        plan,
                        aggregation,
                        row_key: None,
                        column_key: column_context_key(slot),
                    };
                    row.push(finalize_measure_from_states(
                        subtotal_row_slot_states(aggregation, prefix, slot),
                        &plan.measures,
                        row_total,
                        column_slot_total(aggregation, slot),
                        aggregation.grand_total_states(),
                        &context,
                        measure_index,
                        row_aggregate_override.or_else(|| column_slot_aggregate_override(slot)),
                    ));
                }
                self.push_measure_row(row, measure_index);
            }
        }
    }

    fn append_compact_values_on_rows_row_subtotals_without_column_fields(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        next_row_key: Option<&Vec<u32>>,
        empty_column_key: &Vec<u32>,
        row_width: usize,
    ) {
        for position in row_group_end_positions(row_key, next_row_key) {
            if is_row_subtotal_position(row_key, position) && row_subtotal_enabled(plan, position) {
                let prefix = row_key[..=position].to_vec();
                if let Some(states) = aggregation.row_subtotal_states(&prefix) {
                    for subtotal in row_subtotals_for_position(plan, position) {
                        let aggregate_override = subtotal_aggregate_for_field(subtotal);
                        for measure_index in 0..plan.measures.len() {
                            let mut row = vec![
                                compact_subtotal_label_cell(snapshot, plan, &prefix, subtotal),
                                CellValue::string(plan.measures[measure_index].caption()),
                            ];
                            let context = ShowAsContext {
                                snapshot,
                                plan,
                                aggregation,
                                row_key: None,
                                column_key: Some(empty_column_key),
                            };
                            row.push(finalize_measure_from_state_slice(
                                states,
                                &plan.measures,
                                states,
                                aggregation
                                    .column_total_states(empty_column_key)
                                    .map(Vec::as_slice)
                                    .unwrap_or(&[]),
                                aggregation.grand_total_states(),
                                &context,
                                measure_index,
                                aggregate_override,
                            ));
                            self.push_measure_row(row, measure_index);
                        }
                    }
                }
            }
            append_blank_row_after_ended_row_field(&mut self.cells, plan, position, row_width);
            self.sync_unmeasured_rows();
        }
    }

    fn append_compact_values_on_rows_row_subtotals_with_column_fields(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregation: &PivotAggregation,
        row_key: &[u32],
        next_row_key: Option<&Vec<u32>>,
        column_slots: &[ColumnRenderSlot],
        row_width: usize,
    ) {
        for position in row_group_end_positions(row_key, next_row_key) {
            if is_row_subtotal_position(row_key, position) && row_subtotal_enabled(plan, position) {
                let prefix = row_key[..=position].to_vec();
                let row_total = aggregation.row_subtotal_states(&prefix);
                for subtotal in row_subtotals_for_position(plan, position) {
                    let row_aggregate_override = subtotal_aggregate_for_field(subtotal);
                    for measure_index in 0..plan.measures.len() {
                        let mut row = vec![
                            compact_subtotal_label_cell(snapshot, plan, &prefix, subtotal),
                            CellValue::string(plan.measures[measure_index].caption()),
                        ];
                        for slot in column_slots {
                            let context = ShowAsContext {
                                snapshot,
                                plan,
                                aggregation,
                                row_key: None,
                                column_key: column_context_key(slot),
                            };
                            row.push(finalize_measure_from_states(
                                subtotal_row_slot_states(aggregation, &prefix, slot),
                                &plan.measures,
                                row_total,
                                column_slot_total(aggregation, slot),
                                aggregation.grand_total_states(),
                                &context,
                                measure_index,
                                row_aggregate_override
                                    .or_else(|| column_slot_aggregate_override(slot)),
                            ));
                        }
                        self.push_measure_row(row, measure_index);
                    }
                }
            }
            append_blank_row_after_ended_row_field(&mut self.cells, plan, position, row_width);
            self.sync_unmeasured_rows();
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum ColumnRenderSlot {
    Leaf(Vec<u32>),
    Subtotal {
        prefix: Vec<u32>,
        subtotal: PivotSubtotal,
    },
    GrandTotal,
}

pub(crate) fn column_render_slots(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<ColumnRenderSlot> {
    let mut slots = Vec::new();
    for (column_index, column_key) in aggregation.column_order.iter().enumerate() {
        let previous_column_key = column_index
            .checked_sub(1)
            .and_then(|index| aggregation.column_order.get(index));
        for position in column_group_start_positions(column_key, previous_column_key) {
            if column_subtotal_at_top(pivot, plan, position) {
                for subtotal in column_subtotals_for_position(plan, position) {
                    slots.push(ColumnRenderSlot::Subtotal {
                        prefix: column_key[..=position].to_vec(),
                        subtotal,
                    });
                }
            }
        }

        slots.push(ColumnRenderSlot::Leaf(column_key.clone()));

        let next_column_key = aggregation.column_order.get(column_index + 1);
        for position in column_subtotal_positions_to_emit(pivot, plan, column_key, next_column_key)
        {
            for subtotal in column_subtotals_for_position(plan, position) {
                slots.push(ColumnRenderSlot::Subtotal {
                    prefix: column_key[..=position].to_vec(),
                    subtotal,
                });
            }
        }
    }

    if pivot.layout.show_column_grand_totals {
        slots.push(ColumnRenderSlot::GrandTotal);
    }
    slots
}

pub(crate) fn column_slot_label(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    slot: &ColumnRenderSlot,
) -> String {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => {
            key_label(snapshot, plan, &plan.column_indexes, column_key)
        }
        ColumnRenderSlot::Subtotal { prefix, subtotal } => {
            subtotal_key_label(snapshot, plan, &plan.column_indexes, prefix, *subtotal)
        }
        ColumnRenderSlot::GrandTotal => "Grand Total".to_string(),
    }
}

pub(crate) fn column_context_key(slot: &ColumnRenderSlot) -> Option<&Vec<u32>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => Some(column_key),
        ColumnRenderSlot::Subtotal { .. } | ColumnRenderSlot::GrandTotal => None,
    }
}

pub(crate) fn column_slot_aggregate_override(slot: &ColumnRenderSlot) -> Option<PivotAggregate> {
    match slot {
        ColumnRenderSlot::Subtotal { subtotal, .. } => subtotal_aggregate_for_field(*subtotal),
        ColumnRenderSlot::Leaf(_) | ColumnRenderSlot::GrandTotal => None,
    }
}

pub(crate) fn leaf_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    row_key: &[u32],
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.groups.get(&GroupKey {
            rows: row_key.to_vec(),
            columns: column_key.clone(),
        }),
        ColumnRenderSlot::Subtotal { prefix, .. } => {
            aggregation.subtotal_group_states(row_key, prefix)
        }
        ColumnRenderSlot::GrandTotal => aggregation.row_total_states(row_key),
    }
}

pub(crate) fn subtotal_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    row_prefix: &[u32],
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => {
            aggregation.subtotal_group_states(row_prefix, column_key)
        }
        ColumnRenderSlot::Subtotal { prefix, .. } => {
            aggregation.subtotal_group_states(row_prefix, prefix)
        }
        ColumnRenderSlot::GrandTotal => aggregation.row_subtotal_states(row_prefix),
    }
}

pub(crate) fn grand_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.column_total_states(column_key),
        ColumnRenderSlot::Subtotal { prefix, .. } => aggregation.column_subtotal_states(prefix),
        ColumnRenderSlot::GrandTotal => Some(aggregation.grand_total_states()),
    }
}

pub(crate) fn column_slot_total<'a>(
    aggregation: &'a PivotAggregation,
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.column_total_states(column_key),
        ColumnRenderSlot::Subtotal { prefix, .. } => aggregation.column_subtotal_states(prefix),
        ColumnRenderSlot::GrandTotal => Some(aggregation.grand_total_states()),
    }
}

pub(crate) fn append_row_top_subtotals_without_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
    empty_column_key: &Vec<u32>,
) -> bool {
    let mut emitted = false;
    for position in row_group_start_positions(row_key, previous_row_key) {
        if !row_subtotal_at_top(pivot, plan, position) {
            continue;
        }

        let prefix = row_key[..=position].to_vec();
        if let Some(states) = aggregation.row_subtotal_states(&prefix) {
            for subtotal in row_subtotals_for_position(plan, position) {
                let mut row = row_subtotal_label_cells(snapshot, plan, &prefix, subtotal);
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: None,
                    column_key: Some(empty_column_key),
                };
                row.extend(finalize_state_slice_with_context_and_aggregate(
                    states,
                    &plan.measures,
                    states,
                    aggregation
                        .column_total_states(empty_column_key)
                        .map(Vec::as_slice)
                        .unwrap_or(&[]),
                    aggregation.grand_total_states(),
                    &context,
                    subtotal_aggregate_for_field(subtotal),
                ));
                cells.push(row);
            }
            emitted = true;
        }
    }
    emitted
}

pub(crate) fn append_row_top_subtotals_with_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
    column_slots: &[ColumnRenderSlot],
) -> bool {
    let mut emitted = false;
    for position in row_group_start_positions(row_key, previous_row_key) {
        if !row_subtotal_at_top(pivot, plan, position) {
            continue;
        }

        let prefix = row_key[..=position].to_vec();
        let row_total = aggregation.row_subtotal_states(&prefix);

        for subtotal in row_subtotals_for_position(plan, position) {
            let mut row = row_subtotal_label_cells(snapshot, plan, &prefix, subtotal);
            let row_aggregate_override = subtotal_aggregate_for_field(subtotal);
            for slot in column_slots {
                let context = ShowAsContext {
                    snapshot,
                    plan,
                    aggregation,
                    row_key: None,
                    column_key: column_context_key(slot),
                };
                row.extend(finalize_states_with_context_and_aggregate(
                    subtotal_row_slot_states(aggregation, &prefix, slot),
                    &plan.measures,
                    row_total,
                    column_slot_total(aggregation, slot),
                    aggregation.grand_total_states(),
                    &context,
                    row_aggregate_override.or_else(|| column_slot_aggregate_override(slot)),
                ));
            }

            cells.push(row);
        }
        emitted = true;
    }
    emitted
}

pub(crate) fn append_row_subtotals_without_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    empty_column_key: &Vec<u32>,
    row_width: usize,
    pivot: &PivotTable,
) {
    for position in row_group_end_positions(row_key, next_row_key) {
        if is_row_subtotal_position(row_key, position)
            && row_subtotal_enabled(plan, position)
            && !row_subtotal_at_top(pivot, plan, position)
        {
            let prefix = row_key[..=position].to_vec();
            if let Some(states) = aggregation.row_subtotal_states(&prefix) {
                for subtotal in row_subtotals_for_position(plan, position) {
                    let mut row = row_subtotal_label_cells(snapshot, plan, &prefix, subtotal);
                    let context = ShowAsContext {
                        snapshot,
                        plan,
                        aggregation,
                        row_key: None,
                        column_key: Some(empty_column_key),
                    };
                    row.extend(finalize_state_slice_with_context_and_aggregate(
                        states,
                        &plan.measures,
                        states,
                        aggregation
                            .column_total_states(empty_column_key)
                            .map(Vec::as_slice)
                            .unwrap_or(&[]),
                        aggregation.grand_total_states(),
                        &context,
                        subtotal_aggregate_for_field(subtotal),
                    ));
                    cells.push(row);
                }
            }
        }
        append_blank_row_after_ended_row_field(cells, plan, position, row_width);
    }
}

pub(crate) fn append_row_subtotals_with_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    column_slots: &[ColumnRenderSlot],
    row_width: usize,
    pivot: &PivotTable,
) {
    for position in row_group_end_positions(row_key, next_row_key) {
        if is_row_subtotal_position(row_key, position)
            && row_subtotal_enabled(plan, position)
            && !row_subtotal_at_top(pivot, plan, position)
        {
            let prefix = row_key[..=position].to_vec();
            let row_total = aggregation.row_subtotal_states(&prefix);

            for subtotal in row_subtotals_for_position(plan, position) {
                let mut row = row_subtotal_label_cells(snapshot, plan, &prefix, subtotal);
                let row_aggregate_override = subtotal_aggregate_for_field(subtotal);
                for slot in column_slots {
                    let context = ShowAsContext {
                        snapshot,
                        plan,
                        aggregation,
                        row_key: None,
                        column_key: column_context_key(slot),
                    };
                    row.extend(finalize_states_with_context_and_aggregate(
                        subtotal_row_slot_states(aggregation, &prefix, slot),
                        &plan.measures,
                        row_total,
                        column_slot_total(aggregation, slot),
                        aggregation.grand_total_states(),
                        &context,
                        row_aggregate_override.or_else(|| column_slot_aggregate_override(slot)),
                    ));
                }

                cells.push(row);
            }
        }
        append_blank_row_after_ended_row_field(cells, plan, position, row_width);
    }
}

pub(crate) fn append_compact_row_subtotals_without_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    empty_column_key: &Vec<u32>,
    row_width: usize,
) {
    for position in row_group_end_positions(row_key, next_row_key) {
        if is_row_subtotal_position(row_key, position) && row_subtotal_enabled(plan, position) {
            let prefix = row_key[..=position].to_vec();
            if let Some(states) = aggregation.row_subtotal_states(&prefix) {
                for subtotal in row_subtotals_for_position(plan, position) {
                    let mut row = vec![compact_subtotal_label_cell(
                        snapshot, plan, &prefix, subtotal,
                    )];
                    let context = ShowAsContext {
                        snapshot,
                        plan,
                        aggregation,
                        row_key: None,
                        column_key: Some(empty_column_key),
                    };
                    row.extend(finalize_state_slice_with_context_and_aggregate(
                        states,
                        &plan.measures,
                        states,
                        aggregation
                            .column_total_states(empty_column_key)
                            .map(Vec::as_slice)
                            .unwrap_or(&[]),
                        aggregation.grand_total_states(),
                        &context,
                        subtotal_aggregate_for_field(subtotal),
                    ));
                    cells.push(row);
                }
            }
        }
        append_blank_row_after_ended_row_field(cells, plan, position, row_width);
    }
}

pub(crate) fn append_compact_row_subtotals_with_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    column_slots: &[ColumnRenderSlot],
    row_width: usize,
) {
    for position in row_group_end_positions(row_key, next_row_key) {
        if is_row_subtotal_position(row_key, position) && row_subtotal_enabled(plan, position) {
            let prefix = row_key[..=position].to_vec();
            let row_total = aggregation.row_subtotal_states(&prefix);

            for subtotal in row_subtotals_for_position(plan, position) {
                let mut row = vec![compact_subtotal_label_cell(
                    snapshot, plan, &prefix, subtotal,
                )];
                let row_aggregate_override = subtotal_aggregate_for_field(subtotal);
                for slot in column_slots {
                    let context = ShowAsContext {
                        snapshot,
                        plan,
                        aggregation,
                        row_key: None,
                        column_key: column_context_key(slot),
                    };
                    row.extend(finalize_states_with_context_and_aggregate(
                        subtotal_row_slot_states(aggregation, &prefix, slot),
                        &plan.measures,
                        row_total,
                        column_slot_total(aggregation, slot),
                        aggregation.grand_total_states(),
                        &context,
                        row_aggregate_override.or_else(|| column_slot_aggregate_override(slot)),
                    ));
                }

                cells.push(row);
            }
        }
        append_blank_row_after_ended_row_field(cells, plan, position, row_width);
    }
}

pub(crate) fn row_subtotals_for_position(
    plan: &CompiledPivotPlan,
    position: usize,
) -> Vec<PivotSubtotal> {
    plan.row_fields
        .get(position)
        .map(enabled_subtotals_for_field)
        .unwrap_or_default()
}

pub(crate) fn column_subtotals_for_position(
    plan: &CompiledPivotPlan,
    position: usize,
) -> Vec<PivotSubtotal> {
    plan.column_fields
        .get(position)
        .map(enabled_subtotals_for_field)
        .unwrap_or_default()
}

pub(crate) fn enabled_subtotals_for_field(field: &PivotField) -> Vec<PivotSubtotal> {
    if field.subtotals.is_empty() {
        if matches!(field.subtotal, PivotSubtotal::None) {
            Vec::new()
        } else {
            vec![field.subtotal]
        }
    } else {
        field
            .subtotals
            .iter()
            .copied()
            .filter(|subtotal| !matches!(subtotal, PivotSubtotal::None))
            .collect()
    }
}

pub(crate) fn row_item_collapsed(
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    position: usize,
    enabled: bool,
) -> bool {
    axis_item_collapsed(&plan.row_collapsed_item_ids, row_key, position, enabled)
}

pub(crate) fn column_prefix_collapsed(
    plan: &CompiledPivotPlan,
    column_key: &[u32],
    position: usize,
    enabled: bool,
) -> bool {
    enabled
        && plan
            .column_collapsed_item_ids
            .get(position)
            .zip(column_key.get(position))
            .map(|(ids, id)| ids.contains(id))
            .unwrap_or(false)
}

pub(crate) fn axis_item_collapsed(
    collapsed_item_ids: &[AHashSet<u32>],
    key: &[u32],
    position: usize,
    enabled: bool,
) -> bool {
    enabled
        && position < key.len().saturating_sub(1)
        && collapsed_item_ids
            .get(position)
            .map(|ids| ids.contains(&key[position]))
            .unwrap_or(false)
}

pub(crate) fn first_row_collapsed_position(
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    enabled: bool,
) -> Option<usize> {
    first_axis_collapsed_position(&plan.row_collapsed_item_ids, row_key, enabled, false)
}

pub(crate) fn first_column_collapsed_position(
    plan: &CompiledPivotPlan,
    column_key: &[u32],
    enabled: bool,
) -> Option<usize> {
    first_axis_collapsed_position(&plan.column_collapsed_item_ids, column_key, enabled, false)
}

pub(crate) fn first_column_prefix_collapsed_position(
    plan: &CompiledPivotPlan,
    column_key: &[u32],
    enabled: bool,
) -> Option<usize> {
    first_axis_collapsed_position(&plan.column_collapsed_item_ids, column_key, enabled, true)
}

pub(crate) fn first_axis_collapsed_position(
    collapsed_item_ids: &[AHashSet<u32>],
    key: &[u32],
    enabled: bool,
    include_last_position: bool,
) -> Option<usize> {
    if !enabled {
        return None;
    }
    let end = if include_last_position {
        key.len()
    } else {
        key.len().saturating_sub(1)
    };
    (0..end).find(|position| {
        collapsed_item_ids
            .get(*position)
            .map(|ids| ids.contains(&key[*position]))
            .unwrap_or(false)
    })
}

pub(crate) fn position_hidden_by_collapsed(
    collapsed_position: Option<usize>,
    position: usize,
) -> bool {
    collapsed_position
        .map(|collapsed_position| collapsed_position < position)
        .unwrap_or(false)
}

pub(crate) fn row_subtotal_at_top(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    position: usize,
) -> bool {
    matches!(pivot.layout.kind, PivotLayoutKind::Outline)
        && plan
            .row_fields
            .get(position)
            .map(|field| field.subtotal_top)
            .unwrap_or(false)
}

pub(crate) fn column_subtotal_at_top(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    position: usize,
) -> bool {
    matches!(pivot.layout.kind, PivotLayoutKind::Outline)
        && plan
            .column_fields
            .get(position)
            .map(|field| field.subtotal_top)
            .unwrap_or(false)
}

pub(crate) fn column_subtotal_positions_to_emit(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    column_key: &[u32],
    next_column_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    subtotal_positions_to_emit(column_key, next_column_key, |position| {
        column_subtotal_enabled(plan, position) && !column_subtotal_at_top(pivot, plan, position)
    })
}

pub(crate) fn row_group_start_positions(
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    group_start_positions(row_key, previous_row_key)
}

pub(crate) fn column_group_start_positions(
    column_key: &[u32],
    previous_column_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    group_start_positions(column_key, previous_column_key)
}

pub(crate) fn group_start_positions(key: &[u32], previous_key: Option<&Vec<u32>>) -> Vec<usize> {
    if key.len() < 2 {
        return Vec::new();
    }

    (0..(key.len() - 1))
        .filter(|position| {
            previous_key
                .map(|previous| !same_prefix(key, previous, *position + 1))
                .unwrap_or(true)
        })
        .collect()
}

pub(crate) fn subtotal_positions_to_emit(
    key: &[u32],
    next_key: Option<&Vec<u32>>,
    enabled: impl Fn(usize) -> bool,
) -> Vec<usize> {
    if key.len() < 2 {
        return Vec::new();
    }

    (0..(key.len() - 1))
        .rev()
        .filter(|position| enabled(*position))
        .filter(|position| {
            next_key
                .map(|next| !same_prefix(key, next, *position + 1))
                .unwrap_or(true)
        })
        .collect()
}

pub(crate) fn row_subtotal_prefixes(plan: &CompiledPivotPlan, row_key: &[u32]) -> Vec<Vec<u32>> {
    subtotal_prefixes(row_key, |position| {
        row_subtotal_enabled(plan, position) || plan.row_parent_total_positions.contains(&position)
    })
}

pub(crate) fn column_subtotal_prefixes(
    plan: &CompiledPivotPlan,
    column_key: &[u32],
) -> Vec<Vec<u32>> {
    subtotal_prefixes(column_key, |position| {
        column_subtotal_enabled(plan, position)
            || plan.column_parent_total_positions.contains(&position)
    })
}

pub(crate) fn subtotal_prefixes(key: &[u32], enabled: impl Fn(usize) -> bool) -> Vec<Vec<u32>> {
    (1..key.len())
        .filter(|prefix_len| enabled(prefix_len - 1))
        .map(|prefix_len| key[..prefix_len].to_vec())
        .collect()
}

pub(crate) fn row_subtotal_enabled(plan: &CompiledPivotPlan, position: usize) -> bool {
    !row_subtotals_for_position(plan, position).is_empty()
}

pub(crate) fn column_subtotal_enabled(plan: &CompiledPivotPlan, position: usize) -> bool {
    !column_subtotals_for_position(plan, position).is_empty()
}

pub(crate) fn row_group_end_positions(
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    (0..row_key.len())
        .rev()
        .filter(|position| {
            next_row_key
                .map(|next| !same_prefix(row_key, next, *position + 1))
                .unwrap_or(true)
        })
        .collect()
}

pub(crate) fn is_row_subtotal_position(row_key: &[u32], position: usize) -> bool {
    position < row_key.len().saturating_sub(1)
}

pub(crate) fn append_blank_row_after_ended_row_field(
    cells: &mut Vec<Vec<CellValue>>,
    plan: &CompiledPivotPlan,
    position: usize,
    width: usize,
) {
    if row_field_inserts_blank_row(plan, position) {
        cells.push(vec![CellValue::Empty; width]);
    }
}

pub(crate) fn row_field_inserts_blank_row(plan: &CompiledPivotPlan, position: usize) -> bool {
    plan.row_fields
        .get(position)
        .map(|field| field.insert_blank_row)
        .unwrap_or(false)
}

pub(crate) fn row_field_inserts_page_break(plan: &CompiledPivotPlan, position: usize) -> bool {
    plan.row_fields
        .get(position)
        .map(|field| field.insert_page_break)
        .unwrap_or(false)
}

pub(crate) fn subtotal_aggregate_for_field(subtotal: PivotSubtotal) -> Option<PivotAggregate> {
    match subtotal {
        PivotSubtotal::Automatic | PivotSubtotal::None => None,
        PivotSubtotal::Sum => Some(PivotAggregate::Sum),
        PivotSubtotal::Count => Some(PivotAggregate::Count),
        PivotSubtotal::CountNumbers => Some(PivotAggregate::CountNumbers),
        PivotSubtotal::Average => Some(PivotAggregate::Average),
        PivotSubtotal::Min => Some(PivotAggregate::Min),
        PivotSubtotal::Max => Some(PivotAggregate::Max),
        PivotSubtotal::Product => Some(PivotAggregate::Product),
        PivotSubtotal::StdDev => Some(PivotAggregate::StdDev),
        PivotSubtotal::StdDevP => Some(PivotAggregate::StdDevP),
        PivotSubtotal::Var => Some(PivotAggregate::Var),
        PivotSubtotal::VarP => Some(PivotAggregate::VarP),
    }
}

pub(crate) fn row_subtotal_label_cells(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    prefix: &[u32],
    subtotal: PivotSubtotal,
) -> Vec<CellValue> {
    let row_indexes = &plan.row_indexes;
    let mut row = vec![CellValue::Empty; row_indexes.len()];
    if prefix.is_empty() {
        return row;
    }

    let subtotal_position = prefix.len() - 1;
    for index in 0..subtotal_position {
        row[index] = pivot_value_cell(
            plan,
            snapshot.value_by_id(row_indexes[index], prefix[index]),
        );
    }
    let value = pivot_value_label(
        plan,
        snapshot.value_by_id(row_indexes[subtotal_position], prefix[subtotal_position]),
    );
    let field = plan.row_fields.get(subtotal_position);
    row[subtotal_position] = CellValue::string(subtotal_label(plan, field, &value, subtotal));
    row
}

pub(crate) fn subtotal_key_label(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_indexes: &[usize],
    prefix: &[u32],
    subtotal: PivotSubtotal,
) -> String {
    let mut labels = field_indexes
        .iter()
        .zip(prefix.iter())
        .map(|(field_index, id)| pivot_value_label(plan, snapshot.value_by_id(*field_index, *id)))
        .collect::<Vec<_>>();
    if let Some(label) = labels.last_mut() {
        let field = prefix
            .len()
            .checked_sub(1)
            .and_then(|position| plan.column_fields.get(position));
        let value = label.as_ref().to_string();
        *label = Cow::Owned(subtotal_label(plan, field, &value, subtotal));
    }
    labels.join(" | ")
}

pub(crate) fn subtotal_label(
    plan: &CompiledPivotPlan,
    field: Option<&PivotField>,
    value: &str,
    subtotal: PivotSubtotal,
) -> String {
    let suffix = if field
        .map(|field| enabled_subtotals_for_field(field).len() > 1)
        .unwrap_or(false)
    {
        subtotal_aggregate_for_field(subtotal)
            .map(|aggregate| aggregate.caption())
            .unwrap_or("Total")
    } else {
        field
            .and_then(|field| field.subtotal_caption.as_deref())
            .unwrap_or("Total")
    };
    total_caption(plan, &format!("{value} {suffix}"))
}

pub(crate) fn same_prefix(left: &[u32], right: &[u32], len: usize) -> bool {
    left.len() >= len && right.len() >= len && left[..len] == right[..len]
}

pub(crate) fn prepend_page_fields(
    cells: &mut Vec<Vec<CellValue>>,
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
) {
    if plan.page_fields.is_empty() {
        return;
    }

    let display_rows = page_field_display_row_count(pivot, plan);
    let mut rows = vec![Vec::new(); display_rows];
    for (index, (field, field_index)) in plan
        .page_fields
        .iter()
        .zip(plan.page_indexes.iter())
        .enumerate()
    {
        let (row, column) = page_field_grid_position(pivot, index, display_rows);
        let start_column = column * 2;
        let values = &mut rows[row];
        values.resize(start_column + 2, CellValue::Empty);
        values[start_column] = CellValue::string(pivot_field_caption(field));
        values[start_column + 1] = CellValue::string(page_field_caption(
            pivot,
            snapshot,
            plan,
            *field_index,
            &field.field.name,
        ));
    }
    rows.push(Vec::new());
    rows.append(cells);
    *cells = rows;
}

pub(crate) fn page_field_grid_position(
    pivot: &PivotTable,
    index: usize,
    display_rows: usize,
) -> (usize, usize) {
    let wrap = pivot.layout.page_wrap as usize;
    if wrap == 0 {
        (index, 0)
    } else if pivot.layout.page_over_then_down {
        (index / wrap, index % wrap)
    } else {
        debug_assert!(display_rows > 0);
        (index % display_rows, index / display_rows)
    }
}

pub(crate) fn page_field_caption(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_index: usize,
    field_name: &str,
) -> String {
    if !plan
        .filters
        .iter()
        .any(|filter| filter.field_index() == field_index)
    {
        return "(All)".to_string();
    }

    let allowed_item_ids = page_field_allowed_item_ids(snapshot, plan, field_index);
    match allowed_item_ids.as_slice() {
        [] => explicit_single_page_item_caption(pivot, plan, field_name)
            .unwrap_or_else(|| "(All)".to_string()),
        [item_id] => {
            pivot_value_label(plan, snapshot.value_by_id(field_index, *item_id)).into_owned()
        }
        _ => {
            if allowed_item_ids.len() == snapshot.columns[field_index].dictionary.len() {
                "(All)".to_string()
            } else if pivot.layout.show_multiple_label {
                "(Multiple Items)".to_string()
            } else {
                "(All)".to_string()
            }
        }
    }
}

pub(crate) fn explicit_single_page_item_caption(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    field_name: &str,
) -> Option<String> {
    pivot.filters.iter().find_map(|filter| {
        let PivotFilter::FieldItems {
            field,
            allowed_items,
        } = filter
        else {
            return None;
        };
        if field.name.eq_ignore_ascii_case(field_name) && allowed_items.len() == 1 {
            Some(pivot_value_label(plan, &allowed_items[0]).into_owned())
        } else {
            None
        }
    })
}

pub(crate) fn page_field_allowed_item_ids(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_index: usize,
) -> Vec<u32> {
    snapshot.columns[field_index]
        .dictionary
        .iter()
        .enumerate()
        .filter_map(|(item_id, _)| {
            let item_id = item_id as u32;
            plan.filters
                .iter()
                .all(|filter| filter.allows_item(snapshot, field_index, item_id))
                .then_some(item_id)
        })
        .collect()
}

pub(crate) fn decode_key_cells(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_indexes: &[usize],
    key: &[u32],
) -> Vec<CellValue> {
    field_indexes
        .iter()
        .zip(key.iter())
        .map(|(field_index, id)| pivot_value_cell(plan, snapshot.value_by_id(*field_index, *id)))
        .collect()
}

pub(crate) fn row_label_cells(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
) -> Vec<CellValue> {
    let mut cells = decode_key_cells(snapshot, plan, &plan.row_indexes, row_key);
    if !matches!(
        pivot.layout.kind,
        PivotLayoutKind::Tabular | PivotLayoutKind::Outline
    ) || pivot.layout.repeat_item_labels
    {
        return cells;
    }

    if let Some(previous) = previous_row_key {
        for position in 0..row_key.len().saturating_sub(1) {
            if same_prefix(row_key, previous, position + 1) {
                cells[position] = CellValue::Empty;
            }
        }
    }
    cells
}

pub(crate) fn append_compact_group_headers(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
    data_width: usize,
) {
    for position in compact_group_header_positions(row_key, previous_row_key) {
        let mut row = vec![key_position_cell(
            snapshot,
            plan,
            &plan.row_indexes,
            row_key,
            position,
        )];
        row.extend(empty_cells(data_width));
        cells.push(row);
    }
}

pub(crate) fn compact_group_header_positions(
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    if row_key.len() < 2 {
        return Vec::new();
    }

    (0..(row_key.len() - 1))
        .filter(|position| {
            previous_row_key
                .map(|previous| !same_prefix(row_key, previous, *position + 1))
                .unwrap_or(true)
        })
        .collect()
}

pub(crate) fn compact_leaf_label_cell(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
) -> CellValue {
    let position = row_key.len().saturating_sub(1);
    key_position_cell(snapshot, plan, &plan.row_indexes, row_key, position)
}

pub(crate) fn compact_subtotal_label_cell(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    prefix: &[u32],
    subtotal: PivotSubtotal,
) -> CellValue {
    let Some(position) = prefix.len().checked_sub(1) else {
        return CellValue::Empty;
    };
    let value = pivot_value_label(
        plan,
        snapshot.value_by_id(plan.row_indexes[position], prefix[position]),
    );
    let field = plan.row_fields.get(position);
    CellValue::string(subtotal_label(plan, field, &value, subtotal))
}

pub(crate) fn key_position_cell(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_indexes: &[usize],
    key: &[u32],
    position: usize,
) -> CellValue {
    field_indexes
        .get(position)
        .zip(key.get(position))
        .map(|(field_index, id)| pivot_value_cell(plan, snapshot.value_by_id(*field_index, *id)))
        .unwrap_or(CellValue::Empty)
}

pub(crate) fn pivot_value_cell(plan: &CompiledPivotPlan, value: &PivotValue) -> CellValue {
    if matches!(value, PivotValue::Error(_)) {
        if let Some(caption) = &plan.error_caption {
            return CellValue::string(caption);
        }
    }
    value.to_cell_value()
}

pub(crate) fn pivot_value_label<'a>(
    plan: &'a CompiledPivotPlan,
    value: &'a PivotValue,
) -> Cow<'a, str> {
    if matches!(value, PivotValue::Error(_)) {
        if let Some(caption) = &plan.error_caption {
            return Cow::Borrowed(caption.as_str());
        }
    }

    match value {
        PivotValue::Blank => Cow::Borrowed(""),
        PivotValue::Boolean(value) => Cow::Borrowed(if *value { "TRUE" } else { "FALSE" }),
        PivotValue::Number(value) => Cow::Owned(value.to_string()),
        PivotValue::String(value) => Cow::Borrowed(value.as_str()),
        PivotValue::Error(value) => Cow::Borrowed(value.as_str()),
    }
}

pub(crate) fn empty_cells(count: usize) -> impl Iterator<Item = CellValue> {
    std::iter::repeat_with(|| CellValue::Empty).take(count)
}

pub(crate) fn grand_total_caption(pivot: &PivotTable) -> String {
    let caption = pivot
        .layout
        .grand_total_caption
        .as_deref()
        .unwrap_or("Grand Total");
    if pivot.layout.asterisk_totals {
        format!("{caption}*")
    } else {
        caption.to_string()
    }
}

pub(crate) fn grand_total_label_row(label_width: usize, caption: &str) -> Vec<CellValue> {
    if label_width == 0 {
        Vec::new()
    } else {
        let mut row = vec![CellValue::Empty; label_width];
        row[0] = CellValue::string(caption);
        row
    }
}

pub(crate) fn total_caption(plan: &CompiledPivotPlan, caption: &str) -> String {
    if plan.asterisk_totals {
        format!("{caption}*")
    } else {
        caption.to_string()
    }
}

pub(crate) fn values_caption(pivot: &PivotTable) -> &str {
    let caption = pivot.layout.data_caption.trim();
    if caption.is_empty() {
        "Values"
    } else {
        pivot.layout.data_caption.as_str()
    }
}

pub(crate) fn key_label(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    field_indexes: &[usize],
    key: &[u32],
) -> String {
    field_indexes
        .iter()
        .zip(key.iter())
        .map(|(field_index, id)| pivot_value_label(plan, snapshot.value_by_id(*field_index, *id)))
        .collect::<Vec<_>>()
        .join(" | ")
}

pub(crate) fn measure_column_caption(
    column_label: &str,
    measure: &PivotMeasure,
    measure_count: usize,
) -> String {
    if measure_count == 1 {
        column_label.to_string()
    } else {
        format!("{} {}", column_label, measure.caption())
    }
}

pub(crate) fn grand_total_measure_caption(
    caption: &str,
    measure: &PivotMeasure,
    measure_count: usize,
) -> String {
    if measure_count == 1 {
        caption.to_string()
    } else {
        format!("{} {}", caption, measure.caption())
    }
}

pub(crate) fn output_range(
    target: CellAddress,
    row_count: usize,
    col_count: usize,
) -> Result<CellRange> {
    if row_count == 0 || col_count == 0 {
        return Err(Error::other("pivot output cannot be empty"));
    }

    let end_row = target
        .row
        .checked_add(row_count as u32 - 1)
        .ok_or_else(|| Error::RowOutOfBounds(MAX_ROWS, MAX_ROWS - 1))?;
    if end_row >= MAX_ROWS {
        return Err(Error::RowOutOfBounds(end_row, MAX_ROWS - 1));
    }

    let end_col = target.col as u32 + col_count as u32 - 1;
    if end_col >= MAX_COLS as u32 {
        return Err(Error::ColumnOutOfBounds(end_col as u16, MAX_COLS - 1));
    }

    Ok(CellRange::from_indices(
        target.row,
        target.col,
        end_row,
        end_col as u16,
    ))
}
