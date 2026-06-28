use crate::aggregate::*;
use crate::compile::*;
use crate::prelude::*;
use crate::render::*;
use crate::snapshot::*;

#[derive(Debug, Clone, Copy)]
pub(crate) struct ShowAsContext<'a> {
    pub(crate) snapshot: &'a SourceSnapshot,
    pub(crate) plan: &'a CompiledPivotPlan,
    pub(crate) aggregation: &'a PivotAggregation,
    pub(crate) row_key: Option<&'a Vec<u32>>,
    pub(crate) column_key: Option<&'a Vec<u32>>,
}

pub(crate) fn finalize_states_with_context(
    states: Option<&Vec<AggregateState>>,
    measures: &[PivotMeasure],
    row_total: Option<&Vec<AggregateState>>,
    column_total: Option<&Vec<AggregateState>>,
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
) -> Vec<CellValue> {
    finalize_states_with_context_and_aggregate(
        states,
        measures,
        row_total,
        column_total,
        grand_total,
        context,
        None,
    )
}

pub(crate) fn finalize_states_with_context_and_aggregate(
    states: Option<&Vec<AggregateState>>,
    measures: &[PivotMeasure],
    row_total: Option<&Vec<AggregateState>>,
    column_total: Option<&Vec<AggregateState>>,
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
    aggregate_override: Option<PivotAggregate>,
) -> Vec<CellValue> {
    states
        .map(|states| {
            finalize_state_slice_with_context_and_aggregate(
                states,
                measures,
                row_total.map(Vec::as_slice).unwrap_or(&[]),
                column_total.map(Vec::as_slice).unwrap_or(&[]),
                grand_total,
                context,
                aggregate_override,
            )
        })
        .unwrap_or_else(|| vec![missing_value_cell(context.plan); measures.len()])
}

pub(crate) fn finalize_measure_from_states(
    states: Option<&Vec<AggregateState>>,
    measures: &[PivotMeasure],
    row_total: Option<&Vec<AggregateState>>,
    column_total: Option<&Vec<AggregateState>>,
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
    measure_index: usize,
    aggregate_override: Option<PivotAggregate>,
) -> CellValue {
    let Some(states) = states else {
        return missing_value_cell(context.plan);
    };
    finalize_measure_from_state_slice(
        states,
        measures,
        row_total.map(Vec::as_slice).unwrap_or(&[]),
        column_total.map(Vec::as_slice).unwrap_or(&[]),
        grand_total,
        context,
        measure_index,
        aggregate_override,
    )
}

pub(crate) fn finalize_state_slice_with_context(
    states: &[AggregateState],
    measures: &[PivotMeasure],
    row_total: &[AggregateState],
    column_total: &[AggregateState],
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
) -> Vec<CellValue> {
    finalize_state_slice_with_context_and_aggregate(
        states,
        measures,
        row_total,
        column_total,
        grand_total,
        context,
        None,
    )
}

pub(crate) fn finalize_measure_from_state_slice(
    states: &[AggregateState],
    measures: &[PivotMeasure],
    row_total: &[AggregateState],
    column_total: &[AggregateState],
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
    measure_index: usize,
    aggregate_override: Option<PivotAggregate>,
) -> CellValue {
    let Some(measure) = measures.get(measure_index) else {
        return CellValue::Empty;
    };
    let Some(state) = states.get(measure_index) else {
        return missing_value_cell(context.plan);
    };
    let aggregate = aggregate_override.unwrap_or(measure.aggregate);
    finalize_measure_with_context(
        state,
        measure,
        aggregate,
        state_number(row_total, measure_index, aggregate),
        state_number(column_total, measure_index, aggregate),
        state_number(grand_total, measure_index, aggregate),
        measure_index,
        context,
    )
}

pub(crate) fn finalize_state_slice_with_context_and_aggregate(
    states: &[AggregateState],
    measures: &[PivotMeasure],
    row_total: &[AggregateState],
    column_total: &[AggregateState],
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
    aggregate_override: Option<PivotAggregate>,
) -> Vec<CellValue> {
    states
        .iter()
        .enumerate()
        .zip(measures.iter())
        .map(|((index, state), measure)| {
            let aggregate = aggregate_override.unwrap_or(measure.aggregate);
            finalize_measure_with_context(
                state,
                measure,
                aggregate,
                state_number(row_total, index, aggregate),
                state_number(column_total, index, aggregate),
                state_number(grand_total, index, aggregate),
                index,
                context,
            )
        })
        .collect()
}

pub(crate) fn finalize_measure_with_context(
    state: &AggregateState,
    measure: &PivotMeasure,
    aggregate: PivotAggregate,
    row_total: Option<f64>,
    column_total: Option<f64>,
    grand_total: Option<f64>,
    measure_index: usize,
    context: &ShowAsContext<'_>,
) -> CellValue {
    let cell = match &measure.show_as {
        PivotShowAs::Normal => state.finalize(aggregate),
        PivotShowAs::PercentOfGrandTotal => {
            percentage_cell(state.finalize_number(aggregate), grand_total)
        }
        PivotShowAs::PercentOfRowTotal => {
            percentage_cell(state.finalize_number(aggregate), row_total)
        }
        PivotShowAs::PercentOfColumnTotal => {
            percentage_cell(state.finalize_number(aggregate), column_total)
        }
        PivotShowAs::PercentOfParentRowTotal => percentage_cell(
            state.finalize_number(aggregate),
            parent_row_total_value(context, measure_index, aggregate),
        ),
        PivotShowAs::PercentOfParentColumnTotal => percentage_cell(
            state.finalize_number(aggregate),
            parent_column_total_value(context, measure_index, aggregate),
        ),
        PivotShowAs::PercentOfParentTotal { base_field } => percentage_cell(
            state.finalize_number(aggregate),
            parent_base_field_total_value(
                context,
                base_field.name.as_str(),
                measure_index,
                aggregate,
            ),
        ),
        PivotShowAs::Index => index_cell(
            state.finalize_number(aggregate),
            row_total,
            column_total,
            grand_total,
        ),
        PivotShowAs::RunningTotal { base_field } => optional_number_cell(running_total_value(
            context,
            base_field.name.as_str(),
            measure_index,
            aggregate,
        )),
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => {
            let current = state.finalize_number(aggregate);
            let base = base_item_value(
                context,
                base_field.name.as_str(),
                base_item,
                measure_index,
                aggregate,
            );
            optional_number_cell(current.zip(base).map(|(current, base)| current - base))
        }
        PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => {
            let current = state.finalize_number(aggregate);
            let base = base_item_value(
                context,
                base_field.name.as_str(),
                base_item,
                measure_index,
                aggregate,
            );
            match current.zip(base) {
                Some((current, base)) if base != 0.0 => CellValue::Number((current - base) / base),
                _ => CellValue::Empty,
            }
        }
        PivotShowAs::RankAscending { base_field } => optional_number_cell(rank_value(
            context,
            base_field.name.as_str(),
            measure_index,
            aggregate,
            true,
        )),
        PivotShowAs::RankDescending { base_field } => optional_number_cell(rank_value(
            context,
            base_field.name.as_str(),
            measure_index,
            aggregate,
            false,
        )),
    };
    apply_missing_caption(cell, context.plan)
}

pub(crate) fn apply_missing_caption(cell: CellValue, plan: &CompiledPivotPlan) -> CellValue {
    if cell.is_empty() {
        missing_value_cell(plan)
    } else {
        cell
    }
}

pub(crate) fn missing_value_cell(plan: &CompiledPivotPlan) -> CellValue {
    plan.missing_caption
        .as_deref()
        .map(CellValue::string)
        .unwrap_or(CellValue::Empty)
}

pub(crate) fn state_number(
    states: &[AggregateState],
    index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    states
        .get(index)
        .and_then(|state| state.finalize_number(aggregate))
}

pub(crate) fn percentage_cell(numerator: Option<f64>, denominator: Option<f64>) -> CellValue {
    match (numerator, denominator) {
        (Some(numerator), Some(denominator)) if denominator != 0.0 => {
            CellValue::Number(numerator / denominator)
        }
        _ => CellValue::Empty,
    }
}

pub(crate) fn index_cell(
    value: Option<f64>,
    row_total: Option<f64>,
    column_total: Option<f64>,
    grand_total: Option<f64>,
) -> CellValue {
    match (value, row_total, column_total, grand_total) {
        (Some(value), Some(row_total), Some(column_total), Some(grand_total)) => {
            if row_total == 0.0 || column_total == 0.0 {
                CellValue::Empty
            } else {
                CellValue::Number(value * grand_total / (row_total * column_total))
            }
        }
        _ => CellValue::Empty,
    }
}

pub(crate) fn optional_number_cell(value: Option<f64>) -> CellValue {
    value.map(CellValue::Number).unwrap_or(CellValue::Empty)
}

pub(crate) fn parent_row_total_value(
    context: &ShowAsContext<'_>,
    measure_index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    let Some(row_key) = context.row_key else {
        return state_number(
            grand_total_states(context.aggregation),
            measure_index,
            aggregate,
        );
    };
    let states = if row_key.len() <= 1 {
        match context.column_key {
            Some(column_key) if !column_key.is_empty() => {
                column_total_states(context.aggregation, column_key)
            }
            None => Some(grand_total_states(context.aggregation)),
            Some(_) => Some(grand_total_states(context.aggregation)),
        }
    } else {
        parent_row_prefix_states(context, &row_key[..row_key.len() - 1])
    }?;
    state_number(states, measure_index, aggregate)
}

pub(crate) fn parent_column_total_value(
    context: &ShowAsContext<'_>,
    measure_index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    let Some(column_key) = context.column_key else {
        return state_number(
            grand_total_states(context.aggregation),
            measure_index,
            aggregate,
        );
    };
    let states = if column_key.len() <= 1 {
        match context.row_key {
            Some(row_key) if !row_key.is_empty() => row_total_states(context.aggregation, row_key),
            None => Some(grand_total_states(context.aggregation)),
            Some(_) => Some(grand_total_states(context.aggregation)),
        }
    } else {
        parent_column_prefix_states(context, &column_key[..column_key.len() - 1])
    }?;
    state_number(states, measure_index, aggregate)
}

pub(crate) fn parent_base_field_total_value(
    context: &ShowAsContext<'_>,
    base_field: &str,
    measure_index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    match show_as_axis(context, base_field)? {
        ShowAsAxis::Row(position) => {
            let row_key = context.row_key?;
            let prefix = row_key.get(..=position)?;
            let states = if prefix.len() == row_key.len() {
                states_for_row_axis_key(context, row_key)
            } else {
                parent_row_prefix_states(context, prefix)
            }?;
            state_number(states, measure_index, aggregate)
        }
        ShowAsAxis::Column(position) => {
            let column_key = context.column_key?;
            let prefix = column_key.get(..=position)?;
            let states = if prefix.len() == column_key.len() {
                states_for_column_axis_key(context, column_key)
            } else {
                parent_column_prefix_states(context, prefix)
            }?;
            state_number(states, measure_index, aggregate)
        }
    }
}

pub(crate) fn parent_row_prefix_states<'a>(
    context: &'a ShowAsContext<'_>,
    row_prefix: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.column_key {
        Some(column_key) if !column_key.is_empty() => {
            subtotal_group_states(context.aggregation, row_prefix, column_key)
        }
        None => row_subtotal_states(context.aggregation, row_prefix),
        Some(_) => row_subtotal_states(context.aggregation, row_prefix),
    }
}

pub(crate) fn parent_column_prefix_states<'a>(
    context: &'a ShowAsContext<'_>,
    column_prefix: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.row_key {
        Some(row_key) if !row_key.is_empty() => {
            subtotal_group_states(context.aggregation, row_key, column_prefix)
        }
        None => column_subtotal_states(context.aggregation, column_prefix),
        Some(_) => column_subtotal_states(context.aggregation, column_prefix),
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) enum ShowAsAxis {
    Row(usize),
    Column(usize),
}

pub(crate) fn show_as_axis(context: &ShowAsContext<'_>, base_field: &str) -> Option<ShowAsAxis> {
    let field_index = context.snapshot.field_index(base_field)?;
    context
        .plan
        .row_indexes
        .iter()
        .position(|index| *index == field_index)
        .map(ShowAsAxis::Row)
        .or_else(|| {
            context
                .plan
                .column_indexes
                .iter()
                .position(|index| *index == field_index)
                .map(ShowAsAxis::Column)
        })
}

pub(crate) fn base_item_value(
    context: &ShowAsContext<'_>,
    base_field: &str,
    base_item: &PivotValue,
    measure_index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    let field_index = context.snapshot.field_index(base_field)?;
    let base_id = context.snapshot.columns[field_index].id_for_value(base_item)?;
    let state = match show_as_axis(context, base_field)? {
        ShowAsAxis::Row(position) => {
            let mut key = context.row_key?.clone();
            *key.get_mut(position)? = base_id;
            states_for_row_axis_key(context, &key)?
        }
        ShowAsAxis::Column(position) => {
            let mut key = context.column_key?.clone();
            *key.get_mut(position)? = base_id;
            states_for_column_axis_key(context, &key)?
        }
    };
    state_number(state, measure_index, aggregate)
}

pub(crate) fn running_total_value(
    context: &ShowAsContext<'_>,
    base_field: &str,
    measure_index: usize,
    aggregate: PivotAggregate,
) -> Option<f64> {
    let axis = show_as_axis(context, base_field)?;
    let mut total = 0.0;
    let mut found = false;
    match axis {
        ShowAsAxis::Row(position) => {
            let current = context.row_key?;
            for key in &context.aggregation.row_order {
                if !same_peer_key(key, current, position) {
                    continue;
                }
                if let Some(value) = states_for_row_axis_key(context, key)
                    .and_then(|states| state_number(states, measure_index, aggregate))
                {
                    total += value;
                }
                if key == current {
                    found = true;
                    break;
                }
            }
        }
        ShowAsAxis::Column(position) => {
            let current = context.column_key?;
            for key in &context.aggregation.column_order {
                if !same_peer_key(key, current, position) {
                    continue;
                }
                if let Some(value) = states_for_column_axis_key(context, key)
                    .and_then(|states| state_number(states, measure_index, aggregate))
                {
                    total += value;
                }
                if key == current {
                    found = true;
                    break;
                }
            }
        }
    }
    found.then_some(total)
}

pub(crate) fn rank_value(
    context: &ShowAsContext<'_>,
    base_field: &str,
    measure_index: usize,
    aggregate: PivotAggregate,
    ascending: bool,
) -> Option<f64> {
    let axis = show_as_axis(context, base_field)?;
    let current_value = match axis {
        ShowAsAxis::Row(_) => states_for_row_axis_key(context, context.row_key?)?,
        ShowAsAxis::Column(_) => states_for_column_axis_key(context, context.column_key?)?,
    };
    let current_value = state_number(current_value, measure_index, aggregate)?;

    let mut rank = 1u64;
    match axis {
        ShowAsAxis::Row(position) => {
            let current = context.row_key?;
            for key in &context.aggregation.row_order {
                if !same_peer_key(key, current, position) {
                    continue;
                }
                let Some(value) = states_for_row_axis_key(context, key)
                    .and_then(|states| state_number(states, measure_index, aggregate))
                else {
                    continue;
                };
                if rank_precedes(value, current_value, ascending) {
                    rank += 1;
                }
            }
        }
        ShowAsAxis::Column(position) => {
            let current = context.column_key?;
            for key in &context.aggregation.column_order {
                if !same_peer_key(key, current, position) {
                    continue;
                }
                let Some(value) = states_for_column_axis_key(context, key)
                    .and_then(|states| state_number(states, measure_index, aggregate))
                else {
                    continue;
                };
                if rank_precedes(value, current_value, ascending) {
                    rank += 1;
                }
            }
        }
    }
    Some(rank as f64)
}

pub(crate) fn states_for_row_axis_key<'a>(
    context: &'a ShowAsContext<'_>,
    row_key: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.column_key {
        Some(column_key) => context.aggregation.groups.get(&GroupKey {
            rows: row_key.to_vec(),
            columns: column_key.clone(),
        }),
        None => row_total_states(context.aggregation, row_key),
    }
}

pub(crate) fn states_for_column_axis_key<'a>(
    context: &'a ShowAsContext<'_>,
    column_key: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.row_key {
        Some(row_key) => context.aggregation.groups.get(&GroupKey {
            rows: row_key.clone(),
            columns: column_key.to_vec(),
        }),
        None => column_total_states(context.aggregation, column_key),
    }
}

pub(crate) fn same_peer_key(candidate: &[u32], current: &[u32], base_position: usize) -> bool {
    candidate.len() == current.len()
        && candidate
            .iter()
            .zip(current.iter())
            .enumerate()
            .all(|(index, (left, right))| index == base_position || left == right)
}

pub(crate) fn rank_precedes(value: f64, current: f64, ascending: bool) -> bool {
    if ascending {
        value < current
    } else {
        value > current
    }
}
