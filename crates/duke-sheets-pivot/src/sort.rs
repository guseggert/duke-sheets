use crate::aggregate::*;
use crate::prelude::*;
use crate::snapshot::*;

pub(crate) fn sort_key_order(
    order: &mut [Vec<u32>],
    field_indexes: &[usize],
    fields: &[PivotField],
    sort_measure_indexes: &[Option<usize>],
    totals: &AHashMap<Vec<u32>, Vec<AggregateState>>,
    measures: &[PivotMeasure],
    snapshot: &SourceSnapshot,
) {
    if fields
        .iter()
        .all(|field| matches!(field.sort, PivotSort::None))
    {
        return;
    }

    let measure_sort_totals =
        measure_sort_prefix_totals(totals, fields, sort_measure_indexes, measures);

    order.sort_by(|a, b| {
        compare_encoded_key(
            a,
            b,
            field_indexes,
            fields,
            sort_measure_indexes,
            &measure_sort_totals,
            measures,
            snapshot,
        )
    });
}

pub(crate) type MeasureSortPrefixTotals = Vec<Option<AHashMap<Vec<u32>, Vec<AggregateState>>>>;

pub(crate) fn measure_sort_prefix_totals(
    totals: &AHashMap<Vec<u32>, Vec<AggregateState>>,
    fields: &[PivotField],
    sort_measure_indexes: &[Option<usize>],
    measures: &[PivotMeasure],
) -> MeasureSortPrefixTotals {
    let mut prefix_totals = (0..fields.len()).map(|_| None).collect::<Vec<_>>();
    for (field_position, field) in fields.iter().enumerate() {
        if matches!(field.sort, PivotSort::None)
            || sort_measure_indexes
                .get(field_position)
                .and_then(|index| *index)
                .is_none()
        {
            continue;
        }

        let mut scoped = AHashMap::<Vec<u32>, Vec<AggregateState>>::new();
        for (key, states) in totals {
            if key.len() <= field_position {
                continue;
            }
            let entry = scoped
                .entry(key[..=field_position].to_vec())
                .or_insert_with(|| default_states(measures));
            merge_state_slices(entry, states, measures);
        }
        prefix_totals[field_position] = Some(scoped);
    }
    prefix_totals
}

pub(crate) fn order_positions(order: &[Vec<u32>]) -> AHashMap<Vec<u32>, usize> {
    order
        .iter()
        .cloned()
        .enumerate()
        .map(|(index, key)| (key, index))
        .collect()
}

pub(crate) fn compile_axis_sort_measure_indexes(
    pivot_name: &str,
    fields: &[PivotField],
    measures: &[PivotMeasure],
) -> Result<Vec<Option<usize>>> {
    fields
        .iter()
        .map(|field| {
            field
                .sort_by_measure
                .as_ref()
                .map(|sort_measure| {
                    measures
                        .iter()
                        .position(|measure| {
                            pivot_measure_matches_sort_target(measure, sort_measure)
                        })
                        .ok_or_else(|| {
                            Error::other(format!(
                                "pivot table {pivot_name} sorts field {} by an unknown measure",
                                field.field.name
                            ))
                        })
                })
                .transpose()
        })
        .collect()
}

pub(crate) fn pivot_measure_matches_sort_target(
    measure: &PivotMeasure,
    target: &PivotMeasure,
) -> bool {
    measure.field.name.eq_ignore_ascii_case(&target.field.name)
        && measure.aggregate == target.aggregate
        && target
            .name
            .as_ref()
            .is_none_or(|name| measure.name.as_ref() == Some(name))
}

pub(crate) fn compare_encoded_key(
    left: &[u32],
    right: &[u32],
    field_indexes: &[usize],
    fields: &[PivotField],
    sort_measure_indexes: &[Option<usize>],
    totals: &MeasureSortPrefixTotals,
    measures: &[PivotMeasure],
    snapshot: &SourceSnapshot,
) -> Ordering {
    for (index, field_index) in field_indexes.iter().enumerate() {
        if left.get(index) == right.get(index) {
            continue;
        }

        let sort = fields
            .get(index)
            .map(|field| field.sort)
            .unwrap_or(PivotSort::Ascending);
        if matches!(sort, PivotSort::None) {
            return Ordering::Equal;
        }

        let ordering = sort_measure_indexes
            .get(index)
            .and_then(|measure_index| *measure_index)
            .map(|measure_index| {
                compare_measure_sort_values(left, right, index, totals, measures, measure_index)
            })
            .unwrap_or(Ordering::Equal)
            .then_with(|| {
                let Some(left_id) = left.get(index).copied() else {
                    return Ordering::Equal;
                };
                let Some(right_id) = right.get(index).copied() else {
                    return Ordering::Equal;
                };
                compare_pivot_values(
                    snapshot.value_by_id(*field_index, left_id),
                    snapshot.value_by_id(*field_index, right_id),
                )
            });

        if ordering != Ordering::Equal {
            return match sort {
                PivotSort::Ascending => ordering,
                PivotSort::Descending => ordering.reverse(),
                PivotSort::None => Ordering::Equal,
            };
        }
    }

    Ordering::Equal
}

pub(crate) fn compare_measure_sort_values(
    left: &[u32],
    right: &[u32],
    field_position: usize,
    totals: &MeasureSortPrefixTotals,
    measures: &[PivotMeasure],
    measure_index: usize,
) -> Ordering {
    let Some(totals) = totals
        .get(field_position)
        .and_then(|totals| totals.as_ref())
    else {
        return Ordering::Equal;
    };
    if left.len() <= field_position || right.len() <= field_position {
        return Ordering::Equal;
    }
    let aggregate = measures[measure_index].aggregate;
    let left = totals
        .get(&left[..=field_position])
        .and_then(|states| states.get(measure_index))
        .and_then(|state| state.finalize_number(aggregate));
    let right = totals
        .get(&right[..=field_position])
        .and_then(|states| states.get(measure_index))
        .and_then(|state| state.finalize_number(aggregate));
    compare_optional_numbers(left, right)
}

pub(crate) fn compare_optional_numbers(left: Option<f64>, right: Option<f64>) -> Ordering {
    match (left, right) {
        (Some(left), Some(right)) => left.partial_cmp(&right).unwrap_or(Ordering::Equal),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => Ordering::Equal,
    }
}

pub(crate) fn compare_pivot_values(left: &PivotValue, right: &PivotValue) -> Ordering {
    let left_rank = pivot_value_rank(left);
    let right_rank = pivot_value_rank(right);
    left_rank
        .cmp(&right_rank)
        .then_with(|| match (left, right) {
            (PivotValue::Blank, PivotValue::Blank) => Ordering::Equal,
            (PivotValue::Boolean(left), PivotValue::Boolean(right)) => left.cmp(right),
            (PivotValue::Number(left), PivotValue::Number(right)) => {
                left.partial_cmp(right).unwrap_or(Ordering::Equal)
            }
            (PivotValue::String(left), PivotValue::String(right)) => {
                left.to_lowercase().cmp(&right.to_lowercase())
            }
            (PivotValue::Error(left), PivotValue::Error(right)) => left.code().cmp(&right.code()),
            _ => Ordering::Equal,
        })
}

pub(crate) fn pivot_value_rank(value: &PivotValue) -> u8 {
    match value {
        PivotValue::Blank => 0,
        PivotValue::Boolean(_) => 1,
        PivotValue::Number(_) => 2,
        PivotValue::String(_) => 3,
        PivotValue::Error(_) => 4,
    }
}
