use crate::aggregate::*;
use crate::api::*;
use crate::prelude::*;
use crate::runtime_cache::*;
use crate::snapshot::*;

#[derive(Debug, Clone)]
pub(crate) enum CompiledFilter {
    Items {
        field_index: usize,
        allowed_ids: AHashSet<u32>,
    },
    Label {
        field_index: usize,
        operator: PivotFilterOperator,
        value: String,
    },
    LabelBetween {
        field_index: usize,
        lower: String,
        upper: String,
        not_between: bool,
    },
    Date {
        field_index: usize,
        operator: PivotFilterOperator,
        value: f64,
    },
    DateBetween {
        field_index: usize,
        start: f64,
        end: f64,
        not_between: bool,
    },
    DatePeriod {
        field_index: usize,
        period: CompiledDatePeriod,
    },
}

#[derive(Debug, Clone, Copy)]
pub(crate) enum CompiledDatePeriod {
    Range {
        start: f64,
        end: f64,
    },
    Month {
        month: u8,
        date_system: DateSystem,
    },
    Quarter {
        quarter: u8,
        date_system: DateSystem,
    },
}

impl CompiledFilter {
    pub(crate) fn compile(
        filter: &PivotFilter,
        snapshot: &SourceSnapshot,
        pivot_name: &str,
        filter_baselines: &PivotFilterBaselines,
        options: &PivotRefreshOptions,
        date_system: DateSystem,
    ) -> Result<Self> {
        match filter {
            PivotFilter::FieldItems {
                field,
                allowed_items,
            } => {
                let field_index = snapshot.required_field_index(&field.name, pivot_name)?;
                let allowed_ids = allowed_items
                    .iter()
                    .filter_map(|value| snapshot.columns[field_index].id_for_value(value))
                    .collect::<AHashSet<_>>();
                let allowed_ids = allowed_ids_with_new_items(
                    allowed_ids,
                    snapshot,
                    field_index,
                    filter_baselines.known_items(&field.name),
                );
                Ok(Self::Items {
                    field_index,
                    allowed_ids,
                })
            }
            PivotFilter::Label {
                field,
                operator,
                value,
            } => Ok(Self::Label {
                field_index: snapshot.required_field_index(&field.name, pivot_name)?,
                operator: *operator,
                value: value.to_lowercase(),
            }),
            PivotFilter::LabelBetween {
                field,
                start,
                end,
                not_between,
            } => {
                let start = start.to_lowercase();
                let end = end.to_lowercase();
                let (lower, upper) = if start <= end {
                    (start, end)
                } else {
                    (end, start)
                };
                Ok(Self::LabelBetween {
                    field_index: snapshot.required_field_index(&field.name, pivot_name)?,
                    lower,
                    upper,
                    not_between: *not_between,
                })
            }
            PivotFilter::Date {
                field,
                operator,
                value,
            } => Ok(Self::Date {
                field_index: snapshot.required_field_index(&field.name, pivot_name)?,
                operator: *operator,
                value: *value,
            }),
            PivotFilter::DateBetween {
                field,
                start,
                end,
                not_between,
            } => Ok(Self::DateBetween {
                field_index: snapshot.required_field_index(&field.name, pivot_name)?,
                start: *start,
                end: *end,
                not_between: *not_between,
            }),
            PivotFilter::DatePeriod { field, period } => Ok(Self::DatePeriod {
                field_index: snapshot.required_field_index(&field.name, pivot_name)?,
                period: compile_date_period(*period, options, date_system, pivot_name)?,
            }),
            PivotFilter::Value { .. }
            | PivotFilter::ValueBetween { .. }
            | PivotFilter::TopN { .. } => Err(Error::other(format!(
                "pivot table {pivot_name} tried to compile an aggregate filter as a row filter"
            ))),
            PivotFilter::Unsupported { kind, .. } => Err(Error::other(format!(
                "pivot table {pivot_name} contains unsupported filter: {kind}"
            ))),
        }
    }

    pub(crate) fn matches(&self, snapshot: &SourceSnapshot, row: usize) -> bool {
        match self {
            Self::Items {
                field_index,
                allowed_ids,
            } => allowed_ids.contains(&snapshot.columns[*field_index].id_at(row)),
            Self::Label {
                field_index,
                operator,
                value,
            } => {
                let actual = snapshot.value(row, *field_index).to_string();
                label_filter_matches(&actual, *operator, value)
            }
            Self::LabelBetween {
                field_index,
                lower,
                upper,
                not_between,
            } => {
                let actual = snapshot.value(row, *field_index).to_string();
                label_between_filter_matches(&actual, lower, upper, *not_between)
            }
            Self::Date {
                field_index,
                operator,
                value,
            } => pivot_number(snapshot.value(row, *field_index))
                .is_some_and(|actual| date_filter_matches(actual, *operator, *value)),
            Self::DateBetween {
                field_index,
                start,
                end,
                not_between,
            } => pivot_number(snapshot.value(row, *field_index)).is_some_and(|actual| {
                date_between_filter_matches(actual, *start, *end, *not_between)
            }),
            Self::DatePeriod {
                field_index,
                period,
            } => pivot_number(snapshot.value(row, *field_index))
                .is_some_and(|actual| date_period_filter_matches(actual, *period)),
        }
    }

    pub(crate) fn field_index(&self) -> usize {
        match self {
            Self::Items { field_index, .. }
            | Self::Label { field_index, .. }
            | Self::LabelBetween { field_index, .. }
            | Self::Date { field_index, .. }
            | Self::DateBetween { field_index, .. }
            | Self::DatePeriod { field_index, .. } => *field_index,
        }
    }

    pub(crate) fn targets_axis(&self, row_indexes: &[usize], column_indexes: &[usize]) -> bool {
        let field_index = self.field_index();
        row_indexes.contains(&field_index) || column_indexes.contains(&field_index)
    }

    pub(crate) fn allows_item(
        &self,
        snapshot: &SourceSnapshot,
        field_index: usize,
        item_id: u32,
    ) -> bool {
        match self {
            Self::Items {
                field_index: filter_index,
                allowed_ids,
            } if *filter_index == field_index => allowed_ids.contains(&item_id),
            Self::Label {
                field_index: filter_index,
                operator,
                value,
            } if *filter_index == field_index => {
                let actual = snapshot.value_by_id(field_index, item_id).to_string();
                label_filter_matches(&actual, *operator, value)
            }
            Self::LabelBetween {
                field_index: filter_index,
                lower,
                upper,
                not_between,
            } if *filter_index == field_index => {
                let actual = snapshot.value_by_id(field_index, item_id).to_string();
                label_between_filter_matches(&actual, lower, upper, *not_between)
            }
            Self::Date {
                field_index: filter_index,
                operator,
                value,
            } if *filter_index == field_index => {
                pivot_number(snapshot.value_by_id(field_index, item_id))
                    .is_some_and(|actual| date_filter_matches(actual, *operator, *value))
            }
            Self::DateBetween {
                field_index: filter_index,
                start,
                end,
                not_between,
            } if *filter_index == field_index => {
                pivot_number(snapshot.value_by_id(field_index, item_id)).is_some_and(|actual| {
                    date_between_filter_matches(actual, *start, *end, *not_between)
                })
            }
            Self::DatePeriod {
                field_index: filter_index,
                period,
            } if *filter_index == field_index => {
                pivot_number(snapshot.value_by_id(field_index, item_id))
                    .is_some_and(|actual| date_period_filter_matches(actual, *period))
            }
            _ => true,
        }
    }
}

pub(crate) fn allowed_ids_with_new_items(
    mut allowed_ids: AHashSet<u32>,
    snapshot: &SourceSnapshot,
    field_index: usize,
    known_items: Option<&AHashSet<PivotValue>>,
) -> AHashSet<u32> {
    let Some(known_items) = known_items else {
        return allowed_ids;
    };

    for (item_id, value) in snapshot.columns[field_index].dictionary.iter().enumerate() {
        if !known_items.contains(value) {
            allowed_ids.insert(item_id as u32);
        }
    }
    allowed_ids
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum AggregateFilterAxis {
    Row,
    Column,
}

#[derive(Debug, Default, Clone)]
pub(crate) struct AxisItemRestrictions {
    pub(crate) rows: AHashMap<usize, AHashSet<u32>>,
    pub(crate) columns: AHashMap<usize, AHashSet<u32>>,
}

impl AxisItemRestrictions {
    pub(crate) fn restrict(
        &mut self,
        axis: AggregateFilterAxis,
        field_position: usize,
        allowed_item_ids: &AHashSet<u32>,
    ) {
        let target = match axis {
            AggregateFilterAxis::Row => &mut self.rows,
            AggregateFilterAxis::Column => &mut self.columns,
        };
        target
            .entry(field_position)
            .and_modify(|existing| existing.retain(|id| allowed_item_ids.contains(id)))
            .or_insert_with(|| allowed_item_ids.clone());
    }

    pub(crate) fn allows(
        &self,
        axis: AggregateFilterAxis,
        field_position: usize,
        item_id: u32,
    ) -> bool {
        let source = match axis {
            AggregateFilterAxis::Row => &self.rows,
            AggregateFilterAxis::Column => &self.columns,
        };
        source
            .get(&field_position)
            .map(|allowed| allowed.contains(&item_id))
            .unwrap_or(true)
    }
}

#[derive(Debug, Clone)]
pub(crate) enum CompiledAggregateFilter {
    Value {
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
        operator: PivotFilterOperator,
        value: f64,
    },
    ValueBetween {
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
        start: f64,
        end: f64,
        not_between: bool,
    },
    TopN {
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
        n: u32,
        top: bool,
        percent: bool,
    },
}

impl CompiledAggregateFilter {
    pub(crate) fn compile(
        filter: &PivotFilter,
        snapshot: &SourceSnapshot,
        pivot_name: &str,
        row_indexes: &[usize],
        column_indexes: &[usize],
        measures: &[PivotMeasure],
    ) -> Result<Self> {
        match filter {
            PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } => {
                let field_index = snapshot.required_field_index(&field.name, pivot_name)?;
                let (axis, field_position) = aggregate_filter_axis(
                    pivot_name,
                    &field.name,
                    field_index,
                    row_indexes,
                    column_indexes,
                )?;
                let measure_index = measure_index_for_filter(pivot_name, measures, measure)?;
                Ok(Self::Value {
                    axis,
                    field_position,
                    measure_index,
                    aggregate: measure.aggregate,
                    operator: *operator,
                    value: *value,
                })
            }
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                not_between,
            } => {
                let field_index = snapshot.required_field_index(&field.name, pivot_name)?;
                let (axis, field_position) = aggregate_filter_axis(
                    pivot_name,
                    &field.name,
                    field_index,
                    row_indexes,
                    column_indexes,
                )?;
                let measure_index = measure_index_for_filter(pivot_name, measures, measure)?;
                Ok(Self::ValueBetween {
                    axis,
                    field_position,
                    measure_index,
                    aggregate: measure.aggregate,
                    start: *start,
                    end: *end,
                    not_between: *not_between,
                })
            }
            PivotFilter::TopN {
                field,
                measure,
                n,
                top,
                percent,
            } => {
                let field_index = snapshot.required_field_index(&field.name, pivot_name)?;
                let (axis, field_position) = aggregate_filter_axis(
                    pivot_name,
                    &field.name,
                    field_index,
                    row_indexes,
                    column_indexes,
                )?;
                let measure_index = measure_index_for_filter(pivot_name, measures, measure)?;
                Ok(Self::TopN {
                    axis,
                    field_position,
                    measure_index,
                    aggregate: measure.aggregate,
                    n: *n,
                    top: *top,
                    percent: *percent,
                })
            }
            _ => Err(Error::other(format!(
                "pivot table {pivot_name} tried to compile a row filter as an aggregate filter"
            ))),
        }
    }

    pub(crate) fn axis(&self) -> AggregateFilterAxis {
        match self {
            Self::Value { axis, .. }
            | Self::ValueBetween { axis, .. }
            | Self::TopN { axis, .. } => *axis,
        }
    }

    pub(crate) fn field_position(&self) -> usize {
        match self {
            Self::Value { field_position, .. }
            | Self::ValueBetween { field_position, .. }
            | Self::TopN { field_position, .. } => *field_position,
        }
    }

    pub(crate) fn allowed_item_ids(&self, aggregation: &PivotAggregation) -> AHashSet<u32> {
        let item_states = aggregation.item_states_for_filter(
            self.axis(),
            self.field_position(),
            self.measure_index(),
            self.aggregate(),
        );
        match self {
            Self::Value {
                operator, value, ..
            } => item_states
                .into_iter()
                .filter_map(|(item_id, state)| {
                    let actual = state.finalize_number(self.aggregate())?;
                    numeric_filter_matches(actual, *operator, *value).then_some(item_id)
                })
                .collect(),
            Self::ValueBetween {
                start,
                end,
                not_between,
                ..
            } => item_states
                .into_iter()
                .filter_map(|(item_id, state)| {
                    let actual = state.finalize_number(self.aggregate())?;
                    number_between_filter_matches(actual, *start, *end, *not_between)
                        .then_some(item_id)
                })
                .collect(),
            Self::TopN {
                n, top, percent, ..
            } => top_n_item_ids(item_states, self.aggregate(), *n, *top, *percent),
        }
    }

    pub(crate) fn measure_index(&self) -> usize {
        match self {
            Self::Value { measure_index, .. }
            | Self::ValueBetween { measure_index, .. }
            | Self::TopN { measure_index, .. } => *measure_index,
        }
    }

    pub(crate) fn aggregate(&self) -> PivotAggregate {
        match self {
            Self::Value { aggregate, .. }
            | Self::ValueBetween { aggregate, .. }
            | Self::TopN { aggregate, .. } => *aggregate,
        }
    }
}

pub(crate) fn aggregate_filter_axis(
    pivot_name: &str,
    field_name: &str,
    field_index: usize,
    row_indexes: &[usize],
    column_indexes: &[usize],
) -> Result<(AggregateFilterAxis, usize)> {
    row_indexes
        .iter()
        .position(|index| *index == field_index)
        .map(|position| (AggregateFilterAxis::Row, position))
        .or_else(|| {
            column_indexes
                .iter()
                .position(|index| *index == field_index)
                .map(|position| (AggregateFilterAxis::Column, position))
        })
        .ok_or_else(|| {
            Error::other(format!(
                "pivot table {pivot_name} uses aggregate filter field {field_name}, but that field is not on a row or column axis"
            ))
        })
}

pub(crate) fn measure_index_for_filter(
    pivot_name: &str,
    measures: &[PivotMeasure],
    filter_measure: &PivotMeasure,
) -> Result<usize> {
    measures
        .iter()
        .position(|measure| {
            measure.field.name.eq_ignore_ascii_case(&filter_measure.field.name)
                && measure.aggregate == filter_measure.aggregate
                && match filter_measure.name.as_ref() {
                    Some(name) => measure
                        .name
                        .as_ref()
                        .map(|candidate| candidate.eq_ignore_ascii_case(name))
                        .unwrap_or_else(|| measure.caption().eq_ignore_ascii_case(name)),
                    None => true,
                }
        })
        .ok_or_else(|| {
            Error::other(format!(
                "pivot table {pivot_name} uses aggregate filter measure {}, but that measure is not in the pivot",
                filter_measure.caption()
            ))
        })
}

pub(crate) fn label_filter_matches(
    actual: &str,
    operator: PivotFilterOperator,
    expected: &str,
) -> bool {
    let actual_folded = actual.to_lowercase();
    match operator {
        PivotFilterOperator::Equals => actual_folded.as_str() == expected,
        PivotFilterOperator::NotEquals => actual_folded.as_str() != expected,
        PivotFilterOperator::LessThan => actual_folded.as_str() < expected,
        PivotFilterOperator::LessThanOrEqual => actual_folded.as_str() <= expected,
        PivotFilterOperator::GreaterThan => actual_folded.as_str() > expected,
        PivotFilterOperator::GreaterThanOrEqual => actual_folded.as_str() >= expected,
        PivotFilterOperator::BeginsWith => actual_folded.starts_with(expected),
        PivotFilterOperator::DoesNotBeginWith => !actual_folded.starts_with(expected),
        PivotFilterOperator::EndsWith => actual_folded.ends_with(expected),
        PivotFilterOperator::DoesNotEndWith => !actual_folded.ends_with(expected),
        PivotFilterOperator::Contains => actual_folded.contains(expected),
        PivotFilterOperator::DoesNotContain => !actual_folded.contains(expected),
    }
}

pub(crate) fn label_between_filter_matches(
    actual: &str,
    lower: &str,
    upper: &str,
    not_between: bool,
) -> bool {
    let actual = actual.to_lowercase();
    let between = actual.as_str() >= lower && actual.as_str() <= upper;
    if not_between {
        !between
    } else {
        between
    }
}

pub(crate) fn numeric_filter_matches(
    actual: f64,
    operator: PivotFilterOperator,
    expected: f64,
) -> bool {
    match operator {
        PivotFilterOperator::Equals => actual == expected,
        PivotFilterOperator::NotEquals => actual != expected,
        PivotFilterOperator::LessThan => actual < expected,
        PivotFilterOperator::LessThanOrEqual => actual <= expected,
        PivotFilterOperator::GreaterThan => actual > expected,
        PivotFilterOperator::GreaterThanOrEqual => actual >= expected,
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => false,
    }
}

pub(crate) fn number_between_filter_matches(
    actual: f64,
    start: f64,
    end: f64,
    not_between: bool,
) -> bool {
    if !actual.is_finite() || !start.is_finite() || !end.is_finite() {
        return false;
    }
    let lower = start.min(end);
    let upper = start.max(end);
    let between = actual >= lower && actual <= upper;
    if not_between {
        !between
    } else {
        between
    }
}

pub(crate) fn date_filter_matches(
    actual: f64,
    operator: PivotFilterOperator,
    expected: f64,
) -> bool {
    if !actual.is_finite() || !expected.is_finite() {
        return false;
    }
    numeric_filter_matches(actual, operator, expected)
}

pub(crate) fn date_between_filter_matches(
    actual: f64,
    start: f64,
    end: f64,
    not_between: bool,
) -> bool {
    number_between_filter_matches(actual, start, end, not_between)
}

pub(crate) fn compile_date_period(
    period: PivotDatePeriod,
    options: &PivotRefreshOptions,
    date_system: DateSystem,
    pivot_name: &str,
) -> Result<CompiledDatePeriod> {
    match period {
        PivotDatePeriod::Month(month) if (1..=12).contains(&month) => {
            Ok(CompiledDatePeriod::Month { month, date_system })
        }
        PivotDatePeriod::Quarter(quarter) if (1..=4).contains(&quarter) => {
            Ok(CompiledDatePeriod::Quarter {
                quarter,
                date_system,
            })
        }
        PivotDatePeriod::Month(month) => Err(Error::other(format!(
            "pivot table {pivot_name} uses invalid date period month {month}; expected 1..=12"
        ))),
        PivotDatePeriod::Quarter(quarter) => Err(Error::other(format!(
            "pivot table {pivot_name} uses invalid date period quarter {quarter}; expected 1..=4"
        ))),
        _ => {
            let today = options.today.ok_or_else(|| {
                Error::other(format!(
                    "pivot table {pivot_name} uses a relative date period filter but refresh options did not provide today"
                ))
            })?;
            if !today.is_finite() {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} uses a relative date period filter with a non-finite refresh date"
                )));
            }
            relative_date_period_range(period, today, date_system)
                .map(|(start, end)| CompiledDatePeriod::Range { start, end })
                .ok_or_else(|| {
                    Error::other(format!(
                        "pivot table {pivot_name} could not evaluate relative date period filter"
                    ))
                })
        }
    }
}

pub(crate) fn date_period_filter_matches(actual: f64, period: CompiledDatePeriod) -> bool {
    if !actual.is_finite() {
        return false;
    }
    match period {
        CompiledDatePeriod::Range { start, end } => {
            date_between_filter_matches(actual, start, end, false)
        }
        CompiledDatePeriod::Month { month, date_system } => serial_to_date(actual, date_system)
            .map(|(_, actual_month, _)| actual_month == u32::from(month))
            .unwrap_or(false),
        CompiledDatePeriod::Quarter {
            quarter,
            date_system,
        } => serial_to_date(actual, date_system)
            .map(|(_, actual_month, _)| ((actual_month - 1) / 3 + 1) == u32::from(quarter))
            .unwrap_or(false),
    }
}

pub(crate) fn relative_date_period_range(
    period: PivotDatePeriod,
    today: f64,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    let (year, month, day) = serial_to_date(today, date_system)?;
    let today = date_to_serial(year, month, day, date_system);
    match period {
        PivotDatePeriod::Tomorrow => Some((today + 1.0, today + 1.0)),
        PivotDatePeriod::Today => Some((today, today)),
        PivotDatePeriod::Yesterday => Some((today - 1.0, today - 1.0)),
        PivotDatePeriod::NextWeek => week_range(today + 7.0, date_system),
        PivotDatePeriod::ThisWeek => week_range(today, date_system),
        PivotDatePeriod::LastWeek => week_range(today - 7.0, date_system),
        PivotDatePeriod::NextMonth => {
            let (year, month) = add_months(year, month, 1);
            Some(month_range(year, month, date_system))
        }
        PivotDatePeriod::ThisMonth => Some(month_range(year, month, date_system)),
        PivotDatePeriod::LastMonth => {
            let (year, month) = add_months(year, month, -1);
            Some(month_range(year, month, date_system))
        }
        PivotDatePeriod::NextQuarter => {
            let (year, month) = add_months(year, quarter_start_month(month), 3);
            Some(quarter_range(year, month, date_system))
        }
        PivotDatePeriod::ThisQuarter => {
            Some(quarter_range(year, quarter_start_month(month), date_system))
        }
        PivotDatePeriod::LastQuarter => {
            let (year, month) = add_months(year, quarter_start_month(month), -3);
            Some(quarter_range(year, month, date_system))
        }
        PivotDatePeriod::NextYear => Some(year_range(year + 1, date_system)),
        PivotDatePeriod::ThisYear => Some(year_range(year, date_system)),
        PivotDatePeriod::LastYear => Some(year_range(year - 1, date_system)),
        PivotDatePeriod::YearToDate => Some((date_to_serial(year, 1, 1, date_system), today)),
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => None,
    }
}

pub(crate) fn week_range(serial: f64, date_system: DateSystem) -> Option<(f64, f64)> {
    let (year, month, day) = serial_to_date(serial, date_system)?;
    let day_serial = date_to_serial(year, month, day, date_system);
    let weekday = serial_to_weekday(day_serial, date_system);
    let start = day_serial - f64::from(weekday - 1);
    Some((start, start + 6.0))
}

pub(crate) fn month_range(year: i32, month: u32, date_system: DateSystem) -> (f64, f64) {
    (
        date_to_serial(year, month, 1, date_system),
        date_to_serial(year, month, days_in_month(year, month), date_system),
    )
}

pub(crate) fn quarter_range(year: i32, start_month: u32, date_system: DateSystem) -> (f64, f64) {
    let (end_year, end_month) = add_months(year, start_month, 2);
    (
        date_to_serial(year, start_month, 1, date_system),
        date_to_serial(
            end_year,
            end_month,
            days_in_month(end_year, end_month),
            date_system,
        ),
    )
}

pub(crate) fn year_range(year: i32, date_system: DateSystem) -> (f64, f64) {
    (
        date_to_serial(year, 1, 1, date_system),
        date_to_serial(year, 12, 31, date_system),
    )
}

pub(crate) fn quarter_start_month(month: u32) -> u32 {
    ((month - 1) / 3) * 3 + 1
}

pub(crate) fn add_months(year: i32, month: u32, delta: i32) -> (i32, u32) {
    let month_index = year * 12 + month as i32 - 1 + delta;
    let year = month_index.div_euclid(12);
    let month = month_index.rem_euclid(12) + 1;
    (year, month as u32)
}

pub(crate) fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year == 1900 => 29,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 31,
    }
}

pub(crate) fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

pub(crate) fn workbook_date_system(date_1904: bool) -> DateSystem {
    if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    }
}

pub(crate) fn top_n_item_ids(
    item_states: AHashMap<u32, AggregateState>,
    aggregate: PivotAggregate,
    n: u32,
    top: bool,
    percent: bool,
) -> AHashSet<u32> {
    if n == 0 {
        return AHashSet::new();
    }

    let mut values = item_states
        .into_iter()
        .filter_map(|(item_id, state)| {
            state
                .finalize_number(aggregate)
                .map(|value| (item_id, value))
        })
        .collect::<Vec<_>>();
    values.sort_by(|(left_id, left), (right_id, right)| {
        left.partial_cmp(right)
            .unwrap_or(Ordering::Equal)
            .then_with(|| left_id.cmp(right_id))
    });
    if top {
        values.reverse();
    }

    let take = if percent {
        ((values.len() as f64) * (n as f64 / 100.0)).ceil() as usize
    } else {
        n as usize
    }
    .min(values.len());

    values
        .into_iter()
        .take(take)
        .map(|(item_id, _)| item_id)
        .collect()
}
