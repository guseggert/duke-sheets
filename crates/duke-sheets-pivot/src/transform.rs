use crate::prelude::*;
use crate::runtime_cache::*;
use crate::snapshot::*;

pub(crate) fn transformed_snapshot_for_pivot(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    pivot: &PivotTable,
    source_snapshot: CachedSourceSnapshot,
    date_1904: bool,
    cache: &mut PivotRuntimeCache,
) -> Result<Arc<SourceSnapshot>> {
    if pivot.calculated_fields.is_empty()
        && pivot.groupings.is_empty()
        && pivot.calculated_items.is_empty()
    {
        return Ok(source_snapshot.snapshot);
    }

    let calculated_fields_use_workbook_refs =
        calculated_fields_use_workbook_refs(&pivot.name, &pivot.calculated_fields)?;
    let use_cache = !calculated_fields_use_workbook_refs;
    let cache_key = use_cache.then(|| {
        TransformedSnapshotCacheKey::new(
            source_snapshot.key,
            &pivot.calculated_fields,
            &pivot.groupings,
            &pivot.calculated_items,
            date_1904,
        )
    });
    if let Some(cache_key) = &cache_key {
        if let Some(snapshot) = cache.transformed_snapshots.get(cache_key) {
            return Ok(Arc::clone(snapshot));
        }
    }

    let workbook_context =
        calculated_fields_use_workbook_refs.then_some(CalculatedWorkbookContext {
            workbook,
            sheet_index: pivot_sheet_index,
        });
    let calculated_snapshot = if pivot.calculated_fields.is_empty() {
        source_snapshot.snapshot
    } else {
        Arc::new(source_snapshot.snapshot.apply_calculated_fields(
            &pivot.name,
            &pivot.calculated_fields,
            workbook_context,
        )?)
    };
    let grouped_snapshot = if pivot.groupings.is_empty() {
        calculated_snapshot
    } else {
        Arc::new(calculated_snapshot.apply_groupings(&pivot.name, &pivot.groupings, date_1904)?)
    };
    let snapshot = if pivot.calculated_items.is_empty() {
        grouped_snapshot
    } else {
        Arc::new(grouped_snapshot.apply_calculated_items(&pivot.name, &pivot.calculated_items)?)
    };
    if let Some(cache_key) = cache_key {
        cache
            .transformed_snapshots
            .insert(cache_key, Arc::clone(&snapshot));
    }
    Ok(snapshot)
}

pub(crate) fn calculated_fields_use_workbook_refs(
    pivot_name: &str,
    calculated_fields: &[PivotCalculatedField],
) -> Result<bool> {
    for field in calculated_fields {
        let ast = parse_calculated_formula(pivot_name, field)?;
        if calculated_formula_expr_uses_workbook_refs(&ast) {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(crate) fn calculated_formula_expr_uses_workbook_refs(expr: &FormulaExpr) -> bool {
    match expr {
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::ExternalRef(_) => true,
        FormulaExpr::BinaryOp { left, right, .. } => {
            calculated_formula_expr_uses_workbook_refs(left)
                || calculated_formula_expr_uses_workbook_refs(right)
        }
        FormulaExpr::UnaryOp { operand, .. } => calculated_formula_expr_uses_workbook_refs(operand),
        FormulaExpr::Function { args, .. } | FormulaExpr::ExternalFunction { args, .. } => {
            args.iter().any(calculated_formula_expr_uses_workbook_refs)
        }
        FormulaExpr::Array(rows) => rows
            .iter()
            .flatten()
            .any(calculated_formula_expr_uses_workbook_refs),
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::Empty
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_) => false,
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct CalculatedWorkbookContext<'a> {
    pub(crate) workbook: &'a Workbook,
    pub(crate) sheet_index: usize,
}

pub(crate) fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
    }
}

pub(crate) fn parse_calculated_formula(
    pivot_name: &str,
    field: &PivotCalculatedField,
) -> Result<FormulaExpr> {
    let formula = field.formula.trim();
    if formula.is_empty() {
        return Err(Error::other(format!(
            "pivot table {pivot_name} calculated field {} has a blank formula",
            field.name
        )));
    }
    let formula = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    parse_formula(&formula).map_err(|error| {
        Error::other(format!(
            "pivot table {pivot_name} calculated field {} formula did not parse: {error}",
            field.name
        ))
    })
}

pub(crate) fn parse_calculated_item_formula(
    pivot_name: &str,
    item: &PivotCalculatedItem,
) -> Result<FormulaExpr> {
    let formula = item.formula.trim();
    if formula.is_empty() {
        return Err(Error::other(format!(
            "pivot table {pivot_name} calculated item {} has a blank formula",
            item.item
        )));
    }
    let formula = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    parse_formula(&formula).map_err(|error| {
        Error::other(format!(
            "pivot table {pivot_name} calculated item {} formula did not parse: {error}",
            item.item
        ))
    })
}

pub(crate) fn field_lookup(headers: &[String]) -> AHashMap<String, usize> {
    headers
        .iter()
        .enumerate()
        .map(|(index, header)| (header.to_lowercase(), index))
        .collect()
}

pub(crate) fn evaluate_calculated_values(
    pivot_name: &str,
    field: &PivotCalculatedField,
    ast: &FormulaExpr,
    columns: &[EncodedColumn],
    row_count: usize,
    lookup: &AHashMap<String, usize>,
    source_name: Option<&str>,
    workbook_context: Option<CalculatedWorkbookContext<'_>>,
) -> Result<Vec<PivotValue>> {
    #[cfg(feature = "parallel")]
    {
        if row_count >= PARALLEL_ROW_THRESHOLD {
            return (0..row_count)
                .into_par_iter()
                .map(|row| {
                    evaluate_calculated_row(
                        pivot_name,
                        field,
                        ast,
                        columns,
                        row,
                        lookup,
                        source_name,
                        workbook_context,
                    )
                })
                .collect();
        }
    }

    (0..row_count)
        .map(|row| {
            evaluate_calculated_row(
                pivot_name,
                field,
                ast,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )
        })
        .collect()
}

pub(crate) fn evaluate_calculated_row(
    pivot_name: &str,
    field: &PivotCalculatedField,
    ast: &FormulaExpr,
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
    source_name: Option<&str>,
    workbook_context: Option<CalculatedWorkbookContext<'_>>,
) -> Result<PivotValue> {
    let materialized = materialize_calculated_expr(
        pivot_name,
        field,
        ast,
        columns,
        row,
        lookup,
        source_name,
        workbook_context,
    )?;
    let value = if let Some(context) = workbook_context {
        let evaluation_context =
            EvaluationContext::new(Some(context.workbook), context.sheet_index, 0, 0);
        evaluate(&materialized, &evaluation_context)
    } else {
        evaluate(&materialized, &EvaluationContext::simple())
    }
    .map_err(|error| {
        Error::other(format!(
            "pivot table {pivot_name} calculated field {} evaluation failed: {error}",
            field.name
        ))
    })?;
    Ok(formula_value_to_pivot_value(value))
}

pub(crate) fn materialize_calculated_expr(
    pivot_name: &str,
    field: &PivotCalculatedField,
    expr: &FormulaExpr,
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
    source_name: Option<&str>,
    workbook_context: Option<CalculatedWorkbookContext<'_>>,
) -> Result<FormulaExpr> {
    Ok(match expr {
        FormulaExpr::Number(value) => FormulaExpr::Number(*value),
        FormulaExpr::String(value) => FormulaExpr::String(value.clone()),
        FormulaExpr::Boolean(value) => FormulaExpr::Boolean(*value),
        FormulaExpr::Error(value) => FormulaExpr::Error(*value),
        FormulaExpr::Empty => FormulaExpr::Empty,
        FormulaExpr::NameRef(name) => {
            calculated_field_value_expr(pivot_name, field, name, columns, row, lookup)?
        }
        FormulaExpr::StructuredRef(reference) => {
            if let Some(name) = structured_ref_field_name(reference, source_name) {
                calculated_field_value_expr(pivot_name, field, name, columns, row, lookup)?
            } else {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} calculated field {} uses an unsupported structured reference",
                    field.name
                )));
            }
        }
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) if workbook_context.is_some() => {
            expr.clone()
        }
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated field {} uses workbook references, which are not valid pivot source-field references",
                field.name
            )));
        }
        FormulaExpr::ExternalRef(_) => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated field {} uses external workbook references, which are not supported",
                field.name
            )));
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op: *op,
            left: Box::new(materialize_calculated_expr(
                pivot_name,
                field,
                left,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )?),
            right: Box::new(materialize_calculated_expr(
                pivot_name,
                field,
                right,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )?),
        },
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp {
            op: *op,
            operand: Box::new(materialize_calculated_expr(
                pivot_name,
                field,
                operand,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )?),
        },
        FormulaExpr::Function { name, args } => FormulaExpr::Function {
            name: name.clone(),
            args: materialize_calculated_args(
                pivot_name,
                field,
                args,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )?,
        },
        FormulaExpr::ExternalFunction { book, name, args } => FormulaExpr::ExternalFunction {
            book: book.clone(),
            name: name.clone(),
            args: materialize_calculated_args(
                pivot_name,
                field,
                args,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )?,
        },
        FormulaExpr::Array(rows) => {
            let mut materialized_rows = Vec::with_capacity(rows.len());
            for formula_row in rows {
                materialized_rows.push(materialize_calculated_args(
                    pivot_name,
                    field,
                    formula_row,
                    columns,
                    row,
                    lookup,
                    source_name,
                    workbook_context,
                )?);
            }
            FormulaExpr::Array(materialized_rows)
        }
    })
}

pub(crate) fn materialize_calculated_args(
    pivot_name: &str,
    field: &PivotCalculatedField,
    args: &[FormulaExpr],
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
    source_name: Option<&str>,
    workbook_context: Option<CalculatedWorkbookContext<'_>>,
) -> Result<Vec<FormulaExpr>> {
    args.iter()
        .map(|arg| {
            materialize_calculated_expr(
                pivot_name,
                field,
                arg,
                columns,
                row,
                lookup,
                source_name,
                workbook_context,
            )
        })
        .collect()
}

pub(crate) fn calculated_field_value_expr(
    pivot_name: &str,
    field: &PivotCalculatedField,
    name: &str,
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
) -> Result<FormulaExpr> {
    let index = lookup.get(&name.to_lowercase()).copied().ok_or_else(|| {
        Error::other(format!(
            "pivot table {pivot_name} calculated field {} references unknown field: {name}",
            field.name
        ))
    })?;
    Ok(pivot_value_to_formula_expr(columns[index].value(row)))
}

pub(crate) fn structured_ref_field_name<'a>(
    reference: &'a StructuredReference,
    source_name: Option<&str>,
) -> Option<&'a str> {
    if let Some(table) = reference.table.as_deref() {
        let source_name = source_name?;
        if !table.eq_ignore_ascii_case(source_name) {
            return None;
        }
    }
    if !reference
        .specifiers
        .iter()
        .all(|specifier| matches!(specifier, StructuredRefSpecifier::ThisRow))
    {
        return None;
    }
    reference.column.as_deref()
}

pub(crate) fn pivot_value_to_formula_expr(value: &PivotValue) -> FormulaExpr {
    match value {
        PivotValue::Blank => FormulaExpr::Empty,
        PivotValue::Boolean(value) => FormulaExpr::Boolean(*value),
        PivotValue::Number(value) => FormulaExpr::Number(*value),
        PivotValue::String(value) => FormulaExpr::String(value.clone()),
        PivotValue::Error(value) => FormulaExpr::Error(*value),
    }
}

pub(crate) fn formula_value_to_pivot_value(value: FormulaValue) -> PivotValue {
    match value {
        FormulaValue::Empty => PivotValue::Blank,
        FormulaValue::Boolean(value) => PivotValue::Boolean(value),
        FormulaValue::Number(value) => PivotValue::Number(value),
        FormulaValue::String(value) => PivotValue::String(value),
        FormulaValue::Error(value) => PivotValue::Error(value),
        FormulaValue::Array { .. } => PivotValue::Error(CellError::Value),
    }
}

pub(crate) fn grouped_column(
    snapshot: &SourceSnapshot,
    field_index: usize,
    grouping: &PivotGrouping,
    date_1904: bool,
    pivot_name: &str,
) -> Result<EncodedColumn> {
    match grouping {
        PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        } => grouped_number_column(snapshot, field_index, *start, *end, *interval, pivot_name),
        PivotGrouping::Date { units, .. } => {
            Ok(grouped_date_column(snapshot, field_index, units, date_1904))
        }
        PivotGrouping::Manual { groups, .. } => {
            grouped_manual_column(snapshot, field_index, groups, pivot_name)
        }
    }
}

pub(crate) fn grouped_number_column(
    snapshot: &SourceSnapshot,
    field_index: usize,
    start: Option<f64>,
    end: Option<f64>,
    interval: f64,
    pivot_name: &str,
) -> Result<EncodedColumn> {
    if !interval.is_finite() || interval <= 0.0 {
        return Err(Error::other(format!(
            "pivot table {pivot_name} uses an invalid numeric grouping interval"
        )));
    }
    let effective_start =
        start.unwrap_or_else(|| numeric_column_min(snapshot, field_index).unwrap_or(0.0));
    if !effective_start.is_finite() || end.is_some_and(|value| !value.is_finite()) {
        return Err(Error::other(format!(
            "pivot table {pivot_name} uses invalid numeric grouping bounds"
        )));
    }

    Ok(remap_grouped_column(snapshot, field_index, |value| {
        group_number_value(value, effective_start, end, interval)
    }))
}

pub(crate) fn numeric_column_min(snapshot: &SourceSnapshot, field_index: usize) -> Option<f64> {
    (0..snapshot.row_count)
        .filter_map(|row| match snapshot.value(row, field_index) {
            PivotValue::Number(value) if value.is_finite() => Some(*value),
            _ => None,
        })
        .min_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal))
}

pub(crate) fn group_number_value(
    value: &PivotValue,
    start: f64,
    end: Option<f64>,
    interval: f64,
) -> PivotValue {
    let PivotValue::Number(number) = value else {
        return value.clone();
    };
    if !number.is_finite() {
        return value.clone();
    }
    if *number < start {
        return PivotValue::String(format!("<{}", format_group_number(start)));
    }
    if let Some(end) = end {
        if *number > end {
            return PivotValue::String(format!(">{}", format_group_number(end)));
        }
    }

    let bin = start + ((*number - start) / interval).floor() * interval;
    PivotValue::Number(normalize_group_number(bin))
}

pub(crate) fn normalize_group_number(value: f64) -> f64 {
    let rounded = value.round();
    if (value - rounded).abs() < 1e-10 {
        rounded
    } else {
        value
    }
}

pub(crate) fn format_group_number(value: f64) -> String {
    let value = normalize_group_number(value);
    if value.fract() == 0.0 {
        format!("{}", value as i64)
    } else {
        value.to_string()
    }
}

pub(crate) fn grouped_date_column(
    snapshot: &SourceSnapshot,
    field_index: usize,
    units: &[duke_sheets_core::PivotDateGroupUnit],
    date_1904: bool,
) -> EncodedColumn {
    let date_system = if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    };

    remap_grouped_column(snapshot, field_index, |value| {
        group_date_value(value, units, date_system)
    })
}

pub(crate) fn group_date_value(
    value: &PivotValue,
    units: &[duke_sheets_core::PivotDateGroupUnit],
    date_system: DateSystem,
) -> PivotValue {
    use duke_sheets_core::PivotDateGroupUnit;

    let PivotValue::Number(serial) = value else {
        return value.clone();
    };
    if !serial.is_finite() || units.is_empty() {
        return value.clone();
    }
    let Some((year, month, day)) = serial_to_date(*serial, date_system) else {
        return value.clone();
    };
    let (hour, minute, second) = serial_to_time(*serial);

    if units.len() == 1 {
        return match units[0] {
            PivotDateGroupUnit::Years => PivotValue::Number(year as f64),
            PivotDateGroupUnit::Quarters => PivotValue::Number(((month - 1) / 3 + 1) as f64),
            PivotDateGroupUnit::Months => PivotValue::Number(month as f64),
            PivotDateGroupUnit::Days => PivotValue::Number(day as f64),
            PivotDateGroupUnit::Hours => PivotValue::Number(hour as f64),
            PivotDateGroupUnit::Minutes => PivotValue::Number(minute as f64),
            PivotDateGroupUnit::Seconds => PivotValue::Number(second as f64),
        };
    }

    let parts = units
        .iter()
        .map(|unit| match unit {
            PivotDateGroupUnit::Years => format!("{year:04}"),
            PivotDateGroupUnit::Quarters => format!("Q{}", (month - 1) / 3 + 1),
            PivotDateGroupUnit::Months => format!("{month:02}"),
            PivotDateGroupUnit::Days => format!("{day:02}"),
            PivotDateGroupUnit::Hours => format!("{hour:02}"),
            PivotDateGroupUnit::Minutes => format!("{minute:02}"),
            PivotDateGroupUnit::Seconds => format!("{second:02}"),
        })
        .collect::<Vec<_>>();
    PivotValue::String(parts.join("-"))
}

pub(crate) fn grouped_manual_column(
    snapshot: &SourceSnapshot,
    field_index: usize,
    groups: &[PivotManualGroup],
    pivot_name: &str,
) -> Result<EncodedColumn> {
    let lookup = manual_group_lookup(groups, pivot_name)?;

    Ok(remap_grouped_column(snapshot, field_index, |value| {
        group_manual_value(value, &lookup)
    }))
}

pub(crate) fn remap_grouped_column<F>(
    snapshot: &SourceSnapshot,
    field_index: usize,
    group_value: F,
) -> EncodedColumn
where
    F: Fn(&PivotValue) -> PivotValue,
{
    let source_column = &snapshot.columns[field_index];
    source_column.remap_dictionary(group_value)
}

pub(crate) fn manual_group_lookup(
    groups: &[PivotManualGroup],
    pivot_name: &str,
) -> Result<AHashMap<PivotValue, String>> {
    if groups.is_empty() {
        return Err(Error::other(format!(
            "pivot table {pivot_name} uses an empty manual grouping"
        )));
    }

    let mut names = AHashSet::new();
    let mut lookup = AHashMap::new();
    for group in groups {
        if group.name.trim().is_empty() {
            return Err(Error::other(format!(
                "pivot table {pivot_name} has a manual group with a blank name"
            )));
        }
        if group.members.is_empty() {
            return Err(Error::other(format!(
                "pivot table {pivot_name} manual group {} has no members",
                group.name
            )));
        }
        if !names.insert(group.name.to_lowercase()) {
            return Err(Error::other(format!(
                "pivot table {pivot_name} has duplicate manual group name {}",
                group.name
            )));
        }
        for member in &group.members {
            if lookup.insert(member.clone(), group.name.clone()).is_some() {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} assigns pivot item {member} to more than one manual group"
                )));
            }
        }
    }

    Ok(lookup)
}

pub(crate) fn group_manual_value(
    value: &PivotValue,
    lookup: &AHashMap<PivotValue, String>,
) -> PivotValue {
    lookup
        .get(value)
        .map(|group| PivotValue::String(group.clone()))
        .unwrap_or_else(|| value.clone())
}
