use crate::api::*;
use crate::filters::*;
use crate::prelude::*;
use crate::runtime_cache::*;
use crate::snapshot::*;
use crate::sort::*;
use crate::source::*;
use crate::transform::*;

#[derive(Debug, Clone)]
pub(crate) struct CompiledPivotPlan {
    pub(crate) row_indexes: Vec<usize>,
    pub(crate) column_indexes: Vec<usize>,
    pub(crate) page_indexes: Vec<usize>,
    pub(crate) row_fields: Vec<PivotField>,
    pub(crate) column_fields: Vec<PivotField>,
    pub(crate) page_fields: Vec<PivotField>,
    pub(crate) row_collapsed_item_ids: Vec<AHashSet<u32>>,
    pub(crate) column_collapsed_item_ids: Vec<AHashSet<u32>>,
    pub(crate) measure_indexes: Vec<usize>,
    pub(crate) measures: Vec<PivotMeasure>,
    pub(crate) row_sort_measure_indexes: Vec<Option<usize>>,
    pub(crate) column_sort_measure_indexes: Vec<Option<usize>>,
    pub(crate) filters: Vec<CompiledFilter>,
    pub(crate) totals_filters: Vec<CompiledFilter>,
    pub(crate) axis_filters: Vec<CompiledFilter>,
    pub(crate) aggregate_filters: Vec<CompiledAggregateFilter>,
    pub(crate) calculated_items: Vec<CompiledCalculatedItem>,
    pub(crate) row_parent_total_positions: AHashSet<usize>,
    pub(crate) column_parent_total_positions: AHashSet<usize>,
    pub(crate) error_caption: Option<String>,
    pub(crate) missing_caption: Option<String>,
    pub(crate) asterisk_totals: bool,
}

impl CompiledPivotPlan {
    pub(crate) fn compile(
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        filter_baselines: &PivotFilterBaselines,
        options: &PivotRefreshOptions,
        date_system: DateSystem,
    ) -> Result<Self> {
        if pivot.measures.is_empty() {
            return Err(Error::other(format!(
                "pivot table {} must contain at least one measure",
                pivot.name
            )));
        }
        let (row_indexes, row_fields) =
            compile_axis_fields("row", &pivot.name, &pivot.rows, snapshot, &pivot.groupings)?;
        let (column_indexes, column_fields) = compile_axis_fields(
            "column",
            &pivot.name,
            &pivot.columns,
            snapshot,
            &pivot.groupings,
        )?;
        let (page_indexes, page_fields) = compile_axis_fields(
            "page",
            &pivot.name,
            &pivot.page_fields,
            snapshot,
            &pivot.groupings,
        )?;
        let row_collapsed_item_ids =
            compile_collapsed_item_ids("row", &pivot.name, snapshot, &row_indexes, &row_fields)?;
        let column_collapsed_item_ids = compile_collapsed_item_ids(
            "column",
            &pivot.name,
            snapshot,
            &column_indexes,
            &column_fields,
        )?;

        let mut measure_indexes = Vec::with_capacity(pivot.measures.len());
        for measure in &pivot.measures {
            validate_show_as(
                &pivot.name,
                snapshot,
                &row_indexes,
                &column_indexes,
                &measure.show_as,
            )?;
            measure_indexes.push(snapshot.required_field_index(&measure.field.name, &pivot.name)?);
        }
        let row_sort_measure_indexes =
            compile_axis_sort_measure_indexes(&pivot.name, &row_fields, &pivot.measures)?;
        let column_sort_measure_indexes =
            compile_axis_sort_measure_indexes(&pivot.name, &column_fields, &pivot.measures)?;
        let (row_parent_total_positions, column_parent_total_positions) =
            parent_total_subtotal_positions(
                &pivot.name,
                snapshot,
                &row_indexes,
                &column_indexes,
                &pivot.measures,
            )?;

        let mut filters = Vec::new();
        let mut aggregate_filters = Vec::new();
        for filter in &pivot.filters {
            match filter {
                PivotFilter::FieldItems { .. }
                | PivotFilter::Label { .. }
                | PivotFilter::LabelBetween { .. }
                | PivotFilter::Date { .. }
                | PivotFilter::DateBetween { .. }
                | PivotFilter::DatePeriod { .. } => {
                    filters.push(CompiledFilter::compile(
                        filter,
                        snapshot,
                        &pivot.name,
                        filter_baselines,
                        options,
                        date_system,
                    )?);
                }
                PivotFilter::Value { .. }
                | PivotFilter::ValueBetween { .. }
                | PivotFilter::TopN { .. } => {
                    aggregate_filters.push(CompiledAggregateFilter::compile(
                        filter,
                        snapshot,
                        &pivot.name,
                        &row_indexes,
                        &column_indexes,
                        &pivot.measures,
                    )?);
                }
                PivotFilter::Unsupported { kind, .. } => {
                    return Err(Error::other(format!(
                        "pivot table {} contains unsupported filter: {kind}",
                        pivot.name
                    )));
                }
            }
        }
        let totals_filters = filters
            .iter()
            .filter(|filter| !filter.targets_axis(&row_indexes, &column_indexes))
            .cloned()
            .collect();
        let axis_filters = filters
            .iter()
            .filter(|filter| filter.targets_axis(&row_indexes, &column_indexes))
            .cloned()
            .collect();
        let calculated_items = compile_calculated_items(
            &pivot.name,
            &pivot.calculated_items,
            snapshot,
            &row_indexes,
            &column_indexes,
        )?;

        Ok(Self {
            row_indexes,
            column_indexes,
            page_indexes,
            row_fields,
            column_fields,
            page_fields,
            row_collapsed_item_ids,
            column_collapsed_item_ids,
            measure_indexes,
            measures: pivot.measures.clone(),
            row_sort_measure_indexes,
            column_sort_measure_indexes,
            filters,
            totals_filters,
            axis_filters,
            aggregate_filters,
            calculated_items,
            row_parent_total_positions,
            column_parent_total_positions,
            error_caption: pivot
                .layout
                .show_error
                .then(|| pivot.layout.error_caption.clone())
                .flatten(),
            missing_caption: pivot
                .layout
                .show_missing
                .then(|| pivot.layout.missing_caption.clone())
                .flatten(),
            asterisk_totals: pivot.layout.asterisk_totals,
        })
    }
}

impl SourceSnapshot {
    pub(crate) fn required_field_index(&self, field_name: &str, pivot_name: &str) -> Result<usize> {
        self.field_index(field_name).ok_or_else(|| {
            Error::other(format!(
                "pivot table {pivot_name} references unknown source field: {field_name}"
            ))
        })
    }

    pub(crate) fn formula_reference_item_id(
        &self,
        field_index: usize,
        reference: &str,
    ) -> Option<u32> {
        let reference = reference.trim();
        if reference.is_empty() {
            return None;
        }

        let candidate = PivotValue::String(reference.to_string());
        if let Some(id) = self.columns[field_index].id_for_value(&candidate) {
            return Some(id);
        }

        self.columns[field_index]
            .dictionary
            .iter()
            .enumerate()
            .find_map(|(index, value)| match value {
                PivotValue::String(text) if text.eq_ignore_ascii_case(reference) => {
                    Some(index as u32)
                }
                _ if value.to_string().eq_ignore_ascii_case(reference) => Some(index as u32),
                _ => None,
            })
    }

    pub(crate) fn grouped_date_field_index(
        &self,
        field_name: &str,
        unit: duke_sheets_core::PivotDateGroupUnit,
    ) -> Option<(usize, String)> {
        let base = grouped_date_header(field_name, unit);
        self.headers
            .iter()
            .enumerate()
            .rev()
            .find(|(_, header)| grouped_header_matches(header, &base))
            .map(|(index, header)| (index, header.clone()))
    }
}

pub(crate) fn compile_collapsed_item_ids(
    axis_name: &str,
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    indexes: &[usize],
    fields: &[PivotField],
) -> Result<Vec<AHashSet<u32>>> {
    indexes
        .iter()
        .zip(fields.iter())
        .map(|(field_index, field)| {
            field
                .collapsed_items
                .iter()
                .map(|item| {
                    snapshot.columns[*field_index].id_for_value(item).ok_or_else(|| {
                        Error::other(format!(
                            "pivot table {pivot_name} {axis_name} field {} collapsed item was not found in the source data: {item}",
                            field.field.name
                        ))
                    })
                })
                .collect()
        })
        .collect()
}

pub(crate) fn compile_axis_fields(
    axis_name: &str,
    pivot_name: &str,
    fields: &[PivotField],
    snapshot: &SourceSnapshot,
    groupings: &[PivotGrouping],
) -> Result<(Vec<usize>, Vec<PivotField>)> {
    let mut indexes = Vec::new();
    let mut compiled_fields = Vec::new();
    for field in fields {
        if let Some(units) = multi_unit_date_grouping_units(groupings, &field.field.name) {
            for unit in units {
                let (index, header) = snapshot
                    .grouped_date_field_index(&field.field.name, *unit)
                    .ok_or_else(|| {
                        Error::other(format!(
                            "pivot table {pivot_name} references unknown grouped {axis_name} field: {}",
                            grouped_date_header(&field.field.name, *unit)
                        ))
                    })?;
                let mut grouped_field = field.clone();
                grouped_field.field.name = header;
                indexes.push(index);
                compiled_fields.push(grouped_field);
            }
        } else {
            let index = snapshot
                .required_field_index(&field.field.name, pivot_name)
                .map_err(|_| {
                    Error::other(format!(
                        "pivot table {pivot_name} references unknown {axis_name} field: {}",
                        field.field.name
                    ))
                })?;
            indexes.push(index);
            compiled_fields.push(field.clone());
        }
    }
    Ok((indexes, compiled_fields))
}

#[derive(Debug, Clone)]
pub(crate) struct CompiledCalculatedItem {
    pub(crate) axis: AggregateFilterAxis,
    pub(crate) position: usize,
    pub(crate) field_index: usize,
    pub(crate) item_id: u32,
    pub(crate) item: PivotValue,
    pub(crate) ast: FormulaExpr,
}

pub(crate) fn compile_calculated_items(
    pivot_name: &str,
    calculated_items: &[PivotCalculatedItem],
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    column_indexes: &[usize],
) -> Result<Vec<CompiledCalculatedItem>> {
    let mut compiled = Vec::with_capacity(calculated_items.len());
    let mut targets = AHashSet::new();

    for item in calculated_items {
        let field_index = snapshot.required_field_index(&item.field.name, pivot_name)?;
        let (axis, position) = calculated_item_axis(
            pivot_name,
            &item.field.name,
            field_index,
            row_indexes,
            column_indexes,
        )?;
        let item_id = snapshot.columns[field_index]
            .id_for_value(&item.item)
            .ok_or_else(|| {
                Error::other(format!(
                    "pivot table {pivot_name} calculated item {} was not registered in field {}",
                    item.item, item.field.name
                ))
            })?;
        let target_key = (axis_key(axis), position, item_id);
        if !targets.insert(target_key) {
            return Err(Error::other(format!(
                "pivot table {pivot_name} defines calculated item {} more than once in field {}",
                item.item, item.field.name
            )));
        }

        let ast = parse_calculated_item_formula(pivot_name, item)?;
        if calculated_item_formula_references_item(&ast, snapshot, field_index, item_id) {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated item {} references itself",
                item.item
            )));
        }

        compiled.push(CompiledCalculatedItem {
            axis,
            position,
            field_index,
            item_id,
            item: item.item.clone(),
            ast,
        });
    }

    order_calculated_items(pivot_name, compiled, snapshot)
}

pub(crate) fn order_calculated_items(
    pivot_name: &str,
    items: Vec<CompiledCalculatedItem>,
    snapshot: &SourceSnapshot,
) -> Result<Vec<CompiledCalculatedItem>> {
    if items.len() < 2 {
        return Ok(items);
    }

    let target_indexes = items
        .iter()
        .enumerate()
        .map(|(index, item)| ((axis_key(item.axis), item.position, item.item_id), index))
        .collect::<AHashMap<_, _>>();
    let dependencies = items
        .iter()
        .map(|item| calculated_item_dependency_indexes(item, snapshot, &target_indexes))
        .collect::<Vec<_>>();
    let mut visit_state = vec![CalculatedItemVisitState::Unvisited; items.len()];
    let mut ordered = Vec::with_capacity(items.len());

    for index in 0..items.len() {
        visit_calculated_item(
            pivot_name,
            index,
            &dependencies,
            &mut visit_state,
            &mut ordered,
        )?;
    }

    Ok(ordered
        .into_iter()
        .map(|index| items[index].clone())
        .collect())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CalculatedItemVisitState {
    Unvisited,
    Visiting,
    Visited,
}

pub(crate) fn calculated_item_dependency_indexes(
    item: &CompiledCalculatedItem,
    snapshot: &SourceSnapshot,
    target_indexes: &AHashMap<(u8, usize, u32), usize>,
) -> Vec<usize> {
    let mut item_ids = AHashSet::new();
    collect_calculated_item_formula_references(
        &item.ast,
        snapshot,
        item.field_index,
        &mut item_ids,
    );

    let mut indexes = item_ids
        .into_iter()
        .filter_map(|item_id| {
            target_indexes
                .get(&(axis_key(item.axis), item.position, item_id))
                .copied()
        })
        .collect::<Vec<_>>();
    indexes.sort_unstable();
    indexes.dedup();
    indexes
}

pub(crate) fn collect_calculated_item_formula_references(
    expr: &FormulaExpr,
    snapshot: &SourceSnapshot,
    field_index: usize,
    references: &mut AHashSet<u32>,
) {
    match expr {
        FormulaExpr::NameRef(name) | FormulaExpr::String(name) => {
            if let Some(item_id) = snapshot.formula_reference_item_id(field_index, name) {
                references.insert(item_id);
            }
        }
        FormulaExpr::CellRef(reference) => {
            if let Some(name) = calculated_item_cell_reference_name(reference) {
                if let Some(item_id) = snapshot.formula_reference_item_id(field_index, &name) {
                    references.insert(item_id);
                }
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_calculated_item_formula_references(left, snapshot, field_index, references);
            collect_calculated_item_formula_references(right, snapshot, field_index, references);
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            collect_calculated_item_formula_references(operand, snapshot, field_index, references);
        }
        FormulaExpr::Function { args, .. } | FormulaExpr::ExternalFunction { args, .. } => {
            for arg in args {
                collect_calculated_item_formula_references(arg, snapshot, field_index, references);
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for arg in row {
                    collect_calculated_item_formula_references(
                        arg,
                        snapshot,
                        field_index,
                        references,
                    );
                }
            }
        }
        FormulaExpr::Number(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::Empty
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalRef(_) => {}
    }
}

pub(crate) fn visit_calculated_item(
    pivot_name: &str,
    index: usize,
    dependencies: &[Vec<usize>],
    visit_state: &mut [CalculatedItemVisitState],
    ordered: &mut Vec<usize>,
) -> Result<()> {
    match visit_state[index] {
        CalculatedItemVisitState::Visited => return Ok(()),
        CalculatedItemVisitState::Visiting => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated items contain a circular reference"
            )));
        }
        CalculatedItemVisitState::Unvisited => {}
    }

    visit_state[index] = CalculatedItemVisitState::Visiting;
    for dependency in &dependencies[index] {
        visit_calculated_item(pivot_name, *dependency, dependencies, visit_state, ordered)?;
    }
    visit_state[index] = CalculatedItemVisitState::Visited;
    ordered.push(index);
    Ok(())
}

pub(crate) fn calculated_item_axis(
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
                "pivot table {pivot_name} calculated item field {field_name} is not on a row or column axis"
            ))
        })
}

pub(crate) fn axis_key(axis: AggregateFilterAxis) -> u8 {
    match axis {
        AggregateFilterAxis::Row => 0,
        AggregateFilterAxis::Column => 1,
    }
}

pub(crate) fn calculated_item_formula_references_item(
    expr: &FormulaExpr,
    snapshot: &SourceSnapshot,
    field_index: usize,
    item_id: u32,
) -> bool {
    match expr {
        FormulaExpr::NameRef(name) | FormulaExpr::String(name) => {
            snapshot.formula_reference_item_id(field_index, name) == Some(item_id)
        }
        FormulaExpr::CellRef(reference) => {
            calculated_item_cell_reference_name(reference)
                .as_deref()
                .and_then(|name| snapshot.formula_reference_item_id(field_index, name))
                == Some(item_id)
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            calculated_item_formula_references_item(left, snapshot, field_index, item_id)
                || calculated_item_formula_references_item(right, snapshot, field_index, item_id)
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            calculated_item_formula_references_item(operand, snapshot, field_index, item_id)
        }
        FormulaExpr::Function { args, .. } | FormulaExpr::ExternalFunction { args, .. } => {
            args.iter().any(|arg| {
                calculated_item_formula_references_item(arg, snapshot, field_index, item_id)
            })
        }
        FormulaExpr::Array(rows) => rows.iter().any(|row| {
            row.iter().any(|arg| {
                calculated_item_formula_references_item(arg, snapshot, field_index, item_id)
            })
        }),
        FormulaExpr::Number(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::Empty
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalRef(_) => false,
    }
}

pub(crate) fn calculated_item_cell_reference_name(reference: &CellReference) -> Option<String> {
    reference
        .sheet
        .is_none()
        .then(|| reference.address.to_a1_string())
}

pub(crate) fn grouped_header_matches(header: &str, base: &str) -> bool {
    if header.eq_ignore_ascii_case(base) {
        return true;
    }
    header
        .strip_prefix(base)
        .and_then(|suffix| suffix.strip_prefix(' '))
        .is_some_and(|suffix| suffix.parse::<usize>().is_ok())
}

pub(crate) fn multi_unit_date_grouping_units<'a>(
    groupings: &'a [PivotGrouping],
    field_name: &str,
) -> Option<&'a [duke_sheets_core::PivotDateGroupUnit]> {
    groupings.iter().find_map(|grouping| match grouping {
        PivotGrouping::Date { field, units }
            if field.name.eq_ignore_ascii_case(field_name) && units.len() > 1 =>
        {
            Some(units.as_slice())
        }
        _ => None,
    })
}

pub(crate) fn validate_show_as(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    column_indexes: &[usize],
    show_as: &PivotShowAs,
) -> Result<()> {
    match show_as {
        PivotShowAs::Normal
        | PivotShowAs::PercentOfGrandTotal
        | PivotShowAs::PercentOfRowTotal
        | PivotShowAs::PercentOfColumnTotal
        | PivotShowAs::PercentOfParentRowTotal
        | PivotShowAs::PercentOfParentColumnTotal
        | PivotShowAs::Index => Ok(()),
        PivotShowAs::RunningTotal { base_field }
        | PivotShowAs::PercentOfParentTotal { base_field }
        | PivotShowAs::RankAscending { base_field }
        | PivotShowAs::RankDescending { base_field } => validate_base_field(
            pivot_name,
            snapshot,
            row_indexes,
            column_indexes,
            &base_field.name,
        )
        .map(|_| ()),
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        }
        | PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => {
            let field_index = validate_base_field(
                pivot_name,
                snapshot,
                row_indexes,
                column_indexes,
                &base_field.name,
            )?;
            if snapshot.columns[field_index]
                .id_for_value(base_item)
                .is_none()
            {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} references missing show-as base item {} in field {}",
                    base_item, base_field.name
                )));
            }
            Ok(())
        }
    }
}

pub(crate) fn parent_total_subtotal_positions(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    column_indexes: &[usize],
    measures: &[PivotMeasure],
) -> Result<(AHashSet<usize>, AHashSet<usize>)> {
    let mut row_positions = AHashSet::new();
    let mut column_positions = AHashSet::new();

    for measure in measures {
        match &measure.show_as {
            PivotShowAs::PercentOfParentRowTotal => {
                row_positions.extend(0..row_indexes.len().saturating_sub(1));
            }
            PivotShowAs::PercentOfParentColumnTotal => {
                column_positions.extend(0..column_indexes.len().saturating_sub(1));
            }
            PivotShowAs::PercentOfParentTotal { base_field } => {
                let field_index = snapshot.required_field_index(&base_field.name, pivot_name)?;
                if let Some(position) = row_indexes.iter().position(|index| *index == field_index) {
                    row_positions.insert(position);
                } else if let Some(position) = column_indexes
                    .iter()
                    .position(|index| *index == field_index)
                {
                    column_positions.insert(position);
                }
            }
            PivotShowAs::Normal
            | PivotShowAs::PercentOfGrandTotal
            | PivotShowAs::PercentOfRowTotal
            | PivotShowAs::PercentOfColumnTotal
            | PivotShowAs::Index
            | PivotShowAs::RunningTotal { .. }
            | PivotShowAs::DifferenceFrom { .. }
            | PivotShowAs::PercentDifferenceFrom { .. }
            | PivotShowAs::RankAscending { .. }
            | PivotShowAs::RankDescending { .. } => {}
        }
    }

    Ok((row_positions, column_positions))
}

pub(crate) fn validate_base_field(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    column_indexes: &[usize],
    base_field: &str,
) -> Result<usize> {
    let field_index = snapshot.required_field_index(base_field, pivot_name)?;
    if row_indexes.contains(&field_index) || column_indexes.contains(&field_index) {
        Ok(field_index)
    } else {
        Err(Error::other(format!(
            "pivot table {pivot_name} uses show-as base field {base_field}, but that field is not on a row or column axis"
        )))
    }
}
