//! Pivot table refresh engine.
//!
//! This crate refreshes semantic [`duke_sheets_core::PivotTable`] definitions
//! into worksheet cells. The file-format pivot cache objects used by XLSX/XLSB
//! are deliberately kept out of the public authoring API.

use std::any::Any;
use std::cmp::Ordering;
#[cfg(feature = "parallel")]
use std::hash::Hash;
use std::sync::Arc;

use ahash::{AHashMap, AHashSet};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue, Error, PivotAggregate, PivotCalculatedField,
    PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotMeasure,
    PivotOverwritePolicy, PivotRefreshStatus, PivotShowAs, PivotSort, PivotSource, PivotTable,
    PivotValue, Result, Table, Workbook, Worksheet, MAX_COLS, MAX_ROWS,
};
use duke_sheets_formula::{
    evaluate, parse_formula, EvaluationContext, FormulaExpr, FormulaValue, StructuredRefSpecifier,
    StructuredReference,
};
#[cfg(feature = "parallel")]
use rayon::prelude::*;
use ssfmt::{
    date_serial::{serial_to_date, serial_to_time},
    DateSystem,
};

#[cfg(feature = "parallel")]
const PARALLEL_ROW_THRESHOLD: usize = 50_000;
#[cfg(feature = "parallel")]
const PARALLEL_CHUNK_SIZE: usize = 16_384;

/// Result statistics for a pivot refresh operation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PivotRefreshStats {
    /// Pivot tables present in the selected refresh scope.
    pub pivot_count: usize,
    /// Pivot tables successfully refreshed.
    pub pivots_refreshed: usize,
    /// Source data rows scanned.
    pub source_rows: usize,
    /// Output cells written or cleared.
    pub output_cells: usize,
    /// Source snapshot cache hits during this refresh.
    pub cache_hits: usize,
    /// Source snapshot cache misses during this refresh.
    pub cache_misses: usize,
}

impl PivotRefreshStats {
    fn add_rendered(&mut self, rendered: &RenderedPivot) {
        self.pivots_refreshed += 1;
        self.source_rows += rendered.source_rows;
        self.output_cells += rendered.cell_count();
    }
}

/// Extension methods for refreshing pivot tables in a workbook.
pub trait WorkbookPivotExt {
    /// Refresh all pivot tables in the workbook.
    fn refresh_pivots(&mut self) -> Result<PivotRefreshStats>;

    /// Refresh a single pivot table by worksheet index and pivot name.
    fn refresh_pivot(&mut self, sheet_index: usize, pivot_name: &str) -> Result<PivotRefreshStats>;
}

impl WorkbookPivotExt for Workbook {
    fn refresh_pivots(&mut self) -> Result<PivotRefreshStats> {
        let mut cache = take_runtime_cache(self);
        let result = refresh_pivots_inner(self, &mut cache);
        self.set_pivot_runtime_cache(Box::new(cache));
        result
    }

    fn refresh_pivot(&mut self, sheet_index: usize, pivot_name: &str) -> Result<PivotRefreshStats> {
        let mut cache = take_runtime_cache(self);
        let result = refresh_pivot_inner(self, sheet_index, pivot_name, &mut cache);
        self.set_pivot_runtime_cache(Box::new(cache));
        result
    }
}

#[derive(Debug, Clone)]
struct PivotJob {
    sheet_index: usize,
    pivot_index: usize,
    pivot: PivotTable,
}

fn refresh_pivots_inner(
    workbook: &mut Workbook,
    cache: &mut PivotRuntimeCache,
) -> Result<PivotRefreshStats> {
    let jobs = collect_pivot_jobs(workbook);
    let mut stats = PivotRefreshStats {
        pivot_count: jobs.len(),
        ..PivotRefreshStats::default()
    };

    let mut rendered = Vec::with_capacity(jobs.len());
    for job in jobs {
        match build_rendered_pivot(workbook, job.sheet_index, &job.pivot, cache, &mut stats) {
            Ok(output) => {
                stats.add_rendered(&output);
                rendered.push((job, output));
            }
            Err(error) => {
                mark_pivot_failed(
                    workbook,
                    job.sheet_index,
                    job.pivot_index,
                    error.to_string(),
                );
                return Err(error);
            }
        }
    }

    for (job, output) in rendered {
        if let Err(error) = write_rendered_pivot(workbook, &job, output) {
            mark_pivot_failed(
                workbook,
                job.sheet_index,
                job.pivot_index,
                error.to_string(),
            );
            return Err(error);
        }
    }

    Ok(stats)
}

fn refresh_pivot_inner(
    workbook: &mut Workbook,
    sheet_index: usize,
    pivot_name: &str,
    cache: &mut PivotRuntimeCache,
) -> Result<PivotRefreshStats> {
    let worksheet = workbook
        .worksheet(sheet_index)
        .ok_or_else(|| Error::SheetOutOfBounds(sheet_index, workbook.sheet_count()))?;
    let pivot_index = worksheet
        .pivot_tables()
        .iter()
        .position(|pivot| pivot.name.eq_ignore_ascii_case(pivot_name))
        .ok_or_else(|| Error::other(format!("pivot table not found: {pivot_name}")))?;
    let pivot = worksheet.pivot_tables()[pivot_index].clone();

    let mut stats = PivotRefreshStats {
        pivot_count: 1,
        ..PivotRefreshStats::default()
    };

    let output = match build_rendered_pivot(workbook, sheet_index, &pivot, cache, &mut stats) {
        Ok(output) => output,
        Err(error) => {
            mark_pivot_failed(workbook, sheet_index, pivot_index, error.to_string());
            return Err(error);
        }
    };
    stats.add_rendered(&output);
    if let Err(error) = write_rendered_pivot(
        workbook,
        &PivotJob {
            sheet_index,
            pivot_index,
            pivot,
        },
        output,
    ) {
        mark_pivot_failed(workbook, sheet_index, pivot_index, error.to_string());
        return Err(error);
    }

    Ok(stats)
}

fn collect_pivot_jobs(workbook: &Workbook) -> Vec<PivotJob> {
    workbook
        .worksheets()
        .enumerate()
        .flat_map(|(sheet_index, worksheet)| {
            worksheet
                .pivot_tables()
                .iter()
                .cloned()
                .enumerate()
                .map(move |(pivot_index, pivot)| PivotJob {
                    sheet_index,
                    pivot_index,
                    pivot,
                })
        })
        .collect()
}

fn build_rendered_pivot(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    pivot: &PivotTable,
    cache: &mut PivotRuntimeCache,
    stats: &mut PivotRefreshStats,
) -> Result<RenderedPivot> {
    let raw_snapshot =
        snapshot_for_source(workbook, pivot_sheet_index, &pivot.source, cache, stats)?;
    let calculated_snapshot = if pivot.calculated_fields.is_empty() {
        raw_snapshot
    } else {
        Arc::new(raw_snapshot.apply_calculated_fields(&pivot.name, &pivot.calculated_fields)?)
    };
    let snapshot = if pivot.groupings.is_empty() {
        calculated_snapshot
    } else {
        Arc::new(calculated_snapshot.apply_groupings(
            &pivot.name,
            &pivot.groupings,
            workbook.settings().date_1904,
        )?)
    };
    let plan = CompiledPivotPlan::compile(pivot, &snapshot)?;
    let mut aggregation = PivotAggregation::aggregate(&snapshot, &plan);
    aggregation.apply_aggregate_filters(&plan);
    aggregation.sort_orders(&snapshot, &plan);
    render_pivot(pivot, &snapshot, &plan, &aggregation)
}

fn write_rendered_pivot(
    workbook: &mut Workbook,
    job: &PivotJob,
    rendered: RenderedPivot,
) -> Result<()> {
    let sheet_count = workbook.sheet_count();
    let worksheet = workbook
        .worksheet_mut(job.sheet_index)
        .ok_or_else(|| Error::SheetOutOfBounds(job.sheet_index, sheet_count))?;

    if matches!(
        job.pivot.overwrite_policy,
        PivotOverwritePolicy::FailOnOccupied
    ) {
        ensure_output_range_is_available(worksheet, &job.pivot, rendered.range)?;
    }

    if matches!(
        job.pivot.overwrite_policy,
        PivotOverwritePolicy::ClearOwnedRange
    ) {
        if let Some(range) = job.pivot.rendered_range {
            worksheet.clear_range(&range);
        }
    }

    for (row_offset, row) in rendered.cells.iter().enumerate() {
        for (col_offset, value) in row.iter().enumerate() {
            let row = job.pivot.target.row + row_offset as u32;
            let col = job.pivot.target.col + col_offset as u16;
            if value.is_empty() {
                worksheet.clear_cell_at(row, col);
            } else {
                worksheet.set_cell_value_at(row, col, value.clone())?;
            }
        }
    }

    if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(job.pivot_index) {
        pivot.rendered_range = Some(rendered.range);
        pivot.refresh_status = PivotRefreshStatus::Succeeded;
        if let Some(cache_info) = &mut pivot.cache_info {
            cache_info.refresh_status = PivotRefreshStatus::Succeeded;
        }
    }

    Ok(())
}

fn mark_pivot_failed(
    workbook: &mut Workbook,
    sheet_index: usize,
    pivot_index: usize,
    message: String,
) {
    if let Some(worksheet) = workbook.worksheet_mut(sheet_index) {
        if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(pivot_index) {
            let status = PivotRefreshStatus::Failed { message };
            pivot.refresh_status = status.clone();
            if let Some(cache_info) = &mut pivot.cache_info {
                cache_info.refresh_status = status;
            }
        }
    }
}

fn ensure_output_range_is_available(
    worksheet: &Worksheet,
    pivot: &PivotTable,
    output_range: CellRange,
) -> Result<()> {
    for address in output_range.cells() {
        if pivot
            .rendered_range
            .is_some_and(|owned| owned.contains(&address))
        {
            continue;
        }

        if !worksheet.get_value_at(address.row, address.col).is_blank() {
            return Err(Error::other(format!(
                "pivot table {} would overwrite non-empty cell {}",
                pivot.name, address
            )));
        }
    }

    Ok(())
}

#[derive(Debug, Default)]
struct PivotRuntimeCache {
    workbook_nonce: u64,
    structural_generation: u64,
    snapshots: AHashMap<SourceCacheKey, Arc<SourceSnapshot>>,
}

impl PivotRuntimeCache {
    fn for_workbook(workbook: &Workbook) -> Self {
        Self {
            workbook_nonce: workbook.nonce(),
            structural_generation: workbook.structural_generation(),
            snapshots: AHashMap::new(),
        }
    }
}

fn take_runtime_cache(workbook: &mut Workbook) -> PivotRuntimeCache {
    let mut cache = workbook
        .take_pivot_runtime_cache()
        .and_then(|cache| {
            let cache: Box<dyn Any + Send + Sync> = cache;
            cache.downcast::<PivotRuntimeCache>().ok()
        })
        .map(|cache| *cache)
        .unwrap_or_default();

    if cache.workbook_nonce != workbook.nonce()
        || cache.structural_generation != workbook.structural_generation()
    {
        cache = PivotRuntimeCache::for_workbook(workbook);
    }

    cache
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct SourceCacheKey {
    kind: SourceCacheKind,
    sheet_index: usize,
    range: CellRange,
    source_name: Option<String>,
    mutation_count: u64,
    topology_generation: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum SourceCacheKind {
    WorksheetRange,
    Table,
}

#[derive(Debug, Clone)]
struct ResolvedSource {
    kind: SourceCacheKind,
    sheet_index: usize,
    range: CellRange,
    source_name: Option<String>,
    headers: Option<Vec<String>>,
    data_start_row: u32,
    data_end_row: Option<u32>,
    mutation_count: u64,
    topology_generation: u64,
}

impl ResolvedSource {
    fn cache_key(&self) -> SourceCacheKey {
        SourceCacheKey {
            kind: self.kind,
            sheet_index: self.sheet_index,
            range: self.range,
            source_name: self.source_name.clone(),
            mutation_count: self.mutation_count,
            topology_generation: self.topology_generation,
        }
    }
}

fn snapshot_for_source(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    source: &PivotSource,
    cache: &mut PivotRuntimeCache,
    stats: &mut PivotRefreshStats,
) -> Result<Arc<SourceSnapshot>> {
    let resolved = resolve_source(workbook, pivot_sheet_index, source)?;
    let cache_key = resolved.cache_key();

    if let Some(snapshot) = cache.snapshots.get(&cache_key) {
        stats.cache_hits += 1;
        return Ok(Arc::clone(snapshot));
    }

    let worksheet = workbook
        .worksheet(resolved.sheet_index)
        .ok_or_else(|| Error::SheetOutOfBounds(resolved.sheet_index, workbook.sheet_count()))?;
    let snapshot = Arc::new(SourceSnapshot::from_resolved(worksheet, &resolved)?);
    cache.snapshots.insert(cache_key, Arc::clone(&snapshot));
    stats.cache_misses += 1;
    Ok(snapshot)
}

fn resolve_source(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    source: &PivotSource,
) -> Result<ResolvedSource> {
    match source {
        PivotSource::WorksheetRange { sheet, range } => {
            let sheet_index = match sheet {
                Some(name) => workbook
                    .sheet_index(name)
                    .ok_or_else(|| Error::SheetNotFound(name.clone()))?,
                None => pivot_sheet_index,
            };
            let worksheet = workbook
                .worksheet(sheet_index)
                .ok_or_else(|| Error::SheetOutOfBounds(sheet_index, workbook.sheet_count()))?;
            if range.row_count() == 0 || range.col_count() == 0 {
                return Err(Error::other("pivot source range cannot be empty"));
            }

            Ok(ResolvedSource {
                kind: SourceCacheKind::WorksheetRange,
                sheet_index,
                range: *range,
                source_name: sheet.clone(),
                headers: None,
                data_start_row: range.start.row.saturating_add(1),
                data_end_row: if range.end.row > range.start.row {
                    Some(range.end.row)
                } else {
                    None
                },
                mutation_count: worksheet.mutation_count(),
                topology_generation: worksheet.topology_generation(),
            })
        }
        PivotSource::Table { name } => {
            let (sheet_index, worksheet, table) = find_table(workbook, name)
                .ok_or_else(|| Error::other(format!("table not found: {name}")))?;
            let headers = table_headers(table);
            let data_start_row = table
                .reference
                .start
                .row
                .saturating_add(table.header_row_count);
            let data_end_row = table_data_end_row(table);

            Ok(ResolvedSource {
                kind: SourceCacheKind::Table,
                sheet_index,
                range: table.reference,
                source_name: Some(table.name.clone()),
                headers: Some(headers),
                data_start_row,
                data_end_row,
                mutation_count: worksheet.mutation_count(),
                topology_generation: worksheet.topology_generation(),
            })
        }
        PivotSource::External { .. } => Err(Error::other(
            "external pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
        PivotSource::Consolidation { .. } => Err(Error::other(
            "consolidation pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
        PivotSource::Scenario { .. } => Err(Error::other(
            "scenario pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
        PivotSource::Olap { .. } => Err(Error::other(
            "OLAP pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
    }
}

fn find_table<'a>(workbook: &'a Workbook, name: &str) -> Option<(usize, &'a Worksheet, &'a Table)> {
    workbook
        .worksheets()
        .enumerate()
        .find_map(|(sheet_index, worksheet)| {
            worksheet
                .table_by_name(name)
                .map(|table| (sheet_index, worksheet, table))
        })
}

fn table_headers(table: &Table) -> Vec<String> {
    let col_count = table.reference.col_count() as usize;
    (0..col_count)
        .map(|index| {
            table
                .columns
                .get(index)
                .map(|column| column.name.clone())
                .unwrap_or_else(|| format!("Column{}", index + 1))
        })
        .collect()
}

fn table_data_end_row(table: &Table) -> Option<u32> {
    let totals_rows = table.totals_row_count;
    let end_row = table.reference.end.row.saturating_sub(totals_rows);
    if table
        .reference
        .start
        .row
        .saturating_add(table.header_row_count)
        > end_row
    {
        None
    } else {
        Some(end_row)
    }
}

#[derive(Debug, Clone)]
struct SourceSnapshot {
    headers: Vec<String>,
    columns: Vec<EncodedColumn>,
    row_count: usize,
}

impl SourceSnapshot {
    fn from_resolved(worksheet: &Worksheet, source: &ResolvedSource) -> Result<Self> {
        let col_count = source.range.col_count() as usize;
        let headers = match &source.headers {
            Some(headers) => normalize_supplied_headers(headers, col_count),
            None => read_headers_from_sheet(worksheet, source.range)?,
        };
        validate_headers(&headers)?;

        let row_count = source
            .data_end_row
            .map(|end_row| (end_row - source.data_start_row + 1) as usize)
            .unwrap_or(0);
        let mut columns = (0..col_count)
            .map(|_| EncodedColumn::with_capacity(row_count))
            .collect::<Vec<_>>();

        if let Some(data_end_row) = source.data_end_row {
            for source_col in source.range.start.col..=source.range.end.col {
                let col_index = (source_col - source.range.start.col) as usize;
                for row in source.data_start_row..=data_end_row {
                    columns[col_index].push(effective_pivot_value(worksheet, row, source_col));
                }
            }
        }

        Ok(Self {
            headers,
            columns,
            row_count,
        })
    }

    fn field_index(&self, name: &str) -> Option<usize> {
        self.headers
            .iter()
            .position(|header| header.eq_ignore_ascii_case(name))
    }

    fn value(&self, row: usize, col: usize) -> &PivotValue {
        self.columns[col].value(row)
    }

    fn value_by_id(&self, col: usize, id: u32) -> &PivotValue {
        self.columns[col].value_by_id(id)
    }

    fn apply_calculated_fields(
        &self,
        pivot_name: &str,
        calculated_fields: &[PivotCalculatedField],
    ) -> Result<Self> {
        let mut headers = self.headers.clone();
        let mut columns = self.columns.clone();

        for field in calculated_fields {
            if field.name.trim().is_empty() {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} has a calculated field with a blank name"
                )));
            }
            if headers
                .iter()
                .any(|header| header.eq_ignore_ascii_case(&field.name))
            {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} calculated field duplicates source field: {}",
                    field.name
                )));
            }

            let ast = parse_calculated_formula(pivot_name, field)?;
            let lookup = field_lookup(&headers);
            let values = evaluate_calculated_values(
                pivot_name,
                field,
                &ast,
                &columns,
                self.row_count,
                &lookup,
            )?;
            let mut column = EncodedColumn::with_capacity(self.row_count);
            for value in values {
                column.push(value);
            }
            headers.push(field.name.clone());
            columns.push(column);
        }

        Ok(Self {
            headers,
            columns,
            row_count: self.row_count,
        })
    }

    fn apply_groupings(
        &self,
        pivot_name: &str,
        groupings: &[PivotGrouping],
        date_1904: bool,
    ) -> Result<Self> {
        let mut columns = self.columns.clone();
        let mut grouped_fields = AHashSet::new();
        for grouping in groupings {
            let field_name = grouping_field_name(grouping);
            let field_index = self.field_index(field_name).ok_or_else(|| {
                Error::other(format!(
                    "pivot table {pivot_name} references missing grouping field: {field_name}"
                ))
            })?;
            if !grouped_fields.insert(field_index) {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} groups field {field_name} more than once"
                )));
            }
            columns[field_index] =
                grouped_column(self, field_index, grouping, date_1904, pivot_name)?;
        }

        Ok(Self {
            headers: self.headers.clone(),
            columns,
            row_count: self.row_count,
        })
    }
}

fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. } | PivotGrouping::Date { field, .. } => &field.name,
    }
}

fn parse_calculated_formula(pivot_name: &str, field: &PivotCalculatedField) -> Result<FormulaExpr> {
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

fn field_lookup(headers: &[String]) -> AHashMap<String, usize> {
    headers
        .iter()
        .enumerate()
        .map(|(index, header)| (header.to_lowercase(), index))
        .collect()
}

fn evaluate_calculated_values(
    pivot_name: &str,
    field: &PivotCalculatedField,
    ast: &FormulaExpr,
    columns: &[EncodedColumn],
    row_count: usize,
    lookup: &AHashMap<String, usize>,
) -> Result<Vec<PivotValue>> {
    #[cfg(feature = "parallel")]
    {
        if row_count >= PARALLEL_ROW_THRESHOLD {
            return (0..row_count)
                .into_par_iter()
                .map(|row| evaluate_calculated_row(pivot_name, field, ast, columns, row, lookup))
                .collect();
        }
    }

    (0..row_count)
        .map(|row| evaluate_calculated_row(pivot_name, field, ast, columns, row, lookup))
        .collect()
}

fn evaluate_calculated_row(
    pivot_name: &str,
    field: &PivotCalculatedField,
    ast: &FormulaExpr,
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
) -> Result<PivotValue> {
    let materialized = materialize_calculated_expr(pivot_name, field, ast, columns, row, lookup)?;
    let value = evaluate(&materialized, &EvaluationContext::simple()).map_err(|error| {
        Error::other(format!(
            "pivot table {pivot_name} calculated field {} evaluation failed: {error}",
            field.name
        ))
    })?;
    Ok(formula_value_to_pivot_value(value))
}

fn materialize_calculated_expr(
    pivot_name: &str,
    field: &PivotCalculatedField,
    expr: &FormulaExpr,
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
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
            if let Some(name) = structured_ref_field_name(reference) {
                calculated_field_value_expr(pivot_name, field, name, columns, row, lookup)?
            } else {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} calculated field {} uses an unsupported structured reference",
                    field.name
                )));
            }
        }
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::ExternalRef(_) => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated field {} uses workbook references, which are not valid pivot source-field references",
                field.name
            )));
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op: *op,
            left: Box::new(materialize_calculated_expr(
                pivot_name, field, left, columns, row, lookup,
            )?),
            right: Box::new(materialize_calculated_expr(
                pivot_name, field, right, columns, row, lookup,
            )?),
        },
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp {
            op: *op,
            operand: Box::new(materialize_calculated_expr(
                pivot_name, field, operand, columns, row, lookup,
            )?),
        },
        FormulaExpr::Function { name, args } => FormulaExpr::Function {
            name: name.clone(),
            args: materialize_calculated_args(pivot_name, field, args, columns, row, lookup)?,
        },
        FormulaExpr::ExternalFunction { book, name, args } => FormulaExpr::ExternalFunction {
            book: book.clone(),
            name: name.clone(),
            args: materialize_calculated_args(pivot_name, field, args, columns, row, lookup)?,
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
                )?);
            }
            FormulaExpr::Array(materialized_rows)
        }
    })
}

fn materialize_calculated_args(
    pivot_name: &str,
    field: &PivotCalculatedField,
    args: &[FormulaExpr],
    columns: &[EncodedColumn],
    row: usize,
    lookup: &AHashMap<String, usize>,
) -> Result<Vec<FormulaExpr>> {
    args.iter()
        .map(|arg| materialize_calculated_expr(pivot_name, field, arg, columns, row, lookup))
        .collect()
}

fn calculated_field_value_expr(
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

fn structured_ref_field_name(reference: &StructuredReference) -> Option<&str> {
    if reference.table.is_some() {
        return None;
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

fn pivot_value_to_formula_expr(value: &PivotValue) -> FormulaExpr {
    match value {
        PivotValue::Blank => FormulaExpr::Empty,
        PivotValue::Boolean(value) => FormulaExpr::Boolean(*value),
        PivotValue::Number(value) => FormulaExpr::Number(*value),
        PivotValue::String(value) => FormulaExpr::String(value.clone()),
        PivotValue::Error(value) => FormulaExpr::Error(*value),
    }
}

fn formula_value_to_pivot_value(value: FormulaValue) -> PivotValue {
    match value {
        FormulaValue::Empty => PivotValue::Blank,
        FormulaValue::Boolean(value) => PivotValue::Boolean(value),
        FormulaValue::Number(value) => PivotValue::Number(value),
        FormulaValue::String(value) => PivotValue::String(value),
        FormulaValue::Error(value) => PivotValue::Error(value),
        FormulaValue::Array { .. } => PivotValue::Error(CellError::Value),
    }
}

fn grouped_column(
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
    }
}

fn grouped_number_column(
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

    let mut column = EncodedColumn::with_capacity(snapshot.row_count);
    for row in 0..snapshot.row_count {
        column.push(group_number_value(
            snapshot.value(row, field_index),
            effective_start,
            end,
            interval,
        ));
    }
    Ok(column)
}

fn numeric_column_min(snapshot: &SourceSnapshot, field_index: usize) -> Option<f64> {
    (0..snapshot.row_count)
        .filter_map(|row| match snapshot.value(row, field_index) {
            PivotValue::Number(value) if value.is_finite() => Some(*value),
            _ => None,
        })
        .min_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal))
}

fn group_number_value(
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

fn normalize_group_number(value: f64) -> f64 {
    let rounded = value.round();
    if (value - rounded).abs() < 1e-10 {
        rounded
    } else {
        value
    }
}

fn format_group_number(value: f64) -> String {
    let value = normalize_group_number(value);
    if value.fract() == 0.0 {
        format!("{}", value as i64)
    } else {
        value.to_string()
    }
}

fn grouped_date_column(
    snapshot: &SourceSnapshot,
    field_index: usize,
    units: &[duke_sheets_core::PivotDateGroupUnit],
    date_1904: bool,
) -> EncodedColumn {
    let mut column = EncodedColumn::with_capacity(snapshot.row_count);
    let date_system = if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    };
    for row in 0..snapshot.row_count {
        column.push(group_date_value(
            snapshot.value(row, field_index),
            units,
            date_system,
        ));
    }
    column
}

fn group_date_value(
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

#[derive(Debug, Clone)]
struct EncodedColumn {
    values: Vec<u32>,
    dictionary: Vec<PivotValue>,
    lookup: AHashMap<PivotValue, u32>,
}

impl EncodedColumn {
    fn with_capacity(capacity: usize) -> Self {
        Self {
            values: Vec::with_capacity(capacity),
            dictionary: Vec::new(),
            lookup: AHashMap::new(),
        }
    }

    fn push(&mut self, value: PivotValue) {
        let id = if let Some(id) = self.lookup.get(&value) {
            *id
        } else {
            let id = self.dictionary.len() as u32;
            self.dictionary.push(value.clone());
            self.lookup.insert(value, id);
            id
        };
        self.values.push(id);
    }

    fn id_at(&self, row: usize) -> u32 {
        self.values[row]
    }

    fn value(&self, row: usize) -> &PivotValue {
        self.value_by_id(self.id_at(row))
    }

    fn value_by_id(&self, id: u32) -> &PivotValue {
        &self.dictionary[id as usize]
    }

    fn id_for_value(&self, value: &PivotValue) -> Option<u32> {
        self.lookup.get(value).copied()
    }
}

fn normalize_supplied_headers(headers: &[String], col_count: usize) -> Vec<String> {
    (0..col_count)
        .map(|index| {
            headers
                .get(index)
                .filter(|header| !header.trim().is_empty())
                .cloned()
                .unwrap_or_else(|| format!("Column{}", index + 1))
        })
        .collect()
}

fn read_headers_from_sheet(worksheet: &Worksheet, range: CellRange) -> Result<Vec<String>> {
    (range.start.col..=range.end.col)
        .map(|col| {
            let value = effective_pivot_value(worksheet, range.start.row, col);
            let header = value.to_string();
            if header.trim().is_empty() {
                Err(Error::other(format!(
                    "pivot source header cannot be blank at {}",
                    CellAddress::new(range.start.row, col)
                )))
            } else {
                Ok(header)
            }
        })
        .collect()
}

fn validate_headers(headers: &[String]) -> Result<()> {
    let mut seen = AHashSet::new();
    for header in headers {
        if header.trim().is_empty() {
            return Err(Error::other("pivot source headers cannot be blank"));
        }
        let key = header.to_lowercase();
        if !seen.insert(key) {
            return Err(Error::other(format!(
                "pivot source header is duplicated: {header}"
            )));
        }
    }
    Ok(())
}

fn effective_pivot_value(worksheet: &Worksheet, row: u32, col: u16) -> PivotValue {
    worksheet
        .get_calculated_value_at(row, col)
        .map(PivotValue::from_cell_value)
        .unwrap_or_else(|| PivotValue::from_cell_value(&worksheet.get_value_at(row, col)))
}

#[derive(Debug, Clone)]
struct CompiledPivotPlan {
    row_indexes: Vec<usize>,
    column_indexes: Vec<usize>,
    page_indexes: Vec<usize>,
    row_fields: Vec<PivotField>,
    column_fields: Vec<PivotField>,
    page_fields: Vec<PivotField>,
    measure_indexes: Vec<usize>,
    measures: Vec<PivotMeasure>,
    filters: Vec<CompiledFilter>,
    aggregate_filters: Vec<CompiledAggregateFilter>,
}

impl CompiledPivotPlan {
    fn compile(pivot: &PivotTable, snapshot: &SourceSnapshot) -> Result<Self> {
        if pivot.measures.is_empty() {
            return Err(Error::other(format!(
                "pivot table {} must contain at least one measure",
                pivot.name
            )));
        }
        let row_indexes = compile_axis_fields("row", &pivot.name, &pivot.rows, snapshot)?;
        let column_indexes = compile_axis_fields("column", &pivot.name, &pivot.columns, snapshot)?;
        let page_indexes = compile_axis_fields("page", &pivot.name, &pivot.page_fields, snapshot)?;

        let mut measure_indexes = Vec::with_capacity(pivot.measures.len());
        for measure in &pivot.measures {
            validate_show_as(
                &pivot.name,
                snapshot,
                &row_indexes,
                &column_indexes,
                &measure.show_as,
            )?;
            measure_indexes.push(field_index(snapshot, &measure.field.name, &pivot.name)?);
        }

        let mut filters = Vec::new();
        let mut aggregate_filters = Vec::new();
        for filter in &pivot.filters {
            match filter {
                PivotFilter::FieldItems { .. } | PivotFilter::Label { .. } => {
                    filters.push(CompiledFilter::compile(filter, snapshot, &pivot.name)?);
                }
                PivotFilter::Value { .. } | PivotFilter::TopN { .. } => {
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

        Ok(Self {
            row_indexes,
            column_indexes,
            page_indexes,
            row_fields: pivot.rows.clone(),
            column_fields: pivot.columns.clone(),
            page_fields: pivot.page_fields.clone(),
            measure_indexes,
            measures: pivot.measures.clone(),
            filters,
            aggregate_filters,
        })
    }
}

fn compile_axis_fields(
    axis_name: &str,
    pivot_name: &str,
    fields: &[PivotField],
    snapshot: &SourceSnapshot,
) -> Result<Vec<usize>> {
    fields
        .iter()
        .map(|field| {
            field_index(snapshot, &field.field.name, pivot_name).map_err(|_| {
                Error::other(format!(
                    "pivot table {pivot_name} references unknown {axis_name} field: {}",
                    field.field.name
                ))
            })
        })
        .collect()
}

fn validate_show_as(
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
        | PivotShowAs::Index => Ok(()),
        PivotShowAs::RunningTotal { base_field }
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

fn validate_base_field(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    column_indexes: &[usize],
    base_field: &str,
) -> Result<usize> {
    let field_index = field_index(snapshot, base_field, pivot_name)?;
    if row_indexes.contains(&field_index) || column_indexes.contains(&field_index) {
        Ok(field_index)
    } else {
        Err(Error::other(format!(
            "pivot table {pivot_name} uses show-as base field {base_field}, but that field is not on a row or column axis"
        )))
    }
}

fn field_index(snapshot: &SourceSnapshot, field_name: &str, pivot_name: &str) -> Result<usize> {
    snapshot.field_index(field_name).ok_or_else(|| {
        Error::other(format!(
            "pivot table {pivot_name} references unknown source field: {field_name}"
        ))
    })
}

#[derive(Debug, Clone)]
enum CompiledFilter {
    Items {
        field_index: usize,
        allowed_ids: AHashSet<u32>,
    },
    Label {
        field_index: usize,
        operator: PivotFilterOperator,
        value: String,
    },
}

impl CompiledFilter {
    fn compile(filter: &PivotFilter, snapshot: &SourceSnapshot, pivot_name: &str) -> Result<Self> {
        match filter {
            PivotFilter::FieldItems {
                field,
                allowed_items,
            } => {
                let field_index = field_index(snapshot, &field.name, pivot_name)?;
                let allowed_ids = allowed_items
                    .iter()
                    .filter_map(|value| snapshot.columns[field_index].id_for_value(value))
                    .collect();
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
                field_index: field_index(snapshot, &field.name, pivot_name)?,
                operator: *operator,
                value: value.clone(),
            }),
            PivotFilter::Value { .. } | PivotFilter::TopN { .. } => Err(Error::other(format!(
                "pivot table {pivot_name} tried to compile an aggregate filter as a row filter"
            ))),
            PivotFilter::Unsupported { kind, .. } => Err(Error::other(format!(
                "pivot table {pivot_name} contains unsupported filter: {kind}"
            ))),
        }
    }

    fn matches(&self, snapshot: &SourceSnapshot, row: usize) -> bool {
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
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum AggregateFilterAxis {
    Row,
    Column,
}

#[derive(Debug, Clone)]
enum CompiledAggregateFilter {
    Value {
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
        operator: PivotFilterOperator,
        value: f64,
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
    fn compile(
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
                let field_index = field_index(snapshot, &field.name, pivot_name)?;
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
            PivotFilter::TopN {
                field,
                measure,
                n,
                top,
                percent,
            } => {
                let field_index = field_index(snapshot, &field.name, pivot_name)?;
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

    fn axis(&self) -> AggregateFilterAxis {
        match self {
            Self::Value { axis, .. } | Self::TopN { axis, .. } => *axis,
        }
    }

    fn field_position(&self) -> usize {
        match self {
            Self::Value { field_position, .. } | Self::TopN { field_position, .. } => {
                *field_position
            }
        }
    }

    fn allowed_item_ids(&self, aggregation: &PivotAggregation) -> AHashSet<u32> {
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
            Self::TopN {
                n, top, percent, ..
            } => top_n_item_ids(item_states, self.aggregate(), *n, *top, *percent),
        }
    }

    fn measure_index(&self) -> usize {
        match self {
            Self::Value { measure_index, .. } | Self::TopN { measure_index, .. } => *measure_index,
        }
    }

    fn aggregate(&self) -> PivotAggregate {
        match self {
            Self::Value { aggregate, .. } | Self::TopN { aggregate, .. } => *aggregate,
        }
    }
}

fn aggregate_filter_axis(
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

fn measure_index_for_filter(
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

fn label_filter_matches(actual: &str, operator: PivotFilterOperator, expected: &str) -> bool {
    let actual_folded = actual.to_lowercase();
    let expected_folded = expected.to_lowercase();
    match operator {
        PivotFilterOperator::Equals => actual_folded == expected_folded,
        PivotFilterOperator::NotEquals => actual_folded != expected_folded,
        PivotFilterOperator::LessThan => actual_folded < expected_folded,
        PivotFilterOperator::LessThanOrEqual => actual_folded <= expected_folded,
        PivotFilterOperator::GreaterThan => actual_folded > expected_folded,
        PivotFilterOperator::GreaterThanOrEqual => actual_folded >= expected_folded,
        PivotFilterOperator::BeginsWith => actual_folded.starts_with(&expected_folded),
        PivotFilterOperator::DoesNotBeginWith => !actual_folded.starts_with(&expected_folded),
        PivotFilterOperator::EndsWith => actual_folded.ends_with(&expected_folded),
        PivotFilterOperator::DoesNotEndWith => !actual_folded.ends_with(&expected_folded),
        PivotFilterOperator::Contains => actual_folded.contains(&expected_folded),
        PivotFilterOperator::DoesNotContain => !actual_folded.contains(&expected_folded),
    }
}

fn numeric_filter_matches(actual: f64, operator: PivotFilterOperator, expected: f64) -> bool {
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

fn top_n_item_ids(
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

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct GroupKey {
    rows: Vec<u32>,
    columns: Vec<u32>,
}

#[derive(Debug, Clone)]
struct PivotAggregation {
    groups: AHashMap<GroupKey, Vec<AggregateState>>,
    group_order: Vec<GroupKey>,
    row_totals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    row_order: Vec<Vec<u32>>,
    column_totals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    column_order: Vec<Vec<u32>>,
    grand_totals: Vec<AggregateState>,
    matched_rows: usize,
}

impl PivotAggregation {
    fn aggregate(snapshot: &SourceSnapshot, plan: &CompiledPivotPlan) -> Self {
        #[cfg(feature = "parallel")]
        {
            if snapshot.row_count >= PARALLEL_ROW_THRESHOLD {
                return Self::aggregate_parallel(snapshot, plan);
            }
        }

        Self::aggregate_range(snapshot, plan, 0, snapshot.row_count)
    }

    fn aggregate_range(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        start: usize,
        end: usize,
    ) -> Self {
        let mut aggregation = Self {
            groups: AHashMap::new(),
            group_order: Vec::new(),
            row_totals: AHashMap::new(),
            row_order: Vec::new(),
            column_totals: AHashMap::new(),
            column_order: Vec::new(),
            grand_totals: default_states(&plan.measures),
            matched_rows: 0,
        };

        for row in start..end {
            if !plan
                .filters
                .iter()
                .all(|filter| filter.matches(snapshot, row))
            {
                continue;
            }
            aggregation.ingest_row(snapshot, plan, row);
        }

        aggregation
    }

    #[cfg(feature = "parallel")]
    fn aggregate_parallel(snapshot: &SourceSnapshot, plan: &CompiledPivotPlan) -> Self {
        let chunks = (0..snapshot.row_count)
            .step_by(PARALLEL_CHUNK_SIZE)
            .map(|start| (start, (start + PARALLEL_CHUNK_SIZE).min(snapshot.row_count)))
            .collect::<Vec<_>>();

        let partials = chunks
            .par_iter()
            .map(|(start, end)| Self::aggregate_range(snapshot, plan, *start, *end))
            .collect::<Vec<_>>();

        let mut merged = Self {
            groups: AHashMap::new(),
            group_order: Vec::new(),
            row_totals: AHashMap::new(),
            row_order: Vec::new(),
            column_totals: AHashMap::new(),
            column_order: Vec::new(),
            grand_totals: default_states(&plan.measures),
            matched_rows: 0,
        };

        for partial in partials {
            merged.merge_from(partial);
        }

        merged
    }

    fn ingest_row(&mut self, snapshot: &SourceSnapshot, plan: &CompiledPivotPlan, row: usize) {
        self.matched_rows += 1;
        let row_key = encoded_key(snapshot, &plan.row_indexes, row);
        let column_key = encoded_key(snapshot, &plan.column_indexes, row);
        let group_key = GroupKey {
            rows: row_key.clone(),
            columns: column_key.clone(),
        };

        if !self.groups.contains_key(&group_key) {
            self.group_order.push(group_key.clone());
            self.groups
                .insert(group_key.clone(), default_states(&plan.measures));
        }
        update_states(
            self.groups.get_mut(&group_key).expect("group was inserted"),
            snapshot,
            plan,
            row,
        );

        if !self.row_totals.contains_key(&row_key) {
            self.row_order.push(row_key.clone());
            self.row_totals
                .insert(row_key.clone(), default_states(&plan.measures));
        }
        update_states(
            self.row_totals
                .get_mut(&row_key)
                .expect("row total was inserted"),
            snapshot,
            plan,
            row,
        );

        if !self.column_totals.contains_key(&column_key) {
            self.column_order.push(column_key.clone());
            self.column_totals
                .insert(column_key.clone(), default_states(&plan.measures));
        }
        update_states(
            self.column_totals
                .get_mut(&column_key)
                .expect("column total was inserted"),
            snapshot,
            plan,
            row,
        );

        update_states(&mut self.grand_totals, snapshot, plan, row);
    }

    fn apply_aggregate_filters(&mut self, plan: &CompiledPivotPlan) {
        for filter in &plan.aggregate_filters {
            let allowed_item_ids = filter.allowed_item_ids(self);
            self.retain_axis_items(
                filter.axis(),
                filter.field_position(),
                &allowed_item_ids,
                plan,
            );
        }
    }

    fn retain_axis_items(
        &mut self,
        axis: AggregateFilterAxis,
        field_position: usize,
        allowed_item_ids: &AHashSet<u32>,
        plan: &CompiledPivotPlan,
    ) {
        match axis {
            AggregateFilterAxis::Row => {
                self.row_order
                    .retain(|key| allowed_item_ids.contains(&key[field_position]));
                self.group_order
                    .retain(|key| allowed_item_ids.contains(&key.rows[field_position]));
                self.groups
                    .retain(|key, _| allowed_item_ids.contains(&key.rows[field_position]));
            }
            AggregateFilterAxis::Column => {
                self.column_order
                    .retain(|key| allowed_item_ids.contains(&key[field_position]));
                self.group_order
                    .retain(|key| allowed_item_ids.contains(&key.columns[field_position]));
                self.groups
                    .retain(|key, _| allowed_item_ids.contains(&key.columns[field_position]));
            }
        }
        self.rebuild_totals_from_groups(plan);
    }

    fn item_states_for_filter(
        &self,
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
    ) -> AHashMap<u32, AggregateState> {
        let mut item_states = AHashMap::new();
        let (order, totals) = match axis {
            AggregateFilterAxis::Row => (&self.row_order, &self.row_totals),
            AggregateFilterAxis::Column => (&self.column_order, &self.column_totals),
        };

        for key in order {
            let Some(states) = totals.get(key) else {
                continue;
            };
            let Some(state) = states.get(measure_index) else {
                continue;
            };
            item_states
                .entry(key[field_position])
                .or_insert_with(|| AggregateState::new(aggregate))
                .merge(state);
        }
        item_states
    }

    fn rebuild_totals_from_groups(&mut self, plan: &CompiledPivotPlan) {
        self.row_totals.clear();
        self.column_totals.clear();
        self.grand_totals = default_states(&plan.measures);

        for key in &self.group_order {
            let Some(states) = self.groups.get(key) else {
                continue;
            };

            let row_states = self
                .row_totals
                .entry(key.rows.clone())
                .or_insert_with(|| default_states(&plan.measures));
            merge_state_slices(row_states, states);

            let column_states = self
                .column_totals
                .entry(key.columns.clone())
                .or_insert_with(|| default_states(&plan.measures));
            merge_state_slices(column_states, states);

            merge_state_slices(&mut self.grand_totals, states);
        }

        self.row_order
            .retain(|key| self.row_totals.contains_key(key));
        self.column_order
            .retain(|key| self.column_totals.contains_key(key));
    }

    fn sort_orders(&mut self, snapshot: &SourceSnapshot, plan: &CompiledPivotPlan) {
        sort_key_order(
            &mut self.row_order,
            &plan.row_indexes,
            &plan.row_fields,
            snapshot,
        );
        sort_key_order(
            &mut self.column_order,
            &plan.column_indexes,
            &plan.column_fields,
            snapshot,
        );
        self.group_order.sort_by(|a, b| {
            compare_encoded_key(
                &a.rows,
                &b.rows,
                &plan.row_indexes,
                &plan.row_fields,
                snapshot,
            )
            .then_with(|| {
                compare_encoded_key(
                    &a.columns,
                    &b.columns,
                    &plan.column_indexes,
                    &plan.column_fields,
                    snapshot,
                )
            })
        });
    }

    #[cfg(feature = "parallel")]
    fn merge_from(&mut self, other: Self) {
        self.matched_rows += other.matched_rows;
        merge_state_slices(&mut self.grand_totals, &other.grand_totals);

        for key in other.group_order {
            let states = other
                .groups
                .get(&key)
                .expect("ordered group key must exist")
                .clone();
            merge_ordered_bucket(&mut self.groups, &mut self.group_order, key, states);
        }

        for key in other.row_order {
            let states = other
                .row_totals
                .get(&key)
                .expect("ordered row key must exist")
                .clone();
            merge_ordered_bucket(&mut self.row_totals, &mut self.row_order, key, states);
        }

        for key in other.column_order {
            let states = other
                .column_totals
                .get(&key)
                .expect("ordered column key must exist")
                .clone();
            merge_ordered_bucket(&mut self.column_totals, &mut self.column_order, key, states);
        }
    }
}

#[cfg(feature = "parallel")]
fn merge_ordered_bucket<K>(
    map: &mut AHashMap<K, Vec<AggregateState>>,
    order: &mut Vec<K>,
    key: K,
    states: Vec<AggregateState>,
) where
    K: Eq + Hash + Clone,
{
    if let Some(existing) = map.get_mut(&key) {
        merge_state_slices(existing, &states);
    } else {
        order.push(key.clone());
        map.insert(key, states);
    }
}

fn merge_state_slices(target: &mut [AggregateState], source: &[AggregateState]) {
    for (target, source) in target.iter_mut().zip(source.iter()) {
        target.merge(source);
    }
}

fn encoded_key(snapshot: &SourceSnapshot, field_indexes: &[usize], row: usize) -> Vec<u32> {
    field_indexes
        .iter()
        .map(|field_index| snapshot.columns[*field_index].id_at(row))
        .collect()
}

fn default_states(measures: &[PivotMeasure]) -> Vec<AggregateState> {
    measures
        .iter()
        .map(|measure| AggregateState::new(measure.aggregate))
        .collect()
}

fn update_states(
    states: &mut [AggregateState],
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row: usize,
) {
    for ((state, field_index), measure) in states
        .iter_mut()
        .zip(plan.measure_indexes.iter())
        .zip(plan.measures.iter())
    {
        state.update(snapshot.value(row, *field_index), measure.aggregate);
    }
}

#[derive(Debug, Clone)]
struct AggregateState {
    count_non_blank: u64,
    count_numbers: u64,
    sum: f64,
    sum_sq: f64,
    product: f64,
    min: Option<f64>,
    max: Option<f64>,
}

impl AggregateState {
    fn new(_aggregate: PivotAggregate) -> Self {
        Self {
            count_non_blank: 0,
            count_numbers: 0,
            sum: 0.0,
            sum_sq: 0.0,
            product: 1.0,
            min: None,
            max: None,
        }
    }

    fn update(&mut self, value: &PivotValue, _aggregate: PivotAggregate) {
        if !value.is_blank() {
            self.count_non_blank += 1;
        }

        let Some(number) = pivot_number(value) else {
            return;
        };

        self.count_numbers += 1;
        self.sum += number;
        self.sum_sq += number * number;
        self.product *= number;
        self.min = Some(self.min.map_or(number, |current| current.min(number)));
        self.max = Some(self.max.map_or(number, |current| current.max(number)));
    }

    fn finalize(&self, aggregate: PivotAggregate) -> CellValue {
        self.finalize_number(aggregate)
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty)
    }

    fn finalize_number(&self, aggregate: PivotAggregate) -> Option<f64> {
        match aggregate {
            PivotAggregate::Sum => Some(self.sum),
            PivotAggregate::Count => Some(self.count_non_blank as f64),
            PivotAggregate::CountNumbers => Some(self.count_numbers as f64),
            PivotAggregate::Average => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(self.sum / self.count_numbers as f64)
                }
            }
            PivotAggregate::Max => self.max,
            PivotAggregate::Min => self.min,
            PivotAggregate::Product => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(self.product)
                }
            }
            PivotAggregate::StdDev => {
                if self.count_numbers < 2 {
                    None
                } else {
                    Some(sample_variance(self).sqrt())
                }
            }
            PivotAggregate::StdDevP => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(population_variance(self).sqrt())
                }
            }
            PivotAggregate::Var => {
                if self.count_numbers < 2 {
                    None
                } else {
                    Some(sample_variance(self))
                }
            }
            PivotAggregate::VarP => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(population_variance(self))
                }
            }
        }
    }

    fn merge(&mut self, other: &Self) {
        self.count_non_blank += other.count_non_blank;
        self.count_numbers += other.count_numbers;
        self.sum += other.sum;
        self.sum_sq += other.sum_sq;
        self.product *= other.product;
        self.min = match (self.min, other.min) {
            (Some(left), Some(right)) => Some(left.min(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
        self.max = match (self.max, other.max) {
            (Some(left), Some(right)) => Some(left.max(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
    }
}

fn pivot_number(value: &PivotValue) -> Option<f64> {
    match value {
        PivotValue::Number(value) => Some(*value),
        _ => None,
    }
}

fn population_variance(state: &AggregateState) -> f64 {
    let count = state.count_numbers as f64;
    ((state.sum_sq - (state.sum * state.sum / count)) / count).max(0.0)
}

fn sample_variance(state: &AggregateState) -> f64 {
    let count = state.count_numbers as f64;
    ((state.sum_sq - (state.sum * state.sum / count)) / (count - 1.0)).max(0.0)
}

fn sort_key_order(
    order: &mut [Vec<u32>],
    field_indexes: &[usize],
    fields: &[PivotField],
    snapshot: &SourceSnapshot,
) {
    if fields
        .iter()
        .all(|field| matches!(field.sort, PivotSort::None))
    {
        return;
    }

    order.sort_by(|a, b| compare_encoded_key(a, b, field_indexes, fields, snapshot));
}

fn compare_encoded_key(
    left: &[u32],
    right: &[u32],
    field_indexes: &[usize],
    fields: &[PivotField],
    snapshot: &SourceSnapshot,
) -> Ordering {
    for (index, field_index) in field_indexes.iter().enumerate() {
        let sort = fields
            .get(index)
            .map(|field| field.sort)
            .unwrap_or(PivotSort::Ascending);
        if matches!(sort, PivotSort::None) {
            continue;
        }

        let ordering = compare_pivot_values(
            snapshot.value_by_id(*field_index, left[index]),
            snapshot.value_by_id(*field_index, right[index]),
        );

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

fn compare_pivot_values(left: &PivotValue, right: &PivotValue) -> Ordering {
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

fn pivot_value_rank(value: &PivotValue) -> u8 {
    match value {
        PivotValue::Blank => 0,
        PivotValue::Boolean(_) => 1,
        PivotValue::Number(_) => 2,
        PivotValue::String(_) => 3,
        PivotValue::Error(_) => 4,
    }
}

#[derive(Debug, Clone)]
struct RenderedPivot {
    cells: Vec<Vec<CellValue>>,
    range: CellRange,
    source_rows: usize,
}

impl RenderedPivot {
    fn cell_count(&self) -> usize {
        self.cells.iter().map(Vec::len).sum()
    }
}

fn render_pivot(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Result<RenderedPivot> {
    let mut cells = if plan.column_indexes.is_empty() {
        render_without_column_fields(pivot, snapshot, plan, aggregation)
    } else {
        render_with_column_fields(pivot, snapshot, plan, aggregation)
    };
    prepend_page_fields(&mut cells, pivot, snapshot, plan);

    let width = cells.iter().map(Vec::len).max().unwrap_or(1).max(1);
    for row in &mut cells {
        row.resize(width, CellValue::Empty);
    }
    if cells.is_empty() {
        cells.push(vec![CellValue::Empty; width]);
    }

    let range = output_range(pivot.target, cells.len(), width)?;
    Ok(RenderedPivot {
        cells,
        range,
        source_rows: snapshot.row_count,
    })
}

fn render_without_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let mut header = plan
        .row_indexes
        .iter()
        .map(|index| CellValue::string(&snapshot.headers[*index]))
        .collect::<Vec<_>>();
    header.extend(
        plan.measures
            .iter()
            .map(|measure| CellValue::string(measure.caption())),
    );
    cells.push(header);

    let empty_column_key = Vec::new();
    for row_key in &aggregation.row_order {
        let mut row = decode_key_cells(snapshot, &plan.row_indexes, row_key);
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
            aggregation.row_totals.get(row_key),
            aggregation.column_totals.get(&empty_column_key),
            &aggregation.grand_totals,
            &context,
        ));
        cells.push(row);
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = grand_total_label_row(plan.row_indexes.len());
        let context = ShowAsContext {
            snapshot,
            plan,
            aggregation,
            row_key: None,
            column_key: None,
        };
        row.extend(finalize_state_slice_with_context(
            &aggregation.grand_totals,
            &plan.measures,
            &aggregation.grand_totals,
            &aggregation.grand_totals,
            &aggregation.grand_totals,
            &context,
        ));
        cells.push(row);
    }

    cells
}

fn render_with_column_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Vec<CellValue>> {
    let mut cells = Vec::new();
    let mut header = plan
        .row_indexes
        .iter()
        .map(|index| CellValue::string(&snapshot.headers[*index]))
        .collect::<Vec<_>>();

    for column_key in &aggregation.column_order {
        let column_label = key_label(snapshot, &plan.column_indexes, column_key);
        for measure in &plan.measures {
            header.push(CellValue::string(measure_column_caption(
                &column_label,
                measure,
                plan.measures.len(),
            )));
        }
    }
    if pivot.layout.show_column_grand_totals {
        for measure in &plan.measures {
            header.push(CellValue::string(grand_total_measure_caption(
                measure,
                plan.measures.len(),
            )));
        }
    }
    cells.push(header);

    for row_key in &aggregation.row_order {
        let mut row = decode_key_cells(snapshot, &plan.row_indexes, row_key);
        for column_key in &aggregation.column_order {
            let key = GroupKey {
                rows: row_key.clone(),
                columns: column_key.clone(),
            };
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: Some(column_key),
            };
            row.extend(finalize_states_with_context(
                aggregation.groups.get(&key),
                &plan.measures,
                aggregation.row_totals.get(row_key),
                aggregation.column_totals.get(column_key),
                &aggregation.grand_totals,
                &context,
            ));
        }
        if pivot.layout.show_column_grand_totals {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: Some(row_key),
                column_key: None,
            };
            row.extend(finalize_states_with_context(
                aggregation.row_totals.get(row_key),
                &plan.measures,
                aggregation.row_totals.get(row_key),
                Some(&aggregation.grand_totals),
                &aggregation.grand_totals,
                &context,
            ));
        }
        cells.push(row);
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = grand_total_label_row(plan.row_indexes.len());
        for column_key in &aggregation.column_order {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: Some(column_key),
            };
            row.extend(finalize_states_with_context(
                aggregation.column_totals.get(column_key),
                &plan.measures,
                Some(&aggregation.grand_totals),
                aggregation.column_totals.get(column_key),
                &aggregation.grand_totals,
                &context,
            ));
        }
        if pivot.layout.show_column_grand_totals {
            let context = ShowAsContext {
                snapshot,
                plan,
                aggregation,
                row_key: None,
                column_key: None,
            };
            row.extend(finalize_state_slice_with_context(
                &aggregation.grand_totals,
                &plan.measures,
                &aggregation.grand_totals,
                &aggregation.grand_totals,
                &aggregation.grand_totals,
                &context,
            ));
        }
        cells.push(row);
    }

    cells
}

fn prepend_page_fields(
    cells: &mut Vec<Vec<CellValue>>,
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
) {
    if plan.page_fields.is_empty() {
        return;
    }

    let mut rows = Vec::with_capacity(plan.page_fields.len() + 1);
    for (field, field_index) in plan.page_fields.iter().zip(plan.page_indexes.iter()) {
        rows.push(vec![
            CellValue::string(&field.field.name),
            CellValue::string(page_field_caption(
                pivot,
                snapshot,
                *field_index,
                &field.field.name,
            )),
        ]);
    }
    rows.push(Vec::new());
    rows.append(cells);
    *cells = rows;
}

fn page_field_caption(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    field_index: usize,
    field_name: &str,
) -> String {
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    }) else {
        return "(All)".to_string();
    };

    match allowed_items.as_slice() {
        [] => "(All)".to_string(),
        [item] => item.to_string(),
        _ => {
            let selected_count = allowed_items
                .iter()
                .filter(|item| snapshot.columns[field_index].id_for_value(item).is_some())
                .count();
            if selected_count == snapshot.columns[field_index].dictionary.len() {
                "(All)".to_string()
            } else {
                "(Multiple Items)".to_string()
            }
        }
    }
}

fn decode_key_cells(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    key: &[u32],
) -> Vec<CellValue> {
    field_indexes
        .iter()
        .zip(key.iter())
        .map(|(field_index, id)| snapshot.value_by_id(*field_index, *id).to_cell_value())
        .collect()
}

fn grand_total_label_row(label_width: usize) -> Vec<CellValue> {
    if label_width == 0 {
        Vec::new()
    } else {
        let mut row = vec![CellValue::Empty; label_width];
        row[0] = CellValue::string("Grand Total");
        row
    }
}

#[derive(Debug, Clone, Copy)]
struct ShowAsContext<'a> {
    snapshot: &'a SourceSnapshot,
    plan: &'a CompiledPivotPlan,
    aggregation: &'a PivotAggregation,
    row_key: Option<&'a Vec<u32>>,
    column_key: Option<&'a Vec<u32>>,
}

fn finalize_states_with_context(
    states: Option<&Vec<AggregateState>>,
    measures: &[PivotMeasure],
    row_total: Option<&Vec<AggregateState>>,
    column_total: Option<&Vec<AggregateState>>,
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
) -> Vec<CellValue> {
    states
        .map(|states| {
            finalize_state_slice_with_context(
                states,
                measures,
                row_total.map(Vec::as_slice).unwrap_or(&[]),
                column_total.map(Vec::as_slice).unwrap_or(&[]),
                grand_total,
                context,
            )
        })
        .unwrap_or_else(|| vec![CellValue::Empty; measures.len()])
}

fn finalize_state_slice_with_context(
    states: &[AggregateState],
    measures: &[PivotMeasure],
    row_total: &[AggregateState],
    column_total: &[AggregateState],
    grand_total: &[AggregateState],
    context: &ShowAsContext<'_>,
) -> Vec<CellValue> {
    states
        .iter()
        .enumerate()
        .zip(measures.iter())
        .map(|((index, state), measure)| {
            finalize_measure_with_context(
                state,
                measure,
                state_number(row_total, index, measure.aggregate),
                state_number(column_total, index, measure.aggregate),
                state_number(grand_total, index, measure.aggregate),
                index,
                context,
            )
        })
        .collect()
}

fn finalize_measure_with_context(
    state: &AggregateState,
    measure: &PivotMeasure,
    row_total: Option<f64>,
    column_total: Option<f64>,
    grand_total: Option<f64>,
    measure_index: usize,
    context: &ShowAsContext<'_>,
) -> CellValue {
    match &measure.show_as {
        PivotShowAs::Normal => state.finalize(measure.aggregate),
        PivotShowAs::PercentOfGrandTotal => {
            percentage_cell(state.finalize_number(measure.aggregate), grand_total)
        }
        PivotShowAs::PercentOfRowTotal => {
            percentage_cell(state.finalize_number(measure.aggregate), row_total)
        }
        PivotShowAs::PercentOfColumnTotal => {
            percentage_cell(state.finalize_number(measure.aggregate), column_total)
        }
        PivotShowAs::Index => index_cell(
            state.finalize_number(measure.aggregate),
            row_total,
            column_total,
            grand_total,
        ),
        PivotShowAs::RunningTotal { base_field } => optional_number_cell(running_total_value(
            context,
            base_field.name.as_str(),
            measure_index,
            measure.aggregate,
        )),
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => {
            let current = state.finalize_number(measure.aggregate);
            let base = base_item_value(
                context,
                base_field.name.as_str(),
                base_item,
                measure_index,
                measure.aggregate,
            );
            optional_number_cell(current.zip(base).map(|(current, base)| current - base))
        }
        PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => {
            let current = state.finalize_number(measure.aggregate);
            let base = base_item_value(
                context,
                base_field.name.as_str(),
                base_item,
                measure_index,
                measure.aggregate,
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
            measure.aggregate,
            true,
        )),
        PivotShowAs::RankDescending { base_field } => optional_number_cell(rank_value(
            context,
            base_field.name.as_str(),
            measure_index,
            measure.aggregate,
            false,
        )),
    }
}

fn state_number(states: &[AggregateState], index: usize, aggregate: PivotAggregate) -> Option<f64> {
    states
        .get(index)
        .and_then(|state| state.finalize_number(aggregate))
}

fn percentage_cell(numerator: Option<f64>, denominator: Option<f64>) -> CellValue {
    match (numerator, denominator) {
        (Some(numerator), Some(denominator)) if denominator != 0.0 => {
            CellValue::Number(numerator / denominator)
        }
        _ => CellValue::Empty,
    }
}

fn index_cell(
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

fn optional_number_cell(value: Option<f64>) -> CellValue {
    value.map(CellValue::Number).unwrap_or(CellValue::Empty)
}

#[derive(Debug, Clone, Copy)]
enum ShowAsAxis {
    Row(usize),
    Column(usize),
}

fn show_as_axis(context: &ShowAsContext<'_>, base_field: &str) -> Option<ShowAsAxis> {
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

fn base_item_value(
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

fn running_total_value(
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

fn rank_value(
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

fn states_for_row_axis_key<'a>(
    context: &'a ShowAsContext<'_>,
    row_key: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.column_key {
        Some(column_key) => context.aggregation.groups.get(&GroupKey {
            rows: row_key.to_vec(),
            columns: column_key.clone(),
        }),
        None => context.aggregation.row_totals.get(row_key),
    }
}

fn states_for_column_axis_key<'a>(
    context: &'a ShowAsContext<'_>,
    column_key: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    match context.row_key {
        Some(row_key) => context.aggregation.groups.get(&GroupKey {
            rows: row_key.clone(),
            columns: column_key.to_vec(),
        }),
        None => context.aggregation.column_totals.get(column_key),
    }
}

fn same_peer_key(candidate: &[u32], current: &[u32], base_position: usize) -> bool {
    candidate.len() == current.len()
        && candidate
            .iter()
            .zip(current.iter())
            .enumerate()
            .all(|(index, (left, right))| index == base_position || left == right)
}

fn rank_precedes(value: f64, current: f64, ascending: bool) -> bool {
    if ascending {
        value < current
    } else {
        value > current
    }
}

fn key_label(snapshot: &SourceSnapshot, field_indexes: &[usize], key: &[u32]) -> String {
    field_indexes
        .iter()
        .zip(key.iter())
        .map(|(field_index, id)| snapshot.value_by_id(*field_index, *id).to_string())
        .collect::<Vec<_>>()
        .join(" | ")
}

fn measure_column_caption(
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

fn grand_total_measure_caption(measure: &PivotMeasure, measure_count: usize) -> String {
    if measure_count == 1 {
        "Grand Total".to_string()
    } else {
        format!("Grand Total {}", measure.caption())
    }
}

fn output_range(target: CellAddress, row_count: usize, col_count: usize) -> Result<CellRange> {
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

#[cfg(test)]
mod tests {
    use duke_sheets_core::{
        CellRange, PivotAggregate, PivotDateGroupUnit, PivotField, PivotFilter,
        PivotFilterOperator, PivotGrouping, PivotMeasure, PivotShowAs, PivotSort, PivotSource,
        PivotTable, PivotValue, Table, TableColumn, Workbook,
    };
    use pretty_assertions::assert_eq;
    use ssfmt::{date_serial::date_to_serial, DateSystem};

    use super::WorkbookPivotExt;

    fn number(workbook: &Workbook, address: &str) -> f64 {
        workbook
            .worksheet(0)
            .unwrap()
            .get_value(address)
            .unwrap()
            .as_number()
            .unwrap()
    }

    fn text(workbook: &Workbook, address: &str) -> String {
        workbook
            .worksheet(0)
            .unwrap()
            .get_value(address)
            .unwrap()
            .to_string()
    }

    fn assert_close(actual: f64, expected: f64) {
        assert!(
            (actual - expected).abs() < 1e-9,
            "expected {expected}, got {actual}"
        );
    }

    #[test]
    fn refreshes_sum_by_row_field() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", 15.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        let stats = workbook.refresh_pivots().unwrap();

        assert_eq!(stats.pivot_count, 1);
        assert_eq!(stats.pivots_refreshed, 1);
        assert_eq!(stats.source_rows, 3);
        assert_eq!(text(&workbook, "D1"), "Region");
        assert_eq!(text(&workbook, "E1"), "Revenue");
        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 25.0);
        assert_eq!(text(&workbook, "D3"), "West");
        assert_eq!(number(&workbook, "E3"), 20.0);
        assert_eq!(text(&workbook, "D4"), "Grand Total");
        assert_eq!(number(&workbook, "E4"), 45.0);
    }

    #[test]
    fn refreshes_sorted_row_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "North").unwrap();
        sheet.set_cell_value("B4", 15.0).unwrap();

        let mut region = PivotField::new("Region");
        region.sort = PivotSort::Descending;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row(region)
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "West");
        assert_eq!(text(&workbook, "D3"), "North");
        assert_eq!(text(&workbook, "D4"), "East");
    }

    #[test]
    fn refreshes_calculated_field_measure() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Units").unwrap();
        sheet.set_cell_value("C1", "Price").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 2.0).unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", 3.0).unwrap();
        sheet.set_cell_value("C3", 10.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", 7.0).unwrap();
        sheet.set_cell_value("C4", 3.0).unwrap();

        let pivot = PivotTable::builder("CalculatedRevenue")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .calculated_field("Revenue", "=Units*Price")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Region");
        assert_eq!(text(&workbook, "F1"), "Revenue");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 50.0);
        assert_eq!(text(&workbook, "E3"), "West");
        assert_eq!(number(&workbook, "F3"), 21.0);
        assert_eq!(text(&workbook, "E4"), "Grand Total");
        assert_eq!(number(&workbook, "F4"), 71.0);
    }

    #[test]
    fn refreshes_sequential_calculated_fields_with_structured_refs() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Gross Sales").unwrap();
        sheet.set_cell_value("C1", "Rate").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 100.0).unwrap();
        sheet.set_cell_value("C2", 0.1).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", 50.0).unwrap();
        sheet.set_cell_value("C3", 0.2).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", 80.0).unwrap();
        sheet.set_cell_value("C4", 0.25).unwrap();

        let pivot = PivotTable::builder("CalculatedCommission")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .calculated_field("Commission", "=[Gross Sales]*Rate")
            .calculated_field("CommissionWithFee", "=Commission+1")
            .named_measure("CommissionWithFee", PivotAggregate::Sum, "Commission")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 22.0);
        assert_eq!(text(&workbook, "E3"), "West");
        assert_eq!(number(&workbook, "F3"), 21.0);
        assert_eq!(text(&workbook, "E4"), "Grand Total");
        assert_eq!(number(&workbook, "F4"), 43.0);
    }

    #[test]
    fn refreshes_row_and_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Q1").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Q2").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Q1").unwrap();
        sheet.set_cell_value("C4", 7.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source(PivotSource::range(CellRange::parse("A1:C4").unwrap()))
            .target_address("E1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Region");
        assert_eq!(text(&workbook, "F1"), "Q1");
        assert_eq!(text(&workbook, "G1"), "Q2");
        assert_eq!(text(&workbook, "H1"), "Grand Total");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 10.0);
        assert_eq!(number(&workbook, "G2"), 5.0);
        assert_eq!(number(&workbook, "H2"), 15.0);
        assert_eq!(text(&workbook, "E3"), "West");
        assert_eq!(number(&workbook, "F3"), 7.0);
        assert_eq!(number(&workbook, "H3"), 7.0);
        assert_eq!(text(&workbook, "E4"), "Grand Total");
        assert_eq!(number(&workbook, "F4"), 17.0);
        assert_eq!(number(&workbook, "G4"), 5.0);
        assert_eq!(number(&workbook, "H4"), 22.0);
    }

    #[test]
    fn refreshes_percentage_show_as_calculations() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Q1").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Q2").unwrap();
        sheet.set_cell_value("C3", 30.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Q1").unwrap();
        sheet.set_cell_value("C4", 10.0).unwrap();
        sheet.set_cell_value("A5", "West").unwrap();
        sheet.set_cell_value("B5", "Q2").unwrap();
        sheet.set_cell_value("C5", 50.0).unwrap();

        let source = CellRange::parse("A1:C5").unwrap();
        let grand = PivotTable::builder("GrandPercent")
            .source_range(source)
            .target_address("E1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum)
                    .with_show_as(PivotShowAs::PercentOfGrandTotal),
            )
            .build()
            .unwrap();
        let row = PivotTable::builder("RowPercent")
            .source_range(source)
            .target_address("J1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum)
                    .with_show_as(PivotShowAs::PercentOfRowTotal),
            )
            .build()
            .unwrap();
        let column = PivotTable::builder("ColumnPercent")
            .source_range(source)
            .target_address("O1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum)
                    .with_show_as(PivotShowAs::PercentOfColumnTotal),
            )
            .build()
            .unwrap();
        sheet.add_pivot_table(grand).unwrap();
        sheet.add_pivot_table(row).unwrap();
        sheet.add_pivot_table(column).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_close(number(&workbook, "F2"), 0.1);
        assert_close(number(&workbook, "G2"), 0.3);
        assert_close(number(&workbook, "H2"), 0.4);
        assert_close(number(&workbook, "F4"), 0.2);
        assert_close(number(&workbook, "G4"), 0.8);
        assert_close(number(&workbook, "H4"), 1.0);

        assert_close(number(&workbook, "K2"), 0.25);
        assert_close(number(&workbook, "L2"), 0.75);
        assert_close(number(&workbook, "M2"), 1.0);
        assert_close(number(&workbook, "K3"), 1.0 / 6.0);
        assert_close(number(&workbook, "L3"), 5.0 / 6.0);
        assert_close(number(&workbook, "K4"), 0.2);
        assert_close(number(&workbook, "L4"), 0.8);

        assert_close(number(&workbook, "P2"), 0.5);
        assert_close(number(&workbook, "Q2"), 0.375);
        assert_close(number(&workbook, "R2"), 0.4);
        assert_close(number(&workbook, "P3"), 0.5);
        assert_close(number(&workbook, "Q3"), 0.625);
        assert_close(number(&workbook, "P4"), 1.0);
        assert_close(number(&workbook, "Q4"), 1.0);
    }

    #[test]
    fn refreshes_index_show_as_calculation() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Q1").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Q2").unwrap();
        sheet.set_cell_value("C3", 30.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Q1").unwrap();
        sheet.set_cell_value("C4", 20.0).unwrap();
        sheet.set_cell_value("A5", "West").unwrap();
        sheet.set_cell_value("B5", "Q2").unwrap();
        sheet.set_cell_value("C5", 40.0).unwrap();

        let pivot = PivotTable::builder("IndexShowAs")
            .source_range(CellRange::parse("A1:C5").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(PivotShowAs::Index),
            )
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_close(number(&workbook, "F2"), 10.0 * 100.0 / (40.0 * 30.0));
        assert_close(number(&workbook, "G2"), 30.0 * 100.0 / (40.0 * 70.0));
        assert_close(number(&workbook, "F3"), 20.0 * 100.0 / (60.0 * 30.0));
        assert_close(number(&workbook, "G3"), 40.0 * 100.0 / (60.0 * 70.0));
    }

    #[test]
    fn refreshes_numeric_grouping() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Amount").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", 2.0).unwrap();
        sheet.set_cell_value("B2", 1.0).unwrap();
        sheet.set_cell_value("A3", 7.0).unwrap();
        sheet.set_cell_value("B3", 2.0).unwrap();
        sheet.set_cell_value("A4", 12.0).unwrap();
        sheet.set_cell_value("B4", 3.0).unwrap();
        sheet.set_cell_value("A5", 17.0).unwrap();
        sheet.set_cell_value("B5", 4.0).unwrap();
        sheet.set_cell_value("A6", 25.0).unwrap();
        sheet.set_cell_value("B6", 5.0).unwrap();

        let pivot = PivotTable::builder("GroupedAmounts")
            .source_range(CellRange::parse("A1:B6").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Amount")
            .measure("Revenue", PivotAggregate::Sum)
            .grouping(PivotGrouping::Number {
                field: "Amount".into(),
                start: Some(0.0),
                end: None,
                interval: 10.0,
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(number(&workbook, "D2"), 0.0);
        assert_eq!(number(&workbook, "E2"), 3.0);
        assert_eq!(number(&workbook, "D3"), 10.0);
        assert_eq!(number(&workbook, "E3"), 7.0);
        assert_eq!(number(&workbook, "D4"), 20.0);
        assert_eq!(number(&workbook, "E4"), 5.0);
        assert_eq!(text(&workbook, "D5"), "Grand Total");
        assert_eq!(number(&workbook, "E5"), 15.0);
    }

    #[test]
    fn refreshes_date_grouping() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Date").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet
            .set_cell_value("A2", date_to_serial(2024, 1, 15, DateSystem::Date1900))
            .unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet
            .set_cell_value("A3", date_to_serial(2024, 1, 20, DateSystem::Date1900))
            .unwrap();
        sheet.set_cell_value("B3", 5.0).unwrap();
        sheet
            .set_cell_value("A4", date_to_serial(2024, 2, 1, DateSystem::Date1900))
            .unwrap();
        sheet.set_cell_value("B4", 7.0).unwrap();
        sheet
            .set_cell_value("A5", date_to_serial(2025, 1, 1, DateSystem::Date1900))
            .unwrap();
        sheet.set_cell_value("B5", 11.0).unwrap();

        let pivot = PivotTable::builder("GroupedDates")
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Date")
            .measure("Revenue", PivotAggregate::Sum)
            .grouping(PivotGrouping::Date {
                field: "Date".into(),
                units: vec![PivotDateGroupUnit::Years, PivotDateGroupUnit::Months],
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "2024-01");
        assert_eq!(number(&workbook, "E2"), 15.0);
        assert_eq!(text(&workbook, "D3"), "2024-02");
        assert_eq!(number(&workbook, "E3"), 7.0);
        assert_eq!(text(&workbook, "D4"), "2025-01");
        assert_eq!(number(&workbook, "E4"), 11.0);
        assert_eq!(text(&workbook, "D5"), "Grand Total");
        assert_eq!(number(&workbook, "E5"), 33.0);
    }

    #[test]
    fn refreshes_base_field_show_as_calculations() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Period").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", 1.0).unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", 2.0).unwrap();
        sheet.set_cell_value("B3", 15.0).unwrap();
        sheet.set_cell_value("A4", 3.0).unwrap();
        sheet.set_cell_value("B4", 20.0).unwrap();

        let source = CellRange::parse("A1:B4").unwrap();
        let running = PivotTable::builder("Running")
            .source_range(source)
            .target_address("D1")
            .unwrap()
            .row("Period")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                    PivotShowAs::RunningTotal {
                        base_field: "Period".into(),
                    },
                ),
            )
            .build()
            .unwrap();
        let difference = PivotTable::builder("Difference")
            .source_range(source)
            .target_address("G1")
            .unwrap()
            .row("Period")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                    PivotShowAs::DifferenceFrom {
                        base_field: "Period".into(),
                        base_item: PivotValue::Number(1.0),
                    },
                ),
            )
            .build()
            .unwrap();
        let percent_difference = PivotTable::builder("PercentDifference")
            .source_range(source)
            .target_address("J1")
            .unwrap()
            .row("Period")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                    PivotShowAs::PercentDifferenceFrom {
                        base_field: "Period".into(),
                        base_item: PivotValue::Number(1.0),
                    },
                ),
            )
            .build()
            .unwrap();
        let rank = PivotTable::builder("Rank")
            .source_range(source)
            .target_address("M1")
            .unwrap()
            .row("Period")
            .pivot_measure(
                PivotMeasure::new("Revenue", PivotAggregate::Sum).with_show_as(
                    PivotShowAs::RankDescending {
                        base_field: "Period".into(),
                    },
                ),
            )
            .build()
            .unwrap();
        sheet.add_pivot_table(running).unwrap();
        sheet.add_pivot_table(difference).unwrap();
        sheet.add_pivot_table(percent_difference).unwrap();
        sheet.add_pivot_table(rank).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(number(&workbook, "E2"), 10.0);
        assert_eq!(number(&workbook, "E3"), 25.0);
        assert_eq!(number(&workbook, "E4"), 45.0);
        assert_eq!(text(&workbook, "E5"), "");

        assert_eq!(number(&workbook, "H2"), 0.0);
        assert_eq!(number(&workbook, "H3"), 5.0);
        assert_eq!(number(&workbook, "H4"), 10.0);

        assert_eq!(number(&workbook, "K2"), 0.0);
        assert_eq!(number(&workbook, "K3"), 0.5);
        assert_eq!(number(&workbook, "K4"), 1.0);

        assert_eq!(number(&workbook, "N2"), 3.0);
        assert_eq!(number(&workbook, "N3"), 2.0);
        assert_eq!(number(&workbook, "N4"), 1.0);
    }

    #[test]
    fn refresh_applies_item_filters() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", 15.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .filter(PivotFilter::field_items("Region", ["East"]))
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 25.0);
        assert_eq!(text(&workbook, "D3"), "Grand Total");
        assert_eq!(number(&workbook, "E3"), 25.0);
        assert_eq!(text(&workbook, "D4"), "");
    }

    #[test]
    fn refresh_applies_value_filters() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", 15.0).unwrap();
        sheet.set_cell_value("A5", "North").unwrap();
        sheet.set_cell_value("B5", 5.0).unwrap();

        let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .pivot_measure(measure.clone())
            .filter(PivotFilter::Value {
                field: "Region".into(),
                measure,
                operator: PivotFilterOperator::GreaterThanOrEqual,
                value: 20.0,
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 25.0);
        assert_eq!(text(&workbook, "D3"), "West");
        assert_eq!(number(&workbook, "E3"), 20.0);
        assert_eq!(text(&workbook, "D4"), "Grand Total");
        assert_eq!(number(&workbook, "E4"), 45.0);
    }

    #[test]
    fn refresh_applies_top_n_filters() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", 15.0).unwrap();
        sheet.set_cell_value("A5", "North").unwrap();
        sheet.set_cell_value("B5", 5.0).unwrap();

        let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B5").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .pivot_measure(measure.clone())
            .filter(PivotFilter::TopN {
                field: "Region".into(),
                measure,
                n: 2,
                top: true,
                percent: false,
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 25.0);
        assert_eq!(text(&workbook, "D3"), "West");
        assert_eq!(number(&workbook, "E3"), 20.0);
        assert_eq!(text(&workbook, "D4"), "Grand Total");
        assert_eq!(number(&workbook, "E4"), 45.0);
    }

    #[test]
    fn refresh_applies_aggregate_filters_to_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Q1").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Q2").unwrap();
        sheet.set_cell_value("C3", 30.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Q1").unwrap();
        sheet.set_cell_value("C4", 20.0).unwrap();
        sheet.set_cell_value("A5", "West").unwrap();
        sheet.set_cell_value("B5", "Q2").unwrap();
        sheet.set_cell_value("C5", 5.0).unwrap();

        let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C5").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .pivot_measure(measure.clone())
            .filter(PivotFilter::Value {
                field: "Quarter".into(),
                measure,
                operator: PivotFilterOperator::GreaterThan,
                value: 32.0,
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Q2");
        assert_eq!(text(&workbook, "G1"), "Grand Total");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 30.0);
        assert_eq!(number(&workbook, "G2"), 30.0);
        assert_eq!(text(&workbook, "E3"), "West");
        assert_eq!(number(&workbook, "F3"), 5.0);
        assert_eq!(number(&workbook, "G3"), 5.0);
        assert_eq!(text(&workbook, "E4"), "Grand Total");
        assert_eq!(number(&workbook, "F4"), 35.0);
        assert_eq!(number(&workbook, "G4"), 35.0);
    }

    #[test]
    fn refresh_renders_page_fields_above_body() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", "Wholesale").unwrap();
        sheet.set_cell_value("C3", 20.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", 15.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .page("Segment")
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .filter(PivotFilter::field_items("Segment", ["Retail"]))
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Segment");
        assert_eq!(text(&workbook, "F1"), "Retail");
        assert_eq!(text(&workbook, "E3"), "Region");
        assert_eq!(text(&workbook, "E4"), "East");
        assert_eq!(number(&workbook, "F4"), 25.0);
        assert_eq!(text(&workbook, "E5"), "Grand Total");
        assert_eq!(number(&workbook, "F5"), 25.0);
    }

    #[test]
    fn refreshes_table_sources() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();

        let mut table = Table::new(1, "SalesData", CellRange::parse("A1:B3").unwrap());
        table.columns = vec![
            TableColumn::new(1, "Region"),
            TableColumn::new(2, "Revenue"),
        ];
        sheet.add_table(table);

        let pivot = PivotTable::builder("SalesPivot")
            .table_source("SalesData")
            .target_address("D1")
            .unwrap()
            .row("Region")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        let stats = workbook.refresh_pivots().unwrap();

        assert_eq!(stats.source_rows, 2);
        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 10.0);
        assert_eq!(text(&workbook, "D3"), "West");
        assert_eq!(number(&workbook, "E3"), 20.0);
    }

    #[test]
    fn shared_sources_hit_the_internal_snapshot_cache() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();

        let source = CellRange::parse("A1:B2").unwrap();
        let pivot_a = PivotTable::builder("PivotA")
            .source_range(source)
            .target_address("D1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        let pivot_b = PivotTable::builder("PivotB")
            .source_range(source)
            .target_address("G1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot_a).unwrap();
        sheet.add_pivot_table(pivot_b).unwrap();

        let stats = workbook.refresh_pivots().unwrap();

        assert_eq!(stats.cache_misses, 1);
        assert_eq!(stats.cache_hits, 1);
        assert_eq!(number(&workbook, "E2"), 10.0);
        assert_eq!(number(&workbook, "H2"), 10.0);
    }
}
