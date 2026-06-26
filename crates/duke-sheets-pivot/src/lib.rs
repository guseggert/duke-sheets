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
    CellAddress, CellError, CellRange, CellValue, Error, NumberFormat, PivotAggregate,
    PivotCalculatedField, PivotField, PivotFilter, PivotFilterOperator, PivotGrouping,
    PivotLayoutKind, PivotManualGroup, PivotMeasure, PivotOverwritePolicy, PivotRefreshStatus,
    PivotShowAs, PivotSort, PivotSource, PivotSubtotal, PivotTable, PivotValue, Result, Table,
    Workbook, Worksheet, MAX_COLS, MAX_ROWS,
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

#[derive(Debug)]
struct PreparedPivotJob {
    job: PivotJob,
    raw_snapshot: Arc<SourceSnapshot>,
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

    let mut prepared = Vec::with_capacity(jobs.len());
    for job in jobs {
        let raw_snapshot = match snapshot_for_source(
            workbook,
            job.sheet_index,
            &job.pivot.source,
            cache,
            &mut stats,
        ) {
            Ok(snapshot) => snapshot,
            Err(error) => {
                mark_pivot_failed(
                    workbook,
                    job.sheet_index,
                    job.pivot_index,
                    error.to_string(),
                );
                return Err(error);
            }
        };
        prepared.push(PreparedPivotJob { job, raw_snapshot });
    }

    let date_1904 = workbook.settings().date_1904;
    let mut rendered = Vec::with_capacity(prepared.len());
    for (job, output) in render_prepared_pivots(prepared, date_1904) {
        match output {
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

    let mut touched_ranges = Vec::new();
    for (job, output) in rendered {
        let output_range = output.range;
        let job_touched_ranges = pivot_write_touched_ranges(&job, output_range);
        if let Err(error) = write_rendered_pivot(workbook, &job, output) {
            mark_pivot_failed(
                workbook,
                job.sheet_index,
                job.pivot_index,
                error.to_string(),
            );
            return Err(error);
        }
        touched_ranges.extend(job_touched_ranges);
    }
    cache.rebase_untouched_sources(workbook, &touched_ranges);

    Ok(stats)
}

fn render_prepared_pivots(
    prepared: Vec<PreparedPivotJob>,
    date_1904: bool,
) -> Vec<(PivotJob, Result<RenderedPivot>)> {
    #[cfg(feature = "parallel")]
    {
        if prepared.len() > 1 {
            return prepared
                .into_par_iter()
                .map(|prepared| render_prepared_pivot(prepared, date_1904))
                .collect();
        }
    }

    prepared
        .into_iter()
        .map(|prepared| render_prepared_pivot(prepared, date_1904))
        .collect()
}

fn render_prepared_pivot(
    prepared: PreparedPivotJob,
    date_1904: bool,
) -> (PivotJob, Result<RenderedPivot>) {
    let PreparedPivotJob { job, raw_snapshot } = prepared;
    let output = build_rendered_pivot_from_snapshot(&job.pivot, raw_snapshot, date_1904);
    (job, output)
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
    let output_range = output.range;
    let job = PivotJob {
        sheet_index,
        pivot_index,
        pivot,
    };
    let touched_ranges = pivot_write_touched_ranges(&job, output_range);
    if let Err(error) = write_rendered_pivot(workbook, &job, output) {
        mark_pivot_failed(workbook, sheet_index, pivot_index, error.to_string());
        return Err(error);
    }
    cache.rebase_untouched_sources(workbook, &touched_ranges);

    Ok(stats)
}

fn pivot_write_touched_ranges(job: &PivotJob, output_range: CellRange) -> Vec<(usize, CellRange)> {
    let mut ranges = Vec::with_capacity(2);
    if matches!(
        job.pivot.overwrite_policy,
        PivotOverwritePolicy::ClearOwnedRange
    ) {
        if let Some(range) = job.pivot.rendered_range {
            ranges.push((job.sheet_index, range));
        }
    }
    ranges.push((job.sheet_index, output_range));
    ranges
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
    build_rendered_pivot_from_snapshot(pivot, raw_snapshot, workbook.settings().date_1904)
}

fn build_rendered_pivot_from_snapshot(
    pivot: &PivotTable,
    raw_snapshot: Arc<SourceSnapshot>,
    date_1904: bool,
) -> Result<RenderedPivot> {
    let calculated_snapshot = if pivot.calculated_fields.is_empty() {
        raw_snapshot
    } else {
        Arc::new(raw_snapshot.apply_calculated_fields(&pivot.name, &pivot.calculated_fields)?)
    };
    let snapshot = if pivot.groupings.is_empty() {
        calculated_snapshot
    } else {
        Arc::new(calculated_snapshot.apply_groupings(&pivot.name, &pivot.groupings, date_1904)?)
    };
    let plan = CompiledPivotPlan::compile(pivot, &snapshot)?;
    let mut aggregation = PivotAggregation::aggregate(&snapshot, &plan);
    let aggregate_restrictions = aggregation.apply_aggregate_filters(&plan);
    aggregation.expand_show_empty_items(&pivot.name, &snapshot, &plan, &aggregate_restrictions)?;
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
                if row_offset >= rendered.data_start_row {
                    if let Some(format) = rendered
                        .column_number_formats
                        .get(col_offset)
                        .and_then(Option::as_deref)
                    {
                        apply_number_format(worksheet, row, col, format)?;
                    }
                }
            }
        }
    }

    if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(job.pivot_index) {
        pivot.rendered_range = Some(rendered.range);
        pivot.refresh_status = PivotRefreshStatus::Succeeded;
        pivot.set_cache_refresh_status(PivotRefreshStatus::Succeeded);
    }

    Ok(())
}

fn apply_number_format(worksheet: &mut Worksheet, row: u32, col: u16, format: &str) -> Result<()> {
    let mut style = worksheet
        .cell_style_at(row, col)
        .cloned()
        .unwrap_or_default();
    style.number_format = NumberFormat::Custom(format.to_string());
    worksheet.set_cell_style_at(row, col, &style)
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
            pivot.set_cache_refresh_status(status);
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

    fn rebase_untouched_sources(
        &mut self,
        workbook: &Workbook,
        touched_ranges: &[(usize, CellRange)],
    ) {
        let mut snapshots = AHashMap::with_capacity(self.snapshots.len());
        for (mut key, snapshot) in std::mem::take(&mut self.snapshots) {
            let Some(worksheet) = workbook.worksheet(key.sheet_index) else {
                continue;
            };
            let source_touched = touched_ranges.iter().any(|(sheet_index, range)| {
                *sheet_index == key.sheet_index && range.overlaps(&key.range)
            });
            if !source_touched {
                key.mutation_count = worksheet.mutation_count();
                key.topology_generation = worksheet.topology_generation();
            } else if worksheet.mutation_count() != key.mutation_count
                || worksheet.topology_generation() != key.topology_generation
            {
                continue;
            }
            snapshots.insert(key, snapshot);
        }
        self.workbook_nonce = workbook.nonce();
        self.structural_generation = workbook.structural_generation();
        self.snapshots = snapshots;
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
        let columns = source_snapshot_columns(worksheet, source, col_count, row_count);

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
        let mut headers = self.headers.clone();
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
            match grouping {
                PivotGrouping::Date { units, .. } if units.len() > 1 => {
                    for unit in units {
                        headers.push(unique_grouped_header(&headers, field_name, *unit));
                        columns.push(grouped_date_column(self, field_index, &[*unit], date_1904));
                    }
                }
                _ => {
                    columns[field_index] =
                        grouped_column(self, field_index, grouping, date_1904, pivot_name)?;
                }
            }
        }

        Ok(Self {
            headers,
            columns,
            row_count: self.row_count,
        })
    }
}

fn source_snapshot_columns(
    worksheet: &Worksheet,
    source: &ResolvedSource,
    col_count: usize,
    row_count: usize,
) -> Vec<EncodedColumn> {
    let Some(data_end_row) = source.data_end_row else {
        return (0..col_count)
            .map(|_| EncodedColumn::with_capacity(row_count))
            .collect();
    };

    let source_cols = (source.range.start.col..=source.range.end.col).collect::<Vec<_>>();
    #[cfg(feature = "parallel")]
    {
        if row_count >= PARALLEL_ROW_THRESHOLD {
            return source_cols
                .into_par_iter()
                .map(|source_col| {
                    source_snapshot_column(
                        worksheet,
                        source.data_start_row,
                        data_end_row,
                        source_col,
                    )
                })
                .collect();
        }
    }

    source_cols
        .into_iter()
        .map(|source_col| {
            source_snapshot_column(worksheet, source.data_start_row, data_end_row, source_col)
        })
        .collect()
}

fn source_snapshot_column(
    worksheet: &Worksheet,
    data_start_row: u32,
    data_end_row: u32,
    source_col: u16,
) -> EncodedColumn {
    let row_count = (data_end_row - data_start_row + 1) as usize;
    let mut column = EncodedColumn::with_capacity(row_count);
    for row in data_start_row..=data_end_row {
        column.push(effective_pivot_value(worksheet, row, source_col));
    }
    column
}

fn unique_grouped_header(
    headers: &[String],
    field_name: &str,
    unit: duke_sheets_core::PivotDateGroupUnit,
) -> String {
    let base = grouped_date_header(field_name, unit);
    if !headers
        .iter()
        .any(|header| header.eq_ignore_ascii_case(&base))
    {
        return base;
    }

    for suffix in 2.. {
        let candidate = format!("{base} {suffix}");
        if !headers
            .iter()
            .any(|header| header.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded grouped header suffix search should return")
}

fn grouped_date_header(field_name: &str, unit: duke_sheets_core::PivotDateGroupUnit) -> String {
    format!("{field_name} ({})", date_group_unit_name(unit))
}

fn date_group_unit_name(unit: duke_sheets_core::PivotDateGroupUnit) -> &'static str {
    use duke_sheets_core::PivotDateGroupUnit;

    match unit {
        PivotDateGroupUnit::Seconds => "Seconds",
        PivotDateGroupUnit::Minutes => "Minutes",
        PivotDateGroupUnit::Hours => "Hours",
        PivotDateGroupUnit::Days => "Days",
        PivotDateGroupUnit::Months => "Months",
        PivotDateGroupUnit::Quarters => "Quarters",
        PivotDateGroupUnit::Years => "Years",
    }
}

fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
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
        PivotGrouping::Manual { groups, .. } => {
            grouped_manual_column(snapshot, field_index, groups, pivot_name)
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

    Ok(remap_grouped_column(snapshot, field_index, |value| {
        group_number_value(value, effective_start, end, interval)
    }))
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
    let date_system = if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    };

    remap_grouped_column(snapshot, field_index, |value| {
        group_date_value(value, units, date_system)
    })
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

fn grouped_manual_column(
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

fn remap_grouped_column<F>(
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

fn manual_group_lookup(
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

fn group_manual_value(value: &PivotValue, lookup: &AHashMap<PivotValue, String>) -> PivotValue {
    lookup
        .get(value)
        .map(|group| PivotValue::String(group.clone()))
        .unwrap_or_else(|| value.clone())
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

    fn remap_dictionary<F>(&self, group_value: F) -> Self
    where
        F: Fn(&PivotValue) -> PivotValue,
    {
        let mut dictionary = Vec::new();
        let mut lookup = AHashMap::new();
        let mut id_map = Vec::with_capacity(self.dictionary.len());

        for value in &self.dictionary {
            let grouped = group_value(value);
            let id = if let Some(id) = lookup.get(&grouped) {
                *id
            } else {
                let id = dictionary.len() as u32;
                dictionary.push(grouped.clone());
                lookup.insert(grouped, id);
                id
            };
            id_map.push(id);
        }

        let values = remap_column_ids(&self.values, &id_map);
        Self {
            values,
            dictionary,
            lookup,
        }
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

fn remap_column_ids(values: &[u32], id_map: &[u32]) -> Vec<u32> {
    #[cfg(feature = "parallel")]
    {
        if values.len() >= PARALLEL_ROW_THRESHOLD {
            return values
                .par_iter()
                .map(|id| id_map[*id as usize])
                .collect::<Vec<_>>();
        }
    }

    values.iter().map(|id| id_map[*id as usize]).collect()
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
            row_fields,
            column_fields,
            page_fields,
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
    groupings: &[PivotGrouping],
) -> Result<(Vec<usize>, Vec<PivotField>)> {
    let mut indexes = Vec::new();
    let mut compiled_fields = Vec::new();
    for field in fields {
        if let Some(units) = multi_unit_date_grouping_units(groupings, &field.field.name) {
            for unit in units {
                let (index, header) =
                    grouped_date_field_index(snapshot, &field.field.name, *unit).ok_or_else(
                        || {
                            Error::other(format!(
                                "pivot table {pivot_name} references unknown grouped {axis_name} field: {}",
                                grouped_date_header(&field.field.name, *unit)
                            ))
                        },
                    )?;
                let mut grouped_field = field.clone();
                grouped_field.field.name = header;
                indexes.push(index);
                compiled_fields.push(grouped_field);
            }
        } else {
            let index = field_index(snapshot, &field.field.name, pivot_name).map_err(|_| {
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

fn grouped_date_field_index(
    snapshot: &SourceSnapshot,
    field_name: &str,
    unit: duke_sheets_core::PivotDateGroupUnit,
) -> Option<(usize, String)> {
    let base = grouped_date_header(field_name, unit);
    snapshot
        .headers
        .iter()
        .enumerate()
        .rev()
        .find(|(_, header)| grouped_header_matches(header, &base))
        .map(|(index, header)| (index, header.clone()))
}

fn grouped_header_matches(header: &str, base: &str) -> bool {
    if header.eq_ignore_ascii_case(base) {
        return true;
    }
    header
        .strip_prefix(base)
        .and_then(|suffix| suffix.strip_prefix(' '))
        .is_some_and(|suffix| suffix.parse::<usize>().is_ok())
}

fn multi_unit_date_grouping_units<'a>(
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

    fn allows_item(&self, snapshot: &SourceSnapshot, field_index: usize, item_id: u32) -> bool {
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
            _ => true,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum AggregateFilterAxis {
    Row,
    Column,
}

#[derive(Debug, Default, Clone)]
struct AxisItemRestrictions {
    rows: AHashMap<usize, AHashSet<u32>>,
    columns: AHashMap<usize, AHashSet<u32>>,
}

impl AxisItemRestrictions {
    fn restrict(
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

    fn allows(&self, axis: AggregateFilterAxis, field_position: usize, item_id: u32) -> bool {
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
    row_subtotals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    row_order: Vec<Vec<u32>>,
    column_totals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    column_subtotals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    subtotal_groups: AHashMap<GroupKey, Vec<AggregateState>>,
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
            row_subtotals: AHashMap::new(),
            row_order: Vec::new(),
            column_totals: AHashMap::new(),
            column_subtotals: AHashMap::new(),
            subtotal_groups: AHashMap::new(),
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
            row_subtotals: AHashMap::new(),
            row_order: Vec::new(),
            column_totals: AHashMap::new(),
            column_subtotals: AHashMap::new(),
            subtotal_groups: AHashMap::new(),
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

        self.ingest_subtotals(snapshot, plan, row, &row_key, &column_key);

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

    fn ingest_subtotals(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        row_key: &[u32],
        column_key: &[u32],
    ) {
        let row_prefixes = row_subtotal_prefixes(plan, row_key);
        let column_prefixes = column_subtotal_prefixes(plan, column_key);

        for prefix in &row_prefixes {
            self.update_row_subtotal(snapshot, plan, row, prefix.clone());
            if !column_key.is_empty() {
                self.update_subtotal_group(
                    snapshot,
                    plan,
                    row,
                    prefix.clone(),
                    column_key.to_vec(),
                );
            }
        }

        for prefix in &column_prefixes {
            self.update_column_subtotal(snapshot, plan, row, prefix.clone());
            self.update_subtotal_group(snapshot, plan, row, row_key.to_vec(), prefix.clone());
        }

        for row_prefix in &row_prefixes {
            for column_prefix in &column_prefixes {
                self.update_subtotal_group(
                    snapshot,
                    plan,
                    row,
                    row_prefix.clone(),
                    column_prefix.clone(),
                );
            }
        }
    }

    fn update_row_subtotal(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        prefix: Vec<u32>,
    ) {
        let states = self
            .row_subtotals
            .entry(prefix)
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    fn update_column_subtotal(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        prefix: Vec<u32>,
    ) {
        let states = self
            .column_subtotals
            .entry(prefix)
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    fn update_subtotal_group(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        row_key: Vec<u32>,
        column_key: Vec<u32>,
    ) {
        let states = self
            .subtotal_groups
            .entry(GroupKey {
                rows: row_key,
                columns: column_key,
            })
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    fn apply_aggregate_filters(&mut self, plan: &CompiledPivotPlan) -> AxisItemRestrictions {
        let mut restrictions = AxisItemRestrictions::default();
        for filter in &plan.aggregate_filters {
            let allowed_item_ids = filter.allowed_item_ids(self);
            restrictions.restrict(filter.axis(), filter.field_position(), &allowed_item_ids);
            self.retain_axis_items(
                filter.axis(),
                filter.field_position(),
                &allowed_item_ids,
                plan,
            );
        }
        restrictions
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
        self.row_subtotals.clear();
        self.column_totals.clear();
        self.column_subtotals.clear();
        self.subtotal_groups.clear();
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

            let row_prefixes = row_subtotal_prefixes(plan, &key.rows);
            let column_prefixes = column_subtotal_prefixes(plan, &key.columns);

            for prefix in &row_prefixes {
                let row_subtotal = self
                    .row_subtotals
                    .entry(prefix.clone())
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(row_subtotal, states);

                if !key.columns.is_empty() {
                    let subtotal_group = self
                        .subtotal_groups
                        .entry(GroupKey {
                            rows: prefix.clone(),
                            columns: key.columns.clone(),
                        })
                        .or_insert_with(|| default_states(&plan.measures));
                    merge_state_slices(subtotal_group, states);
                }
            }

            for prefix in &column_prefixes {
                let column_subtotal = self
                    .column_subtotals
                    .entry(prefix.clone())
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(column_subtotal, states);

                let subtotal_group = self
                    .subtotal_groups
                    .entry(GroupKey {
                        rows: key.rows.clone(),
                        columns: prefix.clone(),
                    })
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(subtotal_group, states);
            }

            for row_prefix in &row_prefixes {
                for column_prefix in &column_prefixes {
                    let subtotal_group = self
                        .subtotal_groups
                        .entry(GroupKey {
                            rows: row_prefix.clone(),
                            columns: column_prefix.clone(),
                        })
                        .or_insert_with(|| default_states(&plan.measures));
                    merge_state_slices(subtotal_group, states);
                }
            }

            merge_state_slices(&mut self.grand_totals, states);
        }

        self.row_order
            .retain(|key| self.row_totals.contains_key(key));
        self.column_order
            .retain(|key| self.column_totals.contains_key(key));
    }

    fn expand_show_empty_items(
        &mut self,
        pivot_name: &str,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregate_restrictions: &AxisItemRestrictions,
    ) -> Result<()> {
        expand_axis_show_empty_items(
            pivot_name,
            snapshot,
            plan,
            AggregateFilterAxis::Row,
            &plan.row_indexes,
            &plan.row_fields,
            &plan.filters,
            aggregate_restrictions,
            &mut self.row_order,
        )?;
        expand_axis_show_empty_items(
            pivot_name,
            snapshot,
            plan,
            AggregateFilterAxis::Column,
            &plan.column_indexes,
            &plan.column_fields,
            &plan.filters,
            aggregate_restrictions,
            &mut self.column_order,
        )?;
        Ok(())
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

        for (key, states) in other.row_subtotals {
            merge_unordered_bucket(&mut self.row_subtotals, key, states);
        }

        for key in other.column_order {
            let states = other
                .column_totals
                .get(&key)
                .expect("ordered column key must exist")
                .clone();
            merge_ordered_bucket(&mut self.column_totals, &mut self.column_order, key, states);
        }

        for (key, states) in other.column_subtotals {
            merge_unordered_bucket(&mut self.column_subtotals, key, states);
        }

        for (key, states) in other.subtotal_groups {
            merge_unordered_bucket(&mut self.subtotal_groups, key, states);
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

#[cfg(feature = "parallel")]
fn merge_unordered_bucket<K>(
    map: &mut AHashMap<K, Vec<AggregateState>>,
    key: K,
    states: Vec<AggregateState>,
) where
    K: Eq + Hash,
{
    if let Some(existing) = map.get_mut(&key) {
        merge_state_slices(existing, &states);
    } else {
        map.insert(key, states);
    }
}

fn merge_state_slices(target: &mut [AggregateState], source: &[AggregateState]) {
    for (target, source) in target.iter_mut().zip(source.iter()) {
        target.merge(source);
    }
}

fn expand_axis_show_empty_items(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    axis: AggregateFilterAxis,
    field_indexes: &[usize],
    fields: &[PivotField],
    filters: &[CompiledFilter],
    aggregate_restrictions: &AxisItemRestrictions,
    order: &mut Vec<Vec<u32>>,
) -> Result<()> {
    if field_indexes.is_empty() || fields.iter().all(|field| !field.show_empty_items) {
        return Ok(());
    }

    let item_ids = axis_item_ids(
        snapshot,
        axis,
        field_indexes,
        fields,
        filters,
        aggregate_restrictions,
        order,
    );
    if item_ids.iter().any(Vec::is_empty) {
        return Ok(());
    }

    let key_count = cartesian_len(&item_ids)?;
    let limit = show_empty_axis_key_limit(axis, plan);
    if key_count > limit {
        return Err(Error::other(format!(
            "pivot table {pivot_name} show-empty-items expansion would produce {key_count} {} keys, exceeding the worksheet limit {limit}",
            axis_name(axis)
        )));
    }

    let mut seen = order.iter().cloned().collect::<AHashSet<_>>();
    let mut key = Vec::with_capacity(field_indexes.len());
    append_show_empty_axis_keys(&item_ids, 0, &mut key, &mut seen, order);
    Ok(())
}

fn axis_item_ids(
    snapshot: &SourceSnapshot,
    axis: AggregateFilterAxis,
    field_indexes: &[usize],
    fields: &[PivotField],
    filters: &[CompiledFilter],
    aggregate_restrictions: &AxisItemRestrictions,
    order: &[Vec<u32>],
) -> Vec<Vec<u32>> {
    field_indexes
        .iter()
        .enumerate()
        .map(|(position, field_index)| {
            let mut ids = if fields
                .get(position)
                .map(|field| field.show_empty_items)
                .unwrap_or(false)
            {
                visible_dictionary_item_ids(snapshot, *field_index, filters)
            } else {
                observed_axis_item_ids(order, position)
            };
            ids.retain(|id| aggregate_restrictions.allows(axis, position, *id));
            ids
        })
        .collect()
}

fn visible_dictionary_item_ids(
    snapshot: &SourceSnapshot,
    field_index: usize,
    filters: &[CompiledFilter],
) -> Vec<u32> {
    (0..snapshot.columns[field_index].dictionary.len())
        .map(|id| id as u32)
        .filter(|id| {
            filters
                .iter()
                .all(|filter| filter.allows_item(snapshot, field_index, *id))
        })
        .collect()
}

fn observed_axis_item_ids(order: &[Vec<u32>], position: usize) -> Vec<u32> {
    let mut seen = AHashSet::new();
    let mut ids = Vec::new();
    for key in order {
        let Some(id) = key.get(position) else {
            continue;
        };
        if seen.insert(*id) {
            ids.push(*id);
        }
    }
    ids
}

fn cartesian_len(item_ids: &[Vec<u32>]) -> Result<usize> {
    item_ids.iter().try_fold(1usize, |total, ids| {
        total
            .checked_mul(ids.len())
            .ok_or_else(|| Error::other("pivot show-empty-items expansion is too large"))
    })
}

fn show_empty_axis_key_limit(axis: AggregateFilterAxis, plan: &CompiledPivotPlan) -> usize {
    match axis {
        AggregateFilterAxis::Row => MAX_ROWS as usize,
        AggregateFilterAxis::Column => {
            let available_columns = (MAX_COLS as usize).saturating_sub(plan.row_indexes.len());
            available_columns / plan.measures.len().max(1)
        }
    }
}

fn append_show_empty_axis_keys(
    item_ids: &[Vec<u32>],
    position: usize,
    key: &mut Vec<u32>,
    seen: &mut AHashSet<Vec<u32>>,
    order: &mut Vec<Vec<u32>>,
) {
    if position == item_ids.len() {
        if seen.insert(key.clone()) {
            order.push(key.clone());
        }
        return;
    }

    for id in &item_ids[position] {
        key.push(*id);
        append_show_empty_axis_keys(item_ids, position + 1, key, seen, order);
        key.pop();
    }
}

fn axis_name(axis: AggregateFilterAxis) -> &'static str {
    match axis {
        AggregateFilterAxis::Row => "row",
        AggregateFilterAxis::Column => "column",
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
    column_number_formats: Vec<Option<String>>,
    data_start_row: usize,
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
    let mut cells = match (
        compact_row_layout(pivot, plan),
        plan.column_indexes.is_empty(),
    ) {
        (true, true) => render_compact_without_column_fields(pivot, snapshot, plan, aggregation),
        (true, false) => render_compact_with_column_fields(pivot, snapshot, plan, aggregation),
        (false, true) => render_without_column_fields(pivot, snapshot, plan, aggregation),
        (false, false) => render_with_column_fields(pivot, snapshot, plan, aggregation),
    };
    prepend_page_fields(&mut cells, pivot, snapshot, plan);

    let width = cells.iter().map(Vec::len).max().unwrap_or(1).max(1);
    for row in &mut cells {
        row.resize(width, CellValue::Empty);
    }
    if cells.is_empty() {
        cells.push(vec![CellValue::Empty; width]);
    }
    let mut column_number_formats = pivot_column_number_formats(pivot, plan, aggregation);
    column_number_formats.resize(width, None);
    let data_start_row = pivot_data_start_row(plan);

    let range = output_range(pivot.target, cells.len(), width)?;
    Ok(RenderedPivot {
        cells,
        range,
        source_rows: snapshot.row_count,
        column_number_formats,
        data_start_row,
    })
}

fn compact_row_layout(pivot: &PivotTable, plan: &CompiledPivotPlan) -> bool {
    matches!(pivot.layout.kind, PivotLayoutKind::Compact) && plan.row_indexes.len() > 1
}

fn pivot_data_start_row(plan: &CompiledPivotPlan) -> usize {
    if plan.page_fields.is_empty() {
        1
    } else {
        plan.page_fields.len() + 2
    }
}

fn pivot_column_number_formats(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<Option<String>> {
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
    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let mut row = row_label_cells(pivot, snapshot, plan, row_key, previous_row_key);
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

        let next_row_key = aggregation.row_order.get(row_index + 1);
        append_row_subtotals_without_column_fields(
            &mut cells,
            snapshot,
            plan,
            aggregation,
            row_key,
            next_row_key,
            &empty_column_key,
        );
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

fn render_compact_without_column_fields(
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
    cells.push(header);

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
            aggregation.row_totals.get(row_key),
            aggregation.column_totals.get(&empty_column_key),
            &aggregation.grand_totals,
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
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = vec![CellValue::string("Grand Total")];
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
    let column_slots = column_render_slots(pivot, plan, aggregation);
    let mut header = plan
        .row_indexes
        .iter()
        .map(|index| CellValue::string(&snapshot.headers[*index]))
        .collect::<Vec<_>>();

    for slot in &column_slots {
        for measure in &plan.measures {
            let caption = match slot {
                ColumnRenderSlot::GrandTotal => {
                    grand_total_measure_caption(measure, plan.measures.len())
                }
                _ => measure_column_caption(
                    &column_slot_label(snapshot, plan, slot),
                    measure,
                    plan.measures.len(),
                ),
            };
            header.push(CellValue::string(caption));
        }
    }
    cells.push(header);

    for (row_index, row_key) in aggregation.row_order.iter().enumerate() {
        let previous_row_key = row_index
            .checked_sub(1)
            .and_then(|index| aggregation.row_order.get(index));
        let mut row = row_label_cells(pivot, snapshot, plan, row_key, previous_row_key);
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
                aggregation.row_totals.get(row_key),
                column_slot_total(aggregation, slot),
                &aggregation.grand_totals,
                &context,
                column_slot_aggregate_override(plan, slot),
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
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = grand_total_label_row(plan.row_indexes.len());
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
                Some(&aggregation.grand_totals),
                column_slot_total(aggregation, slot),
                &aggregation.grand_totals,
                &context,
                column_slot_aggregate_override(plan, slot),
            ));
        }
        cells.push(row);
    }

    cells
}

fn render_compact_with_column_fields(
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
                ColumnRenderSlot::GrandTotal => {
                    grand_total_measure_caption(measure, plan.measures.len())
                }
                _ => measure_column_caption(
                    &column_slot_label(snapshot, plan, slot),
                    measure,
                    plan.measures.len(),
                ),
            };
            header.push(CellValue::string(caption));
        }
    }
    cells.push(header);

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
                aggregation.row_totals.get(row_key),
                column_slot_total(aggregation, slot),
                &aggregation.grand_totals,
                &context,
                column_slot_aggregate_override(plan, slot),
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
        );
    }

    if pivot.layout.show_row_grand_totals {
        let mut row = vec![CellValue::string("Grand Total")];
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
                Some(&aggregation.grand_totals),
                column_slot_total(aggregation, slot),
                &aggregation.grand_totals,
                &context,
                column_slot_aggregate_override(plan, slot),
            ));
        }
        cells.push(row);
    }

    cells
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum ColumnRenderSlot {
    Leaf(Vec<u32>),
    Subtotal(Vec<u32>),
    GrandTotal,
}

fn column_render_slots(
    pivot: &PivotTable,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
) -> Vec<ColumnRenderSlot> {
    let mut slots = Vec::new();
    for (column_index, column_key) in aggregation.column_order.iter().enumerate() {
        slots.push(ColumnRenderSlot::Leaf(column_key.clone()));

        let next_column_key = aggregation.column_order.get(column_index + 1);
        for position in column_subtotal_positions_to_emit(plan, column_key, next_column_key) {
            slots.push(ColumnRenderSlot::Subtotal(column_key[..=position].to_vec()));
        }
    }

    if pivot.layout.show_column_grand_totals {
        slots.push(ColumnRenderSlot::GrandTotal);
    }
    slots
}

fn column_slot_label(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    slot: &ColumnRenderSlot,
) -> String {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => key_label(snapshot, &plan.column_indexes, column_key),
        ColumnRenderSlot::Subtotal(prefix) => {
            subtotal_key_label(snapshot, &plan.column_indexes, prefix)
        }
        ColumnRenderSlot::GrandTotal => "Grand Total".to_string(),
    }
}

fn column_context_key(slot: &ColumnRenderSlot) -> Option<&Vec<u32>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => Some(column_key),
        ColumnRenderSlot::Subtotal(_) | ColumnRenderSlot::GrandTotal => None,
    }
}

fn column_slot_aggregate_override(
    plan: &CompiledPivotPlan,
    slot: &ColumnRenderSlot,
) -> Option<PivotAggregate> {
    match slot {
        ColumnRenderSlot::Subtotal(prefix) => subtotal_aggregate_for_field(
            plan.column_fields
                .get(prefix.len().saturating_sub(1))
                .map(|field| field.subtotal)
                .unwrap_or(PivotSubtotal::Automatic),
        ),
        ColumnRenderSlot::Leaf(_) | ColumnRenderSlot::GrandTotal => None,
    }
}

fn leaf_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    row_key: &[u32],
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.groups.get(&GroupKey {
            rows: row_key.to_vec(),
            columns: column_key.clone(),
        }),
        ColumnRenderSlot::Subtotal(column_prefix) => {
            subtotal_group_states(aggregation, row_key, column_prefix)
        }
        ColumnRenderSlot::GrandTotal => aggregation.row_totals.get(row_key),
    }
}

fn subtotal_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    row_prefix: &[u32],
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => {
            subtotal_group_states(aggregation, row_prefix, column_key)
        }
        ColumnRenderSlot::Subtotal(column_prefix) => {
            subtotal_group_states(aggregation, row_prefix, column_prefix)
        }
        ColumnRenderSlot::GrandTotal => aggregation.row_subtotals.get(row_prefix),
    }
}

fn grand_row_slot_states<'a>(
    aggregation: &'a PivotAggregation,
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.column_totals.get(column_key),
        ColumnRenderSlot::Subtotal(column_prefix) => {
            aggregation.column_subtotals.get(column_prefix)
        }
        ColumnRenderSlot::GrandTotal => Some(&aggregation.grand_totals),
    }
}

fn column_slot_total<'a>(
    aggregation: &'a PivotAggregation,
    slot: &ColumnRenderSlot,
) -> Option<&'a Vec<AggregateState>> {
    match slot {
        ColumnRenderSlot::Leaf(column_key) => aggregation.column_totals.get(column_key),
        ColumnRenderSlot::Subtotal(column_prefix) => {
            aggregation.column_subtotals.get(column_prefix)
        }
        ColumnRenderSlot::GrandTotal => Some(&aggregation.grand_totals),
    }
}

fn subtotal_group_states<'a>(
    aggregation: &'a PivotAggregation,
    row_key: &[u32],
    column_key: &[u32],
) -> Option<&'a Vec<AggregateState>> {
    aggregation.subtotal_groups.get(&GroupKey {
        rows: row_key.to_vec(),
        columns: column_key.to_vec(),
    })
}

fn append_row_subtotals_without_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    empty_column_key: &Vec<u32>,
) {
    for position in row_subtotal_positions_to_emit(plan, row_key, next_row_key) {
        let prefix = row_key[..=position].to_vec();
        let Some(states) = aggregation.row_subtotals.get(&prefix) else {
            continue;
        };

        let mut row = row_subtotal_label_cells(snapshot, &plan.row_indexes, &prefix);
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
                .column_totals
                .get(empty_column_key)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            &aggregation.grand_totals,
            &context,
            row_subtotal_aggregate_override(plan, position),
        ));
        cells.push(row);
    }
}

fn append_row_subtotals_with_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    column_slots: &[ColumnRenderSlot],
) {
    for position in row_subtotal_positions_to_emit(plan, row_key, next_row_key) {
        let prefix = row_key[..=position].to_vec();
        let mut row = row_subtotal_label_cells(snapshot, &plan.row_indexes, &prefix);
        let row_total = aggregation.row_subtotals.get(&prefix);
        let row_aggregate_override = row_subtotal_aggregate_override(plan, position);

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
                &aggregation.grand_totals,
                &context,
                row_aggregate_override.or_else(|| column_slot_aggregate_override(plan, slot)),
            ));
        }

        cells.push(row);
    }
}

fn append_compact_row_subtotals_without_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    empty_column_key: &Vec<u32>,
) {
    for position in row_subtotal_positions_to_emit(plan, row_key, next_row_key) {
        let prefix = row_key[..=position].to_vec();
        let Some(states) = aggregation.row_subtotals.get(&prefix) else {
            continue;
        };

        let mut row = vec![compact_subtotal_label_cell(snapshot, plan, &prefix)];
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
                .column_totals
                .get(empty_column_key)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            &aggregation.grand_totals,
            &context,
            row_subtotal_aggregate_override(plan, position),
        ));
        cells.push(row);
    }
}

fn append_compact_row_subtotals_with_column_fields(
    cells: &mut Vec<Vec<CellValue>>,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
    column_slots: &[ColumnRenderSlot],
) {
    for position in row_subtotal_positions_to_emit(plan, row_key, next_row_key) {
        let prefix = row_key[..=position].to_vec();
        let mut row = vec![compact_subtotal_label_cell(snapshot, plan, &prefix)];
        let row_total = aggregation.row_subtotals.get(&prefix);
        let row_aggregate_override = row_subtotal_aggregate_override(plan, position);

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
                &aggregation.grand_totals,
                &context,
                row_aggregate_override.or_else(|| column_slot_aggregate_override(plan, slot)),
            ));
        }

        cells.push(row);
    }
}

fn row_subtotal_aggregate_override(
    plan: &CompiledPivotPlan,
    position: usize,
) -> Option<PivotAggregate> {
    subtotal_aggregate_for_field(
        plan.row_fields
            .get(position)
            .map(|field| field.subtotal)
            .unwrap_or(PivotSubtotal::Automatic),
    )
}

fn row_subtotal_positions_to_emit(
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    next_row_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    subtotal_positions_to_emit(row_key, next_row_key, |position| {
        row_subtotal_enabled(plan, position)
    })
}

fn column_subtotal_positions_to_emit(
    plan: &CompiledPivotPlan,
    column_key: &[u32],
    next_column_key: Option<&Vec<u32>>,
) -> Vec<usize> {
    subtotal_positions_to_emit(column_key, next_column_key, |position| {
        column_subtotal_enabled(plan, position)
    })
}

fn subtotal_positions_to_emit(
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

fn row_subtotal_prefixes(plan: &CompiledPivotPlan, row_key: &[u32]) -> Vec<Vec<u32>> {
    subtotal_prefixes(row_key, |position| row_subtotal_enabled(plan, position))
}

fn column_subtotal_prefixes(plan: &CompiledPivotPlan, column_key: &[u32]) -> Vec<Vec<u32>> {
    subtotal_prefixes(column_key, |position| {
        column_subtotal_enabled(plan, position)
    })
}

fn subtotal_prefixes(key: &[u32], enabled: impl Fn(usize) -> bool) -> Vec<Vec<u32>> {
    (1..key.len())
        .filter(|prefix_len| enabled(prefix_len - 1))
        .map(|prefix_len| key[..prefix_len].to_vec())
        .collect()
}

fn row_subtotal_enabled(plan: &CompiledPivotPlan, position: usize) -> bool {
    plan.row_fields
        .get(position)
        .map(|field| !matches!(field.subtotal, PivotSubtotal::None))
        .unwrap_or(false)
}

fn column_subtotal_enabled(plan: &CompiledPivotPlan, position: usize) -> bool {
    plan.column_fields
        .get(position)
        .map(|field| !matches!(field.subtotal, PivotSubtotal::None))
        .unwrap_or(false)
}

fn subtotal_aggregate_for_field(subtotal: PivotSubtotal) -> Option<PivotAggregate> {
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

fn row_subtotal_label_cells(
    snapshot: &SourceSnapshot,
    row_indexes: &[usize],
    prefix: &[u32],
) -> Vec<CellValue> {
    let mut row = vec![CellValue::Empty; row_indexes.len()];
    if prefix.is_empty() {
        return row;
    }

    let subtotal_position = prefix.len() - 1;
    for index in 0..subtotal_position {
        row[index] = snapshot
            .value_by_id(row_indexes[index], prefix[index])
            .to_cell_value();
    }
    let value = snapshot
        .value_by_id(row_indexes[subtotal_position], prefix[subtotal_position])
        .to_string();
    row[subtotal_position] = CellValue::string(format!("{value} Total"));
    row
}

fn subtotal_key_label(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    prefix: &[u32],
) -> String {
    let mut labels = field_indexes
        .iter()
        .zip(prefix.iter())
        .map(|(field_index, id)| snapshot.value_by_id(*field_index, *id).to_string())
        .collect::<Vec<_>>();
    if let Some(label) = labels.last_mut() {
        label.push_str(" Total");
    }
    labels.join(" | ")
}

fn same_prefix(left: &[u32], right: &[u32], len: usize) -> bool {
    left.len() >= len && right.len() >= len && left[..len] == right[..len]
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

fn row_label_cells(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
    previous_row_key: Option<&Vec<u32>>,
) -> Vec<CellValue> {
    let mut cells = decode_key_cells(snapshot, &plan.row_indexes, row_key);
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

fn append_compact_group_headers(
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
            &plan.row_indexes,
            row_key,
            position,
        )];
        row.extend(empty_cells(data_width));
        cells.push(row);
    }
}

fn compact_group_header_positions(
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

fn compact_leaf_label_cell(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row_key: &[u32],
) -> CellValue {
    let position = row_key.len().saturating_sub(1);
    key_position_cell(snapshot, &plan.row_indexes, row_key, position)
}

fn compact_subtotal_label_cell(
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    prefix: &[u32],
) -> CellValue {
    let Some(position) = prefix.len().checked_sub(1) else {
        return CellValue::Empty;
    };
    let value = snapshot
        .value_by_id(plan.row_indexes[position], prefix[position])
        .to_string();
    CellValue::string(format!("{value} Total"))
}

fn key_position_cell(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    key: &[u32],
    position: usize,
) -> CellValue {
    field_indexes
        .get(position)
        .zip(key.get(position))
        .map(|(field_index, id)| snapshot.value_by_id(*field_index, *id).to_cell_value())
        .unwrap_or(CellValue::Empty)
}

fn empty_cells(count: usize) -> impl Iterator<Item = CellValue> {
    std::iter::repeat_with(|| CellValue::Empty).take(count)
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

fn finalize_states_with_context_and_aggregate(
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

fn finalize_state_slice_with_context_and_aggregate(
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

fn finalize_measure_with_context(
    state: &AggregateState,
    measure: &PivotMeasure,
    aggregate: PivotAggregate,
    row_total: Option<f64>,
    column_total: Option<f64>,
    grand_total: Option<f64>,
    measure_index: usize,
    context: &ShowAsContext<'_>,
) -> CellValue {
    match &measure.show_as {
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
        PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind, PivotManualGroup,
        PivotMeasure, PivotShowAs, PivotSort, PivotSource, PivotSubtotal, PivotTable, PivotValue,
        Table, TableColumn, Workbook,
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
    fn remapping_grouped_column_coalesces_dictionary_ids() {
        let mut column = super::EncodedColumn::with_capacity(5);
        for value in ["East", "West", "South", "East", "West"] {
            column.push(PivotValue::String(value.to_string()));
        }

        let grouped = column.remap_dictionary(|value| match value {
            PivotValue::String(region) if region == "East" || region == "West" => {
                PivotValue::String("Coastal".to_string())
            }
            value => value.clone(),
        });

        assert_eq!(
            grouped.dictionary,
            vec![
                PivotValue::String("Coastal".to_string()),
                PivotValue::String("South".to_string()),
            ]
        );
        assert_eq!(grouped.values, vec![0, 0, 1, 0, 0]);
        assert_eq!(
            grouped.id_for_value(&PivotValue::String("Coastal".to_string())),
            Some(0)
        );
    }

    fn tabular_layout() -> PivotLayout {
        let mut layout = PivotLayout::default();
        layout.kind = PivotLayoutKind::Tabular;
        layout.repeat_item_labels = true;
        layout
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

    #[cfg(feature = "parallel")]
    #[test]
    fn parallel_refreshes_large_source_snapshot_and_aggregation() {
        let mut workbook = Workbook::new();
        let data_rows = super::PARALLEL_ROW_THRESHOLD + 7;
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();

        for index in 0..data_rows {
            let row = (index + 2) as u32;
            let region = match index % 3 {
                0 => "East",
                1 => "West",
                _ => "North",
            };
            let quarter = if index % 2 == 0 { "Q1" } else { "Q2" };
            sheet.set_cell_value_at(row - 1, 0, region).unwrap();
            sheet.set_cell_value_at(row - 1, 1, quarter).unwrap();
            sheet.set_cell_value_at(row - 1, 2, 1.0).unwrap();
        }

        let source = CellRange::parse(&format!("A1:C{}", data_rows + 1)).unwrap();
        let pivot = PivotTable::builder("LargeSalesPivot")
            .source_range(source)
            .target_address("E1")
            .unwrap()
            .row("Region")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        let stats = workbook.refresh_pivots().unwrap();

        assert_eq!(stats.source_rows, data_rows);
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 8335.0);
        assert_eq!(number(&workbook, "G2"), 8334.0);
        assert_eq!(number(&workbook, "H2"), 16669.0);
        assert_eq!(text(&workbook, "E3"), "North");
        assert_eq!(number(&workbook, "F3"), 8335.0);
        assert_eq!(number(&workbook, "G3"), 8334.0);
        assert_eq!(text(&workbook, "E4"), "West");
        assert_eq!(number(&workbook, "F4"), 8334.0);
        assert_eq!(number(&workbook, "G4"), 8335.0);
        assert_eq!(text(&workbook, "E5"), "Grand Total");
        assert_eq!(number(&workbook, "F5"), 25004.0);
        assert_eq!(number(&workbook, "G5"), 25003.0);
        assert_eq!(number(&workbook, "H5"), data_rows as f64);
    }

    #[test]
    fn refreshes_tabular_layout_without_repeated_item_labels() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", 7.0).unwrap();

        let mut layout = PivotLayout::default();
        layout.kind = PivotLayoutKind::Tabular;
        layout.repeat_item_labels = false;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(layout)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Region");
        assert_eq!(text(&workbook, "F1"), "Segment");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "Online");
        assert_eq!(text(&workbook, "E3"), "");
        assert_eq!(text(&workbook, "F3"), "Retail");
        assert_eq!(text(&workbook, "E4"), "East Total");
        assert_eq!(number(&workbook, "G4"), 15.0);
        assert_eq!(text(&workbook, "E5"), "West");
        assert_eq!(text(&workbook, "F5"), "Retail");
        assert_eq!(number(&workbook, "G5"), 7.0);
    }

    #[test]
    fn refreshes_tabular_layout_with_repeated_item_labels() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();

        let mut layout = PivotLayout::default();
        layout.kind = PivotLayoutKind::Tabular;
        layout.repeat_item_labels = true;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C3").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(layout)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "Online");
        assert_eq!(text(&workbook, "E3"), "East");
        assert_eq!(text(&workbook, "F3"), "Retail");
        assert_eq!(text(&workbook, "E4"), "East Total");
    }

    #[test]
    fn refreshes_compact_layout_hierarchy_without_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", 7.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Row Labels");
        assert_eq!(text(&workbook, "F1"), "Revenue");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "");
        assert_eq!(text(&workbook, "E3"), "Online");
        assert_eq!(number(&workbook, "F3"), 5.0);
        assert_eq!(text(&workbook, "E4"), "Retail");
        assert_eq!(number(&workbook, "F4"), 10.0);
        assert_eq!(text(&workbook, "E5"), "East Total");
        assert_eq!(number(&workbook, "F5"), 15.0);
        assert_eq!(text(&workbook, "E6"), "West");
        assert_eq!(text(&workbook, "F6"), "");
        assert_eq!(text(&workbook, "E7"), "Retail");
        assert_eq!(number(&workbook, "F7"), 7.0);
        assert_eq!(text(&workbook, "E8"), "West Total");
        assert_eq!(number(&workbook, "F8"), 7.0);
        assert_eq!(text(&workbook, "E9"), "Grand Total");
        assert_eq!(number(&workbook, "F9"), 22.0);
    }

    #[test]
    fn refreshes_compact_layout_hierarchy_with_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Quarter").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", "Q1").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", "Q2").unwrap();
        sheet.set_cell_value("D3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", "Q1").unwrap();
        sheet.set_cell_value("D4", 7.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D4").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Row Labels");
        assert_eq!(text(&workbook, "G1"), "Q1");
        assert_eq!(text(&workbook, "H1"), "Q2");
        assert_eq!(text(&workbook, "I1"), "Grand Total");
        assert_eq!(text(&workbook, "F2"), "East");
        assert_eq!(text(&workbook, "G2"), "");
        assert_eq!(text(&workbook, "H2"), "");
        assert_eq!(text(&workbook, "I2"), "");
        assert_eq!(text(&workbook, "F3"), "Online");
        assert_eq!(text(&workbook, "G3"), "");
        assert_eq!(number(&workbook, "H3"), 5.0);
        assert_eq!(number(&workbook, "I3"), 5.0);
        assert_eq!(text(&workbook, "F4"), "Retail");
        assert_eq!(number(&workbook, "G4"), 10.0);
        assert_eq!(text(&workbook, "H4"), "");
        assert_eq!(number(&workbook, "I4"), 10.0);
        assert_eq!(text(&workbook, "F5"), "East Total");
        assert_eq!(number(&workbook, "G5"), 10.0);
        assert_eq!(number(&workbook, "H5"), 5.0);
        assert_eq!(number(&workbook, "I5"), 15.0);
        assert_eq!(text(&workbook, "F6"), "West");
        assert_eq!(text(&workbook, "G6"), "");
        assert_eq!(text(&workbook, "H6"), "");
        assert_eq!(text(&workbook, "I6"), "");
        assert_eq!(text(&workbook, "F7"), "Retail");
        assert_eq!(number(&workbook, "G7"), 7.0);
        assert_eq!(text(&workbook, "H7"), "");
        assert_eq!(number(&workbook, "I7"), 7.0);
        assert_eq!(text(&workbook, "F8"), "West Total");
        assert_eq!(number(&workbook, "G8"), 7.0);
        assert_eq!(text(&workbook, "H8"), "");
        assert_eq!(number(&workbook, "I8"), 7.0);
        assert_eq!(text(&workbook, "F9"), "Grand Total");
        assert_eq!(number(&workbook, "G9"), 17.0);
        assert_eq!(number(&workbook, "H9"), 5.0);
        assert_eq!(number(&workbook, "I9"), 22.0);
    }

    #[test]
    fn refreshes_show_empty_items_on_row_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 7.0).unwrap();
        sheet.set_cell_value("A4", "North").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", 3.0).unwrap();

        let mut region = PivotField::new("Region");
        region.show_empty_items = true;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row(region)
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .filter(PivotFilter::field_items("Segment", ["Retail"]))
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E1"), "Region");
        assert_eq!(text(&workbook, "F1"), "Revenue");
        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(number(&workbook, "F2"), 10.0);
        assert_eq!(text(&workbook, "E3"), "North");
        assert_eq!(number(&workbook, "F3"), 3.0);
        assert_eq!(text(&workbook, "E4"), "West");
        assert_eq!(text(&workbook, "F4"), "");
        assert_eq!(text(&workbook, "E5"), "Grand Total");
        assert_eq!(number(&workbook, "F5"), 13.0);
    }

    #[test]
    fn refreshes_show_empty_items_on_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Quarter").unwrap();
        sheet.set_cell_value("C1", "Segment").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Q1").unwrap();
        sheet.set_cell_value("C2", "Retail").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Q2").unwrap();
        sheet.set_cell_value("C3", "Online").unwrap();
        sheet.set_cell_value("D3", 5.0).unwrap();

        let mut quarter = PivotField::new("Quarter");
        quarter.show_empty_items = true;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D3").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .column(quarter)
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .filter(PivotFilter::field_items("Segment", ["Retail"]))
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Region");
        assert_eq!(text(&workbook, "G1"), "Q1");
        assert_eq!(text(&workbook, "H1"), "Q2");
        assert_eq!(text(&workbook, "I1"), "Grand Total");
        assert_eq!(text(&workbook, "F2"), "East");
        assert_eq!(number(&workbook, "G2"), 10.0);
        assert_eq!(text(&workbook, "H2"), "");
        assert_eq!(number(&workbook, "I2"), 10.0);
        assert_eq!(text(&workbook, "F3"), "Grand Total");
        assert_eq!(number(&workbook, "G3"), 10.0);
        assert_eq!(text(&workbook, "H3"), "");
        assert_eq!(number(&workbook, "I3"), 10.0);
    }

    #[test]
    fn refreshes_row_field_subtotals() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", 7.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C4").unwrap())
            .target_address("E1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(tabular_layout())
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "Online");
        assert_eq!(number(&workbook, "G2"), 5.0);
        assert_eq!(text(&workbook, "E3"), "East");
        assert_eq!(text(&workbook, "F3"), "Retail");
        assert_eq!(number(&workbook, "G3"), 10.0);
        assert_eq!(text(&workbook, "E4"), "East Total");
        assert_eq!(text(&workbook, "F4"), "");
        assert_eq!(number(&workbook, "G4"), 15.0);
        assert_eq!(text(&workbook, "E5"), "West");
        assert_eq!(text(&workbook, "F5"), "Retail");
        assert_eq!(number(&workbook, "G5"), 7.0);
        assert_eq!(text(&workbook, "E6"), "West Total");
        assert_eq!(number(&workbook, "G6"), 7.0);
        assert_eq!(text(&workbook, "E7"), "Grand Total");
        assert_eq!(number(&workbook, "G7"), 22.0);
    }

    #[test]
    fn refreshes_custom_row_subtotal_function() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 20.0).unwrap();

        let mut region = PivotField::new("Region");
        region.subtotal = PivotSubtotal::Average;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C3").unwrap())
            .target_address("E1")
            .unwrap()
            .row(region)
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(tabular_layout())
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "Online");
        assert_eq!(number(&workbook, "G2"), 20.0);
        assert_eq!(text(&workbook, "E3"), "East");
        assert_eq!(text(&workbook, "F3"), "Retail");
        assert_eq!(number(&workbook, "G3"), 10.0);
        assert_eq!(text(&workbook, "E4"), "East Total");
        assert_eq!(number(&workbook, "G4"), 15.0);
        assert_eq!(text(&workbook, "E5"), "Grand Total");
        assert_eq!(number(&workbook, "G5"), 30.0);
    }

    #[test]
    fn refreshes_row_field_subtotals_with_column_fields() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Quarter").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", "Q1").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", "Q2").unwrap();
        sheet.set_cell_value("D3", 5.0).unwrap();
        sheet.set_cell_value("A4", "West").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", "Q1").unwrap();
        sheet.set_cell_value("D4", 7.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D4").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(tabular_layout())
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Region");
        assert_eq!(text(&workbook, "G1"), "Segment");
        assert_eq!(text(&workbook, "H1"), "Q1");
        assert_eq!(text(&workbook, "I1"), "Q2");
        assert_eq!(text(&workbook, "J1"), "Grand Total");
        assert_eq!(text(&workbook, "F4"), "East Total");
        assert_eq!(number(&workbook, "H4"), 10.0);
        assert_eq!(number(&workbook, "I4"), 5.0);
        assert_eq!(number(&workbook, "J4"), 15.0);
        assert_eq!(text(&workbook, "F6"), "West Total");
        assert_eq!(number(&workbook, "H6"), 7.0);
        assert_eq!(text(&workbook, "I6"), "");
        assert_eq!(number(&workbook, "J6"), 7.0);
        assert_eq!(text(&workbook, "F7"), "Grand Total");
        assert_eq!(number(&workbook, "H7"), 17.0);
        assert_eq!(number(&workbook, "I7"), 5.0);
        assert_eq!(number(&workbook, "J7"), 22.0);
    }

    #[test]
    fn refreshes_column_field_subtotals() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Year").unwrap();
        sheet.set_cell_value("C1", "Quarter").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "2024").unwrap();
        sheet.set_cell_value("C2", "Q1").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "2024").unwrap();
        sheet.set_cell_value("C3", "Q2").unwrap();
        sheet.set_cell_value("D3", 5.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", "2025").unwrap();
        sheet.set_cell_value("C4", "Q1").unwrap();
        sheet.set_cell_value("D4", 7.0).unwrap();
        sheet.set_cell_value("A5", "West").unwrap();
        sheet.set_cell_value("B5", "2024").unwrap();
        sheet.set_cell_value("C5", "Q1").unwrap();
        sheet.set_cell_value("D5", 3.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D5").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .column("Year")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Region");
        assert_eq!(text(&workbook, "G1"), "2024 | Q1");
        assert_eq!(text(&workbook, "H1"), "2024 | Q2");
        assert_eq!(text(&workbook, "I1"), "2024 Total");
        assert_eq!(text(&workbook, "J1"), "2025 | Q1");
        assert_eq!(text(&workbook, "K1"), "2025 Total");
        assert_eq!(text(&workbook, "L1"), "Grand Total");
        assert_eq!(text(&workbook, "F2"), "East");
        assert_eq!(number(&workbook, "G2"), 10.0);
        assert_eq!(number(&workbook, "H2"), 5.0);
        assert_eq!(number(&workbook, "I2"), 15.0);
        assert_eq!(number(&workbook, "J2"), 7.0);
        assert_eq!(number(&workbook, "K2"), 7.0);
        assert_eq!(number(&workbook, "L2"), 22.0);
        assert_eq!(text(&workbook, "F3"), "West");
        assert_eq!(number(&workbook, "G3"), 3.0);
        assert_eq!(number(&workbook, "I3"), 3.0);
        assert_eq!(text(&workbook, "J3"), "");
        assert_eq!(text(&workbook, "K3"), "");
        assert_eq!(number(&workbook, "L3"), 3.0);
        assert_eq!(text(&workbook, "F4"), "Grand Total");
        assert_eq!(number(&workbook, "G4"), 13.0);
        assert_eq!(number(&workbook, "H4"), 5.0);
        assert_eq!(number(&workbook, "I4"), 18.0);
        assert_eq!(number(&workbook, "J4"), 7.0);
        assert_eq!(number(&workbook, "K4"), 7.0);
        assert_eq!(number(&workbook, "L4"), 25.0);
    }

    #[test]
    fn refreshes_custom_column_subtotal_function() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Year").unwrap();
        sheet.set_cell_value("C1", "Quarter").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "2024").unwrap();
        sheet.set_cell_value("C2", "Q1").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "2024").unwrap();
        sheet.set_cell_value("C3", "Q2").unwrap();
        sheet.set_cell_value("D3", 20.0).unwrap();

        let mut year = PivotField::new("Year");
        year.subtotal = PivotSubtotal::Average;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D3").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .column(year)
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Region");
        assert_eq!(text(&workbook, "G1"), "2024 | Q1");
        assert_eq!(text(&workbook, "H1"), "2024 | Q2");
        assert_eq!(text(&workbook, "I1"), "2024 Total");
        assert_eq!(text(&workbook, "J1"), "Grand Total");
        assert_eq!(text(&workbook, "F2"), "East");
        assert_eq!(number(&workbook, "G2"), 10.0);
        assert_eq!(number(&workbook, "H2"), 20.0);
        assert_eq!(number(&workbook, "I2"), 15.0);
        assert_eq!(number(&workbook, "J2"), 30.0);
        assert_eq!(text(&workbook, "F3"), "Grand Total");
        assert_eq!(number(&workbook, "I3"), 15.0);
        assert_eq!(number(&workbook, "J3"), 30.0);
    }

    #[test]
    fn refreshes_row_and_column_subtotal_intersections() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Year").unwrap();
        sheet.set_cell_value("D1", "Quarter").unwrap();
        sheet.set_cell_value("E1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", "2024").unwrap();
        sheet.set_cell_value("D2", "Q1").unwrap();
        sheet.set_cell_value("E2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", "2024").unwrap();
        sheet.set_cell_value("D3", "Q2").unwrap();
        sheet.set_cell_value("E3", 5.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", "Retail").unwrap();
        sheet.set_cell_value("C4", "2025").unwrap();
        sheet.set_cell_value("D4", "Q1").unwrap();
        sheet.set_cell_value("E4", 7.0).unwrap();
        sheet.set_cell_value("A5", "West").unwrap();
        sheet.set_cell_value("B5", "Retail").unwrap();
        sheet.set_cell_value("C5", "2024").unwrap();
        sheet.set_cell_value("D5", "Q1").unwrap();
        sheet.set_cell_value("E5", 3.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:E5").unwrap())
            .target_address("G1")
            .unwrap()
            .row("Region")
            .row("Segment")
            .column("Year")
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(tabular_layout())
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "G1"), "Region");
        assert_eq!(text(&workbook, "H1"), "Segment");
        assert_eq!(text(&workbook, "I1"), "2024 | Q1");
        assert_eq!(text(&workbook, "J1"), "2024 | Q2");
        assert_eq!(text(&workbook, "K1"), "2024 Total");
        assert_eq!(text(&workbook, "L1"), "2025 | Q1");
        assert_eq!(text(&workbook, "M1"), "2025 Total");
        assert_eq!(text(&workbook, "N1"), "Grand Total");
        assert_eq!(text(&workbook, "G4"), "East Total");
        assert_eq!(number(&workbook, "I4"), 10.0);
        assert_eq!(number(&workbook, "J4"), 5.0);
        assert_eq!(number(&workbook, "K4"), 15.0);
        assert_eq!(number(&workbook, "L4"), 7.0);
        assert_eq!(number(&workbook, "M4"), 7.0);
        assert_eq!(number(&workbook, "N4"), 22.0);
        assert_eq!(text(&workbook, "G7"), "Grand Total");
        assert_eq!(number(&workbook, "K7"), 18.0);
        assert_eq!(number(&workbook, "M7"), 7.0);
        assert_eq!(number(&workbook, "N7"), 25.0);
    }

    #[test]
    fn refresh_respects_disabled_row_field_subtotals() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Segment").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "Retail").unwrap();
        sheet.set_cell_value("C2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "Online").unwrap();
        sheet.set_cell_value("C3", 5.0).unwrap();

        let mut region = PivotField::new("Region");
        region.subtotal = PivotSubtotal::None;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:C3").unwrap())
            .target_address("E1")
            .unwrap()
            .row(region)
            .row("Segment")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .layout(tabular_layout())
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "E2"), "East");
        assert_eq!(text(&workbook, "F2"), "Online");
        assert_eq!(text(&workbook, "E3"), "East");
        assert_eq!(text(&workbook, "F3"), "Retail");
        assert_eq!(text(&workbook, "E4"), "Grand Total");
        assert_eq!(number(&workbook, "G4"), 15.0);
    }

    #[test]
    fn refresh_respects_disabled_column_field_subtotals() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Year").unwrap();
        sheet.set_cell_value("C1", "Quarter").unwrap();
        sheet.set_cell_value("D1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", "2024").unwrap();
        sheet.set_cell_value("C2", "Q1").unwrap();
        sheet.set_cell_value("D2", 10.0).unwrap();
        sheet.set_cell_value("A3", "East").unwrap();
        sheet.set_cell_value("B3", "2024").unwrap();
        sheet.set_cell_value("C3", "Q2").unwrap();
        sheet.set_cell_value("D3", 5.0).unwrap();
        sheet.set_cell_value("A4", "East").unwrap();
        sheet.set_cell_value("B4", "2025").unwrap();
        sheet.set_cell_value("C4", "Q1").unwrap();
        sheet.set_cell_value("D4", 7.0).unwrap();

        let mut year = PivotField::new("Year");
        year.subtotal = PivotSubtotal::None;
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:D4").unwrap())
            .target_address("F1")
            .unwrap()
            .row("Region")
            .column(year)
            .column("Quarter")
            .named_measure("Revenue", PivotAggregate::Sum, "Revenue")
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "F1"), "Region");
        assert_eq!(text(&workbook, "G1"), "2024 | Q1");
        assert_eq!(text(&workbook, "H1"), "2024 | Q2");
        assert_eq!(text(&workbook, "I1"), "2025 | Q1");
        assert_eq!(text(&workbook, "J1"), "Grand Total");
        assert_eq!(number(&workbook, "J2"), 22.0);
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
    fn refresh_applies_measure_number_format() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Rate").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 0.25).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 0.5).unwrap();

        let pivot = PivotTable::builder("RatePivot")
            .source_range(CellRange::parse("A1:B3").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .pivot_measure(
                PivotMeasure::new("Rate", PivotAggregate::Sum).with_number_format("0.0%"),
            )
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(sheet.formatted_value("E2").unwrap(), "25.0%");
        assert_eq!(sheet.formatted_value("E3").unwrap(), "50.0%");
        assert_eq!(sheet.formatted_value("E4").unwrap(), "75.0%");
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
    fn refreshes_manual_item_grouping() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
        sheet.set_cell_value("A4", "North").unwrap();
        sheet.set_cell_value("B4", 7.0).unwrap();
        sheet.set_cell_value("A5", "South").unwrap();
        sheet.set_cell_value("B5", 8.0).unwrap();
        sheet.set_cell_value("A6", "Central").unwrap();
        sheet.set_cell_value("B6", 5.0).unwrap();
        sheet.set_cell_value("A7", "East").unwrap();
        sheet.set_cell_value("B7", 3.0).unwrap();

        let pivot = PivotTable::builder("GroupedRegions")
            .source_range(CellRange::parse("A1:B7").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .grouping(PivotGrouping::Manual {
                field: "Region".into(),
                groups: vec![
                    PivotManualGroup::new("Coastal", ["East", "West"]),
                    PivotManualGroup::new("Inland", ["North", "South"]),
                ],
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "Central");
        assert_eq!(number(&workbook, "E2"), 5.0);
        assert_eq!(text(&workbook, "D3"), "Coastal");
        assert_eq!(number(&workbook, "E3"), 33.0);
        assert_eq!(text(&workbook, "D4"), "Inland");
        assert_eq!(number(&workbook, "E4"), 15.0);
        assert_eq!(text(&workbook, "D5"), "Grand Total");
        assert_eq!(number(&workbook, "E5"), 53.0);
    }

    #[test]
    fn refreshes_multi_unit_date_grouping_hierarchy() {
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

        assert_eq!(text(&workbook, "D1"), "Row Labels");
        assert_eq!(text(&workbook, "E1"), "Sum of Revenue");
        assert_eq!(text(&workbook, "D2"), "2024");
        assert_eq!(text(&workbook, "E2"), "");
        assert_eq!(text(&workbook, "D3"), "1");
        assert_eq!(number(&workbook, "E3"), 15.0);
        assert_eq!(text(&workbook, "D4"), "2");
        assert_eq!(number(&workbook, "E4"), 7.0);
        assert_eq!(text(&workbook, "D5"), "2024 Total");
        assert_eq!(number(&workbook, "E5"), 22.0);
        assert_eq!(text(&workbook, "D6"), "2025");
        assert_eq!(text(&workbook, "D7"), "1");
        assert_eq!(number(&workbook, "E7"), 11.0);
        assert_eq!(text(&workbook, "D8"), "2025 Total");
        assert_eq!(number(&workbook, "E8"), 11.0);
        assert_eq!(text(&workbook, "D9"), "Grand Total");
        assert_eq!(number(&workbook, "E9"), 33.0);
    }

    #[test]
    fn refreshes_single_unit_date_grouping() {
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

        let pivot = PivotTable::builder("GroupedDates")
            .source_range(CellRange::parse("A1:B4").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Date")
            .measure("Revenue", PivotAggregate::Sum)
            .grouping(PivotGrouping::Date {
                field: "Date".into(),
                units: vec![PivotDateGroupUnit::Months],
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "1");
        assert_eq!(number(&workbook, "E2"), 15.0);
        assert_eq!(text(&workbook, "D3"), "2");
        assert_eq!(number(&workbook, "E3"), 7.0);
        assert_eq!(text(&workbook, "D4"), "Grand Total");
        assert_eq!(number(&workbook, "E4"), 22.0);
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
    fn show_empty_items_respects_value_filters() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 1.0).unwrap();

        let mut region = PivotField::new("Region");
        region.show_empty_items = true;
        let measure = PivotMeasure::new("Revenue", PivotAggregate::Sum);
        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B3").unwrap())
            .target_address("D1")
            .unwrap()
            .row(region)
            .pivot_measure(measure.clone())
            .filter(PivotFilter::Value {
                field: "Region".into(),
                measure,
                operator: PivotFilterOperator::GreaterThanOrEqual,
                value: 5.0,
            })
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        workbook.refresh_pivots().unwrap();

        assert_eq!(text(&workbook, "D2"), "East");
        assert_eq!(number(&workbook, "E2"), 10.0);
        assert_eq!(text(&workbook, "D3"), "Grand Total");
        assert_eq!(number(&workbook, "E3"), 10.0);
        assert_eq!(text(&workbook, "D4"), "");
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

    #[test]
    fn internal_snapshot_cache_reuses_then_invalidates_on_source_mutation() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();

        let pivot = PivotTable::builder("SalesPivot")
            .source_range(CellRange::parse("A1:B2").unwrap())
            .target_address("D1")
            .unwrap()
            .row("Region")
            .measure("Revenue", PivotAggregate::Sum)
            .build()
            .unwrap();
        sheet.add_pivot_table(pivot).unwrap();

        let first = workbook.refresh_pivots().unwrap();
        assert_eq!(first.cache_misses, 1);
        assert_eq!(first.cache_hits, 0);
        assert_eq!(number(&workbook, "E2"), 10.0);

        let second = workbook.refresh_pivots().unwrap();
        assert_eq!(second.cache_misses, 0);
        assert_eq!(second.cache_hits, 1);
        assert_eq!(number(&workbook, "E2"), 10.0);

        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value("B2", 15.0)
            .unwrap();

        let third = workbook.refresh_pivots().unwrap();
        assert_eq!(third.cache_misses, 1);
        assert_eq!(third.cache_hits, 0);
        assert_eq!(number(&workbook, "E2"), 15.0);
    }
}
