use crate::aggregate::*;
use crate::api::*;
use crate::compile::*;
use crate::filters::*;
use crate::prelude::*;
use crate::render::*;
use crate::runtime_cache::*;
use crate::snapshot::*;
use crate::transform::*;

#[cfg(feature = "parallel")]
pub(crate) fn with_pivot_refresh_pool<T: Send>(
    options: &PivotRefreshOptions,
    refresh: impl FnOnce() -> T + Send,
) -> T {
    if let Some(max_threads) = options.max_threads {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(max_threads.max(1))
            .build()
            .expect("failed to build pivot refresh rayon thread pool");
        pool.install(refresh)
    } else {
        refresh()
    }
}

#[cfg(not(feature = "parallel"))]
pub(crate) fn with_pivot_refresh_pool<T>(
    _options: &PivotRefreshOptions,
    refresh: impl FnOnce() -> T,
) -> T {
    refresh()
}

#[derive(Debug, Clone)]
pub(crate) struct PivotJob {
    pub(crate) sheet_index: usize,
    pub(crate) pivot_index: usize,
    pub(crate) pivot: PivotTable,
}

#[derive(Debug)]
pub(crate) struct PreparedPivotJob {
    pub(crate) job: PivotJob,
    pub(crate) snapshot: Arc<SourceSnapshot>,
    pub(crate) filter_baselines: PivotFilterBaselines,
    pub(crate) date_system: DateSystem,
    pub(crate) options: PivotRefreshOptions,
}

pub(crate) fn refresh_pivots_inner(
    workbook: &mut Workbook,
    cache: &mut PivotRuntimeCache,
    options: &PivotRefreshOptions,
) -> Result<PivotRefreshStats> {
    let jobs = collect_pivot_jobs(workbook);
    let mut stats = PivotRefreshStats {
        pivot_count: jobs.len(),
        ..PivotRefreshStats::default()
    };

    let date_1904 = workbook.settings().date_1904;
    let mut prepared = Vec::with_capacity(jobs.len());
    for job in jobs {
        if source_requires_external_refresh(&job.pivot.source) {
            mark_pivot_external(workbook, job.sheet_index, job.pivot_index);
            continue;
        }
        let source_snapshot = match cache.snapshot_for_source(
            workbook,
            job.sheet_index,
            &job.pivot.source,
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
        let snapshot = match transformed_snapshot_for_pivot(
            workbook,
            job.sheet_index,
            &job.pivot,
            source_snapshot,
            date_1904,
            cache,
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
        let filter_baselines =
            cache.filter_baselines_for_pivot(job.sheet_index, &job.pivot, &snapshot);
        prepared.push(PreparedPivotJob {
            job,
            snapshot,
            filter_baselines,
            date_system: workbook_date_system(date_1904),
            options: options.clone(),
        });
    }

    let mut rendered = Vec::with_capacity(prepared.len());
    for (job, output) in render_prepared_pivots(prepared) {
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

pub(crate) fn render_prepared_pivots(
    prepared: Vec<PreparedPivotJob>,
) -> Vec<(PivotJob, Result<RenderedPivot>)> {
    #[cfg(feature = "parallel")]
    {
        if prepared.len() > 1 {
            return prepared
                .into_par_iter()
                .map(render_prepared_pivot)
                .collect();
        }
    }

    prepared.into_iter().map(render_prepared_pivot).collect()
}

pub(crate) fn render_prepared_pivot(
    prepared: PreparedPivotJob,
) -> (PivotJob, Result<RenderedPivot>) {
    let PreparedPivotJob {
        job,
        snapshot,
        filter_baselines,
        date_system,
        options,
    } = prepared;
    let output = build_rendered_pivot_from_snapshot(
        &job.pivot,
        snapshot,
        &filter_baselines,
        &options,
        date_system,
    );
    (job, output)
}

pub(crate) fn refresh_pivot_inner(
    workbook: &mut Workbook,
    sheet_index: usize,
    pivot_name: &str,
    cache: &mut PivotRuntimeCache,
    options: &PivotRefreshOptions,
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

    if source_requires_external_refresh(&pivot.source) {
        mark_pivot_external(workbook, sheet_index, pivot_index);
        return Ok(stats);
    }

    let output =
        match build_rendered_pivot(workbook, sheet_index, &pivot, cache, &mut stats, options) {
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

pub(crate) fn pivot_write_touched_ranges(
    job: &PivotJob,
    output_range: CellRange,
) -> Vec<(usize, CellRange)> {
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

pub(crate) fn collect_pivot_jobs(workbook: &Workbook) -> Vec<PivotJob> {
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

pub(crate) fn source_requires_external_refresh(source: &PivotSource) -> bool {
    match source {
        PivotSource::External { .. } | PivotSource::Scenario { .. } | PivotSource::Olap { .. } => {
            true
        }
        PivotSource::Consolidation { ranges } => ranges.iter().any(|range| {
            range.sheet.is_none()
                || range.range.is_none()
                || range.external_relationship_id.is_some()
                || range.external_relationship_target.is_some()
        }),
        PivotSource::WorksheetRange { .. } | PivotSource::Table { .. } => false,
    }
}

pub(crate) fn build_rendered_pivot(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    pivot: &PivotTable,
    cache: &mut PivotRuntimeCache,
    stats: &mut PivotRefreshStats,
    options: &PivotRefreshOptions,
) -> Result<RenderedPivot> {
    let source_snapshot =
        cache.snapshot_for_source(workbook, pivot_sheet_index, &pivot.source, stats)?;
    let snapshot = transformed_snapshot_for_pivot(
        workbook,
        pivot_sheet_index,
        pivot,
        source_snapshot,
        workbook.settings().date_1904,
        cache,
    )?;
    let filter_baselines = cache.filter_baselines_for_pivot(pivot_sheet_index, pivot, &snapshot);
    build_rendered_pivot_from_snapshot(
        pivot,
        snapshot,
        &filter_baselines,
        options,
        workbook_date_system(workbook.settings().date_1904),
    )
}

pub(crate) fn build_rendered_pivot_from_snapshot(
    pivot: &PivotTable,
    snapshot: Arc<SourceSnapshot>,
    filter_baselines: &PivotFilterBaselines,
    options: &PivotRefreshOptions,
    date_system: DateSystem,
) -> Result<RenderedPivot> {
    let plan =
        CompiledPivotPlan::compile(pivot, &snapshot, filter_baselines, options, date_system)?;
    let needs_hidden_total_source = pivot_needs_hidden_total_source(pivot);
    let (mut aggregation, hidden_total_source) = if needs_hidden_total_source {
        let (visible, totals) =
            PivotAggregation::aggregate_visible_with_totals_source(&snapshot, &plan);
        (visible, Some(totals))
    } else {
        (PivotAggregation::aggregate_visible(&snapshot, &plan), None)
    };
    let aggregate_restrictions = aggregation.apply_aggregate_filters(&plan);
    aggregation.apply_calculated_items(&pivot.name, &snapshot, &plan)?;
    aggregation.expand_show_empty_items(pivot, &snapshot, &plan, &aggregate_restrictions)?;
    aggregation.sort_orders(&snapshot, &plan);
    if let Some(mut hidden_total_source) = hidden_total_source {
        hidden_total_source.apply_calculated_items(&pivot.name, &snapshot, &plan)?;
        hidden_total_source.sort_orders(&snapshot, &plan);
        aggregation.attach_hidden_total_source(
            hidden_total_source,
            !pivot.layout.visual_totals,
            !pivot.layout.visual_totals || pivot.layout.subtotal_hidden_items,
        );
    }
    render_pivot(pivot, &snapshot, &plan, &aggregation)
}

pub(crate) fn pivot_needs_hidden_total_source(pivot: &PivotTable) -> bool {
    !pivot.layout.visual_totals || pivot.layout.subtotal_hidden_items
}
