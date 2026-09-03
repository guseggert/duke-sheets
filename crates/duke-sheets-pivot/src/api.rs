use crate::prelude::*;
use crate::refresh::*;
use crate::render::*;
use crate::runtime_cache::*;

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
    pub(crate) fn add_rendered(&mut self, rendered: &RenderedPivot) {
        self.pivots_refreshed += 1;
        self.source_rows += rendered.source_rows;
        self.output_cells += rendered.cell_count();
    }
}

/// Options for refreshing pivot tables.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PivotRefreshOptions {
    /// Maximum number of worker threads to use when the `parallel` feature is enabled.
    ///
    /// `None` uses the active Rayon pool. `Some(1)` confines parallel refresh work
    /// to one worker thread. This option has no effect when the `parallel` feature
    /// is disabled.
    pub max_threads: Option<usize>,
    /// Excel serial date used to evaluate relative date-period filters.
    ///
    /// The serial is interpreted in the workbook's date system. Static month and
    /// quarter period filters do not require this option, but relative filters
    /// such as `today`, `thisMonth`, and `yearToDate` do.
    pub today: Option<f64>,
}

/// Extension methods for refreshing pivot tables in a workbook.
///
/// Refresh-all operations are fail-fast. They capture every local source snapshot
/// before writing any pivot output, so pivot-on-pivot dependencies observe
/// pre-refresh source data.
pub trait WorkbookPivotExt {
    /// Refresh all pivot tables in the workbook.
    fn refresh_pivots(&mut self) -> Result<PivotRefreshStats>;

    /// Refresh all pivot tables in the workbook with custom options.
    fn refresh_pivots_with_options(
        &mut self,
        options: &PivotRefreshOptions,
    ) -> Result<PivotRefreshStats>;

    /// Refresh a single pivot table by worksheet index and pivot name.
    fn refresh_pivot(&mut self, sheet_index: usize, pivot_name: &str) -> Result<PivotRefreshStats>;

    /// Refresh a single pivot table by worksheet index and pivot name with custom options.
    fn refresh_pivot_with_options(
        &mut self,
        sheet_index: usize,
        pivot_name: &str,
        options: &PivotRefreshOptions,
    ) -> Result<PivotRefreshStats>;
}

impl WorkbookPivotExt for Workbook {
    fn refresh_pivots(&mut self) -> Result<PivotRefreshStats> {
        self.refresh_pivots_with_options(&PivotRefreshOptions::default())
    }

    fn refresh_pivots_with_options(
        &mut self,
        options: &PivotRefreshOptions,
    ) -> Result<PivotRefreshStats> {
        let mut cache = PivotRuntimeCache::take_from_workbook(self);
        let result =
            with_pivot_refresh_pool(options, || refresh_pivots_inner(self, &mut cache, options));
        self.set_pivot_runtime_cache(Box::new(cache));
        result
    }

    fn refresh_pivot(&mut self, sheet_index: usize, pivot_name: &str) -> Result<PivotRefreshStats> {
        self.refresh_pivot_with_options(sheet_index, pivot_name, &PivotRefreshOptions::default())
    }

    fn refresh_pivot_with_options(
        &mut self,
        sheet_index: usize,
        pivot_name: &str,
        options: &PivotRefreshOptions,
    ) -> Result<PivotRefreshStats> {
        let mut cache = PivotRuntimeCache::take_from_workbook(self);
        let result = with_pivot_refresh_pool(options, || {
            refresh_pivot_inner(self, sheet_index, pivot_name, &mut cache, options)
        });
        self.set_pivot_runtime_cache(Box::new(cache));
        result
    }
}
