//! Pivot table refresh and file-format planning engine.
//!
//! This crate refreshes semantic [`duke_sheets_core::PivotTable`] definitions
//! into worksheet cells and exposes immutable pivot cache plans to file-format
//! writers.

mod prelude;

mod aggregate;
mod api;
mod compile;
mod filters;
mod refresh;
mod render;
mod runtime_cache;
mod show_as;
mod snapshot;
mod sort;
mod source;
mod transform;

/// Immutable, format-neutral pivot cache plans for file-format writers.
#[path = "format_plan.rs"]
pub mod plan;

pub use api::{PivotRefreshOptions, PivotRefreshStats, WorkbookPivotExt};
pub use plan::{
    pivot_date_period_filter_bounds, pivot_measure_matches_target, plan_format_pivots, shift_month,
    visible_row_indexes, FormatPivotAxisTuples, FormatPivotCache, FormatPivotCacheField,
    FormatPivotGroupLevel, FormatPivotGrouping, FormatPivotPlan, FormatPivotSource,
    FormatPivotTable,
};

#[cfg(test)]
mod test_support;
#[cfg(test)]
mod tests;
