//! Pivot table refresh engine.
//!
//! This crate refreshes semantic [`duke_sheets_core::PivotTable`] definitions
//! into worksheet cells. The file-format pivot cache objects used by XLS/XLSB
//! are deliberately kept out of the public authoring API.

mod prelude;

mod aggregate;
mod api;
mod compile;
mod filters;
#[cfg(any(feature = "format-plan", test))]
mod format_plan;
mod refresh;
mod render;
mod runtime_cache;
mod show_as;
mod snapshot;
mod sort;
mod source;
mod transform;

pub use api::{PivotRefreshOptions, PivotRefreshStats, WorkbookPivotExt};
#[cfg(any(feature = "format-plan", test))]
#[doc(hidden)]
pub use format_plan::{
    plan_format_pivots, FormatPivotAxisTuples, FormatPivotCache, FormatPivotCacheField,
    FormatPivotPlan, FormatPivotSource, FormatPivotTable,
};

#[cfg(test)]
mod test_support;
#[cfg(test)]
mod tests;
