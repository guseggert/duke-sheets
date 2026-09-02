use crate::aggregate::*;
use crate::api::*;
use crate::prelude::*;
use crate::refresh::*;
use crate::runtime_cache::*;
use crate::snapshot::*;
use crate::sort::*;
use crate::source::*;
use crate::transform::*;

/// Workbook-level pivot cache plan used by file-format writers.
///
/// This is not an authoring API. It exposes an immutable, format-neutral view
/// of resolved pivot cache data while keeping the mutable runtime caches inside
/// this crate.
#[derive(Debug, Clone)]
pub struct FormatPivotPlan {
    /// Planned cache parts, numbered from 1 in package/write order.
    pub caches: Vec<FormatPivotCache>,
    /// Planned pivot table parts, numbered from 1 in package/write order.
    pub tables: Vec<FormatPivotTable>,
}

/// A resolved pivot cache for file-format writers.
#[derive(Debug, Clone)]
pub struct FormatPivotCache {
    /// One-based cache number used for part names.
    pub cache_num: usize,
    /// Source descriptor for this cache.
    pub source: FormatPivotSource,
    /// Cache fields in source/cache order.
    pub fields: Vec<FormatPivotCacheField>,
    /// Semantic field-name aliases resolved by file-format writers.
    ///
    /// Excel stores consolidation caches with generated field names such as
    /// `Row`, `Column`, and `Value`, while callers may author against source
    /// headers. This mapping keeps that translation out of the semantic pivot
    /// model.
    pub field_aliases: Vec<(String, String)>,
    /// Calculated items registered in this transformed cache.
    pub calculated_items: Vec<PivotCalculatedItem>,
    /// Number of source records.
    pub row_count: usize,
    /// Whether cache records should be written.
    pub save_data: bool,
    /// Refresh-on-open flag from the semantic pivot table.
    pub refresh_on_load: bool,
    /// Background refresh flag from the semantic pivot table.
    pub background_query: bool,
    /// Optional missing-items limit.
    pub missing_items_limit: Option<u32>,
}

impl FormatPivotCache {
    /// Find a field index by case-insensitive cache field name.
    pub fn field_index(&self, name: &str) -> Option<usize> {
        self.fields
            .iter()
            .position(|field| field.name.eq_ignore_ascii_case(name))
            .or_else(|| {
                self.field_aliases
                    .iter()
                    .find(|(alias, _)| alias.eq_ignore_ascii_case(name))
                    .and_then(|(_, target)| {
                        self.fields
                            .iter()
                            .position(|field| field.name.eq_ignore_ascii_case(target))
                    })
            })
    }
}

/// Source descriptor for a planned pivot cache.
#[derive(Debug, Clone)]
pub enum FormatPivotSource {
    /// Worksheet range or table source.
    Worksheet {
        /// Zero-based sheet index containing the source.
        sheet_index: usize,
        /// Source sheet name.
        sheet_name: String,
        /// Source range including the header row.
        range: CellRange,
        /// Table name when the source is a table/list object.
        table_name: Option<String>,
    },
    /// Consolidation source ranges.
    Consolidation {
        /// Original consolidation range descriptors from the semantic pivot.
        ranges: Vec<duke_sheets_core::PivotSourceRange>,
    },
    /// External workbook/database source preserved without local cache records.
    External {
        /// Workbook data connection name or numeric id.
        connection_name: String,
        /// Optional command text associated with the connection.
        command_text: Option<String>,
    },
    /// Scenario source preserved without local cache records.
    Scenario {
        /// Scenario name, when known.
        name: String,
    },
    /// OLAP source preserved without local cache records.
    Olap {
        /// Workbook data connection name or numeric id.
        connection_name: String,
        /// Cube name, when known.
        cube: Option<String>,
        /// Command text, when known.
        command_text: Option<String>,
    },
}

/// A resolved cache field for file-format writers.
#[derive(Debug, Clone)]
pub struct FormatPivotCacheField {
    /// Cache field name.
    pub name: String,
    /// Formula when this field is a calculated field.
    pub formula: Option<String>,
    /// Whether the field maps to a source database field.
    pub database_field: bool,
    /// Shared item dictionary.
    pub shared_items: Vec<PivotValue>,
    /// Field-major item IDs for each source row.
    pub item_ids: Vec<u32>,
    /// Format-neutral grouping data when this is a grouped base field.
    pub grouping: Option<FormatPivotGrouping>,
}

/// Resolved grouping data attached to a planned cache field.
#[derive(Debug, Clone)]
pub struct FormatPivotGrouping {
    /// Semantic grouping definition supplied by the pivot table.
    pub definition: PivotGrouping,
    /// Zero-based cache-field index of the ungrouped base field.
    pub base_field_index: usize,
    /// Zero-based parent field index, when the grouping is nested below one.
    pub parent_field_index: Option<usize>,
    /// Ungrouped source dictionary in first-seen order.
    pub source_items: Vec<PivotValue>,
    /// Ungrouped source item IDs for every cache row.
    pub source_item_ids: Vec<u32>,
    /// One resolved level per numeric/manual grouping or date unit.
    pub levels: Vec<FormatPivotGroupLevel>,
}

/// A resolved grouping level and its source-to-group mapping.
#[derive(Debug, Clone)]
pub struct FormatPivotGroupLevel {
    /// Cache-field index carrying this level's transformed values.
    pub field_index: usize,
    /// Parent cache-field index for hierarchical date levels.
    pub parent_field_index: Option<usize>,
    /// Date unit represented by this level, or `None` for numeric/manual groups.
    pub date_unit: Option<duke_sheets_core::PivotDateGroupUnit>,
    /// Group item dictionary in planned tuple item-ID space.
    pub group_items: Vec<PivotValue>,
    /// Group item ID for each ungrouped source dictionary item.
    pub source_item_group_ids: Vec<u32>,
    /// Group item IDs for every cache row.
    pub item_ids: Vec<u32>,
}

/// A planned pivot table part for file-format writers.
#[derive(Debug, Clone)]
pub struct FormatPivotTable {
    /// Zero-based worksheet index containing the pivot table.
    pub sheet_index: usize,
    /// Zero-based pivot index on the worksheet.
    pub pivot_index: usize,
    /// One-based pivot table number used for part names.
    pub table_num: usize,
    /// One-based cache number used by this pivot.
    pub cache_num: usize,
    /// Source row indexes visible after pivot item filters.
    ///
    /// `None` means every source row is visible. This is a file-format planning
    /// detail, not a cache-authoring API.
    pub visible_rows: Option<Vec<usize>>,
    /// Precomputed axis item tuples in planned cache item-id space.
    ///
    /// For every refreshable local source, both axes are `Some`, including
    /// grouped and measure-sorted axes. `None` is reserved for metadata-only
    /// external, OLAP, scenario, or externally refreshed consolidation caches
    /// that have no cache rows. Grouped tuples use the item-ID spaces exposed
    /// by [`FormatPivotGroupLevel`], with manual groups followed by their base
    /// source item ID.
    pub axis_tuples: FormatPivotAxisTuples,
}

/// Precomputed row/column axis item tuples for file-format writers.
#[derive(Debug, Clone, Default)]
pub struct FormatPivotAxisTuples {
    /// Row-axis item tuples when they are format-neutral.
    pub rows: Option<Vec<Vec<u32>>>,
    /// Column-axis item tuples when they are format-neutral.
    pub columns: Option<Vec<Vec<u32>>>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum FormatPivotCacheKey {
    Transformed(TransformedSnapshotCacheKey, FormatPivotCachePolicyKey),
    MetadataOnly(MetadataOnlyFormatCacheKey, FormatPivotCachePolicyKey),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FormatPivotCachePolicyKey {
    refresh_on_open: bool,
    background_query: bool,
    missing_items_limit: Option<u32>,
}

impl From<&duke_sheets_core::PivotRefreshPolicy> for FormatPivotCachePolicyKey {
    fn from(policy: &duke_sheets_core::PivotRefreshPolicy) -> Self {
        Self {
            refresh_on_open: policy.refresh_on_open,
            background_query: policy.background_query,
            missing_items_limit: policy.missing_items_limit,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct MetadataOnlyFormatCacheKey {
    source: MetadataOnlyFormatSourceCacheKey,
    fields: Vec<String>,
    calculated_fields: Vec<PivotCalculatedField>,
    calculated_items: Vec<PivotCalculatedItem>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum MetadataOnlyFormatSourceCacheKey {
    Consolidation(Vec<duke_sheets_core::PivotSourceRange>),
    External {
        connection_name: String,
        command_text: Option<String>,
    },
    Scenario {
        name: String,
    },
    Olap {
        connection_name: String,
        cube: Option<String>,
        command_text: Option<String>,
    },
}

/// Build immutable file-format pivot plans from a workbook.
pub fn plan_format_pivots(workbook: &Workbook) -> Result<FormatPivotPlan> {
    plan_format_pivots_with_stats(workbook).map(|(plan, _)| plan)
}

/// Return whether a measure identifies the same sort or filter target.
pub fn pivot_measure_matches_target(measure: &PivotMeasure, target: &PivotMeasure) -> bool {
    pivot_measure_matches_sort_target(measure, target)
}

/// Shift a calendar year/month pair by a signed number of months.
pub fn shift_month(year: i32, month: u32, delta: i32) -> Option<(i32, u32)> {
    if !(1..=12).contains(&month) {
        return None;
    }
    let zero_based = year.checked_mul(12)? + month as i32 - 1 + delta;
    Some((
        zero_based.div_euclid(12),
        zero_based.rem_euclid(12) as u32 + 1,
    ))
}

/// Resolve a relative pivot date period to an exclusive serial-number range.
pub fn pivot_date_period_filter_bounds(
    period: PivotDatePeriod,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    use chrono::Datelike;

    let today = chrono::Local::now().date_naive();
    let year = today.year();
    let month = today.month();
    let day = today.day();
    match period {
        PivotDatePeriod::Tomorrow => {
            let date = today.checked_add_signed(chrono::Duration::days(1))?;
            Some(exclusive_day_range(
                date.year(),
                date.month(),
                date.day(),
                date_system,
            ))
        }
        PivotDatePeriod::Today => Some(exclusive_day_range(year, month, day, date_system)),
        PivotDatePeriod::Yesterday => {
            let date = today.checked_sub_signed(chrono::Duration::days(1))?;
            Some(exclusive_day_range(
                date.year(),
                date.month(),
                date.day(),
                date_system,
            ))
        }
        PivotDatePeriod::NextWeek => {
            let date = today.checked_add_signed(chrono::Duration::days(7))?;
            week_filter_bounds(date.year(), date.month(), date.day(), date_system)
        }
        PivotDatePeriod::ThisWeek => week_filter_bounds(year, month, day, date_system),
        PivotDatePeriod::LastWeek => {
            let date = today.checked_sub_signed(chrono::Duration::days(7))?;
            week_filter_bounds(date.year(), date.month(), date.day(), date_system)
        }
        PivotDatePeriod::NextMonth => {
            let (year, month) = shift_month(year, month, 1)?;
            Some(exclusive_month_range(year, month, date_system))
        }
        PivotDatePeriod::ThisMonth => Some(exclusive_month_range(year, month, date_system)),
        PivotDatePeriod::LastMonth => {
            let (year, month) = shift_month(year, month, -1)?;
            Some(exclusive_month_range(year, month, date_system))
        }
        PivotDatePeriod::NextQuarter => {
            let (year, month) = quarter_start_for_shift(year, month, 1)?;
            exclusive_month_span(year, month, 3, date_system)
        }
        PivotDatePeriod::ThisQuarter => {
            exclusive_month_span(year, ((month - 1) / 3) * 3 + 1, 3, date_system)
        }
        PivotDatePeriod::LastQuarter => {
            let (year, month) = quarter_start_for_shift(year, month, -1)?;
            exclusive_month_span(year, month, 3, date_system)
        }
        PivotDatePeriod::NextYear => Some(exclusive_year_range(year + 1, date_system)),
        PivotDatePeriod::ThisYear => Some(exclusive_year_range(year, date_system)),
        PivotDatePeriod::LastYear => Some(exclusive_year_range(year - 1, date_system)),
        PivotDatePeriod::YearToDate => Some((
            date_to_serial(year, 1, 1, date_system),
            date_to_serial(year, month, day, date_system) + 1.0,
        )),
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => None,
    }
}

fn exclusive_day_range(year: i32, month: u32, day: u32, date_system: DateSystem) -> (f64, f64) {
    let start = date_to_serial(year, month, day, date_system);
    (start, start + 1.0)
}

fn week_filter_bounds(
    year: i32,
    month: u32,
    day: u32,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    use chrono::Datelike;

    let date = chrono::NaiveDate::from_ymd_opt(year, month, day)?;
    let start = date.checked_sub_signed(chrono::Duration::days(
        date.weekday().num_days_from_monday() as i64,
    ))?;
    let end = start.checked_add_signed(chrono::Duration::days(7))?;
    Some((
        date_to_serial(start.year(), start.month(), start.day(), date_system),
        date_to_serial(end.year(), end.month(), end.day(), date_system),
    ))
}

fn exclusive_month_range(year: i32, month: u32, date_system: DateSystem) -> (f64, f64) {
    let (end_year, end_month) = shift_month(year, month, 1).unwrap_or((year + 1, 1));
    (
        date_to_serial(year, month, 1, date_system),
        date_to_serial(end_year, end_month, 1, date_system),
    )
}

fn exclusive_month_span(
    year: i32,
    month: u32,
    months: i32,
    date_system: DateSystem,
) -> Option<(f64, f64)> {
    let (end_year, end_month) = shift_month(year, month, months)?;
    Some((
        date_to_serial(year, month, 1, date_system),
        date_to_serial(end_year, end_month, 1, date_system),
    ))
}

fn exclusive_year_range(year: i32, date_system: DateSystem) -> (f64, f64) {
    (
        date_to_serial(year, 1, 1, date_system),
        date_to_serial(year + 1, 1, 1, date_system),
    )
}

fn quarter_start_for_shift(year: i32, month: u32, delta: i32) -> Option<(i32, u32)> {
    shift_month(year, ((month - 1) / 3) * 3 + 1, delta * 3)
}

pub(crate) fn plan_format_pivots_with_stats(
    workbook: &Workbook,
) -> Result<(FormatPivotPlan, PivotRefreshStats)> {
    let mut cache = PivotRuntimeCache::clone_from_workbook(workbook);
    let mut stats = PivotRefreshStats::default();
    let mut cache_by_key: AHashMap<FormatPivotCacheKey, usize> = AHashMap::new();
    let mut caches: Vec<FormatPivotCache> = Vec::new();
    let mut tables = Vec::new();

    for (sheet_index, worksheet) in workbook.worksheets().enumerate() {
        for (pivot_index, pivot) in worksheet.pivot_tables().iter().enumerate() {
            validate_format_pivot(pivot)?;
            if metadata_only_format_source(&pivot.source).is_some() {
                let key = FormatPivotCacheKey::MetadataOnly(
                    metadata_only_format_cache_key(pivot).ok_or_else(|| {
                        Error::other("metadata-only pivot source has no cache key")
                    })?,
                    (&pivot.refresh_policy).into(),
                );
                let cache_num = if let Some(cache_num) = cache_by_key.get(&key).copied() {
                    if let Some(existing) = caches.get_mut(cache_num - 1) {
                        existing.refresh_on_load |= pivot.refresh_policy.refresh_on_open;
                        existing.background_query |= pivot.refresh_policy.background_query;
                    }
                    cache_num
                } else {
                    let cache_num = caches.len() + 1;
                    let planned_cache = build_metadata_only_format_pivot_cache(cache_num, pivot)?;
                    cache_by_key.insert(key, cache_num);
                    caches.push(planned_cache);
                    cache_num
                };

                tables.push(FormatPivotTable {
                    sheet_index,
                    pivot_index,
                    table_num: tables.len() + 1,
                    cache_num,
                    visible_rows: None,
                    axis_tuples: FormatPivotAxisTuples::default(),
                });
                continue;
            }

            let resolved = resolve_source(workbook, sheet_index, &pivot.source)?;
            let source = format_pivot_source(workbook, &resolved, &pivot.source)?;
            let source_snapshot =
                cache.snapshot_for_resolved_source(workbook, resolved, &mut stats)?;
            let raw_snapshot = Arc::clone(&source_snapshot.snapshot);
            let key = FormatPivotCacheKey::Transformed(
                TransformedSnapshotCacheKey::new(
                    source_snapshot.key.clone(),
                    &pivot.calculated_fields,
                    &pivot.groupings,
                    &pivot.calculated_items,
                    workbook.settings().date_1904,
                ),
                (&pivot.refresh_policy).into(),
            );
            let snapshot = transformed_snapshot_for_pivot(
                workbook,
                sheet_index,
                pivot,
                source_snapshot,
                workbook.settings().date_1904,
                &mut cache,
            )?;
            validate_format_pivot_fields(&pivot.name, pivot, &snapshot)?;

            let cache_num = if let Some(cache_num) = cache_by_key.get(&key).copied() {
                if let Some(existing) = caches.get_mut(cache_num - 1) {
                    existing.refresh_on_load |= pivot.refresh_policy.refresh_on_open;
                    existing.background_query |= pivot.refresh_policy.background_query;
                }
                cache_num
            } else {
                let cache_num = caches.len() + 1;
                let planned_cache = build_format_pivot_cache(
                    cache_num,
                    source,
                    &raw_snapshot,
                    &snapshot,
                    pivot,
                    workbook.settings().date_1904,
                )?;
                cache_by_key.insert(key, cache_num);
                caches.push(planned_cache);
                cache_num
            };

            let visible_rows = format_pivot_visible_rows(pivot, &snapshot)?;
            let planned_cache = caches.get(cache_num - 1).ok_or_else(|| {
                Error::other("planned pivot cache disappeared while building axis tuples")
            })?;
            let axis_tuples = format_pivot_axis_tuples(
                pivot,
                &raw_snapshot,
                &snapshot,
                planned_cache,
                visible_rows.as_deref(),
            )?;

            tables.push(FormatPivotTable {
                sheet_index,
                pivot_index,
                table_num: tables.len() + 1,
                cache_num,
                visible_rows,
                axis_tuples,
            });
        }
    }

    Ok((FormatPivotPlan { caches, tables }, stats))
}

pub(crate) fn validate_format_pivot(pivot: &PivotTable) -> Result<()> {
    if pivot.measures.is_empty() {
        return Err(Error::other(format!(
            "pivot table {} has no measures",
            pivot.name
        )));
    }
    Ok(())
}

pub(crate) fn format_pivot_source(
    workbook: &Workbook,
    resolved: &ResolvedPivotSource,
    semantic_source: &PivotSource,
) -> Result<FormatPivotSource> {
    match resolved {
        ResolvedPivotSource::Single(source) => {
            let worksheet = workbook.worksheet(source.sheet_index).ok_or_else(|| {
                Error::SheetOutOfBounds(source.sheet_index, workbook.sheet_count())
            })?;
            Ok(FormatPivotSource::Worksheet {
                sheet_index: source.sheet_index,
                sheet_name: worksheet.name().to_string(),
                range: source.range,
                table_name: matches!(source.kind, SourceCacheKind::Table)
                    .then(|| source.source_name.clone())
                    .flatten(),
            })
        }
        ResolvedPivotSource::Consolidation(_) => {
            let PivotSource::Consolidation { ranges } = semantic_source else {
                return Err(Error::other(
                    "consolidation format source must retain semantic ranges",
                ));
            };
            Ok(FormatPivotSource::Consolidation {
                ranges: ranges.clone(),
            })
        }
    }
}

pub(crate) fn build_format_pivot_cache(
    cache_num: usize,
    source: FormatPivotSource,
    raw_snapshot: &SourceSnapshot,
    snapshot: &SourceSnapshot,
    pivot: &PivotTable,
    date_1904: bool,
) -> Result<FormatPivotCache> {
    if matches!(source, FormatPivotSource::Consolidation { .. }) {
        return build_consolidation_format_pivot_cache(cache_num, source, snapshot, pivot);
    }

    let calculated_fields = pivot
        .calculated_fields
        .iter()
        .map(|field| (field.name.to_lowercase(), field.formula.clone()))
        .collect::<AHashMap<_, _>>();

    let mut fields = snapshot
        .headers
        .iter()
        .zip(snapshot.columns.iter())
        .map(|(name, column)| {
            let formula = calculated_fields.get(&name.to_lowercase()).cloned();
            FormatPivotCacheField {
                name: name.clone(),
                database_field: formula.is_none(),
                formula,
                shared_items: column.dictionary.clone(),
                item_ids: column.values.clone(),
                grouping: None,
            }
        })
        .collect::<Vec<_>>();
    attach_format_groupings(&mut fields, raw_snapshot, snapshot, pivot, date_1904)?;

    Ok(FormatPivotCache {
        cache_num,
        source,
        fields,
        field_aliases: Vec::new(),
        calculated_items: pivot.calculated_items.clone(),
        row_count: snapshot.row_count,
        save_data: true,
        refresh_on_load: pivot.refresh_policy.refresh_on_open,
        background_query: pivot.refresh_policy.background_query,
        missing_items_limit: pivot.refresh_policy.missing_items_limit,
    })
}

fn attach_format_groupings(
    fields: &mut [FormatPivotCacheField],
    raw_snapshot: &SourceSnapshot,
    snapshot: &SourceSnapshot,
    pivot: &PivotTable,
    date_1904: bool,
) -> Result<()> {
    let mut claimed_derived_fields = AHashSet::new();
    for definition in &pivot.groupings {
        let field_name = grouping_field_name(definition);
        let base_field_index = raw_snapshot.field_index(field_name).ok_or_else(|| {
            Error::other(format!(
                "pivot table {} references missing grouping field: {field_name}",
                pivot.name
            ))
        })?;
        let source = raw_snapshot.columns.get(base_field_index).ok_or_else(|| {
            Error::other("pivot grouping base field is missing from the source snapshot")
        })?;
        let mut levels = Vec::new();
        match definition {
            PivotGrouping::Date { units, .. } if units.len() > 1 => {
                let mut parent_field_index = None;
                for unit in units {
                    let grouped =
                        raw_snapshot.grouped_date_column(base_field_index, &[*unit], date_1904);
                    let field_index = snapshot
                        .headers
                        .iter()
                        .enumerate()
                        .skip(raw_snapshot.headers.len())
                        .find(|(index, name)| {
                            !claimed_derived_fields.contains(index)
                                && name.starts_with(&grouped_date_header(field_name, *unit))
                        })
                        .map(|(index, _)| index)
                        .ok_or_else(|| {
                            Error::other(format!(
                                "pivot table {} is missing transformed date grouping field {field_name}",
                                pivot.name
                            ))
                        })?;
                    claimed_derived_fields.insert(field_index);
                    levels.push(format_group_level(
                        source,
                        &grouped,
                        field_index,
                        parent_field_index,
                        Some(*unit),
                    )?);
                    parent_field_index = Some(field_index);
                }
            }
            PivotGrouping::Manual { groups, .. } => {
                levels.push(format_manual_group_level(
                    source,
                    groups,
                    base_field_index,
                    &pivot.name,
                )?);
            }
            PivotGrouping::Number { .. } | PivotGrouping::Date { .. } => {
                let grouped = raw_snapshot.grouped_column(
                    base_field_index,
                    definition,
                    date_1904,
                    &pivot.name,
                )?;
                let date_unit = match definition {
                    PivotGrouping::Date { units, .. } => units.first().copied(),
                    _ => None,
                };
                levels.push(format_group_level(
                    source,
                    &grouped,
                    base_field_index,
                    None,
                    date_unit,
                )?);
            }
        }
        fields[base_field_index].grouping = Some(FormatPivotGrouping {
            definition: definition.clone(),
            base_field_index,
            parent_field_index: None,
            source_items: source.dictionary.clone(),
            source_item_ids: source.values.clone(),
            levels,
        });
    }
    Ok(())
}

fn format_group_level(
    source: &EncodedColumn,
    grouped: &EncodedColumn,
    field_index: usize,
    parent_field_index: Option<usize>,
    date_unit: Option<duke_sheets_core::PivotDateGroupUnit>,
) -> Result<FormatPivotGroupLevel> {
    let source_item_group_ids = source
        .dictionary
        .iter()
        .map(|value| {
            let source_id = source
                .id_for_value(value)
                .ok_or_else(|| Error::other("pivot grouping source dictionary lookup failed"))?;
            let row = source
                .values
                .iter()
                .position(|id| *id == source_id)
                .ok_or_else(|| Error::other("pivot grouping source item has no cache row"))?;
            grouped
                .values
                .get(row)
                .copied()
                .ok_or_else(|| Error::other("pivot grouping row mapping is incomplete"))
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(FormatPivotGroupLevel {
        field_index,
        parent_field_index,
        date_unit,
        group_items: grouped.dictionary.clone(),
        source_item_group_ids,
        item_ids: grouped.values.clone(),
    })
}

fn format_manual_group_level(
    source: &EncodedColumn,
    groups: &[PivotManualGroup],
    field_index: usize,
    pivot_name: &str,
) -> Result<FormatPivotGroupLevel> {
    let lookup = manual_group_lookup(groups, pivot_name)?;
    for group in groups {
        for member in &group.members {
            if !source.dictionary.iter().any(|item| item == member) {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} manual group {} references an item not found in the source field: {member}",
                    group.name
                )));
            }
        }
    }
    let mut group_items = source
        .dictionary
        .iter()
        .filter(|item| !lookup.contains_key(*item))
        .cloned()
        .collect::<Vec<_>>();
    let group_indexes = groups
        .iter()
        .map(|group| {
            let index = group_items.len() as u32;
            group_items.push(PivotValue::String(group.name.clone()));
            (group.name.clone(), index)
        })
        .collect::<AHashMap<_, _>>();
    let source_item_group_ids = source
        .dictionary
        .iter()
        .map(|item| {
            lookup
                .get(item)
                .and_then(|name| group_indexes.get(name).copied())
                .or_else(|| {
                    group_items
                        .iter()
                        .position(|value| value == item)
                        .map(|index| index as u32)
                })
                .ok_or_else(|| Error::other("pivot manual grouping item mapping is incomplete"))
        })
        .collect::<Result<Vec<_>>>()?;
    let item_ids = source
        .values
        .iter()
        .map(|item_id| source_item_group_ids[*item_id as usize])
        .collect();
    Ok(FormatPivotGroupLevel {
        field_index,
        parent_field_index: None,
        date_unit: None,
        group_items,
        source_item_group_ids,
        item_ids,
    })
}

pub(crate) fn build_consolidation_format_pivot_cache(
    cache_num: usize,
    source: FormatPivotSource,
    snapshot: &SourceSnapshot,
    pivot: &PivotTable,
) -> Result<FormatPivotCache> {
    let FormatPivotSource::Consolidation { ranges } = &source else {
        return Err(Error::other(
            "consolidation format cache requires consolidation source metadata",
        ));
    };
    if snapshot.headers.len() < 2 {
        return Err(Error::other(format!(
            "pivot table {} consolidation sources require at least one row-label column and one value column",
            pivot.name
        )));
    }

    let value_column_count = snapshot.headers.len() - 1;
    let row_count = snapshot.row_count * value_column_count;
    let page_count = consolidation_page_count(ranges)?;
    let row_sources = consolidation_snapshot_row_sources(ranges, snapshot.row_count)?;
    let mut row_field = EncodedColumn::with_capacity(row_count);
    let mut column_field = EncodedColumn::with_capacity(row_count);
    let mut value_field = EncodedColumn::with_capacity(row_count);
    let mut page_fields = (0..page_count)
        .map(|_| EncodedColumn::with_capacity(row_count))
        .collect::<Vec<_>>();

    for row in 0..snapshot.row_count {
        let source_index = row_sources[row];
        for source_col in 1..snapshot.headers.len() {
            row_field.push(snapshot.value(row, 0).clone());
            column_field.push(PivotValue::String(snapshot.headers[source_col].clone()));
            value_field.push(snapshot.value(row, source_col).clone());
            for (page_index, page_field) in page_fields.iter_mut().enumerate() {
                let value = ranges[source_index]
                    .page_items
                    .get(page_index)
                    .map(|item| PivotValue::String(item.clone()))
                    .unwrap_or(PivotValue::Blank);
                page_field.push(value);
            }
        }
    }

    let mut fields = vec![
        FormatPivotCacheField {
            name: "Row".to_string(),
            formula: None,
            database_field: true,
            shared_items: row_field.dictionary,
            item_ids: row_field.values,
            grouping: None,
        },
        FormatPivotCacheField {
            name: "Column".to_string(),
            formula: None,
            database_field: true,
            shared_items: column_field.dictionary,
            item_ids: column_field.values,
            grouping: None,
        },
        FormatPivotCacheField {
            name: "Value".to_string(),
            formula: None,
            database_field: true,
            shared_items: value_field.dictionary,
            item_ids: value_field.values,
            grouping: None,
        },
    ];
    for (index, page_field) in page_fields.into_iter().enumerate() {
        fields.push(FormatPivotCacheField {
            name: format!("Page{}", index + 1),
            formula: None,
            database_field: true,
            shared_items: page_field.dictionary,
            item_ids: page_field.values,
            grouping: None,
        });
    }

    let mut field_aliases = Vec::with_capacity(snapshot.headers.len());
    field_aliases.push((snapshot.headers[0].clone(), "Row".to_string()));
    for header in snapshot.headers.iter().skip(1) {
        field_aliases.push((header.clone(), "Value".to_string()));
    }

    Ok(FormatPivotCache {
        cache_num,
        source,
        fields,
        field_aliases,
        calculated_items: pivot.calculated_items.clone(),
        row_count,
        save_data: true,
        refresh_on_load: pivot.refresh_policy.refresh_on_open,
        background_query: pivot.refresh_policy.background_query,
        missing_items_limit: pivot.refresh_policy.missing_items_limit,
    })
}

pub(crate) fn consolidation_page_count(
    ranges: &[duke_sheets_core::PivotSourceRange],
) -> Result<usize> {
    let page_count = ranges
        .iter()
        .map(|range| range.page_items.len())
        .max()
        .unwrap_or(0);
    if page_count > 4 {
        return Err(Error::other(
            "consolidation pivot sources support at most four page fields",
        ));
    }
    for range in ranges {
        for item in &range.page_items {
            if item.trim().is_empty() {
                return Err(Error::other(
                    "consolidation pivot source page item names cannot be blank",
                ));
            }
        }
    }
    Ok(page_count)
}

pub(crate) fn consolidation_snapshot_row_sources(
    ranges: &[duke_sheets_core::PivotSourceRange],
    snapshot_row_count: usize,
) -> Result<Vec<usize>> {
    let mut row_sources = Vec::with_capacity(snapshot_row_count);
    for (range_index, range) in ranges.iter().enumerate() {
        let Some(source_range) = range.range else {
            return Err(Error::other(
                "local consolidation cache planning requires concrete source ranges",
            ));
        };
        let row_count = source_range.row_count().saturating_sub(1) as usize;
        row_sources.extend(std::iter::repeat(range_index).take(row_count));
    }
    if row_sources.len() != snapshot_row_count {
        return Err(Error::other(
            "consolidation source row count changed while planning pivot cache",
        ));
    }
    Ok(row_sources)
}

pub(crate) fn build_metadata_only_format_pivot_cache(
    cache_num: usize,
    pivot: &PivotTable,
) -> Result<FormatPivotCache> {
    let source = metadata_only_format_source(&pivot.source).ok_or_else(|| {
        Error::other("metadata-only format cache requires a non-refreshable pivot source")
    })?;

    Ok(FormatPivotCache {
        cache_num,
        source,
        fields: metadata_only_format_cache_fields(pivot)?,
        field_aliases: Vec::new(),
        calculated_items: pivot.calculated_items.clone(),
        row_count: 0,
        save_data: false,
        refresh_on_load: pivot.refresh_policy.refresh_on_open,
        background_query: pivot.refresh_policy.background_query,
        missing_items_limit: pivot.refresh_policy.missing_items_limit,
    })
}

pub(crate) fn metadata_only_format_source(source: &PivotSource) -> Option<FormatPivotSource> {
    match source {
        PivotSource::External {
            connection_name,
            command_text,
        } => Some(FormatPivotSource::External {
            connection_name: connection_name.clone(),
            command_text: command_text.clone(),
        }),
        PivotSource::Consolidation { ranges } if source_requires_external_refresh(source) => {
            Some(FormatPivotSource::Consolidation {
                ranges: ranges.clone(),
            })
        }
        PivotSource::Scenario { name } => Some(FormatPivotSource::Scenario { name: name.clone() }),
        PivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } => Some(FormatPivotSource::Olap {
            connection_name: connection_name.clone(),
            cube: cube.clone(),
            command_text: command_text.clone(),
        }),
        _ => None,
    }
}

pub(crate) fn metadata_only_format_cache_fields(
    pivot: &PivotTable,
) -> Result<Vec<FormatPivotCacheField>> {
    let calculated_names = pivot
        .calculated_fields
        .iter()
        .map(|field| field.name.to_lowercase())
        .collect::<AHashSet<_>>();
    let mut fields = Vec::new();
    let mut seen = AHashSet::new();

    for field_name in pivot
        .rows
        .iter()
        .map(|field| field.field.name.as_str())
        .chain(pivot.columns.iter().map(|field| field.field.name.as_str()))
        .chain(
            pivot
                .page_fields
                .iter()
                .map(|field| field.field.name.as_str()),
        )
        .chain(
            pivot
                .measures
                .iter()
                .map(|measure| measure.field.name.as_str()),
        )
        .chain(pivot.filters.iter().filter_map(format_filter_field_name))
        .chain(
            pivot
                .filters
                .iter()
                .filter_map(format_filter_measure_field_name),
        )
        .chain(pivot.groupings.iter().map(grouping_field_name))
        .chain(
            pivot
                .measures
                .iter()
                .filter_map(format_measure_show_as_base_field_name),
        )
        .chain(
            pivot
                .calculated_items
                .iter()
                .map(|item| item.field.name.as_str()),
        )
    {
        if calculated_names.contains(&field_name.to_lowercase()) {
            continue;
        }
        push_metadata_only_format_cache_field(&mut fields, &mut seen, field_name);
    }

    for field in &pivot.calculated_fields {
        if field.name.trim().is_empty() {
            return Err(Error::other(format!(
                "pivot table {} has a calculated field with a blank name",
                pivot.name
            )));
        }
        if !seen.insert(field.name.to_lowercase()) {
            return Err(Error::other(format!(
                "pivot table {} calculated field duplicates source field: {}",
                pivot.name, field.name
            )));
        }
        fields.push(FormatPivotCacheField {
            name: field.name.clone(),
            formula: Some(field.formula.clone()),
            database_field: false,
            shared_items: Vec::new(),
            item_ids: Vec::new(),
            grouping: None,
        });
    }

    Ok(fields)
}

pub(crate) fn push_metadata_only_format_cache_field(
    fields: &mut Vec<FormatPivotCacheField>,
    seen: &mut AHashSet<String>,
    name: &str,
) {
    if seen.insert(name.to_lowercase()) {
        fields.push(FormatPivotCacheField {
            name: name.to_string(),
            formula: None,
            database_field: true,
            shared_items: Vec::new(),
            item_ids: Vec::new(),
            grouping: None,
        });
    }
}

fn metadata_only_format_cache_key(pivot: &PivotTable) -> Option<MetadataOnlyFormatCacheKey> {
    let source = match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            MetadataOnlyFormatSourceCacheKey::Consolidation(ranges.clone())
        }
        PivotSource::External {
            connection_name,
            command_text,
        } => MetadataOnlyFormatSourceCacheKey::External {
            connection_name: connection_name.clone(),
            command_text: command_text.clone(),
        },
        PivotSource::Scenario { name } => {
            MetadataOnlyFormatSourceCacheKey::Scenario { name: name.clone() }
        }
        PivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } => MetadataOnlyFormatSourceCacheKey::Olap {
            connection_name: connection_name.clone(),
            cube: cube.clone(),
            command_text: command_text.clone(),
        },
        PivotSource::WorksheetRange { .. } | PivotSource::Table { .. } => return None,
    };
    Some(MetadataOnlyFormatCacheKey {
        source,
        fields: metadata_only_format_cache_field_names_for_key(pivot),
        calculated_fields: pivot.calculated_fields.clone(),
        calculated_items: pivot.calculated_items.clone(),
    })
}

pub(crate) fn metadata_only_format_cache_field_names_for_key(pivot: &PivotTable) -> Vec<String> {
    let mut fields = Vec::new();
    let mut seen = AHashSet::new();
    for field_name in pivot
        .rows
        .iter()
        .map(|field| field.field.name.as_str())
        .chain(pivot.columns.iter().map(|field| field.field.name.as_str()))
        .chain(
            pivot
                .page_fields
                .iter()
                .map(|field| field.field.name.as_str()),
        )
        .chain(
            pivot
                .measures
                .iter()
                .map(|measure| measure.field.name.as_str()),
        )
        .chain(pivot.filters.iter().filter_map(format_filter_field_name))
        .chain(
            pivot
                .filters
                .iter()
                .filter_map(format_filter_measure_field_name),
        )
        .chain(pivot.groupings.iter().map(grouping_field_name))
        .chain(
            pivot
                .measures
                .iter()
                .filter_map(format_measure_show_as_base_field_name),
        )
        .chain(
            pivot
                .calculated_items
                .iter()
                .map(|item| item.field.name.as_str()),
        )
    {
        let key = field_name.to_lowercase();
        if seen.insert(key) {
            fields.push(field_name.to_string());
        }
    }
    fields
}

pub(crate) fn format_pivot_visible_rows(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
) -> Result<Option<Vec<usize>>> {
    let item_filters = pivot
        .filters
        .iter()
        .filter_map(|filter| {
            let PivotFilter::FieldItems {
                field,
                allowed_items,
            } = filter
            else {
                return None;
            };
            let field_index = snapshot.field_index(&field.name)?;
            Some((field_index, allowed_items.as_slice()))
        })
        .collect::<Vec<_>>();

    if item_filters.is_empty() {
        return Ok(None);
    }

    let mut visible_rows = Vec::new();
    'row: for row in 0..snapshot.row_count {
        for (field_index, allowed_items) in &item_filters {
            let Some(column) = snapshot.columns.get(*field_index) else {
                continue 'row;
            };
            let Some(item_id) = column.values.get(row).copied() else {
                continue 'row;
            };
            let Some(item) = column.dictionary.get(item_id as usize) else {
                continue 'row;
            };
            if !allowed_items.iter().any(|allowed| allowed == item) {
                continue 'row;
            }
        }
        visible_rows.push(row);
    }

    Ok(Some(visible_rows))
}

pub(crate) fn format_pivot_axis_tuples(
    pivot: &PivotTable,
    raw_snapshot: &SourceSnapshot,
    snapshot: &SourceSnapshot,
    cache: &FormatPivotCache,
    visible_rows: Option<&[usize]>,
) -> Result<FormatPivotAxisTuples> {
    Ok(FormatPivotAxisTuples {
        rows: format_axis_tuples_for_fields(
            pivot,
            raw_snapshot,
            snapshot,
            cache,
            &pivot.rows,
            visible_rows,
        )?,
        columns: format_axis_tuples_for_fields(
            pivot,
            raw_snapshot,
            snapshot,
            cache,
            &pivot.columns,
            visible_rows,
        )?,
    })
}

pub(crate) fn format_axis_tuples_for_fields(
    pivot: &PivotTable,
    raw_snapshot: &SourceSnapshot,
    snapshot: &SourceSnapshot,
    cache: &FormatPivotCache,
    fields: &[PivotField],
    visible_rows: Option<&[usize]>,
) -> Result<Option<Vec<Vec<u32>>>> {
    if fields.is_empty() {
        return Ok(Some(Vec::new()));
    }

    let (field_indexes, sort_fields) = format_axis_sort_fields(pivot, snapshot, cache, fields)?;
    let mut semantic_tuples = format_unique_axis_tuples(snapshot, &field_indexes, visible_rows);
    sort_format_axis_tuples_by_measure(
        pivot,
        snapshot,
        &field_indexes,
        &sort_fields,
        visible_rows,
        &mut semantic_tuples,
    )?;
    let semantic_positions = semantic_tuples
        .iter()
        .cloned()
        .enumerate()
        .map(|(index, tuple)| (tuple, index))
        .collect::<AHashMap<_, _>>();
    let mut seen = AHashSet::new();
    let mut tuples = Vec::new();
    for row in visible_row_indexes(visible_rows, snapshot.row_count) {
        let semantic = encoded_key(snapshot, &field_indexes, row);
        let expanded = format_expanded_axis_tuple(raw_snapshot, cache, fields, row)?;
        if seen.insert(expanded.clone()) {
            tuples.push((
                semantic_positions
                    .get(&semantic)
                    .copied()
                    .unwrap_or(usize::MAX),
                expanded,
            ));
        }
    }
    tuples.sort_by_key(|(position, _)| *position);
    Ok(Some(tuples.into_iter().map(|(_, tuple)| tuple).collect()))
}

fn format_axis_sort_fields(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    cache: &FormatPivotCache,
    fields: &[PivotField],
) -> Result<(Vec<usize>, Vec<PivotField>)> {
    let mut indexes = Vec::new();
    let mut sort_fields = Vec::new();
    for field in fields {
        let cache_index = cache.field_index(&field.field.name).ok_or_else(|| {
            Error::other(format!(
                "pivot table {} references unknown axis field: {}",
                pivot.name, field.field.name
            ))
        })?;
        if let Some(grouping) = &cache.fields[cache_index].grouping {
            if matches!(grouping.definition, PivotGrouping::Date { ref units, .. } if units.len() > 1)
            {
                for level in &grouping.levels {
                    indexes.push(level.field_index);
                    sort_fields.push(field.clone());
                }
                continue;
            }
        }
        indexes.push(snapshot.required_field_index(&field.field.name, &pivot.name)?);
        sort_fields.push(field.clone());
    }
    Ok((indexes, sort_fields))
}

fn format_expanded_axis_tuple(
    _raw_snapshot: &SourceSnapshot,
    cache: &FormatPivotCache,
    fields: &[PivotField],
    row: usize,
) -> Result<Vec<u32>> {
    let mut tuple = Vec::new();
    for field in fields {
        let cache_index = cache.field_index(&field.field.name).ok_or_else(|| {
            Error::other(format!(
                "pivot references unknown axis field {}",
                field.field.name
            ))
        })?;
        let cache_field = &cache.fields[cache_index];
        if let Some(grouping) = &cache_field.grouping {
            match &grouping.definition {
                PivotGrouping::Manual { .. } => {
                    tuple.push(grouping.levels[0].item_ids[row]);
                    tuple.push(grouping.source_item_ids[row]);
                }
                PivotGrouping::Date { units, .. } if units.len() > 1 => {
                    tuple.extend(grouping.levels.iter().map(|level| level.item_ids[row]));
                }
                PivotGrouping::Number { .. } | PivotGrouping::Date { .. } => {
                    tuple.push(grouping.levels[0].item_ids[row]);
                }
            }
        } else {
            tuple.push(cache_field.item_ids[row]);
        }
    }
    Ok(tuple)
}

pub(crate) fn format_unique_axis_tuples(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    visible_rows: Option<&[usize]>,
) -> Vec<Vec<u32>> {
    let mut seen = AHashSet::new();
    let mut tuples = Vec::new();
    for row in visible_row_indexes(visible_rows, snapshot.row_count) {
        let tuple = encoded_key(snapshot, field_indexes, row);
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    tuples
}

pub(crate) fn sort_format_axis_tuples_by_measure(
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    fields: &[PivotField],
    visible_rows: Option<&[usize]>,
    tuples: &mut [Vec<u32>],
) -> Result<()> {
    if tuples.len() < 2
        || !fields
            .iter()
            .any(|field| !matches!(field.sort, PivotSort::None) && field.sort_by_measure.is_some())
    {
        return Ok(());
    }

    let measure_indexes = pivot
        .measures
        .iter()
        .map(|measure| snapshot.required_field_index(&measure.field.name, &pivot.name))
        .collect::<Result<Vec<_>>>()?;
    let sort_measure_indexes =
        compile_axis_sort_measure_indexes(&pivot.name, fields, &pivot.measures)?;
    let totals = format_axis_tuple_measure_totals(
        snapshot,
        field_indexes,
        &measure_indexes,
        &pivot.measures,
        visible_rows,
    );

    sort_key_order(
        tuples,
        field_indexes,
        fields,
        &sort_measure_indexes,
        &totals,
        &pivot.measures,
        snapshot,
    );
    Ok(())
}

pub(crate) fn format_axis_tuple_measure_totals(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    measure_indexes: &[usize],
    measures: &[PivotMeasure],
    visible_rows: Option<&[usize]>,
) -> AHashMap<Vec<u32>, Vec<AggregateState>> {
    let mut totals = AHashMap::<Vec<u32>, Vec<AggregateState>>::new();
    for row in visible_row_indexes(visible_rows, snapshot.row_count) {
        let tuple = encoded_key(snapshot, field_indexes, row);
        let states = totals
            .entry(tuple)
            .or_insert_with(|| default_states(measures));
        format_update_states(states, snapshot, measure_indexes, measures, row);
    }
    totals
}

pub(crate) fn format_update_states(
    states: &mut [AggregateState],
    snapshot: &SourceSnapshot,
    measure_indexes: &[usize],
    measures: &[PivotMeasure],
    row: usize,
) {
    for ((state, field_index), measure) in states
        .iter_mut()
        .zip(measure_indexes.iter())
        .zip(measures.iter())
    {
        state.update(snapshot.value(row, *field_index), measure.aggregate);
    }
}

/// Iterate all row indexes or a prefiltered subset without allocating.
pub fn visible_row_indexes(
    visible_rows: Option<&[usize]>,
    row_count: usize,
) -> FormatVisibleRowIter<'_> {
    match visible_rows {
        Some(rows) => FormatVisibleRowIter::Filtered(rows.iter().copied()),
        None => FormatVisibleRowIter::All(0..row_count),
    }
}

/// Iterator returned by [`visible_row_indexes`].
pub enum FormatVisibleRowIter<'a> {
    /// Every cache row in source order.
    All(std::ops::Range<usize>),
    /// A prefiltered sequence of cache row indexes.
    Filtered(std::iter::Copied<std::slice::Iter<'a, usize>>),
}

impl Iterator for FormatVisibleRowIter<'_> {
    type Item = usize;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::All(rows) => rows.next(),
            Self::Filtered(rows) => rows.next(),
        }
    }
}

pub(crate) fn validate_format_pivot_fields(
    pivot_name: &str,
    pivot: &PivotTable,
    snapshot: &SourceSnapshot,
) -> Result<()> {
    for field_name in pivot
        .rows
        .iter()
        .map(|field| field.field.name.as_str())
        .chain(pivot.columns.iter().map(|field| field.field.name.as_str()))
        .chain(
            pivot
                .page_fields
                .iter()
                .map(|field| field.field.name.as_str()),
        )
        .chain(
            pivot
                .measures
                .iter()
                .map(|measure| measure.field.name.as_str()),
        )
        .chain(pivot.filters.iter().filter_map(format_filter_field_name))
    {
        if snapshot.field_index(field_name).is_none() {
            return Err(Error::other(format!(
                "pivot table {pivot_name} references unknown source field: {field_name}"
            )));
        }
    }
    Ok(())
}

pub(crate) fn format_filter_field_name(filter: &PivotFilter) -> Option<&str> {
    match filter {
        PivotFilter::FieldItems { field, .. }
        | PivotFilter::Label { field, .. }
        | PivotFilter::LabelBetween { field, .. }
        | PivotFilter::Date { field, .. }
        | PivotFilter::DateBetween { field, .. }
        | PivotFilter::DatePeriod { field, .. }
        | PivotFilter::Value { field, .. }
        | PivotFilter::ValueBetween { field, .. }
        | PivotFilter::TopN { field, .. } => Some(field.name.as_str()),
        PivotFilter::Unsupported { .. } => None,
    }
}

pub(crate) fn format_filter_measure_field_name(filter: &PivotFilter) -> Option<&str> {
    match filter {
        PivotFilter::Value { measure, .. }
        | PivotFilter::ValueBetween { measure, .. }
        | PivotFilter::TopN { measure, .. } => Some(measure.field.name.as_str()),
        _ => None,
    }
}

pub(crate) fn format_measure_show_as_base_field_name(measure: &PivotMeasure) -> Option<&str> {
    match &measure.show_as {
        PivotShowAs::PercentOfParentTotal { base_field }
        | PivotShowAs::RunningTotal { base_field }
        | PivotShowAs::DifferenceFrom { base_field, .. }
        | PivotShowAs::PercentDifferenceFrom { base_field, .. }
        | PivotShowAs::RankAscending { base_field }
        | PivotShowAs::RankDescending { base_field } => Some(base_field.name.as_str()),
        _ => None,
    }
}
