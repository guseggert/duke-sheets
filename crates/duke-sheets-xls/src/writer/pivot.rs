use super::*;

pub(super) struct PivotLayouts(XlsPivotCacheLayouts);

pub(super) fn build_layouts(
    workbook: &Workbook,
    plan: &FormatPivotPlan,
) -> XlsResult<PivotLayouts> {
    build_xls_pivot_cache_layouts(workbook, plan).map(PivotLayouts)
}

pub(super) fn build_cache_streams(
    plan: &FormatPivotPlan,
    layouts: &PivotLayouts,
) -> XlsResult<(Vec<String>, Vec<(String, Vec<u8>)>)> {
    build_pivot_cache_streams(plan, &layouts.0)
}

pub(super) fn sheet_has_tables(plan: &FormatPivotPlan, sheet_idx: usize) -> bool {
    sheet_has_pivot_tables(plan, sheet_idx)
}

pub(super) fn write_table_styles(stream: &mut Vec<u8>, plan: &FormatPivotPlan) {
    write_table_styles_record(stream, plan);
}

pub(super) fn write_pre_boundsheet_records(
    stream: &mut Vec<u8>,
    plan: &FormatPivotPlan,
) -> XlsResult<()> {
    write_pivot_pre_boundsheet_records(stream, plan)
}

pub(super) fn write_global_records(stream: &mut Vec<u8>, plan: &FormatPivotPlan) {
    write_pivot_global_records(stream, plan);
}

pub(super) fn write_workbook_extension_records(
    stream: &mut Vec<u8>,
    plan: &FormatPivotPlan,
) -> XlsResult<()> {
    write_pivot_workbook_extension_records(stream, plan)
}

pub(super) fn write_sheet_records(
    stream: &mut Vec<u8>,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    layouts: &PivotLayouts,
    sheet_idx: usize,
    styles: &StyleTables,
) -> XlsResult<()> {
    write_pivot_sheet_records(stream, workbook, plan, &layouts.0, sheet_idx, styles)
}

pub(super) fn write_sheet_tail_records(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    plan: &FormatPivotPlan,
    sheet_idx: usize,
) {
    write_pivot_sheet_tail_records(stream, sheet, plan, sheet_idx);
}

const PIVOT_SXFMLA_RECORD: u16 = 0x00F9;
const PIVOT_SXNAME_RECORD: u16 = 0x00F6;
const PIVOT_SXPAIR_RECORD: u16 = 0x00F8;
const PIVOT_SXFORMULA_RECORD: u16 = 0x0103;
const PIVOT_SXFDB_RECORD: u16 = records::SXFDB;
const PIVOT_SXDBB_RECORD: u16 = records::SXDBB;
const PIVOT_SXNUM_RECORD: u16 = records::SXNUM;
const PIVOT_SXBOOL_RECORD: u16 = records::SXBOOL;
const PIVOT_SXERR_RECORD: u16 = records::SXERR;
const PIVOT_SXINT_RECORD: u16 = records::SXINT;
const PIVOT_SXSTRING_RECORD: u16 = records::SXSTRING;
const PIVOT_SXDTR_RECORD: u16 = records::SXDTR;
const PIVOT_SXRNG_RECORD: u16 = records::SXRNG;
const PIVOT_SXIDSTM_RECORD: u16 = records::SXIDSTM;

struct XlsPivotGroupingInfo<'a> {
    grouping: &'a PivotGrouping,
    source_numbers: Vec<f64>,
    source_items: Vec<PivotValue>,
    source_item_ids: Vec<u32>,
    group_items: Vec<PivotValue>,
    base_item_group_ids: Vec<u32>,
    group_item_ids: Vec<u32>,
}

#[derive(Debug, Clone)]
struct XlsPivotCacheLayout {
    cache_num: usize,
    row_count: usize,
    base_field_count: usize,
    is_consolidation: bool,
    field_aliases: Vec<(String, String)>,
    fields: Vec<XlsPivotFieldLayout>,
}

impl XlsPivotCacheLayout {
    fn field_index(&self, name: &str) -> Option<usize> {
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

    fn axis_field_index(&self, name: &str) -> Option<usize> {
        self.fields.iter().position(|field| {
            let axis_name = match &field.kind {
                XlsPivotFieldKind::DateSource { .. } | XlsPivotFieldKind::ManualSource { .. } => {
                    return false
                }
                XlsPivotFieldKind::DateGroup {
                    source_field_index, ..
                }
                | XlsPivotFieldKind::ManualGroup {
                    source_field_index, ..
                } => self
                    .fields
                    .get(*source_field_index)
                    .map(|source| source.name.as_str())
                    .unwrap_or(field.name.as_str()),
                _ => field.name.as_str(),
            };
            axis_name.eq_ignore_ascii_case(name)
        })
    }

    fn page_axis_field_index(&self, name: &str) -> Option<usize> {
        self.axis_field_index(name)
    }
}

#[derive(Debug, Clone)]
struct XlsPivotCacheLayouts {
    date_system: DateSystem,
    by_cache_num: HashMap<usize, XlsPivotCacheLayout>,
}

#[derive(Debug)]
struct XlsPivotAxisTuples {
    rows: Vec<Vec<u16>>,
    columns: Vec<Vec<u16>>,
}

impl XlsPivotCacheLayouts {
    fn get(&self, cache_num: usize) -> XlsResult<&XlsPivotCacheLayout> {
        self.by_cache_num.get(&cache_num).ok_or_else(|| {
            XlsError::InvalidFormat(format!("pivot cache layout {cache_num} not found"))
        })
    }
}

#[derive(Debug, Clone)]
struct XlsPivotFieldLayout {
    name: String,
    formula: Option<String>,
    shared_items: Vec<PivotValue>,
    item_ids: Vec<u32>,
    calculated_item_indexes: HashSet<usize>,
    kind: XlsPivotFieldKind,
}

#[derive(Debug, Clone)]
enum XlsPivotFieldKind {
    Regular,
    NumberGroup {
        start: Option<f64>,
        end: Option<f64>,
        interval: f64,
        source_numbers: Vec<f64>,
    },
    DateSource {
        derived_field_indexes: Vec<usize>,
        source_numbers: Vec<f64>,
    },
    DateFilterSource {
        source_numbers: Vec<f64>,
    },
    DateGroup {
        source_field_index: usize,
        unit: PivotDateGroupUnit,
        source_numbers: Vec<f64>,
    },
    ManualSource {
        derived_field_index: usize,
    },
    ManualGroup {
        source_field_index: usize,
        source_item_group_ids: Vec<u32>,
    },
}

#[derive(Debug)]
struct XlsDateSourceData {
    shared_items: Vec<PivotValue>,
    item_ids: Vec<u32>,
    row_numbers: Vec<f64>,
    source_numbers: Vec<f64>,
}

fn sheet_has_pivot_tables(pivot_plan: &FormatPivotPlan, sheet_idx: usize) -> bool {
    pivot_plan
        .tables
        .iter()
        .any(|part| part.sheet_index == sheet_idx)
}

fn build_pivot_cache_streams(
    pivot_plan: &FormatPivotPlan,
    layouts: &XlsPivotCacheLayouts,
) -> XlsResult<(Vec<String>, Vec<(String, Vec<u8>)>)> {
    if pivot_plan.caches.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }

    let storages = vec!["/_SX_DB_CUR".to_string()];
    let mut streams = Vec::with_capacity(pivot_plan.caches.len());
    for cache in &pivot_plan.caches {
        let layout = layouts.get(cache.cache_num)?;
        streams.push((
            format!("/_SX_DB_CUR/{:04}", cache.cache_num),
            build_pivot_cache_stream(cache, layout, layouts.date_system)?,
        ));
    }
    Ok((storages, streams))
}

fn build_xls_pivot_cache_layouts(
    workbook: &Workbook,
    pivot_plan: &FormatPivotPlan,
) -> XlsResult<XlsPivotCacheLayouts> {
    let date_system = workbook_date_system(workbook.settings().date_1904);
    let mut by_cache_num = HashMap::with_capacity(pivot_plan.caches.len());
    for cache in &pivot_plan.caches {
        let groupings = groupings_for_cache(workbook, pivot_plan, cache)?;
        validate_xls_pivot_groupings(cache, groupings)?;
        let grouping_infos = xls_pivot_grouping_infos(cache, groupings, date_system)?;
        let date_filter_fields = date_filter_fields_for_cache(workbook, pivot_plan, cache)?;
        let layout =
            build_xls_pivot_cache_layout(cache, &grouping_infos, &date_filter_fields, date_system)?;
        by_cache_num.insert(cache.cache_num, layout);
    }
    Ok(XlsPivotCacheLayouts {
        date_system,
        by_cache_num,
    })
}

fn build_pivot_cache_stream(
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    date_system: DateSystem,
) -> XlsResult<Vec<u8>> {
    let mut stream = Vec::new();
    let mut sxdb = Vec::new();
    sxdb.extend_from_slice(&(layout.row_count as u32).to_le_bytes());
    sxdb.extend_from_slice(&(layout.cache_num as u16).to_le_bytes());
    sxdb.extend_from_slice(&0x0003u16.to_le_bytes());
    let has_grouped_cache_field = layout.fields.iter().any(|field| {
        matches!(
            field.kind,
            XlsPivotFieldKind::DateGroup { .. } | XlsPivotFieldKind::ManualGroup { .. }
        )
    });
    let has_calculated_field = layout.fields.iter().any(|field| field.formula.is_some());
    let has_calculated_item = !cache.calculated_items.is_empty();
    let has_cache_formula = has_calculated_field || has_calculated_item;
    sxdb.extend_from_slice(&if has_grouped_cache_field || has_calculated_item {
        0x0FFFu16.to_le_bytes()
    } else if has_cache_formula {
        0x0AAAu16.to_le_bytes()
    } else {
        0x1999u16.to_le_bytes()
    });
    sxdb.extend_from_slice(
        &checked_u16(layout.base_field_count, "pivot cache base field count")?.to_le_bytes(),
    );
    sxdb.extend_from_slice(
        &checked_u16(layout.fields.len(), "pivot cache total field count")?.to_le_bytes(),
    );
    let used_row_count = layout.row_count.saturating_add(if has_calculated_item {
        cache.calculated_items.len()
    } else {
        0
    });
    sxdb.extend_from_slice(&if has_grouped_cache_field || has_cache_formula {
        checked_u16(used_row_count, "pivot cache used row count")?.to_le_bytes()
    } else {
        0u16.to_le_bytes()
    });
    sxdb.extend_from_slice(&xls_pivot_cache_source_type(cache)?.to_le_bytes());
    if has_grouped_cache_field || has_cache_formula {
        sxdb.extend_from_slice(&4u16.to_le_bytes());
        sxdb.push(0);
        sxdb.extend_from_slice(b"user");
    } else {
        sxdb.extend_from_slice(&0xFFFFu16.to_le_bytes());
    }
    write_biff_record(&mut stream, records::SXDB, &sxdb);
    if has_calculated_item {
        write_biff_record(
            &mut stream,
            0x0122,
            &[
                0x00, 0x00, 0x00, 0x00, 0x16, 0x8F, 0xE6, 0x40, 0x01, 0x00, 0x00, 0x00,
            ],
        );
    } else {
        write_biff_record(&mut stream, 0x0122, &[0; 12]);
    }

    for item in &cache.calculated_items {
        write_pivot_calculated_item_formula_records(&mut stream, cache, item)?;
    }

    let mut calculated_formulas = Vec::new();
    for (field_index, field) in layout.fields.iter().enumerate() {
        write_sxfdb_record(&mut stream, field_index, field)?;
        if let Some(formula) = &field.formula {
            calculated_formulas.push(formula.clone());
        }
        write_biff_record(&mut stream, 0x01BB, &0u16.to_le_bytes());
        write_pivot_cache_field_items(&mut stream, field, date_system)?;
    }
    for formula in calculated_formulas {
        write_pivot_calculated_field_formula_records(&mut stream, cache, &formula)?;
    }
    write_numeric_cache_records(&mut stream, layout)?;

    write_eof(&mut stream);
    Ok(stream)
}

fn xls_pivot_cache_source_type(cache: &FormatPivotCache) -> XlsResult<u16> {
    match &cache.source {
        FormatPivotSource::Worksheet { .. } => Ok(0x0001),
        FormatPivotSource::External { .. } => unsupported_xls_external_pivot_source(),
        FormatPivotSource::Consolidation { ranges } => {
            validate_xls_consolidation_sources(ranges)?;
            Ok(0x0004)
        }
        FormatPivotSource::Scenario { name } => {
            validate_xls_scenario_source(name)?;
            Ok(0x0008)
        }
        FormatPivotSource::Olap { .. } => unsupported_xls_olap_pivot_source(),
    }
}

fn xls_pivot_view_source_type(cache: &FormatPivotCache) -> XlsResult<u16> {
    match &cache.source {
        FormatPivotSource::Worksheet { .. } => Ok(0x0001),
        FormatPivotSource::External { .. } => unsupported_xls_external_pivot_source(),
        FormatPivotSource::Consolidation { ranges } => {
            validate_xls_consolidation_sources(ranges)?;
            Ok(0x0004)
        }
        // BIFF8 SXVS uses 0x0010 for scenario summaries; SXDB.vsType uses 0x0008.
        FormatPivotSource::Scenario { name } => {
            validate_xls_scenario_source(name)?;
            Ok(0x0010)
        }
        FormatPivotSource::Olap { .. } => unsupported_xls_olap_pivot_source(),
    }
}

fn write_numeric_cache_records(
    stream: &mut Vec<u8>,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    let row_marker_fields = xls_cache_row_marker_fields(layout);
    let numeric_fields = layout
        .fields
        .iter()
        .filter(|field| {
            field.formula.is_none()
                && field_is_numeric(field)
                && matches!(field.kind, XlsPivotFieldKind::Regular)
        })
        .collect::<Vec<_>>();
    if row_marker_fields.is_empty() && numeric_fields.is_empty() {
        return Ok(());
    }

    for row in 0..layout.row_count {
        let row_markers = if row_marker_fields.is_empty() {
            vec![checked_u8(row, "pivot cache row index")?]
        } else {
            row_marker_fields
                .iter()
                .map(|field| {
                    let item_id = field.item_ids.get(row).copied().unwrap_or(0);
                    checked_u8(item_id as usize, "pivot cache row item index")
                })
                .collect::<XlsResult<Vec<_>>>()?
        };
        write_biff_record(stream, PIVOT_SXDBB_RECORD, &row_markers);
        for field in &numeric_fields {
            if let Some(number) = numeric_cache_value(field, row) {
                write_biff_record(stream, PIVOT_SXNUM_RECORD, &number.to_le_bytes());
            }
        }
    }
    Ok(())
}

fn xls_cache_row_marker_fields(layout: &XlsPivotCacheLayout) -> Vec<&XlsPivotFieldLayout> {
    layout
        .fields
        .iter()
        .take(layout.base_field_count)
        .filter(|field| {
            !matches!(field.kind, XlsPivotFieldKind::Regular) || !field_is_numeric(field)
        })
        .collect()
}

fn numeric_cache_value(field: &XlsPivotFieldLayout, row: usize) -> Option<f64> {
    let item_id = *field.item_ids.get(row)? as usize;
    let PivotValue::Number(number) = field.shared_items.get(item_id)? else {
        return None;
    };
    Some(*number)
}

fn groupings_for_cache<'a>(
    workbook: &'a Workbook,
    plan: &'a FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsResult<&'a [PivotGrouping]> {
    let Some(part) = plan
        .tables
        .iter()
        .find(|part| part.cache_num == cache.cache_num)
    else {
        return Ok(&[]);
    };
    let worksheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = worksheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot table not found".into()))?;
    Ok(&pivot.groupings)
}

fn date_filter_fields_for_cache(
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsResult<HashSet<String>> {
    let mut fields = HashSet::new();
    for part in plan
        .tables
        .iter()
        .filter(|part| part.cache_num == cache.cache_num)
    {
        let worksheet = workbook
            .worksheet(part.sheet_index)
            .ok_or_else(|| XlsError::InvalidFormat("pivot table sheet not found".into()))?;
        let pivot = worksheet
            .pivot_tables()
            .get(part.pivot_index)
            .ok_or_else(|| XlsError::InvalidFormat("pivot table not found".into()))?;
        for filter in &pivot.filters {
            match filter {
                PivotFilter::Date { field, .. }
                | PivotFilter::DateBetween { field, .. }
                | PivotFilter::DatePeriod { field, .. } => {
                    fields.insert(field.name.to_lowercase());
                }
                _ => {}
            }
        }
    }
    Ok(fields)
}

fn validate_xls_pivot_groupings(
    cache: &FormatPivotCache,
    groupings: &[PivotGrouping],
) -> XlsResult<()> {
    let mut grouped_fields = HashSet::new();
    for grouping in groupings {
        let field_name = grouping_field_name(grouping);
        if cache.field_index(field_name).is_none() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot grouping references unknown cache field: {field_name}"
            )));
        }
        if !grouped_fields.insert(field_name.to_lowercase()) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot cache has more than one grouping for field {field_name}"
            )));
        }

        match grouping {
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            } => {
                if !interval.is_finite() || *interval <= 0.0 {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot grouping for field {field_name} has an invalid interval"
                    )));
                }
                if start.is_some_and(|value| !value.is_finite())
                    || end.is_some_and(|value| !value.is_finite())
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot grouping for field {field_name} has a non-finite bound"
                    )));
                }
            }
            PivotGrouping::Date { units, .. } => {
                if units.is_empty() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot date grouping has no date units: {field_name}"
                    )));
                }
                let mut seen_units = HashSet::new();
                for unit in units {
                    if !seen_units.insert(*unit) {
                        return Err(XlsError::InvalidFormat(format!(
                            "XLS pivot date grouping for field {field_name} repeats unit {}",
                            xls_date_group_unit_name(*unit)
                        )));
                    }
                }
            }
            PivotGrouping::Manual { groups, .. } => {
                validate_xls_manual_grouping(field_name, groups)?;
            }
        }
    }
    Ok(())
}

fn validate_xls_pivot_grouping_axes(pivot: &duke_sheets_core::PivotTable) -> XlsResult<()> {
    for grouping in &pivot.groupings {
        let PivotGrouping::Manual { field, .. } = grouping else {
            continue;
        };
        let field_name = field.name.as_str();
        let on_rows = pivot_axis_contains_field(&pivot.rows, field_name);
        let on_columns = pivot_axis_contains_field(&pivot.columns, field_name);
        let on_pages = pivot_axis_contains_field(&pivot.page_fields, field_name);
        if on_pages && (on_rows || on_columns) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping does not support field {field_name} on the page axis and another axis"
            )));
        }
        if !on_rows && !on_columns && !on_pages {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping requires a row-, column-, or page-axis field: {field_name}"
            )));
        }
    }
    Ok(())
}

fn pivot_axis_contains_field(fields: &[PivotField], field_name: &str) -> bool {
    fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
}

fn pivot_uses_field_item_filter_axis(
    pivot: &duke_sheets_core::PivotTable,
    field_name: &str,
) -> bool {
    pivot_axis_contains_field(&pivot.rows, field_name)
        || pivot_axis_contains_field(&pivot.columns, field_name)
        || pivot_axis_contains_field(&pivot.page_fields, field_name)
}

fn validate_xls_manual_grouping(field_name: &str, groups: &[PivotManualGroup]) -> XlsResult<()> {
    if groups.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} has no groups"
        )));
    }

    let mut names = HashSet::new();
    let mut members = HashSet::new();
    for group in groups {
        if group.name.trim().is_empty() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping for field {field_name} has a blank group name"
            )));
        }
        if group.members.is_empty() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual group {} has no members",
                group.name
            )));
        }
        if !names.insert(group.name.to_lowercase()) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping for field {field_name} has duplicate group name {}",
                group.name
            )));
        }
        for member in &group.members {
            if !members.insert(member.clone()) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual grouping for field {field_name} assigns item {member} to more than one group"
                )));
            }
        }
    }
    Ok(())
}

fn xls_pivot_grouping_infos<'a>(
    cache: &'a FormatPivotCache,
    groupings: &'a [PivotGrouping],
    date_system: DateSystem,
) -> XlsResult<Vec<XlsPivotGroupingInfo<'a>>> {
    groupings
        .iter()
        .map(|grouping| {
            if let PivotGrouping::Manual { groups, .. } = grouping {
                let field_name = grouping_field_name(grouping);
                let field_index = cache.field_index(field_name).ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot grouping references unknown cache field: {field_name}"
                    ))
                })?;
                let metadata = cache.fields[field_index].grouping.as_ref().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot manual grouping metadata is missing for field {field_name}"
                    ))
                })?;
                let source_items = metadata.source_items.clone();
                let source_item_ids = metadata.source_item_ids.clone();
                let (group_items, base_item_group_ids) =
                    manual_group_items_and_ids(field_name, &source_items, groups)?;
                let group_item_ids = source_item_ids
                    .iter()
                    .map(|item_id| {
                        base_item_group_ids
                            .get(*item_id as usize)
                            .copied()
                            .ok_or_else(|| {
                                XlsError::InvalidFormat(format!(
                                    "XLS pivot manual grouping for field {field_name} has an out-of-range source item index"
                                ))
                            })
                    })
                    .collect::<XlsResult<Vec<_>>>()?;
                return Ok(XlsPivotGroupingInfo {
                    grouping,
                    source_numbers: Vec::new(),
                    source_items,
                    source_item_ids,
                    group_items,
                    base_item_group_ids,
                    group_item_ids,
                });
            }

            Ok(XlsPivotGroupingInfo {
                grouping,
                source_numbers: xls_grouping_source_numbers(
                    cache,
                    grouping,
                    date_system,
                )?,
                source_items: Vec::new(),
                source_item_ids: Vec::new(),
                group_items: Vec::new(),
                base_item_group_ids: Vec::new(),
                group_item_ids: Vec::new(),
            })
        })
        .collect()
}

fn xls_grouping_source_numbers(
    cache: &FormatPivotCache,
    grouping: &PivotGrouping,
    date_system: DateSystem,
) -> XlsResult<Vec<f64>> {
    let field_name = grouping_field_name(grouping);
    let field_index = cache.field_index(field_name).ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot grouping references unknown cache field: {field_name}"
        ))
    })?;
    let metadata = cache.fields[field_index].grouping.as_ref().ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot grouping metadata is missing for field {field_name}"
        ))
    })?;
    let is_date_grouping = matches!(grouping, PivotGrouping::Date { .. });
    let mut seen = HashSet::new();
    let mut numbers = Vec::new();
    for item_id in &metadata.source_item_ids {
        match metadata.source_items.get(*item_id as usize) {
            Some(PivotValue::Number(value)) if value.is_finite() => {
                if is_date_grouping && !valid_xls_date_serial(*value, date_system) {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot date grouping for field {field_name} has invalid date serial: {value}"
                    )));
                }
                if seen.insert(value.to_bits()) {
                    numbers.push(*value);
                }
            }
            Some(PivotValue::Number(value)) if is_date_grouping => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {field_name} has non-finite source value: {value}"
                )));
            }
            Some(PivotValue::Blank) if !is_date_grouping => {}
            _ if is_date_grouping => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {field_name} requires numeric date source values"
                )));
            }
            _ => {}
        }
    }
    if is_date_grouping && numbers.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping for field {field_name} has no source dates"
        )));
    }
    Ok(numbers)
}

fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
    }
}

fn manual_group_items_and_ids(
    field_name: &str,
    source_items: &[PivotValue],
    groups: &[PivotManualGroup],
) -> XlsResult<(Vec<PivotValue>, Vec<u32>)> {
    if source_items.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} has no source items"
        )));
    }
    if source_items
        .iter()
        .any(|item| !xls_manual_group_item_is_supported(item))
    {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} currently supports only text, blank, finite numeric, boolean, or error source items"
        )));
    }

    let mut member_to_group = HashMap::new();
    for group in groups {
        for member in &group.members {
            if !xls_manual_group_item_is_supported(member) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual group {} references an unsupported item in field {field_name}: {member}",
                    group.name
                )));
            }
            if !source_items.iter().any(|item| item == member) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual group {} references item not found in field {field_name}: {member}",
                    group.name
                )));
            }
            member_to_group.insert(member.clone(), group.name.clone());
        }
    }

    let mut group_items = Vec::new();
    let mut ungrouped_item_indexes = HashMap::new();
    for item in source_items {
        if member_to_group.contains_key(item) {
            continue;
        }
        let index = checked_u32(group_items.len(), "pivot manual ungrouped item index")?;
        ungrouped_item_indexes.insert(item.clone(), index);
        group_items.push(item.clone());
    }

    let mut group_name_indexes = HashMap::new();
    for group in groups {
        let index = checked_u32(group_items.len(), "pivot manual group item index")?;
        group_name_indexes.insert(group.name.clone(), index);
        group_items.push(PivotValue::String(group.name.clone()));
    }

    let base_item_group_ids = source_items
        .iter()
        .map(|item| {
            if let Some(group_name) = member_to_group.get(item) {
                group_name_indexes.get(group_name).copied().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot manual grouping for field {field_name} could not map group {group_name}"
                    ))
                })
            } else {
                ungrouped_item_indexes.get(item).copied().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot manual grouping for field {field_name} could not map ungrouped item {item}"
                    ))
                })
            }
        })
        .collect::<XlsResult<Vec<_>>>()?;

    Ok((group_items, base_item_group_ids))
}

fn xls_manual_group_item_is_supported(item: &PivotValue) -> bool {
    matches!(
        item,
        PivotValue::Blank | PivotValue::String(_) | PivotValue::Boolean(_) | PivotValue::Error(_)
    ) || matches!(item, PivotValue::Number(value) if value.is_finite())
}

fn build_xls_pivot_cache_layout(
    cache: &FormatPivotCache,
    grouping_infos: &[XlsPivotGroupingInfo<'_>],
    date_filter_fields: &HashSet<String>,
    date_system: DateSystem,
) -> XlsResult<XlsPivotCacheLayout> {
    let skipped_cache_fields =
        xls_multi_unit_date_group_cache_field_indexes(cache, grouping_infos)?;
    let mut layout = XlsPivotCacheLayout {
        cache_num: cache.cache_num,
        row_count: cache.row_count,
        is_consolidation: matches!(cache.source, FormatPivotSource::Consolidation { .. }),
        field_aliases: cache.field_aliases.clone(),
        base_field_count: cache
            .fields
            .iter()
            .enumerate()
            .filter(|(index, field)| {
                field.formula.is_none() && !skipped_cache_fields.contains(index)
            })
            .count(),
        fields: cache
            .fields
            .iter()
            .enumerate()
            .filter(|(index, _)| !skipped_cache_fields.contains(index))
            .map(|(_, field)| XlsPivotFieldLayout {
                name: field.name.clone(),
                formula: field.formula.clone(),
                shared_items: field.shared_items.clone(),
                item_ids: field.item_ids.clone(),
                calculated_item_indexes: calculated_item_indexes_for_xls_field(
                    cache,
                    &field.name,
                    &field.shared_items,
                ),
                kind: XlsPivotFieldKind::Regular,
            })
            .collect(),
    };

    for info in grouping_infos {
        let field_name = grouping_field_name(info.grouping);
        let cache_field_index = cache.field_index(field_name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot grouping references unknown cache field: {field_name}"
            ))
        })?;
        let field_index = layout.field_index(field_name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot grouping references skipped cache field: {field_name}"
            ))
        })?;
        match info.grouping {
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            } => {
                layout.fields[field_index].kind = XlsPivotFieldKind::NumberGroup {
                    start: *start,
                    end: *end,
                    interval: *interval,
                    source_numbers: info.source_numbers.clone(),
                };
            }
            PivotGrouping::Date { units, .. } => {
                let source = xls_date_source_data(cache, cache_field_index, date_system)?;
                let (start, end) =
                    source_number_min_max(&source.source_numbers).ok_or_else(|| {
                        XlsError::InvalidFormat(format!(
                            "XLS pivot date grouping for field {field_name} has no source dates"
                        ))
                    })?;
                layout.fields[field_index].shared_items = source.shared_items;
                layout.fields[field_index].item_ids = source.item_ids;
                layout.fields[field_index].calculated_item_indexes =
                    calculated_item_indexes_for_xls_field(
                        cache,
                        &layout.fields[field_index].name,
                        &layout.fields[field_index].shared_items,
                    );
                let mut derived_field_indexes = Vec::with_capacity(units.len());
                if units.len() == 1 {
                    let unit = units[0];
                    let derived_items = xls_date_group_shared_items(unit, start, end, date_system)?;
                    let derived_item_ids = source
                        .row_numbers
                        .iter()
                        .map(|serial| {
                            xls_date_group_item_id(unit, *serial, start, end, date_system)
                        })
                        .collect::<XlsResult<Vec<_>>>()?;
                    let derived_field_index = layout.fields.len();
                    derived_field_indexes.push(derived_field_index);
                    let derived_name = unique_xls_date_group_field_name(
                        &layout.fields,
                        &layout.fields[field_index].name,
                        unit,
                    );
                    layout.fields.push(XlsPivotFieldLayout {
                        name: derived_name,
                        formula: None,
                        shared_items: derived_items,
                        item_ids: derived_item_ids,
                        calculated_item_indexes: HashSet::new(),
                        kind: XlsPivotFieldKind::DateGroup {
                            source_field_index: field_index,
                            unit,
                            source_numbers: source.source_numbers.clone(),
                        },
                    });
                } else {
                    let mut indexes_by_unit = HashMap::new();
                    for unit in units.iter().rev() {
                        let derived_items =
                            xls_date_group_shared_items(*unit, start, end, date_system)?;
                        let derived_item_ids = source
                            .row_numbers
                            .iter()
                            .map(|serial| {
                                xls_date_group_item_id(*unit, *serial, start, end, date_system)
                            })
                            .collect::<XlsResult<Vec<_>>>()?;
                        let derived_field_index = layout.fields.len();
                        indexes_by_unit.insert(*unit, derived_field_index);
                        let derived_name = unique_xls_date_group_field_name(
                            &layout.fields,
                            &layout.fields[field_index].name,
                            *unit,
                        );
                        layout.fields.push(XlsPivotFieldLayout {
                            name: derived_name,
                            formula: None,
                            shared_items: derived_items,
                            item_ids: derived_item_ids,
                            calculated_item_indexes: HashSet::new(),
                            kind: XlsPivotFieldKind::DateGroup {
                                source_field_index: field_index,
                                unit: *unit,
                                source_numbers: source.source_numbers.clone(),
                            },
                        });
                    }
                    for unit in units {
                        let derived_field_index =
                            indexes_by_unit.get(unit).copied().ok_or_else(|| {
                                XlsError::InvalidFormat(format!(
                                    "XLS pivot date grouping field missing for {field_name} {}",
                                    xls_date_group_unit_name(*unit)
                                ))
                            })?;
                        derived_field_indexes.push(derived_field_index);
                    }
                }
                layout.fields[field_index].kind = XlsPivotFieldKind::DateSource {
                    derived_field_indexes,
                    source_numbers: source.source_numbers.clone(),
                };
            }
            PivotGrouping::Manual { .. } => {
                let derived_field_index = layout.fields.len();
                layout.fields[field_index].shared_items = info.source_items.clone();
                layout.fields[field_index].item_ids = info.source_item_ids.clone();
                layout.fields[field_index].calculated_item_indexes =
                    calculated_item_indexes_for_xls_field(
                        cache,
                        &layout.fields[field_index].name,
                        &layout.fields[field_index].shared_items,
                    );
                layout.fields[field_index].kind = XlsPivotFieldKind::ManualSource {
                    derived_field_index,
                };
                let derived_name = unique_xls_manual_grouped_field_name(
                    &layout.fields,
                    &layout.fields[field_index].name,
                );
                layout.fields.push(XlsPivotFieldLayout {
                    name: derived_name,
                    formula: None,
                    shared_items: info.group_items.clone(),
                    item_ids: info.group_item_ids.clone(),
                    calculated_item_indexes: HashSet::new(),
                    kind: XlsPivotFieldKind::ManualGroup {
                        source_field_index: field_index,
                        source_item_group_ids: info.base_item_group_ids.clone(),
                    },
                });
            }
        }
    }

    for field_name in date_filter_fields {
        let Some(cache_field_index) = cache
            .fields
            .iter()
            .position(|field| field.name.eq_ignore_ascii_case(field_name))
        else {
            continue;
        };
        let Some(field_index) = layout.field_index(&cache.fields[cache_field_index].name) else {
            continue;
        };
        if !matches!(layout.fields[field_index].kind, XlsPivotFieldKind::Regular) {
            continue;
        }
        let source = xls_date_source_data(cache, cache_field_index, date_system)?;
        layout.fields[field_index].shared_items = source.shared_items;
        layout.fields[field_index].item_ids = source.item_ids;
        layout.fields[field_index].calculated_item_indexes = calculated_item_indexes_for_xls_field(
            cache,
            &layout.fields[field_index].name,
            &layout.fields[field_index].shared_items,
        );
        layout.fields[field_index].kind = XlsPivotFieldKind::DateFilterSource {
            source_numbers: source.source_numbers,
        };
    }

    Ok(layout)
}

fn xls_multi_unit_date_group_cache_field_indexes(
    cache: &FormatPivotCache,
    grouping_infos: &[XlsPivotGroupingInfo<'_>],
) -> XlsResult<HashSet<usize>> {
    let source_field_count = xls_pivot_source_field_count(cache)?;
    let mut claimed = HashSet::new();
    for info in grouping_infos {
        let PivotGrouping::Date { field, units } = info.grouping else {
            continue;
        };
        if units.len() <= 1 {
            continue;
        }

        for unit in units {
            let index = find_xls_multi_unit_date_group_cache_field_index(
                cache,
                &field.name,
                *unit,
                source_field_count,
                &claimed,
            )
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "XLS pivot multi-unit date grouping could not find transformed cache field {} ({})",
                    field.name,
                    xls_date_group_unit_name(*unit)
                ))
            })?;
            claimed.insert(index);
        }
    }
    Ok(claimed)
}

fn xls_pivot_source_field_count(cache: &FormatPivotCache) -> XlsResult<usize> {
    match &cache.source {
        FormatPivotSource::Worksheet { range, .. } => {
            let field_count = u32::from(range.end.col)
                .saturating_sub(u32::from(range.start.col))
                .saturating_add(1);
            usize::try_from(field_count).map_err(|_| {
                XlsError::InvalidFormat("pivot source field count exceeds usize".into())
            })
        }
        FormatPivotSource::Consolidation { .. } => Ok(cache
            .fields
            .iter()
            .filter(|field| field.database_field)
            .count()),
        FormatPivotSource::External { .. } => unsupported_xls_external_pivot_source(),
        FormatPivotSource::Olap { .. } => unsupported_xls_olap_pivot_source(),
        FormatPivotSource::Scenario { name } => {
            validate_xls_scenario_source(name)?;
            Ok(cache
                .fields
                .iter()
                .filter(|field| field.database_field)
                .count())
        }
    }
}

fn unique_xls_manual_grouped_field_name(
    fields: &[XlsPivotFieldLayout],
    source_name: &str,
) -> String {
    for suffix in 2usize.. {
        let candidate = format!("{source_name}{suffix}");
        if fields
            .iter()
            .all(|field| !field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded manual grouped field name suffix search should return")
}

fn xls_date_source_data(
    cache: &FormatPivotCache,
    field_index: usize,
    date_system: DateSystem,
) -> XlsResult<XlsDateSourceData> {
    let field = cache.fields.get(field_index).ok_or_else(|| {
        XlsError::InvalidFormat("XLS pivot date field index is out of range".into())
    })?;
    let (shared_items, item_ids) = field
        .grouping
        .as_ref()
        .map(|grouping| {
            (
                grouping.source_items.clone(),
                grouping.source_item_ids.clone(),
            )
        })
        .unwrap_or_else(|| (field.shared_items.clone(), field.item_ids.clone()));
    let mut row_numbers = Vec::new();
    for item_id in &item_ids {
        let value = match shared_items.get(*item_id as usize) {
            Some(PivotValue::Number(value)) if value.is_finite() => *value,
            Some(PivotValue::Number(value)) => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping has non-finite source value: {value}"
                )));
            }
            _ => {
                return Err(XlsError::InvalidFormat(
                    "XLS pivot date grouping requires numeric date source values".into(),
                ));
            }
        };
        if !valid_xls_date_serial(value, date_system) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot date grouping has invalid date serial: {value}"
            )));
        }
        row_numbers.push(value);
    }
    if row_numbers.len() != cache.row_count {
        return Err(XlsError::InvalidFormat(
            "XLS pivot date source row count does not match the cache".into(),
        ));
    }
    let source_numbers = shared_items
        .iter()
        .filter_map(|value| match value {
            PivotValue::Number(number) => Some(*number),
            _ => None,
        })
        .collect();
    Ok(XlsDateSourceData {
        shared_items,
        item_ids,
        row_numbers,
        source_numbers,
    })
}

fn unique_xls_date_group_field_name(
    fields: &[XlsPivotFieldLayout],
    source_name: &str,
    unit: PivotDateGroupUnit,
) -> String {
    let base = format!("{} ({source_name})", xls_date_group_unit_name(unit));
    if !fields
        .iter()
        .any(|field| field.name.eq_ignore_ascii_case(&base))
    {
        return base;
    }
    for suffix in 2.. {
        let candidate = format!("{base} {suffix}");
        if !fields
            .iter()
            .any(|field| field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded pivot date group name suffix search should return")
}

fn find_xls_multi_unit_date_group_cache_field_index(
    cache: &FormatPivotCache,
    source_name: &str,
    unit: PivotDateGroupUnit,
    start_index: usize,
    claimed: &HashSet<usize>,
) -> Option<usize> {
    let base = format!("{source_name} ({})", xls_date_group_unit_name(unit));
    (1usize..).find_map(|suffix| {
        let candidate = if suffix == 1 {
            base.clone()
        } else {
            format!("{base} {suffix}")
        };
        cache
            .fields
            .iter()
            .enumerate()
            .skip(start_index)
            .find(|(index, field)| {
                !claimed.contains(index)
                    && field.formula.is_none()
                    && field.name.eq_ignore_ascii_case(&candidate)
            })
            .map(|(index, _)| index)
    })
}

fn write_sxfdb_record(
    stream: &mut Vec<u8>,
    field_index: usize,
    field: &XlsPivotFieldLayout,
) -> XlsResult<()> {
    let mut body = Vec::new();
    let mut flags = match field.kind {
        XlsPivotFieldKind::NumberGroup { .. } => 0x0571u16,
        XlsPivotFieldKind::DateSource { .. } => 0x0909u16,
        XlsPivotFieldKind::DateFilterSource { .. } => 0x0901u16,
        XlsPivotFieldKind::DateGroup { .. } => 0x0011u16,
        XlsPivotFieldKind::ManualSource { .. } if field_is_numeric(field) => 0x0569u16,
        XlsPivotFieldKind::ManualSource { .. } => 0x0489u16,
        XlsPivotFieldKind::ManualGroup { .. } => 0x0001u16,
        XlsPivotFieldKind::Regular if field_is_numeric(field) => 0x0560u16,
        XlsPivotFieldKind::Regular => 0x0481u16,
    };
    if field.formula.is_some() {
        flags = 0x8425;
    }
    body.extend_from_slice(&flags.to_le_bytes());
    match &field.kind {
        XlsPivotFieldKind::NumberGroup { source_numbers, .. } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count = checked_u16(field.shared_items.len(), "pivot group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(source_numbers.len(), "pivot grouped source atom count")?
                    .to_le_bytes(),
            );
        }
        XlsPivotFieldKind::DateSource {
            derived_field_indexes,
            source_numbers,
        } => {
            let first_derived = derived_field_indexes.first().copied().ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "XLS pivot date source field {} has no derived fields",
                    field.name
                ))
            })?;
            body.extend_from_slice(
                &checked_i16(first_derived, "pivot date group field index")?.to_le_bytes(),
            );
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count = checked_u16(field.shared_items.len(), "pivot date item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(source_numbers.len(), "pivot date source atom count")?.to_le_bytes(),
            );
        }
        XlsPivotFieldKind::DateFilterSource { source_numbers } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count = checked_u16(field.shared_items.len(), "pivot date item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(source_numbers.len(), "pivot date source atom count")?.to_le_bytes(),
            );
        }
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(
                &checked_i16(*source_field_index, "pivot date source field index")?.to_le_bytes(),
            );
            let item_count = checked_u16(field.shared_items.len(), "pivot date group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
        }
        XlsPivotFieldKind::ManualSource {
            derived_field_index,
        } => {
            body.extend_from_slice(
                &checked_i16(*derived_field_index, "pivot manual group field index")?.to_le_bytes(),
            );
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count =
                checked_u16(field.shared_items.len(), "pivot manual source item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
        }
        XlsPivotFieldKind::ManualGroup {
            source_field_index,
            source_item_group_ids,
        } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(
                &checked_i16(*source_field_index, "pivot manual source field index")?.to_le_bytes(),
            );
            let item_count =
                checked_u16(field.shared_items.len(), "pivot manual group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(
                    source_item_group_ids.len(),
                    "pivot manual grouped source atom count",
                )?
                .to_le_bytes(),
            );
            body.extend_from_slice(&0u16.to_le_bytes());
        }
        XlsPivotFieldKind::Regular if field.formula.is_some() => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(
                &checked_i16(field_index, "pivot calculated field index")?.to_le_bytes(),
            );
            body.extend_from_slice(&2u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
        }
        XlsPivotFieldKind::Regular => {
            let item_count = checked_u16(field.shared_items.len(), "pivot field item count")?;
            body.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
            body.extend_from_slice(
                &checked_u32(field.shared_items.len(), "pivot field item count")?.to_le_bytes(),
            );
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &if field_is_numeric(field) {
                    0u16
                } else {
                    item_count
                }
                .to_le_bytes(),
            );
        }
    }
    push_xlunicode_string(&mut body, &field.name)?;
    write_biff_record(stream, PIVOT_SXFDB_RECORD, &body);
    Ok(())
}

fn write_pivot_cache_field_items(
    stream: &mut Vec<u8>,
    field: &XlsPivotFieldLayout,
    date_system: DateSystem,
) -> XlsResult<()> {
    match &field.kind {
        XlsPivotFieldKind::NumberGroup {
            start,
            end,
            interval,
            source_numbers,
        } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }

            let mut flags = 0x0040u16;
            if start.is_none() {
                flags |= 0x0001;
            }
            if end.is_none() {
                flags |= 0x0002;
            }
            write_biff_record(stream, PIVOT_SXRNG_RECORD, &flags.to_le_bytes());
            write_biff_record(
                stream,
                PIVOT_SXNUM_RECORD,
                &start.unwrap_or(0.0).to_le_bytes(),
            );
            write_biff_record(
                stream,
                PIVOT_SXNUM_RECORD,
                &end.unwrap_or(0.0).to_le_bytes(),
            );
            write_biff_record(stream, PIVOT_SXNUM_RECORD, &interval.to_le_bytes());
            for number in source_numbers {
                write_biff_record(stream, PIVOT_SXNUM_RECORD, &number.to_le_bytes());
            }
        }
        XlsPivotFieldKind::DateSource { source_numbers, .. }
        | XlsPivotFieldKind::DateFilterSource { source_numbers } => {
            for number in source_numbers {
                write_pivot_sxdtr_record(stream, *number, date_system)?;
            }
        }
        XlsPivotFieldKind::DateGroup {
            unit,
            source_numbers,
            ..
        } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }
            let flags = 0x0003u16 | (u16::from(xls_date_group_by(*unit)) << 2);
            write_biff_record(stream, PIVOT_SXRNG_RECORD, &flags.to_le_bytes());
            let (start, end) = source_number_min_max(source_numbers).ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {} has no source dates",
                    field.name
                ))
            })?;
            write_pivot_sxdtr_record(stream, start, date_system)?;
            write_pivot_sxdtr_record(stream, end, date_system)?;
            write_biff_record(stream, PIVOT_SXINT_RECORD, &1u16.to_le_bytes());
        }
        XlsPivotFieldKind::ManualSource { .. } => {
            for value in &field.shared_items {
                write_pivot_cache_shared_item(stream, value)?;
            }
        }
        XlsPivotFieldKind::ManualGroup {
            source_item_group_ids,
            ..
        } => {
            for value in &field.shared_items {
                write_pivot_cache_shared_item(stream, value)?;
            }
            let mut body = Vec::new();
            for item_id in source_item_group_ids {
                body.extend_from_slice(
                    &checked_u16(*item_id as usize, "pivot manual source item group index")?
                        .to_le_bytes(),
                );
            }
            write_biff_record(stream, PIVOT_SXIDSTM_RECORD, &body);
        }
        XlsPivotFieldKind::Regular if !field_is_numeric(field) => {
            for value in &field.shared_items {
                write_pivot_cache_shared_item(stream, value)?;
            }
        }
        XlsPivotFieldKind::Regular if field.item_ids.is_empty() => {
            for value in &field.shared_items {
                write_pivot_cache_shared_item(stream, value)?;
            }
        }
        XlsPivotFieldKind::Regular => {}
    }
    Ok(())
}

fn write_pivot_cache_shared_item(stream: &mut Vec<u8>, value: &PivotValue) -> XlsResult<()> {
    match value {
        PivotValue::Number(value) if value.is_finite() => {
            write_biff_record(stream, PIVOT_SXNUM_RECORD, &value.to_le_bytes());
        }
        PivotValue::Number(value) => {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot cache item has a non-finite number: {value}"
            )));
        }
        PivotValue::Blank | PivotValue::String(_) => {
            write_biff_record(
                stream,
                PIVOT_SXSTRING_RECORD,
                &pivot_value_string_payload(value)?,
            );
        }
        PivotValue::Boolean(value) => {
            write_biff_record(stream, PIVOT_SXBOOL_RECORD, &[u8::from(*value), 0]);
        }
        PivotValue::Error(value) => {
            write_biff_record(stream, PIVOT_SXERR_RECORD, &[value.code(), 0]);
        }
    }
    Ok(())
}

fn xls_date_group_shared_items(
    unit: PivotDateGroupUnit,
    start: f64,
    end: f64,
    date_system: DateSystem,
) -> XlsResult<Vec<PivotValue>> {
    let mut items = Vec::new();
    items.push(PivotValue::String(format!(
        "<{}",
        format_xls_pivot_date_bound(start, date_system)?
    )));
    match unit {
        PivotDateGroupUnit::Seconds | PivotDateGroupUnit::Minutes => {
            for value in 0..=59 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Hours => {
            for value in 0..=23 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Days => {
            for value in 1..=31 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Months => {
            for label in [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ] {
                items.push(PivotValue::String(label.to_string()));
            }
        }
        PivotDateGroupUnit::Quarters => {
            for value in 1..=4 {
                items.push(PivotValue::String(format!("Qtr{value}")));
            }
        }
        PivotDateGroupUnit::Years => {
            let (start_year, _, _) = xls_serial_date_tuple(start, date_system)?;
            let (end_year, _, _) = xls_serial_date_tuple(end, date_system)?;
            for year in start_year..=end_year {
                items.push(PivotValue::String(year.to_string()));
            }
        }
    }
    items.push(PivotValue::String(format!(
        ">{}",
        format_xls_pivot_date_bound(end, date_system)?
    )));
    Ok(items)
}

fn xls_date_group_item_id(
    unit: PivotDateGroupUnit,
    serial: f64,
    start: f64,
    end: f64,
    date_system: DateSystem,
) -> XlsResult<u32> {
    if serial < start {
        return Ok(0);
    }
    let (year, month, day) = xls_serial_date_tuple(serial, date_system)?;
    let (hour, minute, second) = serial_to_time(serial);
    let index = match unit {
        PivotDateGroupUnit::Seconds => 1 + u32::from(second),
        PivotDateGroupUnit::Minutes => 1 + u32::from(minute),
        PivotDateGroupUnit::Hours => 1 + u32::from(hour),
        PivotDateGroupUnit::Days => u32::from(day),
        PivotDateGroupUnit::Months => u32::from(month),
        PivotDateGroupUnit::Quarters => u32::from((month - 1) / 3 + 1),
        PivotDateGroupUnit::Years => {
            let (start_year, _, _) = xls_serial_date_tuple(start, date_system)?;
            1 + (year - start_year) as u32
        }
    };
    if serial > end {
        Ok(index.saturating_add(1))
    } else {
        Ok(index)
    }
}

fn xls_date_group_unit_name(unit: PivotDateGroupUnit) -> &'static str {
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

fn format_xls_pivot_date_bound(serial: f64, date_system: DateSystem) -> XlsResult<String> {
    let (year, month, day) = xls_serial_date_tuple(serial, date_system)?;
    Ok(format!("{month}/{day}/{year}"))
}

fn xls_serial_date_tuple(serial: f64, date_system: DateSystem) -> XlsResult<(i32, u32, u32)> {
    serial_to_date(serial, date_system).ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot date grouping has invalid date serial: {serial}"
        ))
    })
}

fn xls_date_group_by(unit: PivotDateGroupUnit) -> u8 {
    match unit {
        PivotDateGroupUnit::Seconds => 0x01,
        PivotDateGroupUnit::Minutes => 0x02,
        PivotDateGroupUnit::Hours => 0x03,
        PivotDateGroupUnit::Days => 0x04,
        PivotDateGroupUnit::Months => 0x05,
        PivotDateGroupUnit::Quarters => 0x06,
        PivotDateGroupUnit::Years => 0x07,
    }
}

fn write_pivot_sxdtr_record(
    stream: &mut Vec<u8>,
    serial: f64,
    date_system: DateSystem,
) -> XlsResult<()> {
    let Some((year, month, day)) = serial_to_date(serial, date_system) else {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping has invalid date serial: {serial}"
        )));
    };
    if !(1900..=9999).contains(&year) {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping date year is out of range: {year}"
        )));
    }
    let (hour, minute, second) = serial_to_time(serial);
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&(year as u16).to_le_bytes());
    payload.extend_from_slice(&(month as u16).to_le_bytes());
    payload.push(day as u8);
    payload.push(hour as u8);
    payload.push(minute as u8);
    payload.push(second as u8);
    write_biff_record(stream, PIVOT_SXDTR_RECORD, &payload);
    Ok(())
}

fn source_number_min_max(numbers: &[f64]) -> Option<(f64, f64)> {
    let mut min = None::<f64>;
    let mut max = None::<f64>;
    for number in numbers {
        min = Some(min.map_or(*number, |current| current.min(*number)));
        max = Some(max.map_or(*number, |current| current.max(*number)));
    }
    min.zip(max)
}

fn valid_xls_date_serial(serial: f64, date_system: DateSystem) -> bool {
    serial.is_finite()
        && serial_to_date(serial, date_system).is_some_and(|(year, month, day)| {
            (1900..=9999).contains(&year) && (1..=12).contains(&month) && day <= 31
        })
}

fn workbook_date_system(date_1904: bool) -> DateSystem {
    if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    }
}

fn write_pivot_calculated_field_formula_records(
    stream: &mut Vec<u8>,
    cache: &FormatPivotCache,
    formula: &str,
) -> XlsResult<()> {
    let trimmed = formula.trim();
    let parse_input = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={trimmed}")
    };
    let expr = duke_sheets_formula::parse_formula(&parse_input).map_err(|err| {
        XlsError::InvalidFormat(format!(
            "XLS pivot calculated field formula could not be parsed: {err}"
        ))
    })?;

    let mut ptgs = Vec::new();
    let mut sx_names = Vec::new();
    compile_pivot_calculated_formula_expr(&expr, &mut ptgs, &mut sx_names, cache).map_err(
        |_| {
            XlsError::InvalidFormat(format!(
                "XLS pivot calculated field formula uses unsupported syntax: {formula}"
            ))
        },
    )?;

    let mut body = Vec::new();
    body.extend_from_slice(
        &checked_u16(ptgs.len(), "pivot calculated formula token length")?.to_le_bytes(),
    );
    body.extend_from_slice(
        &checked_u16(
            sx_names.len(),
            "pivot calculated formula field reference count",
        )?
        .to_le_bytes(),
    );
    body.extend_from_slice(&ptgs);
    write_biff_record(stream, PIVOT_SXFMLA_RECORD, &body);

    for field_index in sx_names {
        write_pivot_sxname_record(stream, field_index)?;
    }

    Ok(())
}

fn write_pivot_sxname_record(stream: &mut Vec<u8>, field_index: usize) -> XlsResult<()> {
    let field_index = checked_i16(field_index, "pivot calculated formula source field index")?;
    let mut body = Vec::with_capacity(8);
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&field_index.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    write_biff_record(stream, PIVOT_SXNAME_RECORD, &body);
    Ok(())
}

fn write_pivot_calculated_item_formula_records(
    stream: &mut Vec<u8>,
    cache: &FormatPivotCache,
    item: &PivotCalculatedItem,
) -> XlsResult<()> {
    let field_index = cache.field_index(&item.field.name).ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot calculated item references unknown field {}",
            item.field.name
        ))
    })?;
    let field = cache.fields.get(field_index).ok_or_else(|| {
        XlsError::InvalidFormat("XLS pivot calculated item field index is invalid".into())
    })?;
    let target_item_index = field
        .shared_items
        .iter()
        .position(|candidate| candidate == &item.item)
        .ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot calculated item target for field {} is not present in the cache",
                item.field.name
            ))
        })?;
    let calculated_item_indexes = cache
        .calculated_items
        .iter()
        .filter(|calculated| calculated.field.name == item.field.name)
        .filter_map(|calculated| {
            field
                .shared_items
                .iter()
                .position(|candidate| candidate == &calculated.item)
        })
        .collect::<HashSet<_>>();

    let trimmed = item.formula.trim();
    let parse_input = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={trimmed}")
    };
    let expr = duke_sheets_formula::parse_formula(&parse_input).map_err(|err| {
        XlsError::InvalidFormat(format!(
            "XLS pivot calculated item formula could not be parsed: {err}"
        ))
    })?;

    let mut ptgs = Vec::new();
    let mut sx_names = Vec::new();
    compile_pivot_calculated_item_formula_expr(
        &expr,
        &mut ptgs,
        &mut sx_names,
        &field.shared_items,
    )
    .map_err(|_| {
        XlsError::InvalidFormat(format!(
            "XLS pivot calculated item formula uses unsupported syntax: {}",
            item.formula
        ))
    })?;
    if sx_names
        .iter()
        .any(|item_index| calculated_item_indexes.contains(item_index))
    {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot calculated item formula references another calculated item, which Excel does not preserve: {}",
            item.formula
        )));
    }

    let mut body = Vec::new();
    body.extend_from_slice(
        &checked_u16(ptgs.len(), "pivot calculated item formula token length")?.to_le_bytes(),
    );
    body.extend_from_slice(
        &checked_u16(
            sx_names.len(),
            "pivot calculated item formula reference count",
        )?
        .to_le_bytes(),
    );
    body.extend_from_slice(&ptgs);
    write_biff_record(stream, PIVOT_SXFMLA_RECORD, &body);

    for item_index in sx_names {
        write_pivot_sxname_item_record(stream)?;
        write_pivot_sxpair_record(stream, field_index, item_index)?;
    }
    write_biff_record(
        stream,
        0x00F0,
        &[0x00, 0xFF, 0x10, 0x42, 0x00, 0x00, 0x01, 0x00],
    );
    write_biff_record(
        stream,
        0x00F2,
        &[0x01, 0x00, 0x00, 0x04, 0x01, 0x00, 0x01, 0x00],
    );
    write_biff_record(
        stream,
        0x00F5,
        &checked_u16(target_item_index, "pivot calculated item target index")?.to_le_bytes(),
    );
    write_pivot_sxformula_record(stream, target_item_index)?;

    Ok(())
}

fn write_pivot_sxname_item_record(stream: &mut Vec<u8>) -> XlsResult<()> {
    let mut body = Vec::with_capacity(8);
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&(-1i16).to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes());
    write_biff_record(stream, PIVOT_SXNAME_RECORD, &body);
    Ok(())
}

fn write_pivot_sxpair_record(
    stream: &mut Vec<u8>,
    field_index: usize,
    item_index: usize,
) -> XlsResult<()> {
    let mut body = Vec::with_capacity(8);
    body.extend_from_slice(
        &checked_u16(field_index, "pivot calculated item field reference index")?.to_le_bytes(),
    );
    body.extend_from_slice(
        &checked_u16(item_index, "pivot calculated item reference index")?.to_le_bytes(),
    );
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    write_biff_record(stream, PIVOT_SXPAIR_RECORD, &body);
    Ok(())
}

fn write_pivot_sxformula_record(stream: &mut Vec<u8>, _item_index: usize) -> XlsResult<()> {
    write_biff_record(stream, PIVOT_SXFORMULA_RECORD, &[0x00, 0x00, 0xFF, 0xFF]);
    Ok(())
}

fn compile_pivot_calculated_formula_expr(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    cache: &FormatPivotCache,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Number(n) => {
            if let Some(i) = number_as_ptg_int(*n) {
                out.push(0x1E);
                out.extend_from_slice(&i.to_le_bytes());
            } else {
                out.push(0x1F);
                out.extend_from_slice(&n.to_le_bytes());
            }
        }
        FormulaExpr::String(s) => {
            out.push(0x17);
            push_short_xlunicode_string(out, s).map_err(|_| UnsupportedToken)?;
        }
        FormulaExpr::Boolean(b) => {
            out.push(0x1D);
            out.push(if *b { 1 } else { 0 });
        }
        FormulaExpr::Error(e) => {
            out.push(0x1C);
            out.push(e.code());
        }
        FormulaExpr::NameRef(name) => {
            push_pivot_sxname_ptg(out, sx_names, cache, name)?;
        }
        FormulaExpr::StructuredRef(reference) => {
            let Some(column) = reference.column.as_deref() else {
                return Err(UnsupportedToken);
            };
            push_pivot_sxname_ptg(out, sx_names, cache, column)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_pivot_calculated_formula_expr(left, out, sx_names, cache)?;
            compile_pivot_calculated_formula_expr(right, out, sx_names, cache)?;
            out.push(match op {
                BinaryOperator::Add => 0x03,
                BinaryOperator::Subtract => 0x04,
                BinaryOperator::Multiply => 0x05,
                BinaryOperator::Divide => 0x06,
                BinaryOperator::Power => 0x07,
                BinaryOperator::Concat => 0x08,
                BinaryOperator::LessThan => 0x09,
                BinaryOperator::LessEqual => 0x0A,
                BinaryOperator::Equal => 0x0B,
                BinaryOperator::GreaterEqual => 0x0C,
                BinaryOperator::GreaterThan => 0x0D,
                BinaryOperator::NotEqual => 0x0E,
                BinaryOperator::Intersect | BinaryOperator::Union | BinaryOperator::Range => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::UnaryOp { op, operand } => {
            compile_pivot_calculated_formula_expr(operand, out, sx_names, cache)?;
            out.push(match op {
                UnaryOperator::Plus => 0x12,
                UnaryOperator::Negate => 0x13,
                UnaryOperator::Percent => 0x14,
                UnaryOperator::Paren => 0x15,
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::Function { name, args } => {
            compile_pivot_calculated_function_expr(name, args, out, |arg, out| {
                compile_pivot_calculated_formula_expr(arg, out, sx_names, cache)
            })?;
        }
        FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => return Err(UnsupportedToken),
    }

    Ok(())
}

fn compile_pivot_calculated_item_formula_expr(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    shared_items: &[PivotValue],
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Number(n) => {
            if let Some(i) = number_as_ptg_int(*n) {
                out.push(0x1E);
                out.extend_from_slice(&i.to_le_bytes());
            } else {
                out.push(0x1F);
                out.extend_from_slice(&n.to_le_bytes());
            }
        }
        FormulaExpr::String(s) => {
            if !try_push_pivot_calculated_item_sxname_ptg(out, sx_names, shared_items, s)? {
                out.push(0x17);
                push_short_xlunicode_string(out, s).map_err(|_| UnsupportedToken)?;
            }
        }
        FormulaExpr::Boolean(b) => {
            out.push(0x1D);
            out.push(if *b { 1 } else { 0 });
        }
        FormulaExpr::Error(e) => {
            out.push(0x1C);
            out.push(e.code());
        }
        FormulaExpr::NameRef(name) => {
            push_pivot_calculated_item_sxname_ptg(out, sx_names, shared_items, name)?;
        }
        FormulaExpr::StructuredRef(reference) => {
            let Some(column) = reference.column.as_deref() else {
                return Err(UnsupportedToken);
            };
            push_pivot_calculated_item_sxname_ptg(out, sx_names, shared_items, column)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_pivot_calculated_item_formula_expr(left, out, sx_names, shared_items)?;
            compile_pivot_calculated_item_formula_expr(right, out, sx_names, shared_items)?;
            out.push(match op {
                BinaryOperator::Add => 0x03,
                BinaryOperator::Subtract => 0x04,
                BinaryOperator::Multiply => 0x05,
                BinaryOperator::Divide => 0x06,
                BinaryOperator::Power => 0x07,
                BinaryOperator::Concat => 0x08,
                BinaryOperator::LessThan => 0x09,
                BinaryOperator::LessEqual => 0x0A,
                BinaryOperator::Equal => 0x0B,
                BinaryOperator::GreaterEqual => 0x0C,
                BinaryOperator::GreaterThan => 0x0D,
                BinaryOperator::NotEqual => 0x0E,
                BinaryOperator::Intersect | BinaryOperator::Union | BinaryOperator::Range => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::UnaryOp { op, operand } => {
            compile_pivot_calculated_item_formula_expr(operand, out, sx_names, shared_items)?;
            out.push(match op {
                UnaryOperator::Plus => 0x12,
                UnaryOperator::Negate => 0x13,
                UnaryOperator::Percent => 0x14,
                UnaryOperator::Paren => 0x15,
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::Function { name, args } => {
            compile_pivot_calculated_function_expr(name, args, out, |arg, out| {
                compile_pivot_calculated_item_formula_expr(arg, out, sx_names, shared_items)
            })?;
        }
        FormulaExpr::CellRef(reference) => {
            let Some(name) = pivot_calculated_item_cell_ref_name(reference) else {
                return Err(UnsupportedToken);
            };
            push_pivot_calculated_item_sxname_ptg(out, sx_names, shared_items, &name)?;
        }
        FormulaExpr::RangeRef(_)
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => return Err(UnsupportedToken),
    }

    Ok(())
}

fn compile_pivot_calculated_function_expr<F>(
    name: &str,
    args: &[duke_sheets_formula::FormulaExpr],
    out: &mut Vec<u8>,
    mut compile_arg: F,
) -> Result<(), UnsupportedToken>
where
    F: FnMut(&duke_sheets_formula::FormulaExpr, &mut Vec<u8>) -> Result<(), UnsupportedToken>,
{
    let Some(idx) = function_index(name) else {
        return Err(UnsupportedToken);
    };
    if function_is_biff8_addin(idx) || args.len() > 0x7F {
        return Err(UnsupportedToken);
    }
    for arg in args {
        if matches!(arg, duke_sheets_formula::FormulaExpr::Empty) {
            return Err(UnsupportedToken);
        }
        compile_arg(arg, out)?;
    }
    if function_is_fixed_arity(idx, args.len()) {
        out.push(ptg_func_opcode(OperandClass::V));
        out.extend_from_slice(&idx.to_le_bytes());
    } else {
        out.push(ptg_func_var_opcode(OperandClass::V));
        out.push(args.len() as u8);
        out.extend_from_slice(&idx.to_le_bytes());
    }
    Ok(())
}

fn push_pivot_calculated_item_sxname_ptg(
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    shared_items: &[PivotValue],
    item_name: &str,
) -> Result<(), UnsupportedToken> {
    if try_push_pivot_calculated_item_sxname_ptg(out, sx_names, shared_items, item_name)? {
        Ok(())
    } else {
        Err(UnsupportedToken)
    }
}

fn try_push_pivot_calculated_item_sxname_ptg(
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    shared_items: &[PivotValue],
    item_name: &str,
) -> Result<bool, UnsupportedToken> {
    let Some(item_index) = shared_items.iter().position(
        |item| matches!(item, PivotValue::String(text) if text.eq_ignore_ascii_case(item_name)),
    ) else {
        return Ok(false);
    };
    let sx_name_index = checked_u32(sx_names.len(), "pivot calculated item formula SXNAME index")
        .map_err(|_| UnsupportedToken)?;
    sx_names.push(item_index);
    out.extend_from_slice(&[0x18, 0x1D]);
    out.extend_from_slice(&sx_name_index.to_le_bytes());
    Ok(true)
}

fn pivot_calculated_item_cell_ref_name(
    reference: &duke_sheets_formula::ast::CellReference,
) -> Option<String> {
    reference
        .sheet
        .is_none()
        .then(|| reference.address.to_a1_string())
}

fn push_pivot_sxname_ptg(
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    cache: &FormatPivotCache,
    field_name: &str,
) -> Result<(), UnsupportedToken> {
    let field_index = cache.field_index(field_name).ok_or(UnsupportedToken)?;
    let sx_name_index = checked_u32(sx_names.len(), "pivot calculated formula SXNAME index")
        .map_err(|_| UnsupportedToken)?;
    sx_names.push(field_index);
    out.extend_from_slice(&[0x18, 0x1D]);
    out.extend_from_slice(&sx_name_index.to_le_bytes());
    Ok(())
}

fn write_pivot_global_records(stream: &mut Vec<u8>, pivot_plan: &FormatPivotPlan) {
    if pivot_plan.caches.is_empty() {
        return;
    }

    write_biff_record(
        stream,
        0x089A,
        &[
            0x9A, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00,
        ],
    );
    write_biff_record(
        stream,
        0x08A3,
        &[
            0xA3, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
        ],
    );
}

fn write_table_styles_record(stream: &mut Vec<u8>, pivot_plan: &FormatPivotPlan) {
    if pivot_plan.caches.is_empty() {
        return;
    }

    let default_table_style = "TableStyleMedium2";
    let default_pivot_style = "PivotStyleMedium9";
    let mut body = Vec::with_capacity(88);
    body.extend_from_slice(&TABLESTYLES_RECORD.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&[0; 8]);
    body.extend_from_slice(&0u32.to_le_bytes());
    body.extend_from_slice(&(default_table_style.encode_utf16().count() as u16).to_le_bytes());
    body.extend_from_slice(&(default_pivot_style.encode_utf16().count() as u16).to_le_bytes());
    for unit in default_table_style.encode_utf16() {
        body.extend_from_slice(&unit.to_le_bytes());
    }
    for unit in default_pivot_style.encode_utf16() {
        body.extend_from_slice(&unit.to_le_bytes());
    }
    write_biff_record(stream, TABLESTYLES_RECORD, &body);
}

fn write_pivot_workbook_extension_records(
    stream: &mut Vec<u8>,
    pivot_plan: &FormatPivotPlan,
) -> XlsResult<()> {
    if pivot_plan.caches.is_empty() {
        return Ok(());
    }

    write_biff_record(stream, COUNTRY_RECORD, &[0x01, 0x00, 0x01, 0x00]);
    write_biff_record(
        stream,
        RECALCID_RECORD,
        &[0xC1, 0x01, 0x00, 0x00, 0x35, 0xEA, 0x02, 0x00],
    );
    write_biff_record(
        stream,
        BOOKEXT_RECORD,
        &[
            0x63, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x16, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
        ],
    );
    write_biff_record(
        stream,
        COMPRESSPICTURES_RECORD,
        &[
            0x9B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00,
        ],
    );
    write_biff_record(
        stream,
        COMPAT12_RECORD,
        &[
            0x8C, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
        ],
    );
    Ok(())
}

fn write_pivot_pre_boundsheet_records(
    stream: &mut Vec<u8>,
    pivot_plan: &FormatPivotPlan,
) -> XlsResult<()> {
    if pivot_plan.caches.is_empty() {
        return Ok(());
    }

    for cache in &pivot_plan.caches {
        let cache_source_kind = xls_pivot_view_source_type(cache)?;
        write_biff_record(stream, 0x00D5, &(cache.cache_num as u16).to_le_bytes());
        write_biff_record(stream, 0x00E3, &cache_source_kind.to_le_bytes());
        write_pivot_cache_source_records(stream, cache)?;
        for payload in [
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
                0xFF, 0xFF, 0x04, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x18, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x41, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x34, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
        ] {
            write_biff_record(stream, 0x0864, payload);
        }
    }
    write_biff_record(stream, 0x0160, &0u16.to_le_bytes());
    Ok(())
}

fn write_pivot_cache_source_records(
    stream: &mut Vec<u8>,
    cache: &FormatPivotCache,
) -> XlsResult<()> {
    match &cache.source {
        FormatPivotSource::Worksheet {
            sheet_name,
            range,
            table_name,
            ..
        } => {
            if let Some(table_name) = table_name {
                write_dconname_record(stream, table_name)
            } else {
                write_dconref_record(stream, sheet_name, *range, None)
            }
        }
        FormatPivotSource::Consolidation { ranges } => {
            validate_xls_consolidation_sources(ranges)?;
            let pages = xls_consolidation_pages(ranges)?;
            write_sxtbl_record(stream, ranges, &pages)?;
            for range in ranges {
                if let Some(source_range) = range.range {
                    let sheet_name = range.sheet.as_deref().ok_or_else(|| {
                        XlsError::InvalidFormat(
                            "XLS consolidation source range requires a sheet name".into(),
                        )
                    })?;
                    write_dconref_record(
                        stream,
                        sheet_name,
                        source_range,
                        range.external_relationship_target.as_deref(),
                    )?;
                } else if let Some(name) = &range.name {
                    write_dconname_record(stream, name)?;
                }
            }
            for range in ranges {
                write_sxtbpg_record(stream, range, &pages)?;
            }
            write_sxtbrgiitm_records(stream, &pages)?;
            Ok(())
        }
        FormatPivotSource::External { .. } => unsupported_xls_external_pivot_source(),
        FormatPivotSource::Olap { .. } => unsupported_xls_olap_pivot_source(),
        FormatPivotSource::Scenario { name } => write_scenario_pivot_cache_source_record(name),
    }
}

fn write_scenario_pivot_cache_source_record(name: &str) -> XlsResult<()> {
    validate_xls_scenario_source(name)
}

fn validate_xls_scenario_source(name: &str) -> XlsResult<()> {
    if !name.is_empty() {
        return Err(XlsError::InvalidFormat(
            "XLS named scenario pivot source authoring requires Scenario Manager records and is not implemented yet".into(),
        ));
    }
    Ok(())
}

fn unsupported_xls_external_pivot_source<T>() -> XlsResult<T> {
    Err(XlsError::InvalidFormat(
        "XLS external database pivot source authoring requires complete BIFF8 query/data-connection records and is not implemented yet"
            .into(),
    ))
}

fn unsupported_xls_olap_pivot_source<T>() -> XlsResult<T> {
    Err(XlsError::InvalidFormat(
        "XLS OLAP pivot source authoring requires OLAP hierarchy/tuple BIFF8 records and is not implemented yet"
            .into(),
    ))
}

fn validate_xls_consolidation_sources(ranges: &[PivotSourceRange]) -> XlsResult<()> {
    if ranges.is_empty() {
        return Err(XlsError::InvalidFormat(
            "XLS consolidation pivot sources require at least one range".into(),
        ));
    }
    for range in ranges {
        let has_local_range = range.sheet.is_some() && range.range.is_some();
        let has_name = range.name.is_some() && range.range.is_none();
        if !has_local_range && !has_name {
            return Err(XlsError::InvalidFormat(
                "XLS consolidation writing currently requires local worksheet ranges or defined names"
                    .into(),
            ));
        }
        if range.external_relationship_id.is_some() && range.external_relationship_target.is_none()
        {
            return Err(XlsError::InvalidFormat(
                "XLS external consolidation references require a relationship target".into(),
            ));
        }
        if let Some(target) = &range.external_relationship_target {
            if target.trim().is_empty() {
                return Err(XlsError::InvalidFormat(
                    "XLS external consolidation relationship target cannot be blank".into(),
                ));
            }
            if range.sheet.is_none() || range.range.is_none() {
                return Err(XlsError::InvalidFormat(
                    "XLS external consolidation sources require a sheet and range".into(),
                ));
            }
        }
    }
    Ok(())
}

fn xls_consolidation_pages(ranges: &[PivotSourceRange]) -> XlsResult<Vec<Vec<String>>> {
    let page_count = ranges
        .iter()
        .map(|range| range.page_items.len())
        .max()
        .unwrap_or(0);
    if page_count > 4 {
        return Err(XlsError::InvalidFormat(
            "XLS consolidation pivot sources support at most four page fields".into(),
        ));
    }
    if page_count > 0
        && ranges
            .iter()
            .any(|range| range.page_items.len() != page_count)
    {
        return Err(XlsError::InvalidFormat(
            "XLS consolidation page fields require a page item for every source range".into(),
        ));
    }

    let mut pages = vec![Vec::<String>::new(); page_count];
    for range in ranges {
        for (index, item) in range.page_items.iter().enumerate() {
            if item.trim().is_empty() {
                return Err(XlsError::InvalidFormat(
                    "XLS consolidation page item names cannot be blank".into(),
                ));
            }
            if !pages[index].iter().any(|candidate| candidate == item) {
                pages[index].push(item.clone());
            }
        }
    }
    let item_count: usize = pages.iter().map(Vec::len).sum();
    checked_u16(item_count, "XLS consolidation page item count")?;
    Ok(pages)
}

fn write_sxtbl_record(
    stream: &mut Vec<u8>,
    ranges: &[PivotSourceRange],
    pages: &[Vec<String>],
) -> XlsResult<()> {
    let source_count = checked_u16(ranges.len(), "XLS consolidation source count")?;
    let page_field_count = checked_u16(pages.len(), "XLS consolidation page field count")?;
    let mut body = Vec::new();
    body.extend_from_slice(&source_count.to_le_bytes());
    body.extend_from_slice(&source_count.to_le_bytes());
    body.extend_from_slice(&page_field_count.to_le_bytes());
    write_biff_record(stream, 0x00D0, &body);
    Ok(())
}

fn write_sxtbpg_record(
    stream: &mut Vec<u8>,
    range: &PivotSourceRange,
    pages: &[Vec<String>],
) -> XlsResult<()> {
    let mut body = Vec::new();
    for (page_index, page) in pages.iter().enumerate() {
        let Some(item) = range.page_items.get(page_index) else {
            return Err(XlsError::InvalidFormat(
                "XLS consolidation source range is missing a page item".into(),
            ));
        };
        let Some(local_index) = page.iter().position(|candidate| candidate == item) else {
            return Err(XlsError::InvalidFormat(format!(
                "XLS consolidation page item is not declared: {item}"
            )));
        };
        let global_index = pages.iter().take(page_index).map(Vec::len).sum::<usize>() + local_index;
        body.extend_from_slice(
            &checked_u16(global_index, "XLS consolidation page item index")?.to_le_bytes(),
        );
    }
    write_biff_record(stream, 0x00D2, &body);
    Ok(())
}

fn write_sxtbrgiitm_records(stream: &mut Vec<u8>, pages: &[Vec<String>]) -> XlsResult<()> {
    let items = pages.iter().flatten().collect::<Vec<_>>();
    if items.is_empty() {
        return Ok(());
    }
    let mut count = Vec::new();
    count.extend_from_slice(
        &checked_u16(items.len(), "XLS consolidation page item count")?.to_le_bytes(),
    );
    write_biff_record(stream, 0x00D1, &count);
    for item in items {
        let mut body = Vec::new();
        push_xlunicode_string(&mut body, item)?;
        write_biff_record(stream, PIVOT_SXSTRING_RECORD, &body);
    }
    Ok(())
}

fn write_dconref_record(
    stream: &mut Vec<u8>,
    sheet_name: &str,
    range: CellRange,
    external_target: Option<&str>,
) -> XlsResult<()> {
    if range.start.row > u16::MAX as u32 || range.end.row > u16::MAX as u32 {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source range exceeds BIFF8 row limits".into(),
        ));
    }
    if range.start.col > u8::MAX as u16 || range.end.col > u8::MAX as u16 {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source range exceeds BIFF8 column limits".into(),
        ));
    }

    let (source_marker, source_name) = if let Some(target) = external_target {
        (0x01u16, format!("[{target}]{sheet_name}"))
    } else {
        (0x02u16, sheet_name.to_string())
    };

    let encoded_len = 1usize + source_name.encode_utf16().count();
    if encoded_len > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source sheet name is too long".into(),
        ));
    }

    let mut body = Vec::new();
    body.extend_from_slice(&(range.start.row as u16).to_le_bytes());
    body.extend_from_slice(&(range.end.row as u16).to_le_bytes());
    body.push(range.start.col as u8);
    body.push(range.end.col as u8);
    body.extend_from_slice(&(encoded_len as u16).to_le_bytes());
    if source_name.is_ascii() {
        body.push(0);
        body.push(source_marker as u8);
        body.extend_from_slice(source_name.as_bytes());
    } else {
        body.push(1);
        body.extend_from_slice(&source_marker.to_le_bytes());
        for unit in source_name.encode_utf16() {
            body.extend_from_slice(&unit.to_le_bytes());
        }
    }
    write_biff_record(stream, 0x0051, &body);
    Ok(())
}

fn write_dconname_record(stream: &mut Vec<u8>, table_name: &str) -> XlsResult<()> {
    let mut body = Vec::new();
    push_xlunicode_string(&mut body, table_name)?;
    body.extend_from_slice(&0u16.to_le_bytes());
    write_biff_record(stream, 0x0052, &body);
    Ok(())
}

fn write_pivot_sheet_records(
    stream: &mut Vec<u8>,
    workbook: &Workbook,
    pivot_plan: &FormatPivotPlan,
    pivot_layouts: &XlsPivotCacheLayouts,
    sheet_idx: usize,
    styles: &StyleTables,
) -> XlsResult<()> {
    for part in pivot_plan
        .tables
        .iter()
        .filter(|part| part.sheet_index == sheet_idx)
    {
        let cache = pivot_plan
            .caches
            .iter()
            .find(|cache| cache.cache_num == part.cache_num)
            .ok_or_else(|| XlsError::InvalidFormat("pivot cache part not found".into()))?;
        let sheet = workbook
            .worksheet(sheet_idx)
            .ok_or_else(|| XlsError::InvalidFormat("pivot sheet not found".into()))?;
        let pivot = sheet
            .pivot_tables()
            .get(part.pivot_index)
            .ok_or_else(|| XlsError::InvalidFormat("pivot table not found".into()))?;
        let layout = pivot_layouts.get(cache.cache_num)?;
        let date_system = workbook_date_system(workbook.settings().date_1904);
        validate_xls_pivot_groupings(cache, &pivot.groupings)?;
        validate_xls_pivot_grouping_axes(pivot)?;
        validate_xls_pivot_layout(pivot)?;
        validate_xls_pivot_axis_field_options(pivot)?;
        validate_xls_pivot_filters(pivot, layout)?;

        let multi_measure = pivot.measures.len() > 1;
        if pivot.rows.len() != 1
            || pivot.columns.len() > 1
            || pivot.page_fields.len() > 1
            || pivot.measures.is_empty()
            || (multi_measure
                && (!pivot.page_fields.is_empty() || !xls_values_field_on_columns(pivot)))
        {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot table {} uses a layout this BIFF8 writer slice does not encode yet",
                pivot.name
            )));
        }

        write_classic_pivot_view_records(stream, part, pivot, cache, &layout, styles)?;
        let effective_columns = xls_effective_column_fields(pivot, layout);
        let has_expanded_axis = expanded_axis_field_count(layout, &pivot.rows)?
            > checked_u16(pivot.rows.len(), "pivot visible row field count")?
            || expanded_axis_field_count(layout, &effective_columns)?
                > checked_u16(effective_columns.len(), "pivot visible column field count")?
            || has_grouped_page_axis(layout, pivot)
            || xls_cache_has_calculated_field(layout)
            || xls_pivot_has_sxaddl_filter(pivot);
        write_sxex_record_biff8(stream, pivot, &layout)?;
        write_sxview_record(stream, pivot, cache, has_expanded_axis)?;
        write_sxviewex9_record_biff8(stream, pivot, has_expanded_axis)?;
        write_pivot_frt_records(stream, pivot, &layout, has_expanded_axis, date_system)?;
    }
    Ok(())
}

fn validate_xls_pivot_layout(pivot: &duke_sheets_core::PivotTable) -> XlsResult<()> {
    if pivot.layout.subtotal_hidden_items {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot table {} enables hidden-item subtotals, which this BIFF8 writer slice does not encode yet",
            pivot.name
        )));
    }
    Ok(())
}

fn validate_xls_pivot_axis_field_options(pivot: &duke_sheets_core::PivotTable) -> XlsResult<()> {
    for field in pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
    {
        if field.item_page_count == 0 || field.item_page_count > u8::MAX as u32 {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot field {} uses item page count {}; BIFF8 SXVDEX stores this as a 1..=255 count",
                field.field.name, field.item_page_count
            )));
        }
    }
    Ok(())
}

fn validate_xls_pivot_filters(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    for filter in &pivot.filters {
        match filter {
            PivotFilter::FieldItems {
                field,
                allowed_items,
            } if pivot_uses_field_item_filter_axis(pivot, &field.name)
                || layout.field_index(&field.name).is_some() =>
            {
                if allowed_items.is_empty() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot field {} requires at least one selected item",
                        field.name
                    )));
                }
            }
            PivotFilter::FieldItems { field, .. } => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot table {} filters field {} outside the row, column, or page axes, which this BIFF8 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::TopN {
                field,
                measure,
                n,
                percent,
                ..
            } => {
                if *percent && *n > 100 {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses top-N percent threshold {n}; Excel stores pivot percentage filters as 1..=100",
                        pivot.name
                    )));
                }
                if *n == 0 || *n > i32::MAX as u32 {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses top-N threshold {n}; BIFF8 AutoShow requires 1..=2147483647",
                        pivot.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a top-N filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                xls_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Label {
                field,
                operator,
                value,
            } if xls_supported_label_filter_operator(*operator) => {
                if value.is_empty() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses an empty label filter value",
                        pivot.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a label filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Label { field, .. } => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot table {} uses a label filter operator for field {} that this BIFF8 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::LabelBetween {
                field, start, end, ..
            } => {
                if start.is_empty() || end.is_empty() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses an empty label range filter bound on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a label range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } if xls_supported_value_filter_operator(*operator) => {
                if !value.is_finite() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses a non-finite value filter threshold on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a value filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                xls_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Value { field, .. } => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot table {} uses a value filter operator for field {} that this BIFF8 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                ..
            } => {
                if !start.is_finite() || !end.is_finite() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses a non-finite value range filter threshold on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a value range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
                xls_pivot_measure_index_for_filter(pivot, measure)?;
            }
            PivotFilter::Date {
                field,
                operator,
                value,
            } if xls_supported_date_filter_operator(*operator) => {
                if !value.is_finite() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses a non-finite date filter operand on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a date filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Date { field, .. } => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot table {} uses a date filter operator for field {} that this BIFF8 writer slice does not encode yet",
                    pivot.name, field.name
                )));
            }
            PivotFilter::DateBetween {
                field, start, end, ..
            } => {
                if !start.is_finite() || !end.is_finite() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses a non-finite date range filter operand on field {}",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a date range filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::DatePeriod { field, period } => {
                if xls_date_period_filter_codes(*period).is_none() {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} uses a date-period filter for field {} that this BIFF8 writer slice does not encode yet",
                        pivot.name, field.name
                    )));
                }
                if !pivot_axis_contains_field(&pivot.rows, &field.name)
                    && !pivot_axis_contains_field(&pivot.columns, &field.name)
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot table {} applies a date-period filter to field {} outside the row or column axes",
                        pivot.name, field.name
                    )));
                }
            }
            PivotFilter::Unsupported { kind, .. } => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot table {} contains unsupported preserved filter {kind}, which this BIFF8 writer slice does not encode yet",
                    pivot.name
                )));
            }
        }
    }
    Ok(())
}

fn write_classic_pivot_view_records(
    stream: &mut Vec<u8>,
    part: &FormatPivotTable,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    styles: &StyleTables,
) -> XlsResult<()> {
    let effective_columns = xls_effective_column_fields(pivot, layout);
    let axis_tuples = build_xls_axis_tuples(part, pivot, cache, layout, &effective_columns)?;
    write_sxview_record_biff8(
        stream,
        pivot,
        cache,
        layout,
        &axis_tuples,
        &effective_columns,
    )?;
    for (field_index, field) in layout.fields.iter().enumerate() {
        let axis = xls_pivot_field_axis(pivot, layout, field_index);
        let axis_field = xls_axis_field_for_layout_field(pivot, layout, field_index);
        write_sxvd_record(stream, pivot, layout, field_index, field, axis, axis_field)?;
    }
    let values_on_columns = xls_values_field_on_columns(pivot);
    write_sxivd_record(stream, layout, &pivot.rows)?;
    if values_on_columns {
        write_values_sxivd_record(stream);
    } else if !effective_columns.is_empty() {
        write_sxivd_record(stream, layout, &effective_columns)?;
    }
    write_sxpi_records(stream, pivot, layout)?;
    write_sxdi_records(stream, pivot, layout, styles)?;
    write_sxli_collection(stream, pivot, layout, &pivot.rows, &axis_tuples.rows)?;
    if values_on_columns {
        write_values_axis_sxli_collection(stream, pivot)?;
    } else {
        write_sxli_collection(
            stream,
            pivot,
            layout,
            &effective_columns,
            &axis_tuples.columns,
        )?;
    }
    Ok(())
}

fn write_sxview_record_biff8(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    axis_tuples: &XlsPivotAxisTuples,
    effective_columns: &[PivotField],
) -> XlsResult<()> {
    let data_caption = if pivot.layout.data_caption.trim().is_empty() {
        "Values"
    } else {
        pivot.layout.data_caption.as_str()
    };
    let row_line_count = axis_line_count(&pivot.rows, &axis_tuples.rows)?;
    let values_on_columns = xls_values_field_on_columns(pivot);
    let column_line_count = if values_on_columns {
        checked_u16(pivot.measures.len(), "pivot values axis line count")?
    } else {
        axis_line_count(effective_columns, &axis_tuples.columns)?
    };
    let has_calculated_field = xls_cache_has_calculated_field(layout);
    let row_axis_count = expanded_axis_field_count(layout, &pivot.rows)?;
    let visible_row_axis_count = visible_axis_field_count(layout, &pivot.rows)?;
    let column_axis_count = expanded_axis_field_count(layout, effective_columns)?
        .saturating_add(if values_on_columns { 1 } else { 0 });
    let page_axis_count = checked_u16(
        xls_effective_page_field_count(pivot, layout),
        "pivot page field count",
    )?;
    let data_field_count = checked_u16(pivot.measures.len(), "pivot data field count")?;
    let has_calculated_item = !pivot.calculated_items.is_empty();
    let (page_rows, _) = page_field_area_size(pivot, layout);
    let first_row_offset = if page_rows == 0 {
        0
    } else {
        page_rows.saturating_add(1)
    };
    let first_row = checked_biff8_row(
        pivot.target.row.saturating_add(first_row_offset),
        "pivot target row",
    )?;
    let first_col = checked_biff8_col(pivot.target.col, "pivot target column")?;
    let first_header_row = if values_on_columns && effective_columns.is_empty() {
        first_row
    } else {
        first_row.saturating_add(1)
    };
    let first_data_row = if values_on_columns && effective_columns.is_empty() {
        first_row.saturating_add(1)
    } else {
        first_header_row.saturating_add(column_axis_count)
    };
    let first_data_col = first_col.saturating_add(visible_row_axis_count);
    let last_row = if values_on_columns && effective_columns.is_empty() {
        first_row.saturating_add(row_line_count)
    } else {
        first_row
            .saturating_add(column_axis_count)
            .saturating_add(row_line_count)
    };
    let last_col = first_col
        .saturating_add(visible_row_axis_count)
        .saturating_add(column_line_count)
        .saturating_sub(1);
    let mut body = Vec::new();
    body.extend_from_slice(&first_row.to_le_bytes());
    body.extend_from_slice(&last_row.to_le_bytes());
    body.extend_from_slice(&first_col.to_le_bytes());
    body.extend_from_slice(&last_col.to_le_bytes());
    body.extend_from_slice(&first_header_row.to_le_bytes());
    body.extend_from_slice(&first_data_row.to_le_bytes());
    body.extend_from_slice(&first_data_col.to_le_bytes());
    body.extend_from_slice(&((cache.cache_num.saturating_sub(1)) as u16).to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&0x0002u16.to_le_bytes());
    body.extend_from_slice(&(-1i16).to_le_bytes());
    body.extend_from_slice(
        &checked_u16(layout.fields.len(), "pivot view field count")?.to_le_bytes(),
    );
    body.extend_from_slice(&row_axis_count.to_le_bytes());
    body.extend_from_slice(&column_axis_count.to_le_bytes());
    body.extend_from_slice(&page_axis_count.to_le_bytes());
    body.extend_from_slice(&data_field_count.to_le_bytes());
    body.extend_from_slice(&row_line_count.to_le_bytes());
    body.extend_from_slice(&column_line_count.to_le_bytes());
    let mut view_flags: u16 = if column_axis_count > 0
        || page_axis_count > 0
        || row_axis_count > visible_row_axis_count
        || has_calculated_field
        || has_calculated_item
    {
        0x0208
    } else {
        0x0000
    };
    if pivot.layout.show_row_grand_totals {
        view_flags |= 0x0001;
    }
    if pivot.layout.show_column_grand_totals {
        view_flags |= 0x0002;
    }
    body.extend_from_slice(&view_flags.to_le_bytes());
    body.extend_from_slice(&0x0001u16.to_le_bytes());
    body.extend_from_slice(&xlunicode_len_u16(&pivot.name)?.to_le_bytes());
    body.extend_from_slice(&xlunicode_len_u16(data_caption)?.to_le_bytes());
    push_xlunicode_string_no_cch(&mut body, &pivot.name)?;
    push_xlunicode_string_no_cch(&mut body, data_caption)?;
    write_biff_record(stream, records::SXVIEW, &body);
    Ok(())
}

fn write_sxex_record_biff8(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    let error_caption = if pivot.layout.show_error || pivot.layout.error_caption.is_some() {
        Some(pivot.layout.error_caption.as_deref().unwrap_or(""))
    } else {
        None
    };
    let missing_caption = pivot.layout.missing_caption.as_deref();
    let error_len = optional_xlunicode_len_u16(error_caption)?;
    let missing_len = optional_xlunicode_len_u16(missing_caption)?;
    let (page_rows, page_cols) = page_field_area_size(pivot, layout);

    let mut grbit1 = 0x0200u16;
    if pivot.layout.page_over_then_down {
        grbit1 |= 0x0001;
    }
    grbit1 |= ((pivot.layout.page_wrap.min(0xFF) as u16) << 1) & 0x01FE;

    let mut grbit2 = 0u16;
    set_biff_flag(&mut grbit2, 0x0001, pivot.layout.enable_wizard);
    set_biff_flag(&mut grbit2, 0x0002, pivot.layout.enable_drill);
    set_biff_flag(&mut grbit2, 0x0004, pivot.layout.enable_field_properties);
    set_biff_flag(
        &mut grbit2,
        0x0008,
        pivot.refresh_policy.preserve_formatting,
    );
    set_biff_flag(&mut grbit2, 0x0010, pivot.layout.merge_item_labels);
    set_biff_flag(&mut grbit2, 0x0020, pivot.layout.show_error);
    set_biff_flag(&mut grbit2, 0x0040, pivot.layout.show_missing);
    set_biff_flag(&mut grbit2, 0x0080, pivot.layout.subtotal_hidden_items);
    set_biff_flag(&mut grbit2, 0x0200, pivot.layout.edit_data);
    set_biff_flag(&mut grbit2, 0x0400, pivot.layout.disable_field_list);

    let mut body = Vec::new();
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&error_len.to_le_bytes());
    body.extend_from_slice(&missing_len.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(
        &checked_u16(page_rows as usize, "pivot page field rows")?.to_le_bytes(),
    );
    body.extend_from_slice(
        &checked_u16(page_cols as usize, "pivot page field columns")?.to_le_bytes(),
    );
    body.extend_from_slice(&grbit1.to_le_bytes());
    body.extend_from_slice(&grbit2.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    if let Some(caption) = error_caption {
        push_xlunicode_string_no_cch(&mut body, caption)?;
    }
    if let Some(caption) = missing_caption {
        push_xlunicode_string_no_cch(&mut body, caption)?;
    }
    write_biff_record(stream, 0x00F1, &body);
    Ok(())
}

fn write_sxviewex9_record_biff8(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    has_expanded_axis: bool,
) -> XlsResult<()> {
    let mut flags = 0x0004u32;
    if pivot.layout.field_print_titles {
        flags |= 0x0002;
    }
    if has_expanded_axis || pivot.layout.item_print_titles {
        flags |= 0x0020;
    }

    let mut body = Vec::new();
    body.extend_from_slice(&0x0810u16.to_le_bytes());
    body.extend_from_slice(&0x0002u16.to_le_bytes());
    body.extend_from_slice(&0u32.to_le_bytes());
    body.extend_from_slice(&flags.to_le_bytes());
    body.extend_from_slice(&0x0001u16.to_le_bytes());
    push_xlunicode_string(
        &mut body,
        pivot.layout.grand_total_caption.as_deref().unwrap_or(""),
    )?;
    write_biff_record(stream, 0x0810, &body);
    Ok(())
}

fn set_biff_flag(flags: &mut u16, mask: u16, enabled: bool) {
    if enabled {
        *flags |= mask;
    } else {
        *flags &= !mask;
    }
}

fn write_sxvd_record(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    field_index: usize,
    field: &XlsPivotFieldLayout,
    axis: u16,
    axis_field: Option<&PivotField>,
) -> XlsResult<()> {
    if field_index > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(
            "pivot field index exceeds BIFF8 limits".into(),
        ));
    }
    let has_hidden_item_filter = xls_field_has_hidden_item_filter(pivot, layout, field_index);
    let item_count = if matches!(axis, 0x0001 | 0x0002 | 0x0004)
        || has_hidden_item_filter
        || matches!(
            field.kind,
            XlsPivotFieldKind::DateSource { .. } | XlsPivotFieldKind::DateFilterSource { .. }
        ) {
        checked_u16(
            field
                .shared_items
                .len()
                .saturating_add(xls_sxvd_subtotal_count(axis_field) as usize),
            "pivot item count",
        )?
    } else {
        0
    };

    let mut body = Vec::new();
    let data_axis_calculated_field = axis == 0x0008 && field.formula.is_some();
    let (subtotal_count, subtotal_flags) = if data_axis_calculated_field {
        (0, 0)
    } else {
        (
            xls_sxvd_subtotal_count(axis_field),
            xls_sxvd_subtotal_flags(axis_field),
        )
    };
    body.extend_from_slice(&axis.to_le_bytes());
    body.extend_from_slice(&subtotal_count.to_le_bytes());
    body.extend_from_slice(&subtotal_flags.to_le_bytes());
    body.extend_from_slice(&item_count.to_le_bytes());
    if let Some(caption) = axis_field.and_then(|field| field.caption.as_ref()) {
        body.extend_from_slice(&xlunicode_len_u16(caption)?.to_le_bytes());
        push_xlunicode_string_no_cch(&mut body, caption)?;
    } else {
        body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    }
    write_biff_record(stream, records::SXVD, &body);

    let hidden_items = if let Some(filter_field_name) =
        xls_filter_field_name_for_hidden_items(layout, field_index)
    {
        field_filter_hidden_item_indexes(pivot, filter_field_name, field)?
    } else {
        HashSet::new()
    };

    if item_count > 0 {
        for item_index in 0..field.shared_items.len() {
            let item_index = checked_u16(item_index, "pivot item index")?;
            let mut flags = u16::from(hidden_items.contains(&item_index));
            if field
                .calculated_item_indexes
                .contains(&(item_index as usize))
            {
                flags |= 0x0008;
            }
            write_sxvi_record(stream, 0x0000, flags, item_index as i16)?;
        }
        write_sxvi_subtotal_records(stream, axis_field)?;
    }

    write_sxvdex_record(stream, pivot, axis, axis_field, data_axis_calculated_field)?;
    Ok(())
}

fn write_sxvi_record(
    stream: &mut Vec<u8>,
    item_type: u16,
    flags: u16,
    cache_index: i16,
) -> XlsResult<()> {
    let mut item = Vec::new();
    item.extend_from_slice(&item_type.to_le_bytes());
    item.extend_from_slice(&flags.to_le_bytes());
    item.extend_from_slice(&cache_index.to_le_bytes());
    item.extend_from_slice(&0xFFFFu16.to_le_bytes());
    write_biff_record(stream, 0x00B2, &item);
    Ok(())
}

fn write_sxvi_subtotal_records(
    stream: &mut Vec<u8>,
    axis_field: Option<&PivotField>,
) -> XlsResult<()> {
    let subtotals = axis_field
        .map(xls_sxvd_subtotal_items)
        .unwrap_or_else(|| vec![PivotSubtotal::Automatic]);
    for subtotal in subtotals {
        write_sxvi_record(stream, xls_sxvi_subtotal_item_type(subtotal), 0, -1)?;
    }
    Ok(())
}

fn xls_sxvi_subtotal_item_type(subtotal: PivotSubtotal) -> u16 {
    match subtotal {
        PivotSubtotal::Automatic => 0x0001,
        PivotSubtotal::Sum => 0x0002,
        PivotSubtotal::Count => 0x0003,
        PivotSubtotal::Average => 0x0004,
        PivotSubtotal::Max => 0x0005,
        PivotSubtotal::Min => 0x0006,
        PivotSubtotal::Product => 0x0007,
        PivotSubtotal::CountNumbers => 0x0008,
        PivotSubtotal::StdDev => 0x0009,
        PivotSubtotal::StdDevP => 0x000A,
        PivotSubtotal::Var => 0x000B,
        PivotSubtotal::VarP => 0x000C,
        PivotSubtotal::None => 0x0000,
    }
}

fn write_sxvdex_record(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    axis: u16,
    axis_field: Option<&PivotField>,
    data_axis_calculated_field: bool,
) -> XlsResult<()> {
    let top_n_filter = xls_sxvd_top_n_filter(pivot, axis_field)?;
    let mut body = Vec::new();
    body.extend_from_slice(
        &xls_sxvdex_grbit1(axis, axis_field, data_axis_calculated_field, top_n_filter)
            .to_le_bytes(),
    );
    body.extend_from_slice(&xls_sxvd_sort_measure_index(pivot, axis_field)?.to_le_bytes());
    body.extend_from_slice(
        &top_n_filter
            .map(|filter| filter.measure_index)
            .unwrap_or(-1)
            .to_le_bytes(),
    );
    body.extend_from_slice(&0u16.to_le_bytes());
    if let Some(caption) = axis_field.and_then(|field| field.subtotal_caption.as_ref()) {
        body.extend_from_slice(&xlunicode_len_u16(caption)?.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes());
        push_xlunicode_string_no_cch(&mut body, caption)?;
    } else {
        body.extend_from_slice(&0xFFFFu16.to_le_bytes());
        body.extend_from_slice(&0u64.to_le_bytes());
    }
    write_biff_record(stream, 0x0100, &body);
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct XlsSxvdTopNFilter {
    n: u32,
    top: bool,
    measure_index: i16,
}

#[derive(Debug, Clone, Copy)]
struct XlsSxaddlTopNFilter {
    n: u32,
    top: bool,
    percent: bool,
    field_source_index: usize,
    measure_source_index: usize,
}

#[derive(Debug, Clone)]
struct XlsSxaddlLabelFilter {
    kind: XlsSxaddlLabelFilterKind,
    field_source_index: usize,
}

#[derive(Debug, Clone)]
enum XlsSxaddlLabelFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: String,
    },
    Between {
        start: String,
        end: String,
        not_between: bool,
    },
}

#[derive(Debug, Clone, Copy)]
struct XlsSxaddlValueFilter {
    kind: XlsSxaddlValueFilterKind,
    field_source_index: usize,
    measure_index: i16,
}

#[derive(Debug, Clone, Copy)]
struct XlsSxaddlDateFilter {
    kind: XlsSxaddlDateFilterKind,
    field_source_index: usize,
}

#[derive(Debug, Clone, Copy)]
enum XlsSxaddlValueFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: f64,
    },
    Between {
        start: f64,
        end: f64,
        not_between: bool,
    },
}

#[derive(Debug, Clone, Copy)]
enum XlsSxaddlDateFilterKind {
    Comparison {
        operator: PivotFilterOperator,
        value: f64,
    },
    Between {
        start: f64,
        end: f64,
        not_between: bool,
    },
    Period(PivotDatePeriod),
}

#[derive(Debug, Clone)]
enum XlsSxaddlPivotFilter {
    TopN(XlsSxaddlTopNFilter),
    Label(XlsSxaddlLabelFilter),
    Value(XlsSxaddlValueFilter),
    Date(XlsSxaddlDateFilter),
}

impl XlsSxaddlPivotFilter {
    fn field_source_index(&self) -> usize {
        match self {
            Self::TopN(filter) => filter.field_source_index,
            Self::Label(filter) => filter.field_source_index,
            Self::Value(filter) => filter.field_source_index,
            Self::Date(filter) => filter.field_source_index,
        }
    }

    fn collection_filter_type(&self) -> u32 {
        match self {
            Self::TopN(_) => 2,
            Self::Label(filter) => match filter.kind {
                XlsSxaddlLabelFilterKind::Comparison { operator, .. } => match operator {
                    PivotFilterOperator::Equals => 4,
                    PivotFilterOperator::NotEquals => 5,
                    PivotFilterOperator::BeginsWith => 6,
                    PivotFilterOperator::DoesNotBeginWith => 7,
                    PivotFilterOperator::EndsWith => 8,
                    PivotFilterOperator::DoesNotEndWith => 9,
                    PivotFilterOperator::Contains => 10,
                    PivotFilterOperator::DoesNotContain => 11,
                    PivotFilterOperator::GreaterThan => 12,
                    PivotFilterOperator::GreaterThanOrEqual => 13,
                    PivotFilterOperator::LessThan => 14,
                    PivotFilterOperator::LessThanOrEqual => 15,
                },
                XlsSxaddlLabelFilterKind::Between {
                    not_between: false, ..
                } => 16,
                XlsSxaddlLabelFilterKind::Between {
                    not_between: true, ..
                } => 17,
            },
            Self::Value(filter) => match filter.kind {
                XlsSxaddlValueFilterKind::Comparison { operator, .. } => {
                    value_filter_type_and_operator(operator).0
                }
                XlsSxaddlValueFilterKind::Between {
                    not_between: false, ..
                } => 24,
                XlsSxaddlValueFilterKind::Between {
                    not_between: true, ..
                } => 25,
            },
            Self::Date(filter) => match filter.kind {
                XlsSxaddlDateFilterKind::Comparison { operator, .. } => {
                    xls_date_filter_type_and_operator(operator)
                        .expect("validated date filter operator")
                        .0
                }
                XlsSxaddlDateFilterKind::Between {
                    not_between: false, ..
                } => 29,
                XlsSxaddlDateFilterKind::Between {
                    not_between: true, ..
                } => 65,
                XlsSxaddlDateFilterKind::Period(period) => {
                    xls_date_period_filter_codes(period)
                        .expect("validated date period filter")
                        .0
                }
            },
        }
    }

    fn collection_measure_index(&self) -> i32 {
        match self {
            Self::TopN(_) | Self::Label(_) | Self::Date(_) => 0,
            Self::Value(filter) => i32::from(filter.measure_index),
        }
    }

    fn collection_trailing_sentinel(&self) -> i32 {
        match self {
            Self::TopN(_) => -1,
            Self::Label(_) | Self::Date(_) => 0,
            Self::Value(_) => -1,
        }
    }
}

fn xls_sxvd_top_n_filter(
    pivot: &duke_sheets_core::PivotTable,
    axis_field: Option<&PivotField>,
) -> XlsResult<Option<XlsSxvdTopNFilter>> {
    let Some(axis_field) = axis_field else {
        return Ok(None);
    };
    let Some(filter) = pivot.filters.iter().find_map(|filter| match filter {
        PivotFilter::TopN {
            field,
            measure,
            n,
            top,
            percent: false,
        } if field.name.eq_ignore_ascii_case(&axis_field.field.name) => Some((measure, *n, *top)),
        _ => None,
    }) else {
        return Ok(None);
    };
    Ok(Some(XlsSxvdTopNFilter {
        n: filter.1,
        top: filter.2,
        measure_index: xls_pivot_measure_index_for_filter(pivot, filter.0)?,
    }))
}

fn xls_sxaddl_pivot_filters(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<Vec<XlsSxaddlPivotFilter>> {
    let mut filters = Vec::new();
    for filter in &pivot.filters {
        match filter {
            PivotFilter::TopN {
                field,
                measure,
                n,
                top,
                percent,
            } => {
                if !*percent {
                    continue;
                }
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                let measure_source_index =
                    layout.field_index(&measure.field.name).ok_or_else(|| {
                        XlsError::InvalidFormat(format!(
                            "XLS pivot table {} filters by unknown measure field {}",
                            pivot.name, measure.field.name
                        ))
                    })?;
                filters.push(XlsSxaddlPivotFilter::TopN(XlsSxaddlTopNFilter {
                    n: *n,
                    top: *top,
                    percent: *percent,
                    field_source_index,
                    measure_source_index,
                }));
            }
            PivotFilter::Label {
                field,
                operator,
                value,
            } if xls_supported_label_filter_operator(*operator) => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                filters.push(XlsSxaddlPivotFilter::Label(XlsSxaddlLabelFilter {
                    kind: XlsSxaddlLabelFilterKind::Comparison {
                        operator: *operator,
                        value: value.clone(),
                    },
                    field_source_index,
                }));
            }
            PivotFilter::LabelBetween {
                field,
                start,
                end,
                not_between,
            } => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                filters.push(XlsSxaddlPivotFilter::Label(XlsSxaddlLabelFilter {
                    kind: XlsSxaddlLabelFilterKind::Between {
                        start: start.clone(),
                        end: end.clone(),
                        not_between: *not_between,
                    },
                    field_source_index,
                }));
            }
            PivotFilter::Value {
                field,
                measure,
                operator,
                value,
            } if xls_supported_value_filter_operator(*operator) => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                let measure_index = xls_pivot_measure_index_for_filter(pivot, measure)?;
                filters.push(XlsSxaddlPivotFilter::Value(XlsSxaddlValueFilter {
                    kind: XlsSxaddlValueFilterKind::Comparison {
                        operator: *operator,
                        value: *value,
                    },
                    field_source_index,
                    measure_index,
                }));
            }
            PivotFilter::ValueBetween {
                field,
                measure,
                start,
                end,
                not_between,
            } => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                let measure_index = xls_pivot_measure_index_for_filter(pivot, measure)?;
                filters.push(XlsSxaddlPivotFilter::Value(XlsSxaddlValueFilter {
                    kind: XlsSxaddlValueFilterKind::Between {
                        start: *start,
                        end: *end,
                        not_between: *not_between,
                    },
                    field_source_index,
                    measure_index,
                }));
            }
            PivotFilter::Date {
                field,
                operator,
                value,
            } if xls_supported_date_filter_operator(*operator) => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                filters.push(XlsSxaddlPivotFilter::Date(XlsSxaddlDateFilter {
                    kind: XlsSxaddlDateFilterKind::Comparison {
                        operator: *operator,
                        value: *value,
                    },
                    field_source_index,
                }));
            }
            PivotFilter::DateBetween {
                field,
                start,
                end,
                not_between,
            } => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                filters.push(XlsSxaddlPivotFilter::Date(XlsSxaddlDateFilter {
                    kind: XlsSxaddlDateFilterKind::Between {
                        start: *start,
                        end: *end,
                        not_between: *not_between,
                    },
                    field_source_index,
                }));
            }
            PivotFilter::DatePeriod { field, period } => {
                let field_source_index =
                    xls_sxaddl_source_field_index_for_semantic_field(layout, &field.name)
                        .ok_or_else(|| {
                            XlsError::InvalidFormat(format!(
                                "XLS pivot table {} filters unknown field {}",
                                pivot.name, field.name
                            ))
                        })?;
                filters.push(XlsSxaddlPivotFilter::Date(XlsSxaddlDateFilter {
                    kind: XlsSxaddlDateFilterKind::Period(*period),
                    field_source_index,
                }));
            }
            _ => {}
        }
    }
    Ok(filters)
}

fn xls_sxaddl_source_field_index_for_semantic_field(
    layout: &XlsPivotCacheLayout,
    field_name: &str,
) -> Option<usize> {
    let field_index = layout
        .axis_field_index(field_name)
        .or_else(|| layout.field_index(field_name))?;
    let field = layout.fields.get(field_index)?;
    Some(match field.kind {
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        }
        | XlsPivotFieldKind::ManualGroup {
            source_field_index, ..
        } => source_field_index,
        XlsPivotFieldKind::Regular
        | XlsPivotFieldKind::NumberGroup { .. }
        | XlsPivotFieldKind::DateFilterSource { .. }
        | XlsPivotFieldKind::DateSource { .. }
        | XlsPivotFieldKind::ManualSource { .. } => field_index,
    })
}

fn xls_axis_field_for_layout_field<'a>(
    pivot: &'a duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    field_index: usize,
) -> Option<&'a PivotField> {
    let field = layout.fields.get(field_index)?;
    let axis_field_index = match field.kind {
        XlsPivotFieldKind::DateSource { .. } => return None,
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        }
        | XlsPivotFieldKind::ManualGroup {
            source_field_index, ..
        } => source_field_index,
        XlsPivotFieldKind::ManualSource { .. }
            if pivot_axis_contains_field(&pivot.page_fields, &field.name)
                && !pivot_axis_contains_field(&pivot.rows, &field.name)
                && !pivot_axis_contains_field(&pivot.columns, &field.name) =>
        {
            return None;
        }
        _ => field_index,
    };

    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| {
            layout.field_index(&field.field.name) == Some(axis_field_index)
                || layout.axis_field_index(&field.field.name) == Some(axis_field_index)
        })
}

fn xls_sxvd_subtotal_count(axis_field: Option<&PivotField>) -> u16 {
    axis_field
        .map(xls_sxvd_subtotal_items)
        .unwrap_or_else(|| vec![PivotSubtotal::Automatic])
        .len() as u16
}

fn xls_sxvd_subtotal_flags(axis_field: Option<&PivotField>) -> u16 {
    let Some(field) = axis_field else {
        return 0x0001;
    };

    xls_sxvd_subtotal_items(field)
        .into_iter()
        .fold(0u16, |flags, subtotal| {
            flags | xls_sxvd_subtotal_flag(subtotal)
        })
}

fn xls_sxvd_subtotal_items(field: &PivotField) -> Vec<PivotSubtotal> {
    if field.subtotals.is_empty() {
        return match field.subtotal {
            PivotSubtotal::None => Vec::new(),
            subtotal => vec![subtotal],
        };
    }

    let custom = field
        .subtotals
        .iter()
        .copied()
        .filter(|subtotal| subtotal.is_custom_function())
        .collect::<Vec<_>>();
    if !custom.is_empty() {
        custom
    } else if field
        .subtotals
        .iter()
        .any(|subtotal| matches!(subtotal, PivotSubtotal::Automatic))
    {
        vec![PivotSubtotal::Automatic]
    } else {
        Vec::new()
    }
}

fn xls_sxvd_subtotal_flag(subtotal: PivotSubtotal) -> u16 {
    match subtotal {
        PivotSubtotal::Automatic => 0x0001,
        PivotSubtotal::Sum => 0x0002,
        PivotSubtotal::Count => 0x0004,
        PivotSubtotal::Average => 0x0008,
        PivotSubtotal::Max => 0x0010,
        PivotSubtotal::Min => 0x0020,
        PivotSubtotal::Product => 0x0040,
        PivotSubtotal::CountNumbers => 0x0080,
        PivotSubtotal::StdDev => 0x0100,
        PivotSubtotal::StdDevP => 0x0200,
        PivotSubtotal::Var => 0x0400,
        PivotSubtotal::VarP => 0x0800,
        PivotSubtotal::None => 0x0000,
    }
}

fn xls_sxvdex_grbit1(
    axis: u16,
    axis_field: Option<&PivotField>,
    data_axis_calculated_field: bool,
    top_n_filter: Option<XlsSxvdTopNFilter>,
) -> u32 {
    if axis == 0x0008 && data_axis_calculated_field {
        return 0x0AA0_3410;
    }

    let mut flags = 0x0AA0_141Eu32;
    let Some(field) = axis_field else {
        return flags;
    };

    set_u32_flag(&mut flags, 0x0001, field.show_empty_items);
    set_u32_flag(&mut flags, 0x4000, field.insert_page_break);
    set_u32_flag(&mut flags, 0x8000, false);
    set_u32_flag(&mut flags, 0x0040_0000, field.insert_blank_row);
    set_u32_flag(&mut flags, 0x0080_0000, field.subtotal_top);

    match field.sort {
        PivotSort::None => {
            set_u32_flag(&mut flags, 0x0200, false);
            set_u32_flag(&mut flags, 0x0400, false);
        }
        PivotSort::Ascending => {
            set_u32_flag(&mut flags, 0x0200, field.sort_by_measure.is_some());
            set_u32_flag(&mut flags, 0x0400, true);
        }
        PivotSort::Descending => {
            set_u32_flag(&mut flags, 0x0200, true);
            set_u32_flag(&mut flags, 0x0400, false);
        }
    }

    if let Some(filter) = top_n_filter {
        set_u32_flag(&mut flags, 0x0800, true);
        set_u32_flag(&mut flags, 0x1000, filter.top);
    }
    if let Some(field) = axis_field {
        flags = (flags & 0x00FF_FFFF) | ((field.item_page_count & 0xFF) << 24);
    }

    flags
}

fn xls_sxvd_sort_measure_index(
    pivot: &duke_sheets_core::PivotTable,
    axis_field: Option<&PivotField>,
) -> XlsResult<i16> {
    let Some(axis_field) = axis_field else {
        return Ok(-1);
    };
    if matches!(axis_field.sort, PivotSort::None) {
        return Ok(-1);
    }
    let Some(sort_measure) = axis_field.sort_by_measure.as_ref() else {
        return Ok(-1);
    };

    let found = pivot
        .measures
        .iter()
        .any(|measure| pivot_measure_matches_target(measure, sort_measure));
    if !found {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot table {} sorts field {} by an unknown measure",
            pivot.name, axis_field.field.name
        )));
    }

    Ok(-1)
}

fn xls_pivot_measure_index_for_filter(
    pivot: &duke_sheets_core::PivotTable,
    target: &PivotMeasure,
) -> XlsResult<i16> {
    pivot
        .measures
        .iter()
        .position(|measure| pivot_measure_matches_target(measure, target))
        .map(|index| index as i16)
        .ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot table {} filters by an unknown measure {}",
                pivot.name, target.field.name
            ))
        })
}

fn set_u32_flag(flags: &mut u32, mask: u32, value: bool) {
    if value {
        *flags |= mask;
    } else {
        *flags &= !mask;
    }
}

fn write_sxivd_record(
    stream: &mut Vec<u8>,
    layout: &XlsPivotCacheLayout,
    fields: &[duke_sheets_core::PivotField],
) -> XlsResult<()> {
    if fields.is_empty() {
        return Ok(());
    }
    let mut body = Vec::new();
    for field_index in expanded_axis_field_indexes(layout, fields)? {
        body.extend_from_slice(&checked_u16(field_index, "pivot axis field index")?.to_le_bytes());
    }
    write_biff_record(stream, records::SXIVD, &body);
    Ok(())
}

fn write_values_sxivd_record(stream: &mut Vec<u8>) {
    write_biff_record(stream, records::SXIVD, &(-2i16).to_le_bytes());
}

fn write_sxpi_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    for field in &pivot.page_fields {
        let field_index = layout
            .page_axis_field_index(&field.field.name)
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "pivot references unknown page field {}",
                    field.field.name
                ))
            })?;
        let selected_item =
            selected_page_item_index(pivot, &field.field.name, &layout.fields[field_index])?
                .unwrap_or_else(|| default_page_item_index(&layout.fields[field_index]));
        let mut body = Vec::new();
        body.extend_from_slice(&checked_u16(field_index, "pivot page field index")?.to_le_bytes());
        body.extend_from_slice(&selected_item.to_le_bytes());
        body.extend_from_slice(&1u16.to_le_bytes());
        write_biff_record(stream, 0x00B6, &body);
    }
    for field_index in xls_synthetic_consolidation_page_field_indexes(pivot, layout) {
        let mut body = Vec::new();
        body.extend_from_slice(&checked_u16(field_index, "pivot page field index")?.to_le_bytes());
        body.extend_from_slice(&0x7FFDu16.to_le_bytes());
        body.extend_from_slice(&1u16.to_le_bytes());
        write_biff_record(stream, 0x00B6, &body);
    }
    Ok(())
}

fn write_sxdi_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    styles: &StyleTables,
) -> XlsResult<()> {
    if pivot.measures.is_empty() {
        return Err(XlsError::InvalidFormat(
            "pivot table has no data field".into(),
        ));
    }

    for measure in &pivot.measures {
        let field_index = layout.field_index(&measure.field.name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot references unknown measure field {}",
                measure.field.name
            ))
        })?;
        if field_index > u16::MAX as usize {
            return Err(XlsError::InvalidFormat(
                "pivot data field index exceeds BIFF8 limits".into(),
            ));
        }

        let data_field = xls_data_field_options(measure, layout, styles)?;
        let mut body = Vec::new();
        body.extend_from_slice(&(field_index as u16).to_le_bytes());
        body.extend_from_slice(&xls_pivot_aggregate_code(measure.aggregate).to_le_bytes());
        body.extend_from_slice(&data_field.show_as.to_le_bytes());
        body.extend_from_slice(&data_field.base_field.to_le_bytes());
        body.extend_from_slice(&data_field.base_item.to_le_bytes());
        body.extend_from_slice(&data_field.num_format.to_le_bytes());
        push_xlunicode_string(&mut body, &measure.caption())?;
        write_biff_record(stream, 0x00C5, &body);
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct XlsDataFieldOptions {
    show_as: u16,
    base_field: u16,
    base_item: u16,
    num_format: u16,
}

fn xls_data_field_options(
    measure: &PivotMeasure,
    layout: &XlsPivotCacheLayout,
    styles: &StyleTables,
) -> XlsResult<XlsDataFieldOptions> {
    let (show_as, base_field, base_item) = match &measure.show_as {
        PivotShowAs::Normal => (0, None, None),
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        } => (1, Some(base_field), Some(base_item)),
        PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => (3, Some(base_field), Some(base_item)),
        PivotShowAs::RunningTotal { base_field } => (4, Some(base_field), None),
        PivotShowAs::PercentOfRowTotal => (5, None, None),
        PivotShowAs::PercentOfColumnTotal => (6, None, None),
        PivotShowAs::PercentOfGrandTotal => (7, None, None),
        PivotShowAs::Index => (8, None, None),
        PivotShowAs::PercentOfParentRowTotal
        | PivotShowAs::PercentOfParentColumnTotal
        | PivotShowAs::PercentOfParentTotal { .. }
        | PivotShowAs::RankAscending { .. }
        | PivotShowAs::RankDescending { .. } => {
            return Err(XlsError::InvalidFormat(
                "XLS pivot show-as parent/rank variants are not supported by BIFF8 SXDI".into(),
            ))
        }
    };
    let base_field_index = if let Some(base_field) = base_field {
        layout
            .field_index(&base_field.name)
            .or_else(|| layout.axis_field_index(&base_field.name))
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "pivot show-as base field not found: {}",
                    base_field.name
                ))
            })?
    } else {
        0
    };
    let base_item_index = if let Some(base_item) = base_item {
        let field = layout.fields.get(base_field_index).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot show-as base field index out of range: {base_field_index}"
            ))
        })?;
        field
            .shared_items
            .iter()
            .position(|candidate| candidate == base_item)
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "pivot show-as base item not found in field {}: {base_item}",
                    field.name
                ))
            })?
    } else {
        0
    };

    Ok(XlsDataFieldOptions {
        show_as,
        base_field: checked_u16(base_field_index, "pivot show-as base field index")?,
        base_item: checked_u16(base_item_index, "pivot show-as base item index")?,
        num_format: checked_u16(
            pivot_measure_number_format_id(measure, styles)? as usize,
            "pivot measure number format",
        )?,
    })
}

fn pivot_measure_number_format_id(measure: &PivotMeasure, styles: &StyleTables) -> XlsResult<u32> {
    let Some(number_format) = measure.number_format.as_deref() else {
        return Ok(0);
    };
    if let Some(id) = builtin_number_format_id(number_format) {
        return Ok(id);
    }
    styles
        .custom_format_index(number_format)
        .map(u32::from)
        .ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot measure number format was not registered: {number_format}"
            ))
        })
}

fn write_sxli_collection(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
    tuples: &[Vec<u16>],
) -> XlsResult<()> {
    let mut body = Vec::new();
    if fields.is_empty() {
        write_sxli_item(&mut body, &[], false, false);
    } else {
        let calculated_tuples = calculated_item_axis_tuple_set(pivot, layout, fields)?;
        let has_calculated_tuples = !calculated_tuples.is_empty();
        for tuple in tuples {
            if calculated_tuples.contains(tuple) {
                write_sxli_calculated_item(&mut body, tuple);
            } else {
                write_sxli_item(&mut body, tuple, false, has_calculated_tuples);
            }
        }
        write_sxli_item(
            &mut body,
            &vec![0; expanded_axis_field_count(layout, fields)? as usize],
            true,
            has_calculated_tuples,
        );
    }
    write_biff_record(stream, 0x00B5, &body);
    Ok(())
}

fn calculated_item_axis_tuple_set(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<HashSet<Vec<u16>>> {
    let mut out = HashSet::new();
    if fields.len() != 1 {
        return Ok(out);
    }
    let axis_field_name = &fields[0].field.name;
    let Some(field_index) = layout.axis_field_index(axis_field_name) else {
        return Ok(out);
    };
    let Some(layout_field) = layout.fields.get(field_index) else {
        return Ok(out);
    };
    for item in pivot
        .calculated_items
        .iter()
        .filter(|item| item.field.name.eq_ignore_ascii_case(axis_field_name))
    {
        let Some(item_index) = layout_field
            .shared_items
            .iter()
            .position(|candidate| candidate == &item.item)
        else {
            continue;
        };
        out.insert(vec![checked_u16(
            item_index,
            "pivot calculated axis item index",
        )?]);
    }
    Ok(out)
}

fn write_sxli_item(
    body: &mut Vec<u8>,
    item_indexes: &[u16],
    grand_total: bool,
    calculated_item_collection: bool,
) {
    let item_type: u16 = if grand_total { 0x000D } else { 0x0000 };
    let item_flags: u16 = if grand_total { 0x0A00 } else { 0x0000 };
    let isxvi_mac = if calculated_item_collection {
        item_indexes.len() as u16
    } else {
        item_indexes.len().saturating_sub(1) as u16
    };
    write_sxli_structure(body, 0, item_type, isxvi_mac, item_flags, item_indexes);
}

fn write_sxli_calculated_item(body: &mut Vec<u8>, item_indexes: &[u16]) {
    write_sxli_structure(body, 0, 0, 0, 0x0001, item_indexes);
}

fn write_sxli_structure(
    body: &mut Vec<u8>,
    c_sic: u16,
    item_type: u16,
    isxvi_mac: u16,
    item_flags: u16,
    item_indexes: &[u16],
) {
    body.extend_from_slice(&c_sic.to_le_bytes());
    body.extend_from_slice(&item_type.to_le_bytes());
    body.extend_from_slice(&isxvi_mac.to_le_bytes());
    body.extend_from_slice(&item_flags.to_le_bytes());
    for item_index in item_indexes {
        body.extend_from_slice(&item_index.to_le_bytes());
    }
}

fn write_values_axis_sxli_collection(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
) -> XlsResult<()> {
    let mut body = Vec::new();
    for data_item in 0..pivot.measures.len() {
        let data_item = checked_u16(data_item, "pivot data item index")?;
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&(0x1000u16 | data_item.saturating_mul(2)).to_le_bytes());
        body.extend_from_slice(&data_item.to_le_bytes());
    }
    write_biff_record(stream, 0x00B5, &body);
    Ok(())
}

fn axis_line_count(fields: &[PivotField], tuples: &[Vec<u16>]) -> XlsResult<u16> {
    if fields.is_empty() {
        return Ok(1);
    }
    checked_u16(tuples.len().saturating_add(1), "pivot line count")
}

fn build_xls_axis_tuples(
    part: &FormatPivotTable,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    effective_columns: &[PivotField],
) -> XlsResult<XlsPivotAxisTuples> {
    let mut rows = if pivot.rows.is_empty() {
        Vec::new()
    } else if let Some(tuples) = &part.axis_tuples.rows {
        xls_planned_axis_tuples(cache, layout, &pivot.rows, tuples)?
    } else {
        axis_item_tuples(part, layout, &pivot.rows)?
    };
    append_calculated_item_axis_tuples(pivot, layout, &pivot.rows, &mut rows)?;

    let has_synthetic_consolidation_columns = effective_columns.len() != pivot.columns.len();
    let mut columns = if effective_columns.is_empty() {
        Vec::new()
    } else if !has_synthetic_consolidation_columns {
        if let Some(tuples) = &part.axis_tuples.columns {
            xls_planned_axis_tuples(cache, layout, effective_columns, tuples)?
        } else {
            axis_item_tuples(part, layout, effective_columns)?
        }
    } else {
        axis_item_tuples(part, layout, effective_columns)?
    };
    append_calculated_item_axis_tuples(pivot, layout, effective_columns, &mut columns)?;

    Ok(XlsPivotAxisTuples { rows, columns })
}

fn xls_planned_axis_tuples(
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
    tuples: &[Vec<u32>],
) -> XlsResult<Vec<Vec<u16>>> {
    let layout_indexes = expanded_axis_field_indexes(layout, fields)?;
    tuples
        .iter()
        .map(|tuple| {
            let mut source_position = 0usize;
            let mut layout_position = 0usize;
            let mut mapped = Vec::with_capacity(tuple.len());
            for field in fields {
                let cache_index = cache.field_index(&field.field.name).ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "pivot references unknown axis field {}",
                        field.field.name
                    ))
                })?;
                let Some(grouping) = &cache.fields[cache_index].grouping else {
                    mapped.push(checked_u16(
                        *tuple.get(source_position).unwrap_or(&0) as usize,
                        "pivot item index",
                    )?);
                    source_position += 1;
                    layout_position += 1;
                    continue;
                };
                for level in &grouping.levels {
                    let planned_id = *tuple.get(source_position).unwrap_or(&0);
                    let target_field_index =
                        *layout_indexes.get(layout_position).ok_or_else(|| {
                            XlsError::InvalidFormat(
                                "pivot grouped axis layout is incomplete".into(),
                            )
                        })?;
                    let native_id = level
                        .item_ids
                        .iter()
                        .position(|item_id| *item_id == planned_id)
                        .and_then(|row| {
                            layout.fields[target_field_index].item_ids.get(row).copied()
                        })
                        .unwrap_or(planned_id);
                    mapped.push(checked_u16(native_id as usize, "pivot item index")?);
                    source_position += 1;
                    layout_position += 1;
                }
                if matches!(grouping.definition, PivotGrouping::Manual { .. }) {
                    mapped.push(checked_u16(
                        *tuple.get(source_position).unwrap_or(&0) as usize,
                        "pivot item index",
                    )?);
                    source_position += 1;
                    layout_position += 1;
                }
            }
            Ok(mapped)
        })
        .collect()
}

fn axis_item_tuples(
    part: &FormatPivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<Vec<Vec<u16>>> {
    let indexes = expanded_axis_field_indexes(layout, fields)?;

    let mut seen = HashSet::new();
    let mut tuples = Vec::new();
    for row in visible_row_indexes(part.visible_rows.as_deref(), layout.row_count) {
        let tuple = indexes
            .iter()
            .map(|index| {
                let item_id = layout.fields[*index]
                    .item_ids
                    .get(row)
                    .copied()
                    .unwrap_or(0);
                checked_u16(item_id as usize, "pivot item index")
            })
            .collect::<XlsResult<Vec<_>>>()?;
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(tuples)
}

fn append_calculated_item_axis_tuples(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
    tuples: &mut Vec<Vec<u16>>,
) -> XlsResult<()> {
    if fields.len() != 1 {
        return Ok(());
    }
    let axis_field_name = &fields[0].field.name;
    let Some(field_index) = layout.axis_field_index(axis_field_name) else {
        return Ok(());
    };
    let Some(layout_field) = layout.fields.get(field_index) else {
        return Ok(());
    };
    let mut seen = tuples.iter().cloned().collect::<HashSet<_>>();
    for item in pivot
        .calculated_items
        .iter()
        .filter(|item| item.field.name.eq_ignore_ascii_case(axis_field_name))
    {
        let Some(item_index) = layout_field
            .shared_items
            .iter()
            .position(|candidate| candidate == &item.item)
        else {
            continue;
        };
        let tuple = vec![checked_u16(item_index, "pivot calculated axis item index")?];
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(())
}
fn expanded_axis_field_count(
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<u16> {
    checked_u16(
        expanded_axis_field_indexes(layout, fields)?.len(),
        "pivot axis field count",
    )
}

fn expanded_axis_field_indexes(
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<Vec<usize>> {
    let mut indexes = Vec::new();
    for field in fields {
        let field_index = layout
            .field_index(&field.field.name)
            .or_else(|| layout.axis_field_index(&field.field.name))
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "pivot references unknown axis field {}",
                    field.field.name
                ))
            })?;
        match layout.fields.get(field_index).map(|field| &field.kind) {
            Some(XlsPivotFieldKind::DateSource {
                derived_field_indexes,
                ..
            }) => indexes.extend(derived_field_indexes.iter().copied()),
            Some(XlsPivotFieldKind::ManualSource {
                derived_field_index,
            }) => {
                indexes.push(*derived_field_index);
                indexes.push(field_index);
            }
            _ => indexes.push(field_index),
        }
    }
    Ok(indexes)
}

fn visible_axis_field_count(layout: &XlsPivotCacheLayout, fields: &[PivotField]) -> XlsResult<u16> {
    let mut count = 0usize;
    for field in fields {
        let field_index = layout
            .field_index(&field.field.name)
            .or_else(|| layout.axis_field_index(&field.field.name))
            .ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "pivot references unknown axis field {}",
                    field.field.name
                ))
            })?;
        count = count.saturating_add(
            match layout.fields.get(field_index).map(|field| &field.kind) {
                Some(XlsPivotFieldKind::DateSource {
                    derived_field_indexes,
                    ..
                }) => derived_field_indexes.len(),
                _ => 1,
            },
        );
    }
    checked_u16(count, "pivot visible axis field count")
}

fn selected_page_item_index(
    pivot: &duke_sheets_core::PivotTable,
    field_name: &str,
    field: &XlsPivotFieldLayout,
) -> XlsResult<Option<u16>> {
    let Some(filter) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if filter_field.name.eq_ignore_ascii_case(field_name)
        )
    }) else {
        return Ok(None);
    };
    let PivotFilter::FieldItems { allowed_items, .. } = filter else {
        return Ok(None);
    };

    if allowed_items.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot page field {field_name} requires at least one selected item"
        )));
    }

    if allowed_items.len() > 1 {
        field_filter_hidden_item_indexes(pivot, field_name, field)?;
        return Ok(Some(0x7FFD));
    }

    let item = &allowed_items[0];
    let Some(index) = field
        .shared_items
        .iter()
        .position(|candidate| candidate == item)
    else {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot page field {field_name} selected item is not present in the cache"
        )));
    };
    Ok(Some(checked_u16(index, "pivot page item index")?))
}

fn field_filter_hidden_item_indexes(
    pivot: &duke_sheets_core::PivotTable,
    field_name: &str,
    field: &XlsPivotFieldLayout,
) -> XlsResult<HashSet<u16>> {
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if filter_field.name.eq_ignore_ascii_case(field_name)
        )
    }) else {
        return Ok(HashSet::new());
    };

    if pivot_axis_contains_field(&pivot.page_fields, field_name) && allowed_items.len() <= 1 {
        return Ok(HashSet::new());
    }

    let mut allowed_indexes = HashSet::with_capacity(allowed_items.len());
    for item in allowed_items {
        let Some(index) = field
            .shared_items
            .iter()
            .position(|candidate| candidate == item)
        else {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot field {field_name} selected item is not present in the cache"
            )));
        };
        allowed_indexes.insert(checked_u16(index, "pivot page item index")?);
    }

    Ok((0..field.shared_items.len())
        .filter_map(|index| {
            let index = checked_u16(index, "pivot page item index").ok()?;
            (!allowed_indexes.contains(&index)).then_some(index)
        })
        .collect())
}

fn xls_field_has_hidden_item_filter(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    field_index: usize,
) -> bool {
    let Some(filter_field_name) = xls_filter_field_name_for_hidden_items(layout, field_index)
    else {
        return false;
    };
    pivot.filters.iter().any(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(filter_field_name)
        )
    })
}

fn xls_filter_field_name_for_hidden_items<'a>(
    layout: &'a XlsPivotCacheLayout,
    field_index: usize,
) -> Option<&'a str> {
    let field = layout.fields.get(field_index)?;
    match field.kind {
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        }
        | XlsPivotFieldKind::ManualGroup {
            source_field_index, ..
        } => layout
            .fields
            .get(source_field_index)
            .map(|source| source.name.as_str()),
        XlsPivotFieldKind::DateSource { .. } | XlsPivotFieldKind::ManualSource { .. } => None,
        _ => Some(field.name.as_str()),
    }
}

fn default_page_item_index(field: &XlsPivotFieldLayout) -> u16 {
    match field.kind {
        XlsPivotFieldKind::ManualGroup { .. } => 0x7FFD,
        _ => 0xFFFF,
    }
}

fn page_field_area_size(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> (u32, u32) {
    let count = xls_effective_page_field_count(pivot, layout);
    if count == 0 {
        return (0, 0);
    }

    let wrap = pivot.layout.page_wrap as usize;
    let row_count = if wrap == 0 {
        count
    } else if pivot.layout.page_over_then_down {
        (count + wrap - 1) / wrap
    } else {
        wrap.min(count)
    };
    let col_count = if wrap == 0 {
        1
    } else if pivot.layout.page_over_then_down {
        wrap.min(count)
    } else {
        (count + row_count - 1) / row_count
    };
    (row_count as u32, col_count as u32)
}

fn xls_effective_column_fields(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> Vec<PivotField> {
    let mut fields = pivot.columns.clone();
    if let Some(field_index) = xls_synthetic_consolidation_column_field_index(pivot, layout) {
        let mut field = PivotField::new(layout.fields[field_index].name.clone());
        field.sort = PivotSort::None;
        fields.push(field);
    }
    fields
}

fn xls_synthetic_consolidation_column_field_index(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> Option<usize> {
    if !layout.is_consolidation {
        return None;
    }
    let field_index = layout.field_index("Column")?;
    if xls_axis_contains_layout_field(&pivot.rows, layout, field_index)
        || xls_axis_contains_layout_field(&pivot.columns, layout, field_index)
        || xls_axis_contains_layout_field(&pivot.page_fields, layout, field_index)
    {
        return None;
    }
    Some(field_index)
}

fn xls_effective_page_field_count(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> usize {
    pivot.page_fields.len() + xls_synthetic_consolidation_page_field_indexes(pivot, layout).len()
}

fn xls_synthetic_consolidation_page_field_indexes(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> Vec<usize> {
    if !layout.is_consolidation {
        return Vec::new();
    }
    (1usize..=4)
        .filter_map(|index| layout.field_index(&format!("Page{index}")))
        .filter(|field_index| {
            !xls_axis_contains_layout_field(&pivot.page_fields, layout, *field_index)
        })
        .collect()
}

fn checked_u16(value: usize, label: &str) -> XlsResult<u16> {
    if value > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u16)
}

fn checked_u8(value: usize, label: &str) -> XlsResult<u8> {
    if value > u8::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u8)
}

fn checked_i16(value: usize, label: &str) -> XlsResult<i16> {
    if value > i16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as i16)
}

fn checked_u32(value: usize, label: &str) -> XlsResult<u32> {
    if value > u32::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u32)
}

fn checked_biff8_row(value: u32, label: &str) -> XlsResult<u16> {
    if value > u16::MAX as u32 {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 row limits"
        )));
    }
    Ok(value as u16)
}

fn checked_biff8_col(value: u16, label: &str) -> XlsResult<u16> {
    if value > u8::MAX as u16 {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 column limits"
        )));
    }
    Ok(value)
}

fn xls_pivot_field_axis(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    field_index: usize,
) -> u16 {
    let Some(field) = layout.fields.get(field_index) else {
        return 0x0000;
    };
    if matches!(field.kind, XlsPivotFieldKind::ManualSource { .. })
        && pivot_axis_contains_field(&pivot.page_fields, &field.name)
        && !pivot_axis_contains_field(&pivot.rows, &field.name)
        && !pivot_axis_contains_field(&pivot.columns, &field.name)
    {
        return 0x0000;
    }
    let axis_field_index = match field.kind {
        XlsPivotFieldKind::DateSource { .. } => return 0x0000,
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        }
        | XlsPivotFieldKind::ManualGroup {
            source_field_index, ..
        } => source_field_index,
        _ => field_index,
    };
    if xls_axis_contains_layout_field(&pivot.rows, layout, axis_field_index) {
        0x0001
    } else if xls_axis_contains_layout_field(&pivot.columns, layout, axis_field_index)
        || xls_synthetic_consolidation_column_field_index(pivot, layout) == Some(axis_field_index)
    {
        0x0002
    } else if xls_axis_contains_layout_field(&pivot.page_fields, layout, axis_field_index)
        || xls_synthetic_consolidation_page_field_indexes(pivot, layout).contains(&axis_field_index)
    {
        0x0004
    } else if pivot
        .measures
        .iter()
        .any(|measure| layout.field_index(&measure.field.name) == Some(axis_field_index))
    {
        0x0008
    } else {
        0x0000
    }
}

fn xls_axis_contains_layout_field(
    fields: &[PivotField],
    layout: &XlsPivotCacheLayout,
    field_index: usize,
) -> bool {
    fields.iter().any(|field| {
        layout.field_index(&field.field.name) == Some(field_index)
            || layout.axis_field_index(&field.field.name) == Some(field_index)
    })
}

fn has_grouped_page_axis(
    layout: &XlsPivotCacheLayout,
    pivot: &duke_sheets_core::PivotTable,
) -> bool {
    pivot.page_fields.iter().any(|field| {
        layout
            .page_axis_field_index(&field.field.name)
            .and_then(|index| layout.fields.get(index))
            .is_some_and(|field| {
                matches!(
                    field.kind,
                    XlsPivotFieldKind::DateGroup { .. } | XlsPivotFieldKind::ManualGroup { .. }
                )
            })
    })
}

fn xls_cache_has_calculated_field(layout: &XlsPivotCacheLayout) -> bool {
    layout.fields.iter().any(|field| field.formula.is_some())
}

fn calculated_item_indexes_for_xls_field(
    cache: &FormatPivotCache,
    field_name: &str,
    shared_items: &[PivotValue],
) -> HashSet<usize> {
    cache
        .calculated_items
        .iter()
        .filter(|item| item.field.name.eq_ignore_ascii_case(field_name))
        .filter_map(|item| shared_items.iter().position(|value| value == &item.item))
        .collect()
}

fn xls_pivot_has_sxaddl_filter(pivot: &duke_sheets_core::PivotTable) -> bool {
    pivot.filters.iter().any(|filter| match filter {
        PivotFilter::TopN { percent: true, .. } => true,
        PivotFilter::Label { operator, .. } => xls_supported_label_filter_operator(*operator),
        PivotFilter::LabelBetween { .. } => true,
        PivotFilter::Value { operator, .. } => xls_supported_value_filter_operator(*operator),
        PivotFilter::ValueBetween { .. } => true,
        PivotFilter::Date { operator, .. } => xls_supported_date_filter_operator(*operator),
        PivotFilter::DateBetween { .. } | PivotFilter::DatePeriod { .. } => true,
        _ => false,
    })
}

fn xls_supported_label_filter_operator(operator: PivotFilterOperator) -> bool {
    matches!(
        operator,
        PivotFilterOperator::Equals
            | PivotFilterOperator::NotEquals
            | PivotFilterOperator::LessThan
            | PivotFilterOperator::LessThanOrEqual
            | PivotFilterOperator::GreaterThan
            | PivotFilterOperator::GreaterThanOrEqual
            | PivotFilterOperator::BeginsWith
            | PivotFilterOperator::DoesNotBeginWith
            | PivotFilterOperator::EndsWith
            | PivotFilterOperator::DoesNotEndWith
            | PivotFilterOperator::Contains
            | PivotFilterOperator::DoesNotContain
    )
}

fn xls_supported_value_filter_operator(operator: PivotFilterOperator) -> bool {
    matches!(
        operator,
        PivotFilterOperator::Equals
            | PivotFilterOperator::NotEquals
            | PivotFilterOperator::LessThan
            | PivotFilterOperator::LessThanOrEqual
            | PivotFilterOperator::GreaterThan
            | PivotFilterOperator::GreaterThanOrEqual
    )
}

fn xls_supported_date_filter_operator(operator: PivotFilterOperator) -> bool {
    xls_date_filter_type_and_operator(operator).is_some()
}

fn xls_values_field_on_columns(pivot: &duke_sheets_core::PivotTable) -> bool {
    pivot.measures.len() > 1
        && pivot.columns.is_empty()
        && matches!(pivot.layout.values_axis, PivotValuesAxis::Columns)
}

fn xls_pivot_aggregate_code(aggregate: PivotAggregate) -> u16 {
    match aggregate {
        PivotAggregate::Sum => 0,
        PivotAggregate::Count => 1,
        PivotAggregate::Average => 2,
        PivotAggregate::Max => 3,
        PivotAggregate::Min => 4,
        PivotAggregate::Product => 5,
        PivotAggregate::CountNumbers => 6,
        PivotAggregate::StdDev => 7,
        PivotAggregate::StdDevP => 8,
        PivotAggregate::Var => 9,
        PivotAggregate::VarP => 10,
    }
}

fn write_pivot_sheet_tail_records(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    pivot_plan: &FormatPivotPlan,
    sheet_idx: usize,
) {
    for _part in pivot_plan
        .tables
        .iter()
        .filter(|part| part.sheet_index == sheet_idx)
    {
        write_biff_record(
            stream,
            0x088B,
            &[
                0x8B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x12, 0x00,
            ],
        );
        if sheet.selections().is_empty() {
            write_default_selection_record(stream);
        } else {
            write_selection_records(stream, sheet);
        }
        write_biff_record(
            stream,
            0x0867,
            &[
                0x67, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
                0x01, 0xFF, 0xFF, 0xFF, 0xFF, 0x03, 0x44, 0x00, 0x00,
            ],
        );
    }
}

fn write_default_selection_record(stream: &mut Vec<u8>) {
    write_biff_record(
        stream,
        SELECTION_RECORD,
        &[
            0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00,
        ],
    );
}

fn write_sxview_record(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    has_expanded_row_axis: bool,
) -> XlsResult<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&[0x02, 0x08, 0x00, 0x00]);
    body.extend_from_slice(&1u16.to_le_bytes());
    body.extend_from_slice(&2u16.to_le_bytes());
    let future_options = if matches!(cache.source, FormatPivotSource::External { .. }) {
        2u32
    } else {
        0u32
    };
    body.extend_from_slice(&future_options.to_le_bytes());
    body.extend_from_slice(&[0x08, 0x03, 0x10, 0x00]);
    push_xlunicode_string(&mut body, &pivot.name)?;
    body.push(0);
    body.push(if has_expanded_row_axis { 2 } else { 1 });
    write_biff_record(stream, 0x0802, &body);
    Ok(())
}

fn write_pivot_frt_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    has_expanded_row_axis: bool,
    date_system: DateSystem,
) -> XlsResult<()> {
    write_frt0864_name(stream, 0x0000, &pivot.name)?;
    if has_expanded_row_axis {
        write_frt0864_raw(stream, &[0x00, 0x02, 0x08, 0x41, 0x40, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x00, 0x19, 0x9F, 0x00, 0x40, 0x00, 0x00, 0x00]);
    } else {
        write_frt0864_raw(stream, &[0x00, 0x02, 0x00, 0x41, 0x40, 0x01, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x00, 0x19, 0x19, 0x00, 0x40, 0x01, 0x00, 0x00]);
    }
    for (field_index, field) in layout.fields.iter().enumerate() {
        if xls_axis_field_for_layout_field(pivot, layout, field_index)
            .is_some_and(|axis_field| !axis_field.show_drop_downs)
        {
            write_frt0864_field_ver10_info(stream, &field.name, 0x0000_0001)?;
        }
    }
    for (field_index, field) in layout.fields.iter().enumerate() {
        write_frt0864_name(stream, 0x0017, &field.name)?;
        let axis_field = xls_axis_field_for_layout_field(pivot, layout, field_index);
        let include_new_items =
            axis_field.is_some_and(|axis_field| axis_field.include_new_items_in_filter);
        let field_flags = if include_new_items { 0x08u32 } else { 0x28u32 };
        let mut tail = Vec::with_capacity(8);
        tail.extend_from_slice(&[0x17, 0x19]);
        tail.extend_from_slice(&field_flags.to_le_bytes());
        tail.extend_from_slice(&0u16.to_le_bytes());
        write_frt0864_raw(stream, &tail);
        if let Some(filter) = xls_sxvd_top_n_filter(pivot, axis_field)? {
            write_frt0864_field_autoshow_count(stream, filter.n);
        }
        write_frt0864_raw(stream, &[0x17, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x17, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x17, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    }
    write_frt0864_style(stream, &pivot.style)?;
    write_frt0864_raw(stream, &[0x00, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00]);
    let extension_filters = xls_sxaddl_pivot_filters(pivot, layout)?;
    for (index, filter) in extension_filters.iter().enumerate() {
        write_frt0864_raw(stream, &[0x1C, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x1D, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_pivot_filter_collection(stream, filter)?;
        match filter {
            XlsSxaddlPivotFilter::TopN(filter) => {
                write_frt0864_pivot_top_n_filter(stream, filter, index + 1)?;
            }
            XlsSxaddlPivotFilter::Label(filter) => {
                write_frt0864_pivot_label_filter(stream, filter, index + 1)?;
            }
            XlsSxaddlPivotFilter::Value(filter) => {
                write_frt0864_pivot_value_filter(stream, filter, index + 1)?;
            }
            XlsSxaddlPivotFilter::Date(filter) => {
                write_frt0864_pivot_date_filter(stream, filter, index + 1, date_system)?;
            }
        }
        write_frt0864_raw(stream, &[0x1D, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x1C, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    }
    write_frt0864_raw(stream, &[0x00, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00]);
    write_frt0864_raw(stream, &[0x00, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    Ok(())
}

fn write_frt0864_field_ver10_info(stream: &mut Vec<u8>, name: &str, flags: u32) -> XlsResult<()> {
    write_frt0864_name(stream, 0x0001, name)?;
    let mut tail = Vec::with_capacity(8);
    tail.extend_from_slice(&[0x01, 0x02]);
    tail.extend_from_slice(&flags.to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    write_frt0864_raw(stream, &tail);
    write_frt0864_raw(stream, &[0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    Ok(())
}

fn write_frt0864_field_autoshow_count(stream: &mut Vec<u8>, count: u32) {
    let mut tail = Vec::with_capacity(8);
    tail.extend_from_slice(&[0x17, 0x37]);
    tail.extend_from_slice(&count.to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    write_frt0864_raw(stream, &tail);
}

fn write_frt0864_pivot_filter_collection(
    stream: &mut Vec<u8>,
    filter: &XlsSxaddlPivotFilter,
) -> XlsResult<()> {
    let field_index = checked_u32(
        filter.field_source_index(),
        "pivot extension filter field index",
    )?;
    let mut tail = Vec::with_capacity(32);
    tail.extend_from_slice(&[0x1D, 0x38]);
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    tail.extend_from_slice(&field_index.to_le_bytes());
    tail.extend_from_slice(&(-1i32).to_le_bytes());
    tail.extend_from_slice(&filter.collection_filter_type().to_le_bytes());
    tail.extend_from_slice(&(-1i32).to_le_bytes());
    tail.extend_from_slice(&filter.collection_measure_index().to_le_bytes());
    tail.extend_from_slice(&filter.collection_trailing_sentinel().to_le_bytes());
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_pivot_top_n_filter(
    stream: &mut Vec<u8>,
    filter: &XlsSxaddlTopNFilter,
    filter_id: usize,
) -> XlsResult<()> {
    let filter_type = match (filter.top, filter.percent) {
        (true, true) => 3u32,
        (false, true) => 4u32,
        (true, false) => 1u32,
        (false, false) => 2u32,
    };
    let field_index = checked_u32(
        filter.field_source_index + 1,
        "pivot extension filter field index",
    )?;
    let measure_index = checked_u32(
        filter.measure_source_index + 1,
        "pivot extension filter measure index",
    )?;
    let filter_id = checked_u32(filter_id, "pivot extension filter id")?;

    let mut tail = Vec::with_capacity(44);
    tail.extend_from_slice(&[0x1D, 0x3C]);
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    tail.extend_from_slice(&filter_type.to_le_bytes());
    tail.extend_from_slice(&field_index.to_le_bytes());
    tail.extend_from_slice(&measure_index.to_le_bytes());
    tail.extend_from_slice(&filter_id.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    push_pivot_xnum(&mut tail, filter.n as f64);
    tail.extend_from_slice(&0f64.to_le_bytes());
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_pivot_label_filter(
    stream: &mut Vec<u8>,
    filter: &XlsSxaddlLabelFilter,
    filter_id: usize,
) -> XlsResult<()> {
    match &filter.kind {
        XlsSxaddlLabelFilterKind::Comparison { operator, value } => {
            write_frt0864_pivot_filter_string(stream, 0x3A, value)?;

            let filter_id = checked_u32(filter_id, "pivot extension filter id")?;
            let (custom_operator_code, discriminator) = match operator {
                PivotFilterOperator::Equals => (0x02, 0x01),
                PivotFilterOperator::BeginsWith
                | PivotFilterOperator::EndsWith
                | PivotFilterOperator::Contains => (0x02, 0x00),
                PivotFilterOperator::NotEquals => (0x05, 0x01),
                PivotFilterOperator::DoesNotBeginWith
                | PivotFilterOperator::DoesNotEndWith
                | PivotFilterOperator::DoesNotContain => (0x05, 0x00),
                PivotFilterOperator::LessThan => (0x01, 0x01),
                PivotFilterOperator::LessThanOrEqual => (0x03, 0x01),
                PivotFilterOperator::GreaterThan => (0x04, 0x01),
                PivotFilterOperator::GreaterThanOrEqual => (0x06, 0x01),
            };
            let mut tail = Vec::with_capacity(44);
            tail.extend_from_slice(&[0x1D, 0x3C]);
            tail.extend_from_slice(&0u16.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
            tail.extend_from_slice(&filter_id.to_le_bytes());
            tail.extend_from_slice(&[0x06, custom_operator_code, discriminator, 0x00]);
            tail.extend_from_slice(&[0u8; 24]);
            write_frt0864_raw(stream, &tail);

            let criterion = xls_label_filter_criterion(*operator, value);
            write_frt0864_pivot_filter_string(stream, 0x3D, &criterion)?;
        }
        XlsSxaddlLabelFilterKind::Between {
            start,
            end,
            not_between,
        } => {
            write_frt0864_pivot_filter_string(stream, 0x3A, start)?;
            write_frt0864_pivot_filter_string(stream, 0x3B, end)?;
            let mut tail = Vec::with_capacity(44);
            tail.extend_from_slice(&[0x1D, 0x3C]);
            tail.extend_from_slice(&0u16.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
            tail.extend_from_slice(&2u32.to_le_bytes());
            if *not_between {
                tail.extend_from_slice(&[0x06, 0x01, 0x01, 0x00]);
                tail.extend_from_slice(&[0u8; 6]);
                tail.extend_from_slice(&[0x06, 0x04, 0x01, 0x00]);
                tail.extend_from_slice(&[0u8; 6]);
                tail.extend_from_slice(&2u32.to_le_bytes());
            } else {
                tail.extend_from_slice(&[0x06, 0x06, 0x01, 0x00]);
                tail.extend_from_slice(&[0u8; 6]);
                tail.extend_from_slice(&[0x06, 0x03, 0x01, 0x00]);
                tail.extend_from_slice(&[0u8; 6]);
                tail.extend_from_slice(&1u32.to_le_bytes());
            }
            tail.extend_from_slice(&0u32.to_le_bytes());
            write_frt0864_raw(stream, &tail);
            write_frt0864_pivot_filter_string(stream, 0x3D, start)?;
            write_frt0864_pivot_filter_string(stream, 0x3E, end)?;
        }
    }
    Ok(())
}

fn write_frt0864_pivot_value_filter(
    stream: &mut Vec<u8>,
    filter: &XlsSxaddlValueFilter,
    filter_id: usize,
) -> XlsResult<()> {
    let mut tail = Vec::with_capacity(48);
    tail.extend_from_slice(&[0x1D, 0x3C]);
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    match filter.kind {
        XlsSxaddlValueFilterKind::Comparison { operator, value } => {
            let filter_id = checked_u32(filter_id, "pivot extension filter id")?;
            tail.extend_from_slice(&filter_id.to_le_bytes());
            let (_, custom_operator) = value_filter_type_and_operator(operator);
            write_frt0864_numeric_filter_criterion(&mut tail, custom_operator, value);
            tail.extend_from_slice(&[0u8; 18]);
        }
        XlsSxaddlValueFilterKind::Between {
            start,
            end,
            not_between: false,
        } => {
            tail.extend_from_slice(&2u32.to_le_bytes());
            write_frt0864_numeric_filter_criterion(&mut tail, 0x06, start);
            write_frt0864_numeric_filter_criterion(&mut tail, 0x03, end);
            tail.extend_from_slice(&1u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
        }
        XlsSxaddlValueFilterKind::Between {
            start,
            end,
            not_between: true,
        } => {
            tail.extend_from_slice(&2u32.to_le_bytes());
            write_frt0864_numeric_filter_criterion(&mut tail, 0x01, start);
            write_frt0864_numeric_filter_criterion(&mut tail, 0x04, end);
            tail.extend_from_slice(&2u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
        }
    }
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_pivot_date_filter(
    stream: &mut Vec<u8>,
    filter: &XlsSxaddlDateFilter,
    _filter_id: usize,
    date_system: DateSystem,
) -> XlsResult<()> {
    let mut tail = Vec::with_capacity(48);
    tail.extend_from_slice(&[0x1D, 0x3C]);
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    match filter.kind {
        XlsSxaddlDateFilterKind::Comparison { operator, value } => {
            let (_, custom_operator, date_filter_type) =
                xls_date_filter_type_and_operator(operator).expect("validated date operator");
            tail.extend_from_slice(&date_filter_type.to_le_bytes());
            tail.extend_from_slice(&1u32.to_le_bytes());
            write_frt0864_numeric_filter_criterion(&mut tail, custom_operator, value);
            tail.extend_from_slice(&[0u8; 18]);
        }
        XlsSxaddlDateFilterKind::Between {
            start,
            end,
            not_between: false,
        } => {
            tail.extend_from_slice(&7u32.to_le_bytes());
            tail.extend_from_slice(&2u32.to_le_bytes());
            write_frt0864_numeric_filter_criterion(&mut tail, 0x06, start);
            write_frt0864_numeric_filter_criterion(&mut tail, 0x03, end);
            tail.extend_from_slice(&1u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
        }
        XlsSxaddlDateFilterKind::Between {
            start,
            end,
            not_between: true,
        } => {
            tail.extend_from_slice(&43u32.to_le_bytes());
            tail.extend_from_slice(&2u32.to_le_bytes());
            write_frt0864_numeric_filter_criterion(&mut tail, 0x01, start);
            write_frt0864_numeric_filter_criterion(&mut tail, 0x04, end);
            tail.extend_from_slice(&2u32.to_le_bytes());
            tail.extend_from_slice(&0u32.to_le_bytes());
        }
        XlsSxaddlDateFilterKind::Period(period) => {
            let (_, cft) = xls_date_period_filter_codes(period).expect("validated date period");
            tail.extend_from_slice(&cft.to_le_bytes());
            if let Some((start, end)) = pivot_date_period_filter_bounds(period, date_system) {
                tail.extend_from_slice(&2u32.to_le_bytes());
                write_frt0864_numeric_filter_criterion(&mut tail, 0x06, start);
                write_frt0864_numeric_filter_criterion(&mut tail, 0x01, end);
                tail.extend_from_slice(&1u32.to_le_bytes());
                tail.extend_from_slice(&0u32.to_le_bytes());
            } else {
                tail.extend_from_slice(&0u32.to_le_bytes());
                tail.extend_from_slice(&[0u8; 28]);
            }
        }
    }
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_numeric_filter_criterion(out: &mut Vec<u8>, operator: u8, value: f64) {
    out.push(0x04);
    out.push(operator);
    out.extend_from_slice(&value.to_le_bytes());
}

fn value_filter_type_and_operator(operator: PivotFilterOperator) -> (u32, u8) {
    match operator {
        PivotFilterOperator::Equals => (18, 0x02),
        PivotFilterOperator::NotEquals => (19, 0x05),
        PivotFilterOperator::GreaterThan => (20, 0x04),
        PivotFilterOperator::GreaterThanOrEqual => (21, 0x06),
        PivotFilterOperator::LessThan => (22, 0x01),
        PivotFilterOperator::LessThanOrEqual => (23, 0x03),
        _ => unreachable!("unsupported XLS value filter operator"),
    }
}

fn xls_date_filter_type_and_operator(operator: PivotFilterOperator) -> Option<(u32, u8, u32)> {
    Some(match operator {
        PivotFilterOperator::Equals => (26, 0x02, 4),
        PivotFilterOperator::NotEquals => (62, 0x05, 40),
        PivotFilterOperator::LessThan => (27, 0x01, 5),
        PivotFilterOperator::LessThanOrEqual => (63, 0x03, 41),
        PivotFilterOperator::GreaterThan => (28, 0x04, 6),
        PivotFilterOperator::GreaterThanOrEqual => (64, 0x06, 42),
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => return None,
    })
}

fn xls_date_period_filter_codes(period: PivotDatePeriod) -> Option<(u32, u32)> {
    let cft = match period {
        PivotDatePeriod::Tomorrow => 0x08,
        PivotDatePeriod::Today => 0x09,
        PivotDatePeriod::Yesterday => 0x0A,
        PivotDatePeriod::NextWeek => 0x0B,
        PivotDatePeriod::ThisWeek => 0x0C,
        PivotDatePeriod::LastWeek => 0x0D,
        PivotDatePeriod::NextMonth => 0x0E,
        PivotDatePeriod::ThisMonth => 0x0F,
        PivotDatePeriod::LastMonth => 0x10,
        PivotDatePeriod::NextQuarter => 0x11,
        PivotDatePeriod::ThisQuarter => 0x12,
        PivotDatePeriod::LastQuarter => 0x13,
        PivotDatePeriod::NextYear => 0x14,
        PivotDatePeriod::ThisYear => 0x15,
        PivotDatePeriod::LastYear => 0x16,
        PivotDatePeriod::YearToDate => 0x17,
        PivotDatePeriod::Quarter(1) => 0x18,
        PivotDatePeriod::Quarter(2) => 0x19,
        PivotDatePeriod::Quarter(3) => 0x1A,
        PivotDatePeriod::Quarter(4) => 0x1B,
        PivotDatePeriod::Month(1) => 0x1C,
        PivotDatePeriod::Month(2) => 0x1D,
        PivotDatePeriod::Month(3) => 0x1E,
        PivotDatePeriod::Month(4) => 0x1F,
        PivotDatePeriod::Month(5) => 0x20,
        PivotDatePeriod::Month(6) => 0x21,
        PivotDatePeriod::Month(7) => 0x22,
        PivotDatePeriod::Month(8) => 0x23,
        PivotDatePeriod::Month(9) => 0x24,
        PivotDatePeriod::Month(10) => 0x25,
        PivotDatePeriod::Month(11) => 0x26,
        PivotDatePeriod::Month(12) => 0x27,
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => return None,
    };
    Some((cft + 22, cft))
}

fn xls_label_filter_criterion(operator: PivotFilterOperator, value: &str) -> String {
    match operator {
        PivotFilterOperator::BeginsWith | PivotFilterOperator::DoesNotBeginWith => {
            format!("{value}*")
        }
        PivotFilterOperator::EndsWith | PivotFilterOperator::DoesNotEndWith => {
            format!("*{value}")
        }
        PivotFilterOperator::Contains | PivotFilterOperator::DoesNotContain => {
            format!("*{value}*")
        }
        _ => value.to_string(),
    }
}

fn write_frt0864_pivot_filter_string(
    stream: &mut Vec<u8>,
    subtype: u8,
    value: &str,
) -> XlsResult<()> {
    let mut tail = Vec::new();
    tail.extend_from_slice(&[0x1D, subtype]);
    tail.extend_from_slice(&xlunicode_len_u16(value)?.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    push_xlunicode_string(&mut tail, value)?;
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn push_pivot_xnum(out: &mut Vec<u8>, value: f64) {
    let bytes = value.to_le_bytes();
    out.extend_from_slice(&bytes[6..8]);
    out.extend_from_slice(&bytes[4..6]);
    out.extend_from_slice(&bytes[2..4]);
    out.extend_from_slice(&bytes[0..2]);
}

fn write_frt0864_name(stream: &mut Vec<u8>, subtype: u16, name: &str) -> XlsResult<()> {
    let mut tail = Vec::new();
    tail.extend_from_slice(&subtype.to_le_bytes());
    tail.extend_from_slice(&(name.encode_utf16().count() as u32).to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    push_xlunicode_string(&mut tail, name)?;
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_style(
    stream: &mut Vec<u8>,
    style: &duke_sheets_core::PivotStyle,
) -> XlsResult<()> {
    let style_name = style
        .name
        .as_deref()
        .filter(|name| !name.trim().is_empty())
        .unwrap_or("PivotStyleMedium9");
    let mut tail = Vec::new();
    tail.extend_from_slice(&0x1E00u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&pivot_style_flags(style).to_le_bytes());
    tail.extend_from_slice(
        &checked_u16(style_name.encode_utf16().count(), "pivot style name length")?.to_le_bytes(),
    );
    for unit in style_name.encode_utf16() {
        tail.extend_from_slice(&unit.to_le_bytes());
    }
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn pivot_style_flags(style: &duke_sheets_core::PivotStyle) -> u16 {
    let mut flags = 0u16;
    if style.show_last_column {
        flags |= 0x02;
    }
    if style.show_row_stripes {
        flags |= 0x04;
    }
    if style.show_column_stripes {
        flags |= 0x08;
    }
    if style.show_row_headers {
        flags |= 0x10;
    }
    if style.show_column_headers {
        flags |= 0x20;
    }
    flags
}

fn write_frt0864_raw(stream: &mut Vec<u8>, tail: &[u8]) {
    let mut body = Vec::with_capacity(4 + tail.len());
    body.extend_from_slice(&[0x64, 0x08, 0x00, 0x00]);
    body.extend_from_slice(tail);
    write_biff_record(stream, 0x0864, &body);
}

fn pivot_value_string_payload(value: &PivotValue) -> XlsResult<Vec<u8>> {
    let text = match value {
        PivotValue::Blank => String::new(),
        PivotValue::Boolean(value) => {
            if *value {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        PivotValue::Number(value) => value.to_string(),
        PivotValue::String(value) => value.clone(),
        PivotValue::Error(value) => value.to_string(),
    };
    let mut body = Vec::new();
    push_xlunicode_string(&mut body, &text)?;
    Ok(body)
}

fn field_is_numeric(field: &XlsPivotFieldLayout) -> bool {
    !field.shared_items.is_empty()
        && field
            .shared_items
            .iter()
            .all(|value| matches!(value, PivotValue::Number(_)))
}

