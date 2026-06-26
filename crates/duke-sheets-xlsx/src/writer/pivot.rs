use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, Event};

use crate::styles::XlsxStyleTable;
use duke_sheets_core::{
    CellAddress, CellError, CellRange, PivotAggregate, PivotCalculatedField, PivotDateGroupUnit,
    PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotShowAs, PivotSort, PivotSource, PivotSourceRange,
    PivotSubtotal, PivotTable, PivotValue, Table, Workbook, WorkbookConnection,
    WorkbookConnectionKind, Worksheet,
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

use super::{
    write_xml_part, XlsxError, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS,
    NS_SPREADSHEET, RT_PIVOT_CACHE_DEFINITION, RT_PIVOT_CACHE_RECORDS,
};

const NS_SPREADSHEET_X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXT_URI_X14_DATA_FIELD: &str = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}";
#[cfg(feature = "parallel")]
const PARALLEL_CACHE_ROW_THRESHOLD: usize = 50_000;

#[derive(Debug, Clone)]
pub(super) struct PivotNumbering {
    pub(super) cache_parts: Vec<PivotCachePart>,
    pub(super) table_parts: Vec<PivotTablePart>,
}

#[derive(Debug, Clone)]
pub(super) struct PivotCachePart {
    pub(super) cache_num: usize,
    source: PivotSource,
    source_sheet_index: usize,
    fields: Vec<CacheField>,
    rows: Vec<Vec<Option<u32>>>,
    record_count: usize,
    save_data: bool,
    refresh_on_load: bool,
    background_query: bool,
    missing_items_limit: Option<u32>,
}

#[derive(Debug, Clone)]
pub(super) struct PivotTablePart {
    pub(super) sheet_index: usize,
    pub(super) pivot_index: usize,
    pub(super) table_num: usize,
    pub(super) cache_num: usize,
}

#[derive(Debug, Clone)]
struct ResolvedPivotSource {
    key: String,
    source: PivotSource,
    source_sheet_index: usize,
    fields: Vec<CacheField>,
    rows: Vec<Vec<Option<u32>>>,
    record_count: usize,
    save_data: bool,
}

#[derive(Debug, Clone)]
struct CacheField {
    name: String,
    formula: Option<String>,
    database_field: bool,
    metadata_only_shared_items: bool,
    group: Option<CacheFieldGroup>,
    shared_items: Vec<PivotValue>,
    item_lookup: HashMap<PivotValue, u32>,
}

#[derive(Debug, Clone)]
enum CacheFieldGroup {
    Base {
        parent: usize,
    },
    Range(PivotGrouping),
    Manual {
        base: usize,
        item_indexes: Vec<u32>,
        group_items: Vec<PivotValue>,
    },
    DateUnit {
        base: usize,
        parent: Option<usize>,
        unit: PivotDateGroupUnit,
    },
}

impl CacheField {
    fn new(name: String) -> Self {
        Self {
            name,
            formula: None,
            database_field: true,
            metadata_only_shared_items: false,
            group: None,
            shared_items: Vec::new(),
            item_lookup: HashMap::new(),
        }
    }

    fn calculated(name: String, formula: String) -> Self {
        Self {
            name,
            formula: Some(formula),
            database_field: false,
            metadata_only_shared_items: false,
            group: None,
            shared_items: Vec::new(),
            item_lookup: HashMap::new(),
        }
    }

    fn intern(&mut self, value: PivotValue) -> u32 {
        if let Some(index) = self.item_lookup.get(&value) {
            return *index;
        }

        let index = self.shared_items.len() as u32;
        self.shared_items.push(value.clone());
        self.item_lookup.insert(value, index);
        index
    }
}

pub(super) fn workbook_cache_rid(_workbook: &Workbook, cache_num: usize) -> String {
    format!("rIdPivotCache{}", cache_num)
}

pub(super) fn build_pivot_numbering(workbook: &Workbook) -> XlsxResult<PivotNumbering> {
    let mut cache_by_source: BTreeMap<String, usize> = BTreeMap::new();
    let mut cache_parts: Vec<PivotCachePart> = Vec::new();
    let mut table_parts: Vec<PivotTablePart> = Vec::new();

    for (sheet_index, sheet) in workbook.worksheets().enumerate() {
        for (pivot_index, pivot) in sheet.pivot_tables().iter().enumerate() {
            validate_writable_pivot(pivot)?;

            let mut resolved = resolve_pivot_source(workbook, sheet_index, pivot)?;
            apply_calculated_cache_fields(&pivot.name, &mut resolved, &pivot.calculated_fields)?;
            validate_pivot_fields(pivot, &resolved.fields)?;
            validate_pivot_groupings(pivot, &resolved.fields)?;
            apply_grouped_cache_fields(
                &pivot.name,
                &mut resolved,
                &pivot.groupings,
                workbook.settings().date_1904,
            )?;
            mark_metadata_only_measure_fields(pivot, &mut resolved.fields);
            let cache_key = cache_key_for_pivot(&resolved.key, pivot);

            let cache_num = if let Some(cache_num) = cache_by_source.get(&cache_key) {
                if let Some(cache_part) = cache_parts.get_mut(*cache_num - 1) {
                    merge_cache_field_usage(&mut cache_part.fields, &resolved.fields);
                }
                if pivot.refresh_policy.refresh_on_open {
                    if let Some(cache_part) = cache_parts.get_mut(*cache_num - 1) {
                        cache_part.refresh_on_load = true;
                    }
                }
                if pivot.refresh_policy.background_query {
                    if let Some(cache_part) = cache_parts.get_mut(*cache_num - 1) {
                        cache_part.background_query = true;
                    }
                }
                *cache_num
            } else {
                let cache_num = cache_parts.len() + 1;
                cache_by_source.insert(cache_key, cache_num);
                cache_parts.push(PivotCachePart {
                    cache_num,
                    source: resolved.source,
                    source_sheet_index: resolved.source_sheet_index,
                    fields: resolved.fields,
                    rows: resolved.rows,
                    record_count: resolved.record_count,
                    save_data: resolved.save_data,
                    refresh_on_load: pivot.refresh_policy.refresh_on_open,
                    background_query: pivot.refresh_policy.background_query,
                    missing_items_limit: pivot.refresh_policy.missing_items_limit,
                });
                cache_num
            };

            table_parts.push(PivotTablePart {
                sheet_index,
                pivot_index,
                table_num: table_parts.len() + 1,
                cache_num,
            });
        }
    }

    Ok(PivotNumbering {
        cache_parts,
        table_parts,
    })
}

fn validate_writable_pivot(pivot: &PivotTable) -> XlsxResult<()> {
    if pivot.measures.is_empty() {
        return Err(XlsxError::InvalidFormat(format!(
            "pivot table {} has no measures",
            pivot.name
        )));
    }

    if pivot
        .measures
        .iter()
        .any(|measure| !is_writable_show_as(&measure.show_as))
    {
        return Err(XlsxError::InvalidFormat(format!(
            "pivot table {} uses show-as calculations that are not written yet",
            pivot.name
        )));
    }

    if pivot
        .filters
        .iter()
        .any(|filter| !is_writable_filter(filter))
    {
        return Err(XlsxError::InvalidFormat(format!(
            "pivot table {} uses a filter type that is not written yet",
            pivot.name
        )));
    }

    for filter in &pivot.filters {
        match filter {
            PivotFilter::Value { value, .. } if !value.is_finite() => {
                return Err(XlsxError::InvalidFormat(format!(
                    "pivot table {} uses a non-finite value filter operand",
                    pivot.name
                )));
            }
            PivotFilter::TopN { n, .. } if *n == 0 => {
                return Err(XlsxError::InvalidFormat(format!(
                    "pivot table {} uses a top-N filter with a zero threshold",
                    pivot.name
                )));
            }
            _ => {}
        }
    }

    Ok(())
}

fn is_writable_filter(filter: &PivotFilter) -> bool {
    match filter {
        PivotFilter::FieldItems { .. } | PivotFilter::Label { .. } | PivotFilter::TopN { .. } => {
            true
        }
        PivotFilter::Value { operator, .. } => value_filter_type_name(*operator).is_some(),
        PivotFilter::Unsupported { .. } => false,
    }
}

fn validate_pivot_fields(pivot: &PivotTable, fields: &[CacheField]) -> XlsxResult<()> {
    for field in pivot
        .rows
        .iter()
        .map(|field| &field.field)
        .chain(pivot.columns.iter().map(|field| &field.field))
        .chain(pivot.page_fields.iter().map(|field| &field.field))
        .chain(pivot.measures.iter().map(|measure| &measure.field))
        .chain(pivot.filters.iter().filter_map(filter_field_ref))
        .chain(pivot.filters.iter().filter_map(filter_measure_field_ref))
        .chain(pivot.groupings.iter().map(grouping_field_ref))
    {
        if field_index(fields, &field.name).is_none() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {} references unknown source field: {}",
                pivot.name, field.name
            )));
        }
    }

    Ok(())
}

fn validate_pivot_groupings(pivot: &PivotTable, fields: &[CacheField]) -> XlsxResult<()> {
    let mut grouped_fields = HashSet::new();
    for grouping in &pivot.groupings {
        let field = grouping_field_ref(grouping);
        if field_index(fields, &field.name).is_none() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {} references unknown grouped source field: {}",
                pivot.name, field.name
            )));
        }
        if !grouped_fields.insert(field.name.to_lowercase()) {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {} has more than one grouping for field {}",
                pivot.name, field.name
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
                    return Err(XlsxError::InvalidFormat(format!(
                        "pivot table {} has an invalid numeric grouping interval for field {}",
                        pivot.name, field.name
                    )));
                }
                if start.is_some_and(|value| !value.is_finite())
                    || end.is_some_and(|value| !value.is_finite())
                {
                    return Err(XlsxError::InvalidFormat(format!(
                        "pivot table {} has a non-finite numeric grouping bound for field {}",
                        pivot.name, field.name
                    )));
                }
            }
            PivotGrouping::Date { units, .. } => {
                if units.is_empty() {
                    return Err(XlsxError::InvalidFormat(format!(
                        "pivot table {} has an empty date grouping for field {}",
                        pivot.name, field.name
                    )));
                }
            }
            PivotGrouping::Manual { groups, .. } => {
                validate_manual_grouping(&pivot.name, &field.name, groups)?;
            }
        }
    }

    Ok(())
}

fn mark_metadata_only_measure_fields(pivot: &PivotTable, fields: &mut [CacheField]) {
    let measure_fields = pivot
        .measures
        .iter()
        .map(|measure| measure.field.name.to_lowercase())
        .collect::<HashSet<_>>();
    let explicit_item_fields = pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .map(|field| field.field.name.to_lowercase())
        .chain(
            pivot
                .filters
                .iter()
                .filter_map(filter_field_ref)
                .map(|field| field.name.to_lowercase()),
        )
        .chain(
            pivot
                .groupings
                .iter()
                .map(grouping_field_ref)
                .map(|field| field.name.to_lowercase()),
        )
        .collect::<HashSet<_>>();

    for field in fields {
        let name = field.name.to_lowercase();
        field.metadata_only_shared_items =
            measure_fields.contains(&name) && !explicit_item_fields.contains(&name);
    }
}

fn merge_cache_field_usage(existing: &mut [CacheField], incoming: &[CacheField]) {
    for field in existing {
        if let Some(incoming) = incoming
            .iter()
            .find(|incoming| incoming.name.eq_ignore_ascii_case(&field.name))
        {
            field.metadata_only_shared_items &= incoming.metadata_only_shared_items;
        }
    }
}

fn apply_grouped_cache_fields(
    pivot_name: &str,
    resolved: &mut ResolvedPivotSource,
    groupings: &[PivotGrouping],
    date_1904: bool,
) -> XlsxResult<()> {
    let date_system = if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    };

    for grouping in groupings {
        let field_name = &grouping_field_ref(grouping).name;
        let field_index = field_index(&resolved.fields, field_name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} references unknown grouped source field: {field_name}"
            ))
        })?;

        match grouping {
            PivotGrouping::Manual { groups, .. } => {
                let (item_indexes, group_items) = manual_group_cache_items(
                    pivot_name,
                    field_name,
                    &resolved.fields[field_index],
                    groups,
                )?;
                let grouped_name = unique_manual_grouped_header(&resolved.fields, field_name);
                let grouped_index = resolved.fields.len();
                resolved.fields[field_index].group = Some(CacheFieldGroup::Base {
                    parent: grouped_index,
                });
                let mut grouped = CacheField::new(grouped_name);
                grouped.database_field = false;
                grouped.group = Some(CacheFieldGroup::Manual {
                    base: field_index,
                    item_indexes,
                    group_items,
                });
                resolved.fields.push(grouped);
            }
            PivotGrouping::Date { units, .. } if units.len() > 1 => {
                let base_values = resolved
                    .rows
                    .iter()
                    .map(|row| {
                        row.get(field_index)
                            .and_then(|index| *index)
                            .and_then(|index| {
                                resolved.fields[field_index]
                                    .shared_items
                                    .get(index as usize)
                                    .cloned()
                            })
                            .unwrap_or(PivotValue::Blank)
                    })
                    .collect::<Vec<_>>();

                let mut parent = None;
                for unit in units {
                    let header = unique_grouped_header(&resolved.fields, field_name, *unit);
                    let mut grouped = CacheField::new(header);
                    grouped.group = Some(CacheFieldGroup::DateUnit {
                        base: field_index,
                        parent,
                        unit: *unit,
                    });

                    let row_indexes = base_values
                        .iter()
                        .map(|value| {
                            let grouped_value = group_date_value(value, *unit, date_system);
                            grouped.intern(grouped_value)
                        })
                        .collect::<Vec<_>>();
                    for (row, index) in resolved.rows.iter_mut().zip(row_indexes) {
                        row.push(Some(index));
                    }

                    parent = Some(resolved.fields.len());
                    resolved.fields.push(grouped);
                }
            }
            _ => match grouping {
                _ => {
                    resolved.fields[field_index].group =
                        Some(CacheFieldGroup::Range(grouping.clone()));
                }
            },
        }
    }

    Ok(())
}

fn validate_manual_grouping(
    pivot_name: &str,
    field_name: &str,
    groups: &[PivotManualGroup],
) -> XlsxResult<()> {
    if groups.is_empty() {
        return Err(XlsxError::InvalidFormat(format!(
            "pivot table {pivot_name} has an empty manual grouping for field {field_name}"
        )));
    }

    let mut group_names = HashSet::new();
    let mut members = HashSet::new();
    for group in groups {
        if group.name.trim().is_empty() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} has a manual group with a blank name for field {field_name}"
            )));
        }
        if group.members.is_empty() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} manual group {} has no members",
                group.name
            )));
        }
        if !group_names.insert(group.name.to_lowercase()) {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} has duplicate manual group name {}",
                group.name
            )));
        }
        for member in &group.members {
            if !members.insert(member.clone()) {
                return Err(XlsxError::InvalidFormat(format!(
                    "pivot table {pivot_name} assigns pivot item {member} to more than one manual group"
                )));
            }
        }
    }

    Ok(())
}

fn manual_group_cache_items(
    pivot_name: &str,
    field_name: &str,
    base_field: &CacheField,
    groups: &[PivotManualGroup],
) -> XlsxResult<(Vec<u32>, Vec<PivotValue>)> {
    let mut member_to_group = HashMap::new();
    for group in groups {
        for member in &group.members {
            if !base_field.item_lookup.contains_key(member) {
                return Err(XlsxError::InvalidFormat(format!(
                    "pivot table {pivot_name} manual group {} references item not found in field {field_name}: {member}",
                    group.name
                )));
            }
            member_to_group.insert(member.clone(), group.name.clone());
        }
    }

    let mut group_items = Vec::new();
    let mut ungrouped_item_indexes = HashMap::new();

    for item in &base_field.shared_items {
        if member_to_group.contains_key(item) {
            continue;
        }
        let index = group_items.len() as u32;
        ungrouped_item_indexes.insert(item.clone(), index);
        group_items.push(item.clone());
    }

    let mut group_name_indexes = HashMap::new();
    for group in groups {
        let value = PivotValue::String(group.name.clone());
        let index = group_items.len() as u32;
        group_name_indexes.insert(group.name.clone(), index);
        group_items.push(value);
    }

    let item_indexes = base_field
        .shared_items
        .iter()
        .map(|item| {
            if let Some(group_name) = member_to_group.get(item) {
                group_name_indexes
                    .get(group_name)
                    .copied()
                    .ok_or_else(|| {
                        XlsxError::InvalidFormat(format!(
                            "pivot table {pivot_name} could not map manual group for field {field_name}: {group_name}"
                        ))
                    })
            } else {
                ungrouped_item_indexes
                    .get(item)
                    .copied()
                    .ok_or_else(|| {
                        XlsxError::InvalidFormat(format!(
                            "pivot table {pivot_name} could not map ungrouped item for field {field_name}: {item}"
                        ))
                    })
            }
        })
        .collect::<XlsxResult<Vec<_>>>()?;

    Ok((item_indexes, group_items))
}

fn unique_manual_grouped_header(fields: &[CacheField], field_name: &str) -> String {
    for suffix in 2.. {
        let candidate = format!("{field_name}{suffix}");
        if fields
            .iter()
            .all(|field| !field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded manual grouped header suffix search should return")
}

fn unique_grouped_header(
    fields: &[CacheField],
    field_name: &str,
    unit: PivotDateGroupUnit,
) -> String {
    let base = grouped_date_header(field_name, unit);
    if fields
        .iter()
        .all(|field| !field.name.eq_ignore_ascii_case(&base))
    {
        return base;
    }

    for suffix in 2.. {
        let candidate = format!("{base} {suffix}");
        if fields
            .iter()
            .all(|field| !field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded grouped header suffix search should return")
}

fn grouped_date_header(field_name: &str, unit: PivotDateGroupUnit) -> String {
    format!("{field_name} ({})", date_group_unit_name(unit))
}

fn date_group_unit_name(unit: PivotDateGroupUnit) -> &'static str {
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

fn group_date_value(
    value: &PivotValue,
    unit: PivotDateGroupUnit,
    date_system: DateSystem,
) -> PivotValue {
    let PivotValue::Number(serial) = value else {
        return value.clone();
    };
    if !serial.is_finite() {
        return value.clone();
    }
    let Some((year, month, day)) = serial_to_date(*serial, date_system) else {
        return value.clone();
    };
    let (hour, minute, second) = serial_to_time(*serial);

    let value = match unit {
        PivotDateGroupUnit::Years => year as f64,
        PivotDateGroupUnit::Quarters => ((month - 1) / 3 + 1) as f64,
        PivotDateGroupUnit::Months => month as f64,
        PivotDateGroupUnit::Days => day as f64,
        PivotDateGroupUnit::Hours => hour as f64,
        PivotDateGroupUnit::Minutes => minute as f64,
        PivotDateGroupUnit::Seconds => second as f64,
    };
    PivotValue::Number(value)
}

fn grouping_field_ref(grouping: &PivotGrouping) -> &PivotFieldRef {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => field,
    }
}

fn cache_key_for_pivot(source_key: &str, pivot: &PivotTable) -> String {
    if pivot.groupings.is_empty()
        && pivot.calculated_fields.is_empty()
        && !pivot.refresh_policy.refresh_on_open
        && !pivot.refresh_policy.background_query
        && pivot.refresh_policy.missing_items_limit.is_none()
    {
        return source_key.to_string();
    }

    let mut grouping_signatures = pivot
        .groupings
        .iter()
        .map(grouping_signature)
        .collect::<Vec<_>>();
    grouping_signatures.sort();
    let calculated_signatures = pivot
        .calculated_fields
        .iter()
        .map(calculated_field_signature)
        .collect::<Vec<_>>()
        .join(";");
    format!(
        "{source_key}|calculated:{calculated_signatures}|groupings:{}|refreshOnLoad:{}|backgroundQuery:{}|missingItemsLimit:{}",
        grouping_signatures.join(";"),
        pivot.refresh_policy.refresh_on_open,
        pivot.refresh_policy.background_query,
        pivot
            .refresh_policy
            .missing_items_limit
            .map(|limit| limit.to_string())
            .unwrap_or_else(|| "none".to_string())
    )
}

fn calculated_field_signature(field: &PivotCalculatedField) -> String {
    format!(
        "{}:{}",
        field.name.to_lowercase(),
        normalized_formula_for_key(&field.formula)
    )
}

fn normalized_formula_for_key(formula: &str) -> String {
    formula.trim().trim_start_matches('=').to_string()
}

fn grouping_signature(grouping: &PivotGrouping) -> String {
    match grouping {
        PivotGrouping::Number {
            field,
            start,
            end,
            interval,
        } => format!(
            "n:{}:{}:{}:{}",
            field.name.to_lowercase(),
            f64_option_signature(*start),
            f64_option_signature(*end),
            f64_signature(*interval)
        ),
        PivotGrouping::Date { field, units } => {
            let units = units
                .iter()
                .map(|unit| date_group_by_name(*unit))
                .collect::<Vec<_>>()
                .join(",");
            format!("d:{}:{units}", field.name.to_lowercase())
        }
        PivotGrouping::Manual { field, groups } => {
            let groups = groups
                .iter()
                .map(|group| {
                    let members = group
                        .members
                        .iter()
                        .map(ToString::to_string)
                        .collect::<Vec<_>>()
                        .join(",");
                    format!("{}=[{}]", group.name, members)
                })
                .collect::<Vec<_>>()
                .join(",");
            format!("m:{}:{groups}", field.name.to_lowercase())
        }
    }
}

fn f64_option_signature(value: Option<f64>) -> String {
    value
        .map(f64_signature)
        .unwrap_or_else(|| "auto".to_string())
}

fn f64_signature(value: f64) -> String {
    format!("{:016x}", value.to_bits())
}

fn filter_field_ref(filter: &PivotFilter) -> Option<&PivotFieldRef> {
    match filter {
        PivotFilter::FieldItems { field, .. }
        | PivotFilter::Label { field, .. }
        | PivotFilter::Value { field, .. }
        | PivotFilter::TopN { field, .. } => Some(field),
        PivotFilter::Unsupported { .. } => None,
    }
}

fn filter_measure_field_ref(filter: &PivotFilter) -> Option<&PivotFieldRef> {
    match filter {
        PivotFilter::Value { measure, .. } | PivotFilter::TopN { measure, .. } => {
            Some(&measure.field)
        }
        PivotFilter::FieldItems { .. }
        | PivotFilter::Label { .. }
        | PivotFilter::Unsupported { .. } => None,
    }
}

fn show_as_base_field_ref(measure: &PivotMeasure) -> Option<&PivotFieldRef> {
    match &measure.show_as {
        PivotShowAs::RunningTotal { base_field }
        | PivotShowAs::DifferenceFrom { base_field, .. }
        | PivotShowAs::PercentDifferenceFrom { base_field, .. }
        | PivotShowAs::RankAscending { base_field }
        | PivotShowAs::RankDescending { base_field } => Some(base_field),
        PivotShowAs::Normal
        | PivotShowAs::PercentOfGrandTotal
        | PivotShowAs::PercentOfRowTotal
        | PivotShowAs::PercentOfColumnTotal
        | PivotShowAs::Index => None,
    }
}

fn resolve_pivot_source(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    pivot: &PivotTable,
) -> XlsxResult<ResolvedPivotSource> {
    match &pivot.source {
        PivotSource::WorksheetRange { sheet, range } => {
            let source_sheet_index = match sheet {
                Some(sheet_name) => workbook.sheet_index(sheet_name).ok_or_else(|| {
                    XlsxError::InvalidFormat(format!("pivot source sheet not found: {sheet_name}"))
                })?,
                None => pivot_sheet_index,
            };
            let source_sheet = workbook.worksheet(source_sheet_index).ok_or_else(|| {
                XlsxError::InvalidFormat(format!(
                    "pivot source sheet index out of bounds: {source_sheet_index}"
                ))
            })?;
            let sheet_name = source_sheet.name().to_string();
            let key = format!("range:{source_sheet_index}:{}", range.to_a1_string());
            let (fields, rows, record_count) = build_cache_data_from_range(
                source_sheet,
                *range,
                range.start.row + 1,
                range.end.row,
            )?;
            Ok(ResolvedPivotSource {
                key,
                source: PivotSource::WorksheetRange {
                    sheet: Some(sheet_name),
                    range: *range,
                },
                source_sheet_index,
                fields,
                rows,
                record_count,
                save_data: true,
            })
        }
        PivotSource::Table { name } => {
            let (source_sheet_index, source_sheet, table) =
                find_table(workbook, name).ok_or_else(|| {
                    XlsxError::InvalidFormat(format!("pivot source table not found: {name}"))
                })?;
            let key = format!("table:{}", table.name.to_lowercase());
            let headers = table_headers(table, source_sheet);
            let data_start = table.reference.start.row + table.header_row_count;
            let data_end = table
                .reference
                .end
                .row
                .saturating_sub(table.totals_row_count);
            let (fields, rows, record_count) = build_cache_data(
                source_sheet,
                table.reference.start.col,
                headers,
                data_start,
                data_end,
            )?;
            Ok(ResolvedPivotSource {
                key,
                source: PivotSource::Table {
                    name: table.name.clone(),
                },
                source_sheet_index,
                fields,
                rows,
                record_count,
                save_data: true,
            })
        }
        PivotSource::Olap { .. } => Err(XlsxError::InvalidFormat(
            "XLSX OLAP pivot source writing is not supported yet".into(),
        )),
        PivotSource::External {
            connection_name,
            command_text,
        } => {
            validate_external_pivot_connection(workbook, connection_name, command_text.as_deref())?;
            resolve_non_refreshable_pivot_source(pivot_sheet_index, &pivot.source, pivot)
        }
        PivotSource::Consolidation { .. } | PivotSource::Scenario { .. } => {
            resolve_non_refreshable_pivot_source(pivot_sheet_index, &pivot.source, pivot)
        }
    }
}

fn validate_external_pivot_connection(
    workbook: &Workbook,
    connection_name: &str,
    command_text: Option<&str>,
) -> XlsxResult<()> {
    let Some(command_text) = command_text else {
        return Ok(());
    };
    let Some(connection) = find_workbook_connection(workbook, connection_name) else {
        return Err(XlsxError::InvalidFormat(format!(
            "XLSX pivot external source command text requires a matching workbook data connection: {connection_name}"
        )));
    };
    let WorkbookConnectionKind::Database { command, .. } = &connection.kind;
    if command.as_deref() != Some(command_text) {
        return Err(XlsxError::InvalidFormat(format!(
            "XLSX pivot external source command text does not match workbook data connection: {connection_name}"
        )));
    }
    Ok(())
}

fn find_workbook_connection<'a>(
    workbook: &'a Workbook,
    connection_name: &str,
) -> Option<&'a WorkbookConnection> {
    connection_name
        .parse::<u32>()
        .ok()
        .and_then(|id| workbook.data_connection_by_id(id))
        .or_else(|| workbook.data_connection_by_name(connection_name))
}

fn resolve_non_refreshable_pivot_source(
    pivot_sheet_index: usize,
    source: &PivotSource,
    pivot: &PivotTable,
) -> XlsxResult<ResolvedPivotSource> {
    let fields = non_refreshable_source_fields(pivot)?;
    Ok(ResolvedPivotSource {
        key: non_refreshable_source_key(source),
        source: source.clone(),
        source_sheet_index: pivot_sheet_index,
        fields,
        rows: Vec::new(),
        record_count: 0,
        save_data: false,
    })
}

fn non_refreshable_source_key(source: &PivotSource) -> String {
    match source {
        PivotSource::External {
            connection_name,
            command_text,
        } => format!(
            "external:{connection_name}:{}",
            command_text.as_deref().unwrap_or("")
        ),
        PivotSource::Consolidation { ranges } => {
            let ranges = ranges
                .iter()
                .map(|range| format!("{}!{}", range.sheet, range.range.to_a1_string()))
                .collect::<Vec<_>>()
                .join(",");
            format!("consolidation:{ranges}")
        }
        PivotSource::Scenario { name } => format!("scenario:{name}"),
        PivotSource::Olap {
            connection_name,
            cube,
            command_text,
        } => format!(
            "olap:{connection_name}:{}:{}",
            cube.as_deref().unwrap_or(""),
            command_text.as_deref().unwrap_or("")
        ),
        PivotSource::WorksheetRange { .. } | PivotSource::Table { .. } => unreachable!(),
    }
}

fn non_refreshable_source_fields(pivot: &PivotTable) -> XlsxResult<Vec<CacheField>> {
    let calculated_names = pivot
        .calculated_fields
        .iter()
        .map(|field| field.name.to_lowercase())
        .collect::<HashSet<_>>();
    let mut fields = Vec::new();
    let mut seen = HashSet::new();

    for field in pivot
        .rows
        .iter()
        .map(|field| &field.field)
        .chain(pivot.columns.iter().map(|field| &field.field))
        .chain(pivot.page_fields.iter().map(|field| &field.field))
        .chain(pivot.measures.iter().map(|measure| &measure.field))
        .chain(pivot.filters.iter().filter_map(filter_field_ref))
        .chain(pivot.filters.iter().filter_map(filter_measure_field_ref))
        .chain(pivot.groupings.iter().map(grouping_field_ref))
        .chain(pivot.measures.iter().filter_map(show_as_base_field_ref))
    {
        if calculated_names.contains(&field.name.to_lowercase()) {
            continue;
        }
        push_non_refreshable_field(&mut fields, &mut seen, &field.name);
    }

    for field in &pivot.calculated_fields {
        if field.name.trim().is_empty() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {} has a calculated field with a blank name",
                pivot.name
            )));
        }
        if !seen.insert(field.name.to_lowercase()) {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {} calculated field duplicates source field: {}",
                pivot.name, field.name
            )));
        }
        fields.push(CacheField::calculated(
            field.name.clone(),
            formula_for_cache_attr(&field.formula),
        ));
    }

    Ok(fields)
}

fn push_non_refreshable_field(
    fields: &mut Vec<CacheField>,
    seen: &mut HashSet<String>,
    name: &str,
) {
    if seen.insert(name.to_lowercase()) {
        fields.push(CacheField::new(name.to_string()));
    }
}

fn build_cache_data_from_range(
    worksheet: &Worksheet,
    range: CellRange,
    data_start: u32,
    data_end: u32,
) -> XlsxResult<(Vec<CacheField>, Vec<Vec<Option<u32>>>, usize)> {
    let headers = (range.start.col..=range.end.col)
        .map(|col| {
            let value = effective_pivot_value(worksheet, range.start.row, col);
            let header = value.to_string();
            if header.trim().is_empty() {
                Err(XlsxError::InvalidFormat(format!(
                    "pivot source header cannot be blank at {}",
                    duke_sheets_core::CellAddress::new(range.start.row, col)
                )))
            } else {
                Ok(header)
            }
        })
        .collect::<XlsxResult<Vec<_>>>()?;

    build_cache_data(worksheet, range.start.col, headers, data_start, data_end)
}

fn build_cache_data(
    worksheet: &Worksheet,
    start_col: u16,
    headers: Vec<String>,
    data_start: u32,
    data_end: u32,
) -> XlsxResult<(Vec<CacheField>, Vec<Vec<Option<u32>>>, usize)> {
    validate_headers(&headers)?;

    let mut fields = headers.into_iter().map(CacheField::new).collect::<Vec<_>>();
    let mut rows = Vec::new();

    if data_start <= data_end {
        for row_values in
            collect_cache_row_values(worksheet, start_col, fields.len(), data_start, data_end)
        {
            let mut record = Vec::with_capacity(fields.len());
            for (field, value) in fields.iter_mut().zip(row_values) {
                let index = field.intern(value);
                record.push(Some(index));
            }
            rows.push(record);
        }
    }

    let record_count = rows.len();
    Ok((fields, rows, record_count))
}

fn collect_cache_row_values(
    worksheet: &Worksheet,
    start_col: u16,
    field_count: usize,
    data_start: u32,
    data_end: u32,
) -> Vec<Vec<PivotValue>> {
    #[cfg(feature = "parallel")]
    {
        let row_count = (data_end - data_start + 1) as usize;
        if row_count >= PARALLEL_CACHE_ROW_THRESHOLD {
            return (data_start..=data_end)
                .into_par_iter()
                .map(|row| cache_row_values(worksheet, row, start_col, field_count))
                .collect();
        }
    }

    (data_start..=data_end)
        .map(|row| cache_row_values(worksheet, row, start_col, field_count))
        .collect()
}

fn cache_row_values(
    worksheet: &Worksheet,
    row: u32,
    start_col: u16,
    field_count: usize,
) -> Vec<PivotValue> {
    (0..field_count)
        .map(|offset| effective_pivot_value(worksheet, row, start_col + offset as u16))
        .collect()
}

fn validate_headers(headers: &[String]) -> XlsxResult<()> {
    let mut seen = std::collections::HashSet::new();
    for header in headers {
        if header.trim().is_empty() {
            return Err(XlsxError::InvalidFormat(
                "pivot source headers cannot be blank".into(),
            ));
        }
        if !seen.insert(header.to_lowercase()) {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot source header is duplicated: {header}"
            )));
        }
    }
    Ok(())
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

fn table_headers(table: &Table, worksheet: &Worksheet) -> Vec<String> {
    let col_count = table.reference.col_count() as usize;
    (0..col_count)
        .map(|index| {
            table
                .columns
                .get(index)
                .map(|column| column.name.clone())
                .filter(|name| !name.trim().is_empty())
                .unwrap_or_else(|| {
                    effective_pivot_value(
                        worksheet,
                        table.reference.start.row,
                        table.reference.start.col + index as u16,
                    )
                    .to_string()
                })
        })
        .collect()
}

fn effective_pivot_value(worksheet: &Worksheet, row: u32, col: u16) -> PivotValue {
    worksheet
        .get_calculated_value_at(row, col)
        .map(PivotValue::from_cell_value)
        .unwrap_or_else(|| PivotValue::from_cell_value(&worksheet.get_value_at(row, col)))
}

fn apply_calculated_cache_fields(
    pivot_name: &str,
    resolved: &mut ResolvedPivotSource,
    calculated_fields: &[PivotCalculatedField],
) -> XlsxResult<()> {
    for field in calculated_fields {
        if field.name.trim().is_empty() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} has a calculated field with a blank name"
            )));
        }
        if field_index(&resolved.fields, &field.name).is_some() {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} calculated field duplicates source field: {}",
                field.name
            )));
        }

        let ast = parse_calculated_formula(pivot_name, field)?;
        let lookup = cache_field_lookup(&resolved.fields);
        let mut cache_field =
            CacheField::calculated(field.name.clone(), formula_for_cache_attr(&field.formula));
        for row in &mut resolved.rows {
            let value = evaluate_calculated_cache_row(
                pivot_name,
                field,
                &ast,
                &resolved.fields,
                row,
                &lookup,
            )?;
            let index = cache_field.intern(value);
            row.push(Some(index));
        }
        resolved.fields.push(cache_field);
    }

    Ok(())
}

fn parse_calculated_formula(
    pivot_name: &str,
    field: &PivotCalculatedField,
) -> XlsxResult<FormulaExpr> {
    let formula = field.formula.trim();
    if formula.is_empty() {
        return Err(XlsxError::InvalidFormat(format!(
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
        XlsxError::InvalidFormat(format!(
            "pivot table {pivot_name} calculated field {} formula did not parse: {error}",
            field.name
        ))
    })
}

fn cache_field_lookup(fields: &[CacheField]) -> HashMap<String, usize> {
    fields
        .iter()
        .enumerate()
        .map(|(index, field)| (field.name.to_lowercase(), index))
        .collect()
}

fn evaluate_calculated_cache_row(
    pivot_name: &str,
    field: &PivotCalculatedField,
    ast: &FormulaExpr,
    fields: &[CacheField],
    row: &[Option<u32>],
    lookup: &HashMap<String, usize>,
) -> XlsxResult<PivotValue> {
    let materialized = materialize_calculated_expr(pivot_name, field, ast, fields, row, lookup)?;
    let value = evaluate(&materialized, &EvaluationContext::simple()).map_err(|error| {
        XlsxError::InvalidFormat(format!(
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
    fields: &[CacheField],
    row: &[Option<u32>],
    lookup: &HashMap<String, usize>,
) -> XlsxResult<FormulaExpr> {
    Ok(match expr {
        FormulaExpr::Number(value) => FormulaExpr::Number(*value),
        FormulaExpr::String(value) => FormulaExpr::String(value.clone()),
        FormulaExpr::Boolean(value) => FormulaExpr::Boolean(*value),
        FormulaExpr::Error(value) => FormulaExpr::Error(*value),
        FormulaExpr::Empty => FormulaExpr::Empty,
        FormulaExpr::NameRef(name) => {
            calculated_cache_value_expr(pivot_name, field, name, fields, row, lookup)?
        }
        FormulaExpr::StructuredRef(reference) => {
            if let Some(name) = structured_ref_field_name(reference) {
                calculated_cache_value_expr(pivot_name, field, name, fields, row, lookup)?
            } else {
                return Err(XlsxError::InvalidFormat(format!(
                    "pivot table {pivot_name} calculated field {} uses an unsupported structured reference",
                    field.name
                )));
            }
        }
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::ExternalRef(_) => {
            return Err(XlsxError::InvalidFormat(format!(
                "pivot table {pivot_name} calculated field {} uses workbook references, which are not valid pivot source-field references",
                field.name
            )));
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op: *op,
            left: Box::new(materialize_calculated_expr(
                pivot_name, field, left, fields, row, lookup,
            )?),
            right: Box::new(materialize_calculated_expr(
                pivot_name, field, right, fields, row, lookup,
            )?),
        },
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp {
            op: *op,
            operand: Box::new(materialize_calculated_expr(
                pivot_name, field, operand, fields, row, lookup,
            )?),
        },
        FormulaExpr::Function { name, args } => FormulaExpr::Function {
            name: name.clone(),
            args: materialize_calculated_args(pivot_name, field, args, fields, row, lookup)?,
        },
        FormulaExpr::ExternalFunction { book, name, args } => FormulaExpr::ExternalFunction {
            book: book.clone(),
            name: name.clone(),
            args: materialize_calculated_args(pivot_name, field, args, fields, row, lookup)?,
        },
        FormulaExpr::Array(rows) => {
            let mut materialized_rows = Vec::with_capacity(rows.len());
            for formula_row in rows {
                materialized_rows.push(materialize_calculated_args(
                    pivot_name,
                    field,
                    formula_row,
                    fields,
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
    fields: &[CacheField],
    row: &[Option<u32>],
    lookup: &HashMap<String, usize>,
) -> XlsxResult<Vec<FormulaExpr>> {
    args.iter()
        .map(|arg| materialize_calculated_expr(pivot_name, field, arg, fields, row, lookup))
        .collect()
}

fn calculated_cache_value_expr(
    pivot_name: &str,
    field: &PivotCalculatedField,
    name: &str,
    fields: &[CacheField],
    row: &[Option<u32>],
    lookup: &HashMap<String, usize>,
) -> XlsxResult<FormulaExpr> {
    let field_index = lookup.get(&name.to_lowercase()).copied().ok_or_else(|| {
        XlsxError::InvalidFormat(format!(
            "pivot table {pivot_name} calculated field {} references unknown field: {name}",
            field.name
        ))
    })?;
    let value = row
        .get(field_index)
        .and_then(|index| *index)
        .and_then(|index| fields[field_index].shared_items.get(index as usize))
        .unwrap_or(&PivotValue::Blank);
    Ok(pivot_value_to_formula_expr(value))
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

fn formula_for_cache_attr(formula: &str) -> String {
    formula.trim().trim_start_matches('=').to_string()
}

pub(super) fn write_pivot_table_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    part: &PivotTablePart,
    cache_part: &PivotCachePart,
    style_table: &XlsxStyleTable,
) -> XlsxResult<()> {
    let sheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = sheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table not found".into()))?;

    let path = format!("xl/pivotTables/pivotTable{}.xml", part.table_num);
    write_xml_part(zip, &path, |w| {
        let cache_id = part.cache_num.to_string();
        let row_grand = bool_attr(pivot.layout.show_row_grand_totals);
        let col_grand = bool_attr(pivot.layout.show_column_grand_totals);
        let preserve_formatting = bool_attr(pivot.refresh_policy.preserve_formatting);
        let show_headers = bool_attr(pivot.layout.show_field_headers);
        let show_drill = bool_attr(pivot.layout.show_expand_collapse);
        let print_drill = bool_attr(pivot.layout.print_drill_indicators);
        let item_print_titles = bool_attr(pivot.layout.item_print_titles);
        let field_print_titles = bool_attr(pivot.layout.field_print_titles);
        let compact = bool_attr(matches!(pivot.layout.kind, PivotLayoutKind::Compact));
        let outline = bool_attr(matches!(pivot.layout.kind, PivotLayoutKind::Outline));

        let mut tag = BytesStart::new("pivotTableDefinition");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        tag.push_attribute(("name", pivot.name.as_str()));
        tag.push_attribute(("cacheId", cache_id.as_str()));
        tag.push_attribute(("dataCaption", "Values"));
        tag.push_attribute(("updatedVersion", "8"));
        tag.push_attribute(("minRefreshableVersion", "3"));
        tag.push_attribute(("rowGrandTotals", row_grand));
        tag.push_attribute(("colGrandTotals", col_grand));
        tag.push_attribute(("preserveFormatting", preserve_formatting));
        tag.push_attribute(("showHeaders", show_headers));
        tag.push_attribute(("showDrill", show_drill));
        tag.push_attribute(("printDrill", print_drill));
        tag.push_attribute(("itemPrintTitles", item_print_titles));
        tag.push_attribute(("fieldPrintTitles", field_print_titles));
        tag.push_attribute(("compact", compact));
        tag.push_attribute(("outline", outline));
        w.write_event(Event::Start(tag))?;

        write_location(w, pivot, &cache_part.fields, &cache_part.rows)?;
        write_pivot_fields(w, pivot, &cache_part.fields)?;
        write_axis_fields(w, "rowFields", &pivot.rows, &cache_part.fields)?;
        write_axis_items(
            w,
            "rowItems",
            &pivot.rows,
            &cache_part.fields,
            &cache_part.rows,
            pivot.layout.show_row_grand_totals,
            false,
        )?;
        write_axis_fields(w, "colFields", &pivot.columns, &cache_part.fields)?;
        write_axis_items(
            w,
            "colItems",
            &pivot.columns,
            &cache_part.fields,
            &cache_part.rows,
            pivot.layout.show_column_grand_totals,
            true,
        )?;
        write_page_fields(w, pivot, &cache_part.fields)?;
        write_data_fields(w, pivot, &cache_part.fields, style_table)?;
        write_pivot_style(w, pivot)?;
        write_pivot_filters(w, pivot, &cache_part.fields)?;
        write_pivot_extensions(w, pivot)?;

        w.write_event(Event::End(BytesEnd::new("pivotTableDefinition")))?;
        Ok(())
    })
}

pub(super) fn write_pivot_table_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    part: &PivotTablePart,
) -> XlsxResult<()> {
    let path = format!("xl/pivotTables/_rels/pivotTable{}.xml.rels", part.table_num);
    write_xml_part(zip, &path, |w| {
        let mut relationships = BytesStart::new("Relationships");
        relationships.push_attribute(("xmlns", NS_RELATIONSHIPS));
        w.write_event(Event::Start(relationships))?;

        let target = format!("../pivotCache/pivotCacheDefinition{}.xml", part.cache_num);
        w.create_element("Relationship")
            .with_attribute(("Id", "rId1"))
            .with_attribute(("Type", RT_PIVOT_CACHE_DEFINITION))
            .with_attribute(("Target", target.as_str()))
            .write_empty()?;

        w.write_event(Event::End(BytesEnd::new("Relationships")))?;
        Ok(())
    })
}

fn write_location(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    rows: &[Vec<Option<u32>>],
) -> XlsxResult<()> {
    let range = match pivot.rendered_range {
        Some(range) => range,
        None => estimated_pivot_output_range(pivot, fields, rows)?,
    };
    let ref_str = range.to_a1_string();
    let first_data_col = expanded_axis_field_count(&pivot.rows, fields)
        .max(1)
        .to_string();

    let mut location = BytesStart::new("location");
    location.push_attribute(("ref", ref_str.as_str()));
    location.push_attribute(("firstHeaderRow", "1"));
    location.push_attribute(("firstDataRow", "1"));
    location.push_attribute(("firstDataCol", first_data_col.as_str()));
    w.write_event(Event::Empty(location))?;
    Ok(())
}

fn estimated_pivot_output_range(
    pivot: &PivotTable,
    fields: &[CacheField],
    rows: &[Vec<Option<u32>>],
) -> XlsxResult<CellRange> {
    let row_label_width = expanded_axis_field_count(&pivot.rows, fields).max(1);
    let measure_width = pivot.measures.len().max(1);
    let column_tuple_count = axis_item_tuples(&pivot.columns, fields, rows)?.len().max(1);
    let value_width = if pivot.columns.is_empty() {
        measure_width
    } else {
        column_tuple_count * measure_width
    };
    let width = row_label_width + value_width;

    let row_tuple_count = axis_item_tuples(&pivot.rows, fields, rows)?.len();
    let data_rows = if pivot.rows.is_empty() {
        1
    } else {
        row_tuple_count + usize::from(pivot.layout.show_row_grand_totals)
    }
    .max(1);
    let header_rows = if pivot.columns.is_empty() {
        1
    } else {
        expanded_axis_field_count(&pivot.columns, fields).max(1) + 1
    };
    let height = header_rows + data_rows;

    Ok(range_from_size(pivot.target, width, height))
}

fn range_from_size(start: CellAddress, width: usize, height: usize) -> CellRange {
    const MAX_ROWS: usize = 1_048_576;
    const MAX_COLS: usize = 16_384;

    let max_width = MAX_COLS.saturating_sub(start.col as usize).max(1);
    let max_height = MAX_ROWS.saturating_sub(start.row as usize).max(1);
    let width = width.max(1).min(max_width);
    let height = height.max(1).min(max_height);
    let end_row = start.row + (height - 1) as u32;
    let end_col = start.col + (width - 1) as u16;
    CellRange::new(start, CellAddress::new(end_row, end_col))
}

fn write_pivot_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    let count = fields.len().to_string();
    let mut pivot_fields = BytesStart::new("pivotFields");
    pivot_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(pivot_fields))?;

    for (index, field) in fields.iter().enumerate() {
        let mut pivot_field = BytesStart::new("pivotField");
        if pivot_field_is_data_field(pivot, field) {
            pivot_field.push_attribute(("dataField", "1"));
        }
        if let Some(axis) = field_axis(pivot, fields, index) {
            pivot_field.push_attribute(("axis", axis));
        }
        let sort = field_sort(pivot, fields, index);
        if sort != "manual" {
            pivot_field.push_attribute(("sortType", sort));
        }
        if field_is_filtered(pivot, &field.name) {
            pivot_field.push_attribute(("multipleItemSelectionAllowed", "1"));
        }
        if let Some(axis_field) = pivot_axis_field(pivot, fields, index) {
            pivot_field.push_attribute(("showAll", bool_attr(axis_field.show_empty_items)));
            push_subtotal_attrs(&mut pivot_field, axis_field.subtotal);
        }

        let hidden_items = hidden_item_indexes(pivot, fields, index)?;
        let include_default = should_write_pivot_field_items(pivot, fields, index);
        if hidden_items.is_empty() && !include_default {
            w.write_event(Event::Empty(pivot_field))?;
        } else {
            w.write_event(Event::Start(pivot_field))?;
            write_pivot_field_items(w, field, &hidden_items, include_default)?;
            w.write_event(Event::End(BytesEnd::new("pivotField")))?;
        }
    }

    w.write_event(Event::End(BytesEnd::new("pivotFields")))?;
    Ok(())
}

fn pivot_field_is_data_field(pivot: &PivotTable, field: &CacheField) -> bool {
    pivot
        .measures
        .iter()
        .any(|measure| measure.field.name.eq_ignore_ascii_case(&field.name))
}

fn should_write_pivot_field_items(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> bool {
    pivot_axis_field(pivot, fields, field_index).is_some()
        || fields.get(field_index).is_some_and(|field| {
            matches!(
                field.group,
                Some(CacheFieldGroup::Base { .. } | CacheFieldGroup::Manual { .. })
            )
        })
}

fn write_pivot_field_items(
    w: &mut XmlWriter,
    field: &CacheField,
    hidden_items: &[u32],
    include_default: bool,
) -> XlsxResult<()> {
    let item_count = pivot_field_item_count(field);
    let count = (item_count + usize::from(include_default)).to_string();
    let mut items = BytesStart::new("items");
    items.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(items))?;
    for item_index in 0..item_count {
        let x = item_index.to_string();
        let mut item = BytesStart::new("item");
        item.push_attribute(("x", x.as_str()));
        if hidden_items.contains(&(item_index as u32)) {
            item.push_attribute(("h", "1"));
        }
        w.write_event(Event::Empty(item))?;
    }
    if include_default {
        let mut item = BytesStart::new("item");
        item.push_attribute(("t", "default"));
        w.write_event(Event::Empty(item))?;
    }
    w.write_event(Event::End(BytesEnd::new("items")))?;
    Ok(())
}

fn pivot_field_item_count(field: &CacheField) -> usize {
    match &field.group {
        Some(CacheFieldGroup::Manual { group_items, .. }) => group_items.len(),
        _ => field.shared_items.len(),
    }
}

fn field_axis(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> Option<&'static str> {
    let field_name = axis_semantic_field_name(fields, field_index)?;
    if pivot
        .rows
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisRow")
    } else if pivot
        .columns
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisCol")
    } else if pivot
        .page_fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisPage")
    } else {
        None
    }
}

fn field_sort(pivot: &PivotTable, fields: &[CacheField], field_index: usize) -> &'static str {
    let Some(field_name) = axis_semantic_field_name(fields, field_index) else {
        return "manual";
    };
    let sort = pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(&field_name))
        .map(|field| field.sort)
        .unwrap_or(PivotSort::None);

    match sort {
        PivotSort::None => "manual",
        PivotSort::Ascending => "ascending",
        PivotSort::Descending => "descending",
    }
}

fn field_is_filtered(pivot: &PivotTable, field_name: &str) -> bool {
    pivot.filters.iter().any(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    })
}

fn pivot_axis_field<'a>(
    pivot: &'a PivotTable,
    fields: &'a [CacheField],
    field_index: usize,
) -> Option<&'a duke_sheets_core::PivotField> {
    let field_name = axis_semantic_field_name(fields, field_index)?;
    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(&field_name))
}

fn axis_semantic_field_name(fields: &[CacheField], field_index: usize) -> Option<String> {
    let field = fields.get(field_index)?;
    if matches!(field.group, Some(CacheFieldGroup::Base { .. })) {
        return Some(field.name.clone());
    }
    if let Some(CacheFieldGroup::DateUnit { base, .. } | CacheFieldGroup::Manual { base, .. }) =
        &field.group
    {
        return fields.get(*base).map(|field| field.name.clone());
    }
    if has_grouped_children(fields, field_index) {
        return None;
    }
    Some(field.name.clone())
}

fn has_grouped_children(fields: &[CacheField], field_index: usize) -> bool {
    fields.iter().any(|field| {
        matches!(
            field.group,
            Some(CacheFieldGroup::DateUnit { base, .. } | CacheFieldGroup::Manual { base, .. })
                if base == field_index
        )
    })
}

fn push_subtotal_attrs(pivot_field: &mut BytesStart<'_>, subtotal: PivotSubtotal) {
    match subtotal {
        PivotSubtotal::Automatic => {}
        PivotSubtotal::None => pivot_field.push_attribute(("defaultSubtotal", "0")),
        PivotSubtotal::Sum => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("sumSubtotal", "1"));
        }
        PivotSubtotal::Count => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("countASubtotal", "1"));
        }
        PivotSubtotal::CountNumbers => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("countSubtotal", "1"));
        }
        PivotSubtotal::Average => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("avgSubtotal", "1"));
        }
        PivotSubtotal::Min => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("minSubtotal", "1"));
        }
        PivotSubtotal::Max => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("maxSubtotal", "1"));
        }
        PivotSubtotal::Product => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("productSubtotal", "1"));
        }
        PivotSubtotal::StdDev => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("stdDevSubtotal", "1"));
        }
        PivotSubtotal::StdDevP => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("stdDevPSubtotal", "1"));
        }
        PivotSubtotal::Var => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("varSubtotal", "1"));
        }
        PivotSubtotal::VarP => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            pivot_field.push_attribute(("varPSubtotal", "1"));
        }
    }
}

fn hidden_item_indexes(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> XlsxResult<Vec<u32>> {
    let field = &fields[field_index];
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field: filter_field, .. }
                if filter_field.name.eq_ignore_ascii_case(&field.name)
        )
    }) else {
        return Ok(Vec::new());
    };

    let allowed = allowed_items
        .iter()
        .filter_map(|item| field.item_lookup.get(item).copied())
        .collect::<std::collections::HashSet<_>>();

    Ok((0..field.shared_items.len() as u32)
        .filter(|index| !allowed.contains(index))
        .collect())
}

fn write_axis_fields(
    w: &mut XmlWriter,
    tag_name: &str,
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
) -> XlsxResult<()> {
    if axis_fields.is_empty() {
        return Ok(());
    }

    let indexes = expanded_axis_field_indexes(axis_fields, fields)?;
    let count = indexes.len().to_string();
    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(tag))?;
    for index in indexes {
        let x = index.to_string();
        let mut el = BytesStart::new("field");
        el.push_attribute(("x", x.as_str()));
        w.write_event(Event::Empty(el))?;
    }
    w.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn write_axis_items(
    w: &mut XmlWriter,
    tag_name: &str,
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
    rows: &[Vec<Option<u32>>],
    include_grand_total: bool,
    write_empty_item_when_no_fields: bool,
) -> XlsxResult<()> {
    if axis_fields.is_empty() {
        if write_empty_item_when_no_fields {
            let mut tag = BytesStart::new(tag_name);
            tag.push_attribute(("count", "1"));
            w.write_event(Event::Start(tag))?;
            w.write_event(Event::Empty(BytesStart::new("i")))?;
            w.write_event(Event::End(BytesEnd::new(tag_name)))?;
        }
        return Ok(());
    }

    let field_indexes = expanded_axis_field_indexes(axis_fields, fields)?;
    let tuples = axis_item_tuples(axis_fields, fields, rows)?;

    let count = (tuples.len() + usize::from(include_grand_total)).to_string();
    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(tag))?;

    for tuple in tuples {
        write_axis_item(w, None, &tuple)?;
    }
    if include_grand_total {
        let grand = vec![0; field_indexes.len()];
        write_axis_item(w, Some("grand"), &grand)?;
    }

    w.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn axis_item_tuples(
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
    rows: &[Vec<Option<u32>>],
) -> XlsxResult<Vec<Vec<u32>>> {
    if axis_fields.is_empty() {
        return Ok(Vec::new());
    }

    let field_indexes = expanded_axis_field_indexes(axis_fields, fields)?;
    let mut seen = HashSet::new();
    let mut tuples = Vec::new();
    for row in rows {
        let tuple = field_indexes
            .iter()
            .map(|field_index| cache_record_axis_item_index(fields, row, *field_index).unwrap_or(0))
            .collect::<Vec<_>>();
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(tuples)
}

fn cache_record_axis_item_index(
    fields: &[CacheField],
    row: &[Option<u32>],
    field_index: usize,
) -> Option<u32> {
    if let Some(index) = row.get(field_index).and_then(|index| *index) {
        return Some(index);
    }

    match fields.get(field_index)?.group.as_ref()? {
        CacheFieldGroup::Manual {
            base, item_indexes, ..
        } => row
            .get(*base)
            .and_then(|index| *index)
            .and_then(|index| item_indexes.get(index as usize).copied()),
        _ => None,
    }
}

fn write_axis_item(w: &mut XmlWriter, item_type: Option<&str>, indexes: &[u32]) -> XlsxResult<()> {
    let mut item = BytesStart::new("i");
    if let Some(item_type) = item_type {
        item.push_attribute(("t", item_type));
    }
    if indexes.is_empty() {
        w.write_event(Event::Empty(item))?;
        return Ok(());
    }

    w.write_event(Event::Start(item))?;
    for index in indexes {
        let mut x = BytesStart::new("x");
        let value = index.to_string();
        if *index != 0 {
            x.push_attribute(("v", value.as_str()));
        }
        w.write_event(Event::Empty(x))?;
    }
    w.write_event(Event::End(BytesEnd::new("i")))?;
    Ok(())
}

fn expanded_axis_field_count(
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
) -> usize {
    axis_fields
        .iter()
        .map(|field| {
            grouped_cache_field_indexes(fields, &field.field.name)
                .map(|indexes| indexes.len())
                .unwrap_or(1)
        })
        .sum()
}

fn expanded_axis_field_indexes(
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
) -> XlsxResult<Vec<usize>> {
    let mut indexes = Vec::new();
    for field in axis_fields {
        if let Some(grouped_indexes) = grouped_cache_field_indexes(fields, &field.field.name) {
            indexes.extend(grouped_indexes);
            continue;
        }

        let index = field_index(fields, &field.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", field.field.name))
        })?;
        indexes.push(index);
    }
    Ok(indexes)
}

fn grouped_cache_field_indexes(fields: &[CacheField], field_name: &str) -> Option<Vec<usize>> {
    let base = field_index(fields, field_name)?;
    let manual_indexes = fields
        .iter()
        .enumerate()
        .filter_map(|(index, field)| match field.group {
            Some(CacheFieldGroup::Manual {
                base: group_base, ..
            }) if group_base == base => Some(index),
            _ => None,
        })
        .collect::<Vec<_>>();
    if !manual_indexes.is_empty() {
        let mut indexes = manual_indexes;
        indexes.push(base);
        return Some(indexes);
    }

    let indexes = fields
        .iter()
        .enumerate()
        .filter_map(|(index, field)| match field.group {
            Some(CacheFieldGroup::DateUnit {
                base: group_base, ..
            }) if group_base == base => Some(index),
            _ => None,
        })
        .collect::<Vec<_>>();
    (!indexes.is_empty()).then_some(indexes)
}

fn write_page_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    if pivot.page_fields.is_empty() {
        return Ok(());
    }

    let count = pivot.page_fields.len().to_string();
    let mut page_fields = BytesStart::new("pageFields");
    page_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(page_fields))?;
    for field in &pivot.page_fields {
        let index = field_index(fields, &field.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", field.field.name))
        })?;
        let fld = index.to_string();
        let mut el = BytesStart::new("pageField");
        el.push_attribute(("fld", fld.as_str()));
        if let Some(item) = selected_page_item_index(pivot, &field.field.name, &fields[index]) {
            let item = item.to_string();
            el.push_attribute(("item", item.as_str()));
        }
        w.write_event(Event::Empty(el))?;
    }
    w.write_event(Event::End(BytesEnd::new("pageFields")))?;
    Ok(())
}

fn selected_page_item_index(
    pivot: &PivotTable,
    field_name: &str,
    field: &CacheField,
) -> Option<u32> {
    let PivotFilter::FieldItems { allowed_items, .. } = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    })?
    else {
        return None;
    };

    let [item] = allowed_items.as_slice() else {
        return None;
    };
    field.item_lookup.get(item).copied()
}

fn write_data_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    style_table: &XlsxStyleTable,
) -> XlsxResult<()> {
    let count = pivot.measures.len().to_string();
    let mut data_fields = BytesStart::new("dataFields");
    data_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(data_fields))?;
    for measure in &pivot.measures {
        let index = field_index(fields, &measure.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", measure.field.name))
        })?;
        let fld = index.to_string();
        let name = measure.caption();
        let mut data_field = BytesStart::new("dataField");
        data_field.push_attribute(("name", name.as_str()));
        data_field.push_attribute(("fld", fld.as_str()));
        data_field.push_attribute(("subtotal", aggregate_name(measure.aggregate)));
        let num_fmt_id = if let Some(number_format) = &measure.number_format {
            Some(
                style_table
                    .custom_num_fmt_id(number_format)
                    .ok_or_else(|| {
                        XlsxError::InvalidFormat(format!(
                            "pivot measure number format was not registered: {number_format}"
                        ))
                    })?
                    .to_string(),
            )
        } else {
            None
        };
        if let Some(num_fmt_id) = num_fmt_id.as_deref() {
            data_field.push_attribute(("numFmtId", num_fmt_id));
        }
        if let Some(show_data_as) = show_data_as_name(&measure.show_as) {
            data_field.push_attribute(("showDataAs", show_data_as));
        }
        let base_field = show_as_base_field_index(&measure.show_as, fields)?;
        if let Some(base_field) = base_field {
            let base_field = base_field.to_string();
            data_field.push_attribute(("baseField", base_field.as_str()));
        }
        let base_item = show_as_base_item_index(&measure.show_as, fields)?;
        if let Some(base_item) = base_item {
            let base_item = base_item.to_string();
            data_field.push_attribute(("baseItem", base_item.as_str()));
        }
        if let Some(rank_show_as) = rank_show_as_name(&measure.show_as) {
            w.write_event(Event::Start(data_field))?;
            write_data_field_ext(w, rank_show_as)?;
            w.write_event(Event::End(BytesEnd::new("dataField")))?;
        } else {
            w.write_event(Event::Empty(data_field))?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("dataFields")))?;
    Ok(())
}

fn write_pivot_filters(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    let filters = pivot
        .filters
        .iter()
        .filter(|filter| !matches!(filter, PivotFilter::FieldItems { .. }))
        .collect::<Vec<_>>();
    if filters.is_empty() {
        return Ok(());
    }

    let count = filters.len().to_string();
    let mut filters_el = BytesStart::new("filters");
    filters_el.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(filters_el))?;
    for (index, filter) in filters.into_iter().enumerate() {
        write_pivot_filter(w, pivot, fields, filter, index)?;
    }
    w.write_event(Event::End(BytesEnd::new("filters")))?;
    Ok(())
}

fn write_pivot_filter(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    filter: &PivotFilter,
    index: usize,
) -> XlsxResult<()> {
    let field_name = filter_field_ref(filter)
        .map(|field| field.name.as_str())
        .ok_or_else(|| XlsxError::InvalidFormat("unsupported pivot filter".into()))?;
    let field_index = field_index(fields, field_name).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot filter field not found: {field_name}"))
    })?;
    let fld = field_index.to_string();
    let eval_order = index.to_string();
    let id = (index + 1).to_string();

    let mut filter_el = BytesStart::new("filter");
    filter_el.push_attribute(("fld", fld.as_str()));
    filter_el.push_attribute(("evalOrder", eval_order.as_str()));
    filter_el.push_attribute(("id", id.as_str()));

    match filter {
        PivotFilter::Label {
            operator, value, ..
        } => {
            filter_el.push_attribute(("type", label_filter_type_name(*operator)));
            filter_el.push_attribute(("stringValue1", value.as_str()));
            w.write_event(Event::Start(filter_el))?;
            let custom_value = label_custom_filter_value(*operator, value);
            write_pivot_custom_filter(w, custom_filter_operator(*operator), &custom_value)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::Value {
            measure,
            operator,
            value,
            ..
        } => {
            let filter_type = value_filter_type_name(*operator).ok_or_else(|| {
                XlsxError::InvalidFormat("unsupported pivot value filter operator".into())
            })?;
            let measure_index = measure_index_for_filter(pivot, measure)?;
            let i_measure_fld = measure_index.to_string();
            let value = value.to_string();
            filter_el.push_attribute(("type", filter_type));
            filter_el.push_attribute(("iMeasureFld", i_measure_fld.as_str()));
            filter_el.push_attribute(("stringValue1", value.as_str()));
            w.write_event(Event::Start(filter_el))?;
            write_pivot_custom_filter(w, custom_filter_operator(*operator), &value)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::TopN {
            measure,
            n,
            top,
            percent,
            ..
        } => {
            let measure_index = measure_index_for_filter(pivot, measure)?;
            let i_measure_fld = measure_index.to_string();
            filter_el.push_attribute(("type", top_n_filter_type_name(*top, *percent)));
            filter_el.push_attribute(("iMeasureFld", i_measure_fld.as_str()));
            w.write_event(Event::Start(filter_el))?;
            write_pivot_top_filter(w, *n, *top, *percent)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::FieldItems { .. } => {}
        PivotFilter::Unsupported { kind, .. } => {
            return Err(XlsxError::InvalidFormat(format!(
                "unsupported pivot filter: {kind}"
            )));
        }
    }

    Ok(())
}

fn write_pivot_custom_filter(
    w: &mut XmlWriter,
    operator: &'static str,
    value: &str,
) -> XlsxResult<()> {
    write_pivot_auto_filter_start(w)?;

    w.write_event(Event::Start(BytesStart::new("customFilters")))?;
    let mut custom_filter = BytesStart::new("customFilter");
    custom_filter.push_attribute(("operator", operator));
    custom_filter.push_attribute(("val", value));
    w.write_event(Event::Empty(custom_filter))?;
    w.write_event(Event::End(BytesEnd::new("customFilters")))?;

    write_pivot_auto_filter_end(w)?;
    Ok(())
}

fn write_pivot_top_filter(w: &mut XmlWriter, n: u32, top: bool, percent: bool) -> XlsxResult<()> {
    write_pivot_auto_filter_start(w)?;

    let n = n.to_string();
    let mut top10 = BytesStart::new("top10");
    top10.push_attribute(("top", bool_attr(top)));
    top10.push_attribute(("percent", bool_attr(percent)));
    top10.push_attribute(("val", n.as_str()));
    w.write_event(Event::Empty(top10))?;

    write_pivot_auto_filter_end(w)?;
    Ok(())
}

fn write_pivot_auto_filter_start(w: &mut XmlWriter) -> XlsxResult<()> {
    let mut auto_filter = BytesStart::new("autoFilter");
    auto_filter.push_attribute(("ref", "A1"));
    w.write_event(Event::Start(auto_filter))?;

    let mut filter_column = BytesStart::new("filterColumn");
    filter_column.push_attribute(("colId", "0"));
    w.write_event(Event::Start(filter_column))?;
    Ok(())
}

fn write_pivot_auto_filter_end(w: &mut XmlWriter) -> XlsxResult<()> {
    w.write_event(Event::End(BytesEnd::new("filterColumn")))?;
    w.write_event(Event::End(BytesEnd::new("autoFilter")))?;
    Ok(())
}

fn measure_index_for_filter(
    pivot: &PivotTable,
    filter_measure: &duke_sheets_core::PivotMeasure,
) -> XlsxResult<usize> {
    pivot
        .measures
        .iter()
        .position(|measure| {
            measure
                .field
                .name
                .eq_ignore_ascii_case(&filter_measure.field.name)
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
            XlsxError::InvalidFormat(format!(
                "pivot filter measure not found: {}",
                filter_measure.caption()
            ))
        })
}

fn label_filter_type_name(operator: PivotFilterOperator) -> &'static str {
    match operator {
        PivotFilterOperator::Equals => "captionEqual",
        PivotFilterOperator::NotEquals => "captionNotEqual",
        PivotFilterOperator::LessThan => "captionLessThan",
        PivotFilterOperator::LessThanOrEqual => "captionLessThanOrEqual",
        PivotFilterOperator::GreaterThan => "captionGreaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "captionGreaterThanOrEqual",
        PivotFilterOperator::BeginsWith => "captionBeginsWith",
        PivotFilterOperator::DoesNotBeginWith => "captionNotBeginsWith",
        PivotFilterOperator::EndsWith => "captionEndsWith",
        PivotFilterOperator::DoesNotEndWith => "captionNotEndsWith",
        PivotFilterOperator::Contains => "captionContains",
        PivotFilterOperator::DoesNotContain => "captionNotContains",
    }
}

fn value_filter_type_name(operator: PivotFilterOperator) -> Option<&'static str> {
    Some(match operator {
        PivotFilterOperator::Equals => "valueEqual",
        PivotFilterOperator::NotEquals => "valueNotEqual",
        PivotFilterOperator::LessThan => "valueLessThan",
        PivotFilterOperator::LessThanOrEqual => "valueLessThanOrEqual",
        PivotFilterOperator::GreaterThan => "valueGreaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "valueGreaterThanOrEqual",
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => return None,
    })
}

fn top_n_filter_type_name(top: bool, percent: bool) -> &'static str {
    match (top, percent) {
        (true, true) => "topPercent",
        (true, false) => "topCount",
        (false, true) => "bottomPercent",
        (false, false) => "bottomCount",
    }
}

fn custom_filter_operator(operator: PivotFilterOperator) -> &'static str {
    match operator {
        PivotFilterOperator::Equals
        | PivotFilterOperator::BeginsWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::Contains => "equal",
        PivotFilterOperator::NotEquals
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::DoesNotContain => "notEqual",
        PivotFilterOperator::LessThan => "lessThan",
        PivotFilterOperator::LessThanOrEqual => "lessThanOrEqual",
        PivotFilterOperator::GreaterThan => "greaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "greaterThanOrEqual",
    }
}

fn label_custom_filter_value(operator: PivotFilterOperator, value: &str) -> String {
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

fn write_data_field_ext(w: &mut XmlWriter, pivot_show_as: &str) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("extLst")))?;

    let mut ext = BytesStart::new("ext");
    ext.push_attribute(("uri", EXT_URI_X14_DATA_FIELD));
    w.write_event(Event::Start(ext))?;

    let mut data_field = BytesStart::new("x14:dataField");
    data_field.push_attribute(("xmlns:x14", NS_SPREADSHEET_X14));
    data_field.push_attribute(("pivotShowAs", pivot_show_as));
    w.write_event(Event::Empty(data_field))?;

    w.write_event(Event::End(BytesEnd::new("ext")))?;
    w.write_event(Event::End(BytesEnd::new("extLst")))?;
    Ok(())
}

fn write_pivot_extensions(w: &mut XmlWriter, pivot: &PivotTable) -> XlsxResult<()> {
    if pivot
        .extensions
        .iter()
        .all(|extension| extension.payload.is_empty())
    {
        return Ok(());
    }

    w.write_event(Event::Start(BytesStart::new("extLst")))?;
    for extension in &pivot.extensions {
        if !extension.payload.is_empty() {
            w.get_mut().write_all(&extension.payload)?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("extLst")))?;
    Ok(())
}

fn is_writable_show_as(show_as: &PivotShowAs) -> bool {
    matches!(
        show_as,
        PivotShowAs::Normal
            | PivotShowAs::PercentOfGrandTotal
            | PivotShowAs::PercentOfRowTotal
            | PivotShowAs::PercentOfColumnTotal
            | PivotShowAs::Index
            | PivotShowAs::RunningTotal { .. }
            | PivotShowAs::DifferenceFrom { .. }
            | PivotShowAs::PercentDifferenceFrom { .. }
            | PivotShowAs::RankAscending { .. }
            | PivotShowAs::RankDescending { .. }
    )
}

fn show_data_as_name(show_as: &PivotShowAs) -> Option<&'static str> {
    match show_as {
        PivotShowAs::Normal => None,
        PivotShowAs::PercentOfGrandTotal => Some("percentOfTotal"),
        PivotShowAs::PercentOfRowTotal => Some("percentOfRow"),
        PivotShowAs::PercentOfColumnTotal => Some("percentOfCol"),
        PivotShowAs::Index => Some("index"),
        PivotShowAs::RunningTotal { .. } => Some("runTotal"),
        PivotShowAs::DifferenceFrom { .. } => Some("difference"),
        PivotShowAs::PercentDifferenceFrom { .. } => Some("percentDiff"),
        PivotShowAs::RankAscending { .. } | PivotShowAs::RankDescending { .. } => None,
    }
}

fn rank_show_as_name(show_as: &PivotShowAs) -> Option<&'static str> {
    match show_as {
        PivotShowAs::RankAscending { .. } => Some("rankAscending"),
        PivotShowAs::RankDescending { .. } => Some("rankDescending"),
        _ => None,
    }
}

fn show_as_base_field_index(
    show_as: &PivotShowAs,
    fields: &[CacheField],
) -> XlsxResult<Option<usize>> {
    let base_field = match show_as {
        PivotShowAs::RunningTotal { base_field }
        | PivotShowAs::RankAscending { base_field }
        | PivotShowAs::RankDescending { base_field }
        | PivotShowAs::DifferenceFrom { base_field, .. }
        | PivotShowAs::PercentDifferenceFrom { base_field, .. } => &base_field.name,
        _ => return Ok(None),
    };
    field_index(fields, base_field).map(Some).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot base field not found: {base_field}"))
    })
}

fn show_as_base_item_index(
    show_as: &PivotShowAs,
    fields: &[CacheField],
) -> XlsxResult<Option<u32>> {
    let (base_field, base_item) = match show_as {
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        }
        | PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => (&base_field.name, base_item),
        _ => return Ok(None),
    };
    let field_index = field_index(fields, base_field).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot base field not found: {base_field}"))
    })?;
    fields[field_index]
        .item_lookup
        .get(base_item)
        .copied()
        .map(Some)
        .ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot base item not found in field {base_field}: {base_item}"
            ))
        })
}

fn write_pivot_style(w: &mut XmlWriter, pivot: &PivotTable) -> XlsxResult<()> {
    let mut style = BytesStart::new("pivotTableStyleInfo");
    if let Some(name) = &pivot.style.name {
        style.push_attribute(("name", name.as_str()));
    }
    style.push_attribute(("showRowHeaders", bool_attr(pivot.style.show_row_headers)));
    style.push_attribute(("showColHeaders", bool_attr(pivot.style.show_column_headers)));
    style.push_attribute(("showRowStripes", bool_attr(pivot.style.show_row_stripes)));
    style.push_attribute(("showColStripes", bool_attr(pivot.style.show_column_stripes)));
    style.push_attribute(("showLastColumn", bool_attr(pivot.style.show_last_column)));
    w.write_event(Event::Empty(style))?;
    Ok(())
}

pub(super) fn write_pivot_cache_definition_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    part: &PivotCachePart,
) -> XlsxResult<()> {
    let path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", part.cache_num);
    write_xml_part(zip, &path, |w| {
        let record_count = part.record_count.to_string();
        let refresh_on_load = bool_attr(part.refresh_on_load);
        let background_query = bool_attr(part.background_query);
        let save_data = bool_attr(part.save_data);
        let mut tag = BytesStart::new("pivotCacheDefinition");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        tag.push_attribute(("r:id", "rId1"));
        tag.push_attribute(("recordCount", record_count.as_str()));
        tag.push_attribute(("saveData", save_data));
        tag.push_attribute(("refreshOnLoad", refresh_on_load));
        tag.push_attribute(("backgroundQuery", background_query));
        let missing_items_limit = part.missing_items_limit.map(|limit| limit.to_string());
        if let Some(missing_items_limit) = missing_items_limit.as_deref() {
            tag.push_attribute(("missingItemsLimit", missing_items_limit));
        }
        tag.push_attribute(("createdVersion", "8"));
        tag.push_attribute(("refreshedVersion", "8"));
        tag.push_attribute(("minRefreshableVersion", "3"));
        w.write_event(Event::Start(tag))?;

        write_cache_source(w, workbook, part)?;

        let count = part.fields.len().to_string();
        let mut cache_fields = BytesStart::new("cacheFields");
        cache_fields.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(cache_fields))?;
        for field in &part.fields {
            write_cache_field(w, field)?;
        }
        w.write_event(Event::End(BytesEnd::new("cacheFields")))?;

        w.write_event(Event::End(BytesEnd::new("pivotCacheDefinition")))?;
        Ok(())
    })
}

fn cache_source_tag(source_type: &'static str) -> BytesStart<'static> {
    let mut cache_source = BytesStart::new("cacheSource");
    cache_source.push_attribute(("type", source_type));
    cache_source
}

fn write_cache_source(
    w: &mut XmlWriter,
    workbook: &Workbook,
    part: &PivotCachePart,
) -> XlsxResult<()> {
    match &part.source {
        PivotSource::WorksheetRange { .. } | PivotSource::Table { .. } => {
            w.write_event(Event::Start(cache_source_tag("worksheet")))?;
            write_worksheet_source(w, workbook, part)?;
            w.write_event(Event::End(BytesEnd::new("cacheSource")))?;
        }
        PivotSource::External {
            connection_name, ..
        } => {
            let mut cache_source = cache_source_tag("external");
            if let Some(connection_id) = connection_id_attr(workbook, connection_name)? {
                cache_source.push_attribute(("connectionId", connection_id.as_str()));
            }
            w.write_event(Event::Empty(cache_source))?;
        }
        PivotSource::Consolidation { ranges } => {
            w.write_event(Event::Start(cache_source_tag("consolidation")))?;
            write_consolidation_source(w, ranges)?;
            w.write_event(Event::End(BytesEnd::new("cacheSource")))?;
        }
        PivotSource::Scenario { .. } => {
            w.write_event(Event::Empty(cache_source_tag("scenario")))?;
        }
        PivotSource::Olap { .. } => {
            return Err(XlsxError::InvalidFormat(
                "XLSX OLAP pivot source writing is not supported yet".into(),
            ));
        }
    }
    Ok(())
}

fn connection_id_attr(workbook: &Workbook, connection_name: &str) -> XlsxResult<Option<String>> {
    if connection_name.is_empty() {
        return Ok(None);
    }
    if let Some(connection) = find_workbook_connection(workbook, connection_name) {
        return Ok(Some(connection.id.to_string()));
    }
    let connection_id = connection_name.parse::<u32>().map_err(|_| {
        XlsxError::InvalidFormat(format!(
            "XLSX pivot external source connectionId must be numeric or match a workbook data connection: {connection_name}"
        ))
    })?;
    Ok(Some(connection_id.to_string()))
}

fn write_consolidation_source(w: &mut XmlWriter, ranges: &[PivotSourceRange]) -> XlsxResult<()> {
    if ranges.is_empty() {
        return Err(XlsxError::InvalidFormat(
            "XLSX consolidation pivot sources require at least one range".into(),
        ));
    }

    let mut consolidation = BytesStart::new("consolidation");
    consolidation.push_attribute(("autoPage", "0"));
    w.write_event(Event::Start(consolidation))?;

    let count = ranges.len().to_string();
    let mut range_sets = BytesStart::new("rangeSets");
    range_sets.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(range_sets))?;
    for range in ranges {
        let ref_str = range.range.to_a1_string();
        let mut range_set = BytesStart::new("rangeSet");
        range_set.push_attribute(("ref", ref_str.as_str()));
        range_set.push_attribute(("sheet", range.sheet.as_str()));
        w.write_event(Event::Empty(range_set))?;
    }
    w.write_event(Event::End(BytesEnd::new("rangeSets")))?;

    w.write_event(Event::End(BytesEnd::new("consolidation")))?;
    Ok(())
}

fn write_worksheet_source(
    w: &mut XmlWriter,
    workbook: &Workbook,
    part: &PivotCachePart,
) -> XlsxResult<()> {
    let mut source = BytesStart::new("worksheetSource");
    match &part.source {
        PivotSource::WorksheetRange { sheet, range } => {
            let sheet_name = sheet
                .as_deref()
                .or_else(|| {
                    workbook
                        .worksheet(part.source_sheet_index)
                        .map(Worksheet::name)
                })
                .unwrap_or("Sheet1");
            let ref_str = range.to_a1_string();
            source.push_attribute(("ref", ref_str.as_str()));
            source.push_attribute(("sheet", sheet_name));
        }
        PivotSource::Table { name } => {
            source.push_attribute(("name", name.as_str()));
        }
        _ => {}
    }
    w.write_event(Event::Empty(source))?;
    Ok(())
}

fn write_cache_field(w: &mut XmlWriter, field: &CacheField) -> XlsxResult<()> {
    let mut cache_field = BytesStart::new("cacheField");
    cache_field.push_attribute(("name", field.name.as_str()));
    if matches!(
        field.group,
        Some(CacheFieldGroup::Base { .. } | CacheFieldGroup::Manual { .. })
    ) {
        cache_field.push_attribute(("numFmtId", "0"));
    }
    if let Some(formula) = &field.formula {
        cache_field.push_attribute(("formula", formula.as_str()));
    }
    if !field.database_field {
        cache_field.push_attribute(("databaseField", "0"));
    }
    w.write_event(Event::Start(cache_field))?;

    if !matches!(field.group, Some(CacheFieldGroup::Manual { .. }))
        || !field.shared_items.is_empty()
    {
        let mut shared_items = BytesStart::new("sharedItems");
        let metadata_only = shared_items_are_metadata_only(field);
        let count = field.shared_items.len().to_string();
        if !metadata_only {
            shared_items.push_attribute(("count", count.as_str()));
        }
        shared_items.push_attribute(("containsBlank", bool_attr(field_contains_blank(field))));
        shared_items.push_attribute(("containsString", bool_attr(field_contains_string(field))));
        shared_items.push_attribute(("containsNumber", bool_attr(field_contains_number(field))));
        shared_items.push_attribute(("containsMixedTypes", bool_attr(field_contains_mixed(field))));
        let (contains_integer, min_value, max_value) = numeric_shared_item_stats(field);
        if let Some(contains_integer) = contains_integer {
            shared_items.push_attribute(("containsInteger", bool_attr(contains_integer)));
        }
        if let Some(min_value) = min_value.as_deref() {
            shared_items.push_attribute(("minValue", min_value));
        }
        if let Some(max_value) = max_value.as_deref() {
            shared_items.push_attribute(("maxValue", max_value));
        }
        w.write_event(Event::Start(shared_items))?;
        if !metadata_only {
            for value in &field.shared_items {
                write_pivot_value(w, value)?;
            }
        }
        w.write_event(Event::End(BytesEnd::new("sharedItems")))?;
    }

    if let Some(grouping) = &field.group {
        write_field_group(w, grouping)?;
    }

    w.write_event(Event::End(BytesEnd::new("cacheField")))?;
    Ok(())
}

fn shared_items_are_metadata_only(field: &CacheField) -> bool {
    field.metadata_only_shared_items
        && field.group.is_none()
        && !field.shared_items.is_empty()
        && field
            .shared_items
            .iter()
            .all(|value| matches!(value, PivotValue::Number(_)))
}

fn numeric_shared_item_stats(field: &CacheField) -> (Option<bool>, Option<String>, Option<String>) {
    let mut numbers = field.shared_items.iter().filter_map(|value| match value {
        PivotValue::Number(value) if value.is_finite() => Some(*value),
        _ => None,
    });
    let Some(first) = numbers.next() else {
        return (None, None, None);
    };

    let mut min = first;
    let mut max = first;
    let mut contains_integer = first.fract() == 0.0;
    for value in numbers {
        min = min.min(value);
        max = max.max(value);
        contains_integer &= value.fract() == 0.0;
    }

    (
        Some(contains_integer),
        Some(min.to_string()),
        Some(max.to_string()),
    )
}

fn write_field_group(w: &mut XmlWriter, grouping: &CacheFieldGroup) -> XlsxResult<()> {
    let mut field_group = BytesStart::new("fieldGroup");
    let base = match grouping {
        CacheFieldGroup::Base { .. } => None,
        CacheFieldGroup::DateUnit { base, .. } => Some(base.to_string()),
        CacheFieldGroup::Manual { base, .. } => Some(base.to_string()),
        CacheFieldGroup::Range(_) => None,
    };
    let parent = match grouping {
        CacheFieldGroup::Base { parent } => Some(parent.to_string()),
        CacheFieldGroup::DateUnit {
            parent: Some(parent),
            ..
        } => Some(parent.to_string()),
        _ => None,
    };
    if let Some(base) = base.as_deref() {
        field_group.push_attribute(("base", base));
    }
    if let Some(parent) = parent.as_deref() {
        field_group.push_attribute(("par", parent));
    }

    if let CacheFieldGroup::Base { .. } = grouping {
        w.write_event(Event::Empty(field_group))?;
        return Ok(());
    }

    w.write_event(Event::Start(field_group))?;

    if let CacheFieldGroup::Manual {
        item_indexes,
        group_items,
        ..
    } = grouping
    {
        let count = item_indexes.len().to_string();
        let mut discrete_pr = BytesStart::new("discretePr");
        discrete_pr.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(discrete_pr))?;
        for index in item_indexes {
            let value = index.to_string();
            let mut x = BytesStart::new("x");
            x.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(x))?;
        }
        w.write_event(Event::End(BytesEnd::new("discretePr")))?;

        let count = group_items.len().to_string();
        let mut group_items_el = BytesStart::new("groupItems");
        group_items_el.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(group_items_el))?;
        for value in group_items {
            write_pivot_value(w, value)?;
        }
        w.write_event(Event::End(BytesEnd::new("groupItems")))?;
        w.write_event(Event::End(BytesEnd::new("fieldGroup")))?;
        return Ok(());
    }

    let mut range_pr = BytesStart::new("rangePr");
    match grouping {
        CacheFieldGroup::Base { .. } => unreachable!("base field groups return before rangePr"),
        CacheFieldGroup::Range(PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        }) => {
            let auto_start = bool_attr(start.is_none());
            let auto_end = bool_attr(end.is_none());
            let start_num = start.map(|value| value.to_string());
            let end_num = end.map(|value| value.to_string());
            let group_interval = interval.to_string();

            range_pr.push_attribute(("autoStart", auto_start));
            range_pr.push_attribute(("autoEnd", auto_end));
            range_pr.push_attribute(("groupBy", "range"));
            if let Some(start_num) = &start_num {
                range_pr.push_attribute(("startNum", start_num.as_str()));
            }
            if let Some(end_num) = &end_num {
                range_pr.push_attribute(("endNum", end_num.as_str()));
            }
            range_pr.push_attribute(("groupInterval", group_interval.as_str()));
        }
        CacheFieldGroup::Range(PivotGrouping::Date { units, .. }) => {
            range_pr.push_attribute(("groupBy", date_group_by_name(units[0])));
        }
        CacheFieldGroup::Range(PivotGrouping::Manual { .. }) => {
            unreachable!("manual pivot groups use CacheFieldGroup::Manual")
        }
        CacheFieldGroup::Manual { .. } => unreachable!("manual groups return before rangePr"),
        CacheFieldGroup::DateUnit { unit, .. } => {
            range_pr.push_attribute(("groupBy", date_group_by_name(*unit)));
        }
    }
    w.write_event(Event::Empty(range_pr))?;

    w.write_event(Event::End(BytesEnd::new("fieldGroup")))?;
    Ok(())
}

fn date_group_by_name(unit: PivotDateGroupUnit) -> &'static str {
    match unit {
        PivotDateGroupUnit::Seconds => "seconds",
        PivotDateGroupUnit::Minutes => "minutes",
        PivotDateGroupUnit::Hours => "hours",
        PivotDateGroupUnit::Days => "days",
        PivotDateGroupUnit::Months => "months",
        PivotDateGroupUnit::Quarters => "quarters",
        PivotDateGroupUnit::Years => "years",
    }
}

pub(super) fn write_pivot_cache_records_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    _workbook: &Workbook,
    part: &PivotCachePart,
) -> XlsxResult<()> {
    let path = format!("xl/pivotCache/pivotCacheRecords{}.xml", part.cache_num);
    write_xml_part(zip, &path, |w| {
        let count = part.record_count.to_string();
        let mut records = BytesStart::new("pivotCacheRecords");
        records.push_attribute(("xmlns", NS_SPREADSHEET));
        records.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(records))?;

        for row in &part.rows {
            w.write_event(Event::Start(BytesStart::new("r")))?;
            for (field, value_index) in part.fields.iter().zip(row) {
                write_cache_record_value(w, field, *value_index)?;
            }
            w.write_event(Event::End(BytesEnd::new("r")))?;
        }

        w.write_event(Event::End(BytesEnd::new("pivotCacheRecords")))?;
        Ok(())
    })
}

fn write_cache_record_value(
    w: &mut XmlWriter,
    field: &CacheField,
    value_index: Option<u32>,
) -> XlsxResult<()> {
    let Some(index) = value_index else {
        w.write_event(Event::Empty(BytesStart::new("m")))?;
        return Ok(());
    };

    let Some(value) = field.shared_items.get(index as usize) else {
        w.write_event(Event::Empty(BytesStart::new("m")))?;
        return Ok(());
    };

    match value {
        PivotValue::String(_) => {
            let value = index.to_string();
            let mut x = BytesStart::new("x");
            x.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(x))?;
        }
        _ => write_pivot_value(w, value)?,
    }
    Ok(())
}

pub(super) fn write_pivot_cache_definition_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    cache_num: usize,
) -> XlsxResult<()> {
    let path = format!(
        "xl/pivotCache/_rels/pivotCacheDefinition{}.xml.rels",
        cache_num
    );
    write_xml_part(zip, &path, |w| {
        let mut relationships = BytesStart::new("Relationships");
        relationships.push_attribute(("xmlns", NS_RELATIONSHIPS));
        w.write_event(Event::Start(relationships))?;

        let target = format!("pivotCacheRecords{}.xml", cache_num);
        w.create_element("Relationship")
            .with_attribute(("Id", "rId1"))
            .with_attribute(("Type", RT_PIVOT_CACHE_RECORDS))
            .with_attribute(("Target", target.as_str()))
            .write_empty()?;

        w.write_event(Event::End(BytesEnd::new("Relationships")))?;
        Ok(())
    })
}

fn write_pivot_value(w: &mut XmlWriter, value: &PivotValue) -> XlsxResult<()> {
    match value {
        PivotValue::Blank => {
            w.write_event(Event::Empty(BytesStart::new("m")))?;
        }
        PivotValue::Boolean(value) => {
            let mut tag = BytesStart::new("b");
            tag.push_attribute(("v", if *value { "1" } else { "0" }));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::Number(value) => {
            let number = value.to_string();
            let mut tag = BytesStart::new("n");
            tag.push_attribute(("v", number.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::String(value) => {
            let mut tag = BytesStart::new("s");
            tag.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::Error(value) => {
            let mut tag = BytesStart::new("e");
            tag.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
    }
    Ok(())
}

fn field_contains_blank(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::Blank))
}

fn field_contains_string(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::String(_)))
}

fn field_contains_number(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::Number(_)))
}

fn field_contains_mixed(field: &CacheField) -> bool {
    let mut kinds = std::collections::HashSet::new();
    for value in &field.shared_items {
        kinds.insert(std::mem::discriminant(value));
    }
    kinds.len() > 1
}

fn bool_attr(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn field_index(fields: &[CacheField], name: &str) -> Option<usize> {
    fields
        .iter()
        .position(|field| field.name.eq_ignore_ascii_case(name))
}

fn aggregate_name(aggregate: PivotAggregate) -> &'static str {
    match aggregate {
        PivotAggregate::Average => "average",
        PivotAggregate::Count => "count",
        PivotAggregate::CountNumbers => "countNums",
        PivotAggregate::Max => "max",
        PivotAggregate::Min => "min",
        PivotAggregate::Product => "product",
        PivotAggregate::StdDev => "stdDev",
        PivotAggregate::StdDevP => "stdDevp",
        PivotAggregate::Sum => "sum",
        PivotAggregate::Var => "var",
        PivotAggregate::VarP => "varp",
    }
}
