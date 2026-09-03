use crate::prelude::*;
use crate::runtime_cache::*;
use crate::snapshot::*;

#[derive(Debug, Clone)]
pub(crate) enum ResolvedPivotSource {
    Single(ResolvedSource),
    Consolidation(Vec<ResolvedSource>),
}

impl ResolvedPivotSource {
    pub(crate) fn cache_key(&self) -> SourceCacheKey {
        match self {
            Self::Single(source) => SourceCacheKey::Single(source.cache_key()),
            Self::Consolidation(sources) => SourceCacheKey::Consolidation(
                sources.iter().map(ResolvedSource::cache_key).collect(),
            ),
        }
    }
}

pub(crate) fn resolve_source(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    source: &PivotSource,
) -> Result<ResolvedPivotSource> {
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

            Ok(ResolvedPivotSource::Single(ResolvedSource {
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
            }))
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

            Ok(ResolvedPivotSource::Single(ResolvedSource {
                kind: SourceCacheKind::Table,
                sheet_index,
                range: table.reference,
                source_name: Some(table.name.clone()),
                headers: Some(headers),
                data_start_row,
                data_end_row,
                mutation_count: worksheet.mutation_count(),
                topology_generation: worksheet.topology_generation(),
            }))
        }
        PivotSource::External { .. } => Err(Error::other(
            "external pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
        PivotSource::Consolidation { ranges } => resolve_consolidation_source(workbook, ranges),
        PivotSource::Scenario { .. } => Err(Error::other(
            "scenario pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
        PivotSource::Olap { .. } => Err(Error::other(
            "OLAP pivot sources are preserved but cannot be refreshed by the local engine yet",
        )),
    }
}

pub(crate) fn resolve_consolidation_source(
    workbook: &Workbook,
    ranges: &[duke_sheets_core::PivotSourceRange],
) -> Result<ResolvedPivotSource> {
    if ranges.is_empty() {
        return Err(Error::other(
            "consolidation pivot sources must contain at least one range",
        ));
    }

    let mut resolved = Vec::with_capacity(ranges.len());
    for range in ranges {
        let (Some(sheet), Some(source_range)) = (&range.sheet, range.range) else {
            return Err(Error::other(
                "named or external consolidation pivot sources cannot be refreshed by the local engine yet",
            ));
        };
        if source_range.row_count() == 0 || source_range.col_count() == 0 {
            return Err(Error::other(
                "consolidation pivot source range cannot be empty",
            ));
        }
        let sheet_index = workbook
            .sheet_index(sheet)
            .ok_or_else(|| Error::SheetNotFound(sheet.clone()))?;
        let worksheet = workbook
            .worksheet(sheet_index)
            .ok_or_else(|| Error::SheetOutOfBounds(sheet_index, workbook.sheet_count()))?;
        resolved.push(ResolvedSource {
            kind: SourceCacheKind::ConsolidationRange,
            sheet_index,
            range: source_range,
            source_name: Some(consolidation_range_cache_name(range)),
            headers: None,
            data_start_row: source_range.start.row.saturating_add(1),
            data_end_row: if source_range.end.row > source_range.start.row {
                Some(source_range.end.row)
            } else {
                None
            },
            mutation_count: worksheet.mutation_count(),
            topology_generation: worksheet.topology_generation(),
        });
    }

    Ok(ResolvedPivotSource::Consolidation(resolved))
}

pub(crate) fn consolidation_range_cache_name(range: &duke_sheets_core::PivotSourceRange) -> String {
    let mut name = String::new();
    if let Some(sheet) = &range.sheet {
        name.push_str(sheet);
    }
    if let Some(source_range) = range.range {
        name.push('!');
        name.push_str(&source_range.to_a1_string());
    }
    if let Some(display_name) = &range.name {
        name.push('\u{1f}');
        name.push_str(display_name);
    }
    if let Some(external_relationship_id) = &range.external_relationship_id {
        name.push('\u{1f}');
        name.push_str(external_relationship_id);
    }
    if let Some(external_relationship_target) = &range.external_relationship_target {
        name.push('\u{1f}');
        name.push_str(external_relationship_target);
    }
    for page_item in &range.page_items {
        name.push('\u{1f}');
        name.push_str(page_item);
    }
    name
}

pub(crate) fn find_table<'a>(
    workbook: &'a Workbook,
    name: &str,
) -> Option<(usize, &'a Worksheet, &'a Table)> {
    workbook
        .worksheets()
        .enumerate()
        .find_map(|(sheet_index, worksheet)| {
            worksheet
                .table_by_name(name)
                .map(|table| (sheet_index, worksheet, table))
        })
}

pub(crate) fn table_headers(table: &Table) -> Vec<String> {
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

pub(crate) fn table_data_end_row(table: &Table) -> Option<u32> {
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

pub(crate) fn worksheet_for_source<'a>(
    workbook: &'a Workbook,
    source: &ResolvedSource,
) -> Result<&'a Worksheet> {
    workbook
        .worksheet(source.sheet_index)
        .ok_or_else(|| Error::SheetOutOfBounds(source.sheet_index, workbook.sheet_count()))
}

pub(crate) fn same_headers(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| left.eq_ignore_ascii_case(right))
}

pub(crate) fn source_data_row_count(source: &ResolvedSource) -> usize {
    source
        .data_end_row
        .map(|end_row| (end_row - source.data_start_row + 1) as usize)
        .unwrap_or(0)
}

pub(crate) fn consolidation_snapshot_columns(
    workbook: &Workbook,
    sources: &[ResolvedSource],
    col_count: usize,
    row_count: usize,
) -> Result<Vec<EncodedColumn>> {
    #[cfg(feature = "parallel")]
    {
        if row_count >= PARALLEL_ROW_THRESHOLD {
            return (0..col_count)
                .into_par_iter()
                .map(|col_offset| {
                    consolidation_snapshot_column(workbook, sources, row_count, col_offset)
                })
                .collect();
        }
    }

    (0..col_count)
        .map(|col_offset| consolidation_snapshot_column(workbook, sources, row_count, col_offset))
        .collect()
}

pub(crate) fn consolidation_snapshot_column(
    workbook: &Workbook,
    sources: &[ResolvedSource],
    row_count: usize,
    col_offset: usize,
) -> Result<EncodedColumn> {
    let mut column = EncodedColumn::with_capacity(row_count);
    for source in sources {
        let Some(data_end_row) = source.data_end_row else {
            continue;
        };
        let worksheet = worksheet_for_source(workbook, source)?;
        let source_col = source.range.start.col + col_offset as u16;
        for row in source.data_start_row..=data_end_row {
            column.push(effective_pivot_value(worksheet, row, source_col));
        }
    }
    Ok(column)
}

pub(crate) fn source_snapshot_columns(
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

pub(crate) fn source_snapshot_column(
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

pub(crate) fn unique_grouped_header(
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

pub(crate) fn grouped_date_header(
    field_name: &str,
    unit: duke_sheets_core::PivotDateGroupUnit,
) -> String {
    format!("{field_name} ({})", date_group_unit_name(unit))
}

pub(crate) fn date_group_unit_name(unit: duke_sheets_core::PivotDateGroupUnit) -> &'static str {
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

pub(crate) fn normalize_supplied_headers(headers: &[String], col_count: usize) -> Vec<String> {
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

pub(crate) fn read_headers_from_sheet(
    worksheet: &Worksheet,
    range: CellRange,
) -> Result<Vec<String>> {
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

pub(crate) fn validate_headers(headers: &[String]) -> Result<()> {
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

pub(crate) fn effective_pivot_value(worksheet: &Worksheet, row: u32, col: u16) -> PivotValue {
    worksheet
        .get_calculated_value_at(row, col)
        .map(PivotValue::from_cell_value)
        .unwrap_or_else(|| PivotValue::from_cell_value(&worksheet.get_value_at(row, col)))
}
