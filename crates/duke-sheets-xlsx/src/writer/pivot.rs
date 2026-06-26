use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, Event};

use duke_sheets_core::{
    CellError, CellRange, PivotAggregate, PivotCalculatedField, PivotDateGroupUnit, PivotFieldRef,
    PivotFilter, PivotGrouping, PivotLayoutKind, PivotShowAs, PivotSort, PivotSource, PivotTable,
    PivotValue, Table, Workbook, Worksheet,
};
use duke_sheets_formula::{
    evaluate, parse_formula, EvaluationContext, FormulaExpr, FormulaValue, StructuredRefSpecifier,
    StructuredReference,
};

use super::{
    write_xml_part, XlsxError, XlsxResult, XmlWriter, NS_DOC_RELS, NS_RELATIONSHIPS,
    NS_SPREADSHEET, RT_PIVOT_CACHE_RECORDS,
};

const NS_SPREADSHEET_X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXT_URI_X14_DATA_FIELD: &str = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}";

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
    groupings: Vec<PivotGrouping>,
    rows: Vec<Vec<Option<u32>>>,
    record_count: usize,
    refresh_on_load: bool,
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
}

#[derive(Debug, Clone)]
struct CacheField {
    name: String,
    formula: Option<String>,
    database_field: bool,
    shared_items: Vec<PivotValue>,
    item_lookup: HashMap<PivotValue, u32>,
}

impl CacheField {
    fn new(name: String) -> Self {
        Self {
            name,
            formula: None,
            database_field: true,
            shared_items: Vec::new(),
            item_lookup: HashMap::new(),
        }
    }

    fn calculated(name: String, formula: String) -> Self {
        Self {
            name,
            formula: Some(formula),
            database_field: false,
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

            let mut resolved = resolve_pivot_source(workbook, sheet_index, &pivot.source)?;
            apply_calculated_cache_fields(&pivot.name, &mut resolved, &pivot.calculated_fields)?;
            validate_pivot_fields(pivot, &resolved.fields)?;
            validate_pivot_groupings(pivot, &resolved.fields)?;
            let cache_key = cache_key_for_pivot(&resolved.key, pivot);

            let cache_num = if let Some(cache_num) = cache_by_source.get(&cache_key) {
                if pivot.refresh_policy.refresh_on_open {
                    if let Some(cache_part) = cache_parts.get_mut(*cache_num - 1) {
                        cache_part.refresh_on_load = true;
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
                    groupings: pivot.groupings.clone(),
                    rows: resolved.rows,
                    record_count: resolved.record_count,
                    refresh_on_load: pivot.refresh_policy.refresh_on_open,
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

    if pivot.filters.iter().any(|filter| {
        !matches!(
            filter,
            PivotFilter::FieldItems {
                field: _,
                allowed_items: _
            }
        )
    }) {
        return Err(XlsxError::InvalidFormat(format!(
            "pivot table {} uses a filter type that is not written yet",
            pivot.name
        )));
    }

    Ok(())
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
                if units.len() != 1 {
                    return Err(XlsxError::InvalidFormat(format!(
                        "pivot table {} uses multi-unit date grouping for field {}, which this XLSX writer does not emit yet",
                        pivot.name, field.name
                    )));
                }
            }
        }
    }

    Ok(())
}

fn grouping_field_ref(grouping: &PivotGrouping) -> &PivotFieldRef {
    match grouping {
        PivotGrouping::Number { field, .. } | PivotGrouping::Date { field, .. } => field,
    }
}

fn cache_key_for_pivot(source_key: &str, pivot: &PivotTable) -> String {
    if pivot.groupings.is_empty() && pivot.calculated_fields.is_empty() {
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
        "{source_key}|calculated:{calculated_signatures}|groupings:{}",
        grouping_signatures.join(";")
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

fn resolve_pivot_source(
    workbook: &Workbook,
    pivot_sheet_index: usize,
    source: &PivotSource,
) -> XlsxResult<ResolvedPivotSource> {
    match source {
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
            })
        }
        PivotSource::External { .. }
        | PivotSource::Consolidation { .. }
        | PivotSource::Scenario { .. }
        | PivotSource::Olap { .. } => Err(XlsxError::InvalidFormat(
            "this XLSX writer currently supports worksheet and table pivot sources".into(),
        )),
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
        for row in data_start..=data_end {
            let mut record = Vec::with_capacity(fields.len());
            for (offset, field) in fields.iter_mut().enumerate() {
                let value = effective_pivot_value(worksheet, row, start_col + offset as u16);
                let index = field.intern(value);
                record.push(Some(index));
            }
            rows.push(record);
        }
    }

    let record_count = rows.len();
    Ok((fields, rows, record_count))
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
) -> XlsxResult<()> {
    let sheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = sheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table not found".into()))?;
    let cache_part = find_cache_part(workbook, part)?;

    let path = format!("xl/pivotTables/pivotTable{}.xml", part.table_num);
    write_xml_part(zip, &path, |w| {
        let cache_id = part.cache_num.to_string();
        let row_grand = bool_attr(pivot.layout.show_row_grand_totals);
        let col_grand = bool_attr(pivot.layout.show_column_grand_totals);
        let preserve_formatting = bool_attr(pivot.refresh_policy.preserve_formatting);
        let show_headers = bool_attr(pivot.layout.show_field_headers);
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
        tag.push_attribute(("compact", compact));
        tag.push_attribute(("outline", outline));
        w.write_event(Event::Start(tag))?;

        write_location(w, pivot)?;
        write_pivot_fields(w, pivot, &cache_part.fields)?;
        write_axis_fields(w, "rowFields", &pivot.rows, &cache_part.fields)?;
        write_axis_fields(w, "colFields", &pivot.columns, &cache_part.fields)?;
        write_page_fields(w, pivot, &cache_part.fields)?;
        write_data_fields(w, pivot, &cache_part.fields)?;
        write_pivot_style(w, pivot)?;

        w.write_event(Event::End(BytesEnd::new("pivotTableDefinition")))?;
        Ok(())
    })
}

fn find_cache_part(workbook: &Workbook, table_part: &PivotTablePart) -> XlsxResult<PivotCachePart> {
    let numbering = build_pivot_numbering(workbook)?;
    numbering
        .cache_parts
        .into_iter()
        .find(|part| part.cache_num == table_part.cache_num)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot cache part not found".into()))
}

fn write_location(w: &mut XmlWriter, pivot: &PivotTable) -> XlsxResult<()> {
    let range = pivot
        .rendered_range
        .unwrap_or_else(|| CellRange::single(pivot.target));
    let ref_str = range.to_a1_string();
    let first_data_col = pivot.rows.len().max(1).to_string();

    let mut location = BytesStart::new("location");
    location.push_attribute(("ref", ref_str.as_str()));
    location.push_attribute(("firstHeaderRow", "1"));
    location.push_attribute(("firstDataRow", "1"));
    location.push_attribute(("firstDataCol", first_data_col.as_str()));
    w.write_event(Event::Empty(location))?;
    Ok(())
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
        if let Some(axis) = field_axis(pivot, &field.name) {
            pivot_field.push_attribute(("axis", axis));
        }
        let sort = field_sort(pivot, &field.name);
        if sort != "manual" {
            pivot_field.push_attribute(("sortType", sort));
        }
        if field_is_filtered(pivot, &field.name) {
            pivot_field.push_attribute(("multipleItemSelectionAllowed", "1"));
        }
        if field_subtotal_is_none(pivot, &field.name) {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
        }

        let hidden_items = hidden_item_indexes(pivot, fields, index)?;
        if hidden_items.is_empty() {
            w.write_event(Event::Empty(pivot_field))?;
        } else {
            w.write_event(Event::Start(pivot_field))?;
            let count = field.shared_items.len().to_string();
            let mut items = BytesStart::new("items");
            items.push_attribute(("count", count.as_str()));
            w.write_event(Event::Start(items))?;
            for item_index in 0..field.shared_items.len() {
                let x = item_index.to_string();
                let mut item = BytesStart::new("item");
                item.push_attribute(("x", x.as_str()));
                if hidden_items.contains(&(item_index as u32)) {
                    item.push_attribute(("h", "1"));
                }
                w.write_event(Event::Empty(item))?;
            }
            w.write_event(Event::End(BytesEnd::new("items")))?;
            w.write_event(Event::End(BytesEnd::new("pivotField")))?;
        }
    }

    w.write_event(Event::End(BytesEnd::new("pivotFields")))?;
    Ok(())
}

fn field_axis(pivot: &PivotTable, field_name: &str) -> Option<&'static str> {
    if pivot
        .rows
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        Some("axisRow")
    } else if pivot
        .columns
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        Some("axisCol")
    } else if pivot
        .page_fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        Some("axisPage")
    } else {
        None
    }
}

fn field_sort(pivot: &PivotTable, field_name: &str) -> &'static str {
    let sort = pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(field_name))
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

fn field_subtotal_is_none(pivot: &PivotTable, field_name: &str) -> bool {
    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .any(|field| {
            field.field.name.eq_ignore_ascii_case(field_name)
                && matches!(field.subtotal, duke_sheets_core::PivotSubtotal::None)
        })
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

    let count = axis_fields.len().to_string();
    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(tag))?;
    for field in axis_fields {
        let index = field_index(fields, &field.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", field.field.name))
        })?;
        let x = index.to_string();
        let mut el = BytesStart::new("field");
        el.push_attribute(("x", x.as_str()));
        w.write_event(Event::Empty(el))?;
    }
    w.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
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
        let mut tag = BytesStart::new("pivotCacheDefinition");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        tag.push_attribute(("r:id", "rId1"));
        tag.push_attribute(("recordCount", record_count.as_str()));
        tag.push_attribute(("saveData", "1"));
        tag.push_attribute(("refreshOnLoad", refresh_on_load));
        tag.push_attribute(("createdVersion", "8"));
        tag.push_attribute(("refreshedVersion", "8"));
        tag.push_attribute(("minRefreshableVersion", "3"));
        w.write_event(Event::Start(tag))?;

        w.write_event(Event::Start(cache_source_tag()))?;
        write_worksheet_source(w, workbook, part)?;
        w.write_event(Event::End(BytesEnd::new("cacheSource")))?;

        let count = part.fields.len().to_string();
        let mut cache_fields = BytesStart::new("cacheFields");
        cache_fields.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(cache_fields))?;
        for field in &part.fields {
            write_cache_field(w, field, grouping_for_field(&part.groupings, &field.name))?;
        }
        w.write_event(Event::End(BytesEnd::new("cacheFields")))?;

        w.write_event(Event::End(BytesEnd::new("pivotCacheDefinition")))?;
        Ok(())
    })
}

fn cache_source_tag() -> BytesStart<'static> {
    let mut cache_source = BytesStart::new("cacheSource");
    cache_source.push_attribute(("type", "worksheet"));
    cache_source
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

fn write_cache_field(
    w: &mut XmlWriter,
    field: &CacheField,
    grouping: Option<&PivotGrouping>,
) -> XlsxResult<()> {
    let mut cache_field = BytesStart::new("cacheField");
    cache_field.push_attribute(("name", field.name.as_str()));
    if let Some(formula) = &field.formula {
        cache_field.push_attribute(("formula", formula.as_str()));
    }
    if !field.database_field {
        cache_field.push_attribute(("databaseField", "0"));
    }
    w.write_event(Event::Start(cache_field))?;

    let mut shared_items = BytesStart::new("sharedItems");
    let count = field.shared_items.len().to_string();
    shared_items.push_attribute(("count", count.as_str()));
    shared_items.push_attribute(("containsBlank", bool_attr(field_contains_blank(field))));
    shared_items.push_attribute(("containsString", bool_attr(field_contains_string(field))));
    shared_items.push_attribute(("containsNumber", bool_attr(field_contains_number(field))));
    shared_items.push_attribute(("containsMixedTypes", bool_attr(field_contains_mixed(field))));
    w.write_event(Event::Start(shared_items))?;
    for value in &field.shared_items {
        write_pivot_value(w, value)?;
    }
    w.write_event(Event::End(BytesEnd::new("sharedItems")))?;

    if let Some(grouping) = grouping {
        write_field_group(w, grouping)?;
    }

    w.write_event(Event::End(BytesEnd::new("cacheField")))?;
    Ok(())
}

fn grouping_for_field<'a>(
    groupings: &'a [PivotGrouping],
    field_name: &str,
) -> Option<&'a PivotGrouping> {
    groupings.iter().find(|grouping| {
        grouping_field_ref(grouping)
            .name
            .eq_ignore_ascii_case(field_name)
    })
}

fn write_field_group(w: &mut XmlWriter, grouping: &PivotGrouping) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("fieldGroup")))?;

    let mut range_pr = BytesStart::new("rangePr");
    match grouping {
        PivotGrouping::Number {
            start,
            end,
            interval,
            ..
        } => {
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
        PivotGrouping::Date { units, .. } => {
            range_pr.push_attribute(("groupBy", date_group_by_name(units[0])));
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
            for value_index in row {
                match value_index {
                    Some(index) => {
                        let value = index.to_string();
                        let mut x = BytesStart::new("x");
                        x.push_attribute(("v", value.as_str()));
                        w.write_event(Event::Empty(x))?;
                    }
                    None => {
                        w.write_event(Event::Empty(BytesStart::new("m")))?;
                    }
                }
            }
            w.write_event(Event::End(BytesEnd::new("r")))?;
        }

        w.write_event(Event::End(BytesEnd::new("pivotCacheRecords")))?;
        Ok(())
    })
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
