use crate::prelude::*;
use crate::runtime_cache::*;
use crate::source::*;
use crate::transform::*;

#[derive(Debug, Clone)]
pub(crate) struct SourceSnapshot {
    source_name: Option<String>,
    pub(crate) headers: Vec<String>,
    pub(crate) columns: Vec<EncodedColumn>,
    pub(crate) row_count: usize,
}

impl SourceSnapshot {
    pub(crate) fn from_resolved(workbook: &Workbook, source: &ResolvedPivotSource) -> Result<Self> {
        match source {
            ResolvedPivotSource::Single(source) => {
                let worksheet = workbook.worksheet(source.sheet_index).ok_or_else(|| {
                    Error::SheetOutOfBounds(source.sheet_index, workbook.sheet_count())
                })?;
                Self::from_single_source(worksheet, source)
            }
            ResolvedPivotSource::Consolidation(sources) => {
                Self::from_consolidation(workbook, sources)
            }
        }
    }

    pub(crate) fn from_single_source(
        worksheet: &Worksheet,
        source: &ResolvedSource,
    ) -> Result<Self> {
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
            source_name: matches!(source.kind, SourceCacheKind::Table)
                .then(|| source.source_name.clone())
                .flatten(),
            headers,
            columns,
            row_count,
        })
    }

    pub(crate) fn from_consolidation(
        workbook: &Workbook,
        sources: &[ResolvedSource],
    ) -> Result<Self> {
        if sources.is_empty() {
            return Err(Error::other(
                "consolidation pivot sources must contain at least one range",
            ));
        }

        let first = worksheet_for_source(workbook, &sources[0])?;
        let headers = read_headers_from_sheet(first, sources[0].range)?;
        validate_headers(&headers)?;

        let col_count = headers.len();
        let mut row_count = 0usize;
        for source in sources {
            let worksheet = worksheet_for_source(workbook, source)?;
            let source_headers = read_headers_from_sheet(worksheet, source.range)?;
            validate_headers(&source_headers)?;
            if !same_headers(&headers, &source_headers) {
                return Err(Error::other(
                    "consolidation pivot source ranges must have matching headers",
                ));
            }
            row_count += source_data_row_count(source);
        }

        let columns = consolidation_snapshot_columns(workbook, sources, col_count, row_count)?;
        Ok(Self {
            source_name: None,
            headers,
            columns,
            row_count,
        })
    }

    pub(crate) fn field_index(&self, name: &str) -> Option<usize> {
        self.headers
            .iter()
            .position(|header| header.eq_ignore_ascii_case(name))
    }

    pub(crate) fn value(&self, row: usize, col: usize) -> &PivotValue {
        self.columns[col].value(row)
    }

    pub(crate) fn value_by_id(&self, col: usize, id: u32) -> &PivotValue {
        self.columns[col].value_by_id(id)
    }

    pub(crate) fn apply_calculated_fields(
        &self,
        pivot_name: &str,
        calculated_fields: &[PivotCalculatedField],
        workbook_context: Option<CalculatedWorkbookContext<'_>>,
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
                self.source_name.as_deref(),
                workbook_context,
            )?;
            let mut column = EncodedColumn::with_capacity(self.row_count);
            for value in values {
                column.push(value);
            }
            headers.push(field.name.clone());
            columns.push(column);
        }

        Ok(Self {
            source_name: self.source_name.clone(),
            headers,
            columns,
            row_count: self.row_count,
        })
    }

    pub(crate) fn apply_groupings(
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
                        columns.push(self.grouped_date_column(field_index, &[*unit], date_1904));
                    }
                }
                _ => {
                    columns[field_index] =
                        self.grouped_column(field_index, grouping, date_1904, pivot_name)?;
                }
            }
        }

        Ok(Self {
            source_name: self.source_name.clone(),
            headers,
            columns,
            row_count: self.row_count,
        })
    }

    pub(crate) fn apply_calculated_items(
        &self,
        pivot_name: &str,
        calculated_items: &[PivotCalculatedItem],
    ) -> Result<Self> {
        let headers = self.headers.clone();
        let mut columns = self.columns.clone();

        for item in calculated_items {
            if item.formula.trim().is_empty() {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} calculated item {} has a blank formula",
                    item.item
                )));
            }
            let field_index = self.field_index(&item.field.name).ok_or_else(|| {
                Error::other(format!(
                    "pivot table {pivot_name} calculated item references unknown field: {}",
                    item.field.name
                ))
            })?;
            if let Some(existing_id) = columns[field_index].id_for_value(&item.item) {
                if columns[field_index].values.contains(&existing_id) {
                    return Err(Error::other(format!(
                        "pivot table {pivot_name} calculated item {} duplicates a source item in field {}",
                        item.item, item.field.name
                    )));
                }
            }
            columns[field_index].ensure_dictionary_value(item.item.clone());
        }

        Ok(Self {
            source_name: self.source_name.clone(),
            headers,
            columns,
            row_count: self.row_count,
        })
    }
}

#[derive(Debug, Clone)]
pub(crate) struct EncodedColumn {
    pub(crate) values: Vec<u32>,
    pub(crate) dictionary: Vec<PivotValue>,
    lookup: AHashMap<PivotValue, u32>,
}

impl EncodedColumn {
    pub(crate) fn with_capacity(capacity: usize) -> Self {
        Self {
            values: Vec::with_capacity(capacity),
            dictionary: Vec::new(),
            lookup: AHashMap::new(),
        }
    }

    pub(crate) fn push(&mut self, value: PivotValue) {
        let id = self.ensure_dictionary_value(value);
        self.values.push(id);
    }

    pub(crate) fn ensure_dictionary_value(&mut self, value: PivotValue) -> u32 {
        let value = normalize_dictionary_value(value);
        if let Some(id) = self.lookup.get(&value) {
            *id
        } else {
            let id = self.dictionary.len() as u32;
            self.dictionary.push(value.clone());
            self.lookup.insert(value, id);
            id
        }
    }

    pub(crate) fn remap_dictionary<F>(&self, group_value: F) -> Self
    where
        F: Fn(&PivotValue) -> PivotValue,
    {
        let mut dictionary = Vec::new();
        let mut lookup = AHashMap::new();
        let mut id_map = Vec::with_capacity(self.dictionary.len());

        for value in &self.dictionary {
            let grouped = normalize_dictionary_value(group_value(value));
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

        let values = Self::remap_ids(&self.values, &id_map);
        Self {
            values,
            dictionary,
            lookup,
        }
    }

    pub(crate) fn id_at(&self, row: usize) -> u32 {
        self.values[row]
    }

    pub(crate) fn value(&self, row: usize) -> &PivotValue {
        self.value_by_id(self.id_at(row))
    }

    pub(crate) fn value_by_id(&self, id: u32) -> &PivotValue {
        &self.dictionary[id as usize]
    }

    pub(crate) fn id_for_value(&self, value: &PivotValue) -> Option<u32> {
        match value {
            PivotValue::Number(value) if *value == 0.0 => {
                self.lookup.get(&PivotValue::Number(0.0)).copied()
            }
            value => self.lookup.get(value).copied(),
        }
    }

    fn remap_ids(values: &[u32], id_map: &[u32]) -> Vec<u32> {
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
}

fn normalize_dictionary_value(value: PivotValue) -> PivotValue {
    match value {
        PivotValue::Number(value) if value == 0.0 => PivotValue::Number(0.0),
        value => value,
    }
}
