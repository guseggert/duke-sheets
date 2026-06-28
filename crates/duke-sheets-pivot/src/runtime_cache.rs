use crate::api::*;
use crate::prelude::*;
use crate::snapshot::*;
use crate::source::*;

#[derive(Default)]
pub(crate) struct PivotRuntimeCache {
    workbook_nonce: u64,
    structural_generation: u64,
    snapshots: AHashMap<SourceCacheKey, Arc<SourceSnapshot>>,
    transformed_snapshots: AHashMap<TransformedSnapshotCacheKey, Arc<SourceSnapshot>>,
    item_filter_baselines: AHashMap<PivotItemFilterBaselineKey, PivotItemFilterBaseline>,
}

impl PivotRuntimeCache {
    pub(crate) fn for_workbook(workbook: &Workbook) -> Self {
        Self {
            workbook_nonce: workbook.nonce(),
            structural_generation: workbook.structural_generation(),
            snapshots: AHashMap::new(),
            transformed_snapshots: AHashMap::new(),
            item_filter_baselines: AHashMap::new(),
        }
    }

    pub(crate) fn rebase_untouched_sources(
        &mut self,
        workbook: &Workbook,
        touched_ranges: &[(usize, CellRange)],
    ) {
        let mut snapshots = AHashMap::with_capacity(self.snapshots.len());
        for (mut key, snapshot) in std::mem::take(&mut self.snapshots) {
            if !key.rebase_untouched(workbook, touched_ranges) {
                continue;
            }
            snapshots.insert(key, snapshot);
        }
        let mut transformed_snapshots = AHashMap::with_capacity(self.transformed_snapshots.len());
        for (mut key, snapshot) in std::mem::take(&mut self.transformed_snapshots) {
            if !key.rebase_untouched(workbook, touched_ranges) {
                continue;
            }
            transformed_snapshots.insert(key, snapshot);
        }
        self.workbook_nonce = workbook.nonce();
        self.structural_generation = workbook.structural_generation();
        self.snapshots = snapshots;
        self.transformed_snapshots = transformed_snapshots;
    }

    pub(crate) fn take_from_workbook(workbook: &mut Workbook) -> Self {
        let mut cache = workbook
            .take_pivot_runtime_cache()
            .and_then(|cache| {
                let cache: Box<dyn Any + Send + Sync> = cache;
                cache.downcast::<Self>().ok()
            })
            .map(|cache| *cache)
            .unwrap_or_default();

        if cache.workbook_nonce != workbook.nonce()
            || cache.structural_generation != workbook.structural_generation()
        {
            cache = Self::for_workbook(workbook);
        }

        cache
    }

    pub(crate) fn filter_baselines_for_pivot(
        &mut self,
        sheet_index: usize,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
    ) -> PivotFilterBaselines {
        let mut baselines = PivotFilterBaselines::default();
        for filter in &pivot.filters {
            let PivotFilter::FieldItems {
                field,
                allowed_items,
            } = filter
            else {
                continue;
            };
            if !pivot_axis_field_includes_new_items(pivot, &field.name) {
                continue;
            }
            let Some(field_index) = snapshot.field_index(&field.name) else {
                continue;
            };

            let current_items = snapshot.columns[field_index]
                .dictionary
                .iter()
                .cloned()
                .collect::<AHashSet<_>>();
            let allowed_items = allowed_items.iter().cloned().collect::<AHashSet<_>>();
            let key = PivotItemFilterBaselineKey::new(sheet_index, pivot, &field.name);
            let baseline =
                self.item_filter_baselines
                    .entry(key)
                    .or_insert_with(|| PivotItemFilterBaseline {
                        allowed_items: allowed_items.clone(),
                        known_items: current_items.clone(),
                    });
            if baseline.allowed_items != allowed_items {
                baseline.allowed_items = allowed_items;
                baseline.known_items = current_items;
            }
            baselines.insert(&field.name, baseline.known_items.clone());
        }
        baselines
    }

    pub(crate) fn snapshot_for_source(
        &mut self,
        workbook: &Workbook,
        pivot_sheet_index: usize,
        source: &PivotSource,
        stats: &mut PivotRefreshStats,
    ) -> Result<CachedSourceSnapshot> {
        let resolved = resolve_source(workbook, pivot_sheet_index, source)?;
        self.snapshot_for_resolved_source(workbook, resolved, stats)
    }

    pub(crate) fn snapshot_for_resolved_source(
        &mut self,
        workbook: &Workbook,
        resolved: ResolvedPivotSource,
        stats: &mut PivotRefreshStats,
    ) -> Result<CachedSourceSnapshot> {
        let cache_key = resolved.cache_key();

        if let Some(snapshot) = self.snapshots.get(&cache_key) {
            stats.cache_hits += 1;
            return Ok(CachedSourceSnapshot {
                key: cache_key,
                snapshot: Arc::clone(snapshot),
            });
        }

        let snapshot = Arc::new(SourceSnapshot::from_resolved(workbook, &resolved)?);
        self.snapshots
            .insert(cache_key.clone(), Arc::clone(&snapshot));
        stats.cache_misses += 1;
        Ok(CachedSourceSnapshot {
            key: cache_key,
            snapshot,
        })
    }

    pub(crate) fn transformed_snapshot(
        &self,
        key: &TransformedSnapshotCacheKey,
    ) -> Option<Arc<SourceSnapshot>> {
        self.transformed_snapshots.get(key).map(Arc::clone)
    }

    pub(crate) fn insert_transformed_snapshot(
        &mut self,
        key: TransformedSnapshotCacheKey,
        snapshot: Arc<SourceSnapshot>,
    ) {
        self.transformed_snapshots.insert(key, snapshot);
    }

    #[cfg(test)]
    pub(crate) fn snapshot_count(&self) -> usize {
        self.snapshots.len()
    }

    #[cfg(test)]
    pub(crate) fn transformed_snapshot_count(&self) -> usize {
        self.transformed_snapshots.len()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct PivotItemFilterBaselineKey {
    sheet_index: usize,
    pivot_name: String,
    field_name: String,
}

impl PivotItemFilterBaselineKey {
    pub(crate) fn new(sheet_index: usize, pivot: &PivotTable, field_name: &str) -> Self {
        Self {
            sheet_index,
            pivot_name: pivot.name.to_lowercase(),
            field_name: field_name.to_lowercase(),
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct PivotItemFilterBaseline {
    allowed_items: AHashSet<PivotValue>,
    known_items: AHashSet<PivotValue>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct PivotFilterBaselines {
    known_items_by_field: AHashMap<String, AHashSet<PivotValue>>,
}

impl PivotFilterBaselines {
    pub(crate) fn insert(&mut self, field_name: &str, known_items: AHashSet<PivotValue>) {
        self.known_items_by_field
            .insert(field_name.to_lowercase(), known_items);
    }

    pub(crate) fn known_items(&self, field_name: &str) -> Option<&AHashSet<PivotValue>> {
        self.known_items_by_field.get(&field_name.to_lowercase())
    }
}

pub(crate) fn pivot_axis_field_includes_new_items(pivot: &PivotTable, field_name: &str) -> bool {
    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .any(|field| {
            field.field.name.eq_ignore_ascii_case(field_name) && field.include_new_items_in_filter
        })
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) enum SourceCacheKey {
    Single(SourceRangeCacheKey),
    Consolidation(Vec<SourceRangeCacheKey>),
}

impl SourceCacheKey {
    pub(crate) fn rebase_untouched(
        &mut self,
        workbook: &Workbook,
        touched_ranges: &[(usize, CellRange)],
    ) -> bool {
        match self {
            Self::Single(key) => key.rebase_untouched(workbook, touched_ranges),
            Self::Consolidation(keys) => keys
                .iter_mut()
                .all(|key| key.rebase_untouched(workbook, touched_ranges)),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct TransformedSnapshotCacheKey {
    source: SourceCacheKey,
    calculated_fields: Vec<PivotCalculatedField>,
    groupings: Vec<PivotGroupingCacheKey>,
    calculated_items: Vec<PivotCalculatedItem>,
    date_1904: bool,
}

impl TransformedSnapshotCacheKey {
    pub(crate) fn new(
        source: SourceCacheKey,
        calculated_fields: &[PivotCalculatedField],
        groupings: &[PivotGrouping],
        calculated_items: &[PivotCalculatedItem],
        date_1904: bool,
    ) -> Self {
        Self {
            source,
            calculated_fields: calculated_fields.to_vec(),
            groupings: groupings.iter().map(PivotGroupingCacheKey::from).collect(),
            calculated_items: calculated_items.to_vec(),
            date_1904,
        }
    }

    pub(crate) fn rebase_untouched(
        &mut self,
        workbook: &Workbook,
        touched_ranges: &[(usize, CellRange)],
    ) -> bool {
        self.source.rebase_untouched(workbook, touched_ranges)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) enum PivotGroupingCacheKey {
    Number {
        field: String,
        start: Option<u64>,
        end: Option<u64>,
        interval: u64,
    },
    Date {
        field: String,
        units: Vec<duke_sheets_core::PivotDateGroupUnit>,
    },
    Manual {
        field: String,
        groups: Vec<PivotManualGroupCacheKey>,
    },
}

impl From<&PivotGrouping> for PivotGroupingCacheKey {
    fn from(grouping: &PivotGrouping) -> Self {
        match grouping {
            PivotGrouping::Number {
                field,
                start,
                end,
                interval,
            } => Self::Number {
                field: field.name.clone(),
                start: start.map(f64_cache_key),
                end: end.map(f64_cache_key),
                interval: f64_cache_key(*interval),
            },
            PivotGrouping::Date { field, units } => Self::Date {
                field: field.name.clone(),
                units: units.clone(),
            },
            PivotGrouping::Manual { field, groups } => Self::Manual {
                field: field.name.clone(),
                groups: groups.iter().map(PivotManualGroupCacheKey::from).collect(),
            },
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct PivotManualGroupCacheKey {
    name: String,
    members: Vec<PivotValue>,
}

impl From<&PivotManualGroup> for PivotManualGroupCacheKey {
    fn from(group: &PivotManualGroup) -> Self {
        Self {
            name: group.name.clone(),
            members: group.members.clone(),
        }
    }
}

pub(crate) fn f64_cache_key(value: f64) -> u64 {
    value.to_bits()
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct SourceRangeCacheKey {
    kind: SourceCacheKind,
    sheet_index: usize,
    range: CellRange,
    source_name: Option<String>,
    mutation_count: u64,
    topology_generation: u64,
}

impl SourceRangeCacheKey {
    pub(crate) fn rebase_untouched(
        &mut self,
        workbook: &Workbook,
        touched_ranges: &[(usize, CellRange)],
    ) -> bool {
        let Some(worksheet) = workbook.worksheet(self.sheet_index) else {
            return false;
        };
        let source_touched = touched_ranges.iter().any(|(sheet_index, range)| {
            *sheet_index == self.sheet_index && range.overlaps(&self.range)
        });
        if !source_touched {
            self.mutation_count = worksheet.mutation_count();
            self.topology_generation = worksheet.topology_generation();
            true
        } else {
            worksheet.mutation_count() == self.mutation_count
                && worksheet.topology_generation() == self.topology_generation
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum SourceCacheKind {
    WorksheetRange,
    Table,
    ConsolidationRange,
}

#[derive(Debug, Clone)]
pub(crate) struct ResolvedSource {
    pub(crate) kind: SourceCacheKind,
    pub(crate) sheet_index: usize,
    pub(crate) range: CellRange,
    pub(crate) source_name: Option<String>,
    pub(crate) headers: Option<Vec<String>>,
    pub(crate) data_start_row: u32,
    pub(crate) data_end_row: Option<u32>,
    pub(crate) mutation_count: u64,
    pub(crate) topology_generation: u64,
}

impl ResolvedSource {
    pub(crate) fn cache_key(&self) -> SourceRangeCacheKey {
        SourceRangeCacheKey {
            kind: self.kind,
            sheet_index: self.sheet_index,
            range: self.range,
            source_name: self.source_name.clone(),
            mutation_count: self.mutation_count,
            topology_generation: self.topology_generation,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct CachedSourceSnapshot {
    pub(crate) key: SourceCacheKey,
    pub(crate) snapshot: Arc<SourceSnapshot>,
}
