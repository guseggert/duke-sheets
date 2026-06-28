use crate::compile::*;
use crate::filters::*;
use crate::prelude::*;
use crate::render::*;
use crate::show_as::*;
use crate::snapshot::*;
use crate::sort::*;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct GroupKey {
    pub(crate) rows: Vec<u32>,
    pub(crate) columns: Vec<u32>,
}

#[derive(Debug, Clone)]
pub(crate) struct PivotAggregation {
    pub(crate) groups: AHashMap<GroupKey, Vec<AggregateState>>,
    pub(crate) group_order: Vec<GroupKey>,
    pub(crate) row_totals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    pub(crate) row_subtotals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    pub(crate) row_order: Vec<Vec<u32>>,
    pub(crate) column_totals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    pub(crate) column_subtotals: AHashMap<Vec<u32>, Vec<AggregateState>>,
    pub(crate) subtotal_groups: AHashMap<GroupKey, Vec<AggregateState>>,
    pub(crate) column_order: Vec<Vec<u32>>,
    pub(crate) grand_totals: Vec<AggregateState>,
    pub(crate) matched_rows: usize,
    pub(crate) total_source: Option<Arc<PivotAggregation>>,
    pub(crate) subtotal_source: Option<Arc<PivotAggregation>>,
}

impl PivotAggregation {
    pub(crate) fn aggregate_visible(snapshot: &SourceSnapshot, plan: &CompiledPivotPlan) -> Self {
        #[cfg(feature = "parallel")]
        {
            if snapshot.row_count >= PARALLEL_ROW_THRESHOLD {
                return Self::aggregate_visible_parallel(snapshot, plan);
            }
        }

        Self::aggregate_visible_range(snapshot, plan, 0, snapshot.row_count)
    }

    pub(crate) fn aggregate_visible_range(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        start: usize,
        end: usize,
    ) -> Self {
        let mut aggregation = Self::empty(plan);

        for row in start..end {
            if !plan
                .filters
                .iter()
                .all(|filter| filter.matches(snapshot, row))
            {
                continue;
            }
            aggregation.ingest_row(snapshot, plan, row);
        }

        aggregation
    }

    pub(crate) fn aggregate_visible_with_totals_source(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
    ) -> (Self, Self) {
        #[cfg(feature = "parallel")]
        {
            if snapshot.row_count >= PARALLEL_ROW_THRESHOLD {
                return Self::aggregate_visible_with_totals_source_parallel(snapshot, plan);
            }
        }

        Self::aggregate_visible_with_totals_source_range(snapshot, plan, 0, snapshot.row_count)
    }

    pub(crate) fn aggregate_visible_with_totals_source_range(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        start: usize,
        end: usize,
    ) -> (Self, Self) {
        let mut visible = Self::empty(plan);
        let mut totals = Self::empty(plan);

        for row in start..end {
            if !plan
                .totals_filters
                .iter()
                .all(|filter| filter.matches(snapshot, row))
            {
                continue;
            }

            totals.ingest_row(snapshot, plan, row);

            if plan
                .axis_filters
                .iter()
                .all(|filter| filter.matches(snapshot, row))
            {
                visible.ingest_row(snapshot, plan, row);
            }
        }

        (visible, totals)
    }

    #[cfg(feature = "parallel")]
    pub(crate) fn aggregate_visible_parallel(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
    ) -> Self {
        let chunks = (0..snapshot.row_count)
            .step_by(PARALLEL_CHUNK_SIZE)
            .map(|start| (start, (start + PARALLEL_CHUNK_SIZE).min(snapshot.row_count)))
            .collect::<Vec<_>>();

        let partials = chunks
            .par_iter()
            .map(|(start, end)| Self::aggregate_visible_range(snapshot, plan, *start, *end))
            .collect::<Vec<_>>();

        let mut merged = Self::empty(plan);

        for partial in partials {
            merged.merge_from(partial, plan);
        }

        merged
    }

    #[cfg(feature = "parallel")]
    pub(crate) fn aggregate_visible_with_totals_source_parallel(
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
    ) -> (Self, Self) {
        let chunks = (0..snapshot.row_count)
            .step_by(PARALLEL_CHUNK_SIZE)
            .map(|start| (start, (start + PARALLEL_CHUNK_SIZE).min(snapshot.row_count)))
            .collect::<Vec<_>>();

        let partials = chunks
            .par_iter()
            .map(|(start, end)| {
                Self::aggregate_visible_with_totals_source_range(snapshot, plan, *start, *end)
            })
            .collect::<Vec<_>>();

        let mut visible = Self::empty(plan);
        let mut totals = Self::empty(plan);
        for (partial_visible, partial_totals) in partials {
            visible.merge_from(partial_visible, plan);
            totals.merge_from(partial_totals, plan);
        }
        (visible, totals)
    }

    pub(crate) fn empty(plan: &CompiledPivotPlan) -> Self {
        Self {
            groups: AHashMap::new(),
            group_order: Vec::new(),
            row_totals: AHashMap::new(),
            row_subtotals: AHashMap::new(),
            row_order: Vec::new(),
            column_totals: AHashMap::new(),
            column_subtotals: AHashMap::new(),
            subtotal_groups: AHashMap::new(),
            column_order: Vec::new(),
            grand_totals: default_states(&plan.measures),
            matched_rows: 0,
            total_source: None,
            subtotal_source: None,
        }
    }

    pub(crate) fn attach_hidden_total_source(
        &mut self,
        hidden_total_source: Self,
        use_for_totals: bool,
        use_for_subtotals: bool,
    ) {
        let hidden_total_source = Arc::new(hidden_total_source);
        if use_for_totals {
            self.total_source = Some(Arc::clone(&hidden_total_source));
        }
        if use_for_subtotals {
            self.subtotal_source = Some(hidden_total_source);
        }
    }

    pub(crate) fn ingest_row(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
    ) {
        self.matched_rows += 1;
        let row_key = encoded_key(snapshot, &plan.row_indexes, row);
        let column_key = encoded_key(snapshot, &plan.column_indexes, row);
        let group_key = GroupKey {
            rows: row_key.clone(),
            columns: column_key.clone(),
        };

        update_states(
            ordered_bucket_states_mut(
                &mut self.groups,
                &mut self.group_order,
                group_key,
                &plan.measures,
            ),
            snapshot,
            plan,
            row,
        );

        update_states(
            ordered_bucket_states_mut(
                &mut self.row_totals,
                &mut self.row_order,
                row_key.clone(),
                &plan.measures,
            ),
            snapshot,
            plan,
            row,
        );

        self.ingest_subtotals(snapshot, plan, row, &row_key, &column_key);

        update_states(
            ordered_bucket_states_mut(
                &mut self.column_totals,
                &mut self.column_order,
                column_key.clone(),
                &plan.measures,
            ),
            snapshot,
            plan,
            row,
        );

        update_states(&mut self.grand_totals, snapshot, plan, row);
    }

    pub(crate) fn ingest_subtotals(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        row_key: &[u32],
        column_key: &[u32],
    ) {
        let row_prefixes = row_subtotal_prefixes(plan, row_key);
        let column_prefixes = column_subtotal_prefixes(plan, column_key);

        for prefix in &row_prefixes {
            self.update_row_subtotal(snapshot, plan, row, prefix.clone());
            if !column_key.is_empty() {
                self.update_subtotal_group(
                    snapshot,
                    plan,
                    row,
                    prefix.clone(),
                    column_key.to_vec(),
                );
            }
        }

        for prefix in &column_prefixes {
            self.update_column_subtotal(snapshot, plan, row, prefix.clone());
            self.update_subtotal_group(snapshot, plan, row, row_key.to_vec(), prefix.clone());
        }

        for row_prefix in &row_prefixes {
            for column_prefix in &column_prefixes {
                self.update_subtotal_group(
                    snapshot,
                    plan,
                    row,
                    row_prefix.clone(),
                    column_prefix.clone(),
                );
            }
        }
    }

    pub(crate) fn update_row_subtotal(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        prefix: Vec<u32>,
    ) {
        let states = self
            .row_subtotals
            .entry(prefix)
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    pub(crate) fn update_column_subtotal(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        prefix: Vec<u32>,
    ) {
        let states = self
            .column_subtotals
            .entry(prefix)
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    pub(crate) fn update_subtotal_group(
        &mut self,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        row: usize,
        row_key: Vec<u32>,
        column_key: Vec<u32>,
    ) {
        let states = self
            .subtotal_groups
            .entry(GroupKey {
                rows: row_key,
                columns: column_key,
            })
            .or_insert_with(|| default_states(&plan.measures));
        update_states(states, snapshot, plan, row);
    }

    pub(crate) fn apply_aggregate_filters(
        &mut self,
        plan: &CompiledPivotPlan,
    ) -> AxisItemRestrictions {
        let mut restrictions = AxisItemRestrictions::default();
        for filter in &plan.aggregate_filters {
            let allowed_item_ids = filter.allowed_item_ids(self);
            restrictions.restrict(filter.axis(), filter.field_position(), &allowed_item_ids);
            self.retain_axis_items(
                filter.axis(),
                filter.field_position(),
                &allowed_item_ids,
                plan,
            );
        }
        restrictions
    }

    pub(crate) fn apply_calculated_items(
        &mut self,
        pivot_name: &str,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
    ) -> Result<()> {
        if plan.calculated_items.is_empty() {
            return Ok(());
        }

        for item in &plan.calculated_items {
            self.apply_calculated_item(pivot_name, snapshot, plan, item)?;
        }
        self.rebuild_totals_from_groups(plan);
        Ok(())
    }

    pub(crate) fn apply_calculated_item(
        &mut self,
        pivot_name: &str,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        item: &CompiledCalculatedItem,
    ) -> Result<()> {
        let virtual_keys = calculated_item_virtual_keys(item, &self.group_order);
        let evaluated =
            evaluate_calculated_item_groups(pivot_name, snapshot, plan, self, item, &virtual_keys)?;

        for (virtual_key, states) in evaluated {
            self.groups.insert(virtual_key.clone(), states);
            push_unique_group_key(&mut self.group_order, virtual_key.clone());
            match item.axis {
                AggregateFilterAxis::Row => push_unique_key(&mut self.row_order, virtual_key.rows),
                AggregateFilterAxis::Column => {
                    push_unique_key(&mut self.column_order, virtual_key.columns)
                }
            }
        }

        Ok(())
    }

    pub(crate) fn retain_axis_items(
        &mut self,
        axis: AggregateFilterAxis,
        field_position: usize,
        allowed_item_ids: &AHashSet<u32>,
        plan: &CompiledPivotPlan,
    ) {
        match axis {
            AggregateFilterAxis::Row => {
                self.row_order
                    .retain(|key| allowed_item_ids.contains(&key[field_position]));
                self.group_order
                    .retain(|key| allowed_item_ids.contains(&key.rows[field_position]));
                self.groups
                    .retain(|key, _| allowed_item_ids.contains(&key.rows[field_position]));
            }
            AggregateFilterAxis::Column => {
                self.column_order
                    .retain(|key| allowed_item_ids.contains(&key[field_position]));
                self.group_order
                    .retain(|key| allowed_item_ids.contains(&key.columns[field_position]));
                self.groups
                    .retain(|key, _| allowed_item_ids.contains(&key.columns[field_position]));
            }
        }
        self.rebuild_totals_from_groups(plan);
    }

    pub(crate) fn item_states_for_filter(
        &self,
        axis: AggregateFilterAxis,
        field_position: usize,
        measure_index: usize,
        aggregate: PivotAggregate,
    ) -> AHashMap<u32, AggregateState> {
        let mut item_states = AHashMap::new();
        let (order, totals) = match axis {
            AggregateFilterAxis::Row => (&self.row_order, &self.row_totals),
            AggregateFilterAxis::Column => (&self.column_order, &self.column_totals),
        };

        for key in order {
            let Some(states) = totals.get(key) else {
                continue;
            };
            let Some(state) = states.get(measure_index) else {
                continue;
            };
            item_states
                .entry(key[field_position])
                .or_insert_with(|| AggregateState::new(aggregate))
                .merge(state, aggregate);
        }
        item_states
    }

    pub(crate) fn rebuild_totals_from_groups(&mut self, plan: &CompiledPivotPlan) {
        self.row_totals.clear();
        self.row_subtotals.clear();
        self.column_totals.clear();
        self.column_subtotals.clear();
        self.subtotal_groups.clear();
        self.grand_totals = default_states(&plan.measures);

        for key in &self.group_order {
            let Some(states) = self.groups.get(key) else {
                continue;
            };

            let row_states = self
                .row_totals
                .entry(key.rows.clone())
                .or_insert_with(|| default_states(&plan.measures));
            merge_state_slices(row_states, states, &plan.measures);

            let column_states = self
                .column_totals
                .entry(key.columns.clone())
                .or_insert_with(|| default_states(&plan.measures));
            merge_state_slices(column_states, states, &plan.measures);

            let row_prefixes = row_subtotal_prefixes(plan, &key.rows);
            let column_prefixes = column_subtotal_prefixes(plan, &key.columns);

            for prefix in &row_prefixes {
                let row_subtotal = self
                    .row_subtotals
                    .entry(prefix.clone())
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(row_subtotal, states, &plan.measures);

                if !key.columns.is_empty() {
                    let subtotal_group = self
                        .subtotal_groups
                        .entry(GroupKey {
                            rows: prefix.clone(),
                            columns: key.columns.clone(),
                        })
                        .or_insert_with(|| default_states(&plan.measures));
                    merge_state_slices(subtotal_group, states, &plan.measures);
                }
            }

            for prefix in &column_prefixes {
                let column_subtotal = self
                    .column_subtotals
                    .entry(prefix.clone())
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(column_subtotal, states, &plan.measures);

                let subtotal_group = self
                    .subtotal_groups
                    .entry(GroupKey {
                        rows: key.rows.clone(),
                        columns: prefix.clone(),
                    })
                    .or_insert_with(|| default_states(&plan.measures));
                merge_state_slices(subtotal_group, states, &plan.measures);
            }

            for row_prefix in &row_prefixes {
                for column_prefix in &column_prefixes {
                    let subtotal_group = self
                        .subtotal_groups
                        .entry(GroupKey {
                            rows: row_prefix.clone(),
                            columns: column_prefix.clone(),
                        })
                        .or_insert_with(|| default_states(&plan.measures));
                    merge_state_slices(subtotal_group, states, &plan.measures);
                }
            }

            merge_state_slices(&mut self.grand_totals, states, &plan.measures);
        }

        self.row_order
            .retain(|key| self.row_totals.contains_key(key));
        self.column_order
            .retain(|key| self.column_totals.contains_key(key));
    }

    pub(crate) fn expand_show_empty_items(
        &mut self,
        pivot: &PivotTable,
        snapshot: &SourceSnapshot,
        plan: &CompiledPivotPlan,
        aggregate_restrictions: &AxisItemRestrictions,
    ) -> Result<()> {
        expand_axis_show_empty_items(
            &pivot.name,
            snapshot,
            plan,
            AggregateFilterAxis::Row,
            pivot.layout.show_empty_rows,
            &plan.row_indexes,
            &plan.row_fields,
            &plan.filters,
            aggregate_restrictions,
            &mut self.row_order,
        )?;
        expand_axis_show_empty_items(
            &pivot.name,
            snapshot,
            plan,
            AggregateFilterAxis::Column,
            pivot.layout.show_empty_columns,
            &plan.column_indexes,
            &plan.column_fields,
            &plan.filters,
            aggregate_restrictions,
            &mut self.column_order,
        )?;
        Ok(())
    }

    pub(crate) fn sort_orders(&mut self, snapshot: &SourceSnapshot, plan: &CompiledPivotPlan) {
        sort_key_order(
            &mut self.row_order,
            &plan.row_indexes,
            &plan.row_fields,
            &plan.row_sort_measure_indexes,
            &self.row_totals,
            &plan.measures,
            snapshot,
        );
        sort_key_order(
            &mut self.column_order,
            &plan.column_indexes,
            &plan.column_fields,
            &plan.column_sort_measure_indexes,
            &self.column_totals,
            &plan.measures,
            snapshot,
        );
        let row_positions = order_positions(&self.row_order);
        let column_positions = order_positions(&self.column_order);
        self.group_order.sort_by(|a, b| {
            row_positions
                .get(&a.rows)
                .cmp(&row_positions.get(&b.rows))
                .then_with(|| {
                    column_positions
                        .get(&a.columns)
                        .cmp(&column_positions.get(&b.columns))
                })
        });
    }

    #[cfg(feature = "parallel")]
    pub(crate) fn merge_from(&mut self, other: Self, plan: &CompiledPivotPlan) {
        self.matched_rows += other.matched_rows;
        merge_state_slices(&mut self.grand_totals, &other.grand_totals, &plan.measures);

        for key in other.group_order {
            let states = other
                .groups
                .get(&key)
                .expect("ordered group key must exist")
                .clone();
            merge_ordered_bucket(
                &mut self.groups,
                &mut self.group_order,
                key,
                states,
                &plan.measures,
            );
        }

        for key in other.row_order {
            let states = other
                .row_totals
                .get(&key)
                .expect("ordered row key must exist")
                .clone();
            merge_ordered_bucket(
                &mut self.row_totals,
                &mut self.row_order,
                key,
                states,
                &plan.measures,
            );
        }

        for (key, states) in other.row_subtotals {
            merge_unordered_bucket(&mut self.row_subtotals, key, states, &plan.measures);
        }

        for key in other.column_order {
            let states = other
                .column_totals
                .get(&key)
                .expect("ordered column key must exist")
                .clone();
            merge_ordered_bucket(
                &mut self.column_totals,
                &mut self.column_order,
                key,
                states,
                &plan.measures,
            );
        }

        for (key, states) in other.column_subtotals {
            merge_unordered_bucket(&mut self.column_subtotals, key, states, &plan.measures);
        }

        for (key, states) in other.subtotal_groups {
            merge_unordered_bucket(&mut self.subtotal_groups, key, states, &plan.measures);
        }
    }
}

#[cfg(feature = "parallel")]
pub(crate) fn merge_ordered_bucket<K>(
    map: &mut AHashMap<K, Vec<AggregateState>>,
    order: &mut Vec<K>,
    key: K,
    states: Vec<AggregateState>,
    measures: &[PivotMeasure],
) where
    K: Eq + Hash + Clone,
{
    if let Some(existing) = map.get_mut(&key) {
        merge_state_slices(existing, &states, measures);
    } else {
        order.push(key.clone());
        map.insert(key, states);
    }
}

#[cfg(feature = "parallel")]
pub(crate) fn merge_unordered_bucket<K>(
    map: &mut AHashMap<K, Vec<AggregateState>>,
    key: K,
    states: Vec<AggregateState>,
    measures: &[PivotMeasure],
) where
    K: Eq + Hash,
{
    if let Some(existing) = map.get_mut(&key) {
        merge_state_slices(existing, &states, measures);
    } else {
        map.insert(key, states);
    }
}

pub(crate) fn merge_state_slices(
    target: &mut [AggregateState],
    source: &[AggregateState],
    measures: &[PivotMeasure],
) {
    for ((target, source), measure) in target.iter_mut().zip(source.iter()).zip(measures.iter()) {
        target.merge(source, measure.aggregate);
    }
}

pub(crate) fn calculated_item_source_key_matches(
    item: &CompiledCalculatedItem,
    key: &GroupKey,
) -> bool {
    match item.axis {
        AggregateFilterAxis::Row => key
            .rows
            .get(item.position)
            .is_some_and(|id| *id != item.item_id),
        AggregateFilterAxis::Column => key
            .columns
            .get(item.position)
            .is_some_and(|id| *id != item.item_id),
    }
}

pub(crate) fn calculated_item_virtual_key(
    item: &CompiledCalculatedItem,
    source_key: &GroupKey,
) -> GroupKey {
    let mut key = source_key.clone();
    match item.axis {
        AggregateFilterAxis::Row => key.rows[item.position] = item.item_id,
        AggregateFilterAxis::Column => key.columns[item.position] = item.item_id,
    }
    key
}

pub(crate) fn calculated_item_virtual_keys(
    item: &CompiledCalculatedItem,
    source_group_order: &[GroupKey],
) -> Vec<GroupKey> {
    let mut emitted = AHashSet::new();
    let mut virtual_keys = Vec::new();

    for source_key in source_group_order {
        if !calculated_item_source_key_matches(item, source_key) {
            continue;
        }

        let virtual_key = calculated_item_virtual_key(item, source_key);
        if emitted.insert(virtual_key.clone()) {
            virtual_keys.push(virtual_key);
        }
    }

    virtual_keys
}

pub(crate) fn push_unique_key(order: &mut Vec<Vec<u32>>, key: Vec<u32>) {
    if !order.iter().any(|existing| existing == &key) {
        order.push(key);
    }
}

pub(crate) fn push_unique_group_key(order: &mut Vec<GroupKey>, key: GroupKey) {
    if !order.iter().any(|existing| existing == &key) {
        order.push(key);
    }
}

pub(crate) fn evaluate_calculated_item_groups(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    item: &CompiledCalculatedItem,
    group_keys: &[GroupKey],
) -> Result<Vec<(GroupKey, Vec<AggregateState>)>> {
    #[cfg(feature = "parallel")]
    {
        if group_keys.len() >= PARALLEL_CALCULATED_ITEM_GROUP_THRESHOLD {
            return group_keys
                .par_iter()
                .map(|group_key| {
                    evaluate_calculated_group_states(
                        pivot_name,
                        snapshot,
                        plan,
                        aggregation,
                        item,
                        group_key,
                    )
                    .map(|states| (group_key.clone(), states))
                })
                .collect();
        }
    }

    group_keys
        .iter()
        .map(|group_key| {
            evaluate_calculated_group_states(
                pivot_name,
                snapshot,
                plan,
                aggregation,
                item,
                group_key,
            )
            .map(|states| (group_key.clone(), states))
        })
        .collect()
}

pub(crate) fn evaluate_calculated_group_states(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    aggregation: &PivotAggregation,
    item: &CompiledCalculatedItem,
    group_key: &GroupKey,
) -> Result<Vec<AggregateState>> {
    plan.measures
        .iter()
        .enumerate()
        .map(|(measure_index, measure)| {
            let context = CalculatedItemEvalContext {
                snapshot,
                aggregation,
                group_key,
                measure_index,
                aggregate: measure.aggregate,
            };
            let materialized =
                materialize_calculated_item_expr(pivot_name, item, &item.ast, &context)?;
            let value = evaluate(&materialized, &EvaluationContext::simple()).map_err(|error| {
                Error::other(format!(
                    "pivot table {pivot_name} calculated item {} evaluation failed: {error}",
                    item.item
                ))
            })?;
            calculated_item_state_from_value(pivot_name, item, value)
        })
        .collect()
}

pub(crate) struct CalculatedItemEvalContext<'a> {
    pub(crate) snapshot: &'a SourceSnapshot,
    pub(crate) aggregation: &'a PivotAggregation,
    pub(crate) group_key: &'a GroupKey,
    pub(crate) measure_index: usize,
    pub(crate) aggregate: PivotAggregate,
}

pub(crate) fn materialize_calculated_item_expr(
    pivot_name: &str,
    item: &CompiledCalculatedItem,
    expr: &FormulaExpr,
    context: &CalculatedItemEvalContext<'_>,
) -> Result<FormulaExpr> {
    Ok(match expr {
        FormulaExpr::Number(value) => FormulaExpr::Number(*value),
        FormulaExpr::String(value) => {
            if formula_reference_item_id(context.snapshot, item.field_index, value).is_some() {
                calculated_item_reference_expr(pivot_name, item, value, context)?
            } else {
                FormulaExpr::String(value.clone())
            }
        }
        FormulaExpr::Boolean(value) => FormulaExpr::Boolean(*value),
        FormulaExpr::Error(value) => FormulaExpr::Error(*value),
        FormulaExpr::Empty => FormulaExpr::Empty,
        FormulaExpr::NameRef(name) => {
            calculated_item_reference_expr(pivot_name, item, name, context)?
        }
        FormulaExpr::StructuredRef(_) => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated item {} uses a structured reference, which is not valid for item formulas",
                item.item
            )));
        }
        FormulaExpr::CellRef(reference) => {
            if let Some(name) = calculated_item_cell_reference_name(reference) {
                if formula_reference_item_id(context.snapshot, item.field_index, &name).is_some() {
                    calculated_item_reference_expr(pivot_name, item, &name, context)?
                } else {
                    return Err(Error::other(format!(
                        "pivot table {pivot_name} calculated item {} uses workbook references, which are not valid pivot item references",
                        item.item
                    )));
                }
            } else {
                return Err(Error::other(format!(
                    "pivot table {pivot_name} calculated item {} uses workbook references, which are not valid pivot item references",
                    item.item
                )));
            }
        }
        FormulaExpr::RangeRef(_) | FormulaExpr::ExternalRef(_) => {
            return Err(Error::other(format!(
                "pivot table {pivot_name} calculated item {} uses workbook references, which are not valid pivot item references",
                item.item
            )));
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op: *op,
            left: Box::new(materialize_calculated_item_expr(
                pivot_name, item, left, context,
            )?),
            right: Box::new(materialize_calculated_item_expr(
                pivot_name, item, right, context,
            )?),
        },
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp {
            op: *op,
            operand: Box::new(materialize_calculated_item_expr(
                pivot_name, item, operand, context,
            )?),
        },
        FormulaExpr::Function { name, args } => FormulaExpr::Function {
            name: name.clone(),
            args: materialize_calculated_item_args(pivot_name, item, args, context)?,
        },
        FormulaExpr::ExternalFunction { book, name, args } => FormulaExpr::ExternalFunction {
            book: book.clone(),
            name: name.clone(),
            args: materialize_calculated_item_args(pivot_name, item, args, context)?,
        },
        FormulaExpr::Array(rows) => {
            let mut materialized_rows = Vec::with_capacity(rows.len());
            for row in rows {
                materialized_rows.push(materialize_calculated_item_args(
                    pivot_name, item, row, context,
                )?);
            }
            FormulaExpr::Array(materialized_rows)
        }
    })
}

pub(crate) fn materialize_calculated_item_args(
    pivot_name: &str,
    item: &CompiledCalculatedItem,
    args: &[FormulaExpr],
    context: &CalculatedItemEvalContext<'_>,
) -> Result<Vec<FormulaExpr>> {
    args.iter()
        .map(|arg| materialize_calculated_item_expr(pivot_name, item, arg, context))
        .collect()
}

pub(crate) fn calculated_item_reference_expr(
    pivot_name: &str,
    item: &CompiledCalculatedItem,
    reference: &str,
    context: &CalculatedItemEvalContext<'_>,
) -> Result<FormulaExpr> {
    let reference_id = formula_reference_item_id(context.snapshot, item.field_index, reference)
        .ok_or_else(|| {
            Error::other(format!(
                "pivot table {pivot_name} calculated item {} references unknown item: {reference}",
                item.item
            ))
        })?;

    let mut reference_key = context.group_key.clone();
    match item.axis {
        AggregateFilterAxis::Row => reference_key.rows[item.position] = reference_id,
        AggregateFilterAxis::Column => reference_key.columns[item.position] = reference_id,
    }

    let value = context
        .aggregation
        .groups
        .get(&reference_key)
        .and_then(|states| state_number(states, context.measure_index, context.aggregate));
    Ok(value.map(FormulaExpr::Number).unwrap_or(FormulaExpr::Empty))
}

pub(crate) fn calculated_item_state_from_value(
    pivot_name: &str,
    item: &CompiledCalculatedItem,
    value: FormulaValue,
) -> Result<AggregateState> {
    match value {
        FormulaValue::Number(value) => Ok(AggregateState::from_calculated_number(value)),
        FormulaValue::Boolean(value) => Ok(AggregateState::from_calculated_number(if value {
            1.0
        } else {
            0.0
        })),
        FormulaValue::Empty => Ok(AggregateState::new(PivotAggregate::Sum)),
        FormulaValue::Error(error) => Err(Error::other(format!(
            "pivot table {pivot_name} calculated item {} evaluated to {error}",
            item.item
        ))),
        FormulaValue::String(value) => Err(Error::other(format!(
            "pivot table {pivot_name} calculated item {} evaluated to non-numeric value {value}",
            item.item
        ))),
        FormulaValue::Array { .. } => Err(Error::other(format!(
            "pivot table {pivot_name} calculated item {} evaluated to an array",
            item.item
        ))),
    }
}

pub(crate) fn expand_axis_show_empty_items(
    pivot_name: &str,
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    axis: AggregateFilterAxis,
    show_empty_axis: bool,
    field_indexes: &[usize],
    fields: &[PivotField],
    filters: &[CompiledFilter],
    aggregate_restrictions: &AxisItemRestrictions,
    order: &mut Vec<Vec<u32>>,
) -> Result<()> {
    if field_indexes.is_empty()
        || (!show_empty_axis && fields.iter().all(|field| !field.show_empty_items))
    {
        return Ok(());
    }

    let item_ids = axis_item_ids(
        snapshot,
        axis,
        show_empty_axis,
        field_indexes,
        fields,
        filters,
        aggregate_restrictions,
        order,
    );
    if item_ids.iter().any(Vec::is_empty) {
        return Ok(());
    }

    let key_count = cartesian_len(&item_ids)?;
    let limit = show_empty_axis_key_limit(axis, plan);
    if key_count > limit {
        return Err(Error::other(format!(
            "pivot table {pivot_name} show-empty-items expansion would produce {key_count} {} keys, exceeding the worksheet limit {limit}",
            axis_name(axis)
        )));
    }

    let mut seen = order.iter().cloned().collect::<AHashSet<_>>();
    let mut key = Vec::with_capacity(field_indexes.len());
    append_show_empty_axis_keys(&item_ids, 0, &mut key, &mut seen, order);
    Ok(())
}

pub(crate) fn axis_item_ids(
    snapshot: &SourceSnapshot,
    axis: AggregateFilterAxis,
    show_empty_axis: bool,
    field_indexes: &[usize],
    fields: &[PivotField],
    filters: &[CompiledFilter],
    aggregate_restrictions: &AxisItemRestrictions,
    order: &[Vec<u32>],
) -> Vec<Vec<u32>> {
    field_indexes
        .iter()
        .enumerate()
        .map(|(position, field_index)| {
            let mut ids = if show_empty_axis
                || fields
                    .get(position)
                    .map(|field| field.show_empty_items)
                    .unwrap_or(false)
            {
                visible_dictionary_item_ids(snapshot, *field_index, filters)
            } else {
                observed_axis_item_ids(order, position)
            };
            ids.retain(|id| aggregate_restrictions.allows(axis, position, *id));
            ids
        })
        .collect()
}

pub(crate) fn visible_dictionary_item_ids(
    snapshot: &SourceSnapshot,
    field_index: usize,
    filters: &[CompiledFilter],
) -> Vec<u32> {
    (0..snapshot.columns[field_index].dictionary.len())
        .map(|id| id as u32)
        .filter(|id| {
            filters
                .iter()
                .all(|filter| filter.allows_item(snapshot, field_index, *id))
        })
        .collect()
}

pub(crate) fn observed_axis_item_ids(order: &[Vec<u32>], position: usize) -> Vec<u32> {
    let mut seen = AHashSet::new();
    let mut ids = Vec::new();
    for key in order {
        let Some(id) = key.get(position) else {
            continue;
        };
        if seen.insert(*id) {
            ids.push(*id);
        }
    }
    ids
}

pub(crate) fn cartesian_len(item_ids: &[Vec<u32>]) -> Result<usize> {
    item_ids.iter().try_fold(1usize, |total, ids| {
        total
            .checked_mul(ids.len())
            .ok_or_else(|| Error::other("pivot show-empty-items expansion is too large"))
    })
}

pub(crate) fn show_empty_axis_key_limit(
    axis: AggregateFilterAxis,
    plan: &CompiledPivotPlan,
) -> usize {
    match axis {
        AggregateFilterAxis::Row => MAX_ROWS as usize,
        AggregateFilterAxis::Column => {
            let available_columns = (MAX_COLS as usize).saturating_sub(plan.row_indexes.len());
            available_columns / plan.measures.len().max(1)
        }
    }
}

pub(crate) fn append_show_empty_axis_keys(
    item_ids: &[Vec<u32>],
    position: usize,
    key: &mut Vec<u32>,
    seen: &mut AHashSet<Vec<u32>>,
    order: &mut Vec<Vec<u32>>,
) {
    if position == item_ids.len() {
        if seen.insert(key.clone()) {
            order.push(key.clone());
        }
        return;
    }

    for id in &item_ids[position] {
        key.push(*id);
        append_show_empty_axis_keys(item_ids, position + 1, key, seen, order);
        key.pop();
    }
}

pub(crate) fn axis_name(axis: AggregateFilterAxis) -> &'static str {
    match axis {
        AggregateFilterAxis::Row => "row",
        AggregateFilterAxis::Column => "column",
    }
}

pub(crate) fn encoded_key(
    snapshot: &SourceSnapshot,
    field_indexes: &[usize],
    row: usize,
) -> Vec<u32> {
    field_indexes
        .iter()
        .map(|field_index| snapshot.columns[*field_index].id_at(row))
        .collect()
}

pub(crate) fn default_states(measures: &[PivotMeasure]) -> Vec<AggregateState> {
    measures
        .iter()
        .map(|measure| AggregateState::new(measure.aggregate))
        .collect()
}

pub(crate) fn ordered_bucket_states_mut<'a, K>(
    map: &'a mut AHashMap<K, Vec<AggregateState>>,
    order: &mut Vec<K>,
    key: K,
    measures: &[PivotMeasure],
) -> &'a mut [AggregateState]
where
    K: Eq + Hash + Clone,
{
    match map.entry(key) {
        std::collections::hash_map::Entry::Occupied(entry) => entry.into_mut(),
        std::collections::hash_map::Entry::Vacant(entry) => {
            order.push(entry.key().clone());
            entry.insert(default_states(measures))
        }
    }
}

pub(crate) fn update_states(
    states: &mut [AggregateState],
    snapshot: &SourceSnapshot,
    plan: &CompiledPivotPlan,
    row: usize,
) {
    for ((state, field_index), measure) in states
        .iter_mut()
        .zip(plan.measure_indexes.iter())
        .zip(plan.measures.iter())
    {
        state.update(snapshot.value(row, *field_index), measure.aggregate);
    }
}

#[derive(Debug, Clone)]
pub(crate) struct AggregateState {
    pub(crate) count_non_blank: u64,
    pub(crate) count_numbers: u64,
    pub(crate) sum: f64,
    pub(crate) sum_sq: f64,
    pub(crate) product: f64,
    pub(crate) min: Option<f64>,
    pub(crate) max: Option<f64>,
    pub(crate) calculated_value: Option<f64>,
}

impl AggregateState {
    pub(crate) fn new(_aggregate: PivotAggregate) -> Self {
        Self {
            count_non_blank: 0,
            count_numbers: 0,
            sum: 0.0,
            sum_sq: 0.0,
            product: 1.0,
            min: None,
            max: None,
            calculated_value: None,
        }
    }

    pub(crate) fn from_calculated_number(value: f64) -> Self {
        Self {
            count_non_blank: 1,
            count_numbers: 1,
            sum: value,
            sum_sq: value * value,
            product: value,
            min: Some(value),
            max: Some(value),
            calculated_value: Some(value),
        }
    }

    pub(crate) fn from_display_stats(
        stats: &DisplayAggregateStats,
        aggregate: PivotAggregate,
    ) -> Self {
        Self {
            count_non_blank: stats.count,
            count_numbers: stats.count,
            sum: stats.sum,
            sum_sq: stats.sum_sq,
            product: stats.product,
            min: stats.min,
            max: stats.max,
            calculated_value: stats.finalize(aggregate),
        }
    }

    pub(crate) fn update(&mut self, value: &PivotValue, _aggregate: PivotAggregate) {
        self.calculated_value = None;
        if !value.is_blank() {
            self.count_non_blank += 1;
        }

        let Some(number) = pivot_number(value) else {
            return;
        };

        self.count_numbers += 1;
        self.sum += number;
        self.sum_sq += number * number;
        self.product *= number;
        self.min = Some(self.min.map_or(number, |current| current.min(number)));
        self.max = Some(self.max.map_or(number, |current| current.max(number)));
    }

    pub(crate) fn finalize(&self, aggregate: PivotAggregate) -> CellValue {
        self.finalize_number(aggregate)
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty)
    }

    pub(crate) fn finalize_number(&self, aggregate: PivotAggregate) -> Option<f64> {
        if let Some(value) = self.calculated_value {
            return Some(value);
        }

        match aggregate {
            PivotAggregate::Sum => Some(self.sum),
            PivotAggregate::Count => Some(self.count_non_blank as f64),
            PivotAggregate::CountNumbers => Some(self.count_numbers as f64),
            PivotAggregate::Average => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(self.sum / self.count_numbers as f64)
                }
            }
            PivotAggregate::Max => self.max,
            PivotAggregate::Min => self.min,
            PivotAggregate::Product => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(self.product)
                }
            }
            PivotAggregate::StdDev => {
                if self.count_numbers < 2 {
                    None
                } else {
                    Some(sample_variance(self).sqrt())
                }
            }
            PivotAggregate::StdDevP => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(population_variance(self).sqrt())
                }
            }
            PivotAggregate::Var => {
                if self.count_numbers < 2 {
                    None
                } else {
                    Some(sample_variance(self))
                }
            }
            PivotAggregate::VarP => {
                if self.count_numbers == 0 {
                    None
                } else {
                    Some(population_variance(self))
                }
            }
        }
    }

    pub(crate) fn merge(&mut self, other: &Self, aggregate: PivotAggregate) {
        if self.calculated_value.is_some() || other.calculated_value.is_some() {
            self.merge_display_value(other, aggregate);
            return;
        }

        self.count_non_blank += other.count_non_blank;
        self.count_numbers += other.count_numbers;
        self.sum += other.sum;
        self.sum_sq += other.sum_sq;
        self.product *= other.product;
        self.min = match (self.min, other.min) {
            (Some(left), Some(right)) => Some(left.min(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
        self.max = match (self.max, other.max) {
            (Some(left), Some(right)) => Some(left.max(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
    }

    pub(crate) fn merge_display_value(&mut self, other: &Self, aggregate: PivotAggregate) {
        let Some(mut stats) = DisplayAggregateStats::from_state(self, aggregate) else {
            if let Some(other_stats) = DisplayAggregateStats::from_state(other, aggregate) {
                *self = AggregateState::from_display_stats(&other_stats, aggregate);
            }
            return;
        };

        if let Some(other_stats) = DisplayAggregateStats::from_state(other, aggregate) {
            stats.merge(&other_stats);
            *self = AggregateState::from_display_stats(&stats, aggregate);
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct DisplayAggregateStats {
    pub(crate) count: u64,
    pub(crate) sum: f64,
    pub(crate) sum_sq: f64,
    pub(crate) product: f64,
    pub(crate) min: Option<f64>,
    pub(crate) max: Option<f64>,
}

impl DisplayAggregateStats {
    pub(crate) fn single(value: f64) -> Self {
        Self {
            count: 1,
            sum: value,
            sum_sq: value * value,
            product: value,
            min: Some(value),
            max: Some(value),
        }
    }

    pub(crate) fn from_state(state: &AggregateState, aggregate: PivotAggregate) -> Option<Self> {
        if state.calculated_value.is_some() {
            return Some(Self {
                count: state.count_numbers,
                sum: state.sum,
                sum_sq: state.sum_sq,
                product: state.product,
                min: state.min,
                max: state.max,
            });
        }

        state.finalize_number(aggregate).map(Self::single)
    }

    pub(crate) fn merge(&mut self, other: &Self) {
        self.count += other.count;
        self.sum += other.sum;
        self.sum_sq += other.sum_sq;
        self.product *= other.product;
        self.min = match (self.min, other.min) {
            (Some(left), Some(right)) => Some(left.min(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
        self.max = match (self.max, other.max) {
            (Some(left), Some(right)) => Some(left.max(right)),
            (Some(left), None) => Some(left),
            (None, Some(right)) => Some(right),
            (None, None) => None,
        };
    }

    pub(crate) fn finalize(&self, aggregate: PivotAggregate) -> Option<f64> {
        match aggregate {
            PivotAggregate::Sum | PivotAggregate::Count | PivotAggregate::CountNumbers => {
                Some(self.sum)
            }
            PivotAggregate::Average => {
                if self.count == 0 {
                    None
                } else {
                    Some(self.sum / self.count as f64)
                }
            }
            PivotAggregate::Max => self.max,
            PivotAggregate::Min => self.min,
            PivotAggregate::Product => {
                if self.count == 0 {
                    None
                } else {
                    Some(self.product)
                }
            }
            PivotAggregate::StdDev => {
                if self.count < 2 {
                    None
                } else {
                    Some(sample_variance_from_parts(self.sum, self.sum_sq, self.count).sqrt())
                }
            }
            PivotAggregate::StdDevP => {
                if self.count == 0 {
                    None
                } else {
                    Some(population_variance_from_parts(self.sum, self.sum_sq, self.count).sqrt())
                }
            }
            PivotAggregate::Var => {
                if self.count < 2 {
                    None
                } else {
                    Some(sample_variance_from_parts(
                        self.sum,
                        self.sum_sq,
                        self.count,
                    ))
                }
            }
            PivotAggregate::VarP => {
                if self.count == 0 {
                    None
                } else {
                    Some(population_variance_from_parts(
                        self.sum,
                        self.sum_sq,
                        self.count,
                    ))
                }
            }
        }
    }
}

pub(crate) fn pivot_number(value: &PivotValue) -> Option<f64> {
    match value {
        PivotValue::Number(value) => Some(*value),
        _ => None,
    }
}

pub(crate) fn population_variance(state: &AggregateState) -> f64 {
    population_variance_from_parts(state.sum, state.sum_sq, state.count_numbers)
}

pub(crate) fn sample_variance(state: &AggregateState) -> f64 {
    sample_variance_from_parts(state.sum, state.sum_sq, state.count_numbers)
}

pub(crate) fn population_variance_from_parts(sum: f64, sum_sq: f64, count: u64) -> f64 {
    let count = count as f64;
    ((sum_sq - (sum * sum / count)) / count).max(0.0)
}

pub(crate) fn sample_variance_from_parts(sum: f64, sum_sq: f64, count: u64) -> f64 {
    let count = count as f64;
    ((sum_sq - (sum * sum / count)) / (count - 1.0)).max(0.0)
}
