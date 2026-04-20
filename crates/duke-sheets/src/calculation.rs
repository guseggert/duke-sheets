//! Workbook calculation engine
//!
//! Provides workbook-level formula calculation with dependency tracking,
//! circular reference detection, and support for volatile functions.
//!
//! # Example
//!
//! ```rust,ignore
//! use duke_sheets::prelude::*;
//! use duke_sheets::calculation::WorkbookCalculationExt;
//!
//! let mut workbook = Workbook::new();
//! let sheet = workbook.worksheet_mut(0).unwrap();
//! sheet.set_cell_value("A1", 10.0).unwrap();
//! sheet.set_cell_value("A2", 20.0).unwrap();
//! sheet.set_cell_formula("A3", "=A1+A2").unwrap();
//!
//! // Calculate all formulas
//! let stats = workbook.calculate().unwrap();
//! println!("Calculated {} cells", stats.cells_calculated);
//! ```

use crate::{
    evaluate, parse_formula, CellValue, EvaluationContext, FormulaExpr, FormulaValue, ImageInfo,
    Result, Workbook,
};
use ahash::{AHashMap, AHashSet};
use dashmap::DashMap;
use duke_sheets_core::CellAddress;
use duke_sheets_formula::dependency::CellKey;
use duke_sheets_formula::functions::FunctionRegistry;
use duke_sheets_formula::{StructuredRefSpecifier, StructuredReference};
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::sync::{Arc, OnceLock};
use std::time::Instant;

#[cfg(feature = "parallel")]
use rayon::prelude::*;

/// Global function registry for volatile function lookup
static FUNCTION_REGISTRY: OnceLock<FunctionRegistry> = OnceLock::new();

fn get_function_registry() -> &'static FunctionRegistry {
    FUNCTION_REGISTRY.get_or_init(FunctionRegistry::new)
}

/// Options for workbook calculation
#[derive(Clone)]
#[allow(clippy::type_complexity)]
pub struct CalculationOptions {
    /// Enable iterative calculation for circular references
    pub iterative: bool,
    /// Maximum iterations for circular references (default: 100)
    pub max_iterations: u32,
    /// Maximum change threshold for convergence (default: 0.001)
    pub max_change: f64,
    /// Force recalculation of all cells, even if not dirty
    pub force_full_calculation: bool,
    /// Include volatile functions in calculation (NOW, TODAY, RAND, etc.)
    pub calculate_volatile: bool,
    /// Only calculate these sheets (and their transitive cross-sheet dependencies).
    /// If empty, calculate all sheets (default).
    pub sheets: Vec<usize>,
    /// Maximum number of threads for parallel evaluation.
    ///
    /// - `None` (default): use all available cores
    /// - `Some(1)`: force serial evaluation even when the `parallel` feature is enabled
    /// - `Some(n)`: use at most `n` threads
    ///
    /// This option has no effect when the `parallel` feature is not enabled
    /// (e.g. WASM builds).
    pub max_threads: Option<usize>,
    /// Optional callback for WEBSERVICE(url).
    ///
    /// Returning `None` produces `#N/A`.
    pub web_service_fn: Option<Arc<dyn Fn(&str) -> Option<String> + Send + Sync>>,
    /// Optional callback for RTD(prog_id, server, topics...).
    ///
    /// Returning `None` produces `#N/A`.
    pub rtd_fn: Option<Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>>,
}

impl fmt::Debug for CalculationOptions {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("CalculationOptions")
            .field("iterative", &self.iterative)
            .field("max_iterations", &self.max_iterations)
            .field("max_change", &self.max_change)
            .field("force_full_calculation", &self.force_full_calculation)
            .field("calculate_volatile", &self.calculate_volatile)
            .field("sheets", &self.sheets)
            .field("max_threads", &self.max_threads)
            .field(
                "web_service_fn",
                &self.web_service_fn.as_ref().map(|_| "<callback>"),
            )
            .field("rtd_fn", &self.rtd_fn.as_ref().map(|_| "<callback>"))
            .finish()
    }
}

impl Default for CalculationOptions {
    fn default() -> Self {
        Self {
            iterative: false,
            max_iterations: 100,
            max_change: 0.001,
            force_full_calculation: true,
            calculate_volatile: true,
            sheets: vec![],
            max_threads: None,
            web_service_fn: None,
            rtd_fn: None,
        }
    }
}

/// Statistics from a calculation run
#[derive(Debug, Clone, Default)]
pub struct CalculationStats {
    /// Total number of formula cells
    pub formula_count: usize,
    /// Number of cells calculated
    pub cells_calculated: usize,
    /// Number of iterations performed (for circular references)
    pub iterations: u32,
    /// Number of circular references detected
    pub circular_references: usize,
    /// Number of volatile cells recalculated
    pub volatile_cells: usize,
    /// Number of errors encountered during calculation
    pub errors: usize,
    /// Whether calculation converged (for iterative calculation)
    pub converged: bool,
}

/// Extension trait for Workbook to add calculation methods
pub trait WorkbookCalculationExt {
    /// Calculate all formulas in the workbook with default options
    fn calculate(&mut self) -> Result<CalculationStats>;

    /// Calculate all formulas with custom options
    fn calculate_with_options(&mut self, options: &CalculationOptions) -> Result<CalculationStats>;

    /// Calculate only the specified sheets (and their transitive cross-sheet dependencies)
    fn calculate_sheets(&mut self, sheets: &[usize]) -> Result<CalculationStats>;
}

impl WorkbookCalculationExt for Workbook {
    fn calculate(&mut self) -> Result<CalculationStats> {
        self.calculate_with_options(&CalculationOptions::default())
    }

    fn calculate_with_options(&mut self, options: &CalculationOptions) -> Result<CalculationStats> {
        let mut engine = CalculationEngine::new(options.clone());
        engine.calculate_all(self)
    }

    fn calculate_sheets(&mut self, sheets: &[usize]) -> Result<CalculationStats> {
        let options = CalculationOptions {
            sheets: sheets.to_vec(),
            ..Default::default()
        };
        self.calculate_with_options(&options)
    }
}

/// Pre-computed evaluation plan - the expensive DFS result that can be cached.
struct EvalPlan {
    eval_order: Vec<CellKey>,
    idx_to_cell: Vec<CellKey>,
    #[cfg_attr(not(feature = "parallel"), allow(dead_code))]
    cell_to_idx: AHashMap<CellKey, u32>,
    #[cfg_attr(not(feature = "parallel"), allow(dead_code))]
    depth: Vec<u32>,
    #[cfg_attr(not(feature = "parallel"), allow(dead_code))]
    max_depth: u32,
    dependents: DenseDependents,
    input_ranges: DenseInputRanges,
}

/// Persistent calculation cache stored on the `Workbook` between `calculate()` calls.
/// Contains everything needed to skip the parse + DFS phases on repeat calculations.
struct CalcCache {
    /// Requested calculation scope when this cache was built. Empty = full workbook.
    scope_key: Vec<usize>,
    /// Workbook structural generation when this cache was built.
    structural_gen: u64,
    /// Per-sheet topology generations when this cache was built.
    sheet_topology_gens: Vec<(usize, u64)>,
    /// Parsed formula ASTs.
    parsed_formulas: AHashMap<CellKey, FormulaExpr>,
    /// Volatile cells.
    volatile_cells: AHashSet<CellKey>,
    /// Cells involved in circular references.
    circular_cells: AHashSet<CellKey>,
    /// Workbook ranges whose values were consulted during planner-time dependency narrowing.
    value_sensitive_ranges: Vec<SensitiveRange>,
    /// Pre-computed evaluation plan.
    plan: EvalPlan,
}

impl CalcCache {
    fn range_overlaps_dirty(dirty: (u32, u16, u32, u16), sensitive: SensitiveRange) -> bool {
        let (dr1, dc1, dr2, dc2) = dirty;
        let (_, sr1, sc1, sr2, sc2) = sensitive;
        dr1 <= sr2 && dr2 >= sr1 && dc1 <= sc2 && dc2 >= sc1
    }

    /// Check whether this cache is still valid for the given workbook.
    fn is_valid(&self, workbook: &Workbook, scope_key: &[usize]) -> bool {
        if self.scope_key != scope_key {
            return false;
        }
        if self.structural_gen != workbook.structural_generation() {
            return false;
        }
        for &(sheet_idx, gen) in &self.sheet_topology_gens {
            match workbook.worksheet(sheet_idx) {
                Some(ws) if ws.topology_generation() == gen => {}
                _ => return false,
            }
        }

        for &(sheet_idx, sr1, sc1, sr2, sc2) in &self.value_sensitive_ranges {
            let Some(ws) = workbook.worksheet(sheet_idx) else {
                return false;
            };
            if ws
                .dirty_value_ranges()
                .iter()
                .any(|&dirty| Self::range_overlaps_dirty(dirty, (sheet_idx, sr1, sc1, sr2, sc2)))
            {
                return false;
            }
        }
        true
    }
}

fn normalized_scope_key(sheets: &[usize]) -> Vec<usize> {
    let mut key = sheets.to_vec();
    key.sort_unstable();
    key.dedup();
    key
}

/// The calculation engine
struct CalculationEngine {
    options: CalculationOptions,
    /// Parsed formula ASTs, keyed by CellKey
    parsed_formulas: AHashMap<CellKey, FormulaExpr>,
    /// Set of volatile cells
    volatile_cells: AHashSet<CellKey>,
    /// Cells involved in circular references
    circular_cells: AHashSet<CellKey>,
}

type FormulaCellIndex = HashMap<usize, Vec<(u32, u16)>>;
type RangeDependencyCache = AHashMap<(usize, u32, u32, u16, u16), Vec<CellKey>>;
type DensePrecedents = Vec<Box<[u32]>>;
type SensitiveRange = (usize, u32, u16, u32, u16);
type DenseDependents = Vec<Box<[u32]>>;
type DenseInputRanges = Vec<Box<[SensitiveRange]>>;

#[cfg(not(target_arch = "wasm32"))]
fn now() -> Option<Instant> {
    Some(Instant::now())
}

#[cfg(target_arch = "wasm32")]
fn now() -> Option<Instant> {
    None
}

fn elapsed_ms(start: Option<Instant>) -> f64 {
    start
        .map(|t| t.elapsed().as_secs_f64() * 1000.0)
        .unwrap_or(0.0)
}

#[derive(Debug, Clone)]
struct LevelTrace {
    depth: usize,
    width: usize,
    eval_ms: f64,
    write_ms: f64,
}

#[derive(Debug, Clone)]
struct ParallelEvalTrace {
    enabled: bool,
    num_threads: usize,
    total_levels: usize,
    non_empty_levels: usize,
    widest_level_depth: usize,
    widest_level_size: usize,
    avg_non_empty_level_width: f64,
    levels_ge_threads: usize,
    eval_phase_ms: f64,
    write_phase_ms: f64,
    top_levels: Vec<LevelTrace>,
}

#[derive(Debug, Clone)]
struct PlanBuildTrace {
    dep_materialize_ms: f64,
    dfs_ms: f64,
    edge_count: usize,
}

#[derive(Debug, Clone)]
struct CalculationTrace {
    cache_hit: bool,
    parse_ms: f64,
    plan_ms: f64,
    plan_dep_materialize_ms: f64,
    plan_dfs_ms: f64,
    plan_edge_count: usize,
    eval_ms: f64,
    spill_fixup_ms: f64,
    parallel: Option<ParallelEvalTrace>,
}

fn parallel_trace_enabled() -> bool {
    std::env::var_os("DUKE_SHEETS_PARALLEL_TRACE").is_some()
}

fn emit_calculation_trace(trace: &CalculationTrace) {
    eprintln!();
    eprintln!("=== duke-sheets parallel trace ===");
    eprintln!("cache_hit:      {}", trace.cache_hit);
    eprintln!("parse_ms:       {:.2}", trace.parse_ms);
    eprintln!("plan_ms:        {:.2}", trace.plan_ms);
    eprintln!("plan_dep_ms:    {:.2}", trace.plan_dep_materialize_ms);
    eprintln!("plan_dfs_ms:    {:.2}", trace.plan_dfs_ms);
    eprintln!("plan_edges:     {}", trace.plan_edge_count);
    eprintln!("eval_ms:        {:.2}", trace.eval_ms);
    eprintln!("spill_fixup_ms: {:.2}", trace.spill_fixup_ms);
    if let Some(p) = &trace.parallel {
        eprintln!("parallel:       {}", p.enabled);
        eprintln!("threads:        {}", p.num_threads);
        eprintln!("total_levels:   {}", p.total_levels);
        eprintln!("non_empty_lvls: {}", p.non_empty_levels);
        eprintln!(
            "widest_level:   depth {} ({} cells)",
            p.widest_level_depth, p.widest_level_size
        );
        eprintln!("avg_lvl_width:  {:.2}", p.avg_non_empty_level_width);
        eprintln!("lvls>=threads:  {}", p.levels_ge_threads);
        eprintln!("level_eval_ms:  {:.2}", p.eval_phase_ms);
        eprintln!("level_write_ms: {:.2}", p.write_phase_ms);
        if !p.top_levels.is_empty() {
            eprintln!("slowest_levels:");
            for level in &p.top_levels {
                eprintln!(
                    "  depth {:>5} width {:>7} eval_ms {:>8.2} write_ms {:>8.2}",
                    level.depth, level.width, level.eval_ms, level.write_ms
                );
            }
        }
    } else {
        eprintln!("parallel:       false");
    }
}

#[derive(Clone)]
struct ParsedFormulaTemplate {
    base_row: u32,
    ast: FormulaExpr,
    volatile: bool,
}

fn shift_formula_rows(expr: &mut FormulaExpr, row_delta: i32) {
    match expr {
        FormulaExpr::CellRef(cell_ref) => shift_cell_address_row(&mut cell_ref.address, row_delta),
        FormulaExpr::RangeRef(range_ref) => {
            shift_cell_address_row(&mut range_ref.range.start, row_delta);
            shift_cell_address_row(&mut range_ref.range.end, row_delta);
        }
        FormulaExpr::ExternalRef(ext) => shift_cell_address_row(&mut ext.address, row_delta),
        FormulaExpr::BinaryOp { left, right, .. } => {
            shift_formula_rows(left, row_delta);
            shift_formula_rows(right, row_delta);
        }
        FormulaExpr::UnaryOp { operand, .. } => shift_formula_rows(operand, row_delta),
        FormulaExpr::Function { args, .. } => {
            for arg in args {
                shift_formula_rows(arg, row_delta);
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    shift_formula_rows(cell, row_delta);
                }
            }
        }
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::Empty => {}
    }
}

fn shift_cell_address_row(addr: &mut CellAddress, row_delta: i32) {
    if addr.row_absolute || row_delta == 0 {
        return;
    }
    let shifted = addr.row as i32 + row_delta;
    addr.row = shifted.max(0) as u32;
}

fn parse_cell_ref_row_info(s: &str) -> Option<(usize, bool, u32)> {
    let b = s.as_bytes();
    let mut i = 0usize;

    if b.get(i) == Some(&b'$') {
        i += 1;
    }
    let col_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_uppercase() {
            i += 1;
        } else {
            break;
        }
    }
    if i == col_start {
        return None;
    }

    let row_abs = if b.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let row_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_digit() {
            i += 1;
        } else {
            break;
        }
    }
    if i == row_start {
        return None;
    }

    if let Some(&next) = b.get(i) {
        let next = next as char;
        if next.is_ascii_alphanumeric() || next == '_' || next == '.' || next == '(' {
            return None;
        }
    }

    let row0 = s[row_start..i].parse::<u32>().ok()?.saturating_sub(1);
    Some((i, row_abs, row0))
}

fn min_relative_row_delta(formula: &str, current_row: u32) -> i32 {
    let bytes = formula.as_bytes();
    let mut i = 0usize;
    let mut in_string = false;
    let mut min_delta = 0i32;

    while i < bytes.len() {
        if bytes[i] >= 0x80 {
            let ch = formula[i..].chars().next().unwrap();
            i += ch.len_utf8();
            continue;
        }
        let ch = bytes[i] as char;
        if ch == '"' {
            in_string = !in_string;
            i += 1;
            continue;
        }
        if !in_string {
            if i > 0 {
                let prev = bytes[i - 1] as char;
                if prev.is_ascii_alphanumeric() || prev == '_' || prev == '.' {
                    i += 1;
                    continue;
                }
            }
            if let Some((consumed, row_abs, row0)) = parse_cell_ref_row_info(&formula[i..]) {
                if !row_abs {
                    min_delta = min_delta.min(row0 as i32 - current_row as i32);
                }
                i += consumed;
                continue;
            }
        }
        i += 1;
    }

    min_delta
}

fn shift_a1_references_rows(formula: &str, row_delta: i32) -> String {
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;
    let mut in_string = false;

    while i < bytes.len() {
        if bytes[i] >= 0x80 {
            let ch = formula[i..].chars().next().unwrap();
            out.push(ch);
            i += ch.len_utf8();
            continue;
        }
        let ch = bytes[i] as char;
        if ch == '"' {
            in_string = !in_string;
            out.push(ch);
            i += 1;
            continue;
        }
        if !in_string {
            if i > 0 {
                let prev = bytes[i - 1] as char;
                if prev.is_ascii_alphanumeric() || prev == '_' || prev == '.' {
                    out.push(ch);
                    i += 1;
                    continue;
                }
            }
            if let Some((consumed, shifted)) = try_shift_cell_ref_rows(&formula[i..], row_delta) {
                out.push_str(&shifted);
                i += consumed;
                continue;
            }
        }
        out.push(ch);
        i += 1;
    }

    out
}

fn try_shift_cell_ref_rows(s: &str, row_delta: i32) -> Option<(usize, String)> {
    let b = s.as_bytes();
    let mut i = 0usize;

    let col_abs = if b.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let col_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_uppercase() {
            i += 1;
        } else {
            break;
        }
    }
    if i == col_start {
        return None;
    }
    let col_letters = &s[col_start..i];

    let row_abs = if b.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let row_start = i;
    while let Some(&c) = b.get(i) {
        if (c as char).is_ascii_digit() {
            i += 1;
        } else {
            break;
        }
    }
    if i == row_start {
        return None;
    }

    if let Some(&next) = b.get(i) {
        let next = next as char;
        if next.is_ascii_alphanumeric() || next == '_' || next == '.' || next == '(' {
            return None;
        }
    }

    let mut row = s[row_start..i].parse::<i32>().ok()?.saturating_sub(1);
    if !row_abs {
        row += row_delta;
    }

    let mut shifted = String::new();
    if col_abs {
        shifted.push('$');
    }
    shifted.push_str(col_letters);
    if row_abs {
        shifted.push('$');
    }
    if row < 0 {
        shifted.push_str("#REF!");
    } else {
        shifted.push_str(&(row as u32 + 1).to_string());
    }
    Some((i, shifted))
}

fn normalize_formula_row_template(formula: &str, current_row: u32) -> (String, u32) {
    let min_delta = min_relative_row_delta(formula, current_row);
    let base_row = (-min_delta).max(0) as u32;
    let shift = base_row as i32 - current_row as i32;
    (shift_a1_references_rows(formula, shift), base_row)
}

fn parse_with_row_template_cache(
    text: &str,
    row: u32,
    check_volatile: bool,
    template_cache: &DashMap<(String, u32), ParsedFormulaTemplate>,
) -> (Option<FormulaExpr>, bool) {
    let (template_text, base_row) = normalize_formula_row_template(text, row);
    let cache_key = (template_text, base_row);

    if let Some(existing) = template_cache.get(&cache_key) {
        let cached = existing.clone();
        drop(existing);
        let mut ast = cached.ast.clone();
        let delta = row as i32 - cached.base_row as i32;
        if delta != 0 {
            shift_formula_rows(&mut ast, delta);
        }
        return (Some(ast), cached.volatile);
    }

    let template_text_ref = &cache_key.0;
    match parse_formula(template_text_ref) {
        Ok(template_ast) => {
            let volatile = check_volatile && contains_volatile_function(&template_ast);
            template_cache.insert(
                cache_key,
                ParsedFormulaTemplate {
                    base_row,
                    ast: template_ast.clone(),
                    volatile,
                },
            );
            let mut ast = template_ast;
            let delta = row as i32 - base_row as i32;
            if delta != 0 {
                shift_formula_rows(&mut ast, delta);
            }
            (Some(ast), volatile)
        }
        Err(_) => (None, false),
    }
}

impl CalculationEngine {
    fn new(options: CalculationOptions) -> Self {
        Self {
            options,
            parsed_formulas: AHashMap::new(),
            volatile_cells: AHashSet::new(),
            circular_cells: AHashSet::new(),
        }
    }

    /// Calculate all formulas in the workbook.
    ///
    /// When a valid `CalcCache` exists on the workbook (from a previous
    /// `calculate()` call) and no cells/sheets have been mutated since,
    /// the expensive parse + DFS phases are skipped entirely and only
    /// the evaluation phase runs.
    fn calculate_all(&mut self, workbook: &mut Workbook) -> Result<CalculationStats> {
        let mut stats = CalculationStats::default();
        let trace_enabled = parallel_trace_enabled();
        let mut trace = CalculationTrace {
            cache_hit: false,
            parse_ms: 0.0,
            plan_ms: 0.0,
            plan_dep_materialize_ms: 0.0,
            plan_dfs_ms: 0.0,
            plan_edge_count: 0,
            eval_ms: 0.0,
            spill_fixup_ms: 0.0,
            parallel: None,
        };
        let pending_value_sensitive_ranges: Vec<SensitiveRange>;

        let scope_key = normalized_scope_key(&self.options.sheets);
        let cached = if self.options.force_full_calculation {
            let _ = workbook.take_calc_cache();
            None
        } else {
            workbook
                .take_calc_cache()
                .and_then(|c| c.downcast::<CalcCache>().ok())
                .filter(|c| c.is_valid(workbook, &scope_key))
        };
        let dirty_ranges = collect_dirty_value_ranges(workbook);

        let plan;
        let mut affected_flags: Option<Vec<bool>> = None;
        if let Some(mut cache) = cached {
            trace.cache_hit = true;
            // Cache hit - restore parsed formulas + plan from cache.
            self.parsed_formulas = std::mem::take(&mut cache.parsed_formulas);
            self.volatile_cells = std::mem::take(&mut cache.volatile_cells);
            self.circular_cells = std::mem::take(&mut cache.circular_cells);
            pending_value_sensitive_ranges = std::mem::take(&mut cache.value_sensitive_ranges);
            stats.formula_count = self.parsed_formulas.len();
            stats.volatile_cells = self.volatile_cells.len();
            plan = cache.plan;
            if !dirty_ranges.is_empty() {
                affected_flags = Some(collect_affected_formula_flags(&plan, &dirty_ranges));
                if let Some(flags) = affected_flags.as_mut() {
                    for &cell in &self.volatile_cells {
                        if let Some(&idx) = plan.cell_to_idx.get(&cell) {
                            flags[idx as usize] = true;
                        }
                    }
                }
            }
        } else {
            // Cache miss - full parse + DFS.
            let t_parse = now();
            self.parse_formulas(workbook, &mut stats)?;
            if trace_enabled {
                trace.parse_ms = elapsed_ms(t_parse);
            }
            if stats.formula_count == 0 {
                for i in 0..workbook.sheet_count() {
                    if let Some(ws) = workbook.worksheet_mut(i) {
                        ws.clear_dirty_value_ranges();
                    }
                }
                if trace_enabled {
                    emit_calculation_trace(&trace);
                }
                return Ok(stats);
            }
            let t_plan = now();
            let (built_plan, plan_trace, value_sensitive_ranges) = self.build_eval_plan(workbook);
            plan = built_plan;
            if trace_enabled {
                trace.plan_ms = elapsed_ms(t_plan);
                trace.plan_dep_materialize_ms = plan_trace.dep_materialize_ms;
                trace.plan_dfs_ms = plan_trace.dfs_ms;
                trace.plan_edge_count = plan_trace.edge_count;
            }
            pending_value_sensitive_ranges = value_sensitive_ranges;
        }

        #[cfg(debug_assertions)]
        debug_validate_eval_plan(
            &plan,
            &self.parsed_formulas,
            &self.volatile_cells,
            &self.circular_cells,
        );

        // Clear stale image metadata before re-evaluating.
        for i in 0..workbook.sheet_count() {
            if let Some(ws) = workbook.worksheet(i) {
                ws.clear_image_metadata();
            }
        }

        // Evaluate all formulas using the (possibly cached) plan.
        let t_eval = now();
        let (parallel_trace, spill_fixup_ms) =
            self.execute_eval_plan(workbook, &plan, affected_flags.as_deref(), &mut stats)?;
        if trace_enabled {
            trace.eval_ms = elapsed_ms(t_eval);
            trace.parallel = parallel_trace;
            trace.spill_fixup_ms = spill_fixup_ms;
        }

        let sheet_topology_gens: Vec<(usize, u64)> = (0..workbook.sheet_count())
            .map(|i| {
                (
                    i,
                    workbook
                        .worksheet(i)
                        .map_or(0, |ws| ws.topology_generation()),
                )
            })
            .collect();
        let cache = CalcCache {
            scope_key,
            structural_gen: workbook.structural_generation(),
            sheet_topology_gens,
            parsed_formulas: std::mem::take(&mut self.parsed_formulas),
            volatile_cells: std::mem::take(&mut self.volatile_cells),
            circular_cells: std::mem::take(&mut self.circular_cells),
            value_sensitive_ranges: pending_value_sensitive_ranges,
            plan,
        };
        workbook.set_calc_cache(Box::new(cache));

        for i in 0..workbook.sheet_count() {
            if let Some(ws) = workbook.worksheet_mut(i) {
                ws.clear_dirty_value_ranges();
            }
        }

        if trace_enabled {
            emit_calculation_trace(&trace);
        }

        Ok(stats)
    }

    /// Parse all formulas and store ASTs. Does NOT build the dependency graph.
    /// When `options.sheets` is set, discovers cross-sheet refs on-the-fly
    /// so each formula is parsed exactly once.
    ///
    /// Formulas are parsed in parallel (when the `parallel` feature is active
    /// and `max_threads != Some(1)`) using a wave-based approach: each wave
    /// parses the current batch of sheets, then scans for cross-sheet refs
    /// to discover new sheets for the next wave.
    fn parse_formulas(&mut self, workbook: &Workbook, stats: &mut CalculationStats) -> Result<()> {
        let sheet_count = workbook.sheet_count();
        let scoped = !self.options.sheets.is_empty();
        let mut included: HashSet<usize> = if scoped {
            self.options.sheets.iter().copied().collect()
        } else {
            (0..sheet_count).collect()
        };
        let mut pending: Vec<usize> = included.iter().copied().collect();
        pending.sort_unstable();
        let check_volatile = self.options.calculate_volatile;
        let template_cache: DashMap<(String, u32), ParsedFormulaTemplate> = DashMap::new();

        while !pending.is_empty() {
            // Process each sheet in the current wave.  Formulas within
            // a sheet are parsed in parallel (zero-copy: &str borrows from
            // the sheet).  Sheets are processed sequentially so cross-sheet
            // discovery can feed the next wave.
            let wave = std::mem::take(&mut pending);
            for &sheet_idx in &wave {
                if sheet_idx >= sheet_count {
                    continue;
                }
                let sheet = match workbook.worksheet(sheet_idx) {
                    Some(s) => s,
                    None => continue,
                };

                // Collect this sheet's formula cells (borrows &str from sheet).
                let cells: Vec<(u32, u16, &str)> = sheet.formula_cells().collect();
                if cells.is_empty() {
                    continue;
                }

                // Parse in parallel within this sheet when worthwhile.
                let use_par = self.should_use_parallel(cells.len());
                let parsed: Vec<(CellKey, Option<FormulaExpr>, bool)> = if use_par {
                    #[cfg(feature = "parallel")]
                    {
                        cells
                            .par_iter()
                            .map(|&(row, col, text)| {
                                let key = CellKey::new(sheet_idx, row, col);
                                let (ast, vol) = parse_with_row_template_cache(
                                    text,
                                    row,
                                    check_volatile,
                                    &template_cache,
                                );
                                (key, ast, vol)
                            })
                            .collect()
                    }
                    #[cfg(not(feature = "parallel"))]
                    {
                        cells
                            .iter()
                            .map(|&(row, col, text)| {
                                let key = CellKey::new(sheet_idx, row, col);
                                let (ast, vol) = parse_with_row_template_cache(
                                    text,
                                    row,
                                    check_volatile,
                                    &template_cache,
                                );
                                (key, ast, vol)
                            })
                            .collect()
                    }
                } else {
                    cells
                        .iter()
                        .map(|&(row, col, text)| {
                            let key = CellKey::new(sheet_idx, row, col);
                            let (ast, vol) = parse_with_row_template_cache(
                                text,
                                row,
                                check_volatile,
                                &template_cache,
                            );
                            (key, ast, vol)
                        })
                        .collect()
                };

                // Store results and discover cross-sheet refs.
                for (key, ast_opt, is_volatile) in parsed {
                    if let Some(ast) = ast_opt {
                        if scoped {
                            for ref_sheet in extract_sheet_refs(&ast, workbook) {
                                if !included.contains(&ref_sheet) {
                                    included.insert(ref_sheet);
                                    pending.push(ref_sheet);
                                }
                            }
                        }
                        if is_volatile {
                            self.volatile_cells.insert(key);
                        }
                        self.parsed_formulas.insert(key, ast);
                        stats.formula_count += 1;
                    } else {
                        stats.errors += 1;
                    }
                }
            }
        }
        stats.volatile_cells = self.volatile_cells.len();
        Ok(())
    }

    /// Build the evaluation plan via iterative post-order DFS.
    ///
    /// Computes the correct evaluation order by extracting dependencies
    /// on-the-fly from parsed ASTs.  Each formula is visited at most once,
    /// so total work is O(V + E_formula).  The resulting `EvalPlan` can be
    /// cached and reused across `calculate()` calls when the workbook has
    /// not been mutated.
    fn build_eval_plan(
        &mut self,
        workbook: &Workbook,
    ) -> (EvalPlan, PlanBuildTrace, Vec<SensitiveRange>) {
        // Transient spatial index for range→formula-cell lookups.
        let formula_cell_set: AHashSet<CellKey> = self.parsed_formulas.keys().copied().collect();
        let formula_cell_index = build_formula_cell_index(&formula_cell_set);

        // Row-major seed order: financial models mostly flow top-down/
        // left-to-right, so most cells resolve on first visit.
        let mut seed_order: Vec<CellKey> = self.parsed_formulas.keys().copied().collect();
        seed_order.sort_unstable_by(|a, b| {
            a.sheet
                .cmp(&b.sheet)
                .then(a.row.cmp(&b.row))
                .then(a.col.cmp(&b.col))
        });
        let n = self.parsed_formulas.len();
        let cell_to_idx: AHashMap<CellKey, u32> = seed_order
            .iter()
            .enumerate()
            .map(|(i, &k)| (k, i as u32))
            .collect();
        // Dense-indexed AST table: avoids AHashMap lookups during DFS.
        let asts: Vec<&FormulaExpr> = seed_order
            .iter()
            .map(|k| &self.parsed_formulas[k])
            .collect();
        let t_dep = now();
        let (precedents, input_ranges, value_sensitive_ranges) = build_dense_precedents(
            &seed_order,
            &asts,
            workbook,
            &formula_cell_set,
            &formula_cell_index,
            &cell_to_idx,
            self.should_use_parallel(n),
        );
        let dependents = build_dense_dependents(&precedents);
        let dep_materialize_ms = elapsed_ms(t_dep);
        let edge_count = precedents.iter().map(|deps| deps.len()).sum();

        // Phase 1 - build evaluation order via iterative post-order DFS.
        //
        // Each cell's formula-cell dependencies are extracted on-the-fly
        // (no persistent precedent map).  Every cell is visited at most
        // once, so total work is O(V + E_formula).
        //
        // State per cell: 0=unvisited, 1=in_stack, 2=visited.
        let mut state: Vec<u8> = vec![0u8; n];
        let mut is_circular: Vec<bool> = vec![false; n];
        let mut depth: Vec<u32> = vec![0u32; n];
        let mut max_depth: u32 = 0;
        let mut eval_order: Vec<CellKey> = Vec::with_capacity(n);
        let mut stack: Vec<(CellKey, u32, usize)> = Vec::new();

        let t_dfs = now();
        for (seed_idx, &seed) in seed_order.iter().enumerate() {
            let seed_idx = seed_idx as u32;
            if state[seed_idx as usize] == 2 {
                continue;
            }
            state[seed_idx as usize] = 1; // in_stack
            stack.push((seed, seed_idx, 0));

            while !stack.is_empty() {
                let (next_dep, back_edge_idx) = {
                    let (_, si, idx) = stack.last_mut().unwrap();
                    let deps = &precedents[*si as usize];
                    let mut found: Option<(CellKey, u32)> = None;
                    let mut back_idx: Option<u32> = None;
                    while *idx < deps.len() {
                        let di = deps[*idx];
                        *idx += 1;
                        let dep = seed_order[di as usize];
                        let s = state[di as usize];
                        if s == 0 {
                            found = Some((dep, di));
                            break;
                        }
                        if s == 1 {
                            back_idx = Some(di);
                        }
                    }
                    (found, back_idx)
                };

                // Mark all cells on the stack from the back-edge target
                // to the top - they all participate in the cycle.
                if let Some(target_idx) = back_edge_idx {
                    let mut marking = false;
                    for &(_, si, _) in stack.iter() {
                        if si == target_idx {
                            marking = true;
                        }
                        if marking {
                            is_circular[si as usize] = true;
                        }
                    }
                }

                if let Some((dep, di)) = next_dep {
                    state[di as usize] = 1; // in_stack
                    stack.push((dep, di, 0));
                } else {
                    // All deps visited - emit this cell.
                    let (cell, ci, _) = stack.pop().unwrap();
                    state[ci as usize] = 2; // visited
                                            // Depth = 1 + max depth of formula-cell deps.
                    let d = precedents[ci as usize]
                        .iter()
                        .map(|&di| depth[di as usize])
                        .max()
                        .unwrap_or(0)
                        + 1;
                    depth[ci as usize] = d;
                    if d > max_depth {
                        max_depth = d;
                    }
                    eval_order.push(cell);
                }
            }
        }

        // Collect circular cells for stats reporting.
        self.circular_cells = seed_order
            .iter()
            .enumerate()
            .filter(|&(i, _)| is_circular[i])
            .map(|(_, &k)| k)
            .collect();

        let dfs_ms = elapsed_ms(t_dfs);

        (
            EvalPlan {
                eval_order,
                idx_to_cell: seed_order,
                cell_to_idx,
                depth,
                max_depth,
                dependents,
                input_ranges,
            },
            PlanBuildTrace {
                dep_materialize_ms,
                dfs_ms,
                edge_count,
            },
            value_sensitive_ranges,
        )
    }

    /// Execute a pre-computed evaluation plan.
    ///
    /// Evaluates formulas in the order given by `plan.eval_order`, using
    /// parallel level-based evaluation when the `parallel` feature is active
    /// and `max_threads != Some(1)`.  After evaluation, performs targeted
    /// spill fixup for any formulas whose results spilled into adjacent cells.
    fn execute_eval_plan(
        &self,
        workbook: &mut Workbook,
        plan: &EvalPlan,
        affected_flags: Option<&[bool]>,
        stats: &mut CalculationStats,
    ) -> Result<(Option<ParallelEvalTrace>, f64)> {
        let n = plan.eval_order.len();
        let use_parallel = self.should_use_parallel(n);
        let mut spill_ranges: Vec<(usize, u32, u16, u32, u16)> = Vec::new();
        #[allow(unused_mut)]
        let mut parallel_trace: Option<ParallelEvalTrace> = None;

        // Build shared evaluation cache for this calculation pass.
        let sheet_names: ahash::AHashMap<String, usize> = (0..workbook.sheet_count())
            .filter_map(|i| workbook.worksheet(i).map(|ws| (ws.name().to_string(), i)))
            .collect();
        let eval_cache = duke_sheets_formula::EvalCache::new(sheet_names);

        if let Some(flags) = affected_flags {
            if !flags.iter().any(|&v| v) && self.volatile_cells.is_empty() {
                stats.circular_references = self.circular_cells.len();
                stats.iterations = 1;
                stats.converged = true;
                return Ok((parallel_trace, 0.0));
            }
        }

        if use_parallel {
            #[cfg(feature = "parallel")]
            {
                let trace = self.evaluate_parallel(
                    workbook,
                    &plan.eval_order,
                    &plan.idx_to_cell,
                    &plan.cell_to_idx,
                    &plan.depth,
                    plan.max_depth,
                    affected_flags,
                    stats,
                    &mut spill_ranges,
                    &eval_cache,
                );
                parallel_trace = Some(trace);
            }
        } else {
            for &cell_key in &plan.eval_order {
                let include = affected_flags.is_none_or(|flags| {
                    plan.cell_to_idx
                        .get(&cell_key)
                        .is_some_and(|&idx| flags[idx as usize])
                });
                if include {
                    let did_spill = self.evaluate_and_store(workbook, cell_key, stats, &eval_cache);
                    if did_spill {
                        self.record_spill(workbook, cell_key, &mut spill_ranges);
                    }
                }
            }
        }

        // Phase 3 - targeted spill fixup: only re-evaluate formulas whose
        // ASTs reference cells inside a spill range.
        let t_spill = now();
        if !spill_ranges.is_empty() {
            for &cell_key in &plan.eval_order {
                let include = affected_flags.is_none_or(|flags| {
                    plan.cell_to_idx
                        .get(&cell_key)
                        .is_some_and(|&idx| flags[idx as usize])
                });
                if !include {
                    continue;
                }
                if workbook
                    .worksheet(cell_key.sheet)
                    .is_some_and(|s| s.is_spill_source(cell_key.row, cell_key.col))
                {
                    continue;
                }
                if let Some(ast) = self.parsed_formulas.get(&cell_key) {
                    if ast_touches_spill_range(ast, cell_key.sheet, workbook, &spill_ranges) {
                        self.evaluate_and_store(workbook, cell_key, stats, &eval_cache);
                    }
                }
            }
        }
        let _spill_fixup_ms = elapsed_ms(t_spill);

        stats.circular_references = self.circular_cells.len();
        stats.iterations = 1;
        stats.converged = true;
        Ok((parallel_trace, _spill_fixup_ms))
    }

    /// Evaluate a single formula cell and store its result.
    ///
    /// Returns `true` if the formula produced an array that actually spilled
    fn evaluate_and_store(
        &self,
        workbook: &mut Workbook,
        cell_key: CellKey,
        stats: &mut CalculationStats,
        eval_cache: &duke_sheets_formula::EvalCache,
    ) -> bool {
        let ast = match self.parsed_formulas.get(&cell_key) {
            Some(ast) => ast,
            None => return false,
        };

        // Circular reference cells are evaluated normally - their self-references
        // read the cached value from the file (the "previous iteration" result).
        // This handles the common Excel pattern =IF(cond, val, SELF) correctly.

        // Clear any existing spill targets before re-evaluating
        if let Some(sheet) = workbook.worksheet_mut(cell_key.sheet) {
            sheet.clear_spill(cell_key.row, cell_key.col);
        }

        // Evaluate in a block so immutable borrows drop before we mutably store results.
        let result = {
            let wb_ref: &Workbook = workbook;
            let image_sink = |sheet: usize, row: u32, col: u16, info: ImageInfo| {
                if let Some(ws) = wb_ref.worksheet(sheet) {
                    ws.set_image_at(row, col, info);
                }
            };
            let mut ctx =
                EvaluationContext::new(Some(wb_ref), cell_key.sheet, cell_key.row, cell_key.col);
            ctx.web_service_fn = self.options.web_service_fn.as_deref();
            ctx.rtd_fn = self.options.rtd_fn.as_deref();
            ctx.image_sink = Some(&image_sink);
            ctx.eval_cache = Some(eval_cache);

            match evaluate(ast, &ctx) {
                Ok(value) => value,
                Err(_e) => {
                    stats.errors += 1;
                    FormulaValue::Error(duke_sheets_core::CellError::Value)
                }
            }
        };

        // Store the result
        if let Some(sheet) = workbook.worksheet_mut(cell_key.sheet) {
            match result {
                FormulaValue::Array { data: array, .. } => {
                    let cell_array: Vec<Vec<CellValue>> = array
                        .into_iter()
                        .map(|row| row.into_iter().map(|v| v.into()).collect())
                        .collect();
                    let _ = sheet.set_array_formula_result(cell_key.row, cell_key.col, cell_array);
                }
                _ => {
                    let _ = sheet.set_formula_result(cell_key.row, cell_key.col, result.into());
                }
            }
        }
        // Only report a spill if set_array_formula_result actually created
        // spill targets (1×1 arrays are stored as scalars, no spill).
        let did_spill = workbook
            .worksheet(cell_key.sheet)
            .is_some_and(|s| s.is_spill_source(cell_key.row, cell_key.col));
        stats.cells_calculated += 1;
        did_spill
    }

    /// Record a spill range for targeted fixup later.
    fn record_spill(
        &self,
        workbook: &Workbook,
        cell_key: CellKey,
        spill_ranges: &mut Vec<(usize, u32, u16, u32, u16)>,
    ) {
        if let Some(sheet) = workbook.worksheet(cell_key.sheet) {
            if let Some(info) = sheet.get_spill_info(cell_key.row, cell_key.col) {
                let (dr, dc) = info.end_offsets();
                spill_ranges.push((
                    cell_key.sheet,
                    cell_key.row,
                    cell_key.col,
                    cell_key.row + dr,
                    cell_key.col + dc,
                ));
            }
        }
    }

    /// Decide whether to use parallel evaluation.
    #[allow(unused_variables)]
    fn should_use_parallel(&self, formula_count: usize) -> bool {
        // Explicit serial override
        if self.options.max_threads == Some(1) {
            return false;
        }
        // Not enough work to justify thread-pool overhead
        if formula_count < 5_000 {
            return false;
        }
        #[cfg(feature = "parallel")]
        {
            true
        }
        #[cfg(not(feature = "parallel"))]
        {
            false
        }
    }

    #[cfg(feature = "parallel")]
    /// Evaluate a single formula (read-only) and return the result.
    /// Used by the parallel path to separate evaluation from storage.
    /// Returns `(value, was_eval_error)` - the bool is `true` only when
    /// `evaluate()` itself returned `Err`, NOT when the formula legitimately
    /// produces an error value like `=1/0`.
    fn evaluate_formula(
        &self,
        workbook: &Workbook,
        cell_key: CellKey,
        eval_cache: &duke_sheets_formula::EvalCache,
    ) -> Option<(FormulaValue, bool)> {
        let ast = self.parsed_formulas.get(&cell_key)?;
        let image_sink = |sheet: usize, row: u32, col: u16, info: ImageInfo| {
            if let Some(ws) = workbook.worksheet(sheet) {
                ws.set_image_at(row, col, info);
            }
        };
        let mut ctx =
            EvaluationContext::new(Some(workbook), cell_key.sheet, cell_key.row, cell_key.col);
        ctx.web_service_fn = self.options.web_service_fn.as_deref();
        ctx.rtd_fn = self.options.rtd_fn.as_deref();
        ctx.image_sink = Some(&image_sink);
        ctx.eval_cache = Some(eval_cache);
        match evaluate(ast, &ctx) {
            Ok(value) => Some((value, false)),
            Err(_) => Some((
                FormulaValue::Error(duke_sheets_core::CellError::Value),
                true,
            )),
        }
    }

    #[cfg(feature = "parallel")]
    /// Store a pre-computed formula result into the workbook.
    /// Returns `true` if the result spilled.
    fn store_result(workbook: &mut Workbook, cell_key: CellKey, result: FormulaValue) -> bool {
        if let Some(sheet) = workbook.worksheet_mut(cell_key.sheet) {
            sheet.clear_spill(cell_key.row, cell_key.col);
            match result {
                FormulaValue::Array { data: array, .. } => {
                    let cell_array: Vec<Vec<CellValue>> = array
                        .into_iter()
                        .map(|row| row.into_iter().map(|v| v.into()).collect())
                        .collect();
                    let _ = sheet.set_array_formula_result(cell_key.row, cell_key.col, cell_array);
                }
                _ => {
                    let _ = sheet.set_formula_result(cell_key.row, cell_key.col, result.into());
                }
            }
        }
        workbook
            .worksheet(cell_key.sheet)
            .map_or(false, |s| s.is_spill_source(cell_key.row, cell_key.col))
    }

    /// Parallel level-based evaluation.
    ///
    /// Cells are grouped by dependency depth.  All cells at the same depth
    /// have their dependencies satisfied by earlier levels, so they can be
    /// evaluated in parallel.  After each level, results are written
    /// serially to the workbook before the next level starts.
    #[cfg(feature = "parallel")]
    fn evaluate_parallel(
        &self,
        workbook: &mut Workbook,
        eval_order: &[CellKey],
        idx_to_cell: &[CellKey],
        cell_to_idx: &AHashMap<CellKey, u32>,
        depth: &[u32],
        max_depth: u32,
        affected_flags: Option<&[bool]>,
        stats: &mut CalculationStats,
        spill_ranges: &mut Vec<(usize, u32, u16, u32, u16)>,
        eval_cache: &duke_sheets_formula::EvalCache,
    ) -> ParallelEvalTrace {
        let trace_enabled = parallel_trace_enabled();
        // Build levels: group cells by depth.
        let num_levels = max_depth as usize + 1;
        let mut levels: Vec<Vec<CellKey>> = vec![Vec::new(); num_levels];
        if let Some(flags) = affected_flags {
            for (idx, &cell_key) in idx_to_cell.iter().enumerate() {
                if flags[idx] {
                    let d = depth[idx] as usize;
                    levels[d].push(cell_key);
                }
            }
        } else {
            for &cell_key in eval_order {
                if let Some(&idx) = cell_to_idx.get(&cell_key) {
                    let d = depth[idx as usize] as usize;
                    levels[d].push(cell_key);
                }
            }
        }

        let non_empty_levels = levels.iter().filter(|level| !level.is_empty()).count();
        let (widest_level_depth, widest_level_size) = levels
            .iter()
            .enumerate()
            .max_by_key(|(_, level)| level.len())
            .map(|(depth, level)| (depth, level.len()))
            .unwrap_or((0, 0));

        // Build a thread pool with the requested thread count.
        let num_threads = self
            .options
            .max_threads
            .unwrap_or_else(|| rayon::current_num_threads());
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(num_threads)
            .build()
            .expect("failed to build rayon thread pool");

        let mut level_traces: Vec<LevelTrace> = Vec::new();
        let mut total_eval_ms = 0.0f64;
        let mut total_write_ms = 0.0f64;

        pool.install(|| {
            for (depth_idx, level) in levels.iter().enumerate() {
                if level.is_empty() {
                    continue;
                }

                // Evaluate all cells at this level in parallel.
                // Each evaluation only reads from the workbook - all
                // deps at lower depths have already been written.
                let wb_ref: &Workbook = workbook;
                let t_eval = now();
                let results: Vec<(CellKey, FormulaValue, bool)> = level
                    .par_iter()
                    .filter_map(|&cell_key| {
                        self.evaluate_formula(wb_ref, cell_key, eval_cache)
                            .map(|(val, was_error)| (cell_key, val, was_error))
                    })
                    .collect();
                let eval_ms = t_eval
                    .map(|t| t.elapsed().as_secs_f64() * 1000.0)
                    .unwrap_or(0.0);

                // Write results serially and track stats.
                let mut level_errors = 0usize;
                let t_write = now();
                for (cell_key, result, was_error) in results {
                    if was_error {
                        level_errors += 1;
                    }
                    let did_spill = Self::store_result(workbook, cell_key, result);
                    if did_spill {
                        self.record_spill(workbook, cell_key, spill_ranges);
                    }
                    stats.cells_calculated += 1;
                }
                stats.errors += level_errors;

                let write_ms = t_write
                    .map(|t| t.elapsed().as_secs_f64() * 1000.0)
                    .unwrap_or(0.0);
                if trace_enabled {
                    total_eval_ms += eval_ms;
                    total_write_ms += write_ms;
                    level_traces.push(LevelTrace {
                        depth: depth_idx,
                        width: level.len(),
                        eval_ms,
                        write_ms,
                    });
                }
            }
        });

        if trace_enabled {
            level_traces.sort_by(|a, b| {
                (b.eval_ms + b.write_ms)
                    .partial_cmp(&(a.eval_ms + a.write_ms))
                    .unwrap_or(std::cmp::Ordering::Equal)
            });
        }

        let total_scheduled_cells: usize = levels.iter().map(|level| level.len()).sum();
        let avg_non_empty_level_width = if non_empty_levels == 0 {
            0.0
        } else {
            total_scheduled_cells as f64 / non_empty_levels as f64
        };
        let levels_ge_threads = levels
            .iter()
            .filter(|level| level.len() >= num_threads && !level.is_empty())
            .count();

        ParallelEvalTrace {
            enabled: true,
            num_threads,
            total_levels: num_levels,
            non_empty_levels,
            widest_level_depth,
            widest_level_size,
            avg_non_empty_level_width,
            levels_ge_threads,
            eval_phase_ms: total_eval_ms,
            write_phase_ms: total_write_ms,
            top_levels: if trace_enabled {
                level_traces.into_iter().take(5).collect()
            } else {
                Vec::new()
            },
        }
    }
}

/// Check whether any cell or range reference in `expr` overlaps a spill range.
/// Used by the targeted spill fixup to avoid re-evaluating formulas that cannot
/// be affected by newly-created spill targets.
fn ast_touches_spill_range(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    spill_ranges: &[(usize, u32, u16, u32, u16)],
) -> bool {
    match expr {
        FormulaExpr::CellRef(cr) => {
            let si = cr
                .sheet
                .as_ref()
                .and_then(|n| workbook.sheet_index(n))
                .unwrap_or(current_sheet);
            let r = cr.address.row;
            let c = cr.address.col;
            spill_ranges
                .iter()
                .any(|&(s, r1, c1, r2, c2)| s == si && r >= r1 && r <= r2 && c >= c1 && c <= c2)
        }
        FormulaExpr::RangeRef(rr) => {
            let si = rr
                .sheet
                .as_ref()
                .and_then(|n| workbook.sheet_index(n))
                .unwrap_or(current_sheet);
            let sr = rr.range.start.row;
            let er = rr.range.end.row;
            let sc = rr.range.start.col;
            let ec = rr.range.end.col;
            spill_ranges
                .iter()
                .any(|&(s, r1, c1, r2, c2)| s == si && sr <= r2 && er >= r1 && sc <= c2 && ec >= c1)
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            ast_touches_spill_range(left, current_sheet, workbook, spill_ranges)
                || ast_touches_spill_range(right, current_sheet, workbook, spill_ranges)
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            ast_touches_spill_range(operand, current_sheet, workbook, spill_ranges)
        }
        FormulaExpr::Function { args, .. } => args
            .iter()
            .any(|a| ast_touches_spill_range(a, current_sheet, workbook, spill_ranges)),
        FormulaExpr::Array(rows) => rows.iter().any(|row| {
            row.iter()
                .any(|cell| ast_touches_spill_range(cell, current_sheet, workbook, spill_ranges))
        }),
        _ => false,
    }
}

#[cfg(test)]
fn extract_references(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    formula_cells: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
) -> Vec<CellKey> {
    let mut refs = Vec::new();
    let mut range_dep_cache: RangeDependencyCache = AHashMap::new();
    let mut value_sensitive_ranges = Vec::new();
    extract_references_recursive(
        expr,
        current_sheet,
        workbook,
        formula_cells,
        formula_cell_index,
        &mut range_dep_cache,
        &mut value_sensitive_ranges,
        &mut refs,
    );
    refs
}

#[allow(clippy::too_many_arguments)]
fn extract_references_recursive(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    formula_cells: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
    range_dep_cache: &mut RangeDependencyCache,
    value_sensitive_ranges: &mut Vec<SensitiveRange>,
    refs: &mut Vec<CellKey>,
) {
    match expr {
        FormulaExpr::CellRef(cell_ref) => {
            let sheet_idx = cell_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);

            let key = CellKey::new(sheet_idx, cell_ref.address.row, cell_ref.address.col);
            // Only track deps on formula cells - static cells never change
            if formula_cells.contains(&key) {
                refs.push(key);
            }
        }
        FormulaExpr::RangeRef(range_ref) => {
            let sheet_idx = range_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);

            let start_row = range_ref.range.start.row;
            let end_row = range_ref.range.end.row;
            let start_col = range_ref.range.start.col;
            let end_col = range_ref.range.end.col;
            push_range_references(
                sheet_idx,
                start_row,
                end_row,
                start_col,
                end_col,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                refs,
            );
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            extract_references_recursive(
                left,
                current_sheet,
                workbook,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                value_sensitive_ranges,
                refs,
            );
            extract_references_recursive(
                right,
                current_sheet,
                workbook,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                value_sensitive_ranges,
                refs,
            );
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            extract_references_recursive(
                operand,
                current_sheet,
                workbook,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                value_sensitive_ranges,
                refs,
            );
        }
        FormulaExpr::Function { name, args } => {
            if name.eq_ignore_ascii_case("INDEX") {
                if let Some(FormulaExpr::RangeRef(range_ref)) = args.first() {
                    for arg in args.iter().skip(1) {
                        extract_references_recursive(
                            arg,
                            current_sheet,
                            workbook,
                            formula_cells,
                            formula_cell_index,
                            range_dep_cache,
                            value_sensitive_ranges,
                            refs,
                        );
                    }

                    let sheet_idx = range_ref
                        .sheet
                        .as_ref()
                        .and_then(|name| workbook.sheet_index(name))
                        .unwrap_or(current_sheet);

                    let mut row_start = range_ref.range.start.row;
                    let mut row_end = range_ref.range.end.row;
                    let mut col_start = range_ref.range.start.col;
                    let mut col_end = range_ref.range.end.col;

                    if let Some(row_expr) = args.get(1) {
                        let (row_idx, mut sensitive) = try_static_index_coord(
                            row_expr,
                            current_sheet,
                            workbook,
                            formula_cells,
                            formula_cell_index,
                            range_dep_cache,
                        );
                        value_sensitive_ranges.append(&mut sensitive);
                        if let Some(row_idx) = row_idx {
                            let selected = range_ref
                                .range
                                .start
                                .row
                                .saturating_add((row_idx - 1) as u32);
                            if selected <= range_ref.range.end.row {
                                row_start = selected;
                                row_end = selected;
                            }
                        }
                    }

                    if let Some(col_expr) = args.get(2) {
                        let (col_idx, mut sensitive) = try_static_index_coord(
                            col_expr,
                            current_sheet,
                            workbook,
                            formula_cells,
                            formula_cell_index,
                            range_dep_cache,
                        );
                        value_sensitive_ranges.append(&mut sensitive);
                        if let Some(col_idx) = col_idx {
                            let selected = range_ref
                                .range
                                .start
                                .col
                                .saturating_add((col_idx - 1) as u16);
                            if selected <= range_ref.range.end.col {
                                col_start = selected;
                                col_end = selected;
                            }
                        }
                    }

                    push_range_references(
                        sheet_idx,
                        row_start,
                        row_end,
                        col_start,
                        col_end,
                        formula_cells,
                        formula_cell_index,
                        range_dep_cache,
                        refs,
                    );
                    return;
                }
            }

            for arg in args {
                extract_references_recursive(
                    arg,
                    current_sheet,
                    workbook,
                    formula_cells,
                    formula_cell_index,
                    range_dep_cache,
                    value_sensitive_ranges,
                    refs,
                );
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    extract_references_recursive(
                        cell,
                        current_sheet,
                        workbook,
                        formula_cells,
                        formula_cell_index,
                        range_dep_cache,
                        value_sensitive_ranges,
                        refs,
                    );
                }
            }
        }
        // Literals and unresolvable references
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => {}
        FormulaExpr::StructuredRef(sr) => {
            extract_structured_ref_deps(
                sr,
                current_sheet,
                workbook,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                value_sensitive_ranges,
                refs,
            );
        }
    }
}

fn collect_value_sensitive_ranges(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    ranges: &mut Vec<SensitiveRange>,
) -> bool {
    match expr {
        FormulaExpr::CellRef(cell_ref) => {
            let sheet_idx = cell_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);
            ranges.push((
                sheet_idx,
                cell_ref.address.row,
                cell_ref.address.col,
                cell_ref.address.row,
                cell_ref.address.col,
            ));
            true
        }
        FormulaExpr::RangeRef(range_ref) => {
            let sheet_idx = range_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);
            ranges.push((
                sheet_idx,
                range_ref.range.start.row,
                range_ref.range.start.col,
                range_ref.range.end.row,
                range_ref.range.end.col,
            ));
            true
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_value_sensitive_ranges(left, current_sheet, workbook, ranges)
                && collect_value_sensitive_ranges(right, current_sheet, workbook, ranges)
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            collect_value_sensitive_ranges(operand, current_sheet, workbook, ranges)
        }
        FormulaExpr::Function { args, .. } => args
            .iter()
            .all(|arg| collect_value_sensitive_ranges(arg, current_sheet, workbook, ranges)),
        FormulaExpr::Array(rows) => rows.iter().all(|row| {
            row.iter()
                .all(|cell| collect_value_sensitive_ranges(cell, current_sheet, workbook, ranges))
        }),
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::Empty => true,
        FormulaExpr::NameRef(_) | FormulaExpr::ExternalRef(_) | FormulaExpr::StructuredRef(_) => {
            false
        }
    }
}

fn try_static_index_coord(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    formula_cells: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
    range_dep_cache: &mut RangeDependencyCache,
) -> (Option<i64>, Vec<SensitiveRange>) {
    match expr {
        FormulaExpr::Number(n) if n.is_finite() => {
            let i = n.trunc() as i64;
            ((i >= 1).then_some(i), Vec::new())
        }
        FormulaExpr::Function { name, .. }
            if name.eq_ignore_ascii_case("MATCH") || name.eq_ignore_ascii_case("XMATCH") =>
        {
            let mut deps = Vec::new();
            let mut ignored_sensitive = Vec::new();
            extract_references_recursive(
                expr,
                current_sheet,
                workbook,
                formula_cells,
                formula_cell_index,
                range_dep_cache,
                &mut ignored_sensitive,
                &mut deps,
            );
            if !deps.is_empty() {
                return (None, Vec::new());
            }

            let mut sensitive_ranges = Vec::new();
            if !collect_value_sensitive_ranges(expr, current_sheet, workbook, &mut sensitive_ranges)
            {
                return (None, Vec::new());
            }

            let ctx = EvaluationContext::new(Some(workbook), current_sheet, 0, 0);
            let value = match evaluate(expr, &ctx).ok() {
                Some(FormulaValue::Number(n)) if n.is_finite() => {
                    let i = n.trunc() as i64;
                    (i >= 1).then_some(i)
                }
                _ => None,
            };
            (value, sensitive_ranges)
        }
        _ => (None, Vec::new()),
    }
}

fn build_formula_cell_index(formula_cells: &AHashSet<CellKey>) -> FormulaCellIndex {
    let mut index = FormulaCellIndex::new();

    for &cell in formula_cells {
        index
            .entry(cell.sheet)
            .or_default()
            .push((cell.row, cell.col));
    }

    for cells in index.values_mut() {
        cells.sort_unstable();
        cells.dedup();
    }

    index
}

fn collect_input_ranges(
    expr: &FormulaExpr,
    current_sheet: usize,
    workbook: &Workbook,
    formula_cells: &AHashSet<CellKey>,
    ranges: &mut Vec<SensitiveRange>,
) {
    match expr {
        FormulaExpr::CellRef(cell_ref) => {
            let sheet_idx = cell_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);
            let key = CellKey::new(sheet_idx, cell_ref.address.row, cell_ref.address.col);
            if !formula_cells.contains(&key) {
                ranges.push((
                    sheet_idx,
                    cell_ref.address.row,
                    cell_ref.address.col,
                    cell_ref.address.row,
                    cell_ref.address.col,
                ));
            }
        }
        FormulaExpr::RangeRef(range_ref) => {
            let sheet_idx = range_ref
                .sheet
                .as_ref()
                .and_then(|name| workbook.sheet_index(name))
                .unwrap_or(current_sheet);
            ranges.push((
                sheet_idx,
                range_ref.range.start.row,
                range_ref.range.start.col,
                range_ref.range.end.row,
                range_ref.range.end.col,
            ));
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_input_ranges(left, current_sheet, workbook, formula_cells, ranges);
            collect_input_ranges(right, current_sheet, workbook, formula_cells, ranges);
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            collect_input_ranges(operand, current_sheet, workbook, formula_cells, ranges);
        }
        FormulaExpr::Function { args, .. } => {
            for arg in args {
                collect_input_ranges(arg, current_sheet, workbook, formula_cells, ranges);
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    collect_input_ranges(cell, current_sheet, workbook, formula_cells, ranges);
                }
            }
        }
        FormulaExpr::StructuredRef(sr) => {
            collect_structured_ref_input_ranges(sr, current_sheet, workbook, ranges);
        }
        FormulaExpr::NameRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::Empty => {}
    }
}

fn collect_structured_ref_input_ranges(
    sr: &StructuredReference,
    current_sheet: usize,
    workbook: &Workbook,
    ranges: &mut Vec<SensitiveRange>,
) {
    let (reference, columns, header_rows, totals_rows, sheet_idx) = match &sr.table {
        Some(table_name) => {
            let mut found = None;
            for idx in 0..workbook.sheet_count() {
                if let Some(ws) = workbook.worksheet(idx) {
                    if let Some(t) = ws.table_by_name(table_name) {
                        found = Some((
                            t.reference,
                            t.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>(),
                            t.header_row_count,
                            t.totals_row_count,
                            idx,
                        ));
                        break;
                    }
                }
            }
            match found {
                Some(f) => f,
                None => return,
            }
        }
        None => {
            if let Some(ws) = workbook.worksheet(current_sheet) {
                match ws.tables().first() {
                    Some(t) => (
                        t.reference,
                        t.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>(),
                        t.header_row_count,
                        t.totals_row_count,
                        current_sheet,
                    ),
                    None => return,
                }
            } else {
                return;
            }
        }
    };

    let (col_start, col_end) = match &sr.column {
        Some(col_name) => {
            match columns
                .iter()
                .position(|c| c.eq_ignore_ascii_case(col_name))
            {
                Some(i) => {
                    let col = reference.start.col + i as u16;
                    (col, col)
                }
                None => return,
            }
        }
        None => (reference.start.col, reference.end.col),
    };

    let ref_start = reference.start.row;
    let ref_end = reference.end.row;
    let data_start = ref_start + header_rows;
    let data_end = if totals_rows > 0 {
        ref_end - totals_rows
    } else {
        ref_end
    };

    let has_all = sr.specifiers.contains(&StructuredRefSpecifier::All);
    let has_headers = sr.specifiers.contains(&StructuredRefSpecifier::Headers);
    let has_data = sr.specifiers.contains(&StructuredRefSpecifier::Data);
    let has_totals = sr.specifiers.contains(&StructuredRefSpecifier::Totals);
    let has_this_row = sr.specifiers.contains(&StructuredRefSpecifier::ThisRow);

    let (row_start, row_end) = if has_all {
        (ref_start, ref_end)
    } else if has_this_row {
        (data_start, data_end)
    } else if has_headers && has_data && has_totals {
        (ref_start, ref_end)
    } else if has_headers && has_data {
        let end = if has_totals { ref_end } else { data_end };
        (ref_start, end)
    } else if has_data && has_totals {
        (data_start, ref_end)
    } else if has_headers {
        (ref_start, ref_start + header_rows.saturating_sub(1))
    } else if has_totals && totals_rows > 0 {
        (ref_end - totals_rows + 1, ref_end)
    } else {
        (data_start, data_end)
    };

    ranges.push((sheet_idx, row_start, col_start, row_end, col_end));
}

#[cfg(debug_assertions)]
fn debug_validate_eval_plan(
    plan: &EvalPlan,
    parsed_formulas: &AHashMap<CellKey, FormulaExpr>,
    volatile_cells: &AHashSet<CellKey>,
    circular_cells: &AHashSet<CellKey>,
) {
    let n = plan.idx_to_cell.len();
    debug_assert_eq!(plan.depth.len(), n);
    debug_assert_eq!(plan.dependents.len(), n);
    debug_assert_eq!(plan.input_ranges.len(), n);
    debug_assert_eq!(plan.cell_to_idx.len(), n);

    for (idx, &cell) in plan.idx_to_cell.iter().enumerate() {
        debug_assert_eq!(plan.cell_to_idx.get(&cell).copied(), Some(idx as u32));
        debug_assert!(parsed_formulas.contains_key(&cell));
    }

    for deps in &plan.dependents {
        for &dep in deps.iter() {
            debug_assert!((dep as usize) < n);
        }
    }

    for &cell in volatile_cells {
        debug_assert!(plan.cell_to_idx.contains_key(&cell));
    }
    for &cell in circular_cells {
        debug_assert!(plan.cell_to_idx.contains_key(&cell));
    }
}

fn build_dense_dependents(precedents: &DensePrecedents) -> DenseDependents {
    let mut temp: Vec<Vec<u32>> = vec![Vec::new(); precedents.len()];
    for (idx, deps) in precedents.iter().enumerate() {
        for &dep in deps.iter() {
            temp[dep as usize].push(idx as u32);
        }
    }
    temp.into_iter()
        .map(|mut deps| {
            deps.sort_unstable();
            deps.dedup();
            deps.into_boxed_slice()
        })
        .collect()
}

fn collect_dirty_value_ranges(workbook: &Workbook) -> Vec<SensitiveRange> {
    let mut dirty = Vec::new();
    for sheet_idx in 0..workbook.sheet_count() {
        let Some(ws) = workbook.worksheet(sheet_idx) else {
            continue;
        };
        dirty.extend(
            ws.dirty_value_ranges()
                .iter()
                .map(|&(r1, c1, r2, c2)| (sheet_idx, r1, c1, r2, c2)),
        );
    }
    dirty
}

fn collect_affected_formula_flags(plan: &EvalPlan, dirty_ranges: &[SensitiveRange]) -> Vec<bool> {
    let mut affected = vec![false; plan.idx_to_cell.len()];
    let mut stack = Vec::new();

    for (idx, ranges) in plan.input_ranges.iter().enumerate() {
        let overlaps = ranges.iter().any(|&(sheet, r1, c1, r2, c2)| {
            dirty_ranges.iter().any(|&(ds, dr1, dc1, dr2, dc2)| {
                ds == sheet
                    && CalcCache::range_overlaps_dirty(
                        (dr1, dc1, dr2, dc2),
                        (sheet, r1, c1, r2, c2),
                    )
            })
        });
        if overlaps {
            affected[idx] = true;
            stack.push(idx as u32);
        }
    }

    while let Some(idx) = stack.pop() {
        for &dep in plan.dependents[idx as usize].iter() {
            if !affected[dep as usize] {
                affected[dep as usize] = true;
                stack.push(dep);
            }
        }
    }

    affected
}

fn build_dense_precedents(
    seed_order: &[CellKey],
    asts: &[&FormulaExpr],
    workbook: &Workbook,
    formula_cell_set: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
    cell_to_idx: &AHashMap<CellKey, u32>,
    #[allow(unused_variables)] use_parallel: bool,
) -> (DensePrecedents, DenseInputRanges, Vec<SensitiveRange>) {
    #[cfg(feature = "parallel")]
    if use_parallel {
        let per_cell: Vec<(Box<[u32]>, Box<[SensitiveRange]>, Vec<SensitiveRange>)> = seed_order
            .par_iter()
            .enumerate()
            .map_init(
                || {
                    (
                        RangeDependencyCache::new(),
                        Vec::<CellKey>::new(),
                        Vec::<SensitiveRange>::new(),
                        Vec::<SensitiveRange>::new(),
                    )
                },
                |(range_dep_cache, dep_keys, input_ranges, sensitive_ranges), (i, &cell)| {
                    dep_keys.clear();
                    input_ranges.clear();
                    sensitive_ranges.clear();
                    collect_input_ranges(
                        asts[i],
                        cell.sheet,
                        workbook,
                        formula_cell_set,
                        input_ranges,
                    );
                    extract_references_recursive(
                        asts[i],
                        cell.sheet,
                        workbook,
                        formula_cell_set,
                        formula_cell_index,
                        range_dep_cache,
                        sensitive_ranges,
                        dep_keys,
                    );
                    (
                        dep_keys
                            .iter()
                            .filter_map(|dep| cell_to_idx.get(dep).copied())
                            .collect::<Vec<u32>>()
                            .into_boxed_slice(),
                        input_ranges.clone().into_boxed_slice(),
                        sensitive_ranges.clone(),
                    )
                },
            )
            .collect();
        let mut precedents = Vec::with_capacity(per_cell.len());
        let mut input_ranges = Vec::with_capacity(per_cell.len());
        let mut value_sensitive_ranges = Vec::new();
        for (deps, inputs, ranges) in per_cell {
            precedents.push(deps);
            input_ranges.push(inputs);
            value_sensitive_ranges.extend(ranges);
        }
        value_sensitive_ranges.sort_unstable();
        value_sensitive_ranges.dedup();
        return (precedents, input_ranges, value_sensitive_ranges);
    }

    let mut range_dep_cache = RangeDependencyCache::new();
    let mut dep_keys = Vec::new();
    let mut input_ranges = Vec::new();
    let mut value_sensitive_ranges = Vec::new();
    let mut precedents = Vec::with_capacity(seed_order.len());
    let mut all_input_ranges = Vec::with_capacity(seed_order.len());
    for (i, &cell) in seed_order.iter().enumerate() {
        dep_keys.clear();
        input_ranges.clear();
        collect_input_ranges(
            asts[i],
            cell.sheet,
            workbook,
            formula_cell_set,
            &mut input_ranges,
        );
        extract_references_recursive(
            asts[i],
            cell.sheet,
            workbook,
            formula_cell_set,
            formula_cell_index,
            &mut range_dep_cache,
            &mut value_sensitive_ranges,
            &mut dep_keys,
        );
        precedents.push(
            dep_keys
                .iter()
                .filter_map(|dep| cell_to_idx.get(dep).copied())
                .collect::<Vec<u32>>()
                .into_boxed_slice(),
        );
        all_input_ranges.push(input_ranges.clone().into_boxed_slice());
    }
    value_sensitive_ranges.sort_unstable();
    value_sensitive_ranges.dedup();
    (precedents, all_input_ranges, value_sensitive_ranges)
}

#[allow(clippy::too_many_arguments)]
fn push_range_references(
    sheet_idx: usize,
    row_start: u32,
    row_end: u32,
    col_start: u16,
    col_end: u16,
    formula_cells: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
    range_dep_cache: &mut RangeDependencyCache,
    refs: &mut Vec<CellKey>,
) {
    let cache_key = (sheet_idx, row_start, row_end, col_start, col_end);
    if let Some(cached) = range_dep_cache.get(&cache_key) {
        refs.extend(cached.iter().copied());
        return;
    }

    let Some(cells) = formula_cell_index.get(&sheet_idx) else {
        return;
    };

    let start = cells.partition_point(|&(row, _)| row < row_start);
    let end = cells.partition_point(|&(row, _)| row <= row_end);
    if start == end {
        range_dep_cache.insert(cache_key, Vec::new());
        return;
    }

    let mut found = Vec::new();

    for &(row, col) in &cells[start..end] {
        if col >= col_start && col <= col_end {
            let cell_key = CellKey::new(sheet_idx, row, col);
            debug_assert!(formula_cells.contains(&cell_key));
            found.push(cell_key);
        }
    }

    refs.extend(found.iter().copied());
    range_dep_cache.insert(cache_key, found);
}

/// Extract sheet indices referenced by cross-sheet formulas in an AST.
/// Used for discovering transitive sheet dependencies.
fn extract_sheet_refs(expr: &FormulaExpr, workbook: &Workbook) -> HashSet<usize> {
    let mut sheets = HashSet::new();
    extract_sheet_refs_recursive(expr, workbook, &mut sheets);
    sheets
}

fn extract_sheet_refs_recursive(
    expr: &FormulaExpr,
    workbook: &Workbook,
    sheets: &mut HashSet<usize>,
) {
    match expr {
        FormulaExpr::CellRef(cell_ref) => {
            if let Some(name) = &cell_ref.sheet {
                if let Some(idx) = workbook.sheet_index(name) {
                    sheets.insert(idx);
                }
            }
        }
        FormulaExpr::RangeRef(range_ref) => {
            if let Some(name) = &range_ref.sheet {
                if let Some(idx) = workbook.sheet_index(name) {
                    sheets.insert(idx);
                }
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            extract_sheet_refs_recursive(left, workbook, sheets);
            extract_sheet_refs_recursive(right, workbook, sheets);
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            extract_sheet_refs_recursive(operand, workbook, sheets);
        }
        FormulaExpr::Function { args, .. } => {
            for arg in args {
                extract_sheet_refs_recursive(arg, workbook, sheets);
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    extract_sheet_refs_recursive(cell, workbook, sheets);
                }
            }
        }
        FormulaExpr::StructuredRef(sr) => {
            // If the structured ref names a table, find which sheet owns it
            if let Some(table_name) = &sr.table {
                for idx in 0..workbook.sheet_count() {
                    if let Some(ws) = workbook.worksheet(idx) {
                        if ws.table_by_name(table_name).is_some() {
                            sheets.insert(idx);
                            break;
                        }
                    }
                }
            }
        }
        _ => {}
    }
}

/// Extract cell dependencies from a structured table reference.
///
/// Resolves the structured ref to a concrete cell range using the workbook's
/// table definitions and adds all cells in that range to the dependency list.
/// For large ranges, prunes to formula cells only.
#[allow(clippy::too_many_arguments)]
fn extract_structured_ref_deps(
    sr: &StructuredReference,
    current_sheet: usize,
    workbook: &Workbook,
    formula_cells: &AHashSet<CellKey>,
    formula_cell_index: &FormulaCellIndex,
    range_dep_cache: &mut RangeDependencyCache,
    _value_sensitive_ranges: &mut Vec<SensitiveRange>,
    refs: &mut Vec<CellKey>,
) {
    // Find the table and its sheet index.
    let (reference, columns, header_rows, totals_rows, sheet_idx) = match &sr.table {
        Some(table_name) => {
            let mut found = None;
            for idx in 0..workbook.sheet_count() {
                if let Some(ws) = workbook.worksheet(idx) {
                    if let Some(t) = ws.table_by_name(table_name) {
                        found = Some((
                            t.reference,
                            t.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>(),
                            t.header_row_count,
                            t.totals_row_count,
                            idx,
                        ));
                        break;
                    }
                }
            }
            match found {
                Some(f) => f,
                None => return,
            }
        }
        None => {
            if let Some(ws) = workbook.worksheet(current_sheet) {
                match ws.tables().first() {
                    Some(t) => (
                        t.reference,
                        t.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>(),
                        t.header_row_count,
                        t.totals_row_count,
                        current_sheet,
                    ),
                    None => return,
                }
            } else {
                return;
            }
        }
    };

    // Determine column range.
    let (col_start, col_end) = match &sr.column {
        Some(col_name) => {
            match columns
                .iter()
                .position(|c| c.eq_ignore_ascii_case(col_name))
            {
                Some(i) => {
                    let col = reference.start.col + i as u16;
                    (col, col)
                }
                None => return, // column not found
            }
        }
        None => (reference.start.col, reference.end.col),
    };

    // Determine row range based on specifiers.
    let ref_start = reference.start.row;
    let ref_end = reference.end.row;
    let data_start = ref_start + header_rows;
    let data_end = if totals_rows > 0 {
        ref_end - totals_rows
    } else {
        ref_end
    };

    let has_all = sr.specifiers.contains(&StructuredRefSpecifier::All);
    let has_headers = sr.specifiers.contains(&StructuredRefSpecifier::Headers);
    let has_data = sr.specifiers.contains(&StructuredRefSpecifier::Data);
    let has_totals = sr.specifiers.contains(&StructuredRefSpecifier::Totals);
    let has_this_row = sr.specifiers.contains(&StructuredRefSpecifier::ThisRow);

    let (row_start, row_end) = if has_all {
        (ref_start, ref_end)
    } else if has_this_row {
        // ThisRow depends on context; conservatively include all data rows.
        (data_start, data_end)
    } else if has_headers && has_data && has_totals {
        (ref_start, ref_end)
    } else if has_headers && has_data {
        let end = if has_totals { ref_end } else { data_end };
        (ref_start, end)
    } else if has_data && has_totals {
        (data_start, ref_end)
    } else if has_headers {
        (ref_start, ref_start + header_rows.saturating_sub(1))
    } else if has_totals && totals_rows > 0 {
        (ref_end - totals_rows + 1, ref_end)
    } else {
        // Default: #Data (implicit)
        (data_start, data_end)
    };

    push_range_references(
        sheet_idx,
        row_start,
        row_end,
        col_start,
        col_end,
        formula_cells,
        formula_cell_index,
        range_dep_cache,
        refs,
    );
}

/// Check if a formula contains any volatile functions
fn contains_volatile_function(expr: &FormulaExpr) -> bool {
    match expr {
        FormulaExpr::Function { name, args } => {
            // Check if this function is volatile
            let registry = get_function_registry();
            if let Some(func_def) = registry.get(name) {
                if func_def.volatile {
                    return true;
                }
            }
            // Check arguments recursively
            args.iter().any(contains_volatile_function)
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            contains_volatile_function(left) || contains_volatile_function(right)
        }
        FormulaExpr::UnaryOp { operand, .. } => contains_volatile_function(operand),
        FormulaExpr::Array(rows) => rows
            .iter()
            .any(|row| row.iter().any(contains_volatile_function)),
        // These can't contain volatile functions
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellError, ImageInfo, ImageSizing};
    use std::sync::Arc;

    #[test]
    fn test_simple_calculation() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", 10.0).unwrap();
        sheet.set_cell_value("A2", 20.0).unwrap();
        sheet.set_cell_formula("A3", "=A1+A2").unwrap();

        let stats = workbook.calculate().unwrap();

        assert_eq!(stats.formula_count, 1);
        assert_eq!(stats.cells_calculated, 1);
        assert_eq!(stats.errors, 0);

        let sheet = workbook.worksheet(0).unwrap();
        let result = sheet.get_calculated_value_at(2, 0); // A3 is row 2, col 0
        assert_eq!(result, Some(&CellValue::Number(30.0)));
    }

    #[test]
    fn test_chain_calculation() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", 5.0).unwrap();
        sheet.set_cell_formula("A2", "=A1*2").unwrap();
        sheet.set_cell_formula("A3", "=A2+10").unwrap();
        sheet.set_cell_formula("A4", "=A3*A1").unwrap();

        let stats = workbook.calculate().unwrap();

        assert_eq!(stats.formula_count, 3);
        assert_eq!(stats.cells_calculated, 3);

        let sheet = workbook.worksheet(0).unwrap();
        // A2 = 5*2 = 10
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(10.0))
        );
        // A3 = 10+10 = 20
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Number(20.0))
        );
        // A4 = 20*5 = 100
        assert_eq!(
            sheet.get_calculated_value_at(3, 0),
            Some(&CellValue::Number(100.0))
        );
    }

    #[test]
    fn test_sum_range() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", 1.0).unwrap();
        sheet.set_cell_value("A2", 2.0).unwrap();
        sheet.set_cell_value("A3", 3.0).unwrap();
        sheet.set_cell_value("A4", 4.0).unwrap();
        sheet.set_cell_formula("A5", "=SUM(A1:A4)").unwrap();

        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.formula_count, 1);

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(4, 0),
            Some(&CellValue::Number(10.0))
        );
    }

    #[test]
    fn test_circular_reference_detection() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Create a circular reference: A1 = B1, B1 = A1
        sheet.set_cell_formula("A1", "=B1").unwrap();
        sheet.set_cell_formula("B1", "=A1").unwrap();

        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.circular_references, 2, "circular_references");
        // Circular cells evaluate using cached/default values (matching Excel
        // behavior) rather than producing #REF! errors.
        assert_eq!(stats.errors, 0, "errors");
    }

    #[test]
    fn test_iterative_calculation() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Simple iterative calculation: A1 starts at 1, B1 = A1/2 + 0.5
        // This should converge to B1 = 1
        sheet.set_cell_value("A1", 1.0).unwrap();
        sheet.set_cell_formula("B1", "=A1").unwrap();
        sheet.set_cell_formula("A1", "=B1/2+0.5").unwrap();

        let options = CalculationOptions {
            iterative: true,
            max_iterations: 100,
            max_change: 0.0001,
            ..Default::default()
        };

        let stats = workbook.calculate_with_options(&options).unwrap();

        assert!(stats.converged);
    }

    #[test]
    fn test_webservice_callback_via_calculation_options() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("A1", r#"=WEBSERVICE("https://example.com/data")"#)
            .unwrap();

        let options = CalculationOptions {
            web_service_fn: Some(Arc::new(|url| Some(format!("body:{}", url)))),
            ..Default::default()
        };

        let stats = workbook.calculate_with_options(&options).unwrap();
        assert_eq!(stats.errors, 0);

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::String("body:https://example.com/data".into()))
        );
    }

    #[test]
    fn test_rtd_callback_via_calculation_options() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("A1", r#"=RTD("prog","srv","topic1","topic2")"#)
            .unwrap();

        let options = CalculationOptions {
            rtd_fn: Some(Arc::new(|prog_id, server, topics| {
                Some(format!("{}|{}|{}", prog_id, server, topics.join("|")))
            })),
            ..Default::default()
        };

        let stats = workbook.calculate_with_options(&options).unwrap();
        assert_eq!(stats.errors, 0);

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::String("prog|srv|topic1|topic2".into()))
        );
    }

    #[test]
    fn test_image_metadata_on_worksheet() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula(
                "A1",
                r#"=IMAGE("https://example.com/logo.png","Logo",3,48,96)"#,
            )
            .unwrap();

        let _stats = workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::String("Logo".into()))
        );
        assert_eq!(
            sheet.get_image_at(0, 0),
            Some(ImageInfo {
                source: "https://example.com/logo.png".to_string(),
                alt_text: "Logo".to_string(),
                sizing: ImageSizing::Custom,
                width: Some(96.0),
                height: Some(48.0),
            })
        );
    }

    #[test]
    fn test_volatile_function_detection() {
        let ast = parse_formula("=NOW()").unwrap();
        assert!(contains_volatile_function(&ast));

        let ast = parse_formula("=TODAY()").unwrap();
        assert!(contains_volatile_function(&ast));

        let ast = parse_formula("=RAND()").unwrap();
        assert!(contains_volatile_function(&ast));

        let ast = parse_formula("=SUM(A1:A10)").unwrap();
        assert!(!contains_volatile_function(&ast));

        // Nested volatile function
        let ast = parse_formula("=IF(A1>0,NOW(),0)").unwrap();
        assert!(contains_volatile_function(&ast));
    }

    #[test]
    fn test_multiple_sheets() {
        let mut workbook = Workbook::new();

        // First sheet
        let sheet1 = workbook.worksheet_mut(0).unwrap();
        sheet1.set_cell_value("A1", 100.0).unwrap();

        // Add second sheet
        workbook.add_worksheet_with_name("Sheet2").unwrap();
        let sheet2 = workbook.worksheet_mut(1).unwrap();
        sheet2.set_cell_value("A1", 50.0).unwrap();
        sheet2.set_cell_formula("A2", "=Sheet1!A1+A1").unwrap();

        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.formula_count, 1);

        let sheet2 = workbook.worksheet(1).unwrap();
        // Should be 100 + 50 = 150
        assert_eq!(
            sheet2.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(150.0))
        );
    }

    #[test]
    fn test_extract_references() {
        let workbook = Workbook::new();

        // Build a formula_cells set containing the cells we expect to reference.
        // extract_references only returns deps on formula cells (static cells
        // are skipped since they never change during calculation).
        let mut formula_cells: AHashSet<CellKey> = AHashSet::new();
        // A1, A2, A3 for the range test; B2, C3 for the multi-ref test
        formula_cells.insert(CellKey::new(0, 0, 0)); // A1
        formula_cells.insert(CellKey::new(0, 1, 0)); // A2
        formula_cells.insert(CellKey::new(0, 2, 0)); // A3
        formula_cells.insert(CellKey::new(0, 1, 1)); // B2
        formula_cells.insert(CellKey::new(0, 2, 2)); // C3
        let index = build_formula_cell_index(&formula_cells);

        // Simple cell reference
        let ast = parse_formula("=A1").unwrap();
        let refs = extract_references(&ast, 0, &workbook, &formula_cells, &index);
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0], CellKey::new(0, 0, 0));

        // Range reference
        let ast = parse_formula("=SUM(A1:A3)").unwrap();
        let refs = extract_references(&ast, 0, &workbook, &formula_cells, &index);
        assert_eq!(refs.len(), 3);

        // Multiple references
        let ast = parse_formula("=A1+B2*C3").unwrap();
        let refs = extract_references(&ast, 0, &workbook, &formula_cells, &index);
        assert_eq!(refs.len(), 3);

        // Reference to a non-formula cell returns nothing
        let ast = parse_formula("=D4").unwrap();
        let refs = extract_references(&ast, 0, &workbook, &formula_cells, &index);
        assert_eq!(refs.len(), 0);
    }

    #[test]
    fn test_sequence_spilling() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Set up a SEQUENCE formula in A1 that should spill to A1:A5
        sheet.set_cell_formula("A1", "=SEQUENCE(5)").unwrap();

        // Calculate the workbook
        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.formula_count, 1);
        assert_eq!(stats.errors, 0);

        let sheet = workbook.worksheet(0).unwrap();

        // A1 should have the array result stored with value 1.0
        let a1 = sheet.get_calculated_value_at(0, 0);
        assert_eq!(a1, Some(&CellValue::Number(1.0)));

        // A2-A5 should have SpillTarget values that resolve to 2.0, 3.0, 4.0, 5.0
        let a2 = sheet.get_calculated_value_at(1, 0);
        match a2 {
            Some(CellValue::SpillTarget {
                source_row,
                source_col,
                ..
            }) => {
                assert_eq!(*source_row, 0);
                assert_eq!(*source_col, 0);
            }
            Some(CellValue::Number(n)) => {
                assert_eq!(*n, 2.0); // Direct value is also acceptable
            }
            other => panic!("Expected SpillTarget or Number for A2, got {:?}", other),
        }

        // Check A5 (last spilled cell)
        let a5 = sheet.get_calculated_value_at(4, 0);
        match a5 {
            Some(CellValue::SpillTarget {
                source_row,
                source_col,
                ..
            }) => {
                assert_eq!(*source_row, 0);
                assert_eq!(*source_col, 0);
            }
            Some(CellValue::Number(n)) => {
                assert_eq!(*n, 5.0); // Direct value is also acceptable
            }
            other => panic!("Expected SpillTarget or Number for A5, got {:?}", other),
        }

        // A6 should be empty (spill doesn't go past 5 rows)
        let a6 = sheet.get_calculated_value_at(5, 0);
        assert!(a6.is_none() || matches!(a6, Some(CellValue::Empty)));
    }

    #[test]
    fn test_sequence_2d_spilling() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Set up a SEQUENCE formula in A1 that should spill to a 3x4 grid
        sheet.set_cell_formula("A1", "=SEQUENCE(3, 4)").unwrap();

        // Calculate the workbook
        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.formula_count, 1);
        assert_eq!(stats.errors, 0);

        let sheet = workbook.worksheet(0).unwrap();

        // A1 should be the source with value 1.0
        let a1 = sheet.get_calculated_value_at(0, 0);
        assert_eq!(a1, Some(&CellValue::Number(1.0)));

        // D1 (row 0, col 3) should be 4.0
        let d1 = sheet.get_calculated_value_at(0, 3);
        match d1 {
            Some(CellValue::SpillTarget { .. }) | Some(CellValue::Number(4.0)) => {}
            other => panic!("Expected value 4.0 for D1, got {:?}", other),
        }

        // A3 (row 2, col 0) should be 9.0
        let a3 = sheet.get_calculated_value_at(2, 0);
        match a3 {
            Some(CellValue::SpillTarget { .. }) | Some(CellValue::Number(9.0)) => {}
            other => panic!("Expected value 9.0 for A3, got {:?}", other),
        }

        // D3 (row 2, col 3) should be 12.0 (last cell)
        let d3 = sheet.get_calculated_value_at(2, 3);
        match d3 {
            Some(CellValue::SpillTarget { .. }) | Some(CellValue::Number(12.0)) => {}
            other => panic!("Expected value 12.0 for D3, got {:?}", other),
        }
    }

    #[test]
    fn test_sequence_spill_blocked() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Put a value in A3 that will block the spill
        sheet.set_cell_value("A3", 999.0).unwrap();

        // Set up a SEQUENCE formula in A1 that needs to spill to A1:A5
        // This should be blocked because A3 has data
        sheet.set_cell_formula("A1", "=SEQUENCE(5)").unwrap();

        // Calculate the workbook
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();

        // A1 should have a #SPILL! error because the range is blocked
        let a1 = sheet.get_calculated_value_at(0, 0);
        match a1 {
            Some(CellValue::Error(duke_sheets_core::CellError::Spill)) => {
                // Expected - spill was blocked
            }
            Some(CellValue::Number(1.0)) => {
                // Alternative: implementation may allow partial spill or overwrite
                // This depends on the exact implementation
            }
            _ => {
                // For now, accept either error or the value (implementation detail)
            }
        }

        // A3 should still have its original value (not overwritten)
        let a3 = sheet.get_calculated_value_at(2, 0);
        assert_eq!(a3, Some(&CellValue::Number(999.0)));
    }

    /// Helper: create a workbook with a table "Sales" over A1:C6 (header + 4 data + 1 totals).
    /// Columns: Product, Region, Revenue. Data rows 2-5, totals row 6.
    fn workbook_with_sales_table(with_totals: bool) -> Workbook {
        use duke_sheets_core::{CellRange, Table, TableColumn};

        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Header row (row 0)
        sheet.set_cell_value("A1", "Product").unwrap();
        sheet.set_cell_value("B1", "Region").unwrap();
        sheet.set_cell_value("C1", "Revenue").unwrap();

        // Data rows (rows 1-4)
        sheet.set_cell_value("A2", "Widget").unwrap();
        sheet.set_cell_value("B2", "East").unwrap();
        sheet.set_cell_value("C2", 100.0).unwrap();

        sheet.set_cell_value("A3", "Gadget").unwrap();
        sheet.set_cell_value("B3", "West").unwrap();
        sheet.set_cell_value("C3", 200.0).unwrap();

        sheet.set_cell_value("A4", "Widget").unwrap();
        sheet.set_cell_value("B4", "East").unwrap();
        sheet.set_cell_value("C4", 300.0).unwrap();

        sheet.set_cell_value("A5", "Gadget").unwrap();
        sheet.set_cell_value("B5", "West").unwrap();
        sheet.set_cell_value("C5", 400.0).unwrap();

        let totals_count = if with_totals { 1 } else { 0 };
        let end_row = if with_totals { "C6" } else { "C5" };

        if with_totals {
            // Totals row (row 5)
            sheet.set_cell_value("A6", "Total").unwrap();
            sheet.set_cell_value("C6", 1000.0).unwrap();
        }

        let reference = CellRange::parse(&format!("A1:{}", end_row)).unwrap();
        let table = Table {
            id: 1,
            name: "Sales".into(),
            display_name: "Sales".into(),
            reference,
            columns: vec![
                TableColumn::new(1, "Product"),
                TableColumn::new(2, "Region"),
                TableColumn::new(3, "Revenue"),
            ],
            style_info: None,
            header_row_count: 1,
            totals_row_count: totals_count,
            totals_row_shown: true,
        };
        sheet.add_table(table);
        workbook
    }

    #[test]
    fn test_structured_ref_sum_column() {
        // =SUM(Sales[Revenue]) should sum the data column (100+200+300+400 = 1000)
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=SUM(Sales[Revenue])")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4), // E1
            Some(&CellValue::Number(1000.0))
        );
    }

    #[test]
    fn test_structured_ref_sum_column_with_totals() {
        // With totals row, =SUM(Sales[Revenue]) should still only sum data rows
        let mut workbook = workbook_with_sales_table(true);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=SUM(Sales[Revenue])")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4), // E1
            Some(&CellValue::Number(1000.0))
        );
    }

    #[test]
    fn test_structured_ref_this_row() {
        // =Sales[@Revenue] in a data row should return the Revenue value for that row
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        // Put formula in D2 (row 1, col 3) - inside the table's row range
        sheet.set_cell_formula("D2", "=Sales[@Revenue]").unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        // D2 should get the Revenue value from row 1 (C2 = 100)
        assert_eq!(
            sheet.get_calculated_value_at(1, 3), // D2
            Some(&CellValue::Number(100.0))
        );
    }

    #[test]
    fn test_structured_ref_headers() {
        // =Sales[[#Headers],[Revenue]] should return the header text "Revenue"
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=Sales[[#Headers],[Revenue]]")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4), // E1
            Some(&CellValue::String("Revenue".into()))
        );
    }

    #[test]
    fn test_structured_ref_totals() {
        // =Sales[[#Totals],[Revenue]] should return the totals row value
        let mut workbook = workbook_with_sales_table(true);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=Sales[[#Totals],[Revenue]]")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4), // E1
            Some(&CellValue::Number(1000.0))
        );
    }

    #[test]
    fn test_structured_ref_all() {
        // =COUNTA(Sales[#All]) should count all cells in the table (header + 4 data rows)
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=COUNTA(Sales[#All])")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        // 5 rows × 3 cols = 15 cells, all non-empty
        assert_eq!(
            sheet.get_calculated_value_at(0, 4), // E1
            Some(&CellValue::Number(15.0))
        );
    }

    #[test]
    fn test_structured_ref_table_not_found() {
        // Reference to non-existent table should produce an error
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=SUM(NoSuchTable[Revenue])")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        let val = sheet.get_calculated_value_at(0, 4);
        // Should be an error (either #REF! or #NAME?)
        match val {
            Some(CellValue::Error(_)) => {} // expected
            other => panic!("Expected error for missing table, got {:?}", other),
        }
    }

    #[test]
    fn test_structured_ref_column_not_found() {
        // Reference to non-existent column should produce an error
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=SUM(Sales[NoSuchCol])")
            .unwrap();

        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        let val = sheet.get_calculated_value_at(0, 4);
        match val {
            Some(CellValue::Error(_)) => {} // expected
            other => panic!("Expected error for missing column, got {:?}", other),
        }
    }

    #[test]
    fn test_structured_ref_dependency_tracking() {
        // Changing a cell in the table should cause formulas referencing
        // the table via structured refs to recalculate correctly.
        let mut workbook = workbook_with_sales_table(false);
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("E1", "=SUM(Sales[Revenue])")
            .unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4),
            Some(&CellValue::Number(1000.0))
        );

        // Change C2 from 100 to 500
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("C2", 500.0).unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 4),
            Some(&CellValue::Number(1400.0)) // 500+200+300+400
        );
    }

    #[test]
    fn test_spill_value_resolution_get_value_at() {
        // get_value_at resolves SpillTarget cells, but anchor cell is still Formula.
        // Use get_calculated_value_at for anchor, get_value_at for spill targets.
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Anchor cell (Formula): get_calculated_value_at resolves cached_value
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(1.0))
        );
        // SpillTarget cells: get_value_at resolves to actual value
        assert_eq!(sheet.get_value_at(1, 0), CellValue::Number(2.0));
        assert_eq!(sheet.get_value_at(2, 0), CellValue::Number(3.0));
        assert_eq!(sheet.get_value_at(3, 0), CellValue::Empty);
    }

    #[test]
    fn test_spill_value_resolution_get_calculated_value_at() {
        // get_calculated_value_at() should also resolve SpillTarget transparently
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(4)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(2.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Number(3.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(3, 0),
            Some(&CellValue::Number(4.0))
        );
    }

    #[test]
    fn test_spill_2d_value_resolution() {
        // 2D spill: SEQUENCE(2,3) should produce a 2x3 grid
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(2,3)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Row 0: 1, 2, 3 (A1 is anchor, B1/C1 are spill targets)
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(sheet.get_value_at(0, 1), CellValue::Number(2.0));
        assert_eq!(sheet.get_value_at(0, 2), CellValue::Number(3.0));
        // Row 1: 4, 5, 6
        assert_eq!(sheet.get_value_at(1, 0), CellValue::Number(4.0));
        assert_eq!(sheet.get_value_at(1, 1), CellValue::Number(5.0));
        assert_eq!(sheet.get_value_at(1, 2), CellValue::Number(6.0));
        // Row 2: empty (no spill)
        assert_eq!(sheet.get_value_at(2, 0), CellValue::Empty);
    }

    #[test]
    fn test_formula_references_spill_target() {
        // A formula in B1 referencing A2 (a spill target) should get the resolved value
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(5)").unwrap();
        // B1 references A3 which is a spill target (value 3.0)
        sheet.set_cell_formula("B1", "=A3*10").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(30.0))
        );
    }

    #[test]
    fn test_sum_over_spill_range() {
        // SUM over a range that includes spill targets should sum the resolved values
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(5)").unwrap();
        sheet.set_cell_formula("B1", "=SUM(A1:A5)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // SUM(1+2+3+4+5) = 15
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(15.0))
        );
    }

    #[test]
    fn test_spill_blocked_returns_spill_error() {
        // When spill range is blocked, anchor cell should show #SPILL!
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A2", 999.0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Error(duke_sheets_core::CellError::Spill))
        );
        // Original value should be preserved
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(999.0))
        );
    }

    #[test]
    fn test_spill_range_operator() {
        // A1# should return the full spill range as an array
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(4)").unwrap();
        // SUM(A1#) should sum the entire spill range
        sheet.set_cell_formula("B1", "=SUM(A1#)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // SUM(1+2+3+4) = 10
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(10.0))
        );
    }

    #[test]
    fn test_spill_range_operator_non_spill_source() {
        // A1# on a cell that is NOT a spill source returns just A1's value
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", 42.0).unwrap();
        sheet.set_cell_formula("B1", "=SUM(A1#)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(42.0))
        );
    }

    #[test]
    fn test_spill_single_cell_no_spill_targets() {
        // Single-cell array result (1x1) should not create any SpillTarget cells
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(1,1)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(1.0))
        );
        assert!(!sheet.is_spill_source(0, 0));
        assert_eq!(sheet.get_value_at(1, 0), CellValue::Empty);
    }

    #[test]
    fn test_spill_is_spill_source_after_calculate() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert!(sheet.is_spill_source(0, 0));
        assert!(sheet.is_spill_target(1, 0));
        assert!(sheet.is_spill_target(2, 0));
        assert!(!sheet.is_spill_target(3, 0));
    }

    #[test]
    fn test_spill_clear_on_recalculate_to_scalar() {
        // If a formula's result changes from array to scalar, old spill targets
        // should be cleared.
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // First: SEQUENCE(3) spills to A1:A3
        sheet.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();
        workbook.calculate().unwrap();
        {
            let sheet = workbook.worksheet(0).unwrap();
            assert!(sheet.is_spill_source(0, 0));
            assert_eq!(sheet.get_value_at(2, 0), CellValue::Number(3.0));
        }

        // Change formula to a scalar
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=42").unwrap();
        workbook.calculate().unwrap();
        {
            let sheet = workbook.worksheet(0).unwrap();
            assert_eq!(
                sheet.get_calculated_value_at(0, 0),
                Some(&CellValue::Number(42.0))
            );
            // Old spill targets should be gone
            assert!(!sheet.is_spill_target(1, 0));
            assert!(!sheet.is_spill_target(2, 0));
            assert_eq!(sheet.get_value_at(1, 0), CellValue::Empty);
        }
    }

    #[test]
    fn test_spill_filter_function() {
        // Test FILTER producing a dynamic-sized spill
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        // Data in A1:A5
        sheet.set_cell_value("A1", 10.0).unwrap();
        sheet.set_cell_value("A2", 20.0).unwrap();
        sheet.set_cell_value("A3", 30.0).unwrap();
        sheet.set_cell_value("A4", 40.0).unwrap();
        sheet.set_cell_value("A5", 50.0).unwrap();

        // Criteria in B1:B5 (TRUE for values >= 30)
        sheet.set_cell_value("B1", false).unwrap();
        sheet.set_cell_value("B2", false).unwrap();
        sheet.set_cell_value("B3", true).unwrap();
        sheet.set_cell_value("B4", true).unwrap();
        sheet.set_cell_value("B5", true).unwrap();

        // FILTER: keep values where criteria is TRUE
        sheet
            .set_cell_formula("C1", "=FILTER(A1:A5,B1:B5)")
            .unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Should spill 30, 40, 50 into C1:C3
        assert_eq!(
            sheet.get_calculated_value_at(0, 2),
            Some(&CellValue::Number(30.0))
        );
        assert_eq!(sheet.get_value_at(1, 2), CellValue::Number(40.0));
        assert_eq!(sheet.get_value_at(2, 2), CellValue::Number(50.0));
        assert_eq!(sheet.get_value_at(3, 2), CellValue::Empty);
    }

    #[test]
    fn test_spill_sort_function() {
        // Test SORT producing a spill
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", 30.0).unwrap();
        sheet.set_cell_value("A2", 10.0).unwrap();
        sheet.set_cell_value("A3", 50.0).unwrap();
        sheet.set_cell_value("A4", 20.0).unwrap();

        sheet.set_cell_formula("B1", "=SORT(A1:A4)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(10.0))
        );
        assert_eq!(sheet.get_value_at(1, 1), CellValue::Number(20.0));
        assert_eq!(sheet.get_value_at(2, 1), CellValue::Number(30.0));
        assert_eq!(sheet.get_value_at(3, 1), CellValue::Number(50.0));
    }

    #[test]
    fn test_spill_unique_function() {
        // Test UNIQUE deduplication
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", "apple").unwrap();
        sheet.set_cell_value("A2", "banana").unwrap();
        sheet.set_cell_value("A3", "apple").unwrap();
        sheet.set_cell_value("A4", "cherry").unwrap();
        sheet.set_cell_value("A5", "banana").unwrap();

        sheet.set_cell_formula("B1", "=UNIQUE(A1:A5)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Should produce: apple, banana, cherry (3 unique values)
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::String("apple".into()))
        );
        assert_eq!(sheet.get_value_at(1, 1), CellValue::String("banana".into()));
        assert_eq!(sheet.get_value_at(2, 1), CellValue::String("cherry".into()));
        assert_eq!(sheet.get_value_at(3, 1), CellValue::Empty);
    }

    #[test]
    fn test_spill_transpose_function() {
        // TRANSPOSE a column to a row
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", 1.0).unwrap();
        sheet.set_cell_value("A2", 2.0).unwrap();
        sheet.set_cell_value("A3", 3.0).unwrap();

        sheet.set_cell_formula("B1", "=TRANSPOSE(A1:A3)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        // Should spill horizontally: B1=1, C1=2, D1=3
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(sheet.get_value_at(0, 2), CellValue::Number(2.0));
        assert_eq!(sheet.get_value_at(0, 3), CellValue::Number(3.0));
        // B2 should be empty (transposed is 1 row)
        assert_eq!(sheet.get_value_at(1, 1), CellValue::Empty);
    }

    #[test]
    fn test_spill_range_2d() {
        // A1# on a 2D spill should return the full 2D array
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(2,3)").unwrap();
        // SUM of all 6 values: 1+2+3+4+5+6 = 21
        sheet.set_cell_formula("E1", "=SUM(A1#)").unwrap();

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();

        assert_eq!(
            sheet.get_calculated_value_at(0, 4),
            Some(&CellValue::Number(21.0))
        );
    }

    #[test]
    fn test_array_comparison_greater_than() {
        // =SEQUENCE(4)>2 should produce {FALSE,FALSE,TRUE,TRUE}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(4)>2").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Boolean(false))
        ); // 1>2
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Boolean(false))
        ); // 2>2
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Boolean(true))
        ); // 3>2
        assert_eq!(
            sheet.get_calculated_value_at(3, 0),
            Some(&CellValue::Boolean(true))
        ); // 4>2
        assert_eq!(sheet.get_calculated_value_at(4, 0), None); // no spill beyond
    }

    #[test]
    fn test_array_comparison_equal() {
        // =SEQUENCE(3)=2 should produce {FALSE,TRUE,FALSE}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)=2").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Boolean(false))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Boolean(true))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Boolean(false))
        );
    }

    #[test]
    fn test_array_comparison_less_than() {
        // =SEQUENCE(3)<3 should produce {TRUE,TRUE,FALSE}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)<3").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Boolean(true))
        ); // 1<3
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Boolean(true))
        ); // 2<3
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Boolean(false))
        ); // 3<3
    }

    #[test]
    fn test_array_arithmetic_add() {
        // =SEQUENCE(3)+10 should produce {11,12,13}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)+10").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(11.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(12.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Number(13.0))
        );
    }

    #[test]
    fn test_array_arithmetic_multiply() {
        // =SEQUENCE(3)*2 should produce {2,4,6}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)*2").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(2.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(4.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Number(6.0))
        );
    }

    #[test]
    fn test_array_arithmetic_2d() {
        // =SEQUENCE(2,3)*10 should produce {{10,20,30},{40,50,60}}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(2,3)*10").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(10.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(20.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(0, 2),
            Some(&CellValue::Number(30.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(40.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 1),
            Some(&CellValue::Number(50.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 2),
            Some(&CellValue::Number(60.0))
        );
    }

    #[test]
    fn test_array_arithmetic_divide_with_zero() {
        // =SEQUENCE(3)/B1 where B1=0 should produce {#DIV/0!,#DIV/0!,#DIV/0!}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("B1", 0.0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)/B1").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Error(CellError::Div0))
        );
    }

    #[test]
    fn test_array_concat() {
        // =SEQUENCE(3)&"x" should produce {"1x","2x","3x"}
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)&\"x\"").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::String("1x".into()))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::String("2x".into()))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::String("3x".into()))
        );
    }

    #[test]
    fn test_array_comparison_chained_with_sum() {
        // =SUM((SEQUENCE(5)>2)*1) should count values >2 = 3
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("A1", "=SUM((SEQUENCE(5)>2)*1)")
            .unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        // {F,F,T,T,T} * 1 = {0,0,1,1,1}, SUM = 3
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(3.0))
        );
    }

    #[test]
    fn test_array_cross_reference_with_arithmetic() {
        // A1: =SEQUENCE(3)  → {1,2,3}
        // B1: =A1:A3*10     - but this is a range × scalar, not array lifting
        // Instead test: B1: =SEQUENCE(3)*10 and C1: =A1+B1 (spill + spill addition)
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=SEQUENCE(3)").unwrap();
        sheet.set_cell_formula("B1", "=SEQUENCE(3)*10").unwrap();
        workbook.calculate().unwrap();

        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(2.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 0),
            Some(&CellValue::Number(3.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(0, 1),
            Some(&CellValue::Number(10.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(1, 1),
            Some(&CellValue::Number(20.0))
        );
        assert_eq!(
            sheet.get_calculated_value_at(2, 1),
            Some(&CellValue::Number(30.0))
        );
    }

    #[test]
    fn test_calc_cache_survives_value_only_edit() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", 2.0).unwrap();
        sheet.set_cell_formula("A2", "=A1*2").unwrap();

        workbook.calculate().unwrap();

        let cache = workbook
            .take_calc_cache()
            .unwrap()
            .downcast::<CalcCache>()
            .ok()
            .unwrap();
        assert!(cache.is_valid(&workbook, &[]));
        workbook.set_calc_cache(cache);

        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value("A1", 3.0)
            .unwrap();

        let cache = workbook
            .take_calc_cache()
            .unwrap()
            .downcast::<CalcCache>()
            .ok()
            .unwrap();
        assert!(cache.is_valid(&workbook, &[]));
        workbook.set_calc_cache(cache);

        workbook.calculate().unwrap();
        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(1, 0),
            Some(&CellValue::Number(6.0))
        );
    }

    #[test]
    fn test_calc_cache_invalidates_on_formula_edit() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", 2.0).unwrap();
        sheet.set_cell_formula("A2", "=A1*2").unwrap();

        workbook.calculate().unwrap();
        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_formula("A2", "=A1*3")
            .unwrap();

        let cache = workbook
            .take_calc_cache()
            .unwrap()
            .downcast::<CalcCache>()
            .ok()
            .unwrap();
        assert!(!cache.is_valid(&workbook, &[]));
    }

    #[test]
    fn test_calc_cache_invalidates_when_value_edit_hits_sensitive_match_range() {
        let mut workbook = Workbook::new();
        let lookup = workbook.add_worksheet_with_name("Lookup").unwrap();

        {
            let sheet = workbook.worksheet_mut(lookup).unwrap();
            sheet.set_cell_value("A1", "Value").unwrap();
            sheet.set_cell_value("B1", "Key").unwrap();
            sheet.set_cell_value("A2", 10.0).unwrap();
            sheet.set_cell_value("A3", 20.0).unwrap();
            sheet.set_cell_value("B2", "A").unwrap();
            sheet.set_cell_value("B3", "B").unwrap();
        }

        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet.set_cell_value("A1", "A").unwrap();
            sheet
                .set_cell_formula(
                    "B1",
                    "=INDEX(Lookup!$A$2:$A$3,MATCH($A$1,Lookup!$B$2:$B$3,0),1)",
                )
                .unwrap();
        }

        workbook.calculate().unwrap();

        workbook
            .worksheet_mut(lookup)
            .unwrap()
            .set_cell_value("B3", "C")
            .unwrap();

        let cache = workbook
            .take_calc_cache()
            .unwrap()
            .downcast::<CalcCache>()
            .ok()
            .unwrap();
        assert!(!cache.is_valid(&workbook, &[]));
    }

    #[test]
    fn test_volatile_formula_recalculates_on_cache_hit_with_unrelated_dirty_value() {
        let mut workbook = Workbook::new();
        {
            let sheet = workbook.worksheet_mut(0).unwrap();
            sheet.set_cell_value("A1", 2.0).unwrap();
            sheet.set_cell_formula("B1", "=OFFSET(A1,0,0)").unwrap();
            sheet.set_cell_value("D1", 1.0).unwrap();
        }

        workbook.calculate().unwrap();

        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value("D1", 2.0)
            .unwrap();
        let stats = workbook.calculate().unwrap();
        assert_eq!(stats.cells_calculated, 1);
        assert_eq!(
            workbook.worksheet(0).unwrap().get_calculated_value_at(0, 1),
            Some(&CellValue::Number(2.0))
        );
    }

    fn assert_cached_vs_full_equivalence(wb_cached: &Workbook, wb_full: &Workbook) {
        assert_eq!(wb_cached.sheet_count(), wb_full.sheet_count());
        for i in 0..wb_cached.sheet_count() {
            let cached_ws = wb_cached.worksheet(i).unwrap();
            let full_ws = wb_full.worksheet(i).unwrap();
            for (row, col, _) in cached_ws.formula_cells() {
                let cached_val = cached_ws.get_calculated_value_at(row, col);
                let full_val = full_ws.get_calculated_value_at(row, col);
                assert_eq!(
                    cached_val, full_val,
                    "mismatch at sheet {} ({},{}) cached={:?} full={:?}",
                    i, row, col, cached_val, full_val
                );
            }
        }
    }

    #[test]
    fn test_cached_vs_full_equivalence_after_value_edit() {
        let mut wb1 = Workbook::new();
        let mut wb2 = Workbook::new();
        for wb in [&mut wb1, &mut wb2] {
            let sheet = wb.worksheet_mut(0).unwrap();
            sheet.set_cell_value("A1", 10.0).unwrap();
            sheet.set_cell_value("A2", 20.0).unwrap();
            sheet.set_cell_formula("B1", "=A1*2").unwrap();
            sheet.set_cell_formula("B2", "=A2+B1").unwrap();
            sheet.set_cell_formula("C1", "=SUM(B1:B2)").unwrap();
        }

        wb1.calculate().unwrap();
        wb2.calculate().unwrap();

        for wb in [&mut wb1, &mut wb2] {
            wb.worksheet_mut(0)
                .unwrap()
                .set_cell_value("A1", 99.0)
                .unwrap();
        }

        wb1.calculate().unwrap();

        wb2.calculate_with_options(&CalculationOptions {
            force_full_calculation: true,
            ..Default::default()
        })
        .unwrap();

        assert_cached_vs_full_equivalence(&wb1, &wb2);
    }

    #[test]
    fn test_cached_vs_full_equivalence_cross_sheet() {
        let mut wb1 = Workbook::new();
        let mut wb2 = Workbook::new();
        for wb in [&mut wb1, &mut wb2] {
            let _ = wb.add_worksheet_with_name("Data");
            {
                let data = wb.worksheet_mut(1).unwrap();
                data.set_cell_value("A1", 100.0).unwrap();
                data.set_cell_value("A2", 200.0).unwrap();
            }
            {
                let calc = wb.worksheet_mut(0).unwrap();
                calc.set_cell_formula("A1", "=Data!A1+Data!A2").unwrap();
                calc.set_cell_formula("A2", "=A1*2").unwrap();
            }
        }

        wb1.calculate().unwrap();
        wb2.calculate().unwrap();

        for wb in [&mut wb1, &mut wb2] {
            wb.worksheet_mut(1)
                .unwrap()
                .set_cell_value("A1", 999.0)
                .unwrap();
        }

        wb1.calculate().unwrap();

        wb2.calculate_with_options(&CalculationOptions {
            force_full_calculation: true,
            ..Default::default()
        })
        .unwrap();

        assert_cached_vs_full_equivalence(&wb1, &wb2);
    }

    #[test]
    fn test_cross_sheet_constants_many_rows() {
        // Regression: cross-sheet formulas referencing constants in later
        // rows sometimes evaluate to 0 due to non-deterministic ordering.
        let mut wb = Workbook::new();
        let data = wb.worksheet_mut(0).unwrap();
        data.set_name("Data");
        for r in 0u32..60 {
            data.set_cell_value_at(r, 0, CellValue::Number((r + 1) as f64))
                .unwrap();
        }
        // Also a string and boolean for type-checking formulas
        data.set_cell_value_at(50, 0, CellValue::String("hello".into()))
            .unwrap();
        data.set_cell_value_at(51, 0, CellValue::Boolean(true))
            .unwrap();

        wb.add_worksheet_with_name("Tests").unwrap();
        let tests = wb.worksheet_mut(1).unwrap();
        // 60 cross-sheet addition formulas
        for r in 0u32..50 {
            tests
                .set_cell_formula_at(r, 0, &format!("=Data!A{}", r + 1))
                .unwrap();
        }
        tests
            .set_cell_formula_at(50, 0, "=ISTEXT(Data!A51)")
            .unwrap();
        tests
            .set_cell_formula_at(51, 0, "=ISLOGICAL(Data!A52)")
            .unwrap();

        wb.calculate().unwrap();

        let tests = wb.worksheet(1).unwrap();
        for r in 0u32..50 {
            let expected = (r + 1) as f64;
            let val = tests.get_value_at(r, 0);
            assert_eq!(
                val,
                CellValue::Number(expected),
                "Row {r}: Data!A{} should be {expected}",
                r + 1
            );
        }
        assert_eq!(
            tests.get_value_at(50, 0),
            CellValue::Boolean(true),
            "ISTEXT(Data!A51) should be true"
        );
        assert_eq!(
            tests.get_value_at(51, 0),
            CellValue::Boolean(true),
            "ISLOGICAL(Data!A52) should be true"
        );
    }

    #[test]
    fn test_multibyte_chars_in_formula_no_panic() {
        let formula = "=VLOOKUP(B53,[\u{0002}JohnDeere\u{0003}JD Kernersville\u{0003}JDK零件清单.xls]美元!$B$2:$E$244,4,0)";
        let delta = min_relative_row_delta(formula, 52);
        assert!(delta <= 0);
        let shifted = shift_a1_references_rows(formula, 0);
        assert_eq!(shifted, formula);
        let shifted = shift_a1_references_rows(formula, 1);
        assert!(shifted.contains("零件清单"));
        assert!(shifted.contains("美元"));
    }

    #[test]
    fn test_multibyte_formula_calculate_no_panic() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=LEN(\"零件清单\")").unwrap();
        sheet.set_cell_formula("A2", "=LEN(\"零件清单\")").unwrap();
        let stats = wb.calculate().unwrap();
        assert_eq!(stats.errors, 0);
        let sheet = wb.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_calculated_value_at(0, 0),
            Some(&CellValue::Number(4.0))
        );
    }
}
