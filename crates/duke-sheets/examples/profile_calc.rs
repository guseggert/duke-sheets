use duke_sheets::prelude::*;
use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use duke_sheets_formula::dependency::CellKey;
use duke_sheets_formula::{parse_formula, FormulaExpr};
use serde_json::json;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::time::Instant;

#[path = "../perf_fixtures.rs"]
mod perf_fixtures;

enum SourceSpec {
    File(String),
    Fixture(String),
}

struct Cli {
    source: SourceSpec,
    serial: bool,
    json: bool,
    once: bool,
    open_only: bool,
    parallel_report: bool,
    sheet: Option<usize>,
}

struct ParallelReport {
    analyzed_formulas: usize,
    dependency_edges: usize,
    max_depth: usize,
    widest_level_depth: usize,
    widest_level_size: usize,
    top_levels: Vec<(usize, usize)>,
    parse_failures: usize,
    cycle_hits: usize,
    analysis_ms: f64,
}

struct ParallelAnalyzer<'a> {
    workbook: &'a Workbook,
    texts: HashMap<CellKey, String>,
    formula_index: HashMap<usize, BTreeMap<u32, Vec<u16>>>,
    selected_sheet: Option<usize>,
    parsed: HashMap<CellKey, Option<FormulaExpr>>,
    deps: HashMap<CellKey, Vec<CellKey>>,
    depths: HashMap<CellKey, usize>,
    visiting: HashSet<CellKey>,
    parse_failures: usize,
    cycle_hits: usize,
}

fn build_fixture(name: &str) -> Workbook {
    perf_fixtures::build_fixture(name)
}

fn source_label(source: &SourceSpec) -> String {
    match source {
        SourceSpec::File(path) => path.clone(),
        SourceSpec::Fixture(name) => format!("fixture:{name}"),
    }
}

impl<'a> ParallelAnalyzer<'a> {
    fn new(workbook: &'a Workbook, selected_sheet: Option<usize>) -> Self {
        let mut texts = HashMap::new();
        let mut formula_index: HashMap<usize, BTreeMap<u32, Vec<u16>>> = HashMap::new();
        for sheet_idx in 0..workbook.sheet_count() {
            let Some(ws) = workbook.worksheet(sheet_idx) else {
                continue;
            };
            for (row, col, formula) in ws.formula_cells() {
                let key = CellKey::new(sheet_idx, row, col);
                texts.insert(key, formula.to_string());
                formula_index
                    .entry(sheet_idx)
                    .or_default()
                    .entry(row)
                    .or_default()
                    .push(col);
            }
        }
        for rows in formula_index.values_mut() {
            for cols in rows.values_mut() {
                cols.sort_unstable();
            }
        }
        Self {
            workbook,
            texts,
            formula_index,
            selected_sheet,
            parsed: HashMap::new(),
            deps: HashMap::new(),
            depths: HashMap::new(),
            visiting: HashSet::new(),
            parse_failures: 0,
            cycle_hits: 0,
        }
    }

    fn seeds(&self) -> Vec<CellKey> {
        self.texts
            .keys()
            .copied()
            .filter(|key| self.selected_sheet.is_none() || self.selected_sheet == Some(key.sheet))
            .collect()
    }

    fn parse_ast(&mut self, key: CellKey) -> Option<FormulaExpr> {
        if let Some(existing) = self.parsed.get(&key) {
            return existing.clone();
        }
        let parsed = self
            .texts
            .get(&key)
            .and_then(|text| parse_formula(text).ok());
        if parsed.is_none() {
            self.parse_failures += 1;
        }
        self.parsed.insert(key, parsed.clone());
        parsed
    }

    fn push_range_formula_refs(
        &self,
        sheet_idx: usize,
        row_start: u32,
        row_end: u32,
        col_start: u16,
        col_end: u16,
        refs: &mut Vec<CellKey>,
    ) {
        let Some(rows) = self.formula_index.get(&sheet_idx) else {
            return;
        };
        for (&row, cols) in rows.range(row_start..=row_end) {
            let start = cols.partition_point(|&col| col < col_start);
            let end = cols.partition_point(|&col| col <= col_end);
            for &col in &cols[start..end] {
                refs.push(CellKey::new(sheet_idx, row, col));
            }
        }
    }

    fn extract_refs_from_expr(
        &self,
        expr: &FormulaExpr,
        current_sheet: usize,
        refs: &mut Vec<CellKey>,
    ) {
        match expr {
            FormulaExpr::CellRef(cell_ref) => {
                let sheet_idx = cell_ref
                    .sheet
                    .as_ref()
                    .and_then(|name| self.workbook.sheet_index(name))
                    .unwrap_or(current_sheet);
                let key = CellKey::new(sheet_idx, cell_ref.address.row, cell_ref.address.col);
                if self.texts.contains_key(&key) {
                    refs.push(key);
                }
            }
            FormulaExpr::RangeRef(range_ref) => {
                let sheet_idx = range_ref
                    .sheet
                    .as_ref()
                    .and_then(|name| self.workbook.sheet_index(name))
                    .unwrap_or(current_sheet);
                self.push_range_formula_refs(
                    sheet_idx,
                    range_ref.range.start.row,
                    range_ref.range.end.row,
                    range_ref.range.start.col,
                    range_ref.range.end.col,
                    refs,
                );
            }
            FormulaExpr::BinaryOp { left, right, .. } => {
                self.extract_refs_from_expr(left, current_sheet, refs);
                self.extract_refs_from_expr(right, current_sheet, refs);
            }
            FormulaExpr::UnaryOp { operand, .. } => {
                self.extract_refs_from_expr(operand, current_sheet, refs);
            }
            FormulaExpr::Function { args, .. } => {
                for arg in args {
                    self.extract_refs_from_expr(arg, current_sheet, refs);
                }
            }
            FormulaExpr::Array(rows) => {
                for row in rows {
                    for cell in row {
                        self.extract_refs_from_expr(cell, current_sheet, refs);
                    }
                }
            }
            FormulaExpr::ExternalRef(_)
            | FormulaExpr::NameRef(_)
            | FormulaExpr::StructuredRef(_) => {}
            FormulaExpr::Number(_)
            | FormulaExpr::String(_)
            | FormulaExpr::Boolean(_)
            | FormulaExpr::Error(_)
            | FormulaExpr::Empty => {}
        }
    }

    fn deps_for(&mut self, key: CellKey) -> Vec<CellKey> {
        if let Some(existing) = self.deps.get(&key) {
            return existing.clone();
        }
        let Some(ast) = self.parse_ast(key) else {
            self.deps.insert(key, Vec::new());
            return Vec::new();
        };
        let mut refs = Vec::new();
        self.extract_refs_from_expr(&ast, key.sheet, &mut refs);
        refs.sort_unstable_by_key(|k| (k.sheet, k.row, k.col));
        refs.dedup();
        self.deps.insert(key, refs.clone());
        refs
    }

    fn depth_of(&mut self, key: CellKey) -> usize {
        if let Some(depth) = self.depths.get(&key) {
            return *depth;
        }
        if !self.visiting.insert(key) {
            self.cycle_hits += 1;
            return 0;
        }
        let deps = self.deps_for(key);
        let depth = deps
            .into_iter()
            .map(|dep| self.depth_of(dep) + 1)
            .max()
            .unwrap_or(0);
        self.visiting.remove(&key);
        self.depths.insert(key, depth);
        depth
    }
}

fn build_parallel_report(workbook: &Workbook, selected_sheet: Option<usize>) -> ParallelReport {
    let t0 = Instant::now();
    let mut analyzer = ParallelAnalyzer::new(workbook, selected_sheet);
    let seeds = analyzer.seeds();
    for key in seeds {
        let _ = analyzer.depth_of(key);
    }
    let mut histogram: BTreeMap<usize, usize> = BTreeMap::new();
    let mut edge_count = 0usize;
    for (key, depth) in &analyzer.depths {
        *histogram.entry(*depth).or_default() += 1;
        edge_count += analyzer.deps.get(key).map(|d| d.len()).unwrap_or(0);
    }
    let mut top_levels: Vec<(usize, usize)> = histogram.iter().map(|(d, c)| (*d, *c)).collect();
    top_levels.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| a.0.cmp(&b.0)));
    let (widest_level_depth, widest_level_size) = top_levels.first().copied().unwrap_or((0, 0));
    let max_depth = analyzer.depths.values().copied().max().unwrap_or(0);
    ParallelReport {
        analyzed_formulas: analyzer.depths.len(),
        dependency_edges: edge_count,
        max_depth,
        widest_level_depth,
        widest_level_size,
        top_levels: top_levels.into_iter().take(10).collect(),
        parse_failures: analyzer.parse_failures,
        cycle_hits: analyzer.cycle_hits,
        analysis_ms: t0.elapsed().as_secs_f64() * 1000.0,
    }
}

fn parse_args() -> Cli {
    let mut args = std::env::args().skip(1);
    let mut file = None;
    let mut fixture = None;
    let mut serial = false;
    let mut json = false;
    let mut once = false;
    let mut open_only = false;
    let mut parallel_report = false;
    let mut sheet = None;

    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--file" => file = Some(args.next().expect("--file requires a path")),
            "--fixture" => fixture = Some(args.next().expect("--fixture requires a name")),
            "--serial" => serial = true,
            "--json" => json = true,
            "--once" => once = true,
            "--open-only" => open_only = true,
            "--parallel-report" => parallel_report = true,
            "--sheet" => {
                let value = args.next().expect("--sheet requires an index");
                sheet = Some(value.parse::<usize>().expect("invalid --sheet index"));
            }
            _ if arg.starts_with("--sheet=") => {
                sheet = Some(
                    arg.split_once('=')
                        .expect("missing --sheet value")
                        .1
                        .parse::<usize>()
                        .expect("invalid --sheet index"),
                );
            }
            _ if arg.starts_with('-') => panic!("Unknown flag: {arg}"),
            _ => {
                assert!(
                    file.is_none() && fixture.is_none(),
                    "Only one input source is supported"
                );
                file = Some(arg);
            }
        }
    }

    let source = match (file, fixture) {
        (Some(path), None) => SourceSpec::File(path),
        (None, Some(name)) => SourceSpec::Fixture(name),
        (Some(_), Some(_)) => panic!("Provide either a file or a fixture, not both"),
        (None, None) => panic!("Usage: profile_calc <file.xlsx> | --fixture repeated-lookups [--serial] [--json] [--once] [--open-only] [--sheet N]"),
    };

    Cli {
        source,
        serial,
        json,
        once,
        open_only,
        parallel_report,
        sheet,
    }
}

fn main() {
    let cli = parse_args();
    let source_name = source_label(&cli.source);

    let t0 = Instant::now();
    let mut workbook = match &cli.source {
        SourceSpec::File(path) => Workbook::open(path).expect("Failed to open workbook"),
        SourceSpec::Fixture(name) => build_fixture(name),
    };
    let open_time = t0.elapsed();

    let mut total_formulas = 0usize;
    let mut target_formulas = 0usize;
    let mut target_sheet_names = Vec::new();

    if !cli.json {
        eprintln!("Opening {}...", source_name);
        eprintln!(
            "Opened in {:.2?} ({} sheets)",
            open_time,
            workbook.sheet_count()
        );
    }

    for i in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.worksheet(i) {
            let count = sheet.formula_cells().count();
            total_formulas += count;
            let selected = cli.sheet.is_none() || cli.sheet == Some(i);
            if selected {
                target_formulas += count;
                target_sheet_names.push(sheet.name().to_string());
            }
            if !cli.json && count > 0 {
                eprintln!("  Sheet {}: {:>8} formulas  \"{}\"", i, count, sheet.name());
            }
        }
    }

    if let Some(sheet_idx) = cli.sheet {
        assert!(
            sheet_idx < workbook.sheet_count(),
            "sheet index out of bounds"
        );
    }

    if !cli.json {
        eprintln!("Total formulas: {}", total_formulas);
    }

    let parallel_report = if cli.parallel_report {
        Some(build_parallel_report(&workbook, cli.sheet))
    } else {
        None
    };

    if !cli.json {
        if let Some(report) = &parallel_report {
            eprintln!();
            eprintln!("=== Parallelism Report ===");
            eprintln!("Analyzed:    {}", report.analyzed_formulas);
            eprintln!("Edges:       {}", report.dependency_edges);
            eprintln!("Max depth:   {}", report.max_depth);
            eprintln!(
                "Widest:      depth {} ({} formulas)",
                report.widest_level_depth, report.widest_level_size
            );
            eprintln!("Parse fails: {}", report.parse_failures);
            eprintln!("Cycle hits:  {}", report.cycle_hits);
            eprintln!("Analysis:    {:.2} ms", report.analysis_ms);
        }
    }

    let options = CalculationOptions {
        force_full_calculation: true,
        max_threads: if cli.serial { Some(1) } else { None },
        sheets: cli.sheet.map(|s| vec![s]).unwrap_or_default(),
        ..Default::default()
    };

    if cli.open_only {
        if cli.json {
            println!(
                "{}",
                json!({
                    "input": source_name,
                    "fixture": match &cli.source { SourceSpec::Fixture(name) => Some(name.as_str()), _ => None },
                    "open_only": true,
                    "serial": cli.serial,
                    "sheet": cli.sheet,
                    "sheet_names": target_sheet_names,
                    "open_ms": open_time.as_secs_f64() * 1000.0,
                    "total_ms": t0.elapsed().as_secs_f64() * 1000.0,
                    "total_formulas": total_formulas,
                    "target_formulas": target_formulas,
                    "parallel_report": parallel_report.as_ref().map(|report| json!({
                        "analyzed_formulas": report.analyzed_formulas,
                        "dependency_edges": report.dependency_edges,
                        "max_depth": report.max_depth,
                        "widest_level_depth": report.widest_level_depth,
                        "widest_level_size": report.widest_level_size,
                        "top_levels": report.top_levels,
                        "parse_failures": report.parse_failures,
                        "cycle_hits": report.cycle_hits,
                        "analysis_ms": report.analysis_ms,
                    })),
                })
            );
        } else {
            eprintln!();
            eprintln!("=== Open Only ===");
            eprintln!("Open time:   {:.2?}", open_time);
            eprintln!("Total time:  {:.2?}", t0.elapsed());
        }
        return;
    }

    if !cli.json {
        eprintln!();
        eprintln!("Calculating (first run — cold)...");
    }
    let t1 = Instant::now();
    let stats = workbook
        .calculate_with_options(&options)
        .expect("Calculation failed");
    let calc_time = t1.elapsed();

    if cli.once {
        if cli.json {
            println!(
                "{}",
                json!({
                    "input": source_name,
                    "fixture": match &cli.source { SourceSpec::Fixture(name) => Some(name.as_str()), _ => None },
                    "serial": cli.serial,
                    "sheet": cli.sheet,
                    "sheet_names": target_sheet_names,
                    "open_ms": open_time.as_secs_f64() * 1000.0,
                    "calc_ms": calc_time.as_secs_f64() * 1000.0,
                    "total_ms": t0.elapsed().as_secs_f64() * 1000.0,
                    "total_formulas": total_formulas,
                    "target_formulas": target_formulas,
                    "formula_count": stats.formula_count,
                    "cells_calculated": stats.cells_calculated,
                    "iterations": stats.iterations,
                    "converged": stats.converged,
                    "errors": stats.errors,
                    "volatile_cells": stats.volatile_cells,
                    "circular_references": stats.circular_references,
                    "parallel_report": parallel_report.as_ref().map(|report| json!({
                        "analyzed_formulas": report.analyzed_formulas,
                        "dependency_edges": report.dependency_edges,
                        "max_depth": report.max_depth,
                        "widest_level_depth": report.widest_level_depth,
                        "widest_level_size": report.widest_level_size,
                        "top_levels": report.top_levels,
                        "parse_failures": report.parse_failures,
                        "cycle_hits": report.cycle_hits,
                        "analysis_ms": report.analysis_ms,
                    })),
                })
            );
        } else {
            eprintln!();
            eprintln!("=== First Run (cold) ===");
            eprintln!("Formulas:    {}", stats.formula_count);
            eprintln!("Calculated:  {}", stats.cells_calculated);
            eprintln!("Iterations:  {}", stats.iterations);
            eprintln!("Converged:   {}", stats.converged);
            eprintln!("Errors:      {}", stats.errors);
            eprintln!("Volatile:    {}", stats.volatile_cells);
            eprintln!("Circular:    {}", stats.circular_references);
            eprintln!("Calc time:   {:.2?}", calc_time);
            eprintln!("Open time:   {:.2?}", open_time);
            eprintln!("Total time:  {:.2?}", t0.elapsed());
        }
        return;
    }

    if !cli.json {
        eprintln!();
        eprintln!("=== First Run (cold) ===");
        eprintln!("Formulas:    {}", stats.formula_count);
        eprintln!("Calculated:  {}", stats.cells_calculated);
        eprintln!("Iterations:  {}", stats.iterations);
        eprintln!("Converged:   {}", stats.converged);
        eprintln!("Errors:      {}", stats.errors);
        eprintln!("Volatile:    {}", stats.volatile_cells);
        eprintln!("Circular:    {}", stats.circular_references);
        eprintln!("Calc time:   {:.2?}", calc_time);
        eprintln!("Open time:   {:.2?}", open_time);
        eprintln!("Total time:  {:.2?}", t0.elapsed());
        eprintln!();
        eprintln!("Calculating (second run — cached)...");
    }

    let t2 = Instant::now();
    let stats2 = workbook
        .calculate_with_options(&options)
        .expect("Calculation failed");
    let calc_time2 = t2.elapsed();

    if cli.json {
        println!(
            "{}",
            json!({
                "input": source_name,
                "fixture": match &cli.source { SourceSpec::Fixture(name) => Some(name.as_str()), _ => None },
                "serial": cli.serial,
                "sheet": cli.sheet,
                "sheet_names": target_sheet_names,
                "open_ms": open_time.as_secs_f64() * 1000.0,
                "first_run": {
                    "calc_ms": calc_time.as_secs_f64() * 1000.0,
                    "formula_count": stats.formula_count,
                    "cells_calculated": stats.cells_calculated,
                    "iterations": stats.iterations,
                    "converged": stats.converged,
                    "errors": stats.errors,
                    "volatile_cells": stats.volatile_cells,
                    "circular_references": stats.circular_references,
                },
                "second_run": {
                    "calc_ms": calc_time2.as_secs_f64() * 1000.0,
                    "formula_count": stats2.formula_count,
                    "cells_calculated": stats2.cells_calculated,
                    "errors": stats2.errors,
                },
                "speedup": calc_time.as_secs_f64() / calc_time2.as_secs_f64(),
                "total_formulas": total_formulas,
                "target_formulas": target_formulas,
                "parallel_report": parallel_report.as_ref().map(|report| json!({
                    "analyzed_formulas": report.analyzed_formulas,
                    "dependency_edges": report.dependency_edges,
                    "max_depth": report.max_depth,
                    "widest_level_depth": report.widest_level_depth,
                    "widest_level_size": report.widest_level_size,
                    "top_levels": report.top_levels,
                    "parse_failures": report.parse_failures,
                    "cycle_hits": report.cycle_hits,
                    "analysis_ms": report.analysis_ms,
                })),
            })
        );
    } else {
        eprintln!();
        eprintln!("=== Second Run (cached) ===");
        eprintln!("Formulas:    {}", stats2.formula_count);
        eprintln!("Calculated:  {}", stats2.cells_calculated);
        eprintln!("Errors:      {}", stats2.errors);
        eprintln!("Calc time:   {:.2?}", calc_time2);
        eprintln!(
            "Speedup:     {:.1}x",
            calc_time.as_secs_f64() / calc_time2.as_secs_f64()
        );
    }
}
