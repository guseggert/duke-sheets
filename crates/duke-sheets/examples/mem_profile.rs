//! Memory profiling tool - reports cell counts, type sizes, and memory estimates.
//! Build: cargo build --release --features full -p duke-sheets --example mem_profile

use duke_sheets::prelude::*;
use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use std::time::Instant;

fn rss_kb() -> usize {
    std::fs::read_to_string("/proc/self/statm")
        .ok()
        .and_then(|s| s.split_whitespace().nth(1)?.parse::<usize>().ok())
        .map(|pages| pages * 4)
        .unwrap_or(0)
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let path = args.get(1).expect("Usage: mem_profile <file.xlsx>");

    eprintln!("=== Type Sizes ===");
    eprintln!("CellValue:    {} bytes", std::mem::size_of::<CellValue>());
    eprintln!(
        "FormulaExpr:  {} bytes",
        std::mem::size_of::<duke_sheets_formula::FormulaExpr>()
    );
    eprintln!(
        "CellKey:      {} bytes",
        std::mem::size_of::<duke_sheets_formula::dependency::CellKey>()
    );
    eprintln!(
        "FormulaValue: {} bytes",
        std::mem::size_of::<duke_sheets_formula::FormulaValue>()
    );
    eprintln!();

    let rss_before = rss_kb();
    eprintln!("RSS before open: {} MB", rss_before / 1024);

    eprintln!("Opening {}...", path);
    let t0 = Instant::now();
    let mut workbook = Workbook::open(&path).expect("Failed to open workbook");
    let open_time = t0.elapsed();
    let rss_after_open = rss_kb();
    eprintln!("Opened in {:.2?}", open_time);
    eprintln!(
        "RSS after open: {} MB (+{} MB)",
        rss_after_open / 1024,
        (rss_after_open - rss_before) / 1024
    );

    eprintln!("\n=== Per-Sheet Cell Analysis ===");
    let mut grand_cells = 0usize;
    let mut grand_formulas = 0usize;

    for i in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.worksheet(i) {
            let cell_count = sheet.cell_count();
            let formula_count = sheet.formula_cells().count();
            if cell_count == 0 {
                continue;
            }

            let mut numbers = 0usize;
            let mut strings = 0usize;
            let mut booleans = 0usize;
            let mut errors = 0usize;
            let mut empty_styled = 0usize;
            let mut richtext = 0usize;
            let mut spill = 0usize;
            let mut max_col: u16 = 0;

            if let Some(range) = sheet.used_range() {
                for row in range.start.row..=range.end.row {
                    for col in range.start.col..=range.end.col {
                        if let Some(cell) = sheet.cell_at(row, col) {
                            if col > max_col {
                                max_col = col;
                            }
                            match &cell.value {
                                CellValue::Empty => empty_styled += 1,
                                CellValue::Number(_) => numbers += 1,
                                CellValue::String(_) => strings += 1,
                                CellValue::Boolean(_) => booleans += 1,
                                CellValue::Error(_) => errors += 1,
                                CellValue::SpillTarget { .. } => spill += 1,
                                CellValue::RichText(_) => richtext += 1,
                            }
                        }
                    }
                }
            }

            let non_empty = numbers + strings + booleans + errors + richtext + spill;
            eprintln!("Sheet {}: \"{}\" - {} stored, {} non-empty, {} empty-styled, {} formulas, max_col={}",
                i, sheet.name(), cell_count, non_empty, empty_styled, formula_count, max_col);
            if max_col > 50 {
                eprintln!(
                    "  types: num={} str={} bool={} err={} rich={} spill={}",
                    numbers, strings, booleans, errors, richtext, spill
                );
            }

            grand_cells += cell_count;
            grand_formulas += formula_count;
        }
    }

    eprintln!("\n=== Totals ===");
    eprintln!("Total stored cells: {}", grand_cells);
    eprintln!("Total formulas:     {}", grand_formulas);

    // Memory estimates (pre-calc)
    let cv_size = std::mem::size_of::<CellValue>();
    let overhead = 56usize; // AHashMap per-entry
    let per_cell = cv_size + 4 + 6 + overhead; // value + style_idx + key + hashmap
    let per_formula = 48 + 6 + overhead; // String(24) + Option<Vec>(24) + key + hashmap
    let est_cells = grand_cells * per_cell;
    let est_formulas = grand_formulas * per_formula;
    eprintln!("\n=== Pre-Calc Memory Estimates ===");
    eprintln!(
        "Cell grid:     ~{:.1} MB  ({} cells × {} bytes)",
        est_cells as f64 / 1e6,
        grand_cells,
        per_cell
    );
    eprintln!(
        "Formula table: ~{:.1} MB  ({} formulas × {} bytes)",
        est_formulas as f64 / 1e6,
        grand_formulas,
        per_formula
    );
    eprintln!(
        "Actual RSS growth from open: {} MB",
        (rss_after_open - rss_before) / 1024
    );

    // Calculate
    eprintln!("\nCalculating (serial)...");
    let rss_pre = rss_kb();
    let t1 = Instant::now();
    let stats = workbook
        .calculate_with_options(&CalculationOptions {
            force_full_calculation: true,
            max_threads: Some(1),
            ..Default::default()
        })
        .expect("calc failed");
    let calc_time = t1.elapsed();
    let rss_post = rss_kb();

    eprintln!(
        "Calculated {} formulas in {:.2?} ({} errors)",
        stats.cells_calculated, calc_time, stats.errors
    );
    eprintln!("\n=== RSS Summary ===");
    eprintln!("Before open:  {:>5} MB", rss_before / 1024);
    eprintln!(
        "After open:   {:>5} MB  (+{} MB)",
        rss_after_open / 1024,
        (rss_after_open - rss_before) / 1024
    );
    eprintln!(
        "After calc:   {:>5} MB  (+{} MB from calc)",
        rss_post / 1024,
        (rss_post - rss_pre) / 1024
    );
    eprintln!("Total growth: {:>5} MB", (rss_post - rss_before) / 1024);
}
