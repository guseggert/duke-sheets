//! Standalone profiling harness for workbook calculation.
//! Build: cargo build --release --features full -p duke-sheets --example profile_calc
//! Or:   rustc with appropriate deps

use duke_sheets::prelude::*;
use duke_sheets::{CalculationOptions, WorkbookCalculationExt};
use std::time::Instant;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let path = args.get(1).expect("Usage: profile-calc <file.xlsx> [--serial]");
    let force_serial = args.iter().any(|a| a == "--serial");

    eprintln!("Opening {}...", path);
    let t0 = Instant::now();
    let mut workbook = Workbook::open(&path).expect("Failed to open workbook");
    let open_time = t0.elapsed();
    eprintln!(
        "Opened in {:.2?} ({} sheets)",
        open_time,
        workbook.sheet_count()
    );

    // Print per-sheet formula counts
    let mut total_formulas = 0usize;
    for i in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.worksheet(i) {
            let count = sheet.formula_cells().count();
            total_formulas += count;
            if count > 0 {
                eprintln!("  Sheet {}: {:>8} formulas  \"{}\"", i, count, sheet.name());
            }
        }
    }
    eprintln!("Total formulas: {}", total_formulas);

    eprintln!("\nCalculating (first run — cold)...");
    let t1 = Instant::now();
    let options = CalculationOptions {
        force_full_calculation: true,
        max_threads: if force_serial { Some(1) } else { None },
        ..Default::default()
    };
    let stats = workbook
        .calculate_with_options(&options)
        .expect("Calculation failed");
    let calc_time = t1.elapsed();

    eprintln!("\n=== First Run (cold) ===");
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

    // Second calculation: should hit the persistent cache and skip parse+DFS.
    eprintln!("\nCalculating (second run — cached)...");
    let t2 = Instant::now();
    let stats2 = workbook
        .calculate_with_options(&options)
        .expect("Calculation failed");
    let calc_time2 = t2.elapsed();

    eprintln!("\n=== Second Run (cached) ===");
    eprintln!("Formulas:    {}", stats2.formula_count);
    eprintln!("Calculated:  {}", stats2.cells_calculated);
    eprintln!("Errors:      {}", stats2.errors);
    eprintln!("Calc time:   {:.2?}", calc_time2);
    eprintln!("Speedup:     {:.1}x", calc_time.as_secs_f64() / calc_time2.as_secs_f64());
}
