//! Calculation engine benchmarks — dependency resolution and evaluation.

#[path = "helpers.rs"]
mod helpers;

#[path = "../perf_fixtures.rs"]
mod perf_fixtures;

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion};
use duke_sheets::WorkbookCalculationExt;

/// Linear chain: A1=1, A2=A1+1, … AN=A(N-1)+1.
fn bench_calc_linear(c: &mut Criterion) {
    let depths: &[u32] = &[100, 500, 1000];
    let mut group = c.benchmark_group("calculation/linear_chain");

    for &depth in depths {
        group.bench_with_input(BenchmarkId::from_parameter(depth), &depth, |b, &depth| {
            b.iter_batched(
                || helpers::generate_linear_chain(depth),
                |mut wb| {
                    let stats = wb.calculate().unwrap();
                    black_box(stats);
                },
                criterion::BatchSize::SmallInput,
            )
        });
    }
    group.finish();
}

/// Fan-out: N columns in row 1 with values, row 2 each sums A1:X1.
fn bench_calc_fanout(c: &mut Criterion) {
    let widths: &[u16] = &[26, 52, 100, 200];
    let mut group = c.benchmark_group("calculation/fan_out");

    for &width in widths {
        group.bench_with_input(BenchmarkId::from_parameter(width), &width, |b, &width| {
            b.iter_batched(
                || helpers::generate_fanout(width),
                |mut wb| {
                    let stats = wb.calculate().unwrap();
                    black_box(stats);
                },
                criterion::BatchSize::SmallInput,
            )
        });
    }
    group.finish();
}

/// Cross-sheet: Sheet1 has values, Sheet2 references Sheet1.
fn bench_calc_cross_sheet(c: &mut Criterion) {
    let sizes: &[u32] = &[100, 500, 1000, 5000];
    let mut group = c.benchmark_group("calculation/cross_sheet");

    for &rows in sizes {
        group.bench_with_input(BenchmarkId::from_parameter(rows), &rows, |b, &rows| {
            b.iter_batched(
                || helpers::generate_cross_sheet(rows),
                |mut wb| {
                    let stats = wb.calculate().unwrap();
                    black_box(stats);
                },
                criterion::BatchSize::SmallInput,
            )
        });
    }
    group.finish();
}

/// Mixed workbook: values + formulas across sheets.
fn bench_calc_mixed(c: &mut Criterion) {
    let sizes: &[u32] = &[100, 500, 1000];
    let mut group = c.benchmark_group("calculation/mixed");

    for &rows in sizes {
        group.bench_with_input(BenchmarkId::from_parameter(rows), &rows, |b, &rows| {
            b.iter_batched(
                || build_mixed_workbook(rows),
                |mut wb| {
                    let stats = wb.calculate().unwrap();
                    black_box(stats);
                },
                criterion::BatchSize::SmallInput,
            )
        });
    }
    group.finish();
}

fn bench_calc_repeated_lookups(c: &mut Criterion) {
    let mut group = c.benchmark_group("calculation/repeated_lookups");
    group.bench_function("repeated_lookups", |b| {
        b.iter_batched(
            || perf_fixtures::build_fixture("repeated-lookups"),
            |mut wb| {
                let stats = wb.calculate().unwrap();
                black_box(stats);
            },
            criterion::BatchSize::SmallInput,
        )
    });
    group.finish();
}

/// Build a workbook with a mix of values and formulas.
fn build_mixed_workbook(rows: u32) -> duke_sheets::Workbook {
    let mut wb = duke_sheets::Workbook::new();
    let _ = wb.add_worksheet_with_name("Summary");

    // Sheet 0 "Data": numbers and strings
    {
        let s = wb.worksheet_mut(0).unwrap();
        let _ = s.set_name("Data");
        for row in 0..rows {
            let _ = s.set_cell_value_at(row, 0, row as f64 + 1.0);
            let _ = s.set_cell_value_at(row, 1, format!("item_{}", row));
            let _ = s.set_cell_value_at(row, 2, (row as f64 + 1.0) * 10.5);
        }
    }

    // Sheet 1 "Summary": formulas referencing Data
    {
        let s = wb.worksheet_mut(1).unwrap();
        let _ = s.set_cell_formula_at(0, 0, &format!("=SUM(Data!A1:A{})", rows));
        let _ = s.set_cell_formula_at(1, 0, &format!("=AVERAGE(Data!C1:C{})", rows));
        let _ = s.set_cell_formula_at(2, 0, &format!("=COUNT(Data!A1:A{})", rows));
        let _ = s.set_cell_formula_at(3, 0, &format!("=MAX(Data!C1:C{})", rows));
        let _ = s.set_cell_formula_at(4, 0, &format!("=MIN(Data!C1:C{})", rows));
        // Per-row computed column
        for row in 0..rows {
            let _ = s.set_cell_formula_at(row, 1, &format!("=Data!A{}*Data!C{}", row + 1, row + 1));
        }
    }

    wb
}

criterion_group! {
    name = benches;
    config = helpers::fast_criterion();
    targets = bench_calc_linear, bench_calc_fanout, bench_calc_cross_sheet, bench_calc_mixed, bench_calc_repeated_lookups
}
criterion_main!(benches);
