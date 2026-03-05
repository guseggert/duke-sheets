//! Formula parser throughput benchmarks.

#[path = "helpers.rs"]
mod helpers;

use criterion::{black_box, criterion_group, criterion_main, Criterion};
use duke_sheets::parse_formula;

fn bench_parse_simple(c: &mut Criterion) {
    let formulas = helpers::simple_formulas();
    c.bench_function("formula_parse/simple", |b| {
        b.iter(|| {
            for f in &formulas {
                black_box(parse_formula(f).unwrap());
            }
        })
    });
}

fn bench_parse_medium(c: &mut Criterion) {
    let formulas = helpers::medium_formulas();
    c.bench_function("formula_parse/medium", |b| {
        b.iter(|| {
            for f in &formulas {
                black_box(parse_formula(f).unwrap());
            }
        })
    });
}

fn bench_parse_complex(c: &mut Criterion) {
    let formulas = helpers::complex_formulas();
    c.bench_function("formula_parse/complex", |b| {
        b.iter(|| {
            for f in &formulas {
                black_box(parse_formula(f).unwrap());
            }
        })
    });
}

/// Throughput: parse many formulas to get ops/sec.
fn bench_parse_throughput(c: &mut Criterion) {
    // Build a corpus of 1000 formulas
    let mut corpus = Vec::with_capacity(1000);
    for i in 0..1000 {
        let col1 = (b'A' + (i % 26) as u8) as char;
        let col2 = (b'A' + ((i + 1) % 26) as u8) as char;
        let row = (i % 100) + 1;
        match i % 5 {
            0 => corpus.push(format!("={}{}", col1, row)),
            1 => corpus.push(format!("={}{}+{}{}", col1, row, col2, row + 1)),
            2 => corpus.push(format!("=SUM({}1:{}{})", col1, col1, row)),
            3 => corpus.push(format!("=IF({}{}>0,1,0)", col1, row)),
            _ => corpus.push(format!("=AVERAGE({}1:{}{})", col1, col2, row)),
        }
    }

    c.bench_function("formula_parse/throughput_1000", |b| {
        b.iter(|| {
            for f in &corpus {
                black_box(parse_formula(f).unwrap());
            }
        })
    });
}

criterion_group! {
    name = benches;
    config = helpers::fast_criterion();
    targets = bench_parse_simple, bench_parse_medium, bench_parse_complex, bench_parse_throughput
}
criterion_main!(benches);
