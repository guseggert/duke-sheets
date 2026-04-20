//! CSV read and write benchmarks (duke-sheets only - no comparable Rust
//! spreadsheet-oriented CSV competitor).

#[path = "helpers.rs"]
mod helpers;

use criterion::{black_box, criterion_group, criterion_main, Criterion};
use duke_sheets::{CsvReadOptions, CsvReader, CsvWriteOptions, CsvWriter};

fn bench_csv_read(c: &mut Criterion) {
    for &(label, rows, cols) in helpers::SIZES {
        let csv_data = helpers::generate_csv_string(rows, cols);
        let mut group = c.benchmark_group(format!("csv_read/{}", label));

        group.bench_function("duke-sheets", |b| {
            b.iter(|| {
                let cursor = std::io::Cursor::new(csv_data.as_bytes());
                let ws = CsvReader::read(cursor, &CsvReadOptions::default()).unwrap();
                black_box(ws);
            })
        });

        group.finish();
    }
}

fn bench_csv_write(c: &mut Criterion) {
    for &(label, rows, cols) in helpers::SIZES {
        let wb = helpers::generate_workbook(rows, cols);
        let sheet = wb.worksheet(0).unwrap();
        let mut group = c.benchmark_group(format!("csv_write/{}", label));

        group.bench_function("duke-sheets", |b| {
            b.iter(|| {
                let mut buf = Vec::new();
                CsvWriter::write(sheet, &mut buf, &CsvWriteOptions::default()).unwrap();
                black_box(buf);
            })
        });

        group.finish();
    }
}

criterion_group! {
    name = benches;
    config = helpers::fast_criterion();
    targets = bench_csv_read, bench_csv_write
}
criterion_main!(benches);
