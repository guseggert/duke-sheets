//! XLSX read benchmarks - duke-sheets vs calamine vs umya-spreadsheet.

#[path = "helpers.rs"]
mod helpers;

use std::io::Cursor;

use calamine::Reader;
use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn bench_xlsx_read(c: &mut Criterion) {
    for &(label, rows, cols) in helpers::SIZES {
        let wb = helpers::generate_workbook(rows, cols);
        let bytes = helpers::workbook_to_xlsx_bytes(&wb);
        let mut group = c.benchmark_group(format!("xlsx_read/{}", label));

        // -- duke-sheets --
        group.bench_function("duke-sheets", |b| {
            b.iter(|| {
                let cursor = Cursor::new(bytes.as_slice());
                let wb = duke_sheets::XlsxReader::read(cursor).unwrap();
                black_box(wb);
            })
        });

        // -- calamine --
        group.bench_function("calamine", |b| {
            b.iter(|| {
                let cursor = Cursor::new(bytes.as_slice());
                let mut wb: calamine::Xlsx<_> = calamine::open_workbook_from_rs(cursor).unwrap();
                // Force sheet data to be read
                let range = wb.worksheet_range("Data").unwrap();
                for row in range.rows() {
                    black_box(row);
                }
            })
        });

        // -- umya-spreadsheet --
        group.bench_function("umya-spreadsheet", |b| {
            b.iter(|| {
                let cursor = Cursor::new(bytes.as_slice());
                let wb = umya_spreadsheet::reader::xlsx::read_reader(cursor, true).unwrap();
                black_box(wb);
            })
        });

        group.finish();
    }
}

criterion_group! {
    name = benches;
    config = helpers::fast_criterion();
    targets = bench_xlsx_read
}
criterion_main!(benches);
