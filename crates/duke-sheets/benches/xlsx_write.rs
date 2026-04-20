//! XLSX write benchmarks - duke-sheets vs rust_xlsxwriter vs umya-spreadsheet.

#[path = "helpers.rs"]
mod helpers;

use std::io::Cursor;

use criterion::{black_box, criterion_group, criterion_main, Criterion};

/// Benchmark serialize-only: workbook already populated, measure just the
/// XLSX serialization to bytes.  rust_xlsxwriter cannot separate creation
/// from serialization, so it only appears in the create+write benchmark.
fn bench_xlsx_write_serialize(c: &mut Criterion) {
    for &(label, rows, cols) in helpers::SIZES {
        let duke_wb = helpers::generate_workbook(rows, cols);

        // Pre-build an umya workbook with the same data
        let umya_wb = build_umya_workbook(rows, cols);

        let mut group = c.benchmark_group(format!("xlsx_write_serialize/{}", label));

        // -- duke-sheets --
        group.bench_function("duke-sheets", |b| {
            b.iter(|| {
                let mut buf = Cursor::new(Vec::new());
                duke_sheets::XlsxWriter::write(&duke_wb, &mut buf).unwrap();
                black_box(buf.into_inner());
            })
        });

        // -- umya-spreadsheet --
        group.bench_function("umya-spreadsheet", |b| {
            b.iter(|| {
                let mut buf = Vec::new();
                umya_spreadsheet::writer::xlsx::write_writer(&umya_wb, &mut buf).unwrap();
                black_box(buf);
            })
        });

        group.finish();
    }
}

/// Benchmark the full create+populate+serialize pipeline for all three.
fn bench_xlsx_write_full(c: &mut Criterion) {
    for &(label, rows, cols) in helpers::SIZES {
        let mut group = c.benchmark_group(format!("xlsx_write_full/{}", label));

        // -- duke-sheets --
        group.bench_function("duke-sheets", |b| {
            b.iter(|| {
                let wb = helpers::generate_workbook(rows, cols);
                let mut buf = Cursor::new(Vec::new());
                duke_sheets::XlsxWriter::write(&wb, &mut buf).unwrap();
                black_box(buf.into_inner());
            })
        });

        // -- rust_xlsxwriter --
        group.bench_function("rust_xlsxwriter", |b| {
            b.iter(|| {
                let mut workbook = rust_xlsxwriter::Workbook::new();
                let worksheet = workbook.add_worksheet();
                for row in 0..rows {
                    for col in 0..cols {
                        let idx = row * cols as u32 + col as u32;
                        match idx % 4 {
                            0 => {
                                worksheet
                                    .write_string(row, col, format!("str_{}_{}", row, col))
                                    .unwrap();
                            }
                            1 => {
                                worksheet.write_number(row, col, idx as f64 * 1.5).unwrap();
                            }
                            2 => {
                                worksheet.write_number(row, col, idx as f64).unwrap();
                            }
                            _ => {
                                worksheet.write_boolean(row, col, row % 2 == 0).unwrap();
                            }
                        }
                    }
                }
                let bytes = workbook.save_to_buffer().unwrap();
                black_box(bytes);
            })
        });

        // -- umya-spreadsheet --
        group.bench_function("umya-spreadsheet", |b| {
            b.iter(|| {
                let wb = build_umya_workbook(rows, cols);
                let mut buf = Vec::new();
                umya_spreadsheet::writer::xlsx::write_writer(&wb, &mut buf).unwrap();
                black_box(buf);
            })
        });

        group.finish();
    }
}

fn build_umya_workbook(rows: u32, cols: u16) -> umya_spreadsheet::Spreadsheet {
    let mut wb = umya_spreadsheet::new_file();
    let sheet = wb.get_sheet_mut(&0).unwrap();
    for row in 0..rows {
        for col in 0..cols {
            let idx = row * cols as u32 + col as u32;
            // umya uses 1-indexed (col, row)
            let c = (col as u32) + 1;
            let r = row + 1;
            match idx % 4 {
                0 => {
                    sheet
                        .get_cell_mut((c, r))
                        .set_value_string(format!("str_{}_{}", row, col));
                }
                1 => {
                    sheet
                        .get_cell_mut((c, r))
                        .set_value_number(idx as f64 * 1.5);
                }
                2 => {
                    sheet.get_cell_mut((c, r)).set_value_number(idx as f64);
                }
                _ => {
                    sheet.get_cell_mut((c, r)).set_value_bool(row % 2 == 0);
                }
            }
        }
    }
    wb
}

criterion_group! {
    name = benches;
    config = helpers::fast_criterion();
    targets = bench_xlsx_write_serialize, bench_xlsx_write_full
}
criterion_main!(benches);
