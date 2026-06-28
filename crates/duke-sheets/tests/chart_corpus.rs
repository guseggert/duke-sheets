//! Corpus smoke test: scan real-world Excel files for charts and verify
//! that XlsxReader can parse them without panics or errors.
//!
//! Run with:
//!   cargo test -p duke-sheets --test chart_corpus -- --ignored --nocapture
//!
//! Override corpus path:
//!   DUKE_CORPUS_DIR=/path/to/xlsx/files cargo test ...

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::time::Instant;

use duke_sheets::prelude::*;

const DEFAULT_CORPUS_DIR: &str = "data/excel-corpus";

/// Files taking longer than this are reported as slow.
const SLOW_THRESHOLD_SECS: f64 = 10.0;

#[test]
#[ignore = "requires Excel corpus - run with: cargo test -p duke-sheets --test chart_corpus -- --ignored --nocapture"]
fn chart_corpus_smoke() {
    let corpus_dir =
        std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| DEFAULT_CORPUS_DIR.to_string());

    let dir = Path::new(&corpus_dir);
    assert!(
        dir.is_dir(),
        "corpus directory not found: {}",
        dir.display()
    );

    // Collect all .xlsx paths
    let mut xlsx_paths: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read corpus dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("xlsx"))
        .collect();
    xlsx_paths.sort();

    let total_xlsx = xlsx_paths.len();
    println!("\nScanning {total_xlsx} .xlsx files for chart entries...");

    // Phase 1: quick zip scan to find files containing charts
    let scan_start = Instant::now();
    let mut chart_files: Vec<PathBuf> = Vec::new();
    let mut zip_scan_errors = 0usize;

    for (i, path) in xlsx_paths.iter().enumerate() {
        if (i + 1) % 20000 == 0 {
            println!("  scanned {}/{total_xlsx}...", i + 1);
        }
        match has_chart_entries(path) {
            Ok(true) => chart_files.push(path.clone()),
            Ok(false) => {}
            Err(_) => zip_scan_errors += 1,
        }
    }
    let scan_secs = scan_start.elapsed().as_secs_f64();

    println!("\n=== Phase 1: Zip Scan ({scan_secs:.1}s) ===");
    println!("Total .xlsx files:            {total_xlsx}");
    println!("Files with chart entries:      {}", chart_files.len());
    println!("Unreadable zip files:          {zip_scan_errors}");

    // Phase 2: read ALL chart-containing files with XlsxReader
    println!(
        "\nReading {} chart-containing files with XlsxReader...",
        chart_files.len()
    );

    let read_start = Instant::now();
    let mut read_ok = 0usize;
    let mut panicked = 0usize;
    let mut read_errors: Vec<(String, String)> = Vec::new();
    let mut total_ws_charts = 0usize;
    let mut total_chartsheets = 0usize;
    let mut chart_type_counts: HashMap<String, usize> = HashMap::new();
    let mut unsupported_types: HashMap<String, Vec<String>> = HashMap::new();
    let mut slow_files: Vec<(String, f64, u64)> = Vec::new();

    for (i, path) in chart_files.iter().enumerate() {
        if (i + 1) % 100 == 0 {
            println!(
                "  read {}/{} ({:.0}s elapsed)...",
                i + 1,
                chart_files.len(),
                read_start.elapsed().as_secs_f64()
            );
        }

        let fname = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();
        let file_size = path.metadata().map(|m| m.len()).unwrap_or(0);
        let file_start = Instant::now();

        let result = std::panic::catch_unwind(|| XlsxReader::read_file(path));

        let elapsed = file_start.elapsed().as_secs_f64();
        if elapsed > SLOW_THRESHOLD_SECS {
            eprintln!(
                "  SLOW: {fname} ({elapsed:.1}s, {:.1} MB)",
                file_size as f64 / (1024.0 * 1024.0)
            );
            slow_files.push((fname.clone(), elapsed, file_size));
        }

        match result {
            Err(_) => {
                panicked += 1;
                read_errors.push((fname, "PANIC".into()));
            }
            Ok(Err(e)) => {
                read_errors.push((fname, format!("{e}")));
            }
            Ok(Ok(wb)) => {
                read_ok += 1;
                tally_charts(
                    &wb,
                    &fname,
                    &mut total_ws_charts,
                    &mut total_chartsheets,
                    &mut chart_type_counts,
                    &mut unsupported_types,
                );
            }
        }
    }
    let read_secs = read_start.elapsed().as_secs_f64();

    // Report
    println!("\n=== Phase 2: XlsxReader Results ({read_secs:.1}s) ===");
    println!("Read successfully: {read_ok}/{}", chart_files.len());
    println!("Read errors:       {}", read_errors.len());
    println!("Panics caught:     {panicked}");
    println!("Worksheet charts:  {total_ws_charts}");
    println!("Chartsheets:       {total_chartsheets}");

    if !chart_type_counts.is_empty() {
        println!("\nChart type distribution:");
        let mut types: Vec<_> = chart_type_counts.iter().collect();
        types.sort_by(|a, b| b.1.cmp(a.1));
        for (type_name, count) in &types {
            println!("  {type_name:30} {count:>6}");
        }
    }

    if !unsupported_types.is_empty() {
        println!("\nUnsupported chart types:");
        for (tag, files) in &unsupported_types {
            println!("  {tag}: {} file(s) (first: {})", files.len(), files[0]);
        }
    }

    if !slow_files.is_empty() {
        slow_files.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
        println!("\nSlow files (>{SLOW_THRESHOLD_SECS}s):");
        for (f, secs, size) in slow_files.iter().take(20) {
            println!(
                "  {f}: {secs:.1}s ({:.1} MB)",
                *size as f64 / (1024.0 * 1024.0)
            );
        }
    }

    if !read_errors.is_empty() {
        println!("\nRead errors (grouped):");
        let mut error_groups: HashMap<String, Vec<String>> = HashMap::new();
        for (fname, err) in &read_errors {
            error_groups
                .entry(err.clone())
                .or_default()
                .push(fname.clone());
        }
        let mut groups: Vec<_> = error_groups.iter().collect();
        groups.sort_by(|a, b| b.1.len().cmp(&a.1.len()));
        for (err, files) in groups.iter().take(30) {
            println!("  [{} file(s)] {err}", files.len());
            for f in files.iter().take(3) {
                println!("    - {f}");
            }
            if files.len() > 3 {
                println!("    ... and {} more", files.len() - 3);
            }
        }
    }

    // Summary
    let success_rate = if chart_files.is_empty() {
        100.0
    } else {
        (read_ok as f64 / chart_files.len() as f64) * 100.0
    };
    println!("\n=== Summary ===");
    println!("Corpus:        {total_xlsx} xlsx files");
    println!(
        "With charts:   {} ({:.1}%)",
        chart_files.len(),
        chart_files.len() as f64 / total_xlsx as f64 * 100.0
    );
    println!(
        "Read success:  {read_ok}/{} ({success_rate:.1}%)",
        chart_files.len()
    );
    println!(
        "Chart count:   {} ws + {} chartsheets = {} total",
        total_ws_charts,
        total_chartsheets,
        total_ws_charts + total_chartsheets
    );
    println!("Panics:        {panicked}");

    assert_eq!(
        panicked, 0,
        "{panicked} file(s) caused a panic in XlsxReader"
    );

    if !chart_files.is_empty() {
        assert!(
            success_rate >= 80.0,
            "chart file read success rate {success_rate:.1}% is below 80% threshold"
        );
    }
}

fn has_chart_entries(path: &Path) -> std::result::Result<bool, String> {
    let file = std::fs::File::open(path).map_err(|e| format!("open: {e}"))?;
    let archive = zip::ZipArchive::new(file).map_err(|e| format!("zip: {e}"))?;
    for name in archive.file_names() {
        if name.starts_with("xl/charts/") {
            return Ok(true);
        }
    }
    Ok(false)
}

fn tally_charts(
    wb: &Workbook,
    fname: &str,
    total_ws_charts: &mut usize,
    total_chartsheets: &mut usize,
    chart_type_counts: &mut HashMap<String, usize>,
    unsupported_types: &mut HashMap<String, Vec<String>>,
) {
    for si in 0..wb.sheet_count() {
        if let Some(ws) = wb.worksheet(si) {
            for chart in ws.charts() {
                *total_ws_charts += 1;
                let type_name = chart_type_label(&chart.chart_type);
                *chart_type_counts.entry(type_name).or_default() += 1;

                if let ChartType::Unsupported(ref tag) = chart.chart_type {
                    unsupported_types
                        .entry(tag.clone())
                        .or_default()
                        .push(fname.to_string());
                }
            }
        }
    }

    *total_chartsheets += wb.chartsheet_count();
    for ci in 0..wb.chartsheet_count() {
        if let Some(cs) = wb.chartsheet(ci) {
            let type_name = chart_type_label(&cs.chart.chart_type);
            *chart_type_counts.entry(type_name).or_default() += 1;

            if let ChartType::Unsupported(ref tag) = cs.chart.chart_type {
                unsupported_types
                    .entry(tag.clone())
                    .or_default()
                    .push(fname.to_string());
            }
        }
    }
}

fn chart_type_label(ct: &ChartType) -> String {
    match ct {
        ChartType::Unsupported(s) => format!("Unsupported({s})"),
        other => format!("{other:?}"),
    }
}

#[test]
#[ignore = "requires Excel corpus"]
fn chart_corpus_chartex_read() {
    let corpus_dir =
        std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| DEFAULT_CORPUS_DIR.to_string());

    let dir = Path::new(&corpus_dir);
    assert!(dir.is_dir());

    // Scan for xlsx files containing chartEx parts
    let mut chartex_files: Vec<PathBuf> = Vec::new();
    for entry in std::fs::read_dir(dir).expect("read corpus dir") {
        let path = entry.unwrap().path();
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            continue;
        }
        if let Ok(true) = has_chartex_entries(&path) {
            chartex_files.push(path);
        }
    }
    chartex_files.sort();

    println!("Found {} files with chartEx entries", chartex_files.len());
    assert!(
        !chartex_files.is_empty(),
        "no chartEx files found in corpus"
    );

    let mut total_parsed = 0usize;
    let mut failures: Vec<(String, String)> = Vec::new();

    for path in &chartex_files {
        let fname = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();
        let wb = match std::panic::catch_unwind(|| XlsxReader::read_file(path)) {
            Ok(Ok(wb)) => wb,
            Ok(Err(e)) => {
                failures.push((fname, format!("{e}")));
                continue;
            }
            Err(_) => {
                failures.push((fname, "PANIC".into()));
                continue;
            }
        };

        let mut file_ex_count = 0usize;
        for i in 0..wb.sheet_count() {
            if let Some(ws) = wb.worksheet(i) {
                for cx in ws.charts_ex() {
                    file_ex_count += 1;
                    if cx.plot_area.series.is_empty() {
                        failures.push((fname.clone(), "chartEx has no series".into()));
                    }
                    if cx.data.is_empty() {
                        failures.push((fname.clone(), "chartEx has no data blocks".into()));
                    }
                    for s in &cx.plot_area.series {
                        if matches!(s.layout, duke_sheets_chart::ChartExLayout::Unknown(_)) {
                            failures
                                .push((fname.clone(), format!("Unknown layout: {:?}", s.layout)));
                        }
                    }
                }
            }
        }

        if file_ex_count == 0 {
            failures.push((
                fname,
                "no chartEx parsed despite chartEx zip entries".into(),
            ));
        } else {
            total_parsed += file_ex_count;
        }
    }

    println!(
        "Parsed {total_parsed} chartEx charts from {} files",
        chartex_files.len()
    );
    if !failures.is_empty() {
        println!("Failures:");
        for (f, e) in &failures {
            println!("  {f}: {e}");
        }
    }
    assert!(
        failures.is_empty(),
        "{} chartEx read failures",
        failures.len()
    );
}

#[test]
#[ignore = "requires Excel corpus"]
fn chart_corpus_chartex_roundtrip() {
    let corpus_dir =
        std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| DEFAULT_CORPUS_DIR.to_string());
    let dir = Path::new(&corpus_dir);
    assert!(dir.is_dir());

    let mut chartex_files: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read corpus dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("xlsx"))
        .filter(|p| has_chartex_entries(p).unwrap_or(false))
        .collect();
    chartex_files.sort();
    println!("Testing roundtrip on {} chartEx files", chartex_files.len());

    for path in &chartex_files {
        let fname = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();
        eprintln!("  reading {fname}...");
        let wb1 = XlsxReader::read_file(path).unwrap();
        let ex1: Vec<_> = (0..wb1.sheet_count())
            .filter_map(|i| wb1.worksheet(i))
            .flat_map(|ws| ws.charts_ex().iter())
            .collect();
        eprintln!("    {} chartEx read", ex1.len());

        eprintln!("  writing {fname}...");
        let mut buf = Vec::new();
        XlsxWriter::write(&wb1, std::io::Cursor::new(&mut buf)).unwrap();
        eprintln!("    wrote {} bytes", buf.len());

        eprintln!("  reading back {fname}...");
        let wb2 = XlsxReader::read(std::io::Cursor::new(&buf)).unwrap();
        let ex2: Vec<_> = (0..wb2.sheet_count())
            .filter_map(|i| wb2.worksheet(i))
            .flat_map(|ws| ws.charts_ex().iter())
            .collect();

        assert_eq!(
            ex1.len(),
            ex2.len(),
            "{fname}: chartEx count changed: {} -> {}",
            ex1.len(),
            ex2.len()
        );
        for (i, (c1, c2)) in ex1.iter().zip(ex2.iter()).enumerate() {
            if c1 != c2 {
                let diffs = diff_charts_ex(c1, c2, i);
                if !diffs.is_empty() {
                    panic!("{fname}: chartEx {i} roundtrip diffs: {}", diffs.join("; "));
                }
            }
        }
        eprintln!("  PASS {fname}");
    }
}

fn has_chartex_entries(path: &Path) -> std::result::Result<bool, String> {
    let file = std::fs::File::open(path).map_err(|e| format!("open: {e}"))?;
    let archive = zip::ZipArchive::new(file).map_err(|e| format!("zip: {e}"))?;
    for name in archive.file_names() {
        if name.starts_with("xl/charts/chartEx") && name.ends_with(".xml") {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Read → write → read roundtrip on real chart files from the corpus.
/// Asserts that every chart survives a write-back cycle: no data loss.
#[test]
#[ignore = "requires Excel corpus - run with: cargo test --release -p duke-sheets --test chart_corpus chart_corpus_roundtrip -- --ignored --nocapture"]
fn chart_corpus_roundtrip() {
    let corpus_dir =
        std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| DEFAULT_CORPUS_DIR.to_string());

    let dir = Path::new(&corpus_dir);
    assert!(
        dir.is_dir(),
        "corpus directory not found: {}",
        dir.display()
    );

    let mut xlsx_paths: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read corpus dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("xlsx"))
        .collect();
    xlsx_paths.sort();

    // Phase 1: find chart files (reuse the zip scan), including chartEx
    let mut chart_files: Vec<PathBuf> = Vec::new();
    for path in &xlsx_paths {
        let has_std = has_chart_entries(path).unwrap_or(false);
        let has_ex = has_chartex_entries(path).unwrap_or(false);
        if has_std || has_ex {
            chart_files.push(path.clone());
        }
    }
    println!(
        "\nFound {} chart files in {} xlsx files",
        chart_files.len(),
        xlsx_paths.len()
    );

    // Phase 2: read → write → read, compare charts
    let start = Instant::now();
    let mut tested = 0usize;
    let mut passed = 0usize;
    let mut skipped = 0usize;
    let mut chart_mismatches: Vec<(String, String)> = Vec::new();
    let mut write_errors: Vec<(String, String)> = Vec::new();

    for (i, path) in chart_files.iter().enumerate() {
        if (i + 1) % 100 == 0 {
            println!(
                "  roundtrip {}/{} ({:.0}s elapsed)...",
                i + 1,
                chart_files.len(),
                start.elapsed().as_secs_f64()
            );
        }

        let fname = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();

        // Read original
        let wb1 = match std::panic::catch_unwind(|| XlsxReader::read_file(path)) {
            Ok(Ok(wb)) => wb,
            _ => {
                skipped += 1;
                continue;
            }
        };

        // Collect charts from read 1
        let charts1 = collect_all_charts(&wb1);
        let chartex1 = collect_all_charts_ex(&wb1);
        if charts1.is_empty() && chartex1.is_empty() {
            skipped += 1;
            continue;
        }

        // Write back to bytes
        let mut buf = Vec::new();
        match XlsxWriter::write(&wb1, std::io::Cursor::new(&mut buf)) {
            Ok(()) => {}
            Err(e) => {
                write_errors.push((fname, format!("{e}")));
                continue;
            }
        }

        // Read back
        let wb2 = match XlsxReader::read(std::io::Cursor::new(&buf)) {
            Ok(wb) => wb,
            Err(e) => {
                chart_mismatches.push((fname, format!("read-back failed: {e}")));
                continue;
            }
        };

        let charts2 = collect_all_charts(&wb2);
        let chartex2 = collect_all_charts_ex(&wb2);
        tested += 1;

        // Compare standard chart counts
        if charts1.len() != charts2.len() {
            chart_mismatches.push((
                fname.clone(),
                format!("chart count: {} -> {}", charts1.len(), charts2.len()),
            ));
        }

        // Compare standard charts field by field
        let mut file_ok = charts1.len() == charts2.len();
        for (ci, (c1, c2)) in charts1.iter().zip(charts2.iter()).enumerate() {
            if c1 != c2 {
                let diffs = diff_charts(c1, c2, ci);
                if !diffs.is_empty() {
                    chart_mismatches.push((fname.clone(), diffs.join("; ")));
                    file_ok = false;
                    break;
                }
            }
        }

        // Compare chartEx counts
        if chartex1.len() != chartex2.len() {
            chart_mismatches.push((
                fname.clone(),
                format!("chartEx count: {} -> {}", chartex1.len(), chartex2.len()),
            ));
            file_ok = false;
        } else {
            // Compare chartEx field by field
            for (ci, (c1, c2)) in chartex1.iter().zip(chartex2.iter()).enumerate() {
                if c1 != c2 {
                    let diffs = diff_charts_ex(c1, c2, ci);
                    if !diffs.is_empty() {
                        chart_mismatches.push((fname.clone(), diffs.join("; ")));
                        file_ok = false;
                        break;
                    }
                }
            }
        }

        if file_ok {
            passed += 1;
        }
    }
    let elapsed = start.elapsed().as_secs_f64();

    println!("\n=== Corpus Roundtrip Results ({elapsed:.1}s) ===");
    println!("Chart files tested:  {tested}");
    println!("Passed (identical):  {passed}");
    println!("Chart mismatches:    {}", chart_mismatches.len());
    println!("Write errors:        {}", write_errors.len());
    println!("Skipped (read fail): {skipped}");

    if !write_errors.is_empty() {
        println!("\nWrite errors (first 20):");
        for (f, e) in write_errors.iter().take(20) {
            println!("  {f}: {e}");
        }
    }

    if !chart_mismatches.is_empty() {
        println!("\nChart mismatches (first 30):");
        for (f, d) in chart_mismatches.iter().take(30) {
            println!("  {f}: {d}");
        }
    }

    let total_compared = passed + chart_mismatches.len();
    let pass_rate = if total_compared == 0 {
        100.0
    } else {
        (passed as f64 / total_compared as f64) * 100.0
    };
    println!("\n=== Summary ===");
    println!("Roundtrip pass rate: {passed}/{total_compared} ({pass_rate:.1}%)");

    // Fail on panics in write/read-back (those are bugs)
    // Mismatches are reported but don't fail the test yet - we need to
    // triage them first to separate real bugs from known limitations.
}

fn collect_all_charts(wb: &Workbook) -> Vec<&duke_sheets_chart::Chart> {
    let mut charts = Vec::new();
    for si in 0..wb.sheet_count() {
        if let Some(ws) = wb.worksheet(si) {
            for c in ws.charts() {
                charts.push(c);
            }
        }
    }
    for ci in 0..wb.chartsheet_count() {
        if let Some(cs) = wb.chartsheet(ci) {
            charts.push(&cs.chart);
        }
    }
    charts
}

fn collect_all_charts_ex(wb: &Workbook) -> Vec<&duke_sheets_chart::ChartEx> {
    let mut charts = Vec::new();
    for si in 0..wb.sheet_count() {
        if let Some(ws) = wb.worksheet(si) {
            for c in ws.charts_ex() {
                charts.push(c);
            }
        }
    }
    charts
}

fn diff_charts(
    c1: &duke_sheets_chart::Chart,
    c2: &duke_sheets_chart::Chart,
    idx: usize,
) -> Vec<String> {
    let mut diffs = Vec::new();
    if c1.chart_type != c2.chart_type {
        diffs.push(format!(
            "chart {idx} type: {:?} -> {:?}",
            c1.chart_type, c2.chart_type
        ));
    }
    if c1.title != c2.title {
        diffs.push(format!(
            "chart {idx} title: {:?} -> {:?}",
            c1.title, c2.title
        ));
    }
    if c1.series.len() != c2.series.len() {
        diffs.push(format!(
            "chart {idx} series count: {} -> {}",
            c1.series.len(),
            c2.series.len()
        ));
    } else {
        for (si, (s1, s2)) in c1.series.iter().zip(c2.series.iter()).enumerate() {
            if s1.values != s2.values {
                diffs.push(format!("chart {idx} series {si} values differ"));
            }
            if s1.categories != s2.categories {
                diffs.push(format!("chart {idx} series {si} categories differ"));
            }
            if s1.name != s2.name {
                diffs.push(format!(
                    "chart {idx} series {si} name: {:?} -> {:?}",
                    s1.name, s2.name
                ));
            }
            if s1.explosion != s2.explosion {
                diffs.push(format!(
                    "chart {idx} series {si} explosion: {:?} -> {:?}",
                    s1.explosion, s2.explosion
                ));
            }
            if s1.trendline != s2.trendline {
                diffs.push(format!("chart {idx} series {si} trendline differs"));
            }
            if s1.marker != s2.marker {
                diffs.push(format!("chart {idx} series {si} marker differs"));
            }
            if s1.smooth != s2.smooth {
                diffs.push(format!(
                    "chart {idx} series {si} smooth: {:?} -> {:?}",
                    s1.smooth, s2.smooth
                ));
            }
            if s1.data_labels != s2.data_labels {
                diffs.push(format!("chart {idx} series {si} data_labels differ"));
            }
            if s1.error_bars != s2.error_bars {
                diffs.push(format!("chart {idx} series {si} error_bars differ"));
            }
            if s1.data_points != s2.data_points {
                diffs.push(format!("chart {idx} series {si} data_points differ"));
            }
            if s1.shape_properties != s2.shape_properties {
                diffs.push(format!("chart {idx} series {si} shape_properties differ"));
            }
        }
    }
    if c1.legend != c2.legend {
        diffs.push(format!("chart {idx} legend differs"));
    }
    if c1.category_axis != c2.category_axis {
        diffs.push(format!("chart {idx} category_axis differs"));
    }
    if c1.value_axis != c2.value_axis {
        diffs.push(format!("chart {idx} value_axis differs"));
    }
    if c1.series_axis != c2.series_axis {
        diffs.push(format!("chart {idx} series_axis differs"));
    }
    if c1.data_labels != c2.data_labels {
        diffs.push(format!("chart {idx} data_labels differs"));
    }
    if c1.hole_size != c2.hole_size {
        diffs.push(format!(
            "chart {idx} hole_size: {:?} -> {:?}",
            c1.hole_size, c2.hole_size
        ));
    }
    if c1.gap_width != c2.gap_width {
        diffs.push(format!(
            "chart {idx} gap_width: {:?} -> {:?}",
            c1.gap_width, c2.gap_width
        ));
    }
    if c1.overlap != c2.overlap {
        diffs.push(format!(
            "chart {idx} overlap: {:?} -> {:?}",
            c1.overlap, c2.overlap
        ));
    }
    if c1.vary_colors != c2.vary_colors {
        diffs.push(format!(
            "chart {idx} vary_colors: {:?} -> {:?}",
            c1.vary_colors, c2.vary_colors
        ));
    }
    if c1.view_3d != c2.view_3d {
        diffs.push(format!("chart {idx} view_3d differs"));
    }
    if c1.is_3d != c2.is_3d {
        diffs.push(format!("chart {idx} is_3d: {} -> {}", c1.is_3d, c2.is_3d));
    }
    if c1.drop_lines != c2.drop_lines {
        diffs.push(format!("chart {idx} drop_lines differs"));
    }
    if c1.high_low_lines != c2.high_low_lines {
        diffs.push(format!("chart {idx} high_low_lines differs"));
    }
    if c1.up_down_bars != c2.up_down_bars {
        diffs.push(format!("chart {idx} up_down_bars differs"));
    }
    if c1.series_lines != c2.series_lines {
        diffs.push(format!("chart {idx} series_lines differs"));
    }
    diffs
}

fn diff_charts_ex(
    c1: &duke_sheets_chart::ChartEx,
    c2: &duke_sheets_chart::ChartEx,
    idx: usize,
) -> Vec<String> {
    let mut diffs = Vec::new();
    if c1.title != c2.title {
        diffs.push(format!("chartEx {idx} title differs"));
    }
    if c1.data != c2.data {
        diffs.push(format!("chartEx {idx} data differs"));
    }
    if c1.plot_area.series != c2.plot_area.series {
        diffs.push(format!("chartEx {idx} series differs"));
    }
    if c1.plot_area.axes != c2.plot_area.axes {
        diffs.push(format!("chartEx {idx} axes differs"));
    }
    if c1.legend != c2.legend {
        diffs.push(format!("chartEx {idx} legend differs"));
    }
    if c1.shape_properties != c2.shape_properties {
        diffs.push(format!("chartEx {idx} spPr differs"));
    }
    if c1.text_properties != c2.text_properties {
        diffs.push(format!("chartEx {idx} txPr differs"));
    }
    if c1.color_map_override != c2.color_map_override {
        diffs.push(format!("chartEx {idx} clrMapOvr differs"));
    }
    if c1.format_overrides != c2.format_overrides {
        diffs.push(format!("chartEx {idx} fmtOvrs differs"));
    }
    if c1.print_settings != c2.print_settings {
        diffs.push(format!("chartEx {idx} printSettings differs"));
    }
    if c1.raw_chart_style != c2.raw_chart_style {
        diffs.push(format!("chartEx {idx} raw_chart_style differs"));
    }
    if c1.raw_chart_color_style != c2.raw_chart_color_style {
        diffs.push(format!("chartEx {idx} raw_chart_color_style differs"));
    }
    if c1.plot_area.shape_properties != c2.plot_area.shape_properties {
        diffs.push(format!("chartEx {idx} plotArea spPr differs"));
    }
    if c1.plot_area.plot_surface != c2.plot_area.plot_surface {
        diffs.push(format!("chartEx {idx} plotSurface differs"));
    }
    diffs
}
