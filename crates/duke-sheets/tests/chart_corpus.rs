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

/// Above this output size the writer-property checks are skipped unless
/// `DUKE_CORPUS_DEEP=1` is set. See `analyze_one`.
const DEEP_CHECK_MAX_BYTES: u64 = 4 * 1024 * 1024;

fn deep_checks_forced() -> bool {
    std::env::var("DUKE_CORPUS_DEEP").is_ok_and(|v| v != "0")
}

/// The corpus directory, from `DUKE_CORPUS_DIR` or the default.
fn corpus_dir() -> PathBuf {
    PathBuf::from(
        std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| DEFAULT_CORPUS_DIR.to_string()),
    )
}

/// Every .xlsx in the corpus, sorted, optionally narrowed to a
/// reproducible sample.
///
/// `DUKE_CORPUS_LIMIT` takes a deterministic spread across the sorted
/// list rather than a prefix, because file names are grouped by origin
/// and a prefix would sample one producer. `DUKE_CORPUS_SKIP` starts the
/// spread at a different offset, so several runs can cover disjoint
/// slices.
fn corpus_files(dir: &Path) -> Vec<PathBuf> {
    assert!(
        dir.is_dir(),
        "corpus directory not found: {}. Set DUKE_CORPUS_DIR.",
        dir.display()
    );
    let mut paths: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read corpus dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("xlsx"))
        .collect();
    paths.sort();

    let limit = env_usize("DUKE_CORPUS_LIMIT");
    let skip = env_usize("DUKE_CORPUS_SKIP").unwrap_or(0);
    match limit {
        Some(limit) if limit < paths.len() => {
            let stride = paths.len() / limit.max(1);
            paths
                .into_iter()
                .skip(skip)
                .step_by(stride.max(1))
                .take(limit)
                .collect()
        }
        _ => paths.into_iter().skip(skip).collect(),
    }
}

fn env_usize(name: &str) -> Option<usize> {
    std::env::var(name).ok().and_then(|v| v.parse().ok())
}

/// Files whose zip holds a chart part of either kind. The scan is the
/// cheap half of every corpus test, so it runs in parallel.
fn corpus_chart_files(dir: &Path) -> &'static [PathBuf] {
    use rayon::prelude::*;
    static FILES: std::sync::OnceLock<Vec<PathBuf>> = std::sync::OnceLock::new();
    FILES.get_or_init(|| {
        let scan = Instant::now();
        let mut files: Vec<(u64, PathBuf)> = corpus_files(dir)
            .into_par_iter()
            .filter(|p| {
                has_chart_entries(p).unwrap_or(false) || has_chartex_entries(p).unwrap_or(false)
            })
            .map(|p| (std::fs::metadata(&p).map(|m| m.len()).unwrap_or(0), p))
            .collect();
        // Longest job first. One 95 MB workbook takes longer than the
        // other five hundred put together, so if it is picked up last the
        // whole run waits on it; started first, it runs under everything
        // else. Ordering by size is the cheap proxy for duration.
        files.sort_by(|a, b| b.0.cmp(&a.0).then_with(|| a.1.cmp(&b.1)));
        println!(
            "  scanned corpus in {:.1}s, {} chart file(s)",
            scan.elapsed().as_secs_f64(),
            files.len()
        );
        files.into_iter().map(|(_, p)| p).collect()
    })
}

/// Everything one file has to say, computed once and shared by every
/// test in this binary. The pipeline below is the expensive part of the
/// suite, and running it per test would repeat it.
struct Analysis {
    name: String,
    outcome: Outcome,
    /// Whether the writer-property checks ran, which they do not for
    /// files above `DEEP_CHECK_MAX_BYTES`.
    deep: bool,
    /// Divergence between the first and second save, anywhere in the
    /// package, if any.
    package_drift: Option<String>,
    secs: f64,
}

fn analyses() -> &'static [Analysis] {
    use rayon::prelude::*;
    static RESULTS: std::sync::OnceLock<Vec<Analysis>> = std::sync::OnceLock::new();
    RESULTS.get_or_init(|| {
        let files = corpus_chart_files(&corpus_dir());
        let start = Instant::now();
        let done = std::sync::atomic::AtomicUsize::new(0);
        let total = files.len();
        let out: Vec<Analysis> = files
            .par_iter()
            // Split to single files. The default adaptive split hands a
            // worker a contiguous run, and since the list is ordered by
            // size that run is all giants or all minnows; per-file
            // granularity lets idle threads steal the remaining big ones.
            .with_min_len(1)
            .map(|path| {
                let result = analyze_one(path);
                let n = done.fetch_add(1, std::sync::atomic::Ordering::Relaxed) + 1;
                if n % 100 == 0 {
                    println!(
                        "  analysed {n}/{total} ({:.0}s elapsed)...",
                        start.elapsed().as_secs_f64()
                    );
                }
                result
            })
            .collect();
        let wall = start.elapsed().as_secs_f64();
        let cpu: f64 = out.iter().map(|r| r.secs).sum();
        let slowest = out.iter().map(|r| r.secs).fold(0.0, f64::max);
        println!(
            "  analysed {total} file(s) in {wall:.1}s wall, {cpu:.1}s cpu \
             ({:.1}x parallel, slowest file {slowest:.1}s)",
            cpu / wall.max(0.001)
        );
        out
    })
}

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
            for chart in ws.charts().map(|drawn| drawn.payload) {
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
                for cx in ws.charts_ex().map(|drawn| drawn.payload) {
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
            .flat_map(|ws| ws.charts_ex().map(|drawn| drawn.payload))
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
            .flat_map(|ws| ws.charts_ex().map(|drawn| drawn.payload))
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
                let mut diffs = diff_charts_ex(c1, c2, i);
                if diffs.is_empty() {
                    diffs = debug_divergence(c1, c2);
                }
                panic!("{fname}: chartEx {i} roundtrip diffs: {}", diffs.join("; "));
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
    let results = analyses();

    let mut passed = 0usize;
    let mut skipped = 0usize;
    let mut tested = 0usize;
    let mut mismatches: Vec<(&str, &str)> = Vec::new();
    let mut write_errors: Vec<(&str, &str)> = Vec::new();
    let mut slow: Vec<(&str, f64)> = Vec::new();
    for result in results {
        if result.secs > SLOW_THRESHOLD_SECS {
            slow.push((&result.name, result.secs));
        }
        match &result.outcome {
            Outcome::Skipped => skipped += 1,
            Outcome::WriteError(e) => {
                tested += 1;
                write_errors.push((&result.name, e));
            }
            Outcome::Mismatch(d) => {
                tested += 1;
                mismatches.push((&result.name, d));
            }
            Outcome::Passed => {
                tested += 1;
                passed += 1;
            }
        }
    }

    println!("\n=== Corpus Roundtrip Results ===");
    println!("Chart files tested:  {tested}");
    println!("Passed (identical):  {passed}");
    println!("Chart mismatches:    {}", mismatches.len());
    println!("Write errors:        {}", write_errors.len());
    println!("Skipped (read fail): {skipped}");
    let deep = results.iter().filter(|r| r.deep).count();
    println!(
        "Writer-property checks (determinism, fixed point): {deep}/{} \
         (the rest exceed {} MiB; set DUKE_CORPUS_DEEP=1 to include them)",
        results.len(),
        DEEP_CHECK_MAX_BYTES / (1024 * 1024)
    );

    if !write_errors.is_empty() {
        println!("\nWrite errors (first 20):");
        for (f, e) in write_errors.iter().take(20) {
            println!("  {f}: {e}");
        }
    }
    if !mismatches.is_empty() {
        println!("\nMismatches (first 30):");
        for (f, d) in mismatches.iter().take(30) {
            println!("  {f}: {d}");
        }
    }
    if !slow.is_empty() {
        slow.sort_by(|a, b| b.1.total_cmp(&a.1));
        println!("\nSlow files (>{SLOW_THRESHOLD_SECS}s):");
        for (f, secs) in slow.iter().take(10) {
            println!("  {f}: {secs:.1}s");
        }
    }

    let compared = passed + mismatches.len();
    let rate = if compared == 0 {
        100.0
    } else {
        (passed as f64 / compared as f64) * 100.0
    };
    println!("\n=== Summary ===");
    println!("Roundtrip pass rate: {passed}/{compared} ({rate:.1}%)");

    // A mismatch is data the writer lost or changed, and a write error is
    // a chart we can read but not emit. Both are bugs, so both fail here;
    // read failures are a property of the input, not of us, and only the
    // counts above report them.
    assert!(
        write_errors.is_empty(),
        "{} chart file(s) failed to write",
        write_errors.len()
    );
    assert!(
        mismatches.is_empty(),
        "{} chart file(s) changed across a round trip",
        mismatches.len()
    );
}

/// Opening and saving twice must produce the same document as saving
/// once. Anything else means a save is still changing the file, so a
/// user who saves twice gets a third document.
///
/// This is deliberately whole-package rather than chart-only: the
/// invariant needs no expected output, so it finds losses anywhere,
/// including in parts no corpus test targets.
#[test]
#[ignore = "requires Excel corpus"]
fn chart_corpus_package_fixed_point() {
    let results = analyses();
    let offenders: Vec<(&str, &str)> = results
        .iter()
        .filter_map(|r| r.package_drift.as_deref().map(|d| (r.name.as_str(), d)))
        .collect();

    // Group by part so the report names causes, not files.
    let mut by_part: HashMap<&str, (usize, &str)> = HashMap::new();
    for (file, detail) in &offenders {
        let part = detail.split_whitespace().next().unwrap_or("?");
        let entry = by_part.entry(part).or_insert((0, ""));
        entry.0 += 1;
        if entry.1.is_empty() {
            entry.1 = detail;
        }
        let _ = file;
    }
    let mut ranked: Vec<_> = by_part.into_iter().collect();
    ranked.sort_by(|a, b| b.1 .0.cmp(&a.1 .0));

    println!("\n=== Fixed-point failures by part ===");
    for (part, (count, example)) in &ranked {
        println!("  {count:5}  {part}");
        println!("         e.g. {}", &example[..example.len().min(300)]);
    }
    let covered = results.iter().filter(|r| r.deep).count();
    println!(
        "\n{} of {covered} checked files are not a fixed point ({} skipped as too large)",
        offenders.len(),
        results.len() - covered
    );

    // Four parts are known not to be fixed points yet, each for its own
    // reason, so they are listed rather than left to fail the run. The
    // ranking above is the work queue; the assertion below is the ratchet,
    // and any other part that starts drifting fails immediately.
    //
    //   xl/styles.xml            fills and formats are pooled through a
    //                            hash map, so their order - and count -
    //                            comes out differently on a second pass.
    //   xl/worksheets/sheetN.xml data validations are lost: a sheet
    //                            written with 25 comes back with 2.
    //   xl/drawings/drawingN.vml comment box geometry shifts by a fraction
    //   xl/drawings/vmlDrawingN.vml of a point per save, so boxes creep.
    const KNOWN_DRIFT: &[&str] = &[
        "xl/styles.xml",
        "xl/worksheets/sheetN.xml",
        "xl/drawings/drawingN.xml",
        "xl/drawings/vmlDrawingN.vml",
    ];
    let unexpected: Vec<&(&str, (usize, &str))> = ranked
        .iter()
        .filter(|(part, _)| !KNOWN_DRIFT.contains(&normalise_part(part).as_str()))
        .collect();
    assert!(
        unexpected.is_empty(),
        "part(s) outside the known set stopped being a fixed point: {:?}",
        unexpected
            .iter()
            .map(|(part, (count, _))| format!("{part} ({count} files)"))
            .collect::<Vec<_>>()
    );
}

/// Collapse the index out of a part name, so sheet3.xml and sheet17.xml
/// are the same entry in a list of known problems.
fn normalise_part(part: &str) -> String {
    let mut out = String::with_capacity(part.len());
    let mut in_digits = false;
    for c in part.chars() {
        if c.is_ascii_digit() {
            if !in_digits {
                out.push('N');
                in_digits = true;
            }
        } else {
            in_digits = false;
            out.push(c);
        }
    }
    out
}

enum Outcome {
    Skipped,
    WriteError(String),
    Mismatch(String),
    Passed,
}

/// Read, write, read again, and check three things:
///
/// 1. every chart we modelled survives the write unchanged;
/// 2. writing the same model twice produces the same bytes, so nothing
///    leaks iteration order or a clock into the output;
/// 3. writing what we read back reproduces the first output byte for
///    byte. A user who opens and saves twice must not get a third
///    document, and this is the only check that notices a pipeline that
///    oscillates or degrades a little on each pass.
///
/// The first is what the corpus is for and runs on every file. The other
/// two are properties of the writer rather than of any one document, and
/// each costs another full write, so above `DEEP_CHECK_MAX_BYTES` they
/// are skipped: a handful of very large workbooks otherwise account for
/// most of the suite's runtime while testing nothing the small ones do
/// not. `DUKE_CORPUS_DEEP=1` runs them on everything.
fn analyze_one(path: &Path) -> Analysis {
    let name = path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();
    let start = Instant::now();
    let deep_cell = std::cell::Cell::new(false);
    let finish = |outcome, package_drift| Analysis {
        name: name.clone(),
        outcome,
        deep: deep_cell.get(),
        package_drift,
        secs: start.elapsed().as_secs_f64(),
    };

    let wb1 = match std::panic::catch_unwind(|| XlsxReader::read_file(path)) {
        Ok(Ok(wb)) => wb,
        _ => return finish(Outcome::Skipped, None),
    };
    if collect_all_charts(&wb1).is_empty() && collect_all_charts_ex(&wb1).is_empty() {
        return finish(Outcome::Skipped, None);
    }

    let first = match write_to_bytes(&wb1) {
        Ok(bytes) => bytes,
        Err(e) => return finish(Outcome::WriteError(e), None),
    };

    let deep = first.len() as u64 <= DEEP_CHECK_MAX_BYTES || deep_checks_forced();
    deep_cell.set(deep);
    if deep {
        match write_to_bytes(&wb1) {
            Ok(again) if again != first => {
                return finish(
                    Outcome::Mismatch(format!(
                        "writing the same model twice differed ({} vs {} bytes)",
                        first.len(),
                        again.len()
                    )),
                    None,
                );
            }
            Err(e) => return finish(Outcome::WriteError(format!("second write: {e}")), None),
            _ => {}
        }
    }

    let wb2 = match XlsxReader::read(std::io::Cursor::new(&first)) {
        Ok(wb) => wb,
        Err(e) => return finish(Outcome::Mismatch(format!("read-back failed: {e}")), None),
    };

    if let Some(diff) = compare_charts(&wb1, &wb2) {
        return finish(Outcome::Mismatch(diff), None);
    }

    if !deep {
        return finish(Outcome::Passed, None);
    }

    let second = match write_to_bytes(&wb2) {
        Ok(bytes) => bytes,
        Err(e) => return finish(Outcome::WriteError(format!("re-write: {e}")), None),
    };
    // Chart parts must be a fixed point byte for byte. Whole packages are
    // not yet, so the divergence elsewhere is reported rather than
    // asserted; see chart_corpus_package_fixed_point.
    let package_drift = differing_part(&first, &second, |_| true);
    if let Some(part) = differing_part(&first, &second, |name| name.starts_with("xl/charts/")) {
        return finish(
            Outcome::Mismatch(format!("chart part is not a fixed point: {part}")),
            package_drift,
        );
    }

    finish(Outcome::Passed, package_drift)
}

fn write_to_bytes(wb: &Workbook) -> std::result::Result<Vec<u8>, String> {
    let mut buf = Vec::new();
    XlsxWriter::write(wb, std::io::Cursor::new(&mut buf)).map_err(|e| e.to_string())?;
    Ok(buf)
}

/// The first part of two packages that differs, restricted to the parts
/// `keep` selects, so a byte mismatch names a part instead of a size.
fn differing_part(a: &[u8], b: &[u8], keep: impl Fn(&str) -> bool) -> Option<String> {
    use std::io::Read;
    let mut za = zip::ZipArchive::new(std::io::Cursor::new(a)).ok()?;
    let mut zb = zip::ZipArchive::new(std::io::Cursor::new(b)).ok()?;
    let names_a: Vec<String> = za
        .file_names()
        .filter(|n| keep(n))
        .map(str::to_string)
        .collect();
    let names_b: Vec<String> = zb
        .file_names()
        .filter(|n| keep(n))
        .map(str::to_string)
        .collect();
    for name in &names_a {
        if !names_b.contains(name) {
            return Some(format!("{name} is missing from the second write"));
        }
    }
    for name in &names_b {
        if !names_a.contains(name) {
            return Some(format!("{name} appeared only in the second write"));
        }
    }
    for name in &names_a {
        let mut da = Vec::new();
        let mut db = Vec::new();
        za.by_name(name).ok()?.read_to_end(&mut da).ok()?;
        zb.by_name(name).ok()?.read_to_end(&mut db).ok()?;
        if da != db {
            let at = da
                .iter()
                .zip(db.iter())
                .position(|(x, y)| x != y)
                .unwrap_or(da.len().min(db.len()));
            let window = |v: &[u8]| {
                let lo = at.saturating_sub(40);
                String::from_utf8_lossy(&v[lo..(at + 60).min(v.len())]).into_owned()
            };
            return Some(format!(
                "{name} ({} vs {} bytes) at offset {at}: {:?} vs {:?}",
                da.len(),
                db.len(),
                window(&da),
                window(&db)
            ));
        }
    }
    None
}

/// Compare every chart of both kinds, naming the first difference.
fn compare_charts(wb1: &Workbook, wb2: &Workbook) -> Option<String> {
    let charts1 = collect_all_charts(wb1);
    let charts2 = collect_all_charts(wb2);
    if charts1.len() != charts2.len() {
        return Some(format!(
            "chart count: {} -> {}",
            charts1.len(),
            charts2.len()
        ));
    }
    for (ci, (c1, c2)) in charts1.iter().zip(charts2.iter()).enumerate() {
        if c1 != c2 {
            let mut diffs = diff_charts(c1, c2, ci);
            if diffs.is_empty() {
                diffs = debug_divergence(c1, c2)
                    .into_iter()
                    .map(|d| format!("chart {ci} {d}"))
                    .collect();
            }
            return Some(diffs.join("; "));
        }
    }

    let ex1 = collect_all_charts_ex(wb1);
    let ex2 = collect_all_charts_ex(wb2);
    if ex1.len() != ex2.len() {
        return Some(format!("chartEx count: {} -> {}", ex1.len(), ex2.len()));
    }
    for (ci, (c1, c2)) in ex1.iter().zip(ex2.iter()).enumerate() {
        if c1 != c2 {
            let mut diffs = diff_charts_ex(c1, c2, ci);
            if diffs.is_empty() {
                diffs = debug_divergence(c1, c2)
                    .into_iter()
                    .map(|d| format!("chartEx {ci} {d}"))
                    .collect();
            }
            return Some(diffs.join("; "));
        }
    }
    None
}

fn collect_all_charts(wb: &Workbook) -> Vec<&duke_sheets_chart::Chart> {
    let mut charts = Vec::new();
    for si in 0..wb.sheet_count() {
        if let Some(ws) = wb.worksheet(si) {
            for c in ws.charts() {
                charts.push(c.payload);
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
                charts.push(c.payload);
            }
        }
    }
    charts
}

/// Name the first place two values' pretty-Debug renderings diverge.
///
/// The hand-written differs below only report fields they were taught
/// about, so a model that gains a field silently stops being compared.
/// This is the backstop: it needs no maintenance and cannot come up
/// empty for values that are genuinely unequal.
fn debug_divergence<T: std::fmt::Debug>(a: &T, b: &T) -> Vec<String> {
    let (left, right) = (format!("{a:#?}"), format!("{b:#?}"));
    if left == right {
        // Debug agrees but PartialEq does not, which is what a NaN does.
        return vec!["differs in a value that is not equal to itself (NaN?)".to_string()];
    }
    let mut left_lines = left.lines();
    let mut right_lines = right.lines();
    let mut field = "<root>".to_string();
    let mut line_no = 0usize;
    loop {
        match (left_lines.next(), right_lines.next()) {
            (Some(l), Some(r)) => {
                line_no += 1;
                // Track the innermost named field seen so far, so the
                // report points at a field rather than a line number.
                if let Some((name, _)) = l.trim().split_once(':') {
                    if !name.is_empty() && name.chars().all(|c| c.is_alphanumeric() || c == '_') {
                        field = name.to_string();
                    }
                }
                if l != r {
                    return vec![format!(
                        "at {field} (line {line_no}): {} -> {}",
                        l.trim(),
                        r.trim()
                    )];
                }
            }
            (None, None) => return vec!["differs in an unnamed field".to_string()],
            (l, r) => {
                return vec![format!(
                    "structure length differs at line {line_no}: {:?} -> {:?}",
                    l.map(str::trim),
                    r.map(str::trim)
                )]
            }
        }
    }
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
    // Excel requires every chartEx to carry chart style and chart colour
    // style parts, so the writer generates defaults when the source had
    // none; a None -> generated-default transition is expected on round
    // trip, not a loss.
    let style_defaulted = c1.style.is_none()
        && c2.style
            == Some(duke_sheets_chart::ChartStylePart::Typed(Box::new(
                duke_sheets_chart::ChartStyle::default(),
            )));
    if !style_defaulted && c1.style != c2.style {
        diffs.push(format!("chartEx {idx} style differs"));
    }
    let colors_defaulted = c1.color_style.is_none()
        && c2.color_style
            == Some(duke_sheets_chart::ChartColorStylePart::Typed(
                duke_sheets_chart::ChartColorStyle::default(),
            ));
    if !colors_defaulted && c1.color_style != c2.color_style {
        diffs.push(format!("chartEx {idx} color_style differs"));
    }
    if c1.plot_area.shape_properties != c2.plot_area.shape_properties {
        diffs.push(format!("chartEx {idx} plotArea spPr differs"));
    }
    if c1.plot_area.plot_surface != c2.plot_area.plot_surface {
        diffs.push(format!("chartEx {idx} plotSurface differs"));
    }
    diffs
}
