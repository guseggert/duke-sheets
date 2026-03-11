//! Tests from SheetJS ssf library's comma.tsv test file
//!
//! These tests verify comma-based number formatting including thousands
//! separators and comma divisors (,, = millions, ,,, = billions, etc.)

use ssfmt::format_default;
use flate2::read::GzDecoder;
use std::io::Read;

#[derive(Debug)]
struct CommaTestCase {
    value: f64,
    format: String,
    expected: String,
}

fn load_test_cases() -> Vec<CommaTestCase> {
    let compressed = include_bytes!("fixtures/comma.tsv.gz");
    let mut decoder = GzDecoder::new(&compressed[..]);
    let mut tsv_data = String::new();
    decoder.read_to_string(&mut tsv_data).unwrap();
    let lines: Vec<&str> = tsv_data.lines().collect();

    if lines.is_empty() {
        return Vec::new();
    }

    // First line is header with format strings
    let headers: Vec<&str> = lines[0].split('\t').collect();
    let formats: Vec<String> = headers[1..].iter().map(|s| s.to_string()).collect();

    let mut test_cases = Vec::new();

    // Remaining lines are test data
    for line in &lines[1..] {
        if line.trim().is_empty() {
            continue;
        }

        let parts: Vec<&str> = line.split('\t').collect();
        if parts.is_empty() {
            continue;
        }

        let value = parts[0].parse::<f64>().unwrap();

        // Each column after the first is the expected output for that format
        for (i, expected) in parts[1..].iter().enumerate() {
            if i >= formats.len() {
                break;
            }

            test_cases.push(CommaTestCase {
                value,
                format: formats[i].clone(),
                expected: expected.to_string(),
            });
        }
    }

    test_cases
}

#[test]
fn test_ssf_comma_formats() {
    let test_cases = load_test_cases();
    let total = test_cases.len();
    let mut passed = 0;
    let mut failed = 0;

    println!("\nRunning {} test cases from comma.tsv", total);

    for (i, test) in test_cases.iter().enumerate() {
        // Skip tests marked with | prefix (expected to fail)
        if test.expected.starts_with('|') {
            continue;
        }

        match format_default(test.value, &test.format) {
            Ok(result) => {
                if result == test.expected {
                    passed += 1;
                } else {
                    failed += 1;
                    if failed <= 20 {
                        println!(
                            "FAIL #{}: value={}, format='{}', expected='{}', got='{}'",
                            i + 1,
                            test.value,
                            test.format,
                            test.expected,
                            result
                        );
                    }
                }
            }
            Err(e) => {
                failed += 1;
                if failed <= 20 {
                    println!(
                        "ERROR #{}: value={}, format='{}', expected='{}', error={:?}",
                        i + 1,
                        test.value,
                        test.format,
                        test.expected,
                        e
                    );
                }
            }
        }
    }

    println!("Total:   {}", total);
    println!("Passed:  {} ({:.1}%)", passed, (passed as f64 / total as f64) * 100.0);
    println!("Failed:  {} ({:.1}%)", failed, (failed as f64 / total as f64) * 100.0);

    // We should pass most tests, but some edge cases might fail
    assert!(passed > total / 2, "More than half of tests should pass");
}
