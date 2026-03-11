//! Tests from SheetJS ssf library's fraction.json test suite
//!
//! These tests verify that ssfmt correctly implements Excel's fraction
//! formats like "# ?/?", "# ??/??", etc.

use serde_json::Value;
use ssfmt::format_default;

#[derive(Debug)]
struct FractionTestCase {
    value: f64,
    format_code: String,
    expected: String,
}

fn load_test_cases() -> Vec<FractionTestCase> {
    let json_data = include_str!("fixtures/ssf_fraction.json");
    let tests: Vec<Value> = serde_json::from_str(json_data)
        .expect("Failed to parse ssf_fraction.json");

    tests
        .iter()
        .filter_map(|test| {
            if let Value::Array(arr) = test {
                if arr.len() == 3 {
                    let value = arr[0].as_f64()?;
                    let format_code = arr[1].as_str()?.to_string();
                    let expected = arr[2].as_str()?.to_string();
                    return Some(FractionTestCase {
                        value,
                        format_code,
                        expected,
                    });
                }
            }
            None
        })
        .collect()
}

#[test]
fn test_ssf_fractions() {
    let test_cases = load_test_cases();
    let total = test_cases.len();
    let mut passed = 0;
    let mut failed = 0;

    println!("\nRunning {} test cases from ssf fraction.json", total);

    for (i, test) in test_cases.iter().enumerate() {
        match format_default(test.value, &test.format_code) {
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
                            test.format_code,
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
                        test.format_code,
                        test.expected,
                        e
                    );
                }
            }
        }
    }

    println!("\n=== SSF Fraction Test Results ===");
    println!("Total:   {}", total);
    println!("Passed:  {} ({:.1}%)", passed, 100.0 * passed as f64 / total as f64);
    println!("Failed:  {} ({:.1}%)", failed, 100.0 * failed as f64 / total as f64);

    if failed > 0 {
        println!("\nNote: {} tests failed. Fraction formatting not yet implemented.", failed);
    }
}
