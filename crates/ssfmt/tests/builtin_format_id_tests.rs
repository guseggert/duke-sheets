use ssfmt::{format_with_id_default, format_code_from_id};

/// Test built-in format ID 0 (General)
#[test]
fn test_format_id_0_general() {
    // Basic numbers
    assert_eq!(format_with_id_default(0.0, 0).unwrap(), "0");
    assert_eq!(format_with_id_default(1.0, 0).unwrap(), "1");
    assert_eq!(format_with_id_default(1234.56, 0).unwrap(), "1234.56");
    assert_eq!(format_with_id_default(-42.5, 0).unwrap(), "-42.5");

    // No trailing zeros
    assert_eq!(format_with_id_default(1.5, 0).unwrap(), "1.5");
    assert_eq!(format_with_id_default(1.50000, 0).unwrap(), "1.5");
}

/// Test built-in format ID 1 (0)
#[test]
fn test_format_id_1_integer() {
    assert_eq!(format_with_id_default(0.0, 1).unwrap(), "0");
    assert_eq!(format_with_id_default(1.0, 1).unwrap(), "1");
    assert_eq!(format_with_id_default(1234.5, 1).unwrap(), "1235"); // Rounds
    assert_eq!(format_with_id_default(1234.4, 1).unwrap(), "1234");
    assert_eq!(format_with_id_default(-42.0, 1).unwrap(), "-42");
}

/// Test built-in format ID 2 (0.00)
#[test]
fn test_format_id_2_two_decimals() {
    assert_eq!(format_with_id_default(0.0, 2).unwrap(), "0.00");
    assert_eq!(format_with_id_default(1.0, 2).unwrap(), "1.00");
    assert_eq!(format_with_id_default(1234.56, 2).unwrap(), "1234.56");
    assert_eq!(format_with_id_default(1234.567, 2).unwrap(), "1234.57"); // Rounds
    assert_eq!(format_with_id_default(-42.1, 2).unwrap(), "-42.10");
}

/// Test built-in format ID 3 (#,##0)
#[test]
fn test_format_id_3_thousands() {
    assert_eq!(format_with_id_default(0.0, 3).unwrap(), "0");
    assert_eq!(format_with_id_default(999.0, 3).unwrap(), "999");
    assert_eq!(format_with_id_default(1000.0, 3).unwrap(), "1,000");
    assert_eq!(format_with_id_default(1234567.0, 3).unwrap(), "1,234,567");
    assert_eq!(format_with_id_default(-1234.0, 3).unwrap(), "-1,234");
}

/// Test built-in format ID 4 (#,##0.00)
#[test]
fn test_format_id_4_thousands_decimals() {
    assert_eq!(format_with_id_default(0.0, 4).unwrap(), "0.00");
    assert_eq!(format_with_id_default(1234.56, 4).unwrap(), "1,234.56");
    assert_eq!(format_with_id_default(1234567.89, 4).unwrap(), "1,234,567.89");
    assert_eq!(format_with_id_default(-1000.5, 4).unwrap(), "-1,000.50");
}

/// Test built-in format ID 9 (0%)
#[test]
fn test_format_id_9_percent() {
    assert_eq!(format_with_id_default(0.0, 9).unwrap(), "0%");
    assert_eq!(format_with_id_default(0.5, 9).unwrap(), "50%");
    assert_eq!(format_with_id_default(1.0, 9).unwrap(), "100%");
    assert_eq!(format_with_id_default(0.125, 9).unwrap(), "13%"); // Rounds to 13%
    assert_eq!(format_with_id_default(-0.25, 9).unwrap(), "-25%");
}

/// Test built-in format ID 10 (0.00%)
#[test]
fn test_format_id_10_percent_decimals() {
    assert_eq!(format_with_id_default(0.0, 10).unwrap(), "0.00%");
    assert_eq!(format_with_id_default(0.5, 10).unwrap(), "50.00%");
    assert_eq!(format_with_id_default(0.125, 10).unwrap(), "12.50%");
    assert_eq!(format_with_id_default(1.2345, 10).unwrap(), "123.45%");
}

/// Test built-in format ID 11 (0.00E+00)
#[test]
fn test_format_id_11_scientific() {
    let code = format_code_from_id(11).unwrap();
    assert_eq!(code, "0.00E+00");

    // Basic scientific notation formatting
    // Note: Actual scientific formatting may need implementation
}

/// Test built-in format ID 12 (# ?/?)
#[test]
fn test_format_id_12_fraction_single() {
    let code = format_code_from_id(12).unwrap();
    assert_eq!(code, "# ?/?");

    // Fraction formatting
    // Note: Fraction implementation is complex
}

/// Test built-in format ID 14 (m/d/yy) - Date format
#[test]
fn test_format_id_14_date() {
    let code = format_code_from_id(14).unwrap();
    assert_eq!(code, "m/d/yy");

    // Excel serial date: 44927 = 2023-01-01
    // Date formatting depends on date implementation
}

/// Test built-in format ID 49 (@) - Text format
#[test]
fn test_format_id_49_text() {
    let code = format_code_from_id(49).unwrap();
    assert_eq!(code, "@");
}

/// Test invalid format IDs
#[test]
fn test_invalid_format_ids() {
    // ID 5-8 are not defined
    assert!(format_with_id_default(123.0, 5).is_err());
    assert!(format_with_id_default(123.0, 6).is_err());
    assert!(format_with_id_default(123.0, 7).is_err());
    assert!(format_with_id_default(123.0, 8).is_err());

    // ID 164+ are custom formats
    assert!(format_with_id_default(123.0, 164).is_err());
    assert!(format_with_id_default(123.0, 999).is_err());
}

/// Test that all defined format IDs can be looked up
#[test]
fn test_all_defined_format_ids() {
    let defined_ids = vec![
        0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
        37, 38, 39, 40, 45, 46, 47, 48, 49,
    ];

    for id in defined_ids {
        let code = format_code_from_id(id);
        assert!(code.is_some(), "Format ID {} should be defined", id);

        // Verify we can parse and use the format code
        let code_str = code.unwrap();
        assert!(!code_str.is_empty(), "Format code for ID {} should not be empty", id);
    }
}

/// Test accounting formats with parentheses for negatives
#[test]
fn test_accounting_formats() {
    // Format 37: #,##0 ;(#,##0)
    // Positive and negative with parentheses
    let code = format_code_from_id(37).unwrap();
    assert_eq!(code, "#,##0 ;(#,##0)");

    // Format 38: #,##0 ;[Red](#,##0)
    // Negative with red color
    let code = format_code_from_id(38).unwrap();
    assert_eq!(code, "#,##0 ;[Red](#,##0)");
}
