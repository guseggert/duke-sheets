use ssfmt::{format, format_default};

#[test]
fn test_format_convenience() {
    let opts = ssfmt::FormatOptions::default();
    let result = format(1234.5, "#,##0.00", &opts).unwrap();
    assert_eq!(result, "1,234.50");
}

#[test]
fn test_format_default_convenience() {
    let result = format_default(0.42, "0%").unwrap();
    assert_eq!(result, "42%");
}

#[test]
fn test_format_invalid_code() {
    let opts = ssfmt::FormatOptions::default();
    // Empty format should error
    let result = format(42.0, "", &opts);
    assert!(result.is_err());
}
