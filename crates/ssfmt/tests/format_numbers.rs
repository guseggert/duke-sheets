use ssfmt::{FormatOptions, NumberFormat};

#[test]
fn test_format_integer() {
    let fmt = NumberFormat::parse("0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42");
    assert_eq!(fmt.format(42.7, &opts), "43"); // Rounds
}

#[test]
fn test_format_decimal() {
    let fmt = NumberFormat::parse("0.00").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42.00");
    assert_eq!(fmt.format(42.567, &opts), "42.57");
}

#[test]
fn test_format_thousands() {
    let fmt = NumberFormat::parse("#,##0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(1234567.0, &opts), "1,234,567");
    assert_eq!(fmt.format(123.0, &opts), "123");
}

#[test]
fn test_format_percentage() {
    let fmt = NumberFormat::parse("0%").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.42, &opts), "42%");
    assert_eq!(fmt.format(1.5, &opts), "150%");
}

#[test]
fn test_format_hash_placeholder() {
    let fmt = NumberFormat::parse("#.##").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.5, &opts), "42.5");
    assert_eq!(fmt.format(42.0, &opts), "42.");
}

#[test]
fn test_format_negative_section() {
    let fmt = NumberFormat::parse("0;-0").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(42.0, &opts), "42");
    assert_eq!(fmt.format(-42.0, &opts), "-42");
}
