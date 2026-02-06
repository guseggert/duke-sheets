"""
Number formats fixture - tests number formatting.

Tests:
- Built-in formats (General, Number, Currency, Date, Time, Percentage, etc.)
- Custom format strings
- Locale-specific formats
- Accounting formats
- Scientific notation
- Text format
- Special formats (ZIP, phone, SSN)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("number_formats.xlsx")
def number_formats(b: FixtureBuilder):
    """Test number format options."""
    b.set_sheet_name("NumberFormats")

    b.set_column_width("A", 30)
    b.set_column_width("B", 20)
    b.set_column_width("C", 15)

    # === General and number formats ===
    b.add_data("A1", "GENERAL AND NUMBER FORMATS")
    b.add_data("B1", "Value")
    b.add_data("C1", "Formatted")

    test_value = 1234.5678

    number_formats = [
        ("General (default)", None, test_value),
        ("0 (no decimals)", "0", test_value),
        ("0.00 (2 decimals)", "0.00", test_value),
        ("0.0000 (4 decimals)", "0.0000", test_value),
        ("#,##0 (thousands)", "#,##0", test_value),
        ("#,##0.00", "#,##0.00", test_value),
        ("#,##0.00;[Red]-#,##0.00", "#,##0.00;[Red]-#,##0.00", -test_value),
    ]

    for i, (desc, fmt, value) in enumerate(number_formats):
        row = 2 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        if fmt:
            b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Percentage formats ===
    b.add_data("A11", "PERCENTAGE FORMATS")

    pct_value = 0.1234

    pct_formats = [
        ("0%", "0%", pct_value),
        ("0.00%", "0.00%", pct_value),
        ("0.0%", "0.0%", pct_value),
    ]

    for i, (desc, fmt, value) in enumerate(pct_formats):
        row = 12 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Currency formats ===
    b.add_data("A16", "CURRENCY FORMATS")

    currency_value = 1234.56

    currency_formats = [
        ('"$"#,##0', '"$"#,##0', currency_value),
        ('"$"#,##0.00', '"$"#,##0.00', currency_value),
        ('#,##0.00" USD"', '#,##0.00" USD"', currency_value),
        ('[Red]"$"#,##0.00', '"$"#,##0.00;[Red]-"$"#,##0.00', -currency_value),
    ]

    for i, (desc, fmt, value) in enumerate(currency_formats):
        row = 17 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Date formats ===
    b.add_data("A22", "DATE FORMATS")

    # Excel serial date for 2024-03-15
    date_value = 45366

    date_formats = [
        ("MM/DD/YYYY", "MM/DD/YYYY", date_value),
        ("DD/MM/YYYY", "DD/MM/YYYY", date_value),
        ("YYYY-MM-DD", "YYYY-MM-DD", date_value),
        ("DD-MMM-YYYY", "DD-MMM-YYYY", date_value),
        ("MMMM D, YYYY", 'MMMM D", "YYYY', date_value),
        ("Short date", "M/D/YY", date_value),
    ]

    for i, (desc, fmt, value) in enumerate(date_formats):
        row = 23 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Time formats ===
    b.add_data("A30", "TIME FORMATS")

    # Time as fraction of day (14:30:45)
    time_value = 0.604687

    time_formats = [
        ("HH:MM", "HH:MM", time_value),
        ("HH:MM:SS", "HH:MM:SS", time_value),
        ("HH:MM AM/PM", "HH:MM AM/PM", time_value),
    ]

    for i, (desc, fmt, value) in enumerate(time_formats):
        row = 31 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Scientific notation ===
    b.add_data("A35", "SCIENTIFIC NOTATION")

    sci_value = 123456789

    sci_formats = [
        ("0.00E+00", "0.00E+00", sci_value),
        ("0.000E+00", "0.000E+00", sci_value),
        ("##0.0E+0", "##0.0E+0", sci_value),
    ]

    for i, (desc, fmt, value) in enumerate(sci_formats):
        row = 36 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Fraction formats ===
    b.add_data("A40", "FRACTION FORMATS")

    frac_value = 0.625

    frac_formats = [
        ("# ?/?", "# ?/?", frac_value),
        ("# ??/??", "# ??/??", frac_value),
        ("# ?/8 (eighths)", "# ?/8", frac_value),
    ]

    for i, (desc, fmt, value) in enumerate(frac_formats):
        row = 41 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Text format ===
    b.add_data("A45", "TEXT FORMAT")
    b.add_data("A46", "@ (text)")
    b.add_data("B46", 12345)
    b.add_data("C46", 12345)
    b.set_cell_style("C46", StyleSpec(number_format="@"))


@fixture("number_formats_edge_cases.xlsx")
def number_formats_edge_cases(b: FixtureBuilder):
    """Edge cases for number formats."""
    b.set_sheet_name("EdgeCases")

    b.set_column_width("A", 35)
    b.set_column_width("B", 15)
    b.set_column_width("C", 15)

    # === Conditional number formats ===
    b.add_data("A1", "CONDITIONAL FORMATS")
    b.add_data("B1", "Value")
    b.add_data("C1", "Formatted")

    # Format: positive;negative;zero;text
    conditional_formats = [
        ("Positive green, negative red", "[Green]#,##0.00;[Red]-#,##0.00", 1234.56),
        ("Negative value", "[Green]#,##0.00;[Red]-#,##0.00", -1234.56),
        ("Zero in blue", "#,##0.00;-#,##0.00;[Blue]0.00", 0),
    ]

    for i, (desc, fmt, value) in enumerate(conditional_formats):
        row = 2 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Leading zeros ===
    b.add_data("A7", "LEADING ZEROS")

    leading_zero_formats = [
        ("00000 (ZIP code)", "00000", 1234),
        ("000-00-0000 (SSN style)", "000-00-0000", 123456789),
    ]

    for i, (desc, fmt, value) in enumerate(leading_zero_formats):
        row = 8 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Custom text in format ===
    b.add_data("A12", "CUSTOM TEXT")

    custom_text_formats = [
        ("With suffix", '#,##0" units"', 1500),
        ("With prefix", '"Total: "#,##0', 1500),
        ("Degrees", '0.0"°"', 45.5),
    ]

    for i, (desc, fmt, value) in enumerate(custom_text_formats):
        row = 13 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Very large and small numbers ===
    b.add_data("A18", "EXTREME VALUES")

    extreme_formats = [
        ("Very large", "#,##0", 9999999999999),
        ("Very small", "0.0000000000", 0.000000001),
        ("Negative large", "#,##0", -9999999999),
    ]

    for i, (desc, fmt, value) in enumerate(extreme_formats):
        row = 19 + i
        b.add_data(f"A{row}", desc)
        b.add_data(f"B{row}", value)
        b.add_data(f"C{row}", value)
        b.set_cell_style(f"C{row}", StyleSpec(number_format=fmt))

    # === Format combined with styling ===
    b.add_data("A24", "FORMAT + STYLING")

    b.add_data("A25", "Currency, bold, red bg")
    b.add_data("B25", 1234.56)
    b.add_data("C25", 1234.56)
    b.set_cell_style(
        "C25",
        StyleSpec(
            number_format='"$"#,##0.00',
            bold=True,
            fill_color=0xFFCCCC,
        ),
    )
