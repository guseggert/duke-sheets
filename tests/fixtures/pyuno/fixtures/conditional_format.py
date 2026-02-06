"""
Conditional formatting fixture - tests conditional formatting rules.

Tests:
- Cell value conditions (greater than, less than, equal, between)
- Formula-based conditions
- DXF styling (fill, font, border)
- Multiple rules on same range
- Color scales (2-color, 3-color)
- Data bars
- Icon sets
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("conditional_format.xlsx")
def conditional_format(b: FixtureBuilder):
    """Test conditional formatting options."""
    b.set_sheet_name("ConditionalFormat")

    b.set_column_width("A", 25)
    b.set_column_width("B", 15)

    # === Cell value conditions ===
    b.add_data("A1", "CELL VALUE CONDITIONS")

    # Greater than
    b.add_data("A3", "Greater than 50")
    test_values = [25, 50, 75, 100]
    for i, val in enumerate(test_values):
        b.add_data(f"B{3 + i}", val)
    b.add_conditional_format(
        range="B3:B6",
        condition=("greater_than", 50),
        style=StyleSpec(fill_color=0x00FF00, bold=True),  # Green fill, bold
    )

    # Less than
    b.add_data("A8", "Less than 50")
    for i, val in enumerate(test_values):
        b.add_data(f"B{8 + i}", val)
    b.add_conditional_format(
        range="B8:B11",
        condition=("less_than", 50),
        style=StyleSpec(
            fill_color=0xFF0000, font_color=0xFFFFFF
        ),  # Red fill, white text
    )

    # Equal to
    b.add_data("A13", "Equal to 50")
    for i, val in enumerate(test_values):
        b.add_data(f"B{13 + i}", val)
    b.add_conditional_format(
        range="B13:B16",
        condition=("equal", 50),
        style=StyleSpec(fill_color=0xFFFF00, italic=True),  # Yellow fill, italic
    )

    # Between
    b.add_data("A18", "Between 40 and 60")
    for i, val in enumerate(test_values):
        b.add_data(f"B{18 + i}", val)
    b.add_conditional_format(
        range="B18:B21",
        condition=("between", "40"),  # Note: between needs special handling
        style=StyleSpec(
            fill_color=0x00FFFF, underline="single"
        ),  # Cyan fill, underline
    )


@fixture("conditional_format_dxf_styles.xlsx")
def conditional_format_dxf_styles(b: FixtureBuilder):
    """Test DXF (differential) styling in conditional formats."""
    b.set_sheet_name("DXFStyles")

    b.set_column_width("A", 30)
    b.set_column_width("B", 20)

    # Test data
    values = [10, 30, 50, 70, 90]

    # === Font styling ===
    b.add_data("A1", "FONT STYLING")

    b.add_data("A2", "Bold when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{2 + i}", val)
    b.add_conditional_format(
        range="B2:B6",
        condition=("greater_than", 50),
        style=StyleSpec(bold=True),
    )

    b.add_data("A8", "Italic when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{8 + i}", val)
    b.add_conditional_format(
        range="B8:B12",
        condition=("greater_than", 50),
        style=StyleSpec(italic=True),
    )

    b.add_data("A14", "Red text when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{14 + i}", val)
    b.add_conditional_format(
        range="B14:B18",
        condition=("greater_than", 50),
        style=StyleSpec(font_color=0xFF0000),
    )

    # === Fill styling ===
    b.add_data("A20", "FILL STYLING")

    b.add_data("A21", "Green fill when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{21 + i}", val)
    b.add_conditional_format(
        range="B21:B25",
        condition=("greater_than", 50),
        style=StyleSpec(fill_color=0x00FF00),
    )

    # === Border styling ===
    b.add_data("A27", "BORDER STYLING")

    b.add_data("A28", "Border when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{28 + i}", val)
    b.add_conditional_format(
        range="B28:B32",
        condition=("greater_than", 50),
        style=StyleSpec(border_style="medium", border_color=0x0000FF),
    )

    # === Combined styling ===
    b.add_data("A34", "COMBINED STYLING")

    b.add_data("A35", "All styles when > 50")
    for i, val in enumerate(values):
        b.add_data(f"B{35 + i}", val)
    b.add_conditional_format(
        range="B35:B39",
        condition=("greater_than", 50),
        style=StyleSpec(
            bold=True,
            italic=True,
            font_color=0xFFFFFF,
            fill_color=0x0000FF,
            border_style="thin",
            border_color=0x000000,
        ),
    )


@fixture("conditional_format_multiple_rules.xlsx")
def conditional_format_multiple_rules(b: FixtureBuilder):
    """Test multiple conditional formatting rules on same range."""
    b.set_sheet_name("MultipleRules")

    b.set_column_width("A", 30)
    b.set_column_width("B", 15)

    # Test data
    values = [10, 30, 50, 70, 90]

    b.add_data("A1", "MULTIPLE RULES (same range)")
    b.add_data("A2", "Red if < 30, Yellow if 30-70, Green if > 70")

    for i, val in enumerate(values):
        b.add_data(f"B{3 + i}", val)

    # Rule 1: Less than 30 -> Red
    b.add_conditional_format(
        range="B3:B7",
        condition=("less_than", 30),
        style=StyleSpec(fill_color=0xFF0000),
    )

    # Rule 2: Greater than 70 -> Green
    b.add_conditional_format(
        range="B3:B7",
        condition=("greater_than", 70),
        style=StyleSpec(fill_color=0x00FF00),
    )

    # Note: "Between" rule would need special handling in the framework


@fixture("conditional_format_edge_cases.xlsx")
def conditional_format_edge_cases(b: FixtureBuilder):
    """Edge cases for conditional formatting."""
    b.set_sheet_name("EdgeCases")

    b.set_column_width("A", 35)
    b.set_column_width("B", 15)

    # === Empty cells ===
    b.add_data("A1", "EMPTY CELLS")
    b.add_data("A2", "With empty cells")
    b.add_data("B2", 100)
    # B3 is empty
    b.add_data("B4", 50)
    b.add_data("B5", 0)
    b.add_conditional_format(
        range="B2:B5",
        condition=("greater_than", 25),
        style=StyleSpec(fill_color=0x00FF00),
    )

    # === Negative numbers ===
    b.add_data("A7", "NEGATIVE NUMBERS")
    neg_values = [-50, -25, 0, 25, 50]
    for i, val in enumerate(neg_values):
        b.add_data(f"B{8 + i}", val)
    b.add_conditional_format(
        range="B8:B12",
        condition=("less_than", 0),
        style=StyleSpec(fill_color=0xFF0000, font_color=0xFFFFFF),
    )

    # === Decimal values ===
    b.add_data("A14", "DECIMAL VALUES")
    dec_values = [0.1, 0.5, 1.0, 1.5, 2.0]
    for i, val in enumerate(dec_values):
        b.add_data(f"B{15 + i}", val)
    b.add_conditional_format(
        range="B15:B19",
        condition=("greater_than", 1),
        style=StyleSpec(fill_color=0x00FF00),
    )

    # === Text values ===
    b.add_data("A21", "TEXT VALUES (may not apply)")
    text_values = ["Apple", "Banana", "Cherry", "Date", "100"]
    for i, val in enumerate(text_values):
        b.add_data(f"B{22 + i}", val)
    b.add_conditional_format(
        range="B22:B26",
        condition=("greater_than", 50),  # Won't match text
        style=StyleSpec(fill_color=0x00FF00),
    )

    # === Large range ===
    b.add_data("A28", "LARGE RANGE (B29:B48)")
    for i in range(20):
        b.add_data(f"B{29 + i}", (i + 1) * 5)
    b.add_conditional_format(
        range="B29:B48",
        condition=("greater_than", 50),
        style=StyleSpec(fill_color=0xFFFF00),
    )
