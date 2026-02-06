"""
Merged cells fixture - tests cell merging.

Tests:
- Basic horizontal merge
- Basic vertical merge
- Block merge (multiple rows and columns)
- Multiple merges on same sheet
- Merge with styling
- Merge with data
- Adjacent merges
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder, StyleSpec


@fixture("merged_cells.xlsx")
def merged_cells(b: FixtureBuilder):
    """Test merged cell options."""
    b.set_sheet_name("MergedCells")

    b.set_column_width("A", 15)
    b.set_column_width("B", 15)
    b.set_column_width("C", 15)
    b.set_column_width("D", 15)

    # === Horizontal merges ===
    b.add_data("A1", "HORIZONTAL MERGES")

    b.add_data("A3", "2 columns")
    b.add_merge("B3:C3")

    b.add_data("A4", "3 columns")
    b.add_merge("B4:D4")

    b.add_data("A5", "4 columns")
    b.add_merge("B5:E5")

    # === Vertical merges ===
    b.add_data("A8", "VERTICAL MERGES")

    b.add_data("B9", "2 rows")
    b.add_merge("B9:B10")

    b.add_data("C9", "3 rows")
    b.add_merge("C9:C11")

    b.add_data("D9", "4 rows")
    b.add_merge("D9:D12")

    # === Block merges ===
    b.add_data("A15", "BLOCK MERGES")

    b.add_data("B16", "2x2")
    b.add_merge("B16:C17")

    b.add_data("D16", "3x3")
    b.add_merge("D16:F18")

    # === Merges with data ===
    b.add_data("A21", "MERGES WITH DATA")

    b.add_data("B22", "Merged content here")
    b.add_merge("B22:D22")

    b.add_data("B24", "Multi-line\nmerged\ncontent")
    b.add_merge("B24:C26")

    # === Merges with styling ===
    b.add_data("A29", "MERGES WITH STYLING")

    b.add_data("B30", "Bold merged")
    b.set_cell_style("B30", StyleSpec(bold=True))
    b.add_merge("B30:D30")

    b.add_data("B32", "Colored merged")
    b.set_cell_style("B32", StyleSpec(fill_color=0xFFFF00))
    b.add_merge("B32:D32")

    b.add_data("B34", "Full styled merge")
    b.set_cell_style(
        "B34",
        StyleSpec(
            bold=True,
            font_color=0xFFFFFF,
            fill_color=0x0000FF,
            horizontal="center",
            vertical="center",
            border_style="medium",
            border_color=0x000000,
        ),
    )
    b.add_merge("B34:D36")


@fixture("merged_cells_multiple.xlsx")
def merged_cells_multiple(b: FixtureBuilder):
    """Test multiple merges on the same sheet."""
    b.set_sheet_name("MultipleMerges")

    # Create a table-like structure with merges
    b.add_data("A1", "Report Title")
    b.add_merge("A1:F1")
    b.set_cell_style(
        "A1",
        StyleSpec(
            bold=True,
            font_size=16,
            horizontal="center",
            fill_color=0xCCCCCC,
        ),
    )

    # Header row with merges
    b.add_data("A3", "Category")
    b.add_merge("A3:A4")

    b.add_data("B3", "Q1")
    b.add_merge("B3:C3")

    b.add_data("D3", "Q2")
    b.add_merge("D3:E3")

    b.add_data("F3", "Total")
    b.add_merge("F3:F4")

    # Sub-headers
    b.add_data("B4", "Jan")
    b.add_data("C4", "Feb")
    b.add_data("D4", "Mar")
    b.add_data("E4", "Apr")

    # Data rows
    categories = ["Sales", "Expenses", "Profit"]
    for i, cat in enumerate(categories):
        row = 5 + i
        b.add_data(f"A{row}", cat)
        b.add_data(f"B{row}", 100 + i * 10)
        b.add_data(f"C{row}", 110 + i * 10)
        b.add_data(f"D{row}", 120 + i * 10)
        b.add_data(f"E{row}", 130 + i * 10)
        b.add_formula(f"F{row}", f"=SUM(B{row}:E{row})")

    # Footer with merge
    b.add_data("A8", "Summary")
    b.add_merge("A8:F8")
    b.set_cell_style(
        "A8",
        StyleSpec(bold=True, fill_color=0xEEEEEE, horizontal="center"),
    )


@fixture("merged_cells_edge_cases.xlsx")
def merged_cells_edge_cases(b: FixtureBuilder):
    """Edge cases for merged cells."""
    b.set_sheet_name("EdgeCases")

    # === Single cell "merge" (1x1) ===
    # Note: This is technically not a merge, but test the boundary
    b.add_data("A1", "EDGE CASES")

    # === Adjacent merges (horizontal) ===
    b.add_data("A3", "ADJACENT HORIZONTAL")
    b.add_data("B3", "Merge 1")
    b.add_merge("B3:C3")
    b.add_data("D3", "Merge 2")
    b.add_merge("D3:E3")

    # === Adjacent merges (vertical) ===
    b.add_data("A5", "ADJACENT VERTICAL")
    b.add_data("B5", "M1")
    b.add_merge("B5:B6")
    b.add_data("B7", "M2")
    b.add_merge("B7:B8")

    # === Large merge ===
    b.add_data("A10", "LARGE MERGE (10x5)")
    b.add_data("B11", "Large merged region")
    b.add_merge("B11:K15")
    b.set_cell_style(
        "B11",
        StyleSpec(
            horizontal="center",
            vertical="center",
            fill_color=0xCCFFCC,
        ),
    )

    # === Merge at edges of sheet ===
    b.add_data("A17", "MERGE AT ROW 1 (see Sheet2)")
    b.add_sheet("Sheet2")
    b.add_data("A1", "Merged at top")
    b.add_merge("A1:D1")

    # === Merge with formula result ===
    b.select_sheet(0)
    b.add_data("A19", "MERGE WITH FORMULA")
    b.add_formula("B19", "=100+200")
    b.add_merge("B19:D19")

    # === Merge with comment ===
    b.add_data("A21", "MERGE WITH COMMENT")
    b.add_data("B21", "Has comment")
    b.add_comment("B21", "Comment on merged cell")
    b.add_merge("B21:D21")

    # === Merge with conditional format ===
    b.add_data("A23", "MERGE WITH COND FORMAT")
    b.add_data("B23", 100)
    b.add_merge("B23:D23")
    b.add_conditional_format(
        range="B23:D23",
        condition=("greater_than", 50),
        style=StyleSpec(fill_color=0x00FF00),
    )
