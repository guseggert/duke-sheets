"""
Data types fixture - tests all cell value types.

Tests:
- Empty cells
- Numbers (integer, decimal, negative, zero, large)
- Strings (short, long, unicode, special characters)
- Booleans (true, false)
- Formulas (simple arithmetic, functions, references)
- Errors (#DIV/0!, #NAME?, #REF!, #VALUE!, #N/A)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder


@fixture("data_types.xlsx")
def data_types(b: FixtureBuilder):
    """Test all cell data types."""
    b.set_sheet_name("DataTypes")

    # Column A: Labels
    # Column B: Values
    # Column C: Additional test values

    # === Section 1: Numbers ===
    b.add_data("A1", "NUMBERS")
    b.add_data("A2", ["Integer", "Decimal", "Negative", "Zero", "Large", "Small"])
    b.add_data("B2", [42, 3.14159, -100, 0, 1e10, 0.000001])

    # === Section 2: Strings ===
    b.add_data("A9", "STRINGS")
    b.add_data(
        "A10", ["Short", "Long", "Unicode", "Special chars", "With quotes", "Empty-ish"]
    )
    b.add_data(
        "B10",
        [
            "Hello",
            "This is a longer string that spans multiple characters for testing",
            "Unicode: \u65e5\u672c\u8a9e \u4e2d\u6587 \u0420\u0443\u0441\u0441\u043a\u0438\u0439",  # Japanese, Chinese, Russian
            "Special: <>&\"' \t\n",
            'He said "Hello"',
            " ",  # Just a space
        ],
    )

    # === Section 3: Booleans ===
    b.add_data("A17", "BOOLEANS")
    b.add_data("A18", ["True", "False"])
    b.add_data("B18", [True, False])

    # === Section 4: Formulas ===
    b.add_data("A21", "FORMULAS")
    b.add_data(
        "A22", ["Simple add", "Function SUM", "Reference", "Nested IF", "Array-ish"]
    )
    b.add_formula("B22", "=1+1")
    b.add_formula("B23", "=SUM(B2:B7)")  # Sum of numbers section
    b.add_formula("B24", "=B2")  # Reference to integer cell
    b.add_formula("B25", '=IF(B18,"Yes","No")')  # Reference to boolean
    b.add_formula("B26", "=AVERAGE(B2:B7)")

    # === Section 5: Errors ===
    b.add_data("A28", "ERRORS")
    b.add_data("A29", ["#DIV/0!", "#NAME?", "#REF!", "#VALUE!", "#N/A"])
    b.add_formula("B29", "=1/0")  # #DIV/0!
    b.add_formula("B30", "=NOTAFUNCTION()")  # #NAME?
    # Note: #REF! and others are harder to create programmatically
    # We'll use INDIRECT with invalid reference
    b.add_formula("B31", '=INDIRECT("ZZZ99999")')  # May produce #REF!
    b.add_formula("B32", '=VALUE("not a number")')  # #VALUE!
    b.add_formula("B33", "=NA()")  # #N/A

    # === Section 6: Empty cells (implicit) ===
    b.add_data("A35", "EMPTY")
    b.add_data("A36", "C36 is empty ->")
    # C36 is intentionally left empty


@fixture("data_types_edge_cases.xlsx")
def data_types_edge_cases(b: FixtureBuilder):
    """Edge cases for data types."""
    b.set_sheet_name("EdgeCases")

    # Very long string
    b.add_data("A1", "Very long string:")
    b.add_data("B1", "A" * 1000)

    # String that looks like a number
    b.add_data("A2", "String as number:")
    b.add_data("B2", "12345")  # Should be string, not number

    # String that looks like a date
    b.add_data("A3", "String as date:")
    b.add_data("B3", "2024-01-15")

    # Very large number
    b.add_data("A4", "Large number:")
    b.add_data("B4", 9999999999999999.0)

    # Very small number
    b.add_data("A5", "Tiny number:")
    b.add_data("B5", 0.0000000000000001)

    # Negative zero (should be same as zero)
    b.add_data("A6", "Negative zero:")
    b.add_data("B6", -0.0)

    # Infinity (via formula)
    b.add_data("A7", "Large calc:")
    b.add_formula("B7", "=1E+308*10")  # May overflow

    # Formula returning text
    b.add_data("A8", "Formula -> text:")
    b.add_formula("B8", '="Hello" & " World"')

    # Formula returning boolean
    b.add_data("A9", "Formula -> bool:")
    b.add_formula("B9", "=1>0")
