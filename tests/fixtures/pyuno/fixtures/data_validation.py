"""
Data validation fixture - tests cell input validation.

Tests:
- List validation (dropdown)
- Whole number validation
- Decimal validation
- Date validation
- Time validation
- Text length validation
- Custom formula validation
- Error/input messages
- Error styles (stop, warning, info)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from framework import fixture, FixtureBuilder


@fixture("data_validation.xlsx")
def data_validation(b: FixtureBuilder):
    """Test data validation options."""
    b.set_sheet_name("DataValidation")

    b.set_column_width("A", 30)
    b.set_column_width("B", 15)
    b.set_column_width("C", 40)

    # === List validation ===
    b.add_data("A1", "LIST VALIDATION")
    b.add_data("B1", "Input")
    b.add_data("C1", "Description")

    b.add_data("A2", "Dropdown list")
    b.add_data("C2", "Choose: Red, Green, Blue")
    b.add_data_validation(
        "B2",
        validation_type="list",
        formula1='"Red,Green,Blue"',
        show_dropdown=True,
    )

    b.add_data("A3", "List from range")
    # Create the list values
    b.add_data("E1", [["Option A"], ["Option B"], ["Option C"], ["Option D"]])
    b.add_data("C3", "Choose from E1:E4")
    b.add_data_validation(
        "B3",
        validation_type="list",
        formula1="$E$1:$E$4",
        show_dropdown=True,
    )

    # === Whole number validation ===
    b.add_data("A6", "WHOLE NUMBER VALIDATION")

    b.add_data("A7", "Between 1 and 100")
    b.add_data("C7", "Must be whole number 1-100")
    b.add_data_validation(
        "B7",
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="100",
    )

    b.add_data("A8", "Greater than 0")
    b.add_data("C8", "Positive integers only")
    b.add_data_validation(
        "B8",
        validation_type="whole",
        operator="greater_than",
        formula1="0",
    )

    b.add_data("A9", "Equal to 42")
    b.add_data("C9", "Must equal 42")
    b.add_data_validation(
        "B9",
        validation_type="whole",
        operator="equal",
        formula1="42",
    )

    # === Decimal validation ===
    b.add_data("A12", "DECIMAL VALIDATION")

    b.add_data("A13", "Between 0.0 and 1.0")
    b.add_data("C13", "Decimal 0-1 (e.g., percentage)")
    b.add_data_validation(
        "B13",
        validation_type="decimal",
        operator="between",
        formula1="0",
        formula2="1",
    )

    b.add_data("A14", "Less than 100.5")
    b.add_data("C14", "Any decimal < 100.5")
    b.add_data_validation(
        "B14",
        validation_type="decimal",
        operator="less_than",
        formula1="100.5",
    )

    # === Date validation ===
    b.add_data("A17", "DATE VALIDATION")

    b.add_data("A18", "Date after 2020-01-01")
    b.add_data("C18", "Dates in 2020 or later")
    b.add_data_validation(
        "B18",
        validation_type="date",
        operator="greater_or_equal",
        formula1="43831",  # Excel serial for 2020-01-01
    )

    b.add_data("A19", "Date range")
    b.add_data("C19", "2024-01-01 to 2024-12-31")
    b.add_data_validation(
        "B19",
        validation_type="date",
        operator="between",
        formula1="45292",  # 2024-01-01
        formula2="45657",  # 2024-12-31
    )

    # === Time validation ===
    b.add_data("A22", "TIME VALIDATION")

    b.add_data("A23", "Time between 9:00-17:00")
    b.add_data("C23", "Business hours")
    b.add_data_validation(
        "B23",
        validation_type="time",
        operator="between",
        formula1="0.375",  # 9:00 AM
        formula2="0.708333",  # 5:00 PM
    )

    # === Text length validation ===
    b.add_data("A26", "TEXT LENGTH VALIDATION")

    b.add_data("A27", "Max 10 characters")
    b.add_data("C27", "Short text only")
    b.add_data_validation(
        "B27",
        validation_type="text_length",
        operator="less_or_equal",
        formula1="10",
    )

    b.add_data("A28", "Exactly 5 characters")
    b.add_data("C28", "Must be 5 chars (e.g., ZIP)")
    b.add_data_validation(
        "B28",
        validation_type="text_length",
        operator="equal",
        formula1="5",
    )

    # === Custom formula validation ===
    b.add_data("A31", "CUSTOM FORMULA VALIDATION")

    b.add_data("A32", "Value must be even")
    b.add_data("C32", "Custom: MOD(B32,2)=0")
    b.add_data_validation(
        "B32",
        validation_type="custom",
        formula1="MOD(B32,2)=0",
    )


@fixture("data_validation_messages.xlsx")
def data_validation_messages(b: FixtureBuilder):
    """Test data validation with input/error messages."""
    b.set_sheet_name("Messages")

    b.set_column_width("A", 30)
    b.set_column_width("B", 20)

    # === Input messages ===
    b.add_data("A1", "INPUT MESSAGES")

    b.add_data("A2", "With input title & message")
    b.add_data_validation(
        "B2",
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="100",
        input_title="Enter a number",
        input_message="Please enter a whole number between 1 and 100.",
    )

    b.add_data("A3", "Input message only")
    b.add_data_validation(
        "B3",
        validation_type="list",
        formula1='"Yes,No,Maybe"',
        input_message="Select your choice from the dropdown.",
    )

    # === Error messages ===
    b.add_data("A6", "ERROR MESSAGES")

    b.add_data("A7", "Stop error style")
    b.add_data_validation(
        "B7",
        validation_type="whole",
        operator="greater_than",
        formula1="0",
        error_style="stop",
        error_title="Invalid Input",
        error_message="Value must be a positive integer. Please try again.",
    )

    b.add_data("A8", "Warning error style")
    b.add_data_validation(
        "B8",
        validation_type="decimal",
        operator="between",
        formula1="0",
        formula2="100",
        error_style="warning",
        error_title="Out of Range",
        error_message="Value is outside the recommended range (0-100). Continue anyway?",
    )

    b.add_data("A9", "Info error style")
    b.add_data_validation(
        "B9",
        validation_type="text_length",
        operator="less_or_equal",
        formula1="50",
        error_style="info",
        error_title="Long Text",
        error_message="Text exceeds recommended length. It will be accepted.",
    )

    # === Full validation with all messages ===
    b.add_data("A12", "COMPLETE VALIDATION")

    b.add_data("A13", "All features")
    b.add_data_validation(
        "B13",
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="10",
        allow_blank=True,
        input_title="Score Entry",
        input_message="Enter a score from 1 to 10.\nLeave blank if not applicable.",
        error_style="stop",
        error_title="Invalid Score",
        error_message="Score must be between 1 and 10.",
    )


@fixture("data_validation_edge_cases.xlsx")
def data_validation_edge_cases(b: FixtureBuilder):
    """Edge cases for data validation."""
    b.set_sheet_name("EdgeCases")

    b.set_column_width("A", 35)
    b.set_column_width("B", 15)

    # === Allow blank vs not allow blank ===
    b.add_data("A1", "ALLOW BLANK SETTING")

    b.add_data("A2", "Allow blank = True")
    b.add_data_validation(
        "B2",
        validation_type="whole",
        operator="greater_than",
        formula1="0",
        allow_blank=True,
    )

    b.add_data("A3", "Allow blank = False")
    b.add_data_validation(
        "B3",
        validation_type="whole",
        operator="greater_than",
        formula1="0",
        allow_blank=False,
    )

    # === Range validation ===
    b.add_data("A6", "RANGE VALIDATION")

    b.add_data("A7", "Multiple cells")
    b.add_data_validation(
        "B7:B10",
        validation_type="list",
        formula1='"Option1,Option2,Option3"',
    )

    # === Not between ===
    b.add_data("A12", "NOT BETWEEN")

    b.add_data("A13", "Not between 5 and 10")
    b.add_data_validation(
        "B13",
        validation_type="whole",
        operator="not_between",
        formula1="5",
        formula2="10",
    )

    # === Not equal ===
    b.add_data("A15", "NOT EQUAL")

    b.add_data("A16", "Not equal to 0")
    b.add_data_validation(
        "B16",
        validation_type="whole",
        operator="not_equal",
        formula1="0",
    )

    # === Show dropdown setting ===
    b.add_data("A19", "DROPDOWN VISIBILITY")

    b.add_data("A20", "Show dropdown = True")
    b.add_data_validation(
        "B20",
        validation_type="list",
        formula1='"A,B,C"',
        show_dropdown=True,
    )

    b.add_data("A21", "Show dropdown = False")
    b.add_data_validation(
        "B21",
        validation_type="list",
        formula1='"A,B,C"',
        show_dropdown=False,
    )
