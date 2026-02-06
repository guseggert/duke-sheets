"""
Tests for Worksheet class.
"""

import pytest


class TestCellOperations:
    """Test cell get/set operations."""

    def test_set_get_number(self, workbook):
        """Should set and get a number."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        
        value = sheet.get_cell("A1")
        assert value.is_number
        assert value.as_number() == 42.0

    def test_set_get_integer(self, workbook):
        """Should convert integer to float."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42)
        
        value = sheet.get_cell("A1")
        assert value.is_number
        assert value.as_number() == 42.0

    def test_set_get_text(self, workbook):
        """Should set and get text."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Hello, World!")
        
        value = sheet.get_cell("A1")
        assert value.is_text
        assert value.as_text() == "Hello, World!"

    def test_set_get_boolean_true(self, workbook):
        """Should set and get True."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", True)
        
        value = sheet.get_cell("A1")
        assert value.is_boolean
        assert value.as_boolean() == True

    def test_set_get_boolean_false(self, workbook):
        """Should set and get False."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", False)
        
        value = sheet.get_cell("A1")
        assert value.is_boolean
        assert value.as_boolean() == False

    def test_set_none_clears_cell(self, workbook):
        """Setting None should clear the cell."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        sheet.set_cell("A1", None)
        
        value = sheet.get_cell("A1")
        assert value.is_empty

    def test_get_empty_cell(self, workbook):
        """Getting an empty cell should return Empty."""
        sheet = workbook.get_sheet(0)
        value = sheet.get_cell("Z99")
        assert value.is_empty

    def test_invalid_cell_address(self, workbook):
        """Should raise error for invalid address."""
        sheet = workbook.get_sheet(0)
        
        with pytest.raises(ValueError):
            sheet.set_cell("invalid", 42)

    def test_cell_address_case_insensitive(self, workbook):
        """Cell addresses should be case-insensitive."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("a1", 42.0)
        
        value = sheet.get_cell("A1")
        assert value.as_number() == 42.0


class TestUsedRange:
    """Test used range detection."""

    def test_empty_sheet_used_range(self, workbook):
        """Empty sheet should have no used range."""
        sheet = workbook.get_sheet(0)
        assert sheet.used_range is None

    def test_single_cell_used_range(self, workbook):
        """Single cell should define used range."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("B2", 42.0)
        
        used = sheet.used_range
        assert used is not None
        min_row, min_col, max_row, max_col = used
        assert min_row == 1  # B2 is row 1 (0-indexed)
        assert min_col == 1  # B is column 1
        assert max_row == 1
        assert max_col == 1

    def test_multiple_cells_used_range(self, workbook):
        """Multiple cells should expand used range."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 1.0)
        sheet.set_cell("C5", 2.0)
        
        used = sheet.used_range
        assert used is not None
        min_row, min_col, max_row, max_col = used
        assert min_row == 0  # A1
        assert min_col == 0
        assert max_row == 4  # C5 is row 4 (0-indexed)
        assert max_col == 2  # C is column 2


class TestRowColumnDimensions:
    """Test row height and column width."""

    def test_set_row_height(self, workbook):
        """Should set row height."""
        sheet = workbook.get_sheet(0)
        sheet.set_row_height(0, 30.0)
        
        height = sheet.get_row_height(0)
        assert height == 30.0

    def test_set_column_width(self, workbook):
        """Should set column width."""
        sheet = workbook.get_sheet(0)
        sheet.set_column_width(0, 15.0)
        
        width = sheet.get_column_width(0)
        assert width == 15.0

    def test_default_row_height(self, workbook):
        """Unset row height should return None."""
        sheet = workbook.get_sheet(0)
        height = sheet.get_row_height(0)
        assert height is None

    def test_default_column_width(self, workbook):
        """Unset column width should return None."""
        sheet = workbook.get_sheet(0)
        width = sheet.get_column_width(0)
        assert width is None


class TestMergeCells:
    """Test cell merging."""

    def test_merge_cells(self, workbook):
        """Should merge cells."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "Merged")
        sheet.merge_cells("A1:C3")
        # No error means success

    def test_unmerge_cells(self, workbook):
        """Should unmerge cells."""
        sheet = workbook.get_sheet(0)
        sheet.merge_cells("A1:C3")
        sheet.unmerge_cells("A1:C3")
        # No error means success


class TestWorksheetRepr:
    """Test worksheet string representation."""

    def test_worksheet_repr(self, workbook):
        """Worksheet should have useful repr."""
        sheet = workbook.get_sheet(0)
        r = repr(sheet)
        assert "Worksheet" in r
        assert "Sheet1" in r
