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


class TestIterateRows:
    """Test sparse row iteration."""

    def test_iterate_rows_basic(self, workbook):
        """Should iterate sparse rows with for loop."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10)
        sheet.set_cell("C1", "hello")
        sheet.set_cell("A3", 42)
        sheet.set_cell("B5", True)

        rows = list(sheet.iterate_rows())

        assert len(rows) == 3  # rows 0, 2, 4
        assert rows[0].index == 0
        assert len(rows[0].cells) == 2
        assert rows[0].cells[0].col == 0
        assert rows[0].cells[0].value == "10"
        assert rows[0].cells[1].col == 2
        assert rows[0].cells[1].value == "hello"
        assert rows[1].index == 2
        assert rows[1].cells[0].value == "42"
        assert rows[2].index == 4
        assert rows[2].cells[0].col == 1
        assert rows[2].cells[0].value == "TRUE"

    def test_iterate_rows_empty_sheet(self, workbook):
        """Empty sheet should yield no rows."""
        sheet = workbook.get_sheet(0)
        rows = list(sheet.iterate_rows())
        assert len(rows) == 0

    def test_iterate_rows_calculated(self, workbook):
        """Should return calculated values when requested."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10)
        sheet.set_cell("A2", 20)
        sheet.set_formula("A3", "=A1+A2")
        workbook.calculate()

        rows = list(sheet.iterate_rows(use_calculated_values=True))
        a3_row = next(r for r in rows if r.index == 2)
        assert a3_row.cells[0].value == "30"

    def test_get_rows_batch(self, workbook):
        """get_rows_batch should return batched results."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "first")
        sheet.set_cell("A100", "last")

        batch1 = sheet.get_rows_batch(0, 50)
        assert len(batch1) == 1  # only row 0 in range 0..49
        assert batch1[0].index == 0

        batch2 = sheet.get_rows_batch(50, 100)
        assert len(batch2) == 1  # only row 99 in range 50..149
        assert batch2[0].index == 99

        batch3 = sheet.get_rows_batch(100, 100)
        assert len(batch3) == 0

    def test_iterate_rows_skip_empty_values(self, workbook):
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "merged")
        sheet.merge_cells("A1:C1")
        sheet.set_cell("A2", "")

        rows = list(sheet.iterate_rows(include_merge_info=True, skip_empty_values=True))

        assert [row.index for row in rows] == [0, 1]
        assert [cell.col for cell in rows[0].cells] == [0]
        assert rows[0].cells[0].value == "merged"
        assert [cell.col for cell in rows[1].cells] == [0]
        assert rows[1].cells[0].value == ""

    def test_iterate_rows_skip_blank_values(self, workbook):
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10)
        sheet.set_cell("B1", "")
        sheet.set_cell("A2", "merged")
        sheet.merge_cells("A2:C2")

        rows = list(sheet.iterate_rows(include_merge_info=True, skip_blank_values=True))

        assert [row.index for row in rows] == [0, 1]
        assert [cell.col for cell in rows[0].cells] == [0]
        assert rows[0].cells[0].value == "10"
        assert [cell.col for cell in rows[1].cells] == [0]
        assert rows[1].cells[0].value == "merged"

    def test_get_rows_batch_skip_flags(self, workbook):
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", "")
        sheet.set_cell("A2", "merged")
        sheet.merge_cells("A2:C2")
        sheet.set_cell("B3", 42)

        keep_empty_strings = sheet.get_rows_batch(
            0,
            10,
            include_merge_info=True,
            skip_empty_values=True,
        )
        assert [row.index for row in keep_empty_strings] == [0, 1, 2]
        assert keep_empty_strings[0].cells[0].value == ""
        assert [cell.col for cell in keep_empty_strings[1].cells] == [0]

        skip_blanks = sheet.get_rows_batch(
            0,
            10,
            include_merge_info=True,
            skip_blank_values=True,
        )
        assert [row.index for row in skip_blanks] == [1, 2]
        assert [cell.col for cell in skip_blanks[0].cells] == [0]
        assert skip_blanks[0].cells[0].value == "merged"
        assert skip_blanks[1].cells[0].value == "42"


class TestImages:
    """Test embedded image access."""

    def test_images_empty_by_default(self, workbook):
        """Fresh workbook should have no embedded images."""
        sheet = workbook.get_sheet(0)
        images = sheet.images
        assert images == []
