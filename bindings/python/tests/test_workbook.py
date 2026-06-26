"""
Tests for Workbook class.
"""

import pytest
import os


class TestWorkbookCreation:
    """Test workbook creation and basic properties."""

    def test_new_workbook(self):
        """New workbook should have one sheet."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        assert wb.sheet_count == 1

    def test_new_workbook_sheet_name(self):
        """Default sheet should be named 'Sheet1'."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        assert wb.sheet_names == ["Sheet1"]

    def test_workbook_repr(self):
        """Workbook should have a useful repr."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        assert "Workbook" in repr(wb)
        assert "sheets=1" in repr(wb)


class TestSheetManagement:
    """Test adding and removing sheets."""

    def test_add_sheet(self):
        """Should be able to add a new sheet."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        idx = wb.add_sheet("NewSheet")
        
        assert idx == 1
        assert wb.sheet_count == 2
        assert "NewSheet" in wb.sheet_names

    def test_add_multiple_sheets(self):
        """Should be able to add multiple sheets."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        wb.add_sheet("Sheet2")
        wb.add_sheet("Sheet3")
        
        assert wb.sheet_count == 3
        assert wb.sheet_names == ["Sheet1", "Sheet2", "Sheet3"]

    def test_remove_sheet(self):
        """Should be able to remove a sheet."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        wb.add_sheet("ToRemove")
        assert wb.sheet_count == 2
        
        wb.remove_sheet(1)
        assert wb.sheet_count == 1
        assert "ToRemove" not in wb.sheet_names

    def test_get_sheet_by_index(self):
        """Should get sheet by index."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        assert sheet.name == "Sheet1"

    def test_get_sheet_by_name(self):
        """Should get sheet by name."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        wb.add_sheet("MySheet")
        
        sheet = wb.get_sheet("MySheet")
        assert sheet.name == "MySheet"

    def test_get_sheet_invalid_index(self):
        """Should raise IndexError for invalid index."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        
        with pytest.raises(IndexError):
            wb.get_sheet(999)

    def test_get_sheet_invalid_name(self):
        """Should raise IndexError for invalid name."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        
        with pytest.raises(IndexError):
            wb.get_sheet("NonExistent")


class TestFileOperations:
    """Test saving and loading workbooks."""

    def test_save_xlsx(self, temp_dir):
        """Should save workbook as XLSX."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        sheet.set_cell("A1", 42.0)
        
        path = os.path.join(temp_dir, "test.xlsx")
        wb.save(path)
        
        assert os.path.exists(path)
        assert os.path.getsize(path) > 0

    def test_open_xlsx(self, temp_dir):
        """Should open XLSX file."""
        import duke_sheets
        
        # Create and save a workbook
        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        sheet.set_cell("A1", 123.0)
        sheet.set_cell("B1", "Hello")
        
        path = os.path.join(temp_dir, "test.xlsx")
        wb.save(path)
        
        # Open it again
        wb2 = duke_sheets.Workbook.open(path)
        sheet2 = wb2.get_sheet(0)
        
        assert sheet2.get_cell("A1").as_number() == 123.0
        assert sheet2.get_cell("B1").as_text() == "Hello"

    def test_save_csv(self, temp_dir):
        """Should save workbook as CSV."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        sheet.set_cell("A1", 1.0)
        sheet.set_cell("B1", 2.0)
        sheet.set_cell("A2", 3.0)
        sheet.set_cell("B2", 4.0)
        
        path = os.path.join(temp_dir, "test.csv")
        wb.save(path)
        
        assert os.path.exists(path)
        
        # Read the CSV content
        with open(path) as f:
            content = f.read()
        
        assert "1" in content
        assert "2" in content


class TestNamedRanges:
    """Test named range functionality."""

    def test_define_name(self):
        """Should define a named range."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        wb.define_name("TaxRate", "0.05")
        
        result = wb.get_named_range("TaxRate")
        assert result == "0.05"

    def test_define_name_cell_reference(self):
        """Should define a named range with cell reference."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        wb.define_name("Price", "Sheet1!$A$1")
        
        result = wb.get_named_range("Price")
        assert "A" in result and "1" in result

    def test_get_undefined_name(self):
        """Should return None for undefined name."""
        import duke_sheets
        
        wb = duke_sheets.Workbook()
        result = wb.get_named_range("NotDefined")
        assert result is None


class TestPivotTables:
    """Test pivot table authoring through the Python binding."""

    def test_refresh_manual_grouping_from_options(self):
        """Should refresh a manually grouped pivot from semantic options."""
        import duke_sheets

        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        sheet.set_cell("A1", "Region")
        sheet.set_cell("B1", "Revenue")
        sheet.set_cell("A2", "East")
        sheet.set_cell("B2", 10.0)
        sheet.set_cell("A3", "West")
        sheet.set_cell("B3", 20.0)
        sheet.set_cell("A4", "South")
        sheet.set_cell("B4", 5.0)

        sheet.add_pivot_table(
            {
                "name": "ManualGroupedRegions",
                "source_range": "A1:B4",
                "target": "D1",
                "rows": ["Region"],
                "measures": [
                    {"field": "Revenue", "aggregate": "sum", "name": "Revenue"}
                ],
                "groupings": [
                    {
                        "kind": "manual",
                        "field": "Region",
                        "groups": [{"name": "Coastal", "members": ["East", "West"]}],
                    }
                ],
            }
        )

        assert sheet.pivot_count == 1
        assert sheet.pivot_table_names == ["ManualGroupedRegions"]

        stats = wb.refresh_pivots()

        assert stats["pivot_count"] == 1
        assert stats["pivots_refreshed"] == 1
        assert sheet.get_cell("D2").as_text() == "Coastal"
        assert sheet.get_cell("E2").as_number() == 30.0
        assert sheet.get_cell("D3").as_text() == "South"
        assert sheet.get_cell("E3").as_number() == 5.0
        assert sheet.get_cell("E4").as_number() == 35.0
