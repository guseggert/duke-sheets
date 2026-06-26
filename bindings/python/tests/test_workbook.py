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

    def test_pivot_table_definitions(self):
        """Should expose semantic pivot table definitions."""
        import duke_sheets

        wb = duke_sheets.Workbook()
        sheet = wb.get_sheet(0)
        sheet.add_pivot_table(
            {
                "name": "SalesPivot",
                "source_range": "A1:C4",
                "target": "E1",
                "row_fields": [
                    {
                        "field": "Region",
                        "sort": "descending",
                        "subtotal": "none",
                        "show_drop_downs": False,
                        "subtotal_top": False,
                        "insert_blank_row": True,
                        "insert_page_break": True,
                        "include_new_items_in_filter": True,
                        "item_page_count": 25,
                    }
                ],
                "columns": ["Quarter"],
                "measures": [
                    {
                        "field": "Revenue",
                        "aggregate": "sum",
                        "name": "Revenue",
                        "show_as": "percentOfGrandTotal",
                        "number_format": "0.0%",
                    }
                ],
                "filters": [
                    {
                        "kind": "label",
                        "field": "Region",
                        "operator": "beginsWith",
                        "text": "E",
                    }
                ],
                "calculated_fields": [{"name": "Margin", "formula": "=Revenue*0.2"}],
                "calculated_items": [
                    {"field": "Region", "item": "Combined", "formula": "East+West"}
                ],
                "refresh_policy": {"refresh_on_open": True, "missing_items_limit": 25},
                "layout": {
                    "kind": "tabular",
                    "repeat_item_labels": True,
                    "page_wrap": 2,
                    "page_over_then_down": True,
                    "merge_item_labels": True,
                    "data_caption": "Metrics",
                    "grand_total_caption": "Overall",
                    "error_caption": "ERR",
                    "show_error": True,
                    "missing_caption": "N/A",
                    "show_missing": False,
                    "asterisk_totals": True,
                    "show_items": False,
                    "edit_data": True,
                    "disable_field_list": True,
                    "show_calculated_members": False,
                    "visual_totals": False,
                    "show_multiple_label": False,
                    "show_data_drop_down": False,
                    "show_member_property_tips": False,
                    "show_data_tips": False,
                    "enable_wizard": False,
                    "enable_drill": False,
                    "enable_field_properties": False,
                    "subtotal_hidden_items": True,
                    "show_drop_zones": False,
                    "indent": 3,
                    "show_empty_rows": True,
                    "show_empty_columns": True,
                },
                "overwrite_policy": "fail_on_occupied",
            }
        )

        pivot = sheet.get_pivot_table("SalesPivot")
        assert pivot is not None
        assert sheet.pivot_tables == [pivot]
        assert pivot["source"]["kind"] == "worksheet_range"
        assert pivot["source"]["range"] == "A1:C4"
        assert pivot["target"] == "E1"
        assert pivot["rows"][0]["field"] == "Region"
        assert pivot["rows"][0]["sort"] == "descending"
        assert pivot["rows"][0]["subtotal"] == "none"
        assert pivot["rows"][0]["show_drop_downs"] is False
        assert pivot["rows"][0]["subtotal_top"] is False
        assert pivot["rows"][0]["insert_blank_row"] is True
        assert pivot["rows"][0]["insert_page_break"] is True
        assert pivot["rows"][0]["include_new_items_in_filter"] is True
        assert pivot["rows"][0]["item_page_count"] == 25
        assert pivot["columns"][0]["field"] == "Quarter"
        assert pivot["measures"][0]["field"] == "Revenue"
        assert pivot["measures"][0]["aggregate"] == "sum"
        assert pivot["measures"][0]["caption"] == "Revenue"
        assert pivot["measures"][0]["number_format"] == "0.0%"
        assert pivot["measures"][0]["show_as"]["kind"] == "percent_of_grand_total"
        assert pivot["filters"][0]["kind"] == "label"
        assert pivot["filters"][0]["operator"] == "begins_with"
        assert pivot["calculated_fields"][0] == {
            "name": "Margin",
            "formula": "=Revenue*0.2",
        }
        assert pivot["calculated_items"][0] == {
            "field": "Region",
            "item": {"kind": "string", "text": "Combined"},
            "formula": "East+West",
        }
        assert pivot["layout"]["kind"] == "tabular"
        assert pivot["layout"]["repeat_item_labels"] is True
        assert pivot["layout"]["page_wrap"] == 2
        assert pivot["layout"]["page_over_then_down"] is True
        assert pivot["layout"]["merge_item_labels"] is True
        assert pivot["layout"]["data_caption"] == "Metrics"
        assert pivot["layout"]["grand_total_caption"] == "Overall"
        assert pivot["layout"]["error_caption"] == "ERR"
        assert pivot["layout"]["show_error"] is True
        assert pivot["layout"]["missing_caption"] == "N/A"
        assert pivot["layout"]["show_missing"] is False
        assert pivot["layout"]["asterisk_totals"] is True
        assert pivot["layout"]["show_items"] is False
        assert pivot["layout"]["edit_data"] is True
        assert pivot["layout"]["disable_field_list"] is True
        assert pivot["layout"]["show_calculated_members"] is False
        assert pivot["layout"]["visual_totals"] is False
        assert pivot["layout"]["show_multiple_label"] is False
        assert pivot["layout"]["show_data_drop_down"] is False
        assert pivot["layout"]["show_member_property_tips"] is False
        assert pivot["layout"]["show_data_tips"] is False
        assert pivot["layout"]["enable_wizard"] is False
        assert pivot["layout"]["enable_drill"] is False
        assert pivot["layout"]["enable_field_properties"] is False
        assert pivot["layout"]["subtotal_hidden_items"] is True
        assert pivot["layout"]["show_drop_zones"] is False
        assert pivot["layout"]["indent"] == 3
        assert pivot["layout"]["show_empty_rows"] is True
        assert pivot["layout"]["show_empty_columns"] is True
        assert pivot["refresh_policy"]["refresh_on_open"] is True
        assert pivot["refresh_policy"]["missing_items_limit"] == 25
        assert pivot["overwrite_policy"] == "fail_on_occupied"
        assert pivot["refresh_status"]["kind"] == "not_refreshed"
        assert sheet.get_pivot_table("Missing") is None

    def test_external_pivot_database_connection_roundtrip(self, temp_dir):
        """Should save and read external pivot database connection metadata."""
        import os

        import duke_sheets

        path = os.path.join(temp_dir, "external_pivot.xlsx")
        wb = duke_sheets.Workbook()
        wb.add_data_connection(
            {
                "id": 7,
                "name": "SalesConnection",
                "connection": "Provider=MSDASQL;DSN=Sales;",
                "command": "select Region, Revenue from Sales",
                "refresh_on_load": True,
            }
        )
        assert wb.data_connection_count == 1
        assert wb.data_connection_names == ["SalesConnection"]

        sheet = wb.get_sheet(0)
        sheet.add_pivot_table(
            {
                "name": "ExternalSales",
                "external_connection_name": "SalesConnection",
                "external_command_text": "select Region, Revenue from Sales",
                "target": "A1",
                "rows": ["Region"],
                "measures": [
                    {"field": "Revenue", "aggregate": "sum", "name": "Revenue"}
                ],
            }
        )
        assert sheet.pivot_count == 1

        wb.save(path)
        roundtrip = duke_sheets.Workbook.open(path)
        assert roundtrip.data_connection_count == 1
        assert roundtrip.data_connection_names == ["SalesConnection"]
        assert roundtrip.get_sheet(0).pivot_table_names == ["ExternalSales"]

    def test_non_database_data_connection_roundtrip(self, temp_dir):
        """Should save and read web, text, and OLAP connection metadata."""
        import os

        import duke_sheets

        path = os.path.join(temp_dir, "data_connections.xlsx")
        wb = duke_sheets.Workbook()
        wb.add_data_connection(
            {
                "id": 8,
                "name": "WebSales",
                "kind": "web",
                "url": "https://example.test/sales.html",
                "source_data": True,
                "html_tables": True,
            }
        )
        wb.add_data_connection(
            {
                "id": 9,
                "name": "CsvSales",
                "kind": "text",
                "source_file": "/data/sales.csv",
                "delimiter": "|",
                "first_row": 2,
            }
        )
        wb.add_data_connection(
            {
                "id": 10,
                "name": "CubeSales",
                "kind": "olap",
                "local": True,
                "local_connection": "CubeFile=cube.cub",
                "send_locale": True,
            }
        )
        assert wb.data_connection_names == ["WebSales", "CsvSales", "CubeSales"]
        assert [connection["kind"] for connection in wb.data_connections] == [
            "web",
            "text",
            "olap",
        ]
        assert wb.get_data_connection("CsvSales") == {
            "id": 9,
            "name": "CsvSales",
            "kind": "text",
            "refreshed_version": 7,
            "refresh_on_load": False,
            "background": False,
            "save_data": False,
            "source_file": "/data/sales.csv",
            "delimiter": "|",
            "first_row": 2,
            "delimited": True,
            "decimal": None,
            "thousands": None,
        }
        assert wb.get_data_connection_by_id(10)["local_connection"] == "CubeFile=cube.cub"
        assert wb.get_data_connection("Missing") is None

        wb.save(path)
        roundtrip = duke_sheets.Workbook.open(path)
        assert roundtrip.data_connection_count == 3
        assert roundtrip.data_connection_names == ["WebSales", "CsvSales", "CubeSales"]
        assert [connection["kind"] for connection in roundtrip.data_connections] == [
            "web",
            "text",
            "olap",
        ]
        assert roundtrip.get_data_connection("WebSales")["url"] == (
            "https://example.test/sales.html"
        )
        assert roundtrip.get_data_connection_by_id(10)["send_locale"] is True

    def test_olap_pivot_roundtrip(self, temp_dir):
        """Should save and read OLAP pivot source metadata."""
        import os

        import duke_sheets

        path = os.path.join(temp_dir, "olap_pivot.xlsx")
        wb = duke_sheets.Workbook()
        wb.add_data_connection(
            {
                "id": 10,
                "name": "CubeSales",
                "kind": "olap",
                "local": True,
                "local_connection": "CubeFile=cube.cub",
            }
        )

        sheet = wb.get_sheet(0)
        sheet.add_pivot_table(
            {
                "name": "OlapSales",
                "olap_connection_name": "CubeSales",
                "target": "A1",
                "rows": ["Region"],
                "measures": [
                    {"field": "Revenue", "aggregate": "sum", "name": "Revenue"}
                ],
            }
        )
        assert sheet.pivot_count == 1

        wb.save(path)
        roundtrip = duke_sheets.Workbook.open(path)
        assert roundtrip.data_connection_names == ["CubeSales"]
        assert roundtrip.get_sheet(0).pivot_table_names == ["OlapSales"]

    def test_consolidation_pivot_roundtrip(self, temp_dir):
        """Should save and read consolidation pivot source metadata."""
        import os

        import duke_sheets

        path = os.path.join(temp_dir, "consolidation_pivot.xlsx")
        wb = duke_sheets.Workbook()
        wb.add_sheet("North")
        wb.add_sheet("South")

        sheet = wb.get_sheet(0)
        sheet.add_pivot_table(
            {
                "name": "ConsolidatedSales",
                "consolidation_ranges": [
                    {
                        "sheet": "North",
                        "range": "A1:B4",
                        "name": "NorthPlan",
                        "page_items": ["FY2025", "Plan"],
                    },
                    {
                        "sheet": "South",
                        "range": "A1:B4",
                        "name": "SouthActual",
                        "page_items": ["FY2025", "Actual"],
                    },
                ],
                "target": "A1",
                "rows": ["Region"],
                "measures": [
                    {"field": "Revenue", "aggregate": "sum", "name": "Revenue"}
                ],
            }
        )
        assert sheet.pivot_count == 1

        wb.save(path)
        roundtrip = duke_sheets.Workbook.open(path)
        assert roundtrip.get_sheet(0).pivot_table_names == ["ConsolidatedSales"]

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
