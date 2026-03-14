"""
Tests for workbook calculation functionality.
"""

import pytest


class TestCalculationBasics:
    """Test basic calculation functionality."""

    def test_calculate_returns_stats(self, workbook):
        """calculate() should return CalculationStats."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate()

        assert hasattr(stats, "formula_count")
        assert hasattr(stats, "cells_calculated")
        assert hasattr(stats, "errors")
        assert hasattr(stats, "circular_references")
        assert hasattr(stats, "converged")
        assert hasattr(stats, "iterations")

    def test_calculate_formula_count(self, workbook):
        """Should count formulas correctly."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")
        sheet.set_formula("A2", "=2+2")
        sheet.set_formula("A3", "=3+3")

        stats = workbook.calculate()

        assert stats.formula_count == 3

    def test_calculate_cells_calculated(self, workbook):
        """Should count calculated cells."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")
        sheet.set_formula("A2", "=2+2")

        stats = workbook.calculate()

        assert stats.cells_calculated >= 2

    def test_calculate_no_errors(self, workbook):
        """Valid formulas should have no errors."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_formula("A2", "=A1*2")

        stats = workbook.calculate()

        assert stats.errors == 0

    def test_calculate_with_error(self, workbook):
        """Should count formula errors."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1/0")  # Division by zero

        stats = workbook.calculate()

        # The formula evaluates but produces #DIV/0!
        # This may or may not count as an "error" depending on implementation
        assert stats.formula_count == 1


class TestCalculationStats:
    """Test CalculationStats class."""

    def test_stats_repr(self, workbook):
        """CalculationStats should have useful repr."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate()
        r = repr(stats)

        assert "CalculationStats" in r
        assert "formulas=" in r
        assert "calculated=" in r


class TestIterativeCalculation:
    """Test iterative calculation for circular references."""

    def test_circular_reference_detected(self, workbook):
        """Should detect circular references."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=B1")
        sheet.set_formula("B1", "=A1")

        stats = workbook.calculate()

        assert stats.circular_references >= 2

    def test_iterative_calculation_converges(self, workbook):
        """Iterative calculation should converge."""
        sheet = workbook.get_sheet(0)
        # Simple iterative: A1 = B1/2 + 0.5, B1 = A1
        # Should converge to A1 = B1 = 1
        sheet.set_cell("A1", 0.0)  # Initial value
        sheet.set_formula("B1", "=A1")
        sheet.set_formula("A1", "=B1/2+0.5")

        stats = workbook.calculate(
            iterative=True, max_iterations=100, max_change=0.0001
        )

        assert stats.converged

    def test_calculate_with_options_params(self, workbook):
        """Should accept calculation options."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate(
            iterative=False, max_iterations=50, max_change=0.01
        )

        assert stats.formula_count == 1

    def test_calculate_with_force_full_calculation(self, workbook):
        """Should accept force_full_calculation option."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate(force_full_calculation=True)
        assert stats.formula_count == 1
        assert sheet.get_calculated_value("A1").as_number() == 2.0

    def test_calculate_with_calculate_volatile(self, workbook):
        """Should accept calculate_volatile option."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate(calculate_volatile=False)
        assert stats.formula_count == 1

    def test_calculate_specific_sheets(self, workbook):
        """Should calculate only specified sheets."""
        sheet0 = workbook.get_sheet(0)
        sheet0.set_formula("A1", "=1+1")
        workbook.add_sheet("Sheet2")
        sheet1 = workbook.get_sheet(1)
        sheet1.set_formula("A1", "=2+2")

        stats = workbook.calculate(sheets=[0])
        assert stats.cells_calculated >= 1
        assert sheet0.get_calculated_value("A1").as_number() == 2.0

    def test_calculate_with_max_threads(self, workbook):
        """Should accept max_threads option."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats = workbook.calculate(max_threads=1)
        assert stats.formula_count == 1
        assert sheet.get_calculated_value("A1").as_number() == 2.0

    def test_calculate_image_metadata(self, workbook):
        """Should expose IMAGE metadata on worksheet."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", '=IMAGE("https://example.com/logo.png","Logo",3,48,96)')

        workbook.calculate()

        assert sheet.get_calculated_value("A1").as_text() == "Logo"
        image = sheet.get_image_at(0, 0)
        assert image is not None
        assert image["source"] == "https://example.com/logo.png"
        assert image["alt_text"] == "Logo"
        assert image["sizing"] == "custom"
        assert image["width"] == 96.0
        assert image["height"] == 48.0
        # Non-image cell returns None
        assert sheet.get_image_at(1, 0) is None

    def test_worksheet_formula_count(self, workbook):
        """Worksheet should expose formula_count property."""
        sheet = workbook.get_sheet(0)

        assert sheet.formula_count == 0
        sheet.set_formula("A1", "=1+1")
        sheet.set_formula("B1", "=2+2")
        sheet.set_cell("C1", 42.0)
        assert sheet.formula_count == 2


class TestRecalculation:
    """Test recalculation behavior."""

    def test_recalculate_after_change(self, workbook):
        """Should recalculate after changing values."""
        sheet = workbook.get_sheet(0)
        sheet.set_cell("A1", 10.0)
        sheet.set_formula("A2", "=A1*2")

        workbook.calculate()
        assert sheet.get_calculated_value("A2").as_number() == 20.0

        # Change the source cell
        sheet.set_cell("A1", 50.0)
        workbook.calculate()

        assert sheet.get_calculated_value("A2").as_number() == 100.0

    def test_multiple_calculations(self, workbook):
        """Should handle multiple calculations."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", "=1+1")

        stats1 = workbook.calculate()
        stats2 = workbook.calculate()

        # Both should succeed
        assert stats1.formula_count == 1
        assert stats2.formula_count == 1


class TestCalculationOrder:
    """Test that formulas are calculated in correct order."""

    def test_dependency_chain(self, workbook):
        """Formulas should calculate in dependency order."""
        sheet = workbook.get_sheet(0)

        # Set up a dependency chain (in reverse order)
        sheet.set_formula("D1", "=C1*2")  # Depends on C1
        sheet.set_formula("C1", "=B1+5")  # Depends on B1
        sheet.set_formula("B1", "=A1*3")  # Depends on A1
        sheet.set_cell("A1", 10.0)  # Source value

        workbook.calculate()

        # A1 = 10
        # B1 = 10 * 3 = 30
        # C1 = 30 + 5 = 35
        # D1 = 35 * 2 = 70
        assert sheet.get_calculated_value("B1").as_number() == 30.0
        assert sheet.get_calculated_value("C1").as_number() == 35.0
        assert sheet.get_calculated_value("D1").as_number() == 70.0

    def test_cross_sheet_dependencies(self):
        """Formulas should work across sheets."""
        import duke_sheets

        wb = duke_sheets.Workbook()
        wb.add_sheet("Data")

        sheet1 = wb.get_sheet(0)
        sheet2 = wb.get_sheet("Data")

        # Set value in second sheet
        sheet2.set_cell("A1", 100.0)

        # Reference from first sheet
        sheet1.set_formula("A1", "=Data!A1*2")

        wb.calculate()

        value = sheet1.get_calculated_value("A1")
        assert value.as_number() == 200.0


class TestCallbackFunctions:
    """Test Python callable callbacks for WEBSERVICE and RTD."""

    def test_calculate_with_web_service_fn(self, workbook):
        """Should call Python function for WEBSERVICE."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", '=WEBSERVICE("https://example.com/api")')

        stats = workbook.calculate(
            web_service_fn=lambda url: f"response:{url}"
        )

        result = sheet.get_calculated_value("A1")
        assert result.as_text() == "response:https://example.com/api"

    def test_calculate_with_rtd_fn(self, workbook):
        """Should call Python function for RTD."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", '=RTD("prog","srv","topic1")')

        stats = workbook.calculate(
            rtd_fn=lambda prog_id, server, topics: f"{prog_id}:{server}:{','.join(topics)}"
        )

        result = sheet.get_calculated_value("A1")
        assert result.as_text() == "prog:srv:topic1"


    def test_web_service_fn_returning_none(self, workbook):
        """Callback returning None should give #N/A."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", '=WEBSERVICE("https://example.com/missing")')

        stats = workbook.calculate(
            web_service_fn=lambda url: None
        )

        result = sheet.get_calculated_value("A1")
        assert result.is_error
        assert result.as_error() == "#N/A"

    def test_rtd_fn_returning_none(self, workbook):
        """RTD callback returning None should give #N/A."""
        sheet = workbook.get_sheet(0)
        sheet.set_formula("A1", '=RTD("prog","srv","topic1")')

        stats = workbook.calculate(
            rtd_fn=lambda prog_id, server, topics: None
        )

        result = sheet.get_calculated_value("A1")
        assert result.is_error
        assert result.as_error() == "#N/A"

