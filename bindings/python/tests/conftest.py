"""
Pytest configuration and shared fixtures for duke_sheets tests.
"""

import pytest
import tempfile
import os


@pytest.fixture
def workbook():
    """Create a fresh workbook for each test."""
    import duke_sheets
    return duke_sheets.Workbook()


@pytest.fixture
def temp_dir():
    """Create a temporary directory for file tests."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def sample_workbook():
    """Create a workbook with some sample data."""
    import duke_sheets
    
    wb = duke_sheets.Workbook()
    sheet = wb.get_sheet(0)
    
    # Set up some sample data
    sheet.set_cell("A1", 10.0)
    sheet.set_cell("A2", 20.0)
    sheet.set_cell("A3", 30.0)
    sheet.set_cell("B1", "Hello")
    sheet.set_cell("B2", "World")
    sheet.set_cell("C1", True)
    sheet.set_cell("C2", False)
    
    return wb
