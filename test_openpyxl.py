"""Basic openpyxl exporter tests."""
import tempfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl import load_workbook


def test_create_workbook():
    wb = Workbook()
    assert wb is not None
    ws = wb.active
    assert ws is not None


def test_write_read_cell():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 42
    assert ws["A1"].value == 42


def test_cell_row_column():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=2, column=3, value="hello")
    assert ws.cell(row=2, column=3).value == "hello"


def test_append_row():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2, 3])
    assert ws["A1"].value == 1
    assert ws["B1"].value == 2
    assert ws["C1"].value == 3


def test_save_load_roundtrip():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "test"
    ws["B2"] = 100
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = Path(f.name)
    try:
        wb.save(path)
        wb2 = load_workbook(path)
        ws2 = wb2.active
        assert ws2["A1"].value == "test"
        assert ws2["B2"].value == 100
    finally:
        path.unlink(missing_ok=True)
