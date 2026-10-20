import sys
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

import salary_report as app


def make_row(**overrides):
    values = {
        "Mã NV": "NV001",
        "Tên NV": "Nguyễn Văn A",
        "Chức vụ": "Nhân viên",
        "Lương cơ bản": 12_000_000,
        "Lương đóng BHXH": 12_000_000,
        "Ngày công chuẩn": 26,
        "Ngày công đi làm": 26,
        "Thưởng": 1_000_000,
        "Phụ cấp": 500_000,
        "Phạt": 100_000,
    }
    values.update(overrides)
    return pd.Series(values)


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("15.000.000", Decimal("15000000")),
        ("15,000,000", Decimal("15000000")),
        ("25.5", Decimal("25.5")),
        ("1.234,56", Decimal("1234.56")),
        ("1,234.56", Decimal("1234.56")),
        (12_500_000, Decimal("12500000")),
    ],
)
def test_parse_number_supports_common_formats(raw, expected):
    assert app.parse_number(raw, "Trường dữ liệu") == expected


@pytest.mark.parametrize("raw", ["", "không phải số", "NaN", "Infinity", -1, "-25"])
def test_parse_number_rejects_invalid_or_negative_values(raw):
    with pytest.raises(ValueError):
        app.parse_number(raw, "Trường dữ liệu")


@pytest.mark.parametrize(
    ("taxable_income", "expected_tax"),
    [
        (0, 0),
        (10_000_000, 500_000),
        (20_000_000, 1_500_000),
        (30_000_000, 2_500_000),
        (60_000_000, 8_500_000),
        (100_000_000, 20_500_000),
        (120_000_000, 27_500_000),
    ],
)
def test_pit_2026_brackets(taxable_income, expected_tax):
    assert app.calculate_pit(taxable_income) == Decimal(expected_tax)


def test_salary_calculation_is_consistent():
    result = app.calculate_salary(make_row())

    assert result["Lương thực tế"] == Decimal("12000000")
    assert result["Tổng thu nhập"] == Decimal("13500000")
    assert result["BHXH"] == Decimal("960000")
    assert result["BHYT"] == Decimal("180000")
    assert result["BHTN"] == Decimal("120000")
    assert result["Thu nhập tính thuế"] == 0
    assert result["Thuế TNCN"] == 0
    assert result["Tổng khấu trừ"] == Decimal("1360000")
    assert result["Thực lãnh"] == Decimal("12140000")
    assert result["Tổng thu nhập"] - result["Tổng khấu trừ"] == result["Thực lãnh"]


def test_missing_optional_dependents_column_defaults_to_zero():
    result = app.calculate_salary(make_row())

    assert result["Số người phụ thuộc"] == 0


def test_fractional_dependent_count_is_rejected():
    with pytest.raises(ValueError, match="số nguyên"):
        app.calculate_salary(make_row(**{"Số người phụ thuộc": 1.5}))


def test_attendance_greater_than_standard_days_is_rejected():
    with pytest.raises(ValueError, match="ngày công"):
        app.calculate_salary(make_row(**{"Ngày công đi làm": 27}))


def test_duplicate_employee_ids_are_rejected():
    frame = pd.DataFrame({"Mã NV": ["NV001", "NV002", "NV001"]})

    with pytest.raises(ValueError, match="Mã NV bị trùng"):
        app.validate_employee_ids(frame)


def test_empty_input_is_rejected():
    with pytest.raises(ValueError, match="không có bản ghi"):
        app.validate_columns(pd.DataFrame())


def test_pdf_is_written_atomically_and_missing_dependents_column_is_supported(
    tmp_path, monkeypatch
):
    class FakeHTML:
        def __init__(self, string, base_url):
            self.string = string
            self.base_url = base_url

        def write_pdf(self, path):
            Path(path).write_bytes(b"%PDF-1.7\nmock pdf content")

    monkeypatch.setitem(sys.modules, "weasyprint", SimpleNamespace(HTML=FakeHTML))
    monkeypatch.setattr(app, "OUTPUT_DIR", tmp_path)

    row = make_row()
    output_path = app.process_employee(row, excel_row_number=2)

    assert output_path.name == "NV001.pdf"
    assert output_path.is_file()
    assert output_path.read_bytes().startswith(b"%PDF-")
    assert not (tmp_path / ".NV001.tmp.pdf").exists()
