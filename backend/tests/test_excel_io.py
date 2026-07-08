from decimal import Decimal
from io import BytesIO

from openpyxl import load_workbook
import pytest

from app.schemas.labor import LaborInputRow
from app.services.excel_reader import read_labor_rows
from app.services.excel_writer import (
    build_error_test_workbook,
    build_logic_test_workbook,
    build_result_workbook,
    build_template_workbook,
)


def make_row(amount=3000):
    return LaborInputRow(
        year=2026,
        month=6,
        department="事业部A",
        province="北京",
        reimburser="报销人",
        accountant="会计",
        name="张三",
        id_no="110101199001011234",
        after_tax_amount=Decimal(str(amount)),
    )


def test_template_can_be_read_by_reader():
    data = build_template_workbook()
    rows = read_labor_rows(BytesIO(data))
    assert len(rows) == 1
    assert rows[0].year == 2026
    assert rows[0].month == 6
    assert rows[0].name == "张三"
    assert rows[0].after_tax_amount == Decimal("3000")


def test_logic_test_template_covers_key_calculation_scenarios():
    data = build_logic_test_workbook()
    rows = read_labor_rows(BytesIO(data))
    assert len(rows) == 21

    names = [row.name for row in rows]
    assert "测试A-免税累计" in names
    assert "测试B-跨800" in names
    assert "测试F-超过49500" in names
    assert "测试G-增值税临界500" in names
    assert "测试H-增值税超过500" in names

    cross_month = [row for row in rows if row.id_no == "110101199001010009"]
    assert [row.month for row in cross_month] == [6, 7]

    cross_year = [row for row in rows if row.id_no == "110101199001010010"]
    assert [row.year for row in cross_year] == [2026, 2027]

    same_name_rows = [row for row in rows if row.name == "张三"]
    assert len({row.id_no for row in same_name_rows}) == 2


def test_error_test_template_is_intentionally_invalid():
    data = build_error_test_workbook()
    with pytest.raises(ValueError) as exc_info:
        read_labor_rows(BytesIO(data))
    message = str(exc_info.value)
    assert "第 2 行" in message
    assert "year" in message


def test_result_workbook_has_value_sheets_and_formula_sheet():
    data = build_result_workbook([make_row(400), make_row(500)])
    wb = load_workbook(BytesIO(data), data_only=False)
    assert wb.sheetnames == ["劳务费税费换算台账", "清晰版台账", "公式版台账"]

    ws1 = wb["劳务费税费换算台账"]
    assert ws1["A2"].value == "劳务费税前（后）相关税费换算台账"
    assert ws1["J4"].value == "税后劳务金额"
    assert ws1["X4"].value == "核对"
    assert ws1["A6"].value == 1
    assert ws1["J6"].value == Decimal("400.00")
    assert ws1["K7"].value == Decimal("900.00")
    assert not isinstance(ws1["L7"].value, str) or not ws1["L7"].value.startswith("=")

    ws2 = wb["清晰版台账"]
    assert ws2["A1"].value == "基础信息"
    assert ws2["W2"].value == "核对结果"

    ws3 = wb["公式版台账"]
    assert ws3["A2"].value == "劳务费税前（后）相关税费换算台账"
    assert ws3["A6"].value == 1
    assert ws3["J6"].value == Decimal("400.00")
    assert ws3["K7"].value.startswith("=SUMIFS(")
    assert ws3["L7"].value.startswith("=IF(")
    assert ws3["Q7"].value.startswith("=IFERROR(")
    assert ws3["S7"].value.startswith("=MAX(0,")
    assert ws3["X7"].value == "=J7+Q7+R7-O7"
