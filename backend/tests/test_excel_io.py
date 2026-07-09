from decimal import Decimal
from io import BytesIO

from openpyxl import load_workbook
import pytest

from app.schemas.labor import LaborInputRow
from app.services.excel_reader import read_labor_rows
from app.services.test_workbooks import (
    build_error_test_workbook,
    build_logic_test_workbook,
    build_result_workbook,
    build_template_workbook,
)


def make_row(amount=3000, name="张三", id_no="SPK000001", year=2026, month=6, day=1):
    return LaborInputRow(year=year, month=month, day=day, name=name, id_no=id_no, after_tax_amount=Decimal(str(amount)))


def money_cell(cell) -> str:
    return f"{float(cell.value):.2f}"


def test_backend_export_template_can_be_read_by_reader():
    data = build_template_workbook()
    wb = load_workbook(BytesIO(data), data_only=True)
    ws = wb["Worksheet"]
    assert ws["A1"].value == "会议ID"
    assert ws["I1"].value == "会议日期"
    assert ws["P1"].value == "讲者姓名"
    assert ws["Q1"].value == "讲者ID"
    assert ws["V1"].value == "劳务费（实付金额）"

    rows = read_labor_rows(BytesIO(data))
    assert len(rows) == 1
    assert rows[0].year == 2026
    assert rows[0].month == 6
    assert rows[0].day == 1
    assert rows[0].name == "张三"
    assert rows[0].id_no == "SPK000001"
    assert rows[0].after_tax_amount == Decimal("3000")


def test_logic_test_template_matches_backend_export_and_covers_key_scenarios():
    data = build_logic_test_workbook()
    wb = load_workbook(BytesIO(data), data_only=True)
    assert wb.sheetnames == ["Worksheet", "预期结果说明", "说明"]
    ws = wb["Worksheet"]
    assert ws["I1"].value == "会议日期"
    assert ws["P1"].value == "讲者姓名"
    assert ws["Q1"].value == "讲者ID"
    assert ws["V1"].value == "劳务费（实付金额）"

    rows = read_labor_rows(BytesIO(data))
    assert len(rows) == 22
    names = [row.name for row in rows]
    assert "测试I-同日多笔" in names
    assert "测试J-跨日" in names
    assert len([row for row in rows if row.id_no == "SPK009" and row.day == 1]) == 2
    assert len({row.id_no for row in rows if row.name == "张三"}) == 2


def test_error_test_template_reports_invalid_rows_with_readable_messages():
    data = build_error_test_workbook()
    with pytest.raises(ValueError) as exc_info:
        read_labor_rows(BytesIO(data))
    message = str(exc_info.value)
    for row_number in range(2, 10):
        assert f"第 {row_number} 行" in message
    assert "会议日期不能为空" in message
    assert "会议日期月份必须在1-12之间" in message
    assert "讲者姓名不能为空" in message
    assert "讲者ID不能为空" in message
    assert "劳务费（实付金额）不能为空" in message
    assert "劳务费（实付金额）必须大于0" in message
    assert "劳务费（实付金额）必须是数字" in message


def test_result_workbook_has_simplified_ledger_and_rule_sheet():
    data = build_result_workbook([make_row(1000), make_row(1000)])
    wb = load_workbook(BytesIO(data), data_only=False)
    assert wb.sheetnames == ["劳务费税费换算台账", "规则说明"]

    ws = wb["劳务费税费换算台账"]
    assert ws["A1"].value == "劳务费税前（后）相关税费换算台账"
    assert ws["A4"].value == "序号"
    assert ws["D4"].value == "日期"
    assert ws["F4"].value == "讲者ID"
    assert ws["G4"].value == "税后劳务金额"
    assert ws["H5"].value == "本月累计"
    assert ws["J4"].value == "单日累计税前金额\n-含个税不含增值税"
    assert ws["K4"].value == "单次税前金额\n-含个税、增值税"
    assert ws["L5"].value == "增值税"
    assert ws["M5"].value == "附加税"
    assert ws["N5"].value == "个税"

    assert ws["A6"].value == 1
    assert money_cell(ws["G6"]) == "1000.00"
    assert money_cell(ws["H7"]) == "2000.00"
    assert money_cell(ws["I7"]) == "2300.00"
    assert money_cell(ws["J7"]) == "2300.00"
    assert money_cell(ws["K7"]) == "1263.25"
    assert money_cell(ws["L7"]) == "23.00"
    assert money_cell(ws["M7"]) == "1.38"
    assert money_cell(ws["N7"]) == "250.00"

    rule_ws = wb["规则说明"]
    assert rule_ws["A1"].value == "规则项"
    assert "后台导出表格.xlsx" in rule_ws["B2"].value
    assert "年份 + 月份 + 日期 + 讲者ID" in rule_ws["B4"].value


def test_logic_template_export_uses_daily_vat_rule():
    logic_rows = read_labor_rows(BytesIO(build_logic_test_workbook()))
    data = build_result_workbook(logic_rows)
    wb = load_workbook(BytesIO(data), data_only=True)
    ws = wb["劳务费税费换算台账"]

    for row_idx in range(6, ws.max_row + 1):
        if ws.cell(row_idx, 6).value == "SPK008":
            assert ws.cell(row_idx, 12).value > 0
            break
    else:
        raise AssertionError("SPK008 row not found")
