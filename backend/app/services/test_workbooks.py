from __future__ import annotations

from io import BytesIO
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from app.services.excel_writer import TEMPLATE_HEADERS, build_result_workbook, build_template_workbook

LOGIC_TEST_ROWS = [
    [2026, 6, 1, "测试A-免税累计", "SPK001", 400],
    [2026, 6, 1, "测试A-免税累计", "SPK001", 300],
    [2026, 6, 1, "测试B-跨800", "SPK002", 400],
    [2026, 6, 1, "测试B-跨800", "SPK002", 500],
    [2026, 6, 1, "测试C-3360内", "SPK003", 3000],
    [2026, 6, 1, "测试D-21000内", "SPK004", 5000],
    [2026, 6, 2, "测试D-21000内", "SPK004", 10000],
    [2026, 6, 1, "测试E-49500内", "SPK005", 22000],
    [2026, 6, 2, "测试E-49500内", "SPK005", 10000],
    [2026, 6, 1, "测试F-超过49500", "SPK006", 50000],
    [2026, 6, 2, "测试F-超过49500", "SPK006", 10000],
    [2026, 6, 1, "测试G-单日1000临界", "SPK007", 1000],
    [2026, 6, 1, "测试H-单日超过1000", "SPK008", 1001],
    [2026, 6, 1, "测试I-同日多笔", "SPK009", 500],
    [2026, 6, 1, "测试I-同日多笔", "SPK009", 600],
    [2026, 6, 2, "测试J-跨日", "SPK009", 600],
    [2026, 7, 1, "测试K-跨月", "SPK009", 600],
    [2027, 6, 1, "测试L-跨年", "SPK009", 600],
    [2026, 6, 1, "张三", "SPK010", 3000],
    [2026, 6, 1, "张三", "SPK011", 3000],
    [2026, 6, 1, "测试M-不同讲者", "SPK012", 1800],
    [2026, 6, 1, "测试N-不同讲者", "SPK013", 1800],
]

ERROR_TEST_ROWS = [
    [None, 6, 1, "缺年份", "ERR001", 3000],
    [2026, None, 1, "缺月份", "ERR002", 3000],
    [2026, 13, 1, "月份13", "ERR003", 3000],
    [2026, 6, None, "缺日期", "ERR004", 3000],
    [2026, 6, 32, "日期32", "ERR005", 3000],
    [2026, 6, 1, None, "ERR006", 3000],
    [2026, 6, 1, "缺讲者ID", None, 3000],
    [2026, 6, 1, "金额为空", "ERR008", None],
    [2026, 6, 1, "金额为零", "ERR009", 0],
    [2026, 6, 1, "金额为负", "ERR010", -100],
    [2026, 6, 1, "金额文本", "ERR011", "abc"],
]

EXPECTED_ROWS = [
    ["T01", "累计税后≤800", "同月累计税后700，个税0", "验证免税区间"],
    ["T02", "同人同月跨800", "累计税后900，单次个税25", "验证累计反推和个税拆分"],
    ["T03", "累计税后≤3360", "累计税前3550", "验证（税后-160）÷0.8"],
    ["T04", "累计税后3360-21000", "累计税前17857.14", "验证税后÷0.84"],
    ["T05", "累计税后21000-49500", "累计税前39473.68", "验证（税后-2000）÷0.76"],
    ["T06", "累计税后＞49500", "累计税前77941.18", "验证（税后-7000）÷0.68"],
    ["T07", "单日累计税前1000临界", "单日累计税前1000时增值税0", "验证≤1000不计增值税"],
    ["T08", "单日累计税前超过1000", "超过1000后按1%计增值税", "验证新版增值税规则"],
    ["T09", "同一讲者同一天多笔", "按同日累计后差额拆协议金额", "验证协议签订金额差额法"],
    ["T10", "同一讲者跨日", "第二天单日累计重新开始", "验证增值税按天累计"],
    ["T11", "跨月/跨年", "换月换年后月累计重新开始", "验证累计维度"],
    ["T12", "同名不同讲者ID", "两条张三分别累计", "验证不按姓名累计"],
]

HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
PASS_FILL = PatternFill("solid", fgColor="E2F0D9")
ERROR_FILL = PatternFill("solid", fgColor="FCE4D6")
BORDER = Border(*(Side(style="thin", color="000000"),) * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
MONEY_FORMAT = "#,##0.00"
TEXT_FORMAT = "@"


def _save(wb: Workbook) -> bytes:
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()


def _set_widths(ws, widths: dict[str, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _style_table(ws, max_col: int) -> None:
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border = BORDER
            cell.alignment = LEFT if cell.column in (4, 5, 7) else CENTER
            if cell.row == 1:
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)


def _build_input_workbook(title: str, rows: list[list[object]], note: str, expected_rows: list[list[str]] | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(TEMPLATE_HEADERS)
    for row in rows:
        ws.append(row)
    _style_table(ws, len(TEMPLATE_HEADERS))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(TEMPLATE_HEADERS)):
        row[4].number_format = TEXT_FORMAT
        row[5].number_format = MONEY_FORMAT
    _set_widths(ws, {"A": 10, "B": 8, "C": 8, "D": 18, "E": 20, "F": 20})
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"

    if expected_rows:
        expected_ws = wb.create_sheet("预期结果说明")
        expected_ws.append(["场景编号", "场景名称", "预期结果", "验证目的"])
        for row in expected_rows:
            expected_ws.append(row)
        _style_table(expected_ws, 4)
        _set_widths(expected_ws, {"A": 12, "B": 28, "C": 42, "D": 36})
        expected_ws.freeze_panes = "A2"
        expected_ws.auto_filter.ref = f"A1:D{expected_ws.max_row}"

    note_ws = wb.create_sheet("说明")
    note_ws["A1"] = note
    note_ws["A1"].alignment = LEFT
    note_ws["A1"].font = Font(bold=True)
    note_ws.column_dimensions["A"].width = 110
    return _save(wb)


def build_logic_test_workbook() -> bytes:
    return _build_input_workbook(
        "逻辑测试数据",
        LOGIC_TEST_ROWS,
        "逻辑测试模板：基于新版简版0709规则，重点覆盖个税按月累计、增值税按天累计、单日1000临界、同日多笔、跨日、跨月、跨年、同名不同讲者ID等场景。",
        EXPECTED_ROWS,
    )


def build_error_test_workbook() -> bytes:
    return _build_input_workbook(
        "异常测试数据",
        ERROR_TEST_ROWS,
        "异常测试模板：用于测试校验提示，故意包含年份/月/日期/姓名/讲者ID/金额为空，月份和日期超范围，金额为0、负数和文本等错误数据。",
    )
