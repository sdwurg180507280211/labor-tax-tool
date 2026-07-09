from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from app.services.excel_writer import BACKEND_IMPORT_HEADERS, build_result_workbook, build_template_workbook


def _backend_row(index: int, name: str, speaker_id: str | None, amount: int | float | str | None, year: int = 2026, month: int = 6, day: int = 1) -> list[object]:
    return [
        f"M{index:04d}", "线下会", "测试事业部", "慢病管理", "测试产品", "测试发起人", "E0001", f"测试会议{index:02d}",
        f"{year}-{month:02d}-{day:02d}", "09:00", "10:00", 30, 28, "是", f"D{index:04d}", name, speaker_id,
        "A", "讲者", "测试医院", "心内科", amount, None, None,
    ]


LOGIC_TEST_ROWS = [
    _backend_row(1, "测试A-免税累计", "SPK001", 400),
    _backend_row(2, "测试A-免税累计", "SPK001", 300),
    _backend_row(3, "测试B-跨800", "SPK002", 400),
    _backend_row(4, "测试B-跨800", "SPK002", 500),
    _backend_row(5, "测试C-3360内", "SPK003", 3000),
    _backend_row(6, "测试D-21000内", "SPK004", 5000),
    _backend_row(7, "测试D-21000内", "SPK004", 10000, 2026, 6, 2),
    _backend_row(8, "测试E-49500内", "SPK005", 22000),
    _backend_row(9, "测试E-49500内", "SPK005", 10000, 2026, 6, 2),
    _backend_row(10, "测试F-超过49500", "SPK006", 50000),
    _backend_row(11, "测试F-超过49500", "SPK006", 10000, 2026, 6, 2),
    _backend_row(12, "测试G-单日1000临界", "SPK007", 1000),
    _backend_row(13, "测试H-单日超过1000", "SPK008", 1001),
    _backend_row(14, "测试I-同日多笔", "SPK009", 500),
    _backend_row(15, "测试I-同日多笔", "SPK009", 600),
    _backend_row(16, "测试J-跨日", "SPK009", 600, 2026, 6, 2),
    _backend_row(17, "测试K-跨月", "SPK009", 600, 2026, 7, 1),
    _backend_row(18, "测试L-跨年", "SPK009", 600, 2027, 6, 1),
    _backend_row(19, "张三", "SPK010", 3000),
    _backend_row(20, "张三", "SPK011", 3000),
    _backend_row(21, "测试M-不同讲者", "SPK012", 1800),
    _backend_row(22, "测试N-不同讲者", "SPK013", 1800),
]

ERROR_TEST_ROWS = [
    _backend_row(1, "缺日期", "ERR001", 3000),
    _backend_row(2, "日期月份错", "ERR002", 3000),
    _backend_row(3, "缺讲者姓名", "ERR003", 3000),
    _backend_row(4, "缺讲者ID", None, 3000),
    _backend_row(5, "金额为空", "ERR005", None),
    _backend_row(6, "金额为零", "ERR006", 0),
    _backend_row(7, "金额为负", "ERR007", -100),
    _backend_row(8, "金额文本", "ERR008", "abc"),
]
ERROR_TEST_ROWS[0][8] = None
ERROR_TEST_ROWS[1][8] = "2026-13-01"
ERROR_TEST_ROWS[2][15] = None

EXPECTED_ROWS = [
    ["T01", "累计税后小于等于800", "同月累计税后700，个税0", "验证免税区间"],
    ["T02", "同人同月跨800", "累计税后900，单次个税25", "验证累计反推和个税拆分"],
    ["T03", "同日多笔", "按同日累计后差额拆协议金额", "验证新版按天累计"],
    ["T04", "跨日", "第二天单日累计重新开始", "验证增值税按天累计"],
    ["T05", "跨月跨年", "换月换年后月累计重新开始", "验证累计维度"],
    ["T06", "同名不同讲者ID", "两条张三分别累计", "验证不按姓名累计"],
]

HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
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
            cell.alignment = LEFT if cell.column in (8, 16, 20, 21, 24) else CENTER
            if cell.row == 1:
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)


def _build_input_workbook(rows: list[list[object]], note: str, expected_rows: list[list[str]] | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Worksheet"
    ws.append(BACKEND_IMPORT_HEADERS)
    for row in rows:
        ws.append(row)
    _style_table(ws, len(BACKEND_IMPORT_HEADERS))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(BACKEND_IMPORT_HEADERS)):
        row[16].number_format = TEXT_FORMAT
        row[21].number_format = MONEY_FORMAT
        row[22].number_format = MONEY_FORMAT
        row[23].number_format = MONEY_FORMAT
    _set_widths(ws, {"A": 12, "B": 12, "C": 14, "D": 12, "E": 14, "F": 12, "G": 16, "H": 26, "I": 14, "J": 14, "K": 14, "L": 16, "M": 14, "N": 14, "O": 18, "P": 14, "Q": 20, "R": 12, "S": 12, "T": 24, "U": 14, "V": 18, "W": 14, "X": 16})
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:X{ws.max_row}"
    if expected_rows:
        expected_ws = wb.create_sheet("预期结果说明")
        expected_ws.append(["场景编号", "场景名称", "预期结果", "验证目的"])
        for row in expected_rows:
            expected_ws.append(row)
        _style_table(expected_ws, 4)
        _set_widths(expected_ws, {"A": 12, "B": 28, "C": 42, "D": 36})
    note_ws = wb.create_sheet("说明")
    note_ws["A1"] = note
    note_ws["A1"].alignment = LEFT
    note_ws["A1"].font = Font(bold=True)
    note_ws.column_dimensions["A"].width = 110
    return _save(wb)


def build_logic_test_workbook() -> bytes:
    return _build_input_workbook(
        LOGIC_TEST_ROWS,
        "逻辑测试模板：字段结构与后台导出表格.xlsx一致；系统只读取会议日期、讲者姓名、讲者ID、劳务费（实付金额）作为计算入口。",
        EXPECTED_ROWS,
    )


def build_error_test_workbook() -> bytes:
    return _build_input_workbook(
        ERROR_TEST_ROWS,
        "异常测试模板：字段结构与后台导出表格.xlsx一致，用于测试日期、讲者姓名、讲者ID、实付金额等错误提示。",
    )
