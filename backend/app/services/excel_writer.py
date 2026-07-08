from __future__ import annotations

from decimal import Decimal
from io import BytesIO
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.schemas.labor import LaborCalculatedRow, LaborInputRow
from app.services.tax_calculator import calculate_rows, money2_decimal

TEMPLATE_HEADERS = [
    "年份",
    "月份",
    "事业部",
    "省区",
    "报销人",
    "会计",
    "姓名",
    "身份证号码",
    "税后劳务金额",
]

YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
SUBHEADER_FILL = PatternFill("solid", fgColor="EAF4FB")
CLEAR_HEADER_FILL = PatternFill("solid", fgColor="D9EAD3")
ERROR_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="000000")
MEDIUM = Side(style="medium", color="000000")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_MEDIUM = Border(left=MEDIUM, right=MEDIUM, top=MEDIUM, bottom=MEDIUM)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
MONEY_FORMAT = '#,##0.00'
INTEGER_FORMAT = '0'
TEXT_FORMAT = '@'


def _save_workbook(wb: Workbook) -> bytes:
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()


def _style_range(ws, row_start: int, row_end: int, col_start: int, col_end: int, border=BORDER_THIN):
    for row in ws.iter_rows(min_row=row_start, max_row=row_end, min_col=col_start, max_col=col_end):
        for cell in row:
            cell.border = border
            cell.alignment = CENTER


def _set_widths(ws, widths: dict[str, float]):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _money(value: Decimal) -> Decimal:
    return money2_decimal(value)


def build_template_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "基础数据模板"
    ws.append(TEMPLATE_HEADERS)
    ws.append([2026, 6, "事业部A", "北京", "报销人", "会计", "张三", "110101199001011234", 3000])

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = CENTER
        cell.border = BORDER_THIN
    for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=len(TEMPLATE_HEADERS)):
        for cell in row:
            cell.border = BORDER_THIN
            cell.alignment = CENTER
    _set_widths(ws, {"A": 10, "B": 8, "C": 16, "D": 12, "E": 12, "F": 12, "G": 12, "H": 24, "I": 16})
    ws.freeze_panes = "A2"
    return _save_workbook(wb)


def build_result_workbook(input_rows: list[LaborInputRow]) -> bytes:
    calculated_rows = calculate_rows(input_rows)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "劳务费税费换算台账"
    _write_original_style_sheet(ws1, calculated_rows)
    ws2 = wb.create_sheet("清晰版台账")
    _write_clear_sheet(ws2, calculated_rows)
    return _save_workbook(wb)


def _write_original_style_sheet(ws, rows: Iterable[LaborCalculatedRow]) -> None:
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:I1")
    ws["A1"] = "标黄处手动填写"
    ws["A1"].fill = YELLOW_FILL
    ws["A1"].alignment = LEFT
    ws["J1"] = "手动填写支付的劳务费金额"
    ws["J1"].fill = YELLOW_FILL
    ws["J1"].alignment = LEFT

    ws.merge_cells("A2:V2")
    ws["A2"] = "劳务费税前（后）相关税费换算台账"
    ws["A2"].font = Font(bold=True, size=16)
    ws["A2"].alignment = CENTER
    ws["S3"] = "单位：元"
    ws["S3"].alignment = CENTER

    merges = [
        "A4:A5", "B4:B5", "C4:C5", "D4:D5", "E4:E5", "F4:F5", "G4:G5", "H4:H5", "I4:I5",
        "J4:K4", "L4:L5", "M4:M5", "N4:N5", "O4:O5", "P4:P5", "Q4:S4", "T4:V4", "X4:X5",
    ]
    for merged in merges:
        ws.merge_cells(merged)

    header_row4 = {
        "A4": "序号",
        "B4": "年份",
        "C4": "月份",
        "D4": "事业部",
        "E4": "省区",
        "F4": "报销人",
        "G4": "会计",
        "H4": "姓名",
        "I4": "身份证号码",
        "J4": "税后劳务金额",
        "L4": "累计税前金额\n-含个税不含增值税",
        "M4": "累计税前金额\n-含个税、增值税",
        "N4": "单次应开票金额（含税）",
        "O4": "单次应付款金额",
        "P4": "本月累计应付款金额",
        "Q4": "单次代扣代缴税额",
        "T4": "本月累计代扣代缴税额",
        "X4": "核对",
    }
    header_row5 = {
        "J5": "单次",
        "K5": "累计",
        "Q5": "增值税",
        "R5": "附加税",
        "S5": "个税",
        "T5": "增值税",
        "U5": "附加税",
        "V5": "个税",
    }
    for coord, value in {**header_row4, **header_row5}.items():
        ws[coord] = value

    _style_range(ws, 4, 5, 1, 24, BORDER_THIN)
    for row in ws.iter_rows(min_row=4, max_row=5, min_col=1, max_col=24):
        for cell in row:
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)

    widths = {
        "A": 6, "B": 10, "C": 8, "D": 12, "E": 10, "F": 12, "G": 12, "H": 12, "I": 22,
        "J": 14, "K": 14, "L": 22, "M": 18, "N": 18, "O": 16, "P": 18,
        "Q": 14, "R": 14, "S": 14, "T": 14, "U": 14, "V": 14, "W": 12, "X": 12,
    }
    _set_widths(ws, widths)
    ws.row_dimensions[1].height = 43.5
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[4].height = 28
    ws.row_dimensions[5].height = 22

    row_idx = 6
    for item in rows:
        values = [
            item.index,
            item.year,
            item.month,
            item.department,
            item.province,
            item.reimburser,
            item.accountant,
            item.name,
            item.id_no,
            _money(item.after_tax_amount),
            _money(item.cumulative_after_tax_amount),
            _money(item.cumulative_pre_tax_without_vat),
            _money(item.cumulative_pre_tax_with_vat_surcharge),
            _money(item.invoice_amount),
            _money(item.payment_amount),
            _money(item.cumulative_payment_amount),
            _money(item.vat_amount),
            _money(item.surcharge_amount),
            _money(item.individual_tax_amount),
            _money(item.cumulative_vat_amount),
            _money(item.cumulative_surcharge_amount),
            _money(item.cumulative_individual_tax_amount),
            None,
            _money(item.check_amount),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row_idx, col_idx, value)
            cell.border = BORDER_THIN
            cell.alignment = CENTER if col_idx not in (4, 5, 6, 7, 8, 9) else LEFT
            if col_idx in (2, 3):
                cell.number_format = INTEGER_FORMAT
            elif col_idx == 9:
                cell.number_format = TEXT_FORMAT
            elif col_idx >= 10 and col_idx != 23:
                cell.number_format = MONEY_FORMAT
            if col_idx in range(1, 11):
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
        if abs(item.check_amount) > Decimal("0.0049"):
            ws.cell(row_idx, 24).fill = ERROR_FILL
        row_idx += 1

    if row_idx == 6:
        ws.cell(6, 1, "无数据")
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:X{max(row_idx - 1, 6)}"


def _write_clear_sheet(ws, rows: Iterable[LaborCalculatedRow]) -> None:
    ws.sheet_view.showGridLines = False
    headers = [
        "序号", "年份", "月份", "事业部", "省区", "报销人", "会计", "姓名", "身份证号码",
        "单次税后劳务金额", "本月累计税后劳务金额",
        "累计税前金额（含个税、不含增值税）", "累计含税开票口径金额",
        "单次个税", "本月累计个税", "单次增值税", "本月累计增值税", "单次附加税", "本月累计附加税",
        "单次应开票金额", "单次应付款金额", "本月累计应付款金额", "核对结果",
    ]
    group_headers = [
        ("基础信息", 1, 9),
        ("输入金额", 10, 11),
        ("税前与开票口径", 12, 13),
        ("税费明细", 14, 19),
        ("开票付款与核对", 20, 23),
    ]
    for title, start_col, end_col in group_headers:
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        cell = ws.cell(1, start_col, title)
        cell.fill = CLEAR_HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = CENTER
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(2, col_idx, header)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = CENTER
        cell.border = BORDER_THIN
    _style_range(ws, 1, 2, 1, len(headers), BORDER_THIN)

    row_idx = 3
    for item in rows:
        values = [
            item.index, item.year, item.month, item.department, item.province, item.reimburser, item.accountant, item.name, item.id_no,
            _money(item.after_tax_amount), _money(item.cumulative_after_tax_amount),
            _money(item.cumulative_pre_tax_without_vat), _money(item.cumulative_pre_tax_with_vat_surcharge),
            _money(item.individual_tax_amount), _money(item.cumulative_individual_tax_amount),
            _money(item.vat_amount), _money(item.cumulative_vat_amount),
            _money(item.surcharge_amount), _money(item.cumulative_surcharge_amount),
            _money(item.invoice_amount), _money(item.payment_amount), _money(item.cumulative_payment_amount), _money(item.check_amount),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row_idx, col_idx, value)
            cell.border = BORDER_THIN
            cell.alignment = CENTER if col_idx not in (4, 5, 6, 7, 8, 9) else LEFT
            if col_idx in (2, 3):
                cell.number_format = INTEGER_FORMAT
            elif col_idx == 9:
                cell.number_format = TEXT_FORMAT
            elif col_idx >= 10:
                cell.number_format = MONEY_FORMAT
            if col_idx == 23 and abs(item.check_amount) > Decimal("0.0049"):
                cell.fill = ERROR_FILL
        row_idx += 1

    widths = {
        "A": 6, "B": 10, "C": 8, "D": 14, "E": 12, "F": 12, "G": 12, "H": 12, "I": 24,
        "J": 16, "K": 18, "L": 24, "M": 22, "N": 14, "O": 16, "P": 14, "Q": 16, "R": 14, "S": 16,
        "T": 18, "U": 18, "V": 20, "W": 12,
    }
    _set_widths(ws, widths)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 36
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:W{max(row_idx - 1, 3)}"
