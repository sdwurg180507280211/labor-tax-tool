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

LOGIC_TEST_ROWS = [
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试A-免税累计", "110101199001010001", 400],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试A-免税累计", "110101199001010001", 300],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试B-跨800", "110101199001010002", 400],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试B-跨800", "110101199001010002", 500],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试C-3360内", "110101199001010003", 3000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试D-21000内", "110101199001010004", 5000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试D-21000内", "110101199001010004", 10000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试E-49500内", "110101199001010005", 22000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试E-49500内", "110101199001010005", 10000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试F-超过49500", "110101199001010006", 50000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试F-超过49500", "110101199001010006", 10000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试G-增值税临界500", "110101199001010007", 500],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试H-增值税超过500", "110101199001010008", 501],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试I-跨月", "110101199001010009", 3000],
    [2026, 7, "测试事业部", "北京", "测试报销人", "测试会计", "测试I-跨月", "110101199001010009", 3000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试J-跨年", "110101199001010010", 3000],
    [2027, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试J-跨年", "110101199001010010", 3000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "张三", "110101199001010011", 3000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "张三", "110101199001010012", 3000],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试K-不同人同月", "110101199001010013", 1800],
    [2026, 6, "测试事业部", "北京", "测试报销人", "测试会计", "测试L-不同人同月", "110101199001010014", 1800],
]

ERROR_TEST_ROWS = [
    [None, 6, "异常事业部", "北京", "报销人", "会计", "缺年份", "110101199001019001", 3000],
    [2026, None, "异常事业部", "北京", "报销人", "会计", "缺月份", "110101199001019002", 3000],
    [2026, 13, "异常事业部", "北京", "报销人", "会计", "月份13", "110101199001019003", 3000],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "缺身份证", None, 3000],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "缺姓名", "110101199001019005", 3000],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "金额为空", "110101199001019006", None],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "金额为零", "110101199001019007", 0],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "金额为负", "110101199001019008", -100],
    [2026, 6, "异常事业部", "北京", "报销人", "会计", "金额文本", "110101199001019009", "abc"],
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


def _build_input_workbook(title: str, rows: list[list[object]], note: str | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(TEMPLATE_HEADERS)
    for row in rows:
        ws.append(row)

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = CENTER
        cell.border = BORDER_THIN
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), min_col=1, max_col=len(TEMPLATE_HEADERS)):
        for cell in row:
            cell.border = BORDER_THIN
            cell.alignment = CENTER
            if cell.column == 8:
                cell.number_format = TEXT_FORMAT
            if cell.column == 9:
                cell.number_format = MONEY_FORMAT
    _set_widths(ws, {"A": 10, "B": 8, "C": 16, "D": 12, "E": 12, "F": 12, "G": 18, "H": 24, "I": 16})
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{max(ws.max_row, 2)}"

    if note:
        note_ws = wb.create_sheet("说明")
        note_ws["A1"] = note
        note_ws["A1"].alignment = LEFT
        note_ws["A1"].font = Font(bold=True)
        note_ws.column_dimensions["A"].width = 100
    return _save_workbook(wb)


def build_template_workbook() -> bytes:
    return _build_input_workbook(
        "基础数据模板",
        [[2026, 6, "事业部A", "北京", "报销人", "会计", "张三", "110101199001011234", 3000]],
        "正式导入模板：保留表头和一行示例。客户正式使用时，可删除示例行后填入真实数据。",
    )


def build_logic_test_workbook() -> bytes:
    return _build_input_workbook(
        "逻辑测试数据",
        LOGIC_TEST_ROWS,
        "逻辑测试模板：覆盖免税、跨800、3360内、21000内、49500内、超过49500、增值税500临界、跨月、跨年、同名不同身份证、不同人同月等正常计算场景。",
    )


def build_error_test_workbook() -> bytes:
    return _build_input_workbook(
        "异常测试数据",
        ERROR_TEST_ROWS,
        "异常测试模板：用于测试校验提示，故意包含年份/月/姓名/身份证/金额为空、月份13、金额为0、负数和文本等错误数据。",
    )


def build_result_workbook(input_rows: list[LaborInputRow]) -> bytes:
    calculated_rows = calculate_rows(input_rows)
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    ws1 = wb.active
    ws1.title = "劳务费税费换算台账"
    _write_original_style_sheet(ws1, calculated_rows)
    ws2 = wb.create_sheet("清晰版台账")
    _write_clear_sheet(ws2, calculated_rows)
    ws3 = wb.create_sheet("公式版台账")
    _write_formula_sheet(ws3, calculated_rows)
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


def _current_iit_formula(pre_tax_ref: str) -> str:
    return (
        f"IF({pre_tax_ref}<=800,0,"
        f"IF({pre_tax_ref}<=4000,({pre_tax_ref}-800)*20%,"
        f"IF({pre_tax_ref}*80%<=20000,{pre_tax_ref}*80%*20%,"
        f"IF({pre_tax_ref}*80%<=50000,{pre_tax_ref}*80%*30%-2000,{pre_tax_ref}*80%*40%-7000))))"
    )


def _prior_sumifs(sum_col: str, row_idx: int) -> str:
    if row_idx <= 6:
        return "0"
    prev = row_idx - 1
    return (
        f"SUMIFS(${sum_col}$6:{sum_col}{prev},"
        f"$B$6:B{prev},B{row_idx},"
        f"$C$6:C{prev},C{row_idx},"
        f"$I$6:I{prev},I{row_idx})"
    )


def _running_sumifs(sum_col: str, row_idx: int) -> str:
    return (
        f"SUMIFS(${sum_col}$6:{sum_col}{row_idx},"
        f"$B$6:B{row_idx},B{row_idx},"
        f"$C$6:C{row_idx},C{row_idx},"
        f"$I$6:I{row_idx},I{row_idx})"
    )


def _write_formula_sheet(ws, rows: Iterable[LaborCalculatedRow]) -> None:
    row_list = list(rows)
    _write_original_style_sheet(ws, row_list)
    ws["A1"] = "标黄处为基础输入；K-X列保留Excel公式"
    ws["J1"] = "手动填写支付的劳务费金额；计算列为公式"

    for row_idx, _item in enumerate(row_list, start=6):
        k = f"K{row_idx}"
        l = f"L{row_idx}"
        q = f"Q{row_idx}"
        r = f"R{row_idx}"
        s = f"S{row_idx}"
        t = f"T{row_idx}"
        u = f"U{row_idx}"
        v = f"V{row_idx}"
        m = f"M{row_idx}"
        n = f"N{row_idx}"
        o = f"O{row_idx}"
        p = f"P{row_idx}"
        x = f"X{row_idx}"

        formulas = {
            k: f"={_running_sumifs('J', row_idx)}",
            l: f"=IF({k}<=800,{k},IF({k}<=3360,({k}-160)/0.8,IF({k}<=21000,{k}/0.84,IF({k}<=49500,({k}-2000)/0.76,({k}-7000)/0.68))))",
            q: f"=IFERROR(IF(J{row_idx}/{k}*{l}<=500,0,J{row_idx}/{k}*{l}*1%),0)",
            r: f"={q}*12%*50%",
            s: f"=MAX(0,{_current_iit_formula(l)}-{_prior_sumifs('S', row_idx)})",
            t: f"={_running_sumifs('Q', row_idx)}",
            u: f"={_running_sumifs('R', row_idx)}",
            v: f"={_running_sumifs('S', row_idx)}",
            m: f"={l}+{t}+{u}",
            n: f"=MAX(0,{m}-{_prior_sumifs('N', row_idx)})",
            o: f"={n}-{s}",
            p: f"={_running_sumifs('O', row_idx)}",
            x: f"=J{row_idx}+{q}+{r}-{o}",
        }
        for coord, formula in formulas.items():
            cell = ws[coord]
            cell.value = formula
            cell.number_format = MONEY_FORMAT
            cell.border = BORDER_THIN
            cell.alignment = CENTER
        ws[f"W{row_idx}"] = None


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
