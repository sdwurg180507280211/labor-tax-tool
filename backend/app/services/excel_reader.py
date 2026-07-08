from __future__ import annotations

from typing import Any, BinaryIO

from openpyxl import load_workbook
from pydantic import ValidationError

from app.schemas.labor import LaborInputRow

HEADER_ALIASES: dict[str, str] = {
    "年份": "year",
    "年": "year",
    "year": "year",
    "月份": "month",
    "月": "month",
    "month": "month",
    "事业部": "department",
    "部门": "department",
    "department": "department",
    "省区": "province",
    "省份": "province",
    "区域": "province",
    "province": "province",
    "报销人": "reimburser",
    "reimburser": "reimburser",
    "会计": "accountant",
    "accountant": "accountant",
    "姓名": "name",
    "名称": "name",
    "name": "name",
    "身份证号码": "id_no",
    "身份证号": "id_no",
    "身份证": "id_no",
    "id_no": "id_no",
    "idno": "id_no",
    "税后劳务金额": "after_tax_amount",
    "税后劳务金额单次": "after_tax_amount",
    "税后劳务金额-单次": "after_tax_amount",
    "税后劳务金额（单次）": "after_tax_amount",
    "单次税后劳务金额": "after_tax_amount",
    "劳务费金额": "after_tax_amount",
    "after_tax_amount": "after_tax_amount",
}

REQUIRED_FIELDS = {"year", "month", "name", "id_no", "after_tax_amount"}


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    return (
        text.strip()
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("_", "")
        .replace("-", "-")
        .lower()
    )


def find_header_row(ws) -> tuple[int, dict[int, str]]:
    """Find the first row that contains at least 3 known headers."""
    for row_idx in range(1, min(ws.max_row, 20) + 1):
        mapping: dict[int, str] = {}
        for cell in ws[row_idx]:
            normalized = normalize_header(cell.value)
            if normalized in HEADER_ALIASES:
                mapping[cell.column] = HEADER_ALIASES[normalized]
        if len(set(mapping.values())) >= 3:
            return row_idx, mapping
    raise ValueError("未识别到表头。请使用模板，或确保包含：年份、月份、姓名、身份证号码、税后劳务金额。")


def read_labor_rows(file_obj: BinaryIO) -> list[LaborInputRow]:
    wb = load_workbook(file_obj, data_only=True)
    ws = wb.active
    header_row, col_to_field = find_header_row(ws)
    found_fields = set(col_to_field.values())
    missing = REQUIRED_FIELDS - found_fields
    if missing:
        names = {
            "year": "年份",
            "month": "月份",
            "name": "姓名",
            "id_no": "身份证号码",
            "after_tax_amount": "税后劳务金额",
        }
        raise ValueError("缺少必填列：" + "、".join(names[x] for x in sorted(missing)))

    rows: list[LaborInputRow] = []
    errors: list[str] = []
    for excel_row_idx in range(header_row + 1, ws.max_row + 1):
        raw: dict[str, Any] = {}
        empty = True
        for col_idx, field in col_to_field.items():
            value = ws.cell(excel_row_idx, col_idx).value
            if value not in (None, ""):
                empty = False
            raw[field] = value
        if empty:
            continue
        try:
            rows.append(LaborInputRow(**raw))
        except ValidationError as exc:
            msg = "; ".join(f"{'.'.join(map(str, err['loc']))}: {err['msg']}" for err in exc.errors())
            errors.append(f"第 {excel_row_idx} 行：{msg}")

    if errors:
        raise ValueError("\n".join(errors))
    if not rows:
        raise ValueError("没有读取到有效数据行。")
    return rows
