from __future__ import annotations

from typing import Any, BinaryIO

from openpyxl import load_workbook
from pydantic import ValidationError

from app.schemas.labor import LaborInputRow

HEADER_ALIASES: dict[str, str] = {
    "年份": "year",
    "年": "year",
    "月份": "month",
    "月": "month",
    "日期": "day",
    "日": "day",
    "事业部": "department",
    "姓名": "name",
    "讲者姓名": "name",
    "身份证号码": "id_no",
    "身份证号": "id_no",
    "身份证": "id_no",
    "讲者id": "id_no",
    "讲者编号": "id_no",
    "税后劳务金额": "after_tax_amount",
    "税后劳务金额单次": "after_tax_amount",
    "税后劳务金额-单次": "after_tax_amount",
    "税后劳务金额（单次）": "after_tax_amount",
    "单次税后劳务金额": "after_tax_amount",
    "劳务费（实付金额）": "after_tax_amount",
    "劳务费(实付金额)": "after_tax_amount",
    "实付金额": "after_tax_amount",
}

REQUIRED_FIELDS = {"year", "month", "day", "name", "id_no", "after_tax_amount"}
FIELD_LABELS = {
    "year": "年份",
    "month": "月份",
    "day": "日期",
    "name": "姓名",
    "id_no": "身份证号码/讲者ID",
    "after_tax_amount": "税后劳务金额",
}


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    return (
        text.strip()
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("_", "")
        .lower()
    )


def find_header_row(ws) -> tuple[int, dict[int, str]]:
    """Find the customer simplified ledger header row without requiring a changed template."""
    for row_idx in range(1, min(ws.max_row, 20) + 1):
        mapping: dict[int, str] = {}
        for cell in ws[row_idx]:
            normalized = normalize_header(cell.value)
            if normalized in HEADER_ALIASES:
                mapping[cell.column] = HEADER_ALIASES[normalized]
        if len(set(mapping.values())) >= 4:
            return row_idx, mapping
    raise ValueError("未识别到客户简版台账表头。请确保包含：年份、月份、日期、姓名、身份证号码/讲者ID、税后劳务金额。")


def _format_validation_error(err: dict[str, Any]) -> str:
    field = str(err.get("loc", [""])[0])
    message = str(err.get("msg", "校验失败"))
    error_type = str(err.get("type", ""))

    if field == "year":
        if err.get("input") is None or error_type == "missing":
            return "年份不能为空"
        return "年份必须在1900-2999之间"
    if field == "month":
        if err.get("input") is None or error_type == "missing":
            return "月份不能为空"
        return "月份必须在1-12之间"
    if field == "day":
        if err.get("input") is None or error_type == "missing":
            return "日期不能为空"
        return "日期必须在1-31之间"
    if field == "name":
        return "姓名不能为空"
    if field == "id_no":
        return "身份证号码/讲者ID不能为空"
    if field == "after_tax_amount":
        if "不能为空" in message:
            return "税后劳务金额不能为空"
        if "必须是数字" in message:
            return "税后劳务金额必须是数字"
        if "greater than" in message:
            return "税后劳务金额必须大于0"
        return "税后劳务金额格式不正确"
    return f"{FIELD_LABELS.get(field, field)}：{message}"


def read_labor_rows(file_obj: BinaryIO) -> list[LaborInputRow]:
    wb = load_workbook(file_obj, data_only=True)
    ws = wb.active
    header_row, col_to_field = find_header_row(ws)
    found_fields = set(col_to_field.values())
    missing = REQUIRED_FIELDS - found_fields
    if missing:
        raise ValueError("缺少必填列：" + "、".join(FIELD_LABELS[x] for x in sorted(missing)))

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
            msg = "；".join(_format_validation_error(err) for err in exc.errors())
            errors.append(f"第 {excel_row_idx} 行：{msg}")

    if errors:
        raise ValueError("\n".join(errors))
    if not rows:
        raise ValueError("没有读取到有效数据行。")
    return rows
