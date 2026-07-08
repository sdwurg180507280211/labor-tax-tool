from __future__ import annotations

from io import BytesIO
from decimal import Decimal
from typing import Any

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.schemas.labor import LaborInputRow, ManualCalculateRequest
from app.services.excel_reader import read_labor_rows
from app.services.excel_writer import build_result_workbook, build_template_workbook
from app.services.tax_calculator import calculate_rows, money2

router = APIRouter(prefix="/api", tags=["labor-tax"])


def _rows_response(input_rows: list[LaborInputRow]) -> dict[str, Any]:
    calculated = calculate_rows(input_rows)
    zero = Decimal("0")
    total_after_tax = sum((row.after_tax_amount for row in input_rows), start=zero)
    total_invoice = sum((row.invoice_amount for row in calculated), start=zero)
    total_payment = sum((row.payment_amount for row in calculated), start=zero)
    total_iit = sum((row.individual_tax_amount for row in calculated), start=zero)
    return {
        "success": True,
        "input_rows": [row.model_dump(mode="json") for row in input_rows],
        "rows": [row.rounded_dict() for row in calculated],
        "errors": [],
        "summary": {
            "row_count": len(calculated),
            "total_after_tax_amount": money2(total_after_tax) if calculated else "0.00",
            "total_invoice_amount": money2(total_invoice) if calculated else "0.00",
            "total_payment_amount": money2(total_payment) if calculated else "0.00",
            "total_individual_tax_amount": money2(total_iit) if calculated else "0.00",
        },
    }


@router.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@router.get("/template")
def download_template() -> StreamingResponse:
    data = build_template_workbook()
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=labor_fee_input_template.xlsx"},
    )


@router.post("/calculate/upload")
async def calculate_upload(file: UploadFile = File(...)) -> dict[str, Any]:
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="请上传 .xlsx 或 .xlsm 格式的 Excel 文件。")
    try:
        content = await file.read()
        rows = read_labor_rows(BytesIO(content))
        return _rows_response(rows)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@router.post("/calculate/manual")
def calculate_manual(payload: ManualCalculateRequest) -> dict[str, Any]:
    if not payload.rows:
        raise HTTPException(status_code=400, detail="请至少输入一行劳务明细。")
    return _rows_response(payload.rows)


@router.post("/export")
def export_result(payload: ManualCalculateRequest) -> StreamingResponse:
    if not payload.rows:
        raise HTTPException(status_code=400, detail="没有可导出的数据。")
    data = build_result_workbook(payload.rows)
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=labor_fee_tax_ledger.xlsx"},
    )
