from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any, Optional

from pydantic import BaseModel, Field, field_validator


class LaborInputRow(BaseModel):
    """One raw labor fee input row from Excel upload or manual entry."""

    year: int = Field(..., ge=1900, le=2999, description="年份")
    month: int = Field(..., ge=1, le=12, description="月份")
    department: Optional[str] = Field(default="", description="事业部")
    province: Optional[str] = Field(default="", description="省区")
    reimburser: Optional[str] = Field(default="", description="报销人")
    accountant: Optional[str] = Field(default="", description="会计")
    name: str = Field(..., min_length=1, description="姓名")
    id_no: str = Field(..., min_length=1, description="身份证号码")
    after_tax_amount: Decimal = Field(..., gt=Decimal("0"), description="税后劳务金额")

    @field_validator("department", "province", "reimburser", "accountant", mode="before")
    @classmethod
    def normalize_optional_text(cls, value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @field_validator("name", "id_no", mode="before")
    @classmethod
    def normalize_required_text(cls, value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @field_validator("after_tax_amount", mode="before")
    @classmethod
    def parse_decimal(cls, value: Any) -> Decimal:
        if value is None or value == "":
            raise ValueError("税后劳务金额不能为空")
        try:
            return Decimal(str(value).replace(",", "").strip())
        except (InvalidOperation, AttributeError) as exc:
            raise ValueError("税后劳务金额必须是数字") from exc


class LaborCalculatedRow(LaborInputRow):
    """Calculated ledger row. All money fields are high precision Decimal."""

    index: int
    cumulative_after_tax_amount: Decimal
    cumulative_pre_tax_without_vat: Decimal
    cumulative_pre_tax_with_vat_surcharge: Decimal
    invoice_amount: Decimal
    payment_amount: Decimal
    cumulative_payment_amount: Decimal
    vat_amount: Decimal
    surcharge_amount: Decimal
    individual_tax_amount: Decimal
    cumulative_vat_amount: Decimal
    cumulative_surcharge_amount: Decimal
    cumulative_individual_tax_amount: Decimal
    check_amount: Decimal

    def rounded_dict(self) -> dict[str, Any]:
        """Return frontend/export friendly values rounded to two decimals."""
        from app.services.tax_calculator import money2

        return {
            "index": self.index,
            "year": self.year,
            "month": self.month,
            "department": self.department or "",
            "province": self.province or "",
            "reimburser": self.reimburser or "",
            "accountant": self.accountant or "",
            "name": self.name,
            "id_no": self.id_no,
            "after_tax_amount": money2(self.after_tax_amount),
            "cumulative_after_tax_amount": money2(self.cumulative_after_tax_amount),
            "cumulative_pre_tax_without_vat": money2(self.cumulative_pre_tax_without_vat),
            "cumulative_pre_tax_with_vat_surcharge": money2(self.cumulative_pre_tax_with_vat_surcharge),
            "invoice_amount": money2(self.invoice_amount),
            "payment_amount": money2(self.payment_amount),
            "cumulative_payment_amount": money2(self.cumulative_payment_amount),
            "vat_amount": money2(self.vat_amount),
            "surcharge_amount": money2(self.surcharge_amount),
            "individual_tax_amount": money2(self.individual_tax_amount),
            "cumulative_vat_amount": money2(self.cumulative_vat_amount),
            "cumulative_surcharge_amount": money2(self.cumulative_surcharge_amount),
            "cumulative_individual_tax_amount": money2(self.cumulative_individual_tax_amount),
            "check_amount": money2(self.check_amount),
        }


class ManualCalculateRequest(BaseModel):
    rows: list[LaborInputRow]
