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
    reimburser: Optional[str] = Field(default="", description="发起人/报销人")
    accountant: Optional[str] = Field(default="", description="会计")
    name: str = Field(..., min_length=1, description="讲者姓名")
    id_no: str = Field(..., min_length=1, description="讲者ID")
    after_tax_amount: Decimal = Field(..., gt=Decimal("0"), description="劳务费（实付金额）")

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
            raise ValueError("劳务费（实付金额）不能为空")
        try:
            return Decimal(str(value).replace(",", "").strip())
        except (InvalidOperation, AttributeError) as exc:
            raise ValueError("劳务费（实付金额）必须是数字") from exc


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
            "debug_info": self.debug_dict(),
        }

    def debug_dict(self) -> dict[str, Any]:
        """Detailed calculation labels used by the frontend test mode."""
        from app.services.tax_calculator import money2

        allocated_pre_tax = (
            self.after_tax_amount / self.cumulative_after_tax_amount * self.cumulative_pre_tax_without_vat
            if self.cumulative_after_tax_amount
            else Decimal("0")
        )
        taxable_income = self.cumulative_pre_tax_without_vat * Decimal("0.8")

        return {
            "cumulative_key": f"{self.year}-{self.month:02d}-{self.id_no}",
            "pre_tax_bracket": self._pre_tax_bracket_label(),
            "individual_tax_bracket": self._individual_tax_bracket_label(taxable_income),
            "allocated_pre_tax_amount": money2(allocated_pre_tax),
            "vat_rule": "本次对应税前≤500：增值税=0" if allocated_pre_tax <= Decimal("500") else "本次对应税前>500：增值税=本次对应税前×1%",
            "threshold_note": self._threshold_note(),
            "check_ok": "是" if abs(self.check_amount) <= Decimal("0.0049") else "否",
        }

    def _pre_tax_bracket_label(self) -> str:
        amount = self.cumulative_after_tax_amount
        if amount <= Decimal("800"):
            return "累计税后≤800：税前=税后"
        if amount <= Decimal("3360"):
            return "累计税后≤3360：税前=(税后-160)÷0.8"
        if amount <= Decimal("21000"):
            return "累计税后≤21000：税前=税后÷0.84"
        if amount <= Decimal("49500"):
            return "累计税后≤49500：税前=(税后-2000)÷0.76"
        return "累计税后>49500：税前=(税后-7000)÷0.68"

    def _individual_tax_bracket_label(self, taxable_income: Decimal) -> str:
        pre_tax = self.cumulative_pre_tax_without_vat
        if pre_tax <= Decimal("800"):
            return "累计税前≤800：个税=0"
        if pre_tax <= Decimal("4000"):
            return "累计税前≤4000：个税=(税前-800)×20%"
        if taxable_income <= Decimal("20000"):
            return "应纳税所得额≤20000：税率20%"
        if taxable_income <= Decimal("50000"):
            return "应纳税所得额≤50000：税率30%，速算扣除2000"
        return "应纳税所得额>50000：税率40%，速算扣除7000"

    def _threshold_note(self) -> str:
        amount = self.cumulative_after_tax_amount
        thresholds = []
        for threshold in (Decimal("800"), Decimal("3360"), Decimal("21000"), Decimal("49500")):
            if amount > threshold:
                thresholds.append(str(threshold))
        if not thresholds:
            return "未跨过800"
        return "已跨过：" + "、".join(thresholds)


class ManualCalculateRequest(BaseModel):
    rows: list[LaborInputRow]
