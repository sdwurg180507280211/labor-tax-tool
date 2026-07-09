from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any, Optional

from pydantic import BaseModel, Field, field_validator


class LaborInputRow(BaseModel):
    """One raw labor fee input row from Excel upload or manual entry."""

    year: int = Field(..., ge=1900, le=2999, description="年份")
    month: int = Field(..., ge=1, le=12, description="月份")
    day: int = Field(..., ge=1, le=31, description="日期")
    department: Optional[str] = Field(default="", description="事业部")
    name: str = Field(..., min_length=1, description="姓名")
    id_no: str = Field(..., min_length=1, description="讲者ID")
    after_tax_amount: Decimal = Field(..., gt=Decimal("0"), description="税后劳务金额-单次")

    @field_validator("department", mode="before")
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
            raise ValueError("税后劳务金额-单次不能为空")
        try:
            return Decimal(str(value).replace(",", "").strip())
        except (InvalidOperation, AttributeError) as exc:
            raise ValueError("税后劳务金额-单次必须是数字") from exc


class LaborCalculatedRow(LaborInputRow):
    """Calculated ledger row under the simplified 0709 rules."""

    index: int
    cumulative_after_tax_amount: Decimal
    cumulative_pre_tax_without_vat: Decimal
    daily_cumulative_pre_tax_without_vat: Decimal
    contract_amount: Decimal
    vat_amount: Decimal
    surcharge_amount: Decimal
    individual_tax_amount: Decimal
    cumulative_individual_tax_amount: Decimal

    def rounded_dict(self) -> dict[str, Any]:
        """Return frontend/export friendly values rounded to two decimals."""
        from app.services.tax_calculator import money2

        return {
            "index": self.index,
            "year": self.year,
            "month": self.month,
            "day": self.day,
            "department": self.department or "",
            "name": self.name,
            "id_no": self.id_no,
            "after_tax_amount": money2(self.after_tax_amount),
            "cumulative_after_tax_amount": money2(self.cumulative_after_tax_amount),
            "cumulative_pre_tax_without_vat": money2(self.cumulative_pre_tax_without_vat),
            "daily_cumulative_pre_tax_without_vat": money2(self.daily_cumulative_pre_tax_without_vat),
            "contract_amount": money2(self.contract_amount),
            "vat_amount": money2(self.vat_amount),
            "surcharge_amount": money2(self.surcharge_amount),
            "individual_tax_amount": money2(self.individual_tax_amount),
            "cumulative_individual_tax_amount": money2(self.cumulative_individual_tax_amount),
            "debug_info": self.debug_dict(),
        }

    def debug_dict(self) -> dict[str, Any]:
        """Detailed calculation labels used by the frontend test mode."""
        return {
            "monthly_cumulative_key": f"{self.year}-{self.month:02d}-{self.id_no}",
            "daily_cumulative_key": f"{self.year}-{self.month:02d}-{self.day:02d}-{self.id_no}",
            "pre_tax_bracket": self._pre_tax_bracket_label(),
            "individual_tax_bracket": self._individual_tax_bracket_label(self.cumulative_pre_tax_without_vat),
            "vat_rule": "单日累计税前≤1000：增值税=0" if self.daily_cumulative_pre_tax_without_vat <= Decimal("1000") else "单日累计税前>1000：增值税=单日累计税前×1%",
            "threshold_note": self._threshold_note(),
        }

    def _pre_tax_bracket_label(self) -> str:
        amount = self.cumulative_after_tax_amount
        if amount <= Decimal("800"):
            return "本月累计税后≤800：累计税前=累计税后"
        if amount <= Decimal("3360"):
            return "本月累计税后≤3360：累计税前=(累计税后-160)÷0.8"
        if amount <= Decimal("21000"):
            return "本月累计税后≤21000：累计税前=累计税后÷0.84"
        if amount <= Decimal("49500"):
            return "本月累计税后≤49500：累计税前=(累计税后-2000)÷0.76"
        return "本月累计税后>49500：累计税前=(累计税后-7000)÷0.68"

    def _individual_tax_bracket_label(self, cumulative_pre_tax: Decimal) -> str:
        taxable_income = cumulative_pre_tax * Decimal("0.8")
        if cumulative_pre_tax <= Decimal("800"):
            return "累计税前≤800：个税=0"
        if cumulative_pre_tax <= Decimal("4000"):
            return "累计税前≤4000：个税=(累计税前-800)×20%"
        if taxable_income <= Decimal("20000"):
            return "应纳税所得额≤20000：税率20%"
        if taxable_income <= Decimal("50000"):
            return "应纳税所得额≤50000：税率30%，速算扣除2000"
        return "应纳税所得额>50000：税率40%，速算扣除7000"

    def _threshold_note(self) -> str:
        thresholds = []
        for threshold in (Decimal("800"), Decimal("3360"), Decimal("21000"), Decimal("49500")):
            if self.cumulative_after_tax_amount > threshold:
                thresholds.append(str(threshold))
        return "未跨过800" if not thresholds else "本月累计税后已跨过：" + "、".join(thresholds)


class ManualCalculateRequest(BaseModel):
    rows: list[LaborInputRow]
