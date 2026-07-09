from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP, getcontext

from app.schemas.labor import LaborCalculatedRow, LaborInputRow

getcontext().prec = 28

D0 = Decimal("0")
D_01 = Decimal("0.01")
D_06 = Decimal("0.06")
D_08 = Decimal("0.8")
D_084 = Decimal("0.84")
D_076 = Decimal("0.76")
D_068 = Decimal("0.68")
D_20 = Decimal("0.20")
D_30 = Decimal("0.30")
D_40 = Decimal("0.40")


def money2(value: Decimal) -> str:
    """Two-decimal string for display/export without changing internal precision."""
    return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def money2_decimal(value: Decimal) -> Decimal:
    """Two-decimal Decimal for Excel numeric cells."""
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def reverse_pre_tax_from_after_tax(cumulative_after_tax: Decimal) -> Decimal:
    """Reverse monthly cumulative pre-tax amount from monthly cumulative after-tax amount."""
    if cumulative_after_tax <= Decimal("800"):
        return cumulative_after_tax
    if cumulative_after_tax <= Decimal("3360"):
        return (cumulative_after_tax - Decimal("160")) / D_08
    if cumulative_after_tax <= Decimal("21000"):
        return cumulative_after_tax / D_084
    if cumulative_after_tax <= Decimal("49500"):
        return (cumulative_after_tax - Decimal("2000")) / D_076
    return (cumulative_after_tax - Decimal("7000")) / D_068


def cumulative_individual_income_tax(cumulative_pre_tax: Decimal) -> Decimal:
    """Calculate cumulative IIT by labor remuneration rule used by the workbook."""
    if cumulative_pre_tax <= Decimal("800"):
        return D0
    if cumulative_pre_tax <= Decimal("4000"):
        return (cumulative_pre_tax - Decimal("800")) * D_20

    taxable_income = cumulative_pre_tax * D_08
    if taxable_income <= Decimal("20000"):
        return taxable_income * D_20
    if taxable_income <= Decimal("50000"):
        return taxable_income * D_30 - Decimal("2000")
    return taxable_income * D_40 - Decimal("7000")


@dataclass
class MonthlyState:
    cumulative_after_tax: Decimal = D0
    cumulative_iit: Decimal = D0


@dataclass
class DailyState:
    cumulative_after_tax: Decimal = D0
    cumulative_iit: Decimal = D0
    cumulative_contract: Decimal = D0


def calculate_rows(rows: list[LaborInputRow]) -> list[LaborCalculatedRow]:
    """Calculate simplified 0709 ledger rows in input order.

    IIT is calculated by year + month + speaker ID.
    VAT and contract amount are calculated by year + month + day + speaker ID.
    """
    monthly_states: dict[tuple[int, int, str], MonthlyState] = defaultdict(MonthlyState)
    daily_states: dict[tuple[int, int, int, str], DailyState] = defaultdict(DailyState)
    result: list[LaborCalculatedRow] = []

    for index, row in enumerate(rows, start=1):
        monthly_key = (row.year, row.month, row.id_no)
        daily_key = (row.year, row.month, row.day, row.id_no)
        monthly = monthly_states[monthly_key]
        daily = daily_states[daily_key]

        after_tax = row.after_tax_amount

        monthly.cumulative_after_tax += after_tax
        cumulative_pre_tax = reverse_pre_tax_from_after_tax(monthly.cumulative_after_tax)
        current_cumulative_iit = cumulative_individual_income_tax(cumulative_pre_tax)
        iit = current_cumulative_iit - monthly.cumulative_iit
        if iit < D0:
            iit = D0
        monthly.cumulative_iit += iit

        daily.cumulative_after_tax += after_tax
        daily.cumulative_iit += iit
        daily_cumulative_pre_tax = daily.cumulative_after_tax + daily.cumulative_iit
        vat = D0 if daily_cumulative_pre_tax <= Decimal("1000") else daily_cumulative_pre_tax * D_01
        surcharge = vat * D_06
        daily_cumulative_contract = daily_cumulative_pre_tax + vat + surcharge
        contract = daily_cumulative_contract - daily.cumulative_contract
        if contract < D0:
            contract = D0
        daily.cumulative_contract += contract

        result.append(
            LaborCalculatedRow(
                **row.model_dump(),
                index=index,
                cumulative_after_tax_amount=monthly.cumulative_after_tax,
                cumulative_pre_tax_without_vat=cumulative_pre_tax,
                daily_cumulative_pre_tax_without_vat=daily_cumulative_pre_tax,
                contract_amount=contract,
                vat_amount=vat,
                surcharge_amount=surcharge,
                individual_tax_amount=iit,
                cumulative_individual_tax_amount=monthly.cumulative_iit,
            )
        )

    return result
