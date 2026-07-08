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
    """Reverse cumulative pre-tax labor amount from cumulative after-tax amount.

    The brackets follow the confirmed original workbook logic.
    """
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
class RunningState:
    cumulative_after_tax: Decimal = D0
    cumulative_vat: Decimal = D0
    cumulative_surcharge: Decimal = D0
    cumulative_iit: Decimal = D0
    cumulative_invoice: Decimal = D0
    cumulative_payment: Decimal = D0


def calculate_rows(rows: list[LaborInputRow]) -> list[LaborCalculatedRow]:
    """Calculate all ledger rows in input order.

    Cumulative key: year + month + id_no. The original sheet uses month + id_no;
    year is intentionally included to avoid cross-year same-month mixing.
    """
    states: dict[tuple[int, int, str], RunningState] = defaultdict(RunningState)
    result: list[LaborCalculatedRow] = []

    for index, row in enumerate(rows, start=1):
        key = (row.year, row.month, row.id_no)
        state = states[key]

        after_tax = row.after_tax_amount
        state.cumulative_after_tax += after_tax
        cumulative_pre_tax = reverse_pre_tax_from_after_tax(state.cumulative_after_tax)

        allocated_pre_tax = (after_tax / state.cumulative_after_tax * cumulative_pre_tax) if state.cumulative_after_tax else D0
        vat = D0 if allocated_pre_tax <= Decimal("500") else allocated_pre_tax * D_01
        surcharge = vat * D_06

        state.cumulative_vat += vat
        state.cumulative_surcharge += surcharge

        cumulative_pre_tax_with_vat_surcharge = cumulative_pre_tax + state.cumulative_vat + state.cumulative_surcharge
        invoice = cumulative_pre_tax_with_vat_surcharge - state.cumulative_invoice
        if invoice < D0:
            invoice = D0
        state.cumulative_invoice += invoice

        current_cumulative_iit = cumulative_individual_income_tax(cumulative_pre_tax)
        iit = current_cumulative_iit - state.cumulative_iit
        if iit < D0:
            iit = D0
        state.cumulative_iit += iit

        payment = invoice - iit
        state.cumulative_payment += payment
        check = after_tax + vat + surcharge - payment

        result.append(
            LaborCalculatedRow(
                **row.model_dump(),
                index=index,
                cumulative_after_tax_amount=state.cumulative_after_tax,
                cumulative_pre_tax_without_vat=cumulative_pre_tax,
                cumulative_pre_tax_with_vat_surcharge=cumulative_pre_tax_with_vat_surcharge,
                invoice_amount=invoice,
                payment_amount=payment,
                cumulative_payment_amount=state.cumulative_payment,
                vat_amount=vat,
                surcharge_amount=surcharge,
                individual_tax_amount=iit,
                cumulative_vat_amount=state.cumulative_vat,
                cumulative_surcharge_amount=state.cumulative_surcharge,
                cumulative_individual_tax_amount=state.cumulative_iit,
                check_amount=check,
            )
        )

    return result
