from decimal import Decimal

from app.schemas.labor import LaborInputRow
from app.services.tax_calculator import calculate_rows, money2


def row(year, month, id_no, amount, name="张三"):
    return LaborInputRow(
        year=year,
        month=month,
        department="事业部A",
        province="北京",
        reimburser="报销人",
        accountant="会计",
        name=name,
        id_no=id_no,
        after_tax_amount=Decimal(str(amount)),
    )


def test_calculate_rows_cumulative_same_year_month_id():
    result = calculate_rows([
        row(2026, 6, "111", 400),
        row(2026, 6, "111", 500),
        row(2026, 6, "111", 3000),
    ])

    assert money2(result[0].cumulative_after_tax_amount) == "400.00"
    assert money2(result[0].cumulative_pre_tax_without_vat) == "400.00"
    assert money2(result[0].individual_tax_amount) == "0.00"
    assert money2(result[0].payment_amount) == "400.00"
    assert money2(result[0].check_amount) == "0.00"

    assert money2(result[1].cumulative_after_tax_amount) == "900.00"
    assert money2(result[1].cumulative_pre_tax_without_vat) == "925.00"
    assert money2(result[1].vat_amount) == "5.14"
    assert money2(result[1].surcharge_amount) == "0.31"
    assert money2(result[1].individual_tax_amount) == "25.00"
    assert money2(result[1].payment_amount) == "505.45"
    assert money2(result[1].check_amount) == "0.00"

    assert money2(result[2].cumulative_after_tax_amount) == "3900.00"
    assert money2(result[2].cumulative_pre_tax_without_vat) == "4642.86"
    assert money2(result[2].vat_amount) == "35.71"
    assert money2(result[2].surcharge_amount) == "2.14"
    assert money2(result[2].individual_tax_amount) == "717.86"
    assert money2(result[2].payment_amount) == "3037.86"
    assert money2(result[2].check_amount) == "0.00"


def test_cumulative_key_includes_year_and_month():
    result = calculate_rows([
        row(2026, 6, "111", 900),
        row(2026, 7, "111", 900),
        row(2027, 6, "111", 900),
    ])

    assert [money2(x.cumulative_after_tax_amount) for x in result] == ["900.00", "900.00", "900.00"]
    assert [money2(x.individual_tax_amount) for x in result] == ["25.00", "25.00", "25.00"]


def test_rounded_dict_contains_debug_info_for_test_mode():
    result = calculate_rows([row(2026, 6, "111", 400), row(2026, 6, "111", 500)])
    payload = result[1].rounded_dict()
    debug = payload["debug_info"]

    assert debug["cumulative_key"] == "2026-06-111"
    assert debug["pre_tax_bracket"] == "累计税后≤3360：税前=(税后-160)÷0.8"
    assert debug["individual_tax_bracket"] == "累计税前≤4000：个税=(税前-800)×20%"
    assert debug["allocated_pre_tax_amount"] == "513.89"
    assert debug["vat_rule"] == "本次对应税前>500：增值税=本次对应税前×1%"
    assert debug["threshold_note"] == "已跨过：800"
    assert debug["check_ok"] == "是"
