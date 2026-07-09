from decimal import Decimal

from app.schemas.labor import LaborInputRow
from app.services.tax_calculator import calculate_rows, money2


def row(year, month, day, id_no, amount, name="张三"):
    return LaborInputRow(
        year=year,
        month=month,
        day=day,
        name=name,
        id_no=id_no,
        after_tax_amount=Decimal(str(amount)),
    )


def test_monthly_iit_and_daily_vat_follow_simplified_0709_rules():
    result = calculate_rows([
        row(2026, 6, 1, "111", 1000),
        row(2026, 6, 1, "111", 1000),
        row(2026, 6, 2, "111", 1000),
        row(2026, 6, 2, "111", 5000),
        row(2026, 6, 2, "111", 1000),
    ])

    assert money2(result[0].cumulative_after_tax_amount) == "1000.00"
    assert money2(result[0].cumulative_pre_tax_without_vat) == "1050.00"
    assert money2(result[0].daily_cumulative_pre_tax_without_vat) == "1050.00"
    assert money2(result[0].contract_amount) == "1061.13"
    assert money2(result[0].vat_amount) == "10.50"
    assert money2(result[0].surcharge_amount) == "0.63"
    assert money2(result[0].individual_tax_amount) == "50.00"

    assert money2(result[1].cumulative_after_tax_amount) == "2000.00"
    assert money2(result[1].cumulative_pre_tax_without_vat) == "2300.00"
    assert money2(result[1].daily_cumulative_pre_tax_without_vat) == "2300.00"
    assert money2(result[1].contract_amount) == "1263.25"
    assert money2(result[1].vat_amount) == "23.00"
    assert money2(result[1].surcharge_amount) == "1.38"
    assert money2(result[1].individual_tax_amount) == "250.00"

    assert money2(result[2].cumulative_after_tax_amount) == "3000.00"
    assert money2(result[2].cumulative_pre_tax_without_vat) == "3550.00"
    assert money2(result[2].daily_cumulative_pre_tax_without_vat) == "1250.00"
    assert money2(result[2].contract_amount) == "1263.25"
    assert money2(result[2].vat_amount) == "12.50"
    assert money2(result[2].surcharge_amount) == "0.75"
    assert money2(result[2].individual_tax_amount) == "250.00"

    assert money2(result[3].cumulative_after_tax_amount) == "8000.00"
    assert money2(result[3].cumulative_pre_tax_without_vat) == "9523.81"
    assert money2(result[3].daily_cumulative_pre_tax_without_vat) == "7223.81"
    assert money2(result[3].contract_amount) == "6037.13"
    assert money2(result[3].vat_amount) == "72.24"
    assert money2(result[3].surcharge_amount) == "4.33"
    assert money2(result[3].individual_tax_amount) == "973.81"

    assert money2(result[4].cumulative_after_tax_amount) == "9000.00"
    assert money2(result[4].cumulative_pre_tax_without_vat) == "10714.29"
    assert money2(result[4].daily_cumulative_pre_tax_without_vat) == "8414.29"
    assert money2(result[4].contract_amount) == "1203.10"
    assert money2(result[4].vat_amount) == "84.14"
    assert money2(result[4].surcharge_amount) == "5.05"
    assert money2(result[4].individual_tax_amount) == "190.48"


def test_vat_resets_by_day_but_iit_keeps_monthly_accumulation():
    result = calculate_rows([
        row(2026, 6, 1, "111", 600),
        row(2026, 6, 1, "111", 600),
        row(2026, 6, 2, "111", 600),
    ])

    assert money2(result[0].vat_amount) == "0.00"
    assert money2(result[1].daily_cumulative_pre_tax_without_vat) == "1300.00"
    assert money2(result[1].vat_amount) == "13.00"
    assert money2(result[2].daily_cumulative_pre_tax_without_vat) == "750.00"
    assert money2(result[2].vat_amount) == "0.00"
    assert money2(result[2].cumulative_after_tax_amount) == "1800.00"
    assert money2(result[2].individual_tax_amount) == "150.00"


def test_monthly_key_includes_year_month_and_speaker_id():
    result = calculate_rows([
        row(2026, 6, 1, "111", 900),
        row(2026, 7, 1, "111", 900),
        row(2027, 6, 1, "111", 900),
    ])

    assert [money2(x.cumulative_after_tax_amount) for x in result] == ["900.00", "900.00", "900.00"]
    assert [money2(x.individual_tax_amount) for x in result] == ["25.00", "25.00", "25.00"]


def test_rounded_dict_contains_debug_info_for_test_mode():
    result = calculate_rows([row(2026, 6, 1, "111", 600), row(2026, 6, 1, "111", 600)])
    payload = result[1].rounded_dict()
    debug = payload["debug_info"]

    assert debug["monthly_cumulative_key"] == "2026-06-111"
    assert debug["daily_cumulative_key"] == "2026-06-01-111"
    assert debug["pre_tax_bracket"] == "本月累计税后≤3360：累计税前=(累计税后-160)÷0.8"
    assert debug["individual_tax_bracket"] == "累计税前≤4000：个税=(累计税前-800)×20%"
    assert debug["vat_rule"] == "单日累计税前>1000：增值税=单日累计税前×1%"
