from decimal import Decimal

import pytest

from statistika.catalog import Catalog, RecognitionRule
from statistika.report import ReportRow, refresh_groups
from statistika.statement import Product
from statistika.volume import VOLUME_GROUPS, find_container_liters, parse_container_liters


@pytest.mark.parametrize(("name", "liters"), [
    ('Вино «Саперави» 0,75л.', "0.75"),
    ('Виски 0.7 л в кор.', "0.7"),
    ('Бальзам «Белорусский» 0,2л 40%', "0.2"),
    ('Пиво 500 МЛ', "0.5"),
    ('Вода 1,5\u00a0л', "1.5"),
    ('Wine 750ml', "0.75"),
    ('Вода 5л', "5"),
    ('Вино «Старая келья» п.сл. крас.0,7л', "0.7"),
    ('Вино «Шардоне» сух.бел..0,7л', "0.7"),
    ('Настойка «Коноплянка» перец0,5л', "0.5"),
    ('Каор — 0,75л', "0.75"),
    ('Квас .5 л', "0.5"),
    ('Вода 1 литр', "1"),
    ('Сок 250 миллилитров', "0.25"),
])
def test_one_explicit_container_volume(name, liters):
    assert find_container_liters(name) == Decimal(liters)


@pytest.mark.parametrize("name", [
    'Водка 0,5', 'Настойка 40%', 'Коньяк 12 лет', 'Виски 12л 0,7л в кор.',
    'Набор 6 × 0,5л', 'Пиво 6x0,5л', 'Вода 0,75л x 6', 'Напиток 0,5/1л', 'Вода 0,5–1л',
    'Набор 6 бутылок по 0,5л', 'Пиво 0,5л. × 6', 'Вода 0,5 или 1л', 'От 0,5 до 1л', 'Напиток ½л',
    'Набор 6 банок по 0,5 л', 'Вода 1 500 мл',
    'Вода 0л', 'Вода -1л', 'Вода 0,5литой', 'Водка 0,2,5л', 'Напиток 0..75л', 'Модель B12L',
])
def test_no_guess_at_missing_unit_multiple_sizes_or_multipacks(name):
    assert find_container_liters(name) is None


@pytest.mark.parametrize(("text", "liters"), [
    ("0,7", Decimal("0.7")), (" .5 л ", Decimal("0.5")), (" 750 мл ", Decimal("0.75")),
    ("2", Decimal(2)), ("", None), ("  ", None),
])
def test_manual_input_is_liters_or_an_explicit_milliliter_value(text, liters):
    assert parse_container_liters(text) == liters


@pytest.mark.parametrize("text", ["0", "-0,5", "1/2", "большая", "0,7 0,8", "1E+4", "0,5дал", "∞"])
def test_invalid_manual_input_is_not_applied(text):
    with pytest.raises(ValueError, match="объём тары"):
        parse_container_liters(text)


def test_volume_enabled_group_set():
    assert VOLUME_GROUPS == {35, 36, 37, 38, 39, 40, 41, 42, 43, 45}


@pytest.mark.parametrize("group", sorted(VOLUME_GROUPS))
def test_both_volume_totals_multiply_by_liters_without_dividing_by_ten(group):
    product = Product("Напиток 500 мл", Decimal(12), Decimal(48), Decimal(4), Decimal(16))
    assert ReportRow(product, group).volume_values == (Decimal("0.5"), Decimal(6), Decimal(2))


def test_override_works_for_one_row_and_survives_period_change_but_totals_follow_group():
    product = Product("Вино 0,75л", Decimal(3), None, Decimal(8), None)
    row = ReportRow(product, 35)
    other = ReportRow(product, 35)
    row.manual_volume = parse_container_liters("1,25")
    assert row.volume_values == (Decimal("1.25"), Decimal("3.75"), Decimal(10))
    assert other.volume_values == (Decimal("0.75"), Decimal("2.25"), Decimal(6))
    row.product = Product(product.name, Decimal(0), None, None, None)
    assert row.volume_values == (Decimal("1.25"), Decimal(0), None)
    row.assign(53)
    assert row.volume_values == (Decimal("1.25"), None, None)
    row.assign(45)
    row.manual_volume = None
    assert row.volume_values == (Decimal("0.75"), Decimal(0), None)


@pytest.mark.parametrize("group", [None, 44, 46, 53, 57, 63, 100])
def test_only_calculated_volumes_are_limited_to_target_groups(group):
    product = Product("Другой товар 5л", Decimal(6), None, Decimal(7), None)
    assert ReportRow(product, group).volume_values == (Decimal(5), None, None)


def test_catalog_reclassification_recomputes_totals_without_resetting_the_manual_volume():
    product = Product("Напиток Тестовый 0,5 л", Decimal(2), None, Decimal(3), None)
    row = ReportRow(product, 46, manual_volume=Decimal("0.75"))
    assert row.volume_values == (Decimal("0.75"), None, None)
    refresh_groups([row], Catalog(rules={45: (RecognitionRule("напиток тестовый"),)}))
    assert row.group == 45 and not row.manual
    assert row.volume_values == (Decimal("0.75"), Decimal("1.5"), Decimal("2.25"))
    refresh_groups([row], Catalog())
    assert row.group is None and row.volume_values == (Decimal("0.75"), None, None)


def test_missing_quantity_and_missing_volume_are_not_zeros():
    bare = Product("Напиток без объёма", Decimal(6), None, Decimal(0), None)
    assert ReportRow(bare, 45).volume_values == (None, None, None)
    product = Product("Напиток 3 л", None, None, Decimal(0), None)
    assert ReportRow(product, 45).volume_values == (Decimal(3), None, Decimal(0))
