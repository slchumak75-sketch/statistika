from decimal import Decimal

import pytest
from openpyxl import load_workbook

from statistika.catalog import Catalog, RecognitionRule
from statistika.report import RESULT_COLUMNS, ReportRow, display_number, export_xlsx, refresh_groups
from statistika.statement import Product


def test_override_and_export_preserve_order_full_labels_and_blanks(tmp_path):
    catalog = Catalog({35: "Учебное описание 35", 36: "Учебное описание 36"}, {35: (RecognitionRule("салфетка"),)})
    rows = [
        ReportRow(Product("Носки", Decimal("2"), Decimal("4.40"), Decimal("11"), Decimal("24.20"))),
        ReportRow(Product("Салфетка", Decimal("0"), None, Decimal("0"), None)),
        ReportRow(Product("=ИМЯ_ТОВАРА()", None, None, Decimal("1"), Decimal("97.01"))),
    ]
    refresh_groups(rows, catalog)
    assert [row.group for row in rows] == [100, 35, None]
    rows[1].assign(36)
    refresh_groups(rows, catalog)
    assert [row.group for row in rows] == [100, 36, None]
    assert [row.manual for row in rows] == [False, True, False]
    with pytest.raises(ValueError):
        rows[2].assign(44)

    target = tmp_path / "Результат.xlsx"
    export_xlsx(target, rows, catalog)
    workbook = load_workbook(target)
    try:
        sheet = workbook.active
        assert tuple(sheet.values) == (
            RESULT_COLUMNS,
            ("№100 — Чулочно-носочные изделия", "Носки", 2, 4.4, 11, 24.2),
            ("№36 — Учебное описание 36", "Салфетка", 0, None, 0, None),
            (None, "=ИМЯ_ТОВАРА()", None, None, 1, 97.01),
        )
        assert sheet["B4"].data_type == "s"
        assert sheet["D2"].data_type == "n"
        assert sheet.freeze_panes == "C2"
    finally:
        workbook.close()
    assert list(tmp_path.iterdir()) == [target]


def test_manual_clear_remains_blank_and_no_amount_rounding_in_ui():
    row = ReportRow(Product("Чулки", None, None, Decimal("0"), Decimal("2.005")))
    refresh_groups([row], Catalog())
    assert row.group == 100
    row.assign(None)
    refresh_groups([row], Catalog())
    assert row.group is None and row.manual
    assert display_number(row.product.ending_amount) == "2,005"
    assert display_number(row.product.ending_quantity) == "0"
    assert display_number(None) == ""


def test_export_warehouse_scope_and_summary_add_all_units_as_numbers(tmp_path):
    from statistika.catalog import GROUPS, builtin_catalog
    from statistika.report import summarize_groups, warehouses
    from statistika.statement import ReportPeriod

    catalog = builtin_catalog()
    products = [
        Product("Шуруп", Decimal(0), Decimal(0), Decimal(8), Decimal(16), "Мир", "шт."),
        Product("Клей Бустилат 0,9кг", Decimal(3), Decimal(15), Decimal(4), Decimal(20), "Мир", "банка"),
        Product("Доска строганная", Decimal("2.5"), Decimal(10), Decimal("4.5"), Decimal(20), "Мир", "м"),
        Product("Дюбель", None, None, Decimal(2), None, "Мир", "шт."),
        Product("Шайба", Decimal(2), Decimal(4), Decimal(5), Decimal(10), "Мир", ""),
        Product("Саморез", Decimal(5), Decimal(10), Decimal(11), Decimal(22), "Кореличи", "шт."),
        Product("Другой товар", None, None, None, None, "Мир"),
    ]
    rows = [ReportRow(product) for product in products]
    refresh_groups(rows, catalog)
    assert warehouses(rows) == ("Кореличи", "Мир")
    index = GROUPS.index(53)
    summary = summarize_groups(row for row in rows if row.product.warehouse == "Мир")[index]
    assert summary.group == 53 and summary.lines == 5
    assert [total.value for total in summary.metrics] == [Decimal("7.5"), Decimal(29), Decimal("23.5"), Decimal(66)]
    assert [total.known for total in summary.metrics] == [4, 4, 5, 4]
    assert [total.partial for total in summary.metrics] == [True, True, False, True]

    target = tmp_path / "Склад Мир.xlsx"
    export_xlsx(target, rows, catalog, period=ReportPeriod("Ноябрь 2026", ("Ноябрь 2026",), "15.11.2026"), warehouse="Мир")
    workbook = load_workbook(target)
    try:
        detail, group_sheet = workbook.worksheets
        assert detail.max_column == 6 and detail.max_row == 7
        assert [row[1] for row in detail.iter_rows(min_row=2, values_only=True)] == [p.name for p in products if p.warehouse == "Мир"]
        assert detail["C3"].data_type == "n" and detail["C3"].value == 3 and ' банка"' in detail["C3"].number_format
        assert detail["E4"].data_type == "n" and detail["E4"].value == 4.5 and ' м"' in detail["E4"].number_format
        assert detail["A7"].value is None
        assert all("шт." in detail[col].value for col in ("C1", "E1"))
        assert all("руб." in detail[col].value for col in ("D1", "F1"))
        assert detail.auto_filter.ref == "A1:F7"
        assert group_sheet["B1"].value == "Мир" and group_sheet["B2"].value == "Ноябрь 2026"
        assert group_sheet["B3"].value == "15.11.2026"
        report_line = group_sheet[6 + index]
        assert [cell.value for cell in report_line] == ["№53 — Строительные материалы", 5, 7.5, 29, 23.5, 66]
        assert all(cell.data_type == "n" for cell in report_line[1:])
        assert "*" in report_line[2].number_format and "*" not in report_line[4].number_format
        assert report_line[2].comment and report_line[4].comment is None
        assert group_sheet[group_sheet.max_row][0].value == "Без группы"
    finally:
        workbook.close()


def test_quantity_without_a_unit_counts_but_unknown_values_stay_blank():
    from statistika.catalog import GROUPS
    from statistika.report import summarize_groups

    rows = [ReportRow(Product("Товар", Decimal(7), Decimal(9), None, None), 53)] * 2
    total = summarize_groups(rows)[GROUPS.index(53)]
    assert total.sold_quantity.value == 14 and not total.sold_quantity.partial
    assert total.sold_amount.value == 18
    assert total.ending_quantity.value is None and total.ending_amount.value is None


def test_construction_uses_one_group_label_and_preserves_manual_clear(tmp_path):
    from statistika.catalog import GROUPS, builtin_catalog

    catalog = builtin_catalog()
    rows = [ReportRow(Product(name, None, None, None, None)) for name in (
        "Автомат ВА 47-29 1п 16А IEK", "Американка 1/2 Kalde", "Розетка 7-202-01 с/з",
    )]
    refresh_groups(rows, catalog)
    assert [row.group for row in rows] == [53, 53, 53]
    rows[2].assign(None)
    refresh_groups(rows, catalog)
    assert [row.group for row in rows] == [53, 53, None] and rows[2].manual
    target = tmp_path / "Стройматериалы.xlsx"
    export_xlsx(target, rows, catalog)
    workbook = load_workbook(target)
    try:
        detail, summary = workbook.worksheets
        assert detail.max_column == 6
        assert detail["A2"].value == detail["A3"].value == catalog.label(53)
        assert detail["A1"].comment is None and detail["A4"].value is None
        assert tuple(cell.value for cell in summary[6 + GROUPS.index(53)][:2]) == (catalog.short_label(53), 2)
    finally:
        workbook.close()
