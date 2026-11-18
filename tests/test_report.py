from decimal import Decimal

import pytest
from openpyxl import load_workbook

from statistika.catalog import Catalog
from statistika.report import RESULT_COLUMNS, ReportRow, display_number, export_xlsx, refresh_groups
from statistika.statement import Product


def test_override_and_export_preserve_order_full_labels_and_blanks(tmp_path):
    catalog = Catalog({35: "Учебное описание 35", 36: "Учебное описание 36"}, {35: ("салфетка",)})
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
