from datetime import date
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

from statistika.catalog import builtin_catalog
from statistika.report import RESULT_COLUMNS, ReportRow, export_xlsx, refresh_groups
from statistika.statement import Sheet, SourcePeriod, StatementError, read_sheet
from statistika.warehouse import MonthColumn, SalesQuantity, WarehouseLayout, detect_warehouse_layout, import_warehouse_products

_HEADERS = (
    "Склад", "Группа номенклатуры", "Артикул", "Код", "Номенклатура", "Ед.", "Начальный\nостаток",
    "Приход", "Расход", "Конечный\nостаток", "Минимальный резерв", "Цена последнего поставщика", "Оптовая надбавка",
    "Розничная \nцена", "Последний поставщик", "+", "Октябрь 2026 г.", "Ноябрь 2026 г.",
)


def stock_row(name, outflow, ending, price, monthly, section="Хозгруппа", warehouse="Тестовый склад", unit="шт"):
    return (warehouse, section, None, None, name, unit, 0, 6, outflow, ending,
            None, "100.00", None, price, None, None, *monthly)


def test_warehouse_profile_through_xlsx_and_report(tmp_path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Лист_1"
    worksheet.append([None, None, None, None, "Ведомость по остаткам ТМЦ на складах"])
    worksheet.append([None, "С 01.10.26 по 15.11.26"])
    worksheet.append(_HEADERS)
    worksheet.append(stock_row("Муфта PPR 40 мм", "7.00", "2.5", "4,80", (3, 6), "3 металлопласт"))
    worksheet.append(stock_row("Носки женские", "4.00", 0, None, (1, 3)))
    worksheet.append(stock_row("Кастрюля", "2.00", 4, "7.30\n9.76", (0, 2)))
    worksheet.append(stock_row("Пиво светлое 5,0%", "1.00", 3, None, (0, 1)))
    worksheet.append(stock_row("Итого", 14, 10, 50, (4, 12)))
    worksheet.append(_HEADERS)
    path = tmp_path / "Ведомость.xlsx"
    workbook.save(path)
    workbook.close()

    sheet = read_sheet(path, "Лист_1")
    layout = detect_warehouse_layout(sheet)
    assert layout == WarehouseLayout(
        3, 0, 4, 5, 8, 9, 13, (MonthColumn(2026, 10, 16), MonthColumn(2026, 11, 17)),
        SourcePeriod(date(2026, 10, 1), date(2026, 11, 15)),
    )
    assert layout.months_label == "Q:R"
    safe = import_warehouse_products(sheet, layout, SalesQuantity.UNKNOWN)
    assert [row.metrics for row in safe.products] == [(None, None, Decimal("2.5"), 12), (None, None, 0, 0),
                                                       (None, None, 4, None), (None, None, 3, None)]
    assert safe.skipped_totals == safe.skipped_headers == 1
    assert "неоднозначной ценой — 2" in safe.summary
    assert safe.products[0].warehouse == "Тестовый склад"
    assert safe.products[0].unit == "шт."
    months = import_warehouse_products(sheet, layout, SalesQuantity.MONTHS)
    assert [p.sold_quantity for p in months.products] == [9, 4, 2, 1]
    outflow = import_warehouse_products(sheet, layout, SalesQuantity.OUTFLOW)
    assert [p.sold_quantity for p in outflow.products] == [7, 4, 2, 1]
    assert all(p.sold_amount is None for p in months.products + outflow.products)

    rows = [ReportRow(p) for p in months.products]
    catalog = builtin_catalog()
    refresh_groups(rows, catalog)
    assert [row.group for row in rows] == [53, 100, None, 40]
    rows[1].assign(None)
    refresh_groups(rows, catalog)
    target = tmp_path / "Результат.xlsx"
    export_xlsx(target, rows, catalog, months.column_notes, period=months.period)
    workbook = load_workbook(target)
    try:
        sheet = workbook.active
        assert sheet.max_column == 6 and sheet.max_row == 5
        assert tuple(c.value for c in sheet[1]) == RESULT_COLUMNS
        assert sheet["A2"].value == catalog.label(53)
        assert sheet["A3"].value is None
        assert (sheet["C2"].value, sheet["D2"].value, sheet["E2"].value, sheet["F2"].value) == (9, None, 2.5, 12)
        assert "Q:R" in sheet["C1"].comment.text and "не заменяет" in sheet["D1"].comment.text
    finally:
        workbook.close()


def test_blank_months_are_not_silently_converted_to_zero():
    data = Sheet((_HEADERS, stock_row("Товар", 8, 1, 2, (None, 4))))
    layout = detect_warehouse_layout(data)
    assert import_warehouse_products(data, layout, SalesQuantity.MONTHS).products[0].sold_quantity is None


def test_requires_unambiguous_columns_and_reports_bad_qty():
    duplicates = (*_HEADERS, _HEADERS[-1])
    layout = detect_warehouse_layout(Sheet((duplicates,)))
    assert layout is not None and not layout.months
    with pytest.raises(StatementError, match="нет однозначных помесячных"):
        import_warehouse_products(Sheet((duplicates,)), layout, SalesQuantity.MONTHS)
    data = Sheet((_HEADERS, stock_row("Товар", 8, "сломано", 2, (3, 4))))
    with pytest.raises(StatementError, match="Строка 2, столбец J"):
        import_warehouse_products(data, detect_warehouse_layout(data), SalesQuantity.UNKNOWN)
    assert detect_warehouse_layout(Sheet((("Товар", "Расход", "Конечный остаток", "Розничная цена"),))) is None


@pytest.mark.parametrize("include_section", [True, False])
def test_warehouse_section_is_not_used_for_detection_or_classification(include_section):
    rows = (
        _HEADERS,
        stock_row("Антисептик Белмастер Экобио 5л.", 0, 1, 2, (0, 0), "АРХИВ", "Мир"),
        stock_row("Антисептик Белмастер Экобио 5л.", 0, 2, 2, (0, 0), "Антисептики и пропитки", "Кореличи"),
        stock_row("Муфта 40 мм", 0, 2, 2, (0, 0), "Металлопласт", "Мир"),
        stock_row("Носки", 0, 3, 8, (0, 0), "Ликёры и строительные материалы", "Мир"),
    )
    data = Sheet(rows if include_section else tuple((row[0], *row[2:]) for row in rows))
    layout = detect_warehouse_layout(data)
    assert layout is not None
    pairs = [ReportRow(p) for p in import_warehouse_products(data, layout, SalesQuantity.UNKNOWN).products]
    refresh_groups(pairs, builtin_catalog())
    assert [(row.product.warehouse, row.group) for row in pairs] == [("Мир", 53), ("Кореличи", 53), ("Мир", None), ("Мир", 100)]


def test_selected_months_and_end_stock_date_cannot_be_confused():
    data = Sheet((
        ("С 12.10.26 по 15.11.26",), _HEADERS,
        stock_row("Плинтус", 26, 7, 3, (4, 6), unit="пог.м."),
        stock_row("Сверло", 13, 2, 4, (None, 0)),
    ))
    layout = detect_warehouse_layout(data)
    early = import_warehouse_products(data, layout, SalesQuantity.MONTHS, (0, 0))
    assert early.period.sales_months == ("Октябрь 2026",)
    assert early.period.sales_label == "Октябрь 2026 (с 12.10.2026)"
    assert "нет" in early.period.stock_label
    assert [p.metrics for p in early.products] == [(4, None, None, None), (None, None, None, None)]
    assert early.products[0].unit == "пог. м"
    last = import_warehouse_products(data, layout, SalesQuantity.MONTHS, (1, 1))
    assert last.period.sales_months == ("Ноябрь 2026",)
    assert last.period.sales_label == "Ноябрь 2026 (по 15.11.2026)"
    assert last.period.stock_label == "15.11.2026"
    assert [p.metrics for p in last.products] == [(6, None, 7, 21), (0, None, 2, 8)]
    with pytest.raises(StatementError, match="первый месяц"):
        import_warehouse_products(data, layout, SalesQuantity.MONTHS, (1, 0))
    with pytest.raises(StatementError, match="Расход.*за весь период"):
        import_warehouse_products(data, layout, SalesQuantity.OUTFLOW, (0, 0))
