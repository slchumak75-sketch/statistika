from decimal import Decimal

import pytest
import xlwt
from openpyxl import Workbook

from statistika.statement import (
    SOURCE_COLUMNS, ImportPlan, Sheet, StatementError, column_headers, guess_plan, import_products,
    list_sheets, parse_number, read_sheet,
)


@pytest.mark.parametrize(("value", "expected"), [
    ("1\u202f234,50", Decimal("1234.50")), (" 1\u00a0234.567 ", Decimal("1234.567")),
    ("1.234,56", Decimal("1234.56")), ("1,234.56", Decimal("1234.56")),
    ("1,234,567", Decimal("1234567")), ("1.234", Decimal("1.234")),
    ("(4,50)", Decimal("-4.50")), ("−2", Decimal("-2")), ("0", Decimal("0")),
    (12.3, Decimal("12.3")), (None, None), ("–", None), ("  ", None),
])
def test_parse_number(value, expected):
    assert parse_number(value) == expected


@pytest.mark.parametrize("value", [True, float("inf"), "#VALUE!", "12 руб.", "1 23", "13,2.9"])
def test_invalid_number(value):
    with pytest.raises(ValueError):
        parse_number(value)


def test_xlsx_multiline_headers_and_totals(tmp_path):
    path = tmp_path / "Ведомость.xlsx"
    workbook = Workbook()
    workbook.active.title = "Титульный лист"
    sheet = workbook.create_sheet("Товары")
    sheet.append(["Ведомость по товарам за март"])
    major = ("Номенклатура", "Начальный остаток", None, "Приход", None, "Реализация", None, "Конечный остаток")
    minor = (None, "Кол-во", "Сумма", "Кол-во", "Сумма", "Кол-во", "Сумма", "Кол-во", "Сумма")
    sheet.append(major)
    for col in (2, 4, 6, 8):
        sheet.merge_cells(start_row=2, end_row=2, start_column=col, end_column=col + 1)
    sheet.append(minor)
    sheet.append(["Носки", 100, 500, 5, 25, 2, 12.5, 103, 618.7])
    sheet.append(major)  # повтор при печатной выгрузке
    sheet.append(minor)
    sheet.append(["Раздел 2"])
    sheet.append(["Прочий товар", None, None, 8, 33, 0, "—", 1, 5.99])
    sheet.append(["Итого:", 100, 500, 13, 58, 2, 12.5, 104, 624.69])
    workbook.save(path)
    workbook.close()

    assert list_sheets(path) == ("Титульный лист", "Товары")
    data = read_sheet(path, "Товары")
    plan = guess_plan(data)
    assert plan == ImportPlan(2, 3, (0, 5, 6, 7, 8))
    assert column_headers(data, 2, 3)[6] == "Реализация / Сумма"
    assert column_headers(data, 2, 3)[8] == "Конечный остаток / Сумма"
    result = import_products(data, plan)
    assert [product.name for product in result.products] == ["Носки", "Прочий товар"]
    assert result.products[0].metrics == (2, Decimal("12.5"), 103, Decimal("618.7"))
    assert result.products[1].metrics == (0, None, 1, Decimal("5.99"))
    assert (result.skipped_totals, result.skipped_headers, result.skipped_no_values) == (1, 2, 1)


def test_does_not_confuse_shipments_opening_or_duplicates_with_sales(tmp_path):
    sheet = Sheet((
        ("Наименование", "Количество расхода", "Сумма расхода", "Начальный остаток количество",
         "Конечный остаток количество", "Сумма остатка на конец", "Сумма остатка на конец (по учёту)"),
        ("Товар X", 3, 15, 28, 5, 35, 30),
    ))
    assert guess_plan(sheet) == ImportPlan(1, 1, (0, None, None, 4, None))
    # Здесь оператор знает, что «Расход» в этой выгрузке — именно продажи.
    manual = ImportPlan(1, 1, (0, 1, 2, 4, 6))
    assert import_products(sheet, manual).products[0].metrics == (3, 15, 5, 30)


def test_xls(tmp_path):
    path = tmp_path / "ведомость.xls"
    book = xlwt.Workbook()
    sheet = book.add_sheet("Товары")
    for row, items in enumerate((SOURCE_COLUMNS, ("Гольфы", 1, 45.5, 7, 318.5))):
        for col, value in enumerate(items):
            sheet.write(row, col, value)
    book.save(str(path))
    data = read_sheet(path, list_sheets(path)[0])
    assert import_products(data, guess_plan(data)).products[0].metrics == (1, Decimal("45.5"), 7, Decimal("318.5"))


@pytest.mark.parametrize(("filename", "delimiter", "encoding"), [
    ("ведомость.csv", ";", "cp1251"), ("ведомость.CSV", ",", "utf-8-sig"), ("ведомость.tsv", "\t", "utf-16"),
])
def test_delimited_input(tmp_path, filename, delimiter, encoding):
    import csv

    path = tmp_path / filename
    with path.open("w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        writer.writerow(("Период отчёта",))
        writer.writerow(SOURCE_COLUMNS)
        writer.writerow(("Платок, цвет синий", "2", "123,55", "0", "0"))
    data = read_sheet(path, list_sheets(path)[0])
    result = import_products(data, guess_plan(data))
    assert result.products[0].name == "Платок, цвет синий"
    assert result.products[0].metrics == (2, Decimal("123.55"), 0, 0)


def test_mapping_rejects_bad_amount_and_missing_name():
    plan = ImportPlan(0, 0, (0, 1, 2, 3, 4))
    with pytest.raises(StatementError, match=r"Строка 1, столбец C.*58 руб"):
        import_products(Sheet((("A", 1, "58 руб", 2, 100),)), plan)
    with pytest.raises(StatementError, match="нет наименования"):
        import_products(Sheet(((None, 1, 58, 2, 100),)), plan)
    with pytest.raises(StatementError, match="отдельный столбец"):
        import_products(Sheet((("A", 1, 58, 2, 100),)), ImportPlan(0, 0, (0, 1, 2, 1, 4)))
    with pytest.raises(StatementError, match="строки заголовка"):
        import_products(Sheet((("A", 1, 58, 2, 100),)), ImportPlan(0, 1, plan.columns))
