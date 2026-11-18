"""Профиль выгрузки «Ведомость по остаткам ТМЦ на складах»."""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal
from enum import Enum

from openpyxl.utils.cell import get_column_letter

from statistika.statement import ImportedProducts, Product, Sheet, StatementError, cell_at, parse_number


class SalesQuantity(str, Enum):
    UNKNOWN = "unknown"
    MONTHS = "months"
    OUTFLOW = "outflow"


def _text(value: object) -> str:
    return "" if value is None else " ".join(str(value).split())


def _header(value: object) -> str:
    return _text(value).casefold().replace("ё", "е")


_MONTH = re.compile(
    r"^(?:январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь)\s+20\d\d(?:\s*г\.?)?$"
)
_PERIOD = re.compile(r"^с\s+\d{1,2}\.\d{1,2}\.\d{2,4}\s+по\s+\d{1,2}\.\d{1,2}\.\d{2,4}$", re.I)
_TOTAL = re.compile(r"^(?:итог[ои]?|общий итог|всего)\b", re.I)
_REQUIRED = ("склад", "группа номенклатуры", "номенклатура", "начальный остаток", "приход", "расход", "конечный остаток", "розничная цена")


@dataclass(frozen=True)
class WarehouseLayout:
    header_row: int  # с 1
    group_column: int  # индексы столбцов с 0
    name_column: int
    outflow_column: int
    ending_column: int
    price_column: int
    month_columns: tuple[int, ...] = ()
    period: str = ""

    @property
    def months_label(self) -> str:
        if not self.month_columns:
            return ""
        columns = [get_column_letter(col + 1) for col in self.month_columns]
        if len(columns) == 1:
            return columns[0]
        if self.month_columns == tuple(range(self.month_columns[0], self.month_columns[-1] + 1)):
            return f"{columns[0]}:{columns[-1]}"
        return ", ".join(columns)


def detect_warehouse_layout(sheet: Sheet) -> WarehouseLayout | None:
    """Явный профиль по сигнатуре заголовков, независимо от букв столбцов."""
    period = ""
    for index, row in enumerate(sheet.rows[:80], start=1):
        texts = tuple(_header(value) for value in row)
        for text in texts:
            if _PERIOD.fullmatch(text):
                period = text.capitalize()
        if not all(texts.count(title) == 1 for title in _REQUIRED):
            continue
        month_columns = tuple(col for col, name in enumerate(texts) if _MONTH.fullmatch(name))
        # Повторяющиеся заголовки месяцев неоднозначны — не предлагать их суммирование.
        if len({texts[col].rstrip(".").removesuffix(" г") for col in month_columns}) != len(month_columns):
            month_columns = ()
        return WarehouseLayout(
            index, texts.index("группа номенклатуры"), texts.index("номенклатура"), texts.index("расход"),
            texts.index("конечный остаток"), texts.index("розничная цена"), month_columns, period,
        )
    return None


def _number(row: tuple[object, ...], index: int, row_number: int) -> Decimal | None:
    value = cell_at(row, index)
    try:
        return parse_number(value)
    except ValueError as error:
        raise StatementError(f"Строка {row_number}, столбец {get_column_letter(index + 1)}: не число — {value!r}.") from error


def _price(value: object) -> Decimal | None:
    # В отчёте встречается несколько цен в одной ячейке. Не выбираем одну наугад.
    if isinstance(value, str) and len([part for part in value.splitlines() if part.strip()]) > 1:
        return None
    try:
        number = parse_number(value)
    except ValueError:
        return None
    return number if number is not None and number >= 0 else None


def import_warehouse_products(sheet: Sheet, layout: WarehouseLayout, source: SalesQuantity) -> ImportedProducts:
    if source == SalesQuantity.MONTHS and not layout.month_columns:
        raise StatementError("В листе нет однозначных помесячных столбцов.")
    products: list[Product] = []
    totals = repeats = no_values = unknown_price = 0
    outflow = get_column_letter(layout.outflow_column + 1)
    stock = get_column_letter(layout.ending_column + 1)
    price = get_column_letter(layout.price_column + 1)
    for row_number, row in enumerate(sheet.rows[layout.header_row :], start=layout.header_row + 1):
        name = _text(cell_at(row, layout.name_column))
        if _TOTAL.match(name) or (not name and any(_TOTAL.match(_text(value)) for value in row)):
            totals += 1
            continue
        if (
            _header(name) == "номенклатура"
            and _header(cell_at(row, layout.ending_column)) == "конечный остаток"
        ):
            repeats += 1
            continue
        ending_qty = _number(row, layout.ending_column, row_number)
        sold_qty: Decimal | None = None
        if source == SalesQuantity.OUTFLOW:
            sold_qty = _number(row, layout.outflow_column, row_number)
        elif source == SalesQuantity.MONTHS:
            amounts = [_number(row, col, row_number) for col in layout.month_columns]
            # Не подменять пустой месячный показатель нулём, в т. ч. при частичных данных.
            if all(amount is not None for amount in amounts):
                sold_qty = sum((amount for amount in amounts if amount is not None), Decimal(0))
        if not name:
            if ending_qty is not None or sold_qty is not None:
                raise StatementError(f"Строка {row_number}: есть показатели, но нет наименования товара.")
            continue
        if ending_qty is None and sold_qty is None:
            no_values += 1
            continue
        stock_amount = None
        if ending_qty == 0:
            stock_amount = Decimal(0)
        elif ending_qty is not None:
            retail_price = _price(cell_at(row, layout.price_column))
            if retail_price is None:
                unknown_price += 1
            else:
                stock_amount = ending_qty * retail_price
        products.append(Product(name, sold_qty, None, ending_qty, stock_amount, _text(cell_at(row, layout.group_column))))
    if not products:
        raise StatementError("В ведомости ТМЦ нет товарных строк с показателями.")

    quantity_note = {
        SalesQuantity.UNKNOWN: "Не заполнено: ведомость ТМЦ не содержит поля, явно обозначенного как продажи.",
        SalesQuantity.MONTHS: f"Сумма помесячных количеств {layout.months_label}, по выбору пользователя. "
        "Количества считаются продажами только при соответствующей настройке выгрузки. "
        "Строки с пустыми месяцами не рассчитаны.",
        SalesQuantity.OUTFLOW: f"Показатель «Расход» ({outflow}), по выбору пользователя. "
        "Может включать движения, не являющиеся продажами.",
    }[source]
    sales_source = {
        SalesQuantity.UNKNOWN: "не заполнено",
        SalesQuantity.MONTHS: f"из месяцев {layout.months_label} (выбор пользователя)",
        SalesQuantity.OUTFLOW: f"из «Расхода» ({outflow}, выбор пользователя)",
    }[source]
    summary = (
        f"Количество продаж: {sales_source}. Суммы продаж в источнике нет. "
        f"Остаток в розничных ценах ({stock} × {price}): без цены или с неоднозначной ценой — {unknown_price} строк."
    )
    notes = (
        (3, quantity_note),
        (4, "Колонки фактической суммы продаж нет. Текущая розничная цена не заменяет цены отдельных продаж."),
        (5, f"Конечный остаток по ведомости, столбец {stock}." + (f" Период: {layout.period}." if layout.period else "")),
        (6, f"Конечное количество ({stock}) × розничная цена из ведомости ({price}). "
            "Если ненулевой остаток не имеет одной числовой цены, сумма не заполнена."),
    )
    return ImportedProducts(tuple(products), totals, repeats, no_values, summary, notes)
