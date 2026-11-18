"""Профиль выгрузки «Ведомость по остаткам ТМЦ на складах»."""

from __future__ import annotations

import re
from calendar import monthrange
from dataclasses import dataclass
from decimal import Decimal
from enum import Enum

from openpyxl.utils.cell import get_column_letter

from statistika.statement import (
    ImportedProducts, Product, ReportPeriod, Sheet, SourcePeriod, StatementError, cell_at, find_source_period,
    normalize_unit, parse_number,
)


class SalesQuantity(str, Enum):
    UNKNOWN = "unknown"
    MONTHS = "months"
    OUTFLOW = "outflow"


def _text(value: object) -> str:
    return "" if value is None else " ".join(str(value).split())


def _header(value: object) -> str:
    return _text(value).casefold().replace("ё", "е")


_MONTH_NAMES = (
    "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
)
_MONTH = re.compile(rf"^({'|'.join(_MONTH_NAMES)})\s+(20\d\d)(?:\s*г\.?)?$")
_TOTAL = re.compile(r"^(?:итог[ои]?|общий итог|всего)\b", re.I)
_REQUIRED = ("склад", "номенклатура", "начальный остаток", "приход", "расход", "конечный остаток", "розничная цена")


@dataclass(frozen=True)
class MonthColumn:
    year: int
    month: int
    column: int  # с 0

    @property
    def label(self) -> str:
        return f"{_MONTH_NAMES[self.month - 1].capitalize()} {self.year}"

    @property
    def key(self) -> tuple[int, int]:
        return self.year, self.month


def _column_span(months: tuple[MonthColumn, ...]) -> str:
    indexes = tuple(month.column for month in months)
    letters = [get_column_letter(index + 1) for index in indexes]
    if len(letters) > 1 and indexes == tuple(range(indexes[0], indexes[-1] + 1)):
        return f"{letters[0]}:{letters[-1]}"
    return ", ".join(letters)


def _month_span(months: tuple[MonthColumn, ...], period: SourcePeriod | None) -> str:
    first, last = months[0], months[-1]
    total = (last.year - first.year) * 12 + last.month - first.month + 1
    if len(months) == 1:
        label = first.label
    elif total == len(months):
        label = f"{first.label} — {last.label} ({len(months)} мес.)"
    else:
        label = ", ".join(month.label for month in months)  # пропущенные месяцы не включены
    if period:
        if first.key == (period.start.year, period.start.month) and period.start.day > 1:
            label += f" (с {period.start:%d.%m.%Y})"
        if last.key == (period.end.year, period.end.month) and period.end.day < monthrange(last.year, last.month)[1]:
            label += f" (по {period.end:%d.%m.%Y})"
    return label


@dataclass(frozen=True)
class WarehouseLayout:
    header_row: int  # с 1
    warehouse_column: int  # индексы столбцов с 0
    name_column: int
    unit_column: int | None
    outflow_column: int
    ending_column: int
    price_column: int
    months: tuple[MonthColumn, ...] = ()
    source_period: SourcePeriod | None = None

    @property
    def months_label(self) -> str:
        return _column_span(self.months)


def detect_warehouse_layout(sheet: Sheet) -> WarehouseLayout | None:
    """Явный профиль по сигнатуре заголовков, независимо от букв столбцов."""
    for index, row in enumerate(sheet.rows[:80], start=1):
        texts = tuple(_header(value) for value in row)
        if not all(texts.count(title) == 1 for title in _REQUIRED):
            continue
        source_period = find_source_period(sheet.rows[: index - 1])
        months = []
        for col, title in enumerate(texts):
            match = _MONTH.fullmatch(title)
            if match:
                month = MonthColumn(int(match[2]), _MONTH_NAMES.index(match[1]) + 1, col)
                if source_period and not (
                    (source_period.start.year, source_period.start.month) <= month.key
                    <= (source_period.end.year, source_period.end.month)
                ):
                    continue
                months.append(month)
        # Повторяющиеся заголовки месяцев неоднозначны — не предлагать их суммирование.
        if len({month.key for month in months}) != len(months):
            months = []
        units = [col for col, title in enumerate(texts) if title in {"ед.", "ед", "ед. изм.", "единица измерения"}]
        return WarehouseLayout(
            index, texts.index("склад"), texts.index("номенклатура"),
            units[0] if len(units) == 1 else None, texts.index("расход"), texts.index("конечный остаток"),
            texts.index("розничная цена"), tuple(sorted(months, key=lambda month: month.key)), source_period,
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


def _selection(layout: WarehouseLayout, source: SalesQuantity, month_range: tuple[int, int] | None) -> tuple[MonthColumn, ...]:
    if source != SalesQuantity.MONTHS:
        if month_range is not None:
            raise StatementError("Выбирать месяцы можно только для помесячной реализации. «Расход» дан за весь период ведомости.")
        return ()
    if not layout.months:
        raise StatementError("В листе нет однозначных помесячных столбцов.")
    start, end = month_range if month_range is not None else (0, len(layout.months) - 1)
    if not 0 <= start <= end < len(layout.months):
        raise StatementError("Выберите начало и конец реализации: первый месяц должен быть не позже последнего.")
    return layout.months[start : end + 1]


def import_warehouse_products(
    sheet: Sheet, layout: WarehouseLayout, source: SalesQuantity, month_range: tuple[int, int] | None = None,
) -> ImportedProducts:
    months = _selection(layout, source, month_range)
    source_period = layout.source_period
    snapshot = source_period.end if source_period else None
    snapshot_label = f"{snapshot:%d.%m.%Y}" if snapshot else "дата в источнике не указана"
    # J — один снимок, не история запасов. При раннем конце периода не подставляем поздние остатки.
    use_stock = not months or (months[-1].key == (snapshot.year, snapshot.month) if snapshot else months == layout.months)
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
        if _header(name) == "номенклатура" and _header(cell_at(row, layout.ending_column)) == "конечный остаток":
            repeats += 1
            continue
        ending_qty = _number(row, layout.ending_column, row_number) if use_stock else None
        sold_qty: Decimal | None = None
        if source == SalesQuantity.OUTFLOW:
            sold_qty = _number(row, layout.outflow_column, row_number)
        elif months:
            amounts = [_number(row, month.column, row_number) for month in months]
            # Не подменять пустой месячный показатель нулём, в т. ч. при частичных данных.
            if all(amount is not None for amount in amounts):
                sold_qty = sum((amount for amount in amounts if amount is not None), Decimal(0))
        if not name:
            if ending_qty is not None or sold_qty is not None:
                raise StatementError(f"Строка {row_number}: есть показатели, но нет наименования товара.")
            continue
        if ending_qty is None and sold_qty is None and (use_stock or not _text(cell_at(row, layout.ending_column))):
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
        products.append(Product(
            name, sold_qty, None, ending_qty, stock_amount,
            warehouse=_text(cell_at(row, layout.warehouse_column)),
            unit=normalize_unit(cell_at(row, layout.unit_column)) if layout.unit_column is not None else "",
        ))
    if not products:
        raise StatementError("В ведомости ТМЦ нет товарных строк с показателями.")

    if source == SalesQuantity.UNKNOWN:
        quantity_note = "Не заполнено: в ведомости нет поля, явно обозначенного как продажи, и источник не выбран."
        sales_label = "источник не выбран"
    elif source == SalesQuantity.MONTHS:
        quantity_note = (
            f"Сумма количеств из {_column_span(months)} за выбранные месяцы: {', '.join(month.label for month in months)}. "
            "Используется по выбору пользователя; строки с пустыми месяцами не рассчитаны."
        )
        sales_label = _month_span(months, source_period)
    else:
        quantity_note = f"«Расход» ({outflow}) за весь период ведомости. Выбран пользователем; может включать не только продажи."
        sales_label = f"{source_period.label} («Расход»)" if source_period else "весь период ведомости («Расход»)"

    if use_stock:
        stock_label = snapshot_label
        stock_quantity_note = f"Конечный остаток ({stock}): {snapshot_label}."
        stock_amount_note = f"Конечное количество ({stock}) × розничная цена ({price}). Без одной числовой цены сумма не рассчитана."
        summary = f"Суммы продаж в источнике нет. Остаток ({stock} × {price}): без цены или с неоднозначной ценой — {unknown_price} строк."
    else:
        stock_label = "нет на конец выбранного периода"
        stock_quantity_note = stock_amount_note = f"Не рассчитано: ведомость содержит только снимок остатков ({snapshot_label})."
        summary = f"Суммы продаж в источнике нет. Остатки не подставлены: ведомость содержит только снимок ({snapshot_label})."
    return ImportedProducts(
        tuple(products), totals, repeats, no_values, summary,
        ((3, quantity_note),
         (4, "Колонки фактической суммы продаж нет. Текущая розничная цена не заменяет цены отдельных продаж."),
         (5, stock_quantity_note), (6, stock_amount_note)),
        ReportPeriod(sales_label, tuple(month.label for month in months), stock_label),
    )
