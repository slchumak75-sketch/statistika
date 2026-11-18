"""Детализация, сводка по группам и Excel, без привязки к Tkinter."""

from __future__ import annotations

import tempfile
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from statistika.catalog import GROUPS, Catalog
from statistika.classifier import Classifier
from statistika.statement import Product, ReportPeriod, SOURCE_COLUMNS

RESULT_COLUMNS = (
    "№ пункта и описание", SOURCE_COLUMNS[0], "Продано за период,\nшт.*", "Продано за период,\nсумма, руб.",
    "Остаток на конец,\nшт.*", "Остаток на конец,\nсумма, руб.",
)
SUMMARY_COLUMNS = (
    RESULT_COLUMNS[0], "Товарных строк", "Продано, ед.", "Продано, руб.", "Остаток, ед.", "Остаток, руб.",
)
QUANTITY_NOTE = "* Количества в единицах ведомости: штуки, банки, упаковки и др."
SUMMARY_NOTE = "Количества сложены без пересчёта единиц. * — частичный итог: не во всех строках есть значение."
ALL_WAREHOUSES = "Все склады"


@dataclass
class ReportRow:
    product: Product
    group: int | None = None
    manual: bool = False

    def assign(self, group: int | None) -> None:
        if group is not None and group not in GROUPS:
            raise ValueError(f"Группа {group} недоступна для назначения.")
        self.group = group
        self.manual = True


def refresh_groups(rows: Iterable[ReportRow], catalog: Catalog) -> None:
    classifier = Classifier(catalog)
    # Название одинакового товара не получает разные группы из-за склада или исходного раздела.
    by_name: dict[str, int | None] = {}
    for row in rows:
        if not row.manual:
            name = row.product.name
            if name not in by_name:
                by_name[name] = classifier.group_for(name)
            row.group = by_name[name]


def warehouses(rows: Iterable[ReportRow]) -> tuple[str, ...]:
    return tuple(sorted({row.product.warehouse for row in rows if row.product.warehouse}, key=str.casefold))


def in_warehouse(row: ReportRow, warehouse: str | None) -> bool:
    return warehouse is None or row.product.warehouse == warehouse


@dataclass(frozen=True)
class Total:
    value: Decimal | None
    known: int
    lines: int

    @property
    def partial(self) -> bool:
        return 0 < self.known < self.lines


def _total(values: Iterable[Decimal | None]) -> Total:
    result = Decimal(0)
    known = lines = 0
    for value in values:
        lines += 1
        if value is not None:
            result += value
            known += 1
    return Total(result if known else None, known, lines)


@dataclass(frozen=True)
class GroupSummary:
    group: int | None
    lines: int
    sold_quantity: Total
    sold_amount: Total
    ending_quantity: Total
    ending_amount: Total

    @property
    def metrics(self) -> tuple[Total, ...]:
        return self.sold_quantity, self.sold_amount, self.ending_quantity, self.ending_amount


def summarize_groups(rows: Iterable[ReportRow]) -> tuple[GroupSummary, ...]:
    groups: dict[int | None, list[Product]] = defaultdict(list)
    for row in rows:
        groups[row.group].append(row.product)
    summaries = []
    for group in (*GROUPS, None):
        products = groups[group]
        # Количества складываем в единицах ведомости, не делим по единицам и не пересчитываем их.
        summaries.append(GroupSummary(
            group, len(products), _total(p.sold_quantity for p in products), _total(p.sold_amount for p in products),
            _total(p.ending_quantity for p in products), _total(p.ending_amount for p in products),
        ))
    return tuple(summaries)


def display_number(value: Decimal | None) -> str:
    return "" if value is None else format(value, "f").replace(".", ",")


def display_quantity(value: Decimal | None, unit: str) -> str:
    number = display_number(value)
    return f"{number} {unit}" if number and unit and unit != "шт." else number


def display_total(total: Total) -> str:
    return display_number(total.value) + ("*" if total.partial else "")


def _comment(sheet: Worksheet, row: int, col: int, note: str) -> None:
    comment = Comment(note, "Statistika")
    comment.width = 385
    comment.height = min(290, 65 + len(note) // 55 * 15)
    sheet.cell(row, col).comment = comment


def _header_style(sheet: Worksheet, row: int, widths: tuple[int, ...]) -> None:
    for col, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(col)].width = width
        cell = sheet.cell(row, col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="214060")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[row].height = 38


def _number_format(value: Decimal, unit: str = "", partial: bool = False) -> str:
    suffix = unit.replace('"', "").replace("\\", "")[:30]
    suffix += "*" if partial else ""
    if not suffix:
        return "General"
    base = "0" if value == value.to_integral_value() else "0.###############"
    return f'{base}" {suffix}"'


def _write_summary(
    workbook: Workbook, rows: list[ReportRow], catalog: Catalog, period: ReportPeriod, warehouse: str,
) -> None:
    sheet = workbook.create_sheet("Сводка по группам")
    context = (
        ("Склад", warehouse),
        ("Реализация", period.sales_label or "период в источнике не указан"),
        ("Остатки", period.stock_label or "дата в источнике не указана"),
    )
    for index, (title, value) in enumerate(context, start=1):
        sheet.cell(index, 1, title).font = Font(bold=True)
        sheet.merge_cells(start_row=index, start_column=2, end_row=index, end_column=6)
        cell = sheet.cell(index, 2, value)
        cell.data_type = "s"
        cell.alignment = Alignment(wrap_text=True)
        if len(value) > 90:
            sheet.row_dimensions[index].height = 32
    if period.sales_months:
        _comment(sheet, 2, 2, "Месяцы реализации: " + ", ".join(period.sales_months) + ".")
    sheet.merge_cells("A4:F4")
    sheet["A4"] = SUMMARY_NOTE
    sheet["A4"].alignment = Alignment(wrap_text=True)
    sheet.row_dimensions[4].height = 28
    sheet.append(SUMMARY_COLUMNS)
    _header_style(sheet, 5, (44, 16, 26, 22, 26, 24))
    for index, totals in enumerate(summarize_groups(rows), start=6):
        label = catalog.short_label(totals.group) or "Без группы"
        sheet.append((label, totals.lines, *(total.value for total in totals.metrics)))
        sheet.cell(index, 1).data_type = "s"
        sheet.cell(index, 1).alignment = Alignment(wrap_text=True, vertical="top")
        if totals.group is not None and label != catalog.label(totals.group):
            _comment(sheet, index, 1, catalog.label(totals.group))
        for col, total in enumerate(totals.metrics, start=3):
            if total.value is not None and total.partial:
                sheet.cell(index, col).number_format = _number_format(total.value, partial=True)
            if total.known < total.lines:
                _comment(sheet, index, col, f"Показатель есть для {total.known} из {total.lines} товарных строк.")
        sheet.row_dimensions[index].height = 26
    sheet.freeze_panes = "C6"
    sheet.auto_filter.ref = f"A5:F{sheet.max_row}"


def export_xlsx(
    path: Path, rows: Iterable[ReportRow], catalog: Catalog, column_notes: Iterable[tuple[int, str]] = (), *,
    period: ReportPeriod = ReportPeriod(), warehouse: str | None = None,
) -> None:
    rows = [row for row in rows if in_warehouse(row, warehouse)]
    warehouse_label = warehouse if warehouse is not None else ALL_WAREHOUSES
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Продажи и остатки"
    sheet.append(RESULT_COLUMNS)
    _header_style(sheet, 1, (64, 52, 24, 22, 25, 24))
    _comment(sheet, 1, 2, f"Склад: {warehouse_label}. Период реализации: {period.sales_label or 'не указан'}.")
    notes = dict(column_notes)
    for col in (3, 4, 5, 6):
        parts = [notes.get(col, "")]
        if col in (3, 5):
            parts.append(QUANTITY_NOTE)
        if col == 3 and period.sales_months and "выбранные месяцы" not in notes.get(col, ""):
            parts.append("Месяцы: " + ", ".join(period.sales_months) + ".")
        if note := "\n".join(part for part in parts if part):
            _comment(sheet, 1, col, note)
    for index, row in enumerate(rows, start=2):
        sheet.append((catalog.label(row.group) or None, row.product.name, *row.product.metrics))
        for col in (1, 2):
            cell = sheet.cell(index, col)
            # Строки источника и справочника не интерпретируются как формулы.
            if cell.value is not None:
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if row.product.unit and row.product.unit != "шт.":
            for col, value in ((3, row.product.sold_quantity), (5, row.product.ending_quantity)):
                if value is not None:
                    sheet.cell(index, col).number_format = _number_format(value, row.product.unit)
    sheet.freeze_panes = "C2"
    sheet.auto_filter.ref = sheet.dimensions
    _write_summary(workbook, rows, catalog, period, warehouse_label)
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as stream:
        temporary = Path(stream.name)
    try:
        workbook.save(temporary)
        temporary.replace(path)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)
