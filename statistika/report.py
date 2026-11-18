"""Рабочая таблица и её экспорт без привязки к графическому интерфейсу."""

from __future__ import annotations

import tempfile
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.cell import get_column_letter

from statistika.catalog import GROUPS, Catalog
from statistika.classifier import Classifier
from statistika.statement import SOURCE_COLUMNS, Product

RESULT_COLUMNS = ("№ пункта и описание", *SOURCE_COLUMNS)


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
    for row in rows:
        if not row.manual:
            row.group = classifier.group_for(row.product.name, row.product.nomenclature_group)


def display_number(value: Decimal | None) -> str:
    return "" if value is None else format(value, "f").replace(".", ",")


def export_xlsx(
    path: Path, rows: Iterable[ReportRow], catalog: Catalog, column_notes: Iterable[tuple[int, str]] = (),
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Продажи и остатки"
    sheet.append(RESULT_COLUMNS)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="214060")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for col, note in column_notes:
        comment = Comment(note, "Statistika")
        comment.width, comment.height = 350, 90
        sheet.cell(1, col).comment = comment
    sheet.row_dimensions[1].height = 34
    for col, width in enumerate((64, 52, 22, 22, 24, 24), start=1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    for index, row in enumerate(rows, start=2):
        sheet.append((catalog.label(row.group) or None, row.product.name, *row.product.metrics))
        for col in (1, 2):
            cell = sheet.cell(index, col)
            # Имя товара и описание — всегда текст, даже если начинаются с «=».
            if cell.value is not None:
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.freeze_panes = "C2"
    sheet.auto_filter.ref = sheet.dimensions
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as stream:
        temporary = Path(stream.name)
    try:
        workbook.save(temporary)
        temporary.replace(path)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)
