"""Чтение табличных ведомостей и проверка пяти исходных столбцов."""

from __future__ import annotations

import codecs
import csv
import io
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from functools import cached_property
from itertools import islice
from pathlib import Path
from zipfile import BadZipFile

import xlrd
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

SOURCE_COLUMNS = (
    "Наименование товара",
    "Количество, проданное за период",
    "Сумма продаж за период",
    "Количество в остатке на конец периода",
    "Сумма остатка на конец периода",
)


class StatementError(ValueError):
    """Формат или данные ведомости требуют действий пользователя."""


@dataclass(frozen=True)
class Sheet:
    rows: tuple[tuple[object, ...], ...]

    @cached_property
    def width(self) -> int:
        return max(map(len, self.rows), default=0)


@dataclass(frozen=True)
class ImportPlan:
    first_header: int  # Номера строк с 1; (0, 0) = без заголовка.
    last_header: int
    columns: tuple[int | None, ...]  # Индексы колонок с 0, в порядке SOURCE_COLUMNS.


@dataclass(frozen=True)
class Product:
    name: str
    sold_quantity: Decimal | None
    sold_amount: Decimal | None
    ending_quantity: Decimal | None
    ending_amount: Decimal | None
    nomenclature_group: str = ""

    @property
    def metrics(self) -> tuple[Decimal | None, ...]:
        return (self.sold_quantity, self.sold_amount, self.ending_quantity, self.ending_amount)


@dataclass(frozen=True)
class ImportedProducts:
    products: tuple[Product, ...]
    skipped_totals: int = 0
    skipped_headers: int = 0
    skipped_no_values: int = 0
    summary: str = ""
    column_notes: tuple[tuple[int, str], ...] = ()  # Номера столбцов результата с 1.


def list_sheets(path: Path) -> tuple[str, ...]:
    try:
        if path.suffix.lower() in (".csv", ".tsv"):
            return ("Данные",)
        if path.suffix.lower() == ".xls":
            book = xlrd.open_workbook(str(path), on_demand=True)
            try:
                return tuple(book.sheet_names())
            finally:
                book.release_resources()
        if path.suffix.lower() in (".xlsx", ".xlsm"):
            book = load_workbook(path, read_only=True, data_only=True)
            try:
                return tuple(sheet.title for sheet in book.worksheets)
            finally:
                book.close()
    except (OSError, BadZipFile, InvalidFileException, xlrd.XLRDError, ValueError) as error:
        raise StatementError(f"Не удалось открыть ведомость: {error}") from error
    raise StatementError("Загрузите ведомость в формате XLSX, XLS, XLSM, CSV или TSV.")


def _csv_rows(path: Path):
    data = path.read_bytes()
    if data.startswith((codecs.BOM_UTF16_BE, codecs.BOM_UTF16_LE)):
        text = data.decode("utf-16")
    else:
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("cp1251")

    def reader(delimiter: str):
        return csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, strict=True)

    if path.suffix.lower() == ".tsv":
        return reader("\t")

    def score(delimiter: str) -> tuple[int, int, int]:
        try:
            widths = Counter(len(row) for row in islice(reader(delimiter), 100) if len(row) >= 2)
        except csv.Error:
            return (-1, 0, 0)
        return max(((width >= len(SOURCE_COLUMNS), count, width) for width, count in widths.items()), default=(0, 0, 0))

    delimiter = max((";", "\t", ","), key=score)
    return reader(delimiter)


def _xls_value(cell, datemode: int) -> object:
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode)
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(cell.value, "#ERROR")
    return cell.value


def _is_empty(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _trim_row(values) -> tuple[object, ...]:
    row = list(values)
    while row and _is_empty(row[-1]):
        row.pop()
    return tuple(row)


def read_sheet(path: Path, name: str) -> Sheet:
    try:
        suffix = path.suffix.lower()
        if suffix in (".csv", ".tsv"):
            rows = tuple(map(_trim_row, _csv_rows(path)))
        elif suffix == ".xls":
            book = xlrd.open_workbook(str(path), on_demand=True)
            try:
                sheet = book.sheet_by_name(name)
                rows = tuple(_trim_row(_xls_value(cell, book.datemode) for cell in row) for row in sheet.get_rows())
            finally:
                book.release_resources()
        elif suffix in (".xlsx", ".xlsm"):
            book = load_workbook(path, read_only=True, data_only=True)
            try:
                sheet = book[name]
                # Некоторые выгрузки указывают размер всего листа вместо фактических ячеек.
                sheet.reset_dimensions()
                rows = tuple(map(_trim_row, sheet.iter_rows(values_only=True)))
            finally:
                book.close()
        else:
            raise StatementError("Неподдерживаемый формат ведомости.")
    except (OSError, BadZipFile, InvalidFileException, xlrd.XLRDError, ValueError, KeyError, csv.Error) as error:
        raise StatementError(f"Не удалось прочитать лист «{name}»: {error}") from error
    end = len(rows)
    while end and not rows[end - 1]:
        end -= 1
    return Sheet(rows[:end])


def cell_at(row: tuple[object, ...], col: int) -> object:
    return row[col] if col < len(row) else None


def _heading(value: object) -> str:
    return " ".join(value.split()) if isinstance(value, str) else ""


def column_headers(sheet: Sheet, first: int, last: int) -> tuple[str, ...]:
    if first == last == 0:
        return ("",) * sheet.width
    if not (1 <= first <= last <= len(sheet.rows)):
        raise StatementError(f"Укажите строки заголовка от 1 до {len(sheet.rows)} или 0 и 0 (нет заголовка).")
    block = sheet.rows[first - 1 : last]
    by_column: list[list[str]] = [[] for _ in range(sheet.width)]
    for index, row in enumerate(block):
        parent = ""
        for col, parts in enumerate(by_column):
            text = _heading(cell_at(row, col))
            if text:
                parent = text
            elif index < len(block) - 1 and _metric(_words(parent)) is None:
                # Пустая часть объединённого заголовка получает контекст слева,
                # только если под ней действительно есть подпись подстолбца.
                if _period(_words(parent)) and any(_heading(cell_at(lower, col)) for lower in block[index + 1 :]):
                    text = parent
            if text and text not in parts:
                parts.append(text)
    return tuple(" / ".join(parts) for parts in by_column)


def _words(text: str) -> tuple[str, ...]:
    return tuple(re.findall(r"[\w]+", text.lower().replace("ё", "е").replace("_", " ")))


def _starts(words: tuple[str, ...], *stems: str) -> bool:
    return any(word.startswith(stems) for word in words)


def _metric(words: tuple[str, ...]) -> str | None:
    qty = _starts(words, "количеств") or any(word in ("шт", "кол", "колво", "кво") for word in words)
    qty = qty or " к во " in f" {' '.join(words)} "
    amount = _starts(words, "сумм", "стоимост", "выруч")
    return None if qty == amount else "quantity" if qty else "amount"


def _period(words: tuple[str, ...]) -> str | None:
    if _starts(words, "начальн", "начал", "входящ", "приход", "поступлен", "возврат", "списан", "резерв"):
        return "other"
    sold = _starts(words, "продаж", "продан", "реализац", "реализован", "выруч")
    ending = _starts(words, "исходящ") or (_starts(words, "остат") and _starts(words, "конец", "конеч", "конц"))
    if sold != ending:
        return "sold" if sold else "ending"
    return None


def _role(text: str) -> int | None:
    words = _words(text)
    metric = _metric(words)
    period = _period(words)
    if metric and period in ("sold", "ending"):
        return (1 if period == "sold" else 3) + (metric == "amount")
    if not metric and not period and not _starts(words, "код", "артикул", "групп", "единиц", "характеристик"):
        if words in (("товар",), ("товары",), ("номенклатура",), ("наименование",), ("название",)):
            return 0
        if _starts(words, "наименован", "названи") and _starts(words, "товар", "номенклатур", "продукц"):
            return 0
    return None


def guess_plan(sheet: Sheet) -> ImportPlan:
    best = ImportPlan(0, 0, (None,) * len(SOURCE_COLUMNS))
    best_score = (0, 0, 0)
    for end in range(1, min(80, len(sheet.rows)) + 1):
        lower = tuple(_heading(value) for value in sheet.rows[end - 1])
        signals = sum(_metric(_words(text)) is not None or _role(text) == 0 for text in lower)
        for depth in range(1, min(3, end) + 1):
            if depth > 1 and signals < 2:
                continue
            start = end - depth + 1
            candidates: list[list[int]] = [[] for _ in SOURCE_COLUMNS]
            for col, text in enumerate(column_headers(sheet, start, end)):
                role = _role(text)
                if role is not None:
                    candidates[role].append(col)
            columns = tuple(indices[0] if len(indices) == 1 else None for indices in candidates)
            count = sum(col is not None for col in columns)
            score = (count, -depth, -start)
            if score > best_score:
                best = ImportPlan(start, end, columns)
                best_score = score
    return best


_EMPTY_NUMBER = {"", "-", "—", "–"}
_TOTAL = re.compile(r"^(?:(?:общий\s+)?итог(?:о|и)?|всего)\b", re.I)


def parse_number(value: object) -> Decimal | None:
    """Оригинальные значения, без округления/подмены пропусков нулями."""
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError("логическое значение не является числом в ведомости")
    if isinstance(value, (Decimal, int, float)):
        result = Decimal(str(value))
    elif isinstance(value, str):
        text = unicodedata.normalize("NFKC", value).strip().replace("−", "-")
        if text in _EMPTY_NUMBER:
            return None
        if text.startswith("(") and text.endswith(")"):
            text = "-" + text[1:-1].strip()
        if re.search(r"\s", text):
            if not re.fullmatch(r"[+-]?\d{1,3}(?:\s+\d{3})+(?:[.,]\d+)?", text):
                raise ValueError("пробелы не разделяют разряды числа")
            text = "".join(text.split())
        if "," in text and "." in text:
            decimal, grouping = (",", ".") if text.rindex(",") > text.rindex(".") else (".", ",")
            if not re.fullmatch(rf"[+-]?\d{{1,3}}(?:{re.escape(grouping)}\d{{3}})+{re.escape(decimal)}\d+", text):
                raise ValueError("некорректные разделители числа")
            text = text.replace(grouping, "").replace(decimal, ".")
        else:
            sep = "," if "," in text else "."
            if text.count(sep) > 1:
                if not re.fullmatch(rf"[+-]?\d{{1,3}}(?:{re.escape(sep)}\d{{3}})+", text):
                    raise ValueError("некорректные разделители разрядов")
                text = text.replace(sep, "")
            else:
                text = text.replace(sep, ".")
        if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", text):
            raise ValueError("не число")
        try:
            result = Decimal(text)
        except InvalidOperation as error:
            raise ValueError("не число") from error
    else:
        raise ValueError("не число")
    if not result.is_finite():
        raise ValueError("ожидалось конечное число")
    return result


def import_products(sheet: Sheet, plan: ImportPlan) -> ImportedProducts:
    labels = column_headers(sheet, plan.first_header, plan.last_header)
    group_cols = [col for col, label in enumerate(labels) if _words(label) == ("группа", "номенклатуры")]
    if len(plan.columns) != len(SOURCE_COLUMNS) or any(
        col is None or not 0 <= col < sheet.width for col in plan.columns
    ):
        raise StatementError("Укажите все пять столбцов ведомости.")
    columns = tuple(col for col in plan.columns if col is not None)
    if len(set(columns)) != len(columns):
        raise StatementError("Для каждого поля нужен отдельный столбец ведомости.")

    def selected(row: tuple[object, ...]) -> tuple[object, ...]:
        return tuple(cell_at(row, col) for col in columns)

    def signature(row: tuple[object, ...]) -> tuple[str, ...]:
        return tuple(_heading(value).casefold() for value in row)

    headers = {
        signature(selected(row)) for row in sheet.rows[max(0, plan.first_header - 1) : plan.last_header]
    }
    products = []
    totals = repeated = no_values = 0
    for number, row in enumerate(sheet.rows[plan.last_header :], start=plan.last_header + 1):
        if not row:
            continue
        fields = selected(row)
        raw_name, *metrics = fields
        name = "" if _is_empty(raw_name) else " ".join(str(raw_name).split())
        if _TOTAL.match(name) or (not name and any(_TOTAL.match(_heading(value)) for value in row)):
            totals += 1
            continue
        if any(isinstance(value, str) and value.strip() for value in fields) and signature(fields) in headers:
            repeated += 1
            continue
        if all(map(_is_empty, metrics)):
            no_values += bool(name)
            continue
        if not name:
            if all(isinstance(value, str) and value.strip() in _EMPTY_NUMBER or _is_empty(value) for value in metrics):
                continue
            raise StatementError(f"Строка {number}: есть значения, но нет наименования. Проверьте столбец товара.")
        values: list[Decimal | None] = []
        for col, label, value in zip(columns[1:], SOURCE_COLUMNS[1:], metrics):
            try:
                values.append(parse_number(value))
            except ValueError as error:
                letter = get_column_letter(col + 1)
                raise StatementError(f"Строка {number}, столбец {letter} («{label}»): не число — {value!r}.") from error
        group = _heading(cell_at(row, group_cols[0])) if len(group_cols) == 1 else ""
        products.append(Product(name, *values, nomenclature_group=group))
    if not products:
        raise StatementError("Не найдено товарных строк. Проверьте лист, границы заголовка и выбранные столбцы.")
    return ImportedProducts(tuple(products), totals, repeated, no_values)
