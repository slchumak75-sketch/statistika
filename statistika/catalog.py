"""Исходные пункты PDF отделены от списка групп, доступных для назначения."""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from pypdf import PdfReader
from pypdf.errors import PyPdfError

PDF_GROUPS = (35, 36, 37, 38, 39, 40, 41, 42, 43, 45, 46, 53, 57, 63)
HOSIERY = 100
GROUPS = (*PDF_GROUPS, HOSIERY)
HOSIERY_DESCRIPTION = "Чулочно-носочные изделия"
DEFAULT_CATALOG = Path(__file__).parent / "data" / "catalog.json"


class CatalogError(ValueError):
    """Ошибка, которую можно показать пользователю."""


@dataclass(frozen=True)
class Catalog:
    points: dict[int, str] = field(default_factory=dict)
    rules: dict[int, tuple[str, ...]] = field(default_factory=dict)
    source: str = ""

    def description(self, number: int) -> str:
        if number == HOSIERY:
            return HOSIERY_DESCRIPTION
        return self.points.get(number, "описание не загружено")

    def label(self, number: int | None) -> str:
        return "" if number is None else f"№{number} — {self.description(number)}"

    @property
    def missing(self) -> tuple[int, ...]:
        return tuple(number for number in PDF_GROUPS if number not in self.points)


# «35. Описание», «35) Описание», «№ 35 — Описание» и номера в
# первом столбце текстовой таблицы (с пробелами, полученными из PDF).
# Подпункты вроде «35.1» остаются частью текста предыдущего пункта.
_POINT = re.compile(
    r"^\s*(?:№\s*)?(?P<number>[1-9][0-9]{0,3})"
    r"(?:\s*[.)](?!\d)\s*|\s*[-–—]\s*|\s+(?=[^\W\d_]))(?P<text>\S.*)$"
)


def parse_points(text: str) -> dict[int, str]:
    chunks: dict[int, list[str]] = {}
    current: int | None = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.isdecimal():  # в том числе номера страниц
            continue
        match = _POINT.match(line)
        if match:
            current = int(match["number"])
            if current in chunks:
                raise CatalogError(f"Номер {current} повторяется. Оставьте одно описание для этого пункта.")
            chunks[current] = [match["text"]]
        elif current is not None:
            chunks[current].append(line)
    return {number: " ".join(" ".join(parts).split()) for number, parts in chunks.items()}


def format_points(points: dict[int, str]) -> str:
    return "\n".join(f"{number}. {text}" for number, text in sorted(points.items()))


def extract_pdf_text(path: Path) -> tuple[str, int]:
    """Возвращает извлечённый текст и число страниц без текста; OCR не выполняется."""
    try:
        with path.open("rb") as stream:
            reader = PdfReader(stream)
            if reader.is_encrypted:
                raise CatalogError("PDF защищён. Сохраните доступную копию без защиты и загрузите её.")
            pages = [page.extract_text(extraction_mode="layout") or "" for page in reader.pages]
    except (OSError, PyPdfError) as error:
        raise CatalogError(f"Не удалось прочитать PDF: {error}") from error
    if not any(page.strip() for page in pages):
        raise CatalogError("В PDF нет текстового слоя. Выполните OCR и повторите загрузку или введите пункты вручную.")
    return "\n".join(pages), sum(not page.strip() for page in pages)


def local_catalog_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "Statistika" / "catalog.json"


def read_catalog(path: Path) -> Catalog:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or data.get("version") != 1:
            raise ValueError("неподдерживаемая версия")
        points = {int(number): " ".join(text.split()) for number, text in data["points"].items()}
        rules = {}
        for number, phrases in data.get("rules", {}).items():
            if int(number) not in GROUPS or not isinstance(phrases, list):
                raise ValueError(f"недопустимое правило для группы {number}")
            rules[int(number)] = tuple(phrase.strip() for phrase in phrases if phrase.strip())
        if any(number < 1 or not text for number, text in points.items()):
            raise ValueError("непустое описание и положительный номер обязательны")
        source = data.get("source", "")
        if not isinstance(source, str):
            raise ValueError("некорректное имя источника")
    except (OSError, ValueError, TypeError, AttributeError, KeyError) as error:
        raise CatalogError(f"Не удалось открыть справочник «{path}»: {error}") from error
    return Catalog(points, rules, source)


def load_catalog(path: Path) -> Catalog:
    return read_catalog(path if path.exists() else DEFAULT_CATALOG)


def save_catalog(catalog: Catalog, path: Path) -> None:
    data = {
        "version": 1,
        "source": catalog.source,
        "points": catalog.points,
        "rules": catalog.rules,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(data, stream, ensure_ascii=False, indent=2, sort_keys=True)
            stream.write("\n")
        temporary.replace(path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
