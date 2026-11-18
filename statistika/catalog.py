"""Исходные пункты PDF отделены от списка групп, доступных для назначения."""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from dataclasses import asdict, dataclass, field, replace
from functools import lru_cache
from pathlib import Path

from pypdf import PdfReader
from pypdf.errors import PyPdfError

PDF_GROUPS = (35, 36, 37, 38, 39, 40, 41, 42, 43, 45, 46, 53, 57, 63)
HOSIERY = 100
GROUPS = (*PDF_GROUPS, HOSIERY)
HOSIERY_DESCRIPTION = "Чулочно-носочные изделия"
DEFAULT_CATALOG = Path(__file__).parent / "data" / "catalog.json"
# v2 записывал и нетронутые встроенные списки. Обновляем только точную копию прежнего №53.
_ORIGINAL_53_V2 = "0b8fdf9521a8d6ad3715368a9a494923da6cb21f06ad71b23cb9db0ca0ea470a"
GROUP_NAMES_4_TORG = {
    35: "Вина виноградные",
    36: "Вина плодовые и фруктово-ягодные",
    37: "Вина игристые и Советское шампанское",
    38: "Водка",
    39: "Коньяки, бренди, арманьяк, дивины",
    40: "Пиво",
    41: "Ликёры и ликёро-водочные изделия",
    42: "Слабоалкогольные напитки",
    43: "Прочие алкогольные напитки",
    45: "Минеральная и питьевая вода",
    46: "Табачные изделия",
    53: "Строительные материалы",
    57: "Бытовая химия и автокосметика",
    63: "Парфюмерия, косметика и гигиена",
}


class CatalogError(ValueError):
    """Ошибка, которую можно показать пользователю."""


def clean_phrase(text: str) -> str:
    """Убрать символы разметки перед фразой, не меняя обозначения вроде SDS+ и 0%."""
    return " ".join(text.split()).lstrip(" ^!#•✓✔☑").strip()


@dataclass(frozen=True)
class RecognitionRule:
    phrase: str
    at_start: bool = False
    exclude: bool = False
    also: tuple[str, ...] = ()
    also_at_start: tuple[str, ...] = ()
    internal_only: bool = False


def _legacy_rule(line: str) -> RecognitionRule | None:
    """Перенос условий v1/v2 из текстовой записи в отдельные поля."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    excluded = line.startswith("!")
    segments = re.split(r"\s+\+\s+", line[1:].strip() if excluded else line)
    parts = [(clean_phrase(term), term.lstrip().startswith("^")) for term in segments]
    if not all(phrase and any(ch.isalnum() for ch in phrase) for phrase, _ in parts):
        return None
    (phrase, at_start), *extra = parts
    return RecognitionRule(phrase, at_start, excluded, tuple(p for p, start in extra if not start),
                           tuple(p for p, start in extra if start))


def _read_rule(value: object) -> RecognitionRule:
    if not isinstance(value, dict) or not isinstance(value.get("phrase"), str):
        raise ValueError("некорректная фраза распознавания")
    phrase = clean_phrase(value["phrase"])
    at_start, exclude, internal_only = (value.get(key, False) for key in ("at_start", "exclude", "internal_only"))
    if not any(ch.isalnum() for ch in phrase) or not all(isinstance(flag, bool) for flag in (at_start, exclude, internal_only)):
        raise ValueError("некорректные поля правила")

    def extras(key: str) -> tuple[str, ...]:
        phrases = value.get(key, [])
        if not isinstance(phrases, list) or not all(isinstance(p, str) and any(ch.isalnum() for ch in p) for p in phrases):
            raise ValueError("некорректные уточнения правила")
        return tuple(clean_phrase(p) for p in phrases)

    return RecognitionRule(phrase, at_start, exclude, extras("also"), extras("also_at_start"), internal_only)


def _write_rule(rule: RecognitionRule) -> dict[str, object]:
    return {key: value for key, value in asdict(rule).items() if value}


@dataclass(frozen=True)
class Catalog:
    points: dict[int, str] = field(default_factory=dict)
    rules: dict[int, tuple[RecognitionRule, ...]] = field(default_factory=dict)
    source: str = ""

    def description(self, number: int) -> str:
        if number == HOSIERY:
            return HOSIERY_DESCRIPTION
        return self.points.get(number, "описание не загружено")

    def label(self, number: int | None) -> str:
        return "" if number is None else f"№{number} — {self.description(number)}"

    def short_label(self, number: int | None) -> str:
        if number in GROUP_NAMES_4_TORG and self.is_official_group(number):
            return f"№{number} — {GROUP_NAMES_4_TORG[number]}"
        return self.label(number)

    def is_official_group(self, number: int) -> bool:
        return number in self.points and self.points[number] == builtin_catalog().points.get(number)

    def default_rules(self, number: int) -> tuple[RecognitionRule, ...]:
        if number == HOSIERY or self.is_official_group(number):
            return builtin_catalog().rules.get(number, ())
        description = self.points.get(number, "")
        # Для другого PDF предлагаем только явные термины без оговорок.
        if _LIMITS.search(description) or any(char in description for char in "(["):
            return ()
        phrases = (part.strip(" .:-—–") for part in re.split(r"[,;]", description))
        return tuple(RecognitionRule(phrase) for phrase in phrases if len(phrase) >= 3 and phrase.casefold() not in _GENERIC)

    def recognition_rules(self, number: int) -> tuple[RecognitionRule, ...]:
        # Сохранённый пустой список отключает распознавание, а не включает встроенное заново.
        return self.rules[number] if number in self.rules else self.default_rules(number)

    @property
    def missing(self) -> tuple[int, ...]:
        return tuple(number for number in PDF_GROUPS if number not in self.points)


# «35. Описание», «35) Описание», «№ 35 — Описание» и номера в
# первом столбце текстовой таблицы. Не считать «4-торг» и «35.1» пунктами.
_POINT = re.compile(
    r"^\s*(?:№\s*)?(?P<number>[1-9][0-9]{0,3})"
    r"(?:\s*[.)](?!\d)\s*|\s+[-–—]\s*|[-–—]\s+|\s+(?=[^\W\d_]))(?P<text>\S.*)$"
)
_DOTTED_POINT = re.compile(r"^(?P<number>[1-9][0-9]{0,3})\.\s+(?P<text>\S.*)$")
_CHAPTER = re.compile(r"^глава\s+\d+$", re.I)
_LIMITS = re.compile(r"\b(?:кроме|за исключением|исключая|не включая|не учт[её]нн\w*|только)\b", re.I)
_GENERIC = frozenset(("товар", "товары", "изделие", "изделия", "продукция", "прочие", "другие", "наборы", "комплекты", "разные", "в том числе"))


def parse_points(text: str) -> dict[int, str]:
    signature = " ".join(text.casefold().split())
    instructions = "глава 1" in signature and "по заполнению формы" in signature and "4-торг" in signature
    point_pattern = _DOTTED_POINT if instructions else _POINT
    chunks: dict[int, list[str]] = {}
    current: int | None = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.isdecimal():  # в том числе номера страниц
            continue
        if instructions and _CHAPTER.fullmatch(line):
            current = None  # Не переносить название следующей главы в предыдущий пункт.
            continue
        match = point_pattern.match(line)
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


def _read_catalog(path: Path) -> tuple[Catalog, int]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or data.get("version") not in (1, 2, 3):
            raise ValueError("неподдерживаемая версия")
        points = {int(number): " ".join(text.split()) for number, text in data["points"].items()}
        rules = {}
        for number, phrases in data.get("rules", {}).items():
            if int(number) not in GROUPS or not isinstance(phrases, list):
                raise ValueError(f"недопустимое правило для группы {number}")
            if data["version"] < 3:
                entries = tuple(rule for phrase in phrases if (rule := _legacy_rule(phrase)) is not None)
            else:
                entries = tuple(_read_rule(rule) for rule in phrases)
            if int(number) != 53 and any(rule.internal_only for rule in entries):
                raise ValueError("режим внутренней стройгруппы доступен только для №53")
            rules[int(number)] = entries
        if any(number < 1 or not text for number, text in points.items()):
            raise ValueError("непустое описание и положительный номер обязательны")
        source = data.get("source", "")
        if not isinstance(source, str):
            raise ValueError("некорректное имя источника")
    except (OSError, ValueError, TypeError, AttributeError, KeyError) as error:
        raise CatalogError(f"Не удалось открыть справочник «{path}»: {error}") from error
    return Catalog(points, rules, source), data["version"]


@lru_cache(maxsize=1)
def builtin_catalog() -> Catalog:
    return _read_catalog(DEFAULT_CATALOG)[0]


def load_catalog(path: Path) -> Catalog:
    default = builtin_catalog()
    if not path.exists():
        return replace(default, rules={})
    local, version = _read_catalog(path)
    if not local.points and not local.source:
        local = replace(local, points=default.points, source=default.source)
    if version == 1:
        # В v1 пользовательские фразы дополняли встроенные. С v2 сохраняется весь список.
        local = replace(local, rules={
            number: tuple(dict.fromkeys((*local.default_rules(number), *phrases)))
            for number, phrases in local.rules.items()
        })
    if version == 2 and local.is_official_group(53) and local.rules.get(53):
        entries = [_write_rule(rule) for rule in local.rules[53]]
        serialized = json.dumps(entries, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        if hashlib.sha256(serialized.encode("utf-8")).hexdigest() == _ORIGINAL_53_V2:
            local = replace(local, rules={number: rules for number, rules in local.rules.items() if number != 53})
    return local


def save_catalog(catalog: Catalog, path: Path) -> None:
    data = {
        "version": 3,
        "source": catalog.source,
        "points": catalog.points,
        "rules": {number: [_write_rule(rule) for rule in rules] for number, rules in catalog.rules.items()},
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
