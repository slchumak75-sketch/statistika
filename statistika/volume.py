"""Объём одной тары по названию товара или из ручного ввода."""

from __future__ import annotations

import re
import unicodedata
from decimal import Decimal
from functools import lru_cache

VOLUME_GROUPS = frozenset((*range(35, 44), 45))

_DECIMAL = r"\d+(?:[.,]\d+)?"
_MILLILITERS = ("мл", "ml", "миллилитр", "миллилитра", "миллилитров")
_UNIT = "|".join((*_MILLILITERS, "л", "l", "литр", "литра", "литров"))
# В «красн.0,75л» точка — конец сокращения, а не начало числа.
_VOLUME = re.compile(
    rf"(?<![\d,])(?P<number>{_DECIMAL}|(?<![\w.,])[.,]\d+)\s*(?P<unit>{_UNIT})(?!\w)", re.I,
)
_INPUT = re.compile(rf"(?P<number>{_DECIMAL}|[.,]\d+)(?:\s*(?P<unit>{_UNIT}))?", re.I)
# У составной фасовки неизвестно, считается ли упаковка одной штукой.
_CONNECTOR = r"(?:[xх×*/⁄+–—−-]|\b(?:до|или)\b)"
_COMPOUND_BEFORE = re.compile(rf"\d(?:\s*{_CONNECTOR}|\s+(?:[^\W\d_]+\.?\s+)?по)\s*$|[+−-]$", re.I)
_COMPOUND_AFTER = re.compile(rf"^\.?\s*{_CONNECTOR}\s*\d", re.I)
_DECIMAL_PREFIX = re.compile(r"\d[.,]+$")
_THOUSANDS_PREFIX = re.compile(r"(?<!\w)\d{1,3}\s+$")


def _liters(match: re.Match[str]) -> Decimal | None:
    amount = Decimal(match["number"].replace(",", "."))
    if amount <= 0:
        return None
    return amount.scaleb(-3) if (match["unit"] or "").casefold() in _MILLILITERS else amount


@lru_cache(maxsize=4096)
def find_container_liters(name: str) -> Decimal | None:
    """Автозаполнение только при одном явно указанном объёме, не из года или крепости."""
    text = unicodedata.normalize("NFKC", name)
    matches = list(_VOLUME.finditer(text))
    if len(matches) != 1:
        return None
    match = matches[0]
    before, after = text[: match.start()], text[match.end() :]
    if _COMPOUND_BEFORE.search(before) or _COMPOUND_AFTER.match(after) or _DECIMAL_PREFIX.search(before):
        return None
    # Не извлекаем фрагмент кода модели («B12L») или числа с неизвестным разделителем («1 500мл»).
    number = match["number"]
    if number.isdecimal() and (before[-1:].isalnum() or (len(number) == 3 and _THOUSANDS_PREFIX.search(before))):
        return None
    return _liters(match)


def parse_container_liters(text: str) -> Decimal | None:
    """Пустое поле снимает ручную правку. Число без единицы вводится в литрах."""
    text = unicodedata.normalize("NFKC", text).strip()
    if not text:
        return None
    match = _INPUT.fullmatch(text)
    if match is not None and (liters := _liters(match)) is not None:
        return liters
    raise ValueError("Укажите положительный объём тары в литрах, например 0,5, или очистите поле.")
