"""Прозрачная эвристика: совпадения по словам, неоднозначность оставляем оператору."""

from __future__ import annotations

import re
from collections import defaultdict
from functools import lru_cache
from typing import Iterable

from snowballstemmer.russian_stemmer import RussianStemmer

from statistika.catalog import GROUPS, HOSIERY, HOSIERY_DESCRIPTION, PDF_GROUPS, Catalog

_HOSIERY_TERMS = (
    HOSIERY_DESCRIPTION,
    "носки", "носок", "носочки", "носочек", "чулки", "чулок", "колготки", "колготок", "колготы",
    "гольфы", "полугольфы", "получулки", "получулок", "подследники", "следки", "гетры",
    "термоноски", "термоносок", "термогольфы", "термоколготки", "термоколготок",
)
_WORDS = re.compile(r"[^\W_]+", re.UNICODE)
_LIMITS = re.compile(r"\b(?:кроме|за исключением|исключая|не включая|не учтенн\w*|только)\b", re.I)
_GENERIC = {
    "товары", "изделия", "продукция", "прочие", "другие", "наборы", "комплекты", "разные", "в том числе",
}
_STEMMER = RussianStemmer()


@lru_cache(maxsize=32768)
def _stem(word: str) -> str:
    return _STEMMER.stemWord(word) if re.fullmatch(r"[а-я]+", word) else word


def _tokens(text: str) -> tuple[str, ...]:
    return tuple(_stem(word) for word in _WORDS.findall(text.casefold().replace("ё", "е")))


def _description_terms(description: str) -> Iterable[str]:
    # Без точного исходника нельзя додумывать синонимы или толковать исключения.
    # Составное описание с оговоркой требует проверенного ручного правила.
    if _LIMITS.search(description.replace("ё", "е")) or "(" in description or "[" in description:
        return
    for phrase in re.split(r"[,;]", description):
        phrase = phrase.strip(" .:-—–")
        if len(phrase) >= 3 and _tokens(phrase) not in _GENERIC_TOKENS:
            yield phrase


_GENERIC_TOKENS = {_tokens(term) for term in _GENERIC}


class Classifier:
    def __init__(self, catalog: Catalog) -> None:
        entries: dict[str, set[tuple[tuple[str, ...], int]]] = defaultdict(set)

        def add(number: int, phrases: Iterable[str]) -> None:
            for phrase in phrases:
                tokens = _tokens(phrase)
                if tokens:
                    entries[tokens[0]].add((tokens, number))

        for number in PDF_GROUPS:
            add(number, _description_terms(catalog.points.get(number, "")))
        add(HOSIERY, _HOSIERY_TERMS)
        for number in GROUPS:
            add(number, catalog.rules.get(number, ()))
        self._by_start = dict(entries)

    def group_for(self, name: str) -> int | None:
        tokens = _tokens(name)
        match: int | None = None
        for index, token in enumerate(tokens):
            for phrase, number in self._by_start.get(token, ()):
                if tokens[index : index + len(phrase)] == phrase:
                    if match is not None and match != number:
                        return None
                    match = number
        return match
