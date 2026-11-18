"""Сопоставление названий с редактируемыми фразами справочника."""

from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache

from snowballstemmer.russian_stemmer import RussianStemmer

from statistika.catalog import GROUPS, Catalog, RecognitionRule

_WORDS = re.compile(r"\d+(?:[.,]\d+)?\s*%|[^\W_]+", re.UNICODE)
_FOR = re.compile(r"\bд\s*[/\\]\s*")
_VOLUME = re.compile(r"(?<=\d)(?=(?:кг|мл|л)\b)")
_CODE = re.compile(r"\d{1,3}")
_STEMMER = RussianStemmer()


@lru_cache(maxsize=32768)
def _stem(word: str) -> str:
    return _STEMMER.stemWord(word) if re.fullmatch(r"[а-я]+", word) else word


def _tokens(text: str) -> tuple[str, ...]:
    text = unicodedata.normalize("NFKC", text).casefold().replace("ё", "е")
    text = _VOLUME.sub(" ", _FOR.sub("для ", text))

    def normalize(word: str) -> str:
        if word.endswith("%"):
            # «0 %», «0,0%» и «0.00 %» — одно значение, но «0,5%» не ноль.
            whole, _, fraction = word[:-1].strip().replace(",", ".").partition(".")
            whole = whole.lstrip("0") or "0"
            fraction = fraction.rstrip("0")
            return f"{whole}.{fraction}%" if fraction else f"{whole}%"
        return _stem(word)

    return tuple(normalize(word) for word in _WORDS.findall(text))


@dataclass(frozen=True)
class Phrase:
    words: tuple[str, ...]
    at_start: bool

    def matches_at(self, words: tuple[str, ...], index: int, start: int) -> bool:
        return (not self.at_start or index in (0, start)) and words[index : index + len(self.words)] == self.words

    def occurs(self, words: tuple[str, ...], start: int) -> bool:
        if self.at_start:
            return self.matches_at(words, 0, start) or self.matches_at(words, start, start)
        return any(self.matches_at(words, index, start) for index, word in enumerate(words) if word == self.words[0])


@dataclass(frozen=True)
class Rule:
    first: Phrase
    additional: tuple[Phrase, ...]
    group: int
    excludes: bool


def _rule(rule: RecognitionRule, group: int) -> Rule | None:
    first = Phrase(_tokens(rule.phrase), rule.at_start)
    additional = tuple(Phrase(_tokens(phrase), False) for phrase in rule.also) + tuple(
        Phrase(_tokens(phrase), True) for phrase in rule.also_at_start
    )
    if not first.words or any(not part.words for part in additional):
        return None
    return Rule(first, additional, group, rule.exclude)


class Classifier:
    def __init__(self, catalog: Catalog, *, include_electrical: bool = False) -> None:
        by_word: dict[str, set[Rule]] = defaultdict(set)
        for group in GROUPS:
            for entry in catalog.recognition_rules(group):
                if entry.internal_only and not include_electrical:
                    continue
                if rule := _rule(entry, group):
                    by_word[rule.first.words[0]].add(rule)
        self._by_word = dict(by_word)

    def group_for(self, name: str) -> int | None:
        words = _tokens(name)
        # «1 Колено PPR …»: разрешаем числовой префикс перед началом названия.
        start = 1 if words and _CODE.fullmatch(words[0]) else 0
        matches: set[int] = set()
        blocked: set[int] = set()
        for index, word in enumerate(words):
            for rule in self._by_word.get(word, ()):
                marked = blocked if rule.excludes else matches
                if rule.group not in marked and rule.first.matches_at(words, index, start) and all(
                    part.occurs(words, start) for part in rule.additional
                ):
                    marked.add(rule.group)
        matches.difference_update(blocked)
        return matches.pop() if len(matches) == 1 else None
