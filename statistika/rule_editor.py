"""Поиск, включение и удаление фраз справочника."""

from __future__ import annotations

import tkinter as tk
import unicodedata
from dataclasses import replace
from tkinter import messagebox, ttk

from statistika.catalog import GROUPS, Catalog, RecognitionRule, clean_phrase


def _search_key(text: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", text).casefold().replace("ё", "е").split())


def _terms(rule: RecognitionRule) -> tuple[str, ...]:
    return rule.phrase, *rule.also, *rule.also_at_start


def _check(rule: RecognitionRule) -> str:
    return "☐" if rule.disabled else "☑"


class RuleEditor(ttk.Frame):
    def __init__(self, parent: tk.Misc, catalog: Catalog) -> None:
        super().__init__(parent)
        self.catalog = catalog
        self.rules = dict(catalog.rules)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        search = ttk.Frame(self)
        search.grid(row=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(search, text="Поиск по фразам").pack(side="left", padx=(0, 8))
        self.query = tk.StringVar(master=self)
        ttk.Entry(search, textvariable=self.query).pack(side="left", fill="x", expand=True)

        form = ttk.Frame(self)
        form.grid(row=1, columnspan=2, sticky="ew", pady=(0, 10))
        form.columnconfigure(0, weight=1)
        ttk.Label(form, text="Включить фразу").grid(row=0, column=0, sticky="w", pady=(0, 3))
        self.phrase = ttk.Entry(form)
        self.phrase.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.phrase.bind("<Return>", self._enter_phrase)
        labels = tuple(catalog.short_label(number) for number in GROUPS)
        self.group = ttk.Combobox(form, state="readonly", width=29, values=(
            "Выберите группу", *(label if len(label) <= 46 else label[:45] + "…" for label in labels),
        ))
        self.group.current(0)
        self.group.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        self.include_button = ttk.Button(form, text="Включить", command=self.include_phrase)
        self.include_button.grid(row=1, column=2, sticky="e")

        self.tree = ttk.Treeview(self, columns=("enabled", "phrase", "delete"), show="", selectmode="browse", height=14)
        self.tree.column("enabled", width=36, minwidth=36, anchor="center", stretch=False)
        self.tree.column("phrase", width=450, minwidth=220, anchor="w")
        self.tree.column("delete", width=90, minwidth=90, anchor="center", stretch=False)
        self.tree.tag_configure("group", background="#e0eaf5", font=("Segoe UI", 9, "bold"))
        self.tree.tag_configure("disabled", foreground="#617488")
        self.tree.grid(row=2, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scroll.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<ButtonRelease-1>", self._click)
        self.tree.bind("<space>", self._toggle_selected)
        self.tree.bind("<Delete>", self._delete_selected)
        self.query.trace_add("write", lambda *_: self._refresh())
        self._refresh()

    def _rules_for(self, number: int) -> tuple[RecognitionRule, ...]:
        return self.rules[number] if number in self.rules else self.catalog.recognition_rules(number)

    def _refresh(self, *, keep_position: bool = False, selected: str = "") -> None:
        tree = self.tree
        position = tree.yview()[0] if keep_position else 0.0
        query = _search_key(self.query.get())
        tree.delete(*tree.get_children())
        count = 0
        for number in GROUPS:
            title_added = False
            for index, rule in enumerate(self._rules_for(number)):
                terms = _terms(rule)
                if query and query not in _search_key(" ".join(terms)):
                    continue
                if not title_added:
                    tree.insert("", "end", values=("", self.catalog.short_label(number), ""), tags=("group",))
                    title_added = True
                label = " · ".join(terms) + (" (исключение)" if rule.exclude else "")
                tree.insert("", "end", iid=f"{number}:{index}", values=(_check(rule), label, "Удалить"),
                            tags=("disabled",) if rule.disabled else ())
                count += 1
        if not count:
            tree.insert("", "end", values=("", "Фразы не найдены", ""))
        if selected and tree.exists(selected):
            tree.selection_set(selected)
            tree.focus(selected)
            tree.see(selected)
        else:
            tree.yview_moveto(position)

    def _enter_phrase(self, _event) -> str:
        self.include_phrase()
        return "break"

    def include_phrase(self) -> bool:
        raw_phrase = self.phrase.get()
        if not raw_phrase.strip():
            return True
        phrase = clean_phrase(raw_phrase)
        selection = self.group.current()
        if not any(ch.isalnum() for ch in phrase) or selection <= 0:
            messagebox.showwarning("Включить фразу", "Введите фразу с буквами или цифрами и выберите её группу.", parent=self)
            return False
        number = GROUPS[selection - 1]
        rules = list(self._rules_for(number))
        key = _search_key(phrase)
        index = next((i for i, rule in enumerate(rules) if not rule.exclude and key == _search_key(" ".join(_terms(rule)))), None)
        if index is None:
            index = len(rules)
            rules.append(RecognitionRule(phrase))
        elif rules[index].disabled:
            rules[index] = replace(rules[index], disabled=False)
        self.rules[number] = tuple(rules)
        self.phrase.delete(0, "end")
        if _search_key(self.query.get()) not in key:
            self.query.set(phrase)
        self._refresh(selected=f"{number}:{index}")
        return True

    def _edit_row(self, item: str, *, delete: bool) -> None:
        if ":" not in item:
            return  # Не строка фразы: название группы или пустой результат поиска.
        number, index = map(int, item.split(":"))
        rules = list(self._rules_for(number))
        if delete:
            del rules[index]
        else:
            rule = rules[index] = replace(rules[index], disabled=not rules[index].disabled)
            self.tree.set(item, "enabled", _check(rule))
            self.tree.item(item, tags=("disabled",) if rule.disabled else ())
        self.rules[number] = tuple(rules)
        if delete:
            self._refresh(keep_position=True)

    def _click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if column in ("#1", "#3"):
            self._edit_row(item, delete=column == "#3")

    def _toggle_selected(self, _event) -> str:
        self._edit_row(next(iter(self.tree.selection()), ""), delete=False)
        return "break"

    def _delete_selected(self, _event) -> str:
        self._edit_row(next(iter(self.tree.selection()), ""), delete=True)
        return "break"
