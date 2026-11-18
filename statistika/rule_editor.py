"""Редактор фраз: способ совпадения и исключения — поля, а не разметка текста."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

from statistika.catalog import RecognitionRule, clean_phrase

_POSITIONS = ("В названии", "В начале")
_ACTIONS = ("Назначить", "Исключить")


class RuleEditor(ttk.Frame):
    def __init__(self, parent: tk.Misc, on_change: Callable[[tuple[RecognitionRule, ...]], None]) -> None:
        super().__init__(parent)
        self.on_change = on_change
        self.rules: list[RecognitionRule] = []
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        search = ttk.Frame(self)
        search.grid(row=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(search, text="Поиск по фразам").pack(side="left", padx=(0, 6))
        self.query = tk.StringVar(master=self)
        ttk.Entry(search, textvariable=self.query, width=36).pack(side="left")
        self.count = ttk.Label(search, style="Muted.TLabel")
        self.count.pack(side="right")
        self.query.trace_add("write", lambda *_: self._refresh())

        self.tree = ttk.Treeview(self, columns=("phrase", "position", "action", "details"), show="headings",
                                 selectmode="browse", height=8)
        for name, title, width in zip(self.tree["columns"],
                                      ("Фраза", "Где искать", "Действие", "Также должно встретиться"),
                                      (250, 105, 130, 225)):
            self.tree.heading(name, text=title, anchor="w")
            self.tree.column(name, width=width, minwidth=90, stretch=name in ("phrase", "details"))
        self.tree.grid(row=1, column=0, sticky="nsew")
        sy = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        sy.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sy.set)
        self.tree.bind("<<TreeviewSelect>>", self._selected)
        self.tree.bind("<Delete>", self.delete)

        form = ttk.Frame(self)
        form.grid(row=2, columnspan=2, sticky="ew", pady=(8, 4))
        form.columnconfigure(0, weight=1)
        ttk.Label(form, text="Фраза").grid(row=0, column=0, sticky="w")
        self.phrase = ttk.Entry(form)
        self.phrase.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(form, text="Где искать").grid(row=0, column=1, sticky="w")
        self.position = ttk.Combobox(form, values=_POSITIONS, state="readonly", width=16)
        self.position.grid(row=1, column=1, sticky="w", padx=(0, 8))
        ttk.Label(form, text="Действие").grid(row=0, column=2, sticky="w")
        self.action = ttk.Combobox(form, values=_ACTIONS, state="readonly", width=15)
        self.action.grid(row=1, column=2, sticky="w")

        details = ttk.Frame(self)
        details.grid(row=3, columnspan=2, sticky="ew")
        self.also = self._text_field(details, "Также в названии (если нужно)", 0)
        self.also_at_start = self._text_field(details, "Также в начале (если нужно)", 1)

        buttons = ttk.Frame(self)
        buttons.grid(row=4, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(buttons, text="Добавить", command=self.add).pack(side="left")
        self.change_button = ttk.Button(buttons, text="Изменить", command=self.change)
        self.change_button.pack(side="left", padx=6)
        self.delete_button = ttk.Button(buttons, text="Удалить", command=self.delete)
        self.delete_button.pack(side="left")
        ttk.Button(buttons, text="Очистить список…", command=self.clear).pack(side="right")
        self._fill(RecognitionRule(""))
        self._selected()

    @staticmethod
    def _text_field(parent: ttk.Frame, label: str, column: int) -> tk.Text:
        ttk.Label(parent, text=label, style="Muted.TLabel").grid(row=0, column=column, sticky="w")
        parent.columnconfigure(column, weight=1)
        field = tk.Text(parent, height=2, width=24, wrap="word", font="TkDefaultFont")
        field.grid(row=1, column=column, sticky="ew", padx=(0, 6) if column == 0 else 0)
        return field

    def load(self, rules: tuple[RecognitionRule, ...]) -> None:
        self.rules = list(rules)
        self.query.set("")
        self._fill(RecognitionRule(""))

    def _refresh(self, selected: int | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        query = self.query.get().strip().casefold().replace("ё", "е")
        shown = 0
        for index, rule in enumerate(self.rules):
            extra = "; ".join((*rule.also, *(p + " (в начале)" for p in rule.also_at_start)))
            values = rule.phrase, _POSITIONS[rule.at_start], _ACTIONS[rule.exclude], extra
            if query not in " ".join(values).casefold().replace("ё", "е"):
                continue
            self.tree.insert("", "end", iid=str(index), values=values)
            shown += 1
        if selected is not None and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))
            self.tree.see(str(selected))
        self.count.configure(text=f"Правил: {shown} из {len(self.rules)}")
        self._selected()

    def _index(self) -> int | None:
        selection = self.tree.selection()
        return int(selection[0]) if selection else None

    def _selected(self, _event=None) -> None:
        index = self._index()
        for button in (self.change_button, self.delete_button):
            button.state(["disabled" if index is None else "!disabled"])
        if index is not None:
            self._fill(self.rules[index])

    def _fill(self, rule: RecognitionRule) -> None:
        self.phrase.delete(0, "end")
        self.phrase.insert(0, rule.phrase)
        self.position.current(int(rule.at_start))
        self.action.current(int(rule.exclude))
        for field, phrases in ((self.also, rule.also), (self.also_at_start, rule.also_at_start)):
            field.delete("1.0", "end")
            field.insert("1.0", "\n".join(phrases))

    def _read(self) -> RecognitionRule | None:
        phrase = clean_phrase(self.phrase.get())
        also, starts = (
            tuple(dict.fromkeys(p for line in field.get("1.0", "end-1c").splitlines() if (p := clean_phrase(line))))
            for field in (self.also, self.also_at_start)
        )
        if not all(any(ch.isalnum() for ch in value) for value in (phrase, *also, *starts)):
            messagebox.showwarning("Правило распознавания", "Укажите фразу с буквами или цифрами.", parent=self)
            return None
        return RecognitionRule(phrase, self.position.current() == 1, self.action.current() == 1, also, starts)

    def _commit(self, selected: int | None = None) -> None:
        self.on_change(tuple(self.rules))
        self._refresh(selected)

    def add(self) -> None:
        if rule := self._read():
            if rule not in self.rules:
                self.rules.append(rule)
                self._commit(len(self.rules) - 1)
            else:
                self._refresh(self.rules.index(rule))

    def change(self) -> None:
        if (index := self._index()) is not None and (rule := self._read()):
            self.rules[index] = rule
            self._commit(index)

    def delete(self, _event=None) -> None:
        if (index := self._index()) is not None:
            del self.rules[index]
            self._commit(min(index, len(self.rules) - 1) if self.rules else None)
            if not self.rules:
                self._fill(RecognitionRule(""))

    def clear(self) -> None:
        if self.rules and messagebox.askyesno(
            "Очистить список?", "Автоназначение для этой группы будет отключено после сохранения.", parent=self,
        ):
            self.rules.clear()
            self._commit()
            self._fill(RecognitionRule(""))
