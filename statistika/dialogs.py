"""Диалоги справочника и импорта ведомости (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl.utils.cell import get_column_letter

from statistika.catalog import Catalog, save_catalog
from statistika.rule_editor import RuleEditor
from statistika.statement import SOURCE_COLUMNS, ImportedProducts, ImportPlan, Sheet, StatementError, column_headers, import_products


class Modal(tk.Toplevel):
    def __init__(self, parent: tk.Misc, title: str) -> None:
        super().__init__(parent)
        self.withdraw()
        self.title(title)
        self.transient(parent)
        self.bind("<Escape>", lambda _event: self.destroy())

    def show(self) -> None:
        self.update_idletasks()
        parent = self.master
        x = max(0, parent.winfo_rootx() + (parent.winfo_width() - self.winfo_reqwidth()) // 2)
        y = max(0, parent.winfo_rooty() + (parent.winfo_height() - self.winfo_reqheight()) // 2)
        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.wait_visibility()
        self.grab_set()
        self.focus_set()
        self.wait_window()

    def footer(self, body: ttk.Frame, row: int, action: str, command) -> ttk.Button:
        buttons = ttk.Frame(body)
        buttons.grid(row=row, column=0, sticky="ew", pady=(14, 0))
        apply = ttk.Button(buttons, text=action, style="Accent.TButton", command=command)
        apply.pack(side="right")
        ttk.Button(buttons, text="Отмена", command=self.destroy).pack(side="right", padx=(0, 8))
        return apply

    def body(self) -> ttk.Frame:
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        return frame


def _short(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


class CatalogDialog(Modal):
    def __init__(self, parent: tk.Misc, catalog: Catalog, path: Path) -> None:
        super().__init__(parent, "Справочник товарных групп")
        self.original = catalog
        self.path = path
        self.result: Catalog | None = None
        self.geometry("750x520")
        self.minsize(650, 360)
        body = self.body()
        body.rowconfigure(0, weight=1)
        self.rule_editor = RuleEditor(body, catalog)
        self.rule_editor.grid(row=0, sticky="nsew")
        self.save_button = ttk.Button(body, text="Сохранить справочник", style="Accent.TButton", command=self.save)
        self.save_button.grid(row=1, sticky="e", pady=(12, 0))

    def save(self) -> None:
        if not self.rule_editor.include_phrase():
            return
        catalog = replace(self.original, rules=dict(self.rule_editor.rules))
        try:
            save_catalog(catalog, self.path)
        except OSError as error:
            messagebox.showerror("Не удалось сохранить справочник", str(error), parent=self)
            return
        self.result = catalog
        self.destroy()


class SheetDialog(Modal):
    def __init__(self, parent: tk.Misc, names: tuple[str, ...]) -> None:
        super().__init__(parent, "Выбрать лист ведомости")
        self.result: str | None = None
        body = self.body()
        ttk.Label(body, text="Какой лист открыть?", style="DialogTitle.TLabel").grid(row=0, sticky="w")
        self.sheets = ttk.Combobox(body, state="readonly", width=42, values=names)
        self.sheets.grid(row=1, sticky="ew", pady=(8, 0))
        self.sheets.current(0)
        self.footer(body, 2, "Открыть лист", self.submit)

    def submit(self) -> None:
        self.result = self.sheets.get()
        self.destroy()


class ColumnMappingDialog(Modal):
    """Только для нераспознанных заголовков. Без просмотра исходных строк."""

    def __init__(self, parent: tk.Misc, sheet: Sheet, plan: ImportPlan, sheet_name: str) -> None:
        super().__init__(parent, "Указать столбцы ведомости")
        self.sheet = sheet
        self.result: ImportedProducts | None = None
        body = self.body()
        ttk.Label(body, text="Укажите недостающие столбцы", style="DialogTitle.TLabel").grid(row=0, sticky="w")
        ttk.Label(body, text=f"Лист «{_short(sheet_name, 50)}»: не все поля определены автоматически.", style="Muted.TLabel").grid(
            row=1, sticky="w", pady=(3, 10),
        )

        settings = ttk.Frame(body)
        settings.grid(row=2, sticky="ew")
        ttk.Label(settings, text="Строки заголовка с").pack(side="left")
        self.first = tk.StringVar(value=str(plan.first_header))
        self.last = tk.StringVar(value=str(plan.last_header))
        for variable in (self.first, self.last):
            if variable is self.last:
                ttk.Label(settings, text="по").pack(side="left")
            spin = ttk.Spinbox(settings, from_=0, to=len(sheet.rows), width=5, textvariable=variable, command=self.update_header)
            spin.pack(side="left", padx=5)
            spin.bind("<Return>", self.update_header)
        ttk.Button(settings, text="Обновить заголовки", command=self.update_header).pack(side="left", padx=(6, 0))
        ttk.Label(body, text="Для файла без заголовка укажите 0 и 0.", style="Muted.TLabel").grid(
            row=3, sticky="w", pady=(4, 10),
        )
        mapping = ttk.Frame(body)
        mapping.grid(row=4, sticky="ew")
        mapping.columnconfigure(1, weight=1)
        self.columns: list[ttk.Combobox] = []
        for index, title in enumerate(SOURCE_COLUMNS):
            ttk.Label(mapping, text=title).grid(row=index, column=0, sticky="w", padx=(0, 12), pady=2)
            combo = ttk.Combobox(mapping, width=42, state="readonly")
            combo.grid(row=index, column=1, sticky="ew", pady=2)
            self.columns.append(combo)
        self._render_header(plan)
        self.footer(body, 5, "Показать отчёт", self.submit)

    def _current_plan(self) -> ImportPlan:
        try:
            first, last = int(self.first.get()), int(self.last.get())
        except ValueError as error:
            raise StatementError("Введите номера строк заголовка целыми числами.") from error
        return ImportPlan(first, last, tuple(combo.current() - 1 if combo.current() > 0 else None for combo in self.columns))

    def _render_header(self, plan: ImportPlan) -> None:
        labels = column_headers(self.sheet, plan.first_header, plan.last_header)
        values = ["Не выбрано", *(
            get_column_letter(index + 1) + (f" · {_short(label, 85)}" if label else "")
            for index, label in enumerate(labels)
        )]
        for combo, column in zip(self.columns, plan.columns):
            combo.configure(values=values)
            combo.current(0 if column is None else column + 1)

    def update_header(self, _event=None) -> None:
        try:
            self._render_header(self._current_plan())
        except StatementError as error:
            messagebox.showerror("Проверьте заголовок", str(error), parent=self)

    def submit(self) -> None:
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            self.result = import_products(self.sheet, self._current_plan())
        except StatementError as error:
            messagebox.showerror("Проверьте ведомость", str(error), parent=self)
        finally:
            self.configure(cursor="")
        if self.result is not None:
            self.destroy()
