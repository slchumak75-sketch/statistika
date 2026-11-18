"""Диалоги справочника и импорта ведомости (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from openpyxl.utils.cell import get_column_letter

from statistika.catalog import (
    GROUPS, PDF_GROUPS, Catalog, CatalogError, builtin_catalog, extract_pdf_text, format_points, parse_points, save_catalog,
)
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


def _set_text(widget: tk.Text, text: str) -> None:
    widget.delete("1.0", "end")
    widget.insert("1.0", text)
    widget.edit_reset()


class CatalogDialog(Modal):
    def __init__(self, parent: tk.Misc, catalog: Catalog, path: Path) -> None:
        super().__init__(parent, "Справочник товарных групп")
        self.original = catalog
        self.path = path
        self.source = catalog.source
        self.result: Catalog | None = None
        self.geometry(f"{min(850, self.winfo_screenwidth() - 45)}x{min(600, self.winfo_screenheight() - 100)}")
        self.minsize(730, 440)
        body = self.body()
        body.rowconfigure(1, weight=1)
        ttk.Label(body, text="Справочник товарных групп", style="DialogTitle.TLabel").grid(row=0, sticky="w", pady=(0, 12))
        self.tabs = ttk.Notebook(body)
        self.tabs.grid(row=1, sticky="nsew")
        self.points_tab = ttk.Frame(self.tabs, padding=12)
        self.rules_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(self.points_tab, text="Пункты PDF")
        self.tabs.add(self.rules_tab, text="Правила распознавания")

        self._draw_points(self.points_tab, catalog)
        self._draw_rules(self.rules_tab, catalog)
        self.tabs.select(self.rules_tab)
        self.tabs.bind("<<NotebookTabChanged>>", self._tab_changed)
        self.save_button = self.footer(body, 2, "Сохранить справочник", self.save)

    def _draw_points(self, parent: ttk.Frame, catalog: Catalog) -> None:
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="Загрузить PDF…", command=self.open_pdf).pack(side="left")
        ttk.Button(toolbar, text="Встроенный справочник", command=self.restore_points).pack(side="left", padx=(7, 0))
        self.source_text = tk.StringVar(value=self.source or "PDF не загружен")
        ttk.Label(toolbar, textvariable=self.source_text).pack(side="left", padx=10)
        ttk.Label(
            parent, wraplength=675, text="Сверьте текст с PDF. Формат: «35. Описание». "
            "Можно исправить извлечение или ввести точный текст вручную.",
        ).pack(anchor="w", pady=(0, 8))
        self.points_text = ScrolledText(parent, width=74, height=15, wrap="word", undo=True, font="TkDefaultFont")
        self.points_text.pack(fill="both", expand=True)
        _set_text(self.points_text, format_points(catalog.points))
        count = len(PDF_GROUPS)
        self.note = tk.StringVar(value=f"Доступно пунктов PDF из требуемых: {count - len(catalog.missing)} из {count}. №100 встроен.")
        ttk.Label(parent, textvariable=self.note, wraplength=675, style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

    @staticmethod
    def _group_values(catalog: Catalog) -> tuple[str, ...]:
        return "Все группы", *(_short(catalog.short_label(number), 58) for number in GROUPS)

    def _draw_rules(self, parent: ttk.Frame, catalog: Catalog) -> None:
        choice = ttk.Frame(parent)
        choice.pack(fill="x", pady=(0, 10))
        ttk.Label(choice, text="Группа").pack(side="left")
        self.rule_group = ttk.Combobox(choice, width=42, state="readonly", values=self._group_values(catalog))
        self.rule_group.current(0)
        self.rule_group.pack(side="left", padx=8)
        self.rule_group.bind("<<ComboboxSelected>>", self._switch_group)
        self.rule_description = tk.StringVar(master=self)
        description = ttk.Label(
            parent, textvariable=self.rule_description, style="Muted.TLabel", wraplength=680, anchor="w", justify="left",
        )
        description.pack(fill="x", pady=(0, 10))
        parent.bind("<Configure>", lambda event: description.configure(wraplength=max(250, event.width - 24)))
        self.rule_editor = RuleEditor(parent, catalog)
        self.rule_editor.pack(fill="both", expand=True)
        self._describe_group()

    def _selected_group(self) -> int | None:
        index = self.rule_group.current()
        return GROUPS[index - 1] if index > 0 else None

    def _describe_group(self) -> None:
        group = self._selected_group()
        self.rule_description.set(
            self.rule_editor.catalog.label(group) if group is not None else "Выберите группу, чтобы увидеть описание и включить фразу."
        )

    def _switch_group(self, _event=None) -> None:
        self._describe_group()
        self.rule_editor.show_group(self._selected_group())

    def _read_points(self) -> dict[int, str]:
        text = self.points_text.get("1.0", "end-1c")
        points = parse_points(text)
        if text.strip() and not points:
            raise CatalogError("Не найдено ни одного пункта. Нужны строки вида «35. Описание».")
        return points

    def _tab_changed(self, _event=None) -> None:
        if self.tabs.select() != str(self.rules_tab):
            return
        try:
            points = self._read_points()
        except CatalogError:
            return  # При сохранении покажем ошибку; пока оставим прежние подписи.
        catalog = Catalog(points, source=self.source)
        index = max(0, self.rule_group.current())
        self.rule_group.configure(values=self._group_values(catalog))
        self.rule_group.current(index)
        self.rule_editor.update_catalog(catalog)
        self._describe_group()

    def restore_points(self) -> None:
        if not messagebox.askyesno(
            "Восстановить пункты?", "Заменить описания встроенным текстом 4-торг? Ваши правила останутся.", parent=self,
        ):
            return
        catalog = builtin_catalog()
        self.source = catalog.source
        self.source_text.set(self.source)
        _set_text(self.points_text, format_points(catalog.points))
        self.note.set(f"Встроенный справочник: {len(catalog.points)} пунктов. Нажмите «Сохранить справочник».")

    def open_pdf(self) -> None:
        name = filedialog.askopenfilename(parent=self, title="Выберите справочник PDF", filetypes=[("PDF", "*.pdf")])
        if not name:
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            text, blank_pages = extract_pdf_text(Path(name))
        except (CatalogError, OSError) as error:
            messagebox.showerror("Не удалось загрузить PDF", str(error), parent=self)
            return
        finally:
            self.configure(cursor="")
        if self.points_text.get("1.0", "end-1c").strip() and not messagebox.askyesno(
            "Заменить пункты?", "Извлечённый текст заменит пункты в редакторе. Правила останутся. Продолжить?", parent=self,
        ):
            return
        self.source = Path(name).name
        self.source_text.set(self.source)
        try:
            points = parse_points(text)
        except CatalogError as error:
            points = {}
            note = str(error)
        else:
            note = f"Извлечено пунктов: {len(points)}. Сверьте данные с PDF."
            if not points:
                note = "Нумерация не распознана. Исправьте текст так, чтобы каждый пункт начинался с «35. Описание»."
        _set_text(self.points_text, format_points(points) if points else text)
        if blank_pages:
            note += f" Страниц без текста: {blank_pages}."
        self.note.set(note)

    def save(self) -> None:
        try:
            points = self._read_points()
            catalog = Catalog(points, source=self.source if points else "")
            if catalog.missing and (points != self.original.points or catalog.source != self.original.source):
                missing = ", ".join(map(str, catalog.missing))
                if not messagebox.askyesno(
                    "Неполный справочник", f"Нет описаний пунктов: {missing}.\nСохранить как есть?", parent=self,
                ):
                    return
            self.rule_editor.update_catalog(catalog)
            if not self.rule_editor.include_phrase():
                self.tabs.select(self.rules_tab)
                return
            catalog = replace(catalog, rules=dict(self.rule_editor.rules))
            save_catalog(catalog, self.path)
        except (CatalogError, OSError) as error:
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
