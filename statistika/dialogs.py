"""Диалоги проверки импорта и назначения товарной группы (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from openpyxl.utils.cell import get_column_letter

from statistika.catalog import (
    GROUPS, PDF_GROUPS, Catalog, CatalogError, builtin_catalog, extract_pdf_text, format_points, parse_points, save_catalog,
)
from statistika.statement import (
    SOURCE_COLUMNS, ImportedProducts, ImportPlan, Sheet, StatementError, cell_at, column_headers,
    guess_plan, import_products, read_sheet,
)
from statistika.warehouse import SalesQuantity, detect_warehouse_layout, import_warehouse_products


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
        self.rules = dict(catalog.rules)
        self.editing_group = GROUPS[0]
        self.result: Catalog | None = None
        self.minsize(620, 460)
        body = self.body()
        body.rowconfigure(1, weight=1)
        ttk.Label(body, text="Справочник товарных групп", style="DialogTitle.TLabel").grid(row=0, sticky="w", pady=(0, 12))
        tabs = ttk.Notebook(body)
        tabs.grid(row=1, sticky="nsew")
        points = ttk.Frame(tabs, padding=12)
        rules = ttk.Frame(tabs, padding=12)
        tabs.add(points, text="Пункты PDF")
        tabs.add(rules, text="Правила распознавания")
        tabs.bind("<<NotebookTabChanged>>", lambda _event: self._show_rule_description())

        toolbar = ttk.Frame(points)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="Загрузить PDF…", command=self.open_pdf).pack(side="left")
        ttk.Button(toolbar, text="Встроенный справочник", command=self.restore_points).pack(side="left", padx=(7, 0))
        self.source_text = tk.StringVar(value=self.source or "PDF не загружен")
        ttk.Label(toolbar, textvariable=self.source_text).pack(side="left", padx=10)
        ttk.Label(
            points, wraplength=675, text="Сверьте текст с PDF. Формат: «35. Описание». "
            "Можно исправить извлечение или ввести точный текст вручную. Справочник запоминает все внесённые пункты.",
        ).pack(anchor="w", pady=(0, 8))
        self.points_text = ScrolledText(points, width=79, height=15, wrap="word", undo=True, font="TkDefaultFont")
        self.points_text.pack(fill="both", expand=True)
        _set_text(self.points_text, format_points(catalog.points))
        count = len(PDF_GROUPS)
        self.note = tk.StringVar(value=f"Доступно пунктов PDF из требуемых: {count - len(catalog.missing)} из {count}. №100 встроен.")
        ttk.Label(points, textvariable=self.note, wraplength=675, style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        ttk.Label(rules, text="Группа").pack(anchor="w")
        self.rule_group = ttk.Combobox(rules, width=12, state="readonly", values=[f"№{number}" for number in GROUPS])
        self.rule_group.current(0)
        self.rule_group.pack(anchor="w", pady=(3, 7))
        self.rule_group.bind("<<ComboboxSelected>>", self._switch_rule)
        self.rule_description = tk.StringVar()
        ttk.Label(rules, textvariable=self.rule_description, wraplength=675, style="Muted.TLabel").pack(anchor="w")
        ttk.Label(
            rules, wraplength=675, text="Введите характерные слова или фразы (по одной на строку). "
            "Правила дополняют совпадения с описанием PDF. Если совпало несколько групп, назначьте нужную вручную.",
        ).pack(anchor="w", pady=8)
        self.rules_text = ScrolledText(rules, height=9, width=79, wrap="word", undo=True, font="TkDefaultFont")
        self.rules_text.pack(fill="both", expand=True)
        _set_text(self.rules_text, "\n".join(self.rules.get(self.editing_group, ())))
        self._show_rule_description()
        self.footer(body, 2, "Сохранить справочник", self.save)

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

    def _remember_rule(self) -> None:
        text = self.rules_text.get("1.0", "end-1c")
        phrases = tuple(dict.fromkeys(line.strip() for line in text.splitlines() if line.strip()))
        if phrases:
            self.rules[self.editing_group] = phrases
        else:
            self.rules.pop(self.editing_group, None)

    def _switch_rule(self, _event) -> None:
        self._remember_rule()
        self.editing_group = GROUPS[self.rule_group.current()]
        _set_text(self.rules_text, "\n".join(self.rules.get(self.editing_group, ())))
        self._show_rule_description()

    def _show_rule_description(self) -> None:
        try:
            points = parse_points(self.points_text.get("1.0", "end-1c"))
        except CatalogError:
            points = self.original.points
        self.rule_description.set(_short(Catalog(points).label(self.editing_group), 320))

    def save(self) -> None:
        self._remember_rule()
        text = self.points_text.get("1.0", "end-1c")
        try:
            points = parse_points(text)
            if text.strip() and not points:
                raise CatalogError("Не найдено ни одного пункта. Нужны строки вида «35. Описание».")
            catalog = Catalog(points, self.rules, self.source if points else "")
            if catalog.missing and (points != self.original.points or self.source != self.original.source):
                missing = ", ".join(map(str, catalog.missing))
                if not messagebox.askyesno(
                    "Неполный справочник", f"Нет описаний пунктов: {missing}.\nСохранить как есть?", parent=self,
                ):
                    return
            save_catalog(catalog, self.path)
        except (CatalogError, OSError) as error:
            messagebox.showerror("Не удалось сохранить справочник", str(error), parent=self)
            return
        self.result = catalog
        self.destroy()


class StatementDialog(Modal):
    def __init__(self, parent: tk.Misc, path: Path, sheet_names: tuple[str, ...]) -> None:
        super().__init__(parent, "Импорт ведомости")
        self.path = path
        self.sheet = Sheet(())
        self.warehouse = None
        self.use_warehouse = tk.BooleanVar(value=False)
        self.sales_choices = (SalesQuantity.UNKNOWN, SalesQuantity.OUTFLOW)
        self.generic_note = ""
        self.result: ImportedProducts | None = None
        self.sheet_name = sheet_names[0]
        self.minsize(720, 570)
        body = self.body()
        body.rowconfigure(4, weight=1)
        ttk.Label(body, text="Проверьте столбцы ведомости", style="DialogTitle.TLabel").grid(row=0, sticky="w")
        ttk.Label(body, text=_short(path.name, 85), style="Muted.TLabel").grid(row=1, sticky="w", pady=(3, 10))

        settings = ttk.Frame(body)
        settings.grid(row=2, sticky="ew")
        ttk.Label(settings, text="Лист").pack(side="left")
        self.sheets = ttk.Combobox(settings, state="readonly", width=23, values=sheet_names)
        self.sheets.current(0)
        self.sheets.pack(side="left", padx=(5, 14))
        self.sheets.bind("<<ComboboxSelected>>", self.load_sheet)
        ttk.Label(settings, text="Строки заголовка с").pack(side="left")
        self.first = tk.StringVar(value="0")
        self.last = tk.StringVar(value="0")
        self.spins: list[ttk.Spinbox] = []
        for variable in (self.first, self.last):
            if variable is self.last:
                ttk.Label(settings, text="по").pack(side="left")
            spin = ttk.Spinbox(settings, from_=0, to=1, width=5, textvariable=variable, command=self.update_header)
            spin.pack(side="left", padx=5)
            spin.bind("<Return>", self.update_header)
            self.spins.append(spin)
        self.update_button = ttk.Button(settings, text="Обновить", command=self.update_header)
        self.update_button.pack(side="left", padx=(6, 0))
        self.header_hint = tk.StringVar()
        ttk.Label(
            body, textvariable=self.header_hint, style="Muted.TLabel",
        ).grid(row=3, sticky="w", pady=(6, 5))

        view = ttk.Frame(body)
        view.grid(row=4, sticky="nsew")
        view.rowconfigure(0, weight=1)
        view.columnconfigure(0, weight=1)
        self.preview = ttk.Treeview(view, show="headings", height=5, selectmode="none")
        self.preview.grid(row=0, column=0, sticky="nsew")
        sx = ttk.Scrollbar(view, orient="horizontal", command=self.preview.xview)
        sy = ttk.Scrollbar(view, orient="vertical", command=self.preview.yview)
        self.preview.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)
        sx.grid(row=1, column=0, sticky="ew")
        sy.grid(row=0, column=1, sticky="ns")

        panel = ttk.Frame(body)
        panel.grid(row=5, sticky="ew", pady=(10, 0))
        self.profile_switch = ttk.Checkbutton(
            panel, text="Профиль «Остатки ТМЦ на складах»", variable=self.use_warehouse, command=self._sync_mode,
        )
        self.warehouse_panel = ttk.Frame(panel)
        self.warehouse_fields = tk.StringVar()
        ttk.Label(self.warehouse_panel, textvariable=self.warehouse_fields, wraplength=690).pack(anchor="w", pady=(0, 8))
        ttk.Label(self.warehouse_panel, text="Источник количества продаж").pack(anchor="w")
        self.sales = ttk.Combobox(self.warehouse_panel, state="readonly", width=60)
        self.sales.pack(anchor="w", fill="x", pady=(3, 0))
        mapping = self.mapping = ttk.Frame(panel)
        mapping.columnconfigure(1, weight=1)
        self.columns: list[ttk.Combobox] = []
        for index, title in enumerate(SOURCE_COLUMNS):
            ttk.Label(mapping, text=title).grid(row=index, column=0, sticky="w", padx=(0, 12), pady=2)
            combo = ttk.Combobox(mapping, width=42, state="readonly")
            combo.grid(row=index, column=1, sticky="ew", pady=2)
            self.columns.append(combo)
        self.note = tk.StringVar()
        ttk.Label(body, textvariable=self.note, wraplength=690, style="Muted.TLabel").grid(row=6, sticky="w", pady=(7, 0))
        self.apply = self.footer(body, 7, "Построить таблицу", self.submit)
        self.load_sheet()

    def load_sheet(self, _event=None) -> None:
        self.sheet_name = self.sheets.get()
        self.configure(cursor="watch")
        self.update_idletasks()
        self.sheet = Sheet(())
        self.warehouse = None
        try:
            self.sheet = read_sheet(self.path, self.sheet_name)
            plan = guess_plan(self.sheet)
        except StatementError as error:
            self.generic_note = str(error)
            plan = guess_plan(self.sheet)
        else:
            detected = sum(column is not None for column in plan.columns)
            self.generic_note = (
                f"Предложено столбцов: {detected} из 5. Проверьте поля и дополните пропуски. "
                "«Расход» выбирайте только если он означает продажи."
                if self.sheet.width >= 5 else "На листе меньше пяти столбцов. Выберите другой лист или проверьте файл."
            )
            self.warehouse = detect_warehouse_layout(self.sheet)
        finally:
            self.configure(cursor="")
        if self.warehouse is not None and plan.first_header == 0:
            plan = ImportPlan(self.warehouse.header_row, self.warehouse.header_row, plan.columns)
        self.first.set(str(plan.first_header))
        self.last.set(str(plan.last_header))
        for spin in self.spins:
            spin.configure(to=len(self.sheet.rows))
        self.apply.state(["!disabled"] if self.sheet.width >= 5 else ["disabled"])
        self.profile_switch.pack_forget()
        self.mapping.pack_forget()
        self.warehouse_panel.pack_forget()
        self.use_warehouse.set(self.warehouse is not None and any(col is None for col in plan.columns))
        if self.warehouse is not None:
            self.profile_switch.pack(anchor="w", pady=(0, 7))
            layout = self.warehouse
            stock = get_column_letter(layout.ending_column + 1)
            price = get_column_letter(layout.price_column + 1)
            outflow = get_column_letter(layout.outflow_column + 1)
            name = get_column_letter(layout.name_column + 1)
            self.warehouse_fields.set(
                (f"Период: {layout.period}. " if layout.period else "")
                + f"Товар: {name}. Остаток: {stock}. Сумма остатка: {stock} × {price} (розничная цена)."
            )
            choices = ["Не заполнять — продажи явно не указаны"]
            sources = [SalesQuantity.UNKNOWN]
            if layout.month_columns:
                choices.append(f"Сложить помесячные количества ({layout.months_label})")
                sources.append(SalesQuantity.MONTHS)
            choices.append(f"Взять «Расход» ({outflow})")
            sources.append(SalesQuantity.OUTFLOW)
            self.sales_choices = tuple(sources)
            self.sales.configure(values=choices)
            self.sales.current(0)
        self._sync_mode()
        self._render_header(plan)

    def _sync_mode(self) -> None:
        self.mapping.pack_forget()
        self.warehouse_panel.pack_forget()
        active = self.use_warehouse.get() and self.warehouse is not None
        for widget in (*self.spins, self.update_button):
            widget.state(["disabled"] if active else ["!disabled"])
        if active:
            self.warehouse_panel.pack(fill="x")
            self.header_hint.set(f"Распознана ведомость ТМЦ · заголовок в строке {self.warehouse.header_row}.")
            self.note.set(
                "Выбирайте месяцы или «Расход» только если это продажи за нужный период. "
                "Фактической суммы продаж в ведомости нет — её столбец останется пустым."
            )
        else:
            self.mapping.pack(fill="x")
            self.header_hint.set("0 и 0 — без заголовка. Ниже показаны строки исходного листа.")
            self.note.set(self.generic_note)

    def _current_plan(self) -> ImportPlan:
        try:
            first, last = int(self.first.get()), int(self.last.get())
        except ValueError as error:
            raise StatementError("Введите номера строк заголовка целыми числами.") from error
        return ImportPlan(first, last, tuple(combo.current() - 1 if combo.current() > 0 else None for combo in self.columns))

    def _render_header(self, plan: ImportPlan) -> None:
        labels = column_headers(self.sheet, plan.first_header, plan.last_header)
        letters = [get_column_letter(index + 1) for index in range(self.sheet.width)]
        values = ["— выбрать столбец —", *(
            letter + (f" · {_short(label, 95)}" if label else "") for letter, label in zip(letters, labels)
        )]
        for combo, column in zip(self.columns, plan.columns):
            combo.configure(values=values)
            combo.current(0 if column is None else column + 1)

        # Остальные колонки по-прежнему доступны в списках сопоставления.
        letters = letters[:30]
        self.preview.delete(*self.preview.get_children())
        self.preview.configure(columns=("row", *letters))
        self.preview.heading("row", text="Строка")
        self.preview.column("row", width=56, stretch=False, anchor="e")
        for letter in letters:
            self.preview.heading(letter, text=letter)
            self.preview.column(letter, width=125, stretch=False)
        start = max(0, plan.first_header - 2)
        stop = min(len(self.sheet.rows), max(plan.last_header + 5, start + 8))
        indices = sorted(set(range(start, min(start + 8, stop))) | set(range(max(start, stop - 5), stop)))
        for index in indices:
            row = self.sheet.rows[index]
            self.preview.insert("", "end", values=(index + 1, *(
                "" if (value := cell_at(row, col)) is None else _short(" ".join(str(value).split()), 80)
                for col in range(len(letters))
            )))
        self.preview.yview_moveto(0)

    def update_header(self, _event=None) -> None:
        try:
            self._render_header(self._current_plan())
        except StatementError as error:
            messagebox.showerror("Проверьте заголовок", str(error), parent=self)

    def submit(self) -> None:
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            if self.use_warehouse.get() and self.warehouse is not None:
                self.result = import_warehouse_products(self.sheet, self.warehouse, self.sales_choices[self.sales.current()])
            else:
                self.result = import_products(self.sheet, self._current_plan())
        except StatementError as error:
            messagebox.showerror("Проверьте ведомость", str(error), parent=self)
        finally:
            self.configure(cursor="")
        if self.result is not None:
            self.destroy()


class GroupDialog(Modal):
    def __init__(self, parent: tk.Misc, catalog: Catalog, product: str, group: int | None) -> None:
        super().__init__(parent, "Назначение группы")
        self.catalog = catalog
        self.applied = False
        self.number = group
        body = self.body()
        ttk.Label(body, text="Товарная группа", style="DialogTitle.TLabel").grid(row=0, sticky="w")
        ttk.Label(body, text=_short(product, 170), wraplength=555).grid(row=1, sticky="w", pady=(4, 10))
        self.combo = ttk.Combobox(body, width=66, state="readonly", values=(
            "Без группы (пустая ячейка)", *(catalog.short_label(number) for number in GROUPS),
        ))
        self.combo.current(0 if group is None else GROUPS.index(group) + 1)
        self.combo.grid(row=2, sticky="ew")
        self.combo.bind("<<ComboboxSelected>>", self._describe)
        self.description = ScrolledText(body, height=4, width=58, wrap="word", font="TkDefaultFont")
        self.description.grid(row=3, sticky="ew", pady=(8, 0))
        self._describe()
        self.footer(body, 4, "Назначить", self.apply)
        self.bind("<Return>", lambda _event: self.apply())

    def _chosen(self) -> int | None:
        index = self.combo.current()
        return GROUPS[index - 1] if index > 0 else None

    def _describe(self, _event=None) -> None:
        self.description.configure(state="normal")
        _set_text(self.description, self.catalog.label(self._chosen()) or "Первый столбец останется пустым.")
        self.description.configure(state="disabled")

    def apply(self) -> None:
        self.number = self._chosen()
        self.applied = True
        self.destroy()
