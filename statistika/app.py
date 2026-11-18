"""Окно программы; вся обработка файлов находится в отдельных модулях."""

from __future__ import annotations

import argparse
import tempfile
import tkinter as tk
import traceback
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from statistika.catalog import GROUPS, PDF_GROUPS, Catalog, CatalogError, load_catalog, local_catalog_path
from statistika.dialogs import CatalogDialog, ColumnMappingDialog, SheetDialog
from statistika.inline_picker import InlineGroupPicker
from statistika.report import (
    ALL_WAREHOUSES, QUANTITY_NOTE, RESULT_COLUMNS, SUMMARY_COLUMNS, SUMMARY_NOTE, ReportRow,
    display_number, display_quantity, display_total, export_xlsx, in_warehouse, refresh_groups, summarize_groups, warehouses,
)
from statistika.statement import ImportedProducts, ReportPeriod, Sheet, StatementError, guess_plan, import_products, list_sheets, read_sheet
from statistika.warehouse import SalesQuantity, WarehouseLayout, detect_warehouse_layout, import_warehouse_products

_BG = "#f1f5fa"
_INK = "#16304c"
_MUTED = "#466078"
_BLUE = "#174f91"
_ALL_ROWS, _ASSIGNED_ROWS, _UNASSIGNED_ROWS = "all", "assigned", "unassigned"
_GROUP_FILTER_PRESETS = {
    _ALL_ROWS: "Все строки", _ASSIGNED_ROWS: "Только с группой", _UNASSIGNED_ROWS: "Только без группы",
}
_SALES_LABELS = {
    SalesQuantity.MONTHS: "По месяцам",
    SalesQuantity.UNKNOWN: "Не заполнять",
    SalesQuantity.OUTFLOW: "Расход за весь период",
}


def _sales_options(layout: WarehouseLayout) -> tuple[SalesQuantity, ...]:
    return (*((SalesQuantity.MONTHS,) if layout.months else ()), SalesQuantity.UNKNOWN, SalesQuantity.OUTFLOW)


def setup_theme(root: tk.Tk) -> None:
    root.configure(background=_BG)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=_BG, foreground=_INK, font="TkDefaultFont")
    style.configure("TButton", padding=(11, 6))
    style.configure("Accent.TButton", background=_BLUE, foreground="#ffffff", bordercolor=_BLUE)
    style.map("Accent.TButton", background=[("disabled", "#dbe4ed"), ("active", "#123e72")], foreground=[("disabled", _MUTED)])
    style.configure("Kicker.TLabel", foreground=_BLUE, font=("Segoe UI", 9, "bold"))
    style.configure("Title.TLabel", font=("Segoe UI", 24, "bold"))
    style.configure("DialogTitle.TLabel", font=("Segoe UI", 14, "bold"))
    style.configure("Muted.TLabel", foreground=_MUTED)
    style.configure("Counter.TLabel", padding=(10, 5), background="#e0eaf5", foreground=_INK)
    style.configure("Treeview", rowheight=27, background="#ffffff", fieldbackground="#ffffff")
    style.configure("Treeview.Heading", padding=(5, 7), background="#dde8f4", foreground=_INK)
    style.map("Treeview", background=[("selected", "#c6dffa")], foreground=[("selected", _INK)])
    style.configure("Empty.TLabel", background="#ffffff", foreground=_MUTED, padding=10)


class MainWindow:
    def __init__(self, root: tk.Tk, catalog_path: Path) -> None:
        self.root = root
        self.catalog_path = catalog_path
        self.catalog = Catalog()
        self.rows: list[ReportRow] = []
        self.source_path: Path | None = None
        self.column_notes: tuple[tuple[int, str], ...] = ()
        self.unsaved: set[int] = set()
        self.period = ReportPeriod()
        self.warehouse_source: tuple[Sheet, WarehouseLayout] | None = None
        self.sales_options: tuple[SalesQuantity, ...] = ()
        self.applied_source = SalesQuantity.UNKNOWN
        self.month_range: tuple[int, int] | None = None
        self.warehouse_options: tuple[str | None, ...] = (None,)
        self.group_filter = tk.StringVar(value=_ALL_ROWS)
        self.period_text = tk.StringVar()
        self.catalog_summary = tk.StringVar()
        self.source_text = tk.StringVar(value="Ведомость ещё не открыта")
        self.total_text = tk.StringVar(value="Строки: 0")
        self.assigned_text = tk.StringVar(value="С группой: 0")
        self.unassigned_text = tk.StringVar(value="Без группы: 0")
        self.selection_text = tk.StringVar()
        self.export_text = tk.StringVar()
        self.method_text = tk.StringVar()

        root.title("Статистика — продажи и остатки")
        root.geometry(f"{min(1240, root.winfo_screenwidth() - 55)}x{min(760, root.winfo_screenheight() - 100)}")
        root.minsize(850, 540)
        root.protocol("WM_DELETE_WINDOW", self.close)
        setup_theme(root)
        self._draw()
        try:
            self.catalog = load_catalog(catalog_path)
        except CatalogError as error:
            self.export_text.set(str(error))
            messagebox.showwarning("Справочник не прочитан", str(error), parent=root)
        self.update_catalog_summary()
        self.redraw()

    def _draw(self) -> None:
        body = ttk.Frame(self.root, padding=(20, 15))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(5, weight=1)
        top = ttk.Frame(body)
        top.grid(row=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="СТАТИСТИКА  /  ТОВАРНЫЙ ОТЧЁТ", style="Kicker.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text="Продажи и остатки", style="Title.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 1))
        ttk.Label(top, textvariable=self.catalog_summary, style="Muted.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Button(top, text="Справочник групп…", command=self.edit_catalog).grid(row=1, column=1, sticky="e")

        actions = ttk.Frame(body)
        actions.grid(row=1, sticky="ew", pady=(14, 10))
        ttk.Button(actions, text="Открыть ведомость…", style="Accent.TButton", command=self.open_statement).pack(side="left")
        self.export_button = ttk.Button(actions, text="Сохранить Excel…", command=self.save_report)
        self.export_button.pack(side="left", padx=8)
        self.edit_button = ttk.Button(actions, text="Изменить группу…", command=self.edit_selected)
        self.edit_button.pack(side="left")
        ttk.Label(actions, textvariable=self.source_text, style="Muted.TLabel", wraplength=270, justify="right").pack(
            side="right", padx=(8, 0),
        )

        filters = ttk.Frame(body)
        filters.grid(row=2, sticky="ew", pady=(0, 8))
        ttk.Label(filters, text="Склад").pack(side="left")
        self.warehouse_selector = ttk.Combobox(filters, state="disabled", values=(ALL_WAREHOUSES,), width=21)
        self.warehouse_selector.set(ALL_WAREHOUSES)
        self.warehouse_selector.pack(side="left", padx=(6, 16))
        self.warehouse_selector.bind("<<ComboboxSelected>>", lambda _event: self.redraw())

        self.period_fields = ttk.Frame(body)
        self.period_fields.grid(row=3, sticky="ew", pady=(0, 8))
        ttk.Label(self.period_fields, text="Количество").pack(side="left")
        self.sales_selector = ttk.Combobox(self.period_fields, state="readonly", width=25)
        self.sales_selector.pack(side="left", padx=(6, 16))
        self.month_fields = ttk.Frame(self.period_fields)
        ttk.Label(self.month_fields, text="Период с").pack(side="left")
        self.first_month = ttk.Combobox(self.month_fields, state="readonly", width=18)
        self.first_month.pack(side="left", padx=6)
        ttk.Label(self.month_fields, text="по").pack(side="left")
        self.last_month = ttk.Combobox(self.month_fields, state="readonly", width=18)
        self.last_month.pack(side="left", padx=6)
        for combo in (self.sales_selector, self.first_month, self.last_month):
            combo.bind("<<ComboboxSelected>>", self.change_period)
        self.period_fields.grid_remove()

        stats = ttk.Frame(body)
        stats.grid(row=4, sticky="ew", pady=(0, 8))
        for variable in (self.total_text, self.assigned_text, self.unassigned_text):
            ttk.Label(stats, textvariable=variable, style="Counter.TLabel").pack(side="left", padx=(0, 7))
        period_label = ttk.Label(stats, textvariable=self.period_text, style="Muted.TLabel", wraplength=500)
        period_label.pack(side="left", fill="x", expand=True, padx=4)
        period_label.bind("<Configure>", lambda event: period_label.configure(wraplength=max(210, event.width - 2)))

        self.tabs = ttk.Notebook(body)
        self.tabs.grid(row=5, sticky="nsew")
        table = self.detail_tab = ttk.Frame(self.tabs)
        self.summary_tab = ttk.Frame(self.tabs)
        self.tabs.add(table, text="Товары")
        self.tabs.add(self.summary_tab, text="Сводка по группам")
        self.tabs.bind("<<NotebookTabChanged>>", self.refresh_selection)
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        names = ("group", "name", "sold_qty", "sold_amount", "stock_qty", "stock_amount")
        self.tree = ttk.Treeview(table, columns=names, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        headers = (*RESULT_COLUMNS[:2], "Продано, шт.*", "Продано, руб.", "Остаток, шт.*", "Остаток, руб.")
        for index, (name, title, width) in enumerate(zip(names, headers, (228, 246, 111, 111, 111, 111))):
            self.tree.heading(name, text=title, anchor="w" if index < 2 else "center")
            self.tree.column(name, width=width, minwidth=min(width, 172), anchor="w" if index < 2 else "e", stretch=index < 2)
        self.group_menu = tk.Menu(self.tree, tearoff=False)
        self.tree.heading("group", command=self.open_group_filter)
        sy = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.picker = InlineGroupPicker(self.tree, self.assign_group)

        def update_scrollbar(bar, first, last) -> None:
            self.picker.dismiss()
            bar.set(first, last)

        self.tree.configure(yscrollcommand=lambda a, b: update_scrollbar(sy, a, b),
                            xscrollcommand=lambda a, b: update_scrollbar(sx, a, b))
        self.tree.tag_configure("odd", background="#f1f6fc")
        self.tree.bind("<<TreeviewSelect>>", self.refresh_selection)
        self.tree.bind("<Return>", self.edit_selected)
        self.tree.bind("<ButtonRelease-1>", self._click_cell)
        self.empty = ttk.Label(table, text="", style="Empty.TLabel", justify="center")
        self._draw_summary()

        detail = ttk.Label(body, textvariable=self.selection_text, wraplength=810, style="Muted.TLabel")
        detail.grid(row=6, sticky="ew", pady=(10, 2))
        detail.bind("<Configure>", lambda event: detail.configure(wraplength=max(350, event.width - 2)))
        ttk.Label(body, textvariable=self.export_text, style="Muted.TLabel").grid(row=7, sticky="w")
        method = ttk.Label(body, textvariable=self.method_text, wraplength=810, style="Muted.TLabel")
        method.grid(row=8, sticky="ew", pady=(2, 0))
        method.bind("<Configure>", lambda event: method.configure(wraplength=max(350, event.width - 2)))

    def _draw_summary(self) -> None:
        tab = self.summary_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        columns = ("group", "lines", "sold_qty", "sold_amount", "stock_qty", "stock_amount")
        self.summary_tree = ttk.Treeview(tab, columns=columns, show="headings", selectmode="browse")
        self.summary_tree.grid(row=0, column=0, sticky="nsew")
        headers = ("Группа", "Строк", *SUMMARY_COLUMNS[2:])
        for index, (column, label, width) in enumerate(zip(columns, headers, (276, 63, 120, 120, 120, 120))):
            self.summary_tree.column(column, width=width, minwidth=min(width, 175), anchor="w" if not index else "e", stretch=not index)
            self.summary_tree.heading(column, text=label, anchor="w" if not index else "center")
        sy = ttk.Scrollbar(tab, orient="vertical", command=self.summary_tree.yview)
        sx = ttk.Scrollbar(tab, orient="horizontal", command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        ttk.Label(tab, text=SUMMARY_NOTE, style="Muted.TLabel", wraplength=740).grid(row=2, sticky="w", padx=6, pady=4)

    def active_warehouse(self) -> str | None:
        index = self.warehouse_selector.current()
        return self.warehouse_options[index] if index >= 0 else None

    def warehouse_rows(self) -> list[ReportRow]:
        warehouse = self.active_warehouse()
        return [row for row in self.rows if in_warehouse(row, warehouse)]

    def refresh_summary(self, rows: list[ReportRow]) -> None:
        tree = self.summary_tree
        tree.delete(*tree.get_children())
        for group in summarize_groups(rows) if self.rows else ():
            tree.insert("", "end", iid=str(group.group), values=(
                self.catalog.short_label(group.group) or "Без группы", group.lines,
                *(display_total(total) for total in group.metrics),
            ))

    def update_catalog_summary(self) -> None:
        ready = len(PDF_GROUPS) - len(self.catalog.missing)
        self.catalog_summary.set(f"Справочник: {ready} из {len(PDF_GROUPS)} нужных пунктов PDF · №100 — чулочно-носочные изделия")

    def edit_catalog(self) -> None:
        dialog = CatalogDialog(self.root, self.catalog, self.catalog_path)
        dialog.show()
        if dialog.result is not None:
            self.catalog = dialog.result
            self.classify()
            self.update_catalog_summary()
            self.redraw(keep_position=True)
            self.export_text.set("Справочник сохранён. Ручные назначения сохранены.")

    def open_statement(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.root, title="Выберите ведомость",
            filetypes=[("Табличные ведомости", ("*.xlsx", "*.xls", "*.xlsm", "*.csv", "*.tsv")), ("Все файлы", "*.*")],
        )
        if filename:
            self.open_path(Path(filename))

    def open_path(self, path: Path) -> None:
        try:
            names = list_sheets(path)
            if not names:
                raise StatementError("В книге нет листов с данными.")
        except StatementError as error:
            messagebox.showerror("Не удалось открыть ведомость", str(error), parent=self.root)
            return
        if not self.can_discard():
            return
        sheet_name = names[0]
        if len(names) > 1:
            chooser = SheetDialog(self.root, names)
            chooser.show()
            if chooser.result is None:
                return
            sheet_name = chooser.result

        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            sheet = read_sheet(path, sheet_name)
            layout = detect_warehouse_layout(sheet)
            if layout is not None:
                result = import_warehouse_products(sheet, layout, _sales_options(layout)[0])
            else:
                if sheet.width < 5:
                    raise StatementError("На листе меньше пяти столбцов с данными. Проверьте файл или выберите другой лист.")
                plan = guess_plan(sheet)
                result = import_products(sheet, plan) if all(column is not None for column in plan.columns) else None
            if result is not None:
                self.display_report(path, sheet_name, result, (sheet, layout) if layout is not None else None)
                return
        except StatementError as error:
            messagebox.showerror("Не удалось открыть ведомость", str(error), parent=self.root)
            return
        finally:
            self.root.configure(cursor="")
        # Ручной выбор нужен только если заголовки не удалось определить однозначно.
        mapping = ColumnMappingDialog(self.root, sheet, plan, sheet_name)
        mapping.show()
        if mapping.result is not None:
            self.display_report(path, sheet_name, mapping.result)

    def classify(self) -> None:
        refresh_groups(self.rows, self.catalog)

    def display_report(
        self, path: Path, sheet_name: str, result: ImportedProducts,
        warehouse_source: tuple[Sheet, WarehouseLayout] | None = None,
    ) -> None:
        self.source_path = path
        self.warehouse_source = warehouse_source
        self.rows = [ReportRow(product) for product in result.products]
        self.classify()
        self.unsaved.clear()
        self.group_filter.set(_ALL_ROWS)
        self.setup_period()
        self.set_report_metadata(result)
        available = warehouses(self.rows)
        self.warehouse_options = (None, *available)
        self.warehouse_selector.configure(values=(ALL_WAREHOUSES, *available), state="readonly" if available else "disabled")
        self.warehouse_selector.current(0)
        self.tabs.select(self.detail_tab)
        name = f"{path.name} · {sheet_name}"
        self.source_text.set(name if len(name) <= 66 else name[:65] + "…")
        self.export_text.set(f"Пропущено: итогов — {result.skipped_totals}, повторов заголовка — {result.skipped_headers}, "
                             f"строк с наименованием без показателей — {result.skipped_no_values}.")
        self.redraw()

    def setup_period(self) -> None:
        if self.warehouse_source is None:
            self.period_fields.grid_remove()
            self.sales_options = ()
            self.month_range = None
            return
        _, layout = self.warehouse_source
        self.sales_options = _sales_options(layout)
        self.sales_selector.configure(values=tuple(_SALES_LABELS[source] for source in self.sales_options))
        self.applied_source = self.sales_options[0]
        labels = tuple(month.label for month in layout.months)
        self.month_range = (0, len(labels) - 1) if labels else None
        for combo in (self.first_month, self.last_month):
            combo.configure(values=labels)
            combo.set("")
        self.sync_period_fields()
        self.period_fields.grid()

    def sync_period_fields(self) -> None:
        self.sales_selector.current(self.sales_options.index(self.applied_source))
        if self.month_range is not None:
            self.first_month.current(self.month_range[0])
            self.last_month.current(self.month_range[1])
        if self.applied_source == SalesQuantity.MONTHS:
            self.month_fields.pack(side="left")
        else:
            self.month_fields.pack_forget()

    def change_period(self, event=None) -> None:
        if self.warehouse_source is None:
            return
        sheet, layout = self.warehouse_source
        source = self.sales_options[self.sales_selector.current()]
        new_range = self.month_range
        if source == SalesQuantity.MONTHS:
            first, last = self.first_month.current(), self.last_month.current()
            if first > last:  # Редактируемая граница всегда имеет приоритет.
                if event is not None and event.widget is self.last_month:
                    first = last
                else:
                    last = first
            new_range = first, last
        if (source, new_range) == (self.applied_source, self.month_range):
            self.sync_period_fields()
            return
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            result = import_warehouse_products(sheet, layout, source, new_range if source == SalesQuantity.MONTHS else None)
        except StatementError as error:
            self.sync_period_fields()
            messagebox.showerror("Не удалось изменить период", str(error), parent=self.root)
        else:
            self.applied_source, self.month_range = source, new_range
            self.sync_period_fields()
            for row, product in zip(self.rows, result.products, strict=True):
                row.product = product
            self.set_report_metadata(result)
            self.redraw(keep_position=True)
            self.export_text.set("Период обновлён. Склад и ручные назначения сохранены.")
        finally:
            self.root.configure(cursor="")

    def set_report_metadata(self, result: ImportedProducts) -> None:
        self.period = result.period
        self.period_text.set(
            f"Реализация: {self.period.sales_label or 'период не указан'} · "
            f"Остатки: {self.period.stock_label or 'дата не указана'}"
        )
        self.column_notes = result.column_notes
        self.method_text.set("  ".join(part for part in (QUANTITY_NOTE, result.summary) if part))

    def open_group_filter(self) -> None:
        self.picker.dismiss()
        menu = self.group_menu
        menu.delete(0, "end")
        for value, label in _GROUP_FILTER_PRESETS.items():
            menu.add_radiobutton(label=label, variable=self.group_filter, value=value, command=self.redraw)
        menu.add_separator()
        for number in GROUPS:
            label = self.catalog.short_label(number)
            menu.add_radiobutton(label=label if len(label) <= 72 else label[:71] + "…",
                                 variable=self.group_filter, value=str(number), command=self.redraw)
        try:
            menu.tk_popup(*self.tree.winfo_pointerxy())
        finally:
            menu.grab_release()

    @staticmethod
    def matches_group_filter(group: int | None, selected: str) -> bool:
        if selected == _ALL_ROWS:
            return True
        if selected == _ASSIGNED_ROWS:
            return group is not None
        if selected == _UNASSIGNED_ROWS:
            return group is None
        return group is not None and str(group) == selected

    def redraw(self, *, keep_position: bool = False) -> None:
        self.picker.dismiss()
        selection = self.tree.selection() if keep_position else ()
        x = self.tree.xview()[0] if keep_position else 0
        y = self.tree.yview()[0] if keep_position else 0
        self.tree.delete(*self.tree.get_children())
        visible = 0
        selected_group = self.group_filter.get()
        heading = RESULT_COLUMNS[0] if selected_group == _ALL_ROWS else _GROUP_FILTER_PRESETS.get(
            selected_group, f"Группа №{selected_group}",
        )
        self.tree.heading("group", text=f"{heading} ▾")
        warehouse = self.active_warehouse()
        for index, row in enumerate(self.rows):
            if not in_warehouse(row, warehouse) or not self.matches_group_filter(row.group, selected_group):
                continue
            self.tree.insert("", "end", iid=str(index), values=(
                self.catalog.short_label(row.group), row.product.name,
                display_quantity(row.product.sold_quantity, row.product.unit), display_number(row.product.sold_amount),
                display_quantity(row.product.ending_quantity, row.product.unit), display_number(row.product.ending_amount),
            ), tags=("odd",) if visible % 2 else ())
            visible += 1
        if keep_position:
            self.tree.xview_moveto(x)
            self.tree.yview_moveto(y)
            if selection and self.tree.exists(selection[0]):
                self.tree.selection_set(selection[0])
                self.tree.focus(selection[0])
        selected_rows = self.warehouse_rows()
        self.refresh_counters(selected_rows)
        self.refresh_summary(selected_rows)
        self.refresh_selection()

    def refresh_counters(self, rows: list[ReportRow]) -> None:
        unassigned = sum(row.group is None for row in rows)
        self.total_text.set(f"Строки: {len(rows)}")
        self.assigned_text.set(f"С группой: {len(rows) - unassigned}")
        self.unassigned_text.set(f"Без группы: {unassigned}")
        if self.tree.get_children():
            self.empty.place_forget()
        else:
            self.empty.configure(text=("В этом фильтре нет строк." if self.rows else
                                       "Откройте ведомость, чтобы начать.\nСправочник 4-торг уже встроен в программу."))
            self.empty.place(relx=0.5, rely=0.48, anchor="center")
        self.export_button.state(["!disabled"] if rows else ["disabled"])

    def selected_index(self) -> int | None:
        selection = self.tree.selection()
        return int(selection[0]) if selection else None

    def refresh_selection(self, _event=None) -> None:
        showing_detail = self.tabs.select() == str(self.detail_tab)
        if not showing_detail:
            self.picker.dismiss()
            self.edit_button.state(["disabled"])
            self.selection_text.set("Сводка включает все строки выбранного склада. Фильтр группы в шапке действует только на просмотр.")
            return
        index = self.selected_index()
        self.edit_button.state(["disabled"] if index is None else ["!disabled"])
        if index is None:
            text = "Группа ▾ в шапке — фильтр. Нажмите на ячейку первого столбца или Enter, чтобы изменить группу строки."
        else:
            row = self.rows[index]
            group = self.catalog.label(row.group) or "группа не назначена"
            text = f"{row.product.name}  ·  {group}"
            if row.manual:
                text += "  (вручную)"
            if len(text) > 360:
                text = text[:359] + "…"
        self.selection_text.set(text)

    def _click_cell(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell" or self.tree.identify_column(event.x) != "#1":
            return
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.edit_selected()

    def edit_selected(self, _event=None) -> str:
        index = self.selected_index()
        if index is not None:
            self.picker.open(str(index), self.rows[index].group, self.catalog)
        return "break"

    def assign_group(self, item: str, group: int | None) -> None:
        index = int(item)
        row = self.rows[index]
        if row.manual and row.group == group:
            return
        row.assign(group)
        self.unsaved.add(index)
        if not self.matches_group_filter(row.group, self.group_filter.get()):
            next_item = self.tree.next(item) or self.tree.prev(item)
            self.tree.delete(item)
            if next_item:
                self.tree.selection_set(next_item)
                self.tree.focus(next_item)
                self.tree.see(next_item)
        else:
            self.tree.set(item, "group", self.catalog.short_label(row.group))
        selected_rows = self.warehouse_rows()
        self.refresh_counters(selected_rows)
        self.refresh_summary(selected_rows)
        self.refresh_selection()
        self.export_text.set("Назначение изменено. Чтобы сохранить результат, нажмите «Сохранить Excel…».")

    def save_report(self) -> None:
        warehouse = self.active_warehouse()
        if not self.warehouse_rows() or self.source_path is None:
            return
        filename = filedialog.asksaveasfilename(
            parent=self.root, title="Сохранить результат", initialdir=str(self.source_path.parent),
            initialfile=f"{self.source_path.stem} — результат.xlsx", defaultextension=".xlsx",
            filetypes=[("Книга Excel", "*.xlsx")],
        )
        if not filename:
            return
        path = Path(filename)
        if path.resolve() == self.source_path.resolve():
            messagebox.showerror("Выберите другое имя", "Результат не должен заменять исходную ведомость.", parent=self.root)
            return
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            export_xlsx(path, self.rows, self.catalog, self.column_notes, period=self.period, warehouse=warehouse)
        except (OSError, ValueError) as error:
            messagebox.showerror("Не удалось сохранить Excel", str(error), parent=self.root)
        else:
            self.unsaved.difference_update(index for index, row in enumerate(self.rows) if in_warehouse(row, warehouse))
            count = sum(in_warehouse(row, warehouse) for row in self.rows)
            self.export_text.set(f"Сохранено: {path.name} · строк: {count} · {warehouse or ALL_WAREHOUSES}")
        finally:
            self.root.configure(cursor="")

    def can_discard(self) -> bool:
        return not self.unsaved or messagebox.askyesno(
            "Несохранённые назначения", "Ручные назначения не экспортированы. Продолжить без сохранения?", parent=self.root,
        )

    def close(self) -> None:
        if self.can_discard():
            self.root.destroy()


def _self_test(marker: Path) -> int:
    """Проверка запуска собранного EXE, Tk и упакованных данных на Windows-раннере."""
    from unittest.mock import patch

    from openpyxl import Workbook, load_workbook

    from statistika.catalog import RecognitionRule
    from statistika.classifier import Classifier

    def unexpected_dialog(title, message, **_kwargs):
        raise AssertionError(f"{title}: {message}")

    try:
        with (tempfile.TemporaryDirectory() as directory,
              patch.object(messagebox, "showerror", side_effect=unexpected_dialog),
              patch.object(messagebox, "showwarning", side_effect=unexpected_dialog)):
            base = Path(directory)
            catalog_path = base / "catalog.json"
            load_catalog(catalog_path)  # При ошибке ресурса не показывать модальное окно в CI.
            statement = base / "Ведомость.csv"
            statement.write_text(
                "Товар;Продано, шт.;Сумма продаж;Конечный остаток, шт.;Конечный остаток, сумма\n"
                "Носки мужские;2;5,20;3;7,80\n"
                "Неизвестный товар;1;2;0;0\n",
                encoding="utf-8",
            )
            root = tk.Tk()
            root.withdraw()
            try:
                app = MainWindow(root, catalog_path)
                errors: list[str] = []
                root.report_callback_exception = lambda *exc: errors.append("".join(traceback.format_exception(*exc)))
                root.deiconify()
                root.update()
                catalog_editor = CatalogDialog(root, app.catalog, catalog_path)
                catalog_editor.deiconify()
                root.update()
                assert tuple(catalog_editor.tabs.tab(tab_id, "text") for tab_id in catalog_editor.tabs.tabs()) == (
                    "Пункты PDF", "Правила распознавания",
                )
                catalog_editor.tabs.select(catalog_editor.points_tab)
                root.update()
                catalog_editor.points_text.insert("end", "\n75. Дополнительный пункт CI.")
                catalog_editor.tabs.select(catalog_editor.rules_tab)
                root.update()
                editor = catalog_editor.rule_editor
                assert editor.tree["columns"] == ("enabled", "phrase", "exclude", "delete")
                editor.query.set("ГЕЛЬ")
                assert {"57", "63"} <= {item.split(":")[0] for item in editor.tree.get_children()}
                catalog_editor.rule_group.current(1 + GROUPS.index(57))
                catalog_editor.rule_group.event_generate("<<ComboboxSelected>>")
                root.update()
                assert {item.split(":")[0] for item in editor.tree.get_children() if ":" in item} == {"57"}
                catalog_editor.rule_group.current(0)
                catalog_editor.rule_group.event_generate("<<ComboboxSelected>>")
                root.update()

                def click_phrase(item: str, column: str) -> None:
                    editor.tree.see(item)
                    root.update()
                    x, y, width, height = editor.tree.bbox(item, column)
                    editor.tree.event_generate("<ButtonRelease-1>", x=x + width // 2, y=y + height // 2)
                    root.update()

                original_rule = RecognitionRule("шпатель", at_start=True, also=("фасадный",))
                index = app.catalog.recognition_rules(53).index(original_rule)
                rule_item = f"53:{index}"
                editor.query.set("ШПАТЕЛЬ ФАСАДНЫЙ")
                assert editor.tree.set(rule_item, "enabled") == "☑"
                click_phrase(rule_item, "enabled")
                assert editor.tree.set(rule_item, "enabled") == "☐" and editor.rules[53][index].disabled

                editor.group.current(1 + GROUPS.index(53))
                editor.phrase.insert(0, "Пробная фраза CI")
                editor.include_button.invoke()
                enabled_phrase = RecognitionRule("Пробная фраза CI")
                exclusion = RecognitionRule("Пробная фраза CI", exclude=True)
                added = f"53:{len(editor.rules[53]) - 1}"
                assert editor.rules[53][-1] == enabled_phrase and editor.tree.set(added, "exclude") == "Исключение"
                click_phrase(added, "exclude")
                assert editor.rules[53][-1] == exclusion and editor.tree.set(added, "exclude") == "Снять исключение"
                assert editor.tree.set(added, "enabled") == "☑"
                click_phrase(added, "exclude")
                assert editor.rules[53][-1] == enabled_phrase
                click_phrase(added, "exclude")
                editor.phrase.insert(0, "Временная фраза CI")
                editor.include_button.invoke()
                temporary = f"53:{len(editor.rules[53]) - 1}"
                assert editor.tree.set(temporary, "delete") == "Удалить"
                click_phrase(temporary, "delete")
                assert not editor.tree.exists(temporary) and editor.rules[53][-1] == exclusion
                catalog_editor.save_button.invoke()
                saved = load_catalog(catalog_path)
                assert saved == catalog_editor.result
                assert saved.points == {**app.catalog.points, 75: "Дополнительный пункт CI."} and saved.source == app.catalog.source
                assert saved.rules[53][index] == RecognitionRule("шпатель", at_start=True, also=("фасадный",), disabled=True)
                assert saved.rules[53][-1] == exclusion
                classifier = Classifier(saved)
                assert classifier.group_for("Муфта PPR 40 мм") == 53
                assert classifier.group_for("Муфта PPR 40 мм — пробная фраза CI") is None

                assert len(app.catalog.points) == 74 and not app.catalog.missing
                with patch.object(ColumnMappingDialog, "show", side_effect=AssertionError("Лишнее окно сопоставления")):
                    app.open_path(statement)
                assert not app.period_fields.winfo_manager()
                root.update()
                assert [row.group for row in app.rows] == [100, None]
                assert len(app.tree.get_children()) == 2
                for column, unit in (("sold_qty", "шт."), ("sold_amount", "руб."), ("stock_qty", "шт."), ("stock_amount", "руб.")):
                    assert unit in app.tree.heading(column, "text")

                def choose_group_filter(value: str) -> None:
                    with patch.object(app.group_menu, "tk_popup") as popup:
                        root.tk.call(app.tree.heading("group", "command"))
                        popup.assert_called_once()
                    option = next(index for index in range(app.group_menu.index("end") + 1)
                                  if app.group_menu.type(index) == "radiobutton" and app.group_menu.entrycget(index, "value") == value)
                    app.group_menu.invoke(option)
                    root.update()

                choose_group_filter(_UNASSIGNED_ROWS)
                assert app.tree.get_children() == ("1",)
                choose_group_filter(_ALL_ROWS)
                assert app.tree.get_children() == ("0", "1") and app.tree.heading("group", "text").endswith("▾")
                x, y, width, height = app.tree.bbox("1", "#1")
                app.tree.event_generate("<ButtonRelease-1>", x=x + width // 2, y=y + height // 2)
                root.update()
                assert app.picker.item == "1" and app.picker.options.size() == len(GROUPS) + 1
                assert app.picker.options.curselection() == (0,)
                app.picker.options.selection_clear(0, "end")
                app.picker.options.selection_set(1)
                app.picker.choose()
                assert app.rows[1].manual and app.rows[1].group == GROUPS[0]
                app.edit_selected()
                assert app.picker.item == "1"
                app.picker.options.selection_clear(0, "end")
                app.picker.options.selection_set(0)
                app.picker.choose()
                assert app.rows[1].manual and app.rows[1].group is None
                export_xlsx(base / "result.xlsx", app.rows, app.catalog)
                assert (base / "result.xlsx").is_file()

                # Нераспознанный файл всё ещё можно сопоставить вручную, а книга с несколькими листами предлагает только выбор листа.
                manual_path = base / "Без заголовка.xlsx"
                workbook = Workbook()
                workbook.active.title = "Обложка"
                workbook.create_sheet("Данные").append(["Гольфы", 2, 3, 4, 5])
                workbook.save(manual_path)
                workbook.close()

                def choose_second(dialog):
                    dialog.sheets.current(1)
                    dialog.submit()

                def map_columns(dialog):
                    assert not hasattr(dialog, "preview") and dialog.first.get() == "0"
                    for column, combo in enumerate(dialog.columns, 1):
                        combo.current(column)
                    dialog.submit()

                with patch.object(SheetDialog, "show", choose_second), patch.object(ColumnMappingDialog, "show", map_columns), \
                        patch.object(app, "can_discard", return_value=True):
                    app.open_path(manual_path)
                assert len(app.rows) == 1 and app.rows[0].group == 100
                assert app.rows[0].product.metrics == (2, 3, 4, 5)

                warehouse = base / "Остатки ТМЦ.csv"
                warehouse.write_text(
                    "С 01.10.26 по 15.11.26\n"
                    "Склад;Группа номенклатуры;Номенклатура;Ед.;Начальный остаток;Приход;Расход;"
                    "Конечный остаток;Розничная цена;Октябрь 2026 г.;Ноябрь 2026 г.\n"
                    "Мир;АРХИВ;Американка 1/2 прямая Kalde;шт.;2;3;2;3;4,80;4;1\n"
                    "Кореличи;Абразивы;Американка 1/2 прямая Kalde;шт.;5;4;1;8;12;6;3\n"
                    "Мир;Электрика;Автомат ВА 47-29 1п 16А IEK;шт.;1;2;4;5;2,50;0;2\n"
                    "Мир;Клеи;Клей Бустилат 0,9кг;банка;1;5;3;4;3;2;3\n",
                    encoding="utf-8",
                )
                with patch.object(ColumnMappingDialog, "show", side_effect=AssertionError("Лишнее окно импорта")):
                    app.open_path(warehouse)
                root.update()
                assert [row.group for row in app.rows] == [53, 53, 53, 53]
                assert app.applied_source == SalesQuantity.MONTHS and app.period.sales_months == ("Октябрь 2026", "Ноябрь 2026")
                assert app.rows[0].product.metrics == (5, 24, 3, Decimal("14.4"))
                assert app.period_fields.winfo_ismapped() and app.month_fields.winfo_ismapped()
                app.warehouse_selector.current(app.warehouse_options.index("Мир"))
                app.warehouse_selector.event_generate("<<ComboboxSelected>>")
                root.update()
                assert app.tree.get_children() == ("0", "2", "3")
                assert app.summary_tree.set("53", "lines") == "3"
                assert app.summary_tree.set("53", "sold_qty") == "12"
                choose_group_filter("53")
                assert "№53" in app.tree.heading("group", "text")
                assert app.tree.set("2", "group") == app.catalog.short_label(53)
                app.assign_group("0", 100)
                assert app.tree.get_children() == ("2", "3")
                assert [row.group for row in app.rows] == [100, 53, 53, 53]
                assert app.unsaved == {0}

                # Изменение периода не читает файл заново и не теряет склад, фильтр и ручные назначения.
                with patch("statistika.app.read_sheet", side_effect=AssertionError("Повторное чтение файла")):
                    app.last_month.current(0)
                    app.last_month.event_generate("<<ComboboxSelected>>")
                    root.update()
                    assert app.period.sales_months == ("Октябрь 2026",)
                    assert app.rows[0].product.metrics == (4, Decimal("19.2"), None, None)
                    assert app.rows[2].product.metrics == (0, 0, None, None)
                    # Конец раньше нового начала: обе границы становятся ноябрём.
                    app.first_month.current(1)
                    app.first_month.event_generate("<<ComboboxSelected>>")
                    root.update()
                assert app.last_month.current() == 1
                assert app.period.sales_months == ("Ноябрь 2026",) and app.period.stock_label == "15.11.2026"
                assert app.rows[0].product.metrics == (1, Decimal("4.8"), 3, Decimal("14.4"))
                assert app.active_warehouse() == "Мир" and app.group_filter.get() == "53"
                assert app.rows[0].manual and app.rows[0].group == 100 and app.unsaved == {0}
                assert app.tree.get_children() == ("2", "3") and app.tree.set("2", "sold_amount") == "5,00"
                assert app.summary_tree.set("53", "sold_qty") == "5"  # 2 шт. + 3 банки
                assert app.summary_tree.set("100", "sold_qty") == "1"

                app.sales_selector.current(app.sales_options.index(SalesQuantity.OUTFLOW))
                app.change_period()
                assert app.rows[0].product.metrics[:2] == (2, Decimal("9.6")) and not app.month_fields.winfo_manager()
                app.sales_selector.current(app.sales_options.index(SalesQuantity.MONTHS))
                app.change_period()
                assert app.period.sales_months == ("Ноябрь 2026",)

                target = base / "warehouse-result.xlsx"
                with patch.object(filedialog, "asksaveasfilename", return_value=str(target)):
                    app.save_report()
                workbook = load_workbook(target)
                try:
                    detail, summary = workbook.worksheets
                    assert detail.max_row == 4 and detail.max_column == 6
                    assert tuple(cell.value for cell in detail[2][2:]) == (1, 4.8, 3, 14.4)
                    assert detail["A3"].value == app.catalog.label(53)
                    assert detail.auto_filter.ref == "A1:F4"
                    assert tuple(cell.value for cell in summary[6 + GROUPS.index(53)][2:]) == (5, 14, 9, 24.5)
                finally:
                    workbook.close()
                assert not app.unsaved
                assert not errors, "\n".join(errors)
            finally:
                root.destroy()
        marker.write_text("OK\n", encoding="utf-8")
        return 0
    except Exception:
        marker.write_text(traceback.format_exc(), encoding="utf-8")
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Статистика — продажи и остатки")
    parser.add_argument("--self-test", type=Path, metavar="RESULT_FILE", help="проверить запуск и сохранить результат")
    args = parser.parse_args()
    if args.self_test is not None:
        return _self_test(args.self_test)
    root = tk.Tk()
    MainWindow(root, local_catalog_path())
    root.mainloop()
    return 0
