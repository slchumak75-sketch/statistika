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
from statistika.dialogs import CatalogDialog, StatementDialog
from statistika.inline_picker import InlineGroupPicker
from statistika.report import (
    ALL_WAREHOUSES, QUANTITY_NOTE, SUMMARY_NOTE, ReportRow, display_number, display_quantity, display_total,
    display_unit_totals, export_xlsx, in_warehouse, refresh_groups, summarize_groups, unit_totals, warehouses,
)
from statistika.statement import ImportedProducts, ReportPeriod, StatementError, list_sheets

_BG = "#f1f5fa"
_INK = "#16304c"
_MUTED = "#466078"
_BLUE = "#174f91"
_VIEWS = ("Все строки", "Только с группой", "Только без группы")


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
        self.warehouse_options: tuple[str | None, ...] = (None,)
        self.view = tk.StringVar(value=_VIEWS[0])
        self.period_text = tk.StringVar()
        self.catalog_summary = tk.StringVar()
        self.source_text = tk.StringVar(value="Ведомость ещё не открыта")
        self.total_text = tk.StringVar(value="Строки: 0")
        self.assigned_text = tk.StringVar(value="С группой: 0")
        self.unassigned_text = tk.StringVar(value="Без группы: 0")
        self.selection_text = tk.StringVar(value="Выбор группы: нажмите на первый столбец или клавишу Enter на выбранной строке.")
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
        body.rowconfigure(4, weight=1)
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

        filters = ttk.Frame(body)
        filters.grid(row=2, sticky="ew", pady=(0, 8))
        ttk.Label(filters, text="Склад").pack(side="left")
        self.warehouse_selector = ttk.Combobox(filters, state="disabled", values=(ALL_WAREHOUSES,), width=23)
        self.warehouse_selector.set(ALL_WAREHOUSES)
        self.warehouse_selector.pack(side="left", padx=(6, 16))
        self.warehouse_selector.bind("<<ComboboxSelected>>", lambda _event: self.redraw())
        self.view_selector = ttk.Combobox(filters, state="readonly", textvariable=self.view, values=_VIEWS, width=20)
        self.view_selector.pack(side="left")
        self.view_selector.bind("<<ComboboxSelected>>", lambda _event: self.redraw())
        ttk.Label(filters, textvariable=self.source_text, style="Muted.TLabel").pack(side="right", padx=(10, 0))

        stats = ttk.Frame(body)
        stats.grid(row=3, sticky="ew", pady=(0, 8))
        for variable in (self.total_text, self.assigned_text, self.unassigned_text):
            ttk.Label(stats, textvariable=variable, style="Counter.TLabel").pack(side="left", padx=(0, 7))
        period_label = ttk.Label(stats, textvariable=self.period_text, style="Muted.TLabel", wraplength=500)
        period_label.pack(side="left", fill="x", expand=True, padx=4)
        period_label.bind("<Configure>", lambda event: period_label.configure(wraplength=max(210, event.width - 2)))

        self.tabs = ttk.Notebook(body)
        self.tabs.grid(row=4, sticky="nsew")
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
        headers = ("№ пункта и описание", "Наименование товара", "Продано за период\nКоличество, шт.*",
                   "Продано за период\nСумма, руб.", "Остаток на конец\nКоличество, шт.*", "Остаток на конец\nСумма, руб.")
        for index, (name, title, width) in enumerate(zip(names, headers, (268, 268, 131, 131, 136, 136))):
            self.tree.heading(name, text=title, anchor="w" if index < 2 else "center")
            self.tree.column(name, width=width, minwidth=width, anchor="w" if index < 2 else "e", stretch=index < 2)
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
        detail.grid(row=5, sticky="ew", pady=(10, 2))
        detail.bind("<Configure>", lambda event: detail.configure(wraplength=max(350, event.width - 2)))
        ttk.Label(body, textvariable=self.export_text, style="Muted.TLabel").grid(row=6, sticky="w")
        method = ttk.Label(body, textvariable=self.method_text, wraplength=810, style="Muted.TLabel")
        method.grid(row=7, sticky="ew", pady=(2, 0))
        method.bind("<Configure>", lambda event: method.configure(wraplength=max(350, event.width - 2)))

    def _draw_summary(self) -> None:
        tab = self.summary_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        columns = ("#0", "lines", "sold_qty", "sold_amount", "stock_qty", "stock_amount")
        self.summary_tree = ttk.Treeview(tab, columns=columns[1:], show="tree headings", selectmode="browse")
        self.summary_tree.grid(row=0, column=0, sticky="nsew")
        headers = ("Группа / единица", "Строк", "Продано, по ед.", "Продажи, руб.", "Остаток, по ед.", "Остаток, руб.")
        for index, (column, label, width) in enumerate(zip(columns, headers, (294, 63, 142, 125, 142, 125))):
            self.summary_tree.column(column, width=width, minwidth=width, anchor="w" if not index else "e", stretch=not index)
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
        expanded = {item for item in tree.get_children() if tree.item(item, "open")}
        tree.delete(*tree.get_children())
        for total in summarize_groups(rows) if self.rows else ():
            item = str(total.group)
            label = self.catalog.short_label(total.group) or "Без группы"
            quantities = []
            for is_sale in (True, False):
                by_unit = unit_totals(total, sold=is_sale)
                quantities.append(
                    "Раскройте группу" if len(by_unit) > 1 and any(value.known for _, value in by_unit)
                    else display_unit_totals(by_unit)
                )
            tree.insert("", "end", iid=item, text=label, open=item in expanded, values=(
                total.lines, quantities[0], display_total(total.sold_amount), quantities[1], display_total(total.ending_amount),
            ))
            if len(total.units) > 1:
                for unit in total.units:
                    tree.insert(item, "end", text=unit.unit or "Ед. не указана", values=(
                        unit.sold.lines, display_total(unit.sold), "", display_total(unit.ending), "",
                    ))

    def update_catalog_summary(self) -> None:
        ready = len(PDF_GROUPS) - len(self.catalog.missing)
        self.catalog_summary.set(f"Справочник: {ready} из {len(PDF_GROUPS)} нужных пунктов PDF · №100 — чулочно-носочные изделия")

    def edit_catalog(self) -> None:
        dialog = CatalogDialog(self.root, self.catalog, self.catalog_path)
        dialog.show()
        if dialog.result is not None:
            self.catalog = dialog.result
            refresh_groups(self.rows, self.catalog)
            self.update_catalog_summary()
            self.redraw()
            self.export_text.set(f"Справочник сохранён. Пунктов PDF в памяти: {len(self.catalog.points)}. Ручные назначения сохранены.")

    def open_statement(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.root, title="Выберите ведомость",
            filetypes=[("Табличные ведомости", ("*.xlsx", "*.xls", "*.xlsm", "*.csv", "*.tsv")), ("Все файлы", "*.*")],
        )
        if not filename:
            return
        path = Path(filename)
        try:
            names = list_sheets(path)
            if not names:
                raise StatementError("В книге нет листов с данными.")
        except StatementError as error:
            messagebox.showerror("Не удалось открыть ведомость", str(error), parent=self.root)
            return
        if not self.can_discard():
            return
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            dialog = StatementDialog(self.root, path, names)
        finally:
            self.root.configure(cursor="")
        dialog.show()
        if dialog.result is not None:
            self.display_report(path, dialog.sheet_name, dialog.result)

    def display_report(self, path: Path, sheet_name: str, result: ImportedProducts) -> None:
        self.source_path = path
        self.rows = [ReportRow(product) for product in result.products]
        refresh_groups(self.rows, self.catalog)
        self.unsaved.clear()
        self.view.set(_VIEWS[0])
        self.period = result.period
        self.period_text.set(
            f"Реализация: {self.period.sales_label or 'период не указан'} · "
            f"Остатки: {self.period.stock_label or 'дата не указана'}"
        )
        self.column_notes = result.column_notes
        self.method_text.set("  ".join(part for part in (QUANTITY_NOTE, result.summary) if part))
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

    def redraw(self) -> None:
        self.picker.dismiss()
        self.tree.delete(*self.tree.get_children())
        visible = 0
        view = self.view.get()
        warehouse = self.active_warehouse()
        for index, row in enumerate(self.rows):
            if not in_warehouse(row, warehouse) or not self.visible_in_view(row, view):
                continue
            self.tree.insert("", "end", iid=str(index), values=(
                self.catalog.short_label(row.group), row.product.name,
                display_quantity(row.product.sold_quantity, row.product.unit), display_number(row.product.sold_amount),
                display_quantity(row.product.ending_quantity, row.product.unit), display_number(row.product.ending_amount),
            ), tags=("odd",) if visible % 2 else ())
            visible += 1
        selected_rows = self.warehouse_rows()
        self.refresh_counters(selected_rows)
        self.refresh_summary(selected_rows)
        self.refresh_selection()

    @staticmethod
    def visible_in_view(row: ReportRow, view: str) -> bool:
        return view == _VIEWS[0] or (view == _VIEWS[1] and row.group is not None) or (view == _VIEWS[2] and row.group is None)

    def refresh_counters(self, rows: list[ReportRow]) -> None:
        unassigned = sum(row.group is None for row in rows)
        self.total_text.set(f"Строки: {len(rows)}")
        self.assigned_text.set(f"С группой: {len(rows) - unassigned}")
        self.unassigned_text.set(f"Без группы: {unassigned}")
        view = self.view.get()
        available = (len(rows), len(rows) - unassigned, unassigned)[_VIEWS.index(view)]
        if available:
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
        self.view_selector.configure(state="readonly" if showing_detail else "disabled")
        if not showing_detail:
            self.picker.dismiss()
            self.edit_button.state(["disabled"])
            self.selection_text.set("Сводка включает все строки выбранного склада, независимо от фильтра назначений.")
            return
        index = self.selected_index()
        self.edit_button.state(["disabled"] if index is None else ["!disabled"])
        if index is None:
            text = "Нажмите на первый столбец или клавишу Enter на выбранной строке, чтобы изменить группу."
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
        if not self.visible_in_view(row, self.view.get()):
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
    try:
        with tempfile.TemporaryDirectory() as directory:
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
                catalog_editor = CatalogDialog(root, app.catalog, catalog_path)
                for index, number in enumerate(GROUPS):
                    catalog_editor.rule_group.current(index)
                    catalog_editor._switch_rule()
                    assert catalog_editor.rules_text.get("1.0", "end-1c") == "\n".join(app.catalog.recognition_phrases(number))
                    assert catalog_editor.shown_rule
                catalog_editor.rule_group.current(GROUPS.index(53))
                catalog_editor._switch_rule()
                assert "^антисептик\n" in catalog_editor.rules_text.get("1.0", "end-1c")
                catalog_editor.rules_text.insert("end", "\n^Проверка фразы CI")
                catalog_editor.save()
                assert load_catalog(catalog_path).rules[53][-1] == "^Проверка фразы CI"

                mapper = StatementDialog(root, statement, list_sheets(statement))
                assert [combo.current() for combo in mapper.columns] == [1, 2, 3, 4, 5]
                mapper.submit()
                assert mapper.result is not None
                assert len(app.catalog.points) == 74 and not app.catalog.missing
                app.display_report(statement, mapper.sheet_name, mapper.result)
                root.deiconify()
                root.update()
                assert [row.group for row in app.rows] == [100, None]
                assert len(app.tree.get_children()) == 2
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

                warehouse = base / "Остатки ТМЦ.csv"
                warehouse.write_text(
                    "С 01.10.26 по 15.11.26\n"
                    "Склад;Группа номенклатуры;Номенклатура;Ед.;Начальный остаток;Приход;Расход;"
                    "Конечный остаток;Розничная цена;Октябрь 2026 г.;Ноябрь 2026 г.\n"
                    "Мир;АРХИВ;Антисептик Белмастер Экобио 5л.;шт.;2;3;2;3;4,80;4;1\n"
                    "Кореличи;Антисептики;Антисептик Белмастер Экобио 5л.;шт.;5;4;1;8;12;6;3\n",
                    encoding="utf-8",
                )
                mapper = StatementDialog(root, warehouse, list_sheets(warehouse))
                assert mapper.use_warehouse.get() and mapper.sales.current() == 0
                mapper.sales.current(1)
                mapper._sync_sales()
                assert mapper.month_fields.winfo_manager() == "pack"
                mapper.first_month.current(1)
                mapper.submit()
                assert mapper.result is not None and mapper.result.column_notes
                assert mapper.result.period.sales_months == ("Ноябрь 2026",)
                assert mapper.result.products[0].metrics == (1, None, 3, Decimal("14.4"))
                app.display_report(warehouse, mapper.sheet_name, mapper.result)
                assert [row.group for row in app.rows] == [53, 53]
                app.warehouse_selector.current(app.warehouse_options.index("Мир"))
                app.warehouse_selector.event_generate("<<ComboboxSelected>>")
                root.update()
                assert app.tree.get_children() == ("0",)
                assert app.summary_tree.set("53", "lines") == "1"
                app.view.set(_VIEWS[2])
                app.redraw()
                assert not app.tree.get_children()
                assert app.summary_tree.set("53", "lines") == "1"
                export_xlsx(base / "warehouse-result.xlsx", app.rows, app.catalog, app.column_notes,
                            period=app.period, warehouse=app.active_warehouse())
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
