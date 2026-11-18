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
from statistika.dialogs import CatalogDialog, GroupDialog, StatementDialog
from statistika.report import ReportRow, display_number, export_xlsx, refresh_groups
from statistika.statement import ImportedProducts, StatementError, list_sheets

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
        self.dirty = False
        self.view = tk.StringVar(value=_VIEWS[0])
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
        body.rowconfigure(3, weight=1)
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
        view_selector = ttk.Combobox(actions, state="readonly", textvariable=self.view, values=_VIEWS, width=20)
        view_selector.pack(side="right")
        view_selector.bind("<<ComboboxSelected>>", lambda _event: self.redraw())

        stats = ttk.Frame(body)
        stats.grid(row=2, sticky="ew", pady=(0, 8))
        for variable in (self.total_text, self.assigned_text, self.unassigned_text):
            ttk.Label(stats, textvariable=variable, style="Counter.TLabel").pack(side="left", padx=(0, 7))
        ttk.Label(stats, textvariable=self.source_text, style="Muted.TLabel").pack(side="left", padx=4)

        table = ttk.Frame(body)
        table.grid(row=3, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        names = ("group", "name", "sold_qty", "sold_amount", "stock_qty", "stock_amount")
        self.tree = ttk.Treeview(table, columns=names, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        headers = ("№ пункта и описание", "Наименование товара", "Продано за период\nКоличество",
                   "Продано за период\nСумма", "Остаток на конец\nКоличество", "Остаток на конец\nСумма")
        for index, (name, title, width) in enumerate(zip(names, headers, (268, 268, 131, 131, 136, 136))):
            self.tree.heading(name, text=title, anchor="w" if index < 2 else "center")
            self.tree.column(name, width=width, minwidth=width, anchor="w" if index < 2 else "e", stretch=index < 2)
        sy = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.tag_configure("odd", background="#f1f6fc")
        self.tree.bind("<<TreeviewSelect>>", self.refresh_selection)
        self.tree.bind("<Return>", lambda _event: self.edit_selected())
        self.tree.bind("<ButtonRelease-1>", self._click_cell)
        self.empty = ttk.Label(table, text="", style="Empty.TLabel", justify="center")

        detail = ttk.Label(body, textvariable=self.selection_text, wraplength=810, style="Muted.TLabel")
        detail.grid(row=4, sticky="ew", pady=(10, 2))
        detail.bind("<Configure>", lambda event: detail.configure(wraplength=max(350, event.width - 2)))
        ttk.Label(body, textvariable=self.export_text, style="Muted.TLabel").grid(row=5, sticky="w")
        method = ttk.Label(body, textvariable=self.method_text, wraplength=810, style="Muted.TLabel")
        method.grid(row=6, sticky="ew", pady=(2, 0))
        method.bind("<Configure>", lambda event: method.configure(wraplength=max(350, event.width - 2)))

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
        self.dirty = False
        self.view.set(_VIEWS[0])
        self.column_notes = result.column_notes
        self.method_text.set(result.summary)
        name = f"{path.name} · {sheet_name}"
        self.source_text.set(name if len(name) <= 66 else name[:65] + "…")
        self.export_text.set(f"Пропущено: итогов — {result.skipped_totals}, повторов заголовка — {result.skipped_headers}, "
                             f"строк с наименованием без показателей — {result.skipped_no_values}.")
        self.redraw()

    def redraw(self) -> None:
        self.tree.delete(*self.tree.get_children())
        visible = 0
        view = self.view.get()
        for index, row in enumerate(self.rows):
            if not self.visible_in_view(row, view):
                continue
            self.tree.insert("", "end", iid=str(index), values=(
                self.catalog.short_label(row.group), row.product.name, *(display_number(value) for value in row.product.metrics),
            ), tags=("odd",) if visible % 2 else ())
            visible += 1
        self.refresh_counters()
        self.refresh_selection()

    @staticmethod
    def visible_in_view(row: ReportRow, view: str) -> bool:
        return view == _VIEWS[0] or (view == _VIEWS[1] and row.group is not None) or (view == _VIEWS[2] and row.group is None)

    def refresh_counters(self) -> None:
        unassigned = sum(row.group is None for row in self.rows)
        self.total_text.set(f"Строки: {len(self.rows)}")
        self.assigned_text.set(f"С группой: {len(self.rows) - unassigned}")
        self.unassigned_text.set(f"Без группы: {unassigned}")
        view = self.view.get()
        available = (len(self.rows), len(self.rows) - unassigned, unassigned)[_VIEWS.index(view)]
        if available:
            self.empty.place_forget()
        else:
            self.empty.configure(text=("В этом фильтре нет строк." if self.rows else
                                       "Откройте ведомость, чтобы начать.\nСправочник 4-торг уже встроен в программу."))
            self.empty.place(relx=0.5, rely=0.48, anchor="center")
        self.export_button.state(["!disabled"] if self.rows else ["disabled"])

    def selected_index(self) -> int | None:
        selection = self.tree.selection()
        return int(selection[0]) if selection else None

    def refresh_selection(self, _event=None) -> None:
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

    def edit_selected(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        row = self.rows[index]
        dialog = GroupDialog(self.root, self.catalog, row.product.name, row.group)
        dialog.show()
        if not dialog.applied or (row.manual and row.group == dialog.number):
            return
        row.assign(dialog.number)
        self.dirty = True
        item = str(index)
        if not self.visible_in_view(row, self.view.get()):
            next_item = self.tree.next(item) or self.tree.prev(item)
            self.tree.delete(item)
            if next_item:
                self.tree.selection_set(next_item)
                self.tree.focus(next_item)
                self.tree.see(next_item)
        else:
            self.tree.set(item, "group", self.catalog.short_label(row.group))
        self.refresh_counters()
        self.refresh_selection()
        self.tree.focus_set()
        self.export_text.set("Назначение изменено. Чтобы сохранить результат, нажмите «Сохранить Excel…».")

    def save_report(self) -> None:
        if not self.rows or self.source_path is None:
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
            export_xlsx(path, self.rows, self.catalog, self.column_notes)
        except (OSError, ValueError) as error:
            messagebox.showerror("Не удалось сохранить Excel", str(error), parent=self.root)
        else:
            self.dirty = False
            self.export_text.set(f"Сохранено: {path.name} · строк: {len(self.rows)} (включая строки без группы)")
        finally:
            self.root.configure(cursor="")

    def can_discard(self) -> bool:
        return not self.dirty or messagebox.askyesno(
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
                catalog_editor = CatalogDialog(root, app.catalog, catalog_path)
                catalog_editor.destroy()
                mapper = StatementDialog(root, statement, list_sheets(statement))
                assert [combo.current() for combo in mapper.columns] == [1, 2, 3, 4, 5]
                mapper.submit()
                assert mapper.result is not None
                assert len(app.catalog.points) == 74 and not app.catalog.missing
                app.display_report(statement, mapper.sheet_name, mapper.result)
                root.update_idletasks()
                assert [row.group for row in app.rows] == [100, None]
                assert len(app.tree.get_children()) == 2
                editor = GroupDialog(root, app.catalog, "Неизвестный товар", None)
                assert editor.combo.current() == 0
                assert len(root.tk.splitlist(editor.combo.cget("values"))) == len(GROUPS) + 1
                editor.combo.current(1)
                editor.apply()
                assert editor.applied and editor.number == GROUPS[0]
                app.rows[1].assign(editor.number)
                export_xlsx(base / "result.xlsx", app.rows, app.catalog)
                assert (base / "result.xlsx").is_file()
                warehouse = base / "Остатки ТМЦ.csv"
                warehouse.write_text(
                    "Склад;Группа номенклатуры;Номенклатура;Начальный остаток;Приход;Расход;"
                    "Конечный остаток;Розничная цена;Октябрь 2026 г.;Ноябрь 2026 г.\n"
                    "Тест;Хозгруппа;Вино фрукт. 0,75л;2;3;2;3;4,80;4;1\n",
                    encoding="utf-8",
                )
                mapper = StatementDialog(root, warehouse, list_sheets(warehouse))
                assert mapper.use_warehouse.get() and mapper.sales.current() == 0
                mapper.sales.current(1)
                mapper.submit()
                assert mapper.result is not None and mapper.result.column_notes
                assert mapper.result.products[0].metrics == (5, None, 3, Decimal("14.4"))
                app.display_report(warehouse, mapper.sheet_name, mapper.result)
                assert app.rows[0].group == 36
                app.view.set(_VIEWS[2])
                app.redraw()
                assert not app.tree.get_children()
                export_xlsx(base / "warehouse-result.xlsx", app.rows, app.catalog, app.column_notes)
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
