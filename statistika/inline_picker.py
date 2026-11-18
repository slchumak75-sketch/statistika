"""Один выпадающий список поверх таблицы, привязанный к первой ячейке."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable

from statistika.catalog import GROUPS, Catalog

_OPTIONS = (None, *GROUPS)


class InlineGroupPicker(ttk.Frame):
    def __init__(self, tree: ttk.Treeview, on_pick: Callable[[str, int | None], None]) -> None:
        super().__init__(tree, relief="solid", borderwidth=1)
        self.tree = tree
        self.on_pick = on_pick
        self.item: str | None = None
        self.options = tk.Listbox(
            self, exportselection=False, height=len(_OPTIONS), width=52, font="TkDefaultFont",
            borderwidth=0, highlightthickness=0, background="white", foreground="#16304c",
            selectbackground="#c6dffa", selectforeground="#16304c", activestyle="none",
        )
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.options.yview, takefocus=False)
        scrollbar.pack(side="right", fill="y")
        self.options.pack(side="left", fill="both", expand=True)
        self.options.configure(yscrollcommand=scrollbar.set)
        self.options.bind("<ButtonRelease-1>", self._click_option)
        self.options.bind("<Return>", self.choose)
        self.options.bind("<Escape>", self._escape)
        self.options.bind("<FocusOut>", lambda _event: self.dismiss())
        tree.winfo_toplevel().bind("<ButtonPress-1>", self._outside_click, add="+")
        tree.bind("<Configure>", lambda _event: self.dismiss(), add="+")

    def open(self, item: str, current: int | None, catalog: Catalog) -> None:
        self.dismiss()
        self.tree.see(item)
        self.tree.update_idletasks()
        if not (cell := self.tree.bbox(item, "#1")):
            return
        x, y, width, height = cell
        tree_width, tree_height = self.tree.winfo_width(), self.tree.winfo_height()
        if x + width <= 0 or x >= tree_width:
            # Enter also works after horizontally scrolling beyond column 1.
            self.tree.xview_moveto(0)
            self.tree.update_idletasks()
            if not (cell := self.tree.bbox(item, "#1")):
                return
            x, y, width, height = cell
        self.options.delete(0, "end")
        self.options.insert("end", "Без группы (пустая ячейка)", *(catalog.short_label(number) for number in GROUPS))
        self.update_idletasks()
        desired = self.winfo_reqheight()
        below, above = tree_height - y - height - 1, y - 1
        opens_down = below >= desired or below >= above
        popup_height = min(desired, below if opens_down else above)
        popup_width = min(tree_width - 2, max(width, self.winfo_reqwidth()))
        if popup_height <= 0 or popup_width <= 0:
            return
        self.item = item
        self.place(
            x=max(1, min(x, tree_width - popup_width - 1)), y=y + height if opens_down else y - popup_height,
            width=popup_width, height=popup_height,
        )
        self.lift()
        selected = _OPTIONS.index(current)
        self.options.selection_set(selected)
        self.options.activate(selected)
        self.options.see(selected)
        self.options.focus_set()

    def dismiss(self, *, focus_tree: bool = False) -> None:
        if self.item is not None:
            self.item = None
            self.place_forget()
            if focus_tree:
                self.tree.focus_set()

    def _outside_click(self, event) -> None:
        if not str(event.widget).startswith(f"{self}."):
            self.dismiss()

    def _escape(self, _event=None) -> str:
        self.dismiss(focus_tree=True)
        return "break"

    def _click_option(self, event) -> str:
        index = self.options.nearest(event.y)
        box = self.options.bbox(index)
        if box is not None and box[1] <= event.y < box[1] + box[3]:
            self.options.selection_clear(0, "end")
            self.options.selection_set(index)
            return self.choose()
        return "break"

    def choose(self, _event=None) -> str:
        selected = self.options.curselection()
        item = self.item
        if item is not None and selected:
            self.dismiss(focus_tree=True)
            self.on_pick(item, _OPTIONS[selected[0]])
        return "break"
