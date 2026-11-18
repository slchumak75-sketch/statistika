"""Одно поле поверх ячейки: выделение текста или ручной ввод."""

from __future__ import annotations

import tkinter as tk
from tkinter import font, ttk
from typing import Callable


class InlineEntry(ttk.Entry):
    def __init__(self, tree: ttk.Treeview) -> None:
        self.value = tk.StringVar(master=tree)
        super().__init__(tree, textvariable=self.value, exportselection=False, takefocus=False)
        self.tree = tree
        self.text_font = font.nametofont("TkDefaultFont", root=tree)
        self.item: str | None = None
        self.on_apply: Callable[[str, str], bool] | None = None
        self.applying = False
        self.bind("<Return>", self.apply)
        self.bind("<Escape>", self._escape)
        self.bind("<FocusOut>", self._focus_out)
        tree.bind("<Configure>", lambda _event: self.dismiss(), add="+")
        tree.winfo_toplevel().bind("<ButtonPress-1>", self._outside_click, add="+")

    def open(self, item: str, column: str, text: str, on_apply: Callable[[str, str], bool] | None = None) -> None:
        self.dismiss()
        self.tree.see(item)
        self.tree.update_idletasks()
        if not (box := self.tree.bbox(item, column)):
            return
        x, y, width, height = box
        if on_apply is None:
            # Для длинного названия используем свободное место справа. Остаток доступен прокруткой поля.
            width = max(width, self.text_font.measure(text) + 12)
        left = max(1, x)
        right = min(self.tree.winfo_width() - 1, x + width)
        if right <= left:
            return
        self.item, self.on_apply = item, on_apply
        self.state(["readonly" if on_apply is None else "!readonly"])
        self.value.set(text)
        self.place(x=left, y=y, width=right - left, height=height)
        self.lift()
        self.icursor(0 if on_apply is None else "end")
        self.selection_clear()
        if on_apply is None:
            self.xview_moveto(0)
        else:
            self.selection_range(0, "end")
        self.focus_set()

    def dismiss(self, *, focus_tree: bool = False) -> None:
        if self.item is not None:
            self.item, self.on_apply = None, None
            self.place_forget()
            if focus_tree:
                self.tree.focus_set()

    def apply(self, _event=None) -> str:
        if self.item is not None:
            self.applying = True
            try:
                accepted = self.on_apply is None or self.on_apply(self.item, self.get())
            finally:
                self.applying = False
            if accepted:
                self.dismiss(focus_tree=True)
            elif self.item is not None:
                self.focus_set()
        return "break"

    def _focus_out(self, _event) -> None:
        if not self.applying:
            self.dismiss()

    def _outside_click(self, event) -> None:
        if not self.applying and event.widget is not self:
            self.dismiss()

    def _escape(self, _event) -> str:
        self.dismiss(focus_tree=True)
        return "break"
