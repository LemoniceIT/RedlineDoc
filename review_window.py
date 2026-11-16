"""Side-by-side document review in the desktop application."""

import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from comparison import Comparison
from config import GUI_ACCENT_COLOR, GUI_BG_COLOR, GUI_PANEL_BG, GUI_TEXT_COLOR
from i18n import t
from reports import export_report, location_text, status_text, summary_text

logger = logging.getLogger(__name__)


class ReviewWindow(ctk.CTkToplevel):
    def __init__(self, parent, comparison: Comparison, source_paths: tuple[Path, Path]):
        super().__init__(parent)
        self.comparison = comparison
        self._source_paths = tuple(path.resolve() for path in source_paths)
        self.title(t("review.title"))
        self.geometry("1120x720")
        self.minsize(760, 420)
        self.configure(fg_color=GUI_BG_COLOR)
        self._visible_changes = []
        self._change_index = -1
        self._search_position = None

        summary = ctk.CTkLabel(self, text=summary_text(comparison), text_color=GUI_ACCENT_COLOR)
        summary.pack(anchor="w", padx=14, pady=(12, 6))
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=12, pady=(0, 8))

        buttons = [
            ("←", lambda: self._navigate(-1)),
            ("→", lambda: self._navigate(1)),
        ]
        for label, command in buttons:
            ctk.CTkButton(toolbar, text=label, width=34, command=command).pack(side="left", padx=2)

        self._position = ctk.CTkLabel(toolbar, text="0 / 0", width=58)
        self._position.pack(side="left", padx=4)

        self._filters = {
            t("review." + kind): kind
            for kind in ("all", "changes", "inserted", "deleted", "modified", "moved")
        }
        ctk.CTkOptionMenu(
            toolbar, values=list(self._filters), width=145, command=self._render,
        ).pack(side="left", padx=4)

        self._search = ctk.CTkEntry(toolbar, placeholder_text=t("review.search"))
        self._search.pack(side="left", padx=4, fill="x", expand=True)
        self._search.bind("<Return>", self._find)
        self._search.bind("<KeyRelease>", self._search_edited)
        ctk.CTkButton(
            toolbar, text=t("review.find"), width=62, command=self._find,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            toolbar, text=t("review.export"), width=86, command=self._export,
        ).pack(side="right", padx=2)

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        for column in (0, 1):
            content.grid_columnconfigure(column, weight=1, uniform="documents")
        content.grid_rowconfigure(1, weight=1)
        self._panes = []
        for column, name in enumerate((comparison.original_name, comparison.modified_name)):
            ctk.CTkLabel(content, text=name, anchor="w").grid(
                row=0, column=column, sticky="ew", padx=4,
            )
            pane = tk.Text(
                content, wrap="none", width=1, borderwidth=0, padx=8, pady=8,
                font=("Consolas", 11), background=GUI_PANEL_BG, foreground=GUI_TEXT_COLOR,
                highlightthickness=0, undo=False,
            )
            pane.grid(row=1, column=column, sticky="nsew", padx=3)
            horizontal = ctk.CTkScrollbar(content, orientation="horizontal", command=pane.xview)
            horizontal.grid(row=2, column=column, sticky="ew", padx=3)
            pane.configure(xscrollcommand=horizontal.set)
            pane.tag_configure("location", foreground="#84ba95")
            pane.tag_configure("changed", background=("#6b2531", "#214e33")[column])
            pane.tag_configure("moved", foreground="#7ccaff")
            pane.tag_configure("focus", background="#343d35")
            pane.tag_lower("focus")
            pane.tag_configure("search", background="#ffcf73", foreground="#000000")
            pane.bind("<MouseWheel>", self._wheel)
            pane.bind("<Button-4>", lambda _event: self._scroll("scroll", -3, "units") or "break")
            pane.bind("<Button-5>", lambda _event: self._scroll("scroll", 3, "units") or "break")
            self._panes.append(pane)

        vertical = ctk.CTkScrollbar(content, command=self._scroll)
        vertical.grid(row=1, column=2, sticky="ns")
        for pane in self._panes:
            pane.configure(yscrollcommand=lambda first, last, pane=pane: self._sync(
                pane, first, last, vertical,
            ))
        self._render(t("review.all"))

    def _scroll(self, *args):
        self._panes[0].yview(*args)

    def _sync(self, pane, first, last, scrollbar):
        scrollbar.set(first, last)
        peer = self._panes[1 - self._panes.index(pane)]
        if abs(peer.yview()[0] - float(first)) > 1e-6:
            peer.yview_moveto(first)

    def _wheel(self, event):
        self._scroll("scroll", -3 if event.delta > 0 else 3, "units")
        return "break"

    def _render(self, selection):
        kind = self._filters[selection]
        self._visible_changes.clear()
        self._change_index = -1
        self._search_position = None
        self._position.configure(text="0 / 0")

        for pane in self._panes:
            pane.configure(state="normal")
            pane.delete("1.0", "end")
        for row in self.comparison.rows:
            if kind != "all" and kind != row.kind and (kind != "changes" or row.kind == "equal"):
                continue
            start = self._panes[0].index("end-1c")
            if row.kind != "equal":
                self._visible_changes.append(start)
            counts = []
            for pane, spans, side in zip(self._panes, (row.before, row.after), ("original", "modified")):
                location = location_text(row, side)
                label = " · ".join(part for part in (status_text(row), location) if part)
                pane.insert("end", label + "\n", "moved" if row.move else "location")
                count = 0
                for span in spans:
                    pane.insert("end", span.text, "changed" if span.changed else ())
                    count += span.text.count("\n")
                counts.append(count)
            for pane, count in zip(self._panes, counts):
                pane.insert("end", "\n" * (max(counts) - count + 2))
        for pane in self._panes:
            pane.configure(state="disabled")
            pane.yview_moveto(0)
        self._position.configure(text=f"0 / {len(self._visible_changes)}")

    def _navigate(self, direction):
        count = len(self._visible_changes)
        if not count:
            return
        if self._change_index < 0:
            self._change_index = 0 if direction > 0 else count - 1
        else:
            self._change_index = (self._change_index + direction) % count
        start = self._visible_changes[self._change_index]
        for pane in self._panes:
            pane.tag_remove("focus", "1.0", "end")
            pane.tag_add("focus", start, f"{start} lineend+1c")
        self._panes[0].yview(start)
        self._position.configure(text=f"{self._change_index + 1} / {count}")

    def _search_edited(self, event):
        if event.keysym != "Return":
            self._search_position = None
            for pane in self._panes:
                pane.tag_remove("search", "1.0", "end")

    def _find(self, _event=None):
        query = self._search.get().strip()
        if not query:
            return
        for pane in self._panes:
            pane.tag_remove("search", "1.0", "end")
        start_side, start = self._search_position or (0, "1.0")
        ranges = [(start_side, start, "end"), (1 - start_side, "1.0", "end")]
        if start != "1.0":
            ranges.append((start_side, "1.0", start))
        for side, first, stop in ranges:
            pane = self._panes[side]
            length = tk.IntVar(master=self)
            index = pane.search(query, first, stopindex=stop, count=length, nocase=True)
            if index:
                end = f"{index}+{length.get()}c"
                pane.tag_add("search", index, end)
                pane.see(index)
                self._search_position = side, pane.index(end)
                return
        self.bell()

    def _export(self):
        formats = [("HTML", "*.html"), ("PDF", "*.pdf"), ("JSON", "*.json")]
        selected = tk.StringVar(master=self)
        filename = filedialog.asksaveasfilename(
            parent=self, title=t("review.export"), filetypes=formats, typevariable=selected,
            initialfile="redlinedoc", confirmoverwrite=True,
        )
        if not filename:
            return
        path = Path(filename)
        if not path.suffix:
            extensions = {name: pattern[1:] for name, pattern in formats}
            path = path.with_suffix(extensions.get(selected.get(), ".html"))
            if path.exists() and not messagebox.askyesno(
                t("review.export"), t("review.overwrite").format(name=path.name), parent=self,
            ):
                return
        try:
            if path.resolve() in self._source_paths:
                raise ValueError(t("review.source_export"))
            export_report(self.comparison, path)
        except Exception as exc:
            logger.exception("Report export failed")
            messagebox.showerror(t("review.export"), t("compare.error").format(error=exc), parent=self)
