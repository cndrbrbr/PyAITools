#!/usr/bin/env python3

import hashlib
import os
import tkinter as tk

from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "PyWinDiff"


# ----------------------------------------------------------------------
# Hilfsfunktionen
# ----------------------------------------------------------------------

def read_text_file(path: Path):
    """
    Liest eine Textdatei und versucht mehrere typische Encodings.
    """
    encodings = [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ]

    for encoding in encodings:
        try:
            with path.open("r", encoding=encoding) as f:
                return f.readlines(), encoding
        except UnicodeDecodeError:
            pass
        except OSError as e:
            raise RuntimeError(str(e))

    raise RuntimeError(f"Datei konnte nicht gelesen werden: {path}")


def normalize_line(line: str, ignore_case=False, ignore_whitespace=False):
    """
    Normalisiert eine Zeile für den Vergleich.
    Die Anzeige verwendet weiterhin den Originaltext.
    """
    result = line

    if ignore_whitespace:
        result = " ".join(result.split())

    if ignore_case:
        result = result.lower()

    return result


def file_hash(path: Path, block_size=65536):
    """
    Berechnet SHA-256 einer Datei.
    """
    h = hashlib.sha256()

    with path.open("rb") as f:
        while True:
            block = f.read(block_size)

            if not block:
                break

            h.update(block)

    return h.digest()


def files_binary_equal(path1: Path, path2: Path):
    """
    Schneller Vergleich zweier Dateien.
    """
    try:
        if path1.stat().st_size != path2.stat().st_size:
            return False

        return file_hash(path1) == file_hash(path2)

    except OSError:
        return False


# ----------------------------------------------------------------------
# Datenmodell für Verzeichnisvergleich
# ----------------------------------------------------------------------

@dataclass
class DirectoryEntry:
    relative_path: Path
    left_path: Path | None
    right_path: Path | None
    status: str


# ----------------------------------------------------------------------
# Hauptprogramm
# ----------------------------------------------------------------------

class PyWinDiff(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1400x850")
        self.minsize(900, 600)

        self.left_path = tk.StringVar()
        self.right_path = tk.StringVar()

        self.ignore_case = tk.BooleanVar(value=False)
        self.ignore_whitespace = tk.BooleanVar(value=False)
        self.show_identical = tk.BooleanVar(value=True)

        self.directory_entries = {}

        self._build_gui()
        self._configure_tags()

    # ------------------------------------------------------------------
    # GUI
    # ------------------------------------------------------------------

    def _build_gui(self):

        # --------------------------------------------------------------
        # Pfadauswahl
        # --------------------------------------------------------------

        path_frame = ttk.Frame(self, padding=8)
        path_frame.pack(fill=tk.X)

        ttk.Label(path_frame, text="Links:").grid(
            row=0,
            column=0,
            sticky=tk.W,
            padx=4,
            pady=4
        )

        ttk.Entry(
            path_frame,
            textvariable=self.left_path
        ).grid(
            row=0,
            column=1,
            sticky=tk.EW,
            padx=4
        )

        ttk.Button(
            path_frame,
            text="Datei...",
            command=lambda: self.select_file(self.left_path)
        ).grid(row=0, column=2, padx=2)

        ttk.Button(
            path_frame,
            text="Ordner...",
            command=lambda: self.select_directory(self.left_path)
        ).grid(row=0, column=3, padx=2)

        ttk.Label(path_frame, text="Rechts:").grid(
            row=1,
            column=0,
            sticky=tk.W,
            padx=4,
            pady=4
        )

        ttk.Entry(
            path_frame,
            textvariable=self.right_path
        ).grid(
            row=1,
            column=1,
            sticky=tk.EW,
            padx=4
        )

        ttk.Button(
            path_frame,
            text="Datei...",
            command=lambda: self.select_file(self.right_path)
        ).grid(row=1, column=2, padx=2)

        ttk.Button(
            path_frame,
            text="Ordner...",
            command=lambda: self.select_directory(self.right_path)
        ).grid(row=1, column=3, padx=2)

        path_frame.columnconfigure(1, weight=1)

        # --------------------------------------------------------------
        # Optionen
        # --------------------------------------------------------------

        option_frame = ttk.Frame(self, padding=(8, 0))
        option_frame.pack(fill=tk.X)

        ttk.Checkbutton(
            option_frame,
            text="Groß-/Kleinschreibung ignorieren",
            variable=self.ignore_case
        ).pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(
            option_frame,
            text="Leerzeichen ignorieren",
            variable=self.ignore_whitespace
        ).pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(
            option_frame,
            text="Identische Dateien anzeigen",
            variable=self.show_identical
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            option_frame,
            text="Vergleichen",
            command=self.compare
        ).pack(side=tk.RIGHT, padx=5)

        # --------------------------------------------------------------
        # Notebook
        # --------------------------------------------------------------

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(
            fill=tk.BOTH,
            expand=True,
            padx=8,
            pady=8
        )

        self.file_tab = ttk.Frame(self.notebook)
        self.directory_tab = ttk.Frame(self.notebook)

        self.notebook.add(
            self.file_tab,
            text="Dateivergleich"
        )

        self.notebook.add(
            self.directory_tab,
            text="Verzeichnisvergleich"
        )

        self._build_file_compare_tab()
        self._build_directory_compare_tab()

        # --------------------------------------------------------------
        # Statusbar
        # --------------------------------------------------------------

        self.status = tk.StringVar(value="Bereit.")

        statusbar = ttk.Label(
            self,
            textvariable=self.status,
            anchor=tk.W,
            relief=tk.SUNKEN
        )

        statusbar.pack(fill=tk.X, side=tk.BOTTOM)

    # ------------------------------------------------------------------

    def _build_file_compare_tab(self):

        header = ttk.Frame(self.file_tab)
        header.pack(fill=tk.X)

        self.left_label = ttk.Label(
            header,
            text="Links",
            anchor=tk.CENTER
        )

        self.left_label.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True
        )

        self.right_label = ttk.Label(
            header,
            text="Rechts",
            anchor=tk.CENTER
        )

        self.right_label.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True
        )

        text_frame = ttk.Frame(self.file_tab)
        text_frame.pack(
            fill=tk.BOTH,
            expand=True
        )

        text_frame.columnconfigure(0, weight=1)
        text_frame.columnconfigure(1, weight=1)
        text_frame.rowconfigure(0, weight=1)

        left_frame = ttk.Frame(text_frame)
        right_frame = ttk.Frame(text_frame)

        left_frame.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        right_frame.grid(
            row=0,
            column=1,
            sticky="nsew"
        )

        self.left_text = tk.Text(
            left_frame,
            wrap=tk.NONE,
            font=("Consolas", 10),
            undo=False
        )

        self.right_text = tk.Text(
            right_frame,
            wrap=tk.NONE,
            font=("Consolas", 10),
            undo=False
        )

        left_y = ttk.Scrollbar(
            left_frame,
            orient=tk.VERTICAL
        )

        right_y = ttk.Scrollbar(
            right_frame,
            orient=tk.VERTICAL
        )

        left_x = ttk.Scrollbar(
            left_frame,
            orient=tk.HORIZONTAL,
            command=self.left_text.xview
        )

        right_x = ttk.Scrollbar(
            right_frame,
            orient=tk.HORIZONTAL,
            command=self.right_text.xview
        )

        self.left_text.configure(
            xscrollcommand=left_x.set
        )

        self.right_text.configure(
            xscrollcommand=right_x.set
        )

        self.left_text.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        left_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        left_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        self.right_text.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        right_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        right_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        left_frame.rowconfigure(0, weight=1)
        left_frame.columnconfigure(0, weight=1)

        right_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)

        # Synchrones vertikales Scrollen

        def scroll_left(*args):
            self.left_text.yview(*args)
            self.right_text.yview(*args)

        def scroll_right(*args):
            self.left_text.yview(*args)
            self.right_text.yview(*args)

        left_y.configure(command=scroll_left)
        right_y.configure(command=scroll_right)

        self.left_text.configure(
            yscrollcommand=lambda first, last:
            self._sync_scrollbars(left_y, right_y, first, last)
        )

        self.right_text.configure(
            yscrollcommand=lambda first, last:
            self._sync_scrollbars(right_y, left_y, first, last)
        )

    # ------------------------------------------------------------------

    def _build_directory_compare_tab(self):

        frame = ttk.Frame(
            self.directory_tab,
            padding=4
        )

        frame.pack(
            fill=tk.BOTH,
            expand=True
        )

        columns = (
            "status",
            "path",
            "left",
            "right"
        )

        self.tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings"
        )

        self.tree.heading(
            "status",
            text="Status"
        )

        self.tree.heading(
            "path",
            text="Relativer Pfad"
        )

        self.tree.heading(
            "left",
            text="Links"
        )

        self.tree.heading(
            "right",
            text="Rechts"
        )

        self.tree.column(
            "status",
            width=130,
            stretch=False
        )

        self.tree.column(
            "path",
            width=450
        )

        self.tree.column(
            "left",
            width=250
        )

        self.tree.column(
            "right",
            width=250
        )

        y_scroll = ttk.Scrollbar(
            frame,
            orient=tk.VERTICAL,
            command=self.tree.yview
        )

        x_scroll = ttk.Scrollbar(
            frame,
            orient=tk.HORIZONTAL,
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        y_scroll.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        x_scroll.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.bind(
            "<Double-1>",
            self.open_selected_directory_file
        )

        self.tree.tag_configure(
            "different",
            background="#fff1b8"
        )

        self.tree.tag_configure(
            "leftonly",
            background="#ffd6d6"
        )

        self.tree.tag_configure(
            "rightonly",
            background="#d8ffd8"
        )

        self.tree.tag_configure(
            "identical",
            foreground="#777777"
        )

    # ------------------------------------------------------------------

    def _configure_tags(self):

        for widget in (
            self.left_text,
            self.right_text
        ):

            widget.tag_configure(
                "equal",
                background="white",
                foreground="black"
            )

            widget.tag_configure(
                "delete",
                background="yellow",
                foreground="black"
            )

            widget.tag_configure(
                "insert",
                background="yellow",
                foreground="black"
            )

            widget.tag_configure(
                "replace",
                background="yellow",
                foreground="black"
            )

            widget.tag_configure(
                "empty",
                background ="#eeeeee",
                foreground="black"
            )

    # ------------------------------------------------------------------
    # Dateiauswahl
    # ------------------------------------------------------------------

    def select_file(self, variable):

        path = filedialog.askopenfilename()

        if path:
            variable.set(path)

    def select_directory(self, variable):

        path = filedialog.askdirectory()

        if path:
            variable.set(path)

    # ------------------------------------------------------------------
    # Vergleich
    # ------------------------------------------------------------------

    def compare(self):

        left = Path(self.left_path.get())
        right = Path(self.right_path.get())

        if not left.exists():
            messagebox.showerror(
                APP_TITLE,
                f"Linker Pfad existiert nicht:\n{left}"
            )
            return

        if not right.exists():
            messagebox.showerror(
                APP_TITLE,
                f"Rechter Pfad existiert nicht:\n{right}"
            )
            return

        if left.is_file() and right.is_file():

            self.compare_files(left, right)

            self.notebook.select(
                self.file_tab
            )

        elif left.is_dir() and right.is_dir():

            self.compare_directories(
                left,
                right
            )

            self.notebook.select(
                self.directory_tab
            )

        else:

            messagebox.showerror(
                APP_TITLE,
                "Es müssen entweder zwei Dateien "
                "oder zwei Verzeichnisse ausgewählt werden."
            )

    # ------------------------------------------------------------------
    # Dateivergleich
    # ------------------------------------------------------------------

    def compare_files(self, left: Path, right: Path):

        try:
            left_lines, left_encoding = read_text_file(left)
            right_lines, right_encoding = read_text_file(right)

        except Exception as e:

            messagebox.showerror(
                APP_TITLE,
                str(e)
            )

            return

        self.left_label.configure(
            text=f"{left} [{left_encoding}]"
        )

        self.right_label.configure(
            text=f"{right} [{right_encoding}]"
        )

        self.left_text.configure(state=tk.NORMAL)
        self.right_text.configure(state=tk.NORMAL)

        self.left_text.delete(
            "1.0",
            tk.END
        )

        self.right_text.delete(
            "1.0",
            tk.END
        )

        left_normalized = [
            normalize_line(
                line,
                self.ignore_case.get(),
                self.ignore_whitespace.get()
            )
            for line in left_lines
        ]

        right_normalized = [
            normalize_line(
                line,
                self.ignore_case.get(),
                self.ignore_whitespace.get()
            )
            for line in right_lines
        ]

        matcher = SequenceMatcher(
            None,
            left_normalized,
            right_normalized,
            autojunk=False
        )

        different_blocks = 0

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():

            if tag == "equal":

                count = i2 - i1

                for offset in range(count):

                    self.insert_diff_line(
                        left_lines[i1 + offset],
                        right_lines[j1 + offset],
                        "equal",
                        "equal",
                        i1 + offset + 1,
                        j1 + offset + 1
                    )

            elif tag == "replace":

                different_blocks += 1

                left_count = i2 - i1
                right_count = j2 - j1

                max_count = max(
                    left_count,
                    right_count
                )

                for offset in range(max_count):

                    if offset < left_count:
                        left_line = left_lines[i1 + offset]
                        left_no = i1 + offset + 1
                    else:
                        left_line = ""
                        left_no = None

                    if offset < right_count:
                        right_line = right_lines[j1 + offset]
                        right_no = j1 + offset + 1
                    else:
                        right_line = ""
                        right_no = None

                    self.insert_diff_line(
                        left_line,
                        right_line,
                        "replace" if left_no else "empty",
                        "replace" if right_no else "empty",
                        left_no,
                        right_no
                    )

            elif tag == "delete":

                different_blocks += 1

                for index in range(i1, i2):

                    self.insert_diff_line(
                        left_lines[index],
                        "",
                        "delete",
                        "empty",
                        index + 1,
                        None
                    )

            elif tag == "insert":

                different_blocks += 1

                for index in range(j1, j2):

                    self.insert_diff_line(
                        "",
                        right_lines[index],
                        "empty",
                        "insert",
                        None,
                        index + 1
                    )

        self.left_text.configure(
            state=tk.DISABLED
        )

        self.right_text.configure(
            state=tk.DISABLED
        )

        self.left_text.yview_moveto(0)
        self.right_text.yview_moveto(0)

        if different_blocks == 0:

            self.status.set(
                "Dateien sind identisch."
            )

        else:

            self.status.set(
                f"{different_blocks} Unterschiedsblock/"
                f"Unterschiedsblöcke gefunden."
            )

    # ------------------------------------------------------------------

    def insert_diff_line(
        self,
        left_line,
        right_line,
        left_tag,
        right_tag,
        left_number,
        right_number
        ):

        # Newline entfernen, da wir selbst eines ergänzen.
        left_line = left_line.rstrip("\r\n")
        right_line = right_line.rstrip("\r\n")

        if left_number is None:
            left_prefix = "      "
        else:
            left_prefix = f"{left_number:5d} "

        if right_number is None:
            right_prefix = "      "
        else:
            right_prefix = f"{right_number:5d} "

        # Tag DIREKT beim Einfügen anwenden
        self.left_text.insert(
            tk.END,
            left_prefix + left_line + "\n",
            left_tag
        )

        self.right_text.insert(
            tk.END,
            right_prefix + right_line + "\n",
            right_tag
        )

    # ------------------------------------------------------------------
    # Verzeichnisvergleich
    # ------------------------------------------------------------------

    def compare_directories(
        self,
        left_dir: Path,
        right_dir: Path
    ):

        self.tree.delete(
            *self.tree.get_children()
        )

        self.directory_entries.clear()

        left_files = {
            p.relative_to(left_dir): p
            for p in left_dir.rglob("*")
            if p.is_file()
        }

        right_files = {
            p.relative_to(right_dir): p
            for p in right_dir.rglob("*")
            if p.is_file()
        }

        all_paths = sorted(
            set(left_files) |
            set(right_files),
            key=lambda p: str(p).lower()
        )

        identical_count = 0
        different_count = 0
        left_only_count = 0
        right_only_count = 0

        for relative_path in all_paths:

            left_file = left_files.get(
                relative_path
            )

            right_file = right_files.get(
                relative_path
            )

            if left_file is None:

                status = "Nur rechts"
                tag = "rightonly"

                right_only_count += 1

            elif right_file is None:

                status = "Nur links"
                tag = "leftonly"

                left_only_count += 1

            elif files_binary_equal(
                left_file,
                right_file
            ):

                status = "Identisch"
                tag = "identical"

                identical_count += 1

                if not self.show_identical.get():
                    continue

            else:

                status = "Unterschiedlich"
                tag = "different"

                different_count += 1

            item_id = self.tree.insert(
                "",
                tk.END,
                values=(
                    status,
                    str(relative_path),
                    str(left_file or ""),
                    str(right_file or "")
                ),
                tags=(tag,)
            )

            self.directory_entries[item_id] = DirectoryEntry(
                relative_path=relative_path,
                left_path=left_file,
                right_path=right_file,
                status=status
            )

        self.status.set(
            f"Identisch: {identical_count} | "
            f"Unterschiedlich: {different_count} | "
            f"Nur links: {left_only_count} | "
            f"Nur rechts: {right_only_count}"
        )

    # ------------------------------------------------------------------
    # Datei aus Verzeichnisvergleich öffnen
    # ------------------------------------------------------------------

    def open_selected_directory_file(
        self,
        event=None
    ):

        selection = self.tree.selection()

        if not selection:
            return

        item_id = selection[0]

        entry = self.directory_entries.get(
            item_id
        )

        if entry is None:
            return

        if (
            entry.left_path is None or
            entry.right_path is None
        ):

            messagebox.showinfo(
                APP_TITLE,
                "Die Datei existiert nur auf einer Seite."
            )

            return

        self.left_path.set(
            str(entry.left_path)
        )

        self.right_path.set(
            str(entry.right_path)
        )

        self.compare_files(
            entry.left_path,
            entry.right_path
        )

        self.notebook.select(
            self.file_tab
        )

    # ------------------------------------------------------------------
    # Scroll-Synchronisation
    # ------------------------------------------------------------------

    @staticmethod
    def _sync_scrollbars(
        primary,
        secondary,
        first,
        last
    ):

        primary.set(
            first,
            last
        )

        secondary.set(
            first,
            last
        )


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------

def main():

    app = PyWinDiff()
    app.mainloop()


if __name__ == "__main__":
    main()
