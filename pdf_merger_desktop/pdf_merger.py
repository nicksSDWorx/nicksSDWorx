"""PDF Merger — offline desktop app.

A native Windows application for merging PDFs. No internet required, no
browser launch. Pure tkinter UI + pypdf for merging.
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from dataclasses import dataclass, field
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
except ImportError:  # pragma: no cover
    print("pypdf is required. Install it with: pip install pypdf", file=sys.stderr)
    raise

# Optional drag-and-drop support via tkinterdnd2. Falls back gracefully.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _HAS_DND = True
except Exception:
    TkinterDnD = None  # type: ignore
    DND_FILES = None  # type: ignore
    _HAS_DND = False


@dataclass
class PdfEntry:
    path: str
    name: str
    page_count: Optional[int] = None
    error: Optional[str] = None


class PdfMergerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.entries: List[PdfEntry] = []

        root.title("PDF Merger")
        root.geometry("720x520")
        root.minsize(560, 420)

        # Colors / style
        self.bg = "#f5f5fa"
        self.accent = "#5a5ad8"
        self.accent_dark = "#4242b4"
        self.danger = "#c0392b"
        root.configure(bg=self.bg)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=self.bg)
        style.configure("TLabel", background=self.bg, foreground="#2d2d42")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", foreground="#6b6b7b")
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            background=self.accent,
            foreground="white",
        )
        style.map(
            "Accent.TButton",
            background=[("active", self.accent_dark), ("disabled", "#c4c4d4")],
            foreground=[("disabled", "#ffffff")],
        )
        style.configure("TButton", padding=(10, 6))

        # Header
        header = ttk.Frame(root, padding=(16, 14, 16, 4))
        header.pack(fill="x")
        ttk.Label(header, text="PDF Merger", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Combine multiple PDFs into one. 100% offline — no internet required.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        # Drop zone / add button row
        top = ttk.Frame(root, padding=(16, 8))
        top.pack(fill="x")
        drop_text = (
            "Drag PDF files here, or click 'Add PDFs…'"
            if _HAS_DND
            else "Click 'Add PDFs…' to choose files"
        )
        self.drop_label = tk.Label(
            top,
            text=drop_text,
            bg="#ffffff",
            fg="#6b6b7b",
            relief="solid",
            bd=1,
            padx=12,
            pady=18,
            font=("Segoe UI", 10),
            highlightthickness=0,
        )
        self.drop_label.pack(fill="x")
        if _HAS_DND:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
            # Also make the whole window accept drops
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_label.bind("<Button-1>", lambda _e: self.add_files_dialog())

        # File list with scrollbar
        middle = ttk.Frame(root, padding=(16, 8))
        middle.pack(fill="both", expand=True)

        list_frame = ttk.Frame(middle)
        list_frame.pack(side="left", fill="both", expand=True)
        self.listbox = tk.Listbox(
            list_frame,
            activestyle="none",
            selectmode="browse",
            font=("Segoe UI", 10),
            bd=1,
            relief="solid",
            highlightthickness=0,
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        # Side buttons
        side = ttk.Frame(middle, padding=(10, 0, 0, 0))
        side.pack(side="right", fill="y")
        ttk.Button(side, text="Add PDFs…", command=self.add_files_dialog).pack(fill="x", pady=(0, 6))
        ttk.Button(side, text="Move Up", command=self.move_up).pack(fill="x", pady=2)
        ttk.Button(side, text="Move Down", command=self.move_down).pack(fill="x", pady=2)
        ttk.Button(side, text="Remove", command=self.remove_selected).pack(fill="x", pady=2)
        ttk.Button(side, text="Clear All", command=self.clear_all).pack(fill="x", pady=(6, 0))

        # Bottom: status + progress + merge button
        bottom = ttk.Frame(root, padding=(16, 8, 16, 16))
        bottom.pack(fill="x")

        self.status_var = tk.StringVar(value="No files added yet.")
        self.status_label = ttk.Label(bottom, textvariable=self.status_var, style="Sub.TLabel")
        self.status_label.pack(anchor="w")

        self.progress = ttk.Progressbar(bottom, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 10))

        self.merge_btn = ttk.Button(
            bottom,
            text="Merge PDFs",
            style="Accent.TButton",
            command=self.start_merge,
            state="disabled",
        )
        self.merge_btn.pack(anchor="e")

        self._refresh_list()

    # ---------- file management ----------

    def add_files_dialog(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if paths:
            self.add_files(paths)

    def _on_drop(self, event) -> None:
        raw = event.data or ""
        # tkdnd returns files in {curly braces} when paths contain spaces.
        paths = self._parse_dnd_paths(raw)
        if paths:
            self.add_files(paths)

    @staticmethod
    def _parse_dnd_paths(data: str) -> List[str]:
        paths: List[str] = []
        token = ""
        in_braces = False
        for ch in data:
            if ch == "{" and not in_braces:
                in_braces = True
                token = ""
            elif ch == "}" and in_braces:
                in_braces = False
                if token:
                    paths.append(token)
                token = ""
            elif ch == " " and not in_braces:
                if token:
                    paths.append(token)
                    token = ""
            else:
                token += ch
        if token:
            paths.append(token)
        return paths

    def add_files(self, paths) -> None:
        skipped: List[str] = []
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if not os.path.isfile(p):
                skipped.append(os.path.basename(p) + " (not found)")
                continue
            if not p.lower().endswith(".pdf"):
                skipped.append(os.path.basename(p) + " (not a PDF)")
                continue
            entry = PdfEntry(path=p, name=os.path.basename(p))
            try:
                reader = PdfReader(p, strict=False)
                # Access pages to trigger errors on corrupt files.
                entry.page_count = len(reader.pages)
            except (PdfReadError, OSError, ValueError) as exc:
                entry.error = f"Could not read PDF: {exc}"
            except Exception as exc:  # noqa: BLE001
                entry.error = f"Error: {exc}"
            self.entries.append(entry)
            added += 1

        self._refresh_list()

        parts = []
        if added:
            parts.append(f"Added {added} file{'s' if added != 1 else ''}.")
        if skipped:
            parts.append("Skipped: " + ", ".join(skipped))
        if parts:
            self.status_var.set(" ".join(parts))
        if skipped:
            messagebox.showwarning(
                "Some files were skipped",
                "The following files were not added:\n\n" + "\n".join(skipped),
            )

    def remove_selected(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.entries[idx]
        self._refresh_list()
        if self.entries:
            self.listbox.selection_set(min(idx, len(self.entries) - 1))

    def clear_all(self) -> None:
        if not self.entries:
            return
        self.entries.clear()
        self._refresh_list()
        self.status_var.set("Cleared.")

    def move_up(self) -> None:
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self.entries[i - 1], self.entries[i] = self.entries[i], self.entries[i - 1]
        self._refresh_list()
        self.listbox.selection_set(i - 1)

    def move_down(self) -> None:
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.entries) - 1:
            return
        i = sel[0]
        self.entries[i + 1], self.entries[i] = self.entries[i], self.entries[i + 1]
        self._refresh_list()
        self.listbox.selection_set(i + 1)

    def _refresh_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for i, e in enumerate(self.entries, start=1):
            if e.error:
                label = f"{i:>2}.  {e.name}   —  ERROR: {e.error}"
            else:
                pages = e.page_count
                page_txt = f"{pages} page" + ("s" if pages != 1 else "")
                label = f"{i:>2}.  {e.name}   —  {page_txt}"
            self.listbox.insert(tk.END, label)
            if e.error:
                self.listbox.itemconfig(tk.END, fg=self.danger)

        valid = [e for e in self.entries if not e.error]
        self.merge_btn.configure(state=("normal" if len(valid) >= 2 else "disabled"))
        if not self.entries:
            self.status_var.set("No files added yet.")
        else:
            self.status_var.set(
                f"{len(self.entries)} file(s) loaded, {len(valid)} valid. "
                f"{'Ready to merge.' if len(valid) >= 2 else 'Add at least 2 valid PDFs.'}"
            )

    # ---------- merging ----------

    def start_merge(self) -> None:
        valid = [e for e in self.entries if not e.error]
        if len(valid) < 2:
            messagebox.showinfo("Not enough files", "Add at least 2 valid PDFs to merge.")
            return

        out_path = filedialog.asksaveasfilename(
            title="Save merged PDF",
            defaultextension=".pdf",
            initialfile="merged.pdf",
            filetypes=[("PDF file", "*.pdf")],
        )
        if not out_path:
            return

        self.merge_btn.configure(state="disabled")
        self.progress.configure(value=0, maximum=len(valid))
        self.status_var.set("Merging…")

        thread = threading.Thread(
            target=self._merge_worker, args=(valid, out_path), daemon=True
        )
        thread.start()

    def _merge_worker(self, entries: List[PdfEntry], out_path: str) -> None:
        try:
            writer = PdfWriter()
            for i, entry in enumerate(entries, start=1):
                reader = PdfReader(entry.path, strict=False)
                for page in reader.pages:
                    writer.add_page(page)
                self.root.after(0, lambda v=i: self.progress.configure(value=v))
                self.root.after(0, lambda n=entry.name, k=i, total=len(entries):
                                self.status_var.set(f"Merging {k}/{total}: {n}"))

            with open(out_path, "wb") as f:
                writer.write(f)
            writer.close()
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._merge_failed(str(exc), tb))
            return
        self.root.after(0, lambda: self._merge_succeeded(out_path, len(entries)))

    def _merge_succeeded(self, out_path: str, count: int) -> None:
        self.progress.configure(value=self.progress["maximum"])
        self.status_var.set(f"Merged {count} PDFs → {out_path}")
        self.merge_btn.configure(state="normal")
        if messagebox.askyesno(
            "Merge complete",
            f"Merged {count} PDFs successfully:\n{out_path}\n\nOpen the containing folder?",
        ):
            self._open_folder(os.path.dirname(out_path))

    def _merge_failed(self, msg: str, tb: str) -> None:
        self.progress.configure(value=0)
        self.status_var.set("Merge failed.")
        self.merge_btn.configure(state="normal")
        messagebox.showerror("Merge failed", f"{msg}\n\nDetails:\n{tb}")

    @staticmethod
    def _open_folder(path: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", path])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass


def main() -> None:
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    PdfMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
