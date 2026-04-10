"""
Stamkaart Word naar Excel Converter

Extracts employee contract and salary history from a Dutch HR Word
document ("stamkaarten") and exports them to a structured Excel file.

Uses python-docx to walk through paragraphs and tables in document
order, matching each table to the correct employee and section.

Usage: Double-click the .exe or run `python app.py`
"""

import os
import re
import sys
import logging
import traceback
import threading
from datetime import datetime

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

LOG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(sys.argv[0])), "stamkaart_debug.log"
)

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_date(text: str):
    """Try to parse a date string into a datetime object."""
    if text is None:
        return None
    text = text.strip()
    if not text or text == "-":
        return None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_salary(text: str):
    """Parse a salary string like '2.389,44' or '2389.44' into a float."""
    if text is None:
        return None
    text = text.strip().replace(" ", "").replace("\u20ac", "")
    if not text or text == "-":
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def normalize(text: str) -> str:
    """Lowercase + collapse whitespace for matching."""
    return re.sub(r"\s+", " ", text.strip().lower())


def find_col(header_cells: list[str], target: str) -> int | None:
    """
    Find the column index whose normalized header text contains `target`.
    Returns None if not found.
    """
    for i, cell in enumerate(header_cells):
        if target in normalize(cell):
            return i
    return None


# ---------------------------------------------------------------------------
# Word document parsing
# ---------------------------------------------------------------------------

def iter_doc_elements(doc: Document):
    """
    Yield paragraphs and tables in document order.
    Returns tuples of ("paragraph", Paragraph) or ("table", Table).
    """
    for child in doc.element.body:
        tag = child.tag.split("}")[-1]  # strip namespace
        if tag == "p":
            yield ("paragraph", Paragraph(child, doc))
        elif tag == "tbl":
            yield ("table", Table(child, doc))


def process_docx(docx_path: str, progress_callback=None) -> list[dict]:
    """
    Main processing function.

    Walks through the Word document in element order:
      - Paragraphs set the current employee (via "Stamkaart van") and
        the current section ("Contract mutaties" / "Salaris mutaties").
      - Tables are parsed according to the active section, using column
        headers to find the right data.
    """

    def report(msg):
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    report(f"Document openen: {docx_path}")

    doc = Document(docx_path)

    all_rows: list[dict] = []
    current_employee: str | None = None
    current_section: str | None = None  # "contract" or "salary"
    employee_set: set[str] = set()

    for elem_type, elem in iter_doc_elements(doc):

        if elem_type == "paragraph":
            text = elem.text.strip()
            if not text:
                continue

            logger.debug(f"Paragraph: {text!r}")

            # Detect employee name
            m = re.search(r"stamkaart\s+van\s+(.+)", text, re.IGNORECASE)
            if m:
                current_employee = m.group(1).strip()
                current_section = None  # reset section for new employee
                if current_employee not in employee_set:
                    employee_set.add(current_employee)
                    report(f"Medewerker gevonden: {current_employee}")
                continue

            # Detect section headers
            low = text.lower()
            if "contract" in low and "mutatie" in low:
                current_section = "contract"
                logger.debug(f"  -> Section set to: contract")
                continue
            if "salaris" in low and "mutatie" in low:
                current_section = "salary"
                logger.debug(f"  -> Section set to: salary")
                continue

        elif elem_type == "table":
            if current_employee is None:
                logger.debug("Table found but no employee set yet, skipping.")
                continue

            table = elem
            rows_data = []
            for row in table.rows:
                row_cells = [cell.text.strip() for cell in row.cells]
                rows_data.append(row_cells)

            if len(rows_data) < 2:
                logger.debug("Table with < 2 rows, skipping.")
                continue

            header = rows_data[0]
            logger.debug(f"Table header: {header}")

            emp = current_employee

            # --- Try to detect table type from column headers ---
            col_begin_c = find_col(header, "begin contract")
            col_einde_c = find_col(header, "einde contract")
            col_dv = find_col(header, "dienstverband")

            col_begin_s = find_col(header, "begindatum")
            col_einde_s = find_col(header, "einddatum")
            col_sal = find_col(header, "salaris")

            is_contract = col_begin_c is not None
            is_salary = col_begin_s is not None or col_sal is not None

            # If headers don't clearly identify the table, use the
            # current_section set by the preceding paragraph.
            if not is_contract and not is_salary:
                if current_section == "contract":
                    is_contract = True
                elif current_section == "salary":
                    is_salary = True
                else:
                    logger.debug(f"Unrecognised table, skipping: {header}")
                    continue

            if is_contract:
                count = 0
                for row_cells in rows_data[1:]:
                    # Read begin contract
                    begin_val = row_cells[col_begin_c] if col_begin_c is not None and col_begin_c < len(row_cells) else ""
                    begin = parse_date(begin_val)
                    if begin is None:
                        # If no column index, try finding a date in any cell
                        if col_begin_c is None:
                            for c in row_cells:
                                begin = parse_date(c)
                                if begin:
                                    break
                        if begin is None:
                            continue

                    einde_val = row_cells[col_einde_c] if col_einde_c is not None and col_einde_c < len(row_cells) else ""
                    einde = parse_date(einde_val)

                    dv_val = row_cells[col_dv] if col_dv is not None and col_dv < len(row_cells) else ""
                    dv = dv_val.strip() or None

                    all_rows.append({
                        "Naam": emp,
                        "Begin contract": begin,
                        "Einde contract": einde,
                        "Dienstverband": dv,
                        "Begindatum": None,
                        "Einddatum": None,
                        "Salaris": None,
                    })
                    count += 1
                report(f"  {emp}: {count} contractregel(s) gevonden.")

            elif is_salary:
                count = 0
                for row_cells in rows_data[1:]:
                    begin_val = row_cells[col_begin_s] if col_begin_s is not None and col_begin_s < len(row_cells) else ""
                    begin = parse_date(begin_val)
                    if begin is None:
                        if col_begin_s is None:
                            for c in row_cells:
                                begin = parse_date(c)
                                if begin:
                                    break
                        if begin is None:
                            continue

                    einde_val = row_cells[col_einde_s] if col_einde_s is not None and col_einde_s < len(row_cells) else ""
                    einde = parse_date(einde_val)

                    sal_val = row_cells[col_sal] if col_sal is not None and col_sal < len(row_cells) else ""
                    salaris = parse_salary(sal_val)

                    all_rows.append({
                        "Naam": emp,
                        "Begin contract": None,
                        "Einde contract": None,
                        "Dienstverband": None,
                        "Begindatum": begin,
                        "Einddatum": einde,
                        "Salaris": salaris,
                    })
                    count += 1
                report(f"  {emp}: {count} salarisregel(s) gevonden.")

    report(f"\nTotaal: {len(all_rows)} rijen geëxtraheerd "
           f"voor {len(employee_set)} medewerker(s).")

    if not all_rows:
        report("WAARSCHUWING: Geen gegevens geëxtraheerd.")
        report(f"Controleer het logbestand voor details: {LOG_FILE}")

    return all_rows


# ---------------------------------------------------------------------------
# Excel Export
# ---------------------------------------------------------------------------

COLUMNS = [
    "Naam",
    "Begin contract",
    "Einde contract",
    "Dienstverband",
    "Begindatum",
    "Einddatum",
    "Salaris",
]

DATE_COLUMNS = {"Begin contract", "Einde contract", "Begindatum", "Einddatum"}


def write_excel(rows: list[dict], output_path: str):
    """Write the extracted data to an Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    header_font = Font(bold=True)
    header_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            value = row_data.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            if col_name in DATE_COLUMNS and value is not None:
                cell.number_format = "YYYY-MM-DD"
            elif col_name == "Salaris" and value is not None:
                cell.number_format = "#,##0.00"

    col_widths = {
        "Naam": 22, "Begin contract": 16, "Einde contract": 16,
        "Dienstverband": 18, "Begindatum": 14, "Einddatum": 14, "Salaris": 12,
    }
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[letter].width = col_widths.get(col_name, 15)

    wb.save(output_path)
    logger.info(f"Excel bestand opgeslagen: {output_path}")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class StamkaartApp:
    """Dutch-language tkinter GUI for the Word-to-Excel converter."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Stamkaart Word naar Excel")
        self.root.geometry("700x520")
        self.root.resizable(True, True)

        try:
            self.root.tk.call("tk", "scaling", 1.2)
        except tk.TclError:
            pass

        self.docx_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        main = tk.Frame(self.root, padx=15, pady=15)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main,
            text="Stamkaart Word naar Excel Converter",
            font=("Segoe UI", 14, "bold"),
        ).pack(pady=(0, 15))

        # Word input
        docx_frame = tk.LabelFrame(main, text="Invoer", padx=10, pady=8)
        docx_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(docx_frame, text="Word bestand:").grid(row=0, column=0, sticky="w")
        tk.Entry(docx_frame, textvariable=self.docx_path, width=55).grid(
            row=0, column=1, padx=5
        )
        tk.Button(docx_frame, text="Bladeren...", command=self._browse_docx).grid(
            row=0, column=2
        )

        # Excel output
        xlsx_frame = tk.LabelFrame(main, text="Uitvoer", padx=10, pady=8)
        xlsx_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(xlsx_frame, text="Excel bestand:").grid(row=0, column=0, sticky="w")
        tk.Entry(xlsx_frame, textvariable=self.xlsx_path, width=55).grid(
            row=0, column=1, padx=5
        )
        tk.Button(xlsx_frame, text="Bladeren...", command=self._browse_xlsx).grid(
            row=0, column=2
        )

        # Execute button
        btn_frame = tk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=8)
        self.run_btn = tk.Button(
            btn_frame, text="Uitvoeren", command=self._run,
            font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white",
            padx=20, pady=5,
        )
        self.run_btn.pack()

        # Status area
        status_frame = tk.LabelFrame(main, text="Status", padx=10, pady=8)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.status_text = scrolledtext.ScrolledText(
            status_frame, height=12, state=tk.DISABLED, wrap=tk.WORD,
            font=("Consolas", 9),
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)

    def _browse_docx(self):
        path = filedialog.askopenfilename(
            title="Selecteer Word bestand",
            filetypes=[
                ("Word bestanden", "*.docx"),
                ("Alle bestanden", "*.*"),
            ],
        )
        if path:
            self.docx_path.set(path)
            if not self.xlsx_path.get():
                self.xlsx_path.set(os.path.splitext(path)[0] + "_export.xlsx")

    def _browse_xlsx(self):
        path = filedialog.asksaveasfilename(
            title="Opslaan als Excel bestand",
            defaultextension=".xlsx",
            filetypes=[("Excel bestanden", "*.xlsx"), ("Alle bestanden", "*.*")],
        )
        if path:
            self.xlsx_path.set(path)

    def _log(self, message: str):
        def _update():
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
        self.root.after(0, _update)

    def _run(self):
        docx = self.docx_path.get().strip()
        xlsx = self.xlsx_path.get().strip()

        if not docx:
            messagebox.showwarning("Waarschuwing", "Selecteer eerst een Word bestand.")
            return
        if not os.path.isfile(docx):
            messagebox.showerror("Fout", f"Word bestand niet gevonden:\n{docx}")
            return
        if not xlsx:
            messagebox.showwarning("Waarschuwing", "Kies een locatie voor het Excel bestand.")
            return

        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete("1.0", tk.END)
        self.status_text.config(state=tk.DISABLED)

        self.run_btn.config(state=tk.DISABLED, text="Bezig...")
        threading.Thread(target=self._extract, args=(docx, xlsx), daemon=True).start()

    def _extract(self, docx_path: str, xlsx_path: str):
        try:
            rows = process_docx(docx_path, progress_callback=self._log)

            if not rows:
                self._log("\nGeen gegevens gevonden in het document.")
                self._log(f"Controleer het debug logbestand: {LOG_FILE}")
                self.root.after(0, lambda: messagebox.showwarning(
                    "Waarschuwing", "Geen gegevens gevonden. Controleer het logbestand.",
                ))
                return

            write_excel(rows, xlsx_path)
            self._log(f"\nExcel bestand opgeslagen: {xlsx_path}")
            self._log("Klaar!")
            self.root.after(0, lambda: messagebox.showinfo(
                "Gereed",
                f"Export voltooid!\n\n{len(rows)} rijen geschreven naar:\n{xlsx_path}",
            ))

        except Exception as e:
            logger.error(f"Error during extraction:\n{traceback.format_exc()}")
            self._log(f"\nFOUT: {e}")
            self._log(f"\nDetails in logbestand: {LOG_FILE}")
            self.root.after(0, lambda: messagebox.showerror(
                "Fout", f"Er is een fout opgetreden:\n{e}\n\nZie logbestand.",
            ))

        finally:
            self.root.after(
                0, lambda: self.run_btn.config(state=tk.NORMAL, text="Uitvoeren")
            )


def main():
    root = tk.Tk()
    StamkaartApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
