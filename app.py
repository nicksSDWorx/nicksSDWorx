"""
Stamkaart Word naar Excel Converter

Extracts employee contract and salary history from a Dutch HR Word
document ("stamkaarten") and exports them to a structured Excel file.

Walks every element in document order.  Paragraphs containing
"Stamkaart van <Name>" set the current employee.  EVERY table is
scanned for Contract and Salaris sub-sections regardless of size
or structure.

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

DATE_RE = re.compile(r"\d{2}-\d{2}-\d{4}")


def parse_date(text: str):
    """Try to parse a date string (dd-mm-yyyy) into a datetime."""
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


def extract_date_from_text(text: str):
    """Extract the first dd-mm-yyyy date found anywhere in the text."""
    m = DATE_RE.search(text or "")
    if m:
        return parse_date(m.group())
    return None


def parse_salary(text: str):
    """Parse a salary string like '2.389,44' into a float."""
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


# ---------------------------------------------------------------------------
# Document iteration
# ---------------------------------------------------------------------------

def iter_doc_elements(doc: Document):
    """Yield (type, element) in document order."""
    for child in doc.element.body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield ("paragraph", Paragraph(child, doc))
        elif tag == "tbl":
            yield ("table", Table(child, doc))


# ---------------------------------------------------------------------------
# Employee name detection
# ---------------------------------------------------------------------------

def extract_employee_name(text: str) -> str | None:
    """
    If the paragraph text represents a "Stamkaart van <Name>" header,
    return the name.  Returns None for footers (which have a leading
    timestamp) and any non-matching text.
    """
    text = text.strip()
    # Must START with "Stamkaart van" — footers have a timestamp first
    m = re.match(r"stamkaart\s+van\s+(.+)", text, re.IGNORECASE)
    if m:
        name = m.group(1).strip()
        # Reject if the remainder looks like a footer (contains digits
        # that could be a page number right after the name, tab-separated)
        if "\t" in name:
            # e.g. "Name\t2" — the tab+number is a page number, strip it
            name = name.split("\t")[0].strip()
        return name if name else None
    return None


# ---------------------------------------------------------------------------
# Table parsing — section detection within a table
# ---------------------------------------------------------------------------

def _detect_section(cell0_low: str) -> str | None:
    """
    Given the lowercased text of cell[0] of a row, decide if this row
    is a sub-section marker/header.  Returns:
      "contract"  — contract header row
      "rooster"   — rooster section (skip)
      "oe"        — OE/functie section (skip)
      "salary"    — salary header row
      None        — not a section marker
    """
    # Contract header: contains "begin contract" or starts with "werkgever"
    if "begin contract" in cell0_low or "werkgever" in cell0_low:
        return "contract"

    # Salary: contains "salaris" and ("mutatie" or "begindatum")
    if "salaris" in cell0_low and ("mutatie" in cell0_low or "begindatum" in cell0_low):
        return "salary"

    # Rooster: contains "rooster"
    if "rooster" in cell0_low:
        return "rooster"

    # OE/Functie: contains "functie" and "mutatie"
    if "functie" in cell0_low and "mutatie" in cell0_low:
        return "oe"

    return None


def parse_table(table: Table, employee_name: str) -> tuple[list[dict], int, int]:
    """
    Parse a single table, looking for contract and salary sub-sections.

    Returns (rows, contract_count, salary_count).
    """
    rows_data = []
    for row in table.rows:
        cells = [cell.text for cell in row.cells]
        rows_data.append(cells)

    if len(rows_data) < 2:
        return [], 0, 0

    result: list[dict] = []
    section: str | None = None   # "contract", "salary", "rooster", "oe", or None
    is_header_row = False        # True = the current row IS the section header
    contract_count = 0
    salary_count = 0

    for i, cells in enumerate(rows_data):
        cell0 = cells[0].strip() if cells else ""
        cell0_low = cell0.lower()

        # --- Check if this row is a section marker ---
        detected = _detect_section(cell0_low)
        if detected is not None:
            section = detected
            is_header_row = True
            debug_cells = [c.replace(chr(10), "\\n").replace(chr(9), "\\t")[:60] for c in cells]
            logger.debug(f"  Row {i}: section={section} header | {debug_cells}")
            continue

        # The row right after a header may also be a column-labels row
        # (for rooster/OE sections).  Skip rows that contain "begindatum"
        # or "einddatum" in cell0 when in a skip-section.
        if section in ("rooster", "oe"):
            if "begindatum" in cell0_low or "einddatum" in cell0_low:
                continue
            # Data rows in skipped sections — just skip them
            if DATE_RE.search(cell0):
                continue
            # If no date, might be a new section marker we didn't catch
            # Fall through to let _detect_section catch it next iteration

        # --- Contract data rows ---
        if section == "contract":
            # cell[0] or cell[1] contains "CompanyName" + date smashed together
            begin = extract_date_from_text(cell0)
            if begin is None and len(cells) > 1:
                begin = extract_date_from_text(cells[1])
            if begin is None:
                # Not a data row — might be a sub-header or empty row
                continue

            einde = parse_date(cells[2].strip()) if len(cells) > 2 else None
            dv = cells[3].strip() if len(cells) > 3 and cells[3].strip() else None

            result.append({
                "Naam": employee_name,
                "Begin contract": begin,
                "Einde contract": einde,
                "Dienstverband": dv,
                "Begindatum": None,
                "Einddatum": None,
                "Salaris": None,
            })
            contract_count += 1

        # --- Salary data rows ---
        elif section == "salary":
            # cell[0] contains "dd-mm-yyyy\tdd-mm-yyyy" (tab-separated dates)
            dates_in_cell = DATE_RE.findall(cell0)
            if not dates_in_cell:
                continue

            begindatum = parse_date(dates_in_cell[0])
            einddatum = parse_date(dates_in_cell[1]) if len(dates_in_cell) >= 2 else None
            sal_text = cells[1].strip() if len(cells) > 1 else ""
            salaris = parse_salary(sal_text)

            if begindatum is None:
                continue

            result.append({
                "Naam": employee_name,
                "Begin contract": None,
                "Einde contract": None,
                "Dienstverband": None,
                "Begindatum": begindatum,
                "Einddatum": einddatum,
                "Salaris": salaris,
            })
            salary_count += 1

    return result, contract_count, salary_count


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process_docx(docx_path: str, progress_callback=None) -> list[dict]:
    """
    Walk the Word document in element order.
    Paragraphs set the current employee; every table is scanned for data.
    """

    def report(msg):
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    report(f"Document openen: {docx_path}")

    doc = Document(docx_path)
    all_rows: list[dict] = []
    current_employee: str | None = None
    employee_count = 0
    table_idx = 0

    for elem_type, elem in iter_doc_elements(doc):

        if elem_type == "paragraph":
            text = elem.text.strip()
            if not text:
                continue

            name = extract_employee_name(text)
            if name:
                current_employee = name
                employee_count += 1
                report(f"Medewerker gevonden: {name}")
                logger.debug(f"Employee set to: {name!r} (from paragraph: {text!r})")

        elif elem_type == "table":
            table_idx += 1
            nrows = len(elem.rows)
            logger.debug(f"Table #{table_idx}: {nrows} rows")

            if current_employee is None:
                logger.debug(f"  Skipped (no employee set yet)")
                continue

            rows, cc, sc = parse_table(elem, current_employee)

            if rows:
                all_rows.extend(rows)
                report(f"  {current_employee}: {cc} contractregel(s), "
                       f"{sc} salarisregel(s)")
            else:
                logger.debug(f"  No contract/salary data found in this table")

    report(f"\nTotaal: {len(all_rows)} rijen geëxtraheerd "
           f"voor {employee_count} medewerker(s).")

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
            filetypes=[("Word bestanden", "*.docx"), ("Alle bestanden", "*.*")],
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
