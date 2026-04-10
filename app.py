"""
Stamkaart PDF naar Excel Converter

Extracts employee contract and salary history from Dutch HR PDF
("stamkaarten") and exports them to a structured Excel file.

Uses pdfplumber's table extraction for reliable column parsing.

Usage: Double-click the .exe or run `python app.py`
"""

import os
import re
import sys
import logging
import traceback
import threading
from datetime import datetime

import pdfplumber
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


def clean_cell(val) -> str:
    """Normalize a table cell value to a string."""
    if val is None:
        return ""
    return str(val).strip()


def normalize(text: str) -> str:
    """Lowercase + collapse whitespace for header matching."""
    return re.sub(r"\s+", " ", text.strip().lower())


# ---------------------------------------------------------------------------
# PDF extraction — page-by-page with table extraction
# ---------------------------------------------------------------------------

def find_header_index(header_cells: list[str], target: str) -> int | None:
    """
    Find the column index in a header row whose normalized text contains
    the target string.  Returns None if not found.
    """
    for i, cell in enumerate(header_cells):
        if target in normalize(cell):
            return i
    return None


def process_pdf(pdf_path: str, progress_callback=None) -> list[dict]:
    """
    Main processing function.

    Strategy:
      1. Walk through the PDF page by page.
      2. Use extract_text() to detect "Stamkaart van <Name>" markers and
         section headers ("Contract mutaties", "Salaris mutaties").
      3. Use extract_tables() to get structured table data from each page.
      4. Match each table to a section by checking its header row for known
         column names — this avoids all character-position guessing.
    """

    def report(msg):
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    report(f"PDF openen: {pdf_path}")

    all_rows: list[dict] = []
    warnings: list[str] = []
    current_employee: str | None = None
    employee_set: set[str] = set()

    with pdfplumber.open(pdf_path) as pdf:
        report(f"{len(pdf.pages)} pagina('s) gevonden.")

        for page_num, page in enumerate(pdf.pages, start=1):
            # --- Text: detect employees and sections ---
            text = page.extract_text() or ""
            logger.debug(
                f"=== PAGE {page_num} RAW TEXT ===\n{text}\n{'=' * 60}"
            )

            # Find all "Stamkaart van <Name>" on this page.
            # The LAST one seen before / around a table is the active employee.
            for m in re.finditer(
                r"stamkaart\s+van\s+(.+)", text, re.IGNORECASE
            ):
                name = m.group(1).strip()
                current_employee = name
                if name not in employee_set:
                    employee_set.add(name)
                    report(f"Medewerker gevonden: {name}")

            # --- Tables: extract structured data ---
            tables = page.extract_tables() or []
            logger.debug(
                f"Page {page_num}: {len(tables)} table(s) found"
            )

            for t_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue  # need at least a header + one data row

                # The first row is assumed to be column headers
                raw_header = [clean_cell(c) for c in table[0]]
                norm_header = [normalize(c) for c in raw_header]
                logger.debug(
                    f"  Table {t_idx} headers: {raw_header}"
                )

                # --- Try to identify this table as CONTRACT ---
                col_begin_c = find_header_index(raw_header, "begin contract")
                col_einde_c = find_header_index(raw_header, "einde contract")
                col_dv = find_header_index(raw_header, "dienstverband")

                is_contract_table = col_begin_c is not None

                # --- Try to identify this table as SALARY ---
                col_begin_s = find_header_index(raw_header, "begindatum")
                col_einde_s = find_header_index(raw_header, "einddatum")
                col_sal = find_header_index(raw_header, "salaris")

                is_salary_table = col_begin_s is not None and col_sal is not None

                emp = current_employee or "Onbekend"

                if is_contract_table:
                    count = 0
                    for row in table[1:]:
                        cells = [clean_cell(c) for c in row]
                        begin = parse_date(
                            cells[col_begin_c] if col_begin_c < len(cells) else ""
                        )
                        if begin is None:
                            continue  # skip non-data rows
                        einde = parse_date(
                            cells[col_einde_c] if col_einde_c is not None and col_einde_c < len(cells) else ""
                        )
                        dv = (
                            cells[col_dv] if col_dv is not None and col_dv < len(cells) else None
                        ) or None

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

                elif is_salary_table:
                    count = 0
                    for row in table[1:]:
                        cells = [clean_cell(c) for c in row]
                        begin = parse_date(
                            cells[col_begin_s] if col_begin_s < len(cells) else ""
                        )
                        if begin is None:
                            continue
                        einde = parse_date(
                            cells[col_einde_s] if col_einde_s is not None and col_einde_s < len(cells) else ""
                        )
                        salaris = parse_salary(
                            cells[col_sal] if col_sal is not None and col_sal < len(cells) else ""
                        )

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

                else:
                    logger.debug(
                        f"  Table {t_idx} on page {page_num} not recognised — "
                        f"headers: {raw_header}"
                    )

    # --- Fallback: text-based parsing if no tables were extracted ---
    if not all_rows:
        report("Geen tabellen gevonden via tabelextractie. Probeer tekst-gebaseerde extractie...")
        all_rows, extra_warnings = _text_based_fallback(pdf_path, progress_callback=report)
        warnings.extend(extra_warnings)

    report(f"\nTotaal: {len(all_rows)} rijen geëxtraheerd.")

    if not all_rows:
        report("WAARSCHUWING: Geen gegevens geëxtraheerd.")
        report(f"Controleer het logbestand voor details: {LOG_FILE}")

    if warnings:
        for w in warnings:
            report(f"  {w}")

    return all_rows


# ---------------------------------------------------------------------------
# Fallback: text-based parsing (if extract_tables() yields nothing)
# ---------------------------------------------------------------------------

def _text_based_fallback(pdf_path: str, progress_callback=None) -> tuple[list[dict], list[str]]:
    """
    Fallback parser that works on raw extracted text line-by-line.
    Used when pdfplumber's table detector finds no tables.
    """
    pages_text: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_text.append(page.extract_text() or "")

    full_text = "\n".join(pages_text)
    lines = full_text.split("\n")

    # Split into employee blocks
    block_starts: list[tuple[int, str]] = []
    for i, line in enumerate(lines):
        m = re.search(r"stamkaart\s+van\s+(.+)", line, re.IGNORECASE)
        if m:
            block_starts.append((i, m.group(1).strip()))

    if not block_starts:
        return [], ["Geen 'Stamkaart van' gevonden in de PDF."]

    all_rows: list[dict] = []
    warnings: list[str] = []

    for idx, (start, emp_name) in enumerate(block_starts):
        end = block_starts[idx + 1][0] if idx + 1 < len(block_starts) else len(lines)
        block_lines = lines[start:end]

        if progress_callback:
            progress_callback(f"  (tekst-fallback) Verwerken: {emp_name}")

        # Identify sections within the block
        section = None  # "contract" or "salary"
        date_re = re.compile(r"\d{2}[-/]\d{2}[-/]\d{4}")

        for line in block_lines:
            low = line.lower()

            # Detect section headers
            if "contract" in low and "mutatie" in low:
                section = "contract"
                continue
            if "salaris" in low and "mutatie" in low:
                section = "salary"
                continue

            # Skip non-data lines (no dates)
            if not date_re.search(line):
                continue

            # Split by 2+ spaces
            parts = re.split(r"\s{2,}|\t", line.strip())
            parts = [p.strip() for p in parts if p.strip()]

            if section == "contract" and parts:
                dates = [parse_date(p) for p in parts]
                texts = [p for p, d in zip(parts, dates) if d is None]
                dates_found = [d for d in dates if d is not None]
                if dates_found:
                    all_rows.append({
                        "Naam": emp_name,
                        "Begin contract": dates_found[0],
                        "Einde contract": dates_found[1] if len(dates_found) >= 2 else None,
                        "Dienstverband": texts[0] if texts else None,
                        "Begindatum": None,
                        "Einddatum": None,
                        "Salaris": None,
                    })

            elif section == "salary" and parts:
                dates_found = []
                salary_val = None
                for p in parts:
                    d = parse_date(p)
                    if d:
                        dates_found.append(d)
                    else:
                        s = parse_salary(p)
                        if s is not None:
                            salary_val = s
                if dates_found:
                    all_rows.append({
                        "Naam": emp_name,
                        "Begin contract": None,
                        "Einde contract": None,
                        "Dienstverband": None,
                        "Begindatum": dates_found[0],
                        "Einddatum": dates_found[1] if len(dates_found) >= 2 else None,
                        "Salaris": salary_val,
                    })

    return all_rows, warnings


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
    """Dutch-language tkinter GUI for the PDF-to-Excel converter."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Stamkaart PDF naar Excel")
        self.root.geometry("700x520")
        self.root.resizable(True, True)

        try:
            self.root.tk.call("tk", "scaling", 1.2)
        except tk.TclError:
            pass

        self.pdf_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        main = tk.Frame(self.root, padx=15, pady=15)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main,
            text="Stamkaart PDF naar Excel Converter",
            font=("Segoe UI", 14, "bold"),
        ).pack(pady=(0, 15))

        # PDF input
        pdf_frame = tk.LabelFrame(main, text="Invoer", padx=10, pady=8)
        pdf_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(pdf_frame, text="PDF bestand:").grid(row=0, column=0, sticky="w")
        tk.Entry(pdf_frame, textvariable=self.pdf_path, width=55).grid(
            row=0, column=1, padx=5
        )
        tk.Button(pdf_frame, text="Bladeren...", command=self._browse_pdf).grid(
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

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Selecteer PDF bestand",
            filetypes=[("PDF bestanden", "*.pdf"), ("Alle bestanden", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
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
        pdf = self.pdf_path.get().strip()
        xlsx = self.xlsx_path.get().strip()

        if not pdf:
            messagebox.showwarning("Waarschuwing", "Selecteer eerst een PDF bestand.")
            return
        if not os.path.isfile(pdf):
            messagebox.showerror("Fout", f"PDF bestand niet gevonden:\n{pdf}")
            return
        if not xlsx:
            messagebox.showwarning("Waarschuwing", "Kies een locatie voor het Excel bestand.")
            return

        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete("1.0", tk.END)
        self.status_text.config(state=tk.DISABLED)

        self.run_btn.config(state=tk.DISABLED, text="Bezig...")
        threading.Thread(target=self._extract, args=(pdf, xlsx), daemon=True).start()

    def _extract(self, pdf_path: str, xlsx_path: str):
        try:
            rows = process_pdf(pdf_path, progress_callback=self._log)

            if not rows:
                self._log("\nGeen gegevens gevonden in de PDF.")
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
