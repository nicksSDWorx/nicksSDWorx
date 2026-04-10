"""
Stamkaart PDF naar Excel Converter

Extracts employee contract and salary history from Dutch HR PDF
("stamkaarten") and exports them to a structured Excel file.

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
    """Try to parse a Dutch date string into a datetime object."""
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
    text = text.strip().replace(" ", "").replace("\u20ac", "")  # strip € sign
    if not text or text == "-":
        return None
    # Dutch format: 2.389,44  ->  remove dots, replace comma with dot
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return round(float(text), 2)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# PDF Parsing — employee block splitting
# ---------------------------------------------------------------------------

def extract_pages_text(pdf_path: str) -> list[str]:
    """Extract text from each page of the PDF using pdfplumber."""
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            logger.debug(f"=== PAGE {i + 1} RAW TEXT ===\n{text}\n{'=' * 60}")
            pages_text.append(text)
    return pages_text


def split_into_employee_blocks(full_text: str) -> list[tuple[str, str]]:
    """
    Split the full text into (employee_name, block_text) tuples.
    Each employee section starts with a line containing "Stamkaart van <Name>".
    """
    lines = full_text.split("\n")

    # Find every line that contains "Stamkaart van"
    block_starts: list[tuple[int, str]] = []
    for i, line in enumerate(lines):
        m = re.search(r"stamkaart\s+van\s+(.+)", line, re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            block_starts.append((i, name))

    if not block_starts:
        logger.warning("No 'Stamkaart van' headers found in the PDF text.")
        return []

    blocks: list[tuple[str, str]] = []
    for idx, (start, name) in enumerate(block_starts):
        end = block_starts[idx + 1][0] if idx + 1 < len(block_starts) else len(lines)
        block_text = "\n".join(lines[start:end])
        blocks.append((name, block_text))
        logger.debug(
            f"Employee block [{idx}]: '{name}' — lines {start}..{end} "
            f"({end - start} lines)"
        )

    return blocks


# ---------------------------------------------------------------------------
# PDF Parsing — section detection
# ---------------------------------------------------------------------------

# Section header keywords (lowercased). Order matters: first match wins.
CONTRACT_KEYWORDS = ["contract mutaties", "contract mutatie", "contract historie"]
SALARY_KEYWORDS = ["salaris mutaties", "salaris mutatie", "salaris historie"]

# Generic markers that signal the START of a new section (used to detect
# where the current section ends).
SECTION_BOUNDARY_RE = re.compile(
    r"(mutaties|mutatie|historie|overzicht|gegevens|stamkaart\s+van)",
    re.IGNORECASE,
)


def find_section(text: str, keywords: list[str]) -> str | None:
    """
    Return the text from the first matching section header down to the next
    section boundary (or end of text).
    """
    lines = text.split("\n")
    start_idx: int | None = None

    for i, line in enumerate(lines):
        low = line.lower()
        for kw in keywords:
            if kw in low:
                start_idx = i
                break
        if start_idx is not None:
            break

    if start_idx is None:
        return None

    # Walk forward to find where this section ends
    end_idx = len(lines)
    for i in range(start_idx + 1, len(lines)):
        stripped = lines[i].strip()
        if not stripped:
            continue
        # A new section header is a short-ish line that matches the boundary
        # regex AND has fewer than 4 digit characters (so data rows don't
        # accidentally match).
        if SECTION_BOUNDARY_RE.search(stripped):
            digit_count = sum(c.isdigit() for c in stripped)
            if digit_count < 4:
                end_idx = i
                break

    section = "\n".join(lines[start_idx:end_idx])
    logger.debug(f"Section found (lines {start_idx}..{end_idx}):\n{section}")
    return section


# ---------------------------------------------------------------------------
# PDF Parsing — column-header-aware row extraction
# ---------------------------------------------------------------------------

def locate_columns(header_line: str, col_names: list[str]) -> list[tuple[str, int]]:
    """
    Given a header line and a list of expected column names (lowercased),
    return [(col_name, char_position), ...] sorted by position.

    Matching is case-insensitive and allows partial overlap
    (e.g. "begindatum" matches "begindatum").
    """
    low = header_line.lower()
    found: list[tuple[str, int]] = []
    for cn in col_names:
        pos = low.find(cn)
        if pos != -1:
            found.append((cn, pos))
    found.sort(key=lambda x: x[1])
    return found


def slice_row_by_columns(
    line: str,
    col_positions: list[tuple[str, int]],
) -> dict[str, str]:
    """
    Slice a text line into fields based on column start positions.
    Each field runs from its start position to the start of the next column
    (or end of line for the last column).
    """
    result: dict[str, str] = {}
    for idx, (col_name, start) in enumerate(col_positions):
        if idx + 1 < len(col_positions):
            end = col_positions[idx + 1][1]
        else:
            end = len(line)
        value = line[start:end].strip() if start < len(line) else ""
        result[col_name] = value
    return result


def parse_contract_section(section_text: str) -> list[dict]:
    """
    Parse contract rows from a "Contract mutaties" section.

    Looks for columns: begin contract, einde contract, dienstverband.
    """
    lines = section_text.split("\n")

    # --- Locate the column header line ---
    target_cols = ["begin contract", "einde contract", "dienstverband"]
    col_positions: list[tuple[str, int]] = []
    header_line_idx: int | None = None

    for i, line in enumerate(lines):
        cols = locate_columns(line, target_cols)
        # Accept if we find at least 2 of the 3 expected columns
        if len(cols) >= 2:
            col_positions = cols
            header_line_idx = i
            logger.debug(
                f"Contract header at line {i}: {cols} | raw: {line!r}"
            )
            break

    if header_line_idx is None:
        # Fallback: try to parse lines that contain dates
        logger.warning("Could not find contract column headers; using fallback.")
        return _fallback_parse_contract(lines)

    # --- Parse data rows below the header ---
    rows: list[dict] = []
    date_re = re.compile(r"\d{2}[-/]\d{2}[-/]\d{4}")

    for line in lines[header_line_idx + 1 :]:
        stripped = line.strip()
        if not stripped:
            continue
        # A data row must contain at least one date
        if not date_re.search(stripped):
            continue

        fields = slice_row_by_columns(line, col_positions)

        begin = parse_date(fields.get("begin contract", ""))
        einde = parse_date(fields.get("einde contract", ""))
        dv = fields.get("dienstverband", "").strip() or None

        if begin is not None:
            rows.append(
                {"begin_contract": begin, "einde_contract": einde, "dienstverband": dv}
            )

    return rows


def _fallback_parse_contract(lines: list[str]) -> list[dict]:
    """Date-based fallback when column headers are not detected."""
    rows: list[dict] = []
    date_re = re.compile(r"\d{2}[-/]\d{2}[-/]\d{4}")

    for line in lines:
        stripped = line.strip()
        dates = date_re.findall(stripped)
        if not dates:
            continue

        parts = re.split(r"\s{2,}|\t", stripped)
        parts = [p.strip() for p in parts if p.strip()]

        dates_parsed = [parse_date(p) for p in parts]
        texts = [p for p, d in zip(parts, dates_parsed) if d is None]
        dates_found = [d for d in dates_parsed if d is not None]

        if dates_found:
            rows.append({
                "begin_contract": dates_found[0],
                "einde_contract": dates_found[1] if len(dates_found) >= 2 else None,
                "dienstverband": texts[0] if texts else None,
            })
    return rows


def parse_salary_section(section_text: str) -> list[dict]:
    """
    Parse salary rows from a "Salaris mutaties" section.

    Looks for columns: begindatum, einddatum, salaris.
    """
    lines = section_text.split("\n")

    # --- Locate the column header line ---
    target_cols = ["begindatum", "einddatum", "salaris"]
    col_positions: list[tuple[str, int]] = []
    header_line_idx: int | None = None

    for i, line in enumerate(lines):
        cols = locate_columns(line, target_cols)
        if len(cols) >= 2:
            col_positions = cols
            header_line_idx = i
            logger.debug(
                f"Salary header at line {i}: {cols} | raw: {line!r}"
            )
            break

    if header_line_idx is None:
        logger.warning("Could not find salary column headers; using fallback.")
        return _fallback_parse_salary(lines)

    # --- Parse data rows below the header ---
    rows: list[dict] = []
    date_re = re.compile(r"\d{2}[-/]\d{2}[-/]\d{4}")

    for line in lines[header_line_idx + 1 :]:
        stripped = line.strip()
        if not stripped:
            continue
        if not date_re.search(stripped):
            continue

        fields = slice_row_by_columns(line, col_positions)

        begin = parse_date(fields.get("begindatum", ""))
        einde = parse_date(fields.get("einddatum", ""))
        salaris = parse_salary(fields.get("salaris", ""))

        if begin is not None:
            rows.append(
                {"begindatum": begin, "einddatum": einde, "salaris": salaris}
            )

    return rows


def _fallback_parse_salary(lines: list[str]) -> list[dict]:
    """Date-based fallback when column headers are not detected."""
    rows: list[dict] = []
    date_re = re.compile(r"\d{2}[-/]\d{2}[-/]\d{4}")

    for line in lines:
        stripped = line.strip()
        dates = date_re.findall(stripped)
        if not dates:
            continue

        parts = re.split(r"\s{2,}|\t", stripped)
        parts = [p.strip() for p in parts if p.strip()]

        dates_found: list = []
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
            rows.append({
                "begindatum": dates_found[0],
                "einddatum": dates_found[1] if len(dates_found) >= 2 else None,
                "salaris": salary_val,
            })
    return rows


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process_pdf(pdf_path: str, progress_callback=None) -> list[dict]:
    """
    Main entry: extract all employee data from the PDF.
    Returns a list of row dicts ready for Excel export.
    """

    def report(msg):
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    report(f"PDF openen: {pdf_path}")

    pages = extract_pages_text(pdf_path)
    report(f"{len(pages)} pagina('s) gevonden.")

    full_text = "\n".join(pages)

    employee_blocks = split_into_employee_blocks(full_text)
    report(f"{len(employee_blocks)} medewerker(s) gevonden.")

    if not employee_blocks:
        report("WAARSCHUWING: Geen 'Stamkaart van' headers gevonden in de PDF.")
        return []

    all_rows: list[dict] = []
    warnings: list[str] = []

    for emp_name, block_text in employee_blocks:
        report(f"Verwerken: {emp_name}")

        # ---- Contract mutaties ----
        contract_section = find_section(block_text, CONTRACT_KEYWORDS)
        if contract_section:
            contract_rows = parse_contract_section(contract_section)
            report(f"  {len(contract_rows)} contractregel(s) gevonden.")
            for crow in contract_rows:
                all_rows.append({
                    "Naam": emp_name,
                    "Begin contract": crow["begin_contract"],
                    "Einde contract": crow["einde_contract"],
                    "Dienstverband": crow["dienstverband"],
                    "Begindatum": None,
                    "Einddatum": None,
                    "Salaris": None,
                })
        else:
            msg = f"  WAARSCHUWING: Geen 'Contract mutaties' gevonden voor {emp_name}."
            report(msg)
            warnings.append(msg)

        # ---- Salaris mutaties ----
        salary_section = find_section(block_text, SALARY_KEYWORDS)
        if salary_section:
            salary_rows = parse_salary_section(salary_section)
            report(f"  {len(salary_rows)} salarisregel(s) gevonden.")
            for srow in salary_rows:
                all_rows.append({
                    "Naam": emp_name,
                    "Begin contract": None,
                    "Einde contract": None,
                    "Dienstverband": None,
                    "Begindatum": srow["begindatum"],
                    "Einddatum": srow["einddatum"],
                    "Salaris": srow["salaris"],
                })
        else:
            msg = f"  WAARSCHUWING: Geen 'Salaris mutaties' gevonden voor {emp_name}."
            report(msg)
            warnings.append(msg)

    report(f"\nTotaal: {len(all_rows)} rijen geëxtraheerd.")
    if warnings:
        report(f"\n{len(warnings)} waarschuwing(en):")
        for w in warnings:
            report(f"  {w}")

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

    # Headers
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    # Data rows
    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            value = row_data.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            if col_name in DATE_COLUMNS and value is not None:
                cell.number_format = "YYYY-MM-DD"
            elif col_name == "Salaris" and value is not None:
                cell.number_format = "#,##0.00"

    # Column widths
    col_widths = {
        "Naam": 22,
        "Begin contract": 16,
        "Einde contract": 16,
        "Dienstverband": 18,
        "Begindatum": 14,
        "Einddatum": 14,
        "Salaris": 12,
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

        # --- PDF input ---
        pdf_frame = tk.LabelFrame(main, text="Invoer", padx=10, pady=8)
        pdf_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(pdf_frame, text="PDF bestand:").grid(row=0, column=0, sticky="w")
        tk.Entry(pdf_frame, textvariable=self.pdf_path, width=55).grid(
            row=0, column=1, padx=5
        )
        tk.Button(pdf_frame, text="Bladeren...", command=self._browse_pdf).grid(
            row=0, column=2
        )

        # --- Excel output ---
        xlsx_frame = tk.LabelFrame(main, text="Uitvoer", padx=10, pady=8)
        xlsx_frame.pack(fill=tk.X, pady=(0, 8))
        tk.Label(xlsx_frame, text="Excel bestand:").grid(row=0, column=0, sticky="w")
        tk.Entry(xlsx_frame, textvariable=self.xlsx_path, width=55).grid(
            row=0, column=1, padx=5
        )
        tk.Button(xlsx_frame, text="Bladeren...", command=self._browse_xlsx).grid(
            row=0, column=2
        )

        # --- Execute button ---
        btn_frame = tk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=8)
        self.run_btn = tk.Button(
            btn_frame,
            text="Uitvoeren",
            command=self._run,
            font=("Segoe UI", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=5,
        )
        self.run_btn.pack()

        # --- Status area ---
        status_frame = tk.LabelFrame(main, text="Status", padx=10, pady=8)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.status_text = scrolledtext.ScrolledText(
            status_frame,
            height=12,
            state=tk.DISABLED,
            wrap=tk.WORD,
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
        """Append a message to the status area (thread-safe)."""

        def _update():
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)

        self.root.after(0, _update)

    def _run(self):
        """Start extraction in a background thread."""
        pdf = self.pdf_path.get().strip()
        xlsx = self.xlsx_path.get().strip()

        if not pdf:
            messagebox.showwarning("Waarschuwing", "Selecteer eerst een PDF bestand.")
            return
        if not os.path.isfile(pdf):
            messagebox.showerror("Fout", f"PDF bestand niet gevonden:\n{pdf}")
            return
        if not xlsx:
            messagebox.showwarning(
                "Waarschuwing", "Kies een locatie voor het Excel bestand."
            )
            return

        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete("1.0", tk.END)
        self.status_text.config(state=tk.DISABLED)

        self.run_btn.config(state=tk.DISABLED, text="Bezig...")
        threading.Thread(target=self._extract, args=(pdf, xlsx), daemon=True).start()

    def _extract(self, pdf_path: str, xlsx_path: str):
        """Run the extraction (background thread)."""
        try:
            rows = process_pdf(pdf_path, progress_callback=self._log)

            if not rows:
                self._log("\nGeen gegevens gevonden in de PDF.")
                self._log("Controleer het debug logbestand voor meer informatie:")
                self._log(f"  {LOG_FILE}")
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "Waarschuwing",
                        "Geen gegevens gevonden. Controleer het logbestand.",
                    ),
                )
                return

            write_excel(rows, xlsx_path)
            self._log(f"\nExcel bestand opgeslagen: {xlsx_path}")
            self._log("Klaar!")
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Gereed",
                    f"Export voltooid!\n\n"
                    f"{len(rows)} rijen geschreven naar:\n{xlsx_path}",
                ),
            )

        except Exception as e:
            logger.error(f"Error during extraction:\n{traceback.format_exc()}")
            self._log(f"\nFOUT: {e}")
            self._log(f"\nDetails in logbestand: {LOG_FILE}")
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Fout",
                    f"Er is een fout opgetreden:\n{e}\n\nZie logbestand.",
                ),
            )

        finally:
            self.root.after(
                0, lambda: self.run_btn.config(state=tk.NORMAL, text="Uitvoeren")
            )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    StamkaartApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
