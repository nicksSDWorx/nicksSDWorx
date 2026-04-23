"""
Builder for Payroll Discovery Document MKB (v2, Dutch-only).
Phased build.
Phase 3 status: Start + '1. Bedrijf & Contact' sheets filled.
"""
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

OUTPUT = "/home/user/nicksSDWorx/outputs/Payroll_Discovery_Document_MKB.xlsx"

# --- Styling constants ---
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
DEFAULT_FILL = PatternFill("solid", fgColor="F2F2F2")

HEADER_FONT = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
SUBHEADER_FONT = Font(name="Calibri", size=12, bold=True, color="1F3864")
BASE_FONT = Font(name="Calibri", size=11)
LABEL_FONT = Font(name="Calibri", size=11)
REQ_FONT = Font(name="Calibri", size=11, color="C00000", bold=True)
HELP_FONT = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
DEFAULT_FONT = Font(name="Calibri", size=11, italic=True, color="595959")
INTRO_FONT = Font(name="Calibri", size=11)

THIN = Side(border_style="thin", color="BFBFBF")
BORDER_INPUT = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

UNLOCKED = Protection(locked=False)
LOCKED = Protection(locked=True)

# --- Sheet structure ---
SHEETS = [
    "Start",
    "1. Bedrijf & Contact",
    "2. Payroll basis",
    "3. Loonheffing",
    "4. Reserveringen & Pensioen",
    "5. Verlof & Grootboek",
    "6. Akkoord",
]
SHEET_TITLES = {
    "Start": "Payroll Discovery Document - MKB",
    "1. Bedrijf & Contact": "1. Bedrijf & Contact",
    "2. Payroll basis": "2. Payroll basis",
    "3. Loonheffing": "3. Loonheffing",
    "4. Reserveringen & Pensioen": "4. Reserveringen & Pensioen",
    "5. Verlof & Grootboek": "5. Verlof & Grootboek",
    "6. Akkoord": "6. Akkoord",
}

VALIDATIONS = {
    "val_JaNee": ["Ja", "Nee"],
    "val_Salarisperiode": ["Week", "4 weken", "Maand"],
    "val_Taal": ["Nederlands", "Engels"],
    "val_Indeling": ["Klein", "Middel", "Groot"],
    "val_Oproep": [
        "Nee", "Ja, met verplichting te komen",
        "Ja, zonder verplichting te komen", "Ja, beide",
    ],
    "val_Maand": [
        "Januari", "Februari", "Maart", "April", "Mei", "Juni",
        "Juli", "Augustus", "September", "Oktober", "November", "December",
        "Periode 1", "Periode 2", "Periode 3", "Periode 4", "Periode 5",
        "Periode 6", "Periode 7", "Periode 8", "Periode 9", "Periode 10",
        "Periode 11", "Periode 12", "Periode 13",
    ],
    "val_Zichtbaarheid": ["Op de betaaldatum", "# dagen na de betaaldatum"],
    "val_Formaat": ["XML", "Excel", "CSV", "TXT", "Anders"],
    "val_Basisperiode": [
        "In de periode", "In de maand", "Per 1 januari", "Per 31 december",
        "Geboren voor", "Geboren op of na", "AOW Gerechtigd",
    ],
    "val_FinSysteem": [
        "Exact Online", "AFAS", "Twinfield", "Visma.net", "SnelStart", "Unit4", "Anders",
    ],
    "val_30pct": ["Nee", "Ja, bruto-aftopping", "Ja, netto-vergoeding"],
    "val_Toggle": ["Basis", "Uitgebreid"],
    "val_Checkbox": ["☐", "☑"],
    "val_Sector": [
        "1 - Agrarisch bedrijf", "2 - Tabakverwerkende industrie",
        "3 - Bouwbedrijf", "4 - Baggerbedrijf",
        "5 - Hout en emballage-industrie", "6 - Timmerindustrie",
        "7 - Meubel- en orgelbouw industrie",
        "8 - Groothandel hout, zagerijen, schaverijen en houtbereidingsindustrie",
        "9 - Grafische industrie", "10 - Metaal industrie",
        "11 - Elektronische industrie", "12 - Metaal- en technische bedrijfstakken",
        "13 - Bakkerijen", "14 - Suikerverwerkende industrie",
        "15 - Slagersbedrijven", "16 - Slagers overig",
        "17 - Detailhandel en ambachten", "18 - Reiniging",
        "19 - Grootwinkelbedrijf", "20 - Havenbedrijven",
        "21 - Havenclassificeerders", "22 - Binnenscheepvaart",
        "23 - Visserij", "24 - Koopvaardij",
        "25 - Vervoer KLM", "26 - Vervoer NS",
        "27 - Vervoer posterijen", "28 - Taxivervoer",
        "29 - Openbaar vervoer", "30 - Besloten busvervoer",
        "31 - Overig personenvervoer te land en in de lucht",
        "32 - Overig goederenvervoer te land en in de lucht",
        "33 - Horeca algemeen", "34 - Horeca catering",
        "35 - Gezondheid, geestelijke en maatschappelijke belangen",
        "38 - Banken", "39 - Verzekeringswezen en ziekenfondsen",
        "40 - Uitgeverij", "41 - Groothandel I", "42 - Groothandel II",
        "43 - Zakelijke dienstverlening I", "44 - Zakelijke dienstverlening II",
        "45 - Zakelijke dienstverlening III", "46 - Zuivelindustrie",
        "47 - Textielindustrie",
        "48 - Steen-, cement-, glas- en keramische industrie",
        "49 - Chemische industrie", "50 - Voedingsindustrie",
        "51 - Algemene industrie", "52 - Uitzendbedrijven",
        "53 - Bewakingsondernemingen", "54 - Culturele instellingen",
        "55 - Overige takken van bedrijf en beroep",
        "56 - Schildersbedrijf", "57 - Stukadoorsbedrijf",
        "58 - Dakdekkersbedrijf", "59 - Mortelbedrijf",
        "60 - Steenhouwersbedrijf",
        "61 - Overheid, onderwijs en wetenschappen",
        "62 - Overheid, rijk, politie en rechterlijke macht",
        "63 - Overheid, defensie",
        "64 - Overheid, provincies, gemeenten en waterschappen",
        "65 - Overheid, openbare nutsbedrijven",
        "66 - Overheid, overige instellingen",
        "67 - Werk en (re)integratie",
        "68 - Railbouw", "69 - Telecommunicatie",
    ],
}

# --- Global registries ---
REQUIRED_REFS = []  # list of (sheet_name, coord)
DV_CACHE = {}       # (sheet_name, list_name) -> DataValidation


# --- Helpers ---
def setup_sheet(ws, title):
    ws.column_dimensions["A"].width = 42
    for col in ("B", "C", "D", "E", "F", "G", "H"):
        ws.column_dimensions[col].width = 28
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = title
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.protection.sheet = True


def subheader(ws, row, text, span=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = SUBHEADER_FILL
    c.font = SUBHEADER_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def label(ws, row, text, required=False, col=1):
    display = (text + " *") if required else text
    c = ws.cell(row=row, column=col, value=display)
    c.font = REQ_FONT if required else LABEL_FONT
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)


def plain(ws, coord, text, font=None, alignment=None):
    c = ws[coord]
    c.value = text
    c.font = font or BASE_FONT
    if alignment:
        c.alignment = alignment


def input_cell(ws, coord, required=False, comment=None, default=None, is_default=False):
    c = ws[coord]
    if default is not None:
        c.value = default
    c.fill = DEFAULT_FILL if is_default else INPUT_FILL
    c.font = DEFAULT_FONT if is_default else BASE_FONT
    c.border = BORDER_INPUT
    c.protection = UNLOCKED
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    if comment:
        c.comment = Comment(comment, "SD Worx")
    if required:
        REQUIRED_REFS.append((ws.title, coord))


def dropdown(ws, coord, listname, required=False, comment=None, default=None, is_default=False):
    input_cell(ws, coord, required=required, comment=comment, default=default, is_default=is_default)
    key = (ws.title, listname)
    dv = DV_CACHE.get(key)
    if dv is None:
        dv = DataValidation(
            type="list", formula1=f"={listname}", allow_blank=True, showDropDown=False
        )
        dv.error = "Ongeldige keuze. Kies uit de lijst."
        dv.errorTitle = "Ongeldige invoer"
        ws.add_data_validation(dv)
        DV_CACHE[key] = dv
    dv.add(coord)


def help_line(ws, coord, text):
    c = ws[coord]
    c.value = text
    c.font = HELP_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def opmerkingen_block(ws, start_row, rows=4, span=8):
    subheader(ws, start_row, "Opmerkingen", span=span)
    ws.merge_cells(
        start_row=start_row + 1, start_column=1,
        end_row=start_row + rows, end_column=span,
    )
    c = ws.cell(row=start_row + 1, column=1)
    c.fill = INPUT_FILL
    c.border = BORDER_INPUT
    c.protection = UNLOCKED
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)


def footnote(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = HELP_FONT


# --- Validations sheet ---
def build_validations(wb):
    vs = wb.create_sheet("_Validations")
    for col_idx, (name, values) in enumerate(VALIDATIONS.items(), start=1):
        letter = get_column_letter(col_idx)
        vs.column_dimensions[letter].width = 40
        vs.cell(row=1, column=col_idx, value=name).font = Font(bold=True)
        for i, v in enumerate(values, start=2):
            vs.cell(row=i, column=col_idx, value=v)
        ref = f"'_Validations'!${letter}$2:${letter}${len(values) + 1}"
        wb.defined_names[name] = DefinedName(name=name, attr_text=ref)
    vs.sheet_state = "hidden"


# --- Start sheet ---
def build_start(wb):
    ws = wb["Start"]
    # Intro
    ws.merge_cells("A3:H3")
    plain(ws, "A3", "Welkom! Dit document is de basis voor de inrichting van uw loonadministratie in Cobra (SD Worx).",
          font=Font(name="Calibri", size=12, bold=True))
    ws.merge_cells("A4:H4")
    plain(ws, "A4",
          "Vul per tabblad de gegevens in. Geschatte invultijd: 30 - 45 minuten. Velden met * zijn verplicht.",
          font=INTRO_FONT)
    ws.merge_cells("A5:H5")
    plain(ws, "A5",
          "Geel = door u in te vullen.   Lichtgrijs = voorstel (overschrijfbaar).   Bij twijfel: leeglaten en contact opnemen met uw consultant.",
          font=HELP_FONT)

    # Consultant block
    subheader(ws, 7, "Uw contactpersoon bij SD Worx")
    label(ws, 8, "Naam consultant")
    input_cell(ws, "B8", comment="In te vullen door SD Worx-consultant.")
    label(ws, 9, "E-mail")
    input_cell(ws, "B9", comment="In te vullen door SD Worx-consultant.")
    label(ws, 10, "Telefoon")
    input_cell(ws, "B10", comment="In te vullen door SD Worx-consultant.")

    # Bijlagen checklist
    subheader(ws, 12, "Benodigde bijlagen  (aanvinken wat is meegestuurd)")
    attachments = [
        "Kopie laatste loonstrook van ca. 5 medewerkers (incl. 1 met variabele beloning, 1 met pensioen)",
        "Kopie WHK-beschikking lopend jaar",
        "Kopie grootboekschema + laatste journaalpost (alleen als u GL-export wilt)",
        "Kopie pensioenregeling (alleen als u NIET bij een bedrijfstakpensioenfonds zit)",
        "CAO (alleen als van toepassing)",
        "Functielijst (functiecode + omschrijving)",
    ]
    for i, text in enumerate(attachments):
        r = 13 + i
        dropdown(ws, f"A{r}", "val_Checkbox", default="☐", is_default=True)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        tc = ws.cell(row=r, column=2, value=text)
        tc.font = BASE_FONT
        tc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

    # Progress indicator
    subheader(ws, 20, "Voortgang verplichte velden")
    label(ws, 21, "Aantal ingevuld / totaal")
    # Placeholder; filled at end
    c = ws["B21"]
    c.fill = DEFAULT_FILL
    c.font = DEFAULT_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws["B21"].value = "(wordt berekend)"
    label(ws, 22, "Percentage ingevuld")
    c = ws["B22"]
    c.fill = DEFAULT_FILL
    c.font = DEFAULT_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws["B22"].value = "(wordt berekend)"
    help_line(ws, "C21",
              "Tip: sla het bestand op of druk op F9 om de teller te verversen.")

    opmerkingen_block(ws, 24)


# --- 1. Bedrijf & Contact ---
def build_bedrijf(wb):
    ws = wb["1. Bedrijf & Contact"]

    subheader(ws, 3, "Bedrijfsgegevens")
    rows_company = [
        ("Bedrijfsnaam", "B4", True, None),
        ("Straat + huisnummer", "B5", True, None),
        ("Postcode", "B6", True, "Formaat: 1234 AB"),
        ("Plaats", "B7", True, None),
        ("KvK-nummer", "B8", False, "8 cijfers, zonder vestigingsnummer."),
        ("Loonheffingsnummer", "B9", False,
         "Formaat: 123456789L01. Wordt ook gevraagd op tab 3. Loonheffing; hier mag u overslaan."),
    ]
    for i, (lab, cell, req, cmt) in enumerate(rows_company):
        label(ws, 4 + i, lab, required=req)
        input_cell(ws, cell, required=req, comment=cmt)

    subheader(ws, 11, "Contactpersoon HR / salarisadministratie")
    rows_hr = [
        ("Naam", "B12", True),
        ("Functie", "B13", False),
        ("Telefoon", "B14", True),
        ("E-mail", "B15", True),
    ]
    for i, (lab, cell, req) in enumerate(rows_hr):
        label(ws, 12 + i, lab, required=req)
        input_cell(ws, cell, required=req)

    subheader(ws, 17, "Contactpersoon Finance / boekhouding")
    rows_fin = [
        ("Naam", "B18", True),
        ("Functie", "B19", False),
        ("Telefoon", "B20", True),
        ("E-mail", "B21", True),
    ]
    for i, (lab, cell, req) in enumerate(rows_fin):
        label(ws, 18 + i, lab, required=req)
        input_cell(ws, cell, required=req)

    subheader(ws, 23, "Bankgegevens (voor SEPA-betaalbestand nettolonen)")
    label(ws, 24, "IBAN-rekeningnummer", required=True)
    input_cell(ws, "B24", required=True,
               comment="Nederlands IBAN, bijv. NL91 ABNA 0417 1643 00. Spaties mogen, worden genegeerd.")
    label(ws, 25, "Naam bank")
    input_cell(ws, "B25")
    label(ws, 26, "BIC / SWIFT-code")
    input_cell(ws, "B26",
               comment="Alleen invullen als de bank NIET in de IBAN-zone (Europa) zit.")

    footnote(ws, 28, "* = verplicht veld. Niet-verplichte velden mogen leeg blijven.")
    opmerkingen_block(ws, 30)


# --- Progress formula wiring ---
def wire_progress(wb):
    ws = wb["Start"]
    if not REQUIRED_REFS:
        return
    refs = ",".join(f"'{s}'!{coord_abs(c)}" for s, c in REQUIRED_REFS)
    total = len(REQUIRED_REFS)
    # Filled count / total as text
    ws["B21"].value = f'=COUNTA({refs})&" / {total}"'
    ws["B22"].value = f'=IFERROR(TEXT(COUNTA({refs})/{total},"0%"),"0%")'


def coord_abs(coord):
    """Convert 'B4' -> '$B$4'."""
    import re
    m = re.match(r"([A-Z]+)(\d+)", coord)
    return f"${m.group(1)}${m.group(2)}"


# --- Main ---
def build():
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEETS:
        ws = wb.create_sheet(name)
        setup_sheet(ws, SHEET_TITLES[name])
    build_validations(wb)
    build_start(wb)
    build_bedrijf(wb)
    wire_progress(wb)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    path = build()
    size = os.path.getsize(path)
    print(f"SAVED: {path} ({size} bytes, {size/1024:.1f} KB)")
    print(f"Required fields registered: {len(REQUIRED_REFS)}")
