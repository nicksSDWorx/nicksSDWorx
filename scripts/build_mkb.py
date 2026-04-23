"""
Builder for Payroll Discovery Document MKB (v2, Dutch-only).
Phased build: each phase extends this script.
Phase 2 status: skeleton only — 7 visible sheets + _Validations (hidden) with named ranges.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties

OUTPUT = "/home/user/nicksSDWorx/outputs/Payroll_Discovery_Document_MKB.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
BASE_FONT = Font(name="Calibri", size=11)

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
        "Nee",
        "Ja, met verplichting te komen",
        "Ja, zonder verplichting te komen",
        "Ja, beide",
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


def build_validations(wb):
    vs = wb.create_sheet("_Validations")
    vs.column_dimensions["A"].width = 25
    for col_idx, (name, values) in enumerate(VALIDATIONS.items(), start=1):
        letter = get_column_letter(col_idx)
        vs.column_dimensions[letter].width = 40
        vs.cell(row=1, column=col_idx, value=name).font = Font(bold=True)
        for i, v in enumerate(values, start=2):
            vs.cell(row=i, column=col_idx, value=v)
        ref = f"'_Validations'!${letter}$2:${letter}${len(values) + 1}"
        wb.defined_names[name] = DefinedName(name=name, attr_text=ref)
    vs.sheet_state = "hidden"


def build():
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEETS:
        ws = wb.create_sheet(name)
        setup_sheet(ws, SHEET_TITLES[name])
    build_validations(wb)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    path = build()
    size = os.path.getsize(path)
    print(f"SAVED: {path} ({size} bytes, {size/1024:.1f} KB)")
