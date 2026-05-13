# Voortgangsrapport-tool

Genereert wekelijks een Markdown-voortgangsrapport uit een implementatie-tracking
Excel. Werkt lokaal op Windows, draait offline, en wijzigt het bronbestand nooit.

## Wat doet de tool?

1. Leest het Excel-bestand **read-only** in (collega's kunnen blijven werken).
2. Bouwt een snapshot van alle implementaties en bewaart die als JSON.
3. Vergelijkt met de snapshot van vorige week.
4. Genereert een rapport met:
   - Samenvatting (totalen, veranderingen sinds vorige week)
   - Tabel met status per implementatie
   - Wijzigingen (nieuw, status-veranderingen, afgerond, mogelijk stagnerend)
   - Knelpunten (deadlines, ontbrekende eigenaars, kritieke datums op "Ntb")
   - Per-implementatie detailblokken voor relevante gevallen
5. Opent het rapport automatisch.

## Eenmalige installatie

1. Installeer Python 3.10 of hoger (https://www.python.org). Bij installatie:
   vink "Add Python to PATH" aan.
2. Open een Command Prompt in deze projectmap.
3. Maak een virtuele omgeving en installeer dependencies:

   ```bat
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

4. Sluit het venster.

## Draaien

**Dubbelklik `run_report.bat`.**

Het script activeert de virtuele omgeving (indien aanwezig), draait `monitor.py`
en opent het rapport. Bij een fout blijft het venster open zodat je de melding kunt lezen.

## Belangrijke locatie

Plaats deze projectmap **niet** in een OneDrive- of SharePoint-gesynchroniseerde
map. Aanbevolen locatie:

```
C:\Tools\voortgangsrapport\
```

In een gesynchroniseerde map kunnen snapshots ongewenst uploaden naar de cloud.

## Configuratie (`config.json`)

| Veld | Betekenis |
|---|---|
| `excel_path` | Pad naar het Excel-bestand. Mag absoluut zijn (`C:\\...\\tracker.xlsx`) of relatief aan deze projectmap. |
| `general_sheet` | Naam van het overzichtstabblad (standaard `GENERAL`). |
| `columns.*` | Kolomnamen op het GENERAL-tabblad. Pas aan als kolommen worden hernoemd. |
| `customer_sheet.metadata_cells` | Celposities (bv. `B4`) voor metadata op klant-tabbladen. |
| `customer_sheet.phase_scorecard` | Celposities van de scorecard-cellen per fase. |
| `thresholds.deadline_warning_days` | Aantal dagen voor Go-live waarop deadline-waarschuwing triggert. |
| `thresholds.stagnation_weeks` | Aantal weken zelfde status voor een stagnatie-signaal. |
| `thresholds.ntb_values` | Tekstwaardes die als "nog te bepalen" tellen voor kritieke datums. |
| `report.open_after_generate` | `true` opent het rapport automatisch na generatie. |
| `report.per_customer_only_with_changes_or_bottlenecks` | `true` toont per-klant blokken alleen voor relevante gevallen. |

## Overschakelen van dummy- naar productie-bestand

1. Open `config.json` in Kladblok.
2. Wijzig `"excel_path"` naar het volledige pad van het echte bestand, bv:

   ```json
   "excel_path": "C:\\Users\\<jouwgebruiker>\\Teams\\Implementaties\\tracker.xlsx"
   ```

   Let op dubbele backslashes (`\\`) in JSON.
3. Sla op en draai `run_report.bat`.

## Wekelijks automatisch draaien (Windows Taakplanner)

1. Druk Windows-toets, typ "Taakplanner", open de app.
2. Klik rechts op **Eenvoudige taak maken**.
3. Naam: "Voortgangsrapport wekelijks". Klik volgende.
4. Trigger: **Wekelijks**. Kies dag en tijd (bv. maandag 08:00). Volgende.
5. Actie: **Een programma starten**. Volgende.
6. Programma/script: bladeren naar `run_report.bat` in deze projectmap.
7. **Beginnen in**: het pad van deze projectmap (zonder bestandsnaam).
8. Voltooien.

Tip: zorg dat je computer aanstaat op het geplande moment, anders draait de
taak bij de volgende keer aanmelden.

## Troubleshooting

**"Excel-bestand niet gevonden"**
Controleer `excel_path` in `config.json`. Voor Teams-bestanden: het lokale pad
staat na het synchroniseren in `C:\Users\<naam>\<Teams-naam>\...`.

**"Kan tabblad 'GENERAL' niet lezen"**
Het tabblad is hernoemd of verwijderd. Pas `general_sheet` aan in `config.json`.

**"config.json bevat een syntaxfout op regel X"**
Open `config.json` in Kladblok, controleer komma's, dubbele aanhalingstekens
en accolades. Een online JSON-validator (offline beschikbaar) kan helpen.

**Het rapport opent niet automatisch**
Zet `report.open_after_generate` op `true`, of open handmatig
`data\reports\YYYY-MM-DD_voortgangsrapport.md`.

**Onverwachte fout met kolomnaam**
Het Excel-bestand heeft een kolomnaam gekregen die niet in `config.json` staat.
Pas de mapping aan in `config.json` onder `columns`.

**Bestand vergrendeld door Teams**
Sluit het Excel-bestand bij collega's, of wacht enkele seconden en draai
opnieuw. De tool opent het bestand alleen-lezen — een leesvergrendeling
hoort niet te ontstaan, maar Teams kan tijdens sync kort blokkeren.

## Bestandsstructuur

```
config.json              configuratie
monitor.py               hoofdscript
tools/simulate_week.py   testhulp (dev-only, mag weg in productie)
data/snapshots/          wekelijkse JSON-snapshots (niet syncen)
data/reports/            gegenereerde Markdown-rapporten
dummy_tracker.xlsx       geanonimiseerd testbestand
run_report.bat           dubbelklik om te draaien
requirements.txt         Python-dependencies
README.md                deze tekst
COMPLIANCE.md            audit-document voor IT/Security
```

## Compliance

Voor IT/Security: zie `COMPLIANCE.md`.
