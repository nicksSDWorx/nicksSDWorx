# Compliance-document — Voortgangsrapport-tool

**Versie**: 1.0
**Bouw-datum**: 2026-05-13

Dit document beschrijft de tool voor IT- en Security-review. Alle hieronder
genoemde eigenschappen zijn verifieerbaar in de broncode in deze projectmap.

## 1. Dependencies

De tool gebruikt uitsluitend `pandas`, `openpyxl` en de Python-standaardlibrary.

| Package | Doel | Netwerk-gebruik |
|---|---|---|
| `pandas` | Inlezen en transformatie van Excel-data | Geen bij normaal gebruik. |
| `openpyxl` | Read-only inlezen van .xlsx | Geen. |
| Stdlib (`json`, `os`, `sys`, `pathlib`, `datetime`, `dataclasses`) | Standaardfuncties | Geen. |

Volledige lijst staat in `requirements.txt`. Er worden geen andere packages
geïnstalleerd of dynamisch geladen.

## 2. Geen netwerk-calls

De code importeert **geen** van de volgende modules: `requests`, `httpx`,
`urllib`, `socket`, `http.client`, `aiohttp`, `ftplib`, `smtplib`. Te
verifiëren met:

```bat
findstr /S /N /R "import requests import httpx import urllib import socket import http.client" *.py
```

De enige externe systeem-aanroep is `os.startfile()` in `monitor.py:open_report()`,
die het lokale rapport opent met de standaard-applicatie van Windows. Dit
veroorzaakt geen netwerk-verkeer.

## 3. Read-only Excel-toegang

Het Excel-bronbestand wordt geopend met:

- `openpyxl.load_workbook(..., read_only=True, data_only=True)` in
  `monitor.py:read_excel()`
- `pandas.read_excel(..., engine="openpyxl")` in `monitor.py:read_general()`

Beide laden het bestand in het geheugen zonder schrijven. De `read_only=True`
flag van openpyxl voorkomt zelfs onbedoelde schrijfacties.

**Verificatie**: vergelijk de modificatiedatum van het Excel-bestand voor en
na een run. Deze moet identiek blijven.

## 4. Bestanden die de tool leest

- `config.json` (configuratie)
- Het in `config.json` geconfigureerde Excel-bestand (read-only)
- Eerdere snapshot-bestanden in `data/snapshots/` (read-only, eigen output)

Geen andere bestanden worden gelezen.

## 5. Bestanden die de tool schrijft

- `data/snapshots/YYYY-MM-DD_snapshot.json` — leesbare JSON-snapshot. Geen
  pickle of andere code-executable serialisatie.
- `data/reports/YYYY-MM-DD_voortgangsrapport.md` — Markdown-rapport.

Beide bestanden bevatten klantdata uit het bronbestand in geaggregeerde vorm
en blijven lokaal. Plaats de projectmap **niet** in een gesynchroniseerde
OneDrive- of SharePoint-locatie.

## 6. Logging

Alle log- en foutberichten in `monitor.py` bevatten uitsluitend technische
informatie: bestandsnamen, regelaantallen, fouttypes (`ValueError`,
`FileNotFoundError`, etc.). Inhoud van Excel-cellen (klantnamen, opmerkingen,
statussen) wordt **niet** gelogd. Helper-functie `log()` in `monitor.py`
verzorgt dit consistent.

## 7. Geen geheimen

De tool gebruikt geen credentials, API-keys, tokens of environment variables.
Er is niets om te roteren of veilig op te slaan.

## 8. Geen externe scripts of binaries

De tool roept geen `subprocess` aan om andere programma's te starten (op
`os.startfile()` voor het openen van het rapport na). Er worden geen scripts
gedownload of dynamisch geëxecuteerd.

## 9. Geen pickle of vergelijkbare risico's

Snapshots gebruiken `json` voor serialisatie. Geen `pickle`, `dill`, `marshal`
of `shelve`. De JSON-bestanden zijn met een teksteditor leesbaar en
controleerbaar.

## 10. Verificatie zonder netwerk

De tool moet identiek functioneren met netwerk uitgeschakeld. Te testen:

1. Schakel netwerkadapter uit (Windows: instellingen → Netwerk → adapter
   uitschakelen, of vliegtuigmodus).
2. Dubbelklik `run_report.bat`.
3. Verwachte uitkomst: rapport gegenereerd en geopend, identiek aan een run
   met netwerk aan.

## 11. Broncode-transparantie

Alle bestanden zijn platte tekst (.py, .json, .md, .bat). Geen `.pyc`-only
distributie, geen obfuscation, geen compiled C-extensions ontwikkeld door
de auteur. De gebruikte packages (`pandas`, `openpyxl`) bevatten compiled
delen — dit zijn standaardpackages uit PyPI.

## 12. Verifieerbare grep-checks

Voor de reviewer, te draaien in de projectmap:

```bat
REM Geen netwerk-imports
findstr /R /C:"^import requests" /C:"^import urllib" /C:"^import http" /C:"^import socket" *.py
REM Geen pickle
findstr /R /C:"^import pickle" /C:"^from pickle" *.py
REM Geen subprocess
findstr /R /C:"^import subprocess" /C:"^from subprocess" *.py
```

Alle bovenstaande commando's horen geen resultaten op te leveren in `monitor.py`
of `tools/simulate_week.py`.

## 13. Klantdata in output

`data/reports/*.md` en `data/snapshots/*.json` bevatten geaggregeerde
implementatie-gegevens uit het bronbestand: klantnummers, klantnamen,
eigenaars, deadlines, fase-status. Dit is de beoogde output van de tool. Vrije
tekstvelden (`Opmerkingen`) op de takenlijst worden **niet** in snapshot of
rapport opgenomen — alleen geaggregeerde tellingen.

## 14. Aanbevelingen voor de gebruiker

- Plaats de projectmap op een niet-gesynchroniseerde locatie (bv. `C:\Tools\`).
- Beperk toegang tot de projectmap tot bevoegde collega's via NTFS-rechten.
- Verwijder `tools/simulate_week.py` en `dummy_tracker.xlsx` voordat de tool
  in productie naast het echte bestand draait — deze zijn alleen voor de
  ontwikkelingsfase.
