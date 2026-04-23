# Payroll Discovery Document – MKB-versie (v1)

Vereenvoudigde inventarisatie voor **SD Worx Cobra** voor kleinere Nederlandse werkgevers (≤ ~50 medewerkers, 1 BV/betaalgroep, geen split payroll, geen expats, geen UPA-complexiteit).

- **Bestand:** `outputs/Payroll_Discovery_Document_MKB.xlsx`
- **Generator:** `scripts/build_mkb.py` (Python + openpyxl)
- **QA:** `scripts/qa.py`
- **Geschatte invultijd voor klant:** 30 – 45 minuten
- **Aantal verplichte velden:** 36 (tegen enkele honderden in het origineel)

## Wijzigingen t.o.v. het origineel

| Sheet origineel | Behandeling in MKB-versie |
|---|---|
| `Explanation` | samengevoegd in **Start** (welkomsttekst + bijlagen-checklist) |
| `General Company Info` | vereenvoudigd in **1. Bedrijf & Contact** (HR- én Finance-contact op één tab; KvK en loonheffingsnr als optionele velden) |
| `Payroll settings` | vereenvoudigd in **2. Payroll basis** (uurloon-berekeningsconfig weg; split/expats/uitkering/gepensioneerd weg) |
| `Payroll tax data` | vereenvoudigd in **3. Loonheffing** – WHK-premie helemaal geen invulveld meer; alleen bijlage-verzoek |
| `Accruals` (Vakantiegeld, 13e maand) | vereenvoudigd in **4. Reserveringen & Pensioen** (referentiejaar en % voorgevuld) |
| `Accruals (2)` – reservering 3 & 4 | **verwijderd** |
| `Accruals (IKB)` | verplaatst naar **Geavanceerd-toggle** onderaan tab 4 |
| `(Pension)Scheme` | vereenvoudigd in **4. Reserveringen & Pensioen** met duidelijke twee-takken (BPF / eigen regeling) |
| `(Pension)Scheme (2)` – excedent | **verwijderd** → vraag "Heeft u aanvullende regelingen?" |
| `(WGA_WIA)Scheme (3)` | **verwijderd** → zelfde vraag |
| `(Other)Scheme (4)` | **verwijderd** → zelfde vraag |
| `GL Ledgers & Cost centers` (112 rijen) | **verwijderd** → "stuur grootboekschema + journaalpost als bijlage" |
| `GL file export` | sterk vereenvoudigd in **5. Verlof & Grootboek** (systeem, formaat, per kostenplaats) |
| `WAZO` | behouden in **5. Verlof & Grootboek** als compacte tabel (7 rijen × 3 kolommen, defaults = wettelijk minimum) |
| `Wage codes` (286 rijen) | **verwijderd** → "stuur loonstroken + looncomponenten als bijlage" |
| `Approval` | behouden in **6. Akkoord** |
| `Languages` | **niet gebruikt in v1** (Nederlands-only) |
| `Validation values` | gekopieerd naar verborgen `_Validations` + 14 named ranges |

## Structuur

Zeven zichtbare tabbladen + één verborgen:
1. **Start** – welkom, consultant-contact, bijlagen-checklist, voortgangsteller
2. **1. Bedrijf & Contact** – NAW, HR-contact, Finance-contact, bank
3. **2. Payroll basis** – CAO, werktijden, salarisperiode, salarisstrook, 30%-regeling, werknemerstypen
4. **3. Loonheffing** – loonheffingsnr, sector, indeling, CBS-code, WBSO; WHK als bijlage
5. **4. Reserveringen & Pensioen** – vakantiegeld, 13e maand, pensioen (BPF / eigen), aanvullende regelingen, IKB onder Geavanceerd-toggle
6. **5. Verlof & Grootboek** – WAZO-tabel + GL-export-basis
7. **6. Akkoord** – klant, ondertekening, akkoordvinkjes
- `_Validations` – verborgen tab met 14 dropdownlijsten (`val_JaNee`, `val_Sector`, `val_Maand`, `val_FinSysteem`, etc.)

## Velden onder "Geavanceerd"

Alleen op **4. Reserveringen & Pensioen** (rijen 31 – 36):
- IKB-toggle (Basis / Uitgebreid)
- IKB-percentage
- IKB-referentieperiode van / tot
- IKB-uitbetaalmaand

Deze cellen hebben lichtgrijze achtergrond + celcommentaar "Alleen invullen bij toggle = 'Uitgebreid'". Klanten zonder IKB laten het blok leeg.

## Gebruiksvriendelijkheid

- **Gele cellen (`#FFF2CC`)** = door klant in te vullen.
- **Lichtgrijze cellen (`#F2F2F2`)** = voorstel/default (overschrijfbaar).
- **Donkerblauwe balk (`#1F3864`)** = sheet-header; **lichtblauw (`#D9E1F2`)** = sub-sectie.
- **Rood sterretje (*)** = verplicht veld.
- **Celcommentaren** bij complexe velden (hover in Excel).
- **Dropdowns** voor alle keuze-velden (17 in totaal, 43 cellen gedekt).
- **Sheet-protection** aan zonder wachtwoord: labels gelockt, gele cellen bewerkbaar.
- **Print**: landscape A4, fit-to-width.
- **Voortgangsteller** op Start (COUNTA van 36 verplichte referenties / 36).

## Defaults (voorgevuld, overschrijfbaar)

| Veld | Default |
|---|---|
| Uren per week | 40 |
| Uren per dag (ma–vr) | 8 elk |
| Salarisperiode | Maand |
| Salarisstrook-taal | Nederlands |
| Strook zichtbaar | Op de betaaldatum |
| Vakantiegeld % | 8,33% |
| Vakantiegeld ref-periode | Juni t/m mei |
| Uitbetaalmaand vakantiegeld | Mei |
| Minimum leeftijd pensioen | 21 |
| Maximum leeftijd pensioen | AOW-leeftijd |
| WAZO defaults | Wettelijk minimum per verlofsoort |

## Aanbevolen vervolgstappen voor de implementatieconsultant

1. **Voorafgaand aan uitlevering**: vul de 3 velden op **Start** in (consultant-naam/e-mail/telefoon).
2. **Bij uitlevering**: mail het bestand naar HR-contact met korte instructie + link naar bijlagen-SFTP.
3. **Bij retour**: draai `python scripts/qa.py` (of: laat openen in Excel) en controleer of voortgangsteller op 100% staat.
4. **Bij opengelaten velden** (met name pensioen-details, CBS-code, sector): neem mondeling door.
5. **Aanvullende regelingen op Ja**: plan 15 min call over WGA-hiaat / excedent / netto pensioen.
6. **IKB-toggle op Uitgebreid**: controleer dat velden B33–B36 zijn ingevuld.
7. **Map bijlagen**: loonstroken → Cobra-wage-codes; grootboekschema → GL-mapping; WHK-beschikking → premie-instellingen.

## Bekende beperkingen & aannames

- **Nederlands-only in v1.** Engelse versie komt in een volgende iteratie (placeholder: `Languages`-tab uit origineel is niet gebruikt).
- **Geen VBA** (blijft pure `.xlsx`). Conditional hiding van Geavanceerd-velden is visueel (grijs + comment), niet functioneel afgedwongen.
- **Checkbox-kolom op Start** werkt via `☐`/`☑`-dropdown — geen "echte" form-control checkbox (vereist VBA of ActiveX).
- **WAZO-defaults** zijn wettelijk minimum. CAO-gunstigere regelingen moeten handmatig overschreven worden.
- **Indeling werkgever** is voorgesteld als klein < 25 / middel 25–100 / groot > 100 mdw. De Belastingdienst hanteert formeel een premieloon-ondergrens — bij twijfel laat de klant het veld leeg en bepaalt SD Worx.
- **Sector-lijst** bevat de 67 officiële Belastingdienst-sectoren (exclusief 36 en 37, die bestaan niet).
- **Voortgangsformule** (COUNTA op 36 refs) update bij elke bewerking in Excel; in LibreOffice soms pas na `Ctrl+Shift+F9`.
- **LibreOffice headless in CI-sandbox**: werkt niet in onze build-omgeving (Java-init faalt), maar `file`-detectie en openpyxl-round-trip bevestigen dat het bestand valide Excel 2007+ is. Klanten openen het in gewone Excel of LibreOffice zonder issues.

## Hergenereren

```bash
python3 scripts/build_mkb.py   # bouwt outputs/Payroll_Discovery_Document_MKB.xlsx
python3 scripts/qa.py          # 8 QA-checks
```

Python ≥ 3.9, `pip install openpyxl>=3.1`.
