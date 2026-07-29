# 4-ogen loonreview / 4-eyes payroll review

**Bestand:** [`4-eyes-review.html`](4-eyes-review.html) — download het bestand en open het in een browser (Edge/Chrome). Geen installatie nodig; alle verwerking gebeurt lokaal in de browser, er worden geen gegevens verstuurd.

## Wat doet de tool?

De tool vergelijkt het overzicht uit het payrollsysteem met de input van de klant, per werknemer en per looncode, zodat de 4-ogen reviewer alleen nog naar de **afwijkingen** hoeft te kijken in plaats van alles regel voor regel te controleren.

## De drie bestanden

| # | Bestand | Inhoud |
|---|---|---|
| 1 | **Systeem — aantallen / uren / dagen** | Export uit het payrollsysteem met de aantallen, uren en dagen per looncode |
| 2 | **Systeem — bedragen (€)** | Export uit het payrollsysteem met de bedragen per looncode |
| 3 | **Klant — input** | Het aangeleverde bestand van de klant (aantallen en/of bedragen) |

De uploadvakken in stap 1 zijn duidelijk gelabeld (Bestand 1 · Systeem, Bestand 2 · Systeem, Bestand 3 · Klant). Alle formaten: `.xlsx`, `.xls`, `.csv`, `.txt`. Je kunt ook met slechts één van de twee systeembestanden vergelijken (dan wordt alleen die dimensie vergeleken).

**Bestand 1 gaat voor.** Per looncode geldt één controle: staat de looncode in bestand 1, dan wordt alleen het aantal gecontroleerd en telt bestand 2 voor die code niet mee. Staat de looncode niet in bestand 1, dan wordt het bedrag uit bestand 2 gecontroleerd. In het resultaat zie je per looncode dus óf de aantal-kolommen óf de bedrag-kolommen gevuld.

## Werkwijze

1. **Dossier & bestanden** — vul klant en periode in, laad de bestanden. Toleranties zijn instelbaar (standaard: bedrag ± € 0,05, aantal ± 0,01) zodat afrondingsverschillen niet als fout worden gemeld.
2. **Kolomtoewijzing** — per bestand (onder elkaar) toont de tool een voorbeeld en een voorstel voor de kolomtoewijzing op basis van de kolomkoppen (NL en EN). Looncodes met tekst erachter (bijv. "1000 Salaris") worden op het nummer gematcht; de omschrijving wordt in het resultaat en het rapport getoond. Voor het aantallenbestand kies je **drie waardekolommen** (aantal, uren, dagen) — per regel worden die opgeteld tot één aantal. Meerdere regels per werknemer+looncode worden gesommeerd; bestanden zonder werknemerkolom (totalen per looncode) werken ook.
3. **Resultaat** — samenvatting + twee weergaven:
   - **Per looncode**: totalen per code, klik op een regel voor detail per werknemer;
   - **Per werknemer**: alle afwijkende regels per werknemer.

## Wat wordt gesignaleerd?

| Status | Betekenis |
|---|---|
| **OK** | Waarden gelijk binnen de tolerantie |
| **Verschil** | Aantal en/of bedrag wijkt af boven de tolerantie |
| **Alleen systeem** | Regel/looncode staat wel in het systeem maar niet in de klantinput |
| **Alleen klant** | Regel/looncode staat wel in de klantinput maar niet in het systeem |
| **Onbekende code** | Looncode komt niet voor in het Master loonmodel 2026 (2.064 codes, ingebouwd) |

**Alleen looncodes uit klantbestand** — met dit vinkje (boven de resultaattabel) worden alleen looncodes gecontroleerd die in het klantbestand voorkomen; systeemcodes die de klant niet aanlevert (bijv. door het systeem berekende codes) blijven volledig buiten de vergelijking. Looncodes die de klant wél aanlevert maar in het systeem ontbreken, blijven zichtbaar. De keuze wordt in het rapport vastgelegd.

**Negeer ontbrekende looncodes** — met dit vinkje (boven de resultaattabel) tellen regels die maar aan één kant voorkomen niet meer mee als afwijking. Ze blijven zichtbaar met een grijze status "(genegeerd)", en echte waardeverschillen en onbekende codes blijven gewoon gemarkeerd. De keuze wordt ook in het rapport vastgelegd.

## Rapport & audit trail

- **Download reviewrapport (CSV)** — volledig rapport (klant, periode, datum/tijd, de drie bestandsnamen, toleranties, negeer-instelling, samenvatting en alle regels) als bewijs voor het dossier; opent direct correct in Excel (NL-notatie).
- **Print / PDF** — printvriendelijke weergave van het resultaat.
- **Nieuwe review** — wist alles en start direct met de volgende klant (toleranties blijven staan).

De taal van de interface is schakelbaar (NL/EN, rechtsboven).

## Instellingen — looncodes uit/aan

Via de knop **⚙ Instellingen** (rechtsboven) kun je individuele looncodes uit- en weer aanzetten. Uitgeschakelde codes blijven volledig buiten de vergelijking, in alle drie de bestanden. De lijst toont alle looncodes uit de geladen bestanden (met omschrijving), is doorzoekbaar en heeft Alles aan/Alles uit; een code die (nog) niet in de bestanden zit, kun je handmatig uitschakelen. De keuze wordt in de browser onthouden voor volgende reviews (de teller op de knop laat zien hoeveel codes uit staan) en de uitgeschakelde codes worden in het CSV-rapport vermeld.

## Master loonmodel bijwerken

De lijst met geldige looncodes (bron: [`Master_loonmodel_2026.csv`](Master_loonmodel_2026.csv)) is in het HTML-bestand ingebouwd in de regel die begint met `var MASTER_CODES_RAW = "1000 1001 …"`. Voor een nieuw loonmodeljaar: vervang die reeks door de nieuwe codes (spatiegescheiden).
