# 4-ogen loonreview / 4-eyes payroll review

**Bestand:** [`4-eyes-review.html`](4-eyes-review.html) — download het bestand en open het in een browser (Edge/Chrome). Geen installatie nodig; alle verwerking gebeurt lokaal in de browser, er worden geen gegevens verstuurd.

## Wat doet de tool?

De tool vergelijkt het overzicht uit het payrollsysteem met de input van de klant, per werknemer en per looncode, zodat de 4-ogen reviewer alleen nog naar de **afwijkingen** hoeft te kijken in plaats van alles regel voor regel te controleren.

## Werkwijze

1. **Dossier & bestanden** — vul klant en periode in, laad beide bestanden (`.xlsx`, `.xls`, `.csv`, `.txt`). Toleranties zijn instelbaar (standaard: bedrag ± € 0,05, aantal ± 0,01) zodat afrondingsverschillen niet als fout worden gemeld.
2. **Kolomtoewijzing** — de tool herkent automatisch de kolommen (werknemer, looncode, aantal, bedrag) op basis van de kolomkoppen (NL en EN); controleer/corrigeer waar nodig. Werkt met elk bestandsformaat van de klant. Meerdere regels per werknemer+looncode worden gesommeerd. Bestanden zonder werknemerkolom (totalen per looncode) werken ook.
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

## Rapport & audit trail

- **Download reviewrapport (CSV)** — volledig rapport (klant, periode, datum/tijd, bestandsnamen, toleranties, samenvatting en alle regels) als bewijs voor het dossier; opent direct correct in Excel (NL-notatie).
- **Print / PDF** — printvriendelijke weergave van het resultaat.
- **Nieuwe review** — wist alles en start direct met de volgende klant (toleranties blijven staan).

De taal van de interface is schakelbaar (NL/EN, rechtsboven).

## Master loonmodel bijwerken

De lijst met geldige looncodes (bron: [`Master_loonmodel_2026.csv`](Master_loonmodel_2026.csv)) is in het HTML-bestand ingebouwd in de regel die begint met `var MASTER_CODES_RAW = "1000 1001 …"`. Voor een nieuw loonmodeljaar: vervang die reeks door de nieuwe codes (spatiegescheiden).
