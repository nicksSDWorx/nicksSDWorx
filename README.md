# SME Omzetmonitor

Eén zelfstandige, lokaal werkende webtool waarmee SME structureel kan monitoren hoe klanten groeien of
krimpen, of geleverde diensten volledig en juist zijn gefactureerd, en waar kortingen, credits,
prijswijzigingen, dubbele facturatie, gemiste facturatie, groeikansen en retentierisico's aandacht vragen.

## Gebruik

1. Download **`SME_Omzetmonitor.html`** en open het bestand in **Chrome of Edge** (dubbelklikken volstaat —
   geen installatie, geen internet nodig).
2. Sleep het omzetbestand (`.xlsx`, `.xls` of `.csv`) in het uploadvak.
3. Controleer de kolomherkenning en koppel zo nodig handmatig een kolom.
4. Klik op **Analyse uitvoeren** en gebruik de tabs Overzicht, Klanten, Signalen, Vergelijkingen en
   Instellingen.

### Verwachte kolommen

| Kolom | Betekenis |
|---|---|
| Debiteur nr | Unieke klantidentifier |
| Debiteur | Klantnaam |
| Periode jaar / Periode maand | Facturatieperiode |
| Product | Gefactureerd product of dienst |
| Comm omschrijving | Commerciële omschrijving (korting, credit, uitvoeringsperiode, …) |
| Aantal | Gefactureerd volume (betekenis productafhankelijk) |
| Verk prijs ex. BTW | Verkoopprijs per eenheid |
| Totaal verkoop company | Gerealiseerde omzet ex. btw |

De herkenning is hoofdletter- en spatie-ongevoelig en accepteert kleine schrijfvariaties; een optionele
kortingskolom wordt automatisch meegenomen.

## Privacy

Alle verwerking gebeurt **100% lokaal in de browser**. De tool maakt geen enkele externe verbinding
(SheetJS en Chart.js zijn in het bestand ingebed), slaat niets op (ook niet in localStorage) en verstuurt
niets. Sluiten van het tabblad wist alle gegevens. Daarmee is de tool geschikt voor gevoelige klant-,
omzet- en facturatiegegevens.

## Wat signaleert de tool?

Elf analyseregels (fase 1), met vaste prioriteitsvolgorde en ontdubbeling:

1. Korting controleren
2. Creditnota of correctie onderzoeken
3. Mogelijk dubbel gefactureerd
4. Mogelijk niet gefactureerd (volume zonder omzet)
5. Factuurwaarde controleren (aantal × prijs ≠ omzet)
6. Mogelijk gestopt of ontbrekend product
7. Prijswijziging controleren (binnen debiteur + product)
8. Mogelijk omzet- of retentierisico
9. Mogelijke commerciële groeikans
10. Periodeverdeling controleren
11. Mogelijk nieuw product of nieuwe dienstverlening

Iedere combinatie van debiteur, product en verkoopprijs wordt vergeleken met de **vorige beschikbare
maand**, **dezelfde maand vorig jaar** en het **gemiddelde van de drie voorgaande beschikbare maanden**.
Alle drempels (omzet- 20%, volume- 20% en prijsafwijking 10%, minimale omzet €250, tolerantie
rekenverschil €1, aantal maanden voor "gestopt product") zijn instelbaar via het scherm Instellingen.

## Techniek

- Eén HTML-bestand: HTML + moderne CSS + vanilla JavaScript.
- [SheetJS](https://sheetjs.com/) 0.20.3 (Apache-2.0) voor Excel/CSV, ingebed.
- [Chart.js](https://www.chartjs.org/) 4.4.3 (MIT) voor grafieken, ingebed.
- Geen build-omgeving, backend of database.
- Vormgeving volgens de SD Worx-huisstijl.

Een uitgebreide README staat als commentaar boven in de broncode van `SME_Omzetmonitor.html`.
