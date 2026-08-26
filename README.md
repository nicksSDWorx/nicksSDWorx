# Expatregeling herrekenen — 30%-regeling jaarcontrole

Tool voor de jaarlijkse controle/herrekening van de 30%-regeling (expatregeling),
gebaseerd op het rekenmodel **“Expatregeling herrekenen 2025”** (tabblad *30% regeling*).
Wat de payco nu handmatig per medewerker in de Excel invult, doet de tool voor alle
medewerkers tegelijk.

## Gebruik

1. Open **`expatregeling-tool.html`** in een browser (dubbelklikken volstaat, geen installatie nodig).
   Alle verwerking gebeurt lokaal in de browser — er worden **geen gegevens verstuurd**.
2. Upload de twee CSV-exports:
   - **Expatregeling overzicht** (bijv. `Expatregeling 2025.csv`)
   - **Historisch overzicht** (bijv. `Historisch Overzicht Export 2025.csv`)
3. Controleer zo nodig de **instellingen** (standaard ingeklapt; klik op *Instellingen* om ze te openen) en klik **Bereken**.
4. Per medewerker verschijnt de uitkomst met status; klik een regel aan voor de volledige
   berekening in het stramien van de Excel, inclusief de actie- en controlesectie.
5. **Exporteer resultaten (CSV)** of **Kopieer naar klembord** (plakt direct in Excel).

## Hoe de velden worden bepaald

Medewerkers worden gematcht op **personeelsnummer** (expat kolom J `Persnr` ↔ historie
kolom E `Persnr`). Per medewerker wordt de expat-regel van de **laatste periode** gebruikt
(hoogste `Periode`, dan hoogste `Runnr`).

| Excel-cel | Bron |
|---|---|
| B5 · Datum start | laatste van: expat kolom U `Regeling Vanaf` en de instelbare standaard startdatum (standaard 1 januari van het berekeningsjaar) |
| B6 · Datum uit dienst of einde 30% | eerste van: expat kolom N `Datum uit Dienst`, kolom V `Regeling Tm` en 31 december van het berekeningsjaar |
| B3 · Toetsloon | op basis van expat kolom T `Expatregeling`: code 913 → instelling *Toetsloon regeling 913*, code 914 → instelling *Toetsloon regeling 914*. Andere of ontbrekende code → status *Handmatig controleren*. Wijkt kolom AD `Grenswaarde` af van het ingestelde bedrag, dan volgt een aandachtspunt (het ingestelde bedrag blijft leidend) |
| B14 · LC 9970 (totaal SVW-loon) | historie: regel met `MasterLooncode` 9970 (kolom L), waarde uit `cumulatief` (kolom AB) |
| B15 · LC 5990 (netto 30%) | historie: regel met `MasterLooncode` 5990 (kolom L), waarde uit `cumulatief` (kolom AB) |

Kolommen worden herkend op **kolomnaam**; ontbreekt een naam, dan valt de tool terug op de
vaste kolomposities hierboven en meldt dat.

## Berekening

Eén-op-één overgenomen uit het Excel-rekenmodel, inclusief de NETWORKDAYS/DATEDIF-logica
voor gebroken maanden (B7–B9), toetsloon pro rato (B12), niet afgelaagd belastbaar loon
(B16 = B14 + B15), reeds toegepaste aflaging (B18), verschil (B20, afgerond op hele euro's),
nog extra af te lagen (B21) en het **nieuwe percentage** (B22, afgekapt op het maximum en
berekend op 2 decimalen). Bij een percentage onder het maximum toont de tool de actie:
*TWK aanmaken met het nieuwe percentage op LC 9535 (SYSLC904), alle maanden.*

### Balkenendenorm (WNT-norm)

Aanvullend op het Excel-model past de tool de maximering van de Balkenendenorm toe:
de onbelaste vergoeding is gemaximeerd op *maximaal percentage × min(niet afgelaagd
loon, norm pro rata)*. De norm (instelbaar; 2025: € 246.000) wordt pro rata herleid
naar de looptijd met dezelfde SV-dagenmethodiek als het toetsloon. Ligt het hierdoor
begrensde percentage onder het reguliere nieuwe percentage, dan geldt het begrensde
percentage en markeert de tool dit bij de medewerker. Veld leeg of 0 = norm niet
toepassen.

## Instellingen

- **Berekeningsjaar** (A1) en **standaard startdatum** (ondergrens voor B5, aanpasbaar).
- **Toetsloon regeling 913** en **Toetsloon regeling 914** — jaarlijks wettelijke
  bedragen (2025: € 35.468 laag, € 46.660 hoog); per medewerker gekozen op basis van
  kolom T. Het afgeleide *toetsloon maximaal 30%* (÷ 0,7) wordt per veld live getoond.
- **Maximaal percentage** — 30; per 2027 wordt dit 27, dan hier aan te passen.
- **Balkenendenorm** — WNT-norm op jaarbasis (2025: € 246.000); 0 = niet toepassen.
- **Looncodes** voor SVW-loon (9970) en netto 30% (5990).
- **Melding bij te hoge aflaging** — geeft (standaard aan) een aandachtspunt per
  medewerker wanneer het reeds toegepaste percentage (B18) hoger is dan het nieuwe
  percentage (B22), vergeleken op 2 decimalen: er is dan al meer afgelaagd dan
  toegestaan en de TWK leidt tot een correctie omlaag.
- Instellingen worden lokaal onthouden (localStorage van de browser).

## Aandachtspunten / bewuste keuzes

- De Balkenendenorm-maximering zit — anders dan in de oorspronkelijke Excel — wél in de
  berekening, op basis van het instelbare normbedrag (jaarlijks controleren).
- Percentages worden berekend en getoond met 2 decimalen.
- NETWORKDAYS rekent zonder feestdagenkalender, exact zoals het Excel-model.
- Bekende beperking uit het Excel-model is overgenomen én wordt gesignaleerd: start op de
  eerste *werkdag* van een maand die niet de 1e is, telt die maand niet als volledige maand.
- Wisselen regeling-/dienstverbandgegevens tussen perioden (herrekeningen), dan gebruikt de
  tool de laatste periode en markeert dit als aandachtspunt.
- Medewerkers zonder match of zonder looncode 9970/5990 in de historie krijgen de status
  *Handmatig controleren*; er wordt dan niets berekend.

## Validatie

`referentie/berekening_referentie.py` is een onafhankelijke Python-implementatie van
dezelfde berekening, bedoeld voor controle/audit:

```
python3 referentie/berekening_referentie.py "Expatregeling 2025.csv" "Historisch Overzicht Export 2025.csv" --jaar 2025
```

De browserlogica en dit script geven identieke uitkomsten; beide reproduceren het
rekenvoorbeeld uit de originele Excel exact.
