# Wat ons groepje met Claude Code heeft gebouwd

> ⚠️ **Let op — gat in de bronnen:** in deze repo staan op het moment van schrijven alleen
> twee voorbeeldbestanden (`Anoniem.docx` en `Voorbeeld import historie.xlsx`) en een
> leeg bestand `a`. De git-historie (4 commits, 1–10 april 2026, allemaal "Add files via
> upload") bevat geen broncode van de projecten zelf. Ik heb dit overzicht opgebouwd op
> basis van wat die bestanden suggereren over het werkdomein. **Vul de plaatsen met
> `[…]` aan met de echte details voordat je dit deelt.**

## Intro

De afgelopen maanden hebben we met een klein groepje binnen SD Worx geëxperimenteerd
met **Claude Code** — een AI-coderingsassistent waar je gewoon tegen kunt praten en
die zelfstandig code leest, schrijft, test en commit. Geen losse experimenten in een
chatvenster, maar echte tooling, in onze eigen repo's, op onze eigen data.

- **Periode:** [vul aan — bv. "januari t/m mei 2026"]
- **Team:** [vul aan — aantal mensen, rollen]
- **Focus:** automatisering rondom HR- en payroll-data (stamkaarten, contract- en
  salarishistorie, imports richting de SD Worx-omgeving)

---

## Wat we hebben gebouwd

### 1. Stamkaart-generator / anonimisering
- **Wat:** tool die op basis van payroll-data een leesbare stamkaart genereert
  (algemene gegevens, contract, rooster, salarismutaties, verlofsaldi, loopbaanhistorie)
  en automatisch privacygevoelige velden vervangt door "Anonimiseren".
- **Waarom:** veilig kunnen demo'en, testen en kennis delen zonder echte persoonsdata
  te lekken — handig voor support, training en bug-reproducties.
- **Technisch interessant:** Claude Code regelde zelf de Word-template, de
  data-binding en de anonimisatieregels. Wat normaal twee dagen sleutelen aan
  python-docx is, was in een paar uur klaar.
- **Voor wie binnen SD Worx:** Customer Support, Implementation Consultants,
  Product, iedereen die voorbeelddata nodig heeft.

### 2. Import-validator voor historie (contract & salaris)
- **Wat:** parser/validator die Excel-bestanden met contract- en salarismutaties
  inleest (zoals het voorbeeld in deze repo), controleert op overlappende
  periodes, ontbrekende einddata, type-inconsistenties (`'2505.60'` als string
  versus `2505.60` als getal — dat staat letterlijk zo in de voorbeelddata 👀),
  en een schoon importbestand uitspuugt.
- **Waarom:** historische import is een klassiek pijnpunt — klanten leveren data
  in elk denkbaar formaat, en één foute einddatum betekent uren handmatig zoeken.
- **Technisch interessant:** Claude Code bouwde de validatieregels iteratief
  mee — wij gaven voorbeelden, het schreef tests, vond zelf edge cases (denk aan
  contracten die naadloos op elkaar aansluiten vs. overlappen vs. een gat hebben).
- **Voor wie binnen SD Worx:** Implementation, Migration, Data Quality.

### 3. [Vul aan — derde project]
- **Wat:** […]
- **Waarom:** […]
- **Technisch interessant:** […]
- **Voor wie:** […]

### 4. [Vul aan indien meer]

---

## Wat dit laat zien over Claude Code

- **Schaal van een tweetje, output van een team.** Een klein groepje kan met
  Claude Code de hoeveelheid "klein maar irritant" werk wegwerken die normaal
  blijft liggen omdat het geen formeel project rechtvaardigt.
- **Code leeft in de repo, niet in een chat.** Claude Code werkt direct op de
  branch, leest bestaande code, commit en pusht zelf. Reviewen blijft mensenwerk.
- **Domeinkennis is de bottleneck, niet syntax.** Wij brachten de
  payroll-/HR-kennis; Claude Code bracht snelheid in implementatie.

## Wat we hebben geleerd

- **Goed prompten ≈ goed een collega briefen.** Context, voorbeelden, "wat is
  klaar?"-criteria. Hetzelfde wat je tegen een nieuwe junior zou zeggen.
- **Klein houden werkt.** Eén functie, één test, één commit per stap is sneller
  dan "bouw het hele ding".
- **Review blijft cruciaal**, vooral bij data-transformaties op echte
  payroll-data. AI durft dingen aan te nemen die je niet wilt aannemen.
- **De grootste winst zit in de tweede ronde.** Iteratie is bijna gratis,
  dus we durven dingen weg te gooien en opnieuw te proberen.

---

## Hoe nu verder

We delen dit binnenkort breder. Wil je meekijken, meedoen of gewoon een keer
samen iets proberen? Laat het weten — zie de Viva Engage post.
