# SD Worx · Notulentool

Een hulpmiddel voor management assistenten dat vergadernotities, een Teams-transcript en een Copilot-recap omzet naar notulen in het vaste sjabloon **SD Worx Nederland - Notulen MT**. De tool bestaat uit één bestand (`index.html`), draait volledig in de browser en heeft geen AI-koppeling, API-key of installatie nodig. Je voert de bronnen in, klikt op **Maak notulen** en downloadt het Word-bestand.

## Wat de tool doet

1. **Bronnen invoeren.** Je plakt of sleept je bronnen in drie vakken: **Notities** (verplicht), **Transcript** (optioneel) en **Copilot-recap** (optioneel). Ondersteunde bestanden: `.txt`, `.md`, `.docx` en voor het transcript ook `.vtt` (Teams-transcript, wordt automatisch omgezet naar regels per spreker). Optioneel vul je de kopgegevens in: titel, locatie, datum, notulist, aanwezig, gedeeltelijk aanwezig, afwezig en de link naar de MT Prep. Wat je leeg laat, haalt de tool uit de kop van je notities.
2. **Maak notulen.** De tool herkent de kopjes en opsommingen in je notities, koppelt ze aan de vaste agendapunten, haalt acties (met eigenaar en deadline), besluiten en afgeronde punten uit alle bronnen en zet alles in de structuur van het sjabloon. Het transcript en de recap vullen aan; de notities zijn leidend.
3. **Controleren en downloaden.** Je ziet de notulen direct, past ze aan via **Bewerken** en downloadt ze als Word-bestand in het sjabloon (`.docx`), als markdown (`.md`) of kopieert ze als platte tekst voor Teams of e-mail.

De tool herschrijft niets: elke zin in de notulen komt letterlijk uit een bron (alleen datums als `8/9` worden voluit geschreven en zinnen uit het transcript worden van "ik" naar de naam van de spreker gezet). Wat de tool niet kan afleiden, krijgt `[NOG AAN TE VULLEN]`. Controleer de notulen dus altijd; de statusregel telt hoeveel acties zijn gevonden en hoeveel invulplekken er zijn.

De output volgt de structuur van het Word-sjabloon `Notulensjabloon_1.docx`:

| Onderdeel | Inhoud |
|---|---|
| Kopblok | Titel, locatie en datum, aanwezig / gedeeltelijk aanwezig / afwezig, notulist, link naar de MT Prep |
| 1. Opening & doelstellingen | Vast agendapunt |
| 2. Prep MBR | Vast agendapunt |
| 3. KPI's | Vast agendapunt |
| 4. Lunch | Vast agendapunt ("Niet besproken." als er niets over in de bronnen staat) |
| 5. Strategische initiatieven | Vast agendapunt |
| 6. MT Topics & Besluitvorming | Vast agendapunt |
| 7, 8, … | Extra agendapunten: kopjes uit de notities die bij geen vast agendapunt horen, plus "Overige punten" voor losse punten uit transcript of recap |
| Openstaande acties | Tabel van alle acties (Actie, Eigenaar, Deadline, Bron) en de afgeronde actiepunten |
| Afsluiting | Altijd het laatste agendapunt |

Onder elk agendapunt staan eerst de keypoints, dan de besluiten (`**Besluit:** …`) en dan de acties (`**Actie:** wat – eigenaar – deadline`). Alle acties komen daarna samen in de tabel onder "Openstaande acties", met in de kolom Bron uit welke bronnen ze komen.

## De tool openen

- **Dubbelklikken.** Sla `index.html` op je laptop op en dubbelklik erop. De tool opent in je standaardbrowser (Edge of Chrome). Je ziet dan `file:///...` in de adresbalk; dat is de bedoeling.
- **Gehost.** `index.html` kan ook op een interne webserver of SharePoint-pagina staan die bestanden ongewijzigd serveert. Er is geen backend nodig.

Vereisten: een recente versie van Edge, Chrome of Firefox, en internettoegang naar de CDN's `cdnjs.cloudflare.com`, `unpkg.com` en `fonts.googleapis.com` voor de bibliotheken en het lettertype. Zonder CDN blijft de tool werken, maar dan zonder `.docx`-import en Word-export en met een eenvoudige tekstweergave.

Met **Probeer met voorbeeldbronnen** (of `index.html?voorbeeld=1`) vult de tool de vakken met een fictief MT-overleg en maakt hij meteen de notulen, zodat je kunt zien wat hij herkent en de downloads kunt uitproberen.

## Hoe de tool je bronnen leest

**Notities.** De tool zoekt kopjes: genummerde regels (`1. Opening`), korte regels gevolgd door opsommingstekens, regels in hoofdletters, regels die op een dubbele punt eindigen en regels als `Extra: …`. Elk kopje wordt aan een vast agendapunt gekoppeld op basis van herkenwoorden (bijvoorbeeld "MBR" bij Prep MBR, "KPI" of "NPS" bij KPI's, "besluit" of "budget" bij MT Topics). Kopjes die nergens bij horen, worden extra agendapunten. Alles vóór het eerste kopje is de kop van de notities: daar zoekt de tool naar `Aanwezig:`, `Afwezig:`, `Notulist:`, een datum, een locatie en een link. Staan er geen kopjes in, dan verdeelt de tool de losse regels op trefwoorden over de agendapunten en zet de rest onder "Overige punten".

**Per regel** bepaalt de tool wat het is:

| Soort | Herkend aan |
|---|---|
| Actie | Een naam (uit de aanwezigheidslijst) of groep ("alle directors", "iedereen") plus een datum of een actiewoord (levert, stuurt, plant, checkt, rondt af, …); of een regel die met "actie" begint; of een regel in een sectie "Acties" / "Actiepunten". Het deel na een pijl (`->`) wordt apart bekeken: "40 dossiers incompleet -> Fatima checkt, uiterlijk 8/9" geeft een keypoint én een actie. |
| Besluit | Een besluitwoord: akkoord, besloten, goedgekeurd, uitgesteld, doorgeschoven, vastgesteld, … (niet in vragen), of een regel die met "Besluit:" begint. |
| Afgerond actiepunt | Een regel die met "afgerond" of "gedaan" begint, of "is getekend / afgerond / opgeleverd / …" bevat, of in een sectie "Afgerond" staat. |
| Keypoint | Al het andere. |

Bij een actie zoekt de tool de **eigenaar** (de naam of namen in de regel, of de naam vóór de dubbele punt zoals in `Sanne: planning delen`) en de **deadline** (bij voorkeur de datum na "uiterlijk", "deadline", "voor" of "op"; anders de laatste datum in de regel). Ontbreekt een van beide, dan staat er `[NOG AAN TE VULLEN]`. Namen komen uit de kopvelden, de aanwezigheidsregels in de notities, de sprekers in het transcript en de deelnemerslijst in de recap; een voornaam in de notities wordt aangevuld tot de volledige naam.

**Copilot-recap.** Secties als "Belangrijkste punten", "Actiepunten", "Afgerond" en "Open vragen" worden herkend. Punten die al in de notities staan, worden overgeslagen; nieuwe punten komen bij het agendapunt met de meeste gedeelde woorden, anders onder "Overige punten". Acties uit de recap vullen ontbrekende eigenaren en deadlines aan.

**Transcript.** Alleen toezeggingen ("ik stuur … uiterlijk 12 september", "kun je … rapporteren?" gevolgd door "ja"), besluiten, afgeronde punten en de afsluiting ("het volgende MT is op …") worden gebruikt. Zinnen worden van de eerste naar de derde persoon gezet. Gewone gespreksregels blijven buiten de notulen.

**Samenvoegen en controleren.** Dezelfde actie uit meerdere bronnen wordt één regel in de tabel (zelfde eigenaar en overlappende woorden of dezelfde deadline), met alle bronnen in de kolom Bron. Noemen bronnen verschillende getallen bij hetzelfde begrip (bijvoorbeeld NPS 42 in de notities en 44 in de recap), dan komt er een regel "Let op: de bronnen verschillen over …" bij het betreffende agendapunt.

### Tips voor notities die goed worden herkend

- Gebruik kopjes die op de agenda lijken: `1. Opening`, `2. Prep MBR`, `3. KPI's`, … Een extra onderwerp zet je als `Extra: <onderwerp>` of als eigen genummerd kopje.
- Zet per kopje korte opsommingsregels (`- …`).
- Schrijf acties met naam en datum: `- Sanne stuurt de planning, uiterlijk 12/9` of `- actie Sanne: planning delen, 12/9`. Een pijl werkt ook: `- 40 dossiers incompleet -> Fatima checkt, 8/9`.
- Schrijf besluiten met "akkoord", "besloten" of `Besluit: …`.
- Zet bovenaan `Aanwezig: …`, `Afwezig: …`, `Notulist: …` en de datum. Of vul de kopvelden in de tool in; die gaan altijd voor.

## Configuratie (beheerder)

Bovenaan het script in `index.html` staat een blok dat begint met `// === CONFIGURATIE ===` en eindigt met `// === EINDE CONFIGURATIE ===`. Open het bestand in Kladblok, Notepad++ of VS Code.

- **`STANDAARD_TITEL`**: de vooringevulde titel.
- **`AGENDA`**: de vaste agendapunten, in volgorde, elk met een reguliere expressie (`herken`) waarmee kopjes uit de notities worden herkend. Voeg een agendapunt toe of pas de herkenwoorden aan. `AGENDA_ACTIES` en `AGENDA_AFSLUITING` zijn de twee vaste slotpunten; `TITEL_OVERIG` is de naam van het opvangpunt.
- **`ACTIE_WOORDEN`**, **`BESLUIT_WOORDEN`**, **`AFGEROND_WOORDEN`**, **`AFSLUIT_WOORDEN`**, **`DEADLINE_WOORDEN`**: de signaalwoorden. Voeg woorden toe die in jullie notities gebruikelijk zijn.
- **`GROEPEN`**: groepen die eigenaar van een actie kunnen zijn ("Alle directors", "Iedereen").
- **`VOORBEELD_BRONNEN`**: de fictieve bronnen achter "Probeer met voorbeeldbronnen".
- **`KOPVELDEN`** en **`SJABLOON_DOCX_BASE64`**: de koppeling met het Word-sjabloon, zie hieronder.

Sla het bestand op als UTF-8 en herlaad de pagina. Test met de voorbeeldbronnen en met echte notities.

### Hoe de Word-export werkt

**Download .docx** vult het echte Word-sjabloon (kop- en voettekst met logo, stijlen en nummering van `Notulensjabloon_1.docx`). Het sjabloon zit als base64-tekst in `index.html`, met plaatshouders op de plekken waar tekst komt:

| Plaatshouder | Wordt gevuld met |
|---|---|
| `{{titel}}` | De titel (kop `#` van de notulen) |
| `{{subtitel}}` | Locatie en datum uit de kopregels |
| `{{aanwezig}}`, `{{gedeeltelijk}}`, `{{afwezig}}` | De drie regels achter "Aan:" |
| `{{notulist}}` | De regel "Van:" |
| `{{link}}` | De regel "Link:", als klikbare hyperlink wanneer er een URL is |
| `{{inhoud}}` | Alles vanaf het eerste agendapunt |
| `{{voetdatum}}` (voettekst) | De datum als dd/mm/jjjj |

Bij het vullen van `{{inhoud}}` gebruikt de tool de stijlen van het sjabloon: agendapunten (`##`) worden genummerde koppen (stijl Heading 3; Word nummert zelf), subonderwerpen (`###`) worden Subheading 3, opsommingen krijgen de opsommingstekens van het sjabloon (het pijltje voor `Actie:` en `Besluit:`), en de actietabel wordt een Word-tabel. Lukt het vullen niet (bijvoorbeeld omdat een bibliotheek niet geladen is), dan maakt de tool een eenvoudig Word-bestand zonder sjabloon en meldt dat in de statusregel.

Een ander of aangepast sjabloon gebruiken:

1. Haal het huidige sjabloon uit `index.html`: kopieer de tekst tussen de aanhalingstekens van `SJABLOON_DOCX_BASE64` naar een bestand `sjabloon.b64` en voer in PowerShell uit: `[IO.File]::WriteAllBytes("sjabloon.docx", [Convert]::FromBase64String((Get-Content sjabloon.b64 -Raw)))`.
2. Open `sjabloon.docx` in Word en pas de opmaak aan. Laat de plaatshouders staan (elke plaatshouder in één stuk tekst, dus niet half vet of half in een ander lettertype) en houd de stijlen Heading 3, Subheading 3 en List Paragraph en de nummeringen van het sjabloon in stand.
3. Codeer het bestand opnieuw: `[Convert]::ToBase64String([IO.File]::ReadAllBytes("sjabloon.docx")) | Set-Content sjabloon.b64` en plak de inhoud in `SJABLOON_DOCX_BASE64`.

De markdown die de tool maakt, is ook de basis voor de export: de kopregels (`**Locatie:**` enzovoort, zie `KOPVELDEN`) en de agendapunten (`## 1. …`) moet je in **Bewerken** laten staan als je wilt dat de Word-export ze herkent.

## Privacy

- Alles gebeurt in de browser. Je invoer (notities, transcript, recap, kopgegevens) verlaat de laptop niet en wordt nergens opgeslagen: niet op schijf, niet in de browseropslag, niet bij een dienst. Na het sluiten van het tabblad is alles weg.
- De CDN's (cdnjs, unpkg) en Google Fonts leveren alleen bibliotheken en het lettertype. Daar wordt geen invoer naartoe gestuurd; ze zien wel je IP-adres en het verzoek om die bestanden.
- Gedownloade bestanden komen in je map Downloads terecht; die beheer je zelf.

## Handmatig testen

In de map `test-input/` staan fictieve bronnen van een MT-overleg van SD Worx Nederland: `notities.txt`, `transcript.vtt` (Teams-formaat) en `recap.txt`. Dezelfde bronnen zitten in de tool achter **Probeer met voorbeeldbronnen**. De recap noemt bewust een andere NPS dan de notities, zodat je de waarschuwing over verschillende bronnen ziet.

1. Open `index.html` en sleep de drie bestanden in de vakken. Controleer dat het transcript is omgezet naar regels `Naam: tekst`.
2. Klik op **Maak notulen** en bekijk wat er per agendapunt is herkend. Probeer **Bewerken**, **Weergave**, beide downloads en het kopiëren.
3. Open het Word-bestand en controleer dat kopblok, nummering en actietabel in het sjabloon staan.

## Gemaakte aannames

- Er is bewust geen AI-koppeling: de tool werkt met herkenregels (kopjes, opsommingen, namen, datums, signaalwoorden). Dat maakt hem voorspelbaar en volledig lokaal, maar hij vat niet samen en herschrijft niet. De kwaliteit van de notulen hangt af van de structuur van de notities; zie de tips hierboven.
- De structuur is afgeleid van `Notulensjabloon_1.docx`: de agendapunten 1 tot en met 6 zijn vast, kopjes uit de notities die nergens bij horen worden extra agendapunten, "Openstaande acties" en "Afsluiting" sluiten af. Acties staan zowel bij het agendapunt als in de tabel; de kolom "Bron" is toegevoegd om te zien uit welke bronnen een actie komt.
- Namen worden alleen herkend als ze in een aanwezigheidslijst, kopveld, sprekerlabel of deelnemerslijst voorkomen. Een regel met een onbekende naam en een datum wordt een keypoint, geen actie. Vul dus de aanwezigen in als de notities geen `Aanwezig:`-regel hebben.
- Datums als `8/9` krijgen het jaar van de vergadering (uit het datumveld of de kop van de notities), anders het huidige jaar. "In september" zonder dag geldt niet als deadline; "eind september" wel.
- "Gedeeltelijk aanwezig" en "afwezig" worden "geen" als er een aanwezigheidslijst is maar niemand als gedeeltelijk aanwezig of afwezig wordt genoemd.
- De Word-export vult het originele sjabloon via plaatshouders (bibliotheek `docx`, functie `patchDocument`). De hyperlinks naar SharePoint uit het originele sjabloon zijn vervangen door de link uit de kopregel. De datum in de voettekst komt uit het datumveld, anders uit de kopregel "Datum", anders is het vandaag. Het logo in de koptekst (EMF) blijft ongewijzigd.
- Bibliotheken: `mammoth` 1.12.2 en `marked` 18.0.11 via cdnjs (met unpkg als reserve), `docx` 9.7.1 via unpkg (niet beschikbaar op cdnjs). De bibliotheken laden asynchroon; als een CDN niet bereikbaar is, blijft de rest van de tool werken. Inter komt van Google Fonts, met Segoe UI als fallback.
- Een bestand dat je in een vak laadt, vervangt de tekst die er al stond. Een `.txt` in het transcriptvak dat met `WEBVTT` begint, wordt ook als VTT verwerkt. In de VTT-verwerking worden opeenvolgende cues van dezelfde spreker samengevoegd en wordt een cue zonder sprekerlabel bij de vorige regel gevoegd.
- Bestandsnamen volgen `JJJJ-MM-DD_<titel>_notulen.md/.docx`: de datum uit het veld, anders uit de kopregel "Datum" van de notulen, anders vandaag; de titel uit het veld (zonder het woord "notulen"), anders uit de eerste kop, anders `mt`.
- Kopiëren naar het klembord werkt bij dubbelklikken (`file://`) en op HTTPS. Op een `http://`-adres zonder HTTPS valt de tool terug op de oudere kopieerfunctie van de browser.
- `SKILL.md` in deze repository (de SD Worx-brandbook) hoort niet bij de tool en is ongewijzigd gelaten.
