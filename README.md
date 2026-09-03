# SD Worx · Notulentool

Een hulpmiddel voor management assistenten dat vergadernotities, een Teams-transcript en een Copilot-recap omzet naar notulen in een vaste structuur: samenvatting, keypoints per onderwerp, besluiten, acties (tabel) en open punten. De tool bestaat uit één bestand (`index.html`), draait volledig in de browser en roept een Claude-model aan via de API van Anthropic of via Microsoft Foundry.

## Wat de tool doet

1. Je plakt of sleept je bronnen in drie vakken: **Notities** (verplicht), **Transcript** (optioneel) en **Copilot-recap** (optioneel). Ondersteunde bestanden: `.txt`, `.md`, `.docx` en voor het transcript ook `.vtt` (Teams-transcript).
2. Optioneel vul je vergadertitel, datum en aanwezigen in. Laat je die leeg, dan leidt het model ze af uit de bronnen.
3. Je klikt op **Genereer notulen**. De tool stuurt de bronnen in één aanroep naar het model en toont de notulen terwijl ze binnenkomen.
4. Je controleert de notulen, past ze zo nodig aan via **Bewerken** en kopieert of downloadt ze: als markdown, als platte tekst (voor Teams of e-mail), als `.md` of als Word-bestand (`.docx`).

Regels die het model meekrijgt: de vaste structuur wordt exact gevolgd, er wordt niets verzonnen, de notities zijn leidend, tegenstrijdigheden tussen bronnen worden benoemd, en ontbrekende eigenaren of deadlines krijgen de tekst `[NOG AAN TE VULLEN]`. Controleer de notulen altijd zelf; het model kan fouten maken.

## De tool openen

- **Dubbelklikken.** Sla `index.html` op je laptop op en dubbelklik erop. De tool opent in je standaardbrowser (Edge of Chrome). Je ziet dan `file:///...` in de adresbalk; dat is de bedoeling.
- **Gehost.** `index.html` kan ook op een interne webserver of SharePoint-pagina staan die bestanden ongewijzigd serveert. Er is geen backend nodig.

Vereisten: een recente versie van Edge, Chrome of Firefox, en internettoegang naar (1) het gekozen model-endpoint en (2) de CDN's `cdnjs.cloudflare.com`, `unpkg.com` en `fonts.googleapis.com` voor de bibliotheken en het lettertype. Zonder CDN blijft de tool werken, maar dan zonder `.docx`-import/-export en met een eenvoudige tekstweergave.

Eerste keer: klik rechtsboven op **Instellingen**, kies de provider, vul de API-key (en bij Foundry de endpoint-URL) in en klik op **Test verbinding**. Wil je de tool eerst zonder key uitproberen, zet dan **Mock-modus** aan of open de tool met `?mock=1` achter het adres. In mock-modus komt na één seconde een vast voorbeeld terug zonder API-call.

## Een API-key en endpoint verkrijgen

### Anthropic API

1. Ga naar de Claude Console van Anthropic (https://platform.claude.com, voorheen console.anthropic.com) en log in op de organisatie van SD Worx, of vraag de beheerder van die organisatie om een key.
2. Maak onder **API keys** een nieuwe key aan. De key begint met `sk-ant-` en wordt maar één keer getoond; bewaar hem veilig (bijvoorbeeld in een wachtwoordmanager).
3. De organisatie moet een betaalmethode of tegoed hebben; anders geeft de API een fout 400 met een melding over het saldo.
4. Kies in de tool de provider **Anthropic API**, plak de key en laat de modelnaam op `claude-sonnet-5` staan (of vul een andere modelnaam in).

De tool stuurt bij deze provider de headers `x-api-key`, `anthropic-version` en `anthropic-dangerous-direct-browser-access: true` mee. Die laatste is verplicht voor aanroepen rechtstreeks vanuit een browser.

### Microsoft Foundry

1. Je hebt een Azure-abonnement nodig en een rol op de Foundry-resource zoals **Foundry User** (voorheen Azure AI User) of **Cognitive Services User**. Meestal regelt IT dit.
2. Open het Foundry-portaal (https://ai.azure.com), ga naar **Discover** › **Models**, zoek een Claude-model (bijvoorbeeld `claude-sonnet-5`) en klik op **Deploy**. De deploymentnaam is standaard gelijk aan de modelnaam; je kunt een eigen naam kiezen.
3. Na het deployen ga je naar **Build** › **Models**, opent je deployment en kijkt op het tabblad **Details**. Daar staan de **Target URI** (het endpoint) en de **Key**.
4. Kies in de tool de provider **Microsoft Foundry** en vul in:
   - **Endpoint-URL**: de URL van je resource, in de vorm `https://<resource>.services.ai.azure.com/anthropic`. Je mag ook de volledige URL tot en met `/v1/messages` plakken; de tool vult het pad zelf aan en toont onder het veld naar welke URL de aanroepen gaan.
   - **Naam van de auth-header**: standaard `api-key`. Foundry accepteert ook `x-api-key`. Gebruik je een Microsoft Entra ID-token in plaats van een key, kies dan `Authorization`; de tool zet er automatisch `Bearer ` voor. Zo'n token haal je bijvoorbeeld op met `az account get-access-token --resource https://ai.azure.com --query accessToken -o tsv` en verloopt na ongeveer een uur.
   - **API-key**: de Key uit het portaal (of het Entra ID-token).
   - **Modelnaam**: de naam van je deployment.

Bij Foundry stuurt de tool alleen de auth-header, `anthropic-version` en `content-type` mee.

## Test verbinding en CORS

**Test verbinding** doet één minimale aanroep naar het endpoint (een vraag van één woord, hooguit 32 tokens) en vertaalt het resultaat naar gewone taal:

| Resultaat | Betekenis | Wat je doet |
|---|---|---|
| OK | Het endpoint antwoordde en staat aanroepen vanuit de browser toe. | Niets, je kunt aan de slag. |
| 401 | De key is onjuist, verlopen of hoort niet bij dit endpoint. | Controleer de key; bij Foundry ook of de key bij deze resource hoort. |
| 403 | De key of het account heeft geen rechten. | Vraag IT om de juiste rol of toegang tot het model. |
| 404 | Endpoint-URL of modelnaam klopt niet. | Controleer de URL en de model- of deploymentnaam. |
| 429 | Te veel verzoeken (rate limit). | Wacht een minuut en probeer opnieuw. |
| CORS-fout / niet bereikbaar | De browser kon het endpoint niet bereiken. | Zie hieronder. |

Een browser mag alleen met een ander adres praten als dat adres dat expliciet toestaat (dat heet CORS). De API van Anthropic staat browser-aanroepen toe. Voor Foundry hangt het af van hoe de resource is ingericht; dat hebben wij niet kunnen verifiëren. Krijg je bij Test verbinding de melding dat het endpoint niet bereikbaar is, loop dan dit na:

1. **URL en netwerk.** Controleer de endpoint-URL op typefouten. Probeer het buiten het bedrijfsnetwerk of zonder VPN; een bedrijfsproxy kan `api.anthropic.com` of `*.services.ai.azure.com` blokkeren. Vraag IT die adressen toe te staan.
2. **Anthropic API.** Werkt het ook buiten het bedrijfsnetwerk niet, controleer dan de key en de modelnaam (een verkeerde modelnaam geeft een 404, geen CORS-fout).
3. **Foundry.** Blijft de melding komen terwijl dezelfde key wel werkt vanuit een script of vanuit het Foundry-portaal, dan staat de resource geen browser-aanroepen toe. Vraag IT of de Azure-beheerder om CORS toe te staan voor het adres waarop de tool draait (bij dubbelklikken is dat de origin `null`), bijvoorbeeld door de resource achter Azure API Management met een CORS-beleid te zetten of door de tool op een toegestaan intern adres te hosten. Lukt dat niet, gebruik dan de provider Anthropic API. Een eigen tussenlaag (proxy of backend) valt buiten deze versie.
4. Wil je intussen de werking laten zien, zet dan Mock-modus aan.

Fouten tijdens het genereren worden op dezelfde manier leesbaar in het scherm getoond, met de melding van het endpoint erbij.

## STRUCTUUR en VOORBEELDEN invullen (beheerder)

Bovenaan het script in `index.html` staat een blok dat begint met `// === CONFIGURATIE ===` en eindigt met `// === EINDE CONFIGURATIE ===`. Alles wat je aan de notulen wilt veranderen, staat daar; de rest van de code hoef je niet aan te raken. Open het bestand in Kladblok, Notepad++ of VS Code.

- **`SYSTEEMPROMPT`**: de instructies voor het model (Nederlands). Pas regels aan of voeg ze toe.
- **`STRUCTUUR`**: het vaste sjabloon in markdown. Vervang de default door de echte opzet van SD Worx. Houd de kolommen van de actietabel (`Actie | Eigenaar | Deadline | Bron`) intact, want de systeemprompt verwijst ernaar. Tekst tussen `[ ]` is een invulplek.
- **`VOORBEELDEN`**: een lijst van complete voorbeeldnotulen die als few-shot meegaan. Standaard leeg. Zo vul je hem:

  ```js
  const VOORBEELDEN = [
  `# Notulen: Weekoverleg HR-team

  **Datum:** 12 mei 2026
  ... (het volledige voorbeeld, in exact dezelfde structuur als STRUCTUUR) ...`,

  `# Notulen: Stuurgroep payroll
  ... (tweede voorbeeld) ...`,
  ];
  ```

  Gebruik geanonimiseerde, echte notulen van goede kwaliteit. Eén tot drie voorbeelden volstaan. Houd ze in dezelfde structuur als `STRUCTUUR`, anders spreken sjabloon en voorbeelden elkaar tegen.
- **`MOCK_OUTPUT`**: het vaste voorbeeld dat in mock-modus wordt getoond. Vervang dit door notulen in jullie structuur, zodat de demo klopt.
- Verder staan er constanten voor het standaardmodel, de maximale outputlengte (`MAX_TOKENS`) en de waarschuwingsgrens voor grote invoer.

Let op: de teksten staan tussen backticks (`` ` ``). Een backtick ín een tekst schrijf je als `` \` ``. Sla het bestand op als UTF-8, herlaad de pagina en test eerst in mock-modus (laadt het bestand nog?) en daarna met een echte aanroep.

## Privacy

- Je invoer (notities, transcript, recap, metadata) blijft in het geheugen van de browser en wordt nergens opgeslagen: niet op schijf, niet in de browseropslag, niet bij ons. Na het sluiten van het tabblad is ze weg.
- De invoer verlaat de browser alleen richting het gekozen model-endpoint: `api.anthropic.com` (Anthropic API) of je eigen Foundry-resource (`*.services.ai.azure.com`), altijd via HTTPS. Daar gelden de gegevensvoorwaarden van die aanbieder; bij Foundry blijven prompts en antwoorden binnen Azure (bij een deployment die op Azure wordt gehost).
- De CDN's (cdnjs, unpkg) en Google Fonts leveren alleen bibliotheken en het lettertype. Daar wordt geen invoer naartoe gestuurd; ze zien wel je IP-adres en het verzoek om die bestanden.
- Alleen de **instellingen** worden bewaard, in de `localStorage` van deze browser op deze laptop: provider, endpoint-URL, naam van de auth-header, modelnaam, mock-modus en de API-key. De key staat daar leesbaar. Deel de laptop dus niet zonder eerst **Wis alle instellingen** te gebruiken, en behandel de key als een wachtwoord.
- Gedownloade `.md`- en `.docx`-bestanden komen in je map Downloads terecht; die beheer je zelf.

## Handmatig testen

In de map `test-input/` staan fictieve bronnen van een korte HR-implementatievergadering: `notities.txt`, `transcript.vtt` (Teams-formaat) en `recap.txt`. De recap bevat bewust een afwijkende go-livedatum, zodat je kunt zien of het model de tegenstrijdigheid benoemt. Gebruik ze zo:

1. Open `index.html?mock=1` (of zet Mock-modus aan) en sleep de drie bestanden in de vakken. Controleer dat het transcript is omgezet naar regels `Naam: tekst`.
2. Klik op **Genereer notulen** en probeer bewerken, kopiëren en beide downloads.
3. Zet Mock-modus uit, vul een key in, klik op **Test verbinding** en genereer opnieuw.

## Gemaakte aannames

- Standaardmodel is `claude-sonnet-5`, zoals gevraagd; de naam is vrij aan te passen. De tool stuurt geen extra parameters mee (geen thinking-, effort- of temperature-instellingen), zodat elke modelnaam werkt. Maximale outputlengte is 16.000 tokens (`MAX_TOKENS`).
- Foundry: het endpointformaat `https://<resource>.services.ai.azure.com/anthropic/v1/messages`, de auth-headers `api-key`/`x-api-key`/`Authorization: Bearer` en het gebruik van de deploymentnaam als model komen uit de documentatie "Claude in Microsoft Foundry" van Anthropic, geraadpleegd op 3 september 2026. Of een Foundry-resource browser-aanroepen (CORS) toestaat, is niet geverifieerd; daarom is Test verbinding de eerste stap.
- Entra ID-tokens worden ondersteund door ze handmatig te plakken; er zit geen aanmeldflow (MSAL) in deze versie. Zo'n token verloopt na ongeveer een uur.
- Bibliotheken: `mammoth` 1.12.2 en `marked` 18.0.11 via cdnjs (met unpkg als reserve), `docx` 9.7.1 via unpkg (niet beschikbaar op cdnjs). De bibliotheken laden asynchroon; als een CDN niet bereikbaar is, blijft de rest van de tool werken. Inter komt van Google Fonts, met Segoe UI als fallback.
- Het model-antwoord wordt als SSE-stream gelezen. Geeft een proxy geen stream door, dan wordt het antwoord in één keer verwerkt.
- Een bestand dat je in een vak laadt, vervangt de tekst die er al stond. Een `.txt` in het transcriptvak dat met `WEBVTT` begint, wordt ook als VTT verwerkt. In de VTT-verwerking worden opeenvolgende cues van dezelfde spreker samengevoegd en wordt een cue zonder sprekerlabel bij de vorige regel gevoegd.
- De waarschuwing bij grote invoer telt tekens (300.000), geen tokens, en blokkeert niet. Lange transcripten worden niet opgeknipt; bij zeer grote invoer kan het endpoint traag zijn of het verzoek weigeren.
- De datum uit het metadataveld gaat als "3 september 2026" naar het model. Bestandsnamen volgen `JJJJ-MM-DD_<titel>_notulen.md/.docx`: de datum uit het veld, anders vandaag; de titel uit het veld, anders uit de eerste kop van de notulen, anders `vergadering`.
- De Word-export zet koppen, alinea's, opsommingen en tabellen om, met Inter als lettertype en SD Worx-blauw voor koppen. Verdere opmaak is bewust weggelaten.
- Kopiëren naar het klembord werkt bij dubbelklikken (`file://`) en op HTTPS. Op een `http://`-adres zonder HTTPS valt de tool terug op de oudere kopieerfunctie van de browser.
- `SKILL.md` in deze repository (de SD Worx-brandbook) hoort niet bij de tool en is ongewijzigd gelaten.
