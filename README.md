# Brandweer Soest — Bezetting­simulatie

Interactieve website om te onderzoeken hoe vaak er "ruis" ontstaat tussen
**vaste piket­leden** (8 personen: 6× TS, 2× RV) en **overige
vrijwilligers** bij verschillende oproep­regels, op basis van de echte
incidenten in [`Incidenten.xlsx`](Incidenten.xlsx).

## Snel starten

De site is statisch (HTML + JS, geen server nodig). Twee opties:

### 1. Lokaal openen via mini-webserver

Een browser mag `Incidenten.xlsx` niet rechtstreeks van het bestandssysteem
laden, dus serveer de map even kort:

```powershell
cd SimulationBRW
python -m http.server 8000
# open http://localhost:8000/website/
```

### 2. GitHub Pages

Push de repo en zet GitHub Pages aan op de `main`-branch (root). De site
staat dan op `https://<jouwgebruiker>.github.io/<repo>/SimulationBRW/website/`.

> Werkt geen webserver? Klik op **"Of upload eigen .xlsx"** en kies
> `Incidenten.xlsx` zelf — dan heb je de server niet nodig.

## Hoe werkt de simulatie?

1. De website leest het werkblad `Incidenten` uit `Incidenten.xlsx`.
2. Per incident worden de 6-cijferige P2000-codes (`09xxxx`) uit de melding
   getrokken; alleen incidenten met minstens één Soest-voertuig blijven over.
3. Voor elk incident wordt op basis van de gekozen **regels** uitgerekend
   welke piket­leden en welke overige vrijwilligers nodig zijn.
4. Een incident telt als **ruis** wanneer er tegelijk
   *piket­leden thuis blijven* én *overige vrijwilligers worden opgeroepen*.

### Scenario's (paneel 3)

| Scenario              | Gedrag |
|-----------------------|--------|
| **Huidig**            | 1e TS → vaste TS-piket; 2e TS, RV (na piket vol) en HA → overige vrijwilligers. Bij deelalarm hebben de thuisblijvers wél eerste recht op de HA. |
| **Vast eerst altijd** | Vaste piket­leden krijgen voorrang op élk voertuig. Overige vrijwilligers vullen pas aan als alle 8 piket­leden bezet zijn. |
| **Aangepast**         | Alleen 1e TS is voor het TS-piket; verder is alles vrij. |

Alle aantallen (6 voor TS, 2 voor RV, capaciteiten, deelalarm-grootte) zijn
in paneel 3 instelbaar.

### Voertuig-mapping (paneel 2)

Standaard­mapping van de Soest-codes:

| Code     | Rol     | Opmerking                  |
|----------|---------|----------------------------|
| `093334` | HP-TS   | Huidige HP-TS (266)        |
| `093341` | HP-TS   | Oude HP-TS                 |
| `093336` | HP-HA   | HP-HA (814)                |
| `093351` | HP-RV   | HW (haakwagen, 257)        |
| `093352` | HP-RV   | AL (autoladder)            |
| `093431` | NP-TS   | Huidige NP-TS (265)        |
| `093441` | NP-TS   | Oude NP-TS                 |

Pas labels of rollen aan in de tabel; vink **deelalarm** aan voor codes
waarbij maar 4 van de 6 TS-piket­leden meegaan.

## Bestandsstructuur

```
SimulationBRW/
├── Incidenten.xlsx          # bron-data
├── README.md
└── website/
    ├── index.html
    ├── app.js               # parsing + simulatie + grafieken
    └── style.css
```

Externe libraries (via CDN, geen install nodig):
[SheetJS](https://sheetjs.com/) voor het inlezen van Excel,
[Chart.js](https://www.chartjs.org/) voor de grafieken.

## Privacy

Alle verwerking gebeurt in de browser. Het Excel-bestand wordt nooit naar
een server gestuurd.
