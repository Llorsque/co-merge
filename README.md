# Vereniging Matcher (webtool)

Een eenvoudige, browser-gebaseerde tool om twee Excel-bestanden met verenigingen te vergelijken:

- **Aanbieders / KVK-export** (bijv. vanuit KVK / eigen systeem)
- **CO non-profit** (bijv. dataset van non-profitorganisaties)

De tool draait volledig in de browser (geen server nodig) en kan op GitHub Pages worden gehost.

## Functionaliteit

Je uploadt:

1. Een **Aanbieders-bestand** (Excel, `.xlsx`) met o.a. kolommen:

   - `DOSSIERNR` – KVK-nummer (mag ook numeriek zijn)
   - `FANAAM` – officiële naam
   - `VENNAAM` – handelsnaam (optioneel)
   - `RECHTSVORM`
   - `GEMEENTE`
   - `WOONPLAATS`
   - `POSTCODE`
   - `STRAAT`
   - `HUISNR`
   - `HUISNRTOEV`
   - `TELNR`
   - `EMAIL`

2. Een **CO non-profit-bestand** met o.a.:

   - `Nr.` – volgnummer
   - `KvK-nummer`
   - `Naam`
   - `Subsoort organisatie`
   - `Vestigingsgemeente`
   - `Telefoonnr.`
   - `E-mail`
   - `Postadres`
   - `Geblokkeerd`

De tool:

1. Normaliseert KVK-nummers naar een 8-cijferige sleutel (`KVK8`).
2. Maakt drie groepen:
   - **Tabblad 1 – KVK-match**  
     CO-rijen met KVK8 die ook in Aanbieders voorkomen.  
     Gegevens uit beide bestanden worden naast elkaar gezet.
   - **Tabblad 2 – Geen KVK in CO**  
     CO-rijen zonder KVK-nummer.
   - **Tabblad 3 – Wel KVK in CO, geen match in Aanbieders**  
     CO-rijen met KVK8 die niet in Aanbieders zitten.
3. Voor Tabblad 3 voert de tool een **naam-match** uit tegen alle namen in Aanbieders:
   - Namen worden opgeschoond (lowercase, quotes weg, spaties netjes)
   - Similarity wordt berekend o.b.v. bigram-Dice-coëfficiënt (0–100)
   - Alleen matches met **score ≥ 90** worden meegenomen.
   - Er wordt een extra controle gedaan op **gemeente** (genormaliseerd).
   - Deze resultaten komen in **Tabblad 4 – Naam-match**.
4. Voor Tabblad 1 wordt een kolom `NaamStatus` toegevoegd:
   - `Exact` – naam in CO en officiële naam in Aanbieders zijn 1-op-1 gelijk
   - `Bijna` – na kleine opschoning (lowercase, quotes weg, spaties) gelijk
   - `Anders` – inhoudelijk andere naam

Resultaat: een nieuw Excel-bestand (`resultaat_matcher.xlsx`) met:

- `Samenvatting`
- `Tabblad1_KVK_match`
- `Tabblad2_Geen_KVK`
- `Tabblad3_Wel_KVK_geen_match`
- `Tabblad4_Naam_match`

## Gebruiksaanwijzing (lokaal)

1. Download deze map (of de `.zip`) en pak uit.
2. Open `index.html` in je browser (Chrome / Edge / Firefox).
3. Kies:
   - Aanbieders-bestand
   - CO non-profit-bestand
4. Klik op **"Vergelijk bestanden"**.
5. Wacht tot de status aangeeft dat het klaar is.
6. Klik op **"Download resultaat.xlsx"**.

## Host op GitHub Pages

1. Maak een nieuwe GitHub-repository (bijv. `vereniging-matcher`).
2. Upload de bestanden:
   - `index.html`
   - `script.js`
   - `styles.css`
   - `README.md`
3. In GitHub:
   - Ga naar **Settings → Pages**
   - Kies:
     - Source: **Deploy from a branch**
     - Branch: `main` (of `master`), directory `/root`
4. Na enkele minuten staat de site live op een URL zoals:

   `https://jouw-username.github.io/vereniging-matcher`

## Aannames / beperkingen

- De tool gaat ervan uit dat de kolomnamen zoals hierboven aanwezig zijn.
- Alleen het **eerste werkblad** van ieder Excel-bestand wordt gelezen.
- Fuzzy naam-match is bewust streng ingesteld (score ≥ 90) om onzinmatches te voorkomen.
- Alle logica draait in de browser – geen data wordt naar een server gestuurd.

## Aanpassen

- De drempel voor naam-match aanpassen:

  In `script.js`:

  ```js
  const threshold = 90;
  ```

- Extra kolommen tonen of combineren kan eenvoudig door de JSON-objecten bij `tab1`, `tab2`, `tab3`, `tab4` uit te breiden.

