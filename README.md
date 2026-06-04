# Commuter Flow Web

This folder contains a self-contained static browser app. You can upload the whole folder to a GitHub repository and host it with GitHub Pages.

## What it does

- loads local gmina boundary data
- loads local powiat boundary data
- loads local flow data from `XLSX`, `CSV`, or `TXT`
- lets the user choose gmina or powiat scope
- applies transfer, distance, and Top N filters
- exports a downloadable ZIP bundle with:
  - line `GeoJSON`
  - polygon `GeoJSON`
  - summary `TXT`
  - multi-sheet `XLSX`

## Important first-version constraint

This version is fully dependency-free and browser-only, so it does **not** read or write `GPKG` yet.

Instead it expects:

- gmina boundaries as `GeoJSON`
- powiat boundaries as `GeoJSON`
- flow data as `XLSX`, `CSV`, or `TXT`

That tradeoff is what keeps the app:

- hostable on GitHub Pages
- fully contained in one folder
- free of CDNs, build steps, and installs

## Expected input fields

### Gmina GeoJSON

Feature properties should include:

- `JPT_KOD_JE`
- `JPT_NAZWA_`

Fallback property names also supported:

- `kod_teryt`
- `code`
- `name`
- `nazwa_gminy`

### Powiat GeoJSON

Feature properties should include:

- `JPT_KOD_JE`
- `JPT_NAZWA_`

### Flow workbook

The app uses the first matching worksheet named like `Macierz przeplywow`, or the first worksheet if no close match exists.

The first five columns are interpreted as:

1. residence code
2. residence name
3. work code
4. work name
5. transfers

## Hosting on GitHub Pages

1. Put this folder in your repository.
2. Make sure `index.html` stays at the published root or published subfolder.
3. Enable GitHub Pages in repo settings.

## Next upgrade path

If you want full parity with the desktop tool, the next step is bundling a browser SQLite/WASM runtime inside this folder so the app can ingest and emit real `GPKG` files too.
