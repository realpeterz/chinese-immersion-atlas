# Chinese Immersion Atlas

An interactive web map of Chinese immersion programs.

## Files

- `Chinese-Immersion-Atlas.html` — standalone interactive atlas page.
- `data/source/mip-list-2026-1-17-1.xlsx` — source spreadsheet downloaded from the Mandarin Immersion Parents Council.
- `data/source/SOURCES.txt` — source URLs.

## Run locally

Open `Chinese-Immersion-Atlas.html` in a browser, or serve the folder locally:

```bash
python3 -m http.server 8000
```

Then visit <http://localhost:8000/Chinese-Immersion-Atlas.html>.

## Ingest a new spreadsheet

The atlas HTML currently embeds school data in a JavaScript `SCHOOLS` constant. To refresh it from a newer MIP spreadsheet:

```bash
uv run ./scripts/ingest_spreadsheet.py data/source/mip-list-2026-1-17-1.xlsx
```

The spreadsheet does not include latitude/longitude, so the script uses `data/geocodes.json` plus coordinates preserved from the existing HTML. Run with `--dry-run` first to check for unmatched schools.

To fetch coordinates for new schools via Nominatim/OpenStreetMap:

```bash
uv run ./scripts/ingest_spreadsheet.py data/source/mip-list-2026-1-17-1.xlsx --geocode-missing
```

Review new `data/geocodes.json` entries after geocoding; automated geocoding can return approximate or incorrect matches.

## Tests

```bash
uv run python -m unittest discover -s tests -v
```

## Data source

See `data/source/SOURCES.txt`.
