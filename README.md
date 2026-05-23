# Chinese Immersion Atlas

An interactive web map of Chinese immersion programs.

## Files

- `Chinese-Immersion-Atlas.html` — standalone interactive atlas page.
- `mip-list-2026-1-17-1.xlsx` — source spreadsheet downloaded from the Mandarin Immersion Parents Council.

## Run locally

Open `Chinese-Immersion-Atlas.html` in a browser, or serve the folder locally:

```bash
python3 -m http.server 8000
```

Then visit <http://localhost:8000/Chinese-Immersion-Atlas.html>.

## Ingest a new spreadsheet

The atlas HTML currently embeds school data in a JavaScript `SCHOOLS` constant. To refresh it from a newer MIP spreadsheet:

```bash
python3 -m pip install -r requirements.txt
./scripts/ingest_spreadsheet.py path/to/latest-mip-list.xlsx
```

The spreadsheet does not include latitude/longitude, so the script uses `data/geocodes.json` plus coordinates preserved from the existing HTML. Run with `--dry-run` first to check for unmatched schools.

To fetch coordinates for new schools via Nominatim/OpenStreetMap:

```bash
./scripts/ingest_spreadsheet.py path/to/latest-mip-list.xlsx --geocode-missing
```

Review new `data/geocodes.json` entries after geocoding; automated geocoding can return approximate or incorrect matches.

## Tests

```bash
python3 -m unittest discover -s tests -v
```

## Data source

Spreadsheet downloaded from:

<https://miparentscouncil.org/wp-content/uploads/2026/01/mip-list-2026-1-17-1.xlsx>
