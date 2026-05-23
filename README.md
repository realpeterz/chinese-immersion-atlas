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

The spreadsheet does not include latitude/longitude, so the script preserves coordinates from the existing HTML by matching school names and city/state/country. Run with `--dry-run` first to check for unmatched schools.

## Data source

Spreadsheet downloaded from:

<https://miparentscouncil.org/wp-content/uploads/2026/01/mip-list-2026-1-17-1.xlsx>
