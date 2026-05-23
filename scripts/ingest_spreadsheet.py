#!/usr/bin/env python3
"""Ingest a MIP spreadsheet into the standalone atlas HTML.

The spreadsheet does not include latitude/longitude. This script preserves geo data
from the current HTML by matching schools, with city-level fallback coordinates.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

FIELDS = [
    "name", "district", "address", "city", "state", "zip", "country", "grades",
    "type", "phone", "year", "mandarin", "model", "otherLang", "website",
    "language", "lat", "lng", "approx",
]


def clean(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).replace("\xa0", " ").strip()


def title_model(v: str) -> str:
    v = clean(v)
    return {"strand": "Strand", "whole": "Whole School", "whole school": "Whole School"}.get(v.lower(), v)


COUNTRY_ALIASES = {
    "u.k.": "United Kingdom",
    "uk": "United Kingdom",
    "usa": "USA",
    "u.s.a.": "USA",
    "united states": "USA",
}


def canonical_country(v: str) -> str:
    v = clean(v)
    return COUNTRY_ALIASES.get(v.lower(), v)


def norm(*parts: str) -> str:
    fixed = [canonical_country(p) if i == len(parts) - 1 else clean(p) for i, p in enumerate(parts)]
    return "|".join(re.sub(r"[^a-z0-9]+", " ", p.lower()).strip() for p in fixed)


def extract_const(html: str, name: str, next_name: str) -> Any:
    m = re.search(rf"const {name} = (.*?);\nconst {next_name}\b", html, re.S)
    if not m:
        raise SystemExit(f"Could not find const {name} before const {next_name}")
    return json.loads(m.group(1))


def read_existing(html_path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    html = html_path.read_text(encoding="utf-8")
    schools = extract_const(html, "SCHOOLS", "WORLD")
    labels = extract_const(html, "LABELS", "C_MANDARIN")
    return schools, labels


def base_school(language: str = "Mandarin") -> dict[str, Any]:
    return {
        "name": "", "district": "", "address": "", "city": "", "state": "", "zip": "",
        "country": "USA", "grades": "", "type": "", "phone": "", "year": "",
        "mandarin": "", "model": "", "otherLang": "", "website": "",
        "language": language, "lat": None, "lng": None, "approx": True,
    }


def parse_us(ws) -> list[dict[str, Any]]:
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not clean(row[0] if row else None):
            continue
        s = base_school("Mandarin")
        s.update({
            "name": clean(row[0]), "district": clean(row[1]), "address": clean(row[2]),
            "city": clean(row[3]), "state": clean(row[4]), "zip": clean(row[5]),
            "grades": clean(row[6]), "type": clean(row[7]), "phone": clean(row[8]),
            "year": clean(row[9]), "mandarin": clean(row[10]), "model": title_model(row[11]),
            "otherLang": clean(row[13]), "website": clean(row[14]),
        })
        out.append(s)
    return out


def parse_cantonese(ws) -> list[dict[str, Any]]:
    out = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not clean(row[0] if row else None):
            continue
        s = base_school("Cantonese")
        district = clean(row[2]) or clean(row[1])
        address = clean(row[3])
        s.update({
            "name": clean(row[0]), "district": district, "address": address,
            "city": clean(row[4]), "state": clean(row[5]), "zip": clean(row[6]),
            "grades": clean(row[13]) or clean(row[9]), "type": clean(row[7]), "phone": clean(row[8]),
            "year": clean(row[10]), "mandarin": "", "model": title_model(row[11]),
            "otherLang": "", "website": clean(row[14]),
        })
        out.append(s)
    return out


def parse_international(ws) -> list[dict[str, Any]]:
    out = []
    current_country = ""
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = clean(row[0] if row else None)
        nonempty = [clean(v) for v in row if clean(v)]
        if not nonempty:
            continue
        # Country separator rows have only the country name in column A.
        if name and len(nonempty) == 1:
            current_country = name
            continue
        if not name:
            continue
        s = base_school("Mandarin")
        country = canonical_country(clean(row[6]) or current_country)
        s.update({
            "name": name, "district": clean(row[2]), "address": clean(row[3]),
            "city": clean(row[4]), "state": clean(row[5]), "zip": clean(row[7]),
            "country": country or "", "grades": clean(row[10]), "type": clean(row[8]),
            "phone": clean(row[9]), "year": clean(row[11]), "mandarin": clean(row[14]),
            "model": title_model(row[15]), "otherLang": "", "website": clean(row[12]),
        })
        out.append(s)
    return out


def geo_indexes(existing: list[dict[str, Any]]):
    exact, by_name, by_name_only, by_city = {}, {}, {}, defaultdict(list)
    for s in existing:
        if s.get("lat") is None or s.get("lng") is None:
            continue
        exact[norm(s.get("name", ""), s.get("city", ""), s.get("state", ""), s.get("country", ""))] = s
        by_name.setdefault(norm(s.get("name", ""), s.get("country", "")), s)
        by_name_only.setdefault(norm(s.get("name", "")), s)
        by_city[norm(s.get("city", ""), s.get("state", ""), s.get("country", ""))].append(s)
    return exact, by_name, by_name_only, by_city


def attach_geo(schools: list[dict[str, Any]], existing: list[dict[str, Any]]) -> list[dict[str, Any]]:
    exact, by_name, by_name_only, by_city = geo_indexes(existing)
    missing = []
    for s in schools:
        hit = exact.get(norm(s["name"], s["city"], s["state"], s["country"]))
        if not hit:
            hit = by_name.get(norm(s["name"], s["country"]))
        if not hit:
            hit = by_name_only.get(norm(s["name"]))
        if hit:
            s["lat"], s["lng"], s["approx"] = hit["lat"], hit["lng"], bool(hit.get("approx", False))
            continue
        city_hits = by_city.get(norm(s["city"], s["state"], s["country"]), [])
        if city_hits:
            s["lat"] = round(sum(float(x["lat"]) for x in city_hits) / len(city_hits), 4)
            s["lng"] = round(sum(float(x["lng"]) for x in city_hits) / len(city_hits), 4)
            s["approx"] = True
        else:
            missing.append(s)
    return missing


def make_city_labels(schools: list[dict[str, Any]]) -> list[dict[str, Any]]:
    buckets = defaultdict(list)
    for s in schools:
        if s.get("lat") is None or s.get("lng") is None or not s.get("city"):
            continue
        buckets[(s["city"], s.get("state", ""), s.get("country", ""))].append(s)
    labels = []
    for (city, state, country), rows in buckets.items():
        labels.append({
            "name": city,
            "lat": round(sum(float(r["lat"]) for r in rows) / len(rows), 3),
            "lng": round(sum(float(r["lng"]) for r in rows) / len(rows), 3),
            "n": len(rows),
        })
    labels.sort(key=lambda x: (-x["n"], x["name"]))
    return labels[:100]


def replace_const(html: str, name: str, next_name: str, value: Any) -> str:
    payload = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    pattern = rf"const {name} = .*?;\nconst {next_name}\b"
    replacement = f"const {name} = {payload};\nconst {next_name}"
    return re.sub(pattern, lambda _m: replacement, html, count=1, flags=re.S)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("spreadsheet", type=Path)
    ap.add_argument("--html", type=Path, default=Path("Chinese-Immersion-Atlas.html"))
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--allow-missing-geo", action="store_true")
    args = ap.parse_args()

    existing, labels = read_existing(args.html)
    wb = load_workbook(args.spreadsheet, read_only=True, data_only=True)
    schools = parse_us(wb["MIP List"]) + parse_international(wb["International"]) + parse_cantonese(wb["Cantonese"])
    missing = attach_geo(schools, existing)

    if missing and not args.allow_missing_geo:
        print(f"Missing geo for {len(missing)} schools; HTML not changed:", file=sys.stderr)
        for s in missing[:25]:
            print(f"- {s['name']} ({s.get('city')}, {s.get('state')}, {s.get('country')})", file=sys.stderr)
        if len(missing) > 25:
            print(f"... and {len(missing)-25} more", file=sys.stderr)
        print("Add matching schools/coords to the current HTML or rerun with --allow-missing-geo.", file=sys.stderr)
        return 2

    schools = [{k: s[k] for k in FIELDS} for s in schools if s.get("lat") is not None and s.get("lng") is not None]
    labels["cities"] = make_city_labels(schools)

    print(f"Parsed {len(schools)} schools ({Counter(s['language'] for s in schools)})")
    print(f"Missing geo: {len(missing)}")
    if args.dry_run:
        return 0

    html = args.html.read_text(encoding="utf-8")
    html = replace_const(html, "SCHOOLS", "WORLD", schools)
    html = replace_const(html, "LABELS", "C_MANDARIN", labels)
    args.html.write_text(html, encoding="utf-8")
    print(f"Updated {args.html}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
