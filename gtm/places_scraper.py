#!/usr/bin/env python3
"""
Influence Hub — Google Places lead scraper (Dubai F&B venues).

Pulls cafes / restaurants by district from the Google Places API (New, v1
Text Search) and writes a CSV whose columns match the Leads template, so you
can paste the rows straight into inflnc_leads_template.xlsx.

Places gives you: name, type, district, rating, reviews, website, phone, address.
It does NOT give Instagram / UGC / decision maker / email — those columns are
left blank on purpose for manual enrichment (that's where the deal is won).

--------------------------------------------------------------------------
SETUP (one time)
  1. Google Cloud Console -> create/select a project.
  2. Enable "Places API (New)".   (APIs & Services -> Library)
  3. Create an API key (APIs & Services -> Credentials). Restrict it to the
     Places API. Billing must be enabled — Places is a paid API with a
     monthly free tier; this script uses the cheap "Text Search (Essentials/
     Pro)" SKU via a tight field mask. Watch your usage.
  4. Export the key:   export GOOGLE_MAPS_API_KEY="AIza...."

RUN
  python3 gtm/places_scraper.py                    # default districts + types
  python3 gtm/places_scraper.py --districts JVC "Business Bay" --max-per-query 60
  python3 gtm/places_scraper.py --types "specialty coffee" brunch --out gtm/leads_jvc.csv

Then open the CSV, sanity-check, and paste the rows under the header in the
Leads sheet. Fill the blank enrichment columns (IG, UGC, decision maker) and
the Score/Tier formulas will grade each lead automatically.
--------------------------------------------------------------------------
Stdlib only (urllib) — no pip install needed.
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.request
import urllib.error

ENDPOINT = "https://places.googleapis.com/v1/places:searchText"

# Field mask keeps the request on the cheaper SKU. Add fields only if you need
# them — each extra field can bump you to a pricier billing tier.
FIELD_MASK = ",".join([
    "places.id",
    "places.displayName",
    "places.formattedAddress",
    "places.primaryType",
    "places.primaryTypeDisplayName",
    "places.rating",
    "places.userRatingCount",
    "places.websiteUri",
    "places.nationalPhoneNumber",
    "places.internationalPhoneNumber",
    "places.businessStatus",
    "places.googleMapsUri",
    "nextPageToken",
])

# Default search seeds — segments that fit the ICP (visual, footfall-driven).
DEFAULT_TYPES = [
    "specialty coffee shop",
    "brunch cafe",
    "dessert cafe",
    "restaurant",
]
DEFAULT_DISTRICTS = [
    "JVC Dubai",
    "Business Bay Dubai",
    "Dubai Marina",
    "DIFC Dubai",
]

# CSV header — mirrors the INPUT columns of the Leads template, in order.
# (Score & Tier are formula columns in the sheet, so they are NOT emitted here.)
CSV_HEADER = [
    "Venue name", "Type", "District", "Locations", "IG handle", "IG followers",
    "Posts UGC?", "Google rating", "Google reviews", "New (<6 mo)?", "Website",
    "Phone", "WhatsApp", "Email", "Decision maker", "Role", "Best channel",
    "Status", "Last touch", "Next step", "Notes",
]

# Map a raw search-type seed to the template's Type dropdown value.
def map_type(seed, primary_type_display):
    s = (seed + " " + (primary_type_display or "")).lower()
    if "coffee" in s or "cafe" in s:
        if "dessert" in s or "bakery" in s or "patisserie" in s:
            return "Dessert / bakery"
        if "brunch" in s:
            return "Brunch / cafe"
        if "coffee" in s:
            return "Specialty coffee"
        return "Brunch / cafe"
    if "dessert" in s or "bakery" in s or "ice cream" in s:
        return "Dessert / bakery"
    if "bar" in s or "lounge" in s:
        return "Bar / lounge"
    if "restaurant" in s or "dining" in s:
        return "Restaurant (dine-in)"
    return "Other"


def post_json(payload, api_key):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(ENDPOINT, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("X-Goog-Api-Key", api_key)
    req.add_header("X-Goog-FieldMask", FIELD_MASK)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def search(seed_type, district, api_key, max_results, region="ae"):
    """Text Search with pagination. Returns list of place dicts."""
    query = f"{seed_type} in {district}"
    collected = []
    page_token = None
    while len(collected) < max_results:
        payload = {
            "textQuery": query,
            "regionCode": region,
            "languageCode": "en",
            "pageSize": 20,
        }
        if page_token:
            payload["pageToken"] = page_token
        try:
            body = post_json(payload, api_key)
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "ignore")[:300]
            print(f"  ! HTTP {e.code} for '{query}': {detail}", file=sys.stderr)
            break
        except urllib.error.URLError as e:
            print(f"  ! network error for '{query}': {e}", file=sys.stderr)
            break

        places = body.get("places", []) or []
        collected.extend(places)
        page_token = body.get("nextPageToken")
        if not page_token:
            break
        time.sleep(2)  # nextPageToken needs a short warm-up + be polite
    return collected[:max_results]


def to_row(place, seed_type, district_label):
    name = (place.get("displayName") or {}).get("text", "").strip()
    ptype_display = (place.get("primaryTypeDisplayName") or {}).get("text", "")
    rating = place.get("rating", "")
    reviews = place.get("userRatingCount", "")
    website = place.get("websiteUri", "")
    phone = place.get("nationalPhoneNumber") or place.get("internationalPhoneNumber") or ""
    address = place.get("formattedAddress", "")
    maps = place.get("googleMapsUri", "")
    status = place.get("businessStatus", "")

    note_bits = []
    if address:
        note_bits.append(address)
    if status and status != "OPERATIONAL":
        note_bits.append(f"[{status}]")
    if maps:
        note_bits.append(maps)
    notes = " | ".join(note_bits)

    return {
        "Venue name": name,
        "Type": map_type(seed_type, ptype_display),
        "District": district_label,
        "Locations": "",              # enrich manually (1 / 2-6 / 7+)
        "IG handle": "",              # enrich manually
        "IG followers": "",           # enrich manually
        "Posts UGC?": "",             # enrich manually (Yes/No)
        "Google rating": rating,
        "Google reviews": reviews,
        "New (<6 mo)?": "",           # enrich manually (Yes/No)
        "Website": website,
        "Phone": phone,
        "WhatsApp": "",               # often == phone; confirm before use
        "Email": "",                  # enrich manually
        "Decision maker": "",         # enrich manually
        "Role": "",
        "Best channel": "",
        "Status": "0 - New",
        "Last touch": "",
        "Next step": "Enrich: find IG + decision maker",
        "Notes": notes,
    }


def clean_district_label(d):
    # "Business Bay Dubai" -> "Business Bay"; helps match the sheet dropdown.
    return d.replace(" Dubai", "").replace("Dubai ", "").strip() or d


def main():
    ap = argparse.ArgumentParser(description="Scrape Dubai F&B venues from Google Places into a leads CSV.")
    ap.add_argument("--types", nargs="+", default=DEFAULT_TYPES,
                    help="Search seeds, e.g. 'specialty coffee' brunch dessert")
    ap.add_argument("--districts", nargs="+", default=DEFAULT_DISTRICTS,
                    help="Districts, e.g. JVC 'Business Bay' 'Dubai Marina'")
    ap.add_argument("--max-per-query", type=int, default=40,
                    help="Max results per (type x district) query (default 40)")
    ap.add_argument("--region", default="ae", help="Region bias code (default ae)")
    ap.add_argument("--out", default="gtm/leads_places.csv", help="Output CSV path")
    ap.add_argument("--min-reviews", type=int, default=0,
                    help="Drop venues with fewer than N Google reviews (default 0 = keep all)")
    args = ap.parse_args()

    api_key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if not api_key:
        print("ERROR: set GOOGLE_MAPS_API_KEY env var first. See header of this file.",
              file=sys.stderr)
        sys.exit(1)

    seen = set()          # dedup by place id
    rows = []
    for district in args.districts:
        label = clean_district_label(district)
        for seed in args.types:
            print(f"-> {seed} in {district} ...", file=sys.stderr)
            for place in search(seed, district, api_key, args.max_per_query, args.region):
                pid = place.get("id")
                if pid in seen:
                    continue
                if args.min_reviews and (place.get("userRatingCount") or 0) < args.min_reviews:
                    continue
                seen.add(pid)
                rows.append(to_row(place, seed, label))

    # Sort: most-reviewed first (rough proxy for "real, active venue").
    rows.sort(key=lambda r: (r["Google reviews"] or 0), reverse=True)

    out_dir = os.path.dirname(args.out)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_HEADER)
        w.writeheader()
        w.writerows(rows)

    print(f"\nDone. {len(rows)} unique venues -> {args.out}", file=sys.stderr)
    print("Next: open the CSV, sanity-check, paste rows under the header in the "
          "Leads sheet, then enrich the blank IG / decision-maker columns.",
          file=sys.stderr)


if __name__ == "__main__":
    main()
