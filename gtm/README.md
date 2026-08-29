# GTM — lead base + scraper

Two tools for building the Influence Hub lead list (F&B venues, Dubai first).

## 1. Lead base template — `inflnc_leads_template.xlsx`
Rebuild it any time with:
```bash
python3 gtm/build_leads_template.py
```
Sheets:
- **Leads** — the working list. Fill the **blue** columns; **Score** and **Tier**
  are computed automatically (A ≥6, B 4–5, C 1–3). Dropdowns + auto-filter built in.
- **ICP & Scoring** — the 5-point ICP checklist, the scoring table, and UAE
  compliance notes.
- **Lists (ref)** — dropdown source values (edit here to change the dropdowns).

Work top-down by Tier A → B → C. Score rewards: active IG but weak UGC (+3),
new venue (+2), small operator (+1), visual segment (+1), strong social proof
(+1), reachable decision maker (+1).

## 2. Places scraper — `places_scraper.py`
Pulls cafés/restaurants by district from the Google Places API (New) into a CSV
whose columns match the Leads sheet.

```bash
export GOOGLE_MAPS_API_KEY="AIza..."          # Places API (New) enabled + billing on
python3 gtm/places_scraper.py                 # default types × districts
python3 gtm/places_scraper.py --districts JVC "Business Bay" --min-reviews 30
```
Output: `gtm/leads_places.csv`. Places supplies name, type, rating, reviews,
website, phone, address. It does **not** supply Instagram / UGC / decision maker
— those columns stay blank for manual enrichment (that's where deals are won).

**Paste flow:** open the CSV → sanity-check → copy the rows → paste under the
header in the **Leads** sheet, starting at `A4`. The 21 CSV columns line up 1:1
with the sheet's input columns; the Score/Tier formula columns sit further right
and are untouched. Then enrich the blank columns and the tiering grades itself.

> Places is a paid API with a monthly free tier. The script uses a tight field
> mask to stay on the cheaper SKU — still, watch your Cloud billing usage.
