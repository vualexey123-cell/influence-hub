#!/usr/bin/env python3
"""
Populate the Leads template with the 10 real Dubai venues enriched from Google
Maps + Instagram (Aug 2026). Writes:
  - gtm/inflnc_leads_10_enriched.xlsx  (template + 10 filled rows, Score/Tier auto)
  - gtm/leads_enriched_10.csv          (same rows, template column order)

Blank cells = data not publicly available (IG followers for some, decision
maker/email) and left for manual enrichment on purpose — nothing is invented.

Run:  python3 gtm/build_leads_template.py   # ensure base template exists
      python3 gtm/populate_enriched_10.py
"""
import csv
import openpyxl
from openpyxl.styles import Font

TEMPLATE = "gtm/inflnc_leads_template.xlsx"
OUT_XLSX = "gtm/inflnc_leads_10_enriched.xlsx"
OUT_CSV = "gtm/leads_enriched_10.csv"

HEADER = [
    "Venue name", "Type", "District", "Locations", "IG handle", "IG followers",
    "Posts UGC?", "Google rating", "Google reviews", "New (<6 mo)?", "Website",
    "Phone", "WhatsApp", "Email", "Decision maker", "Role", "Best channel",
    "Status", "Last touch", "Next step", "Notes",
]

# Each row: only verified fields filled. "" = enrich manually.
ROWS = [
    {
        "Venue name": "Stella's Coffee House - JVC", "Type": "Specialty coffee",
        "District": "JVC", "Locations": "1 venue", "IG handle": "@stellas_coffeehouse",
        "IG followers": "", "Posts UGC?": "", "Google rating": 4.9, "Google reviews": 300,
        "New (<6 mo)?": "", "Website": "stellasdxb.com", "Phone": "+971 4 566 1080",
        "WhatsApp": "", "Email": "bassel@stellasdxb.com", "Decision maker": "Bassel",
        "Role": "Owner", "Best channel": "Email", "Status": "1 - Enriched",
        "Next step": "Personal email to Bassel + 2 profiles",
        "Notes": "Dog cafe, owner actively replies to reviews. Personal email bassel@ on the domain = likely owner (role inferred). Also TikTok @stellas_coffeehouse. Al Barsha South Fourth, JVC. Landline only.",
    },
    {
        "Venue name": "VITOR CASTRO - JVC", "Type": "Dessert / bakery",
        "District": "JVC", "Locations": "1 venue", "IG handle": "@vitorcastro.space",
        "IG followers": "", "Posts UGC?": "", "Google rating": 4.9, "Google reviews": 168,
        "New (<6 mo)?": "Yes", "Website": "", "Phone": "+971 52 957 7969",
        "WhatsApp": "+971 52 957 7969", "Email": "", "Best channel": "WhatsApp",
        "Status": "1 - Enriched", "Next step": "WhatsApp opener (mobile) + 2 profiles",
        "Notes": "Portuguese bakery (pastel de nata) + specialty coffee. Recently opened (Curly Tales). NO website -> socials-first, strong trial-month fit. The Icon Casa 4.",
    },
    {
        "Venue name": "Loukie's Speciality Coffee", "Type": "Specialty coffee",
        "District": "JVC", "Locations": "1 venue", "IG handle": "",
        "IG followers": "", "Posts UGC?": "", "Google rating": 4.9, "Google reviews": 135,
        "New (<6 mo)?": "", "Website": "loukies.coffee", "Phone": "+971 58 584 5945",
        "WhatsApp": "+971 58 584 5945", "Email": "", "Best channel": "WhatsApp",
        "Status": "1 - Enriched", "Next step": "Verify IG handle, then WhatsApp opener",
        "Notes": "Sgal Community Park, district 15. IG handle not confirmed in search - verify on the venue.",
    },
    {
        "Venue name": "Village Hangout Cafe (JVC)", "Type": "Brunch / cafe",
        "District": "JVC", "Locations": "1 venue", "IG handle": "@villagehangoutcafe",
        "IG followers": 641, "Posts UGC?": "", "Google rating": 4.9, "Google reviews": 1481,
        "New (<6 mo)?": "", "Website": "villagehangoutcafe.com", "Phone": "+971 56 650 6045",
        "WhatsApp": "+971 56 650 6045", "Email": "villagehangoutjvc@gmail.com",
        "Decision maker": "Varsha", "Role": "Owner", "Best channel": "WhatsApp",
        "Status": "1 - Enriched", "Next step": "WhatsApp opener to Varsha + 2 profiles",
        "Notes": "All-day breakfast + specialty coffee. 1481 Google reviews but only ~641 IG followers = weak social flow -> ideal pitch. 178 IG posts. Owner: Varsha. Reef Residence.",
    },
    {
        "Venue name": "Roasters Specialty Coffee House Dubai Hills", "Type": "Specialty coffee",
        "District": "Dubai Hills", "Locations": "7+ venues", "IG handle": "@roasterscoffee_dxb",
        "IG followers": 69000, "Posts UGC?": "Yes", "Google rating": 4.8, "Google reviews": 886,
        "New (<6 mo)?": "Yes", "Website": "roasterscoffee.ae", "Phone": "+971 58 514 4267",
        "WhatsApp": "+971 58 514 4267", "Email": "", "Best channel": "WhatsApp",
        "Status": "1 - Enriched", "Next step": "Lower priority - large chain, already saturated with UGC",
        "Notes": "Big local chain (Guinness-awarded), 69K IG, Dubai Hills branch just opened. Already gets heavy UGC -> weak fit for trial pitch. Included for contrast.",
    },
    {
        "Venue name": "The Coffee Merchant", "Type": "Specialty coffee",
        "District": "Business Bay", "Locations": "1 venue", "IG handle": "@thecoffeemerchant.ae",
        "IG followers": 3075, "Posts UGC?": "", "Google rating": 4.9, "Google reviews": 397,
        "New (<6 mo)?": "", "Website": "thecoffeemerchant.ae", "Phone": "+971 56 741 6664",
        "WhatsApp": "+971 56 741 6664", "Email": "thecoffeemerchantae@gmail.com",
        "Best channel": "WhatsApp", "Status": "1 - Enriched",
        "Next step": "WhatsApp opener + 2 profiles (email as backup)",
        "Notes": "Own roastery. 3.0K IG. AG Tower, Marasi Dr. Email + mobile both public.",
    },
    {
        "Venue name": "Title Brew", "Type": "Specialty coffee",
        "District": "Business Bay", "Locations": "1 venue", "IG handle": "",
        "IG followers": "", "Posts UGC?": "", "Google rating": 4.7, "Google reviews": 160,
        "New (<6 mo)?": "", "Website": "", "Phone": "+971 4 447 4703",
        "WhatsApp": "", "Email": "", "Best channel": "Instagram DM", "Status": "1 - Enriched",
        "Next step": "Find IG handle, then IG DM opener",
        "Notes": "Indie / vinyl vibe, built for creatives. Bay Square 08, Marasi Dr. NO website -> socials-first. IG handle to confirm on venue.",
    },
    {
        "Venue name": "Estekana Cafe", "Type": "Specialty coffee",
        "District": "Business Bay", "Locations": "1 venue", "IG handle": "",
        "IG handle": "@estekanacafe.ae", "IG followers": 389, "Posts UGC?": "",
        "Google rating": 4.2, "Google reviews": 69,
        "New (<6 mo)?": "", "Website": "l.wl.co (linktree)", "Phone": "+971 4 423 2239",
        "WhatsApp": "", "Email": "", "Best channel": "Instagram DM", "Status": "1 - Enriched",
        "Next step": "IG DM opener + 2 profiles",
        "Notes": "Blue Bay Tower, Marasi Dr. Burj Khalifa view, Turkish coffee. Lowest social proof of the set (4.2 / 69). 389 IG followers, 135 posts. Link-tree, not a real site.",
    },
    {
        "Venue name": "Baristas' Corner Business Bay", "Type": "Specialty coffee",
        "District": "Business Bay", "Locations": "1 venue", "IG handle": "@baristascornerofficial",
        "IG followers": 3576, "Posts UGC?": "", "Google rating": 4.4, "Google reviews": 158,
        "New (<6 mo)?": "", "Website": "baristascorner.com", "Phone": "+971 56 446 4461",
        "WhatsApp": "+971 56 446 4461", "Email": "info@baristascorner.com", "Best channel": "WhatsApp", "Status": "1 - Enriched",
        "Next step": "WhatsApp opener (mobile from site) + 2 profiles",
        "Notes": "24h, Canal Central Hotel, Marasi Dr. 3.6K IG, 317 posts. Mobile +971 56 446 4461 from website (Google listing shows landline +971 4 873 2191). 'Brainchild of an Emirati entrepreneur' - name not public. Owner claimed the Google listing.",
    },
    {
        "Venue name": "Kimi - Speciality Coffee and Food", "Type": "Brunch / cafe",
        "District": "Business Bay", "Locations": "1 venue", "IG handle": "@kimicafedubai",
        "IG followers": "", "Posts UGC?": "", "Google rating": 4.8, "Google reviews": 176,
        "New (<6 mo)?": "", "Website": "kimicafe.com", "Phone": "+971 4 572 3647",
        "WhatsApp": "", "Email": "", "Best channel": "Instagram DM", "Status": "1 - Enriched",
        "Next step": "IG DM opener + 2 profiles",
        "Notes": "Design-forward cafe (award-noted interior). Bay Square, Building 09. Owner claimed the Google listing.",
    },
]

# ---- CSV ----
with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=HEADER)
    w.writeheader()
    for r in ROWS:
        w.writerow({k: r.get(k, "") for k in HEADER})
print("saved", OUT_CSV)

# ---- XLSX ----
wb = openpyxl.load_workbook(TEMPLATE)
ws = wb["Leads"]
INPUT_FONT = Font(name="Arial", color="0000FF", size=10)
FIRST = 4  # first data row (row 3 = header)

# header name -> column index
col_of = {}
for c in range(1, ws.max_column + 1):
    v = ws.cell(3, c).value
    if v:
        col_of[v] = c

for i, r in enumerate(ROWS):
    row = FIRST + i
    for name, cidx in col_of.items():
        if name in ("Score", "Tier"):
            continue  # keep formula
        cell = ws.cell(row, cidx)
        val = r.get(name, "")
        cell.value = val if val != "" else None
        cell.font = INPUT_FONT
    ws.cell(row, col_of["IG followers"]).number_format = "#,##0"
    ws.cell(row, col_of["Google reviews"]).number_format = "#,##0"
    ws.cell(row, col_of["Google rating"]).number_format = "0.0"

# ---- Deep enrichment sheet (6 layers) ----
from openpyxl.styles import PatternFill, Alignment, Border, Side
HDR_FILL = PatternFill("solid", fgColor="16324F")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
GUESS_FONT = Font(name="Arial", color="B45309", size=10)   # amber = hypothesis, verify
OK_FONT = Font(name="Arial", color="1A1A1A", size=10)
MANUAL_FONT = Font(name="Arial", italic=True, color="9A9A9A", size=10)
thin = Side(style="thin", color="D0D0D0")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)

deep_cols = [
    ("Venue", 26),
    ("L1 · IG (followers / posts)", 24),
    ("L1 · Creator-readiness", 26),
    ("L2 · Email (confirmed)", 26),
    ("L2 · Email hypothesis (verify!)", 26),
    ("L3 · Momentum / opened", 26),
    ("L4 · Delivery presence", 20),
    ("L5 · Website / tech", 22),
    ("L6 · Cross-rating", 20),
]

# g = confirmed/ok, ? = hypothesis(amber), m = manual/login-gated(grey)
DEEP = [
    ["Stella's Coffee House - JVC",
     ("@stellas_coffeehouse +TikTok; count login-gated","m"),
     ("Owner replies to every review; UGC cadence = manual","m"),
     ("bassel@stellasdxb.com (owner)","g"),
     ("firstname@stellasdxb.com pattern","?"),
     ("Dog-cafe hook; opened date = manual","m"),
     ("Talabat unclear (namesake boutique) — check","m"),
     ("stellasdxb.com (custom site)","g"),
     ("Google 4.9/300","g")],
    ["VITOR CASTRO - JVC",
     ("@vitorcastro.space; count login-gated","m"),
     ("Portuguese pastel de nata = highly shootable","g"),
     ("— (no domain)","m"),
     ("IG DM only","?"),
     ("NEW: opened ~Jul 2026 (Curly Tales) — HOT","g"),
     ("Not confirmed on aggregators","m"),
     ("No website → socials-first","g"),
     ("Google 4.9/168","g")],
    ["Loukie's Speciality Coffee",
     ("Handle unconfirmed — verify on site","m"),
     ("Manual (need handle first)","m"),
     ("— none confirmed","m"),
     ("info@loukies.coffee","?"),
     ("Site = 'coming soon' splash → rebrand/new?","g"),
     ("Manual","m"),
     ("loukies.coffee = placeholder page","g"),
     ("Google 4.9/135","g")],
    ["Village Hangout Cafe (JVC)",
     ("@villagehangoutcafe · 641 / 178","g"),
     ("1481 reviews vs 641 IG = weak flow → IDEAL","g"),
     ("villagehangoutjvc@gmail.com (owner Varsha)","g"),
     ("varsha@villagehangoutcafe.com","?"),
     ("Established, high review momentum","g"),
     ("On aggregators (contactless delivery)","g"),
     ("villagehangoutcafe.com (real site + About)","g"),
     ("Google 4.9/1481; on Tripadvisor","g")],
    ["Roasters Specialty Coffee House Dubai Hills",
     ("@roasterscoffee_dxb · 69K (chain)","g"),
     ("Already saturated with UGC → weak fit","g"),
     ("— (co-founder/CEO: Konstantin Harbuz)","m"),
     ("info@roasterscoffee.ae","?"),
     ("Dubai Hills branch just opened; big chain","g"),
     ("Yes (chain, all aggregators)","g"),
     ("roasterscoffee.ae + linktree","g"),
     ("Google 4.8/886","g")],
    ["The Coffee Merchant",
     ("@thecoffeemerchant.ae · 3.0K","g"),
     ("3K audience, UGC cadence = manual","m"),
     ("thecoffeemerchantae@gmail.com","g"),
     ("hello@thecoffeemerchant.ae","?"),
     ("Own roastery; sells beans online","g"),
     ("Own delivery (free UAE) via site","g"),
     ("Shopify e-commerce → cares about reach","g"),
     ("Google 4.9/397","g")],
    ["Title Brew",
     ("Handle unconfirmed — verify on site","m"),
     ("Indie/vinyl 'built for creatives' — good angle","g"),
     ("— (no domain)","m"),
     ("IG DM / phone","?"),
     ("NEW-ish: Zomato 0 ratings; ~AED115/2","g"),
     ("Minimal (takeaway; no delivery ratings)","g"),
     ("No website → socials-first","g"),
     ("Google 4.7/160; Zomato 0","g")],
    ["Estekana Cafe",
     ("@estekanacafe.ae · 389 / 135","g"),
     ("Burj Khalifa view = shootable; small audience","g"),
     ("— none confirmed","m"),
     ("via linktree l.wl.co","?"),
     ("Turkish coffee, view spot","g"),
     ("Drive-thru + Drivu/EatEasy listed","g"),
     ("Linktree, no real site","g"),
     ("Google 4.2/69 (lowest of set)","g")],
    ["Baristas' Corner Business Bay",
     ("@baristascornerofficial · 3.6K / 317","g"),
     ("3.6K + FB/Twitter/Snapchat; cadence manual","m"),
     ("info@baristascorner.com (decoded)","g"),
     ("mobile +971 56 446 4461 (WhatsApp)","g"),
     ("24h; 'Emirati entrepreneur' (name not public)","g"),
     ("Yes (magicpin / Zomato)","g"),
     ("baristascorner.com (real site)","g"),
     ("Google 4.4/158","g")],
    ["Kimi - Speciality Coffee and Food",
     ("@kimicafedubai; count login-gated","m"),
     ("Design-forward interior = shootable","g"),
     ("— none confirmed (site SSL error)","m"),
     ("hello@kimicafe.com","?"),
     ("Award-noted interior design","g"),
     ("Listed (Business Bay aggregators)","g"),
     ("kimicafe.com","g"),
     ("Google 4.8/176; 2GIS 3.5 (divergence!)","g")],
]

deep = wb.create_sheet("Deep enrichment")
deep.sheet_view.showGridLines = False
deep.cell(1, 1, "Deep enrichment — 6 layers (Aug 2026)").font = Font(name="Arial", bold=True, size=14, color="16324F")
deep.cell(2, 1, "Green = confirmed/public · Amber = hypothesis, VERIFY before use (PDPL/spam risk) · Grey = login-gated or needs a call/visit").font = Font(name="Arial", italic=True, size=9, color="6B6B6B")
for c, (name, w) in enumerate(deep_cols, start=1):
    cell = deep.cell(4, c, name)
    cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BD
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    deep.column_dimensions[deep.cell(4, c).column_letter].width = w
deep.row_dimensions[4].height = 30
font_of = {"g": OK_FONT, "?": GUESS_FONT, "m": MANUAL_FONT}
for ri, row in enumerate(DEEP, start=5):
    deep.cell(ri, 1, row[0]).font = Font(name="Arial", bold=True, size=10)
    deep.cell(ri, 1).border = BD
    deep.cell(ri, 1).alignment = Alignment(vertical="top", wrap_text=True)
    for ci, (text, kind) in enumerate(row[1:], start=2):
        cell = deep.cell(ri, ci, text)
        cell.font = font_of[kind]; cell.border = BD
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    deep.row_dimensions[ri].height = 42
deep.freeze_panes = "B5"

wb.save(OUT_XLSX)
print("saved", OUT_XLSX, f"({len(ROWS)} rows) + Deep enrichment sheet")
