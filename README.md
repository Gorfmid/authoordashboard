# Author Dashboard v1.1

A Google Sheets + Apps Script publishing dashboard with one editable sheet and automatic Amazon sales-rank tracking.

## Sheet layout

1. Manual Entry — the only sheet you edit
2. Dashboard — automatic and protected
3. Catalog Summary — automatic and protected
4. Sales History — automatic and protected
5. Rank History — automatic and protected
6. Marketing History — automatic and protected

## Before you start

Install Node.js, then install clasp:

```bash
npm install -g @google/clasp
```

On Windows PowerShell, if `clasp` is blocked by execution policy, use:

```powershell
clasp.cmd login
clasp.cmd push
```

If OAuth/token requests fail with a certificate error, set:

```powershell
$env:NODE_OPTIONS = "--use-system-ca"
```

## Connect this folder to your existing Google Sheet

1. Open the Google Sheet.
2. Open **Extensions → Apps Script**.
3. In Apps Script, open **Project Settings**.
4. Copy the **Script ID**.
5. In this folder, copy `.clasp.json.example` to `.clasp.json`.
6. Replace the placeholder with your Script ID.

Enable the Apps Script API:

https://script.google.com/home/usersettings

Then run:

```bash
clasp login
clasp push
```

When clasp asks whether to overwrite remote files, approve it.

Reload the Google Sheet after pushing so the **Author Dashboard** menu refreshes.

## Initialize the workbook

1. Reload the Google Sheet.
2. Open **Author Dashboard → Initialize — ERASE AND REBUILD**.
3. Approve permissions and confirm the reset.

You can also run `initializeDashboard` directly from Apps Script.

## Where data goes

All manual data goes into **Manual Entry**.

- ASIN or ISBN: column M (`Identifier / ASIN / ISBN`)
- For Amazon rank tracking: Store = `Amazon`, Identifier Type = `ASIN`, Listing Status = `Live` or `In Review`
- Lifetime units: column Q
- Lifetime KU pages: column R
- Lifetime royalties: column S
- Current overall rank / rating / reviews: filled automatically by Amazon updates when available
- Process Status: shows the latest Amazon fetch result for that listing

Book ID and Listing ID are generated automatically. Do not overwrite them.

## Automatic Amazon rank updates

Menu:

- **Author Dashboard → Update Amazon Rankings Now**
- **Author Dashboard → Install Weekly Rank Update**
- **Author Dashboard → Remove Weekly Rank Update**

What the updater does:

1. Finds active Amazon ASIN listings on Manual Entry
2. Fetches each Amazon product page
3. Extracts overall rank, category ranks, rating, and review count when available
4. Writes current values back to Manual Entry
5. Appends permanent Rank History rows
6. Refreshes Catalog Summary and Dashboard

Rank History stores values (not formulas), including Overall and Category rows.

Amazon page scraping can fail. Amazon may return CAPTCHA/robot-check pages, HTTP errors, or change page markup. Failures are written to Process Status and Apps Script logs; one failed listing does not stop the rest.

## KDP sales upload (units, KU pages, royalties)

Apps Script cannot download KDP files while you are logged into Amazon, but upload is one click:

1. Download a KDP **Dashboard** `.xlsx` using **All time / full life** (not a short custom range)
2. In your Sheet: **Author Dashboard → Upload KDP Sales Report**
3. Choose the `.xlsx` file

That single action:

- Parses the workbook in a dialog (no clutter sheets added)
- Matches ASINs/ISBNs to Manual Entry Amazon rows
- Fills **Lifetime** Units, KU Pages (KENP), and Royalties (USD) only when the report is not detected as partial
- Splits royalties into **eBook / Print / KENP** (USD) and stores a **Royalty Periods** row for that reporting window
- Refuses to lower existing lifetime totals; blocks partial date-range overwrites
- Records a Sales History snapshot using the **report end date** for Week Ending when dates are present
- Refreshes Catalog Summary, Dashboard, and Reconciliation

Enter **KENPC** (Kindle Edition Normalized Page Count) on Manual Entry for the Kindle listing so Dashboard can show Estimated Full KU Read Royalty and KU Equivalent Reads. The **Estimated KENP Royalty Rate** is period-matched (KENP $ ÷ KENP pages from the same upload) and is not the finalized monthly Global Fund rate.

Ranks are not changed. Use **Open KDP Reports Page** anytime for the download site + QR code.

## Dashboard vs Visual Dashboard

- **Dashboard** — portfolio numbers, **KU Estimates**, Catalog Performance (incl. KENP), category ranks, Meta sync health  
- **Visual Dashboard** — charts only (rank score trend, orders, KENP), tab immediately after Dashboard  
- **Royalty Periods** — historical estimated KU rate / royalty mix per KDP report window (foundation for future daily snapshots)

## Events

Editable **Events** sheet (and **Author Dashboard → Add Event**) for a **timeline**: book releases, Meta campaign start/end, price changes, free days, newsletters, big social posts. Optional `SOL-###` Book ID. Later these can appear as markers on charts. It is **not** a sales or ad-spend ledger.

## Meta Daily vs Marketing History

| Sheet | Purpose |
|--------|---------|
| **Meta Daily** | Daily API metrics (spend, impressions, clicks, link clicks) |
| **Marketing History** | Manual one-off notes from Manual Entry (newsletter, promo, etc.) |

Meta does **not** auto-write Marketing History (different grain; would flood it with daily rows). Use **Events** for “campaign started”, Meta Daily for spend/clicks.

## Meta ads (local pull → CSV upload)

Secrets stay in a local `.env` (see `.env.example`). Never put tokens in the Sheet.

1. `pip install -r meta/requirements.txt`
2. Copy `.env.example` → `.env` and set `META_ACCESS_TOKEN` + `META_AD_ACCOUNT_ID`
3. `python meta/sync_meta_insights.py`
4. **Author Dashboard → Upload Meta Insights CSV**

Details: `meta/README.md`. Metrics are ad spend/clicks only — not Amazon conversions.

## Stable internal IDs

- **Book ID** = `SOL-001`, `SOL-002`, … (permanent). Do not use title as the relationship key.
- **Listing ID** = `SOL-001-AMAZ-KIND-xxxx` (book + store + format + suffix).
- ASIN/ISBN, marketplace, and format stay in their own columns.
- One-time: **Author Dashboard → Migrate Stable Book IDs (SOL-001…)**.

## Sales metrics (do not confuse these)

| Concept | Where | Meaning |
|--------|--------|---------|
| Lifetime cumulative | Manual Entry + Sales History “Lifetime …” columns | All-time total at that moment |
| Units/KENP/Royalties since previous snapshot | Sales History period columns | Change since the prior snapshot for that listing (first snapshot = 0) |
| Reporting period from KDP | Report date range on import | Used for Week Ending / Last Data Date — **not** inferred from import day alone |

## Sales History shading and yearly reports

- **Sales History** rows are lightly color-shaded by Book ID (adjacent books never share a color).
- Auto-built year sheets named **Sales 2026**, **Sales 2027**, … show period units, lifetime units, period KENP, and lifetime KENP with clear labels.
- **Year over Year** summarizes **period** units/royalties by book, plus a **KENP by year** line chart (2026 vs 2027 as separate lines when both exist).
- **Reconciliation** flags mismatches (does not auto-correct).

After upgrading to Phase 0, run once: **Author Dashboard → Migrate Sales History (Phase 0)**.

These sheets regenerate on **Refresh Everything**, sales snapshots, KDP upload, and the weekly sales trigger.

## Weekly sales snapshots

Update lifetime totals in Manual Entry (manually or via KDP import), then choose:

**Author Dashboard → Record Current Snapshot**

Or install:

**Author Dashboard → Install Weekly Saturday Night Update**

That runs at **Saturday night midnight Mountain time** (Sunday 12:00 AM `America/Boise`). Week ending is Saturday.

Also install **Install Weekly Saturday Night Rank Update** for the Amazon rank fetch on the same schedule.

Period changes are calculated by comparing current lifetime totals to the previous sales snapshot (not by guessing a calendar week from the import date).

## Testing Amazon fetch

In Apps Script, run:

- `testAmazonRankFetch()` — fetches the first valid Amazon ASIN from Manual Entry and shows a summary alert
- `testRankHistoryAppend()` — validates duplicate prevention and removes its temporary test row

## Inspecting logs

1. Open **Extensions → Apps Script**
2. Open **Executions** (clock icon) or the execution log for a run
3. Look for `Amazon fetch ...` and `Amazon rank update summary` messages

## Notes

- Manual Entry remains the only editable sheet
- Dashboard, Catalog Summary, Sales History, Rank History, and Marketing History stay protected
- Rank History is migrated in place if older columns are missing (`Category`, `Source URL`, `Fetch Status`)
- Lower Amazon rank numbers are better; Dashboard labels and the rank chart make that explicit
