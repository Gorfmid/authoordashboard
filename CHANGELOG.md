# Author Dashboard — Change Log

## Backup (required before Phase 0)

| Item | Value |
|------|--------|
| Local Apps Script backup | `C:\Users\gorfm\Documents\projects\AuthorDashboard-backup-20260726-080638` |
| Google Sheet backup | **Required manually:** File → Make a copy → name `Author Dashboard BACKUP 2026-07-26` |
| Rollback | Restore from Sheet copy; restore `.gs`/`.html` from the local backup folder; `clasp.cmd push` |

Do not treat production as safe until the Google Sheet copy exists.

---

## Phase 0 — Sales history model (2026-07-26)

**Goal:** Fix confusion between lifetime totals, period sales, and snapshot deltas. No Meta integration.

### Why

Sales year charts showed “weekly” numbers that were snapshot-to-snapshot deltas of lifetime totals. First snapshots incorrectly treated full lifetime as a “weekly” sale. KDP imports could overwrite lifetime with a partial date-range report, and Week Ending was stamped from import day, not the report period.

### Files changed

| File | Change |
|------|--------|
| `Config.gs` | Version/schema 1.4.0; renamed Sales History period columns; Reconciliation sheet name |
| `History.gs` | Period change = 0 on first snapshot; schema migrate; recompute from lifetime |
| `SalesReport.gs` | Sales YYYY: period + lifetime units + KENP; labels; KENP-by-year overlay on YoY |
| `KdpImport.gs` | Report date range; partial-report block; no lifetime decrease; snapshot dates from report |
| `KdpUpload.html` | Date normalization; lifetime-report guidance |
| `Reconciliation.gs` | **New** — checks + System version block |
| `Tests.gs` | Phase 0 test suite (pure + sheet-safe) |
| `Main.gs` | Menu: migrate, reconcile, Phase 0 tests; refresh hooks |
| `Validation.gs` | Clearer Manual Entry notes for lifetime fields |
| `Protection.gs` | (via `getAutomaticSheetNames_`) Reconciliation protected |
| `README.md` | Phase 0 semantics documented |
| `CHANGELOG.md` | This file |

### Sheets changed

| Sheet | Change |
|-------|--------|
| Sales History | Header rename (period columns); optional recompute of period columns from lifetime series |
| Sales YYYY | Layout: legend + period units + lifetime units + period KENP + lifetime KENP + charts |
| Year over Year | Clearer labels; KENP-by-year line overlay (2026 vs 2027 when present) |
| Reconciliation | **New** — discrepancy flags + workbook/schema/script version |

### Formulas / triggers / named ranges

- No named-range changes.
- Existing weekly triggers unchanged (still call snapshot helpers).
- Dashboard still rebuilt from Catalog / Manual Entry (not from import sheets).

### Rollback path

1. Open Sheet backup copy, or restore Sales History from that copy.
2. Copy `.gs`/`.html` from `AuthorDashboard-backup-20260726-080638`.
3. `clasp.cmd push` and reload the Sheet.
4. Run **Refresh Everything** if needed.

### Phase 2d — KENP / KU royalty metrics (2026-07-26)

| Item | Change |
|------|--------|
| Manual Entry | Lifetime eBook / Print / KENP Royalties (USD) + **KENPC** (manual); total = sum of three |
| Import | Persist eBook vs Print vs KENP buckets; total no longer drops KENP when format ≥ Combined |
| Royalty Periods | New sheet — period-matched KENP pages + KENP $, estimated rate, mix %, reconcile |
| Dashboard | **KU Estimates** block: Estimated KENP Royalty Rate, $/1k, full-read, equivalent reads, mix % |
| Labels | Rate labeled estimated (not finalized Global Fund); note on calculation / month finalization |
| Reconciliation | eBook + Print + KENP ≈ Total (±$0.01) on Manual Entry + latest period |

### Hotfix — USD royalties + Sales week from orders (2026-07-26)

| Item | Change |
|------|--------|
| Royalties | Catalog shows **USD only**; matches KDP Summary Royalty (USD) = **$110.23**. CAD/GBP excluded. KENP pages ≠ $ until Amazon posts KU royalty lines |
| Sales History week | Snapshot / Week Ending anchored on **max order/royalty date**, not KENP. Fixes bogus `7/26 → week ending 8/1` when Sunday has KENP but no orders |
| Upsert | KDP import upserts by Week Ending + Listing; refresh consolidates duplicate week rows and recomputes period deltas |
| Labels | Royalties (USD) on Dashboard / Manual Entry header |

### Phase 2c — Royalties + auto Events/Meta marketing (2026-07-26)

| Item | Change |
|------|--------|
| KDP royalties | Merge complementary sources (Combined units + format/KENP KU), not `max` alone |
| Dashboard title | Merged through column K |
| Events | Auto book-release + Meta campaign start on Refresh (manual rows kept) |
| Marketing History | One upserted row per Meta campaign (`AUTO_META\|…`) on Refresh / Meta upload |

### Phase 2b — Visual Dashboard + UX fixes (2026-07-26)

| Item | Change |
|------|--------|
| Reconciliation | Force-hidden after every refresh/reorder |
| Visual Dashboard | New sheet (tab right of Dashboard) with all charts |
| Rank chart | Readable **rank score** (higher=better) + real-rank table |
| Catalog Performance | Starts at column F; adds **KENP Read** |
| KDP royalties | `max(Combined, format/KENP sheets)` so KU earnings are not dropped |
| Meta vs Marketing | Documented as separate (Meta Daily ≠ Marketing History) |

### Phase 2 — Rank chart fix + Meta insights (2026-07-26)

| Item | Change |
|------|--------|
| Rank chart | Sheets cannot reverse vAxis — chart now plots **−rank** so up = better |
| Meta | Local `meta/sync_meta_insights.py` + `.env`; **Upload Meta Insights CSV** → Meta Daily upsert |
| Dashboard | Last Meta Sync / Meta Sync Status (stale warning >3 days) |
| Guardrails | IDs as text; clicks vs link clicks; no Amazon “conversion” labeling |
| Files | `Dashboard.gs`, `MetaImport.gs`, `MetaUpload.html`, `meta/*`, `Config.gs`, `Main.gs`, `.gitignore` |

### Phase 1b — Events + rank chart / label fixes (2026-07-26)

| Item | Change |
|------|--------|
| Dashboard / Catalog | `Current Best Rank` → **Current Rank** |
| Rank chart | Y-axis inverted (`vAxis.direction: -1`) so lower/better ranks plot upward |
| MANUAL cleanup | Only deletes Manual Snapshot when Amazon Overall exists for same date+listing |
| 7/25 paperback gap | `repairMissingJuly25OverallRanks_` restores known Overall ranks by ASIN |
| Events sheet | New editable **Events** table + **Add Event** menu (`EVT-001…`) |
| Files | `Dashboard.gs`, `Config.gs`, `Init.gs`, `History.gs`, `Events.gs` (new), `Main.gs`, `SalesReport.gs` |

### Phase 1a — Stable SOL book IDs (2026-07-26)

| Item | Change |
|------|--------|
| Book ID | Permanent `SOL-001`, `SOL-002`, … (prefix configurable) |
| Listing ID | `{SOL-xxx}-{STORE}-{FORM}-{suffix}`; suffix preserved on migrate |
| Migration | Menu: **Migrate Stable Book IDs (SOL-001…)** rewrites Manual Entry + Sales/Rank/Marketing history |
| Preferred map | SOL-001 Kestrel Veil, SOL-002 Helion Accord, SOL-003 Book 3 |
| New books | `assignInternalIds_` issues next SOL number (no more `BK-` uuid) |
| Files | `StableIds.gs` (new), `Data.gs`, `Helpers.gs`, `Config.gs`, `Main.gs`, `Validation.gs`, `SalesReport.gs`, `Tests.gs` |
| Rollback | Restore Sheet backup; restore pre-1.5.0 scripts; listing/book ids revert with data |

### Follow-up — dashboard charts + sales/rank fixes (2026-07-26)

| Item | Change |
|------|--------|
| Reconciliation | Kept for diagnostics; **hidden** after rebuild |
| Dashboard | Orders (stacked) + KENP-by-year (lines) charts; Sales/YoY keep tables only |
| Week totals | Period Orders/KENP from **book-level lifetime week-over-week** (deduped per listing/week) |
| Rank MANUAL rows | Stop writing from sales snapshot; cleanup deletes `Manual Snapshot` / `MANUAL` rows |
| Files | `Dashboard.gs`, `SalesReport.gs`, `History.gs`, `Triggers.gs`, `Reconciliation.gs`, `Main.gs` |

### Hotfix — migrate freeze/merge (2026-07-26)

| Item | Change |
|------|--------|
| Error | `You can't merge frozen and non-frozen columns` on Migrate Sales History |
| Cause | Year over Year (and Sales YYYY banners) merged across columns while column A was frozen |
| Fix | `prepareSheetForRebuild_` unfreezes/unmerges; banners use `setBannerRow_` (no merge); YoY freezes header row only |
| Files | `Helpers.gs`, `SalesReport.gs`, `Reconciliation.gs` |

### Remaining risks (Phase 0)

- Historical rows already written with inflated first-period deltas need **Migrate Sales History (Phase 0)** once.
- Partial vs lifetime detection is heuristic (date span + totals vs current); edge-case reports may need a lifetime “All time” export.
- Book title still used for year-sheet pivots (stable SOL- IDs are a later phase).
- KENP year chart uses period changes by week-of-year; weeks without snapshots show 0.
