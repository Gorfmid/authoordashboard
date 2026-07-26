# Meta Marketing API sync

Pull **daily** ad insights locally, then upload the CSV into the Author Dashboard.

Secrets stay in a local `.env` — never in the Google Sheet or git.

## Setup

```powershell
cd C:\Users\gorfm\Documents\projects\AuthorDashboard
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r meta\requirements.txt
copy meta\.env.example .env
```

Edit `.env`:

- `META_ACCESS_TOKEN` — long-lived Marketing API token  
- `META_AD_ACCOUNT_ID` — e.g. `act_1234567890`  
- Optional `META_DATE_START` / `META_DATE_END` (defaults: last 30 days)  
- Optional `META_CAMPAIGN_ID` to limit to one campaign  

## Run

```powershell
python meta\sync_meta_insights.py
```

Writes `meta/out/meta_daily_insights.csv` with:

- `time_increment=1` (daily rows)
- Pagination + retries / rate-limit handling
- Token-expiry errors called out clearly
- IDs forced as text
- **Clicks (all)** vs **Link Clicks** / landing page views from `actions`

## Import into the Sheet

1. Reload the Google Sheet  
2. **Author Dashboard → Upload Meta Insights CSV**  
3. Choose the CSV  

Upserts into **Meta Daily** (keyed by date + campaign + ad set + ad).  

## Attribution warning

Meta clicks are **not** Amazon orders. Dashboard language uses sync health only. Prefer Events (“Meta campaign started”) and phrases like *sales during campaign* / *correlated sales* — never “Meta conversions” unless you add real conversion tracking later.
