#!/usr/bin/env python3
"""
Pull Meta Marketing API daily insights (time_increment=1) and write a CSV
for Author Dashboard → Upload Meta Insights CSV.

Secrets: load from project-root .env (never the Google Sheet).

Usage (from AuthorDashboard root):
  pip install -r meta/requirements.txt
  copy meta\\.env.example .env   # then fill tokens
  python meta/sync_meta_insights.py
"""

from __future__ import annotations

import csv
import json
import os
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]
GRAPH = "https://graph.facebook.com"


class MetaApiError(Exception):
    def __init__(self, message: str, status: Optional[int] = None, payload: Any = None):
        super().__init__(message)
        self.status = status
        self.payload = payload


def env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def as_text_id(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float):
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0") and s.replace(".", "", 1).isdigit():
        return s[:-2]
    return s


def action_value(actions: Any, action_type: str) -> float:
    if not isinstance(actions, list):
        return 0.0
    for item in actions:
        if isinstance(item, dict) and item.get("action_type") == action_type:
            try:
                return float(item.get("value") or 0)
            except (TypeError, ValueError):
                return 0.0
    return 0.0


def default_dates() -> tuple[str, str]:
    end = date.today()
    start = end - timedelta(days=30)
    return start.isoformat(), end.isoformat()


def request_with_retry(session: requests.Session, url: str, params: Dict[str, Any], max_attempts: int = 6) -> Dict[str, Any]:
    delay = 2.0
    last_err: Optional[Exception] = None
    for attempt in range(1, max_attempts + 1):
        try:
            resp = session.get(url, params=params, timeout=60)
            data = resp.json() if resp.content else {}
            if resp.status_code == 200 and "error" not in data:
                return data

            err = (data or {}).get("error") or {}
            code = err.get("code")
            msg = err.get("message") or resp.text or f"HTTP {resp.status_code}"

            # Token expired / invalid
            if code in (190, 102) or resp.status_code == 401:
                raise MetaApiError(
                    f"Meta token error ({code}): {msg}. Refresh META_ACCESS_TOKEN in .env.",
                    status=resp.status_code,
                    payload=data,
                )

            # Rate limit / transient
            if resp.status_code in (429, 500, 502, 503, 504) or code in (4, 17, 32, 613):
                last_err = MetaApiError(msg, status=resp.status_code, payload=data)
                time.sleep(delay)
                delay = min(delay * 2, 60)
                continue

            raise MetaApiError(msg, status=resp.status_code, payload=data)
        except requests.RequestException as exc:
            last_err = exc
            time.sleep(delay)
            delay = min(delay * 2, 60)
    raise MetaApiError(f"Meta API failed after retries: {last_err}")


def paginate_insights(session: requests.Session, url: str, params: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    next_url: Optional[str] = url
    next_params: Optional[Dict[str, Any]] = dict(params)

    while next_url:
        data = request_with_retry(session, next_url, next_params or {})
        chunk = data.get("data") or []
        if isinstance(chunk, list):
            rows.extend(chunk)
        paging = data.get("paging") or {}
        next_url = paging.get("next")
        next_params = None  # next URL already includes query string
        if next_url:
            time.sleep(0.25)
    return rows


def fetch_insights(
    token: str,
    ad_account_id: str,
    api_version: str,
    date_start: str,
    date_end: str,
    campaign_id: str = "",
) -> List[Dict[str, Any]]:
    account = ad_account_id if ad_account_id.startswith("act_") else f"act_{ad_account_id}"
    fields = ",".join(
        [
            "campaign_id",
            "campaign_name",
            "adset_id",
            "adset_name",
            "ad_id",
            "ad_name",
            "spend",
            "impressions",
            "clicks",
            "reach",
            "ctr",
            "cpc",
            "cpm",
            "actions",
            "date_start",
            "date_stop",
        ]
    )
    params: Dict[str, Any] = {
        "access_token": token,
        "level": "ad",
        "time_increment": 1,
        "limit": 100,
        "fields": fields,
        "time_range": json.dumps({"since": date_start, "until": date_end}),
    }
    if campaign_id:
        params["filtering"] = json.dumps(
            [{"field": "campaign.id", "operator": "EQUAL", "value": as_text_id(campaign_id)}]
        )

    url = f"{GRAPH}/{api_version}/{account}/insights"
    session = requests.Session()
    return paginate_insights(session, url, params)


def to_csv_rows(insights: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    for row in insights:
        actions = row.get("actions") or []
        link_clicks = action_value(actions, "link_click")
        # Some accounts report outbound_click / landing page views under different action types
        if link_clicks <= 0:
            link_clicks = action_value(actions, "outbound_click")
        lpv = action_value(actions, "landing_page_view")
        out.append(
            {
                "Date": row.get("date_start") or "",
                "Campaign ID": as_text_id(row.get("campaign_id")),
                "Campaign Name": str(row.get("campaign_name") or ""),
                "Ad Set ID": as_text_id(row.get("adset_id")),
                "Ad Set Name": str(row.get("adset_name") or ""),
                "Ad ID": as_text_id(row.get("ad_id")),
                "Ad Name": str(row.get("ad_name") or ""),
                "Spend": str(row.get("spend") or "0"),
                "Impressions": str(row.get("impressions") or "0"),
                "Clicks (all)": str(row.get("clicks") or "0"),
                "Link Clicks": str(int(link_clicks) if link_clicks == int(link_clicks) else link_clicks),
                "Landing Page Views": str(int(lpv) if lpv == int(lpv) else lpv),
                "Reach": str(row.get("reach") or "0"),
                "CTR": str(row.get("ctr") or ""),
                "CPC": str(row.get("cpc") or ""),
                "CPM": str(row.get("cpm") or ""),
                "Action Types (JSON)": json.dumps(actions, separators=(",", ":")),
                "Book ID": "",
            }
        )
    return out


def write_csv(path: Path, rows: List[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "Date",
        "Campaign ID",
        "Campaign Name",
        "Ad Set ID",
        "Ad Set Name",
        "Ad ID",
        "Ad Name",
        "Spend",
        "Impressions",
        "Clicks (all)",
        "Link Clicks",
        "Landing Page Views",
        "Reach",
        "CTR",
        "CPC",
        "CPM",
        "Action Types (JSON)",
        "Book ID",
    ]
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            # Force IDs as text for Sheets (prefix tab)
            for id_col in ("Campaign ID", "Ad Set ID", "Ad ID"):
                if row.get(id_col):
                    row[id_col] = "'" + as_text_id(row[id_col])
            writer.writerow(row)


def main() -> int:
    load_dotenv(ROOT / ".env")
    token = env("META_ACCESS_TOKEN")
    account = env("META_AD_ACCOUNT_ID")
    if not token or not account:
        print("Missing META_ACCESS_TOKEN or META_AD_ACCOUNT_ID in .env", file=sys.stderr)
        return 2

    api_version = env("META_API_VERSION", "v21.0")
    start = env("META_DATE_START")
    end = env("META_DATE_END")
    if not start or not end:
        start, end = default_dates()
    campaign_id = env("META_CAMPAIGN_ID")
    out_rel = env("META_OUT_CSV", "meta/out/meta_daily_insights.csv")
    out_path = (ROOT / out_rel).resolve()

    print(f"Fetching Meta insights {start} → {end} for {account} (level=ad, time_increment=1)…")
    try:
        insights = fetch_insights(token, account, api_version, start, end, campaign_id)
    except MetaApiError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    rows = to_csv_rows(insights)
    write_csv(out_path, rows)
    spend = sum(float(r.get("Spend") or 0) for r in rows)
    print(f"Wrote {len(rows)} daily ad rows to {out_path}")
    print(f"Spend sum (daily rows): ${spend:,.2f}")
    print("Next: Author Dashboard → Upload Meta Insights CSV")
    print("Reminder: clicks/spend are NOT Amazon conversions.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
