"""EDINET API v2 client for fetching annual securities reports."""

import time
from datetime import date, timedelta
from typing import Optional

import requests


ANNUAL_REPORT_CODE = "120"


class EdinetClient:
    BASE_URL = "https://disclosure.edinet-api.go.jp/api/v2"

    def __init__(self, api_key: str = ""):
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "FinancialAnalysisTool/1.0"

    # ------------------------------------------------------------------ #
    # Internal helpers                                                     #
    # ------------------------------------------------------------------ #

    def _params(self, extra: dict = None) -> dict:
        p = dict(extra or {})
        if self.api_key:
            p["Subscription-Key"] = self.api_key
        return p

    def _get(self, path: str, params: dict = None) -> Optional[requests.Response]:
        url = f"{self.BASE_URL}/{path}"
        for attempt in range(4):
            try:
                r = self.session.get(url, params=self._params(params), timeout=30)
                r.raise_for_status()
                return r
            except requests.RequestException as e:
                if attempt < 3:
                    time.sleep(2 ** attempt)
                else:
                    print(f"    API error ({path}): {e}")
        return None

    def _docs_for_date(self, d: date) -> list:
        r = self._get("documents.json", {"date": d.strftime("%Y-%m-%d"), "type": 2})
        if r is None:
            return []
        try:
            return r.json().get("results") or []
        except Exception:
            return []

    # ------------------------------------------------------------------ #
    # Public API                                                           #
    # ------------------------------------------------------------------ #

    def search_company(self, name: str) -> Optional[dict]:
        """Search for a listed company by name and return its EDINET metadata.

        Strategy: check weekly intervals going back 7 years, prioritising the
        June-August window where most March-FY companies file their annual
        reports (~70 % of Tokyo-listed companies).
        """
        today = date.today()

        # Build a prioritised date list: check every week, but pre-prioritise
        # the common filing months so we find the company faster.
        def date_range_weekly(start: date, end: date):
            d = start
            while d >= end:
                yield d
                d -= timedelta(days=7)

        priority_dates = []
        other_dates = []
        cutoff = today - timedelta(days=365 * 7)

        d = today
        while d >= cutoff:
            if d.month in (6, 7, 8, 3, 4):
                priority_dates.append(d)
            else:
                other_dates.append(d)
            d -= timedelta(days=7)

        search_dates = priority_dates + other_dates

        print(f"  企業を検索中", end="", flush=True)
        for i, d in enumerate(search_dates):
            if i % 15 == 0:
                print(".", end="", flush=True)

            docs = self._docs_for_date(d)
            time.sleep(0.35)

            for doc in docs:
                if doc.get("docTypeCode") != ANNUAL_REPORT_CODE:
                    continue
                filer = doc.get("filerName", "")
                if name in filer or filer in name:
                    print(" 見つかりました!")
                    return doc

        print(" 見つかりませんでした")
        return None

    def get_annual_reports(self, edinet_code: str, years: int = 5) -> list:
        """Return up to *years* annual-report metadata dicts for the company."""
        today = date.today()
        found: dict[str, dict] = {}  # period_end -> doc

        d = today
        cutoff = today - timedelta(days=365 * (years + 2))

        print(f"  有価証券報告書を検索中", end="", flush=True)
        dot_counter = 0
        while d >= cutoff and len(found) < years:
            if dot_counter % 15 == 0:
                print(".", end="", flush=True)
            dot_counter += 1

            docs = self._docs_for_date(d)
            time.sleep(0.35)

            for doc in docs:
                if (
                    doc.get("edinetCode") == edinet_code
                    and doc.get("docTypeCode") == ANNUAL_REPORT_CODE
                ):
                    period = doc.get("periodEnd", "")
                    if period and period not in found:
                        found[period] = doc

            d -= timedelta(days=7)

        print(f" {len(found)}件見つかりました")
        return sorted(found.values(), key=lambda x: x.get("periodEnd", ""), reverse=True)[:years]

    def download_xbrl(self, doc_id: str) -> Optional[bytes]:
        """Download the XBRL zip (type=5) for a document."""
        r = self._get(f"documents/{doc_id}", {"type": 5})
        if r is not None:
            return r.content
        return None
