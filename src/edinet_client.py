"""EDINET API v2 client for fetching annual securities reports."""

import calendar
import time
from datetime import date, timedelta
from typing import Optional

import requests


ANNUAL_REPORT_CODE = "120"

# Filing month order by likelihood (March FY = June/July most common in Japan)
_MONTH_PRIORITY = [6, 7, 8, 12, 1, 9, 10, 11, 2, 3, 4, 5]


class EdinetApiError(Exception):
    pass


class EdinetClient:
    BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"

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
            data = r.json()
        except Exception:
            return []

        # EDINET returns HTTP 200 even for auth errors; check metadata.status
        meta = data.get("metadata", {})
        status = str(meta.get("status", "200"))
        if status != "200":
            msg = meta.get("message", "")
            if status in ("400", "401"):
                raise EdinetApiError(
                    f"EDINET APIエラー (status={status}): {msg}\n"
                    "  → APIキーが未設定または無効です。\n"
                    "  → GitHub Secrets に EDINET_API_KEY を設定してください。\n"
                    "  → https://disclosure.edinet-api.go.jp/ でAPIキーを取得できます。"
                )
            print(f"    EDINET warning: status={status} message={msg}")
            return []

        return data.get("results") or []

    def validate_api_key(self) -> None:
        """Call once at startup to fail fast on auth errors."""
        test_date = date.today() - timedelta(days=30)
        r = self._get("documents.json", {"date": test_date.strftime("%Y-%m-%d"), "type": 2})
        if r is None:
            raise EdinetApiError("EDINET APIに接続できませんでした。")
        data = r.json()
        status = str(data.get("metadata", {}).get("status", "200"))
        if status in ("400", "401"):
            msg = data.get("metadata", {}).get("message", "")
            raise EdinetApiError(
                f"EDINET API認証エラー (status={status}): {msg}\n"
                "  → GitHub Secrets に EDINET_API_KEY を設定してください。\n"
                "  → https://disclosure.edinet-api.go.jp/ でAPIキーを取得できます。"
            )

    # ------------------------------------------------------------------ #
    # Date generation helpers                                              #
    # ------------------------------------------------------------------ #

    def _all_dates_for_years(self, years_back: int) -> list[date]:
        """Return all dates (ordered by filing likelihood) for the past *years_back* years.

        Ordering: for each year, iterate months in _MONTH_PRIORITY order (June-first)
        so that March-FY companies (most common) are found quickly.
        """
        today = date.today()
        dates = []
        for year_offset in range(years_back):
            year = today.year - year_offset
            for month in _MONTH_PRIORITY:
                _, last_day = calendar.monthrange(year, month)
                for day in range(1, last_day + 1):
                    try:
                        d = date(year, month, day)
                        if d <= today:
                            dates.append(d)
                    except ValueError:
                        pass
        return dates

    # ------------------------------------------------------------------ #
    # Public API                                                           #
    # ------------------------------------------------------------------ #

    def search_company(self, name: str) -> Optional[dict]:
        """Search for a listed company by name and return its EDINET metadata."""
        print("  企業を検索中", end="", flush=True)
        for i, d in enumerate(self._all_dates_for_years(7)):
            if i % 20 == 0:
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
        """Return up to *years* annual-report documents for the company.

        Uses daily search ordered by filing-month likelihood so that common
        March-FY filers are found quickly and the loop exits early.
        """
        found: dict[str, dict] = {}

        print("  有価証券報告書を検索中", end="", flush=True)
        for i, d in enumerate(self._all_dates_for_years(years + 2)):
            if len(found) >= years:
                break
            if i % 20 == 0:
                print(".", end="", flush=True)

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

        print(f" {len(found)}件見つかりました")
        return sorted(found.values(), key=lambda x: x.get("periodEnd", ""), reverse=True)[:years]

    def download_xbrl(self, doc_id: str) -> Optional[bytes]:
        """Download the XBRL zip (type=5) for a document."""
        r = self._get(f"documents/{doc_id}", {"type": 5})
        if r is not None:
            return r.content
        return None
