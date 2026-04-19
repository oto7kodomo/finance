"""Parse EDINET XBRL zip files and extract key financial figures."""

import io
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

from lxml import etree


XBRLI_NS = "http://www.xbrl.org/2003/instance"

# Ordered by preference: first match wins.
ELEMENT_MAP: dict[str, list[str]] = {
    # Income statement
    "net_sales": [
        "NetSales",
        "NetSalesSummaryOfBusinessResults",
        "RevenueIFRS",
        "Revenue",
        "Revenues",
        "SalesAndOperatingRevenues",
        "NetRevenues",
        "SalesOfProductsAndServices",
    ],
    "gross_profit": [
        "GrossProfit",
        "GrossProfitLoss",
    ],
    "operating_income": [
        "OperatingIncome",
        "OperatingIncomeLoss",
        "OperatingProfit",
        "OperatingProfitLoss",
        "ProfitFromOperatingActivitiesIFRS",
        "ProfitLossFromOperatingActivities",
    ],
    "ordinary_income": [
        "OrdinaryIncome",
        "OrdinaryIncomeLoss",
        "OrdinaryProfit",
    ],
    "net_income": [
        "ProfitAttributableToOwnersOfParent",
        "ProfitLossAttributableToOwnersOfParent",
        "NetIncome",
        "NetIncomeLoss",
        "ProfitLoss",
        "NetProfit",
    ],
    # Balance sheet
    "total_assets": ["Assets", "TotalAssets"],
    "current_assets": ["CurrentAssets"],
    "noncurrent_assets": ["NoncurrentAssets", "NonCurrentAssets"],
    "total_liabilities": ["Liabilities", "TotalLiabilities"],
    "current_liabilities": ["CurrentLiabilities"],
    "noncurrent_liabilities": ["NoncurrentLiabilities", "NonCurrentLiabilities"],
    "net_assets": [
        "NetAssets",
        "TotalNetAssets",
        "Equity",
        "TotalEquity",
        "EquityAttributableToOwnersOfParent",
    ],
    "interest_bearing_debt": [
        "InterestBearingDebt",
        "BorrowingsAndBonds",
        "NotesAndAccountsPayableTrade",
    ],
    "cash_and_equivalents": [
        "CashAndCashEquivalents",
        "CashAndCashEquivalentsAtEndOfPeriod",
        "CashAndDeposits",
    ],
    # Cash flow
    "operating_cf": [
        "NetCashProvidedByUsedInOperatingActivities",
        "CashFlowsFromUsedInOperatingActivities",
        "CashProvidedByUsedInOperatingActivities",
    ],
    "investing_cf": [
        "NetCashProvidedByUsedInInvestingActivities",
        "CashFlowsFromUsedInInvestingActivities",
        "CashProvidedByUsedInInvestingActivities",
    ],
    "financing_cf": [
        "NetCashProvidedByUsedInFinancingActivities",
        "CashFlowsFromUsedInFinancingActivities",
        "CashProvidedByUsedInFinancingActivities",
    ],
    # Per-share
    "eps": [
        "BasicEarningsLossPerShare",
        "EarningsPerShare",
        "NetIncomePerShare",
        "BasicEarningsPerShare",
    ],
    "bvps": ["BookValuePerShare", "NetAssetsPerShare"],
    "dividends_per_share": [
        "DividendsPerShare",
        "CashDividendsPerShare",
        "AnnualDividendsPerShare",
    ],
}

# Balance-sheet items use instant contexts; the rest use duration contexts.
BS_KEYS = {
    "total_assets",
    "current_assets",
    "noncurrent_assets",
    "total_liabilities",
    "current_liabilities",
    "noncurrent_liabilities",
    "net_assets",
    "interest_bearing_debt",
    "cash_and_equivalents",
}


class XBRLParser:
    def parse(self, zip_bytes: bytes, period_end: str) -> Optional[dict]:
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                xbrl_name = self._find_instance_doc(zf)
                if not xbrl_name:
                    return None
                xbrl_bytes = zf.read(xbrl_name)
        except (zipfile.BadZipFile, KeyError) as e:
            print(f"    ZIPエラー: {e}")
            return None

        return self._parse_xbrl(xbrl_bytes, period_end)

    # ------------------------------------------------------------------ #
    # Zip helpers                                                          #
    # ------------------------------------------------------------------ #

    def _find_instance_doc(self, zf: zipfile.ZipFile) -> Optional[str]:
        names = zf.namelist()
        candidates = [
            n for n in names
            if n.endswith(".xbrl")
            and not any(n.endswith(s) for s in ("_cal.xbrl", "_def.xbrl", "_lab.xbrl"))
        ]
        # Prefer annual-report instance files (jpcrp / asr in name)
        preferred = [n for n in candidates if "asr" in n or "jpcrp" in n]
        return preferred[0] if preferred else (candidates[0] if candidates else None)

    # ------------------------------------------------------------------ #
    # XBRL parsing                                                         #
    # ------------------------------------------------------------------ #

    def _parse_xbrl(self, content: bytes, period_end: str) -> Optional[dict]:
        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError as e:
            print(f"    XMLパースエラー: {e}")
            return None

        contexts = self._parse_contexts(root)
        facts = self._parse_facts(root)

        target_end = period_end[:10]  # YYYY-MM-DD

        duration_ctx = self._duration_contexts(contexts, target_end)
        instant_ctx = self._instant_contexts(contexts, target_end)

        # Fallback: use any full-year / any instant context
        if not duration_ctx:
            duration_ctx = [
                cid for cid, c in contexts.items()
                if c["type"] == "duration" and self._is_full_year(c)
            ]
        if not instant_ctx:
            instant_ctx = [
                cid for cid, c in contexts.items() if c["type"] == "instant"
            ]

        data: dict = {"period": period_end}
        for key, names in ELEMENT_MAP.items():
            ctx_ids = instant_ctx if key in BS_KEYS else duration_ctx
            value = self._find_value(facts, names, ctx_ids)
            if value is not None:
                data[key] = value

        return data if len(data) > 1 else None

    # ------------------------------------------------------------------ #
    # Context helpers                                                       #
    # ------------------------------------------------------------------ #

    def _parse_contexts(self, root: etree._Element) -> dict:
        ns = {"x": XBRLI_NS}
        contexts: dict[str, dict] = {}
        for ctx in root.findall(".//x:context", ns):
            cid = ctx.get("id")
            period = ctx.find("x:period", ns)
            if period is None:
                continue
            instant_el = period.find("x:instant", ns)
            start_el = period.find("x:startDate", ns)
            end_el = period.find("x:endDate", ns)
            if instant_el is not None and instant_el.text:
                contexts[cid] = {"type": "instant", "date": instant_el.text.strip()}
            elif start_el is not None and end_el is not None:
                contexts[cid] = {
                    "type": "duration",
                    "start": start_el.text.strip(),
                    "end": end_el.text.strip(),
                }
        return contexts

    def _is_full_year(self, ctx: dict) -> bool:
        try:
            s = datetime.strptime(ctx["start"], "%Y-%m-%d").date()
            e = datetime.strptime(ctx["end"], "%Y-%m-%d").date()
            return 340 <= (e - s).days <= 380
        except Exception:
            return False

    def _duration_contexts(self, contexts: dict, target_end: str) -> list[str]:
        return [
            cid for cid, c in contexts.items()
            if c["type"] == "duration"
            and c.get("end", "") == target_end
            and self._is_full_year(c)
        ]

    def _instant_contexts(self, contexts: dict, target_date: str) -> list[str]:
        return [
            cid for cid, c in contexts.items()
            if c["type"] == "instant" and c.get("date", "") == target_date
        ]

    # ------------------------------------------------------------------ #
    # Fact parsing                                                          #
    # ------------------------------------------------------------------ #

    def _parse_facts(self, root: etree._Element) -> dict[str, list[tuple]]:
        """Return {local_name: [(context_id, float_value), ...]}."""
        facts: dict[str, list] = {}
        for elem in root.iter():
            ctx_ref = elem.get("contextRef")
            if ctx_ref is None or not (elem.text or "").strip():
                continue
            # Skip nil values
            if elem.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true":
                continue
            try:
                value = float(Decimal(elem.text.strip()))
            except (InvalidOperation, ValueError):
                continue
            local = etree.QName(elem.tag).localname
            facts.setdefault(local, []).append((ctx_ref, value))
        return facts

    def _find_value(
        self,
        facts: dict,
        element_names: list[str],
        context_ids: list[str],
    ) -> Optional[float]:
        ctx_set = set(context_ids)

        for name in element_names:
            entries = facts.get(name)
            if not entries:
                continue

            in_ctx = [(c, v) for c, v in entries if c in ctx_set]
            pool = in_ctx if in_ctx else entries

            # Prefer consolidated (exclude NonConsolidated contexts)
            consolidated = [
                (c, v) for c, v in pool
                if "NonConsolidated" not in c and "Nonconsolidated" not in c
            ]
            selected = consolidated if consolidated else pool

            if selected:
                return selected[0][1]

        return None
