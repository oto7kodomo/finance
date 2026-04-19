"""Parse EDINET XBRL / Inline XBRL (iXBRL) zip files and extract financial figures."""

import io
import re
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

from lxml import etree


XBRLI_NS = "http://www.xbrl.org/2003/instance"
IX_NS = "http://www.xbrl.org/2013/inlineXBRL"
IX_NS_OLD = "http://www.xbrl.org/2008/inlineXBRL"  # older submissions

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
    "gross_profit": ["GrossProfit", "GrossProfitLoss"],
    "operating_income": [
        "OperatingIncome",
        "OperatingIncomeLoss",
        "OperatingProfit",
        "OperatingProfitLoss",
        "ProfitFromOperatingActivitiesIFRS",
        "ProfitLossFromOperatingActivities",
    ],
    "ordinary_income": ["OrdinaryIncome", "OrdinaryIncomeLoss", "OrdinaryProfit"],
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

BS_KEYS = {
    "total_assets",
    "current_assets",
    "noncurrent_assets",
    "total_liabilities",
    "current_liabilities",
    "noncurrent_liabilities",
    "net_assets",
    "cash_and_equivalents",
}


class XBRLParser:
    def parse(self, zip_bytes: bytes, period_end: str) -> Optional[dict]:
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                fname, fmt = self._find_instance_doc(zf)
                if not fname:
                    names = zf.namelist()
                    print(f"    インスタンス文書が見つかりません。ZIPの内容: {names[:8]}")
                    return None
                content = zf.read(fname)
        except (zipfile.BadZipFile, KeyError) as e:
            print(f"    ZIPエラー: {e}")
            return None

        if fmt == "ixbrl":
            return self._parse_ixbrl(content, period_end)
        return self._parse_xbrl(content, period_end)

    # ------------------------------------------------------------------ #
    # File discovery                                                        #
    # ------------------------------------------------------------------ #

    def _find_instance_doc(self, zf: zipfile.ZipFile) -> tuple[Optional[str], str]:
        names = zf.namelist()

        # 1. Pure XBRL instance documents (.xbrl)
        xbrl = [
            n for n in names
            if n.endswith(".xbrl")
            and not re.search(r"_(cal|def|lab|pre|ref)\.", n)
        ]
        preferred = [n for n in xbrl if "asr" in n or "jpcrp" in n or "ifrs" in n.lower()]
        if preferred:
            return preferred[0], "xbrl"
        if xbrl:
            return xbrl[0], "xbrl"

        # 2. Inline XBRL (.htm / .xhtml) – EDINET format since ~2022
        htm = [
            n for n in names
            if (n.lower().endswith(".htm") or n.lower().endswith(".xhtml"))
            and not n.lower().endswith("_lab.htm")
        ]
        preferred_htm = [n for n in htm if "asr" in n or "jpcrp" in n or "ifrs" in n.lower()]
        if preferred_htm:
            return preferred_htm[0], "ixbrl"
        if htm:
            return htm[0], "ixbrl"

        return None, ""

    # ------------------------------------------------------------------ #
    # Context helpers (shared between XBRL and iXBRL)                     #
    # ------------------------------------------------------------------ #

    def _parse_contexts(self, root: etree._Element) -> dict:
        contexts: dict[str, dict] = {}
        for ctx in root.iter(f"{{{XBRLI_NS}}}context"):
            cid = ctx.get("id")
            if not cid:
                continue
            period = ctx.find(f"{{{XBRLI_NS}}}period")
            if period is None:
                continue
            instant = period.find(f"{{{XBRLI_NS}}}instant")
            start = period.find(f"{{{XBRLI_NS}}}startDate")
            end = period.find(f"{{{XBRLI_NS}}}endDate")
            if instant is not None and instant.text:
                contexts[cid] = {"type": "instant", "date": instant.text.strip()}
            elif start is not None and end is not None and start.text and end.text:
                contexts[cid] = {
                    "type": "duration",
                    "start": start.text.strip(),
                    "end": end.text.strip(),
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

    def _find_value(
        self, facts: dict, element_names: list[str], context_ids: list[str]
    ) -> Optional[float]:
        ctx_set = set(context_ids)
        for name in element_names:
            entries = facts.get(name)
            if not entries:
                continue
            in_ctx = [(c, v) for c, v in entries if c in ctx_set]
            pool = in_ctx if in_ctx else entries
            consolidated = [
                (c, v) for c, v in pool
                if "NonConsolidated" not in c and "Nonconsolidated" not in c
            ]
            selected = consolidated if consolidated else pool
            if selected:
                return selected[0][1]
        return None

    def _extract_data(self, facts: dict, contexts: dict, period_end: str) -> Optional[dict]:
        target_end = period_end[:10]
        duration_ctx = self._duration_contexts(contexts, target_end)
        instant_ctx = self._instant_contexts(contexts, target_end)

        # Fallbacks
        if not duration_ctx:
            duration_ctx = [
                cid for cid, c in contexts.items()
                if c["type"] == "duration" and self._is_full_year(c)
            ]
        if not instant_ctx:
            instant_ctx = [cid for cid, c in contexts.items() if c["type"] == "instant"]

        data: dict = {"period": period_end}
        for key, names in ELEMENT_MAP.items():
            ctx_ids = instant_ctx if key in BS_KEYS else duration_ctx
            value = self._find_value(facts, names, ctx_ids)
            if value is not None:
                data[key] = value

        return data if len(data) > 1 else None

    # ------------------------------------------------------------------ #
    # Pure XBRL parser                                                     #
    # ------------------------------------------------------------------ #

    def _parse_xbrl(self, content: bytes, period_end: str) -> Optional[dict]:
        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError as e:
            print(f"    XBRL XMLエラー: {e}")
            return None

        contexts = self._parse_contexts(root)
        facts: dict[str, list] = {}

        for elem in root.iter():
            ctx_ref = elem.get("contextRef")
            if ctx_ref is None or not (elem.text or "").strip():
                continue
            if elem.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true":
                continue
            try:
                value = float(Decimal(elem.text.strip()))
            except (InvalidOperation, ValueError):
                continue
            local = etree.QName(elem.tag).localname
            facts.setdefault(local, []).append((ctx_ref, value))

        return self._extract_data(facts, contexts, period_end)

    # ------------------------------------------------------------------ #
    # Inline XBRL (iXBRL) parser                                          #
    # ------------------------------------------------------------------ #

    def _parse_ixbrl(self, content: bytes, period_end: str) -> Optional[dict]:
        # iXBRL is XHTML – try strict XML parse first, fall back to HTML parser
        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError:
            try:
                root = etree.fromstring(content, etree.HTMLParser(encoding="utf-8"))
            except Exception as e:
                print(f"    iXBRL パースエラー: {e}")
                return None

        contexts = self._parse_contexts(root)

        facts: dict[str, list] = {}
        for ns in (IX_NS, IX_NS_OLD):
            for elem in root.iter(f"{{{ns}}}nonFraction"):
                name_attr = elem.get("name", "")
                local = name_attr.split(":")[-1] if ":" in name_attr else name_attr
                if not local:
                    continue
                ctx_ref = elem.get("contextRef", "")

                # Collect text (join all text nodes, skip child element tags)
                raw = "".join(elem.itertext()).strip()
                # Remove formatting characters
                cleaned = re.sub(r"[,，\s\u00a0]", "", raw)
                if not cleaned or cleaned in ("-", "－", "—", "△"):
                    continue

                # Handle Japanese minus sign △
                negative = cleaned.startswith("△") or elem.get("sign") == "-"
                cleaned = cleaned.lstrip("△")

                try:
                    value = Decimal(cleaned)
                    if negative:
                        value = -value

                    # Apply iXBRL scale factor (e.g. scale="6" → × 10^6)
                    scale = elem.get("scale")
                    if scale:
                        value = value * Decimal(10) ** int(scale)

                    facts.setdefault(local, []).append((ctx_ref, float(value)))
                except (InvalidOperation, ValueError, ArithmeticError):
                    continue

        if not facts:
            print("    iXBRL: 数値データが抽出できませんでした")
            return None

        return self._extract_data(facts, contexts, period_end)
