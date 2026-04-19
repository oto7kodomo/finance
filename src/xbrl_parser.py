"""Parse EDINET XBRL / iXBRL / XBRL-to-CSV zip files and extract financial figures."""

import io
import re
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

import pandas as pd
from lxml import etree


XBRLI_NS = "http://www.xbrl.org/2003/instance"
IX_NS    = "http://www.xbrl.org/2013/inlineXBRL"
IX_NS_OLD = "http://www.xbrl.org/2008/inlineXBRL"

ELEMENT_MAP: dict[str, list[str]] = {
    "net_sales": [
        "NetSales", "NetSalesSummaryOfBusinessResults",
        "RevenueIFRS", "Revenue", "Revenues",
        "SalesAndOperatingRevenues", "NetRevenues",
    ],
    "gross_profit": ["GrossProfit", "GrossProfitLoss"],
    "operating_income": [
        "OperatingIncome", "OperatingIncomeLoss",
        "OperatingProfit", "OperatingProfitLoss",
        "ProfitFromOperatingActivitiesIFRS",
    ],
    "ordinary_income": ["OrdinaryIncome", "OrdinaryIncomeLoss", "OrdinaryProfit"],
    "net_income": [
        "ProfitAttributableToOwnersOfParent",
        "ProfitLossAttributableToOwnersOfParent",
        "NetIncome", "NetIncomeLoss", "ProfitLoss",
    ],
    "total_assets": ["Assets", "TotalAssets"],
    "current_assets": ["CurrentAssets"],
    "noncurrent_assets": ["NoncurrentAssets", "NonCurrentAssets"],
    "total_liabilities": ["Liabilities", "TotalLiabilities"],
    "current_liabilities": ["CurrentLiabilities"],
    "noncurrent_liabilities": ["NoncurrentLiabilities", "NonCurrentLiabilities"],
    "net_assets": [
        "NetAssets", "TotalNetAssets", "Equity", "TotalEquity",
        "EquityAttributableToOwnersOfParent",
    ],
    "cash_and_equivalents": [
        "CashAndCashEquivalents",
        "CashAndCashEquivalentsAtEndOfPeriod",
        "CashAndDeposits",
    ],
    "operating_cf": [
        "NetCashProvidedByUsedInOperatingActivities",
        "CashFlowsFromUsedInOperatingActivities",
    ],
    "investing_cf": [
        "NetCashProvidedByUsedInInvestingActivities",
        "CashFlowsFromUsedInInvestingActivities",
    ],
    "financing_cf": [
        "NetCashProvidedByUsedInFinancingActivities",
        "CashFlowsFromUsedInFinancingActivities",
    ],
    "eps": [
        "BasicEarningsLossPerShare", "EarningsPerShare",
        "NetIncomePerShare", "BasicEarningsPerShare",
    ],
    "bvps": ["BookValuePerShare", "NetAssetsPerShare"],
    "dividends_per_share": ["DividendsPerShare", "CashDividendsPerShare"],
}

BS_KEYS = {
    "total_assets", "current_assets", "noncurrent_assets",
    "total_liabilities", "current_liabilities", "noncurrent_liabilities",
    "net_assets", "cash_and_equivalents",
}


class XBRLParser:

    def parse(self, zip_bytes: bytes, period_end: str) -> Optional[dict]:
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                fname, fmt = self._find_instance_doc(zf)
                if not fname:
                    names = zf.namelist()
                    print(f"    インスタンス文書なし。ZIP内容: {names[:6]}")
                    return None
                content = zf.read(fname)
        except (zipfile.BadZipFile, KeyError) as e:
            print(f"    ZIPエラー: {e}")
            return None

        if fmt == "csv":
            return self._parse_csv(content, period_end)
        if fmt == "ixbrl":
            return self._parse_ixbrl(content, period_end)
        return self._parse_xbrl(content, period_end)

    # ------------------------------------------------------------------ #
    # File discovery                                                        #
    # ------------------------------------------------------------------ #

    def _find_instance_doc(self, zf: zipfile.ZipFile) -> tuple[Optional[str], str]:
        names = zf.namelist()

        # 1. Pure XBRL (.xbrl)
        xbrl = [
            n for n in names
            if n.endswith(".xbrl") and not re.search(r"_(cal|def|lab|pre)\.", n)
        ]
        pref = [n for n in xbrl if "asr" in n or "jpcrp" in n]
        if pref:
            return pref[0], "xbrl"
        if xbrl:
            return xbrl[0], "xbrl"

        # 2. Inline XBRL (.htm / .xhtml)
        htm = [n for n in names if n.lower().endswith((".htm", ".xhtml"))]
        pref = [n for n in htm if "asr" in n or "jpcrp" in n]
        if pref:
            return pref[0], "ixbrl"
        if htm:
            return htm[0], "ixbrl"

        # 3. XBRL-to-CSV (.csv) — EDINET newer format
        csv_files = [n for n in names if n.lower().endswith(".csv")]
        pref = [n for n in csv_files if "asr" in n or "jpcrp" in n]
        if pref:
            return pref[0], "csv"
        if csv_files:
            return csv_files[0], "csv"

        return None, ""

    # ------------------------------------------------------------------ #
    # Shared context / fact extraction                                     #
    # ------------------------------------------------------------------ #

    def _parse_contexts_from_root(self, root: etree._Element) -> dict:
        contexts: dict[str, dict] = {}
        for ctx in root.iter(f"{{{XBRLI_NS}}}context"):
            cid = ctx.get("id")
            if not cid:
                continue
            period = ctx.find(f"{{{XBRLI_NS}}}period")
            if period is None:
                continue
            instant = period.find(f"{{{XBRLI_NS}}}instant")
            start   = period.find(f"{{{XBRLI_NS}}}startDate")
            end     = period.find(f"{{{XBRLI_NS}}}endDate")
            if instant is not None and instant.text:
                contexts[cid] = {"type": "instant", "date": instant.text.strip()}
            elif start is not None and end is not None and start.text and end.text:
                contexts[cid] = {
                    "type": "duration",
                    "start": start.text.strip(),
                    "end":   end.text.strip(),
                }
        return contexts

    def _is_full_year(self, ctx: dict) -> bool:
        try:
            s = datetime.strptime(ctx["start"], "%Y-%m-%d").date()
            e = datetime.strptime(ctx["end"],   "%Y-%m-%d").date()
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
            # Prefer consolidated (exclude NonConsolidated contexts)
            consolidated = [
                (c, v) for c, v in pool
                if "NonConsolidated" not in c and "Nonconsolidated" not in c
            ]
            selected = consolidated if consolidated else pool
            if selected:
                return selected[0][1]
        return None

    def _extract_data(
        self, facts: dict, contexts: dict, period_end: str
    ) -> Optional[dict]:
        target = period_end[:10]
        dur_ctx  = self._duration_contexts(contexts, target)
        inst_ctx = self._instant_contexts(contexts, target)

        if not dur_ctx:
            dur_ctx = [
                cid for cid, c in contexts.items()
                if c["type"] == "duration" and self._is_full_year(c)
            ]
        if not inst_ctx:
            inst_ctx = [cid for cid, c in contexts.items() if c["type"] == "instant"]

        data: dict = {"period": period_end}
        for key, names in ELEMENT_MAP.items():
            ctx_ids = inst_ctx if key in BS_KEYS else dur_ctx
            v = self._find_value(facts, names, ctx_ids)
            if v is not None:
                data[key] = v

        return data if len(data) > 1 else None

    # ------------------------------------------------------------------ #
    # XBRL-to-CSV parser                                                   #
    # ------------------------------------------------------------------ #

    def _parse_csv(self, content: bytes, period_end: str) -> Optional[dict]:
        df = None
        for enc in ("utf-8-sig", "utf-8", "cp932", "shift-jis"):
            try:
                df = pd.read_csv(io.BytesIO(content), encoding=enc, dtype=str, header=0)
                break
            except Exception:
                continue

        if df is None or df.empty:
            print("    CSV読み込み失敗")
            return None

        cols = list(df.columns)

        # Identify columns by name patterns
        def find_col(*keywords):
            for kw in keywords:
                for c in cols:
                    if kw in c:
                        return c
            return None

        elem_col  = find_col("要素", "element", "Element")
        ctx_col   = find_col("コンテキスト", "context", "Context")
        start_col = find_col("開始", "start", "Start")
        end_col   = find_col("終了", "時点", "end", "End")
        val_col   = find_col("値", "value", "Value")

        # Positional fallback (standard EDINET XBRL_TO_CSV layout)
        # [要素ID, コンテキストID, 開始日, 終了日/時点, 単位, 小数点桁数, 値]
        if not elem_col or not val_col:
            if len(cols) >= 7:
                elem_col, ctx_col, start_col, end_col = cols[0], cols[1], cols[2], cols[3]
                val_col = cols[6]
            elif len(cols) >= 5:
                elem_col, ctx_col, end_col, val_col = cols[0], cols[1], cols[3], cols[4]
            else:
                print(f"    CSV列不明: {cols}")
                return None

        facts: dict[str, list] = {}
        contexts: dict[str, dict] = {}

        for _, row in df.iterrows():
            elem_id   = str(row.get(elem_col,  "")).strip()
            ctx_id    = str(row.get(ctx_col,   "")).strip()
            val_str   = str(row.get(val_col,   "")).strip()
            start_str = str(row.get(start_col, "")).strip() if start_col else ""
            end_str   = str(row.get(end_col,   "")).strip() if end_col   else ""

            if not elem_id or not val_str or val_str in ("nan", ""):
                continue

            local = elem_id.split(":")[-1] if ":" in elem_id else elem_id

            # Build context entry once
            if ctx_id and ctx_id not in contexts:
                if start_str and end_str and start_str not in ("nan", "") and end_str not in ("nan", ""):
                    contexts[ctx_id] = {
                        "type": "duration",
                        "start": start_str,
                        "end":   end_str,
                    }
                elif end_str and end_str not in ("nan", ""):
                    contexts[ctx_id] = {"type": "instant", "date": end_str}

            # Parse numeric value
            cleaned = re.sub(r"[,，\s\u00a0]", "", val_str)
            negative = cleaned.startswith("△") or cleaned.startswith("-")
            cleaned = cleaned.lstrip("△")
            try:
                v = float(Decimal(cleaned))
                if negative and v > 0:
                    v = -v
                facts.setdefault(local, []).append((ctx_id, v))
            except (InvalidOperation, ValueError):
                continue

        if not facts:
            print("    CSV: 数値データなし")
            return None

        return self._extract_data(facts, contexts, period_end)

    # ------------------------------------------------------------------ #
    # Pure XBRL parser                                                     #
    # ------------------------------------------------------------------ #

    def _parse_xbrl(self, content: bytes, period_end: str) -> Optional[dict]:
        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError as e:
            print(f"    XBRL XMLエラー: {e}")
            return None

        contexts = self._parse_contexts_from_root(root)
        facts: dict[str, list] = {}

        for elem in root.iter():
            ctx_ref = elem.get("contextRef")
            if ctx_ref is None or not (elem.text or "").strip():
                continue
            if elem.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true":
                continue
            try:
                v = float(Decimal(elem.text.strip()))
            except (InvalidOperation, ValueError):
                continue
            local = etree.QName(elem.tag).localname
            facts.setdefault(local, []).append((ctx_ref, v))

        return self._extract_data(facts, contexts, period_end)

    # ------------------------------------------------------------------ #
    # Inline XBRL (iXBRL) parser                                          #
    # ------------------------------------------------------------------ #

    def _parse_ixbrl(self, content: bytes, period_end: str) -> Optional[dict]:
        try:
            root = etree.fromstring(content)
        except etree.XMLSyntaxError:
            try:
                root = etree.fromstring(content, etree.HTMLParser(encoding="utf-8"))
            except Exception as e:
                print(f"    iXBRL パースエラー: {e}")
                return None

        contexts = self._parse_contexts_from_root(root)
        facts: dict[str, list] = {}

        for ns in (IX_NS, IX_NS_OLD):
            for elem in root.iter(f"{{{ns}}}nonFraction"):
                name_attr = elem.get("name", "")
                local = name_attr.split(":")[-1] if ":" in name_attr else name_attr
                ctx_ref = elem.get("contextRef", "")

                raw = "".join(elem.itertext()).strip()
                cleaned = re.sub(r"[,，\s\u00a0]", "", raw)
                if not cleaned or cleaned in ("-", "－", "—", "△"):
                    continue

                negative = cleaned.startswith("△") or elem.get("sign") == "-"
                cleaned = cleaned.lstrip("△")

                try:
                    v = Decimal(cleaned)
                    if negative:
                        v = -v
                    scale = elem.get("scale")
                    if scale:
                        v = v * Decimal(10) ** int(scale)
                    facts.setdefault(local, []).append((ctx_ref, float(v)))
                except (InvalidOperation, ValueError, ArithmeticError):
                    continue

        if not facts:
            print("    iXBRL: 数値データなし")
            return None

        return self._extract_data(facts, contexts, period_end)
