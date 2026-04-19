"""Compute financial ratios and trend indicators from parsed XBRL data."""

from typing import Optional


def _pct(numerator, denominator) -> Optional[float]:
    if numerator is None or denominator is None or denominator == 0:
        return None
    return numerator / denominator * 100


def _growth(current, previous) -> Optional[float]:
    if current is None or previous is None or previous == 0:
        return None
    return (current - previous) / abs(previous) * 100


class FinancialAnalyzer:
    def analyze(self, data_list: list[dict]) -> dict:
        """
        *data_list* is sorted ascending by period.
        Returns a dict with:
          - periods: list of period labels
          - raw: list of raw financial dicts
          - ratios: list of ratio dicts (one per period)
          - trends: growth rates between consecutive periods
          - summary: dict of latest-year highlights
        """
        periods = [d["period"][:7] for d in data_list]  # YYYY-MM
        ratios = [self._calc_ratios(d) for d in data_list]
        trends = self._calc_trends(data_list)
        summary = self._summarize(data_list[-1], ratios[-1], trends)

        return {
            "periods": periods,
            "raw": data_list,
            "ratios": ratios,
            "trends": trends,
            "summary": summary,
        }

    # ------------------------------------------------------------------ #
    # Per-period ratios                                                    #
    # ------------------------------------------------------------------ #

    def _calc_ratios(self, d: dict) -> dict:
        sales = d.get("net_sales")
        op = d.get("operating_income")
        ni = d.get("net_income")
        assets = d.get("total_assets")
        net_assets = d.get("net_assets")
        liabilities = d.get("total_liabilities")
        ocf = d.get("operating_cf")

        return {
            "period": d["period"],
            # Profitability
            "gross_margin": _pct(d.get("gross_profit"), sales),
            "operating_margin": _pct(op, sales),
            "net_margin": _pct(ni, sales),
            "roe": _pct(ni, net_assets),
            "roa": _pct(ni, assets),
            "ebitda_margin": None,  # requires depreciation data not in scope
            # Stability
            "equity_ratio": _pct(net_assets, assets),
            "debt_ratio": _pct(liabilities, assets),
            "current_ratio": _pct(d.get("current_assets"), d.get("current_liabilities")),
            "interest_coverage": (
                op / (op - ocf) if op and ocf and (op - ocf) != 0 else None
            ),
            # Efficiency
            "asset_turnover": (sales / assets if sales and assets else None),
            "ocf_margin": _pct(ocf, sales),
        }

    # ------------------------------------------------------------------ #
    # Growth trends                                                        #
    # ------------------------------------------------------------------ #

    def _calc_trends(self, data_list: list[dict]) -> list[dict]:
        trends = []
        for i in range(1, len(data_list)):
            prev, curr = data_list[i - 1], data_list[i]
            trends.append(
                {
                    "period": curr["period"][:7],
                    "net_sales_growth": _growth(curr.get("net_sales"), prev.get("net_sales")),
                    "operating_income_growth": _growth(
                        curr.get("operating_income"), prev.get("operating_income")
                    ),
                    "net_income_growth": _growth(
                        curr.get("net_income"), prev.get("net_income")
                    ),
                    "total_assets_growth": _growth(
                        curr.get("total_assets"), prev.get("total_assets")
                    ),
                    "net_assets_growth": _growth(
                        curr.get("net_assets"), prev.get("net_assets")
                    ),
                }
            )
        return trends

    # ------------------------------------------------------------------ #
    # Summary                                                              #
    # ------------------------------------------------------------------ #

    def _summarize(self, latest: dict, latest_ratios: dict, trends: list[dict]) -> dict:
        last_trend = trends[-1] if trends else {}
        return {
            "net_sales": latest.get("net_sales"),
            "operating_income": latest.get("operating_income"),
            "net_income": latest.get("net_income"),
            "total_assets": latest.get("total_assets"),
            "net_assets": latest.get("net_assets"),
            "operating_margin": latest_ratios.get("operating_margin"),
            "net_margin": latest_ratios.get("net_margin"),
            "roe": latest_ratios.get("roe"),
            "roa": latest_ratios.get("roa"),
            "equity_ratio": latest_ratios.get("equity_ratio"),
            "net_sales_growth": last_trend.get("net_sales_growth"),
            "operating_income_growth": last_trend.get("operating_income_growth"),
            "ocf": latest.get("operating_cf"),
            "eps": latest.get("eps"),
            "dividends_per_share": latest.get("dividends_per_share"),
        }
