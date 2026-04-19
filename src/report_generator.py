"""Generate HTML analysis report with embedded charts and a console summary."""

import base64
import io
import os
from datetime import datetime
from typing import Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib import rcParams

# Use Japanese fonts when available (installed by GitHub Actions workflow)
for font in ["Noto Sans CJK JP", "Noto Sans JP", "IPAexGothic", "Hiragino Kaku Gothic Pro", "Meiryo"]:
    try:
        rcParams["font.family"] = font
        plt.figure()
        plt.close()
        break
    except Exception:
        continue


# --------------------------------------------------------------------------- #
# Formatting helpers                                                            #
# --------------------------------------------------------------------------- #

def _yen(value: Optional[float], unit: str = "百万円") -> str:
    if value is None:
        return "—"
    v = value / 1_000_000  # convert raw yen → millions
    if abs(v) >= 1_000_000:
        return f"{v / 1_000_000:,.1f} 兆円"
    if abs(v) >= 100:
        return f"{v / 100:,.1f} 億円"
    return f"{v:,.0f} {unit}"


def _pct_str(v: Optional[float]) -> str:
    return f"{v:.1f}%" if v is not None else "—"


def _chg(v: Optional[float]) -> str:
    if v is None:
        return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.1f}%"


def _m(value: Optional[float]) -> Optional[float]:
    """Convert raw yen to millions of yen."""
    return value / 1_000_000 if value is not None else None


# --------------------------------------------------------------------------- #
# Chart helpers                                                                 #
# --------------------------------------------------------------------------- #

COLORS = ["#2196F3", "#4CAF50", "#FF9800", "#F44336", "#9C27B0"]


def _fig_to_b64(fig: plt.Figure) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def _bar_line_chart(
    periods: list[str],
    bar_data: list[tuple[str, list]],
    line_data: list[tuple[str, list]],
    ylabel_bar: str,
    ylabel_line: str,
    title: str,
) -> str:
    fig, ax1 = plt.subplots(figsize=(9, 4.5))
    ax2 = ax1.twinx()

    x = range(len(periods))
    width = 0.35
    n_bars = len(bar_data)
    offsets = [i * width - (n_bars - 1) * width / 2 for i in range(n_bars)]

    for (label, values), offset, color in zip(bar_data, offsets, COLORS):
        vals = [v if v is not None else 0 for v in values]
        ax1.bar([xi + offset for xi in x], vals, width, label=label, color=color, alpha=0.8)

    for (label, values), color in zip(line_data, COLORS[len(bar_data):]):
        vals = [v if v is not None else float("nan") for v in values]
        ax2.plot(x, vals, "o-", label=label, color=color, linewidth=2, markersize=5)

    ax1.set_xticks(list(x))
    ax1.set_xticklabels(periods, fontsize=9)
    ax1.set_ylabel(ylabel_bar, fontsize=9)
    ax2.set_ylabel(ylabel_line, fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax1.set_title(title, fontsize=11, pad=10)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc="upper left")
    fig.tight_layout()
    return _fig_to_b64(fig)


def _line_chart(
    periods: list[str],
    series: list[tuple[str, list]],
    ylabel: str,
    title: str,
) -> str:
    fig, ax = plt.subplots(figsize=(9, 4))
    x = range(len(periods))
    for (label, values), color in zip(series, COLORS):
        vals = [v if v is not None else float("nan") for v in values]
        ax.plot(x, vals, "o-", label=label, color=color, linewidth=2, markersize=5)
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, pad=10)
    ax.legend(fontsize=8)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _fig_to_b64(fig)


def _stacked_bar_chart(
    periods: list[str],
    segments: list[tuple[str, list]],
    ylabel: str,
    title: str,
) -> str:
    fig, ax = plt.subplots(figsize=(9, 4.5))
    x = range(len(periods))
    bottoms = [0.0] * len(periods)
    for (label, values), color in zip(segments, COLORS):
        vals = [v if v is not None else 0 for v in values]
        ax.bar(list(x), vals, bottom=bottoms, label=label, color=color, alpha=0.85)
        bottoms = [b + v for b, v in zip(bottoms, vals)]
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, pad=10)
    ax.legend(fontsize=8, loc="upper left")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _fig_to_b64(fig)


# --------------------------------------------------------------------------- #
# Written analysis                                                              #
# --------------------------------------------------------------------------- #

def _written_analysis(company: str, analysis: dict) -> str:
    s = analysis["summary"]
    raw = analysis["raw"]
    ratios = analysis["ratios"]
    trends = analysis["trends"]
    periods = analysis["periods"]

    paragraphs = []

    # Revenue trend
    sales_vals = [d.get("net_sales") for d in raw]
    valid_sales = [v for v in sales_vals if v is not None]
    if len(valid_sales) >= 2:
        total_growth = (valid_sales[-1] - valid_sales[0]) / abs(valid_sales[0]) * 100
        direction = "増収" if total_growth > 0 else "減収"
        paragraphs.append(
            f"【売上高】{company}の売上高は過去{len(raw)}期で"
            f"{abs(total_growth):.1f}%の{direction}となりました。"
            f"最新期の売上高は{_yen(valid_sales[-1])}です。"
        )

    # Profitability
    op_margin = s.get("operating_margin")
    net_margin = s.get("net_margin")
    if op_margin is not None:
        assessment = "高収益" if op_margin >= 10 else ("標準的" if op_margin >= 5 else "低収益")
        paragraphs.append(
            f"【収益性】最新期の営業利益率は{op_margin:.1f}%（{assessment}水準）、"
            f"純利益率は{_pct_str(net_margin)}です。"
        )

    # ROE / ROA
    roe = s.get("roe")
    roa = s.get("roa")
    if roe is not None:
        roe_comment = "良好（8%超）" if roe >= 8 else "要改善"
        paragraphs.append(
            f"【効率性】ROEは{roe:.1f}%（{roe_comment}）、ROAは{_pct_str(roa)}です。"
        )

    # Financial soundness
    eq_ratio = s.get("equity_ratio")
    if eq_ratio is not None:
        health = "健全" if eq_ratio >= 40 else ("要注意" if eq_ratio >= 20 else "脆弱")
        paragraphs.append(
            f"【財務健全性】自己資本比率は{eq_ratio:.1f}%（{health}水準）です。"
        )

    # Cash flow
    ocf = s.get("ocf")
    if ocf is not None:
        ocf_comment = "安定的な営業キャッシュフローを創出しています。" if ocf > 0 else "営業キャッシュフローがマイナスとなっており、注意が必要です。"
        paragraphs.append(f"【キャッシュフロー】最新期の営業CFは{_yen(ocf)}で、{ocf_comment}")

    return "\n".join(f"<p>{p}</p>" for p in paragraphs)


# --------------------------------------------------------------------------- #
# Main report generator                                                         #
# --------------------------------------------------------------------------- #

class ReportGenerator:
    def generate(self, company: str, analysis: dict) -> str:
        periods = analysis["periods"]
        raw = analysis["raw"]
        ratios = analysis["ratios"]
        trends = analysis["trends"]

        # Prepare million-yen series
        sales_m = [_m(d.get("net_sales")) for d in raw]
        op_m = [_m(d.get("operating_income")) for d in raw]
        ni_m = [_m(d.get("net_income")) for d in raw]
        assets_m = [_m(d.get("total_assets")) for d in raw]
        net_assets_m = [_m(d.get("net_assets")) for d in raw]
        liabilities_m = [_m(d.get("total_liabilities")) for d in raw]
        ocf_m = [_m(d.get("operating_cf")) for d in raw]
        icf_m = [_m(d.get("investing_cf")) for d in raw]
        fcf_m = [_m(d.get("financing_cf")) for d in raw]

        op_margins = [r.get("operating_margin") for r in ratios]
        net_margins = [r.get("net_margin") for r in ratios]
        roes = [r.get("roe") for r in ratios]
        roas = [r.get("roa") for r in ratios]
        eq_ratios = [r.get("equity_ratio") for r in ratios]

        # Chart 1: Revenue & profits
        chart1 = _bar_line_chart(
            periods,
            bar_data=[("売上高", sales_m), ("営業利益", op_m), ("純利益", ni_m)],
            line_data=[("営業利益率(%)", op_margins)],
            ylabel_bar="百万円",
            ylabel_line="利益率(%)",
            title="売上高・利益の推移",
        )

        # Chart 2: Profitability ratios
        chart2 = _line_chart(
            periods,
            series=[
                ("営業利益率(%)", op_margins),
                ("純利益率(%)", net_margins),
                ("ROE(%)", roes),
                ("ROA(%)", roas),
            ],
            ylabel="%",
            title="収益性指標の推移",
        )

        # Chart 3: Financial structure (latest year)
        chart3 = _stacked_bar_chart(
            periods,
            segments=[
                ("純資産", net_assets_m),
                ("負債", liabilities_m),
            ],
            ylabel="百万円",
            title="資産・負債・純資産の推移",
        )

        # Chart 4: Cash flow
        chart4 = _bar_line_chart(
            periods,
            bar_data=[
                ("営業CF", ocf_m),
                ("投資CF", icf_m),
                ("財務CF", fcf_m),
            ],
            line_data=[],
            ylabel_bar="百万円",
            ylabel_line="",
            title="キャッシュフローの推移",
        )

        # Chart 5: Equity ratio
        chart5 = _line_chart(
            periods,
            series=[("自己資本比率(%)", eq_ratios)],
            ylabel="%",
            title="自己資本比率の推移",
        )

        analysis_text = _written_analysis(company, analysis)
        html = self._build_html(company, analysis, [chart1, chart2, chart3, chart4, chart5], analysis_text)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"report_{company}_{ts}.html"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html)
        return filename

    # ------------------------------------------------------------------ #

    def _build_html(self, company: str, analysis: dict, charts: list[str], analysis_text: str) -> str:
        periods = analysis["periods"]
        raw = analysis["raw"]
        ratios = analysis["ratios"]
        trends = analysis["trends"]
        s = analysis["summary"]

        trend_map = {t["period"]: t for t in trends}

        def row(label, values, fmt=_yen):
            cells = "".join(f"<td>{fmt(v)}</td>" for v in values)
            return f"<tr><th>{label}</th>{cells}</tr>"

        def ratio_row(label, key, fmt=_pct_str):
            cells = "".join(f"<td>{fmt(r.get(key))}</td>" for r in ratios)
            return f"<tr><th>{label}</th>{cells}</tr>"

        period_headers = "".join(f"<th>{p}</th>" for p in periods)

        chart_tags = "".join(
            f'<div class="chart"><img src="data:image/png;base64,{c}" alt="chart"></div>'
            for c in charts
        )

        highlights = f"""
        <div class="highlights">
            <div class="kpi"><span class="kpi-label">売上高</span><span class="kpi-value">{_yen(s.get('net_sales'))}</span></div>
            <div class="kpi"><span class="kpi-label">営業利益</span><span class="kpi-value">{_yen(s.get('operating_income'))}</span></div>
            <div class="kpi"><span class="kpi-label">純利益</span><span class="kpi-value">{_yen(s.get('net_income'))}</span></div>
            <div class="kpi"><span class="kpi-label">総資産</span><span class="kpi-value">{_yen(s.get('total_assets'))}</span></div>
            <div class="kpi"><span class="kpi-label">営業利益率</span><span class="kpi-value">{_pct_str(s.get('operating_margin'))}</span></div>
            <div class="kpi"><span class="kpi-label">ROE</span><span class="kpi-value">{_pct_str(s.get('roe'))}</span></div>
            <div class="kpi"><span class="kpi-label">自己資本比率</span><span class="kpi-value">{_pct_str(s.get('equity_ratio'))}</span></div>
            <div class="kpi"><span class="kpi-label">EPS</span><span class="kpi-value">{f"{s.get('eps'):.1f}円" if s.get('eps') else '—'}</span></div>
        </div>"""

        return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{company} 財務分析レポート</title>
<style>
  body {{ font-family: "Noto Sans JP", "Hiragino Kaku Gothic Pro", Meiryo, sans-serif;
         margin: 0; background: #f4f6f9; color: #333; }}
  .container {{ max-width: 1100px; margin: 0 auto; padding: 24px; }}
  h1 {{ font-size: 1.6rem; color: #1a237e; border-bottom: 3px solid #1a237e; padding-bottom: 8px; }}
  h2 {{ font-size: 1.15rem; color: #283593; margin-top: 32px; border-left: 4px solid #3f51b5;
        padding-left: 10px; }}
  .meta {{ color: #666; font-size: 0.85rem; margin-bottom: 24px; }}
  .highlights {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 24px; }}
  .kpi {{ background: #fff; border-radius: 8px; padding: 14px 20px; min-width: 140px;
           box-shadow: 0 1px 4px rgba(0,0,0,.1); flex: 1; }}
  .kpi-label {{ display: block; font-size: 0.75rem; color: #888; margin-bottom: 4px; }}
  .kpi-value {{ display: block; font-size: 1.2rem; font-weight: bold; color: #1a237e; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           box-shadow: 0 1px 4px rgba(0,0,0,.08); border-radius: 6px;
           overflow: hidden; margin-bottom: 24px; font-size: 0.88rem; }}
  th, td {{ padding: 9px 12px; text-align: right; border-bottom: 1px solid #e8eaf6; }}
  th:first-child {{ text-align: left; background: #e8eaf6; font-weight: 600; width: 200px; }}
  tr:hover td {{ background: #f5f5f5; }}
  thead th {{ background: #3f51b5; color: #fff; font-weight: 600; }}
  .chart {{ background: #fff; border-radius: 8px; padding: 16px;
            box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 20px; }}
  .chart img {{ width: 100%; height: auto; }}
  .analysis {{ background: #fff; border-radius: 8px; padding: 20px 24px;
               box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 24px; line-height: 1.8; }}
  .analysis p {{ margin: 0 0 10px; }}
  .charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
  .charts-grid .chart:first-child {{ grid-column: 1 / -1; }}
  @media (max-width: 700px) {{ .charts-grid {{ grid-template-columns: 1fr; }} }}
  footer {{ text-align: center; font-size: 0.78rem; color: #aaa; margin-top: 32px; }}
</style>
</head>
<body>
<div class="container">
  <h1>{company} 財務分析レポート（過去{len(periods)}期）</h1>
  <p class="meta">データソース: EDINET（金融庁電子開示システム）　作成日: {datetime.now().strftime('%Y年%m月%d日')}</p>

  <h2>最新期ハイライト（{periods[-1]}期）</h2>
  {highlights}

  <h2>チャート</h2>
  <div class="charts-grid">
    {''.join(f'<div class="chart"><img src="data:image/png;base64,{c}" alt="chart"></div>' for c in charts)}
  </div>

  <h2>過去{len(periods)}期 財務サマリー（単位: 百万円）</h2>
  <table>
    <thead><tr><th>項目</th>{period_headers}</tr></thead>
    <tbody>
      {row("売上高", [_m(d.get("net_sales")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("営業利益", [_m(d.get("operating_income")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("経常利益", [_m(d.get("ordinary_income")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("当期純利益", [_m(d.get("net_income")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("総資産", [_m(d.get("total_assets")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("純資産", [_m(d.get("net_assets")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("営業CF", [_m(d.get("operating_cf")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("投資CF", [_m(d.get("investing_cf")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
      {row("財務CF", [_m(d.get("financing_cf")) for d in raw], fmt=lambda v: f"{v:,.0f}" if v else "—")}
    </tbody>
  </table>

  <h2>財務指標</h2>
  <table>
    <thead><tr><th>指標</th>{period_headers}</tr></thead>
    <tbody>
      {ratio_row("営業利益率(%)", "operating_margin")}
      {ratio_row("純利益率(%)", "net_margin")}
      {ratio_row("ROE(%)", "roe")}
      {ratio_row("ROA(%)", "roa")}
      {ratio_row("自己資本比率(%)", "equity_ratio")}
      {ratio_row("流動比率(%)", "current_ratio")}
      {ratio_row("営業CF/売上高(%)", "ocf_margin")}
    </tbody>
  </table>

  <h2>前期比成長率</h2>
  <table>
    <thead><tr><th>項目</th>{"".join(f"<th>{t['period']}</th>" for t in trends)}</tr></thead>
    <tbody>
      <tr><th>売上高成長率</th>{"".join(f"<td>{_chg(t.get('net_sales_growth'))}</td>" for t in trends)}</tr>
      <tr><th>営業利益成長率</th>{"".join(f"<td>{_chg(t.get('operating_income_growth'))}</td>" for t in trends)}</tr>
      <tr><th>純利益成長率</th>{"".join(f"<td>{_chg(t.get('net_income_growth'))}</td>" for t in trends)}</tr>
      <tr><th>総資産成長率</th>{"".join(f"<td>{_chg(t.get('total_assets_growth'))}</td>" for t in trends)}</tr>
    </tbody>
  </table>

  <h2>分析コメント</h2>
  <div class="analysis">
    {analysis_text}
  </div>

  <footer>本レポートはEDINET公開データに基づく自動生成レポートです。投資判断の参考情報として御覧ください。</footer>
</div>
</body>
</html>"""

    # ------------------------------------------------------------------ #
    # Console summary                                                       #
    # ------------------------------------------------------------------ #

    def print_summary(self, company: str, analysis: dict) -> None:
        s = analysis["summary"]
        periods = analysis["periods"]

        print(f"\n{'='*55}")
        print(f"  {company}  財務分析サマリー")
        print(f"  対象期間: {periods[0]} ～ {periods[-1]}")
        print(f"{'='*55}")
        print(f"  売上高          : {_yen(s.get('net_sales'))}")
        print(f"  営業利益        : {_yen(s.get('operating_income'))}")
        print(f"  純利益          : {_yen(s.get('net_income'))}")
        print(f"  総資産          : {_yen(s.get('total_assets'))}")
        print(f"  純資産          : {_yen(s.get('net_assets'))}")
        print(f"{'─'*55}")
        print(f"  営業利益率      : {_pct_str(s.get('operating_margin'))}")
        print(f"  純利益率        : {_pct_str(s.get('net_margin'))}")
        print(f"  ROE             : {_pct_str(s.get('roe'))}")
        print(f"  ROA             : {_pct_str(s.get('roa'))}")
        print(f"  自己資本比率    : {_pct_str(s.get('equity_ratio'))}")
        print(f"{'─'*55}")
        print(f"  売上高成長率    : {_chg(s.get('net_sales_growth'))}")
        print(f"  営業利益成長率  : {_chg(s.get('operating_income_growth'))}")
        print(f"  営業CF          : {_yen(s.get('ocf'))}")
        if s.get("eps"):
            print(f"  EPS             : {s['eps']:.1f}円")
        if s.get("dividends_per_share"):
            print(f"  1株配当         : {s['dividends_per_share']:.1f}円")
        print(f"{'='*55}\n")
