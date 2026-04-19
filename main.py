#!/usr/bin/env python3
"""
有価証券報告書 財務分析ツール
================================
上場企業名を入力すると EDINET から過去5期分の財務データを取得し、
HTML レポートとコンソール要約を生成します。

使い方:
    python main.py <企業名>
    python main.py <企業名> --edinet-code <E12345>  # コード既知の場合
    python main.py <企業名> --years 3               # 取得期数を変更

事前準備:
    1. pip install -r requirements.txt
    2. .env.example を .env にコピーして EDINET_API_KEY を設定
       (APIキーは https://disclosure.edinet-api.go.jp/ から申請)
"""

import argparse
import os
import sys

from dotenv import load_dotenv

from src.edinet_client import EdinetClient
from src.financial_analyzer import FinancialAnalyzer
from src.report_generator import ReportGenerator
from src.xbrl_parser import XBRLParser


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="上場企業の有価証券報告書から財務分析レポートを生成します",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("company", nargs="?", help="企業名（例: トヨタ自動車）")
    p.add_argument("--edinet-code", metavar="CODE", help="EDINETコード（例: E02144）")
    p.add_argument("--years", type=int, default=5, help="取得する期数 (デフォルト: 5)")
    return p.parse_args()


def main() -> None:
    load_dotenv()
    args = parse_args()

    # Company name
    if args.company:
        company_name = args.company
    else:
        try:
            company_name = input("企業名を入力してください: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n中断しました")
            sys.exit(0)

    if not company_name:
        print("企業名を入力してください", file=sys.stderr)
        sys.exit(1)

    api_key = os.getenv("EDINET_API_KEY", "")
    if not api_key:
        print(
            "警告: EDINET_API_KEY が設定されていません。"
            "2024年4月以降はAPIキーが必須です。\n"
            "  → https://disclosure.edinet-api.go.jp/ からAPIキーを取得し、"
            ".env ファイルに設定してください。\n"
        )

    client = EdinetClient(api_key=api_key)

    # ------------------------------------------------------------------ #
    # 1. Resolve EDINET code                                               #
    # ------------------------------------------------------------------ #
    if args.edinet_code:
        company_info = {
            "edinetCode": args.edinet_code,
            "filerName": company_name,
            "periodEnd": "",
        }
        print(f"\nEDINETコード {args.edinet_code} を使用します")
    else:
        print(f"\n「{company_name}」の企業情報を EDINET で検索します...")
        company_info = client.search_company(company_name)
        if not company_info:
            print(
                f"\n「{company_name}」に該当する企業が見つかりませんでした。\n"
                "ヒント:\n"
                "  • 正式な会社名（法人格を含む）で検索してください（例: トヨタ自動車株式会社）\n"
                "  • --edinet-code オプションで直接コードを指定することもできます"
            )
            sys.exit(1)
        print(f"  見つかりました: {company_info['filerName']} (コード: {company_info['edinetCode']})")

    edinet_code = company_info["edinetCode"]
    display_name = company_info.get("filerName", company_name)

    # ------------------------------------------------------------------ #
    # 2. Fetch annual report list                                          #
    # ------------------------------------------------------------------ #
    print(f"\n有価証券報告書（過去{args.years}期）を検索しています...")
    reports = client.get_annual_reports(edinet_code, years=args.years)

    if not reports:
        print("有価証券報告書が見つかりませんでした")
        sys.exit(1)

    print(f"  {len(reports)}件 見つかりました")

    # ------------------------------------------------------------------ #
    # 3. Download & parse XBRL                                             #
    # ------------------------------------------------------------------ #
    parser = XBRLParser()
    financial_data: list[dict] = []

    print("\nXBRL財務データを取得・解析しています...")
    for report in reports:
        period = report.get("periodEnd", "不明")[:10]
        print(f"  [{period}] ダウンロード中...", end="", flush=True)
        zip_bytes = client.download_xbrl(report["docID"])

        if zip_bytes is None:
            print(" スキップ（ダウンロード失敗）")
            continue

        print(" 解析中...", end="", flush=True)
        data = parser.parse(zip_bytes, period)

        if data is None:
            print(" スキップ（解析失敗）")
            continue

        print(f" 完了 ({len(data) - 1}項目取得)")
        financial_data.append(data)

    if not financial_data:
        print("\n財務データの取得に失敗しました。APIキーや企業名をご確認ください。")
        sys.exit(1)

    if len(financial_data) < 2:
        print(
            f"\n警告: 取得できた期数が{len(financial_data)}期のみです。"
            "トレンド分析には2期以上必要です。"
        )

    # Sort ascending by period for charts / trend calculations
    financial_data.sort(key=lambda x: x["period"])

    # ------------------------------------------------------------------ #
    # 4. Analyze                                                           #
    # ------------------------------------------------------------------ #
    print("\n財務分析を実行しています...")
    analyzer = FinancialAnalyzer()
    analysis = analyzer.analyze(financial_data)

    # ------------------------------------------------------------------ #
    # 5. Generate report                                                   #
    # ------------------------------------------------------------------ #
    print("HTMLレポートを生成しています...")
    generator = ReportGenerator()
    output_path = generator.generate(display_name, analysis)

    print(f"\nレポートが生成されました: {output_path}")
    print("ブラウザで開いてご確認ください。\n")

    generator.print_summary(display_name, analysis)


if __name__ == "__main__":
    main()
