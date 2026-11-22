#!/usr/bin/env python3
"""Generate the fictional sample CSVs for the csv-normalizer-reconciler demo.

The three store files deliberately use different column names, column
orders, date formats and text encodings (UTF-8 / Shift-JIS / UTF-8 BOM),
which is exactly the mess this tool cleans up. The ledger file contains
planted discrepancies so the reconciliation report has something to find.

All data is fictional. Usage: python3 make_sample_data.py [output_dir]
"""
import csv
import io
import os
import sys

TOKYO = [
    ["日付", "店舗", "商品名", "数量", "単価", "金額"],
    ["2026/10/01", "東京店", "りんごジュース 1L", "12", "380", "4560"],
    ["2026/10/01", "東京店", "食パン 6枚切", "30", "210", "6300"],
    ["2026/10/01", "東京店", "食パン 6枚切", "30", "210", "6300"],  # exact duplicate
    ["2026/10/02", "東京店", "牛乳 1L", "25", "240", "6000"],
    ["2026/10/02", "東京店", "りんごジュース 1L", "10", "380", "3800"],
]

OSAKA = [
    ["date", "store", "product", "qty", "unit_price", "amount"],
    ["2026-10-01", "大阪店", "りんごジュース 1L", "8", "380", "3040"],
    ["2026-10-01", "大阪店", "牛乳 1L", "20", "240", "4800"],
    ["2026-10-03", "大阪店", "食パン 6枚切", "18", "210", "3780"],
    ["2026-10-03", "大阪店", "卵 10個入", "15", "290", "4350"],
]

FUKUOKA = [  # different names, different order, two date styles
    ["商品", "個数", "売上日", "単価（円）", "合計（円）", "店舗名"],
    ["卵 10個入", "22", "Oct 1, 2026", "290", "6380", "福岡店"],
    ["牛乳 1L", "12", "Oct 2, 2026", "240", "2880", "福岡店"],
    ["りんごジュース 1L", "5", "2026年10月3日", "380", "1900", "福岡店"],
    ["ヨーグルト 4個パック", "10", "Oct 3, 2026", "320", "3000", "福岡店"],  # 10x320 != 3000
]

LEDGER = [
    ["日付", "店舗", "商品名", "数量", "金額"],
    ["2026-10-01", "東京店", "りんごジュース 1L", "12", "4560"],
    ["2026-10-01", "東京店", "食パン 6枚切", "30", "6300"],
    ["2026-10-02", "東京店", "牛乳 1L", "25", "6000"],
    ["2026-10-02", "東京店", "りんごジュース 1L", "10", "3800"],
    ["2026-10-01", "大阪店", "りんごジュース 1L", "8", "3040"],
    ["2026-10-01", "大阪店", "牛乳 1L", "20", "4560"],   # amount differs (store: 4800)
    ["2026-10-03", "大阪店", "食パン 6枚切", "18", "3780"],
    ["2026-10-03", "大阪店", "卵 10個入", "15", "4350"],
    ["2026-10-01", "福岡店", "卵 10個入", "22", "6380"],
    ["2026-10-02", "福岡店", "牛乳 1L", "12", "2880"],
    ["2026-10-03", "福岡店", "ヨーグルト 4個パック", "10", "3000"],
    # 福岡店 2026-10-03 りんごジュース is missing here on purpose.
    ["2026-10-03", "東京店", "納豆 3個パック", "9", "1350"],  # ledger only
]


def write_csv(path, rows, encoding):
    buf = io.StringIO()
    csv.writer(buf).writerows(rows)
    with open(path, "wb") as f:
        f.write(buf.getvalue().encode(encoding))
    print(f"{path} ({encoding}, {len(rows) - 1} data rows)")


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "sample_data"
    os.makedirs(out, exist_ok=True)
    write_csv(os.path.join(out, "sales_tokyo.csv"), TOKYO, "utf-8")
    write_csv(os.path.join(out, "sales_osaka.csv"), OSAKA, "shift_jis")
    write_csv(os.path.join(out, "sales_fukuoka.csv"), FUKUOKA, "utf-8-sig")
    write_csv(os.path.join(out, "ledger.csv"), LEDGER, "utf-8")


if __name__ == "__main__":
    main()
