#!/usr/bin/env python3
"""normalize_csv.py — Normalize messy CSVs and reconcile two data sources.

What it does:
  1. Reads every store CSV in the input folder. Column names, column
     order, date formats and encodings (UTF-8 / UTF-8 BOM / Shift-JIS)
     may all differ between files; a header mapping (column_map.json)
     unifies them into one schema:
         date, store, product, quantity, unit_price, amount, source_file
  2. Flags data-quality problems (quantity x unit_price != amount).
  3. Removes exact duplicate rows and reports them.
  4. Reconciles the cleaned data against a ledger CSV, aggregated by
     (date, store, product): rows missing on either side and
     quantity / amount mismatches are reported.

Outputs (output/):
  normalized_sales.csv / .xlsx   the unified, de-duplicated data
  reconciliation_report.md       human-readable reconciliation report
  (the .xlsx also contains 突合結果 / 重複 / データ品質 sheets)

Usage:
    python3 normalize_csv.py [input_dir] [ledger_csv]
    (defaults: sample_data/  sample_data/ledger.csv)
"""
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%b %d, %Y", "%Y-%m-%d %H:%M:%S")
CANONICAL = ["date", "store", "product", "quantity", "unit_price", "amount"]


def read_text_auto(path):
    """Decode a CSV file, trying UTF-8 (with/without BOM) then Shift-JIS."""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "shift_jis", "cp932"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Could not decode {path}")


def parse_date(text):
    text = text.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Unknown date format: {text!r}")


def parse_number(text):
    return int(re.sub(r"[^0-9-]", "", str(text)) or 0)


def load_rows(path, column_map):
    text, enc = read_text_auto(path)
    reader = csv.DictReader(text.splitlines())
    rows = []
    for raw in reader:
        row = {}
        for header, value in raw.items():
            field = column_map.get((header or "").strip())
            if field:
                row[field] = value.strip() if isinstance(value, str) else value
        if not row.get("date") or not row.get("product"):
            continue
        rows.append({
            "date": parse_date(row["date"]),
            "store": row.get("store", ""),
            "product": row["product"],
            "quantity": parse_number(row.get("quantity", 0)),
            "unit_price": parse_number(row.get("unit_price", 0)),
            "amount": parse_number(row.get("amount", 0)),
            "source_file": path.name,
            "_encoding": enc,
        })
    return rows


def aggregate(rows):
    agg = {}
    for r in rows:
        key = (r["date"], r["store"], r["product"])
        if key not in agg:
            agg[key] = {"quantity": 0, "amount": 0}
        agg[key]["quantity"] += r["quantity"]
        agg[key]["amount"] += r["amount"]
    return agg


def reconcile(store_rows, ledger_rows):
    store_agg, ledger_agg = aggregate(store_rows), aggregate(ledger_rows)
    results = {"match": [], "amount_mismatch": [], "qty_mismatch": [],
               "only_in_store": [], "only_in_ledger": []}
    for key in sorted(set(store_agg) | set(ledger_agg)):
        s, l = store_agg.get(key), ledger_agg.get(key)
        entry = {"date": key[0], "store": key[1], "product": key[2],
                 "store_qty": s["quantity"] if s else None,
                 "ledger_qty": l["quantity"] if l else None,
                 "store_amount": s["amount"] if s else None,
                 "ledger_amount": l["amount"] if l else None}
        if s is None:
            results["only_in_ledger"].append(entry)
        elif l is None:
            results["only_in_store"].append(entry)
        elif s["amount"] != l["amount"]:
            results["amount_mismatch"].append(entry)
        elif s["quantity"] != l["quantity"]:
            results["qty_mismatch"].append(entry)
        else:
            results["match"].append(entry)
    return results


HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def add_sheet(wb, title, headers, rows, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def write_outputs(rows, duplicates, quality, results, out_dir):
    out_dir.mkdir(parents=True, exist_ok=True)

    with open(out_dir / "normalized_sales.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(CANONICAL + ["source_file"])
        for r in rows:
            w.writerow([r[k] for k in CANONICAL] + [r["source_file"]])

    wb = Workbook()
    add_sheet(wb, "統合データ",
              ["日付", "店舗", "商品名", "数量", "単価", "金額", "元ファイル"],
              [[r[k] for k in CANONICAL] + [r["source_file"]] for r in rows],
              [11, 9, 24, 7, 9, 10, 20], first=True)
    recon_rows = []
    labels = {"amount_mismatch": "金額不一致", "qty_mismatch": "数量不一致",
              "only_in_store": "台帳に無い", "only_in_ledger": "店舗データに無い"}
    for kind, label in labels.items():
        for e in results[kind]:
            recon_rows.append([label, e["date"], e["store"], e["product"],
                               e["store_qty"], e["ledger_qty"],
                               e["store_amount"], e["ledger_amount"]])
    add_sheet(wb, "突合結果",
              ["種別", "日付", "店舗", "商品名", "数量（店舗）", "数量（台帳）",
               "金額（店舗）", "金額（台帳）"],
              recon_rows, [14, 11, 9, 24, 12, 12, 13, 13])
    add_sheet(wb, "重複",
              ["日付", "店舗", "商品名", "数量", "単価", "金額", "元ファイル"],
              [[d[k] for k in CANONICAL] + [d["source_file"]] for d in duplicates],
              [11, 9, 24, 7, 9, 10, 20])
    add_sheet(wb, "データ品質",
              ["日付", "店舗", "商品名", "数量", "単価", "金額（記載）",
               "金額（計算値）", "元ファイル"],
              [[q[k] for k in ("date", "store", "product", "quantity",
                               "unit_price", "amount")] +
               [q["quantity"] * q["unit_price"], q["source_file"]] for q in quality],
              [11, 9, 24, 7, 9, 13, 13, 20])
    wb.save(out_dir / "normalized_sales.xlsx")

    lines = ["# 突合レポート（サンプル）", "",
             f"- 統合レコード数: {len(rows)} 件（重複 {len(duplicates)} 件を除去）",
             f"- 突合一致: {len(results['match'])} 件",
             f"- 金額不一致: {len(results['amount_mismatch'])} 件",
             f"- 数量不一致: {len(results['qty_mismatch'])} 件",
             f"- 台帳に無い（店舗データのみ）: {len(results['only_in_store'])} 件",
             f"- 店舗データに無い（台帳のみ）: {len(results['only_in_ledger'])} 件",
             f"- データ品質警告（数量×単価≠金額）: {len(quality)} 件", ""]
    for kind, label in labels.items():
        if results[kind]:
            lines += [f"## {label}", "",
                      "| 日付 | 店舗 | 商品名 | 数量（店舗） | 数量（台帳） | 金額（店舗） | 金額（台帳） |",
                      "|---|---|---|---|---|---|---|"]
            for e in results[kind]:
                f_ = lambda v: f"{v:,}" if isinstance(v, int) else "-"
                lines.append(f"| {e['date']} | {e['store']} | {e['product']} | "
                             f"{f_(e['store_qty'])} | {f_(e['ledger_qty'])} | "
                             f"{f_(e['store_amount'])} | {f_(e['ledger_amount'])} |")
            lines.append("")
    if duplicates:
        lines += ["## 除去した重複行", ""]
        for d in duplicates:
            lines.append(f"- {d['date']} / {d['store']} / {d['product']} "
                         f"（{d['source_file']}）")
        lines.append("")
    if quality:
        lines += ["## データ品質警告", ""]
        for q in quality:
            lines.append(f"- {q['date']} / {q['store']} / {q['product']}: "
                         f"{q['quantity']} × {q['unit_price']:,} = "
                         f"{q['quantity'] * q['unit_price']:,} ≠ 記載 {q['amount']:,}")
    (out_dir / "reconciliation_report.md").write_text("\n".join(lines) + "\n",
                                                      encoding="utf-8")


def main():
    base = Path(__file__).parent
    input_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else base / "sample_data"
    ledger_path = Path(sys.argv[2]) if len(sys.argv) > 2 else input_dir / "ledger.csv"
    column_map = json.loads((base / "column_map.json").read_text(encoding="utf-8"))

    store_rows, seen, duplicates = [], set(), []
    for path in sorted(input_dir.glob("*.csv")):
        if path == ledger_path:
            continue
        for row in load_rows(path, column_map):
            key = tuple(row[k] for k in CANONICAL)
            if key in seen:
                duplicates.append(row)
            else:
                seen.add(key)
                store_rows.append(row)
    store_rows.sort(key=lambda r: (r["date"], r["store"], r["product"]))

    quality = [r for r in store_rows
               if r["unit_price"] and r["quantity"] * r["unit_price"] != r["amount"]]
    ledger_rows = load_rows(ledger_path, column_map)
    results = reconcile(store_rows, ledger_rows)
    write_outputs(store_rows, duplicates, quality, results, base / "output")

    print(f"normalized: {len(store_rows)} rows, duplicates removed: {len(duplicates)}, "
          f"quality warnings: {len(quality)}")
    print(f"reconcile: match={len(results['match'])} "
          f"amount_mismatch={len(results['amount_mismatch'])} "
          f"qty_mismatch={len(results['qty_mismatch'])} "
          f"only_in_store={len(results['only_in_store'])} "
          f"only_in_ledger={len(results['only_in_ledger'])}")
    print("Wrote output/normalized_sales.csv, output/normalized_sales.xlsx, "
          "output/reconciliation_report.md")


if __name__ == "__main__":
    main()
