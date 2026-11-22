#!/usr/bin/env python3
"""Episode 3 — read the project's COST sheet (any layout) into rates.json.

  python3 boq_rates.py inspect file.xlsx [sheet]
      list the sheets, or show the first rows of one sheet with column letters and a guess
      of the DESCRIPTION / QUANTITY / COST columns. Sell / revenue / client columns are flagged.

  python3 boq_rates.py extract file.xlsx "Sheet" first last DESC QTY COST rates.json [UNIT] [--cost-is-rate]
      copy description, quantity and cost of rows first..last into rates.json.
      COST is the total study cost of the line; budget rate = cost / quantity.
      With --cost-is-rate the COST column already holds a cost rate per unit.

Only descriptions, quantities and costs are copied — never the file name, owner names or sell rates.
"""
import json
import re
import sys

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

SELL = re.compile(r"sell|selling|revenue|client|contract\s*price|tender|بيع|العميل|المالك|ايراد|إيراد", re.I)
DESC = re.compile(r"desc|item|بيان|وصف|البند", re.I)
QTY = re.compile(r"qty|quant|كمي", re.I)
COST = re.compile(r"cost|تكلف|budget|ميزاني", re.I)
UNIT = re.compile(r"^\s*(unit|uom|الوحدة|وحدة)\s*$", re.I)


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def norm_unit(u):
    if u is None:
        return None
    s = str(u).strip().lower().replace("³", "3").replace("²", "2").replace(" ", "")
    if s in ("m3", "cum", "م3", "مكعب", "cu.m", "cu.m."):
        return "m3"
    if s in ("m2", "sqm", "م2", "مسطح", "sq.m", "sq.m."):
        return "m2"
    if s in ("ton", "t", "طن"):
        return "t"
    if s in ("m", "lm", "م.ط", "مط"):
        return "m"
    return s or None


def inspect(path, sheet=None, nrows=40):
    wb = load_workbook(path, data_only=True, read_only=True)
    if not sheet:
        for ws in wb.worksheets:
            print(f"{ws.title!r}: {ws.max_row} rows x {ws.max_column} cols")
        print("\nnext: python3 boq_rates.py inspect", path, '"<sheet>"')
        return
    ws = wb[sheet]
    guess = {}
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), values_only=True))
    for ri, r in enumerate(rows[:60], 1):
        for ci, v in enumerate(r, 1):
            if not isinstance(v, str):
                continue
            L = get_column_letter(ci)
            if SELL.search(v):
                guess.setdefault("SELL (do not use)", set()).add(L)
            elif DESC.search(v):
                guess.setdefault("DESCRIPTION", set()).add(L)
            elif QTY.search(v):
                guess.setdefault("QUANTITY", set()).add(L)
            elif COST.search(v):
                guess.setdefault("COST", set()).add(L)
            elif UNIT.search(v):
                guess.setdefault("UNIT", set()).add(L)
    width = max((len(r) for r in rows), default=0)
    print("row | " + " | ".join(get_column_letter(i + 1) for i in range(width)))
    for ri, r in enumerate(rows[:nrows], 1):
        cells = ["" if v is None else str(v)[:28] for v in r]
        if any(cells):
            print(f"{ri:>3} | " + " | ".join(cells))
    first = last = None
    for ri, r in enumerate(rows, 1):
        nums = sum(1 for v in r if num(v) is not None)
        txt = any(isinstance(v, str) and len(v) > 8 for v in r)
        if nums >= 2 and txt:
            first = first or ri
            last = ri
    print("\nguess:", {k: sorted(v) for k, v in guess.items()}, "| data rows ~", first, "to", last)
    if "SELL (do not use)" in guess:
        print("WARNING: sell / client price columns found — pick the COST column, never these.")


def extract(path, sheet, first, last, dcol, qcol, ccol, out, ucol=None, cost_is_rate=False):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    di, qi, ci = (column_index_from_string(x) - 1 for x in (dcol, qcol, ccol))
    ui = column_index_from_string(ucol) - 1 if ucol else None
    head = " ".join(str(r[ci]) for r in ws.iter_rows(min_row=max(1, first - 3), max_row=first - 1, values_only=True)
                    if ci < len(r) and r[ci] is not None)      # header text above the chosen COST column
    lines, skipped, nocalc = [], 0, 0
    for ri, r in enumerate(ws.iter_rows(min_row=first, max_row=last, values_only=True), first):
        desc = r[di] if di < len(r) else None
        q, c = num(r[qi] if qi < len(r) else None), num(r[ci] if ci < len(r) else None)
        if desc is None or q in (None, 0) or c is None:
            if desc and (r[ci] if ci < len(r) else None) is None and (r[qi] if qi < len(r) else None) is not None:
                nocalc += 1
            skipped += 1
            continue
        cost = c * q if cost_is_rate else c
        lines.append(dict(desc=" ".join(str(desc).split()), qty=q, cost=round(cost, 2),
                          unit=norm_unit(r[ui]) if ui is not None and ui < len(r) else None,
                          rate=round(cost / q, 4), row=ri))
    if SELL.search(head):
        print("WARNING: the COST column header mentions sell / client prices — pick the study-cost column.")
    if nocalc:
        print(f"WARNING: {nocalc} rows have no cached cost value (formulas never calculated). "
              "Open and save the file in Excel / LibreOffice first.")
    json.dump(dict(sheet=sheet, rows=[first, last], columns=dict(desc=dcol, qty=qcol, cost=ccol, unit=ucol),
                   cost_is_rate=cost_is_rate, lines=lines), open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{len(lines)} cost lines -> {out} ({skipped} rows skipped: no description / quantity / cost)")
    for l in lines:
        print(f"  {l['rate']:>12,.2f} /{l['unit'] or '?':<3} {l['desc'][:70]}")


def main(a):
    flags = [x for x in a if x.startswith("--")]
    a = [x for x in a if not x.startswith("--")]
    if not a or a[0] not in ("inspect", "extract"):
        print(__doc__)
        sys.exit(1)
    if a[0] == "inspect":
        inspect(a[1], a[2] if len(a) > 2 else None)
    else:
        if len(a) < 9:
            print(__doc__)
            sys.exit(1)
        extract(a[1], a[2], int(a[3]), int(a[4]), a[5], a[6], a[7], a[8], a[9] if len(a) > 9 else None,
                "--cost-is-rate" in flags)


if __name__ == "__main__":
    main(sys.argv[1:])
