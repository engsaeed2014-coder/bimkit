#!/usr/bin/env python3
"""Fill the kit Update sheet from an actualized REVEAL Master (dates actualized from IR / CPR logs).

  python3 adapters/reveal_progress.py --template out/REVEAL_EB_Update.xlsx \
      --master master_actualized.xlsx --data-date 2026-10-03 --out examples/reveal-eb/REVEAL_EB_Update.xlsx

Rules: finish on/before the data date = complete (100%, actual dates = the actualized dates);
start on/before the data date and finish after it = in progress, % = elapsed working days (ASSUMED);
later = forecast dates from the actualized master. Crews are left as they are in the template.
"""
import argparse
import datetime as dt

from openpyxl import load_workbook


def wd(a, b, off=(4,)):
    n, x = 0, a
    while x <= b:
        if x.weekday() not in off:
            n += 1
        x += dt.timedelta(days=1)
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--master", required=True)
    ap.add_argument("--sheet", default="Master Schedule موحد")
    ap.add_argument("--data-date", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    dd = dt.date.fromisoformat(a.data_date)
    ws = load_workbook(a.master, read_only=True, data_only=True)[a.sheet]
    rows = ws.iter_rows(values_only=True)
    H = list(next(rows))
    M = {}
    for r in rows:
        if r and r[0]:
            x = dict(zip(H, r))
            if x.get("Start") and x.get("Finish"):
                M[str(x["Activity ID"])] = (x["Start"].date(), x["Finish"].date())
    wb = load_workbook(a.template)
    up = wb["Update"]
    up["D2"] = dd
    n = dict(done=0, prog=0, fc=0, missing=0)
    for r in range(6, up.max_row + 1):
        code = up.cell(r, 1).value
        if not code:
            continue
        if code not in M:
            n["missing"] += 1
            continue
        s, f = M[code]
        up.cell(r, 9).value, up.cell(r, 10).value = s, f
        if f <= dd:
            pct = 1.0
            n["done"] += 1
        elif s <= dd:
            pct = round(wd(s, dd) / max(wd(s, f), 1), 2)
            n["prog"] += 1
        else:
            pct = 0.0
            n["fc"] += 1
        up.cell(r, 11).value = pct
    wb.save(a.out)
    print(f"data date {dd} | complete {n['done']} | in progress {n['prog']} (% ASSUMED by elapsed days) | "
          f"forecast {n['fc']} | not in the master {n['missing']} -> {a.out}")


if __name__ == "__main__":
    main()
