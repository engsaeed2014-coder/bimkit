#!/usr/bin/env python3
"""Episode 3 — cost on the model (5D).

  python3 cost.py project.json out [--ac Cost_Report.xlsx] [--data-date YYYY-MM-DD]

Needs out/feed.json (update.py) and the rates file named in project.json (boq_rates.py).
Budget rate = the project's own cost-sheet rate (cost / quantity). EV = budget rate x executed
quantity. No client / sell rates anywhere. AC comes from the Cost Report (column O) when given,
otherwise from cost.ac_factors (ASSUMED), otherwise AC = EV.
Writes <Prefix>_Budget_Baseline.xlsx, <Prefix>_Cost_Report.xlsx and out/cost.json.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (Calendar, DATE_FMT, MONEY_FMT, d, die, fname, head_row, iso, load_json, load_project, ppath,
                    save_json, styles)


# ---------------------------------------------------------------- rules
def _norm(s):
    return " ".join(str(s).lower().split())


def find_line(lines, prefix):
    p = _norm(prefix)
    hits = [l for l in lines if _norm(l["desc"]).startswith(p)]
    if not hits:
        hits = [l for l in lines if p in _norm(l["desc"])]
    if len(hits) > 1:
        exact = [l for l in hits if _norm(l["desc"]) == p]
        if len(exact) == 1:
            return exact[0]
        raise ValueError(f"rate '{prefix}' matches {len(hits)} cost lines — make the prefix longer: "
                         + " | ".join(h["desc"][:50] for h in hits))
    return hits[0] if hits else None


def _match(where, el, group):
    for k, v in where.items():
        val = group if k == "group" else el.get(k)
        vs = v if isinstance(v, list) else [v]
        if val not in vs:
            return False
    return True


def _thk_key(t):
    return f"{float(t):g}"


def price(P, data, lines):
    """Element budget: returns {id: dict(bc, label, parts)} and problems list."""
    S = P["schedule"]
    groups = S["groups"]
    rules = P["cost"]["rules"]
    out, unmatched, missing = {}, [], []
    used = {}
    for el in data["elements"]:
        g = groups.get(el["type"])
        rule = next((r for r in rules if _match(r.get("where", {}), el, g)), None)
        if not rule:
            unmatched.append(el["id"])
            continue
        parts = []
        spec = rule["rate"]
        main_basis = rule.get("basis", "volume_m3")
        if isinstance(spec, dict) and spec.get("by") == "t_m":
            t = el.get("dims", {}).get("t_m")
            key = _thk_key(t) if t is not None else None
            table = {_thk_key(k): v for k, v in spec["table"].items()}
            if key not in table:
                missing.append(f"{el['id']}: no cost line for thickness {t} m")
                continue
            prefixes = [table[key]]
        else:
            prefixes = spec if isinstance(spec, list) else [spec]
        items = [(px, main_basis) for px in prefixes] + [(x["rate"], x.get("basis", "volume_m3")) for x in rule.get("also", [])]
        ok = True
        for px, basis in items:
            line = find_line(lines, px)
            if not line:
                missing.append(f"{el['id']}: cost line '{px}' not in the cost sheet")
                ok = False
                continue
            qty = el["volume_m3"] if basis == "volume_m3" else el.get("dims", {}).get(basis)
            if qty is None:
                missing.append(f"{el['id']}: no {basis} in the model")
                ok = False
                continue
            want = "m3" if basis == "volume_m3" else ("m2" if basis == "area_m2" else None)
            if line.get("unit") and want and line["unit"] != want:
                missing.append(f"unit check: '{line['desc'][:40]}' is per {line['unit']} but priced on {basis}")
            rate = line["cost"] / line["qty"]
            parts.append(dict(desc=line["desc"], rate=rate, basis=basis, qty=qty, cost=rate * qty))
            used[line["desc"]] = line
        if not ok:
            continue
        t = el.get("dims", {}).get("t_m")
        label = rule.get("label", parts[0]["desc"] if parts else el["type"])
        label = label.replace("{t_mm}", str(int(round(t * 1000))) if t else "")
        out[el["id"]] = dict(bc=round(sum(p["cost"] for p in parts), 2), label=label, parts=parts)
    return out, unmatched, sorted(set(missing)), used


# ---------------------------------------------------------------- activities
def frac(cal, s, f, dd, typ):
    if typ != "task":
        return 0.0
    s, f = d(s), d(f)
    if dd < s:
        return 0.0
    if dd >= f:
        return 1.0
    return cal.networkdays(s, dd) / cal.networkdays(s, f)


def load_ac(path, codes):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    if "Activity Cost" not in wb.sheetnames:
        die("sheet 'Activity Cost' not found in " + path)
    ws = wb["Activity Cost"]
    ac = {}
    for r in ws.iter_rows(min_row=6, values_only=True):
        if r[0] in codes and r[14] not in (None, ""):
            ac[r[0]] = float(r[14])
    return d(ws["D2"].value), ac


def run(P, feed, data, rates, ac_file=None, data_date=None):
    S, C = P["schedule"], P["calendar"]
    cal = Calendar(S["start"], C.get("off_weekdays", [4]), C.get("holidays", []), C.get("name"))
    lines = rates["lines"]
    E, unmatched, missing, used = price(P, data, lines)
    W = S["weights"]
    seq = S["sequence"]
    pos = {s: (g, i) for g, ss in seq.items() for i, s in enumerate(ss)}
    acts = feed["activities"]
    bc = {a["code"]: 0.0 for a in acts}
    q = {a["code"]: 0.0 for a in acts}
    vol = {e["id"]: e["volume_m3"] for e in data["elements"]}
    for eid, l in feed["elements"].items():
        if eid not in E:
            continue
        for t in l["tasks"]:
            st = next(a["step"] for a in acts if a["code"] == t)
            g, i = pos[st]
            w = W[g][i]
            bc[t] += E[eid]["bc"] * w
            q[t] += vol[eid] * w
    dd = d(data_date or feed["data_date"])
    ac_in, ac_src = {}, None
    if ac_file:
        add, ac_in = load_ac(ac_file, set(bc))
        dd = d(data_date) if data_date else (add or dd)
        ac_src = "Cost Report"
    facs = P["cost"].get("ac_factors") or {}
    rows = []
    for a in acts:
        c = a["code"]
        pv = frac(cal, a["baseline_start"], a["baseline_finish"], dd, a["type"])
        ev = frac(cal, a["upd_start"], a["upd_finish"], dd, a["type"])
        b = round(bc[c], 2)
        af = facs.get(a["step"], 1.0) if facs else 1.0
        if c in ac_in:
            acv = ac_in[c]
        elif ac_file:
            acv = 0.0
        else:
            acv = round(b * af * ev, 2)
        rows.append(dict(code=c, name=a["name_ar"], floor=a["floor"], step=a["step"], type=a["type"],
                         bs=a["baseline_start"], bf=a["baseline_finish"], us=a["upd_start"], uf=a["upd_finish"],
                         bc=b, q=round(q[c], 4), pv_pct=pv, ev_pct=ev, pv=b * pv, ev=b * ev, ac=acv, af=af))
    T = dict(BAC=sum(r["bc"] for r in rows), PV=sum(r["pv"] for r in rows), EV=sum(r["ev"] for r in rows),
             AC=sum(r["ac"] for r in rows))
    T["CPI"] = T["EV"] / T["AC"] if T["AC"] else 0
    T["SPI"] = T["EV"] / T["PV"] if T["PV"] else 0
    T["CV"], T["SV"] = T["EV"] - T["AC"], T["EV"] - T["PV"]
    T["EAC"] = T["AC"] + (T["BAC"] - T["EV"]) / T["CPI"] if T["CPI"] else 0
    T["VAC"] = T["BAC"] - T["EAC"]
    if ac_src:
        T["ac_source"] = ac_src
    elif facs:
        T["ac_source"] = "ASSUMED factors"
    else:
        T["ac_source"] = "AC = budget cost of executed work (CPI 1.0)"
    return dict(cal=cal, E=E, rows=rows, T=T, dd=dd, unmatched=unmatched, missing=missing, used=used, lines=lines,
                ac_assumed=(not ac_src and bool(facs)))


# ---------------------------------------------------------------- workbooks
def write_budget(P, R, data, feed, path):
    from openpyxl import Workbook
    st = styles()
    cur = P["cost"].get("currency", P["project"].get("currency", "EGP"))
    wb = Workbook()
    sm = wb.active
    sm.title = "Summary"
    ur = wb.create_sheet("Unit Rates")
    eb = wb.create_sheet("Element Budget")
    ab = wb.create_sheet("Activity Budget")
    ur["A1"] = f"{P['project']['name']} — Structural Budget (5D input) / ميزانية الهيكل"
    ur["A1"].font = st["title"]
    ur["A2"] = ("Budget rates = study-cost rates READ from the project cost sheet (cost ÷ quantity). "
                "Units: m3 for RC lines, m2 for blinding / waterproofing. Direct cost only.")
    head_row(ur, 4, ["Cost line / بند التكلفة", "Quantity", f"Budget cost ({cur})", f"Budget rate ({cur}/unit)", "Unit", "Used"],
             st, [70, 12, 16, 16, 7, 7])
    for i, l in enumerate(R["lines"], 5):
        for c, v in enumerate([l["desc"], l["qty"], l["cost"], f"=C{i}/B{i}", l.get("unit"),
                               "Y" if l["desc"] in R["used"] else ""], 1):
            cell = ur.cell(i, c, v)
            if c in (3, 4):
                cell.number_format = "#,##0.00" if c == 4 else MONEY_FMT
    head_row(eb, 1, ["Element ID", "Floor", "Type", "Mark", "Volume (m3)", "Cost line", f"Budget rate ({cur}/m3 eq.)",
                     f"Budget cost ({cur})", "Linked activities"], st, [26, 7, 6, 8, 11, 26, 14, 14, 24])
    r = 2
    links = feed["elements"]
    for el in data["elements"]:
        e = R["E"].get(el["id"])
        if not e:
            continue
        row = [el["id"], el["floor"], el["type"], el["mark"], el["volume_m3"], e["label"], f"=H{r}/E{r}", e["bc"],
               ", ".join(links.get(el["id"], {}).get("tasks", []))]
        for c, v in enumerate(row, 1):
            cell = eb.cell(r, c, v)
            if c in (7, 8):
                cell.number_format = MONEY_FMT
        r += 1
    eb.cell(r, 1, "TOTAL").font = st["bold"]
    eb.cell(r, 5, f"=SUM(E2:E{r - 1})")
    eb.cell(r, 8, f"=SUM(H2:H{r - 1})").number_format = MONEY_FMT
    etot = r
    eb.freeze_panes = "B2"
    head_row(ab, 1, ["Activity", "Name", "Floor", f"Budget cost ({cur})", "Budget qty (m3 eq.)", "BL Start", "BL Finish"],
             st, [9, 52, 7, 15, 13, 11, 11])
    r = 2
    for x in R["rows"]:
        if x["type"] != "task":
            continue
        for c, v in enumerate([x["code"], x["name"], x["floor"], x["bc"], x["q"], d(x["bs"]), d(x["bf"])], 1):
            cell = ab.cell(r, c, v)
            if c == 4:
                cell.number_format = MONEY_FMT
            if c in (6, 7):
                cell.number_format = DATE_FMT
        r += 1
    ab.cell(r, 2, "TOTAL").font = st["bold"]
    ab.cell(r, 4, f"=SUM(D2:D{r - 1})").number_format = MONEY_FMT
    atot = r
    W = P["schedule"]["weights"]
    wtxt = "; ".join(f"{g} " + "/".join(str(int(round(w * 100))) for w in ws_) + f" ({'/'.join(P['schedule']['sequence'][g])})"
                     for g, ws_ in W.items())
    rows = [("Budget summary / ملخص الميزانية", None),
            (f"Budget cost — BAC ({cur})", f"='Element Budget'!H{etot}"),
            ("Budget quantity (m3 eq.)", f"='Element Budget'!E{etot}"),
            ("Check: activity budget = element budget", f"=IF(ABS('Activity Budget'!D{atot}-B2)<1,\"OK\",\"CHECK\")"),
            (None, None), (None, None),
            ("EV (cost) = budget rate x executed quantity. No client rates in this workbook.", None),
            (f"Actual cost (AC) is entered in {fname(P, 'Cost_Report.xlsx')} (sheet Activity Cost, column O).", None),
            (f"Stage weights used to load element budget on activities: {wtxt}.", None)]
    for i, (k, v) in enumerate(rows, 1):
        sm.cell(i, 1, k)
        if v:
            sm.cell(i, 2, v).number_format = MONEY_FMT
    sm["A1"].font = st["title"]
    sm.column_dimensions["A"].width = 60
    sm.column_dimensions["B"].width = 18
    wb.save(path)


def write_report(P, R, path):
    from openpyxl import Workbook
    st = styles()
    cur = P["cost"].get("currency", P["project"].get("currency", "EGP"))
    mask = R["cal"].excel_mask()
    wb = Workbook()
    ws = wb.active
    ws.title = "Activity Cost"
    ws["A1"] = f"{P['project']['name']} — Cost Report (5D) / كوست ريبورت"
    ws["A1"].font = st["title"]
    ws["A2"] = "Data date / تاريخ البيانات"
    ws["D2"] = R["dd"]
    ws["D2"].number_format = DATE_FMT
    ws["D2"].fill, ws["D2"].font = st["input_fill"], st["input_font"]
    ac_note = ("والـ AC الحالي تقديري بمعاملات أسعار (ASSUMED)." if R["ac_assumed"]
               else "والـ AC من الموقع.")
    ws["A3"] = ("الأصفر = إدخال (AC = التكلفة الفعلية حتى تاريخه). الأزرق الفاتح = بيانات من البرنامج والميزانية. "
                f"الباقي معادلات. أسعار الميزانية (تكلفة الدراسة) من شيت التكاليف (READ)، {ac_note}")
    hd = ["Code", "Activity / النشاط", "Floor", "Step", "BL Start", "BL Finish", "Upd Start", "Upd Finish",
          f"Budget cost ({cur})", "Budget qty (m3 eq.)", "PV %", "EV %", f"PV ({cur})", f"EV ({cur})",
          f"AC to date ({cur})", "CPI", "CV = EV-AC", "SV = EV-PV", f"Budget rate ({cur}/m3 eq.)"]
    head_row(ws, 5, hd, st, [8, 52, 6, 6, 11, 11, 11, 11] + [14] * 11)
    r = 6
    for x in R["rows"]:
        if x["type"] != "task":
            continue
        pv = f'=IF($D$2<E{r},0,IF($D$2>=F{r},1,NETWORKDAYS.INTL(E{r},$D$2,"{mask}")/NETWORKDAYS.INTL(E{r},F{r},"{mask}")))'
        ev = f'=IF($D$2<G{r},0,IF($D$2>=H{r},1,NETWORKDAYS.INTL(G{r},$D$2,"{mask}")/NETWORKDAYS.INTL(G{r},H{r},"{mask}")))'
        row = [x["code"], x["name"], x["floor"], x["step"], d(x["bs"]), d(x["bf"]), d(x["us"]), d(x["uf"]), x["bc"], x["q"],
               pv, ev, f"=I{r}*K{r}", f"=S{r}*J{r}*L{r}", round(x["ac"], 2), f"=IF(O{r}>0,N{r}/O{r},0)",
               f"=N{r}-O{r}", f"=N{r}-M{r}", f"=IF(J{r}>0,I{r}/J{r},0)"]
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = st["border"]
            if 5 <= c <= 10:
                cell.fill = st["base_fill"]
            if c == 15:
                cell.fill, cell.font = st["input_fill"], st["input_font"]
            if 5 <= c <= 8:
                cell.number_format = DATE_FMT
            elif c in (11, 12):
                cell.number_format = "0%"
            elif c == 16:
                cell.number_format = "0.00"
            elif c >= 9:
                cell.number_format = MONEY_FMT
        r += 1
    last = r - 1
    tot = r
    ws.cell(tot, 2, "TOTAL").font = st["bold"]
    for col in "IJMNOQR":
        ws[f"{col}{tot}"] = f"=SUM({col}6:{col}{last})"
        ws[f"{col}{tot}"].number_format = MONEY_FMT
        ws[f"{col}{tot}"].fill = st["total_fill"]
    ws.freeze_panes = "C6"
    sm = wb.create_sheet("Summary")
    A = "'Activity Cost'!"
    rows = [("Cost summary / ملخص التكلفة", None), (f"BAC — Budget cost ({cur})", f"={A}I{tot}"),
            ("Budget quantity (m3 eq.)", f"={A}J{tot}"), ("PV — planned value to data date", f"={A}M{tot}"),
            ("EV — earned value", f"={A}N{tot}"), ("AC — actual cost to date", f"={A}O{tot}"),
            ("CPI = EV / AC", "=IF(B6>0,B5/B6,0)"), ("SPI = EV / PV", "=IF(B4>0,B5/B4,0)"),
            ("CV = EV − AC", "=B5-B6"), ("SV = EV − PV", "=B5-B4"),
            ("EAC = AC + (BAC − EV)/CPI", "=IF(B7>0,B6+(B2-B5)/B7,0)"), ("VAC = BAC − EAC", "=B2-B11"),
            ("Progress by value (EV/BAC)", "=B5/B2"),
            ("Executed quantity (m3 eq.)", f"=SUMPRODUCT({A}J6:J{last},{A}L6:L{last})"),
            ("AC source", R["T"]["ac_source"])]
    for i, (k, v) in enumerate(rows, 1):
        sm.cell(i, 1, k)
        if v is not None:
            c = sm.cell(i, 2, v)
            c.number_format = "0.00" if i in (7, 8) else ("0.0%" if i == 13 else MONEY_FMT)
    sm["A1"].font = st["title"]
    sm.column_dimensions["A"].width = 34
    sm.column_dimensions["B"].width = 18
    bf = wb.create_sheet("By Floor")
    head_row(bf, 1, ["Floor", "Budget", "PV", "EV", "AC", "CPI"], st, [8, 15, 15, 15, 15, 8])
    floors = [P["schedule"]["fnd_floor"]] + P["schedule"]["floors"]
    for i, f in enumerate(floors, 2):
        bf.cell(i, 1, f)
        for c, col in zip("BCDE", "IMNO"):
            bf[f"{c}{i}"] = f"=SUMIFS({A}{col}$6:{col}${last},{A}$C$6:$C${last},$A{i})"
            bf[f"{c}{i}"].number_format = MONEY_FMT
        bf[f"F{i}"] = f"=IF(E{i}>0,D{i}/E{i},0)"
        bf[f"F{i}"].number_format = "0.00"
    bl = wb.create_sheet("Budget Lines")
    head_row(bl, 1, ["Budget line / بند الميزانية", "Model qty (m3)", f"Budget cost ({cur})", f"Budget rate ({cur}/m3)"],
             st, [30, 14, 16, 16])
    agg = {}
    for e_id, e in R["E"].items():
        a = agg.setdefault(e["label"], [0.0, 0.0])
        a[1] += e["bc"]
        a[0] += sum(p["qty"] for p in e["parts"] if p["basis"] == "volume_m3")
    for i, (k, (qv, cv)) in enumerate(sorted(agg.items()), 2):
        for c, v in enumerate([k, round(qv, 1), round(cv), f"=C{i}/B{i}"], 1):
            cell = bl.cell(i, c, v)
            if c >= 3:
                cell.number_format = MONEY_FMT
    hu = wb.create_sheet("How to use")
    for i, t in enumerate([
        "1) في شيت Activity Cost: عدّل AC (العمود O الأصفر) بالتكلفة الفعلية حتى تاريخه، وتاريخ البيانات D2.",
        "2) EV = سعر الميزانية × الكمية المنفذة فعليًا (مفيش أسعار عميل). PV/EV بيتحسبوا من تواريخ البرنامج (البيز لاين والتحديث). تواريخ التحديث بتتعدل من شيت الـ Update.",
        "3) ابعت الملف كما هو أو استورده في تاب 5D: الصفحة بتقراه مباشرة وتحدّث CPI وEAC والمنحنيات والتدفق النقدي.",
        "4) متغيّرش الأكواد ولا ترتيب الأعمدة ولا أسماء الشيتات. مفيش مصاريف عامة أو أرباح — تكلفة مباشرة فقط."], 1):
        hu.cell(i, 1, t)
    hu.column_dimensions["A"].width = 130
    wb.save(path)


def cost_json(P, R):
    rates = {}
    for e in R["E"].values():
        v = sum(p["qty"] for p in e["parts"] if p["basis"] == "volume_m3")
        if e["label"] not in rates and e["parts"] and e["parts"][0]["basis"] == "volume_m3":
            rates[e["label"]] = round(e["parts"][0]["rate"], 1)
    return dict(
        elements={k: dict(bc=v["bc"], bl=v["label"]) for k, v in R["E"].items()},
        activities={x["code"]: dict(bc=x["bc"], q=x["q"], af=x["af"], ac=round(x["ac"], 2)) for x in R["rows"]},
        rates=rates,
        lines=[dict(desc=l["desc"], qty=l["qty"], cost=l["cost"], unit=l.get("unit")) for l in R["lines"]],
        totals={k: (round(v, 4) if isinstance(v, float) else v) for k, v in R["T"].items()},
        data_date=iso(R["dd"]), ac_assumed=R["ac_assumed"],
        src=("Budget rates = cost rates READ from the cost sheet (client sell rates not used); "
             "EV = budget rate x executed quantity; AC = " +
             ("imagined price factors (ASSUMED)" if R["ac_assumed"] else R["T"]["ac_source"])))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("out")
    ap.add_argument("--ac", help="Cost Report with actual cost in column O")
    ap.add_argument("--data-date")
    a = ap.parse_args()
    P = load_project(a.project)
    fp = os.path.join(a.out, "feed.json")
    if not os.path.exists(fp):
        die("out/feed.json not found — run schedule.py and update.py first")
    feed = load_json(fp)
    data = load_json(ppath(P, P["inputs"]["data"]))
    rp = ppath(P, P["inputs"].get("rates"))
    if not rp or not os.path.exists(rp):
        die("rates file not found — run boq_rates.py extract first and set inputs.rates")
    rates = load_json(rp)
    try:
        R = run(P, feed, data, rates, a.ac or ppath(P, P["inputs"].get("cost_report")), a.data_date)
    except ValueError as e:
        die(str(e))
    write_budget(P, R, data, feed, os.path.join(a.out, fname(P, "Budget_Baseline.xlsx")))
    write_report(P, R, os.path.join(a.out, fname(P, "Cost_Report.xlsx")))
    save_json(cost_json(P, R), os.path.join(a.out, "cost.json"))
    T = R["T"]
    cur = P["cost"].get("currency", "EGP")
    print(f"BAC {T['BAC']:,.2f} {cur} | data date {iso(R['dd'])}")
    print(f"PV {T['PV']:,.0f} | EV {T['EV']:,.0f} | AC {T['AC']:,.0f} ({T['ac_source']})")
    print(f"CPI {T['CPI']:.3f} | SPI {T['SPI']:.3f} | CV {T['CV']:,.0f} | EAC {T['EAC']:,.0f} | VAC {T['VAC']:,.0f}")
    print(f"unmatched elements: {len(R['unmatched'])}" + (f" -> {R['unmatched'][:10]}" if R["unmatched"] else ""))
    print("rates used:")
    for desc, l in R["used"].items():
        print(f"  {l['rate']:>12,.2f}  {desc[:80]}")
    for m in R["missing"]:
        print("MISSING:", m)
    s_act = sum(x["bc"] for x in R["rows"])
    s_el = sum(e["bc"] for e in R["E"].values())
    if abs(s_act - s_el) > 1:
        print(f"CHECK: activity budget {s_act:,.2f} != element budget {s_el:,.2f} (elements without activity?)")


if __name__ == "__main__":
    main()
