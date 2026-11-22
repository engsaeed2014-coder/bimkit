#!/usr/bin/env python3
"""Episode 2 — read the site Update sheet and build the 4D feed.

  python3 update.py project.json out [Update.xlsx] [--data-date YYYY-MM-DD]

Needs out/programme.json (run schedule.py first). Writes out/feed.json and refreshes
<Prefix>_Programme.xlsx with the actual / forecast rows.
Rules: actual dates and % complete come from the sheet (yellow cells). A not-started
activity with empty dates, or an in-progress one with no finish, is forecast from its
predecessors after the data date. Dates on a day off move to the next working day.
"""
import argparse
import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import Calendar, d, die, fname, iso, load_json, load_project, ppath, save_json, task_dates, topo
import progxlsx


def run(P, prog, sheet=None, data_date=None, keep_sheet=False):
    S, C = P["schedule"], P["calendar"]
    cal = Calendar(S["start"], C.get("off_weekdays", [4]), C.get("holidays", []), C.get("name"))
    acts = prog["activities"]
    A = {a["code"]: a for a in acts}
    rows, dd = {}, d(data_date) if data_date else None
    if sheet:
        sdd, rows = progxlsx.read_update(P, sheet)
        dd = dd or sdd
    dd = dd or d(S["start"])
    ddi = cal.index(dd, "prev") if dd >= cal.start else -1
    tr = progxlsx.labour_trades(P)
    warn, resched = [], []
    U = {}
    unknown = [c for c in rows if c not in A]
    if unknown:
        warn.append(f"codes in the sheet but not in the programme (ignored): {unknown[:10]}")

    for c in topo(acts):
        a = A[c]
        r = rows.get(c, {})
        crew_act = dict(a["crew"])
        for k, v in r.get("crew", {}).items():
            crew_act[k] = int(v) if float(v).is_integer() else v
        crew_act = {k: v for k, v in crew_act.items() if v or k not in tr}   # drop empty trades
        pct = r.get("pct", 0.0)
        if a["type"] == "start_milestone":
            s = f = cal.index(d(a["baseline_start"]))
            U[c] = dict(s=s, f=s, pct=1.0 if ddi >= s else 0.0, crew=crew_act)
            continue
        # earliest start allowed by predecessors on the update dates
        es = 0
        for p in a["preds"]:
            b = U[p["task"]]
            lag, t = p.get("lag", 0), p.get("type", "FS")
            bend = b["f"] + 1 if A[p["task"]]["type"] == "task" else b["f"]
            if t == "FS":
                es = max(es, bend + lag)
            elif t == "SS":
                es = max(es, b["s"] + lag)
            elif t == "FF":
                es = max(es, bend + lag - a["dur"])
        if a["type"] == "finish_milestone":
            f = max(U[p["task"]]["f"] for p in a["preds"])
            allp = all(U[p["task"]]["pct"] >= 1 for p in a["preds"])
            U[c] = dict(s=f, f=f, pct=1.0 if allp else 0.0, crew={})
            continue
        st, fi = r.get("start"), r.get("finish")
        si = cal.index(st) if st else None
        fn = cal.index(fi) if fi else None
        if pct >= 1:
            if si is None or fn is None:
                warn.append(f"{c}: 100% but actual start/finish missing — forecast used")
                pct = 0.0
            elif fn > ddi:
                warn.append(f"{c}: 100% with finish after the data date")
        if 0 < pct < 1 and si is None:
            warn.append(f"{c}: in progress without actual start — data date used")
            si = ddi
        if pct >= 1:
            pass
        elif pct > 0:
            rem = max(1, math.ceil(a["dur"] * (1 - pct)))
            if fn is None or fn <= ddi:
                fn = ddi + rem
                resched.append(c)
        else:
            if si is None or fn is None or si <= ddi:
                if si is not None and si <= ddi:
                    warn.append(f"{c}: not started but start on/before the data date — re-forecast from logic")
                si = max(es, ddi + 1)
                fn = si + a["dur"] - 1
                resched.append(c)
            elif si < es and not keep_sheet:
                # sheet forecast is earlier than its predecessors allow: push it, keep its duration
                d_ = fn - si
                si = es
                fn = si + d_
                resched.append(c)
        if fn < si:
            warn.append(f"{c}: finish before start — finish set to start + duration")
            fn = si + a["dur"] - 1
        U[c] = dict(s=si, f=fn, pct=pct, crew=crew_act)

    out_acts = []
    for a in acts:
        u = U[a["code"]]
        if a["type"] == "task":
            us, uf = cal.date(u["s"]), cal.date(u["f"])
        else:
            us = uf = cal.date(u["s"])
        pct = u["pct"]
        started = pct > 0 or (a["type"] == "start_milestone" and pct >= 1)
        o = dict(code=a["code"], name=a["name"], name_ar=a["name_ar"], type=a["type"], wbs=a["wbs"], floor=a["floor"],
                 wing=a["wing"], part=a["part"], step=a["step"], crew=a["crew"],
                 baseline_start=a["baseline_start"], baseline_finish=a["baseline_finish"],
                 actual_start=iso(us) if started and a["type"] == "task" else None,
                 actual_finish=iso(uf) if pct >= 1 and a["type"] == "task" else None,
                 forecast_finish=iso(uf) if a["type"] != "start_milestone" else None,
                 upd_start=iso(us), upd_finish=iso(uf), percent_complete=int(round(pct * 100)),
                 total_float_days=a["tf"], predecessors=a["preds"])
        if u["crew"] != a["crew"] and a["type"] == "task":
            o["crew_actual"] = u["crew"]
        out_acts.append(o)

    # daily series
    pad = P.get("viewer", {}).get("pad_days", 3)
    last = max(max(a["ef"] - 1 for a in acts), max(U[c]["f"] for c in U)) + pad
    n = last + 1
    dates = [iso(cal.date(i)) for i in range(n)]
    z = lambda: [0.0] * n
    pt, at, ut = {k: z() for k in tr}, {k: z() for k in tr}, {k: z() for k in tr}
    crane = z()
    # optional ASSUMED factor on actual head-count until site records exist (number or {step: factor})
    alf = P.get("update", {}).get("actual_labour_factor", 1.0)
    for a in acts:
        if a["type"] != "task":
            continue
        u = U[a["code"]]
        for i in range(a["es"], a["ef"]):
            for k in tr:
                pt[k][i] += a["crew"].get(k, 0)
            crane[i] += a["crew"].get("TWC", 0)
        for i in range(u["s"], u["f"] + 1):
            for k in tr:
                ut[k][i] += u["crew"].get(k, 0)
            if u["pct"] > 0 and i <= ddi:
                f_ = alf.get(a["step"], 1.0) if isinstance(alf, dict) else alf
                for k in tr:
                    at[k][i] += u["crew"].get(k, 0) * f_
                if i <= ddi:
                    for k in tr:
                        ut[k][i] += u["crew"].get(k, 0) * (f_ - 1)
    r1 = lambda v: round(v, 1)
    sumt = lambda T: [r1(sum(T[k][i] for k in tr)) for i in range(n)]
    cut = lambda L: [v if i <= ddi else None for i, v in enumerate(L)]   # no actuals after the data date
    daily = dict(dates=dates, planned_labour=sumt(pt), actual_labour=cut(sumt(at)), planned_crane=[r1(v) for v in crane],
                 upd_labour=sumt(ut), upd_by_trade={k: [r1(v) for v in ut[k]] for k in tr},
                 planned_by_trade={k: [r1(v) for v in pt[k]] for k in tr},
                 actual_by_trade={k: cut([r1(v) for v in at[k]]) for k in tr})

    O = {o["code"]: o for o in out_acts}
    cr = dict(S.get("crane", {}))
    if cr.get("enabled") and "A1020" in O:
        cr.update(erect_upd=O["A1020"]["upd_start"], erect_upd_finish=O["A1020"]["upd_finish"],
                  dismantle_upd_start=O.get("A9000", {}).get("upd_start"),
                  dismantle_upd_finish=O.get("A9000", {}).get("upd_finish"),
                  erect_task="A1020", dismantle_task="A9000" if "A9000" in O else None)
    bfin = max(d(a["baseline_finish"]) for a in acts)
    ufin = max(d(o["upd_finish"]) for o in out_acts)
    feed = dict(project=P["project"]["name"], calendar=cal.name, start=S["start"], contract_end=S.get("contract_end"),
                data_date=iso(dd), duration_scale=S.get("duration_scale", 1),
                duration_scale_note=S.get("duration_scale_note", ""), resources=P["resources"], crane=cr,
                zones=[{k: z_[k] for k in ("code", "name", "name_ar")} for z_ in S["zones"]], floors=S["floors"],
                floor_names_ar={k: v for k, v in S.get("floor_names_ar", {}).items() if k in S["floors"]},
                activities=out_acts, elements=prog["elements"], daily=daily,
                delay_days=(ufin - bfin).days, baseline_finish=iso(bfin), forecast_finish=iso(ufin))
    return feed, cal, warn, resched


def delay_sources(feed, cal):
    """Rank activities by own duration overrun and by late start after predecessors finished."""
    A = {a["code"]: a for a in feed["activities"]}
    out = []
    for a in feed["activities"]:
        if a["type"] != "task":
            continue
        bd = cal.networkdays(a["baseline_start"], a["baseline_finish"])
        ud = cal.networkdays(a["upd_start"], a["upd_finish"])
        ready = None
        for p in a["predecessors"]:
            b = A[p["task"]]
            if p.get("type", "FS") == "FS":
                x = cal.index(d(b["upd_finish"])) + (1 if b["type"] == "task" else 0) + p.get("lag", 0)
            else:
                x = cal.index(d(b["upd_start"])) + p.get("lag", 0)
            ready = x if ready is None else max(ready, x)
        late = max(0, cal.index(d(a["upd_start"])) - ready) if ready is not None and a["actual_start"] else 0
        out.append(dict(code=a["code"], name=a["name_ar"], overrun=max(0, ud - bd), late_start=late))
    out.sort(key=lambda x: -(x["overrun"] + x["late_start"]))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("out")
    ap.add_argument("sheet", nargs="?", help="Update sheet (default: inputs.update in project.json, else out/<Prefix>_Update.xlsx)")
    ap.add_argument("--data-date")
    ap.add_argument("--keep-sheet-forecast", action="store_true",
                    help="use forecast dates from the sheet even when they break the logic")
    a = ap.parse_args()
    P = load_project(a.project)
    pj = os.path.join(a.out, "programme.json")
    if not os.path.exists(pj):
        die("out/programme.json not found — run schedule.py first")
    prog = load_json(pj)
    sheet = a.sheet or ppath(P, P["inputs"].get("update"))
    if not sheet or not os.path.exists(sheet):
        alt = os.path.join(a.out, fname(P, "Update.xlsx"))
        sheet = alt if os.path.exists(alt) else None
    feed, cal, warn, resched = run(P, prog, sheet, a.data_date, a.keep_sheet_forecast)
    save_json(feed, os.path.join(a.out, "feed.json"))
    # refresh programme workbook with actual rows
    from common import cpm
    acts = prog["activities"]
    data = load_json(ppath(P, P["inputs"]["data"]))
    progxlsx.write_programme(P, acts, cal, prog["elements"], data, os.path.join(a.out, fname(P, "Programme.xlsx")), feed)
    done = sum(1 for x in feed["activities"] if x["type"] == "task" and x["percent_complete"] >= 100)
    prog_ = sum(1 for x in feed["activities"] if x["type"] == "task" and 0 < x["percent_complete"] < 100)
    print(f"sheet: {sheet or '(none — no progress)'} | data date {feed['data_date']}")
    print(f"done {done} | in progress {prog_} | forecast finish {feed['forecast_finish']} | "
          f"baseline finish {feed['baseline_finish']} | delay {feed['delay_days']} calendar days")
    if resched:
        print(f"forecast from logic: {len(resched)} activities ({', '.join(resched[:8])}{'...' if len(resched) > 8 else ''})")
    top = [x for x in delay_sources(feed, cal) if x["overrun"] or x["late_start"]][:5]
    for x in top:
        print(f"  delay source {x['code']}: overrun {x['overrun']} wd, late start {x['late_start']} wd")
    for w in warn:
        print("WARNING:", w)


if __name__ == "__main__":
    main()
