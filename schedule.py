#!/usr/bin/env python3
"""Episode 2 — baseline programme from project.json + the 3D model.

  python3 schedule.py project.json out            build programme from the step/logic rules
  python3 schedule.py project.json out --xer F    use an existing XER instead (codes/logic/durations kept)

Writes out/programme.json, <Prefix>_Programme.xlsx, <Prefix>_Baseline_Programme.xer and,
if it does not exist yet, <Prefix>_Update.xlsx (the progress sheet the site fills in).
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (Calendar, cpm, d, die, fname, iso, load_json, load_project, open_ends,
                    ppath, save_json, scaled, task_dates)


# ---------------------------------------------------------------- element zones / groups
def zone_of(el, zones):
    x = el["x"]
    for z in zones:
        lo, hi = z.get("x_min", -1e18), z.get("x_max", 1e18)
        ylo, yhi = z.get("y_min", -1e18), z.get("y_max", 1e18)
        if lo <= x < hi and ylo <= el["y"] < yhi:   # upper bound exclusive
            return z["code"]
    return None


def el_group(el, S):
    return S["groups"].get(el["type"])


def el_floor(el, S):
    return S["fnd_floor"] if el_group(el, S) == "FND" else el["floor"]


# ---------------------------------------------------------------- generator
def generate(P):
    S = P["schedule"]
    K = S.get("duration_scale", 1.0)
    rnd = S.get("rounding", "ceil")
    L = S["logic"]
    steps = S["steps"]
    zones = S["zones"]
    crane_on = S.get("crane", {}).get("enabled", False)

    def dur(step):
        st = steps[step]
        return st["dur"] if "dur" in st else scaled(st["base"], K, rnd)

    acts = []
    names = {z["code"]: z for z in zones}
    fl_ar = S.get("floor_names_ar", {})

    def add(code, step, typ, wbs, floor=None, zone=None, preds=(), name=None, name_ar=None, part=None):
        st = steps.get(step, {}) if step else {}
        a = {
            "code": code, "type": typ, "wbs": wbs, "floor": floor, "wing": zone, "step": step,
            "part": part if part is not None else st.get("part"),
            "dur": 0 if typ != "task" else dur(step),
            "crew": dict(st.get("crew", {})) if typ == "task" else {},
            "preds": [dict(task=t, type=ty, lag=lg) for t, ty, lg in preds],
            "name": name or st.get("name"), "name_ar": name_ar or st.get("name_ar"),
        }
        acts.append(a)
        return code

    ntp = add("A1000", None, "start_milestone", "1.1.1", name="Notice to proceed", name_ar="إشعار بدء العمل")
    mob = add("A1010", "MOB", "task", "1.1.1", preds=[(ntp, "FS", 0)])
    crn = add("A1020", "CRN", "task", "1.1.2", preds=[(mob, "FS", 0)]) if crane_on else None

    # raft per zone
    n = 1100
    fnd_last = {}
    seq = S["sequence"]
    for zi, z in enumerate(zones):
        prev = None
        first0 = None if zi == 0 else f"A{1100}"
        for si, step in enumerate(seq["FND"]):
            code = f"A{n}"; n += 10
            pr = []
            if si == 0:
                pr = [(mob, "FS", 0)] if zi == 0 else [(first0, "SS", L["zone_ss_lag"])]
            else:
                pr = [(prev, "FS", 0)]
            if si == len(seq["FND"]) - 1 and zi > 0:
                pr.append((fnd_last[zones[zi - 1]["code"]], "FS", L["raft_pour_gap"]))
            st = steps[step]
            add(code, step, "task", f"1.2.{zi + 1}", S["fnd_floor"], z["code"], pr,
                f"{st['name']} - {z['name']}", f"{st['name_ar']} - {z['name_ar']}")
            prev = code
        fnd_last[z["code"]] = prev

    # floors
    n = 2000
    last_of_zone = dict(fnd_last)        # last activity per zone so far
    on_raft = {z["code"]: True for z in zones}
    for fi, fl in enumerate(S["floors"]):
        parts = S.get("floor_parts", {}).get(fl, S.get("default_floor_parts", ["SLB", "VERT"]))
        zone0_first = None
        for zi, z in enumerate(zones):
            prev = None
            for pi, part in enumerate(parts):
                for si, step in enumerate(seq[part]):
                    code = f"A{n}"; n += 10
                    pr = []
                    if si == 0 and pi == 0:
                        if on_raft[z["code"]]:
                            pr.append((last_of_zone[z["code"]], "FS", L["vertical_after_raft_lag"]))
                            if crn:
                                pr.append((crn, "FS", 0))
                        else:
                            pr.append((last_of_zone[z["code"]], "FS", L["slab_after_vertical_lag"] if part == "SLB" else L["vertical_after_slab_lag"]))
                        if zi == 0:
                            zone0_first = code
                        else:
                            pr.append((zone0_first, "SS", L["zone_ss_lag"]))
                    elif si == 0:
                        pr.append((prev, "FS", L["vertical_after_slab_lag"] if part == "VERT" else L["slab_after_vertical_lag"]))
                    else:
                        pr.append((prev, "FS", 0))
                    st = steps[step]
                    add(code, step, "task", f"1.3.{fi + 1}", fl, z["code"], pr,
                        f"{fl} - {z['name']} - {st['name']}",
                        f"{fl_ar.get(fl, fl)} - {z['name_ar']} - {st['name_ar']}")
                    prev = code
            last_of_zone[z["code"]] = prev
            on_raft[z["code"]] = False

    lasts = [last_of_zone[z["code"]] for z in zones]
    closers = []
    if crn:
        closers.append(add("A9000", "DIS", "task", "1.4.1", preds=[(c, "FS", 0) for c in lasts]))
    closers.append(add("A9010", "QA", "task", "1.4.1", preds=[(lasts[-1], "FS", 0)]))
    add("A9020", None, "finish_milestone", "1.4.1", preds=[(c, "FS", 0) for c in closers],
        name="Substantial completion of structure", name_ar="الاستلام الابتدائي للهيكل")
    return acts


# ---------------------------------------------------------------- XER import
def from_xer(P, path):
    from xer import read_xer
    X = read_xer(path)
    hpd = P["calendar"].get("hours_per_day", 8)
    S = P["schedule"]
    wbs = {w["wbs_id"]: w for w in X.get("PROJWBS", [])}

    def wbs_code(wid):
        w = wbs.get(wid)
        return w["wbs_short_name"] if w else ""
    tasks = X["TASK"]
    by_id = {t["task_id"]: t for t in tasks}
    rs = {r["rsrc_id"]: r["rsrc_short_name"] for r in X.get("RSRC", [])}
    crews = {}
    for tr in X.get("TASKRSRC", []):
        crews.setdefault(tr["task_id"], {})[rs.get(tr["rsrc_id"], tr["rsrc_id"])] = float(tr.get("target_qty_per_hr") or 0)
    preds = {}
    for p in X.get("TASKPRED", []):
        preds.setdefault(p["task_id"], []).append(dict(
            task=by_id[p["pred_task_id"]]["task_code"], type=p["pred_type"].replace("PR_", ""),
            lag=int(round(float(p.get("lag_hr_cnt") or 0) / hpd))))
    mapping = S.get("xer_map", {})
    zname = {z["name"].lower(): z["code"] for z in S["zones"]}
    sname = {v["name"].lower(): k for k, v in S["steps"].items()}
    part_of = {s: g for g, ss in S["sequence"].items() for s in ss}

    def guess(name):
        """'GF - Wing A (axes 01-15) - Slab formwork' or 'Raft concrete pour - Wing A ...' -> floor/zone/step."""
        bits = [b.strip() for b in name.split(" - ")]
        low = [b.lower() for b in bits]
        step = next((sname[b] for b in low if b in sname), None)
        zone = next((zname[b] for b in low if b in zname), None)
        floor = next((b for b in bits if b in S["floors"]), None)
        if step and part_of.get(step) == "FND":
            floor = S["fnd_floor"]
        if step and zone and floor:
            st = S["steps"][step]
            return dict(floor=floor, zone=zone, step=step, part=st.get("part"))
        if step and not part_of.get(step):
            return dict(step=step)
        return {}

    acts = []
    for t in tasks:
        typ = {"TT_Mile": "start_milestone", "TT_FinMile": "finish_milestone"}.get(t["task_type"], "task")
        m = mapping.get(t["task_code"]) or guess(t["task_name"])
        if m.get("floor") and "name_ar" not in m:
            st = S["steps"][m["step"]]
            zz = next(z for z in S["zones"] if z["code"] == m["zone"])
            fl = S.get("floor_names_ar", {}).get(m["floor"], m["floor"])
            m["name_ar"] = (f"{st['name_ar']} - {zz['name_ar']}" if m["floor"] == S["fnd_floor"]
                            else f"{fl} - {zz['name_ar']} - {st['name_ar']}")
        elif m.get("step") and "name_ar" not in m:
            m["name_ar"] = S["steps"][m["step"]]["name_ar"]
        elif typ != "task" and "name_ar" not in m:
            m["name_ar"] = {"start_milestone": "إشعار بدء العمل", "finish_milestone": "الاستلام الابتدائي للهيكل"}[typ] \
                if t["task_name"] in ("Notice to proceed", "Substantial completion of structure") else t["task_name"]
        acts.append(dict(code=t["task_code"], type=typ, wbs=wbs_code(t["wbs_id"]), floor=m.get("floor"),
                         wing=m.get("zone"), step=m.get("step"), part=m.get("part"),
                         dur=0 if typ != "task" else int(round(float(t.get("target_drtn_hr_cnt") or 0) / hpd)),
                         crew={k: (int(v) if v == int(v) else v) for k, v in crews.get(t["task_id"], {}).items()},
                         preds=preds.get(t["task_id"], []), name=t["task_name"], name_ar=m.get("name_ar", t["task_name"])))
    unmapped = [a["code"] for a in acts if a["type"] == "task" and not a["step"]]
    if unmapped:
        print(f"WARNING: {len(unmapped)} XER activities have no schedule.xer_map entry "
              f"(floor/zone/step) and will not drive model elements: {', '.join(unmapped[:12])}...")
    return acts


# ---------------------------------------------------------------- supplied programme
def from_activities(P, path):
    """Activities supplied as JSON (an approved master converted by an adapter).

    Each record: code, name, name_ar, type (task / start_milestone / finish_milestone), wbs, floor, zone,
    step, start, finish (YYYY-MM-DD), optional dur, preds [{task,type,lag}], crew {trade: n},
    budget {qty, unit, rate, cost, ref, source, status}. Dates are kept as given (no CPM) unless
    schedule.recalc_supplied is true and logic is present.
    """
    src = load_json(path)
    recs = src["activities"] if isinstance(src, dict) else src
    acts = []
    for r in recs:
        typ = r.get("type", "task")
        a = dict(code=str(r["code"]), type=typ, wbs=r.get("wbs") or "", floor=r.get("floor"), wing=r.get("zone"),
                 step=r.get("step"), part=r.get("part"), dur=r.get("dur"), crew=dict(r.get("crew") or {}),
                 preds=list(r.get("preds") or []), name=r.get("name") or r["code"],
                 name_ar=r.get("name_ar") or r.get("name") or r["code"],
                 given_start=r.get("start"), given_finish=r.get("finish"))
        if r.get("budget"):
            a["budget"] = r["budget"]
        for k in ("group", "package", "note"):
            if r.get(k) is not None:
                a[k] = r[k]
        acts.append(a)
    return acts


def links_from_file(acts, data, path):
    """Element -> activity links supplied by an adapter.

    {element_id: code} | {element_id: [codes]} | {element_id: {"tasks": [codes], "weights": [w, ...]}}.
    Weights are the share of the element's progress carried by each activity (equal when omitted).
    """
    raw = load_json(path)
    A = {a["code"]: a for a in acts}
    links, orphans_el, missing = {}, [], set()
    for el in data["elements"]:
        v = raw.get(el["id"])
        wts = None
        if isinstance(v, dict):
            tasks, wts = list(v.get("tasks") or []), v.get("weights")
        else:
            tasks = [v] if isinstance(v, str) else list(v or [])
        bad = [c for c in tasks if c not in A]
        missing.update(bad)
        keep = [i for i, c in enumerate(tasks) if c in A]
        tasks = [tasks[i] for i in keep]
        if not tasks:
            orphans_el.append(el["id"])
            continue
        if wts:
            wts = [float(wts[i]) for i in keep]
            s = sum(wts)
            wts = [w / s for w in wts] if s > 0 else None
        rec = dict(first=tasks[0], pour=tasks[-1], tasks=tasks, group=f"W{len(tasks)}",
                   zone=el.get("zone") or A[tasks[0]].get("wing"))
        if wts:
            rec["weights"] = [round(w, 6) for w in wts]
            rec["group"] = "E:" + el["id"]
        links[el["id"]] = rec
    used = {t for l in links.values() for t in l["tasks"]}
    orphans_act = [a["code"] for a in acts if a["type"] == "task" and a["code"] not in used]
    if missing:
        print(f"WARNING: links name {len(missing)} codes that are not activities: {sorted(missing)[:8]}")
    return links, orphans_el, orphans_act


# ---------------------------------------------------------------- links
def link_elements(P, acts, data):
    S = P["schedule"]
    idx = {}
    for a in acts:
        if a["type"] == "task" and a["floor"] and a["wing"] and a["step"]:
            idx.setdefault((a["floor"], a["wing"]), []).append(a)
    part_of = {s: g for g, ss in S["sequence"].items() for s in ss}
    links, orphans_el = {}, []
    for el in data["elements"]:
        g = el_group(el, S)
        z = zone_of(el, S["zones"])
        fl = el_floor(el, S)
        tasks = [a["code"] for a in idx.get((fl, z), []) if part_of.get(a["step"]) == g]
        if not tasks:
            orphans_el.append(el["id"])
            continue
        links[el["id"]] = dict(first=tasks[0], pour=tasks[-1], tasks=tasks, group=g, zone=z)
    used = {t for l in links.values() for t in l["tasks"]}
    orphans_act = [a["code"] for a in acts if a["type"] == "task" and a["floor"] and a["code"] not in used]
    return links, orphans_el, orphans_act


# ---------------------------------------------------------------- main
def build(P, acts, data, links_path=None):
    S = P["schedule"]
    C = P["calendar"]
    supplied = any("given_start" in a for a in acts)
    start = S.get("start")
    if supplied:
        given = [d(a["given_start"]) for a in acts if a.get("given_start")]
        start = min([d(start)] + given if start else given).isoformat()
    cal = Calendar(start, C.get("off_weekdays", [4]), C.get("holidays", []), C.get("name"))
    if supplied and not (S.get("recalc_supplied") and any(a["preds"] for a in acts)):
        for a in acts:
            if not a.get("given_start") or not a.get("given_finish"):
                raise ValueError(f"{a['code']}: supplied activity without start/finish")
            a["es"] = cal.index(a["given_start"])
            fi = max(cal.index(a["given_finish"], "prev"), a["es"])
            a["ef"] = a["es"] if a["type"] != "task" else fi + 1
            a["dur"] = a["ef"] - a["es"]
            a["ls"] = a["lf"] = a["tf"] = None
    else:
        cpm(acts)
    for a in acts:
        bs, bf = task_dates(cal, a, a["es"], a["ef"])
        a["baseline_start"], a["baseline_finish"] = iso(bs), iso(bf)
        a["total_float_days"] = a["tf"]
        a.pop("given_start", None)
        a.pop("given_finish", None)
    if links_path:
        links, oe, oa = links_from_file(acts, data, links_path)
    else:
        links, oe, oa = link_elements(P, acts, data)
    return cal, links, oe, oa


def checks(P, acts, links, oe, oa, data):
    msgs = []
    supplied = all(a.get("tf") is None for a in acts)
    if supplied:
        msgs.append("supplied programme: dates kept as given, no CPM logic (float not calculated)")
    npred, nsucc = open_ends(acts)
    if (npred or nsucc) and not supplied:
        msgs.append(f"open ends: no predecessor {npred}, no successor {nsucc}")
    if any(p.get("lag", 0) < 0 for a in acts for p in a["preds"]):
        msgs.append("negative lags present")
    mx = P["schedule"].get("max_activity_days", 20)
    long = [a["code"] for a in acts if a["dur"] > mx]
    if long and not supplied:
        msgs.append(f"activities longer than {mx} working days: {long}")
    if oe:
        msgs.append(f"{len(oe)} model elements without an activity: {oe[:10]}")
    if oa:
        msgs.append(f"{len(oa)} floor activities without model elements: {oa[:10]}")
    starts = [a for a in acts if a["type"] == "start_milestone"]
    fins = [a for a in acts if a["type"] == "finish_milestone"]
    if (len(starts) != 1 or len(fins) != 1) and not supplied:
        msgs.append(f"expected one start and one finish milestone, got {len(starts)}/{len(fins)}")
    return msgs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("out")
    ap.add_argument("--xer", help="existing baseline XER to use instead of the generator")
    ap.add_argument("--activities", help="supplied programme JSON (approved master via an adapter); dates kept")
    ap.add_argument("--links", help="element -> activity links JSON (from the adapter)")
    ap.add_argument("--force-update-sheet", action="store_true", help="overwrite an existing Update sheet")
    a = ap.parse_args()
    P = load_project(a.project)
    data = load_json(ppath(P, P["inputs"]["data"]))
    acts_path = a.activities or ppath(P, P["inputs"].get("activities"))
    links_path = a.links or ppath(P, P["inputs"].get("links"))
    if a.xer:
        acts = from_xer(P, a.xer)
    elif acts_path:
        acts = from_activities(P, acts_path)
    else:
        acts = generate(P)
    try:
        cal, links, oe, oa = build(P, acts, data, links_path)
    except ValueError as e:
        die(str(e))
    os.makedirs(a.out, exist_ok=True)
    prog = dict(project=P["project"]["name"], calendar=cal.name, start=cal.start.isoformat(),
                contract_end=P["schedule"].get("contract_end"), duration_scale=P["schedule"].get("duration_scale", 1),
                activities=acts, elements=links, orphans=dict(elements=oe, activities=oa))
    save_json(prog, os.path.join(a.out, "programme.json"))

    import progxlsx
    import xer
    xer.write_xer(P, acts, cal, os.path.join(a.out, fname(P, "Baseline_Programme.xer")))
    up = os.path.join(a.out, fname(P, "Update.xlsx"))
    if a.force_update_sheet or not os.path.exists(up):
        progxlsx.write_update_template(P, acts, cal, up)
        print("Update sheet:", up)
    feed = None
    fp = os.path.join(a.out, "feed.json")
    if os.path.exists(fp):
        feed = load_json(fp)
    progxlsx.write_programme(P, acts, cal, links, data, os.path.join(a.out, fname(P, "Programme.xlsx")), feed)

    fin = max(a_["baseline_finish"] for a_ in acts)
    crit = sum(1 for x in acts if x["tf"] is not None and x["tf"] <= 0)
    print(f"activities {len(acts)} | baseline finish {fin} | contract end {P['schedule'].get('contract_end')} | critical {crit}")
    print(f"elements linked {len(links)}/{len(data['elements'])}")
    for m in checks(P, acts, links, oe, oa, data):
        print("CHECK:", m)


if __name__ == "__main__":
    main()
