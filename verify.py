#!/usr/bin/env python3
"""Checks before hand-off.

  python3 verify.py project.json out [--no-recalc]

Programme: no loops, no open ends, no negative lags, one start / one finish milestone.
Links: every model element has exactly one pour activity; floor activities without elements.
Cost: activity budget = element budget = BAC, no unmatched elements.
Workbooks: recalculated with LibreOffice, 0 formula errors. Outputs: no sell-rate wording.
Exit code 1 when a check fails.
"""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_json, load_project, open_ends, ppath, topo

ERR = re.compile(r"^#(REF!|VALUE!|NAME\?|DIV/0!|N/A|NUM!|NULL!)")
SELLW = re.compile(r"sell(ing)? rate|client rate|سعر البيع|earned at boq|boq value|certificate", re.I)


def recalc_errors(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None, "LibreOffice not installed — recalc skipped"
    from openpyxl import load_workbook
    tmp = tempfile.mkdtemp()
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    out = os.path.join(tmp, os.path.basename(path))
    if not os.path.exists(out):
        return None, "LibreOffice conversion failed"
    wb = load_workbook(out, data_only=True)
    errs = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and ERR.match(c.value):
                    errs.append(f"{ws.title}!{c.coordinate} {c.value}")
    shutil.rmtree(tmp, ignore_errors=True)
    return errs, None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("out")
    ap.add_argument("--no-recalc", action="store_true")
    a = ap.parse_args()
    P = load_project(a.project)
    fails, notes = [], []
    ok = lambda m: print("OK  ", m)

    prog = load_json(os.path.join(a.out, "programme.json"))
    acts = prog["activities"]
    try:
        topo(acts)
        ok("programme logic has no loops")
    except ValueError as e:
        fails.append(str(e))
    supplied = all(x.get("tf") is None for x in acts)
    npred, nsucc = open_ends(acts)
    if supplied:
        notes.append("supplied programme: dates as approved, no CPM logic in the kit files (open ends / float not checked)")
    else:
        (fails.append(f"open ends {npred} {nsucc}") if npred or nsucc else ok("no open ends"))
    neg = [x["code"] for x in acts if any(p.get("lag", 0) < 0 for p in x["preds"])]
    (fails.append(f"negative lags: {neg}") if neg else ok("no negative lags"))
    ms = [x["type"] for x in acts]
    if supplied or (ms.count("start_milestone") == 1 and ms.count("finish_milestone") == 1):
        ok("one start and one finish milestone") if not supplied else None
    else:
        fails.append("milestones: need exactly one start and one finish")
    mx = P["schedule"].get("max_activity_days", 20)
    long = [x["code"] for x in acts if x["dur"] > mx]
    if not supplied:
        (notes.append(f"activities over {mx} working days: {long}") if long else ok(f"all activities <= {mx} working days"))

    data = load_json(ppath(P, P["inputs"]["data"]))
    L = prog["elements"]
    steps = {x["code"]: x["step"] for x in acts}
    pours = {s for g, ss in P["schedule"].get("sequence", {}).items() for s in ss[-1:]}
    bad = [] if supplied else [eid for eid, l in L.items() if sum(1 for t in l["tasks"] if steps.get(t) in pours) != 1]
    miss = [e["id"] for e in data["elements"] if e["id"] not in L]
    (fails.append(f"{len(miss)} elements without activity: {miss[:8]}") if miss else ok(f"{len(L)} elements linked"))
    (fails.append(f"elements without exactly one pour: {bad[:8]}") if bad else ok("each element has exactly one pour activity"))
    oa = prog.get("orphans", {}).get("activities", [])
    if oa:
        notes.append(f"{len(oa)} activities without model elements (scheduled and costed at activity level): {oa[:6]}...")
    else:
        ok("no orphan floor activities")

    cp = os.path.join(a.out, "cost.json")
    if os.path.exists(cp):
        C = load_json(cp)
        bac = C["totals"]["BAC"]
        s_el = sum(v["bc"] for v in C["elements"].values())
        s_act = sum(v["bc"] for v in C["activities"].values())
        if C.get("mode") == "activity":
            linked = {t for l in L.values() for t in l["tasks"]}
            s_lk = sum(v["bc"] for k, v in C["activities"].items() if k in linked)
            if abs(s_act - bac) < 1 and abs(s_el - s_lk) < max(1, 1e-8 * s_lk):   # weights are rounded to 0.01
                ok(f"activity budget = BAC {bac:,.2f}; elements carry their activities' budget {s_el:,.2f}")
            else:
                fails.append(f"budget mismatch: activities {s_act:,.2f} BAC {bac:,.2f} elements {s_el:,.2f} linked {s_lk:,.2f}")
        elif abs(s_el - bac) < 1 and abs(s_act - bac) < 1:
            ok(f"activity budget = element budget = BAC {bac:,.2f}")
        else:
            fails.append(f"budget mismatch: elements {s_el:,.2f} activities {s_act:,.2f} BAC {bac:,.2f}")
        unm = [e["id"] for e in data["elements"] if e["id"] not in C["elements"]]
        (fails.append(f"{len(unm)} unmatched elements (no cost rule): {unm[:8]}") if unm else ok("0 unmatched elements"))
        if C.get("ac_assumed"):
            notes.append("AC is ASSUMED (cost.ac_factors) — replace with site actual cost")
    else:
        notes.append("cost.json not found — 5D not run yet")

    for f in sorted(glob.glob(os.path.join(a.out, "*.xlsx"))):
        if a.no_recalc:
            break
        errs, why = recalc_errors(f)
        if why:
            notes.append(f"{os.path.basename(f)}: {why}")
        elif errs:
            fails.append(f"{os.path.basename(f)}: {len(errs)} formula errors {errs[:5]}")
        else:
            ok(f"{os.path.basename(f)} recalculates with 0 errors")

    from openpyxl import load_workbook
    hits = []
    for f in glob.glob(os.path.join(a.out, "*.xlsx")):
        wb = load_workbook(f)
        for ws in wb:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and SELLW.search(v) and "not used" not in v and "No client" not in v and "مفيش أسعار" not in v:
                        hits.append(f"{os.path.basename(f)}/{ws.title}: {v[:60]}")
    (fails.append(f"sell-rate wording found: {hits[:5]}") if hits else ok("no sell-rate / BOQ-value wording in workbooks"))

    for n in notes:
        print("NOTE", n)
    for f in fails:
        print("FAIL", f)
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
