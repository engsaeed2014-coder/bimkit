#!/usr/bin/env python3
"""Adapter: REVEAL repos -> kit inputs (data.json, activities.json, links.json).

  python3 adapters/reveal4d.py --reveal4d ../reveal-4d --master master.xlsx --area EB --out examples/reveal-eb
        [--files EB-SUB EB123 EB456] [--sheet "Master Schedule موحد"]

- Model: reveal-4d data/geometry/<file>.json. Elements of the same file|level|category|zone|part are merged
  into one kit element (the reveal-4d mapping links exactly that key to one Master activity), so a
  6-building block stays a few hundred elements. Volume = plan area x height of each piece (CALC).
- Programme + budget: the approved Master sheet (cached values). Superseded rows are left out.
- Links: reveal-4d data/mapping.json (building|level|category|zone|part -> activity).
Coordinates are moved to a local origin (shift recorded in data.json cfg.qa.origin_shift).
"""
import argparse
import json
import os
import sys
from collections import defaultdict

TYPE = {"Column": "COL", "Beam": "BEAM", "Slab": "SLB", "Footing": "FND", "Stair": "STR", "BlockWall": "BLK",
        "FacadeRAL9010": "FAC", "FacadeRAL7035": "FAC2", "Window": "WIN", "WindowFrame": "WIN", "Railing": "RAIL",
        "TopRail": "RAIL", "SteelBox": "STL", "Wall": "WAL", "RetWall": "WAL", "Drop": "SLB"}
LEVEL_MAP = {"FND": "FND", "SOG": "SOG", "BM": "BM", "GF": "GF", "01": "FF", "02": "SF", "03": "TF", "04": "4F",
             "05": "5F", "RF": "RF"}
FLOORS = ["FND", "SOG", "BM", "GF", "FF", "SF", "TF", "4F", "5F", "RF"]


def area(pts):
    s = 0.0
    for i in range(len(pts)):
        x0, y0 = pts[i - 1]
        x1, y1 = pts[i]
        s += x0 * y1 - x1 * y0
    return abs(s) / 2


def load_master(path, sheet, area_code):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    H = list(next(rows))
    acts, skipped = [], 0
    for r in rows:
        if not r or not r[0]:
            continue
        x = dict(zip(H, r))
        if x.get("Area") != area_code:
            continue
        if x.get("Package") == "Superseded" or not x.get("Start") or not x.get("Finish"):
            skipped += 1
            continue
        part = x.get("Building / Part") or x.get("Group")
        lv = LEVEL_MAP.get(str(x.get("Level") or ""))
        pkg = x.get("Package") or "-"
        name = str(x.get("Activity") or x["Activity ID"]).strip()
        label = f"{part} · {x.get('Level') or ''} · {name}".replace(" ·  · ", " · ")
        b = None
        if x.get("Budget Cost") not in (None, "") or x.get("Qty") not in (None, ""):
            b = dict(qty=x.get("Qty"), unit=x.get("Unit"), rate=x.get("Budget Rate"),
                     cost=float(x.get("Budget Cost") or 0), ref=x.get("BOQ Ref"), source=x.get("Cost Source"),
                     status=x.get("Cost Loading Status"), basis=x.get("Qty Basis"))
        acts.append(dict(code=str(x["Activity ID"]), name=label, name_ar=label, type="task",
                         wbs=f"{x.get('Group') or area_code}.{part}.{pkg}", floor=lv, zone=part,
                         step=pkg.replace(" ", "_"), package=pkg, group=x.get("Group"),
                         start=x["Start"].date().isoformat(), finish=x["Finish"].date().isoformat(),
                         budget=b, note=x.get("Predecessor / Rule")))
    return acts, skipped


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--reveal4d", required=True)
    ap.add_argument("--master", required=True)
    ap.add_argument("--sheet", default="Master Schedule موحد")
    ap.add_argument("--area", default="EB")
    ap.add_argument("--files", nargs="*")
    ap.add_argument("--out", required=True)
    ap.add_argument("--name", default=None)
    a = ap.parse_args()
    try:
        from shapely.geometry import MultiPoint
    except ImportError:
        sys.exit("pip install shapely --break-system-packages")
    files = a.files or [f"{a.area}-SUB", f"{a.area}123", f"{a.area}456"]
    G = os.path.join(a.reveal4d, "data", "geometry")
    mapping = json.load(open(os.path.join(a.reveal4d, "data", "mapping.json"), encoding="utf-8"))

    groups = defaultdict(list)
    levels, tops = {}, {}
    for f in files:
        g = json.load(open(os.path.join(G, f + ".json"), encoding="utf-8"))
        for lv, L in g.get("levels", {}).items():
            levels[lv] = min(levels.get(lv, 1e9), L["z"])
            tops[lv] = max(tops.get(lv, -1e9), L["z"] + L.get("h", 0))
        for e in g["elements"]:
            if e["cat"] not in TYPE or len(e.get("p", [])) < 3:
                continue
            groups[(f, e["lv"], e["cat"], e.get("zn"), e.get("part"))].append(e)

    allx = [p[0] for es in groups.values() for e in es for p in e["p"]]
    ally = [p[1] for es in groups.values() for e in es for p in e["p"]]
    ox, oy = min(allx), min(ally)
    R = lambda v: round(v, 2)
    elements, links, orphans = [], {}, []
    for key, es in sorted(groups.items(), key=lambda kv: (FLOORS.index(kv[0][1]) if kv[0][1] in FLOORS else 99, kv[0])):
        f, lv, cat, zn, part = key
        typ = TYPE[cat]
        if cat == "Beam" and lv == "FND":
            typ = "TIE"
        parts, vol, ar = [], 0.0, 0.0
        pts_all = []
        est = False
        for e in es:
            pts = [[R(x - ox), R(y - oy)] for x, y in e["p"]]
            a_ = area(pts)
            h = max(e["z1"] - e["z0"], 0)
            vol += a_ * h
            ar += a_
            parts.append(dict(outer=pts, holes=[], z0=R(e["z0"]), z1=R(e["z1"])))
            pts_all += pts
            est = est or bool(e.get("est"))
        hull = MultiPoint(pts_all).convex_hull
        outer = [[R(x), R(y)] for x, y in hull.exterior.coords] if hull.geom_type == "Polygon" else parts[0]["outer"]
        cx, cy = hull.centroid.x, hull.centroid.y
        eid = f"{part}-{lv}-{cat}"
        elements.append(dict(id=eid, floor=lv, type=typ, mark=cat, axis=part, x=R(cx), y=R(cy),
                             z0=min(p["z0"] for p in parts), z1=max(p["z1"] for p in parts),
                             dims=dict(area_m2=round(ar, 3), pieces=len(parts)), volume_m3=round(vol, 3),
                             source_sheet=f"reveal-4d/{f}", confidence="ASSUMED" if est else "CALC",
                             note="merged pieces (file|level|category|zone|part)", zone=part,
                             geom=dict(outer=outer, holes=[], parts=parts)))
        act = mapping.get("|".join(str(k) for k in key))
        if act:
            links[eid] = [act]
        else:
            orphans.append(eid)

    acts, skipped = load_master(a.master, a.sheet, a.area)
    codes = {x["code"] for x in acts}
    bad = sorted({v[0] for v in links.values() if v[0] not in codes})
    for k in [k for k, v in links.items() if v[0] not in codes]:
        orphans.append(k)
        del links[k]
    floors = [f for f in FLOORS if f in levels or any(e["floor"] == f for e in elements)]
    xs = [p[0] for e in elements for p in e["geom"]["outer"]]
    ys = [p[1] for e in elements for p in e["geom"]["outer"]]
    data = dict(levels={f: levels.get(f, 0) for f in floors}, top={f: tops.get(f, 0) for f in floors},
                grid=dict(x=[], xl=[], y=[], yl=[]), elements=elements,
                cfg=dict(name=a.name or f"REVEAL {a.area}", subtitle="REVEAL Residential Compound · Al-Obour",
                         floors=floors, bbox=[min(xs), max(xs), min(ys), max(ys)],
                         qa=dict(origin_shift=[ox, oy], source="reveal-4d " + " ".join(files),
                                 unmapped_elements=orphans)))
    os.makedirs(a.out, exist_ok=True)
    json.dump(data, open(os.path.join(a.out, "data.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    json.dump(dict(source=os.path.basename(a.master), area=a.area, activities=acts),
              open(os.path.join(a.out, "activities.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    json.dump(links, open(os.path.join(a.out, "links.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    bud = sum((x["budget"] or {}).get("cost", 0) for x in acts)
    print(f"elements {len(elements)} (from {sum(len(v) for v in groups.values())} pieces) | linked {len(links)} | "
          f"unmapped {len(orphans)}" + (f" {orphans[:6]}" if orphans else ""))
    print(f"activities {len(acts)} (skipped superseded/undated {skipped}) | budget {bud:,.0f}")
    if bad:
        print("WARNING: mapping names activities missing from the master:", bad[:8])
    print(f"floors {floors} | origin shift ({ox:.3f}, {oy:.3f})")


if __name__ == "__main__":
    main()
