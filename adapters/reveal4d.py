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
        "TopRail": "RAIL", "SteelBox": "STL", "Wall": "WAL", "RetWall": "WAL", "Drop": "SLB",
        "BlockWall12": "BLK", "BlockWall25": "BLK", "RetainingWall": "WAL", "SoilLayer": "SOIL",
        "Road": "ROAD", "Paving": "PAVE", "Curb": "CURB", "Planter": "PLANT", "Pool": "POOL",
        "Lawn": "LAWN", "Hedge": "HEDGE", "TreeTrunk": "TRUNK", "TreeCanopy": "TREE",
        "ACPLight": "FAC", "ACPDark": "FAC2", "CurtainWall": "WIN", "Shopfront": "WIN", "Coping": "FAC2"}
FLOOR_ALIAS = {"B1": "BM", "TR": "RF"}   # Mall basement / roof top-room shown on the block floor list
LEVEL_MAP = {"FND": "FND", "SOG": "SOG", "BM": "BM", "GF": "GF", "01": "FF", "02": "SF", "03": "TF", "04": "4F",
             "05": "5F", "RF": "RF"}
FLOORS = ["FND", "SOG", "BM", "GF", "FF", "SF", "TF", "4F", "5F", "RF", "EXT"]


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
        if x.get("Area") not in area_code:
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


def _poly_pieces(geom):
    out = []
    for g in getattr(geom, "geoms", [geom]):
        if g.is_empty or g.area < 0.05:
            continue
        out.append(dict(p=[list(c) for c in g.exterior.coords][:-1], holes=[[list(c) for c in h.coords][:-1] for h in g.interiors]))
    return out


def add_site(groups, root):
    """Site works as drawn (SITE.json: roads, paving, kerbs, planters, pools) + soft landscape (landscape.json, ESTIMATE:
    no planting plan issued). The Master has no site activities, so main() links them to its last activity."""
    import math
    f = os.path.join(root, "data", "geometry", "SITE.json")
    if not os.path.exists(f):
        return False
    for e in json.load(open(f, encoding="utf-8"))["elements"]:
        if e["cat"] in TYPE and len(e.get("p", [])) >= 3 and e["z1"] - e["z0"] >= 0.01:
            groups[("SITE", "EXT", e["cat"], None, "SITE")].append(e)
    f = os.path.join(root, "data", "landscape.json")
    if os.path.exists(f):
        L = json.load(open(f, encoding="utf-8"))
        for cat, key, h in (("Lawn", "lawn", 0.08), ("Hedge", "hedge", 1.0)):
            for a in L.get(key, []):
                if len(a["p"]) >= 3:
                    groups[("SITE", "EXT", cat, None, "SITE")].append(
                        dict(p=a["p"], h=a.get("h", []), z0=a["z"], z1=a["z"] + h, est="landscape ESTIMATE"))
        ring = lambda x, y, r, n: [[x + r * math.cos(2 * math.pi * i / n), y + r * math.sin(2 * math.pi * i / n)] for i in range(n)]
        RAD = {"palm": 1.7, "broad": 2.6, "cypress": 0.8}
        for x, y, z, kind, h in L.get("trees", []):
            groups[("SITE", "EXT", "TreeTrunk", None, "SITE")].append(
                dict(p=ring(x, y, 0.16, 6), z0=z, z1=z + h * (0.82 if kind == "palm" else 0.55), est="landscape ESTIMATE"))
            r = RAD.get(kind, 2)
            # canopy as stacked tiers so it reads as a crown, not a drum: (bottom, top, radius factor) of the height
            tiers = {"palm": [(0.80, 0.88, 1.0), (0.88, 0.95, 0.7), (0.95, 1.0, 0.35)],
                     "cypress": [(0.20, 0.45, 1.0), (0.45, 0.70, 0.8), (0.70, 0.88, 0.55), (0.88, 1.0, 0.25)]
                     }.get(kind, [(0.50, 0.62, 0.7), (0.62, 0.82, 1.0), (0.82, 0.93, 0.75), (0.93, 1.0, 0.4)])
            for b0, b1, k in tiers:
                groups[("SITE", "EXT", "TreeCanopy", None, "SITE")].append(
                    dict(p=ring(x, y, r * k, 10), z0=z + h * b0, z1=z + h * b1, est="landscape ESTIMATE"))
    return True


def add_assumed(groups, files):
    for f in files:
        if f.endswith("SUB"):
            _add_assumed(groups, f)


def _add_assumed(groups, sub):
    """ASSUMED geometry for scope the drawings do not carry: basement retaining wall + soil layer under footings."""
    from shapely.geometry import MultiPoint, Polygon
    from shapely.ops import unary_union
    bm = [(k, e) for k, es in groups.items() if k[0] == sub and k[1] == "BM" and k[2] == "Slab" for e in es]
    sog = [e for k, es in groups.items() if k[0] == sub and k[1] == "SOG" and k[2] == "Slab" for e in es]
    if bm:
        allp = unary_union([Polygon(e["p"]).buffer(0.05) for _, e in bm]).buffer(-0.05)
        outer = max(getattr(allp, "geoms", [allp]), key=lambda g: g.area)
        outer = Polygon(outer.exterior)
        ring = outer.difference(outer.buffer(-0.3))
        z0 = max(e["z1"] for e in sog) if sog else min(e["z0"] for _, e in bm) - 3
        z1 = min(e["z0"] for _, e in bm)
        parts = defaultdict(list)
        for k, e in bm:
            parts[(k[3], k[4])].append(Polygon(e["p"]).buffer(0.6))
        for (zn, part), ps in parts.items():
            piece = ring.intersection(unary_union(ps))
            for pc in _poly_pieces(piece):
                groups[("ASSUMED", "BM", "RetainingWall", zn, part)].append(
                    dict(id=f"{part}-RW", lv="BM", cat="RetainingWall", z0=z0, z1=z1, p=pc["p"], est=True,
                         zn=zn, part=part, _orig="RetainingWall"))
    fnd = defaultdict(list)
    for k, es in groups.items():
        if k[0] == sub and k[1] == "FND" and k[2] == "Footing":
            fnd[(k[3], k[4])].extend(es)
    for (zn, part), es in fnd.items():
        hull = MultiPoint([pt for e in es for pt in e["p"]]).convex_hull.buffer(1.0, join_style=2)
        zb = min(e["z0"] for e in es)
        for pc in _poly_pieces(hull):
            groups[("ASSUMED", "FND", "SoilLayer", zn, part)].append(
                dict(id=f"{part}-SOIL", lv="FND", cat="SoilLayer", z0=zb - 0.5, z1=zb, p=pc["p"], est=True,
                     zn=zn, part=part, _orig="SoilLayer"))


def _find(index, head, name):
    """Codes whose activity name (text after the last ' · ') ends with `name`, under the key prefix `head`."""
    out = []
    for k, codes in index.items():
        if k[:-1] == head and k[-1].endswith(name):
            out += codes
    return out


def apply_stages(rules, meta, links, acts):
    """Add the activities each element carries (finishing stages); weights = activity budget share on the element."""
    by_floor, by_bld, by_grp, grp_of = defaultdict(list), defaultdict(list), defaultdict(list), {}
    for x in acts:
        suf = x["name"].split(" · ")[-1].strip()
        if x["zone"] and x.get("group") and x["zone"] != x["group"]:
            grp_of[x["zone"]] = x["group"]
        if x["floor"]:
            by_floor[(x["zone"], x["floor"], suf)].append(x["code"])
        elif x["zone"] == x.get("group"):
            by_grp[(x["zone"], suf)].append(x["code"])
        else:
            by_bld[(x["zone"], suf)].append(x["code"])
    A = {x["code"]: x for x in acts}
    out = {}
    for eid, m in meta.items():
        base = list(links.get(eid, []))
        tasks = list(base)
        part = A[base[0]]["zone"] if base and base[0] in A and A[base[0]]["zone"] != A[base[0]].get("group") else m["part"]
        for r in rules:
            if r["cat"] != m["cat"] or (r.get("lv") and r["lv"] != m["lv"]):
                continue
            if r.get("only") and not str(part).startswith(r["only"]):
                continue
            if r.get("replace"):
                tasks = []
            lv = r.get("floor_as") or FLOOR_ALIAS.get(m["lv"], m["lv"])
            for nm in r["names"]:
                g_ = grp_of.get(part)
                if r["floor"] == "same":       # building activity first, else the group's activity on that floor
                    tasks += _find(by_floor, (part, lv), nm) or _find(by_floor, (g_, lv), nm)
                elif r["floor"] == "below":
                    i = FLOORS.index(lv) if lv in FLOORS else -1
                    if 0 <= i < len(FLOORS) - 1:
                        tasks += _find(by_floor, (part, FLOORS[i + 1]), nm) or _find(by_floor, (g_, FLOORS[i + 1]), nm)
                elif r["floor"] == "building":
                    tasks += _find(by_bld, (part,), nm) or _find(by_grp, (g_,), nm)
                elif r["floor"] == "group":
                    tasks += _find(by_grp, (grp_of.get(part),), nm)
        if not any(c in A for c in tasks):
            tasks = base                    # rules found nothing (e.g. Mall): keep the reveal-4d mapping
        seen = []
        for c in tasks:
            if c in A and c not in seen:
                seen.append(c)
        if seen:
            out[eid] = sorted(seen, key=lambda c: (A[c]["start"], A[c]["finish"], c))
    vol_of = defaultdict(float)
    for eid, ts in out.items():
        for c in ts:
            vol_of[c] += max(meta[eid]["vol"], 1e-6)
    res = {}
    for eid, ts in out.items():
        w = [((A[c]["budget"] or {}).get("cost", 0) or 0) * max(meta[eid]["vol"], 1e-6) / vol_of[c] for c in ts]
        res[eid] = dict(tasks=ts, weights=[round(x, 2) for x in w]) if sum(w) > 0 else dict(tasks=ts)
    return res


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--reveal4d", required=True)
    ap.add_argument("--master", required=True)
    ap.add_argument("--sheet", default="Master Schedule موحد")
    ap.add_argument("--area", nargs="+", default=["EB"], help="one or more master Area codes (EB WB NB SB Mall)")
    ap.add_argument("--files", nargs="*")
    ap.add_argument("--out", required=True)
    ap.add_argument("--name", default=None)
    ap.add_argument("--site", action="store_true", help="add site works (SITE.json) and soft landscape (landscape.json)")
    ap.add_argument("--stages", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "reveal_stages.json"),
                    help="extra activities per element (finishing stages); '' to use the reveal-4d mapping only")
    a = ap.parse_args()
    try:
        from shapely.geometry import MultiPoint
    except ImportError:
        sys.exit("pip install shapely --break-system-packages")
    G = os.path.join(a.reveal4d, "data", "geometry")
    FILES = {"Mall": ["MALL"]}
    files = a.files or [os.path.splitext(n)[0] for ar in a.area for n in sorted(os.listdir(G))
                        if n.endswith(".json") and (n[:-5] in FILES.get(ar, []) or
                                                    (ar not in FILES and n.startswith(ar)))]
    G = os.path.join(a.reveal4d, "data", "geometry")
    mapping = json.load(open(os.path.join(a.reveal4d, "data", "mapping.json"), encoding="utf-8"))

    ST = json.load(open(a.stages, encoding="utf-8")) if a.stages else {"rules": [], "split": {}}
    split = ST.get("split", {})
    groups = defaultdict(list)
    dropped = []
    levels, tops = {}, {}
    for f in files:
        g = json.load(open(os.path.join(G, f + ".json"), encoding="utf-8"))
        for lv, L in g.get("levels", {}).items():
            lv = FLOOR_ALIAS.get(lv, lv)
            levels[lv] = min(levels.get(lv, 1e9), L["z"])
            tops[lv] = max(tops.get(lv, -1e9), L["z"] + L.get("h", 0))
        for e in g["elements"]:
            if e["cat"] not in TYPE or len(e.get("p", [])) < 3:
                continue
            if e["z1"] - e["z0"] < 0.01:
                dropped.append(e["id"])
                continue
            cat = e["cat"]
            sp = split.get(cat)
            if sp and e.get(sp["by"]) is not None:
                cat = sp["low"] if e[sp["by"]] < sp["limit"] else sp["high"]
            e["_orig"] = e["cat"]
            groups[(f, e["lv"], cat, e.get("zn"), e.get("part"))].append(e)

    if ST.get("assumed_geometry"):
        add_assumed(groups, files)

    site = a.site and add_site(groups, a.reveal4d)
    if site:
        levels["EXT"], tops["EXT"] = 0, 0.3

    allx = [p[0] for es in groups.values() for e in es for p in e["p"]]
    ally = [p[1] for es in groups.values() for e in es for p in e["p"]]
    ox, oy = min(allx), min(ally)
    R = lambda v: round(v, 2)
    elements, links, orphans, meta = [], {}, [], {}
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
            hs = [[[R(x - ox), R(y - oy)] for x, y in hh] for hh in e.get("h", []) if len(hh) >= 3]
            a_ -= sum(area(hh) for hh in hs)
            parts.append(dict(outer=pts, holes=hs, z0=R(e["z0"]), z1=R(e["z1"])))
            pts_all += pts
            est = est or bool(e.get("est"))
        hull = MultiPoint(pts_all).convex_hull
        outer = [[R(x), R(y)] for x, y in hull.exterior.coords] if hull.geom_type == "Polygon" else parts[0]["outer"]
        cx, cy = hull.centroid.x, hull.centroid.y
        eid = f"{part}-{lv}-{cat}"
        elements.append(dict(id=eid, floor=FLOOR_ALIAS.get(lv, lv), type=typ, mark=cat, axis=part, x=R(cx), y=R(cy),
                             z0=min(p["z0"] for p in parts), z1=max(p["z1"] for p in parts),
                             dims=dict(area_m2=round(ar, 3), pieces=len(parts)), volume_m3=round(vol, 3),
                             source_sheet=f"reveal-4d/{f}" if f != "ASSUMED" else "ASSUMED geometry (adapter)",
                             confidence="ASSUMED" if est or f == "ASSUMED" else "CALC",
                             note="merged pieces (file|level|category|zone|part)", zone=part,
                             geom=dict(outer=outer, holes=[], parts=parts)))
        orig = es[0].get("_orig", cat)
        act = mapping.get("|".join(str(k) for k in (f, lv, orig, zn, part)))
        meta[eid] = dict(cat=cat, lv=lv, part=part, vol=vol)
        if act:
            links[eid] = [act]

    acts, skipped = load_master(a.master, a.sheet, set(a.area))
    codes = {x["code"] for x in acts}
    links = apply_stages(ST.get("rules", []), meta, links, acts)
    if site:   # TEMPORARY: no site activities in the Master -> shown from its last activity (as reveal-4d does)
        last = max(acts, key=lambda x: (x["finish"], x["code"]))["code"]
        for eid, m in meta.items():
            if m["part"] == "SITE":
                links[eid] = dict(tasks=[last])
        print(f"site works + landscape linked to the last Master activity {last} (TEMPORARY, no site activities in the Master)")
    drop = {e["id"] for e in elements if e["confidence"] == "ASSUMED" and e["id"] not in links}
    if drop:   # ASSUMED geometry with no Master activity (e.g. WB earthworks not in the Master): leave it out
        elements = [e for e in elements if e["id"] not in drop]
        print(f"ASSUMED elements left out (no Master activity): {len(drop)}")
    for e in elements:
        if e["id"] not in links:
            orphans.append(e["id"])
    bad = sorted({c for v in links.values() for c in v["tasks"] if c not in codes})
    floors = [f for f in FLOORS if f in levels or any(e["floor"] == f for e in elements)]
    xs = [p[0] for e in elements for p in e["geom"]["outer"]]
    ys = [p[1] for e in elements for p in e["geom"]["outer"]]
    data = dict(levels={f: levels.get(f, 0) for f in floors}, top={f: tops.get(f, 0) for f in floors},
                grid=dict(x=[], xl=[], y=[], yl=[]), elements=elements,
                cfg=dict(name=a.name or ("REVEAL " + " ".join(a.area)), subtitle="REVEAL Residential Compound · Al-Obour",
                         floors=floors, bbox=[min(xs), max(xs), min(ys), max(ys)],
                         qa=dict(origin_shift=[ox, oy], source="reveal-4d " + " ".join(files),
                                 unmapped_elements=orphans, zero_height_pieces_dropped=len(dropped),
                                 assumed_geometry=ST.get("assumed_geometry", {}))))
    os.makedirs(a.out, exist_ok=True)
    json.dump(data, open(os.path.join(a.out, "data.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    json.dump(dict(source=os.path.basename(a.master), area=" ".join(a.area), activities=acts),
              open(os.path.join(a.out, "activities.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    json.dump(links, open(os.path.join(a.out, "links.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    bud = sum((x["budget"] or {}).get("cost", 0) for x in acts)
    print(f"elements {len(elements)} (from {sum(len(v) for v in groups.values())} pieces) | linked {len(links)} | "
          f"unmapped {len(orphans)}" + (f" {orphans[:6]}" if orphans else "") +
          f" | zero-height pieces dropped {len(dropped)}")
    used = {c for v in links.values() for c in v["tasks"]}
    on_model = sum((x["budget"] or {}).get("cost", 0) for x in acts if x["code"] in used)
    print(f"activities on the model {len(used)} | budget on the model {on_model:,.0f}")
    print(f"activities {len(acts)} (skipped superseded/undated {skipped}) | budget {bud:,.0f}")
    if bad:
        print("WARNING: mapping names activities missing from the master:", bad[:8])
    print(f"floors {floors} | origin shift ({ox:.3f}, {oy:.3f})")


if __name__ == "__main__":
    main()
