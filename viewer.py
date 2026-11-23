#!/usr/bin/env python3
"""Build the single 3D / 4D / 5D page (out/index.html) from the approved viewer templates.

  python3 viewer.py project.json out

Uses data.json (3D), out/feed.json (4D) and out/cost.json (5D, optional — without it the
5D tab shows zero budget). Publish out/index.html as the project artifact.
"""
import argparse
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import d, die, load_json, load_project, ppath


def rings_for(P, data, off):
    """Outline of the building per zone (union of the slabs on one typical floor)."""
    try:
        from shapely.geometry import Polygon
        from shapely.ops import unary_union
    except ImportError:
        die("shapely missing: pip install shapely --break-system-packages")
    S = P["schedule"]
    V = P.get("viewer", {})
    fl = V.get("rings_floor") or next((f for f in S["floors"] if any(e["floor"] == f and e["type"] == "SLB" for e in data["elements"])), None)
    from schedule import zone_of
    out = []
    for z in S["zones"]:
        polys = []
        for e in data["elements"]:
            ez = e.get("zone") or zone_of(e, [z])
            if e["floor"] != fl or S["groups"].get(e["type"]) != "SLB" or ez != z["code"]:
                continue
            g = e.get("geom") or {}
            rings_ = [pp["outer"] for pp in g.get("parts", [])] or [g.get("outer", [])]
            for ring in rings_:
                if len(ring) >= 3:
                    p = Polygon(ring).buffer(0)
                    if not p.is_empty:
                        polys.append(p)
        if not polys:
            continue
        u = unary_union([p.buffer(0.05, join_style=2) for p in polys]).buffer(-0.05, join_style=2)
        geoms = getattr(u, "geoms", [u])
        big = max(geoms, key=lambda g: g.area)
        big = big.simplify(V.get("rings_simplify", 0.02))
        out.append([[round(x + off[0], 2), round(y + off[1], 2)] for x, y in big.exterior.coords])
    return out


def build_sd(P, data, feed, cost):
    S = P["schedule"]
    V = P.get("viewer", {})
    off = V.get("offset", [0, 0])
    dates = feed["daily"]["dates"]
    di = {x: i for i, x in enumerate(dates)}

    def ix(x):
        if x in di:
            return di[x]
        later = [i for i, y in enumerate(dates) if y > x]
        return later[0] if later else len(dates) - 1

    tr = [k for k, r in P["resources"].items() if r.get("kind") == "labour"]
    C = cost or {}
    CA = C.get("activities", {})
    CE = C.get("elements", {})
    acts = {}
    for a in feed["activities"]:
        if a["type"] == "start_milestone":
            continue
        cr = a.get("crew_actual", a["crew"])
        o = dict(cr=cr, a=[ix(a["baseline_start"]), ix(a["baseline_finish"]), ix(a["upd_start"]), ix(a["upd_finish"])],
                 s=a["step"], f=a["floor"], w=a["wing"], n=a["name_ar"], p=[p["task"] for p in a["predecessors"]],
                 t=a["type"], l=round(sum(v for k, v in cr.items() if k in tr), 1))
        if "crew_actual" in a:
            o["cb"] = a["crew"]
        c = CA.get(a["code"], {})
        o.update(bc=c.get("bc", 0), q=c.get("q", 0), af=c.get("af", 1.0), ac=c.get("ac", 0))
        acts[a["code"]] = o
    links = feed["elements"]
    groups = S["groups"]
    els = []
    for e in data["elements"]:
        g = e.get("geom") or {"outer": [], "holes": []}
        mv = lambda pts: [[round(x + off[0], 2), round(y + off[1], 2)] for x, y in pts]
        ce = CE.get(e["id"], {})
        gg = dict(o=mv(g.get("outer", [])), h=[mv(h) for h in g.get("holes", [])])
        if g.get("parts"):     # merged element: several extrusions shown as one
            gg["p"] = [[mv(pp["outer"]), [mv(h) for h in pp.get("holes", [])], pp["z0"], pp["z1"]] for pp in g["parts"]]
        lk = links.get(e["id"], {})
        els.append(dict(id=e["id"], f=e["floor"], t=e["type"], m=e["mark"], ax=e.get("axis"),
                        x=round(e["x"] + off[0], 2), y=round(e["y"] + off[1], 2), z0=e["z0"], z1=e["z1"], v=e["volume_m3"],
                        g=gg, k=lk.get("tasks", []), c=e.get("confidence"), gr=lk.get("group") or groups.get(e["type"]),
                        bl=ce.get("bl"), bc=ce.get("bc", 0)))
    # weights per element group; supplied links use groups W1..Wn = equal shares
    weights = dict(S.get("weights", {}))
    for l in links.values():
        gname = l.get("group")
        if gname and gname not in weights:
            n = len(l["tasks"])
            weights[gname] = l.get("weights") or [round(1 / n, 6)] * n
    kind = {k: v.get("part") for k, v in S.get("steps", {}).items() if v.get("part")}
    for a in feed["activities"]:
        if a.get("step") and a["step"] not in kind:
            kind[a["step"]] = V.get("kind", {}).get(a["step"], "pour")
    rings = rings_for(P, data, off)
    xs = [p[0] for r in rings for p in r] or [0]
    ys = [p[1] for r in rings for p in r] or [0]
    # the 4D stage puts update / baseline side by side along x: turn a block that is long in y by 90 degrees
    rot = V.get("rotate")
    if rot is None:
        rot = (max(ys) - min(ys)) > 2 * (max(xs) - min(xs))
    if rot:
        x1_, y0_ = max(xs), min(ys)
        rt = lambda pt: [round(pt[1] - y0_, 2), round(x1_ - pt[0], 2)]
        trl = lambda pts: [rt(q) for q in pts]
        rings = [trl(r) for r in rings]
        for e in els:
            e["x"], e["y"] = rt([e["x"], e["y"]])
            e["g"]["o"] = trl(e["g"]["o"])
            e["g"]["h"] = [trl(h) for h in e["g"]["h"]]
            if "p" in e["g"]:
                e["g"]["p"] = [[trl(o), [trl(h) for h in hs], z0, z1] for o, hs, z0, z1 in e["g"]["p"]]
        xs = [p[0] for r in rings for p in r] or [0]
        ys = [p[1] for r in rings for p in r] or [0]
    stands = []
    for s in data.get("stands", []):
        q, ang = [s["x"] + off[0], s["y"] + off[1]], s.get("a", 0)
        if rot:
            q, ang = rt(q), ang - math.pi / 2
        stands.append(dict(n=s.get("name", ""), l=s.get("label") or "PUMP " + s.get("name", ""), k=s.get("kind", "pump"),
                           x=round(q[0], 2), y=round(q[1], 2), z=s.get("z", 0), a=round(ang, 3), f=s.get("f"), t=s.get("t"),
                           c=s.get("c", []), p=s.get("p", []), b=s.get("serves") or s.get("name", "").split("/")))
    pits = []
    for p_ in data.get("pits", []):
        tq = (lambda pts: [rt([x + off[0], y + off[1]]) for x, y in pts]) if rot else (lambda pts: [[round(x + off[0], 2), round(y + off[1], 2)] for x, y in pts])
        pits.append(dict(id=p_["id"], r=[tq(r) for r in p_["r"]], zb=p_["zb"], zf=p_["zf"], d=p_["d"],
                         site=p_.get("site_level", 0), s=[dict(o=tq(s["o"]), h=[tq(h) for h in s["h"]]) for s in p_["s"]],
                         e=p_["e"], rp=p_["rp"], c=p_["c"], o=p_["o"]))
    terrain = None
    if data.get("terrain"):
        tr_ = data["terrain"]            # m: grid (gx, gy) -> stage (X, Y)
        m = [0, 1, off[1] - y0_, -1, 0, x1_ - off[0]] if rot else [1, 0, off[0], 0, 1, off[1]]
        # keep the part of the grid around the model: the update and baseline copies stand side by side and must not overlap
        st_, nx_, ny_ = tr_["step"], tr_["nx"], tr_["ny"]
        X = lambda gx, gy: m[0] * gx + m[1] * gy + m[2]
        Y = lambda gx, gy: m[3] * gx + m[4] * gy + m[5]
        ok = lambda j, i: (min(xs) - 6 <= X(tr_["x0"] + j * st_, tr_["y0"] + i * st_) <= max(xs) + 6
                           and min(ys) - 25 <= Y(tr_["x0"] + j * st_, tr_["y0"] + i * st_) <= max(ys) + 25)
        js = [j for j in range(nx_) if any(ok(j, i) for i in (0, ny_ // 2, ny_ - 1))]
        is_ = [i for i in range(ny_) if any(ok(j, i) for j in (0, nx_ // 2, nx_ - 1))]
        j0, j1, i0, i1 = min(js), max(js), min(is_), max(is_)
        cut = lambda a: [a[i * nx_ + j] for i in range(i0, i1 + 1) for j in range(j0, j1 + 1)]
        terrain = dict(x0=round(tr_["x0"] + j0 * st_, 2), y0=round(tr_["y0"] + i0 * st_, 2), step=st_, nx=j1 - j0 + 1, ny=i1 - i0 + 1,
                       z=cut(tr_["z"]), pit=cut(tr_["pit"]), m=m)
    floors = [S["fnd_floor"]] + [f for f in S["floors"] if f != S["fnd_floor"]]
    lv = data["levels"]
    below = {}
    for i, f in enumerate(floors):
        below[f] = lv.get(floors[i - 1] if i > 0 else f, lv.get(f))
    cr = S.get("crane", {})
    cfg = dict(name=P["project"]["name"], name_ar=P["project"].get("name_ar", P["project"]["name"]),
               subtitle=P["project"].get("subtitle", ""), code=P["project"]["code"], floors=floors,
               floor_names_ar=S.get("floor_names_ar", {}), groups=groups, weights=weights,
               kind=kind,
               trades={k: P["resources"][k].get("color") or "#cccccc" for k in tr}, zones=feed["zones"],
               bbox=[min(xs), max(xs), min(ys), max(ys)], off_weekdays=P["calendar"].get("off_weekdays", [4]),
               cx=cr.get("x", (min(xs) + max(xs)) / 2) if not rot else (min(xs) + max(xs)) / 2,
               crane_on=bool(cr.get("enabled")), rotated=bool(rot),
               currency=P["cost"].get("currency", P["project"].get("currency", "EGP")), below=below)
    dd = feed["data_date"]
    ddi = di.get(dd)
    if ddi is None:
        prior = [i for i, y in enumerate(dates) if y <= dd]
        ddi = prior[-1] if prior else 0
    dl = feed["daily"]
    tot = C.get("totals", {})
    sd = dict(cfg=cfg, rings=rings, stands=stands, pits=pits, terrain=terrain, dates=dates, acts=acts, els=els, dd=ddi, start=feed["start"],
              contract_end=feed.get("contract_end"),
              crane=dict(x=cr.get("x", 0) + off[0], y=round(cr.get("y", 0) + off[1], 2), jib=cr.get("jib_radius_m", 50)),
              levels=data["levels"], top=data["top"], pl=dl["planned_labour"], up=dl["upd_labour"],
              plt=dl["planned_by_trade"], upt=dl["upd_by_trade"], res={k: r["name_ar"] for k, r in P["resources"].items()},
              delay=feed.get("delay_days", 0), rates=C.get("rates", {}), lines=C.get("lines", []),
              cost=dict(bac_cost=int(round(tot.get("BAC", 0))), cur=cfg["currency"], ac_scenario=bool(C.get("ac_assumed")),
                        src=C.get("src", "no cost data yet (run cost.py)")))
    return sd


def build_html(P, data, sd):
    T = os.path.join(HERE, "templates")
    shell = open(os.path.join(T, "shell.html"), encoding="utf-8").read()
    p3 = open(os.path.join(T, "page_3d.txt"), encoding="utf-8").read()
    p4 = open(os.path.join(T, "page_4d5d.txt"), encoding="utf-8").read()
    bim = json.dumps(data, ensure_ascii=True, separators=(",", ":")).replace("</", "<\\/")
    p3 = p3.replace("@@BIMDATA@@", bim)
    sdj = json.dumps(sd, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    name = P["project"]["name"]
    html = (shell.replace("@@TITLE@@", f"{name} BIM")
            .replace("@@BAR@@", f"{P['project'].get('name_ar') or name} · BIM")
            .replace("@@KEY@@", f"{P['project']['code'].lower()}4d_override_v1"))
    for k, v in (("@@S3@@", p3), ("@@S4@@", p4), ("@@SD@@", sdj)):
        html = html.replace(k, v, 1)
    return html


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("out")
    a = ap.parse_args()
    P = load_project(a.project)
    data = load_json(ppath(P, P["inputs"]["data"]))
    fp = os.path.join(a.out, "feed.json")
    if not os.path.exists(fp):
        die("out/feed.json not found — run schedule.py and update.py first")
    feed = load_json(fp)
    cp = os.path.join(a.out, "cost.json")
    cost = load_json(cp) if os.path.exists(cp) else None
    sd = build_sd(P, data, feed, cost)
    html = build_html(P, data, sd)
    out = os.path.join(a.out, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    with open(os.path.join(a.out, "viewer_data.json"), "w", encoding="utf-8") as f:
        json.dump(sd, f, ensure_ascii=False)
    print(f"{out} ({len(html) / 1024:.0f} KB) | elements {len(sd['els'])} | activities {len(sd['acts'])} | "
          f"data date {sd['dates'][sd['dd']]} | BAC {sd['cost']['bac_cost']:,}")


if __name__ == "__main__":
    main()
