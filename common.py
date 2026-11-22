"""Shared helpers for the BIM kit: project file, working-day calendar, CPM, xlsx styles."""
import datetime as dt
import json
import math
import os
import sys

# ---------------------------------------------------------------- project
def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_json(obj, path, indent=None):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=indent)


def load_project(path):
    """Read project.json; relative input paths are resolved against its folder."""
    p = load_json(path)
    p["_dir"] = os.path.dirname(os.path.abspath(path))
    return p


def ppath(p, rel):
    if rel is None:
        return None
    return rel if os.path.isabs(rel) else os.path.join(p["_dir"], rel)


def fname(p, suffix):
    """Output file name prefix, e.g. Bicon_Building_Budget_Baseline.xlsx."""
    base = p["project"].get("file_prefix") or p["project"]["name"].replace(" ", "_")
    return f"{base}_{suffix}"


def die(msg):
    print("ERROR:", msg, file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------- dates
def d(s):
    if s is None or s == "":
        return None
    if isinstance(s, dt.datetime):
        return s.date()
    if isinstance(s, dt.date):
        return s
    return dt.date.fromisoformat(str(s)[:10])


def iso(x):
    return x.isoformat() if x else None


class Calendar:
    """Working-day calendar. off_weekdays use Python weekday numbers (Mon=0 ... Fri=4, Sun=6)."""

    def __init__(self, start, off_weekdays=(4,), holidays=(), name="Working calendar"):
        self.start = d(start)
        self.off = set(off_weekdays)
        self.hol = {d(h) for h in holidays}
        self.name = name
        self.days = []          # index -> date
        self.idx = {}
        self._extend(800)

    def working(self, x):
        return x.weekday() not in self.off and x not in self.hol

    def _extend(self, n):
        x = self.days[-1] + dt.timedelta(days=1) if self.days else self.start
        while len(self.days) < n:
            if self.working(x):
                self.idx[x] = len(self.days)
                self.days.append(x)
            x += dt.timedelta(days=1)

    def date(self, i):
        if i >= len(self.days):
            self._extend(i + 200)
        return self.days[i]

    def index(self, x, snap="next"):
        """Index of a date; a non-working date snaps to the next (or previous) working day."""
        x = d(x)
        if x < self.start:
            raise ValueError(f"{x} is before the calendar start {self.start}")
        while x > self.days[-1]:
            self._extend(len(self.days) + 200)
        step = 1 if snap == "next" else -1
        while x not in self.idx:
            x += dt.timedelta(days=step)
        return self.idx[x]

    def networkdays(self, a, b):
        return self.index(b, "prev") - self.index(a) + 1

    # Excel NETWORKDAYS.INTL weekend mask, Monday first ("0000100" = Friday off)
    def excel_mask(self):
        return "".join("1" if i in self.off else "0" for i in range(7))


# ---------------------------------------------------------------- CPM
def cpm(acts, start_index=0):
    """Forward/backward pass on working-day indices.

    acts: ordered list of dicts with code, dur (working days, 0 = milestone),
    type (task/start_milestone/finish_milestone), preds [{task,type,lag}].
    Time model: a task occupies [s, s+dur); finish point f = s+dur.
    Sets es, ef, ls, lf, tf (working days). Raises on cycles / unknown preds.
    """
    by = {a["code"]: a for a in acts}
    for a in acts:
        for p in a["preds"]:
            if p["task"] not in by:
                raise ValueError(f"{a['code']}: unknown predecessor {p['task']}")
    order = topo(acts)
    for c in order:
        a = by[c]
        es = start_index
        for p in a["preds"]:
            b = by[p["task"]]
            lag = p.get("lag", 0)
            t = p.get("type", "FS")
            if t == "FS":
                es = max(es, b["ef"] + lag)
            elif t == "SS":
                es = max(es, b["es"] + lag)
            elif t == "FF":
                es = max(es, b["ef"] + lag - a["dur"])
            elif t == "SF":
                es = max(es, b["es"] + lag - a["dur"])
        a["es"], a["ef"] = es, es + a["dur"]
    finish = max(a["ef"] for a in acts)
    succ = {c: [] for c in by}
    for a in acts:
        for p in a["preds"]:
            succ[p["task"]].append((a["code"], p))
    for c in reversed(order):
        a = by[c]
        lf = finish
        for s, p in succ[c]:
            b = by[s]
            lag = p.get("lag", 0)
            t = p.get("type", "FS")
            if t == "FS":
                lf = min(lf, b["ls"] - lag)
            elif t == "SS":
                lf = min(lf, b["ls"] - lag + a["dur"])
            elif t == "FF":
                lf = min(lf, b["lf"] - lag)
            elif t == "SF":
                lf = min(lf, b["lf"] - lag + a["dur"])
        a["lf"], a["ls"] = lf, lf - a["dur"]
        a["tf"] = a["ls"] - a["es"]
    return finish


def topo(acts):
    by = {a["code"]: a for a in acts}
    state, out = {}, []

    def visit(c, path):
        if state.get(c) == 1:
            raise ValueError("logic loop: " + " -> ".join(path + [c]))
        if state.get(c) == 2:
            return
        state[c] = 1
        for p in by[c]["preds"]:
            visit(p["task"], path + [c])
        state[c] = 2
        out.append(c)

    sys.setrecursionlimit(10000)
    for a in acts:
        visit(a["code"], [])
    return out


def task_dates(cal, a, s, f):
    """Calendar dates for an activity occupying [s, f)."""
    if a["type"] == "finish_milestone":
        x = cal.date(max(f - 1, 0))
        return x, x
    if a["type"] == "start_milestone" or f == s:
        x = cal.date(s)
        return x, x
    return cal.date(s), cal.date(f - 1)


def open_ends(acts):
    """Activities without predecessor / successor (other than the two milestones)."""
    has_succ = {p["task"] for a in acts for p in a["preds"]}
    no_pred = [a["code"] for a in acts if not a["preds"] and a["type"] != "start_milestone"]
    no_succ = [a["code"] for a in acts if a["code"] not in has_succ and a["type"] != "finish_milestone"]
    return no_pred, no_succ


def scaled(base, k, rounding="ceil"):
    v = base * k
    if rounding == "round":
        return int(round(v))
    if rounding == "none":
        return v
    return int(math.ceil(v - 1e-9))


# ---------------------------------------------------------------- xlsx styling
def styles():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style="thin", color="BFC5CC")
    return {
        "title": Font(bold=True, size=13),
        "bold": Font(bold=True),
        "head_fill": PatternFill("solid", fgColor="1F3A5F"),
        "head_font": Font(bold=True, color="FFFFFF"),
        "input_fill": PatternFill("solid", fgColor="FFF2B3"),
        "input_font": Font(color="0000FF"),
        "base_fill": PatternFill("solid", fgColor="DDEBF7"),
        "total_fill": PatternFill("solid", fgColor="E7E9EC"),
        "border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "wrap": Alignment(wrap_text=True, vertical="top"),
        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
    }


def head_row(ws, row, headers, st, widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.fill, c.font, c.alignment, c.border = st["head_fill"], st["head_font"], st["center"], st["border"]
    if widths:
        from openpyxl.utils import get_column_letter
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


DATE_FMT = "dd-mmm-yy"
MONEY_FMT = "#,##0"
