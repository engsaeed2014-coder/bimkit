"""Excel outputs for Episode 2: Update sheet (site input) and Programme workbook (Gantt, manpower, crews)."""
import datetime as dt

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as CL

from common import DATE_FMT, d, head_row, styles


def labour_trades(P):
    return [k for k, r in P["resources"].items() if r.get("kind") == "labour"]


# ---------------------------------------------------------------- Update sheet
def write_update_template(P, acts, cal, path, feed=None):
    """Update sheet: row 5 headers, data from row 6, D2 = data date. Yellow cells are inputs.

    Without a feed the actual/forecast columns start equal to the baseline (no progress yet).
    """
    st = styles()
    tr = labour_trades(P)
    nt = len(tr)
    mask = cal.excel_mask()
    wb = Workbook()
    ws = wb.active
    ws.title = "Update"
    ws["A1"] = f"{P['project']['name']} — 4D Update Sheet / شيت تحديث البرنامج"
    ws["A1"].font = st["title"]
    ws["A2"] = "Data date / تاريخ البيانات"
    dd = d(feed["data_date"]) if feed else d(P["schedule"]["start"])
    ws["D2"] = dd
    ws["D2"].number_format = DATE_FMT
    ws["D2"].fill, ws["D2"].font = st["input_fill"], st["input_font"]
    ws["A3"] = "الخانات الصفراء (زرقاء الخط) = إدخال. الأزرق الفاتح = البيز لاين (لا يتغير). باقي الأعمدة معادلات."
    names_ar = {k: P["resources"][k]["name_ar"] for k in tr}
    for j, k in enumerate(tr):
        ws.cell(4, 17 + j, names_ar[k])
        ws.cell(4, 17 + nt + j, names_ar[k])
    hd = ["Code", "Activity / النشاط", "Floor", "Zone", "Step", "BL Start", "BL Finish", "BL Dur (d)",
          "Actual/Fcst Start", "Actual/Fcst Finish", "% Complete", "Dur (d)", "Start slip (d)", "Finish slip (d)",
          "Dur overrun (d)", "Status"] + [f"BL {k}" for k in tr] + [f"Act {k}" for k in tr]
    head_row(ws, 5, hd, st, [8, 52, 7, 6, 7, 12, 12, 9, 13, 13, 9, 8, 9, 9, 9, 10] + [9] * (2 * nt))
    F = {a["code"]: a for a in (feed["activities"] if feed else [])}
    r = 6
    for a in acts:
        if a["type"] != "task":
            continue
        u = F.get(a["code"], {})
        vals = [a["code"], a["name_ar"], a["floor"], a["wing"], a["step"], d(a["baseline_start"]), d(a["baseline_finish"]),
                f'=NETWORKDAYS.INTL(F{r},G{r},"{mask}")',
                d(u.get("upd_start") or a["baseline_start"]), d(u.get("upd_finish") or a["baseline_finish"]),
                (u.get("percent_complete", 0) or 0) / 100,
                f'=NETWORKDAYS.INTL(I{r},J{r},"{mask}")', f"=I{r}-F{r}", f"=J{r}-G{r}", f"=L{r}-H{r}",
                f'=IF(OR(K{r}>=1,J{r}<$D$2),"منفذ",IF(OR(K{r}>0,AND(I{r}<=$D$2,J{r}>=$D$2)),"جاري","لم يبدأ"))']
        bl = a.get("crew_baseline", a["crew"])
        act = u.get("crew_actual", a["crew"])
        vals += [bl.get(k, 0) for k in tr] + [act.get(k, 0) for k in tr]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = st["border"]
            if c in (6, 7) or 17 <= c < 17 + nt:
                cell.fill = st["base_fill"]
            if c in (9, 10, 11) or c >= 17 + nt:
                cell.fill, cell.font = st["input_fill"], st["input_font"]
            if c in (6, 7, 9, 10):
                cell.number_format = DATE_FMT
            if c == 11:
                cell.number_format = "0%"
        r += 1
    last = r - 1
    ws.freeze_panes = "C6"
    s = wb.create_sheet("Summary")
    rows = [("Summary / الملخص", None),
            ("Activities / عدد الأنشطة", f"=COUNTA(Update!A6:A{last})"),
            ("Done / منفذ", f'=COUNTIF(Update!P6:P{last},"منفذ")'),
            ("In progress / جاري", f'=COUNTIF(Update!P6:P{last},"جاري")'),
            ("Not started / لم يبدأ", f'=COUNTIF(Update!P6:P{last},"لم يبدأ")'),
            ("Baseline finish", f"=MAX(Update!G6:G{last})"),
            ("Actual/Forecast finish", f"=MAX(Update!J6:J{last})"),
            ("Project delay (days) / التأخير", "=B7-B6"),
            ("Late activities (finish slip>0)", f'=COUNTIF(Update!N6:N{last},">0")')]
    for i, (k, v) in enumerate(rows, 1):
        s.cell(i, 1, k)
        c = s.cell(i, 2, v)
        if i in (6, 7):
            c.number_format = DATE_FMT
    s["A1"].font = st["title"]
    s.column_dimensions["A"].width = 34
    s.column_dimensions["B"].width = 14
    h = wb.create_sheet("How to use")
    off = ", ".join(dt.date(2024, 1, 1 + w).strftime("%A") for w in P["calendar"].get("off_weekdays", [4]))
    for i, t in enumerate([
        f"1) عدّل في شيت Update الخانات الصفراء فقط: تاريخ البيانات (D2)، البداية/النهاية الفعلية أو المتوقعة، نسبة الإنجاز، وأعداد العمالة الفعلية لكل تخصص (Act ...).",
        f"2) المعادلات (المدة، التأخير، الحالة) بتتحسب لوحدها. الإجازة الأسبوعية: {off}.",
        "3) احفظ الملف وابعته كما هو، أو استورده في تاب 4D: هيتحدث الـ 4D والتأخير ومصادر التأخير والعمالة.",
        "4) متغيّرش الأكواد (Code) ولا ترتيب الأعمدة ولا أسماء الشيتات.",
        "5) تأخير النشاط = Finish slip. مصدر التأخير = زيادة المدة (Dur overrun) + بداية متأخرة بعد انتهاء السوابق.",
        "6) نشاط لسه مبدأش: لو سبت البداية/النهاية فاضية، السكريبت بيحسبها من المنطق (السوابق) بعد تاريخ البيانات."], 1):
        h.cell(i, 1, t)
    h.column_dimensions["A"].width = 130
    wb.save(path)
    return path


def read_update(P, path):
    """Return (data_date, {code: dict(start, finish, pct, crew)}) from an Update sheet."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    if "Update" not in wb.sheetnames:
        raise ValueError("sheet 'Update' not found")
    ws = wb["Update"]
    tr = labour_trades(P)
    nt = len(tr)
    heads = [c.value for c in ws[5]]
    act_cols = []
    for k in tr:
        h = f"Act {k}"
        act_cols.append(heads.index(h) if h in heads else 16 + nt + tr.index(k))
    rows = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        code = row[0]
        if not code:
            continue
        pct = row[10]
        if isinstance(pct, str):
            pct = pct.strip().rstrip("%") or 0
        pct = float(pct or 0)
        if pct > 1:          # typed as 75 instead of 75 %
            pct = pct / 100
        crew = {}
        for k, ci in zip(tr, act_cols):
            v = row[ci] if ci < len(row) else None
            if v not in (None, ""):
                crew[k] = float(v)
        rows[str(code).strip()] = dict(start=d(row[8]), finish=d(row[9]), pct=pct, crew=crew)
    return d(ws["D2"].value), rows


# ---------------------------------------------------------------- Programme workbook
def write_programme(P, acts, cal, links, data, path, feed=None):
    st = styles()
    S = P["schedule"]
    tr = labour_trades(P)
    eq = [k for k, r in P["resources"].items() if r.get("kind") != "labour"]
    F = {a["code"]: a for a in (feed["activities"] if feed else [])}
    dd = d(feed["data_date"]) if feed else None
    wb = Workbook()
    sm = wb.active
    sm.title = "Summary"

    # ---- Programme (Gantt) ----
    ws = wb.create_sheet("Programme")
    ws["A1"] = f"{P['project']['name']} - Structural Works: Baseline vs Actual / Forecast programme"
    ws["A1"].font = st["title"]
    ws["A2"] = "Data date"
    ws["B2"] = dd or d(S["start"])
    ws["B2"].number_format = DATE_FMT
    ws["B2"].fill, ws["B2"].font = st["input_fill"], st["input_font"]
    ws["C2"] = (f"Calendar: {cal.name}. Blue bar = baseline. Green = actual. Orange = forecast. "
                "Change the data date to move the status line.")
    hd = ["Code", "Activity", "النشاط", "WBS", "Row", "Start", "Finish", "Dur (wd)", "Status", "% complete",
          "Var finish (days)", "Total float (wd)", "Critical", "Predecessors", "Labour on activity", ""]
    head_row(ws, 4, hd, st, [8, 44, 44, 7, 8, 11, 11, 8, 11, 9, 9, 9, 7, 22, 9, 2])
    start = d(S["start"])
    wk0 = start - dt.timedelta(days=(start.weekday() + 1) % 7)       # week starts Sunday
    ends = [d(a["baseline_finish"]) for a in acts] + [d(F[c]["upd_finish"]) for c in F if F[c].get("upd_finish")]
    endd = max(ends)
    weeks = []
    w = wk0
    while w <= endd:
        weeks.append(w)
        w += dt.timedelta(days=7)
    C0 = 17
    for i, w in enumerate(weeks):
        c = ws.cell(4, C0 + i, w)
        c.number_format = "dd"
        c.font = Font(size=8, color="FFFFFF", bold=True)
        c.fill = st["head_fill"]
        ws.column_dimensions[CL(C0 + i)].width = 2.6
        if i == 0 or w.month != weeks[i - 1].month:
            ws.cell(3, C0 + i, w.strftime("%b %y")).font = Font(size=8, bold=True)
    r = 5
    wbs_names = _wbs_names(P, acts)
    cur_wbs = None
    first_row = r
    for a in acts:
        if a["wbs"] != cur_wbs:
            cur_wbs = a["wbs"]
            en, ar = wbs_names.get(cur_wbs, (cur_wbs, cur_wbs))
            for c, v in enumerate([cur_wbs, en, ar], 1):
                ws.cell(r, c, v).font = st["bold"]
                ws.cell(r, c).fill = st["total_fill"]
            r += 1
        lab = sum(v for k, v in a["crew"].items() if k in tr)
        base = [a["code"], a["name"], a["name_ar"], a["wbs"], "Baseline", d(a["baseline_start"]), d(a["baseline_finish"]),
                a["dur"], None, None, None, a["tf"], "Y" if a["tf"] <= 0 else None,
                ", ".join(p["task"] + ("" if p.get("type", "FS") == "FS" and not p.get("lag") else
                                       f" {p.get('type', 'FS')}{'+' + str(p['lag']) if p.get('lag') else ''}") for p in a["preds"]),
                round(lab, 1)]
        for c, v in enumerate(base, 1):
            cell = ws.cell(r, c, v)
            if c in (6, 7):
                cell.number_format = DATE_FMT
        u = F.get(a["code"])
        r2 = r + 1
        if u:
            us, uf = d(u["upd_start"]), d(u["upd_finish"])
            status = "Complete" if u["percent_complete"] >= 100 else ("In progress" if u["percent_complete"] > 0 else "Not started")
            row2 = [None, None, None, None, "Actual" if u["percent_complete"] > 0 else "Forecast", us, uf,
                    (cal.networkdays(us, uf) if a["type"] == "task" else 0), status, u["percent_complete"] / 100,
                    f"=G{r2}-G{r}"]
        else:
            row2 = [None, None, None, None, "Forecast", d(a["baseline_start"]), d(a["baseline_finish"]), a["dur"],
                    "Not started", 0, f"=G{r2}-G{r}"]
        for c, v in enumerate(row2, 1):
            cell = ws.cell(r2, c, v)
            if c in (6, 7):
                cell.number_format = DATE_FMT
            if c == 10:
                cell.number_format = "0%"
        r += 2
    last_row = r - 1
    rng = f"{CL(C0)}{first_row}:{CL(C0 + len(weeks) - 1)}{last_row}"
    q = CL(C0)
    blue = PatternFill("solid", fgColor="2F75B5", bgColor="2F75B5")
    green = PatternFill("solid", fgColor="3CB56A", bgColor="3CB56A")
    orange = PatternFill("solid", fgColor="F28C28", bgColor="F28C28")
    red = PatternFill("solid", fgColor="E5594F", bgColor="E5594F")
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($E{first_row}="Actual",$F{first_row}<>"",{q}$4<=$B$2,{q}$4+6>=$F{first_row},{q}$4<=$G{first_row})'],
        fill=green, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(OR($E{first_row}="Actual",$E{first_row}="Forecast"),$F{first_row}<>"",{q}$4+6>=$F{first_row},{q}$4<=$G{first_row})'],
        fill=orange, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($E{first_row}="Baseline",$M{first_row}="Y",{q}$4+6>=$F{first_row},{q}$4<=$G{first_row})'],
        fill=red, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($E{first_row}="Baseline",{q}$4+6>=$F{first_row},{q}$4<=$G{first_row})'], fill=blue, stopIfTrue=True))
    ws.freeze_panes = "F5"

    # ---- Manpower (weekly) ----
    mp = wb.create_sheet("Manpower")
    mp["A1"] = "Weekly labour on site: planned (baseline) vs actual"
    mp["A1"].font = st["title"]
    wd_week = 7 - len(P["calendar"].get("off_weekdays", [4]))
    mp["A2"] = (f"Average daily head-count per week (labour only, equipment excluded). Planned = crew table x baseline dates. "
                f"Actual = Update sheet crews to the data date. {wd_week} working days per week.")
    hdm = ["Week start", "Planned avg/day", "Actual avg/day", "Variance (actual - planned)", "Planned man-days",
           "Actual man-days"] + [f"{P['resources'][k]['name']} (planned)" for k in eq[:1]] + \
          [f"{P['resources'][k]['name']} (planned)" for k in tr]
    head_row(mp, 4, hdm, st, [12, 11, 11, 12, 11, 11] + [13] * (len(tr) + 1))
    daily = feed["daily"] if feed else None
    plan_d = {}
    plan_t = {k: {} for k in tr}
    eq_d = {}
    for a in acts:
        if a["type"] != "task":
            continue
        for i in range(a["es"], a["ef"]):
            x = cal.date(i)
            for k in tr:
                plan_t[k][x] = plan_t[k].get(x, 0) + a["crew"].get(k, 0)
            plan_d[x] = plan_d.get(x, 0) + sum(a["crew"].get(k, 0) for k in tr)
            if eq:
                eq_d[x] = eq_d.get(x, 0) + a["crew"].get(eq[0], 0)
    act_d = {}
    if daily:
        for x, v in zip(daily["dates"], daily["actual_labour"]):
            act_d[d(x)] = v or 0
    r = 5
    for w in weeks:
        days = [w + dt.timedelta(days=i) for i in range(7)]
        pm = sum(plan_d.get(x, 0) for x in days)
        am = sum(act_d.get(x, 0) for x in days)
        row = [w, f'=IF(COUNT(E{r})=0,"",ROUND(E{r}/{wd_week},1))', f"=ROUND(F{r}/{wd_week},1)", f"=C{r}-B{r}",
               round(pm, 1), round(am, 1)]
        if eq:
            row.append(round(sum(eq_d.get(x, 0) for x in days) / wd_week, 1))
        row += [round(sum(plan_t[k].get(x, 0) for x in days) / wd_week, 1) for k in tr]
        for c, v in enumerate(row, 1):
            cell = mp.cell(r, c, v)
            if c == 1:
                cell.number_format = DATE_FMT
        r += 1
    mp.cell(r, 1, "Total").font = st["bold"]
    mp.cell(r, 5, f"=SUM(E5:E{r - 1})")
    mp.cell(r, 6, f"=SUM(F5:F{r - 1})")
    mp.cell(r + 1, 1, "Peak planned avg/day").font = st["bold"]
    mp.cell(r + 1, 2, f"=MAX(B5:B{r - 1})")
    mp_total, mp_peak = r, r + 1

    # ---- Crews ----
    cw = wb.create_sheet("Crews")
    cw["A1"] = "Planned crew per activity type (workers per zone, per day)"
    cw["A1"].font = st["title"]
    cw["A2"] = "ASSUMED planning crews. The XER carries the same crews as resource assignments."
    res = tr + eq
    hdc = ["Step", "Activity", "النشاط", "Duration (wd)"] + \
          [f"{P['resources'][k]['name']} / {P['resources'][k]['name_ar']}" for k in res] + ["Labour total"]
    head_row(cw, 4, hdc, st, [7, 40, 34, 10] + [13] * len(res) + [10])
    durs = {}
    for a in acts:
        if a["step"]:
            durs.setdefault(a["step"], a["dur"])
    r = 5
    for k, s in S["steps"].items():
        row = [k, s["name"], s["name_ar"], durs.get(k)] + [s.get("crew", {}).get(x) for x in res]
        for c, v in enumerate(row, 1):
            cw.cell(r, c, v)
        cw.cell(r, 5 + len(res), f"=SUM({CL(5)}{r}:{CL(4 + len(tr))}{r})")
        r += 1

    # ---- Crane ----
    cr = S.get("crane", {})
    if cr.get("enabled"):
        ck = wb.create_sheet("Crane")
        ck["A1"] = "Tower crane plan"
        ck["A1"].font = st["title"]
        head_row(ck, 3, ["Item", "From", "To", "Basis"], st, [46, 12, 12, 14])
        A = {a["code"]: a for a in acts}
        vert = [a for a in acts if a["step"] in S["sequence"].get("VERT", [])]
        rows = [("Erection (A1020)", A["A1020"]["baseline_start"], A["A1020"]["baseline_finish"], "Baseline logic")]
        if vert:
            rows.append(("Working period (first to last vertical activity)", min(a["baseline_start"] for a in vert),
                         max(a["baseline_finish"] for a in vert), "CALC"))
        if "A9000" in A:
            rows.append(("Dismantling (A9000)", A["A9000"]["baseline_start"], A["A9000"]["baseline_finish"], "CALC"))
        if "A1020" in F and F["A1020"].get("actual_start"):
            rows.append(("Actual erection (progress update)", F["A1020"]["actual_start"],
                         F["A1020"].get("actual_finish") or F["A1020"]["upd_finish"], "Actual"))
        for i, row in enumerate(rows, 4):
            for c, v in enumerate(row, 1):
                cell = ck.cell(i, c, d(v) if c in (2, 3) else v)
                if c in (2, 3):
                    cell.number_format = DATE_FMT
        i = 4 + len(rows) + 2
        ck.cell(i, 1, "Crane data used in the 4D").font = st["bold"]
        for j, row in enumerate([("Position (model coordinates, m)", f"x {cr['x']}, y {cr['y']}", None, cr.get("position_confidence", "ASSUMED")),
                                 ("Jib radius", f"{cr.get('jib_radius_m')} m", None, cr.get("position_confidence", "ASSUMED")),
                                 ("Hook height", f"{cr.get('hook_height_m')} m", None, cr.get("position_confidence", "ASSUMED")),
                                 ("Sharing", "One crane serves all zones (see Crews sheet for units per activity)", None, "ASSUMED")], i + 1):
            for c, v in enumerate(row, 1):
                ck.cell(j, c, v)

    # ---- Element links ----
    el = wb.create_sheet("Element_Links")
    el["A1"] = "Activity to model-element links (for the 4D)"
    el["A1"].font = st["title"]
    zs = "; ".join(f"{z['code']} = {z['name']}" for z in S["zones"])
    el["A2"] = f"Rule: floor + zone + element group. Zones: {zs}. Every model element has exactly one pour activity."
    head_row(el, 4, ["Code", "Activity", "Floor", "Zone", "Group", "Elements", "Concrete m3 (part of work)"], st,
             [8, 50, 7, 6, 7, 9, 14])
    vol = {e["id"]: e["volume_m3"] for e in data["elements"]}
    per = {}
    for eid, l in links.items():
        for t in l["tasks"]:
            per.setdefault(t, []).append(eid)
    part_of = {s: g for g, ss in S["sequence"].items() for s in ss}
    r = 5
    for a in acts:
        if a["code"] not in per:
            continue
        ids = per[a["code"]]
        row = [a["code"], a["name"], a["floor"], a["wing"], part_of.get(a["step"]), len(ids), round(sum(vol[i] for i in ids), 1)]
        for c, v in enumerate(row, 1):
            el.cell(r, c, v)
        r += 1

    # ---- Assumptions ----
    asn = wb.create_sheet("Assumptions")
    asn["A1"] = "Assumptions and notes"
    asn["A1"].font = st["title"]
    K = S.get("duration_scale", 1)
    cyc = ", ".join(f"{k} {durs.get(k)}" for g in ("SLB", "VERT") for k in S["sequence"].get(g, []))
    raft = ", ".join(f"{k} {durs.get(k)}" for k in S["sequence"].get("FND", []))
    notes = [
        f"1. Programme start {d(S['start']):%d-%b-%Y}, contract completion {d(S['contract_end']):%d-%b-%Y}." +
        (f" Baseline durations scaled x{K} ({S.get('duration_scale_note', 'project.json')}); crews are the Crews sheet figures (ASSUMED)." if K != 1 else ""),
        f"2. Work week: {cal.name}, {P['calendar'].get('hours_per_day', 8)} hours. Holidays loaded: {len(P['calendar'].get('holidays', []))}.",
        f"3. Floor cycle (working days): {cyc}. Zones follow the first zone by an SS lag of {S['logic']['zone_ss_lag']} days.",
        f"4. Raft: {raft} working days per zone (ASSUMED).",
        "5. Crews and equipment (Crews sheet) are ASSUMED planning figures.",
        "6. Tower crane position, jib radius and hook height are as in project.json (Crane sheet).",
        "7. Excavation, MEP and finishes are outside this structural programme.",
        "8. Actual / forecast rows come from the Update sheet at the data date; replace them with site records.",
        "9. The XER carries durations, logic and resource assignments; P6 calculates dates on import (F9). "
        "If the import complains about resources, import without them first."]
    for i, t in enumerate(notes, 3):
        asn.cell(i, 1, t)
    asn.column_dimensions["A"].width = 140

    # ---- Summary ----
    sm["A1"] = f"{P['project']['name']} - Structural Works programme"
    sm["A1"].font = st["title"]
    bfin = max(d(a["baseline_finish"]) for a in acts)
    ffin = max(d(F[c]["upd_finish"]) for c in F) if F else bfin
    rows = [("Project start", d(S["start"])), ("Contract completion", d(S["contract_end"])), ("Baseline finish", bfin),
            ("Float to contract (calendar days)", "=B4-B5"), ("Forecast finish (current logic)", ffin),
            ("Forecast slip vs baseline (calendar days)", "=B7-B5"), ("Data date", "=Programme!B2"),
            ("Activities (incl. milestones)", len(acts)),
            ("Critical activities (float <= 0)", f'=COUNTIF(Programme!L5:L{last_row},"<=0")'),
            ("Baseline: activities planned to be complete at data date",
             sum(1 for a in acts if dd and d(a["baseline_finish"]) <= dd) if dd else None),
            ("Actual: activities complete at data date",
             sum(1 for c in F if F[c]["percent_complete"] >= 100 and F[c]["type"] == "task") if F else None),
            ("Peak planned labour (avg/day)", f"=Manpower!B{mp_peak}"),
            ("Planned man-days (labour, total)", f"=Manpower!E{mp_total}"),
            ("Actual man-days to data date", f"=Manpower!F{mp_total}"),
            ("Model elements linked", len(links)),
            ("Concrete in model (m3)", round(sum(e["volume_m3"] for e in data["elements"]), 1))]
    for i, (k, v) in enumerate(rows, 3):
        sm.cell(i, 1, k)
        c = sm.cell(i, 2, v)
        if isinstance(v, dt.date) or (isinstance(v, str) and v == "=Programme!B2"):
            c.number_format = DATE_FMT
    sm.column_dimensions["A"].width = 52
    sm.column_dimensions["B"].width = 14
    wb.save(path)
    return path


def _wbs_names(P, acts):
    S = P["schedule"]
    z = {x["code"]: x for x in S["zones"]}
    fl_ar = S.get("floor_names_ar", {})
    out = {"1.1.1": ("Mobilisation", "التجهيز"), "1.1.2": ("Tower crane", "الونش البرجي"),
           "1.4.1": ("Closeout", "الإنهاء والتسليم")}
    for a in acts:
        if a["wbs"] in out:
            continue
        if a["floor"] == S.get("fnd_floor") and a["wing"] in z:
            out[a["wbs"]] = (f"Raft - {z[a['wing']]['name']}", f"لبشة - {z[a['wing']]['name_ar']}")
        elif a["floor"]:
            out[a["wbs"]] = (f"{a['floor']} floor", fl_ar.get(a["floor"], a["floor"]))
    return out
