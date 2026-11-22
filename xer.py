"""Primavera P6 XER read / write (cp1252, CRLF, tab separated). English names only in the XER."""
import datetime as dt

F = {
    "CURRTYPE": "curr_id decimal_digit_cnt curr_symbol decimal_symbol digit_group_symbol pos_curr_fmt_type neg_curr_fmt_type curr_type curr_short_name group_digit_cnt base_exch_rate",
    "CALENDAR": "clndr_id default_flag clndr_name proj_id base_clndr_id last_chng_date clndr_type day_hr_cnt week_hr_cnt month_hr_cnt year_hr_cnt rsrc_private clndr_data",
    "PROJECT": "proj_id fy_start_month_num rsrc_self_add_flag allow_complete_flag rsrc_multi_assign_flag checkout_flag project_flag step_complete_flag cost_qty_recalc_flag batch_sum_flag name_sep_char def_complete_pct_type proj_short_name acct_id orig_proj_id source_proj_id base_type_id clndr_id sum_base_proj_id task_code_base task_code_step priority_num wbs_max_sum_level risk_level last_checksum critical_drtn_hr_cnt def_duration_type task_code_prefix guid def_qty_type add_by_name web_local_root_path proj_url def_rate_type add_act_remain_flag act_this_per_link_flag def_task_type act_pct_link_flag add_date last_tasksum_date fcst_start_date def_rsrc_roll_on_flag term_login last_fin_dates_id last_baseline_update_date cr_external_key apply_actuals_date location_id loaded_scope_level export_flag new_fin_dates_id baselines_to_export baseline_names_to_export next_data_date close_period_flag sum_refresh_date trsrcsum_loaded plan_start_date plan_end_date scd_end_date last_recalc_date",
    "PROJWBS": "wbs_id proj_id obs_id seq_num est_wt proj_node_flag sum_data_flag status_code wbs_short_name wbs_name phase_id parent_wbs_id ev_user_pct ev_etc_user_value orig_cost indep_remain_total_cost ann_dscnt_rate_pct dscnt_period_type indep_remain_work_qty anticip_start_date anticip_end_date ev_compute_type ev_etc_compute_type guid tmpl_guid plan_open_state",
    "RSRC": "rsrc_id parent_rsrc_id clndr_id rsrc_seq_num rsrc_name rsrc_short_name rsrc_type active_flag curr_id def_qty_per_hr cost_qty_type ot_factor auto_compute_act_flag def_cost_qty_link_flag ot_flag load_tasks_flag level_flag",
    "TASK": "task_id proj_id wbs_id clndr_id phys_complete_pct rev_fdbk_flag est_wt lock_plan_flag auto_compute_act_flag complete_pct_type task_type duration_type status_code task_code task_name rsrc_id total_float_hr_cnt free_float_hr_cnt remain_drtn_hr_cnt act_work_qty remain_work_qty target_work_qty target_drtn_hr_cnt target_equip_qty act_equip_qty remain_equip_qty cstr_date act_start_date act_end_date late_start_date late_end_date expect_end_date early_start_date early_end_date restart_date reend_date target_start_date target_end_date rem_late_start_date rem_late_end_date cstr_type priority_type suspend_date resume_date float_path float_path_order guid tmpl_guid cstr_date2 cstr_type2 driving_path_flag act_this_per_work_qty act_this_per_equip_qty external_early_start_date external_late_end_date create_date update_date create_user update_user location_id",
    "TASKPRED": "task_pred_id task_id pred_task_id proj_id pred_proj_id pred_type lag_hr_cnt float_path aref arls comments",
    "TASKRSRC": "taskrsrc_id task_id proj_id rsrc_id remain_qty target_qty remain_qty_per_hr target_qty_per_hr target_lag_drtn_hr_cnt cost_per_qty target_cost rsrc_type",
}
F = {k: v.split() for k, v in F.items()}

WBS_NAMES = {"1.1": "Preliminaries", "1.1.1": "Mobilisation", "1.1.2": "Tower crane", "1.2": "Substructure",
             "1.3": "Superstructure", "1.4": "Completion", "1.4.1": "Closeout"}


def _clean(s):
    return str(s).encode("cp1252", "replace").decode("cp1252").replace("\t", " ")


def _calendar_data(off_weekdays, hpd):
    # P6 day numbers: 1=Sunday ... 7=Saturday. Python weekday: Mon=0 ... Sun=6
    p6_off = {((w + 1) % 7) + 1 for w in off_weekdays}
    half = hpd / 2
    m_end = f"{8 + int(half):02d}:00"
    a_start = f"{9 + int(half):02d}:00"
    a_end = f"{9 + int(hpd):02d}:00"
    days = ""
    for n in (1, 2, 3, 4, 5, 7, 6):
        if n in p6_off:
            days += f"(0||{n}())"
        else:
            days += f"(0||{n}()(0||0(s|08:00|f|{m_end}))(0||1(s|{a_start}|f|{a_end})))"
    return f"(0||CalendarData()(0||DaysOfWeek(){days})(0||VIEW(ShowTotal|Y))(0||Exceptions()))"


def write_xer(P, acts, cal, path):
    S, C = P["schedule"], P["calendar"]
    hpd = C.get("hours_per_day", 8)
    work_days = 7 - len(C.get("off_weekdays", [4]))
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    start = S["start"] + " 00:00"
    cur = P["project"].get("currency", "EGP")
    T = {k: [] for k in F}
    T["CURRTYPE"].append(dict(curr_id=1, decimal_digit_cnt=2, curr_symbol=cur[:1], decimal_symbol=".", digit_group_symbol=",",
                              pos_curr_fmt_type="#1.1", neg_curr_fmt_type="(#1.1)", curr_type=cur, curr_short_name=cur,
                              group_digit_cnt=3, base_exch_rate=1))
    T["CALENDAR"].append(dict(clndr_id=1, default_flag="Y", clndr_name=C.get("name", "Calendar"), last_chng_date=now,
                              clndr_type="CA_Base", day_hr_cnt=f"{hpd:.1f}", week_hr_cnt=f"{hpd * work_days:.1f}",
                              month_hr_cnt=f"{hpd * work_days * 52 / 12:.1f}", year_hr_cnt=f"{hpd * work_days * 52:.1f}",
                              rsrc_private="N", clndr_data=_calendar_data(C.get("off_weekdays", [4]), hpd)))
    T["PROJECT"].append(dict(proj_id=1, fy_start_month_num=1, rsrc_self_add_flag="N", allow_complete_flag="N",
                             rsrc_multi_assign_flag="Y", checkout_flag="N", project_flag="Y", step_complete_flag="N",
                             cost_qty_recalc_flag="N", batch_sum_flag="N", name_sep_char=".", def_complete_pct_type="CP_Phys",
                             proj_short_name=P["project"].get("short_name", P["project"]["code"]), clndr_id=1,
                             task_code_base=1000, task_code_step=10, priority_num=10, wbs_max_sum_level=2,
                             critical_drtn_hr_cnt=0, def_duration_type="DT_FixedDUR2", task_code_prefix="A",
                             def_qty_type="QT_Hour", add_by_name="admin", def_rate_type="COST_PER_QTY",
                             add_act_remain_flag="N", act_this_per_link_flag="N", def_task_type="TT_Task",
                             act_pct_link_flag="N", add_date=now, def_rsrc_roll_on_flag="N", loaded_scope_level=7,
                             export_flag="Y", close_period_flag="N", trsrcsum_loaded="N", plan_start_date=start,
                             plan_end_date=(S.get("contract_end") or S["start"]) + " 00:00", last_recalc_date=start))
    # WBS tree
    codes = sorted({a["wbs"] for a in acts if a["wbs"]}, key=lambda s: [int(x) for x in s.split(".")])
    allc = set()
    for c in codes:
        parts = c.split(".")
        for i in range(1, len(parts) + 1):
            allc.add(".".join(parts[:i]))
    allc = sorted(allc, key=lambda s: [int(x) for x in s.split(".")])
    zones = {z["code"]: z for z in S["zones"]}
    wid = {}
    for i, c in enumerate(allc):
        wid[c] = 100 + i
        name = WBS_NAMES.get(c)
        if c == "1":
            name = f"{P['project']['name']} - Structural Works"
        if not name:
            a = next((x for x in acts if x["wbs"] == c), None)
            if a and a["floor"] == S.get("fnd_floor"):
                name = f"Raft - {zones[a['wing']]['name']}" if a["wing"] in zones else "Raft"
            elif a and a["floor"]:
                name = f"{a['floor']} floor"
            else:
                name = c
        parent = ".".join(c.split(".")[:-1])
        T["PROJWBS"].append(dict(wbs_id=wid[c], proj_id=1, obs_id=1, seq_num=(i + 1) * 10, est_wt=1,
                                 proj_node_flag="Y" if c == "1" else "N", sum_data_flag="N", status_code="WS_Open",
                                 wbs_short_name=c, wbs_name=_clean(name), parent_wbs_id=wid.get(parent, ""),
                                 ev_user_pct=6, ev_etc_user_value=0.88, orig_cost=0, indep_remain_total_cost=0,
                                 indep_remain_work_qty=0, ev_compute_type="EC_Cmp_pct", ev_etc_compute_type="EE_Rem_hr"))
    rid = {}
    for i, (k, r) in enumerate(P["resources"].items()):
        rid[k] = i + 1
        T["RSRC"].append(dict(rsrc_id=i + 1, clndr_id=1, rsrc_seq_num=(i + 1) * 10, rsrc_name=_clean(r["name"]),
                              rsrc_short_name=k, rsrc_type="RT_Labor" if r.get("kind") == "labour" else "RT_Equip",
                              active_flag="Y", curr_id=1, def_qty_per_hr=1, cost_qty_type="QT_Hour", ot_factor=0,
                              auto_compute_act_flag="N", def_cost_qty_link_flag="N", ot_flag="N", load_tasks_flag="N",
                              level_flag="N"))
    tid = {}
    for i, a in enumerate(acts):
        tid[a["code"]] = 1000 + i
        hrs = f"{a['dur'] * hpd:.1f}"
        T["TASK"].append(dict(task_id=1000 + i, proj_id=1, wbs_id=wid.get(a["wbs"], 100), clndr_id=1, phys_complete_pct=0,
                              rev_fdbk_flag="N", est_wt=1, lock_plan_flag="N", auto_compute_act_flag="N",
                              complete_pct_type="CP_Phys",
                              task_type={"start_milestone": "TT_Mile", "finish_milestone": "TT_FinMile"}.get(a["type"], "TT_Task"),
                              duration_type="DT_FixedDUR2", status_code="TK_NotStart", task_code=a["code"],
                              task_name=_clean(a["name"]), remain_drtn_hr_cnt=hrs, act_work_qty=0, remain_work_qty=0,
                              target_work_qty=0, target_drtn_hr_cnt=hrs, target_equip_qty=0, act_equip_qty=0,
                              remain_equip_qty=0, priority_type="PT_Normal", driving_path_flag="N",
                              act_this_per_work_qty=0, act_this_per_equip_qty=0, create_date=now, update_date=now,
                              create_user="admin", update_user="admin"))
    n = 0
    for a in acts:
        for p in a["preds"]:
            n += 1
            T["TASKPRED"].append(dict(task_pred_id=n, task_id=tid[a["code"]], pred_task_id=tid[p["task"]], proj_id=1,
                                      pred_proj_id=1, pred_type="PR_" + p.get("type", "FS"),
                                      lag_hr_cnt=f"{p.get('lag', 0) * hpd:.1f}"))
    n = 0
    for a in acts:
        for k, q in a["crew"].items():
            if k not in rid or not q:
                continue
            n += 1
            tot = q * a["dur"] * hpd
            fmt = lambda v: int(v) if float(v).is_integer() else round(v, 3)
            T["TASKRSRC"].append(dict(taskrsrc_id=n, task_id=tid[a["code"]], proj_id=1, rsrc_id=rid[k],
                                      remain_qty=fmt(tot), target_qty=fmt(tot), remain_qty_per_hr=fmt(q),
                                      target_qty_per_hr=fmt(q), target_lag_drtn_hr_cnt=0, cost_per_qty=0, target_cost=0,
                                      rsrc_type="RT_Labor" if P["resources"][k].get("kind") == "labour" else "RT_Equip"))
    lines = ["\t".join(["ERMHDR", "19.12", dt.date.today().isoformat(), "Project", "admin", "Admin", "Primavera",
                        "Project Management", cur])]
    for t, fields in F.items():
        lines.append("%T\t" + t)
        lines.append("%F\t" + "\t".join(fields))
        for r in T[t]:
            lines.append("%R\t" + "\t".join(_clean(r.get(f, "")) if r.get(f, "") != "" else "" for f in fields))
    lines.append("%E")
    with open(path, "wb") as f:
        f.write(("\r\n".join(lines) + "\r\n").encode("cp1252", "replace"))
    return path


def read_xer(path):
    raw = open(path, "rb").read()
    for enc in ("cp1252", "utf-8", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    out, cur, fields = {}, None, []
    for ln in txt.splitlines():
        p = ln.split("\t")
        if p[0] == "%T":
            cur = p[1]
            out[cur] = []
        elif p[0] == "%F":
            fields = p[1:]
        elif p[0] == "%R" and cur:
            out[cur].append(dict(zip(fields, p[1:])))
    return out
