# -*- coding: utf-8 -*-
import io
from typing import Dict, List, Tuple, Set, Optional
from collections import defaultdict

import streamlit as st
import pandas as pd
import altair as alt

# ---- add near the top (Excel I/O helpers) ----
import io
from datetime import date, timedelta
import numpy as np
import pandas as pd

# ---- PuLP (ILP solver) ----
try:
    import pulp
except ImportError:
    st.stop()


def make_template_excel_with_priority() -> bytes:
    rng = np.random.default_rng(7)

    employees = ["Aki", "Ben", "Chika", "Daichi", "Emi"]
    start = date(2025, 8, 1)
    end = date(2025, 8, 31)
    days = [
        (start + timedelta(days=i)).isoformat() for i in range((end - start).days + 1)
    ]
    tasks = ["Cashier", "Kitchen", "Floor", "MachineOp"]

    df_employees = pd.DataFrame({"employee": employees})
    df_days = pd.DataFrame({"day": days})
    df_tasks = pd.DataFrame({"task": tasks})

    # Availability (~75% available; log only False + some True)
    avail_rows = []
    for e in employees:
        for d in days:
            available = rng.random() > 0.25
            if not available or (available and rng.random() < 0.08):
                avail_rows.append({"employee": e, "day": d, "available": available})
    df_availability = pd.DataFrame(avail_rows)

    # Demand with weekend tilt
    def is_weekend(iso):
        return date.fromisoformat(iso).weekday() >= 5

    need_rows = []
    for d in days:
        p = [0.25, 0.55, 0.20] if not is_weekend(d) else [0.15, 0.60, 0.25]
        for t in tasks:
            need_rows.append(
                {"day": d, "task": t, "need": int(rng.choice([0, 1, 2], p=p))}
            )
    df_demand = pd.DataFrame(need_rows)
    if df_demand["need"].sum() < 30:
        idx = rng.choice(df_demand.index, size=30, replace=False)
        df_demand.loc[idx, "need"] = df_demand.loc[idx, "need"].clip(lower=1)

    # Incompatibilities
    pairs = [
        ("Aki", "Ben"),
        ("Ben", "Chika"),
        ("Chika", "Daichi"),
        ("Ben", "Emi"),
        ("Aki", "Emi"),
    ]
    pairs = sorted({tuple(sorted(p)) for p in pairs})
    df_incompat_global = pd.DataFrame(pairs, columns=["employee_a", "employee_b"])

    ibd_rows = []
    for _ in range(14):
        d = rng.choice(days)
        a, b = sorted(rng.choice(employees, size=2, replace=False))
        ibd_rows.append({"day": d, "employee_a": a, "employee_b": b})
    df_incompat_by_day = pd.DataFrame(ibd_rows)

    # CanDo (some False)
    cando_rows = []
    for e in employees:
        if rng.random() < 0.55:
            cannot = "MachineOp" if rng.random() < 0.7 else rng.choice(tasks)
            cando_rows.append({"employee": e, "task": cannot, "can_do": False})
    df_cando = pd.DataFrame(cando_rows)

    # MinMax
    mins = rng.integers(6, 11, size=len(employees))
    maxs = mins + rng.integers(6, 9, size=len(employees))
    df_minmax = pd.DataFrame(
        {"employee": employees, "min_shifts": mins, "max_shifts": maxs}
    )

    # Options (+ weight_pref)
    df_options = pd.DataFrame(
        {
            "incompat_level": ["day"],
            "weight_unmet": [1200.0],
            "weight_fair": [1.0],
            "allow_unmet_via_slack": [True],
            "time_limit_sec": [20],
            "weight_pref": [1.5],
        }
    )

    # Priority
    weekend_days = [d for d in days if is_weekend(d)]
    pr = [
        {"employee": "Emi", "day": "", "task": "", "score": 1.0},
        {"employee": "Aki", "day": "", "task": "Cashier", "score": 1.5},
        {"employee": "Daichi", "day": "", "task": "MachineOp", "score": 2.0},
    ]
    for d in rng.choice(weekend_days, size=min(8, len(weekend_days)), replace=False):
        pr.append({"employee": "Aki", "day": d, "task": "", "score": 1.0})
    for d in rng.choice(weekend_days, size=min(8, len(weekend_days)), replace=False):
        pr.append({"employee": "Emi", "day": d, "task": "Floor", "score": 1.2})
    df_priority = pd.DataFrame(pr)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_employees.to_excel(w, "Employees", index=False)
        df_days.to_excel(w, "Days", index=False)
        df_tasks.to_excel(w, "Tasks", index=False)
        df_availability.to_excel(w, "Availability", index=False)
        df_demand.to_excel(w, "Demand", index=False)
        df_incompat_global.to_excel(w, "IncompatibilitiesGlobal", index=False)
        df_incompat_by_day.to_excel(w, "IncompatibilitiesByDay", index=False)
        df_cando.to_excel(w, "CanDo", index=False)
        df_minmax.to_excel(w, "MinMax", index=False)
        df_options.to_excel(w, "Options", index=False)
        df_priority.to_excel(w, "Priority", index=False)
    buf.seek(0)
    return buf.read()


# ====================== 最適化ロジック ======================


def solve_shift_scheduling(
    employees: List[str],
    days: List[str],
    tasks: List[str],
    availability: Optional[Dict[Tuple[str, str], bool]] = None,
    demand: Optional[Dict[Tuple[str, str], int]] = None,
    incompatible_pairs_global: Optional[Set[Tuple[str, str]]] = None,
    incompatible_pairs_by_day: Optional[Dict[str, Set[Tuple[str, str]]]] = None,
    incompat_level: str = "day",  # "day" or "task"
    can_do: Optional[Dict[Tuple[str, str], bool]] = None,
    min_shifts_per_emp: Optional[Dict[str, int]] = None,
    max_shifts_per_emp: Optional[Dict[str, int]] = None,
    weight_unmet: float = 1000.0,
    weight_fair: float = 1.0,
    allow_unmet_via_slack: bool = True,
    time_limit_sec: Optional[int] = 20,
    # ★ 追加: 優先スコア
    priority_scores: Optional[Dict[Tuple[str, str, str], float]] = None,
    weight_pref: float = 1.0,
):
    employees = list(
        dict.fromkeys([str(x).strip() for x in employees if str(x).strip()])
    )
    days = list(dict.fromkeys([str(x).strip() for x in days if str(x).strip()]))
    tasks = list(dict.fromkeys([str(x).strip() for x in tasks if str(x).strip()]))

    availability = availability or {}
    demand = demand or {}
    incompatible_pairs_global = incompatible_pairs_global or set()
    incompatible_pairs_by_day = incompatible_pairs_by_day or {}
    can_do = can_do or {}
    min_shifts_per_emp = min_shifts_per_emp or {}
    max_shifts_per_emp = max_shifts_per_emp or {}
    priority_scores = priority_scores or {}

    total_demand = sum(max(0, int(demand.get((d, t), 0))) for d in days for t in tasks)
    avg_target = total_demand / max(1, len(employees))

    def norm_pair(a: str, b: str) -> Tuple[str, str]:
        return tuple(sorted((a, b)))  # type: ignore

    incompatible_pairs_global = {
        norm_pair(a, b) for (a, b) in incompatible_pairs_global
    }
    incompatible_pairs_by_day = {
        d: {norm_pair(a, b) for (a, b) in pairs}
        for d, pairs in incompatible_pairs_by_day.items()
    }

    model = pulp.LpProblem("ShiftScheduling", pulp.LpMinimize)

    x = pulp.LpVariable.dicts(
        "x",
        ((e, d, t) for e in employees for d in days for t in tasks),
        0,
        1,
        cat=pulp.LpBinary,
    )

    u = {}
    if allow_unmet_via_slack:
        u = pulp.LpVariable.dicts(
            "unmet", ((d, t) for d in days for t in tasks), 0, None, cat=pulp.LpInteger
        )

    y = pulp.LpVariable.dicts(
        "y", (e for e in employees), 0, None, cat=pulp.LpContinuous
    )
    p_dev = pulp.LpVariable.dicts(
        "pos_dev", (e for e in employees), 0, None, cat=pulp.LpContinuous
    )
    n_dev = pulp.LpVariable.dicts(
        "neg_dev", (e for e in employees), 0, None, cat=pulp.LpContinuous
    )

    # 需要
    for d in days:
        for t in tasks:
            need = int(demand.get((d, t), 0))
            if allow_unmet_via_slack:
                model += pulp.lpSum(x[(e, d, t)] for e in employees) + u[(d, t)] == need
            else:
                model += pulp.lpSum(x[(e, d, t)] for e in employees) == need

    # 可用性・1日1タスク・適正
    for e in employees:
        for d in days:
            can_work_today = availability.get((e, d), True)
            if not can_work_today:
                for t in tasks:
                    model += x[(e, d, t)] == 0
            model += pulp.lpSum(x[(e, d, t)] for t in tasks) <= (
                1 if can_work_today else 0
            )
            for t in tasks:
                if can_do and (e, t) in can_do and not can_do[(e, t)]:
                    model += x[(e, d, t)] == 0

    # 相性NG
    if incompat_level not in ("day", "task"):
        raise ValueError("incompat_level は 'day' か 'task' を指定してください。")

    for d in days:
        pairs_today = incompatible_pairs_global.union(
            incompatible_pairs_by_day.get(d, set())
        )
        for a, b in pairs_today:
            if incompat_level == "day":
                model += (
                    pulp.lpSum(x[(a, d, t)] for t in tasks)
                    + pulp.lpSum(x[(b, d, t)] for t in tasks)
                    <= 1
                )
            else:
                for t in tasks:
                    model += x[(a, d, t)] + x[(b, d, t)] <= 1

    # 勤務数と偏差
    for e in employees:
        model += y[e] == pulp.lpSum(x[(e, d, t)] for d in days for t in tasks)
        model += y[e] - avg_target == p_dev[e] - n_dev[e]

    # 上下限制約
    for e in employees:
        if e in min_shifts_per_emp:
            model += y[e] >= int(min_shifts_per_emp[e])
        if e in max_shifts_per_emp:
            model += y[e] <= int(max_shifts_per_emp[e])

    # 目的関数
    obj = 0
    if allow_unmet_via_slack:
        obj += weight_unmet * pulp.lpSum(u[(d, t)] for d in days for t in tasks)
    obj += weight_fair * pulp.lpSum(p_dev[e] + n_dev[e] for e in employees)

    # ★ 優先ボーナス（score が大きい割当を“得”にする、最小化なので“-”）
    if priority_scores:
        obj += -weight_pref * pulp.lpSum(
            (priority_scores.get((e, d, t), 0.0)) * x[(e, d, t)]
            for e in employees
            for d in days
            for t in tasks
        )

    model += obj

    solver = (
        pulp.PULP_CBC_CMD(msg=False, timeLimit=time_limit_sec)
        if time_limit_sec
        else pulp.PULP_CBC_CMD(msg=False)
    )
    status = model.solve(solver)

    result = {
        "status": pulp.LpStatus[status],
        "objective": pulp.value(model.objective),
        "assignments": [],
        "unmet": {},
        "total_by_employee": {},
    }

    for e in employees:
        val = y[e].value()
        result["total_by_employee"][e] = int(round(val if val is not None else 0))

    for d in days:
        for t in tasks:
            for e in employees:
                val = x[(e, d, t)].value()
                if val is not None and val > 0.5:
                    result["assignments"].append({"employee": e, "day": d, "task": t})
            if allow_unmet_via_slack:
                uv = u[(d, t)].value()
                result["unmet"][(d, t)] = int(round(uv if uv is not None else 0))
    return result


# ====================== Excel I/O ======================


def read_bool(val, default=True):
    if pd.isna(val) or val == "":
        return default
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    if s in ("true", "t", "1", "yes", "y", "出勤可", "可"):
        return True
    if s in ("false", "f", "0", "no", "n", "出勤不可", "不可"):
        return False
    return default


def parse_input_excel(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    required = ["Employees", "Days", "Tasks", "Demand"]
    for s in required:
        if s not in xls.sheet_names:
            raise ValueError(f"必須シート '{s}' が見つかりません。")

    employees = (
        pd.read_excel(xls, "Employees")["employee"].dropna().astype(str).tolist()
    )
    days = pd.read_excel(xls, "Days")["day"].dropna().astype(str).tolist()
    tasks = pd.read_excel(xls, "Tasks")["task"].dropna().astype(str).tolist()

    # Demand
    df_demand = pd.read_excel(xls, "Demand")
    demand: Dict[Tuple[str, str], int] = {}
    for _, r in df_demand.iterrows():
        d = str(r["day"])
        t = str(r["task"])
        need = int(r["need"])
        demand[(d, t)] = need

    # Availability（任意）
    availability: Dict[Tuple[str, str], bool] = {}
    if "Availability" in xls.sheet_names:
        df_av = pd.read_excel(xls, "Availability")
        for _, r in df_av.iterrows():
            e = str(r["employee"])
            d = str(r["day"])
            available = read_bool(r.get("available", ""), default=True)
            availability[(e, d)] = available

    # Incompatibilities Global（任意）
    ig: Set[Tuple[str, str]] = set()
    if "IncompatibilitiesGlobal" in xls.sheet_names:
        df_ig = pd.read_excel(xls, "IncompatibilitiesGlobal")
        for _, r in df_ig.iterrows():
            a = str(r["employee_a"])
            b = str(r["employee_b"])
            if a and b:
                ig.add(tuple(sorted((a, b))))

    # Incompatibilities By Day（任意）
    ibd: Dict[str, Set[Tuple[str, str]]] = {}
    if "IncompatibilitiesByDay" in xls.sheet_names:
        df_ibd = pd.read_excel(xls, "IncompatibilitiesByDay")
        for _, r in df_ibd.iterrows():
            d = str(r["day"])
            a = str(r["employee_a"])
            b = str(r["employee_b"])
            if d and a and b:
                ibd.setdefault(d, set()).add(tuple(sorted((a, b))))

    # CanDo（任意）
    can_do: Dict[Tuple[str, str], bool] = {}
    if "CanDo" in xls.sheet_names:
        df_cd = pd.read_excel(xls, "CanDo")
        for _, r in df_cd.iterrows():
            e = str(r["employee"])
            t = str(r["task"])
            c = read_bool(r.get("can_do", ""), default=True)
            if not c:
                can_do[(e, t)] = False

    # MinMax（任意）
    min_shifts: Dict[str, int] = {}
    max_shifts: Dict[str, int] = {}
    if "MinMax" in xls.sheet_names:
        df_mm = pd.read_excel(xls, "MinMax")
        for _, r in df_mm.iterrows():
            e = str(r["employee"])
            if "min_shifts" in r and not pd.isna(r["min_shifts"]):
                min_shifts[e] = int(r["min_shifts"])
            if "max_shifts" in r and not pd.isna(r["max_shifts"]):
                max_shifts[e] = int(r["max_shifts"])

    # Options（任意）
    incompat_level = "day"
    weight_unmet = 1000.0
    weight_fair = 1.0
    allow_unmet_via_slack = True
    time_limit_sec = 15
    weight_pref = 1.0
    df_opt = None
    if "Options" in xls.sheet_names:
        df_opt = pd.read_excel(xls, "Options")
        if not df_opt.empty:
            r = df_opt.iloc[0]
            incompat_level = str(r.get("incompat_level", incompat_level)).strip()
            weight_unmet = float(r.get("weight_unmet", weight_unmet))
            weight_fair = float(r.get("weight_fair", weight_fair))
            allow_unmet_via_slack = read_bool(
                r.get("allow_unmet_via_slack", allow_unmet_via_slack),
                allow_unmet_via_slack,
            )
            if not pd.isna(r.get("time_limit_sec", None)):
                time_limit_sec = int(r["time_limit_sec"])
            if "weight_pref" in r and not pd.isna(r["weight_pref"]):
                weight_pref = float(r["weight_pref"])

    # ★ Priority（任意）: ワイルドカード展開（空欄→全日/全タスク）
    priority_scores: Dict[Tuple[str, str, str], float] = {}
    if "Priority" in xls.sheet_names:
        df_pr = pd.read_excel(xls, "Priority").fillna(
            {"employee": "", "day": "", "task": "", "score": 0.0}
        )
        df_pr["employee"] = df_pr["employee"].astype(str).str.strip()
        df_pr["day"] = df_pr["day"].astype(str).str.strip()
        df_pr["task"] = df_pr["task"].astype(str).str.strip()
        for _, r in df_pr.iterrows():
            e = r["employee"]
            d = r["day"]
            t = r["task"]
            try:
                s = float(r["score"])
            except:
                continue
            if not e:
                continue
            days_iter = [d] if d else days
            tasks_iter = [t] if t else tasks
            for dd in days_iter:
                for tt in tasks_iter:
                    key = (e, dd, tt)
                    priority_scores[key] = priority_scores.get(key, 0.0) + s

    return dict(
        employees=employees,
        days=days,
        tasks=tasks,
        availability=availability,
        demand=demand,
        incompatible_pairs_global=ig,
        incompatible_pairs_by_day=ibd,
        incompat_level=incompat_level,
        can_do=can_do,
        min_shifts_per_emp=min_shifts,
        max_shifts_per_emp=max_shifts,
        weight_unmet=weight_unmet,
        weight_fair=weight_fair,
        allow_unmet_via_slack=allow_unmet_via_slack,
        time_limit_sec=time_limit_sec,
        # ★
        priority_scores=priority_scores,
        weight_pref=weight_pref,
    )


def build_output_excels(result: dict):
    df_assign = pd.DataFrame(result["assignments"])
    if df_assign.empty:
        df_assign = pd.DataFrame(columns=["day", "task", "employee"])
    else:
        df_assign = df_assign[["day", "task", "employee"]].sort_values(
            ["day", "task", "employee"]
        )

    unmet_rows = [
        {"day": d, "task": t, "unmet": v} for (d, t), v in result["unmet"].items()
    ]
    df_unmet = (
        pd.DataFrame(unmet_rows).sort_values(["day", "task"])
        if unmet_rows
        else pd.DataFrame(columns=["day", "task", "unmet"])
    )

    tbe = [
        {"employee": e, "total_shifts": c}
        for e, c in result["total_by_employee"].items()
    ]
    df_total = (
        pd.DataFrame(tbe).sort_values(["employee"])
        if tbe
        else pd.DataFrame(columns=["employee", "total_shifts"])
    )

    if not df_assign.empty:
        df_pivot = df_assign.pivot_table(
            index="day",
            columns="task",
            values="employee",
            aggfunc=lambda x: ", ".join(sorted(x)),
        ).fillna("")
        df_pivot = df_pivot.reset_index()
    else:
        df_pivot = pd.DataFrame(columns=["day"])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_assign.to_excel(writer, sheet_name="Assignments", index=False)
        df_pivot.to_excel(writer, sheet_name="Pivot", index=False)
        df_unmet.to_excel(writer, sheet_name="Unmet", index=False)
        df_total.to_excel(writer, sheet_name="TotalsByEmployee", index=False)
        meta = pd.DataFrame(
            [{"status": result.get("status"), "objective": result.get("objective")}]
        )
        meta.to_excel(writer, sheet_name="Meta", index=False)

    out.seek(0)
    return out.read(), df_assign, df_unmet, df_total, df_pivot


# ====================== Streamlit UI ======================

st.set_page_config(
    page_title="シフト最適化（Excel入出力＋可視化＋優先割当）", layout="wide"
)
st.title("シフト自動割当（Excel入出力＋可視化＋優先割当）")

st.markdown(
    """
- ① テンプレートをダウンロード → Excel で編集  
- ② ここにアップロード → 最適化 → 結果Excelをダウンロード  
- ③ **Charts** タブで可視化を確認

**前提**  
- 1人は1日に最大1タスク  
- 需要未充足は許容（ペナルティ最小化、Optionsで変更可）  
- 相性NGは「同日一緒に不可（day）」または「同日同タスク不可（task）」  
- **Priority.score が大きいほど積極的に割当（weight_pref で重み付け）**
"""
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("1) テンプレート")
    tpl_bytes = make_template_excel_with_priority()
    st.download_button(
        label="Excelテンプレート（優先つき）をダウンロード",
        data=tpl_bytes,
        file_name="shift_template_aug2025_with_priority.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col2:
    st.subheader("2) 入力Excelのアップロード")
    file = st.file_uploader("テンプレート体裁のExcel(.xlsx)を選択", type=["xlsx"])

if file:
    try:
        params = parse_input_excel(file.read())
    except Exception as e:
        st.error(f"入力の解析に失敗しました: {e}")
        st.stop()

    with st.expander("読み込んだパラメタを確認する", expanded=False):
        skip_keys = (
            "availability",
            "demand",
            "incompatible_pairs_by_day",
            "can_do",
            "min_shifts_per_emp",
            "max_shifts_per_emp",
            "priority_scores",
        )
        st.write({k: v for k, v in params.items() if k not in skip_keys})
        st.write(
            "availability (一部):", dict(list(params["availability"].items())[:10])
        )
        st.write("demand (一部):", dict(list(params["demand"].items())[:10]))
        st.write(
            "incompatible_pairs_by_day:",
            {k: list(v) for k, v in params["incompatible_pairs_by_day"].items()},
        )
        st.write(
            "can_do (Falseのみ):",
            [k for k, v in params["can_do"].items() if v is False],
        )
        st.write("min_shifts_per_emp:", params["min_shifts_per_emp"])
        st.write("max_shifts_per_emp:", params["max_shifts_per_emp"])
        st.write("priority_scores（サイズ）:", len(params.get("priority_scores", {})))

    # オプション微調整（UI側でも上書き可能）
    st.subheader("3) オプション")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        incompat_level = st.selectbox(
            "相性NGのレベル",
            ["day", "task"],
            index=["day", "task"].index(params["incompat_level"]),
            help=(
                "相性NG（Incompatibilities）の適用範囲。\n"
                "- day: 同じ日に同時出勤させない（タスクが違っても不可）\n"
                "- task: 同じ日・同じタスクだけ同時不可（別タスクなら可）"
            ),
        )
    with c2:
        weight_unmet = st.number_input(
            "未充足ペナルティ（weight_unmet）",
            value=float(params["weight_unmet"]),
            step=100.0,
            min_value=0.0,
            help=(
                "需要（Demand）を1人ぶん満たせなかったときのペナルティ係数。\n"
                "数値が大きいほど『欠員ゼロ』を優先します。実務では公平性や優先よりも大きく設定するのが普通。\n"
                "目安: 1000〜5000。"
            ),
        )
    with c3:
        weight_fair = st.number_input(
            "公平性の重み（weight_fair）",
            value=float(params["weight_fair"]),
            step=0.5,
            min_value=0.0,
            help=(
                "各従業員の総勤務回数 y[e] と平均値 avg の偏差 |y[e]-avg| の合計に掛ける係数。\n"
                "大きいほど勤務回数の偏りを減らします（＝均等割に近づく）。\n"
                "目安: 0〜5。"
            ),
        )
    with c4:
        allow_unmet_via_slack = st.checkbox(
            "未充足を許容（allow_unmet_via_slack）",
            value=bool(params["allow_unmet_via_slack"]),
            help=(
                "オン: 需要を満たせない場合に『未充足（欠員）』を許容し、目的関数で最小化します。\n"
                "オフ: 常に完全充足を強制（不可能だと実行不能 Infeasible）。"
            ),
        )
    with c5:
        weight_pref = st.number_input(
            "優先割当の重み（weight_pref）",
            value=float(params.get("weight_pref", 1.0)),
            step=0.5,
            min_value=0.0,
            help=(
                "Priority シートの score をどれだけ重視するかの係数。\n"
                "目的関数では『− weight_pref × Σ(score×割当)』として効きます（最小化なので“ボーナス”）。\n"
                "未充足ペナルティより大きすぎると、欠員を残してでも好みを優先しうるので注意。\n"
                "目安: 0〜5（まずは 1.0〜2.0 から）。"
            ),
        )
    with c6:
        time_limit_sec = st.number_input(
            "ソルバー制限秒（time_limit_sec）",
            value=int(params["time_limit_sec"]),
            step=5,
            min_value=1,
            help=(
                "求解に使う最大秒数。短すぎると近似解・途中解になることがあります。\n"
                "目安: 10〜60 秒。規模が大きいほど延ばしてください。"
            ),
        )

    if st.button("4) 最適化を実行", type="primary"):
        with st.spinner("最適化中..."):
            result = solve_shift_scheduling(
                employees=params["employees"],
                days=params["days"],
                tasks=params["tasks"],
                availability=params["availability"],
                demand=params["demand"],
                incompatible_pairs_global=params["incompatible_pairs_global"],
                incompatible_pairs_by_day=params["incompatible_pairs_by_day"],
                incompat_level=incompat_level,
                can_do=params["can_do"],
                min_shifts_per_emp=params["min_shifts_per_emp"],
                max_shifts_per_emp=params["max_shifts_per_emp"],
                weight_unmet=weight_unmet,
                weight_fair=weight_fair,
                allow_unmet_via_slack=allow_unmet_via_slack,
                time_limit_sec=time_limit_sec,
                # ★
                priority_scores=params.get("priority_scores", {}),
                weight_pref=weight_pref,
            )

        st.success(
            f"ステータス: {result['status']} / 目的関数: {result['objective']:.2f}"
        )

        # Excel出力とテーブル
        out_bytes, df_assign, df_unmet, df_total, df_pivot = build_output_excels(result)

        tab_tbl, tab_charts = st.tabs(["🧾 Tables", "📊 Charts"])

        with tab_tbl:
            st.markdown("#### ピボット（日×タスク）")
            st.dataframe(df_pivot, use_container_width=True)

            st.markdown("#### 未充足（Unmet）")
            st.dataframe(df_unmet, use_container_width=True)

            st.markdown("#### 従業員別合計（TotalsByEmployee）")
            st.dataframe(df_total, use_container_width=True)

            st.download_button(
                label="結果Excelをダウンロード",
                data=out_bytes,
                file_name="shift_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab_charts:
            # ---- Coverage heatmap ----
            def build_coverage_df():
                if df_assign.empty:
                    base = pd.DataFrame(
                        [(d, t) for d in params["days"] for t in params["tasks"]],
                        columns=["day", "task"],
                    )
                    base["assigned"] = 0
                else:
                    cnt = (
                        df_assign.groupby(["day", "task"])
                        .size()
                        .reset_index(name="assigned")
                    )
                    base = pd.DataFrame(
                        [(d, t) for d in params["days"] for t in params["tasks"]],
                        columns=["day", "task"],
                    )
                    base = base.merge(cnt, on=["day", "task"], how="left").fillna(
                        {"assigned": 0}
                    )
                need_rows = [
                    {"day": d, "task": t, "need": int(params["demand"].get((d, t), 0))}
                    for d in params["days"]
                    for t in params["tasks"]
                ]
                need_df = pd.DataFrame(need_rows)
                cov = base.merge(need_df, on=["day", "task"], how="left").fillna(
                    {"need": 0}
                )
                cov["gap"] = cov["need"] - cov["assigned"]
                return cov

            cov = build_coverage_df()

            st.markdown("##### Unmet heatmap（未充足）")
            if not cov.empty:
                chart_unmet = (
                    alt.Chart(cov)
                    .mark_rect()
                    .encode(
                        y=alt.X("task:N", title="Task"),
                        x=alt.Y("day:N", title="Day"),
                        color=alt.Color(
                            "gap:Q",
                            title="Unmet (need-assigned)",
                            scale=alt.Scale(scheme="blues"),
                        ),
                        tooltip=["day:N", "task:N", "need:Q", "assigned:Q", "gap:Q"],
                    )
                    .properties(width="container", height=280)
                )
                st.altair_chart(chart_unmet, use_container_width=True)

            st.markdown("##### Workload by employee（総勤務回数）")
            if not df_total.empty:
                chart_work = (
                    alt.Chart(df_total)
                    .mark_bar()
                    .encode(
                        y=alt.X(
                            "employee:N",
                            sort=alt.SortField("employee", order="ascending"),
                        ),
                        x=alt.Y("total_shifts:Q"),
                        tooltip=["employee:N", "total_shifts:Q"],
                    )
                    .properties(width="container", height=320)
                )
                st.altair_chart(chart_work, use_container_width=True)

            st.markdown("##### Assignment grid（誰が入っているか）")
            if not df_assign.empty:
                label_df = (
                    df_assign.groupby(["day", "task"])
                    .agg(employee_list=("employee", lambda x: ", ".join(sorted(x))))
                    .reset_index()
                )
                chart_grid = (
                    alt.Chart(label_df)
                    .mark_rect(stroke="gray")
                    .encode(
                        y=alt.X("task:N", title="Task"),
                        x=alt.Y("day:N", title="Day"),
                        tooltip=["day:N", "task:N", "employee_list:N"],
                    )
                    .properties(width="container", height=280)
                )
                text = (
                    alt.Chart(label_df)
                    .mark_text(baseline="middle", align="center", dy=0, size=12)
                    .encode(y="task:N", x="day:N", text="employee_list:N")
                )
                st.altair_chart(chart_grid + text, use_container_width=True)

            # ★ Priority score heatmap（割当の好みスコア合計）
            if not df_assign.empty and params.get("priority_scores"):
                scored = []
                for _, r in df_assign.iterrows():
                    e, d, t = r["employee"], r["day"], r["task"]
                    scored.append(
                        {
                            "day": d,
                            "task": t,
                            "score": params["priority_scores"].get((e, d, t), 0.0),
                        }
                    )
                df_sc = (
                    pd.DataFrame(scored)
                    .groupby(["day", "task"], as_index=False)["score"]
                    .sum()
                )
                st.markdown("##### Priority score heatmap（割当の好みスコア合計）")
                st.altair_chart(
                    alt.Chart(df_sc)
                    .mark_rect()
                    .encode(
                        y="task:N",
                        x="day:N",
                        color=alt.Color(
                            "score:Q", title="Score", scale=alt.Scale(scheme="greens")
                        ),
                        tooltip=["day:N", "task:N", "score:Q"],
                    )
                    .properties(height=280, width="container"),
                    use_container_width=True,
                )

else:
    st.info("テンプレートをダウンロードし、Excelを編集してアップロードしてください。")
