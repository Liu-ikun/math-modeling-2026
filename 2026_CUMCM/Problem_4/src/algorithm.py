# -*- coding: utf-8 -*-
"""问题4 编排层（改进版·核心闭环）
第一部分 Q4-2: 全年一次性全局LP(完美信息) + 无储能基线 + 逐日短视对照 + α_d电价结构分析;
第二部分 Q4-3: MPC滚动(0/6/12/18) + 实际执行结算 + 时移对齐 + 完美预报自洽;
第三部分: 侧面验证 + 附件5模板回填(格式零改动) + 覆盖率门禁 + 完整存档。
修复说明:
 (1) 原文件 verify_fill42 函数体内误嵌套了 fill_wide43/verify_fill43/write_full42,
     导致模块级缺少这些名字、main.py 导入失败; 本版全部提升到模块级并去重;
 (2) 删除改进版不需要的 robust_adj(报童校正)与 build_virtual_fc(虚拟发布),
     二者仅服务于已裁剪的频率实验与qcheck链。
路径: 2026_CUMCM/Problem_4/src/algorithm.py"""
import re
import shutil

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook

from model import (DT, E0, E_MAX, E_MIN, ETA_CH, ETA_DIS, build_stage, solve,
                   solve_global, solve_stage)

TOL = 1e-6
UPDATES = (0, 6, 12, 18)
DATE_KEYS = ["2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"]
WIN_START = "2025-02-01"            # 报告窗口(1月仅作SOC预热)
Q2_PERFECT_COST = 13751266.38       # 问题2完美信息全局最优(回归门禁基准)
Q3_B_COST = 19674667.34             # 问题3提交策略(附件1电价) — 固定vs波动对照
Q3_A_COST = 19856349.12             # 问题3朴素日前(附件1电价)
BREAKEVEN_RATIO = 1.0 / (ETA_CH * ETA_DIS)     # 充放循环保本峰谷价比 = 1.2346

SLOT_LABELS = ["10:00-10:10", "12:00-12:10", "14:00-14:10",
               "16:00-16:10", "18:00-18:10", "20:00-20:10"]
SLOT_HOURS = [10, 12, 14, 16, 18, 20]
BLOCK_LABELS = ["0:00-4:00", "4:00-8:00", "8:00-12:00",
                "12:00-16:00", "16:00-20:00", "20:00-24:00"]
BLOCK_RANGE = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
_LAB_RE = re.compile(r"(\d{1,2}):(\d{1,2})\s*-\s*(\d{1,2}):(\d{1,2})")
SHIFT_BY_H = {0: 0, 6: 0, 12: 0, 18: 0}


# ---------------------------------------------------------------- 通用工具
def _pick(vs):
    return [v.value() or 0.0 for v in vs]


def _norm(v):
    return re.sub(r"\s+", "", str(v)) if v is not None else ""


def _norm_date(v):
    if v is None:
        return None
    ts = pd.to_datetime(v, errors="coerce")
    return None if pd.isna(ts) else ts.strftime("%Y-%m-%d")


def _parse_lab(s):
    """时段标签 -> t(1..144), 以区间右端点定t(与Q1/Q2/Q3一致); '0:00'右端点视为24:00"""
    m = _LAB_RE.search(str(s))
    if not m:
        return None
    h1, m1, h2, m2 = map(int, m.groups())
    end = h2 * 60 + m2 or 1440
    t = end // 10
    return t if 1 <= t <= 144 else None


def flatten(M):
    """(D,144)矩阵 -> 长度 D*144 的平铺列表(按日顺序)"""
    return [x for row in M for x in row]


# ================================================== 第一部分: Q4-2 全局LP
def alpha_analysis(c4, c1, L, PV, dates):
    """附件4电价结构分解: c_{d,t} ≈ α_d × 附件1分时形状 + 噪声
    输出逐日表(α_d/日均价/日内峰谷比/单位套利利润/日均净负荷) + 汇总指标"""
    A4, A1 = np.asarray(c4, float), np.asarray(c1, float)
    base = float(A1.mean())
    alpha = A4.mean(1) / base
    ratio = A4.max(1) / np.maximum(A4.min(1), 1e-9)
    profit = A4.max(1) - A4.min(1) / (ETA_CH * ETA_DIS)
    nl = (np.asarray(L, float) - np.asarray(PV, float)).mean(1)
    pvd = np.asarray(PV, float).mean(1)
    months = np.array([int(d[5:7]) for d in dates])
    df = pd.DataFrame({"日期": dates, "月份": months, "α_d": np.round(alpha, 4),
                       "日均电价": np.round(A4.mean(1), 6),
                       "日内峰谷比": np.round(ratio, 4),
                       "单位套利利润(元/kWh)": np.round(profit, 4),
                       "日均净负荷kW": np.round(nl, 1),
                       "日均光伏kW": np.round(pvd, 1)})
    info = {
        "附件1年均电价": round(base, 6),
        "附件4年均电价": round(float(A4.mean()), 6),
        "α_d范围": (round(float(alpha.min()), 4), round(float(alpha.max()), 4)),
        "α_d极差比": round(float(alpha.max() / max(alpha.min(), 1e-9)), 4),
        "保本峰谷价比": round(BREAKEVEN_RATIO, 4),
        "逐日峰谷比min/中位/max": (round(float(ratio.min()), 3),
                                 round(float(np.median(ratio)), 3),
                                 round(float(ratio.max()), 3)),
        "峰谷比超保本线天数": f"{int((ratio > BREAKEVEN_RATIO).sum())}/{len(ratio)}",
        "套利利润为正天数": f"{int((profit > 0).sum())}/{len(profit)}",
        "corr(α_d,日均净负荷)": round(float(np.corrcoef(alpha, nl)[0, 1]), 4),
        "corr(α_d,日均光伏)": round(float(np.corrcoef(alpha, pvd)[0, 1]), 4),
        "月度α_d": {int(m): round(float(alpha[months == m].mean()), 4)
                  for m in range(1, 13)},
    }
    return df, info, alpha


def solve_myopic(c4, L, PV, method="highs-ds"):
    """对照策略C: 逐日独立全局LP + SOC链式递推(无末态估值) -> 每天午夜放空到E_MIN的短视策略。
    与 solve_global(全年一次) 的差值 = 跨日套利价值。
    附件1(电价逐日重复)下该值应≈0; 附件4(α_d波动)下应显著>0 —— Q4核心新增结论"""
    D = len(L)
    g, eme, spv, pch, pdis, E = [], [], [], [], [], [E0]
    e_cur, obj = E0, 0.0
    for d in range(D):
        r = solve_global(c4[d], L[d], PV[d], storage=True, e_init=e_cur, method=method)
        g += r["G"]; eme += r["Eme"]; spv += r["Sp"]
        pch += r["Pch"]; pdis += r["Pdis"]
        E += r["E"][1:]
        e_cur = r["E"][-1]
        obj += r["obj"]
    simult = sum(1 for n in range(len(g)) if pch[n] > TOL and pdis[n] > TOL)
    return dict(status="Optimal(逐日短视)", obj=obj, D=D, N=len(g), G=g, Eme=eme,
                Sp=spv, Pch=pch, Pdis=pdis, E=E, n_simult=simult,
                buy=sum(g) * DT, emer=sum(eme) * DT, spill=sum(spv) * DT)


def validate_global(res, L, PV):
    """Q4-2侧面验证: 功率平衡双口径残差/SOC在界/同时充放/紧急购电(完美信息应=0)/弃光"""
    T = 144
    errs = []
    for n in range(res["N"]):
        d, t = divmod(n, T)
        errs.append(abs(res["G"][n] + res["Eme"][n] + PV[d][t] + res["Pdis"][n]
                        - L[d][t] - res["Pch"][n] - res["Sp"][n]))
    bal = max(errs) if errs else 0.0
    scale = max(1e-9, max(max(r) for r in L))
    inb = res["E"] is None or all(E_MIN - 1e-6 <= x <= E_MAX + 1e-6 for x in res["E"])
    return {"功率平衡最大绝对残差kW": bal,
            "功率平衡相对残差": bal / scale,
            "SOC全程在界": inb,
            "同时充放时段数": res["n_simult"],
            "紧急购电总量kWh": round(res["emer"], 4),
            "弃光总量kWh": round(res["spill"], 2),
            "购电总量kWh": round(res["buy"], 2)}


def daily_agg_global(res, c4, dates, alpha=None):
    T = 144
    rows = []
    for d in range(res["D"]):
        s = slice(d * T, (d + 1) * T)
        g, sp = res["G"][s], res["Sp"][s]
        pch, pdis, eme = res["Pch"][s], res["Pdis"][s], res["Eme"][s]
        rows.append([str(dates[d]),
                     round(sum(c4[d][t] * g[t] * DT for t in range(T)), 2),
                     round(sum(g) * DT, 2), round(sum(pch) * DT, 2),
                     round(sum(pdis) * DT, 2), round(sum(sp) * DT, 2),
                     round(sum(eme) * DT, 4),
                     round(float(alpha[d]), 4) if alpha is not None else None])
    return pd.DataFrame(rows, columns=["日期", "购电费(元)", "购电量kWh", "充电量kWh",
                                       "放电量kWh", "弃光电量kWh", "紧急购电kWh", "α_d"])


def economy_check_global(res_g, res_n, res_m):
    out = {"有储能全局优化(元)": round(res_g["obj"], 2),
           "无储能基线(元)": round(res_n["obj"], 2),
           "逐日短视对照(元)": round(res_m["obj"], 2),
           "储能节省(元)": round(res_n["obj"] - res_g["obj"], 2),
           "储能节省比例": round((res_n["obj"] - res_g["obj"]) / res_n["obj"], 4)
           if res_n["obj"] else None,
           "跨日套利价值(元)": round(res_m["obj"] - res_g["obj"], 2),
           "跨日套利价值占全局成本比": round(
               (res_m["obj"] - res_g["obj"]) / res_g["obj"], 5) if res_g["obj"] else None}
    return out


def win_sum(daily, col, start=WIN_START):
    m = daily["日期"] >= start
    return float(daily.loc[m, col].sum())


def long_table_g(res, c4, dates, start=None):
    T = 144
    rows = []
    for d in range(res["D"]):
        if start and str(dates[d]) < start:
            continue
        for t in range(T):
            n = d * T + t
            rows.append([str(dates[d]),
                         f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
                         c4[d][t], res["G"][n], res["Pch"][n], res["Pdis"][n],
                         res["Sp"][n], res["Eme"][n],
                         res["E"][n + 1] if res["E"] else None])
    return pd.DataFrame(rows, columns=["日期", "时刻", "电价(元/kWh)", "购电功率kW",
                                       "充电kW", "放电kW", "弃光kW", "紧急购电kW",
                                       "时段末储电量kWh"])


def _day_tables_g(res, c_d, d):
    """第d天的表1/表2/表3数据; c_d为该日电价向量(附件4逐日不同, 不能再用统一c_day)"""
    T = 144
    n0 = d * T
    g, pch, pdis = res["G"][n0:n0 + T], res["Pch"][n0:n0 + T], res["Pdis"][n0:n0 + T]
    eme = res["Eme"][n0:n0 + T]
    e = res["E"][n0:n0 + T + 1] if res["E"] else [None] * (T + 1)
    buy = [x * DT for x in g]
    return dict(slots={lab: buy[h * 6] for lab, h in zip(SLOT_LABELS, SLOT_HOURS)},
                buy=sum(buy), cost=sum(c_d[t] * g[t] * DT for t in range(T)),
                blocks={lab: (sum(pch[a:b]) * DT, sum(pdis[a:b]) * DT)
                        for lab, (a, b) in zip(BLOCK_LABELS, BLOCK_RANGE)},
                e0=e[0], e24=e[-1],
                t3=[(f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
                     round(eme[t] * DT, 4)) for t in range(T)])


def make_date_tables_g(res, c4, dates):
    return {k: _day_tables_g(res, c4[dates.index(k)], dates.index(k))
            for k in DATE_KEYS if k in dates}


# ================================================== 第二部分: Q4-3 MPC
def set_fc_shift(shift_map):
    SHIFT_BY_H.update(shift_map)


def calibrate_shift(fcmap, PV, dates, lo=-8, hi=8):
    """宽域RMSE自动对齐: 逐发布时扫描整点位移s, 取 fc(lead k) - act(小时 H+k-1+s) 的RMSE最小者。
    RMSE对错位敏感(爬坡处数千kW)而对系统性偏差钝感; 2%容差内取最小s。
    与问题3同一判据 => s 可直接沿用(预期 (0,0,0,+1))"""
    best_map = {}
    for H in (0, 6, 12, 18):
        cand = []
        for s in range(lo, hi + 1):
            se, n = 0.0, 0
            for di, date in enumerate(dates):
                leads = fcmap[date].get(H)
                if not leads:
                    continue
                for k in range(1, 25):
                    hh = H + k - 1 + s
                    if 0 <= hh <= 23:
                        act = sum(PV[di][hh * 6:hh * 6 + 6]) / 6.0
                        se += (leads[k - 1] - act) ** 2
                        n += 1
            rmse = (se / n) ** 0.5 if n else 9e9
            cand.append((rmse, s))
        best = min(r for r, _ in cand)
        s_sel = min(s for r, s in cand if r <= best * 1.02)
        best_map[H] = s_sel
        print(f"[自动对齐] 发布时{H}: 最优shift={s_sel:+d}h RMSE={best:.1f} kW "
              f"(s=0时RMSE={next(r for r, s in cand if s == 0):.1f})")
    return best_map


def verify_alignment(fcmap, PV, dates):
    """日出边沿门禁(无参): 应用校准位移后, fc首个>50kW的lead映射小时与实际首个>50kW小时
    的中位差应≈0; 不为0则口径未定, 拒绝进入求解"""
    diffs = []
    for di, date in enumerate(dates):
        leads = fcmap[date].get(0)
        if not leads:
            continue
        k1 = next((k for k in range(1, 25) if leads[k - 1] > 50), None)
        h1 = next((h for h in range(24)
                   if sum(PV[di][h * 6:h * 6 + 6]) / 6.0 > 50), None)
        if k1 is not None and h1 is not None:
            diffs.append((k1 - 1 + SHIFT_BY_H[0]) - h1)
    diffs.sort()
    med = diffs[len(diffs) // 2] if diffs else 0
    print(f"[对齐门禁] 校准后日出边沿中位差 {med:+d} h (应≈0)")
    return med


def fc_series(fcmap, date, H, t0=0, vmap=None, fadj=None, timing="ceil", interp=False):
    """发布时H可用的10min光伏预报序列(覆盖时段 t0+1..144)。

    时移口径 timing:
      'ceil'  k = ceil(t/6) - h - s  —— 正确口径。数据语义为 lead k ↔ 小时 H+k-1+s,
              而时段t所属小时 = floor((t-1)/6) = ceil(t/6)-1, 故必须用ceil才能逐时段命中;
              同时消除 t%6!=0 时 k 偏小1 造成的'每小时5/6时段滞后1小时'与阶段起点零预报真空。
      'floor' k = t//6 - h - s       —— 问题3旧口径, 仅作对照链(Bleg)量化时移影响。
    多发布回退: 新->旧, 保证任一时段都有可用预报(杜绝零预报真空)。
    interp=True: 同一发布内相邻lead线性插值降尺度(团队数据分析: MAE 226.0 vs ZOH 361.9 kW);
                 严禁跨发布时点插值(问题3踩过的processed宽表污染坑)。"""
    issues = fcmap[date]
    virt = vmap is not None and H in vmap.get(date, {})
    src_map = dict(issues)
    if virt:
        src_map.update(vmap.get(date, {}))
    order = [(h, SHIFT_BY_H[h] if h in SHIFT_BY_H else 0)
             for h in sorted(src_map) if h <= H][::-1]
    out = []
    for t in range(t0 + 1, 145):
        base = (t + 5) // 6 if timing == "ceil" else t // 6
        v, used, uh, row = 0.0, None, None, None
        for h, s in order:
            k = base - h - s
            if 1 <= k <= 24 and k - 1 < len(src_map[h]):
                v, used, uh, row = float(src_map[h][k - 1]), (h, k), h, src_map[h]
                break
        if used is not None and interp and timing == "ceil":
            delta = ((t - 1) % 6 + 0.5) / 6.0 - 0.5      # 时段中心相对所属小时中心的偏移
            if delta >= 0 and k < len(row):
                v += delta * (float(row[k]) - v)          # 向 lead k+1 插值
            elif delta < 0 and k >= 2:
                v += (-delta) * (float(row[k - 2]) - v)   # 向 lead k-1 插值
        if fadj and used:
            key = ("v", k) if virt else used
            v = max(0.0, v - fadj.get(key, 0.0))
        out.append(max(0.0, v))
    return out


def diagnose_timing(fcmap, PV, dates, timings=("ceil", "floor"), shifts=(-1, 0, 1)):
    """D1诊断: 实测 fc_series 输出相对实际PV的等效时移(小时)。
    对每个发布时H与每种口径, 计算 RMSE( fc[t] , PV[t-6*sh] ) 在 sh∈shifts 上的最小值;
    sh=0 => 时移正确; sh=+1 => 预报滞后1小时(白天爬坡段将被迫用5c紧急购电买单)。
    仅统计白天(实际或预报>20kW)时段, 避免夜间零光伏造成无区分度。"""
    report = {}
    for timing in timings:
        for H in (0, 6, 12, 18):
            se = {sh: 0.0 for sh in shifts}
            n = {sh: 0 for sh in shifts}
            for di, date in enumerate(dates):
                if H not in fcmap[date]:
                    continue
                pvfc = fc_series(fcmap, date, H, H * 6, timing=timing)
                for j, fv in enumerate(pvfc):
                    t = H * 6 + 1 + j
                    for sh in shifts:
                        i0 = t - 1 - 6 * sh
                        if 0 <= i0 < 144:
                            av = PV[di][i0]
                            if fv > 20.0 or av > 20.0:
                                se[sh] += (fv - av) ** 2
                                n[sh] += 1
            rm = {sh: (se[sh] / n[sh]) ** 0.5 if n[sh] else float("nan")
                  for sh in shifts}
            best = min((r for r in rm.values() if r == r))
            lag = min(sh for sh, r in rm.items() if r == r and r <= best * 1.02)
            report[(timing, H)] = (rm, lag)
            print(f"[时移诊断] timing={timing:5s} 发布时{H:2d}: "
                  + " ".join(f"sh{sh:+d}={rm[sh]:7.1f}" for sh in shifts)
                  + f" kW => 等效滞后 {lag:+d} h {'✓正确' if lag == 0 else '✗存在时移'}")
    return report


def build_fc_stats(fcmap, PV, dates):
    """逐(发布时H, lead k)统计预报误差 err = fc - act (act取小时 H+k-1+s 的实际均值)"""
    import math
    s1, cnt, errs = {}, {}, {}
    for di, date in enumerate(dates):
        for H, leads in fcmap[date].items():
            for k in range(1, 25):
                hh = H + k - 1 + SHIFT_BY_H[H]
                if 0 <= hh <= 23:
                    e = leads[k - 1] - sum(PV[di][hh * 6:hh * 6 + 6]) / 6.0
                    s1[(H, k)] = s1.get((H, k), 0.0) + e
                    cnt[(H, k)] = cnt.get((H, k), 0) + 1
                    errs.setdefault((H, k), []).append(e)
    bias = {q: s1[q] / cnt[q] for q in cnt}
    std = {q: math.sqrt(max(0.0, sum(x * x for x in errs[q]) / cnt[q] - bias[q] ** 2))
           for q in cnt}
    return bias, std, errs


def build_perfect_fc(PV, dates, timing="ceil"):
    """完美预报fcmap(仅四个真实发布时): 用于BP链 —— 理论上MPC在完美预报下
    G_adj≡G_plan(ΔG=0)、紧急购电≡0, 其全年成本应严格等于Q4-2的逐日短视LP,
    从而交叉验证两套代码(Q4-2全局LP / Q4-3阶段LP)的一致性"""
    pmap = {}
    for di, date in enumerate(dates):
        vh = {}
        for H in UPDATES:
            s = SHIFT_BY_H[H]
            row = []
            for k in range(1, 25):
                hh = H + k - 1 + s if timing == "ceil" else H + k + s
                if not 0 <= hh <= 23:
                    break
                row.append(sum(PV[di][hh * 6:hh * 6 + 6]) / 6.0)
            if row:
                vh[H] = row
        pmap[date] = vh
    return pmap


def run_day(c_d, L_d, PVact, fcmap, date, e_init, updates=UPDATES, solver="highs",
            vmap=None, fadj=None, timing="ceil", interp=False):
    """单日MPC滚动决策 + 实际执行结算。c_d为该日144个时段的附件4电价(逐日不同)。
    结算: c·G_plan + 1.5c·ΔG⁺ - 0.5c·ΔG⁻ + 5c·Eme, 四项分别累计供成本分解"""
    T = 144
    plan = [0.0] * T
    gadj = [0.0] * T
    pch = [0.0] * T
    pdis = [0.0] * T
    E = [e_init] * (T + 1)
    obj_sum = 0.0
    for i, H in enumerate(updates):
        t0 = H * 6
        t1 = updates[i + 1] * 6 if i + 1 < len(updates) else T
        mode = "plan" if i == 0 else "adj"
        pvfc = fc_series(fcmap, date, H, t0, vmap, fadj, timing, interp)
        plan_arg = None if mode == "plan" else plan[t0:]
        if solver == "highs":
            s = solve_stage(c_d[t0:], L_d[t0:], pvfc, plan=plan_arg,
                            e_init=E[t0], mode=mode)
            g, pc, pd_, e = s["G"], s["Pch"], s["Pdis"], s["E"]
            obj_sum += s["obj"]
        else:
            prob, v = build_stage(c_d[t0:], L_d[t0:], pvfc, plan=plan_arg,
                                  e_init=E[t0], mode=mode)
            st, ov = solve(prob, time_limit=120)
            g, pc, pd_, e = (_pick(v[k]) for k in ("G", "Pch", "Pdis", "E"))
            obj_sum += ov
        if i == 0:
            plan[:] = g
        gadj[t0:t1] = g[:t1 - t0]
        pch[t0:t1] = pc[:t1 - t0]
        pdis[t0:t1] = pd_[:t1 - t0]
        E[t0 + 1:t1 + 1] = e[1:t1 - t0 + 1]
    emer = [0.0] * T
    sp = [0.0] * T
    c_plan = c_up = c_dn = c_eme = 0.0
    for t in range(T):
        net = gadj[t] + PVact[t] + pdis[t] - L_d[t] - pch[t]
        emer[t] = max(-net, 0.0)          # 缺口 -> 紧急购电(5c)
        sp[t] = max(net, 0.0)             # 盈余 -> 弃光(零成本)
        E[t + 1] = E[t] + ETA_CH * pch[t] * DT - pdis[t] * DT / ETA_DIS
        dp = max(gadj[t] - plan[t], 0.0)
        dm = max(plan[t] - gadj[t], 0.0)
        c_plan += c_d[t] * plan[t] * DT
        c_up += 1.5 * c_d[t] * dp * DT
        c_dn += -0.5 * c_d[t] * dm * DT
        c_eme += 5.0 * c_d[t] * emer[t] * DT
    cost = c_plan + c_up + c_dn + c_eme
    simult = sum(1 for t in range(T) if pch[t] > TOL and pdis[t] > TOL)
    return dict(plan=plan, gadj=gadj, pch=pch, pdis=pdis, E=E, emer=emer, sp=sp,
                cost=cost, cost_plan=c_plan, cost_up=c_up, cost_dn=c_dn,
                cost_eme=c_eme, simult=simult, buy=sum(gadj) * DT,
                plan_kwh=sum(plan) * DT, emer_kwh=sum(emer) * DT,
                spill=sum(sp) * DT, ch_kwh=sum(pch) * DT,
                dis_kwh=sum(pdis) * DT, lp_obj=obj_sum)


def run_chain(c4, L, PV, fcmap, dates, idxs, updates, solver="highs", vmap=None,
              fadj=None, timing="ceil", interp=False):
    """一条独立SOC链: 链内按日串行递推e_init(不可并行), 链间可多进程并发"""
    days, e_cur = [], E0
    for d in idxs:
        r = run_day(c4[d], L[d], PV[d], fcmap, dates[d], e_cur, updates, solver,
                    vmap, fadj, timing, interp)
        e_cur = r["E"][-1]
        days.append(r)
    return days


def chain_cost(days, dates=None, start=None):
    if start is None or dates is None:
        return sum(r["cost"] for r in days)
    return sum(r["cost"] for r, d in zip(days, dates) if d >= start)


def validate3(days, L, PVact, fcmap, dates, timing="ceil"):
    """Q4-3侧面验证: 执行平衡双口径残差/SOC在界/同时充放/紧急购电与预报偏差集中度"""
    errs, e_all, simult = [], [], 0
    for di, rec in enumerate(days):
        for t in range(144):
            errs.append(abs(rec["gadj"][t] + rec["emer"][t] + PVact[di][t]
                            + rec["pdis"][t] - L[di][t] - rec["pch"][t] - rec["sp"][t]))
        e_all += rec["E"]
        simult += rec["simult"]
    pairs = []
    for di, (rec, date) in enumerate(zip(days, dates)):
        fc_used = []
        for i, H in enumerate(UPDATES):
            t0 = H * 6
            t1 = UPDATES[i + 1] * 6 if i + 1 < len(UPDATES) else 144
            fc_used += fc_series(fcmap, date, H, t0, timing=timing)[:t1 - t0]
        for t in range(144):
            pairs.append((abs(fc_used[t] - PVact[di][t]), rec["emer"][t] * DT))
    pairs.sort(key=lambda x: -x[0])
    tot = sum(p[1] for p in pairs) or 1.0
    top20 = sum(p[1] for p in pairs[:len(pairs) // 5]) / tot
    scale = max(1e-9, max(max(r) for r in L))
    return {"执行平衡最大绝对残差kW": max(errs) if errs else 0.0,
            "执行平衡相对残差": (max(errs) if errs else 0.0) / scale,
            "SOC全程在界": all(E_MIN - 1e-6 <= x <= E_MAX + 1e-6 for x in e_all),
            "同时充放时段数": simult,
            "紧急购电总量kWh": round(sum(r["emer_kwh"] for r in days), 2),
            "紧急购电时段数": sum(1 for r in days for x in r["emer"] if x > TOL),
            "Top20%偏差时段承载紧急电量占比": round(top20, 4)}


def make_date_tables(rec, c_d, d):
    T = 144
    buy_p = [x * DT for x in rec["plan"]]
    buy_a = [x * DT for x in rec["gadj"]]
    return dict(slots_p={lab: buy_p[h * 6] for lab, h in zip(SLOT_LABELS, SLOT_HOURS)},
                slots_a={lab: buy_a[h * 6] for lab, h in zip(SLOT_LABELS, SLOT_HOURS)},
                totals={"buy_p": sum(buy_p), "buy_a": sum(buy_a), "cost": rec["cost"]},
                blocks={lab: (sum(rec["pch"][a:b]) * DT, sum(rec["pdis"][a:b]) * DT)
                        for lab, (a, b) in zip(BLOCK_LABELS, BLOCK_RANGE)},
                e0=rec["E"][0], e24=rec["E"][-1],
                t3=[(f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
                     round(rec["emer"][t] * DT, 4)) for t in range(T)])


# ================================================== 第三部分: 模板回填
def fill_wide42(template, out, res, dates, c4):
    """result4-2模板回填(复制模板按位置写值, 格式零改动), 沿用问题2已验证的双模式识别:
    第1行含>=100个时段标签 => 网格模式(日期行×144时段列 + 右侧购电量/购电费汇总列);
    否则竖排块模式: '充放'=6个4h块+0:00/24:00储电量, '紧急'=稀疏列表(完美信息下写'无/0')"""
    shutil.copyfile(template, out)
    wb = load_workbook(out)
    T = 144
    dset = set(dates)
    report = []
    for ws in wb.worksheets:
        title = ws.title
        nw = 0
        cols, csum = {}, {}
        for c in range(2, ws.max_column + 1):
            hv = ws.cell(row=1, column=c).value
            p = _parse_lab(hv)
            if p:
                cols[p] = c
            elif hv is not None:
                k = _norm(hv)
                if "购电费" in k:
                    csum["cost"] = c
                elif "购电量" in k:
                    csum["buy"] = c
        grid = len(cols) >= 100
        if grid:
            for r in range(2, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d not in dset:
                    continue
                di = dates.index(d)
                tb = _day_tables_g(res, c4[di], di)
                g = res["G"][di * T: di * T + T]
                for t, c in cols.items():
                    ws.cell(row=r, column=c).value = round(g[t - 1] * DT, 4)
                    nw += 1
                if "buy" in csum:
                    ws.cell(row=r, column=csum["buy"]).value = round(tb["buy"], 4)
                    nw += 1
                if "cost" in csum:
                    ws.cell(row=r, column=csum["cost"]).value = round(tb["cost"], 4)
                    nw += 1
        elif "紧急" in title:
            cur = None
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur = d
                    ws.cell(row=r, column=2).value = "无"
                    ws.cell(row=r, column=3).value = 0.0
                    nw += 2
                    continue
                if cur is None:
                    continue
                p = _parse_lab(ws.cell(row=r, column=2).value)
                if p:
                    n0 = dates.index(cur) * T
                    ws.cell(row=r, column=3).value = round(res["Eme"][n0 + p - 1] * DT, 4)
                    nw += 1
        elif "充放" in title:
            cur, mom = None, 0
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur, mom = d, 0
                if cur is None:
                    continue
                di = dates.index(cur)
                tb = _day_tables_g(res, c4[di], di)
                lab = _norm(ws.cell(row=r, column=2).value)
                if lab in BLOCK_LABELS:
                    ch, dis = tb["blocks"][lab]
                    ws.cell(row=r, column=3).value = round(ch, 4)
                    ws.cell(row=r, column=4).value = round(dis, 4)
                    nw += 2
                ev = ws.cell(row=r, column=5).value
                en = _norm(ev)
                if ev is not None and mom < 2 \
                        and en not in ("时刻", "⋮", "︙", "…", "...", "‥", ":"):
                    mom += 1
                    ws.cell(row=r, column=6).value = round(
                        tb["e0"] if mom == 1 else tb["e24"], 4)
                    nw += 1
        miss = [t for t in range(1, 145) if t not in cols] if grid else []
        report.append(f"{title}: 写入{nw}格({'网格' if grid else '竖排'})"
                      + (f" 缺失时段{miss[:5]}" if miss else ""))
        print(f"[result4-2/{title}] {report[-1]}")
    wb.save(out)
    return report

def verify_fill42(out, dates):
    """覆盖率门禁: 读回输出逐格检查应有值非空(网格sheet全格 + 竖排块/时刻/紧急日期行)"""
    wb = load_workbook(out)
    dset = set(dates)
    problems = []
    for ws in wb.worksheets:
        title = ws.title
        cols, csum = {}, []
        for c in range(2, ws.max_column + 1):
            hv = ws.cell(row=1, column=c).value
            p = _parse_lab(hv)
            if p:
                cols[p] = c
            elif hv is not None and ("购电费" in _norm(hv) or "购电量" in _norm(hv)):
                csum.append(c)
        if len(cols) >= 100:
            for r in range(2, ws.max_row + 1):
                if _norm_date(ws.cell(row=r, column=1).value) not in dset:
                    continue
                for c in list(cols.values()) + csum:
                    if ws.cell(row=r, column=c).value is None:
                        problems.append((title, r, c))
        else:
            cur = None
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur = d
                if cur is None:
                    continue
                if "充放" in title:
                    lab = _norm(ws.cell(row=r, column=2).value)
                    if lab in BLOCK_LABELS and (ws.cell(row=r, column=3).value is None
                                                or ws.cell(row=r, column=4).value is None):
                        problems.append((title, r, 3))
                    ev = ws.cell(row=r, column=5).value
                    en = _norm(ev)
                    if en in ("⋮", "︙", "…", "...", "‥"):
                        if ws.cell(row=r, column=6).value is not None \
                                or ws.cell(row=r, column=3).value is not None:
                            problems.append((title, r, 6))
                    elif ev is not None and en != "时刻" \
                            and ws.cell(row=r, column=6).value is None:
                        problems.append((title, r, 6))
                else:
                    if d in dset and ws.cell(row=r, column=3).value is None:
                        problems.append((title, r, 3))
    if problems:
        print(f"[result4-2覆盖率校验] 缺失 {len(problems)} 格, 样例: {problems[:10]}")
        return False
    print("[result4-2覆盖率校验] 全部应填单元格均已回填, 无缺失")
    return True


def fill_wide43(template, out, days, dates):
    """result4-3模板回填(复制模板按位置写值, 格式零改动), 沿用问题3已验证的自适应逻辑:
    每sheet自动判别 —— 第1行含>=100个时段标签=网格模式(按sheet关键字选计划/调整/紧急序列),
    否则竖排块模式: '充放'=6个4h块+0:00/24:00储电量(时刻封顶2), '紧急'=稀疏列表
    (非零按行写时段, 全0写'无/0')"""
    shutil.copyfile(template, out)
    wb = load_workbook(out)
    dset = set(dates)
    report = []
    for ws in wb.worksheets:
        title = ws.title
        nw = 0
        cols = {}
        for cidx in range(2, ws.max_column + 1):
            p = _parse_lab(ws.cell(row=1, column=cidx).value)
            if p:
                cols[p] = cidx
        grid = len(cols) >= 100
        if grid:
            key = "emer" if "紧急" in title else ("adj" if "调整" in title else "plan")
            for r in range(2, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d not in dset:
                    continue
                rec = days[dates.index(d)]
                seq = rec["emer"] if key == "emer" else (
                    rec["gadj"] if key == "adj" else rec["plan"])
                for t, cidx in cols.items():
                    ws.cell(row=r, column=cidx).value = round(seq[t - 1] * DT, 4)
                    nw += 1
        elif "充放" in title:
            cur, mom = None, 0
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur, mom = d, 0
                if cur is None:
                    continue
                tb = make_date_tables(days[dates.index(cur)], None, 0)
                lab = _norm(ws.cell(row=r, column=2).value)
                if lab in BLOCK_LABELS:
                    ch, dis = tb["blocks"][lab]
                    ws.cell(row=r, column=3).value = round(ch, 4)
                    ws.cell(row=r, column=4).value = round(dis, 4)
                    nw += 2
                ev = ws.cell(row=r, column=5).value
                en = _norm(ev)
                if ev is not None and mom < 2 and en not in ("时刻", "⋮", "︙", "…", "..."):
                    mom += 1
                    ws.cell(row=r, column=6).value = round(
                        tb["e0"] if mom == 1 else tb["e24"], 4)
                    nw += 1
        else:
            cur, drow, brows = None, None, []

            def _flush():
                nonlocal nw
                if cur is None:
                    return
                rec = days[dates.index(cur)]
                ws.cell(row=drow, column=2).value = "合计"
                ws.cell(row=drow, column=3).value = round(rec["emer_kwh"], 4)
                nw += 2
                nz = sorted(((rec["emer"][t] * DT, t) for t in range(144)
                             if rec["emer"][t] > TOL), reverse=True)
                if not nz and brows:
                    ws.cell(row=brows[0], column=2).value = "无"
                    ws.cell(row=brows[0], column=3).value = 0.0
                    nw += 2
                for (val, t), rr in zip(nz, brows):
                    ws.cell(row=rr, column=2).value = (
                        f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}")
                    ws.cell(row=rr, column=3).value = round(val, 4)
                    nw += 2
                if len(nz) > len(brows):
                    print(f"[result4-3/{title}] 注: {cur} 非零紧急时段{len(nz)}个, "
                          f"模板仅示Top{len(brows)}, 全量见result4_full.xlsx")

            for r in range(1, ws.max_row + 2):
                d = (_norm_date(ws.cell(row=r, column=1).value)
                     if r <= ws.max_row else "END")
                if d in dset or d == "END":
                    _flush()
                    cur, drow, brows = (d if d in dset else None), r, []
                elif cur is not None and r <= ws.max_row:
                    brows.append(r)
        report.append(f"{title}: 写入{nw}格({'网格' if grid else '竖排'})")
        print(f"[result4-3/{title}] {report[-1]}")
    wb.save(out)
    return report


def verify_fill43(out, dates):
    """覆盖率门禁: 读回输出, 网格sheet逐格非空检查; 竖排充放块/时刻与紧急日期行非空检查"""
    wb = load_workbook(out)
    dset = set(dates)
    problems = []
    for ws in wb.worksheets:
        title = ws.title
        cols = {}
        for cidx in range(2, ws.max_column + 1):
            p = _parse_lab(ws.cell(row=1, column=cidx).value)
            if p:
                cols[p] = cidx
        if len(cols) >= 100:
            for r in range(2, ws.max_row + 1):
                if _norm_date(ws.cell(row=r, column=1).value) not in dset:
                    continue
                for cidx in cols.values():
                    if ws.cell(row=r, column=cidx).value is None:
                        problems.append((title, r, cidx))
        else:
            cur = None
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur = d
                if cur is None:
                    continue
                if "充放" in title:
                    lab = _norm(ws.cell(row=r, column=2).value)
                    if lab in BLOCK_LABELS and (ws.cell(row=r, column=3).value is None
                                                or ws.cell(row=r, column=4).value is None):
                        problems.append((title, r, 3))
                else:
                    if d in dset and ws.cell(row=r, column=3).value is None:
                        problems.append((title, r, 3))
    if problems:
        print(f"[result4-3覆盖率校验] 缺失 {len(problems)} 格, 样例: {problems[:10]}")
        return False
    print("[result4-3覆盖率校验] 全部应填单元格均已回填, 无缺失")
    return True


def write_full42(path, daily_g, long_df, eco, chk, alpha_df, daily3=None):
    """Q4完整存档(result4_full.xlsx): 每日汇总 + 全年10min明细 + 验证对比
    + α_d逐日表 + (可选)Q4-3逐日成本分解"""
    wb = Workbook()
    ws = wb.active
    ws.title = "每日汇总"
    ws.append(list(daily_g.columns))
    for r in daily_g.itertuples(index=False):
        ws.append(list(r))
    ws2 = wb.create_sheet("全年10min明细")
    ws2.append(list(long_df.columns))
    for r in long_df.itertuples(index=False):
        ws2.append(list(r))
    ws3 = wb.create_sheet("验证对比")
    ws3.append(["指标", "数值"])
    for k, v in list(eco.items()) + list(chk.items()):
        if isinstance(v, (list, tuple, dict)):
            v = " / ".join(str(x) for x in v)
        ws3.append([k, v])
    ws4 = wb.create_sheet("alpha_d逐日")
    ws4.append(list(alpha_df.columns))
    for r in alpha_df.itertuples(index=False):
        ws4.append(list(r))
    if daily3 is not None:
        ws5 = wb.create_sheet("Q43成本分解")
        ws5.append(list(daily3.columns))
        for r in daily3.itertuples(index=False):
            ws5.append(list(r))
    wb.save(path)