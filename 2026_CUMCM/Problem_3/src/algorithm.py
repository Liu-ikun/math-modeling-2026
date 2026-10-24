# -*- coding: utf-8 -*-
"""问题3 编排: MPC滚动 vs 日前对照 vs 频率实验 + 实际执行结算 + 侧面验证 + 模板回填(含覆盖率门禁)"""
import re
import shutil

import pandas as pd
from openpyxl import load_workbook

from model import DT, E0, build_stage, solve, solve_stage

TOL = 1e-6
UPDATES = (0, 6, 12, 18)
DATE_KEYS = ["2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"]
Q2_PERFECT_COST = 13751266.38   # 问题2完美信息全局最优 = 问题3的完美信息下界
BLOCK_LABELS = ["0:00-4:00", "4:00-8:00", "8:00-12:00",
                "12:00-16:00", "16:00-20:00", "20:00-24:00"]
BLOCK_RANGE = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
_LAB_RE = re.compile(r"(\d{1,2}):(\d{1,2})\s*-\s*(\d{1,2}):(\d{1,2})")
SHIFT_BY_H = {0: 0, 6: 0, 12: 0, 18: 0}


def set_fc_shift(shift_map):
    SHIFT_BY_H.update(shift_map)


def _shift_for(H):
    h0 = max(h for h in SHIFT_BY_H if h <= H)
    return SHIFT_BY_H[h0]


def calibrate_shift(fcmap, PV, dates, lo=-8, hi=8):
    """宽域RMSE自动对齐: 逐发布时扫描整点位移, 取fc-act的RMSE最小者。
    RMSE对错位敏感(爬坡处数千kW)而对系统性偏差钝感, 判据强于|平均偏差|;
    自动吸收'lead语义滞后'与'附件2光伏数值旋转'两类口径差异"""
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
    """日出边沿门禁: 应用校准位移后, fc首个>50kW的lead映射小时与实际首个>50kW小时
    的中位差应≈0; 不为0则口径仍错, 拒绝进入求解"""
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
    m = _LAB_RE.search(str(s))
    if not m:
        return None
    h1, m1, h2, m2 = map(int, m.groups())
    end = h2 * 60 + m2 or 1440
    t = end // 10
    return t if 1 <= t <= 144 else None


def fc_series(fcmap, date, H, t0=0, vmap=None, fadj=None):
    """发布时H可用的10min光伏预报(ZOH): 虚拟发布(完全信息上界, 语义自定shift=0)优先,
    不覆盖时段按新->旧回退到真实发布, 杜绝'零预报真空'"""
    issues = fcmap[date]
    virt = vmap is not None and H in vmap.get(date, {})
    if virt:
        merged = dict(issues)
        merged.update(vmap.get(date, {}))
        src_map = merged
    else:
        src_map = issues
    order = [(h, SHIFT_BY_H[h] if h in SHIFT_BY_H else 0)
             for h in sorted(src_map) if h <= H][::-1]
    out = []
    for t in range(t0 + 1, 145):
        v, used = 0.0, None
        for h, s in order:
            k = t // 6 - h - s
            if 1 <= k <= 24 and k - 1 < len(src_map[h]):
                v, used = src_map[h][k - 1], (h, k)
                break
        if fadj and used:
            key = ("v", k) if virt else used
            v = max(0.0, v - fadj.get(key, 0.0))
        out.append(v)
    return out


def build_fc_stats(fcmap, PV, dates):
    """逐(发布时H,lead k)统计预报误差err=fc-act: 均值/标准差/全样本(供经验分位)"""
    import math
    s1, cnt, errs = {}, {}, {}
    for di, date in enumerate(dates):
        for H, leads in fcmap[date].items():
            for k in range(1, 25):
                hh = H + k - 1 + SHIFT_BY_H[H]
                if 0 <= hh <= 23:
                    act = sum(PV[di][hh * 6:hh * 6 + 6]) / 6.0
                    e = leads[k - 1] - act
                    key = (H, k)
                    s1[key] = s1.get(key, 0.0) + e
                    cnt[key] = cnt.get(key, 0) + 1
                    errs.setdefault(key, []).append(e)
    bias = {q: s1[q] / cnt[q] for q in cnt}
    std = {q: math.sqrt(max(0.0, sum(x * x for x in errs[q]) / cnt[q]
                            - bias[q] ** 2)) for q in cnt}
    return bias, std, errs


def robust_adj(errs, q_plan=0.8, q_adj=0.7):
    """分阶段报童经验分位校正: plan段理论F=0.8(过剩c/缺电增量4c),
    adj段理论F=0.7(过剩1.5c/缺电增量3.5c); 虚拟发布按lead k池化('v'键),
    与虚拟预报抽样同池 => 校正后belief无系统失真; q由主流程扫描经验确定"""
    import numpy as np
    adj = {}
    for (H, k), v in errs.items():
        adj[(H, k)] = float(np.quantile(v, q_plan if H == 0 else q_adj))
    byk = {}
    for (H, k), v in errs.items():
        byk.setdefault(k, []).extend(v)
    for k, v in byk.items():
        adj[("v", k)] = float(np.quantile(v, q_adj))
    return adj


def build_virtual_fc(PV, dates, errs, seed=2026):
    """虚拟发布预报(完全信息上界): 虚拟发布时H lead k -> 目标小时τ=H+k,
    与fc_series运行约定(k=t//6-h-s)严格一致——若用τ=H+k-1则虚拟预报比实际
    滞后一小时, 爬坡段引发5c紧急购电使'完美信息'成本反超真实预报(V6=64.2k>F4=53.9k);
    预报=act(τ)即ε≡0, 作为'新预报产品误差为零'的理想上界, 与F曲线(真实4发布·仅重优化)
    夹逼出加密发布的真实价值区间; 确定性无随机数。
    注: 蒙特卡洛ε抽样方案已否决——独立逐小时噪声在5c/1c不对称惩罚下使成本随更新次数
    单调暴涨, 噪声代价淹没信息价值, 非真实产品的特性"""
    vmap = {}
    for di, date in enumerate(dates):
        vh = {}
        for H in range(24):
            if H in (0, 6, 12, 18):
                continue
            row = []
            for k in range(1, 25):
                tau = H + k
                if tau > 23:
                    break
                row.append(sum(PV[di][tau * 6:tau * 6 + 6]) / 6.0)
            if row:
                vh[H] = row
        vmap[date] = vh
    return vmap


def run_day(c, L, PVact, fcmap, date, e_init, updates=UPDATES,
            solver="highs", vmap=None, fadj=None):
    """MPC滚动决策 + 实际执行结算; lp_obj=阶段LP目标之和(求解器等价性判据)"""
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
        pvfc = fc_series(fcmap, date, H, t0, vmap, fadj)
        plan_arg = None if mode == "plan" else plan[t0:]
        if solver == "highs":
            s = solve_stage(c[t0:], L[t0:], pvfc, plan=plan_arg,
                            e_init=E[t0], mode=mode)
            g, pc, pd_, e = s["G"], s["Pch"], s["Pdis"], s["E"]
            obj_sum += s["obj"]
        else:
            prob, v = build_stage(c[t0:], L[t0:], pvfc, plan=plan_arg,
                                  e_init=E[t0], mode=mode)
            st, ov = solve(prob)
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
    cost = 0.0
    for t in range(T):
        net = gadj[t] + PVact[t] + pdis[t] - L[t] - pch[t]
        emer[t] = max(-net, 0.0)
        sp[t] = max(net, 0.0)
        E[t + 1] = E[t] + 0.9 * pch[t] * DT - pdis[t] * DT / 0.9
        dp = max(gadj[t] - plan[t], 0.0)
        dm = max(plan[t] - gadj[t], 0.0)
        cost += (c[t] * plan[t] + 1.5 * c[t] * dp - 0.5 * c[t] * dm
                 + 5.0 * c[t] * emer[t]) * DT
    simult = sum(1 for t in range(T) if pch[t] > TOL and pdis[t] > TOL)
    return dict(plan=plan, gadj=gadj, pch=pch, pdis=pdis, E=E, emer=emer, sp=sp,
                cost=cost, simult=simult, buy=sum(gadj) * DT,
                emer_kwh=sum(emer) * DT, spill=sum(sp) * DT, lp_obj=obj_sum)


def validate(days, c, L, PVact, fcmap, dates):
    """侧面验证: 执行平衡双口径残差/SOC在界/同时充放/紧急购电与预报偏差的集中度"""
    errs, e_all, simult = [], [], 0
    for rec, date in zip(days, dates):
        T = 144
        for t in range(T):
            errs.append(abs(rec["gadj"][t] + rec["emer"][t] + PVact[dates.index(date)][t]
                            + rec["pdis"][t] - L[dates.index(date)][t]
                            - rec["pch"][t] - rec["sp"][t]))
        e_all += rec["E"]
        simult += rec["simult"]
    pairs = []
    for rec, date in zip(days, dates):
        di = dates.index(date)
        fc_used = []
        for i, H in enumerate(UPDATES):
            t0, t1 = H * 6, (UPDATES[i + 1] * 6 if i + 1 < len(UPDATES) else 144)
            fc_used += fc_series(fcmap, date, H, t0)[:t1 - t0]
        for t in range(144):
            pairs.append((abs(fc_used[t] - PVact[di][t]), rec["emer"][t] * DT))
    pairs.sort(key=lambda x: -x[0])
    tot_e = sum(p[1] for p in pairs) or 1.0
    err_share_top20 = sum(p[1] for p in pairs[:len(pairs) // 5]) / tot_e
    scale = max(1e-9, max(max(r) for r in L))
    return {"执行平衡最大绝对残差kW": max(errs),
            "执行平衡相对残差": max(errs) / scale,
            "SOC全程在界": all(1200 - 1e-6 <= x <= 10800 + 1e-6 for x in e_all),
            "同时充放时段数": simult,
            "紧急购电总量kWh": round(sum(r["emer_kwh"] for r in days), 2),
            "紧急购电时段数": sum(1 for r in days for x in r["emer"] if x > TOL),
            "Top20%偏差时段承载紧急电量占比": round(err_share_top20, 4)}


def make_date_tables(rec, c, d):
    T = 144
    buy_p = [x * DT for x in rec["plan"]]
    buy_a = [x * DT for x in rec["gadj"]]
    return dict(
        slots_p={lab: buy_p[h * 6] for lab, h in
                 zip(["10:00-10:10", "12:00-12:10", "14:00-14:10", "16:00-16:10",
                      "18:00-18:10", "20:00-20:10"], [10, 12, 14, 16, 18, 20])},
        slots_a={lab: buy_a[h * 6] for lab, h in
                 zip(["10:00-10:10", "12:00-12:10", "14:00-14:10", "16:00-16:10",
                      "18:00-18:10", "20:00-20:10"], [10, 12, 14, 16, 18, 20])},
        totals={"buy_p": sum(buy_p), "buy_a": sum(buy_a),
                "cost": rec["cost"]},
        blocks={lab: (sum(rec["pch"][a:b]) * DT, sum(rec["pdis"][a:b]) * DT)
                for lab, (a, b) in zip(BLOCK_LABELS, BLOCK_RANGE)},
        e0=rec["E"][0], e24=rec["E"][-1],
        t3=[(f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
             round(rec["emer"][t] * DT, 4)) for t in range(T)])


def fill_wide3(template, out, days, dates):
    """result3模板回填(复制模板按位置写值): 每sheet自适应——第1行含时段标签=网格模式,
    否则竖排块模式; 网格按sheet关键字选序列(计划/调整/紧急), 竖排充放=块+时刻封顶2,
    竖排紧急=稀疏列表(全0写'无/0', 非零按行写入)"""
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
                        f"{t // 6}:{t % 6 * 10:02d}-"
                        f"{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}")
                    ws.cell(row=rr, column=3).value = round(val, 4)
                    nw += 2
                if len(nz) > len(brows):
                    print(f"[{title}] 注: {cur} 非零紧急时段{len(nz)}个, 模板仅示Top"
                          f"{len(brows)}, 全量明细见result3_full")

            for r in range(1, ws.max_row + 2):
                d = (_norm_date(ws.cell(row=r, column=1).value)
                     if r <= ws.max_row else "END")
                if d in dset or d == "END":
                    _flush()
                    cur, drow, brows = (d if d in dset else None), r, []
                elif cur is not None and r <= ws.max_row:
                    brows.append(r)
        report.append(f"{title}: 写入{nw}格({'网格' if grid else '竖排'})")
        print(report[-1])
    wb.save(out)
    return report


def verify_fill3(out, dates):
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
        print(f"[覆盖率校验] 缺失 {len(problems)} 格, 样例: {problems[:10]}")
        return False
    print("[覆盖率校验] 全部应填单元格均已回填, 无缺失")
    return True


def run_chain(c, L, PV, fcmap, dates, idxs, updates, solver="highs",
              vmap=None, fadj=None):
    """一条独立SOC链: 按日序串行递推e_init(链内不可并行), 链间可多进程"""
    days, e_cur = [], E0
    for d in idxs:
        r = run_day(c, L[d], PV[d], fcmap, dates[d], e_cur, updates, solver,
                    vmap, fadj)
        e_cur = r["E"][-1]
        days.append(r)
    return days