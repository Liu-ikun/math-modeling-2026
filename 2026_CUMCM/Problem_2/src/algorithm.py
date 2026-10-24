# -*- coding: utf-8 -*-
"""问题2 求解编排: 三策略(无储能基线/全年全局优化/逐日滚动对照) + 侧面验证 + 模板回填
路径: 2026_CUMCM/Problem_2/src/algorithm.py"""
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from model import DT, E0, build_horizon, solve

TOL = 1e-6
SLOT_LABELS = ["10:00-10:10", "12:00-12:10", "14:00-14:10",
               "16:00-16:10", "18:00-18:10", "20:00-20:10"]
SLOT_HOURS = [10, 12, 14, 16, 18, 20]
BLOCK_LABELS = ["0:00-4:00", "4:00-8:00", "8:00-12:00",
                "12:00-16:00", "16:00-20:00", "20:00-24:00"]
BLOCK_RANGE = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]
DATE_KEYS = ["2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"]


def _pick(vs):
    """提取变量值, None 视为 0"""
    return [v.value() or 0.0 for v in vs]


def solve_strategy(c_day, L, PV, storage=True, cyclic=False):
    """在全年时间轴上一次建模求解 (策略A: storage=False; 策略B: storage=True)"""
    D = len(L)
    c_all = c_day * D
    L_all = [x for row in L for x in row]
    PV_all = [x for row in PV for x in row]
    prob, G, Eme, Pch, Pdis, E, Sp = build_horizon(
        c_all, L_all, PV_all, storage=storage, cyclic=cyclic)
    status, obj = solve(prob)
    g, eme, sp = _pick(G), _pick(Eme), _pick(Sp)
    if storage:
        pch, pdis, e = _pick(Pch), _pick(Pdis), _pick(E)
    else:
        pch = pdis = [0.0] * len(g)
        e = None
    simult = sum(1 for n in range(len(g)) if pch[n] > TOL and pdis[n] > TOL)
    return dict(status=status, obj=obj, D=D, G=g, Eme=eme, Pch=pch, Pdis=pdis,
                E=e, Sp=sp, n_simult=simult,
                spill=sum(sp) * DT, emer=sum(eme) * DT, buy=sum(g) * DT)


def solve_myopic(c_day, L, PV):
    """策略C(对照实验): 逐日独立LP, 无末态约束 -> 每天午夜放空储能的短视策略"""
    D = len(L)
    g, eme, pch, pdis, sp, e = [], [], [], [], [], [E0]
    e_cur, obj = E0, 0.0
    for d in range(D):
        prob, G, Eme, Pch, Pdis, E, Sp = build_horizon(
            c_day, L[d], PV[d], storage=True, e_init=e_cur)
        _, o = solve(prob)
        g += _pick(G); eme += _pick(Eme); pch += _pick(Pch)
        pdis += _pick(Pdis); sp += _pick(Sp)
        e_day = _pick(E); e += e_day[1:]; e_cur = e_day[-1]; obj += o
    simult = sum(1 for n in range(len(g)) if pch[n] > TOL and pdis[n] > TOL)
    return dict(status="OK", obj=obj, D=D, G=g, Eme=eme, Pch=pch, Pdis=pdis,
                E=e, Sp=sp, n_simult=simult,
                spill=sum(sp) * DT, emer=sum(eme) * DT, buy=sum(g) * DT)


def validate(res, c_day, L, PV):
    """侧面验证: 功率平衡双口径残差/SOC在界/同时充放/紧急购电/弃光"""
    T = 144
    g, pch, pdis, sp = res["G"], res["Pch"], res["Pdis"], res["Sp"]
    eme = res["Eme"]
    errs = []
    for n in range(res["D"] * T):
        d, t = divmod(n, T)
        errs.append(abs(g[n] + eme[n] + PV[d][t] + pdis[n] - L[d][t] - pch[n] - sp[n]))
    bal = max(errs)
    scale = max(1e-9, max(max(r) for r in L))
    in_bounds = res["E"] is None or all(1200 - 1e-6 <= x <= 10800 + 1e-6 for x in res["E"])
    return {
        "功率平衡最大绝对残差kW": bal,
        "功率平衡相对残差": bal / scale,
        "SOC全程在界": in_bounds,
        "同时充放时段数": res["n_simult"],
        "紧急购电总量kWh": round(res["emer"], 2),
        "弃光总量kWh": round(res["spill"], 2),
    }


def daily_agg(res, c_day, dates):
    """把全年最优解聚合为每日一行: 日期/购电费/购电量/弃光/吞吐/紧急购电"""
    T = 144
    rows = []
    for d in range(res["D"]):
        g = res["G"][d * T:(d + 1) * T]
        sp = res["Sp"][d * T:(d + 1) * T]
        pdis_d = res["Pdis"][d * T:(d + 1) * T]
        eme = res["Eme"][d * T:(d + 1) * T]
        rows.append([str(dates[d]),
                     round(sum(c_day[t] * g[t] * DT for t in range(T)), 2),
                     round(sum(g) * DT, 2), round(sum(sp) * DT, 2),
                     round(sum(pdis_d) * DT, 2), round(sum(eme) * DT, 4)])
    return pd.DataFrame(rows, columns=["日期", "购电费(元)", "购电量kWh",
                                       "弃光电量kWh", "吞吐电量kWh", "紧急购电kWh"])


def economy_check(res_b, res_a, res_c=None):
    """三策略经济对比: 全局优化 vs 无储能基线 vs (可选)逐日滚动"""
    out = {"有储能全年购电费(元)": round(res_b["obj"], 2),
           "无储能全年购电费(元)": round(res_a["obj"], 2),
           "节省购电费(元)": round(res_a["obj"] - res_b["obj"], 2)}
    out["节省比例"] = round(out["节省购电费(元)"] / res_a["obj"], 4) if res_a["obj"] else None
    if res_c is not None:
        out["逐日滚动全年购电费(元)"] = round(res_c["obj"], 2)
        out["全局优化优于逐日滚动(元)"] = round(res_c["obj"] - res_b["obj"], 2)
    return out


def fill_template(template: Path, out: Path, daily: pd.DataFrame):
    """按附件5模板原格式回填 result2.xlsx: 复制模板 -> 表头关键字定位列 -> 写每日行
    若模板无'日期'表头则跳过并返回提示, 由调用方决定是否人工核对模板结构。"""
    shutil.copyfile(template, out)
    wb = load_workbook(out)
    report = []
    for ws in wb.worksheets:
        header_row, colmap = None, {}
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30)):
            for c in row:
                if not isinstance(c.value, str):
                    continue
                t = "".join(str(c.value).split())
                if "日期" in t:
                    colmap["date"] = c.column
                elif "购电费" in t:
                    colmap["cost"] = c.column
                elif "购电量" in t:
                    colmap["buy"] = c.column
                elif "弃光" in t:
                    colmap["spill"] = c.column
                elif "放电" in t:
                    colmap["dis"] = c.column
                elif "紧急" in t:
                    colmap["eme"] = c.column
            if "date" in colmap:
                header_row = colmap["date"]
                break
        if header_row:
            for i, rec in enumerate(daily.to_dict("records")):
                rr = header_row + 1 + i
                if "date" in colmap:
                    ws.cell(row=rr, column=colmap["date"], value=rec["日期"])
                if "cost" in colmap:
                    ws.cell(row=rr, column=colmap["cost"], value=rec["购电费(元)"])
                if "buy" in colmap:
                    ws.cell(row=rr, column=colmap["buy"], value=rec["购电量kWh"])
                if "spill" in colmap:
                    ws.cell(row=rr, column=colmap["spill"], value=rec["弃光电量kWh"])
                if "dis" in colmap:
                    ws.cell(row=rr, column=colmap["dis"], value=rec["吞吐电量kWh"])
                if "eme" in colmap:
                    ws.cell(row=rr, column=colmap["eme"], value=rec["紧急购电kWh"])
            report.append(f"{ws.title}: 每日汇总 {len(daily)} 行已回填")
    wb.save(out)
    return report


def write_full(path, daily_b, long_df, eco, chk):
    """完整结果文件: 每日汇总 + 全年10min明细 + 验证对比 (提交模板之外的存档)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "每日汇总"
    ws.append(list(daily_b.columns))
    for r in daily_b.itertuples(index=False):
        ws.append(list(r))
    ws2 = wb.create_sheet("全年10min明细")
    ws2.append(list(long_df.columns))
    for r in long_df.itertuples(index=False):
        ws2.append(list(r))
    ws3 = wb.create_sheet("验证对比")
    ws3.append(["指标", "数值"])
    for k, v in list(eco.items()) + list(chk.items()):
        if isinstance(v, (list, tuple)):
            v = " / ".join(str(x) for x in v)
        ws3.append([k, v])
    wb.save(path)


def long_table(res, c_day, dates, start=None):
    """策略B的全年 10min 明细长表 (日期, 时刻, 电价, 购电, 充电, 放电, 弃光, 储电量)"""
    T = 144
    rows = []
    for d in range(res["D"]):
        if start and str(dates[d]) < start:
            continue
        for t in range(T):
            n = d * T + t
            rows.append([str(dates[d]),
                         f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
                         c_day[t], res["G"][n], res["Pch"][n], res["Pdis"][n],
                         res["Sp"][n], res["E"][n + 1] if res["E"] else None])
    return pd.DataFrame(rows, columns=["日期", "时刻", "电价(元/kWh)", "购电功率kW",
                                       "充电kW", "放电kW", "弃光kW", "时段末储电量kWh"])


def _day_tables(res, c_day, d):
    """第 d 天的表1/表2/表3数据(任意日期可算, 不局限于四个指定日期)"""
    T = 144
    n0 = d * T
    g = res["G"][n0:n0 + T]
    pch = res["Pch"][n0:n0 + T]
    pdis = res["Pdis"][n0:n0 + T]
    eme = res["Eme"][n0:n0 + T]
    e = res["E"][n0:n0 + T + 1]
    buy = [x * DT for x in g]
    cost = sum(c_day[t] * g[t] * DT for t in range(T))
    return dict(
        slots={lab: buy[h * 6] for lab, h in zip(SLOT_LABELS, SLOT_HOURS)},
        totals={"buy": sum(buy), "cost": cost}, buy=sum(buy), cost=cost,
        blocks={lab: (sum(pch[a:b]) * DT, sum(pdis[a:b]) * DT)
                for lab, (a, b) in zip(BLOCK_LABELS, BLOCK_RANGE)},
        e0=e[0], e24=e[-1],
        t3=[(f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}",
             round(eme[t] * DT, 4)) for t in range(T)])


def make_date_tables(res, c_day, dates):
    """题面表3指定四日期的表1/表2/表3数据(论文用)"""
    return {k: _day_tables(res, c_day, dates.index(k))
            for k in DATE_KEYS if k in dates}


def fill_wide(template, out, res, dates, c_day):
    """result2模板回填v3(复制模板按位置写值, 格式零改动):
    - 计划购电量sheet: 日期行×144时段列网格 + 右侧全天购电量/购电费汇总列(按行直算)
    - 充放电量sheet: 竖排表2块, 按模板现有日期块(示意压缩版)逐块回填
    - 紧急购电量sheet: 竖排表3块; 紧急购电恒0 -> 每日期块首行写显式零行"无/0",
      若模板带时段标签则按标签回填非零值"""
    shutil.copyfile(template, out)
    wb = load_workbook(out)
    T = 144
    dset = set(dates)
    report = []
    for ws in wb.worksheets:
        title = ws.title
        nw = 0
        if "紧急" in title:
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
                    ws.cell(row=r, column=3).value = \
                        round(res["Eme"][n0 + p[0] - 1] * DT, 4)
                    nw += 1
        elif "充放" in title:
            cur, mom = None, 0
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur = d
                    mom = 0
                if cur is None:
                    continue
                tb = _day_tables(res, c_day, dates.index(cur))
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
        else:
            cols, csum = {}, {}
            for c in range(2, ws.max_column + 1):
                hv = ws.cell(row=1, column=c).value
                p = _parse_lab(hv)
                if p:
                    cols[p[0]] = c
                elif hv is not None:
                    k = _norm(hv)
                    if "购电费" in k:
                        csum["cost"] = c
                    elif "购电量" in k:
                        csum["buy"] = c
            for r in range(2, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d not in dset:
                    continue
                tb = _day_tables(res, c_day, dates.index(d))
                g = res["G"][dates.index(d) * T: dates.index(d) * T + T]
                for t, c in cols.items():
                    ws.cell(row=r, column=c).value = round(g[t - 1] * DT, 4)
                    nw += 1
                if "buy" in csum:
                    ws.cell(row=r, column=csum["buy"]).value = round(tb["buy"], 4)
                    nw += 1
                if "cost" in csum:
                    ws.cell(row=r, column=csum["cost"]).value = round(tb["cost"], 4)
                    nw += 1
        miss = []
        if "充放" not in title and "紧急" not in title:
            miss = [t for t in range(1, 145) if t not in cols]
        report.append(f"{title}: 写入{nw}格, 缺失时段{miss[:5]}")
        print(f"[{title}] 写入单元格={nw} 缺失时段={miss[:5]}")
    wb.save(out)
    print(f"已按宽表网格回填: {out}")
    return report


def verify_fill(out, dates):
    """回填后覆盖率门禁: 重新读回输出文件, 逐格检查'应有值'的单元格非空,
    杜绝静默缺失; 有缺则打印(sheet,行,列)坐标并返回False"""
    wb = load_workbook(out)
    dset = set(dates)
    problems = []
    for ws in wb.worksheets:
        title = ws.title
        if "充放" in title or "紧急" in title:
            cur = None
            for r in range(1, ws.max_row + 1):
                d = _norm_date(ws.cell(row=r, column=1).value)
                if d in dset:
                    cur = d
                if cur is None:
                    continue
                if "紧急" in title:
                    is_date_row = d in dset
                    has_lab = _parse_lab(ws.cell(row=r, column=2).value) is not None
                    if (is_date_row or has_lab) \
                            and ws.cell(row=r, column=3).value is None:
                        problems.append((title, r, 3))
                else:
                    lab = _norm(ws.cell(row=r, column=2).value)
                    if lab in BLOCK_LABELS and (
                            ws.cell(row=r, column=3).value is None
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
            cols, csum = {}, []
            for c in range(2, ws.max_column + 1):
                hv = ws.cell(row=1, column=c).value
                p = _parse_lab(hv)
                if p:
                    cols[p[0]] = c
                elif hv is not None and ("购电费" in _norm(hv)
                                         or "购电量" in _norm(hv)):
                    csum.append(c)
            for r in range(2, ws.max_row + 1):
                if _norm_date(ws.cell(row=r, column=1).value) not in dset:
                    continue
                for c in list(cols.values()) + csum:
                    if ws.cell(row=r, column=c).value is None:
                        problems.append((title, r, c))
    if problems:
        print(f"[覆盖率校验] 缺失 {len(problems)} 格, 坐标样例: {problems[:10]}")
        return False
    print("[覆盖率校验] 全部应填单元格均已回填, 无缺失")
    return True


import re

_LAB_RE = re.compile(r"(\d{1,2}):(\d{1,2})\s*-\s*(\d{1,2}):(\d{1,2})")


def _norm(v):
    return re.sub(r"\s+", "", str(v)) if v is not None else ""


def _norm_date(v):
    """日期归一化: 兼容 datetime / 2025-03-20 / 2025.3.20 / 2025/3/20 等写法"""
    if v is None:
        return None
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.strftime("%Y-%m-%d")


def _parse_lab(s):
    """时段标签 -> (t, kind): 以区间右端点定 t(右端点约定, 与Q1一致);
    标签含'充电'/'放电'字样时区分充放列; '0:00'右端点视为24:00(t=144)"""
    m = _LAB_RE.search(str(s))
    if not m:
        return None
    h1, m1, h2, m2 = map(int, m.groups())
    end = h2 * 60 + m2
    if end == 0:
        end = 1440
    t = end // 10
    if not 1 <= t <= 144:
        return None
    kind = "ch" if "充电" in str(s) else ("dis" if "放电" in str(s) else "net")
    return t, kind
