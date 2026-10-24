# -*- coding: utf-8 -*-
"""问题4 主流程（改进版·核心闭环）
Q4-2: Q2框架(完美信息全年一次性全局LP) + 附件4波动电价, 含无储能基线/逐日短视对照/跨日套利价值;
Q4-3: Q3框架(MPC滚动 0/6/12/18) + 附件4波动电价, 含时移诊断/日前对照/完美预报自洽/成本分解。
默认即核心闭环(无需 --fast), 预计 ≈2-3 分钟; --ipm/--check 为可选加速与验证开关。
运行: python main.py [--ipm] [--check]
--ipm   : 全局LP改用 highs-ipm(更快, 边界值可能非精确顶点; 默认 highs-ds)
--check : 双求解器(HiGHS vs CBC) 7天窗口全局LP + 阶段LP 隔离对照(验证体系, 平时可不跑)
路径: 2026_CUMCM/Problem_4/src/main.py"""
import sys
import time
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import concurrent.futures as cf
from algorithm import (DATE_KEYS, Q2_PERFECT_COST, Q3_A_COST, Q3_B_COST,
                       UPDATES, WIN_START, alpha_analysis, build_fc_stats,
                       build_perfect_fc, calibrate_shift, chain_cost,
                       daily_agg_global, diagnose_timing, economy_check_global,
                       fc_series, fill_wide42, fill_wide43, flatten,
                       long_table_g, make_date_tables, make_date_tables_g,
                       run_chain, set_fc_shift, solve_myopic, validate3,
                       validate_global, verify_alignment, verify_fill42,
                       verify_fill43, win_sum, write_full42)
from model import DT, E0, solve_global

ROOT = Path(__file__).resolve().parents[2]
OUT_T = ROOT / "Problem_4" / "outputs" / "tables"
OUT_F = ROOT / "Problem_4" / "outputs" / "figures"
OUT_R = ROOT / "Problem_4" / "outputs" / "report"
for d in (OUT_T, OUT_F, OUT_R):
    d.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({"font.sans-serif": ["SimHei", "Microsoft YaHei"],
                     "axes.unicode_minus": False, "axes.grid": True, "grid.alpha": 0.3,
                     "savefig.dpi": 200, "axes.spines.top": False,
                     "axes.spines.right": False})

# 问题2(附件1, 电价逐日重复)逐日短视 - 全局 的成本差, 引自问题2输出;
# 用于跨日套利2×2对照(附件1该值≈0, 附件4应显著>0), 避免重复跑附件1全局LP。
Q2_XDAY_VALUE = 26894.74


def _tpl(name):
    """附件5模板定位: 三处候选 + 排除输出文件自身(防 SameFileError)"""
    out = OUT_T / name
    cands = [ROOT / "Shared_Utils" / "data" / "raw" / "附件5" / name,
             ROOT / "Material" / "附件" / "附件5" / name,
             ROOT / "Problem_4" / "data" / "raw" / "附件5" / name]
    for p in cands:
        if p.exists() and p.resolve() != out.resolve():
            return p
    return None


def pick_col(df, keys):
    for k in keys:
        for c in df.columns:
            if k in str(c):
                return c
    return None


def _raw_path(name):
    for p in (ROOT / "Problem_4" / "data" / "raw" / name,
              ROOT / "Material" / "附件" / name,
              ROOT / "Problem_2" / "data" / "raw" / name,
              ROOT / "Problem_3" / "data" / "raw" / name):
        if p.exists():
            return p
    raise SystemExit(f"未找到 {name}")


def load_price4():
    """附件4原始直读: 宽表 (365,145) = [日期(合并单元格需ffill), 144个10min电价]。
    已由数据分析确认10分钟粒度, 无需任何插值; 探针打印表头防格式误判"""
    df = pd.read_excel(_raw_path("附件4.xlsx"), header=0)
    cols = [str(c) for c in df.columns]
    print(f"[探针·附件4] shape={df.shape} 列数={len(cols)} "
          f"首列='{cols[0]}' 次列='{cols[1]}' 末列='{cols[-1]}'")
    dcol = df.columns[0]
    dates = pd.to_datetime(df[dcol], errors="coerce").ffill() \
        .dt.strftime("%Y-%m-%d").tolist()
    vals = df.drop(columns=[dcol]).astype(float).to_numpy()
    if vals.shape[1] != 144:
        raise SystemExit(f"[附件4] 数值列数={vals.shape[1]}, 预期144; 请核对表头后调整解析")
    print(f"[附件4] {vals.shape[0]}天×{vals.shape[1]}时段 | "
          f"min={vals.min():.4f} max={vals.max():.4f} 年均={vals.mean():.6f} 元/kWh")
    return dates, vals.tolist()


def load_price1():
    """附件1分时电价(144,): 首选processed对齐表, 回退原始附件1首sheet按关键字挑列。
    用途: α_d分解的形状基准(shape) 与 跨日套利2×2的 k=0 参照"""
    proc = ROOT / "Problem_2" / "data" / "processed" / "附件1_10min对齐.xlsx"
    if proc.exists():
        df = pd.read_excel(proc)
        col = pick_col(df, ["电价"])
        v = df[col].astype(float).tolist()
        if len(v) == 144:
            print(f"[附件1] 来源: {proc.name} ({len(v)}时段)")
            return v
        print("[附件1] processed列长异常, 回退原始直读")
    p = _raw_path("附件1.xlsx")
    sheets = pd.ExcelFile(p).sheet_names
    df = pd.read_excel(p, header=0)
    print(f"[探针·附件1] sheets={sheets} 首sheet列名={[str(c) for c in df.columns][:6]}")
    col = pick_col(df, ["电价"])
    if col is None:
        raise SystemExit("[附件1] 未找到'电价'列")
    v = df[col].astype(float).tolist()
    if len(v) != 144:
        raise SystemExit(f"[附件1] 电价列长={len(v)}, 预期144")
    return v


def load_actual():
    """附件2双sheet直读: 实际负载/光伏 (D,144), 列序天然 00:10..23:50,0:00+1"""
    raw = _raw_path("附件2.xlsx")

    def parse(sn):
        df = pd.read_excel(raw, sheet_name=sn, header=0)
        df = df.rename(columns={df.columns[0]: "日期"})
        df["日期"] = pd.to_datetime(df["日期"])
        df = df.sort_values("日期")
        return (df["日期"].dt.strftime("%Y-%m-%d").tolist(),
                df.drop(columns=["日期"]).astype(float).to_numpy().tolist())

    sheets = pd.ExcelFile(raw).sheet_names
    dates, L = parse(next(s for s in sheets if any(k in s for k in ("负荷", "负载"))))
    PV = parse(next(s for s in sheets if any(k in s for k in ("光伏", "发电"))))[1]
    print(f"[附件2] {len(dates)}天×144时段 | sheets={sheets}")
    return L, PV, dates


def load_fc():
    """附件3原始直读: [日期(ffill), 预报时刻, 预报1小时..预报24小时]。
    严禁使用processed'附件3_10min宽表'——跨发布时点线性插值污染(问题3已踩坑)"""
    df = pd.read_excel(_raw_path("附件3.xlsx"), header=0)
    cols = [str(c) for c in df.columns]
    print(f"[探针·附件3] 列名({len(cols)}列): {cols[:4]} ... {cols[-2:]}")

    def lead_no(c):
        m = re.search(r"预报\s*(\d+)\s*小时", c)
        return int(m.group(1)) if m else None

    cand = [(lead_no(c), c) for c in cols]
    lead_cols = [c for n, c in sorted((x for x in cand if x[0] is not None),
                                      key=lambda x: x[0])]
    if len(lead_cols) != 24:
        lead_cols = cols[2:2 + 24]
        print(f"[load_fc] 正则识别到{len(lead_cols)}列, 退回位置法取第3~26列")
    date = pd.to_datetime(df.iloc[:, 0].ffill())
    hh = pd.to_numeric(df.iloc[:, 1].astype(str).str.extract(r"(\d{1,2})")[0],
                       errors="coerce")
    fcmap = {}
    for i in range(len(df)):
        H = hh.iloc[i]
        if pd.isna(H) or int(H) not in (0, 6, 12, 18):
            continue
        leads = [float(df[c].iloc[i]) for c in lead_cols]
        fcmap.setdefault(date.iloc[i].strftime("%Y-%m-%d"), {})[int(H)] = leads
    print(f"[附件3] {len(fcmap)} 天 × "
          f"{sorted({H for v in fcmap.values() for H in v})} 时点发布 | lead {len(lead_cols)}列")
    return fcmap


_G = {}


def _init(CC, LL, PP, FF, DD):
    _G.update(c4=CC, L=LL, PV=PP, fc=FF, dates=DD)


def _work(item):
    name, sch, idxs, vmap, fadj, timing, interp = item
    return name, run_chain(_G["c4"], _G["L"], _G["PV"], _G["fc"], _G["dates"],
                           idxs, sch, vmap=vmap, fadj=fadj, timing=timing,
                           interp=interp)


def main():
    t0 = time.time()
    method = "highs-ipm" if "--ipm" in sys.argv else "highs-ds"

    # ================= 数据装载 =================
    c1 = load_price1()
    d4, c4 = load_price4()
    L, PV, dates = load_actual()
    fcmap = load_fc()
    D = len(dates)
    if len(c4) != D:
        print(f"[附件4] 行数{len(c4)} != 附件2天数{D}, 按索引截断")
        c4 = c4[:D]
        d4 = d4[:D]
    if [str(x) for x in d4] != [str(x) for x in dates]:
        print("[附件4] 日期序列与附件2不一致, 按行序对齐(附件4行序视为与附件2相同)")
    print(f"数据: {D} 天 ({dates[0]} ~ {dates[-1]}) | 全局LP方法: {method}")

    # ================= 附件4电价结构(α_d) =================
    alpha_df, ainfo, alpha = alpha_analysis(c4, c1, L, PV, dates)
    print("[电价结构α_d] " + " | ".join(
        f"{k}={v}" for k, v in ainfo.items() if k != "月度α_d"))
    print("[月度α_d] " + " ".join(f"{m}月:{ainfo['月度α_d'][m]:.3f}"
                                for m in range(1, 13)))

    # ================= Q4-2: 全年一次性全局LP 三策略 =================
    c4_all, L_all, PV_all = flatten(c4), flatten(L), flatten(PV)
    res_g4 = solve_global(c4_all, L_all, PV_all, storage=True, method=method)
    res_n4 = solve_global(c4_all, L_all, PV_all, storage=False, method=method)
    res_m4 = solve_myopic(c4, L, PV, method=method)
    print(f"[Q4-2] 附件4全局LP obj={res_g4['obj']:.2f} | 无储能={res_n4['obj']:.2f} "
          f"| 逐日短视={res_m4['obj']:.2f} | 耗时 {time.time() - t0:.1f}s")

    # ================= 双求解器隔离对照(--check) =================
    if "--check" in sys.argv:
        from model import build_horizon, solve as psolve
        from model import build_stage as bstage, solve_stage
        w = 7
        cw, Lw, Pw = flatten(c4[:w]), flatten(L[:w]), flatten(PV[:w])
        r_h = solve_global(cw, Lw, Pw, storage=True)
        prob, *_ = build_horizon(cw, Lw, Pw, storage=True)
        st, o_p = psolve(prob, time_limit=600)
        rel1 = abs(r_h["obj"] - o_p) / max(1.0, abs(o_p))
        print(f"[双求解器·Q4-2全局LP {w}天] HiGHS={r_h['obj']:.4f} CBC={o_p:.4f} "
              f"相对差={rel1:.2e} (应<1e-6)")
        r_h2 = solve_global(cw, Lw, Pw, storage=False)
        prob2, *_ = build_horizon(cw, Lw, Pw, storage=False)
        st2, o_p2 = psolve(prob2, time_limit=600)
        rel2 = abs(r_h2["obj"] - o_p2) / max(1.0, abs(o_p2))
        print(f"[双求解器·Q4-2无储能 {w}天] HiGHS={r_h2['obj']:.4f} CBC={o_p2:.4f} "
              f"相对差={rel2:.2e} (应<1e-6)")
        pv0 = fc_series(fcmap, dates[0], 0)
        s1 = solve_stage(c4[0], L[0], pv0, e_init=E0, mode="plan")
        st1, ov1 = psolve(bstage(c4[0], L[0], pv0, e_init=E0, mode="plan")[0])
        r1 = abs(s1["obj"] - ov1) / max(1.0, abs(ov1))
        pv6 = fc_series(fcmap, dates[0], 6, 36)
        s2 = solve_stage(c4[0][36:], L[0][36:], pv6, plan=s1["G"][36:],
                         e_init=s1["E"][37], mode="adj")
        st2, ov2 = psolve(bstage(c4[0][36:], L[0][36:], pv6, plan=s1["G"][36:],
                                 e_init=s1["E"][37], mode="adj")[0])
        r2 = abs(s2["obj"] - ov2) / max(1.0, abs(ov2))
        print(f"[双求解器·Q4-3阶段LP] plan段 {r1:.2e} | adj段 {r2:.2e} (应<1e-6)")

    # ================= Q4-2 验证/经济性/日汇总 =================
    chk_g = validate_global(res_g4, L, PV)
    eco = economy_check_global(res_g4, res_n4, res_m4)
    daily_g = daily_agg_global(res_g4, c4, dates, alpha)
    daily_n = daily_agg_global(res_n4, c4, dates)
    daily_m = daily_agg_global(res_m4, c4, dates)
    daily_g.to_csv(OUT_T / "daily_summary_q42.csv", index=False, encoding="utf-8-sig")
    alpha_df.to_csv(OUT_T / "alpha_daily_q4.csv", index=False, encoding="utf-8-sig")
    win_g = win_sum(daily_g, "购电费(元)")
    win_m = win_sum(daily_m, "购电费(元)")
    win_n = win_sum(daily_n, "购电费(元)")
    print(f"[Q4-2窗口] 334天报告窗: 全局{win_g:.2f} | 短视{win_m:.2f} | 无储能{win_n:.2f}")
    print(f"[跨日套利价值] 附件4: 短视{res_m4['obj']:.0f} - 全局{res_g4['obj']:.0f} "
          f"= {eco['跨日套利价值(元)']:.2f} 元 (占全局 {eco['跨日套利价值占全局成本比']}) "
          f"| 附件1参照≈{Q2_XDAY_VALUE:.0f} 元(≈0)")

    # ================= Q4-2 模板回填 =================
    tpl2 = _tpl("result4-2.xlsx")
    if tpl2:
        report2 = fill_wide42(tpl2, OUT_T / "result4-2.xlsx", res_g4, dates, c4)
        ok2 = verify_fill42(OUT_T / "result4-2.xlsx", dates)
    else:
        report2, ok2 = [], None
        print("未找到模板 result4-2.xlsx, 跳过回填")

    # ================= Q4-2 四日期表 =================
    r1, r2, r3 = [], [], []
    dt4 = make_date_tables_g(res_g4, c4, dates)
    for key in DATE_KEYS:
        tb = dt4[key]
        for lab, v in tb["slots"].items():
            r1.append([key, lab, round(v, 4)])
        r1.append([key, "全天购电量", round(tb["buy"], 4)])
        r1.append([key, "全天购电费", round(tb["cost"], 4)])
        for lab, (ch, dis) in tb["blocks"].items():
            r2.append([key, lab, round(ch, 4), round(dis, 4)])
        r2.append([key, "0:00储电量", round(tb["e0"], 4), ""])
        r2.append([key, "24:00储电量", round(tb["e24"], 4), ""])
        for lab, v in tb["t3"]:
            r3.append([key, lab, v])
    pd.DataFrame(r1, columns=["日期", "时间段/项目", "计划购电量kWh"]).to_csv(
        OUT_T / "q42_table1_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r2, columns=["日期", "时间段/项目", "充电量kWh", "放电量kWh"]).to_csv(
        OUT_T / "q42_table2_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r3, columns=["日期", "时间段", "紧急购电量kWh"]).to_csv(
        OUT_T / "q42_table3_4dates.csv", index=False, encoding="utf-8-sig")

    # ================= Q4-2 图 =================
    xd = pd.to_datetime(dates)
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.plot(xd, alpha_df["α_d"], color="#4c72b0", lw=1.0, label="日水平因子 α_d")
    ax.axhline(1.0, color="k", lw=0.8, ls="--")
    ax.set_xlabel("日期"); ax.set_ylabel("α_d = 日均电价 / 附件1均价")
    ax.set_title("图1 附件4日水平因子α_d时序 (冬高春低, 与光伏季节反相)")
    ax.legend(frameon=False, fontsize=8)
    fig.autofmt_xdate()
    fig.tight_layout(); fig.savefig(OUT_F / "fig1_alpha_series.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 4.2))
    bars = ax.bar(["附件1电价\n(逐日重复·引自问题2)", "附件4电价\n(α_d波动·本问)"],
                  [Q2_XDAY_VALUE, eco["跨日套利价值(元)"]],
                  color=["#9aa0a6", "#dd8452"], width=0.5)
    ax.bar_label(bars, fmt="%.0f 元", fontsize=9)
    ax.set_ylabel("逐日短视 vs 全年全局 成本差 (元)")
    ax.set_title("图2 跨日套利价值: 波动电价激活跨日电量搬移")
    fig.tight_layout(); fig.savefig(OUT_F / "fig2_xday_value.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.plot(range(D), daily_n["购电费(元)"], color="#9aa0a6", lw=1.0, label="无储能基线")
    ax.plot(range(D), daily_m["购电费(元)"], color="#d62728", lw=1.0, label="逐日短视")
    ax.plot(range(D), daily_g["购电费(元)"], color="#4c72b0", lw=1.0, label="全年全局")
    ax.set_xlabel("日序号"); ax.set_ylabel("日购电费 (元)")
    ax.set_title("图3 逐日购电费对比 (全局 vs 无储能 vs 短视, 附件4)")
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(OUT_F / "fig3_daily_cost.png"); plt.close(fig)

    for key in DATE_KEYS:
        di = dates.index(key)
        fig, ax = plt.subplots(figsize=(11, 4.0))
        ax.plot(res_g4["E"][di * 144:(di + 1) * 144 + 1], color="#2ca02c", lw=1.6)
        ax.set_xlabel("10min 时段"); ax.set_ylabel("储电量 (kWh)", color="#2ca02c")
        ax.set_ylim(0, 11500)
        ax2 = ax.twinx()
        ax2.step(range(144), c4[di], where="post", color="#c44e52", lw=1.2)
        ax2.set_ylabel("附件4电价 (元/kWh)", color="#c44e52")
        ax.set_title(f"图4 {key} SOC与附件4电价 (低充高放)")
        fig.tight_layout(); fig.savefig(OUT_F / f"fig4_soc_price_{key}.png"); plt.close(fig)

    # ================= Q4-3: 对齐与门禁 =================
    smap = calibrate_shift(fcmap, PV, dates)
    set_fc_shift(smap)
    med = verify_alignment(fcmap, PV, dates)
    if abs(med) > 1:
        raise SystemExit("[对齐门禁] 校准后日出边沿仍不齐, 口径未定, 拒绝求解")
    bias, std, errs = build_fc_stats(fcmap, PV, dates)
    for H in (0, 6, 12, 18):
        ks = [k for (h, k) in bias if h == H]
        if ks:
            print(f"[预报诊断] 发布时{H}(shift{smap[H]:+d}): 平均偏差 "
                  f"{sum(bias[H, k] for k in ks) / len(ks):+.1f} kW | 平均标准差 "
                  f"{sum(std[H, k] for k in ks) / len(ks):.1f} kW | lead数 {len(ks)}")
    _ = diagnose_timing(fcmap, PV, dates)      # D1时移诊断: ceil应全0, floor应部分+1

    # ================= Q4-3: 核心5链并发 =================
    pmap = build_perfect_fc(PV, dates, timing="ceil")
    full = list(range(D))
    chains = [("B", UPDATES, full, None, None, "ceil", False),      # 提交策略
              ("Bleg", UPDATES, full, None, None, "floor", False),  # 时移对照
              ("A", (0,), full, None, None, "ceil", False),         # 朴素日前
              ("A2", (0,), full, None, None, "ceil", False),        # 确定性重复链
              ("BP", UPDATES, full, pmap, None, "ceil", False)]     # 完美预报自洽
    with cf.ProcessPoolExecutor(initializer=_init,
                                initargs=(c4, L, PV, fcmap, dates)) as ex:
        out = dict(ex.map(_work, chains))
    print(f"Q4-3 求解完成: {time.time() - t0:.1f} s (含{len(chains)}条并发链)")

    cost = {k: chain_cost(v) for k, v in out.items()}
    cost_b, cost_a = cost["B"], cost["A"]
    best_key = "B"
    days_b = out[best_key]
    print(f"[Q4-3提交策略] {best_key} = {cost[best_key]:.2f} 元")
    print(f"[Q4-3成本分解(365天)] B: 计划 {sum(r['cost_plan'] for r in out['B']):.2f} | "
          f"调整加项 {sum(r['cost_up'] for r in out['B']):.2f} | "
          f"调整减项 {sum(r['cost_dn'] for r in out['B']):.2f} | "
          f"紧急 {sum(r['cost_eme'] for r in out['B']):.2f}")
    print(f"[D1时移影响] B(ceil)={cost_b:.2f} vs Bleg(floor)={cost['Bleg']:.2f} "
          f"=> 时移修复价值 {cost['Bleg'] - cost_b:.2f} 元")
    # BP链=完美"小时级"预报MPC。附件3预报本质为小时分辨率(24 lead), 经ZOH降尺度到10min后
    # 与结算用的实际10min PV存在"小时内偏差", 故即便完美预报也会触发紧急购电 => BP > 短视LP。
    # 二者之差非代码bug, 而是"小时级预报分辨率的固有代价"; BP与B之差才是"预报误差的额外代价"。
    res_penalty = cost["BP"] - res_m4["obj"]
    fc_penalty = cost_b - cost["BP"]
    print(f"[信息代价三段分解] 完美10min短视LP={res_m4['obj']:.2f} "
          f"→ 完美小时预报MPC(BP)={cost['BP']:.2f} (分辨率代价 {res_penalty:.2f}) "
          f"→ 真实预报MPC(B)={cost_b:.2f} (误差代价 {fc_penalty:.2f})")
    ok_order = res_m4["obj"] <= cost["BP"] <= cost_b
    print(f"[BP排序门禁] 短视LP ≤ BP ≤ B: {res_m4['obj']:.0f} ≤ {cost['BP']:.0f} ≤ "
          f"{cost_b:.0f} {'✓' if ok_order else '✗ 异常'}")
    print(f"[自洽检查] A vs A2 重复链相对差 "
          f"{abs(cost['A'] - cost['A2']) / max(1.0, cost['A']):.2e} (应=0)")

    chk3 = validate3(days_b, L, PV, fcmap, dates)
    win_b = chain_cost(out["B"], dates, start=WIN_START)
    win_a = chain_cost(out["A"], dates, start=WIN_START)
    print(f"[Q4-3窗口] 334天报告窗: B={win_b:.2f} | A={win_a:.2f} | "
          f"阶梯(365天): 完美{res_g4['obj']:.0f} ≤ B{cost_b:.0f} ≤ A{cost_a:.0f}")

    # ================= Q4-3 模板回填 =================
    tpl3 = _tpl("result4-3.xlsx")
    if tpl3:
        report3 = fill_wide43(tpl3, OUT_T / "result4-3.xlsx", days_b, dates)
        ok3 = verify_fill43(OUT_T / "result4-3.xlsx", dates)
    else:
        report3, ok3 = [], None
        print("未找到模板 result4-3.xlsx, 跳过回填")

    rows = []
    for rec, dstr in zip(days_b, dates):
        rows.append([dstr, round(rec["cost"], 2), round(rec["cost_plan"], 2),
                     round(rec["cost_up"], 2), round(rec["cost_dn"], 2),
                     round(rec["cost_eme"], 2), round(rec["buy"], 2),
                     round(rec["emer_kwh"], 4), round(rec["spill"], 2)])
    daily3 = pd.DataFrame(rows, columns=["日期", "结算费用(元)", "计划购电费(元)",
                                         "调整加项(元)", "调整减项(元)", "紧急购电费(元)",
                                         "调整购电量kWh", "紧急购电量kWh", "弃光电量kWh"])
    daily3.to_csv(OUT_T / "daily_summary_q43.csv", index=False, encoding="utf-8-sig")

    # ================= Q4-3 四日期表 =================
    r1, r2, r3 = [], [], []
    for key in DATE_KEYS:
        di = dates.index(key)
        tb = make_date_tables(days_b[di], c4[di], di)
        for lab, v in tb["slots_p"].items():
            r1.append([key, lab, round(v, 4), round(tb["slots_a"][lab], 4)])
        for lab, (ch, dis) in tb["blocks"].items():
            r2.append([key, lab, round(ch, 4), round(dis, 4)])
        r2.append([key, "0:00储电量", round(tb["e0"], 4), ""])
        r2.append([key, "24:00储电量", round(tb["e24"], 4), ""])
        for lab, v in tb["t3"]:
            if v > 0:
                r3.append([key, lab, v])
        if not any(x[0] == key for x in r3):
            r3.append([key, "无", 0.0])
    pd.DataFrame(r1, columns=["日期", "时间段", "计划购电量kWh", "调整购电量kWh"]).to_csv(
        OUT_T / "q43_table1_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r2, columns=["日期", "时间段/项目", "充电量kWh", "放电量kWh"]).to_csv(
        OUT_T / "q43_table2_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r3, columns=["日期", "时间段", "紧急购电量kWh"]).to_csv(
        OUT_T / "q43_table3_4dates.csv", index=False, encoding="utf-8-sig")

    # ================= Q4-3 图 =================
    fig, ax = plt.subplots(figsize=(8, 4.4))
    names = ["Q4-2完美信息下界", "Q4-3朴素MPC(B)", "Q4-3朴素日前(A)"]
    vals = [res_g4["obj"], cost_b, cost_a]
    bars = ax.bar(names, vals, color=["#9aa0a6", "#4c72b0", "#d62728"], width=0.55)
    ax.bar_label(bars, fmt="%.0f 元", fontsize=9)
    ax.set_ylabel("全年结算费用 (元)")
    ax.set_title("图5 成本阶梯: 完美信息 ≤ MPC ≤ 日前 (附件4)")
    fig.tight_layout(); fig.savefig(OUT_F / "fig5_cost_ladder.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4.2))
    for key in DATE_KEYS:
        ax.plot(days_b[dates.index(key)]["E"], label=key)
    ax.axhline(1200, ls="--", lw=0.8, color="k")
    ax.axhline(10800, ls="--", lw=0.8, color="k")
    ax.set_xlabel("10min 时段"); ax.set_ylabel("储电量 (kWh)")
    ax.set_title("图6 Q4-3 指定日期 SOC 轨迹"); ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(OUT_F / "fig6_soc_dates.png"); plt.close(fig)

    # ================= 存档与摘要 =================
    long_g = long_table_g(res_g4, c4, dates, start=WIN_START)
    write_full42(OUT_T / "result4_full.xlsx", daily_g, long_g, eco, chk_g,
                 alpha_df, daily3)

    lines = [
        "# 问题4 求解摘要（改进版核心闭环：建模 → 求解 → 结果分析含侧面验证）",
        "",
        f"- 数据: {D} 天 ({dates[0]} ~ {dates[-1]}) | 报告窗口: {WIN_START}起(334天) | "
        f"总耗时 {time.time() - t0:.1f} s | 全局LP方法: {method}",
        f"- result4-2.xlsx 回填: {len(report2)}个sheet 覆盖率{'通过' if ok2 else '未通过/未做'} | "
        f"result4-3.xlsx 回填: {len(report3)}个sheet 覆盖率{'通过' if ok3 else '未通过/未做'}",
        "",
        "## 一、附件4电价结构(α_d)",
        f"- α_d∈{ainfo['α_d范围']} 极差比 {ainfo['α_d极差比']} | 保本峰谷比 {ainfo['保本峰谷价比']}",
        f"- 逐日峰谷比 {ainfo['逐日峰谷比min/中位/max']} | 超保本线 {ainfo['峰谷比超保本线天数']}",
        f"- corr(α_d,日均净负荷)={ainfo['corr(α_d,日均净负荷)']} (净需求定价结构)",
        "",
        "## 二、Q4-2 全年全局LP(完美信息, 附件4)",
        f"- 全年全局最优: {res_g4['obj']:.2f} 元 (334天窗 {win_g:.2f}) | "
        f"无储能基线: {res_n4['obj']:.2f}",
        f"- 逐日短视对照: {res_m4['obj']:.2f} (334天窗 {win_m:.2f}) | "
        f"储能节省: {eco['储能节省(元)']:.2f} 元 ({eco['储能节省比例']})",
        f"- 跨日套利价值(附件4): {eco['跨日套利价值(元)']:.2f} 元 "
        f"(占全局 {eco['跨日套利价值占全局成本比']}) | 附件1参照≈{Q2_XDAY_VALUE:.0f}元(≈0)",
        f"- 紧急购电(完美信息应=0): {chk_g['紧急购电总量kWh']} kWh | "
        f"弃光: {chk_g['弃光总量kWh']} kWh",
        "",
        "## 三、Q4-3 MPC(附件4)",
        f"- 提交策略: {best_key} = {cost[best_key]:.2f} 元 (365天) | "
        f"334天窗 B={win_b:.2f}",
        f"- 朴素MPC(B): {cost_b:.2f} | 朴素日前(A): {cost_a:.2f} "
        f"(滚动更新信息价值 {cost_a - cost_b:.2f} 元)",
        f"- 成本分解(B): 计划 {sum(r['cost_plan'] for r in out['B']):.2f} | "
        f"调整加项 {sum(r['cost_up'] for r in out['B']):.2f} | "
        f"调整减项 {sum(r['cost_dn'] for r in out['B']):.2f} | "
        f"紧急 {sum(r['cost_eme'] for r in out['B']):.2f}",
        f"- D1时移: B(ceil)={cost_b:.2f} vs Bleg(floor)={cost['Bleg']:.2f} "
        f"=> 时移修复价值 {cost['Bleg'] - cost_b:.2f} 元",
        f"- 信息代价分解: 完美10min {res_m4['obj']:.0f} → 完美小时BP {cost['BP']:.0f} "
        f"(分辨率代价 {res_penalty:.0f}) → 真实B {cost_b:.0f} (预报误差代价 {fc_penalty:.0f})",
        f"- 确定性: A≡A2 相对差 {abs(cost['A'] - cost['A2']) / max(1.0, cost['A']):.2e} (应=0)",
        "",
        "## 四、侧面验证(Q4-3提交策略)",
        f"- 执行平衡最大绝对残差 {chk3['执行平衡最大绝对残差kW']:.2e} kW, "
        f"相对残差 {chk3['执行平衡相对残差']:.2e}",
        f"- SOC全程在界: {chk3['SOC全程在界']} | 同时充放时段数: {chk3['同时充放时段数']}",
        f"- 紧急购电: {chk3['紧急购电总量kWh']} kWh / {chk3['紧急购电时段数']} 时段; "
        f"Top20%预报偏差时段承载 {chk3['Top20%偏差时段承载紧急电量占比']:.1%}",
        f"- 固定电价参照(365天): Q3提交 {Q3_B_COST:.2f} / Q3日前 {Q3_A_COST:.2f} | "
        f"Q2完美 {Q2_PERFECT_COST:.2f}",
    ]
    (OUT_R / "summary.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print("\n模板回填清单:", "\n".join(report2 + report3) if (report2 or report3) else "(空)")


if __name__ == "__main__":
    main()