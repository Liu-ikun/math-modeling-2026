# -*- coding: utf-8 -*-
"""问题3 主流程: 取数(电价/实际负载光伏/四时点预报) -> MPC滚动+日前对照+频率实验
   + 校正强度q扫描 -> 验证 -> 回填 -> 图表
运行: python main.py [--fast(跳过频率实验)] [--freq-sample N] [--q 0.6(单点校正)]
      [--check(双求解器单阶段隔离对照)]"""
import sys
import time
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import concurrent.futures as cf
from algorithm import (DATE_KEYS, Q2_PERFECT_COST, UPDATES, build_fc_stats,
                       build_virtual_fc, calibrate_shift, fc_series, fill_wide3,
                       make_date_tables, robust_adj, run_chain, set_fc_shift,
                       validate, verify_alignment, verify_fill3)
from model import DT, E0

ROOT = Path(__file__).resolve().parents[2]
PROC = ROOT / "Problem_2" / "data" / "processed"
OUT_T = ROOT / "Problem_3" / "outputs" / "tables"
OUT_F = ROOT / "Problem_3" / "outputs" / "figures"
OUT_R = ROOT / "Problem_3" / "outputs" / "report"
TPL = next((p for p in (ROOT / "Shared_Utils" / "data" / "raw" / "附件5" / "result3.xlsx",
                        ROOT / "Material" / "附件" / "附件5" / "result3.xlsx")
            if p.exists()), None)
for d in (OUT_T, OUT_F, OUT_R):
    d.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({"font.sans-serif": ["SimHei", "Microsoft YaHei"],
                     "axes.unicode_minus": False, "axes.grid": True, "grid.alpha": 0.3,
                     "savefig.dpi": 200, "axes.spines.top": False, "axes.spines.right": False})


def pick_col(df, keys):
    for k in keys:
        for c in df.columns:
            if k in str(c):
                return c
    return None


def load_price():
    df = pd.read_excel(PROC / "附件1_10min对齐.xlsx")
    return df[pick_col(df, ["电价"])].astype(float).tolist()


def load_actual():
    """附件2双sheet直读: 实际负载/光伏 (D,144)"""
    raw = ROOT / "Problem_2" / "data" / "raw" / "附件2.xlsx"

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
    return L, PV, dates


def load_fc():
    """附件3原始直读: 列布局=[日期(稀疏,需ffill), 预报时刻, 预报1小时..预报24小时]。
    严禁使用processed'附件3_10min宽表'——它经long_to_10min在发布时点之间线性插值,
    按行覆盖后留下XX:50行 ≈ 0.86·本次发布 + 0.14·下次发布(lead指向6h后时段),
    日出/日落爬坡段系统性失真(RMSE地板700kW、日出边沿-6h假错位均源于此)"""
    df = pd.read_excel(ROOT / "Problem_3" / "data" / "raw" / "附件3.xlsx", header=0)
    cols = [str(c) for c in df.columns]
    print(f"附件3原始列名({len(cols)}列): {cols[:4]} ... {cols[-2:]}")

    def lead_no(c):
        m = re.search(r"预报\s*(\d+)\s*小时", c)
        return int(m.group(1)) if m else None

    cand = [(lead_no(c), c) for c in cols]
    lead_cols = [c for n, c in sorted((x for x in cand if x[0] is not None),
                                      key=lambda x: x[0])]
    if len(lead_cols) != 24:
        lead_cols = cols[2:2 + 24]
        print(f"[load_fc] '预报k小时'正则识别到{len(lead_cols)}列, "
              f"退回位置法取第3~26列作为lead1~24")
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
    print(f"附件3预报(原始直读): {len(fcmap)} 天 × "
          f"{sorted({H for v in fcmap.values() for H in v})} 时点发布 | lead列数 {len(lead_cols)}")
    return fcmap


_G = {}


def _init(cc, LL, PP, FF, DD, VV, AA):
    _G.update(c=cc, L=LL, PV=PP, fc=FF, dates=DD, vmap=VV, fadj=AA)


def _work(item):
    name, sch, idxs, use_v, qkey = item
    fadj = _G["fadj"].get(qkey) if qkey is not None else None
    return name, run_chain(_G["c"], _G["L"], _G["PV"], _G["fc"], _G["dates"],
                           idxs, sch, vmap=_G["vmap"] if use_v else None,
                           fadj=fadj)


def main():
    t0 = time.time()
    fast = "--fast" in sys.argv
    n_smp = 60
    if "--freq-sample" in sys.argv:
        n_smp = int(sys.argv[sys.argv.index("--freq-sample") + 1])
    c = load_price()
    L, PV, dates = load_actual()
    fcmap = load_fc()
    D = len(dates)
    print(f"数据: {D} 天 ({dates[0]} ~ {dates[-1]})")

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
                  f"{sum(std[H, k] for k in ks) / len(ks):.1f} kW | 样本lead数 {len(ks)}")
    qs = (0.5, 0.6, 0.7, 0.8, 0.9)
    if "--q" in sys.argv:
        qs = (float(sys.argv[sys.argv.index("--q") + 1]),)
    fadj_multi = {q: robust_adj(errs, q_plan=q, q_adj=q) for q in qs}
    print(f"[校正强度q扫描] q∈{qs}, 经验确定最优校正强度(报童理论仅给出点估计)")

    if "--check" in sys.argv:
        from model import build_stage, solve as psolve, solve_stage
        d, pv0 = 0, fc_series(fcmap, dates[0], 0)
        s1 = solve_stage(c, L[d], pv0, e_init=E0, mode="plan")
        st1, ov1 = psolve(build_stage(c, L[d], pv0, e_init=E0, mode="plan")[0])
        r1 = abs(s1["obj"] - ov1) / max(1.0, abs(ov1))
        pv6 = fc_series(fcmap, dates[0], 6, 36)
        s2 = solve_stage(c[36:], L[d][36:], pv6, plan=s1["G"][36:],
                         e_init=s1["E"][37], mode="adj")
        st2, ov2 = psolve(build_stage(c[36:], L[d][36:], pv6,
                                      plan=s1["G"][36:], e_init=s1["E"][37],
                                      mode="adj")[0])
        r2 = abs(s2["obj"] - ov2) / max(1.0, abs(ov2))
        print(f"[双求解器单阶段隔离对照] plan段 {r1:.2e} | adj段 {r2:.2e} (应<1e-6)")

    vmap = None if fast else build_virtual_fc(PV, dates, errs)
    full = list(range(D))
    chains = [("B", UPDATES, full, False, None),
              ("A", (0,), full, False, None),
              ("A2", (0,), full, False, None)]
    for q in qs:
        chains.append((f"AQ{int(q * 10):02d}", (0,), full, False, q))
        chains.append((f"BQ{int(q * 10):02d}", UPDATES, full, False, q))
    if not fast:
        step = max(1, D // n_smp)
        sample = list(range(0, D, step))
        chains += [("F1", (0,), sample, False, None),
                   ("F4", UPDATES, sample, False, None),
                   ("F6", tuple(range(0, 24, 4)), sample, False, None),
                   ("F12", tuple(range(0, 24, 2)), sample, False, None),
                   ("F24", tuple(range(24)), sample, False, None),
                   ("V6", tuple(range(0, 24, 4)), sample, True, None),
                   ("V12", tuple(range(0, 24, 2)), sample, True, None),
                   ("V24", tuple(range(24)), sample, True, None)]
    with cf.ProcessPoolExecutor(initializer=_init,
                                initargs=(c, L, PV, fcmap, dates, vmap,
                                          fadj_multi)) as ex:
        out = dict(ex.map(_work, chains))
    cost = {k: sum(r["cost"] for r in out[k]) for k in out}
    cost_b, cost_a = cost["B"], cost["A"]
    best_key = min((k for k in cost if not k.startswith(("F", "V", "A2"))),
                   key=lambda k: cost[k])
    days_b = out[best_key]
    days_a = out["A"]
    print(f"[提交策略] 全链最低成本: {best_key} = {cost[best_key]:.2f} 元")
    print(f"[校正强度q扫描] " + " | ".join(
        f"q={q}: A={cost[f'AQ{int(q * 10):02d}']:.2f} "
        f"B={cost[f'BQ{int(q * 10):02d}']:.2f}" for q in qs))
    print(f"求解完成: {time.time() - t0:.1f} s (含{len(chains)}条并发链)")

    freq, vfreq = {}, {}
    if not fast:
        fmap = {"F1": "仅0:00(日前)", "F4": "4次(0/6/12/18)", "F6": "6次(每4h)",
                "F12": "12次(每2h)", "F24": "24次(每1h)"}
        for k, name in fmap.items():
            freq[name] = sum(r["cost"] for r in out[k]) / len(out[k])
            print(f"[频率实验·仅重优化] {name}: 样本{len(out[k])}天日均 {freq[name]:.2f} 元")
        for k, name in (("V6", "6次(每4h)"), ("V12", "12次(每2h)"), ("V24", "24次(每1h)")):
            vfreq[name] = sum(r["cost"] for r in out[k]) / len(out[k])
            print(f"[频率实验·理想发布(完全信息上界)] {name}: 样本{len(out[k])}天日均 {vfreq[name]:.2f} 元")
    d_aa = abs(cost["A"] - cost["A2"]) / max(1.0, cost["A"])
    print(f"[自洽检查] A vs A2 重复链相对差 {d_aa:.2e} (并发确定性, 应=0)")

    chk = validate(days_b, c, L, PV, fcmap, dates)
    report = fill_wide3(TPL, OUT_T / "result3.xlsx", days_b, dates) if TPL else []
    ok = verify_fill3(OUT_T / "result3.xlsx", dates) if TPL else None

    rows = []
    for rec, dstr in zip(days_b, dates):
        rows.append([dstr, round(rec["cost"], 2), round(rec["buy"], 2),
                     round(rec["emer_kwh"], 4), round(rec["spill"] * DT, 2)])
    daily = pd.DataFrame(rows, columns=["日期", "结算费用(元)", "调整购电量kWh",
                                        "紧急购电量kWh", "弃光电量kWh"])
    daily.to_csv(OUT_T / "daily_summary_q3.csv", index=False, encoding="utf-8-sig")
    r1, r2, r3 = [], [], []
    for key in DATE_KEYS:
        tb = make_date_tables(days_b[dates.index(key)], c, dates.index(key))
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
        OUT_T / "table1_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r2, columns=["日期", "时间段/项目", "充电量kWh", "放电量kWh"]).to_csv(
        OUT_T / "table2_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(r3, columns=["日期", "时间段", "紧急购电量kWh"]).to_csv(
        OUT_T / "table3_4dates.csv", index=False, encoding="utf-8-sig")

    fig, ax = plt.subplots(figsize=(8, 4.4))
    names = ["问题2完美信息下界", "朴素MPC(B)", "朴素日前(A)", f"提交策略({best_key})"]
    vals = [Q2_PERFECT_COST, cost_b, cost_a, cost[best_key]]
    bars = ax.bar(names, vals, color=["#9aa0a6", "#4c72b0", "#d62728", "#55a868"],
                  width=0.55)
    ax.bar_label(bars, fmt="%.0f 元", fontsize=9)
    ax.set_ylabel("全年结算费用 (元)")
    ax.set_title("图1 成本阶梯: 完美信息 ≤ MPC ≤ 日前 + 提交策略")
    fig.tight_layout(); fig.savefig(OUT_F / "fig1_cost_ladder.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 4))
    qxs = list(qs)
    ax.plot(qxs, [cost[f"AQ{int(q * 10):02d}"] / 1e6 for q in qxs], "o-",
            label="日前A(q)")
    ax.plot(qxs, [cost[f"BQ{int(q * 10):02d}"] / 1e6 for q in qxs], "s-",
            label="MPC-B(q)")
    ax.set_xlabel("校正分位 q"); ax.set_ylabel("全年成本 (百万元)")
    ax.set_title("图4 校正强度q敏感性: 经验最优q*")
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(OUT_F / "fig4_q_sweep.png"); plt.close(fig)

    if freq:
        fig, ax = plt.subplots(figsize=(8, 4.2))
        xs = [1, 4, 6, 12, 24]
        ys1 = [freq[k] for k in ("仅0:00(日前)", "4次(0/6/12/18)", "6次(每4h)",
                                 "12次(每2h)", "24次(每1h)")]
        ax.plot(xs, ys1, "o-", color="#4c72b0", label="仅加密重优化(预报仍4次发布)")
        if vfreq:
            ys2 = [ys1[0], ys1[1], vfreq["6次(每4h)"], vfreq["12次(每2h)"],
                   vfreq["24次(每1h)"]]
            ax.plot(xs, ys2, "s--", color="#dd8452", label="加密预报发布·理想信息上界")
        ax.set_xscale("log", base=2)
        ax.set_xlabel("每日更新/发布次数 (log2)")
        ax.set_ylabel("样本日均结算费用 (元)")
        ax.set_title("图2 更新频率双曲线: 重优化价值 vs 新信息价值")
        ax.legend(frameon=False, fontsize=8)
        fig.tight_layout(); fig.savefig(OUT_F / "fig2_freq_value.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4.2))
    for key in DATE_KEYS:
        rec = days_b[dates.index(key)]
        ax.plot(rec["E"], label=key)
    ax.axhline(1200, ls="--", lw=0.8, color="k")
    ax.axhline(10800, ls="--", lw=0.8, color="k")
    ax.set_xlabel("10min 时段"); ax.set_ylabel("储电量 (kWh)")
    ax.set_title("图3 指定日期 SOC 轨迹"); ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(OUT_F / "fig3_soc_dates.png"); plt.close(fig)

    lines = [
        "# 问题3 求解摘要（闭环：建模 → 求解 → 结果分析含侧面验证）",
        "",
        f"- 数据: {D} 天 | MPC状态: 完成 | 总耗时 {time.time() - t0:.1f} s",
        f"- result3.xlsx 回填: {len(report)} 个sheet | 覆盖率校验: "
        f"{'通过' if ok else '未通过/未做'}",
        "",
        "## 一、成本阶梯与风险校正价值(创新点核心)",
        f"- 问题2完美信息下界: {Q2_PERFECT_COST:.2f} 元",
        f"- 朴素日前A: {cost_a:.2f} | 朴素MPC(B): {cost_b:.2f} "
        f"(滚动更新信息价值 {cost_a - cost_b:.2f} 元)",
        f"- 提交策略: {best_key} = {cost[best_key]:.2f} 元 "
        f"(较朴素日前省 {cost_a - cost[best_key]:.2f} 元)",
        f"- 校正强度q扫描: " + " | ".join(
            f"q={q}: A={cost[f'AQ{int(q * 10):02d}']:.2f} "
            f"B={cost[f'BQ{int(q * 10):02d}']:.2f}" for q in qs),
        "",
        "## 二、侧面验证",
        f"- 执行平衡最大绝对残差 {chk['执行平衡最大绝对残差kW']:.2e} kW, "
        f"相对残差 {chk['执行平衡相对残差']:.2e}",
        f"- SOC全程在界: {chk['SOC全程在界']} | 同时充放时段数: {chk['同时充放时段数']}",
        f"- 紧急购电: {chk['紧急购电总量kWh']} kWh / {chk['紧急购电时段数']} 时段; "
        f"Top20%预报偏差时段承载 {chk['Top20%偏差时段承载紧急电量占比']:.1%}",
    ]
    if freq:
        lines += ["", "## 三、更新频率实验(样本日均, 元)",
                  "- 仅重优化(信息不变):"]
        lines += [f"  - {k}: {v:.2f}" for k, v in freq.items()]
        if vfreq:
            lines += ["- 理想发布(完全信息上界, 虚拟预报ε≡0):"]
            lines += [f"  - {k}: {v:.2f}" for k, v in vfreq.items()]
    (OUT_R / "summary.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))


if __name__ == "__main__":
    main()