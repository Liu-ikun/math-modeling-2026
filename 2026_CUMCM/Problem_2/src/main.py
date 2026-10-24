# -*- coding: utf-8 -*-
"""问题2 主流程: 取数(附件1电价+附件2负载/光伏) -> 三策略求解 -> 侧面验证 -> 模板回填 -> 图表
路径: 2026_CUMCM/Problem_2/src/main.py
运行: python main.py [--pv-from-fc] [--report-start 2025-02-01]"""
import sys
import time
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from algorithm import (daily_agg, economy_check, fill_wide, long_table,
                       make_date_tables, solve_myopic, solve_strategy,
                       validate, verify_fill, write_full)
from model import DT

ROOT = Path(__file__).resolve().parents[2]
PROC = ROOT / "Problem_2" / "data" / "processed"
OUT_T = ROOT / "Problem_2" / "outputs" / "tables"
OUT_F = ROOT / "Problem_2" / "outputs" / "figures"
OUT_R = ROOT / "Problem_2" / "outputs" / "report"
TPL = next((p for p in (ROOT / "Shared_Utils" / "data" / "raw" / "附件5" / "result2.xlsx",
                        ROOT / "Material" / "附件" / "附件5" / "result2.xlsx")
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


def to_matrix(df, val_col):
    """长表 -> (D,144) 矩阵: 按完整时间戳排序后 reshape。
    '0:00+1' 行时间戳为次日00:00, 排序后恰落在当天区块末尾(第144步), 无错位。"""
    ts = pd.to_datetime(df["时间"])
    df = df.assign(_ts=ts).sort_values("_ts")
    vals = df[val_col].astype(float).to_numpy()
    if len(vals) % 144:
        raise SystemExit(f"长表行数 {len(vals)} 不是 144 的整数倍, 无法整形成 (D,144)")
    dates = df["_ts"].dt.strftime("%Y-%m-%d").to_numpy()[::144].tolist()
    return vals.reshape(-1, 144).tolist(), dates


def load_year_raw():
    """原始附件2.xlsx 直读: 两个 sheet(小区负载/光伏发电实际功率),
    列序天然为 00:10..23:50,0:00+1, 即 t=1..144 右端点约定, 无需重排。"""
    raw = ROOT / "Problem_2" / "data" / "raw" / "附件2.xlsx"
    sheets = pd.ExcelFile(raw).sheet_names
    load_sn = next((s for s in sheets if any(k in s for k in ("负荷", "负载"))), None)
    pv_sn = next((s for s in sheets if any(k in s for k in ("光伏", "发电"))), None)
    if not load_sn:
        return None, None, None

    def parse(sn):
        df = pd.read_excel(raw, sheet_name=sn, header=0)
        df = df.rename(columns={df.columns[0]: "日期"})
        df["日期"] = pd.to_datetime(df["日期"])
        df = df.sort_values("日期")
        dates = df["日期"].dt.strftime("%Y-%m-%d").tolist()
        vals = df.drop(columns=["日期"]).astype(float).to_numpy().tolist()
        return dates, vals

    dates, L = parse(load_sn)
    PV = parse(pv_sn)[1] if pv_sn else None
    return L, PV, dates


def load_year_data(use_fc_pv):
    """装载全年 L/PV: 首选 processed 长表(关键字定位序列);
    长表缺光伏序列时回退原始附件2双sheet直读; 最后兜底 --pv-from-fc(附件3预报)"""
    d2 = pd.read_excel(PROC / "附件2_长表.xlsx")
    print("附件2长表列名:", list(d2.columns))
    val_col = pick_col(d2, ["数值", "값"])
    var_col = pick_col(d2, ["变量", "序列"])
    series = {}
    if var_col:
        for name, grp in d2.groupby(var_col):
            series[str(name)] = grp
    else:
        series["负荷"] = d2
    keys = list(series)
    print("附件2长表变量:", keys)
    # 按关键字定位序列, 不能按排序取首个: "光"(U+5149) < "小"(U+5C0F), 光伏会排在负载前
    l_key = next((k for k in keys if any(x in k for x in ("负荷", "负载"))), keys[0])
    pv_key = next((k for k in keys if any(x in k for x in ("光伏", "发电"))), None)
    L, dates = to_matrix(series[l_key], val_col)
    if pv_key:
        PV, _ = to_matrix(series[pv_key], val_col)
    else:
        Lr, PVr, datesr = load_year_raw()
        if PVr is not None:
            print("长表无光伏序列 -> 改用原始附件2双sheet直读(负荷与光伏同源, 口径一致)")
            L, PV, dates = Lr, PVr, datesr
        elif use_fc_pv:
            d3 = pd.read_excel(PROC / "附件3_10min宽表.xlsx")
            d3["日期"] = pd.to_datetime(d3["时间"]).dt.strftime("%Y-%m-%d")
            d3["时刻"] = pd.to_datetime(d3["时间"]).dt.strftime("%H:%M")
            fc = d3[d3["时刻"] == "00:00"].set_index("日期")
            PV = []
            for d in dates:
                leads = [float(fc.loc[d, f"预报{h}小时"]) for h in range(1, 25)]
                PV.append([leads[min(23, t // 6)] for t in range(144)])
            print("警告: 光伏采用附件3零时发布预报兜底口径")
        else:
            raise SystemExit("长表与原始附件2均无光伏序列, 可加 --pv-from-fc 用附件3预报兜底")
    return L, PV, dates


def main():
    t0 = time.time()
    use_fc_pv = "--pv-from-fc" in sys.argv
    report_start = None
    if "--report-start" in sys.argv:
        report_start = sys.argv[sys.argv.index("--report-start") + 1]
    c_day = load_price()
    L, PV, dates = load_year_data(use_fc_pv)
    D = len(dates)
    print(f"数据: {D} 天 ({dates[0]} ~ {dates[-1]}), 报告起始: {report_start or '全部'}")

    res_b = solve_strategy(c_day, L, PV, storage=True)     # 策略B: 全年全局优化
    res_a = solve_strategy(c_day, L, PV, storage=False)    # 策略A: 无储能基线
    res_c = solve_myopic(c_day, L, PV)                     # 策略C: 逐日滚动对照
    chk = validate(res_b, c_day, L, PV)
    eco = economy_check(res_b, res_a, res_c)
    daily_b = daily_agg(res_b, c_day, dates)
    daily_a = daily_agg(res_a, c_day, dates)
    win = [d >= "2025-02-01" for d in dates]
    cost_win = float(daily_b["购电费(元)"][win].sum())
    buy_win = float(daily_b["购电量kWh"][win].sum())
    base_spill = sum(max(PV[n][t] - L[n][t], 0.0) * DT
                     for n in range(D) for t in range(144))
    spill_b = chk["弃光总量kWh"]
    if report_start:
        daily_b = daily_b[daily_b["日期"] >= report_start].reset_index(drop=True)
        daily_a = daily_a[daily_a["日期"] >= report_start].reset_index(drop=True)

    if TPL is None:
        print("未找到附件5模板 result2.xlsx, 跳过模板回填")
        report, ok = [], None
    else:
        report = fill_wide(TPL, OUT_T / "result2.xlsx", res_b, dates, c_day)
        ok = verify_fill(OUT_T / "result2.xlsx", dates)
    write_full(OUT_T / "result2_full.xlsx", daily_b, long_table(res_b, c_day, dates, report_start),
               eco, chk)
    daily_b.to_csv(OUT_T / "daily_summary.csv", index=False, encoding="utf-8-sig")

    dt4 = make_date_tables(res_b, c_day, dates)
    rows1, rows2, rows3 = [], [], []
    for key, tb in dt4.items():
        for lab, v in tb["slots"].items():
            rows1.append([key, lab, round(v, 4)])
        rows1.append([key, "全天购电量", round(tb["totals"]["buy"], 4)])
        rows1.append([key, "全天购电费", round(tb["totals"]["cost"], 4)])
        for lab, (ch, dis) in tb["blocks"].items():
            rows2.append([key, lab, round(ch, 4), round(dis, 4)])
        rows2.append([key, "0:00储电量", round(tb["e0"], 4), ""])
        rows2.append([key, "24:00储电量", round(tb["e24"], 4), ""])
        for lab, v in tb["t3"]:
            rows3.append([key, lab, v])
    pd.DataFrame(rows1, columns=["日期", "时间段/项目", "计划购电量kWh"]).to_csv(
        OUT_T / "table1_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows2, columns=["日期", "时间段/项目", "充电量kWh", "放电量kWh"]).to_csv(
        OUT_T / "table2_4dates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(rows3, columns=["日期", "时间段", "紧急购电量kWh"]).to_csv(
        OUT_T / "table3_4dates.csv", index=False, encoding="utf-8-sig")

    x = range(len(daily_b))
    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(x, daily_a["购电费(元)"], color="#9aa0a6", lw=1.2, label="无储能基线")
    ax.plot(x, daily_b["购电费(元)"], color="#4c72b0", lw=1.2, label="全年全局优化")
    ax.set_xlabel("日序号")
    ax.set_ylabel("日购电费 (元)")
    ax.set_title("图1 逐日购电费对比（全局优化 vs 无储能）")
    ax.legend(frameon=False)
    fig.tight_layout(); fig.savefig(OUT_F / "fig1_daily_cost.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(x, daily_a["购电费(元)"].cumsum(), color="#9aa0a6", lw=1.6, label="无储能基线")
    ax.plot(x, daily_b["购电费(元)"].cumsum(), color="#4c72b0", lw=1.6, label="全年全局优化")
    ax.set_xlabel("日序号"); ax.set_ylabel("累计购电费 (元)")
    ax.set_title("图2 全年累计购电费对比")
    ax.legend(frameon=False)
    fig.tight_layout(); fig.savefig(OUT_F / "fig2_cumulative_cost.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(x, daily_a["弃光电量kWh"], color="#9aa0a6", lw=1.2, label="无储能弃光")
    ax.plot(x, daily_b["弃光电量kWh"], color="#2ca02c", lw=1.2, label="有储能弃光")
    ax.set_xlabel("日序号"); ax.set_ylabel("弃光电量 (kWh)")
    ax.set_title("图3 逐日弃光对比（储能消纳盈余光伏）")
    ax.legend(frameon=False)
    fig.tight_layout(); fig.savefig(OUT_F / "fig3_spill.png"); plt.close(fig)

    T = 144
    for pick, name in [(0, "首日"), (len(daily_b) // 2, "年中"), (len(daily_b) - 1, "末日")]:
        n0 = pick * T
        fig, ax = plt.subplots(figsize=(11, 4.2))
        ax.plot(res_b["E"][n0:n0 + T], color="#2ca02c", lw=1.6, label="储电量")
        ax2 = ax.twinx()
        ax2.step(range(T), c_day, where="post", color="#c44e52", lw=1.2, label="电价")
        ax.set_xlabel("10min 时段"); ax.set_ylabel("储电量 (kWh)")
        ax2.set_ylabel("电价 (元/kWh)")
        ax.set_title(f"图4 {name} SOC 与电价（低充高放）")
        fig.tight_layout(); fig.savefig(OUT_F / f"fig4_soc_{name}.png"); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 4.2))
    bars = ax.bar(["无储能基线", "全年全局优化", "逐日滚动"],
                  [res_a["obj"], res_b["obj"], res_c["obj"]],
                  color=["#9aa0a6", "#4c72b0", "#d62728"], width=0.55)
    ax.bar_label(bars, fmt="%.0f 元", fontsize=9)
    ax.set_ylabel("全年购电费 (元)")
    ax.set_title("图5 三策略经济性对比（全局优于短视）")
    fig.tight_layout(); fig.savefig(OUT_F / "fig5_strategy_compare.png"); plt.close(fig)

    lines = [
        "# 问题2 求解摘要（闭环：建模 → 求解 → 结果分析含侧面验证）",
        "",
        f"- 数据: {D} 天 ({dates[0]} ~ {dates[-1]}), 报告窗口: {report_start or '全部'}",
        f"- 策略B 状态: {res_b['status']} | 求解耗时 {time.time() - t0:.1f} s",
        f"- result2.xlsx 回填: {len(report)} 个sheet | 覆盖率校验: "
        f"{'通过' if ok else '未通过(见上方坐标)'}",
        "",
        "## 一、求解结果",
        f"- 全年购电费: {res_b['obj']:.2f} 元 | 全年购电量: {res_b['buy']:.2f} kWh",
        f"- 紧急购电总量: {chk['紧急购电总量kWh']} kWh (完美信息下应为 0)",
        "",
        "## 二、侧面验证（合理性）",
        f"- 功率平衡最大绝对残差 {chk['功率平衡最大绝对残差kW']:.2e} kW, "
        f"相对残差 {chk['功率平衡相对残差']:.2e}",
        f"- SOC 全程在界: {chk['SOC全程在界']} | 同时充放时段数: {chk['同时充放时段数']}",
        f"- 弃光总量: {chk['弃光总量kWh']} kWh",
        "",
        "## 三、侧面验证（经济性与优越性）",
        f"- 无储能基线: {eco['无储能全年购电费(元)']} 元",
        f"- 有储能全局优化: {eco['有储能全年购电费(元)']} 元, "
        f"节省 {eco['节省购电费(元)']} 元 ({eco['节省比例']})",
        f"- 逐日滚动对照: {eco['逐日滚动全年购电费(元)']} 元, "
        f"全局优化再省 {eco['全局优化优于逐日滚动(元)']} 元",
        "",
        "## 四、模板回填",
        f"- 模板窗口(2025-02-01~12-31, 334天, 与result2.xlsx回填行一致): "
        f"购电费 {cost_win:.2f} 元 | 购电量 {buy_win:.2f} kWh",
        f"- 弃光消纳: 无储能基线弃光 {base_spill:.2f} kWh -> 有储能 {spill_b:.2f} kWh, "
        f"消纳率提升 {(base_spill - spill_b) / base_spill:.2%}",
    ]
    (OUT_R / "summary.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print("\n模板回填清单:", "\n".join(report) if report else "(空)")


if __name__ == "__main__":
    main()