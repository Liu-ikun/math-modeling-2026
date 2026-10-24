# -*- coding: utf-8 -*-
"""
问题1 主流程(闭环): 建模->求解->模板原格式回填->润色图表->摘要含侧面验证
路径: D:/Python projects/Math Modeling Project 2026/2026_CUMCM/Problem_1/src/main.py

【执行步骤】(与论文 5.2 节结构一一对应)
  ① 取数   : 附件1_10min对齐.xlsx (processed) -> 失败则原始附件1直读 (双保险)
  ② 求解   : LP(必要时MILP) 最优购电策略 + 无储能基线对照
  ③ 验证   : 合理性/经济性/优越性三组指标 (侧面验证, 当问完成)
  ④ 输出   : result1.xlsx(附件5模板原格式回填) + 验证表 + 论文用csv + 5张图 + summary.md

【关键口径说明】(答辩要点)
  LP 输入必须用"附件1_10min对齐"表: 其 144 行的时刻标签为区间右端点 00:10..24:00,
  行序恰与决策变量下标 t=1..144 一一对应;
  "典型日"表标签为 00:00..23:50 (首行 00:00 存的是 24:00 的值), 直接使用会整体错位一个时段。
"""
import sys
import time
from pathlib import Path

import matplotlib
matplotlib.use("Agg")               # 无界面后端: 只存文件不弹窗, 服务器/批处理环境安全
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))   # 保证同目录 model/algorithm 可导入
from algorithm import (economy_check, fill_result1, make_results,
                       solve_no_storage, solve_storage, validate,
                       write_verification)
from model import DT

# ------------------------- 路径约定 (相对项目根, 可整体迁移) -------------------------
ROOT = Path(__file__).resolve().parents[2]                 # = 2026_CUMCM/
PROC = ROOT / "Problem_1" / "data" / "processed"           # 清洗后数据 (preprocessing.py 产出)
RAW = ROOT / "Problem_1" / "data" / "raw"                  # 原始附件
OUT_T = ROOT / "Problem_1" / "outputs" / "tables"          # 表格输出
OUT_F = ROOT / "Problem_1" / "outputs" / "figures"         # 图形输出
OUT_R = ROOT / "Problem_1" / "outputs" / "report"          # 报告输出
# 附件5 提交模板 (两个候选位置取先存在者); result1.xlsx 必须按其原格式回填
TPL = next((p for p in (ROOT / "Shared_Utils" / "data" / "raw" / "附件5" / "result1.xlsx",
                        ROOT / "Material" / "附件" / "附件5" / "result1.xlsx")
            if p.exists()), None)
for d in (OUT_T, OUT_F, OUT_R):
    d.mkdir(parents=True, exist_ok=True)
# ------------------------- 绘图全局风格 (论文出图统一) -------------------------
plt.rcParams.update({
    "font.sans-serif": ["SimHei", "Microsoft YaHei"],      # 中文字体, 防图内乱码
    "axes.unicode_minus": False,                           # 负号正常显示
    "axes.grid": True, "grid.alpha": 0.3,                  # 淡网格, 便于读数
    "figure.dpi": 110, "savefig.dpi": 200,                 # 存图 200dpi, 满足印刷要求
    "axes.spines.top": False, "axes.spines.right": False,  # 去掉上/右边框, 简洁风格
})
# 五条功率曲线的固定配色 (全论文统一): 负载/光伏/购电/充电/放电
CL, CPV, CG, CCH, CDIS = "#1f77b4", "#2ca02c", "#d62728", "#ff7f0e", "#9467bd"


def pick_col(df, keys):
    """按关键字模糊选列: 依次尝试 keys 中的关键字, 返回第一个命中的列名。
    目的: 对列名细微差异(如"小区负载"/"负载")鲁棒, 避免硬编码列名导致的脆性。"""
    for k in keys:
        for c in df.columns:
            if k in str(c):
                return c
    return None


def from_frame(df, src):
    """从 DataFrame 提取 LP 所需三条 144 长度序列 (电价/负载/光伏)。
    校验: 三列齐备且行数恰为 144 才返回, 否则 None (交给下一级取数兜底)。"""
    cc, lc, pc = (pick_col(df, ["电价"]), pick_col(df, ["负载", "负荷"]),
                  pick_col(df, ["光伏", "发电"]))
    if cc and lc and pc and len(df) == 144:
        return (df[cc].astype(float).tolist(), df[lc].astype(float).tolist(),
                df[pc].astype(float).tolist(), src)
    return None


def load_raw_direct():
    """第二级取数兜底: 绕过 processed, 直接解析原始附件1.xlsx。
    技术要点: 原始时间列是 datetime.time 对象且末行 24:00 被存为 time(0,0),
    直接排序会把 24:00 排到最前 -> 用"时间回退检测"(diff<0 即跨日)累加天数修正,
    使时间轴单调递增为 00:10..24:00, 与 t=1..144 对齐。"""
    df = pd.read_excel(RAW / "附件1.xlsx", sheet_name=0)
    tc = pick_col(df, ["时间"])
    ts = pd.to_datetime(pd.Series([
        pd.Timestamp("2000-01-01") + pd.Timedelta(hours=x.hour, minutes=x.minute,
                                                  seconds=x.second)
        if hasattr(x, "hour") else pd.Timestamp(str(x)) for x in df[tc]]))
    wrap = ts.diff() < pd.Timedelta(0)                     # 时间回退 = 跨日(24:00存成0:00)
    df = df.assign(_t=ts + pd.to_timedelta(wrap.cumsum(), unit="D")).sort_values("_t")
    return from_frame(df, "原始附件1直读")


def load_inputs():
    """两级取数链: processed 对齐表(首选, 清洗管道产物) -> 原始附件1直读(兜底)。
    两级都失败说明清洗管道未运行或附件1损坏, 直接终止并提示。"""
    p = PROC / "附件1_10min对齐.xlsx"
    if p.exists():
        got = from_frame(pd.read_excel(p), "processed附件1_10min对齐")
        if got:
            return got
    got = load_raw_direct()
    if got:
        return got
    raise SystemExit("processed 与原始附件1 均取数失败, 请先运行 preprocessing.py")


def main():
    t0 = time.time()
    # ---------- ① 取数 ----------
    c, L, PV, src = load_inputs()

    # ---------- ② 求解: 有储能最优策略 + 无储能基线 ----------
    res = solve_storage(c, L, PV)
    res["_L"], res["_PV"] = L, PV                # 注入原始输入, 供 validate 复算功率平衡
    chk = validate(res, c)                       # 合理性验证指标
    no_res = solve_no_storage(c, L, PV)          # 无储能基线 (优越性对比用)
    eco = economy_check(res, no_res)             # 经济性验证指标
    slots, totals, blocks, e0, e24, full = make_results(res, c)   # 表1/表2/明细

    # ---------- ③ 输出: result1.xlsx (模板原格式回填) + 验证表 + 论文csv ----------
    if TPL is None:
        raise SystemExit("未找到附件5模板 result1.xlsx, 无法原格式回填")
    report = fill_result1(TPL, OUT_T / "result1.xlsx",
                          slots, totals, blocks, e0, e24, full)
    write_verification(OUT_T / "result1_verification.xlsx", eco, chk)
    # 论文表1/表2 的 csv 版 (utf-8-sig 带 BOM, Excel 双击打开中文不乱码)
    pd.DataFrame([[k, v] for k, v in slots.items()],
                 columns=["时间段", "购电量kWh"]).to_csv(
        OUT_T / "table1_paper.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame([[k, ch, dis] for k, (ch, dis) in blocks.items()],
                 columns=["时间段", "充电量kWh", "放电量kWh"]).to_csv(
        OUT_T / "table2_paper.csv", index=False, encoding="utf-8-sig")

    buy = full["购电量kWh"].tolist()
    # ---------- ④ 五张论文图 ----------
    # 图1 功率平衡: 光伏面积填充 + 负载/购电/充/放四线, 直观展示"供给=需求"逐时段成立
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.fill_between(range(144), PV, color=CPV, alpha=0.25, label="光伏预测")
    ax.plot(L, color=CL, lw=1.6, label="小区负载")
    ax.plot(res["G"], color=CG, lw=1.6, label="购电功率")
    ax.plot(res["Pch"], "--", color=CCH, lw=1.2, label="充电")
    ax.plot(res["Pdis"], "--", color=CDIS, lw=1.2, label="放电")
    ax.set_xlabel("10min 时段")
    ax.set_ylabel("功率 (kW)")
    ax.set_title("图1 典型日功率平衡（LP 最优策略）")
    ax.legend(ncol=5, loc="upper center", bbox_to_anchor=(0.5, -0.12), frameon=False)
    fig.tight_layout()
    fig.savefig(OUT_F / "fig1_power_balance.png", bbox_inches="tight")
    plt.close(fig)

    # 图2 电价-购电量双轴: 柱=购电量, 阶梯线=电价, 验证"低价多购、高价少购"套利形态
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.bar(range(144), buy, color="#4c72b0", alpha=0.85, label="购电量")
    k = int(max(range(144), key=lambda i: buy[i]))         # 购电量峰值时段(自动标注)
    ax.annotate(f"峰值 {buy[k]:.0f} kWh\n时段{k + 1}", (k, buy[k]),
                xytext=(k + 8, buy[k] * 0.9), arrowprops=dict(arrowstyle="->", lw=0.8),
                fontsize=9)
    ax2 = ax.twinx()                                       # 右轴: 电价阶梯线
    ax2.step(range(144), c, where="post", color="#c44e52", lw=1.4, label="电价")
    ax2.set_ylabel("电价 (元/kWh)")
    ax2.spines["right"].set_visible(True)
    ax.set_xlabel("10min 时段")
    ax.set_ylabel("购电量 (kWh)")
    ax.set_title("图2 分时电价与购电量（低价多购、高价少购）")
    fig.tight_layout()
    fig.savefig(OUT_F / "fig2_price_purchase.png")
    plt.close(fig)

    # 图3 SOC 轨迹: 底色带=[1200,10800]可行域, 验证"凌晨充电、晚峰放电、首末=6000"
    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.axhspan(1200, 10800, color="#2ca02c", alpha=0.07)
    ax.plot(res["E"], color="#2ca02c", lw=1.8, label="储电量 E(t)")
    ax.axhline(6000, color="gray", ls="--", lw=1)          # E0 参考线
    ax.scatter([0, 143], [e0, e24], color="#d62728", zorder=5,
               label=f"E0=E144={e0:.0f} kWh")
    ax.set_xlabel("10min 时段")
    ax.set_ylabel("储电量 (kWh)")
    ax.set_title("图3 储能 SOC 轨迹（凌晨充电、晚峰放电）")
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(OUT_F / "fig3_soc.png")
    plt.close(fig)

    # 图4 经济价值柱状对比: 无储能基线 vs LP有储能, 柱顶标注具体金额
    fig, ax = plt.subplots(figsize=(6.4, 4.2))
    bars = ax.bar(["无储能基线", "LP 有储能"], [no_res["obj"], res["obj"]],
                  color=["#9aa0a6", "#4c72b0"], width=0.55)
    ax.bar_label(bars, fmt="%.0f 元", fontsize=10)
    ax.set_ylabel("全天购电费 (元)")
    ax.set_title(f"图4 储能经济价值：节省 {eco['节省购电费(元)']} 元"
                 f"（{eco['节省比例']:.1%}）")
    ax.set_ylim(0, no_res["obj"] * 1.15)
    fig.tight_layout()
    fig.savefig(OUT_F / "fig4_cost_comparison.png")
    plt.close(fig)

    # 图5 累计成本曲线: 两线间距=逐时段节省的累积, 阴影面积直观呈现节省发生的时间分布
    cum_w = pd.Series(c).mul(res["G"]).mul(DT).cumsum()    # 有储能累计购电费
    cum_n = pd.Series(c).mul(no_res["G"]).mul(DT).cumsum() # 无储能累计购电费
    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.plot(cum_n, color="#9aa0a6", lw=1.6, label="无储能基线")
    ax.plot(cum_w, color="#4c72b0", lw=1.6, label="LP 有储能")
    ax.fill_between(range(144), cum_w, cum_n, color="#4c72b0", alpha=0.12)
    ax.set_xlabel("10min 时段")
    ax.set_ylabel("累计购电费 (元)")
    ax.set_title("图5 累计购电费对比（阴影=储能节省）")
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(OUT_F / "fig5_cumulative_cost.png")
    plt.close(fig)

    # ---------- ⑤ summary.md: 三段式摘要(结果->合理性验证->经济性验证), 可直接改写进论文 ----------
    lines = [
        "# 问题1 求解摘要（闭环：建模 → 求解 → 结果分析含侧面验证）",
        "",
        f"- 数据口径：{src}",
        f"- 求解模式：{res['mode']}｜求解状态：{res['status']}｜耗时 {time.time() - t0:.2f} s",
        f"- result1.xlsx：按附件5模板原格式回填（回填 {len(report)} 处, 格式零改动）",
        "",
        "## 一、求解结果",
        f"- 全天购电量：{totals['buy']:.2f} kWh",
        f"- 全天购电费：{totals['cost']:.2f} 元",
        f"- 指定时段购电量： " + "；".join(f"{k}={v:.2f}" for k, v in slots.items()),
        "",
        "## 二、侧面验证（合理性）",
        f"- 功率平衡最大绝对残差 {chk['功率平衡最大绝对残差kW']:.2e} kW，"
        f"相对残差 {chk['功率平衡相对残差']:.2e}（求解器容差内，能量守恒成立）",
        f"- SOC 全程在界：{chk['SOC全程在界']}｜E0/E144：{chk['E0/E144'][0]:.0f}/"
        f"{chk['E0/E144'][1]:.0f} kWh",
        f"- 同时充放时段数：{chk['同时充放时段数']}｜弃光电量：{chk['弃光电量kWh']} kWh",
        f"- 低价段充电占比 {chk['低价段充电占比']}｜高价段放电占比 {chk['高价段放电占比']}"
        "（接近 1，符合低充高放规律）",
        "",
        "## 三、侧面验证（经济性与优越性）",
        f"- 有储能 {eco['有储能购电费(元)']} 元 vs 无储能 {eco['无储能购电费(元)']} 元",
        f"- 节省 {eco['节省购电费(元)']} 元（{eco['节省比例']:.1%}）",
        f"- 储能吞吐电量 {chk['储能吞吐电量kWh']} kWh",
    ]
    (OUT_R / "summary.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print("\n回填坐标清单:")            # 审计日志: 每处回填的 sheet!坐标<-标签, 供人工核对
    print("\n".join(report))


if __name__ == "__main__":
    main()