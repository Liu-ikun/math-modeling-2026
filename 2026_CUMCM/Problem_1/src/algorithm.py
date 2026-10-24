# -*- coding: utf-8 -*-
"""
问题1 求解/验证编排 + 附件5模板原格式回填
路径: D:/Python projects/Math Modeling Project 2026/2026_CUMCM/Problem_1/src/algorithm.py

【本层职责】(对应论文 5.2.2 求解 与 5.2.4 结果分析)
  1. 求解编排  : LP 优先求解 -> 自动检测"同时充放" -> 必要时升级 MILP 重解 (双模型自校验)
  2. 基线对照  : 无储能模型重解, 得到优越性对比的成本基准
  3. 侧面验证  : 合理性(功率平衡残差/SOC/套利形态) + 经济性(吞吐/弃光) + 优越性(节省额)
  4. 表格生成  : 题面表1(指定时段购电量)、表2(4h段充放电量)、144行完整策略
  5. 模板回填  : 复制附件5模板 result1.xlsx, 只按标签定位写数值, 模板格式零改动
"""
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from model import DT, E_MAX, E_MIN, build_model, solve

# ------------------------- 输出表结构常数 (题面给定) -------------------------
TOL = 1e-6          # 判定"同时充放"的数值容差: min(Pch,Pdis) > TOL 视为同时发生
# 表1 的 6 个指定时段。区间为左开右闭, "10:00-10:10" 即第 61 个时段(1-based),
# 对应 0-based 索引 h*6 (h 为整点小时数), 故购电量取 buy[h*6]
SLOT_LABELS = ["10:00-10:10", "12:00-12:10", "14:00-14:10",
               "16:00-16:10", "18:00-18:10", "20:00-20:10"]
SLOT_HOURS = [10, 12, 14, 16, 18, 20]
# 表2 的 6 个 4h 聚合段: 每段 = 24 个 10min 步, 用 0-based 切片 [a, b) 求和
BLOCK_LABELS = ["0:00-4:00", "4:00-8:00", "8:00-12:00",
                "12:00-16:00", "16:00-20:00", "20:00-24:00"]
BLOCK_RANGE = [(0, 24), (24, 48), (48, 72), (72, 96), (96, 120), (120, 144)]


def solve_storage(c, L, PV):
    """含储能模型的两阶段求解 (论文 5.2.2)

    第一阶段: LP 松弛求解 (不含互斥约束, 秒级)
    第二阶段: 检查 LP 解是否存在同一时段既充又放 (物理上无意义);
              若存在 -> 加 0-1 变量升级为 MILP 重解, 修正为互斥策略。
    本问实测同时充放 = 0, LP 解即原问题全局最优 (充放效率不对称使同时充放必然净亏损,
    求解器不会选择), 这一结论本身写入论文作为模型性质说明。
    """
    pick = lambda vs: [v.value() or 0.0 for v in vs]   # 提取变量值, None 视为 0
    prob, G, Pch, Pdis, E, Sp = build_model(c, L, PV, binary=False)   # ① LP
    status, obj = solve(prob)
    g, pch, pdis, e, sp = pick(G), pick(Pch), pick(Pdis), pick(E), pick(Sp)
    simult = [t for t in range(len(c)) if pch[t] > TOL and pdis[t] > TOL]  # ② 同时充放检测
    mode = "LP"
    if simult:
        # ③ 检测到同时充放 -> MILP 修正 (Big-M 互斥约束), 保证解的物理可行性
        mode = "MILP(修正同时充放)"
        prob, G, Pch, Pdis, E, Sp = build_model(c, L, PV, binary=True)
        status, obj = solve(prob)
        g, pch, pdis, e, sp = pick(G), pick(Pch), pick(Pdis), pick(E), pick(Sp)
    return dict(status=status, obj=obj, mode=mode, n_simult=len(simult),
                G=g, Pch=pch, Pdis=pdis, E=e, Sp=sp)


def solve_no_storage(c, L, PV):
    """无储能基线: 令 Pch = Pdis = 0 重解同一 LP, 得到"不装储能"的最低购电费。
    用途: 侧面验证中的优越性对比 (储能的经济价值 = 基线成本 - 最优成本)。"""
    prob, G, _, _, _, Sp = build_model(c, L, PV, storage=False)
    status, obj = solve(prob)
    return dict(status=status, obj=obj,
                G=[v.value() or 0.0 for v in G],
                Sp=[v.value() or 0.0 for v in Sp])


def validate(res, c):
    """侧面验证-合理性指标 (论文 5.2.4 ①②)

    输出 8 项可直接写入论文的检验指标:
      - 功率平衡残差(绝对/相对) : 能量守恒校验, 相对残差应在求解器容差内(~1e-8)
      - SOC 在界 / E0=E144      : 储能物理约束与跨日循环条件校验
      - 同时充放时段数 / 弃光量  : 解的物理意义与容量配置合理性校验
      - 低价段充电占比 / 高价段放电占比 : 将"低充高放"套利规律量化
        (以全天电价中位数为分界, 占比越接近 1 说明策略越符合经济调度规律)
      - 储能吞吐电量            : 全天放电总量, 衡量储能利用强度
    """
    g, pch, pdis, e, sp = res["G"], res["Pch"], res["Pdis"], res["E"], res["Sp"]
    L, PV = res.get("_L", []), res.get("_PV", [])      # main.py 注入的原始输入, 用于复算平衡
    # 逐时段复算 |G + PV + Pdis - L - Pch - Sp|, 取最大者为绝对残差
    errs = [abs(g[t] + PV[t] + pdis[t] - L[t] - pch[t] - sp[t]) for t in range(len(g))]
    bal = max(errs)
    scale = max(1e-9, max(L)) if L else 1.0            # 以最大负载为量纲基准, 得相对残差
    med = sorted(c)[len(c) // 2]                       # 全天电价中位数: 高/低价段分界
    ch_energy = sum(pch) * DT                          # 全天充电电量 (kWh)
    dis_energy = sum(pdis) * DT                        # 全天放电电量 (kWh)
    ch_low = sum(pch[t] for t in range(len(c)) if c[t] < med) * DT    # 低价段充电量
    dis_high = sum(pdis[t] for t in range(len(c)) if c[t] > med) * DT # 高价段放电量
    return {
        "功率平衡最大绝对残差kW": bal,
        "功率平衡相对残差": bal / scale,
        "SOC全程在界": bool(all(E_MIN - 1e-6 <= x <= E_MAX + 1e-6 for x in e)),
        "E0/E144": (e[0], e[-1]),
        "同时充放时段数": res["n_simult"],
        "弃光电量kWh": round(sum(sp) * DT, 4),
        "低价段充电占比": round(ch_low / ch_energy, 4) if ch_energy > 0 else None,
        "高价段放电占比": round(dis_high / dis_energy, 4) if dis_energy > 0 else None,
        "储能吞吐电量kWh": round(dis_energy, 2),
    }


def economy_check(res, no_res):
    """侧面验证-经济性与优越性指标 (论文 5.2.4 ③)
    对比"有储能最优解"与"无储能基线": 节省额 = 基线成本 - 最优成本, 节省比例 = 节省额/基线。
    注意: 储能并不减少总用电能量(往返损耗反而多购 ~19% 吞吐), 节省全部来自峰谷价差套利。"""
    save = no_res["obj"] - res["obj"]
    return {
        "有储能购电费(元)": round(res["obj"], 2),
        "无储能购电费(元)": round(no_res["obj"], 2),
        "节省购电费(元)": round(save, 2),
        "节省比例": round(save / no_res["obj"], 4) if no_res["obj"] else None,
        "有储能购电量kWh": round(sum(res["G"]) * DT, 2),
        "无储能购电量kWh": round(sum(no_res["G"]) * DT, 2),
    }


def make_results(res, c):
    """把 144 步的最优解组织为题面要求的三种输出形态

    返回
    ----
    slots  : 表1 六个指定时段的购电量 {标签: kWh}
    totals : 表1 汇总 {"buy": 全天购电量, "cost": 全天购电费}
    blocks : 表2 六个 4h 段的充/放电量 {标签: (充电kWh, 放电kWh)}
    e0/e24 : 表2 首末储电量 (0:00 与 24:00, kWh)
    full   : 完整策略明细表 (144 行 × 8 列, 供模板回填与论文附录)
    """
    g, pch, pdis, e = res["G"], res["Pch"], res["Pdis"], res["E"]
    buy = [x * DT for x in g]                          # 功率(kW) × Δt(h) = 电量(kWh)
    slots = {lab: buy[h * 6] for lab, h in zip(SLOT_LABELS, SLOT_HOURS)}
    totals = {"buy": sum(buy), "cost": res["obj"]}
    blocks = {lab: (sum(pch[a:b]) * DT, sum(pdis[a:b]) * DT)
              for lab, (a, b) in zip(BLOCK_LABELS, BLOCK_RANGE)}
    # 明细表的"时间段"列为区间标签, 如 t=0 -> "0:00-0:10", ..., t=143 -> "23:50-24:00"
    full = pd.DataFrame({
        "序号": range(1, 145),
        "时间段": [f"{t // 6}:{t % 6 * 10:02d}-{(t + 1) // 6}:{(t + 1) % 6 * 10:02d}"
                 for t in range(144)],
        "电价(元/kWh)": c,
        "购电功率kW": g, "购电量kWh": buy,
        "充电kW": pch, "放电kW": pdis, "时段末储电量kWh": e[1:],   # E_1..E_144 对齐各时段
    })
    return slots, totals, blocks, e[0], e[-1], full


def _norm(s):
    """标签归一化: 去除所有空白字符, 避免模板中"10:00 - 10:10"之类写法导致匹配失败"""
    return "".join(str(s).split())


def fill_result1(template: Path, out: Path, slots, totals, blocks, e0, e24, full):
    """按附件5模板原格式回填 result1.xlsx (格式零改动策略)

    实现要点 (答辩可述):
      1. 先整体复制模板文件, 再用 openpyxl 打开副本 -> sheet名/合并单元格/边框/列宽全部保留;
      2. 遍历所有字符串单元格, 命中表1/表2标签(如"10:00-10:10"、"0:00-4:00")时,
         向其右侧相邻单元格写入对应数值(4h段为右侧两格: 充电量/放电量);
      3. 单独识别"完整策略"明细表: 在前20行内找同时含"时间段"与"购电"字样的表头行,
         按表头关键字(放电优先于充电判断, 避免"充放电"误配)建立列映射, 逐行回填144行;
      4. 每次写入都记录"sheet!坐标<-标签"到审计清单, 控制台打印供人工核对, 杜绝写错格。
    """
    shutil.copyfile(template, out)                     # ① 复制模板, 保证原格式
    wb = load_workbook(out)
    report = []                                        # 回填审计清单
    slot_map = {_norm(k): v for k, v in slots.items()}
    block_map = {_norm(k): v for k, v in blocks.items()}
    for ws in wb.worksheets:
        # ② 标签定位回填: 表1 指定时段 / 表2 4h段 / 汇总与首末储电量
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue                           # 只处理文本单元格(标签)
                key = _norm(cell.value)
                if key in slot_map:
                    # 表1 时段标签 -> 右邻格写购电量
                    t = ws.cell(row=cell.row, column=cell.column + 1,
                                value=round(slot_map[key], 4))
                    report.append(f"{ws.title}!{t.coordinate}<-{key}")
                elif key in block_map:
                    # 表2 4h段标签 -> 右邻两格分别写充电量/放电量
                    ch, dis = block_map[key]
                    a = ws.cell(row=cell.row, column=cell.column + 1, value=round(ch, 4))
                    b = ws.cell(row=cell.row, column=cell.column + 2, value=round(dis, 4))
                    report.append(f"{ws.title}!{a.coordinate},{b.coordinate}<-{key}")
                elif key == _norm("全天购电量"):
                    t = ws.cell(row=cell.row, column=cell.column + 1,
                                value=round(totals["buy"], 4))
                    report.append(f"{ws.title}!{t.coordinate}<-全天购电量")
                elif key == _norm("全天购电费"):
                    t = ws.cell(row=cell.row, column=cell.column + 1,
                                value=round(totals["cost"], 4))
                    report.append(f"{ws.title}!{t.coordinate}<-全天购电费")
                elif key == _norm("0:00储电量"):
                    t = ws.cell(row=cell.row, column=cell.column + 1, value=round(e0, 4))
                    report.append(f"{ws.title}!{t.coordinate}<-0:00储电量")
                elif key == _norm("24:00储电量"):
                    t = ws.cell(row=cell.row, column=cell.column + 1, value=round(e24, 4))
                    report.append(f"{ws.title}!{t.coordinate}<-24:00储电量")
        # ③ 完整策略明细表回填: 先在前20行内定位表头行, 建立 列名关键字->列号 映射
        header_row, colmap = None, {}
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20)):
            texts = [_norm(c.value) if isinstance(c.value, str) else "" for c in row]
            if any(("时间段" in t or t == "时间") for t in texts) and any("购电" in t for t in texts):
                header_row = row[0].row
                for c, t in zip(row, texts):
                    # 判断顺序有讲究: "放电"必须先于"充电", 否则"充放电量"类表头会被误配
                    if "放电" in t:
                        colmap["pdis"] = c.column
                    elif "充电" in t:
                        colmap["pch"] = c.column
                    elif "购电量" in t:
                        colmap["buy"] = c.column
                    elif "购电功率" in t:
                        colmap["G"] = c.column
                    elif "储电量" in t:
                        colmap["E"] = c.column
                    elif "电价" in t:
                        colmap["c"] = c.column
                    elif "时间段" in t or t == "时间":
                        colmap["label"] = c.column
                break
        if header_row:
            # ④ 从表头下一行起逐行写 144 个时段 (模板预留空行则填充, 不足则顺延追加)
            for i, rec in enumerate(full.to_dict("records")):
                rr = header_row + 1 + i
                if "label" in colmap:
                    ws.cell(row=rr, column=colmap["label"], value=rec["时间段"])
                if "c" in colmap:
                    ws.cell(row=rr, column=colmap["c"], value=round(rec["电价(元/kWh)"], 6))
                if "G" in colmap:
                    ws.cell(row=rr, column=colmap["G"], value=round(rec["购电功率kW"], 4))
                if "buy" in colmap:
                    ws.cell(row=rr, column=colmap["buy"], value=round(rec["购电量kWh"], 4))
                if "pch" in colmap:
                    ws.cell(row=rr, column=colmap["pch"], value=round(rec["充电kW"], 4))
                if "pdis" in colmap:
                    ws.cell(row=rr, column=colmap["pdis"], value=round(rec["放电kW"], 4))
                if "E" in colmap:
                    ws.cell(row=rr, column=colmap["E"], value=round(rec["时段末储电量kWh"], 4))
            report.append(f"{ws.title}: 完整策略144行已回填")
    wb.save(out)
    return report


def write_verification(path, eco, chk):
    """把所有侧面验证指标写入独立的 result1_verification.xlsx
    (与 result1.xlsx 分离, 保证提交模板格式零改动; 本表仅供论文引用与答辩备查)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "验证对比"
    ws.append(["指标", "数值"])
    for k, v in list(eco.items()) + list(chk.items()):
        if isinstance(v, (list, tuple)):
            v = " / ".join(str(x) for x in v)          # openpyxl 不能写元组, 转成文本
        ws.append([k, v])
    wb.save(path)