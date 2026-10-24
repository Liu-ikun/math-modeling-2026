# -*- coding: utf-8 -*-
"""
问题1 LP/MILP 模型: 含储能/无储能两套, 弃光变量保证可行性
路径: D:/Python projects/Math Modeling Project 2026/2026_CUMCM/Problem_1/src/model.py

【模型定位】(对应论文 5.2.1 模型建立)
  微网单日购电策略优化: 一天划分为 144 个 10min 时段 (t = 1..144),
  在满足功率平衡与储能物理约束的前提下, 最小化全天购电费。

【决策变量】
  G_t    : 时段 t 的计划购电功率 (kW),            G_t >= 0
  Pch_t  : 时段 t 的储能充电功率 (kW),            0 <= Pch_t  <= 5000
  Pdis_t : 时段 t 的储能放电功率 (kW),            0 <= Pdis_t <= 5000
  E_t    : 时段 t 末的储电量 (kWh),               1200 <= E_t <= 10800, t = 0..144
  Sp_t   : 时段 t 的弃光功率 (kW, 零成本辅助变量), 仅在光伏盈余无法消纳时启用,
           作用是保证模型在任何数据工况下均可行 (本问数据下最优解 Sp ≡ 0)
  u_t    : 0-1 变量 (仅 MILP 模式), 强制充/放电互斥

【模型参数】(全部来自题面, 不含任何假设值)
  c_t   : 分时电价 (元/kWh)  —— 附件1
  L_t   : 小区负载 (kW)      —— 附件1
  PV_t  : 光伏预测出力 (kW)  —— 附件1
"""
import pulp

# ------------------------- 模型常数 (题面给定) -------------------------
DT = 1.0 / 6.0                    # 时段长度: 10min = 1/6 h, 用于功率(kW)->电量(kWh)换算
ETA_CH, ETA_DIS = 0.9, 0.9        # 储能充电/放电效率 (题面: 均为 90%)
E_MIN, E_MAX, E0 = 1200.0, 10800.0, 6000.0   # 储电量下限/上限/初始值 (kWh); 题面要求 E_0 = E_144 = 6000
P_MAX = 5000.0                    # 充/放电功率上限 (kW)


def build_model(c, L, PV, binary=False, storage=True):
    """构建问题1的线性规划模型 (论文式(1)-(5))

    参数
    ----
    c, L, PV : 长度144的列表 (元/kWh, kW, kW), 行序 = 时段 t = 1..144
               (输入取自 附件1_10min对齐.xlsx, 其第 t 行标签恰为区间右端点 t*10 分钟)
    binary   : True 时加入 0-1 变量 u_t 强制充放互斥 (LP 解出现"同时充放"时的修正模式)
    storage  : False 时退化为"无储能基线"模型 (Pch = Pdis = 0),
               用于侧面验证中的优越性对比 (有储能 vs 无储能成本差)

    返回
    ----
    (prob, G, Pch, Pdis, E, Sp) : PuLP 问题对象与四组变量列表 (E 长度 145, 含 E_0..E_144)
    """
    T = len(c)                                          # 时段数 = 144
    prob = pulp.LpProblem("P1_microgrid", pulp.LpMinimize)

    # ---- 决策变量 ----
    G = [pulp.LpVariable(f"G_{t}", lowBound=0) for t in range(T)]          # 购电功率, 非负
    Sp = [pulp.LpVariable(f"Sp_{t}", lowBound=0) for t in range(T)]        # 弃光功率, 非负, 目标中不计费
    if storage:
        # 储能三组变量: 充/放功率受 5000 kW 上限, 储电量受 [1200, 10800] 区间约束
        Pch = [pulp.LpVariable(f"Pch_{t}", lowBound=0, upBound=P_MAX) for t in range(T)]
        Pdis = [pulp.LpVariable(f"Pdis_{t}", lowBound=0, upBound=P_MAX) for t in range(T)]
        E = [pulp.LpVariable(f"E_{t}", lowBound=E_MIN, upBound=E_MAX) for t in range(T + 1)]
    else:
        # 无储能基线: 充放功率恒为 0 (常数, 不建变量), 无 SOC 状态
        Pch, Pdis, E = [0.0] * T, [0.0] * T, None

    # ---- 目标函数 (论文式(1)): min 全天购电费 = Σ c_t * G_t * Δt ----
    prob += pulp.lpSum(c[t] * G[t] * DT for t in range(T)), "obj_cost"

    # ---- 约束条件 ----
    for t in range(T):
        # 式(2) 功率平衡: 购电 + 光伏 + 放电 = 负载 + 充电 + 弃光
        # (取等式: 供电不低于负载; 多余电力要么充电、要么弃光, 多购即浪费, 最优解处必然取等)
        prob += (G[t] + PV[t] + (Pdis[t] if storage else 0.0)
                 == L[t] + (Pch[t] if storage else 0.0) + Sp[t]), f"balance_{t}"
        if storage:
            # 式(3) SOC 递推: E_t = E_{t-1} + 0.9*Pch_t*Δt - Pdis_t*Δt/0.9
            # (充电乘效率、放电除效率, 体现往返损耗 ~19%)
            prob += E[t + 1] == E[t] + ETA_CH * Pch[t] * DT - Pdis[t] * DT / ETA_DIS, f"soc_{t}"
            if binary:
                # 式(4) 充放互斥 (Big-M): u_t=1 只许充电, u_t=0 只许放电
                # 仅在 LP 松弛解出现同时充放时启用; 本问数据下 LP 解天然互斥, 未触发
                u = pulp.LpVariable(f"u_{t}", cat="Binary")
                prob += Pch[t] <= P_MAX * u, f"ch_bigM_{t}"
                prob += Pdis[t] <= P_MAX * (1 - u), f"dis_bigM_{t}"
    if storage:
        # 式(5) 首末储电量相等 (题面: E_0 = E_144 = 6000 kWh, 保证策略跨日可循环执行)
        prob += E[0] == E0, "soc_init"
        prob += E[T] == E0, "soc_end"
    return prob, G, Pch, Pdis, E, Sp


def solve(prob, time_limit=60):
    """调用 CBC 求解器 (PuLP 内置, 开源可复现)

    msg=0        : 静默模式, 不打印求解器日志
    timeLimit=60 : 超时保护; 本问题规模仅 576 个连续变量, 实测 < 1 s 求得最优
    返回 (求解状态字符串, 目标函数值)
    """
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=time_limit))
    return pulp.LpStatus[prob.status], pulp.value(prob.objective)