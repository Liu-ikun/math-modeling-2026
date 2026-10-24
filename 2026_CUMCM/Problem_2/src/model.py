# -*- coding: utf-8 -*-
"""问题2 全年确定性LP模型 (含储能/无储能两套, 支持任意长度时间轴)
与问题1共用储能参数: 1200-10800 kWh, 5000 kW, 效率0.9/0.9, E0=6000
路径: 2026_CUMCM/Problem_2/src/model.py"""
import pulp

DT = 1.0 / 6.0                     # 10min 时段长 (h)
ETA_CH, ETA_DIS = 0.9, 0.9         # 充/放电效率
E_MIN, E_MAX, E0 = 1200.0, 10800.0, 6000.0   # SOC 下限/上限/初始 (kWh)
P_MAX = 5000.0                     # 充放电功率上限 (kW)
EMER_MULT = 5.0                    # 紧急购电 = 5 倍分时电价


def build_horizon(c_all, L_all, PV_all, storage=True, binary=False,
                  cyclic=False, include_emergency=True, e_init=E0):
    """在任意长度时间轴 N(=天数*144) 上一次性建模。
    c_all/L_all/PV_all: 长度 N 的平铺列表(按日顺序拼接)。
    storage=False -> 无储能基线; binary=True -> MILP(仅应急, 全年规模慎用);
    cyclic=True -> 末态回 E0(可选, 默认自由); e_init -> 首时段储电量。"""
    N = len(c_all)
    prob = pulp.LpProblem("P2_microgrid", pulp.LpMinimize)
    G = [pulp.LpVariable(f"G_{n}", lowBound=0) for n in range(N)]          # 计划购电
    Sp = [pulp.LpVariable(f"Sp_{n}", lowBound=0) for n in range(N)]        # 弃光(零成本)
    Eme = ([pulp.LpVariable(f"Eme_{n}", lowBound=0) for n in range(N)]
           if include_emergency else [0.0] * N)                            # 紧急购电(5倍价)
    if storage:
        Pch = [pulp.LpVariable(f"Pch_{n}", lowBound=0, upBound=P_MAX) for n in range(N)]
        Pdis = [pulp.LpVariable(f"Pdis_{n}", lowBound=0, upBound=P_MAX) for n in range(N)]
        E = [pulp.LpVariable(f"E_{n}", lowBound=E_MIN, upBound=E_MAX) for n in range(N + 1)]
    else:
        Pch, Pdis, E = [0.0] * N, [0.0] * N, None
    prob += pulp.lpSum(c_all[n] * (G[n] + EMER_MULT * (Eme[n] if include_emergency else 0.0))
                       * DT for n in range(N)), "obj_cost"
    for n in range(N):
        prob += (G[n] + (Eme[n] if include_emergency else 0.0) + PV_all[n]
                 + (Pdis[n] if storage else 0.0)
                 == L_all[n] + (Pch[n] if storage else 0.0) + Sp[n]), f"balance_{n}"
        if storage:
            prob += E[n + 1] == E[n] + ETA_CH * Pch[n] * DT - Pdis[n] * DT / ETA_DIS, f"soc_{n}"
            if binary:
                u = pulp.LpVariable(f"u_{n}", cat="Binary")
                prob += Pch[n] <= P_MAX * u, f"ch_{n}"
                prob += Pdis[n] <= P_MAX * (1 - u), f"dis_{n}"
    if storage:
        prob += E[0] == e_init, "soc_init"
        if cyclic:
            prob += E[N] == e_init, "soc_end"
    return prob, G, Eme, Pch, Pdis, E, Sp


def solve(prob, time_limit=600):
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=time_limit))
    return pulp.LpStatus[prob.status], pulp.value(prob.objective)