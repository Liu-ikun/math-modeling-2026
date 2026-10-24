# -*- coding: utf-8 -*-
"""问题3 阶段LP: MPC滚动优化的剩余时域子模型
mode='plan': 0:00阶段, G即计划(=假定执行), 无调整项;
mode='adj' : 更新阶段, G为调整购电, 含0.5c|G-plan|偏差罚(等价于1.5c/-0.5c结算)"""
import pulp
import numpy as np
import scipy.sparse as sp
from scipy.optimize import linprog

DT = 1.0 / 6.0
ETA_CH, ETA_DIS = 0.9, 0.9
E_MIN, E_MAX, E0 = 1200.0, 10800.0, 6000.0
P_MAX = 5000.0
EMER_MULT = 5.0
ADJ_PEN = 0.5


def build_stage(c, L, PVfc, plan=None, e_init=E0, mode="plan"):
    T = len(c)
    prob = pulp.LpProblem(f"stage_{mode}", pulp.LpMinimize)
    G = [pulp.LpVariable(f"G_{t}", lowBound=0) for t in range(T)]
    Eme = [pulp.LpVariable(f"Eme_{t}", lowBound=0) for t in range(T)]
    Sp = [pulp.LpVariable(f"Sp_{t}", lowBound=0) for t in range(T)]
    Pch = [pulp.LpVariable(f"Pch_{t}", 0, P_MAX) for t in range(T)]
    Pdis = [pulp.LpVariable(f"Pdis_{t}", 0, P_MAX) for t in range(T)]
    E = [pulp.LpVariable(f"E_{t}", E_MIN, E_MAX) for t in range(T + 1)]
    obj = pulp.lpSum(c[t] * G[t] * DT for t in range(T))
    if mode == "adj":
        Dp = [pulp.LpVariable(f"Dp_{t}", lowBound=0) for t in range(T)]
        Dm = [pulp.LpVariable(f"Dm_{t}", lowBound=0) for t in range(T)]
        obj += pulp.lpSum(ADJ_PEN * c[t] * (Dp[t] + Dm[t]) * DT for t in range(T))
        for t in range(T):
            prob += Dp[t] >= G[t] - plan[t]
            prob += Dm[t] >= plan[t] - G[t]
    obj += pulp.lpSum(EMER_MULT * c[t] * Eme[t] * DT for t in range(T))
    prob += obj
    for t in range(T):
        prob += G[t] + Eme[t] + PVfc[t] + Pdis[t] == L[t] + Pch[t] + Sp[t], f"bal_{t}"
        prob += E[t + 1] == E[t] + ETA_CH * Pch[t] * DT - Pdis[t] * DT / ETA_DIS, f"soc_{t}"
    prob += E[0] == e_init
    return prob, dict(G=G, Eme=Eme, Sp=Sp, Pch=Pch, Pdis=Pdis, E=E)


def solve(prob, time_limit=120):
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=time_limit, options=["feasTol 1e-9"]))
    return pulp.LpStatus[prob.status], pulp.value(prob.objective)


_STAGE_CACHE = {}


def _stage_matrices(T, mode):
    """按(时域长度,模式)缓存阶段LP的稀疏约束矩阵/界/变量索引:
    约束结构只与(T,mode)有关, 与数据无关 -> 建模一次, 复用全年"""
    key = (T, mode)
    if key in _STAGE_CACHE:
        return _STAGE_CACHE[key]
    if mode == "adj":
        n = 8 * T + 1
        iG, iDp, iDm = 0, T, 2 * T
        iEme, iSp, iPch, iPdis, iE = 3 * T, 4 * T, 5 * T, 6 * T, 7 * T
    else:
        n = 6 * T + 1
        iG, iEme, iSp, iPch, iPdis, iE = 0, T, 2 * T, 3 * T, 4 * T, 5 * T
    rows, cols, vals = [], [], []
    for t in range(T):
        rows += [t, t, t, t, t]
        cols += [iG + t, iEme + t, iPdis + t, iPch + t, iSp + t]
        vals += [1.0, 1.0, 1.0, -1.0, -1.0]
        r = T + t
        rows += [r, r, r, r]
        cols += [iE + t + 1, iE + t, iPch + t, iPdis + t]
        vals += [1.0, -1.0, -ETA_CH * DT, DT / ETA_DIS]
    rows.append(2 * T)
    cols.append(iE)
    vals.append(1.0)
    A = sp.csr_matrix((vals, (rows, cols)), shape=(2 * T + 1, n))
    A_ub = None
    if mode == "adj":
        ur, uc, uv = [], [], []
        for t in range(T):
            ur += [t, t]
            uc += [iG + t, iDp + t]
            uv += [1.0, -1.0]          # G - Dp <= plan  <=>  Dp >= G - plan
            r2 = T + t
            ur += [r2, r2]
            uc += [iG + t, iDm + t]
            uv += [-1.0, -1.0]         # -G - Dm <= -plan <=> Dm >= plan - G
        A_ub = sp.csr_matrix((uv, (ur, uc)), shape=(2 * T, n))
    bounds = [(0, None)] * n
    for t in range(T):
        bounds[iPch + t] = (0, P_MAX)
        bounds[iPdis + t] = (0, P_MAX)
    for t in range(T + 1):
        bounds[iE + t] = (E_MIN, E_MAX)
    idx = dict(G=iG, Eme=iEme, Pch=iPch, Pdis=iPdis, E=iE)
    if mode == "adj":
        idx.update(Dp=iDp, Dm=iDm)
    _STAGE_CACHE[key] = (A, A_ub, bounds, idx)
    return _STAGE_CACHE[key]


def solve_stage(c, L, PVfc, plan=None, e_init=E0, mode="plan"):
    """HiGHS直解阶段LP: 与build_stage数学等价(同变量/约束/目标), 仅跳过PuLP建模;
    每次调用只组装目标向量与rhs(平衡右端=L-PVfc, SOC右端=0, E0=e_init)"""
    T = len(c)
    A, A_ub, bounds, idx = _stage_matrices(T, mode)
    cvec = np.zeros(A.shape[1])
    carr = np.asarray(c, dtype=float)
    cvec[idx["G"]:idx["G"] + T] = carr * DT
    cvec[idx["Eme"]:idx["Eme"] + T] = EMER_MULT * carr * DT
    if mode == "adj":
        cvec[idx["Dp"]:idx["Dp"] + T] = ADJ_PEN * carr * DT
        cvec[idx["Dm"]:idx["Dm"] + T] = ADJ_PEN * carr * DT
    b_ub = None
    if mode == "adj":
        parr = np.asarray(plan, dtype=float)
        b_ub = np.concatenate([parr, -parr])
    rhs = np.empty(2 * T + 1)
    rhs[:T] = np.asarray(L, dtype=float) - np.asarray(PVfc, dtype=float)
    rhs[T:2 * T] = 0.0
    rhs[2 * T] = e_init
    res = linprog(cvec, A_ub=A_ub, b_ub=b_ub, A_eq=A, b_eq=rhs,
                  bounds=bounds, method="highs-ds")
    if not res.success:
        raise SystemExit(f"阶段LP求解失败: {res.message}")
    x = res.x
    return dict(G=x[idx["G"]:idx["G"] + T].tolist(),
                Pch=x[idx["Pch"]:idx["Pch"] + T].tolist(),
                Pdis=x[idx["Pdis"]:idx["Pdis"] + T].tolist(),
                E=x[idx["E"]:idx["E"] + T + 1].tolist(),
                obj=float(res.fun))
