# -*- coding: utf-8 -*-
"""问题4 模型层
(A) Q4-2: 全年一次性全局LP(完美信息), scipy-HiGHS稀疏直解 + PuLP/CBC对照;
(B) Q4-3: MPC阶段LP(与问题3数学等价, 仅电价由逐日重复改为逐日波动)。
储能物理参数与问题1/2/3完全一致: 1200-10800 kWh, 5000 kW, 0.9/0.9, E0=6000。
路径: 2026_CUMCM/Problem_4/src/model.py"""
import numpy as np
import pulp
import scipy.sparse as sp
from scipy.optimize import linprog

DT = 1.0 / 6.0                      # 10min 时段长 (h)
ETA_CH, ETA_DIS = 0.9, 0.9
E_MIN, E_MAX, E0 = 1200.0, 10800.0, 6000.0
P_MAX = 5000.0
EMER_MULT = 5.0                     # 紧急购电 = 5 倍当时电价
ADJ_PEN = 0.5                       # 调整偏差 |ΔG| 的净惩罚系数(1.5c/-0.5c 等价形式)
TOL_M = 1e-6

_GLOBAL_CACHE = {}
_STAGE_CACHE = {}


# ============================================================ Q4-2 全局LP
def _global_matrices(N, storage):
    """全年(或任意长度)时间轴上的稀疏约束矩阵。
    变量序: [G(N), Eme(N), Sp(N), Pch(N), Pdis(N), E(N+1)] -> n = 6N+1
    约束序: [功率平衡(N), SOC递推(N), 初始SOC(1)]        -> m = 2N+1
    结构只与(N, storage)有关, 与电价/负荷数据无关 => 建模一次, 全年及各对照链复用"""
    key = (N, storage)
    if key in _GLOBAL_CACHE:
        return _GLOBAL_CACHE[key]
    t = np.arange(N)
    one = np.ones(N)
    iG, iEme, iSp = 0, N, 2 * N
    if storage:
        iPch, iPdis, iE = 3 * N, 4 * N, 5 * N
        n, m = 6 * N + 1, 2 * N + 1
        rows = np.concatenate([t, t, t, t, t,
                               N + t, N + t, N + t, N + t,
                               np.array([2 * N])])
        cols = np.concatenate([iG + t, iEme + t, iPdis + t, iPch + t, iSp + t,
                               iE + t + 1, iE + t, iPch + t, iPdis + t,
                               np.array([iE])])
        vals = np.concatenate([one, one, one, -one, -one,
                               one, -one, np.full(N, -ETA_CH * DT),
                               np.full(N, DT / ETA_DIS), np.array([1.0])])
        bnds = np.zeros((n, 2))
        bnds[:, 1] = np.inf
        bnds[iPch:iPch + N, 1] = P_MAX
        bnds[iPdis:iPdis + N, 1] = P_MAX
        bnds[iE:iE + N + 1, 0] = E_MIN
        bnds[iE:iE + N + 1, 1] = E_MAX
        idx = dict(G=iG, Eme=iEme, Sp=iSp, Pch=iPch, Pdis=iPdis, E=iE)
    else:
        n, m = 3 * N, N
        rows = np.concatenate([t, t, t])
        cols = np.concatenate([iG + t, iEme + t, iSp + t])
        vals = np.concatenate([one, one, -one])
        bnds = np.zeros((n, 2))
        bnds[:, 1] = np.inf
        idx = dict(G=iG, Eme=iEme, Sp=iSp, Pch=-1, Pdis=-1, E=-1)
    A = sp.csr_matrix((vals, (rows, cols)), shape=(m, n))
    _GLOBAL_CACHE[key] = (A, bnds, idx)
    return _GLOBAL_CACHE[key]


def solve_global(c_all, L_all, PV_all, storage=True, e_init=E0, method="highs-ds"):
    """HiGHS直解全局LP: min Σ c_t(G_t + 5·Eme_t)Δt
    s.t. G+Eme+PV+Pdis = L+Pch+Sp ; E_{t+1}=E_t+0.9·Pch·Δt-Pdis·Δt/0.9 ; E_0=e_init
    c_all 为逐时段电价(附件4展平 或 附件1平铺), 长度必须等于 L_all/PV_all"""
    N = len(c_all)
    if not (len(L_all) == N and len(PV_all) == N):
        raise SystemExit(f"全局LP输入长度不一致: c={N} L={len(L_all)} PV={len(PV_all)}")
    A, bnds, idx = _global_matrices(N, storage)
    ca = np.asarray(c_all, dtype=float)
    cvec = np.zeros(A.shape[1])
    cvec[idx["G"]:idx["G"] + N] = ca * DT
    cvec[idx["Eme"]:idx["Eme"] + N] = EMER_MULT * ca * DT
    rhs = np.empty(A.shape[0])
    rhs[:N] = np.asarray(L_all, dtype=float) - np.asarray(PV_all, dtype=float)
    if storage:
        rhs[N:2 * N] = 0.0
        rhs[2 * N] = e_init
    res = linprog(cvec, A_eq=A, b_eq=rhs, bounds=bnds, method=method)
    if not res.success:
        raise SystemExit(f"全局LP求解失败: {res.message}")
    x = res.x
    g = x[idx["G"]:idx["G"] + N]
    eme = x[idx["Eme"]:idx["Eme"] + N]
    spv = x[idx["Sp"]:idx["Sp"] + N]
    if storage:
        pch = x[idx["Pch"]:idx["Pch"] + N]
        pdis = x[idx["Pdis"]:idx["Pdis"] + N]
        E = x[idx["E"]:idx["E"] + N + 1].tolist()
        n_sim = int(((pch > TOL_M) & (pdis > TOL_M)).sum())
    else:
        pch = pdis = np.zeros(N)
        E, n_sim = None, 0
    return dict(status="Optimal", obj=float(res.fun), N=N, D=N // 144,
                G=g.tolist(), Eme=eme.tolist(), Sp=spv.tolist(),
                Pch=pch.tolist(), Pdis=pdis.tolist(), E=E, n_simult=n_sim,
                buy=float(g.sum() * DT), emer=float(eme.sum() * DT),
                spill=float(spv.sum() * DT))


def build_horizon(c_all, L_all, PV_all, storage=True, e_init=E0):
    """PuLP/CBC版全局LP(与solve_global数学等价), 仅用于双求解器对照(建议<=7天窗口)"""
    N = len(c_all)
    prob = pulp.LpProblem("P4_global", pulp.LpMinimize)
    G = [pulp.LpVariable(f"G_{n}", lowBound=0) for n in range(N)]
    Sp = [pulp.LpVariable(f"Sp_{n}", lowBound=0) for n in range(N)]
    Eme = [pulp.LpVariable(f"Eme_{n}", lowBound=0) for n in range(N)]
    if storage:
        Pch = [pulp.LpVariable(f"Pch_{n}", 0, P_MAX) for n in range(N)]
        Pdis = [pulp.LpVariable(f"Pdis_{n}", 0, P_MAX) for n in range(N)]
        E = [pulp.LpVariable(f"E_{n}", E_MIN, E_MAX) for n in range(N + 1)]
    else:
        Pch, Pdis, E = [0.0] * N, [0.0] * N, None
    prob += pulp.lpSum(c_all[n] * (G[n] + EMER_MULT * Eme[n]) * DT for n in range(N))
    for n in range(N):
        prob += (G[n] + Eme[n] + PV_all[n] + (Pdis[n] if storage else 0.0)
                 == L_all[n] + (Pch[n] if storage else 0.0) + Sp[n]), f"bal_{n}"
        if storage:
            prob += (E[n + 1] == E[n] + ETA_CH * Pch[n] * DT
                     - Pdis[n] * DT / ETA_DIS), f"soc_{n}"
    if storage:
        prob += E[0] == e_init, "soc_init"
    return prob, G, Eme, Pch, Pdis, E, Sp


def solve(prob, time_limit=600):
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=time_limit,
                                 options=["feasTol 1e-9"]))
    return pulp.LpStatus[prob.status], pulp.value(prob.objective)


# ============================================================ Q4-3 阶段LP
def build_stage(c, L, PVfc, plan=None, e_init=E0, mode="plan"):
    """PuLP版阶段LP(仅用于--check双求解器隔离对照)
    mode='plan': 0:00阶段, G即计划; mode='adj': 更新阶段, G为调整购电+0.5c|G-plan|罚"""
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


def _stage_matrices(T, mode):
    """阶段LP稀疏矩阵缓存(结构与数据无关); adj模式的 Dp/Dm 两条不等式必须在A_ub中,
    漏写会使 Dp≡Dm≡0 -> 调整'免费' -> 成本倒挂(问题3已踩过并修复)"""
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
    A_ub, b_ub = None, None
    if mode == "adj":
        ur, uc, uv = [], [], []
        for t in range(T):
            ur += [t, t]
            uc += [iG + t, iDp + t]
            uv += [1.0, -1.0]          # G - Dp <= plan  <=> Dp >= G - plan
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
    """HiGHS直解阶段LP(与build_stage数学等价), 每次只组装目标向量与rhs"""
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