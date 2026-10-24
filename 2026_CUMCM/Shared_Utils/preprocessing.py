# -*- coding: utf-8 -*-
r"""
2026 CUMCM 数据清洗脚本 (完整版)
文件路径: 2026_CUMCM/Shared_Utils/preprocessing.py

清洗流程:
  读取附件1/2/3/4 -> 统一时间轴 -> 附件2/4宽转长
  -> 附件3日期ffill、长宽转换、10min线性插值 -> 附件1生成典型日10min表
  -> 按问题合并输入 -> 保存到各Problem_x/data/processed/ -> 生成cleaning_report.md

关键约定:
  1. 时间列兼容: datetime / datetime.time / date / Excel序列号 / 日内小数 /
     纯时刻字符串("0:10","00:10:00") / 次日标记时刻("0:00+1"=次日00:00)。
     纯时刻统一挂 BASE_DATE, 避免被 pandas 解析成"运行当天"或时区偏移。
  2. 附件2/4 为"行=日期, 列=时刻"宽表: 列名即时刻(可带+N次日标记),
     时间轴 = 日期 + 列名时刻; 数据本身10min对齐, 不做插值。
  3. 附件3 为6h间隔预报, 需10min线性插值, 输出宽表(避免百万行长表)。
  4. 保存: 行/列/单元格数超 Excel 上限自动改存 CSV(utf-8-sig, Excel可直接打开);
     时间列统一写为 "YYYY-MM-DD HH:MM" 字符串, 保证 Excel 中可见。
  5. 每次运行先清理 processed 目录旧产出, 避免残留文件。

运行 (PowerShell, 项目根目录):
  .venv\Scripts\python.exe 2026_CUMCM\Shared_Utils\preprocessing.py --preview
  .venv\Scripts\python.exe 2026_CUMCM\Shared_Utils\preprocessing.py
"""

import argparse
import re
import sys
import time
from datetime import datetime, time as _time, date as _date
from pathlib import Path

import numpy as np
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = Path(__file__).resolve().parent / "cleaning_report.md"
LOG_PATH = Path(__file__).resolve().parent / "preprocessing.log"
TIME_KEYWORDS = ("时间", "时刻", "日期", "date", "time", "timestamp")
BASE_DATE = "2000-01-01"
TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\+(\d+))?$")
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384
EXCEL_CELL_LIMIT = 5_000_000
OLD_PATTERNS = ("附件1_*", "附件2_*", "附件3_*", "附件4_*", "附件_清洗后.*")

CONFIG = {
    1: {"wide": False, "ffill_date": False, "typical_day": True,
        "time_col": None, "sheet": None},
    2: {"wide": True, "ffill_date": False, "typical_day": False,
        "time_col": None, "sheet": None, "resample": False, "series_name": "负荷"},
    3: {"wide": True, "ffill_date": True, "typical_day": False,
        "time_col": None, "date_col": None, "sheet": None, "resample": True},
    4: {"wide": True, "ffill_date": False, "typical_day": False,
        "time_col": None, "sheet": None, "resample": False, "series_name": "电价"},
}


def log(msg: str):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def read_excel_smart(path: Path, sheet=None) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet or 0, header=None)
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        raise ValueError(f"{path.name} 读取后为空, 请检查文件")
    best = max(range(min(10, len(raw))), key=lambda i: raw.iloc[i].notna().sum())
    df = raw.iloc[best + 1:].reset_index(drop=True)
    header = raw.iloc[best].astype(str)
    header = (header.str.replace("\n", "", regex=False)
              .str.strip()
              .str.replace("Unnamed.*", "nan", regex=True))
    header = header.replace({"nan": pd.NA, "": pd.NA}).ffill()
    header = header.fillna(pd.Series([f"列{i}" for i in range(len(header))]))
    df.columns = _dedupe_columns(header.tolist())
    return df.dropna(how="all").reset_index(drop=True)


def _dedupe_columns(names):
    seen, out = {}, []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n}_{seen[n]}")
    return out


def _parse_time_str(s: str):
    """解析 'H:MM[:SS][+N]' 时刻字符串, 返回 (时,分,秒,次日偏移) 或 None"""
    m = TIME_RE.match(str(s).strip())
    if not m:
        return None
    hh, mm, ss, off = m.groups()
    return int(hh), int(mm), int(ss or 0), int(off or 0)


def _parse_one(x):
    """逐元素解析时间为无时区 Timestamp, 失败返回 NaT"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return pd.NaT
    if isinstance(x, _time):
        return (pd.Timestamp(BASE_DATE)
                + pd.Timedelta(hours=x.hour, minutes=x.minute, seconds=x.second))
    if isinstance(x, _date) and not isinstance(x, datetime):
        return pd.Timestamp(x.isoformat())
    if isinstance(x, datetime):
        ts = pd.Timestamp(x)
        return ts.tz_localize(None) if ts.tzinfo is not None else ts
    if isinstance(x, (int, float, np.integer, np.floating)):
        v = float(x)
        if pd.isna(v):
            return pd.NaT
        if v >= 25569:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=v)
        if 0 <= v <= 1:
            return pd.Timestamp(BASE_DATE) + pd.Timedelta(days=v)
        return pd.NaT
    tod = _parse_time_str(x)
    if tod is not None:
        hh, mm, ss, off = tod
        return (pd.Timestamp(BASE_DATE)
                + pd.Timedelta(days=off, hours=hh, minutes=mm, seconds=ss))
    try:
        ts = pd.Timestamp(str(x).strip())
    except (ValueError, TypeError, OverflowError):
        return pd.NaT
    return ts.tz_localize(None) if ts.tzinfo is not None else ts


def to_datetime_series(s: pd.Series) -> pd.Series:
    s = pd.Series(s)
    if pd.api.types.is_datetime64_any_dtype(s):
        out = pd.to_datetime(s, errors="coerce")
        if getattr(out.dt, "tz", None) is not None:
            out = out.dt.tz_localize(None)
        return out
    if pd.api.types.is_numeric_dtype(s):
        v = s.dropna()
        if not v.empty:
            if float(v.median()) >= 25569:
                return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
            if 0 <= float(v.min()) and float(v.max()) <= 1:
                return pd.Series(pd.Timestamp(BASE_DATE) + pd.to_timedelta(s, unit="D"),
                                 index=s.index)
    return pd.to_datetime(s.map(_parse_one), errors="coerce")


def _coerce_numeric(s: pd.Series) -> pd.Series:
    """文本数值强制转换: 去千分位/空白; 仍失败则提取字符串中首个数字(兼容带单位文本)"""
    if pd.api.types.is_numeric_dtype(s):
        return s
    t = s.astype(str).str.replace(",", "", regex=False).str.strip()
    out = pd.to_numeric(t, errors="coerce")
    if out.notna().mean() < 0.5:
        alt = pd.to_numeric(t.str.extract(r"(-?\d+(?:\.\d+)?)", expand=False),
                            errors="coerce")
        if alt.notna().mean() > out.notna().mean():
            out = alt
    return out


def detect_time_column(df: pd.DataFrame, hint=None, exclude=None):
    if hint and hint in df.columns:
        return hint
    for col in df.columns:
        if exclude and col == exclude:
            continue
        if any(k in str(col).lower() for k in TIME_KEYWORDS):
            return col
    best_col, best_ratio = None, 0.0
    for col in df.columns:
        if exclude and col == exclude:
            continue
        sample = df[col].dropna().head(500)
        if sample.empty:
            continue
        ratio = to_datetime_series(sample).notna().mean()
        if ratio > best_ratio:
            best_col, best_ratio = col, ratio
    return best_col if best_ratio >= 0.8 else None


def resample_10min(df: pd.DataFrame, time_col: str):
    t = to_datetime_series(df[time_col])
    df = df[t.notna()].copy()
    df = df.infer_objects()
    for c in df.columns:
        if c != time_col and df[c].dtype == object:
            v = pd.to_numeric(df[c], errors="coerce")
            if v.notna().mean() >= 0.9:
                df[c] = v
    ts = pd.to_datetime(t[t.notna()])
    wrap = ts.diff() < pd.Timedelta(0)
    df["_t"] = ts + pd.to_timedelta(wrap.cumsum(), unit="D")
    df = df.sort_values("_t").drop_duplicates(subset="_t", keep="last")
    for c in [c for c in df.columns if c not in ("_t", time_col)]:
        sample = df[c].dropna().head(200)
        if sample.empty:
            continue
        ratio = sample.astype(str).str.strip().map(
            lambda s: _parse_time_str(s) is not None).mean()
        if ratio >= 0.8:
            tods = df[c].astype(str).str.strip().map(_parse_time_str)
            delta = pd.to_timedelta(
                [pd.Timedelta(days=p[3], hours=p[0], minutes=p[1], seconds=p[2])
                 if p is not None else pd.NaT for p in tods])
            df["_t"] = df["_t"].dt.normalize() + delta
            df = df.drop(columns=[c]).dropna(subset=["_t"])
            break
    numeric_cols = []
    for c in df.columns:
        if c in ("_t", time_col):
            continue
        coerced = _coerce_numeric(df[c])
        if pd.api.types.is_numeric_dtype(coerced) and coerced.notna().any():
            df[c] = coerced
            numeric_cols.append(c)
    log(f"resample: time_col={time_col} numeric_cols={numeric_cols}")
    if df.empty:
        dropped = [c for c in df.columns if c not in ("_t", time_col)]
        return pd.DataFrame(columns=["时间"]), dropped
    grid = pd.date_range(df["_t"].min(), df["_t"].max(), freq="10min")
    if len(grid) > 1_000_000:
        raise ValueError(f"10min网格过大({len(grid)}点, 跨度{grid[-1] - grid[0]}), "
                         f"时间列解析异常, 请检查 {time_col} 列格式")
    out = pd.DataFrame({"_t": grid})
    for c in numeric_cols:
        s = df.set_index("_t")[c]
        s = s[~s.index.duplicated(keep="last")].reindex(grid)
        out[c] = s.interpolate(method="linear", limit_direction="both").to_numpy()
    out = out.dropna(axis=1, how="all")
    dropped = [c for c in df.columns if c not in numeric_cols and c not in ("_t", time_col)]
    log(f"resample: grid={len(grid)} dropped={dropped}")
    return out.rename(columns={"_t": "时间"}), dropped


def wide_to_long(df: pd.DataFrame, cfg: dict, time_col: str) -> pd.DataFrame:
    """宽转长; '行=日期, 列=时刻'结构自动合成 日期+列名时刻(含+N次日) 时间轴"""
    df = df.copy()
    date_col = cfg.get("date_col")
    if cfg.get("ffill_date") and date_col and date_col in df.columns:
        df[date_col] = df[date_col].ffill()
    t_full = to_datetime_series(df[time_col])
    if date_col and date_col in df.columns:
        combined = to_datetime_series(
            df[date_col].astype(str) + " " + df[time_col].astype(str))
        if combined.notna().mean() >= 0.8:
            t_full = combined
    drop_cols = {time_col}
    if date_col:
        drop_cols.add(date_col)
    value_cols = [c for c in df.columns if c not in drop_cols]
    for c in value_cols:
        df[c] = _coerce_numeric(df[c])
    names = [str(c).strip() for c in value_cols]
    parsed_names = [_parse_time_str(n) for n in names]
    ok = sum(p is not None for p in parsed_names)
    if names and ok >= 0.8 * len(names):
        base_date = pd.to_datetime(t_full).dt.normalize()
        series_name = cfg.get("series_name") or f"附件{cfg.get('_idx', '')}序列"
        frames = []
        for c, n, p in zip(value_cols, names, parsed_names):
            if p is None:
                continue
            hh, mm, ss, off = p
            t = base_date + pd.Timedelta(days=off, hours=hh, minutes=mm, seconds=ss)
            frames.append(pd.DataFrame({"时间": t, "数值": df[c].to_numpy()}))
        long = pd.concat(frames, ignore_index=True)
        long = long[long["时间"].notna()].dropna(subset=["数值"])
        long["变量"] = series_name
        return long.sort_values("时间").reset_index(drop=True)[["时间", "变量", "数值"]]
    long = df.melt(value_vars=value_cols, var_name="变量", value_name="数值")
    long["时间"] = np.tile(t_full.to_numpy(), len(value_cols))
    long = long[long["时间"].notna()].dropna(subset=["数值"])
    return long[["时间", "变量", "数值"]].reset_index(drop=True)


def long_to_10min(long: pd.DataFrame) -> pd.DataFrame:
    if long.empty:
        return long
    grid = pd.date_range(long["时间"].min(), long["时间"].max(), freq="10min")
    parts = []
    for var, g in long.groupby("变量"):
        s = g.set_index("时间")["数值"]
        s = s[~s.index.duplicated(keep="last")]
        g_grid = pd.date_range(max(s.index.min(), grid[0]),
                               min(s.index.max(), grid[-1]), freq="10min")
        interp = s.reindex(g_grid).interpolate(method="linear", limit_direction="both")
        tmp = interp.rename("数值").reset_index().rename(columns={"index": "时间"})
        tmp["变量"] = var
        parts.append(tmp[["时间", "变量", "数值"]])
    return pd.concat(parts, ignore_index=True)


def pivot_wide(long: pd.DataFrame, prefix: str = ""):
    if long is None or long.empty:
        return None
    long = long.copy()
    long["变量"] = long["变量"].astype(str).replace({"时间": "数值_时间"})
    w = long.pivot_table(index="时间", columns="变量", values="数值", aggfunc="mean")
    if prefix:
        w.columns = [f"{prefix}_{c}" for c in w.columns]
    w = w.reset_index()
    w = w.loc[:, ~w.columns.duplicated()]
    if "时间" not in w.columns:
        return None
    return w


def typical_day_10min(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.infer_objects()
    for c in df.columns:
        if c != "时间" and df[c].dtype == object:
            v = pd.to_numeric(df[c], errors="coerce")
            if v.notna().mean() >= 0.9:
                df[c] = v
    df["_t"] = to_datetime_series(df["时间"])
    df = df[df["_t"].notna()]
    df["hour"] = df["_t"].dt.hour
    df["minute"] = df["_t"].dt.minute
    numeric_cols = [c for c in df.columns if c not in ("时间", "_t", "hour", "minute")
                    and pd.api.types.is_numeric_dtype(df[c])]
    g = df.groupby(["hour", "minute"], as_index=False)[numeric_cols].mean()
    g = g.sort_values(["hour", "minute"]).reset_index(drop=True)
    g["hour"] = g["hour"].round().astype(int)
    g["minute"] = g["minute"].round().astype(int)
    g["时刻"] = (g["hour"].astype(str).str.zfill(2) + ":"
                 + g["minute"].astype(str).str.zfill(2))
    return g[["时刻"] + numeric_cols]


def merge_by_problem(part1, wides: dict):
    frames = []
    if part1 is not None and not part1.empty:
        frames.append(part1.copy())
    for idx in sorted(wides):
        if wides[idx] is not None:
            frames.append(wides[idx].copy())
    if not frames:
        return None
    for f in frames:
        f["时间"] = pd.to_datetime(f["时间"])
        dup = f.columns.duplicated()
        if dup.any():
            f = f.loc[:, ~dup]
    merged = frames[0]
    for f in frames[1:]:
        merged = merged.merge(f, on="时间", how="outer")
    merged = merged.loc[:, ~merged.columns.duplicated()]
    p1_cols = [c for c in part1.columns if c != "时间"] if part1 is not None else []
    if p1_cols and len(merged):
        t1 = pd.to_datetime(part1["时间"])
        tm = pd.to_datetime(merged["时间"])
        if (t1.max() - t1.min()) <= pd.Timedelta("2D") and \
                (tm.max() - tm.min()) > pd.Timedelta("2D"):
            hhmm = t1.dt.strftime("%H:%M")
            for c in p1_cols:
                mapping = dict(zip(hhmm, part1[c]))
                merged[c] = [mapping.get(t.strftime("%H:%M")) for t in tm]
            other = [c for c in merged.columns if c not in p1_cols and c != "时间"]
            if other:
                merged = merged.dropna(subset=other, how="all")
    return merged.sort_values("时间").reset_index(drop=True)


def _format_for_save(df: pd.DataFrame) -> pd.DataFrame:
    """datetime 列统一写成 'YYYY-MM-DD HH:MM' 字符串, 保证 Excel 中时间可见"""
    df = df.copy()
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = df[c].dt.strftime("%Y-%m-%d %H:%M")
    return df


def save_processed(prob_dir: Path, outputs: dict) -> tuple:
    processed = prob_dir / "data" / "processed"
    processed.mkdir(parents=True, exist_ok=True)
    for pat in OLD_PATTERNS:
        for old in processed.glob(pat):
            old.unlink(missing_ok=True)
    saved = []
    for name, df in outputs.items():
        if df is None or df.empty:
            continue
        try:
            df = _format_for_save(df)
            cells = len(df) * df.shape[1]
            if (len(df) <= EXCEL_MAX_ROWS and df.shape[1] <= EXCEL_MAX_COLS
                    and cells <= EXCEL_CELL_LIMIT):
                df.to_excel(processed / f"{name}.xlsx", index=False)
                saved.append(f"{name}.xlsx ({len(df)}行x{df.shape[1]}列)")
            else:
                df.to_csv(processed / f"{name}.csv", index=False, encoding="utf-8-sig")
                saved.append(f"{name}.csv ({len(df)}行x{df.shape[1]}列, 超Excel上限改存CSV)")
        except Exception as e:
            saved.append(f"{name} 保存失败: {type(e).__name__}: {e}")
            log(f"保存失败 {name}: {type(e).__name__}: {e}")
    return processed, saved


def process_problem(prob_dir: Path, lines: list):
    raw_dir = prob_dir / "data" / "raw"
    lines.append(f"\n## {prob_dir.name}")
    parts_long = {}
    part1 = None
    for idx in (1, 2, 3, 4):
        path = raw_dir / f"附件{idx}.xlsx"
        cfg = dict(CONFIG[idx])
        lines.append(f"\n### 附件{idx}")
        if not path.exists():
            lines.append(f"- ⚠ 文件不存在: {path}")
            continue
        try:
            df = read_excel_smart(path, cfg.get("sheet"))
            if cfg.get("ffill_date"):
                cfg["date_col"] = (cfg.get("date_col")
                                   or next((c for c in df.columns if "日期" in str(c)), None))
            time_col = detect_time_column(df, cfg.get("time_col"), exclude=cfg.get("date_col"))
            lines.append(f"- 原始形状: {df.shape} | 识别时间列: {time_col}")
            if time_col is None:
                lines.append("- ⚠ 未识别到时间列, 跳过该附件 (可在 CONFIG 手动指定 time_col)")
                continue
            before_na = int(df.isna().sum().sum())
            if cfg.get("wide"):
                cfg["_idx"] = idx
                long = wide_to_long(df, cfg, time_col)
                if cfg.get("resample"):
                    long = long_to_10min(long)
                    lines.append(f"- 宽转长+10min线性插值: {len(long)} 行 | "
                                 f"缺失值 {before_na} -> {int(long['数值'].isna().sum())}")
                else:
                    lines.append(f"- 宽转长: {len(long)} 行 (时间轴已10min对齐, 不插值) | "
                                 f"缺失值 {before_na} -> {int(long['数值'].isna().sum())}")
                parts_long[idx] = long
            else:
                part1, dropped1 = resample_10min(df, time_col)
                if dropped1:
                    lines.append(f"- ⚠ 以下列因无可解析数值被剔除: {dropped1}")
                t = part1["时间"]
                span = f"{t.min():%Y-%m-%d %H:%M} ~ {t.max():%Y-%m-%d %H:%M}" if len(t) else "空"
                lines.append(f"- 10min对齐后: {part1.shape} | 时间范围: {span} | "
                             f"缺失值 {before_na} -> {int(part1.isna().sum().sum())}")
        except Exception as e:
            lines.append(f"- ❌ 该附件清洗失败: {type(e).__name__}: {e}")
            log(f"{prob_dir.name}/附件{idx} 清洗失败: {type(e).__name__}: {e}")

    wides = {}
    for idx, long in parts_long.items():
        try:
            wides[idx] = pivot_wide(long, f"附件{idx}")
        except Exception as e:
            lines.append(f"- ❌ 附件{idx} 透视宽表失败: {type(e).__name__}: {e}")
            log(f"{prob_dir.name}/附件{idx} 透视失败: {type(e).__name__}: {e}")
    typical = None
    if part1 is not None:
        try:
            typical = typical_day_10min(part1)
            lines.append(f"\n- 附件1 典型日10min表: {typical.shape} (期望 144 行)")
        except Exception as e:
            lines.append(f"- ❌ 典型日表生成失败: {type(e).__name__}: {e}")
    merged = None
    try:
        merged = merge_by_problem(part1, wides)
        if merged is not None:
            lines.append(f"- 按问题合并后: {merged.shape}")
    except Exception as e:
        lines.append(f"- ❌ 按问题合并失败: {type(e).__name__}: {e}")
        log(f"{prob_dir.name} 合并失败: {type(e).__name__}: {e}")
    wide3 = None
    if 3 in parts_long:
        try:
            wide3 = pivot_wide(parts_long[3], "")
        except Exception as e:
            lines.append(f"- ❌ 附件3 宽表产出失败: {type(e).__name__}: {e}")

    outputs = {
        "附件1_10min对齐": part1,
        "附件1_典型日_10min": typical,
        "附件2_长表": parts_long.get(2),
        "附件4_长表": parts_long.get(4),
        "附件3_10min宽表": wide3,
        "附件_清洗后": merged,
    }
    try:
        processed_dir, saved = save_processed(prob_dir, outputs)
        lines.append(f"- 已保存到: {processed_dir}")
        lines.append("- 输出文件: " + "; ".join(saved))
    except Exception as e:
        lines.append(f"- ❌ 保存失败: {type(e).__name__}: {e}")
        log(f"{prob_dir.name} 保存失败: {type(e).__name__}: {e}")


def generate_report(lines: list, problems: list):
    header = [
        "# 数据清洗报告 cleaning_report.md",
        "",
        f"- 生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"- 清洗脚本: `2026_CUMCM/Shared_Utils/preprocessing.py`",
        f"- 处理问题: {', '.join(p.name for p in problems)}",
        "- 清洗流程: 读取附件1/2/3/4 -> 统一时间轴 -> 附件2/4宽转长 -> "
        "附件3日期ffill、长宽转换、10min线性插值 -> 附件1典型日10min表 -> "
        "按问题合并 -> 保存processed -> 生成本报告",
        "",
    ]
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(header + lines) + "\n")
    log(f"报告已生成: {REPORT_PATH}")


def preview_all(problems: list):
    for prob in problems:
        print(f"\n===== {prob.name} =====")
        raw_dir = prob / "data" / "raw"
        for idx in (1, 2, 3, 4):
            path = raw_dir / f"附件{idx}.xlsx"
            if not path.exists():
                print(f"附件{idx}: 不存在")
                continue
            try:
                df = read_excel_smart(path, CONFIG[idx].get("sheet"))
                print(f"\n[附件{idx}] {path.name}  shape={df.shape}")
                print("列名:", list(df.columns))
                print(df.head(3).to_string())
            except Exception as e:
                print(f"\n[附件{idx}] 读取失败: {type(e).__name__}: {e}")


def main():
    parser = argparse.ArgumentParser(description="2026 CUMCM 数据清洗")
    parser.add_argument("--problems", nargs="*", default=None,
                        help="指定问题目录, 默认全部 Problem_*")
    parser.add_argument("--preview", action="store_true",
                        help="只预览各附件表结构, 不执行清洗")
    args = parser.parse_args()

    if args.problems:
        problems = [ROOT / p for p in args.problems]
    else:
        problems = sorted(ROOT.glob("Problem_*"))
    problems = [p for p in problems if p.is_dir()]
    if not problems:
        print("未找到任何 Problem_* 目录")
        return
    print(f"项目根目录: {ROOT}")
    print(f"处理问题: {[p.name for p in problems]}")

    if args.preview:
        preview_all(problems)
        return

    lines = []
    t0 = time.time()
    for prob in problems:
        log(f"开始清洗 {prob.name}")
        try:
            process_problem(prob, lines)
        except Exception as e:
            lines.append(f"\n## {prob.name}\n- ❌ 该问题清洗中断: {type(e).__name__}: {e}")
            log(f"{prob.name} 中断: {type(e).__name__}: {e}")
    generate_report(lines, problems)
    log(f"全部完成, 耗时 {time.time() - t0:.2f} 秒")


if __name__ == "__main__":
    main()
