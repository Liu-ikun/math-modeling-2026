# -*- coding: utf-8 -*-
#校验清洗后数据
"""核验所有 Problem_x/data/processed 输出: 形状、每列非空数、示例值"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]

for prob in sorted(ROOT.glob("Problem_*")):
    pdir = prob / "data" / "processed"
    print(f"\n===== {prob.name} =====")
    if not pdir.exists():
        print("  processed 目录不存在")
        continue
    for f in sorted(pdir.iterdir()):
        if f.suffix.lower() == ".xlsx":
            df = pd.read_excel(f)
        elif f.suffix.lower() == ".csv":
            df = pd.read_csv(f)
        else:
            continue
        print(f"\n[{f.name}] shape={df.shape}")
        for c in df.columns[:12]:
            nn = int(df[c].notna().sum())
            sample = df[c].dropna().head(2).tolist()
            print(f"  {c}: dtype={df[c].dtype} | 非空 {nn}/{len(df)} | 示例 {sample}")
        if df.shape[1] > 12:
            rest = [c for c in df.columns[12:]]
            empty_rest = [c for c in rest if df[c].notna().sum() == 0]
            print(f"  ... 其余 {len(rest)} 列; 其中全空列: {empty_rest if empty_rest else '无'}")