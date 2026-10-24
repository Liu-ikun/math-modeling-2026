# -*- coding: utf-8 -*-
"""dump 原始附件的每个工作表: 形状、前8行原样、每列非空数与取值样本
用法: python dump_raw.py 1   (参数=附件编号, 默认1)"""
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
idx = int(sys.argv[1]) if len(sys.argv) > 1 else 1
path = ROOT / "Problem_1" / "data" / "raw" / f"附件{idx}.xlsx"
xl = pd.ExcelFile(path)
print(f"文件: {path}")
print("工作表:", xl.sheet_names)
for sn in xl.sheet_names:
    raw = pd.read_excel(path, sheet_name=sn, header=None)
    print(f"\n===== sheet [{sn}] shape={raw.shape} =====")
    print(raw.head(8).to_string())
    for j in range(min(raw.shape[1], 10)):
        col = raw[j]
        uniq = col.dropna().unique()[:5].tolist()
        print(f"  col{j}: 非空={col.notna().sum()}/{len(col)} dtype={col.dtype} 取值样本={uniq}")