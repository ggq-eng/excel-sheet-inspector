# -*- coding: utf-8 -*-
"""只读检查两个 Excel 文件的结构：工作表、行列规模、原始网格、逐列统计。不修改原文件。"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import pandas as pd

FILES = [
    r"D:\VX\xwechat_files\wxid_w9d35pgiofaa22_4d29\msg\file\2026-09\4.02.12.2.71Z2H.0384-报价.xlsx",
    r"D:\VX\xwechat_files\wxid_w9d35pgiofaa22_4d29\msg\file\2026-09\4.02.12.2.71Z2H.D.0384系统导出原表.xls",
]


def cell_str(v):
    if v is None:
        return ""
    try:
        if isinstance(v, float) and pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v)[:26].replace("\n", "⏎")


def sniff(path):
    with open(path, "rb") as f:
        return f.read(8)


def dump_sheet(df, name):
    print(f"\n--- SHEET [{name}]: {df.shape[0]} 行 x {df.shape[1]} 列 ---")
    n = len(df)
    if n <= 60:
        show_idx = list(range(n))
    else:
        show_idx = list(range(15)) + list(range(n - 3, n))
    prev = None
    for i in show_idx:
        if prev is not None and i != prev + 1:
            print(f"[r{prev+2}..r{i}] ……（中间行省略）")
        row = df.iloc[i]
        print(f"[r{i+1}] " + " | ".join(cell_str(v) for v in row.tolist()))
        prev = i
    stats = []
    for c in df.columns:
        col = df[c]
        nn = int(col.notna().sum())
        nu = int(col.nunique())
        first = col.dropna().iloc[0] if nn else ""
        stats.append(f"c{c}:非空{nn} 唯一{nu} 样例[{cell_str(first)}]")
    print("COLS: " + " || ".join(stats))


for path in FILES:
    print("=" * 100)
    print("FILE:", path.split("\\")[-1])
    head = sniff(path)
    print("format sniff:", head[:8].hex())
    try:
        if path.lower().endswith(".xls"):
            if head.startswith(b"\xd0\xcf\x11\xe0"):
                sheets = pd.read_excel(path, sheet_name=None, engine="xlrd", header=None)
            else:
                with open(path, "rb") as f:
                    raw = f.read()
                tables = pd.read_html(io.BytesIO(raw), header=None)
                sheets = {f"html_table_{i}": t for i, t in enumerate(tables)}
        else:
            sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None)
        for name, df in sheets.items():
            dump_sheet(df, str(name))
    except Exception:
        import traceback
        traceback.print_exc()
