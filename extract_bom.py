# -*- coding: utf-8 -*-
"""只读提取：两张 BOM 报表的物料行、状态分布、零价行、同名相邻主/替对、两版差异。不修改原文件。"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import pandas as pd

QUOTE = r"D:\VX\xwechat_files\wxid_w9d35pgiofaa22_4d29\msg\file\2026-09\4.02.12.2.71Z2H.0384-报价.xlsx"
ORIG = r"D:\VX\xwechat_files\wxid_w9d35pgiofaa22_4d29\msg\file\2026-09\4.02.12.2.71Z2H.D.0384系统导出原表.xls"

HDR_ALIASES = {
    "序号": "seq", "子件编码": "code", "子件名称": "name", "子件规格": "spec",
    "状态": "status", "主料含税最新采购价(含损耗)": "p_loss", "主料含税最新采购价(无损耗)": "p_noloss",
    "单位": "unit", "用量/分母": "qty", "单位损耗": "loss_rate", "合计用量(含损耗)": "qty_total",
    "材料价(不含税本币)": "mat_price", "材料成本(不含税本币)(主料)": "mat_cost_main",
    "本位币(净价)": "base_net", "原币(净价)": "orig_net", "本位币(含税价)": "base_tax", "原币(含税)": "orig_tax",
    "人工成本": "labor", "制造成本": "mfg", "备注": "remark",
}


def norm_hdr(s):
    return str(s).replace(chr(10), "").replace(chr(13), "").replace("⏎", "").replace(" ", "").strip()


def load(path):
    if path.lower().endswith(".xls"):
        raw = pd.read_excel(path, sheet_name=None, engine="xlrd", header=None)
    else:
        raw = pd.read_excel(path, sheet_name=None, engine="openpyxl", header=None)
    df = list(raw.values())[0]
    meta = {}
    # 元数据：前 9 行扫关键字
    for i in range(min(9, len(df))):
        vals = [str(v).strip() for v in df.iloc[i].tolist() if str(v).strip() not in ("", "nan")]
        line = " | ".join(vals)
        for key in ("产品编号", "版本号", "单位", "成本合计(主替料)(最小)", "成本合计(主替料)(最大)",
                    "成本合计(主料)(不含税)", "成本合计(主料)(含税)最新采购价含损耗率", "成本合计(主料)(含税)最新采购价无损耗率"):
            if norm_hdr(line).startswith(norm_hdr(key)):
                parts = line.split("|")
                meta[key] = parts[1].strip() if len(parts) > 1 else ""
    # 表头行：包含“序号”且包含“子件编码”的行
    hdr_i = None
    for i in range(len(df)):
        cells = [norm_hdr(v) for v in df.iloc[i].tolist()]
        if "序号" in cells and "子件编码" in cells:
            hdr_i = i
            break
    colmap = {}
    for j, v in enumerate(df.iloc[hdr_i].tolist()):
        h = norm_hdr(v)
        if h in HDR_ALIASES and HDR_ALIASES[h] not in colmap.values():
            colmap[j] = HDR_ALIASES[h]
    rows = []
    for i in range(hdr_i + 1, len(df)):
        rec = {}
        seqv = df.iloc[i][0]
        try:
            float(seqv)
        except (TypeError, ValueError):
            continue
        for j, key in colmap.items():
            rec[key] = df.iloc[i][j]
        rows.append(rec)
    return meta, pd.DataFrame(rows)


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def analyze(tag, path):
    meta, df = load(path)
    print("=" * 90)
    print(f"### {tag}")
    print("元数据:", {k: v for k, v in meta.items()})
    print(f"物料行数: {len(df)}")
    sc = df["status"].astype(str).str.strip().value_counts(dropna=False).to_dict()
    print("状态分布:", sc)
    zero = []
    for _, r in df.iterrows():
        st = str(r.get("status", "")).strip()
        p = fnum(r.get("p_loss"))
        if st == "主料" and (p is None or p == 0):
            zero.append(f"#{r['seq']} {r['name']}({r['code']}) 价={p}")
    print(f"主料零价/缺价行({len(zero)}): " + "; ".join(zero) if zero else "主料零价/缺价行: 无")
    # 同名相邻 主料/替代料 对
    pairs = []
    rows = df.to_dict("records")
    for a, b in zip(rows, rows[1:]):
        if str(a["name"]).strip() == str(b["name"]).strip():
            sa, sb = str(a["status"]).strip(), str(b["status"]).strip()
            if {sa, sb} == {"主料", "替代料"}:
                pairs.append(f"#{a['seq']}/{b['seq']} {a['name']} [{a['code']}({sa},{fnum(a.get('p_loss'))}) vs {b['code']}({sb},{fnum(b.get('p_loss'))})]")
    print(f"同名相邻主/替对({len(pairs)}):")
    for p in pairs:
        print("   ", p)
    # 替代料行其主料价列是否为0
    alt_zero = sum(1 for r in rows if str(r["status"]).strip() == "替代料" and (fnum(r.get("p_loss")) in (0, None)))
    print(f"替代料行主料价=0 的数量: {alt_zero}")
    nan_rows = [f"#{r['seq']} {r['name']} 编码{r['code']}" for _, r in df.iterrows() if str(r.get('status', '')).strip() in ('', 'nan', 'None')]
    print(f'状态为空的行({len(nan_rows)}): ' + '; '.join(nan_rows))
    # 备注非空样例
    rem = [f"#{r['seq']} {r['name']}:{r['remark']}" for r in rows if str(r.get("remark", "")).strip() not in ("", "nan")]
    print(f"备注非空({len(rem)}): " + "; ".join(rem[:8]))
    return df


dq = analyze("报价表 4.02.12.2.71Z2H.0384-报价.xlsx", QUOTE)
do = analyze("系统原表 4.02.12.2.71Z2H.D.0384系统导出原表.xls", ORIG)

# 两版差异（按子件名称+状态）
nq = set(zip(dq["name"].astype(str).str.strip(), dq["status"].astype(str).str.strip()))
no = set(zip(do["name"].astype(str).str.strip(), do["status"].astype(str).str.strip()))
print("=" * 90)
print("仅在报价表中的 (名称,状态):", sorted(nq - no, key=str))
print("仅在系统原表中的 (名称,状态):", sorted(no - nq, key=str))
# 编码交叉
cq, co = set(dq["code"].astype(str).str.strip()), set(do["code"].astype(str).str.strip())
print(f"编码交集 {len(cq & co)} / 报价独有 {len(cq - co)} / 原表独有 {len(co - cq)}")
print("报价独有编码样例:", sorted(cq - co)[:6])
print("原表独有编码样例:", sorted(co - cq)[:6])
