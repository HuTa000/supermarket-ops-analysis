# -*- coding: utf-8 -*-
"""
工具 · 把旧版 .xls（Excel 2003 二进制）转成 .xlsx，并打印结构摘要。

用途：表格编辑器 / openpyxl 读不了 .xls，先用它转一次。
换数据时也可以先跑这个看看新文件的表名和列名，再决定 config.json 怎么写。

用法：
    python convert_xls.py                          # 自动取 data/ 下的 .xls
    python convert_xls.py --input D:/某个报表.xls
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import _paths as P

ap = P.make_parser("把 .xls 转成 .xlsx 并打印结构摘要")
ap.add_argument("--output", "-o", default="", help="输出 .xlsx 路径，默认写到 data/ 下同名文件")
args = ap.parse_args()

SRC = P.find_input(args.input)
if not SRC.lower().endswith(".xls"):
    print(f"[提示] 输入不是 .xls：{SRC}\n       仍然可以解析并打印结构。")
DST = args.output or os.path.join(P.DATA, os.path.splitext(os.path.basename(SRC))[0] + ".xlsx")
if not os.path.isabs(DST):
    DST = os.path.join(P.ROOT, DST)
os.makedirs(os.path.dirname(DST), exist_ok=True)

eng = "xlrd" if SRC.lower().endswith(".xls") else "openpyxl"
xls = pd.ExcelFile(SRC, engine=eng)
print(f"源文件：{SRC}")
print("SHEETS:", xls.sheet_names)
print("=" * 70)

frames = {}
for name in xls.sheet_names:
    frames[name] = xls.parse(name, header=0)

with pd.ExcelWriter(DST, engine="openpyxl") as writer:
    for name, df in frames.items():
        df.to_excel(writer, sheet_name=name[:31], index=False)
        print(f"\n### sheet: {name}  shape={df.shape}")
        print("columns:", list(df.columns))
        print("dtypes:\n", df.dtypes.to_string())
        print("head(3):\n", df.head(3).to_string())
        print("nulls:", df.isna().sum().to_dict())
        print("-" * 70)

print("\nWROTE:", DST)

# ---------------------------------------------------------------
# 结构对照：告诉用户 config.json 该怎么填
# ---------------------------------------------------------------
cfg = P.load_config()
print("\n" + "=" * 70)
print("【结构对照】把你的表/字段对上脚本要求")
print("=" * 70)
for std, actual in cfg["表名"].items():
    hit = actual in frames or any(actual in s or s in actual for s in frames)
    print(f"  表「{std}」 ← 源文件「{actual}」  {'✅ 命中' if hit else '❌ 未命中，请改 config.json 的「表名」'}")

miss = []
for std, need_cols in P.REQUIRED.items():
    src_sheet = cfg["表名"].get(std)
    if src_sheet not in frames:
        miss += [(std, c) for c in need_cols]
        continue
    cols = list(frames[src_sheet].columns)
    for c in need_cols:
        if c not in cols:
            miss.append((std, c))
if miss:
    print("\n  以下必需字段在源文件中找不到，需在 config.json 的「字段映射」里指认：")
    for s, c in miss:
        print(f"    · 表「{s}」缺「{c}」")
else:
    print("\n  ✅ 全部必需字段齐备，config.json 无需改动。")
print("=" * 70)
