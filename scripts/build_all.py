# -*- coding: utf-8 -*-
"""
一键复现 · 超市运营数据分析项目
从原始 .xls 一路跑到三份交付物 + 全部校验通过。

用法：
    python build_all.py
"""
import sys, io, os, subprocess, time, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import _paths as P

PY = sys.executable

ap = argparse.ArgumentParser(description="一键复现：从原始数据跑到三份交付物 + 全部校验")
P.add_input_arg(ap)
ARGS = ap.parse_args()
NAME = P.prefix()

STEPS = [
    ("etl.py",             "W1 数据预处理：三表整合 / 退货去重 / 派生字段 / 对账"),
    ("analysis.py",        "W2-W3 六大业务模块分析 → metrics.json"),
    ("build_dashboard.py", "W4 生成单文件 HTML 交互看板"),
    ("build_excel.py",     "W5-A 生成 Excel 多 Sheet 汇总表"),
    ("crosscheck.py",      "W5-C 交付物交叉核对"),
]

print("#" * 74)
print("#  超市运营数据分析项目 · 一键复现")
print("#" * 74)
src = P.find_input(ARGS.input)
print(f"#  输入：{src}")
print("#" * 74)

failed = []
t0 = time.time()
for script, desc in STEPS:
    print(f"\n{'=' * 74}\n▶ {desc}\n{'=' * 74}")
    cmd = [PY, os.path.join(HERE, script)]
    if script == "etl.py" and ARGS.input:
        cmd += ["--input", ARGS.input]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (r.stdout or "") + (r.stderr or "")
    tail = "\n".join(out.strip().splitlines()[-6:])
    print(tail)
    if r.returncode != 0:
        failed.append(script)
        print(f"  ❌ {script} 失败（退出码 {r.returncode}）")

print(f"\n{'#' * 74}")
if failed:
    print(f"#  复现中断：失败 {len(failed)} 步 → {', '.join(failed)}")
    print("#" * 74)
    sys.exit(1)

print(f"#  ✅ 全部 {len(STEPS)} 步通过，耗时 {time.time()-t0:.1f} 秒")
print("#" * 74)
print("\n交付物：")
for f in [f"{NAME}_交互看板.html", f"{NAME}_汇总表.xlsx", f"{NAME}报告.md"]:
    p = os.path.join(ROOT, f)
    print(f"  · {f}  ({os.path.getsize(p)/1024:.0f} KB)" if os.path.exists(p) else f"  · {f}  缺失")
print("\n补充验证（需 Node）：node scripts/verify_dashboard.cjs  → 看板真实浏览器断言")
