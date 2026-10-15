# -*- coding: utf-8 -*-
"""
W4 · 生成单文件 HTML 交互看板
把 ECharts 与行级数据内联进模板 → 输出可离线打开的单文件看板。
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _paths as P

ROOT = P.ROOT
TPL = os.path.join(ROOT, "templates", "dashboard.html")
ECHARTS = os.path.join(ROOT, "vendor", "echarts.min.js")
DATA = os.path.join(ROOT, "outputs", "dashboard_data.json")
OUT = os.path.join(ROOT, f"{P.prefix()}_交互看板.html")

tpl = open(TPL, encoding="utf-8").read()
echarts = open(ECHARTS, encoding="utf-8").read()
data = open(DATA, encoding="utf-8").read()

assert "__ECHARTS__" in tpl and "__DATA__" in tpl, "模板占位符缺失"

html = tpl.replace("__ECHARTS__", echarts).replace("__DATA__", data)
open(OUT, "w", encoding="utf-8").write(html)

size = os.path.getsize(OUT)
print("=" * 72)
print("W4 · 单文件看板生成完成")
print("=" * 72)
print(f"  输出：{OUT}")
print(f"  体积：{size/1024/1024:.2f} MB（内联 ECharts {len(echarts)/1024/1024:.2f} MB + 数据 {len(data)/1024/1024:.2f} MB）")

# 自检：占位符必须已全部替换，且关键结构存在
assert "__ECHARTS__" not in html and "__DATA__" not in html, "占位符未替换干净"
for key in ['id="c1"', 'id="c2"', 'id="c3"', 'id="c4"', 'id="c5"',
            'id="c6"', 'id="c7"', 'id="c8"', 'id="losstbl"', 'id="kpis"', 'id="filters"']:
    assert key in html, f"缺少结构：{key}"
print("  ✅ 自检：占位符已替换，8 个图表容器 + KPI + 筛选器 + 亏损表结构齐备")
