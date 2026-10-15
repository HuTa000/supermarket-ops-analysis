# -*- coding: utf-8 -*-
"""
W5-C · 交付物交叉核对
1) 核对 Excel 汇总表中的关键数值 == metrics.json == 原始数据重算
2) 核对报告文本中引用的关键数字与 metrics.json 一致
"""
import sys, io, os, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd, openpyxl
import _paths as P

ROOT = P.ROOT
OUT = P.OUT
NAME = P.prefix()
M = json.load(open(os.path.join(OUT, "metrics.json"), encoding="utf-8"))
df = pd.read_csv(os.path.join(OUT, "clean_orders.csv"), encoding="utf-8-sig")
report = open(os.path.join(ROOT, f"{NAME}报告.md"), encoding="utf-8").read()

ok = bad = 0


def chk(name, cond, detail=""):
    global ok, bad
    if cond:
        ok += 1; print(f"  ✅ {name}")
    else:
        bad += 1; print(f"  ❌ {name}  {detail}")


print("=" * 74)
print("W5-C · 交付物交叉核对")
print("=" * 74)

# ---------- 1. 从原始宽表独立重算，与 metrics.json 比对 ----------
print("\n【1】metrics.json ←→ 原始数据独立重算")
chk("销售额合计", abs(df.销售额.sum() - M["baseline"]["销售额"]) < 0.5,
    f'{df.销售额.sum():,.2f} vs {M["baseline"]["销售额"]:,}')
chk("利润合计", abs(df.利润.sum() - M["baseline"]["利润"]) < 0.5,
    f'{df.利润.sum():,.2f} vs {M["baseline"]["利润"]:,}')
chk("订单数", df["订单 ID"].nunique() == M["baseline"]["订单数"])
chk("客户数", df["客户 ID"].nunique() == M["baseline"]["客户数"])
chk("明细行", len(df) == M["baseline"]["明细行"])

# 折扣档
for x in M["M2_discount"]:
    dv = x["折扣值"]
    sub = df[df.折扣 == dv]
    chk(f'折扣 {x["折扣"]} 利润', abs(sub.利润.sum() - x["利润"]) < 1,
        f'{sub.利润.sum():,.0f} vs {x["利润"]:,}')
    chk(f'折扣 {x["折扣"]} 销售额占比', abs(sub.销售额.sum()/df.销售额.sum()*100 - x["销售额占比"]) < 0.02)

# ≥40%
high = df[df.折扣 >= 0.4]
chk("≥40% 销售额占比", abs(high.销售额.sum()/df.销售额.sum()*100 - M["M2_high_discount"]["销售额占比"]) < 0.02)
chk("≥40% 明细行占比", abs(len(high)/len(df)*100 - M["M2_high_discount"]["明细行占比"]) < 0.02,
    f'重算 {len(high)/len(df)*100:.2f}% vs 记录 {M["M2_high_discount"]["明细行占比"]}')
chk("≥40% 亏损占全网亏损比",
    abs(high.利润.sum()/df.loc[df.利润 < 0, "利润"].sum()*100 - M["M2_high_discount"]["亏损占全网亏损比"]) < 0.02)

# 类目
for x in M["M3_category"]:
    sub = df[df.类别 == x["类别"]]
    chk(f'{x["类别"]} 利润率', abs(sub.利润.sum()/sub.销售额.sum()*100 - x["利润率"]) < 0.02)
# 子类
for name in ("桌子", "美术"):
    sx = next(y for y in M["M3_subcategory"] if y["子类别"] == name)
    sub = df[df.子类别 == name]
    chk(f'{name} 利润', abs(sub.利润.sum() - sx["利润"]) < 1)

# 大区
for x in M["M5_region"]:
    sub = df[df.地区 == x["地区"]]
    chk(f'{x["地区"]} 利润率', abs(sub.利润.sum()/sub.销售额.sum()*100 - x["利润率"]) < 0.02)
    chk(f'{x["地区"]} 销售额占比', abs(sub.销售额.sum()/df.销售额.sum()*100 - x["销售额占比"]) < 0.02)

# 客户
cust = df.groupby("客户 ID").利润.sum()
chk("净亏损客户数", (cust < 0).sum() == M["M4_customer_stats"]["亏损客户数"])
chk("净亏损客户合计亏损", abs(cust[cust < 0].sum() - M["M4_customer_stats"]["亏损客户合计亏损"]) < 1)

# 退货
ol = df.drop_duplicates(subset=["订单 ID"])
chk("真实退货率", abs(ol.是否退货.mean()*100 - M["M6_return"]["真实退货率"]) < 0.01)

# ---------- 2. Excel ←→ metrics.json ----------
print("\n【2】Excel 汇总表 ←→ metrics.json")
wb = openpyxl.load_workbook(os.path.join(ROOT, f"{NAME}_汇总表.xlsx"))
chk("Sheet 数量 = 9", len(wb.sheetnames) == 9, str(wb.sheetnames))

ws = wb["关键指标总览"]
kv = {}
for row in ws.iter_rows(min_row=5, values_only=True):
    if row[0] and row[1] is not None:
        kv[row[0]] = row[1]
chk("Excel 销售额 = metrics", kv.get("销售额（元）") == M["baseline"]["销售额"],
    f'{kv.get("销售额（元）")} vs {M["baseline"]["销售额"]}')
chk("Excel 利润 = metrics", kv.get("利润（元）") == M["baseline"]["利润"])
chk("Excel 退货率 = metrics", abs(kv.get("真实退货率", 0) - M["M6_return"]["真实退货率"]) < 0.01)

ws = wb["M2-折扣与利润"]
d2 = [r for r in ws.iter_rows(min_row=5, values_only=True)
      if r[0] and re.fullmatch(r"\d+%", str(r[0]))]
chk("Excel 折扣行数 = 7", len(d2) == 7, str([r[0] for r in d2]))
chk("Excel 折扣40% 利润 = metrics",
    any(abs((r[4] or 0) - next(x for x in M["M2_discount"] if x["折扣"] == "40%")["利润"]) < 1 for r in d2))

# ---------- 3. 报告正文 ←→ metrics.json ----------
print("\n【3】报告正文数字反查（关键引用）")
def w(v):  # 元 → 万元字符串（1 位小数）
    return f"{v/10000:.1f}"

def g(d, k, v):
    return next(x for x in M[d] if x[k] == v)


claims = {
    "2019 销售额 546.2 万": "546.2 万",
    "整体利润率 13.36%": "13.36%",
    "≥40% 销售额占比 18.96%": f'{M["M2_high_discount"]["销售额占比"]}%',
    "≥40% 亏损占比 90.86%": f'{M["M2_high_discount"]["亏损占全网亏损比"]}%',
    "≥40% 明细行占比 22.79%": f'{M["M2_high_discount"]["明细行占比"]}%',
    "桌子累计亏损 13.34 万": "−13.34 万",
    "美术亏损 1.83 万": "1.83 万",
    "退货率 10.69%": f'{M["M6_return"]["真实退货率"]}%',
    "错误口径 41.99%": f'{M["M6_return"]["错误口径_按退货表行数"]}%',
    "净亏损客户 139 位": "139 位",
    "净亏损客户占比 17.59%": f'{M["M4_customer_stats"]["亏损客户占比"]}%',
    "净亏损合计 26.9 万": "26.9 万",
    "家具利润率 11.14%": f'{g("M3_category","类别","家具")["利润率"]}%',
    "技术利润率 13.73%": f'{g("M3_category","类别","技术")["利润率"]}%',
    "办公用品利润率 15.57%": f'{g("M3_category","类别","办公用品")["利润率"]}%',
    "家具销售额占比 35.69%": f'{g("M3_category","类别","家具")["销售额占比"]}%',
    "华东利润率 12.96%": f'{g("M5_region","地区","华东")["利润率"]}%',
    "中南利润率 16.22%": f'{g("M5_region","地区","中南")["利润率"]}%',
    "东北利润率 9.03%": f'{g("M5_region","地区","东北")["利润率"]}%',
    "华北利润率 17.61%": f'{g("M5_region","地区","华北")["利润率"]}%',
    "西南利润率 7.49%": f'{g("M5_region","地区","西南")["利润率"]}%',
    "西北利润率 12.09%": f'{g("M5_region","地区","西北")["利润率"]}%',
    "西南退货率 12.45%": f'{g("M5_region","地区","西南")["退货率"]}%',
    "华北退货率 8.62%": f'{g("M5_region","地区","华北")["退货率"]}%',
    "消费者利润率 13.12%": f'{g("M4_segment","细分","消费者")["利润率"]}%',
    "公司利润率 13.23%": f'{g("M4_segment","细分","公司")["利润率"]}%',
    "小型企业利润率 14.27%": f'{g("M4_segment","细分","小型企业")["利润率"]}%',
    "退货·办公用品 10.77%": f'{g("M6_return_by_category","类别","办公用品")["退货率"]}%',
    "退货·家具 11.06%": f'{g("M6_return_by_category","类别","家具")["退货率"]}%',
    "退货·技术 10.07%": f'{g("M6_return_by_category","类别","技术")["退货率"]}%',
    "当日退货率 9.27%": f'{g("M6_shipping","邮寄方式","当日")["退货率"]}%',
    "二级退货率 11.89%": f'{g("M6_shipping","邮寄方式","二级")["退货率"]}%',
    "标准级退货率 10.47%": f'{g("M6_shipping","邮寄方式","标准级")["退货率"]}%',
    "一级退货率 10.24%": f'{g("M6_shipping","邮寄方式","一级")["退货率"]}%',
    "当日利润率 16.24%": f'{g("M6_shipping","邮寄方式","当日")["利润率"]}%',
    "what-if 保留利润率 23.35%": f'{M["M2_whatif"]["scenario_40"]["保留利润率"]}%',
    "what-if 销售额 -18.96%": "18.96%",
    "what-if 利润 +41.63%": f'{M["M2_whatif"]["scenario_40"]["利润变化率"]}%',
    "2019 销售额同比 +28.7%": "28.7%",
    "2019 利润同比 +9.5%": "9.5%",
    "Q1 247.4 万": "247.4 万",
    "Q4 513.4 万": "513.4 万",
}
for name, val in claims.items():
    chk(f"报告含「{name}」", val in report, f"未找到 {val!r}")

# 反向检查：已知的错误写法不得残留
print("\n【4】已知错误写法残留扫描")
for wrong, right in [("占明细行 20.7%", "22.79%"), ("35.68%", "35.69%"), ("13.24%", "13.23%"),
                     ("16.24% | 10.33%", "当日退货率应为 9.27%")]:
    chk(f"报告已不含「{wrong}」", wrong not in report, f"仍残留，应改为 {right}")

print("\n" + "=" * 74)
print(f"结果：通过 {ok} 项，失败 {bad} 项")
print("=" * 74)
sys.exit(1 if bad else 0)
