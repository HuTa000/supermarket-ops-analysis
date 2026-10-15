# -*- coding: utf-8 -*-
"""
W2-W3 · 六大业务模块分析（M1-M6）
输入：outputs/clean_orders.csv（W1 产出）
输出：outputs/metrics.json（供 Excel 与报告使用，Python 口径为权威口径）
      outputs/dashboard_data.json（紧凑行级数据，供看板前端联动计算）
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import numpy as np
import _paths as P

ROOT = P.ROOT
OUT = P.OUT
df = pd.read_csv(os.path.join(OUT, "clean_orders.csv"), encoding="utf-8-sig")


def r0(x):
    return round(float(x), 0)


def r2(x):
    return round(float(x), 2)


def pct(a, b):
    return round(float(a) / float(b) * 100, 2) if b else 0.0


M = {}

# ============================================================
# 全局基线
# ============================================================
TOTAL_SALES = df["销售额"].sum()
TOTAL_PROFIT = df["利润"].sum()
TOTAL_ROWS = len(df)
TOTAL_ORDERS = df["订单 ID"].nunique()
M["baseline"] = {
    "销售额": r0(TOTAL_SALES), "利润": r0(TOTAL_PROFIT),
    "利润率": pct(TOTAL_PROFIT, TOTAL_SALES),
    "明细行": TOTAL_ROWS, "订单数": TOTAL_ORDERS,
    "客户数": int(df["客户 ID"].nunique()), "产品数": int(df["产品 ID"].nunique()),
    "起止": f"{df['订单日期'].min()[:7]} ~ {df['订单日期'].max()[:7]}",
}

# ============================================================
# M1 · 经营趋势（这四年生意变好了吗）
# ============================================================
y = df.groupby("年").agg(销售额=("销售额", "sum"), 利润=("利润", "sum"),
                         订单数=("订单 ID", "nunique")).reset_index()
y["利润率"] = (y["利润"] / y["销售额"] * 100).round(2)
y["销售额同比"] = (y["销售额"].pct_change() * 100).round(2)
y["利润同比"] = (y["利润"].pct_change() * 100).round(2)
y["客单价"] = (y["销售额"] / y["订单数"]).round(0)
M["M1_yearly"] = [
    {"年": int(row.年), "销售额": r0(row.销售额), "利润": r0(row.利润),
     "利润率": r2(row.利润率), "订单数": int(row.订单数), "客单价": r0(row.客单价),
     "销售额同比": (None if pd.isna(row.销售额同比) else r2(row.销售额同比)),
     "利润同比": (None if pd.isna(row.利润同比) else r2(row.利润同比))}
    for row in y.itertuples()
]

q = df.groupby("季").agg(销售额=("销售额", "sum"), 利润=("利润", "sum")).reset_index()
M["M1_quarterly"] = [{"季": f"Q{int(row.季)}", "销售额": r0(row.销售额),
                      "利润": r0(row.利润), "利润率": pct(row.利润, row.销售额)}
                     for row in q.itertuples()]

m = df.groupby("月").agg(销售额=("销售额", "sum"), 利润=("利润", "sum")).reset_index()
M["M1_monthly"] = [{"月": f"{int(row.月)}月", "销售额": r0(row.销售额), "利润": r0(row.利润)}
                   for row in m.itertuples()]

ym = df.groupby("年月").agg(销售额=("销售额", "sum"), 利润=("利润", "sum")).reset_index()
M["M1_monthly_series"] = [{"年月": row.年月, "销售额": r0(row.销售额), "利润": r0(row.利润)}
                          for row in ym.itertuples()]

# 2019 vs 2018 拆解：增收不增利
y18 = y[y.年 == 2018].iloc[0]
y19 = y[y.年 == 2019].iloc[0]
M["M1_gap"] = {
    "销售额增长": r0(y19.销售额 - y18.销售额),
    "销售额增幅": pct(y19.销售额 - y18.销售额, y18.销售额),
    "利润增长": r0(y19.利润 - y18.利润),
    "利润增幅": pct(y19.利润 - y18.利润, y18.利润),
    "利润率变化": r2(y19.利润率 - y18.利润率),
}
# 四年里"低折扣(<=0.2) vs 高折扣(>=0.4) 销售额占比"演变，解释利润率下滑
for yr in (2016, 2019):
    sub = df[df.年 == yr]
    M["M1_gap"][f"{yr}_高折扣销售额占比"] = pct(sub.loc[sub.折扣 >= 0.4, "销售额"].sum(), sub["销售额"].sum())
    M["M1_gap"][f"{yr}_利润率"] = pct(sub["利润"].sum(), sub["销售额"].sum())

# M1 归因：2019 利润率下滑到底是谁造成的（排除法）
CATS = ["家具", "技术", "办公用品"]
attr_disc, attr_nodisc, attr_cat_share, attr_cat_margin, attr_desk = {}, {}, {}, {}, {}
for yr in (2016, 2017, 2018, 2019):
    s = df[df.年 == yr]
    attr_disc[str(yr)] = r2((s.折扣 * s.销售额).sum() / s.销售额.sum() * 100)
    z = s[s.折扣 == 0]
    attr_nodisc[str(yr)] = pct(z.利润.sum(), z.销售额.sum())
    attr_cat_share[str(yr)] = {c: pct(s.loc[s.类别 == c, "销售额"].sum(), s["销售额"].sum()) for c in CATS}
    attr_cat_margin[str(yr)] = {c: pct(s.loc[s.类别 == c, "利润"].sum(), s.loc[s.类别 == c, "销售额"].sum()) for c in CATS}
    attr_desk[str(yr)] = r0(s.loc[s.子类别 == "桌子", "利润"].sum())
M["M1_attribution"] = {
    "加权平均折扣": attr_disc,
    "无折扣业务利润率": attr_nodisc,
    "类别销售额占比": attr_cat_share,
    "类别利润率": attr_cat_margin,
    "桌子利润": attr_desk,
    "排除项": "折扣结构（加权平均折扣四年稳定在 8.5%~9.7%）与无折扣业务利润率（稳定在 24% 上下）均无恶化，"
              "故 2019 利润率下滑不是折扣造成的。",
    "真因": f"家具销售额占比从 2018 的 {attr_cat_share['2018']['家具']}% 升至 2019 的 {attr_cat_share['2019']['家具']}%，"
            f"同时家具利润率从 {attr_cat_margin['2018']['家具']}% 跌至 {attr_cat_margin['2019']['家具']}%；"
            f"其子类「桌子」的亏损从 {abs(attr_desk['2018']):,.0f} 元扩大到 {abs(attr_desk['2019']):,.0f} 元（扩大约 3.5 倍）。",
}

# ============================================================
# M2 · 折扣与利润（利润被什么吃掉了）
# ============================================================
d = df.groupby("折扣").agg(明细行=("行 ID", "count"), 销售额=("销售额", "sum"),
                            利润=("利润", "sum")).reset_index()
d["销售额占比"] = (d["销售额"] / TOTAL_SALES * 100).round(2)
d["利润率"] = (d["利润"] / d["销售额"] * 100).round(2)
d["平均单笔销售额"] = (d["销售额"] / d["明细行"]).round(0)
d["累计利润贡献"] = d["利润"].cumsum().round(0)
M["M2_discount"] = [
    {"折扣": f"{int(round(row.折扣*100))}%", "折扣值": r2(row.折扣),
     "明细行": int(row.明细行), "销售额": r0(row.销售额), "销售额占比": r2(row.销售额占比),
     "利润": r0(row.利润), "利润率": r2(row.利润率), "平均单笔销售额": r0(row.平均单笔销售额)}
    for row in d.itertuples()
]

loss_total = float(df.loc[df["利润"] < 0, "利润"].sum())
high = df[df["折扣"] >= 0.4]
M["M2_high_discount"] = {
    "阈值": "折扣 ≥ 40%",
    "销售额": r0(high["销售额"].sum()),
    "销售额占比": pct(high["销售额"].sum(), TOTAL_SALES),
    "利润": r0(high["利润"].sum()),
    "亏损占全网亏损比": pct(high["利润"].sum(), loss_total),
    "明细行占比": pct(len(high), TOTAL_ROWS),
    "平均利润率": pct(high["利润"].sum(), high["销售额"].sum()),
}
mid = df[df["折扣"] >= 0.2]
M["M2_mid_discount"] = {
    "阈值": "折扣 ≥ 20%",
    "销售额占比": pct(mid["销售额"].sum(), TOTAL_SALES),
    "利润": r0(mid["利润"].sum()),
    "平均利润率": pct(mid["利润"].sum(), mid["销售额"].sum()),
}


def whatif(thr):
    keep = df[df["折扣"] < thr]
    return {
        "剔除阈值": f"折扣 ≥ {int(thr*100)}%",
        "保留销售额": r0(keep["销售额"].sum()),
        "销售额变化": r0(keep["销售额"].sum() - TOTAL_SALES),
        "销售额变化率": pct(keep["销售额"].sum() - TOTAL_SALES, TOTAL_SALES),
        "保留利润": r0(keep["利润"].sum()),
        "利润变化": r0(keep["利润"].sum() - TOTAL_PROFIT),
        "利润变化率": pct(keep["利润"].sum() - TOTAL_PROFIT, TOTAL_PROFIT),
        "保留利润率": pct(keep["利润"].sum(), keep["销售额"].sum()),
    }


M["M2_whatif"] = {"scenario_40": whatif(0.4), "scenario_20": whatif(0.2)}

# 类别 × 折扣档 利润矩阵（定位"哪个品类在哪个折扣档失血"）
cat_rows = []
for cat in ["家具", "技术", "办公用品"]:
    row = {"类别": cat}
    for dv in sorted(df["折扣"].unique()):
        sub = df[(df.类别 == cat) & (df.折扣 == dv)]
        row[f"{int(round(dv*100))}%"] = r0(sub["利润"].sum())
    cat_rows.append(row)
M["M2_cat_discount_profit"] = cat_rows

# ============================================================
# M3 · 品类结构（哪些品类在失血）
# ============================================================
c = df.groupby("类别").agg(销售额=("销售额", "sum"), 利润=("利润", "sum"),
                           明细行=("行 ID", "count")).reset_index()
c["利润率"] = (c["利润"] / c["销售额"] * 100).round(2)
c["销售额占比"] = (c["销售额"] / TOTAL_SALES * 100).round(2)
c["平均折扣"] = df.groupby("类别")["折扣"].mean().values
M["M3_category"] = [
    {"类别": row.类别, "销售额": r0(row.销售额), "销售额占比": r2(row.销售额占比),
     "利润": r0(row.利润), "利润率": r2(row.利润率), "平均折扣": r2(row.平均折扣),
     "明细行": int(row.明细行)}
    for row in c.sort_values("销售额", ascending=False).itertuples()
]

s = df.groupby("子类别").agg(类别=("类别", "first"), 销售额=("销售额", "sum"),
                              利润=("利润", "sum"), 明细行=("行 ID", "count")).reset_index()
s["利润率"] = (s["利润"] / s["销售额"] * 100).round(2)
s["平均折扣"] = df.groupby("子类别")["折扣"].mean().values
med_sales = s["销售额"].median()
med_margin = s["利润率"].median()
s["象限"] = s.apply(
    lambda x: ("明星·高销高利" if x.销售额 >= med_sales and x.利润率 >= med_margin else
               "问题·高销低利" if x.销售额 >= med_sales else
               "潜力·低销高利" if x.利润率 >= med_margin else "瘦狗·低销低利"), axis=1)
M["M3_subcategory"] = [
    {"子类别": row.子类别, "类别": row.类别, "销售额": r0(row.销售额), "利润": r0(row.利润),
     "利润率": r2(row.利润率), "平均折扣": r2(row.平均折扣),
     "明细行": int(row.明细行), "象限": row.象限}
    for row in s.sort_values("利润").itertuples()
]
M["M3_quadrant_median"] = {"销售额中位数": r0(med_sales), "利润率中位数": r2(med_margin)}

# ============================================================
# M4 · 客户价值（谁值得投，谁在倒贴）
# ============================================================
cust = df.groupby("客户 ID").agg(订单数=("订单 ID", "nunique"), 销售额=("销售额", "sum"),
                                 利润=("利润", "sum"), 平均折扣=("折扣", "mean"),
                                 最近购买=("订单日期", "max"), 首次购买=("订单日期", "min")).reset_index()
cust["利润率"] = cust["利润"] / cust["销售额"]

# 客户匿名化映射（保护隐私：原字段含人名）
cust_map = {cid: f"C{i+1:04d}" for i, cid in enumerate(sorted(cust["客户 ID"].unique()))}
cust["客户编号"] = cust["客户 ID"].map(cust_map)

loss_cust = cust[cust["利润"] < 0].sort_values("利润")

M["M4_customer_stats"] = {
    "客户总数": int(len(cust)),
    "亏损客户数": int(len(loss_cust)),
    "亏损客户占比": pct(len(loss_cust), len(cust)),
    "亏损客户合计亏损": r0(loss_cust["利润"].sum()),
    "客均订单数": r2(cust["订单数"].mean()),
    "客均销售额": r0(cust["销售额"].mean()),
    "客均利润": r0(cust["利润"].mean()),
    "订单数中位数": r2(cust["订单数"].median()),
    "订单数最多": int(cust["订单数"].max()),
    "一次下单客户占比": pct((cust["订单数"] == 1).sum(), len(cust)),
    "销售额分位": {"P25": r0(cust["销售额"].quantile(.25)), "P50": r0(cust["销售额"].quantile(.50)),
                   "P75": r0(cust["销售额"].quantile(.75)), "P90": r0(cust["销售额"].quantile(.90)),
                   "P99": r0(cust["销售额"].quantile(.99))},
}
M["M4_loss_customers"] = [
    {"客户编号": row.客户编号, "订单数": int(row.订单数), "销售额": r0(row.销售额),
     "利润": r0(row.利润), "利润率": r2(row.利润率 * 100), "平均折扣": r2(row.平均折扣)}
    for row in loss_cust.head(15).itertuples()
]

sg = df.groupby("细分").agg(客户数=("客户 ID", "nunique"), 订单数=("订单 ID", "nunique"),
                            销售额=("销售额", "sum"), 利润=("利润", "sum"),
                            平均折扣=("折扣", "mean")).reset_index()
sg["客单价"] = sg["销售额"] / sg["订单数"]
sg["利润率"] = sg["利润"] / sg["销售额"] * 100
M["M4_segment"] = [
    {"细分": row.细分, "客户数": int(row.客户数), "订单数": int(row.订单数),
     "客单价": r0(row.客单价), "销售额": r0(row.销售额), "利润": r0(row.利润),
     "利润率": r2(row.利润率), "平均折扣": r2(row.平均折扣)}
    for row in sg.sort_values("销售额", ascending=False).itertuples()
]

# ============================================================
# M5 · 区域业绩（哪个大区该换打法）
# ============================================================
staff = df.drop_duplicates(subset=["订单 ID"])[["订单 ID", "地区", "地区经理", "是否退货"]]
reg_ret = staff.groupby("地区")["是否退货"].mean() * 100
r = df.groupby("地区").agg(销售额=("销售额", "sum"), 利润=("利润", "sum"),
                           订单数=("订单 ID", "nunique"), 客户数=("客户 ID", "nunique"),
                           平均折扣=("折扣", "mean")).reset_index()
r["利润率"] = (r["利润"] / r["销售额"] * 100).round(2)
r["销售额占比"] = (r["销售额"] / TOTAL_SALES * 100).round(2)
r["退货率"] = r["地区"].map(reg_ret).round(2)
r["地区经理"] = r["地区"].map(dict(zip(df["地区"], df["地区经理"])))
M["M5_region"] = [
    {"地区": row.地区, "地区经理": row.地区经理, "销售额": r0(row.销售额),
     "销售额占比": r2(row.销售额占比), "利润": r0(row.利润), "利润率": r2(row.利润率),
     "订单数": int(row.订单数), "客户数": int(row.客户数),
     "平均折扣": r2(row.平均折扣), "退货率": r2(row.退货率)}
    for row in r.sort_values("销售额", ascending=False).itertuples()
]

# 年度 × 大区 销售额矩阵
pivot = df.pivot_table(index="地区", columns="年", values="销售额", aggfunc="sum").fillna(0).round(0)
M["M5_region_year"] = [{"地区": idx, **{str(c): r0(v) for c, v in row.items()}}
                       for idx, row in pivot.iterrows()]

# 高折扣在各地区的渗透（解释低利区）
hd = df[df.折扣 >= 0.4]
hd_share = (hd.groupby("地区")["销售额"].sum() / df.groupby("地区")["销售额"].sum() * 100).round(2)
M["M5_region_high_discount"] = [{"地区": k, "高折扣销售额占比": r2(v)} for k, v in hd_share.items()]

# ============================================================
# M6 · 退货与履约（退货是真问题还是噪音）
# ============================================================
qj = json.load(open(os.path.join(OUT, "quality.json"), encoding="utf-8"))["质量台账"]
ol = df.drop_duplicates(subset=["订单 ID"])
M["M6_return"] = {
    "真实退货率": qj["真实退货率_按订单"],
    "错误口径_按退货表行数": qj["错误口径_按退货表行数"],
    "退货表原始行数": qj["退货表_原始行数"],
    "退货表去重后": qj["退货表_去重后行数"],
    "删除重复行": qj["退货表_删除重复行数"],
}
M["M6_return_by_category"] = [
    {"类别": k, "订单数": int(v["size"]), "退货率": r2(v["mean"] * 100)}
    for k, v in ol.groupby("类别")["是否退货"].agg(["mean", "size"]).iterrows()
]
M["M6_return_by_region"] = [
    {"地区": k, "订单数": int(v["size"]), "退货率": r2(v["mean"] * 100)}
    for k, v in ol.groupby("地区")["是否退货"].agg(["mean", "size"]).sort_values("mean", ascending=False).iterrows()
]
ship = df.groupby("邮寄方式").agg(明细行=("行 ID", "count"), 销售额=("销售额", "sum"),
                                  利润=("利润", "sum"), 平均发货天数=("发货天数", "mean")).reset_index()
ship["利润率"] = (ship["利润"] / ship["销售额"] * 100).round(2)
order_ship = ol.groupby("邮寄方式")["是否退货"].mean() * 100
ship["退货率"] = ship["邮寄方式"].map(order_ship).round(2)
M["M6_shipping"] = [
    {"邮寄方式": row.邮寄方式, "明细行": int(row.明细行), "平均发货天数": r2(row.平均发货天数),
     "销售额": r0(row.销售额), "利润率": r2(row.利润率), "退货率": r2(row.退货率)}
    for row in ship.sort_values("利润率", ascending=False).itertuples()
]

# ============================================================
# 写 metrics.json
# ============================================================
with open(os.path.join(OUT, "metrics.json"), "w", encoding="utf-8") as f:
    json.dump(M, f, ensure_ascii=False, indent=2)

# ============================================================
# 写 dashboard_data.json（紧凑行级数据 + 匿名客户编号）
# ============================================================
df["客户编号"] = df["客户 ID"].map(cust_map)
df["订单序号"] = df["订单 ID"].map({oid: i for i, oid in enumerate(sorted(df["订单 ID"].unique()))})
cols = ["年", "季", "月", "地区", "类别", "子类别", "细分", "邮寄方式",
        "折扣", "销售额", "数量", "利润", "发货天数", "是否退货", "客户编号", "订单序号"]
rows = []
for rec in df[cols].itertuples(index=False):
    rows.append([int(rec[0]), int(rec[1]), int(rec[2]), rec[3], rec[4], rec[5], rec[6], rec[7],
                 round(float(rec[8]), 2), round(float(rec[9]), 2), int(rec[10]), round(float(rec[11]), 2),
                 int(rec[12]), 1 if rec[13] else 0, rec[14], int(rec[15])])
dash = {
    "columns": cols,
    "rows": rows,
    "dims": {
        "年": sorted(df["年"].unique().tolist()),
        "地区": df.groupby("地区")["销售额"].sum().sort_values(ascending=False).index.tolist(),
        "类别": df.groupby("类别")["销售额"].sum().sort_values(ascending=False).index.tolist(),
        "子类别": df.groupby("子类别")["销售额"].sum().sort_values(ascending=False).index.tolist(),
        "细分": df.groupby("细分")["销售额"].sum().sort_values(ascending=False).index.tolist(),
        "邮寄方式": ["标准级", "二级", "一级", "当日"],
        "折扣": sorted(df["折扣"].unique().tolist()),
        "地区经理": {k: v for k, v in zip(df["地区"], df["地区经理"])},
    },
}
with open(os.path.join(OUT, "dashboard_data.json"), "w", encoding="utf-8") as f:
    json.dump(dash, f, ensure_ascii=False, separators=(",", ":"))

# ============================================================
# 自检与打印
# ============================================================
assert abs(sum(x["销售额"] for x in M["M2_discount"]) - TOTAL_SALES) < 1, "M2 折扣销售额合计对不上"
assert abs(sum(x["利润"] for x in M["M2_discount"]) - TOTAL_PROFIT) < 1, "M2 折扣利润合计对不上"
assert abs(sum(x["销售额"] for x in M["M3_category"]) - TOTAL_SALES) < 1, "M3 类别销售额合计对不上"
assert abs(sum(x["销售额"] for x in M["M5_region"]) - TOTAL_SALES) < 1, "M5 大区销售额合计对不上"

print("=" * 72)
print("W2-W3 六大业务模块 · 分析完成")
print("=" * 72)
print(f"基线：销售额 {TOTAL_SALES:,.0f} | 利润 {TOTAL_PROFIT:,.0f} | 利润率 {pct(TOTAL_PROFIT, TOTAL_SALES)}%")
print()
print("【M1 经营趋势】")
for x in M["M1_yearly"]:
    print(f"  {x['年']}: 销售额 {x['销售额']:>10,.0f} | 利润 {x['利润']:>9,.0f} | "
          f"利润率 {x['利润率']:>5.2f}% | 销售额同比 {x['销售额同比']}")
print(f"  2019 vs 2018：销售额 +{M['M1_gap']['销售额增幅']}%，利润 +{M['M1_gap']['利润增幅']}%，"
      f"利润率 {M['M1_gap']['利润率变化']:+.2f}pct")
print(f"  高折扣(>=40%)销售额占比：2016 {M['M1_gap']['2016_高折扣销售额占比']}% -> "
      f"2019 {M['M1_gap']['2019_高折扣销售额占比']}%")
print()
print("【M2 折扣与利润】")
for x in M["M2_discount"]:
    print(f"  折扣 {x['折扣']:>4}: 占销售额 {x['销售额占比']:>5.2f}% | 利润 {x['利润']:>10,.0f} | 利润率 {x['利润率']:>7.2f}%")
h = M["M2_high_discount"]
print(f"  ≥40% 档：销售额占比 {h['销售额占比']}%，亏损占全网亏损 {h['亏损占全网亏损比']}%，合计亏 {h['利润']:,.0f}")
print(f"  what-if 剔除 ≥40%：销售额 {M['M2_whatif']['scenario_40']['销售额变化率']}%，"
      f"利润 {M['M2_whatif']['scenario_40']['利润变化率']}%")
print()
print("【M3 品类】")
for x in M["M3_category"]:
    print(f"  {x['类别']}: 销售额 {x['销售额']:>10,.0f} | 利润 {x['利润']:>9,.0f} | 利润率 {x['利润率']:>5.2f}%")
print("  亏损子类别：", [(x["子类别"], x["利润"]) for x in M["M3_subcategory"] if x["利润"] < 0])
print()
print("【M4 客户】")
st = M["M4_customer_stats"]
print(f"  客户 {st['客户总数']} 人 | 亏损客户 {st['亏损客户数']} 人（{st['亏损客户占比']}%）"
      f" 合计亏 {st['亏损客户合计亏损']:,.0f} | 客均订单 {st['客均订单数']}")
print()
print("【M5 大区】")
for x in M["M5_region"]:
    print(f"  {x['地区']}({x['地区经理']}): 销售额 {x['销售额']:>10,.0f} | 利润率 {x['利润率']:>5.2f}% | 退货率 {x['退货率']:>5.2f}%")
print()
print("【M6 退货】")
print(f"  真实退货率 {M['M6_return']['真实退货率']}%（错误口径 {M['M6_return']['错误口径_按退货表行数']}%）")
print("=" * 72)
print("✅ metrics.json / dashboard_data.json 已写出")
