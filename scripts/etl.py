# -*- coding: utf-8 -*-
"""
W1 · 数据预处理（ETL）
读原始表格 → 三表整合 → 退货去重 → 派生字段 → 质量核查 → 输出干净宽表。

路径与表名/字段名全部走 config.json，项目可整体移动到任意盘符。

业务口径（已确认）：
- 利润率 = 利润 / 销售额
- 负利润是业务真实值，不是异常值，必须保留
- 退货率分母 = 唯一订单数，分子 = 退货表中的唯一订单数（退货表需先去重）
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import numpy as np
import _paths as P

CFG = P.load_config()
OUT = P.OUT
os.makedirs(OUT, exist_ok=True)


def resolve_sheets(xls, cfg):
    """把配置里的表名映射解析到实际 sheet，并给出可读的报错。"""
    want = cfg["表名"]
    got = {}
    for std, actual in want.items():
        if actual in xls.sheet_names:
            got[std] = actual
        else:
            # 容错：按包含关系匹配
            hit = [s for s in xls.sheet_names if actual in s or s in actual]
            got[std] = hit[0] if hit else None
        if got[std] is None:
            sys.exit(
                f"[错误] 找不到工作表「{actual}」。\n"
                f"        文件实际包含：{xls.sheet_names}\n"
                f"        请修改 config.json 的「表名」映射。"
            )
    return got


def resolve_columns(df, std_sheet, cfg_map):
    """按字段映射重命名列；返回 (新df, 缺失字段清单)。"""
    ren = {v: k for k, v in cfg_map.items()
           if not k.startswith("_") and isinstance(v, str)}
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    need = P.REQUIRED.get(std_sheet, [])
    missing = [c for c in need if c not in df.columns]
    return df, missing


def main():
    ap = P.make_parser("W1 数据预处理")
    args = ap.parse_args()

    src = P.find_input(args.input)
    eng = P.engine_for(src)
    xls = pd.ExcelFile(src, engine=eng)
    src_desc = f"{os.path.basename(src)}（{'Excel 2003 .xls' if eng == 'xlrd' else 'Excel .xlsx'}）"

    sheets = resolve_sheets(xls, CFG)
    cmap = CFG.get("字段映射", {})

    orders_raw, miss_o = resolve_columns(xls.parse(sheets["订单"]), "订单", cmap)
    returns_raw, miss_r = resolve_columns(xls.parse(sheets["退货"]), "退货", cmap)
    staff_raw, miss_s = resolve_columns(xls.parse(sheets["销售人员"]), "销售人员", cmap)

    miss = [("订单", m) for m in miss_o] + [("退货", m) for m in miss_r] + [("销售人员", m) for m in miss_s]
    if miss:
        sys.exit(
            "[错误] 源数据缺少必需字段，无法继续：\n"
            + "\n".join(f"        · 表「{s}」缺字段「{c}」" for s, c in miss)
            + "\n        请在 config.json 的「字段映射」里把标准名指到源文件的实际列名。\n"
            + "        各表可用列：\n"
            + f"          {sheets['订单']}: {list(xls.parse(sheets['订单']).columns)}\n"
            + f"          {sheets['退货']}: {list(xls.parse(sheets['退货']).columns)}\n"
            + f"          {sheets['销售人员']}: {list(xls.parse(sheets['销售人员']).columns)}"
        )

    q = {}  # 数据质量台账

    # ---------- 1. 退货表去重 ----------
    returns = returns_raw.drop_duplicates().copy()
    q["退货表_原始行数"] = int(len(returns_raw))
    q["退货表_去重后行数"] = int(len(returns))
    q["退货表_删除重复行数"] = int(len(returns_raw) - len(returns))
    return_order_ids = set(returns["订单 ID"])
    q["退货涉及唯一订单数"] = len(return_order_ids)

    # ---------- 2. 订单表质量核查 ----------
    df = orders_raw.copy()
    q["订单表_明细行数"] = int(len(df))
    q["订单表_唯一订单数"] = int(df["订单 ID"].nunique())
    q["订单表_唯一客户数"] = int(df["客户 ID"].nunique())
    q["订单表_唯一产品数"] = int(df["产品 ID"].nunique())
    q["订单表_全字段重复行"] = int(df.duplicated().sum())
    q["订单表_缺失值总数"] = int(df.isna().sum().sum())
    q["订单表_负销售额行"] = int((df["销售额"] < 0).sum())
    q["订单表_负数量行"] = int((df["数量"] < 0).sum())
    q["订单表_负利润行"] = int((df["利润"] < 0).sum())
    q["订单表_负利润金额"] = round(float(df.loc[df["利润"] < 0, "利润"].sum()), 2)

    # 日期逻辑
    df["订单日期"] = pd.to_datetime(df["订单日期"])
    df["发货日期"] = pd.to_datetime(df["发货日期"])
    bad_date = int((df["发货日期"] < df["订单日期"]).sum())
    q["订单表_发货早于下单行"] = bad_date

    # 字段取值普查
    countries = sorted(df["国家/地区"].dropna().unique().tolist()) if "国家/地区" in df.columns else []
    prefixes = df["订单 ID"].astype(str).str.split("-").str[0].value_counts().to_dict()
    q["国家地区_取值"] = countries
    q["订单ID_前缀分布"] = {k: int(v) for k, v in prefixes.items()}
    q["订单ID_前缀异常说明"] = (
        "订单编号存在多种前缀；判定为编号规则不统一（数据治理问题），"
        "不影响金额口径，分析中不使用订单号做地理划分。"
    )

    # ---------- 3. 派生字段 ----------
    df["年"] = df["订单日期"].dt.year
    df["季"] = df["订单日期"].dt.quarter
    df["月"] = df["订单日期"].dt.month
    df["年月"] = df["订单日期"].dt.to_period("M").astype(str)
    df["发货天数"] = (df["发货日期"] - df["订单日期"]).dt.days
    df["利润率"] = df["利润"] / df["销售额"]
    df["是否亏损行"] = df["利润"] < 0
    df["折扣档"] = df["折扣"].map(lambda x: f"{int(round(x*100))}%")

    # ---------- 4. 三表整合 ----------
    df["是否退货"] = df["订单 ID"].isin(return_order_ids)
    staff_map = dict(zip(staff_raw["地区"], staff_raw["地区经理"]))
    df["地区经理"] = df["地区"].map(staff_map)
    q["地区经理_缺失行"] = int(df["地区经理"].isna().sum())

    # ---------- 5. 退货率（订单粒度） ----------
    order_level = df.drop_duplicates(subset=["订单 ID"]).copy()
    q["真实退货率_按订单"] = round(len(return_order_ids) / order_level["订单 ID"].nunique() * 100, 2)
    q["错误口径_按退货表行数"] = round(len(returns_raw) / order_level["订单 ID"].nunique() * 100, 2)

    # ---------- 6. 输出 ----------
    clean_path = os.path.join(OUT, "clean_orders.csv")
    df.to_csv(clean_path, index=False, encoding="utf-8-sig")

    with open(os.path.join(OUT, "quality.json"), "w", encoding="utf-8") as f:
        json.dump({"数据源": src_desc, "数据源路径": src,
                   "时间范围": f"{df['订单日期'].min().date()} ~ {df['订单日期'].max().date()}",
                   "质量台账": q}, f, ensure_ascii=False, indent=2)

    # ---------- 7. 对账 ----------
    print("=" * 72)
    print("W1 数据预处理 · 对账报告")
    print("=" * 72)
    print(f"数据源：{src_desc}")
    print(f"路径  ：{src}")
    print()
    print("【对账】清洗前后金额与行数必须一致（清洗只去重退货表，不动订单表）")
    print(f"  订单明细行   : {len(orders_raw)} -> {len(df)}  {'✅' if len(orders_raw)==len(df) else '❌'}")
    print(f"  销售额合计   : {orders_raw['销售额'].sum():,.2f} -> {df['销售额'].sum():,.2f}  "
          f"{'✅' if abs(orders_raw['销售额'].sum()-df['销售额'].sum())<0.01 else '❌'}")
    print(f"  利润合计     : {orders_raw['利润'].sum():,.2f} -> {df['利润'].sum():,.2f}  "
          f"{'✅' if abs(orders_raw['利润'].sum()-df['利润'].sum())<0.01 else '❌'}")
    print()
    print("【数据质量台账】")
    for k, v in q.items():
        if k in ("订单ID_前缀分布", "国家地区_取值", "订单ID_前缀异常说明"):
            continue
        print(f"  {k:<26}: {v}")
    print(f"  {'国家地区_取值':<26}: {q['国家地区_取值']}")
    print(f"  {'订单ID_前缀分布':<26}: {q['订单ID_前缀分布']}")
    print()
    print(f"✅ 干净宽表已写出：{clean_path}")
    print(f"✅ 质量台账已写出：{os.path.join(OUT, 'quality.json')}")


if __name__ == "__main__":
    main()
