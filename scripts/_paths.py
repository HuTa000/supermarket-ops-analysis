# -*- coding: utf-8 -*-
"""
项目路径与配置解析（与项目所在盘符无关）。

设计目标：
1. 项目整体移动到任意目录/盘符都能直接运行 —— 所有路径基于 __file__ 反推，不写死盘符。
2. 换一组数据时，只改 config.json（或命令行 --input），不改脚本。

目录约定：
    <项目根>/
        config.json     配置（输入文件、表名映射、字段映射）
        data/           放原始数据
        scripts/        脚本
        templates/      看板模板
        vendor/         第三方库（ECharts）
        outputs/        中间产物（干净宽表、指标 JSON）
"""
import os
import sys
import json
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))   # .../scripts
ROOT = os.path.dirname(HERE)                        # 项目根
OUT = os.path.join(ROOT, "outputs")
DATA = os.path.join(ROOT, "data")
CONFIG_PATH = os.path.join(ROOT, "config.json")

# 支持的表格后缀（按优先级）
EXCEL_EXT = [".xlsx", ".xlsm", ".xls"]

# ------------------------------------------------------------------
# 默认配置：换数据时在 config.json 里覆盖对应键即可
# ------------------------------------------------------------------
DEFAULTS = {
    # 输入文件：相对项目根 或 绝对路径；留空 = 自动取 data/ 下唯一的表格文件
    "输入文件": "",
    # 三份交付物的文件名前缀
    "输出前缀": "超市运营分析",
    # 工作表名映射：{标准名: 源文件中的实际表名}
    "表名": {"订单": "订单", "退货": "退货", "销售人员": "销售人员"},
    # 字段映射：{标准名: 源文件中的实际列名}。源文件列名已是标准名时不必写。
    "字段映射": {},
}

# 脚本要求的字段（缺一个就无法继续，etl 会打印清单）
REQUIRED = {
    "订单": ["订单 ID", "客户 ID", "产品 ID", "行 ID", "订单日期", "发货日期",
             "销售额", "数量", "利润", "折扣", "类别", "子类别", "细分", "邮寄方式", "地区"],
    "退货": ["订单 ID"],
    "销售人员": ["地区", "地区经理"],
}


def load_config():
    cfg = {k: (dict(v) if isinstance(v, dict) else v) for k, v in DEFAULTS.items()}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8") as f:
            user = json.load(f)
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


def _abs(p):
    return p if os.path.isabs(p) else os.path.join(ROOT, p)


def find_input(cli_path=None):
    """定位输入文件。优先级：命令行 > config.json > data/ 下唯一表格文件。"""
    cfg = load_config()
    cand = cli_path or cfg.get("输入文件") or ""

    if cand:
        p = _abs(cand)
        if not os.path.exists(p):
            sys.exit(f"[错误] 输入文件不存在：{p}\n"
                     f"        请检查 config.json 的「输入文件」，或改用 --input 指定。")
        return p

    if not os.path.isdir(DATA):
        sys.exit(f"[错误] 未找到数据目录：{DATA}\n"
                 f"        请把原始数据放进 data/，或在 config.json 里指定「输入文件」。")

    files = [f for f in sorted(os.listdir(DATA))
             if os.path.splitext(f)[1].lower() in EXCEL_EXT and not f.startswith("~$")]
    # 优先 .xlsx/.xlsm，其次 .xls（旧版二进制）
    files.sort(key=lambda f: EXCEL_EXT.index(os.path.splitext(f)[1].lower()))
    if not files:
        sys.exit(f"[错误] data/ 目录下没有可读的表格文件（支持 {', '.join(EXCEL_EXT)}）：{DATA}")
    return os.path.join(DATA, files[0])


def engine_for(path):
    return "xlrd" if os.path.splitext(path)[1].lower() == ".xls" else "openpyxl"


def add_input_arg(parser):
    parser.add_argument("--input", "-i", default=None,
                        help="输入文件路径（覆盖 config.json）。不传则自动在 data/ 下寻找。")


def prefix():
    """三份交付物的文件名前缀。"""
    return load_config().get("输出前缀") or "超市运营分析"


def make_parser(desc):
    p = argparse.ArgumentParser(description=desc)
    add_input_arg(p)
    return p


# 兼容旧的硬编码常量名，方便逐步迁移
RAW_XLSX = os.path.join(DATA, "原始数据.xlsx")
