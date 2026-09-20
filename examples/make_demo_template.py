# -*- coding: utf-8 -*-
"""
生成演示模板 demo_template.xlsx —— 纯虚构的电商定价测算场景。

⚠️ 本模板的所有数据（商品、类目、平台、金额）均为编造，不含任何真实业务信息。
   examples/ 目录下的全部产物都由这份模板跑出来。

场景结构
--------
主表「定价测算」     1 行表头 + 1 行样板行（公式 + 缓存值）
外部表「运费价目」   类目 → 头程运费
外部表「平台费率」   平台 → 佣金率

为什么用 xlsxwriter
-------------------
本工具依赖「双读」：公式文本 + data_only 的缓存值，后者是对账基线。
xlsxwriter 的 `write_formula(..., value=)` 可以同时写入公式与缓存值，
正好用来造一份「已经算好的」Excel，无需真的打开 Excel 重算。

缓存值按 Excel 语义手工计算并写入（ROUND 远离零、CEILING 向上取整），
运行后可用 `verify` 校验生成代码与之一致。
"""

from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import numpy_financial as nf
import xlsxwriter

HERE = Path(__file__).resolve().parent
OUT = HERE / "demo_template.xlsx"

# ---------------------------------------------------------------- 样板行输入值

SKU = "SKU-10001"
NAME = "示例商品"
CATEGORY = "家居"
PLATFORM = "平台A"
COST = 42.5          # 成本价
QTY = 100            # 采购数量
ADS = 8.0            # 推广费
LIST = 99.0          # 参考售价

FREIGHT_TABLE = [("家居", 25.0), ("服饰", 18.0), ("数码", 32.0), ("美妆", 22.0)]
RATE_TABLE = [("平台A", 0.08), ("平台B", 0.12), ("平台C", 0.15)]


# ---------------------------------------------------------------- Excel 语义


def xl_round(x, n=0):
    """Excel ROUND：四舍五入、远离零（与 Python 银行家舍入不同）。"""
    q = Decimal(1).scaleb(-n)
    return float(Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP))


def xl_ceiling(x, sig):
    """Excel CEILING：向上取整到 sig 的倍数。"""
    import math

    return math.ceil(round(x / sig, 10)) * sig


# ---------------------------------------------------------------- 计算缓存值

FREIGHT = dict(FREIGHT_TABLE)[CATEGORY]
RATE = dict(RATE_TABLE)[PLATFORM]

V_G = FREIGHT                                    # =VLOOKUP(C2,运费价目!$A:$B,2,0)
V_H = xl_round(COST * 1.13, 2)                   # =ROUND(E2*1.13,2)
V_I = RATE                                       # =VLOOKUP(D2,平台费率!$A:$B,2,0)
V_K = xl_round((V_H + V_G + QTY * 0.05) / QTY, 2)  # =ROUND((H2+G2+F2*0.05)/F2,2)
V_M = xl_round(LIST * V_I, 2)                    # =ROUND(L2*I2,2)
V_N = xl_round(LIST - V_M - V_K - ADS, 2)        # =ROUND(L2-M2-K2-J2,2)
V_O = xl_round(V_N / LIST, 4)                    # =ROUND(N2/L2,4)
V_P = "A" if V_N >= 30 else "B" if V_N >= 20 else "C" if V_N >= 10 else "D"
V_Q = xl_ceiling(LIST * 1.2, 5)                  # =CEILING(L2*1.2,5)
V_R = xl_round(nf.pmt(0.06 / 12, 12, -LIST), 2)  # =ROUND(PMT(0.06/12,12,-L2),2)
V_S = "是" if V_O >= 0.3 else "否"                # =IF(O2>=0.3,"是","否")


# ---------------------------------------------------------------- 写模板

wb = xlsxwriter.Workbook(str(OUT))
bold = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
num2 = wb.add_format({"num_format": "0.00"})
num4 = wb.add_format({"num_format": "0.0000"})
pct = wb.add_format({"num_format": "0.0%"})

# ---- 主表 ----
ws = wb.add_worksheet("定价测算")
headers = [
    "商品编号", "商品名称", "类目", "平台", "成本价", "采购数量",
    "头程运费", "含税成本", "佣金率", "推广费", "单件成本", "参考售价",
    "平台佣金", "毛利", "毛利率", "定价档位", "建议售价", "月供测算",
    "是否达标", "备注",
]
for c, h in enumerate(headers):
    ws.write(0, c, h, bold)
ws.set_row(0, 20)

# 数据列（作为输入参数，不生成赋值）
ws.write(1, 0, SKU)
ws.write(1, 1, NAME)
ws.write(1, 2, CATEGORY)
ws.write(1, 3, PLATFORM)
ws.write(1, 4, COST)
ws.write(1, 5, QTY)
ws.write(1, 9, ADS)
ws.write(1, 11, LIST)

# 公式列（公式 + 缓存值）
ws.write_formula(1, 6, "=VLOOKUP(C2,运费价目!$A:$B,2,0)", num2, V_G)
ws.write_formula(1, 7, "=ROUND(E2*1.13,2)", num2, V_H)
ws.write_formula(1, 8, "=VLOOKUP(D2,平台费率!$A:$B,2,0)", pct, V_I)
ws.write_formula(1, 10, "=ROUND((H2+G2+F2*0.05)/F2,2)", num2, V_K)
ws.write_formula(1, 12, "=ROUND(L2*I2,2)", num2, V_M)
ws.write_formula(1, 13, "=ROUND(L2-M2-K2-J2,2)", num2, V_N)
ws.write_formula(1, 14, "=ROUND(N2/L2,4)", num4, V_O)
ws.write_formula(1, 15, '=IF(N2>=30,"A",IF(N2>=20,"B",IF(N2>=10,"C","D")))', None, V_P)
ws.write_formula(1, 16, "=CEILING(L2*1.2,5)", num2, V_Q)
ws.write_formula(1, 17, "=ROUND(PMT(0.06/12,12,-L2),2)", num2, V_R)
ws.write_formula(1, 18, '=IF(O2>=0.3,"是","否")', None, V_S)

ws.set_column(0, 0, 12)
ws.set_column(1, 1, 14)
ws.set_column(2, 19, 11)

# ---- 外部表 1：运费价目 ----
ws2 = wb.add_worksheet("运费价目")
ws2.write(0, 0, "类目", bold)
ws2.write(0, 1, "头程运费", bold)
for i, (k, v) in enumerate(FREIGHT_TABLE, start=1):
    ws2.write(i, 0, k)
    ws2.write(i, 1, v, num2)
ws2.set_column(0, 1, 14)

# ---- 外部表 2：平台费率 ----
ws3 = wb.add_worksheet("平台费率")
ws3.write(0, 0, "平台", bold)
ws3.write(0, 1, "佣金率", bold)
for i, (k, v) in enumerate(RATE_TABLE, start=1):
    ws3.write(i, 0, k)
    ws3.write(i, 1, v, pct)
ws3.set_column(0, 1, 14)

wb.close()

# ---------------------------------------------------------------- 回显

print(f"已生成 {OUT}")
print(f"  主表「定价测算」  {len(headers)} 列 = 8 数据列 + 11 公式列 + 1 空列")
print("  外部表「运费价目」/「平台费率」")
print("\n样板行缓存值：")
for col, val in zip("GHIKMNO P Q R S".replace(" ", ""), [V_G, V_H, V_I, V_K, V_M, V_N, V_O, V_P, V_Q, V_R, V_S]):
    print(f"  {col}2 = {val!r}")
