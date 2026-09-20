# excel2pandas/_run.py
# -*- coding: utf-8 -*-
"""
本地调试入口 —— 在 VSCode 里右键 "Run Python File" 直接运行。

原理
----
右键运行时，Python 把本文件当作**独立脚本**执行（不是包的一部分），
所以 `from .xxx` 这类相对导入会失败。这里做两件事解决：

  ① 把包的上一级目录加进 sys.path，让 `import excel2pandas` 能找到包
  ② 全程使用绝对导入（from excel2pandas.xxx import ...）

用法
----
改下面 PARAMS 里的参数 → 保存 → 右键 → "Run Python File"。

默认跑的是仓库自带的**虚构演示案例** `examples/demo_template.xlsx`（电商定价场景）。
换成自己的表时，把 EXCEL / SHEET / ROW / KEY 改掉即可；
所有路径都相对**项目根目录**解析，所以在 VSCode 里从任意位置点运行都不会跑偏。
"""

import sys
from pathlib import Path

# ⭐ 关键：把包的上一级加进 sys.path，否则右键运行时找不到 excel2pandas 包
_HERE = Path(__file__).resolve()
_PROJ_ROOT = _HERE.parent.parent          # excel2pandas/ 的上一级

if str(_PROJ_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJ_ROOT))

# Windows 控制台默认 GBK，中文/特殊字符输出会直接抛 UnicodeEncodeError
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

# ---- 绝对导入（不是相对导入）----
import pandas as pd

from excel2pandas import reader as R
from excel2pandas.generator import generate
from excel2pandas.runner import run_single
from excel2pandas.verifier import verify, write_verify_report


# ============================================================
# 改这里：参数写死，右键跑
# ============================================================
EXCEL = _PROJ_ROOT / "examples" / "demo_template.xlsx"   # Excel 文件（仓库自带演示案例）
SHEET = "定价测算"                                        # sheet 名
ROW = 2                                                  # 样板行号（带公式的那一行）
HEADER_ROW = 1                                           # 表头行号
OUT_DIR = _PROJ_ROOT / "output"                          # 输出目录（已被 .gitignore 忽略）
KEY = "SKU-10001"                                        # 用于 run / verify 的主键
# ============================================================


def main():
    # ---- ⓪ 前置检查：路径不对就说清楚，别让人对着 traceback 猜 ----
    if not Path(EXCEL).exists():
        print(f"[错误] 找不到 Excel 文件：{EXCEL}")
        print("       两种可能：")
        print("       1) 你还没生成演示案例 —— 先跑一次：python examples/build_examples.py")
        print("       2) 你想跑自己的表 —— 把上面的 EXCEL / SHEET / KEY 改掉")
        return 1

    # ---- ① 生成代码 ----
    print("=" * 60)
    print("① 生成代码")
    print("=" * 60)
    rep = generate(
        excel_path=EXCEL,
        sheet_name=SHEET,
        target_row=ROW,
        header_row=HEADER_ROW,
        out_dir=OUT_DIR,
        force=True,                 # 调试时强制重新生成，避免缓存
    )
    print(f"  列数    : {rep.n_columns}（公式 {rep.n_formula} / 数据 {rep.n_data}）")
    print(f"  成功    : {rep.n_success}")
    print(f"  失败    : {rep.n_fail}")
    print(f"  CONFIG  : {rep.n_config} 个常量")
    print(f"  代码    : {rep.code_path}")
    print(f"  报告    : {rep.migration_path}")
    print(f"  规则    : {rep.rules_path}")
    if rep.errors:
        print("\n  [失败列]")
        for letter, name, err in rep.errors:
            print(f"    {letter:>3} {name:<20} {err}")
    if rep.warnings:
        print(f"\n  [自动纠正 {len(rep.warnings)} 处]")
        for w in dict.fromkeys(rep.warnings):
            print(f"    - {w}")

    # ---- ② 运行单条 ----
    print()
    print("=" * 60)
    print(f"② 运行单条：{KEY}")
    print("=" * 60)
    try:
        df = run_single(rep.code_path, KEY)
        pd.set_option("display.width", 220)
        pd.set_option("display.max_columns", 60)
        print(df.T)
    except Exception as e:  # noqa: BLE001
        print(f"  [错误] {type(e).__name__}: {e}")
        return 1

    # ---- ③ 对账 ----
    print()
    print("=" * 60)
    print("③ 与 Excel 对账")
    print("=" * 60)
    try:
        model = R.read_model(EXCEL, SHEET, target_row=ROW, header_row=HEADER_ROW)
        vrep = verify(df, EXCEL, SHEET, model, header_row=HEADER_ROW)
        report_path = Path(OUT_DIR) / "verify_report.md"
        write_verify_report(
            vrep, report_path,
            code_path=rep.code_path,
            excel_path=str(EXCEL),
            sheet_name=SHEET,
        )
        print(f"  比较单元格: {vrep.compared_cells}")
        print(f"  一致率    : {vrep.ok_rate:.2f}%")
        for k, v in vrep.counts.items():
            print(f"    {k}: {v}")
        print(f"  报告      : {report_path}")
    except Exception as e:  # noqa: BLE001
        print(f"  [错误] {type(e).__name__}: {e}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
