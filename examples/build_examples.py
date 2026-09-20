# -*- coding: utf-8 -*-
"""
一键重建 examples/ 下的全部产物。

流程
----
1. 造演示模板              make_demo_template.py  →  demo_template.xlsx
2. generate                模板 → generated_code.py + 两份报告
3. 路径归一化              把生成产物里的本机绝对路径改写为仓库内相对路径
4. verify                  生成代码 vs 模板缓存值 逐格对账
5. run --keys              批量跑 keys.csv

用法
----
    python examples/build_examples.py

产物全部是虚构的电商场景数据，可安全提交到公开仓库。
"""

import os
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

SHEET = "定价测算"
ROW = "2"

# 本机绝对路径 → 仓库相对路径
# 匹配形如 X:\...\excel2pandas\ 的前缀。
# 故意不写死具体的用户目录名，这样脚本本身也不含任何本机路径特征。
ABS_RE = re.compile(r"[A-Za-z]:[\\/][^\"'\n]*?excel2pandas[\\/]+")


def _env():
    e = dict(os.environ)
    e["PYTHONIOENCODING"] = "utf-8"      # Windows 控制台默认 GBK，会炸中文输出
    return e


def cli(*args, quiet=False):
    """调用 python -m excel2pandas ..."""
    r = subprocess.run(
        [sys.executable, "-m", "excel2pandas", *args],
        cwd=str(ROOT), env=_env(), text=True,
        encoding="utf-8", errors="replace", capture_output=True,
    )
    if r.returncode != 0:
        print(r.stdout or "", r.stderr or "")
        raise SystemExit(f"命令失败: {' '.join(args)}")
    if not quiet:
        print((r.stdout or "").rstrip())
    return r.stdout or ""


def normalize_paths():
    """产物里不要出现本机绝对路径。"""
    n = 0
    for p in sorted(HERE.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in {".py", ".md", ".csv", ".json"}:
            continue
        if p.name == Path(__file__).name:
            continue
        t0 = t = p.read_text(encoding="utf-8")
        t = ABS_RE.sub("", t)
        if t != t0:
            p.write_text(t, encoding="utf-8")
            n += 1
    print(f"[ok] 路径归一化：改写 {n} 个文件")


def main():
    print("=" * 64)
    print("① 造演示模板")
    print("=" * 64)
    subprocess.run([sys.executable, str(HERE / "make_demo_template.py")],
                   cwd=str(ROOT), env=_env(), check=True)

    print()
    print("=" * 64)
    print("② generate：模板 → 代码 + 报告")
    print("=" * 64)
    cli("generate", "--file", "examples/demo_template.xlsx",
        "--sheet", SHEET, "--row", ROW, "--out", "examples/", "--force")

    print()
    print("=" * 64)
    print("③ 路径归一化")
    print("=" * 64)
    normalize_paths()

    print()
    print("=" * 64)
    print("④ verify：生成代码 vs 模板缓存值")
    print("=" * 64)
    cli("verify", "--code", "examples/generated_code.py",
        "--excel", "examples/demo_template.xlsx", "--sheet", SHEET, "--row", ROW,
        "--key", "SKU-10001", "--report", "examples/verify_report.md")

    print()
    print("=" * 64)
    print("⑤ run --keys：批量执行")
    print("=" * 64)
    (HERE / "keys.csv").write_text("商品编号\nSKU-10001\nSKU-10002\nSKU-10003\n",
                                   encoding="utf-8")
    cli("run", "--code", "examples/generated_code.py", "--keys", "examples/keys.csv",
        "--progress", "examples/progress.json", "--out", "examples/result.xlsx")

    # 缓存文件属于本地运行态，不进仓库
    for junk in (HERE / ".excel2pandas_cache.json",):
        if junk.exists():
            junk.unlink()

    print()
    print("=" * 64)
    print("完成。examples/ 产物：")
    for p in sorted(HERE.iterdir()):
        if p.is_file():
            print(f"  {p.stat().st_size:>8}  {p.name}")


if __name__ == "__main__":
    main()
