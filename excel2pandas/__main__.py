# -*- coding: utf-8 -*-
"""
__main__.py —— 命令行入口（薄壳，只有 argparse 路由，无业务逻辑）

    python -m excel2pandas generate --file examples/demo_template.xlsx --sheet "定价测算" --row 2 --out output/
    python -m excel2pandas run      --code output/generated_code.py --key "SKU-10001"
    python -m excel2pandas run      --code output/generated_code.py --keys keys.csv --progress output/progress.json --resume
    python -m excel2pandas verify   --code output/generated_code.py --excel examples/demo_template.xlsx --sheet "定价测算" --report output/verify_report.md
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd


def _keys_from_args(args) -> list:
    if getattr(args, "keys", None):
        p = Path(args.keys)
        df = pd.read_csv(p) if p.suffix.lower() == ".csv" else pd.read_excel(p)
        return df.iloc[:, 0].dropna().astype(str).tolist()
    if getattr(args, "key", None):
        return list(args.key)
    return []


def cmd_generate(args) -> int:
    from .generator import generate

    names = {}
    cfg = Path(args.config_names) if args.config_names else Path(args.out) / "config_names.json"
    if cfg.exists():
        names = json.loads(cfg.read_text(encoding="utf-8"))

    rep = generate(
        excel_path=args.file,
        sheet_name=args.sheet,
        target_row=args.row,
        header_row=args.header_row,
        out_dir=args.out,
        force=args.force,
        scan_external=args.scan_external,
        config_names=names,
        comment_mode=args.comments,
        runtime_mode=args.runtime_mode,
    )
    if rep.skipped:
        print("[cache] 命中缓存，未重新生成（--force 可绕过）")
        print(f"        {rep.code_path}")
        return 0

    print(f"[ok] 输出目录  : {rep.out_dir}")
    print(f"     列数      : {rep.n_columns}（公式 {rep.n_formula} / 数据 {rep.n_data}）")
    print(f"     转换成功  : {rep.n_success}    失败: {rep.n_fail}")
    print(f"     CONFIG    : {rep.n_config} 个常量")
    print(f"     runtime   : {rep.runtime_mode}"
          f"（内联 {len(rep.runtime_helpers)} 个 helper）")
    print(f"     generated_code.py   {rep.code_path}")
    print(f"     business_rules.md   {rep.rules_path}")
    print(f"     migration_report.md {rep.migration_path}")
    if rep.errors:
        print("\n[失败列]")
        for letter, name, err in rep.errors:
            print(f"   {letter:>3} {name:<20} {err}")
    if rep.warnings:
        print(f"\n[自动纠正 {len(rep.warnings)} 处]")
        for w in dict.fromkeys(rep.warnings):
            print(f"   - {w}")
    return 0


def cmd_run(args) -> int:
    from .runner import run_batch, run_single

    keys = _keys_from_args(args)
    if not keys:
        print("需要 --key 或 --keys", file=sys.stderr)
        return 2

    if len(keys) == 1 and not args.progress:
        df = run_single(args.code, keys[0], file_paths=None)
        pd.set_option("display.width", 220)
        pd.set_option("display.max_columns", 60)
        print(df.T)
        return 0

    df, failed = run_batch(args.code, keys, progress_file=args.progress, resume=args.resume)
    print(f"[ok] 共 {len(df)} 行，失败 {len(failed)} 条")
    if failed:
        Path("failed.json").write_text(json.dumps(failed, ensure_ascii=False, indent=1), encoding="utf-8")
        print("     失败清单 -> failed.json")
    if args.out:
        df.to_excel(args.out, index=False)
        print(f"     结果 -> {args.out}")
    return 0


def cmd_verify(args) -> int:
    from . import reader as R
    from .runner import load_module
    from .verifier import verify, write_verify_report

    keys = _keys_from_args(args)
    mod = load_module(args.code)
    df = mod.calc(keys) if keys else mod.calc(["__none__"])

    model = R.read_model(args.excel, args.sheet, target_row=args.row, header_row=args.header_row)
    rep = verify(df, args.excel, args.sheet, model, header_row=args.header_row)
    out = write_verify_report(rep, args.report, code_path=str(args.code),
                              excel_path=str(args.excel), sheet_name=args.sheet)
    print(f"[ok] 对账报告 -> {out}")
    print(f"     比较单元格 {rep.compared_cells}，一致率 {rep.ok_rate:.2f}%")
    for k, v in rep.counts.items():
        print(f"       {k}: {v}")
    return 0


def cmd_report(args) -> int:
    from . import reader as R

    model = R.read_model(args.file, args.sheet, target_row=args.row,
                         header_row=args.header_row, scan_external=True)
    out = R.write_migration_report(model, args.out)
    print(f"[ok] 迁移评估 -> {out}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser("excel2pandas", description="把 Excel 模板公式翻译成 Pandas 代码")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("generate", help="生成代码 + 报告")
    g.add_argument("--file", required=True)
    g.add_argument("--sheet", required=True)
    g.add_argument("--row", type=int, default=2, help="样板行号")
    g.add_argument("--header-row", type=int, default=1)
    g.add_argument("--out", default="output")
    g.add_argument("--config-names", default=None)
    g.add_argument("--force", action="store_true", help="绕过缓存强制重新生成")
    g.add_argument("--scan-external", action="store_true", help="统计外部表行数（慢）")
    g.add_argument("--comments", default="excel", choices=["excel", "both", "none"])
    g.add_argument("--runtime-mode", default="minimal", choices=["minimal", "full"],
                   help="minimal=只内联用到的 helper（默认）；full=全量内联（回退用）")
    g.set_defaults(func=cmd_generate)

    r = sub.add_parser("run", help="运行生成的代码")
    r.add_argument("--code", required=True)
    r.add_argument("--key", action="append", help="单个主键（可重复）")
    r.add_argument("--keys", help="主键清单 CSV/XLSX")
    r.add_argument("--out", help="结果输出 xlsx")
    r.add_argument("--progress", help="进度文件（断点续传）")
    r.add_argument("--resume", action="store_true")
    r.set_defaults(func=cmd_run)

    v = sub.add_parser("verify", help="一键对账")
    v.add_argument("--code", required=True)
    v.add_argument("--excel", required=True)
    v.add_argument("--sheet", required=True)
    v.add_argument("--row", type=int, default=2)
    v.add_argument("--header-row", type=int, default=1)
    v.add_argument("--key", action="append")
    v.add_argument("--keys")
    v.add_argument("--report", default="output/verify_report.md")
    v.set_defaults(func=cmd_verify)

    m = sub.add_parser("report", help="只出迁移评估报告")
    m.add_argument("--file", required=True)
    m.add_argument("--sheet", required=True)
    m.add_argument("--row", type=int, default=2)
    m.add_argument("--header-row", type=int, default=1)
    m.add_argument("--out", default="output/migration_report.md")
    m.set_defaults(func=cmd_report)

    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as e:  # noqa: BLE001
        print(f"[error] {type(e).__name__}: {e}", file=sys.stderr)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
