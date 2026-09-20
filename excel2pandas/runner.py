# -*- coding: utf-8 -*-
"""
runner.py —— 运行器（单条 / 批量 / 断点续传 / 异常映射）

职责：
  1. import 生成的模块 -> 喂主键 -> 得到 df
  2. 单条调试 / 批量执行 / 断点续传
  3. 异常映射回 Excel 列 + 原公式

不做：对账（那是 verifier 的事）。
"""
from __future__ import annotations

import importlib.util
import json
import re
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd


class CalcError(Exception):
    """把运行期的 Python 异常映射回 Excel 列 + 原公式。"""

    def __init__(self, message: str, excel_col: str = "", column_name: str = "",
                 original_formula: str = "", primary_key: Any = None, tb: str = ""):
        super().__init__(message)
        self.excel_col = excel_col
        self.column_name = column_name
        self.original_formula = original_formula
        self.primary_key = primary_key
        self.tb = tb

    def __str__(self) -> str:
        loc = f" [{self.excel_col} {self.column_name}]" if self.excel_col else ""
        return (f"{super().__str__()}{loc}"
                + (f"\n    原公式：{self.original_formula}" if self.original_formula else "")
                + (f"\n    主键  ：{self.primary_key}" if self.primary_key is not None else ""))


def load_module(module_path: str | Path):
    p = Path(module_path).resolve()
    spec = importlib.util.spec_from_file_location(f"e2p_generated_{abs(hash(str(p))) % 10**8}", p)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块：{p}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


# ---------------------------------------------------------------- 异常映射


_LINE_RE = re.compile(r"^\s*(\d+)\s+(.*)$")


def map_error(module_path: str | Path, exc: BaseException) -> Tuple[str, str, str, int]:
    """从 traceback 里定位到生成代码的那一行 -> (列字母, 列名, 原公式, 行号)"""
    lines = Path(module_path).read_text(encoding="utf-8").splitlines()
    src_lines: Dict[int, str] = {}

    tb = exc.__traceback__
    while tb is not None:
        fn = tb.tb_frame.f_code.co_filename
        if Path(fn).resolve() == Path(module_path).resolve():
            src_lines[tb.tb_lineno] = lines[tb.tb_lineno - 1] if 0 < tb.tb_lineno <= len(lines) else ""
        tb = tb.tb_next

    if not src_lines:
        return "", "", "", 0

    ln = max(src_lines)
    src = src_lines[ln]

    letter, name, formula = "", "", ""
    m = re.search(r"df\[\s*'([^']+)'\s*\]\s*=", src)
    if m:
        name = m.group(1)
    # 往上找 `# B2: =...` 注释
    for i in range(ln - 1, max(0, ln - 8), -1):
        c = lines[i - 1]
        cm = re.match(r"\s*#\s*([A-Z]{1,3})(\d+):\s*(=.*)$", c)
        if cm:
            letter, formula = cm.group(1), cm.group(3)
            break
    if not name and letter:
        name = ""
    return letter, name, formula, ln


# ---------------------------------------------------------------- 进度


@dataclass
class Progress:
    path: str
    done: List[str] = field(default_factory=list)
    failed: List[Dict[str, Any]] = field(default_factory=list)

    @staticmethod
    def load(path: str | Path) -> "Progress":
        p = Path(path)
        if p.exists():
            try:
                d = json.loads(p.read_text(encoding="utf-8"))
                return Progress(path=str(p), done=list(d.get("done", [])),
                                failed=list(d.get("failed", [])))
            except Exception:
                pass
        return Progress(path=str(p))

    def save(self) -> None:
        Path(self.path).write_text(
            json.dumps({"done": self.done, "failed": self.failed}, ensure_ascii=False, indent=1),
            encoding="utf-8")


# ---------------------------------------------------------------- 运行


def run_single(module_path: str | Path, primary_value: Any,
               file_paths: Optional[Dict[str, str]] = None,
               inputs: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """单主键调试。"""
    mod = load_module(module_path)
    try:
        return mod.calc([primary_value], file_paths=file_paths, inputs=inputs)
    except Exception as e:  # noqa: BLE001
        letter, name, formula, _ = map_error(module_path, e)
        raise CalcError(str(e), letter, name, formula, primary_value,
                        tb=traceback.format_exc()) from e


def run_batch(module_path: str | Path,
              primary_values: Sequence[Any],
              file_paths: Optional[Dict[str, str]] = None,
              inputs: Optional[Dict[str, Any]] = None,
              progress_file: Optional[str | Path] = None,
              resume: bool = False,
              chunk_size: int = 200,
              batch_log_interval: int = 100,
              log: Optional[Callable[[str], None]] = None) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    """
    批量执行。
    · 分块向量化（不是逐条循环）—— 每个 chunk 调一次 calc，避免重复读表
    · progress_file + resume：断点续传
    · 单条失败不中断，记录失败清单
    """
    log = log or (lambda m: print(m, file=sys.stderr))
    mod = load_module(module_path)

    prog = Progress.load(progress_file) if (progress_file and resume) else (
        Progress(path=str(progress_file)) if progress_file else None)

    done_set = set(map(str, prog.done)) if prog else set()
    todo = [k for k in primary_values if str(k) not in done_set]
    if done_set:
        log(f"[resume] 已完成 {len(done_set)} 条，剩余 {len(todo)} 条")

    frames: List[pd.DataFrame] = []
    failed: List[Dict[str, Any]] = list(prog.failed) if prog else []

    for i in range(0, len(todo), chunk_size):
        chunk = list(todo[i:i + chunk_size])
        try:
            df = mod.calc(chunk, file_paths=file_paths, inputs=inputs)
            frames.append(df)
            if prog:
                prog.done.extend(map(str, chunk))
                prog.save()
        except Exception as e:  # noqa: BLE001
            # 整块失败 -> 拆成单条，定位到具体是哪条 / 哪列
            log(f"[warn] 分块失败（{len(chunk)} 条），降级为单条重试：{e}")
            for k in chunk:
                try:
                    df = mod.calc([k], file_paths=file_paths, inputs=inputs)
                    frames.append(df)
                    if prog:
                        prog.done.append(str(k))
                except Exception as e2:  # noqa: BLE001
                    letter, name, formula, ln = map_error(module_path, e2)
                    failed.append({"key": str(k), "column": f"{letter} {name}".strip(),
                                   "formula": formula, "line": ln, "error": str(e2)})
                    log(f"[fail] {k} -> {letter} {name}: {e2}")
        if prog:
            prog.save()

        if batch_log_interval and (i + chunk_size) % batch_log_interval < chunk_size:
            log(f"[progress] {min(i + chunk_size, len(todo))}/{len(todo)}")

    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return out, failed
