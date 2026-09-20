# -*- coding: utf-8 -*-
"""
generator.py —— 编排 + 渲染 + 缓存（原 generator + CLI 编排）

职责：
  1. 串起 reader -> converter
  2. 渲染 generated_code.py（自包含裸代码，内嵌 runtime）
  3. 生成 business_rules.md（CONFIG 文档）
  4. 生成 migration_report.md（委托 reader）
  5. 基于 Excel hash 的缓存，避免重复生成

不做：执行、对账。
"""
from __future__ import annotations

import ast
import hashlib
import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from . import converter as C
from . import reader as R

VERSION = "0.2.0"

# 生成代码的 runtime 嵌入模式
RUNTIME_MODES = ("minimal", "full")


# ---------------------------------------------------------------- 缓存


def excel_fingerprint(path: str | Path, sheet: str, target_row: int,
                      runtime_mode: str = "minimal") -> str:
    p = Path(path)
    st = p.stat()
    h = hashlib.sha256()
    h.update(f"{p.resolve()}|{st.st_size}|{int(st.st_mtime)}|{sheet}|{target_row}"
             f"|{runtime_mode}|{VERSION}".encode())
    return h.hexdigest()[:16]


def _cache_file(out_dir: Path) -> Path:
    return out_dir / ".excel2pandas_cache.json"


def _load_cache(out_dir: Path) -> Dict[str, Any]:
    f = _cache_file(out_dir)
    if f.exists():
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_cache(out_dir: Path, data: Dict[str, Any]) -> None:
    _cache_file(out_dir).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------------------------------------------------------------- 结果


@dataclass
class GenerateReport:
    skipped: bool = False
    out_dir: str = ""
    code_path: str = ""
    rules_path: str = ""
    migration_path: str = ""
    n_columns: int = 0
    n_formula: int = 0
    n_data: int = 0
    n_success: int = 0
    n_fail: int = 0
    n_config: int = 0
    runtime_mode: str = "minimal"
    runtime_helpers: List[str] = field(default_factory=list)
    fingerprints: Dict[str, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    errors: List[Tuple[str, str, str]] = field(default_factory=list)


# ---------------------------------------------------------------- 默认配置命名

# 魔法值 → 可读业务名 的映射表。
#
# ⚠️ 这里**刻意留空**，不是忘了写。
#
#    常量名本身就是业务信息。一个「某某费率」「某某分成」「某某系数」的名字，
#    连同它绑定的数值一起，等于把你的计费口径与商务条款写进代码——而这是一个
#    会被 git 提交、会被别人 clone 的公开文件。所以业务命名一律从库里拿掉，
#    改由使用者在**输出目录的 config_names.json** 里自己填（见 _load_config_names）。
#    未命中的值由 _auto_key() 按中性启发式命名（_rate_* / _fee_* / _c_*），
#    含义由使用者自行补；库代码本身不携带任何业务语义。
#
# 键格式由 _skey(值) 生成：小数点写成下划线，负号写成 "neg"。
#     0.12 → "0_12"      0.05 → "0_05"      -0.5 → "neg0_5"      3600 → "3600"
# ⚠️ 写成 "0.12" 是查不到的 —— 查找侧走 _skey()，键必须同格式。
DEFAULT_CONFIG_NAMES: Dict[str, str] = {}


def _skey(v: float) -> str:
    s = str(int(v)) if float(v).is_integer() else repr(v)
    return s.replace("-", "neg").replace(".", "_")


def _runtime_source() -> str:
    src = (Path(__file__).parent / "_runtime.py").read_text(encoding="utf-8")
    marker = "import datetime as _dt"
    i = src.index(marker)
    return src[i:].rstrip() + "\n"


# ---------------------------------------------------------------- 按需嵌入
#
# 生成代码必须**自包含**：helper 源码一律内联，绝不出现 `from excel2pandas import ...`。
# 在「全量内联」之上再加一层「按需内联」——只嵌正文真正调用到的 helper。
#
# 关键取舍：**不手工维护依赖表**。
#   `DEFAULT_CONFIG_NAMES` 那类「注册表 → helper」的映射表一旦与实现脱节，就会漏嵌，
#   而漏嵌的后果是生成代码在运行时抛 NameError（用户拿到文件才炸）。
#   这里改成从**已渲染的正文**出发做不动点推导：
#     ① 扫正文里出现的 `_xxx` 标识符 → 命中哪个单元就收哪个
#     ② 把该单元源码也纳入扫描面 → 递归补齐它的依赖
#     ③ 直到不再有新增（闭包完备性由构造保证，而不是靠人记得住）
#   收完再做一次「调用了但没找到」的检查，命中就**在生成期报错**，而不是留下隐患。

_RUNTIME_START = "import datetime as _dt"

# 恒定的两个基础包；_dt / _re / _nf 视实际用到的 helper 决定
_RUNTIME_CORE_IMPORTS = ["import numpy as np", "import pandas as pd"]

# 这几个是 import 别名，不是 helper
_RUNTIME_EXT_ALIASES = {"_dt", "_re", "_nf"}


def _strip_noise(src: str) -> str:
    """去掉注释与字符串内容，只留标识符骨架。

    不能直接对源码做正则：注释和 docstring 里出现的 `_xxx`（例如「# 用 _round 处理」）
    会造成误命中，从而多嵌一堆用不到的函数。
    """
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return re.sub(r"#[^\n]*", "", src)
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            node.value = ""
    try:
        return ast.unparse(tree)
    except Exception:   # pragma: no cover - 兜底，不该发生
        return re.sub(r"#[^\n]*", "", src)


def _module_level_names(src: str) -> Set[str]:
    """源码里模块级定义的名字（def / 赋值）。"""
    out: Set[str] = set()
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return out
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            out.add(node.name)
        elif isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name):
                    out.add(t.id)
    return out


def _runtime_parts() -> Tuple[List[str], Dict[str, str], List[str]]:
    """把 _runtime.py 切成 (恒定代码行, {单元名: 源码}, 单元顺序)。

    单元 = 顶层 `def`，或带名字的顶层赋值（模块级状态、常量表）。
    紧邻其上的空行与注释并入该单元 —— 这样分节标题会跟着函数一起被带上。
    """
    src = (Path(__file__).parent / "_runtime.py").read_text(encoding="utf-8")
    src = src[src.index(_RUNTIME_START):]
    lines = src.splitlines(keepends=True)
    tree = ast.parse(src)

    units: Dict[str, str] = {}
    order: List[str] = []
    taken = [False] * len(lines)
    prev_end = 0

    for node in tree.body:
        name: Optional[str] = None
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            name = node.name
        elif (isinstance(node, ast.Assign) and len(node.targets) == 1
              and isinstance(node.targets[0], ast.Name)):
            name = node.targets[0].id

        start, end = node.lineno - 1, node.end_lineno
        if name is None:
            prev_end = end
            continue

        while (start - 1 >= prev_end
               and (not lines[start - 1].strip() or lines[start - 1].lstrip().startswith("#"))):
            start -= 1
        for i in range(start, end):
            taken[i] = True
        units[name] = "".join(lines[start:end]).rstrip() + "\n"
        order.append(name)
        prev_end = end

    always: List[str] = []
    for i, ln in enumerate(lines):
        if taken[i]:
            continue
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        if ln.startswith("import ") or ln.startswith("from "):
            continue          # import 交给 _runtime_imports 按需生成
        always.append(ln.rstrip())

    return always, units, order


def _runtime_imports(scan: str) -> List[str]:
    """按实际用到的 helper 决定 import 哪些包。"""
    out = list(_RUNTIME_CORE_IMPORTS)
    if re.search(r"\b_dt\.", scan):
        out.append("import datetime as _dt")
    if re.search(r"\b_re\.", scan):
        out.append("import re as _re")
    if re.search(r"\b_nf\.", scan):
        out.append("import numpy_financial as _nf")
    return out


def _render_runtime(body: str, mode: str = "minimal"
                    ) -> Tuple[List[str], str, List[str], int]:
    """渲染 runtime 段。

    → (import 行, runtime 源码, 内联的 helper 名列表, 行数)
    """
    if mode == "full":
        text = _runtime_source()
        return [], text, [], text.count("\n") + 1

    always, units, order = _runtime_parts()

    scan = _strip_noise(body)
    chosen: Set[str] = set()
    changed = True
    while changed:
        changed = False
        for name in order:
            if name in chosen:
                continue
            if re.search(r"\b" + re.escape(name) + r"\b", scan):
                chosen.add(name)
                scan += "\n" + _strip_noise(units[name]) + "\n"
                changed = True

    # 兜底：正文调用了、但 runtime 里没有 —— 生成期直接报错，不留 NameError 给用户
    defined = _module_level_names(body)
    missing = sorted({n for n in re.findall(r"\b(_[A-Za-z]\w*)\s*\(", scan)
                      if n not in chosen and n not in defined
                      and n not in _RUNTIME_EXT_ALIASES})
    if missing:
        raise RuntimeError(
            f"按需嵌入失败：正文引用了 _runtime.py 里不存在的辅助函数 {missing}。"
            f"请检查拼写；或改用 generate(..., runtime_mode='full') 全量嵌入。")

    names = sorted(chosen, key=order.index)
    src = "\n".join(_runtime_imports(scan) + always)
    src += "\n\n\n" + "\n\n\n".join(units[n].rstrip() for n in names) + "\n"
    return [], src, names, src.count("\n") + 1


# ---------------------------------------------------------------- 生成


def generate(
    excel_path: str | Path,
    sheet_name: str,
    target_row: int = 2,
    header_row: int = 1,
    out_dir: str | Path = "output",
    force: bool = False,
    scan_external: bool = False,
    config_names: Optional[Dict[str, str]] = None,
    comment_mode: str = "excel",
    runtime_mode: str = "minimal",
) -> GenerateReport:
    if runtime_mode not in RUNTIME_MODES:
        raise ValueError(f"runtime_mode 只能是 {RUNTIME_MODES}，收到 {runtime_mode!r}")
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    fp = excel_fingerprint(excel_path, sheet_name, target_row, runtime_mode)
    cache = _load_cache(out)
    if not force and cache.get("fingerprint") == fp and (out / "generated_code.py").exists():
        return GenerateReport(skipped=True, out_dir=str(out), fingerprints={"excel": fp},
                              code_path=str(out / "generated_code.py"))

    rep = GenerateReport(out_dir=str(out), fingerprints={"excel": fp},
                         runtime_mode=runtime_mode)

    # ---- Step 1-3: reader（读 + 分层 + 评估）----
    model = R.read_model(excel_path, sheet_name, target_row=target_row,
                         header_row=header_row, scan_external=scan_external)
    rep.n_columns = len(model.columns)
    rep.n_formula = len(model.formula_cols)
    rep.n_data = len(model.data_cols)

    # ---- Step 4-5: converter（拓扑 + 逐列转换）----
    opts = C.ConverterOptions(target_row=target_row, header_row=header_row)
    names = dict(DEFAULT_CONFIG_NAMES)
    names.update(config_names or {})
    ctx = C.RenderCtx(model, opts, names)

    order = C.topo_order(model)
    results: Dict[str, C.ConvertResult] = {}
    for letter in order:
        col = model.pick(letter)
        ctx.begin_column()
        res = C.convert_column(col, model, opts, ctx)
        results[letter] = res
        if not res.success:
            rep.errors.append((letter, col.name, res.error or "未知错误"))
        rep.warnings.extend(res.warnings)
    rep.n_success = sum(1 for r in results.values() if r.success)
    rep.n_fail = len(order) - rep.n_success
    rep.n_config = len(ctx.config_values)

    # ---- Step 6: 渲染 ----
    pk_letter = model.primary_key_letter
    pk_name = model.resolve_name(pk_letter) if pk_letter else None
    input_cols = [c.name for c in model.data_cols if c.name != pk_name]
    fail_cols = [c.name for c in model.columns
                 if c.formula and not results.get(c.letter, C.ConvertResult(c.letter, c.name, c.layer)).success]

    code, rt_names = _render_code(model, order, results, ctx, pk_name, pk_letter,
                                  input_cols, fail_cols, comment_mode, runtime_mode)
    rep.runtime_helpers = rt_names
    code_path = out / "generated_code.py"
    code_path.write_text(code, encoding="utf-8")
    rep.code_path = str(code_path)

    rules_path = out / "business_rules.md"
    rules_path.write_text(_render_rules(model, results, ctx, names), encoding="utf-8")
    rep.rules_path = str(rules_path)

    mig_path = out / "migration_report.md"
    R.write_migration_report(model, mig_path)
    rep.migration_path = str(mig_path)

    # ---- Step 7: 缓存 ----
    cache.update({"fingerprint": fp, "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                  "excel": str(Path(excel_path).resolve()), "sheet": sheet_name,
                  "target_row": target_row})
    _save_cache(out, cache)

    return rep


# ---------------------------------------------------------------- 渲染代码


def _render_code(model, order, results, ctx, pk_name, pk_letter, input_cols,
                 fail_cols, comment_mode, runtime_mode="minimal"
                 ) -> Tuple[str, List[str]]:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    L: List[str] = []
    A = L.append

    all_out = [c.name for c in model.columns if c.kind != "empty" or c.header]

    A("# " + "=" * 72)
    A("# 自动生成 by excel2pandas " + VERSION)
    A(f"# 源文件   : {model.path}")
    A(f"# Sheet    : {model.sheet_name}")
    A(f"# 表头行   : {model.header_row}    样板行: {model.target_row}")
    A(f"# 生成时间 : {ts}")
    A("#")
    A("# 列分类：")
    A(f"#   数据列（输入参数）: {len(model.data_cols):>3} 列")
    A(f"#   成功列（已转换）  : {sum(1 for r in results.values() if r.success):>3} 列")
    A(f"#   失败列（_fail）   : {len(fail_cols):>3} 列")
    A("#")
    A(f"# 魔法值常量 : {len(ctx.config_values)} 个（已抽到 CONFIG）")
    A(f"# 分层冲突   : {len(model.conflicts)} 处 → 执行顺序由全局拓扑排序决定")
    A(f"# 主键列     : {pk_letter or '-'} {pk_name or ''}")
    A("# " + "=" * 72)
    hdr_end = len(L)        # runtime 段的信息行，等闭包算完再插进来
    A("# 源文件路径：换数据源只改这里")
    A("FILE_PATHS = {")
    A(f'    "main": r"{model.path}",')
    A("}")
    A("")
    A("")
    A("# " + "=" * 72)
    A("# 配置常量（业务规则集中在这里 —— 改规则只改这一段）")
    A("# " + "=" * 72)
    A("CONFIG = {")
    for k, v in ctx.config_values.items():
        A(f'    "{k}": {v!r},')
    A("}")
    A("")
    A("")
    # ---- 模块常量 ----
    A("# ---- 模块常量 ----")
    A(f'_TARGET_SHEET = "{model.sheet_name}"')
    A(f"_HEADER_ROW = {model.header_row}")
    A(f"_TARGET_ROW = {model.target_row}")
    A(f'_PK_NAME = "{pk_name}"')
    A("")
    A("# 模板里的手填输入列（批量跑多条记录时由 inputs 覆盖 / 不传则用这里的默认值）")
    A("_DEFAULTS = {")
    for c in model.data_cols:
        if c.name == pk_name:
            continue
        v = c.cached
        # 数据列的缓存值也可能是 #N/A 这种错误值，不能当默认值传下去
        if isinstance(v, str) and v.strip().upper() in R.EXCEL_ERRORS:
            A(f'    "{c.name}": np.nan,   # 模板里的原值是 {v}')
            continue
        # 数字样式的文本（'12.5%' / '1,280'）一律还原成数值
        nv = R.parse_numeric_text(v)
        if nv is not None:
            A(f'    "{c.name}": {nv!r},   # 模板里原样是 {v!r}')
            continue
        A(f'    "{c.name}": {v!r},')
    A("}")
    A("")
    A("_INPUT_COLS = [")
    for n in input_cols:
        A(f'    "{n}",')
    A("]")
    A("")
    # 目标表读取位置（0-based，供 pd.read_excel(usecols=...) 用）
    tgt_names = [pk_name] + input_cols + [f"{n}_fail" for n in fail_cols]
    tgt_pos = []
    for n in [pk_name] + input_cols + fail_cols:
        c = model.by_name.get(n)
        tgt_pos.append(c.index - 1 if c else -1)
    A("# 目标表读取位置（0-based 列序号 -> 列名），用于把模板里的手填值 / 失败列兜底值捞回来")
    A("_TARGET_COLS = " + repr(tgt_names))
    A("_TARGET_POS  = " + repr(tgt_pos))
    A("_FAIL_COLS = " + repr(fail_cols))
    A("")
    A("")
    # ---- 正文切到独立列表：runtime 段要等正文渲染完才能裁剪 ----
    B: List[str] = []
    A = B.append

    # ---- 数据层 ----
    A("def _load_target(file_paths, primary_values):")
    A('    """读回模板 sheet 里这些主键对应的行（手填输入 + 失败列兜底值）。"""')
    A("    pos = [p for p in _TARGET_POS if p >= 0]")
    A("    if not pos:")
    A("        return pd.DataFrame()")
    A("    t = _parse(_TARGET_SHEET, header=_HEADER_ROW - 1, usecols=pos)")
    A("    t.columns = _TARGET_COLS")
    A("    return t[t[_PK_NAME].isin(list(primary_values))]")
    A("")
    A("")
    A("# " + "=" * 72)
    A(f"# 主计算函数：primary_values = {pk_name or '主键'} 列表")
    A("# " + "=" * 72)
    A("def calc(primary_values, file_paths=None, inputs=None):")
    A('    """')
    A("    primary_values : 主键列表")
    A("    file_paths     : {'main': xlsx 路径}，默认用 FILE_PATHS")
    A("    inputs         : 手填参数覆盖，{列名: 标量 或 {主键: 值}}")
    A('    """')
    A("    global _FILE_PATHS, _NROW, _XLSX")
    A("    _paths = dict(file_paths) if file_paths else dict(FILE_PATHS)")
    A("    # 只有换了数据源才丢弃工作簿句柄和 sheet 缓存（批量分块时复用）")
    A("    if _FILE_PATHS.get('main') != _paths.get('main'):")
    A("        _XLSX = None")
    A("        _SHEET_CACHE.clear()")
    A("    _FILE_PATHS.update(_paths)")
    A("    pks = list(primary_values)")
    A("    _NROW = len(pks)")
    A("")
    A(f"    df = pd.DataFrame({{{pk_name!r}: pks}})")
    A("")
    A("    # ---- 数据层：手填输入参数 ----")
    A("    for _c in _INPUT_COLS:")
    A("        df[_c] = np.nan")
    A("    _tgt = _load_target(_FILE_PATHS, pks)")
    A("    if not _tgt.empty:")
    A("        _tgt = _tgt.drop_duplicates(subset=[_PK_NAME], keep='first').set_index(_PK_NAME)")
    A("        for _c in _INPUT_COLS:")
    A("            if _c in _tgt.columns:")
    A(f"                df[_c] = df[{pk_name!r}].map(_tgt[_c])")
    A("    for _c in _INPUT_COLS:")
    A("        if inputs and _c in inputs:")
    A("            _v = inputs[_c]")
    A("            df[_c] = df[_PK_NAME].map(_v) if isinstance(_v, dict) else _v")
    A("        elif _c in _DEFAULTS:")
    A("            df[_c] = df[_c].fillna(_DEFAULTS[_c])")
    A("")
    A("    # " + "-" * 60)
    A(f"    # Layer 1/2/3：按全局拓扑序逐列计算")
    A("    # " + "-" * 60)

    last_layer = None
    n_done = 0
    for letter in order:
        col = model.pick(letter)
        res = results[letter]
        if comment_mode != "none" and col.layer != last_layer:
            A("")
            A(f"    # ---- {col.layer} ----")
            last_layer = col.layer
        A("")
        if res.success:
            if comment_mode in ("excel", "both", "names", "letters"):
                A(f"    # {letter}{model.target_row}: {col.formula}")
            for st in res.pre_stmts:
                A("    " + st)
            A(f"    df[{col.name!r}] = {res.expr}")
            n_done += 1
            # 逐列 insert 会让 DataFrame 高度碎片化（100+ 列时触发 PerformanceWarning
            # 并显著拖慢后续运算），定期去碎片
            if n_done % 30 == 0:
                A("")
                A("    df = df.copy()   # 去碎片：逐列 insert 会降低后续运算性能")
                A("")
        else:
            A(f"    # ⚠ 转换失败: {res.error}")
            A(f"    #    原公式：{col.formula}")
            A(f"    #    此列保留 Excel 缓存值，名称加 `_fail` 后缀")

    A("")
    A("    # " + "-" * 60)
    A("    # 失败列：不生成赋值，保留 Excel 兜底值 + `_fail` 后缀")
    A("    # " + "-" * 60)
    if fail_cols:
        A("    for _c in _FAIL_COLS:")
        A("        df[_c + '_fail'] = np.nan")
        A("    if not _tgt.empty:")
        A("        for _c in _FAIL_COLS:")
        A("            if _c in _tgt.columns:")
        A(f"                df[_c + '_fail'] = df[{pk_name!r}].map(_tgt[_c])")
    else:
        A("    # （本次生成：无失败列）")

    A("")
    A("    # ---- 清理临时列 + 返回完整业务表 ----")
    A("    df = df.drop(columns=[c for c in df.columns if c.startswith('__tmp_')])")
    A("")
    A("    target_cols = [")
    for n in all_out:
        if n in fail_cols:
            n = f"{n}_fail"
        A(f"        {n!r},")
    A("    ]")
    A("    return df[[c for c in target_cols if c in df.columns]]")
    A("")
    A("")
    A('if __name__ == "__main__":')
    A("    import sys")
    A("    _keys = sys.argv[1:] or ['SAMPLE-0001']")
    A("    pd.set_option('display.width', 220)")
    A("    pd.set_option('display.max_columns', 60)")
    A("    _out = calc(_keys)")
    A("    print(_out.T)")
    A("")

    # ---- runtime 段：按正文实际用到的 helper 裁剪后内联 ----
    rt_imports, rt_src, rt_names, rt_lines = _render_runtime("\n".join(B), runtime_mode)

    info: List[str] = []
    if runtime_mode == "full":
        info.append("# 渲染模式   : full —— 全量内联 runtime（最保险，文件最长）")
    else:
        info.append(f"# 渲染模式   : minimal —— 只内联正文用到的 {len(rt_names)} 个 helper")
        info.append('#              要全量嵌入：generate(..., runtime_mode="full")')
    info.append("# 内联 helper: " + ("、".join(rt_names) if rt_names else "（全量，不裁剪）"))
    info.append(f"# runtime 行数: {rt_lines}")
    deps = "pip install numpy pandas"
    if runtime_mode == "full" or any("numpy_financial" in x for x in rt_imports):
        deps += " numpy_financial"
    info.append(f"# 外部依赖   : {deps}")
    info.append("#              本文件自包含，**不需要**安装 excel2pandas")

    rt_block = ["", "",
                "# " + "=" * 72,
                f"# 运行时辅助库（内联，自包含 —— {rt_lines} 行）",
                "# " + "=" * 72]
    rt_block += rt_src.rstrip("\n").split("\n")

    final = (L[:hdr_end] + info + [""]
             + rt_imports + ["", ""]
             + L[hdr_end:]
             + rt_block + [""] + B)
    return "\n".join(final), rt_names


# ---------------------------------------------------------------- 渲染 CONFIG 文档


def _render_rules(model, results, ctx, names) -> str:
    # CONFIG key -> 使用的列
    usage: Dict[str, List[str]] = {}
    for letter, r in results.items():
        if not r.success:
            continue
        col = model.pick(letter)
        for k in r.config_keys:
            usage.setdefault(k, []).append(f"{letter} {col.name}")

    # 值 -> 用了同一 key 的列（用于发现语义碰撞）
    by_value: Dict[float, List[str]] = {}
    for k, v in ctx.config_values.items():
        by_value.setdefault(v, []).append(k)

    L: List[str] = []
    A = L.append
    A("# 业务规则常量（CONFIG）说明")
    A("")
    A(f"- 源文件：`{model.path}`")
    A(f"- Sheet：`{model.sheet_name}`　样板行：{model.target_row}")
    A(f"- 共 {len(ctx.config_values)} 个常量")
    A("")
    A("> 这些常量是从样板行公式里抽出来的「魔法值」。改业务规则只改 `generated_code.py` 的 `CONFIG` 字典。")
    A("")
    A("---")
    A("")
    A("## 常量清单")
    A("")
    A("| CONFIG 键 | 值 | 原始字面量 | 被哪些列使用 |")
    A("|---|---|---|---|")
    for k, v in ctx.config_values.items():
        raw = str(int(v)) if float(v).is_integer() else repr(v)
        used = usage.get(k, [])
        if not used:
            used = ["（未被任何列的顶层表达式直接使用，可能只出现在条件分支里）"]
        elif len(used) > 6:
            used = used[:6] + [f"…共 {len(used)} 列"]
        A(f"| `{k}` | `{v}` | `{raw}` | {'<br>'.join(used)} |")
    A("")
    A("---")
    A("")
    A("## 需要人工确认的点")
    A("")
    n = 0
    # 1) 多列复用的常量 —— 只给汇总，不逐条列（逐条会在清单表里重复）
    shared = {k: cols for k, cols in usage.items() if len(cols) >= 2}
    if shared:
        n += 1
        top = ", ".join(f"`{k}`({len(v)})" for k, v in
                        sorted(shared.items(), key=lambda x: -len(x[1]))[:6])
        A(f"- 🔎 **{len(shared)}** 个常量被多列复用：{top}")
        A(f"  —— 请对照上面的「常量清单」逐条核对：**若这些列的业务含义不同**，说明同一个数值"
          f"被合并成了一个常量（同一个数字在不同列可能代表完全不同的口径），需要拆开分别命名。")
    # 2) 隐含查找表：同一列引用了一组同前缀常量
    PREFIX_HINT = {"_rate_": "费率/系数表", "_fee_": "金额档位表", "_c_": "数值档位表"}
    for letter, r in results.items():
        if not r.success:
            continue
        groups: Dict[str, List[str]] = {}
        for k in r.config_keys:
            for pre in PREFIX_HINT:
                if k.startswith(pre):
                    groups.setdefault(pre, []).append(k)
        for pre, ks in groups.items():
            if len(ks) >= 5:
                n += 1
                col = model.pick(letter)
                A(f"- 📊 `{letter} {col.name}` 引用了 **{len(ks)}** 个 `{pre}*` 常量"
                  f"（{PREFIX_HINT[pre]}）—— 建议在 CONFIG 里改写成一个字典，")
                A(f"  让代码从 {len(ks)} 行平铺常量 + 深层嵌套 IF 变成一次表查找。")
                break
    # 3) 自动纠正
    for letter, r in results.items():
        for w in r.warnings:
            n += 1
            A(f"- ⚠️ `{letter}` {model.pick(letter).name}：{w}")
    # 4) 未命名常量（还是启发式 key）
    heur = [k for k in ctx.config_values if k.startswith(("_rate_", "_fee_", "_c_"))]
    if heur:
        n += 1
        A(f"- 📝 还有 {len(heur)} 个常量用的是启发式命名（`{heur[0]}` 这种），"
          f"建议在 `config_names.json` 里补上业务名后重新生成。")
    if not n:
        A("- ✅ 无")
    A("")
    A("---")
    A("")
    A("## 怎么改规则")
    A("")
    A("1. 直接改 `generated_code.py` 里 `CONFIG` 字典的值（最快）。")
    A("2. 想保留改动：在输出目录放一个 `config_names.json`（`{\"0_12\": \"包装费率\"}` 这种，键用下划线代替小数点），重新生成即可复用命名。")
    A("3. **不要**回头改 Excel 再生成 —— 生成器的输入是 Excel 模板，输出是代码；改代码就是改代码。")
    A("")
    return "\n".join(L)
