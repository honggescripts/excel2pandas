# ========================================================================
# 自动生成 by excel2pandas 0.1.0
# 源文件   : examples\demo_template.xlsx
# Sheet    : 定价测算
# 表头行   : 1    样板行: 2
# 生成时间 : 2026-09-20 15:59:22
#
# 列分类：
#   数据列（输入参数）:   8 列
#   成功列（已转换）  :  11 列
#   失败列（_fail）   :   0 列
#
# 魔法值常量 : 4 个（已抽到 CONFIG）
# 分层冲突   : 0 处 → 执行顺序由全局拓扑排序决定
# 主键列     : A 商品编号
# ========================================================================

import numpy as np
import pandas as pd
import numpy_financial as nf  # noqa: F401  （_runtime 里的财务函数用）


# 源文件路径：换数据源只改这里
FILE_PATHS = {
    "main": r"examples\demo_template.xlsx",
}


# ========================================================================
# 配置常量（业务规则集中在这里 —— 改规则只改这一段）
# ========================================================================
CONFIG = {
    "包装费率": 0.12,
    "仓储费率": 0.05,
    "建议售价加成": 1.2,
    "达标毛利率": 0.3,
}


# ---- 模块常量 ----
_TARGET_SHEET = "定价测算"
_HEADER_ROW = 1
_TARGET_ROW = 2
_PK_NAME = "商品编号"

# 模板里的手填输入列（批量跑多条记录时由 inputs 覆盖 / 不传则用这里的默认值）
_DEFAULTS = {
    "商品名称": '示例商品',
    "类目": '家居',
    "平台": '平台A',
    "成本价": 36,
    "采购数量": 100,
    "推广费": 8,
    "参考售价": 99,
}

_INPUT_COLS = [
    "商品名称",
    "类目",
    "平台",
    "成本价",
    "采购数量",
    "推广费",
    "参考售价",
]

# 目标表读取位置（0-based 列序号 -> 列名），用于把模板里的手填值 / 失败列兜底值捞回来
_TARGET_COLS = ['商品编号', '商品名称', '类目', '平台', '成本价', '采购数量', '推广费', '参考售价']
_TARGET_POS  = [0, 1, 2, 3, 4, 5, 9, 11]
_FAIL_COLS = []


# ========================================================================
# 运行时辅助库（自包含，无需安装 excel2pandas）
# ========================================================================
import datetime as _dt
import re as _re

import numpy as np
import numpy_financial as _nf
import pandas as pd

# ---- 由 calc() 在运行期注入的模块级状态 ----
_FILE_PATHS = {}
_HEADER_ROW = 1
_TARGET_ROW = 2
_NROW = 1
_SHEET_CACHE = {}
_XLSX = None      # 复用同一个工作簿句柄 —— 否则每读一个 sheet 都要重开整个工作簿

# Excel 错误值字符串：读进来的数据里可能混着这些，统一清洗成 NaN，
# 否则它们会以 str 形式漏进数值运算，直接炸掉。
_ERROR_STRINGS = ["#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!"]


def _clean(s):
    """把 Excel 错误值字符串清洗成 NaN。"""
    if isinstance(s, pd.Series):
        if s.dtype.kind in "OSU":
            return s.replace(_ERROR_STRINGS, np.nan)
        return s
    if isinstance(s, np.ndarray) and s.dtype.kind in "OSU":
        return pd.Series(s).replace(_ERROR_STRINGS, np.nan).to_numpy()
    return s



# ================================================================ 基础

def _unbox(x):
    """Series / Index -> ndarray；其余原样。"""
    if isinstance(x, (pd.Series, pd.Index)):
        return x.to_numpy()
    return np.asarray(x)


def _fit(x, n=None):
    """把标量广播到长度 n。"""
    n = _NROW if n is None else n
    a = _unbox(x)
    if a.ndim == 0:
        return np.full(n, a.item(), dtype=object if a.dtype == object else a.dtype)
    return a


def _obj(x):
    a = _unbox(x)
    if a.ndim == 0:
        return np.full(_NROW, a.item() if a.dtype != object else None, dtype=object)
    return a.astype(object)


def _num(x):
    a = _unbox(x)
    if a.dtype.kind in "iufc":
        return a.astype(float)
    if a.dtype.kind in "OSU":
        a = pd.to_numeric(pd.Series(a), errors="coerce").to_numpy()
    return a.astype(float)


def _n(x):
    """Excel 的隐式数值转换：'12.5%' -> 0.125、'1,280' -> 1280。

    只在「缓存值本身就是数字样式文本」的列上由生成器自动加壳，不滥用。
    """
    a = _unbox(x)
    if a.dtype.kind in "iufc":
        return a
    s = pd.Series(np.atleast_1d(a)).astype(object).astype(str).str.strip()
    pct = s.str.endswith("%").to_numpy()
    base = s.str.rstrip("%").str.replace(",", "", regex=False)
    num = pd.to_numeric(base, errors="coerce").to_numpy(dtype=float)
    return np.where(pct, num / 100.0, num)


# ================================================================ 数据源

def _book():
    """复用同一个 ExcelFile 句柄（性能关键：大工作簿重开一次要好几秒）。"""
    global _XLSX
    if _XLSX is None:
        _XLSX = pd.ExcelFile(_FILE_PATHS["main"])
    return _XLSX


def _parse(sheet_name, **kw):
    d = _book().parse(sheet_name=sheet_name, **kw)
    for c in d.columns:
        if d[c].dtype.kind in "OSU":
            d[c] = d[c].replace(_ERROR_STRINGS, np.nan)
    return d


def _sheet(name):
    """按需读取并缓存一张外部 sheet（同工作簿）。"""
    if name not in _SHEET_CACHE:
        _SHEET_CACHE[name] = _parse(name, header=_HEADER_ROW - 1)
    return _SHEET_CACHE[name]


def _pick(src, letter):
    """按 Excel 列字母取列（位置寻址，免疫表头重名/空表头）。"""
    from openpyxl.utils import column_index_from_string

    return src.iloc[:, column_index_from_string(letter) - 1]


# ================================================================ 查找 / 聚合

def _lookup(src, key, val, keys, not_found=None):
    """XLOOKUP / VLOOKUP / INDEX+MATCH 的统一实现（首次命中）。"""
    k = _pick(src, key)
    v = _pick(src, val)
    m = pd.DataFrame({"__k": k.to_numpy(), "__v": v.to_numpy()})
    m = m[m["__k"].notna()]
    m = m.drop_duplicates(subset="__k", keep="first")
    s = pd.Series(m["__v"].to_numpy(), index=m["__k"].to_numpy())
    if s.dtype.kind in "OSU":
        s = s.replace(_ERROR_STRINGS, np.nan)
    K = pd.Series(_unbox(keys))
    if K.dtype.kind in "OSU":
        K = K.replace(_ERROR_STRINGS, np.nan)
    K = K.to_numpy()
    if K.ndim == 0:
        K = np.full(_NROW, K.item(), dtype=object)
    out = pd.Series(K).map(s)
    if not_found is not None and not (not_found is None):
        out = out.fillna(not_found)
    return out.to_numpy()


def _aggif(src, val, pairs, how="sum"):
    """SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS 的统一实现。

    注意 Excel 语义差异（踩过坑）：
      · SUMIFS / COUNTIFS / MAXIFS / MINIFS 无匹配时返回 **0**（不是空）
      · AVERAGEIFS 无匹配时返回 #DIV/0!（用 NaN 表示）
    这与 XLOOKUP/VLOOKUP 无匹配返回 #N/A 的行为不同。
    """
    dfm = pd.DataFrame()
    keycols = []
    for i, (kl, _kv) in enumerate(pairs):
        col = f"__k{i}"
        dfm[col] = np.asarray(_unbox(_pick(src, kl)), dtype=object)
        keycols.append(col)

    if how == "size":
        res = dfm.dropna(subset=keycols).groupby(keycols, dropna=False).size()
    else:
        dfm["__v"] = _num(_pick(src, val))
        res = dfm.dropna(subset=keycols).groupby(keycols, dropna=False)["__v"].agg(how)

    K = [np.asarray(_fit(_unbox(kv)), dtype=object) for _kl, kv in pairs]
    if len(K) == 1:
        idx = pd.Index(K[0])
    else:
        idx = pd.MultiIndex.from_arrays(K)
    try:
        out = res.reindex(idx).to_numpy(dtype=float)
    except (TypeError, KeyError, ValueError):
        out = np.full(_NROW, np.nan)

    if len(out) != _NROW:
        out = np.full(_NROW, np.nan)
    if how in ("sum", "size", "max", "min"):
        # Excel：无匹配返回 0
        out = np.where(pd.isna(out), 0.0, out)
    return out


# ================================================================ 逻辑

def _is_arr(x):
    return isinstance(x, (np.ndarray, pd.Series, pd.Index, list, tuple))


def _scalar_eq(x, y):
    if x is None or y is None:
        return False
    for v in (x, y):
        if isinstance(v, float) and v != v:
            return False
        if isinstance(v, (np.floating,)) and np.isnan(v):
            return False
    if isinstance(x, (bool, np.bool_)) and isinstance(y, (bool, np.bool_)):
        return bool(x) == bool(y)
    if isinstance(x, str) and isinstance(y, str):
        return x.strip() == y.strip()
    try:
        return float(x) == float(y)
    except (TypeError, ValueError):
        return str(x) == str(y)


def _eq(a, b):
    """Excel 的 `=`（含类型宽松比较）。"""
    if _is_arr(a) or _is_arr(b):
        A = _obj(a) if _is_arr(a) else _fit(a)
        B = _obj(b) if _is_arr(b) else _fit(b)
        A, B = np.broadcast_arrays(A, B)
        return np.array([_scalar_eq(x, y) for x, y in zip(A.ravel(), B.ravel())],
                        dtype=bool).reshape(A.shape)
    return _scalar_eq(a, b)


def _ne(a, b):
    return ~_eq(a, b)


def _bool(x):
    if _is_arr(x):
        return _unbox(x).astype(bool)
    return bool(x)


def _if(cond, a, b):
    """IF：标量判断短路，数组判断逐元素。"""
    if not _is_arr(cond):
        return a if bool(cond) else b
    C = _bool(cond)
    n = C.shape[0] if C.ndim else 1
    A = _fit(a, n)
    B = _fit(b, n)
    if A.dtype == object or B.dtype == object or A.dtype.kind in "US" or B.dtype.kind in "US":
        out = np.where(C, A.astype(object), B.astype(object)).astype(object)
    else:
        out = np.where(C, A, B)
    return out


def _and(items):
    if any(_is_arr(x) for x in items):
        n = _NROW
        out = np.ones(n, dtype=bool)
        for x in items:
            out &= np.asarray(_bool(_fit(x, n)), dtype=bool)
        return out
    return all(bool(x) for x in items)


def _or(items):
    if any(_is_arr(x) for x in items):
        n = _NROW
        out = np.zeros(n, dtype=bool)
        for x in items:
            out |= np.asarray(_bool(_fit(x, n)), dtype=bool)
        return out
    return any(bool(x) for x in items)


def _iferror(a, b):
    try:
        if _is_arr(a):
            arr = _unbox(a)
            if arr.dtype.kind == "f":
                bad = ~np.isfinite(arr)
                if bad.any():
                    return np.where(bad, np.asarray(_fit(b), dtype=float), arr)
            return arr
        if a is None or (isinstance(a, float) and a != a):
            return b
        return a
    except Exception:
        return b


# ================================================================ 数学

def _div(a, b):
    A = _num(a)
    B = _num(b)
    with np.errstate(divide="ignore", invalid="ignore"):
        if A.ndim == 0 and B.ndim == 0:
            return np.nan if (B == 0 or B != B) else float(A / B)
        out = A / B
    out = np.where((B == 0) | (B != B), np.nan, out)
    return out


def _round(x, n=0):
    """Excel ROUND：四舍五入「远离零」，与 numpy 的银行家舍入不同。"""
    V = _num(x)
    f = 10.0 ** float(_unbox(n) if np.ndim(n) == 0 else 0)
    out = np.sign(V) * np.floor(np.abs(V) * f + 0.5) / f
    return out.item() if out.ndim == 0 else out


def _roundup(x, n=0):
    V = _num(x)
    f = 10.0 ** float(n)
    out = np.sign(V) * np.ceil(np.abs(V) * f) / f
    return out.item() if out.ndim == 0 else out


def _rounddown(x, n=0):
    V = _num(x)
    f = 10.0 ** float(n)
    out = np.sign(V) * np.floor(np.abs(V) * f) / f
    return out.item() if out.ndim == 0 else out


def _ceiling(x, n):
    V = _num(x)
    N = _num(n)
    if N.ndim == 0 and N == 0:
        return np.zeros_like(V)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.ceil(V / N) * N
    return out.item() if out.ndim == 0 else out


def _floor_(x, n):
    V = _num(x)
    N = _num(n)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.floor(V / N) * N
    return out.item() if out.ndim == 0 else out


def _psum(items):
    arrs = [_num(x) for x in items]
    if all(a.ndim == 0 for a in arrs):
        return float(np.sum(arrs))
    return np.nansum(np.stack(np.broadcast_arrays(*arrs)), axis=0)


def _pprod(items):
    arrs = [_num(x) for x in items]
    if all(a.ndim == 0 for a in arrs):
        return float(np.prod(arrs))
    return np.nanprod(np.stack(np.broadcast_arrays(*arrs)), axis=0)


def _pmax(items):
    arrs = [_num(x) for x in items]
    if all(a.ndim == 0 for a in arrs):
        return float(np.max(arrs))
    return np.nanmax(np.stack(np.broadcast_arrays(*arrs)), axis=0)


def _pmin(items):
    arrs = [_num(x) for x in items]
    if all(a.ndim == 0 for a in arrs):
        return float(np.min(arrs))
    return np.nanmin(np.stack(np.broadcast_arrays(*arrs)), axis=0)


def _map(key, table, default=np.nan):
    """嵌套 IF 查找表 -> 字典映射。"""
    K = _unbox(key)
    if K.ndim == 0:
        K = np.full(_NROW, K.item(), dtype=object)
    s = pd.Series(K).map(table)
    if _is_arr(default):
        d = np.asarray(_fit(_unbox(default)))
        out = s.to_numpy().astype(object)
        mask = pd.isna(out)
        out = np.where(mask, d, out)
        return out
    return s.fillna(default).to_numpy()


# ================================================================ 文本 / 转换

def _xstr(x):
    """Excel 的 `&` 语义：数字不带小数点、日期格式化、None 变空串。"""
    a = np.atleast_1d(_unbox(x))
    out = []
    for v in a:
        if v is None:
            out.append("")
        elif isinstance(v, (bool, np.bool_)):
            out.append("TRUE" if v else "FALSE")
        elif isinstance(v, (float, np.floating)):
            if v != v:
                out.append("")
            elif float(v).is_integer():
                out.append(str(int(v)))
            else:
                out.append(repr(float(v)))
        elif isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)):
            out.append(pd.Timestamp(v).strftime("%Y-%m-%d %H:%M:%S"))
        else:
            out.append(str(v))
    return np.array(out, dtype=object)


def _str(x):
    return pd.Series(_xstr(x))


def _cat(items):
    if len(items) == 1:
        return _xstr(items[0])
    return _cat2(_cat(items[:-1]), items[-1])


def _cat2(a, b):
    A = _xstr(a)
    B = _xstr(b)
    A, B = np.broadcast_arrays(A, B)
    return np.array([f"{x}{y}" for x, y in zip(A.ravel(), B.ravel())], dtype=object).reshape(A.shape)


def _textbefore(x, sep):
    a = _unbox(x)
    if a.dtype.kind == "M":
        s = pd.Series(a).dt.strftime("%Y-%m-%d %H:%M:%S")
    else:
        s = pd.Series(np.atleast_1d(a)).astype(str)
    return s.str.split(_re.escape(str(sep))).str[0].to_numpy()


def _textafter(x, sep):
    a = _unbox(x)
    if a.dtype.kind == "M":
        s = pd.Series(a).dt.strftime("%Y-%m-%d %H:%M:%S")
    else:
        s = pd.Series(np.atleast_1d(a)).astype(str)
    return s.str.split(_re.escape(str(sep))).str[-1].to_numpy()


def _tonum(x):
    a = _unbox(x)
    if a.dtype.kind in "OSU":
        return pd.to_numeric(pd.Series(np.atleast_1d(a)), errors="coerce").to_numpy()
    return a


def _neg2(x):
    """Excel 的 `--x`：文本转数值 / 文本转日期。"""
    a = _unbox(x)
    if a.dtype.kind == "M":
        return a
    if a.dtype.kind in "OSU":
        arr = np.atleast_1d(a)
        nonnull = int(pd.notna(pd.Series(arr)).sum())
        d = pd.to_datetime(pd.Series(arr), errors="coerce")
        if nonnull and int(d.notna().sum()) >= nonnull:
            return d.to_numpy()
        return pd.to_numeric(pd.Series(arr), errors="coerce").to_numpy()
    return a


def _row(df):
    return df.index.to_numpy() + _TARGET_ROW


def _column(df, ref):
    return np.full(len(df), np.nan)


# ================================================================ 财务函数

def _vfin(fn, args):
    """把标量财务函数向量化；无解时返回 NaN（对应 Excel 的 #NUM!）。"""
    arrs = [np.asarray(_num(x), dtype=float) for x in args]
    arrs = np.broadcast_arrays(*arrs)
    shape = arrs[0].shape
    out = np.full(shape if shape else (), np.nan, dtype=float)
    for idx in (np.ndindex(shape) if shape else [()]):
        vals = tuple(float(a[idx]) for a in arrs)
        if any(v != v for v in vals):
            continue
        try:
            r = float(fn(*vals))
        except Exception:
            r = np.nan
        out[idx] = r if np.isfinite(r) else np.nan
    return out.item() if out.ndim == 0 else out


def _rate(nper, pmt, pv, fv=0, typ=0):
    return _vfin(lambda a, b, c, d: _nf.rate(a, b, c, d, when="begin" if typ else "end"),
                 (nper, pmt, pv, fv))


def _fv(rate_, nper, pmt, pv=0, typ=0):
    return _vfin(lambda a, b, c, d: _nf.fv(a, b, c, d, when="begin" if typ else "end"),
                 (rate_, nper, pmt, pv))


def _pv(rate_, nper, pmt, fv=0, typ=0):
    return _vfin(lambda a, b, c, d: _nf.pv(a, b, c, d, when="begin" if typ else "end"),
                 (rate_, nper, pmt, fv))


def _pmt(rate_, nper, pv, fv=0, typ=0):
    return _vfin(lambda a, b, c, d: _nf.pmt(a, b, c, d, when="begin" if typ else "end"),
                 (rate_, nper, pv, fv))


def _nper(rate_, pmt, pv, fv=0, typ=0):
    return _vfin(lambda a, b, c, d: _nf.nper(a, b, c, d, when="begin" if typ else "end"),
                 (rate_, pmt, pv, fv))


def _ipmt(rate_, per, nper, pv, fv=0, typ=0):
    return _vfin(lambda a, b, c, d: float(_nf.ipmt(a, b, c, d, fv=0,
                                                  when="begin" if typ else "end")),
                 (rate_, per, nper, pv))


def _cumipmt(rate_, nper, pv, start, end, typ=0):
    """numpy_financial 没有 cumipmt —— 用 ipmt 逐期累加实现（已实测复现 Excel）。"""
    def one(r, n, p, s, e):
        n_i = int(round(n))
        s_i = int(round(s))
        e_i = int(round(e))
        pers = np.arange(s_i, e_i + 1)
        if len(pers) == 0:
            return 0.0
        return float(np.sum(_nf.ipmt(r, pers, n_i, p, fv=0,
                                     when="begin" if typ else "end")))
    return _vfin(one, (rate_, nper, pv, start, end))



def _load_target(file_paths, primary_values):
    """读回模板 sheet 里这些主键对应的行（手填输入 + 失败列兜底值）。"""
    pos = [p for p in _TARGET_POS if p >= 0]
    if not pos:
        return pd.DataFrame()
    t = _parse(_TARGET_SHEET, header=_HEADER_ROW - 1, usecols=pos)
    t.columns = _TARGET_COLS
    return t[t[_PK_NAME].isin(list(primary_values))]


# ========================================================================
# 主计算函数：primary_values = 商品编号 列表
# ========================================================================
def calc(primary_values, file_paths=None, inputs=None):
    """
    primary_values : 主键列表
    file_paths     : {'main': xlsx 路径}，默认用 FILE_PATHS
    inputs         : 手填参数覆盖，{列名: 标量 或 {主键: 值}}
    """
    global _FILE_PATHS, _NROW, _XLSX
    _paths = dict(file_paths) if file_paths else dict(FILE_PATHS)
    # 只有换了数据源才丢弃工作簿句柄和 sheet 缓存（批量分块时复用）
    if _FILE_PATHS.get('main') != _paths.get('main'):
        _XLSX = None
        _SHEET_CACHE.clear()
    _FILE_PATHS.update(_paths)
    pks = list(primary_values)
    _NROW = len(pks)

    df = pd.DataFrame({'商品编号': pks})

    # ---- 数据层：手填输入参数 ----
    for _c in _INPUT_COLS:
        df[_c] = np.nan
    _tgt = _load_target(_FILE_PATHS, pks)
    if not _tgt.empty:
        _tgt = _tgt.drop_duplicates(subset=[_PK_NAME], keep='first').set_index(_PK_NAME)
        for _c in _INPUT_COLS:
            if _c in _tgt.columns:
                df[_c] = df['商品编号'].map(_tgt[_c])
    for _c in _INPUT_COLS:
        if inputs and _c in inputs:
            _v = inputs[_c]
            df[_c] = df[_PK_NAME].map(_v) if isinstance(_v, dict) else _v
        elif _c in _DEFAULTS:
            df[_c] = df[_c].fillna(_DEFAULTS[_c])

    # ------------------------------------------------------------
    # Layer 1/2/3：按全局拓扑序逐列计算
    # ------------------------------------------------------------

    # ---- L2 ----

    # G2: =VLOOKUP(C2,运费价目!$A:$B,2,0)
    df['__tmp_lk_01'] = _lookup(_sheet('运费价目'), 'A', 'B', df['类目'], not_found=None)
    df['头程运费'] = df['__tmp_lk_01']

    # ---- L3 ----

    # H2: =ROUND(E2*0.12,2)
    df['包装成本'] = _round(((df['成本价']) * (CONFIG['包装费率'])), 2)

    # ---- L2 ----

    # I2: =VLOOKUP(D2,平台费率!$A:$B,2,0)
    df['__tmp_lk_02'] = _lookup(_sheet('平台费率'), 'A', 'B', df['平台'], not_found=None)
    df['佣金率'] = df['__tmp_lk_02']

    # ---- L3 ----

    # K2: =ROUND(H2+G2+E2*0.05,2)
    df['单件成本'] = _round(((((df['包装成本']) + (df['头程运费']))) + (((df['成本价']) * (CONFIG['仓储费率'])))), 2)

    # M2: =ROUND(L2*I2,2)
    df['平台佣金'] = _round(((df['参考售价']) * (df['佣金率'])), 2)

    # N2: =ROUND(L2-M2-K2-J2,2)
    df['毛利'] = _round(((((((df['参考售价']) - (df['平台佣金']))) - (df['单件成本']))) - (df['推广费'])), 2)

    # O2: =ROUND(N2/L2,4)
    df['毛利率'] = _round(_div(df['毛利'], df['参考售价']), 4)

    # P2: =IF(N2>=30,"A",IF(N2>=20,"B",IF(N2>=10,"C","D")))
    df['定价档位'] = _if(((df['毛利']) >= (30)), 'A', _if(((df['毛利']) >= (20)), 'B', _if(((df['毛利']) >= (10)), 'C', 'D')))

    # Q2: =CEILING(L2*1.2,5)
    df['建议售价'] = _ceiling(((df['参考售价']) * (CONFIG['建议售价加成'])), 5)

    # R2: =ROUND(N2/J2,2)
    df['毛利投产比'] = _round(_div(df['毛利'], df['推广费']), 2)

    # S2: =IF(O2>=0.3,"是","否")
    df['是否达标'] = _if(((df['毛利率']) >= (CONFIG['达标毛利率'])), '是', '否')

    # ------------------------------------------------------------
    # 失败列：不生成赋值，保留 Excel 兜底值 + `_fail` 后缀
    # ------------------------------------------------------------
    # （本次生成：无失败列）

    # ---- 清理临时列 + 返回完整业务表 ----
    df = df.drop(columns=[c for c in df.columns if c.startswith('__tmp_')])

    target_cols = [
        '商品编号',
        '商品名称',
        '类目',
        '平台',
        '成本价',
        '采购数量',
        '头程运费',
        '包装成本',
        '佣金率',
        '推广费',
        '单件成本',
        '参考售价',
        '平台佣金',
        '毛利',
        '毛利率',
        '定价档位',
        '建议售价',
        '毛利投产比',
        '是否达标',
        '备注',
    ]
    return df[[c for c in target_cols if c in df.columns]]


if __name__ == "__main__":
    import sys
    _keys = sys.argv[1:] or ['SAMPLE-0001']
    pd.set_option('display.width', 220)
    pd.set_option('display.max_columns', 60)
    _out = calc(_keys)
    print(_out.T)
