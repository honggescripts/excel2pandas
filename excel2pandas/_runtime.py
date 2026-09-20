# -*- coding: utf-8 -*-
"""
_runtime.py —— 生成代码用的运行时辅助库（模板资源，不是包模块）

generator 会把本文件的源码原样嵌入生成代码的尾部，
使 generated_code.py 成为**自包含的裸代码**（不含对本包的依赖），
人可以读、可以改、可以进版本库。

约定：所有辅助函数对「标量 / numpy 数组」双态兼容，统一返回 numpy 数组或标量。
"""
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
    """Excel CEILING：向**远离零**方向取整到倍数。

    ⚠️ 不能写成 np.ceil(V/N)*N —— 那个是"向正无穷"，对负数是错的：
       Excel CEILING(-2.1, 1) = -3（远离零），np.ceil 给的是 -2。
    """
    V = _num(x)
    N = _num(n)
    if N.ndim == 0 and N == 0:
        return np.zeros_like(V)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.sign(V) * np.ceil(np.abs(V / N)) * np.abs(N)
    return out.item() if out.ndim == 0 else out


def _floor_(x, n):
    """Excel FLOOR：向**零**方向取整到倍数。

    同理不能写 np.floor(V/N)*N —— 对负数得到的是"向负无穷"：
       Excel FLOOR(-2.1, 1) = -2（向零），np.floor 给的是 -3。
    """
    V = _num(x)
    N = _num(n)
    if N.ndim == 0 and N == 0:
        return np.zeros_like(V)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.sign(V) * np.floor(np.abs(V / N)) * np.abs(N)
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


# ================================================================ 广播工具
#
# 下面这些只服务于「按行向量化」的新增函数：把标量参数摊到与行数等长，
# 保证每个函数都返回长度 = _NROW 的数组（与其余 helper 的输出约定一致）。


def _rep1(a, n):
    """长度 1 的数组复制成 n 份；其余原样。"""
    a = np.atleast_1d(a)
    if a.shape[0] != n and a.shape[0] == 1:
        return np.repeat(a, n)
    return a


def _bn(x, n=None):
    """数值参数 -> 长度 n 的 float 数组。"""
    n = _NROW if n is None else n
    a = _num(x)
    if a.ndim == 0:
        return np.full(n, float(a))
    return _rep1(np.atleast_1d(a).astype(float), n)


def _bvec(x, n=None):
    """布尔参数 -> 长度 n 的 bool 数组。"""
    n = _NROW if n is None else n
    a = _unbox(x)
    if a.ndim == 0:
        return np.full(n, bool(a), dtype=bool)
    return _rep1(np.array([bool(v) for v in np.atleast_1d(a)], dtype=bool), n)


# ================================================================ 日期
#
# 输入可以是 datetime64 / Timestamp / date / 日期文本 / **Excel 序列号（数字）**。
# 返回统一是 datetime64[ns] 数组。
#
# ⚠️ 两个需要知道的边界：
#   · 序列号 < 61（1900-01-01 ~ 1900-02-28）落在 Excel 的「1900 闰年 bug」区间，不保证正确；
#     现代业务日期（>= 1900-03-01）用 1899-12-30 为基准换算，与 Excel 一致。
#   · 返回的是 datetime64，不是 Excel 的序列号。对账时 Excel 侧单元格需为日期格式，
#     否则会「值相等但类型不同」而判为不一致。

# Excel 序列号基准：2024-01-01 = 45292
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def _d1(v):
    """单个值 -> datetime64[ns]；无法解析 -> NaT。数字按 Excel 序列号解释。"""
    if v is None or isinstance(v, (bool, np.bool_)):
        return np.datetime64("NaT")
    if isinstance(v, (int, float, np.integer, np.floating)):
        if v != v:
            return np.datetime64("NaT")
        return np.datetime64(_EXCEL_EPOCH + pd.Timedelta(days=float(v)), "ns")
    t = pd.to_datetime(v, errors="coerce")
    return np.datetime64("NaT") if pd.isna(t) else np.datetime64(t, "ns")


def _dvec(x, n=None):
    """日期参数 -> 长度 n 的 datetime64[ns] 数组（标量摊到 n）。"""
    n = _NROW if n is None else n
    a = _unbox(x)
    if a.ndim == 0:
        return np.full(n, _d1(a.item()), dtype="datetime64[ns]")
    a = np.atleast_1d(a)
    if a.dtype.kind == "M":
        return _rep1(a.astype("datetime64[ns]"), n)
    return _rep1(np.array([_d1(v) for v in a], dtype="datetime64[ns]"), n)


def _dct2(a, b):
    """两个日期参数 -> 对齐到同一长度的一对数组 (+ 该长度)。"""
    A = _dvec(a)
    B = _dvec(b)
    n = max(len(A), len(B))
    return _rep1(A, n), _rep1(B, n), n


def _lastday(y, m):
    return int((pd.Timestamp(y, m, 1) + pd.offsets.MonthEnd(0)).day)


def _date(y, m, d):
    """DATE：允许月/日越界滚动（Excel 语义：DATE(2026,13,1) = 2027-01-01）。"""
    Y = _bn(y)
    M = _bn(m)
    D = _bn(d)
    n = max(len(Y), len(M), len(D))
    Y, M, D = _rep1(Y, n), _rep1(M, n), _rep1(D, n)
    out = np.empty(n, dtype="datetime64[ns]")
    for i in range(n):
        yy, mm, dd = Y[i], M[i], D[i]
        if yy != yy or mm != mm or dd != dd:
            out[i] = np.datetime64("NaT")
            continue
        y2 = int(yy) + (int(mm) - 1) // 12
        m2 = (int(mm) - 1) % 12 + 1
        if not (1900 <= y2 <= 9999):
            out[i] = np.datetime64("NaT")
            continue
        out[i] = np.datetime64(pd.Timestamp(y2, m2, 1) + pd.Timedelta(days=int(dd) - 1), "ns")
    return out


def _datepart(x, part):
    """YEAR / MONTH / DAY：取日期分量，返回 float 数组（NaT -> NaN）。"""
    s = pd.Series(_dvec(x))
    if part == "year":
        return s.dt.year.to_numpy(dtype=float)
    if part == "month":
        return s.dt.month.to_numpy(dtype=float)
    if part == "day":
        return s.dt.day.to_numpy(dtype=float)
    if part == "weekday":
        # Excel WEEKDAY 默认 type=1：周日=1 … 周六=7
        return (s.dt.weekday.to_numpy(dtype=float) + 2.0) % 7.0 + 1.0
    raise ValueError(f"未知日期分量: {part!r}")


def _eomonth(start, months=0):
    """EOMONTH：偏移 months 个月后的月末。"""
    D = _dvec(start)
    M = _bn(months, len(D))
    out = np.empty(len(D), dtype="datetime64[ns]")
    for i, (d, m) in enumerate(zip(D, M)):
        t = pd.Timestamp(d)
        if pd.isna(t) or m != m:
            out[i] = np.datetime64("NaT")
            continue
        k = int(round(float(m)))
        y = t.year + (t.month - 1 + k) // 12
        mo = (t.month - 1 + k) % 12 + 1
        if not (1900 <= y <= 9999):
            out[i] = np.datetime64("NaT")
            continue
        out[i] = np.datetime64(pd.Timestamp(y, mo, 1) + pd.offsets.MonthEnd(0), "ns")
    return out


def _edate(start, months=0):
    """EDATE：偏移 months 个月，日号不变；目标月没有该日则取月末。"""
    D = _dvec(start)
    M = _bn(months, len(D))
    out = np.empty(len(D), dtype="datetime64[ns]")
    for i, (d, m) in enumerate(zip(D, M)):
        t = pd.Timestamp(d)
        if pd.isna(t) or m != m:
            out[i] = np.datetime64("NaT")
            continue
        k = int(round(float(m)))
        y = t.year + (t.month - 1 + k) // 12
        mo = (t.month - 1 + k) % 12 + 1
        if not (1900 <= y <= 9999):
            out[i] = np.datetime64("NaT")
            continue
        out[i] = np.datetime64(pd.Timestamp(y, mo, min(t.day, _lastday(y, mo))), "ns")
    return out


def _days360_1(t1, t2, euro=False):
    """单对日期的 30/360 天数。euro=False 走 NASD 美式规则（Excel 默认）。"""
    y1, m1, d1 = t1.year, t1.month, t1.day
    y2, m2, d2 = t2.year, t2.month, t2.day
    if euro:
        if d1 == 31:
            d1 = 30
        if d2 == 31:
            d2 = 30
    else:
        if d1 == _lastday(y1, m1):
            d1 = 30
        if d2 == _lastday(y2, m2):
            if d1 < 30:
                d2, m2 = 1, m2 + 1
                if m2 == 13:
                    m2, y2 = 1, y2 + 1
            else:
                d2 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def _days360(start, end, method=False):
    """DAYS360(start, end, [method])：method=True 用欧式 30/360。"""
    D1, D2, n = _dct2(start, end)
    euro = _bvec(method, n)
    out = np.full(n, np.nan)
    for i in range(n):
        t1, t2 = pd.Timestamp(D1[i]), pd.Timestamp(D2[i])
        if pd.isna(t1) or pd.isna(t2):
            continue
        out[i] = float(_days360_1(t1, t2, bool(euro[i])))
    return out


def _yearfrac(start, end, basis=0):
    """YEARFRAC：basis 0=US 30/360，2=Actual/360，3=Actual/365，4=欧式 30/360。

    ⚠️ basis=1（Actual/Actual）**刻意不实现**：Excel 对它有多条与版本相关的边界规则
    （跨闰年时的分母取值），没有逐例校准就实现会静默算错。碰到就抛错，不猜。
    """
    D1, D2, n = _dct2(start, end)
    B = _bn(basis, n)
    out = np.full(n, np.nan)
    for i in range(n):
        t1, t2 = pd.Timestamp(D1[i]), pd.Timestamp(D2[i])
        b = B[i]
        if pd.isna(t1) or pd.isna(t2) or b != b:
            continue
        k = int(round(float(b)))
        if k == 1:
            raise NotImplementedError(
                "YEARFRAC basis=1（Actual/Actual）未实现 —— 该口径含版本相关的闰年边界规则，"
                "需先与目标 Excel 逐例校准。请改用 basis=0/2/3/4，或手动补上该列。")
        if k == 0:
            out[i] = _days360_1(t1, t2, False) / 360.0
        elif k == 2:
            out[i] = float((t2 - t1).days) / 360.0
        elif k == 3:
            out[i] = float((t2 - t1).days) / 365.0
        elif k == 4:
            out[i] = _days360_1(t1, t2, True) / 360.0
    return out


def _holidays(h):
    """节假日参数 -> set[date]。None / 空 -> 空集。"""
    if h is None:
        return set()
    a = _unbox(h)
    if a.ndim == 0:
        if a.item() is None:
            return set()
        a = np.atleast_1d(a)
    out = set()
    for v in a:
        if v is None:
            continue
        t = pd.Timestamp(_d1(v))
        if not pd.isna(t):
            out.add(t.normalize().date())
    return out


def _networkdays(start, end, holidays=None):
    """NETWORKDAYS：起止区间内的周一~周五天数（含两端），扣除节假日。"""
    D1, D2, n = _dct2(start, end)
    hol = _holidays(holidays)
    out = np.full(n, np.nan)
    for i in range(n):
        t1, t2 = pd.Timestamp(D1[i]), pd.Timestamp(D2[i])
        if pd.isna(t1) or pd.isna(t2):
            continue
        a, b, sign = t1.normalize(), t2.normalize(), 1
        if a > b:
            a, b, sign = b, a, -1
        cnt = 0
        for d in pd.date_range(a, b, freq="D"):
            if d.weekday() < 5 and d.date() not in hol:
                cnt += 1
        out[i] = float(sign * cnt)
    return out


def _workday(start, days, holidays=None):
    """WORKDAY：按工作日偏移（days 可为负）。"""
    D = _dvec(start)
    N = _bn(days, len(D))
    hol = _holidays(holidays)
    out = np.empty(len(D), dtype="datetime64[ns]")
    for i, (d, k) in enumerate(zip(D, N)):
        t = pd.Timestamp(d)
        if pd.isna(t) or k != k:
            out[i] = np.datetime64("NaT")
            continue
        step = 1 if k >= 0 else -1
        rem = abs(int(round(float(k))))
        cur = t.normalize()
        while rem > 0:
            cur = cur + pd.Timedelta(days=step)
            if cur.weekday() < 5 and cur.date() not in hol:
                rem -= 1
        out[i] = np.datetime64(cur, "ns")
    return out


# Excel WEEKNUM 的 type -> 一周的起始日（Python weekday：周一=0 … 周日=6）
# 11~17 分别是 周一~周日；21 = ISO 8601。Excel 支持的类型全部覆盖。
_WEEK_START = {1: 6, 2: 0, 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5, 17: 6, 21: 0}


def _weeknum(date, type=1):
    """WEEKNUM：周数。type 1/2/11~17/21，非法 type 返回 NaN（对应 Excel 的 #NUM!）。"""
    D = _dvec(date)
    T = _bn(type, len(D))
    out = np.full(len(D), np.nan)
    for i, (d, t) in enumerate(zip(D, T)):
        ts = pd.Timestamp(d)
        if pd.isna(ts) or t != t:
            continue
        k = int(round(float(t)))
        if k not in _WEEK_START:
            continue
        if k == 21:
            out[i] = float(ts.isocalendar()[1])
            continue
        ws = _WEEK_START[k]
        off = (pd.Timestamp(ts.year, 1, 1).weekday() - ws) % 7
        out[i] = float((ts.dayofyear + off - 1) // 7 + 1)
    return out


# ================================================================ 数学（补）


def _mround(x, multiple):
    """MROUND：舍入到倍数。规则与 Excel 一致：
      · .5 向远离零方向进位（MROUND(1.5,1)=2、MROUND(-1.5,1)=-2）
      · 结果的符号跟随被舍入的数（MROUND(-10,-3)=-9，不是 +9）
      · 两个参数异号 -> #NUM!（用 NaN 表示）；MROUND(x,0)=0

    加 1e-9 是为了吸收 IEEE754 误差：6.05/0.1 在双精度下是 60.4999…，
    直接 +0.5 会得到 6.0，而 Excel 的十进制口径给 6.1。
    """
    X = _num(x)
    M = _num(multiple)
    X, M = np.broadcast_arrays(X, M)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.sign(X) * np.floor(np.abs(X / M) + 0.5 + 1e-9) * np.abs(M)
    out = np.where(M == 0, 0.0, out)                 # MROUND(x, 0) = 0
    out = np.where(np.isnan(X) | np.isnan(M), np.nan, out)
    out = np.where((X * M) < 0, np.nan, out)         # 异号 -> #NUM!
    return out.item() if out.ndim == 0 else out


# ================================================================ 文本（补）


def _textjoin(delimiter, ignore_empty, items):
    """TEXTJOIN：按行拼接（分隔符与 ignore_empty 也会广播）。"""
    texts = [_xstr(it) for it in items]
    if not texts:
        return np.full(_NROW, "", dtype=object)
    n = max(len(t) for t in texts)
    texts = [_rep1(t, n) for t in texts]
    D = _rep1(_xstr(delimiter), n)
    ie = _bvec(ignore_empty, n)
    out = []
    for i in range(n):
        parts = ["" if t[i] is None else str(t[i]) for t in texts]
        if bool(ie[i]):
            parts = [p for p in parts if p != ""]
        out.append(str(D[i]).join(parts))
    return np.array(out, dtype=object)


def _nth_replace(s, o, v, k):
    """把 s 里第 k 次出现的 o 换成 v（k 从 1 开始；找不到则原样返回）。"""
    if k <= 0:
        return s
    idx = -1
    start = 0
    for _ in range(k):
        idx = s.find(o, start)
        if idx < 0:
            return s
        start = idx + len(o)
    return s[:idx] + v + s[idx + len(o):]


def _substitute(text, old, new, instance=None):
    """SUBSTITUTE：替换全部，或只替换第 instance 次。"""
    T = _xstr(text)
    O = _xstr(old)
    N = _xstr(new)
    n = max(len(T), len(O), len(N))
    T, O, N = _rep1(T, n), _rep1(O, n), _rep1(N, n)
    iv = _bn(np.nan if instance is None else instance, n)
    out = []
    for i in range(n):
        s, o, v = str(T[i]), str(O[i]), str(N[i])
        if o == "":
            out.append(s)
        elif iv[i] != iv[i]:
            out.append(s.replace(o, v))
        else:
            out.append(_nth_replace(s, o, v, int(round(float(iv[i])))))
    return np.array(out, dtype=object)


def _proper(text):
    """PROPER：每个「词」首字母大写、其余小写（Excel 把非字母都当分隔符）。"""
    S = _xstr(text)
    return np.array(
        [_re.sub(r"[A-Za-z]+",
                 lambda m: m.group(0)[:1].upper() + m.group(0)[1:].lower(), str(s))
         for s in S],
        dtype=object)


def _exact(a, b):
    """EXACT：区分大小写的相等比较。"""
    A = _xstr(a)
    B = _xstr(b)
    n = max(len(A), len(B))
    A, B = _rep1(A, n), _rep1(B, n)
    return np.array([str(A[i]) == str(B[i]) for i in range(n)], dtype=bool)


# Excel TEXT 的**受限子集**：只支持下表这些格式代码。
# converter 在转换期做白名单校验 —— 不支持的格式会让该列进入「失败列」，而不是猜。
_TEXT_DATES = {
    "yyyy-mm-dd": lambda t: f"{t.year:04d}-{t.month:02d}-{t.day:02d}",
    "yyyy/m/d": lambda t: f"{t.year:04d}/{t.month}/{t.day}",
    "m/d/yyyy": lambda t: f"{t.month}/{t.day}/{t.year:04d}",
    "mm/dd/yyyy": lambda t: f"{t.month:02d}/{t.day:02d}/{t.year:04d}",
    "yyyy-mm": lambda t: f"{t.year:04d}-{t.month:02d}",
    "yyyy": lambda t: f"{t.year:04d}",
    "yyyy年m月d日": lambda t: f"{t.year:04d}年{t.month}月{t.day}日",
    "yyyy年m月": lambda t: f"{t.year:04d}年{t.month}月",
}


def _text(value, fmt):
    """TEXT：格式化（金额 / 百分比 / 日期）。格式白名单见 _TEXT_DATES 与 converter。"""
    f = str(fmt)
    if f in _TEXT_DATES:
        D = _dvec(value)
        return np.array(
            [("" if pd.isna(d) else _TEXT_DATES[f](pd.Timestamp(d))) for d in D],
            dtype=object)
    if f.endswith("%"):
        dec = len(f.split(".")[1]) - 1 if "." in f else 0
        V = _bn(value) * 100.0
        out = []
        for v in V:
            if v != v:
                out.append("")
                continue
            out.append(f"{float(_round(v, dec)):.{dec}f}%")
        return np.array(out, dtype=object)
    comma = "," in f
    dec = len(f.split(".")[1]) if "." in f else 0
    V = _bn(value)
    out = []
    for v in V:
        if v != v:
            out.append("")
            continue
        r = float(_round(v, dec))     # 走 Excel 的「.5 远离零」，不是 Python 的银行家舍入
        s = f"{abs(r):,.{dec}f}" if comma else f"{abs(r):.{dec}f}"
        out.append(("-" if r < 0 else "") + s)
    return np.array(out, dtype=object)
