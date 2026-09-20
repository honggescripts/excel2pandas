# ========================================================================
# 自动生成 by excel2pandas 0.2.0
# 源文件   : examples\demo_template.xlsx
# Sheet    : 定价测算
# 表头行   : 1    样板行: 2
# 生成时间 : 2026-09-20 16:28:33
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
# 渲染模式   : minimal —— 只内联正文用到的 20 个 helper
#              要全量嵌入：generate(..., runtime_mode="full")
# 内联 helper: _FILE_PATHS、_HEADER_ROW、_NROW、_SHEET_CACHE、_XLSX、_ERROR_STRINGS、_unbox、_fit、_num、_book、_parse、_sheet、_pick、_lookup、_is_arr、_bool、_if、_div、_round、_ceiling
# runtime 行数: 201
# 外部依赖   : pip install numpy pandas
#              本文件自包含，**不需要**安装 excel2pandas



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
# 运行时辅助库（内联，自包含 —— 201 行）
# ========================================================================
import numpy as np
import pandas as pd



# ---- 由 calc() 在运行期注入的模块级状态 ----
_FILE_PATHS = {}


_HEADER_ROW = 1


_NROW = 1


_SHEET_CACHE = {}


_XLSX = None      # 复用同一个工作簿句柄 —— 否则每读一个 sheet 都要重开整个工作簿



# Excel 错误值字符串：读进来的数据里可能混着这些，统一清洗成 NaN，
# 否则它们会以 str 形式漏进数值运算，直接炸掉。
_ERROR_STRINGS = ["#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!"]





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




def _num(x):
    a = _unbox(x)
    if a.dtype.kind in "iufc":
        return a.astype(float)
    if a.dtype.kind in "OSU":
        a = pd.to_numeric(pd.Series(a), errors="coerce").to_numpy()
    return a.astype(float)




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




# ================================================================ 逻辑

def _is_arr(x):
    return isinstance(x, (np.ndarray, pd.Series, pd.Index, list, tuple))




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
