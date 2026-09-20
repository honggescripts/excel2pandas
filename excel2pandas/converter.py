# -*- coding: utf-8 -*-
"""
converter.py —— 依赖拓扑 + 公式转换 + 函数插件（合并原设计的 dependency + converter + plugin）

职责：
  1. Excel 公式 -> AST（分词 + 递归下降解析）
  2. AST -> pandas 表达式；跨表操作（lookup / 聚合）提升为临时列 + merge
  3. 函数注册表（可扩展 / 覆写）
  4. 全局拓扑排序（分层降级后的执行顺序）
  5. 魔法值常量抽取 -> CONFIG

不做：文件 IO、代码渲染、执行。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

# ================================================================ AST


@dataclass
class Num:
    value: float


@dataclass
class Str:
    value: str


@dataclass
class Bool:
    value: bool


@dataclass
class ErrLit:
    value: str


@dataclass
class Ref:
    col: str
    row: Optional[int] = None          # None = 整列
    sheet: Optional[str] = None
    abs_col: bool = False
    abs_row: bool = False
    whole_col: bool = False


@dataclass
class Range:
    c1: str
    c2: str
    r1: Optional[int] = None
    r2: Optional[int] = None
    sheet: Optional[str] = None


@dataclass
class Unary:
    op: str
    operand: Any


@dataclass
class BinOp:
    op: str
    left: Any
    right: Any


@dataclass
class Call:
    name: str
    args: List[Any]


@dataclass
class MapTable:
    """嵌套 IF(键=常量, 值, ...) 退化成的「查找表」——渲染为字典映射。"""
    key: Any
    pairs: List[Tuple[float, Any]]
    default: Any


# ================================================================ 分词

_T_PATTERNS: List[Tuple[str, str]] = [
    ("QUAL_REF", r"(?:'(?P<q>[^']+)'|(?P<p>[A-Za-z0-9\u4e00-\u9fa5_\.]+))\s*!\s*"
                 r"(?P<ref>\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+"
                 r"|\$?[A-Z]{1,3}\s*:\s*\$?[A-Z]{1,3}(?![A-Za-z0-9_])"
                 r"|\$?[A-Z]{1,3}\$?\d+)"),
    ("EXTERNAL", r"\[[^\]]+\][^!]*!\s*\$?[A-Z]{1,3}\$?\d+"),
    ("ERR", r"#[A-Z0-9/!?]+"),
    ("STR", r'"(?:[^"]|"")*"'),
    ("RANGE", r"\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+"),
    ("WCOL", r"\$?[A-Z]{1,3}\s*:\s*\$?[A-Z]{1,3}(?![A-Za-z0-9_])"),
    ("REF", r"\$?[A-Z]{1,3}\$?\d+"),
    ("NUM", r"(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?"),
    ("IDENT", r"[A-Za-z_\\][A-Za-z0-9_\.\\]*"),
    ("OP", r"<>|<=|>=|[=<>+\-*/^&%]"),
    ("LP", r"\("),
    ("RP", r"\)"),
    ("COMMA", r"[,;]"),
    ("SP", r"\s+"),
    ("UNK", r"."),
]
_TOKEN_RE = re.compile("|".join(f"(?P<{n}>{p})" for n, p in _T_PATTERNS))


@dataclass
class Token:
    kind: str
    text: str
    pos: int
    groups: Dict[str, Any] = field(default_factory=dict)


class FormulaError(Exception):
    pass


def tokenize(src: str) -> List[Token]:
    toks: List[Token] = []
    for m in _TOKEN_RE.finditer(src):
        kind = m.lastgroup
        if kind == "SP":
            continue
        toks.append(Token(kind, m.group(0), m.start(), m.groupdict()))
    return toks


# ================================================================ 解析

_CMP = {"=", "<>", "<", ">", "<=", ">="}

# openpyxl 读出的新版函数带 `_xlfn.` 前缀，必须剥掉，否则 26 处 XLOOKUP 全部失败
_FUNC_PREFIXES = ("_XLFN.", "XLFN.", "_XLWS.", "_XLFN.")


def _norm_func(raw: str) -> str:
    n = raw.upper()
    for pre in _FUNC_PREFIXES:
        if n.startswith(pre):
            n = n[len(pre):]
            break
    return n


class Parser:
    """递归下降解析 Excel 公式。"""

    def __init__(self, formula: str):
        self.src = formula.lstrip("=").strip()
        self.toks = tokenize(self.src)
        self.i = 0

    # ---- 工具 ----
    def peek(self, k: int = 0) -> Optional[Token]:
        j = self.i + k
        return self.toks[j] if j < len(self.toks) else None

    def next(self) -> Token:
        if self.i >= len(self.toks):
            raise FormulaError(f"公式意外结束：{self.src}")
        t = self.toks[self.i]
        self.i += 1
        return t

    def eat(self, kind: str, text: Optional[str] = None) -> Optional[Token]:
        t = self.peek()
        if t and t.kind == kind and (text is None or t.text == text):
            self.i += 1
            return t
        return None

    def expect(self, kind: str) -> Token:
        t = self.peek()
        if not t or t.kind != kind:
            raise FormulaError(f"期望 {kind}，实到 {t.text if t else 'EOF'}（{self.src}）")
        return self.next()

    # ---- 入口 ----
    def parse(self) -> Any:
        node = self.expr()
        if self.i != len(self.toks):
            rest = "".join(t.text for t in self.toks[self.i:])
            raise FormulaError(f"公式未完全解析，剩余：{rest}")
        return node

    # ---- 优先级分层 ----
    def expr(self) -> Any:
        return self.cmp()

    def cmp(self) -> Any:
        node = self.concat()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.text in _CMP:
                self.next()
                node = BinOp(t.text, node, self.concat())
            else:
                return node

    def concat(self) -> Any:
        node = self.add()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.text == "&":
                self.next()
                node = BinOp("&", node, self.add())
            else:
                return node

    def add(self) -> Any:
        node = self.mul()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.text in ("+", "-"):
                self.next()
                node = BinOp(t.text, node, self.mul())
            else:
                return node

    def mul(self) -> Any:
        node = self.unary()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.text in ("*", "/"):
                self.next()
                node = BinOp(t.text, node, self.unary())
            else:
                return node

    def unary(self) -> Any:
        t = self.peek()
        if t and t.kind == "OP" and t.text in ("+", "-"):
            self.next()
            return Unary(t.text, self.unary())
        return self.power()

    def power(self) -> Any:
        node = self.postfix()
        t = self.peek()
        if t and t.kind == "OP" and t.text == "^":
            self.next()
            return BinOp("^", node, self.unary())   # 右结合
        return node

    def postfix(self) -> Any:
        node = self.primary()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.text == "%":
                self.next()
                node = Unary("%", node)
            else:
                return node

    def primary(self) -> Any:
        t = self.peek()
        if t is None:
            raise FormulaError(f"公式意外结束：{self.src}")

        if t.kind == "LP":
            self.next()
            node = self.expr()
            self.expect("RP")
            return node

        if t.kind == "NUM":
            self.next()
            return Num(float(t.text))

        if t.kind == "STR":
            self.next()
            return Str(t.text[1:-1].replace('""', '"'))

        if t.kind == "ERR":
            self.next()
            return ErrLit(t.text)

        if t.kind in ("REF", "WCOL", "RANGE", "QUAL_REF"):
            self.next()
            return self._make_ref(t)

        if t.kind == "EXTERNAL":
            self.next()
            raise FormulaError(f"不支持跨工作簿引用：{t.text}")

        if t.kind == "IDENT":
            self.next()
            name = _norm_func(t.text)
            # TRUE / FALSE 字面量
            if name in ("TRUE", "FALSE") and not (self.peek() and self.peek().kind == "LP"):
                return Bool(name == "TRUE")
            if self.peek() and self.peek().kind == "LP":
                self.next()
                args: List[Any] = []
                if not (self.peek() and self.peek().kind == "RP"):
                    args.append(self.expr())
                    while self.eat("COMMA"):
                        if self.peek() and self.peek().kind == "RP":
                            break
                        args.append(self.expr())
                self.expect("RP")
                return Call(name, args)
            # 命名区域 —— 不支持
            raise FormulaError(f"不支持命名区域：{t.text}")

        raise FormulaError(f"无法解析的记号 {t.text!r}（{self.src}）")

    # ---- 引用构造 ----
    @staticmethod
    def _split_ref(token: str, sheet: Optional[str]) -> Any:
        token = token.replace(" ", "")
        if ":" in token:
            a, b = token.split(":", 1)
            ma = re.match(r"^(\$?)([A-Z]{1,3})(\$?)(\d*)$", a)
            mb = re.match(r"^(\$?)([A-Z]{1,3})(\$?)(\d*)$", b)
            if not ma or not mb:
                raise FormulaError(f"无法解析区域：{token}")
            return Range(
                c1=ma.group(2), c2=mb.group(2),
                r1=int(ma.group(4)) if ma.group(4) else None,
                r2=int(mb.group(4)) if mb.group(4) else None,
                sheet=sheet,
            )
        m = re.match(r"^(\$?)([A-Z]{1,3})(\$?)(\d*)$", token)
        if not m:
            raise FormulaError(f"无法解析引用：{token}")
        return Ref(
            col=m.group(2),
            row=int(m.group(4)) if m.group(4) else None,
            sheet=sheet,
            abs_col=bool(m.group(1)),
            abs_row=bool(m.group(3)),
            whole_col=(m.group(4) == ""),
        )

    def _make_ref(self, t: Token) -> Any:
        if t.kind == "QUAL_REF":
            sheet = t.groups["q"] or t.groups["p"]
            return self._split_ref(t.groups["ref"], sheet)
        return self._split_ref(t.text, None)


def parse(formula: str) -> Any:
    return Parser(formula).parse()


# ================================================================ 遍历工具


def walk(node: Any):
    yield node
    if isinstance(node, Unary):
        yield from walk(node.operand)
    elif isinstance(node, BinOp):
        yield from walk(node.left)
        yield from walk(node.right)
    elif isinstance(node, Call):
        for a in node.args:
            yield from walk(a)


def func_names(node: Any) -> Set[str]:
    out: Set[str] = set()

    def rec(n: Any):
        if isinstance(n, Call):
            out.add(n.name)
            for a in n.args:
                rec(a)
        elif isinstance(n, Unary):
            rec(n.operand)
        elif isinstance(n, BinOp):
            rec(n.left)
            rec(n.right)

    rec(node)
    return out


# ================================================================ 配置


@dataclass
class ConverterOptions:
    min_literal_to_extract: float = 100.0     # 数字 >= 该值 -> 抽到 CONFIG
    extract_rates: bool = True                # 费率样式(0.xx / 1.xx) -> 抽到 CONFIG
    div_guard: bool = True                    # 除数为动态值时用 _div 防零除
    target_row: int = 2
    header_row: int = 1


@dataclass
class TempCol:
    name: str
    kind: str            # lookup / aggif / diref
    code: str            # 生成的赋值语句
    comment: str         # 对应 Excel 片段
    warning: str = ""    # 自动纠正等提示


@dataclass
class ConvertResult:
    letter: str
    name: str
    layer: str
    expr: str = ""
    pre_stmts: List[str] = field(default_factory=list)
    temp_columns: List[str] = field(default_factory=list)
    success: bool = True
    error: Optional[str] = None
    original_formula: str = ""
    config_keys: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ================================================================ 函数注册表

LOOKUP_FUNCS = {"VLOOKUP", "XLOOKUP", "HLOOKUP", "LOOKUP"}
AGGIF_FUNCS = {"SUMIFS", "COUNTIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"}
AGGIF_HOW = {"SUMIFS": "sum", "COUNTIFS": "size", "AVERAGEIFS": "mean", "MAXIFS": "max", "MINIFS": "min"}

FuncImpl = Callable[[List[str], Any, "RenderCtx"], str]
_FUNC_REGISTRY: Dict[str, FuncImpl] = {}


def register_function(*names: str):
    """插件接口：@register_function("PMT") -> 注册 / 覆写一个 Excel 函数实现。"""
    def deco(fn: FuncImpl) -> FuncImpl:
        for n in names:
            _FUNC_REGISTRY[n.upper()] = fn
        return fn
    return deco


def registered_names() -> List[str]:
    return sorted(_FUNC_REGISTRY)


def _join(args: Sequence[str], sep: str) -> str:
    return sep.join(f"({a})" for a in args)


@register_function("ABS")
def _f_abs(a, n, c): return f"np.abs({a[0]})"

@register_function("INT")
def _f_int(a, n, c): return f"np.floor({a[0]})"

@register_function("ROUND")
def _f_round(a, n, c): return f"_round({a[0]}, {a[1] if len(a) > 1 else 0})"

@register_function("ROUNDUP")
def _f_roundup(a, n, c): return f"_roundup({a[0]}, {a[1] if len(a) > 1 else 0})"

@register_function("ROUNDDOWN")
def _f_rounddown(a, n, c): return f"_rounddown({a[0]}, {a[1] if len(a) > 1 else 0})"

@register_function("CEILING", "CEILING.MATH")
def _f_ceiling(a, n, c): return f"_ceiling({a[0]}, {a[1]})"

@register_function("FLOOR", "FLOOR.MATH")
def _f_floor(a, n, c): return f"_floor_({a[0]}, {a[1]})"

@register_function("SQRT")
def _f_sqrt(a, n, c): return f"np.sqrt({a[0]})"

@register_function("MOD")
def _f_mod(a, n, c): return f"np.mod({a[0]}, {a[1]})"

@register_function("POWER")
def _f_power(a, n, c): return f"(({a[0]}) ** ({a[1]}))"

@register_function("MAX")
def _f_max(a, n, c):
    return f"_pmax([{', '.join(a)}])" if len(a) > 1 else f"({a[0]})"

@register_function("MIN")
def _f_min(a, n, c):
    return f"_pmin([{', '.join(a)}])" if len(a) > 1 else f"({a[0]})"

@register_function("SUM")
def _f_sum(a, n, c):
    if len(a) == 1:
        return a[0]
    return "_psum([" + ", ".join(a) + "])"

@register_function("PRODUCT")
def _f_product(a, n, c):
    if len(a) == 1:
        return a[0]
    return "_pprod([" + ", ".join(a) + "])"

@register_function("IF")
def _f_if(a, n, c):
    # Excel 允许 IF(cond, a) 两参形式，缺省第三参 -> FALSE
    if len(a) == 2:
        return f"_if({a[0]}, {a[1]}, False)"
    if len(a) < 3:
        raise FormulaError("IF 至少需要 2 个参数")
    return f"_if({a[0]}, {a[1]}, {a[2]})"

@register_function("IFS")
def _f_ifs(a, n, c):
    if len(a) % 2:
        raise FormulaError("IFS 参数个数必须为偶数")
    expr = a[-1]
    for i in range(len(a) - 2, -1, -2):
        expr = f"_if({a[i]}, {a[i+1]}, {expr})"
    return expr

@register_function("IFERROR", "IFNA")
def _f_iferror(a, n, c): return f"_iferror({a[0]}, {a[1]})"

@register_function("AND")
def _f_and(a, n, c): return "_and([" + ", ".join(a) + "])"

@register_function("OR")
def _f_or(a, n, c): return "_or([" + ", ".join(a) + "])"

@register_function("NOT")
def _f_not(a, n, c): return f"(~_bool({a[0]}))"

@register_function("ROW")
def _f_row(a, n, c): return "_row(df)"

@register_function("COLUMN")
def _f_column(a, n, c): return "_column(df, " + (a[0] if a else "0") + ")"

@register_function("VALUE")
def _f_value(a, n, c): return f"_tonum({a[0]})"

@register_function("TEXTBEFORE")
def _f_textbefore(a, n, c): return f"_textbefore({a[0]}, {a[1]})"

@register_function("TEXTAFTER")
def _f_textafter(a, n, c): return f"_textafter({a[0]}, {a[1]})"

@register_function("CONCAT", "CONCATENATE")
def _f_concat(a, n, c): return "_cat([" + ", ".join(a) + "])" if len(a) > 1 else f"_cat([{a[0]}])"

@register_function("TRIM")
def _f_trim(a, n, c): return f"_str({a[0]}).str.strip()"

@register_function("LEFT")
def _f_left(a, n, c): return f"_str({a[0]}).str.slice(0, {a[1]})"

@register_function("RIGHT")
def _f_right(a, n, c): return f"_str({a[0]}).str.slice(-({a[1]}))"

@register_function("LEN")
def _f_len(a, n, c): return f"_str({a[0]}).str.len()"

# ---- 财务函数 ----

@register_function("RATE")
def _f_rate(a, n, c): return f"_rate({', '.join(a)})"

@register_function("FV")
def _f_fv(a, n, c): return f"_fv({', '.join(a)})"

@register_function("PV")
def _f_pv(a, n, c): return f"_pv({', '.join(a)})"

@register_function("PMT")
def _f_pmt(a, n, c): return f"_pmt({', '.join(a)})"

@register_function("NPER")
def _f_nper(a, n, c): return f"_nper({', '.join(a)})"

@register_function("IPMT")
def _f_ipmt(a, n, c): return f"_ipmt({', '.join(a)})"

@register_function("CUMIPMT")
def _f_cumipmt(a, n, c): return f"_cumipmt({', '.join(a)})"

@register_function("TEXT")
def _f_text(a, n, c):
    raise FormulaError("不支持 TEXT 格式化（格式化规则无法自动翻译）")

# ================================================================ 渲染上下文


class RenderCtx:
    def __init__(self, model, opts: ConverterOptions, config_names: Optional[Dict[str, str]] = None):
        self.model = model
        self.opts = opts
        self.config_names = config_names or {}
        self.temps: List[TempCol] = []
        self._seq = 0
        self.literals: Dict[float, str] = {}     # 值 -> CONFIG key
        self.config_values: Dict[str, float] = {}
        self.warnings: List[str] = []
        self._used_this_col: Set[str] = set()    # 本列实际引用到的 CONFIG key

    # ---- 逐列开始 ----
    def begin_column(self) -> None:
        self.temps = []
        self.warnings = []
        self._used_this_col = set()

    # ---- 临时列 ----
    def new_temp(self, kind: str, code_fmt: str, comment: str, warning: str = "") -> str:
        self._seq += 1
        name = f"__tmp_{kind}_{self._seq:02d}"
        code = code_fmt.format(tmp=name)
        self.temps.append(TempCol(name=name, kind=kind, code=code, comment=comment, warning=warning))
        return name

    # ---- 魔法值 ----
    def _rate_like(self, v: float) -> bool:
        """费率样式：0.xx（不含整数）或 1.xx（不含整数）。"""
        if not self.opts.extract_rates:
            return False
        a = abs(v)
        if a == 0 or a >= 10:
            return False
        if float(a).is_integer():
            return False
        return (0 < a < 1) or (1 < a < 2)

    def literal(self, v: float) -> str:
        """数值 -> CONFIG 引用 或 内联字面量。"""
        if v != v or v in (float("inf"), float("-inf")):
            return "np.nan"
        # 抹掉浮点噪声：1.35% 会算出 0.013500000000000002
        v = float(f"{v:.12g}")
        should = (abs(v) >= self.opts.min_literal_to_extract) or self._rate_like(v)
        if not should:
            if float(v).is_integer() and abs(v) < 1e15:
                return str(int(v))
            return repr(v)
        if v in self.literals:
            self._used_this_col.add(self.literals[v])
            return f"CONFIG[{self.literals[v]!r}]"
        key = self.config_names.get(_skey(v)) or self._auto_key(v)
        # 去重：同名但不同值 -> 加后缀
        base = key
        k = 2
        while key in self.config_values and self.config_values[key] != v:
            key = f"{base}_{k}"
            k += 1
        self.literals[v] = key
        self.config_values[key] = v
        self._used_this_col.add(key)
        return f"CONFIG[{key!r}]"

    def _auto_key(self, v: float) -> str:
        if self._rate_like(v) and abs(v) < 2:
            return f"_rate_{_skey(v)}"
        if abs(v) >= 1000:
            return f"_fee_{_skey(v)}"
        return f"_c_{_skey(v)}"

    def col_name(self, letter: str) -> str:
        m = self.model.pick(letter)
        if m is None:
            raise FormulaError(f"目标表不存在列 {letter}")
        return m.name

    def col_expr(self, letter: str) -> str:
        c = self.model.pick(letter)
        expr = f"df[{c.name!r}]"
        # 缓存值是「数字样式的文本」（如 '7.5%'、'3,900'）-> 参与运算前先数值化，
        # 复现 Excel 的隐式转换，否则 '7.5%'/12 会抛 TypeError。
        if getattr(c, "cached_numeric_text", False):
            expr = f"_n({expr})"
        return expr


def _skey(v: float) -> str:
    """数值 -> 安全 key 片段。"""
    if float(v).is_integer():
        s = str(int(v))
    else:
        s = repr(v)
    return s.replace("-", "neg").replace(".", "_")


# ================================================================ 跨表提升


def _sheet_of(node: Any) -> Optional[str]:
    return getattr(node, "sheet", None)


def _colrange(node: Any) -> Tuple[str, str]:
    """Ref/Range -> (起始列, 终止列)"""
    if isinstance(node, Ref):
        return node.col, node.col
    if isinstance(node, Range):
        return node.c1, node.c2
    raise FormulaError("需要引用或区域")


def _is_whole_col_same_sheet(node: Any, model) -> Optional[str]:
    """`E:E` 这种同表整列 —— 语义是「当前行的 E 值」。"""
    if isinstance(node, Range) and node.sheet is None and node.r1 is None and node.r2 is None:
        if node.c1 == node.c2:
            return node.c1
    return None


def _lift(node: Any, ctx: RenderCtx, result: ConvertResult) -> Any:
    """把跨表操作 / 查找 / 聚合 提升为临时列，返回替换后的 AST。"""
    model = ctx.model

    if isinstance(node, Unary):
        return Unary(node.op, _lift(node.operand, ctx, result))
    if isinstance(node, BinOp):
        return BinOp(node.op, _lift(node.left, ctx, result), _lift(node.right, ctx, result))
    if not isinstance(node, Call):
        return node

    name = node.name
    lifted_args = [_lift(a, ctx, result) for a in node.args]

    # ---- 查找类 ----
    if name in LOOKUP_FUNCS:
        return _lift_lookup(name, lifted_args, ctx, result)

    # ---- 条件聚合类 ----
    if name in AGGIF_FUNCS:
        return _lift_aggif(name, lifted_args, ctx, result)

    if name == "INDEX":
        return _lift_index(lifted_args, ctx, result)

    if name == "MATCH":
        raise FormulaError("MATCH 需与 INDEX 组合使用")

    return Call(name, lifted_args)


def _key_expr(node: Any, ctx: RenderCtx) -> str:
    """查找值：整列引用 -> 当前行值；其他 -> 正常渲染。"""
    wc = _is_whole_col_same_sheet(node, ctx.model)
    if wc:
        return ctx.col_expr(wc)
    return render(node, ctx)


def _lift_lookup(name: str, args: List[Any], ctx: RenderCtx, result: ConvertResult) -> Any:
    if name == "XLOOKUP":
        if len(args) < 3:
            raise FormulaError("XLOOKUP 至少需要 3 个参数")
        key_arr, ret_arr = args[1], args[2]
        not_found = render(args[3], ctx) if len(args) > 3 and not isinstance(args[3], ErrLit) else "None"
        keys = _key_expr(args[0], ctx)
    elif name == "VLOOKUP":
        if len(args) < 3:
            raise FormulaError("VLOOKUP 至少需要 3 个参数")
        table = args[1]
        if not isinstance(table, (Ref, Range)) or _sheet_of(table) is None:
            raise FormulaError("VLOOKUP 的第二个参数必须是跨表区域")
        c1, c2 = _colrange(table)
        idx_node = args[2]
        if not isinstance(idx_node, Num):
            raise FormulaError("VLOOKUP 的列序号必须是常量")
        idx = int(idx_node.value)
        i1 = _col_index(c1)
        i2 = _col_index(c2)
        width = i2 - i1 + 1
        target_i = i1 + idx - 1
        warn = ""
        if idx > width:
            warn = (f"列索引 {idx} 超出区域 {c1}:{c2} 宽度 {width}，"
                    f"已按「{c1} 起偏移 {idx-1} 列」自动纠正为 {_col_letter(target_i)} 列（原表为 #REF!）")
            ctx.warnings.append(warn)
        key_arr = Ref(col=c1, sheet=_sheet_of(table))
        ret_arr = Ref(col=_col_letter(target_i), sheet=_sheet_of(table))
        keys = _key_expr(args[0], ctx)
        not_found = "None"
    else:
        raise FormulaError(f"暂不支持 {name}")

    sheet = _sheet_of(key_arr) or _sheet_of(ret_arr)
    if sheet is None:
        raise FormulaError(f"{name} 需引用跨表数据")
    kc, _ = _colrange(key_arr)
    vc, _ = _colrange(ret_arr)

    comment = f"{name}({_brief(args[0])}, {sheet}!{kc}:{kc} → {vc}:{vc})"
    warn = locals().get("warn", "")
    tmp = ctx.new_temp("lk",
                       f"df['{{tmp}}'] = _lookup(_sheet({sheet!r}), {kc!r}, {vc!r}, {keys}, not_found={not_found})",
                       comment, warn)
    if warn:
        result.warnings.append(warn)
    result.temp_columns.append(tmp)
    return Ref(col=f"@{tmp}")     # 特殊标记：列名以 @ 开头表示临时列


def _lift_index(args: List[Any], ctx: RenderCtx, result: ConvertResult) -> Any:
    """INDEX(返回列, MATCH(键, 查找列, 0)) -> lookup"""
    if len(args) != 2:
        raise FormulaError("暂只支持 INDEX(列, MATCH(...)) 两参数形式")
    ret, mat = args
    if not isinstance(mat, Call) or mat.name != "MATCH":
        raise FormulaError("INDEX 的第二个参数必须是 MATCH")
    if not isinstance(ret, (Ref, Range)) or not isinstance(mat.args[1], (Ref, Range)):
        raise FormulaError("INDEX/MATCH 参数必须是引用")
    sheet = _sheet_of(ret)
    if sheet is None or sheet != _sheet_of(mat.args[1]):
        raise FormulaError("INDEX/MATCH 必须引用同一张跨表")
    kc, _ = _colrange(mat.args[1])
    vc, _ = _colrange(ret)
    keys = render(mat.args[0], ctx)
    tmp = ctx.new_temp("lk",
                       f"df['{{tmp}}'] = _lookup(_sheet({sheet!r}), {kc!r}, {vc!r}, {keys})",
                       f"INDEX/MATCH({sheet}!{kc} → {vc})")
    result.temp_columns.append(tmp)
    return Ref(col=f"@{tmp}")


def _lift_aggif(name: str, args: List[Any], ctx: RenderCtx, result: ConvertResult) -> Any:
    if len(args) < 3 or len(args) % 2 == 0:
        raise FormulaError(f"{name} 参数个数必须是 奇数 且 >= 3")
    val_node = args[0]
    if not isinstance(val_node, (Ref, Range)) or _sheet_of(val_node) is None:
        raise FormulaError(f"{name} 的第一个参数必须是跨表列")
    sheet = _sheet_of(val_node)
    vc, _ = _colrange(val_node)

    pairs = []
    for i in range(1, len(args), 2):
        k_node, c_node = args[i], args[i + 1]
        if not isinstance(k_node, (Ref, Range)) or _sheet_of(k_node) is None:
            raise FormulaError(f"{name} 的条件区域必须是跨表列")
        if _sheet_of(k_node) != sheet:
            raise FormulaError(f"{name} 的所有引用必须来自同一张表")
        kc, _ = _colrange(k_node)
        pairs.append((kc, _key_expr(c_node, ctx)))

    pairs_src = ", ".join(f"({k!r}, {v})" for k, v in pairs)
    tmp = ctx.new_temp("ag",
                       f"df['{{tmp}}'] = _aggif(_sheet({sheet!r}), {vc!r}, [{pairs_src}], {AGGIF_HOW[name]!r})",
                       f"{name}({sheet}!{vc}, {len(pairs)} 个条件)")
    result.temp_columns.append(tmp)
    return Ref(col=f"@{tmp}")


def _brief(node: Any) -> str:
    if isinstance(node, Ref):
        if node.whole_col:
            return f"{node.sheet}!{node.col}:{node.col}" if node.sheet else f"{node.col}:{node.col}"
        return f"{node.col}{node.row or ''}"
    if isinstance(node, Range):
        return f"{node.sheet}!{node.c1}:{node.c2}"
    return "..."


def _col_index(letter: str) -> int:
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def _col_letter(idx: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx)


# ================================================================ 渲染


def render(node: Any, ctx: RenderCtx) -> str:
    """AST -> pandas 表达式字符串。"""
    model = ctx.model
    opts = ctx.opts

    if isinstance(node, Num):
        return ctx.literal(node.value)

    if isinstance(node, Str):
        return repr(node.value)

    if isinstance(node, Bool):
        return repr(node.value)

    if isinstance(node, ErrLit):
        raise FormulaError(f"公式内含 Excel 错误值 {node.value}")

    if isinstance(node, Ref):
        # 临时列（提升后的跨表结果）
        if node.col.startswith("@"):
            return f'df[{node.col[1:]!r}]'
        if node.sheet is not None and node.sheet != model.sheet_name:
            raise FormulaError(f"未提升的跨表引用 {node.sheet}!{node.col}")
        if node.row is None:
            raise FormulaError(f"整列引用 {node.col}:{node.col} 只能出现在查找/聚合函数里")
        # 表头行常量（如 A$1，把表头文字当作参数值取出来）
        if node.row == model.header_row:
            m = model.pick(node.col)
            raw = m.header if m else ""
            try:
                return ctx.literal(float(raw))
            except ValueError:
                return repr(raw)
        if node.row != opts.target_row:
            raise FormulaError(f"跨行引用 {node.col}{node.row}（样板行 {opts.target_row}）不支持")
        return ctx.col_expr(node.col)

    if isinstance(node, Range):
        if node.sheet is not None and node.sheet != model.sheet_name:
            raise FormulaError(f"跨表区域 {node.sheet}!{node.c1}:{node.c2} 只能用于 SUMIFS 等聚合函数")
        raise FormulaError(
            f"同表区域 {node.c1}{node.r1 or ''}:{node.c2}{node.r2 or ''} 只能用于 SUM/AVERAGE/COUNT/MAX/MIN"
        )

    if isinstance(node, Unary):
        if node.op == "%":
            inner = node.operand
            if isinstance(inner, Num):
                return ctx.literal(inner.value / 100.0)
            return f"(({render(inner, ctx)}) / 100)"
        if node.op == "-":
            # Excel 的 `--x` 是「文本/日期转数值」，不能简化成负负得正
            if isinstance(node.operand, Unary) and node.operand.op == "-":
                return f"_neg2({render(node.operand.operand, ctx)})"
            return f"(-{render(node.operand, ctx)})"
        return f"(+{render(node.operand, ctx)})"

    if isinstance(node, MapTable):
        table = ", ".join(f"{ctx.literal(k)}: {render(v, ctx)}" for k, v in node.pairs)
        return f"_map({render(node.key, ctx)}, {{{table}}}, {render(node.default, ctx)})"

    if isinstance(node, BinOp):
        return _render_binop(node, ctx)

    if isinstance(node, Call):
        return _render_call(node, ctx)

    raise FormulaError(f"无法渲染的节点：{node!r}")


def _render_binop(node: BinOp, ctx: RenderCtx) -> str:
    op = node.op
    L = render(node.left, ctx)
    R = render(node.right, ctx)

    # 双负号：Excel 的 `--x` 语义是「文本转数值 / 日期」
    if op == "-" and isinstance(node.left, Unary) and node.left.op == "-":
        return f"_neg2({render(node.left.operand, ctx)})"

    if op == "&":
        return f"_cat2({L}, {R})"
    if op == "^":
        return f"(({L}) ** ({R}))"
    if op == "=":
        return f"_eq({L}, {R})"
    if op == "<>":
        return f"(_ne({L}, {R}))"
    if op in _CMP:
        return f"(({L}) {op} ({R}))"
    if op == "/":
        const = _num_literal(node.right)
        if const is not None and const != 0:
            return f"(({L}) / ({R}))"
        if ctx.opts.div_guard:
            return f"_div({L}, {R})"
        return f"(({L}) / ({R}))"
    return f"(({L}) {op} ({R}))"


def _render_call(node: Call, ctx: RenderCtx) -> str:
    name = node.name
    impl = _FUNC_REGISTRY.get(name)
    if impl is None:
        raise FormulaError(f"不支持函数 {name}")

    # SUM(区域) 特例：区域被渲染成 __SUMCOL__ 前缀的整行求和
    if name in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN") and len(node.args) == 1 and isinstance(node.args[0], Range):
        rng = node.args[0]
        if rng.sheet is None and rng.r1 is None:
            i1, i2 = _col_index(rng.c1), _col_index(rng.c2)
            if i2 < i1:
                i1, i2 = i2, i1
            names = [ctx.col_name(_col_letter(i)) for i in range(i1, i2 + 1)]
            agg = {"SUM": "sum", "AVERAGE": "mean", "COUNT": "count", "MAX": "max", "MIN": "min"}[name]
            return f"df[{names!r}].{agg}(axis=1)"
        if rng.r1 is not None and rng.r1 == rng.r2 == ctx.opts.target_row:
            i1, i2 = _col_index(rng.c1), _col_index(rng.c2)
            if i2 < i1:
                i1, i2 = i2, i1
            names = [ctx.col_name(_col_letter(i)) for i in range(i1, i2 + 1)]
            agg = {"SUM": "sum", "AVERAGE": "mean", "COUNT": "count", "MAX": "max", "MIN": "min"}[name]
            return f"df[{names!r}].{agg}(axis=1)"

    rendered = [render(a, ctx) for a in node.args]
    return impl(rendered, node, ctx)


def _ast_sig(n: Any) -> str:
    """AST 结构签名（忽略常量），用于判断嵌套 IF 的键表达式是否一致。"""
    if isinstance(n, Ref):
        return f"R{n.sheet or ''}!{n.col}{'' if n.row is None else n.row}"
    if isinstance(n, Range):
        return f"G{n.sheet or ''}!{n.c1}:{n.c2}"
    if isinstance(n, Num):
        return "N"
    if isinstance(n, Str):
        return "S"
    if isinstance(n, Bool):
        return "B"
    if isinstance(n, ErrLit):
        return "E"
    if isinstance(n, Unary):
        return f"({n.op}{_ast_sig(n.operand)})"
    if isinstance(n, BinOp):
        return f"({_ast_sig(n.left)}{n.op}{_ast_sig(n.right)})"
    if isinstance(n, Call):
        return f"{n.name}({','.join(_ast_sig(a) for a in n.args)})"
    return repr(n)


def _is_const_node(n: Any) -> bool:
    return _num_literal(n) is not None


def _optimize(node: Any, min_pairs: int = 3) -> Any:
    """
    把 `IF(x=0,v0,IF(x=1,v1,IF(x=2,v2,...)))` 这种「嵌套 IF 当查找表」
    退化成 MapTable —— 生成的代码从 14 层嵌套变成一行字典映射，可读性天差地别。
    """
    pairs: List[Tuple[float, Any]] = []
    key_sig: Optional[str] = None
    cur = node

    while isinstance(cur, Call) and cur.name == "IF" and len(cur.args) == 3:
        cond, val, els = cur.args
        if not (isinstance(cond, BinOp) and cond.op == "="):
            break
        kv = _num_literal(cond.right)
        if kv is None:
            break
        sig = _ast_sig(cond.left)
        if key_sig is None:
            key_sig = sig
        elif sig != key_sig:
            break
        if not _is_const_node(val):
            break
        kk = _num_literal(cond.left) if _is_const_node(cond.left) else None
        pairs.append((kv, val))
        cur = els
        # 值也是常量但键不是 -> 继续；值非常量 -> 上面已 break

    if len(pairs) < min_pairs or key_sig is None:
        return node

    # 键必须全部不重复
    keys = [k for k, _ in pairs]
    if len(set(keys)) != len(keys):
        return node

    # 还原键表达式：从原 AST 里取第一个 cond.left
    first_cond = node.args[0]
    return MapTable(key=first_cond.left, pairs=pairs, default=cur)


def _num_literal(node: Any) -> Optional[float]:
    """纯数值字面量（含 % / 正负号），用于决定是否走 _div 保护。"""
    if isinstance(node, Num):
        return node.value
    if isinstance(node, Unary) and node.op == "-" and isinstance(node.operand, Num):
        return -node.operand.value
    if isinstance(node, Unary) and node.op == "+" and isinstance(node.operand, Num):
        return node.operand.value
    if isinstance(node, Unary) and node.op == "%" and isinstance(node.operand, Num):
        return node.operand.value / 100.0
    return None



# ================================================================ 单列转换


def convert_column(col, model, opts: ConverterOptions, ctx: RenderCtx) -> ConvertResult:
    res = ConvertResult(letter=col.letter, name=col.name, layer=col.layer,
                        original_formula=col.formula or "")
    if not col.formula:
        res.success = False
        res.error = "非公式列"
        return res
    try:
        ast = parse(col.formula)
        ast = _lift(ast, ctx, res)
        ast = _optimize(ast)
        res.expr = render(ast, ctx)
        res.pre_stmts = [t.code for t in ctx.temps]
        res.warnings = list(ctx.warnings)
        res.config_keys = sorted(ctx._used_this_col)
    except Exception as e:  # noqa: BLE001
        res.success = False
        res.error = f"{type(e).__name__}: {e}"
        res.pre_stmts = [t.code for t in ctx.temps]
    return res


# ================================================================ 拓扑排序


def topo_order(model) -> List[str]:
    """
    全局拓扑排序：返回公式列的列字母执行顺序。
    分层模型降级后仍用这个（结果等价）。
    """
    cols = {c.letter: c for c in model.columns}
    formula = [c.letter for c in model.formula_cols]
    fset = set(formula)

    indeg: Dict[str, int] = {c: 0 for c in formula}
    nxt: Dict[str, List[str]] = {c: [] for c in formula}
    for c in formula:
        for d in cols[c].depends_on:
            if d in fset:
                nxt[d].append(c)
                indeg[c] += 1

    ready = sorted([c for c in formula if indeg[c] == 0])
    out: List[str] = []
    while ready:
        cur = ready.pop(0)
        out.append(cur)
        for n in sorted(nxt[cur]):
            indeg[n] -= 1
            if indeg[n] == 0:
                ready.append(n)
        ready.sort()

    if len(out) != len(formula):
        stuck = [c for c in formula if c not in out]
        raise FormulaError(f"依赖图存在循环引用，涉及列：{sorted(stuck)}")
    return out


def find_cycles(model) -> List[str]:
    try:
        topo_order(model)
        return []
    except FormulaError as e:
        return [str(e)]

