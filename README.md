# excel2pandas

**把 Excel 模板公式翻译成可读、可对账、可维护的 Pandas 代码。**

读一遍模板样板行的公式 → 生成一份自包含的 Python 文件 → 批量跑任意多条记录 → 逐格对账，证明生成代码没有算错。

不是通用 Excel 转换器，不是执行引擎，不做双向同步。

---

## 文档

| 文档 | 内容 |
|---|---|
| [`docs/产品设计方案.md`](docs/产品设计方案.md) | **主文档。** 产品定位 + 5 模块架构 + 核心设计决策 + Excel 语义陷阱清单 + 验收口径 |
| [`examples/`](examples/) | 可直接运行的演示案例（虚构的电商定价场景） |
| `README.md`（本文） | 快速开始 + 支持边界 + 已知坑 |

---

## 一、安装

```bash
pip install -r requirements.txt
# 国内建议走镜像：
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

依赖：`pandas` / `openpyxl` / `numpy-financial`（财务函数用）。

也可 `pip install -e .` 后直接用 `excel2pandas` 命令。

---

## 二、快速开始

仓库自带一个**虚构的电商定价测算模板** `examples/demo_template.xlsx`，可以直接跑通全流程：

```bash
# ① 生成：读 Excel 样板行 → 出代码 + 3 份报告
python -m excel2pandas generate \
    --file examples/demo_template.xlsx \
    --sheet "定价测算" --row 2 \
    --out examples/

# ② 单条调试：跑一条，打印结果
python -m excel2pandas run --code examples/generated_code.py --key "SKU-10001"

# ③ 批量：CSV 清单 + 断点续传
python -m excel2pandas run --code examples/generated_code.py \
    --keys examples/keys.csv --progress examples/progress.json --resume \
    --out examples/result.xlsx

# ④ 一键对账：生成代码 vs 模板缓存值，逐格 diff
python -m excel2pandas verify --code examples/generated_code.py \
    --excel examples/demo_template.xlsx \
    --sheet "定价测算" --row 2 --key "SKU-10001" \
    --report examples/verify_report.md
```

生成默认走缓存（Excel 未变则跳过）。`--force` 强制重新生成。

**想从零重建上面全部产物**（含演示模板本身）：

```bash
python examples/build_examples.py
```

---

## 三、项目结构

```
excel2pandas/
├── excel2pandas/                 包本体（5 模块 + 运行时库 + CLI）
│   ├── reader.py                 读簿建模型（双读 + 列建模 + 分层 + 评估报告）
│   ├── converter.py              公式 → pandas 表达式（AST 转换 + 拓扑 + 函数注册表）
│   ├── generator.py              编排 + 渲染 + 报告 + 缓存
│   ├── runner.py                 执行（单条 / 批量 / 断点续传 / 异常映射）
│   ├── verifier.py               对账（逐记录 + 逐字段 diff）
│   ├── _runtime.py               运行时辅助库（会被嵌入生成代码）
│   └── __main__.py               CLI 路由
├── docs/产品设计方案.md           主文档
├── examples/                     演示案例（全部为虚构数据）
│   ├── make_demo_template.py     造演示模板
│   ├── build_examples.py         一键重建全部产物
│   └── demo_template.xlsx        造出来的演示模板
├── probe/probe_template.py       模板探针（接手新模板时先跑这个）
├── requirements.txt / pyproject.toml / .gitignore
└── README.md
```

---

## 四、产物

```
output/
├── generated_code.py        # 生成的代码（自包含，无需安装本包）
├── business_rules.md        # CONFIG 说明：每个业务常量的值 + 被哪些列使用
├── migration_report.md      # 迁移评估：规模/依赖/分层/风险/覆盖率
├── verify_report.md         # 对账报告：一致率 + 差异明细
└── progress.json            # 批量进度（断点续传用）
```

### generated_code.py 长什么样

下面是 `examples/` 里真实生成的内容（节选）：

```python
CONFIG = {
    "含税系数": 1.13,
    "年化费率": 0.06,
    "_rate_0_05": 0.05,          # 未命中命名映射 -> 启发式命名
    "_rate_1_2": 1.2,
    "_rate_0_3": 0.3,
}

# 源文件路径：换数据源只改这里
FILE_PATHS = {
    "main": r"examples\demo_template.xlsx",
}

def calc(primary_values, file_paths=None, inputs=None):
    df = pd.DataFrame({'商品编号': pks})
    # ---- 数据层：手填输入参数 ----
    df['成本价'] = _inputs('成本价', pks, default=_DEFAULTS['成本价'])
    ...

    # ---- L2 ----
    # G2: =VLOOKUP(C2,运费价目!$A:$B,2,0)
    df['__tmp_lk_01'] = _lookup(_sheet('运费价目'), 'A', 'B', df['类目'], not_found=None)
    df['头程运费'] = df['__tmp_lk_01']

    # ---- L3 ----
    # H2: =ROUND(E2*1.13,2)
    df['含税成本'] = _round(((df['成本价']) * (CONFIG['含税系数'])), 2)
```

**每列上方保留原公式**，对账时能直接追溯。换数据源只改 `FILE_PATHS`。

### `calc()` 三个参数

| 参数 | 含义 |
|---|---|
| `primary_values` | 主键值列表，批量时一次传一批 |
| `file_paths` | 文件路径字典，覆盖生成的 `FILE_PATHS` |
| `inputs` | **手填参数列的逐行覆盖值**——手工定价这类列每行都不同，必须能逐行传入 |

> `inputs` 是很容易漏的一个参数。样板行的手填值只作为 `CONFIG` 兜底默认值，
> 批量跑不同记录时必须逐行覆盖，否则整批都用同一个值算。

---

## 五、演示案例（`examples/`）

`examples/` 是一个**虚构的电商定价测算场景**，用来展示产物形态：

| 文件 | 内容 |
|---|---|
| `demo_template.xlsx` | 3 个 sheet：主表「定价测算」+ 外部表「运费价目」「平台费率」 |
| `make_demo_template.py` | 造模板的脚本（含公式与缓存值） |
| `build_examples.py` | 一键重建全部产物 |
| `generated_code.py` | 生成的自包含代码 |
| `business_rules.md` / `migration_report.md` / `verify_report.md` | 三份报告 |
| `keys.csv` / `result.xlsx` | 批量输入清单与结果表 |

主表 20 列 = **8 数据列 + 11 公式列 + 1 空列**，覆盖：跨表 `VLOOKUP`、条件逻辑、`ROUND`/`CEILING`、财务函数 `PMT`、同表引用链。

实测结果：**转换 11/11（失败 0）、一致率 100.00%、批量 3 条 0 失败。**

---

## 六、3+1 层执行模型

```
数据层        从 Excel 读输入参数（手填列）
   ↓
Layer 1      直接引用        =B2
Layer 2      跨表操作        VLOOKUP / XLOOKUP / INDEX+MATCH / SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS
Layer 3      其他公式        算术 / IF / ROUND / CEILING / RATE / FV / CUMIPMT / SUM(区域) ...
```

分层冲突（低层依赖高层）时**自动降级为全局拓扑排序**，结果等价。

> 为什么必须降级：按分层顺序输出会让 `df["X"]` 的赋值排在它依赖的 `df["Y"]` 之前，
> 而 Y 列已被读成全 NaN——**pandas 不报错，静默算错**。

---

## 七、三分类输出

| 类型 | 生成代码 | 输出 df 列名 |
|---|---|---|
| 数据列 | 不生成赋值（作为输入参数） | 原列名 |
| 成功列 | 生成赋值语句 | 原列名 |
| 失败列 | 不生成赋值，保留 Excel 兜底值 | 原列名 + `_fail` |

失败列不填 `NaN` 而是保留兜底值，这样批量结果整体仍可用，且一眼看出哪些列不可信。

---

## 八、支持的边界

**支持**（内置 42 个函数）

| 类别 | 覆盖 |
|---|---|
| 算术/比较 | `+ - * / ^ %`、`= <> < > <= >=`、字符串拼接 `&` |
| 逻辑 | `IF / IFS / AND / OR / NOT / IFERROR / IFNA` |
| 数学 | `ABS / ROUND / ROUNDUP / ROUNDDOWN / INT / CEILING / FLOOR / MAX / MIN / SUM / PRODUCT / SQRT / MOD / POWER` |
| 文本 | `TEXTBEFORE / TEXTAFTER / CONCAT / TRIM / LEFT / RIGHT / LEN / VALUE` |
| 引用 | `ROW / COLUMN`、同表引用、`$` 绝对/混合引用、表头行引用（`$A$1`） |
| 查找 | `XLOOKUP`（含反向 / `if_not_found`）/ `VLOOKUP`(精确) / `INDEX+MATCH` |
| 条件聚合 | `SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS` |
| 财务 | `RATE / FV / PV / PMT / NPER / IPMT / CUMIPMT` |

**不支持**（→ 失败列 + `_fail` 后缀 + 报告说明）

- 跨工作簿引用 `[Book1.xlsx]Sheet1!A1`
- 同表跨行引用（仅支持上一行 `shift(1)`）
- 数组公式、动态数组
- `TEXT` 格式化（格式规则无法自动翻译）
- 命名区域
- 循环引用（拓扑排序会直接报错）

**新增函数**：用插件接口

```python
from excel2pandas.converter import register_function

@register_function("MY_FUNC")
def _f_my_func(args, node, ctx):
    return f"_my_func({', '.join(args)})"
```

对应补 `_runtime.py` 里的运行时实现；生成代码会内嵌被用到的 helper。

---

## 九、几个必须知道的坑

### 1. `_xlfn.` 前缀

openpyxl 读新版函数会带前缀（`_xlfn.XLOOKUP`），转换器会自动剥掉。
**不处理的话所有 XLOOKUP 一个都转不了。**

### 2. 越界 `VLOOKUP` 自动纠正

`VLOOKUP(A2,'台账'!B:H,8,0)` —— `B:H` 只有 7 列却取第 8 列，Excel 返回 `#REF!`。
生成器按「从起始列偏移 7 列」自动纠正，并在 `business_rules.md` 里提示人工确认。

### 3. 文本数字的隐式转换

单元格里存的是文本 `'7.5%'` 时，Excel 参与算术会隐式转成 `0.075`；Python 不转会直接 `TypeError`。
生成器**只在「缓存值本身就是数字样式文本」的列上**加数值化包装，不滥用。

### 4. `SUMIFS` 无匹配返回 0，`XLOOKUP` 无匹配返回 `#N/A`

两者语义不同，混同会让批量结果整片变 NaN：

| 函数 | 无匹配时 Excel 返回 | 实现 |
|---|---|---|
| `SUMIFS` / `COUNTIFS` / `MAXIFS` / `MINIFS` | `0` | 填 0 |
| `AVERAGEIFS` | `#DIV/0!` | NaN |
| `XLOOKUP` / `VLOOKUP` | `#N/A` | NaN |

### 5. `ROUND` 不是 numpy 的 round

Excel 的 `ROUND` 是「四舍五入远离零」（`ROUND(2.5,0)=3`），numpy 是银行家舍入（`=2`）。
运行时用 `_round` / `_roundup` / `_rounddown` 复现 Excel 语义。

### 6. Excel 侧错误值要单独归类

pandas 会把 `#REF!` / `#N/A` 读成 NaN，导致这些单元格被误判成「空值」或「逻辑差异」。
对账器额外用 openpyxl 读原始缓存值，把这类单元格归为 **Excel侧错误**，不计入一致率分母。

### 7. 常量命名映射的键格式

`generator.DEFAULT_CONFIG_NAMES` 与 `config_names.json` 的键由 `_skey(值)` 生成：
**小数点写成下划线**。`1.13 → "1_13"`、`0.05 → "0_05"`、`-0.5 → "neg0_5"`。
写成 `"1.13"` 是查不到的（查找侧走 `_skey()`，键必须同格式）。

---

## 十、批量性能

生成的代码是**向量化**的：跨表查找与聚合由 `_lookup` / `_aggif` 一次性映射到全表，
**不在生成代码里做 merge**（因此不存在多次 merge 打乱行序的问题）。

`run_batch` 按 chunk（默认 200 条）调用 `calc`，工作簿句柄与 sheet 缓存跨 chunk 复用。

> 在一个已脱敏的真实项目上验证：100+ 列 / 13 张外部表 / 最大 8 层依赖 / 依赖表最大 8.8 万行，
> 批量 50 条记录约 30 秒。

---

## 十一、验收口径

一次迁移只有同时满足三条才算完成：

1. **转换覆盖率** —— 公式列转换成功数 / 公式列总数
2. **一致率（排除 Excel 侧错误）** —— 生成代码 vs Excel 缓存值逐格对比
3. **可批量** —— 生成的代码能直接跑 N 条主键并产出结果表

其中第 2 条不可替代，它是对「生成代码是否算错」的唯一客观证据。
`examples/verify_report.md` 是一份真实的对账报告样例（一致率 100.00%）。
