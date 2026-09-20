# -*- coding: utf-8 -*-
"""
_runtime.py 的单元测试。

    python -m unittest discover -s tests -v
    python tests/test_runtime.py

为什么这些测试值得存在：这些函数的「正确」定义是 **Excel 的行为**，不是 Python 的直觉。
  · ROUND 是「.5 远离零」，numpy 是银行家舍入
  · CEILING 对负数是「远离零」，np.ceil 是「向正无穷」
  · FLOOR 对负数是「向零」，np.floor 是「向负无穷」
  · MROUND 异号要报 #NUM!，且符号跟随被舍入的数
  · DAYS360 的月末规则、NETWORKDAYS 的含端与扣节假日
写成 numpy 默认行为，跑起来不报错，但结果和 Excel 对不上 —— 那是最难查的一类 bug。
"""

import importlib.util
import math
import pathlib
import sys
import unittest

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "_e2p_runtime", ROOT / "excel2pandas" / "_runtime.py")
rt = importlib.util.module_from_spec(_spec)
sys.modules["_e2p_runtime"] = rt
_spec.loader.exec_module(rt)


# ---------------------------------------------------------------- 测试工具


def dstr(v):
    """datetime64 / Timestamp -> 'YYYY-MM-DD'。"""
    return pd.Timestamp(v).strftime("%Y-%m-%d")


def d1(fn, *args):
    """按「一行数据」调用一个日期函数，返回该行的日期字符串。"""
    rt._NROW = 1
    out = fn(*args)
    return dstr(np.atleast_1d(out)[0])


def one(fn, *args):
    """按「一行数据」调用一个函数，返回标量。"""
    rt._NROW = 1
    out = fn(*args)
    return out.item() if hasattr(out, "item") and np.ndim(out) == 0 else out


def arr(fn, n, *args):
    """按「n 行数据」调用，返回当前行数下正常结果的列表。"""
    rt._NROW = n
    out = fn(*args)
    return np.atleast_1d(out)


# ================================================================ 舍入 / 数学


class TestRounding(unittest.TestCase):
    def test_round_is_away_from_zero(self):
        # Excel 的 .5 一律远离零；Python 内置 round 会走银行家舍入（2.5 -> 2）
        self.assertEqual(one(rt._round, 2.5), 3.0)
        self.assertEqual(one(rt._round, -2.5), -3.0)
        self.assertEqual(one(rt._round, 1.2345, 2), 1.23)
        self.assertEqual(one(rt._round, -1.235, 2), -1.24)

    def test_roundup_rounddown(self):
        self.assertEqual(one(rt._roundup, 1.001, 2), 1.01)
        self.assertEqual(one(rt._rounddown, 1.009, 2), 1.00)
        self.assertEqual(one(rt._roundup, -1.001, 2), -1.01)
        self.assertEqual(one(rt._rounddown, -1.009, 2), -1.00)

    def test_ceiling_away_from_zero(self):
        self.assertEqual(one(rt._ceiling, 2.1, 1), 3.0)
        self.assertEqual(one(rt._ceiling, 5, 5), 5.0)
        # 回归点：负数必须「远离零」。写成 np.ceil(V/N)*N 会得到 -2，Excel 是 -3
        self.assertEqual(one(rt._ceiling, -2.1, 1), -3.0)
        self.assertEqual(one(rt._ceiling, 0, 5), 0.0)

    def test_floor_toward_zero(self):
        self.assertEqual(one(rt._floor_, 2.9, 1), 2.0)
        # 回归点：负数必须「向零」。写成 np.floor(V/N)*N 会得到 -3，Excel 是 -2
        self.assertEqual(one(rt._floor_, -2.1, 1), -2.0)

    def test_mround(self):
        self.assertEqual(one(rt._mround, 10, 3), 9.0)
        self.assertEqual(one(rt._mround, 1.5, 1), 2.0)
        self.assertEqual(one(rt._mround, -1.5, -1), -2.0)
        # 符号跟随被舍入的数：Excel 给 -9，不是 +9
        self.assertEqual(one(rt._mround, -10, -3), -9.0)
        self.assertEqual(one(rt._mround, 0, 5), 0.0)
        # 异号 -> #NUM!（NaN）。Excel 的 MROUND 要求两个参数同号，
        # 所以 MROUND(-1.5, 1) 也是 #NUM!，而不是 -2
        self.assertTrue(math.isnan(one(rt._mround, -10, 3)))
        self.assertTrue(math.isnan(one(rt._mround, -1.5, 1)))
        # 十进制半值：6.05/0.1 在双精度下是 60.4999…，靠 1e-9 纠回 Excel 的 6.1
        self.assertAlmostEqual(one(rt._mround, 6.05, 0.1), 6.1, places=10)

    def test_div_zero(self):
        self.assertTrue(math.isnan(one(rt._div, 1, 0)))
        self.assertEqual(one(rt._div, 6, 3), 2.0)


# ================================================================ 日期


class TestDates(unittest.TestCase):
    def test_excel_serial(self):
        # 已知锚点：2024-01-01 = 45292（Excel 序列号）
        self.assertEqual(dstr(rt._d1(45292)), "2024-01-01")
        self.assertEqual(dstr(rt._d1(46023)), "2026-01-01")
        self.assertEqual(dstr(rt._d1("2026-01-01")), "2026-01-01")

    def test_date(self):
        rt._NROW = 1
        self.assertEqual(dstr(rt._date(2026, 1, 15)[0]), "2026-01-15")
        # Excel 的月/日越界会滚动
        self.assertEqual(dstr(rt._date(2026, 13, 1)[0]), "2027-01-01")
        self.assertEqual(dstr(rt._date(2026, 1, 32)[0]), "2026-02-01")

    def test_datepart(self):
        rt._NROW = 1
        self.assertEqual(one(rt._datepart, "2026-03-15", "year"), 2026.0)
        self.assertEqual(one(rt._datepart, "2026-03-15", "month"), 3.0)
        self.assertEqual(one(rt._datepart, "2026-03-15", "day"), 15.0)

    def test_eomonth(self):
        self.assertEqual(d1(rt._eomonth, "2026-01-15", 0), "2026-01-31")
        self.assertEqual(d1(rt._eomonth, "2026-01-31", 1), "2026-02-28")
        self.assertEqual(d1(rt._eomonth, "2024-01-31", 1), "2024-02-29")   # 闰年
        self.assertEqual(d1(rt._eomonth, "2026-03-15", -1), "2026-02-28")
        self.assertEqual(d1(rt._eomonth, "2026-12-15", 1), "2027-01-31")    # 跨年

    def test_edate(self):
        self.assertEqual(d1(rt._edate, "2026-01-15", 2), "2026-03-15")
        # 目标月没有该日 -> 取月末
        self.assertEqual(d1(rt._edate, "2026-01-31", 1), "2026-02-28")
        self.assertEqual(d1(rt._edate, "2024-01-31", 1), "2024-02-29")
        self.assertEqual(d1(rt._edate, "2026-03-31", -1), "2026-02-28")

    def test_days360(self):
        # 美式（NASD）：起始 1/1 早于 30 号，终止 12/31 是月末 -> 变成次年 1/1
        self.assertEqual(one(rt._days360, "2026-01-01", "2026-12-31"), 360.0)
        self.assertEqual(one(rt._days360, "2026-01-01", "2026-07-01"), 180.0)
        # 欧式：两端的 31 号都压成 30 号
        self.assertEqual(one(rt._days360, "2026-01-01", "2026-12-31", True), 359.0)

    def test_yearfrac(self):
        self.assertAlmostEqual(one(rt._yearfrac, "2026-01-01", "2026-07-01", 0), 0.5)
        # 2026-01-01 -> 2026-07-01 实际 181 天
        self.assertAlmostEqual(one(rt._yearfrac, "2026-01-01", "2026-07-01", 2), 181 / 360)
        self.assertAlmostEqual(one(rt._yearfrac, "2026-01-01", "2026-07-01", 3), 181 / 365)

    def test_yearfrac_basis1_is_loud_not_silent(self):
        # basis=1 刻意不实现：宁可报错，也不要静默算错
        with self.assertRaises(NotImplementedError):
            one(rt._yearfrac, "2026-01-01", "2026-07-01", 1)

    def test_networkdays(self):
        # 2026-01-01(四) ~ 01-07(三)：四、五、一、二、三 = 5 天
        self.assertEqual(one(rt._networkdays, "2026-01-01", "2026-01-07"), 5.0)
        self.assertEqual(
            one(rt._networkdays, "2026-01-01", "2026-01-07", ["2026-01-05"]), 4.0)
        # 起止颠倒 -> 负数（与 Excel 一致）
        self.assertEqual(one(rt._networkdays, "2026-01-07", "2026-01-01"), -5.0)

    def test_workday(self):
        self.assertEqual(d1(rt._workday, "2026-01-01", 1), "2026-01-02")   # 周四 -> 周五
        self.assertEqual(d1(rt._workday, "2026-01-02", 1), "2026-01-05")   # 周五 -> 下周一
        self.assertEqual(d1(rt._workday, "2026-01-05", -1), "2026-01-02")  # 周一 -> 上周五
        self.assertEqual(
            d1(rt._workday, "2026-01-02", 1, ["2026-01-05"]), "2026-01-06")  # 跳过节假日

    def test_weeknum(self):
        # type=1 一周从周日开始；type=2 从周一开始；type=21 是 ISO 8601
        self.assertEqual(one(rt._weeknum, "2026-01-01", 1), 1.0)
        self.assertEqual(one(rt._weeknum, "2026-01-04", 1), 2.0)   # 那年的第一个周日
        self.assertEqual(one(rt._weeknum, "2026-01-05", 2), 2.0)   # 那年的第一个周一
        self.assertEqual(one(rt._weeknum, "2026-01-01", 21), 1.0)
        self.assertEqual(one(rt._weeknum, "2026-01-04", 21), 1.0)  # ISO 第 1 周含 1/4

    def test_datetime_roundtrip_vectorized(self):
        rt._NROW = 3
        out = rt._eomonth(pd.to_datetime(["2026-01-05", "2026-02-05", "2026-03-05"]), 0)
        self.assertEqual([dstr(v) for v in out],
                         ["2026-01-31", "2026-02-28", "2026-03-31"])


# ================================================================ 文本


class TestText(unittest.TestCase):
    def test_text_number(self):
        self.assertEqual(one(rt._text, 1234.5, "#,##0.00"), "1,234.50")
        self.assertEqual(one(rt._text, 1234.4, "0"), "1234")
        self.assertEqual(one(rt._text, -1234.5, "#,##0.00"), "-1,234.50")
        # .5 走 Excel 的远离零，Python 的 f-string 会给 "1234"
        self.assertEqual(one(rt._text, 1234.5, "0"), "1235")

    def test_text_percent(self):
        self.assertEqual(one(rt._text, 0.125, "0.0%"), "12.5%")
        self.assertEqual(one(rt._text, 0.126, "0%"), "13%")

    def test_text_date(self):
        self.assertEqual(one(rt._text, "2026-01-31", "yyyy-mm-dd"), "2026-01-31")
        self.assertEqual(one(rt._text, "2026-03-05", "yyyy年m月"), "2026年3月")
        self.assertEqual(one(rt._text, "2026-03-05", "yyyy/m/d"), "2026/3/5")

    def test_textjoin(self):
        rt._NROW = 1
        self.assertEqual(one(rt._textjoin, "-", True, ["a", "", "b"]), "a-b")
        self.assertEqual(one(rt._textjoin, "-", False, ["a", "", "b"]), "a--b")
        self.assertEqual(one(rt._textjoin, ",", True, ["a", "b"]), "a,b")

    def test_substitute(self):
        self.assertEqual(one(rt._substitute, "a-b-c", "-", "+"), "a+b+c")
        self.assertEqual(one(rt._substitute, "a-b-c", "-", "+", 2), "a-b+c")
        self.assertEqual(one(rt._substitute, "aaa", "a", "b", 1), "baa")

    def test_proper(self):
        self.assertEqual(one(rt._proper, "hello world"), "Hello World")
        # Excel 把数字也当分隔符：PROPER("2nd place") = "2Nd Place"
        self.assertEqual(one(rt._proper, "2nd place"), "2Nd Place")

    def test_exact_case_sensitive(self):
        self.assertFalse(one(rt._exact, "abc", "ABC"))
        self.assertTrue(one(rt._exact, "abc", "abc"))

    def test_textbefore_textafter(self):
        rt._NROW = 1
        self.assertEqual(one(rt._textbefore, "a@b.com", "@"), "a")
        self.assertEqual(one(rt._textafter, "a@b.com", "@"), "b.com")


# ================================================================ 其它既有语义


class TestCore(unittest.TestCase):
    def test_if_scalar_and_array(self):
        rt._NROW = 3
        self.assertEqual(one(rt._if, True, "x", "y"), "x")
        out = rt._if(np.array([True, False, True]), "x", "y")
        self.assertEqual(list(out), ["x", "y", "x"])

    def test_clean_error_strings(self):
        s = pd.Series(["1", "#DIV/0!", "#N/A"])
        self.assertEqual(int(rt._clean(s).isna().sum()), 2)

    def test_aggif_sumifs_no_match_is_zero(self):
        rt._NROW = 1
        src = pd.DataFrame({"k": ["a", "a", "b"], "v": [1.0, 2.0, 3.0]})
        # 无匹配时 SUMIFS 返回 0（不是 NaN），AVERAGEIFS 返回 NaN
        self.assertEqual(rt._aggif(src, "B", [("A", "z")], "sum")[0], 0.0)
        self.assertTrue(math.isnan(rt._aggif(src, "B", [("A", "z")], "mean")[0]))
        self.assertEqual(rt._aggif(src, "B", [("A", "a")], "sum")[0], 3.0)

    def test_lookup_first_match_and_not_found(self):
        rt._NROW = 3
        src = pd.DataFrame({"k": ["a", "a", "b"], "v": [1.0, 9.0, 3.0]})
        # 键重复时取首次命中（Excel 的 VLOOKUP/XLOOKUP 语义），而不是报错或取最后一个
        out = rt._lookup(src, "A", "B", ["a", "b", "zzz"], not_found=None)
        self.assertEqual(float(out[0]), 1.0)
        self.assertEqual(float(out[1]), 3.0)
        self.assertTrue(pd.isna(out[2]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
