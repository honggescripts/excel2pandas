# -*- coding: utf-8 -*-
"""
自包含性验证 —— 生成的代码能不能**脱离本项目**独立运行。

    python tests/test_selfcontained.py

为什么单独测这一条：
  「按需嵌入」最容易犯的错，是把 helper 改成从本包 import
  （`from excel2pandas._runtime import _round`）。那样在仓库里跑一切正常，
  但用户把 generated_code.py 单独拿走就崩 —— 而用户拿到文件就走的场景，
  正是这个项目存在的理由。

这里做三层验证：
  ① 源码层面：生成文件里不出现任何对本包的 import
  ② 解释器层面：在**本包不可导入**的环境里执行（cwd 隔离 + PYTHONPATH 清空）
  ③ 结果层面：独立环境跑出来的值与仓库内跑出来的一致
"""

import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from excel2pandas.generator import generate  # noqa: E402

SHEET = "定价测算"
KEYS = ["SKU-10001", "SKU-10002", "SKU-10003"]

CHILD = '''\
# -*- coding: utf-8 -*-
"""在「本包不可导入」的环境里运行生成代码。"""
import importlib.util
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

# ① 先证明本包确实不可导入 —— 否则这次验证没有说服力
try:
    import excel2pandas  # noqa: F401
    print(json.dumps({"error": "excel2pandas 竟然可导入，自包含性无法证明"}))
    raise SystemExit(3)
except ImportError:
    pass

# ② 直接按路径加载生成的模块（不依赖任何包安装）
spec = importlib.util.spec_from_file_location("generated_code", HERE / "generated_code.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

df = mod.calc(%(keys)r)
print(json.dumps({
    "rows": int(len(df)),
    "cols": int(len(df.columns)),
    "毛利率": [round(float(v), 6) for v in df["毛利率"]],
    "建议售价": [round(float(v), 6) for v in df["建议售价"]],
}, ensure_ascii=False))
'''


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class TestSelfContained(unittest.TestCase):
    def test_generated_code_runs_without_the_package(self):
        # Windows 上 pandas 会持有 xlsx 句柄，临时目录清理要容错
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tmp = Path(td)

            # 把模板也拷进去 —— 模拟「用户只拿到一个文件 + 自己的 Excel」
            tpl = tmp / "demo_template.xlsx"
            shutil.copy2(ROOT / "examples" / "demo_template.xlsx", tpl)

            # ① 在这里生成，产物落到临时目录
            rep = generate(excel_path=tpl, sheet_name=SHEET, target_row=2,
                           out_dir=tmp, force=True, runtime_mode="minimal")
            code_path = Path(rep.code_path)
            self.assertTrue(code_path.exists())

            # ② 源码层面：不得有任何对本包的引用
            src = code_path.read_text(encoding="utf-8")
            for bad in ("from excel2pandas", "import excel2pandas"):
                self.assertNotIn(bad, src, f"生成代码里出现了 {bad!r}，破坏了自包含")

            # ③ 解释器层面：cwd 隔离 + 清空 PYTHONPATH 后运行
            child = tmp / "_run_child.py"
            child.write_text(CHILD % {"keys": KEYS}, encoding="utf-8")
            env = dict(os.environ)
            env["PYTHONPATH"] = ""
            env["PYTHONIOENCODING"] = "utf-8"
            p = subprocess.run([sys.executable, str(child)], cwd=str(tmp), env=env,
                               capture_output=True, text=True, encoding="utf-8",
                               errors="replace", timeout=300)
            self.assertEqual(p.returncode, 0,
                             f"独立环境运行失败：\n{p.stdout}\n{p.stderr}")
            got = json.loads(p.stdout.strip().splitlines()[-1])

            # ④ 结果层面：与仓库内跑出来的逐值一致
            mod = _load(code_path, "gen_in_repo")
            df = mod.calc(KEYS)
            self.assertEqual(got["rows"], len(df))
            self.assertEqual(got["cols"], len(df.columns))
            self.assertEqual(got["毛利率"], [round(float(v), 6) for v in df["毛利率"]])
            self.assertEqual(got["建议售价"], [round(float(v), 6) for v in df["建议售价"]])

            # ⑤ 顺带确认这次跑的就是裁剪后的 minimal 产物
            self.assertEqual(rep.runtime_mode, "minimal")
            self.assertGreater(len(rep.runtime_helpers), 0)

            # 放掉 pandas 持有的 xlsx 句柄，否则临时目录删不掉（Windows）
            h = getattr(mod, "_XLSX", None)
            if h is not None:
                h.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
