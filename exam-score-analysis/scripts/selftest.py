# -*- coding: utf-8 -*-
"""安装自检：依赖是否齐全 + 三个脚本的核心函数是否正确。

任何宿主（Codex / Claude Code / WorkBuddy / Gemini CLI …）装好本技能后，
先跑这一句确认可用：

    python3 scripts/selftest.py

全部 OK 才继续；有 FAIL 按提示补依赖。
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

FAILED = []


def check(name, fn):
    try:
        fn()
        print(f"  OK   {name}")
    except Exception as e:
        FAILED.append((name, repr(e)))
        print(f"  FAIL {name}: {e!r}")


# ---------------------------------------------------------------- 依赖
def dep_openpyxl():
    import openpyxl  # noqa: F401


def dep_xlrd():
    import xlrd  # noqa: F401


def dep_pptx():
    import pptx  # noqa: F401


def dep_lxml():
    import lxml  # noqa: F401


# ---------------------------------------------------------------- 逻辑
def logic_round():
    from reader import r2
    assert r2(21.625) == 21.63, "ROUND_HALF_UP 失效：21.625 应为 21.63"
    assert r2(19.45) == 19.45


def logic_wavg():
    from reader import wavg
    assert abs(wavg([(80, 10), (90, 30)]) - 87.5) < 1e-9


def logic_rank():
    from reader import rank_desc
    got = rank_desc({"a": 90, "b": 90, "c": 80, "d": 80, "e": 70})
    assert got == {"a": 1, "b": 1, "c": 3, "d": 3, "e": 5}, f"竞赛式排名错误：{got}"


def logic_cls():
    from reader import cls_norm
    assert cls_norm("高三（314）班") == "314"
    assert cls_norm("年级") is None


def logic_num():
    from reader import num
    assert num("缺考") is None and num("380.08") == 380.08


def logic_mode():
    from reader import detect_mode
    with tempfile.TemporaryDirectory() as d:
        os.makedirs(os.path.join(d, "高三物理类（共性分析）", "2.各校成绩分析"))
        r = detect_mode(d)
        assert r["mode"] == "joint", r
    with tempfile.TemporaryDirectory() as d:
        os.makedirs(os.path.join(d, "总分"))
        r = detect_mode(d)
        assert r["mode"] == "single", r


def logic_count_schools():
    from reader import count_schools
    rows = [
        ["某某考试总分基础指标对比"],
        ["学校代码", "学校", "统计人数"],
        ["学校代码", "学校", "统计人数"],
        ["", "甲市第一中学", 25],
        ["1135", "乙市第一中学", 176],
        ["1702", "丙市实验中学", 120],
        ["", "整体", 8265],
    ]
    n, names = count_schools(rows)
    assert n == 3, f"校数应为 3，实得 {n}：{names}"
    assert "整体" not in names and "学校代码" not in names, names


def logic_audit():
    from audit_scaffold import Checker
    c = Checker()
    assert c.num("x", 340.49, 340.49)
    assert not c.num("y", 340.49, 341.00)
    c.text("t", "  30校排名 ", "30校排名")
    assert c.total >= 3 and len(c.issues) == 1


def logic_patch_help():
    """xlsx_patch 至少能被导入且暴露三个子命令。"""
    import xlsx_patch as xp
    assert hasattr(xp, "cmd_patch") and hasattr(xp, "find_cell") and hasattr(xp, "enc")
    assert xp.dec(xp.enc("赋分")) == "赋分"


def main():
    print("【依赖】")
    for n, f in [("openpyxl", dep_openpyxl), ("xlrd", dep_xlrd),
                 ("python-pptx", dep_pptx), ("lxml", dep_lxml)]:
        check(n, f)

    print("【reader】")
    for n, f in [("ROUND_HALF_UP", logic_round), ("加权均值", logic_wavg),
                 ("竞赛式排名", logic_rank), ("班级归一化", logic_cls),
                 ("数值解析", logic_num), ("联考/单校判定", logic_mode),
                 ("参评校数统计", logic_count_schools)]:
        check(n, f)

    print("【核对与修补】")
    check("Checker 断言", logic_audit)
    check("xlsx_patch 实体编解码", logic_patch_help)

    print()
    if FAILED:
        print(f"结论：{len(FAILED)} 项失败 →")
        for n, e in FAILED:
            print(f"  - {n}: {e}")
        print("\n缺依赖时执行：python3 -m pip install -r requirements.txt")
        return 1
    print("结论：全部通过，本技能在当前环境可用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
