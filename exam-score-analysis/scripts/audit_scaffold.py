# -*- coding: utf-8 -*-
"""独立核对骨架（Phase 5）

用法：复制本文件到项目 scripts/ 下，改 TODO 段落，即可得到一个可复跑的核对脚本。

设计约束（不要改）：
1. 只读原始成绩包与成品文件，**不读生成阶段的中间 JSON**；
2. 数值比到单元格，容差默认 0.005；
3. 三类差异单列，不计入"不符"：同分并列行序、x.xx5 边界、显示值与原值相减；
4. 结论三态：全部通过 / 存在 n 项待修 / 存在 n 项需用户裁决。

报告示例：
    合计 1406 项核对
    - 不符：2 项
      [PPT] p23 临界生人数 现值=18 复算=17  差异：口径未同步（-15~-10 开闭区间）
    - 精度备注：1 条
    - 白名单：2 条
"""
from decimal import Decimal, ROUND_HALF_UP
import sys


def r2(x):
    if x is None:
        return None
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


class Checker:
    """核对器。所有断言都带位置描述，报告可直接定位到格子。"""

    def __init__(self, tol=0.005):
        self.tol = tol
        self.total = 0
        self.issues = []      # (分类, 位置, 现值, 复算值, 原因)
        self.notes = []       # 精度备注
        self.white = []       # 白名单（已知合理差异）
        self.scope = ""

    # ---------- 基础断言
    def num(self, where, cur, calc, tol=None, note=""):
        """数值比对。cur=成品现值，calc=独立复算值。"""
        t = self.tol if tol is None else tol
        self.total += 1
        cur, calc = r2(cur), r2(calc)
        if cur is None and calc is None:
            return True
        if cur is None or calc is None:
            self.issues.append(("缺失", where, cur, calc,
                                note or "一方为空"))
            return False
        d = abs(cur - calc)
        if d <= t:
            return True
        if d <= t + 0.005 and _is_half_edge(cur, calc):
            self.notes.append((where, cur, calc, "x.xx5 边界舍入差"))
            return True
        self.issues.append(("数值", where, cur, calc, note or f"差 {d:.4f}"))
        return False

    def text(self, where, cur, expect, note=""):
        """文本比对（去首尾空白与全角空格）。"""
        self.total += 1
        a = "" if cur is None else str(cur).replace("\u3000", " ").strip()
        b = "" if expect is None else str(expect).replace("\u3000", " ").strip()
        if a == b:
            return True
        self.issues.append(("文字", where, a, b, note))
        return False

    def contain(self, where, cur, needle, note=""):
        """包含断言：用于结论句中关键量是否出现。"""
        self.total += 1
        a = "" if cur is None else str(cur)
        if needle and str(needle) in a:
            return True
        self.issues.append(("文字缺失", where, a[:60], str(needle),
                            note or "结论未包含预期关键量"))
        return False

    def near(self, where, cur, calc, tol, note=""):
        """宽容差比对，用于百分比/占比类。"""
        return self.num(where, cur, calc, tol=tol, note=note)

    def skip(self, where, reason):
        """登记白名单（不计入总数）。"""
        self.white.append((where, reason))

    # ---------- 报告
    def report(self, title=""):
        lines = []
        if title:
            lines.append(title)
        lines.append(f"合计 {self.total} 项核对")
        lines.append(f"- 不符：{len(self.issues)} 项")
        for kind, where, cur, calc, note in self.issues:
            lines.append(f"  [{kind}] {where} 现值={cur!r} 复算={calc!r}  {note}")
        lines.append(f"- 精度备注：{len(self.notes)} 条")
        for where, cur, calc, note in self.notes:
            lines.append(f"  {where} {cur} vs {calc}  {note}")
        lines.append(f"- 白名单：{len(self.white)} 条")
        for where, reason in self.white:
            lines.append(f"  {where}  {reason}")
        if not self.issues:
            verdict = "全部通过"
        elif all(k == "备注" for k, *_ in self.issues):
            verdict = "全部通过（仅有备注）"
        else:
            verdict = f"存在 {len(self.issues)} 项待修"
        lines.append(f"结论：{verdict}")
        return "\n".join(lines)

    def save(self, path, title=""):
        txt = self.report(title)
        with open(path, "w", encoding="utf-8") as f:
            f.write(txt + "\n")
        return txt

    def exit_code(self):
        return 0 if not [i for i in self.issues if i[0] != "备注"] else 1


def _is_half_edge(a, b):
    """判断两个两位小数是否是同一个 x.xx5 母值的两种舍入结果。"""
    if a is None or b is None:
        return False
    s = (a + b) / 2
    return abs(abs(a - b) - 0.01) < 1e-9 and abs(s * 1000 - round(s * 1000)) < 0.5


# ============================================================ 使用示例
if __name__ == "__main__":
    c = Checker()

    # ---- TODO 1: 载入原始数据（不要读生成阶段中间文件）
    # from reader import read_any, num, cls_norm, wavg, r2
    # IND = read_any(".../总分各项指标对比.xlsx", "总分")

    # ---- TODO 2: 按锁定口径独立重算
    # calc_total_avg = wavg([(num(r[12]), num(r[2])) for r in IND[3:33]])
    # calc_rank = ...

    # ---- TODO 3: 与成品逐格比对
    # c.num("汇总/总分排名!F5", 340.49, calc_total_avg)
    # c.text("PPT/p3 表头", "30校排名", "30校排名")

    # ---- TODO 4: 核对结论句中的关键量（最易出错，必做）
    # c.contain("PPT/p6 结论", slide_text, "27.30%")

    # ---- TODO 5: 登记白名单
    # c.skip("汇总/九科对比!D7", "423.5 同分并列，行序差异，数值一致")

    print(c.report("独立核对报告"))
    sys.exit(c.exit_code())
