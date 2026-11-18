# -*- coding: utf-8 -*-
"""通用读取与归一化工具（联考成绩分析流水线 Phase 1）

设计目标：让"读表"这件事不再成为每次任务的成本。
- 自动兼容 OLE2 伪装的 .xlsx
- 统一数值解析（含缺考占位符）
- 班级名归一化
- 按列名定位（避免数据源改列序后全面错位）
- 结构指纹（数据源版本漂移时先比指纹）
- 加权均值与 ROUND_HALF_UP 舍入

用法：
    from reader import read_any, num, cls_norm, header_index, cols, wavg, r2, fingerprint
"""
import os
import re
import hashlib
import collections
from decimal import Decimal, ROUND_HALF_UP

# 缺考/空值占位符：出现即视为无有效分数
NULL_TOKENS = {"", "-", "—", "——", "/", "\\", "缺考", "缺", "无卡", "未考",
               "None", "none", "NULL", "null", "nan", "NaN"}


# ---------------------------------------------------------------- 读取
def read_any(path, sheet=None):
    """读取任意 Excel（.xls / .xlsx / 伪装成 .xlsx 的 .xls），返回 list[list]。

    - sheet 为 None 时取第一张表；不存在则回退第一张
    - 优先 openpyxl，失败（BadZipFile / InvalidFileException）回退 xlrd
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    except Exception:
        import xlrd
        book = xlrd.open_workbook(path)
        sh = (book.sheet_by_name(sheet)
              if sheet and sheet in book.sheet_names() else book.sheet_by_index(0))
        return [[sh.cell_value(i, c) for c in range(sh.ncols)] for i in range(sh.nrows)]


def sheet_names(path):
    """返回工作表名列表（同样兼容 OLE2）。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        import xlrd
        return list(xlrd.open_workbook(path).sheet_names())


# ---------------------------------------------------------------- 归一化
def num(v):
    """转 float；缺考/空值/占位符返回 None。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in NULL_TOKENS:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def cls_norm(v):
    """班级名归一化：'314班' / '高314班' / '高三（314）班' → '314'。"""
    if v is None:
        return None
    m = re.search(r"(\d{2,3})", str(v))
    return m.group(1) if m else None


def is_summary_row(value):
    """汇总行判据：班级列不是 2–3 位数字即视为'年级/整体/合计'等汇总行。"""
    return cls_norm(value) is None


def text(v):
    """安全转字符串并去首尾空白与全角空格。"""
    if v is None:
        return ""
    return str(v).replace("\u3000", " ").strip()


# ---------------------------------------------------------------- 列定位
def header_index(rows, names, max_scan=8, exact=False):
    """在表头区（前 max_scan 行）找列名为 names 之一的列号。

    names 可为 str 或 list。返回 (header_row, col_index)；找不到返回 (-1, -1)。
    优先精确匹配，其次包含匹配。表头常跨 2–3 行合并，故逐行扫描。
    """
    keys = [names] if isinstance(names, str) else list(names)
    for r in range(min(max_scan, len(rows))):
        row = [text(c) for c in rows[r]]
        for i, cell in enumerate(row):
            if not cell:
                continue
            for k in keys:
                if cell == k:
                    return r, i
    if exact:
        return -1, -1
    for r in range(min(max_scan, len(rows))):
        row = [text(c) for c in rows[r]]
        for i, cell in enumerate(row):
            for k in keys:
                if k and k in cell:
                    return r, i
    return -1, -1


def cols(rows, mapping, header_row=0):
    """按 {字段: 列名} 建索引，返回 {字段: 列号}。缺列直接报错，避免静默错位。"""
    out, missing = {}, []
    header = [text(c) for c in rows[header_row]] if rows else []
    for field, name in mapping.items():
        keys = [name] if isinstance(name, str) else list(name)
        hit = -1
        for k in keys:
            if k in header:
                hit = header.index(k)
                break
        if hit < 0:                      # 回退：包含匹配
            for i, cell in enumerate(header):
                if any(k and k in cell for k in keys):
                    hit = i
                    break
        if hit < 0:
            missing.append(field)
        out[field] = hit
    if missing:
        raise KeyError(f"表头缺少必要列 {missing}；实际表头={header}")
    return out


def get(row, idx, default=None):
    """按列号安全取值。"""
    if idx is None or idx < 0 or idx >= len(row):
        return default
    return row[idx]


def getnum(row, idx):
    return num(get(row, idx))


# ---------------------------------------------------------------- 计算
def r2(x, digits=2):
    """教务系统口径的四舍五入（ROUND_HALF_UP）。

    注意：不要用内置 round()——它是银行家舍入，x.xx5 时与系统差 0.01。
    """
    if x is None:
        return None
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(str(x)).quantize(q, rounding=ROUND_HALF_UP))


def r2s(x, digits=2):
    """四舍五入并格式化为定长字符串（用于与显示值比对）。"""
    v = r2(x, digits)
    return "" if v is None else f"{v:.{digits}f}"


def wavg(pairs):
    """加权平均：pairs = [(值, 人数), ...]，自动跳过 None 与 0 人。"""
    sv = sn = 0.0
    for v, w in pairs:
        if v is None:
            continue
        w = num(w) or 0
        if w <= 0:
            continue
        sv += v * w
        sn += w
    return sv / sn if sn else None


def pct(part, whole, digits=2):
    """百分比（返回数值，不带 % 号）。"""
    if not whole:
        return None
    return r2(part / whole * 100, digits)


def rank_desc(values):
    """降序名次（竞赛式排名，同分并列后跳号：90,90,80 → 1,1,3）。

    values: {key: 分数}。前 N 名统计需用此口径——并列占位，后续名次顺延。
    """
    present = {k: x for k, x in values.items() if x is not None}
    counts = collections.Counter(present.values())
    base, out = 0, {}
    for x in sorted(counts, reverse=True):
        out[x] = base + 1
        base += counts[x]
    return {k: out[x] for k, x in present.items()}


# ---------------------------------------------------------------- 模式判定
# 不是每次考试都有外校成绩。先判联考/单校，再决定做不做校际对比。
SCHOOL_KEYS = ["学校", "参评校", "校名", "学校名称", "单位"]
SUMMARY_WORDS = ["整体", "合计", "总计", "全年级", "年级平均"]
JOINT_DIR_HINTS = ["各校成绩分析", "各项指标对比", "共性分析", "校际", "各校对比", "各校"]


def _school_col(rows, max_scan=6):
    """找"学校名"所在列：优先精确 '学校'/'学校名称'，明确排除"学校代码"。

    指标表常见结构：第 1 行大标题、第 2–3 行重复表头、第 4 行起数据。
    返回 (header_row, col)；找不到返回 (-1, -1)。
    """
    best = (0, -1, -1)
    for r in range(min(max_scan, len(rows))):
        for i, cell in enumerate(text(c) for c in rows[r]):
            if not cell:
                continue
            if cell in ("学校代码", "代码", "学校编号", "编号") or "代码" in cell:
                continue
            if cell in ("学校", "学校名称", "参评校", "校名", "学校名"):
                score = 5
            elif "学校名称" in cell:
                score = 4
            elif any(k in cell for k in ("学校", "参评校", "校名")):
                score = 2
            else:
                continue
            if score > best[0]:
                best = (score, r, i)
    return (best[1], best[2]) if best[0] else (-1, -1)


def count_schools(rows, max_scan=6):
    """统计指标表参评校数 N（去重），用于替换"写死 30"。

    返回 (N, [校名...])；找不到"学校"列或只有表头无数据时返回 (0, [])。
    "整体/合计/总计"等汇总行不计入；重复表头行自动跳过。
    """
    hr, ci = _school_col(rows, max_scan)
    if hr < 0:
        return 0, []
    names, seen_header = [], False
    for row in rows[hr + 1:]:
        nm = text(get(row, ci))
        if not nm:
            continue
        if nm in ("学校", "学校名称", "参评校", "校名", "学校名") or "代码" in nm:
            seen_header = True          # 重复表头行，跳过
            continue
        if any(w in nm for w in SUMMARY_WORDS):
            continue
        if not seen_header and nm in names:
            continue
        if nm not in names:
            names.append(nm)
    return len(names), names


def detect_mode(root_dir, max_depth=3):
    """粗判考试类型：目录里有"各校/共性分析"类子目录 → 联考，否则单校。

    返回 {"mode": "joint"|"single", "evidence": [...命中目录...], "hint": str}
    注意：这是**粗判**。联考还须用 count_schools() 确认指标表里确有 ≥2 所学校；
    单校则反过来确认全校只有一所学校的数据。判据三条见 references/exam-mode.md。
    """
    root_dir = os.path.abspath(root_dir)
    evidence = []
    base_depth = root_dir.rstrip(os.sep).count(os.sep)
    for cur, dirs, _files in os.walk(root_dir):
        if cur.count(os.sep) - base_depth >= max_depth:
            dirs[:] = []
            continue
        for d in dirs:
            if any(h in d for h in JOINT_DIR_HINTS):
                evidence.append(os.path.join(cur, d).replace(root_dir + os.sep, ""))
    mode = "joint" if evidence else "single"
    hint = ("联考：检测到各校对比类目录，需再用 count_schools() 确认校数 N"
            if evidence else
            "单校：未检测到各校对比目录，按单校模式（从班级分析起），不做校际对比")
    return {"mode": mode, "evidence": evidence, "hint": hint}


# ---------------------------------------------------------------- 结构指纹
def fingerprint(path):
    """数据源结构指纹：文件 sha256 前 12 位 + 表名 + 每表前 4 行文字摘要。

    用途：数据源被重生成时，先比指纹判断"变了什么"，再决定重做范围。
    """
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    fp = {"file": os.path.basename(path), "sha256_12": h.hexdigest()[:12], "sheets": {}}
    for name in sheet_names(path):
        rows = read_any(path, name)
        head = []
        for r in rows[:4]:
            head.append([text(c) for c in r if text(c)][:12])
        fp["sheets"][name] = {"rows": len(rows), "head": head}
    return fp


def diff_fingerprint(old, new):
    """比较两份指纹，返回人类可读的差异描述列表（空列表=结构未变）。"""
    out = []
    if old.get("sha256_12") != new.get("sha256_12"):
        out.append(f"文件哈希变化 {old.get('sha256_12')} → {new.get('sha256_12')}")
    so, sn = old.get("sheets", {}), new.get("sheets", {})
    for k in so.keys() - sn.keys():
        out.append(f"工作表被删除：{k}")
    for k in sn.keys() - so.keys():
        out.append(f"工作表新增：{k}")
    for k in so.keys() & sn.keys():
        if so[k]["rows"] != sn[k]["rows"]:
            out.append(f"{k} 行数 {so[k]['rows']} → {sn[k]['rows']}")
        if so[k]["head"] != sn[k]["head"]:
            out.append(f"{k} 表头变化")
    return out


if __name__ == "__main__":
    import sys
    args = sys.argv[1:]
    if not args:
        print("用法:\n"
              "  python3 reader.py --mode <数据包目录>        # 判联考/单校\n"
              "  python3 reader.py <xlsx路径> [工作表名]     # 预览前 5 行")
        raise SystemExit(1)
    if args[0] == "--mode":
        res = detect_mode(args[1])
        print(f"mode = {res['mode']}")
        print(f"hint = {res['hint']}")
        for e in res["evidence"][:8]:
            print("  命中：", e)
        raise SystemExit(0)
    p = args[0]
    for nm in (args[1:2] or sheet_names(p)):
        rows = read_any(p, nm)
        print(f"[{nm}] {len(rows)} 行")
        for r in rows[:5]:
            print("   ", [text(c) for c in r][:10])
