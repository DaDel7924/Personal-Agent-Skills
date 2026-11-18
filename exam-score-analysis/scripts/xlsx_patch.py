# -*- coding: utf-8 -*-
"""单元格级 xlsx 修补（保留图表、条件格式、筛选、冻结窗格）

为什么不用 openpyxl：load_workbook + save 会丢失 xl/charts/*、部分条件格式与
drawing 引用。本工具直接改写 xl/worksheets/sheetN.xml，其余 zip 条目按原
date_time / compress_type 逐字节复制，保存后条目数、图表、关系文件全部不变。

特性：
- 默认 dry-run，只报告将改什么；加 --apply 才落盘
- 旧文必须**唯一命中**，命中 0 次或多次直接报错退出（禁止盲改）
- 中文以 &#NNNN; 实体存储时自动解码比对、编码写回
- 落盘前后自动校验：zip 条目数不变、chart 仍在、openpyxl 可读

用法：
    python3 xlsx_patch.py list  book.xlsx
    python3 xlsx_patch.py dump  book.xlsx "17 得分率差明细" K160
    python3 xlsx_patch.py patch book.xlsx "17 得分率差明细" K160 --old " " --new "22" --apply
    python3 xlsx_patch.py patch book.xlsx "说明" B11 --old "9月 439" --new "9月 438"
"""
import argparse
import re
import shutil
import sys
import zipfile
from datetime import datetime

ENT_RE = re.compile(r"&#(\d+);")


def enc(s):
    """把非 ASCII 字符编码为数字实体，与 Excel 自身存储风格一致。"""
    return "".join(ch if ord(ch) < 128 else f"&#{ord(ch)};" for ch in str(s))


def dec(s):
    return ENT_RE.sub(lambda m: chr(int(m.group(1))), s)


def read_entries(path):
    with zipfile.ZipFile(path) as z:
        return [(i, z.read(i.filename)) for i in z.infolist()]


def write_entries(path, entries):
    """按原序写入，逐条目还原 ZipInfo 的 date_time / compress_type / external_attr。"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in entries:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.create_system = info.create_system
            z.writestr(zi, data)


def sheet_map(entries):
    """返回 {表名: 'xl/worksheets/sheetN.xml'}。"""
    files = {i.filename: d for i, d in entries}
    wb = files.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    rels = files.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "ignore")
    rid2target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid2target.update({k: v for v, k in
                       re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid2target.get(rid, "")
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[dec(name)] = t
    return out


def find_cell(xml, ref):
    """返回目标单元格 XML 片段的 (start, end)；支持自闭合与成对标签。"""
    pat = re.compile(r'<c r="%s"(?:\s[^>]*)?(?:/>|>.*?</c>)' % re.escape(ref), re.S)
    m = pat.search(xml)
    return (m.start(), m.end()) if m else None


def cmd_list(path):
    entries = read_entries(path)
    for name, target in sheet_map(entries).items():
        print(f"{name}\t{target}")
    charts = [i.filename for i, _ in entries if "chart" in i.filename.lower()]
    print(f"\n图表部件 {len(charts)} 个：{charts or '无'}")
    print(f"zip 条目 {len(entries)} 个")


def cmd_dump(path, sheet, ref):
    entries = read_entries(path)
    target = sheet_map(entries).get(sheet)
    if not target:
        sys.exit(f"找不到工作表：{sheet}")
    xml = dict((i.filename, d) for i, d in entries)[target].decode("utf-8", "ignore")
    span = find_cell(xml, ref)
    if not span:
        sys.exit(f"找不到单元格 {ref}")
    print(dec(xml[span[0]:span[1]]))


def cmd_patch(path, sheet, ref, old, new, apply_):
    entries = read_entries(path)
    target = sheet_map(entries).get(sheet)
    if not target:
        sys.exit(f"找不到工作表：{sheet}；可用：{list(sheet_map(entries))}")

    idx = next(k for k, (i, _) in enumerate(entries) if i.filename == target)
    raw = entries[idx][1].decode("utf-8")
    span = find_cell(raw, ref)
    if not span:
        sys.exit(f"找不到单元格 {ref}")
    block = dec(raw[span[0]:span[1]])

    hits = block.count(old)
    if hits != 1:
        sys.exit(f"旧文在该单元格命中 {hits} 次（要求恰好 1 次），已中止。\n"
                 f"单元格内容：{block[:400]}")
    global_hits = dec(raw).count(old)
    if global_hits != 1:
        print(f"⚠ 旧文在该工作表全文命中 {global_hits} 次；仅替换目标单元格，请复核是否漏改")

    if old == new:
        sys.exit("新旧文本相同，无需修改")

    new_block = block.replace(old, enc(new) if any(ord(c) > 127 for c in new) else new)
    new_raw = raw[:span[0]] + new_block + raw[span[1]:]

    print(f"[{'APPLY' if apply_ else 'DRY-RUN'}] {sheet}!{ref}")
    print(f"  旧：{old!r}")
    print(f"  新：{new!r}")
    print(f"  单元格：{block[:200]}")

    if not apply_:
        print("\n未落盘。确认无误后加 --apply。")
        return

    bak = f"{path}.bak_{datetime.now():%Y%m%d_%H%M%S}"
    shutil.copy2(path, bak)
    print(f"  备份：{bak}")

    entries[idx] = (entries[idx][0], new_raw.encode("utf-8"))
    write_entries(path, entries)

    # ---- 落盘校验
    after = read_entries(path)
    before = read_entries(bak)
    assert len(after) == len(before), "zip 条目数变化，已中止（请从备份还原）"
    charts_before = [i.filename for i, _ in before if "chart" in i.filename.lower()]
    charts_after = [i.filename for i, _ in after if "chart" in i.filename.lower()]
    assert charts_before == charts_after, "图表部件丢失，请从备份还原"
    missing = {i.filename for i, _ in before} - {i.filename for i, _ in after}
    assert not missing, f"条目丢失：{missing}"

    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        n = len(wb.sheetnames)
        wb.close()
        print(f"  ✓ 校验通过：{len(after)} 条目、图表 {len(charts_after)} 个、{n} 张表可读")
    except Exception as e:
        print(f"  ✗ openpyxl 回读失败：{e}\n  请从备份还原：{bak}")
        sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description="保留图表的单元格级 xlsx 修补")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list", help="列出工作表与图表部件")
    p.add_argument("book")

    p = sub.add_parser("dump", help="查看单元格原始 XML")
    p.add_argument("book"); p.add_argument("sheet"); p.add_argument("ref")

    p = sub.add_parser("patch", help="替换单元格内的文本（默认 dry-run）")
    p.add_argument("book"); p.add_argument("sheet"); p.add_argument("ref")
    p.add_argument("--old", required=True)
    p.add_argument("--new", required=True)
    p.add_argument("--apply", action="store_true")

    a = ap.parse_args()
    if a.cmd == "list":
        cmd_list(a.book)
    elif a.cmd == "dump":
        cmd_dump(a.book, a.sheet, a.ref)
    else:
        cmd_patch(a.book, a.sheet, a.ref, a.old, a.new, a.apply)


if __name__ == "__main__":
    main()
