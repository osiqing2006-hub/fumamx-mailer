# -*- coding: utf-8 -*-
"""极简 xlsx 读写（纯标准库，零依赖）。

孚盟群发助手只用到两件事：
  1. 读出工作表里的文本（国家 / 邮箱 / 客户名 / 公司）
  2. 把发送日期写回指定列（如 L 列），其余内容原样保留

xlsx 本质是 zip + XML，这里直接操作 XML，不动样式、公式和其他单元格。
"""
from __future__ import annotations

import csv
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def col_to_index(col: str) -> int:
    """A->0, B->1, L->11, AA->26"""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def index_to_col(idx: int) -> str:
    """0->A, 11->L"""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_names(path: str | Path) -> list[str]:
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        return ["Sheet1"]
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/workbook.xml"))
    return [s.get("name", "") for s in root.iter(f"{NS_MAIN}sheet")]


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    try:
        data = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    out = []
    for si in root.iter(f"{NS_MAIN}si"):
        out.append("".join(t.text or "" for t in si.iter(f"{NS_MAIN}t")))
    return out


def _sheet_path(z: zipfile.ZipFile, sheet: str | int | None) -> str:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_map = {r.get("Id"): r.get("Target") for r in rels.iter(f"{NS_PKGREL}Relationship")}
    sheets = list(wb.iter(f"{NS_MAIN}sheet"))
    if sheet is None:
        target = sheets[0]
    elif isinstance(sheet, int):
        target = sheets[sheet]
    else:
        target = next((s for s in sheets if s.get("name") == sheet), sheets[0])
    tgt = rid_map.get(target.get(f"{NS_REL}id"), "worksheets/sheet1.xml")
    if not tgt.startswith("/"):
        tgt = "xl/" + tgt.lstrip("./")
    else:
        tgt = tgt.lstrip("/")
    return tgt


def read_sheet(path: str | Path, sheet: str | int | None = None) -> list[list[str]]:
    """返回二维字符串数组（空单元格为空串）。"""
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            return [list(r) for r in csv.reader(f)]
    with zipfile.ZipFile(path) as z:
        shared = _shared_strings(z)
        sp = _sheet_path(z, sheet)
        root = ET.fromstring(z.read(sp))
    rows: list[list[str]] = []
    max_c = 0
    for row in root.iter(f"{NS_MAIN}row"):
        cells: list[str] = []
        for c in row.iter(f"{NS_MAIN}c"):
            ref = c.get("r") or ""
            m = _CELL_RE.match(ref)
            if m:
                ci = col_to_index(m.group(1))
                while len(cells) < ci:
                    cells.append("")
            t = c.get("t")
            if t == "inlineStr":
                v = "".join(x.text or "" for x in c.iter(f"{NS_MAIN}t"))
            else:
                ve = c.find(f"{NS_MAIN}v")
                v = ve.text or "" if ve is not None else ""
                if t == "s" and v.isdigit():
                    v = shared[int(v)] if int(v) < len(shared) else ""
                elif t == "b":
                    v = "TRUE" if v == "1" else "FALSE"
            cells.append(v if v is not None else "")
        max_c = max(max_c, len(cells))
        rows.append(cells)
    for r in rows:
        if len(r) < max_c:
            r.extend([""] * (max_c - len(r)))
    return rows


def write_cells(path: str | Path, sheet: str | int | None, cells: dict[str, str],
                backup: bool = True) -> str:
    """把 {单元格引用: 文本} 写回原文件（原地修改 zip 里的 sheet XML）。

    例：write_cells(f, "2604广交会", {"L2": "26.9.20", "L3": "26.9.20"})
    只改这些单元格，样式/公式/其他内容全部保留。
    """
    path = Path(path)
    if backup:
        bak = path.with_suffix(path.suffix + ".bak")
        if not bak.exists():
            shutil.copy2(path, bak)
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        sp = _sheet_path(z, sheet)
        xml = z.read(sp).decode("utf-8")
        payload = {n: z.read(n) for n in names}

    def repl_row(m: re.Match) -> str:
        row_no = m.group(2)
        inner = m.group(3)
        for ref, val in cells.items():
            mm = _CELL_RE.match(ref)
            if not mm or mm.group(2) != row_no:
                continue
            col = mm.group(1)
            esc = (val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
            new_c = (f'<c r="{col}{row_no}" t="inlineStr"><is><t xml:space="preserve">'
                     f"{esc}</t></is></c>")
            pat = re.compile(r'<c r="%s%s"(?:\s[^>]*)?(?:/>|>.*?</c>)' % (col, row_no), re.S)
            if pat.search(inner):
                inner = pat.sub(new_c, inner, count=1)
            else:
                inner += new_c
        return f"{m.group(1)}{inner}</row>"

    xml = re.sub(r"(<row\b[^>]*>)(.*?)(</row>)",
                 lambda m: m.group(1) + m.group(2) + m.group(3), xml, flags=re.S)

    def patch(m: re.Match) -> str:
        head, inner, tail = m.group(1), m.group(2), m.group(3)
        row_no = re.search(r'r="(\d+)"', head)
        if not row_no:
            return m.group(0)
        rn = row_no.group(1)
        for ref, val in cells.items():
            mm = _CELL_RE.match(ref)
            if not mm or mm.group(2) != rn:
                continue
            col = mm.group(1)
            esc = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            new_c = (f'<c r="{col}{rn}" t="inlineStr"><is><t xml:space="preserve">'
                     f"{esc}</t></is></c>")
            pat = re.compile(r'<c r="%s%s"(?:\s[^>]*)?(?:/>|>.*?</c>)' % (col, rn), re.S)
            inner = pat.sub(new_c, inner, count=1) if pat.search(inner) else inner + new_c
        return head + inner + tail

    xml = re.sub(r"(<row\b[^>]*>)(.*?)(</row>)", patch, xml, flags=re.S)

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, xml.encode("utf-8") if n == sp else payload[n])
    shutil.move(str(tmp), str(path))
    return str(path)
