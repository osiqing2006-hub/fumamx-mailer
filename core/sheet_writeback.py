# -*- coding: utf-8 -*-
"""发送记录回写客户表（2026-09-24 新增）

严格照**原客户表里已有的记录格式**来写：
    · 不新建行、不新建表 —— 在客户所在的那一行里追加
    · 找到该行最后一条记录所在的列，从**下一列**写，绝不覆盖已有单元格
    · 格式与既有记录一致：单个日期，YY.M.D（如 26.8.24、26.9.24）
    · 列不够就照同样格式向右扩展，并把新扩了哪些列报出来

原表只读不改，结果另存一份到 out_dir（默认 output/）。
"""
from __future__ import annotations

import datetime
import json
import re
import shutil
from pathlib import Path

DATE_RE = re.compile(r"^\d{2}\.\d{1,2}\.\d{1,2}$")     # 26.8.24 / 26.9.24


def fmt_date(d: datetime.date | datetime.datetime) -> str:
    """表格既有格式：YY.M.D（年两位，月和日无前导零）"""
    return f"{d.year % 100}.{d.month}.{d.day}"


def find_mail_col(header: list) -> int:
    """在表头里找邮箱列。各分表列名不统一：邮箱 / 邮箱\\/电话 / 邮箱\\n2026/09/18 9:00"""
    best, score = -1, -1
    for i, h in enumerate(header):
        t = str(h or "").strip()
        if not t:
            continue
        low = t.lower()
        s = 3 if "邮箱" in t else (2 if ("mail" in low or "e-mail" in low) else 0)
        if s > score:
            best, score = i, s
    return best


def next_record_col(row: list, header_n: int, mail_col: int) -> int:
    """该行下一条记录该写在哪一列（0-based）。

    优先接在最后一个日期格之后；该行还没有任何日期时，从表头之后开始
    （同时保证不覆盖该行已有的其它内容）。
    """
    last_date, last_any = -1, -1
    for i, v in enumerate(row):
        if v is None:
            continue
        t = str(v).strip()
        if not t:
            continue
        if i > last_any:
            last_any = i
        if DATE_RE.match(t) and i != mail_col and i > last_date:
            last_date = i
    if last_date >= 0:
        return last_date + 1
    return max(header_n, last_any + 1 if last_any >= header_n else header_n)


def _norm_emails(cell) -> list[str]:
    try:
        from .engine import normalize_emails
        return [m.lower() for m in normalize_emails(cell)]
    except Exception:
        s = str(cell or "")
        return [m.lower() for m in re.findall(r"[\w.\-+]+@[\w.\-]+\.\w+", s)]


def append_send_record(src: str | Path,
                       emails,
                       send_date: datetime.date | datetime.datetime | None = None,
                       out_dir: str | Path = "",
                       out_name: str = "") -> dict:
    """把这一批已发客户的发送日期追加写回客户表，另存一份到 out_dir。

    emails: 可迭代的邮箱（或 {邮箱: 日期} 字典，用于不同客户不同日期）
    返回 dict：ok / out / written / per_sheet / new_cols / unmatched / stamp
    """
    import openpyxl

    src = Path(src)
    if not str(out_dir or "").strip():
        out_dir = Path(__file__).resolve().parent.parent / "output"
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if isinstance(emails, dict):
        pairs = {str(k).strip().lower(): (v or send_date) for k, v in emails.items()}
    else:
        pairs = {str(e).strip().lower(): send_date for e in (emails or [])}
    pairs = {k: v for k, v in pairs.items() if k}
    if not pairs:
        return {"ok": False, "msg": "没有要写入的邮箱"}

    if send_date is None:
        send_date = datetime.date.today()

    wb = openpyxl.load_workbook(src)
    per_sheet: dict[str, int] = {}
    new_cols: dict[str, list[int]] = {}
    unmatched = 0
    written = 0

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            continue
        header = list(rows[0])
        header_n = 0
        for i, h in enumerate(header):
            if str(h or "").strip():
                header_n = i + 1
        mail_col = find_mail_col(header)
        if mail_col < 0:
            continue

        n = 0
        for r_i, row in enumerate(rows[1:], start=2):   # Excel 行号从 2 起
            row = list(row)
            if mail_col >= len(row):
                continue
            mails = _norm_emails(row[mail_col])
            if not mails:
                continue
            hits = [m for m in mails if m in pairs]
            if not hits:
                unmatched += 1
                continue
            d = pairs[hits[0]]
            stamp = fmt_date(d if d else send_date)
            # 同一天不重复记
            if any(str(v).strip() == stamp for v in row if v is not None):
                continue
            col = next_record_col(row, header_n, mail_col)
            # 保险：目标格若被别的内容占着（备注、电话…），继续往右让位，绝不覆盖
            while col < 200 and ws.cell(row=r_i, column=col + 1).value not in (None, ""):
                t = str(ws.cell(row=r_i, column=col + 1).value).strip()
                if DATE_RE.match(t):     # 已是日期格 → 说明重复，放弃
                    col = -1
                    break
                col += 1
            if col < 0:
                continue
            ws.cell(row=r_i, column=col + 1).value = stamp
            n += 1
            if col + 1 > ws.max_column:
                new_cols.setdefault(ws.title, [])
                if (col + 1) not in new_cols[ws.title]:
                    new_cols[ws.title].append(col + 1)
        if n:
            per_sheet[ws.title] = n
            written += n

    today = datetime.date.today()
    name = out_name or f"{src.stem}_已同步发送记录_{today:%Y%m%d}.xlsx"
    out = out_dir / name
    wb.save(out)

    return {
        "ok": True,
        "out": str(out),
        "written": written,
        "per_sheet": per_sheet,
        "new_cols": new_cols,
        "unmatched": unmatched,
        "stamp": fmt_date(send_date),
    }
