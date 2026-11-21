# -*- coding: utf-8 -*-
"""发送结果导出：把客户表的某个分表 + 发送状态 + 定时时间，导出成新的 xlsx。

产出目录默认是本程序目录下的 output/发送结果/（可用环境变量 FUMAMX_OUT_DIR 覆盖）
文件名：<分表名>_发送结果_<日期>.xlsx
在原表所有列后面追加三列：发送状态 / 定时发送（收件人当地时间）/ 完成日期
"""
import json
import os
import re
import datetime
import pathlib

OUT_DIR = os.environ.get("FUMAMX_OUT_DIR") or str(
    pathlib.Path(__file__).resolve().parent.parent / "output" / "发送结果")

# 客户表里简短的国家名 -> 孚盟下拉里能搜到的完整名称
COUNTRY_ALIAS = {
    "印尼": "印度尼西亚", "印度尼亚": "印度尼西亚", "indo": "印度尼西亚",
    "俄国": "俄罗斯", "俄": "俄罗斯", "俄罗斯联邦": "俄罗斯",
    "南韩": "韩国", "南朝鲜": "韩国",
    "阿联酋": "阿拉伯联合酋长国", "uae": "阿拉伯联合酋长国",
    "南非": "南非", "纽西兰": "新西兰", "澳洲": "澳大利亚",
    "美利坚": "美国", "usa": "美国", "uk": "英国", "大不列颠": "英国",
    "英格兰": "英国", "苏格兰": "英国", "尼德兰": "荷兰",
    "义大利": "意大利", "法兰西": "法国", "马来": "马来西亚",
    "捷克共和国": "捷克", "丹麦王国": "丹麦",
}


def fix_country(name: str) -> str:
    """把客户表里不完整的国家名补成孚盟能识别的写法。"""
    s = (name or "").strip()
    if not s:
        return s
    if s in COUNTRY_ALIAS:
        return COUNTRY_ALIAS[s]
    k = s.lower()
    if k in COUNTRY_ALIAS:
        return COUNTRY_ALIAS[k]
    return s


def load_sent(data_dir) -> dict:
    p = pathlib.Path(data_dir) / "sent_history.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def parse_log_times(log_paths) -> dict:
    """从运行日志里解析 邮箱 -> 实际定时时间（收件人当地）。"""
    out = {}
    pat = re.compile(r"✓[^|]*?(\S+@\S+?)[，,（(].*?当地\s*([\d/]{10}\s[\d:]{5})")
    pat2 = re.compile(r"(\S+@\S+?)[，,].*?已定时（.*?当地\s*([\d/]{10}\s[\d:]{5})")
    for lp in log_paths:
        try:
            txt = pathlib.Path(lp).read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        for line in txt.splitlines():
            for p in (pat2, pat):
                m = p.search(line)
                if m:
                    out[m.group(1).strip().lower()] = m.group(2)
                    break
    return out


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", (s or "Sheet").strip())[:40]


def export_sheet(xlsx_path: str, sheet_idx: int, sheet_name: str,
                 data_dir, when_default: str = "2026/09/22 09:00",
                 out_dir: str = OUT_DIR, log_paths=()) -> str:
    """导出单个分表的结果表，返回生成的文件路径。"""
    from core import xlsxlight
    from openpyxl import Workbook

    rows = xlsxlight.read_sheet(xlsx_path, sheet_idx)
    header, data = rows[0], rows[1:]
    hdr = [str(h).strip() for h in header]
    ci_mail = hdr.index("邮箱") if "邮箱" in hdr else 8
    ci_ctry = hdr.index("国家") if "国家" in hdr else 2

    sent = load_sent(data_dir)
    times = parse_log_times(log_paths)

    os.makedirs(out_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "发送结果"
    ws.append(list(header) + ["发送状态", "定时发送（当地时间）", "完成日期", "孚盟国家名"])

    n_done = 0
    for r in data:
        email = str(r[ci_mail] or "").strip()
        raw_country = str(r[ci_ctry] or "").strip()
        if not email or "@" not in email:
            continue
        rec = sent.get(email.lower()) or sent.get(email)
        if rec:
            status = "已定时" if "sched" in str(rec) else "已发送"
            done = str(rec).replace("(sched)", "").strip()
            n_done += 1
        else:
            status, done = "未发送", ""
        when = times.get(email.lower(), "") or (when_default if rec else "")
        ws.append(list(r) + [status, when, done, fix_country(raw_country)])

    day = datetime.date.today().strftime("%Y%m%d")
    out = os.path.join(out_dir, f"{safe_name(sheet_name)}_发送结果_{day}.xlsx")
    wb.save(out)
    return out, n_done


def export_today_list(data_dir, out_dir: str = OUT_DIR, log_paths=()) -> str:
    """导出今天（最近一批）成功定时/发送的客户清单。"""
    from openpyxl import Workbook

    sent = load_sent(data_dir)
    times = parse_log_times(log_paths)
    os.makedirs(out_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "已定时清单"
    ws.append(["邮箱", "状态", "完成日期", "定时发送（当地时间）"])
    n = 0
    for email, rec in sent.items():
        kind = "已定时" if "sched" in str(rec) else "已发送"
        ws.append([email, kind, str(rec).replace("(sched)", "").strip(),
                   times.get(str(email).lower(), "")])
        n += 1
    day = datetime.date.today().strftime("%Y%m%d")
    out = os.path.join(out_dir, f"今日已定时清单_{day}.xlsx")
    wb.save(out)
    return out, n
