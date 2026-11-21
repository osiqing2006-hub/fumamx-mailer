# -*- coding: utf-8 -*-
"""发送流水 + 历史发送记录查询（2026-09-26）

用户要的三件事：
  ① 每个「模板 × 主题 × 客户分表」发了多少次（能按模板/主题分组汇总）
  ② 能筛选（模板/主题/分表/日期区间）+ 关键字搜索 + 分页 + 按发送次数排序
  ③ 每行能「重发」——复用原模板+主题，向指定分表重发，重发前给出预估条数

数据层：data/send_flow.jsonl（只追加流水，一行一条）
    {"ts","date","email","sheet","tpl","subject","ok","note"}
不动 sent_history.json —— 原有去重逻辑完全不受影响。

【不重复计数】是用户点名要求的，两处最容易出错，这里都堵了：
  · 分桶键 =（模板, 主题, 分表）三元组，桶内用「邮箱集合」去重。
    ——「发送次数」=流水条数（同一人重发两次算 2）；
      「去重人数」=桶内不同邮箱数（同一人重发两次算 1）。
    两个都给，避免"同一封被重发过导致数字虚高"。
  · 跨分表合并 / 总计：同一个人可能同时在两个分表里，直接 sum 会 double count，
    所以合并与总计一律走 **集合并集**，并额外给出 raw_sum 方便人工核对。

【性能】流水是单行 JSON，量大时每次都全读会卡。这里按
    (文件路径, mtime_ns, 文件大小) 做键缓存一次解析结果；
    解析 → 聚合是 O(n) 单趟；筛选/排序/分页只作用在**聚合后的分组**（几百条）。
"""
from __future__ import annotations

import ast
import json
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# 旧数据没有模板/主题字段时的占位（不冒充真实模板名）
OLD_TPL = "（旧记录·未记模板）"
NO_SUBJ = "（旧记录·未记主题）"
NO_SHEET = "（未记分表）"

# 日志里的成功行：[定时] ✓ 第 3/132 位 a@b.com（西班牙）已定时（…
#                 [补发] ✓ 第 1/39 位 a@b.com（西班牙）补发成功｜…
RE_OK = re.compile(r"\[(?:定时|补发)\]\s*✓\s*第\s*\d+/\d+\s*位\s+([\w.\-+]+@[\w.\-]+\.\w+)")
RE_SHEET = re.compile(r"已恢复上次的客户表：.*?/\s*(\S+)\s*$")
RE_EXPORT = re.compile(r"已生成结果表：.*?([^\\/]+?)_发送结果_\d+\.xlsx")
# 旧日志没记模板名，但记了本轮主题 —— 用它把不同轮次区分开
RE_SUBJECT = re.compile(r"生成任务：.*?主题「(.+?)」")
# 「② 客户表」载入时会打这一行：载入 667 行，列：['名字','国家',…]
# ——> 行数 + 表头 = 分表指纹。比"最后一条恢复提示"准得多：
#     app 启动那句「已恢复上次的客户表」往往写在当天日志的**最后**，
#     用它归因会把早上发的记录算到后半夜才打开的分表上去。
RE_LOAD = re.compile(r"载入\s*(\d+)\s*行，列：(\[[^\]]*\])")
RE_XLSX = re.compile(r"已恢复上次的客户表：\s*([^/\\]+?)\s*/")


def sheet_fingerprints(xlsx_paths) -> dict:
    """给客户表的每个分表做「指纹」，用来认出日志里当时载入的是哪一栏。

    指纹 = (数据行数, 表头元组)。两个维度一起撞车的概率极低
    （用户这张表 12 个分表零冲突），比只看行数或只看表头都稳。
    """
    try:
        from core import xlsxlight
    except Exception:
        return {}
    out: dict = {}
    for p in xlsx_paths or []:
        try:
            names = xlsxlight.sheet_names(str(p))
        except Exception:
            continue
        for n in names:
            try:
                rows = xlsxlight.read_sheet(str(p), n)
            except Exception:
                continue
            if not rows or not rows[0]:
                continue
            cols = tuple(str(h) or f"第{i + 1}列" for i, h in enumerate(rows[0]))
            # setdefault：多份客户表里有同名同结构的，保留先出现的那份
            out.setdefault((len(rows) - 1, cols), str(n))
    return out


def candidate_xlsx(xlsx_path: str = "", logs_dir=None) -> list[str]:
    """可能用过哪些客户表：当前这张 + 日志里提到过的文件名（在常见目录里找）。"""
    paths: list[str] = []
    if xlsx_path and Path(xlsx_path).exists():
        paths.append(str(xlsx_path))
    if not logs_dir:
        return paths
    names: set = set()
    for fp in sorted(Path(logs_dir).glob("run_*.log")):
        try:
            txt = fp.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue
        for m in RE_XLSX.finditer(txt):
            names.add(m.group(1).strip())
    roots = []
    if xlsx_path:
        roots.append(Path(xlsx_path).parent)
    roots += [Path(__file__).resolve().parent.parent / "output",
              Path.home() / "Desktop", Path.home() / "Documents"]
    for n in names:
        for r in roots:
            p = r / n
            if p.exists():
                paths.append(str(p))
                break
        for follow in Path(xlsx_path).parent.rglob(n) if xlsx_path else []:
            paths.append(str(follow))
            break
    seen, out = set(), []
    for p in paths:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


# ---------------------------------------------------------------- 基础读写
def flow_path(data_dir) -> Path:
    return Path(data_dir) / "send_flow.jsonl"


def invalidate() -> None:
    """刚写完流水就把缓存作废，保证界面立刻读到新数据。"""
    _CACHE["key"] = None


def record(data_dir, email: str, tpl: str = "", sheet: str = "",
           subject: str = "", ok: bool = True, note: str = "",
           when: datetime | None = None) -> None:
    """追加一条发送流水。任何时候出错都不影响主流程（统计不能绊倒发送）。"""
    try:
        d = when or datetime.now()
        row = {"ts": d.strftime("%Y-%m-%d %H:%M:%S"), "date": d.strftime("%Y-%m-%d"),
               "email": (email or "").strip().lower(), "sheet": (sheet or "").strip(),
               "tpl": (tpl or "").strip(), "subject": (subject or "").strip()[:80],
               "ok": bool(ok), "note": note}
        p = flow_path(data_dir)
        p.parent.mkdir(parents=True, exist_ok=True)
        with p.open("a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
    except Exception:
        pass
    finally:
        invalidate()


_CACHE: dict = {"key": None, "data": []}


def rows(data_dir, use_cache: bool = True) -> list[dict]:
    """读全部流水（带 (mtime_ns, size) 缓存；文件变了自动重读）。"""
    p = flow_path(data_dir)
    try:
        st = p.stat()
        key = (str(p), st.st_mtime_ns, st.st_size)
    except OSError:
        return [] if not use_cache else []
    if use_cache and _CACHE["key"] == key:
        return _CACHE["data"]
    out: list[dict] = []
    try:
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except Exception:
                continue
    except Exception:
        pass
    _CACHE["key"], _CACHE["data"] = key, out
    return out


def load(data_dir) -> list[dict]:
    return rows(data_dir)


# ---------------------------------------------------------------- 聚合核心
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def _bucket(t: str, s: str, sh: str) -> tuple[str, str, str]:
    """分组三元组：模板 / 主题 / 分表。缺什么就补占位，保证键唯一。"""
    return (_norm(t) or "（未填模板）", _norm(s) or NO_SUBJ, _norm(sh) or NO_SHEET)


def aggregate(recs: list[dict]) -> dict[tuple, dict]:
    """核心聚合：单趟遍历，按（模板,主题,分表）分桶。

    —— 这里就是「不重复计数」的第一道：桶内 people 用 set，
       同一同一个人重发多次只在 people 里留一份，sends 里照实累加。
    —— 失败行只进 fail 计数，不进 sends/people/dates，
       免得一个反复失败的客户把「本已发了多少」抬上去。
    """
    agg: dict[tuple, dict] = {}
    for r in recs:
        k = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))
        g = agg.get(k)
        if g is None:
            g = agg[k] = {"tpl": k[0], "subject": k[1], "sheet": k[2],
                          "sends": 0, "fail": 0, "people": set(),
                          "first": "", "last": ""}
        if not r.get("ok", True):
            g["fail"] += 1
            continue
        g["sends"] += 1
        e = (r.get("email") or "").strip().lower()
        if e:
            g["people"].add(e)
        d = (r.get("date") or (r.get("ts") or "")[:10]).strip()
        if d:
            if not g["first"] or d < g["first"]:
                g["first"] = d
            if not g["last"] or d > g["last"]:
                g["last"] = d
    return agg


def _merge_sheets(agg: dict[tuple, dict]) -> dict[tuple, dict]:
    """跨分表合并：把同一（模板,主题）下所有分表合成一行。

    ★ 合并必须用**集合并集**求人数：同一个邮箱可能同时存在于多个分表，
      直接把各分表的 sends 相加会把一个人算好几遍（这是最容易虚高的地方）。
    """
    merged: dict[tuple, dict] = {}
    for (t, s, sh), g in agg.items():
        k = (t, s)
        m = merged.get(k)
        if m is None:
            m = merged[k] = {"tpl": t, "subject": s, "sheet": "（全部分表）",
                             "sends": 0, "fail": 0, "people": set(),
                             "first": "", "last": "", "sheets": set()}
        m["sends"] += g["sends"]
        m["fail"] += g["fail"]
        m["people"] |= g["people"]          # 并集去重 ← 关键
        m["sheets"].add(sh)
        for fld in ("first", "last"):
            if g[fld]:
                if not m[fld] or (g[fld] < m[fld] if fld == "first" else g[fld] > m[fld]):
                    m[fld] = g[fld]
    for m in merged.values():
        # 合并后 shelf 数 >1 时把"覆盖了哪几个分表"写进 sheet 列，方便一眼看懂
        ss = sorted(m.pop("sheets") or [])
        if ss and len(ss) != 1:
            m["sheet"] = "（合并 %d 个分表）%s" % (len(ss), "、".join(ss[:3]) + ("…" if len(ss) > 3 else ""))
        elif ss:
            m["sheet"] = ss[0]
    return merged


# ---------------------------------------------------------------- 查询
def query(data_dir, *, tpl: str = "", subject: str = "", sheet: str = "",
          date_from: str = "", date_to: str = "", kw: str = "",
          only_ok: bool = True, merge: bool = False,
          sort: str = "count", desc: bool = True,
          page: int = 1, page_size: int = 50) -> dict:
    """历史发送记录查询（ filters → 聚合 → 排序 → 分页 → 总计 ）。

    返回 {"rows":[...], "total","page","pages","grand":{...}}
    rows 每行 = {key,tpl,subject,sheet,sends,people,fail,first,last,emails}
    """
    allrows = rows(data_dir)

    # ① 行级过滤（日期区间 / 只看成功）—— 先砍掉一大截再聚合
    recs = []
    for r in allrows:
        if only_ok and not r.get("ok", True):
            continue
        d = (r.get("date") or (r.get("ts") or "")[:10]).strip()
        if date_from and d and d < date_from:
            continue
        if date_to and d and d > date_to:
            continue
        recs.append(r)

    agg = aggregate(recs)
    if merge:
        agg = _merge_sheets(agg)

    # ② 分组级过滤（模板/主题/分表/关键字）
    ft, fs, fsh = _norm(tpl), _norm(subject).lower(), _norm(sheet)
    kwl = _norm(kw).lower()
    groups = []
    for key, g in agg.items():
        if ft and ft != g["tpl"] and ft not in g["tpl"]:
            continue
        if fs and fs not in (g["subject"] or "").lower():
            continue
        if fsh and fsh != g["sheet"]:
            continue
        if kwl:
            # 关键字同时覆盖：模板 / 主题 / 分表 / 该组收件人邮箱 / 备注轮次
            hay = " ".join((g["tpl"], g["subject"], g["sheet"])).lower()
            hit = kwl in hay or any(kwl in e for e in g["people"])
            if not hit:
                continue
        groups.append(g)

    # ③ 排序（可选的几个维度）
    keyfuncs = {
        "count": lambda g: g["sends"],
        "people": lambda g: len(g["people"]),
        "last": lambda g: g["last"] or "",
        "first": lambda g: g["first"] or "",
        "tpl": lambda g: g["tpl"],
        "sheet": lambda g: g["sheet"],
    }
    groups.sort(key=keyfuncs.get(sort, keyfuncs["count"]), reverse=desc)

    # ④ 分页
    total = len(groups)
    size = max(1, int(page_size or 50))
    pages = max(1, (total + size - 1) // size)
    page = min(max(1, int(page or 1)), pages)
    sl = groups[(page - 1) * size: page * size]

    out = []
    for g in sl:
        out.append({"tpl": g["tpl"], "subject": g["subject"], "sheet": g["sheet"],
                    "sends": g["sends"], "people": len(g["people"]), "fail": g["fail"],
                    "first": g["first"], "last": g["last"],
                    "emails": sorted(g["people"])})

    # ⑤ 总计：**对全部结果**（不是当前这一页）统计。
    #    人数跨组也用集合并集 —— 同一个人出现在两个主题/分表也只能算一个人。
    union_people: set = set()
    for g in groups:
        union_people |= g["people"]
    grand = {
        "groups": total,
        "sends": sum(g["sends"] for g in groups),
        "fail": sum(g["fail"] for g in groups),
        "sends_raw_sum": sum(g["sends"] for g in groups),   # 留档：未去重的条数合计
        "people": len(union_people),                        # 去重后的真实人数
        "sheets": len({g["sheet"] for g in groups}),
        "tpls": len({g["tpl"] for g in groups}),
        "scanned": len(allrows),                            # 参与扫描的流水总行数
    }
    return {"rows": out, "total": total, "page": page, "pages": pages,
            "grand": grand, "sort": sort, "desc": desc}


def template_stats(data_dir) -> list[dict]:
    """每个模板的「首次使用 / 最近使用 / 发送条数 / 去重人数」。

    ⑥ 页模板下拉给历史模板排序用：按最近使用时间倒序
    （没有最近时间的退化到首次使用时间，再退化到模板名），
    这样用户最近在用的模板总是排在前面。
    """
    agg: dict[str, dict] = {}
    for r in rows(data_dir):
        t = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))[0]
        g = agg.get(t)
        if g is None:
            g = agg[t] = {"tpl": t, "first": "", "last": "", "sends": 0,
                          "people": set()}
        if not r.get("ok", True):
            continue
        g["sends"] += 1
        e = (r.get("email") or "").strip().lower()
        if e:
            g["people"].add(e)
        d = (r.get("date") or (r.get("ts") or "")[:10]).strip()
        if d:
            if not g["first"] or d < g["first"]:
                g["first"] = d
            if not g["last"] or d > g["last"]:
                g["last"] = d
    out = []
    for g in agg.values():
        out.append({"tpl": g["tpl"], "first": g["first"], "last": g["last"],
                    "sends": g["sends"], "people": len(g["people"])})
    out.sort(key=lambda x: (x["last"] or x["first"] or "", x["tpl"]), reverse=True)
    return out


def meta(data_dir) -> dict:
    """给筛选下拉框用的枚举：模板 / 主题 / 分表（按出现频次倒序）。"""
    cnt_t, cnt_s, cnt_sh = defaultdict(int), defaultdict(int), defaultdict(int)
    for r in rows(data_dir):
        k = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))
        cnt_t[k[0]] += 1
        cnt_s[k[1]] += 1
        cnt_sh[k[2]] += 1
    top = lambda d: [k for k, _ in sorted(d.items(), key=lambda kv: -kv[1])]
    return {"templates": top(cnt_t), "subjects": top(cnt_s),
            "sheets": top(cnt_sh), "n": len(rows(data_dir))}


def _subj80(s) -> str:
    """主题统一按「前 80 字」比较。

    record() 落盘时会把主题截到 80 字；如果外面拿**完整的**主题字符串来比对，
    「完整版 in 截断版」永远不成立 → 明明发过却判成没发过 → 重复发送。
    所以凡是拿主题做比较的地方，两边都先过这个函数。
    """
    return _norm(s).lower()[:80]


def same_subject(a: str, b: str) -> bool:
    """两个主题算不算同一轮（双向包含，容忍截断/首尾空格差异）。"""
    x, y = _subj80(a), _subj80(b)
    return bool(x) and bool(y) and (x == y or x in y or y in x)


def sent_emails(data_dir, tpl: str = "", subject: str = "",
                sheet: str = "") -> set:
    """某「模板+主题（+分表）」已经发过的邮箱集合 —— 重发预估要用。

    ★ 这里**不带分表条件时是跨分表合并**的：同一个模板+主题，无论在哪个
      分表发过，都不再重发 —— 与查询端「跨分表不重复计数」口径一致。
    """
    out = set()
    ft, fs, fsh = _norm(tpl), _subj80(subject), _norm(sheet)
    for r in rows(data_dir):
        if not r.get("ok", True):
            continue
        k = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))
        if ft and ft != k[0] and ft not in k[0]:
            continue
        # 主题用 same_subject 双向比（流水里存的是前 80 字，直接子集匹配会漏）
        if fs and not same_subject(fs, k[1]):
            continue
        if fsh and fsh != k[2]:
            continue
        e = (r.get("email") or "").strip().lower()
        if e:
            out.add(e)
    return out


def sent_subjects_of(data_dir, email: str) -> set:
    """这个邮箱历史上发过哪些主题 —— 新一轮」模式下判断要不要跳过用它。"""
    e = (email or "").strip().lower()
    out = set()
    if not e:
        return out
    for r in rows(data_dir):
        if not r.get("ok", True):
            continue
        if (r.get("email") or "").strip().lower() != e:
            continue
        out.add(_bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))[1])
    return out


def detail(data_dir, tpl: str, subject: str, sheet: str) -> list[dict]:
    """某一行的明细：谁、什么时候发的、成功还是失败（按时间倒序）。"""
    wt, ws, wsh = _bucket(tpl, subject, sheet)
    out = []
    for r in rows(data_dir):
        k = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))
        # 模板/分表要完全相等，主题用 same_subject 容忍截断（外面可能拿完整主题来查）
        if k[0] != wt or k[2] != wsh or not same_subject(ws, k[1]):
            continue
        out.append({"email": r.get("email", ""), "ts": r.get("ts", ""),
                    "ok": bool(r.get("ok", True)), "note": r.get("note", "")})
    out.sort(key=lambda x: x["ts"], reverse=True)
    return out


# ---------------------------------------------------------------- 兼容老调用
def matrix(data_dir, only_ok: bool = True) -> dict:
    """（老接口，保留）{tpl: {sheet: 次数}} + 合计。"""
    tpl_sheet: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    emails: dict[str, set] = defaultdict(set)
    dates: dict[str, set] = defaultdict(set)
    total = 0
    for r in rows(data_dir):
        if only_ok and not r.get("ok", True):
            continue
        k = _bucket(r.get("tpl"), r.get("subject"), r.get("sheet"))
        tpl_sheet[k[0]][k[2]] += 1
        total += 1
        if r.get("email"):
            emails[(k[0], k[2])].add(r["email"])
        if r.get("date"):
            dates[(k[0], k[2])].add(r["date"])
    sheets = sorted({s for d in tpl_sheet.values() for s in d})
    return {
        "cells": {t: dict(d) for t, d in tpl_sheet.items()},
        "sheets": sheets,
        "emails": {f"{k[0]}||{k[1]}": len(v) for k, v in emails.items()},
        "dates": {f"{k[0]}||{k[1]}": sorted(v) for k, v in dates.items()},
        "total": total,
    }


# ------------------------------------------------- 旧日志回溯（可重复执行）
def backfill(data_dir, logs_dir, force: bool = False,
             xlsx_path: str = "") -> dict:
    """从 logs/run_*.log 把历史轮次的成功记录补进流水。

    · 去重键是（邮箱, 主题）—— 同一客户不同轮次是两条独立记录，不会被吃掉；
      同一个 customer 同一主题反复执行也只会补一次（随时可以再跑，不会翻倍）。
    · 旧日志确实没有模板名，模板列一律写 OLD_TPL（注明来源，不冒充真实模板名），
      主题列用日志里「生成任务：…主题「xxx」」这一句 —— 这能把不同轮次分开。
    """
    data_dir, logs_dir = Path(data_dir), Path(logs_dir)
    known = {(r.get("email"), _bucket(r.get("tpl"), r.get("subject"),
                                      r.get("sheet"))[1]) for r in rows(data_dir)}
    added, dup, files = 0, 0, []
    fps = sheet_fingerprints(candidate_xlsx(xlsx_path, logs_dir))   # 分表指纹表
    carry = NO_SHEET          # 跨文件沿用：本文件开头还没有任何线索时用上一份的
    for fp in sorted(logs_dir.glob("run_*.log")):
        try:
            txt = fp.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue
        files.append(fp.name)
        date = fp.stem.replace("run_", "")
        if len(date) == 8:
            date = f"{date[:4]}-{date[4:6]}-{date[6:]}"

        lines = txt.splitlines()
        # ---- 第一遍：把「主题切换点」和「分表线索点」按行号记下来 ----
        # 为什么不能一把梭顺序扫：app 启动时的「已恢复上次的客户表」往往写在
        # 当天日志的**最后**（那是下一次启动追加的），若按顺序一把梭，
        # 当天早上发的所有记录都会被算到后半夜才打开的那个分表上去。
        subj_marks: list[tuple[int, str]] = []      # (行号, 主题)
        sheet_marks: list[tuple[int, str]] = []     # (行号, 分表)
        oks: list[tuple[int, str]] = []             # (行号, 邮箱)
        for i, line in enumerate(lines):
            m = RE_SUBJECT.search(line)
            if m:
                subj_marks.append((i, m.group(1).strip()[:80] or NO_SUBJ))
                continue
            m = RE_SHEET.search(line)
            if m and m.group(1).strip():
                sheet_marks.append((i, m.group(1).strip()))
                continue
            m = RE_EXPORT.search(line)
            if m and m.group(1).strip():
                sheet_marks.append((i, m.group(1).strip()))
                continue
            m = RE_LOAD.search(line)
            if m and fps:
                # 「载入 N 行，列：[…]」→ 用指纹反查当时载入的是哪个分表。
                # 这是最强的一条线索：它就在真正读表的那一刻打出来的。
                try:
                    cols = tuple(str(c) for c in ast.literal_eval(m.group(2)))
                    who = fps.get((int(m.group(1)), cols))
                except Exception:
                    who = None
                if who:
                    sheet_marks.append((i, who))
                continue
            m = RE_OK.search(line)
            if m:
                oks.append((i, m.group(1).strip().lower()))
        sheet_marks.sort(key=lambda x: x[0])      # 两种线索按时间合并排序

        def _sheet_at(i: int) -> tuple[str, str]:
            """第 i 行属于哪个分表：取它**之前**最近的一条线索；没有就用 carry。
            返回 (sheet, 来源) —— 来源是「推断」还是「沿用上一份日志」，写进备注。"""
            cur, src = carry, "沿用"
            for mi, sh in sheet_marks:
                if mi < i:
                    cur, src = sh, "推断"
                else:
                    break
            return cur, src

        # ---- 第二遍：按成功行回填 ----
        for i, email in oks:
            if not email:
                continue
            sheet, src = _sheet_at(i)
            subj = subj_marks[-1][1] if subj_marks and subj_marks[0][0] < i else NO_SUBJ
            for mi, sj in subj_marks:                # 取这行之前最近的一次主题切换
                if mi < i:
                    subj = sj
                else:
                    break
            key = (email, subj)
            if key in known:
                dup += 1
                continue
            known.add(key)
            try:
                when = datetime.strptime(date, "%Y-%m-%d")
            except Exception:
                when = None
            record(data_dir, email, tpl=OLD_TPL, sheet=sheet, subject=subj,
                   ok=True, note=f"旧日志补录（分表为{src}）", when=when)
            added += 1
        if sheet_marks:
            carry = sheet_marks[-1][1]      # 把本文件最后一条线索带到下一个文件
    return {"added": added, "skipped_dup": dup, "files": files}


# ============ 2026-10-25 新增：分表跟进建议 / 按日期明细 ============

def sheet_advice(data_dir, sheet: str, library: dict | None = None) -> dict:
    """某个分表「上次用了什么模板 + 下一步建议用什么」。

    数据来源：data/send_flow.jsonl（发送流水）+ 可选的本地模板库 library。

    library 结构（data/tpl_library.json）：
      {"templates":[{"name","stage","stage_label","for_sheets","universal","retired",
                     "subject_options","body","note"}],
       "retired_templates":[{"name","reason"}]}

    推荐优先级（谁先命中用谁）：
      ① 库内匹配：for_sheets 命中当前分表、且这个分表用得最少的模板
      ② 历史转移：该分表「上次模板 → 下一个模板」最高频的那个（**排除停用模板**）
      ③ 库内通用：标记 universal 且这个分表没用过的
      ④ 空

    ⚠️ retired_templates 里的模板（目前是广交会邀请系列）**永远不会被推荐**。
    """
    sh = _norm(sheet)
    out = {"sheet": sheet, "last_tpl": "", "last_at": "", "sends": 0,
           "used": [], "next_tpl": "", "reason": "", "chain": [],
           "next_info": {}, "retired_hit": "",
           "next_subject": "", "subject_used_before": 0, "subject_options": []}
    if not sh:
        out["reason"] = "未选择分表"
        return _fin(out)

    retired = {_norm(r.get("name")) for r in (library or {}).get("retired_templates", [])}
    lib = [t for t in (library or {}).get("templates", [])
           if t.get("name") and not t.get("retired")]

    recs = [r for r in rows(data_dir) if _norm(r.get("sheet")) == sh]
    recs.sort(key=lambda r: (r.get("ts") or ""))
    out["sends"] = len(recs)

    # 别名 → 主名（流水里的老模板名要和库里的新名字对上，否则会被当成"没用过"而重复推荐）
    alias_map: dict[str, str] = {}
    for _t in lib:
        main = _norm(_t["name"])
        alias_map[main] = main
        for _a in (_t.get("aliases") or []):
            alias_map[_norm(_a)] = main

    def canon(x: str) -> str:
        k = _norm(x)
        return alias_map.get(k, k)

    cnt: dict[str, int] = {}
    for r in recs:
        t = canon(r.get("tpl"))
        if t:
            cnt[t] = cnt.get(t, 0) + 1
    out["used"] = sorted(cnt.items(), key=lambda kv: -kv[1])

    if recs:
        last = recs[-1]
        out["last_tpl"] = _norm(last.get("tpl"))
        out["last_at"] = (last.get("ts") or "")[:16]
        chain: list[str] = []
        for r in recs:
            t = _norm(r.get("tpl"))
            if t and (not chain or chain[-1] != t):
                chain.append(t)
        out["chain"] = chain[-8:]

    # 该分表上次用的模板在库里的「阶段」。用来防止推荐回退
    # （例如「一直未回复」分表已经走完第二轮，就不该再被推荐首封冷客户信）
    last_stage = 0
    if out["last_tpl"]:
        _lt = canon(out["last_tpl"])
        for _t in lib:
            if _norm(_t["name"]) == _lt:
                try:
                    last_stage = int(_t.get("stage") or 0)
                except Exception:
                    last_stage = 0
                break

    def _stage_ok(t: dict) -> bool:
        try:
            return int(t.get("stage") or 0) >= last_stage
        except Exception:
            return True

    def pick(t: dict) -> dict:
        return {"name": t["name"], "stage": t.get("stage"),
                "stage_label": t.get("stage_label", ""),
                "subject_options": t.get("subject_options", []),
                "body": t.get("body", ""), "note": t.get("note", "")}

    def _subj_key(x) -> str:
        """主题归一化：去编号前缀 "1. "、去 Re:/RE:、去空白、小写。"""
        v = str(x or "").strip()
        v = re.sub(r"^\d+\s*[\.、\)]\s*", "", v)
        v = re.sub(r"^(re|fw|fwd|回复|转发)\s*[:：]\s*", "", v, flags=re.I)
        return re.sub(r"\s+", " ", v).strip().lower()

    def _fin(o: dict) -> dict:
        """给结果补上「推荐主题」：在该模板的主题候选里，选这个分表用得最少的那个。
        全部都用过时就取最早那个（等于轮换回开头）。"""
        opts = (o.get("next_info") or {}).get("subject_options") or []
        if opts:
            used: dict[str, int] = {}
            for r in recs:
                k = _subj_key(r.get("subject"))
                if k:
                    used[k] = used.get(k, 0) + 1
            pick_i, pick_n = 0, None
            for i, opt in enumerate(opts):
                n = used.get(_subj_key(opt), 0)
                if pick_n is None or n < pick_n:
                    pick_i, pick_n = i, n
            o["next_subject"] = opts[pick_i]
            o["subject_used_before"] = pick_n or 0
            o["subject_options"] = opts
        return o

    # ---- ① 库内匹配：for_sheets 命中、且这个分表用得最少 ----
    if lib:
        hit = [t for t in lib
               if _stage_ok(t)
               and any(_norm(x) == sh or _norm(x) in sh or sh in _norm(x)
                       for x in (t.get("for_sheets") or []))]
        if hit:
            # 排序：先看 for_sheets 里有没有「精确等于」当前分表名的
            #（精确命中的比"包含式命中"更贴，例如「持续发信但从未被查看」
            #  应该先配「持续发信 从未被查看 v1」，而不是泛泛的首封信）
            def _rank(t):
                names = [_norm(x) for x in (t.get("for_sheets") or [])]
                exact = 0 if sh in names else 1
                return (exact, cnt.get(_norm(t["name"]), 0))
            hit.sort(key=_rank)
            best = hit[0]
            if cnt.get(_norm(best["name"]), 0) == 0:
                out["next_tpl"] = best["name"]
                out["next_info"] = pick(best)
                out["reason"] = (f"这个分表还没用过「{best['name']}」"
                                 f"（{best.get('stage_label','')}），按模板库的位置它就是下一步")
                return _fin(out)

    # ---- ② 历史转移（排除停用）----
    chain = out["chain"]
    trans: dict[tuple[str, str], int] = {}
    for i in range(len(chain) - 1):
        a, b = chain[i], chain[i + 1]
        trans[(a, b)] = trans.get((a, b), 0) + 1
    if out["last_tpl"]:
        cand = [(k[1], v) for k, v in trans.items()
                if k[0] == out["last_tpl"] and _norm(k[1]) not in retired]
        if cand:
            nxt = max(cand, key=lambda x: x[1])[0]
            out["next_tpl"] = nxt
            out["next_info"] = next((pick(t) for t in lib if _norm(t["name"]) == _norm(nxt)), {})
            out["reason"] = (f"上次用的是「{out['last_tpl']}」（{out['last_at']}）；"
                             f"按这个分表的历史顺序，下一步通常是「{nxt}」")
            return _fin(out)
        # 历史里唯一的后继是停用模板 → 明确告诉用户，而不是硬推
        blocked = [(k[1], v) for k, v in trans.items()
                   if k[0] == out["last_tpl"] and _norm(k[1]) in retired]
        if blocked:
            nm = max(blocked, key=lambda x: x[1])[0]
            out["retired_hit"] = nm
            out["reason"] = (f"上次用的是「{out['last_tpl']}」"
                             f"（{out['last_at'] or '时间未知'}）；"
                             f"历史上它后面接的是「{nm}」，但这个模板已停用"
                             f"（广交会邀请已发过），所以不推荐它。请从模板库另选一个。")
        if out["last_tpl"] and not out["reason"]:
            out["reason"] = (f"上次用的是「{out['last_tpl']}」"
                             f"（{out['last_at']}），历史上没有它之后接哪个模板的记录")

    # ---- ③ 库内通用：没用过的 universal ----
    if lib:
        uni = [t for t in lib if t.get("universal") and _stage_ok(t)
               and cnt.get(_norm(t["name"]), 0) == 0]
        if uni:
            best = uni[0]
            out["next_tpl"] = best["name"]
            out["next_info"] = pick(best)
            out["reason"] = (out["reason"] + "；" if out["reason"] else "") + \
                            f"库里还有通用模板「{best['name']}」没用过，可作下一封"
            return _fin(out)

    # ---- ④ 同一阶段隔满 14 天：允许换主题再跟一次（两周跟一次是外贸常规节奏）----
    # 场景：已发过 PI 的那 21 家，已经没有更靠后的模板了，但他们是最该再问一次的。
    # 规则上不再硬性判"走完"，而是"距上次满三周 → 换个主题再发同一模板"。
    if not out["next_tpl"] and out["last_tpl"] and out["last_at"]:
        try:
            import datetime as _dtm
            _last = _dtm.datetime.strptime(out["last_at"][:10], "%Y-%m-%d")
            _days = (_dtm.datetime.now() - _last).days
        except Exception:
            _days = 0
        _t = next((x for x in lib if _norm(x["name"]) == canon(out["last_tpl"])), None)
        if _t and _days >= 14:
            out["next_tpl"] = _t["name"]
            out["next_info"] = pick(_t)
            out["reason"] = (f"上次用「{_t['name']}」是 {out['last_at'][:10]}，"
                             f"已经过去 {_days} 天。同一模板、换个主题再跟一次，"
                             f"主题会和上次不同。")

    if last_stage > 0 and not out["next_tpl"]:
        out["reason"] = (f"这个分表已经用过第 {last_stage} 阶段的模板"
                         f"（{out['last_tpl']}），库里没有更靠后的可用模板 —— "
                         f"本轮算走完了。要继续只能换角度重发，建议隔 2–4 周。")
        return _fin(out)
    if not out["reason"]:
        out["reason"] = "这个分表还没有发送记录" if not recs else "推不出下一步，请手动选模板"
    return _fin(out)


def by_date(data_dir, *, date_from: str = "", date_to: str = "",
            tpl: str = "", sheet: str = "", only_ok: bool = True,
            sort: str = "date", desc: bool = True,
            page: int = 1, page_size: int = 50) -> dict:
    """按「日期」聚合的历史视图 —— 一天一行，看当天发了什么。

    与 query() 的区别：query 按「模板+主题+分表」聚合，这里按「日期」聚合。
    返回 {"rows":[{date,sends,people,fail,tpls,sheets,first,last}...],
          "total","page","pages","grand"}
    """
    allrows = rows(data_dir)
    ft, fsh = _norm(tpl), _norm(sheet)
    days: dict[str, dict] = {}
    for r in allrows:
        if only_ok and not r.get("ok", True):
            continue
        d = (r.get("date") or (r.get("ts") or "")[:10]).strip()
        if not d:
            continue
        if date_from and d < date_from:
            continue
        if date_to and d > date_to:
            continue
        if ft and ft not in _norm(r.get("tpl")):
            continue
        if fsh and fsh != _norm(r.get("sheet")):
            continue
        g = days.setdefault(d, {"date": d, "sends": 0, "fail": 0,
                                "people": set(), "tpls": {}, "sheets": {},
                                "first": "", "last": ""})
        ok = bool(r.get("ok", True))
        g["sends"] += 1
        if not ok:
            g["fail"] += 1
        e = (r.get("email") or "").strip().lower()
        if e:
            g["people"].add(e)
        t = _norm(r.get("tpl")) or "（未记模板）"
        g["tpls"][t] = g["tpls"].get(t, 0) + 1
        s = _norm(r.get("sheet")) or "（未记分表）"
        g["sheets"][s] = g["sheets"].get(s, 0) + 1
        ts = (r.get("ts") or "")[:16]
        if ts:
            if not g["first"] or ts < g["first"]:
                g["first"] = ts
            if not g["last"] or ts > g["last"]:
                g["last"] = ts

    groups = list(days.values())
    groups.sort(key=lambda g: g["date"], reverse=desc)

    total = len(groups)
    size = max(1, int(page_size or 50))
    pages = max(1, (total + size - 1) // size)
    page = min(max(1, int(page or 1)), pages)
    sl = groups[(page - 1) * size: page * size]

    out = []
    for g in sl:
        top = sorted(g["tpls"].items(), key=lambda kv: -kv[1])
        out.append({
            "date": g["date"], "sends": g["sends"], "people": len(g["people"]),
            "fail": g["fail"],
            "tpls": " / ".join(f"{t}({c})" for t, c in top[:3]),
            "tpl_count": len(g["tpls"]),
            "sheets": " / ".join(sorted(g["sheets"])[:3]),
            "sheet_count": len(g["sheets"]),
            "first": g["first"], "last": g["last"],
        })

    union: set = set()
    for g in groups:
        union |= g["people"]
    grand = {"days": total,
             "sends": sum(g["sends"] for g in groups),
             "fail": sum(g["fail"] for g in groups),
             "people": len(union),
             "scanned": len(allrows)}
    return {"rows": out, "total": total, "page": page, "pages": pages, "grand": grand}


def day_records(data_dir, date: str, *, tpl: str = "", sheet: str = "",
                only_ok: bool = False) -> list[dict]:
    """某一天的全部发送明细（给「双击某天看客户」用）。

    返回 [{ts,email,tpl,subject,sheet,ok,note}...]，按时间正序。
    """
    d = (date or "").strip()
    ft, fsh = _norm(tpl), _norm(sheet)
    out = []
    for r in rows(data_dir):
        rd = (r.get("date") or (r.get("ts") or "")[:10]).strip()
        if rd != d:
            continue
        if only_ok and not r.get("ok", True):
            continue
        if ft and ft not in _norm(r.get("tpl")):
            continue
        if fsh and fsh != _norm(r.get("sheet")):
            continue
        out.append({"ts": r.get("ts", ""), "email": r.get("email", ""),
                    "tpl": _norm(r.get("tpl")), "subject": r.get("subject", ""),
                    "sheet": r.get("sheet", ""), "ok": bool(r.get("ok", True)),
                    "note": r.get("note", "")})
    out.sort(key=lambda x: x["ts"])
    return out
