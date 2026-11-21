# -*- coding: utf-8 -*-
"""孚盟MX 群发助手 —— 双击「启动.bat」即可使用，不需要懂命令行。

界面四步：①连接孚盟 → ②选客户表 → ③设置 → ④开始发送
"""
from __future__ import annotations

import csv
import json
import os
import queue
import random
import re
import shutil
import sys
import threading

try:                                   # 结果导出（分表跑完自动出表）
    from core.result_export import (export_sheet, export_today_list,
                                    OUT_DIR as RESULT_DIR, fix_country)
except Exception:                      # 缺失时给出空实现，不影响主程序
    RESULT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "output", "发送结果")

    def export_sheet(*a, **k):
        return "", 0

    def export_today_list(*a, **k):
        return "", 0

    def fix_country(s):
        return s

# 2026-09-24：发送记录回写客户表（照原表既有格式向右追加，不覆盖、不新增行）
try:
    from core.sheet_writeback import append_send_record
except Exception:                      # 缺失时给空实现，不影响主程序
    def append_send_record(*a, **k):
        return {"ok": False, "msg": "缺少 core/sheet_writeback.py"}

# 2026-09-26：历史发送记录（模板×主题×分表 的发送次数查询 + 重发）
try:
    from core import send_stats
except Exception:
    send_stats = None

import time
import tkinter as tk
from datetime import timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))

from core import xlsxlight  # noqa: E402
from core.engine import CN, Engine, EngineConfig, normalize_emails, today_cn  # noqa: E402
from core.scheduler import Scheduler, TaskStore, now_cn, parse_when  # noqa: E402


# --------------------------------------------------------------------------
# 轻提示不许抢焦点（2026-10-06）
# --------------------------------------------------------------------------
def _fg_hwnd() -> int:
    """当前前台窗口句柄（非 Windows 或出错返回 0）。"""
    if os.name != "nt":
        return 0
    try:
        import ctypes
        return int(ctypes.windll.user32.GetForegroundWindow() or 0)
    except Exception:
        return 0


def _fg_back(hwnd: int) -> None:
    """把键盘焦点还给 hwnd。

    轻提示（_toast）是个置顶小窗，Windows 在它弹出时会把焦点抢过来 ——
    用户正在别的程序里打字时就被打断了。弹完立刻把焦点还回去。
    """
    if os.name != "nt" or not hwnd:
        return
    try:
        import ctypes
        u = ctypes.windll.user32
        if int(u.GetForegroundWindow() or 0) != hwnd:
            u.SetForegroundWindow(hwnd)
    except Exception:
        pass

# 2026-09-23：失败原因分类器（网络/页面/数据/服务端拒绝 → 决定重发策略）
try:
    from core.failure import classify, backoff_for, CATS
except Exception:                        # 分类器缺失时退化为"全部重试 2 次"
    CATS = {}

    def classify(msg="", err=None):
        return {"cat": "other", "name": "未知原因", "retry": True, "max_try": 2,
                "backoff": 8.0, "action": "少量重试后标人工",
                "reason": (msg or "")[:120], "raw": (msg or "")[:200]}

    def backoff_for(cat, attempt):
        return min(8.0 * (2 ** max(0, attempt - 1)), 60.0)

APP = Path(__file__).resolve().parent
SETTINGS = APP / "settings.json"
DATA = APP / "data"                 # 2026-09-26：流水 / 已发记录统一放这儿

# ---------------- 2026-10-25 新增：模板偏好（通用模板标记）----------------
# 单独一个小文件，不动 settings.json / ui.json，避免影响既有读写逻辑。
#   {"universal": ["模板名", ...]}   → 标为「通用」的模板，任何分表都能用
TPLPREFS = DATA / "tpl_prefs.json"
# 2026-10-25：本地模板库（模板名 + 正文 + 阶段 + 停用标记）
TPLLIB = DATA / "tpl_library.json"


def load_tpl_library() -> dict:
    """读模板库。读不到就返回空壳，绝不影响启动。"""
    try:
        if TPLLIB.exists():
            d = json.loads(TPLLIB.read_text(encoding="utf-8"))
            if isinstance(d, dict):
                d.setdefault("templates", [])
                d.setdefault("retired_templates", [])
                return d
    except Exception:
        pass
    return {"templates": [], "retired_templates": []}


def load_tpl_prefs() -> dict:
    try:
        if TPLPREFS.exists():
            d = json.loads(TPLPREFS.read_text(encoding="utf-8"))
            if isinstance(d, dict):
                d.setdefault("universal", [])
                return d
    except Exception:
        pass
    return {"universal": []}


def save_tpl_prefs(d: dict) -> None:
    try:
        TPLPREFS.parent.mkdir(parents=True, exist_ok=True)
        TPLPREFS.write_text(json.dumps(d, ensure_ascii=False, indent=1),
                            encoding="utf-8")
    except Exception:
        pass
OUT_DIR = Path(os.environ.get("FUMAMX_OUT_DIR")
               or (Path(__file__).resolve().parent / "output" / "发送结果"))

COUNTRY_KEYS = ("国家", "地区", "country", "国别")
MAIL_KEYS = ("邮箱", "email", "e-mail", "mail", "邮件")
NAME_KEYS = ("名字", "姓名", "联系人", "客户名称", "客户名", "name", "contact", "客户")
COMPANY_KEYS = ("公司名称", "公司", "company", "企业", "单位")
NOISE = ("等级", "类型", "时间", "日期", "备注")     # 这些列不算客户名


def _tpl_key(name: str) -> str:
    """模板名归一化键（只用于去重）：去首尾空格、合并中间空白、不分大小写。

    孚盟目前只给模板名、不给模板 ID，所以「同名或同 ID 去重」里
    现在只能按名字归一化解；将来流水里若记了模板 ID，
    把这个函数改成返回 f"{tid}|{名字归一}" 即可，其余代码不用动。
    """
    return re.sub(r"\s+", " ", str(name or "").strip()).casefold()


def guess_col(header: list[str], keys: tuple[str, ...]) -> int:
    """优先整列头完全相等，其次包含关键词；含"等级/类型"等噪音词的一律跳过。"""
    cols = [str(h).strip() for h in header]
    low = [c.lower() for c in cols]

    def usable(i: int) -> bool:
        return not any(n in cols[i] for n in NOISE)

    for k in keys:
        k = k.lower()
        for i, h in enumerate(low):
            if h == k and usable(i):
                return i
    for k in keys:
        k = k.lower()
        for i, h in enumerate(low):
            if k in h and usable(i):
                return i
    return -1


def pick_header_row(rows: list) -> int:
    """自动判断表头在第几行（返回 0 基行号）。

    2026-09-29 用户反馈「表格列名和里面的内容对不上」。根因：
    原来固定 `self.header = rows[0]`，但客户表里两种排版都有——
      · 多数分表：第 1 行就是表头（名字 / 国家 / 公司 / 邮箱 …）；
      · 2510广交会、Sheet2：第 1 行是**标题行**（「重要客户」+ 末尾一列
        「国家_原值(备份)」），真正的表头在**第 2 行**（公司/名字/邮箱/电话/网站…）。
    后一种被固定取行 0 后：国家列被认成「国家_原值(备份)」那一列，公司列整个认不出来。

    判据：客户表一定有邮箱列，所以哪一行出现「邮箱 / mail / e-mail」就用哪一行。
    """
    for i in range(min(3, len(rows))):
        for v in rows[i]:
            t = str(v or "").strip().lower()
            if ("邮箱" in t or "mail" in t or "e-mail" in t):
                return i
    return 0


def _content_kind(vals: list) -> str:
    """看一列的实际内容，判断它究竟是哪种列。

    返回 country / mail / url / company / other。
    只用于**纠正列名显示**，不参与任何发送逻辑。
    """
    v = [str(x or "").strip() for x in vals if str(x or "").strip()]
    if len(v) < 3:
        return "other"
    n = len(v)
    if sum(1 for x in v if "@" in x) / n >= 0.5:
        return "mail"
    if sum(1 for x in v if x.lower().startswith(("http", "www."))) / n >= 0.5:
        return "url"
    try:
        from core import geo_alias
        hit = sum(1 for x in v if geo_alias.resolve_country(x)[0])
    except Exception:
        hit = 0
    rate = hit / n
    if rate >= 0.7:
        return "country"
    if rate <= 0.2:
        # 不像国家 → 再看像不像公司名（不是纯数字 / 日期）
        solid = sum(1 for x in v
                    if len(x) >= 2
                    and not x.replace(".", "").replace("/", "").replace("-", "").isdigit())
        if solid / n >= 0.6:
            return "company"
    return "other"


# 列名关键词 → 它「自称」是哪种列（用于和内容比对）
_NAME_KINDS = (
    ("country", ("国家", "地区", "country", "国别")),
    ("company", ("公司", "company", "企业", "单位")),
    ("mail", ("邮箱", "mail", "e-mail")),
)
_KIND2NAME = {"country": "国家", "company": "公司", "mail": "邮箱", "url": "网站"}


def quick_kind(vals: list) -> str:
    """极简版内容判型 —— 只为「一遍扫完就认出国家列」，绝不调用 geo_alias。

    2026-10-09 加。原来 _content_kind 逐格调 geo_alias.resolve_country()，
    一万格的列就是一万次调用，慢到不能用。这里只数三种一眼能认出来的形状：
      · 带 @      → 邮箱列
      · http/www  → 网站列
      · 2~3 个字母（US / UK / DE …）→ 典型国家写法
    够用就行：认不出就返回 other，由上层决定是否交给慢的 _content_kind。
    """
    v = [str(x or "").strip() for x in vals if str(x or "").strip()]
    if len(v) < 3:
        return "other"
    n = len(v)
    if sum(1 for x in v if "@" in x) / n >= 0.4:
        return "mail"
    if sum(1 for x in v if x.lower().startswith(("http", "www."))) / n >= 0.4:
        return "url"
    if sum(1 for x in v if re.fullmatch(r"[A-Za-z]{2,3}", x)) / n >= 0.3:
        return "country"
    return "other"


def fix_col_names(header: list, body: list, above: list = None) -> tuple[list, list]:
    """按实际内容纠正列名 —— 只改**显示用的列名**，绝不改原始数据值。

    用户要求：「把本应表示公司信息、列名却写成『国家』的那一列，列名改为『公司』；
               把国家信息的那一列，列名改为『国家』。」

    三条规则：
      ① 列名空 → 按内容补一个（国家 / 公司 / 邮箱 / 网站）；
      ② 列名自称「国家/地区/country」，但内容 70%+ 是公司名 → 改名「公司」；
      ③ 列名自称「公司/company」，但内容 70%+ 是国家 → 改名「国家」；
      ④ 出现重名列（例：Sheet2 里第 1 列和第 11 列都叫「国家」）→
         借用上一行（标题行）对应列的名字区分，如「国家_原值(备份)」。
    只动 country↔company 这一对（其余列名原样保留），避免误伤正常列。

    above：表头**上一行**的单元格（有的表那一行才是完整列名），用于重名消歧。

    返回 (纠正后的列名表, 改动说明列表)
    """
    out = [str(h or "").strip() for h in header]
    notes = []
    seen: dict[str, int] = {}
    sample = body[:60]
    for ci in range(len(out)):
        vals = [r[ci] for r in sample if ci < len(r)]
        kind = _content_kind(vals)
        if kind == "other":
            if out[ci]:
                seen.setdefault(out[ci], ci)
            continue
        name = out[ci]
        low = name.lower()
        declared = ""
        for k, keys in _NAME_KINDS:
            if any(x in low for x in keys):
                declared = k
                break
        if not name:
            # 只给「判定极可靠」的三类补列名：国家（城市/带后缀都能认）、邮箱、网址。
            # company **不自动补名** —— 实测会把「名字 / 客户详情 / 备注 / 发信日期」
            # 一律判成 company，容易给空列乱起名（潜力商机客户 第 10 列就中过招）。
            if kind not in ("country", "mail", "url"):
                continue
            new = _KIND2NAME[kind]
            if new in out[:ci]:
                continue          # 已经有一列叫这个名字了，别再造出重名列
            out[ci] = new
            notes.append(f"第 {ci + 1} 列原本没有列名，内容像「{kind}」→ 命名为「{new}」")
        elif (declared in ("country", "company") and kind in ("country", "company")
              and declared != kind):
            new = _KIND2NAME[kind]
            out[ci] = new
            notes.append(f"第 {ci + 1} 列列名「{name}」与内容不符"
                         f"（内容实为{'公司信息' if kind == 'company' else '国家信息'}）"
                         f"→ 显示为「{new}」")
        # ---- 重名消歧：同表里两列同名会让人分不清哪个是哪个 ----
        nm = out[ci]
        if nm and nm in seen:
            alt = str(above[ci] or "").strip() if (above and ci < len(above)) else ""
            if alt and alt != nm:
                out[ci] = alt
                notes.append(f"第 {ci + 1} 列与前面的列重名「{nm}」，"
                             f"改用上一行标题行的名字「{alt}」")
            else:
                out[ci] = f"{nm}(2)"
                notes.append(f"第 {ci + 1} 列与前面的列重名「{nm}」，显示为「{nm}(2)」")
        if out[ci]:
            seen.setdefault(out[ci], ci)
    return out, notes


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("孚盟MX 群发助手")
        # 2026-09-26：Tk 回调里抛异常时默认只写 stderr（pythonw 下等于石沉大海），
        # 表现就是「点按钮毫无反应」。这里接管，让异常弹窗提示 + 落日志。
        self.report_callback_exception = self._on_tk_error
        self.geometry("980x720")
        self.q: queue.Queue = queue.Queue()
        self.engine: Engine | None = None
        self.worker: threading.Thread | None = None
        self.rows: list[list[str]] = []
        self.header: list[str] = []
        self.results: list[dict] = []
        self.cfg = self._load_cfg()
        self.ui = self._load_ui()
        # 分批发送（2026-10-09）：本次运行内每批的位次记账。
        # 只记「本分表从第几位开始跑、跑了几位」等计数，真正的落盘进度在 ui.json。
        self._batch_run = False         # 本轮是否启用了「本批发送数量」
        self._batch_hint = ""           # 最近一次批量进度提示（刷新横幅用）
        self._cur_prog_extra = ""       # 当前分表的已完成批次数文字
        self._groups: list[dict] = []          # 折叠组：{"btn":…, "body":…, "open":bool}
        (APP / "logs").mkdir(exist_ok=True)
        self._logf = open(APP / "logs" / ("run_" + now_cn().strftime("%Y%m%d") + ".log"),
                          "a", encoding="utf-8")

        # 定时队列：先建好（建页时要刷新列表）；重启后恢复上次没发完的任务
        self.store = TaskStore(APP / "data" / "schedule.json")
        self.store.recover_interrupted(self.log)
        self.scheduler = Scheduler(None, self.store, log=self.log)
        self.scheduler.on_change = lambda: self.q.put(("call", self.refresh_sched))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.nb = nb
        self.f1 = self._tab_connect(nb)
        self.f2 = self._tab_sheet(nb)
        self.f3 = self._tab_settings(nb)
        self.f4 = self._tab_sched(nb)
        self.f5 = self._tab_run(nb)
        self.f6 = self._tab_history(nb)
        nb.add(self.f1, text="① 连接孚盟")
        nb.add(self.f2, text="② 客户表")
        nb.add(self.f3, text="③ 发送设置")
        nb.add(self.f4, text="④ 定时发送")
        nb.add(self.f5, text="⑤ 立即发送")
        nb.add(self.f6, text="⑥ 历史发送记录")

        self.scheduler.start()

        # ---- 2026-09-29：一轮跑完的「现场」----
        # 用户诉求①：「自查并重发失败邮件」只补**刚刚发完的当前分表**，
        #   不要一按就把待人工清单里其他分表的陈年失败一起翻出来。
        # 所以每开一轮就清空重建下面这三个，收尾时供补发按钮取用。
        self._last_round_sheet = ""     # 刚发完的分表名
        self._last_round_at = ""        # 本轮开始时间（HH:MM:SS）
        self._last_round_failed = {}    # {邮箱小写: {email,country,reason,name,action,retry,src}}

        self.bind("<Configure>", self._on_resize)
        # 2026-09-26：一切到⑥页就重读一遍（新模板、刚发完的记录马上能看到）
        nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # 底部常驻工具条：置顶开关 + 窗口尺寸（用户说过：操作完网站还得重新打开窗口）
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        self.var_top = tk.BooleanVar(value=bool(self.ui.get("topmost", True)))
        ttk.Checkbutton(bar, text="窗口置顶（操作孚盟网站时也压在最前面，不用重新打开）",
                        variable=self.var_top, command=self._apply_topmost).pack(side="left")
        ttk.Button(bar, text="紧凑小窗", command=self.do_compact).pack(side="left", padx=10)
        ttk.Button(bar, text="还原大窗", command=self.do_expand).pack(side="left")

        geo = self.ui.get("geometry", "")
        # 2026-09-29：ui.json 里若存的是 "1x1+…"（最小化时存坏的），别用它，
        # 否则一打开就是个小方块。
        if geo and not str(geo).startswith("1x1"):
            try:
                self.geometry(geo)
            except Exception:
                pass
        self._apply_topmost()
        self.refresh_sent_count()        # 已发记录条数
        self._restore_session()          # 恢复上次的客户表/分表/时间，并提示断点

        self.after(150, self._drain)
        # 启动后自动对一次账：日志里成功、但已发记录里没有的补进去（防漏记导致重发）
        self.after(1200, self._reconcile_sent)
        self.after(1500, self._sched_tick)   # 定时列表每 2 秒自动刷新进度
        self.after(1800, self._hist_boot)    # ⑥ 历史发送记录：开界面自动查一次
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- 窗口 ----------------
    def _load_ui(self) -> dict:
        p = APP / "ui.json"
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _save_ui(self):
        try:
            self.ui["topmost"] = bool(self.var_top.get())
            # 2026-09-29 修复：窗口**最小化**时 Tk 的 geometry() 会返回 "1x1+…"，
            # 原样存盘 → 下次打开变成一个 1x1 的小方块（用户的 ui.json 就中招过）。
            # 这里只在窗口正常显示时才记录尺寸；最小化/最大化时保留上次的尺寸。
            try:
                st = str(self.state())
            except Exception:
                st = "normal"
            g = ""
            try:
                g = self.geometry()
            except Exception:
                pass
            if st == "normal" and g and not str(g).startswith("1x1"):
                self.ui["geometry"] = g
            self._snapshot()          # 把当前「客户表/分表/模板/主题/时间/模式」记进 ui.json
            (APP / "ui.json").write_text(
                json.dumps(self.ui, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:
            pass

    # ---------------- 现场记忆（关掉重开自动恢复）----------------
    def _prog_map(self) -> dict:
        """ui.json 里的「各分表进度表」：{"<客户表>|<分表>": {...}}。

        2026-10-09 新增。以前进度只存在 ui["sess"] 这一个键里 —— 换一张分表
        就把上一张的进度冲掉了，用户切回来时断点是空的。
        现在按「客户表全路径 + 分表名」分别记，各分表互不干扰。
        """
        m = self.ui.get("prog")
        if not isinstance(m, dict):
            m = {}
            self.ui["prog"] = m
        return m

    def _prog_key(self, file: str = None, sheet: str = None) -> str:
        """一个分表的唯一标识。不带参数 = 当前界面上选的那张。"""
        f = self.var_file.get() if file is None and hasattr(self, "var_file") else (file or "")
        s = (str(self.cmb_sheet.get())
             if sheet is None and hasattr(self, "cmb_sheet") else (sheet or ""))
        return f"{f}|{s}"

    def _prog(self, file: str = None, sheet: str = None) -> dict:
        """取某个分表的进度字典（没有就按旧 sess 兼容升级一份）。"""
        k = self._prog_key(file, sheet)
        m = self._prog_map()
        d = m.get(k)
        if not isinstance(d, dict):
            d = {}
            # 兼容：旧版本只存了 sess 一个分表。若它对得上当前这张，就认领过来，
            # 用户不会因为这次升级把断点丢了。
            s = self.ui.get("sess") or {}
            if (k and s.get("file") and s.get("sheet")
                    and k == f"{s.get('file')}|{s.get('sheet')}"):
                d = {kk: s.get(kk) for kk in
                     ("pos", "total", "done", "at", "batches", "finished")
                     if s.get(kk) is not None}
            m[k] = d
        return d

    def _prog_mark(self, key: str, **kw):
        """按分表 key 记进度（纯 dict 操作 → 后台线程也能安全调）。"""
        try:
            m = self._prog_map()
            d = m.get(key)
            if not isinstance(d, dict):
                d = {}
                m[key] = d
            d.update(kw)
        except Exception:
            pass

    def _sess(self) -> dict:
        """ui.json 里的「上次现场」字典（纯 dict 操作，后台线程也能安全改 pos/total）"""
        s = self.ui.get("sess")
        if not isinstance(s, dict):
            s = {}
            self.ui["sess"] = s
        return s

    def _sess_mark(self, **kw):
        self._sess().update(kw)       # 只动 dict，不碰 Tk 控件 → 线程安全

    def _snapshot(self) -> dict:
        """抓当前界面上的现场（只能在主线程调，因为要读 Tk 变量）"""
        s = self._sess()
        try:
            s["file"] = self.var_file.get()
            s["sheet"] = self.cmb_sheet.get()
            s["tpl"] = self.cmb_tpl.get()
            s["subj"] = self.cmb_subj.get()
            s["when"] = self.var_when.get()
            s["mode"] = self.cmb_mode.current()
            if getattr(self, "var_batch", None) is not None:
                s["batch"] = self.var_batch.get().strip()
            # ★ 2026-10-06：时间基准也记进现场。以前只写在 settings.json，而
            #   settings.json 只在「关窗口/确定设置」时才落盘 —— 用户选了
            #   「北京时间」但没走那两步的话，重启就悄悄退回「当地时间」，
            #   表现就是"我明明选了，发出来还是当地 13:46"。
            s["basis"] = str(
                (getattr(self, "var_basis", None) and self.var_basis.get()) or "")
            # ★ 2026-10-09：当前分表的进度也同步进 sess（旧结构，兼容用）
            try:
                d = self._prog(self.var_file.get(), self.cmb_sheet.get())
                for kk in ("pos", "total", "done", "batches", "finished"):
                    if kk in d:
                        s[kk] = d[kk]
            except Exception:
                pass
            s["at"] = now_cn().strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        return s

    def _restore_session(self):
        """重开时把上次的表/分表/模板/主题/时间填回去，并给出断点提示"""
        s = self.ui.get("sess") or {}
        if s.get("when"):
            try:
                self.var_when.set(s["when"])
            except Exception:
                pass
        if isinstance(s.get("mode"), int) and s["mode"] >= 0:
            try:
                self.cmb_mode.current(s["mode"])
            except Exception:
                pass
        # ★ 2026-10-06：恢复时间基准（现场里优先，其次 settings.json）
        b = str(s.get("basis") or getattr(self.cfg, "when_basis", "local") or "local")
        if b in ("beijing", "local"):
            try:
                if getattr(self, "var_basis", None) is not None:
                    self.var_basis.set(b)
                self.cfg.when_basis = b
                if self.engine:
                    self.engine.cfg.when_basis = b
            except Exception:
                pass
        # 模板/主题要等连上孚盟读到列表后才能回填，先挂起来
        self._pending_tpl = s.get("tpl", "") or ""
        self._pending_subj = s.get("subj", "") or ""

        f = s.get("file", "")
        if f and Path(f).exists():
            try:
                names = xlsxlight.sheet_names(f)
                self.var_file.set(f)
                self.cmb_sheet["values"] = names
                want = s.get("sheet", "")
                self.cmb_sheet.set(want if want in names else names[0])
                self.do_load_sheet()
                self.log(f"[现场] 已恢复上次的客户表：{Path(f).name} / {self.cmb_sheet.get()}")
            except Exception as e:
                self.log("[现场] 恢复客户表失败：", e)
        elif f:
            self.log(f"[现场] 上次的客户表找不到了：{f}（请重新浏览选择）")
        self.refresh_resume()

    def refresh_resume(self):
        """判断「断点续跑」能不能用：只有当表+分表和上次完全一致时才允许。

        2026-10-09：改成读**本分表自己的**进度（ui["prog"]["<表>|<分表>"]），
        换一张分表不会再把上一张的进度冲掉；本分表已跑完时明确提示「已完成」。
        """
        if not hasattr(self, "var_resume_tip"):
            return
        cur_file = self.var_file.get() if hasattr(self, "var_file") else ""
        cur_sheet = str(self.cmb_sheet.get()) if hasattr(self, "cmb_sheet") else ""
        d = self._prog(cur_file, cur_sheet)
        pos, total = int(d.get("pos") or 0), int(d.get("total") or 0)
        done = int(d.get("done") or pos)
        nb = int(d.get("batches") or 0)
        fin = bool(d.get("finished")) or (total > 0 and pos >= total)
        at = d.get("at", "")
        bt = f"（已发 {nb} 批）" if nb else ""
        try:
            if not cur_file or total <= 0:
                self.var_resume_tip.set("")
                self.cb_resume.config(text="断点续跑（换表/换分表后不生效）",
                                      state="disabled")
                self.var_resume.set(False)
            elif fin:
                self.var_resume_tip.set(
                    f"✔ 本分表已全部发完：{done}/{total} 位{bt}"
                    + (f"（{at}）" if at else "")
                    + " ｜想再跑一轮请点「标为未完成」")
                self.cb_resume.config(text="从第 1 位重新开始", state="normal")
                self.var_resume.set(False)
            elif 0 < pos < total:
                self.var_resume_tip.set(
                    f"本分表上次跑到第 {pos}/{total} 位{bt}"
                    + (f"（{at}）" if at else "")
                    + f"，后面还有 {total - pos} 位没处理")
                self.cb_resume.config(
                    text=f"从第 {pos + 1} 位开始（跳过前面 {pos} 位）", state="normal")
                self.var_resume.set(True)
            else:
                self.var_resume_tip.set("")
                self.cb_resume.config(text="断点续跑（换表/换分表后不生效）",
                                      state="disabled")
                self.var_resume.set(False)
        except Exception:
            pass

    def _apply_topmost(self):
        try:
            self.attributes("-topmost", bool(self.var_top.get()))
        except Exception:
            pass

    # ---- 弹窗统一入口：必须叠在主窗口之上，否则模态框躲在置顶主窗后面 = 整窗假死 ----
    # 2026-10-04 修：原来这里第一行是 attributes("-topmost", False)（先取消置顶），
    # 浏览器最大化在前时 lift/focus_force 又抢不过它 → 模态框整个躲在浏览器后面，
    # 主线程被 messagebox 锁死、宿主窗口被禁用，表现为"卡死且关不掉"。
    # 现在改为：弹窗期间强制置顶（弹完由 _apply_topmost 还原用户设置）。
    def _front(self):
        try:
            self.attributes("-topmost", True)
            self.lift()
            self.focus_force()
            self.update_idletasks()
        except Exception:
            pass

    def _toast(self, msg, kind="info", ms=8000):
        """非模态轻提示：2026-10-04 加（用户要求：不要锁住主窗口）。

        凡是"告知类"提示（登录成功、发送结果、导出结果、错误…）一律走这里：
        overrideredirect 的 Toplevel，不 grab_set、不阻塞主线程、可点"知道了"关闭、
        到时自动消失、同时最多 3 条（多了自动收掉最老的）。
        只有 _ask（必须拿到是/否来继续流程）才保留模态框，且走 _front 强制置顶。
        """
        palette = {
            "info":  ("#fffbe6", "#5a4a00", "#c8a200"),
            "ok":    ("#eaf7ea", "#1f4d1f", "#3a9b3a"),
            "warn":  ("#fff2e0", "#6b3b00", "#e08020"),
            "error": ("#fdeaea", "#6b1010", "#c02020"),
        }
        bg, fg, bd = palette.get(kind, palette["info"])
        life = ms if kind != "error" else max(ms, 14000)

        def build():
            try:
                if not hasattr(self, "_toasts"):
                    self._toasts = []
                # 最多同时 3 条，多了先收掉最老的
                while len(self._toasts) >= 3:
                    old = self._toasts.pop(0)
                    try:
                        old.destroy()
                    except Exception:
                        pass

                prev_fg = _fg_hwnd()          # 记下用户此刻正在用的窗口
                win = tk.Toplevel(self)
                win.overrideredirect(True)
                try:
                    win.attributes("-topmost", True)
                except Exception:
                    pass
                outer = tk.Frame(win, bg=bd)
                outer.pack(fill="both", expand=True)
                lab = tk.Label(outer, text=msg, justify="left", wraplength=400,
                               bg=bg, fg=fg, padx=14, pady=10)
                lab.pack(fill="both", expand=True)
                btn = tk.Button(outer, text="知道了", command=win.destroy,
                                bg=bg, fg=fg, relief="raised", bd=1,
                                activebackground=bg, activeforeground=fg)
                btn.pack(fill="x")

                def bye():
                    try:
                        self._toasts.remove(win)
                    except Exception:
                        pass
                    try:
                        win.destroy()
                    except Exception:
                        pass

                btn.config(command=bye)
                lab.bind("<Button-1>", lambda e: bye())
                win.update_idletasks()
                w = max(win.winfo_reqwidth(), 280)
                h = win.winfo_reqheight()
                x = max(0, self.winfo_screenwidth() - w - 30)
                y = max(0, self.winfo_screenheight() - h - 80)
                win.geometry("%dx%d+%d+%d" % (w, h, x, y))
                win.after(life, bye)
                self._toasts.append(win)
                # 轻提示弹出时 Windows 会把焦点抢过去 —— 弹完立刻还给用户，
                # 免得他正在别的程序里打字时被打断（2026-10-06）。
                win.after(50, lambda: _fg_back(prev_fg))
            except Exception:
                pass
        try:
            self.after(0, build)
        except Exception:
            pass

    def _info(self, title, msg):
        """2026-10-04：告知类提示不再用 messagebox（会锁死主线程），改轻提示。"""
        self._toast(msg, kind="info")

    def _err(self, title, msg):
        self._toast(msg, kind="error")

    def _warn(self, title, msg):
        self._toast(msg, kind="warn")

    def _ask(self, title, msg):
        """必须拿到是/否才能继续流程的确认 —— 保留模态框。
        靠 _front() 强制置顶保证它不会被浏览器盖住（盖住 = 看着像卡死）。"""
        self._front()
        try:
            return messagebox.askyesno(title, msg, parent=self)
        finally:
            self._apply_topmost()

    def do_compact(self):
        """缩成一个角上的小条：自动折叠所有分组，内容再多也不会溢出。"""
        self._collapse_all(closed=True)
        self.geometry("560x300+{}+60".format(max(0, self.winfo_screenwidth() - 580)))
        self._save_ui()

    def do_expand(self):
        self._collapse_all(closed=False)
        self.geometry("980x720+80+60")
        self._save_ui()

    def _come_back(self):
        """把本窗口提到同层最前 —— 但不置顶、不抢焦点（2026-10-06 重写）。

        原来这里是 attributes("-topmost", True) + lift() + focus_force()：
          · topmost 设上之后**再也没有还原**过 → 本窗口永远压在所有程序之上；
          · focus_force() 会抢走当前前台窗口的键盘焦点。
        两者叠加，用户开着别的程序干活时会被反复打断（表现为刚打开的窗口
        "几秒就没了"）。现在只 lift()：只调整本程序内部的层级，既不置顶、
        也不碰别人的焦点；要不要置顶完全跟随界面上那个「置顶」勾选框。
        """
        try:
            self.lift()
        except Exception:
            pass

    # ---------------- 基础 ----------------
    def log(self, *a):
        line = " ".join(str(x) for x in a)
        self.q.put(("log", line))
        try:
            self._logf.write(now_cn().strftime("%H:%M:%S ") + line + "\n")
            self._logf.flush()
        except Exception:
            pass

    # ---------------- 折叠组 ----------------
    def _collapsible(self, parent, title: str, open_: bool = True):
        """▸/▾ 折叠组：小窗口下把不相关的收起来，内容再多也不溢出。"""
        box = ttk.Frame(parent)
        lab = tk.Label(box, text=("▾ " if open_ else "▸ ") + title,
                       anchor="w", cursor="hand2", font=("Microsoft YaHei UI", 10, "bold"))
        lab.pack(fill="x")
        body = ttk.Frame(box, padding=(14, 2))
        if open_:
            body.pack(fill="x")

        def toggle(_e=None):
            open_ = not bool(body.winfo_ismapped())
            lab.config(text=("▾ " if open_ else "▸ ") + title)
            if open_:
                body.pack(fill="x")
            else:
                body.pack_forget()
            self._save_ui()

        lab.bind("<Button-1>", toggle)
        g = {"title": title, "lab": lab, "body": body, "open": open_}
        self._groups.append(g)
        return box, body

    def _collapse_all(self, closed: bool):
        for g in self._groups:
            if closed:
                g["body"].pack_forget()
                g["lab"].config(text="▸ " + g["title"])
            else:
                g["body"].pack(fill="x")
                g["lab"].config(text="▾ " + g["title"])

    def _on_resize(self, e):
        """窗口变窄时把长说明文字的换行宽度跟着收，避免横向溢出。"""
        w = max(240, e.width - 60)
        for lab in getattr(self, "_wraplabels", []):
            try:
                lab.configure(wraplength=w)
            except Exception:
                pass

    def _drain(self):
        # 2026-09-22 加固：任何一条消息处理出错都不能让更新循环死掉
        # （之前循环一旦抛异常就永久停摆 → 后台在发、界面却永远不动）
        try:
            while True:
                kind, payload = self.q.get_nowait()
                try:
                    self._dispatch(kind, payload)
                except Exception:
                    # 界面消息出错只记文件日志，吞掉保证循环存活。
                    # 2026-09-23：必须把异常详情写进去，否则只看到"[drain错误] mxrun"
                    # 却不知道哪里错，等于没法排查。
                    try:
                        import traceback as _tb
                        self._logf.write(now_cn().strftime("%H:%M:%S ")
                                         + "[drain错误] " + kind + " → "
                                         + _tb.format_exc(limit=3).replace("\n", " | ")
                                         + "\n")
                        self._logf.flush()
                    except Exception:
                        pass
        except queue.Empty:
            pass
        self.after(150, self._drain)

    def _dispatch(self, kind, payload):
        if kind == "log":
            self.txt_log.insert("end", payload + "\n")
            self.txt_log.see("end")
        elif kind == "status":
            self.var_status.set(payload)
        elif kind == "mxrun":
            self._mx_update(payload)
        elif kind == "progress":
            self.pb["value"] = payload
        elif kind == "msg":
            self._err("孚盟MX 群发助手", payload)
        elif kind == "info":
            self._info("孚盟MX 群发助手", payload)
        elif kind == "toast":
            # 2026-10-04：非模态轻提示。走这里的消息绝不能弹 messagebox，
            # 否则会锁死主线程（浏览器在前时 = 整窗假死且关不掉）。
            # payload 可以是纯字符串，也可以是 (msg, kind, 停留毫秒)。
            if isinstance(payload, (tuple, list)):
                self._toast(*payload)
            else:
                self._toast(payload)
        elif kind == "sent_report":
            self._on_sent_report(payload)
        elif kind == "call":
            payload()          # 工作线程要做的界面更新，统一在主线程执行
        elif kind == "done":
            self.btn_start.config(state="normal")
            # 停止键保持可用：任何时候都能点，不再变灰

    def _load_cfg(self) -> EngineConfig:
        c = EngineConfig(profile_dir=str(APP / "data" / "mxprofile"))
        if SETTINGS.exists():
            try:
                d = json.loads(SETTINGS.read_text(encoding="utf-8"))
                for k, v in d.items():
                    if hasattr(c, k):
                        setattr(c, k, v)
            except Exception:
                pass
        return c

    def _save_cfg(self):
        d = {k: getattr(self.cfg, k) for k in
             ("port", "login_url", "login_wait_min", "step_delay", "interval_min",
              "interval_max", "human_mode", "spread_gap", "spread_mode",
              "clear_cc", "work_start", "work_end",
              "only_work_hours", "skip_sent", "dry_run", "browser_exe",
              "fast_mode", "verify_every", "require_chip", "auto_restore",
              "auto_revive", "auto_raise", "when_basis")}
        # 2026-10-09：分批发送「整轮已跑完」的标记，跟进度一起持久化，
        # 重开程序也认得哪些分表已经发完（避免重复跑）。
        d["batch_done"] = dict(getattr(self.cfg, "batch_done", {}) or {})
        SETTINGS.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")

    def _on_close(self):
        try:
            self._save_cfg()
            self._save_ui()
        except Exception:
            pass
        self.destroy()

    # ---------------- ① 连接 ----------------
    def _tab_connect(self, nb):
        f = ttk.Frame(nb, padding=12)
        self.var_status = tk.StringVar(value="未连接")
        ttk.Label(f, text="状态：").grid(row=0, column=0, sticky="w")
        ttk.Label(f, textvariable=self.var_status, foreground="#c0392b").grid(row=0, column=1, sticky="w")
        ttk.Button(f, text="打开孚盟MX 并登录", command=self.do_connect).grid(row=0, column=2, padx=6)
        ttk.Button(f, text="我登录好了，检测一下", command=self.do_check).grid(row=0, column=3, padx=6)
        ttk.Button(f, text="读取模板列表", command=self.do_templates).grid(row=0, column=4, padx=6)
        ttk.Button(f, text="读取历史主题", command=self.do_subjects).grid(row=0, column=5, padx=6)

        g_tpl, gf = self._collapsible(f, "模板与主题", open_=True)
        g_tpl.grid(row=1, column=0, columnspan=5, sticky="ew")
        f.columnconfigure(0, weight=1)
        ttk.Label(gf, text="邮件模板（正文/附件由孚盟模板带入，本地不改）").grid(
            row=0, column=0, sticky="w")
        self.cmb_tpl = ttk.Combobox(gf, width=48, state="readonly")
        self.cmb_tpl.grid(row=1, column=0, sticky="ew")
        self.cmb_tpl.set(self.cfg.extras.get("template", "") if hasattr(self.cfg, "extras") else "")
        # 2026-10-25：换模板后刷新③页的跟进建议（③页还没建时静默跳过）
        self.cmb_tpl.bind(
            "<<ComboboxSelected>>",
            lambda e: (self.refresh_tpl_advice()
                       if hasattr(self, "var_tpl_advice") else None))

        ttk.Label(gf, text="主题（可下拉选历史主题，也可自己写）").grid(
            row=2, column=0, sticky="w", pady=(10, 0))
        self.cmb_subj = ttk.Combobox(gf, width=70)
        self.cmb_subj.grid(row=3, column=0, sticky="ew")

        note = ttk.Label(f, text="说明：登录只在您打开的浏览器窗口里手工进行，本程序不保存、不读取您的账号密码。",
                         foreground="#666", wraplength=900, justify="left")
        note.grid(row=5, column=0, columnspan=5, sticky="w", pady=(18, 0))
        self._wraplabels = getattr(self, "_wraplabels", []) + [note]
        return f

    def do_connect(self):
        def run():
            try:
                self.cfg.profile_dir = str(APP / "data" / "mxprofile")
                self.cfg.browser_exe = ""          # 用本机 Edge（不再用 OCS 的浏览器）
                self.engine = Engine(self.cfg, log=self.log, data_dir=APP / "data")
                ok = self.engine.connect(wait_login=True)
                # 2026-10-04：置顶必须排在 info 之前。
                # _drain 是顺序 while 循环取队列，info 会弹模态框把主线程卡住，
                # 排在它后面的 call 永远执行不到 → 弹窗与主窗一起被浏览器盖住 = 整窗假死。
                self.q.put(("call", lambda: self.after(300, self._come_back)))
                if ok:
                    self.q.put(("status", "已登录"))
                    self.log("已连接孚盟MX，接下来点「读取模板列表」和「读取历史主题」。")
                    self.q.put(("toast", "已连接孚盟MX。\n接下来点「读取模板列表」和「读取历史主题」。\n"
                                         "（本窗口会置顶，浏览器在它后面，不用重新打开本程序）"))
                else:
                    self.q.put(("status", "未检测到登录，请在浏览器里登录后再点「检测一下」"))
                # 浏览器最大化后把本窗口拉回前面（after 不能在线程里调，走队列回主线程）
                self.q.put(("call", lambda: self.after(300, self._come_back)))
            except Exception as e:
                import traceback
                traceback.print_exc()          # 落到 logs/app_err.log
                self.log("连接失败:", repr(e))
                self.q.put(("status", "连接失败"))
                self.q.put(("msg", "连接失败：\n" + str(e)))
        threading.Thread(target=run, daemon=True).start()

    def do_check(self):
        def run():
            try:
                if not self.engine:
                    self.q.put(("msg", "请先点一次「打开孚盟MX 并登录」，\n再点本按钮检测。"))
                    return
                ok = self.engine.check_login()
                if ok:
                    self.q.put(("status", "已登录"))
                    self.q.put(("info", "已登录！接下来点「读取模板列表」和「读取历史主题」。"))
                else:
                    self.q.put(("status", "未登录"))
                    self.q.put(("msg", "还是没检测到登录。\n请确认浏览器窗口里已经进入孚盟MX 主界面"
                                       "（地址栏不含 /login），再点本按钮。"))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.log("检测失败:", repr(e))
                self.q.put(("status", "检测失败"))
                self.q.put(("msg", "检测失败：\n" + str(e)))
        threading.Thread(target=run, daemon=True).start()

    def _guard(self) -> bool:
        """读取类操作互斥：连点会在同一个标签页上抢 CDP，各自等 20~25 秒超时，
        表现就是"反应很久还读不出来"。正在跑的时候直接挡掉后来的点击。"""
        if getattr(self, "_reading", False):
            self._info("稍等", "上一次读取还没结束，正在操作孚盟页面。\n"
                                      "请勿重复点击（重复点会互相抢，反而更慢/失败）。")
            return False
        self._reading = True
        return True

    def do_templates(self):
        def run():
            try:
                names = self.engine.templates()

                def apply():
                    self.cmb_tpl["values"] = names
                    want = getattr(self, "_pending_tpl", "") or ""
                    self.cmb_tpl.set(want if want in names else (names[0] if names else ""))
                    if want and want in names:
                        self.log(f"[现场] 已恢复上次的模板：{want}")
                    # 2026-09-26：把这份"当前模板"也喂给⑥页下拉（只写缓存+刷新下拉，
                    # 不动 cmb_tpl 本身 → ① 页的读取/选择/保存逻辑完全不受影响）
                    try:
                        self.ui["tpl_list"] = [str(x) for x in (names or [])]
                        self._hist_fill_filters()
                    except Exception:
                        pass
                # Tkinter 控件只能在主线程动，线程里直接改会锁死整个窗口
                self.q.put(("call", apply))
                self.log("模板:", names)
                # 2026-10-04：这些提示原来走 messagebox（会锁死主线程）。
                # 按用户要求改成右下角非模态提示。
                if not names:
                    self.q.put(("toast", (
                        "没读到模板。\n"
                        "请到浏览器里确认左侧能看到「孚盟邮」，再点一次；\n"
                        "还不行就刷新孚盟页面后再点。", "warn", 20000)))
                elif len(names) == 1 and ("账号" in names[0]):
                    self.q.put(("toast", (
                        f"只读到 1 个：{names[0]}\n"
                        "这多半是孚盟的「账号切换」下拉，不是模板列表。\n"
                        "请到浏览器里点一下左侧「孚盟邮」，再回来重点一次。",
                        "warn", 20000)))
                else:
                    self.q.put(("toast", (f"读到模板 {len(names)} 个：\n"
                                          + "、".join(names[:6])
                                          + ("…" if len(names) > 6 else ""),
                                          "ok", 9000)))
            except Exception as e:
                self.log("读取模板失败:", e)
                self.q.put(("toast", (f"读取模板失败：{type(e).__name__}: {e}",
                                      "warn", 20000)))
            finally:
                self._reading = False
        if not self.engine:
            self._info("提示", "请先点「打开孚盟MX 并登录」")
            return
        if not self._guard():
            return
        threading.Thread(target=run, daemon=True).start()

    def do_subjects(self):
        def run():
            try:
                items = self.engine.subjects(self.cmb_tpl.get())

                def apply():
                    self.cmb_subj["values"] = items
                    want = getattr(self, "_pending_subj", "") or ""
                    self.cmb_subj.set(want if want in items else (items[0] if items else ""))
                    if want and want in items:
                        self.log(f"[现场] 已恢复上次的主题：{want[:40]}")
                self.q.put(("call", apply))   # 同上：控件更新回主线程
                self.log("历史主题:", items[:5])
                if not items:
                    self.q.put(("toast", ("没读到历史主题。\n"
                                          "（写信页可能没打开成功，看日志里有没有"
                                          "「未能进入写信页」）", "warn", 18000)))
                else:
                    self.q.put(("toast", (f"读到历史主题 {len(items)} 条", "ok", 8000)))
            except Exception as e:
                self.log("读取历史主题失败:", e)
                self.q.put(("toast", (f"读取历史主题失败：{type(e).__name__}: {e}\n"
                                      "常见原因：模板下拉里选的不是真模板。\n"
                                      "先点「读取模板列表」，再从下拉选一个。",
                                      "warn", 22000)))
            finally:
                self._reading = False
        if not self.engine:
            self._info("提示", "请先点「打开孚盟MX 并登录」")
            return
        if not self._guard():
            return
        threading.Thread(target=run, daemon=True).start()

    # ---------------- ② 客户表 ----------------
    def _tab_sheet(self, nb):
        f = ttk.Frame(nb, padding=12)
        self.var_file = tk.StringVar()
        ttk.Label(f, text="客户表：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.var_file, width=70).grid(row=0, column=1, sticky="w")
        ttk.Button(f, text="浏览...", command=self.do_browse).grid(row=0, column=2, padx=6)

        ttk.Label(f, text="工作表：").grid(row=1, column=0, sticky="w", pady=6)
        self.cmb_sheet = ttk.Combobox(f, width=30, state="readonly")
        self.cmb_sheet.grid(row=1, column=1, sticky="w")
        self.cmb_sheet.bind("<<ComboboxSelected>>", lambda e: self.do_load_sheet())

        g_col, gf = self._collapsible(f, "列对应（自动识别，可手改）", open_=False)
        g_col.grid(row=2, column=0, columnspan=5, sticky="ew")
        self.cmb_country = ttk.Combobox(gf, width=18, state="readonly")
        self.cmb_mail = ttk.Combobox(gf, width=18, state="readonly")
        self.cmb_name = ttk.Combobox(gf, width=18, state="readonly")
        self.cmb_company = ttk.Combobox(gf, width=18, state="readonly")
        for i, (lab, c) in enumerate((("国家", self.cmb_country), ("邮箱", self.cmb_mail),
                                      ("客户名", self.cmb_name), ("公司", self.cmb_company))):
            ttk.Label(gf, text=lab).grid(row=0, column=i, sticky="w")
            c.grid(row=1, column=i, sticky="ew", padx=2)
            gf.columnconfigure(i, weight=1)

        # 2026-09-28 用户反馈：客户表列多（国家/公司/名字/邮箱/电话/网址…），
        # 右边几列被窗口截断，鼠标又划不过去。
        # 原因有两层：① 只挂了纵向滚动条；② Treeview 的"请求宽度"= 列宽合计，
        #   会顺着 grid 传给父容器把整页撑宽（实测撑到 1549px），表格和容器一样宽，
        #   自然就"没有东西可滚"。
        # 修法：让 Treeview 用 **place** 铺满它的容器 —— place 不参与尺寸协商，
        # 所以列宽合计再大也不会把页面撑开；超出部分由横向滚动条接管。
        tree_wrap = ttk.Frame(f)
        tree_wrap.grid(row=5, column=0, columnspan=4, sticky="nsew", pady=10)
        # 关键：不接受子控件的尺寸请求 → 列宽合计再大也不会把页面撑开，
        # 而自己的高度/宽度仍由 grid 的 nsew + rowconfigure(5, weight=1) 决定。
        tree_wrap.grid_propagate(False)
        self.tree = ttk.Treeview(tree_wrap, height=14, show="headings")
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        # 右侧留 18px 给纵向滚动条，底部留 18px 给横向滚动条
        self.tree.place(x=0, y=0, relwidth=1.0, relheight=1.0, width=-18, height=-18)
        vsb.place(relx=1.0, x=-18, y=0, relheight=1.0, height=-18)
        hsb.place(x=0, rely=1.0, y=-18, relwidth=1.0, width=-18)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        # 鼠标滚轮：默认只上下滚；按住 Shift 滚 = 左右滚（表格类界面的通用习惯）
        self.tree.bind("<Shift-MouseWheel>", self._wheel_h)
        f.rowconfigure(5, weight=1)
        return f

    def _wheel_h(self, event):
        """Shift + 滚轮 → 横向滚动（表格类界面的通用习惯）。

        Windows 上 event.delta 正数=向上滚。绑到哪个 Treeview 就用哪个，
        所以②页客户表、④页结果表、⑥页历史记录表都能共用这一个函数。
        """
        try:
            tv = event.widget
            tv.xview_scroll(-1 * (event.delta // 120), "units")
        except Exception:
            pass
        return "break"

    # 客户表各列的建议宽度（2026-09-28 新增）：内容长的列给宽些，
    # 配合 stretch=False，总宽超出窗口时由横向滚动条接管。
    _COL_W = (
        (("邮箱", "mail", "email", "电子"), 210),
        (("网址", "网站", "url", "website", "官网"), 200),
        (("电话", "手机", "whatsapp", "wechat", "微信"), 150),
        (("公司", "company"), 190),
        (("名字", "客户名", "联系人", "name", "contact"), 130),
        (("国家", "地区", "country"), 90),
        (("地址", "address"), 200),
    )

    def _col_width(self, name: str) -> int:
        """按列名猜一个合适的显示宽度。认不出来就给 110。"""
        n = str(name or "").strip().lower()
        for keys, w in self._COL_W:
            for k in keys:
                if k in n:
                    return w
        return 110

    def do_browse(self):
        p = filedialog.askopenfilename(parent=self,
            title="选择客户表", initialdir=str(Path.home() / "Desktop"),
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("所有文件", "*.*")])
        if not p:
            return
        self.var_file.set(p)
        try:
            names = xlsxlight.sheet_names(p)
            self.cmb_sheet["values"] = names
            self.cmb_sheet.set(names[0])
            self.do_load_sheet()
        except Exception as e:
            self._err("打不开", str(e))

    def do_load_sheet(self):
        p = self.var_file.get()
        if not p:
            return
        try:
            rows = xlsxlight.read_sheet(p, self.cmb_sheet.get() or 0)
        except Exception as e:
            self._err("读取失败", str(e))
            return
        if not rows:
            return
        # 2026-09-29：表头不再固定取第 1 行 —— 2510广交会 / Sheet2 这类分表
        # 第 1 行是标题行，真正的列名在第 2 行，固定取行 0 会让国家列被认成
        # 「国家_原值(备份)」、公司列整个认不出来（用户反馈的「列名和内容对不上」）。
        hr = pick_header_row(rows)
        self.header = [str(h or "").strip() for h in rows[hr]]
        self.rows = rows[hr + 1:]
        if hr:
            self.log(f"[表头] 第 1 行是标题行，已自动改用第 {hr + 1} 行作为列名")
        # 列名与实际内容不符时纠正显示（空列名补齐、国家↔公司 写反的对调、重名消歧）
        self.header, _notes = fix_col_names(
            self.header, self.rows, above=(rows[hr - 1] if hr else None))
        for t in _notes:
            self.log(f"[列名] {t}")
        cols = [str(h) or f"第{i+1}列" for i, h in enumerate(self.header)]
        for c in (self.cmb_country, self.cmb_mail, self.cmb_name, self.cmb_company):
            c["values"] = cols
        self.cmb_country.set(cols[guess_col(self.header, COUNTRY_KEYS)] if guess_col(self.header, COUNTRY_KEYS) >= 0 else "")
        i = guess_col(self.header, MAIL_KEYS)
        self.cmb_mail.set(cols[i] if i >= 0 else "")
        i = guess_col(self.header, NAME_KEYS)
        self.cmb_name.set(cols[i] if i >= 0 else "")
        i = guess_col(self.header, COMPANY_KEYS)
        self.cmb_company.set(cols[i] if i >= 0 else "")
        # 2026-10-09：如果国家列「没认出来」或「认出来的列里几乎没一个国家写法」，
        #   就扫一遍所有列，把国家写法最多的那列认成国家列。
        #   起因：新分表（227 行）列名是「名字/国家/客户等级/邮箱/…」且夹杂
        #   "Paper Straw Group" 这类非国家文本，关键词匹配失败 → MX 定时整批被挡。
        #   这里只填下拉框，不改任何原始数据。
        try:
            self._autofill_country(cols)
        except Exception:
            pass

        self.tree["columns"] = cols
        # 2026-09-28 用户反馈：「右边的列划不过去」。
        # 根源是这里 stretch=True —— 所有列都拉伸去填满窗口，于是列宽被自动压扁、
        # 内容被截断，而且总宽永远不超出容器，横向滚动条就成了摆设。
        # 改成 stretch=False + 按列名给合理宽度：总宽该超出就超出，交给横向滚动条。
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=self._col_width(c), minwidth=48, stretch=False)
        self.tree.delete(*self.tree.get_children())
        for r in rows[1:41]:
            self.tree.insert("", "end", values=r[:len(cols)])
        self.log(f"载入 {len(self.rows)} 行，列：{cols}")
        self._country_check(cols)      # 2026-09-29：顺手体检国家列
        self.refresh_resume()          # 换了表/换了分表，断点提示跟着变
        self._save_ui()                # 记住这次选的表，重开自动填回
        # 2026-10-25：换了分表 → 立刻刷新「模板跟进建议」（③ 页）
        try:
            self.refresh_tpl_advice()
        except Exception:
            pass

    def _autofill_country(self, cols: list):
        """国家列兜底认领：哪一列的国家写法最多，就让它当国家列。

        只在「国家列没认出来」或「认出来的那列内容不像国家」时才动手，
        绝不会覆盖用户已经正确选好的列。顺带把「邮箱列」也补一次。
        """
        def col_vals(i):
            return [r[i] for r in self.rows[:600] if i < len(r)]

        def score(i):
            v = [str(x or "").strip() for x in col_vals(i) if str(x or "").strip()]
            if len(v) < 5:
                return 0.0
            return sum(1 for x in v if re.fullmatch(r"[A-Za-z]{2,3}", x)) / len(v)

        best_i, best_r = -1, 0.0
        for i in range(len(cols)):
            r = score(i)
            if r > best_r:
                best_i, best_r = i, r
        if best_i < 0 or best_r < 0.3:
            return
        cur = []
        try:
            _v = self.cmb_country.get()
            if _v in cols:
                cur = [str(x or "").strip() for x in col_vals(cols.index(_v))
                       if str(x or "").strip()]
        except Exception:
            cur = []
        cur_r = (sum(1 for x in cur if re.fullmatch(r"[A-Za-z]{2,3}", x)) / len(cur)
                 if len(cur) >= 5 else 0.0)
        if cur_r >= best_r:
            return                     # 现在这个已经比它好，不动
        self.cmb_country.set(cols[best_i])
        self.log(f"[国家列] 原来那列不太像国家（像的只占 {int(cur_r * 100)}%），"
                 f"已自动改用「{cols[best_i]}」列（像的占 {int(best_r * 100)}%）。"
                 f"不对的话请在下面手动改回来。")
        # 邮箱列没认出来时一起补
        try:
            if not self.cmb_mail.get():
                for i, c in enumerate(cols):
                    v = [str(x or "") for x in col_vals(i)]
                    if v and sum(1 for x in v if "@" in x) / max(len(v), 1) >= 0.4:
                        self.cmb_mail.set(c)
                        self.log(f"[邮箱列] 已自动认成「{c}」列")
                        break
        except Exception:
            pass
        self._country_check(cols)

    def _country_check(self, cols=None):
        """载入客户表后体检「国家」列：列出脚本认不出来的写法。

        2026-09-29 用户反馈：「国家后面只要跟了其他内容就识别不出来」。
        脚本现在能自动剥后缀（美国广交会客户 → 美国、沙特sa → 沙特阿拉伯），
        但**地区泛称**（欧盟/中东）和**填错列的公司名**仍必须人工改客户表。
        这里在日志里直接列出来，省得用户一封封等到孚盟报错才发现。
        """
        try:
            from core import geo_alias
        except Exception:
            return
        try:
            hdr = [str(h or "").strip() for h in (self.header or [])]
            sel = (self.cmb_country.get() or "").strip()
            ic = (hdr.index(sel) if sel in hdr
                  else (hdr.index("国家") if "国家" in hdr else -1))
            if ic < 0:
                self.log("[国家体检] 没找到「国家」列，跳过")
                return
            vals, bad_amb, bad_unk = {}, [], []
            for r in (self.rows or []):
                if not isinstance(r, (list, tuple)) or ic >= len(r):
                    continue
                v = str(r[ic] or "").strip()
                if v:
                    vals[v] = vals.get(v, 0) + 1
            for v in vals:
                c, note = geo_alias.resolve_country(v)
                if c:
                    continue
                if note == "amb":
                    bad_amb.append(v)
                else:
                    bad_unk.append(v)
            if not (bad_amb or bad_unk):
                self.log(f"[国家体检] ✓ {len(vals)} 个国家写法全部能识别")
                return
            self.log(f"[国家体检] ⚠ {len(vals)} 个国家写法里有 "
                     f"{len(bad_amb) + len(bad_unk)} 种认不出来（这些客户会被跳过/报错）：")
            for v in bad_amb:
                self.log(f"    · 「{v}」是地区泛称 → 请在客户表改成**具体国家名**")
            for v in bad_unk[:15]:
                self.log(f"    · 「{v}」认不出 → 多半是国家列填错了（填成公司名/备注）")
            if len(bad_unk) > 15:
                self.log(f"    · ……还有 {len(bad_unk) - 15} 种，同上")
        except Exception as e:
            self.log(f"[国家体检] 跳过（{e}）")

    # ---------------- ③ 设置 ----------------
    def _tab_settings(self, nb):
        f = ttk.Frame(nb, padding=12)
        self.var_gap = tk.DoubleVar(value=float(self.cfg.interval_min))
        self.var_ws = tk.IntVar(value=self.cfg.work_start)
        self.var_we = tk.IntVar(value=self.cfg.work_end)
        self.var_only = tk.BooleanVar(value=self.cfg.only_work_hours)
        self.var_skip = tk.BooleanVar(value=self.cfg.skip_sent)
        self.var_dry = tk.BooleanVar(value=self.cfg.dry_run)
        self.var_test = tk.StringVar(value="")
        self.var_fast = tk.BooleanVar(value=bool(getattr(self.cfg, "fast_mode", True)))
        self.var_human = tk.BooleanVar(value=bool(getattr(self.cfg, "human_mode", True)))
        self.var_clear_cc = tk.BooleanVar(value=bool(getattr(self.cfg, "clear_cc", False)))
        self.var_verify = tk.StringVar(
            value=f"{int(getattr(self.cfg, 'verify_every', 10) or 0)}"
                  + {"1": "=每位都核对（最慢最稳）", "10": "=每10位核对一次（推荐）",
                     "20": "=每20位核对一次", "0": "=不核对，整轮结束统一核对（最快）"}.get(
                      str(int(getattr(self.cfg, 'verify_every', 10) or 0)), "=每10位核对一次（推荐）"))

        items = [
            ("封间等待（秒；仅关闭「随机间隔」时生效）", self.var_gap),
            ("发送时段开始（点）", self.var_ws), ("发送时段结束（点）", self.var_we),
        ]
        g1, b1 = self._collapsible(f, "发送节奏", open_=True)
        g1.grid(row=1, column=0, columnspan=3, sticky="ew")
        f.columnconfigure(0, weight=1)
        for i, (lab, var) in enumerate(items):
            ttk.Label(b1, text=lab).grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(b1, textvariable=var, width=10).grid(row=i, column=1, sticky="w")
        ttk.Checkbutton(b1, text="只在中国时间的工作时段发送", variable=self.var_only).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Checkbutton(
            b1, text="跳过已经发过的邮箱（**不勾 = 允许重复发**："
                     "同一位客户、同一个模板+主题也照样再发一遍）",
            variable=self.var_skip).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(b1, text="（这一条是总开关：定时发送 / 立即发送 / 补发 / ⑥页重发全都听它）",
                  foreground="#555555").grid(row=13, column=0, columnspan=2, sticky="w")
        # ---- 随机间隔（2026-09-30 防限流）----
        ttk.Checkbutton(
            b1, text="🧍 随机间隔（避免触发平台限流，建议开）：封间 45–110 秒随机 + "
                     "每连发 5–8 封休息 3–7 分钟 + 偶发多停半分钟",
            variable=self.var_human).grid(
            row=14, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(b1, text="（背景：09-29 一轮 114 封、封间中位仅 17 秒、还都带 1.9M 附件 → "
                           "孚盟服务器 Connection reset，发件箱大量「发送失败」。"
                           "开启后不看上面的「封间等待」）",
                  foreground="#555555", wraplength=520, justify="left").grid(
            row=15, column=0, columnspan=2, sticky="w")
        # ---- 清抄送（2026-10-04 防限流）----
        ttk.Checkbutton(
            b1, text="发送前清掉模板自带的抄送（模板若抄送自己，等于每封投 2 个地址、"
                     "投递量翻倍）",
            variable=self.var_clear_cc).grid(
            row=16, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(b1, text="（09-29 那批 246 个失败现场里 193 个带抄送；"
                           "要留底就别勾，默认关）",
                  foreground="#555555", wraplength=520, justify="left").grid(
            row=17, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(b1, text="演练模式（只填写不点发送，用来看效果）", variable=self.var_dry).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=2)

        # ---- 提速（2026-09-22）----
        ttk.Checkbutton(b1, text="⚡ 极速模式：页面一到位就走，不再死等固定秒数（快 40% 以上）",
                        variable=self.var_fast).grid(
            row=6, column=0, columnspan=2, sticky="w", pady=2)
        fr_v = ttk.Frame(b1)
        fr_v.grid(row=7, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(fr_v, text="发件箱核对频率：").pack(side="left")
        ttk.Combobox(fr_v, width=30, state="readonly", textvariable=self.var_verify,
                     values=["1=每位都核对（最慢最稳）", "10=每10位核对一次（推荐）",
                             "20=每20位核对一次", "0=不核对，整轮结束统一核对（最快）"]
                     ).pack(side="left")
        ttk.Label(b1, text="（不管选哪个，整轮跑完都会统一核对一次并给清单）",
                  foreground="#555555").grid(row=8, column=0, columnspan=2, sticky="w")
        self.var_auto = tk.BooleanVar(value=bool(getattr(self.cfg, "auto_restore", True)))
        ttk.Checkbutton(b1, text="🪟 Edge 窗口被最小化时自动恢复（最小化会让鼠标点击落空，建议开）",
                        variable=self.var_auto).grid(
            row=9, column=0, columnspan=2, sticky="w", pady=2)
        # ---- 允许拉前台（2026-10-05，默认关）----
        self.var_raise = tk.BooleanVar(value=bool(getattr(self.cfg, "auto_raise", False)))
        ttk.Checkbutton(
            b1, text="🚨 允许把浏览器拉到前台（不勾 = 绝不打断你）",
            variable=self.var_raise).grid(
            row=18, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(b1, text="（2026-10-05 起选模板改走页面内部方法，不需要前台，"
                           "浏览器被别的窗口挡着也照样发。开着别的窗口干活请保持不勾——"
                           "勾了会出现任务栏图标闪红、自己打开的窗口被盖住/最小化。"
                           "只有排查页面问题才需要临时勾上）",
                  foreground="#555555", wraplength=520, justify="left").grid(
            row=19, column=0, columnspan=2, sticky="w")
        self.var_revive = tk.BooleanVar(value=bool(getattr(self.cfg, "auto_revive", True)))
        ttk.Checkbutton(b1, text="⟳ 页面僵死/白屏时自动刷新恢复（省得手工按 F5，建议开）",
                        variable=self.var_revive).grid(
            row=10, column=0, columnspan=2, sticky="w", pady=2)

        # ---- 新一轮（2026-09-26：广交会第二轮 / 换模板主题重发同一批客户）----
        # 默认【不勾】：勾上的话，同一位客户就算发件箱里已有旧记录也会再发一轮。
        # 这是用户明确要的能力（"我现在发的是第二轮，不要再跳过去"），
        # 但默认是关的，避免手滑重复发送。
        self.var_new_round = tk.BooleanVar(value=bool(getattr(self.cfg, "new_round", False)))
        ttk.Checkbutton(
            b1, text="🆕 本轮是新的一轮（同一批客户换模板/主题再发，不要因为发过就跳过）",
            variable=self.var_new_round).grid(
            row=11, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(b1, text="（勾上后：只有「同模板+同主题」发过的才跳过；"
                           "看之以往各轮发了多少 → 去「⑥ 历史发送记录」）",
                  foreground="#555555").grid(row=12, column=0, columnspan=2, sticky="w")

        g2, b2 = self._collapsible(f, "试发一封", open_=True)
        g2.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(b2, text="试发收件人（可直接粘贴，多个用逗号/分号/换行隔开）").grid(
            row=0, column=0, sticky="w")
        ent = ttk.Entry(b2, textvariable=self.var_test, width=34)
        ent.grid(row=1, column=0, sticky="ew")
        ent.bind("<KeyRelease>", lambda e: self._peek_mail())
        ent.bind("<<Paste>>", lambda e: self.after(50, self._peek_mail))
        ttk.Button(b2, text="试发一封", command=self.do_test_send).grid(row=1, column=1, padx=8)
        self.lab_peek = ttk.Label(b2, text="", foreground="#1a7f37", wraplength=420, justify="left")
        self.lab_peek.grid(row=2, column=0, columnspan=2, sticky="w")

        g3, b3 = self._collapsible(f, "把发送日期写回 Excel", open_=False)
        g3.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(b3, text="结果写回原表的列（如 L；留空=不回写）").grid(row=0, column=0, sticky="w")
        self.var_col = tk.StringVar(value="L")
        ttk.Entry(b3, textvariable=self.var_col, width=6).grid(row=1, column=0, sticky="w")

        # ---- 模板跟进建议（2026-10-25）----
        # 目的：每个分表「上次用了什么模板、下一步该用什么」一目了然，
        #      再给一个「通用模板」标记（任何分表都能发的那种，比如展会邀请）。
        g4, b4 = self._collapsible(
            f, "模板跟进建议（这个分表上次发了什么、下一步该发什么）", open_=True)
        # 2026-10-29：原来放在 row=3（最底部）。但「发送节奏」这组本身就有 15 行控件，
        # ③ 页又没有滚动条 —— 结果整块被挤出窗口外，用户重启后看不到，以为没做。
        # 现在提到最顶部 row=0，一进 ③ 页就能看见。
        g4.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.var_tpl_advice = tk.StringVar(value="在「② 客户表」选一个分表后，这里会显示建议")
        self._tpl_advice: dict = {}
        ttk.Label(b4, textvariable=self.var_tpl_advice, foreground="#1a4f8a",
                  wraplength=580, justify="left").grid(
            row=0, column=0, columnspan=4, sticky="w")
        row_btn = ttk.Frame(b4)
        row_btn.grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 0))
        ttk.Button(row_btn, text="用建议的模板",
                   command=self.do_use_suggested_tpl).pack(side="left")
        ttk.Button(row_btn, text="把当前模板标为通用",
                   command=self.do_mark_universal).pack(side="left", padx=6)
        ttk.Button(row_btn, text="取消通用标记",
                   command=self.do_unmark_universal).pack(side="left")
        ttk.Button(row_btn, text="看模板正文",
                   command=self.do_show_tpl_body).pack(side="left", padx=6)
        ttk.Button(row_btn, text="模板库…",
                   command=self.do_open_library).pack(side="left")
        ttk.Button(row_btn, text="刷新",
                   command=self.refresh_tpl_advice).pack(side="left", padx=6)
        ttk.Label(b4, text="（「通用模板」= 任何分表都能发的模板，比如展会邀请、节前问候；"
                           "标记存在 data/tpl_prefs.json，不动原有配置）",
                  foreground="#555555", wraplength=580, justify="left").grid(
            row=2, column=0, columnspan=4, sticky="w")
        # ③ 页刚建好时先把建议刷出来（分表可能早在启动时就载入了）
        try:
            self.after(120, self.refresh_tpl_advice)
        except Exception:
            pass
        return f

    # ---------------- 模板跟进建议（2026-10-25）----------------
    def refresh_tpl_advice(self):
        """读当前分表的发送历史，算出「上次模板 / 建议下一封」。"""
        if send_stats is None:
            self.var_tpl_advice.set("（缺 core/send_stats.py，读不到发送历史）")
            return
        sheet = ""
        try:
            sheet = str(self.cmb_sheet.get()).strip()
        except Exception:
            pass
        try:
            a = send_stats.sheet_advice(DATA, sheet, load_tpl_library())
        except Exception as e:
            self.var_tpl_advice.set(f"读取失败：{type(e).__name__}: {e}")
            return
        self._tpl_advice = a
        uni = load_tpl_prefs().get("universal", [])
        cur = ""
        try:
            cur = str(self.cmb_tpl.get()).strip()
        except Exception:
            pass
        L = [f"分表：{sheet or '（未选）'} ｜ 累计已发 {a['sends']} 条"]
        if a["last_tpl"]:
            L.append(f"上次模板：{a['last_tpl']}（{a['last_at'] or '时间未知'}）")
        if a["next_tpl"]:
            L.append(f"建议模板：{a['next_tpl']}")
            if a.get("next_subject"):
                n = a.get("subject_used_before", 0)
                L.append(f"建议主题：{a['next_subject']}"
                         + (f"（这个分表还没用过）" if not n else f"（用过 {n} 次，轮换用）"))
        L.append(f"依据：{a['reason'] or '—'}")
        L.append("通用模板：" + ("、".join(uni) if uni else "（未标记）"))
        if cur:
            L.append(f"当前选中：{cur}" + ("   ✔ 已标为通用" if cur in uni else ""))
        self.var_tpl_advice.set("\n".join(L))

    def do_show_tpl_body(self):
        """看模板正文 + 主题，可一键复制（拿去孚盟建模板用）。"""
        info = (self._tpl_advice or {}).get("next_info") or {}
        if not info.get("body"):
            name = ""
            try:
                name = str(self.cmb_tpl.get()).strip()
            except Exception:
                pass
            for t in load_tpl_library().get("templates", []):
                if t.get("name") == name:
                    info = t
                    break
        if not info.get("body"):
            self._info("没有正文",
                       "当前没有推荐模板，也没在库里找到对应正文。\n\n"
                       "可以先点「刷新」，或点「模板库…」看看有哪些模板。")
            return
        win = tk.Toplevel(self)
        win.title(f"模板正文：{info.get('name', '')}")
        win.geometry("700x540")
        ttk.Label(win, text=f"阶段：{info.get('stage_label', '')}　｜　"
                            f"孚盟端模板名必须完全一致：{info.get('name', '')}",
                  foreground="#1a4f8a").pack(anchor="w", padx=10, pady=(10, 4))
        subs = info.get("subject_options") or []
        rec_subj = (self._tpl_advice or {}).get("next_subject") or ""
        if subs:
            ttk.Label(win, text="主题（★ = 本次推荐，其余留着轮换）：").pack(anchor="w", padx=10)
            for i, s in enumerate(subs, 1):
                star = "  ★ 推荐" if s == rec_subj else ""
                ttk.Label(win, text=f"   {i}. {s}{star}",
                          foreground=("#1a7f37" if star else "#000000")).pack(anchor="w", padx=16)
        txt = tk.Text(win, wrap="word", height=17)
        txt.pack(fill="both", expand=True, padx=10, pady=8)
        txt.insert("1.0", info["body"])

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=10, pady=(0, 10))

        def _copy():
            try:
                win.clipboard_clear()
                win.clipboard_append(info["body"])
                self._info("已复制", "正文已复制到剪贴板，可直接粘进孚盟模板。\n"
                                     "开头是 Hi there, —— 发信时脚本会自动换成客户姓名。")
            except Exception as e:
                self._err("复制失败", str(e))

        ttk.Button(bar, text="复制正文", command=_copy).pack(side="left")
        if info.get("note"):
            ttk.Label(bar, text=info["note"], foreground="#666666",
                      wraplength=440, justify="left").pack(side="left", padx=10)

    def do_open_library(self):
        """模板库总览：哪些能用、哪些停用、每个分表已用几次。"""
        lib = load_tpl_library()
        win = tk.Toplevel(self)
        win.title("模板库")
        win.geometry("900x520")
        sheet = ""
        try:
            sheet = str(self.cmb_sheet.get()).strip()
        except Exception:
            pass
        used = {}
        try:
            if send_stats is not None and sheet:
                adv = send_stats.sheet_advice(DATA, sheet, lib)
                used = dict(adv.get("used") or [])
        except Exception:
            pass

        ttk.Label(win, text=f"启用中的模板（当前分表：{sheet or '未选'}）",
                  foreground="#1a4f8a").pack(anchor="w", padx=10, pady=(10, 2))
        cols = ("stage", "name", "univ", "used", "note")
        tv = ttk.Treeview(win, show="headings", columns=cols, height=9)
        for c, t, w in zip(cols, ("阶段", "孚盟模板名", "通用", "本分表已用", "说明"),
                           (130, 240, 56, 90, 300)):
            tv.heading(c, text=t)
            tv.column(c, width=w, minwidth=40, stretch=(c == "note"))
        tv.pack(fill="both", expand=True, padx=10)
        for t in lib.get("templates", []):
            nm = t.get("name", "")
            tv.insert("", "end", values=(
                t.get("stage_label", ""), nm, "是" if t.get("universal") else "",
                used.get(nm, 0), t.get("note", "")))

        ttk.Label(win, text="已停用（不会出现在推荐里）：",
                  foreground="#c0392b").pack(anchor="w", padx=10, pady=(10, 2))
        cols2 = ("name", "why")
        tv2 = ttk.Treeview(win, show="headings", columns=cols2, height=6)
        for c, t, w in zip(cols2, ("模板名", "停用原因"), (260, 560)):
            tv2.heading(c, text=t)
            tv2.column(c, width=w, stretch=(c == "why"))
        tv2.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for r in lib.get("retired_templates", []):
            tv2.insert("", "end", values=(r.get("name", ""), r.get("reason", "")))

    def do_use_suggested_tpl(self):
        """把①页的模板下拉设成建议的那个。"""
        nxt = (self._tpl_advice or {}).get("next_tpl") or ""
        if not nxt:
            self._info("没有建议", "这个分表推不出下一封用什么模板。\n"
                                   "可能历史记录里没有它的发送记录，手动选一个即可。")
            return
        try:
            self.cmb_tpl.set(nxt)
        except Exception as e:
            self._err("设置失败", str(e))
            return
        # 2026-10-29：模板之外，主题也一并填上（用户要的是"主题也给推荐"）
        subj = (self._tpl_advice or {}).get("next_subject") or ""
        if subj:
            try:
                self.cmb_subj.set(subj)
            except Exception:
                pass
        self.refresh_tpl_advice()
        self.log(f"[模板建议] 已选用「{nxt}」"
                 + (f" + 主题「{subj}」" if subj else ""))

    def do_mark_universal(self):
        cur = ""
        try:
            cur = str(self.cmb_tpl.get()).strip()
        except Exception:
            pass
        if not cur:
            self._info("提示", "先在上面的模板下拉里选一个模板")
            return
        pref = load_tpl_prefs()
        uni = [x for x in pref.get("universal", []) if x != cur]
        uni.append(cur)
        pref["universal"] = uni
        save_tpl_prefs(pref)
        self.refresh_tpl_advice()
        self.log(f"[通用模板] 已标记「{cur}」为通用（任何分表可用）")

    def do_unmark_universal(self):
        cur = ""
        try:
            cur = str(self.cmb_tpl.get()).strip()
        except Exception:
            pass
        pref = load_tpl_prefs()
        uni = [x for x in pref.get("universal", []) if x != cur]
        pref["universal"] = uni
        save_tpl_prefs(pref)
        self.refresh_tpl_advice()
        self.log(f"[通用模板] 已取消「{cur}」的通用标记")

    def _apply_settings(self):
        g = float(self.var_gap.get() or 0.0)
        if g < 0:
            g = 0.0
        # 封间等待：min=max。设 0 即「不额外等待」，只靠发送内部必要的孚盟反应时间。
        self.cfg.interval_min = g
        self.cfg.interval_max = g
        # 数字框填空/乱字时不崩，保留原值
        def _int(var, old):
            try:
                return int(float(var.get()))
            except Exception:
                return old
        self.cfg.work_start = _int(self.var_ws, self.cfg.work_start)
        self.cfg.work_end = _int(self.var_we, self.cfg.work_end)
        self.cfg.only_work_hours = bool(self.var_only.get())
        self.cfg.skip_sent = bool(self.var_skip.get())
        # 随机间隔（2026-09-30 防限流）：开关存配置；开着时引擎忽略「封间等待」
        self.cfg.human_mode = bool(self.var_human.get())
        # 定时时间错开（分钟）：0 = 全部同一时刻
        try:
            self.cfg.spread_gap = max(0, int(float(self.var_spread.get() or 0)))
        except Exception:
            pass
        # 2026-10-06：错开方式 same / even / window（业务要求当地 9 点 → 默认 same）
        try:
            _i = int(self.cmb_spread.current())
            self.cfg.spread_mode = self._spread_modes[_i] if 0 <= _i < 3 else "same"
        except Exception:
            self.cfg.spread_mode = str(getattr(self.cfg, "spread_mode", "same") or "same")
        if self.engine:
            try:
                self.engine.cfg.spread_mode = self.cfg.spread_mode
            except Exception:
                pass
        # 发送前清抄送（2026-10-04 防限流）
        try:
            self.cfg.clear_cc = bool(self.var_clear_cc.get())
        except Exception:
            self.cfg.clear_cc = False
        self.cfg.dry_run = bool(self.var_dry.get())
        # 提速两项
        self.cfg.fast_mode = bool(self.var_fast.get())
        try:
            self.cfg.verify_every = int(str(self.var_verify.get() or "10").split("=")[0].strip())
        except Exception:
            pass
        if self.engine:                       # 立刻生效，不用重连
            self.engine.cfg.fast_mode = self.cfg.fast_mode
            self.engine.cfg.verify_every = self.cfg.verify_every
            try:
                if self.engine.w:
                    self.engine.w.s.fast = self.cfg.fast_mode
            except Exception:
                pass
        self._save_cfg()

        self.cfg.auto_restore = bool(getattr(self, "var_auto", None)
                                     and self.var_auto.get())
        self.cfg.auto_revive = bool(getattr(self, "var_revive", None)
                                    and self.var_revive.get())
        self.cfg.new_round = bool(getattr(self, "var_new_round", None)
                                  and self.var_new_round.get())
        # 2026-10-05：默认关。勾上才会把浏览器拽到用户面前。
        self.cfg.auto_raise = bool(getattr(self, "var_raise", None)
                                   and self.var_raise.get())
        if self.engine:                       # 立刻生效，不用重连
            self.engine.cfg.auto_restore = self.cfg.auto_restore
            self.engine.cfg.auto_revive = self.cfg.auto_revive
            try:
                if self.engine.w:
                    self.engine.w.s.auto_restore = self.cfg.auto_restore
                    self.engine.w.s.auto_raise = self.cfg.auto_raise
            except Exception:
                pass
        self._save_cfg()

    def _peek_mail(self):
        """边打字/边粘贴就把识别结果显示出来，免得用户以为"识别不了"。"""
        got = normalize_emails(self.var_test.get())
        if not self.var_test.get().strip():
            self.lab_peek.config(text="", foreground="#666")
        elif got:
            self.lab_peek.config(text=f"已识别 {len(got)} 个：" + "  ".join(got[:3]) +
                                 (" ..." if len(got) > 3 else ""),
                                 foreground="#1a7f37")
        else:
            self.lab_peek.config(text="没识别出邮箱（检查有没有 @、是不是全角＠）",
                                 foreground="#c0392b")

    def do_test_send(self):
        self._apply_settings()
        if not self.engine:
            self._info("提示", "请先连接孚盟")
            return
        raw = self.var_test.get()
        got = normalize_emails(raw)
        if not got:
            self._err(
                "认不出邮箱",
                "您填的内容里没找到邮箱：\n" + repr(raw)[:200] +
                "\n\n可以这样填（多个用逗号/分号/换行隔开）：\nname@example.com")
            return
        self.var_test.set("; ".join(got))
        if len(got) > 1 and not self._ask(
                "识别到多个邮箱",
                f"识别到 {len(got)} 个：\n" + "\n".join(got[:10]) +
                "\n\n要依次发给全部吗？（选「否」就只发第一个）"):
            got = got[:1]
        targets, subj, tpl = got, self.cmb_subj.get().strip(), self.cmb_tpl.get().strip()

        def run():
            for e in targets:
                r = self.engine.send_one(e, subj, tpl)
                self.log("试发结果:", r)
                self.q.put(("info", f"{e}\n结果：{r.get('msg')}"))
        threading.Thread(target=run, daemon=True).start()

    # ---------------- ④ 定时发送 ----------------
    def _tab_sched(self, nb):
        f = ttk.Frame(nb, padding=12)
        self.var_sched_stat = tk.StringVar(value="任务 0（待发 0）")
        # 2026-09-23：待人工补发清单（自动重发解决不了的）
        self.manual = self._manual_load()
        self.var_manual = tk.StringVar(value="")
        top = ttk.Frame(f)
        top.pack(fill="x")
        ttk.Label(top, text="时间（立即发=北京时间；MX定时=收件人当地时间）：").pack(side="left")
        self.var_when = tk.StringVar(
            value=(now_cn() + timedelta(days=1)).strftime("%Y-%m-%d 09:00"))
        ent_when = ttk.Entry(top, textvariable=self.var_when, width=17)
        ent_when.pack(side="left", padx=4)
        try:    # 改时间/改基准就刷新那行换算说明
            ent_when.bind("<KeyRelease>", self._refresh_tz_preview)
        except Exception:
            pass
        ttk.Label(top, text="模式：").pack(side="left")
        self.cmb_mode = ttk.Combobox(top, width=26, state="readonly",
                                     values=["到点立即发（程序到点替您点发送）",
                                             "孚盟MX定时发送（按收件人当地时间）"])
        self.cmb_mode.current(0)
        self.cmb_mode.pack(side="left")
        ttk.Button(top, text="把客户表生成定时任务", command=self.do_add_schedule).pack(
            side="left", padx=8)
        # 2026-09-30：定时时间自动错开——不然 N 位客户挤在同一秒投递，
        # 孚盟服务器会 Connection reset（发件箱一片「发送失败」）。
        fr_sp = ttk.Frame(f)
        fr_sp.pack(fill="x", pady=(2, 0))
        ttk.Label(fr_sp, text="错开方式：").pack(side="left")
        # 2026-10-06：常见需求是「邮件在收件人当地 9 点定时发送」→ 默认改成「同一时刻」。
        #   原来只有「顺序错开」，第 1 位之后每封往后推，119 位里绝大多数都收不到
        #   当地 9 点（第 67 位 13:46、末位漂到次日）→ 实测反馈「这个时间还是错的」。
        self._spread_modes = ["same", "even", "window"]
        self.cmb_spread = ttk.Combobox(
            fr_sp, width=32, state="readonly",
            values=["同一时刻（所有人都写您填的时间）← 推荐",
                    "顺序错开（每封往后推，末位会漂到次日）",
                    "窗口内随机（都落在基准时间起 N 分钟内）"])
        try:
            _i = self._spread_modes.index(str(getattr(self.cfg, "spread_mode", "same")))
        except Exception:
            _i = 0
        self.cmb_spread.current(_i)
        self.cmb_spread.pack(side="left", padx=4)
        self.var_spread = tk.IntVar(value=int(getattr(self.cfg, "spread_gap", 5) or 0))
        ttk.Label(fr_sp, text="间隔/窗口（分钟）：").pack(side="left", padx=(10, 0))
        ttk.Entry(fr_sp, textvariable=self.var_spread, width=5).pack(side="left", padx=4)
        ttk.Label(fr_sp, text="（排队完就可以关电脑，孚盟按点发）",
                  foreground="#1a7f37").pack(side="left")

        # ---- 2026-10-05：时间基准（实测反馈「发送的当地时间和我设置的不一样」）----
        fr_bs = ttk.Frame(f)
        fr_bs.pack(fill="x", pady=(2, 0))
        ttk.Label(fr_bs, text="时间基准：").pack(side="left")
        self.var_basis = tk.StringVar(
            value=("beijing" if str(getattr(self.cfg, "when_basis", "local")) == "beijing"
                   else "local"))
        ttk.Radiobutton(fr_bs, text="北京时间（我填几点，就几点发）",
                        value="beijing", variable=self.var_basis,
                        command=self._refresh_tz_preview).pack(side="left")
        ttk.Radiobutton(fr_bs, text="收件人当地时间（各国都等自己的钟点走到这儿）",
                        value="local", variable=self.var_basis,
                        command=self._refresh_tz_preview).pack(side="left", padx=10)
        self.var_tz_prev = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.var_tz_prev, wraplength=780, justify="left",
                  foreground="#1a7f37").pack(fill="x", pady=(2, 4))
        try:
            self.after(300, self._refresh_tz_preview)
        except Exception:
            pass

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=6)
        self.btn_sched = ttk.Button(bar, text="暂停", command=self.do_sched_toggle)
        self.btn_sched.pack(side="left")
        ttk.Button(bar, text="清空已完成/失败", command=self.do_sched_clean).pack(
            side="left", padx=8)
        ttk.Button(bar, text="导出已发送清单", command=self.do_export_today).pack(
            side="left", padx=4)
        ttk.Button(bar, text="导出本分表结果", command=self.do_export_sheet).pack(
            side="left", padx=4)
        ttk.Button(bar, text="核对发件箱记录", command=self.do_verify_outbox).pack(
            side="left", padx=4)
        ttk.Button(bar, text="同步发送记录回客户表", command=self.do_sync_sheet).pack(
            side="left", padx=4)
        ttk.Label(bar, textvariable=self.var_sched_stat).pack(side="left", padx=8)

        # ---- 待人工补发清单（2026-09-23：自动重发解决不了的，列出来交用户处理）----
        mb = ttk.Frame(f)
        mb.pack(fill="x", pady=(0, 4))
        # 2026-09-29：默认只补「刚刚发完的当前分表」，按钮名也写清楚
        ttk.Button(mb, text="补发本分表失败的", command=self.do_recheck_resend).pack(
            side="left")
        ttk.Button(mb, text="导出待补发清单", command=self.do_export_manual).pack(
            side="left", padx=4)
        ttk.Button(mb, text="清空清单", command=self.do_clear_manual).pack(
            side="left", padx=4)
        ttk.Label(mb, textvariable=self.var_manual, foreground="#B04000").pack(
            side="left", padx=8)
        self.after(100, self.refresh_manual_count)

        # ---- 分批发送（2026-10-09 用户要求）----
        #   「在定时发送设置中，添加一个输入框，用于指定每次发送该分表的数量。
        #     同时需要记录当前发送进度，下次打开窗口能读取上次记录，
        #     从上次未发送完的位置继续，直到全部发完。不填默认全部发送。」
        frb = ttk.Frame(f)
        frb.pack(fill="x", pady=(4, 0))
        ttk.Label(frb, text="本批发送数量：",
                  font=("Microsoft YaHei UI", 9, "bold")).pack(side="left")
        self.var_batch = tk.StringVar(
            value=str((self.ui.get("sess") or {}).get("batch") or ""))
        ent_batch = ttk.Entry(frb, textvariable=self.var_batch, width=6)
        ent_batch.pack(side="left", padx=4)
        ttk.Label(frb, text="封（留空 / 填 0 = 这一分表剩下的全部发完）").pack(side="left")
        try:
            ent_batch.bind("<KeyRelease>", self._refresh_batch_preview)
        except Exception:
            pass
        ttk.Label(frb, text="　你填几次、就点几次「生成定时任务」；"
                            "每次都从上次停下的下一位接着走。",
                  foreground="#1a7f37").pack(side="left", padx=(12, 0))

        # ---- 断点提示：本分表上次跑到第几位（按分表分开记）----
        rbar = ttk.Frame(f)
        rbar.pack(fill="x")
        self.var_resume_tip = tk.StringVar(value="")
        ttk.Label(rbar, textvariable=self.var_resume_tip, foreground="#b45309",
                  font=("Microsoft YaHei UI", 9, "bold")).pack(side="left")
        self.var_resume = tk.BooleanVar(value=False)
        self.cb_resume = ttk.Checkbutton(rbar, text="断点续跑", variable=self.var_resume,
                                         state="disabled")
        self.cb_resume.pack(side="left", padx=8)
        ttk.Button(rbar, text="标为未完成（重新从第 1 位发）",
                   command=self.do_reset_progress).pack(side="left", padx=4)
        self.var_batch_prev = tk.StringVar(value="")
        ttk.Label(rbar, textvariable=self.var_batch_prev, foreground="#1668c4",
                  font=("Microsoft YaHei UI", 9)).pack(side="left", padx=10)

        # ---- 实时进度面板（逐个写入时的主视觉）----
        pan = ttk.LabelFrame(f, text="实时进度", padding=8)
        pan.pack(fill="x", pady=(6, 4))
        # 2026-10-05：暂停 / 继续 —— 用户盯着这个面板看，按钮就该放在这儿。
        # 原来只有顶部那排里的「暂停」两个字会变成「继续」，夹在 6 个按钮中间，
        # 状态变化又只有文字不一样，很容易看不见 → "暂停了不知道点哪儿继续"。
        prow = ttk.Frame(pan)
        prow.pack(fill="x", pady=(0, 6))
        self.btn_pause = tk.Button(
            prow, text="⏸ 暂停发送", command=self.do_sched_toggle,
            bg="#8a8a8a", fg="white", state="disabled",
            font=("Microsoft YaHei UI", 10, "bold"), padx=14, pady=4,
            relief="raised", cursor="hand2")
        self.btn_pause.pack(side="left")
        self.var_pause_tip = tk.StringVar(value="")
        ttk.Label(prow, textvariable=self.var_pause_tip,
                  font=("Microsoft YaHei UI", 9),
                  foreground="#b45309").pack(side="left", padx=10)
        # 2026-09-29：日志框在⑤页，用户在④页操作时看不见「本轮结束」，
        # 所以这里放一条醒目的收尾横幅 —— 一眼就知道这一轮完了、可以换分表了。
        self.var_mx_done = tk.StringVar(value="")
        ttk.Label(pan, textvariable=self.var_mx_done,
                  font=("Microsoft YaHei UI", 10, "bold"),
                  foreground="#0a6c2a", wraplength=880,
                  justify="left").pack(anchor="w", pady=(0, 4))
        self.var_mx_prog = tk.StringVar(value="尚未开始")
        ttk.Label(pan, textvariable=self.var_mx_prog,
                  font=("Microsoft YaHei UI", 12, "bold"),
                  foreground="#0a6c2a").pack(anchor="w")
        self.pb_mx = ttk.Progressbar(pan, mode="determinate", maximum=100)
        self.pb_mx.pack(fill="x", pady=(4, 2))
        self.var_mx_stat = tk.StringVar(value="成功 0 ｜失败 0 ｜跳过 0")
        ttk.Label(pan, textvariable=self.var_mx_stat,
                  font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        self.var_mx_cur = tk.StringVar(value="")
        ttk.Label(pan, textvariable=self.var_mx_cur, font=("Microsoft YaHei UI", 9),
                  foreground="#1668c4").pack(anchor="w", pady=(2, 0))

        # ---- 已发记录管理：备份 / 清空（想重发某一批时用）----
        sbar = ttk.Frame(f)
        sbar.pack(fill="x", pady=(6, 2))
        ttk.Label(sbar, text="已发记录：").pack(side="left")
        ttk.Button(sbar, text="备份一份", command=self.do_sent_backup).pack(side="left", padx=2)
        ttk.Button(sbar, text="清空全部", command=self.do_sent_clear).pack(side="left", padx=2)
        ttk.Button(sbar, text="只清本分表", command=self.do_sent_clear_sheet).pack(side="left", padx=2)
        self.var_sent_n = tk.StringVar(value="已发记录 0 条")
        ttk.Label(sbar, textvariable=self.var_sent_n, foreground="#555555").pack(side="left", padx=8)

        # ---- 本轮逐位结果（替换掉没用的队列任务表）----
        rcols = ("#", "客户（邮箱 · 国家）", "结果")
        # 2026-09-28：同②页 —— 用 place 铺满，避免列宽把页面撑开；超出交给横向滚动条。
        # pack_propagate(False) 关掉后，高度仍由 pack(fill="both", expand=True) 决定。
        mx_wrap = ttk.Frame(f)
        mx_wrap.pack(fill="both", expand=True, pady=(4, 0))
        mx_wrap.pack_propagate(False)
        self.tree_mx = ttk.Treeview(mx_wrap, columns=rcols, show="headings", height=12)
        mx_vsb = ttk.Scrollbar(mx_wrap, orient="vertical", command=self.tree_mx.yview)
        mx_hsb = ttk.Scrollbar(mx_wrap, orient="horizontal", command=self.tree_mx.xview)
        self.tree_mx.place(x=0, y=0, relwidth=1.0, relheight=1.0, width=-18, height=-18)
        mx_vsb.place(relx=1.0, x=-18, y=0, relheight=1.0, height=-18)
        mx_hsb.place(x=0, rely=1.0, y=-18, relwidth=1.0, width=-18)
        self.tree_mx.configure(yscrollcommand=mx_vsb.set, xscrollcommand=mx_hsb.set)
        self.tree_mx.bind("<Shift-MouseWheel>", self._wheel_h)
        w2 = {"#": 40, "客户（邮箱 · 国家）": 260, "结果": 360}
        for c in rcols:
            self.tree_mx.heading(c, text=c)
            # stretch=False：让总宽可以超出容器，横向滚动条才有意义（同②页）
            self.tree_mx.column(c, width=w2.get(c, 120), minwidth=36, stretch=False)
        self.tree_mx.tag_configure("ok", foreground="#1a7f37")
        self.tree_mx.tag_configure("bad", foreground="#c0392b")
        self.tree_mx.tag_configure("skip", foreground="#8a6d3b")

        # 旧队列任务表：MX 已不走队列，保留对象但不再占用界面
        cols = ("状态", "发送时间", "邮箱", "客户名", "国家", "次数", "结果")
        self.tree_sched = ttk.Treeview(f, columns=cols, show="headings", height=1)
        w = {"状态": 62, "发送时间": 118, "邮箱": 150, "客户名": 80,
             "国家": 80, "次数": 42, "结果": 200}
        for c in cols:
            self.tree_sched.heading(c, text=c)
            self.tree_sched.column(c, width=w.get(c, 90), minwidth=40, stretch=True)
        self.tree_sched.tag_configure("fail", foreground="#c0392b")
        self.tree_sched.tag_configure("done", foreground="#1a7f37")
        self.tree_sched.tag_configure("doing", foreground="#1668c4")
        self.var_progress = tk.StringVar(value="")
        self._sched_notified = False
        self.refresh_sched()
        try:
            self._refresh_batch_preview()
        except Exception:
            pass
        return f

    # ---- 分批发送：预览 / 重置（2026-10-09）----
    def _batch_n(self) -> int:
        """用户填的「本批发送数量」。留空 / 0 / 乱填 → 0（= 全部发完）。"""
        try:
            v = str(self.var_batch.get() or "").strip()
            if not v:
                return 0
            n = int(float(v))
            return n if n > 0 else 0
        except Exception:
            return 0

    def _refresh_batch_preview(self, *_):
        """把「本批填 N 封」换算成用户看得懂的一句话（不写盘、不发信）。"""
        if not hasattr(self, "var_batch_prev"):
            return
        try:
            n = self._batch_n()
        except Exception:
            return
        try:
            d = self._prog()
        except Exception:
            d = {}
        pos, total = int(d.get("pos") or 0), int(d.get("total") or 0)
        cur_cnt = 0
        try:
            cur_cnt = len(self._collect_rows())
        except Exception:
            cur_cnt = 0
        try:
            if n <= 0:
                self.var_batch_prev.set(
                    f"→ 本批把这一分表剩下的全部发完"
                    + (f"（约 {max(total - pos, 0)} 位）" if total > pos else ""))
                return
            if total > 0 and pos < total:
                left = total - pos
                self.var_batch_prev.set(
                    f"→ 本批发第 {pos + 1}~{min(pos + n, total)} 位"
                    f"（剩 {left} 位，约需 {max(1, -(-left // n))} 批发完）")
            elif total > 0 and pos >= total:
                self.var_batch_prev.set(
                    f"→ 本分表 {total} 位已发完；想重发请点「标为未完成」")
            elif cur_cnt:
                self.var_batch_prev.set(
                    f"→ 本批发第 1~{min(n, cur_cnt)} 位（这一分表共 {cur_cnt} 位）")
            else:
                self.var_batch_prev.set("→ 请先在「② 客户表」载入分表")
        except Exception:
            pass

    def do_reset_progress(self):
        """把「当前分表」的进度清零，下次从第 1 位重新发。

        只清这一张分表的记录，不动已发记录（sent_history）、不动别的分表。
        """
        p = self.var_file.get() if hasattr(self, "var_file") else ""
        if not p:
            self._info("提示", "请先在「② 客户表」里选择客户表文件")
            return
        sh = str(self.cmb_sheet.get())
        if not self._ask("确认", f"把「{sh}」这张分表的进度清零？\n\n"
                                "只是让下次从第 1 位重新开始排一遍，\n"
                                "已发记录（哪些邮箱发过）不会被清掉。\n\n"
                                "注意：孚盟「发件箱」里已经排好的定时邮件不会撤回，\n"
                                "若③页没勾「跳过已经发过的邮箱」，重排会把它们再发一遍。\n确定？"):
            return
        try:
            k = self._prog_key()
            m = self._prog_map()
            if k in m:
                m.pop(k, None)
            s = self._sess()
            for kk in ("pos", "total", "batches", "finished", "done"):
                s.pop(kk, None)
            self.cfg.batch_done = {kk: v for kk, v in
                                   (getattr(self.cfg, "batch_done", {}) or {}).items()
                                   if kk != k}
            self._save_ui()
            self._save_cfg()
            self.log(f"[分批] 「{sh}」的进度已清零，下次从第 1 位重新开始")
        except Exception as e:
            self._err("清不掉", str(e))
            return
        self.refresh_resume()
        self._refresh_batch_preview()

    # ---- 实时进度面板更新（工作线程 -> 队列 -> 主线程）----
    def _mx_update(self, p: dict):
        if not hasattr(self, "pb_mx"):
            return
        i, total = p.get("i", 0), max(p.get("total", 0), 1)
        pct = int(i * 100 / total)
        self.var_mx_prog.set(f"进度 {i} / {total}（{pct}%）")
        self.pb_mx["value"] = pct
        eta = p.get("eta")
        self.var_mx_stat.set(
            f"成功 {p.get('ok', 0)} ｜失败 {p.get('fail', 0)} ｜跳过 {p.get('skip', 0)}"
            + (f" ｜预计还需 {eta} 分钟" if eta else ""))
        self.var_mx_cur.set(p.get("cur", "") or "")
        # 2026-09-23 加固：逐位结果那一行是"锦上添花"，任何意外都不许影响进度条
        row = p.get("row")
        if row:
            try:
                r = tuple(row) + ("",) * 4          # 补齐，避免长度不够时索引报错
                self.tree_mx.insert("", "end", values=(r[0], r[1], r[3]),
                                    tags=(r[2],))
                self.tree_mx.see("end")
            except Exception:
                pass

    def _sched_tick(self):
        """每 2 秒自动刷一次列表，保证进度实时可见（不依赖回调是否触发）"""
        try:
            if self.winfo_exists():
                self.refresh_sched()
        except Exception:
            return
        self.after(2000, self._sched_tick)

    def refresh_sched(self):
        if not hasattr(self, "tree_sched"):
            return
        tasks = sorted(self.store.all(), key=lambda t: (t.get("run_at", ""), t.get("email", "")))
        sel = self.tree_sched.get_children()
        pos = None
        try:
            if sel and self.tree_sched.yview():
                pos = self.tree_sched.yview()[0]
        except Exception:
            pos = None
        self.tree_sched.delete(*self.tree_sched.get_children())
        n_pend = n_done = n_fail = n_doing = 0
        for t in tasks:
            st = t.get("status", "?")
            tag = ""
            if st == "pending":
                n_pend += 1
                st = "待发" + ("（暂停）" if self.scheduler.paused else "")
            elif st == "sending":
                n_doing += 1
                st, tag = "写入中…", "doing"
            elif st == "done":
                n_done += 1
                st, tag = "已定时", "done"
            elif st == "failed":
                n_fail += 1
                st, tag = "失败", "fail"
            when = t.get("run_at", "")
            if t.get("mode") == "mx":
                when = f"{t.get('mx_local_time', '')}（当地）"
            self.tree_sched.insert("", "end", values=(
                st, when, t.get("email", ""), t.get("name", ""),
                t.get("country", ""), t.get("attempts", 0),
                (t.get("result") or t.get("last_error") or "")[:60]), tags=(tag,))

        total = len(tasks)
        self.var_sched_stat.set(f"任务 {total}（待发 {n_pend}）"
                                + ("｜调度已暂停" if self.scheduler.paused else ""))

        busy = n_pend + n_doing
        if total == 0:
            self.var_progress.set("")
            self._sched_notified = False
        elif busy == 0:
            tail = "，请到孚盟「发件箱」核对定时记录" if n_fail == 0 else "，失败的可点「清空已完成/失败」后重排"
            self.var_progress.set(
                f"✅ 全部处理完毕：共 {total} ｜已定时 {n_done} ｜失败 {n_fail}{tail}")
            if not self._sched_notified:
                self._sched_notified = True
                try:
                    self._info("部署完成",
                                        f"共 {total} 个任务已处理完：\n"
                                        f"已定时（孚盟发件箱可见）{n_done} 个\n"
                                        f"失败 {n_fail} 个\n\n"
                                        "请到孚盟「发件箱」核对定时记录；\n"
                                        "到点后邮件发出，则会出现在「已发件箱」。")
                except Exception:
                    pass
        else:
            self._sched_notified = False
            self.var_progress.set(
                f"⏳ 处理中：共 {total} ｜已定时 {n_done} ｜写入中 {n_doing} "
                f"｜待发 {n_pend} ｜失败 {n_fail}")
        try:
            if pos is not None:
                self.tree_sched.yview_moveto(pos)
        except Exception:
            pass

    def do_add_schedule(self):
        self._apply_settings()
        if not self.cmb_subj.get().strip():
            self._info("提示", "请先在「① 连接孚盟」里选好主题")
            return
        try:
            when = parse_when(self.var_when.get())
        except ValueError as e:
            self._err("时间不对", str(e))
            return
        rows = self._collect_rows()
        if not rows:
            self.log("[定时] 没读到客户邮箱，已中止（请到②页载入客户表并选好邮箱列）")
            self._info("提示", "没读到客户邮箱，请先在「② 客户表」载入并选好邮箱列")
            return
        mode = "mx" if self.cmb_mode.current() == 1 else "send"
        # 防重复总开关（③页「跳过已经发过的邮箱」）：没勾 = 允许重复发。
        # 必须在弹确认框之前抓出来 —— 弹窗文案要用，线程里也要用。
        dedupe = bool(getattr(self.cfg, "skip_sent", True))
        self._dedupe_on = dedupe
        # 新一轮模式也必须在弹窗前抓出来（弹窗文案要用）。
        # 2026-09-26 事故：这里漏了赋值，而下面弹窗文案引用了 new_round，
        # 导致 UnboundLocalError → 点「把客户表生成定时任务」毫无反应。
        new_round = self._round_mode()
        self.log(f"[定时] 生成任务：模式={'MX定时' if mode == 'mx' else '到点立即发'}，"
                 f"客户 {len(rows)} 位，主题「{self.cmb_subj.get()[:36]}」"
                 f"｜防重复={'开（发过就跳过）' if dedupe else '关（允许重复发）'}")
        if mode == "mx":
            # 有国家才进定时队列；个别行缺国家只跳过那一行，不再整批拒绝。
            with_c = [r for r in rows if (r.get("country") or "").strip()]
            n_no = len(rows) - len(with_c)
            if not with_c:
                self.log("[定时] 一行国家都没认出来，已中止（请到②页选好国家列）")
                self._err("缺国家",
                          "MX 定时发送要按「国家地区」选时区，客户表里没有一行认出国家。\n"
                          "请在「② 客户表」的列对应里选好国家列，或改用「到点立即发」。")
                return
            rows = with_c                   # 只保留认得出国家的那些位
            # ---- 断点续跑 + 分批上限（2026-10-09 重写）----
            #   · 起点 start：本分表上次跑到第几位（读本分表自己的进度表）
            #   · 本批上限 batch：用户填的「本批发送数量」，0 = 不限制
            #   两者叠加 = 「从上次停下的下一位起，本批最多发 batch 位」。
            n_all = len(rows)              # 这一分表总共认得出多少位（含国家）
            kprog = self._prog_key()
            dprog = self._prog(self.var_file.get(), str(self.cmb_sheet.get()))
            start = 0
            if getattr(self, "var_resume", None) and self.var_resume.get():
                start = max(0, int(dprog.get("pos") or 0))
                start = min(start, max(n_all - 1, 0))
            if start:
                rows = rows[start:]
                self.log(f"[定时] 断点续跑：跳过前面 {start} 位，本次从第 {start + 1} 位开始")
            batch = self._batch_n()
            if batch:
                self._batch_run = True
            else:
                self._batch_run = False
            if batch and len(rows) > batch:
                self.log(f"[分批] 本批上限 {batch} 位：本分表共 {n_all} 位、"
                         f"从第 {start + 1} 位起共 {len(rows)} 位待发 → "
                         f"本批只发到第 {start + batch} 位，剩下的下次再点一次继续")
                rows = rows[:batch]
            elif batch:
                self.log(f"[分批] 本批上限 {batch} 位，"
                         f"剩下的 {len(rows)} 位本批一次发完（发完即整表完成）")
            # 2026-10-05：时间基准（Tk 变量只在主线程读一次）
            basis = self._apply_when_basis()
            mx_local = parse_when(self.var_when.get()).strftime("%Y/%m/%d %H:%M")
            # ★ 2026-10-04：过去的时间 = 齐发 = 风控，自动推到下一个安全时刻
            mx_local, _past_warn = self._fix_past_when(mx_local)
            if _past_warn:
                self.log("[定时] " + _past_warn.replace("\n", " "))
                try:
                    self.var_when.set(mx_local.replace("/", "-"))
                except Exception:
                    pass
                self.q.put(("toast", (_past_warn, "warn", 22000)))
            if n_no:
                self.log(f"[定时] 提示：{n_no} 位客户缺国家，将被跳过，只定时其余 {len(rows)} 位")
            # ★ 2026-10-09：本分表已经整轮跑完了，再点就是「重发整张表」——
            #   必须让用户明确知道，不能默默又发一遍。
            _fin = bool(dprog.get("finished")) or (
                int(dprog.get("total") or 0) > 0
                and int(dprog.get("pos") or 0) >= int(dprog.get("total") or 0))
            batch_line = ""
            if batch:
                _end = start + len(rows)
                _left = max(n_all - _end, 0)
                batch_line = (f"【分批】本批发第 {start + 1}~{_end} 位（共 {len(rows)} 位）"
                              + (f"，还剩 {_left} 位 → 发完点「否」没戏，要再点一次本按钮接着发"
                                 if _left else "，本批发完这张分表就全发完了 ✓")
                              + "\n")
            if _fin and not start and not batch:
                if not self._ask(
                        "注意：这一分表上轮已发完",
                        f"「{self.cmb_sheet.get()}」上轮已经全部发完了"
                        f"（{dprog.get('total', 0)}/{dprog.get('total', 0)} 位，"
                        f"{dprog.get('at', '')}）。\n\n"
                        "现在再点，等于把整张分表**重发一遍**。\n"
                        "· 想重发 → 点「是」\n"
                        "· 只是想在别的时间再发一轮 → 确认③页「跳过已经发过的邮箱」\n"
                        "  是否勾着（勾着 = 发过的会被跳过，实际不会重复投递）\n"
                        "· 不想重发 → 点「否」\n\n确定要重发这一整张分表？"):
                    self.log(f"[定时] 「{self.cmb_sheet.get()}」已发完，用户点了「否」，未重发")
                    return
            # 2026-10-05：换算预览 —— 让用户点「是」之前就看清各国实际几点发
            tz_prev = self._tz_preview_text(mx_local, rows)
            # ★ 2026-10-06：错开范围也提前算好放进确认框，并且**这次用的就是
            #   真正写进孚盟的那一份**（work() 里复用同一列表，所以预览=实际）。
            self._spread_list = self._spread_local_times(mx_local, max(1, len(rows)))
            sp_txt = self._spread_preview_text(self._spread_list)
            self.log(f"[定时] 等待您在弹窗里点「是」…（共 {len(rows)} 位，"
                     f"{'北京时间' if basis == 'beijing' else '当地时间'} {mx_local}）")
            if sp_txt:
                self.log("[定时] " + sp_txt.replace("\n", " "))
            if tz_prev:
                self.log("[定时] " + tz_prev.replace("\n", " "))
            if not self._ask(
                    "确认", f"共 {len(rows)} 位客户，将逐个写入孚盟「定时发送」：\n"
                    f"{'北京时间' if basis == 'beijing' else '当地时间'} {mx_local}"
                    + ("（按每位客户国家换算成当地时间）" if basis == "beijing"
                       else "（各国按自己的钟点发，换算到北京时间会差 6~12 小时）")
                    + "\n"
                    + (tz_prev or "")
                    + sp_txt
                    + batch_line
                    + (f"断点续跑：从第 {start + 1} 位开始（前面 {start} 位已跳过）\n" if start else "")
                    + f"模板「{self.cmb_tpl.get()}」主题「{self.cmb_subj.get()[:40]}」\n"
                    + (("⚠ 新一轮模式：这同一批客户如果只在**别的主题**下发过，会照发；\n"
                        "只有「同模板+同主题」发过的才跳过。\n") if new_round else "")
                    + (("⚠⚠ 防重复已关闭（③页「跳过已经发过的邮箱」没勾）：\n"
                        "发过的、发件箱里已经排队的，本轮都会**再发一遍**。\n")
                       if not dedupe else "")
                    + self._spread_mode_line()
                    + "点「是」后 Edge 会自动逐个写入（几分钟~半小时+）。\n"
                      "写入期间您可以正常用电脑干别的：脚本不会把浏览器拽到前台，\n"
                      "被别的窗口挡着也照样跑（2026-10-05 起）。\n"
                    "写好的即使关掉本程序，孚盟也会按时发出。确定？"):
                self.log("[定时] 您在弹窗点了「否」，已取消（什么都没排队）")
                return
            self.log("[定时] 已确认，立即开始逐个写入孚盟定时发送（不走队列）…")
            self.log(f"[定时] 去重模式："
                     f"{'已关闭（允许重复发）' if not dedupe else ('新一轮' if new_round else '常规（发过就跳过）')}"
                     f"｜主题「{self.cmb_subj.get()[:30]}」")
            self._round_mode_on = new_round      # 给 _send_one_with_retry 里的核对用
            self._dedupe_on = dedupe
            if getattr(self, "_mx_running", False):
                self._err("正在写入", "上一轮逐个写入还在进行中。\n"
                          "想停掉重来：点「停止」，等日志打出「本轮结束」后再操作。")
                return
            self._mx_stop = False
            self._mx_pause = False
            if self.engine:
                self.engine.resume()          # 清掉上一轮可能的停止/暂停标记
            try:
                self.btn_stop.config(state="normal")   # 停止键永不置灰，随时可点
                self.btn_sched.config(text="暂停")
                self._sync_pause_btn()      # 2026-10-05：进度面板那个大按钮亮起来
            except Exception:
                pass
            self._mx_running = True
            try:                                  # 新一轮：清空上一轮的结果行
                self.tree_mx.delete(*self.tree_mx.get_children())
                self.pb_mx["value"] = 0
                self.var_mx_prog.set("准备开始…")
                self.var_mx_cur.set("")
                self.var_mx_done.set("")     # 新一轮开始，撤掉上一轮的收尾横幅
            except Exception:
                pass
            # 2026-09-29：开新的一轮 → 清掉上一轮的失败现场，
            # 让「自查并重发」按钮只盯**这一轮这一张分表**。
            self._last_round_failed = {}
            self._last_round_sheet = self._cur_sheet()
            self._last_round_at = time.strftime("%H:%M:%S")
            rows2 = list(rows)
            subject = self.cmb_subj.get().strip()
            template = self.cmb_tpl.get().strip()
            # new_round 已在方法开头抓好（弹窗文案要用），这里不再重复抓。

            def work():
                ok = fail = skip = 0
                # 注意：rows2 是外层方法的变量，这里**绝不能**给它重新赋值
                # （2026-09-23 事故：赋值后 Python 把它当成本函数局部变量，
                #  下面读完就 UnboundLocalError → 线程静默猝死 → 界面卡在准备开始）。
                # 本轮真正要跑的名单统一用 todo。
                todo = list(rows2)
                # ---- 防重复第一道闸：开跑前扫一次孚盟「发件箱」，
                #      已有定时记录的客户直接跳过并记入已发记录（绝不再发）----
                #      dedupe=False（③页没勾"跳过已经发过的"）时整道闸跳过：
                #      连发件箱都不扫，省时间，也绝不会有人被挡下来。
                if not dedupe:
                    self.log("[防重复] 已关闭（③页「跳过已经发过的邮箱」没勾）"
                             "→ 本轮不跳过任何人，发过的也照发")
                    pre = {"found": []}
                else:
                    try:
                        pre = self.engine.preflight_scheduled(
                            [(r.get("email") or "") for r in todo],
                            subject=(subject if new_round else ""))
                    except Exception as _pe:
                        pre = {"found": []}
                        self.log(f"[防重复] 预检出错，照常跑：{type(_pe).__name__}: {_pe}")
                if pre.get("found"):
                    pre_set = {x.strip().lower() for x in pre["found"]}
                    for x in pre["found"]:
                        try:
                            self.engine.mark_scheduled(x)
                        except Exception:
                            pass
                    n0 = len(todo)
                    todo = [r for r in todo
                            if (r.get("email") or "").strip().lower() not in pre_set]
                    self.log(f"[防重复] 跳过 {n0 - len(todo)} 位"
                             f"（发件箱已有定时记录；之前显示\"失败\"的其实是成功，本次不重发）")
                    if not todo:
                        self.log("[定时] 本轮客户全部已有定时记录，无需再发。")
                        self.log("[定时] —— 本轮结束 ——（没有需要发的客户，未发送任何人）")
                        self.q.put(("call", self.refresh_sent_count))
                        # 提前返回同样要复位，否则下次点按钮会被「上一轮还在进行中」挡住
                        self._mx_running = False
                        self._mx_stop = False
                        self._mx_pause = False
                        self.q.put(("call", self._mx_finish_ui))
                        return
                sched_ok: list[str] = []      # 本轮定时成功的邮箱（结束统一核对用）
                ve = int(getattr(self.cfg, "verify_every", 1) or 0)
                # ★ 2026-09-30：每位客户的定时时间依次错开，不再所有人挤在同一秒
                #   （同一秒投递 81 封带大附件的邮件 → 孚盟服务器 Connection reset）。
                #   错开后投递全交给孚盟服务器，脚本排队完就可以关电脑。
                #   2026-10-06：改用 _take_spread_list —— 复用确认框里预览过的那份，
                #   保证「弹窗上看到的时间」和「真正写进孚盟的时间」完全一致。
                when_list = self._take_spread_list(mx_local, len(todo))
                if when_list:
                    self.log(f"[定时] 时间安排：第 1 位当地 {when_list[0]} → "
                             f"第 {len(when_list)} 位当地 {when_list[-1]}"
                             f"｜{self._spread_mode_line().strip()}"
                             f"（投递由孚盟服务器按点完成，本轮排队完即可关电脑）")
                t0 = time.time()
                base = start                  # 断点续跑时的绝对位次偏移
                rel = len(todo)               # 本轮要处理的位数
                total = base + rel            # 位次按整张表算，断点才不会错位
                last = 0                       # 已处理完的位数（断点用）
                # ★ 2026-10-09：分批发送 —— 总位数要按**整张分表**算，而不是
                #   本轮 todo 的长度。否则每批都显示「1/100」，用户看不出整表进度。
                g_total = int(dprog.get("total") or 0) or total
                if base == 0 and g_total < total:
                    g_total = total            # 表变长了（换表/加了客户），以本次为准
                g_batches = int(dprog.get("batches") or 0) + (1 if self._batch_run else 0)
                _kprog = kprog
                try:
                    for i, r in enumerate(todo, 1):
                        ai = base + i                 # 客户表里的绝对位次
                        last = i - 1
                        # ★ 2026-10-06 修：原来这里是 `if i % 10 == 0 or i == rel:`
                        #   每 10 位才记一次断点 —— 中途暂停/关程序时 pos 会**落后
                        #   好几位**（真机实测：日志已跑到第 82 位，ui.json 里的 pos
                        #   却停在 79），重启「断点续跑」就会把 80~82 再排一遍 = 重复
                        #   投递。现在改成每位都记，误差最多 1 位。
                        self._prog_mark(_kprog, pos=base + last, total=g_total,
                                        batches=g_batches, finished=False,
                                        at=now_cn().strftime("%Y-%m-%d %H:%M"))
                        self._sess_mark(pos=base + last, total=g_total,
                                        batches=g_batches, finished=False,
                                        at=now_cn().strftime("%Y-%m-%d %H:%M"))
                        self.q.put(("call", self._save_ui))   # 存盘回主线程做
                        if getattr(self, "_mx_stop", False):
                            self.log(f"[定时] 已停止：处理到第 {ai - 1}/{total} 位")
                            break
                        # 中途可暂停：就地等待，点「继续」接着跑
                        while getattr(self, "_mx_pause", False) and not getattr(self, "_mx_stop", False):
                            self.q.put(("status", f"已暂停（第 {ai}/{total} 位）｜成功 {ok} "
                                                  f"｜失败 {fail} ｜跳过 {skip}"))
                            time.sleep(0.5)
                        if getattr(self, "_mx_stop", False):
                            self.log(f"[定时] 已停止：处理到第 {ai - 1}/{total} 位")
                            break
                        email = (r.get("email") or "").strip()
                        country = (r.get("country") or "").strip()
                        if not country:
                            skip += 1
                            self.log(f"[定时] 第 {ai}/{total} 位 {email} 缺国家，跳过（补好国家后可重新生成）")
                            self.q.put(("mxrun", {"i": ai, "total": total, "ok": ok, "fail": fail,
                                                  "skip": skip, "cur": "",
                                                  "row": (ai, f"{email}（无国家）", "skip",
                                                          "⊘ 跳过：客户表国家列为空")}))
                            continue
                        skip_this, why = self._already_done(email, subject, new_round, dedupe)
                        if skip_this:
                            skip += 1
                            self.log(f"[定时] 第 {ai}/{total} 位 {email} {why}，跳过")
                            self.q.put(("mxrun", {"i": ai, "total": total, "ok": ok, "fail": fail,
                                                  "skip": skip, "cur": "",
                                                  "row": (ai, f"{email}（{country}）", "skip",
                                                          f"⊘ 跳过：{why}")}))
                            continue
                        self.q.put(("mxrun", {"i": ai - 1, "total": total, "ok": ok, "fail": fail,
                                              "skip": skip,
                                              "cur": f"正在处理：{email}（{country}）"}))
                        # 核对频率：1=每位；N=每 N 位（第 1 位必核）；0=中途不核（最快）
                        do_verify = True if ve == 1 else (False if ve <= 0 else (i % ve == 1))
                        # 发送这一位（分类重试 + 防重复 + 网络挂起；
                        # 补发失败邮件走的是同一个方法，行为完全一致）
                        when_i = when_list[i - 1] if i <= len(when_list) else mx_local
                        res, cls, sec = self._send_one_with_retry(
                            email, country, subject, template, when_i,
                            do_verify, tag=f"[定时] 第 {ai}/{total} 位",
                            name=(r.get("name") or ""))
                        if res.get("ok"):
                            ok += 1
                            sched_ok.append(email)
                            self.log(f"[定时] ✓ 第 {ai}/{total} 位 {email}（{country}）已定时"
                                     f"（{country} 当地 {when_i}，由孚盟发出）| 用时 {sec}s")
                        else:
                            fail += 1
                            self.log(f"[定时] ✗ 第 {ai}/{total} 位 {email}（{country}）最终失败"
                                     f"| 用时 {sec}s | {res.get('msg', '')[:90]}")
                            # 记进「本轮失败现场」，供「自查并重发」按钮只补这一批
                            cls = cls or {}
                            self._last_round_failed[email.lower()] = {
                                "email": email.lower(), "country": country,
                                "reason": cls.get("reason", "") or res.get("msg", ""),
                                "name": cls.get("name", "未分类"),
                                "action": cls.get("action", ""),
                                "retry": bool(cls.get("retry", True)),
                                "src": f"本轮第 {ai} 位",
                            }
                        self.q.put(("status", f"定时写入中：{ai}/{total} ｜成功 {ok} ｜失败 {fail} ｜跳过 {skip}"))
                        try:
                            eta = round((total - ai) * ((time.time() - t0) / max(i, 1)) / 60, 1)
                        except Exception:
                            eta = None
                        self.q.put(("mxrun", {
                            "i": ai, "total": total, "ok": ok, "fail": fail, "skip": skip,
                            "eta": eta, "cur": f"刚完成：{email}（{country}）",
                            "row": (ai, f"{email}（{country}）",
                                    "ok" if res.get("ok") else "bad",
                                    (f"✓ 已定时（{country} 当地 {mx_local}）｜用时 {sec}s"
                                     if res.get("ok") else
                                     f"✗ 失败｜用时 {sec}s｜{res.get('msg', '')[:60]}"))}))
                    last = rel                      # 整轮跑完，断点记到末尾
                finally:
                    mins = round((time.time() - t0) / 60, 1)
                    # ★ 2026-10-09：本批收尾 —— 算清「整张分表走到哪了、还剩多少」
                    _pos_now = base + last
                    _fin_now = bool(_pos_now >= g_total and g_total > 0)
                    self._prog_mark(_kprog, pos=_pos_now, total=g_total,
                                    batches=g_batches, finished=_fin_now,
                                    at=now_cn().strftime("%Y-%m-%d %H:%M"))
                    if _fin_now:
                        try:
                            bd = dict(getattr(self.cfg, "batch_done", {}) or {})
                            bd[_kprog] = {"total": g_total,
                                          "at": now_cn().strftime("%Y-%m-%d %H:%M")}
                            self.cfg.batch_done = bd
                            self.q.put(("call", self._save_cfg))
                        except Exception:
                            pass
                    _left_now = max(g_total - _pos_now, 0)
                    self._batch_hint = (
                        (f"✔ 本分表全部发完：{_pos_now}/{g_total} 位，共 {g_batches} 批"
                         if _fin_now else
                         f"⏸ 本批结束：整表已走 {_pos_now}/{g_total} 位，"
                         f"还剩 {_left_now} 位 → 再点一次「把客户表生成定时任务」接着发")
                        if self._batch_run else "")
                    self._sess_mark(pos=_pos_now, total=g_total,
                                    batches=g_batches, finished=_fin_now,
                                    at=now_cn().strftime("%Y-%m-%d %H:%M"))
                    self.q.put(("call", self._save_ui))   # 断点落盘
                    self.q.put(("call", self.refresh_resume))  # 断点提示/横幅跟着更新
                    self.q.put(("call", self.refresh_sent_count))
                    # 本轮发过的模板可能第一次出现 → 让⑥页下拉也认得它
                    self.q.put(("call", self._hist_fill_filters))
                    # 整轮结束：只打开发件箱一次，统一核对本轮所有客户（省一大半时间）
                    try:
                        if sched_ok and self.engine:
                            vb = self.engine.verify_batch_scheduled(sched_ok, folder="发件箱")
                            miss = vb.get("missing") or []
                            self.log(f"[定时] 统一核对：本轮定时成功 {len(sched_ok)} 位 → "
                                     f"发件箱核对到 {len(vb.get('found') or [])} 位"
                                     + (f"，未核对到 {len(miss)} 位" if miss else "，全部核对到 ✓"))
                            self.q.put(("status", f"统一核对：{len(vb.get('found') or [])}"
                                                  f"/{len(sched_ok)} 位在发件箱已确认"))
                    except Exception as e:
                        self.log("[定时] 统一核对出错：", e)
                    # 发件箱巡检（2026-09-30）：定时成功 ≠ 投递成功，
                    # 孚盟投递时可能 Connection reset，把「发送失败」的挑出来提醒
                    try:
                        if self.engine and not self._mx_stop:
                            self._outbox_sweep()
                    except Exception:
                        pass
                    # 2026-09-27 用户要求：跑完不再自动输出「发送结果」Excel，
                    # 历史记录只留在助手内部（⑥ 历史发送记录 / data/send_flow.jsonl）。
                    # 需要出表时手动点④页的「导出已发送清单 / 导出本分表结果」。
                    # 本轮跑完：按原表既有格式把发送日期回写到客户表（另存一份到 output/）
                    try:
                        self._writeback_after_round(
                            {str(e).strip().lower() for e in sched_ok} or None)
                    except Exception as e:
                        self.log("[回写] 自动回写出错：", e)
                    self.log(f"[定时] —— 本轮结束 ——成功 {ok}，失败 {fail}，跳过 {skip}，"
                             f"总用时 {mins} 分钟。请到孚盟「发件箱」核对定时记录条数")
                    self.q.put(("info", f"MX 定时写入结束\n成功 {ok}，失败 {fail}，跳过 {skip}，"
                                f"用时 {mins} 分钟\n请到孚盟「发件箱」核对定时记录条数是否与成功数一致"))
                    self.q.put(("status", f"定时写入完成：成功 {ok} ｜失败 {fail} ｜跳过 {skip}｜可以接着发下一个分表"))
                    # ④页顶部横幅：不用切到⑤页看日志也知道这一轮完了
                    self.q.put(("call", lambda: self._mx_show_done(
                        f"✔ 本轮已结束：「{self._last_round_sheet or '?'}」"
                        f"成功 {ok} ｜失败 {fail} ｜跳过 {skip} ｜用时 {mins} 分钟\n"
                        + (self._batch_hint + "\n" if self._batch_hint else "")
                        + f"→ 可直接换分表接着发"
                        + (f"；本分表失败 {len(self._last_round_failed)} 位，"
                           f"点「自查并重发失败邮件」可只补这批"
                           if self._last_round_failed else "；本分表没有失败客户 ✓"))))
                    # ★ 2026-09-29 用户反馈：发完一个分表后，不重启就发不了下一个
                    #   —— 根因就是这里。正常跑完只打了日志，_mx_running 一直留着 True，
                    #   下次点按钮就被「上一轮还在进行中」挡回去，只能重启程序。
                    #   现在无论**跑完 / 点停止 / 中途出错**，都在 finally 里统一复位。
                    self._mx_running = False
                    self._mx_stop = False
                    self._mx_pause = False
                    self.q.put(("call", self._mx_finish_ui))
                    self.log(f"[定时] 本轮现场已复位，可直接换分表接着发"
                             f"（本分表失败 {len(self._last_round_failed)} 位，"
                             f"点「自查并重发失败邮件」可只补这批）")

            def _run():
                # 2026-09-23：后台线程的任何异常都必须可见，否则界面会
                # 静默卡在「准备开始…」，用户根本不知道出了什么事。
                try:
                    work()
                except Exception:
                    import traceback
                    tb = traceback.format_exc()
                    self.log("[致命] 写入线程异常退出：\n" + tb)
                    self.q.put(("msg", "写入线程出错了（详见日志末尾）：\n"
                                + tb.strip().splitlines()[-1]))
                    self._mx_running = False
                    self._mx_stop = False
                    self._mx_pause = False
                    try:
                        self.btn_stop.config(state="normal")
                        self.btn_sched.config(text="暂停")
                    except Exception:
                        pass
                    self.q.put(("call", self._sync_pause_btn))   # 后台线程 → 回主线程
                    self.q.put(("status", "已中断：写入线程异常，请看日志"))

            threading.Thread(target=_run, daemon=True).start()
            self._info("开始", f"从第 1 位开始逐个写入（共 {len(rows2)} 位），\n"
                       "底部日志会实时打出每位客户的结果和耗时；\n"
                       "中途想停：点「停止」立刻生效。")
            return
        if not self._ask(
                "确认", f"共 {len(rows)} 位客户，定时 {when.strftime('%Y-%m-%d %H:%M')}（北京时间）\n"
                        f"模板「{self.cmb_tpl.get()}」主题「{self.cmb_subj.get()[:40]}」\n"
                        "已发过的和重复任务会自动跳过。确定排队？"):
            return
        r = self.scheduler.add_batch(rows, when, self.cmb_subj.get().strip(),
                                     self.cmb_tpl.get().strip(), mode)
        self.refresh_sched()
        self._info("好了", f"已排队 {r['added']} 个任务\n"
                                f"（已发过跳过 {r['sent']}，重复跳过 {r['dup']}）\n"
                                "任务保存在 data/schedule.json，关掉程序重开也会继续按点发送。")

    def _sync_pause_btn(self):
        """按当前状态刷新「暂停 / 继续」按钮（文字 + 颜色 + 可用状态）。

        必须在主线程调（Tk 控件）。后台线程请用
        self.q.put(("call", self._sync_pause_btn))。
        """
        try:
            p = bool(getattr(self, "_mx_pause", False))
            running = bool(getattr(self, "_mx_running", False))
            b = getattr(self, "btn_pause", None)
            if b is not None:
                if p:
                    b.config(text="▶ 继续发送（点这里接着发）",
                             bg="#e08a00", fg="white", state="normal")
                elif running:
                    b.config(text="⏸ 暂停发送", bg="#c0392b", fg="white",
                             state="normal")
                else:
                    b.config(text="⏸ 暂停发送", bg="#8a8a8a", fg="white",
                             state="disabled")
            try:
                self.var_pause_tip.set(
                    "⏸ 已暂停：当前这封写完就停住，点左边的橙色按钮接着跑"
                    if p else "")
            except Exception:
                pass
            try:                       # 顶部那个老按钮保持同步
                self.btn_sched.config(text="继续" if p else "暂停")
            except Exception:
                pass
        except Exception:
            pass

    def do_sched_toggle(self):
        # 同时控制「定时队列」和「直接逐个写入」两条通道
        paused = self.scheduler.toggle_pause()
        self._mx_pause = paused
        if self.engine:
            self.engine.pause(paused)
        self._sync_pause_btn()
        self.log("[定时]", "已暂停：当前这封写完就停住（下面橙色按钮变「继续发送」，"
                 "点了接着跑）" if paused else "已继续")
        self.refresh_sched()

    def do_sched_clean(self):
        n = 0
        for t in self.store.all():
            if t.get("status") in ("done", "failed", "cancelled"):
                self.store.tasks.pop(t["id"], None)
                n += 1
        self.store.save()
        self.log(f"[定时] 清掉了 {n} 个已完成/失败的任务")
        self.refresh_sched()

    # ---- 已发记录管理（备份 / 清空 / 只清本分表）----
    def _sent_file(self) -> Path:
        return APP / "data" / "sent_history.json"

    def _load_sent_dict(self) -> dict:
        p = self._sent_file()
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _reconcile_sent(self):
        """启动对账：把「日志里判定成功、但已发记录里没有」的客户自动补进记录。

        为什么必须有这一步：2026-09-23 发现 15~16 时那轮 55 位客户日志明明成功，
        记录里却没有（被别的实例整批覆盖掉了）。漏记的后果不是数字不好看，而是
        **下次重跑会把他们再发一遍**。现在每次启动自动对一次账，就算记录以后又被
        改坏也能自动补回来。
        """
        try:
            pat = re.compile(
                r"\[(?:定时|补发)\] ✓ 第 \d+/\d+ 位 ([\w.\-+]+@[\w.\-]+\.\w+)")
            logs = sorted((Path(__file__).resolve().parent / "logs").glob("run_*.log"),
                          key=lambda p: p.stat().st_mtime, reverse=True)[:3]
            succ = set()
            for lp in logs:
                try:
                    for line in lp.read_text(encoding="utf-8",
                                             errors="ignore").splitlines():
                        m = pat.search(line)
                        if m:
                            succ.add(m.group(1).strip().lower())
                except Exception:
                    continue
            if not succ:
                return
            d = self._load_sent_dict()
            miss = sorted(e for e in succ if e not in d)
            if not miss:
                return
            stamp = now_cn().strftime("%Y-%m-%d") + "(sched)"
            for e in miss:
                d[e] = stamp
            self._sent_file().write_text(
                json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
            try:
                if self.engine:
                    self.engine.reload_sent()
            except Exception:
                pass
            self.log(f"[对账] 已发记录里缺 {len(miss)} 位（日志显示他们其实已定时成功），"
                     f"已自动补上，避免下次重发。当前记录 {len(d)} 条")
            self.refresh_sent_count()
        except Exception as e:
            self.log(f"[对账] 跳过（{type(e).__name__}: {e}）")

    def refresh_sent_count(self):
        if not hasattr(self, "var_sent_n"):
            return
        try:
            n = len(self._load_sent_dict())
        except Exception:
            n = 0
        self.var_sent_n.set(f"已发记录 {n} 条")

    def _backup_sent(self) -> Path:
        src = self._sent_file()
        dst = APP / "data" / f"sent_history_备份_{now_cn().strftime('%Y%m%d_%H%M%S')}.json"
        if src.exists():
            shutil.copy2(src, dst)
        return dst

    def do_sent_backup(self):
        try:
            dst = self._backup_sent()
            self.log(f"[记录] 已备份已发记录 → {dst}")
            self._info("已备份", f"已发记录已备份一份：\n{dst}")
        except Exception as e:
            self._err("备份失败", str(e))

    def do_sent_clear(self):
        d = self._load_sent_dict()
        if not d:
            self._info("无需清空", "已发记录本来就是空的。")
            return
        if not self._ask("确认清空",
                         f"当前 {len(d)} 条已发记录将被清空（会先自动备份一份）。\n"
                         "清空后再跑，这些客户会重新发一次。确定？"):
            return
        try:
            self._backup_sent()
            self._sent_file().write_text("{}", encoding="utf-8")
            if self.engine:
                self.engine.reload_sent()
            self.refresh_sent_count()
            self.log(f"[记录] 已清空已发记录（原 {len(d)} 条，已备份）")
            self._info("已清空", f"清掉了 {len(d)} 条已发记录。\n现在重跑会重新发一遍。")
        except Exception as e:
            self._err("清空失败", str(e))

    def do_sent_clear_sheet(self):
        if not self.rows:
            self._info("提示", "请先在「② 客户表」载入客户表")
            return
        d = self._load_sent_dict()
        if not d:
            self._info("无需清理", "已发记录本来就是空的。")
            return
        keys = set()
        for r in self._collect_rows():
            for m in normalize_emails(r.get("email") or ""):
                keys.add(m.lower())
        hit = [k for k in d if str(k).lower() in keys]
        if not hit:
            self._info("无需清理", "当前这张分表的客户都不在已发记录里。")
            return
        if not self._ask("确认清理",
                         f"当前分表里有 {len(hit)} 位已在已发记录中，将被移除（先自动备份）。\n"
                         "移除后重跑本分表，这几位会重新发一次。确定？"):
            return
        try:
            self._backup_sent()
            for k in hit:
                d.pop(k, None)
            self._sent_file().write_text(json.dumps(d, ensure_ascii=False, indent=1),
                                         encoding="utf-8")
            if self.engine:
                self.engine.reload_sent()
            self.refresh_sent_count()
            self.log(f"[记录] 已移除本分表的 {len(hit)} 条已发记录（已备份）")
            self._info("已清理", f"移除了 {len(hit)} 条（本分表）。\n现在重跑本分表会重发这几位。")
        except Exception as e:
            self._err("清理失败", str(e))

    # ---- 结果导出 ----
    def do_export_today(self):
        def run():
            try:
                out, n = export_today_list(Path(__file__).resolve().parent / "data")
                self.log(f"[导出] 已发送清单：{out}（{n} 位）")
                self.q.put(("info", f"已导出已发送清单（{n} 位）\n{out}"))
            except Exception as e:
                self.log("[导出] 失败：", e)
                self.q.put(("msg", f"导出失败：{e}"))
        threading.Thread(target=run, daemon=True).start()

    def do_export_sheet(self):
        path = self.var_file.get() if hasattr(self, "var_file") else ""
        if not path:
            self._info("提示", "请先在「② 客户表」里选择客户表文件")
            return

        def run():
            try:
                sheet = self.cmb_sheet.get() if hasattr(self, "cmb_sheet") else 0
                try:
                    sidx = int(sheet)
                except Exception:
                    sidx = 0
                out, n = export_sheet(path, sidx, str(sheet),
                                      Path(__file__).resolve().parent / "data")
                self.log(f"[导出] 分表结果：{out}（已发/已定时 {n} 位）")
                self.q.put(("info", f"已导出本分表结果（已发/已定时 {n} 位）\n{out}"))
            except Exception as e:
                self.log("[导出] 失败：", e)
                self.q.put(("msg", f"导出失败：{e}"))
        threading.Thread(target=run, daemon=True).start()

    # ---- 发送记录回写客户表（2026-09-24） ----
    def _sent_pairs(self, only=None) -> dict:
        """把已发记录整理成 {邮箱: 日期}，供回写用。

        sent_history 里的值是 "2026-09-24" 或 "2026-09-24(sched)"，
        这里统一剥掉 (sched) 后缀再转成 date。
        """
        import datetime as _dt
        pairs = {}
        src = {}
        try:
            if self.engine:
                src = dict(self.engine.sent or {})
        except Exception:
            src = {}
        if not src:                       # 引擎没起来就直接读盘，保证同步不漏
            try:
                src = json.loads(Path(self._sent_file()).read_text(encoding="utf-8"))
            except Exception:
                src = {}
        for e, v in src.items():
            em = str(e or "").strip().lower()
            if not em or "@" not in em:
                continue
            if only is not None and em not in only:
                continue
            s = str(v or "").replace("(sched)", "").strip()[:10]
            try:
                d = _dt.datetime.strptime(s, "%Y-%m-%d").date()
            except Exception:
                d = now_cn().date()
            pairs[em] = d
        return pairs

    @staticmethod
    def _col_letter(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def do_sync_sheet(self):
        """手动触发：把助手里所有发送记录按原表格式回写到客户表。"""
        path = self.var_file.get() if hasattr(self, "var_file") else ""
        if not path:
            self._info("提示", "请先在「② 客户表」里选择客户表文件")
            return
        self.log("[回写] 开始把已发记录同步回客户表…")

        def run():
            try:
                r = self._writeback(path, None)
                self.q.put(("info", self._writeback_msg(r)))
            except Exception as e:
                import traceback
                self.log("[回写] 失败：\n" + traceback.format_exc(limit=3))
                self.q.put(("msg", f"回写失败：{e}"))
        threading.Thread(target=run, daemon=True).start()

    def _writeback(self, path, emails=None) -> dict:
        """真正执行回写（在后台线程里跑）。返回 sheet_writeback 的结果 dict。"""
        pairs = self._sent_pairs(only=emails)
        if not pairs:
            self.log("[回写] 没有可写入的已发记录，跳过")
            return {"ok": False, "msg": "没有已发记录"}
        r = append_send_record(path, pairs)
        return r or {}

    def _writeback_msg(self, r: dict) -> str:
        """把回写结果整理成人话（含新扩了哪些列）。"""
        if not r.get("ok"):
            return f"回写未完成：{r.get('msg','未知原因')}"
        per = r.get("per_sheet") or {}
        msg = (f"已把发送记录回写到客户表\n"
               f"本次写入 {r.get('written',0)} 格｜日期格式 {r.get('stamp','')}\n"
               f"输出：{r.get('out','')}")
        if per:
            msg += "\n\n各分表写入："
            for k, v in per.items():
                msg += f"\n· {k}：{v} 位"
        nc = r.get("new_cols") or {}
        if nc:
            msg += "\n\n⚠ 原表列不够，已按相同格式向右扩展了列："
            for k, cols in nc.items():
                letters = "、".join(self._col_letter(c) for c in sorted(cols))
                msg += f"\n· {k}：新增 {letters} 列（第 {min(cols)}–{max(cols)} 列）"
        else:
            msg += "\n\n没有新增列（原表已有空列可用）。"
        if r.get("unmatched"):
            msg += (f"\n\n表里还有 {r.get('unmatched')} 位客户不在已发记录里"
                    f"（未发过，本次没动）")
        return msg

    def _writeback_after_round(self, emails=None):
        """每轮发送/补发结束后自动调用：回写 + 弹报告（后台线程）。"""
        path = self.var_file.get() if hasattr(self, "var_file") else ""
        if not path:
            self.log("[回写] 未选择客户表，跳过回写")
            return

        def run():
            try:
                r = self._writeback(path, emails)
                if not r.get("ok"):
                    self.log(f"[回写] 跳过：{r.get('msg','')}")
                    return
                self.log(f"[回写] 完成：写入 {r.get('written',0)} 格 → {r.get('out','')}")
                nc = r.get("new_cols") or {}
                if nc:
                    for k, cols in nc.items():
                        letters = "、".join(self._col_letter(c) for c in sorted(cols))
                        self.log(f"[回写] ⚠ 表「{k}」列不够，已向右扩展：{letters}")
                self.q.put(("info", self._writeback_msg(r)))
                self.q.put(("status", f"已回写客户表：{r.get('written',0)} 格"))
            except Exception as e:
                import traceback
                self.log("[回写] 出错：\n" + traceback.format_exc(limit=3))
        threading.Thread(target=run, daemon=True).start()

    # ---------------- 失败自检 / 自动重发支撑（2026-09-23） ----------------
    def _outbox_hit(self, email: str, subject: str = "", for_verify: bool = False) -> bool:
        """去孚盟「发件箱」看这位客户是否已有定时记录。

        这是**唯一可信凭据**：孚盟可能先落库再报错，也可能报错但其实成功。
        命中就顺手记入已发记录，保证绝不再发第二遍。

        2026-09-26：新一轮模式下这里也要「认主题」——
        发件箱里那条如果**不是本轮主题**（是上一轮留下的），就不能算命中，
        否则新一轮会被这条旧记录挡住（用户说的"第二轮全被跳过"就是这么来的）。

        2026-09-26 之二：③页「跳过已经发过的邮箱」没勾（=允许重复发）时，
        这里也**不再把发件箱里的旧记录当成"已经发成功"**——
        否则您想重复发的那些内容，脚本会拿旧记录当免死金牌、一封都不真发。

        2026-09-28 修正（实测反馈：孚盟网页里明明定时成功，脚本记「失败」）：
        上面那条规则有个副作用 —— 关掉防重复后，连"**失败了要复核一下到底成没成**"
        也被一起掐断了。于是孚盟报「邮件发送失败」时，脚本不去发件箱确认，
        直接判失败。但孚盟常见"先落库再报错 / 报错但其实成功"，
        这一判就可能把真成功的记成失败。
        所以加 for_verify 参数：**失败后的真伪复核**永远查发件箱（不受防重复开关影响）；
        而"要不要跳过已发客户"仍听防重复开关。
        """
        if not self.engine or not email:
            return False
        if getattr(self, "_dedupe_on", True) is False and not for_verify:
            return False
        strict = bool(getattr(self, "_round_mode_on", False))
        try:
            v = self.engine.w.verify_sent(email, subject or "", folder="发件箱",
                                          strict_subject=strict)
        except Exception:
            return False
        if v.get("loose") and not v.get("found") and strict:
            self.log(f"（核对）{email} 发件箱里有记录，但不是本轮主题 → 判为上一轮，本轮照发")
        if not v.get("found"):
            return False
        try:
            self.engine.mark_scheduled(email)
        except Exception:
            pass
        return True

    def _nap(self, sec: float) -> bool:
        """可中断的等待：点了停止返回 False，暂停时就地等待。"""
        if sec <= 0:
            return True
        end = time.time() + sec
        while time.time() < end:
            if getattr(self, "_mx_stop", False):
                return False
            while getattr(self, "_mx_pause", False) and not getattr(self, "_mx_stop", False):
                time.sleep(0.4)
            if getattr(self, "_mx_stop", False):
                return False
            time.sleep(0.4)
        return not getattr(self, "_mx_stop", False)

    def _wait_network(self, max_minutes: int = 10) -> bool:
        """网络断了就在这里等恢复：每 15 秒探一次活，恢复返回 True。

        期间客户位不前进、不计失败，用户点停止可立即中断。
        """
        deadline = time.time() + max_minutes * 60
        probes = 0
        while time.time() < deadline:
            if getattr(self, "_mx_stop", False):
                return False
            while getattr(self, "_mx_pause", False) and not getattr(self, "_mx_stop", False):
                time.sleep(0.5)
            probes += 1
            try:
                href = self.engine._probe() if self.engine else ""
                if href and "/login" not in href and "fumamx" in href:
                    # 网络刚恢复时页面常常还是白屏（断网期间接口全挂，
                    # 页面不会自己重拉数据），先刷一次再继续，否则接下来会连着失败
                    try:
                        if (self.engine and self.engine.w
                                and getattr(self.engine.cfg, "auto_revive", True)):
                            self.q.put(("status", "网络已恢复，正在刷新页面…"))
                            self.engine.w.ensure_healthy("网络刚恢复")
                    except Exception:
                        pass
                    self.q.put(("status", "网络已恢复，继续发送"))
                    return True
            except Exception:
                pass
            self.q.put(("status", f"网络中断，等待恢复中…（已等 {probes * 15 // 60} 分"
                                  f"{probes * 15 % 60} 秒，最多 {max_minutes} 分钟）"))
            if not self._nap(15):
                return False
        return False

    def _round_mode(self) -> bool:
        """是不是"新一轮发送"（③ 页那个勾）。"""
        return bool(getattr(self, "var_new_round", None) and self.var_new_round.get())

    def _already_done(self, email: str, subject: str, new_round: bool,
                      dedupe: bool = True) -> tuple[bool, str]:
        """这一位本轮要不要跳过 —— 防重复的第二道闸。

        dedupe=False（③页「跳过已经发过的邮箱」没勾）→ **一律不跳过**，
            这就是用户要的"有些内容我需要重复发的"。
        dedupe=True：
          普通模式：sent_history 里有这个邮箱 → 直接跳过（最保守，等同 2026-09-22 前）。
          新一轮模式：只有当「同一个邮箱 + 同一个主题」在流水里已存在才跳过；
            发过的是别的主题 → 那是上一轮，本轮照样发。
            （旧流水没主题的记录不参与判定，否则换主题后就永远发不出去）
        返回 (是否跳过, 原因文字)。
        """
        e = (email or "").strip().lower()
        if not e or not dedupe:
            return False, ""
        if self.engine and self.engine.already_sent(e):
            if not new_round:
                return True, "已发过/已定时"
            try:
                if send_stats is None:
                    return True, "已发过（缺 send_stats 模块，按保守处理）"
                subs = send_stats.sent_subjects_of(DATA, e)
                cur = re.sub(r"\s+", " ", (subject or "").strip())
                # 同一主题才算"本轮发过"；subject 在流水里只存前 80 字，
                # 所以用 same_subject 双向比，别用字符串直接 in。
                if cur and any(send_stats.same_subject(cur, s) for s in subs):
                    return True, f"本轮（主题「{cur[:24]}」）已发过"
            except Exception:
                pass
            return False, "上一轮发过，本轮照发"
        return False, ""

    def _mx_show_done(self, text: str):
        """在④页顶部打出「上一轮已结束」横幅（主线程执行）。"""
        try:
            self.var_mx_done.set(text)
        except Exception:
            pass

    def _fix_past_when(self, when_str: str) -> tuple[str, str]:
        """定时时间已经过去时，自动推到下一个能安全发送的时刻。

        ⚠ 为什么必须管（2026-10-04）：
          ui.json 一直存着上一次的定时时间，实测残留「2026-09-30 09:00」，
          整整过期 4 天。拿过去的时间提交给孚盟，服务器会**立刻**投递 ——
          66 位客户在同一秒齐发，正是 09-29 那场 Connection reset 的复刻
          （发件箱里一片「发送失败」）。错开排布再好也救不了一个过去的时间。

        规则：
          · 至少要留出 15 分钟给脚本排队（排几十位本身就要十几分钟）；
          · 不够就按「同一时刻」往后推一天，最多推 7 天；
          · 落在 09:00 以前或 17:30 以后的，统一拉回当天 09:00。
        返回 (新时间, 提示语)；没问题时提示语为空串。
        """
        from datetime import datetime as _dt
        try:
            t = _dt.strptime((when_str or "").strip(), "%Y/%m/%d %H:%M")
        except Exception:
            return when_str, ""
        try:
            now = now_cn().replace(tzinfo=None, second=0, microsecond=0)
        except Exception:
            return when_str, ""
        soonest = now + timedelta(minutes=15)
        if t >= soonest:
            return when_str, ""
        cand, day = t, 1
        while cand < soonest and day <= 7:
            cand = (t + timedelta(days=day)).replace(
                hour=t.hour, minute=t.minute, second=0, microsecond=0)
            day += 1
        if cand.hour < 9 or (cand.hour > 17) or (cand.hour == 17 and cand.minute > 30):
            cand = cand.replace(hour=9, minute=0)
        new = cand.strftime("%Y/%m/%d %H:%M")
        return new, (f"⚠ 定时时间 {when_str} 已经是过去的时间，"
                     f"孚盟收到过去时间会立刻投递（= 所有人同时齐发 = 触发风控）。\n"
                     f"已自动改到 {new}，不合适请在④页自己改。")

    def _apply_when_basis(self) -> str:
        """读④页「时间基准」单选并写进 cfg。**必须在主线程调**（Tk 变量）。

        beijing = 填的是北京时间，写入孚盟前按客户国家换算成当地时间；
        local   = 旧行为，填什么就当各客户自己的当地时间。
        """
        basis = "local"
        try:
            basis = str((getattr(self, "var_basis", None) and self.var_basis.get()) or "local")
        except Exception:
            basis = "local"
        try:
            self.cfg.when_basis = basis
            if self.engine:
                self.engine.cfg.when_basis = basis
            # ★ 2026-10-06：选完立刻落盘 + 记进现场。以前只在「关窗口/确定设置」
            #   时写 settings.json，中途重启就丢 → 用户以为选了「北京时间」，
            #   实际跑的还是「当地时间」。
            self.ui["basis"] = basis
            self._save_cfg()
        except Exception:
            pass
        return basis

    def _tz_preview_text(self, when: str, rows) -> str:
        """给确认框/日志用的换算预览文本；local 模式返回说明，beijing 模式返回各国时刻。"""
        basis = str(getattr(self.cfg, "when_basis", "local") or "local")
        if basis != "beijing":
            return (f"⚠ 当前是「当地时间」模式：各国都等到自己当地 {when} 才发，\n"
                    f"换算到北京时间各国差 6~12 小时（美国客户会在北京晚上才发）。\n"
                    f"想让所有人同一时刻收到，请回④页改选「北京时间」。\n")
        try:
            from core.tzmap import preview as _tzp
            cs = [(r.get("国家") or "").strip() for r in (rows or [])]
            lines = _tzp(when, cs, limit=6)
        except Exception:
            lines = []
        if not lines:
            return ""
        return "按北京时间 " + when + " 发出 → " + "；".join(lines) + "\n"

    def _refresh_tz_preview(self, *_):
        """④页：改时间基准/改时间时刷新那行绿色说明。"""
        try:
            from core.tzmap import preview as _tzp
        except Exception:
            return
        try:
            basis = str((getattr(self, "var_basis", None) and self.var_basis.get()) or "local")
            self.cfg.when_basis = basis
            when = (getattr(self, "var_when", None) and self.var_when.get() or "").strip()
            if not when:
                self.var_tz_prev.set("")
                return
            if basis == "beijing":
                cs = []
                try:
                    cs = [(r.get("国家") or "").strip() for r in (getattr(self, "rows", None) or [])]
                except Exception:
                    cs = []
                if not cs:
                    cs = ["美国", "英国", "德国", "法国", "澳大利亚", "日本"]
                self.var_tz_prev.set(
                    "按北京时间 %s 发出 → %s" % (when, "；".join(_tzp(when, cs, limit=6))))
            else:
                self.var_tz_prev.set(
                    "按各国当地 %s 发：每个国家的时钟走到这个点才发，"
                    "换算到北京时间各国差 6~12 小时。想同一时刻收到就改选「北京时间」。" % when)
        except Exception:
            pass

    def _spread_local_times(self, base: str, n: int) -> list[str]:
        """把 n 位客户的定时时间依次错开，返回每人一个「当地时间」字符串。

        为什么必须错开（2026-09-30）：
          原来 n 位客户共用**同一个**当地时间（如当地 09:00）。孚盟在定时那一刻
          要同时投递 81 封带 1.9M 附件的邮件 → 服务器连接被重置（Connection reset），
          发件箱里一片「发送失败」。错开之后每封之间隔几分钟，投递是孚盟服务器
          按点完成的 ——**脚本只负责排队，排完就能关电脑**，正好解决
          "不能一直挂机"的问题。

        规则：
          · 第 1 位用基准时间，之后每位累加一个随机间隔（spread_gap ± 40%）；
          · 超过当地 17:30 就顺延到次日 09:00 继续（跨天自动排）；
          · 早于当地 09:00 的（比如基准填了凌晨）统一推到 09:00。
        """
        from datetime import datetime as _dt
        try:
            t0 = _dt.strptime((base or "").strip(), "%Y/%m/%d %H:%M")
        except Exception:
            return [base] * n
        n = max(1, n)
        mode = str(getattr(self.cfg, "spread_mode", "same") or "same")
        gap = int(getattr(self.cfg, "spread_gap", 5) or 0)

        # ---- ① 同一时刻：所有人写同一个时间（要「当地 9 点」就走这条）----
        #   2026-10-06 加。定时通道的投递由**孚盟服务器**按点完成，脚本只是把
        #   时间写进去；所以"齐发"的并发压力在孚盟那一侧，不在本机。
        if mode == "same" or (mode == "even" and gap <= 0):
            t = self._clamp_when(t0)
            return [t.strftime("%Y/%m/%d %H:%M")] * n

        # ---- ② 窗口内随机：都落在 [基准, 基准+gap] 内，既不齐发也不漂远 ----
        if mode == "window" and gap > 0:
            out: list[str] = []
            for _ in range(n):
                cur = t0 + timedelta(minutes=random.uniform(0, gap))
                out.append(self._clamp_when(cur).strftime("%Y/%m/%d %H:%M"))
            return out

        # ---- ③ 顺序错开（旧行为）：第 1 位用基准，之后每位累加 gap±40% ----
        out = []
        cur = t0
        for i in range(n):
            if i and gap > 0:
                cur = cur + timedelta(minutes=gap * random.uniform(0.6, 1.4))
            # ⚠ 2026-10-06 修：秒**不能**在这里归零。
            #   原来每轮都 replace(second=0, microsecond=0)，等于把每封的零头
            #   （平均 30 秒）丢进垃圾桶，逐封累加 → 整体比设定值快约 10%。
            #   实测 119 位平均间隔 4.49 分钟（应为 5 分钟）。
            #   到分钟只在**输出那一下**截断，绝不动 cur。
            out.append(self._clamp_when(cur).strftime("%Y/%m/%d %H:%M"))
        return out

    @staticmethod
    def _clamp_when(cur):
        """把时间拉回可发送区间：早于 09:00 推到 09:00；超过 17:30 顺延次日 09:00。"""
        if (cur.hour > 17) or (cur.hour == 17 and cur.minute > 30):
            return (cur + timedelta(days=1)).replace(
                hour=9, minute=0, second=0, microsecond=0)
        if cur.hour < 9:
            return cur.replace(hour=9, minute=0, second=0, microsecond=0)
        return cur

    def _spread_preview_text(self, lst: list[str]) -> str:
        """确认框里用的「时间安排」预告。2026-10-06 加。

        实测反馈「我设 09:00，孚盟里却显示 13:46」——根因就是顺序错开：只有第 1 位
        等于用户填的时间，之后每位都往后推。这段提前把结果讲清楚。
        """
        try:
            if not lst:
                return ""
            mode = str(getattr(self.cfg, "spread_mode", "same") or "same")
            gap = int(getattr(self.cfg, "spread_gap", 5) or 0)
            if mode == "same" or (mode == "even" and gap <= 0):
                return (f"⏱ 错开方式：同一时刻 → 这 {len(lst)} 位**全部写 {lst[0]}**，\n"
                        f"   和您在④页填的时间一模一样（外贸常见的「当地 9 点」）。\n"
                        f"   投递由孚盟服务器按点完成，排队完即可关电脑。\n")
            if mode == "window" and gap > 0:
                return (f"⏱ 错开方式：窗口内随机 → {len(lst)} 位落在 "
                        f"{min(lst)} ~ {max(lst)}（{gap} 分钟窗口内）\n")
            first, last = lst[0], lst[-1]
            txt = (f"⏱ 错开方式：顺序错开 → 第 1 位 {first} → 第 {len(lst)} 位 {last}"
                   f"（每封约 {gap} 分钟）\n"
                   f"   注意：**只有第 1 位是您填的时间**，之后每一位都会往后推，"
                   f"孚盟里看到的当地时间因此越来越大。\n")
            if first[:10] != last[:10]:
                txt += ("   ⚠ 最后一位已经排到**第二天**了。要「当地 9 点」的话，"
                        "请把④页错开方式改成「同一时刻」。\n")
            return txt
        except Exception:
            return ""

    def _spread_mode_line(self) -> str:
        """确认框/日志里那行「错开方式」说明（一行版）。"""
        mode = str(getattr(self.cfg, "spread_mode", "same") or "same")
        gap = int(getattr(self.cfg, "spread_gap", 5) or 0)
        if mode == "same" or (mode == "even" and gap <= 0):
            return "错开方式：同一时刻（所有人都写您填的那一个时间）\n"
        if mode == "window":
            return f"错开方式：窗口内随机（都落在基准时间起 {gap} 分钟内）\n"
        return (f"错开方式：顺序错开（第 1 位之后每封 +{gap} 分钟，"
                f"末位会漂到次日）\n")

    def _take_spread_list(self, base: str, n: int) -> list[str]:
        """给 work() 用：优先复用确认框里预览过的那份列表 → 预览=实际。

        去重后 todo 只会变少，所以 pre[:n] 一定够；不够就重新生成。
        """
        pre = list(getattr(self, "_spread_list", None) or [])
        if len(pre) >= n:
            return pre[:n]
        return self._spread_local_times(base, n)

    def _outbox_sweep(self) -> int:
        """整轮结束后扫一遍孚盟「发件箱」（定时任务列表），找出「发送失败」的条目。

        背景（2026-09-30 实测反馈）：**定时提交成功 ≠ 投递成功** —— 孚盟服务器在实际
        投递时可能 Connection reset（发太快 / 附件太大被限制），发件箱里留下
        「发送失败 ①失败重发」，而脚本流水里全是"成功"，完全无感。
        这里每轮结束后扫一次列表（只读列表、不点任何邮件，绝对安全），
        把带失败标记的条目挑出来提醒用户，并落盘 data/outbox_failed.json。
        """
        try:
            w = self.engine.w if self.engine else None
            if w is None:
                return -1
            items = w.scan_sent(folder="发件箱", max_scroll=5) or []
        except Exception as e:
            self.log(f"[发件箱巡检] 跳过（{type(e).__name__}: {e}）")
            return -1
        bad = [t for t in items if ("发送失败" in t or "失败重发" in t)]
        try:
            (DATA / "outbox_failed.json").write_text(
                json.dumps({"checked_at": now_cn().strftime("%Y-%m-%d %H:%M:%S"),
                            "scanned": len(items), "failed": bad},
                           ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:
            pass
        if bad:
            self.log(f"[发件箱巡检] ⚠ 发现 {len(bad)} 条「发送失败」的定时邮件"
                     f"（脚本定时成功、孚盟实际投递时被服务器断开——多发于发送太快/附件太大）")
            for t in bad[:12]:
                self.log("    · " + (t or "").replace("\n", " ")[:90])
            if len(bad) > 12:
                self.log(f"    · ……还有 {len(bad) - 12} 条（完整清单在 data/outbox_failed.json）")
            self.log("[发件箱巡检] 处理：① 孚盟发件箱里逐条点「失败重发」；"
                     "② 保持「随机间隔」开启放慢速度；③ 附件尽量压到 1M 以内")
        else:
            self.log(f"[发件箱巡检] 发件箱最近 {len(items)} 条里没有「发送失败」标记 ✓")
        return len(bad)

    def _mx_finish_ui(self):
        """一轮跑完（正常结束 / 点停止 / 中途出错）后，把界面状态复位。

        必须在**主线程**执行 —— 后台线程不能直接碰 Tk 控件，
        所以调用方一律用 self.q.put(("call", self._mx_finish_ui))。

        2026-09-29：以前只复位 _mx_running，按钮文字、进度条还停在上一轮，
        用户看着像"还没结束"。现在连界面一起复位。
        2026-10-05：顺带把「暂停 / 继续」按钮复位成灰色不可用。
        """
        try:
            self.btn_sched.config(text="暂停")
        except Exception:
            pass
        self._sync_pause_btn()
        try:
            self.btn_stop.config(state="normal")
        except Exception:
            pass
        try:
            self.var_mx_cur.set("")
        except Exception:
            pass

    def _cur_sheet(self) -> str:
        """当前应该记进哪个分表。

        普通发送 = 界面上选的分表；「⑥ 历史发送记录」里点重发时，
        目标分表可能不等于界面上那个，所以用 _resend_sheet 覆盖，
        保证流水记的是**真正发到的那一栏**（统计才不会串）。
        """
        return str(getattr(self, "_resend_sheet", "") or "") \
            or (str(self.cmb_sheet.get()) if hasattr(self, "cmb_sheet") else "")

    def _record_send(self, email: str, template: str, subject: str,
                     ok: bool, note: str = "") -> None:
        """写一条「模板+主题+分表」的发送流水（历史发送记录的数据源）。

        只在这里统一落单点，各条发送路径（定时/补发/试发/重发）行为完全一致，
        统计口径不会出现"有的记了有的没记"。统计失败绝不能绊倒发送。
        """
        if send_stats is None:
            return
        try:
            send_stats.record(DATA, email, tpl=template, subject=subject,
                              sheet=self._cur_sheet(), ok=ok, note=note)
        except Exception:
            pass

    def _send_one_with_retry(self, email: str, country: str, subject: str,
                             template: str, mx_local: str,
                             do_verify: bool = True, tag: str = "",
                             name: str = "") -> tuple:
        """发送一位客户：分类重试 + 防重复核对 + 网络挂起。

        主流程与「补发失败邮件」共用这一个方法，保证两者行为完全一致。
        返回 (res, cls, sec)：结果 dict / 失败分类 dict（成功时为 None）/ 耗时秒
        """
        pre = tag or "[定时]"
        ts = time.time()
        res = {"ok": False, "msg": "未执行"}
        attempt = 0
        net_streak = 0
        cls = None
        while True:
            if getattr(self, "_mx_stop", False):
                break
            attempt += 1
            exc = None
            try:
                res = self.engine.send_scheduled_mx(
                    email, subject, template, mx_local, country, verify=do_verify,
                    name=name)
            except Exception as e:
                exc = e
                res = {"ok": False, "msg": f"{type(e).__name__}: {e}"}
            if res.get("ok"):
                break
            cls = classify(res.get("msg", ""), exc)
            max_try = max(1, int(cls.get("max_try") or 1))

            # 防重复第一原则：无论准备重试还是准备放弃，先查一次发件箱。
            # 2026-09-28：这里传 for_verify=True —— 失败后的真伪复核**不受**
            # 「跳过已经发过的邮箱」开关影响。原因是用户关掉防重复后，
            # 孚盟报错时脚本不再去发件箱确认，会把"其实已成功"的记成失败。
            if self._outbox_hit(email, subject, for_verify=True):
                res = {"ok": True, "msg": "已定时（核对发件箱：已有记录，未重发）"}
                self.log(f"{pre} {email} 核对发件箱：已有定时记录 → 判定成功，不重发")
                break

            # 不可自动恢复（数据问题/服务端拒绝）→ 直接进人工清单
            if not cls.get("retry"):
                self._manual_add(email, country, cls, tries=attempt)
                self.log(f"{pre} {email} 判为「{cls['name']}」，不自动重发 "
                         f"→ 已列入待人工清单：{cls['reason']}")
                break

            # 网络类：连续 2 次就挂起等网络，别白白吃掉后面的客户
            if cls["cat"] == "net":
                net_streak += 1
                if net_streak >= 2:
                    self.log(f"[网络] 连续 {net_streak} 次网络类失败，挂起等待网络恢复"
                             f"（最多 10 分钟），当前客户不跳过：{email}")
                    if self._wait_network(10):
                        self.log("[网络] 孚盟已恢复，继续处理当前客户"
                                 "（网络问题不计入重试次数）")
                        attempt = 0
                        net_streak = 0
                        continue
                    self.log("[网络] 等待超时仍未恢复，放弃这一位 → 列入待人工清单")
                    self._manual_add(email, country, cls, tries=attempt)
                    break

            # 页面类失败：第一次先刷新页面再重试（多半是页面卡住没渲染，干等没用）
            if cls["cat"] == "ui" and attempt == 1 and self.engine is not None \
                    and getattr(self.engine.cfg, "auto_revive", True):
                try:
                    self.log(f"{pre} {email} 页面未就绪，先刷新页面再重试…")
                    self.engine.w.revive(cls.get("reason", "") or "页面未就绪")
                except Exception as e:
                    self.log(f"{pre} （刷新页面失败）{type(e).__name__}: {e}")

            if attempt >= max_try:
                self._manual_add(email, country, cls, tries=attempt)
                self.log(f"{pre} {email} 重试 {attempt} 次仍失败（{cls['name']}）"
                         f"→ 已列入待人工补发清单")
                break

            wait = backoff_for(cls["cat"], attempt)
            self.log(f"{pre} {email} 第 {attempt}/{max_try} 次失败"
                     f"（{cls['name']}：{cls['reason']}），{wait:.0f}s 后重试…")
            if not self._nap(wait):
                break

        # 最后一道核对（仅当循环里没查过，例如被停止打断）
        if not res.get("ok") and self.engine is not None \
                and not getattr(self, "_mx_stop", False) and cls is None:
            try:
                v = self.engine.w.verify_sent(email, subject or "", folder="发件箱")
                if v.get("found"):
                    res = {"ok": True, "msg": "已定时（最终核对：发件箱已有记录）"}
                    try:
                        self.engine.mark_scheduled(email)
                    except Exception:
                        pass
                    self.log(f"{pre} {email} 最终核对：发件箱已有记录 → 判定成功（之前是误报）")
            except Exception:
                pass
        # ---- 落一条发送流水：成功才计数 ----
        # res["msg"] 里带"核对发件箱"的就是"其实上一轮已经定时成功"，也算成功一次。
        self._record_send(email, template, subject, bool(res.get("ok")),
                          note=(pre or "").strip("[]") or "发送")
        return res, cls, round(time.time() - ts, 1)

    # ---------------- 待人工补发清单 ----------------
    def _manual_path(self) -> Path:
        return Path(__file__).resolve().parent / "data" / "manual_retry.json"

    def _manual_load(self) -> list:
        p = self._manual_path()
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return []
        return []

    def _manual_save(self) -> None:
        try:
            p = self._manual_path()
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.manual, ensure_ascii=False, indent=1),
                         encoding="utf-8")
        except Exception:
            pass

    def _manual_add(self, email: str, country: str, cls: dict, tries: int = 1) -> None:
        """把"自动重发解决不了"的客户记进清单（按邮箱去重，保留最新）。"""
        try:
            e = (email or "").strip().lower()
            if not e:
                return
            try:
                sheet = str(self.cmb_sheet.get())
            except Exception:
                sheet = ""
            item = {"email": e, "country": country or "", "cat": cls.get("cat", "other"),
                    "name": cls.get("name", "未知"), "reason": cls.get("reason", ""),
                    "action": cls.get("action", ""), "tries": tries,
                    "retry": bool(cls.get("retry", True)),
                    "at": now_cn().strftime("%Y-%m-%d %H:%M"), "sheet": sheet}
            self.manual = [x for x in self.manual
                           if (x.get("email") or "").lower() != e]
            self.manual.append(item)
            self._manual_save()
            self.q.put(("call", self.refresh_manual_count))
        except Exception:
            pass

    def refresh_manual_count(self):
        try:
            n = len(self.manual)
            cats = {}
            for x in self.manual:
                cats[x.get("name", "?")] = cats.get(x.get("name", "?"), 0) + 1
            detail = "、".join(f"{k} {v}" for k, v in sorted(cats.items(), key=lambda kv: -kv[1]))
            self.var_manual.set(f"待人工补发 {n} 位" + (f"（{detail}）" if detail else ""))
        except Exception:
            pass

    def do_export_manual(self):
        """导出待人工补发清单（xlsx；没装 openpyxl 时自动改存 csv，保证拿得到）。"""
        if not self.manual:
            self._info("提示", "目前没有需要人工补发的客户 ✓")
            return
        outdir = Path(RESULT_DIR)
        outdir.mkdir(parents=True, exist_ok=True)
        stamp = now_cn().strftime("%Y%m%d_%H%M")
        head = ["邮箱", "国家", "失败类别", "具体原因", "已尝试次数",
                "建议处理动作", "发生时间", "分表"]
        rowsx = [[x.get("email", ""), x.get("country", ""), x.get("name", ""),
                  x.get("reason", ""), x.get("tries", 1), x.get("action", ""),
                  x.get("at", ""), x.get("sheet", "")] for x in self.manual]
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
            wb = Workbook()
            ws = wb.active
            ws.title = "待人工补发"
            ws.append(head)
            for r in rowsx:
                ws.append(r)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="C00000")
            for col in "ABCDEFGH":
                ws.column_dimensions[col].width = 30
            ws.column_dimensions["D"].width = 46
            ws.column_dimensions["F"].width = 46
            fp = outdir / f"待人工补发清单_{stamp}.xlsx"
            wb.save(fp)
        except Exception:
            fp = outdir / f"待人工补发清单_{stamp}.csv"
            with open(fp, "w", encoding="utf-8-sig", newline="") as f:
                csv.writer(f).writerow(head)
                csv.writer(f).writerows(rowsx)
        self.log(f"[导出] 待人工补发清单：{fp}（{len(rowsx)} 位）")
        self._info("已导出", f"待人工补发清单共 {len(rowsx)} 位：\n{fp}")

    def do_clear_manual(self):
        if not self.manual:
            return
        if not self._ask("确认", f"清空待人工补发清单（{len(self.manual)} 位）？\n"
                                 "建议先导出一份留底。"):
            return
        self.manual = []
        self._manual_save()
        self.refresh_manual_count()
        self.log("[清单] 待人工补发清单已清空")

    # ---------------- 一键自查并重发失败邮件（2026-09-23） ----------------
    def _collect_failed(self) -> list:
        """汇总"发失败过"的客户：待人工清单 + 最近日志里的「最终失败」。

        返回 [{email, country, reason, name, src}]，按邮箱去重（保留最新一条）。
        """
        out: dict[str, dict] = {}

        # ① 程序自动记的待人工清单
        for x in self.manual:
            e = (x.get("email") or "").strip().lower()
            if e and "@" in e:
                out[e] = {"email": e, "country": x.get("country", ""),
                          "reason": x.get("reason", ""),
                          "name": x.get("name", "未分类"),
                          "action": x.get("action", ""),
                          "retry": bool(x.get("retry", True)),
                          "src": "待人工清单"}

        # ② 最近 2 天的运行日志里解析「最终失败」
        pat = re.compile(
            r"^(\d{2}:\d{2}:\d{2})\s+\[定时\] ✗ 第 (\d+)/\d+ 位 "
            r"([\w\.\-\+]+@[\w\.\-]+\.\w+)（(.+?)）最终失败\|\s*用时 [\d\.]+s \|\s*(.+)$")
        # 只看最近那一天的日志：补发只针对上一轮，不翻旧账
        # （更早的失败会一直保存在 data/manual_retry.json 里，不会丢）
        logs = sorted((Path(__file__).resolve().parent / "logs").glob("run_*.log"),
                      key=lambda p: p.stat().st_mtime, reverse=True)[:1]
        for lp in logs:
            day = lp.stem.replace("run_", "")
            try:
                for line in lp.read_text(encoding="utf-8", errors="ignore").splitlines():
                    m = pat.match(line.strip())
                    if not m:
                        continue
                    t, idx, email, country, msg = m.groups()
                    e = email.strip().lower()
                    cls = classify(msg)
                    out[e] = {"email": e, "country": country,
                              "reason": cls.get("reason", ""),
                              "name": cls.get("name", "未分类"),
                              "action": cls.get("action", ""),
                              "retry": bool(cls.get("retry", True)),
                              "src": f"{day} {t} 第{idx}位"}
            except Exception:
                continue

        # ③ 客户表里能查到国家的，以客户表为准（防止日志里记错）
        #    注意 self.rows 是**二维列表**（原始行），不是 dict，只能按列索引取
        by_mail = {}
        try:
            hdr = [str(h or "").strip() for h in (getattr(self, "header", []) or [])]
            cm = (self.cmb_mail.get() or "").strip() if hasattr(self, "cmb_mail") else ""
            cc = (self.cmb_country.get() or "").strip() if hasattr(self, "cmb_country") else ""
            im = (hdr.index(cm) if cm in hdr
                  else (hdr.index("邮箱") if "邮箱" in hdr else -1))
            ic = (hdr.index(cc) if cc in hdr
                  else (hdr.index("国家") if "国家" in hdr else -1))
            if im >= 0:
                for r in (getattr(self, "rows", []) or []):
                    if not isinstance(r, (list, tuple)) or im >= len(r):
                        continue
                    em = str(r[im] or "").strip().lower()
                    if em and "@" in em:
                        by_mail[em] = str(r[ic] or "").strip() if 0 <= ic < len(r) else ""
        except Exception:
            pass
        for e, item in out.items():
            if by_mail.get(e):
                item["country"] = by_mail[e]
        return list(out.values())

    def _ask_resend_scope(self, cur_failed: list):
        """问「补发范围」：返回 "sheet"（只补刚发完的分表）/ "all"（全部）/ None（取消）。

        2026-09-29 用户要求：默认只补**刚刚发完的那一张分表**，
        别一按就把待人工清单里其他分表的旧失败一起翻出来。
        """
        if not cur_failed:
            # 本轮没记录到失败（刚重启过 / 上一轮全成功）→ 退回老逻辑，但先说清楚
            if not self._ask("补发范围",
                             "本轮（本次启动以来）没有记录到失败客户。\n"
                             "可能是刚重启过，或上一轮全部成功。\n\n"
                             "改从「待人工清单 + 最近日志」里找失败客户来补发吗？"):
                return None
            return "all"
        sheet = str(getattr(self, "_last_round_sheet", "") or "(未记录分表)")
        win = tk.Toplevel(self)
        win.title("补发范围")
        win.transient(self)
        win.resizable(False, False)
        var = tk.StringVar(value="sheet")     # 默认：只补刚发完的分表
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="要补发哪些失败客户？").pack(anchor="w")
        ttk.Radiobutton(
            frm, value="sheet", variable=var,
            text=f"只补「{sheet}」本轮失败的 {len(cur_failed)} 位　← 推荐"
        ).pack(anchor="w", pady=(10, 2))
        ttk.Radiobutton(
            frm, value="all", variable=var,
            text="补全部（待人工清单 + 最近日志，含其他分表）"
        ).pack(anchor="w", pady=2)
        ans = {"v": None}

        def _ok():
            ans["v"] = var.get()
            win.destroy()

        def _no():
            ans["v"] = None
            win.destroy()

        bb = ttk.Frame(frm)
        bb.pack(fill="x", pady=(14, 0))
        ttk.Button(bb, text="取消", command=_no).pack(side="right", padx=4)
        ttk.Button(bb, text="确定", command=_ok).pack(side="right")
        win.grab_set()
        self.wait_window(win)
        return ans["v"]

    def do_recheck_resend(self):
        """一键：先自查失败原因 → 发件箱核对 → 只补发确实没发出去的。

        2026-09-29：默认只补**刚刚发完的当前分表**（用户要求），
        想补其他分表的旧失败时，在范围弹窗里选「补全部」。
        """
        if not self.engine:
            self._info("提示", "请先在「① 连接孚盟」里连接浏览器")
            return
        # ---- 先定范围：本轮刚发完的分表 / 全部 ----
        cur_failed = list((getattr(self, "_last_round_failed", {}) or {}).values())
        scope = self._ask_resend_scope(cur_failed)
        if scope is None:
            self.log("[补发] 已取消")
            return
        if scope == "sheet":
            cands = cur_failed
            src_note = f"刚刚发完的分表「{self._last_round_sheet or '(未记录)'}」"
            # 流水要记到真正发到的那一栏，别记成界面上后来选的分表
            self._resend_sheet = str(self._last_round_sheet or "")
        else:
            cands = self._collect_failed()
            src_note = "全部待人工清单 + 最近日志"
        if not cands:
            self._info("没有失败记录 ✓",
                       f"补发范围：{src_note}\n"
                       "待人工清单是空的，最近日志里也没有「最终失败」记录。\n"
                       "不需要补发。")
            return

        subject = (self.cmb_subj.get() if hasattr(self, "cmb_subj") else "").strip()
        template = (self.cmb_tpl.get() if hasattr(self, "cmb_tpl") else "").strip()
        try:
            mx_local = parse_when(self.var_when.get()).strftime("%Y/%m/%d %H:%M")
        except Exception:
            mx_local = ""
        self._apply_when_basis()          # 2026-10-05：补发同样跟随④页的时间基准
        # 过去的时间会让孚盟立刻投递（= 齐发 = 风控），自动推到下一个安全时刻
        if mx_local:
            mx_local, _past_warn = self._fix_past_when(mx_local)
            if _past_warn:
                self.log("[补发] " + _past_warn.replace("\n", " "))
                try:
                    self.var_when.set(mx_local.replace("/", "-"))
                except Exception:
                    pass
                self._toast(_past_warn, "warn", 22000)
        if not template or not subject or not mx_local:
            self._info("还缺设置", "请先填好：模板、主题、发送时间（在「③ 发送设置」页），\n"
                                   "补发要用和原来一样的模板与时间。")
            return

        # 分成两类：能自动补发的（网络/页面类）／注定补不成的（数据问题、被拒收）
        auto = [c for c in cands if c.get("retry", True)]
        stuck = [c for c in cands if not c.get("retry", True)]

        cats = {}
        for c in cands:
            cats[c["name"]] = cats.get(c["name"], 0) + 1
        detail = "\n".join(f"  · {k}：{v} 位" for k, v in
                           sorted(cats.items(), key=lambda kv: -kv[1]))
        block = ""
        if stuck:
            block = ("\n\n以下 {} 位属「自动重发也没用」，本次跳过、需您先修数据：\n".format(len(stuck))
                     + "\n".join(f"  · {c['email']}（{c['country'] or '缺国家'}）"
                                 f"{c['name']}：{c.get('reason', '')[:40]}"
                                 for c in stuck[:6])
                     + ("\n  …" if len(stuck) > 6 else ""))
        if not auto:
            self._info("没有可自动补发的",
                       f"找到 {len(cands)} 位失败客户，但全部属于「自动重发也没用」的类型。\n"
                       f"{block}\n\n修好数据后（比如把「欧盟」改成具体国家名）再点这个按钮。")
            return
        if not self._ask("自查并重发失败邮件",
                         f"补发范围：{src_note}\n"
                         f"共找到 {len(cands)} 位发送失败过的客户：\n{detail}\n"
                         f"其中可自动补发 {len(auto)} 位" +
                         (f"，需先修数据 {len(stuck)} 位" if stuck else "") + f"\n{block}\n\n"
                         f"将按以下顺序处理：\n"
                         f"① 先去孚盟「发件箱」自查——其实已定时成功的直接判定成功、绝不重发；\n"
                         f"② 剩下的逐个补发（模板「{template}」、当地时间 {mx_local}）；\n"
                         f"③ 仍失败的会写明原因，留在「待人工补发清单」里。\n\n"
                         f"确定开始？"):
            self.log("[补发] 已取消")
            return

        if getattr(self, "_mx_running", False):
            self._err("正在发送", "上一轮还在进行中，请先「停止」再补发。")
            return

        self._mx_stop = False
        self._mx_pause = False
        self._round_mode_on = False      # 补发永远是同一轮，去重回到最保守
        self._dedupe_on = bool(getattr(self.cfg, "skip_sent", True))
        self.engine.resume()
        self._mx_running = True
        try:
            self.btn_stop.config(state="normal")
            self.btn_sched.config(text="暂停")
            self._sync_pause_btn()          # 2026-10-05
            self.tree_mx.delete(*self.tree_mx.get_children())
            self.pb_mx["value"] = 0
            self.var_mx_done.set("正在补发…")
        except Exception:
            pass

        cands = list(auto)           # 只补发可自动恢复的（线程内只读，不重赋值外层变量）
        blocked = list(stuck)        # 需要先修数据的，收尾时一并报告

        def work():
            ok_n = skip_n = bad_n = 0
            bad_list = []
            ok_emails = []               # 本轮确认已发的（含发件箱里已排队的），用于回写
            t0 = time.time()
            self.log(f"[补发] —— 开始自查并重发，共 {len(cands)} 位 ——")
            # ① 自查：发件箱里已有的，判成功、不重发
            todo = list(cands)
            if not getattr(self, "_dedupe_on", True):
                # 防重复关着 → 补发也不自查，清单里的人一律真发一遍
                self.log("[补发] 防重复已关闭（③页没勾「跳过已经发过的邮箱」）"
                         f"→ 不自查发件箱，{len(todo)} 位全部真发一遍")
                pre = {"found": []}
            else:
                try:
                    pre = self.engine.preflight_scheduled([c["email"] for c in todo])
                    found = {str(x).strip().lower() for x in (pre.get("found") or [])}
                    if found:
                        for e in found:
                            try:
                                self.engine.mark_scheduled(e)
                            except Exception:
                                pass
                        for c in todo:
                            if c["email"].lower() in found:
                                self._manual_drop(c["email"])
                        already = [c for c in todo if c["email"].lower() in found]
                        todo = [c for c in todo if c["email"].lower() not in found]
                        self.log(f"[补发] 自查结果：{len(already)} 位其实已在发件箱排队"
                                 f"（判成功、不重发）；真正需要补发 {len(todo)} 位")
                        for c in already:
                            self.log(f"[补发]   ✓ 已有定时记录，跳过：{c['email']}")
                            ok_emails.append(c["email"])
                    else:
                        self.log(f"[补发] 自查结果：发件箱里没有这批记录，"
                                 f"全部 {len(todo)} 位都需要补发")
                except Exception as e:
                    self.log(f"[补发] 自查发件箱出错：{type(e).__name__}: {e}"
                             f"（继续按原名单补发，发送中仍会逐位核对）")

            # ② 逐个补发
            total = len(todo)
            # 补发也按「每人错开几分钟」排时间，避免同一秒齐发被服务器掐断
            when_list = self._spread_local_times(mx_local, len(todo))
            for i, c in enumerate(todo, 1):
                if getattr(self, "_mx_stop", False):
                    self.log(f"[补发] 已停止：处理到第 {i - 1}/{total} 位")
                    break
                while getattr(self, "_mx_pause", False) and not getattr(self, "_mx_stop", False):
                    self.q.put(("status", f"补发已暂停（第 {i}/{total} 位）"))
                    time.sleep(0.5)
                if getattr(self, "_mx_stop", False):
                    break
                email, country = c["email"], c["country"]
                when_i = when_list[i - 1] if i <= len(when_list) else mx_local
                self.q.put(("status", f"补发中：{i}/{total} ｜成功 {ok_n} ｜失败 {bad_n}"))
                self.q.put(("mxrun", {"i": i, "total": total, "ok": ok_n, "fail": bad_n,
                                      "skip": skip_n,
                                      "cur": f"正在补发：{email}（{country}）"}))
                res, cls, sec = self._send_one_with_retry(
                    email, country, subject, template, when_i, True,
                    tag=f"[补发] 第 {i}/{total} 位",
                    name=(c.get("name") or ""))
                if res.get("ok"):
                    ok_n += 1
                    ok_emails.append(email)
                    self._manual_drop(email)          # 补发成功 → 从人工清单移除
                    # 本轮失败现场里也划掉，下次点补发不会再翻出他
                    self._last_round_failed.pop(email.lower(), None)
                    self.log(f"[补发] ✓ 第 {i}/{total} 位 {email}（{country}）补发成功"
                             f"｜用时 {sec}s")
                else:
                    bad_n += 1
                    nm = (cls or {}).get("name", "未分类")
                    why = (cls or {}).get("reason", res.get("msg", ""))[:70]
                    act = (cls or {}).get("action", "")
                    bad_list.append((email, country, nm, why, act))
                    # 仍失败的留在现场里，方便再点一次继续补
                    self._last_round_failed[email.lower()] = {
                        "email": email.lower(), "country": country,
                        "reason": why, "name": nm, "action": act,
                        "retry": bool((cls or {}).get("retry", True)),
                        "src": f"补发第 {i} 位",
                    }
                    self.log(f"[补发] ✗ 第 {i}/{total} 位 {email}（{country}）仍失败"
                             f"｜{nm}｜{why}")
                self.q.put(("mxrun", {"i": i, "total": total, "ok": ok_n, "fail": bad_n,
                                      "skip": skip_n,
                                      "cur": f"刚完成：{email}（{country}）",
                                      "row": (i, f"{email}（{country}）",
                                              "ok" if res.get("ok") else "bad",
                                              (f"✓ 补发成功｜用时 {sec}s" if res.get("ok")
                                               else f"✗ 仍失败｜{res.get('msg', '')[:50]}"))}))

            # ③ 收尾报告
            mins = round((time.time() - t0) / 60, 1)
            self.q.put(("call", self._save_ui))
            self.q.put(("call", self.refresh_sent_count))
            self.q.put(("call", self.refresh_manual_count))
            left = len(self.manual)
            msg = (f"补发完成\n"
                   f"补发成功 {ok_n} 位｜仍失败 {bad_n} 位｜总用时 {mins} 分钟\n"
                   f"待人工清单还剩 {left} 位")
            if blocked:
                msg += ("\n\n需您先修数据再补发（本次未动）：")
                for c in blocked[:8]:
                    msg += (f"\n· {c['email']}（{c['country'] or '缺国家'}）"
                            f"{c['name']}｜{c.get('action', '')}")
                if len(blocked) > 8:
                    msg += f"\n…（其余 {len(blocked) - 8} 位见日志）"
            if bad_list:
                msg += "\n\n仍失败的（已写明原因与处理动作）："
                for e, ctry, nm, why, act in bad_list[:8]:
                    msg += f"\n· {e}（{ctry}）{nm}：{why}"
                if len(bad_list) > 8:
                    msg += f"\n…（其余 {len(bad_list) - 8} 位见日志）"
                msg += "\n\n点「导出待补发清单」可拿到完整表格。"
            self.log(f"[补发] —— 本轮结束 ——补发成功 {ok_n}，仍失败 {bad_n}，"
                     f"用时 {mins} 分钟，待人工清单剩 {left} 位")
            # 补发结束：同样把发送记录回写到客户表（另存一份到 output/）
            try:
                self._writeback_after_round(
                    {str(e).strip().lower() for e in ok_emails} or None)
            except Exception as e:
                self.log("[回写] 自动回写出错：", e)
            self.q.put(("info", msg))
            self.q.put(("status", f"补发完成：成功 {ok_n} ｜失败 {bad_n}｜可以接着发下一个分表"))
            self.q.put(("call", lambda: self._mx_show_done(
                f"✔ 补发结束：成功 {ok_n} ｜仍失败 {bad_n} ｜用时 {mins} 分钟"
                f"｜待人工清单剩 {left} 位")))
            self._mx_running = False

        def _run():
            try:
                work()
            except Exception:
                import traceback
                self.log("[补发] 后台异常：\n" + traceback.format_exc())
                self.q.put(("msg", "补发过程中出错，详情见日志末尾"))
            finally:
                # 无论正常/停止/出错都要复位，否则下次点按钮被「上一轮还在进行中」挡住
                self._mx_running = False
                self._mx_stop = False
                self._mx_pause = False
                self._resend_sheet = ""          # 补发专用分表覆盖用完了就清掉
                self.q.put(("call", self._mx_finish_ui))
                self.q.put(("status", "补发结束"))

        threading.Thread(target=_run, daemon=True).start()
        self.log("[补发] 已启动：先自查发件箱，再逐个补发…")

    def _manual_drop(self, email: str) -> None:
        """补发成功 / 发现其实已成功 → 从待人工清单里移除。"""
        try:
            e = (email or "").strip().lower()
            n0 = len(self.manual)
            self.manual = [x for x in self.manual if (x.get("email") or "").lower() != e]
            if len(self.manual) != n0:
                self._manual_save()
                self.q.put(("call", self.refresh_manual_count))
        except Exception:
            pass

    def do_verify_outbox(self):
        """手动核对：扫一遍孚盟「发件箱」，看本分表客户哪些已有定时记录。"""
        if not self.engine:
            self._info("提示", "请先在「① 连接孚盟」里连接浏览器")
            return
        rows = getattr(self, "rows", []) or []
        emails = [(r.get("email") or "").strip() for r in rows]
        emails = [e for e in emails if e]
        if not emails:
            self._info("提示", "请先在「② 客户表」里载入客户表")
            return

        def run():
            try:
                pre = self.engine.preflight_scheduled(emails)
                found = {x.strip().lower() for x in pre.get("found", [])}
                miss = [e for e in emails if e.strip().lower() not in found]
                self.log(f"[核对] —— 发件箱核对结果 ——")
                if pre.get("found"):
                    for x in pre["found"]:
                        self.log(f"[核对]   ✓ 已有定时记录：{x}")
                    try:      # 顺手记入已发记录，防止下次重发
                        for x in pre["found"]:
                            self.engine.mark_scheduled(x)
                        self.log("[核对] 以上客户已记入「已发记录」，之后不会再重发")
                        self.q.put(("call", self.refresh_sent_count))
                    except Exception:
                        pass
                self.log(f"[核对] 未见到记录 {len(miss)} 位"
                         + (f"（前 8 位：{'、'.join(m.split('@')[0] for m in miss[:8])}…）"
                            if len(miss) > 8 else
                            (f"（{'、'.join(m.split('@')[0] for m in miss)}）" if miss else "")))
                self.q.put(("info", f"核对完成\n已有定时记录：{len(pre.get('found', []))} 位\n"
                                    f"未见到记录：{len(miss)} 位\n详见日志"))
            except Exception as e:
                self.log(f"[核对] 出错：{type(e).__name__}: {e}")
                self.q.put(("msg", f"核对失败：{e}"))

        threading.Thread(target=run, daemon=True).start()
        self.log("[核对] 正在打开孚盟「发件箱」扫描定时记录（几秒钟）…")

    def _on_tk_error(self, exc, val, tb):
        """Tk 回调里的未捕获异常统一在这里收口。

        没这个函数时：pythonw 没有控制台，异常只写进 app_err.log，
        界面上「点按钮毫无反应」—— 用户无从判断是卡了还是崩了。
        2026-09-26 就因此排查了一次 do_add_schedule 的 UnboundLocalError。
        """
        import traceback
        txt = "".join(traceback.format_exception(exc, val, tb))
        # 1) 落日志（含完整调用栈）
        try:
            self.log("[界面错误] " + txt.strip().replace("\n", "\n           "))
        except Exception:
            pass
        try:
            with open(APP / "logs" / "app_err.log", "a", encoding="utf-8") as f:
                f.write("\n=== " + now_cn().strftime("%Y-%m-%d %H:%M:%S") + " Tk回调异常 ===\n")
                f.write(txt)
        except Exception:
            pass
        # 2) 定位出错的最后一行代码，方便一眼看出问题
        last = ""
        try:
            for fr in traceback.extract_tb(tb)[::-1]:
                if "app.py" in fr.filename:
                    last = f"{Path(fr.filename).name} 第 {fr.lineno} 行：{fr.line}"
                    break
        except Exception:
            pass
        # 3) 弹窗告知（本次操作没生效，别让用户反复点）
        try:
            self._err(
                "操作没能完成",
                "这一步执行出错，所以界面上看不出变化（不是卡住）。\n\n"
                + (f"出错位置：\n{last}\n\n" if last else "")
                + f"错误类型：{val.__class__.__name__}: {val}\n\n"
                + "完整堆栈已写进 logs/app_err.log。\n"
                + "请把这句话截图发我，我来定位。")
        except Exception:
            pass

    def _on_tab_changed(self, event=None):
        """切到「⑥ 历史发送记录」时刷新一次：新加的模板、刚发完的记录都能同步进来。"""
        try:
            if self.nb.index("current") == 5:      # 第 6 个标签页
                self.do_hist_query(self._h_page)
        except Exception:
            pass

    def _hist_boot(self):
        """开界面时自动把历史发送记录查一遍（用户不用先点查询）。"""
        try:
            self._hist_fill_filters()
            self.do_hist_query(1)
        except Exception:
            pass

    # ---------------- ⑥ 历史发送记录 ----------------
    # 数据源：core/send_stats.py → data/send_flow.jsonl（每封成功发完就追加一条：
    # 时间 / 邮箱 / 分表 / 模板 / 主题）。所有历史轮的 matrix 都从这份流水算，
    # 不碰 sent_history.json —— 去重逻辑完全不受影响。
    #
    # 聚合口径（防重复计数）：
    #   行 = （模板, 主题, 分表）三元组；行内「去重人数」按邮箱集合算；
    #   跨分表合并 / 底部总计 一律用集合并集，同一客户不会在不同分表被算两遍。
    SORT_MAP = {
        "发送次数 ↓": ("count", True), "发送次数 ↑": ("count", False),
        "去重人数 ↓": ("people", True), "去重人数 ↑": ("people", False),
        "最近发送 ↓": ("last", True), "最早发送 ↑": ("first", False),
        "模板 A→Z": ("tpl", False), "分表 A→Z": ("sheet", False),
    }

    def _tab_history(self, nb):
        f = ttk.Frame(nb, padding=10)
        self.var_h_tpl = tk.StringVar()
        self.var_h_subj = tk.StringVar()
        self.var_h_sheet = tk.StringVar()
        self.var_h_from = tk.StringVar()
        self.var_h_to = tk.StringVar()
        self.var_h_kw = tk.StringVar()
        self.var_h_sort = tk.StringVar(value="发送次数 ↓")
        # 2026-10-25 新增：视图模式。「模板聚合」= 原来的一行一个模板+主题；
        # 「按日期」= 一天一行，能看到那天一共发了多少、用了哪几个模板。
        self.var_h_view = tk.StringVar(value="模板聚合（一行=一个模板+主题）")
        self.var_h_size = tk.StringVar(value="50")
        self.var_h_page = tk.StringVar(value="第 1 / 1 页")
        self.var_h_sum = tk.StringVar(value="点「查询」看结果")
        self.var_h_merge = tk.BooleanVar(value=False)
        self.var_h_hasfail = tk.BooleanVar(value=False)
        self.var_h_tplinfo = tk.StringVar(value="")
        self._h_page, self._h_pages, self._h_rows = 1, 1, []
        self._h_tpl_map: dict[str, str] = {}       # 模板下拉「显示值 → 真名」（历史项带前缀）

        # ---- ① 筛选区：模板 / 主题 / 分表 ----
        bar1 = ttk.Frame(f)
        bar1.grid(row=0, column=0, columnspan=6, sticky="ew")
        self.cmb_h_tpl = ttk.Combobox(bar1, width=26, textvariable=self.var_h_tpl)
        self.cmb_h_subj = ttk.Combobox(bar1, width=26, textvariable=self.var_h_subj)
        self.cmb_h_sheet = ttk.Combobox(bar1, width=26, textvariable=self.var_h_sheet)
        for i, (lab, cb) in enumerate((("模板", self.cmb_h_tpl),
                                       ("主题", self.cmb_h_subj),
                                       ("分表", self.cmb_h_sheet))):
            ttk.Label(bar1, text=lab + "：").grid(row=0, column=i * 2, sticky="w", pady=2)
            cb.grid(row=0, column=i * 2 + 1, sticky="w", padx=(0, 12))
            cb["values"] = []

        ttk.Label(f, textvariable=self.var_h_tplinfo, foreground="#666666",
                  wraplength=900, justify="left").grid(
            row=1, column=0, columnspan=6, sticky="w")

        # ---- ② 日期区间 + 关键字 ----
        bar2 = ttk.Frame(f)
        bar2.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        ttk.Label(bar2, text="日期从").pack(side="left")
        ttk.Entry(bar2, width=11, textvariable=self.var_h_from).pack(side="left")
        ttk.Label(bar2, text="到").pack(side="left")
        ttk.Entry(bar2, width=11, textvariable=self.var_h_to).pack(side="left")
        ttk.Label(bar2, text=" 关键字：").pack(side="left", padx=(10, 0))
        ent_kw = ttk.Entry(bar2, width=22, textvariable=self.var_h_kw)
        ent_kw.pack(side="left")
        ent_kw.bind("<Return>", lambda e: self.do_hist_query(1))

        # ---- ③ 排序 / 分页 / 合并开关 ----
        bar3 = ttk.Frame(f)
        bar3.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        ttk.Label(bar3, text="视图").pack(side="left")
        cb_view = ttk.Combobox(bar3, width=30, state="readonly", textvariable=self.var_h_view,
                               values=["模板聚合（一行=一个模板+主题）", "按日期（一天一行）"])
        cb_view.pack(side="left", padx=(0, 12))
        cb_view.bind("<<ComboboxSelected>>", lambda e: self._hist_view_changed())
        ttk.Label(bar3, text="每页").pack(side="left")
        ttk.Spinbox(bar3, from_=10, to=500, width=5,
                    textvariable=self.var_h_size).pack(side="left")
        ttk.Label(bar3, text="条  排序").pack(side="left", padx=(8, 0))
        ttk.Combobox(bar3, width=14, state="readonly", textvariable=self.var_h_sort,
                     values=list(self.SORT_MAP)).pack(side="left")
        ttk.Checkbutton(bar3, text="跨分表合并同一模板+主题",
                        variable=self.var_h_merge,
                        command=lambda: self.do_hist_query(1)).pack(side="left", padx=10)
        ttk.Checkbutton(bar3, text="含失败记录", variable=self.var_h_hasfail,
                        command=lambda: self.do_hist_query(1)).pack(side="left")

        # ---- ④ 结果表 ----
        cols = ("no", "tpl", "subj", "sheet", "sends", "people", "fail", "first", "last")
        self.tree_h = ttk.Treeview(f, height=18, show="headings", columns=cols)
        for c, txt, w in zip(cols, ("#", "模板", "主题", "分表", "发送次数",
                                    "去重人数", "失败", "首次发送", "最近发送"),
                             (40, 150, 240, 150, 78, 78, 56, 96, 96)):
            self.tree_h.heading(c, text=txt)
            # 2026-09-28：原来 tpl/subj/sheet 是 stretch=True，会把列压扁导致右侧列
            # 既看不全、又不产生横向溢出。统一 stretch=False，交给横向滚动条。
            self.tree_h.column(c, width=w, minwidth=40, stretch=False)
        self.tree_h.grid(row=4, column=0, columnspan=6, sticky="nsew", pady=8)
        # 2026-09-28：⑥页有 9 列，窗口一窄右边（首次/最近发送）就看不见。
        # 补横向滚动条，和②页客户表一个做法。
        hsb_h = ttk.Scrollbar(f, orient="horizontal", command=self.tree_h.xview)
        hsb_h.grid(row=5, column=0, columnspan=6, sticky="ew")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree_h.yview)
        sb.grid(row=4, column=6, sticky="ns")
        self.tree_h.configure(yscrollcommand=sb.set, xscrollcommand=hsb_h.set)
        self.tree_h.bind("<Shift-MouseWheel>", self._wheel_h)
        f.rowconfigure(4, weight=1)
        f.columnconfigure(0, weight=1)
        self.tree_h.bind("<Double-1>", self._hist_dblclick)

        # ---- ⑤ 分页 + 总计 ----
        bar4 = ttk.Frame(f)
        bar4.grid(row=5, column=0, columnspan=6, sticky="ew")
        ttk.Button(bar4, text="◀ 上一页", command=lambda: self.do_hist_query(
            max(1, self._h_page - 1))).pack(side="left")
        ttk.Label(bar4, textvariable=self.var_h_page, width=26).pack(side="left", padx=6)
        ttk.Button(bar4, text="下一页 ▶", command=lambda: self.do_hist_query(
            min(self._h_pages, self._h_page + 1))).pack(side="left")
        ttk.Label(bar4, textvariable=self.var_h_sum,
                  foreground="#1a4f8a").pack(side="left", padx=12)

        # ---- ⑥ 操作按钮 ----
        bar5 = ttk.Frame(f)
        bar5.grid(row=6, column=0, columnspan=6, sticky="ew", pady=(8, 0))
        ttk.Button(bar5, text="🔍 查询", command=lambda: self.do_hist_query(1)).pack(side="left")
        ttk.Button(bar5, text="清空筛选", command=self.do_hist_reset).pack(side="left", padx=6)
        ttk.Button(bar5, text="重发这一行…", command=self.do_hist_resend).pack(side="left", padx=6)
        ttk.Button(bar5, text="查看明细", command=self.do_hist_detail).pack(side="left", padx=6)
        ttk.Button(bar5, text="导出 CSV", command=self.do_hist_export).pack(side="left", padx=6)
        ttk.Button(bar5, text="从旧日志补录", command=self.do_hist_backfill).pack(side="left", padx=6)
        ttk.Label(f, text="双击一行也能看明细。旧记录的模板栏写「（旧记录·未记模板）」，"
                          "分表栏是按日志推断的（备注里会写「推断/沿用」）—— 这两种都不是真实采集值。",
                  foreground="#666666", wraplength=900, justify="left").grid(
            row=7, column=0, columnspan=6, sticky="w", pady=(6, 0))
        return f

    def _hist_template_options(self) -> list[tuple[str, str]]:
        """⑥ 页「模板」下拉的候选 = **当前模板 + 历史模板**，去重后排好序。

        数据来源（两条，缺谁都不报错）：
          ① 当前模板：① 页从孚盟读到的 self.cmb_tpl 列表；
             连上孚盟读到新列表后自动同步（do_templates 里回调刷新）；
             还没连孚盟时退回 ui.json 里缓存的上一次列表（只是缓存，不写回孚盟）。
             另外把「上次用过的模板」也补进来，避免没读到列表时选不中。
          ② 历史模板：data/send_flow.jsonl 里出现过的模板名
             （send_stats.template_stats，带首次/最近使用时间）。

        去重：按 _tpl_key（名字归一化）只留一条，**当前优先**。
        排序：①当前模板在前（保持孚盟给的顺序）→ ②历史模板在后（最近使用倒序，
              没有最近时间就退到首次使用时间，再退到模板名）。

        返回 [(显示文字, 实际模板名)]。只在历史里出现的加「（历史）」前缀标明来源，
        查询时用 self._h_tpl_map 把显示值还原成真名再传给 send_stats。
        """
        cur: list[str] = []
        # ①-a 孚盟实时列表（① 页读过才有）
        try:
            vals = self.cmb_tpl.cget("values") if hasattr(self, "cmb_tpl") else ()
            for v in (vals or ()):
                v = str(v).strip()
                if v:
                    cur.append(v)
        except Exception:
            pass
        # ①-b 没连孚盟 → 用上次读到的缓存（同样算"当前"这一类）
        if not cur:
            for v in (self.ui.get("tpl_list") or []):
                v = str(v).strip()
                if v:
                    cur.append(v)
        # ①-c 上次用过的模板（列表还没读回来时也能选到）
        pend = str(getattr(self, "_pending_tpl", "") or "").strip()
        if pend and pend not in cur:
            cur.append(pend)

        seen: dict[str, str] = {}          # 归一化键 → 真名（去重表）
        out: list[tuple[str, str]] = []
        for t in cur:
            k = _tpl_key(t)
            if not k or k in seen:
                continue
            seen[k] = t
            out.append((t, t))             # 当前模板：显示名 = 真名

        # ② 历史模板（当前列表里没有的那些），按最近使用倒序
        hist: list[tuple[str, str, str, str]] = []   # (显示, 真名, last, first)
        if send_stats is not None:
            try:
                for g in send_stats.template_stats(DATA):
                    t = str(g.get("tpl") or "").strip()
                    k = _tpl_key(t)
                    if not t or k in seen:
                        continue
                    seen[k] = t
                    hist.append((f"（历史）{t}", t, g.get("last", "") or "",
                                 g.get("first", "") or ""))
            except Exception:
                hist = []
        hist.sort(key=lambda x: (x[2] or x[3], x[1]), reverse=True)
        out += [(d, t) for d, t, _l, _f in hist]
        return out

    def _hist_fill_filters(self):
        """刷新⑥页三个下拉：模板（当前+历史）/ 主题 / 分表。

        ★ 只**读** self.cmb_tpl，不写它 —— ① 页模板的读取、选择、保存逻辑完全不动。
        ★ 任何一步出错都吞掉，历史为空/数据缺失时下拉照常渲染（只是项目少）。
        """
        # ---- 模板：当前 + 历史 ----
        try:
            opts = self._hist_template_options()
        except Exception:
            opts = []
        self._h_tpl_map = {d: t for d, t in opts}      # 显示值 → 真名
        try:
            self.cmb_h_tpl["values"] = [""] + [d for d, _ in opts]
        except Exception:
            pass
        # 保持原选中项：显示值还在就原样保留；显示值变了（例如加了「（历史）」前缀）
        # 就按真名找回同一个条目；实在找不到才清空（=全部）。
        try:
            cur = (self.var_h_tpl.get() or "").strip()
            vals = set(self.cmb_h_tpl["values"] or ())
            if cur and cur not in vals:
                real = self._h_tpl_map.get(cur) or cur
                for d, t in opts:
                    if t == real:
                        self.var_h_tpl.set(d)
                        break
                else:
                    self.var_h_tpl.set("")
        except Exception:
            pass
        try:
            n_cur = sum(1 for d, _ in opts if not str(d).startswith("（历史）"))
            self.var_h_tplinfo.set(
                f"模板下拉 = 当前 {n_cur} 个 + 历史 {len(opts) - n_cur} 个"
                f"（当前优先；历史按最近使用倒序；留空=不筛选）")
        except Exception:
            pass

        # ---- 主题 / 分表：照旧从流水里取 ----
        if send_stats is None:
            return
        try:
            m = send_stats.meta(DATA)
            for cb, vals in ((self.cmb_h_subj, m["subjects"]),
                             (self.cmb_h_sheet, m["sheets"])):
                cb["values"] = [""] + vals
        except Exception:
            pass

    def _hist_args(self) -> dict:
        """把界面上的筛选条件收集成 send_stats.query 的参数。"""
        sort, desc = self.SORT_MAP.get(self.var_h_sort.get(), ("count", True))
        try:
            size = int(float(self.var_h_size.get() or 50))
        except Exception:
            size = 50
        sel = (self.var_h_tpl.get() or "").strip()
        # 下拉里历史模板带「（历史）」前缀，这里用映射还原成真名再交给查询
        tpl = self._h_tpl_map.get(sel, sel)
        return dict(
            tpl="" if tpl in ("（全部）",) else tpl,
            subject=self.var_h_subj.get().strip(),
            sheet=self.var_h_sheet.get().strip(),
            date_from=self.var_h_from.get().strip(),
            date_to=self.var_h_to.get().strip(),
            kw=self.var_h_kw.get().strip(),
            only_ok=not self.var_h_hasfail.get(),
            merge=self.var_h_merge.get(),
            sort=sort, desc=desc, page_size=size)

    def do_hist_reset(self):
        for v in (self.var_h_tpl, self.var_h_subj, self.var_h_sheet,
                  self.var_h_from, self.var_h_to, self.var_h_kw):
            v.set("")
        self.var_h_sort.set("发送次数 ↓")
        self.var_h_merge.set(False)
        self.var_h_hasfail.set(False)
        # 2026-10-25：清筛选时顺手回到「模板聚合」视图（列标题也要跟着换回来）
        if self.var_h_view.get().startswith("按日期"):
            self.var_h_view.set("模板聚合（一行=一个模板+主题）")
            self._hist_view_changed()
        else:
            self.do_hist_query(1)

    def do_hist_query(self, page: int = 1):
        """查一页结果。数据量大时也不会卡：聚合是单趟 O(n)，
        排序/分页只作用在聚合后的几百个分组上（不是十几万行流水）。

        2026-10-25：加「按日期」视图 —— 一天一行，双击看当天全部客户。
        """
        if send_stats is None:
            self._info("缺模块", "没找到 core/send_stats.py，历史记录功能不可用。")
            return
        self._hist_fill_filters()
        self._h_page = max(1, int(page or 1))
        is_date = self.var_h_view.get().startswith("按日期")
        try:
            if is_date:
                res = send_stats.by_date(DATA, page=self._h_page, **self._hist_args_date())
            else:
                res = send_stats.query(DATA, page=self._h_page, **self._hist_args())
        except Exception as e:
            self._err("查询失败", f"{type(e).__name__}: {e}")
            return
        self._h_rows, self._h_page, self._h_pages = res["rows"], res["page"], res["pages"]
        self.tree_h.delete(*self.tree_h.get_children())
        base = (res["page"] - 1) * int(float(self.var_h_size.get() or 50))
        for i, g in enumerate(self._h_rows):
            if is_date:
                vals = (base + i + 1, g["date"], g["tpls"][:70], g["sheets"][:40],
                        g["sends"], g["people"], g["fail"], g["first"], g["last"])
            else:
                vals = (base + i + 1, g["tpl"], g["subject"][:60], g["sheet"],
                        g["sends"], g["people"], g["fail"], g["first"], g["last"])
            self.tree_h.insert("", "end", iid=str(i), values=vals)
        self.var_h_page.set(f"第 {res['page']} / {res['pages']} 页（共 {res['total']} 行）")
        gr = res["grand"]
        if is_date:
            self.var_h_sum.set(
                f"总计：{gr['days']} 天 ｜发送 {gr['sends']} 次 ｜实际收件人 {gr['people']} 位"
                f"（已去重）｜失败 {gr['fail']} ｜流水 {gr['scanned']} 行")
        else:
            self.var_h_sum.set(
                f"总计：{gr['groups']} 行 ｜发送 {gr['sends']} 次 ｜实际收件人 {gr['people']} 位"
                f"（已去重）｜失败 {gr['fail']} ｜{gr['sheets']} 个分表 / {gr['tpls']} 个模板"
                f"｜流水 {gr['scanned']} 行")

    def _hist_args_date(self) -> dict:
        """「按日期」视图的参数：日期区间 / 模板 / 分表 / 只看成功 / 每页条数。"""
        try:
            size = int(float(self.var_h_size.get() or 50))
        except Exception:
            size = 50
        sel = (self.var_h_tpl.get() or "").strip()
        tpl = self._h_tpl_map.get(sel, sel)
        return dict(
            tpl="" if tpl in ("（全部）",) else tpl,
            sheet=self.var_h_sheet.get().strip(),
            date_from=self.var_h_from.get().strip(),
            date_to=self.var_h_to.get().strip(),
            only_ok=not self.var_h_hasfail.get(),
            page_size=size)

    def _hist_view_changed(self):
        """切换「模板聚合 / 按日期」时把列标题换掉，再重新查询。"""
        is_date = self.var_h_view.get().startswith("按日期")
        if is_date:
            heads = ("#", "日期", "涉及模板（条数）", "涉及分表",
                     "发送条数", "去重人数", "失败", "当天首封", "当天末封")
        else:
            heads = ("#", "模板", "主题", "分表", "发送次数",
                     "去重人数", "失败", "首次发送", "最近发送")
        cols = ("no", "tpl", "subj", "sheet", "sends", "people", "fail", "first", "last")
        for c, t in zip(cols, heads):
            self.tree_h.heading(c, text=t)
        # 合并开关只在聚合视图有意义
        try:
            self.var_h_merge.set(False)
        except Exception:
            pass
        self.do_hist_query(1)

    def _hist_dblclick(self, _e=None):
        """双击分派：聚合视图看「这个模板发给了谁」，按日期视图看「这一天发给了谁」。"""
        if self.var_h_view.get().startswith("按日期"):
            self.do_day_detail()
        else:
            self.do_hist_detail()

    def do_day_detail(self):
        """某一天的全部发送明细（按日期视图双击 / 「查看当天客户」按钮）。"""
        if send_stats is None:
            return
        sel = self.tree_h.selection()
        if not sel:
            self._info("提示", "先在列表里点一行日期")
            return
        try:
            g = self._h_rows[int(sel[0])]
            day = g["date"]
        except Exception:
            return
        a = self._hist_args_date()
        try:
            recs = send_stats.day_records(DATA, day, tpl=a["tpl"],
                                          sheet=a["sheet"], only_ok=a["only_ok"])
        except Exception as e:
            self._err("读取失败", f"{type(e).__name__}: {e}")
            return

        win = tk.Toplevel(self)
        win.title(f"{day} 当天明细（{len(recs)} 条）")
        win.geometry("880x540")
        cols = ("no", "ts", "email", "tpl", "subject", "sheet", "st")
        tv = ttk.Treeview(win, show="headings", height=22, columns=cols)
        for c, t, w in zip(cols, ("#", "时间", "收件人", "模板", "主题", "分表", "状态"),
                           (42, 66, 220, 130, 210, 110, 52)):
            tv.heading(c, text=t)
            tv.column(c, width=w, minwidth=40, stretch=(c == "subject"))
        tv.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        for i, r in enumerate(recs):
            tv.insert("", "end", values=(
                i + 1,
                (r["ts"] or "")[11:19] if len(r["ts"] or "") > 10 else "",
                r["email"], (r["tpl"] or "")[:24], (r["subject"] or "")[:60],
                (r["sheet"] or "")[:18], "✓" if r["ok"] else "✗"))
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(bar, text=f"{day} ｜ 共 {len(recs)} 条 ｜ "
                            f"成功 {sum(1 for r in recs if r['ok'])} ｜ "
                            f"失败 {sum(1 for r in recs if not r['ok'])}",
                  foreground="#1a4f8a").pack(side="left")

        def _export():
            from tkinter import filedialog
            import csv
            p = filedialog.asksaveasfilename(
                defaultextension=".csv", initialfile=f"发送明细_{day}.csv",
                filetypes=[("CSV", "*.csv")])
            if not p:
                return
            try:
                with open(p, "w", encoding="utf-8-sig", newline="") as fp:
                    w = csv.writer(fp)
                    w.writerow(["时间", "收件人", "模板", "主题", "分表", "状态"])
                    for r in recs:
                        w.writerow([r["ts"], r["email"], r["tpl"], r["subject"],
                                    r["sheet"], "成功" if r["ok"] else "失败"])
                self._info("已导出", p)
            except Exception as e:
                self._err("导出失败", str(e))

        ttk.Button(bar, text="导出 CSV", command=_export).pack(side="right")

    def _hist_sel(self) -> dict | None:
        sel = self.tree_h.selection()
        if not sel:
            self._info("提示", "先在列表里点一行")
            return None
        try:
            return self._h_rows[int(sel[0])]
        except Exception:
            return None

    def do_hist_detail(self):
        """某一行都发给谁了（按时间倒序，最多列 500 条）。"""
        g = self._hist_sel()
        if not g or send_stats is None:
            return
        if g.get("sheet", "").startswith("（合并"):
            self._info("明细", "当前是「跨分表合并」视图，这一行是多个分表合并出来的。\n"
                               "请取消勾选「跨分表合并」后，再对单个分表看明细。")
            return
        rows = send_stats.detail(DATA, g["tpl"], g["subject"], g["sheet"])
        win = tk.Toplevel(self)
        win.title(f"明细：{g['subject'][:30]} ｜ {g['sheet']}")
        win.geometry("620x460")
        tv = ttk.Treeview(win, show="headings", height=20,
                          columns=("no", "email", "ts", "st"))
        for c, t, w in zip(("no", "email", "ts", "st"), ("#", "收件人", "时间", "状态"),
                           (45, 250, 145, 70)):
            tv.heading(c, text=t)
            tv.column(c, width=w)
        for i, r in enumerate(rows[:500], 1):
            tv.insert("", "end", values=(i, r["email"], r["ts"],
                                         "成功" if r["ok"] else "失败"))
        tv.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(win, text=f"共 {len(rows)} 条" +
                  ("（这里只显示前 500 条）" if len(rows) > 500 else "")
                  + f"｜重复出现的收件人 = 同一模板+主题给他发过多次").pack(pady=(0, 8))

    def do_hist_export(self):
        """导出当前筛选结果（全部页，不只是这一页）到 CSV。
        2026-10-25：跟着当前视图走 —— 按日期视图就导每天一行。"""
        if send_stats is None:
            return
        is_date = self.var_h_view.get().startswith("按日期")
        try:
            if is_date:
                args = self._hist_args_date()
                args["page_size"] = 100000
                res = send_stats.by_date(DATA, page=1, **args)
            else:
                args = self._hist_args()
                args["page_size"] = 100000          # 导出不分页
                res = send_stats.query(DATA, page=1, **args)
        except Exception as e:
            self._err("导出失败", str(e))
            return
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = now_cn().strftime("%Y%m%d_%H%M")
        tag = "按日期" if is_date else "按模板"
        out = OUT_DIR / f"历史发送记录_{tag}_{stamp}.csv"
        try:
            with out.open("w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                if is_date:
                    w.writerow(["日期", "发送条数", "去重人数", "失败",
                                "涉及模板", "涉及分表", "当天首封", "当天末封"])
                    for g in res["rows"]:
                        w.writerow([g["date"], g["sends"], g["people"], g["fail"],
                                    g["tpls"], g["sheets"], g["first"], g["last"]])
                    gr = res["grand"]
                    w.writerow([])
                    w.writerow(["合计", gr["sends"], gr["people"], gr["fail"],
                                f"{gr['days']} 天", "", "", ""])
                else:
                    w.writerow(["模板", "主题", "分表", "发送次数", "去重人数", "失败",
                                "首次发送", "最近发送"])
                    for g in res["rows"]:
                        w.writerow([g["tpl"], g["subject"], g["sheet"], g["sends"],
                                    g["people"], g["fail"], g["first"], g["last"]])
                    gr = res["grand"]
                    w.writerow([])
                    w.writerow(["合计", "", "", gr["sends"], gr["people"], gr["fail"], "", ""])
        except Exception as e:
            self._err("导出失败", str(e))
            return
        self.log(f"[历史记录] 已导出：{out}")
        self._info("已导出", f"{len(res['rows'])} 行\n{out}")

    def do_hist_backfill(self):
        """把 logs/run_*.log 里以前的成功记录补进流水（可以反复点，不会翻倍）。"""
        if send_stats is None:
            return
        try:
            r = send_stats.backfill(DATA, APP / "logs",
                                    xlsx_path=self.var_file.get())
        except Exception as e:
            self._err("补录失败", str(e))
            return
        self.do_hist_query(1)
        self._info("补录完成",
                   f"新增 {r['added']} 条，跳过重复 {r['skipped_dup']} 条\n"
                   f"扫了 {len(r['files'])} 份日志：{'、'.join(r['files'])}\n\n"
                   "注：旧日志没写模板名，模板栏统一填「（旧记录·未记模板）」；\n"
                   "主题栏取自日志里的「生成任务…主题「…」」这一句 —— 靠它区分第几轮。")

    def _sheet_people(self, path: str, sheet) -> list[dict]:
        """读某个分表的客户名单。

        ★ 刻意复用「② 客户表」那套读表路由（xlsxlight.read_sheet + guess_col 认列
          + normalize_emails 拆多邮箱），保证「统计看到的名单」和「真正发的时候」
          用的是同一把尺子，不会出现两边人数对不上。
        """
        try:
            rows = xlsxlight.read_sheet(path, sheet)
        except Exception as e:
            self.log("[重发] 读客户表失败：", e)
            return []
        if not rows:
            return []
        # 2026-09-29：同 do_load_sheet —— 表头自动识别 + 列名按内容纠正
        hr = pick_header_row(rows)
        header, _n = fix_col_names([str(h or "").strip() for h in rows[hr]],
                                   rows[hr + 1:],
                                   above=(rows[hr - 1] if hr else None))
        i_mail = guess_col(header, MAIL_KEYS)
        i_cty = guess_col(header, COUNTRY_KEYS)
        i_name = guess_col(header, NAME_KEYS)
        i_comp = guess_col(header, COMPANY_KEYS)
        if i_mail < 0:
            return []
        out, seen = [], set()
        for r in rows[hr + 1:]:
            cell = r[i_mail] if i_mail < len(r) else ""
            nm = str(r[i_name] if 0 <= i_name < len(r) else "").strip()
            cp = str(r[i_comp] if 0 <= i_comp < len(r) else "").strip()
            if nm and cp and nm.lower().replace(" ", "") == cp.lower().replace(" ", ""):
                nm = ""
            for e in normalize_emails(cell):
                if e in seen:
                    continue
                seen.add(e)
                out.append({"email": e,
                            "country": str(r[i_cty] if 0 <= i_cty < len(r) else "").strip(),
                            "name": nm})
        return out

    def do_hist_resend(self):
        """重发：复用这一行的「模板 + 主题」，发到指定的客户表/分表。

        逻辑要点（用户要的"重发前先看到底发多少人"）：
          ① 目标范围 = 指定客户表里的指定分表（可以和原分表不同）；
          ② 默认排除「同模板+同主题已经发过」的人 —— 这里是**跨分表**查的，
             因为在 A 分表发过、B 分表里又是同一个人，不该再挨一封；
          ③ 缺国家的先剔出来报个数（MX 定时要按国家算时区，没有国家发不了）；
          ④ 确认后才开跑，跑的时候把 _resend_sheet 设成目标分表，
             这样新流水会记在目标分表名下，统计才不会串到别的栏。
        """
        g = self._hist_sel()
        if not g:
            return
        # 2026-10-25：按日期视图的行只有「日期」，没有模板/主题，
        # 重发要靠模板+主题，所以这里挡一下，让用户切回聚合视图。
        if self.var_h_view.get().startswith("按日期"):
            self._info("提示", "「重发」需要模板和主题。\n"
                               "请先把视图切回「模板聚合」，再选那一行重发。\n"
                               "（按日期视图双击某天，可以看当天发给了谁、并导出 CSV）")
            return
        if not self.engine:
            self._info("提示", "请先到「① 连接孚盟」连上账号，重发要往孚盟写定时记录。")
            return
        if g.get("sheet", "").startswith("（合并"):
            self._info("提示", "这是合并出来的行，请先取消「跨分表合并」，"
                               "再对单个分表做重发。")
            return

        win = tk.Toplevel(self)
        win.title("重发：复用原模板 + 原主题")
        win.geometry("620x430")
        win.transient(self)
        f = ttk.Frame(win, padding=12)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text=f"模板：{g['tpl']}", wraplength=580, justify="left").grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(f, text=f"主题：{g['subject']}", wraplength=580, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w")
        ttk.Label(f, text=f"历史上这一组：已发 {g['sends']} 次 / {g['people']} 位"
                          f"（{g['first']} ~ {g['last']}）").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

        var_file = tk.StringVar(value=self.var_file.get())
        ttk.Label(f, text="发到这张客户表：").grid(row=3, column=0, sticky="w")
        ttk.Entry(f, textvariable=var_file, width=44).grid(row=3, column=1, sticky="w")
        var_sheet = tk.StringVar()
        cmb = ttk.Combobox(f, width=26, state="readonly", textvariable=var_sheet)
        cmb.grid(row=4, column=1, sticky="w", pady=4)
        ttk.Label(f, text="发到这个分表：").grid(row=4, column=0, sticky="w")

        def browse():
            p = filedialog.askopenfilename(
                parent=win, title="选择客户表",
                initialdir=str(Path(var_file.get()).parent if var_file.get()
                               else Path.home() / "Desktop"),
                filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("所有文件", "*.*")])
            if not p:
                return
            var_file.set(p)
            try:
                names = xlsxlight.sheet_names(p)
                cmb["values"] = names
                want = g.get("sheet", "")
                var_sheet.set(want if want in names else (names[0] if names else ""))
            except Exception as e:
                self._err("打不开", str(e))

        ttk.Button(f, text="浏览…", command=browse).grid(row=3, column=2, padx=6)

        var_when = tk.StringVar(value=self.var_when.get())
        ttk.Label(f, text="定时发送时间（当地时间，按客户国家换算）：").grid(
            row=5, column=0, sticky="w")
        ttk.Entry(f, textvariable=var_when, width=22).grid(row=5, column=1, sticky="w", pady=4)

        # 默认跟③页的防重复总开关：没勾「跳过已经发过的」→ 这里也默认不排除
        var_only_new = tk.BooleanVar(value=bool(getattr(self.cfg, "skip_sent", True)))
        ttk.Checkbutton(f, text="只发「本模板+本主题」还没发过的（跨分表查；"
                                "想重复发就取消勾选）",
                        variable=var_only_new).grid(
            row=6, column=0, columnspan=3, sticky="w")

        var_sum = tk.StringVar(value="")
        lab = ttk.Label(f, textvariable=var_sum, wraplength=580, justify="left",
                        foreground="#1a4f8a")
        lab.grid(row=7, column=0, columnspan=3, sticky="w", pady=8)

        win_plan: dict = {"rows": []}     # 预估出来的目标名单，点「开始」时直接用

        def load_sheets():
            p = var_file.get().strip()
            if not p:
                return
            try:
                names = xlsxlight.sheet_names(p)
                cmb["values"] = names
                want = g.get("sheet", "")
                var_sheet.set(want if want in names else (names[0] if names else ""))
            except Exception as e:
                self.log("[重发] 读分表名失败：", e)
        load_sheets()

        def recalc(*_):
            win.update_idletasks()
            p, sh = var_file.get().strip(), var_sheet.get().strip()
            if not p or not sh:
                var_sum.set("先把客户表和分表选好，这里会自动算出要发多少人")
                win_plan["rows"] = []
                return
            people = self._sheet_people(p, sh)
            if not people:
                var_sum.set(f"「{sh}」里没读到邮箱，检查一下这一栏是不是邮箱列")
                win_plan["rows"] = []
                return
            done = set()
            if var_only_new.get() and send_stats is not None:
                try:
                    # ★ 跨分表查：只要这个模板+主题给这个人发过（不管在哪个分表）
                    # 就排除 —— 和查询端「跨分表不重复计数」是同一套口径。
                    done = send_stats.sent_emails(DATA, tpl=g["tpl"],
                                                  subject=g["subject"])
                except Exception:
                    done = set()
            targets = [p2 for p2 in people if p2["email"] not in done]
            no_cty = [p2 for p2 in targets if not p2["country"]]
            ok_rows = [p2 for p2 in targets if p2["country"]]
            win_plan["rows"] = ok_rows
            var_sum.set(
                f"目标范围：「{sh}」共 {len(people)} 位"
                f"｜本模板+主题已发过 {len(done)} 位（自动排除）\n"
                f"预估发送 {len(ok_rows)} 位"
                f"｜另有 {len(no_cty)} 位缺国家，发不了会被跳过"
                + ("" if var_only_new.get() else "\n⚠ 没勾选排除项：已发过的也会重发一遍"))

        # win_plan 已在上面定义好 ← chỗ này重要：var_sheet.trace_add 会立刻触发一次 recalc
        var_sheet.trace_add("write", recalc)
        var_only_new.trace_add("write", recalc)
        recalc()

        def start():
            targets = list(win_plan.get("rows") or [])
            try:
                when = parse_when(var_when.get())
            except ValueError as e:
                messagebox.showerror("时间不对", str(e), parent=win)
                return
            if not targets:
                messagebox.showinfo("没有可发的", "预估人数为 0，先换个分表或取消排除试试。",
                                    parent=win)
                return
            if getattr(self, "_mx_running", False):
                self._err("正在发送", "上一轮还在跑，请先「停止」。")
                return
            sh = var_sheet.get().strip()
            only_new = bool(var_only_new.get())     # Tk 变量只在这里（主线程）读一次
            self._apply_when_basis()                # 2026-10-05：重发也跟随时间基准
            basis6 = str(getattr(self.cfg, "when_basis", "local") or "local")
            mx_local = when.strftime("%Y/%m/%d %H:%M")
            mx_local, _past_warn = self._fix_past_when(mx_local)   # 过去时间 = 齐发 = 风控
            if _past_warn:
                self.log("[重发] " + _past_warn.replace("\n", " "))
                try:
                    var_when.set(mx_local.replace("/", "-"))
                except Exception:
                    pass
                self._toast(_past_warn, "warn", 22000)
            if not self._ask(
                    "确认重发",
                    f"模板「{g['tpl']}」\n主题「{g['subject'][:40]}」\n"
                    f"目标：{Path(var_file.get()).name} / {sh}\n"
                    f"预计 {len(targets)} 位，"
                    f"{'北京时间' if basis6 == 'beijing' else '当地时间'} {mx_local}\n"
                    + (self._tz_preview_text(mx_local, targets) or "")
                    + "\n"
                    "会像正常发送一样逐个写入孚盟定时（可随时「停止」），\n"
                    "发完会在「⑥ 历史发送记录」里记到这个分表名下。\n\n确定？"):
                return
            win.destroy()

            def work():
                ok = fail = 0
                n = len(targets)
                # _resend_sheet 会盖掉界面上的分表名 → 新流水落到目标分表
                # _round_mode_on=True → 发件箱里上一轮的同名旧记录不再当成"已发过"
                self._resend_sheet = sh
                self._round_mode_on = True
                # 去重跟弹窗里那个勾：勾了「只发没发过的」才启用防重复，
                # 没勾（=要重复发）就一路放行，发件箱里有旧记录也不算数。
                # （Tk 变量不能在工作线程读，所以 only_new 已在外面抓成普通 bool）
                self._dedupe_on = only_new
                self._mx_stop = self._mx_pause = False
                self._mx_running = True
                self.log(f"[重发] 开始：模板「{g['tpl']}」主题「{g['subject'][:30]}」"
                         f"→ {Path(var_file.get()).name} / {sh}，共 {n} 位")
                try:
                    when_list = self._spread_local_times(mx_local, n)   # 同样错开，别挤同一秒
                    for i, p2 in enumerate(targets, 1):
                        if getattr(self, "_mx_stop", False):
                            self.log(f"[重发] 已停止：处理到第 {i - 1}/{n} 位")
                            break
                        while getattr(self, "_mx_pause", False) \
                                and not getattr(self, "_mx_stop", False):
                            time.sleep(0.5)
                        self.q.put(("status", f"重发中：{i}/{n} ｜成功 {ok} ｜失败 {fail}"))
                        res, cls, sec = self._send_one_with_retry(
                            p2["email"], p2["country"], g["subject"], g["tpl"],
                            when_list[i - 1] if i <= len(when_list) else mx_local,
                            True, tag=f"[重发] 第 {i}/{n} 位",
                            name=(p2.get("name") or ""))
                        if res.get("ok"):
                            ok += 1
                            self.log(f"[重发] ✓ 第 {i}/{n} 位 {p2['email']} 已定时｜用时 {sec}s")
                        else:
                            fail += 1
                            why = (cls or {}).get("name", "未分类")
                            self.log(f"[重发] ✗ 第 {i}/{n} 位 {p2['email']} 失败"
                                     f"｜{why}｜{res.get('msg', '')[:60]}")
                finally:
                    self._resend_sheet = None
                    self._round_mode_on = False
                    # 2026-09-29：同定时通道一样，跑完要把运行态和界面一起复位
                    self._mx_running = False
                    self._mx_stop = False
                    self._mx_pause = False
                    self.log(f"[重发] —— 结束 ——成功 {ok}，失败 {fail}，共 {n} 位")
                    self.q.put(("call", lambda: self._mx_show_done(
                        f"✔ 重发结束：成功 {ok} ｜失败 {fail} ｜共 {n} 位")))
                    self.q.put(("status", f"重发完成：成功 {ok} ｜失败 {fail}"))
                    self.q.put(("call", self._mx_finish_ui))
                    self.q.put(("call", lambda: self.do_hist_query(1)))   # 刷新统计

            threading.Thread(target=work, daemon=True).start()

        bar = ttk.Frame(f)
        bar.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Button(bar, text="开始重发", command=start).pack(side="left")
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="left", padx=8)

    # ---------------- ⑤ 运行 ----------------
    def _tab_run(self, nb):
        f = ttk.Frame(nb, padding=12)
        bar = ttk.Frame(f)
        bar.pack(fill="x")
        self.btn_start = ttk.Button(bar, text="开始发送", command=self.do_start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(bar, text="停止", command=self.do_stop)
        self.btn_stop.pack(side="left", padx=8)
        ttk.Button(bar, text="导出结果 CSV", command=self.do_export).pack(side="left", padx=8)
        ttk.Button(bar, text="按发送结果回写", command=self.do_writeback).pack(side="left", padx=8)
        ttk.Button(bar, text="按已发件箱核对并回写", command=self.do_writeback_sent).pack(side="left", padx=8)
        self.pb = ttk.Progressbar(f, maximum=100)
        self.pb.pack(fill="x", pady=8)
        self.txt_log = tk.Text(f, height=26, wrap="word")
        self.txt_log.pack(fill="both", expand=True)
        return f

    def _collect_rows(self) -> list[dict]:
        cols = [str(h) for h in self.header]
        def idx(cmb):
            v = cmb.get()
            return cols.index(v) if v in cols else -1
        i_mail, i_name = idx(self.cmb_mail), idx(self.cmb_name)
        i_country, i_comp = idx(self.cmb_country), idx(self.cmb_company)
        out, bad = [], []
        for n, r in enumerate(self.rows, start=2):   # 第 1 行是表头
            cell = r[i_mail] if 0 <= i_mail < len(r) else ""
            got = normalize_emails(cell)
            if not got:
                if str(cell).strip():
                    bad.append((n, str(cell)))
                continue
            nm = r[i_name].strip() if 0 <= i_name < len(r) else ""
            cp = r[i_comp].strip() if 0 <= i_comp < len(r) else ""
            # 名字列里填的其实是公司名时（「谷歌搜索客户」整列如此），
            # 不拿它当称呼 —— 交给称呼模块降级成通用称呼，别发出 "Hi körschgen,"。
            if nm and cp and nm.lower().replace(" ", "") == cp.lower().replace(" ", ""):
                nm = ""
            # 同理：名字列填的是国家名（表格错列时会有）也不当称呼
            cty_chk = r[i_country].strip() if 0 <= i_country < len(r) else ""
            if nm and cty_chk and nm.lower() == cty_chk.lower():
                nm = ""
            for e in got:
                out.append({"row_index": n, "email": e,
                            "name": nm,
                            "country": r[i_country].strip() if 0 <= i_country < len(r) else "",
                            "company": cp})
        if bad:
            self.log(f"有 {len(bad)} 行的邮箱格子里认不出邮箱，已跳过：{bad[:5]}")
        return out

    def do_start(self):
        self._apply_settings()
        if not self.engine:
            self._info("提示", "请先连接孚盟")
            return
        rows = self._collect_rows()
        if not rows:
            self._info("提示", "没读到客户邮箱，请先载入客户表并选好邮箱列")
            return
        if not self._ask(
                "确认【立即发送】",
                "注意：这是「立即发送」，点「是」马上逐封发出，不是定时！\n"
                "（想按客户当地时间定时，请去「④ 定时发送」页。）\n\n"
                f"共 {len(rows)} 位客户，模板「{self.cmb_tpl.get()}」\n"
                f"主题「{self.cmb_subj.get()[:40]}」\n确定立即发出？"):
            return
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        subj = self.cmb_subj.get().strip()
        tpl = self.cmb_tpl.get().strip()

        def on_prog(i, total, res):
            self.q.put(("progress", i * 100 / total))

        def run():
            try:
                summary = self.engine.run_batch(rows, subj, tpl, on_progress=on_prog)
                self.results = summary["results"]
                # 立即发送也落流水（统一口径：模板+主题+分表）
                try:
                    for r in (summary.get("results") or []):
                        if isinstance(r, dict) and r.get("ok"):
                            self._record_send(r.get("email", ""), tpl, subj, True, "立即发送")
                except Exception:
                    pass
                self.log(f"完成：成功 {summary['ok']} / 失败 {summary['fail']} / 跳过 {summary['skip']}")
            except Exception as e:
                self.log("运行出错:", e)
            finally:
                self.q.put(("done", ""))
        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def do_stop(self):
        self._mx_stop = True            # 直接串行写入通道：立刻停
        self._mx_pause = False          # 2026-10-05：停止时把暂停标记一并清掉，
                                        # 免得按钮停在「继续发送」的橙色状态
        if self.engine:
            self.engine.stop()
        self.scheduler.request_stop()
        self.btn_sched.config(text="暂停")
        self._sync_pause_btn()
        self.log("已请求停止：正在处理的那位客户收尾后停下（最多几秒）。")

    def do_export(self):
        if not self.results:
            self._info("提示", "还没有发送结果")
            return
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".csv",
                                         initialfile=f"发送结果_{today_cn()}.csv",
                                         filetypes=[("CSV", "*.csv")])
        if not p:
            return
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["行号", "邮箱", "客户名", "国家", "是否成功", "说明", "时间"])
            for r in self.results:
                w.writerow([r.get("row_index"), r.get("email"), r.get("name"),
                            r.get("country"), "成功" if r.get("ok") else "失败",
                            r.get("msg"), r.get("time")])
        self._info("好了", p)

    def do_writeback(self):
        if not self.results:
            self._info("提示", "还没有发送结果")
            return
        col = (self.var_col.get() or "L").strip().upper()
        path = self.var_file.get()
        if not path:
            self._info("提示", "没有原表路径")
            return
        cells = {f"{col}{r['row_index']}": today_cn()
                 for r in self.results if r.get("ok") and not r.get("dry")}
        self._guarded_writeback(path, col, cells)

    def _guarded_writeback(self, path: str, col: str, cells: dict):
        """回写的统一安全闸：目标格有内容先列出来确认；空 cells 直接提示。"""
        if not path:
            self._info("提示", "没有原表路径")
            return
        if not cells:
            self._info("提示", "没有需要回写的行")
            return

        # 安全闸：目标格必须都是空的，否则会覆盖原有数据
        try:
            cur = xlsxlight.read_sheet(path, self.cmb_sheet.get() or 0)
        except Exception as e:
            self._err("读表失败", str(e))
            return
        ci = xlsxlight.col_to_index(col)
        occupied = []
        for ref in cells:
            rn = int("".join(ch for ch in ref if ch.isdigit()))
            row = cur[rn - 1] if rn - 1 < len(cur) else []
            if ci < len(row) and str(row[ci]).strip():
                occupied.append(f"{ref}（现在是“{str(row[ci])[:12]}”）")
        if occupied and not self._ask(
                "目标列不是空的",
                f"{col} 列这些格子已经有内容，回写会覆盖它们：\n" +
                "\n".join(occupied[:8]) +
                ("\n..." if len(occupied) > 8 else "") +
                "\n\n仍然继续吗？（建议先在 Excel 里在 K 列后插入一个空列）"):
            return
        try:
            xlsxlight.write_cells(path, self.cmb_sheet.get() or 0, cells)
            self._info("好了", f"已把 {len(cells)} 个日期（{today_cn()}）写回 {col} 列")
        except Exception as e:
            self._err("回写失败", str(e))

    def do_writeback_sent(self):
        """按「已发件箱」的真实记录核对后再回写——草稿/没发出去的一律不算。"""
        if not self.engine:
            self._info("提示", "请先连接孚盟")
            return
        rows = self._collect_rows()
        if not rows:
            self._info("提示", "没读到客户邮箱，请先载入客户表并选好邮箱列")
            return
        path = self.var_file.get()
        if not path:
            self._info("提示", "没有原表路径")
            return
        col = (self.var_col.get() or "L").strip().upper()
        subj = self.cmb_subj.get().strip()

        def run():
            try:
                res = self.engine.confirm_sent_from_outbox(rows, subj)
                self.q.put(("sent_report", {"res": res, "col": col, "path": path}))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.q.put(("msg", "核对已发件箱失败：\n" + str(e)))
        threading.Thread(target=run, daemon=True).start()

    def _on_sent_report(self, payload: dict):
        res, col, path = payload["res"], payload["col"], payload["path"]
        rows = res["rows"]
        ok_rows = [r for r in rows if r["confirmed"]]
        bad_rows = [r for r in rows if not r["confirmed"] and r.get("email")]
        if res.get("same_prefix"):
            self._warn(
                "注意", "已发件箱里只显示邮箱 @ 前缀，这些前缀有多位客户共用：\n"
                + "、".join(res["same_prefix"][:8]) +
                "\n核对按前缀+主题匹配，请回写后抽查这几行。")
        if not self._ask(
                "核对结果",
                f"已发件箱共 {len(res['texts'])} 条记录。\n"
                f"确认发出：{len(ok_rows)}；没找到记录：{len(bad_rows)}。\n"
                f"只把确认发出的 {len(ok_rows)} 行的日期写回 {col} 列，继续？"):
            return
        if bad_rows:
            preview = "\n".join(f"  第{r['row_index']}行 {r['email']}" for r in bad_rows[:8])
            self._info(
                "这些行没在已发件箱找到，不回写",
                preview + ("\n..." if len(bad_rows) > 8 else "") +
                "\n（定时未到点/发送失败的都在这里；确认发出去之后再核对一次即可）")
        cells = {f"{col}{r['row_index']}": today_cn() for r in ok_rows}
        self._guarded_writeback(path, col, cells)


def main():
    # pythonw 没有控制台，任何异常都要落到文件里，否则出错无迹可寻
    (APP / "logs").mkdir(exist_ok=True)
    errf = open(APP / "logs" / "app_err.log", "a", encoding="utf-8")
    sys.stdout = _Tee(sys.__stdout__, errf)
    sys.stderr = errf
    try:
        App().mainloop()
    except Exception:
        import traceback
        traceback.print_exc(file=errf)
        errf.flush()


class _Tee:
    def __init__(self, *streams):
        self.streams = [s for s in streams if s is not None]

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
                st.flush()
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass


if __name__ == "__main__":
    main()
