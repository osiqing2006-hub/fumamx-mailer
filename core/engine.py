# -*- coding: utf-8 -*-
"""发信引擎：把 WebMX 包装成"给一批客户依次发一封"的批处理。

设计红线：
  - 正文一律由孚盟 MX 端模板生成，本地绝不拼 HTML
  - 不注入任何自动化邮件头
  - 账号密码只由用户在浏览器里手工输入，程序不接触、不落盘
  - 每一封都有发送前自检，任一不过就不发
"""
from __future__ import annotations

import json
import random
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

from . import cdp
from .mx_web import MXWebError, WebMX, WebSettings

try:                                   # 国家名补全（印尼 -> 印度尼西亚 等）
    from .result_export import fix_country
except Exception:                      # 模块缺失时不影响主流程
    def fix_country(name: str) -> str:
        return name

CN = timezone(timedelta(hours=8))          # 中国时间（记录发送日用它）

_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def normalize_emails(text) -> list[str]:
    """把各种花样的粘贴内容洗成干净邮箱列表。

    能处理：
      全角＠、中文分号/顿号/逗号、换行、首尾空格和引号、
      "张三 <a@b.com>" 这种带名字的、mailto: 前缀、一行里多个地址。
    """
    if text is None:
        return []
    s = str(text)
    s = (s.replace("＠", "@").replace("＠", "@")
          .replace("；", ";").replace("，", ",").replace("、", ",")
          .replace("，", ",").replace("\n", ";").replace("\r", ";")
          .replace("\t", ";"))
    s = re.sub(r"mailto:\s*", "", s, flags=re.I)
    out: list[str] = []
    for m in _EMAIL_RE.finditer(s):
        e = m.group(0).strip().strip(".,;\"'()<>[]{} ").lower()
        if e and e not in out:
            out.append(e)
    return out


def today_cn() -> str:
    """yy.m.d 格式，例：26.9.20"""
    n = datetime.now(CN)
    return f"{n.year % 100}.{n.month}.{n.day}"


@dataclass
class EngineConfig:
    profile_dir: str = ""
    browser_exe: str = ""            # 留空=自动找本机的 Edge/Chrome；装进 OCS 后会指向 OCS 的 Chrome
    port: int = 9333
    login_url: str = "https://www.fumamx.com/"
    login_wait_min: int = 20
    step_delay: float = 1.2
    # 封间等待（秒）：**只在关闭随机间隔时生效**。
    # 2026-09-21 用户曾要求去掉随机等待；2026-09-30 因孚盟风控（发太快 →
    # 服务器 Connection reset、发件箱一片「发送失败」）要求加回随机间隔，
    # 开关就是下面的 human_mode（默认开）。
    require_chip: bool = False       # 库外地址不成芯片时仍放行定时（孚盟实际会发）
    interval_min: float = 0.6
    interval_max: float = 1.5
    human_mode: bool = True          # 随机间隔（防限流）：封间 45-110 秒随机 +
                                     # 连发一拨后休息几分钟 + 偶发多停一会
    # 定时发送时，每位客户的「当地时间」依次错开多少分钟（0=全部同一时刻）。
    # 2026-09-30：原来 N 位客户共用一个当地时间 → 孚盟服务器同一秒要投递 N 封
    # 带大附件的邮件 → Connection reset（发件箱一片「发送失败」）。
    # 错开之后投递由孚盟服务器按点完成，**排队完就可以关电脑**。
    # 默认 5 分钟：81 位 ≈ 6.7 小时（当地 9:00 起，当天下午发完）。
    spread_gap: int = 5
    # 2026-10-06：错开方式。常见需求是「邮件必须在**收件人当地 9 点**定时发送」，
    #   而原来的「顺序错开」把第 1 位之后的每一封都往后推（第 67 位 13:46、
    #   末位直接漂到次日），等于绝大多数客户都收不到"当地 9 点"。
    #   "same"  ：所有人写**同一个时间**（= 用户填的那个），一行不差。默认。
    #   "even"  ：旧的顺序错开（每封往后推 spread_gap 分钟，末位会漂到次日）。
    #   "window"：都落在 [基准时间, 基准+spread_gap] 内随机，既不齐发也不漂远。
    # 说明：定时通道的投递是**孚盟服务器**按点完成的，脚本只负责排队写进去，
    #   所以"齐发"的压力在孚盟那一侧，不在本机。
    spread_mode: str = "same"
    # 发送前清掉模板自带的抄送（2026-10-04）：
    # 实测大量邮件抄送了一个公共地址，等于每封投 2 个地址、投递量翻倍
    # （246 个失败现场里 193 个带抄送）。默认 False = 保持原样，需要时勾上。
    clear_cc: bool = False
    # 2026-10-05：④页定时时间到底是「北京时间」还是「收件人当地时间」。
    #   "local"（默认/旧行为）：填什么就当各客户自己的当地时间写进孚盟
    #        → 用户填 09:00 时，英国客户当地 9 点发（=北京 16 点），
    #          美国客户当地 9 点发（=北京 22 点），各国发作息差 6~12 小时。
    #   "beijing"：用户填北京时间，脚本按客户国家时区换算成当地时间再写。
    #        → 填 09:00 就是北京时间 9 点整发出，各国同一时刻收到。
    # 错开（spread_gap）在两种模式下都成立：序列在 UTC 轴上始终是均匀递增的。
    when_basis: str = "local"
    # 2026-10-05：是否允许把浏览器拉到前台（镜像 WebSettings.auto_raise，
    # 放在这里是为了能存进 settings.json 持久化）。默认关 = 绝不打断用户。
    auto_raise: bool = False
    # 2026-10-09：分批发送 —— 哪些分表「已经整轮跑完」。
    #   键 = "<客户表全路径>|<分表名>"，值 = {"total":N,"at":"YYYY-MM-DD HH:MM"}。
    #   用户填了「本批发送数量」时，跑完最后一批就写进来；下次再点生成任务会
    #   提示「本分表已发完」，避免手滑把整张表重发一遍。
    batch_done: dict = field(default_factory=dict)
    work_start: int = 9               # 只在这个时间段发（中国时间）
    work_end: int = 18
    only_work_hours: bool = True
    skip_sent: bool = True            # 跳过已发过的地址
    dry_run: bool = False             # 演练：只填不点发送
    # 2026-09-22 提速用：
    fast_mode: bool = True            # 极速模式：页面到位就走，不再死等固定秒数
    verify_every: int = 10            # 发件箱核对频率：1=每位都核对；N=每 N 位核对一次；
                                      # 0=中途不核对，整轮跑完统一核对一次（最快）
    auto_restore: bool = True         # Edge 窗口被最小化时自动恢复（最小化会让鼠标点击落空）
    auto_revive: bool = True          # 页面僵死/白屏时自动刷新恢复（等于自动按 F5）
    # 2026-09-26：新一轮——同一批客户换个模板/主题再发一轮。
    # False（默认）= 老规矩：这个邮箱发过就跳过；True = 只有同主题发过才跳过。
    new_round: bool = False
    extras: dict = field(default_factory=dict)


class Engine:
    def __init__(self, cfg: EngineConfig, log=print, data_dir: str | Path = ""):
        self.cfg = cfg
        self.log = log
        self.data_dir = Path(data_dir or ".")
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.sent_file = self.data_dir / "sent_history.json"
        self.w: WebMX | None = None
        self._stop = False
        self.sent: dict[str, str] = self._load_sent()

    # ---------------- 发送历史 ----------------
    def _load_sent(self) -> dict[str, str]:
        if self.sent_file.exists():
            try:
                return json.loads(self.sent_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _save_sent(self) -> None:
        """落盘已发记录。必须是**合并写**，不能全量覆盖。

        2026-09-23 事故：原写法是 `write_text(dumps(self.sent))`，而 self.sent 只是
        本实例的内存快照。若有多个实例/脚本同时跑（或本程序前后两次启动），各自
        快照不同，后写的那次会把先写的记录**整批冲掉** —— 当天 15~16 时那轮
        55 位客户的记录就是这么丢的（日志明明成功，记录里却没有 → 下次会被重发）。
        现在改成：先读盘上最新 → 合并内存 → 原子写回，任何实例都不会覆盖别人。
        """
        try:
            disk = self._load_sent()
        except Exception:
            disk = {}
        if isinstance(disk, dict):
            disk.update(self.sent)      # 内存里的新记录优先
            self.sent = disk            # 内存也同步成合并结果，避免下次又覆盖
        try:
            tmp = self.sent_file.with_suffix(self.sent_file.suffix + ".tmp")
            tmp.write_text(json.dumps(self.sent, ensure_ascii=False, indent=1),
                           encoding="utf-8")
            tmp.replace(self.sent_file)  # 原子替换，防止写一半被别的进程读到
        except Exception:
            # 兜底：原子替换失败就退化为直接写
            self.sent_file.write_text(
                json.dumps(self.sent, ensure_ascii=False, indent=1), encoding="utf-8")

    def already_sent(self, email: str) -> bool:
        """查是否已发过。会先看一眼文件有没有被别人改过（改过就重载），
        避免本实例拿着旧快照把别人的新记录当不存在 → 重复发送。"""
        self._sync_sent_if_changed()
        return email.strip().lower() in self.sent

    def _sync_sent_if_changed(self) -> None:
        """盘上文件比内存新（别的实例/手工补录改过）就把差异合进内存。"""
        try:
            mt = self.sent_file.stat().st_mtime
        except Exception:
            return
        if mt == getattr(self, "_sent_mtime", None):
            return
        try:
            disk = self._load_sent()
            if isinstance(disk, dict) and len(disk) > len(self.sent):
                self.sent = disk          # 以盘上最新的为准（只会多不会少）
        except Exception:
            pass
        self._sent_mtime = mt

    def reload_sent(self) -> None:
        """外面动过 sent_history.json（清空/清分表）之后，重新读回内存。"""
        self.sent = self._load_sent()

    def sent_today(self) -> int:
        d = datetime.now(CN).strftime("%Y-%m-%d")
        return sum(1 for v in self.sent.values() if v == d)

    # ---------------- 连接 ----------------
    def connect(self, wait_login: bool = True) -> bool:
        b = cdp.Browser(self.cfg.port)
        if not b.alive():
            self.log("正在打开浏览器 ...")
            b = cdp.launch(port=self.cfg.port,
                           edge_path=self.cfg.browser_exe or None,
                           user_data_dir=self.cfg.profile_dir,
                           url=self.cfg.login_url)
        tab = None
        # 2026-09-26 修复：原来无脑绑第一个标签页。Edge 里若 about:blank 或别的页
        # 排在第一位，脚本就绑到空白页，用户看的孚盟页是另一个标签页 ——
        # 表现是"页面明明加载好了，却读不到模板/显示未登录"。现在优先挑孚盟的页。
        try:
            fts = [t for t in b.targets() if "fumamx" in (t.get("url") or "")]
            if fts:
                tab = b.attach(fts[-1])
        except Exception:
            tab = None
        if tab is None:
            tab = b.first_tab()
        if tab is None:
            raise MXWebError("没有可用的标签页")
        if "fumamx" not in (tab.url or ""):
            self.log("当前标签页不在孚盟上，正在带浏览器去孚盟首页 …")
            try:
                tab.goto(self.cfg.login_url, timeout=60)
            except Exception as e:
                self.log("（打开孚盟失败，继续用当前页试）", e)
        try:
            # 2026-10-05：原来无条件 maximized —— 每次连接都把浏览器铺满整屏，
            # 直接盖住用户正在看的窗口。改成只有真被最小化时才恢复成普通窗口。
            # 2026-10-06：恢复必须走"不激活"（Win32 SW_SHOWNOACTIVATE）——
            # CDP 的 setWindowBounds 会顺手把 Edge 抢到前台，一连接就把用户的
            # 窗口挤走。见 cdp._win_restore_no_activate 的注释与实测记录。
            if b.window_state(tab.id) == "minimized":
                if not b.restore_no_activate(tab.id):
                    b.set_window_state(tab.id, "normal")
        except Exception:
            pass
        ws = WebSettings(browser="", port=self.cfg.port,
                         profile_dir=self.cfg.profile_dir,
                         login_url=self.cfg.login_url,
                         login_wait_minutes=self.cfg.login_wait_min,
                         step_delay=self.cfg.step_delay,
                         fast=bool(self.cfg.fast_mode),
                         auto_restore=bool(getattr(self.cfg, "auto_restore", True)),
                         # 默认关：不允许把浏览器拽到用户面前
                         auto_raise=bool(getattr(self.cfg, "auto_raise", False)))
        self.w = WebMX(ws, log=self.log)
        self.w.browser, self.w.tab = b, tab
        if wait_login and not self.w.is_logged_in():
            self.log("请在打开的浏览器窗口里登录孚盟MX（程序不碰您的账号密码）...")
            if not self._wait_login():
                return False
        self.log("已登录：" + (self.w.tab.url or ""))
        return True

    def _wait_login(self) -> bool:
        """轮询登录状态。页面抖动/超时一律按"还没登录"继续等，绝不向上抛。"""
        deadline = time.time() + self.cfg.login_wait_min * 60
        while time.time() < deadline:
            if self._stop:
                return False
            try:
                if self.w.is_logged_in():
                    return True
            except Exception as e:
                self.log("（检测登录时页面未就绪，继续等待）", type(e).__name__)
            time.sleep(2)
        self.log("等待登录超时（超过 %d 分钟）" % self.cfg.login_wait_min)
        return False

    def check_login(self) -> bool:
        """只查一次当前登录状态（给界面上的「检测登录状态」按钮用），不等待。"""
        self._check()
        try:
            st = self.w.login_state()
        except Exception:
            st = "ERR"
        ok = (st == "APP")
        # 2026-10-12：把判定依据也打出来。上次孚盟改版就是靠这个才看出
        # 「有密码框但看不见」，否则日志只有一句"未登录"，根本没法排查。
        _why = {
            "OFFSITE": "当前标签页不在孚盟网站上",
            "PWD": "页面上有看得见的密码框（这是登录页的特征）",
            "LOGINROUTE": "地址栏里有 /login（还没跳进系统）",
            "APP": "主框架已加载，人在系统里",
            "EMPTY": "页面框架是空的，可能还没加载完（稍后再点一次试试）",
            "ERR": "读不到页面状态（标签页可能被关掉或正在跳转）",
        }.get(st, st)
        self.log("登录状态:", "已登录" if ok else "未登录",
                 "｜判定依据：" + _why + "｜当前地址：" + (self.w.tab.url or "?"))
        return ok

    def _home(self) -> None:
        """确保停在「孚盟邮」邮件模块里。

        2026-10-12 孚盟改版后重写（真机实测）：
          · 老路由 #/main/mail/home/index **已经不存在了**，直接设 hash 页面
            毫无反应（实测改了 hash，app 内容一个字节都没变）→ 新版路由是 #/main/mail；
          · 而且光设 hash 也进不去，必须**点左侧菜单的「孚盟邮」**才真的切模块
            （实测点完 hash 变 #/main/mail，.mail_dir-operat 立刻出现）。
          · 「写邮件」右边的下拉箭头也没了（.mail_dir-operat .el-dropdown 匹配 0 个），
            所以等待判据从"下拉图标出现"放宽到"邮件操作条出现"。
        这里只做导航，模板下拉的读取另见 mx_web.template_names()。
        """
        if not self.w:
            return
        w = self.w
        try:
            in_mail = w.tab.evaluate(
                "(()=>{const u=location.hash||'';"
                "return u.indexOf('/mail')===0 || !!document.querySelector('.mail_dir-operat');})()",
                timeout=15)
        except Exception:
            in_mail = False
        if not in_mail:
            try:
                r = w._js_click_text("孚盟邮")
                if r.startswith("NOT_FOUND"):
                    self.log("[邮件页] 左侧菜单找不到「孚盟邮」，页面可能还没加载完")
            except Exception as e:
                self.log("[邮件页] 点「孚盟邮」菜单失败：", type(e).__name__, e)
        # 2026-09-26 修复：原来这里死等 2.5 秒就往下走。SPA 渲染慢（或刚从
        # #/login 弹回来）时邮件模块根本没挂载 → 后面点啥啥没有、读啥啥为空。
        # 改成轮询等邮件模块的操作条真的出现（最多 20 秒），到位才继续。
        # 判据兼容两版：新版是 .mail_dir-operat，旧版另有 .el-dropdown。
        try:
            w._settle(2.5,
                      "(()=>!!document.querySelector('.mail_dir-operat')"
                      "||!!document.querySelector('.mail_dir-operat .el-dropdown'))()",
                      lambda v: v is True, timeout=20.0)
        except Exception:
            pass

    # ---------------- 读取 ----------------
    def templates(self) -> list[str]:
        self._check()
        self._home()
        names = self.w.template_names()
        self.log("读到模板 " + str(len(names)) + " 个")
        return names

    def subjects(self, template: str = "") -> list[str]:
        """打开一次写信页（带模板），读主题历史。

        2026-10-04 修复「读取历史主题失败：未能进入写信页」：
          ① 先拿当前**真实**模板列表核对一遍。界面下拉里若残留上次读错的
             假模板名（实测抓到过「全部账号」），直接拿去点必然跳不进去 ——
             这种情况改为不带模板进写信页，先把主题历史读出来。
          ② 读历史主题是**只读**用途，允许最后裸跳写信页兜底（allow_blank=True）。
             发送流程绝不传这个参数，免得发出空白邮件。
        """
        self._check()
        self._home()
        tpl = template or ""
        if tpl:
            try:
                real = self.w.template_names()
                if real and tpl not in real:
                    self.log(f"（历史主题）模板「{tpl}」不在当前模板列表里，"
                             f"改成不带模板进写信页｜可用: {real}")
                    tpl = ""
            except Exception:
                pass
        self.w.open_compose(tpl, allow_blank=True)
        items = self.w.subject_history()
        self.log("读到历史主题 " + str(len(items)) + " 条")
        return items

    def _check(self, need_login: bool = False):
        """每次操作前先探一下连接：断网/浏览器重启后自动重连，登录失效也给明确提示。"""
        if not self.w:
            raise MXWebError("还没连接孚盟，请先点「连接孚盟」")
        href = self._probe()
        if not href:
            self.log("（与孚盟的连接已断开，正在自动重连 …）")
            try:
                self.connect(wait_login=False)
            except Exception as e:
                # 重连本身失败：很可能是页面被冻结/卡死，先刷新自救一次再判失败
                self.log("（重连未成功，尝试刷新页面 …）")
                try:
                    if self.w is not None and self.w.revive("重连失败"):
                        href = self._probe()
                except Exception:
                    pass
                if not href:
                    raise MXWebError(
                        f"自动重连失败：{type(e).__name__}: {e}\n"
                        "已尝试自动刷新页面仍未恢复。\n"
                        "请确认 Edge 还开着并停在孚盟页面，或点「连接孚盟」手动重连。")
            if not href:
                href = self._probe()
            if not href:
                # 连上了浏览器但页面仍无响应 = 页面僵死，只能靠刷新救（用户平时按 F5 的场景）
                self.log("（重连后页面仍无响应，尝试刷新 …）")
                try:
                    if self.w is not None and self.w.revive("重连后页面无响应"):
                        href = self._probe()
                except Exception as e:
                    self.log("（刷新自救失败）" + str(e))
            if not href:
                raise MXWebError(
                    "自动重连 + 刷新后仍连不上孚盟页面"
                    "（请检查 Edge 是否打开、网络是否恢复）")
            self.log("已自动重新连上孚盟 ✓")
        if "/login" in href:
            raise MXWebError("孚盟登录已失效（页面停在登录页），请在 Edge 里重新登录，"
                             "再回本窗口点「检测一下」")
        if need_login:
            try:
                if not self.w.is_logged_in():
                    raise MXWebError("未检测到孚盟登录状态，请在 Edge 里登录后点「检测一下」")
            except MXWebError:
                raise
            except Exception:
                pass

    def _probe(self) -> str:
        """探一次活，返回当前页面地址；连不通返回空串（不抛异常）。"""
        try:
            return str(self.w.tab.evaluate("location.href", timeout=6) or "")
        except Exception:
            return ""

    # ---------------- 单封 ----------------
    def send_one(self, email: str, subject: str = "", template: str = "",
                 name: str = "") -> dict:
        self._check()
        # 不管外面喂进来的是不是"张三 <a@b.com>"、全角＠、中文分号，一律先洗干净
        got = normalize_emails(email)
        if not got:
            res = {"email": str(email or ""), "ok": False,
                   "msg": "邮箱地址不合法（识别不出邮箱，原文：" + repr(email)[:60] + "）",
                   "time": ""}
            return res
        email = got[0]
        res = {"email": email, "ok": False, "msg": "", "time": ""}
        if len(got) > 1:
            self.log(f"  这一个格子里有 {len(got)} 个邮箱，本次只发第一个：{email}，其余：{got[1:]}")
        if self.cfg.skip_sent and self.already_sent(email):
            res["msg"] = "已发过，跳过"
            res["ok"] = True
            res["skipped"] = True
            return res
        try:
            self._home()
            self.w.open_compose(template)
            if getattr(self.cfg, "clear_cc", False):
                # 同上：模板自带抄送 = 每封实际投 2 个地址，风控压力翻倍
                n = self.w.clear_cc()
                if n:
                    self.log(f"  清除抄送芯片 {n} 个（③页「发送前清掉抄送」已开）")
            if subject:
                self.w.set_subject(subject)
                time.sleep(1)
            chip = self.w.add_recipient(email)
            if not chip:
                # 2026-09-21 加固：文本停在输入框里没成芯片时，点「发送」孚盟不认账，
                # 邮件就"卡在最后一步"。补一刀回车提交+失焦再看。
                self.log("  收件人未成芯片，补一次回车提交 ...")
                chip = self.w.commit_recipient(email)
                self.log("  重提后芯片:", chip)

            # 称呼个性化（2026-10-17）：把模板开头的通用称呼换成客户姓名。
            # 失败只记一行日志，不影响后面的自检与发送。
            try:
                _g = self.w.personalize_greeting(name)
                if _g and ("已替换" in _g or "已插入" in _g):
                    self.log("  " + _g)
            except Exception as _e:
                self.log("  称呼个性化跳过：%s" % type(_e).__name__)

            info = self.w.snapshot()
            in_box = any(email in (i or "") for i in self._input_values())
            gate = {
                "收件人": chip or in_box,
                "主题": bool((info.get("subject") or "").strip()),
                "正文": int(info.get("bodyLen") or 0) > 20,
            }
            self.log(f"  自检 {gate}" + ("" if chip else "（注意：收件人未成芯片，风险高）"))
            if not all(gate.values()):
                png = self.w.capture_fail(email + "_gate")
                res["msg"] = f"自检未通过 {gate}，现场已存 {png}"
                return res
            if self.cfg.dry_run:
                res.update({"ok": True, "msg": "演练（未真正发送）", "dry": True})
                return res

            r = self.w._js_click_text("发送", restrict_xy=False)
            if r.startswith("NOT_FOUND"):
                res["msg"] = "找不到发送按钮"
                return res
            # 轮询最多 ~30 秒：确认弹窗、toast、是否离开写信页
            toasts: list[str] = []
            ok = False
            for rnd in range(10):
                if self._stop:
                    # 用户点了停止：发送指令可能已经点出去，提示去孚盟里核对
                    res["msg"] = "已手动停止（发送键可能已点下，请到孚盟已发件箱核对这封）"
                    return res
                time.sleep(2.5)
                try:
                    toasts += self.w._confirm_dialog(rounds=2)
                except Exception as e:
                    self.log("  （读弹窗出错，继续）", type(e).__name__)
                url = self.w.tab.url or ""
                if "/mail/new" not in url:
                    ok = True
                    self.log(f"  第 {rnd + 1} 轮：已离开写信页")
                    break
                if any(any(k in t for k in ("发送成功", "已发送", "发送中")) for t in toasts):
                    ok = True
                    break
            if not ok:
                # 终审：直接去「已发件箱」核对，不靠提示条
                try:
                    v = self.w.verify_sent(email, subject or "")
                    self.log("  已发件箱核对:", v.get("found"), v.get("top", [])[:2])
                    ok = bool(v.get("found"))
                    if not ok:
                        png = self.w.capture_fail(email + "_nosend")
                        self.log("  发送未生效，现场已存:", png)
                except Exception as e:
                    self.log("  已发件箱核对失败:", type(e).__name__, e)
            res["ok"] = bool(ok)
            res["msg"] = "发送成功" if ok else f"未见成功提示 {toasts[-3:]}"
            if not chip and ok:
                res["msg"] += "（收件人是文本态发出的，建议核对该地址）"
            if ok:
                self.sent[email.lower()] = datetime.now(CN).strftime("%Y-%m-%d")
                self._save_sent()
        except Exception as e:
            try:
                png = self.w.capture_fail(email + "_exc") if self.w else ""
                if png:
                    self.log("  异常现场已存:", png)
            except Exception:
                pass
            res["msg"] = f"{type(e).__name__}: {e}"
        res["time"] = datetime.now(CN).strftime("%Y-%m-%d %H:%M:%S")
        return res

    def send_scheduled_mx(self, email: str, subject: str, template: str,
                          when_local: str, country: str,
                          verify: bool = True, name: str = "") -> dict:
        """串行处理单个客户的完整定时流程（一次只处理这一个客户）：

        ①打开模板写信页 → ②填主题/收件人 → ③定时弹窗选客户国家（时区孚盟带出）
        → ④填当地时间（如明早 9:00）→ ⑤确定 → ⑥等弹窗关闭
        → ⑦去「发件箱」核对定时记录是否出现（成功凭据）
        → ⑧回到孚盟初始页面并等真的回去（下一客户才开场，避免标签页堆积）。

        when_local: 收件人当地时间 "YYYY/MM/DD HH:MM"。缺国家 → 直接跳过并记录。
        成功后记入 sent_history（带 sched 标记），防止同一客户被重复排队。
        """
        self._check()
        # 页面僵死/白屏自查：不健康就先自动刷新，别等它自己失败
        # （长时间挂机时 Edge 会冻结标签页、或断网后页面不重拉数据 → 必须 F5）
        try:
            if getattr(self.cfg, "auto_revive", True) and not self.w.ensure_healthy("发送前体检"):
                return {"ok": False, "msg": "页面无响应且刷新后未恢复（请手工看一眼 Edge）"}
        except MXWebError:
            raise
        except Exception:
            pass
        t0 = time.time()
        country = fix_country(country)      # 简短国家名补全成孚盟能搜到的写法

        # 2026-09-29：客户表「国家」列常写成「国家名 + 后缀」（美国广交会客户、
        # 沙特sa、（阿联酋）迪拜、德国\n开发客户），也可能写的是地区泛称（欧盟/中东）。
        # 前者在这里清洗掉后缀，后者直接跳过 —— 别让孚盟白搜十几秒再报错。
        _amb = False
        try:
            from .geo_alias import resolve_country
            _std, _note = resolve_country(country)
            if _note == "amb":
                _amb = True
            elif _std and _std != country:
                self.log(f"[国家] 「{country}」→ 识别为「{_std}」（客户表里带了别的内容）")
                country = _std
        except Exception:
            pass
        got = normalize_emails(email)
        if not got:
            self.log(f"[定时] ✗ {email} 邮箱地址不合法，跳过（用时 0s）")
            return {"ok": False, "msg": "邮箱地址不合法：" + repr(email)[:50]}
        email = got[0]
        res = {"email": email, "ok": False, "msg": "", "time": ""}

        # 缺国家/时区信息：不排队、记日志（排队环节已过滤，这里是兜底）
        if not (country or "").strip():
            self.log(f"[定时] ✗ {email} 缺国家信息，无法选时区 → 跳过并记录（用时 0s）")
            res["msg"] = "缺国家信息，已跳过（可在客户表补国家后重排）"
            res["time"] = datetime.now(CN).strftime("%Y-%m-%d %H:%M:%S")
            return res

        # 国家写的是地区泛称（欧盟/中东/非洲/东南亚…）：孚盟一定搜不到，别白跑
        if _amb:
            self.log(f"[定时] ✗ {email} 国家写的是地区泛称「{country}」，"
                     f"孚盟只认具体国家 → 跳过（用时 0s）")
            res["msg"] = f"国家是地区泛称「{country}」，请在客户表改成具体国家名后重排"
            res["time"] = datetime.now(CN).strftime("%Y-%m-%d %H:%M:%S")
            return res

        # 2026-10-05：用户选了「按北京时间定」时，把北京时间换算成该国当地时间。
        cn_when = when_local
        if str(getattr(self.cfg, "when_basis", "local") or "local") == "beijing":
            try:
                from .tzmap import cn_to_local as _c2l, parse as _tzp, fmt as _tzf
                _t = _tzp(when_local)
                if _t is not None:
                    _loc = _c2l(_t, country)
                    if _loc is not None:
                        when_local = _tzf(_loc)
                        self.log(f"  🕐 北京 {cn_when} → {country} 当地 {when_local}")
                    else:
                        self.log(f"  ⚠ 「{country}」不在时区表里，按原值 {when_local} 写入"
                                 f"（请核对客户表国家名）")
            except Exception as e:
                self.log(f"  ⚠ 时区换算失败（{type(e).__name__}），按原值 {when_local} 写入")

        verified = "未核对"
        back = False
        try:
            # ①-⑤ 完整写信 + 定时弹窗
            info = self.w.schedule_one(email, subject, template,
                                       country=country, when_local=when_local,
                                       require_chip=self.cfg.require_chip,
                                       clear_cc=bool(getattr(self.cfg, "clear_cc", False)),
                                       recipient_name=name)
            eff = info.get("when_local") or when_local   # 弹窗里可能顺延过的时间
            if eff != when_local:
                self.log(f"  ⏱ 该客户实际发送时间已顺延：当地 {eff}（原设定 {when_local}）")
            res["when"] = eff
            res["cn_when"] = cn_when
            # ⑥ 弹窗已关。先清掉可能堆积的多余标签，再等回初始页面
            try:
                closed = self.w.close_other_tabs()
                if closed:
                    self.log(f"  清理多余标签 {closed} 个")
            except Exception:
                pass
            back = self.w.back_to_home()

            # ⑦ 去「发件箱」核对定时记录（按核对频率决定这封核不核对）
            if verify:
                for attempt in (1, 2):
                    try:
                        v = self.w.verify_sent(email, subject or "", folder="发件箱")
                        if v.get("found"):
                            verified = f"发件箱可见（{v['top'][0][:40]}…）" if v.get("top") else "发件箱可见"
                            break
                        verified = "发件箱未见该记录"
                        if attempt == 1:
                            time.sleep(4)
                    except Exception as e:
                        verified = f"核对出错 {type(e).__name__}"
                        if attempt == 1:
                            time.sleep(4)
            else:
                verified = "本封跳过核对（按核对频率）"

            res["ok"] = True
            if verified.startswith("发件箱可见"):
                res["msg"] = f"已定时（{country} 当地 {eff}，发件箱已核对）"
            elif verified.startswith("本封跳过"):
                res["msg"] = f"已定时（{country} 当地 {eff}，本封未核对，整轮结束后统一核对）"
            else:
                # 弹窗流程成功但列表没核对到：不再重发（防重复定时），标记请人工确认
                res["msg"] = (f"已提交定时（{country} 当地 {eff}），"
                              f"但发件箱{verified}，请稍后人工确认，程序不会重发")
            # scheduled 邮件不算"已发"，但必须挡住重复排队
            self.sent[email.lower()] = datetime.now(CN).strftime("%Y-%m-%d") + "(sched)"
            self._save_sent()
        except Exception as e:
            try:
                png = self.w.capture_fail(email + "_sched")
                self.log("  定时失败现场已存:", png)
            except Exception:
                pass
            res["msg"] = f"{type(e).__name__}: {e}"
            # 失败也尽量把页面拉回初始状态，别让下一位客户在脏页面上开场
            try:
                self.w.back_to_home()
            except Exception:
                pass
        sec = round(time.time() - t0, 1)
        res["sec"] = sec
        res["time"] = datetime.now(CN).strftime("%Y-%m-%d %H:%M:%S")
        self.log(f"[定时] {'✓' if res['ok'] else '✗'} {email}（{country}）"
                 f"{'，' + res['msg'] if res['msg'] else ''} | 用时 {sec}s | "
                 f"回到初始页={'是' if back else '否/超时'} | 发件箱核对={verified} | 标签数={self.w.open_tab_count()}")
        return res

    def mark_scheduled(self, email: str) -> None:
        """把地址记为「已定时」（发件箱里确认有记录时用），防下次重发。"""
        e = (email or "").strip().lower()
        if not e:
            return
        self.sent[e] = datetime.now(CN).strftime("%Y-%m-%d") + "(sched)"
        self._save_sent()

    def preflight_scheduled(self, emails: list[str], max_scroll: int = 25,
                            subject: str = "") -> dict:
        """开跑前防重复闸门：打开一次「发件箱」，扫一遍现有定时记录，
        看本轮客户里哪些在孚盟已经有定时任务了（这些**绝不能**再发）。

        背景：2026-09-22 事故——之前被误判"失败"的客户其实定时成功了，
        但没进已发记录 → 下轮重发 → 同一客户定时两三遍。

        2026-09-26 加 **subject**：
          平时不传 → 和以前一样，只看邮箱前缀（最保守）。
          新一轮（用户要的"第二轮"）才传本轮主题 → 发件箱里那条必须同时
          出现「收件人前缀」+「本轮主题前 12 字」才算重复。
          上一轮留下的旧记录主题不同 → 不算重复 → 第二轮能正常发。
        """
        self._check()
        emails = [e for e in (emails or []) if (e or "").strip()]
        if not emails:
            return {"found": [], "n_items": 0}
        try:
            items = self.w.scan_sent(subject="", folder="发件箱", max_scroll=max_scroll)
            if not items:               # 页面可能还没渲染完，再给一次机会
                time.sleep(2.5)
                items = self.w.scan_sent(subject="", folder="发件箱", max_scroll=max_scroll)
        except Exception as e:
            self.log(f"[防重复] 扫描发件箱失败：{type(e).__name__}: {e}（跳过预检，继续跑）")
            return {"found": [], "n_items": 0}
        if not items:
            # 2026-09-23 实测：预检读到 0 条不代表发件箱是空的（列表渲染慢 /
            # 孚盟列表一次只显示有限条）。这时预检无效，但本地已发记录仍然挡得住。
            self.log("[防重复] 本次没读到发件箱列表，预检无效（不影响发送；"
                     "防重仍以本地已发记录为准）")
            return {"found": [], "n_items": 0}
        low = [(t or "").lower() for t in items]
        key = (subject or "").strip()[:12].lower()
        other_round = 0            # 有这个人的记录、但主题不是本轮的（上一轮留下的）
        found = []
        for e in emails:
            local = e.split("@")[0].strip().lower()
            if not local:
                continue
            hit = False
            for t in low:
                if local in t:
                    if not key or key in t:
                        hit = True
                        break
                    other_round += 1
            if hit:
                found.append(e)
        self.log(f"[防重复] 发件箱现有定时记录 {len(items)} 条；"
                 f"本轮客户中 {len(found)} 位已有记录"
                 + (f"（另有 {other_round} 处是**别的主题**的旧记录，"
                    f"按新一轮处理、不算重复）" if other_round else "")
                 + (f"：{'、'.join(x.split('@')[0] for x in found[:6])}{'…' if len(found) > 6 else ''}"
                    if found else ""))
        return {"found": found, "n_items": len(items), "other_round": other_round}

    def verify_batch_scheduled(self, emails: list[str], folder: str = "发件箱") -> dict:
        """整轮跑完统一核对：只打开发件箱一次，比对本轮全部客户。

        省时点：原来每位客户都要切一次发件箱（约 5~6 秒/位），400 多位就是半小时。
        返回 {"found": [...], "missing": [...], "n_items": N}。
        """
        self._check()
        emails = [e for e in (emails or []) if e]
        if not emails:
            return {"found": [], "missing": [], "n_items": 0}
        try:
            r = self.w.verify_batch(emails, folder=folder)
        except Exception as e:
            self.log(f"[核对] 统一核对失败：{type(e).__name__}: {e}")
            return {"found": [], "missing": emails, "n_items": 0,
                    "err": f"{type(e).__name__}: {e}"}
        tail = "，全部核对到 ✓" if not r.get("missing") else f"，未核对到 {len(r['missing'])} 位"
        self.log(f"[核对] 「{folder}」共 {r['n_items']} 条记录；本轮 {len(emails)} 位"
                 f"核对到 {len(r['found'])} 位{tail}")
        if r.get("missing"):
            self.log("[核对] 未核对到的（请到孚盟发件箱人工看一眼）："
                     + "、".join(r["missing"][:12])
                     + ("…" if len(r["missing"]) > 12 else ""))
        return r

    def confirm_sent_from_outbox(self, rows: list[dict], subject: str = "") -> dict:
        """以「已发件箱」的真实记录为准，核对哪些客户真的发出去了。

        返回 {"texts": 已发件箱条目, "rows": [{...row, confirmed, item}]}。
        匹配规则：条目文本包含 邮箱@前缀 +（若有）主题前 24 字。
        同名前缀（@前相同）的客户会提示人工注意，避免张冠李戴。
        """
        self._check()
        texts = self.w.scan_sent(subject=subject)
        low = [t.lower() for t in texts]
        out = []
        for r in rows:
            email = (r.get("email") or "").strip().lower()
            local = email.split("@")[0]
            hit = ""
            if email:
                for t in low:
                    if local and local in t and (not subject or subject[:24].lower() in t):
                        hit = t
                        break
            out.append({**r, "confirmed": bool(hit), "item": hit[:110]})
        twins = {}
        for r in rows:
            local = (r.get("email") or "").split("@")[0].lower()
            twins[local] = twins.get(local, 0) + 1
        dup = [k for k, v in twins.items() if v > 1]
        res = {"texts": texts, "rows": out, "same_prefix": dup}
        n_ok = sum(1 for r in out if r["confirmed"])
        self.log(f"[核对] 已发件箱共读到 {len(texts)} 条；"
                 f"确认发出 {n_ok}/{len(out)}" + (f"；@前缀重复 {dup}" if dup else ""))
        return res

    def _input_values(self) -> list[str]:
        try:
            raw = self.w.tab.evaluate(
                "JSON.stringify(Array.from(document.querySelectorAll('input.input-box__item'))"
                ".map(e=>e.value||''))", timeout=15)
            return json.loads(raw)
        except Exception:
            return []

    # ---------------- 批处理 ----------------
    def stop(self):
        self._stop = True
        try:
            self.w.abort = True        # 浏览器层面的等待也会立刻中断
        except Exception:
            pass

    def resume(self):
        """继续（点「继续」或新起一轮时调用）。"""
        self._stop = False
        try:
            self.w.abort = False
            self.w.pause = False
        except Exception:
            pass

    def pause(self, on: bool = True):
        """暂停：浏览器操作停在当前这一步，点继续再走（不中断这一位客户）。"""
        try:
            self.w.pause = bool(on)
        except Exception:
            pass

    def run_batch(self, rows: list[dict], subject: str, template: str,
                  on_progress=None, write_col: str = "") -> dict:
        """rows: [{email, name, country, row_index}]"""
        self._stop = False
        total = len(rows)
        done = ok = fail = skip = 0
        results = []
        for i, r in enumerate(rows, 1):
            if self._stop:
                self.log("已手动停止")
                break
            raw = (r.get("email") or "")
            got = normalize_emails(raw)
            email = got[0] if got else ""
            if got and len(got) > 1 and got[1:]:
                # 一个格子里写了好几个邮箱：后面几个按独立一行排队，本轮先发第一个
                self.log(f"  第 {r.get('row_index')} 行有 {len(got)} 个邮箱，"
                         f"本次发 {got[0]}，其余请拆成单独一行：{got[1:]}")
            label = f"[{i}/{total}] {r.get('name', '')} <{email}>"
            if not email:
                skip += 1
                results.append({**r, "ok": False, "msg": "无邮箱"})
                self.log(label + " 跳过：无邮箱")
                continue
            if self.cfg.only_work_hours:
                self._wait_work_hours()

            self.log(label + " 开始")
            res = self.send_one(email, subject, template, name=(r.get("name") or ""))
            results.append({**r, **res})
            done += 1
            if res.get("skipped"):
                skip += 1
            elif res.get("ok"):
                ok += 1
            else:
                fail += 1
            self.log(f"   -> {res.get('msg')}")
            if on_progress:
                on_progress(i, total, res)
            if i < total and not self._stop:
                gap = self._inter_send_gap()
                self.log(f"   等待 {gap:.0f} 秒（约 {gap / 60:.1f} 分钟）再发下一位"
                         f"{'（随机间隔）' if getattr(self.cfg, 'human_mode', True) else ''}")
                self._sleep(gap)
        return {"total": total, "done": done, "ok": ok, "fail": fail,
                "skip": skip, "results": results}

    # ---------------- 随机间隔（2026-09-30 新增，防限流）----------------
    # 背景：09-29 实测一轮 114 封、封间中位只有 17 秒、90 个间隔不到 30 秒，
    # 每封还带 1.92M 附件 → 孚盟发送服务器 Connection reset，
    # 发件箱里大量「发送失败」。真人不可能 17 秒发一封带附件的开发信。
    # 随机三件套（开启后**不看**设置页的「封间等待」，那只在关闭随机时生效）：
    #   ① 封间基础间隔 45–110 秒随机；
    #   ② 连发 5–8 封（每次随机）后「离开一会儿」：休息 3–7 分钟；
    #   ③ 每封之间有 8% 概率额外多停 30–90 秒。
    _HUMAN_GAP = (45.0, 110.0)
    _HUMAN_BURST = (5, 8)          # 连发多少封算一拨（随机取）
    _HUMAN_REST = (180.0, 420.0)   # 一拨之后的休息时长（秒）
    _HUMAN_WANDER_P = 0.08         # 额外停顿概率
    _HUMAN_WANDER = (30.0, 90.0)   # 额外停留区间（秒）

    def _inter_send_gap(self) -> float:
        """这一封发完后该等多久（秒）。human_mode=False 时退回设置里的封间等待。"""
        if not getattr(self.cfg, "human_mode", True):
            return random.uniform(self.cfg.interval_min, self.cfg.interval_max)
        self._burst_left = getattr(self, "_burst_left", 0) or random.randint(*self._HUMAN_BURST)
        gap = random.uniform(*self._HUMAN_GAP)
        if random.random() < self._HUMAN_WANDER_P:
            w = random.uniform(*self._HUMAN_WANDER)
            self.log(f"   （随机：这封之后多歇 {w:.0f} 秒）")
            gap += w
        self._burst_left -= 1
        if self._burst_left <= 0:
            rest = random.uniform(*self._HUMAN_REST)
            self._burst_left = random.randint(*self._HUMAN_BURST)
            self.log(f"   （随机：已连发一拨，离开 {rest / 60:.1f} 分钟再回来，"
                     f"之后连发 {self._burst_left} 封左右再休息）")
            return rest
        return gap

    def _sleep(self, sec: float):
        end = time.time() + sec
        while time.time() < end:
            if self._stop:
                return
            time.sleep(1)

    def _wait_work_hours(self):
        while not self._stop:
            h = datetime.now(CN).hour
            if self.cfg.work_start <= h < self.cfg.work_end:
                return
            self.log(f"当前 {h} 点不在发送时段（{self.cfg.work_start}-{self.cfg.work_end} 点），等待中 ...")
            self._sleep(600)
