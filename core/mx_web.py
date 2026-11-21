# -*- coding: utf-8 -*-
"""
孚盟MX 网页端驱动层（WebMX）。

本模块里的每一个选择器、每一步流程，都是 2026-09-19/20 在真实孚盟MX 网页端
（https://www.fumamx.com/，Edge 153）上逐条实测验证过的，不是猜的。

标准发信流程：
  1. 启动/复用本机 Edge（独立 profile；注意登录态不跨重启，见 ensure_login）
  2. 左侧「孚盟邮」→「写邮件」右边的小下拉图标 → 选模板（正文+附件由 MX 端模板带入）
  3. 清理（可选）：删抄送芯片 / 删附件
  4. 主题：可直接写，也可取「历史记录」第 N 条
  5. 收件人：input 里插入地址 → 回车 → 失焦成芯片
  6. 点「发送」→ 处理确认弹窗 → 校验「发送成功」提示
  7. 去「已发件箱」复核（最靠得住的成功判定）

2026-09-20 真机实测结论（写死在代码里，别再踩）：
  - 电脑锁屏 / 窗口被完全遮挡时 document.hidden=true，Vue 的展开动画（下拉、
    弹窗、toast）永远不跑。此时**不要按可见性过滤元素**，直接对隐藏的 <li>
    派发 click 同样能触发 Vue 命令；发送是否成功改判"是否离开 /mail/new"。
  - 页面上常驻大量 w=0/h=0 的隐藏 .el-dialog（菜单栏显示设置、添加菜单、
    选择图片、定时发送……）。弹窗识别**必须按实际尺寸过滤**，否则会误点
    其中的「确定」，把用户的菜单设置恢复成默认。
  - 主题历史必须把主题**清空**才弹；容器 .recent-subject，末尾 <p> 是"到底了"。
  - 库外地址（如 QQ）失焦后会生成「陌生」前缀的芯片，可以正常发送。
  - 模板下拉第一组即「2610广交会邀请邮件模板」，templateId=18714953。

设计红线（务必遵守）：
  - 正文完全由 MX 端模板生成，本地绝不拼 HTML。
  - 不注入任何 X-Mailer / List-Unsubscribe 之类的自动化邮件头。
  - 账号密码只由用户在浏览器窗口里手工输入，脚本不接触、不落盘、不打印。
"""
from __future__ import annotations

import json
import random
import re
import time
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import cdp


class MXWebError(RuntimeError):
    pass


# --------------------------------------------------------------------------
# 选择器 / 文案（孚盟改版时优先改这里）
# --------------------------------------------------------------------------
SEL = {
    # 导航
    "login_password_input": "input[type=password]",
    "menu_mail": "孚盟邮",
    # 写邮件按钮右边的下拉图标（选模板的正确入口，实测确认）
    "compose_dropdown": ".mail_dir-operat .el-dropdown",
    "compose_dropdown_icon": ".mail_dir-operat .el-dropdown i.iconfont",
    "template_item": ".el-dropdown-menu li",
    "compose_new": "写邮件",
    # 写信页
    "sender_input": "input.el-input__inner[placeholder=请选择发件人]",
    "recipient_input": "input.input-box__item",
    "subject": ".mail-info__item_subject",
    "subject_history": ".recent-subject",
    "body_iframe": "iframe.tox-edit-area__iframe",
    "btn_send": "发送",
    "btn_send_span": "span.mail_sendBtn",   # 真正的发送按钮（内含 span.box1）
    "folder_sent": "已发件箱",                # 注意不是「已发送」
    "chip": ".address-item",
    "chip_delete": "address-item__delete",
    "attach_item": ".FileItem",
    "attach_delete": "i.icon-del",
}

CONFIRM_WORDS = ["确定", "确 定", "确认", "继续发送", "发 送", "发送"]
SUCCESS_WORDS = ["发送成功", "已发送", "发送中"]

# --------------------------------------------------------------------------
# 极速模式的「到位判断」：页面真的准备好了就立刻走，不再死等固定秒数。
# 每条都返回可直接判真假的值（布尔 / 数字 / 字符串）。
# --------------------------------------------------------------------------
_DLG_PICK = ("const ok=e=>{const r=e.getBoundingClientRect();"
             " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
             "const W=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
             ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
             ".trim()==='定时发送');")
JS_SCHED_OPEN = "(()=>{" + _DLG_PICK + "return !!W;})()"
JS_COMPOSE_READY = (
    "(()=>{const ins=Array.from(document.querySelectorAll('input.input-box__item'))"
    ".filter(e=>{const r=e.getBoundingClientRect();return r.width>50&&r.y>0;});"
    "return ins.length>0;})()")
JS_TPL_COUNT = (
    "(()=>Array.from(document.querySelectorAll('ul.el-dropdown-menu li'))"
    ".filter(e=>(e.innerText||'').trim()).length)()")

# 2026-10-04：模板下拉扫描器（重写版）
# 老版本只扫 ul.el-dropdown-menu，且只按「不含动作词」筛 → 实测把孚盟的
# **账号切换下拉**当成模板列表抓了回来（读到 1 个「全部账号」），
# 接着 open_compose 去点「全部账号」这个 li，自然跳不进写信页 →
# 「读取历史主题失败：未能进入写信页」。模板和历史主题两个功能一起废掉。
# 这一版改三点：多容器扫描 / 显式排除账号·信箱类分组 / 输出诊断信息。
JS_TPL_SCAN = (
    "(()=>{const txt=e=>((e.innerText||'').trim()||(e.textContent||'').trim());"
    # 顺带把「写邮件」下拉图标的位置一起带回去：
    # 模板菜单是挂在这个图标**正下方、左对齐**的 popper（实测 ul.left == 图标.left，
    # li 中心 x=305.5 宽 219 → left=196，图标 left=196 完全吻合）。
    # 页面上 300+ 个下拉里，只有它是这个几何关系 —— 比"取最大组"可靠得多。
    "const dd=document.querySelector('" + SEL["compose_dropdown"] + "');"
    "const ic=dd&&(dd.querySelector('i.iconfont')||dd);"
    "let icon=null;"
    "if(ic){const b=ic.getBoundingClientRect();"
    " if(b.width>0&&b.height>0) icon={x:Math.round(b.x),y:Math.round(b.y),"
    " w:Math.round(b.width),h:Math.round(b.height)};}"
    # 动作词：出现在这些词里的 li 一定不是模板名。
    # 2026-10-04 真机补充：页面上有大量右键菜单（SingleOperation_dropdown 10 项：
    # 标记为已读/批注/作为附件发送/导出eml文件/内分发…），不扩这张表就会被它们赢走。
    "const BAD=/(回复|转发|删除|标记|批注|文件夹|审批|归并|已读|未读|排序|关闭|移入|移出|导出|"
    "新增|设置|打印|举报|移动|清空|导入|处理|待办|本人|撤回|查看|打开|授权|分配|"
    "转交|共享|同步|刷新|展开|收起|全选|反选|置顶|上移|下移|预览|下载|另存|"
    "内分发|日志|转存|复制|新标签页|作为附件|附件|签名|草稿|邮件夹|星标|抄送|密送)/;"
    # 账号 / 信箱 / 时间筛选类下拉：这些**绝不可能是模板**
    "const ACC=/(全部账号|所有账号|选择账号|切换账号|账号|发件人|收件箱|发件箱|"
    "已发件箱|草稿箱|垃圾箱|垃圾邮件|已删除|广告邮件|其他文件夹|全部邮件|"
    "未读邮件|已读|星标邮件|今日|本周|本月|最近|全部|所有)/;"
    # 2026-10-08 关键补充：**有些模板名本身含有动作词**，会被 BAD 误杀。
    # 真机实例：孚盟里有个模板就叫「未回复」（templateId 18873624），
    # "回复"二字命中 BAD → 脚本只读到 7 个模板，漏了它。
    # 判据修正：这个菜单里的 li 如果**自带 Vue command（templateId + type）**，
    # 那就是孚盟自己的模板对象 —— 铁证，比名字像不像模板可靠得多。
    # 所以：有 templateId 的一律保留；没有的才退回名字过滤。
    "const sels=['ul.el-dropdown-menu','div.el-dropdown-menu','ul[role=menu]',"
    "'.el-popper ul','ul.el-select-dropdown__list','div.el-poper ul'];"
    "const seen=new Set();const out=[];"
    "for(const s of sels){for(const u of document.querySelectorAll(s)){"
    " if(seen.has(u))continue;seen.add(u);"
    " const lis=Array.from(u.querySelectorAll('li'));"
    " const raw=lis.map(txt).filter(Boolean);"
    " if(!raw.length)continue;"
    " const ok=lis.filter(li=>{const t=txt(li);"
    "   if(!t||t.length>60) return false;"
    "   const c=li.__vue__&&li.__vue__.command;"
    "   if(c&&typeof c==='object'&&c.templateId) return true;"   # ★ 有模板 ID → 铁定是模板
    "   return !BAD.test(t)&&!ACC.test(t);"                     # 否则退回名字过滤
    "  }).map(txt);"
    " const r=u.getBoundingClientRect();"
    " const nid=lis.filter(li=>{const c=li.__vue__&&li.__vue__.command;"
    "   return c&&typeof c==='object'&&c.templateId;}).length;"
    " out.push({n:ok.length,raw:raw.length,nid:nid,"
    "  bad:raw.filter(t=>BAD.test(t)).length,"
    "  acc:raw.filter(t=>ACC.test(t)).length,"
    "  items:ok.slice(0,80),cls:(u.className||'').toString().slice(0,50),"
    "  rect:{x:Math.round(r.x),y:Math.round(r.y),w:Math.round(r.width),h:Math.round(r.height)},"
    "  popper:u.hasAttribute('x-placement')?1:0});}}"
    "return JSON.stringify({icon:icon,groups:out});})()")
# 精确版：写邮件下拉图标 aria-owns 指向的那个菜单（Element UI 标准，最准）
JS_TPL_OWNS = (
    "(()=>{const txt=e=>((e.innerText||'').trim()||(e.textContent||'').trim());"
    "const dd=document.querySelector('" + SEL["compose_dropdown"] + "');"
    "if(!dd) return '[]';"
    "const ids=[dd.getAttribute('aria-owns'),dd.getAttribute('aria-controls')];"
    "const seen=new Set();const out=[];"
    "for(const id of ids){"
    " const ul=id&&document.getElementById(id);"
    " if(!ul||seen.has(ul))continue;seen.add(ul);"
    " const g=Array.from(ul.querySelectorAll('li')).map(txt).filter(Boolean);"
    " if(g.length)out.push({n:g.length,raw:g.length,bad:0,acc:0,items:g.slice(0,80),cls:'aria-owns'});}"
    "return JSON.stringify(out);})()")
JS_SEARCHBOX = (
    "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
    " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
    "return Array.from(document.querySelectorAll('.searchBox input')).filter(ok).length;})()")
JS_COUNTRY_VAL = (
    "(()=>{" + _DLG_PICK + "if(!W) return '';"
    "const it=Array.from(W.querySelectorAll('.el-form-item'))"
    ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='国家地区');"
    "return it? (it.querySelector('.el-select input.el-input__inner').value||'') : '';})()")
JS_TIME_VAL = (
    "(()=>{" + _DLG_PICK + "if(!W) return '';"
    "const it=Array.from(W.querySelectorAll('.el-form-item'))"
    ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='当地时间');"
    "return it? (it.querySelector('input').value||'') : '';})()")


JS_MAIL_ITEMS = (
    "(()=>Array.from(document.querySelectorAll('.mail-item'))"
    ".filter(e=>e.getBoundingClientRect().width>100).length)()")


def _js_country_hit(country: str) -> str:
    """国家下拉里是否已经出现可选项（精确或包含）。"""
    return ("((c)=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>5&&r.height>5&&getComputedStyle(e).display!=='none';};"
            "const lis=Array.from(document.querySelectorAll('li.el-select-dropdown__item')).filter(ok);"
            "return lis.some(l=>((l.title||'')+(l.innerText||'')).indexOf(c)>=0);})("
            + json.dumps(country, ensure_ascii=False) + ")")


def _js_recip_ready(email: str) -> str:
    """收件人是否已就位：成了芯片，或文本还在输入框里（库外地址常常只到这一步）。"""
    e = json.dumps(email.lower(), ensure_ascii=False)
    return ("(()=>{const ins=Array.from(document.querySelectorAll('input.input-box__item'))"
            ".filter(x=>x.getBoundingClientRect().width>50);"
            "if(ins.some(x=>(x.value||'').toLowerCase().includes(" + e + "))) return true;"
            "return Array.from(document.querySelectorAll('.address-item'))"
            ".some(x=>(x.textContent||'').toLowerCase().includes(" + e + "));})()")


def _js_chip(email: str) -> str:
    """收件人芯片是否已经生成（库外地址可能一直不生成，靠超时兜底）。"""
    return ("(()=>Array.from(document.querySelectorAll('.address-item'))"
            ".some(x=>(x.textContent||'').toLowerCase().includes("
            + json.dumps(email.lower(), ensure_ascii=False) + ")))()")


def _digits(s) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())


@dataclass
class WebSettings:
    browser: str = ""                       # 空 = 自动找 Edge/Chrome
    port: int = 9333
    profile_dir: str = ""
    login_url: str = "https://www.fumamx.com/"
    headless: bool = False
    login_wait_minutes: int = 20
    step_delay: float = 1.0                 # 每步之间的停顿（等 SPA 渲染）
    # 极速模式：把「固定睡 N 秒」换成「页面到位就立刻走」的轮询等待（2026-09-22）
    # 关掉它就退回原来那种保险的死等，用于孚盟卡顿时排障。
    fast: bool = True
    # 2026-09-23：窗口被最小化时自动恢复。最小化会让鼠标事件全部落空、
    # Page.captureScreenshot 挂死（老坑），所以宁可自己弹回来也不要卡住。
    auto_restore: bool = True
    # 2026-10-05：是否允许把浏览器**拉到前台**。默认**关**。
    # 开着的时候每选一次模板就 bringToFront 一次，用户开着别的窗口干活时会被
    # 不停打断（任务栏图标闪红、自己打开的窗口被盖住/最小化）。
    # 关掉后一律不抢焦点：选模板改走 Vue 组件方法（见 _click_template_li），
    # 不需要展开动画、不需要点坐标 —— 浏览器被别的窗口挡着也照样跑。
    auto_raise: bool = False
    extras: dict[str, Any] = field(default_factory=dict)


# --------------------------------------------------------------------------
class WebMX:
    """网页端发信会话。

        w = WebMX(WebSettings(profile_dir=r"...\\mxprofile"))
        w.ensure_login()
        print(w.subject_history())                       # 看主题历史
        r = w.send_one(to_email="a@b.com",
                       subject=w.subject_history()[0],
                       template_name="2610广交会邀请邮件模板")
    """

    abort: bool = False    # 用户点「停止」：正在进行的浏览器操作尽快中断
    pause: bool = False    # 用户点「暂停」：停在当前这一步等待，点「继续」再走

    def _sleep(self, sec: float) -> None:
        """可中断的等待：停止立刻跳出，暂停就地等待（0.2 秒粒度响应）。"""
        end = time.time() + sec
        while True:
            if self.abort:
                raise MXWebError("已手动停止")
            if self.pause:
                time.sleep(0.2)
                continue
            rest = end - time.time()
            if rest <= 0:
                return
            time.sleep(min(0.2, rest))

    def _human_pause(self, lo: float, hi: float) -> None:
        """随机微停顿：随机 lo–hi 秒。

        2026-09-30 防限流：fast 模式下每步都是固定 0.35/0.5 秒，
        一天几百封的页面操作节奏完全一致，本身就是机器特征。
        关键步骤换成随机区间，让每封邮件的耗时各不相同。
        """
        lo = max(0.2, float(lo))
        hi = max(lo + 0.1, float(hi))
        self._sleep(random.uniform(lo, hi))

    def _wait(self, js: str, want, timeout: float = 8.0, poll: float = 0.25) -> bool:
        """轮询等页面到位就立刻走（极速模式用）。

        原来每步都是「固定睡 N 秒」——孚盟其实常常 0.3 秒就好了，白等的都是整批时间。
        want(v) 成立马上返回 True；超时返回 False（不抛错，交给上层兜底重试）。
        """
        deadline = time.time() + timeout
        while True:
            try:
                v = self.tab.evaluate(js, timeout=8)
            except Exception:
                v = None
            try:
                if want(v):
                    return True
            except Exception:
                pass
            if time.time() >= deadline:
                return False
            self._sleep(poll)

    def _settle(self, slow: float, js: str = "", want=None, timeout: float = 6.0) -> bool:
        """一步等待的统一入口：极速模式轮询到位，否则退回原来的固定秒数。"""
        if self.s.fast and js:
            return self._wait(js, want or (lambda v: bool(v)), timeout=timeout)
        self._sleep(slow)
        return True

    def __init__(self, settings: WebSettings, log=print):
        self.s = settings
        self.browser: cdp.Browser | None = None
        self.tab: cdp.Tab | None = None
        # 2026-09-21 补：_pick_country / _set_local_time 里用 self.log 打日志，
        # 以前没这个属性 → 定时流程选完国家就 AttributeError 崩掉（定时全军覆没的真凶）。
        self.log = log or print

    def _say(self, *a):
        print("[WebMX]", *a, flush=True)

    # ---------------- 让页面真的"活"过来（2026-10-04 真机取证）----------------
    # 实测：助手窗口挡在浏览器前面时 document.hidden = true，
    # Vue 的下拉展开动画**永远不跑** → 菜单项 getBoundingClientRect().height 恒为 0
    # → JS 派发 click 和 CDP 真实鼠标点击**全部落空**（点了没反应）。
    # 只要等到 height > 0 再点，JS 派发一下就跳转（实测 1 秒进写信页）。
    def _bring_front(self):
        """把孚盟标签页带到前台。

        ⚠ 2026-10-05：默认**不允许**（WebSettings.auto_raise=False）。
        每次 bringToFront 都会把 Edge 拽到用户面前 —— 用户正在别的窗口干活时
        表现为"任务栏图标闪红 + 自己打开的窗口被盖住/最小化"，119 封就是抢 119 次。
        现在选模板走 Vue 组件方法，不需要动画和坐标，所以根本不用抢焦点。
        只有用户自己在③页勾了「允许拉前台」才真的拉（排障时用）。
        """
        if not getattr(self.s, "auto_raise", False):
            return False
        try:
            self.tab.send("Page.bringToFront", {}, timeout=8)
        except Exception:
            pass
        try:
            self.ensure_visible()
        except Exception:
            pass
        return True

    # 2026-10-05 真机取证：模板 li 上挂着 Vue2 组件实例，
    # li.__vue__.command 就是 {"templateId":18714953,"type":"W"}，
    # li.__vue__.handleClick() 会走 Element UI 的 menu-item-click → 路由跳转。
    # 关键：**它不依赖展开动画、不依赖 getBoundingClientRect、不依赖焦点**
    # （实测浏览器被挡住、li 高度还是 0 的时候，handleClick 照样 1 秒进写信页，
    # 而 li.click()/真实鼠标点全部落空）。这就把"必须抢前台"这个死结解开了。
    def template_ids(self) -> dict[str, dict]:
        """不展开下拉，直接从常驻 DOM 读 {模板名: {templateId, type}}。

        同一个模板在孚盟里有多份（写信 W / 收信 R / 转发 F …），
        写信必须用 type=='W' 那一份。
        """
        try:
            raw = self.tab.evaluate(
                "(()=>{const out={};"
                "document.querySelectorAll('li').forEach(li=>{"
                " const v=li.__vue__; if(!v) return;"
                " const c=v.command;"
                " if(!c||typeof c!=='object'||!c.templateId) return;"
                " const n=(li.innerText||'').trim(); if(!n) return;"
                " const cur=out[n];"
                " if(!cur||(c.type==='W'&&cur.type!=='W')) out[n]={templateId:c.templateId,type:c.type};"
                "});"
                "return JSON.stringify(out);})()", timeout=20)
            d = json.loads(raw or "{}")
            return d if isinstance(d, dict) else {}
        except Exception:
            return {}

    def _click_template_vue(self, name: str) -> str:
        """直接调 Vue 组件方法选模板。返回 'CALLED' / 'NOT_FOUND' / 'ERR ..'。"""
        try:
            return str(self.tab.evaluate(
                "((name)=>{const li=Array.from(document.querySelectorAll('li'))"
                ".filter(e=>(e.innerText||'').trim()===name"
                " && e.__vue__ && e.__vue__.command"
                " && e.__vue__.command.templateId"
                " && (e.__vue__.command.type==='W'||!e.__vue__.command.type))"
                ".sort((a,b)=>(a.__vue__.command.type==='W'?0:1)"
                "-(b.__vue__.command.type==='W'?0:1)).pop();"
                "if(!li) return 'NOT_FOUND';"
                "try{ li.__vue__.handleClick(); return 'CALLED'; }"
                "catch(e){ return 'ERR '+e.message; }})("
                + json.dumps(name, ensure_ascii=False) + ")", timeout=20))
        except Exception as e:
            return "ERR " + str(e)

    def _li_ready(self, name: str, timeout: float = 6.0) -> bool:
        """等某个下拉项真的展开出高度（>5px）。等不到返回 False。"""
        js = ("((name)=>{const li=Array.from(document.querySelectorAll('li'))"
              ".filter(e=>(e.innerText||'').trim()===name).pop();"
              "if(!li) return 0;const b=li.getBoundingClientRect();"
              "return Math.round(b.height*10)/10;})("
              + json.dumps(name, ensure_ascii=False) + ")")
        end = time.time() + timeout
        h = 0.0
        while time.time() < end:
            try:
                h = float(self.tab.evaluate(js, timeout=8) or 0)
            except Exception:
                h = 0.0
            if h > 5:
                return True
            self._sleep(0.4)
        self._say("[模板] 菜单项「%s」一直没展开出高度（最后 %spx）——"
                  "多半是浏览器在后台，Vue 动画没跑" % (name, h))
        return False

    def _click_template_li(self, name: str) -> bool:
        """选下拉里的模板并进写信页。成功返回 True。

        2026-10-05 重写（真机实测，顺序即优先级）：
          ① **Vue 组件方法** li.__vue__.handleClick()  ← 主力
             不展开菜单、不等动画、不点坐标、不抢焦点。
             浏览器被别的窗口挡着（document.hidden=true、li 高度 0）时，
             只有这条路能成 —— 实测 li.click() 和真实鼠标点全部落空，
             而 handleClick() 1 秒进写信页，正文/附件都正确套用模板。
          ② 兜底：JS 派发 click（页面恰好可见时有效）
          ③ 最后：真实鼠标点（要 bringToFront，只有用户勾了③页
             「允许把浏览器拉到前台」才会走到这里）
        """
        before = self.tab.url or ""
        moved = lambda v: "/mail/new" in str(v) and str(v) != before

        # ① Vue 组件方法
        r = self._click_template_vue(name)
        if r == "CALLED":
            if self._wait("location.href", moved, timeout=8.0):
                self._say("[模板] ① Vue 方法选模板成功:", name)
                return True
            self._say("[模板] ① Vue 方法已调用但没跳转，试老办法")
        else:
            self._say("[模板] ① Vue 方法不可用（%s），试老办法" % r)

        # ② JS 派发 click
        try:
            self.tab.evaluate(
                "((name)=>{const lis=Array.from(document.querySelectorAll('li'))"
                ".filter(e=>(e.innerText||'').trim()===name);"
                "if(!lis.length) return 'NOT_FOUND'; const li=lis[lis.length-1];"
                "['mouseenter','mouseover'].forEach(tn=>li.dispatchEvent(new MouseEvent(tn,{bubbles:true})));"
                "li.click(); return 'CLICKED';})("
                + json.dumps(name, ensure_ascii=False) + ")", timeout=20)
        except Exception:
            pass
        if self._wait("location.href", moved, timeout=6.0):
            self._say("[模板] ② JS click 成功:", name)
            return True

        # ③ 真实鼠标（需要前台；auto_raise 关着就不做，免得打断用户）
        if not getattr(self.s, "auto_raise", False):
            self._say("[模板] ③ 需要真实鼠标点击，但没开「允许拉前台」→ 放弃，"
                      "不让浏览器跳出来打断你。要排查请到③页勾上那个开关。")
            return False
        self._bring_front()
        self._li_ready(name, timeout=6.0)
        r = self._rect_center(
            "((name)=>{const li=Array.from(document.querySelectorAll('li'))"
            ".filter(e=>(e.innerText||'').trim()===name"
            "&&e.getBoundingClientRect().height>5).pop();"
            "if(!li) return 'null';const b=li.getBoundingClientRect();"
            "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,"
            "w:b.width,h:b.height});})("
            + json.dumps(name, ensure_ascii=False) + ")")
        if r is not None:
            try:
                self._real_click(r["x"], r["y"])
            except Exception:
                pass
            return self._wait("location.href", lambda v: "/mail/new" in str(v), timeout=7.0)
        return False

    # ---------------- 窗口可见性 ----------------
    def ensure_visible(self) -> bool:
        """窗口被最小化就自动恢复 —— 但**绝不抢用户正在用的窗口**。

        为什么必须恢复：最小化状态下鼠标事件全部落空（选国家那一步就是真鼠标点），
        Page.captureScreenshot 也会一直挂住导致超时 —— 表现为"卡住不动"。

        ⚠ 2026-10-06 真机取证后重写：
        原来这里走 CDP 的 Browser.setWindowBounds("normal")。实测它会把 Edge
        **激活到前台**（调用后 1.2 秒，前台窗口就从别的程序变成 Edge）。脚本每封
        邮件都会走到这一步，用户正在别的窗口干活时就被反复抢走、窗口被挤到后面，
        表现就是"我一打开窗口，几秒就没了"。

        现在首选 Win32 ShowWindow(SW_SHOWNOACTIVATE)：窗口恢复成普通大小、
        页面 document.hidden 变回 false（真实鼠标点击/截图照常可用），
        但前台窗口完全不动。只有拿不到窗口句柄时才退回 CDP（会抢前台）。
        """
        if not self.s.auto_restore:
            return False
        try:
            st = self.browser.window_state(self.tab.id)
        except Exception:
            return False
        if st != "minimized":
            return False
        # ① 后台恢复：不碰用户的前台窗口
        try:
            n = self.browser.restore_no_activate(self.tab.id)
            if n:
                self.log("  ⚠ Edge 被最小化了 → 已在后台恢复"
                         "（不会抢走你正在用的窗口）")
                self._sleep(0.8)
                return True
        except Exception as e:
            self.log("  （后台恢复窗口失败，改用 CDP）", type(e).__name__)
        # ② 兜底：CDP 恢复。会抢前台，但总比卡死强。
        try:
            self.log("  ⚠ Edge 被最小化了 → 用 CDP 恢复"
                     "（这一步可能会把 Edge 提到前台）")
            self.browser.set_window_state(self.tab.id, "normal")
            self._sleep(0.8)
            return True
        except Exception as e:
            self.log("  （自动恢复窗口失败）", type(e).__name__)
            return False

    # ---------------- 连接 ----------------
    def connect(self) -> None:
        if not self.s.profile_dir:
            raise MXWebError("必须提供独立的 profile_dir（不能占用日常浏览器配置）")
        self.browser = cdp.launch(
            edge_path=self.s.browser or None,
            port=self.s.port,
            user_data_dir=self.s.profile_dir,
            url=self.s.login_url,       # 2026-09-26：新开就直接开孚盟，别开空白页
            headless=self.s.headless,
        )
        # 2026-09-26 修复：原来无脑绑第一个标签页——Edge 里若开着 about:blank
        # 或别的页（排在第一位），脚本就绑到了空白页上，用户看的孚盟页反而是
        # 另一个标签页 → 表现就是"页面明明加载好了，却读不到模板/显示未登录"。
        # 现在优先挑孚盟标签页；实在没有，就把第一个标签页带去孚盟。
        self.tab = self._pick_fumamox_tab()
        if self.tab is None:
            self.tab = self.browser.first_tab()
        if self.tab is None:
            raise MXWebError("没有可用的标签页")
        if "fumamx" not in (self.tab.url or ""):
            try:
                self.tab.goto(self.s.login_url, timeout=60)
            except Exception as exc:
                self._say("打开孚盟首页失败（稍后重试）:", exc)
        # 最小化窗口会让鼠标事件落空、截图挂死 —— 必须先恢复。
        # 2026-10-06：原来这里**无条件 maximized**，一连接就把浏览器铺满整屏、
        # 把用户正在看的窗口盖住。改成只有真被最小化时才恢复，而且用"不激活"
        # 的方式，绝不动用户的前台窗口。
        try:
            if self.browser.window_state(self.tab.id) == "minimized":
                if not self.browser.restore_no_activate(self.tab.id):
                    self.browser.set_window_state(self.tab.id, "normal")
        except Exception as exc:
            self._say("恢复窗口失败:", exc)

    def _pick_fumamox_tab(self):
        """在所有标签页里挑孚盟的（优先用户正在看的那个 = 列表最后一个）。"""
        try:
            ts = [t for t in self.browser.targets() if "fumamx" in (t.get("url") or "")]
            if ts:
                return self.browser.attach(ts[-1])   # 多个孚盟页时绑最新的
        except Exception:
            pass
        return None

    # ---------------- 页面僵死自救（2026-09-23 新增） ----------------
    def page_health(self) -> tuple[bool, str]:
        """体检当前页面，返回 (是否健康, 状态描述)。

        三种状态：
          dead  —— 连 evaluate 都超时：标签页被冻结/丢弃，或 JS 卡死
          blank —— 地址读得到但页面没内容：白屏、SPA 组件没挂载、
                   或断网期间接口全挂、网络恢复后页面不会自己重拉数据
          ok
        """
        try:
            v = self.tab.evaluate(
                "(()=>{try{return JSON.stringify({h:location.href||'',"
                "r:document.readyState||'',"
                "n:(document.body?document.body.innerText.trim().length:-1),"
                "c:(document.body?document.body.children.length:-1)});}"
                "catch(e){return 'ERR';}})()", timeout=8)
        except Exception as e:
            return False, "dead:%s" % type(e).__name__
        if not v or v == "ERR":
            return False, "dead:no-response"
        try:
            d = json.loads(v)
        except Exception:
            return False, "dead:bad-json"
        href = d.get("h") or ""
        if not href or href == "about:blank":
            return False, "dead:blank-url"
        # 白屏判定：正文几乎没字 且 body 几乎没有子节点
        try:
            n, c = int(d.get("n", 0)), int(d.get("c", 0))
        except Exception:
            n, c = 0, 0
        if n < 30 and c <= 2:
            return False, "blank:text=%s,children=%s" % (n, c)
        return True, "ok"

    def revive(self, reason: str = "", wait: float = 40.0) -> bool:
        """页面无响应时自救：强制刷新（等于自动帮用户按 F5）并等孚盟重新渲染。

        为什么必须有这一步：页面被冻结或 SPA 卡死后，**光重连 CDP 是救不回来的**
        ——重连只是重新绑到同一个已经僵死的标签页。用户平时遇到这种情况都是
        手工按 F5 才恢复，挂机几小时没人盯着时就会白白失败一批客户。
        """
        self.log("  ⟳ 页面无响应（%s），自动刷新恢复…" % (reason or "未知原因"))
        try:
            self.ensure_visible()          # 最小化状态下刷新也渲染不出来
        except Exception:
            pass
        try:
            ok = self.tab.reload(ignore_cache=True, timeout=60)
        except Exception as e:
            self.log("  （刷新失败）%s: %s" % (type(e).__name__, e))
            ok = False
        if not ok:
            return False
        deadline = time.time() + max(5.0, wait)
        while time.time() < deadline:
            if self.abort:
                raise MXWebError("已手动停止")
            if self.pause:
                time.sleep(0.2)
                continue
            healthy, st = self.page_health()
            if healthy:
                self.log("  ✓ 刷新后页面已恢复（%s），继续发送" % st)
                return True
            time.sleep(1.0)
        self.log("  ✗ 刷新后页面仍未恢复")
        return False

    def ensure_healthy(self, reason: str = "") -> bool:
        """每次关键操作前调一次：页面僵死/白屏就自动刷新救回来。
        返回 True=健康（或已救回），False=救不回来。"""
        try:
            healthy, st = self.page_health()
        except Exception as e:
            healthy, st = False, "dead:%s" % type(e).__name__
        if healthy:
            return True
        return self.revive(reason or st)

    # ---------------- JS 小工具 ----------------
    def _js_click_text(self, text: str, restrict_xy: bool = True) -> str:
        return self.tab.evaluate(
            "((txt, rxy) => {"
            "  const els = Array.from(document.querySelectorAll('button,span,div,a,li,p,label,em,i'));"
            "  const c = els.filter(e => {"
            "    const r = e.getBoundingClientRect();"
            "    if (r.width<=0 || r.height<=0) return false;"
            "    if (rxy && (r.x<0 || r.y<0)) return false;"
            "    return (e.innerText||'').trim() === txt;"
            "  });"
            "  c.sort((a,b)=>a.querySelectorAll('*').length-b.querySelectorAll('*').length);"
            "  if (!c.length) return 'NOT_FOUND:' + txt;"
            "  c[0].scrollIntoView({block:'center'}); c[0].click();"
            "  return 'CLICKED<' + c[0].tagName + '>';"
            "})(" + json.dumps(text, ensure_ascii=False)
            + ", " + ("true" if restrict_xy else "false") + ")", timeout=25)

    def _visible_dialogs(self) -> dict:
        """只看"真·看得见"的弹窗。

        2026-09-20 踩坑：孚盟页面里常驻着一堆 w=0/h=0 的隐藏 .el-dialog
        （菜单栏显示设置 / 添加菜单 / 选择图片 / 定时发送 ……），只判
        display!=='none' 会把它们当成真弹窗，导致误点其中的「确定」。
        所以这里必须按**实际尺寸**过滤（>20x20）。
        """
        raw = self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>20&&r.height>20&&getComputedStyle(e).display!=='none';};"
            "const dlgs=Array.from(document.querySelectorAll('.el-dialog,.el-message-box'))"
            ".filter(ok).map(e=>({t:(e.innerText||'').replace(/\\s+/g,' ').slice(0,250),"
            " btns:Array.from(e.querySelectorAll('button')).map(b=>(b.innerText||'').trim()).filter(Boolean)}));"
            "const toasts=Array.from(document.querySelectorAll('.el-message,.el-notification'))"
            ".filter(ok).map(e=>(e.innerText||'').replace(/\\s+/g,' ').trim().slice(0,120));"
            "return JSON.stringify({dialogs:dlgs,toasts});})()", timeout=25)
        try:
            return json.loads(raw)
        except Exception:
            return {}

    def _confirm_dialog(self, rounds: int = 4) -> list[str]:
        """处理发送确认弹窗。只点"与发信相关"的弹窗按钮，杜绝误点页面残留弹窗。"""
        toasts: list[str] = []
        REL_KEYS = ("发送", "收件人", "主题", "附件", "确定", "确认", "继续")
        for _ in range(rounds):
            d = self._visible_dialogs()
            toasts += d.get("toasts", [])
            rel = [dg for dg in d.get("dialogs", [])
                   if any(k in dg.get("t", "") for k in REL_KEYS)]
            btns = [b for dg in rel for b in dg.get("btns", [])]
            pick = next((w for w in CONFIRM_WORDS if w in btns), None)
            if not pick:
                break
            self._say("确认弹窗:", btns, "-> 点", pick)
            self.tab.evaluate(
                "((n)=>{const vis=e=>{const r=e.getBoundingClientRect();return r.width>0&&r.height>0;};"
                "const bs=Array.from(document.querySelectorAll('.el-dialog button,.el-message-box button'))"
                ".filter(vis).filter(b=>(b.innerText||'').trim()===n);"
                "if(bs.length){bs[0].click(); return 'ok';} return 'nf';})("
                + json.dumps(pick, ensure_ascii=False) + ")", timeout=20)
            self._sleep(1.5)
        return toasts

    # ---------------- 登录 ----------------
    # 2026-10-12 孚盟系统更新后重写（真机 CDP 取证，不是猜的）。
    #
    # 现场证据（连到用户正在用的浏览器，端口 9333）：
    #   URL    https://www.fumamx.com/#/main/workbench   ← 工作台，不是原来的 mail/home
    #   根节点 div.app（旧版是 #app，现在 #app 根本不存在）
    #   页面里**有 3 个 input[type=password]**，placeholder 分别是
    #       「输入原密码」「长度至少6位，需要同时含字母加数字」「确认新密码」
    #       外层 div.nuc-basic-info__input-wide.el-input —— 是「修改密码」表单
    #   但这 3 个框 getClientRects().length == 0、宽高都是 0，**看不见**。
    #
    # 旧判据是「页面上有密码框 = 没登录」→ 把这 3 个隐藏框当成了登录页，
    # 于是用户明明登录了，脚本永远说「未检测到登录」。
    #
    # 新判据分四层，从"确定没登录"到"确定已登录"依次排：
    #   OFFSITE     不在孚盟域名上            → 未登录（2026-09-26 的老修复，保留）
    #   PWD         有**看得见**的密码框       → 未登录（登录页的密码框一定是可见的）
    #   LOGINROUTE  地址里有 /login           → 未登录（老修复，保留）
    #   APP         主框架 div.app 有实质内容   → 已登录（新版已登录页 innerHTML 约 20 万字符）
    #   EMPTY       框架是空的/还没渲染        → 未登录（宁可多等，不误判已登录）
    JS_LOGIN_STATE = (
        "(()=>{const u=location.href||'';"
        "if(u.indexOf('fumamx')<0) return 'OFFSITE';"
        # 只认看得见的密码框：getClientRects 对 position:fixed 也有效，
        # 比 offsetParent 可靠（顶部栏可能用 fixed，会把可见元素误判成隐藏）。
        "const vis=el=>{try{return el.getClientRects().length>0;}catch(e){return false;}};"
        "const pwd=Array.from(document.querySelectorAll('input[type=password]')).filter(vis);"
        "if(pwd.length) return 'PWD';"
        "if((location.hash||'').indexOf('/login')>=0) return 'LOGINROUTE';"
        "const app=document.querySelector('div.app')||document.querySelector('#app');"
        "if(app&&app.innerHTML.length>2000) return 'APP';"
        "return 'EMPTY';})()"
    )

    def login_state(self) -> str:
        """登录状态机：OFFSITE / PWD / LOGINROUTE / APP / EMPTY / ERR。

        只有 APP 算已登录。返回具体状态而不是一个布尔值，是为了让日志能说清
        「为什么判定未登录」—— 上次改版就是靠这个才定位到隐藏密码框的。
        """
        try:
            return str(self.tab.evaluate(self.JS_LOGIN_STATE, timeout=15))
        except Exception:
            return "ERR"

    def is_logged_in(self) -> bool:
        """是否已登录。任何异常/拿不到状态一律按未登录，绝不向上抛。

        历史修复记录：
          · 2026-09-20：原 JS 正则 /#\\/login/i 是坏的（浏览器报 Invalid regular
            expression flags）→ 永远抛 CDPError、永远返回 False。
          · 2026-09-26：about:blank / 非孚盟域名原来会被误判"已登录"（没密码框、
            hash 也不含 /login）→ 绑错标签页还自以为正常。
          · 2026-10-12：孚盟改版后工作台常驻一份隐藏的「修改密码」表单（3 个
            input[type=password]，0×0 不可见），旧判据把它当登录页 → 永远未登录。
            改为只认**可见**密码框 + 用 div.app 有内容做正向确认。
        """
        return self.login_state() == "APP"

    def ensure_login(self) -> None:
        """注意：实测登录态不跨浏览器重启（令牌是会话级的）。
        所以重启后要么这里等待人工登录，要么调用方先自动填表登录。"""
        self.connect()
        self.ensure_visible()          # 最小化会导致后面所有鼠标操作落空
        if self.is_logged_in():
            self._say("已登录:", self.tab.url)
            return
        self._say("未登录，请在浏览器窗口里手工登录（脚本不接触账号密码）...")
        if "fumamx.com" not in (self.tab.url or ""):
            self.tab.goto(self.s.login_url, timeout=60)
        deadline = time.time() + self.s.login_wait_minutes * 60
        while time.time() < deadline:
            if self.is_logged_in():
                self._say("登录成功:", self.tab.url)
                return
            self._sleep(2)
        raise MXWebError(f"{self.s.login_wait_minutes} 分钟内未检测到登录")

    # ---------------- 进写信页 ----------------
    def _open_template_dropdown(self) -> list[str]:
        """点开「写邮件」右边的下拉图标，返回模板名列表。

        ★ 2026-10-12 新版快速通道（实测，放在最前）：
          孚盟改版后模板菜单**常驻 DOM** —— popper 挂在 `body` 下
          （class: `el-dropdown-menu MXscroll new2026-el-dropdown`），
          收起时 display:none 但 li 全在，且每项都带
          `__vue__.command.templateId`（实测 8 项，含「未回复」18873624）。
          所以根本不用展开、不用派发鼠标事件、不用把浏览器抢到前台，
          直接全页面扫 li 就能拿到全部模板。
          零副作用 —— 用户盯着屏幕时也不会被抢焦点。
          扫不到（旧版环境）才走下面那套老逻辑。

          注意：菜单**不在** `.mail_dir-operat` 里（第一版快速通道写成
          `.mail_dir-operat .el-dropdown-menu`，实测返回空 —— popper 是
          append-to-body 的）。所以这里按 templateId 全页面扫 + 按 id 去重。
        """
        try:
            quick = self.tab.evaluate(
                "(()=>{const seen={},out=[];"
                "document.querySelectorAll('li').forEach(li=>{"
                "  const c=li.__vue__&&li.__vue__.command;"
                "  if(c&&typeof c==='object'&&c.templateId){"
                "    const t=(li.innerText||'').trim();"
                "    if(t&&!seen[c.templateId]){seen[c.templateId]=1;out.push(t);}}});"
                "return JSON.stringify(out);})()", timeout=15)
            quick_items = json.loads(quick) if quick else []
        except Exception:
            quick_items = []
        if len(quick_items) >= 1:
            self._say("[模板扫描] 新版直读（菜单常驻，零打扰）: %d 项: %s"
                      % (len(quick_items), quick_items[:8]))
            return [str(x) for x in quick_items]

        def scan() -> tuple[list[dict], dict]:
            try:
                d = json.loads(self.tab.evaluate(JS_TPL_SCAN, timeout=25))
                if isinstance(d, dict):
                    return (d.get("groups") or []), (d.get("icon") or {})
            except Exception:
                pass
            return [], {}

        def key(g: dict):
            return (g.get("cls", ""), tuple(g.get("items") or []))

        def pick(groups: list[dict]) -> list[str]:
            cand = [g for g in groups if g.get("n", 0) > 0 and g.get("items")]
            if not cand:
                return []
            cand.sort(key=lambda g: (int(g.get("n", 0)), int(g.get("raw", 0))))
            return [str(x) for x in cand[-1]["items"]]

        def pick_geo(groups: list[dict], icon: dict) -> list[str]:
            """★ 主力：按几何位置选组。

            模板菜单是挂在「写邮件」下拉图标**正下方、左边缘对齐**的那个 popper。
            实测：图标 left=196；菜单 ul left=196（li 中心 x=305.5 宽 219）。
            页面上 300+ 个下拉里只有它满足这个关系，而且**菜单收起时 left 依旧对齐**
            —— 所以不管动画跑没跑、菜单是不是刚创建，这一条都成立。
            """
            if not icon:
                return []
            ix, iy = icon.get("x", 0) + icon.get("w", 0) / 2.0, icon.get("y", 0)
            ib = iy + icon.get("h", 0)
            cand = []
            for g in groups:
                if g.get("n", 0) <= 0 or not g.get("items"):
                    continue
                r = g.get("rect") or {}
                if not r:
                    continue
                d_left = abs(r.get("x", 9999) - (ix - icon.get("w", 0) / 2.0))
                d_ctr = abs((r.get("x", 0) + r.get("w", 0) / 2.0) - ix)
                dy = r.get("y", 9999) - ib          # 菜单应在图标下方
                if (d_left < 45 or d_ctr < 130) and -80 <= dy <= 220:
                    cand.append(g)
            if not cand:
                return []
            # 排序优先级（从低到高，取最后一个）：
            #   ① popper（真实弹出层）
            #   ② nid = 组里**带 templateId 的 li 个数** ← 2026-10-08 新增，最关键
            #      模板菜单的每一项都带 templateId；其它同名/同位置的菜单不带。
            #   ③ n  = 过滤后的项数
            # 以前只按 ①③ 排，若某干扰组项数更多就会赢走。
            cand.sort(key=lambda g: (int(g.get("popper", 0)),
                                     int(g.get("nid", 0)),
                                     int(g.get("n", 0))))
            return [str(x) for x in cand[-1]["items"]]

        def open_dd(real: bool) -> bool:
            """点开「写邮件」右边的下拉图标。real=True 用 CDP 真实鼠标。"""
            if not real:
                try:
                    self.tab.evaluate(
                        "(()=>{const dd=document.querySelector('"
                        + SEL["compose_dropdown"] + "');"
                        "if(!dd) return 'NO_DD';"
                        "const ic=dd.querySelector('i.iconfont')||dd;"
                        "const rc=ic.getBoundingClientRect();"
                        "['mouseenter','mouseover','mousemove'].forEach(tn=>"
                        "  ic.dispatchEvent(new MouseEvent(tn,{bubbles:true,"
                        "    clientX:rc.x+rc.width/2, clientY:rc.y+rc.height/2})));"
                        "ic.click(); return 'OK';})()", timeout=25)
                    return True
                except Exception:
                    return False
            r = self._rect_center(
                "(()=>{const t=document.querySelector('"
                + SEL["compose_dropdown_icon"] + "')||"
                "document.querySelector('" + SEL["compose_dropdown"] + "');"
                "if(!t) return 'null';const b=t.getBoundingClientRect();"
                "if(b.width<=0||b.height<=0) return 'null';"
                "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,"
                "w:b.width,h:b.height});})()")
            if r is None:
                return False
            try:
                self._real_click(r["x"], r["y"])
                return True
            except Exception:
                return False

        base_g, _ = scan()
        base = {key(g) for g in base_g}        # 点开前的基线
        best: list[str] = []
        how = "无"
        for rnd in range(3):
            if rnd == 0:
                # 先把浏览器带前台：后台状态下 Vue 动画不跑，下拉展开不到位
                self._bring_front()
                open_dd(real=False)
            elif rnd == 1 and not best:
                open_dd(real=True)             # JS 派发不灵时换真实鼠标
            # 不能用 _settle(JS_TPL_COUNT)——那个条件页面本来就为真，等了等于没等
            self._sleep(1.2 if rnd < 2 else 0.3)
            groups, icon = scan()
            # ① 几何定位（主力，菜单常驻时也能认出来）
            got = pick_geo(groups, icon)
            if got:
                how = "几何定位"
            else:
                # ② 差集：菜单是"点开后才创建"的（首次展开时有效）
                fresh = [g for g in groups if key(g) not in base]
                got = pick(fresh)
                if got:
                    how = "差集"
            if not got and rnd == 2:
                # ③ 全量兜底
                got = pick(groups)
                how = "全量兜底（请核对是不是真模板）"
            if got and len(got) > len(best):
                best = got
            if len(best) >= 2:
                break
        self._say("[模板扫描] 图标=%s → 读到 %d 项（%s）: %s"
                  % (icon or "无", len(best), how, best[:8]))
        return best

    def template_names(self) -> list[str]:
        """读取「写邮件」下拉里可用的模板名（不点进去，安全只读）。

        2026-09-26 修复：原来不管页面在哪，先点一次「孚盟邮」再说。
        真机实测：已经在邮件模块里时再点，会把左侧菜单折叠/切走，
        模板下拉反而从 DOM 里消失（第 1 次点完下拉图标在，第 2 次点完就没了）。
        现在：下拉图标在就不点；不在才点，并轮询等它真的出现（最多 15 秒）。

        2026-10-12 孚盟改版：「写邮件」右边的下拉箭头整个没了，
        .mail_dir-operat .el-dropdown 匹配 0 个 → 判据放宽成"邮件操作条出现"，
        否则这里会白等 15 秒再报错。
        """
        try:
            has = self.tab.evaluate(
                "(()=>!!document.querySelector('.mail_dir-operat')"
                "||!!document.querySelector('.mail_dir-operat .el-dropdown'))()", timeout=15)
        except Exception:
            has = False
        if not has:
            self._js_click_text(SEL["menu_mail"])
            self._settle(self.s.step_delay + 2,
                         "(()=>!!document.querySelector('.mail_dir-operat')"
                         "||!!document.querySelector('.mail_dir-operat .el-dropdown'))()",
                         lambda v: v is True, timeout=15.0)
        items = self._open_template_dropdown()
        # 2026-10-08：「回复」「转发」原来是硬编码排除的，但那两个词同样可能是
        # 真模板名的一部分（比如「未回复」）。扫描阶段已按 templateId 判过真伪，
        # 这里只剔除明确的占位项，不再按动作词猜。
        return [i for i in items if i not in ("选择模板", "请选择模板", "暂无模板")]

    def open_compose(self, template_name: str = "", _retry: bool = False,
                     allow_blank: bool = False) -> None:
        """进写信页。给了 template_name 就走「写邮件」右边的下拉选模板（推荐）。

        2026-09-20 实测：即使下拉没真正展开，直接对隐藏的 <li> 派发 click 也能
        触发 Vue 的命令，跳转到 /mail/new?templateId=xxx。所以这里不做可见性强校验。
        2026-09-21 补充：连续跑很多封后，孚盟偶发停在「孚盟助手/快速进入平台」
        引导页、邮件模块没渲染出来 → 刷新整页后重试一次。
        2026-10-04 修复「读取历史主题失败：未能进入写信页」：
          根因是模板下拉抓错了组（抓到账号下拉的「全部账号」），
          而老代码的兜底是「全页面找同名 li」→ 正好命中那个假 li →
          点下去当然不跳转。**兜底已从"全页面匹配"改成"只认模板组"**，
          抓不到就明确报错，绝不拿一个不是模板的条目去点。
        allow_blank=True 时（只读用途，如读历史主题）允许最后裸跳写信页兜底；
        发送用途**绝不能**开——那会发出没有正文的空白邮件。
        """
        if "/mail/new" in (self.tab.url or "") and not template_name:
            return
        r = self._js_click_text(SEL["menu_mail"])
        if r.startswith("NOT_FOUND"):
            raise MXWebError(f"找不到菜单「{SEL['menu_mail']}」: {r}")
        # 2026-10-12：判据放宽。新版「写邮件」是 div.mail_dir-operat-write，
        # 旁边没有 .el-dropdown 箭头了；旧版仍有。两个都认，避免白等 8 秒。
        self._settle(self.s.step_delay + 2,
                     "(()=>!!document.querySelector('.mail_dir-operat')"
                     "||!!document.querySelector('.mail_dir-operat .el-dropdown'))()",
                     lambda v: v is True, timeout=8.0)

        if template_name:
            items = self._open_template_dropdown()
            if template_name not in items:
                items = self._open_template_dropdown()   # 页面可能刚渲染完，再读一次
            if template_name not in items:
                # 2026-10-04：这里原来会去"全页面找同名 li"兜底，结果点到账号下拉
                # 的「全部账号」，点了不跳转，白等 10 秒还报"未能进入写信页"。
                # 现在直接把话说清楚，让用户去重读模板，而不是拿假条目去试。
                raise MXWebError(
                    f"模板下拉里没有「{template_name}」。\n"
                    f"这次真正读到的模板：{items or '（一个都没有）'}\n"
                    "请回①页重点「读取模板列表」，再从下拉里选一个真模板。"
                    + ("\n（下拉里若只有「全部账号」这类东西，说明抓到的是账号切换下拉，"
                       "请先到浏览器里点一下左侧「孚盟邮」再回来读。）"
                       if len(items) <= 1 else ""))
            self._say("选择模板:", template_name, "/ 可选:", items)
            self._click_template_li(template_name)
        else:
            r = self._js_click_text(SEL["compose_new"])
            if r.startswith("NOT_FOUND"):
                raise MXWebError(f"找不到「{SEL['compose_new']}」: {r}")

        if self.s.fast:
            self._wait("location.href", lambda v: "/mail/new" in str(v), timeout=10.0)
            self._wait(JS_COMPOSE_READY, lambda v: v is True, timeout=6.0)
        else:
            self._sleep(self.s.step_delay + 4)

        if "/mail/new" not in (self.tab.url or ""):
            if not _retry:
                self.log("（没进写信页：孚盟可能停在引导页，刷新后重试一次）")
                try:
                    self.tab.evaluate("location.reload()", timeout=20)
                except Exception:
                    pass
                self._sleep(7)
                try:
                    self.ensure_login()
                except Exception:
                    pass
                return self.open_compose(template_name, _retry=True,
                                         allow_blank=allow_blank)
            if allow_blank:
                # 只读用途（读历史主题）的最后兜底：裸跳写信页，不套模板。
                # 发送用途不会走到这里（allow_blank 默认 False）。
                self.log("（兜底：直接跳写信页，本次不带模板）")
                try:
                    self.tab.evaluate("location.hash='#/main/mail/new'", timeout=20)
                except Exception:
                    pass
                self._wait("location.href", lambda v: "/mail/new" in str(v), timeout=10.0)
                self._wait(JS_COMPOSE_READY, lambda v: v is True, timeout=6.0)
        if "/mail/new" not in (self.tab.url or ""):
            raise MXWebError(f"未能进入写信页，当前 {self.tab.url}")

    # ---------------- 主题历史 ----------------
    def _focus_subject(self) -> None:
        self.tab.evaluate(
            "(()=>{const s=Array.from(document.querySelectorAll('" + SEL["subject"] + "'))"
            ".find(x=>{const rc=x.getBoundingClientRect();return rc.width>50&&rc.height>10;});"
            "if(!s) return; s.focus(); s.click();})()", timeout=20)

    def _subject_text(self) -> str:
        v = self.tab.evaluate(
            "(()=>{const s=Array.from(document.querySelectorAll('" + SEL["subject"] + "'))"
            ".find(x=>{const rc=x.getBoundingClientRect();return rc.width>50&&rc.height>10;});"
            "return s ? (s.innerText||'').trim() : '';})()", timeout=15)
        return v or ""

    def _clear_subject(self) -> None:
        """Ctrl+A 再退格，把主题清空（实测：清空才会弹历史记录）。"""
        self._focus_subject()
        self._sleep(0.4)
        for typ in ("keyDown", "keyUp"):
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "a", "code": "KeyA",
                "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65,
                "modifiers": 2}, timeout=10)
        self._sleep(0.4)
        for typ in ("keyDown", "keyUp"):
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "Backspace", "code": "Backspace",
                "windowsVirtualKeyCode": 8, "nativeVirtualKeyCode": 8}, timeout=10)

    def subject_history(self) -> list[str]:
        """触发并返回「历史记录」主题列表（最上面的是第 1 条）。

        2026-09-20 实测要点：
          1. 必须把主题**清空**（Ctrl+A + Backspace）才会弹历史；仅"输入再删除"不奏效。
          2. 列表容器是 .recent-subject，子 div 是主题，末尾 <p> 是"到底了, 显示最近20条主题~"。
          3. 兜底：找全页 innerText 最短且含"到底了"的容器。
        """
        for attempt in (1, 2):
            self._clear_subject()
            self._sleep(2.5 if attempt == 1 else 3.5)
            raw = self.tab.evaluate(
                "(()=>{const box=document.querySelector('" + SEL["subject_history"] + "');"
                "if(!box) return 'NO_BOX';"
                "const kids=Array.from(box.children).filter(e=>e.tagName!=='P')"
                ".map(e=>(e.innerText||'').replace(/\\s+/g,' ').trim()).filter(Boolean);"
                "return JSON.stringify(kids);})()", timeout=25)
            if raw != "NO_BOX":
                try:
                    items = json.loads(raw)
                except Exception:
                    items = []
                items = [x for x in items if x and "到底了" not in x]
                if items:
                    return items
            # 兜底：全页找含"到底了"且文本最短的容器
            raw2 = self.tab.evaluate(
                "(()=>{let best=null;"
                "document.querySelectorAll('div,ul,section').forEach(e=>{"
                " const tx=e.innerText||''; if(!tx.includes('到底了')) return;"
                " if(!best||tx.length<best.innerText.length) best=e;});"
                "if(!best) return '[]';"
                "const kids=Array.from(best.children).filter(e=>e.tagName!=='P')"
                ".map(e=>(e.innerText||'').replace(/\\s+/g,' ').trim()).filter(Boolean);"
                "return JSON.stringify(kids);})()", timeout=25)
            try:
                items = json.loads(raw2)
            except Exception:
                items = []
            items = [x for x in items if x and "到底了" not in x]
            if items:
                return items
        return []

    def set_subject(self, subject: str) -> None:
        self.tab.evaluate(
            "((s)=>{const l=Array.from(document.querySelectorAll('" + SEL["subject"] + "'))"
            ".find(x=>{const rc=x.getBoundingClientRect();return rc.width>50&&rc.height>10;});"
            "if(!l) return 'NO_SUBJ'; l.focus(); l.innerText=s;"
            "l.dispatchEvent(new InputEvent('input',{bubbles:true})); return 'OK';})("
            + json.dumps(subject, ensure_ascii=False) + ")", timeout=25)

    # ---------------- 清理 ----------------
    def clear_cc(self) -> int:
        n = self.tab.evaluate(
            "(()=>{const chips=Array.from(document.querySelectorAll('" + SEL["chip"] + "'))"
            ".filter(e=>(e.textContent||'').includes('@'));"
            "let n=0; for(const c of chips){"
            " const d=c.querySelector('[class*=" + SEL["chip_delete"] + "]') || c.querySelector('[class*=delete]');"
            " if(d){d.click(); n++;}} return n;})()", timeout=25)
        if n:
            self._sleep(1)
        return n

    def clear_attachments(self) -> int:
        n = 0
        while True:
            r = self.tab.evaluate(
                "(()=>{const it=document.querySelector('" + SEL["attach_item"] + "');"
                "if(!it) return 'NONE';"
                "const d=it.querySelector('" + SEL["attach_delete"] + "');"
                "if(!d) return 'NO_DEL'; d.click(); return 'OK';})()", timeout=20)
            if r != "OK":
                break
            n += 1
            self._sleep(1.2)
        return n

    # ---------------- 收件人 ----------------
    def add_recipient(self, email: str) -> bool:
        """写入收件人。返回是否生成了芯片。

        实测：MX 认识的地址（自己的邮箱 / 客户库里的）失焦后变成 .address-item 芯片；
        库外地址会保持纯文本、不生成芯片 —— 但仍在输入框里，能否发出由 MX 决定。
        """
        r = self.tab.evaluate(
            "(()=>{const ins=Array.from(document.querySelectorAll('" + SEL["recipient_input"] + "'))"
            ".filter(e=>{const rc=e.getBoundingClientRect();return rc.width>50&&rc.y>0;});"
            "ins.sort((a,b)=>a.getBoundingClientRect().y-b.getBoundingClientRect().y);"
            "if(!ins.length) return 'NO_INPUT'; const e=ins[0];"
            "e.scrollIntoView({block:'center'}); e.focus(); e.click(); return 'OK';})()", timeout=25)
        if r != "OK":
            raise MXWebError(f"找不到收件人输入框: {r}")
        self.tab.send("Input.insertText", {"text": email}, timeout=15)
        self._settle(0.8, _js_recip_ready(email), lambda v: v is True, timeout=3.0)
        for typ in ("keyDown", "keyUp"):
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "Enter", "code": "Enter",
                "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13}, timeout=10)
        self._human_pause(0.6, 1.6)
        # 失焦促使成芯片
        self.tab.evaluate(
            "(()=>{const s=Array.from(document.querySelectorAll('" + SEL["subject"] + "'))"
            ".find(x=>{const rc=x.getBoundingClientRect();return rc.width>50&&rc.height>10;});"
            "if(s) s.click(); return 'OK';})()", timeout=20)
        self._settle(1.0, _js_chip(email), lambda v: v is True, timeout=2.5)
        return self.tab.evaluate(
            "(()=>{return Array.from(document.querySelectorAll('" + SEL["chip"] + "'))"
            ".filter(e=>(e.textContent||'').toLowerCase().includes("
            + json.dumps(email.lower(), ensure_ascii=False) + ")).length;})()", timeout=20) > 0

    def commit_recipient(self, email: str) -> bool:
        """不重复输入，只补一刀"回车提交 + 失焦"，再看芯片出来没有。

        2026-09-21 加固：批量发送时发现有的地址 insertText 之后回车时序不稳，
        文本停在输入框里没成芯片，自检的兜底条件（输入框里有这段字）又放行了，
        结果点「发送」时孚盟根本没按收件人算 —— 邮件就"卡在最后一步"。
        """
        self.tab.evaluate(
            "(()=>{const ins=Array.from(document.querySelectorAll('" + SEL["recipient_input"] + "'))"
            ".filter(e=>{const rc=e.getBoundingClientRect();return rc.width>50&&rc.y>0;});"
            "if(!ins.length) return 'NO_INPUT'; ins[0].focus(); return 'OK';})()", timeout=20)
        for typ in ("keyDown", "keyUp"):
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "Enter", "code": "Enter",
                "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13}, timeout=10)
        self._sleep(0.5 if self.s.fast else 0.8)
        self.tab.evaluate(
            "(()=>{const s=Array.from(document.querySelectorAll('" + SEL["subject"] + "'))"
            ".find(x=>{const rc=x.getBoundingClientRect();return rc.width>50&&rc.height>10;});"
            "if(s) s.click(); return 'OK';})()", timeout=20)
        self._settle(1.0, _js_chip(email), lambda v: v is True, timeout=2.5)
        return self.tab.evaluate(
            "(()=>{return Array.from(document.querySelectorAll('" + SEL["chip"] + "'))"
            ".filter(e=>(e.textContent||'').toLowerCase().includes("
            + json.dumps(email.lower(), ensure_ascii=False) + ")).length;})()", timeout=20) > 0

    def capture_fail(self, tag: str, out_dir: str = "") -> str:
        """失败现场留存：截图 + 页面状态 JSON，返回文件前缀路径。"""
        import binascii
        d = Path(out_dir) if out_dir else Path("logs") / "fail"
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:
            d = Path(".") / "logs" / "fail"
            d.mkdir(parents=True, exist_ok=True)
        stem = time.strftime("%Y%m%d-%H%M%S") + "_" + re.sub(r"[^A-Za-z0-9_.@-]", "_", tag)[:40]
        png = d / (stem + ".png")
        js = d / (stem + ".json")
        try:
            r = self.tab.send("Page.captureScreenshot", {"format": "png"}, timeout=30)
            png.write_bytes(binascii.a2b_base64(r["data"]))
        except Exception:
            pass
        try:
            state = {
                "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "url": self.tab.url,
                "page": self._visible_dialogs(),
                "recipients": self.tab.evaluate(
                    "JSON.stringify(Array.from(document.querySelectorAll('input.input-box__item'))"
                    ".map(e=>e.value||''))", timeout=15),
                "snapshot": self.snapshot(),
            }
            js.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:
            pass
        return str(png)


    # ---------------- 校验 ----------------
    def snapshot(self) -> dict:
        raw = self.tab.evaluate(
            "(()=>{const vis=e=>{const r=e.getBoundingClientRect();return r.width>0&&r.height>0;};"
            "const subj=Array.from(document.querySelectorAll('" + SEL["subject"] + "')).find(vis);"
            "let body=''; const f=document.querySelector('" + SEL["body_iframe"] + "');"
            "if(f){try{body=f.contentDocument.body.innerText.trim();}catch(e){}}"
            "const chips=Array.from(document.querySelectorAll('" + SEL["chip"] + "')).filter(vis)"
            ".map(e=>(e.textContent||'').replace(/\\s+/g,' ').trim());"
            "const files=Array.from(document.querySelectorAll('" + SEL["attach_item"] + "')).filter(vis)"
            ".map(e=>(e.textContent||'').replace(/\\s+/g,' ').trim().slice(0,40));"
            "const sender=document.querySelector('" + SEL["sender_input"] + "');"
            "return JSON.stringify({url: location.href,"
            " subject: subj?(subj.innerText||'').trim():'', bodyLen: body.length,"
            " bodyHead: body.slice(0,120), chips: chips, files: files,"
            " sender: sender?sender.value:''});})()", timeout=30)
        try:
            return json.loads(raw)
        except Exception:
            return {"raw": raw}

    # ---------------- 称呼个性化（2026-10-17 新增）----------------
    # 模板正文开头的称呼是模板里写死的，多半是 "Dear Sir/Madam"。
    # 这里在「模板已带入正文 + 收件人已填」之后、点发送之前，把开头称呼换成客户姓名。
    # 取不到姓名 / 姓名不可用 → 保持原样，绝不阻断发送。

    _NAME_BAD = {"none", "null", "nan", "-", "--", "—", "无", "/", "na", "n/a",
                 "unknown", "0", "0.0", "false", "tbd"}
    # 称呼个性化总开关：改成 False 就完全退回原有行为（正文一个字都不动）
    GREETING_ENABLED = True
    _NAME_COMPANY_RE = re.compile(
        r"(?i)\b(ltd|limited|co\.?|inc\.?|gmbh|s\.?a\.?|s\.?r\.?l|llc|pvt|corp|company"
        r"|trading|import|export|group|holdings?|enterprise|industries)\b")
    _NAME_GENERIC_RE = re.compile(
        r"(?i)^(sir|madam|sir\s*[\/&]\s*madam|dear\s*sir\s*[\/&]\s*madam|hi|hello"
        r"|hey|there|customer|buyer|friend|all)$")

    def clean_name(self, raw) -> str:
        """把表格「名字」列清洗成能直接放进称呼的字符串；不可用返回空串。"""
        if raw is None:
            return ""
        try:
            s = str(raw)
        except Exception:
            return ""
        s = (s.replace("\u00a0", " ").replace("\u3000", " ").replace("\n", " ")
              .replace("\r", " ").replace("\t", " "))
        s = re.sub(r"\s+", " ", s).strip()
        s = s.strip("\"'“”‘’()（）[]【】<>《》")
        # 名字中间夹的引号/括号一律去掉（"John "Smith" → John Smith），撇号 O'Brien 保留
        s = re.sub("[\\\"“”《》<>【】\\[\\]()（）]", "", s).strip()
        s = s.strip()
        if not s or s.lower() in self._NAME_BAD:
            return ""
        if "@" in s or "http" in s.lower() or "www." in s.lower():
            return ""
        # UTF-8 被当 latin-1 读过的典型痕迹（Wójcicki → WÃ³jcicki），拿来当称呼很难看
        if re.search(r"Ã|Â|â€", s):
            return ""
        if len(s) > 40:                      # 多半是整句备注，不是名字
            return ""
        if not re.search(r"[A-Za-z\u00c0-\u024f\u0400-\u04ff\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]", s):
            return ""
        # 纯中文长串多半是备注（"请尽快回复我们需要报价"）。
        # 注意只在"整串都是汉字"时才拦，别误伤俄语/日语/阿拉伯语的名字。
        if (re.search(r"[\u4e00-\u9fff]", s)
                and not re.search(r"[^\u4e00-\u9fff\s]", s)
                and len(re.sub(r"\s", "", s)) > 8):
            return ""
        if self._NAME_GENERIC_RE.match(s):   # 表格里本来就填的 "Sir"/"Hi"
            return ""
        if self._NAME_COMPANY_RE.search(s):  # 公司名当称呼很怪 → 降级为通用称呼
            return ""
        s = s.rstrip(",.;:，。；：、 ")
        return s.strip()

    def personalize_greeting(self, name: str = "") -> str:
        """把正文开头称呼换成客户姓名，返回一句人话（给日志用）。
        内部吞掉所有异常——这个功能绝不能影响发信主流程。"""
        if not self.GREETING_ENABLED:
            return ""
        clean = self.clean_name(name)
        js = r"""
(() => {
  try {
    const f = document.querySelector("iframe.tox-edit-area__iframe")
           || document.querySelector("iframe[id*='tox-edit-area']")
           || document.querySelector("iframe[title*='Rich Text']");
    if (!f) return JSON.stringify({ok:0, why:"no_iframe"});
    let doc = null;
    try { doc = f.contentDocument || (f.contentWindow && f.contentWindow.document); }
    catch (e) { return JSON.stringify({ok:0, why:"xdom"}); }
    if (!doc || !doc.body) return JSON.stringify({ok:0, why:"no_body"});

    const NAME = "__NAME__";
    const body = doc.body;
    const GREET = /^(dear\s+(sir\s*[\/&,]\s*madam|sirs?|madam(e)?|customer|buyer|purchasing(\s+manager)?|purchaser|gentlemen|all|team|friend|partner)[,]?|hi(\s+there)?[,!]?|hello(\s+there)?[,!]?|hey[,!]?|to\s+whom\s+it\s+may\s+concern[,:]?|good\s+(morning|afternoon|day)[,]?)\s*$/i;

    const fire = () => {
      try { body.dispatchEvent(new Event("input", {bubbles:true})); } catch (e) {}
      try { body.dispatchEvent(new Event("keyup", {bubbles:true})); } catch (e) {}
      try { body.dispatchEvent(new Event("change", {bubbles:true})); } catch (e) {}
    };

    const w = doc.createTreeWalker(body, NodeFilter.SHOW_TEXT, null);
    let tn = null, node;
    while ((node = w.nextNode())) {
      if ((node.nodeValue || "").replace(/\u00a0/g, " ").trim()) { tn = node; break; }
    }
    if (!tn) return JSON.stringify({ok:0, why:"no_text"});

    const txt = (tn.nodeValue || "").replace(/\u00a0/g, " ").trim();

    // A. 通用称呼 → 换成客户姓名（没有可用姓名就原样不动）
    if (GREET.test(txt)) {
      if (!NAME) return JSON.stringify({ok:1, mode:"keep_generic", before:txt.slice(0,40)});
      const isDear = /^dear/i.test(txt);
      const out = (isDear ? "Dear " + NAME : "Hi " + NAME) + ",";
      tn.nodeValue = out;
      fire();
      return JSON.stringify({ok:1, mode:"replace", before:txt.slice(0,40), after:out});
    }

    // B. 第一行本身已是具体称呼 → 保留，尊重手工定制
    //    （用 "Dear Mr. Smith," 这种含缩写点的也要认出来，否则会多插一行）
    if (txt.length <= 40 && /,$/.test(txt) && !/[?!]/.test(txt) &&
        /^(dear|hi|hello|hey|attn|attention|to)\b/i.test(txt)) {
      return JSON.stringify({ok:1, mode:"keep", before:txt.slice(0,40)});
    }

    // C. 没有称呼行 → 在正文最前面插一行
    if (!NAME) return JSON.stringify({ok:1, mode:"none"});
    let host = tn.parentElement;
    while (host && host.parentElement && host.parentElement !== body &&
           ["SPAN","B","I","U","STRONG","EM","FONT","A","SMALL","MARK"].indexOf(host.tagName) >= 0) {
      host = host.parentElement;
    }
    let block = null;
    if (host && ["P","DIV","H1","H2","H3","H4","LI"].indexOf(host.tagName) >= 0) {
      block = host.cloneNode(false);
    } else {
      block = doc.createElement("p");
    }
    block.textContent = "Hi " + NAME + ",";
    body.insertBefore(block, body.firstChild);
    fire();
    return JSON.stringify({ok:1, mode:"insert", after:"Hi " + NAME + ","});
  } catch (e) {
    return JSON.stringify({ok:0, why:String(e).slice(0, 90)});
  }
})()
"""
        js = js.replace("__NAME__", clean.replace("\\", "\\\\").replace('"', '\\"'))
        try:
            raw = self.tab.evaluate(js, timeout=20)
        except Exception as e:
            return "称呼个性化跳过（%s）" % type(e).__name__
        try:
            r = json.loads(raw) if isinstance(raw, str) else (raw or {})
        except Exception:
            return "称呼个性化跳过（返回值无法解析）"
        if not r.get("ok"):
            return "称呼未改动：%s" % (r.get("why") or "未知原因")
        m = r.get("mode")
        if m == "replace":
            return "称呼已替换：%s → %s" % (r.get("before"), r.get("after"))
        if m == "insert":
            return "正文无称呼行，已在开头插入：%s" % (r.get("after"),)
        if m == "keep":
            return "称呼已是具体人名，保留：%s" % (r.get("before"),)
        if m == "keep_generic":
            return "无可用姓名，保留通用称呼：%s" % (r.get("before"),)
        return "称呼未改动"

    # ---------------- 主入口 ----------------
    def send_one(self, to_email: str, subject: str = "", template_name: str = "",
                 clear_cc: bool = False, clear_files: bool = False,
                 dry_run: bool = False, require_chip: bool = True,
                 recipient_name: str = "") -> dict:
        """走完整个写信流程并发送。

        dry_run=True 只填不点发送，返回 snapshot 供人工核对。
        require_chip=False 时允许"地址停留在输入框、未成芯片"也继续（库外地址常见）。
        """
        self.ensure_login()
        self.open_compose(template_name)
        if clear_cc:
            n = self.clear_cc()
            if n:
                self._say(f"清除抄送芯片 {n} 个")
        if clear_files:
            n = self.clear_attachments()
            if n:
                self._say(f"清除附件 {n} 个")

        if subject:
            self.set_subject(subject)
            self._human_pause(0.5, 1.4)

        chip_ok = self.add_recipient(to_email)

        # 称呼个性化（2026-10-17）：把模板正文开头的通用称呼换成客户姓名。
        # 任何失败都在 personalize_greeting 内部吞掉，不影响后面的自检与发送。
        try:
            _g = self.personalize_greeting(recipient_name)
            if _g and ("已替换" in _g or "已插入" in _g):
                self._say("  " + _g)
        except Exception:
            pass

        info = self.snapshot()
        info["recipient_chip"] = chip_ok

        if dry_run:
            info["sent"] = False
            return info

        if require_chip and not chip_ok:
            raise MXWebError(f"收件人 {to_email} 未生成芯片，疑似无效地址；"
                             f"如确要发送请设 require_chip=False")
        if not info.get("subject", "").strip():
            raise MXWebError("主题为空，拒绝发送")
        if not info.get("bodyLen"):
            raise MXWebError("正文为空，拒绝发送")

        # 随机：发送前停一拍（真人会最后扫一眼邮件再点发送）
        self._human_pause(0.5, 1.5)
        r = self._js_click_text(SEL["btn_send"], restrict_xy=False)
        if r.startswith("NOT_FOUND"):
            raise MXWebError(f"找不到「发送」按钮: {r}")
        self._sleep(2.5)
        toasts = self._confirm_dialog()
        self._sleep(2)
        toasts += self._confirm_dialog()
        self._sleep(3)
        left = "/mail/new" not in (self.tab.url or "")
        ok = left or any(any(w in x for w in SUCCESS_WORDS) for x in toasts)
        # 2026-09-20 实测：页面不可见时 toast 动画不跑，只能靠"是否离开写信页"判定
        info.update({"sent": ok, "toasts": toasts, "left_compose": left})
        if not ok:
            raise MXWebError(f"未捕获「发送成功」提示，请人工确认。提示: {toasts}")
        return info

    # ---------------- 定时发送 ----------------
    # 2026-09-21 实测（真机拆解）：
    #   「发送」按钮 = span.mail_sendBtn，右边的下拉箭头 = span.box2，
    #   菜单(立即发送/定时发送/延时发送)打开后 popper 挂在 body 上（class zh-cn），
    #   所以找菜单项必须全文档找**可见**的 li，不能限定在按钮子树里。
    #   弹窗 .el-dialog--small.width420，标题「定时发送」，内含三个表单项：
    #   国家地区(el-select 带搜索框) / 时区(选国家后自动带) / 当地时间(文本可输)，
    #   按钮 取消 / 确定。
    #   **没有收件人时点「定时发送」不弹窗**——这是之前邮件"卡在最后一步"的
    #   根因之一：收件人没成芯片，点什么都没用。

    def _real_click(self, x: float, y: float):
        for typ in ("mouseMoved", "mousePressed", "mouseReleased"):
            self.tab.send("Input.dispatchMouseEvent",
                          {"type": typ, "x": x, "y": y,
                           "button": "left", "clickCount": 1}, timeout=10)
            self._sleep(0.12)

    def _rect_center(self, js: str):
        raw = self.tab.evaluate(js, timeout=20)
        try:
            r = json.loads(raw)
        except Exception:
            return None
        if not isinstance(r, dict) or r.get("w", 0) <= 0:
            return None
        return r

    def _sched_dialog_open(self) -> bool:
        return self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "return Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).some(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');})()", timeout=20) is True

    def _open_schedule_dialog(self, tries: int = 3):
        # 2026-09-21 实测：只要收件人已成芯片，直接 JS 点隐藏的菜单 li 就能弹窗，
        # 不用真实鼠标去展开下拉（那个还时灵时不灵）。收件人为空时点了也不弹。
        for attempt in range(tries):
            if self._sched_dialog_open():
                return
            self.tab.evaluate(
                "(()=>{const li=Array.from(document.querySelectorAll('ul.el-dropdown-menu li'))"
                ".find(l=>(l.innerText||'').trim()==='定时发送');"
                "if(li) li.click(); return 'ok';})()", timeout=20)
            self._settle(2.2, JS_SCHED_OPEN, lambda v: v is True, timeout=5.0)
            if self._sched_dialog_open():
                return
            # 兜底：真实鼠标走一遍「下拉箭头 → 菜单项」
            r = self._rect_center(
                "(()=>{const t=document.querySelector('span.mail_sendBtn .box2');"
                "if(!t) return 'null';const b=t.getBoundingClientRect();"
                "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height});})()")
            if r is not None:
                self._real_click(r["x"], r["y"])
                self._sleep(0.8 if self.s.fast else 1.3)
                r2 = self._rect_center(
                    "(()=>{const li=Array.from(document.querySelectorAll('ul.el-dropdown-menu li'))"
                    ".find(l=>(l.innerText||'').trim()==='定时发送'&&l.getBoundingClientRect().width>0);"
                    "if(!li) return 'null';const b=li.getBoundingClientRect();"
                    "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height});})()")
                if r2 is not None:
                    self._real_click(r2["x"], r2["y"])
                    self._settle(2.2, JS_SCHED_OPEN, lambda v: v is True, timeout=5.0)
        if not self._sched_dialog_open():
            raise MXWebError("「定时发送」弹窗没有打开（注意：没有收件人时孚盟不弹这个窗）")

    def _form_item(self, label: str) -> str:
        """弹窗里按 label 找表单项，返回其 outerHTML 片段要不要用再说，这里只当定位器。"""
        return (
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "if(!w) return 'null';"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='"
            + label + "');"
            "if(!it) return 'null';const b=it.getBoundingClientRect();"
            "return JSON.stringify({x:b.x,y:b.y,w:b.width,h:b.height});})()")

    def _pick_country_word(self, word: str) -> str:
        """用某个关键词走一遍「输入 → 选中最匹配的项」，返回 OK / NONE:..."""
        self.tab.send("Input.insertText", {"text": word}, timeout=15)
        self._settle(1.8, _js_country_hit(word), lambda v: v is True, timeout=4.0)
        return self.tab.evaluate(
            "((c)=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>5&&r.height>5&&getComputedStyle(e).display!=='none';};"
            "const lis=Array.from(document.querySelectorAll('li.el-select-dropdown__item')).filter(ok);"
            "const li=lis.find(l=>((l.title||'')||(l.innerText||'')).trim()===c)"
            " || lis.find(l=>((l.title||'')+(l.innerText||'')).indexOf(c)>=0);"
            "if(!li) return 'NONE:'+lis.map(l=>(l.innerText||'').trim()).slice(0,10).join(',');"
            "li.click(); return 'OK';})(" + json.dumps(word, ensure_ascii=False) + ")",
            timeout=25)

    def _pick_country(self, country: str):
        """在「国家地区」里搜索并选中；选完校验回显。时区由孚盟自动带出。

        2026-09-26：客户表里常写城市名（迪拜、纽约…），孚盟只认国家 →
        原来直接报错失败。现在先按原名搜；搜不到就把城市换算成国家
        （迪拜 → 阿联酋）再搜一遍；都搜不到才报错，并在报错里给出建议。
        """
        if not country:
            raise MXWebError("没有国家/地区，定时发送不知道按哪里的时间（请补全客户表的国家列）")
        # 点开下拉（el-select 只认真实鼠标事件，JS click 无效）
        r = self._rect_center(self._form_item("国家地区"))
        if r is None:
            raise MXWebError("找不到「国家地区」下拉")
        rin = self._rect_center(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='国家地区');"
            "const inp=it.querySelector('.el-select input.el-input__inner');"
            "const b=inp.getBoundingClientRect();"
            "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height});})()")
        if rin is None:
            raise MXWebError("找不到「国家地区」输入框")
        self._real_click(rin["x"], rin["y"])
        self._settle(1.3, JS_SEARCHBOX, lambda v: isinstance(v, int) and v > 0, timeout=4.0)
        # 搜索框输入国家
        rs = self._rect_center(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const s=Array.from(document.querySelectorAll('.searchBox input')).filter(ok)[0];"
            "if(!s) return 'null';const b=s.getBoundingClientRect();"
            "return JSON.stringify({x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height});})()")
        if r is None or rs is None:
            raise MXWebError("国家地区下拉/搜索框没打开")
        self.ensure_visible()          # 真鼠标点击前再确认一次窗口没被最小化
        self._real_click(rs["x"], rs["y"])
        self._sleep(0.25 if self.s.fast else 0.4)
        # ---- 搜索词顺序 ----
        # 2026-09-29：客户表国家列常写成「国家名+后缀」（美国广交会客户 / 沙特sa /
        # （阿联酋）迪拜 / 德国\n开发客户）。原串带后缀时，拿它去孚盟搜必然落空，
        # 白白等十几秒。现在交给 geo_alias.search_terms 排序：
        # 干净国家名优先，原串兜底。
        terms = [country]
        try:
            from .geo_alias import search_terms, suggest_alias
            terms = search_terms(country) or terms
        except Exception:
            suggest_alias = None
        hit = "NONE:"
        for i, w in enumerate(terms):
            if i > 0:                  # 换下一个搜索词：先清空输入框
                self.tab.evaluate(
                    "(()=>{const s=document.querySelectorAll('.searchBox input')[0];"
                    "if(!s) return ''; const r=s.getBoundingClientRect();"
                    "return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2});})()", timeout=15)
                self.tab.send("Input.dispatchKeyEvent", {
                    "type": "keyDown", "key": "a", "code": "KeyA",
                    "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65,
                    "modifiers": 2}, timeout=10)
                self.tab.send("Input.dispatchKeyEvent", {
                    "type": "keyUp", "key": "a", "code": "KeyA",
                    "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65,
                    "modifiers": 2}, timeout=10)
                self.tab.send("Input.dispatchKeyEvent", {
                    "type": "keyDown", "key": "Backspace", "code": "Backspace",
                    "windowsVirtualKeyCode": 8, "nativeVirtualKeyCode": 8}, timeout=10)
                self.tab.send("Input.dispatchKeyEvent", {
                    "type": "keyUp", "key": "Backspace", "code": "Backspace",
                    "windowsVirtualKeyCode": 8, "nativeVirtualKeyCode": 8}, timeout=10)
                self._sleep(0.3)
                if w != terms[0]:
                    self.log(f"  国家换算：客户表写「{country}」 → 孚盟按「{w}」搜索")
            hit = self._pick_country_word(w)
            if str(hit).startswith("OK"):
                break
        if not str(hit).startswith("OK"):
            tip = ""
            try:
                tip = suggest_alias(country) if suggest_alias else ""
            except Exception:
                tip = ""
            raise MXWebError(f"在孚盟的国家列表里没找到「{country}」：{hit}"
                             + (f" ｜提示：{tip}" if tip else ""))
        self._settle(1.2, JS_COUNTRY_VAL, lambda v: bool(str(v or "").strip()), timeout=4.0)
        val = self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='国家地区');"
            "return it.querySelector('.el-select input.el-input__inner').value;})()", timeout=20)
        if country[:2] not in (val or "") and (val or "").strip() == "":
            raise MXWebError(f"国家没选上（回显为空），想要 {country}")
        self.log(f"  定时-国家地区: {val}")

    @staticmethod
    def _roll_nine(when_local: str, local_now_raw: str) -> str:
        """设定时间已过（不晚于当地当前时间）→ 顺延为该时区下一个上午 9:00。

        当地还没到今天 9:00 → 用今天 9:00；已过今天 9:00 → 用明天 9:00。
        解析失败（读不到当地现在）就原样返回，交给孚盟自己报错。
        """
        def parse(s: str):
            s = (s or "").replace("-", "/").replace("：", ":").strip()
            for fmt in ("%Y/%m/%d %H:%M", "%Y/%m/%d %H:%M:%S"):
                try:
                    return datetime.strptime(s, fmt)
                except ValueError:
                    continue
            return None
        want = parse(when_local)
        now = parse(local_now_raw)
        if not want or not now or want > now:
            return when_local
        nine = now.replace(hour=9, minute=0, second=0, microsecond=0)
        target = nine if now < nine else nine + timedelta(days=1)
        return target.strftime("%Y/%m/%d %H:%M")

    def _read_time_input(self) -> str:
        """读「当地时间」输入框当前值（选完国家后默认值 = 收件人当地现在）。"""
        return str(self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "if(!w) return '';"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='当地时间');"
            "return it ? it.querySelector('input').value : '';})()", timeout=20) or "")

    def _set_local_time(self, when_local: str) -> str:
        """「当地时间」是普通文本框（默认填了当地当前时间），清空重输+回车确认。

        过时自动顺延：先读默认值（=当地现在），若设定时间已过，按规则
        顺延为该时区下一个上午 9:00，并把实际写入的时间返回给上层记录。
        """
        r = self._rect_center(self._form_item("当地时间"))
        if r is None:
            raise MXWebError("找不到「当地时间」输入框")
        local_now_raw = self._read_time_input()
        eff = self._roll_nine(when_local, local_now_raw)
        if eff != when_local:
            self.log(f"  ⏱ 设定时间 {when_local} 已过（当地现在 {local_now_raw}），"
                     f"自动顺延为 {eff}")
        when_local = eff
        self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='当地时间');"
            "const inp=it.querySelector('input'); inp.focus(); inp.click(); return 'ok';})()",
            timeout=20)
        self._sleep(0.35 if self.s.fast else 0.5)
        for typ in ("keyDown", "keyUp"):     # Ctrl+A（modifiers: Ctrl=2）
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "modifiers": 2, "key": "a", "code": "KeyA",
                "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65}, timeout=10)
        for typ in ("keyDown", "keyUp"):     # Backspace
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "Backspace", "code": "Backspace",
                "windowsVirtualKeyCode": 8, "nativeVirtualKeyCode": 8}, timeout=10)
        self._sleep(0.25 if self.s.fast else 0.3)
        self.tab.send("Input.insertText", {"text": when_local}, timeout=15)
        self._settle(0.8, JS_TIME_VAL, lambda v: bool(str(v or "").strip()), timeout=3.0)
        for typ in ("keyDown", "keyUp"):     # Enter 确认日期
            self.tab.send("Input.dispatchKeyEvent", {
                "type": typ, "key": "Enter", "code": "Enter",
                "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13}, timeout=10)
        self._settle(1.2, JS_TIME_VAL,
                     lambda v: _digits(v) == _digits(when_local), timeout=3.0)
        got = self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='当地时间');"
            "return it.querySelector('input').value;})()", timeout=20)
        digits = lambda s: "".join(ch for ch in (s or "") if ch.isdigit())
        if digits(got) != digits(when_local):
            raise MXWebError(f"当地时间没写进去：想要 {when_local}，框里是 {got!r}")
        # 加固：向输入框补发 input/change 事件，确保 Vue 模型真正收到值，
        # 否则偶发「发送时间不能小于当前时间」（框里看着对、模型里还是旧值）
        self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "const it=Array.from(w.querySelectorAll('.el-form-item'))"
            ".find(x=>((x.querySelector('.el-form-item__label')||{}).innerText||'').trim()==='当地时间');"
            "if(!it) return 'no';const inp=it.querySelector('input');"
            "inp.dispatchEvent(new Event('input',{bubbles:true}));"
            "inp.dispatchEvent(new Event('change',{bubbles:true}));return 'ok';})()", timeout=15)
        self._sleep(0.35 if self.s.fast else 0.6)
        self.log("  定时-当地时间:", got)
        return when_local

    def schedule_one(self, to_email: str, subject: str = "", template_name: str = "",
                     country: str = "", when_local: str = "",
                     require_chip: bool = True, clear_cc: bool = False,
                     recipient_name: str = "") -> dict:
        """走完整写信流程，但最后走「定时发送」弹窗，由孚盟服务器按时发送。

        when_local: 收件人当地时间，格式 YYYY/MM/DD HH:MM（弹窗里就这个格式）。
        时区概念全部交给孚盟（国家地区下拉里选），本地不做任何换算。

        require_chip=False：地址不在孚盟通讯录（库外地址）时常常不生成收件人芯片，
        但孚盟实际仍会发送。此时放行继续定时，不再误判为失败。
        """
        self.ensure_login()
        self.open_compose(template_name)
        if clear_cc:
            # 2026-10-04：模板自带的抄送（实测大量邮件抄送了一个公共地址）
            # 会让每一封实际投递 2 个地址 —— 246 个失败现场里 193 个带抄送。
            # 可在③页勾「发送前清掉抄送」，默认关（不改变原有行为）。
            n = self.clear_cc()
            if n:
                self._say(f"清除抄送芯片 {n} 个")
        if subject:
            self.set_subject(subject)
            self._human_pause(0.5, 1.4)
        chip = self.add_recipient(to_email)
        if not chip:
            chip = self.commit_recipient(to_email)

        # 称呼个性化（2026-10-17）：定时路径与即时发送保持一致。
        try:
            _g = self.personalize_greeting(recipient_name)
            if _g and ("已替换" in _g or "已插入" in _g):
                self._say("  " + _g)
        except Exception:
            pass

        info = self.snapshot()
        info["recipient_chip"] = chip
        if not chip:
            if require_chip:
                raise MXWebError(f"收件人 {to_email} 没成芯片（孚盟会报「收件人不能为空」），已取消定时")
            self.log(f"  （库外地址：{to_email} 未生成芯片，按放行继续定时）")
        if not (info.get("subject") or "").strip():
            raise MXWebError("主题为空，拒绝定时")
        if not info.get("bodyLen"):
            raise MXWebError("正文为空，拒绝定时")
        self._open_schedule_dialog()
        self._pick_country(country)
        eff_when = self._set_local_time(when_local)
        # 随机：真人会先核对一遍弹窗里的国家/时间再点确定，不会 0 秒手滑点下去
        self._human_pause(0.6, 1.8)
        toasts_before = self._visible_dialogs().get("toasts", [])
        r = self.tab.evaluate(
            "(()=>{const ok=e=>{const r=e.getBoundingClientRect();"
            " return r.width>10&&r.height>10&&getComputedStyle(e).display!=='none';};"
            "const w=Array.from(document.querySelectorAll('.el-dialog__wrapper'))"
            ".filter(ok).find(w=>((w.querySelector('.el-dialog__title')||{}).innerText||'')"
            ".trim()==='定时发送');"
            "if(!w) return 'NO_DIALOG';"
            "const btn=Array.from(w.querySelectorAll('button'))"
            ".find(b=>(b.innerText||'').trim()==='确定'&&b.className.indexOf('primary')>=0);"
            "if(!btn) return 'NO_BTN'; btn.click(); return 'OK';})()", timeout=25)
        if r != "OK":
            # 没点到「确定」→ 邮件还没提交，这种情况重试是安全的
            raise MXWebError(f"没点到「确定」: {r}")
        # ---- 从这一刻起「确定」已点击，邮件正在提交 ----
        # 2026-09-22 教训：原来只等 6 秒就判失败 → 孚盟慢一点就被冤枉成失败
        # → 上层重试 → 同一客户被定时两三次（重复邮件事故）。
        # 现在耐心等 20 秒：出报错 toast 才算真失败；只是弹窗关得慢、
        # 没有任何报错 → 一律按"已提交"处理，绝不重发。
        #
        # 2026-09-28 教训（实测反馈：孚盟网页里明明定时成功，脚本却记「失败」）：
        # 原来只要视野里**存在**含「失败/错误」字样的 toast 就算失败，
        # 但 toasts_before（点击前的快照）算了却没用 —— 于是**上一次操作残留的
        # 报错 toast** 会把这一次成功误判成失败。
        # 改成：① 只认点击后**新出现**的 toast（排除 toasts_before）。
        #       ② 不因为"看不懂"就放行 —— 只要新增的 toast 带「失败/错误/不能为空」，
        #          一律按失败处理（宁可让上层去核对发件箱，也不能把真失败当成功）。
        #          只放过一眼就能看出与发信无关的（见 IGNORE_HINTS）。
        # 注：判定为失败并不等于这封没发出去 —— 上层 _outbox_hit(for_verify=True)
        #     会去发件箱复核，孚盟"报错但已落库"的情况会被正确地改判为成功。
        IGNORE_HINTS = ("保存", "模板", "菜单", "头像", "设置", "上传", "导入", "导出")
        fail_txt = ""
        gone = False
        deadline = time.time() + 20
        while time.time() < deadline:
            self._sleep(0.4)
            try:
                toasts = self._visible_dialogs().get("toasts", [])
            except Exception:
                toasts = []
            fresh = [t for t in toasts if t not in toasts_before]     # ← 只看新增的
            hard = [t for t in fresh if ("失败" in t or "错误" in t or "不能为空" in t)]
            # 与发信无关的（模板/菜单/上传…）不判失败，只记日志
            bad = [t for t in hard
                   if not any(k in t for k in IGNORE_HINTS)]
            if bad:
                fail_txt = "；".join(bad)[:120]
                break
            ignored = [t for t in hard if t not in bad]
            if ignored:
                self.log(f"  （忽略与发信无关的提示：{'；'.join(ignored)[:80]}）")
            try:
                if not self._sched_dialog_open():
                    gone = True
                    break
            except Exception:
                gone = True
                break
        if fail_txt:
            raise MXWebError(f"孚盟拒绝定时：{fail_txt}（未提交，可安全重试）")
        if not gone:
            self.log("  ⚠ 定时弹窗 20s 未关闭但无任何报错，按「已提交」处理（不重发），整轮结束后统一核对")
        toasts = self._visible_dialogs().get("toasts", [])
        info.update({"scheduled": True, "when_local": eff_when,
                     "when_requested": when_local, "country": country,
                     "dialog_gone": gone, "toasts": toasts})
        return info

    def back_to_home(self, timeout: float = 30.0) -> bool:
        """定时写完后回到孚盟初始页面（邮件首页），并**一直等到真的回去了**才返回。

        这是串行处理的关键：不回到首页就不开始下一个客户，避免写信页越堆越多。
        """
        try:
            self.tab.evaluate(
                "(()=>{if(location.hash.indexOf('/mail/home')<0){"
                "location.hash='#/main/mail/home/index';}return location.hash;})()",
                timeout=20)
        except Exception as e:
            self.log("（回首页导航失败）", type(e).__name__)
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                u = str(self.tab.evaluate("location.hash", timeout=8) or "")
            except Exception:
                u = ""
            if "/mail/home" in u:
                self._sleep(0.6 if self.s.fast else 1.5)   # 让首页 DOM 真正渲染完
                return True
            self._sleep(0.4 if self.s.fast else 1.0)
        self.log("（等待回到初始页面超时）")
        return False

    def open_tab_count(self) -> int:
        """当前浏览器开了几个页面标签（用于观察是否堆积）。"""
        try:
            return len(self.browser.targets())
        except Exception:
            return -1

    def close_other_tabs(self) -> int:
        """关掉除当前以外的所有页面标签，防止标签堆积把孚盟拖崩。失败不影响主流程。"""
        try:
            cur = self.tab.id
            n = 0
            for t in self.browser.targets():
                if t.get("id") != cur:
                    self.browser.close_tab(t["id"])
                    n += 1
            return n
        except Exception as e:
            self.log("（清理多余标签失败，忽略）", type(e).__name__)
            return 0

    def open_folder(self, folder: str) -> int:
        """点开左侧某个文件夹（已发件箱 / 发件箱 / 草稿箱…），返回匹配到的元素数。"""
        return self.tab.evaluate(
            "((txt)=>{const c=Array.from(document.querySelectorAll('div.content-name,span.content-name,li,div'))"
            ".filter(e=>{const r=e.getBoundingClientRect();"
            " return r.width>0&&r.height>0&&(e.innerText||'').trim().indexOf(txt)===0;});"
            "c.sort((a,b)=>a.querySelectorAll('*').length-b.querySelectorAll('*').length);"
            "if(c.length) c[0].click(); return c.length;})("
            + json.dumps(folder, ensure_ascii=False) + ")", timeout=25) or 0

    def scan_sent(self, subject: str = "", max_scroll: int = 12,
                  folder: str = "已发件箱") -> list[str]:
        """扫某个文件夹的邮件列表，返回每条的摘要文本（含收件人前缀+日期+主题）。

        folder="已发件箱"：唯一可信的"真发出"凭据；
        folder="发件箱"：核对定时任务的凭据（定时提交后立刻出现带未来时间的记录）。
        """
        self.open_folder(folder)
        self._settle(3.5, JS_MAIL_ITEMS, lambda v: isinstance(v, int) and v > 0, timeout=6.0)
        seen: list[str] = []
        for _ in range(max_scroll):
            raw = self.tab.evaluate(
                "(()=>{const its=Array.from(document.querySelectorAll('.mail-item'))"
                ".filter(e=>e.getBoundingClientRect().width>100);"
                "return JSON.stringify(its.map(e=>(e.innerText||'').replace(/\\s+/g,' ').trim()));})()",
                timeout=30)
            try:
                items = json.loads(raw)
            except Exception:
                items = []
            for it in items:
                if it and it not in seen:
                    seen.append(it)
            # 往下滚，加载更多
            moved = self.tab.evaluate(
                "(()=>{const its=Array.from(document.querySelectorAll('.mail-item'));"
                "if(!its.length) return 'none';"
                "let p=its[0].parentElement;"
                "for(let i=0;i<4&&p;i++){"
                "  if(p.scrollHeight>p.clientHeight+10){const t=p.scrollTop;p.scrollTop=t+p.clientHeight;"
                "    return 'moved';}"
                "  p=p.parentElement;}"
                "return 'end';})()", timeout=20)
            if moved != "moved":
                break
            self._sleep(0.5 if self.s.fast else 1.0)
        if subject:
            key = subject[:24]
            seen = [t for t in seen if key in t] or seen
        return seen

    def verify_sent(self, to_email: str, subject: str = "",
                    folder: str = "已发件箱", strict_subject: bool = False) -> dict:
        """核对一封邮件是否出现在指定文件夹（默认「已发件箱」=真发出）。

        folder="发件箱" 用来核对定时任务：定时提交成功后，孚盟「发件箱」
        会立刻出现一条带未来时间的定时记录——这是"定时成功"的凭据。
        """
        self.open_folder(folder)
        self._settle(4, JS_MAIL_ITEMS, lambda v: isinstance(v, int) and v > 0, timeout=6.0)
        local = to_email.split("@")[0].strip().lower()
        # 2026-09-22 修正：列表里收件人和主题都是**截断显示**
        # （如「Purchase @leone…」「Invitation to the 137th…」），
        # 旧版还要求主题前 30 字完整出现 + 不分大小写都没做 → 永远匹配不上
        # → 明明已经定时成功却判成"没核对到" → 上层重试 → 重复发送。
        # 现在：只按 @前缀、不区分大小写匹配；主题不参与判定。
        head = self.tab.evaluate(
            "(()=>{const items=Array.from(document.querySelectorAll('.mail-item'))"
            ".filter(e=>e.getBoundingClientRect().width>100);"
            "const out=[];for(const e of items){const t=(e.innerText||'').replace(/\\s+/g,' ').trim();"
            " if(t) out.push(t.slice(0,160)); if(out.length>=8) break;}"
            "return JSON.stringify(out);})()", timeout=30)
        try:
            top = json.loads(head)
        except Exception:
            top = []
        low = [(t or "").lower() for t in top]
        hit_any = any(local in t for t in low)
        # 2026-09-26：strict_subject=True 时还要「这条记录是本轮主题」才算命中。
        # 列表里主题会截断显示（如「Invitation to the 137th…」），所以只比前 12 个字符。
        if strict_subject:
            key = (subject or "").strip()[:12].lower()
            hit = bool(key) and any(local in t and key in t for t in low)
            return {"top": top, "found": hit, "loose": hit_any}
        return {"top": top, "found": bool(top) and hit_any, "loose": hit_any}

    def verify_batch(self, emails: list[str], folder: str = "发件箱",
                     max_scroll: int = 10) -> dict:
        """一次打开文件夹、一次取列表，批量比对哪些邮箱在里头。

        原来每位客户都单独切一次发件箱核对（约 5~6 秒/位），400 多位就是半小时。
        改成整轮跑完统一核对一次，省下来的都是真金白银。
        注意：按 @前缀 匹配，前缀相同的客户无法区分（与单封核对同样的局限）。
        """
        items = self.scan_sent(subject="", folder=folder, max_scroll=max_scroll)
        low = [(t or "").lower() for t in items]
        found, missing = [], []
        for e in emails or []:
            local = (e or "").split("@")[0].strip().lower()
            if not local:
                continue
            (found if any(local in t for t in low) else missing).append(e)
        return {"found": found, "missing": missing,
                "n_items": len(items), "folder": folder}

    def close(self) -> None:
        if self.tab:
            self.tab.close()
