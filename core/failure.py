# -*- coding: utf-8 -*-
"""失败原因分类器（2026-09-23 新增）

目的：把「发一封失败」这件事拆成可自动恢复的 / 必须人工处理的，
      并给出每一类的重试上限、退避间隔和对应处理动作。

设计原则（血泪教训）：
  1. **提交类动作**（点确定/发送）默认按"疑似成功"处理，失败判定要保守；
  2. **网络类**失败最容易被误判成"客户有问题"——必须单独识别，走长退避；
  3. **数据类**失败（国家名不对、地址无效）重试一万次也没用，直接标人工。
"""
from __future__ import annotations

import re

# ---------------------------------------------------------------- 类别定义
# key: (中文名, 可否自动重发, 建议最大尝试次数, 退避基数秒, 处理动作)
CATS = {
    "net":  ("网络/超时",        True,  5, 15.0,
             "网络恢复后自动重发；连续失败会挂起等待，不会跳过客户"),
    "ui":   ("页面未就绪",      True,  3,  4.0,
             "孚盟页面偶发没渲染出来，自动重试；多次仍失败则标人工"),
    "data": ("数据问题",        False, 1,  0.0,
             "需人工修正客户表（补国家名/改地址）后再补发"),
    "perm": ("服务端拒绝",      False, 1,  0.0,
             "孚盟/对方服务器拒绝（附件超限、认证失败、退信），需人工确认"),
    "skip": ("已跳过/无需重发", False, 1,  0.0,
             "无需处理"),
    "other": ("未知原因",       True,  2,  8.0,
             "未在白名单内的错误，少量重试后标人工"),
}

# ------------------------------------------------------- 关键词 → 类别
# 顺序敏感：越具体的放前面
_RULES: list[tuple[str, str, str]] = [
    # ---- 网络 / 超时 / 连接 ----
    (r"timed?\s*out|超时", "net", "网络超时，孚盟页面没在规定时间内响应"),
    # 2026-09-30：孚盟服务器投递时报「SocketException: Connection reset」——
    # 发太快/附件太大触发的断连，必须归网络类走长退避，不能当"服务端拒绝"放弃。
    # 注意要放在 MessagingException 规则之前（那条规定会把整串文本抢先吃掉）。
    (r"connection\s*reset|连接被重置|SocketException", "net",
     "连接被服务器重置（多为发送太快/附件太大触发限制），退避后重试"),
    (r"自动重连失败", "net", "与孚盟的连接断开，自动重连也没连上"),
    (r"cannot read from timed out", "net", "浏览器通道读超时（网络中断的典型表现）"),
    (r"ConnectionReset|WinError 10054|远程主机", "net", "连接被远程主机强制关闭"),
    (r"ConnectionRefused|WinError 10061", "net", "连接被拒绝（孚盟服务短暂不可用）"),
    (r"ConnectionAborted|BrokenPipe|WinError 10053", "net", "连接被中止"),
    (r"socket|Socket|websocket|WebSocket", "net", "底层通道异常"),
    (r"连不上孚盟|连接已断开|已断开", "net", "与孚盟的连接断开"),
    (r"Target closed|target closed", "net", "浏览器标签被关闭"),
    (r"NameResolution|getaddrinfo|DNS", "net", "域名解析失败（网络问题）"),

    # ---- 服务端 / 永久拒绝 ----
    (r"MessagingException", "perm", "孚盟服务端报错（常见于附件过大或服务端 IO 异常）"),
    (r"附件|超过.*[大限]|尺寸超限|too large", "perm", "附件过大被拒（建议改小附件或改用外链）"),
    (r"认证|鉴权|Authentication|auth fail|401|403", "perm", "发送账号认证/授权失败"),
    (r"退信|被拒收|bounce|rejected|550|554|553", "perm", "对方服务器拒收（地址不存在或被拉黑）"),
    (r"限流|频率|too many|429|发送过于频繁", "perm", "触发平台限流，需放缓节奏"),
    (r"余额|配额|quota|超出.*上限", "perm", "账号配额/余额不足"),

    # ---- 数据问题（改客户表才能解决）----
    (r"没找到「.+?」|国家列表里没找到", "data", "客户表里的国家名孚盟搜不到（如「欧盟」不是国家，要写具体国家）"),
    (r"缺国家|国家列为空|country.*空", "data", "客户表该行国家为空，需补上国家"),
    (r"地址无效|邮箱.*无效|不是有效.*地址|Invalid.*address|malformed", "data",
     "邮箱地址格式无效，需人工核对客户表"),

    # ---- 页面未就绪（重试多半能过）----
    (r"下拉/搜索框没打开|下拉没打开", "ui", "孚盟国家下拉没渲染出来（页面偶发）"),
    (r"国家没选上|回显为空", "ui", "国家已点但回显没生效（页面偶发）"),
    (r"找不到菜单|NOT_FOUND", "ui", "孚盟菜单没渲染出来（常发生在页面刚跳转时）"),
    (r"弹窗没有打开|没有收件人", "ui", "定时弹窗没打开（多为收件人未就位或页面未渲染完）"),
    (r"当地时间没写进去|框里是", "ui", "时间输入框内容异常（页面偶发）"),
    (r"模板下拉里没有", "ui", "模板列表没加载出来（页面偶发）"),
    (r"没点到|点击失败", "ui", "按钮没点到（页面遮挡或未渲染）"),
    (r"回到初始页面超时|回首页导航失败", "ui", "发送后没回到初始页（多为页面卡顿）"),
]


# 这些不是主权国家，孚盟国家列表里必然搜不到 → 真·数据问题，改客户表才行
NOT_A_COUNTRY = {
    "欧盟", "欧洲", "欧洲联盟", "eu", "欧州", "中东", "东南亚", "非洲", "南美",
    "北美", "拉丁美洲", "中南美", "亚洲", "西欧", "东欧", "中东欧", "大洋洲",
    "全球", "其他", "其它", "未知", "不详", "none", "null", "无", "",
}


def classify(msg: str = "", err: BaseException | None = None) -> dict:
    """把一条失败信息归类。

    返回 dict：
      cat      类别 key（net/ui/data/perm/other）
      name     中文名（网络/超时 …）
      retry    是否可自动重发
      max_try  建议最大尝试次数
      backoff  退避基数秒
      action   处理动作说明
      reason   这条失败的具体原因（可读）
    """
    text = (msg or "")
    if err is not None:
        text = f"{text} {type(err).__name__}: {err}"
    text = text.strip()

    cat, reason = "other", (text[:120] or "未记录错误信息")
    for pat, c, r in _RULES:
        if re.search(pat, text, flags=re.I):
            cat, reason = c, r
            break

    name, retry, max_try, backoff, action = CATS[cat]

    # "国家列表没找到 X" 要分辨两种情况：
    #   ① 消息里同时提到下拉没打开 → 列表压根没加载，属页面问题，可重试
    #   ② X 本身不是国家（如「欧盟」）→ 真数据问题，必须改客户表
    #   ③ 除①②外（如「美国」这种孚盟肯定有的国家）→ 多半是列表没加载的误报，
    #      按页面问题重试，否则会被误判成"数据问题"而放弃本可成功的客户
    if cat == "data":
        m = re.search(r"没找到「(.+?)」", text)
        what = (m.group(1) if m else "").strip().lower()
        if re.search(r"下拉|NOT_FOUND|没打开", text) or what not in NOT_A_COUNTRY:
            cat, name, retry, max_try, backoff, action = "ui", *CATS["ui"]
            reason = ("国家列表没加载出来（页面偶发），并非客户表数据错误；" + reason)

    return {"cat": cat, "name": name, "retry": bool(retry), "max_try": max_try,
            "backoff": backoff, "action": action, "reason": reason,
            "raw": text[:200]}


def backoff_for(cat: str, attempt: int) -> float:
    """第 attempt 次失败后应等待的秒数（指数退避，封顶 120 秒）。"""
    base = CATS.get(cat, CATS["other"])[3]
    if base <= 0:
        return 0.0
    return min(base * (2 ** max(0, attempt - 1)), 120.0)


def is_net(msg: str = "", err: BaseException | None = None) -> bool:
    return classify(msg, err)["cat"] == "net"
