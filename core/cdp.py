# -*- coding: utf-8 -*-
"""
极简 Chrome DevTools Protocol (CDP) 客户端 —— 纯 Python 标准库实现。

为什么不用 playwright / selenium：
  1. 它们都要额外下载浏览器内核（几百 MB），而本机已经有 Edge；
  2. 打包成独立 .exe 时体积会暴涨、且需要联网初始化；
  3. 本脚本只需要"打开页面 + 找元素 + 点/输入 + 读取"这几件事，CDP 足够。

因此这里手写了一个最小 WebSocket 客户端（RFC 6455），通过 Edge 的
--remote-debugging-port 直接对话。依赖为零，可直接打进独立程序。

用法要点：
  - Edge/Chrome 必须用**独立的 --user-data-dir** 启动（不能占用用户正在用的配置），
    否则 remote debugging 会被拒绝，而且会和用户正在开的浏览器抢配置。
  - 首次使用需要用户在弹出的窗口里手工登录一次；之后 profile 目录里保留了
    登录态（Cookie），后续无需再登录。
"""
from __future__ import annotations

import base64
import json
import os
import socket
import struct
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

DEFAULT_PORT = 9333


class CDPError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# 最小 WebSocket 客户端
# ---------------------------------------------------------------------------
class _WS:
    """只实现文本帧收发 + ping/pong + 关闭，够 CDP 用。"""

    def __init__(self, url: str, timeout: float = 30.0):
        self.url = url
        self.sock, self._f = _handshake(url, timeout)
        self.timeout = timeout

    # ---- 底层 ----
    def _read(self, n: int) -> bytes:
        buf = bytearray()
        while len(buf) < n:
            chunk = self._f.read(n - len(buf))
            if not chunk:
                raise CDPError("WebSocket 连接已断开")
            buf += chunk
        return bytes(buf)

    def send(self, text: str, opcode: int = 0x1) -> None:
        data = text.encode("utf-8")
        header = bytearray([0x80 | opcode])
        n = len(data)
        if n < 126:
            header.append(0x80 | n)
        elif n < 65536:
            header.append(0x80 | 126)
            header += struct.pack(">H", n)
        else:
            header.append(0x80 | 127)
            header += struct.pack(">Q", n)
        mask = os.urandom(4)
        header += mask
        payload = bytes(b ^ mask[i % 4] for i, b in enumerate(data))
        self.sock.sendall(bytes(header) + payload)

    def recv(self) -> str:
        while True:
            b0, b1 = self._read(2)
            opcode = b0 & 0x0F
            masked = b1 & 0x80
            n = b1 & 0x7F
            if n == 126:
                n = struct.unpack(">H", self._read(2))[0]
            elif n == 127:
                n = struct.unpack(">Q", self._read(8))[0]
            mask = self._read(4) if masked else b""
            payload = self._read(n) if n else b""
            if masked:
                payload = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
            if opcode == 0x1:          # text
                return payload.decode("utf-8", "replace")
            if opcode == 0x8:          # close
                raise CDPError("WebSocket 被服务端关闭")
            if opcode == 0x9:          # ping -> pong
                self.send(payload.decode("utf-8", "replace"), 0xA)
            # 0x0 续帧 / 0x2 二进制：CDP 用不到，直接忽略继续读

    def close(self) -> None:
        try:
            self.send("", 0x8)
        except Exception:
            pass
        try:
            self.sock.close()
        except Exception:
            pass


def _handshake(url: str, timeout: float):
    if not url.startswith("ws://"):
        raise CDPError(f"不支持的 WebSocket 地址: {url}")
    rest = url[5:]
    hostport, _, path = rest.partition("/")
    path = "/" + path
    host, _, port_s = hostport.partition(":")
    port = int(port_s or 80)
    key = base64.b64encode(os.urandom(16)).decode("ascii")
    req = (
        f"GET {path} HTTP/1.1\r\n"
        f"Host: {hostport}\r\n"
        f"Upgrade: websocket\r\n"
        f"Connection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {key}\r\n"
        f"Sec-WebSocket-Version: 13\r\n\r\n"
    )
    sock = socket.create_connection((host, port), timeout=timeout)
    sock.sendall(req.encode("ascii"))
    f = sock.makefile("rb")
    status = f.readline().decode("latin1")
    if "101" not in status:
        raise CDPError(f"WebSocket 握手失败: {status.strip()}")
    while True:
        line = f.readline()
        if line in (b"\r\n", b"\n", b""):
            break
    sock.settimeout(timeout)
    return sock, f


# ---------------------------------------------------------------------------
# 标签页
# ---------------------------------------------------------------------------
class Tab:
    """一个 CDP 会话（对应浏览器里的一个标签页）。"""

    def __init__(self, ws: _WS, target: dict[str, Any], browser: "Browser"):
        self._ws = ws
        self.target = target
        self.id = target.get("id", "")
        self.browser = browser
        self._mid = 0

    # ---- 基础 ----
    def send(self, method: str, params: dict[str, Any] | None = None,
             timeout: float = 30.0) -> dict[str, Any]:
        self._mid += 1
        msg = json.dumps({"id": self._mid, "method": method,
                          "params": params or {}}, ensure_ascii=False)
        self._ws.sock.settimeout(timeout)
        self._ws.send(msg)
        deadline = time.time() + timeout
        while True:
            raw = self._ws.recv()
            data = json.loads(raw)
            if data.get("id") == self._mid:
                if "error" in data:
                    err = data["error"]
                    raise CDPError(f"{method} 失败: {err.get('message')} "
                                   f"(code={err.get('code')})")
                return data.get("result", {})
            if time.time() > deadline:
                raise CDPError(f"{method} 超时")
            # 其余是事件通知，忽略

    def evaluate(self, expr: str, timeout: float = 30.0) -> Any:
        r = self.send("Runtime.evaluate", {
            "expression": expr,
            "returnByValue": True,
            "awaitPromise": True,
            "userGesture": True,
        }, timeout=timeout)
        if "exceptionDetails" in r:
            d = r["exceptionDetails"]
            txt = d.get("text", "")
            ex = (d.get("exception") or {}).get("description") or (d.get("exception") or {}).get("value")
            raise CDPError(f"JS 执行异常 {txt}: {ex}")
        return (r.get("result") or {}).get("value")

    # ---- 导航 ----
    def goto(self, url: str, wait_load: bool = True, timeout: float = 60.0) -> None:
        self.send("Page.enable", timeout=10)
        self.send("Page.navigate", {"url": url}, timeout=timeout)
        if wait_load:
            self.wait_js("document.readyState === 'complete' || document.readyState === 'interactive'",
                         timeout=timeout)

    def reload(self, ignore_cache: bool = True, timeout: float = 60.0) -> bool:
        """刷新当前页面（等价于手工按 F5），返回是否等到加载完成。

        为什么需要它：Edge 长时间挂机后标签页可能被冻结/丢弃，或单页应用（孚盟）
        的 JS 卡住，表现是页面不动、evaluate 一直超时。这种"页面僵死"光靠重连
        CDP 是救不回来的 —— 只有刷新页面才行（用户平时就是手工按 F5 解决的）。
        """
        try:
            self.send("Page.enable", timeout=10)
        except Exception:
            pass
        try:
            self.send("Page.reload", {"ignoreCache": bool(ignore_cache)}, timeout=25)
        except Exception:
            # 少数情况下 reload 命令本身会超时，退回到重新导航当前地址
            try:
                u = self.target.get("url") or ""
                if u and u != "about:blank":
                    self.send("Page.navigate", {"url": u}, timeout=timeout)
                else:
                    return False
            except Exception:
                return False
        return self.wait_js(
            "document.readyState === 'complete' || document.readyState === 'interactive'",
            timeout=timeout)

    def wait_js(self, expr: str, timeout: float = 60.0, interval: float = 0.5) -> bool:
        """轮询执行 JS 表达式，直到返回真值。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                if self.evaluate(f"!!({expr})", timeout=15):
                    return True
            except CDPError:
                pass
            time.sleep(interval)
        return False

    @property
    def url(self) -> str:
        try:
            return self.evaluate("location.href", timeout=10) or ""
        except CDPError:
            return self.target.get("url", "")

    @property
    def title(self) -> str:
        try:
            return self.evaluate("document.title", timeout=10) or ""
        except CDPError:
            return self.target.get("title", "")

    # ---- 交互 ----
    def rect_of(self, selector: str, index: int = 0, frame: int | None = None) -> dict | None:
        js = """
        ((sel, idx, inFrame) => {
          let doc = document;
          if (inFrame !== null) {
            const fs = document.querySelectorAll('iframe');
            if (!fs[inFrame]) return null;
            try { doc = fs[inFrame].contentDocument; } catch(e) { return null; }
          }
          const els = doc.querySelectorAll(sel);
          const el = els[idx];
          if (!el) return null;
          el.scrollIntoView({block:'center'});
          const r = el.getBoundingClientRect();
          return {x: r.x + r.width/2, y: r.y + r.height/2,
                  w: r.width, h: r.height, visible: r.width>0 && r.height>0};
        })
        """
        return self.evaluate(f"({js})({json.dumps(selector, ensure_ascii=False)}, {index}, "
                             f"{'null' if frame is None else frame})", timeout=15)

    def click(self, selector: str, index: int = 0, frame: int | None = None,
              settle: float = 0.6) -> bool:
        r = self.rect_of(selector, index, frame)
        if not r or not r.get("visible"):
            return False
        x, y = float(r["x"]), float(r["y"])
        for etype in ("mouseMoved", "mousePressed", "mouseReleased"):
            self.send("Input.dispatchMouseEvent", {
                "type": etype, "x": x, "y": y,
                "button": "left", "clickCount": 1, "buttons": 1 if etype != "mouseReleased" else 0,
            }, timeout=15)
        time.sleep(settle)
        return True

    def hover(self, selector: str, index: int = 0, frame: int | None = None,
              hold: float = 0.8, warm: float = 0.0) -> bool:
        """把鼠标移到元素上悬停一会儿（不点击）。

        2026-10-12：新版孚盟的「写邮件」模板入口是**悬停展开**的
        （鼠标放上去等一下自动弹出，没有下拉箭头可点）。
        这里用真实 CDP 鼠标事件派发 mouseMoved —— Vue 的 hover 指令
        监听真实鼠标事件，JS 合成事件不一定触发，所以必须走 Input 域。

        warm：先在这个点上做几次微小抖动。有些 el-popper 要求「鼠标真的
        在动」才认为用户悬停意图明确，只发一次 mouseMoved 会被忽略。
        """
        r = self.rect_of(selector, index, frame)
        if not r or not r.get("visible"):
            return False
        x, y = float(r["x"]), float(r["y"])
        # 先从页面外的一个安全点移入，保证是一次完整的 enter（不是瞬移）
        self.send("Input.dispatchMouseEvent", {
            "type": "mouseMoved", "x": 2.0, "y": 2.0, "button": "left", "buttons": 0,
        }, timeout=15)
        time.sleep(0.12)
        self.send("Input.dispatchMouseEvent", {
            "type": "mouseMoved", "x": x, "y": y, "button": "left", "buttons": 0,
        }, timeout=15)
        if warm > 0:
            for k in range(3):
                time.sleep(0.08)
                self.send("Input.dispatchMouseEvent", {
                    "type": "mouseMoved", "x": x + (1 if k % 2 else -1), "y": y,
                    "button": "left", "buttons": 0,
                }, timeout=15)
        if hold > 0:
            time.sleep(hold)
        return True

    def click_text(self, text: str, tags: str = "button,a,span,div,li,[role=button]",
                   frame: int | None = None, settle: float = 0.6) -> bool:
        """按可见文本点元素（中文界面最常用）。"""
        js = """
        ((txt, tags, inFrame) => {
          let doc = document;
          if (inFrame !== null) {
            const fs = document.querySelectorAll('iframe');
            if (!fs[inFrame]) return false;
            try { doc = fs[inFrame].contentDocument; } catch(e) { return false; }
          }
          const els = doc.querySelectorAll(tags);
          for (const el of els) {
            const t = (el.innerText || el.textContent || '').trim();
            if (t === txt || t.includes(txt)) {
              const r = el.getBoundingClientRect();
              if (r.width <= 0 || r.height <= 0) continue;
              el.scrollIntoView({block:'center'});
              const r2 = el.getBoundingClientRect();
              return {x: r2.x + r2.width/2, y: r2.y + r2.height/2};
            }
          }
          return false;
        })
        """
        pos = self.evaluate(f"({js})({json.dumps(text, ensure_ascii=False)}, "
                            f"{json.dumps(tags)}, {'null' if frame is None else frame})", timeout=15)
        if not pos:
            return False
        x, y = float(pos["x"]), float(pos["y"])
        for etype in ("mouseMoved", "mousePressed", "mouseReleased"):
            self.send("Input.dispatchMouseEvent", {
                "type": etype, "x": x, "y": y,
                "button": "left", "clickCount": 1, "buttons": 1 if etype != "mouseReleased" else 0,
            }, timeout=15)
        time.sleep(settle)
        return True

    def fill(self, selector: str, text: str, index: int = 0, frame: int | None = None,
             clear: bool = True) -> bool:
        """聚焦元素后插入文本（走的是真实输入事件，React/Vue 受控组件也能收到）。"""
        if not self.click(selector, index, frame, settle=0.2):
            return False
        if clear:
            self.send("Input.dispatchKeyEvent", {
                "type": "keyDown", "key": "a", "code": "KeyA",
                "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65,
                "modifiers": 2,   # Ctrl
            }, timeout=10)
            self.send("Input.dispatchKeyEvent", {
                "type": "keyUp", "key": "a", "code": "KeyA",
                "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65,
                "modifiers": 2,
            }, timeout=10)
            self.send("Input.dispatchKeyEvent", {
                "type": "keyDown", "key": "Delete", "code": "Delete",
                "windowsVirtualKeyCode": 46, "nativeVirtualKeyCode": 46,
            }, timeout=10)
            self.send("Input.dispatchKeyEvent", {
                "type": "keyUp", "key": "Delete", "code": "Delete",
                "windowsVirtualKeyCode": 46, "nativeVirtualKeyCode": 46,
            }, timeout=10)
        if text:
            self.send("Input.insertText", {"text": text}, timeout=15)
        return True

    def window_id(self) -> int:
        r = self.send("Browser.getWindowForTarget", {"targetId": self.id}, timeout=15)
        return r["windowId"]

    def set_window_state(self, state: str = "maximized") -> bool:
        """恢复/最大化/最小化窗口。

        坑：用 DETACHED_PROCESS 启动的 Edge 窗口会以最小化形态出现，
        此时鼠标事件全部落空、Page.captureScreenshot 会一直挂住（表现是超时）。
        交互前必须先把窗口恢复成正常或最大化。
        """
        try:
            wid = self.window_id()
            self.send("Browser.setWindowBounds", {"windowId": wid,
                                                  "bounds": {"windowState": state}}, timeout=15)
            return True
        except CDPError as exc:
            raise CDPError(f"设置窗口状态失败: {exc}")

    def screenshot(self, path: str) -> str:
        r = self.send("Page.captureScreenshot", {"format": "png"}, timeout=30)
        import binascii
        with open(path, "wb") as f:
            f.write(binascii.a2b_base64(r["data"]))
        return path

    def close(self) -> None:
        try:
            self.browser.close_tab(self.id)
        except Exception:
            pass
        self._ws.close()


# ---------------------------------------------------------------------------
# 浏览器
# ---------------------------------------------------------------------------
class BrowserSession:
    """浏览器级会话（连 /json/version 的 WS），用于窗口控制等 Browser.* 域命令。"""

    def __init__(self, ws: _WS):
        self._ws = ws
        self._mid = 0

    def send(self, method: str, params: dict[str, Any] | None = None,
             timeout: float = 20.0) -> dict[str, Any]:
        self._mid += 1
        self._ws.sock.settimeout(timeout)
        self._ws.send(json.dumps({"id": self._mid, "method": method,
                                  "params": params or {}}, ensure_ascii=False))
        while True:
            data = json.loads(self._ws.recv())
            if data.get("id") == self._mid:
                if "error" in data:
                    raise CDPError(f"{method} 失败: {data['error'].get('message')}")
                return data.get("result", {})

    def close(self) -> None:
        self._ws.close()


# ---------------------------------------------------------------------------
# Windows 专用：把窗口"恢复但不激活"
# ---------------------------------------------------------------------------
_SW_SHOWNOACTIVATE = 4


def _win_restore_no_activate(pid: int) -> int:
    """把属于 pid 的浏览器顶层窗口从最小化恢复出来，但**不激活**。

    2026-10-06 真机取证：
      · CDP 的 Browser.setWindowBounds 恢复最小化窗口时，Chromium 会顺手把它
        激活到前台 —— 实测调用后 1.2 秒，前台窗口就从别的程序变成了 Edge。
        脚本每发一封都会恢复一次窗口，用户开着别的程序干活时就被反复打断，
        表现为"我一打开窗口，几秒就没了"。
      · 改用 Win32 ShowWindow(SW_SHOWNOACTIVATE) 后：窗口回到普通大小、
        页面 document.hidden 变回 false（CDP 真实鼠标点击、截图照常可用），
        而**前台窗口纹丝不动**（实测前台仍是原来的程序）。

    返回实际恢复的窗口数；非 Windows / 找不到窗口时返回 0（调用方自行兜底）。
    """
    if os.name != "nt" or not pid:
        return 0
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return 0
    u = ctypes.windll.user32
    hits: list[int] = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def _cb(hwnd, _lparam):
        owner = wintypes.DWORD()
        u.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
        if owner.value != pid:
            return True
        cls = ctypes.create_unicode_buffer(256)
        u.GetClassNameW(hwnd, cls, 256)
        # Edge/Chrome 的顶层窗口都是 Chrome_WidgetWin_1（个别版本 _0）
        if not cls.value.startswith("Chrome_WidgetWin"):
            return True
        hits.append(int(hwnd))
        return True

    try:
        u.EnumWindows(WNDENUMPROC(_cb), 0)
    except Exception:
        return 0
    n = 0
    for hwnd in hits:
        try:
            if u.IsIconic(hwnd):          # 只动最小化的；普通窗口绝不碰
                u.ShowWindow(hwnd, _SW_SHOWNOACTIVATE)
                n += 1
        except Exception:
            pass
    return n


class Browser:
    def __init__(self, port: int = DEFAULT_PORT):
        self.port = port
        self._pid_cache = 0          # 浏览器主进程 PID（懒查一次，见 browser_pid）

    def session(self) -> BrowserSession:
        return BrowserSession(_WS(self.version()["webSocketDebuggerUrl"]))

    def window_state(self, target_id: str) -> str:
        """查询窗口状态：normal / minimized / maximized / fullscreen（查不到返回空）。"""
        s = self.session()
        try:
            r = s.send("Browser.getWindowForTarget", {"targetId": target_id})
            b = s.send("Browser.getWindowBounds", {"windowId": r["windowId"]})
            return (b.get("bounds") or {}).get("windowState", "") or ""
        except Exception:
            return ""
        finally:
            s.close()

    def set_window_state(self, target_id: str, state: str = "maximized") -> bool:
        """恢复/最大化指定标签页所在窗口（绕过页面会话，直接走浏览器级 WS）。"""
        s = self.session()
        try:
            r = s.send("Browser.getWindowForTarget", {"targetId": target_id})
            wid = r["windowId"]
            # 坑：最小化的窗口不能直接最大化，必须先恢复成 normal
            if state in ("maximized", "fullscreen"):
                try:
                    s.send("Browser.setWindowBounds",
                           {"windowId": wid, "bounds": {"windowState": "normal"}})
                    time.sleep(0.6)
                except CDPError:
                    pass
            s.send("Browser.setWindowBounds",
                   {"windowId": wid, "bounds": {"windowState": state}})
            return True
        finally:
            s.close()

    # ---- 恢复窗口但不抢前台（2026-10-06）----------------------------------
    def browser_pid(self) -> int:
        """浏览器主进程 PID（用来定位它自己的窗口）。取不到返回 0。"""
        if self._pid_cache:
            return self._pid_cache
        s = self.session()
        try:
            r = s.send("SystemInfo.getProcessInfo", {})
            for p in (r.get("processInfo") or []):
                if str(p.get("type") or "") == "browser":
                    self._pid_cache = int(p.get("id") or 0)
                    return self._pid_cache
        except Exception:
            pass
        finally:
            s.close()
        return 0

    def restore_no_activate(self, target_id: str = "") -> int:
        """把最小化的浏览器窗口恢复成普通窗口，**但不激活、不抢前台**。

        返回实际恢复了几扇窗口；0 表示没恢复（调用方应回退到
        set_window_state，虽然那个会抢前台）。
        """
        pid = self.browser_pid()
        if not pid:
            return 0
        return _win_restore_no_activate(pid)

    # ---- HTTP 端点 ----
    def _http(self, path: str, method: str = "GET", timeout: float = 8.0) -> Any:
        url = f"http://127.0.0.1:{self.port}{path}"
        req = urllib.request.Request(url, method=method)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def alive(self) -> bool:
        try:
            self._http("/json/version", timeout=3)
            return True
        except Exception:
            return False

    def version(self) -> dict:
        return self._http("/json/version")

    def targets(self) -> list[dict]:
        try:
            return [t for t in self._http("/json/list") if t.get("type") == "page"]
        except Exception:
            return []

    def close_tab(self, target_id: str) -> None:
        try:
            self._http(f"/json/close/{target_id}", timeout=5)
        except Exception:
            pass

    def new_tab(self, url: str = "about:blank") -> Tab:
        path = "/json/new?" + urllib.parse.quote(url, safe="")
        info = self._http(path, method="PUT", timeout=20)
        return self.attach(info)

    def first_tab(self) -> Tab | None:
        ts = self.targets()
        if not ts:
            return None
        return self.attach(ts[0])

    def attach(self, target: dict) -> Tab:
        ws = _WS(target["webSocketDebuggerUrl"])
        return Tab(ws, target, self)


# ---------------------------------------------------------------------------
# 启动 / 查找浏览器
# ---------------------------------------------------------------------------
def find_browser() -> str:
    """按 Edge -> Chrome 顺序找本机浏览器可执行文件。"""
    import winreg

    cands: list[str] = []
    # 1) 注册表 App Paths
    for root, sub in (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
    ):
        try:
            with winreg.OpenKey(root, sub) as k:
                cands.append(winreg.QueryValue(k, None))
        except OSError:
            pass
    # 2) 常见安装路径
    for base in (os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                 os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("LOCALAPPDATA", "")):
        for rel in (r"Microsoft\Edge\Application\msedge.exe",
                    r"Google\Chrome\Application\chrome.exe"):
            p = os.path.join(base, rel)
            if base:
                cands.append(p)
    for p in cands:
        if p and os.path.isfile(p):
            return p
    raise CDPError("没找到 Edge 或 Chrome，请手动指定浏览器路径")


def launch(edge_path: str | None = None, port: int = DEFAULT_PORT,
           user_data_dir: str = "", url: str = "about:blank",
           headless: bool = False, wait: float = 40.0,
           extra_args: list[str] | None = None, detach: bool = False) -> Browser:
    """启动带远程调试端口的浏览器；若端口已在监听则直接复用。"""
    b = Browser(port)
    if b.alive():
        return b
    exe = edge_path or find_browser()
    if not user_data_dir:
        raise CDPError("必须指定独立的 --user-data-dir（不能占用用户日常用的浏览器配置）")
    os.makedirs(user_data_dir, exist_ok=True)
    args = [
        exe,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        # ---- 防冻结 / 防后台降频（2026-09-23 加）----
        # 挂机群发时 Edge 会把长时间不活动的标签页"睡眠"或"丢弃"，页面 JS 停摆，
        # 表现就是页面不动、必须手工按 F5 才恢复。下面这些开关直接关掉这套机制。
        "--disable-features=Translate,OptimizationHints,"
        "SleepingTabs,TabDiscarding,TabFreezing,"
        "CalculateNativeWinOcclusion,IntensiveWakeUpThrottling,"
        "HighEfficiencyModeAvailable,WebContentsOcclusion",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-background-timer-throttling",
        "--disable-ipc-flooding-protection",
        "--lang=zh-CN",
        "--window-size=1440,900",
    ]
    if headless:
        args.append("--headless=new")
    args.extend(extra_args or [])
    if url:
        args.append(url)

    # 关键：默认不做进程脱离。DETACHED_PROCESS 会让 Edge 窗口以最小化形态出现，
    # 导致鼠标点击落空、截图挂死 —— 交互前还得再恢复窗口，得不偿失。
    # 只有确实需要"父进程退出后浏览器仍存活"时才 detach=True。
    flags = 0
    if os.name == "nt" and detach:
        flags |= 0x00000008          # DETACHED_PROCESS：不继承控制台
        flags |= 0x00000200          # CREATE_NEW_PROCESS_GROUP
        try:
            subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             close_fds=True, creationflags=flags | 0x01000000)  # BREAKAWAY_FROM_JOB
        except OSError:
            subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             close_fds=True, creationflags=flags)
    elif os.name == "nt":
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         close_fds=True, creationflags=0x00000200)  # CREATE_NEW_PROCESS_GROUP
    else:
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         close_fds=True, start_new_session=True)
    deadline = time.time() + wait
    while time.time() < deadline:
        if b.alive():
            return b
        time.sleep(0.4)
    raise CDPError(f"浏览器调试端口 {port} 未就绪（等待 {wait}s）")
