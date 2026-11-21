# -*- coding: utf-8 -*-
"""定时发送队列：任务持久化、重启恢复、失败重试、去重、留痕。

设计：
  - 任务存 `data/schedule.json`（人类可读，可直接检查/手改）
  - 时间一律用**北京时间墙上时钟**（YYYY-MM-DD HH:MM）。时区换算交给孚盟MX
    端（用户明确说了 MX 里可以选时区），本地不做任何时区数学。
  - 执行模式 mode：
      "send"  到点后立即走一遍普通发送（写日志+已发件箱核对）
      "mx"    到点后走孚盟的「定时发送」弹窗（由 MX 服务器端定时，程序关了也能发）
  - 每次状态变化都追加一行到 `data/send_log.jsonl`（留痕，可追溯）
  - 去重三层：已发过(sent_history) / 同邮箱同主题同日的 pending 任务 / 发送前再查一次
  - 失败重试：最多 max_attempts 次，间隔 retry_delay 分钟
  - 崩溃恢复：上次退出时停在 "sending" 的任务，启动时回到 "pending"
    （即使上一进程其实已发出，发送前的 already_sent 检查也会拦住重复发送）
"""
from __future__ import annotations

import json
import threading
import time
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

CN = timezone(timedelta(hours=8))


def now_cn() -> datetime:
    return datetime.now(CN)


def parse_when(s: str) -> datetime:
    """'2026-09-22 09:00' / '2026/9/22 9:00' / '明天9:00' 这类都试着认。"""
    s = str(s or "").strip().replace("，", " ").replace("/", "-")
    if not s:
        raise ValueError("时间不能为空")
    if s.startswith("明天"):
        h, m = 9, 0
        rest = s[2:].replace("点", ":")
        if ":" in rest:
            parts = rest.split(":")
            h = int(parts[0] or 9)
            m = int(parts[1] if len(parts) > 1 and parts[1] else 0)
        base = now_cn() + timedelta(days=1)
        return base.replace(hour=h, minute=m, second=0, microsecond=0)
    s = s.replace("点", ":00")
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%m-%d %H:%M", "%d %H:%M", "%H:%M"):
        try:
            dt = datetime.strptime(s, fmt)
            if fmt == "%m-%d %H:%M":
                dt = dt.replace(year=now_cn().year)
            elif fmt == "%d %H:%M":
                dt = dt.replace(year=now_cn().year, month=now_cn().month)
            elif fmt == "%H:%M":
                base = now_cn()
                dt = base.replace(hour=dt.hour, minute=dt.minute, second=0, microsecond=0)
                if dt <= base:
                    dt += timedelta(days=1)       # 只给了几点：已过就排明天
            return dt
        except ValueError:
            continue
    raise ValueError("认不出这个时间写法：" + s)


class TaskStore:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.tasks: dict[str, dict] = {}
        self._lock = threading.Lock()
        self._load()

    def _load(self):
        if self.path.exists():
            try:
                self.tasks = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception:
                self.tasks = {}

    def save(self):
        with self._lock:
            try:
                self.path.write_text(
                    json.dumps(self.tasks, ensure_ascii=False, indent=1), encoding="utf-8")
            except Exception:
                pass

    def add(self, task: dict) -> str:
        tid = task.get("id") or uuid.uuid4().hex[:12]
        task["id"] = tid
        with self._lock:
            self.tasks[tid] = task
        self.save()
        return tid

    def update(self, tid: str, **kw):
        with self._lock:
            t = self.tasks.get(tid)
            if t:
                t.update(kw)
        self.save()

    def get(self, tid: str) -> dict | None:
        return self.tasks.get(tid)

    def all(self) -> list[dict]:
        return list(self.tasks.values())

    def recover_interrupted(self, log=print):
        """崩溃恢复：把上代进程停在 sending 的任务放回 pending。"""
        n = 0
        for t in self.tasks.values():
            if t.get("status") == "sending":
                t["status"] = "pending"
                t["note"] = "上次程序退出时正在发这封，已放回队列（发送前会再查重，不会重复发）"
                n += 1
        if n:
            self.save()
            log(f"恢复了 {n} 个上次没发完的定时任务")
        return n


class Scheduler(threading.Thread):
    """后台线程：每 15 秒扫一次到期任务。daemon=True，随界面退出。"""

    def __init__(self, engine, store: TaskStore, log=print,
                 max_attempts: int = 3, retry_delay_min: float = 10.0):
        super().__init__(daemon=True, name="scheduler")
        self.engine = engine
        self.store = store
        self.log = log
        self.max_attempts = max_attempts
        self.retry_delay_min = retry_delay_min
        self._stop = False
        self.paused = False
        self.halt = False              # 「停止」按钮：当前这封收尾后不再取下一个任务
        self.on_change = None          # 界面回调（刷新列表用）
        self._busy = False
        self._seq = 1                  # 串行计数：第几个客户

    def toggle_pause(self) -> bool:
        self.paused = not self.paused
        if not self.paused:
            self.halt = False          # 手动恢复时清掉停止标记
        return self.paused

    def request_stop(self):
        """停止按钮：处理完当前这一封就不再取新任务，并把调度一并暂停。

        已入队的任务保留在 schedule.json 里，点「恢复调度」后继续。
        """
        self.halt = True
        self.paused = True
        self.log("[定时] 已请求停止：当前这封收尾后暂停队列（任务保留，点「恢复调度」继续）")
        self._notify()

    # ------------- 留痕 -------------
    def _trace(self, task: dict, event: str, **kw):
        row = {"time": now_cn().strftime("%Y-%m-%d %H:%M:%S"),
               "id": task.get("id"), "email": task.get("email"),
               "event": event, **kw}
        try:
            p = Path(self.store.path).parent / "send_log.jsonl"
            with open(p, "a", encoding="utf-8") as f:
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
        except Exception:
            pass
        self.log(f"[定时] {task.get('email')} {event} " +
                 (" ".join(f"{k}={v}" for k, v in kw.items() if v != "")))

    # ------------- 任务生成 -------------
    def add_batch(self, rows: list[dict], when: datetime,
                  subject: str, template: str, mode: str = "send",
                  mx_local_time: str = "") -> dict:
        """rows: [{email,name,country,...}]

        mode="send": 到点后立即发送（when=北京时间执行时刻）
        mode="mx":   立即把任务交给孚盟「定时发送」，mx_local_time=收件人当地时间
                     （YYYY/MM/DD HH:MM），时区由孚盟按国家地区换算
        """
        added = dup = sent = 0
        self.halt = False               # 新任务入队 = 明确要跑，清掉停止标记
        day = when.strftime("%Y-%m-%d")
        pending_keys = {(t.get("email"), t.get("subject"), t.get("day"))
                        for t in self.store.all() if t.get("status") in ("pending", "sending")}
        for r in rows:
            email = (r.get("email") or "").strip().lower()
            if not email:
                continue
            if self.engine.already_sent(email):
                sent += 1
                continue
            if (email, subject, day) in pending_keys:
                dup += 1
                continue
            self.store.add({"email": email, "name": r.get("name", ""),
                            "country": r.get("country", ""),
                            "run_at": when.strftime("%Y-%m-%d %H:%M"),
                            "day": day, "subject": subject, "template": template,
                            "mode": mode, "mx_local_time": mx_local_time,
                            "status": "pending",
                            "attempts": 0, "max_attempts": self.max_attempts,
                            "last_error": "", "created": now_cn().strftime("%Y-%m-%d %H:%M")})
            pending_keys.add((email, subject, day))
            added += 1
        self.log(f"[定时] 新增 {added} 个任务（已发过跳过 {sent}，重复任务跳过 {dup}），"
                 f"发送时间 {when.strftime('%Y-%m-%d %H:%M')}（北京时间）")
        self._notify()
        return {"added": added, "dup": dup, "sent": sent}

    def _notify(self):
        try:
            if self.on_change:
                self.on_change()
        except Exception:
            pass

    # ------------- 主循环 -------------
    def stop(self):
        self._stop = True

    def run(self):
        self.log("[定时] 调度线程已启动（每 15 秒检查一次到期任务）")
        while not self._stop:
            try:
                self.tick()
            except Exception as e:
                self.log("[定时] 轮询出错:", type(e).__name__, e)
            for _ in range(15):
                if self._stop:
                    return
                time.sleep(1)

    def tick(self):
        if self._busy or self.paused:
            return
        if self.engine is None:
            return          # 还没连上孚盟：任务留在队列里，连上后自然继续
        due = [t for t in self.store.all()
               if t.get("status") == "pending" and self._due(t)]
        for i, t in enumerate(due):
            # 暂停/停止必须立刻生效：每个任务开始前都再看一眼
            if self._stop or self.halt or self.paused or self._busy:
                return
            if i == 0 and due:
                self.log(f"[调度] 开始逐个处理，本次共 {len(due)} 位客户"
                         "（严格串行：一位走完全流程并回到初始页，才处理下一位）")
            self._run_task(t)

    @staticmethod
    def _due(t: dict) -> bool:
        try:
            # 任务里存的是北京墙上时钟（无时区标记），比较时也要用无标记的
            now = datetime.now(CN).replace(tzinfo=None)
            run_at = str(t["run_at"])
            fmt = "%Y-%m-%d %H:%M:%S" if len(run_at) > 16 else "%Y-%m-%d %H:%M"
            return now >= datetime.strptime(run_at, fmt)
        except Exception:
            return False

    def _run_task(self, t: dict):
        if t.get("status") != "pending":
            return
        if self.halt or self.paused or self._stop:
            return
        email = t.get("email", "")
        # 发送前最后一道查重（哪怕重启恢复的也一样拦）
        if self.engine.already_sent(email):
            self.store.update(t["id"], status="done", result="跳过：已发过")
            self._trace(t, "skip-duplicate")
            self._notify()
            return
        self.store.update(t["id"], status="sending", attempts=t.get("attempts", 0) + 1)
        self._trace(t, "start", attempt=t.get("attempts", 0) + 1, mode=t.get("mode"))
        t0 = time.time()
        self._busy = True
        try:
            if t.get("mode") == "mx":
                res = self.engine.send_scheduled_mx(
                    email, t.get("subject", ""), t.get("template", ""),
                    t.get("mx_local_time", ""), t.get("country", ""))
            else:
                res = self.engine.send_one(email, t.get("subject", ""), t.get("template", ""))
        except Exception as e:
            res = {"ok": False, "msg": f"{type(e).__name__}: {e}"}
        finally:
            self._busy = False
        sec = round(time.time() - t0, 1)
        ok = bool(res.get("ok"))
        attempts = t.get("attempts", 0) + 1
        if ok:
            self.store.update(t["id"], status="done", result=res.get("msg", ""),
                              finished=now_cn().strftime("%Y-%m-%d %H:%M:%S"))
            self._trace(t, "done", msg=res.get("msg", ""), sec=sec)
            self.log(f"[调度] 第 {self._seq} 个完成 ✓ {email} 用时 {sec}s — {res.get('msg', '')[:60]}")
        elif attempts >= t.get("max_attempts", self.max_attempts):
            self.store.update(t["id"], status="failed", result=res.get("msg", ""),
                              finished=now_cn().strftime("%Y-%m-%d %H:%M:%S"))
            self._trace(t, "failed-final", msg=res.get("msg", ""), sec=sec)
            self.log(f"[调度] 第 {self._seq} 个最终失败 ✗ {email} 用时 {sec}s — {res.get('msg', '')[:80]}"
                     f"（已重试 {attempts - 1} 次）")
        else:
            nxt = (now_cn() + timedelta(minutes=self.retry_delay_min)).strftime("%Y-%m-%d %H:%M:%S")
            self.store.update(t["id"], status="pending", run_at=nxt,
                              last_error=res.get("msg", ""))
            self._trace(t, "retry-scheduled", msg=res.get("msg", ""), next=nxt, sec=sec)
            self.log(f"[调度] 第 {self._seq} 个失败，{self.retry_delay_min} 分钟后重试"
                     f"（第 {attempts}/{t.get('max_attempts', self.max_attempts)} 次）{email} — {res.get('msg', '')[:60]}")
        self._seq += 1
        self._notify()
