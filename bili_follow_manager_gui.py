#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""B站关注列表清理与分析工具（图形界面）"""

import json
import ctypes
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import unicodedata
from tkinter import messagebox, scrolledtext, ttk
from openpyxl import load_workbook

import bili_follow_manager as bm

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(BASE_DIR, "bili_follow_manager.py")
DATA_DIR = os.path.join(BASE_DIR, "data")
COOKIE_FILE = os.path.join(DATA_DIR, "cookies.json")
FOLLOW_FILE = os.path.join(DATA_DIR, "followings.json")
ANALYSIS_FILE = os.path.join(DATA_DIR, "analysis.json")
PLAN_FILE = os.path.join(BASE_DIR, "unfollow_plan.xlsx")
REPORT_FILE = os.path.join(BASE_DIR, "bili_follow_report.xlsx")

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
SELECTED_FILE = os.path.join(DATA_DIR, "selected_unfollow.json")
QR_FILE = os.path.join(DATA_DIR, "login_qr.png")
CURRENT_ACCOUNT_FILE = os.path.join(DATA_DIR, "current_account.json")
GUIDE_FLAG = os.path.join(DATA_DIR, "guide_shown.flag")


def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def current_uid():
    info = load_json(CURRENT_ACCOUNT_FILE, {})
    uid = info.get("uid")
    if not uid:
        legacy = load_json(os.path.join(DATA_DIR, "cookies.json"), {})
        uid = legacy.get("DedeUserID") or legacy.get("DedeUserID__ckMd5")
    return str(uid) if uid else ""


def days_since(ts):
    try:
        return max(0, (int(time.time()) - int(ts)) // 86400)
    except (TypeError, ValueError):
        return None


def human_days(days):
    if days is None:
        return "无"
    days = int(days)
    if days < 30:
        return f"{days}天"
    years, rem = divmod(days, 365)
    months, rem = divmod(rem, 30)
    parts = []
    if years:
        parts.append(f"{years}年")
    if months:
        parts.append(f"{months}个月")
    if rem or not parts:
        parts.append(f"{rem}天")
    return "".join(parts)


def suggest_reason(rec, threshold):
    if rec.get("status") == "cancelled":
        return "账号已注销", "建议取关"
    if rec.get("special") == 1:
        return "特别关注（白名单）", "保留"
    if rec.get("attribute") == 6:
        return "互相关注（白名单）", "保留"
    days = days_since(rec.get("latest_post_ts"))
    if rec.get("latest_post_ts") is None:
        return "无公开投稿/动态", "建议取关"
    if days is not None and days >= threshold:
        reason = f"超过{threshold}天未更新（{human_days(days)}）"
        if int(rec.get("fans") or 0) < 1000:
            reason += "，且粉丝较少"
        return reason, "建议取关"
    return "", "保留"


class _LogWriter:
    def __init__(self, app):
        self.app = app
        self.buf = ""

    def write(self, text):
        self.buf += text
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            self.app.log_queue.put(line)
            self.app._parse_progress(line)

    def flush(self):
        if self.buf:
            self.app.log_queue.put(self.buf)
            self.app._parse_progress(self.buf)
            self.buf = ""


class App:
    def __init__(self, root):
        self.root = root
        self.proc = None
        self.log_queue = queue.Queue()
        self.qr_win = None
        self.last_shown_mid = None
        self.last_shown_updated = 0
        self.seen_mids = set()
        self.batch_start_time = None
        self.batch_total = 0
        self.batch_progress = 0
        self.worker_running = False
        self.worker_thread = None
        root.title("B站关注清理与分析")
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        init_w = min(860, max(640, screen_w - 80))
        init_h = min(620, max(460, screen_h - 140))
        root.geometry(f"{init_w}x{init_h}")
        root.minsize(640, 460)

        self.status_var = tk.StringVar(value="状态：准备就绪")
        self.login_var = tk.StringVar(value="登录：未登录")
        self.login_btn_text = tk.StringVar(value="扫码登录")
        self.count_var = tk.StringVar(value="关注总数：0")
        self.done_var = tk.StringVar(value="已分析：0")
        self.cancel_var = tk.StringVar(value="已注销：0")
        self.error_var = tk.StringVar(value="错误：0")
        self.backfill_var = tk.StringVar(value="待回填：0")
        self.now_var = tk.StringVar(value="最近分析：—")
        self.video_var = tk.StringVar(value="最新视频：—")
        self.dyn_text_var = tk.StringVar(value="最新动态：—")
        self.eta_var = tk.StringVar(value="剩余完成时间：—")

        top = ttk.Frame(root, padding=(10, 6))
        top.pack(fill="x")
        top_items = [
            (self.status_var, ""),
            (self.login_var, ""),
            (self.count_var, ""),
            (self.done_var, ""),
            (self.cancel_var, ""),
            (self.error_var, ""),
            (self.backfill_var, ""),
        ]
        for i, (var, text) in enumerate(top_items):
            if i < 3:
                ttk.Label(top, textvariable=var).grid(row=0, column=i, sticky="w", padx=(0, 20), pady=1)
            else:
                ttk.Label(top, textvariable=var).grid(row=1, column=i - 3, sticky="w", padx=(0, 20), pady=1)
        top.columnconfigure(3, weight=1)
        tk.Button(
            top,
            textvariable=self.login_btn_text,
            command=self.login_logout,
            bg="#FB7299",
            fg="white",
            activebackground="#E05E80",
            activeforeground="white",
            font=("Microsoft YaHei UI", 10, "bold"),
            padx=18,
            pady=4,
            relief="flat",
            cursor="hand2",
        ).grid(row=0, column=3, sticky="e")

        self.progress = ttk.Progressbar(root, maximum=1, length=820)
        self.progress.pack(fill="x", padx=10, pady=6)

        now_frame = ttk.LabelFrame(root, text="最近分析", padding=6)
        now_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.now_label = ttk.Label(now_frame, textvariable=self.now_var)
        self.now_label.pack(anchor="w", fill="x")
        self.video_label = ttk.Label(now_frame, textvariable=self.video_var)
        self.video_label.pack(anchor="w", fill="x")
        self.dyn_label = ttk.Label(now_frame, textvariable=self.dyn_text_var)
        self.dyn_label.pack(anchor="w", fill="x")
        self.eta_label = ttk.Label(now_frame, textvariable=self.eta_var)
        self.eta_label.pack(anchor="w", fill="x")

        opts = ttk.LabelFrame(root, text="选项", padding=(10, 6))
        opts.pack(fill="x", padx=10)

        ttk.Label(opts, text="未更新天数").grid(row=0, column=0, sticky="w")
        self.days_var = tk.StringVar(value="180")
        ttk.Entry(opts, textvariable=self.days_var, width=6).grid(row=0, column=1, sticky="w", padx=(4, 12))

        ttk.Label(opts, text="热度视频数").grid(row=0, column=2, sticky="w")
        self.top_n_var = tk.StringVar(value="10")
        ttk.Entry(opts, textvariable=self.top_n_var, width=6).grid(row=0, column=3, sticky="w", padx=(4, 12))

        self.dynamic_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="含最新动态", variable=self.dynamic_var).grid(row=0, column=4, sticky="w", padx=(0, 12))

        self.refresh_plan_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="生成报告时刷新取关计划", variable=self.refresh_plan_var).grid(row=0, column=5, sticky="w")

        self.hint_label = ttk.Label(
            opts,
            text="提示：\n"
                 "· 含最新动态：勾选后抓取每个UP主的最新动态时间；不勾选只按最新视频判断，速度更快。\n"
                 "· 生成报告时刷新取关计划：勾选后生成报告会重置可编辑取关计划，覆盖手动勾选；不勾选则保留。\n"
                 "· 修改热度视频数后点“开始分析”会自动补齐；改阈值会重新计算建议取关名单。\n"
                 "· 阈值设 0 天≈清空关注（互关、特别关注保留），取关不可逆。",
            foreground="#B00020",
            justify="left",
            font=("Microsoft YaHei UI", 8),
        )
        self.hint_label.grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 0))
        ttk.Button(opts, text="操作指引", command=self.open_guide).grid(row=1, column=6, sticky="e", pady=(6, 0))

        btns = ttk.Frame(root, padding=10)
        btns.pack(fill="x")
        buttons = [
            ("拉取关注列表", lambda: self.run(["fetch"])),
            ("开始分析（续传）", self.run_analyze),
            ("生成报告", self.run_export),
            ("增量同步", self.run_incremental_sync),
            ("打开报告", lambda: self.open_account_file("bili_follow_report.xlsx")),
            ("打开取关计划", lambda: self.open_account_file("unfollow_plan.xlsx")),
            ("取关管理", self.open_unfollow_manager),
            ("批量取关（按计划）", self.run_unfollow),
            ("取关注销账号", self.run_unfollow_cancelled),
            ("停止当前任务", self.stop),
        ]
        for i, (text, cmd) in enumerate(buttons):
            ttk.Button(btns, text=text, command=cmd).grid(row=i // 5, column=i % 5, padx=4, pady=4, sticky="ew")
        for c in range(5):
            btns.columnconfigure(c, weight=1)

        log = ttk.LabelFrame(root, text="运行日志", padding=6)
        log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log = scrolledtext.ScrolledText(log, wrap="word", height=14, state="disabled")
        self.log.pack(fill="both", expand=True)

        self._update_wrap()
        self.root.bind("<Configure>", self._update_wrap)
        self.refresh_state()
        self.root.after(1200, self.poll)
        if not os.path.exists(GUIDE_FLAG):
            self.root.after(800, self._show_first_guide)

    def open_file(self, path):
        def _open():
            if os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception:
                    messagebox.showinfo("提示", f"请手动打开：\n{path}")
            else:
                messagebox.showinfo("提示", "文件还没生成，请先运行对应步骤。")
        return _open

    def account_file(self, name):
        uid = current_uid()
        if uid:
            d = os.path.join(DATA_DIR, "accounts", uid)
            os.makedirs(d, exist_ok=True)
            return os.path.join(d, name)
        return os.path.join(DATA_DIR, name)

    def open_account_file(self, name):
        path = self.account_file(name)
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception:
                messagebox.showinfo("提示", f"请手动打开：\n{path}")
        else:
            messagebox.showinfo("提示", "文件还没生成，请先运行对应步骤。")

    def login_logout(self):
        if current_uid():
            self.run(["logout"])
        else:
            self.run(["login"])

    def open_guide(self):
        guide = (
            "B站关注清理与分析 · 操作指引\n"
            "\n"
            "一、登录\n"
            "点击顶部“扫码登录”，用手机 B站 App 扫码并确认。登录后按钮会变成“退出登录”。\n"
            "每个账号的数据会独立保存在 data/accounts/你的UID/ 文件夹里。\n"
            "\n"
            "二、首次使用\n"
            "1. 点“拉取关注列表”，把当前关注的全部账号抓下来。\n"
            "2. 点“开始分析（续传）”，逐个分析每个账号的粉丝、分类、热度视频、关注时长和最新动态。\n"
            "3. 分析完成后点“生成报告”，得到 Excel 表格和取关建议。\n"
            "\n"
            "三、日常维护\n"
            "取关或关注了新账号后，点“增量同步”：\n"
            "它会删除已取关的账号，完整分析新增账号，并按你的选择更新现有账号的最新动态时间。\n"
            "\n"
            "四、确认取关对象\n"
            "点“取关管理”打开候选名单，可以：\n"
            "· 点击“选择”列勾选或取消；\n"
            "· 用“全选 / 全不选 / 反选”批量处理；\n"
            "· 输入“距今天数 ≥ N”后点“按天数批量选择”；\n"
            "· 点“保存选择”，再点“开始取关”。\n"
            "取关不可逆，执行前会再次确认；互相关注和特别关注默认不会建议取关。\n"
            "\n"
            "五、查看结果\n"
            "· 点“打开报告”查看完整分析表；\n"
            "· 点“打开取关计划”查看可编辑的候选名单。\n"
            "\n"
            "六、常见问题\n"
            "· 遇到 412/-352 风控：等待 10～30 分钟再继续，工具会自动冷却重试。\n"
            "· 想清空关注：把“未更新天数”设为 0 天，但互关和特别关注仍会保留，且取关不可逆。\n"
            "· 报告/计划打开失败：先点“生成报告”生成对应文件。\n"
        )
        win = tk.Toplevel(self.root)
        win.title("操作指引")
        win.geometry("760x560")
        txt = scrolledtext.ScrolledText(win, wrap="word", font=("Microsoft YaHei UI", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", guide)
        txt.configure(state="disabled")

    def _show_first_guide(self):
        self.open_guide()
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(GUIDE_FLAG, "w", encoding="utf-8") as f:
            f.write("shown")

    def open_unfollow_manager(self):
        analysis = load_json(self.account_file("analysis.json"), {})
        try:
            threshold = int(self.days_var.get() or 180)
        except ValueError:
            threshold = 180

        candidates = []
        for mid, rec in analysis.items():
            reason, action = suggest_reason(rec, threshold)
            if action != "建议取关":
                continue
            days = days_since(rec.get("latest_post_ts"))
            top = rec.get("top_videos") or []
            tops = []
            for v in top[:3]:
                title = (v.get("title") or "（无标题）")[:42]
                tops.append(f"{title}（播放{v.get('play') or 0}）")
            while len(tops) < 3:
                tops.append("")
            name = "".join(ch for ch in str(rec.get("uname") or "") if unicodedata.category(ch)[0] != "C").strip()
            name = name or "（无显示昵称）"
            candidates.append({
                "mid": mid,
                "name": name,
                "fans": rec.get("fans") or 0,
                "days": days,
                "time": human_days(days),
                "reason": reason,
                "category": rec.get("category") or "",
                "tops": tops,
                "selected": False,
            })
        candidates.sort(key=lambda c: c["days"] if c["days"] is not None else 10**9, reverse=True)

        win = tk.Toplevel(self.root)
        win.title("取关管理")
        win.geometry("1200x600")

        toolbar = ttk.Frame(win, padding=6)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="全选", command=lambda: self._set_all(candidates, tree, True, update_count)).pack(side="left", padx=3)
        ttk.Button(toolbar, text="全不选", command=lambda: self._set_all(candidates, tree, False, update_count)).pack(side="left", padx=3)
        ttk.Button(toolbar, text="反选", command=lambda: self._invert(candidates, tree, update_count)).pack(side="left", padx=3)
        ttk.Label(toolbar, text="距今天数 ≥").pack(side="left", padx=(12, 3))
        day_entry = ttk.Entry(toolbar, width=6)
        day_entry.insert(0, "365")
        day_entry.pack(side="left")
        ttk.Label(toolbar, text="天").pack(side="left")
        ttk.Button(toolbar, text="按天数批量选择", command=lambda: self._select_by_days(candidates, tree, day_entry.get(), update_count)).pack(side="left", padx=3)
        ttk.Button(toolbar, text="保存选择", command=lambda: self._save_selection(candidates)).pack(side="left", padx=12)
        ttk.Button(toolbar, text="开始取关", command=lambda: self._start_selected(candidates, win)).pack(side="left", padx=3)
        count_var = tk.StringVar(value="已选择：0")
        ttk.Label(toolbar, textvariable=count_var).pack(side="left", padx=12)

        def update_count():
            count_var.set(f"已选择：{sum(1 for c in candidates if c['selected'])}")

        tree_area = ttk.Frame(win, padding=6)
        tree_area.pack(fill="both", expand=True)

        cols = ["选择", "UID", "名字", "粉丝数", "距今天数(天)", "未更新时长", "建议理由", "主要分类", "热门视频Top1", "热门视频Top2", "热门视频Top3"]
        tree = ttk.Treeview(tree_area, columns=cols, show="headings", height=20)
        widths = {"选择": 45, "UID": 90, "名字": 140, "粉丝数": 80, "距今天数(天)": 90, "未更新时长": 100, "建议理由": 180, "主要分类": 70, "热门视频Top1": 220, "热门视频Top2": 220, "热门视频Top3": 220}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 120), anchor="w")
        tree.column("选择", anchor="center")

        vsb = ttk.Scrollbar(tree_area, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_area, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree_area.grid_rowconfigure(0, weight=1)
        tree_area.grid_columnconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for c in candidates:
            iid = tree.insert("", "end", values=["☐", c["mid"], c["name"], c["fans"],
                                                 c["days"] if c["days"] is not None else "无",
                                                 c["time"], c["reason"], c["category"],
                                                 c["tops"][0], c["tops"][1], c["tops"][2]])
            c["iid"] = iid

        def toggle(event):
            region = tree.identify("region", event.x, event.y)
            column = tree.identify_column(event.x)
            if region != "cell" or column != "#1":
                return
            iid = tree.identify_row(event.y)
            if not iid:
                return
            c = next((x for x in candidates if x["iid"] == iid), None)
            if not c:
                return
            c["selected"] = not c["selected"]
            tree.set(iid, "选择", "☑" if c["selected"] else "☐")
            update_count()

        tree.bind("<Button-1>", toggle, add="+")

    def _set_all(self, candidates, tree, value, on_change=None):
        for c in candidates:
            c["selected"] = value
            tree.set(c["iid"], "选择", "☑" if value else "☐")
        if on_change:
            on_change()

    def _invert(self, candidates, tree, on_change=None):
        for c in candidates:
            c["selected"] = not c["selected"]
            tree.set(c["iid"], "选择", "☑" if c["selected"] else "☐")
        if on_change:
            on_change()

    def _select_by_days(self, candidates, tree, text, on_change=None):
        digits = "".join(ch for ch in str(text) if ch.isdigit())
        if not digits:
            messagebox.showinfo("提示", "请输入正确的天数。")
            return
        n = int(digits)
        for c in candidates:
            c["selected"] = False
            tree.set(c["iid"], "选择", "☐")
        for c in candidates:
            if c["days"] is None or c["days"] >= n:
                c["selected"] = True
                tree.set(c["iid"], "选择", "☑")
        if on_change:
            on_change()

    def _save_selection(self, candidates):
        selected = [c["mid"] for c in candidates if c["selected"]]
        path = self.account_file("selected_unfollow.json")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(selected, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("已保存", f"已选择 {len(selected)} 个取关对象。")

    def _start_selected(self, candidates, win):
        self._save_selection(candidates)
        win.destroy()
        self.run(["unfollow", "--from-selected"], confirm=True)

    def run(self, args, confirm=False):
        if self.worker_running:
            messagebox.showwarning("提示", "已有任务在运行，请先等待或停止。")
            return
        if confirm and not messagebox.askyesno("确认", "取关不可逆，确定继续吗？"):
            return
        if args and args[0] == "login":
            try:
                os.remove(QR_FILE)
            except OSError:
                pass
        self.batch_start_time = time.time()
        self.batch_total = 0
        self.batch_progress = 0
        self.eta_var.set("剩余完成时间：—")
        self.worker_running = True
        self.status_var.set("状态：运行中")
        self.worker_thread = threading.Thread(target=self._run_worker, args=(args,), daemon=True)
        self.worker_thread.start()
        if args and args[0] == "login":
            threading.Thread(target=self._wait_for_qr, daemon=True).start()
            threading.Thread(target=self._wait_for_login, daemon=True).start()

    def _run_worker(self, args):
        writer = _LogWriter(self)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = writer
        sys.stderr = writer
        try:
            ns = bm.build_parser().parse_args(args)
            client = bm.BiliClient()
            cmd = ns.cmd
            if cmd == "login":
                bm.cmd_login(client, ns)
            elif cmd == "logout":
                bm.cmd_logout(client, ns)
            elif cmd == "fetch":
                bm.cmd_fetch(client, ns)
            elif cmd == "analyze":
                bm.cmd_analyze(client, ns)
            elif cmd == "export":
                bm.cmd_export(client, ns)
            elif cmd == "unfollow":
                bm.cmd_unfollow(client, ns)
            elif cmd == "unfollow-cancelled":
                bm.cmd_unfollow_cancelled(client, ns)
            elif cmd == "status":
                bm.cmd_status(client, ns)
            elif cmd == "incremental-sync":
                bm.cmd_incremental_sync(client, ns)
            elif cmd == "run":
                bm.cmd_run(client, ns)
        except SystemExit:
            pass
        except Exception as e:
            self.log_queue.put(f"[错误] {e}")
        finally:
            sys.stdout = old_out
            sys.stderr = old_err
            writer.flush()
            self.log_queue.put("[任务结束]")
            self.worker_running = False

    def _wait_for_qr(self):
        for _ in range(120):
            if os.path.exists(QR_FILE):
                self.root.after(0, self._show_qr)
                return
            time.sleep(0.5)

    def _show_qr(self):
        if not os.path.exists(QR_FILE):
            messagebox.showinfo("提示", "二维码生成失败，请查看运行日志里的链接。")
            return
        win = tk.Toplevel(self.root)
        self.qr_win = win
        win.title("扫码登录")
        try:
            img = tk.PhotoImage(file=QR_FILE)
        except Exception:
            messagebox.showinfo("提示", "无法显示二维码图片，请查看运行日志里的链接。")
            win.destroy()
            return
        label = ttk.Label(win, image=img)
        label.image = img
        label.pack(padx=12, pady=12)
        ttk.Label(win, text="请用手机 B站 App 扫码登录，并在手机上确认").pack(pady=(0, 12))

    def _wait_for_login(self):
        for _ in range(720):
            if current_uid():
                self.root.after(0, self._on_login_success)
                return
            time.sleep(0.5)

    def _on_login_success(self):
        if self.qr_win is not None and self.qr_win.winfo_exists():
            self.qr_win.destroy()
        self.qr_win = None
        messagebox.showinfo("登录成功", "已登录成功。")

    def run_analyze(self):
        args = ["analyze", "--only-new"]
        if not self.dynamic_var.get():
            args.append("--no-dynamic")
        args += ["--top-n", self.top_n_var.get() or "10"]
        self.run(args)

    def run_export(self):
        args = ["export", "--days", self.days_var.get() or "180"]
        if self.refresh_plan_var.get():
            args.append("--refresh-plan")
        self.run(args)

    def run_incremental_sync(self):
        update_dynamic = messagebox.askyesno(
            "增量同步",
            "是否同时更新每个账号的最新动态时间？\n\n"
            "选择“是”：新增账号完整分析，现有账号也会逐个刷新最新动态，耗时较长。\n"
            "选择“否”：只处理新增和已取关账号，速度更快。",
        )
        args = ["incremental-sync", "--days", self.days_var.get() or "180", "--export"]
        if not update_dynamic:
            args.append("--no-dynamic-update")
        self.run(args)

    def run_unfollow(self):
        self.run(["unfollow"], confirm=True)

    def run_unfollow_cancelled(self):
        self.run(["unfollow-cancelled", "--yes"], confirm=True)

    def stop(self):
        if self.worker_running:
            self._force_stop_thread(self.worker_thread)
            self.worker_running = False
            self.log_queue.put("[已请求强制停止]")
        else:
            messagebox.showinfo("提示", "当前没有运行中的任务。")

    def _force_stop_thread(self, thread):
        if thread is None or thread.ident is None:
            return
        tid = ctypes.c_long(thread.ident)
        exc = ctypes.py_object(SystemExit)
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, exc)
        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)

    def _reader(self):
        try:
            for line in self.proc.stdout:
                text = line.rstrip("\n")
                self.log_queue.put(text)
                self._parse_progress(text)
        except Exception as e:
            self.log_queue.put(f"[读取日志出错] {e}")
        self.log_queue.put("[任务结束]")
        self.proc = None

    def _parse_progress(self, text):
        patterns = [
            (r"已拉取\s+(\d+)/(\d+)", None),
            (r"更新动态进度：(\d+)/(\d+)", None),
            (r"已写入\s+(\d+)/(\d+)", None),
            (r"已处理\s+(\d+)/(\d+)", None),
            (r"\[(\d+)/(\d+)\]\s*分析", None),
            (r"\[(\d+)/(\d+)\]\s*取关", None),
        ]
        for pat, _ in patterns:
            m = re.search(pat, text)
            if m:
                self.batch_progress = int(m.group(1))
                self.batch_total = int(m.group(2))
                return

    def append_log(self, line):
        self.log.configure(state="normal")
        self.log.insert("end", line + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _update_wrap(self, event=None):
        width = max(300, self.root.winfo_width() - 40)
        for label in (self.hint_label, self.now_label, self.video_label, self.dyn_label):
            label.configure(wraplength=width)

    def refresh_state(self):
        uid = current_uid()
        self.login_var.set(f"登录：UID {uid}" if uid else "登录：未登录")
        self.login_btn_text.set("退出登录" if uid else "扫码登录")
        follows = load_json(self.account_file("followings.json"), [])
        analysis = load_json(self.account_file("analysis.json"), {})
        total = len(follows)
        done = sum(1 for r in analysis.values() if r.get("done"))
        cancelled = sum(1 for r in analysis.values() if r.get("status") == "cancelled")
        errors = sum(1 for r in analysis.values() if r.get("status") == "error")
        self.count_var.set(f"关注总数：{total}")
        self.done_var.set(f"已分析：{done}")
        self.cancel_var.set(f"已注销：{cancelled}")
        self.error_var.set(f"错误：{errors}")
        self.progress.configure(maximum=max(1, total), value=done)

        try:
            top_n = max(1, int(self.top_n_var.get() or 10))
        except ValueError:
            top_n = 10
        need_backfill = 0
        for r in analysis.values():
            if not (r.get("done") and r.get("status") == "ok"):
                continue
            have = len(r.get("top_videos") or [])
            total = r.get("video_total")
            if total is None:
                need_backfill += have < top_n
            else:
                need_backfill += have < min(top_n, int(total))
        self.backfill_var.set(f"待回填：{need_backfill}")

        if analysis:
            latest = max(analysis.values(), key=lambda r: r.get("updated") or 0)
            updated = latest.get("updated") or 0
            name = latest.get("uname") or latest.get("mid") or "—"
            kind = "回填" if latest.get("mid") in self.seen_mids else "分析"
            self.now_var.set(f"最近{kind}：{name}｜分类 {latest.get('category') or '-'}｜粉丝 {latest.get('fans') or 0}")
            self.video_var.set(f"最新视频：{latest.get('latest_video_title') or '—'}")
            dyn = latest.get("latest_dynamic_text") or latest.get("latest_dynamic_type") or "—"
            self.dyn_text_var.set(f"最新动态：{dyn}")
            if updated != self.last_shown_updated:
                self.last_shown_updated = updated
                title = latest.get("latest_video_title") or "—"
                self.append_log(
                    f"[最近{kind}] {name}｜分类 {latest.get('category') or '-'}｜粉丝 {latest.get('fans') or 0}｜"
                    f"最新视频：{title}｜最新动态：{dyn}"
                )
            self.seen_mids.add(latest.get("mid"))

    def poll(self):
        while True:
            try:
                line = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.append_log(line)
        self.refresh_state()
        if self.worker_running:
            self.status_var.set("状态：运行中")
            if self.batch_total > 0:
                self.progress.configure(maximum=self.batch_total, value=self.batch_progress)
                if self.batch_progress > 0:
                    elapsed = time.time() - (self.batch_start_time or time.time())
                    eta = elapsed * (self.batch_total - self.batch_progress) / self.batch_progress
                    if eta < 60:
                        eta_text = f"{int(eta)}秒"
                    elif eta < 3600:
                        eta_text = f"{int(eta // 60)}分{int(eta % 60)}秒"
                    else:
                        eta_text = f"{int(eta // 3600)}小时{int((eta % 3600) // 60)}分"
                    self.eta_var.set(f"剩余完成时间：{eta_text}")
        else:
            self.status_var.set("状态：准备就绪")
        self.root.after(1200, self.poll)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
