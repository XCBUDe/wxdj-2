#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
网销点检 — 任务控制面板（图形界面）

双击「启动控制面板.bat」即可打开。功能：
  ① 依赖安装情况检测 + 一键安装（随包附带 vendor/ 离线依赖，pip 失败也能跑）
  ② 输入文件选择（经销商名单 / 报价标准），点击选择 + 拖拽
  ③ 平台选择：汽车之家 / 易车 / 懂车帝（预留接口）
  ④ 任务勾选：报价 / 软文 / 截图 / 核价 / 打包，只勾需要的，避免全量白跑
  ⑤ 启动前依赖预检：缺什么提示自动安装，装完自动继续执行

实现方式：面板把所选配置写成 JSON，经环境变量 WXDJ_CONFIG_JSON
注入 task.py / yiche_task.py，不修改原命令行用法。
"""

import os
import sys
import json
import shutil
import threading
import subprocess
import importlib.util
from pathlib import Path

HERE = Path(__file__).resolve().parent

# vendor 离线依赖（bs4 / tkinterdnd2 等随包附带，pip 装不上也能用）
_VENDOR = HERE / "vendor"
if _VENDOR.is_dir() and str(_VENDOR) not in sys.path:
    sys.path.append(str(_VENDOR))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =============================================================================
# ★ 平台注册表：以后接入懂车帝，把 enabled 改 True 并实现 dongchedi_task.py ★
# =============================================================================
PLATFORMS = {
    "autohome": {
        "label": "汽车之家",
        "script": "task.py",
        "output_dir": "output",
        "default_concurrency": 2,
        "max_concurrency": 4,
        "requires": ["requests", "openpyxl", "playwright", "parsel"],
        "needs_browser": True,
        "enabled": True,
    },
    "yiche": {
        "label": "易车",
        "script": "yiche_task.py",
        "output_dir": "output_yiche",
        "default_concurrency": 3,
        "max_concurrency": 5,
        "requires": ["requests", "openpyxl", "bs4"],
        "needs_browser": False,
        "enabled": True,
    },
    "dongchedi": {
        "label": "懂车帝（预留）",
        "script": "dongchedi_task.py",
        "output_dir": "output_dongchedi",
        "default_concurrency": 2,
        "max_concurrency": 3,
        "requires": ["requests", "openpyxl"],
        "needs_browser": False,
        "enabled": False,
    },
}

DEPENDENCIES = [
    ("requests", "requests"), ("openpyxl", "openpyxl"),
    ("beautifulsoup4", "bs4"), ("pillow", "PIL"),
    ("parsel", "parsel"), ("playwright", "playwright"),
    ("拖拽支持", "tkinterdnd2"),
]
IMP2DISP = {imp: disp for disp, imp in DEPENDENCIES}

PIP_MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
PW_MIRROR = "https://npmmirror.com/mirrors/playwright/"

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    _HAS_DND = False

# ── SmileCare 风格配色 ───────────────────────────────────────────────────────
C = {
    "bg":        "#EDF2F7",   # 页面底色：浅灰蓝
    "card":      "#FFFFFF",   # 卡片
    "border":    "#E2E8F0",
    "primary":   "#0E9F8A",   # 主色：青绿
    "primary_d": "#0B7F6E",   # 主色加深（hover）
    "primary_l": "#E4F5F2",   # 主色浅底
    "ink":       "#16283A",   # 主文字：深藏蓝
    "muted":     "#64748B",   # 次要文字
    "ok":        "#16A34A",
    "bad":       "#DC2626",
    "danger_l":  "#FDECEC",
    "log_bg":    "#102433",   # 控制台深底
    "log_fg":    "#BFE8DE",
}
FONT = ("Microsoft YaHei UI", 10)
FONT_S = ("Microsoft YaHei UI", 9)
FONT_T = ("Microsoft YaHei UI", 15, "bold")
FONT_H = ("Microsoft YaHei UI", 10, "bold")


def _module_ok(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def _find_browser() -> str:
    cands = [
        ("Edge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        ("Edge", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        ("Chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        ("Chrome", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    ]
    for name, p in cands:
        if Path(p).exists():
            return name
    for name, exe in [("Edge", "msedge"), ("Chrome", "chrome"), ("Chromium", "chromium")]:
        if shutil.which(exe):
            return name
    for d in (Path.home() / "AppData/Local/ms-playwright",
              Path.home() / ".cache/ms-playwright"):
        if d.exists() and any(d.glob("chromium*")):
            return "Playwright-Chromium"
    return ""


# ── 圆角风按钮（tk.Button 扁平化模拟） ──────────────────────────────────────
class Btn(tk.Button):
    def __init__(self, master, text, command=None, kind="ghost", **kw):
        style = {
            "primary": dict(bg=C["primary"], fg="white",
                            activebackground=C["primary_d"], activeforeground="white"),
            "ghost":   dict(bg=C["card"], fg=C["ink"],
                            activebackground=C["primary_l"], activeforeground=C["ink"],
                            highlightbackground=C["border"], highlightthickness=1),
            "danger":  dict(bg=C["danger_l"], fg=C["bad"],
                            activebackground="#F8D7D7", activeforeground=C["bad"]),
        }[kind]
        super().__init__(master, text=text, command=command, font=FONT_S,
                         relief="flat", bd=0, cursor="hand2",
                         padx=14, pady=6, **style, **kw)


class ControlPanel:

    def __init__(self):
        self.root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
        self.root.title("网销点检 · 任务控制面板")
        self.root.geometry("900x820")
        self.root.minsize(780, 660)
        self.root.configure(bg=C["bg"])

        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure(".", font=FONT, background=C["card"], foreground=C["ink"])
        st.configure("Card.TFrame", background=C["card"])
        st.configure("Bg.TFrame", background=C["bg"])
        st.configure("TCheckbutton", background=C["card"], foreground=C["ink"])
        st.map("TCheckbutton", background=[("active", C["card"])])
        st.configure("TRadiobutton", background=C["card"], foreground=C["ink"])
        st.map("TRadiobutton", background=[("active", C["card"])])
        st.configure("TEntry", fieldbackground="#F7FAFC", bordercolor=C["border"])
        st.configure("TSpinbox", fieldbackground="#F7FAFC")

        self.proc = None
        self.dep_labels = {}
        self._pending_start = False     # 自动装完依赖后续跑

        self._build_ui()
        self.refresh_deps()
        self.autodetect_files()

    # ──────────────────────────────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────────────────────────────
    def _card(self, parent, num, title, hint=""):
        wrap = tk.Frame(parent, bg=C["bg"])
        wrap.pack(fill="x", padx=16, pady=(0, 10))
        card = tk.Frame(wrap, bg=C["card"], highlightbackground=C["border"],
                        highlightthickness=1)
        card.pack(fill="x")
        head = tk.Frame(card, bg=C["card"])
        head.pack(fill="x", padx=14, pady=(10, 2))
        badge = tk.Label(head, text=str(num), bg=C["primary"], fg="white",
                         font=("Microsoft YaHei UI", 9, "bold"), width=2)
        badge.pack(side="left")
        tk.Label(head, text=" " + title, bg=C["card"], fg=C["ink"],
                 font=FONT_H).pack(side="left")
        if hint:
            tk.Label(head, text=hint, bg=C["card"], fg=C["muted"],
                     font=FONT_S).pack(side="left", padx=8)
        body = tk.Frame(card, bg=C["card"])
        body.pack(fill="x", padx=14, pady=(4, 12))
        return body

    def _build_ui(self):
        # ===== 顶栏 =====
        top = tk.Frame(self.root, bg=C["card"], highlightbackground=C["border"],
                       highlightthickness=1)
        top.pack(fill="x")
        tk.Label(top, text="●", fg=C["primary"], bg=C["card"],
                 font=("Arial", 16)).pack(side="left", padx=(18, 6), pady=12)
        tk.Label(top, text="网销点检 · 任务控制面板", bg=C["card"], fg=C["ink"],
                 font=FONT_T).pack(side="left")
        tk.Label(top, text="经销商网销价格执行核查工作台", bg=C["card"],
                 fg=C["muted"], font=FONT_S).pack(side="left", padx=12)
        self.lbl_state = tk.Label(top, text="● 就绪", bg=C["primary_l"],
                                  fg=C["primary_d"], font=FONT_S, padx=10, pady=3)
        self.lbl_state.pack(side="right", padx=18)

        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True, pady=(12, 0))

        # ===== ① 依赖 =====
        b1 = self._card(body, 1, "依赖安装情况",
                        "随包自带 vendor 离线依赖，缺项会在启动前自动安装")
        grid = tk.Frame(b1, bg=C["card"]); grid.pack(fill="x")
        for i, (disp, imp) in enumerate(DEPENDENCIES):
            r, c = divmod(i, 4)
            lbl = tk.Label(grid, text="● " + disp, bg=C["card"], font=FONT_S,
                           anchor="w", width=18)
            lbl.grid(row=r, column=c, sticky="w", padx=2, pady=2)
            self.dep_labels[imp] = (lbl, disp)
        self.lbl_browser = tk.Label(grid, text="● 浏览器", bg=C["card"],
                                    font=FONT_S, anchor="w", width=18)
        self.lbl_browser.grid(row=len(DEPENDENCIES) // 4,
                              column=len(DEPENDENCIES) % 4, sticky="w", padx=2)
        btns = tk.Frame(b1, bg=C["card"]); btns.pack(fill="x", pady=(8, 0))
        Btn(btns, "重新检测", self.refresh_deps).pack(side="left", padx=(0, 6))
        Btn(btns, "一键安装依赖", self.install_deps).pack(side="left", padx=6)
        Btn(btns, "安装 Chromium（仅汽车之家）",
            self.install_browser).pack(side="left", padx=6)

        # ===== ② 输入文件 =====
        dnd = "支持拖拽 xlsx 到输入框" if _HAS_DND else "点「选择文件」浏览"
        b2 = self._card(body, 2, "输入文件", dnd)
        self.var_list, self.var_std = tk.StringVar(), tk.StringVar()
        self._file_row(b2, "经销商名单 *", self.var_list)
        self._file_row(b2, "报价标准（不核价可留空）", self.var_std)
        Btn(b2, "自动识别本文件夹 xlsx",
            self.autodetect_files).pack(anchor="w", pady=(6, 0))

        # ===== ③ 平台 =====
        b3 = self._card(body, 3, "平台", "单次任务只跑单个平台")
        self.var_platform = tk.StringVar(value="yiche")
        row = tk.Frame(b3, bg=C["card"]); row.pack(fill="x")
        for key, info in PLATFORMS.items():
            rb = ttk.Radiobutton(row, text=info["label"], value=key,
                                 variable=self.var_platform,
                                 command=self.on_platform_change)
            rb.pack(side="left", padx=(0, 24))
            if not info["enabled"]:
                rb.state(["disabled"])

        # ===== ④ 任务 =====
        b4 = self._card(body, 4, "任务选择", "只勾需要的，避免全量白跑")
        self.var_pricing = tk.BooleanVar(value=True)
        self.var_article = tk.BooleanVar(value=True)
        self.var_shot = tk.BooleanVar(value=True)
        self.var_compare = tk.BooleanVar(value=True)
        self.var_zip = tk.BooleanVar(value=True)
        self.var_clear_ckpt = tk.BooleanVar(value=False)

        r1 = tk.Frame(b4, bg=C["card"]); r1.pack(fill="x", pady=2)
        ttk.Checkbutton(r1, text="车型报价采集", variable=self.var_pricing,
                        command=self._sync_task_state).pack(side="left", padx=(0, 18))
        ttk.Checkbutton(r1, text="最新软文日期",
                        variable=self.var_article).pack(side="left", padx=18)
        ttk.Checkbutton(r1, text="头图截图", variable=self.var_shot,
                        command=self._sync_task_state).pack(side="left", padx=18)
        r2 = tk.Frame(b4, bg=C["card"]); r2.pack(fill="x", pady=2)
        self.cb_compare = ttk.Checkbutton(
            r2, text="核价对比着色（需报价采集 + 标准表）", variable=self.var_compare)
        self.cb_compare.pack(side="left", padx=(0, 18))
        self.cb_zip = ttk.Checkbutton(
            r2, text="截图打包 ZIP（需头图截图）", variable=self.var_zip)
        self.cb_zip.pack(side="left", padx=18)
        r3 = tk.Frame(b4, bg=C["card"]); r3.pack(fill="x", pady=(6, 0))
        Btn(r3, "全选", lambda: self._set_all(True)).pack(side="left", padx=(0, 6))
        Btn(r3, "全不选", lambda: self._set_all(False)).pack(side="left", padx=4)
        tk.Label(r3, text="并发数", bg=C["card"], fg=C["muted"],
                 font=FONT_S).pack(side="left", padx=(20, 6))
        self.var_conc = tk.IntVar(value=PLATFORMS["yiche"]["default_concurrency"])
        self.spin_conc = ttk.Spinbox(r3, from_=1, to=5, width=4,
                                     textvariable=self.var_conc)
        self.spin_conc.pack(side="left")
        ttk.Checkbutton(r3, text="清除断点重新采集（换任务组合时建议勾选）",
                        variable=self.var_clear_ckpt).pack(side="left", padx=24)

        # ===== 运行区 =====
        run = tk.Frame(body, bg=C["bg"]); run.pack(fill="x", padx=16, pady=(2, 8))
        self.btn_run = Btn(run, "▶  开始执行", self.start_task, kind="primary")
        self.btn_run.configure(font=FONT_H, padx=22, pady=8)
        self.btn_run.pack(side="left")
        self.btn_stop = Btn(run, "■ 停止", self.stop_task, kind="danger")
        self.btn_stop.pack(side="left", padx=8)
        self.btn_stop.configure(state="disabled")
        Btn(run, "打开输出文件夹", self.open_output).pack(side="left", padx=8)

        # ===== 日志（控制台风格） =====
        logwrap = tk.Frame(body, bg=C["bg"]); logwrap.pack(
            fill="both", expand=True, padx=16, pady=(0, 14))
        logcard = tk.Frame(logwrap, bg=C["log_bg"], highlightbackground=C["border"],
                           highlightthickness=1)
        logcard.pack(fill="both", expand=True)
        tk.Label(logcard, text="  运行日志", bg=C["log_bg"], fg="#6E8FA3",
                 font=FONT_S, anchor="w").pack(fill="x", pady=(6, 0))
        inner = tk.Frame(logcard, bg=C["log_bg"]); inner.pack(
            fill="both", expand=True, padx=8, pady=6)
        self.log = tk.Text(inner, height=12, state="disabled", bd=0,
                           bg=C["log_bg"], fg=C["log_fg"],
                           insertbackground=C["log_fg"],
                           font=("Consolas", 9), wrap="word")
        sb = ttk.Scrollbar(inner, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        self.on_platform_change()

    def _file_row(self, parent, label, var):
        row = tk.Frame(parent, bg=C["card"]); row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=C["card"], fg=C["ink"], font=FONT_S,
                 width=22, anchor="w").pack(side="left")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True, padx=6)
        Btn(row, "选择文件…", lambda: self._pick_file(var)).pack(side="left")
        if _HAS_DND:
            ent.drop_target_register(DND_FILES)
            ent.dnd_bind("<<Drop>>",
                         lambda e, v=var: v.set(e.data.strip("{}").strip()))

    def _pick_file(self, var):
        p = filedialog.askopenfilename(
            title="选择 Excel 文件", initialdir=str(HERE),
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if p:
            var.set(p)

    # ──────────────────────────────────────────────────────────────────
    # 状态
    # ──────────────────────────────────────────────────────────────────
    def _set_state(self, text, running=False):
        self.lbl_state.config(
            text="● " + text,
            bg="#FFF4E0" if running else C["primary_l"],
            fg="#B45309" if running else C["primary_d"])

    def logln(self, text=""):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def refresh_deps(self):
        for imp, (lbl, disp) in self.dep_labels.items():
            ok = _module_ok(imp)
            lbl.config(text=("● " if ok else "✕ ") + disp,
                       fg=C["ok"] if ok else C["bad"])
        b = _find_browser()
        self.lbl_browser.config(
            text=("● 浏览器(" + b + ")") if b else "✕ 浏览器(未检测到)",
            fg=C["ok"] if b else C["bad"])

    def autodetect_files(self):
        cands = [p for p in HERE.glob("*.xlsx")
                 if not p.name.startswith("~$") and "result" not in p.name.lower()]
        if not self.var_list.get():
            for p in cands:
                if any(k in p.name for k in ("检核", "名单", "uid", "UID", "经销商")):
                    self.var_list.set(str(p)); break
        if not self.var_std.get():
            for p in cands:
                if any(k in p.name for k in ("报价", "标准", "建议")):
                    self.var_std.set(str(p)); break

    def on_platform_change(self):
        info = PLATFORMS[self.var_platform.get()]
        self.var_conc.set(info["default_concurrency"])
        self.spin_conc.config(to=info["max_concurrency"])
        self._sync_task_state()

    def _sync_task_state(self):
        if not self.var_pricing.get():
            self.var_compare.set(False); self.cb_compare.state(["disabled"])
        else:
            self.cb_compare.state(["!disabled"])
        if not self.var_shot.get():
            self.var_zip.set(False); self.cb_zip.state(["disabled"])
        else:
            self.cb_zip.state(["!disabled"])

    def _set_all(self, value):
        for v in (self.var_pricing, self.var_article, self.var_shot,
                  self.var_compare, self.var_zip):
            v.set(value)
        self._sync_task_state()

    def open_output(self):
        out = HERE / PLATFORMS[self.var_platform.get()]["output_dir"]
        out.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(out)            # noqa
        else:
            subprocess.Popen(["xdg-open", str(out)])

    # ──────────────────────────────────────────────────────────────────
    # 子进程
    # ──────────────────────────────────────────────────────────────────
    def _run_subprocess(self, cmd, env=None, on_done=None):
        def _reader():
            rc = -1
            try:
                e = dict(os.environ, PYTHONIOENCODING="utf-8",
                         PYTHONUTF8="1", **(env or {}))
                self.proc = subprocess.Popen(
                    cmd, cwd=str(HERE), env=e,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1)
                for line in self.proc.stdout:
                    self.root.after(0, self.logln, line.rstrip())
                self.proc.wait()
                rc = self.proc.returncode
                self.root.after(0, self.logln, f"—— 进程结束（返回码 {rc}） ——")
            except Exception as ex:
                self.root.after(0, self.logln, f"[面板错误] {ex}")
            finally:
                self.proc = None
                if on_done:
                    self.root.after(0, on_done, rc)
        threading.Thread(target=_reader, daemon=True).start()

    def install_deps(self, on_done=None):
        if self.proc:
            messagebox.showwarning("提示", "已有任务在运行，请先停止。"); return
        self._set_state("安装依赖中…", running=True)
        self.logln("[依赖] 开始安装（清华镜像，失败自动回退官方源）…")
        py = sys.executable
        cmd = (f'"{py}" -m pip install -r requirements.txt -i {PIP_MIRROR}'
               f' || "{py}" -m pip install -r requirements.txt')
        shell = ["cmd", "/c", cmd] if sys.platform == "win32" \
            else ["bash", "-c", cmd.replace('"', "'")]

        def _done(rc):
            self.refresh_deps()
            self._set_state("就绪")
            if on_done:
                on_done(rc)
        self._run_subprocess(shell, on_done=_done)

    def install_browser(self):
        if self.proc:
            messagebox.showwarning("提示", "已有任务在运行，请先停止。"); return
        self._set_state("安装浏览器中…", running=True)
        self.logln("[浏览器] 开始安装 Playwright Chromium（国内镜像）…")
        self._run_subprocess(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            env={"PLAYWRIGHT_DOWNLOAD_HOST": PW_MIRROR},
            on_done=lambda rc: (self.refresh_deps(), self._set_state("就绪")))

    # ──────────────────────────────────────────────────────────────────
    # 启动任务（含依赖预检）
    # ──────────────────────────────────────────────────────────────────
    def _missing_modules(self, info, features):
        need = list(info["requires"])
        if features.get("screenshot"):
            need.append("PIL")           # 截图嵌 Excel 需要 pillow
        return [m for m in dict.fromkeys(need) if not _module_ok(m)]

    def start_task(self):
        if self.proc:
            messagebox.showwarning("提示", "已有任务在运行。"); return

        pf_key = self.var_platform.get()
        info = PLATFORMS[pf_key]
        if not info["enabled"]:
            messagebox.showinfo("预留接口", f"{info['label']} 尚未实现。"); return

        dealer = self.var_list.get().strip()
        std = self.var_std.get().strip()
        if not dealer or not Path(dealer).exists():
            messagebox.showerror("缺少输入", "请先选择经销商名单 xlsx。"); return

        features = {
            "pricing": self.var_pricing.get(),
            "article_date": self.var_article.get(),
            "screenshot": self.var_shot.get(),
        }
        if not any(features.values()):
            messagebox.showerror("未选任务", "至少勾选一个采集任务。"); return

        # ── 依赖预检：缺什么 → 自动安装 → 装完自动续跑 ──
        missing = self._missing_modules(info, features)
        if missing:
            names = "、".join(IMP2DISP.get(m, m) for m in missing)
            if not messagebox.askyesno(
                    "缺少依赖",
                    f"运行「{info['label']}」还缺少：{names}\n\n"
                    "是否现在自动安装？安装完成后将自动开始任务。"):
                return
            self.logln(f"[预检] 缺少依赖：{names} → 自动安装后续跑")

            def _after(rc):
                still = self._missing_modules(info, features)
                if still:
                    self.logln("[预检] 仍缺少：" +
                               "、".join(IMP2DISP.get(m, m) for m in still))
                    messagebox.showerror(
                        "依赖安装未完成",
                        "自动安装后仍缺少依赖，请检查网络后点「一键安装依赖」重试。")
                else:
                    self.logln("[预检] 依赖已就绪，自动开始任务。")
                    self.start_task()
            self.install_deps(on_done=_after)
            return

        if info["needs_browser"] and not _find_browser():
            if not messagebox.askyesno(
                    "未检测到浏览器",
                    "汽车之家平台需要 Edge/Chrome/Chromium。\n"
                    "未检测到浏览器，仍然尝试运行吗？\n"
                    "（建议先点「安装 Chromium」）"):
                return

        compare = self.var_compare.get()
        if compare and not (std and Path(std).exists()):
            if not messagebox.askyesno(
                    "缺少报价标准",
                    "勾选了核价对比，但未选择报价标准表。\n继续将跳过核价，是否继续？"):
                return
            compare = False

        out_dir = HERE / info["output_dir"]
        out_dir.mkdir(parents=True, exist_ok=True)
        ckpt = out_dir / "checkpoint.jsonl"

        # 断点与任务组合一致性
        feat_file = out_dir / "last_features.json"
        if ckpt.exists():
            last = None
            try:
                last = json.loads(feat_file.read_text(encoding="utf-8"))
            except Exception:
                pass
            if self.var_clear_ckpt.get():
                ckpt.unlink(missing_ok=True)
                self.logln("[断点] 已清除，本次将重新采集全部经销商。")
            elif last is not None and last != features:
                if messagebox.askyesno(
                        "任务组合已变化",
                        "本次勾选的任务与上次不同。\n"
                        "若沿用断点，已完成的经销商会被跳过、新勾选项不会补采。\n\n"
                        "是否清除断点重新采集？（推荐：是）"):
                    ckpt.unlink(missing_ok=True)
                    self.logln("[断点] 已清除（任务组合变化）。")
        feat_file.write_text(json.dumps(features, ensure_ascii=False),
                             encoding="utf-8")

        ov = {
            "dealer_list_xlsx": dealer,
            "standard_xlsx": std if std else None,
            "concurrency": int(self.var_conc.get()),
            "features": features,
            "compare_prices": compare,
            "zip_screenshots": self.var_zip.get(),
        }
        cfg_path = out_dir / "panel_config.json"
        cfg_path.write_text(json.dumps(ov, ensure_ascii=False, indent=2),
                            encoding="utf-8")

        picked = [n for n, v in (("报价", features["pricing"]),
                                 ("软文", features["article_date"]),
                                 ("截图", features["screenshot"]),
                                 ("核价", compare),
                                 ("打包ZIP", self.var_zip.get())) if v]
        self.logln("─" * 56)
        self.logln(f"[启动] 平台={info['label']}  任务={'+'.join(picked)}  "
                   f"并发={ov['concurrency']}")
        self.logln(f"[启动] 名单={Path(dealer).name}  "
                   f"标准={Path(std).name if std else '（无）'}")
        self.logln("─" * 56)

        self._set_state("任务运行中…", running=True)
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self._run_subprocess(
            [sys.executable, "-u", info["script"]],
            env={"WXDJ_CONFIG_JSON": str(cfg_path)},
            on_done=self._task_done)

    def _task_done(self, rc=None):
        self.btn_run.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self._set_state("就绪")

    def stop_task(self):
        if self.proc:
            self.proc.terminate()
            self.logln("[面板] 已发送停止信号（断点已保存，可续跑）。")


def main():
    ControlPanel().root.mainloop()


if __name__ == "__main__":
    main()
