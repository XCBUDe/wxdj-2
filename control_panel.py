#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
网销点检 — 任务控制面板（图形界面）

双击「启动控制面板.bat」即可打开。功能：
  1. 依赖安装情况检测 + 一键安装
  2. 输入文件选择（经销商名单 / 报价标准），支持点击选择和拖拽（需 tkinterdnd2）
  3. 平台选择：汽车之家 / 易车 / 懂车帝（预留接口，未实现）
  4. 任务勾选：报价采集 / 软文日期 / 头图截图 / 核价对比 / 截图打包
     —— 只勾「软文日期」就只跑软文，不再全量白跑
  5. 实时日志输出 + 断点管理

实现方式：面板把所选配置写成 JSON，通过环境变量 WXDJ_CONFIG_JSON
注入 task.py / yiche_task.py（两脚本均已支持外部配置覆盖），
因此不修改、不破坏原有命令行用法。
"""

import os
import sys
import json
import shutil
import threading
import subprocess
import importlib.util
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

HERE = Path(__file__).resolve().parent

# =============================================================================
# ★ 平台注册表：以后接入懂车帝，只需把 enabled 改 True 并实现 dongchedi_task.py ★
# =============================================================================
PLATFORMS = {
    "autohome": {
        "label": "汽车之家",
        "script": "task.py",
        "output_dir": "output",
        "default_concurrency": 2,
        "max_concurrency": 4,
        "needs_browser": True,      # Playwright 渲染，需要 Chromium/Edge
        "enabled": True,
    },
    "yiche": {
        "label": "易车",
        "script": "yiche_task.py",
        "output_dir": "output_yiche",
        "default_concurrency": 3,
        "max_concurrency": 5,
        "needs_browser": False,     # 纯 requests，无需浏览器
        "enabled": True,
    },
    "dongchedi": {
        "label": "懂车帝（预留接口，未实现）",
        "script": "dongchedi_task.py",
        "output_dir": "output_dongchedi",
        "default_concurrency": 2,
        "max_concurrency": 3,
        "needs_browser": False,
        "enabled": False,           # ← 实现后改 True 即可在面板启用
    },
}

# 依赖清单：(显示名, import 名, 用途说明)
DEPENDENCIES = [
    ("requests",       "requests", "网络请求（两平台必需）"),
    ("openpyxl",       "openpyxl", "Excel 读写（必需）"),
    ("beautifulsoup4", "bs4",      "易车 HTML 解析（必需）"),
    ("pillow",         "PIL",      "图片嵌入 Excel（截图任务）"),
    ("parsel",         "parsel",   "页面解析辅助"),
    ("playwright",     "playwright", "汽车之家浏览器采集（仅汽车之家）"),
    ("tkinterdnd2",    "tkinterdnd2", "面板拖拽支持（可选）"),
]

PIP_MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
PW_MIRROR = "https://npmmirror.com/mirrors/playwright/"

# ── 拖拽支持（可选） ────────────────────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    _HAS_DND = False


def _module_ok(import_name: str) -> bool:
    try:
        return importlib.util.find_spec(import_name) is not None
    except Exception:
        return False


def _find_browser() -> str:
    """检测系统可用浏览器（汽车之家用）。"""
    candidates = [
        ("Edge",   r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        ("Edge",   r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        ("Chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        ("Chrome", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    ]
    for name, p in candidates:
        if Path(p).exists():
            return name
    for name, exe in [("Edge", "msedge"), ("Chrome", "chrome"),
                      ("Chromium", "chromium"), ("Chromium", "chromium-browser")]:
        if shutil.which(exe):
            return name
    # Playwright 自带 Chromium
    pw_dirs = [
        Path.home() / "AppData/Local/ms-playwright",
        Path.home() / ".cache/ms-playwright",
    ]
    for d in pw_dirs:
        if d.exists() and any(d.glob("chromium*")):
            return "Playwright-Chromium"
    return ""


class ControlPanel:

    def __init__(self):
        if _HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("网销点检 · 任务控制面板")
        self.root.geometry("860x780")
        self.root.minsize(760, 640)

        self.proc = None          # 当前运行的子进程
        self.dep_labels = {}      # 依赖状态标签

        self._build_ui()
        self.refresh_deps()
        self.autodetect_files()

    # ──────────────────────────────────────────────────────────────────
    # UI 构建
    # ──────────────────────────────────────────────────────────────────
    def _build_ui(self):
        pad = dict(padx=10, pady=4)

        # ===== 1. 依赖状态 =====
        f_dep = ttk.LabelFrame(self.root, text=" ① 依赖安装情况 ")
        f_dep.pack(fill="x", **pad)

        grid = ttk.Frame(f_dep)
        grid.pack(fill="x", padx=8, pady=4)
        cols = 4
        for i, (disp, imp, note) in enumerate(DEPENDENCIES):
            r, c = divmod(i, cols)
            lbl = ttk.Label(grid, text=f"… {disp}", width=24)
            lbl.grid(row=r, column=c, sticky="w", padx=4, pady=2)
            self.dep_labels[imp] = (lbl, disp)
        self.lbl_browser = ttk.Label(grid, text="… 浏览器", width=24)
        self.lbl_browser.grid(row=(len(DEPENDENCIES)) // cols,
                              column=(len(DEPENDENCIES)) % cols,
                              sticky="w", padx=4, pady=2)

        btns = ttk.Frame(f_dep)
        btns.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(btns, text="重新检测", command=self.refresh_deps).pack(side="left", padx=4)
        ttk.Button(btns, text="一键安装依赖（清华镜像）",
                   command=self.install_deps).pack(side="left", padx=4)
        ttk.Button(btns, text="安装 Chromium 浏览器（仅汽车之家需要）",
                   command=self.install_browser).pack(side="left", padx=4)

        # ===== 2. 输入文件 =====
        dnd_hint = "（可直接拖文件到输入框）" if _HAS_DND else "（安装 tkinterdnd2 后可拖拽）"
        f_in = ttk.LabelFrame(self.root, text=f" ② 输入文件 {dnd_hint} ")
        f_in.pack(fill="x", **pad)

        self.var_list = tk.StringVar()
        self.var_std = tk.StringVar()
        self._file_row(f_in, "经销商名单 *", self.var_list)
        self._file_row(f_in, "报价标准（不核价可留空）", self.var_std)
        ttk.Button(f_in, text="自动识别本文件夹 xlsx",
                   command=self.autodetect_files).pack(anchor="w", padx=12, pady=(0, 6))

        # ===== 3. 平台 =====
        f_pf = ttk.LabelFrame(self.root, text=" ③ 平台（单次任务只跑单个平台） ")
        f_pf.pack(fill="x", **pad)
        self.var_platform = tk.StringVar(value="yiche")
        row = ttk.Frame(f_pf)
        row.pack(fill="x", padx=8, pady=4)
        for key, info in PLATFORMS.items():
            rb = ttk.Radiobutton(row, text=info["label"], value=key,
                                 variable=self.var_platform,
                                 command=self.on_platform_change)
            rb.pack(side="left", padx=10)
            if not info["enabled"]:
                rb.state(["disabled"])

        # ===== 4. 任务选择 =====
        f_task = ttk.LabelFrame(
            self.root, text=" ④ 任务选择（只勾需要的，避免全量白跑） ")
        f_task.pack(fill="x", **pad)

        self.var_pricing = tk.BooleanVar(value=True)
        self.var_article = tk.BooleanVar(value=True)
        self.var_shot = tk.BooleanVar(value=True)
        self.var_compare = tk.BooleanVar(value=True)
        self.var_zip = tk.BooleanVar(value=True)
        self.var_clear_ckpt = tk.BooleanVar(value=False)

        row1 = ttk.Frame(f_task); row1.pack(fill="x", padx=8, pady=2)
        ttk.Checkbutton(row1, text="车型报价采集", variable=self.var_pricing,
                        command=self._sync_task_state).pack(side="left", padx=8)
        ttk.Checkbutton(row1, text="最新软文日期", variable=self.var_article).pack(side="left", padx=8)
        self.cb_shot = ttk.Checkbutton(row1, text="头图截图", variable=self.var_shot,
                                       command=self._sync_task_state)
        self.cb_shot.pack(side="left", padx=8)

        row2 = ttk.Frame(f_task); row2.pack(fill="x", padx=8, pady=2)
        self.cb_compare = ttk.Checkbutton(row2, text="核价对比着色（需报价采集+标准表）",
                                          variable=self.var_compare)
        self.cb_compare.pack(side="left", padx=8)
        self.cb_zip = ttk.Checkbutton(row2, text="截图打包 ZIP（需头图截图）",
                                      variable=self.var_zip)
        self.cb_zip.pack(side="left", padx=8)

        row3 = ttk.Frame(f_task); row3.pack(fill="x", padx=8, pady=(2, 6))
        ttk.Button(row3, text="全选", width=6,
                   command=lambda: self._set_all(True)).pack(side="left", padx=8)
        ttk.Button(row3, text="全不选", width=6,
                   command=lambda: self._set_all(False)).pack(side="left", padx=4)
        ttk.Label(row3, text="并发数:").pack(side="left", padx=(20, 4))
        self.var_conc = tk.IntVar(value=PLATFORMS["yiche"]["default_concurrency"])
        self.spin_conc = ttk.Spinbox(row3, from_=1, to=5, width=4,
                                     textvariable=self.var_conc)
        self.spin_conc.pack(side="left")
        ttk.Checkbutton(row3, text="清除断点重新采集（换任务组合时建议勾选）",
                        variable=self.var_clear_ckpt).pack(side="left", padx=20)

        # ===== 5. 执行 =====
        f_run = ttk.Frame(self.root)
        f_run.pack(fill="x", **pad)
        self.btn_run = ttk.Button(f_run, text="▶ 开始执行", command=self.start_task)
        self.btn_run.pack(side="left", padx=8)
        self.btn_stop = ttk.Button(f_run, text="■ 停止", command=self.stop_task,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        ttk.Button(f_run, text="打开输出文件夹",
                   command=self.open_output).pack(side="left", padx=12)

        # ===== 6. 日志 =====
        f_log = ttk.LabelFrame(self.root, text=" 运行日志 ")
        f_log.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(f_log, height=14, state="disabled",
                                             font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

        self.on_platform_change()

    def _file_row(self, parent, label, var):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=22).pack(side="left")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="选择文件…", width=10,
                   command=lambda: self._pick_file(var)).pack(side="left", padx=4)
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
    # 逻辑
    # ──────────────────────────────────────────────────────────────────
    def logln(self, text=""):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def refresh_deps(self):
        for imp, (lbl, disp) in self.dep_labels.items():
            ok = _module_ok(imp)
            lbl.config(text=("✅ " if ok else "❌ ") + disp,
                       foreground=("green" if ok else "red"))
        b = _find_browser()
        if b:
            self.lbl_browser.config(text=f"✅ 浏览器({b})", foreground="green")
        else:
            self.lbl_browser.config(text="❌ 浏览器(未检测到)", foreground="red")

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
        # 核价依赖报价采集；ZIP 依赖截图
        if not self.var_pricing.get():
            self.var_compare.set(False)
            self.cb_compare.state(["disabled"])
        else:
            self.cb_compare.state(["!disabled"])
        if not self.var_shot.get():
            self.var_zip.set(False)
            self.cb_zip.state(["disabled"])
        else:
            self.cb_zip.state(["!disabled"])

    def _set_all(self, value: bool):
        for v in (self.var_pricing, self.var_article, self.var_shot,
                  self.var_compare, self.var_zip):
            v.set(value)
        self._sync_task_state()

    def open_output(self):
        out = HERE / PLATFORMS[self.var_platform.get()]["output_dir"]
        out.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(out)          # noqa
        else:
            subprocess.Popen(["xdg-open", str(out)])

    # ── 子进程统一执行（依赖安装 / 任务运行共用） ──────────────────────
    def _run_subprocess(self, cmd, env=None, on_done=None):
        def _reader():
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
                self.root.after(0, self.logln,
                                f"—— 进程结束（返回码 {rc}） ——")
            except Exception as ex:
                self.root.after(0, self.logln, f"[面板错误] {ex}")
            finally:
                self.proc = None
                if on_done:
                    self.root.after(0, on_done)
        threading.Thread(target=_reader, daemon=True).start()

    def install_deps(self):
        if self.proc:
            messagebox.showwarning("提示", "已有任务在运行，请先停止。"); return
        self.logln("[依赖] 开始安装（清华镜像，失败自动回退官方源）…")
        py = sys.executable
        cmd = (f'"{py}" -m pip install -r requirements.txt -i {PIP_MIRROR}'
               f' || "{py}" -m pip install -r requirements.txt')
        self._run_subprocess(["cmd", "/c", cmd] if sys.platform == "win32"
                             else ["bash", "-c", cmd.replace('"', "'")],
                             on_done=self.refresh_deps)

    def install_browser(self):
        if self.proc:
            messagebox.showwarning("提示", "已有任务在运行，请先停止。"); return
        self.logln("[浏览器] 开始安装 Playwright Chromium（国内镜像）…")
        self._run_subprocess([sys.executable, "-m", "playwright", "install", "chromium"],
                             env={"PLAYWRIGHT_DOWNLOAD_HOST": PW_MIRROR},
                             on_done=self.refresh_deps)

    # ── 启动点检任务 ───────────────────────────────────────────────────
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

        # 断点 与 任务组合 一致性检查
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
                        "检测到本次勾选的任务与上次不同，\n"
                        "若沿用断点，已完成的经销商会被跳过、新勾选项不会补采。\n\n"
                        "是否清除断点重新采集？（推荐：是）"):
                    ckpt.unlink(missing_ok=True)
                    self.logln("[断点] 已清除（任务组合变化）。")
        feat_file.write_text(json.dumps(features, ensure_ascii=False),
                             encoding="utf-8")

        # 写面板配置 → 环境变量注入入口脚本
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
        self.logln("=" * 56)
        self.logln(f"[启动] 平台={info['label']}  任务={'+'.join(picked)}  "
                   f"并发={ov['concurrency']}")
        self.logln(f"[启动] 名单={Path(dealer).name}  "
                   f"标准={Path(std).name if std else '（无）'}")
        self.logln("=" * 56)

        self.btn_run.state(["disabled"])
        self.btn_stop.state(["!disabled"])
        self._run_subprocess(
            [sys.executable, "-u", info["script"]],
            env={"WXDJ_CONFIG_JSON": str(cfg_path)},
            on_done=self._task_done)

    def _task_done(self):
        self.btn_run.state(["!disabled"])
        self.btn_stop.state(["disabled"])

    def stop_task(self):
        if self.proc:
            self.proc.terminate()
            self.logln("[面板] 已发送停止信号（断点已保存，可续跑）。")


def main():
    app = ControlPanel()
    app.root.mainloop()


if __name__ == "__main__":
    main()
