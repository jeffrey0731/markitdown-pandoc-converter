#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF/DOCX -> Markdown 双引擎批量转换工具  V3.0
（MarkItDown + Pandoc，表格错位优化 + 全功能增强版）

在 V1.0 基础上整合了《V1.0修改建议.md》中的全部要点：
  1. 自动根据输入文件生成输出路径
  2. 自动识别文件类型并推荐转换引擎
  3. 实时日志窗口
  4. 启动时自动检测 markitdown / pandoc 是否安装
  5. 拖拽文件支持（可选依赖 tkinterdnd2，未安装则自动降级）
  6. 转换完成后可自动打开输出目录
  7. 底部状态栏
  8. 配置持久化（引擎、输出目录、开关项、窗口大小）
  9. 多文件批量转换
  10. 扫描版 PDF 自动检测提示（可选依赖 pdfplumber）
  11. Markdown 后处理（统一换行符、清理连续空行）

用法：python V3_0.py
依赖：
  必需：Python 3.9+ (tkinter 自带)
  外部命令行工具（至少装一个）：markitdown、pandoc
  可选 pip 包：tkinterdnd2（拖拽）、pdfplumber（扫描件检测）
"""

import os
import sys
import json
import shutil
import platform
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ────────────────────────── 可选依赖：拖拽支持 ──────────────────────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ────────────────────────── 可选依赖：扫描版PDF检测 ──────────────────────────
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".md_converter_v3_config.json")

DOC_EXTS = (".pdf", ".docx")


# ============================================================
#  工具函数
# ============================================================

def get_encoding():
    """适配Windows中文编码，避免中文路径、日志乱码"""
    if sys.platform == "win32":
        return "gbk"
    return "utf-8"


def check_tool_installed(name):
    return shutil.which(name) is not None


def suggest_engine_for_file(path):
    """根据扩展名自动推荐引擎：PDF -> markitdown，DOCX -> pandoc"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return "markitdown"
    if ext == ".docx":
        return "pandoc"
    return "markitdown"


def is_scanned_pdf(path, text_threshold=50):
    """粗略判断PDF是否为扫描件：首页可提取文字过少即视为扫描件。
    未安装 pdfplumber 时返回 None，表示无法判断。"""
    if not PDFPLUMBER_AVAILABLE:
        return None
    if not path.lower().endswith(".pdf"):
        return False
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return False
            first_page_text = pdf.pages[0].extract_text() or ""
            return len(first_page_text.strip()) < text_threshold
    except Exception:
        return None


def post_process_markdown(md_path):
    """Markdown 后处理：统一换行符、清理连续空行"""
    try:
        with open(md_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        while "\n\n\n" in text:
            text = text.replace("\n\n\n", "\n\n")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(text)
        return True
    except Exception:
        return False


def open_folder(path):
    """跨平台打开文件所在目录"""
    folder = os.path.dirname(path) or "."
    try:
        if platform.system() == "Windows":
            os.startfile(folder)  # noqa
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception:
        pass


def auto_output_path(in_path, out_dir=None):
    """根据输入文件名自动生成输出 .md 路径"""
    base = os.path.splitext(os.path.basename(in_path))[0]
    if out_dir:
        return os.path.join(out_dir, base + ".md")
    return os.path.splitext(in_path)[0] + ".md"


# ============================================================
#  配置持久化
# ============================================================

class ConfigManager:
    DEFAULTS = {
        "engine": "markitdown",
        "last_output_dir": "",
        "same_dir_as_input": True,
        "auto_open_folder": True,
        "post_process": True,
        "detect_scanned_pdf": True,
        "geometry": "980x680",
    }

    def __init__(self, path):
        self.path = path
        self.data = dict(self.DEFAULTS)
        self.load()

    def load(self):
        if os.path.isfile(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                self.data.update(saved)
            except Exception:
                pass

    def save(self):
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get(self, key):
        return self.data.get(key, self.DEFAULTS.get(key))

    def set(self, key, value):
        self.data[key] = value


# ============================================================
#  转换核心（子进程调用 markitdown / pandoc CLI）
# ============================================================

def build_command(engine, in_path, out_path):
    out_dir = os.path.dirname(out_path)
    if engine == "markitdown":
        return [
            "markitdown",
            in_path,
            "--enable-table-plugin",
            "--gfm",
            "--preserve-layout",
            "--no-wrap",
            "-o", out_path,
        ]
    else:
        base = os.path.splitext(os.path.basename(out_path))[0]
        media_dir = os.path.join(out_dir, base + "_media")
        os.makedirs(media_dir, exist_ok=True)
        return [
            "pandoc",
            in_path,
            "-t", "gfm",
            "--wrap=none",
            f"--extract-media={media_dir}",
            "-o", out_path,
        ]


def convert_one(engine, in_path, out_path, post_process=True):
    """转换单个文件，返回 (success: bool, message: str)"""
    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    cmd = build_command(engine, in_path, out_path)

    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding=get_encoding(),
            errors="ignore",
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "未知错误").strip()
            return False, err[:1800]

        if post_process:
            post_process_markdown(out_path)

        return True, "转换成功"

    except FileNotFoundError:
        return False, f"系统未找到命令行工具 [{cmd[0]}]，请确认已安装并加入环境变量"
    except Exception as e:
        return False, str(e)


# ============================================================
#  GUI 主程序
# ============================================================

class ConverterApp:
    def __init__(self, root, config: ConfigManager):
        self.root = root
        self.cfg = config
        self.files = []  # 待转换文件列表
        self._build_ui()
        self._check_tools()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------------------------------------------
    # 界面构建
    # ---------------------------------------------------------------
    def _build_ui(self):
        self.root.title("PDF/DOCX -> Markdown 双引擎批量转换器 V3.0")
        self.root.geometry(self.cfg.get("geometry"))
        self.root.minsize(820, 560)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # ---------- 顶部：引擎选择 ----------
        frame_engine = ttk.LabelFrame(self.root, text="转换引擎", padding=10)
        frame_engine.pack(fill=tk.X, padx=12, pady=(12, 6))

        self.engine_var = tk.StringVar(value=self.cfg.get("engine"))
        self.rb_md = ttk.Radiobutton(
            frame_engine, text="MarkItDown（适合 PDF / 扫描件 / OCR，自带表格解析）",
            variable=self.engine_var, value="markitdown",
        )
        self.rb_md.pack(side=tk.LEFT, padx=(0, 16))

        self.rb_pandoc = ttk.Radiobutton(
            frame_engine, text="Pandoc（适合 DOCX / 复杂表格，行列错位更少）",
            variable=self.engine_var, value="pandoc",
        )
        self.rb_pandoc.pack(side=tk.LEFT, padx=(0, 16))

        self.auto_engine_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_engine, text="按文件类型自动推荐引擎（PDF→MarkItDown，DOCX→Pandoc）",
            variable=self.auto_engine_var,
        ).pack(side=tk.LEFT)

        self.tool_status_label = ttk.Label(frame_engine, text="", foreground="#b00000")
        self.tool_status_label.pack(side=tk.RIGHT)

        # ---------- 文件列表区 ----------
        frame_files = ttk.LabelFrame(
            self.root,
            text="待转换文件列表" + ("（支持拖拽文件到此处）" if DND_AVAILABLE else "（未安装 tkinterdnd2，暂不支持拖拽，可用“添加文件”按钮）"),
            padding=10,
        )
        frame_files.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        list_container = ttk.Frame(frame_files)
        list_container.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(list_container, selectmode=tk.EXTENDED, height=8)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)

        if DND_AVAILABLE:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)

        btn_row = ttk.Frame(frame_files)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="添加文件...", command=self._add_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="移除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="清空列表", command=self._clear_files).pack(side=tk.LEFT)

        # ---------- 输出设置 ----------
        frame_out = ttk.LabelFrame(self.root, text="输出设置", padding=10)
        frame_out.pack(fill=tk.X, padx=12, pady=6)

        self.same_dir_var = tk.BooleanVar(value=self.cfg.get("same_dir_as_input"))
        ttk.Checkbutton(
            frame_out, text="输出到源文件所在目录（自动生成同名 .md）",
            variable=self.same_dir_var, command=self._toggle_output_dir,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(frame_out, text="自定义输出目录：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.out_dir_var = tk.StringVar(value=self.cfg.get("last_output_dir"))
        self.out_dir_entry = ttk.Entry(frame_out, textvariable=self.out_dir_var, width=60)
        self.out_dir_entry.grid(row=1, column=1, sticky="we", padx=6, pady=(6, 0))
        self.out_dir_btn = ttk.Button(frame_out, text="浏览...", command=self._choose_out_dir)
        self.out_dir_btn.grid(row=1, column=2, pady=(6, 0))
        frame_out.grid_columnconfigure(1, weight=1)
        self._toggle_output_dir()

        # ---------- 选项 ----------
        frame_opts = ttk.LabelFrame(self.root, text="选项", padding=10)
        frame_opts.pack(fill=tk.X, padx=12, pady=6)

        self.auto_open_var = tk.BooleanVar(value=self.cfg.get("auto_open_folder"))
        ttk.Checkbutton(frame_opts, text="转换完成后自动打开输出目录",
                         variable=self.auto_open_var).pack(side=tk.LEFT, padx=(0, 16))

        self.post_process_var = tk.BooleanVar(value=self.cfg.get("post_process"))
        ttk.Checkbutton(frame_opts, text="自动后处理（统一换行符/清理多余空行）",
                         variable=self.post_process_var).pack(side=tk.LEFT, padx=(0, 16))

        self.detect_scan_var = tk.BooleanVar(value=self.cfg.get("detect_scanned_pdf"))
        cb_scan = ttk.Checkbutton(
            frame_opts,
            text="检测扫描版PDF" + ("" if PDFPLUMBER_AVAILABLE else "（未安装pdfplumber，不可用）"),
            variable=self.detect_scan_var,
        )
        cb_scan.pack(side=tk.LEFT)
        if not PDFPLUMBER_AVAILABLE:
            cb_scan.config(state=tk.DISABLED)
            self.detect_scan_var.set(False)

        # ---------- 转换按钮 + 进度条 ----------
        frame_action = ttk.Frame(self.root, padding="12 4 12 4")
        frame_action.pack(fill=tk.X)

        self.convert_btn = ttk.Button(
            frame_action, text="▶  开始批量转换", command=self._start_convert,
        )
        self.convert_btn.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(frame_action, mode="determinate", length=400)
        self.progress.pack(side=tk.LEFT, padx=12, fill=tk.X, expand=True)

        self.progress_label = ttk.Label(frame_action, text="")
        self.progress_label.pack(side=tk.LEFT)

        # ---------- 日志窗口 ----------
        frame_log = ttk.LabelFrame(self.root, text="日志", padding=6)
        frame_log.pack(fill=tk.BOTH, expand=False, padx=12, pady=(6, 6))

        log_container = ttk.Frame(frame_log)
        log_container.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(log_container, height=8, wrap=tk.WORD,
                                bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=log_scroll.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        log_container.grid_rowconfigure(0, weight=1)
        log_container.grid_columnconfigure(0, weight=1)

        # ---------- 状态栏 ----------
        self.status_var = tk.StringVar(value="就绪 — 请添加文件后点击“开始批量转换”")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN,
                                anchor=tk.W, padding="4 2")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # ---------- 底部提示 ----------
        tip_text = (
            "【表格错位处理提示】 扫描版PDF优先用 MarkItDown；表格密集的 DOCX 优先用 Pandoc；\n"
            "转换后建议用 VSCode + Markdown All in One（Shift+Alt+F）再次格式化对齐表格。"
        )
        ttk.Label(self.root, text=tip_text, foreground="#555555", justify=tk.LEFT,
                  padding="10 4").pack(fill=tk.X, side=tk.BOTTOM)

    # ---------------------------------------------------------------
    # 文件列表操作
    # ---------------------------------------------------------------
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件（可多选）",
            filetypes=[
                ("PDF / Word 文档", "*.pdf;*.docx"),
                ("PDF 文件", "*.pdf"),
                ("Word DOCX", "*.docx"),
                ("所有文件", "*.*"),
            ],
        )
        self._append_files(paths)

    def _on_drop(self, event):
        # tkinterdnd2 的拖拽数据可能包含花括号包裹的多个路径
        raw = event.data
        paths = self.root.tk.splitlist(raw)
        self._append_files(paths)

    def _append_files(self, paths):
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            if not os.path.isfile(p):
                continue
            if p.lower().endswith(DOC_EXTS) and p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, p)
                added += 1
        if added:
            self.status_var.set(f"已添加 {added} 个文件，共 {len(self.files)} 个待转换")
            if self.auto_engine_var.get() and self.files:
                self.engine_var.set(suggest_engine_for_file(self.files[-1]))
        elif paths:
            messagebox.showinfo("提示", "仅支持 .pdf / .docx 文件")

    def _remove_selected(self):
        sel = list(self.listbox.curselection())
        for i in reversed(sel):
            self.listbox.delete(i)
            del self.files[i]
        self.status_var.set(f"共 {len(self.files)} 个待转换文件")

    def _clear_files(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self.status_var.set("已清空文件列表")

    # ---------------------------------------------------------------
    # 输出目录
    # ---------------------------------------------------------------
    def _toggle_output_dir(self):
        state = tk.DISABLED if self.same_dir_var.get() else tk.NORMAL
        self.out_dir_entry.config(state=state)
        self.out_dir_btn.config(state=state)

    def _choose_out_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir_var.set(os.path.normpath(d))

    # ---------------------------------------------------------------
    # 工具检测
    # ---------------------------------------------------------------
    def _check_tools(self):
        has_md = check_tool_installed("markitdown")
        has_pd = check_tool_installed("pandoc")

        missing = []
        if not has_md:
            self.rb_md.config(state=tk.DISABLED)
            missing.append("markitdown")
        if not has_pd:
            self.rb_pandoc.config(state=tk.DISABLED)
            missing.append("pandoc")

        if missing:
            self.tool_status_label.config(text=f"⚠ 未检测到：{', '.join(missing)}（对应功能已禁用）")
        else:
            self.tool_status_label.config(text="✓ markitdown 与 pandoc 均已就绪", foreground="#0a7d0a")

        # 若当前选中的引擎不可用，自动切到可用的那个
        if self.engine_var.get() == "markitdown" and not has_md and has_pd:
            self.engine_var.set("pandoc")
        elif self.engine_var.get() == "pandoc" and not has_pd and has_md:
            self.engine_var.set("markitdown")

    # ---------------------------------------------------------------
    # 日志 / 状态
    # ---------------------------------------------------------------
    def _log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    # ---------------------------------------------------------------
    # 转换主流程
    # ---------------------------------------------------------------
    def _start_convert(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加至少一个文件")
            return

        out_dir = None if self.same_dir_var.get() else self.out_dir_var.get().strip()
        if out_dir and not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录：{e}")
                return

        self.convert_btn.config(state=tk.DISABLED)
        self.progress.config(value=0, maximum=len(self.files))
        self.progress_label.config(text=f"0 / {len(self.files)}")
        self.status_var.set("正在转换...")
        self._log(f"===== 开始批量转换，共 {len(self.files)} 个文件 =====")

        files_snapshot = list(self.files)
        auto_engine = self.auto_engine_var.get()
        fixed_engine = self.engine_var.get()
        post_process = self.post_process_var.get()
        detect_scan = self.detect_scan_var.get() and PDFPLUMBER_AVAILABLE
        auto_open = self.auto_open_var.get()

        def worker():
            success_count = 0
            last_out_path = None
            for idx, in_path in enumerate(files_snapshot, start=1):
                engine = suggest_engine_for_file(in_path) if auto_engine else fixed_engine
                out_path = auto_output_path(in_path, out_dir)

                self.root.after(0, self._log, f"[{idx}/{len(files_snapshot)}] {os.path.basename(in_path)} -> 引擎:{engine}")

                if detect_scan and in_path.lower().endswith(".pdf"):
                    scanned = is_scanned_pdf(in_path)
                    if scanned:
                        self.root.after(0, self._log, "  提示：检测到该PDF可能是扫描件，建议使用 MarkItDown 以获得更好的OCR效果")

                ok, msg = convert_one(engine, in_path, out_path, post_process=post_process)

                if ok:
                    success_count += 1
                    last_out_path = out_path
                    self.root.after(0, self._log, f"  ✓ 成功 -> {out_path}")
                else:
                    self.root.after(0, self._log, f"  ✗ 失败：{msg}")

                self.root.after(0, self._update_progress, idx, len(files_snapshot))

            self.root.after(0, self._convert_finished, success_count, len(files_snapshot), last_out_path, auto_open)

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, done, total):
        self.progress.config(value=done)
        self.progress_label.config(text=f"{done} / {total}")

    def _convert_finished(self, success_count, total, last_out_path, auto_open):
        self.convert_btn.config(state=tk.NORMAL)
        self.status_var.set(f"转换完成：成功 {success_count} / {total}")
        self._log(f"===== 批量转换完成：成功 {success_count} / {total} =====")

        if success_count > 0 and auto_open and last_out_path:
            open_folder(last_out_path)

        if success_count == total:
            messagebox.showinfo("转换完成", f"全部 {total} 个文件转换成功！")
        elif success_count == 0:
            messagebox.showerror("转换失败", "全部文件转换失败，请查看日志了解详情")
        else:
            messagebox.showwarning("部分成功", f"成功 {success_count} 个，失败 {total - success_count} 个，详情见日志")

    # ---------------------------------------------------------------
    # 关闭时保存配置
    # ---------------------------------------------------------------
    def _on_close(self):
        self.cfg.set("engine", self.engine_var.get())
        self.cfg.set("same_dir_as_input", self.same_dir_var.get())
        self.cfg.set("last_output_dir", self.out_dir_var.get().strip())
        self.cfg.set("auto_open_folder", self.auto_open_var.get())
        self.cfg.set("post_process", self.post_process_var.get())
        self.cfg.set("detect_scanned_pdf", self.detect_scan_var.get())
        self.cfg.set("geometry", self.root.geometry())
        self.cfg.save()
        self.root.destroy()


# ============================================================
#  入口
# ============================================================

def main():
    config = ConfigManager(CONFIG_PATH)

    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    ConverterApp(root, config)
    root.mainloop()


if __name__ == "__main__":
    main()
