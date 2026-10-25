# -*- coding: utf-8 -*-
"""订单 / 发票比对工具 — 图形界面。"""

from __future__ import annotations

import os
import shutil
import sys
import threading
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from reconcile import (
    DEFAULT_ACCOUNT_MONTHS,
    peek_period_from_invoice_file,
    run_reconcile,
)

APP_TITLE = "订单发票比对工具"
DEFAULT_RESULT_NAME = "比对结果.xlsx"

# 简约扁平配色
COLOR_BG = "#F7F8FA"
COLOR_SURFACE = "#FFFFFF"
COLOR_TEXT = "#1F2937"
COLOR_MUTED = "#6B7280"
COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_HOVER = "#1D4ED8"
COLOR_BORDER = "#E5E7EB"
COLOR_LOG_BG = "#F9FAFB"
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 10, "bold")
FONT_SMALL = ("Microsoft YaHei UI", 9)

FORMAT_HELP_TEXT = """【订单数据】
首行表头须包含以下列（可有多余列，列顺序不限）：
  · 交货日期 — 建议 YYYY-MM-DD / YYYY/MM/DD
  · 型号
  · 数量 — 数值
  · 用户价格 — 数值（申请含税价）

【发票数据】
首行表头须包含以下列（可有多余列，列顺序不限）：
  · 开票日期 — 建议 YYYY-MM-DD / YYYY/MM/DD
  · 规格型号
  · 数量 — 数值
  · 单价 — 数值（未税）；程序按 单价×1.13 计算含税单价

【其他说明】
  · 自动选用含上述必填列的第一张工作表
  · 型号与规格型号匹配时忽略大小写与全部空白
  · 可点击「下载订单模板 / 下载发票模板」获取带样例的 Excel
"""


def _parse_date(text: str, label: str):
    s = (text or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"{label}格式无效，请使用 YYYY-MM-DD：{s}")


def _resource_dir() -> Path:
    """开发态为脚本目录；打包后为 PyInstaller 解压目录或 exe 旁。"""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            p = Path(meipass) / "templates"
            if p.is_dir():
                return p
        return Path(sys.executable).resolve().parent / "templates"
    return Path(__file__).resolve().parent / "templates"


def _template_path(filename: str) -> Path:
    path = _resource_dir() / filename
    if not path.is_file():
        raise FileNotFoundError(f"未找到模板文件：{path}")
    return path


class ReconcileApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("820x640")
        self.minsize(720, 520)
        self.resizable(True, True)
        self.configure(bg=COLOR_BG)

        self.order_var = tk.StringVar()
        self.invoice_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.period_start_var = tk.StringVar(value="")
        self.period_end_var = tk.StringVar(value="")
        self.account_months_var = tk.StringVar(value=str(DEFAULT_ACCOUNT_MONTHS))
        self.status_var = tk.StringVar(
            value="请先选择发票与订单文件；上传发票后将自动推断查询周期。"
        )
        self.progress_pct_var = tk.StringVar(value="0%")
        self._busy = False

        self._apply_theme()
        self._build_ui()
        self._set_busy(False)

    def _apply_theme(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=FONT_UI, background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_SURFACE)
        style.configure(
            "TLabel",
            background=COLOR_BG,
            foreground=COLOR_TEXT,
            font=FONT_UI,
        )
        style.configure(
            "Card.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            font=FONT_UI,
        )
        style.configure(
            "Title.TLabel",
            background=COLOR_BG,
            foreground=COLOR_TEXT,
            font=FONT_TITLE,
        )
        style.configure(
            "Section.TLabel",
            background=COLOR_BG,
            foreground=COLOR_TEXT,
            font=FONT_SECTION,
        )
        style.configure(
            "Muted.TLabel",
            background=COLOR_BG,
            foreground=COLOR_MUTED,
            font=FONT_SMALL,
        )
        style.configure(
            "CardMuted.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_MUTED,
            font=FONT_SMALL,
        )
        style.configure(
            "Link.TLabel",
            background=COLOR_BG,
            foreground=COLOR_PRIMARY,
            font=FONT_SMALL,
            cursor="hand2",
        )
        style.configure(
            "Pct.TLabel",
            background=COLOR_BG,
            foreground=COLOR_MUTED,
            font=FONT_SMALL,
            width=5,
            anchor="e",
        )

        style.configure(
            "TLabelframe",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            relief="solid",
            borderwidth=1,
        )
        style.configure(
            "TLabelframe.Label",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            font=FONT_SECTION,
        )

        style.configure(
            "TEntry",
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            insertcolor=COLOR_TEXT,
            padding=6,
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", COLOR_PRIMARY)],
            lightcolor=[("focus", COLOR_PRIMARY)],
            darkcolor=[("focus", COLOR_PRIMARY)],
        )

        style.configure(
            "TButton",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            focuscolor=COLOR_BORDER,
            padding=(12, 6),
            font=FONT_UI,
        )
        style.map(
            "TButton",
            background=[("active", COLOR_BORDER), ("disabled", COLOR_BG)],
            foreground=[("disabled", COLOR_MUTED)],
        )

        style.configure(
            "Primary.TButton",
            background=COLOR_PRIMARY,
            foreground="#FFFFFF",
            bordercolor=COLOR_PRIMARY,
            lightcolor=COLOR_PRIMARY,
            darkcolor=COLOR_PRIMARY,
            focuscolor=COLOR_PRIMARY,
            padding=(16, 8),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COLOR_PRIMARY_HOVER),
                ("disabled", "#93C5FD"),
            ],
            foreground=[("disabled", "#FFFFFF")],
            bordercolor=[
                ("active", COLOR_PRIMARY_HOVER),
                ("disabled", "#93C5FD"),
            ],
        )

        style.configure(
            "Ghost.TButton",
            background=COLOR_BG,
            foreground=COLOR_MUTED,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=(12, 8),
            font=FONT_UI,
        )
        style.map(
            "Ghost.TButton",
            background=[("active", COLOR_BORDER)],
            foreground=[("active", COLOR_TEXT)],
        )

        style.configure(
            "Link.TButton",
            background=COLOR_BG,
            foreground=COLOR_PRIMARY,
            borderwidth=0,
            focuscolor=COLOR_BG,
            padding=(4, 2),
            font=FONT_SMALL,
        )
        style.map(
            "Link.TButton",
            foreground=[("active", COLOR_PRIMARY_HOVER)],
            background=[("active", COLOR_BG)],
        )

        style.configure(
            "Horizontal.TProgressbar",
            troughcolor=COLOR_BORDER,
            background=COLOR_PRIMARY,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_PRIMARY,
            darkcolor=COLOR_PRIMARY,
            thickness=8,
        )

        style.configure(
            "Vertical.TScrollbar",
            background=COLOR_BORDER,
            troughcolor=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            arrowcolor=COLOR_MUTED,
        )

    def _build_ui(self):
        root = ttk.Frame(self, padding=20)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor=tk.W)

        help_row = ttk.Frame(root)
        help_row.pack(fill=tk.X, pady=(0, 14))
        ttk.Button(
            help_row, text="格式说明", style="Link.TButton", command=self._show_format_help
        ).pack(side=tk.LEFT)
        ttk.Label(help_row, text="·", style="Muted.TLabel").pack(side=tk.LEFT, padx=4)
        ttk.Button(
            help_row,
            text="下载订单模板",
            style="Link.TButton",
            command=lambda: self._download_template("订单数据模板.xlsx", "订单数据模板.xlsx"),
        ).pack(side=tk.LEFT)
        ttk.Label(help_row, text="·", style="Muted.TLabel").pack(side=tk.LEFT, padx=4)
        ttk.Button(
            help_row,
            text="下载发票模板",
            style="Link.TButton",
            command=lambda: self._download_template("发票数据模板.xlsx", "发票数据模板.xlsx"),
        ).pack(side=tk.LEFT)

        form = ttk.Frame(root)
        form.pack(fill=tk.X)

        self._add_file_row(
            form,
            0,
            "订单数据文件",
            self.order_var,
            self._browse_order,
            hint="需含列：交货日期、型号、数量、用户价格",
        )
        self._add_file_row(
            form,
            1,
            "发票数据文件",
            self.invoice_var,
            self._browse_invoice,
            hint="需含列：开票日期、规格型号、数量、单价；选择后自动推断周期",
        )
        self._add_file_row(
            form,
            2,
            "结果保存路径",
            self.output_var,
            self._browse_output,
            hint="导出为 .xlsx",
            save=True,
        )

        period = ttk.LabelFrame(root, text="  发票查询周期与账期  ", padding=12)
        period.pack(fill=tk.X, pady=(14, 0))

        ttk.Label(period, text="周期起始日", style="Card.TLabel").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Entry(period, textvariable=self.period_start_var, width=14).grid(
            row=0, column=1, sticky=tk.W, padx=(8, 16)
        )
        ttk.Label(period, text="周期结束日", style="Card.TLabel").grid(
            row=0, column=2, sticky=tk.W
        )
        ttk.Entry(period, textvariable=self.period_end_var, width=14).grid(
            row=0, column=3, sticky=tk.W, padx=(8, 16)
        )
        ttk.Label(period, text="账期(月)", style="Card.TLabel").grid(
            row=0, column=4, sticky=tk.W
        )
        ttk.Entry(period, textvariable=self.account_months_var, width=6).grid(
            row=0, column=5, sticky=tk.W, padx=(8, 0)
        )
        ttk.Label(
            period,
            text="默认空；选择发票后按开票日所在月月初～月末自动填入，可再手动修改。账期默认 1 个月。",
            style="CardMuted.TLabel",
        ).grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=(8, 0))
        ttk.Label(
            period,
            text="边界：期初豁免截止=起始+N月-1天；期末豁免起始=结束-N月+1天",
            style="CardMuted.TLabel",
        ).grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(4, 0))

        btns = ttk.Frame(root)
        btns.pack(fill=tk.X, pady=(16, 10))
        ttk.Frame(btns).pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.exit_btn = ttk.Button(
            btns, text="退出", style="Ghost.TButton", command=self.destroy, width=10
        )
        self.open_result_btn = ttk.Button(
            btns,
            text="打开比对结果",
            style="Ghost.TButton",
            command=self._on_open_result,
            width=14,
        )
        self.export_btn = ttk.Button(
            btns,
            text="导出比对结果",
            style="Primary.TButton",
            command=self._on_export,
            width=16,
        )
        # pack 顺序：先退出 → 打开 → 导出，配合 side=RIGHT → [导出][打开][退出]
        self.exit_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.open_result_btn.pack(side=tk.RIGHT, padx=(8, 0))
        self.export_btn.pack(side=tk.RIGHT)
        self.output_var.trace_add("write", lambda *_: self._refresh_open_result_btn())

        prog_row = ttk.Frame(root)
        prog_row.pack(fill=tk.X, pady=(0, 10))
        self.progress = ttk.Progressbar(
            prog_row, mode="determinate", maximum=100, value=0
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Label(prog_row, textvariable=self.progress_pct_var, style="Pct.TLabel").pack(
            side=tk.RIGHT
        )

        log_frame = ttk.LabelFrame(root, text="  运行日志  ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log = tk.Text(
            log_frame,
            height=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg=COLOR_LOG_BG,
            fg=COLOR_TEXT,
            insertbackground=COLOR_TEXT,
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            font=FONT_SMALL,
            padx=6,
            pady=4,
        )
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(root, textvariable=self.status_var, style="Muted.TLabel").pack(
            anchor=tk.W, pady=(10, 0)
        )

    def _add_file_row(self, parent, row, label, var, command, hint="", save=False):
        ttk.Label(parent, text=label, style="Section.TLabel").grid(
            row=row * 2, column=0, sticky=tk.W, pady=(0, 2)
        )
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row * 2, column=1, sticky=tk.EW, padx=(12, 8), pady=(0, 2))
        ttk.Button(parent, text="浏览…", command=command, width=10).grid(
            row=row * 2, column=2, pady=(0, 2)
        )
        if hint:
            ttk.Label(parent, text=hint, style="Muted.TLabel").grid(
                row=row * 2 + 1, column=1, sticky=tk.W, pady=(0, 10)
            )
        parent.columnconfigure(1, weight=1)

    def _show_format_help(self):
        win = tk.Toplevel(self)
        win.title("数据格式说明")
        win.configure(bg=COLOR_BG)
        win.transient(self)
        win.grab_set()
        win.geometry("520x420")
        win.minsize(420, 320)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="数据格式要求", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="请按下列必填列准备 Excel，也可直接下载模板后替换样例。",
            style="Muted.TLabel",
        ).pack(anchor=tk.W, pady=(4, 10))

        text = tk.Text(
            frame,
            wrap=tk.WORD,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            relief=tk.FLAT,
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            font=FONT_SMALL,
            padx=10,
            pady=10,
        )
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(tk.END, FORMAT_HELP_TEXT)
        text.configure(state=tk.DISABLED)

        ttk.Button(frame, text="关闭", style="Ghost.TButton", command=win.destroy).pack(
            anchor=tk.E, pady=(12, 0)
        )
        win.focus_set()

    def _download_template(self, resource_name: str, default_name: str):
        try:
            src = _template_path(resource_name)
        except FileNotFoundError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        path = filedialog.asksaveasfilename(
            title=f"保存{default_name}",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            shutil.copy2(src, path)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"保存模板失败：\n{exc}")
            return
        self._append_log(f"已下载模板：{path}")
        self.status_var.set(f"模板已保存：{Path(path).name}")
        messagebox.showinfo(APP_TITLE, f"模板已保存到：\n{path}")

    def _browse_order(self):
        path = filedialog.askopenfilename(
            title="选择订单数据文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.order_var.set(path)
            self._maybe_default_output()

    def _browse_invoice(self):
        path = filedialog.askopenfilename(
            title="选择发票数据文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.invoice_var.set(path)
            self._maybe_default_output()
            self._infer_period_from_invoice(path)

    def _infer_period_from_invoice(self, invoice_path: str):
        """上传发票后：按开票日所在月月初～月末填入周期起止。"""
        try:
            months = int(self.account_months_var.get().strip() or DEFAULT_ACCOUNT_MONTHS)
            if months < 0:
                months = DEFAULT_ACCOUNT_MONTHS
                self.account_months_var.set(str(months))
        except ValueError:
            months = DEFAULT_ACCOUNT_MONTHS
            self.account_months_var.set(str(months))

        try:
            cfg = peek_period_from_invoice_file(Path(invoice_path), months)
        except Exception as exc:
            self._append_log(f"自动推断查询周期失败：{exc}")
            messagebox.showwarning(APP_TITLE, f"无法根据发票推断查询周期：\n{exc}")
            return

        self.period_start_var.set(str(cfg.start))
        self.period_end_var.set(str(cfg.end))
        self._append_log(
            f"已根据发票推断查询周期：{cfg.start} ~ {cfg.end}（账期 {months} 个月）；"
            f"期初豁免截止 {cfg.grace_invoice_until}，期末豁免起始 {cfg.grace_order_from}"
        )
        self.status_var.set(
            f"已推断周期 {cfg.start} ~ {cfg.end}，可手动修改后导出。"
        )

    def _browse_output(self):
        initial = self.output_var.get().strip()
        initialdir = str(Path(initial).parent) if initial else None
        initialfile = Path(initial).name if initial else DEFAULT_RESULT_NAME
        path = filedialog.asksaveasfilename(
            title="保存比对结果",
            defaultextension=".xlsx",
            initialfile=initialfile,
            initialdir=initialdir,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_var.set(path)

    def _maybe_default_output(self):
        if self.output_var.get().strip():
            return
        base = self.order_var.get().strip() or self.invoice_var.get().strip()
        if base:
            self.output_var.set(str(Path(base).with_name(DEFAULT_RESULT_NAME)))

    def _append_log(self, text: str):
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _set_progress(self, pct: int, message: str = ""):
        pct = max(0, min(100, int(pct)))
        self.progress["value"] = pct
        self.progress_pct_var.set(f"{pct}%")
        if message:
            self.status_var.set(message)

    def _on_progress(self, pct: int, message: str):
        self.after(0, lambda p=pct, m=message: self._set_progress(p, m))

    def _set_busy(self, busy: bool):
        self._busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self.export_btn.configure(state=state)
        if busy:
            self.open_result_btn.configure(state=tk.DISABLED)
            self._set_progress(0, "正在比对，请稍候…")
        else:
            self._refresh_open_result_btn()

    def _refresh_open_result_btn(self):
        if getattr(self, "_busy", False):
            self.open_result_btn.configure(state=tk.DISABLED)
            return
        path = self.output_var.get().strip()
        ok = bool(path) and Path(path).is_file()
        self.open_result_btn.configure(state=tk.NORMAL if ok else tk.DISABLED)

    def _on_open_result(self):
        path = self.output_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "请先指定结果保存路径，并完成一次导出。")
            return
        p = Path(path)
        if not p.is_file():
            messagebox.showwarning(
                APP_TITLE,
                "比对结果文件尚不存在。\n\n"
                "请先点击「导出比对结果」生成文件，或确认路径是否正确。\n"
                f"当前路径：{p}",
            )
            self._refresh_open_result_btn()
            return
        try:
            os.startfile(str(p))  # noqa: S606 — Windows 用默认程序打开
        except OSError as exc:
            messagebox.showerror(
                APP_TITLE,
                "无法打开比对结果文件。\n\n"
                "可能原因：\n"
                "· 未安装 Excel / WPS，或未关联 .xlsx 默认打开方式\n"
                "· 文件路径无效或已被移动/删除\n"
                "· 系统权限不足\n\n"
                f"路径：{p}\n详情：{exc}",
            )

    def _explain_error(self, exc: BaseException) -> str:
        """根据异常类型给出可读原因说明。"""
        msg = str(exc)
        if isinstance(exc, PermissionError) or (
            isinstance(exc, OSError)
            and getattr(exc, "errno", None) in (13, 32)
        ):
            return (
                "无法写入结果文件（权限不足 / 文件被占用）。\n\n"
                "可能原因：\n"
                "1. 该结果文件正被 Excel 或 WPS 打开（最常见）\n"
                "2. 文件被其他程序锁定，或只读属性已勾选\n"
                "3. 对目标文件夹没有写入权限\n"
                "4. 杀毒/同步盘（如 OneDrive）正在占用该文件\n\n"
                "建议：关闭已打开的同名工作簿后重试，或更换「结果保存路径」另存。\n\n"
                f"详情：{msg}"
            )
        if isinstance(exc, FileNotFoundError):
            return (
                "找不到所需文件。\n\n"
                "可能原因：\n"
                "1. 订单或发票路径不正确，文件已移动/删除\n"
                "2. 结果目录不存在且无法自动创建\n"
                "3. 网络盘/U 盘已断开\n\n"
                f"详情：{msg}"
            )
        if isinstance(exc, ValueError):
            return (
                "数据内容不符合要求。\n\n"
                "可能原因：\n"
                "1. 缺少必填列（订单：交货日期/型号/数量/用户价格；"
                "发票：开票日期/规格型号/数量/单价）\n"
                "2. 工作表无有效数据行\n"
                "3. 日期、数量、价格格式无法解析\n"
                "4. 查询周期/账期填写无效\n\n"
                "可点「格式说明」或「下载模板」对照检查。\n\n"
                f"详情：{msg}"
            )
        if isinstance(exc, KeyError):
            return (
                "读取表格列时出错。\n\n"
                "可能原因：表头名称与要求不一致，或列被误删/改名。\n"
                "请对照「格式说明」检查首行表头。\n\n"
                f"详情：{msg}"
            )
        return (
            "导出过程中发生错误。\n\n"
            "可能原因：\n"
            "1. 订单/发票 Excel 损坏或正在被占用\n"
            "2. 表头或数据格式不符合要求\n"
            "3. 磁盘空间不足，或目标路径不可写\n"
            "4. 程序异常（详见下方信息）\n\n"
            f"详情：{msg}"
        )

    def _on_export(self):
        order = self.order_var.get().strip()
        invoice = self.invoice_var.get().strip()
        output = self.output_var.get().strip()

        if not order:
            messagebox.showwarning(APP_TITLE, "请先选择订单数据文件。")
            return
        if not invoice:
            messagebox.showwarning(APP_TITLE, "请先选择发票数据文件。")
            return
        if not output:
            messagebox.showwarning(APP_TITLE, "请先指定结果保存路径。")
            return

        try:
            months = int(self.account_months_var.get().strip())
            if months < 0:
                raise ValueError("账期不能为负")
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, f"账期月数无效：{exc}")
            return

        try:
            period_start = _parse_date(self.period_start_var.get(), "周期起始日")
            period_end = _parse_date(self.period_end_var.get(), "周期结束日")
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        # 若仍为空（未成功推断），导出时再按发票推断
        if period_start is None or period_end is None:
            try:
                cfg = peek_period_from_invoice_file(Path(invoice), months)
                period_start, period_end = cfg.start, cfg.end
                self.period_start_var.set(str(cfg.start))
                self.period_end_var.set(str(cfg.end))
                self._append_log(
                    f"导出前补全推断周期：{cfg.start} ~ {cfg.end}"
                )
            except Exception as exc:
                messagebox.showwarning(
                    APP_TITLE,
                    f"查询周期为空且无法推断，请填写周期起止日。\n{exc}",
                )
                return

        out_path = Path(output)
        if out_path.exists():
            if not messagebox.askyesno(
                APP_TITLE, f"文件已存在，是否覆盖？\n{out_path}"
            ):
                return
            if not self._can_write_path(out_path):
                if messagebox.askyesno(
                    APP_TITLE,
                    "无法覆盖写入结果文件。\n\n"
                    "可能原因：\n"
                    "1. 该文件正被 Excel / WPS 打开（最常见）\n"
                    "2. 文件为只读，或对文件夹无写入权限\n"
                    "3. 杀毒/网盘同步正在占用文件\n\n"
                    f"路径：{out_path}\n\n"
                    "是否另存为其他文件名？",
                ):
                    alt = self._ask_alternate_output(out_path)
                    if not alt:
                        return
                    out_path = alt
                    output = str(alt)
                    self.output_var.set(output)
                else:
                    return

        self._start_export(
            order, invoice, output, period_start, period_end, months
        )

    def _can_write_path(self, path: Path) -> bool:
        """探测目标文件是否可覆盖写入（被 Excel 打开时通常会 PermissionError）。"""
        try:
            with open(path, "a+b"):
                return True
        except OSError:
            return False

    def _ask_alternate_output(self, current: Path) -> Path | None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested = current.with_name(f"{current.stem}_{stamp}{current.suffix}")
        path = filedialog.asksaveasfilename(
            title="另存比对结果",
            defaultextension=".xlsx",
            initialfile=suggested.name,
            initialdir=str(current.parent),
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return None
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return Path(path)

    def _start_export(
        self, order, invoice, output, period_start, period_end, months
    ):
        self._set_busy(True)
        self._append_log("-" * 48)
        self._append_log(f"订单：{order}")
        self._append_log(f"发票：{invoice}")
        self._append_log(f"输出：{output}")
        self._append_log(
            f"查询周期：{period_start} ~ {period_end}；账期：{months}月"
        )

        def worker():
            try:
                stats = run_reconcile(
                    order,
                    invoice,
                    output,
                    save_preprocess_files=False,
                    period_start=period_start,
                    period_end=period_end,
                    account_months=months,
                    progress_callback=self._on_progress,
                )
                self.after(0, lambda s=stats: self._on_success(s))
            except PermissionError as exc:
                err = "".join(
                    traceback.format_exception(type(exc), exc, exc.__traceback__)
                )
                self.after(
                    0,
                    lambda e=exc, d=err: self._on_permission_error(e, d),
                )
            except OSError as exc:
                err = "".join(
                    traceback.format_exception(type(exc), exc, exc.__traceback__)
                )
                if getattr(exc, "errno", None) in (13, 32):
                    self.after(
                        0,
                        lambda e=exc, d=err: self._on_permission_error(e, d),
                    )
                else:
                    self.after(0, lambda e=exc, d=err: self._on_error(e, d))
            except Exception as exc:
                err = "".join(
                    traceback.format_exception(type(exc), exc, exc.__traceback__)
                )
                self.after(0, lambda e=exc, d=err: self._on_error(e, d))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, stats: dict):
        self._set_progress(100, "导出成功。")
        self._set_busy(False)
        self.output_var.set(stats["output_path"])
        self._refresh_open_result_btn()
        lines = [
            "比对完成。",
            f"查询周期：{stats['period_start']} ~ {stats['period_end']}"
            f"{'（自动推断）' if stats.get('period_inferred') else ''}",
            f"账期：{stats['account_months']}月；"
            f"期初豁免截止 {stats['grace_invoice_until']}；"
            f"期末豁免起始 {stats['grace_order_from']}",
            f"订单工作表：{stats['order_sheet']}（{stats['order_rows']} 行）",
            f"发票工作表：{stats['invoice_sheet']}（{stats['invoice_rows']} 行）",
            f"结果行数：{stats['result_rows']}",
            f"需人工核对：{stats['manual_count']}",
        ]
        for k, v in sorted(stats["status_counts"].items()):
            lines.append(f"  · {k}: {v}")
        lines.append(f"已保存：{stats['output_path']}")
        for line in lines:
            self._append_log(line)
        if messagebox.askyesno(
            APP_TITLE,
            "比对结果已导出。\n\n"
            + "\n".join(lines[1:6] + [lines[-1]])
            + "\n\n是否立即打开比对结果？",
        ):
            self._on_open_result()

    def _on_permission_error(self, exc: BaseException, detail: str):
        self._set_busy(False)
        tip = self._explain_error(exc)
        self._append_log("导出失败（文件被占用）：")
        self._append_log(tip)
        self._append_log(detail)
        self.status_var.set("导出失败：目标文件被占用。")
        if messagebox.askyesno(
            APP_TITLE, tip + "\n\n是否现在另选保存路径并重试？"
        ):
            current = Path(self.output_var.get().strip() or DEFAULT_RESULT_NAME)
            alt = self._ask_alternate_output(current)
            if not alt:
                return
            self.output_var.set(str(alt))
            self._on_export()

    def _on_error(self, exc: BaseException, detail: str):
        self._set_busy(False)
        tip = self._explain_error(exc)
        self._append_log("导出失败：")
        self._append_log(tip)
        self._append_log(detail)
        self.status_var.set("导出失败。")
        messagebox.showerror(APP_TITLE, f"导出失败\n\n{tip}")


def main():
    app = ReconcileApp()
    app.mainloop()


if __name__ == "__main__":
    main()
