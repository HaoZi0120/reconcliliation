# -*- coding: utf-8 -*-
"""订单 / 发票比对工具 — 图形界面（加载 ui.html Fluent 界面）。"""

from __future__ import annotations

import json
import os
import shutil
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import webview

from reconcile import (
    DEFAULT_ACCOUNT_DAYS,
    peek_period_from_invoice_file,
    run_reconcile,
)

APP_TITLE = "订单发票比对工具"
DEFAULT_RESULT_NAME = "比对结果.xlsx"
APP_VERSION = "Pro3.0"

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


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ui_html_path() -> Path:
    p = _app_dir() / "ui.html"
    if not p.is_file():
        raise FileNotFoundError(f"未找到界面文件：{p}")
    return p


def _template_dir() -> Path:
    base = _app_dir()
    p = base / "templates"
    if p.is_dir():
        return p
    return Path(__file__).resolve().parent / "templates"


def _template_path(filename: str) -> Path:
    path = _template_dir() / filename
    if not path.is_file():
        raise FileNotFoundError(f"未找到模板文件：{path}")
    return path


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


def _explain_error(exc: BaseException) -> str:
    msg = str(exc)
    if isinstance(exc, PermissionError) or (
        isinstance(exc, OSError) and getattr(exc, "errno", None) in (13, 32)
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


def _can_write_path(path: Path) -> bool:
    try:
        with open(path, "a+b"):
            return True
    except OSError:
        return False


class Api:
    """供 ui.html 通过 pywebview.js_api 调用。"""

    def __init__(self):
        self._window = None
        self._busy = False

    def set_window(self, window):
        self._window = window

    def _js(self, script: str):
        if self._window is None:
            return
        try:
            self._window.evaluate_js(script)
        except Exception:
            pass

    def _log(self, level: str, message: str):
        payload = json.dumps({"level": level, "message": message}, ensure_ascii=False)
        self._js(f"window.__hostLog && window.__hostLog({payload})")

    def _progress(self, pct: int, message: str = ""):
        payload = json.dumps(
            {"pct": int(pct), "message": message}, ensure_ascii=False
        )
        self._js(f"window.__hostProgress && window.__hostProgress({payload})")

    def get_defaults(self):
        return {
            "title": APP_TITLE,
            "version": APP_VERSION,
            "accountDays": DEFAULT_ACCOUNT_DAYS,
            "resultName": DEFAULT_RESULT_NAME,
        }

    def get_format_help(self):
        return FORMAT_HELP_TEXT

    def browse_order(self):
        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=False,
            file_types=("Excel 文件 (*.xlsx;*.xlsm)", "所有文件 (*.*)"),
        )
        if not result:
            return {"ok": False}
        path = str(result[0])
        self._log("INFO", f"已选择订单文件：{path}")
        out = self._default_output_if_empty(path)
        return {"ok": True, "path": path, "output": out}

    def browse_invoice(self, account_days: int | None = None):
        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=False,
            file_types=("Excel 文件 (*.xlsx;*.xlsm)", "所有文件 (*.*)"),
        )
        if not result:
            return {"ok": False}
        path = str(result[0])
        self._log("INFO", f"已选择发票文件：{path}")
        out = self._default_output_if_empty(path)
        days = account_days if account_days is not None else DEFAULT_ACCOUNT_DAYS
        try:
            days = int(days)
            if days < 0:
                days = DEFAULT_ACCOUNT_DAYS
        except (TypeError, ValueError):
            days = DEFAULT_ACCOUNT_DAYS

        period = {}
        try:
            cfg = peek_period_from_invoice_file(Path(path), days)
            period = {
                "start": str(cfg.start),
                "end": str(cfg.end),
                "graceInvoiceUntil": str(cfg.grace_invoice_until),
                "graceOrderFrom": str(cfg.grace_order_from),
                "days": days,
            }
            self._log(
                "SUCCESS",
                f"自动推断查询周期：{cfg.start} ~ {cfg.end}（账期 {days} 天）；"
                f"期初豁免截止 {cfg.grace_invoice_until}，期末豁免起始 {cfg.grace_order_from}",
            )
        except Exception as exc:
            self._log("WARN", f"自动推断查询周期失败：{exc}")
            period = {"error": str(exc)}

        return {"ok": True, "path": path, "output": out, "period": period}

    def browse_output(self, current: str = ""):
        initial = (current or "").strip()
        directory = ""
        save_filename = DEFAULT_RESULT_NAME
        if initial:
            p = Path(initial)
            if p.parent.exists():
                directory = str(p.parent)
            save_filename = p.name or DEFAULT_RESULT_NAME
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory=directory,
            save_filename=save_filename,
            file_types=("Excel 工作簿 (*.xlsx)",),
        )
        if not result:
            return {"ok": False}
        path = str(result if isinstance(result, str) else result[0])
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._log("INFO", f"结果保存路径：{path}")
        return {"ok": True, "path": path}

    def _default_output_if_empty(self, base_path: str) -> str | None:
        # 由前端决定是否覆盖；这里仅提供建议路径
        return str(Path(base_path).with_name(DEFAULT_RESULT_NAME))

    def download_template(self, kind: str):
        name_map = {
            "order": "订单数据模板.xlsx",
            "invoice": "发票数据模板.xlsx",
            "订单数据模板.xlsx": "订单数据模板.xlsx",
            "发票数据模板.xlsx": "发票数据模板.xlsx",
        }
        filename = name_map.get(kind, kind)
        try:
            src = _template_path(filename)
        except FileNotFoundError as exc:
            self._log("ERROR", str(exc))
            return {"ok": False, "error": str(exc)}

        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=filename,
            file_types=("Excel 工作簿 (*.xlsx)",),
        )
        if not result:
            return {"ok": False, "cancelled": True}
        path = str(result if isinstance(result, str) else result[0])
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            shutil.copy2(src, path)
        except OSError as exc:
            self._log("ERROR", f"保存模板失败：{exc}")
            return {"ok": False, "error": str(exc)}
        self._log("SUCCESS", f"模板已保存：{path}")
        return {"ok": True, "path": path}

    def open_result(self, path: str):
        path = (path or "").strip()
        if not path:
            return {"ok": False, "error": "请先指定结果保存路径，并完成一次导出。"}
        p = Path(path)
        if not p.is_file():
            return {
                "ok": False,
                "error": f"比对结果文件尚不存在。\n当前路径：{p}",
            }
        try:
            os.startfile(str(p))  # noqa: S606
            self._log("INFO", f"正在打开：{p}")
            return {"ok": True}
        except OSError as exc:
            tip = (
                "无法打开比对结果文件。\n\n"
                "可能原因：\n"
                "· 未安装 Excel / WPS，或未关联 .xlsx 默认打开方式\n"
                "· 文件路径无效或已被移动/删除\n"
                "· 系统权限不足\n\n"
                f"路径：{p}\n详情：{exc}"
            )
            self._log("ERROR", tip)
            return {"ok": False, "error": tip}

    def result_exists(self, path: str):
        path = (path or "").strip()
        return bool(path) and Path(path).is_file()

    def start_export(
        self,
        order: str,
        invoice: str,
        output: str,
        period_start: str,
        period_end: str,
        account_days,
        overwrite_confirmed: bool = False,
    ):
        if self._busy:
            return {"ok": False, "error": "正在比对中，请稍候。"}

        order = (order or "").strip()
        invoice = (invoice or "").strip()
        output = (output or "").strip()

        if not order:
            return {"ok": False, "error": "请先选择订单数据文件。"}
        if not invoice:
            return {"ok": False, "error": "请先选择发票数据文件。"}
        if not output:
            return {"ok": False, "error": "请先指定结果保存路径。"}

        try:
            days = int(str(account_days).strip())
            if days < 0:
                raise ValueError("账期不能为负")
        except ValueError as exc:
            return {"ok": False, "error": f"账期天数无效：{exc}"}

        try:
            start = _parse_date(period_start, "周期起始日")
            end = _parse_date(period_end, "周期结束日")
        except ValueError as exc:
            return {"ok": False, "error": str(exc)}

        if start is None or end is None:
            try:
                cfg = peek_period_from_invoice_file(Path(invoice), days)
                start, end = cfg.start, cfg.end
                self._log("INFO", f"导出前补全推断周期：{cfg.start} ~ {cfg.end}")
                self._js(
                    "window.__hostSetPeriod && window.__hostSetPeriod("
                    + json.dumps({"start": str(cfg.start), "end": str(cfg.end)}, ensure_ascii=False)
                    + ")"
                )
            except Exception as exc:
                return {
                    "ok": False,
                    "error": f"查询周期为空且无法推断，请填写周期起止日。\n{exc}",
                }

        out_path = Path(output)
        if out_path.exists() and not overwrite_confirmed:
            if not _can_write_path(out_path):
                return {
                    "ok": False,
                    "needAlternate": True,
                    "error": (
                        "无法覆盖写入结果文件。\n\n"
                        "可能原因：\n"
                        "1. 该文件正被 Excel / WPS 打开（最常见）\n"
                        "2. 文件为只读，或对文件夹无写入权限\n"
                        "3. 杀毒/网盘同步正在占用文件\n\n"
                        f"路径：{out_path}"
                    ),
                }
            return {
                "ok": False,
                "needConfirmOverwrite": True,
                "path": str(out_path),
            }

        if out_path.exists() and not _can_write_path(out_path):
            return {
                "ok": False,
                "needAlternate": True,
                "error": (
                    "无法覆盖写入结果文件（可能正被 Excel/WPS 打开）。\n"
                    f"路径：{out_path}"
                ),
            }

        self._busy = True
        self._progress(0, "正在比对，请稍候…")
        self._log("INFO", "=" * 48)
        self._log("INFO", f"订单：{order}")
        self._log("INFO", f"发票：{invoice}")
        self._log("INFO", f"输出：{output}")
        self._log("INFO", f"查询周期：{start} ~ {end}；账期：{days}天")

        def worker():
            try:
                def cb(pct, message):
                    self._progress(pct, message)
                    if message:
                        self._log("INFO", message)

                stats = run_reconcile(
                    order,
                    invoice,
                    output,
                    save_preprocess_files=False,
                    period_start=start,
                    period_end=end,
                    account_days=days,
                    progress_callback=cb,
                )
                self._on_success(stats)
            except Exception as exc:
                err = "".join(
                    traceback.format_exception(type(exc), exc, exc.__traceback__)
                )
                self._on_failure(exc, err)

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "started": True}

    def pick_alternate_output(self, current: str = ""):
        current_path = Path((current or "").strip() or DEFAULT_RESULT_NAME)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested = f"{current_path.stem}_{stamp}{current_path.suffix or '.xlsx'}"
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory=str(current_path.parent) if current_path.parent.exists() else "",
            save_filename=suggested,
            file_types=("Excel 工作簿 (*.xlsx)",),
        )
        if not result:
            return {"ok": False}
        path = str(result if isinstance(result, str) else result[0])
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return {"ok": True, "path": path}

    def _on_success(self, stats: dict):
        self._busy = False
        self._progress(100, "比对成功完成！")
        lines = [
            "比对完成。",
            f"查询周期：{stats['period_start']} ~ {stats['period_end']}"
            f"{'（自动推断）' if stats.get('period_inferred') else ''}",
            f"账期：{stats['account_days']}天；"
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
            self._log("SUCCESS" if "已保存" in line or line == "比对完成。" else "INFO", line)

        payload = json.dumps(
            {
                "output": stats["output_path"],
                "summary": "\n".join(lines[1:6] + [lines[-1]]),
            },
            ensure_ascii=False,
        )
        self._js(f"window.__hostExportDone && window.__hostExportDone({payload})")

    def _on_failure(self, exc: BaseException, detail: str):
        self._busy = False
        tip = _explain_error(exc)
        is_perm = isinstance(exc, PermissionError) or (
            isinstance(exc, OSError) and getattr(exc, "errno", None) in (13, 32)
        )
        self._log("ERROR", "导出失败" + ("（文件被占用）" if is_perm else ""))
        self._log("ERROR", tip)
        self._log("ERROR", detail)
        self._progress(0, "导出失败。")
        payload = json.dumps(
            {"error": tip, "permission": is_perm},
            ensure_ascii=False,
        )
        self._js(f"window.__hostExportFail && window.__hostExportFail({payload})")

    def exit_app(self):
        self._log("WARN", "正在关闭订单发票比对工具...")
        if self._window is not None:
            self._window.destroy()


def main():
    api = Api()
    html = _ui_html_path()
    window = webview.create_window(
        APP_TITLE,
        url=html.as_uri(),
        js_api=api,
        width=1100,
        height=860,
        min_size=(900, 700),
        background_color="#F3F4F6",
    )
    api.set_window(window)
    webview.start(debug=False)


if __name__ == "__main__":
    main()
