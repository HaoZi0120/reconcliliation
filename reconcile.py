# -*- coding: utf-8 -*-
"""订单 vs 发票比对脚本。

规则锁定：
1. 日期：交货日严格早于开票日；或开票早于交货≤3天且开票日≥接单日
2. 型号匹配键：去全部空白后 casefold（大小写、空格不敏感）；展示仍用原文
3. 已核销（原完全匹配）：匹配键同 + 日期合法 + 数量相等 + |用户价格-含税单价|<=0.02
4. 已配对但价格不合格细分：实际售价高于申请价格 / 实际售价低于申请价格 / 价格缺失（均需人工）
5. 业务约束：含税单价不得高于用户价格；超申请用 PRICE_OVER，低于申请用 PRICE_UNDER
6. 多候选 FIFO 成功：算已核销，写 MULTI_CAND
7. 原文仅大小写/空格差异：写 CASE_DIFF / SPACE_DIFF，仍可已核销
8. 订单预处理：用户后增加「E率」=(价格/1.13/E价)-1
9. 发票数量为负或状态为「已红冲-全额」、订单号以S开头：不参与比对但保留数据；「已红冲-部分」仍参与比对
10. 发票列绿底、订单列蓝底；核实高时发票/订单红字、比对结果列浅橙底；数据异常仅比对结果列浅红
"""

from __future__ import annotations

import calendar
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "datas"
ORDER_FILE = DATA_DIR / "订单数据3.xlsx"
INVOICE_FILE = DATA_DIR / "发票数据3.xlsx"
OUTPUT_FILE = DATA_DIR / "比对结果.xlsx"

ORDER_REQUIRED_HEADERS = {"交货日期", "型号", "数量", "用户价格"}
INVOICE_REQUIRED_HEADERS = {"开票日期", "规格型号", "数量", "单价"}

# 仅「已红冲-全额」排除；「已红冲-部分」仍参与比对
EXCLUDED_INVOICE_STATUS_FULL_FLUSH = "已红冲-全额"
PRICE_TOLERANCE = 0.02
# 开票可早于交货的最大天数（须同时满足开票日≥接单日）
INVOICE_EARLY_DAYS = 3
# 已配对但价格相关的细分状态（均 need_manual=True）
STATUS_MATCHED = "已核销"  # 原「完全匹配」
STATUS_PRICE_OVER = "实际售价高于申请价格"
STATUS_PRICE_UNDER = "实际售价低于申请价格"
STATUS_PRICE_MISSING = "价格缺失"
NEED_EXPLAIN_YES = "需解释"

# 未配 / 账期相关状态
STATUS_GRACE_ORDER = "开票在下一年"  # 原「账期内待开票」
STATUS_GRACE_INVOICE = "订单在上一年"  # 原「账期内无订单」
STATUS_ORDER_NO_INV = "订单未开票"
STATUS_INV_NO_ORDER = "发票无订单"
STATUS_SAME_DATE = "同型号开票早于发货"
STATUS_SAME_QTY = "同型号数量无法对齐"
STATUS_SAME_OCCUPIED = "同型号行已被其他票/单占用"
# 匹配状态需红字加粗
STATUS_RED_BOLD = frozenset(
    {STATUS_PRICE_OVER, STATUS_ORDER_NO_INV, STATUS_INV_NO_ORDER}
)

BIAS_ORDER = "订单未开票"
BIAS_INV = "发票无订单"
BIAS_OTHER = "其他"

# 核实优先级（筛选用，找数据提供方沟通）
VERIFY_HIGH = "高"
VERIFY_MID = "中"
VERIFY_NONE = ""  # 不沟通

VERIFY_PRIORITY_HIGH_STATUSES = {
    STATUS_PRICE_OVER,
    STATUS_SAME_DATE,
    STATUS_SAME_OCCUPIED,
    STATUS_ORDER_NO_INV,
    STATUS_INV_NO_ORDER,
    STATUS_SAME_QTY,
}
VERIFY_PRIORITY_MID_STATUSES = {
    "数据异常",
    STATUS_PRICE_MISSING,
}

DEFAULT_ACCOUNT_DAYS = 30

# 监察方式
MODE_DAILY = "daily"
MODE_STRATEGIC = "strategic"
MODE_LABELS = {
    MODE_DAILY: "日常监察",
    MODE_STRATEGIC: "战略品监察",
}


def verify_priority_for(status: str) -> str:
    """根据匹配状态返回核实优先级：高 / 中 / 空。"""
    if status in VERIFY_PRIORITY_HIGH_STATUSES:
        return VERIFY_HIGH
    if status in VERIFY_PRIORITY_MID_STATUSES:
        return VERIFY_MID
    return VERIFY_NONE


def need_explain_for(verify_priority: str) -> str:
    """核实优先级为「高」时填「需解释」，否则为空。"""
    return NEED_EXPLAIN_YES if verify_priority == VERIFY_HIGH else ""


@dataclass
class PeriodConfig:
    """发票查询周期 + 账期推算的豁免边界。"""

    start: date
    end: date
    account_days: int
    grace_invoice_until: date  # 开票日<=此日且无订单 → 订单在上一年
    grace_order_from: date  # 交货日>=此日且无发票 → 开票在下一年


def to_date_only(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    dt = to_datetime(v)
    return dt.date() if dt else None


def build_period_config(start, end, account_days: int) -> PeriodConfig:
    start_d = to_date_only(start)
    end_d = to_date_only(end)
    if start_d is None or end_d is None:
        raise ValueError("发票查询周期起止日无效")
    if end_d < start_d:
        raise ValueError("发票查询周期结束日不能早于起始日")
    days = int(account_days)
    if days < 0:
        raise ValueError("账期天数不能为负")
    # 期初豁免截止日 = 起始日 + N 天 - 1 天；期末豁免起始日 = 结束日 - N 天 + 1 天
    grace_invoice_until = start_d + timedelta(days=days) - timedelta(days=1)
    grace_order_from = end_d - timedelta(days=days) + timedelta(days=1)
    return PeriodConfig(
        start=start_d,
        end=end_d,
        account_days=days,
        grace_invoice_until=grace_invoice_until,
        grace_order_from=grace_order_from,
    )


def month_first_day(d: date) -> date:
    return date(d.year, d.month, 1)


def month_last_day(d: date) -> date:
    return date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def infer_period_from_invoices(invoices, account_days: int) -> PeriodConfig:
    """按发票开票日 min/max 所在月：起始=月初第一天，结束=月末最后一天。"""
    dates = [to_date_only(i.get("_inv_dt") or i.get("开票日期")) for i in invoices]
    dates = [d for d in dates if d is not None]
    if not dates:
        raise ValueError("无法从发票数据推断查询周期：无有效开票日期")
    start = month_first_day(min(dates))
    end = month_last_day(max(dates))
    return build_period_config(start, end, account_days)


def bias_for(status: str, side: str | None) -> str:
    """side: order / invoice / None。配对成功等非未配场景返回空串。"""
    if status in (STATUS_GRACE_ORDER, STATUS_ORDER_NO_INV):
        return BIAS_ORDER
    if status in (STATUS_GRACE_INVOICE, STATUS_INV_NO_ORDER):
        return BIAS_INV
    if status == STATUS_SAME_DATE:
        return BIAS_ORDER if side == "order" else BIAS_INV
    if status == STATUS_SAME_QTY:
        return BIAS_ORDER if side == "order" else BIAS_INV
    if status == STATUS_SAME_OCCUPIED:
        return BIAS_OTHER
    if status == "数据异常":
        return BIAS_OTHER
    return ""
ORDER_FILL = PatternFill("solid", fgColor="DDEBF7")  # 浅蓝（预处理等）
MANUAL_FILL = PatternFill("solid", fgColor="FFF2CC")  # 浅黄（汇总提示等）
ERROR_FILL = PatternFill("solid", fgColor="FCE4EC")  # 浅红 数据异常行
HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
# 匹配明细视觉规范
FILL_ZONE_A_HEADER = PatternFill("solid", fgColor="5B2C6F")  # 匹配结果表头深紫
FILL_ZONE_B_INV_HEADER = PatternFill("solid", fgColor="70AD47")
FILL_ZONE_B_ORD_HEADER = PatternFill("solid", fgColor="5B9BD5")
FILL_ZONE_C_HEADER = PatternFill("solid", fgColor="A9D08E")
FILL_ZONE_D_HEADER = PatternFill("solid", fgColor="9DC3E6")
FILL_ZONE_A_DATA = PatternFill("solid", fgColor="F5EEF8")  # 匹配结果默认浅紫
FILL_META_HIGH = PatternFill("solid", fgColor="FCE4D6")  # 核实=高：仅比对结果列浅橙底
FILL_INV_DATA = PatternFill("solid", fgColor="E2EFDA")  # 发票列统一浅绿
FILL_ORD_DATA = PatternFill("solid", fgColor="DDEBF7")  # 订单列统一浅蓝
# 兼容旧名（隐藏列与可见列数据底色已统一）
FILL_ZONE_B_INV_DATA = FILL_INV_DATA
FILL_ZONE_B_ORD_DATA = FILL_ORD_DATA
FILL_ZONE_C_DATA = FILL_INV_DATA
FILL_ZONE_D_DATA = FILL_ORD_DATA
# 匹配明细统一等线字体
FONT_NAME = "等线"


def _font(bold=False, color=None, size=None):
    kw = {"name": FONT_NAME}
    if bold:
        kw["bold"] = True
    if color is not None:
        kw["color"] = color
    if size is not None:
        kw["size"] = size
    return Font(**kw)


FONT_DEFAULT = _font()
FONT_HEADER_WHITE = _font(bold=True, color="FFFFFF")
FONT_HEADER_DARK = _font(bold=True, color="1F4E79")
FONT_VERIFY_HIGH = _font(bold=True, color="C00000")  # 核实=高：红字
FONT_VERIFY_MID = _font(bold=True, color="833C0C")
FONT_HIGH_INV_ORD = _font(color="C00000")  # 核实=高：发票/订单列红字
FONT_EXPLAIN = _font(bold=True, color="C65911")  # 「需解释」橘色加粗
FONT_STATUS_OVER = _font(bold=True, color="C00000")
FONT_BIAS_OTHER = _font(color="7030A0")
FONT_REASON = _font(color="595959", size=9)
FILL_MODEL_SUMMARY = PatternFill("solid", fgColor="D9D9D9")  # 型号汇总行浅灰底
FONT_MODEL_SUMMARY = _font(bold=True, color="000000")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def to_float(v):
    """文本/数字转 float；支持去千分位逗号与首尾空白。"""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("，", "").replace(" ", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_int_qty(v):
    """数量转整数（先解析为数字再四舍五入）。"""
    f = to_float(v)
    if f is None:
        return None
    return int(round(f))


def to_datetime(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip().replace("/", "-")
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S.%f"):
            try:
                return datetime.strptime(s[:26], fmt)
            except ValueError:
                continue
    return None


def norm_model_key(v):
    """型号匹配键：去掉全部空白后 casefold，大小写与空格不敏感。"""
    if v is None:
        return None
    s = re.sub(r"\s+", "", str(v))
    s = s.casefold()
    return s or None


def append_model_text_diffs(reasons, order_raw, inv_raw):
    """匹配键相同但原文不一致时记录 CASE_DIFF / SPACE_DIFF（不强制人工）。"""
    o = "" if order_raw is None else str(order_raw)
    i = "" if inv_raw is None else str(inv_raw)
    if o == i:
        return
    o_c = re.sub(r"\s+", "", o)
    i_c = re.sub(r"\s+", "", i)
    # 压缩后仍不同 → 仅大小写差异（匹配键已相同）
    if o_c != i_c:
        reasons.append(f"CASE_DIFF:型号大小写不一致：订单{o} / 发票{i}")
    # 任一侧含空白，或去空白后相同但原文不同 → 空格差异
    if o_c != o or i_c != i or o_c == i_c:
        # o!=i 已在上方保证；o_c==i_c 表示仅空格写法不同
        reasons.append(f"SPACE_DIFF:型号含空格差异：订单{o} / 发票{i}")


def _header_set(row):
    return {str(h).strip() for h in row if h is not None and str(h).strip()}


def find_sheet_name(path: Path, required_headers: set[str]) -> str:
    """按表头自动定位工作表；找不到则抛出明确错误。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if row and required_headers.issubset(_header_set(row)):
                return name
        raise ValueError(
            f"在文件中未找到含必要列 {sorted(required_headers)} 的工作表：{path.name}"
        )
    finally:
        wb.close()


def read_sheet_rows(path: Path, sheet_name=None, required_headers=None):
    path = Path(path)
    if sheet_name is None and required_headers:
        sheet_name = find_sheet_name(path, required_headers)
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    used_sheet = ws.title
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], [], used_sheet
    headers = [h if h is not None else f"空列{i+1}" for i, h in enumerate(rows[0])]
    while headers and str(headers[-1]).startswith("空列"):
        headers.pop()
    data = []
    for r in rows[1:]:
        if r is None or all(c is None or str(c).strip() == "" for c in r):
            continue
        item = {}
        for i, h in enumerate(headers):
            item[h] = r[i] if i < len(r) else None
        data.append(item)
    return headers, data, used_sheet


def order_exclude_reasons(row) -> list[str]:
    """配对前排除规则（数据仍保留，仅不参与匹配）。

    订单号去掉前导空白后若以 S/s 开头，则不参与比对。
    """
    reasons = []
    ono = row.get("订单号")
    if ono is None or str(ono).strip() == "":
        return reasons
    s = str(ono).lstrip()
    if s and s[0].casefold() == "s":
        reasons.append(
            f"EXCL_ORDER_NO_S:订单号以S开头（实际为「{str(ono).strip()}」），不参与比对"
        )
    return reasons


def preprocess_orders(rows):
    """返回带价格合计的订单列表（已排序）；标记不参与比对的行。"""
    out = []
    for idx, r in enumerate(rows):
        row = dict(r)
        qty = to_float(row.get("数量"))
        price = to_float(row.get("用户价格"))
        if qty is not None and price is not None:
            row["价格合计"] = round(qty * price, 6)
        else:
            row["价格合计"] = None
        # E率 = round(((价格/1.13/E价)-1)*100)，整数百分比，不留小数
        list_price = to_float(row.get("价格"))
        e_price = to_float(row.get("E价"))
        if list_price is not None and e_price is not None and e_price != 0:
            row["E率"] = int(round(((list_price / 1.13 / e_price) - 1) * 100))
        else:
            row["E率"] = None
        # 交货日期 / 接单日期：只保留年月日
        ship_d = to_date_only(row.get("交货日期"))
        row["交货日期"] = ship_d
        recv_d = to_date_only(row.get("接单日期"))
        if "接单日期" in row or recv_d is not None:
            row["接单日期"] = recv_d
        row["_src_idx"] = idx
        row["_model_key"] = norm_model_key(row.get("型号"))
        row["_qty"] = qty
        row["_user_price"] = price
        row["_ship_dt"] = (
            datetime(ship_d.year, ship_d.month, ship_d.day) if ship_d else None
        )
        row["_recv_dt"] = (
            datetime(recv_d.year, recv_d.month, recv_d.day) if recv_d else None
        )
        row["_exclude_reasons"] = order_exclude_reasons(row)
        out.append(row)
    out.sort(
        key=lambda x: (
            x["_model_key"] or "",
            x["_ship_dt"] or datetime.max,
            x["_src_idx"],
        )
    )
    return out


def invoice_exclude_reasons(row) -> list[str]:
    """配对前排除规则（数据仍保留，仅不参与匹配）。"""
    reasons = []
    qty = row.get("_qty")
    if qty is not None and qty < 0:
        reasons.append(f"EXCL_NEG_QTY:发票数量为负数（{qty:g}），不参与比对")
    status = row.get("发票状态")
    if status is not None and EXCLUDED_INVOICE_STATUS_FULL_FLUSH in str(status):
        reasons.append(
            f"EXCL_RED_FLUSH:发票状态为「{EXCLUDED_INVOICE_STATUS_FULL_FLUSH}」"
            f"（实际为「{str(status).strip()}」），不参与比对；"
            "「已红冲-部分」仍参与比对"
        )
    return reasons


def preprocess_invoices(rows):
    """返回带含税单价的发票列表（已排序）；标记不参与比对的行。

    数量、单价若为文本则转为数字：数量保留整数，单价保留 2 位小数。
    """
    out = []
    for idx, r in enumerate(rows):
        row = dict(r)
        qty = to_int_qty(row.get("数量"))
        row["数量"] = qty
        unit = to_float(row.get("单价"))
        if unit is not None:
            unit = round(unit, 2)
            row["单价"] = unit
            row["含税单价"] = round(unit * 1.13, 10)
        else:
            row["单价"] = None
            row["含税单价"] = None
        # 开票日期：只保留年月日
        inv_d = to_date_only(row.get("开票日期"))
        row["开票日期"] = inv_d
        row["_src_idx"] = idx
        row["_model_key"] = norm_model_key(row.get("规格型号"))
        row["_qty"] = float(qty) if qty is not None else None
        row["_tax_price"] = row["含税单价"]
        row["_inv_dt"] = (
            datetime(inv_d.year, inv_d.month, inv_d.day) if inv_d else None
        )
        row["_exclude_reasons"] = invoice_exclude_reasons(row)
        out.append(row)
    out.sort(
        key=lambda x: (
            x["_model_key"] or "",
            x["_inv_dt"] or datetime.max,
            x["_src_idx"],
        )
    )
    return out


def split_for_match(rows):
    """拆成参与比对 / 不参与比对（后者仍输出到结果）。"""
    active = [r for r in rows if not r.get("_exclude_reasons")]
    excluded = [r for r in rows if r.get("_exclude_reasons")]
    return active, excluded


def split_invoices_for_match(invoices):
    return split_for_match(invoices)


def split_orders_for_match(orders):
    return split_for_match(orders)


def model_qty_totals(orders, invoices):
    so = defaultdict(float)
    si = defaultdict(float)
    for o in orders:
        if o.get("_exclude_reasons"):
            continue
        if o["_model_key"] and o["_qty"] is not None:
            so[o["_model_key"]] += o["_qty"]
    for i in invoices:
        if i.get("_exclude_reasons"):
            continue
        if i["_model_key"] and i["_qty"] is not None:
            si[i["_model_key"]] += i["_qty"]
    return so, si


def date_pair_ok(ship_dt, inv_dt, recv_dt=None) -> bool:
    """订单-发票日期是否合法。

    - 常规：交货日严格早于开票日
    - 软窗口：开票早于交货 1～INVOICE_EARLY_DAYS 天，且接单日存在、开票日≥接单日
    - 同日开票=交货：非法（保持现网）
    - 软窗口但接单日缺失：非法
    """
    ship_d = to_date_only(ship_dt)
    inv_d = to_date_only(inv_dt)
    recv_d = to_date_only(recv_dt)
    if ship_d is None or inv_d is None:
        return False
    if ship_d < inv_d:
        return True
    early = (ship_d - inv_d).days
    if 1 <= early <= INVOICE_EARLY_DAYS and recv_d is not None and inv_d >= recv_d:
        return True
    return False


def date_pair_used_soft_window(ship_dt, inv_dt, recv_dt=None) -> bool:
    """是否靠「提前≤3天 + 接单日」软窗口合法（非常规 ship<inv）。"""
    ship_d = to_date_only(ship_dt)
    inv_d = to_date_only(inv_dt)
    recv_d = to_date_only(recv_dt)
    if ship_d is None or inv_d is None or recv_d is None:
        return False
    if ship_d < inv_d:
        return False
    early = (ship_d - inv_d).days
    return 1 <= early <= INVOICE_EARLY_DAYS and inv_d >= recv_d


def build_model_index(rows):
    """按 _model_key 建索引：key -> [行下标]。"""
    idx: dict[str, list[int]] = defaultdict(list)
    for i, row in enumerate(rows):
        k = row.get("_model_key")
        if k:
            idx[k].append(i)
    return idx


def find_candidates(order, invoices, used_inv, candidate_js=None):
    """数量相等且日期合法的候选发票下标列表。

    candidate_js: 可选，同型号发票下标列表（有索引时传入，避免全表扫描）。
    """
    if candidate_js is None:
        candidate_js = range(len(invoices))
    cands = []
    for j in candidate_js:
        if j in used_inv:
            continue
        inv = invoices[j]
        if inv["_model_key"] != order["_model_key"]:
            continue
        if order["_ship_dt"] is None or inv["_inv_dt"] is None:
            continue
        if not date_pair_ok(order["_ship_dt"], inv["_inv_dt"], order.get("_recv_dt")):
            continue
        if order["_qty"] is None or inv["_qty"] is None:
            continue
        if abs(order["_qty"] - inv["_qty"]) > 1e-9:
            continue
        cands.append(j)
    return cands


def pick_best_candidate(order, invoices, cand_idxs):
    """在数量已相等的候选中：价格接近优先，再 FIFO（开票日早优先）。"""
    scored = []
    for j in cand_idxs:
        inv = invoices[j]
        price_ok = False
        price_diff = None
        if order["_user_price"] is not None and inv["_tax_price"] is not None:
            price_diff = abs(order["_user_price"] - inv["_tax_price"])
            price_ok = price_diff <= PRICE_TOLERANCE
        # 排序：价格合格优先；差价小优先；开票日期早优先；原序
        scored.append(
            (
                0 if price_ok else 1,
                price_diff if price_diff is not None else 1e18,
                inv["_inv_dt"] or datetime.max,
                j,
            )
        )
    scored.sort()
    best_j = scored[0][3]
    return best_j, len(cand_idxs) > 1


def _dedupe_reasons(reasons):
    seen = set()
    uniq = []
    for r in reasons:
        if r not in seen:
            seen.add(r)
            uniq.append(r)
    return uniq


def classify_unmatched_order(
    order,
    invoices,
    so,
    si,
    used_inv,
    period: PeriodConfig | None,
    same_model_js=None,
):
    """返回 (status, need_manual, reasons, bias)。

    same_model_js: 同型号发票下标（来自型号索引）；缺省则全表扫描（兼容旧调用）。
    """
    reasons = []
    model = order["_model_key"]
    if not model:
        return "数据异常", True, ["BAD_FIELD:关键字段缺失或无法解析：型号"], BIAS_OTHER
    if order["_ship_dt"] is None:
        reasons.append("BAD_FIELD:关键字段缺失或无法解析：交货日期")
    if order["_qty"] is None:
        reasons.append("BAD_FIELD:关键字段缺失或无法解析：数量")
    if any(r.startswith("BAD_FIELD") for r in reasons):
        return "数据异常", True, _dedupe_reasons(reasons), BIAS_OTHER

    ship_d = to_date_only(order["_ship_dt"])
    # 账期豁免优先
    if period and ship_d and ship_d >= period.grace_order_from:
        reasons.append(
            f"GRACE_ORDER:交货日{ship_d}≥期末豁免起始日{period.grace_order_from}，"
            f"开票在下一年（查询周期{period.start}~{period.end}，账期{period.account_days}天）"
        )
        return STATUS_GRACE_ORDER, False, reasons, BIAS_ORDER

    if same_model_js is None:
        same_model_js = [j for j, i in enumerate(invoices) if i["_model_key"] == model]
    same_model_invs = [invoices[j] for j in same_model_js]
    if not same_model_invs:
        reasons.append("ONLY_ORDER:该型号仅出现在订单中")
        status_flag = order.get("订单状态")
        if status_flag and str(status_flag).strip() == "已开票":
            reasons.append("STATUS_MISMATCH:订单状态为「已开票」但未找到匹配发票")
        return STATUS_ORDER_NO_INV, False, _dedupe_reasons(reasons), BIAS_ORDER

    # 同型号细分：占用 > 日期 > 数量
    legal_all = find_candidates(order, invoices, used_inv=set(), candidate_js=same_model_js)
    legal_free = find_candidates(order, invoices, used_inv, candidate_js=same_model_js)
    if legal_all and not legal_free:
        reasons.append(
            "OCCUPIED:存在数量相等且日期合法的同型号发票，但已被其他订单行占用"
        )
        reasons.append("偏向说明:表面像订单未开票，根因更偏配对竞争(其他)")
        return STATUS_SAME_OCCUPIED, True, _dedupe_reasons(reasons), BIAS_OTHER

    equal_qty = [
        i
        for i in same_model_invs
        if order["_qty"] is not None
        and i["_qty"] is not None
        and abs(order["_qty"] - i["_qty"]) < 1e-9
    ]
    if equal_qty and not legal_all:
        reasons.append(
            "DATE_ORDER:同型号存在等量发票，但不满足日期规则"
            f"（交货早于开票，或开票早于交货≤{INVOICE_EARLY_DAYS}天且开票≥接单）"
        )
        for inv in equal_qty:
            if order["_ship_dt"] and inv["_inv_dt"]:
                recv = order.get("_recv_dt")
                recv_s = f"{recv:%Y-%m-%d}" if recv else "空"
                reasons.append(
                    f"DATE_ORDER:交货{order['_ship_dt']:%Y-%m-%d} / "
                    f"开票{inv['_inv_dt']:%Y-%m-%d} / 接单{recv_s}"
                )
                break
        reasons.append("偏向说明:订单侧未配上，更偏向订单未开票")
        return STATUS_SAME_DATE, True, _dedupe_reasons(reasons), BIAS_ORDER

    o_total = so.get(model, 0)
    i_total = si.get(model, 0)
    reasons.append(
        f"QTY_MISALIGN:同型号无等量可配发票行（本行数量{order['_qty']}；"
        f"发票数量{[x['_qty'] for x in same_model_invs]}）"
    )
    if abs(o_total - i_total) > 1e-9:
        reasons.append(
            f"MODEL_QTY_GAP:该型号订单数量合计与发票数量合计不等"
            f"（订单合计{o_total:g} / 发票合计{i_total:g}）"
        )
        reasons.append(
            "QTY_POOL_HINT:型号级数量可冲减（合计接近），但行级一对一失败，建议人工拆分/合并核对"
        )
    status_flag = order.get("订单状态")
    if status_flag and str(status_flag).strip() == "已开票":
        reasons.append("STATUS_MISMATCH:订单状态为「已开票」但未找到匹配发票")
    reasons.append("偏向说明:订单侧未配上，更偏向订单未开票")
    return STATUS_SAME_QTY, True, _dedupe_reasons(reasons), BIAS_ORDER


def classify_unmatched_invoice(
    inv,
    orders,
    so,
    si,
    used_ord,
    period: PeriodConfig | None,
    same_model_ois=None,
):
    """返回 (status, need_manual, reasons, bias)。used_ord 为已占用订单下标集合。

    same_model_ois: 同型号订单下标（来自型号索引）；缺省则全表扫描。
    """
    reasons = []
    model = inv["_model_key"]
    if not model:
        return "数据异常", True, ["BAD_FIELD:关键字段缺失或无法解析：规格型号"], BIAS_OTHER
    if inv["_inv_dt"] is None:
        reasons.append("BAD_FIELD:关键字段缺失或无法解析：开票日期")
    if inv["_qty"] is None:
        reasons.append("BAD_FIELD:关键字段缺失或无法解析：数量")
    if any(r.startswith("BAD_FIELD") for r in reasons):
        return "数据异常", True, _dedupe_reasons(reasons), BIAS_OTHER

    inv_d = to_date_only(inv["_inv_dt"])
    if period and inv_d and inv_d <= period.grace_invoice_until:
        reasons.append(
            f"GRACE_INVOICE:开票日{inv_d}≤期初豁免截止日{period.grace_invoice_until}，"
            f"订单在上一年（查询周期{period.start}~{period.end}，账期{period.account_days}天）"
        )
        return STATUS_GRACE_INVOICE, False, reasons, BIAS_INV

    if same_model_ois is None:
        same_model_ois = [oi for oi, o in enumerate(orders) if o["_model_key"] == model]
    same_model_orders = [orders[oi] for oi in same_model_ois]
    if not same_model_orders:
        reasons.append("ONLY_INVOICE:该型号仅出现在发票中")
        return STATUS_INV_NO_ORDER, False, reasons, BIAS_INV

    # 找「对本发票合法」的订单：等量 + 日期规则（含软窗口）
    def legal_order_indexes(ignore_used: bool):
        idxs = []
        for oi in same_model_ois:
            if not ignore_used and oi in used_ord:
                continue
            o = orders[oi]
            if o["_ship_dt"] is None or inv["_inv_dt"] is None:
                continue
            if not date_pair_ok(o["_ship_dt"], inv["_inv_dt"], o.get("_recv_dt")):
                continue
            if o["_qty"] is None or inv["_qty"] is None:
                continue
            if abs(o["_qty"] - inv["_qty"]) > 1e-9:
                continue
            idxs.append(oi)
        return idxs

    legal_all = legal_order_indexes(ignore_used=True)
    legal_free = legal_order_indexes(ignore_used=False)
    if legal_all and not legal_free:
        reasons.append(
            "OCCUPIED:存在数量相等且日期合法的同型号订单，但已被其他发票行占用"
        )
        reasons.append("偏向说明:表面像发票无订单，根因更偏配对竞争(其他)")
        return STATUS_SAME_OCCUPIED, True, _dedupe_reasons(reasons), BIAS_OTHER

    equal_qty = [
        o
        for o in same_model_orders
        if inv["_qty"] is not None
        and o["_qty"] is not None
        and abs(o["_qty"] - inv["_qty"]) < 1e-9
    ]
    if equal_qty and not legal_all:
        reasons.append(
            "DATE_ORDER:同型号存在等量订单，但不满足日期规则"
            f"（交货早于开票，或开票早于交货≤{INVOICE_EARLY_DAYS}天且开票≥接单）"
        )
        for o in equal_qty:
            if o["_ship_dt"] and inv["_inv_dt"]:
                recv = o.get("_recv_dt")
                recv_s = f"{recv:%Y-%m-%d}" if recv else "空"
                reasons.append(
                    f"DATE_ORDER:交货{o['_ship_dt']:%Y-%m-%d} / "
                    f"开票{inv['_inv_dt']:%Y-%m-%d} / 接单{recv_s}"
                )
                break
        reasons.append("偏向说明:发票侧未配上，更偏向发票无订单")
        return STATUS_SAME_DATE, True, _dedupe_reasons(reasons), BIAS_INV

    o_total = so.get(model, 0)
    i_total = si.get(model, 0)
    reasons.append(
        f"QTY_MISALIGN:同型号无等量可配订单行（本行数量{inv['_qty']}；"
        f"订单数量{[x['_qty'] for x in same_model_orders]}）"
    )
    if abs(o_total - i_total) > 1e-9:
        reasons.append(
            f"MODEL_QTY_GAP:该型号订单数量合计与发票数量合计不等"
            f"（订单合计{o_total:g} / 发票合计{i_total:g}）"
        )
        reasons.append(
            "QTY_POOL_HINT:型号级数量可冲减（合计接近），但行级一对一失败，建议人工拆分/合并核对"
        )
    reasons.append("偏向说明:发票侧未配上，更偏向发票无订单")
    return STATUS_SAME_QTY, True, _dedupe_reasons(reasons), BIAS_INV


def make_result_row(order, invoice, status, need_manual, reasons, bias=None, **extra):
    side = "order" if order and not invoice else ("invoice" if invoice and not order else None)
    if bias is None:
        bias = bias_for(status, side)
    row = {
        "order": order,
        "invoice": invoice,
        "status": status,
        "need_manual": need_manual,
        "reasons": reasons or [],
        "bias": bias,
        "verify_priority": verify_priority_for(status),
    }
    row.update(extra)
    return row


def reconcile(
    orders,
    invoices,
    period: PeriodConfig | None = None,
    progress_callback=None,
):
    """orders/invoices 可为全量（内部只对未排除行做匹配；排除行仍输出）。

    progress_callback 可选：callback(pct: int, message: str)，匹配阶段建议占用 55～74。
    """
    active_orders, excluded_orders = split_orders_for_match(orders)
    active_invoices, excluded_invoices = split_invoices_for_match(invoices)
    so, si = model_qty_totals(active_orders, active_invoices)
    inv_by_model = build_model_index(active_invoices)
    ord_by_model = build_model_index(active_orders)
    used_inv = set()  # 已成功配对的发票下标（相对 active_invoices）
    matched_ord = set()  # 已成功配对的订单下标（相对 active_orders）
    results = []

    last_pct = -1

    def _match_progress(done: int, total: int, lo: int, hi: int, label: str):
        nonlocal last_pct
        if progress_callback is None:
            return
        if total <= 0:
            pct = hi
        else:
            pct = lo + int((hi - lo) * (done / total))
        pct = max(lo, min(hi, pct))
        # 仅在百分比前进时回调，避免刷屏
        if pct <= last_pct and done < total:
            return
        last_pct = pct
        progress_callback(pct, f"{label}（{done}/{total}）")

    n_ord = len(active_orders)
    for oi, order in enumerate(active_orders):
        reasons = []
        model = order.get("_model_key")
        same_js = inv_by_model.get(model, []) if model else []

        if not order["_model_key"] or order["_qty"] is None or order["_ship_dt"] is None:
            bad = []
            if not order["_model_key"]:
                bad.append("型号")
            if order["_qty"] is None:
                bad.append("数量")
            if order["_ship_dt"] is None:
                bad.append("交货日期")
            results.append(
                make_result_row(
                    order,
                    None,
                    "数据异常",
                    True,
                    [f"BAD_FIELD:关键字段缺失或无法解析：{'/'.join(bad)}"],
                )
            )
            _match_progress(oi + 1, n_ord, 55, 68, "正在匹配订单")
            continue

        cands = find_candidates(
            order, active_invoices, used_inv, candidate_js=same_js
        )
        if not cands:
            status, need, rs, bias = classify_unmatched_order(
                order,
                active_invoices,
                so,
                si,
                used_inv,
                period,
                same_model_js=same_js,
            )
            results.append(
                make_result_row(order, None, status, need, rs, bias=bias)
            )
            _match_progress(oi + 1, n_ord, 55, 68, "正在匹配订单")
            continue

        best_j, multi = pick_best_candidate(order, active_invoices, cands)
        inv = active_invoices[best_j]
        used_inv.add(best_j)
        matched_ord.add(oi)

        price_diff = None
        if order["_user_price"] is not None and inv["_tax_price"] is not None:
            price_diff = order["_user_price"] - inv["_tax_price"]

        if multi:
            reasons.append(
                "MULTI_CAND:同型号存在多条等量候选，按价格优先+FIFO取一条，可能配错"
            )

        if date_pair_used_soft_window(
            order["_ship_dt"], inv["_inv_dt"], order.get("_recv_dt")
        ):
            ship_d = to_date_only(order["_ship_dt"])
            inv_d = to_date_only(inv["_inv_dt"])
            recv_d = to_date_only(order.get("_recv_dt"))
            early = (ship_d - inv_d).days if ship_d and inv_d else 0
            reasons.append(
                f"DATE_EARLY_OK:开票早于交货{early}天（≤{INVOICE_EARLY_DAYS}），"
                f"且开票日{inv_d}≥接单日{recv_d}，按软窗口判定日期合法"
            )

        append_model_text_diffs(reasons, order.get("型号"), inv.get("规格型号"))

        if abs(so.get(model, 0) - si.get(model, 0)) > 1e-9:
            reasons.append(
                f"MODEL_QTY_GAP:该型号订单数量合计与发票数量合计不等"
                f"（订单合计{so.get(model,0):g} / 发票合计{si.get(model,0):g}）"
            )

        if order["_user_price"] is None or inv["_tax_price"] is None:
            reasons.append("BAD_FIELD:关键字段缺失或无法解析：用户价格或含税单价")
            status = STATUS_PRICE_MISSING
            need_manual = True
        elif price_diff < -PRICE_TOLERANCE:
            over = inv["_tax_price"] - order["_user_price"]
            reasons.append(
                f"PRICE_OVER:实际含税单价高于申请用户价格：用户价格{order['_user_price']:g} vs "
                f"含税单价{inv['_tax_price']:g}，超{over:g}"
            )
            status = STATUS_PRICE_OVER
            need_manual = True
        elif price_diff > PRICE_TOLERANCE:
            under = order["_user_price"] - inv["_tax_price"]
            reasons.append(
                f"PRICE_UNDER:实际含税单价低于申请用户价格：用户价格{order['_user_price']:g} vs "
                f"含税单价{inv['_tax_price']:g}，低{under:g}"
            )
            status = STATUS_PRICE_UNDER
            need_manual = True
        else:
            status = STATUS_MATCHED
            need_manual = False

        results.append(
            make_result_row(
                order,
                inv,
                status,
                need_manual,
                reasons,
                bias="",
                price_diff=price_diff,
            )
        )
        _match_progress(oi + 1, n_ord, 55, 68, "正在匹配订单")

    unmatched_inv_js = [j for j in range(len(active_invoices)) if j not in used_inv]
    n_uinv = len(unmatched_inv_js)
    for ui, j in enumerate(unmatched_inv_js):
        inv = active_invoices[j]
        if not inv["_model_key"] or inv["_qty"] is None or inv["_inv_dt"] is None:
            bad = []
            if not inv["_model_key"]:
                bad.append("规格型号")
            if inv["_qty"] is None:
                bad.append("数量")
            if inv["_inv_dt"] is None:
                bad.append("开票日期")
            results.append(
                make_result_row(
                    None,
                    inv,
                    "数据异常",
                    True,
                    [f"BAD_FIELD:关键字段缺失或无法解析：{'/'.join(bad)}"],
                )
            )
            _match_progress(ui + 1, n_uinv, 68, 74, "正在归类未配发票")
            continue
        status, need, rs, bias = classify_unmatched_invoice(
            inv,
            active_orders,
            so,
            si,
            matched_ord,
            period,
            same_model_ois=ord_by_model.get(inv["_model_key"], []),
        )
        results.append(make_result_row(None, inv, status, need, rs, bias=bias))
        _match_progress(ui + 1, n_uinv, 68, 74, "正在归类未配发票")

    for order in excluded_orders:
        results.append(
            make_result_row(
                order,
                None,
                "不参与比对",
                False,
                list(order.get("_exclude_reasons") or []),
                bias="",
            )
        )

    for inv in excluded_invoices:
        results.append(
            make_result_row(
                None,
                inv,
                "不参与比对",
                False,
                list(inv.get("_exclude_reasons") or []),
                bias="",
            )
        )

    if progress_callback is not None:
        progress_callback(74, "匹配收尾完成")

    return results, so, si


# 预处理/明细用的全量字段（顺序供预处理 sheet；明细顺序见下方 DETAIL_*）
ORDER_OUT_COLS = [
    "交货日期",
    "订单号",
    "型号",
    "数量",
    "用户价格",
    "E率",
    "价格合计",
    "价格",
    "E价",
    "接单日期",
    "订单状态",
    "用户",
    "战略标记",
    "统计对象",
    "产品二级分类",
]

INVOICE_OUT_COLS = [
    "数电发票号码",
    "货物或应税劳务名称",
    "规格型号",
    "数量",
    "含税单价",
    "序号",
    "发票代码",
    "发票号码",
    "单位",
    "单价",
    "金额",
    "税率",
    "税额",
    "价税合计",
    "发票状态",
    "购买方名称",
    "销方名称",
    "备注",
    "开票日期",
]

META_COLS = [
    "核实优先级",
    "是否需要解释",
    "匹配状态",
    "问题偏向",
    "是否需人工核对",
    "异议原因",
]

# 匹配明细：按下列显式顺序排布；HIDDEN 集合内的列写入但默认隐藏
DETAIL_INVOICE_COLS = list(INVOICE_OUT_COLS)
DETAIL_ORDER_COLS = list(ORDER_OUT_COLS)
DETAIL_META_COLS = list(META_COLS)

INVOICE_HIDDEN_FIELDS = {
    "序号",
    "发票代码",
    "发票号码",
    "单位",
    "单价",
    "金额",
    "税率",
    "税额",
    "价税合计",
    "发票状态",
    "购买方名称",
    "销方名称",
    "备注",
}
ORDER_HIDDEN_FIELDS = {
    "价格合计",
    "价格",
    "E价",
    "接单日期",
    "订单状态",
    "用户",
}
META_HIDDEN_FIELDS: set[str] = set()

# 兼容旧名：可见列 = 明细列中未隐藏者
INVOICE_VISIBLE_COLS = [f for f in DETAIL_INVOICE_COLS if f not in INVOICE_HIDDEN_FIELDS]
ORDER_VISIBLE_COLS = [f for f in DETAIL_ORDER_COLS if f not in ORDER_HIDDEN_FIELDS]
META_VISIBLE_COLS = [f for f in DETAIL_META_COLS if f not in META_HIDDEN_FIELDS]


def build_result_columns():
    """返回 [(kind, field, header), ...]：发票 → 订单 → 匹配结果。"""
    cols = [("发票", f, f"发票:{f}") for f in DETAIL_INVOICE_COLS]
    cols += [("订单", f, f"订单:{f}") for f in DETAIL_ORDER_COLS]
    cols += [("meta", f, f) for f in DETAIL_META_COLS]
    return cols


RESULT_COLUMNS = build_result_columns()
RESULT_HEADERS = [c[2] for c in RESULT_COLUMNS]
ORDER_COL_INDEXES = {
    i + 1 for i, (k, _, _) in enumerate(RESULT_COLUMNS) if k == "订单"
}
DATE_COL_INDEXES = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if f in ("交货日期", "接单日期", "开票日期")
}

ZONE_INV_VISIBLE_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "发票" and f not in INVOICE_HIDDEN_FIELDS
}
ZONE_INV_HIDDEN_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "发票" and f in INVOICE_HIDDEN_FIELDS
}
ZONE_ORD_VISIBLE_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "订单" and f not in ORDER_HIDDEN_FIELDS
}
ZONE_ORD_HIDDEN_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "订单" and f in ORDER_HIDDEN_FIELDS
}
ZONE_META_VISIBLE_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "meta" and f not in META_HIDDEN_FIELDS
}
ZONE_META_HIDDEN_COLS = {
    i + 1
    for i, (k, f, _) in enumerate(RESULT_COLUMNS)
    if k == "meta" and f in META_HIDDEN_FIELDS
}
HIDDEN_COL_INDEXES = (
    ZONE_INV_HIDDEN_COLS | ZONE_ORD_HIDDEN_COLS | ZONE_META_HIDDEN_COLS
)
META_COL_INDEX = {
    f: i + 1 for i, (k, f, _) in enumerate(RESULT_COLUMNS) if k == "meta"
}
# 冻结在整段发票列之后，便于对照订单与匹配结果
FREEZE_AFTER_COL = len(DETAIL_INVOICE_COLS) + 1


def fmt_cell(v):
    if isinstance(v, datetime):
        return v
    return v


def write_header(ws, headers, start_col=1):
    for i, h in enumerate(headers):
        cell = ws.cell(1, start_col + i, h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, horizontal="center")
        cell.border = THIN


def result_cell_value(r, kind, field):
    if r.get("_summary"):
        if kind == "发票" and field == "数电发票号码":
            return "汇总"
        if kind == "发票" and field == "规格型号":
            return r.get("summary_model") or ""
        if kind == "发票" and field == "数量":
            return r.get("summary_inv_qty")
        if kind == "订单" and field == "数量":
            return r.get("summary_ord_qty")
        return None
    if kind == "meta":
        if field == "匹配状态":
            return r["status"]
        if field == "核实优先级":
            return r.get("verify_priority") or ""
        if field == "是否需要解释":
            return need_explain_for(r.get("verify_priority") or "")
        if field == "是否需人工核对":
            return "是" if r["need_manual"] else "否"
        if field == "问题偏向":
            return r.get("bias") or ""
        if field == "异议原因":
            return "; ".join(r.get("reasons") or [])
        return None
    src = r["order"] if kind == "订单" else r["invoice"]
    if not src:
        return None
    return fmt_cell(src.get(field))


def _zone_header_style(col_idx: int):
    # 发票可见：绿底白字；发票其余：浅绿深字；订单可见：蓝底白字；订单其余：浅蓝深字；匹配：深紫白字
    if col_idx in ZONE_INV_VISIBLE_COLS:
        return FILL_ZONE_B_INV_HEADER, FONT_HEADER_WHITE
    if col_idx in ZONE_INV_HIDDEN_COLS:
        return FILL_ZONE_C_HEADER, FONT_HEADER_DARK
    if col_idx in ZONE_ORD_VISIBLE_COLS:
        return FILL_ZONE_B_ORD_HEADER, FONT_HEADER_WHITE
    if col_idx in ZONE_ORD_HIDDEN_COLS:
        return FILL_ZONE_D_HEADER, FONT_HEADER_DARK
    if col_idx in ZONE_META_VISIBLE_COLS or col_idx in ZONE_META_HIDDEN_COLS:
        return FILL_ZONE_A_HEADER, FONT_HEADER_WHITE
    return HEADER_FILL, Font(bold=True)


def _zone_data_fill(col_idx: int):
    """发票统一绿底、订单统一蓝底、匹配结果默认浅紫（核实高时另涂浅橙底）。"""
    if col_idx in ZONE_INV_VISIBLE_COLS or col_idx in ZONE_INV_HIDDEN_COLS:
        return FILL_INV_DATA
    if col_idx in ZONE_ORD_VISIBLE_COLS or col_idx in ZONE_ORD_HIDDEN_COLS:
        return FILL_ORD_DATA
    return FILL_ZONE_A_DATA


def write_detail_header(ws):
    for i, h in enumerate(RESULT_HEADERS, start=1):
        cell = ws.cell(1, i, h)
        fill, font = _zone_header_style(i)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = THIN
    ws.row_dimensions[1].height = 32


def write_result_rows(
    ws,
    rows,
    progress_callback=None,
    progress_lo=80,
    progress_hi=88,
):
    """写入匹配明细样式行。

    发票列恒绿底、订单列恒蓝底（不整行黄/橙覆盖）；
    核实优先级「高」：发票/订单列红字，比对结果列浅橙底。
    """
    write_detail_header(ws)
    n_cols = len(RESULT_COLUMNS)
    idx_verify = META_COL_INDEX.get("核实优先级")
    idx_explain = META_COL_INDEX.get("是否需要解释")
    idx_status = META_COL_INDEX.get("匹配状态")
    idx_bias = META_COL_INDEX.get("问题偏向")
    idx_reason = META_COL_INDEX.get("异议原因")

    n_rows = len(rows)
    last_pct = -1

    def _row_progress(done: int):
        nonlocal last_pct
        if progress_callback is None or n_rows <= 0:
            return
        pct = progress_lo + int((progress_hi - progress_lo) * (done / n_rows))
        pct = max(progress_lo, min(progress_hi, pct))
        if pct <= last_pct and done < n_rows:
            return
        last_pct = pct
        progress_callback(pct, f"正在写入匹配明细（{done}/{n_rows}）")

    for ri, r in enumerate(rows, start=2):
        if r.get("_summary"):
            for ci, (kind, field, _) in enumerate(RESULT_COLUMNS, start=1):
                cell = ws.cell(ri, ci, result_cell_value(r, kind, field))
                cell.border = THIN
                cell.fill = FILL_MODEL_SUMMARY
                cell.font = FONT_MODEL_SUMMARY
                if kind != "meta" and field == "数量" and isinstance(cell.value, (int, float)):
                    cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="right")
            _row_progress(ri - 1)
            continue

        vp = r.get("verify_priority") or ""
        is_high = vp == VERIFY_HIGH
        is_error = r["status"] == "数据异常"

        for ci, (kind, field, _) in enumerate(RESULT_COLUMNS, start=1):
            cell = ws.cell(ri, ci, result_cell_value(r, kind, field))
            cell.border = THIN
            cell.font = FONT_DEFAULT
            if ci in DATE_COL_INDEXES and isinstance(cell.value, (datetime, date)):
                cell.number_format = "yyyy-mm-dd"
            if kind == "订单" and field == "E率" and isinstance(cell.value, (int, float)):
                cell.number_format = '0"%"'

            if kind == "发票":
                cell.fill = FILL_INV_DATA
            elif kind == "订单":
                cell.fill = FILL_ORD_DATA
            elif is_high:
                cell.fill = FILL_META_HIGH
            elif is_error:
                cell.fill = ERROR_FILL
            else:
                cell.fill = FILL_ZONE_A_DATA

            if is_high and kind in ("发票", "订单"):
                cell.font = FONT_HIGH_INV_ORD
            if idx_verify and ci == idx_verify:
                if is_high:
                    cell.font = FONT_VERIFY_HIGH
                elif vp == VERIFY_MID:
                    cell.font = FONT_VERIFY_MID
            elif idx_explain and ci == idx_explain and cell.value == NEED_EXPLAIN_YES:
                cell.font = FONT_EXPLAIN
            elif idx_status and ci == idx_status and r["status"] in STATUS_RED_BOLD:
                cell.font = FONT_STATUS_OVER
            elif idx_bias and ci == idx_bias and (r.get("bias") or "") == BIAS_OTHER:
                cell.font = FONT_BIAS_OTHER
            elif idx_reason and ci == idx_reason:
                cell.font = FONT_REASON

            if kind != "meta" and field in ("数量", "用户价格", "含税单价", "价税合计", "价格合计", "单价"):
                cell.alignment = Alignment(horizontal="right")
            if kind != "meta" and field == "数量" and isinstance(cell.value, (int, float)):
                cell.number_format = "0"
            if (
                kind != "meta"
                and field in ("单价", "含税单价", "用户价格")
                and isinstance(cell.value, (int, float))
            ):
                cell.number_format = "0.00"

        _row_progress(ri - 1)

    last = max(1, ws.max_row)
    ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{last}"
    ws.freeze_panes = f"{get_column_letter(FREEZE_AFTER_COL)}2"
    for ci in HIDDEN_COL_INDEXES:
        ws.column_dimensions[get_column_letter(ci)].hidden = True
    if progress_callback is not None:
        progress_callback(progress_hi, "正在调整列宽…")
    _autosize(ws, max_width=28)


def _row_sort_datetime(r) -> datetime:
    """同型号组内排序主日期：优先交货日期；交货日期为空（如仅发票行）则用开票日期。"""
    order = r.get("order")
    inv = r.get("invoice")
    if order:
        ship_dt = order.get("_ship_dt")
        if ship_dt is not None:
            return ship_dt
    if inv:
        inv_dt = inv.get("_inv_dt")
        if inv_dt is not None:
            return inv_dt
    return datetime.max


def _row_secondary_datetime(r) -> datetime:
    """次要日期：另一侧日期，便于同日行稳定对照。"""
    order = r.get("order")
    inv = r.get("invoice")
    ship_dt = order.get("_ship_dt") if order else None
    inv_dt = inv.get("_inv_dt") if inv else None
    # 主排序用了交货日 → 次要用开票日；主排序用了开票日 → 次要用交货日（通常为空）
    if ship_dt is not None:
        return inv_dt or datetime.max
    return ship_dt or datetime.max


def _detail_model_key(r) -> str:
    o = r.get("order")
    i = r.get("invoice")
    return (
        (o.get("_model_key") if o else None)
        or (i.get("_model_key") if i else None)
        or ""
    )


def _detail_model_label(rows) -> str:
    """汇总行展示用型号原文：优先发票规格型号，其次订单型号。"""
    for r in rows:
        inv = r.get("invoice")
        if inv and inv.get("规格型号") not in (None, ""):
            return str(inv.get("规格型号"))
    for r in rows:
        order = r.get("order")
        if order and order.get("型号") not in (None, ""):
            return str(order.get("型号"))
    return _detail_model_key(rows[0]) if rows else ""


def make_model_summary_row(group_rows):
    """每个型号组结束后的汇总行（深灰底，仅填数电票号/规格型号/双方数量）。"""
    inv_qty = 0.0
    ord_qty = 0.0
    has_inv = False
    has_ord = False
    for r in group_rows:
        order = r.get("order")
        inv = r.get("invoice")
        if order and order.get("_qty") is not None:
            ord_qty += float(order["_qty"])
            has_ord = True
        if inv and inv.get("_qty") is not None:
            inv_qty += float(inv["_qty"])
            has_inv = True
    return {
        "_summary": True,
        "summary_model": _detail_model_label(group_rows),
        "summary_inv_qty": int(round(inv_qty)) if has_inv else None,
        "summary_ord_qty": int(round(ord_qty)) if has_ord else None,
        "order": None,
        "invoice": None,
        "status": "",
        "need_manual": False,
        "reasons": [],
        "bias": "",
        "verify_priority": "",
    }


def with_model_summary_rows(sorted_rows):
    """在每个型号组之后插入一行汇总。"""
    out = []
    i = 0
    n = len(sorted_rows)
    while i < n:
        key = _detail_model_key(sorted_rows[i])
        group = []
        while i < n and _detail_model_key(sorted_rows[i]) == key:
            group.append(sorted_rows[i])
            i += 1
        out.extend(group)
        out.append(make_model_summary_row(group))
    return out


def write_detail_sheet(wb, results, progress_callback=None):
    ws = wb.create_sheet("匹配明细", 0)

    def sort_key(r):
        """相同型号聚拢；组内优先按交货日期升序，无交货日期（仅发票）则按开票日期。"""
        return (
            _detail_model_key(r),
            _row_sort_datetime(r),
            _row_secondary_datetime(r),
            r.get("status") or "",
        )

    if progress_callback is not None:
        progress_callback(80, "正在排序匹配明细…")
    sorted_rows = sorted(results, key=sort_key)
    write_result_rows(
        ws,
        with_model_summary_rows(sorted_rows),
        progress_callback=progress_callback,
        progress_lo=80,
        progress_hi=88,
    )
    return ws


def write_summary_sheet(wb, results, orders, invoices, so, si):
    ws = wb.create_sheet("型号汇总", 1)
    headers = [
        "型号",
        "订单行数",
        "发票行数",
        "订单数量合计",
        "发票数量合计",
        "数量差(订-票)",
        "已核销",
        "实际售价高于申请价格",
        "实际售价低于申请价格",
        "价格缺失",
        "同型号开票早于发货",
        "同型号数量无法对齐",
        "同型号行已被其他票/单占用",
        "开票在下一年",
        "订单在上一年",
        "订单未开票",
        "发票无订单",
        "不参与比对",
        "数据异常",
        "需人工核对行数",
        "汇总备注",
    ]
    write_header(ws, headers)

    models = sorted(
        set(so)
        | set(si)
        | {
            r["order"]["_model_key"]
            for r in results
            if r["order"] and r["order"]["_model_key"]
        }
        | {
            r["invoice"]["_model_key"]
            for r in results
            if r["invoice"] and r["invoice"]["_model_key"]
        }
    )
    # 汇总展示用原文（优先订单型号）
    key_label = {}
    for o in orders:
        k = o.get("_model_key")
        if k and k not in key_label and o.get("型号") is not None:
            key_label[k] = str(o.get("型号"))
    for i in invoices:
        k = i.get("_model_key")
        if k and k not in key_label and i.get("规格型号") is not None:
            key_label[k] = str(i.get("规格型号"))

    stats = defaultdict(lambda: defaultdict(int))
    manual_by_model = defaultdict(int)
    for r in results:
        m = None
        if r["order"]:
            m = r["order"]["_model_key"]
        elif r["invoice"]:
            m = r["invoice"]["_model_key"]
        if not m:
            m = "(空型号)"
        stats[m][r["status"]] += 1
        if r["need_manual"]:
            manual_by_model[m] += 1

    order_lines = defaultdict(int)
    inv_lines = defaultdict(int)
    for o in orders:
        if o["_model_key"]:
            order_lines[o["_model_key"]] += 1
    for i in invoices:
        if i["_model_key"]:
            inv_lines[i["_model_key"]] += 1

    ri = 2
    for m in models:
        oq = so.get(m, 0)
        iq = si.get(m, 0)
        notes = []
        if abs(oq - iq) > 1e-9:
            notes.append(f"MODEL_QTY_GAP:数量差{oq-iq:g}")
        if order_lines[m] and not inv_lines[m]:
            notes.append("ONLY_ORDER")
        if inv_lines[m] and not order_lines[m]:
            notes.append("ONLY_INVOICE")
        vals = [
            key_label.get(m, m),
            order_lines[m],
            inv_lines[m],
            oq,
            iq,
            oq - iq,
            stats[m][STATUS_MATCHED],
            stats[m][STATUS_PRICE_OVER],
            stats[m][STATUS_PRICE_UNDER],
            stats[m][STATUS_PRICE_MISSING],
            stats[m][STATUS_SAME_DATE],
            stats[m][STATUS_SAME_QTY],
            stats[m][STATUS_SAME_OCCUPIED],
            stats[m][STATUS_GRACE_ORDER],
            stats[m][STATUS_GRACE_INVOICE],
            stats[m][STATUS_ORDER_NO_INV],
            stats[m][STATUS_INV_NO_ORDER],
            stats[m]["不参与比对"],
            stats[m]["数据异常"],
            manual_by_model[m],
            "; ".join(notes),
        ]
        for ci, v in enumerate(vals, start=1):
            cell = ws.cell(ri, ci, v)
            cell.border = THIN
            if abs(oq - iq) > 1e-9 or manual_by_model[m]:
                if ci == 1 or ci >= 6:
                    cell.fill = MANUAL_FILL
        ri += 1

    # 总计
    ws.cell(ri + 1, 1, "【全局状态统计】").font = Font(bold=True)
    from collections import Counter

    c = Counter(r["status"] for r in results)
    r2 = ri + 2
    ws.cell(r2, 1, "状态")
    ws.cell(r2, 2, "行数")
    write_header(ws, ["状态", "行数"])
    # overwrite nicely
    for i, (k, v) in enumerate(sorted(c.items()), start=r2):
        ws.cell(i, 1, k).border = THIN
        ws.cell(i, 2, v).border = THIN
    ws.cell(r2 + len(c) + 1, 1, "需人工核对合计")
    ws.cell(r2 + len(c) + 1, 2, sum(1 for r in results if r["need_manual"]))
    _autosize(ws, max_width=36)


def write_rules_sheet(
    wb,
    period: PeriodConfig | None = None,
    mode: str = MODE_DAILY,
    stat_objects: list[str] | None = None,
):
    ws = wb.create_sheet("规则说明", 2)
    mode_label = MODE_LABELS.get(mode, mode)
    mode_lines = [
        f"监察方式：{mode_label}",
    ]
    if mode == MODE_STRATEGIC:
        cats = "、".join(stat_objects or []) or "（未指定）"
        mode_lines.append(f"战略品统计对象（已选）：{cats}")
        mode_lines.append(
            "战略品监察：仅比对所选统计对象下的订单型号及同型号发票；未选中数据不写入结果"
        )
    else:
        mode_lines.append("日常监察：全量型号比对")
    mode_lines.append("")

    period_lines = []
    if period:
        period_lines = [
            f"发票查询周期：{period.start} ~ {period.end}",
            f"账期：{period.account_days} 天",
            f"期初豁免截止日（开票日≤此日无订单→订单在上一年）：{period.grace_invoice_until}",
            f"期末豁免起始日（交货日≥此日无发票→开票在下一年）：{period.grace_order_from}",
            "",
        ]
    lines = [
        "比对规则说明",
        "",
        *mode_lines,
        *period_lines,
        "一、预处理",
        "1. 订单：新增「价格合计」= 数量 × 用户价格；「E率」= round(((价格/1.13/E价)-1)×100) 整数百分比；交货日期/接单日期仅保留年月日；按型号、交货日期升序",
        "2. 发票：数量/单价由文本转数字（数量取整、单价保留2位小数）；新增「含税单价」= 单价 × 1.13；"
        "开票日期仅保留年月日；按规格型号、开票日期升序",
        "3. 配对前排除（不删除）：发票数量为负，或发票状态含「已红冲-全额」，"
        "或订单号去前导空白后以S/s开头 → 状态「不参与比对」，仍输出；"
        "「已红冲-部分」仍参与比对",
        "",
        "二、匹配键与约束",
        "1. 关联键：订单.型号 ↔ 发票.规格型号（去全部空白后 casefold，大小写与空格不敏感）",
        f"2. 日期规则：交货日 < 开票日；或开票早于交货 1～{INVOICE_EARLY_DAYS} 天且开票日≥接单日"
        "（同日开票=交货仍不合法；软窗口缺接单日不合法）",
        "3. 行级一对一；优先数量完全相等，再比价格接近，再 FIFO",
        f"4. 价格容差：|用户价格 - 含税单价| ≤ {PRICE_TOLERANCE}",
        "5. 业务：含税单价不得高于用户价格（申请含税价）",
        "6. 账期：起始日+N天-1天=期初豁免截止；结束日-N天+1天=期末豁免起始",
        "",
        "三、状态定义",
        "已核销：日期合法 + 数量相等 + |用户价格-含税单价|≤容差",
        "实际售价高于申请价格 / 实际售价低于申请价格 / 价格缺失：已配对价格问题 → 需人工",
        "开票在下一年：订单未配上且交货日≥期末豁免起始日 → 不强制人工",
        "订单在上一年：发票未配上且开票日≤期初豁免截止日 → 不强制人工",
        "订单未开票 / 发票无订单：豁免区外的一侧缺失 → 不强制人工",
        "同型号开票早于发货：同型号有等量对方但不满足先发货后开票 → 需人工；偏向见「问题偏向」列",
        "同型号数量无法对齐：同型号无等量可配行 → 需人工；偏向见「问题偏向」列",
        "同型号行已被其他票/单占用：本可配但候选被占用 → 需人工；问题偏向=其他",
        "不参与比对 / 数据异常",
        "",
        "四、核实优先级（筛选用：找数据提供方沟通）",
        "高：实际售价高于申请价格；订单未开票；发票无订单；同型号开票早于发货；同型号数量无法对齐；同型号行已被其他票/单占用",
        "中：数据异常；价格缺失",
        "空（不沟通）：已核销；开票在下一年；订单在上一年；不参与比对；实际售价低于申请价格",
        "核实优先级为「高」时，「是否需要解释」填「需解释」",
        "筛选用法：在「匹配明细」对「核实优先级」筛选「高」或「中」",
        "",
        "四-附、匹配明细列序（取消成对对照区；其余列保留但默认隐藏）",
        "1. 发票列序：数电发票号码、货物或应税劳务名称、规格型号、数量、含税单价、"
        "序号/代码/号码/单位/单价/金额/税率/税额/价税合计/状态/购销方/备注（隐藏）、开票日期",
        "2. 订单列序：交货日期、订单号、型号、数量、用户价格、E率、"
        "价格合计/价格/E价/接单日期/订单状态/用户（隐藏）、战略标记、统计对象、产品二级分类",
        "3. 匹配结果列：核实优先级、是否需要解释、匹配状态、问题偏向、是否需人工核对、异议原因",
        "明细排序：同型号内优先按交货日期升序；交货日期为空（有发票无对应订单）则按开票日期升序",
        "",
        "五、问题偏向（三选一，仅未配相关行填写）",
        "订单未开票：订单侧未配或同型号日期/数量问题在订单行",
        "发票无订单：发票侧未配或同型号日期/数量问题在发票行",
        "其他：配对竞争占用、数据异常等",
        "",
        "六、匹配明细颜色规则",
        "【表头】发票可见=绿底白字；发票其余=浅绿底深字；订单可见=蓝底白字；订单其余=浅蓝底深字；匹配结果=深紫底白字",
        "【数据底色】发票列统一浅绿；订单列统一浅蓝；匹配结果默认浅紫；发票/订单列不做黄/橙整行覆盖",
        "【核实高】发票/订单列=红字+原绿/蓝底；比对结果列=浅橙底；核实优先级=红字加粗；「需解释」=橘色加粗",
        "【型号汇总行】每型号组末：数电发票号码=汇总，规格型号=该型号，发票数量/订单数量=组内合计；浅灰底黑字加粗",
        "【字体】匹配明细统一使用「等线」字体",
        "【其他】数据异常仅比对结果列浅红底；核实「中」仅核实优先级褐字加粗；"
        "匹配状态「实际售价高于申请价格」「订单未开票」「发票无订单」=红色加粗；"
        "问题偏向「其他」=灰紫字；异议原因=灰色小字",
        "【不做】已核销不用绿色字；不为每种匹配状态单独设行色",
        "",
        "七、异议原因代码（节选）",
        "PRICE_OVER / PRICE_UNDER / MULTI_CAND / CASE_DIFF / SPACE_DIFF",
        "GRACE_ORDER / GRACE_INVOICE / OCCUPIED / QTY_MISALIGN / DATE_ORDER / DATE_EARLY_OK",
        "EXCL_NEG_QTY / EXCL_RED_FLUSH / EXCL_ORDER_NO_S / MODEL_QTY_GAP / BAD_FIELD",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(i, 1, line)
        if i == 1:
            ws.cell(i, 1).font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 100


def write_preprocessed_sheets(wb, orders, invoices, progress_callback=None):
    def _p(pct, msg):
        if progress_callback is not None:
            progress_callback(pct, msg)

    ws_o = wb.create_sheet("预处理-订单", 3)
    cols = ORDER_OUT_COLS
    write_header(ws_o, cols)
    n_o = len(orders)
    last_pct = -1
    for ri, o in enumerate(orders, start=2):
        for ci, col in enumerate(cols, start=1):
            cell = ws_o.cell(ri, ci, fmt_cell(o.get(col)))
            cell.fill = ORDER_FILL
            cell.border = THIN
            if col in ("交货日期", "接单日期") and isinstance(cell.value, (datetime, date)):
                cell.number_format = "yyyy-mm-dd"
            if col == "E率" and isinstance(cell.value, (int, float)):
                cell.number_format = '0"%"'
        if progress_callback is not None and n_o > 0:
            done = ri - 1
            pct = 91 + int(2 * (done / n_o))  # 91→93
            if pct > last_pct or done >= n_o:
                last_pct = pct
                _p(min(93, pct), f"正在写入预处理订单（{done}/{n_o}）")
    _autosize(ws_o)

    ws_i = wb.create_sheet("预处理-发票", 4)
    # 发票输出用较全列：原关键列 + 含税单价
    inv_cols = [
        "序号",
        "发票代码",
        "发票号码",
        "数电发票号码",
        "销方识别号",
        "销方名称",
        "购方识别号",
        "购买方名称",
        "开票日期",
        "税收分类编码",
        "特定业务类型",
        "货物或应税劳务名称",
        "规格型号",
        "单位",
        "数量",
        "单价",
        "含税单价",
        "金额",
        "税率",
        "税额",
        "价税合计",
        "发票来源",
        "发票票种",
        "发票状态",
        "是否正数发票",
        "发票风险等级",
        "开票人",
        "备注",
    ]
    write_header(ws_i, inv_cols)
    n_i = len(invoices)
    last_pct = -1
    for ri, inv in enumerate(invoices, start=2):
        for ci, col in enumerate(inv_cols, start=1):
            cell = ws_i.cell(ri, ci, fmt_cell(inv.get(col)))
            cell.border = THIN
            if col == "开票日期" and isinstance(cell.value, (datetime, date)):
                cell.number_format = "yyyy-mm-dd"
            if col == "数量" and isinstance(cell.value, (int, float)):
                cell.number_format = "0"
            if col in ("单价", "含税单价") and isinstance(cell.value, (int, float)):
                cell.number_format = "0.00"
        if progress_callback is not None and n_i > 0:
            done = ri - 1
            pct = 93 + int(2 * (done / n_i))  # 93→95
            if pct > last_pct or done >= n_i:
                last_pct = pct
                _p(min(95, pct), f"正在写入预处理发票（{done}/{n_i}）")
    _autosize(ws_i, max_width=24)


def _autosize(ws, max_width=24, sample_rows=80):
    """按前 sample_rows 行估算列宽，避免 ws.columns 物化整表。"""
    max_col = ws.max_column or 1
    max_row = min(ws.max_row or 1, max(1, sample_rows))
    for ci in range(1, max_col + 1):
        length = 0
        for ri in range(1, max_row + 1):
            val = ws.cell(ri, ci).value
            if val is None:
                continue
            length = max(length, min(max_width, len(str(val))))
        ws.column_dimensions[get_column_letter(ci)].width = max(10, length + 2)


def save_preprocessed_sources(orders, invoices, out_dir: Path):
    """可选：另存预处理后的源文件副本到指定目录。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    write_header(ws, ORDER_OUT_COLS)
    for ri, o in enumerate(orders, start=2):
        for ci, col in enumerate(ORDER_OUT_COLS, start=1):
            cell = ws.cell(ri, ci, fmt_cell(o.get(col)))
            cell.fill = ORDER_FILL
    out_o = out_dir / "订单数据_预处理.xlsx"
    wb.save(out_o)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    inv_cols = [
        "序号",
        "发票代码",
        "发票号码",
        "数电发票号码",
        "销方识别号",
        "销方名称",
        "购方识别号",
        "购买方名称",
        "开票日期",
        "税收分类编码",
        "特定业务类型",
        "货物或应税劳务名称",
        "规格型号",
        "单位",
        "数量",
        "单价",
        "含税单价",
        "金额",
        "税率",
        "税额",
        "价税合计",
        "发票来源",
        "发票票种",
        "发票状态",
        "是否正数发票",
        "发票风险等级",
        "开票人",
        "备注",
    ]
    write_header(ws, inv_cols)
    for ri, inv in enumerate(invoices, start=2):
        for ci, col in enumerate(inv_cols, start=1):
            cell = ws.cell(ri, ci, fmt_cell(inv.get(col)))
            if col == "数量" and isinstance(cell.value, (int, float)):
                cell.number_format = "0"
            if col in ("单价", "含税单价") and isinstance(cell.value, (int, float)):
                cell.number_format = "0.00"
    out_i = out_dir / "发票数据_预处理.xlsx"
    wb.save(out_i)
    return out_o, out_i


def run_reconcile(
    order_path,
    invoice_path,
    output_path,
    save_preprocess_files: bool = False,
    period_start=None,
    period_end=None,
    account_days: int = DEFAULT_ACCOUNT_DAYS,
    progress_callback=None,
    mode: str = MODE_DAILY,
    stat_objects: list[str] | None = None,
):
    """执行比对并写出结果。返回统计信息字典。

    period_start/period_end 为空时，按发票开票日 min/max 所在月的月初～月末推断查询周期。
    account_days 为账期天数。
    mode: daily=日常监察全量；strategic=战略品监察（按统计对象过滤，未选中数据不写入）。
    progress_callback 可选，签名为 callback(pct: int, message: str)，百分比单调递增。
    """
    def _progress(pct: int, message: str):
        if progress_callback is None:
            return
        progress_callback(max(0, min(100, int(pct))), message)

    mode = (mode or MODE_DAILY).strip().lower()
    if mode not in (MODE_DAILY, MODE_STRATEGIC):
        raise ValueError(f"未知监察方式：{mode}（应为 daily 或 strategic）")
    selected_stats = [
        s
        for s in (normalize_stat_object(x) for x in (stat_objects or []))
        if s
    ]

    order_path = Path(order_path)
    invoice_path = Path(invoice_path)
    output_path = Path(output_path)

    if not order_path.is_file():
        raise FileNotFoundError(f"订单文件不存在：{order_path}")
    if not invoice_path.is_file():
        raise FileNotFoundError(f"发票文件不存在：{invoice_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    _progress(5, "正在读取订单…")
    _, order_raw, order_sheet = read_sheet_rows(
        order_path, required_headers=ORDER_REQUIRED_HEADERS
    )
    _progress(10, "订单读取完成")
    _progress(12, "正在读取发票…")
    _, inv_raw, inv_sheet = read_sheet_rows(
        invoice_path, required_headers=INVOICE_REQUIRED_HEADERS
    )
    _progress(20, "发票读取完成")
    if not order_raw:
        raise ValueError("订单文件没有有效数据行")
    if not inv_raw:
        raise ValueError("发票文件没有有效数据行")

    _progress(25, "正在预处理订单…")
    orders = preprocess_orders(order_raw)
    order_rows_raw = len(orders)
    _progress(35, "订单预处理完成")
    _progress(38, "正在预处理发票…")
    invoices = preprocess_invoices(inv_raw)
    invoice_rows_raw = len(invoices)

    strategic_model_keys: list[str] = []
    if mode == MODE_STRATEGIC:
        _progress(42, "战略品过滤中…")
        orders, invoices, strategic_model_keys = filter_for_strategic(
            orders, invoices, selected_stats
        )
        _progress(
            44,
            f"战略品过滤完成：订单 {order_rows_raw}→{len(orders)}，"
            f"发票 {invoice_rows_raw}→{len(invoices)}，型号 {len(strategic_model_keys)}",
        )

    _active_orders, excluded_orders = split_orders_for_match(orders)
    _active_inv, excluded_invoices = split_invoices_for_match(invoices)
    _progress(45, "发票预处理完成")

    period_inferred = False
    if period_start is not None and period_end is not None:
        period = build_period_config(period_start, period_end, account_days)
    else:
        period = infer_period_from_invoices(invoices, account_days)
        period_inferred = True
    _progress(50, "查询周期已确定")

    _progress(55, "正在匹配订单与发票…")
    results, so, si = reconcile(
        orders, invoices, period=period, progress_callback=_progress
    )
    _progress(75, "匹配完成")

    _progress(80, "正在写入结果…")
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    write_detail_sheet(wb, results, progress_callback=_progress)
    _progress(88, "匹配明细已写完")
    write_summary_sheet(wb, results, orders, invoices, so, si)
    _progress(90, "已写型号汇总")
    write_rules_sheet(
        wb, period, mode=mode, stat_objects=selected_stats if mode == MODE_STRATEGIC else None
    )
    _progress(91, "正在写入预处理表…")
    write_preprocessed_sheets(
        wb, orders, invoices, progress_callback=_progress
    )
    _progress(96, "正在保存文件…")

    wb.save(output_path)
    _progress(100, "导出完成")

    preprocess_paths = None
    if save_preprocess_files:
        preprocess_paths = save_preprocessed_sources(
            orders, invoices, output_path.parent
        )

    status_counts = dict(Counter(r["status"] for r in results))
    stats = {
        "mode": mode,
        "mode_label": MODE_LABELS.get(mode, mode),
        "stat_objects": selected_stats if mode == MODE_STRATEGIC else [],
        "strategic_model_count": len(strategic_model_keys),
        "order_rows_before_filter": order_rows_raw,
        "invoice_rows_before_filter": invoice_rows_raw,
        "order_rows": len(orders),
        "invoice_rows": len(invoices),
        "order_excluded_rows": len(excluded_orders),
        "invoice_excluded_rows": len(excluded_invoices),
        "result_rows": len(results),
        "status_counts": status_counts,
        "manual_count": sum(1 for r in results if r["need_manual"]),
        "multi_cand_count": sum(
            1
            for r in results
            if any(str(x).startswith("MULTI_CAND") for x in (r.get("reasons") or []))
        ),
        "price_over_count": sum(
            1 for r in results if r["status"] == STATUS_PRICE_OVER
        ),
        "price_under_count": sum(
            1 for r in results if r["status"] == STATUS_PRICE_UNDER
        ),
        "price_missing_count": sum(
            1 for r in results if r["status"] == STATUS_PRICE_MISSING
        ),
        "period_start": str(period.start),
        "period_end": str(period.end),
        "account_days": period.account_days,
        "grace_invoice_until": str(period.grace_invoice_until),
        "grace_order_from": str(period.grace_order_from),
        "period_inferred": period_inferred,
        "order_sheet": order_sheet,
        "invoice_sheet": inv_sheet,
        "output_path": str(output_path),
        "preprocess_paths": (
            [str(p) for p in preprocess_paths] if preprocess_paths else None
        ),
    }
    return stats


def parse_cli_date(text: str, label: str) -> date:
    s = (text or "").strip()
    if not s:
        raise ValueError(f"{label}不能为空")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"{label}格式无效，请使用 YYYY-MM-DD：{s}")


def prompt_with_default(label: str, default: str) -> str:
    raw = input(f"{label} [{default}]: ").strip()
    return raw if raw else default


def peek_period_from_invoice_file(invoice_path: Path, account_days: int) -> PeriodConfig:
    """读取发票文件并按开票日推断查询周期（不写结果）。"""
    _, inv_raw, _ = read_sheet_rows(
        invoice_path, required_headers=INVOICE_REQUIRED_HEADERS
    )
    if not inv_raw:
        raise ValueError("发票文件没有有效数据行，无法推断查询周期")
    invoices = preprocess_invoices(inv_raw)
    return infer_period_from_invoices(invoices, account_days)


def normalize_stat_object(v) -> str | None:
    """统计对象规范化：去首尾空白；空值返回 None。"""
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def collect_stat_objects_from_rows(rows) -> list[str]:
    """从订单行收集非空统计对象（去重排序）。"""
    seen: set[str] = set()
    for row in rows:
        s = normalize_stat_object(row.get("统计对象"))
        if s:
            seen.add(s)
    return sorted(seen)


def list_stat_objects(order_path) -> list[str]:
    """读取订单文件，返回可用于战略品监察的「统计对象」列表。"""
    path = Path(order_path)
    if not path.is_file():
        raise FileNotFoundError(f"订单文件不存在：{path}")
    _, order_raw, _ = read_sheet_rows(path, required_headers=ORDER_REQUIRED_HEADERS)
    if not order_raw:
        raise ValueError("订单文件没有有效数据行")
    values = collect_stat_objects_from_rows(order_raw)
    if not values:
        raise ValueError(
            "订单数据中未找到非空的「统计对象」列值，无法进行战略品监察。"
            "请确认表头含「统计对象」且至少有一行填写了类别。"
        )
    return values


def filter_for_strategic(orders, invoices, stat_objects: list[str]):
    """按所选统计对象过滤：仅保留对应订单型号及同型号发票；其余丢弃。"""
    selected = {
        normalize_stat_object(x)
        for x in (stat_objects or [])
        if normalize_stat_object(x)
    }
    if not selected:
        raise ValueError("战略品监察请至少选择一个统计对象")

    filtered_orders = [
        o for o in orders if normalize_stat_object(o.get("统计对象")) in selected
    ]
    model_keys = {o["_model_key"] for o in filtered_orders if o.get("_model_key")}
    if not model_keys:
        raise ValueError("所选统计对象下没有可比对的订单型号")

    filtered_invoices = [
        i for i in invoices if i.get("_model_key") in model_keys
    ]
    return filtered_orders, filtered_invoices, sorted(model_keys)


def resolve_period_interactively(
    invoice_path: Path,
    account_days: int,
    period_start=None,
    period_end=None,
    non_interactive: bool = False,
) -> tuple[date, date, int, bool]:
    """
    上传/选定发票后：按发票开票日推断周期，打印到 CLI，并允许手动修改。
    返回 (起始日, 结束日, 账期天数, 是否仍为推断未改周期边界)。
    """
    inferred = peek_period_from_invoice_file(invoice_path, account_days)
    start = to_date_only(period_start) or inferred.start
    end = to_date_only(period_end) or inferred.end
    days = int(account_days)

    print("")
    print("=== 发票查询周期 / 账期 ===")
    print(f"发票文件：{invoice_path}")
    print(
        f"已根据发票开票日推断查询周期（起始月月初～结束月月末）："
        f"{inferred.start} ~ {inferred.end}"
        f"（账期默认 {days} 天）"
    )
    preview = build_period_config(start, end, days)
    print(f"当前将使用：{preview.start} ~ {preview.end}，账期 {preview.account_days} 天")
    print(f"  期初豁免截止日：{preview.grace_invoice_until}")
    print(f"  期末豁免起始日：{preview.grace_order_from}")

    if non_interactive:
        still_inferred = (
            start == inferred.start and end == inferred.end and period_start is None and period_end is None
        )
        return start, end, days, still_inferred

    print("可直接回车保留当前值，或输入新值后回车：")
    while True:
        try:
            start_s = prompt_with_default("发票查询周期起始日", str(start))
            end_s = prompt_with_default("发票查询周期结束日", str(end))
            days_s = prompt_with_default("账期天数", str(days))
            start = parse_cli_date(start_s, "发票查询周期起始日")
            end = parse_cli_date(end_s, "发票查询周期结束日")
            days = int(days_s)
            if days < 0:
                raise ValueError("账期天数不能为负")
            cfg = build_period_config(start, end, days)
            break
        except (ValueError, TypeError) as exc:
            print(f"输入无效：{exc}，请重新输入。")

    print(
        f"已确认：查询周期 {cfg.start} ~ {cfg.end}，账期 {cfg.account_days} 天；"
        f"期初豁免截止 {cfg.grace_invoice_until}，期末豁免起始 {cfg.grace_order_from}"
    )
    still_inferred = start == inferred.start and end == inferred.end
    return start, end, days, still_inferred


def build_arg_parser():
    import argparse

    p = argparse.ArgumentParser(
        description=(
            "Order/invoice reconciliation. After loading the invoice file, "
            "the query period is inferred and can be confirmed or edited interactively."
        )
    )
    p.add_argument("--order", default=str(ORDER_FILE), help="Order Excel path")
    p.add_argument("--invoice", default=str(INVOICE_FILE), help="Invoice Excel path")
    p.add_argument("--output", default=str(OUTPUT_FILE), help="Result Excel path")
    p.add_argument(
        "--period-start",
        default=None,
        help="Query period start YYYY-MM-DD; omit to infer from invoice dates",
    )
    p.add_argument(
        "--period-end",
        default=None,
        help="Query period end YYYY-MM-DD; omit to infer from invoice dates",
    )
    p.add_argument(
        "--account-days",
        type=int,
        default=DEFAULT_ACCOUNT_DAYS,
        help=f"Account period in days (default {DEFAULT_ACCOUNT_DAYS})",
    )
    p.add_argument(
        "--non-interactive",
        action="store_true",
        help="Do not prompt; use inferred period or provided --period-start/--period-end",
    )
    p.add_argument(
        "--save-preprocess",
        action="store_true",
        help="Also write preprocessed order/invoice xlsx files",
    )
    p.add_argument(
        "--mode",
        choices=[MODE_DAILY, MODE_STRATEGIC],
        default=MODE_DAILY,
        help="Monitoring mode: daily (full) or strategic (filter by 统计对象)",
    )
    p.add_argument(
        "--stat-object",
        action="append",
        default=None,
        dest="stat_objects",
        help="Strategic mode: 统计对象 category (repeatable)",
    )
    return p


def main(argv=None):
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    order_path = Path(args.order)
    invoice_path = Path(args.invoice)
    output_path = Path(args.output)
    days = int(args.account_days)
    start_arg = args.period_start
    end_arg = args.period_end
    non_interactive = bool(args.non_interactive)

    if (start_arg and not end_arg) or (end_arg and not start_arg):
        parser.error(
            "--period-start and --period-end must be provided together, "
            "or both omitted (infer from invoice)"
        )
    if days < 0:
        parser.error("--account-days cannot be negative")

    if not invoice_path.is_file():
        raise SystemExit(f"发票文件不存在：{invoice_path}")
    if not order_path.is_file():
        raise SystemExit(f"订单文件不存在：{order_path}")

    period_start = parse_cli_date(start_arg, "发票查询周期起始日") if start_arg else None
    period_end = parse_cli_date(end_arg, "发票查询周期结束日") if end_arg else None

    period_start, period_end, days, still_inferred = resolve_period_interactively(
        invoice_path,
        days,
        period_start=period_start,
        period_end=period_end,
        non_interactive=non_interactive,
    )

    if args.mode == MODE_STRATEGIC and not args.stat_objects:
        parser.error("战略品监察(--mode strategic)需至少指定一个 --stat-object")

    stats = run_reconcile(
        order_path,
        invoice_path,
        output_path,
        save_preprocess_files=bool(args.save_preprocess),
        period_start=period_start,
        period_end=period_end,
        account_days=days,
        mode=args.mode,
        stat_objects=args.stat_objects,
    )
    stats["period_inferred"] = still_inferred

    print("")
    print("=== 比对完成 ===")
    print(f"监察方式: {stats.get('mode_label', stats.get('mode'))}")
    if stats.get("mode") == MODE_STRATEGIC:
        print(f"统计对象: {'、'.join(stats.get('stat_objects') or [])}")
        print(
            f"过滤前订单/发票: {stats.get('order_rows_before_filter')}/"
            f"{stats.get('invoice_rows_before_filter')} → "
            f"{stats['order_rows']}/{stats['invoice_rows']}"
        )
    print(
        f"订单行: {stats['order_rows']}  发票行: {stats['invoice_rows']}  "
        f"结果行: {stats['result_rows']}"
    )
    print(
        f"查询周期: {stats['period_start']} ~ {stats['period_end']}  "
        f"账期: {stats['account_days']}天"
        f"{'（发票推断）' if stats['period_inferred'] else ''}"
    )
    print(
        f"期初豁免截止: {stats['grace_invoice_until']}  "
        f"期末豁免起始: {stats['grace_order_from']}"
    )
    print(f"订单工作表: {stats['order_sheet']}  发票工作表: {stats['invoice_sheet']}")
    for k, v in sorted(stats["status_counts"].items()):
        print(f"  {k}: {v}")
    print(f"需人工核对: {stats['manual_count']}")
    print(f"含 MULTI_CAND: {stats['multi_cand_count']}")
    print(f"实际售价高于申请价格: {stats['price_over_count']}")
    print(f"实际售价低于申请价格: {stats['price_under_count']}")
    print(f"价格缺失: {stats['price_missing_count']}")
    print(f"输出: {stats['output_path']}")
    if stats["preprocess_paths"]:
        print(f"预处理: {', '.join(stats['preprocess_paths'])}")


if __name__ == "__main__":
    main()
