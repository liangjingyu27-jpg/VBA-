from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Optional

from app.core.models.batch_context import ActualFreightTask

ADDR_MISSING = "[地址缺失]"
REMARK_MISSING = "[备注缺失]"


@dataclass
class Rule:
    row: int
    keyword: str
    scope: str
    action: str
    priority: int
    start_month: str
    end_month: str


@dataclass
class Term:
    row: int
    code: str
    name: str
    settle_type: str
    ratio: Optional[float]
    fixed: Optional[float]
    no_charge_reason: str
    remark: str
    start_month: str
    end_month: str


@dataclass
class SelfPickup:
    decision: str
    row: int


@dataclass
class ActualFreightRecord:
    amount: Optional[float]
    basis: str
    source: str
    row: int


def s(v: Any) -> str:
    if v is None:
        return ""
    t = str(v).strip()
    if t.lower() in {"nan", "none"}:
        return ""
    return t


def norm(v: Any) -> str:
    t = s(v)
    if not t:
        return ""
    t = re.sub(r"\s+", " ", t)
    return t.lower()


def contains(hay: Any, needle: str) -> bool:
    if not needle:
        return False
    return needle.lower() in norm(hay)


def to_month(v: Any) -> str:
    if v is None:
        return ""
    if hasattr(v, "year") and hasattr(v, "month"):
        try:
            return f"{int(v.year):04d}-{int(v.month):02d}"
        except Exception:
            pass
    t = s(v)
    if not t:
        return ""
    m = re.match(r"^\s*(\d{4})[-/](\d{1,2})", t)
    if m:
        y = int(m.group(1))
        mm = int(m.group(2))
        return f"{y:04d}-{mm:02d}"
    try:
        dt = datetime.fromisoformat(t.replace("/", "-"))
        return f"{dt.year:04d}-{dt.month:02d}"
    except Exception:
        return ""


def month_key(m: str) -> int:
    if not m:
        return -1
    mm = re.match(r"^(\d{4})-(\d{2})$", m)
    if not mm:
        return -1
    y = int(mm.group(1))
    mo = int(mm.group(2))
    return y * 12 + mo


def num(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = s(v).replace(",", "")
    if t == "":
        return None
    try:
        return float(t)
    except Exception:
        return None


def pick_first(row: dict[str, Any], candidates: list[str]) -> Any:
    for name in candidates:
        if name in row:
            return row.get(name)
    return None


def detect_run_month_from_raw_rows(raw_rows: list[dict[str, Any]]) -> str:
    for row in raw_rows:
        m = to_month(pick_first(row, ["业务日期", "日期", "月份"]))
        if m:
            return m
    return ""


def load_rules_from_rows(rule_rows: list[dict[str, Any]], run_month: str) -> list[Rule]:
    run_k = month_key(run_month)
    rules: list[Rule] = []
    action_seen = {"保留": 0, "排除": 0, "待确认": 0}

    for i, row in enumerate(rule_rows, start=3):
        enabled = norm(pick_first(row, ["是否启用", "启用", "启用(Y/N)"])) == "y"
        keyword = s(pick_first(row, ["关键词", "关键字"]))
        if not (enabled and keyword):
            continue

        scope = s(pick_first(row, ["范围", "作用范围", "字段范围", "扫描范围"])) or "全部"
        action = s(pick_first(row, ["动作", "处理动作"])) or "待确认"
        start_m = to_month(pick_first(row, ["生效月份", "开始月份", "起始月份"]))
        end_m = to_month(pick_first(row, ["失效月份", "失效月份(可空)", "结束月份"])) or "9999-99"

        if run_k != -1:
            s_k = month_key(start_m) if start_m else -1
            e_k = month_key(end_m) if end_m else 9999 * 12 + 99
            if s_k != -1 and s_k > run_k:
                continue
            if e_k != -1 and e_k < run_k:
                continue

        base = 900
        if action == "保留":
            base = 100
        elif action == "排除":
            base = 200
        elif action == "待确认":
            base = 300

        pri = base - len(keyword) - (0 if scope == "全部" else 10) + action_seen.get(action, 0)
        action_seen[action] = action_seen.get(action, 0) + 1

        rules.append(
            Rule(
                row=i,
                keyword=keyword,
                scope=scope,
                action=action,
                priority=int(pri),
                start_month=start_m,
                end_month=end_m,
            )
        )

    rules.sort(key=lambda x: (x.priority, x.row))
    return rules


def load_terms_from_rows(term_rows: list[dict[str, Any]], run_month: str) -> dict[str, Term]:
    run_k = month_key(run_month)
    best: dict[str, Term] = {}

    for i, row in enumerate(term_rows, start=3):
        enabled = norm(pick_first(row, ["是否启用", "启用", "启用(Y/N)"])) == "y"
        code = s(pick_first(row, ["客户编码"]))
        if not (enabled and code):
            continue

        name = s(pick_first(row, ["客户名称(可选)", "客户名称"]))
        settle_type = s(pick_first(row, ["结算类型"]))
        ratio = num(pick_first(row, ["费比(如7.9%)", "费比"]))
        fixed = num(pick_first(row, ["固定金额(可空)", "固定金额"]))
        no_charge_reason = s(pick_first(row, ["不计费原因(可空)", "不计费原因"]))
        remark = s(pick_first(row, ["备注"]))
        start_m = to_month(pick_first(row, ["生效月份", "开始月份"]))
        end_m = to_month(pick_first(row, ["失效月份(可空)", "失效月份"])) or "9999-99"

        if run_k != -1:
            s_k = month_key(start_m) if start_m else -1
            e_k = month_key(end_m) if end_m else 9999 * 12 + 99
            if s_k != -1 and s_k > run_k:
                continue
            if e_k != -1 and e_k < run_k:
                continue

        t = Term(
            row=i,
            code=code,
            name=name,
            settle_type=settle_type,
            ratio=ratio,
            fixed=fixed,
            no_charge_reason=no_charge_reason,
            remark=remark,
            start_month=start_m,
            end_month=end_m,
        )
        cur = best.get(code)
        if cur is None:
            best[code] = t
        else:
            if month_key(t.start_month) > month_key(cur.start_month):
                best[code] = t

    return best


def load_self_pickup_from_rows(sp_rows: list[dict[str, Any]], run_month: str) -> dict[tuple[str, str, str, str, str], SelfPickup]:
    mp: dict[tuple[str, str, str, str, str], SelfPickup] = {}
    for i, row in enumerate(sp_rows, start=5):
        month = to_month(pick_first(row, ["月份"]))
        cust = s(pick_first(row, ["收货客户"]))
        line = s(pick_first(row, ["运输线路"]))
        mode = s(pick_first(row, ["运输方式"]))
        addr = s(pick_first(row, ["收货渠道地址"])) or ADDR_MISSING
        remark = s(pick_first(row, ["备注"])) or REMARK_MISSING
        decision = s(pick_first(row, ["判定(计费/不计费)", "判定"]))
        if not cust:
            continue
        if decision not in ("计费", "不计费"):
            continue
        if month and run_month and month != run_month:
            continue
        mp[(norm(cust), norm(line), norm(mode), norm(addr), norm(remark))] = SelfPickup(
            decision=decision,
            row=i,
        )
    return mp


def load_actual_freight_from_rows(af_rows: list[dict[str, Any]], run_month: str) -> dict[tuple[str, str], ActualFreightRecord]:
    mp: dict[tuple[str, str], ActualFreightRecord] = {}
    for i, row in enumerate(af_rows, start=5):
        code = s(pick_first(row, ["客户编码"]))
        if not code:
            continue
        month = to_month(pick_first(row, ["月份"])) or run_month
        amt = num(pick_first(row, ["实际运费金额", "实际结算金额", "实际金额"]))
        basis = s(pick_first(row, ["测算依据", "依据"]))
        source = s(pick_first(row, ["依据来源"]))
        mp[(month, code)] = ActualFreightRecord(amount=amt, basis=basis, source=source, row=i)
    return mp


FIELD_ORDER = [("运输线路", "line"), ("运输方式", "mode"), ("收货渠道地址", "addr"), ("备注", "remark")]


def scope_allows(scope: str, field_cn: str) -> bool:
    return (not scope) or scope == "全部" or scope == field_cn


def eval_rules(rules: list[Rule], line: str, mode: str, addr: str, remark: str) -> tuple[str, Optional[Rule], str]:
    vals = {"line": line, "mode": mode, "addr": addr, "remark": remark}
    for field_cn, key in FIELD_ORDER:
        v = vals[key]
        matched: list[Rule] = []
        for ru in rules:
            if not scope_allows(ru.scope, field_cn):
                continue
            if contains(v, ru.keyword):
                matched.append(ru)
        if not matched:
            continue

        for act in ("保留", "排除", "待确认"):
            cand = [x for x in matched if x.action == act]
            if cand:
                cand.sort(key=lambda x: (x.priority, x.row))
                return act, cand[0], field_cn

        matched.sort(key=lambda x: (x.priority, x.row))
        return matched[0].action, matched[0], field_cn

    return "待确认", None, ""


def build_actual_freight_tasks_from_raw_rows(
    raw_rows: list[dict[str, Any]],
    rule_rows: list[dict[str, Any]],
    term_rows: list[dict[str, Any]],
    self_pickup_rows: list[dict[str, Any]] | None = None,
    actual_freight_rows: list[dict[str, Any]] | None = None,
) -> list[ActualFreightTask]:
    if not raw_rows:
        return []

    run_month = detect_run_month_from_raw_rows(raw_rows)
    rules = load_rules_from_rows(rule_rows, run_month)
    terms = load_terms_from_rows(term_rows, run_month)
    sp_map = load_self_pickup_from_rows(self_pickup_rows or [], run_month)
    af_map = load_actual_freight_from_rows(actual_freight_rows or [], run_month)

    a3_rows: list[dict[str, Any]] = []

    for raw_index, row in enumerate(raw_rows, start=2):
        doc_no = s(pick_first(row, ["单据编号"]))
        if not doc_no:
            continue

        status = s(pick_first(row, ["单据状态"]))
        cust_code = s(pick_first(row, ["客户编码"]))
        cust_name = s(pick_first(row, ["收货客户"]))
        line = s(pick_first(row, ["运输线路"]))
        mode = s(pick_first(row, ["运输方式"]))
        addr = s(pick_first(row, ["收货渠道地址"]))
        remark = s(pick_first(row, ["备注"]))
        qty = num(pick_first(row, ["数量"]))
        amt = num(pick_first(row, ["价税合计"]))

        audited = norm(status) == "已审核"
        if not audited:
            continue

        amount_missing = (amt is None) and (qty is not None and abs(qty) > 1e-9)
        if amount_missing:
            # D1 风险，不进入 A3
            continue

        rule_action, _, _ = eval_rules(rules, line, mode, addr, remark)

        raw_scan_all_blank = (norm(line) == "" and norm(mode) == "" and norm(addr) == "" and norm(remark) == "")
        if raw_scan_all_blank:
            rule_action = "保留"

        if rule_action == "排除":
            continue

        if rule_action == "待确认":
            # U1 / R1，先不进入 A3
            continue

        self_pickup_trigger = any(contains(x, "自提") for x in (line, mode, addr, remark))
        sp_decision = ""
        if self_pickup_trigger:
            addr_k = addr if addr else ADDR_MISSING
            remark_k = remark if remark else REMARK_MISSING
            sp = sp_map.get((norm(cust_name), norm(line), norm(mode), norm(addr_k), norm(remark_k)))
            if sp:
                sp_decision = sp.decision

            if sp_decision == "":
                # B1，先不进入 A3
                continue
            if sp_decision == "不计费":
                continue
            # 计费才继续

        term = terms.get(cust_code)
        if term is None:
            # C1，先不进入 A3
            continue

        settle_type = term.settle_type
        if settle_type != "按实际":
            continue

        a3_rows.append(
            {
                "月份": run_month,
                "客户编码": cust_code,
                "收货客户": cust_name,
                "价税合计": 0.0 if amt is None else float(amt),
                "源行号": raw_index,
            }
        )

    if not a3_rows:
        return []

    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in a3_rows:
        customer_code = s(row.get("客户编码"))
        if customer_code == "":
            continue
        grouped.setdefault(customer_code, []).append(row)

    tasks: list[ActualFreightTask] = []
    for customer_code, rows in grouped.items():
        names = [s(r.get("收货客户")) for r in rows if s(r.get("收货客户"))]
        customer_name = Counter(names).most_common(1)[0][0] if names else ""
        monthly_sales = sum(float(r.get("价税合计", 0.0) or 0.0) for r in rows)
        month = s(rows[0].get("月份"))

        record = af_map.get((month, customer_code))
        is_done = record is not None and record.amount is not None and s(record.basis) != ""

        tasks.append(
            ActualFreightTask(
                month=month,
                customer_code=customer_code,
                customer_name=customer_name,
                monthly_sales_audited=round(monthly_sales, 2),
                actual_freight_amount=None if record is None else record.amount,
                basis="" if record is None else record.basis,
                basis_source="" if record is None else record.source,
                status="已录入" if is_done else "待录入",
                raw_count=len(rows),
                trace_codes=["A3"],
                extra={"rows": rows},
            )
        )

    # 待录入排前面，已录入排后面
    tasks.sort(key=lambda item: (item.status == "已录入", item.customer_name, item.customer_code))
    return tasks


def summarize_actual_freight_tasks(tasks: list[ActualFreightTask]) -> dict[str, int]:
    total = len(tasks)
    done = sum(1 for x in tasks if x.status == "已录入")
    pending = sum(1 for x in tasks if x.status != "已录入")
    return {
        "total": total,
        "done": done,
        "pending": pending,
    }