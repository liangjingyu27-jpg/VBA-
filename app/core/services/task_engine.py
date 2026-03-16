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
            year = int(v.year)
            month = int(v.month)
            if 1 <= month <= 12:
                return f"{year:04d}-{month:02d}"
            return ""
        except Exception:
            return ""

    t = s(v)
    if not t:
        return ""

    m = re.match(r"^(\d{4})(\d{2})$", t)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
        return ""

    m = re.match(r"^(\d{4})[-/](\d{1,2})$", t)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
        return ""

    try:
        dt = datetime.fromisoformat(t.replace("/", "-"))
        if 1 <= dt.month <= 12:
            return f"{dt.year:04d}-{dt.month:02d}"
        return ""
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
    if t.endswith("%"):
        t = t[:-1].strip()
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


def build_exclude_rule_tasks_from_raw_rows(
    raw_rows: list[dict[str, Any]],
    rule_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not raw_rows:
        return []

    run_month = detect_run_month_from_raw_rows(raw_rows)
    rules = load_rules_from_rows(rule_rows, run_month)

    tasks: list[dict[str, Any]] = []
    scope_map = {"运输线路": "运输线路", "运输方式": "运输方式", "收货渠道地址": "收货渠道地址", "备注": "备注"}
    for raw_index, row in enumerate(raw_rows, start=2):
        doc_no = s(pick_first(row, ["单据编号"]))
        if not doc_no:
            continue

        status = s(pick_first(row, ["单据状态"]))
        if norm(status) != "已审核":
            continue

        line = s(pick_first(row, ["运输线路"]))
        mode = s(pick_first(row, ["运输方式"]))
        addr = s(pick_first(row, ["收货渠道地址"]))
        remark = s(pick_first(row, ["备注"]))

        raw_scan_all_blank = (norm(line) == "" and norm(mode) == "" and norm(addr) == "" and norm(remark) == "")
        if raw_scan_all_blank:
            continue

        rule_action, hit_rule, hit_field = eval_rules(rules, line, mode, addr, remark)
        if rule_action != "待确认":
            continue

        tasks.append(
            {
                "task_id": f"ER-{raw_index}-{doc_no}",
                "raw_row": raw_index,
                "doc_no": doc_no,
                "customer_code": s(pick_first(row, ["客户编码"])),
                "customer_name": s(pick_first(row, ["收货客户"])),
                "hit_field": hit_field or "未命中字段",
                "hit_keyword": "" if hit_rule is None else hit_rule.keyword,
                "reason": "命中待确认规则，需维护候选规则",
                "line": line,
                "mode": mode,
                "addr": addr,
                "remark": remark,
                "candidate_scope": scope_map.get(hit_field, "全部"),
                "candidate_action": "排除",
                "candidate_keyword": "" if hit_rule is None else hit_rule.keyword,
                "candidate_remark": f"来源单据 {doc_no}，源行 {raw_index}",
                "saved": False,
                "status": "待处理",
            }
        )

    return tasks


def build_settlement_term_tasks_from_raw_rows(
    raw_rows: list[dict[str, Any]],
    term_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not raw_rows:
        return []

    run_month = detect_run_month_from_raw_rows(raw_rows)
    terms = load_terms_from_rows(term_rows, run_month)

    grouped: dict[str, dict[str, Any]] = {}
    row_level_tasks: list[dict[str, Any]] = []

    for raw_index, row in enumerate(raw_rows, start=2):
        doc_no = s(pick_first(row, ["单据编号"]))
        if not doc_no:
            continue

        status = s(pick_first(row, ["单据状态"]))
        if norm(status) != "已审核":
            continue

        customer_code = s(pick_first(row, ["客户编码"]))
        customer_name = s(pick_first(row, ["收货客户"]))
        has_term = customer_code != "" and customer_code in terms
        if has_term:
            continue

        if customer_code == "":
            row_level_tasks.append(
                {
                    "task_id": f"ST-ROW-{raw_index}-{doc_no}",
                    "raw_row": raw_index,
                    "doc_no": doc_no,
                    "customer_code": "[缺失客户编码]",
                    "customer_name": customer_name,
                    "reason": "源数据缺失客户编码，无法匹配结算条款",
                    "status": "待处理",
                }
            )
            continue

        group_key = customer_code
        existing = grouped.get(group_key)
        if existing is None:
            grouped[group_key] = {
                "task_id": f"ST-{run_month or 'NA'}-{customer_code}",
                "raw_row": raw_index,
                "doc_no": doc_no,
                "customer_code": customer_code,
                "customer_name": customer_name,
                "reason": "当前客户未匹配到有效结算条款",
                "status": "待处理",
                "raw_count": 1,
            }
        else:
            existing["raw_count"] = int(existing.get("raw_count", 1)) + 1
            if s(existing.get("customer_name")) == "" and customer_name != "":
                existing["customer_name"] = customer_name

    tasks = list(grouped.values()) + row_level_tasks
    tasks.sort(key=lambda x: (str(x.get("customer_code") == "[缺失客户编码]"), str(x.get("customer_code")), str(x.get("task_id"))))
    return tasks


def summarize_settlement_term_tasks(tasks: list[dict[str, Any]]) -> dict[str, int]:
    total = len(tasks)
    done = sum(1 for task in tasks if str(task.get("status")) == "已处理")
    pending = total - done
    return {
        "total": total,
        "done": done,
        "pending": pending,
    }


def can_mark_settlement_term_done(task: dict[str, Any], payload_written: bool) -> tuple[bool, str]:
    payload = task.get("maintain_payload")
    if not isinstance(payload, dict):
        return False, "当前客户尚未保存结算条款填写内容。"
    if not payload_written:
        return False, "当前维护值未写入结果层，不能标记已处理。"

    def get_value(en_key: str, zh_key: str) -> Any:
        return payload.get(en_key) if en_key in payload else payload.get(zh_key)

    required = [
        ("enabled", "是否启用"),
        ("customer_code", "客户编码"),
        ("settle_type", "结算类型"),
        ("start_month", "生效月份"),
    ]
    for en_key, zh_key in required:
        if s(get_value(en_key, zh_key)) == "":
            return False, f"字段“{zh_key}”缺失，不能标记已处理。"

    settle_type = s(get_value("settle_type", "结算类型"))
    if settle_type == "费比" and num(get_value("ratio", "费比")) is None:
        return False, "结算类型为费比时，必须填写费比。"
    if settle_type == "固定金额" and num(get_value("fixed_amount", "固定金额")) is None:
        return False, "结算类型为固定金额时，必须填写固定金额。"
    if settle_type == "不计费" and s(get_value("no_charge_reason", "不计费原因")) == "":
        return False, "结算类型为不计费时，必须填写不计费原因。"

    return True, "OK"
