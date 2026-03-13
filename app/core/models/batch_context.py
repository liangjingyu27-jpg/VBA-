from dataclasses import dataclass, field
from typing import Any


@dataclass
class ActualFreightTask:
    month: str = ""
    customer_code: str = ""
    customer_name: str = ""
    monthly_sales_audited: float = 0.0

    actual_freight_amount: float | None = None
    basis: str = ""
    basis_source: str = ""

    status: str = "待录入"
    raw_count: int = 0
    trace_codes: list[str] = field(default_factory=list)
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class BatchContext:
    batch_no: str = "未创建批次"
    file_path: str = ""
    file_name: str = "尚未导入"
    source_sheet: str = "尚未识别"
    source_reason: str = "尚未识别"

    available_sheet_names: list[str] = field(default_factory=list)
    source_rows: int = 0
    source_cols: int = 0
    preview_rows: list[list[str]] = field(default_factory=list)

    batch_state: str = "未导入"
    gate_state: str = "未通过"
    last_run_summary: str = "尚未运行"
    feedback: str = "当前提示：请先导入原始数据文件。"

    actual_freight_tasks: list[ActualFreightTask] = field(default_factory=list)

    @property
    def actual_freight_count(self) -> int:
        return len(self.actual_freight_tasks)