import os
import sys
from datetime import datetime

from PySide6.QtCore import QPoint, QRect, QSize, Qt
from PySide6.QtGui import QResizeEvent
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLayout,
    QLayoutItem,
    QMainWindow,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.core.services.excel_service import (
    close_workbook_safely,
    find_sheet_by_keywords,
    open_excel_with_source,
    sheet_to_dict_rows,
)
from app.core.services.task_engine import (
    build_actual_freight_tasks_from_raw_rows,
    summarize_actual_freight_tasks,
)


class FlowLayout(QLayout):
    def __init__(self, parent: QWidget | None = None, margin: int = 0, hspacing: int = 8, vspacing: int = 8) -> None:
        super().__init__(parent)
        self.item_list: list[QLayoutItem] = []
        self._hspacing = hspacing
        self._vspacing = vspacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item: QLayoutItem) -> None:
        self.item_list.append(item)

    def count(self) -> int:
        return len(self.item_list)

    def itemAt(self, index: int) -> QLayoutItem | None:
        if 0 <= index < len(self.item_list):
            return self.item_list[index]
        return None

    def takeAt(self, index: int) -> QLayoutItem | None:
        if 0 <= index < len(self.item_list):
            return self.item_list.pop(index)
        return None

    def expandingDirections(self) -> Qt.Orientations:
        return Qt.Orientation(0)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return self._do_layout(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect: QRect) -> None:
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self) -> QSize:
        return self.minimumSize()

    def minimumSize(self) -> QSize:
        size = QSize()
        for item in self.item_list:
            size = size.expandedTo(item.minimumSize())
        margins = self.contentsMargins()
        size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        return size

    def _smart_spacing(self, pm: QStyle.PixelMetric) -> int:
        parent = self.parent()
        if parent is None:
            return -1
        if parent.isWidgetType():
            return parent.style().pixelMetric(pm, None, parent)
        return parent.spacing()

    def horizontalSpacing(self) -> int:
        if self._hspacing >= 0:
            return self._hspacing
        return self._smart_spacing(QStyle.PixelMetric.PM_LayoutHorizontalSpacing)

    def verticalSpacing(self) -> int:
        if self._vspacing >= 0:
            return self._vspacing
        return self._smart_spacing(QStyle.PixelMetric.PM_LayoutVerticalSpacing)

    def _do_layout(self, rect: QRect, test_only: bool) -> int:
        left, top, right, bottom = self.getContentsMargins()
        effective = rect.adjusted(left, top, -right, -bottom)
        x = effective.x()
        y = effective.y()
        line_height = 0
        spacing_x = self.horizontalSpacing()
        spacing_y = self.verticalSpacing()

        for item in self.item_list:
            widget = item.widget()
            if widget is not None and not widget.isVisible():
                continue

            hint = item.sizeHint()
            next_x = x + hint.width() + spacing_x
            if next_x - spacing_x > effective.right() and line_height > 0:
                x = effective.x()
                y = y + line_height + spacing_y
                next_x = x + hint.width() + spacing_x
                line_height = 0

            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))

            x = next_x
            line_height = max(line_height, hint.height())

        return y + line_height - rect.y() + bottom


class ResponsivePanelRow(QFrame):
    def __init__(self) -> None:
        super().__init__()
        self._panels: list[QWidget] = []
        self._grid = QGridLayout(self)
        self._grid.setContentsMargins(0, 0, 0, 0)
        self._grid.setHorizontalSpacing(10)
        self._grid.setVerticalSpacing(10)

    def add_panel(self, panel: QWidget) -> None:
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        panel.setMinimumWidth(320)
        self._panels.append(panel)
        self._rebuild()

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._rebuild()

    def _rebuild(self) -> None:
        while self._grid.count():
            item = self._grid.takeAt(0)
            if item is not None and item.widget() is not None:
                item.widget().setParent(self)

        width = max(1, self.width())
        if width >= 1320:
            columns = 3
        elif width >= 900:
            columns = 2
        else:
            columns = 1

        for index, panel in enumerate(self._panels):
            row = index // columns
            col = index % columns
            self._grid.addWidget(panel, row, col)

        for col in range(columns):
            self._grid.setColumnStretch(col, 1)


class InfoBlock(QFrame):
    def __init__(self, title: str, value: str = "") -> None:
        super().__init__()
        self.setObjectName("InfoBlock")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(2)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("InfoBlockTitle")

        self.value_label = QLabel(value)
        self.value_label.setObjectName("InfoBlockValue")
        self.value_label.setWordWrap(True)

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("昆分运费账单工作台 V3")
        self.resize(1440, 960)

        self.current_batch_no = "未创建批次"
        self.current_file_path = ""
        self.current_file_name = "尚未导入"
        self.current_source_sheet = "尚未识别"
        self.current_source_reason = "尚未识别"
        self.available_sheet_names: list[str] = []
        self.source_preview_rows: list[list[str]] = []
        self.source_rows = 0
        self.source_cols = 0

        self.current_batch_state = "未导入"
        self.gate_state = "未通过"
        self.last_run_summary = "尚未运行"
        self.current_feedback = "当前提示：请先导入原始数据文件。"
        self.total_progress_value = 0
        self.current_stage_key = "overview"

        self.task_pool = {
            "exclude_rules": 0,
            "settlement_terms": 0,
            "self_pickup": 0,
            "actual_freight": 0,
        }

        self.actual_freight_tasks = []
        self.actual_freight_edited_rows: set[int] = set()
        self._is_refreshing_actual_freight_table = False

        self.batch_states = [
            "已导入",
            "已运行识别",
            "待维护",
            "维护中",
            "可再运行",
            "可导出",
            "已导出",
        ]

        self.task_meta = {
            "overview": {"title": "批次总览", "counted": False},
            "exclude_rules": {"title": "排除规则", "counted": True},
            "settlement_terms": {"title": "结算条款", "counted": True},
            "self_pickup": {"title": "自提判定", "counted": True},
            "actual_freight": {"title": "按实际运费", "counted": True},
            "run_check": {"title": "运行与校验", "counted": False},
            "export": {"title": "导出结果", "counted": False},
        }

        self._apply_style()
        self._build_ui()
        self._refresh_ui()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: #f7f8fc;
            }
            QWidget {
                font-size: 14px;
                color: #2f3441;
            }
            QFrame#PanelShell, QFrame#InfoBlock, QFrame#TaskCardShell {
                background: #ffffff;
                border: 1px solid #e6ebf3;
                border-radius: 14px;
            }
            QLabel#AppTitle {
                font-size: 24px;
                font-weight: 700;
                color: #1e2430;
            }
            QLabel#PanelTitle {
                font-size: 15px;
                font-weight: 700;
                color: #1e2430;
            }
            QLabel#SectionNote {
                color: #6b7280;
                font-size: 12px;
            }
            QLabel#InfoBlockTitle {
                color: #6b7280;
                font-size: 11px;
            }
            QLabel#InfoBlockValue {
                font-size: 15px;
                font-weight: 700;
                color: #1f2430;
            }
            QLabel#TrackText {
                font-size: 12px;
                font-weight: 700;
                color: #4a5565;
            }
            QLabel#TaskCardTitle {
                font-size: 18px;
                font-weight: 700;
                color: #1f2430;
            }
            QLabel#TaskCardBody {
                font-size: 13px;
                color: #425066;
            }
            QPushButton {
                min-height: 38px;
                border-radius: 10px;
                padding: 0 14px;
                border: 1px solid #d8deea;
                background: #ffffff;
                font-weight: 700;
            }
            QPushButton:hover {
                background: #f8faff;
            }
            QPushButton:disabled {
                color: #a2aaba;
                background: #f2f4f8;
                border: 1px solid #e4e8f0;
            }
            QPushButton#PrimaryButton {
                background: #2f6fed;
                color: #ffffff;
                border: none;
            }
            QPushButton#PrimaryButton:disabled {
                background: #b8caee;
                color: #eef3ff;
            }
            QPushButton#TaskTabButton {
                text-align: center;
                min-height: 40px;
                min-width: 150px;
            }
            QPushButton#TaskTabButton:checked {
                background: #eef4ff;
                border: 1px solid #7ca5ff;
                color: #1e4ea8;
            }
            QProgressBar {
                min-height: 16px;
                border-radius: 8px;
                background: #eef2f8;
                border: 1px solid #e1e6ef;
                text-align: center;
                font-weight: 700;
                color: #314155;
            }
            QProgressBar::chunk {
                border-radius: 7px;
                background: #82aaf8;
            }
            QTableWidget {
                border: 1px solid #e8edf5;
                border-radius: 10px;
                background: #ffffff;
                gridline-color: #eff3f8;
            }
            QHeaderView::section {
                background: #f6f8fc;
                border: none;
                border-bottom: 1px solid #e8edf5;
                padding: 8px;
                font-weight: 700;
                color: #566172;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            """
        )

    def _build_ui(self) -> None:
        page_shell = QWidget()
        outer = QVBoxLayout(page_shell)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(10)

        self.top_status_bar = self._build_top_status_bar()
        self.status_and_task_bar = self._build_status_and_task_bar()
        self.current_task_card = self._build_current_task_card()
        self.main_stage_stack = self._build_main_stage_stack()
        self.bottom_feedback_bar = self._build_bottom_feedback_bar()

        outer.addWidget(self.top_status_bar)
        outer.addWidget(self.status_and_task_bar)
        outer.addWidget(self.current_task_card)
        outer.addWidget(self.main_stage_stack, 1)
        outer.addWidget(self.bottom_feedback_bar)

        self.setCentralWidget(page_shell)

    def _build_top_status_bar(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        app_title = QLabel("昆分运费账单工作台 V3")
        app_title.setObjectName("AppTitle")
        title_row.addWidget(app_title)
        title_row.addStretch()
        layout.addLayout(title_row)

        self.top_panel_row = ResponsivePanelRow()

        self.top_identity_panel = QFrame()
        self.top_identity_panel.setObjectName("PanelShell")
        left_layout = QVBoxLayout(self.top_identity_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(8)
        left_title = QLabel("身份区")
        left_title.setObjectName("PanelTitle")
        left_layout.addWidget(left_title)

        identity_grid = QGridLayout()
        identity_grid.setHorizontalSpacing(8)
        identity_grid.setVerticalSpacing(8)
        self.info_batch = InfoBlock("当前批次号", self.current_batch_no)
        self.info_file = InfoBlock("当前文件名", self.current_file_name)
        self.info_source = InfoBlock("当前源表", self.current_source_sheet)
        self.info_source_reason = InfoBlock("命中原因", self.current_source_reason)
        identity_grid.addWidget(self.info_batch, 0, 0)
        identity_grid.addWidget(self.info_file, 0, 1)
        identity_grid.addWidget(self.info_source, 1, 0)
        identity_grid.addWidget(self.info_source_reason, 1, 1)
        identity_grid.setColumnStretch(0, 1)
        identity_grid.setColumnStretch(1, 1)
        left_layout.addLayout(identity_grid)

        self.top_batch_state_panel = QFrame()
        self.top_batch_state_panel.setObjectName("PanelShell")
        mid_layout = QVBoxLayout(self.top_batch_state_panel)
        mid_layout.setContentsMargins(10, 10, 10, 10)
        mid_layout.setSpacing(8)
        mid_title = QLabel("批次状态区")
        mid_title.setObjectName("PanelTitle")
        mid_layout.addWidget(mid_title)

        status_grid = QGridLayout()
        status_grid.setHorizontalSpacing(8)
        status_grid.setVerticalSpacing(8)
        self.info_batch_state = InfoBlock("当前批次状态", self.current_batch_state)
        self.info_gate_state = InfoBlock("当前门禁状态", self.gate_state)
        self.info_last_run = InfoBlock("最近一次运行结果", self.last_run_summary)
        status_grid.addWidget(self.info_batch_state, 0, 0)
        status_grid.addWidget(self.info_gate_state, 0, 1)
        status_grid.addWidget(self.info_last_run, 1, 0, 1, 2)
        status_grid.setColumnStretch(0, 1)
        status_grid.setColumnStretch(1, 1)
        mid_layout.addLayout(status_grid)

        self.top_progress_panel = QFrame()
        self.top_progress_panel.setObjectName("PanelShell")
        right_layout = QVBoxLayout(self.top_progress_panel)
        right_layout.setContentsMargins(10, 10, 10, 10)
        right_layout.setSpacing(8)
        right_title = QLabel("进度区")
        right_title.setObjectName("PanelTitle")
        right_layout.addWidget(right_title)

        progress_grid = QGridLayout()
        progress_grid.setHorizontalSpacing(8)
        progress_grid.setVerticalSpacing(8)
        self.progress_current_state = InfoBlock("当前工作焦点", "尚未进入处理")
        self.progress_task_pool = InfoBlock("当前任务池摘要", "当前待处理：0 类 / 0 条")
        self.progress_ratio = InfoBlock("完成度数字", "0 / 7")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        progress_grid.addWidget(self.progress_current_state, 0, 0, 1, 2)
        progress_grid.addWidget(self.progress_task_pool, 1, 0)
        progress_grid.addWidget(self.progress_ratio, 1, 1)
        progress_grid.addWidget(self.progress_bar, 2, 0, 1, 2)
        progress_grid.setColumnStretch(0, 1)
        progress_grid.setColumnStretch(1, 1)
        right_layout.addLayout(progress_grid)

        self.top_panel_row.add_panel(self.top_identity_panel)
        self.top_panel_row.add_panel(self.top_batch_state_panel)
        self.top_panel_row.add_panel(self.top_progress_panel)
        layout.addWidget(self.top_panel_row)
        return shell

    def _build_status_and_task_bar(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        state_track_box = QFrame()
        state_track_box.setObjectName("PanelShell")
        state_layout = QVBoxLayout(state_track_box)
        state_layout.setContentsMargins(10, 10, 10, 10)
        state_layout.setSpacing(8)
        state_title = QLabel("批次状态轨")
        state_title.setObjectName("PanelTitle")
        state_layout.addWidget(state_title)

        self.batch_track_flow = FlowLayout(hspacing=8, vspacing=8)
        self.batch_state_track_items: dict[str, QLabel] = {}
        for state in self.batch_states:
            item = QLabel(state)
            item.setObjectName("TrackText")
            item.setAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setMinimumSize(110, 32)
            item.setStyleSheet(
                "background:#f6f8fc;border:1px solid #e4e9f2;border-radius:10px;padding:4px 8px;"
            )
            self.batch_state_track_items[state] = item
            self.batch_track_flow.addWidget(item)
        state_layout.addLayout(self.batch_track_flow)

        tabs_box = QFrame()
        tabs_box.setObjectName("PanelShell")
        tabs_layout = QVBoxLayout(tabs_box)
        tabs_layout.setContentsMargins(10, 10, 10, 10)
        tabs_layout.setSpacing(8)
        tabs_title = QLabel("任务分类签")
        tabs_title.setObjectName("PanelTitle")
        tabs_layout.addWidget(tabs_title)

        self.tabs_flow = FlowLayout(hspacing=8, vspacing=8)
        self.task_category_buttons: dict[str, QPushButton] = {}
        for key in self.task_meta.keys():
            btn = QPushButton(self.task_meta[key]["title"])
            btn.setObjectName("TaskTabButton")
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked=False, current_key=key: self._switch_stage(current_key))
            self.task_category_buttons[key] = btn
            self.tabs_flow.addWidget(btn)
        tabs_layout.addLayout(self.tabs_flow)

        layout.addWidget(state_track_box)
        layout.addWidget(tabs_box)
        return shell

    def _build_current_task_card(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("TaskCardShell")
        layout = QGridLayout(shell)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(6)

        self.current_task_title = QLabel("当前推荐任务：先导入原始数据")
        self.current_task_title.setObjectName("TaskCardTitle")
        layout.addWidget(self.current_task_title, 0, 0)

        self.current_task_remaining = QLabel("剩余：1 项")
        self.current_task_remaining.setObjectName("TaskCardBody")
        self.current_task_remaining.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.current_task_remaining, 0, 1)

        self.current_task_reason = QLabel("推荐原因：当前还没有建立批次，无法进入后续状态驱动处理。")
        self.current_task_reason.setObjectName("TaskCardBody")
        self.current_task_reason.setWordWrap(True)
        layout.addWidget(self.current_task_reason, 1, 0, 1, 2)

        self.current_task_meta = QLabel("当前影响：所有任务分类都尚未进入。｜完成后：进入批次总览并开始识别当前批次。")
        self.current_task_meta.setObjectName("TaskCardBody")
        self.current_task_meta.setWordWrap(True)
        layout.addWidget(self.current_task_meta, 2, 0)

        self.btn_task_go = QPushButton("打开当前推荐任务")
        self.btn_task_go.clicked.connect(self._open_recommended_task)
        layout.addWidget(self.btn_task_go, 2, 1)

        layout.setColumnStretch(0, 4)
        layout.setColumnStretch(1, 1)
        return shell

    def _wrap_scroll_stage(self, content: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(content)
        return scroll

    def _build_main_stage_stack(self) -> QWidget:
        self.stage_stack = QStackedWidget()
        self.stage_stack.setObjectName("main_stage_stack")
        self.stage_stack.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.stage_index: dict[str, int] = {}

        self.stage_index["overview"] = self.stage_stack.addWidget(self._wrap_scroll_stage(self._build_overview_stage()))
        self.stage_index["exclude_rules"] = self.stage_stack.addWidget(
            self._wrap_scroll_stage(
                self._build_task_stage(
                    stage_key="exclude_rules",
                    stage_title="排除规则维护面板",
                    stage_note="这里承载当前批次的排除规则候选维护。当前仍为演示占位。",
                )
            )
        )
        self.stage_index["settlement_terms"] = self.stage_stack.addWidget(
            self._wrap_scroll_stage(
                self._build_task_stage(
                    stage_key="settlement_terms",
                    stage_title="结算条款维护面板",
                    stage_note="这里承载当前批次的结算条款待补维护。当前仍为演示占位。",
                )
            )
        )
        self.stage_index["self_pickup"] = self.stage_stack.addWidget(
            self._wrap_scroll_stage(
                self._build_task_stage(
                    stage_key="self_pickup",
                    stage_title="自提判定维护面板",
                    stage_note="这里承载当前批次的自提待判定维护。当前仍为演示占位。",
                )
            )
        )
        self.stage_index["actual_freight"] = self.stage_stack.addWidget(
            self._wrap_scroll_stage(
                self._build_task_stage(
                    stage_key="actual_freight",
                    stage_title="按实际运费维护面板",
                    stage_note="这里承载当前批次的真实 A3 按实际录入任务（已录入 / 待录入联动）。",
                )
            )
        )
        self.stage_index["run_check"] = self.stage_stack.addWidget(self._wrap_scroll_stage(self._build_run_check_stage()))
        self.stage_index["export"] = self.stage_stack.addWidget(self._wrap_scroll_stage(self._build_export_stage()))
        return self.stage_stack

    def _build_overview_stage(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        shell.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title = QLabel("批次总览面板")
        title.setObjectName("PanelTitle")
        note = QLabel("这是默认入口，当前阶段先承载真实导入、真实源表识别，以及新骨架下的批次摘要。")
        note.setObjectName("SectionNote")
        note.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(note)

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        self.btn_import_file = QPushButton("导入原始数据")
        self.btn_import_file.setObjectName("PrimaryButton")
        self.btn_import_file.clicked.connect(self._handle_import_file)
        action_row.addWidget(self.btn_import_file)
        action_row.addStretch()
        layout.addLayout(action_row)

        self.overview_panel_row = ResponsivePanelRow()
        self.overview_state_summary = InfoBlock("当前批次状态摘要", "当前批次尚未建立。")
        self.overview_task_pool_summary = InfoBlock("当前任务池摘要", "当前待处理：0 类 / 0 条")
        self.overview_run_summary = InfoBlock("最近一次运行结果摘要", self.last_run_summary)
        self.overview_export_summary = InfoBlock("当前导出状态", "当前尚未解锁导出。")
        self.overview_panel_row.add_panel(self.overview_state_summary)
        self.overview_panel_row.add_panel(self.overview_task_pool_summary)
        self.overview_panel_row.add_panel(self.overview_run_summary)
        self.overview_panel_row.add_panel(self.overview_export_summary)
        layout.addWidget(self.overview_panel_row)

        self.sheet_preview_table = QTableWidget(0, 6)
        self.sheet_preview_table.setHorizontalHeaderLabels(["样本1", "样本2", "样本3", "样本4", "样本5", "样本6"])
        self.sheet_preview_table.verticalHeader().setVisible(False)
        self.sheet_preview_table.setMinimumHeight(340)
        layout.addWidget(self.sheet_preview_table)
        layout.addStretch()
        return shell

    def _build_task_stage(self, stage_key: str, stage_title: str, stage_note: str) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        shell.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title = QLabel(stage_title)
        title.setObjectName("PanelTitle")
        note = QLabel(stage_note)
        note.setObjectName("SectionNote")
        note.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(note)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        count_label = QLabel("当前待处理：0 条")
        count_label.setObjectName("TaskCardBody")
        setattr(self, f"{stage_key}_count_label", count_label)

        btn_clear = QPushButton("模拟清零当前分类")
        btn_clear.clicked.connect(lambda checked=False, key=stage_key: self._clear_task_category(key))

        top_row.addWidget(count_label)
        top_row.addStretch()
        top_row.addWidget(btn_clear)
        layout.addLayout(top_row)

        if stage_key == "actual_freight":
            self.btn_save_actual_freight_current = QPushButton("保存当前行")
            self.btn_save_actual_freight_current.setObjectName("PrimaryButton")
            self.btn_save_actual_freight_current.clicked.connect(self._save_actual_freight_current_row)

            self.btn_save_actual_freight_all = QPushButton("保存全部已编辑")
            self.btn_save_actual_freight_all.clicked.connect(self._save_actual_freight_all_edited)

            top_row.addWidget(self.btn_save_actual_freight_current)
            top_row.addWidget(self.btn_save_actual_freight_all)

            table = QTableWidget(0, 8)
            table.setHorizontalHeaderLabels(
                ["月份", "客户编码", "客户名称", "本月销售额(已审核)", "实际运费金额", "测算依据", "状态", "命中行数"]
            )
            table.itemChanged.connect(self._on_actual_freight_item_changed)
        else:
            table = QTableWidget(0, 4)
            table.setHorizontalHeaderLabels(["字段1", "字段2", "字段3", "字段4"])

        table.verticalHeader().setVisible(False)
        table.setMinimumHeight(420)
        setattr(self, f"{stage_key}_table", table)
        layout.addWidget(table)
        layout.addStretch()
        return shell

    def _build_run_check_stage(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        shell.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title = QLabel("运行与校验面板")
        title.setObjectName("PanelTitle")
        note = QLabel("这里承载可反复触发的运行动作与结果摘要。当前阶段先用演示运行逻辑表达循环感。")
        note.setObjectName("SectionNote")
        note.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(note)

        self.run_panel_row = ResponsivePanelRow()
        self.run_advice = InfoBlock("当前运行建议", "当前尚未进入运行阶段。")
        self.run_last_summary = InfoBlock("最近一次运行摘要", self.last_run_summary)
        self.run_gate_summary = InfoBlock("当前校验结论", "当前尚未满足导出条件。")
        self.run_next_summary = InfoBlock("下一步建议", "请先导入并建立任务池。")
        self.run_panel_row.add_panel(self.run_advice)
        self.run_panel_row.add_panel(self.run_last_summary)
        self.run_panel_row.add_panel(self.run_gate_summary)
        self.run_panel_row.add_panel(self.run_next_summary)
        layout.addWidget(self.run_panel_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_run_refresh = QPushButton("运行刷新（演示）")
        self.btn_run_refresh.setObjectName("PrimaryButton")
        self.btn_run_refresh.clicked.connect(self._run_refresh)
        btn_row.addWidget(self.btn_run_refresh)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        layout.addStretch()
        return shell

    def _build_export_stage(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        shell.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(shell)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        title = QLabel("导出结果面板")
        title.setObjectName("PanelTitle")
        note = QLabel("这个面板不再是首页固定按钮区，而是通关门。当前阶段先展示解锁 / 未解锁状态。")
        note.setObjectName("SectionNote")
        note.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(note)

        self.export_unlock_summary = InfoBlock("当前解锁状态", "当前尚未解锁导出。")
        self.export_reason_summary = InfoBlock("未解锁原因 / 导出前确认", "请先完成当前批次任务与运行校验。")
        layout.addWidget(self.export_unlock_summary)
        layout.addWidget(self.export_reason_summary)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_export_final = QPushButton("导出最终版")
        self.btn_export_full = QPushButton("导出全母版")
        self.btn_export_final.setObjectName("PrimaryButton")
        self.btn_export_full.setObjectName("PrimaryButton")
        self.btn_export_final.setEnabled(False)
        self.btn_export_full.setEnabled(False)
        self.btn_export_final.clicked.connect(lambda: self._show_feedback("当前状态：本轮暂不执行真实导出写出。"))
        self.btn_export_full.clicked.connect(lambda: self._show_feedback("当前状态：本轮暂不执行真实导出写出。"))
        btn_row.addWidget(self.btn_export_final)
        btn_row.addWidget(self.btn_export_full)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        layout.addStretch()
        return shell

    def _build_bottom_feedback_bar(self) -> QWidget:
        shell = QFrame()
        shell.setObjectName("PanelShell")
        layout = QHBoxLayout(shell)
        layout.setContentsMargins(14, 8, 14, 8)
        layout.setSpacing(8)
        title = QLabel("底部轻反馈")
        title.setObjectName("PanelTitle")
        self.bottom_feedback_label = QLabel(self.current_feedback)
        self.bottom_feedback_label.setObjectName("TaskCardBody")
        self.bottom_feedback_label.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(self.bottom_feedback_label, 1)
        return shell

    def _task_pool_summary_text(self) -> str:
        counted_items = [
            ("排除规则", self.task_pool["exclude_rules"]),
            ("结算条款", self.task_pool["settlement_terms"]),
            ("自提判定", self.task_pool["self_pickup"]),
            ("按实际运费", self.task_pool["actual_freight"]),
        ]
        active = [(name, count) for name, count in counted_items if count > 0]
        total_categories = len(active)
        total_items = sum(count for _, count in active)
        return f"当前待处理：{total_categories} 类 / {total_items} 条"

    def _compute_progress(self) -> int:
        if self.current_batch_state == "未导入":
            return 0
        if self.current_batch_state in {"可导出", "已导出"}:
            return 100

        total_items = sum(self.task_pool.values())
        base = {
            "已导入": 12,
            "已运行识别": 25,
            "待维护": 35,
            "维护中": 45,
            "可再运行": 80,
            "可导出": 100,
            "已导出": 100,
        }.get(self.current_batch_state, 0)

        if self.current_batch_state in {"待维护", "维护中"}:
            if total_items == 0:
                return 80
            solved_ratio = max(0.0, min(1.0, 1 - total_items / max(total_items, 24)))
            return max(base, min(79, int(35 + solved_ratio * 40)))
        return base

    def _current_focus_text(self) -> str:
        mapping = {
            "overview": "当前焦点：批次总览",
            "exclude_rules": "当前焦点：排除规则维护",
            "settlement_terms": "当前焦点：结算条款维护",
            "self_pickup": "当前焦点：自提判定维护",
            "actual_freight": "当前焦点：按实际运费维护",
            "run_check": "当前焦点：运行与校验",
            "export": "当前焦点：导出结果",
        }
        return mapping.get(self.current_stage_key, "当前焦点：批次总览")

    def _recommended_task(self) -> dict[str, str]:
        if self.current_file_path == "":
            return {
                "key": "overview",
                "title": "当前推荐任务：先导入原始数据",
                "reason": "推荐原因：当前还没有建立批次，无法进入后续状态驱动处理。",
                "impact": "当前影响：没有文件就没有任务池，也无法进入运行与校验。",
                "remaining": "剩余：1 项",
                "next": "完成后：进入批次总览并建立当前批次。",
            }

        if self.gate_state == "已通过" or self.current_batch_state in {"可导出", "已导出"}:
            if self.current_batch_state == "已导出":
                return {
                    "key": "export",
                    "title": "当前推荐任务：本轮批次已完成",
                    "reason": "推荐原因：当前批次已经完成导出。",
                    "impact": "当前影响：本轮处理已结束。",
                    "remaining": "剩余：0 项",
                    "next": "完成后：如需处理下一批次，请重新导入文件。",
                }
            return {
                "key": "export",
                "title": "当前推荐任务：进入导出结果",
                "reason": "推荐原因：当前批次门禁已通过，可以执行导出。",
                "impact": "当前影响：完成后本轮批次正式结束。",
                "remaining": "剩余：0 项",
                "next": "完成后：批次状态将更新为已导出。",
            }

        blocking_order = [
            ("settlement_terms", "结算条款"),
            ("actual_freight", "按实际运费"),
            ("self_pickup", "自提判定"),
            ("exclude_rules", "排除规则"),
        ]
        for key, title in blocking_order:
            count = self.task_pool[key]
            if count > 0:
                return {
                    "key": key,
                    "title": f"当前推荐任务：先处理{title}",
                    "reason": "推荐原因：这是当前最阻塞批次推进的任务类别。",
                    "impact": "当前影响：未完成前无法稳定进入再次运行。",
                    "remaining": f"剩余：{count} 条",
                    "next": "完成后：更接近“可再运行”状态。",
                }

        return {
            "key": "run_check",
            "title": "当前推荐任务：执行再次运行",
            "reason": "推荐原因：当前维护项已清零，需要刷新批次状态与门禁结果。",
            "impact": "当前影响：不运行则无法确认是否进入可导出状态。",
            "remaining": "剩余：1 项",
            "next": "完成后：更新任务池与门禁状态。",
        }

    def _task_category_display(self, key: str) -> str:
        title = self.task_meta[key]["title"]
        if self.task_meta[key]["counted"]:
            return f"{title} ({self.task_pool[key]})"
        if key == "run_check":
            if self.current_file_path == "":
                return f"{title}｜未进入"
            if sum(self.task_pool.values()) == 0 and self.gate_state != "已通过":
                return f"{title}｜建议关注"
            return f"{title}｜查看"
        if key == "export":
            return f"{title}｜{'已解锁' if self.gate_state == '已通过' else '未解锁'}"
        return title

    def _stage_status_style(self, state: str, is_current: bool) -> str:
        if is_current:
            return "background:#eef4ff;border:1px solid #7ca5ff;border-radius:10px;padding:4px 8px;color:#1e4ea8;"
        states_order = {name: idx for idx, name in enumerate(self.batch_states)}
        current_index = states_order.get(self.current_batch_state, -1)
        target_index = states_order.get(state, -1)
        if target_index <= current_index and current_index != -1:
            return "background:#f3f9f3;border:1px solid #b6dfb8;border-radius:10px;padding:4px 8px;color:#2f6b31;"
        return "background:#f6f8fc;border:1px solid #e4e9f2;border-radius:10px;padding:4px 8px;color:#667186;"

    def _handle_import_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择原始 Excel 文件",
            "",
            "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)",
        )
        if not file_path:
            self._show_feedback("当前提示：你取消了导入，系统仍停留在当前批次外。")
            return

        workbook = None
        try:
            result = open_excel_with_source(file_path)
            workbook = result.workbook

            self.current_file_path = file_path
            self.current_file_name = os.path.basename(file_path)
            self.current_batch_no = datetime.now().strftime("BATCH-%Y%m%d-%H%M%S")

            self.available_sheet_names = result.available_sheet_names
            self.current_source_sheet = result.source_sheet_name
            self.current_source_reason = result.source_sheet_reason
            self.source_rows = result.source_rows
            self.source_cols = result.source_cols
            self.source_preview_rows = result.preview_rows

            raw_sheet = workbook[result.source_sheet_name]
            raw_rows = sheet_to_dict_rows(raw_sheet, header_row=1)

            rules_sheet = find_sheet_by_keywords(workbook, ["排除规则"])
            terms_sheet = find_sheet_by_keywords(workbook, ["结算条款"])
            self_pickup_sheet = find_sheet_by_keywords(workbook, ["自提判定"])
            actual_freight_sheet = find_sheet_by_keywords(workbook, ["按实际运费录入"])

            rule_rows = sheet_to_dict_rows(rules_sheet, header_row=2) if rules_sheet is not None else []
            term_rows = sheet_to_dict_rows(terms_sheet, header_row=2) if terms_sheet is not None else []
            self_pickup_rows = sheet_to_dict_rows(self_pickup_sheet, header_row=4) if self_pickup_sheet is not None else []
            actual_freight_rows = sheet_to_dict_rows(actual_freight_sheet, header_row=4) if actual_freight_sheet is not None else []

            self.actual_freight_tasks = build_actual_freight_tasks_from_raw_rows(
                raw_rows=raw_rows,
                rule_rows=rule_rows,
                term_rows=term_rows,
                self_pickup_rows=self_pickup_rows,
                actual_freight_rows=actual_freight_rows,
            )

            aft_summary = summarize_actual_freight_tasks(self.actual_freight_tasks)

            self.task_pool = {
                "exclude_rules": 0,
                "settlement_terms": 0,
                "self_pickup": 0,
                "actual_freight": aft_summary["pending"],
            }

            self.gate_state = "未通过"
            self.current_stage_key = "overview"

            if aft_summary["pending"] > 0:
                self.current_batch_state = "待维护"
                self.last_run_summary = (
                    f"原始数据现算完成：A3 总数 {aft_summary['total']} 条，待录入 {aft_summary['pending']} 条，已录入 {aft_summary['done']} 条。"
                )
                self._show_feedback(
                    f"最近一次运行：已导入《{self.current_file_name}》，A3 总数 {aft_summary['total']} 条，待录入 {aft_summary['pending']} 条，已录入 {aft_summary['done']} 条。"
                )
            elif aft_summary["total"] > 0:
                self.current_batch_state = "可再运行"
                self.last_run_summary = (
                    f"原始数据现算完成：A3 总数 {aft_summary['total']} 条，已全部录入，可进入再次运行。"
                )
                self._show_feedback(
                    f"最近一次运行：已导入《{self.current_file_name}》，A3 已全部录入，可进入再次运行。"
                )
            else:
                self.current_batch_state = "已运行识别"
                self.last_run_summary = "原始数据现算完成：当前未识别到 A3 按实际任务。"
                self._show_feedback(
                    f"最近一次运行：已导入《{self.current_file_name}》，原始数据现算后当前未识别到 A3 按实际任务。"
                )

            self._fill_sheet_preview_table()
            self._refresh_ui()
            self._scroll_current_stage_to_top()

        except Exception as exc:
            self.actual_freight_tasks = []
            self.task_pool["actual_freight"] = 0
            self.current_batch_state = "未导入"
            self.gate_state = "未通过"
            self._show_feedback(f"当前卡点：导入失败，错误为 {exc}")

        finally:
            close_workbook_safely(workbook)

    def _switch_stage(self, key: str) -> None:
        self.current_stage_key = key
        self.stage_stack.setCurrentIndex(self.stage_index[key])
        self._scroll_current_stage_to_top()
        self._refresh_ui()

    def _on_actual_freight_item_changed(self, item: QTableWidgetItem) -> None:
        if self._is_refreshing_actual_freight_table:
            return
        table: QTableWidget = getattr(self, "actual_freight_table", None)
        if table is None or item is None or table is not item.tableWidget():
            return
        if item.column() not in {4, 5}:
            return
        if item.row() < 0 or item.row() >= len(self.actual_freight_tasks):
            return
        self.actual_freight_edited_rows.add(item.row())

    def _apply_actual_freight_row_to_task(self, row_index: int) -> tuple[bool, str]:
        table: QTableWidget = getattr(self, "actual_freight_table", None)
        if table is None:
            return False, "当前维护面板未就绪。"
        if row_index < 0 or row_index >= table.rowCount() or row_index >= len(self.actual_freight_tasks):
            return False, "当前选中行无效，请重新选择。"

        amount_item = table.item(row_index, 4)
        basis_item = table.item(row_index, 5)
        amount_text = "" if amount_item is None else amount_item.text().strip()
        basis_text = "" if basis_item is None else basis_item.text().strip()

        if amount_text == "":
            return False, f"第 {row_index + 1} 行保存失败：请填写“实际运费金额”。"
        try:
            amount_value = float(amount_text.replace(",", ""))
        except Exception:
            return False, f"第 {row_index + 1} 行保存失败：“实际运费金额”必须是数字。"

        if basis_text == "":
            return False, f"第 {row_index + 1} 行保存失败：请填写“测算依据”。"

        task = self.actual_freight_tasks[row_index]
        task.actual_freight_amount = amount_value
        task.basis = basis_text
        task.status = "已录入"
        return True, ""

    def _save_actual_freight_current_row(self) -> None:
        if self.current_file_path == "":
            self._show_feedback("当前提示：请先导入原始数据，再进行按实际运费维护。")
            return
        table: QTableWidget = getattr(self, "actual_freight_table", None)
        if table is None or table.rowCount() == 0:
            self._show_feedback("当前提示：当前没有可维护的按实际运费任务。")
            return

        row_index = table.currentRow()
        if row_index < 0:
            self._show_feedback("当前提示：请先选中一条任务，再执行“保存当前行”。")
            return

        ok, message = self._apply_actual_freight_row_to_task(row_index)
        if not ok:
            self._show_feedback(message)
            return

        self.actual_freight_edited_rows.discard(row_index)
        aft_summary = summarize_actual_freight_tasks(self.actual_freight_tasks)
        self.task_pool["actual_freight"] = aft_summary["pending"]
        self._show_feedback(f"当前状态：第 {row_index + 1} 行已保存，按实际运费待录入剩余 {aft_summary['pending']} 条。")
        self._refresh_ui()

    def _save_actual_freight_all_edited(self) -> None:
        if self.current_file_path == "":
            self._show_feedback("当前提示：请先导入原始数据，再进行按实际运费维护。")
            return
        table: QTableWidget = getattr(self, "actual_freight_table", None)
        if table is None or table.rowCount() == 0:
            self._show_feedback("当前提示：当前没有可维护的按实际运费任务。")
            return

        target_rows = sorted(idx for idx in self.actual_freight_edited_rows if 0 <= idx < table.rowCount())
        if len(target_rows) == 0:
            self._show_feedback("当前提示：没有检测到已编辑行，请先修改“实际运费金额”或“测算依据”。")
            return

        validated_updates: list[tuple[int, float, str]] = []
        for row_index in target_rows:
            amount_item = table.item(row_index, 4)
            basis_item = table.item(row_index, 5)
            amount_text = "" if amount_item is None else amount_item.text().strip()
            basis_text = "" if basis_item is None else basis_item.text().strip()

            if amount_text == "":
                self._show_feedback(f"第 {row_index + 1} 行保存失败：请填写“实际运费金额”。")
                return
            try:
                amount_value = float(amount_text.replace(",", ""))
            except Exception:
                self._show_feedback(f"第 {row_index + 1} 行保存失败：“实际运费金额”必须是数字。")
                return

            if basis_text == "":
                self._show_feedback(f"第 {row_index + 1} 行保存失败：请填写“测算依据”。")
                return

            validated_updates.append((row_index, amount_value, basis_text))

        for row_index, amount_value, basis_text in validated_updates:
            task = self.actual_freight_tasks[row_index]
            task.actual_freight_amount = amount_value
            task.basis = basis_text
            task.status = "已录入"

        self.actual_freight_edited_rows.clear()
        aft_summary = summarize_actual_freight_tasks(self.actual_freight_tasks)
        self.task_pool["actual_freight"] = aft_summary["pending"]
        self._show_feedback(
            f"当前状态：已保存 {len(target_rows)} 条编辑记录，按实际运费待录入剩余 {aft_summary['pending']} 条。"
        )
        self._refresh_ui()

    def _open_recommended_task(self) -> None:
        recommendation = self._recommended_task()
        self._switch_stage(recommendation["key"])

    def _scroll_current_stage_to_top(self) -> None:
        current_widget = self.stage_stack.currentWidget()
        if isinstance(current_widget, QScrollArea):
            current_widget.verticalScrollBar().setValue(0)

    def _clear_task_category(self, key: str) -> None:
        if self.current_file_path == "":
            self._show_feedback("当前提示：请先导入原始数据，再进入任务维护。")
            return
        if key not in self.task_pool:
            return

        if key == "actual_freight":
            self.actual_freight_tasks = []
        self.task_pool[key] = 0

        remaining = sum(self.task_pool.values())
        if remaining > 0:
            self.current_batch_state = "维护中"
            self.last_run_summary = f"{self.task_meta[key]['title']} 已清零，当前批次仍有其他待处理任务。"
            self._show_feedback(f"当前提示：{self.task_meta[key]['title']} 已清零，建议继续处理剩余分类。")
        else:
            self.current_batch_state = "可再运行"
            self.last_run_summary = "当前维护项已清零，建议再次运行刷新批次状态。"
            self._show_feedback("当前提示：当前维护项已清零，建议进入“运行与校验”并执行再次运行。")

        self._refresh_ui()
        self._scroll_current_stage_to_top()

    def _run_refresh(self) -> None:
        if self.current_file_path == "":
            self._show_feedback("当前提示：请先导入原始数据，再执行运行刷新。")
            return

        remaining = sum(self.task_pool.values())
        if remaining > 0:
            self.current_batch_state = "待维护"
            self.gate_state = "未通过"
            self.last_run_summary = "运行完成：当前仍有待处理任务，请先继续维护后再尝试刷新。"
            self._show_feedback("当前卡点：当前仍有待处理任务，运行后批次继续停留在待维护状态。")
        else:
            self.current_batch_state = "可导出"
            self.gate_state = "已通过"
            self.last_run_summary = "运行完成：当前批次门禁通过，可进入导出结果。"
            self._show_feedback("当前状态：门禁通过，可进入导出结果。")
        self._refresh_ui()
        self._scroll_current_stage_to_top()

    def _show_feedback(self, text: str) -> None:
        self.current_feedback = text
        self.bottom_feedback_label.setText(text)

    def _refresh_ui(self) -> None:
        self.total_progress_value = self._compute_progress()
        recommendation = self._recommended_task()
        aft_summary = summarize_actual_freight_tasks(self.actual_freight_tasks)

        self.info_batch.set_value(self.current_batch_no)
        self.info_file.set_value(self.current_file_name)
        self.info_source.set_value(self.current_source_sheet)
        self.info_source_reason.set_value(self.current_source_reason)
        self.info_batch_state.set_value(self.current_batch_state)
        self.info_gate_state.set_value(self.gate_state)
        self.info_last_run.set_value(self.last_run_summary)
        self.progress_current_state.set_value(self._current_focus_text())
        self.progress_task_pool.set_value(self._task_pool_summary_text())

        current_index = self.batch_states.index(self.current_batch_state) + 1 if self.current_batch_state in self.batch_states else 0
        self.progress_ratio.set_value(f"{current_index} / {len(self.batch_states)}")
        self.progress_bar.setValue(self.total_progress_value)
        self.progress_bar.setFormat(f"{self.total_progress_value}%")

        for state, label in self.batch_state_track_items.items():
            label.setStyleSheet(self._stage_status_style(state, state == self.current_batch_state))

        for key, button in self.task_category_buttons.items():
            button.setText(self._task_category_display(key))
            button.setChecked(key == self.current_stage_key)
            button.setEnabled(self.current_file_path != "" or key == "overview")

        if hasattr(self, "btn_save_actual_freight_current"):
            has_file = self.current_file_path != ""
            has_rows = len(self.actual_freight_tasks) > 0
            self.btn_save_actual_freight_current.setEnabled(has_file and has_rows)
            self.btn_save_actual_freight_all.setEnabled(has_file and has_rows)

        self.current_task_title.setText(recommendation["title"])
        self.current_task_reason.setText(recommendation["reason"])
        self.current_task_remaining.setText(recommendation["remaining"])
        self.current_task_meta.setText(f"{recommendation['impact']}｜{recommendation['next']}")

        overview_state_text = (
            f"当前批次状态：{self.current_batch_state}。当前源表：{self.current_source_sheet}。"
            f"源表规模：{self.source_rows} 行 / {self.source_cols} 列。"
        )
        self.overview_state_summary.set_value(overview_state_text)
        self.overview_task_pool_summary.set_value(
            f"{self._task_pool_summary_text()}｜A3总数：{aft_summary['total']}｜A3已录入：{aft_summary['done']}"
        )
        self.overview_run_summary.set_value(self.last_run_summary)
        export_text = "当前门禁已通过，可进入导出结果。" if self.gate_state == "已通过" else "当前尚未满足导出条件。"
        self.overview_export_summary.set_value(export_text)

        self._refresh_task_stage("exclude_rules")
        self._refresh_task_stage("settlement_terms")
        self._refresh_task_stage("self_pickup")
        self._refresh_task_stage("actual_freight")

        self.run_advice.set_value(self._run_advice_text())
        self.run_last_summary.set_value(self.last_run_summary)
        self.run_gate_summary.set_value(
            "当前校验结论：门禁通过，可进入导出。" if self.gate_state == "已通过" else "当前校验结论：门禁尚未通过。"
        )
        self.run_next_summary.set_value(recommendation["next"])
        self.btn_run_refresh.setEnabled(self.current_file_path != "")

        unlocked = self.gate_state == "已通过"
        self.export_unlock_summary.set_value("当前已解锁导出。" if unlocked else "当前尚未解锁导出。")
        self.export_reason_summary.set_value(
            "当前批次门禁已通过，可执行导出最终版 / 全母版。" if unlocked else recommendation["impact"]
        )
        self.btn_export_final.setEnabled(unlocked)
        self.btn_export_full.setEnabled(unlocked)

        self.stage_stack.setCurrentIndex(self.stage_index[self.current_stage_key])
        self.bottom_feedback_label.setText(self.current_feedback)

    def _refresh_task_stage(self, key: str) -> None:
        count = self.task_pool[key]
        count_label: QLabel = getattr(self, f"{key}_count_label")
        table: QTableWidget = getattr(self, f"{key}_table")

        if key == "actual_freight":
            aft_summary = summarize_actual_freight_tasks(self.actual_freight_tasks)
            count_label.setText(
                f"当前待处理：{aft_summary['pending']} 条｜已录入：{aft_summary['done']} 条｜A3总数：{aft_summary['total']} 条"
            )
            rows = []
            for task in self.actual_freight_tasks:
                amount_text = "" if task.actual_freight_amount is None else f"{task.actual_freight_amount:.2f}"
                rows.append([
                    task.month,
                    task.customer_code,
                    task.customer_name,
                    f"{task.monthly_sales_audited:.2f}",
                    amount_text,
                    task.basis,
                    task.status,
                    str(task.raw_count),
                ])
            self._is_refreshing_actual_freight_table = True
            try:
                self._fill_table(table, rows, 8)
                for row_index in range(table.rowCount()):
                    for col_index in (4, 5):
                        item = table.item(row_index, col_index)
                        if item is None:
                            continue
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            finally:
                self._is_refreshing_actual_freight_table = False
            return

        count_label.setText(f"当前待处理：{count} 条")
        rows = []
        for index in range(count):
            rows.append([
                f"{self.task_meta[key]['title']}项 {index + 1}",
                "当前批次",
                self.current_batch_no,
                "演示数据",
            ])
        self._fill_table(table, rows, 4)

    def _run_advice_text(self) -> str:
        if self.current_file_path == "":
            return "当前尚未建立批次，请先导入原始数据。"
        if sum(self.task_pool.values()) > 0:
            return "当前仍有待处理任务，建议先完成任务维护，再执行运行刷新。"
        if self.gate_state == "已通过":
            return "当前批次已通过门禁，可直接进入导出结果。"
        return "当前维护项已清零，建议执行再次运行刷新批次状态。"

    def _fill_sheet_preview_table(self) -> None:
        rows = self.source_preview_rows
        self._fill_table(self.sheet_preview_table, rows, 6)

    def _fill_table(self, table: QTableWidget, rows: list[list[str]], col_count: int) -> None:
        table.setRowCount(len(rows))
        table.setColumnCount(col_count)
        for row_index, row_data in enumerate(rows):
            for col_index in range(col_count):
                value = row_data[col_index] if col_index < len(row_data) else ""
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                table.setItem(row_index, col_index, item)
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
