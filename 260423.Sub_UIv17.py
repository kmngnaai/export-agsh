import os
import sys
import json
import base64
import time
import traceback
import threading
import importlib.util
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict

from PySide6.QtCore import Qt, QTimer, Signal, QObject, QSize, QEvent
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QLabel, QPushButton, QLineEdit, QComboBox, QTextEdit, QTreeWidget,
    QTreeWidgetItem, QProgressBar, QVBoxLayout, QHBoxLayout, QGridLayout,
    QFrame, QSizePolicy, QSplitter, QCheckBox, QGroupBox, QToolButton,
    QPlainTextEdit, QHeaderView, QStyle, QTableWidget, QTableWidgetItem
)

BASE_DIR = os.path.dirname(__file__)
BACKEND_PATH = os.path.join(BASE_DIR, "00.Detaisublog_v26.py")
APP_TITLE = "Sub Detail Collector - Bảng điều khiển Audit"
CONFIG_FILE = os.path.join(BASE_DIR, "config_detail_sub_dashboard.json")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

ICON_PICK_B64 = """iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAACXBIWXMAAAsTAAALEwEAmpwYAAAEa0lEQVR4nO2b3WsUVxjG31ZjwcsqlZK0pb3IThIh0KVqKjYwszkTcAUrLHrd/hHW3myvam8SU1Q0BHrONoZA+pnGxuKmSqjG3cHSL9P7ll6lCrYWmqLNU84mWIzR3Z3Z3TO75/3BA4F8sOf5zZw5c2ZCxDAMwzAMwzAMwzAMwzAMwzAMsw4o6oGiE5B0A5L+giLUNJLeMT3GWIIp2gJJJyHpXs1LVyyhkvLzdS9esYQNKR35jSxfsYQH5/xGTDuKJWwIJI0YK1/dl/Ae2QoULRoXoCw+EyDpjvHylcVnwkMlTD4FXHgG+OZFoNAJBI7dKXTqLpaRb/8WszsO1E9A7gngqx3mBxzEPHMd3+OLZ7fWVoAu//Jz5gcXNEnmX7hZMwklAXzko+rMdXxXGwF6zjc9mKAJU0wAF7bvjy6Aj36Ezlz79egC9GrH9ECCJs2Vl/6OLoCXmgidQudKdAGmBxE0d1hAwAKsDrEAhwXYHGIBDguwOcQCHBZgc4gFOCzA5hALcFiAzSEW4LAAm0MswGEBNodYgMMCbA6xAIcF2BxiAQ4LsDnEAhwW0Gz5N+jC4nwvJmZ3I/vpPrwx6eLghwPwc34p+us3J93S9/TP/Dzfi5ViFwtAxCxd7cHYTB8On/PgKb+qHDnnYex8H5YWdrIAVJnbCz0Ynt4LPyeqLn599N8YmX4Vt6/1sABUkK/zSRwcH4hc/Pq8Pj6AS/mXWQAekbtFB0Of76158Q9FirPJ0WQbCwj+z3KhC8c+fq3+5d+P+DI9mg73L0uteOQfa2j5q3GlyGemMlusFzDUiGnn0WfCGasFzOWTVZdWjqol5MRhKwX8ca0bh8ZTxgW4yr/VP5Hebp2A4elwU0/Nz4DV68FpqwQsLewMfZNVJwH/pHKp560RMDbTF/rCWQ8Ba1PRcSsErBS7Svs0cRPgSfFbZiqzqeUFLM73hi+pngJKZ4HY1fICJmZ3x1ZASoq3KhHwp+kSESHZT/ZFKjgqj5+G/I8qEJC4YbpERIh+cBJXAa4SP5YXUEwMmy4REVJuq9mkAE+K38sLKHR3I3DumS4SIVNu/W9SgKvEclkBa9eB900XCasFXE+2oZi4aLpM2DgFPSjBGUGQuGu6VNh0Ed74mpAYQjHxE4qJO6YLRqsvQ5sdV/lvR7lZilRwmbhKHKVWJyUH98RVgPfB4CvU6mSz2Sdd5f8aNwGu9H/Rn41swFX+8fgJEO+SLfSPD3bohyBxEaDX/67a30424SlxJi4CPOmfJNsQY+JpfeNjWoAr/ZtVPZRvJVw5cMS8gMEM2YwnxdmwF87oEafIdjJTmU2e8j9rdPmuEuf7L/dvNj3+WJAeTW/VL8w2rHzpz4R+ObdVSY4m28KujKqddvjIfwz6Xc0wq6MKppwl6y+4leLlvG36dUF9g1SD4pf1Ol8veyv+AMwq+u5Ub1uE2TvSv6O3F6y7w6U6sLqBJ3bp93b0nr1+cKLfatbbGaUo/5YnxQ9r3zuqdzXDbqz9BwliJk9Gv54+AAAAAElFTkSuQmCC"""
ICON_RUN_B64 = """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"""
ICON_OPEN_B64 = """iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAABwElEQVR4nGNgGAWjAAoMVDX/E8L6KprXLWRkOBkGygEGENxBUwc8W66BFR/q1fhvrA5W81dfTc2a7g54tlzjf2sS0aGEE1PkgIeLNf77Ww+gA55RgAeNAwxQc9VsnA54vkr///+/v/9TC2B1gKpGDk4HvN4V9v/X2ys0CY0Qe6gD1JByE7oDPp5p/v/19gqqW/54qcZ/U02wXf/MVFT4cDrg2711/z+cqqG6A45NgKQHfVWNW3gT4e8Pt/6/3hFAdQcsLoclSI1VOB3wfJXB/3+/v/5/tlKH6g5oioclQM1qnA54syf6/683F2iSAOPd4TnAG6cDPp5r///l5mKaOMBWD2KPobq6FE4HfHuw+f/7ExVUt/zCdFgC1HyFtyT8/en+/1fbfKjugHV18EJoJ04HPF9j+v/fr8//n63QproDutPgOaATpwPe7Ev4//PVaZrEf5YfNARU1KNwOuDThe7/X27Mp4kDPMygCVBVWxOnA74/2v7//bFiqlt+e77Gf0M1cAL8FsrAwIzTAe+eXf9/d5Xr/1vzNKiKNzXBgl/jJIblIEBpU4torKIxgwEbMFDVeEcXrKaRgtUBo2DEAgDGFMGJkyqrNAAAAABJRU5ErkJggg=="""


def load_backend_module():
    if not os.path.exists(BACKEND_PATH):
        raise FileNotFoundError(f"Khong tim thay file backend: {BACKEND_PATH}")
    spec = importlib.util.spec_from_file_location("sub_detail_backend", BACKEND_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Không thể nạp backend module")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


backend = load_backend_module()


PALETTE = {
    "bg": "#F3F6FA",
    "panel": "#FFFFFF",
    "panel_alt": "#F7FAFD",
    "text": "#0F172A",
    "muted": "#5F7087",
    "line": "#D2DAE6",
    "line_soft": "#E5ECF4",
    "primary": "#6F92F3",
    "primary_hover": "#5F85EE",
    "primary_soft": "#E5EEFF",
    "primary_soft_2": "#F3F7FF",
    "success_soft": "#EDF8F2",
    "warning_soft": "#FCF7EA",
    "error_soft": "#FDF0F2",
    "success_text": "#17603C",
    "warning_text": "#8A6112",
    "error_text": "#B24A5F",
    "success": "#22A06B",
    "warning": "#D28A17",
    "error": "#D4556A",
}

STATUS_COLORS = {
    "OK": (PALETTE["success_soft"], PALETTE["success_text"]),
    "WARNING": (PALETTE["warning_soft"], PALETTE["warning_text"]),
    "ERROR": (PALETTE["error_soft"], PALETTE["error_text"]),
}


@dataclass
class IssueRecord:
    file_path: str
    status: str
    detail: str

    @property
    def file_name(self) -> str:
        return backend.get_file_name_from_path(self.file_path)

    @property
    def normalized_status(self) -> str:
        return backend.nz_str(self.status).upper().strip()

    @property
    def search_blob(self) -> str:
        return " | ".join([
            backend.nz_str(self.file_path),
            self.file_name,
            backend.nz_str(self.status),
            backend.nz_str(self.detail),
        ]).lower()


class BackendLogger:
    def __init__(self, sink):
        self.sink = sink
        self._lock = threading.Lock()

    def log(self, msg: str, level: str = "info"):
        with self._lock:
            self.sink(msg, level)


class WorkerSignals(QObject):
    progress = Signal(int, int, str)
    finished = Signal(dict)
    error = Signal(str, str)
    append_log = Signal(str, str)
    refresh_done = Signal(list, str)
    refresh_error = Signal(str)


class ProcessorWorker(threading.Thread):
    def __init__(self, processor, folder_path: str, output_path: str, repair_options, signals: WorkerSignals):
        super().__init__(daemon=True)
        self.processor = processor
        self.folder_path = folder_path
        self.output_path = output_path
        self.repair_options = repair_options
        self.signals = signals

    def run(self):
        try:
            throttle_seconds = 0.5
            last_emit = 0.0

            def progress_callback(done: int, total: int, file_path: str):
                nonlocal last_emit
                now = time.monotonic()
                force_emit = done >= total
                if force_emit or (now - last_emit) >= throttle_seconds:
                    last_emit = now
                    self.signals.progress.emit(done, total, file_path)
                self.signals.append_log.emit(f"[{done}/{total}] {backend.get_file_name_from_path(file_path)}", "info")

            summary = self.processor.run(
                self.folder_path,
                self.output_path,
                progress_callback=progress_callback,
                repair_options=self.repair_options,
            )
            self.signals.finished.emit(summary)
        except Exception as exc:
            self.signals.error.emit(str(exc), traceback.format_exc())


class RefreshWorker(threading.Thread):
    def __init__(self, owner, output_path: str, signals: WorkerSignals):
        super().__init__(daemon=True)
        self.owner = owner
        self.output_path = output_path
        self.signals = signals

    def run(self):
        try:
            items = self.owner._read_issue_items_from_output(self.output_path)
            self.signals.refresh_done.emit(items, self.output_path)
        except Exception as exc:
            self.signals.refresh_error.emit(str(exc))


class BadgeWidget(QFrame):
    def __init__(self, text: str, bg: str, fg: str):
        super().__init__()
        self.setObjectName("badge")
        self.setStyleSheet(f"QFrame#badge {{ background:{bg}; border:1px solid {bg}; border-radius:6px; }}")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 7, 14, 7)
        self.label = QLabel(text)
        self.label.setStyleSheet(f"color:{fg}; font: 600 10pt 'Segoe UI';")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

    def setText(self, text: str):
        self.label.setText(text)


class ClickablePathEdit(QFrame):
    clicked = Signal()
    focusChanged = Signal(bool)

    def __init__(self, placeholder: str = "", show_button: bool = False, button_text: str = "Chon"):
        super().__init__()
        self.setObjectName("pathEdit")
        self.setStyleSheet(
            f"""
            QFrame#pathEdit {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #FBFCFE);
                border:1px solid {PALETTE['line']};
                border-radius:8px;
            }}
            QFrame#pathEdit[focused='true'] {{
                border:1px solid {PALETTE['primary']};
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #F5F8FF);
            }}
            """
        )
        self.setProperty("focused", False)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 3, 10, 3)
        lay.setSpacing(8)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder)
        self.edit.setFrame(False)
        self.edit.setMinimumHeight(24)
        self.edit.setStyleSheet("QLineEdit { border:none; background:transparent; font: 9.2pt 'Segoe UI'; color:#0F172A; }")
        lay.addWidget(self.edit, 1)
        self.button = None
        if show_button:
            self.button = QPushButton(button_text)
            self.button.setCursor(Qt.PointingHandCursor)
            self.button.setMinimumHeight(28)
            self.button.setStyleSheet(
                f"""
                QPushButton {{ background:{PALETTE['primary']}; color:white; border:none; border-radius:6px; padding:0 12px; font:600 9.2pt 'Segoe UI'; }}
                QPushButton:hover {{ background:{PALETTE['primary_hover']}; }}
                """
            )
            lay.addWidget(self.button, 0)
        self._focus_state = False
        self.edit.installEventFilter(self)

    def text(self) -> str:
        return self.edit.text()

    def setText(self, text: str):
        self.edit.setText(text)

    def setPlaceholderText(self, text: str):
        self.edit.setPlaceholderText(text)

    def eventFilter(self, obj, event):
        if obj is self.edit:
            et = event.type()
            if et == QEvent.FocusIn and not self._focus_state:
                self._focus_state = True
                self.setProperty("focused", True)
                self.style().polish(self)
                self.update()
                self.focusChanged.emit(True)
            elif et == QEvent.FocusOut and self._focus_state:
                self._focus_state = False
                self.setProperty("focused", False)
                self.style().polish(self)
                self.update()
                self.focusChanged.emit(False)
        return QFrame.eventFilter(self, obj, event)



class SelectableTableCell(QLineEdit):
    clicked = Signal(int)
    doubleClicked = Signal(int)

    def __init__(self, row_index: int, text: str = "", parent=None):
        super().__init__(parent)
        self.row_index = row_index
        self.setText(text)
        self.setReadOnly(True)
        self.setFrame(False)
        self.setCursorPosition(0)
        self.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.setStyleSheet(
            "QLineEdit { background: transparent; border:none; padding: 1px 6px; color:#1F2937; selection-background-color:#2F6FDB; selection-color:#FFFFFF; }"
        )

    def mousePressEvent(self, event):
        self.clicked.emit(self.row_index)
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        self.doubleClicked.emit(self.row_index)
        super().mouseDoubleClickEvent(event)


class IssueGridTable(QTableWidget):
    rowClicked = Signal(int)
    rowDoubleClicked = Signal(int)

    SELECTED_BG = "#DDEEFF"
    SELECTED_BORDER = "#B9D7FF"
    DEFAULT_WIDTHS = [92, 220, 520]

    def __init__(self, parent=None):
        super().__init__(0, 3, parent)
        self.records: List[IssueRecord] = []
        self.setHorizontalHeaderLabels(["Status", "File", "Loi chi tiet"])
        self.setAlternatingRowColors(False)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setSelectionMode(QTableWidget.SingleSelection)
        self.setEditTriggers(QTableWidget.NoEditTriggers)
        self.setShowGrid(True)
        self.setWordWrap(False)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setStretchLastSection(False)
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.apply_column_widths(self.DEFAULT_WIDTHS)
        self.setStyleSheet(
            "QTableWidget { background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #FBFCFE); border:1px solid #E4EAF2; border-radius:8px; gridline-color:#E8EDF4; outline:0; }"
            "QTableWidget::item:selected { background:#DDEEFF; color:#0F172A; }"
            "QHeaderView::section { background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FDFEFF, stop:1 #F1F5FA); color:#58708C; border:none; border-bottom:1px solid #D8E1EC; border-right:1px solid #E6EDF5; padding:8px 8px; font:600 9pt 'Segoe UI'; }"
        )
        self.currentCellChanged.connect(self._emit_row_clicked)

    def apply_column_widths(self, widths):
        if not isinstance(widths, (list, tuple)) or len(widths) < 3:
            widths = self.DEFAULT_WIDTHS
        safe_widths = []
        for i, default in enumerate(self.DEFAULT_WIDTHS):
            try:
                value = int(widths[i])
            except Exception:
                value = default
            safe_widths.append(max(50, min(value, 2000)))
        for col, width in enumerate(safe_widths):
            self.setColumnWidth(col, width)

    def column_widths(self) -> List[int]:
        return [int(self.columnWidth(i)) for i in range(3)]

    def _emit_row_clicked(self, currentRow, currentColumn, previousRow, previousColumn):
        self._refresh_row_highlight()
        if 0 <= currentRow < len(self.records):
            self.rowClicked.emit(currentRow)

    def _row_palette(self, rec: 'IssueRecord', idx: int):
        if rec.normalized_status == 'OK':
            return ('#F5FBF6' if idx % 2 == 0 else '#FFFFFF', PALETTE['success_text'])
        if rec.normalized_status == 'WARNING':
            return ('#FBF7EC' if idx % 2 == 0 else '#F9F3E8', PALETTE['warning_text'])
        if rec.normalized_status == 'ERROR':
            return ('#FCEFF2' if idx % 2 == 0 else '#FAEDEF', PALETTE['error_text'])
        return ('#F9FAFB' if idx % 2 == 0 else '#FFFFFF', PALETTE['text'])

    def _cell_style(self, bg: str, fg: str, selected: bool = False) -> str:
        font_weight = "600" if selected else "400"
        border_color = self.SELECTED_BORDER if selected else bg
        return (
            f"QLineEdit {{ background:{bg}; border:1px solid {border_color}; "
            f"padding: 1px 6px; color:{fg}; font-weight:{font_weight}; "
            "selection-background-color:#2F6FDB; selection-color:#FFFFFF; }}"
        )

    def _apply_row_style(self, row: int, selected: bool = False):
        if row < 0 or row >= len(self.records):
            return
        rec = self.records[row]
        base_bg, status_fg = self._row_palette(rec, row)
        bg = self.SELECTED_BG if selected else base_bg
        for col in range(3):
            item = self.item(row, col)
            if item is not None:
                item.setBackground(QColor(bg))
            cell = self.cellWidget(row, col)
            if isinstance(cell, QLineEdit):
                fg = status_fg if col == 0 else '#0F172A'
                cell.setStyleSheet(self._cell_style(bg, fg, selected=selected))

    def _refresh_row_highlight(self):
        selected_row = self.currentRow()
        for row in range(len(self.records)):
            self._apply_row_style(row, selected=(row == selected_row))

    def set_records(self, records: List['IssueRecord'], selected_index: int = -1):
        old_widths = self.column_widths()
        self.clearContents()
        self.records = list(records or [])
        self.setRowCount(len(self.records))
        for idx, rec in enumerate(self.records):
            bg, _fg = self._row_palette(rec, idx)
            base_item = QTableWidgetItem()
            base_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            base_item.setBackground(QColor(bg))
            self.setVerticalHeaderItem(idx, base_item)
            for col, value in enumerate([rec.normalized_status, rec.file_name, rec.detail]):
                item = QTableWidgetItem('')
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item.setBackground(QColor(bg))
                self.setItem(idx, col, item)
                cell = SelectableTableCell(idx, backend.nz_str(value))
                cell.clicked.connect(self._cell_clicked)
                cell.doubleClicked.connect(self._cell_double_clicked)
                self.setCellWidget(idx, col, cell)
            self.setRowHeight(idx, 28)
            self._apply_row_style(idx, selected=False)
        self.apply_column_widths(old_widths)
        if self.records:
            if selected_index < 0 or selected_index >= len(self.records):
                selected_index = 0
            self.selectRow(selected_index)
            self.setCurrentCell(selected_index, 0)
            self._refresh_row_highlight()
        else:
            self.clearSelection()

    def _cell_clicked(self, row_index: int):
        if 0 <= row_index < len(self.records):
            self.selectRow(row_index)
            self.setCurrentCell(row_index, 0)
            self._refresh_row_highlight()
            self.rowClicked.emit(row_index)

    def _cell_double_clicked(self, row_index: int):
        if 0 <= row_index < len(self.records):
            self.selectRow(row_index)
            self.setCurrentCell(row_index, 0)
            self._refresh_row_highlight()
            self.rowDoubleClicked.emit(row_index)

    def current_row_index(self) -> int:
        row = self.currentRow()
        return row if 0 <= row < len(self.records) else -1


class DashboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(800, 730)
        self.setMinimumSize(800, 640)

        self.folder_path = ""
        self.output_path = ""
        self.is_running = False
        self.log_visible = False
        self.repair_visible = False
        self.refreshing_output = False
        self.detail_panel_visible = False

        self.issue_items_master: List[IssueRecord] = []
        self.filtered_issue_items: List[IssueRecord] = []

        self.signals = WorkerSignals()
        self.signals.progress.connect(self._apply_progress)
        self.signals.finished.connect(self._finish_run)
        self.signals.error.connect(self._handle_worker_error)
        self.signals.append_log.connect(self._append_log)
        self.signals.refresh_done.connect(self._finish_refresh_output)
        self.signals.refresh_error.connect(self._handle_refresh_error)

        self.logger = BackendLogger(self._append_log)
        self.processor = backend.Processor(self.logger)

        self._setup_styles()
        self._build_ui()
        self._apply_button_icons()
        self._load_config()
        self._refresh_summary_labels()

    def _icon_from_b64(self, data_b64: str) -> Optional[QIcon]:
        try:
            raw = base64.b64decode(data_b64)
            pixmap = QPixmap()
            if pixmap.loadFromData(raw):
                return QIcon(pixmap)
        except Exception:
            pass
        return None

    def _set_button_icon(self, btn, standard_pixmap=None, size=16, custom_b64: Optional[str] = None):
        try:
            icon = None
            if custom_b64:
                icon = self._icon_from_b64(custom_b64)
            if icon is None and standard_pixmap is not None:
                icon = self.style().standardIcon(standard_pixmap)
            if icon is not None:
                btn.setIcon(icon)
                btn.setIconSize(QSize(size, size))
            btn.setCursor(Qt.PointingHandCursor)
        except Exception:
            pass

    def _apply_button_icons(self):
        self._set_button_icon(self.btn_pick_folder, size=18, custom_b64=ICON_PICK_B64)
        self._set_button_icon(self.btn_pick_output, size=18, custom_b64=ICON_PICK_B64)
        self._set_button_icon(self.btn_run, size=20, custom_b64=ICON_RUN_B64)
        self._set_button_icon(self.btn_open_output, size=18, custom_b64=ICON_OPEN_B64)
        self._set_button_icon(self.btn_refresh, QStyle.SP_BrowserReload, size=14)
        self._set_button_icon(self.btn_log, QStyle.SP_FileDialogDetailedView, size=14)
        self._set_button_icon(self.btn_detail, QStyle.SP_MessageBoxWarning, size=14)
        self._set_button_icon(self.btn_clear, QStyle.SP_DialogResetButton, size=14)
        self._set_button_icon(self.btn_open_file, QStyle.SP_FileIcon, size=14)
        self._set_button_icon(self.btn_open_folder, QStyle.SP_DirOpenIcon, size=14)
        self._set_button_icon(self.btn_copy_path, QStyle.SP_FileDialogContentsView, size=14)
        self._set_button_icon(self.btn_repair_toggle, QStyle.SP_ArrowRight, size=12)
        self._set_button_icon(self.btn_save_cfg, QStyle.SP_DialogSaveButton, size=14)
        self._set_button_icon(self.btn_reset_cfg, QStyle.SP_DialogResetButton, size=14)

    def _setup_styles(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background:{PALETTE['bg']};
                color:{PALETTE['text']};
                font: 10pt 'Segoe UI';
            }}

            QFrame[card='true'] {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #F9FBFE);
                border:1px solid {PALETTE['line']};
                border-radius:16px;
            }}
            QFrame[alt='true'] {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FCFDFE, stop:1 #F4F8FC);
                border:1px solid {PALETTE['line_soft']};
                border-radius:10px;
            }}
            QLabel#title {{ font:700 14pt 'Segoe UI'; color:#10233F; }}
            QLabel#section {{ font:700 10pt 'Segoe UI'; color:#122645; }}
            QLabel#muted {{ color:{PALETTE['muted']}; font: 8.7pt 'Segoe UI'; }}

            QPushButton {{ outline:none; }}
            QPushButton#btnPick {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ABC0FF, stop:1 #8DA7F3);
                color:#FFFFFF;
                border:1px solid #7F99E7;
                border-radius:9px;
                padding: 6px 14px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnPick:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #BDD0FF, stop:1 #9AB0F5); }}
            QPushButton#btnPick:pressed {{ background:#84A0F0; }}

            QPushButton#btnRun {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #8DB2FF, stop:1 #6C91F2);
                color:#FFFFFF;
                border:1px solid #5E83E8;
                border-radius:10px;
                padding: 6px 15px;
                font:700 9.2pt 'Segoe UI';
            }}
            QPushButton#btnRun:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #9FC0FF, stop:1 #7C9DF4); }}
            QPushButton#btnRun:pressed {{ background:#6689ED; }}

            QPushButton#btnOpenOutput {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F1FAF5, stop:1 #DFF1E8);
                color:#216547;
                border:1px solid #B8DFC8;
                border-radius:10px;
                padding: 6px 14px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnOpenOutput:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F8FDF9, stop:1 #E9F7F0); }}
            QPushButton#btnOpenOutput:pressed {{ background:#D7ECE0; }}

            QPushButton#btnRefresh {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #EDF5FF, stop:1 #D9E9FF);
                color:#25538C;
                border:1px solid #B8D1F1;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnRefresh:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F4F9FF, stop:1 #E5F1FF); }}
            QPushButton#btnRefresh:pressed {{ background:#D7E8FF; }}

            QPushButton#btnLog {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F4F1FF, stop:1 #E6DFFF);
                color:#6347AE;
                border:1px solid #D4C6FA;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnLog:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F9F7FF, stop:1 #EEE8FF); }}
            QPushButton#btnLog:pressed {{ background:#E4DBFF; }}

            QPushButton#btnDetail, QPushButton#btnOpenFile, QPushButton#btnOpenFolder {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFF8E8, stop:1 #FCEBC3);
                color:#8A6112;
                border:1px solid #EFD391;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnDetail:hover, QPushButton#btnOpenFile:hover, QPushButton#btnOpenFolder:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFDF4, stop:1 #FDEFCF); }}
            QPushButton#btnDetail:pressed, QPushButton#btnOpenFile:pressed, QPushButton#btnOpenFolder:pressed {{ background:#F7E3B5; }}

            QPushButton#btnCopyPath {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F0F8FF, stop:1 #DDEDFB);
                color:#2E5F8F;
                border:1px solid #BFDAF3;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnCopyPath:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F7FCFF, stop:1 #E8F4FF); }}
            QPushButton#btnCopyPath:pressed {{ background:#D7EAF9; }}

            QPushButton#btnSaveCfg {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #EEF9F1, stop:1 #DDF2E5);
                color:#17603C;
                border:1px solid #BAE0C8;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnSaveCfg:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #F5FCF7, stop:1 #E7F6ED); }}
            QPushButton#btnResetCfg {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFF6EB, stop:1 #FCE8D1);
                color:#985A1A;
                border:1px solid #F0D0A6;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnResetCfg:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFF9F0, stop:1 #FDEEDB); }}
            QPushButton#btnClear {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FDF2F4, stop:1 #F8E1E6);
                color:{PALETTE['error_text']};
                border:1px solid #F0C5CF;
                border-radius:9px;
                padding: 6px 12px;
                font:600 9pt 'Segoe UI';
            }}
            QPushButton#btnClear:hover {{ background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFF7F8, stop:1 #FBE8EC); }}
            QPushButton#btnClear:pressed {{ background:#F6D8DE; }}

            QLineEdit.rounded, QComboBox.rounded {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #F9FBFE);
                border:1px solid #CAD7E6;
                border-radius:8px;
                padding: 4px 10px;
                min-height:24px;
                color:#22324A;
            }}
            QLineEdit.rounded:focus, QComboBox.rounded:focus {{
                border:1px solid {PALETTE['primary']};
                background:#FFFFFF;
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 22px;
                border-left: 1px solid #DEE6F0;
                background:#F5F8FC;
                border-top-right-radius:8px;
                border-bottom-right-radius:8px;
            }}
            QComboBox::down-arrow {{ image:none; width:8px; height:8px; }}

            QTreeWidget {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #FBFCFE);
                border:1px solid #E4EAF2;
                border-radius:8px;
                outline:0;
            }}
            QTreeWidget::item {{ padding: 3px 7px; border:none; }}
            QTreeWidget::item:selected {{ background:#E6F0FB; color:#0F172A; }}
            QHeaderView::section {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FDFEFF, stop:1 #F1F5FA);
                color:#58708C;
                border:none;
                border-bottom:1px solid #D8E1EC;
                padding:8px 8px;
                font:600 9pt 'Segoe UI';
            }}

            QPlainTextEdit, QTextEdit {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #FBFCFE);
                border:1px solid #DDE5EF;
                border-radius:8px;
                padding:6px;
            }}
            QScrollBar:vertical {{ background:transparent; width:14px; margin:6px 2px 6px 2px; }}
            QScrollBar::handle:vertical {{ background:#B7C0CB; min-height:36px; border-radius:6px; }}
            QScrollBar::handle:vertical:hover {{ background:#A4AFBC; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0px; background:none; border:none; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background:transparent; }}
            QScrollBar:horizontal {{ background:transparent; height:14px; margin:2px 6px 2px 6px; }}
            QScrollBar::handle:horizontal {{ background:#B7C0CB; min-width:36px; border-radius:6px; }}
            QScrollBar::handle:horizontal:hover {{ background:#A4AFBC; }}
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width:0px; background:none; border:none; }}
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background:transparent; }}

            QGroupBox {{
                border:1px solid {PALETTE['line']};
                border-radius:10px;
                margin-top:12px;
                padding-top:8px;
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #FAFCFE);
                font:600 10pt 'Segoe UI';
            }}
            QGroupBox::title {{ subcontrol-origin:margin; left:12px; padding:0 6px; }}
            QProgressBar {{
                border:1px solid #DDE5EF;
                border-radius:6px;
                text-align:center;
                background:#EEF2F6;
                min-height:8px;
                max-height:8px;
            }}
            QProgressBar::chunk {{
                border-radius:6px;
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #89AFFF, stop:1 #6C91F2);
            }}
        """)

    def _build_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        root = QVBoxLayout(cw)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(5)

        # Header card
        header_card = QFrame()
        header_card.setProperty("card", True)
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(8, 7, 8, 7)
        header_layout.setSpacing(5)
        title = QLabel("Sub Detail Collector")
        title.setObjectName("title")
        header_layout.addWidget(title)

        info = QGridLayout()
        info.setHorizontalSpacing(8)
        info.setVerticalSpacing(8)
        info.setColumnStretch(1, 1)

        lbl_folder = QLabel("Thư mục:")
        lbl_folder.setObjectName("muted")
        lbl_folder.setFixedWidth(56)
        info.addWidget(lbl_folder, 0, 0)
        self.folder_input = ClickablePathEdit("Chon thu muc nguon", show_button=False)
        self.folder_input.setFixedHeight(36)
        self.folder_input.edit.returnPressed.connect(self._commit_path_inputs)
        self.folder_input.edit.editingFinished.connect(self._commit_path_inputs)
        info.addWidget(self.folder_input, 0, 1)
        self.btn_pick_folder = QPushButton("Chọn")
        self.btn_pick_folder.setObjectName("btnPick")
        self.btn_pick_folder.setCursor(Qt.PointingHandCursor)
        self.btn_pick_folder.setFixedSize(92, 34)
        self.btn_pick_folder.clicked.connect(self.select_folder)
        info.addWidget(self.btn_pick_folder, 0, 2)
        self.btn_run = QPushButton("Bắt đầu")
        self.btn_run.setObjectName("btnRun")
        self.btn_run.setFixedSize(148, 36)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        self.btn_run.clicked.connect(self.re_run)
        info.addWidget(self.btn_run, 0, 3)

        lbl_output = QLabel("Kết quả:")
        lbl_output.setObjectName("muted")
        lbl_output.setFixedWidth(56)
        info.addWidget(lbl_output, 1, 0)
        self.output_input = ClickablePathEdit("Chon hoac tao file ket qua", show_button=False)
        self.output_input.setFixedHeight(36)
        self.output_input.edit.returnPressed.connect(self._commit_path_inputs)
        self.output_input.edit.editingFinished.connect(self._commit_path_inputs)
        info.addWidget(self.output_input, 1, 1)
        self.btn_pick_output = QPushButton("Chọn")
        self.btn_pick_output.setObjectName("btnPick")
        self.btn_pick_output.setCursor(Qt.PointingHandCursor)
        self.btn_pick_output.setFixedSize(92, 34)
        self.btn_pick_output.clicked.connect(self.select_output_file)
        info.addWidget(self.btn_pick_output, 1, 2)
        self.btn_open_output = QPushButton("Mở output")
        self.btn_open_output.setCursor(Qt.PointingHandCursor)
        self.btn_open_output.setFixedSize(148, 36)
        self.btn_open_output.setObjectName("btnOpenOutput")
        self.btn_open_output.clicked.connect(self.open_output_file)
        info.addWidget(self.btn_open_output, 1, 3)

        header_layout.addLayout(info)
        root.addWidget(header_card)

        # Toolbar
        toolbar_card = QFrame()
        toolbar_card.setProperty("card", True)
        toolbar_layout = QHBoxLayout(toolbar_card)
        toolbar_layout.setContentsMargins(8, 7, 8, 7)
        toolbar_layout.setSpacing(6)

        filter_label = QLabel("Bộ lọc:")
        filter_label.setObjectName("muted")
        toolbar_layout.addWidget(filter_label)
        self.status_cb = QComboBox()
        self.status_cb.addItems(["ISSUES", "WARNING", "ERROR", "OK", "ALL"])
        self.status_cb.setCurrentText("ISSUES")
        self.status_cb.setFixedWidth(84)
        self.status_cb.setProperty("class", "rounded")
        self.status_cb.setStyleSheet("QComboBox { background:#FFFFFF; border:1px solid #D1D1D1; border-radius:6px; padding: 3px 8px; min-height:22px; } QComboBox:focus { border:1px solid %s; }" % PALETTE['primary'])
        self.status_cb.currentTextChanged.connect(self.refresh_issue_tree)
        toolbar_layout.addWidget(self.status_cb)

        search_label = QLabel("Tìm kiếm:")
        search_label.setObjectName("muted")
        toolbar_layout.addSpacing(8)
        toolbar_layout.addWidget(search_label)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Tìm theo file, lỗi, path...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setFixedWidth(175)
        self.search_edit.setProperty("class", "rounded")
        self.search_edit.setStyleSheet("QLineEdit { background:#FFFFFF; border:1px solid #D1D1D1; border-radius:6px; padding: 3px 8px; min-height:22px; } QLineEdit:focus { border:1px solid %s; }" % PALETTE['primary'])
        self.search_edit.textChanged.connect(self.refresh_issue_tree)
        toolbar_layout.addWidget(self.search_edit)
        toolbar_layout.addStretch(1)

        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.setObjectName("btnRefresh")
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.btn_refresh.clicked.connect(self.reload_issue_list_from_output_async)
        toolbar_layout.addWidget(self.btn_refresh)

        self.btn_log = QPushButton("Xem log")
        self.btn_log.setObjectName("btnLog")
        self.btn_log.setCursor(Qt.PointingHandCursor)
        self.btn_log.clicked.connect(self.toggle_log_panel)
        toolbar_layout.addWidget(self.btn_log)

        self.btn_detail = QPushButton("Lỗi")
        self.btn_detail.setObjectName("btnDetail")
        self.btn_detail.setCursor(Qt.PointingHandCursor)
        self.btn_detail.clicked.connect(self.toggle_detail_panel)
        toolbar_layout.addWidget(self.btn_detail)

        self.btn_clear = QPushButton("Clear")
        self.btn_clear.setObjectName("btnClear")
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear_issue_view)
        toolbar_layout.addWidget(self.btn_clear)
        root.addWidget(toolbar_card)

        # Main splitter
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(8)

        left_card = QFrame()
        left_card.setProperty("card", True)
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(8, 7, 8, 7)
        left_layout.setSpacing(4)
        sec = QLabel("Bang WARNING / ERROR")
        sec.setObjectName("section")
        left_layout.addWidget(sec)
        hint = QLabel("To den trong tung o de Ctrl+C copy. Keo gian cot tu do. Double click de mo file. Click 1 dong de xem chi tiet.")
        hint.setObjectName("muted")
        left_layout.addWidget(hint)

        self.issue_tree = IssueGridTable()
        self.issue_tree.rowClicked.connect(lambda *_: self._on_issue_selected())
        self.issue_tree.rowDoubleClicked.connect(lambda *_: self.open_selected_issue_file())
        self.issue_column_width_timer = QTimer(self)
        self.issue_column_width_timer.setSingleShot(True)
        self.issue_column_width_timer.timeout.connect(self._save_issue_column_widths_silent)
        self.issue_tree.horizontalHeader().sectionResized.connect(lambda *_: self._schedule_issue_column_width_save())
        left_layout.addWidget(self.issue_tree, 1)
        self.main_splitter.addWidget(left_card)

        self.detail_card = QFrame()
        self.detail_card.setProperty("card", True)
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(12, 10, 12, 12)
        detail_layout.setSpacing(8)
        dtitle = QLabel("Chi tiet audit")
        dtitle.setObjectName("section")
        detail_layout.addWidget(dtitle)

        self.detail_head = QFrame()
        self.detail_head.setProperty("alt", True)
        dh = QHBoxLayout(self.detail_head)
        dh.setContentsMargins(0, 0, 0, 0)
        dh.setSpacing(0)
        self.detail_accent = QFrame()
        self.detail_accent.setFixedWidth(5)
        self.detail_accent.setStyleSheet("background:#E5E7EB; border-top-left-radius:14px; border-bottom-left-radius:14px;")
        dh.addWidget(self.detail_accent)
        dtext_wrap = QWidget()
        dtext_lay = QVBoxLayout(dtext_wrap)
        dtext_lay.setContentsMargins(10, 10, 10, 10)
        dtext_lay.setSpacing(4)
        self.selected_file_label = QLabel("(Chua chon dong)")
        self.selected_file_label.setStyleSheet("font: 700 10pt 'Segoe UI';")
        self.selected_path_label = QLabel("-")
        self.selected_path_label.setObjectName("muted")
        dtext_lay.addWidget(self.selected_file_label)
        dtext_lay.addWidget(self.selected_path_label)
        dh.addWidget(dtext_wrap, 1)
        detail_layout.addWidget(self.detail_head)

        self.detail_box = QPlainTextEdit()
        self.detail_box.setReadOnly(True)
        self.detail_box.setStyleSheet("QPlainTextEdit { background:#FFFFFF; border:1px solid %s; border-radius:6px; padding:6px; }" % PALETTE['line_soft'])
        detail_layout.addWidget(self.detail_box, 1)

        action_row = QGridLayout()
        action_row.setHorizontalSpacing(8)
        action_row.setVerticalSpacing(0)
        self.btn_open_file = QPushButton("Mở file")
        self.btn_open_file.setObjectName("btnOpenFile")
        self.btn_open_file.setFixedHeight(36)
        self.btn_open_file.clicked.connect(self.open_selected_issue_file)
        action_row.addWidget(self.btn_open_file, 0, 0)
        self.btn_open_folder = QPushButton("Mở thư mục")
        self.btn_open_folder.setObjectName("btnOpenFolder")
        self.btn_open_folder.setFixedHeight(36)
        self.btn_open_folder.clicked.connect(self.open_selected_issue_folder)
        action_row.addWidget(self.btn_open_folder, 0, 1)
        self.btn_copy_path = QPushButton("Copy path")
        self.btn_copy_path.setObjectName("btnCopyPath")
        self.btn_copy_path.setFixedHeight(36)
        self.btn_copy_path.clicked.connect(self.copy_selected_issue_path)
        action_row.addWidget(self.btn_copy_path, 0, 2)
        action_row.setColumnStretch(0, 1)
        action_row.setColumnStretch(1, 1)
        action_row.setColumnStretch(2, 1)
        detail_layout.addLayout(action_row)

        self.main_splitter.addWidget(self.detail_card)
        self.detail_card.hide()
        self.main_splitter.setSizes([900, 0])
        root.addWidget(self.main_splitter, 1)

        # Progress card
        progress_card = QFrame()
        progress_card.setProperty("card", True)
        p_lay = QVBoxLayout(progress_card)
        p_lay.setContentsMargins(8, 7, 8, 7)
        p_lay.setSpacing(4)
        self.progress_text = QLabel("Tien do: San sang")
        p_lay.addWidget(self.progress_text)
        pwrap = QHBoxLayout()
        pwrap.setSpacing(8)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        pwrap.addWidget(self.progress, 1)
        self.progress_percent = QLabel("0%")
        self.progress_percent.setObjectName("muted")
        self.progress_percent.setStyleSheet("font: 700 9pt 'Segoe UI'; color:%s;" % PALETTE['muted'])
        pwrap.addWidget(self.progress_percent)
        p_lay.addLayout(pwrap)
        self.summary_line = QLabel("")
        self.summary_line.setObjectName("muted")
        self.summary_line.setStyleSheet("font: 9pt 'Segoe UI'; color:%s;" % PALETTE["muted"])
        self.summary_line.setWordWrap(True)
        p_lay.addWidget(self.summary_line)

        root.addWidget(progress_card)

        # Repair options
        self.btn_repair_toggle = QPushButton("Tùy chọn sửa ▸")
        self.btn_repair_toggle.setObjectName("btnSaveCfg")
        self.btn_repair_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_repair_toggle.clicked.connect(self.toggle_repair_options)
        root.addWidget(self.btn_repair_toggle, 0, Qt.AlignLeft)

        self.repair_group = QGroupBox("Tùy chọn sửa")
        repair_lay = QGridLayout(self.repair_group)
        repair_lay.setContentsMargins(6, 6, 6, 6)
        repair_lay.setHorizontalSpacing(6)
        repair_lay.setVerticalSpacing(4)
        self.chk_fix_folder = QCheckBox("Sua theo chuan ten folder")
        self.chk_fix_folder.setChecked(True)
        self.chk_fix_manual = QCheckBox("Sua theo thong tin nhap tay")
        repair_lay.addWidget(self.chk_fix_folder, 0, 0)
        repair_lay.addWidget(self.chk_fix_manual, 0, 1)
        repair_lay.addWidget(QLabel("Invoice No"), 1, 0)
        repair_lay.addWidget(QLabel("Invoice Date (dd/mm/yyyy)"), 1, 1)
        repair_lay.addWidget(QLabel("Destination"), 1, 2)
        self.ent_manual_inv = QLineEdit()
        self.ent_manual_inv.setProperty("class", "rounded")
        self.ent_manual_date = QLineEdit()
        self.ent_manual_date.setProperty("class", "rounded")
        self.ent_manual_dest = QLineEdit()
        self.ent_manual_dest.setProperty("class", "rounded")
        for ent in [self.ent_manual_inv, self.ent_manual_date, self.ent_manual_dest]:
            ent.setStyleSheet("QLineEdit { background:#FFFFFF; border:1px solid #D1D1D1; border-radius:6px; padding: 3px 8px; min-height:22px; } QLineEdit:focus { border:1px solid %s; }" % PALETTE['primary'])
        repair_lay.addWidget(self.ent_manual_inv, 2, 0)
        repair_lay.addWidget(self.ent_manual_date, 2, 1)
        repair_lay.addWidget(self.ent_manual_dest, 2, 2)
        btn_row = QHBoxLayout()
        self.btn_save_cfg = QPushButton("Lưu mặc định")
        self.btn_save_cfg.setObjectName("btnSaveCfg")
        self.btn_save_cfg.clicked.connect(self.save_config)
        btn_row.addWidget(self.btn_save_cfg)
        self.btn_reset_cfg = QPushButton("Reset")
        self.btn_reset_cfg.setObjectName("btnResetCfg")
        self.btn_reset_cfg.clicked.connect(self.reset_config)
        btn_row.addWidget(self.btn_reset_cfg)
        btn_row.addStretch(1)
        repair_lay.addLayout(btn_row, 3, 0, 1, 3)
        self.repair_group.hide()
        root.addWidget(self.repair_group)

        # Log panel
        self.log_card = QFrame()
        self.log_card.setProperty("card", True)
        log_lay = QVBoxLayout(self.log_card)
        log_lay.setContentsMargins(8, 7, 8, 7)
        log_title = QLabel("Log realtime")
        log_title.setObjectName("section")
        log_lay.addWidget(log_title)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("QPlainTextEdit { background:#0F172A; color:#E2E8F0; border:none; border-radius:6px; padding:8px; font: 9pt 'Consolas'; }")
        log_lay.addWidget(self.log_box)
        self.log_card.hide()
        root.addWidget(self.log_card)

    # ---------- helpers ----------
    def _append_log(self, msg: str, level: str = "info"):
        self.log_box.appendPlainText(msg)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _refresh_labels(self):
        self.folder_input.setText(self.folder_path or "")
        self.output_input.setText(self.output_path or "")

    def _commit_path_inputs(self):
        self.folder_path = self.folder_input.text().strip()
        self.output_path = self.output_input.text().strip()

    def _update_badges(self, total: int, ok: int, warning: int, error: int):
        return

    def _set_progress_style(self, mode: str):
        color = PALETTE['primary']
        pct_color = PALETTE['muted']
        if mode == 'success':
            color = PALETTE['success']
            pct_color = PALETTE['success_text']
        elif mode == 'warning':
            color = PALETTE['warning']
            pct_color = PALETTE['warning_text']
        elif mode == 'error':
            color = PALETTE['error']
            pct_color = PALETTE['error_text']
        self.progress.setStyleSheet(
            f"QProgressBar {{ border:none; border-radius:6px; text-align:center; background:#E5E7EB; min-height:7px; max-height:7px; }} QProgressBar::chunk {{ border-radius:6px; background:{color}; }}"
        )
        self.progress_percent.setStyleSheet(f"font: 700 9pt 'Segoe UI'; color:{pct_color};")

    def _compact_path(self, path: str) -> str:
        p = backend.nz_str(path).replace("\\", "/")
        if not p:
            return "-"
        parts = [x for x in p.split("/") if x]
        if len(parts) >= 2:
            folder = parts[-2]
            file_name = parts[-1]
            short_folder = (folder[:14] + "...") if len(folder) > 14 else folder
            short_file = (file_name[:24] + "...") if len(file_name) > 24 else file_name
            return f"{short_folder}/{short_file}"
        return (p[:36] + "...") if len(p) > 39 else p

    def _format_detail_text(self, text: str) -> str:
        t = backend.nz_str(text).replace(" | ", "\n")
        t = t.replace(" missing ", "\n")
        t = t.replace(" - FOLDER:", "\nFOLDER:")
        t = t.replace(" | FILE:", "\nFILE:")
        t = t.replace("; ", "\n")
        return t.strip() or "(Khong co noi dung chi tiet)"

    def _set_status_badge(self, status: str):
        bg, fg = STATUS_COLORS.get(status, ("#E5E7EB", PALETTE["muted"]))
        self.detail_accent.setStyleSheet(f"background:{fg if status in STATUS_COLORS else bg}; border-top-left-radius:14px; border-bottom-left-radius:14px;")
        self.selected_file_label.setStyleSheet(f"font: 700 10pt 'Segoe UI'; color:{fg if status in STATUS_COLORS else PALETTE['text']};")

    def _show_issue_detail(self, rec: Optional[IssueRecord]):
        if rec is None:
            self.selected_file_label.setText("(Chua chon dong)")
            self.selected_path_label.setText("-")
            self._set_status_badge("-")
            text = "Chon 1 dong de xem chi tiet."
        else:
            self.selected_file_label.setText(rec.file_name)
            self.selected_path_label.setText(self._compact_path(rec.file_path))
            self._set_status_badge(rec.normalized_status)
            text = self._format_detail_text(rec.detail if rec.normalized_status != "OK" else "(Khong co loi)")
        self.detail_box.setPlainText(text)

    # ---------- ui actions ----------
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chon thu muc nguon")
        if folder:
            self.folder_path = folder
            self._refresh_labels()
            self._append_log(f"Da chon thu muc: {folder}")

    def select_output_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Chon hoac tao file ket qua", self.output_path or "SUB_DETAIL_OUTPUT.xlsx", "Excel Files (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                QMessageBox.warning(self, "Sai dinh dang", "Chi duoc chon file .xlsx")
                return
            self.output_path = path
            self._refresh_labels()
            self._append_log(f"Da chon file ket qua: {path}")

    def open_output_file(self):
        if not self.output_path:
            QMessageBox.warning(self, "Chua co file", "Chua co duong dan file output.")
            return
        try:
            backend.open_file_with_default_app(self.output_path)
        except Exception as exc:
            QMessageBox.critical(self, "Khong mo duoc file output", str(exc))

    def toggle_log_panel(self):
        self.log_visible = not self.log_visible
        self.log_card.setVisible(self.log_visible)

    def show_detail_panel(self):
        self.detail_panel_visible = True
        self.detail_card.show()
        self.main_splitter.setSizes([1, 1])

    def hide_detail_panel(self):
        self.detail_panel_visible = False
        self.detail_card.hide()
        self.main_splitter.setSizes([1, 0])

    def toggle_detail_panel(self):
        if self.detail_panel_visible:
            self.hide_detail_panel()
        else:
            self.show_detail_panel()
            rec = self.get_selected_issue(show_warning=False)
            self._show_issue_detail(rec)

    def clear_issue_view(self):
        self.issue_items_master = []
        self.filtered_issue_items = []
        self.search_edit.clear()
        self.status_cb.setCurrentText("ISSUES")
        self.issue_tree.set_records([])
        self._show_issue_detail(None)
        self.progress_text.setText("Tien do: Da clear UI")
        self.progress_percent.setText("0%")
        self.progress.setValue(0)
        self._set_progress_style('modern')
        self.summary_line.setText("")
        self._refresh_summary_labels()
        self.logger.log("Da clear bang hien thi tren UI.")

    def _set_running_state(self, running: bool):
        self.is_running = running
        self.btn_run.setEnabled(not running)
        self.btn_refresh.setEnabled(not running and not self.refreshing_output)

    def re_run(self):
        self._commit_path_inputs()
        if self.is_running:
            QMessageBox.warning(self, "Dang xu ly", "Tien trinh dang chay. Hay cho hoan tat truoc.")
            return
        if not self.folder_path:
            QMessageBox.warning(self, "Thieu thu muc", "Hay chon thu muc truoc.")
            return
        if not self.output_path:
            QMessageBox.warning(self, "Thieu file ket qua", "Hay chon file ket qua truoc.")
            return
        self.start_process()

    def start_process(self):
        self._set_running_state(True)
        self._set_progress_style('modern')
        self.progress.setValue(0)
        self.progress_percent.setText("0%")
        self.progress_text.setText("Tien do: Dang bat dau xu ly...")
        self.log_box.setPlainText("")
        self._append_log(f"Bat dau xu ly...\n{self.folder_path}\n{self.output_path}")
        self.summary_line.setText("Dang xu ly...")
        self._append_log("Goi y: voi file .xls, hay dung bang audit de mo file va sua tay nhanh hon.")
        repair_options = backend.RepairOptions(
            use_folder_truth=self.chk_fix_folder.isChecked(),
            use_manual_values=self.chk_fix_manual.isChecked(),
            manual_invoice_no=self.ent_manual_inv.text().strip(),
            manual_invoice_date=self.ent_manual_date.text().strip(),
            manual_destination=self.ent_manual_dest.text().strip(),
        )
        self.worker = ProcessorWorker(self.processor, self.folder_path, self.output_path, repair_options, self.signals)
        self.worker.start()

    def _apply_progress(self, done: int, total: int, file_path: str):
        percent = 0 if total <= 0 else int(done / total * 100)
        self.progress.setValue(percent)
        self.progress_percent.setText(f"{percent}%")
        self.progress_text.setText(f"Tien do: Dang xu ly {done}/{total} | {backend.get_file_name_from_path(file_path)}")

    def _finish_run(self, summary: dict):
        self._set_running_state(False)
        done_value = 100 if summary.get("total", 0) > 0 else 0
        err = int(summary.get("error", 0) or 0)
        warn = int(summary.get("warning", 0) or 0)
        if done_value >= 100 and err > 0:
            self._set_progress_style('error')
        elif done_value >= 100 and warn > 0:
            self._set_progress_style('warning')
        elif done_value >= 100:
            self._set_progress_style('success')
        else:
            self._set_progress_style('modern')
        self.progress.setValue(done_value)
        self.progress_percent.setText("100%" if summary.get("total", 0) > 0 else "0%")
        self.progress_text.setText("Tien do: Hoan tat")
        self.merge_issue_list(summary.get("issues", []), summary.get("scanned_files", []))
        self._update_stat_cards(summary)

        msg = (
            f"Hoan tat | Tong file quet: {summary.get('total', 0)} | "
            f"OK: {summary.get('ok', 0)} | WARNING: {summary.get('warning', 0)} | "
            f"ERROR: {summary.get('error', 0)} | Skip Path: {summary.get('skip_path', 0)} | "
            f"Skip Signature: {summary.get('skip_signature', 0)} | Repaired: {summary.get('repaired', 0)}"
        )
        self.summary_line.setText(msg)
        self._append_log(msg)

    def _handle_worker_error(self, exc_text: str, tb: str):
        self._set_running_state(False)
        self.progress_text.setText("Tien do: Phat sinh loi")
        self.progress_percent.setText("0%")
        self._append_log(f"Loi: {exc_text}\n\n{tb}")
        QMessageBox.critical(self, "Lỗi", exc_text)

    def _update_stat_cards(self, summary: dict):
        self._update_badges(
            int(summary.get('total', 0) or 0),
            int(summary.get('ok', 0) or 0),
            int(summary.get('warning', 0) or 0),
            int(summary.get('error', 0) or 0),
        )

    def reload_issue_list_from_output_async(self):
        if self.refreshing_output or self.is_running:
            return
        self._commit_path_inputs()
        if not self.output_path:
            QMessageBox.warning(self, "Chua co file output", "Hay chon file output truoc.")
            return
        if not os.path.exists(self.output_path):
            QMessageBox.warning(self, "Khong tim thay file", f"Khong tim thay file output:\n{self.output_path}")
            return
        self.refreshing_output = True
        self.btn_refresh.setEnabled(False)
        self.progress_text.setText("Tien do: Dang doc lai SUB_DETAIL + LOG_SUB_DETAIL...")
        self.progress_percent.setText("...")
        self._append_log(f"Dang refresh tu output: {self.output_path}")
        self.refresh_worker = RefreshWorker(self, self.output_path, self.signals)
        self.refresh_worker.start()

    def _finish_refresh_output(self, items, output_path):
        self.refreshing_output = False
        self.btn_refresh.setEnabled(True)
        self.status_cb.setCurrentText("ISSUES")
        self.load_issue_list(items)
        self.progress_text.setText("Tien do: Da refresh output")
        self.progress_percent.setText("100%" if items else "0%")
        self.summary_line.setText(f"Refresh output xong | So dong hien thi: {len(items)} | File: {backend.get_file_name_from_path(output_path)}")
        self._append_log(f"Da refresh danh sach loi tu file output: {output_path}")

    def _handle_refresh_error(self, err_text: str):
        self.refreshing_output = False
        self.btn_refresh.setEnabled(True)
        self.progress_text.setText("Tien do: Refresh output loi")
        self.progress_percent.setText("0%")
        self.summary_line.setText(f"Refresh output loi: {err_text}")
        self._append_log(f"Refresh output loi: {err_text}")
        QMessageBox.critical(self, "Khong doc duoc output", err_text)

    # ---------- issue data ----------
    def load_issue_list(self, items):
        self.issue_items_master = [IssueRecord(*x) for x in list(items or [])]
        self.refresh_issue_tree()

    def merge_issue_list(self, items, scanned_files=None):
        new_items = [IssueRecord(*x) for x in list(items or [])]
        scanned_set = {backend.normalize_path_text(x) for x in (scanned_files or []) if backend.nz_str(x)}

        current_master = list(self.issue_items_master)
        if scanned_set:
            current_master = [x for x in current_master if backend.normalize_path_text(x.file_path) not in scanned_set]

        dedup = {}
        for rec in current_master + new_items:
            key = (
                backend.normalize_path_text(rec.file_path),
                rec.normalized_status,
                backend.normalize_simple_text(rec.detail),
            )
            dedup[key] = rec
        self.issue_items_master = list(dedup.values())
        self.refresh_issue_tree()

    def _get_filtered_issue_data(self) -> List[IssueRecord]:
        keyword = backend.nz_str(self.search_edit.text()).lower()
        wanted_status = backend.nz_str(self.status_cb.currentText()).upper().strip() or "ISSUES"
        result = []
        for rec in self.issue_items_master:
            status = rec.normalized_status
            if wanted_status == "ISSUES":
                if status not in ("WARNING", "ERROR"):
                    continue
            elif wanted_status != "ALL" and status != wanted_status:
                continue
            if keyword and keyword not in rec.search_blob:
                continue
            result.append(rec)
        return result

    def refresh_issue_tree(self):
        selected_idx = self.issue_tree.current_row_index() if hasattr(self.issue_tree, "current_row_index") else -1
        selected_key = None
        if 0 <= selected_idx < len(self.filtered_issue_items):
            rec0 = self.filtered_issue_items[selected_idx]
            selected_key = (rec0.normalized_status, rec0.file_name, rec0.detail)

        self.filtered_issue_items = self._get_filtered_issue_data()
        reselection_index = -1
        if selected_key:
            for i, rec in enumerate(self.filtered_issue_items):
                if selected_key == (rec.normalized_status, rec.file_name, rec.detail):
                    reselection_index = i
                    break

        self.issue_tree.set_records(self.filtered_issue_items, reselection_index)
        if self.filtered_issue_items:
            self._on_issue_selected()
        else:
            self._show_issue_detail(None)
        self._refresh_summary_labels()

    def _on_issue_selected(self):
        rec = self.get_selected_issue(show_warning=False)
        self._show_issue_detail(rec)

    def get_selected_issue(self, show_warning: bool = True) -> Optional[IssueRecord]:
        idx = self.issue_tree.current_row_index() if hasattr(self.issue_tree, "current_row_index") else -1
        if idx < 0 or idx >= len(self.filtered_issue_items):
            if show_warning:
                QMessageBox.warning(self, "Chua chon", "Hay chon 1 dong trong bang.")
            return None
        return self.filtered_issue_items[idx]

    def open_selected_issue_file(self):
        rec = self.get_selected_issue()
        if not rec:
            return
        try:
            backend.open_file_with_default_app(rec.file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Khong mo duoc file", str(exc))

    def open_selected_issue_folder(self):
        rec = self.get_selected_issue()
        if not rec:
            return
        try:
            backend.open_folder_and_select_file(rec.file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Khong mo duoc thu muc", str(exc))

    def copy_selected_issue_path(self):
        rec = self.get_selected_issue()
        if not rec:
            return
        QApplication.clipboard().setText(rec.file_path)
        self._append_log(f"Da copy duong dan: {rec.file_path}")

    def toggle_repair_options(self):
        visible = not self.repair_group.isVisible()
        self.repair_group.setVisible(visible)
        self.btn_repair_toggle.setText("Repair options ▾" if visible else "Repair options ▸")

    # ---------- output reading ----------
    def _parse_log_run_time(self, value):
        if isinstance(value, backend.datetime):
            return value
        text = backend.nz_str(value)
        if not text:
            return backend.datetime.min
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
            try:
                return backend.datetime.strptime(text, fmt)
            except Exception:
                pass
        return backend.datetime.min

    def _norm_issue_detail_key(self, detail: str) -> str:
        s = backend.normalize_simple_text(detail).upper()
        s = s.replace(" | ", "|").replace(" :", ":").replace(": ", ":")
        s = s.replace(" ; ", ";").replace("; ", ";")
        return s

    def _build_log_issue_key(self, file_path: str, file_name: str, detail: str) -> str:
        p = backend.normalize_path_text(file_path)
        f = backend.normalize_simple_text(file_name).upper()
        d = self._norm_issue_detail_key(detail)
        if p:
            return f"PATH::{p}||DETAIL::{d}"
        return f"FILE::{f}||DETAIL::{d}"

    def _choose_better_log_issue(self, old_rec: Dict, new_rec: Dict) -> Dict:
        if new_rec["run_dt"] > old_rec["run_dt"]:
            return new_rec
        if new_rec["run_dt"] < old_rec["run_dt"]:
            return old_rec
        old_step = 3 if backend.nz_str(old_rec.get("step")).upper() == "SUMMARY" else 2 if backend.nz_str(old_rec.get("step")).upper() == "CHECK" else 1
        new_step = 3 if backend.nz_str(new_rec.get("step")).upper() == "SUMMARY" else 2 if backend.nz_str(new_rec.get("step")).upper() == "CHECK" else 1
        if new_step > old_step:
            return new_rec
        if new_step < old_step:
            return old_rec
        old_res = 2 if backend.nz_str(old_rec.get("status")).upper() == "ERROR" else 1 if backend.nz_str(old_rec.get("status")).upper() == "WARNING" else 0
        new_res = 2 if backend.nz_str(new_rec.get("status")).upper() == "ERROR" else 1 if backend.nz_str(new_rec.get("status")).upper() == "WARNING" else 0
        if new_res > old_res:
            return new_rec
        if new_res < old_res:
            return old_rec
        old_score = sum(1 for x in (old_rec.get("file_path"), old_rec.get("detail"), old_rec.get("file"), old_rec.get("run_time")) if backend.nz_str(x))
        new_score = sum(1 for x in (new_rec.get("file_path"), new_rec.get("detail"), new_rec.get("file"), new_rec.get("run_time")) if backend.nz_str(x))
        if new_score > old_score:
            return new_rec
        return old_rec

    def _read_ok_items_from_sub_detail(self, wb):
        if "SUB_DETAIL" not in wb.sheetnames:
            return []
        ws = wb["SUB_DETAIL"]
        row_iter = ws.iter_rows(values_only=True)
        header_row = next(row_iter, None)
        if not header_row:
            return []
        header_map = {backend.nz_str(v): idx for idx, v in enumerate(header_row)}
        status_idx = header_map.get("Status")
        error_idx = header_map.get("Error")
        path_idx = header_map.get("File Path")
        if status_idx is None or error_idx is None or path_idx is None:
            raise ValueError("Sheet SUB_DETAIL thieu cot Status / Error / File Path.")
        items = []
        for row in row_iter:
            if not row:
                continue
            status = backend.nz_str(row[status_idx] if status_idx < len(row) else "").upper().strip()
            if status != "OK":
                continue
            detail = backend.nz_str(row[error_idx] if error_idx < len(row) else "")
            path = backend.nz_str(row[path_idx] if path_idx < len(row) else "")
            if not path:
                continue
            items.append((path, "OK", detail))
        return items

    def _read_issue_items_from_log(self, wb):
        if "LOG_SUB_DETAIL" not in wb.sheetnames:
            return []
        ws = wb["LOG_SUB_DETAIL"]
        row_iter = ws.iter_rows(values_only=True)
        header_row = next(row_iter, None)
        if not header_row:
            return []
        header_map = {backend.nz_str(v): idx for idx, v in enumerate(header_row)}
        run_idx = header_map.get("Run Time")
        file_idx = header_map.get("File")
        step_idx = header_map.get("Step")
        result_idx = header_map.get("Result")
        detail_idx = header_map.get("Detail")
        path_idx = header_map.get("File Path")
        if any(x is None for x in (run_idx, file_idx, step_idx, result_idx, detail_idx, path_idx)):
            raise ValueError("Sheet LOG_SUB_DETAIL thieu cot Run Time / File / Step / Result / Detail / File Path.")
        best = {}
        for row in row_iter:
            if not row:
                continue
            step = backend.nz_str(row[step_idx] if step_idx < len(row) else "").upper().strip()
            status = backend.nz_str(row[result_idx] if result_idx < len(row) else "").upper().strip()
            if status not in ("WARNING", "ERROR"):
                continue
            if step not in ("CHECK", "SUMMARY"):
                continue
            file_name = backend.nz_str(row[file_idx] if file_idx < len(row) else "")
            detail = backend.nz_str(row[detail_idx] if detail_idx < len(row) else "")
            file_path = backend.nz_str(row[path_idx] if path_idx < len(row) else "")
            run_time = backend.nz_str(row[run_idx] if run_idx < len(row) else "")
            key = self._build_log_issue_key(file_path, file_name, detail)
            rec = {
                "run_time": run_time,
                "run_dt": self._parse_log_run_time(row[run_idx] if run_idx < len(row) else ""),
                "file": file_name,
                "step": step,
                "status": status,
                "detail": detail,
                "file_path": file_path,
            }
            old = best.get(key)
            best[key] = rec if old is None else self._choose_better_log_issue(old, rec)
        final = sorted(best.values(), key=lambda x: (x["run_dt"], 1 if x["status"] == "ERROR" else 0, backend.normalize_simple_text(x["file"]).upper(), self._norm_issue_detail_key(x["detail"])), reverse=True)
        return [(rec["file_path"], rec["status"], rec["detail"]) for rec in final if backend.nz_str(rec["file_path"]) or backend.nz_str(rec["file"])]

    def _read_issue_items_from_output(self, output_path):
        snapshot_items = []
        try:
            if hasattr(backend, 'load_issue_snapshot_items'):
                snapshot_items = list(backend.load_issue_snapshot_items(output_path) or [])
        except Exception:
            snapshot_items = []

        wb = backend.openpyxl.load_workbook(output_path, read_only=True, data_only=True)
        try:
            ok_items = self._read_ok_items_from_sub_detail(wb)
            if snapshot_items:
                return ok_items + snapshot_items
            issue_items = self._read_issue_items_from_log(wb)
            return ok_items + issue_items
        finally:
            wb.close()

    # ---------- summary/config ----------
    def _refresh_summary_labels(self):
        total = len(self.issue_items_master)
        ok = sum(1 for x in self.issue_items_master if x.normalized_status == "OK")
        warning = sum(1 for x in self.issue_items_master if x.normalized_status == "WARNING")
        error = sum(1 for x in self.issue_items_master if x.normalized_status == "ERROR")
        self._update_badges(total, ok, warning, error)

    def _load_config(self):
        try:
            if not os.path.exists(CONFIG_FILE):
                return
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.folder_path = backend.nz_str(cfg.get("folder_path", ""))
            self.output_path = backend.nz_str(cfg.get("output_path", ""))
            self.chk_fix_folder.setChecked(bool(cfg.get("fix_by_folder", True)))
            self.chk_fix_manual.setChecked(bool(cfg.get("fix_by_manual", False)))
            self.ent_manual_inv.setText(backend.nz_str(cfg.get("manual_invoice_no", "")))
            self.ent_manual_date.setText(backend.nz_str(cfg.get("manual_invoice_date", "")))
            self.ent_manual_dest.setText(backend.nz_str(cfg.get("manual_destination", "")))
            self.issue_tree.apply_column_widths(cfg.get("issue_table_widths", IssueGridTable.DEFAULT_WIDTHS))
            self._refresh_labels()
        except Exception as exc:
            self._append_log(f"Khong load duoc config: {exc}")

    def _schedule_issue_column_width_save(self):
        try:
            if hasattr(self, "issue_column_width_timer"):
                self.issue_column_width_timer.start(600)
        except Exception:
            pass

    def _save_issue_column_widths_silent(self):
        try:
            cfg = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f) or {}
            cfg["issue_table_widths"] = self.issue_tree.column_widths()
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def closeEvent(self, event):
        self._save_issue_column_widths_silent()
        super().closeEvent(event)

    def save_config(self):
        self._commit_path_inputs()
        try:
            cfg = {
                "folder_path": self.folder_path,
                "output_path": self.output_path,
                "fix_by_folder": self.chk_fix_folder.isChecked(),
                "fix_by_manual": self.chk_fix_manual.isChecked(),
                "manual_invoice_no": self.ent_manual_inv.text().strip(),
                "manual_invoice_date": self.ent_manual_date.text().strip(),
                "manual_destination": self.ent_manual_dest.text().strip(),
                "issue_table_widths": self.issue_tree.column_widths(),
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self._append_log(f"Da luu mac dinh vao: {CONFIG_FILE}")
            QMessageBox.information(self, "Da luu", "Da luu cau hinh mac dinh.")
        except Exception as exc:
            QMessageBox.critical(self, "Khong luu duoc", str(exc))

    def reset_config(self):
        self.chk_fix_folder.setChecked(True)
        self.chk_fix_manual.setChecked(False)
        self.ent_manual_inv.clear()
        self.ent_manual_date.clear()
        self.ent_manual_dest.clear()
        self._append_log("Da reset tuy chon Repair options tren giao dien.")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = DashboardWindow()
    win.show()
    win.raise_()
    win.activateWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
