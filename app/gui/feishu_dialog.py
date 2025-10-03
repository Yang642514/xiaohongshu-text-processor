from typing import List, Dict, Callable
import os
import json
import logging

from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QWidget, QHeaderView, QApplication, QProgressBar
)
from PyQt5.QtGui import QFontMetrics

from app.core.feishu_client import FeishuClient
from app.gui.feishu_config_dialog import FeishuConfigDialog
from app.core.utils import save_settings


class FeishuDialog(QDialog):
    def __init__(self, settings: Dict, on_edit: Callable[[Dict], None], parent=None, settings_path: str = None):
        super().__init__(parent)
        self.settings = settings
        self.on_edit = on_edit
        self.settings_path = settings_path
        self.setWindowTitle("同步飞书数据")
        # 初始尺寸：优先使用记忆尺寸，其次自适应父窗口比例
        self.resize(760, 520)
        try:
            saved_size = self.settings.get("feishu_dialog_size")
            if isinstance(saved_size, dict):
                w = int(saved_size.get("width", self.width()))
                h = int(saved_size.get("height", self.height()))
                if w > 0 and h > 0:
                    self.resize(w, h)
        except Exception:
            pass
        try:
            if parent is not None:
                pgeom = parent.geometry()
                max_w = int(pgeom.width() * 0.9)
                max_h = int(pgeom.height() * 0.9)
                new_w = min(self.width(), max_w)
                new_h = min(self.height(), max_h)
                self.resize(new_w, new_h)
                self.move(pgeom.center().x() - new_w // 2, pgeom.center().y() - new_h // 2)
        except Exception:
            pass

        layout = QVBoxLayout(self)
        # 顶部工具条
        top = QHBoxLayout()
        self.btn_refresh = QPushButton("刷新")
        self.btn_refresh.setObjectName("InlineButton")
        self.btn_config = QPushButton("配置")
        self.btn_config.setObjectName("InlineButton")
        self.info = QLabel("准备就绪")
        self.info.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        # 限制宽度并开启省略显示与点击复制
        self.info.setFixedWidth(320)
        self.info.setCursor(Qt.PointingHandCursor)
        self._info_full = "准备就绪"
        try:
            self.info.setStyleSheet("font-weight: 600; color: #2f3e4d;")
        except Exception:
            pass
        def _copy_info_to_clipboard(event):
            try:
                cb = QApplication.clipboard()
                cb.setText(self._info_full or "")
            except Exception:
                pass
        self.info.mousePressEvent = _copy_info_to_clipboard
        top.addWidget(self.btn_refresh)
        top.addWidget(self.btn_config)
        top.addStretch(1)
        top.addWidget(self.info)
        # 加载指示器：不确定进度条，显示“正在加载”动效
        try:
            self.loading_bar = QProgressBar()
            self.loading_bar.setRange(0, 0)
            self.loading_bar.setFixedWidth(100)
            self.loading_bar.hide()
            top.addWidget(self.loading_bar)
        except Exception:
            self.loading_bar = None
        layout.addLayout(top)

        # 表格
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "文案预览（去敏感词）", "状态", "更新时间", "操作"])
        self.table.verticalHeader().setVisible(False)
        # 控制默认行高，避免操作按钮过大导致表格行溢出
        try:
            self.table.verticalHeader().setDefaultSectionSize(30)
        except Exception:
            pass
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        # 禁止换行以实现单行显示
        try:
            self.table.setWordWrap(False)
        except Exception:
            pass
        # 允许用户手动调整列宽，默认内容列自适应填充，其余可调整；若存在记忆列宽则优先应用并全部可调整
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        try:
            widths = self.settings.get("feishu_dialog_table_widths")
        except Exception:
            widths = None
        if isinstance(widths, list) and len(widths) == 5:
            # 应用记忆宽度，并允许用户继续微调
            try:
                for idx, w in enumerate(widths):
                    self.table.setColumnWidth(idx, int(w))
            except Exception:
                pass
            for c in range(5):
                header.setSectionResizeMode(c, QHeaderView.Interactive)
        else:
            # 默认初始宽度：ID/状态/时间/操作固定，内容列自适应
            base_id, base_status, base_time, base_ops = 56, 96, 128, 120
            try:
                self.table.setColumnWidth(0, base_id)
                self.table.setColumnWidth(2, base_status)
                self.table.setColumnWidth(3, base_time)
                self.table.setColumnWidth(4, base_ops)
            except Exception:
                pass
            header.setSectionResizeMode(0, QHeaderView.Interactive)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            header.setSectionResizeMode(2, QHeaderView.Interactive)
            header.setSectionResizeMode(3, QHeaderView.Interactive)
            header.setSectionResizeMode(4, QHeaderView.Interactive)
        try:
            # 标题栏居中显示
            header.setDefaultAlignment(Qt.AlignCenter)
            header.sectionResized.connect(self._on_section_resized)
        except Exception:
            pass
        # 强化表头与网格样式，区分标题与数据区
        try:
            self.table.setStyleSheet(
                "QHeaderView::section{background:#f5f6f8; padding:4px 8px; font-weight:600; color:#333; border-bottom:1px solid #d9d9de;}\n"
                "QTableWidget{gridline-color:#e5e7eb;}\n"
                "QTableWidget::item{padding:2px;}\n"
            )
        except Exception:
            pass
        layout.addWidget(self.table)

        # 底部关闭按钮移除（按需求）

        self.btn_refresh.clicked.connect(self.load_data)
        self.btn_config.clicked.connect(self.open_config)

        # 首次打开异步加载，避免阻塞对话框显示
        try:
            QTimer.singleShot(0, self.load_data)
        except Exception:
            self.load_data()

        # 后台线程引用，避免 GC 导致线程提前结束
        self._fetch_thread = None

    def _set_info(self, text: str):
        """设置顶部提示：过长文本省略显示，悬浮可见完整内容，点击可复制。"""
        try:
            full = str(text)
            self._info_full = full
            self.info.setToolTip(full)
            fm = QFontMetrics(self.info.font())
            # 预留 12px 内边距，避免边界截断
            width = max(50, self.info.width() - 12)
            elided = fm.elidedText(full, Qt.ElideRight, width)
            self.info.setText(elided)
        except Exception:
            # 兜底直接显示原文本
            self.info.setText(str(text))

    class _FetchThread(QThread):
        result_ready = pyqtSignal(list)
        error = pyqtSignal(str)

        def __init__(self, settings: Dict):
            super().__init__()
            self._settings = settings

        def run(self):
            try:
                client = FeishuClient(self._settings)
                items = client.search_done_records()
                self.result_ready.emit(items)
            except Exception as e:
                self.error.emit(str(e))

    def load_data(self):
        self._set_info("加载中...")
        # 先尝试使用缓存填充，提升首屏体验
        try:
            cached = self.settings.get("feishu_cached_items")
            if isinstance(cached, list) and cached:
                self._fill_table(cached)
        except Exception:
            pass
        # 显示加载指示器
        try:
            if getattr(self, "loading_bar", None):
                self.loading_bar.show()
        except Exception:
            pass
        # 启动后台线程加载，避免 UI 卡顿
        try:
            # 如有旧线程，尝试停止
            if getattr(self, "_fetch_thread", None):
                try:
                    self._fetch_thread.quit()
                    self._fetch_thread.wait(100)
                except Exception:
                    pass
            self._fetch_thread = self._FetchThread(self.settings)
            # 记录启动时间
            import datetime as _dt
            self._fetch_time_str = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            def _on_result(items: List[Dict]):
                try:
                    self._fill_table(items)
                    # 写入缓存供下次快速显示
                    try:
                        self.settings["feishu_cached_items"] = items
                        if self.settings_path:
                            save_settings(self.settings_path, self.settings)
                    except Exception:
                        pass
                    # —— 调试输出：记录刷新获取到的数据 ——
                    try:
                        logger = logging.getLogger("xhs_tool")
                        logger.info("[FeishuDialog] 刷新获取到 %d 条记录", len(items))
                        try:
                            if logger.isEnabledFor(logging.DEBUG):
                                logger.debug("[FeishuDialog] items=%s", json.dumps(items, ensure_ascii=False))
                        except Exception:
                            pass
                        log_dir = self.settings.get("log_dir", "logs") or "logs"
                        os.makedirs(log_dir, exist_ok=True)
                        snapshot_path = os.path.join(log_dir, "feishu_refresh_latest.json")
                        with open(snapshot_path, 'w', encoding='utf-8') as f:
                            json.dump(items, f, ensure_ascii=False, indent=2)
                        self._set_info(f"共 {len(items)} 条 | 快照: {os.path.abspath(snapshot_path)}")
                    except Exception:
                        self._set_info(f"共 {len(items)} 条")
                finally:
                    try:
                        if getattr(self, "loading_bar", None):
                            self.loading_bar.hide()
                    except Exception:
                        pass

            def _on_error(err: str):
                self._set_info(f"错误: {err}")
                try:
                    if getattr(self, "loading_bar", None):
                        self.loading_bar.hide()
                except Exception:
                    pass

            self._fetch_thread.result_ready.connect(_on_result)
            self._fetch_thread.error.connect(_on_error)
            self._fetch_thread.start()
        except Exception as e:
            self._set_info(f"错误: {e}")
            try:
                if getattr(self, "loading_bar", None):
                    self.loading_bar.hide()
            except Exception:
                pass

    def open_config(self):
        # 打开配置对话框，保存后刷新当前 settings 并重新加载
        dlg = FeishuConfigDialog(self.settings, self.settings_path or "", self)
        if dlg.exec_() == dlg.Accepted:
            # 重新加载设置文件以保持与主窗口一致
            from app.core.utils import load_settings
            if self.settings_path:
                self.settings.clear()
                self.settings.update(load_settings(self.settings_path))
            self.load_data()

    def _fill_table(self, items: List[Dict]):
        self.table.setRowCount(0)
        # 读取本地已使用记录ID集合（持久化在 settings）
        try:
            used_records = set(self.settings.get("feishu_used_records", []))
        except Exception:
            used_records = set()
        for i, it in enumerate(items):
            self.table.insertRow(i)
            # 序号ID（1,2,3...）
            id_item = QTableWidgetItem(str(i + 1))
            try:
                id_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            except Exception:
                pass
            self.table.setItem(i, 0, id_item)
            # 文案预览：设置 tooltip 显示完整内容；初次渲染保证不为空
            content_text = str(it.get("content", ""))
            single_line = content_text.replace("\r", " ").replace("\n", " ")
            # 单行省略显示：根据当前列宽截断文本；若列宽尚未就绪，则回退为完整单行文本
            try:
                fm = QFontMetrics(self.table.font())
                col_w = int(self.table.columnWidth(1)) - 16
                if col_w <= 20:
                    display_text = single_line
                else:
                    display_text = fm.elidedText(single_line, Qt.ElideRight, col_w)
            except Exception:
                display_text = single_line
            content_item = QTableWidgetItem(display_text)
            try:
                # 显示内容居中
                content_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            except Exception:
                pass
            content_item.setToolTip(content_text)
            # 存储整条记录以便预览与编辑回调使用
            try:
                content_item.setData(Qt.UserRole, it)
            except Exception:
                pass
            self.table.setItem(i, 1, content_item)
            # 状态：本地记忆“是否已使用”，默认未使用
            rid = str(it.get("record_id", ""))
            status_text = "已使用" if rid in used_records else "未使用"
            self.table.setItem(i, 2, QTableWidgetItem(status_text))
            # 更新时间：显示为本次获取时间
            fetch_ts = getattr(self, "_fetch_time_str", "")
            self.table.setItem(i, 3, QTableWidgetItem(str(fetch_ts)))

            # 操作列
            cell = QWidget()
            h = QHBoxLayout(cell)
            h.setContentsMargins(4, 2, 4, 2)
            try:
                h.setSpacing(4)
                # 操作按钮居中且保留边距，不填满单元格
                h.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            except Exception:
                pass
            btn = QPushButton("去编辑")
            btn.setObjectName("InlineButton")
            # 缩小按钮尺寸，避免溢出单元格
            try:
                btn.setFixedHeight(22)
                btn.setFixedWidth(66)
                # 局部样式覆盖全局 InlineButton 的最小尺寸和内边距
                btn.setStyleSheet("min-height:22px; min-width:66px; padding:2px 6px; font-size:12px;")
            except Exception:
                pass
            def make_cb(payload: Dict, row_index: int):
                def _cb():
                    # 本地标记为已使用并持久化，避免刷新后状态复原
                    try:
                        rid_local = str(payload.get("record_id", ""))
                        if rid_local:
                            used_records.add(rid_local)
                            try:
                                self.table.setItem(row_index, 2, QTableWidgetItem("已使用"))
                            except Exception:
                                pass
                            self.settings["feishu_used_records"] = sorted(list(used_records))
                            if self.settings_path:
                                save_settings(self.settings_path, self.settings)
                    except Exception:
                        pass
                    try:
                        # 更新状态为“已使用”，并关闭窗口后填充编辑框
                        client = FeishuClient(self.settings)
                        client.mark_record_as_used(payload)
                    except Exception:
                        # 状态更新失败不阻塞编辑注入流程
                        pass
                    try:
                        self.accept()
                    finally:
                        try:
                            self.on_edit(payload)
                        except Exception:
                            pass
                return _cb
            btn.clicked.connect(make_cb(it, i))
            try:
                h.addWidget(btn)
            except Exception:
                h.addWidget(btn)
            # 不再添加伸缩项，保持按钮在单元格中居中显示
            self.table.setCellWidget(i, 4, cell)

        # 默认选中第一行
        try:
            if items:
                self.table.setCurrentCell(0, 1)
        except Exception:
            pass
        # 根据最终列宽刷新省略显示
        try:
            # 布局完成后刷新一次，保证省略效果与列宽一致
            QTimer.singleShot(0, self._update_content_ellipses)
        except Exception:
            pass
    def _on_section_resized(self, logical_index: int, old_size: int, new_size: int):
        try:
            self._snapshot_widths()
        except Exception:
            pass
        # 内容列宽变化时刷新省略显示
        try:
            if logical_index == 1:
                self._update_content_ellipses()
        except Exception:
            pass

    def _snapshot_widths(self):
        try:
            widths = [self.table.columnWidth(i) for i in range(5)]
            self.settings["feishu_dialog_table_widths"] = widths
        except Exception:
            pass

    def _update_content_ellipses(self):
        try:
            fm = QFontMetrics(self.table.font())
            col_w = max(0, self.table.columnWidth(1) - 16)
            rows = self.table.rowCount()
            for r in range(rows):
                item = self.table.item(r, 1)
                if not item:
                    continue
                payload = item.data(Qt.UserRole)
                if isinstance(payload, dict):
                    full = str(payload.get("content", ""))
                else:
                    full = item.toolTip() or item.text()
                display_text = fm.elidedText(full.replace("\r", " ").replace("\n", " "), Qt.ElideRight, col_w)
                item.setText(display_text)
                item.setToolTip(full)
        except Exception:
            pass

    def _save_layout(self):
        try:
            # 保存列宽与窗口尺寸
            self._snapshot_widths()
            self.settings["feishu_dialog_size"] = {"width": self.width(), "height": self.height()}
            if self.settings_path:
                save_settings(self.settings_path, self.settings)
        except Exception:
            pass

    def accept(self):
        try:
            self._save_layout()
        except Exception:
            pass
        super().accept()

    def closeEvent(self, event):
        try:
            self._save_layout()
        except Exception:
            pass
        return super().closeEvent(event)