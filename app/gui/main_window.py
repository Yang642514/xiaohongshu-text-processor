import os
from datetime import datetime
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QTextEdit, QPushButton,
    QAction, QLabel, QHBoxLayout, QSizePolicy
)

from app.core.logger import setup_logger
from app.core.punctuation import normalize_punctuation
from app.core.parser import extract_title_and_points, format_paragraphs, render_processed_template, parse_processed_template, configure_tail_filter
from app.core.excel_writer import write_to_template
from app.core.zipper import make_zip
from app.core.utils import load_settings, save_settings
from app.gui.settings_dialog import SettingsDialog
from app.gui.about_dialog import AboutDialog
from app.gui.message_dialog import MessageDialog


class MainWindow(QMainWindow):
    def __init__(self, settings_path: str):
        super().__init__()
        self.setWindowTitle("文案转模板")

        self.settings_path = settings_path
        self.settings = load_settings(settings_path)
        # 应用尾段过滤配置
        configure_tail_filter(self.settings)

        self.logger = setup_logger(self.settings.get("log_dir", "logs"), self.settings.get("log_level", "INFO"))

        central = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(12)
        self.info_label = QLabel("文案转模板")
        self.info_label.setObjectName("TitleLabel")
        self.info_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        top_row.addWidget(self.info_label)
        # 当前主题标签，显示最近一次处理的文案标题
        self.theme_label = QLabel("当前主题：")
        self.theme_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        top_row.addWidget(self.theme_label)
        top_row.addStretch(1)
        # 右侧提示信息区域
        self.status_label = QLabel("提示：请先点击“处理文案”，再写入模板并压缩")
        self.status_label.setWordWrap(False)
        self.status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        top_row.addWidget(self.status_label)
        layout.addLayout(top_row)

        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("粘贴文案...")
        layout.addWidget(self.text_input)

        # 进度条已移除，改用状态文本提示

        btn_row = QHBoxLayout()
        self.process_btn = QPushButton("处理文案")
        self.process_btn.setFixedHeight(34)
        self.process_btn.clicked.connect(self.process_text)
        self.write_zip_btn = QPushButton("写入模板并压缩")
        self.write_zip_btn.setFixedHeight(34)
        self.write_zip_btn.clicked.connect(self.write_and_zip)
        # 写入功能随时可用
        self.write_zip_btn.setEnabled(True)
        # 新增：清空按钮（清空编辑框与会话标题）
        self.clear_btn = QPushButton("清空")
        self.clear_btn.setFixedHeight(34)
        self.clear_btn.clicked.connect(self.clear_text_and_title)
        btn_row.addWidget(self.process_btn)
        btn_row.addWidget(self.write_zip_btn)
        btn_row.addWidget(self.clear_btn)
        layout.addLayout(btn_row)

        central.setLayout(layout)
        self.setCentralWidget(central)
        # 存储最近一次处理的标题（会话级，初始化为空）
        self.last_title = ""

        # Menu actions
        menubar = self.menuBar()
        settings_action = QAction("设置", self)
        settings_action.triggered.connect(self.open_settings)
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.open_about)
        menubar.addAction(settings_action)
        menubar.addAction(about_action)

    def open_settings(self):
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec_() == dlg.Accepted:
            updated = dlg.get_settings()
            self.settings.update(updated)
            save_settings(self.settings_path, self.settings)
            # 更新解析器配置
            configure_tail_filter(self.settings)
            self.logger.info("设置已更新: %s", updated)

    def open_about(self):
        contact_url = self.settings.get("contact_url", "")
        dlg = AboutDialog(self.settings.get("about_text", ""), contact_url, self)
        dlg.exec_()
        # 关于窗口关闭后不更改标题记忆

    def process_text(self):
        try:
            self.process_btn.setEnabled(False)
            raw = self.text_input.toPlainText().strip()
            if not raw:
                MessageDialog.warning(self, "提示", "请输入文案内容")
                return
            self.logger.info("开始处理文案")

            # 固定逻辑：始终避免嵌套，且尾段过滤按解析器内置开关启用
            already_pairs = parse_processed_template(raw)
            if already_pairs:
                self.logger.info("检测到已处理模板，执行去嵌套的重新排版。分论点数: %d", len(already_pairs))
                points = already_pairs
                title = getattr(self, "last_title", "")
            else:
                title, points = extract_title_and_points(raw)
                self.logger.info("解析标题: %s, 分论点数: %d", title, len(points))
                # 记忆标题（用于命名与再次处理场景）
                self.last_title = title
            # 更新当前主题标签（使用记忆的标题）
            self.theme_label.setText(f"当前主题：{self.last_title}")

            # 上述分支已完成解析，这里不重复解析
            # 解析完成，开始规范化与排版
            # 标点与排版（再次处理也只对内容做规整，模板不嵌套）
            formatted_points = []
            for pt_title, pt_content in points:
                norm_title = normalize_punctuation(pt_title)
                norm_content = normalize_punctuation(pt_content)
                fmt_content = format_paragraphs(
                    norm_content,
                    # 不限制行数：避免删除文案内容
                    max_lines=self.settings.get("max_lines_per_paragraph", 0),
                    # 取消按字符折行：设置为0表示不做字符切分，仅按句号分行
                    max_chars=self.settings.get("max_chars_per_line", 0)
                )
                formatted_points.append((norm_title, fmt_content))

            # 渲染处理模板并显示供用户微调
            processed = render_processed_template(formatted_points)
            self.text_input.setPlainText(processed)
            MessageDialog.info(self, "完成", "已生成处理模板（去嵌套），可微调后写入模板")
            # 更新提示（写入功能随时可用）
            self.status_label.setText("处理完成：现在或稍后均可写入模板并压缩")
            self.process_btn.setEnabled(True)

        except Exception as e:
            self.logger.exception("处理失败: %s", e)
            MessageDialog.error(self, "错误", f"处理失败：{e}")
            self.process_btn.setEnabled(True)

    def write_and_zip(self):
        try:
            # 写入功能随时可用：不禁用按钮
            # 解析编辑框中的处理模板
            processed_text = self.text_input.toPlainText().strip()
            if not processed_text:
                # 空文本也允许写入：提示用户需提供内容（保留弹窗以避免空文件）
                MessageDialog.warning(self, "提示", "编辑框为空，请粘贴或处理文案后再写入")
                return
            if not getattr(self, "last_title", ""):
                # 未有标题：从当前文本尝试提取一次标题用于命名
                tmp_title, _ = extract_title_and_points(processed_text)
                self.last_title = tmp_title or self.last_title or "输出"
            pairs = parse_processed_template(processed_text)
            # 若不是处理模板或解析为空，自动执行一次处理流程（随时可用）
            if not pairs:
                self.logger.info("当前内容非处理模板或为空，自动执行处理流程以写入")
                title, points = extract_title_and_points(processed_text)
                self.last_title = title or self.last_title or "输出"
                # 做与处理流程一致的排版规整
                formatted_points = []
                for pt_title, pt_content in points:
                    norm_title = normalize_punctuation(pt_title)
                    norm_content = normalize_punctuation(pt_content)
                    fmt_content = format_paragraphs(
                        norm_content,
                        max_lines=self.settings.get("max_lines_per_paragraph", 0),
                        max_chars=self.settings.get("max_chars_per_line", 0)
                    )
                    formatted_points.append((norm_title, fmt_content))
                pairs = formatted_points
            # 更新当前主题（写入前确保同步显示）
            self.theme_label.setText(f"当前主题：{self.last_title}")
            self.logger.info("处理模板解析出分论点: %d", len(pairs))

            # 写入 Excel（移除标题列，使用页面/文本_1/文本_2）
            out_xlsx = write_to_template(
                template_path=self.settings.get("template_excel_path"),
                output_dir=self.settings.get("zip_output_dir", "output"),
                title=self.last_title or "",
                points=pairs,
                column_map={
                    "page_column": self.settings.get("page_column", "页面"),
                    "point_title_column": self.settings.get("point_title_column", "文本_1"),
                    "point_content_column": self.settings.get("point_content_column", "文本_2"),
                    # 新增：文本_3列与默认内容
                    "extra_text_column": self.settings.get("extra_text_column", "文本_3"),
                    "extra_text_default": self.settings.get("extra_text_default", "内容仅供参考，身体不适请及时就医!")
                }
            )
            self.logger.info("已生成Excel: %s", out_xlsx)

            # 生成 ZIP
            zip_path = make_zip(out_xlsx, self.settings.get("zip_output_dir", "output"))
            self.logger.info("已生成ZIP: %s", zip_path)

            # 根据设置删除Excel，仅保留zip
            if self.settings.get("delete_excel_after_zip", True):
                try:
                    os.remove(out_xlsx)
                    self.logger.info("已删除Excel，仅保留zip")
                except Exception as de:
                    self.logger.warning("删除Excel失败: %s", de)

            MessageDialog.info(self, "完成", f"已生成压缩包：\n{zip_path}")
            self.status_label.setText(f"完成：压缩包已生成 → {os.path.basename(zip_path)}")
            # 按钮保持可用
            self.write_zip_btn.setEnabled(True)

        except Exception as e:
            self.logger.exception("写入压缩失败: %s", e)
            MessageDialog.error(self, "错误", f"写入压缩失败：{e}")
            self.write_zip_btn.setEnabled(True)

    def clear_text_and_title(self):
        """清空编辑框与会话标题，作为一次处理会话的结束。"""
        self.text_input.clear()
        self.last_title = ""
        self.theme_label.setText("当前主题：")
        self.status_label.setText("已清空：可粘贴新文案进行处理")