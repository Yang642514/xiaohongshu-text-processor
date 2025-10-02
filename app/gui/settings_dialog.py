from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QFileDialog, QMessageBox, QHBoxLayout, QWidget
)
from openpyxl import load_workbook


class SettingsDialog(QDialog):
    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.settings = settings.copy()

        layout = QVBoxLayout()
        form = QFormLayout()

        self.tpl_edit = QLineEdit(self.settings.get("template_excel_path", ""))
        self.zip_edit = QLineEdit(self.settings.get("zip_output_dir", "output"))
        self.page_col = QLineEdit(self.settings.get("page_column", "页面"))
        self.point_title_col = QLineEdit(self.settings.get("point_title_column", "文本_1"))
        self.point_content_col = QLineEdit(self.settings.get("point_content_column", "文本_2"))
        # 新增：与默认内容
        self.extra_text_col = QLineEdit(self.settings.get("extra_text_column", "文本_3"))
        self.extra_text_default = QLineEdit(self.settings.get("extra_text_default", "内容仅供参考，身体不适请及时就医!"))

        # 以下逻辑已内置：再次处理不嵌套与尾段过滤始终启用；不在设置中展示。

        # 模板路径（右侧内联浏览按钮）
        tpl_row = QHBoxLayout()
        tpl_row.addWidget(self.tpl_edit)
        btn_browse_tpl = QPushButton("浏览")
        btn_browse_tpl.clicked.connect(self._browse_template)
        tpl_row.addWidget(btn_browse_tpl)
        tpl_container = QWidget()
        tpl_container.setLayout(tpl_row)
        form.addRow("模板Excel路径", tpl_container)

        # 压缩包目录（右侧内联浏览按钮）
        zip_row = QHBoxLayout()
        zip_row.addWidget(self.zip_edit)
        btn_browse_zip = QPushButton("浏览")
        btn_browse_zip.clicked.connect(self._browse_zip_dir)
        zip_row.addWidget(btn_browse_zip)
        zip_container = QWidget()
        zip_container.setLayout(zip_row)
        form.addRow("压缩包目录", zip_container)

        form.addRow("页面列名", self.page_col)
        form.addRow("分论点标题列名", self.point_title_col)
        form.addRow("分论点内容列名", self.point_content_col)
        form.addRow("页面注释列名", self.extra_text_col)
        form.addRow("页面注释默认内容", self.extra_text_default)
        # 不展示：再次处理不嵌套、尾段过滤开关，以及尾段关键词相关配置
        
        layout.addLayout(form)

        # 底部仅保留“检测模板”和“保存”按钮，并适当增大高度与右对齐
        btns = QHBoxLayout()
        btn_check = QPushButton("检测模板")
        btn_check.setFixedHeight(40)
        btn_check.clicked.connect(self._check_template)
        btn_ok = QPushButton("保存")
        btn_ok.setFixedHeight(40)
        btn_ok.clicked.connect(self.accept)
        btns.addStretch(1)
        btns.addWidget(btn_check)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)

        self.setLayout(layout)

    def _browse_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择模板Excel", "", "Excel Files (*.xlsx)")
        if path:
            self.tpl_edit.setText(path)

    def _browse_zip_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择压缩包目录", "")
        if path:
            self.zip_edit.setText(path)

    def get_settings(self) -> dict:
        return {
            "template_excel_path": self.tpl_edit.text().strip(),
            "zip_output_dir": self.zip_edit.text().strip(),
            "page_column": self.page_col.text().strip() or "页面",
            "point_title_column": self.point_title_col.text().strip() or "文本_1",
            "point_content_column": self.point_content_col.text().strip() or "文本_2",
            "extra_text_column": self.extra_text_col.text().strip() or "文本_3",
            "extra_text_default": self.extra_text_default.text().strip() or "内容仅供参考，身体不适请及时就医!",
        }

    def _check_template(self):
        path = self.tpl_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "提示", "请先选择模板Excel路径")
            return
        try:
            wb = load_workbook(path)
            ws = wb.active
            need_cols = [
                self.page_col.text().strip() or "页面",
                self.point_title_col.text().strip() or "文本_1",
                self.point_content_col.text().strip() or "文本_2",
                self.extra_text_col.text().strip() or "文本_3",
            ]
            header_row = None
            for r in range(1, min(10, ws.max_row) + 1):
                row_vals = [c.value for c in ws[r]]
                if row_vals and all(col in row_vals for col in need_cols):
                    header_row = r
                    break
            if header_row is None:
                QMessageBox.critical(self, "模板不可用", f"未在前10行找到包含列：{', '.join(need_cols)} 的表头")
            else:
                QMessageBox.information(self, "模板可用", f"模板检测通过（表头行：第{header_row}行）")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法读取模板：{e}")