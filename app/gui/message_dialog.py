from PyQt5.QtWidgets import QDialog, QHBoxLayout, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt


class MessageDialog(QDialog):
    def __init__(self, title: str, text: str, icon: str = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        layout = QHBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 左侧图标（可选）：使用 Emoji，字号与正文接近稍大
        if icon:
            icon_label = QLabel(icon)
            icon_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            icon_label.setStyleSheet("font-size: 16px;")
            layout.addWidget(icon_label, alignment=Qt.AlignTop)

        # 中间文字区域
        text_label = QLabel(text)
        text_label.setWordWrap(True)
        text_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        layout.addWidget(text_label)

        # 去掉 OK 按钮：改为点击任意位置左键即可关闭

        self.setLayout(layout)
        self.setModal(True)

    def mousePressEvent(self, event):
        # 左键点击任意位置关闭对话框
        try:
            if event.button() == Qt.LeftButton:
                self.accept()
                return
        except Exception:
            pass
        super().mousePressEvent(event)

    @staticmethod
    def info(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="ℹ️", parent=parent)
        dlg.exec_()

    @staticmethod
    def warning(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="⚠️", parent=parent)
        dlg.exec_()

    @staticmethod
    def error(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="❌", parent=parent)
        dlg.exec_()