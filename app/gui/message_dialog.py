from PyQt5.QtWidgets import QDialog, QHBoxLayout, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt, QTimer


class MessageDialog(QDialog):
    def __init__(self, title: str, text: str, icon: str = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        # 非模态，支持点击外部关闭
        try:
            self.setModal(False)
            self.setWindowModality(Qt.NonModal)
        except Exception:
            pass
        layout = QHBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 左侧图标（可选）：使用 Emoji，字号与正文接近稍大
        if icon:
            icon_label = QLabel(icon)
            icon_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            icon_label.setStyleSheet("font-size: 15px;")
            layout.addWidget(icon_label, alignment=Qt.AlignTop)

        # 中间文字区域
        text_label = QLabel(text)
        text_label.setWordWrap(True)
        text_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        layout.addWidget(text_label)

        # 去掉 OK 按钮：改为点击任意位置左键即可关闭

        self.setLayout(layout)

        # 3秒后自动关闭
        try:
            QTimer.singleShot(3000, self.accept)
        except Exception:
            pass

    def mousePressEvent(self, event):
        # 左键点击弹框内部关闭对话框
        try:
            if event.button() == Qt.LeftButton:
                self.accept()
                return
        except Exception:
            pass
        super().mousePressEvent(event)

    def focusOutEvent(self, event):
        # 失去焦点（点击弹框外部）自动关闭
        try:
            self.accept()
            return
        except Exception:
            pass
        return super().focusOutEvent(event)

    @staticmethod
    def info(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="ℹ️", parent=parent)
        try:
            dlg.show()
        except Exception:
            try:
                dlg.exec_()
            except Exception:
                pass

    @staticmethod
    def warning(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="⚠️", parent=parent)
        try:
            dlg.show()
        except Exception:
            try:
                dlg.exec_()
            except Exception:
                pass

    @staticmethod
    def error(parent, title: str, text: str):
        dlg = MessageDialog(title, text, icon="❌", parent=parent)
        try:
            dlg.show()
        except Exception:
            try:
                dlg.exec_()
            except Exception:
                pass