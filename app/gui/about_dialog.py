from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt


class AboutDialog(QDialog):
    def __init__(self, about_text: str, contact_url: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setObjectName("AboutDialog")
        layout = QVBoxLayout()
        label = QLabel(about_text)
        label.setObjectName("AboutLabel")
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        layout.addWidget(label)
        # 超链接文字：联系我
        link = QLabel(f'<a href="{contact_url}">联系我</a>')
        link.setObjectName("AboutLink")
        link.setTextFormat(Qt.RichText)
        link.setTextInteractionFlags(Qt.TextBrowserInteraction)
        link.setOpenExternalLinks(True)
        link.setAlignment(Qt.AlignLeft)
        layout.addWidget(link)
        self.setLayout(layout)