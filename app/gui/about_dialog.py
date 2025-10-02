from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton


class AboutDialog(QDialog):
    def __init__(self, about_text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        layout = QVBoxLayout()
        label = QLabel(about_text)
        label.setWordWrap(True)
        layout.addWidget(label)
        btn = QPushButton("关闭")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)
        self.setLayout(layout)