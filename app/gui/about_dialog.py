from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel
from PyQt5.QtCore import Qt


class AboutDialog(QDialog):
    def __init__(self, about_text: str, contact_url: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setObjectName("AboutDialog")
        layout = QVBoxLayout()
        # 解析为富文本：标签加粗、值着色，并分行显示
        try:
            import re
            from html import escape
            raw = about_text or ""
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
            ver = auth = usage = title = None
            for l in lines:
                m = re.match(r"^(版本)[:：]\s*(.+)$", l)
                if m:
                    ver = m.group(2)
                    continue
                m = re.match(r"^(作者)[:：]\s*(.+)$", l)
                if m:
                    auth = m.group(2)
                    continue
                m = re.match(r"^(使用说明)[:：]\s*(.+)$", l)
                if m:
                    usage = m.group(2)
                    continue
            for l in lines:
                if not re.match(r"^(版本|作者|使用说明)[:：]", l):
                    title = l
                    break
            parts = []
            if title:
                parts.append(f'<h3 style="margin:0 0 8px 0;">{escape(title)}</h3>')
            if ver is not None:
                parts.append(f'<div><b>版本：</b><span style="color:#2079DA;">{escape(ver)}</span></div>')
            if auth is not None:
                parts.append(f'<div><b>作者：</b><span style="color:#2079DA;">{escape(auth)}</span></div>')
            if usage is not None:
                parts.append(f'<div style="margin-top:8px;"><b>使用说明：</b><span style="color:#3B82F6;">{escape(usage)}</span></div>')
            text = "".join(parts) if parts else escape(raw).replace("\n", "<br/>")
        except Exception:
            text = about_text
        label = QLabel(text)
        label.setObjectName("AboutLabel")
        label.setWordWrap(True)
        label.setTextFormat(Qt.RichText)
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