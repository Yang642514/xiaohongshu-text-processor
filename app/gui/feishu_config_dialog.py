from typing import Dict, Optional

import re
import requests
from urllib.parse import urlparse, parse_qs

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QHBoxLayout,
    QLabel, QSizePolicy, QWidget, QComboBox
)
from app.gui.message_dialog import MessageDialog

from app.core.utils import save_settings
from app.core.config_manager import ConfigManager


class FeishuConfigDialog(QDialog):
    """飞书配置对话框：

    - 连接配置（App ID、App Secret、Bitable 链接），通过链接自动派生 app_token/table_id
    - 其他配置（筛选字段名、筛选字段值、文案字段名、分页大小）
    - 保存到 settings 并返回
    """

    def __init__(self, settings: Dict, settings_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("飞书配置")
        # 使用共享设置对象
        self.settings = ConfigManager.instance().settings
        self.settings_path = settings_path

        layout = QVBoxLayout()
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setFormAlignment(Qt.AlignLeft)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        # 连接配置
        self.app_id = QLineEdit(self.settings.get("feishu_app_id", ""))
        self.app_id.setFixedHeight(28)
        self.app_id.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.app_secret = QLineEdit(self.settings.get("feishu_app_secret", ""))
        self.app_secret.setFixedHeight(28)
        self.app_secret.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.bitable_link = QLineEdit(self.settings.get("feishu_bitable_link", ""))
        self.bitable_link.setPlaceholderText("粘贴飞书多维表格完整链接")
        self.bitable_link.setFixedHeight(28)
        self.bitable_link.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        form.addRow(self._vlabel("App ID"), self.app_id)
        form.addRow(self._vlabel("App Secret"), self.app_secret)
        form.addRow(self._vlabel("多维表格链接"), self.bitable_link)

        # 其他配置
        self.status_field = QComboBox()
        self.status_field.setEditable(True)
        self.status_field.addItem(self.settings.get("feishu_status_field_name", "笔记状态"))
        self.done_value = QComboBox()
        self.done_value.setEditable(True)
        self.done_value.addItem(self.settings.get("feishu_status_done_value", "已完成"))
        self.content_field = QComboBox()
        self.content_field.setEditable(True)
        self.content_field.addItem(self.settings.get("feishu_content_field_name", "去敏感词"))
        self.page_size = QLineEdit(str(self.settings.get("feishu_page_size", 50)))
        form.addRow(self._vlabel("筛选字段名"), self.status_field)
        form.addRow(self._vlabel("筛选字段值"), self.done_value)
        form.addRow(self._vlabel("文案字段名"), self.content_field)
        form.addRow(self._vlabel("分页大小"), self.page_size)

        layout.addLayout(form)

        # 底部按钮
        btns = QHBoxLayout()
        btns.setSpacing(12)
        btns.setAlignment(Qt.AlignCenter)
        btn_test = QPushButton("测试连接性")
        btn_test.setFixedHeight(28)
        btn_test.clicked.connect(self._test_connectivity)
        btn_save = QPushButton("保存")
        btn_save.setFixedHeight(28)
        btn_save.clicked.connect(self._save)
        btns.addWidget(btn_test)
        btns.addWidget(btn_save)
        layout.addLayout(btns)

        self.setLayout(layout)
        self.resize(560, 360)
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)

        self._fields_meta: Dict[str, Dict] = {}
        def _on_status_changed(text: str):
            try:
                meta = self._fields_meta.get(text) or {}
                opts = []
                prop = meta.get("property") or {}
                if isinstance(prop, dict):
                    options = prop.get("options") or []
                    if isinstance(options, list):
                        for o in options:
                            name = o.get("name") if isinstance(o, dict) else None
                            if name:
                                opts.append(str(name))
                prev = self.done_value.currentText().strip()
                self.done_value.clear()
                if opts:
                    for n in opts:
                        self.done_value.addItem(n)
                # 保留原有值作为可选
                if prev and prev not in opts:
                    self.done_value.addItem(prev)
                # 设置回当前文本
                if prev:
                    idx = self.done_value.findText(prev)
                    self.done_value.setCurrentIndex(idx if idx >= 0 else 0)
            except Exception:
                pass
        self.status_field.currentTextChanged.connect(_on_status_changed)

    def _vlabel(self, text: str) -> QLabel:
        lab = QLabel(text)
        lab.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lab.setFixedHeight(24)
        return lab

    @staticmethod
    def _derive_from_link(link: str) -> Optional[Dict[str, str]]:
        """尽力从链接中解析 app_token 与 table_id（base 形态）。

        规则说明：
        - 若为 base 形态链接（路径含 /base/），app_token 通常为路径段中形如
          "app..." 或 "bascn..." 的标识（官方示例："appbcbWCzen6D8dezhoCH2RpMAh"）。
        - table_id 通常形如 "tbl..."，可能出现在查询参数（table/table_id）或路径段中。
        - 若为 wiki 形态链接（路径含 /wiki/），官方文档说明无法直接从链接拿到 app_token，
          需调用知识空间接口获取；此处返回 None 提示用户改用 base 链接。
        参考：获取多维表格元数据（飞书服务端文档）。
        """
        if not link:
            return None
        try:
            parsed = urlparse(link)
            qs = parse_qs(parsed.query or "")
            path = parsed.path or ""
            segments = [seg for seg in path.split('/') if seg]

            # wiki 链接不直接包含 app_token
            if any(seg.lower() == "wiki" for seg in segments):
                return None

            # 先从查询参数取 table_id
            table_id = None
            for k in ("table", "table_id", "tbl"):
                v = qs.get(k, [])
                if v:
                    table_id = v[0]
                    break

            # 从路径段识别 app_token/table_id
            app_token = None
            # 优先：若存在 /base/<token> 结构，则取该段为 app_token
            try:
                base_idx = next(i for i, seg in enumerate(segments) if seg.lower() == "base")
                if base_idx + 1 < len(segments) and not app_token:
                    candidate = segments[base_idx + 1]
                    # 排除明显的表或视图段
                    if not re.fullmatch(r"tbl[A-Za-z0-9]{4,}", candidate) and not candidate.lower().startswith("vew"):
                        app_token = candidate
            except StopIteration:
                pass
            for seg in segments:
                if not app_token and re.fullmatch(r"(app|bascn)[A-Za-z0-9]{6,}", seg):
                    app_token = seg
                if not table_id and re.fullmatch(r"tbl[A-Za-z0-9]{4,}", seg):
                    table_id = seg

            # 兜底：全文匹配（避免误识别参数名）
            if not app_token:
                m_app = re.search(r"(app|bascn)[A-Za-z0-9]{6,}", link)
                if m_app:
                    app_token = m_app.group(0)
            if not table_id:
                m_tbl = re.search(r"tbl[A-Za-z0-9]{4,}", link)
                if m_tbl:
                    table_id = m_tbl.group(0)

            if app_token and table_id:
                return {"feishu_bitable_app_token": app_token, "feishu_bitable_table_id": table_id}
        except Exception:
            pass
        return None

    def _save(self):
        try:
            app_id = self.app_id.text().strip()
            app_secret = self.app_secret.text().strip()
            link = self.bitable_link.text().strip()
            status_field = self.status_field.currentText().strip() or "笔记状态"
            done_value = self.done_value.currentText().strip() or "已完成"
            content_field = self.content_field.currentText().strip() or "去敏感词"
            page_size_text = self.page_size.text().strip() or "50"
            try:
                page_size = int(page_size_text)
            except ValueError:
                raise ValueError("分页大小需为整数")

            if not app_id or not app_secret:
                raise ValueError("请填写 App ID 和 App Secret")

            # 尝试从链接派生 app_token/table_id
            derived = self._derive_from_link(link)
            updated = {
                "feishu_app_id": app_id,
                "feishu_app_secret": app_secret,
                "feishu_bitable_link": link,
                "feishu_status_field_name": status_field,
                "feishu_status_done_value": done_value,
                "feishu_content_field_name": content_field,
                "feishu_page_size": page_size,
            }
            if derived:
                updated.update(derived)

            # 合并保存：统一通过 ConfigManager.save()
            self.settings.update(updated)
            try:
                ConfigManager.instance().save()
            except Exception:
                # 回退以避免写盘失败导致未保存
                save_settings(self.settings_path, self.settings)

            MessageDialog.info(self, "成功", "配置已保存")
            self.accept()
        except Exception as e:
            MessageDialog.error(self, "错误", f"保存失败：{e}")

    def _test_connectivity(self):
        try:
            app_id = self.app_id.text().strip()
            app_secret = self.app_secret.text().strip()
            link = self.bitable_link.text().strip()
            if not app_id or not app_secret:
                raise ValueError("请填写 App ID 和 App Secret")

            # 1) 测试鉴权
            auth_url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
            resp = requests.post(auth_url, json={"app_id": app_id, "app_secret": app_secret}, timeout=15)
            data = resp.json()
            if data.get("code") != 0:
                raise RuntimeError(f"鉴权失败: {data}")
            token = data.get("tenant_access_token")
            if not token:
                raise RuntimeError("鉴权成功但未返回 token")

            # 2) 测试表访问（若能派生 app_token/table_id）
            derived = self._derive_from_link(link) or {}
            app_token = derived.get("feishu_bitable_app_token", "")
            table_id = derived.get("feishu_bitable_table_id", "")
            if app_token and table_id:
                fields_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
                headers = {"Authorization": f"Bearer {token}"}
                f_resp = requests.get(fields_url, headers=headers, timeout=15)
                f_data = f_resp.json()
                if f_data.get("code") != 0:
                    raise RuntimeError(f"表访问失败: {f_data}")
                items = f_data.get("data", {}).get("items", [])
                # 填充下拉框（仅在确有字段时）
                try:
                    names = []
                    meta_map: Dict[str, Dict] = {}
                    for it in items:
                        # 兼容不同返回结构：优先使用 name，其次 field_name
                        name = it.get("name") or it.get("field_name")
                        if name:
                            n = str(name)
                            names.append(n)
                            meta_map[n] = it
                    if names:
                        self._fields_meta = meta_map
                        prev_status = self.status_field.currentText().strip()
                        prev_content = self.content_field.currentText().strip()
                        self.status_field.clear()
                        self.content_field.clear()
                        for n in names:
                            self.status_field.addItem(n)
                            self.content_field.addItem(n)
                        if prev_status:
                            idx = self.status_field.findText(prev_status)
                            self.status_field.setCurrentIndex(idx if idx >= 0 else 0)
                        if prev_content:
                            idx2 = self.content_field.findText(prev_content)
                            self.content_field.setCurrentIndex(idx2 if idx2 >= 0 else 0)
                        # 触发一次联动以填充筛选值
                        self.status_field.currentTextChanged.emit(self.status_field.currentText())
                        MessageDialog.info(self, "成功", f"鉴权成功，表访问成功：字段 {len(items)} 个，已填充下拉选项")
                    else:
                        # 无字段返回：保留原有输入，不清空下拉框
                        MessageDialog.info(
                            self,
                            "成功",
                            "鉴权成功，已访问表，但未返回字段列表。可能原因：权限不足、表为空或链接未对应 app_token。"
                        )
                except Exception:
                    # 即使解析失败也不影响鉴权提示
                    MessageDialog.info(self, "成功", f"鉴权成功，表访问成功：但字段解析异常，请手动输入字段名。")
            else:
                MessageDialog.info(
                    self,
                    "成功",
                    "鉴权成功（未识别表信息）。如果是 wiki 链接，无法直接解析 app_token，"
                    "请改用 base 形态链接并确保其中包含形如 app... 与 tbl... 的标识。"
                )
        except Exception as e:
            MessageDialog.error(self, "错误", f"测试失败：{e}")