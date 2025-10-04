from typing import List, Tuple, Dict, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QCursor, QPixmap, QImageReader
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QPushButton, QTableWidget, QTableWidgetItem,
    QWidget, QSizePolicy, QHeaderView, QFrame
)

from app.core.utils import save_settings
from app.core.config_manager import ConfigManager


class ImageDialog(QDialog):
    """配图页面（可选功能）：

    - 顶部区块：图片来源与裁剪/抠图配置，状态提示与“保存配置”按钮
    - 主表格：每行一个分论点，含标题、可编辑关键词、缩略图预览、操作按钮、状态
    - 底部操作：批量自动配图、全部抠图、全部裁剪、写入模板并压缩（配图版）、返回
    """

    def __init__(self, settings: Dict, points: List[Tuple[str, str]], settings_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配图页面")
        # 使用共享设置对象，避免副本导致覆盖
        self.settings = settings
        self.settings_path = settings_path
        self.points = points or []
        # 记录每行本地图片路径，供写入模板与打包使用
        self.local_images: List[Optional[str]] = [None] * len(self.points)

        layout = QVBoxLayout()
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # 顶部配置与状态
        top = QWidget()
        top_layout = QVBoxLayout(top)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(6)

        form = QFormLayout()
        # 统一左侧标签列宽，并居中显示
        form.setLabelAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        form.setFormAlignment(Qt.AlignLeft)
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(4)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        # 计算“关键词分析”的宽度作为统一列宽基准
        try:
            fm = self.fontMetrics()
            self._label_w = max(88, fm.width("关键词分析：") + 16)
        except Exception:
            self._label_w = 100

        def _col_label(text: str) -> QLabel:
            lbl = QLabel(text)
            lbl.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            try:
                lbl.setFixedWidth(self._label_w)
                lbl.setFixedHeight(24)
            except Exception:
                pass
            return lbl

        # 第一行：图片来源 + API Key（同一行）
        self.image_source = QComboBox()
        self.image_source.addItems(["pixabay", "pexels", "ai"])
        self.image_source.setCurrentText(self.settings.get("image_source", "pixabay"))
        try:
            self.image_source.setFixedHeight(24)
        except Exception:
            pass
        # 初始化按来源读取已保存的 API Key
        try:
            _init_source = self.settings.get("image_source", "pixabay")
            _init_key = self.settings.get(f"{_init_source}_api_key", self.settings.get("image_api_key", ""))
        except Exception:
            _init_key = self.settings.get("image_api_key", "")
        self.image_api_key = QLineEdit(_init_key)
        self.image_api_key.setPlaceholderText("按来源填写APIKey")
        self.image_api_key.setFixedHeight(24)
        self.image_api_key.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.image_api_key.setMaximumWidth(320)
        # 跟踪并联动来源变化时的 API Key 显示
        self._last_image_source = self.image_source.currentText().strip()
        try:
            self.image_source.currentTextChanged.connect(self._on_image_source_changed)
        except Exception:
            pass
        src_row = QWidget()
        src_h = QHBoxLayout(src_row)
        src_h.setContentsMargins(0, 0, 0, 0)
        src_h.setSpacing(8)
        src_h.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        src_h.addWidget(self.image_source)
        # 可切换显隐的通用 API Key 标签
        self.lbl_api_key = QLabel("API Key:")
        src_h.addWidget(self.lbl_api_key)
        src_h.addWidget(self.image_api_key)
        src_h.addStretch(1)
        # 第一行：左侧统一列标签 + 同行来源与 API Key
        form.addRow(_col_label("图片来源:"), src_row)
        # 仅显示通用 API Key 输入框
        try:
            self.image_api_key.setVisible(True)
            self.lbl_api_key.setVisible(True)
        except Exception:
            pass

        # 第二行：抠图相关设置 + API Key
        self.remove_bg_enabled = QCheckBox("启用抠图")
        self.remove_bg_enabled.setChecked(bool(self.settings.get("remove_bg_enabled", True)))
        self.remove_bg_method = QComboBox()
        self.remove_bg_method.addItems(["rembg", "api"])
        self.remove_bg_method.setCurrentText(self.settings.get("remove_bg_method", "rembg"))
        try:
            self.remove_bg_method.setFixedHeight(24)
        except Exception:
            pass
        self.remove_bg_api_key = QLineEdit(self.settings.get("remove_bg_api_key", ""))
        self.remove_bg_api_key.setPlaceholderText("抠图 API Key")
        self.remove_bg_api_key.setFixedHeight(24)
        self.remove_bg_api_key.setMaximumWidth(260)

        # 自动裁剪与尺寸/模式
        self.auto_crop_enabled = QCheckBox("自动裁剪")
        self.auto_crop_enabled.setChecked(bool(self.settings.get("auto_crop_enabled", True)))
        self.crop_size = QLineEdit(self.settings.get("crop_size", "1024x1024"))
        self.crop_size.setFixedHeight(24)
        self.crop_size.setMaximumWidth(120)
        self.crop_mode = QComboBox()
        self.crop_mode.addItems(["cover", "contain"])
        self.crop_mode.setCurrentText(self.settings.get("crop_mode", "cover"))
        try:
            self.crop_mode.setFixedHeight(24)
        except Exception:
            pass
        self.contain_bg_color = QLineEdit(self.settings.get("contain_bg_color", "#FFFFFF"))
        self.contain_bg_color.setFixedHeight(24)
        self.contain_bg_color.setMaximumWidth(120)

        # 预览缩略图尺寸与输出格式
        self.thumbnail_size = QLineEdit(self.settings.get("thumbnail_size", "160x120"))
        self.thumbnail_size.setFixedHeight(24)
        self.thumbnail_size.setMaximumWidth(120)
        self.image_format = QComboBox()
        self.image_format.addItems(["png", "jpg"])
        self.image_format.setCurrentText(self.settings.get("image_format", "png"))
        try:
            self.image_format.setFixedHeight(24)
        except Exception:
            pass

        # 每论点图片数（保留作为扩展）
        self.image_per_point = QSpinBox()
        self.image_per_point.setRange(1, 5)
        self.image_per_point.setValue(int(self.settings.get("image_per_point", 1)))
        self.image_per_point.setFixedHeight(24)

        # 关键词分析来源（本地/AI）控件（需在使用前初始化）
        self.keyword_source = QComboBox()
        self.keyword_source.addItems(["本地分析", "AI分析"])
        self.keyword_source.setCurrentText(self.settings.get("keyword_source", "本地分析"))
        try:
            self.keyword_source.setFixedHeight(24)
        except Exception:
            pass
        # 切换关键词来源时，重新为空/标题同名的行填充本地生成的关键词
        try:
            self.keyword_source.currentTextChanged.connect(self._on_keyword_source_changed)
        except Exception:
            pass
        self.keyword_ai_model = QLineEdit(self.settings.get("keyword_ai_model", ""))
        self.keyword_ai_model.setPlaceholderText("模型名称")
        self.keyword_ai_model.setFixedHeight(24)
        self.keyword_ai_model.setMaximumWidth(200)
        self.keyword_ai_api_key = QLineEdit(self.settings.get("keyword_ai_api_key", ""))
        self.keyword_ai_api_key.setPlaceholderText("大模型API Key")
        self.keyword_ai_api_key.setFixedHeight(24)
        self.keyword_ai_api_key.setMaximumWidth(260)

        # 使用一行容纳抠图开关与方法
        rb_row = QWidget()
        rb_h = QHBoxLayout(rb_row)
        rb_h.setContentsMargins(0, 0, 0, 0)
        rb_h.setSpacing(8)
        rb_h.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        rb_h.addWidget(self.remove_bg_enabled)
        rb_h.addWidget(QLabel("方式"))
        rb_h.addWidget(self.remove_bg_method)
        rb_h.addWidget(QLabel("API Key"))
        rb_h.addWidget(self.remove_bg_api_key)
        rb_h.addStretch(1)
        form.addRow(_col_label("抠  图"), rb_row)

        # 自动裁剪配置
        crop_row = QWidget()
        crop_h = QHBoxLayout(crop_row)
        crop_h.setContentsMargins(0, 0, 0, 0)
        crop_h.setSpacing(8)
        crop_h.addWidget(self.auto_crop_enabled)
        crop_h.addWidget(QLabel("尺寸："))
        crop_h.addWidget(self.crop_size)
        crop_h.addWidget(QLabel("模式："))
        crop_h.addWidget(self.crop_mode)
        crop_h.addWidget(QLabel("填充色："))
        crop_h.addWidget(self.contain_bg_color)
        crop_h.addStretch(1)
        # 第三行：裁剪设置
        form.addRow(_col_label("裁  剪："), crop_row)

        thumb_row = QWidget()
        thumb_h = QHBoxLayout(thumb_row)
        thumb_h.setContentsMargins(0, 0, 0, 0)
        thumb_h.setSpacing(8)
        thumb_h.addWidget(QLabel("缩略图尺寸："))
        thumb_h.addWidget(self.thumbnail_size)
        thumb_h.addWidget(QLabel("输出格式："))
        thumb_h.addWidget(self.image_format)
        thumb_h.addWidget(QLabel("每论点图片数："))
        thumb_h.addWidget(self.image_per_point)
        thumb_h.addStretch(1)
        # 第四行：预览与输出
        form.addRow(_col_label("预览与输出："), thumb_row)

        # 关键词分析来源行
        kw_row = QWidget()
        kw_h = QHBoxLayout(kw_row)
        kw_h.setContentsMargins(0, 0, 0, 0)
        kw_h.setSpacing(8)
        kw_h.addWidget(QLabel("来源："))
        kw_h.addWidget(self.keyword_source)
        kw_h.addWidget(QLabel("AI模型："))
        kw_h.addWidget(self.keyword_ai_model)
        kw_h.addWidget(QLabel("API Key："))
        kw_h.addWidget(self.keyword_ai_api_key)
        kw_h.addStretch(1)
        # 第五行：关键词分析
        form.addRow(_col_label("关键词分析："), kw_row)

        top_layout.addLayout(form)
        # 顶部区域固定高度，避免与表格抢占伸展空间
        try:
            top.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        except Exception:
            pass

        layout.addWidget(top)

        # 第六行：分隔线
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Sunken)
        layout.addWidget(divider)

        # 第七行：提示信息 与 保存配置按钮（分割线之下）
        status_row = QHBoxLayout()
        status_row.setContentsMargins(0, 0, 0, 0)
        status_row.setSpacing(8)
        self.info_label = QLabel(f"分论点数：{len(self.points)} | 提示：配置在上方，表格在下方")
        self.info_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        status_row.addWidget(self.info_label)
        status_row.addStretch(1)
        btn_save_cfg = QPushButton("保存配置")
        btn_save_cfg.setObjectName("InlineButton")
        # 统一顶部两个按钮的高度与字体大小
        btn_save_cfg.setFixedHeight(28)
        try:
            btn_save_cfg.setMinimumWidth(90)
            btn_save_cfg.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            btn_save_cfg.setStyleSheet("padding:0 14px; font-size:13px;")
        except Exception:
            pass
        btn_save_cfg.clicked.connect(self._save_config)
        status_row.addWidget(btn_save_cfg)
        # 右上角：全屏/退出全屏切换
        self.btn_fullscreen = QPushButton("全屏")
        try:
            self.btn_fullscreen.setObjectName("InlineButton")
            self.btn_fullscreen.setFixedHeight(28)
            self.btn_fullscreen.setMinimumWidth(72)
            self.btn_fullscreen.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.btn_fullscreen.setStyleSheet("font-size:13px;")
        except Exception:
            pass
        try:
            self.btn_fullscreen.clicked.connect(self._toggle_fullscreen)
        except Exception:
            pass
        status_row.addWidget(self.btn_fullscreen)
        layout.addLayout(status_row)

        # 主表格
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["论点", "关键词", "预览", "操作", "状态"])
        # 显示垂直表头，允许调整每行高度
        self.table.verticalHeader().setVisible(True)
        try:
            self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
            # 默认行高先设为一个合理值，稍后根据缩略图尺寸再调整
            self.table.verticalHeader().setDefaultSectionSize(36)
            # 行高变化时，更新预览控件尺寸并保存布局
            try:
                self.table.verticalHeader().sectionResized.connect(self._on_row_resized)
            except Exception:
                pass
            self.table.verticalHeader().sectionResized.connect(lambda *_: self._snapshot_row_heights())
        except Exception:
            pass
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # 边框与网格线颜色：浅黑色
        try:
            self.table.setStyleSheet(
                "QTableWidget{border:1px solid #444; gridline-color:#444;} "
                "QHeaderView::section{background:#f5f6f8; padding:4px 8px; font-weight:600; color:#333; border-bottom:1px solid #d9d9de;} "
                "QTableWidget::item:selected{background: transparent; color: inherit;} "
                "QTableWidget::item:selected:!active{background: transparent; color: inherit;}"
            )
            self.table.setShowGrid(True)
        except Exception:
            pass
        hh = self.table.horizontalHeader()
        try:
            # 论点/关键词/预览支持手动拖动调整，操作/状态固定宽度
            hh.setSectionResizeMode(0, QHeaderView.Interactive)
            hh.setSectionResizeMode(1, QHeaderView.Stretch)
            # 预览列默认拉伸，填充剩余宽度，避免窗口变化时表格不随之扩展
            hh.setSectionResizeMode(2, QHeaderView.Stretch)
            hh.setSectionResizeMode(3, QHeaderView.Fixed)
            hh.setSectionResizeMode(4, QHeaderView.Fixed)
            # 最后一列不自动拉伸，保持固定宽度
            hh.setStretchLastSection(False)
            try:
                hh.setDefaultAlignment(Qt.AlignCenter)
                hh.sectionResized.connect(lambda *_: self._snapshot_widths())
            except Exception:
                pass
        except Exception:
            pass
        # 预览/操作/状态初始宽度（紧凑）
        self._thumb_w, self._thumb_h = self._parse_size(self.settings.get("thumbnail_size", "160x120"), default=(140, 100))
        # 根据缩略图高度调整默认行高，避免预览错位与边框缺失
        try:
            vh = self.table.verticalHeader()
            min_h = max(48, min(self._thumb_h + 6, 140))
            vh.setDefaultSectionSize(min_h)
        except Exception:
            pass
        # 应用记忆的列宽；若无，则设定默认宽并进行一次内容自适应
        self._apply_saved_widths(hh)
        # 隐藏水平与垂直滚动条
        try:
            self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        except Exception:
            pass
        # 移除首行强制低高度设置，统一由默认与内容自适应控制
        # 让表格在垂直方向优先获得多余空间
        layout.addWidget(self.table, 1)

        # 填充行
        self._fill_rows()
        # 应用记忆的行高
        try:
            self._apply_saved_row_heights()
        except Exception:
            pass

        # 底部操作
        bottom = QWidget()
        b = QHBoxLayout(bottom)
        b.setAlignment(Qt.AlignCenter)
        b.setSpacing(10)
        # 底部操作区固定高度，避免与表格争夺伸展空间
        try:
            bottom.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        except Exception:
            pass

        self.btn_batch_pick = QPushButton("批量自动配图")
        self.btn_batch_pick.setFixedHeight(28)
        self.btn_remove_bg = QPushButton("全部抠图")
        self.btn_remove_bg.setFixedHeight(28)
        self.btn_crop = QPushButton("全部裁剪")
        self.btn_crop.setFixedHeight(28)
        self.btn_write_zip = QPushButton("写入模板并压缩（配图版）")
        self.btn_write_zip.setFixedHeight(28)
        self.btn_back = QPushButton("返回")
        self.btn_back.setFixedHeight(28)

        # 仅占位：功能稍后实现
        self.btn_batch_pick.clicked.connect(self._on_batch_pick)
        self.btn_remove_bg.clicked.connect(self._on_remove_bg_all)
        self.btn_crop.clicked.connect(self._on_crop_all)
        self.btn_write_zip.clicked.connect(self._write_zip_with_images)
        self.btn_back.clicked.connect(self.accept)

        b.addWidget(self.btn_batch_pick)
        b.addWidget(self.btn_remove_bg)
        b.addWidget(self.btn_crop)
        b.addWidget(self.btn_write_zip)
        b.addWidget(self.btn_back)
        layout.addWidget(bottom)

        self.setLayout(layout)
        # 默认尺寸与居中显示
        self.resize(840, 560)
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)

    def _vlabel(self, text: str) -> QLabel:
        lab = QLabel(text)
        lab.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return lab

    def _info(self, text: str):
        try:
            self.info_label.setText(text)
        except Exception:
            pass

    def _save_config(self):
        src = self.image_source.currentText().strip() or "pixabay"
        cur_key = self.image_api_key.text().strip()
        updated = {
            "enable_image_feature": True,
            "image_source": src,
            # 仍保留通用字段，便于向后兼容
            "image_api_key": cur_key,
            # 按来源分别保存 API Key
            f"{src}_api_key": cur_key,
            "remove_bg_enabled": bool(self.remove_bg_enabled.isChecked()),
            "remove_bg_method": self.remove_bg_method.currentText().strip() or "rembg",
            "remove_bg_api_key": self.remove_bg_api_key.text().strip(),
            "auto_crop_enabled": bool(self.auto_crop_enabled.isChecked()),
            "crop_size": self.crop_size.text().strip() or "1024x1024",
            "crop_mode": self.crop_mode.currentText().strip() or "cover",
            "contain_bg_color": self.contain_bg_color.text().strip() or "#FFFFFF",
            "thumbnail_size": self.thumbnail_size.text().strip() or "160x120",
            "image_format": self.image_format.currentText().strip() or "png",
            "image_per_point": int(self.image_per_point.value()),
            # 关键词分析来源
            "keyword_source": self.keyword_source.currentText().strip() or "本地分析",
            "keyword_ai_model": self.keyword_ai_model.text().strip(),
            "keyword_ai_api_key": self.keyword_ai_api_key.text().strip(),
        }
        # 保持其他来源的 API Key（若此前已保存）不丢失
        for s in ("pixabay", "pexels", "ai"):
            k = f"{s}_api_key"
            if k not in updated:
                try:
                    updated[k] = self.settings.get(k, "")
                except Exception:
                    updated[k] = ""
        try:
            self.settings.update(updated)
            # 统一保存入口
            try:
                ConfigManager.instance().save()
            except Exception:
                save_settings(self.settings_path, self.settings)
            self._info("配置已保存")
        except Exception:
            self._info("配置保存失败")

    def _toggle_fullscreen(self):
        try:
            if self.isFullScreen():
                self.showNormal()
                try:
                    self.btn_fullscreen.setText("全屏")
                except Exception:
                    pass
            else:
                self.showFullScreen()
                try:
                    self.btn_fullscreen.setText("退出全屏")
                except Exception:
                    pass
        except Exception:
            pass

    def keyPressEvent(self, event):
        try:
            if event.key() == Qt.Key_F11:
                # F11 切换全屏/退出全屏
                self._toggle_fullscreen()
                event.accept()
                return
            if event.key() == Qt.Key_Escape and self.isFullScreen():
                # Esc 在全屏时退出全屏
                self.showNormal()
                try:
                    self.btn_fullscreen.setText("全屏")
                except Exception:
                    pass
                event.accept()
                return
        except Exception:
            pass
        try:
            return super().keyPressEvent(event)
        except Exception:
            pass

    def _fill_rows(self):
        self.table.setRowCount(len(self.points))
        for i, (pt_title, pt_content) in enumerate(self.points):
            # 论点标题
            title_item = QTableWidgetItem(pt_title)
            title_item.setFlags(title_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(i, 0, title_item)

            # 关键词：按来源决定是否使用本地生成
            try:
                use_local = (self.keyword_source.currentText().strip() == "本地分析")
            except Exception:
                use_local = True
            kw_init = (self._generate_keywords_local(pt_title, pt_content) if use_local else pt_title.strip())
            kw_item = QTableWidgetItem(kw_init)
            kw_item.setToolTip("可编辑，按 Enter 保存；用于搜索图片")
            self.table.setItem(i, 1, kw_item)

            # 预览：占位缩略图（紧凑）
            prev = QLabel("(暂无)")
            prev.setAlignment(Qt.AlignCenter)
            # 保证预览控件在单元格内完整显示：不小于 48，高度与默认行高匹配
            try:
                row_h = self.table.verticalHeader().defaultSectionSize()
            except Exception:
                row_h = 48
            prev.setFixedSize(max(64, self._thumb_w), max(48, min(self._thumb_h, row_h - 4)))
            prev.setStyleSheet("border:1px solid #e5e7eb; color:#6b7280;")
            prev.setToolTip("(暂无预览) 点击或悬浮将显示预览")
            # 点击预览（将来有图片时弹出大图预览）
            try:
                prev.mousePressEvent = (lambda event, row=i, label=prev: self._on_preview_click(event, row, label))
            except Exception:
                pass
            # 悬浮预览：进入时显示、离开时关闭
            try:
                prev.enterEvent = (lambda event, row=i, label=prev: self._on_preview_hover(event, row, label, True))
                prev.leaveEvent = (lambda event, row=i, label=prev: self._on_preview_hover(event, row, label, False))
            except Exception:
                pass
            self.table.setCellWidget(i, 2, prev)

            # 操作：内联按钮容器
            cell = QWidget()
            h = QHBoxLayout(cell)
            h.setContentsMargins(6, 2, 6, 2)
            h.setSpacing(3)
            h.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            btn_refresh = QPushButton("重新配图")
            btn_refresh.setObjectName("InlineButton")
            btn_refresh.setFixedHeight(18)
            btn_refresh.setMaximumWidth(74)
            btn_refresh.setStyleSheet("QPushButton#InlineButton{padding:2px 6px; font-size:12px; min-width:0;}")
            btn_replace = QPushButton("替换本地图")
            btn_replace.setObjectName("InlineButton")
            btn_replace.setFixedHeight(18)
            btn_replace.setMaximumWidth(88)
            btn_replace.setStyleSheet("QPushButton#InlineButton{padding:2px 6px; font-size:12px; min-width:0;}")
            # 重新配图：根据关键词重新抓取并更新预览
            btn_refresh.clicked.connect(lambda _, row=i: self._on_refresh_image(row))
            # 修正事件绑定，确保点击对应行生效
            btn_replace.clicked.connect(lambda _, r=i: self._on_replace_local_image(r))
            h.addWidget(btn_refresh)
            h.addWidget(btn_replace)
            self.table.setCellWidget(i, 3, cell)

            # 状态
            st_item = QTableWidgetItem("未配图")
            st_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            self.table.setItem(i, 4, st_item)

            # 移除首行强调色，统一由表头样式区分

        # 一次性内容自适应（中间列）：关键词、预览、操作
        try:
            # 仅让关键词与预览列按内容自适应；操作列保持既定紧凑宽度
            for c in [1, 2]:
                self.table.resizeColumnToContents(c)
            self.table.resizeRowsToContents()
        except Exception:
            pass

    def _on_preview_click(self, event, row: int, label: QLabel):
        """点击缩略图进行预览：若存在图片，弹窗显示较大预览；否则给出提示。"""
        try:
            pm = getattr(label, 'pixmap', None)
            pm = pm() if callable(pm) else pm
            if pm:
                dlg = QDialog(self)
                dlg.setWindowTitle(f"预览：第{row+1}行")
                v = QVBoxLayout(dlg)
                v.setContentsMargins(10,10,10,10)
                v.setSpacing(8)
                img_label = QLabel()
                img_label.setAlignment(Qt.AlignCenter)
                img_label.setPixmap(pm)
                v.addWidget(img_label)
                dlg.resize(max(360, self._thumb_w*2), max(280, self._thumb_h*2))
                dlg.setModal(True)
                dlg.exec_()
            else:
                self._info(f"第{row+1}行暂无图片可预览")
        except Exception:
            self._info("预览失败")

    def _on_replace_local_image(self, row: int):
        """选择本地图片，显示缩略图，并更新状态。"""
        try:
            from PyQt5.QtWidgets import QFileDialog
            fmt = self.image_format.currentText() if hasattr(self, 'image_format') else 'png'
            filters = f"Images (*.{fmt} *.jpg *.jpeg *.png *.webp);;All Files (*.*)"
            fpath, _ = QFileDialog.getOpenFileName(self, f"选择第{row+1}行图片", "", filters)
            if not fpath:
                return
            # 保存路径
            try:
                fpath_abs = os.path.abspath(fpath)
            except Exception:
                fpath_abs = fpath
            self.local_images[row] = fpath_abs
            # 渲染缩略图到预览列
            label = self.table.cellWidget(row, 2)
            if isinstance(label, QLabel):
                pm = self._safe_load_pixmap(self.local_images[row])
                if not pm.isNull():
                    try:
                        row_h = self.table.rowHeight(row)
                        target_w = max(64, self._thumb_w)
                        target_h = max(48, min(self._thumb_h, row_h - 4))
                    except Exception:
                        target_w, target_h = self._thumb_w, self._thumb_h
                    pm_scaled = pm.scaled(target_w, target_h, Qt.KeepAspectRatio, transformMode=Qt.SmoothTransformation)
                    label.setPixmap(pm_scaled)
                    label.setToolTip("点击或悬浮查看预览")
                    label.setStyleSheet("border:1px solid #e5e7eb; color:#111827;")
                else:
                    # 加载失败时给出提示，便于定位问题
                    try:
                        self._info(f"第{row+1}行图片加载失败：{self.local_images[row]}")
                    except Exception:
                        pass
                    label.setText("(加载失败)")
            # 更新状态列
            self.table.setItem(row, 4, QTableWidgetItem("已配图"))
            self._info(f"已替换第{row+1}行本地图")
        except Exception as e:
            self._info(f"替换本地图失败：{e}")

    def _write_zip_with_images(self):
        """写入有图模板并压缩：Excel与图片在压缩包同级。"""
        try:
            # 组装写入参数
            from app.core.excel_writer import write_to_template
            from app.core.zipper import make_zip
            from app.gui.message_dialog import MessageDialog
            import os, shutil, tempfile

            # 标题从父窗口的 last_title 或者用首行标题
            title = getattr(self.parent(), 'last_title', None) or (self.points[0][0] if self.points else '输出')
            # 仅使用“有图模板”路径；未配置则阻止写入
            template_path = self.settings.get('template_excel_image_path')
            if not template_path:
                MessageDialog.warning(self, '提示', '未配置“有图模板”路径，请在设置中填写后重试')
                return
            output_dir = self.settings.get('zip_output_dir', 'output')

            # 检测图片文件名（不含后缀）是否重复，若重复则阻止写入
            non_empty_imgs = [p for p in self.local_images if p]
            base_names_wo_ext = [os.path.splitext(os.path.basename(p))[0] for p in non_empty_imgs]
            if len(set(base_names_wo_ext)) != len(base_names_wo_ext):
                # 找出重复项
                seen = set()
                dups = []
                for name in base_names_wo_ext:
                    if name in seen and name not in dups:
                        dups.append(name)
                    seen.add(name)
                MessageDialog.error(self, '错误', f'图片文件名（不含后缀）重复：{", ".join(dups)}\n请修改后再试')
                return

            # 写入 Excel（图片列：图片_1）
            out_xlsx = write_to_template(
                template_path=template_path,
                output_dir=output_dir,
                title=title,
                points=self.points,
                column_map={
                    'page_column': self.settings.get('page_column', '页面'),
                    'point_title_column': self.settings.get('point_title_column', '文本_1'),
                    'point_content_column': self.settings.get('point_content_column', '文本_2'),
                    'extra_text_column': self.settings.get('extra_text_column', '文本_3'),
                    'extra_text_default': self.settings.get('extra_text_default', '内容仅供参考，身体不适请及时就医!'),
                    'image_column': self.settings.get('image_column', '图片_1')
                },
                images=self.local_images
            )

            # 将本地图片复制到临时目录以便打包（与Excel同级写入到zip根目录）
            tmp_images_dir = None
            if any(self.local_images):
                tmp_images_dir = tempfile.mkdtemp(prefix='images_')
                for p in filter(None, self.local_images):
                    try:
                        shutil.copy(p, os.path.join(tmp_images_dir, os.path.basename(p)))
                    except Exception:
                        pass

            # 生成配图版 ZIP
            zip_path = make_zip(out_xlsx, output_dir, images_dir=tmp_images_dir, zip_name_suffix='配图版')

            # 根据设置删除Excel，仅保留zip
            if self.settings.get('delete_excel_after_zip', True):
                try:
                    os.remove(out_xlsx)
                except Exception:
                    pass

            # 清理临时目录
            if tmp_images_dir and os.path.isdir(tmp_images_dir):
                try:
                    shutil.rmtree(tmp_images_dir)
                except Exception:
                    pass

            MessageDialog.info(self, '完成', f'已生成配图版压缩包：\n{zip_path}')
            self._info(f"完成：压缩包已生成 → {os.path.basename(zip_path)}")
        except Exception as e:
            try:
                from app.gui.message_dialog import MessageDialog
                MessageDialog.error(self, '错误', f'写入压缩失败：{e}')
            except Exception:
                pass
            self._info(f"写入压缩失败：{e}")

    # ========== 配图与抠图功能 ==========
    def _on_refresh_image(self, row: int):
        """单行重新配图：改为后台线程，避免阻塞 UI。"""
        try:
            kw_item = self.table.item(row, 1)
            keyword = kw_item.text().strip() if kw_item else ""
            if not keyword:
                self._info(f"第{row+1}行关键词为空，无法配图")
                return
            # 提示处理中
            self._info(f"正在为第{row+1}行配图…")
            import concurrent.futures
            # 创建单次执行器，避免与批量任务互相影响
            execu = concurrent.futures.ThreadPoolExecutor(max_workers=1)
            self._single_executor = execu

            fut = execu.submit(self._download_image_for_keyword, row, keyword)

            def _post_ui(fn):
                try:
                    QTimer.singleShot(0, fn)
                except Exception:
                    try:
                        fn()
                    except Exception:
                        pass

            def _on_done(f):
                try:
                    fpath = f.result()
                except Exception:
                    fpath = None

                def _update():
                    try:
                        if fpath:
                            self._update_preview_label(row, fpath)
                            self.table.setItem(row, 4, QTableWidgetItem("已配图"))
                            self._info(f"已为第{row+1}行配图")
                        else:
                            self.table.setItem(row, 4, QTableWidgetItem("配图失败"))
                            self._info(f"第{row+1}行配图失败")
                    finally:
                        try:
                            execu.shutdown(wait=False)
                            if getattr(self, '_single_executor', None) is execu:
                                self._single_executor = None
                        except Exception:
                            pass

                _post_ui(_update)

            try:
                fut.add_done_callback(_on_done)
            except Exception:
                pass
        except Exception as e:
            self._info(f"重新配图失败：{e}")

    def _on_batch_pick(self):
        """批量根据关键词自动配图（线程池并发），避免阻塞 UI。"""
        try:
            rows = self.table.rowCount()
        except Exception:
            rows = len(self.points)
        # 收集需要处理的行
        tasks = []
        for r in range(rows):
            kw_item = self.table.item(r, 1)
            keyword = kw_item.text().strip() if kw_item else ""
            if keyword:
                tasks.append((r, keyword))
        total = len(tasks)
        if total == 0:
            self._info("无可用关键词，无法批量配图")
            return
        # 禁用按钮，提示处理中
        try:
            self.btn_batch_pick.setEnabled(False)
        except Exception:
            pass
        self._info(f"开始批量配图：共 {total} 行，已提交任务…")
        # 线程池并发，限制同时下载数以减轻卡顿
        import concurrent.futures
        self._batch_success = 0
        self._batch_done = 0
        self._batch_total = total
        self._batch_futures = []
        self._batch_executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)

        def _post_ui(fn):
            try:
                QTimer.singleShot(0, fn)
            except Exception:
                try:
                    fn()
                except Exception:
                    pass

        for r, kw in tasks:
            fut = self._batch_executor.submit(self._download_image_for_keyword, r, kw)
            self._batch_futures.append(fut)

            def _on_done(f, row=r, keyword=kw):
                try:
                    fpath = f.result()
                except Exception:
                    fpath = None
                def _update():
                    if fpath:
                        try:
                            self._update_preview_label(row, fpath)
                            self.table.setItem(row, 4, QTableWidgetItem("已配图"))
                            self._batch_success += 1
                        except Exception:
                            pass
                    else:
                        try:
                            self.table.setItem(row, 4, QTableWidgetItem("配图失败"))
                        except Exception:
                            pass
                    self._batch_done += 1
                    # 更新提示与完成状态（显示 成功/总，已完成/总）
                    try:
                        self._info(f"批量配图进度：成功 {self._batch_success}/{self._batch_total}，完成 {self._batch_done}/{self._batch_total}")
                    except Exception:
                        pass
                    if (self._batch_done >= self._batch_total):
                        try:
                            self._info(f"批量配图完成：成功 {self._batch_success}/{self._batch_total}")
                            self.btn_batch_pick.setEnabled(True)
                            try:
                                execu = getattr(self, '_batch_executor', None)
                                if execu:
                                    execu.shutdown(wait=False)
                                    self._batch_executor = None
                            except Exception:
                                pass
                        except Exception:
                            pass
                _post_ui(_update)

            try:
                fut.add_done_callback(_on_done)
            except Exception:
                pass

        # 不阻塞主线程，让下载在后台进行
        # 线程池会在应用退出时自动清理；如需提前取消，可在其它位置 shutdown。

    def _on_remove_bg_all(self):
        """为已配图的本地图片执行抠图（线程池并发），避免阻塞 UI。"""
        try:
            remove_method = (self.remove_bg_method.currentText().strip() if hasattr(self, 'remove_bg_method') else 'rembg')
            api_key = (self.remove_bg_api_key.text().strip() if hasattr(self, 'remove_bg_api_key') else '')
        except Exception:
            remove_method, api_key = 'rembg', ''

        from os import path
        tasks = []
        for idx, img_path in enumerate(self.local_images):
            if img_path and path.isfile(img_path):
                tasks.append((idx, img_path))
        total = len(tasks)
        if total == 0:
            self._info("无可用本地图片，无法执行全部抠图")
            return

        try:
            self.btn_remove_bg.setEnabled(False)
        except Exception:
            pass
        self._info(f"开始全部抠图：共 {total} 行，已提交任务…")

        import concurrent.futures
        self._rm_bg_success = 0
        self._rm_bg_done = 0
        self._rm_bg_total = total
        self._rm_bg_futures = []
        self._rm_bg_executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)

        def _post_ui(fn):
            try:
                QTimer.singleShot(0, fn)
            except Exception:
                try:
                    fn()
                except Exception:
                    pass

        def _do_remove(img_path: str):
            try:
                if remove_method == 'rembg':
                    try:
                        from rembg import remove  # type: ignore
                        from PIL import Image
                        im = Image.open(img_path).convert('RGBA')
                        im_out = remove(im)
                        out_path = self._derived_path(img_path, suffix='_nobg')
                        im_out.save(out_path)
                        return out_path
                    except Exception:
                        # 回退到简单白底抠图
                        return self._simple_remove_bg(img_path)
                else:
                    # API 抠图（未实现）：直接使用简单白底抠图
                    return self._simple_remove_bg(img_path)
            except Exception:
                return None

        for idx, img_path in tasks:
            fut = self._rm_bg_executor.submit(_do_remove, img_path)
            self._rm_bg_futures.append(fut)

            def _on_done(f, row=idx):
                try:
                    out_path = f.result()
                except Exception:
                    out_path = None

                def _update():
                    if out_path:
                        try:
                            self.local_images[row] = out_path
                            self._update_preview_label(row, out_path)
                            self.table.setItem(row, 4, QTableWidgetItem("已抠图"))
                            self._rm_bg_success += 1
                        except Exception:
                            pass
                    else:
                        try:
                            self.table.setItem(row, 4, QTableWidgetItem("抠图失败"))
                        except Exception:
                            pass
                    self._rm_bg_done += 1
                    try:
                        self._info(f"全部抠图进度：成功 {self._rm_bg_success}/{self._rm_bg_total}，完成 {self._rm_bg_done}/{self._rm_bg_total}")
                    except Exception:
                        pass
                    if self._rm_bg_done >= self._rm_bg_total:
                        try:
                            self._info(f"全部抠图完成：成功 {self._rm_bg_success}/{self._rm_bg_total}")
                            self.btn_remove_bg.setEnabled(True)
                            execu = getattr(self, '_rm_bg_executor', None)
                            if execu:
                                execu.shutdown(wait=False)
                                self._rm_bg_executor = None
                        except Exception:
                            pass

                _post_ui(_update)

            try:
                fut.add_done_callback(_on_done)
            except Exception:
                pass

        # 后台执行，不阻塞主线程

    def _simple_remove_bg(self, img_path: str) -> Optional[str]:
        """将近白色背景置为透明的简易抠图。"""
        try:
            from PIL import Image
            img = Image.open(img_path).convert('RGBA')
            pixels = img.getdata()
            new_pixels = []
            for p in pixels:
                r, g, b, a = p
                if r > 240 and g > 240 and b > 240:
                    new_pixels.append((r, g, b, 0))
                else:
                    new_pixels.append(p)
            img.putdata(new_pixels)
            out_path = self._derived_path(img_path, suffix='_nobg')
            img.save(out_path)
            return out_path
        except Exception:
            return None

    def _derived_path(self, src_path: str, suffix: str) -> str:
        try:
            import os
            base, ext = os.path.splitext(src_path)
            fmt = (self.image_format.currentText().strip() if hasattr(self, 'image_format') else 'png')
            # 若原始扩展与目标不一致，改用目标格式扩展
            dst_ext = f".{fmt.lower()}"
            return f"{base}{suffix}{dst_ext}"
        except Exception:
            return src_path

    def _download_image_for_keyword(self, row: int, keyword: str) -> Optional[str]:
        """下载一张图片并返回本地路径。
        支持 pixabay/pexels，缺少 API Key 时回退到 picsum。
        """
        try:
            import os
            import io
            import json
            import requests
            from PIL import Image

            # 目标尺寸：使用缩略图尺寸的 3 倍，保证清晰
            try:
                tw, th = self._parse_size(self.settings.get("thumbnail_size", "160x120"), default=(160, 120))
                dw, dh = max(320, tw * 3), max(240, th * 3)
            except Exception:
                dw, dh = 640, 480

            source = (self.image_source.currentText().strip() if hasattr(self, 'image_source') else 'pixabay')
            # 按来源读取 API Key（不再支持 Unsplash 特殊字段）
            try:
                if hasattr(self, 'image_api_key'):
                    api_key = self.image_api_key.text().strip()
                else:
                    api_key = self.settings.get(f"{source}_api_key", self.settings.get("image_api_key", ""))
            except Exception:
                api_key = ''

            def _fallback_url() -> str:
                seed = f"{keyword.replace(' ', '_')}_{row}"
                return f"https://picsum.photos/seed/{seed}/{dw}/{dh}"

            img_url = None
            q = keyword.strip()
            if source == "pixabay" and api_key:
                # Pixabay: https://pixabay.com/api/?key=KEY&q=QUERY&image_type=photo&orientation=horizontal&per_page=10&safesearch=true
                try:
                    resp = requests.get(
                        "https://pixabay.com/api/",
                        params={
                            "key": api_key,
                            "q": q,
                            "image_type": "photo",
                            "orientation": "horizontal",
                            "safesearch": "true",
                            "per_page": 5,
                        },
                        timeout=12,
                    )
                    if resp.ok:
                        data = resp.json()
                        hits = data.get("hits") or []
                        if hits:
                            # 优先选择较大图
                            best = hits[0]
                            img_url = best.get("largeImageURL") or best.get("webformatURL")
                except Exception:
                    img_url = None
            elif source == "pexels" and api_key:
                # Pexels: GET https://api.pexels.com/v1/search?query=QUERY&per_page=1  (Header: Authorization)
                try:
                    resp = requests.get(
                        "https://api.pexels.com/v1/search",
                        params={"query": q, "per_page": 5},
                        headers={"Authorization": api_key},
                        timeout=12,
                    )
                    if resp.ok:
                        data = resp.json()
                        photos = data.get("photos") or []
                        if photos:
                            first = photos[0]
                            src = first.get("src") or {}
                            img_url = src.get("large") or src.get("large2x") or src.get("original")
                except Exception:
                    img_url = None

            # 回退到占位图
            if not img_url:
                img_url = _fallback_url()

            # 下载图片并统一保存为配置的格式（失败则回退占位图）
            try:
                r = requests.get(img_url, headers={"Accept": "image/*"}, timeout=(5, 15))
                r.raise_for_status()
                data = r.content
            except Exception:
                try:
                    r = requests.get(_fallback_url(), headers={"Accept": "image/*"}, timeout=(5, 12))
                    r.raise_for_status()
                    data = r.content
                except Exception:
                    return None
            img = Image.open(io.BytesIO(data))
            # 可选：统一尺寸
            try:
                img = img.convert("RGBA") if img.mode in ("P", "LA") else img.convert("RGB")
                img = img.resize((dw, dh))
            except Exception:
                pass

            out_dir = self.settings.get('zip_output_dir', 'output')
            # 使用绝对路径，避免因工作目录变化导致 QPixmap 无法读取
            try:
                out_dir = os.path.abspath(out_dir)
            except Exception:
                pass
            cache_dir = os.path.join(out_dir, 'images_cache')
            os.makedirs(cache_dir, exist_ok=True)
            fmt = (self.image_format.currentText().strip() if hasattr(self, 'image_format') else 'png')
            safe_kw = ''.join(c for c in keyword if c.isalnum() or c in ('_', '-'))[:40]
            fname = f"row{row+1}_{safe_kw}.{fmt.lower()}"
            fpath = os.path.join(cache_dir, fname)
            try:
                fpath = os.path.abspath(fpath)
            except Exception:
                pass
            try:
                img.save(fpath)
            except Exception:
                # 如果保存失败，回退为 PNG
                base, _ = os.path.splitext(fpath)
                fpath = f"{base}.png"
                img.save(fpath)
            self.local_images[row] = fpath
            return fpath
        except Exception:
            return None

    # ========== 关键词生成与来源切换 ==========
    def _extract_tokens(self, text: str) -> List[str]:
        try:
            import re
            # 使用常见汉字/英文分隔符切分
            parts = re.split(r"[\s，。,!！？?、；;：:·\-—\(\)\[\]\{\}\"']+", text)
            toks = []
            for p in parts:
                t = p.strip()
                if len(t) >= 2 and not t.isdigit():
                    toks.append(t)
            # 去重保持顺序
            seen = set()
            uniq = []
            for t in toks:
                if t not in seen:
                    seen.add(t)
                    uniq.append(t)
            return uniq
        except Exception:
            return [text.strip()] if text else []

    def _generate_keywords_local(self, title: str, content: str) -> str:
        try:
            # 简单策略：优先取标题中的词，补充少量正文中的词
            from app.core.punctuation import normalize_punctuation
            t = normalize_punctuation(title or "")
            c = normalize_punctuation((content or "")[:120])
            toks_t = self._extract_tokens(t)
            toks_c = [w for w in self._extract_tokens(c) if w not in toks_t]
            # 取前几项，避免过长
            sel = (toks_t[:3] + toks_c[:2]) or [t.strip() for t in [title] if t.strip()]
            return " ".join(sel)
        except Exception:
            return (title or "").strip()

    def _on_keyword_source_changed(self, src: str):
        try:
            use_local = (src.strip() == "本地分析")
        except Exception:
            use_local = True
        try:
            rows = self.table.rowCount()
        except Exception:
            rows = 0
        for r in range(rows):
            try:
                title_item = self.table.item(r, 0)
                kw_item = self.table.item(r, 1)
                title = title_item.text().strip() if title_item else ""
                kw_cur = kw_item.text().strip() if kw_item else ""
                # 仅在关键词为空或与标题完全一致时，进行填充，避免覆盖用户已编辑内容
                if use_local and (not kw_cur or kw_cur == title):
                    # 获取内容以辅助生成
                    pt_content = self.points[r][1] if r < len(self.points) else ""
                    new_kw = self._generate_keywords_local(title, pt_content)
                    self.table.setItem(r, 1, QTableWidgetItem(new_kw))
            except Exception:
                pass

    def _on_image_source_changed(self, new_source: str):
        # 保存上一来源的 key，并显示新来源下的已保存 key（移除 Unsplash 特殊处理）
        try:
            prev = getattr(self, '_last_image_source', '') or ''
            prev_key = self.image_api_key.text().strip() if hasattr(self, 'image_api_key') else ''
            if prev:
                self.settings[f"{prev}_api_key"] = prev_key
                try:
                    ConfigManager.instance().save()
                except Exception:
                    try:
                        if self.settings_path:
                            save_settings(self.settings_path, self.settings)
                    except Exception:
                        pass
            self._last_image_source = new_source.strip()
            try:
                self.image_api_key.setVisible(True)
                self.lbl_api_key.setVisible(True)
            except Exception:
                pass
            new_key = self.settings.get(f"{self._last_image_source}_api_key", "")
            if hasattr(self, 'image_api_key'):
                self.image_api_key.setText(new_key)
            self._info(f"已切换来源：{self._last_image_source}，API Key 已联动")
        except Exception:
            pass

    def _on_crop_all(self):
        """遍历所有已配图，使用线程池进行裁剪并更新预览，避免阻塞 UI。"""
        try:
            tasks = []
            for idx, img_path in enumerate(self.local_images):
                if img_path:
                    tasks.append((idx, img_path))
            total = len(tasks)
            if total == 0:
                self._info("无可用本地图片，无法执行全部裁剪")
                return

            try:
                self.btn_crop.setEnabled(False)
            except Exception:
                pass
            self._info(f"开始全部裁剪：共 {total} 行，已提交任务…")

            import concurrent.futures
            self._crop_success = 0
            self._crop_done = 0
            self._crop_total = total
            self._crop_futures = []
            self._crop_executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)

            def _post_ui(fn):
                try:
                    QTimer.singleShot(0, fn)
                except Exception:
                    try:
                        fn()
                    except Exception:
                        pass

            def _do_crop(img_path: str):
                try:
                    from PIL import Image
                    is_nobg = ("_nobg" in img_path)
                    im = Image.open(img_path).convert('RGBA')
                    alpha = im.split()[3]
                    has_transparency = alpha.getextrema()[0] < 255
                    if is_nobg or has_transparency:
                        return self._crop_to_content(im, img_path)
                    else:
                        try:
                            w, h = self._parse_size(self.settings.get("crop_size", "1024x1024"), default=(1024, 1024))
                        except Exception:
                            w, h = 1024, 1024
                        mode = self.settings.get("crop_mode", "cover")
                        bg = self.settings.get("contain_bg_color", "#FFFFFF")
                        return self._resize_or_pad_to_size(im, img_path, (w, h), mode, bg)
                except Exception:
                    return None

            for idx, img_path in tasks:
                fut = self._crop_executor.submit(_do_crop, img_path)
                self._crop_futures.append(fut)

                def _on_done(f, row=idx):
                    try:
                        out_path = f.result()
                    except Exception:
                        out_path = None

                    def _update():
                        if out_path:
                            try:
                                self.local_images[row] = out_path
                                self._update_preview_label(row, out_path)
                                self.table.setItem(row, 4, QTableWidgetItem("已裁剪"))
                                self._crop_success += 1
                            except Exception:
                                pass
                        else:
                            try:
                                self.table.setItem(row, 4, QTableWidgetItem("裁剪失败"))
                            except Exception:
                                pass
                        self._crop_done += 1
                        try:
                            self._info(f"全部裁剪进度：成功 {self._crop_success}/{self._crop_total}，完成 {self._crop_done}/{self._crop_total}")
                        except Exception:
                            pass
                        if self._crop_done >= self._crop_total:
                            try:
                                self._info(f"全部裁剪完成：成功 {self._crop_success}/{self._crop_total}")
                                self.btn_crop.setEnabled(True)
                                execu = getattr(self, '_crop_executor', None)
                                if execu:
                                    execu.shutdown(wait=False)
                                    self._crop_executor = None
                            except Exception:
                                pass

                    _post_ui(_update)

                try:
                    fut.add_done_callback(_on_done)
                except Exception:
                    pass

            # 后台执行，不阻塞主线程
        except Exception:
            self._info("裁剪过程中出现错误")

    def _crop_to_content(self, im, src_path: str) -> Optional[str]:
        """基于透明像素的内容边缘裁剪。"""
        try:
            from PIL import Image
            alpha = im.split()[3]
            bbox = alpha.getbbox()
            if not bbox:
                return None
            im_cropped = im.crop(bbox)
            out_path = self._derived_path(src_path, suffix='_crop')
            im_cropped.save(out_path)
            return out_path
        except Exception:
            return None

    def _resize_or_pad_to_size(self, im, src_path: str, dst_size: Tuple[int, int], mode: str, bg_color: str) -> Optional[str]:
        """按 cover/contain 逻辑生成目标尺寸图。"""
        try:
            from PIL import Image
            dw, dh = dst_size
            iw, ih = im.size
            if mode == 'contain':
                # 缩放以适配，留边填充背景色
                ratio = min(dw / iw, dh / ih)
                nw, nh = max(1, int(iw * ratio)), max(1, int(ih * ratio))
                im_resized = im.resize((nw, nh), Image.LANCZOS)
                bg = Image.new('RGBA', (dw, dh), self._hex_to_rgba(bg_color))
                x = (dw - nw) // 2
                y = (dh - nh) // 2
                bg.paste(im_resized, (x, y), im_resized)
                out = bg
            else:
                # cover：先等比放大填满，再居中裁剪
                ratio = max(dw / iw, dh / ih)
                nw, nh = max(1, int(iw * ratio)), max(1, int(ih * ratio))
                im_resized = im.resize((nw, nh), Image.LANCZOS)
                x = (nw - dw) // 2
                y = (nh - dh) // 2
                out = im_resized.crop((x, y, x + dw, y + dh))
            out_path = self._derived_path(src_path, suffix='_crop')
            out.save(out_path)
            return out_path
        except Exception:
            return None

    @staticmethod
    def _hex_to_rgba(s: str) -> Tuple[int, int, int, int]:
        try:
            s = s.strip().lstrip('#')
            if len(s) == 3:
                r, g, b = [int(ch * 2, 16) for ch in s]
            else:
                r = int(s[0:2], 16)
                g = int(s[2:4], 16)
                b = int(s[4:6], 16)
            return (r, g, b, 255)
        except Exception:
            return (255, 255, 255, 255)

    def _update_preview_label(self, row: int, fpath: str):
        try:
            label = self.table.cellWidget(row, 2)
            if not isinstance(label, QLabel):
                return
            pm = self._safe_load_pixmap(fpath)
            if pm.isNull():
                label.setText("(加载失败)")
                try:
                    self._info(f"第{row+1}行预览图片加载失败：{fpath}")
                except Exception:
                    pass
                return
            try:
                row_h = self.table.rowHeight(row)
                target_w = max(64, self._thumb_w)
                target_h = max(48, min(self._thumb_h, row_h - 4))
            except Exception:
                target_w, target_h = self._thumb_w, self._thumb_h
            pm_scaled = pm.scaled(target_w, target_h, Qt.KeepAspectRatio, transformMode=Qt.SmoothTransformation)
            # 同步更新控件尺寸，避免缩略图被裁切或留白
            try:
                label.setFixedSize(target_w, target_h)
            except Exception:
                pass
            label.setPixmap(pm_scaled)
            label.setToolTip("点击或悬浮查看预览")
            label.setStyleSheet("border:1px solid #e5e7eb; color:#111827;")
        except Exception:
            pass

    def _safe_load_pixmap(self, fpath: str) -> QPixmap:
        """更稳健地加载图片：
        - 使用绝对路径；
        - 尝试通过 QImageReader 读取，兼容多格式与带透明度图像；
        - 回退到 QPixmap 直接加载；
        - 失败时返回 isNull 的 QPixmap。
        """
        try:
            import os
            # 规范绝对路径
            try:
                fpath_abs = os.path.abspath(fpath)
            except Exception:
                fpath_abs = fpath
            # 先用 QImageReader，提高格式兼容性
            try:
                reader = QImageReader(fpath_abs)
                reader.setAutoTransform(True)
                image = reader.read()
                if not image.isNull():
                    pm = QPixmap.fromImage(image)
                    if not pm.isNull():
                        return pm
            except Exception:
                pass
            # 回退直接加载
            pm2 = QPixmap(fpath_abs)
            return pm2
        except Exception:
            return QPixmap()

    # ========== 布局记忆 ==========
    def _snapshot_widths(self):
        try:
            widths = [self.table.columnWidth(i) for i in range(5)]
            self.settings["image_dialog_table_widths"] = widths
        except Exception:
            pass

    def _apply_saved_widths(self, header: QHeaderView):
        try:
            widths = self.settings.get("image_dialog_table_widths")
        except Exception:
            widths = None
        if isinstance(widths, list) and len(widths) == 5:
            try:
                for idx, w in enumerate(widths):
                    self.table.setColumnWidth(idx, int(w))
            except Exception:
                pass
            # 论点允许手动调整，关键词/预览自动拉伸，操作/状态固定
            try:
                header.setSectionResizeMode(0, QHeaderView.Interactive)
                header.setSectionResizeMode(1, QHeaderView.Stretch)
                # 预览列保持拉伸，填充剩余空间，确保状态列贴齐右侧
                header.setSectionResizeMode(2, QHeaderView.Stretch)
                header.setSectionResizeMode(3, QHeaderView.Fixed)
                header.setSectionResizeMode(4, QHeaderView.Fixed)
            except Exception:
                pass
        else:
            # 默认：首列固定较宽，末列固定较窄；中间列按内容自适应一次
            try:
                self.table.setColumnWidth(0, 220)
                self._thumb_w, self._thumb_h = self._parse_size(self.settings.get("thumbnail_size", "160x120"), default=(140, 100))
                self.table.setColumnWidth(2, max(80, min(self._thumb_w + 8, 140)))
                # 操作列固定为紧凑宽度，显示两个按钮且留出间距
                self.table.setColumnWidth(3, 180)
                self.table.setColumnWidth(4, 72)
            except Exception:
                pass
            try:
                for c in [1, 2]:
                    self.table.resizeColumnToContents(c)
            except Exception:
                pass
            # 应用交互模式，关键词/预览随窗口拉伸，填充剩余宽度
            try:
                header.setSectionResizeMode(0, QHeaderView.Interactive)
                header.setSectionResizeMode(1, QHeaderView.Stretch)
                # 预览列保持拉伸
                header.setSectionResizeMode(2, QHeaderView.Stretch)
                header.setSectionResizeMode(3, QHeaderView.Fixed)
                header.setSectionResizeMode(4, QHeaderView.Fixed)
            except Exception:
                pass

    def _snapshot_row_heights(self):
        """保存当前所有行的高度到设置中。"""
        try:
            rows = self.table.rowCount()
            heights = [self.table.rowHeight(r) for r in range(rows)]
            self.settings["image_dialog_row_heights"] = heights
        except Exception:
            pass

    def _apply_saved_row_heights(self):
        """应用已保存的行高设定，如果存在。"""
        try:
            heights = self.settings.get("image_dialog_row_heights")
        except Exception:
            heights = None
        try:
            rows = self.table.rowCount()
        except Exception:
            rows = 0
        if isinstance(heights, list) and len(heights) == rows:
            try:
                # 应用已保存行高时，确保不低于当前默认行高
                try:
                    default_h = self.table.verticalHeader().defaultSectionSize()
                except Exception:
                    default_h = 48
                for r, h in enumerate(heights):
                    self.table.setRowHeight(r, max(default_h, int(h)))
                    # 行高应用后，同步更新预览控件尺寸
                    try:
                        label = self.table.cellWidget(r, 2)
                        if isinstance(label, QLabel):
                            target_w = max(64, self._thumb_w)
                            target_h = max(48, min(self._thumb_h, self.table.rowHeight(r) - 4))
                            label.setFixedSize(target_w, target_h)
                    except Exception:
                        pass
            except Exception:
                pass
        else:
            # 无保存时，按内容自适应行高，再确保不低于默认值
            try:
                self.table.resizeRowsToContents()
                default_h = self.table.verticalHeader().defaultSectionSize()
                for r in range(rows):
                    if self.table.rowHeight(r) < default_h:
                        self.table.setRowHeight(r, default_h)
                    try:
                        label = self.table.cellWidget(r, 2)
                        if isinstance(label, QLabel):
                            target_w = max(64, self._thumb_w)
                            target_h = max(48, min(self._thumb_h, self.table.rowHeight(r) - 4))
                            label.setFixedSize(target_w, target_h)
                    except Exception:
                        pass
            except Exception:
                pass

    def _on_row_resized(self, logicalIndex: int, oldSize: int, newSize: int):
        """当某行被用户拖动调整高度时，按新高度同步预览控件大小。"""
        try:
            label = self.table.cellWidget(logicalIndex, 2)
            if isinstance(label, QLabel):
                target_w = max(64, self._thumb_w)
                target_h = max(48, min(self._thumb_h, newSize - 4))
                label.setFixedSize(target_w, target_h)
        except Exception:
            pass
        try:
            self._snapshot_row_heights()
        except Exception:
            pass

    def _save_layout(self):
        try:
            self._snapshot_widths()
            self._snapshot_row_heights()
            self.settings["image_dialog_size"] = {"width": self.width(), "height": self.height()}
            if self.settings_path:
                try:
                    ConfigManager.instance().save()
                except Exception:
                    save_settings(self.settings_path, self.settings)
        except Exception:
            pass

    def accept(self):
        # 关闭前自动保存当前配置，避免未点击“保存配置”导致丢失
        try:
            self._save_config()
        except Exception:
            pass
        try:
            self._save_layout()
        except Exception:
            pass
        super().accept()

    def closeEvent(self, event):
        # 关闭事件中也进行一次自动保存以确保持久化
        try:
            self._save_config()
        except Exception:
            pass
        try:
            self._save_layout()
        except Exception:
            pass
        return super().closeEvent(event)

    def _on_preview_hover(self, event, row: int, label: QLabel, entering: bool):
        """悬浮时的预览窗口（轻量级）。"""
        try:
            if entering:
                pm = getattr(label, 'pixmap', None)
                pm = pm() if callable(pm) else pm
                if not pm:
                    return
                # 创建轻量级提示窗口
                dlg = QDialog(self)
                dlg.setWindowFlags(Qt.ToolTip)
                v = QVBoxLayout(dlg)
                v.setContentsMargins(6,6,6,6)
                img_label = QLabel()
                img_label.setAlignment(Qt.AlignCenter)
                img_label.setPixmap(pm)
                v.addWidget(img_label)
                self._hover_preview_dlg = dlg
                # 放到鼠标附近
                pos = QCursor.pos()
                dlg.move(pos.x() + 12, pos.y() + 12)
                dlg.show()
            else:
                dlg = getattr(self, '_hover_preview_dlg', None)
                if dlg:
                    dlg.close()
                    self._hover_preview_dlg = None
        except Exception:
            pass

    @staticmethod
    def _parse_size(s: str, default=(160, 120)):
        try:
            parts = s.lower().replace("x", " ").split()
            w = int(parts[0])
            h = int(parts[1]) if len(parts) > 1 else w
            return max(32, w), max(32, h)
        except Exception:
            return default