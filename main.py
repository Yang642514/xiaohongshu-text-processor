import os
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from app.core.utils import load_settings, save_settings
from app.core.config_manager import ConfigManager
from app.gui.main_window import MainWindow
from typing import Optional


def _ensure_icon_files(icon_dir: str) -> Optional[str]:
    """Ensure logo.png and logo.ico exist, converting from logo.jpg/png if possible.
    Returns the best available icon path for runtime use."""
    icon_ico = os.path.join(icon_dir, "logo.ico")
    icon_png = os.path.join(icon_dir, "logo.png")
    icon_jpg = os.path.join(icon_dir, "logo.jpg")
    # If ICO exists, prefer it
    if os.path.exists(icon_ico):
        return icon_ico
    # Try to generate from JPG or PNG using Pillow if available
    src = icon_png if os.path.exists(icon_png) else (icon_jpg if os.path.exists(icon_jpg) else None)
    if src:
        try:
            from PIL import Image
            img = Image.open(src).convert("RGBA")
            # Save PNG (standardized size)
            try:
                img_png = img.copy()
                img_png = img_png.resize((256, 256), Image.LANCZOS)
                img_png.save(icon_png, format="PNG")
            except Exception:
                pass
            # Save ICO with multiple sizes
            sizes = [(16, 16), (32, 32), (64, 64), (128, 128), (256, 256)]
            img.save(icon_ico, sizes=sizes)
            return icon_ico
        except Exception:
            # Pillow not available or conversion failed; fall back to existing file
            pass
    # Fallback: return any existing image
    if os.path.exists(icon_png):
        return icon_png
    if os.path.exists(icon_jpg):
        return icon_jpg
    return None


def ensure_dirs(settings: dict):
    for k in ["output_dir", "zip_output_dir", "log_dir"]:
        d = settings.get(k)
        if d:
            os.makedirs(d, exist_ok=True)


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    settings_path = os.path.join(base_dir, "app", "config", "default_settings.json")

    # 初始化统一配置管理器并确保目录存在
    cfg = ConfigManager.initialize(settings_path)
    ensure_dirs(cfg.settings)

    app = QApplication(sys.argv)
    # 设置全局应用图标
    icon_dir = os.path.join(base_dir, "app", "static", "logo")
    icon_path = _ensure_icon_files(icon_dir)
    if icon_path and os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    # 加载样式表
    style_path = os.path.join(base_dir, "app", "gui", "style.qss")
    if os.path.exists(style_path):
        with open(style_path, 'r', encoding='utf-8') as f:
            app.setStyleSheet(f.read())
    win = MainWindow(settings_path=settings_path)
    # 同步窗口图标，确保任务栏显示图标
    if icon_path and os.path.exists(icon_path):
        win.setWindowIcon(QIcon(icon_path))
    win.resize(800, 600)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()