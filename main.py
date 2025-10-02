import os
import sys
from PyQt5.QtWidgets import QApplication
from app.core.utils import load_settings, save_settings
from app.gui.main_window import MainWindow


def ensure_dirs(settings: dict):
    for k in ["output_dir", "zip_output_dir", "log_dir"]:
        d = settings.get(k)
        if d:
            os.makedirs(d, exist_ok=True)


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    settings_path = os.path.join(base_dir, "app", "config", "default_settings.json")

    # 加载设置并确保目录存在
    settings = load_settings(settings_path)
    ensure_dirs(settings)

    app = QApplication(sys.argv)
    # 加载样式表
    style_path = os.path.join(base_dir, "app", "gui", "style.qss")
    if os.path.exists(style_path):
        with open(style_path, 'r', encoding='utf-8') as f:
            app.setStyleSheet(f.read())
    win = MainWindow(settings_path=settings_path)
    win.resize(800, 600)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()