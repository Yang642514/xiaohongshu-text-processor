from typing import Callable, List, Optional, Dict
import threading

from .utils import load_settings, save_settings


class ConfigManager:
    """全局配置管理器（单例）：统一管理 settings 的加载、保存与订阅通知。

    - 使用共享的字典对象，避免各对话框持有不同副本导致覆盖问题
    - 提供 reload/save/update 方法；如有需要可订阅变更事件
    """

    _instance: Optional["ConfigManager"] = None

    def __init__(self, settings_path: str):
        self.settings_path = settings_path
        self.settings: Dict = load_settings(settings_path)
        # 补齐可能缺失的默认项，确保初始化稳定
        try:
            self._ensure_defaults()
        except Exception:
            pass
        self._subscribers: List[Callable[[Dict], None]] = []
        # 线程安全：统一通过锁保护 settings 的读写与保存
        self._lock = threading.RLock()

    @classmethod
    def initialize(cls, settings_path: str) -> "ConfigManager":
        if cls._instance is None:
            cls._instance = ConfigManager(settings_path)
        else:
            cls._instance.settings_path = settings_path
            cls._instance.reload()
        return cls._instance

    @classmethod
    def instance(cls) -> "ConfigManager":
        if cls._instance is None:
            raise RuntimeError("ConfigManager not initialized. Call initialize(settings_path) first.")
        return cls._instance

    def subscribe(self, cb: Callable[[Dict], None]) -> None:
        try:
            if not cb:
                return
            with self._lock:
                if cb not in self._subscribers:
                    self._subscribers.append(cb)
        except Exception:
            pass

    def _notify(self) -> None:
        # 复制订阅者列表，避免回调中修改列表导致并发问题
        try:
            with self._lock:
                subscribers = list(self._subscribers)
                current_settings = self.settings
        except Exception:
            subscribers = []
            current_settings = self.settings
        for cb in subscribers:
            try:
                cb(current_settings)
            except Exception:
                # 忽略订阅回调中的异常
                pass

    def reload(self) -> Dict:
        with self._lock:
            self.settings = load_settings(self.settings_path)
            try:
                self._ensure_defaults()
            except Exception:
                pass
        self._notify()
        return self.settings

    def save(self) -> None:
        # 统一保存入口：加锁、保存并通知
        with self._lock:
            save_settings(self.settings_path, self.settings)
        self._notify()

    def update(self, updated: Dict, save: bool = True) -> None:
        with self._lock:
            try:
                self.settings.update(updated or {})
            finally:
                pass
        if save:
            self.save()

    def _ensure_defaults(self) -> None:
        """确保关键默认键存在，避免老版本设置文件缺失导致异常。

        仅补齐键，不覆盖已有值。
        """
        # 现版本不再补齐 Unsplash 相关默认项；依赖默认设置文件即可
        try:
            pass
        except Exception:
            pass