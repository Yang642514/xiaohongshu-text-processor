import json
import os


def load_settings(path: str) -> dict:
    if not os.path.exists(path):
        raise FileNotFoundError(f"设置文件不存在: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_settings(path: str, data: dict):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)