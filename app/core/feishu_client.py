import time
from typing import Dict, List, Optional, Tuple
import os
import json as _json
import logging

import requests


class FeishuClient:
    """简单封装飞书鉴权与多维表格查询。

    依赖 settings 中的以下字段：
    - feishu_app_id
    - feishu_app_secret
    - feishu_bitable_app_token
    - feishu_bitable_table_id
    - feishu_status_field_name (默认 "笔记状态")
    - feishu_status_done_value (默认 "已完成")
    - feishu_content_field_name (默认 "去敏感词")
    - feishu_page_size (默认 50)
    """

    AUTH_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    FIELDS_URL_TPL = "https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
    SEARCH_URL_TPL = "https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/search"

    def __init__(self, settings: Dict[str, str]):
        self.settings = settings
        self._token_cache: Tuple[str, float] = ("", 0.0)

    def _get_token(self) -> str:
        cached, ts = self._token_cache
        # 简单的 1 小时缓存
        if cached and (time.time() - ts) < 3600:
            return cached
        app_id = self.settings.get("feishu_app_id", "").strip()
        app_secret = self.settings.get("feishu_app_secret", "").strip()
        if not app_id or not app_secret:
            raise ValueError("缺少飞书 app_id 或 app_secret")
        resp = requests.post(self.AUTH_URL, json={"app_id": app_id, "app_secret": app_secret}, timeout=15)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"获取租户 token 失败: {data}")
        token = data.get("tenant_access_token")
        if not token:
            raise RuntimeError("未返回 tenant_access_token")
        self._token_cache = (token, time.time())
        return token

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self._get_token()}", "Content-Type": "application/json"}

    def get_fields_map(self, app_token: str, table_id: str) -> Dict[str, str]:
        """返回 字段名->字段ID 的映射。"""
        url = self.FIELDS_URL_TPL.format(app_token=app_token, table_id=table_id)
        resp = requests.get(url, headers=self._headers(), timeout=15)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"获取字段失败: {data}")
        fields = data.get("data", {}).get("items", [])
        # 兼容不同结构键名：使用 name/field_name 与 id/field_id
        mapping: Dict[str, str] = {}
        for f in fields:
            name = f.get("name") or f.get("field_name")
            fid = f.get("id") or f.get("field_id")
            if name and fid:
                mapping[str(name)] = str(fid)
        return mapping

    def search_done_records(self) -> List[Dict]:
        """按设置查询“笔记状态=已完成”的记录，返回记录列表。"""
        app_token = self.settings.get("feishu_bitable_app_token", "").strip()
        table_id = self.settings.get("feishu_bitable_table_id", "").strip()
        if not app_token or not table_id:
            raise ValueError("缺少飞书多维表格 app_token 或 table_id，请在配置中明确填写。")

        status_field_name = self.settings.get("feishu_status_field_name", "笔记状态")
        status_done_value = self.settings.get("feishu_status_done_value", "已完成")
        content_field_name = self.settings.get("feishu_content_field_name", "去敏感词")
        page_size = int(self.settings.get("feishu_page_size", 50))
        # 兼容 SDK 示例的可选参数：视图、字段名列表、排序、是否返回自动字段
        view_id = str(self.settings.get("feishu_view_id", "") or "").strip()
        field_names_setting = self.settings.get("feishu_field_names", "")
        sort_field_name = str(self.settings.get("feishu_sort_field_name", "") or "").strip()
        sort_desc_setting = str(self.settings.get("feishu_sort_desc", "true") or "true").strip().lower()
        automatic_fields_setting = str(self.settings.get("feishu_automatic_fields", "false") or "false").strip().lower()
        # 解析字段名列表：支持逗号分隔或 JSON 数组文本
        field_names: Optional[List[str]] = None
        try:
            if isinstance(field_names_setting, list):
                field_names = [str(x).strip() for x in field_names_setting if str(x).strip()]
            else:
                s = str(field_names_setting or "").strip()
                if s:
                    if s.startswith("[") and s.endswith("]"):
                        import json as __json
                        arr = __json.loads(s)
                        if isinstance(arr, list):
                            field_names = [str(x).strip() for x in arr if str(x).strip()]
                    else:
                        field_names = [t.strip() for t in s.split(',') if t.strip()]
        except Exception:
            field_names = None
        sort_desc = sort_desc_setting in ("1", "true", "yes", "y")
        automatic_fields = automatic_fields_setting in ("1", "true", "yes", "y")

        fields_map = self.get_fields_map(app_token, table_id)
        status_field_id = fields_map.get(status_field_name)
        content_field_id = fields_map.get(content_field_name)
        # 标题字段：尝试多候选名称
        title_field_id = None
        for _cand in ("标题", "笔记标题", "标题（人工审核）"):
            _fid = fields_map.get(_cand)
            if _fid:
                title_field_id = _fid
                break
        if not status_field_id:
            raise ValueError(f"未找到状态字段: {status_field_name}")
        if not content_field_id:
            raise ValueError(f"未找到内容字段: {content_field_name}")
        # 记录关键映射，方便调试
        try:
            logger = logging.getLogger("xhs_tool")
            logger.info(
                "[FeishuClient] 字段映射: status='%s'->%s, content='%s'->%s, title='标题'->%s",
                status_field_name,
                status_field_id,
                content_field_name,
                content_field_id,
                title_field_id,
            )
        except Exception:
            pass

        url = self.SEARCH_URL_TPL.format(app_token=app_token, table_id=table_id)
        items: List[Dict] = []
        raw_items: List[Dict] = []
        last_payload_snapshot: Dict = {}
        page_token: Optional[str] = None
        # 使用 REST 路径：强制字段名条件筛选 + 分页
        while True:
            payload = {
                "filter": {
                    "conjunction": "and",
                    "conditions": [
                        {
                            "field_name": status_field_name,
                            "operator": "is",
                            "value": [status_done_value],
                        }
                    ],
                },
                "page_size": page_size,
            }
            if view_id:
                payload.setdefault("view_id", view_id)
            if field_names:
                payload.setdefault("field_names", field_names)
            if sort_field_name:
                payload.setdefault("sort", [{"field_name": sort_field_name, "desc": sort_desc}])
            payload.setdefault("automatic_fields", automatic_fields)
            if page_token:
                payload["page_token"] = page_token

            last_payload_snapshot = payload
            resp = requests.post(url, headers=self._headers(), json=payload, timeout=20)
            data = resp.json()
            if data.get("code") != 0:
                raise RuntimeError(f"搜索记录失败: {data} | payload={payload}")

            batch = data.get("data", {}).get("items", [])
            if isinstance(batch, list):
                raw_items.extend(batch)
            for rec in batch:
                fields = rec.get("fields", {})
                # 同时支持按字段名与字段ID读取（API可能返回 name-keyed 或 id-keyed）
                def _get_field_value(_fields: Dict, fid: Optional[str], fname: Optional[str]):
                    if not isinstance(_fields, dict):
                        return None
                    # 优先按字段名读取（更直观，且与过滤保持一致）
                    if fname and fname in _fields:
                        return _fields.get(fname)
                    if fid and fid in _fields:
                        return _fields.get(fid)
                    return None
                content_val = _get_field_value(fields, content_field_id, content_field_name)
                status_val = _get_field_value(fields, status_field_id, status_field_name)
                title_val = None
                # 先按候选名称读取标题；未命中再尝试候选ID
                for _cand in ("标题", "笔记标题", "标题（人工审核）"):
                    if _cand in fields:
                        title_val = fields.get(_cand)
                        break
                if title_val is None and (title_field_id is not None):
                    title_val = fields.get(title_field_id)
                def _to_text(v):
                    if v is None:
                        return ""
                    if isinstance(v, list):
                        try:
                            if all(isinstance(x, dict) for x in v):
                                parts: List[str] = []
                                for x in v:
                                    extracted = None
                                    for key in ("text", "value", "content", "name"):
                                        if key in x and isinstance(x[key], (str, int, float)):
                                            extracted = str(x[key])
                                            break
                                    parts.append(extracted if extracted is not None else str(x))
                                return "\n".join(parts)
                        except Exception:
                            pass
                        return "\n".join(str(x) for x in v)
                    if isinstance(v, dict):
                        try:
                            for key in ("text", "value", "content"):
                                if key in v and isinstance(v[key], (str, int, float)):
                                    return str(v[key])
                        except Exception:
                            pass
                        try:
                            return _json.dumps(v, ensure_ascii=False)
                        except Exception:
                            return str(v)
                    return str(v)
                try:
                    logger = logging.getLogger("xhs_tool")
                    if logger.isEnabledFor(logging.DEBUG):
                        logger.debug(
                            "[FeishuClient] rec_id=%s fields_keys=%s content_raw_type=%s status_raw_type=%s title_raw_type=%s",
                            rec.get("record_id"),
                            list(fields.keys()) if isinstance(fields, dict) else type(fields),
                            type(content_val).__name__,
                            type(status_val).__name__,
                            type(title_val).__name__ if title_field_id else None,
                        )
                except Exception:
                    pass
                content_text = _to_text(content_val)
                if not content_text:
                    for _cand_name in ("发布内容一键复制", "笔记信息"):
                        # 先按字段名尝试
                        if _cand_name in fields:
                            _alt = fields.get(_cand_name)
                            _alt_text = _to_text(_alt)
                            if _alt_text:
                                try:
                                    logger = logging.getLogger("xhs_tool")
                                    logger.info(
                                        "[FeishuClient] 内容为空，回退字段 '%s' 命中（按名称），rec_id=%s",
                                        _cand_name,
                                        rec.get("record_id"),
                                    )
                                except Exception:
                                    pass
                                content_val = _alt
                                content_text = _alt_text
                                break
                        # 再按字段ID尝试
                        _fid = fields_map.get(_cand_name)
                        if _fid and _fid in fields:
                            _alt = fields.get(_fid)
                            _alt_text = _to_text(_alt)
                            if _alt_text:
                                try:
                                    logger = logging.getLogger("xhs_tool")
                                    logger.info(
                                        "[FeishuClient] 内容为空，回退字段 '%s' 命中（按ID），rec_id=%s",
                                        _cand_name,
                                        rec.get("record_id"),
                                    )
                                except Exception:
                                    pass
                                content_val = _alt
                                content_text = _alt_text
                                break
                items.append({
                    "record_id": rec.get("record_id"),
                    "title": _to_text(title_val) or rec.get("record_id"),
                    "status": _to_text(status_val),
                    "content": content_text,
                    "updated_time": rec.get("updated_time"),
                })
            page_token = data.get("data", {}).get("page_token")
            if not page_token:
                break
        # 输出原始响应快照与映射辅助信息与搜索请求快照
        try:
            log_dir = self.settings.get("log_dir", "logs") or "logs"
            os.makedirs(log_dir, exist_ok=True)
            raw_path = os.path.join(log_dir, "feishu_refresh_raw_response.json")
            meta_path = os.path.join(log_dir, "feishu_refresh_debug_meta.json")
            req_path = os.path.join(log_dir, "feishu_refresh_search_payload.json")
            with open(raw_path, "w", encoding="utf-8") as rf:
                _json.dump(raw_items, rf, ensure_ascii=False, indent=2)
            with open(meta_path, "w", encoding="utf-8") as mf:
                _json.dump({
                    "status_field_name": status_field_name,
                    "status_field_id": status_field_id,
                    "content_field_name": content_field_name,
                    "content_field_id": content_field_id,
                    "title_field_id": title_field_id,
                    "fields_map": fields_map,
                    "view_id": view_id,
                    "field_names": field_names,
                    "sort_field_name": sort_field_name,
                    "sort_desc": sort_desc,
                    "automatic_fields": automatic_fields,
                    "items_count": len(items),
                }, mf, ensure_ascii=False, indent=2)
            with open(req_path, "w", encoding="utf-8") as qf:
                _json.dump(last_payload_snapshot, qf, ensure_ascii=False, indent=2)
            try:
                logger = logging.getLogger("xhs_tool")
                logger.info(
                    "[FeishuClient] 原始响应快照: %s | 映射信息: %s | 请求快照: %s",
                    os.path.abspath(raw_path),
                    os.path.abspath(meta_path),
                    os.path.abspath(req_path),
                )
            except Exception:
                pass
        except Exception:
            pass
        return items

    def mark_record_as_used(self, payload: Dict) -> None:
        """将指定记录的状态字段更新为“已使用”。

        - 使用 settings 中的 feishu_status_field_name（默认“笔记状态”）
        - 使用 settings 中的 feishu_status_used_value（默认“已使用”）
        - 单选字段需写入选项ID列表；若找不到选项ID，则不更新（避免报错）
        """
        app_token = self.settings.get("feishu_bitable_app_token", "").strip()
        table_id = self.settings.get("feishu_bitable_table_id", "").strip()
        if not app_token or not table_id:
            # 若未设置，尝试派生（与 search 中一致）
            link = self.settings.get("feishu_bitable_link", "").strip()
            if not link:
                raise ValueError("缺少飞书多维表格链接或派生参数")
            # 复用派生逻辑：简单调用一次搜索以确保派生填充
            _ = self.search_done_records()
            app_token = self.settings.get("feishu_bitable_app_token", "").strip()
            table_id = self.settings.get("feishu_bitable_table_id", "").strip()
            if not app_token or not table_id:
                raise ValueError("无法派生 app_token/table_id")

        record_id = payload.get("record_id")
        if not record_id:
            raise ValueError("缺少 record_id，无法更新状态")

        status_field_name = self.settings.get("feishu_status_field_name", "笔记状态")
        used_value_name = self.settings.get("feishu_status_used_value", "已使用")
        fields_map = self.get_fields_map(app_token, table_id)
        status_field_id = fields_map.get(status_field_name)
        if not status_field_id:
            raise ValueError(f"未找到状态字段: {status_field_name}")

        # 获取单选选项ID
        opt_id = None
        try:
            meta_url = self.FIELDS_URL_TPL.format(app_token=app_token, table_id=table_id)
            m_resp = requests.get(meta_url, headers=self._headers(), timeout=15)
            m_data = m_resp.json()
            if m_data.get("code") == 0:
                items = m_data.get("data", {}).get("items", []) or []
                target_meta = None
                for it in items:
                    fid = it.get("id") or it.get("field_id")
                    fname = it.get("name") or it.get("field_name")
                    if fid == status_field_id or fname == status_field_name:
                        target_meta = it
                        break
                if isinstance(target_meta, dict):
                    prop = target_meta.get("property") or {}
                    options = prop.get("options") or []
                    for o in options:
                        if not isinstance(o, dict):
                            continue
                        name = o.get("name")
                        oid = o.get("id") or o.get("option_id")
                        if name and oid and str(name) == str(used_value_name):
                            opt_id = str(oid)
                            break
        except Exception:
            pass

        if not opt_id:
            # 找不到目标选项时不更新，直接返回
            return

        url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
        body = {"fields": {status_field_id: [opt_id]}}
        resp = requests.patch(url, headers=self._headers(), json=body, timeout=15)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"更新记录状态失败: {data} | body={body}")