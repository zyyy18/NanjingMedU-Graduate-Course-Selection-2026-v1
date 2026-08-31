from __future__ import annotations

import json
import os
from urllib import parse, request

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


class FeishuClientError(RuntimeError):
    pass


class FeishuClient:
    BASE = "https://open.feishu.cn/open-apis"

    def __init__(self, app_id=None, app_secret=None, app_token=None):
        self.app_id = app_id or os.getenv("FEISHU_APP_ID", "")
        self.app_secret = app_secret or os.getenv("FEISHU_APP_SECRET", "")
        self.app_token = app_token or os.getenv("FEISHU_BITABLE_APP_TOKEN", "")
        self._token = ""

    def _json_request(self, method, url, body=None, token=None):
        data = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        try:
            with request.urlopen(request.Request(url, data=data, headers=headers, method=method), timeout=60) as r:
                payload = json.loads(r.read().decode("utf-8"))
        except Exception as exc:
            raise FeishuClientError(f"飞书 API 请求失败：{exc}") from exc
        if payload.get("code", 0) != 0:
            raise FeishuClientError(f"飞书 API 返回错误：{payload.get('code')} {payload.get('msg')}")
        return payload

    def tenant_access_token(self) -> str:
        if self._token:
            return self._token
        if not self.app_id or not self.app_secret:
            raise FeishuClientError("未配置 FEISHU_APP_ID / FEISHU_APP_SECRET")
        p = self._json_request(
            "POST",
            f"{self.BASE}/auth/v3/tenant_access_token/internal",
            {"app_id": self.app_id, "app_secret": self.app_secret},
        )
        self._token = p["tenant_access_token"]
        return self._token

    def list_records(self, table_id: str, page_size: int = 500) -> list[dict]:
        rows, page_token = [], ""
        while True:
            q = {"page_size": str(min(max(page_size, 1), 500))}
            if page_token:
                q["page_token"] = page_token
            url = f"{self.BASE}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records?{parse.urlencode(q)}"
            p = self._json_request("GET", url, token=self.tenant_access_token())
            data = p.get("data", {})
            rows.extend(data.get("items", []))
            if not data.get("has_more"):
                break
            page_token = data.get("page_token", "")
            if not page_token:
                break
        return rows

    def create_record(self, table_id: str, fields: dict) -> dict:
        p = self._json_request(
            "POST",
            f"{self.BASE}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records",
            {"fields": fields},
            token=self.tenant_access_token(),
        )
        return p.get("data", {}).get("record", {})

    def update_record(self, table_id: str, record_id: str, fields: dict) -> dict:
        p = self._json_request(
            "PUT",
            f"{self.BASE}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records/{record_id}",
            {"fields": fields},
            token=self.tenant_access_token(),
        )
        return p.get("data", {}).get("record", {})

    def upsert_records(self, table_id: str, records: list[dict], key_field: str = "课程编号", replace_existing: bool = False) -> dict:
        existing = self.list_records(table_id)
        index = {}
        for item in existing:
            fields = item.get("fields", {})
            value = fields.get(key_field)
            if value not in (None, ""):
                index[str(value)] = item.get("record_id") or item.get("id")
        created, updated, skipped = [], [], []
        for record in records:
            key = str(record.get(key_field, "")).strip()
            if key and key in index:
                if replace_existing:
                    self.update_record(table_id, index[key], record)
                    updated.append(key)
                else:
                    skipped.append(key)
            else:
                created.append(self.create_record(table_id, record))
        return {"created": created, "updated": updated, "skipped": skipped}

    def send_message(self, receive_id: str, text: str, receive_id_type: str = "open_id"):
        return self._json_request(
            "POST",
            f"{self.BASE}/im/v1/messages?{parse.urlencode({'receive_id_type': receive_id_type})}",
            {"receive_id": receive_id, "msg_type": "text", "content": json.dumps({'text': text}, ensure_ascii=False)},
            token=self.tenant_access_token(),
        )
