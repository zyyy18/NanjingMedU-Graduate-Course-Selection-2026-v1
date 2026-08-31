from __future__ import annotations
import json, os
from urllib import parse, request

class FeishuClientError(RuntimeError): pass
class FeishuClient:
    BASE="https://open.feishu.cn/open-apis"
    def __init__(self,app_id=None,app_secret=None,app_token=None):
        self.app_id=app_id or os.getenv("FEISHU_APP_ID",""); self.app_secret=app_secret or os.getenv("FEISHU_APP_SECRET",""); self.app_token=app_token or os.getenv("FEISHU_BITABLE_APP_TOKEN",""); self._token=""
    def _json_request(self,method,url,body=None,token=None):
        data=None if body is None else json.dumps(body,ensure_ascii=False).encode(); headers={"Content-Type":"application/json; charset=utf-8"}
        if token: headers["Authorization"]=f"Bearer {token}"
        try:
            with request.urlopen(request.Request(url,data=data,headers=headers,method=method),timeout=30) as r: p=json.loads(r.read().decode())
        except Exception as e: raise FeishuClientError(f"飞书 API 请求失败：{e}") from e
        if p.get("code",0)!=0: raise FeishuClientError(f"飞书 API 返回错误：{p.get('code')} {p.get('msg')}")
        return p
    def tenant_access_token(self):
        if self._token:return self._token
        if not self.app_id or not self.app_secret: raise FeishuClientError("未配置 FEISHU_APP_ID / FEISHU_APP_SECRET")
        p=self._json_request("POST",f"{self.BASE}/auth/v3/tenant_access_token/internal",{"app_id":self.app_id,"app_secret":self.app_secret})
        self._token=p["tenant_access_token"]; return self._token
    def list_records(self,table_id,page_size=500):
        rows=[]; page_token=""; token=self.tenant_access_token()
        while True:
            q={"page_size":str(min(page_size,500))};
            if page_token:q["page_token"]=page_token
            p=self._json_request("GET",f"{self.BASE}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records?{parse.urlencode(q)}",token=token); d=p.get("data",{}); rows.extend(d.get("items",[]))
            if not d.get("has_more"): break
            page_token=d.get("page_token","");
            if not page_token: break
        return rows
    def create_record(self,table_id,fields):
        p=self._json_request("POST",f"{self.BASE}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records",{"fields":fields},token=self.tenant_access_token()); return p.get("data",{}).get("record",{})
    def send_message(self,receive_id,text,receive_id_type="open_id"):
        p=self._json_request("POST",f"{self.BASE}/im/v1/messages?{parse.urlencode({'receive_id_type':receive_id_type})}",{"receive_id":receive_id,"msg_type":"text","content":json.dumps({'text':text},ensure_ascii=False)},token=self.tenant_access_token()); return p
