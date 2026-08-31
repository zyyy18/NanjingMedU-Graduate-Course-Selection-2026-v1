from __future__ import annotations
import json, os, re
from typing import Any
from urllib import request

class AIClientError(RuntimeError): pass

class AIClient:
    def __init__(self, base_url=None, api_key=None, model=None):
        self.base_url=(base_url or os.getenv("AI_BASE_URL","https://api.openai.com/v1")).rstrip("/")
        self.api_key=api_key or os.getenv("AI_API_KEY","")
        self.model=model or os.getenv("AI_MODEL","gpt-4o-mini")
        self.timeout=float(os.getenv("AI_TIMEOUT","60"))
    def _post(self,payload):
        if not self.api_key: raise AIClientError("未配置 AI_API_KEY")
        req=request.Request(f"{self.base_url}/chat/completions",data=json.dumps(payload,ensure_ascii=False).encode(),headers={"Authorization":f"Bearer {self.api_key}","Content-Type":"application/json"},method="POST")
        try:
            with request.urlopen(req,timeout=self.timeout) as r: data=json.loads(r.read().decode())
        except Exception as e: raise AIClientError(f"AI 请求失败：{e}") from e
        return data
    @staticmethod
    def _extract_json(text):
        text=text.strip()
        if text.startswith("```"): text=re.sub(r"^```(?:json)?\s*|\s*```$","",text,flags=re.I)
        try: return json.loads(text)
        except json.JSONDecodeError:
            m=re.search(r"\{.*\}",text,re.S)
            if not m: raise AIClientError("AI 未返回可解析 JSON")
            return json.loads(m.group(0))
    def chat_json(self,system_prompt,user_prompt):
        payload={"model":self.model,"messages":[{"role":"system","content":system_prompt},{"role":"user","content":user_prompt}],"temperature":.1,"response_format":{"type":"json_object"}}
        try: data=self._post(payload)
        except AIClientError:
            payload.pop("response_format",None); data=self._post(payload)
        try: return self._extract_json(data["choices"][0]["message"]["content"])
        except (KeyError,IndexError,TypeError) as e: raise AIClientError("AI 响应格式异常") from e
    def parse_curriculum(self,pdf_text):
        system='''你是高校研究生培养方案结构化专家。把培养方案原文转换成严格 JSON，不编造课程。保留分组规则原文。无法确认的信息置空并写入 warnings。输出 {"school":"","program":"","semester":"","courses":[{"code":"","name":"","credit":0,"category":"","group_id":null,"group_rule":"","required":false,"semester":null}],"warnings":[]}'''
        return self.chat_json(system,pdf_text)
    def parse_user_intent(self,natural_language,context=""):
        system='''你是智能选课需求分析助手。把自然语言转换成结构化约束，不自行决定具体课程。评分 0-100，100 为硬性必选。输出 avoid_time、preferred_campus、max_other_campus_days、exempt_courses、unwanted_courses、preferred_courses、preferred_classes、required_courses、required_classes、objective_weights、warnings。'''
        return self.chat_json(system,f"培养方案/课程上下文：\n{context}\n\n用户需求：\n{natural_language}")
    def explain_solution(self,request_text,solution,alternatives):
        data=self._post({"model":self.model,"messages":[{"role":"system","content":"你是选课决策解释助手，只能基于给定结果解释，不得虚构课程信息。用中文说明推荐理由、满足的偏好、妥协与备选方案差异。"},{"role":"user","content":json.dumps({"request":request_text,"recommended":solution,"alternatives":alternatives},ensure_ascii=False)}],"temperature":.2})
        try: return data["choices"][0]["message"]["content"].strip()
        except (KeyError,IndexError,TypeError) as e: raise AIClientError("AI 解释响应格式异常") from e
