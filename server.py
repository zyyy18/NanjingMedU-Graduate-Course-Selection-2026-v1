from __future__ import annotations
import json,os
from fastapi import FastAPI,Request
from fastapi.responses import JSONResponse
from ai_coursepilot.service import CoursePilotService
app=FastAPI(title="AI CoursePilot for Feishu")
@app.get("/health")
def health(): return {"ok":True,"service":"ai-coursepilot"}
@app.post("/feishu/webhook")
async def feishu_webhook(request:Request):
    p=await request.json()
    if p.get("type")=="url_verification": return JSONResponse({"challenge":p.get("challenge","")})
    m=p.get("event",{}).get("message",{}); content=m.get("content","")
    try: text=(json.loads(content) if isinstance(content,str) else content).get("text","")
    except: text=str(content)
    try:
        r=CoursePilotService().run_from_feishu(os.getenv("FEISHU_CURRICULUM_TABLE_ID",""),os.getenv("FEISHU_CLASS_TABLE_ID",""),os.getenv("FEISHU_RESULT_TABLE_ID",""),text.strip() or "请生成最合适的选课方案。",n=5); reply=r["explanation"]
    except Exception as e: reply=f"选课助手执行失败：{e}"
    return JSONResponse({"msg_type":"text","content":{"text":reply}})
