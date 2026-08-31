from __future__ import annotations
from pathlib import Path
from typing import Any
from .ai_client import AIClient
from .schema import CurriculumCourse,CurriculumSchema

def extract_pdf_text(pdf_path: str|Path)->str:
    from pypdf import PdfReader
    return "\n".join(page.extract_text() or "" for page in PdfReader(Path(pdf_path)).pages).strip()

def curriculum_from_ai(raw:dict[str,Any])->CurriculumSchema:
    courses=[]
    for x in raw.get("courses",[]):
        code=str(x.get("code","")).strip(); name=str(x.get("name","")).strip()
        if not code or not name: continue
        gid=x.get("group_id")
        try: gid=int(gid) if gid not in (None,"") else None
        except: gid=None
        courses.append(CurriculumCourse(code,name,float(x.get("credit",0) or 0),str(x.get("category","") or ""),gid,str(x.get("group_rule","") or ""),bool(x.get("required",False)),int(x.get("semester",0) or 0) or None))
    return CurriculumSchema(str(raw.get("school","") or ""),str(raw.get("program","") or ""),str(raw.get("semester","") or ""),courses,warnings=[str(x) for x in raw.get("warnings",[])])

def ingest_curriculum(pdf_path,ai=None):
    text=extract_pdf_text(pdf_path)
    if not text: raise ValueError("PDF 未提取到文本；扫描件需要先接 OCR")
    return curriculum_from_ai((ai or AIClient()).parse_curriculum(text))
