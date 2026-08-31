from __future__ import annotations
import json
try:
    from course_selector_gui_v15 import Planner,UserPreference,special_requirements,parse_simple_block
except ImportError:
    from legacy_course_selector import Planner,UserPreference,special_requirements,parse_simple_block
from .adapter import build_engine_from_records
from .ai_client import AIClient
from .feishu_client import FeishuClient
from .schema import UserIntent

class CoursePilotService:
    def __init__(self,ai=None,feishu=None): self.ai=ai or AIClient(); self.feishu=feishu or FeishuClient()
    @staticmethod
    def _intent(raw,original):
        return UserIntent(original,[str(x) for x in raw.get("avoid_time",[])],str(raw.get("preferred_campus","") or ""),int(raw.get("max_other_campus_days",3) or 0),[str(x) for x in raw.get("exempt_courses",[])],[str(x) for x in raw.get("unwanted_courses",[])],{str(k):int(v) for k,v in (raw.get("preferred_courses") or {}).items()},{str(k):int(v) for k,v in (raw.get("preferred_classes") or {}).items()},[str(x) for x in raw.get("required_courses",[])],{str(k):int(v) for k,v in (raw.get("required_classes") or {}).items()},{str(k):float(v) for k,v in (raw.get("objective_weights") or {}).items()})
    @staticmethod
    def _resolve(engine,keys):
        out=set()
        for key in keys:
            if key in engine.plan_by_code: out.add(key); continue
            for code,c in engine.plan_by_code.items():
                if key.replace(" ","")==str(c["课程名称"]).replace(" ",""): out.add(code)
        return out
    def _pref(self,engine,i):
        avoid=set()
        for x in i.avoid_time: avoid|=parse_simple_block(x)
        scores={}
        for k,v in i.preferred_courses.items():
            for c in self._resolve(engine,[k]): scores[c]=max(scores.get(c,50),v)
        for k in i.required_courses:
            for c in self._resolve(engine,[k]): scores[c]=100
        classes={}
        for k,v in i.preferred_classes.items():
            if "|" in k:
                code,cls=k.split("|",1)
                for c in self._resolve(engine,[code]): classes[(c,cls)]=max(classes.get((c,cls),50),v)
        for k in i.required_classes:
            if "|" in k:
                code,cls=k.split("|",1)
                for c in self._resolve(engine,[code]): classes[(c,cls)]=100
        return UserPreference(avoid_slots=avoid,preferred_campus=i.preferred_campus,max_other_campus_days=i.max_other_campus_days,exempt_codes=self._resolve(engine,i.exempt_courses),unwanted_codes=self._resolve(engine,i.unwanted_courses),preferred_scores=scores,class_scores=classes)
    def run_from_feishu(self,curriculum_table,class_table,result_table,user_request,n=5):
        engine=build_engine_from_records(self.feishu.list_records(curriculum_table),self.feishu.list_records(class_table))
        raw=self.ai.parse_user_intent(user_request,context=json.dumps({"课程数":len(engine.plan),"课程":engine.plan},ensure_ascii=False)); intent=self._intent(raw,user_request); pref=self._pref(engine,intent)
        results=Planner(engine,pref).generate(n=n)
        if not results: raise RuntimeError("当前约束下没有找到可行选课方案。")
        rows=[]
        for rank,(score,selected) in enumerate(results,1):
            row={"方案":rank,"偏好得分":-float(score[0]),"课程数":len(selected),"总学分":sum(float(x.credit) for x in selected),"特殊需求":special_requirements(selected,pref),"课程":[x.as_dict() for x in selected]}; rows.append(row)
        explanation=self.ai.explain_solution(user_request,rows[0],rows[1:]); rows[0]["AI推荐理由"]=explanation
        for r in rows: self.feishu.create_record(result_table,{"方案":r["方案"],"总学分":r["总学分"],"课程数":r["课程数"],"AI推荐理由":r.get("AI推荐理由",""),"课程JSON":json.dumps(r["课程"],ensure_ascii=False),"特殊需求":json.dumps(r["特殊需求"],ensure_ascii=False),"用户需求":user_request})
        return {"intent":raw,"solutions":rows,"explanation":explanation}
