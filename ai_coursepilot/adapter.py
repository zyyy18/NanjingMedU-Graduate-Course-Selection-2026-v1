from __future__ import annotations
from collections import defaultdict
from types import SimpleNamespace
from typing import Any
try:
    from course_selector_gui_v15 import CourseOption, parse_time
except ImportError:
    from legacy_course_selector import CourseOption, parse_time

def _v(r,*names,default=""):
    f=r.get("fields",r)
    for n in names:
        x=f.get(n)
        if x not in (None,""): return x
    return default

def build_engine_from_records(curriculum_records,class_records):
    plan=[]
    for r in curriculum_records:
        code=str(_v(r,"课程编号","code")).strip(); name=str(_v(r,"课程名称","name")).strip()
        if not code or not name: continue
        gid=_v(r,"分组","group_id",default="")
        try: gid=int(gid) if gid not in ("",None) else None
        except: gid=None
        plan.append({"课程编号":code,"课程名称":name,"学分":float(_v(r,"学分","credit",default=0) or 0),"课程类别":str(_v(r,"课程类别","category",default="")),"分组":gid,"分组规则":str(_v(r,"分组规则","group_rule",default="")),"学期":int(_v(r,"学期","semester",default=0) or 0),"学时":int(_v(r,"学时","hours",default=0) or 0)})
    by={x["课程编号"]:x for x in plan}; rows=defaultdict(list)
    for r in class_records:
        code=str(_v(r,"课程编号","code")).strip()
        if code: rows[code].append(r)
    options={}
    for code,c in by.items():
        rs=rows.get(code,[])
        if not rs:
            options[code]=[CourseOption(plan_course=c,class_name="网课/待确认",course_code=code,course_name=c["课程名称"],credit=c["学分"],time_text="网课（时间不冲突）",address="线上",campus="网课",slots=set())]; continue
        grouped=defaultdict(list)
        for r in rs: grouped[str(_v(r,"班级名称","class_name",default="待定"))].append(r)
        opts=[]
        for cls,items in grouped.items():
            slots=set(); times=[]; addrs=[]; campuses=[]
            for r in items:
                tt=str(_v(r,"时间","time",default="")); slots|=parse_time(tt)
                if tt:times.append(tt)
                a=str(_v(r,"上课地址","address",default=""));
                if a:addrs.append(a)
                cp=str(_v(r,"校区","campus",default="网课"));
                if cp:campuses.append(cp)
            opts.append(CourseOption(plan_course=c,class_name=cls,course_code=code,course_name=str(_v(items[0],"课程名称","name",default=c["课程名称"])),credit=float(_v(items[0],"学分","credit",default=c["学分"]) or c["学分"]),time_text=" | ".join(times),address="; ".join(dict.fromkeys(addrs)),campus=campuses[0] if campuses else "网课",slots=slots))
        options[code]=opts
    return SimpleNamespace(plan=plan,plan_by_code=by,options_by_code=options,courses_with_schedule={c:o for c,o in options.items() if o and not o[0].is_virtual},courses_without_schedule=[c for c,o in options.items() if o and o[0].is_virtual])
