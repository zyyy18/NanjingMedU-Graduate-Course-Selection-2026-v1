from __future__ import annotations
import argparse,json
from ai_coursepilot.feishu_client import FeishuClient
from ai_coursepilot.ingest import ingest_curriculum

def main():
    ap=argparse.ArgumentParser(description="AI 培养方案解析器"); ap.add_argument("pdf"); ap.add_argument("--write-feishu-table",default=""); args=ap.parse_args()
    s=ingest_curriculum(args.pdf); out={"school":s.school,"program":s.program,"semester":s.semester,"courses":[c.to_legacy_dict() for c in s.courses],"warnings":s.warnings}
    if args.write_feishu_table:
        fs=FeishuClient()
        for c in s.courses: fs.create_record(args.write_feishu_table,{"课程编号":c.code,"课程名称":c.name,"学分":c.credit,"课程类别":c.category,"分组":c.group_id if c.group_id is not None else "","分组规则":c.group_rule,"必修":c.required,"学期":c.semester or ""})
        out["feishu_written"]=True
    print(json.dumps(out,ensure_ascii=False,indent=2))
if __name__=="__main__": main()
