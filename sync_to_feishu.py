from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

from ai_coursepilot.ai_client import AIClient
from ai_coursepilot.feishu_client import FeishuClient
from ai_coursepilot.feishu_import import curriculum_to_feishu_records, make_feishu_course_records, parse_course_xlsx
from ai_coursepilot.ingest import ingest_curriculum


def main() -> int:
    parser = argparse.ArgumentParser(description="AI解析培养方案 + 读取课程Excel + 写入飞书多维表格")
    parser.add_argument("--pdf", required=True, help="培养方案 PDF")
    parser.add_argument("--xlsx", required=True, help="课程目录及课程安排 Excel")
    parser.add_argument("--curriculum-table", required=True, help="培养方案多维表格 table_id")
    parser.add_argument("--course-table", required=True, help="课程目录多维表格 table_id")
    parser.add_argument("--class-table", required=True, help="课程班级多维表格 table_id")
    parser.add_argument("--replace", action="store_true", help="遇到同课程编号时更新已有记录")
    parser.add_argument("--output-json", default="ai_curriculum_result.json")
    parser.add_argument("--dry-run", action="store_true", help="只解析和转换，不写飞书")
    args = parser.parse_args()

    print("[1/4] AI 解析培养方案 PDF …")
    schema = ingest_curriculum(args.pdf, AIClient())
    print(f"      学校：{schema.school or '未识别'}")
    print(f"      专业：{schema.program or '未识别'}")
    print(f"      课程：{len(schema.courses)} 门")
    if schema.warnings:
        print("      AI 警告：")
        for w in schema.warnings[:20]:
            print(f"        - {w}")

    print("[2/4] 读取课程 Excel …")
    directory, schedule = parse_course_xlsx(args.xlsx)
    course_records, class_records = make_feishu_course_records(directory, schedule)
    curriculum_records = curriculum_to_feishu_records(schema)
    print(f"      课程目录：{len(course_records)} 条")
    print(f"      班级安排：{len(class_records)} 条")

    result = {
        "curriculum": schema.__dict__ | {"courses": [c.__dict__ for c in schema.courses]},
        "feishu_records": {
            "curriculum": curriculum_records,
            "course_directory": course_records,
            "class_schedule": class_records,
        },
    }
    Path(args.output_json).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.dry_run:
        print(f"[3/4] dry-run：结果已保存到 {args.output_json}")
        print("[4/4] 未写入飞书")
        return 0

    print("[3/4] 连接飞书并写入多维表格 …")
    fs = FeishuClient()
    r1 = fs.upsert_records(args.curriculum_table, curriculum_records, key_field="课程编号", replace_existing=args.replace)
    r2 = fs.upsert_records(args.course_table, course_records, key_field="课程编号", replace_existing=args.replace)
    # 班级表不能只以课程编号去重，因此按“课程编号+班级名称”检查后逐条创建/更新。
    existing = fs.list_records(args.class_table)
    existing_index = {}
    for item in existing:
        f = item.get("fields", {})
        key = f"{f.get('课程编号', '')}||{f.get('班级名称', '')}"
        if key.strip("|"):
            existing_index[key] = item.get("record_id") or item.get("id")
    created, updated, skipped = 0, 0, 0
    for record in class_records:
        key = f"{record.get('课程编号', '')}||{record.get('班级名称', '')}"
        if key in existing_index:
            if args.replace:
                fs.update_record(args.class_table, existing_index[key], record)
                updated += 1
            else:
                skipped += 1
        else:
            fs.create_record(args.class_table, record)
            created += 1

    print("[4/4] 完成")
    print(json.dumps({
        "curriculum_table": r1,
        "course_table": r2,
        "class_table": {"created": created, "updated": updated, "skipped": skipped},
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
