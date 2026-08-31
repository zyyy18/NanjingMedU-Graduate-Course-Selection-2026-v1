from __future__ import annotations

import argparse
import json

from .service import CoursePilotService


def main():
    ap = argparse.ArgumentParser(description="AI + 飞书 + CP-SAT 智能选课助手")
    ap.add_argument("--curriculum-table", required=True)
    ap.add_argument("--class-table", required=True)
    ap.add_argument("--result-table", required=True)
    ap.add_argument("--request", required=True)
    ap.add_argument("-n", type=int, default=5)
    args = ap.parse_args()
    result = CoursePilotService().run_from_feishu(
        args.curriculum_table, args.class_table, args.result_table, args.request, n=args.n
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
