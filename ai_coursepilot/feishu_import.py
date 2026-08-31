from __future__ import annotations

"""Convert the user's real PDF + XLSX inputs into Feishu-ready records.

The parser intentionally keeps the raw timetable strings. AI is used for the
curriculum interpretation; timetable data remains deterministic and is never
invented by the model.
"""

import csv
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Any

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _col_to_idx(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx_sheets(path: str | Path) -> dict[str, list[list[str]]]:
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", NS):
                shared.append("".join(t.text or "" for t in si.iterfind(".//main:t", NS)))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rel}
        out: dict[str, list[list[str]]] = {}
        for sh in wb.find("main:sheets", NS):
            name = sh.attrib["name"]
            target = rel_map[sh.attrib[f"{{{NS['rel']}}}id"]]
            if not target.startswith("/"):
                target = "xl/" + target.lstrip("./")
            target = target.replace("xl/xl/", "xl/")
            root = ET.fromstring(z.read(target))
            rows: list[list[str]] = []
            for row in root.findall(".//main:sheetData/main:row", NS):
                cells: dict[int, str] = {}
                max_idx = -1
                for c in row.findall("main:c", NS):
                    m = re.match(r"([A-Z]+)\d+", c.attrib.get("r", ""))
                    if not m:
                        continue
                    idx = _col_to_idx(m.group(1))
                    max_idx = max(max_idx, idx)
                    typ = c.attrib.get("t")
                    if typ == "s":
                        v = c.find("main:v", NS)
                        value = shared[int(v.text)] if v is not None and v.text else ""
                    elif typ == "inlineStr":
                        value = "".join(t.text or "" for t in c.iterfind(".//main:t", NS))
                    else:
                        v = c.find("main:v", NS)
                        value = v.text if v is not None and v.text is not None else ""
                    cells[idx] = str(value).strip()
                row_values = [""] * (max_idx + 1)
                for idx, value in cells.items():
                    row_values[idx] = value
                rows.append(row_values)
            out[name] = rows
        return out


def _find_header(rows: list[list[str]], required: list[str]) -> tuple[int, list[str]]:
    wanted = set(required)
    for i, row in enumerate(rows):
        if wanted.issubset(set(row)):
            return i, row
    raise ValueError(f"找不到 Excel 表头：{required}")


def parse_course_xlsx(path: str | Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    sheets = read_xlsx_sheets(path)
    if "课程目录" not in sheets or "课程安排表" not in sheets:
        raise ValueError("Excel 必须包含“课程目录”和“课程安排表”两个工作表")

    di, dh = _find_header(sheets["课程目录"], ["课程编号", "课程名称", "学分"])
    si, sh = _find_header(
        sheets["课程安排表"],
        ["班级名称", "课程编号", "课程名称", "学分", "时间", "上课地址", "校区"],
    )

    directory: list[dict[str, Any]] = []
    for row in sheets["课程目录"][di + 1 :]:
        row = row + [""] * max(0, len(dh) - len(row))
        record = dict(zip(dh, row[: len(dh)]))
        if record.get("课程编号"):
            directory.append(record)

    schedule: list[dict[str, Any]] = []
    for row in sheets["课程安排表"][si + 1 :]:
        row = row + [""] * max(0, len(sh) - len(row))
        record = dict(zip(sh, row[: len(sh)]))
        if record.get("课程编号"):
            schedule.append(record)
    return directory, schedule


def make_feishu_course_records(
    directory: list[dict[str, Any]],
    schedule: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Return (course_directory_records, class_schedule_records)."""
    directory_records: list[dict[str, Any]] = []
    for r in directory:
        directory_records.append(
            {
                "课程编号": r.get("课程编号", ""),
                "课程名称": r.get("课程名称", ""),
                "学分": float(r.get("学分") or 0),
                "校区": r.get("校区", ""),
                "来源": "2026-2027秋季研究生课程目录及课程安排表",
            }
        )

    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for r in schedule:
        grouped[(r.get("课程编号", ""), r.get("班级名称", ""))].append(r)

    class_records: list[dict[str, Any]] = []
    for (code, class_name), rows in grouped.items():
        first = rows[0]
        times = " | ".join(x.get("时间", "") for x in rows if x.get("时间"))
        addresses = "; ".join(dict.fromkeys(x.get("上课地址", "") for x in rows if x.get("上课地址")))
        class_records.append(
            {
                "班级名称": class_name,
                "课程编号": code,
                "课程名称": first.get("课程名称", ""),
                "学分": float(first.get("学分") or 0),
                "时间": times,
                "上课地址": addresses,
                "校区": first.get("校区", ""),
                "数据状态": "已读取",
            }
        )
    return directory_records, class_records


def curriculum_to_feishu_records(schema: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for c in schema.courses:
        records.append(
            {
                "课程编号": c.code,
                "课程名称": c.name,
                "学分": c.credit,
                "课程类别": c.category,
                "分组": c.group_id if c.group_id is not None else "",
                "分组规则": c.group_rule,
                "必修": "是" if c.required else "否",
                "学期": c.semester or "",
                "AI解析状态": "已解析",
            }
        )
    return records
