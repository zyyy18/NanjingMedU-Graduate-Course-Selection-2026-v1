# -*- coding: utf-8 -*-
"""
研究生智能选课决策器 GUI v11
------------------------------------------------------------
功能：
1. 读取培养方案 PDF + 课程目录 Excel；
2. 弹窗输入：
   - 避开时间段
   - 首选校区
   - 每周最多 0/1/2/3 天去其他校区
   - 免修课程
   - 不想选的课程
   - 倾向课程 + 倾向评分；100% = 必选，其余默认 50%
3. 自动生成多个满足培养方案约束、课程时间不冲突的选课方案；
4. 同一天尽量集中同一校区；
5. 自动判断特殊需求是否满足；
6. 用 matplotlib 绘制多个方案的周课程表；
7. 导出 JSON / CSV / PNG；方案总览支持分页 PDF。
 8. 优先使用 OR-Tools CP-SAT 做约束优化；无 OR-Tools 时使用内置剪枝回退。
------------------------------------------------------------
依赖：
    pip install pypdf matplotlib
    推荐：pip install ortools

说明：
- Excel 读取使用标准 zip/xml，不依赖 pandas/openpyxl。
- 若某些课程在培养方案里存在、但当前课表中没有对应编号，程序会标记“待确认”，不会猜测。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import math
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path
from typing import Dict, List, Tuple, Set, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages

import matplotlib
import platform
import time

# OR-Tools 是首选求解器；未安装时自动退回内置分支定界求解器。
try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except Exception:
    cp_model = None
    ORTOOLS_AVAILABLE = False

# 自动根据系统选择中文字体，防止方块字
_system = platform.system()
if _system == "Windows":
    _fonts = ["SimHei", "Microsoft YaHei", "STHeiti"]
elif _system == "Darwin":  # macOS
    _fonts = ["PingFang SC", "Heiti SC", "Arial Unicode MS", "STHeiti"]
else:  # Linux
    _fonts = ["WenQuanYi Micro Hei", "Noto Sans CJK SC", "SimHei", "DejaVu Sans"]

_available = [f.name for f in matplotlib.font_manager.fontManager.ttflist]
_chosen = None
for f in _fonts:
    if f in _available:
        _chosen = f
        break

if _chosen:
    matplotlib.rcParams["font.family"] = _chosen
    matplotlib.rcParams["axes.unicode_minus"] = False
else:
    for f in _fonts:
        try:
            matplotlib.rcParams["font.sans-serif"] = [f]
            matplotlib.rcParams["axes.unicode_minus"] = False
            break
        except Exception:
            continue



# =========================
# 1. Excel reader
# =========================

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def col_to_idx(col_letters: str) -> int:
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def read_xlsx_sheets(path: Path) -> Dict[str, List[List[str]]]:
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", NS):
                shared.append(
                    "".join(t.text or "" for t in si.iterfind(".//main:t", NS))
                )

        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rel_root}

        out = {}
        for sh in wb_root.find("main:sheets", NS):
            name = sh.attrib["name"]
            rid = sh.attrib[f"{{{NS['rel']}}}id"]
            target = rel_map[rid]
            if not target.startswith("/"):
                target = "xl/" + target.lstrip("./")
            target = target.replace("xl/xl/", "xl/")

            root = ET.fromstring(z.read(target))
            rows = []

            for row in root.findall(".//main:sheetData/main:row", NS):
                cells = {}
                current_max = -1

                for c in row.findall("main:c", NS):
                    ref = c.attrib.get("r", "")
                    m = re.match(r"([A-Z]+)(\d+)", ref)
                    if not m:
                        continue

                    idx = col_to_idx(m.group(1))
                    current_max = max(current_max, idx)
                    typ = c.attrib.get("t")

                    if typ == "s":
                        v = c.find("main:v", NS)
                        value = shared[int(v.text)] if v is not None and v.text else ""
                    elif typ == "inlineStr":
                        value = "".join(
                            t.text or "" for t in c.iterfind(".//main:t", NS)
                        )
                    else:
                        v = c.find("main:v", NS)
                        value = v.text if v is not None and v.text is not None else ""

                    cells[idx] = value

                row_vals = [""] * (current_max + 1)
                for i, v in cells.items():
                    row_vals[i] = str(v).strip()
                rows.append(row_vals)

            out[name] = rows

        return out


def find_header(rows: List[List[str]], required: List[str]):
    for i, row in enumerate(rows):
        if all(x in row for x in required):
            return i, row
    raise ValueError(f"找不到表头：{required}")


# =========================
# 2. PDF / 培养方案
# =========================

def extract_pdf_text(path: Path) -> str:
    from pypdf import PdfReader

    text = "\n".join(page.extract_text() or "" for page in PdfReader(path).pages)
    return unicodedata.normalize("NFKC", text)


def parse_training_plan(pdf_text: str) -> List[dict]:
    text = re.sub(r"[ \t]+", " ", pdf_text)
    code_re = re.compile(r"[A-Z]\d{3}[a-z]{2}\d{3}[a-z]")
    matches = list(code_re.finditer(text))

    records = []

    for i, m in enumerate(matches):
        code = m.group(0)
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        seg = text[m.end():end]

        nums = list(re.finditer(r"(\d+(?:\.\d+)?)\s+(\d+)\s+(\d+)", seg))
        if not nums:
            continue

        nm = nums[-1]
        name = re.sub(r"\s+", "", seg[:nm.start()]).strip()
        name = re.sub(r"\d{4}/\d+/\d+\s+\d+:\d+\s+\S+", "", name)

        records.append(
            {
                "课程编号": code,
                "课程名称": name,
                "学分": float(nm.group(1)),
                "学时": int(float(nm.group(2))),
                "学期": int(nm.group(3)),
            }
        )

    # 针对本培养方案进行结构化
    for r in records:
        c = r["课程编号"]

        if c in {
            "G210gw004a", "G220my025a", "G220wy003a",
            "G310jc009a", "X210yj001a", "X210yj002a", "X210yj003a"
        }:
            r["课程类别"] = "A"
        elif c == "Z320gx187a":
            r["课程类别"] = "B"
        elif c in {
            "Z320gw098a", "Z320jc133a", "Z320jc134a", "Z320jc143a",
            "Z320jc148a", "G310yl039a", "G310yl040a", "G310yl043a",
            "Z110yl264a", "Z120jc147a", "Z120yl259a", "Z310ek015a",
            "Z310yl248a", "Z310yl251a", "Z310yl254a", "Z310yl256a",
            "Z310yl257a", "Z310yl258a", "Z310yl262a", "Z310yl265a",
            "Z310yl306a"
        }:
            r["课程类别"] = "G"
        elif (
            c in {"G310ty044a", "G310ty045a", "X310jc043a"}
            or c.startswith("X310yj")
        ):
            r["课程类别"] = "H"
        else:
            r["课程类别"] = "未识别"

        r["分组"] = None
        r["分组规则"] = ""

        if c in {"X210yj001a", "X210yj002a", "X210yj003a"}:
            r["分组"] = 1
            r["分组规则"] = "第1组：3选1"
        elif c in {
            "Z320gw098a", "Z320jc133a", "Z320jc134a",
            "Z320jc143a", "Z320jc148a"
        }:
            r["分组"] = 3
            r["分组规则"] = "第3组：选2学分"
        elif r["课程类别"] == "G" and c != "Z320gx187a":
            r["分组"] = 4
            r["分组规则"] = "第4组：选2-3学分，至少2学分"
        elif r["课程类别"] == "H":
            r["分组"] = 2
            r["分组规则"] = "第2组：45选1"

    return records


# =========================
# 3. 时间解析
# =========================

DAY_MAP = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7}
DAY_NAME = {v: k for k, v in DAY_MAP.items()}


def parse_time(s: str) -> Set[Tuple[int, int, int]]:
    """返回 {(周, 星期, 节次)}"""
    slots = set()

    week_matches = list(re.finditer(r"(\d+)-(\d+)周", s))
    single_week = re.search(r"(\d+)周", s)

    if week_matches:
        w1, w2 = map(int, week_matches[-1].groups())
    elif single_week:
        w1 = w2 = int(single_week.group(1))
    else:
        # 某些系统用“7-9周;连续周”后不一定紧跟在前
        w1, w2 = 1, 20

    for dm in re.finditer(
        r"星期([一二三四五六日])\s*第(\d+)节-第(\d+)节", s
    ):
        day = DAY_MAP[dm.group(1)]
        p1, p2 = map(int, dm.group(2, 3))
        for w in range(w1, w2 + 1):
            for p in range(p1, p2 + 1):
                slots.add((w, day, p))

    return slots


def parse_simple_block(block: str) -> Set[Tuple[int, int]]:
    """
    给用户输入的避开时间解析成 {(星期, 节次)}。
    支持：
      周一7-10
      星期一 第7节-第10节
      一 7-10
    """
    out = set()
    day_m = re.search(r"(?:星期|周)?([一二三四五六日])", block)
    period_m = re.search(r"第?(\d+)\s*[-~到]\s*第?(\d+)", block)

    if day_m and period_m:
        d = DAY_MAP[day_m.group(1)]
        p1, p2 = map(int, period_m.groups())
        for p in range(p1, p2 + 1):
            out.add((d, p))
        return out

    # 例如“周一上午”
    if day_m:
        d = DAY_MAP[day_m.group(1)]
        if "上午" in block:
            for p in range(1, 7):
                out.add((d, p))
        elif "下午" in block:
            for p in range(8, 12):
                out.add((d, p))
        elif "晚上" in block:
            for p in range(11, 16):
                out.add((d, p))

    return out


# =========================
# 学期日期 / 节次时间
# =========================
SEMESTER_WEEK1_MONDAY = dt.date(2026, 8, 3)
PERIOD_STARTS = {
    1: "08:00", 2: "08:50", 3: "09:40", 4: "10:30",
    5: "11:20", 6: "12:10", 7: "14:00", 8: "14:50",
    9: "15:40", 10: "16:30", 11: "17:20", 12: "18:10",
    13: "19:00", 14: "19:50", 15: "20:40", 16: "21:30",
}
DAY_NAME_FULL = {1: "星期一", 2: "星期二", 3: "星期三", 4: "星期四", 5: "星期五", 6: "星期六", 7: "星期日"}


def period_time_range(p1: int, p2: int) -> str:
    """按每节40分钟、节间10分钟，并按用户给定的第1/7/12节时间表计算。"""
    if p1 not in PERIOD_STARTS or p2 not in PERIOD_STARTS:
        return f"第{p1}-{p2}节"
    start = dt.datetime.strptime(PERIOD_STARTS[p1], "%H:%M")
    end = dt.datetime.strptime(PERIOD_STARTS[p2], "%H:%M") + dt.timedelta(minutes=40)
    return f"{start:%H:%M}-{end:%H:%M}"


def week_day_date(week: int, day: int) -> dt.date:
    return SEMESTER_WEEK1_MONDAY + dt.timedelta(weeks=week - 1, days=day - 1)


def slot_to_text(week: int, day: int, p1: int, p2: int) -> str:
    d = week_day_date(week, day)
    return f"第{week}周 {d:%Y/%m/%d} {DAY_NAME_FULL[day]} 第{p1}-{p2}节 {period_time_range(p1, p2)}"


def format_option_schedule(opt: "CourseOption") -> str:
    """将合并后的课程班级时间展开成具体周次、日期和时段。"""
    if not opt.slots:
        return "网课/待确认（无固定时间）"
    by_week_day = defaultdict(list)
    for w, d, p in sorted(opt.slots):
        by_week_day[(w, d)].append(p)
    parts = []
    for (w, d), periods in sorted(by_week_day.items()):
        periods = sorted(set(periods))
        groups = []
        start = prev = periods[0]
        for p in periods[1:]:
            if p == prev + 1:
                prev = p
            else:
                groups.append((start, prev))
                start = prev = p
        groups.append((start, prev))
        for p1, p2 in groups:
            parts.append(slot_to_text(w, d, p1, p2))
    return "；".join(parts)


# =========================
# 4. 课程/方案
# =========================

@dataclass
class CourseOption:
    plan_course: dict
    class_name: str
    course_code: str
    course_name: str
    credit: float
    time_text: str
    address: str
    campus: str
    slots: Set[Tuple[int, int, int]] = field(default_factory=set)

    @property
    def days(self):
        return {d for _, d, _ in self.slots}

    @property
    def is_virtual(self):
        return self.campus == "网课" or not self.slots

    @property
    def exact_schedule_text(self):
        return format_option_schedule(self)

    def as_dict(self):
        return {
            "课程类别": self.plan_course.get("课程类别", ""),
            "分组": self.plan_course.get("分组", ""),
            "分组规则": self.plan_course.get("分组规则", ""),
            "班级名称": self.class_name,
            "课程编号": self.course_code,
            "课程名称": self.course_name,
            "学分": self.credit,
            "时间": self.time_text,
            "具体日期与时间": self.exact_schedule_text,
            "上课地址": self.address,
            "校区": self.campus,
        }


class CourseEngine:
    def __init__(self, pdf_path: Path, xlsx_path: Path):
        self.pdf_path = pdf_path
        self.xlsx_path = xlsx_path

        self.plan = parse_training_plan(extract_pdf_text(pdf_path))
        self.plan_by_code = {x["课程编号"]: x for x in self.plan}

        sheets = read_xlsx_sheets(xlsx_path)
        di, dh = find_header(
            sheets["课程目录"], ["课程编号", "课程名称", "学分", "校区"]
        )
        si, sh = find_header(
            sheets["课程安排表"],
            ["班级名称", "课程编号", "课程名称", "学分", "时间", "上课地址", "校区"],
        )

        self.directory = []
        for row in sheets["课程目录"][di + 1:]:
            row = [str(x).strip() for x in row]
            if len(row) < len(dh):
                row += [""] * (len(dh) - len(row))
            if row and row[0]:
                self.directory.append(dict(zip(dh, row[:len(dh)])))

        self.schedule = []
        for row in sheets["课程安排表"][si + 1:]:
            row = [str(x).strip() for x in row]
            if len(row) < len(sh):
                row += [""] * (len(sh) - len(row))
            if len(row) >= len(sh) and row[2]:
                self.schedule.append(dict(zip(sh, row[:len(sh)])))

        self.sched_by_code = defaultdict(list)
        for r in self.schedule:
            self.sched_by_code[r["课程编号"]].append(r)

        self.options_by_code = {
            code: self.build_options(code)
            for code in self.plan_by_code
        }

    def build_options(self, code: str) -> List[CourseOption]:
        rows = self.sched_by_code.get(code, [])

        # 如果课表中没有该课程的任何安排，自动设为网课（时间不冲突）
        if not rows:
            plan = self.plan_by_code.get(code)
            if plan:
                return [
                    CourseOption(
                        plan_course=plan,
                        class_name="网课/待确认",
                        course_code=code,
                        course_name=plan["课程名称"],
                        credit=float(plan["学分"]),
                        time_text="网课（时间不冲突）",
                        address="线上",
                        campus="网课",
                        slots=set(),
                    )
                ]
            return []

        by_class = defaultdict(list)

        for r in rows:
            by_class[r["班级名称"]].append(r)

        options = []
        for cls, rs in by_class.items():
            slots = set()
            for r in rs:
                slots |= parse_time(r["时间"])

            options.append(
                CourseOption(
                    plan_course=self.plan_by_code[code],
                    class_name=cls,
                    course_code=code,
                    course_name=rs[0]["课程名称"],
                    credit=float(rs[0]["学分"]),
                    time_text=" | ".join(r["时间"] for r in rs),
                    address="; ".join(dict.fromkeys(r["上课地址"] for r in rs)),
                    campus=rs[0]["校区"],
                    slots=slots,
                )
            )

        return options

    def courses_with_schedule(self):
        return {
            code: opts
            for code, opts in self.options_by_code.items()
            if opts and opts[0].class_name != "网课/待确认"
        }

    def courses_without_schedule(self):
        return [
            code for code, opts in self.options_by_code.items()
            if opts and opts[0].class_name == "网课/待确认"
        ]

    def group_courses(self, group_id: int):
        return [
            c for c in self.plan
            if c.get("分组") == group_id and c["课程编号"] in self.options_by_code
        ]


# =========================
# 5. 约束与评分
# =========================

@dataclass
class UserPreference:
    avoid_slots: Set[Tuple[int, int]] = field(default_factory=set)
    preferred_campus: str = ""
    max_other_campus_days: int = 0
    exempt_codes: Set[str] = field(default_factory=set)
    unwanted_codes: Set[str] = field(default_factory=set)
    # 课程级偏好：默认 50。100=必选，<50=负偏好，50=中性。
    preferred_scores: Dict[str, int] = field(default_factory=dict)
    # 班级级偏好：(课程编号, 班级名称) -> 0~100，默认 50。
    class_scores: Dict[Tuple[str, str], int] = field(default_factory=dict)


@dataclass
class GroupRule:
    """一个分组的需求规则。count 和 credit 可以二选一或同时存在。"""
    group_id: int
    min_count: Optional[int] = None
    max_count: Optional[int] = None
    min_credit: Optional[float] = None
    max_credit: Optional[float] = None
    description: str = ""

    def adjusted_for_exempt(self, exempt_count: int, exempt_credit: float) -> "GroupRule":
        min_count = self.min_count
        max_count = self.max_count
        min_credit = self.min_credit
        max_credit = self.max_credit

        if min_count is not None:
            min_count = max(0, min_count - exempt_count)
        if max_count is not None:
            max_count = max(0, max_count - exempt_count)
        if min_credit is not None:
            min_credit = max(0.0, min_credit - exempt_credit)
        if max_credit is not None:
            max_credit = max(0.0, max_credit - exempt_credit)

        return GroupRule(
            group_id=self.group_id,
            min_count=min_count,
            max_count=max_count,
            min_credit=min_credit,
            max_credit=max_credit,
            description=self.description,
        )


def effective_preference_score(opt: CourseOption, pref: UserPreference) -> int:
    """
    计算某个“课程-班级”的最终偏好分。

    课程默认 50，班级默认 50。
    班级分相对于 50 的偏移叠加到课程分上，并限制在 0~100。

    例如：
      课程80 + 班级50 -> 80
      课程80 + 班级30 -> 60
      课程50 + 班级100 -> 100
    """
    course_score = int(pref.preferred_scores.get(opt.course_code, 50))
    class_score = int(pref.class_scores.get((opt.course_code, opt.class_name), 50))
    return max(0, min(100, course_score + class_score - 50))


def preference_score(opt: CourseOption, pref: UserPreference) -> float:
    """将最终偏好分映射到 -1~1：50 为中性，低于50为负偏好。"""
    return (effective_preference_score(opt, pref) - 50) / 50.0


def has_avoid_time(opt: CourseOption, pref: UserPreference):
    return any((d, p) in pref.avoid_slots for _, d, p in opt.slots)


def day_campus_consistency(selected: List[CourseOption]) -> float:
    """同一天存在多个校区，视为惩罚。"""
    day_campus = defaultdict(set)
    for opt in selected:
        for d in opt.days:
            day_campus[d].add(opt.campus)

    penalty = 0
    for campuses in day_campus.values():
        if len(campuses) > 1:
            penalty += len(campuses) - 1
    return penalty


def other_campus_days(selected: List[CourseOption], preferred_campus: str) -> int:
    if not preferred_campus:
        return len(
            {
                d
                for opt in selected
                for d in opt.days
                if opt.campus != preferred_campus
            }
        )

    other = set()
    for opt in selected:
        if opt.campus != preferred_campus:
            other |= opt.days
    return len(other)


def classes_conflict(a: CourseOption, b: CourseOption):
    return bool(a.slots & b.slots)


def special_requirements(selected: List[CourseOption], pref: UserPreference):
    selected_codes = {x.course_code for x in selected}

    required_course_ok = all(
        c in selected_codes for c, score in pref.preferred_scores.items() if score >= 100
    )
    required_class_ok = all(
        any(
            x.course_code == code and x.class_name == class_name
            for x in selected
        )
        for (code, class_name), score in pref.class_scores.items()
        if score >= 100
    )
    required_ok = required_course_ok and required_class_ok
    exempt_ok = selected_codes.isdisjoint(pref.exempt_codes)
    unwanted_ok = selected_codes.isdisjoint(pref.unwanted_codes)
    avoid_ok = not any(has_avoid_time(x, pref) for x in selected)
    campus_days = other_campus_days(selected, pref.preferred_campus)
    campus_ok = campus_days <= pref.max_other_campus_days if pref.preferred_campus else True

    return {
        "100%倾向课程均已选择": required_ok,
        "免修课程均未选择": exempt_ok,
        "不想选课程均未选择": unwanted_ok,
        "避开时间均已避开": avoid_ok,
        "其他校区天数限制满足": campus_ok,
        "其他校区天数": campus_days,
    }


# =========================
# 6. 方案生成
# =========================

def physical_signature(selected: List[CourseOption]) -> Tuple[Tuple[str, str], ...]:
    """网课不参与方案身份判定；方案身份只由有真实时间/地点的课程班级组成。"""
    return tuple(sorted((o.course_code, o.class_name) for o in selected if not o.is_virtual))


def virtual_course_summary(selected: List[CourseOption], engine: CourseEngine) -> List[str]:
    """把网课从“不同方案”转换成分组级的选课要求摘要。"""
    virtual = [o for o in selected if o.is_virtual]
    if not virtual:
        return []
    groups = defaultdict(list)
    for o in virtual:
        gid = o.plan_course.get("分组")
        groups[gid].append(o)
    out = []
    for gid, opts in sorted(groups.items(), key=lambda z: (999 if z[0] is None else z[0])):
        if gid is None:
            names = "、".join(o.course_name for o in opts)
            out.append(f"网课：{names}")
            continue
        rules = [c.get("分组规则", "") for c in engine.plan if c.get("分组") == gid]
        rule = rules[0] if rules else ""
        names = "、".join(o.course_name for o in opts)
        out.append(f"网课第{gid}组：{rule or '已满足分组要求'}（代表选择：{names}）")
    return out


class Planner:
    """
    选课优化器。

    首选：OR-Tools CP-SAT
      - 每个“课程-班级”是一个 0/1 决策变量
      - 同一课程最多选一个班
      - 时间冲突通过“时间槽容量 <= 1”一次性建模
      - 培养方案分组通过学分/门数约束建模
      - 100% 倾向 = 硬约束
      - 其它倾向 = 目标函数
      - 校区天数、同日校区切换均作为可优化指标
      - 多次求解 + no-good / diversity 约束生成多个方案

    如果没有安装 OR-Tools，则退回内置 DFS + 剪枝，功能保持可用，
    但大规模数据下速度会明显弱于 CP-SAT。
    """

    def __init__(self, engine: CourseEngine, pref: UserPreference):
        self.engine = engine
        self.pref = pref
        self.options: List[CourseOption] = []
        self.option_index: Dict[Tuple[str, str], int] = {}
        self._build_option_index()

    def _build_option_index(self):
        self.options = []
        self.option_index = {}
        for code in self.engine.plan_by_code:
            for opt in self.engine.options_by_code.get(code, []):
                idx = len(self.options)
                self.options.append(opt)
                self.option_index[(opt.course_code, opt.class_name)] = idx

    def mandatory_codes(self):
        return [
            c["课程编号"]
            for c in self.engine.plan
            if c.get("课程类别") in {"A", "B"}
            and c.get("分组") is None
        ]

    def _is_required_course(self, code: str) -> bool:
        """判断课程是否必须进入最终方案。

        无分组的课程只有在“培养方案必修”或用户设置为100%必选时才允许进入方案；
        有分组的课程则由分组学分/门数规则决定。
        """
        if code in self.mandatory_codes():
            return True
        if self.pref.preferred_scores.get(code, 50) >= 100:
            return True
        return any(
            c == code and score >= 100
            for (c, _), score in self.pref.class_scores.items()
        )

    def _group_required_codes(self, group_id: int):
        """返回该分组中真正的硬性必选课程。"""
        out = []
        for c in self.engine.plan:
            if c.get("分组") != group_id:
                continue
            code = c["课程编号"]
            if code in self.pref.exempt_codes or code in self.pref.unwanted_codes:
                continue
            if self._is_required_course(code):
                out.append(code)
        return sorted(set(out))

    def _group_required_credit100(self, group_id: int) -> int:
        return sum(
            int(round(float(self.engine.plan_by_code[c].get("学分", 0) or 0) * 100))
            for c in self._group_required_codes(group_id)
        )

    def _group_required_count(self, group_id: int) -> int:
        return len(self._group_required_codes(group_id))

    def _group_residual_rule(self, group_id: int, rule: GroupRule):
        """
        将分组总规则扣除硬性必选课程后的“剩余选择规则”计算出来。

        核心原则：
        - residual_min = max(0, 总最低学分 - 必选课程学分)
        - residual_max = max(0, 总最大学分 - 必选课程学分)
        - 对门数同理。

        之后的可选课程必须形成一个“最小充分集合”：达到 residual_min
        后立即停止，不允许再加入任何额外课程。
        """
        fixed_credit = self._group_required_credit100(group_id)
        fixed_count = self._group_required_count(group_id)
        min_credit = int(round(rule.min_credit * 100)) if rule.min_credit is not None else None
        max_credit = int(round(rule.max_credit * 100)) if rule.max_credit is not None else None
        min_count = int(rule.min_count) if rule.min_count is not None else None
        max_count = int(rule.max_count) if rule.max_count is not None else None

        residual_min_credit = None if min_credit is None else max(0, min_credit - fixed_credit)
        residual_max_credit = None if max_credit is None else max(0, max_credit - fixed_credit)
        residual_min_count = None if min_count is None else max(0, min_count - fixed_count)
        residual_max_count = None if max_count is None else max(0, max_count - fixed_count)

        return {
            "fixed_credit": fixed_credit,
            "fixed_count": fixed_count,
            "min_credit": residual_min_credit,
            "max_credit": residual_max_credit,
            "min_count": residual_min_count,
            "max_count": residual_max_count,
        }

    def _is_group_required_code(self, code: str) -> bool:
        c = self.engine.plan_by_code.get(code)
        if not c or c.get("分组") is None:
            return False
        return code in self._group_required_codes(c.get("分组"))

    def _parse_group_rule(self, group_id: int, description: str) -> GroupRule:
        """把当前培养方案中的规则文本转成通用约束。"""
        text = description or ""

        # “3选1”“45选1” -> 选 1 门
        m = re.search(r"(\d+)\s*选\s*(\d+)", text)
        if m:
            count = int(m.group(2))
            return GroupRule(
                group_id=group_id,
                min_count=count,
                max_count=count,
                description=text,
            )

        # “选2-3学分”
        m = re.search(r"选\s*(\d+(?:\.\d+)?)\s*[-~到]\s*(\d+(?:\.\d+)?)\s*学分", text)
        if m:
            lo, hi = map(float, m.groups())
            return GroupRule(group_id=group_id, min_credit=lo, max_credit=hi, description=text)

        # “选2学分”“至少2学分”
        m = re.search(r"至少\s*(\d+(?:\.\d+)?)\s*学分", text)
        if m:
            lo = float(m.group(1))
            exact = re.search(r"选\s*(\d+(?:\.\d+)?)\s*学分", text)
            hi = float(exact.group(1)) if exact else None
            return GroupRule(group_id=group_id, min_credit=lo, max_credit=hi, description=text)

        m = re.search(r"选\s*(\d+(?:\.\d+)?)\s*学分", text)
        if m:
            credit = float(m.group(1))
            return GroupRule(group_id=group_id, min_credit=credit, max_credit=credit, description=text)

        # 未识别的规则：不擅自猜测。
        return GroupRule(group_id=group_id, description=text)

    def group_rules(self) -> Dict[int, GroupRule]:
        raw = {}
        for c in self.engine.plan:
            gid = c.get("分组")
            if gid is not None and gid not in raw:
                raw[gid] = self._parse_group_rule(gid, c.get("分组规则", ""))
        out = {}
        for gid, rule in raw.items():
            group_courses = [c for c in self.engine.plan if c.get("分组") == gid]
            exempt = [c for c in group_courses if c["课程编号"] in self.pref.exempt_codes]
            exempt_count = len(exempt)
            exempt_credit = sum(float(c.get("学分", 0) or 0) for c in exempt)
            out[gid] = rule.adjusted_for_exempt(exempt_count, exempt_credit)
        return out

    def group_courses(self, group_id: int):
        return [c for c in self.engine.plan if c.get("分组") == group_id]

    def _course_option_indices(self, code: str) -> List[int]:
        return [i for i, o in enumerate(self.options) if o.course_code == code]

    def _usable_options(self, code: str) -> List[int]:
        out = []
        for i in self._course_option_indices(code):
            opt = self.options[i]
            if code in self.pref.exempt_codes or code in self.pref.unwanted_codes:
                continue
            if has_avoid_time(opt, self.pref):
                continue
            out.append(i)
        return out

    def _hard_precheck(self):
        """快速发现明显的无解原因，避免把明显无解的问题交给求解器。"""
        issues = []
        for code in self.mandatory_codes():
            if code in self.pref.exempt_codes:
                continue
            if code in self.pref.unwanted_codes:
                name = self.engine.plan_by_code.get(code, {}).get("课程名称", code)
                issues.append(f"必修课程“{name}”同时被设为不想选")
            elif not self._usable_options(code):
                name = self.engine.plan_by_code.get(code, {}).get("课程名称", code)
                issues.append(f"必修课程“{name}”没有可用班级（可能都被避开时间过滤）")

        for code, score in self.pref.preferred_scores.items():
            if score >= 100:
                if code in self.pref.exempt_codes:
                    name = self.engine.plan_by_code.get(code, {}).get("课程名称", code)
                    issues.append(f"100%必选课程“{name}”被设置为免修")
                elif code in self.pref.unwanted_codes:
                    name = self.engine.plan_by_code.get(code, {}).get("课程名称", code)
                    issues.append(f"100%必选课程“{name}”同时被设为不想选")
                elif not self._usable_options(code):
                    name = self.engine.plan_by_code.get(code, {}).get("课程名称", code)
                    issues.append(f"100%必选课程“{name}”没有可用班级")
        return issues

    @staticmethod
    def _option_score(opt: CourseOption, pref: UserPreference) -> int:
        # 整数化，方便 CP-SAT。
        return int(round(preference_score(opt, pref) * 1000))

    def _build_cp_model(self, banned_signatures=None):
        model = cp_model.CpModel()
        x = [model.NewBoolVar(f"x_{i}") for i in range(len(self.options))]

        # 1) 同一课程最多一个班；必修/100%课程准确选一个。
        codes = sorted(self.engine.plan_by_code)
        for code in codes:
            inds = self._usable_options(code)
            exempt_or_unwanted = code in self.pref.exempt_codes or code in self.pref.unwanted_codes
            if exempt_or_unwanted:
                for i in self._course_option_indices(code):
                    model.Add(x[i] == 0)
                continue

            # 无分组的非必修课程不应被“偏好目标”额外选入。
            # 它们不承担任何培养方案学分要求，达到所需学分后无需继续加课。
            if self.engine.plan_by_code[code].get("分组") is None and not self._is_required_course(code):
                for i in self._course_option_indices(code):
                    model.Add(x[i] == 0)
                continue

            # 100%偏好：硬约束；固定必修：硬约束；普通分组课程：由分组规则决定。
            is_mandatory = code in self.mandatory_codes()
            is_required_pref = self.pref.preferred_scores.get(code, 50) >= 100
            required_class_indices = [
                i for i in inds
                if self.pref.class_scores.get((self.options[i].course_code, self.options[i].class_name), 50) >= 100
            ]
            if len(required_class_indices) > 1:
                raise ValueError(
                    f"课程 {self.engine.plan_by_code.get(code, {}).get('课程名称', code)} 同时有多个100%必选班级。"
                )
            if is_mandatory or is_required_pref:
                model.Add(sum(x[i] for i in inds) == 1)
            elif required_class_indices:
                # 指定某个班级为100%必选：必须选该班，且同一课程只能选一个班。
                model.Add(sum(x[i] for i in required_class_indices) == 1)
                model.Add(sum(x[i] for i in inds) == 1)
            else:
                model.Add(sum(x[i] for i in self._course_option_indices(code)) <= 1)
                # 禁用避开时间的班级。
                for i in self._course_option_indices(code):
                    if i not in inds:
                        model.Add(x[i] == 0)

        # 2) 时间冲突：对每个真实时间槽，最多一个班。
        slot_to_options = defaultdict(list)
        for i, opt in enumerate(self.options):
            if opt.course_code in self.pref.exempt_codes or opt.course_code in self.pref.unwanted_codes:
                continue
            if has_avoid_time(opt, self.pref):
                continue
            for slot in opt.slots:
                slot_to_options[slot].append(i)
        for inds in slot_to_options.values():
            if len(inds) > 1:
                model.Add(sum(x[i] for i in inds) <= 1)

        # 3) 分组规则：先扣除硬性必选课程，只对“剩余可选课程”施加规则。
        #    对学分型分组，额外加入“最小充分集合”约束：
        #    若选中了某门可选课，则删除它后必须使剩余学分低于 residual_min。
        for gid, rule in self.group_rules().items():
            residual = self._group_residual_rule(gid, rule)
            required_codes = set(self._group_required_codes(gid))
            inds = [
                i for i, o in enumerate(self.options)
                if o.plan_course.get("分组") == gid
                and o.course_code not in required_codes
                and o.course_code not in self.pref.exempt_codes
                and o.course_code not in self.pref.unwanted_codes
                and not has_avoid_time(o, self.pref)
            ]

            optional_count_expr = sum(x[i] for i in inds)
            optional_credit_expr = sum(
                int(round(self.options[i].credit * 100)) * x[i] for i in inds
            )

            if residual["max_count"] is not None:
                if residual["max_count"] < 0:
                    model.Add(0 == 1)
                else:
                    model.Add(optional_count_expr <= residual["max_count"])
            if residual["min_count"] is not None:
                if residual["min_count"] < 0:
                    model.Add(0 == 1)
                else:
                    model.Add(optional_count_expr >= residual["min_count"])

            if residual["max_credit"] is not None:
                if residual["max_credit"] < 0:
                    model.Add(0 == 1)
                else:
                    model.Add(optional_credit_expr <= residual["max_credit"])
            if residual["min_credit"] is not None:
                if residual["min_credit"] < 0:
                    model.Add(0 == 1)
                else:
                    model.Add(optional_credit_expr >= residual["min_credit"])

                # 最小充分集合约束。
                # 对任一被选课程 i：删除 i 后，剩余可选学分必须不足 residual_min。
                # 这会自动排除：residual_min=1 时同时选择两门1学分课程。
                for i in inds:
                    credit_i = int(round(self.options[i].credit * 100))
                    model.Add(
                        optional_credit_expr - credit_i <= residual["min_credit"] - 1
                    ).OnlyEnforceIf(x[i])

            # 若是门数型规则，同样要求达到最低门数后立即停止。
            if residual["min_count"] is not None and residual["min_credit"] is None:
                for i in inds:
                    model.Add(optional_count_expr <= residual["min_count"]).OnlyEnforceIf(x[i])

        # 4) 其他校区天数硬约束。
        other_day_vars = {}
        extra_campus_active = []
        if self.pref.preferred_campus:
            for d in range(1, 8):
                y = model.NewBoolVar(f"other_day_{d}")
                other_day_vars[d] = y
                day_inds = [
                    i for i, o in enumerate(self.options)
                    if o.campus != self.pref.preferred_campus and d in o.days
                ]
                for i in day_inds:
                    model.Add(y >= x[i])
                if day_inds:
                    model.Add(sum(x[i] for i in day_inds) >= y)
                else:
                    model.Add(y == 0)
                extra_campus_active.append(y)
            model.Add(sum(extra_campus_active) <= int(self.pref.max_other_campus_days))

        # 5) 同日多校区：active(d,campus)，惩罚额外校区数。
        campus_active = []
        day_used = []
        campus_names = sorted({o.campus for o in self.options})
        for d in range(1, 8):
            used = model.NewBoolVar(f"day_used_{d}")
            day_used.append(used)
            day_all_inds = [i for i, o in enumerate(self.options) if d in o.days]
            if day_all_inds:
                for i in day_all_inds:
                    model.Add(used >= x[i])
                model.Add(sum(x[i] for i in day_all_inds) >= used)
            else:
                model.Add(used == 0)

            for campus in campus_names:
                active = model.NewBoolVar(f"campus_{d}_{len(campus_active)}")
                inds = [i for i, o in enumerate(self.options) if o.campus == campus and d in o.days]
                if inds:
                    for i in inds:
                        model.Add(active >= x[i])
                    model.Add(sum(x[i] for i in inds) >= active)
                else:
                    model.Add(active == 0)
                campus_active.append(active)

        # 6) 排除以前得到的“实体选课方案”。网课不参与方案身份。
        banned_signatures = banned_signatures or []
        physical_indices = [i for i, o in enumerate(self.options) if not o.is_virtual]
        physical_set = set(physical_indices)
        for selected_indices in banned_signatures:
            sig = set(selected_indices) & physical_set
            terms = [x[i] for i in sig] + [1 - x[i] for i in physical_indices if i not in sig]
            if terms:
                model.Add(sum(terms) >= 1)

        # 7) 目标函数：先最小化无必要学分，再最少课程数，再考虑用户偏好与校区。
        preference_expr = sum(self._option_score(o, self.pref) * x[i] for i, o in enumerate(self.options))
        switch_penalty_expr = sum(campus_active) - sum(day_used)
        other_days_expr = sum(extra_campus_active) if extra_campus_active else 0
        course_count_expr = sum(x)
        total_credit_expr = sum(int(round(o.credit * 100)) * x[i] for i, o in enumerate(self.options))
        objective = (
            -total_credit_expr * 10_000_000
            -course_count_expr * 10_000
            + preference_expr * 100
            - switch_penalty_expr * 10
            - other_days_expr
        )
        model.Maximize(objective)
        return model, x

    def _virtual_group_states(self):
        """
        仅为“非必选网课”计算所有可达组合状态。

        必选网课不再混入 DP 状态，而是在完成阶段作为固定学分先扣除。
        每个状态记录：(门数, 学分*100) -> (偏好分, 课程索引列表)。
        """
        states_by_group = {}
        for gid, rule in self.group_rules().items():
            required_codes = set(self._group_required_codes(gid))
            virtual_codes = []
            for c in self.group_courses(gid):
                code = c["课程编号"]
                if code in required_codes:
                    continue
                opts = self._usable_options(code)
                if opts and all(self.options[i].is_virtual for i in opts):
                    virtual_codes.append(code)

            states = {(0, 0): (0.0, [])}
            for code in sorted(virtual_codes):
                inds = self._usable_options(code)
                if not inds:
                    continue
                i = inds[0]
                opt = self.options[i]
                credit100 = int(round(opt.credit * 100))
                score = preference_score(opt, self.pref)
                next_states = dict(states)
                for (cnt, cr), (old_score, chosen) in states.items():
                    key = (cnt + 1, cr + credit100)
                    cand = (old_score + score, chosen + [i])
                    if key not in next_states or cand[0] > next_states[key][0]:
                        next_states[key] = cand
                states = next_states
            states_by_group[gid] = states
        return states_by_group

    def _best_virtual_state(self, gid: int, rule: GroupRule, fixed_credit100: int, fixed_count: int):
        """在扣除必选课程后，选择最小充分的网课补充组合。"""
        residual = self._group_residual_rule(gid, rule)
        states = self._virtual_states_cache.get(gid, {(0, 0): (0.0, [])})
        candidates = []
        for (v_count, v_credit), (v_score, v_indices) in states.items():
            if residual["min_count"] is not None and v_count < residual["min_count"]:
                continue
            if residual["max_count"] is not None and v_count > residual["max_count"]:
                continue
            if residual["min_credit"] is not None and v_credit < residual["min_credit"]:
                continue
            if residual["max_credit"] is not None and v_credit > residual["max_credit"]:
                continue

            # 对学分型规则：优先最少满足学分，再最少门数，再偏好。
            # 对门数型规则：优先最少门数，再学分，再偏好。
            if residual["min_credit"] is not None:
                rank = (v_credit, v_count, -v_score, tuple(v_indices))
            else:
                rank = (v_count, v_credit, -v_score, tuple(v_indices))
            candidates.append((rank, v_indices, v_score))

        if not candidates:
            return None
        candidates.sort(key=lambda z: z[0])
        _, indices, score = candidates[0]
        return indices, score

    @staticmethod
    def _rule_accepts(rule: GroupRule, count: int, credit100: int) -> bool:
        credit = credit100 / 100.0
        if rule.min_count is not None and count < rule.min_count:
            return False
        if rule.max_count is not None and count > rule.max_count:
            return False
        if rule.min_credit is not None and credit + 1e-9 < rule.min_credit:
            return False
        if rule.max_credit is not None and credit - 1e-9 > rule.max_credit:
            return False
        return True

    def _complete_physical_selection(self, selected):
        """给定实体课程组合，补入固定必选网课及“最小充分”网课组合。"""
        total_score = sum(preference_score(o, self.pref) for o in selected)
        completed = list(selected)
        virtual_states = self._virtual_states_cache

        # 先固定加入所有必选网课。
        selected_codes0 = {o.course_code for o in completed}
        required_codes_all = set(self.mandatory_codes()) | {
            c for c, s in self.pref.preferred_scores.items() if s >= 100
        } | {
            c for (c, _), s in self.pref.class_scores.items() if s >= 100
        }
        required_codes_all -= self.pref.exempt_codes
        required_codes_all -= self.pref.unwanted_codes

        for code in sorted(required_codes_all):
            if code in selected_codes0:
                continue
            opts = self._usable_options(code)
            virtual_opts = [i for i in opts if self.options[i].is_virtual]
            if virtual_opts and not any(not self.options[i].is_virtual for i in opts):
                # 同一课程只有一个网课占位时直接加入；若存在多个网课班，取偏好最高者。
                best_i = max(virtual_opts, key=lambda i: self._option_score(self.options[i], self.pref))
                completed.append(self.options[best_i])
                total_score += preference_score(self.options[best_i], self.pref)

        for gid, rule in self.group_rules().items():
            group_all = [o for o in completed if o.plan_course.get("分组") == gid]
            fixed_required_codes = set(self._group_required_codes(gid))
            fixed_required = [o for o in group_all if o.course_code in fixed_required_codes]
            optional_physical = [o for o in group_all if o.course_code not in fixed_required_codes]

            fixed_credit = self._group_required_credit100(gid)
            fixed_count = self._group_required_count(gid)
            # 防止重复/异常情况下漏计真正的实体必选课程；统一以培养方案定义为准。
            if fixed_required:
                fixed_credit = sum(int(round(o.credit * 100)) for o in fixed_required) + sum(
                    int(round(self.engine.plan_by_code[c]["学分"] * 100))
                    for c in fixed_required_codes - {o.course_code for o in fixed_required}
                )

            residual = self._group_residual_rule(gid, rule)
            p_optional_credit = sum(int(round(o.credit * 100)) for o in optional_physical)
            p_optional_count = len(optional_physical)

            # 实体部分已经达到最低要求时，绝不再补任何网课。
            if residual["min_credit"] is not None:
                if p_optional_credit > residual["max_credit"] if residual["max_credit"] is not None else False:
                    return None
            if residual["min_count"] is not None and p_optional_count > residual["max_count"] if residual["max_count"] is not None else False:
                return None

            need_credit = None if residual["min_credit"] is None else max(0, residual["min_credit"] - p_optional_credit)
            need_count = None if residual["min_count"] is None else max(0, residual["min_count"] - p_optional_count)
            max_credit_left = None if residual["max_credit"] is None else residual["max_credit"] - p_optional_credit
            max_count_left = None if residual["max_count"] is None else residual["max_count"] - p_optional_count

            # 如果实体课程已经满足最低要求，则网课必须是空集。
            if (need_credit is not None and need_credit == 0) or (need_count is not None and need_count == 0 and residual["min_credit"] is None):
                if residual["min_credit"] is not None and p_optional_credit < residual["min_credit"]:
                    return None
                continue

            # 为剩余网课重新筛选“最小充分”状态。
            states = virtual_states.get(gid, {(0, 0): (0.0, [])})
            candidates = []
            for (vcnt, vcredit), (vscore, v_indices) in states.items():
                total_optional_credit = p_optional_credit + vcredit
                total_optional_count = p_optional_count + vcnt
                if residual["min_credit"] is not None and total_optional_credit < residual["min_credit"]:
                    continue
                if residual["max_credit"] is not None and total_optional_credit > residual["max_credit"]:
                    continue
                if residual["min_count"] is not None and total_optional_count < residual["min_count"]:
                    continue
                if residual["max_count"] is not None and total_optional_count > residual["max_count"]:
                    continue

                # 最小充分集合：对于最终加入的每个网课，去掉它后必须低于最低要求。
                minimal = True
                for idx in v_indices:
                    ci = int(round(self.options[idx].credit * 100))
                    if residual["min_credit"] is not None and total_optional_credit - ci >= residual["min_credit"]:
                        minimal = False
                        break
                    if residual["min_credit"] is None and residual["min_count"] is not None and total_optional_count - 1 >= residual["min_count"]:
                        minimal = False
                        break
                if not minimal:
                    continue

                if residual["min_credit"] is not None:
                    rank = (total_optional_credit, total_optional_count, -vscore, tuple(v_indices))
                else:
                    rank = (total_optional_count, total_optional_credit, -vscore, tuple(v_indices))
                candidates.append((rank, v_indices, vscore))

            if not candidates:
                return None
            candidates.sort(key=lambda z: z[0])
            _, v_indices, v_score = candidates[0]
            completed.extend(self.options[i] for i in v_indices)
            total_score += v_score

        if not self._all_required_selected({o.course_code for o in completed}):
            return None
        return completed, total_score

    def _generate_exhaustive_exact(self):
        """完整、确定性的实体方案枚举。

        与简单笛卡尔积相比，这里使用：
        1. 冲突位图：O(1) 判断班级是否与已选班级冲突；
        2. 动态 MRV：优先展开当前可行选项最少的课程；
        3. 分组上下界剪枝：剩余课程即使全部选入也无法满足时立即回退；
        4. 校区天数剪枝；
        5. 失败状态缓存（memoization）：相同“剩余课程+占用时间+分组状态+校区状态”不重复搜索；
        6. 网课只通过 DP 补足分组，不进入实体方案身份。

        这些都是安全剪枝：不会删除任何实际上可以形成合法实体方案的分支。
        因而“尽可能全部”模式仍然保持完整性，但通常比朴素全枚举快很多。
        """
        pre = self._hard_precheck()
        if pre:
            raise ValueError("\n".join("- " + x for x in pre))

        self._virtual_states_cache = self._virtual_group_states()
        group_rules = self.group_rules()
        group_ids = sorted(group_rules)
        group_pos = {gid: j for j, gid in enumerate(group_ids)}

        mandatory = set(self.mandatory_codes())
        required_courses = {
            c for c, score in self.pref.preferred_scores.items() if score >= 100
        }
        required_courses |= {
            code for (code, _), score in self.pref.class_scores.items() if score >= 100
        }
        required_courses -= self.pref.exempt_codes
        required_courses -= self.pref.unwanted_codes

        # 只考虑存在真实班级的课程。纯网课课程由 DP 处理。
        physical_codes = []
        raw_domains = {}
        for code in sorted(self.engine.plan_by_code):
            if code in self.pref.exempt_codes or code in self.pref.unwanted_codes:
                continue

            # 无分组且非必修/非100%必选的课程不进入搜索空间，避免在学分已满足后继续加课。
            if self.engine.plan_by_code[code].get("分组") is None and not self._is_required_course(code):
                raw_domains[code] = []
                continue

            inds = [i for i in self._usable_options(code) if not self.options[i].is_virtual]
            # 对100%必选班级做硬过滤。
            required_class_names = {
                cls for (c, cls), score in self.pref.class_scores.items()
                if c == code and score >= 100
            }
            if required_class_names:
                inds = [i for i in inds if self.options[i].class_name in required_class_names]
            raw_domains[code] = inds
            if inds:
                physical_codes.append(code)
            elif code in mandatory or code in required_courses:
                # 没有实体班级：如果有网课会由补充阶段处理，否则最终完整性检查会判无解。
                continue

        # 预计算：课程 -> 组，以及每组课程的学分/门数上界信息。
        code_group = {
            c["课程编号"]: c.get("分组") for c in self.engine.plan
            if c.get("课程编号") in raw_domains
        }
        code_credit100 = {
            c["课程编号"]: int(round(float(c.get("学分", 0) or 0) * 100))
            for c in self.engine.plan
            if c.get("课程编号") in raw_domains
        }

        # 冲突位图：一个 Python int 表示一组班级冲突关系。
        option_count = len(self.options)
        conflict_masks = [0] * option_count
        slot_map = defaultdict(list)
        for i, opt in enumerate(self.options):
            if opt.is_virtual:
                continue
            if opt.course_code in self.pref.exempt_codes or opt.course_code in self.pref.unwanted_codes:
                continue
            if has_avoid_time(opt, self.pref):
                continue
            for slot in opt.slots:
                slot_map[slot].append(i)
        for inds in slot_map.values():
            if len(inds) > 1:
                mask = 0
                for i in inds:
                    mask |= 1 << i
                for i in inds:
                    conflict_masks[i] |= mask ^ (1 << i)

        # 每个班级的“其他校区天”位图。
        option_other_day_mask = [0] * option_count
        preferred = self.pref.preferred_campus
        for i, opt in enumerate(self.options):
            if preferred and not opt.is_virtual and opt.campus != preferred:
                mask = 0
                for d in opt.days:
                    mask |= 1 << (d - 1)
                option_other_day_mask[i] = mask

        # 每组已选状态：(count, credit100)。
        # 先把“纯网课的硬性必选课程”视为固定学分/门数，避免实体可选课把它忽略。
        initial_group_state = []
        for gid in group_ids:
            initial_group_state.append((
                self._group_required_count(gid) if all(
                    self.options[i].is_virtual
                    for c in self._group_required_codes(gid)
                    for i in self._usable_options(c)
                ) and self._group_required_codes(gid) else 0,
                self._group_required_credit100(gid) if all(
                    self.options[i].is_virtual
                    for c in self._group_required_codes(gid)
                    for i in self._usable_options(c)
                ) and self._group_required_codes(gid) else 0,
            ))
        zero_group_state = tuple(initial_group_state)

        # 对每个课程，缓存“按当前占用时间/组上限可用的实体班级”的过滤函数结果。
        # 失败状态缓存只存不可能产生解的状态，不会影响完整性。
        dead_states = set()
        results = []
        seen = set()

        def feasible_inds(code, selected_mask, group_state):
            gid = code_group.get(code)
            if gid is not None and gid in group_pos:
                gi = group_pos[gid]
                rule = group_rules[gid]
                count, credit = group_state[gi]
                required_code = self._is_group_required_code(code)

                if not required_code:
                    # 已达到最低要求：这个分组立即关闭，不再允许额外课程。
                    if rule.min_credit is not None and credit >= int(round(rule.min_credit * 100)):
                        return []
                    if rule.min_count is not None and rule.min_credit is None and count >= int(rule.min_count):
                        return []
                    # 不能超过扣除必选课程后的上限。
                    residual = self._group_residual_rule(gid, rule)
                    if residual["max_credit"] is not None and credit - self._group_required_credit100(gid) >= residual["max_credit"]:
                        return []
                    if residual["max_count"] is not None and count - self._group_required_count(gid) >= residual["max_count"]:
                        return []

            out = []
            for i in raw_domains.get(code, []):
                if conflict_masks[i] & selected_mask:
                    continue
                opt = self.options[i]
                if gid is not None and gid in group_pos:
                    new_count, new_credit = update_group_state(group_state, opt)[group_pos[gid]]
                    if rule.max_count is not None and new_count > int(rule.max_count):
                        continue
                    if rule.max_credit is not None and new_credit > int(round(rule.max_credit * 100)):
                        continue
                out.append(i)
            return out

        def update_group_state(group_state, opt):
            gid = opt.plan_course.get("分组")
            if gid is None or gid not in group_pos:
                return group_state
            gi = group_pos[gid]
            state = list(group_state)
            count, credit = state[gi]
            state[gi] = (count + 1, credit + int(round(opt.credit * 100)))
            return tuple(state)

        def group_bounds_possible(remaining_codes, group_state):
            """用乐观上界检查每个分组是否仍可能达到最低门数/学分。"""
            remaining_by_group = {gid: [] for gid in group_ids}
            for code in remaining_codes:
                gid = code_group.get(code)
                if gid in remaining_by_group:
                    remaining_by_group[gid].append(code)

            for gid in group_ids:
                rule = group_rules[gid]
                if gid not in group_pos:
                    continue
                count, credit = group_state[group_pos[gid]]
                # 先检查当前已经超上限。
                if rule.max_count is not None and count > rule.max_count:
                    return False
                if rule.max_credit is not None and credit > int(round(rule.max_credit * 100)):
                    return False

                rem = remaining_by_group[gid]

                # 实体课程的乐观上界：忽略时间冲突，假设所有剩余课程都可以选择。
                physical_count_cap = len(rem)
                if rule.max_count is not None:
                    physical_count_cap = min(physical_count_cap, max(0, rule.max_count - count))
                physical_credits = sorted((code_credit100.get(c, 0) for c in rem), reverse=True)[:physical_count_cap]
                physical_credit_cap = sum(physical_credits)

                # 网课也可能满足同一分组，因此必须纳入上界；否则会错误剪掉“实体0课 + 网课满足”的分支。
                virtual_states = self._virtual_states_cache.get(gid, {(0, 0): (0, [])})
                virtual_count_cap = 0
                virtual_credit_cap = 0
                for (vcnt, vcredit), _ in virtual_states.items():
                    if rule.max_count is not None and count + vcnt > rule.max_count:
                        continue
                    if rule.max_credit is not None and credit + vcredit > int(round(rule.max_credit * 100)):
                        continue
                    virtual_count_cap = max(virtual_count_cap, vcnt)
                    virtual_credit_cap = max(virtual_credit_cap, vcredit)

                if rule.min_count is not None:
                    need_count = rule.min_count - count
                    if need_count > 0 and physical_count_cap + virtual_count_cap < need_count:
                        return False

                if rule.min_credit is not None:
                    need_credit = int(round(rule.min_credit * 100)) - credit
                    if need_credit > 0 and physical_credit_cap + virtual_credit_cap < need_credit:
                        return False
            return True

        def mandatory_available(remaining_codes, selected_codes, selected_mask, group_state):
            # 所有还未选择的必选课程必须仍存在至少一个可行班级。
            remaining_required = (mandatory | required_courses) - selected_codes
            for code in remaining_required:
                if code in self.pref.exempt_codes or code in self.pref.unwanted_codes:
                    continue
                if code not in remaining_codes:
                    # 没有实体分支可能是网课；留给最终补充阶段判断。
                    if not any(
                        self.options[i].is_virtual and self.options[i].course_code == code
                        for i in self._usable_options(code)
                    ):
                        return False
                    continue
                if not feasible_inds(code, selected_mask, group_state):
                    return False
            return True

        def choose_next_code(remaining_codes, selected_mask, group_state):
            """MRV：选择当前可行实体班级数量最少的课程。"""
            best_code = None
            best_count = None
            best_required = False
            required_pool = [
                code for code in remaining_codes
                if code in mandatory or code in required_courses
            ]
            pool = required_pool if required_pool else list(remaining_codes)
            for code in pool:
                inds = feasible_inds(code, selected_mask, group_state)
                required = code in mandatory or code in required_courses
                if required and not inds:
                    return code, [], True
                # 可选课程永远有“不选”分支，因此至少一个分支可走。
                branch_count = len(inds) + (0 if required else 1)
                if best_count is None or branch_count < best_count or (
                    branch_count == best_count and required and not best_required
                ):
                    best_code = code
                    best_count = branch_count
                    best_required = required
            return best_code, (feasible_inds(best_code, selected_mask, group_state) if best_code else []), best_required

        def complete_and_record(selected):
            completed = self._complete_physical_selection(selected)
            if completed is None:
                return
            final_selected, total_pref = completed
            if any(has_avoid_time(o, self.pref) for o in final_selected):
                return
            if self.pref.preferred_campus:
                other_days = other_campus_days(final_selected, self.pref.preferred_campus)
                if other_days > self.pref.max_other_campus_days:
                    return
            key = physical_signature(final_selected)
            if key in seen:
                return
            seen.add(key)
            score = (
                -total_pref,
                day_campus_consistency(final_selected),
                other_campus_days(final_selected, self.pref.preferred_campus) if self.pref.preferred_campus else 0,
                len(key),
            )
            results.append((score, final_selected))

        def dfs(remaining_codes, selected, selected_codes, selected_mask, group_state, other_day_mask):
            # 安全的全局剪枝。
            if not group_bounds_possible(remaining_codes, group_state):
                return
            if not mandatory_available(remaining_codes, selected_codes, selected_mask, group_state):
                return

            if not remaining_codes:
                complete_and_record(selected)
                return

            # 失败状态缓存：只记录确定无解的状态。
            # remaining_codes 已排序，group_state 与占用/校区状态共同决定未来可行性。
            state_key = (
                tuple(sorted(remaining_codes)),
                selected_mask,
                group_state,
                other_day_mask,
                tuple(sorted(selected_codes & (mandatory | required_courses))),
            )
            if state_key in dead_states:
                return

            code, inds, is_required = choose_next_code(remaining_codes, selected_mask, group_state)
            if code is None:
                complete_and_record(selected)
                return

            next_remaining = tuple(c for c in remaining_codes if c != code)
            # 先高偏好班级，最后“不选”；仅影响搜索顺序，不影响完整性。
            inds = sorted(
                inds,
                key=lambda i: (-self._option_score(self.options[i], self.pref), self.options[i].class_name)
            )

            found_solution = False
            branches = inds if is_required else [None] + inds
            for i in branches:
                if i is None:
                    before = len(results)
                    dfs(next_remaining, selected, selected_codes, selected_mask, group_state, other_day_mask)
                    found_solution = found_solution or len(results) > before
                    continue

                opt = self.options[i]
                new_mask = selected_mask | (1 << i)
                new_day_mask = other_day_mask | option_other_day_mask[i]
                if preferred and new_day_mask.bit_count() > self.pref.max_other_campus_days:
                    continue

                new_group_state = update_group_state(group_state, opt)
                gid = opt.plan_course.get("分组")
                if gid is not None and gid in group_pos:
                    rule = group_rules[gid]
                    gi = group_pos[gid]
                    cnt, cr = new_group_state[gi]
                    if rule.max_count is not None and cnt > rule.max_count:
                        continue
                    if rule.max_credit is not None and cr > int(round(rule.max_credit * 100)):
                        continue

                before = len(results)
                dfs(
                    next_remaining,
                    selected + [opt],
                    selected_codes | {code},
                    new_mask,
                    new_group_state,
                    new_day_mask,
                )
                found_solution = found_solution or len(results) > before

            if not found_solution:
                dead_states.add(state_key)

        dfs(tuple(sorted(physical_codes)), [], set(), 0, zero_group_state, 0)
        results.sort(key=lambda z: (z[0], physical_signature(z[1])))
        return results

    def _solve_cp_sat(self, n: Optional[int]):
        """CP-SAT 用于有限数量的优选方案；“尽可能全部”改走严格全枚举。

        这样避免把“多次求一个可行解”误认为数学意义上的全枚举。
        """
        if n is None:
            return self._generate_exhaustive_exact()
        if not ORTOOLS_AVAILABLE:
            return None

        pre = self._hard_precheck()
        if pre:
            raise ValueError("\n".join("- " + x for x in pre))

        banned = []
        raw = []
        target = max(1, int(n))

        while len(raw) < target:
            model, x = self._build_cp_model(banned)
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 30.0
            solver.parameters.num_search_workers = 1
            solver.parameters.randomize_search = False
            solver.parameters.log_search_progress = False

            status = solver.Solve(model)
            if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                break

            selected_indices = [i for i, var in enumerate(x) if solver.Value(var)]
            selected = [self.options[i] for i in selected_indices]
            physical_sig = physical_signature(selected)
            if not physical_sig:
                break
            if physical_sig in {physical_signature(s) for _, s in raw}:
                banned.append([i for i, o in enumerate(self.options) if not o.is_virtual and (o.course_code, o.class_name) in set(physical_sig)])
                continue

            raw.append((selected_indices, selected))
            sig_set = set(physical_sig)
            banned.append([
                i for i, o in enumerate(self.options)
                if not o.is_virtual and (o.course_code, o.class_name) in sig_set
            ])

        scored = []
        for _, selected in raw:
            pref_score = sum(preference_score(o, self.pref) for o in selected)
            campus_penalty = day_campus_consistency(selected)
            other_days = other_campus_days(selected, self.pref.preferred_campus) if self.pref.preferred_campus else 0
            scored.append(((-pref_score, campus_penalty, other_days, len(selected)), selected))
        scored.sort(key=lambda z: (z[0], physical_signature(z[1])))
        return scored[:target]

    def _all_required_selected(self, codes):
        required = set(self.mandatory_codes()) | {
            c for c, s in self.pref.preferred_scores.items() if s >= 100
        } | {
            c for (c, _), s in self.pref.class_scores.items() if s >= 100
        }
        required -= self.pref.exempt_codes
        return required.issubset(codes)

    def _check_group_fallback(self, selected_codes, selected):
        """fallback 路径使用与主 DFS/CP-SAT 一致的“最小充分集合”规则。"""
        for gid, rule in self.group_rules().items():
            required_codes = set(self._group_required_codes(gid))
            group_selected = [o for o in selected if o.plan_course.get("分组") == gid]
            fixed_credit = self._group_required_credit100(gid)
            fixed_count = self._group_required_count(gid)
            optional = [o for o in group_selected if o.course_code not in required_codes]
            oc = sum(int(round(o.credit * 100)) for o in optional)
            on = len(optional)
            residual = self._group_residual_rule(gid, rule)
            if residual["min_credit"] is not None:
                if oc < residual["min_credit"] or (residual["max_credit"] is not None and oc > residual["max_credit"]):
                    return False
                # 最小充分：删除任意一门后必须不足最低学分。
                if any(oc - int(round(o.credit * 100)) >= residual["min_credit"] for o in optional):
                    return False
            if residual["min_count"] is not None:
                if on < residual["min_count"] or (residual["max_count"] is not None and on > residual["max_count"]):
                    return False
                if residual["min_credit"] is None and any(on - 1 >= residual["min_count"] for _ in optional):
                    return False
        return True

    def _fallback_generate(self, n=8):
        if n is None:
            return self._generate_exhaustive_exact()
        """无 OR-Tools 时的改进版 DFS：先选约束最强的课程，并在中途剪枝。"""
        pre = self._hard_precheck()
        if pre:
            raise ValueError("\n".join("- " + x for x in pre))

        candidate_codes = []
        mandatory = set(self.mandatory_codes())
        required = {c for c, s in self.pref.preferred_scores.items() if s >= 100}
        required_class_codes = {
            code for (code, _), score in self.pref.class_scores.items() if score >= 100
        }
        required |= required_class_codes
        for code in sorted(self.engine.plan_by_code):
            if code in self.pref.exempt_codes or code in self.pref.unwanted_codes:
                continue
            # 无分组且非必修/非100%必选的课程不进入搜索空间。
            if self.engine.plan_by_code[code].get("分组") is None and not self._is_required_course(code):
                continue
            if self._usable_options(code):
                candidate_codes.append(code)

        def domain_size(code):
            return len(self._usable_options(code))
        candidate_codes.sort(key=lambda c: (0 if c in mandatory or c in required else 1, domain_size(c)))

        results = []
        seen = set()
        target = max(1, int(n))

        def dfs(pos, selected, selected_codes):
            if len(results) >= target:
                return
            if pos == len(candidate_codes):
                if not self._check_group_fallback(selected_codes, selected):
                    return
                if not self._check_campus_days_fallback(selected):
                    return
                if not self._all_required_selected(selected_codes):
                    return
                score = self._fallback_score(selected)
                key = physical_signature(selected)
                if key not in seen:
                    seen.add(key)
                    results.append((score, selected[:]))
                return

            code = candidate_codes[pos]
            is_required = code in mandatory or code in required
            opts = self._usable_options(code)
            required_class_names = {
                cls for (c, cls), score in self.pref.class_scores.items()
                if c == code and score >= 100
            }
            if required_class_names:
                opts = [
                    i for i in opts
                    if self.options[i].class_name in required_class_names
                ]

            options_order = sorted(opts, key=lambda i: self._option_score(self.options[i], self.pref), reverse=True)
            if not is_required:
                options_order = options_order + [None]

            for i in options_order:
                if i is None:
                    dfs(pos + 1, selected, selected_codes)
                    continue
                opt = self.options[i]
                if any(classes_conflict(opt, x) for x in selected):
                    continue
                new_selected = selected + [opt]
                if not self._check_campus_days_fallback(new_selected):
                    continue
                dfs(pos + 1, new_selected, selected_codes | {code})

        dfs(0, [], set())
        results.sort(key=lambda z: (z[0], physical_signature(z[1])))
        return results[:target]

    def generate(self, n=8):
        started = time.perf_counter()
        if ORTOOLS_AVAILABLE:
            results = self._solve_cp_sat(n)
            if results is not None:
                return results
        return self._fallback_generate(n)


# =========================
# 7. GUI
# =========================

# 固定的校区配色（用于课表可视化，方便一眼区分校区）
CAMPUS_COLORS = {
    "江宁校区": "#4C72B0",   # 蓝
    "五台校区": "#DD8452",   # 橙
    "网课": "#95A5A6",       # 灰
}
_EXTRA_PALETTE = ["#55A868", "#C44E52", "#8172B2", "#937860", "#DA8BC3"]
_campus_color_cache: Dict[str, str] = {}


def campus_color(campus: str) -> str:
    if campus in CAMPUS_COLORS:
        return CAMPUS_COLORS[campus]
    if campus not in _campus_color_cache:
        idx = len(_campus_color_cache) % len(_EXTRA_PALETTE)
        _campus_color_cache[campus] = _EXTRA_PALETTE[idx]
    return _campus_color_cache[campus]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("研究生智能选课决策器")
        self.geometry("1450x900")
        self.minsize(1200, 760)

        self.pdf_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        self.n_plans = tk.StringVar(value="尽可能全部")

        self.avoid_var = tk.StringVar()
        self.campus_var = tk.StringVar(value="江宁校区")
        self.other_campus_var = tk.IntVar(value=0)

        self.engine: Optional[CourseEngine] = None
        self.results = []
        self.current_plan_idx = 0
        self.current_week = 1
        self.weeks_in_plan = []

        self.pref_score_vars = {}
        self.class_score_vars = {}
        self.exempt_vars = {}
        self.unwanted_vars = {}

        self.build_file_frame()
        self.build_constraints_frame()
        self.build_course_frame()
        self.build_actions()
        self.build_results()

    def build_file_frame(self):
        frm = ttk.LabelFrame(self, text="① 数据文件")
        frm.pack(fill="x", padx=10, pady=8)

        ttk.Label(frm, text="培养方案 PDF：").grid(row=0, column=0, padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.pdf_path, width=90).grid(row=0, column=1, padx=6)
        ttk.Button(frm, text="选择文件", command=self.select_pdf).grid(row=0, column=2, padx=6)

        ttk.Label(frm, text="课程目录 Excel：").grid(row=1, column=0, padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.xlsx_path, width=90).grid(row=1, column=1, padx=6)
        ttk.Button(frm, text="选择文件", command=self.select_xlsx).grid(row=1, column=2, padx=6)

        ttk.Button(
            frm, text="读取培养方案与课程目录",
            command=self.load_data
        ).grid(row=0, column=3, rowspan=2, padx=10)

    def build_constraints_frame(self):
        frm = ttk.LabelFrame(self, text="② 个人时间 / 校区需求")
        frm.pack(fill="x", padx=10, pady=8)

        ttk.Label(frm, text="避开时间段（可用分号分隔，如：周一7-10；周五下午）：").grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        ttk.Entry(frm, textvariable=self.avoid_var, width=80).grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=6
        )

        ttk.Label(frm, text="希望校区：").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Combobox(
            frm, textvariable=self.campus_var,
            values=["江宁校区", "五台校区", "随意"],
            state="readonly", width=15
        ).grid(row=1, column=1, sticky="w", padx=6)

        ttk.Label(frm, text="每周最多其他校区上课天数：").grid(
            row=1, column=2, sticky="e", padx=6
        )
        ttk.Combobox(
            frm, textvariable=self.other_campus_var,
            values=[0, 1, 2, 3], state="readonly", width=8
        ).grid(row=1, column=3, sticky="w", padx=6)

        ttk.Button(
            frm, text="校区偏好设置窗口…",
            command=self.open_campus_dialog
        ).grid(row=1, column=4, sticky="w", padx=10)

        ttk.Label(frm, text="生成方案数量（可选“尽可能全部”）：").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        ttk.Combobox(
            frm, textvariable=self.n_plans,
            values=[5, 10, 20, 50, 100, "尽可能全部"],
            state="readonly", width=10
        ).grid(row=2, column=1, sticky="w", padx=6)

    def build_course_frame(self):
        frm = ttk.LabelFrame(self, text="③ 课程偏好")
        frm.pack(fill="both", expand=False, padx=10, pady=8)

        columns = ("course", "code", "group", "credit", "exempt", "unwanted", "score")
        self.course_tree = ttk.Treeview(
            frm, columns=columns, show="headings", height=13
        )

        headings = {
            "course": "课程名称",
            "code": "课程编号",
            "group": "类别/分组",
            "credit": "学分",
            "exempt": "免修",
            "unwanted": "不想选",
            "score": "倾向评分"
        }

        widths = {
            "course": 240, "code": 120, "group": 120,
            "credit": 60, "exempt": 70, "unwanted": 80, "score": 90
        }

        for c in columns:
            self.course_tree.heading(c, text=headings[c])
            self.course_tree.column(c, width=widths[c], anchor="center")

        self.course_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frm, orient="vertical", command=self.course_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.course_tree.configure(yscrollcommand=scrollbar.set)

        self.course_tree.bind("<Double-1>", self.edit_course_preference)

        ttk.Label(
            frm,
            text="双击课程行可进入课程详情，并设置课程级及具体班级偏好。\n课程与班级评分默认均为50：100=必选；50=中性；<50=不想上（负偏好）；‘不想选’勾选仍表示硬性排除。",
            justify="left"
        ).pack(side="bottom", anchor="w", padx=8, pady=4)

    def build_actions(self):
        frm = ttk.Frame(self)
        frm.pack(fill="x", padx=10, pady=4)

        ttk.Button(
            frm, text="生成多个方案",
            command=self.generate_plans
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="查看上一方案",
            command=lambda: self.switch_plan(-1)
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="查看下一方案",
            command=lambda: self.switch_plan(1)
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="导出当前方案",
            command=self.export_current
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="导出全部方案",
            command=self.export_all
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="方案总览图",
            command=self.show_overview
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="导出总览PDF（全部分页）",
            command=self.export_overview_pdf
        ).pack(side="left", padx=5)

        ttk.Button(
            frm, text="导出总览CSV",
            command=self.export_overview_csv
        ).pack(side="left", padx=5)

        self.status = ttk.Label(frm, text="请先读取文件。")
        self.status.pack(side="right", padx=8)

    def build_results(self):
        frm = ttk.LabelFrame(self, text="④ 方案可视化与特殊需求检查")
        frm.pack(fill="both", expand=True, padx=10, pady=8)

        left = ttk.Frame(frm)
        left.pack(side="left", fill="both", expand=True)

        # 周选择器
        week_bar = ttk.Frame(left)
        week_bar.pack(fill="x", padx=6, pady=(4, 0))
        ttk.Button(week_bar, text="← 上一周", command=lambda: self.switch_week(-1)).pack(side="left")
        ttk.Label(week_bar, text="  第 ").pack(side="left")
        self.week_combo = ttk.Combobox(week_bar, values=[], width=6, state="readonly")
        self.week_combo.pack(side="left")
        self.week_combo.bind("<<ComboboxSelected>>", self.on_week_selected)
        ttk.Label(week_bar, text=" 周  ").pack(side="left")
        ttk.Button(week_bar, text="下一周 →", command=lambda: self.switch_week(1)).pack(side="left")
        self.week_info_label = ttk.Label(week_bar, text="")
        self.week_info_label.pack(side="left", padx=10)

        right = ttk.Frame(frm, width=400)
        right.pack(side="right", fill="y")

        self.fig, self.ax = plt.subplots(figsize=(11, 6.2))
        self.canvas = FigureCanvasTkAgg(self.fig, master=left)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        ttk.Label(right, text="全部方案（点击可直接查看）：").pack(
            anchor="w", padx=6, pady=(6, 2)
        )

        list_frame = ttk.Frame(right)
        list_frame.pack(fill="x", padx=6)

        self.plan_listbox = tk.Listbox(list_frame, height=8, exportselection=False)
        self.plan_listbox.pack(side="left", fill="x", expand=True)
        plan_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.plan_listbox.yview
        )
        plan_scroll.pack(side="right", fill="y")
        self.plan_listbox.configure(yscrollcommand=plan_scroll.set)
        self.plan_listbox.bind("<<ListboxSelect>>", self.on_plan_listbox_select)

        self.req_text = tk.Text(right, width=45, height=26, wrap="word")
        self.req_text.pack(fill="both", expand=True, padx=6, pady=6)

        # 用于给“特殊需求”结果标绿/标红
        self.req_text.tag_configure("ok", foreground="#2E8B57")
        self.req_text.tag_configure("bad", foreground="#C0392B")
        self.req_text.tag_configure("bold", font=("Arial", 10, "bold"))

    def select_pdf(self):
        p = filedialog.askopenfilename(
            title="选择培养方案 PDF",
            filetypes=[("PDF", "*.pdf")]
        )
        if p:
            self.pdf_path.set(p)

    def select_xlsx(self):
        p = filedialog.askopenfilename(
            title="选择课程目录 Excel",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            self.xlsx_path.set(p)

    def load_data(self):
        try:
            if not self.pdf_path.get() or not self.xlsx_path.get():
                raise ValueError("请先选择 PDF 和 Excel。")

            self.engine = CourseEngine(
                Path(self.pdf_path.get()),
                Path(self.xlsx_path.get())
            )

            self.populate_course_tree()

            missing = self.engine.courses_without_schedule()
            self.status.config(
                text=f"读取成功：{len(self.engine.plan)} 门培养方案课程；"
                     f"{len(missing)} 门暂未在课表中找到。"
            )

            messagebox.showinfo(
                "读取成功",
                f"培养方案：{len(self.engine.plan)} 门课程\n"
                f"课表中可匹配：{len(self.engine.courses_with_schedule())} 门\n"
                f"暂未匹配：{len(missing)} 门"
            )

            # 读取成功后自动弹出校区偏好设置窗口
            self.open_campus_dialog()

        except Exception as e:
            messagebox.showerror("读取失败", str(e))

    def open_campus_dialog(self):
        """
        读取培养方案后弹出的窗口：设置希望校区，
        以及每周最多允许在其他校区上课的天数（0/1/2/3）。
        与主界面②中的下拉框保持双向同步。
        """
        win = tk.Toplevel(self)
        win.title("设置校区偏好")
        win.geometry("380x220")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        campus_local = tk.StringVar(value=self.campus_var.get())
        other_local = tk.IntVar(value=self.other_campus_var.get())

        ttk.Label(
            win, text="培养方案已读取完成，请设置本学期的校区偏好：",
            wraplength=340, justify="left", font=("Arial", 10, "bold")
        ).pack(padx=16, pady=(16, 10), anchor="w")

        row1 = ttk.Frame(win)
        row1.pack(fill="x", padx=16, pady=6)
        ttk.Label(row1, text="希望校区：").pack(side="left")
        ttk.Combobox(
            row1, textvariable=campus_local,
            values=["江宁校区", "五台校区", "随意"],
            state="readonly", width=15
        ).pack(side="left", padx=8)

        row2 = ttk.Frame(win)
        row2.pack(fill="x", padx=16, pady=6)
        ttk.Label(row2, text="每周最多去其他校区天数：").pack(side="left")
        ttk.Combobox(
            row2, textvariable=other_local,
            values=[0, 1, 2, 3], state="readonly", width=8
        ).pack(side="left", padx=8)

        ttk.Label(
            win,
            text="说明：选“随意”表示不限定主校区；\n"
                 "选具体校区后，超出上述天数上限的方案不会被生成。",
            wraplength=340, justify="left", foreground="#555555"
        ).pack(padx=16, pady=(10, 6), anchor="w")

        def confirm():
            self.campus_var.set(campus_local.get())
            self.other_campus_var.set(other_local.get())
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=confirm).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="稍后在主界面设置", command=win.destroy).pack(side="left", padx=8)

        win.protocol("WM_DELETE_WINDOW", confirm)

    def populate_course_tree(self):
        for item in self.course_tree.get_children():
            self.course_tree.delete(item)

        for c in self.engine.plan:
            code = c["课程编号"]
            if code not in self.pref_score_vars:
                self.pref_score_vars[code] = tk.IntVar(value=50)
                self.class_score_vars[code] = {}
                self.exempt_vars[code] = tk.BooleanVar(value=False)
                self.unwanted_vars[code] = tk.BooleanVar(value=False)

            # 为该课程的每一个班级初始化独立偏好（默认50）。
            for opt in self.engine.options_by_code.get(code, []):
                self.class_score_vars[code].setdefault(
                    opt.class_name, tk.IntVar(value=50)
                )

            group_text = c["课程类别"]
            if c.get("分组"):
                group_text += f" / 第{c['分组']}组"

            self.course_tree.insert(
                "",
                "end",
                iid=code,
                values=(
                    c["课程名称"], code, group_text, c["学分"],
                    "否", "否", self.pref_score_vars[code].get()
                )
            )

    def edit_course_preference(self, event=None):
        item = self.course_tree.focus()
        if not item or self.engine is None:
            return

        code = item
        values = self.course_tree.item(item, "values")
        course = self.engine.plan_by_code.get(code, {})
        course_name = course.get("课程名称", values[0] if values else code)
        options = self.engine.options_by_code.get(code, [])

        win = tk.Toplevel(self)
        win.title(f"设置课程/班级偏好：{course_name}")
        win.geometry("1000x620")
        win.transient(self)
        win.grab_set()

        # ===== 课程级设置 =====
        top = ttk.LabelFrame(win, text="课程级设置")
        top.pack(fill="x", padx=12, pady=10)

        ex = tk.BooleanVar(value=self.exempt_vars[code].get())
        un = tk.BooleanVar(value=self.unwanted_vars[code].get())
        course_score = tk.IntVar(value=self.pref_score_vars[code].get())

        ttk.Label(top, text=course_name, font=("Arial", 12, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 6)
        )
        ttk.Checkbutton(top, text="免修这门课", variable=ex).grid(
            row=1, column=0, sticky="w", padx=12, pady=5
        )
        ttk.Checkbutton(top, text="不想选这门课（硬性排除）", variable=un).grid(
            row=1, column=1, sticky="w", padx=12, pady=5
        )

        ttk.Label(top, text="课程偏好评分：").grid(row=2, column=0, sticky="w", padx=12, pady=8)
        course_scale = ttk.Scale(
            top, from_=0, to=100, variable=course_score, orient="horizontal"
        )
        course_scale.grid(row=2, column=1, sticky="ew", padx=8, pady=8)
        course_value = ttk.Label(top, textvariable=course_score, width=5)
        course_value.grid(row=2, column=2, sticky="w")
        ttk.Label(
            top,
            text="50=中性；>50倾向；<50不想上；100=必须选。",
            foreground="#555555"
        ).grid(row=2, column=3, sticky="w", padx=10)
        top.columnconfigure(1, weight=1)

        # ===== 班级级设置 =====
        box = ttk.LabelFrame(win, text="该课程的全部班级（Excel 中同一班级的多行已合并）")
        box.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        columns = ("class", "time", "address", "campus", "score")
        tree = ttk.Treeview(box, columns=columns, show="headings", height=16, selectmode="browse")
        headings = {
            "class": "班级名称",
            "time": "具体时间 / 周次",
            "address": "教室/地点",
            "campus": "校区",
            "score": "班级偏好"
        }
        widths = {"class": 150, "time": 390, "address": 180, "campus": 90, "score": 90}
        for c in columns:
            tree.heading(c, text=headings[c])
            tree.column(c, width=widths[c], anchor="center")

        tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)

        # 记录 Treeview 行与班级名称的对应关系。
        class_names = []
        for opt in options:
            if code not in self.class_score_vars:
                self.class_score_vars[code] = {}
            if opt.class_name not in self.class_score_vars[code]:
                self.class_score_vars[code][opt.class_name] = tk.IntVar(value=50)
            score_var = self.class_score_vars[code][opt.class_name]
            class_names.append(opt.class_name)
            tree.insert(
                "", "end",
                iid=str(len(class_names) - 1),
                values=(
                    opt.class_name,
                    format_option_schedule(opt),
                    opt.address,
                    opt.campus,
                    score_var.get(),
                )
            )

        ttk.Label(
            box,
            text="说明：同一班级若在不同周次有不同时间/地点，已合并显示；班级评分默认50。100表示优先锁定该班，低于50表示不喜欢该班。",
            foreground="#555555",
            wraplength=900,
            justify="left"
        ).pack(anchor="w", padx=8, pady=6)

        # ===== 当前选中班级的评分编辑区 =====
        edit = ttk.Frame(win)
        edit.pack(fill="x", padx=12, pady=4)
        ttk.Label(edit, text="选中班级的评分：").pack(side="left")
        selected_score = tk.IntVar(value=50)
        selected_label = ttk.Label(edit, text="请选择班级")
        selected_label.pack(side="left", padx=8)
        selected_scale = ttk.Scale(edit, from_=0, to=100, variable=selected_score, orient="horizontal")
        selected_scale.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Label(edit, textvariable=selected_score, width=5).pack(side="left")

        def load_selected(event=None):
            sel = tree.selection()
            if not sel:
                selected_label.config(text="请选择班级")
                selected_score.set(50)
                return
            idx = int(sel[0])
            cls = class_names[idx]
            selected_label.config(text=cls)
            selected_score.set(self.class_score_vars[code][cls].get())

        def write_selected(event=None):
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            cls = class_names[idx]
            self.class_score_vars[code][cls].set(int(round(selected_score.get())))
            tree.set(sel[0], "score", int(round(selected_score.get())))

        tree.bind("<<TreeviewSelect>>", load_selected)
        selected_scale.bind("<ButtonRelease-1>", write_selected)
        selected_score.trace_add("write", lambda *_: write_selected())

        buttons = ttk.Frame(win)
        buttons.pack(pady=10)

        def save():
            # 将当前编辑值写回最后选中的班级。
            write_selected()
            self.exempt_vars[code].set(ex.get())
            self.unwanted_vars[code].set(un.get())
            self.pref_score_vars[code].set(int(round(course_score.get())))
            self.course_tree.item(
                item,
                values=(
                    values[0], values[1], values[2], values[3],
                    "是" if ex.get() else "否",
                    "是" if un.get() else "否",
                    int(round(course_score.get()))
                )
            )
            win.destroy()

        def reset_classes():
            for cls in class_names:
                self.class_score_vars[code][cls].set(50)
            for iid in tree.get_children():
                tree.set(iid, "score", 50)
            if tree.selection():
                selected_score.set(50)

        ttk.Button(buttons, text="所有班级恢复50", command=reset_classes).pack(side="left", padx=6)
        ttk.Button(buttons, text="保存", command=save).pack(side="left", padx=6)
        ttk.Button(buttons, text="取消", command=win.destroy).pack(side="left", padx=6)

    def collect_preferences(self):
        pref = UserPreference()

        # 避开时间
        for block in re.split(r"[;；]+", self.avoid_var.get()):
            block = block.strip()
            if block:
                pref.avoid_slots |= parse_simple_block(block)

        campus = self.campus_var.get()
        pref.preferred_campus = "" if campus == "随意" else campus
        pref.max_other_campus_days = int(self.other_campus_var.get())

        pref.exempt_codes = {
            code for code, var in self.exempt_vars.items()
            if var.get()
        }

        pref.unwanted_codes = {
            code for code, var in self.unwanted_vars.items()
            if var.get()
        }

        # 所有课程始终保留评分；默认50为中性。
        pref.preferred_scores = {
            code: int(var.get())
            for code, var in self.pref_score_vars.items()
        }

        # 班级偏好同样保留全部评分；默认50表示不改变课程级偏好。
        pref.class_scores = {
            (code, class_name): int(var.get())
            for code, class_map in self.class_score_vars.items()
            for class_name, var in class_map.items()
        }

        overlap = pref.exempt_codes & pref.unwanted_codes
        if overlap:
            raise ValueError(
                "以下课程同时被设置为“免修”和“不想选”：" +
                ", ".join(overlap)
            )

        hard_conflict = {
            code for code, score in pref.preferred_scores.items()
            if score >= 100 and (code in pref.exempt_codes or code in pref.unwanted_codes)
        }
        if hard_conflict:
            raise ValueError(
                "以下课程被设为100%必选，但同时被设置为免修/不想选：" +
                ", ".join(sorted(hard_conflict))
            )

        class_required = defaultdict(list)
        for (code, class_name), score in pref.class_scores.items():
            if score >= 100:
                class_required[code].append(class_name)
        ambiguous = {code: names for code, names in class_required.items() if len(names) > 1}
        if ambiguous:
            msg = []
            for code, names in sorted(ambiguous.items()):
                msg.append(f"{code}: " + ", ".join(names))
            raise ValueError(
                "同一课程不能同时将多个班级设为100%必选，请只保留一个班级：\n" +
                "\n".join(msg)
            )

        class_required_conflict = {
            code for code, names in class_required.items()
            if code in pref.exempt_codes or code in pref.unwanted_codes
        }
        if class_required_conflict:
            raise ValueError(
                "以下课程存在100%必选班级，但课程同时被设为免修/不想选：" +
                ", ".join(sorted(class_required_conflict))
            )

        return pref

    def generate_plans(self):
        try:
            if self.engine is None:
                raise ValueError("请先读取培养方案与课程目录。")

            pref = self.collect_preferences()

            planner = Planner(self.engine, pref)
            raw_n = self.n_plans.get()
            n = None if raw_n == "尽可能全部" else int(raw_n)
            self.results = planner.generate(n=n)

            if not self.results:
                self.status.config(text="没有找到同时满足所有硬约束的方案。")
                self.req_text.delete("1.0", "end")
                self.req_text.insert(
                    "end",
                    "没有可行方案。\n\n"
                    "优先检查：\n"
                    "1. 培养方案的课程分组/分组规则是否解析正确；\n"
                    "2. 免修课程是否已正确从对应分组要求中扣除；\n"
                    "3. 100%必选、免修、不想选是否互相冲突；\n"
                    "4. 避开时间与其他校区天数是否过严；\n"
                    "5. 课程安排表中的时间格式是否被正确解析。"
                )
                self.fig.clear()
                self.canvas.draw()
                return

            self.populate_plan_listbox()

            self.current_plan_idx = 0
            self.show_plan()

            solver_name = "CP-SAT" if ORTOOLS_AVAILABLE else "内置剪枝求解器"
            self.status.config(
                text=f"已生成 {len(self.results)} 个不同的实体选课方案；当前：方案1；求解器：{solver_name}"
            )

        except Exception as e:
            messagebox.showerror("生成方案失败", str(e))

    def populate_plan_listbox(self):
        self.plan_listbox.delete(0, "end")
        pref = self.collect_preferences()

        for i, (score, selected) in enumerate(self.results, 1):
            req = special_requirements(selected, pref)
            all_ok = all(v for k, v in req.items() if k != "其他校区天数")
            mark = "✓ 全部满足" if all_ok else "✗ 部分未满足"
            campus_days = req["其他校区天数"]
            vsummary = virtual_course_summary(selected, self.engine)
            vtext = "；".join(vsummary) if vsummary else "无网课占位"
            self.plan_listbox.insert(
                "end", f"方案{i}  [{mark}]  其他校区{campus_days}天  |  {vtext}"
            )

    def on_plan_listbox_select(self, event=None):
        sel = self.plan_listbox.curselection()
        if not sel or not self.results:
            return
        self.current_plan_idx = sel[0]
        self.current_week = 1
        self.show_plan()
        self.status.config(
            text=f"当前：方案 {self.current_plan_idx + 1}/{len(self.results)}"
        )

    def show_plan(self):
        if not self.results:
            return

        score, selected = self.results[self.current_plan_idx]

        # 提取该方案涉及的所有周
        all_weeks = sorted({w for opt in selected for w, _, _ in opt.slots})
        if not all_weeks:
            all_weeks = [1]
        self.weeks_in_plan = all_weeks

        # 初始化周选择器
        self.week_combo.config(values=[str(w) for w in all_weeks])
        if str(self.current_week) not in [str(w) for w in all_weeks]:
            self.current_week = all_weeks[0]
        self.week_combo.set(str(self.current_week))

        n_courses = self.draw_schedule(selected, self.current_plan_idx + 1, week=self.current_week)
        self.week_info_label.config(text=f"（本周有 {n_courses} 门课）")

        # 同步高亮左侧方案列表中的当前项
        self.plan_listbox.selection_clear(0, "end")
        self.plan_listbox.selection_set(self.current_plan_idx)
        self.plan_listbox.see(self.current_plan_idx)

        # 特殊需求状态
        pref = self.collect_preferences()
        req = special_requirements(selected, pref)
        all_ok = all(v for k, v in req.items() if k != "其他校区天数")

        self.req_text.delete("1.0", "end")

        self.req_text.insert(
            "end",
            f"【方案 {self.current_plan_idx + 1} / {len(self.results)}】\n\n",
            "bold"
        )

        total_credit = sum(x.credit for x in selected)
        self.req_text.insert("end", f"总学分：{total_credit:g}\n")
        self.req_text.insert("end", f"目标函数：{score}\n\n")

        self.req_text.insert(
            "end",
            ("全部特殊需求均满足\n\n" if all_ok else "存在未满足的特殊需求\n\n"),
            ("ok" if all_ok else "bad")
        )

        for k, v in req.items():
            if k == "其他校区天数":
                continue
            symbol = "✓" if v else "✗"
            self.req_text.insert("end", f"{symbol} {k}\n", "ok" if v else "bad")

        self.req_text.insert(
            "end",
            f"\n其他校区上课天数：{req['其他校区天数']}\n"
            f"允许最多：{pref.max_other_campus_days}\n"
        )

        vsummary = virtual_course_summary(selected, self.engine)
        if vsummary:
            self.req_text.insert("end", "\n网课选课汇总（网课组合不作为不同方案）：\n" + "-" * 35 + "\n")
            for item in vsummary:
                self.req_text.insert("end", item + "\n")

        self.req_text.insert("end", "\n实际到课课程/班级：\n" + "-" * 35 + "\n")

        for x in sorted([o for o in selected if not o.is_virtual], key=lambda z: (min(z.days) if z.days else 99, z.course_name)):
            self.req_text.insert(
                "end",
                f"{x.course_name}\n"
                f"  {x.class_name} | {x.credit:g}学分 | {x.campus}\n"
                f"  {x.time_text}\n"
                f"  {x.exact_schedule_text}\n\n"
            )

    def draw_schedule(self, selected, plan_no, week=None):
        """按周绘制课表。week=None 时显示所有周，week=int 时只显示该周。"""
        self.fig.clear()
        ax = self.fig.add_subplot(111)

        campuses_seen = set()
        courses_this_week = 0

        for idx, opt in enumerate(selected):
            if not opt.slots:
                continue

            if week is not None:
                week_slots = {(d, p) for w, d, p in opt.slots if w == week}
                if not week_slots:
                    continue
            else:
                week_slots = {(d, p) for w, d, p in opt.slots}

            courses_this_week += 1
            color = campus_color(opt.campus)
            campuses_seen.add(opt.campus)

            day_to_periods = defaultdict(list)
            for d, p in week_slots:
                day_to_periods[d].append(p)

            for day, periods in day_to_periods.items():
                periods = sorted(set(periods))

                groups = []
                start = prev = periods[0]
                for p in periods[1:]:
                    if p == prev + 1:
                        prev = p
                    else:
                        groups.append((start, prev))
                        start = prev = p
                groups.append((start, prev))

                for p1, p2 in groups:
                    ax.bar(
                        day, p2 - p1 + 1,
                        bottom=p1 - 1,
                        width=0.78,
                        alpha=0.85,
                        color=color,
                        edgecolor="white",
                        linewidth=1.0,
                    )

                    height = p2 - p1 + 1
                    if height >= 4:
                        label = f"{opt.course_name}\n第{p1}-{p2}节\n{opt.campus}"
                    elif height >= 2:
                        label = f"{opt.course_name}\n{opt.campus}"
                    else:
                        label = opt.course_name[:6]

                    ax.text(
                        day,
                        (p1 + p2) / 2 - 0.5,
                        label,
                        ha="center",
                        va="center",
                        fontsize=7 if height >= 2 else 6,
                        color="white",
                        fontweight="bold",
                    )

        ax.set_xlim(0.5, 7.5)
        ax.set_ylim(15, 0)
        ax.set_xticks(range(1, 8))
        ax.set_xticklabels(
            ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        )
        ax.set_yticks(range(0, 16))
        ax.set_yticklabels([f"第{i}节 {period_time_range(i, i)}" for i in range(1, 17)])

        ax.grid(axis="y", linestyle="--", alpha=0.4)

        if week is not None:
            ax.set_title(f"选课方案 {plan_no}  ·  第 {week} 周", fontsize=14)
        else:
            ax.set_title(f"选课方案 {plan_no}", fontsize=14)

        ax.set_xlabel("星期", fontsize=11)
        ax.set_ylabel("节次", fontsize=11)

        if campuses_seen:
            handles = [
                plt.Rectangle((0, 0), 1, 1, color=campus_color(c))
                for c in sorted(campuses_seen)
            ]
            ax.legend(
                handles, sorted(campuses_seen),
                loc="upper right", bbox_to_anchor=(1.18, 1.0),
                fontsize=9, title="校区"
            )

        self.fig.tight_layout()
        self.canvas.draw()
        return courses_this_week

    def switch_week(self, delta):
        if not self.weeks_in_plan:
            return
        idx = self.weeks_in_plan.index(self.current_week)
        new_idx = (idx + delta) % len(self.weeks_in_plan)
        self.current_week = self.weeks_in_plan[new_idx]
        self.week_combo.set(str(self.current_week))
        self.show_week()

    def on_week_selected(self, event=None):
        try:
            w = int(self.week_combo.get())
            if w in self.weeks_in_plan:
                self.current_week = w
                self.show_week()
        except ValueError:
            pass

    def show_week(self):
        if not self.results:
            return
        score, selected = self.results[self.current_plan_idx]
        n_courses = self.draw_schedule(selected, self.current_plan_idx + 1, week=self.current_week)
        self.week_info_label.config(text=f"（本周有 {n_courses} 门课）")

    def switch_plan(self, delta):
        if not self.results:
            return

        self.current_plan_idx = (
            self.current_plan_idx + delta
        ) % len(self.results)

        self.current_week = 1
        self.show_plan()
        self.status.config(
            text=f"当前：方案 {self.current_plan_idx + 1}/{len(self.results)}"
        )


    def _overview_matrix(self):
        """返回总览图数据。

        每个单元格只展示“当前方案实际选中的课程”：
        - 未选择：留空，不再用“—”标记；
        - 与上一方案相比发生变化：只给当前方案中“实际选中的课程”加红框；
        - 课程填充色按课程分组区分，采用低饱和度莫兰迪色系。
        """
        row_keys = []
        for c in self.engine.plan:
            code = c["课程编号"]
            opts = self.engine.options_by_code.get(code, [])
            if any(not o.is_virtual for o in opts):
                row_keys.append(("course", code))

        virtual_rows = sorted(
            {x for _, selected in self.results for x in virtual_course_summary(selected, self.engine)}
        )
        row_keys.extend(("virtual", x) for x in virtual_rows)

        labels = []
        matrix = []
        groups = []

        for kind, key in row_keys:
            if kind == "course":
                plan_course = self.engine.plan_by_code[key]
                labels.append(plan_course["课程名称"])
                groups.append(plan_course.get("分组"))

                row = []
                for _, selected in self.results:
                    opt = next(
                        (o for o in selected if o.course_code == key and not o.is_virtual),
                        None,
                    )
                    row.append(opt.class_name if opt else "")
                matrix.append(row)
            else:
                labels.append(key)
                # 从摘要文本反查分组；没有明确分组时保持 None。
                gid = None
                m = re.search(r"网课第(\d+)组", key)
                if m:
                    gid = int(m.group(1))
                groups.append(gid)

                matrix.append([
                    "✓" if key in virtual_course_summary(selected, self.engine) else ""
                    for _, selected in self.results
                ])

        return labels, matrix, groups

    def _make_overview_figure(self, start_idx=0, end_idx=None, row_start=0, row_end=None):
        labels, matrix, groups = self._overview_matrix()
        end_idx = len(self.results) if end_idx is None else min(end_idx, len(self.results))
        row_end = len(labels) if row_end is None else min(row_end, len(labels))

        sub_labels = labels[row_start:row_end]
        sub_matrix = [row[start_idx:end_idx] for row in matrix[row_start:row_end]]
        sub_groups = groups[row_start:row_end]

        ncols = max(1, end_idx - start_idx)
        nrows = max(1, row_end - row_start)

        fig, ax = plt.subplots(
            figsize=(max(10, 1.0 + 1.25 * ncols), max(6.5, 0.45 * nrows + 1.8))
        )
        ax.set_xlim(-0.5, ncols - 0.5)
        ax.set_ylim(nrows - 0.5, -0.5)
        ax.set_xticks(range(ncols))
        ax.set_xticklabels(
            [f"方案{start_idx + i + 1}" for i in range(ncols)],
            rotation=30, ha="right", fontsize=9
        )
        ax.set_yticks(range(nrows))
        ax.set_yticklabels(sub_labels, fontsize=9)

        # 莫兰迪色系：低饱和、偏灰，透明度中高，适合密集总览图。
        morandi = [
            "#A8B0B8",  # 灰蓝
            "#B7A99A",  # 灰棕
            "#A9B7A5",  # 灰绿
            "#C3A6A6",  # 灰粉
            "#A9A4B8",  # 灰紫
            "#C0B59F",  # 卡其
            "#9FAFB0",  # 青灰
            "#B7AAA0",  # 暖灰
        ]
        group_color = {
            gid: morandi[i % len(morandi)]
            for i, gid in enumerate(sorted({g for g in sub_groups if g is not None}))
        }
        neutral_color = "#D3D0CA"
        fill_alpha = 0.68

        for r, row in enumerate(sub_matrix):
            gid = sub_groups[r]
            fill = group_color.get(gid, neutral_color)

            for c, value in enumerate(row):
                # 不论是否选课，都按照课程所属分组进行莫兰迪色填充。
                ax.add_patch(
                    plt.Rectangle(
                        (c - 0.48, r - 0.43), 0.96, 0.86,
                        facecolor=fill,
                        edgecolor="none",
                        alpha=fill_alpha,
                        zorder=1,
                    )
                )

                if value:
                    # 与上一方案相比有差异时，只标出“当前方案已选”的课程。
                    # 如果当前方案未选，即使上一方案选了，也不做红框。
                    # 只要当前已选课程与“任一其他方案”不同，就标出当前单元格。
                    # 注意：当前方案未选时永远不会进入这里，因此不会标记“没选的课”。
                    changed = any(
                        j != c and other_value != value
                        for j, other_value in enumerate(row)
                    )

                    if changed:
                        ax.add_patch(
                            plt.Rectangle(
                                (c - 0.48, r - 0.43), 0.96, 0.86,
                                fill=False,
                                linewidth=2.0,
                                edgecolor="#B85C5C",
                                zorder=3,
                            )
                        )

                    ax.text(
                        c, r, value,
                        ha="center", va="center",
                        fontsize=8.2, wrap=True,
                        color="#333333",
                        zorder=4,
                    )

        # 细网格仅作为阅读辅助，不覆盖课程填充。
        ax.set_xticks([x - 0.5 for x in range(1, ncols)], minor=True)
        ax.set_yticks([y - 0.5 for y in range(1, nrows)], minor=True)
        ax.grid(which="minor", color="white", linewidth=1.2)
        ax.tick_params(which="minor", bottom=False, left=False)
        ax.grid(False, axis="both")

        ax.set_title(
            f"选课方案总览（方案 {start_idx + 1}-{end_idx}；红框仅表示当前方案选中的差异课程）",
            fontsize=13,
        )
        ax.set_xlabel("方案")
        ax.set_ylabel("课程 / 网课分组")

        # 图例：说明不同分组对应不同莫兰迪填充色。
        legend_items = []
        for gid in sorted({g for g in sub_groups if g is not None}):
            label = f"第{gid}组"
            legend_items.append(
                plt.Rectangle((0, 0), 1, 1, facecolor=group_color[gid], alpha=fill_alpha, edgecolor="none", label=label)
            )
        if any(g is None for g in sub_groups):
            legend_items.append(
                plt.Rectangle((0, 0), 1, 1, facecolor=neutral_color, alpha=fill_alpha, edgecolor="none", label="无分组")
            )
        if legend_items:
            ax.legend(
                handles=legend_items,
                loc="upper left",
                bbox_to_anchor=(1.01, 1.0),
                fontsize=8.5,
                title="课程分组",
                frameon=False,
            )

        fig.tight_layout()
        return fig

    def _overview_pages(self):
        """计算方案总览的分页坐标列表 [(c0, c1, r0, r1), ...]；供窗口显示和导出共用。"""
        labels, matrix, groups = self._overview_matrix()
        row_page_size = 24
        col_page_size = 12
        pages = []
        for c0 in range(0, len(self.results), col_page_size):
            for r0 in range(0, len(labels), row_page_size):
                pages.append((c0, min(c0 + col_page_size, len(self.results)), r0, min(r0 + row_page_size, len(labels))))
        return pages or [(0, len(self.results), 0, len(labels))]

    def export_overview_pdf(self):
        """导出完整方案总览 PDF：包含全部分页（所有方案 × 所有课程行），可放大查看、打印。"""
        if not self.results:
            messagebox.showwarning("没有方案", "请先生成方案。")
            return
        path = filedialog.asksaveasfilename(
            title="导出完整方案总览 PDF（含全部分页）",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        pages = self._overview_pages()
        with PdfPages(path) as pdf:
            for pc0, pc1, pr0, pr1 in pages:
                f = self._make_overview_figure(pc0, pc1, pr0, pr1)
                pdf.savefig(f, bbox_inches="tight")
                plt.close(f)
        messagebox.showinfo("导出成功", f"已导出完整 PDF（共 {len(pages)} 页）：\n{path}")

    def export_overview_csv(self):
        """导出完整总览 CSV：包含全部课程行 × 全部方案列，可用 Excel/表格软件进一步分析。"""
        if not self.results:
            messagebox.showwarning("没有方案", "请先生成方案。")
            return
        path = filedialog.asksaveasfilename(
            title="导出完整总览 CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        labels2, matrix2, _groups2 = self._overview_matrix()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["课程/网课分组"] + [f"方案{i+1}" for i in range(len(self.results))])
            for label, row in zip(labels2, matrix2):
                w.writerow([label] + row)
        messagebox.showinfo("导出成功", str(path))

    def show_overview(self):
        """分页显示方案总览（带滚动条，避免图表比窗口大时内容被裁切/按钮消失）。"""
        if not self.results:
            messagebox.showwarning("没有方案", "请先生成方案。")
            return

        pages = self._overview_pages()

        win = tk.Toplevel(self)
        win.title("方案总览图（分页，可滚动查看；导出请用主界面按钮）")
        win.geometry("1250x850")
        win.minsize(700, 500)

        # 底部导航栏先固定占位（side="bottom"），保证任何情况下都可见，不会被图表挤走。
        nav = ttk.Frame(win)
        nav.pack(side="bottom", fill="x", padx=8, pady=6)
        page_var = tk.IntVar(value=0)
        page_label = ttk.Label(nav, text=f"第1页 / 共{len(pages)}页")
        page_label.pack(side="left", padx=8)
        ttk.Label(nav, text="（内容超出窗口时可用鼠标滚轮/右侧及下方滚动条查看）", foreground="#666").pack(side="left", padx=8)

        # 可滚动区域：外层 tk.Canvas + 横竖滚动条，内部放 matplotlib 的 FigureCanvasTkAgg。
        outer = ttk.Frame(win)
        outer.pack(side="top", fill="both", expand=True)
        h_scroll = ttk.Scrollbar(outer, orient="horizontal")
        v_scroll = ttk.Scrollbar(outer, orient="vertical")
        scroll_canvas = tk.Canvas(
            outer, xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set, highlightthickness=0
        )
        h_scroll.config(command=scroll_canvas.xview)
        v_scroll.config(command=scroll_canvas.yview)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        scroll_canvas.pack(side="left", fill="both", expand=True)

        inner = ttk.Frame(scroll_canvas)
        inner_window = scroll_canvas.create_window((0, 0), window=inner, anchor="nw")

        state = {"fig": None, "mpl_canvas": None}

        def render_page(idx):
            idx = max(0, min(idx, len(pages) - 1))
            c0, c1, r0, r1 = pages[idx]
            if state["fig"] is not None:
                plt.close(state["fig"])
            if state["mpl_canvas"] is not None:
                state["mpl_canvas"].get_tk_widget().destroy()
            fig = self._make_overview_figure(c0, c1, r0, r1)
            mpl_canvas = FigureCanvasTkAgg(fig, master=inner)
            mpl_canvas.get_tk_widget().pack()
            mpl_canvas.draw()
            state["fig"] = fig
            state["mpl_canvas"] = mpl_canvas
            page_var.set(idx)
            page_label.config(text=f"第{idx + 1}页 / 共{len(pages)}页")
            # 让滚动区域适配实际内容大小
            inner.update_idletasks()
            scroll_canvas.config(scrollregion=scroll_canvas.bbox("all"))

        def on_mousewheel(event):
            delta = -1 if event.num == 5 or event.delta < 0 else 1
            scroll_canvas.yview_scroll(-delta, "units")

        scroll_canvas.bind_all("<MouseWheel>", on_mousewheel)
        scroll_canvas.bind_all("<Button-4>", on_mousewheel)
        scroll_canvas.bind_all("<Button-5>", on_mousewheel)

        def on_close():
            scroll_canvas.unbind_all("<MouseWheel>")
            scroll_canvas.unbind_all("<Button-4>")
            scroll_canvas.unbind_all("<Button-5>")
            if state["fig"] is not None:
                plt.close(state["fig"])
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)

        ttk.Button(nav, text="← 上一页", command=lambda: render_page(page_var.get() - 1)).pack(side="left")
        ttk.Button(nav, text="下一页 →", command=lambda: render_page(page_var.get() + 1)).pack(side="left", padx=6)
        ttk.Button(nav, text="关闭", command=on_close).pack(side="right", padx=6)

        render_page(0)

    def export_current(self):
        if not self.results:
            messagebox.showwarning("没有方案", "请先生成方案。")
            return

        _, selected = self.results[self.current_plan_idx]

        path = filedialog.asksaveasfilename(
            title="导出当前方案",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("JSON", "*.json")]
        )
        if not path:
            return

        p = Path(path)

        if p.suffix.lower() == ".json":
            payload = [x.as_dict() for x in selected]
            p.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
        else:
            rows = [x.as_dict() for x in selected]
            cols = [
                "课程类别", "分组", "分组规则", "班级名称",
                "课程编号", "课程名称", "学分", "时间", "具体日期与时间",
                "上课地址", "校区"
            ]
            with p.open("w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)

        messagebox.showinfo("导出成功", str(p))

    def export_all(self):
        if not self.results:
            messagebox.showwarning("没有方案", "请先生成方案。")
            return

        path = filedialog.asksaveasfilename(
            title="导出全部方案",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if not path:
            return

        data = []
        pref = self.collect_preferences()

        for i, (score, selected) in enumerate(self.results, 1):
            data.append(
                {
                    "方案": i,
                    "目标函数": score,
                    "特殊需求": special_requirements(selected, pref),
                    "课程": [x.as_dict() for x in selected],
                }
            )

        Path(path).write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        messagebox.showinfo("导出成功", str(path))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", default="", help="培养方案PDF")
    ap.add_argument("--xlsx", default="", help="课程目录Excel")
    args = ap.parse_args()

    app = App()

    if args.pdf:
        app.pdf_path.set(args.pdf)
    if args.xlsx:
        app.xlsx_path.set(args.xlsx)

    app.mainloop()


if __name__ == "__main__":
    main()
