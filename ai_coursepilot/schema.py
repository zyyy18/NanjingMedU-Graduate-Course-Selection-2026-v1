"""统一的培养方案与选课数据结构。"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any

@dataclass
class CurriculumCourse:
    code: str
    name: str
    credit: float
    category: str = ""
    group_id: int | None = None
    group_rule: str = ""
    required: bool = False
    semester: int | None = None
    def to_legacy_dict(self) -> dict[str, Any]:
        return {"课程编号": self.code, "课程名称": self.name, "学分": self.credit, "课程类别": "A" if self.required else self.category, "分组": self.group_id, "分组规则": self.group_rule, "学期": self.semester or 0, "学时": 0}

@dataclass
class CurriculumSchema:
    school: str = ""
    program: str = ""
    semester: str = ""
    courses: list[CurriculumCourse] = field(default_factory=list)
    source_text_summary: str = ""
    warnings: list[str] = field(default_factory=list)
    def to_legacy_plan(self) -> list[dict[str, Any]]:
        return [c.to_legacy_dict() for c in self.courses]

@dataclass
class UserIntent:
    natural_language: str
    avoid_time: list[str] = field(default_factory=list)
    preferred_campus: str = ""
    max_other_campus_days: int = 3
    exempt_courses: list[str] = field(default_factory=list)
    unwanted_courses: list[str] = field(default_factory=list)
    preferred_courses: dict[str, int] = field(default_factory=dict)
    preferred_classes: dict[str, int] = field(default_factory=dict)
    required_courses: list[str] = field(default_factory=list)
    required_classes: dict[str, int] = field(default_factory=dict)
    objective_weights: dict[str, float] = field(default_factory=lambda: {"preference": .5, "campus": .2, "days": .2, "course_count": .1})
