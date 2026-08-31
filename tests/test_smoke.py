from ai_coursepilot.schema import CurriculumCourse,CurriculumSchema

def test_schema():
    c=CurriculumCourse(code="A",name="课程A",credit=2,group_id=1,group_rule="3选1")
    s=CurriculumSchema(courses=[c])
    assert s.to_legacy_plan()[0]["课程编号"]=="A"
    assert s.to_legacy_plan()[0]["分组规则"]=="3选1"
