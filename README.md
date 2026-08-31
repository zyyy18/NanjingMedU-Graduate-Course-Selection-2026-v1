# AI CoursePilot for Feishu

**AI 驱动的研究生智能选课决策器｜飞书 × AI × OR-Tools**

> 把复杂的培养方案交给 AI 理解，把课程与偏好交给飞书管理，把真正能毕业、能上课、时间不冲突的组合交给约束优化器求解。

这是在原有研究生智能选课程序基础上的 Hackathon 版本。原程序已经具备 PDF/Excel 读取、课程时间冲突检测、培养方案分组约束、免修/排除、课程偏好、校区限制、多方案生成和 OR-Tools CP-SAT 求解能力。新版本不推倒这些能力，而是在其上增加**通用培养方案 Schema、AI 自然语言理解、飞书多维表格读写和飞书机器人入口**。

## ✨ 核心闭环

```text
┌─────────────── 飞书 ───────────────┐
│ 多维表格：培养方案 / 课程班级 / 结果 │
│              ↑             ↓       │
│         用户自然语言      方案回写   │
└──────────────┬──────────────┘
               ↓
        ┌──────────────┐
        │      AI      │
        │ 培养方案理解 │
        │ 需求理解     │
        │ 结果解释     │
        └──────┬───────┘
               ↓
        ┌──────────────┐
        │  Constraint  │
        │   Engine     │
        │ OR-Tools     │
        │ CP-SAT       │
        └──────┬───────┘
               ↓
        合法 + 优化的多方案
               ↓
        AI 解释 → 飞书回写
```

作品运行时**同时真正使用飞书和 AI**：

- 飞书负责课程数据读取、用户入口和结果沉淀；
- AI 负责从自然语言/非结构化文本中理解约束，并解释结果；
- OR-Tools 负责严格判断“是否可行”。

AI 不直接“猜答案”，最终课程组合必须经过约束求解器验证。

## 🚀 为什么这个版本比原版更通用

原版培养方案解析里存在针对特定培养方案的课程编号集合和分组规则。这是原版只能适用于某一套培养方案的关键原因之一。

新版本把培养方案统一成：

```json
{
  "code": "COURSE001",
  "name": "课程A",
  "credit": 2,
  "group_id": 1,
  "group_rule": "3选1",
  "required": false
}
```

AI 将学校/专业的培养方案转换成这个 Schema，再交给现有约束求解器。更换学校时，核心求解器不需要跟着学校规则重写。

## 🤖 AI 在运行时做什么

### 1. 解析培养方案

培养方案 PDF → 文本 → AI → `CurriculumSchema`。

AI 会识别课程编号、课程名称、学分、课程类别、分组、分组规则、必修课程、学期，并输出无法确认的信息到 `warnings`。

### 2. 理解自然语言需求

例如用户在飞书里说：

> 我不想周一早八，最好不要去五台校区，王老师的课我比较喜欢，课程尽量少一点。

AI 会把它转换成结构化需求，再交给求解器。

### 3. 解释方案

求解器生成方案后，AI 根据真实结果解释为什么推荐方案 1、哪些偏好被满足、有哪些妥协以及备选方案区别。

## 🧠 为什么不能让 AI 直接生成选课结果

选课是严格的约束优化问题，例如：时间不能冲突、培养方案学分必须满足、必修课程必须选择、分组选课规则必须满足、跨校区天数不能超过上限。

因此系统采用：

**AI 提出结构化约束 → Solver 精确求解 → AI 解释结果**。

## 📚 飞书数据表建议

### 1. 培养方案表

| 字段 | 示例 |
|---|---|
| 课程编号 | COURSE001 |
| 课程名称 | 医学统计学 |
| 学分 | 2 |
| 课程类别 | 专业基础 |
| 分组 | 2 |
| 分组规则 | 3选1 |
| 必修 | ✅ |
| 学期 | 1 |

### 2. 课程班级表

| 字段 | 示例 |
|---|---|
| 班级名称 | 01班 |
| 课程编号 | COURSE001 |
| 课程名称 | 医学统计学 |
| 学分 | 2 |
| 时间 | 星期一 第7节-第8节 |
| 上课地址 | XX教学楼 |
| 校区 | 江宁校区 |

### 3. 方案结果表

| 字段 | 示例 |
|---|---|
| 方案 | 1 |
| 总学分 | 10 |
| 课程数 | 5 |
| 用户需求 | 不想周一早八…… |
| AI推荐理由 | …… |
| 特殊需求 | JSON |
| 课程JSON | JSON |

## 🛠️ 安装

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
pip install -r requirements.txt
```

复制 `.env.example` 为 `.env` 并配置 AI 与飞书凭证。`.env` 不应提交到 GitHub。

## ▶️ AI 培养方案解析

```bash
python ingest_curriculum.py your_training_plan.pdf
```

解析结果可以直接写入飞书：

```bash
python ingest_curriculum.py your_training_plan.pdf --write-feishu-table tblxxxx
```

执行链路：

```text
PDF
 ↓
AI 识别培养方案
 ↓
Schema 验证
 ↓
飞书多维表格沉淀
```

## ▶️ 飞书 + AI 选课 Demo

```bash
python -m ai_coursepilot \
  --curriculum-table tblxxxx \
  --class-table tblxxxx \
  --result-table tblxxxx \
  --request "我不想周一早八，尽量不要去其他校区，比较喜欢王老师的课" \
  -n 5
```

运行链路：

```text
飞书读取课程数据
        ↓
AI 解析用户需求
        ↓
OR-Tools 生成方案
        ↓
AI 解释方案
        ↓
飞书写回结果
```

## 🤖 飞书机器人 Demo

启动：

```bash
uvicorn server:app --host 0.0.0.0 --port 8000
```

然后将 `/feishu/webhook` 配置到飞书应用的事件订阅/回调地址。用户在飞书里 @机器人并发送自然语言需求后，机器人会调用 AI、调用 OR-Tools，并将结果解释返回。

## 🔌 飞书能力

本项目使用飞书开放平台服务端 API：

- 获取 `tenant_access_token`
- 多维表格列出记录
- 多维表格新增记录
- 消息发送

飞书开放平台目前提供机器人、网页应用、多维表格、电子表格等开放能力，适合把本项目做成完整的数据与交互闭环。

## 🧩 代码结构

```text
.
├── ai_coursepilot/
│   ├── __init__.py
│   ├── __main__.py
│   ├── schema.py
│   ├── ai_client.py
│   ├── feishu_client.py
│   ├── ingest.py
│   └── adapter.py / service.py
├── course_selector_gui_v15.py
├── server.py
├── ingest_curriculum.py
├── run_feishu_coursepilot.py
├── examples/curriculum.schema.json
├── tests/test_smoke.py
├── .env.example
├── .gitignore
└── requirements.txt
```

`course_selector_gui_v15.py` 保留原有成熟的 GUI + CP-SAT 求解内核；新模块负责通用 Schema、AI、飞书和业务编排。

## 🏆 Hackathon Demo

推荐现场演示：飞书多维表格中准备培养方案与课程班级 → 用户在飞书机器人说“我不想周一早八，尽量不去其他校区，喜欢王老师的课” → AI 解析自然语言 → OR-Tools 生成 3–5 个真正可行的方案 → AI 解释推荐理由 → 结果自动回写飞书。

这样能够直接证明飞书和 AI 都是作品运行链路中的核心组成部分。

## ⚠️ 当前边界

1. AI 输出必须经过 Schema/求解器验证。
2. 不同飞书租户需要根据实际权限配置多维表格和消息权限。
3. 扫描型培养方案 PDF 需要下一阶段接 OCR。
4. 原有 GUI 中仍保留南京医科大学特定培养方案的旧解析逻辑；Hackathon 通用模式走新的 Schema + AI 链路。

## 🔗 相关资料

飞书开放平台：<https://open.feishu.cn/>
MD