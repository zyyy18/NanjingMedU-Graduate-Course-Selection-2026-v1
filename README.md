# AI CoursePilot for Feishu

**AI 驱动的研究生智能选课决策器｜飞书 × AI × OR-Tools**

> 把复杂的培养方案交给 AI 理解，把课程与偏好交给飞书管理，把真正能毕业、能上课、时间不冲突的组合交给约束优化器求解。

这是在原有研究生智能选课程序基础上的 Hackathon 版本。原程序已经具备 PDF/Excel 读取、课程时间冲突检测、培养方案分组约束、免修/排除、课程偏好、校区限制、多方案生成和 OR-Tools CP-SAT 求解能力。新版本在其上增加**通用培养方案 Schema、AI 自然语言理解、真实 PDF/XLSX → 飞书多维表格导入和飞书机器人入口**。

## ✨ 核心闭环

```text
┌─────────────── 飞书 ───────────────┐
│ 多维表格：培养方案 / 课程目录 / 班级 │
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
        │ Constraint   │
        │ Engine       │
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

## 🚀 当前 Hackathon 版本：真实 PDF + Excel 导入飞书

新增 `sync_to_feishu.py`，实现：

```text
培养方案 PDF
   ↓
PyPDF
   ↓
AI 结构化解析
   ↓
CurriculumSchema
   ↓
飞书培养方案表

课程 Excel
   ↓
标准库 XLSX 解析
   ↓
课程目录 + 班级安排
   ↓
飞书课程目录表 + 课程班级表
```

其中**课程时间、地址、校区全部来自真实 Excel，不由 AI 猜测**；AI 专注于培养方案规则理解。这可以避免模型幻觉直接破坏排课结果。

### 使用真实数据

```bash
python sync_to_feishu.py \
  --pdf "pyfady_show(2).pdf" \
  --xlsx "2026-2027学年第一学期（秋季）研究生课程目录及课程安排表(2).xlsx" \
  --curriculum-table tbl培养方案 \
  --course-table tbl课程目录 \
  --class-table tbl课程班级
```

第一次建议：

```bash
python sync_to_feishu.py ... --dry-run
```

确认生成的 `ai_curriculum_result.json` 后再正式写入飞书。重复执行默认跳过已有课程；需要重新同步时使用 `--replace`。

完整配置说明见 [`docs/FEISHU_REAL_DATA_SETUP.md`](docs/FEISHU_REAL_DATA_SETUP.md)。

## 🤖 AI 在运行时做什么

### 1. 解析培养方案

培养方案 PDF → 文本 → AI → `CurriculumSchema`。

AI 会识别课程编号、课程名称、学分、课程类别、分组、分组规则、必修课程、学期，并输出无法确认的信息到 `warnings`。

### 2. 理解自然语言需求

例如用户在飞书里说：

> 我不想周一早八，最好不要去五台校区，王老师的课我比较喜欢，课程尽量少一点。

AI 会把它转换成结构化需求，再交给求解器。

### 3. 解释方案

求解器生成方案后，AI 根据真实结果解释为什么推荐方案、哪些偏好被满足、有哪些妥协以及备选方案差异。

## 🧠 为什么不能让 AI 直接生成选课结果

选课是严格的约束优化问题，例如：时间不能冲突、培养方案学分必须满足、必修课程必须选择、分组选课规则必须满足、跨校区天数不能超过上限。

因此系统采用：

**AI 提出结构化约束 → Solver 精确求解 → AI 解释结果**。

## 📚 飞书数据表

建议建立 4 张表：

- 培养方案
- 课程目录
- 课程班级
- 方案结果

字段模板见 [`docs/FEISHU_TABLE_TEMPLATE.md`](docs/FEISHU_TABLE_TEMPLATE.md)。

## 🛠️ 安装

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
pip install -r requirements.txt
```

复制 `.env.example` 为 `.env`。程序会自动读取 `.env`。

## ▶️ 飞书 + AI 选课 Demo

```bash
python -m ai_coursepilot \
  --curriculum-table tblxxxx \
  --class-table tblxxxx \
  --result-table tblxxxx \
  --request "我不想周一早八，尽量不要去其他校区，比较喜欢王老师的课" \
  -n 5
```

## 🤖 飞书机器人 Demo

启动：

```bash
uvicorn server:app --host 0.0.0.0 --port 8000
```

将 `/feishu/webhook` 配置到飞书应用事件回调后，用户可以直接在飞书中发送自然语言选课需求。

## 🧩 代码结构

```text
.
├── ai_coursepilot/
│   ├── schema.py
│   ├── ai_client.py
│   ├── feishu_client.py
│   ├── feishu_import.py
│   ├── ingest.py
│   └── ...
├── course_selector_gui_v15.py
├── sync_to_feishu.py
├── ingest_curriculum.py
├── server.py
├── docs/
├── examples/
├── tests/
├── .env.example
└── requirements.txt
```

`course_selector_gui_v15.py` 保留原有 GUI + CP-SAT 求解内核；Hackathon 新入口通过 Schema + AI + 飞书把它扩展为通用数据闭环。

## 🏆 Hackathon Demo 推荐流程

1. 在飞书多维表格准备空的“培养方案 / 课程目录 / 课程班级 / 方案结果”表。
2. 运行 `sync_to_feishu.py`，展示真实培养方案 PDF 被 AI 解析并写入飞书。
3. 展示真实课程 Excel 被确定性读取并写入飞书。
4. 在飞书机器人中输入自然语言偏好。
5. AI 将偏好转换为结构化约束。
6. OR-Tools 生成多个真正可行的方案。
7. AI 解释推荐理由并将结果沉淀回飞书。

这个流程能够直接体现：**飞书不是结果展示页，AI 也不是贴上去的文本，两者都在作品运行链路中承担实际职责。**

## ⚠️ 当前边界

1. AI 输出必须经过 Schema/求解器验证。
2. 不同飞书租户需要根据实际权限配置多维表格和消息权限。
3. 扫描型培养方案 PDF 需要 OCR；当前版本针对可提取文本的 PDF。
4. 原有 GUI 中仍保留南京医科大学特定培养方案的旧解析逻辑；Hackathon 新模式走通用 Schema + AI 链路。
5. 当前仓库没有提交你的真实 App Secret、API Key、培养方案 PDF 和课程 Excel，以避免泄露凭证或原始业务数据。

## 🔗 相关资料

飞书开放平台：<https://open.feishu.cn/>
