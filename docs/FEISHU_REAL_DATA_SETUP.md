# 飞书真实数据接入指南

本项目已经支持以下真实数据链路：

```text
培养方案 PDF
   ↓
PyPDF 提取文本
   ↓
AI 结构化解析
   ↓
CurriculumSchema
   ↓
飞书「培养方案」表

课程 Excel
   ↓
标准库 XLSX 解析
   ↓
课程目录 + 班级安排
   ↓
飞书「课程目录」表 + 「课程班级」表
```

飞书开放平台支持通过应用凭证获取 tenant_access_token，再调用多维表格记录接口读写数据。项目采用这一服务端模式。citehttps://open.feishu.cn/?lang=zh-CN

## 一、创建飞书企业自建应用

在飞书开放平台创建企业自建应用，并开启机器人能力（机器人不是导入数据所必需，但用于 Hackathon 对话 Demo）。应用需要发布后 API 权限才会生效。

必须妥善保管：

- App ID
- App Secret

不要把真实 Secret 提交到 GitHub。

## 二、创建一个多维表格

建议创建一个“AI智能选课”多维表格，并建立 4 张数据表。

### 表 A：培养方案

主字段建议：

```text
课程编号        文本
课程名称        文本
学分            数字
课程类别        文本
分组            数字或文本
分组规则        文本
必修            文本
学期            数字
AI解析状态      文本
```

### 表 B：课程目录

```text
课程编号        文本
课程名称        文本
学分            数字
校区            文本
来源            文本
```

### 表 C：课程班级

```text
班级名称        文本
课程编号        文本
课程名称        文本
学分            数字
时间            文本
上课地址        文本
校区            文本
数据状态        文本
```

### 表 D：方案结果

后续 AI 求解功能建议使用：

```text
方案编号
总学分
课程数
满意度
跨校区天数
AI推荐理由
课程JSON
约束JSON
```

## 三、获取多维表格 App Token 和 Table ID

打开多维表格后，从 URL/API 配置中取得：

```text
FEISHU_BITABLE_APP_TOKEN
培养方案 table_id
课程目录 table_id
课程班级 table_id
方案结果 table_id
```

项目不把这些值写死在源码中。

## 四、配置环境变量

复制：

```bash
cp .env.example .env
```

填写：

```dotenv
FEISHU_APP_ID=cli_xxx
FEISHU_APP_SECRET=xxx
FEISHU_BITABLE_APP_TOKEN=basexxx

AI_BASE_URL=https://api.openai.com/v1
AI_API_KEY=xxx
AI_MODEL=gpt-4o-mini
```

也可以使用兼容 OpenAI Chat Completions API 的模型服务，只要 `AI_BASE_URL` 与鉴权格式兼容。

## 五、用你的真实培养方案运行

本项目测试数据对应的实际培养方案为：

> 南京医科大学 2026 年度全日制专业学位博士研究生课程方案，专业学位类别为临床医学，领域为皮肤病与性病学。

该 PDF 中实际包含公共必修、专业必修、专业选修和公共选修等课程；例如公共必修中有医学论文写作与学术规范、中国马克思主义与当代、博士英语、学术规范与实验室安全，以及“大数据算法/人工智能与创新/智能法理”三选一；专业选修又包含第3组和第4组规则。fileciteturn15file0L11-L31 fileciteturn15file0L32-L51

运行：

```bash
python sync_to_feishu.py \
  --pdf "pyfady_show(2).pdf" \
  --xlsx "2026-2027学年第一学期（秋季）研究生课程目录及课程安排表(2).xlsx" \
  --curriculum-table tbl培养方案 \
  --course-table tbl课程目录 \
  --class-table tbl课程班级
```

第一次建议先：

```bash
python sync_to_feishu.py ... --dry-run
```

确认 `ai_curriculum_result.json` 后再正式写入飞书。

再次执行时默认按 `课程编号` 跳过已存在课程，因此不会因为重复运行而大量重复创建。需要重新同步时使用：

```bash
--replace
```

## 六、真实课程 Excel 如何进入飞书

项目不会让 AI 猜测时间和地点。Excel 中的：

- 班级名称
- 课程编号
- 课程名称
- 学分
- 时间
- 上课地址
- 校区

由确定性的 XLSX 解析器读取后直接进入飞书；例如实际课程数据包含五台校区、江宁校区，以及分段周次课程。fileciteturn13file8L272-L280

这点非常重要：

> **AI 负责理解培养方案，真实课表数据由程序原样读取。**

因此不会因为大模型幻觉而生成不存在的上课时间。

## 七、Hackathon 建议 Demo

### Demo 1：培养方案自动入库

把 PDF 放到本地：

```text
南京医科大学培养方案.pdf
```

运行：

```text
PDF → AI → 培养方案表
```

现场展示飞书表格中自动出现全部课程和分组规则。

### Demo 2：自然语言选课

用户在飞书机器人发送：

> 我不想周一早八，尽量不要去五台校区，比较喜欢王老师的课，课程尽可能少一点。

AI 将自然语言转成结构化约束，再交给 OR-Tools。

### Demo 3：自动解释

飞书返回：

> 推荐方案 2。该方案满足所有培养方案硬约束，无时间冲突；上课 3 天，跨校区 0 天，并保留了你偏好的课程。

## 八、权限与安全

程序使用企业自建应用的 tenant_access_token 调用多维表格接口。飞书开放平台的开发者资料显示，该凭证可用于服务端代表租户执行 API 操作；多维表格记录可通过开放接口读取和写入。citehttps://open.feishu.cn/community/articles/7271149634339438594

不要把：

```text
AI_API_KEY
FEISHU_APP_SECRET
```

写入 Python 文件、README、Issue 或 GitHub Actions 日志。
