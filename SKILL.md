---
name: "huicheng-ppt"
description: "按照汇成医美教育公司模板格式生成PPT。基于公司品牌PPT模板（绿色主题+金色点缀），严格复刻模板的色调、字体、坐标和元素结构。支持封面、目录、章节分隔页、16种内容页布局和结束页。生成可编辑的 .pptx 文件。当用户要求制作汇成风格PPT、公司模板PPT、医美教育PPT时使用。"
---

---
name: huicheng-ppt
description: 按照汇成医美教育公司模板格式生成PPT。严格复刻模板色调、字体、坐标和元素结构，支持封面/目录/章节/内容/结束页，生成可编辑 .pptx。已内置 dashiai-ppt 核心布局概念。
---

# 汇成医美教育 PPT 生成器

基于汇成医美教育 (HCYM EDUCATION GROUP) 公司模板，生成品牌一致的可编辑 PPTX 文件。

**本 skill 已分析 dashi-ppt 1020种布局，精选16种适合企业汇报的布局**（流程页、数据指标页、对比页、引言页等），无需安装 dashiai-ppt 即可使用。

## Skill 目录

当前 SKILL.md 所在目录为 `<skill-root>`。

| 路径 | 用途 |
|------|------|
| `<skill-root>/templates/template.pptx` | 原始模板文件（5页） |
| `<skill-root>/media/` | 12个品牌素材文件 |
| `<skill-root>/scripts/generate.py` | PPT 生成脚本（包含所有布局逻辑） |

依赖 pptx skill 的工具链（add_slide.py、clean.py、validate.py）。

## 核心原则

1. **严格保持模板色调** — 所有页面使用模板原图作为全屏背景，禁止添加遮罩层、半透明覆盖或纯色背景
2. **精确复刻坐标** — 模板已有的元素位置、大小、样式不得改动，仅替换文本内容
3. **内容只在正文区** — 所有新增文字放在 y=1.5in 以下的正文区域，不得遮挡页眉（y<1.2in）或底部栏（y>6.5in）
4. **自包含布局系统** — 已整合 dashiai-ppt 核心布局概念（流程、指标、对比、引言），无需额外依赖

## 模板结构

模板共 5 页，尺寸 13.33" × 7.50"（宽屏）。

| 页 | 类型 | 可编辑文本 | 背景图 |
|----|------|-----------|--------|
| 1 | 封面 | 主标题、英文副标题(2行)、主讲人 | image1.jpeg |
| 2 | 目录 | 5个条目标题("目录标题1"~"目录标题5") | image3.jpeg |
| 3 | 章节分隔页 | 年份、章节标题、PART编号 | image8.jpeg |
| 4 | 内容展示页 | 无文本（纯图+底部栏） | image10.jpeg + image11.png |
| 5 | 结束页 | 无文本 | image1.jpeg + image12.png |

## 设计规范

### 配色

| 用途 | 色值 | 说明 |
|------|------|------|
| 品牌绿 | `#46A53B` | PART标签、kicker强调色 |
| 深绿 | `#213A25` | 封面主讲人标签 |
| 金色 | `#FBB03B` | 分割线、数字强调、项目符号、小标题 |
| 白色 | `#FFFFFF` | 主文字色 |

### 字体

- 中文正文：微软雅黑
- 英文/数字强调：Arial
- 封面/目录/章节页文本：继承模板原样式（主题字体 `+mj-ea`）

### 字号规范（内容页）

| 元素 | 字号(pt) | 粗体 | 颜色 |
|------|---------|------|------|
| 标题 | 36 | 是 | 白色 |
| 金色分割线 | — | — | #FBB03B, 宽2.5in |
| 正文/要点 | 18 | 否 | 白色 |
| 大数字(stat) | 72 | 是 | 金色 |
| 指标数字(metrics) | 48 | 是 | 白色 |
| 流程编号 | 36 | 是 | 金色 |
| 单位 | 24 | 否 | 白色/金色 |
| 双栏/对比小标题 | 22 | 是 | 金色 |
| 对比/双栏正文 | 14-15 | 否 | 白色 |
| 引用文字 | 28 | 否 | 白色 |
| 引用来源 | 14 | 否 | 金色 |
| Kicker标签 | 16 | 是 | 绿色 #46A53B |
| Lead导语 | 18 | 否 | 白色 |

## JSON 计划格式

```json
{
  "title": "PPT主标题",
  "subtitle": "HCYM EDUCATION GROUP",
  "subtitle2": "ANNUAL REPORT",
  "presenter": "主讲人姓名",
  "output": "output.pptx",
  "slides": [
    {"type": "cover", "title": "年度汇报", "subtitle": "HCYM EDUCATION GROUP", "subtitle2": "ANNUAL REPORT", "presenter": "张三"},
    {"type": "toc", "items": ["业绩回顾", "核心指标", "流程展示", "未来展望"]},
    {"type": "section", "part": "01", "year": "2026年", "title": "业绩回顾"},
    {"type": "content", "title": "年度营收", "layout": "stat", "number": "3.2亿", "unit": "元", "desc": "同比增长 25%"},
    {"type": "content", "title": "课程亮点", "layout": "bullets", "bullets": ["全新课程体系上线", "学员满意度 98%"]},
    {"type": "content", "title": "团队架构", "layout": "two_column", "left_title": "教学团队", "left_items": ["主讲讲师 12 人"], "right_title": "运营团队", "right_items": ["课程顾问 8 人"]},
    {"type": "content", "title": "金句", "layout": "quote", "quote": "医美教育的核心是真诚", "source": "汇成医美教育"},
    {"type": "content", "title": "核心数据", "layout": "metrics", "metrics": [{"label": "年营收", "value": "3.2", "unit": "亿元", "sub": "同比+25%"}]},
    {"type": "content", "title": "培训流程", "layout": "process", "steps": [{"title": "需求调研", "desc": "了解企业需求"}, {"title": "课程设计", "desc": "定制方案"}]},
    {"type": "content", "title": "方案对比", "layout": "comparison", "left_header": "传统模式", "right_header": "汇成模式", "comparison_items": [{"left": "统一教材", "right": "定制课程"}]},
    {"type": "content", "title": "发展愿景", "layout": "intro", "kicker": "VISION 2027", "lead": "成为医美教育领域最具影响力的培训平台"},
    {"type": "content", "title": "详细说明", "layout": "title_content", "content": "这里是详细的文字内容..."},
    {"type": "content", "title": "图文展示", "layout": "image_text", "side_title": "产品特点", "side_items": ["专业", "高效"]},
    {"type": "content", "title": "发展历程", "layout": "timeline", "events": [{"year": "2020", "event": "公司成立"}, {"year": "2023", "event": "全国布局"}]},
    {"type": "end"}
  ]
}
```

## 内容页布局（16种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `bullets` | title, bullets[] | 标题+金色分割线+要点列表（默认） |
| `stat` | title, number, unit?, desc? | 大数字展示+单位+描述 |
| `two_column` | title, left_title, left_items[], right_title, right_items[] | 左右双栏对比 |
| `title_content` | title, content | 标题+大段文字 |
| `quote` | title, quote, source? | 引用金句+来源 |
| `image_text` | title, side_title, side_items[] | 左图右文 |
| `timeline` | title, events[{year, event}] | 时间线 |
| `metrics` | title, metrics[{label, value, unit, sub}] | 多指标卡片横排 |
| `process` | title, steps[{title, desc}] | 编号流程步骤（最多4步） |
| `comparison` | title, left_header, right_header, comparison_items[{left, right}] | 方案对比表 |
| `intro` | title, kicker?, lead? | 引言页（绿色kicker+导语） |
| `table` | title, headers[], rows[][] | 数据表格 |
| `team` | title, members[{name, role, desc}] | 团队成员卡片 |
| `case` | title, case_name, case_result, case_detail | 案例展示 |
| `funnel` | title, stages[{label, value}] | 漏斗/金字塔分级 |
| `takeaway` | title, takeaways[] | 核心结论要点 |

## 工作流程

1. **理解需求** — 从用户输入中提取 PPT 标题、各页内容
2. **生成 JSON 计划** — 按上述格式构建 slides 数组，选择合适的 layout
3. **运行生成脚本**：
```bash
python3 <skill-root>/scripts/generate.py <plan.json> \
  --pptx-scripts <pptx-skill-scripts-dir> \
  --output <output.pptx> \
  --skill-root <skill-root>
```
4. **验证** — 用 pptx skill 的 validate.py 检查
5. **交付** — 提供 .pptx 文件

## 素材清单

| 文件 | 用途 |
|------|------|
| image1.jpeg | 封面/结束页全屏背景 |
| image2.png | 封面/结束页右上角 Logo |
| image3.jpeg | 目录页全屏背景 |
| image4.png ~ image7.png | 目录页编号圆圈装饰 |
| image8.jpeg | 章节分隔页全屏背景 |
| image9.png | 章节页 PART 标签装饰 |
| image10.jpeg | 内容页全屏背景 |
| image11.png | 内容页底部品牌栏 |
| image12.png | 结束页中央图片 |

## 技术实现

使用 unzip → XML 编辑 → repack 方式：
1. 解压模板 PPTX
2. 封面/目录/章节页：直接替换模板文本（保留原位置、原样式）
3. 内容页：克隆 slide4，清除原元素，按布局添加文本框和金色分割线
4. 结束页：保留模板原样
5. 重新排序 sldIdLst，清理无用文件，打包输出
