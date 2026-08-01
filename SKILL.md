---
name: huicheng-ppt
description: 按照汇成医美教育公司模板格式生成PPT。严格复刻模板色调、字体、坐标和元素结构，支持封面/目录/章节/内容/结束页，23种内容页布局（含7种 dashiai-ppt 风格卡片布局），生成可编辑 .pptx。
---

# 汇成医美教育 PPT 生成器

基于汇成医美教育 (HCYM EDUCATION GROUP) 公司模板，生成品牌一致的可编辑 PPTX 文件。

**本 skill 已整合 dashiai-ppt 核心布局概念**，在原有16种基础布局上新增7种卡片/形状布局（对比卡片、流程圆圈、标签列表、高亮数据框、双指标、卡片网格、Before/After），共 **23种内容页布局**，无需安装 dashiai-ppt 即可使用。

## Skill 目录

当前 SKILL.md 所在目录为 `<skill-root>`。

| 路径 | 用途 |
|------|------|
| `<skill-root>/templates/template.pptx` | 原始模板文件（5页） |
| `<skill-root>/media/` | 12个品牌素材文件 |
| `<skill-root>/scripts/generate.py` | PPT 生成脚本（1191行，包含所有布局逻辑） |

依赖 pptx skill 的工具链（add_slide.py、clean.py、validate.py）。

## 核心原则

1. **严格保持模板色调** — 所有页面使用模板原图作为全屏背景，禁止添加遮罩层、半透明覆盖或纯色背景
2. **精确复刻坐标** — 模板已有的元素位置、大小、样式不得改动，仅替换文本内容
3. **内容只在正文区** — 所有新增文字放在 y=1.5in 以下的正文区域，不得遮挡页眉（y<1.2in）或底部栏（y>6.5in）
4. **自包含布局系统** — 已整合 dashiai-ppt 核心布局概念（流程、指标、对比、引言、卡片），无需额外依赖

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
| 品牌绿 | `#46A53B` | PART标签、kicker强调色、卡片填充 |
| 深绿 | `#213A25` | 封面主讲人标签 |
| 金色 | `#FBB03B` | 分割线、数字强调、项目符号、小标题 |
| 白色 | `#FFFFFF` | 主文字色 |
| 浅灰 | `#F5F5F5` | 卡片背景 |
| 浅绿 | `#F0F7EE` | 卡片背景（绿色系） |
| 浅黄 | `#FFF8EC` | 卡片背景（金色系） |
| 正文色 | `#333333` | 内容页正文/要点 |
| 辅助色 | `#555555` | 说明文字/描述 |

### 字体

- 中文正文：微软雅黑
- 英文/数字强调：Arial
- 封面/目录/章节页文本：继承模板原样式（主题字体 `+mj-ea`）

### 字号规范（内容页）

| 元素 | 字号(pt) | 粗体 | 颜色 |
|------|---------|------|------|
| 标题 | 36 | 是 | 白色 |
| 金色分割线 | — | — | #FBB03B, 宽2.5in |
| 正文/要点 | 18 | 否 | #333333 |
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
  "slides": [
    {"type": "cover", "title": "年度汇报", "subtitle": "HCYM EDUCATION GROUP", "subtitle2": "ANNUAL REPORT", "presenter": "张三"},
    {"type": "toc", "items": ["业绩回顾", "核心指标", "流程展示", "未来展望"]},
    {"type": "section", "part": "01", "year": "2026年", "title": "业绩回顾"},
    {"type": "content", "title": "年度营收", "layout": "stat", "number": "3.2亿", "unit": "元", "desc": "同比增长 25%"},
    {"type": "content", "title": "对比卡片", "layout": "comparison_cards", "left_title": "方案A", "left_items": ["优点1"], "right_title": "方案B", "right_items": ["优点2"]},
    {"type": "content", "title": "三步法", "layout": "process_circles", "steps": [{"title": "第一步", "desc": "描述"}]},
    {"type": "content", "title": "标签列表", "layout": "tag_list", "tags": [{"term": "标签名", "note": "说明文字"}]},
    {"type": "content", "title": "高亮数据", "layout": "highlight_box", "big_number": "10x", "big_label": "效率提升", "secondaries": [{"value": "70%", "label": "时间节省"}]},
    {"type": "content", "title": "双指标", "layout": "dual_stat", "left_value": "30min", "left_label": "手工", "right_value": "10sec", "right_label": "AI"},
    {"type": "content", "title": "卡片网格", "layout": "card_grid", "cards": [{"title": "标题", "desc": "描述"}]},
    {"type": "content", "title": "前后对比", "layout": "before_after", "before_title": "Before", "before_items": ["旧方式"], "after_title": "After", "after_items": ["新方式"]},
    {"type": "end"}
  ]
}
```

## 内容页布局（23种）

### 基础布局（16种）

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

### dashiai-ppt 风格新增布局（7种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `comparison_cards` | title, left_title, left_items[], right_title, right_items[] | **两张圆角卡片对比**，左绿右黄边框+背景色，视觉区分明确 |
| `process_circles` | title, steps[{title, desc}] | **编号圆圈+连接线**，绿色圆形编号+金色箭头+描述文字 |
| `tag_list` | title, tags[{term, note}] | **彩色标签+描述**，每个标签为圆角矩形，6色轮转（绿/金/蓝/橙/青/红） |
| `highlight_box` | title, big_number, big_label, secondaries[{value, label}] | **大数字高亮框+副指标卡片**，主数字突出+下方3个辅助指标 |
| `dual_stat` | title, left_value, left_label, right_value, right_label | **两个大数字并排**，左绿右金圆角卡片对比 |
| `card_grid` | title, cards[{title, desc}] | **网格卡片**，2×3或3×N排列，每张卡片有彩色背景+左侧竖条装饰 |
| `before_after` | title, before_title, before_items[], after_title, after_items[] | **Before/After对比**，左红右绿卡片+中间绿色箭头 |

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

## 布局选择建议

| 内容类型 | 推荐布局 |
|---------|---------|
| 方案对比/优劣分析 | `comparison_cards` 或 `before_after` |
| 步骤/流程/方法论 | `process_circles` |
| 多个并列概念/标签 | `tag_list` |
| 突出核心数据 | `highlight_box` 或 `dual_stat` |
| 多项能力/功能展示 | `card_grid` |
| 简单要点罗列 | `bullets` 或 `takeaway` |
| 大数字/KPI | `stat` |
| 多指标横排 | `metrics` |
| 时间线/里程碑 | `timeline` |
| 双栏对照 | `two_column` 或 `comparison` |
| 引言/开场 | `intro` |
| 金句/名言 | `quote` |

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
3. 内容页：克隆 slide4，清除原元素，按布局添加文本框、形状和金色分割线
4. 新增形状原语：`make_rect`（圆角矩形）、`make_circle`（圆形），支持填充色+文本
5. 结束页：保留模板原样
6. 重新排序 sldIdLst，清理无用文件，打包输出
