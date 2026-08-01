---
name: huicheng-ppt
description: 按照汇成医美教育公司模板格式生成PPT。严格复刻模板色调、字体、坐标和元素结构，支持封面/目录/章节/内容/结束页，126种内容页布局（含图表、仪表盘、网格、雷达图、瀑布图、Marimekko、思维导图、K线图、玫瑰图、子弹图、波特五力、组合图、主题河流等），生成可编辑 .pptx。
---

# 汇成医美教育 PPT 生成器

基于汇成医美教育 (HCYM EDUCATION GROUP) 公司模板，生成品牌一致的可编辑 PPTX 文件。

**本 skill 已整合 dashiai-ppt 核心布局概念**，在原有16种基础布局上新增17种卡片/形状布局，再加24种进阶图表布局，再加10种实用分析布局，再加11种专业图表布局，再加15种高频补充布局（Marimekko变宽图、分组柱状图、趋势折线图、产业链图、日历热力图、轨道枢纽图、三联面板、计量条、帕累托图、变化量对比、里程碑、光谱定位图、Logo墙、瀑布流网格、转化阶梯），再加11种动画布局静态化适配（思维导图、网络节点图、图片拼贴、卫星数据、气泡时间线、冰柱图、K线图、成熟度曲线、字阵、比例带、唱片），再加20种 dashi-ppt 蒸馏迁移布局（玫瑰图、直方图、群言墙、地铁线路图、天平对比、波特五力、术语表、专辑清单、分组括号、三视野、架构栈、分层防线、三角串联、闭环循环、生态网络、编年史、宣言主张、特性对照表、点阵计数、子弹图），再加2种最终补齐布局（组合图/柱线双轴、主题河流/中心流式堆叠），共 **126种内容页布局**，无需安装 dashiai-ppt 即可使用。

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
3. **内容只在正文区** — 所有新增文字放在 y=1.5in 以下的正文区域，不得遮挡页眉（y<1.2in）或底部栏（y>6.57in，即 EMU 6007100 以上）
4. **文字自动适配** — `make_textbox` 和 `make_rect` 的 bodyPr 均含 `normAutofit`，文字超出时 PowerPoint 自动等比缩小，避免溢出
5. **自包含布局系统** — 已整合 dashiai-ppt 核心布局概念（流程、指标、对比、引言、卡片），无需额外依赖

## 模板结构

模板共 5 页，尺寸 13.33" × 7.50"（宽屏）。

| 页 | 类型 | 可编辑文本 | 背景图 |
|----|------|-----------|--------|
| 1 | 封面 | 主标题、英文副标题(2行)、主讲人 | image1.jpeg |
| 2 | 目录 | 5个条目标题("目录标题1"~"目录标题5") | image3.jpeg |
| 3 | 章节分隔页 | 年份、章节标题、PART编号 | image8.jpeg |
| 4 | 内容展示页 | 无文本（纯背景+底部栏，底部栏 rId2=image11.png） | image10.jpeg + image11.png |
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
| 标题 | 36 | 是 | 品牌绿 #46A53B |
| 正文/要点 | 18 | 否 | #333333 |
| 大数字(stat) | 72 | 是 | 品牌绿 #46A53B |
| 指标数字(metrics) | 48 | 是 | 品牌绿 #46A53B |
| 流程编号 | 36 | 是 | 金色 #FBB03B |
| 单位 | 24 | 否 | 金色 #FBB03B |
| 双栏/对比小标题 | 22 | 是 | 金色 #FBB03B |
| 对比/双栏正文 | 14-15 | 否 | #333333 |
| 引用文字 | 28 | 否 | #333333 |
| 引用来源 | 14 | 否 | 金色 #FBB03B |
| Kicker标签 | 16 | 是 | 品牌绿 #46A53B |
| Lead导语 | 18 | 否 | #333333 |

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

## 内容页布局（126种）

### 基础布局（16种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `bullets` | title, bullets[] | 标题+要点列表（默认） |
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

### 新增布局（参考 dashi-ppt 语义角色，10种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `swot` | quadrants{S,W,O,T}[], labels? | **SWOT四象限**，四色实填卡片（绿/红/金/深绿）+ 白字 |
| `quadrant` | items{q1,q2,q3,q4}[], labels? | **四象限矩阵**，类似SWOT但通用象限标签 |
| `checklist` | items[{text, done?}] | **清单打卡**，✓已完成项绿色高亮，未完成灰色 |
| `scorecard` | items[{label, score, max, desc?}] | **评分卡**，进度条按得分比例填充，颜色自动切换 |
| `stair` | steps[{title, desc}] | **阶梯递进**，逐级升高的品牌色台阶 |
| `flywheel` | items[{label}][] + center? | **飞轮循环**，中心圆+周围节点环形排列 |
| `statement` | text, sub? | **大字宣言**，金色引号装饰+深绿大文字 |
| `journey` | stages[{stage}] + rows[{label, cells[]}] | **用户旅程图**，阶段标签+行维度矩阵 |
| `pricing` | plans[{name, price, unit, features[], highlight?}] | **价格方案**，多列卡片，主推方案高亮 |
| `faq` | items[{q, a}] | **FAQ常见问题**，Q绿色标签+A金色标签交替 |

### 本轮新增布局（图表 / 仪表盘 / 网格，9种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `pie_chart` | slices[{label, value}], donut? | **饼图/环形图**，左侧扇形图+右侧图例，设 `donut:true` 变环形 |
| `bar_chart` | bars[{label, value, unit?, color?}] | **横向条形图**，左侧标签+背景条+数据条+右侧数值 |
| `dashboard` | cards[{label, value, unit?, trend?}] | **仪表盘 2×3 指标卡**，6色卡片网格，显示趋势文字 |
| `kpi_card` | cards[{label, value, unit?, trend?, trend_val?}] | **KPI 指标卡**，最多4列，大字数值+趋势箭头 |
| `hero_banner` | kicker?, title, subtitle?, tagline? | **英雄横幅**，全宽绿色大卡，金色装饰线+标题+副文 |
| `numbered_list` | items[{text, desc?}] | **编号列表**，绿色数字圆圈+文字，最多8项 |
| `matrix_2x2` | quadrants{q1,q2,q3,q4} | **2×2 矩阵**，四色边框卡片，顶部色条+标题+内容列表 |
| `chart_placeholder` | chart_type?, caption?, note? | **图表占位**，虚线框+示意图形（bar 柱状/line 折线），`line` 类型使用对角线连接数据点 |
| `feature_grid` | items[{icon?, title, desc?}] | **特性网格**，2×3卡片，圆形图标+标题+描述，顶部色条装饰 |

### 本轮补齐布局（dashi-ppt 语义角色全覆盖，15种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `radar_chart` | axes[{label, value}] 或 axes[]+values[] | **雷达图 / 能力图**，同心圆网格+轴线+数据多边形连线+标签，右侧图例 |
| `pyramid` | levels[{label, value?}] | **金字塔**，逐级缩小的层级条，6色填充，顶部最宽 |
| `roadmap` | phases[{title, period?, desc?, items?}] | **路线图**，阶段编号圆圈+时间标签+内容卡片，最多5阶段 |
| `venn` | groups[{label, items?}] | **维恩图**，2~3组半透明圆圈品字排列+右侧说明 |
| `ranking` | items[{rank?, name, value, desc?}] | **排行榜**，金/银/铜奖牌+名称+数值+描述，最多6项 |
| `waterfall` | values[], labels[] | **瀑布图**，正负值累积柱状图，绿色正向/红色负向 |
| `heatmap` | rows[{label, values[]}], col_labels? | **热力矩阵**，颜色深浅按数值（0~100），最多5行×8列 |
| `gantt` | tasks[{name, start, duration}], periods? | **甘特图 / 排期**，任务名+彩色进度条+时间段标签 |
| `cycle` | steps[{label}], center? | **循环图 / 闭环**，中心圆+周围节点环形排列 |
| `big_number` | number, label?, desc?, suffix? | **超大数字海报**，深绿背景卡片+超大数字+金色装饰线 |
| `gallery` | items[{title, desc?, tag?}] | **作品/案例画廊**，图片占位+标签+标题+描述，最多4列 |
| `layers` | layers[{label, desc?}] | **层级架构图**，横向层级条+标签+描述，最多5层 |
| `bento` | cells[{title, value?, desc?}] | **便当格卡片**，2行混合布局+彩色边框+数值展示 |
| `gauge` | value, label?, target?, unit?, sub_metrics? | **仪表盘 / 达成率**，中心大数值+百分比+副指标卡片 |
| `testimonial` | quotes[{quote, author?, role?}] 或 quote+author | **证言/引述卡**，1~3张引述卡片+引号装饰+作者署名 |

### 再次补齐布局（dashi-ppt 高频通用场景，10种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `treemap` | items[{label, value}] | **矩形树图**，按数值比例排列嵌套矩形，两行布局，6色轮转 |
| `scatter` | points[{label, x, y, size?}], x_label?, y_label? | **散点/气泡图**，2D坐标轴+彩色圆点，支持尺寸变化 |
| `stacked_bar` | categories[], series[{label, values[], color?}] | **百分比堆叠柱状图**，横向堆叠条+百分比标注+图例 |
| `profile` | name, subtitle?, metrics[{label, value}], desc?, tags[]? | **人物/企业档案卡**，头像区+标签+指标行+描述 |
| `spotlight` | big_stat, stat_label?, desc?, highlights[{label, value}] | **特写聚焦**，深绿大数字区+右侧高亮指标列表 |
| `risk` | risks[{name, probability, impact, desc?}] | **风险矩阵**，3×3概率-影响网格+右侧风险列表 |
| `swimlane` | lanes[{label, steps[{text}]}] | **泳道流程**，多角色横向泳道，每行4色标签+步骤条 |
| `overview` | summary?, key_points[{label, value, desc?}] | **全局概览**，摘要横幅+4列彩色指标卡 |
| `principles` | items[{title, desc?}] | **核心原则**，编号圆圈+标题+描述，最多5条竖排 |
| `org_chart` | nodes[{label, level, desc?}] | **组织架构/生态图**，层级节点居中排列，最多4层×5节点 |

### 补齐 dashi-ppt 专业图表（11种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `bump` | series[{label, ranks[]/values[]}], periods[]? | **名次变迁图**，多系列排名变化+对角连线，圆点+标签 |
| `dumbbell` | items[{label, before, after}] | **哑铃图**，前后两色圆点+连接线，适合对比变化 |
| `lollipop` | items[{label, value}] | **棒棒糖图**，细杆+圆点，清洁的条形图替代 |
| `waffle` | items[{label, value, color?}], total? | **华夫饼图**，10×10 方格按百分比填色 |
| `radial_bar` | items[{label, value, max?}] | **径向条形图**，从中心向外辐射柱条，长度=数值比例，背景圆+端点圆点+标签 |
| `diverging` | items[{label, value}] | **正负双向条形**，中心线两侧，绿正红负 |
| `tornado` | items[{label, left_value, right_value}], left_label?, right_label? | **龙卷风图**，中心标签+左右背对背条形 |
| `honeycomb` | items[{label, value?, color?}] | **蜂巢图**，圆角矩形近似蜂窝，交错排列 |
| `slope` | items[{label, left, right}], left_label?, right_label? | **斜率图**，左右两轴+连线圆点，比较变化趋势 |
| `pictogram` | items[{label, value}], total? | **象形图**，圆形单元按百分比填色，100格 |
| `sunburst` | inner[{label, value}], outer[{label, parent_idx, value}]? | **旭日图**，内圈分类+外圈细分，同心圆环+图例 |

### 补齐 dashi-ppt 剩余高频布局（15种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `mekko` | items[{label, width, segments[{label, value}]}] | **Marimekko 变宽图**，宽度+高度双维度编码数据 |
| `grouped` | categories[], series[{label, values[], color?}] | **分组柱状图**，多系列并列柱+图例 |
| `trend` | series[{label, values[], color?}], labels[]? | **趋势折线图**，数据点+连线+X轴标签+图例 |
| `chain` | stages[{label, items[{name}]}] | **产业链/价值链**，阶段标题+项目卡片+箭头连接 |
| `calendar` | weeks[{values[7]}], day_labels[]? | **日历热力图**，7×N色块网格，颜色深浅=数值 |
| `orbit` | center, nodes[{label, orbit?}] | **轨道枢纽图**，中心圆+多圈层轨道+节点 |
| `triptych` | panels[{title, items[], color?}] | **三联面板**，三栏并列+标题条+内容列表 |
| `meter` | items[{label, value, max?, benchmark?}] | **计量条/进度条**，填充条+基准线+数值 |
| `pareto` | items[{label, value}] | **帕累托图**，排序柱+累积百分比线+80%标注 |
| `delta` | items[{label, before, after, unit?}] | **变化量对比**，前后数值+涨跌箭头+百分比 |
| `milestones` | milestones[{date, title, desc?}] | **里程碑时间轴**，轴线+圆点标记+日期+描述 |
| `spectrum` | items[{label, position}], left_label?, right_label? | **光谱定位图**，渐变条+标记圆点+标签 |
| `logowall` | items[{label}], cols? | **Logo 墙/伙伴墙**，网格排列+名称标签 |
| `masonry` | items[{title, desc?}] | **瀑布流网格**，不等高卡片交错排列 |
| `ladder` | stages[{label, value, dropoff?}] | **转化阶梯**，递降条形+流失标注 |

### 动画布局静态化适配（11种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `mindmap` | center, branches[{label, leaves[{label}]}] | **思维导图 / 放射树**，中心圆+放射分支+对角连线+叶子节点，6色轮转 |
| `network` | nodes[{label}], edges[{source, target}], center? | **网络节点图**，环形节点+连线+可选中心节点 |
| `mosaic` | items[{title}] | **图片拼贴 / 马赛克**，1大+5小布局，6色填充 |
| `sticker_bubble` | big_number, big_label, satellites[{label, value}] | **卫星数据 / 估值泡沫**，左侧大数字+右侧卫星卡 |
| `bubbletl` | items[{label, value, date}] | **气泡时间线**，基线+气泡圆点（大小=数值）+日期标签 |
| `icicle` | items[{label, value, children[{label, value}]}] | **冰柱图 / 层级分解**，顶层分类+子层细分矩形 |
| `candles` | items[{open, close, high, low, label}] | **K线图 / 蜡烛图**，实体+影线，绿涨红跌 |
| `hypecycle` | items[{label, position}], phases[] | **成熟度曲线 / Hype Cycle**，S型曲线+阶段标签+数据点 |
| `typeriver` | words[{text, size?}], lead? | **字阵 / 标语流**，多字号+多颜色关键词流 |
| `ribbon` | items[{label, value}] | **全幅比例带**，宽度=占比，横向铺满 |
| `vinyl` | title, tracks[{label, duration}] | **唱片 / 播放列表**，左侧唱片造型+右侧曲目列表 |

### dashi-ppt 蒸馏迁移（20种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `polar_rose` | items[{label, value}] | **玫瑰图 / 南丁格尔图**，变半径扇形+图例 |
| `histogram` | bins[{range, count}], x_label?, y_label? | **直方图 / 频率分布**，相邻竖条+区间标签 |
| `quotewall` | quotes[{text, author?}] | **群言墙 / 多引述马赛克**，2×3卡片+引号装饰 |
| `metro` | lines[{name, color?, stations[{name}]}] | **地铁线路图**，多条线路+站点圆点 |
| `balance` | left{title, items[]}, right{title, items[]} | **天平 / 权衡对比**，横梁+左右两栏 |
| `fiveforces` | center, forces[{name, level?}] | **波特五力模型**，中心框+5周边框+连线 |
| `glossary` | items[{term, definition}] | **术语表 / 词汇表**，交替行+术语/定义两栏 |
| `album` | title, tracks[{name, detail?, year?}] | **专辑 / 成就清单**，封面区+编号曲目列表 |
| `bracket` | groups[{label, items[]}] | **分组括号图**，左侧标签+竖线+右侧子项 |
| `horizon` | views[{title, items[]}] | **三视野 / 地平线**，三栏并列+色条+列表 |
| `stack` | layers[{label, items[]}] | **架构栈 / 技术栈**，垂直层叠+标签+子项 |
| `gate` | layers[{label, desc?}] | **分层防线 / 门控模型**，同心嵌套矩形+标签 |
| `triad` | items[{label, desc?}], center? | **三角串联 / 三球**，三圆三角排列+连线 |
| `loop` | steps[{label}], center? | **闭环循环**，中心圆+周围节点环形 |
| `ecosystem` | center, nodes[{label, orbit?}] | **生态网络**，中心+辐射连线+节点 |
| `chronicle` | events[{year, title, desc?}] | **编年史 / 纵向时间线**，纵轴+年份+事件 |
| `manifesto` | text, sub? | **宣言 / 主张页**，深绿全宽卡+大文字+金装饰线 |
| `comparetable` | headers[], rows[{feature, values[]}] | **特性对照表**，表头+数据行+色标 |
| `dotfield` | items[{label, value, total?}] | **点阵计数 / 单位图**，彩色圆点网格 |
| `bullet` | items[{label, value, target, max?}] | **子弹图 / 目标达成**，填充条+目标线+数值 |

### 最终补齐（2种）

| 布局 | 字段 | 说明 |
|------|------|------|
| `combo` | categories[], bars[{label, values[], color?}], line{label, values[], color?} | **组合图 / 柱线双轴**，多系列柱+叠加折线+图例 |
| `stream` | series[{label, values[], color?}], labels[] | **主题河流 / 中心流式堆叠**，居中对称堆叠柱+时间标签+图例 |

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
| 多项能力/功能展示 | `card_grid` 或 `feature_grid` |
| 简单要点罗列 | `bullets` 或 `takeaway` |
| SWOT / 四象限分析 | `swot` 或 `quadrant` |
| 任务清单/里程碑 | `checklist` |
| 多维度评分 | `scorecard` |
| 阶段递进/成长路径 | `stair` |
| 循环模型/飞轮效应 | `flywheel` |
| 核心主张/使命宣言 | `statement` |
| 用户旅程/体验地图 | `journey` |
| 产品定价/会员方案 | `pricing` |
| 常见问题解答 | `faq` |
| 大数字/KPI | `stat` |
| 多指标横排 | `metrics` |
| 时间线/里程碑 | `timeline` |
| 双栏对照 | `two_column` 或 `comparison` |
| 引言/开场 | `intro` |
| 金句/名言 | `quote` |
| 比例/占比分布 | `pie_chart` |
| 数值对比排行 | `bar_chart` |
| 6 项综合指标 | `dashboard` |
| KPI 追踪/趋势 | `kpi_card` |
| 战略/愿景横幅 | `hero_banner` |
| 有序步骤说明 | `numbered_list` |
| 2×2 分析矩阵 | `matrix_2x2` |
| 图表占位（后期插入） | `chart_placeholder` |
| 多维度能力评估 | `radar_chart` |
| 层级/优先级结构 | `pyramid` 或 `layers` |
| 阶段规划/里程碑路线 | `roadmap` |
| 概念交集/关联分析 | `venn` |
| 排名/Top N | `ranking` |
| 数值累积/增减分解 | `waterfall` |
| 多维数据矩阵 | `heatmap` |
| 项目排期/时间规划 | `gantt` |
| 循环模型/闭环流程 | `cycle` |
| 核心大数字突出展示 | `big_number` 或 `stat` |
| 案例/作品展示 | `gallery` |
| 模块化数据速览 | `bento` |
| 达成率/完成度 | `gauge` |
| 用户评价/见证 | `testimonial` |
| 比例嵌套/面积分布 | `treemap` |
| 两变量相关性分析 | `scatter` |
| 多类别结构占比 | `stacked_bar` |
| 人物/公司档案介绍 | `profile` |
| 核心数据特写/案例聚焦 | `spotlight` |
| 风险识别与优先级 | `risk` |
| 跨职能流程/职责分工 | `swimlane` |
| 项目/业务摘要概览 | `overview` |
| 价值观/核心准则 | `principles` |
| 组织层级/生态关系 | `org_chart` |
| 排名变化/竞争态势 | `bump` |
| 前后对比/变化幅度 | `dumbbell` |
| 简洁数值排行 | `lollipop` |
| 百分比/完成度可视化 | `waffle` 或 `pictogram` |
| 多维度达标率 | `radial_bar` |
| 正负值/盈亏对比 | `diverging` |
| 两组数据背对背对比 | `tornado` |
| 蜂窝状分类展示 | `honeycomb` |
| 两时点趋势对比 | `slope` |
| 层级结构/分类细分 | `sunburst` |
| 双维度占比/市场规模 | `mekko` |
| 多系列分类对比 | `grouped` |
| 时间序列趋势 | `trend` |
| 产业/价值链关系 | `chain` |
| 日度/周度数据热力 | `calendar` |
| 核心-外围关系 | `orbit` |
| 三栏并列对比 | `triptych` |
| 指标达标度/进度 | `meter` |
| 关键少数/二八分析 | `pareto` |
| 前后变化量/增长率 | `delta` |
| 关键节点里程碑 | `milestones` |
| 定位/光谱/程度 | `spectrum` |
| 合作伙伴/客户展示 | `logowall` |
| 不等高卡片陈列 | `masonry` |
| 转化漏斗/留存阶梯 | `ladder` |
| 思维导图/放射树/发散结构 | `mindmap` |
| 网络关系/联盟生态 | `network` |
| 图片拼贴/案例画廊 | `mosaic` |
| 核心数据+子项展开 | `sticker_bubble` |
| 事件时间线+数值规模 | `bubbletl` |
| 层级分解/分类结构 | `icicle` |
| 价格波动/金融走势 | `candles` |
| 技术成熟度/创新周期 | `hypecycle` |
| 关键词/核心概念流 | `typeriver` |
| 全幅占比/简单比例 | `ribbon` |
| 列表/播放清单/成就列表 | `vinyl` |
| 极坐标/玫瑰图分布 | `polar_rose` |
| 频率分布/数据区间 | `histogram` |
| 多人引言/观点汇总 | `quotewall` |
| 多线路/站点关系 | `metro` |
| 两方案权衡/优劣对比 | `balance` |
| 行业竞争/战略分析 | `fiveforces` |
| 专业术语/概念解释 | `glossary` |
| 专辑/作品/成就编号列表 | `album` |
| 分组归类/括号展开 | `bracket` |
| 短中长期/多视野规划 | `horizon` |
| 技术架构/层级堆叠 | `stack` |
| 安全防护/分层门控 | `gate` |
| 三角关系/三要素 | `triad` |
| 循环闭环/PDCA | `loop` |
| 生态关系/核心辐射 | `ecosystem` |
| 年度大事记/纵向时间 | `chronicle` |
| 使命宣言/核心主张 | `manifesto` |
| 功能特性/版本对比 | `comparetable` |
| 达成率/百分比可视化 | `dotfield` |
| 目标达成/KPI追踪 | `bullet` |
| 柱状+折线双轴组合 | `combo` |
| 多主题流式变化趋势 | `stream` |

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
3. 内容页：克隆 slide4，清除原元素，复用 rId2 底部栏（避免重复），按布局添加文本框和形状
4. 新增形状原语：`make_rect`（圆角矩形）、`make_circle`（圆形）、`make_line`（水平线）、`_diag_line`（任意角度对角线，通过旋转矩形实现），支持填充色+文本
5. 结束页：保留模板原样
6. 重新排序 sldIdLst，清理无用文件，打包输出

### 章节页文本替换规则（edit_section）

模板 slide3 有 4 个文本区域，按以下条件顺序匹配替换，互不干扰：

| 条件 | 匹配目标 | 替换为 |
|------|---------|--------|
| 文本含 "PART" | PART 标签（id=10） | `PART`（固定，不加序号） |
| `text.strip().isdigit()` | 左上角序号（id=9，原值 "01"） | `data['part']`（如 "02"、"03"） |
| 以 "20" 开头 | 年份（id=30 段落1） | `data['year']` |
| 其余非数字文本 | 章节标题（id=30 段落2） | `data['title']` |

**标题换行处理**：标题文本框（id=30）做以下调整以支持长标题：
- `wrap="none"` → `wrap="square"`（开启换行）
- `cx` 从 1805305 扩到 **3500000**（到金色边框右边界，x 位置不变）
- 字号按标题长度自动缩放，保证最多两行：
  - ≤10 字 → 28pt
  - 11~20 字 → 24pt
  - \>20 字 → 20pt

### 内容页底部栏

内容页克隆 slide4 后，底部栏（image11.png）由模板 rels 中的 `rId2` 提供，**不得额外添加 rId5**（否则重复渲染）。生成代码中底部栏尺寸使用模板精确值：
- `BAR_Y = 6007100`（EMU）
- `BAR_H = 698500`（EMU）

### 内容页颜色方案

内容页背景（image10.jpeg）为**白色**，所有元素颜色须适配白底：

| 元素类型 | 颜色策略 |
|---------|---------|
| 标题 | 品牌绿 #46A53B（白底高辨识度） |
| 正文/要点 | 深灰 #333333 |
| 辅助说明 | 中灰 #555555 |
| 强调数字/指标 | 品牌绿 #46A53B 或金色 #FBB03B |
| 卡片填充（dashiai风格） | 品牌色系实填：`46A53B`/`213A25`/`FBB03B`/`2D7A35`/`D4940A`/`3B8A52` |
| 卡片内文字 | 白色 #FFFFFF（品牌色卡片上白字） |
| 金色分割线 | 已移除（不在内容页添加） |

### 目录页（TOC）处理规则

模板 slide2 有 5 个固定占位：`目录标题1`~`目录标题5`，各带对应数字序号圆圈。`edit_toc` 的处理逻辑：
- 条目数 ≤ 5：替换前 N 个标题，**删除多余的标题 sp 和对应序号圆圈 sp**，不留空白占位
- 条目数 > 5：仅替换前 5 个（模板上限）
- 序号圆圈与标题 sp 是独立元素，均需单独删除

### 内容页布局高度约束

内容页可用垂直空间：`BODY_Y=emu(2.8)` 到 `BAR_Y=6007100`（约 6.57in），共约 **3.77in**。各布局必须遵守：

| 布局 | 最大高度约束 | 说明 |
|------|------------|------|
| swot / quadrant | 每象限 `emu(1.6)`，总高 ≈3.4in | 2行×1.6in + 0.2in间距 |
| pricing | 卡片高 `emu(3.5)` | 3列并排，底部距底栏 ≥0.3in |
| flywheel | 区域 `emu(3.5)` | 从 BODY_Y 起算，总高 3.5in |
| statement | 主文字 `emu(2.8)` + 副标题 `emu(0.5)` | 合计约 3.6–4.0in（含引号装饰） |
| highlight_box | 主框 `emu(2.2)` + 副卡 `emu(1.3)` | 合计约 3.8in |
| stair | 底部对齐 `y + emu(3.0)` | 最高台阶不超过 3.2in |
| principles | 5行固定 `emu(3.0)` | 每行 `emu(0.5)+emu(0.1)` 间距 |
| org_chart | 4层固定 `emu(3.15)` | `node_h=emu(0.6) + gap=emu(0.65)` |
| risk | 网格+列表 `emu(3.0)` | 3×3网格 `emu(2.4)` + 右侧列表 |
| swimlane | 4行 `emu(3.2)` | 每行 `emu(0.7)+emu(0.1)` 间距 |
| bump | 排名区 `emu(2.8)` + 标签 | 水平排名线间距按系列数均分 |
| dumbbell | 每行 `emu(0.45)+emu(0.15)` | 6行合计约 `emu(3.6)` |
| slope | 区域 `emu(2.8)` + 标签 | 6条连线+左右标签 |
| tornado | 每行 `emu(0.38)+emu(0.1)` | 7行+图例约 `emu(3.6)` |
