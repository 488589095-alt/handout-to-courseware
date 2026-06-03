---
name: handout-to-courseware
description: 小学数学讲义+教学设计稿→课件 PPT 的渲染 skill。输入「讲义 docx + 教学设计稿 docx + PPT 模板 pptx」，输出完整 .pptx 课件。当用户说"讲义转课件"、"出课件"、"做课件 PPT"、"按教学设计生成 PPT"、"把讲义转成 PPT"、"小学数学课件生成"、"基于教学设计稿做 PPT"时主动触发。注意：这个 skill 不创造教学内容（故事/思考题/笔记），那些必须从教学设计稿原文里来。核心任务是把"剧本（教学设计稿）+ 题库（讲义）+ 舞台（模板）"严格映射并渲染成最终 PPT。
---

# 讲义+教学设计稿 → 课件 PPT (v1)

## 这个 skill 在做什么

```
讲义 docx           ┐
教学设计稿 docx     ┼─→  解析  →  生成 slide_structure  →  渲染  →  .pptx
PPT 模板 pptx       ┘
```

**核心定位**：剧本 → 舞台渲染。教学设计稿是剧本（有故事、思考题、笔记），讲义是题库（提供例题/练习的题目本身），模板是舞台（提供版式）。skill 不创造内容，只做映射 + 渲染。

## 三件硬性原则（小学课件甲方要求，不可妥协）

1. **正文字体统一微软雅黑 24 号**：小学生课件硬要求，字号小看不清。占位符标题、表格、封面/结束页大字按各自规则（详见 [references/typography_rules.md](references/typography_rules.md)）。

2. **严格按教学设计的"对应课件 PN"标注拆页**：教学设计若标了 `（对应课件P9-P13）`，那这 5 页就是 5 张 slide，**不能合并成 1 张**。slide_num 严格等于 P 号。详见 [references/p_number_mapping.md](references/p_number_mapping.md)。

3. **不创造内容**：故事/思考题/笔记/口诀的文字必须来自教学设计稿原文；例题/练习题干必须来自讲义。教学设计没标"课后任务"就不要自己加。教学设计写了"凑3口诀"就照原文用。

## 四阶段流程

### Phase 1：解析三个输入（脚本，确定性）

```bash
python3 {SKILL_DIR}/scripts/parse_handout.py <handout.docx> \
  -o output/<case>/handout.json
python3 {SKILL_DIR}/scripts/parse_teaching_design.py <design.docx> \
  -o output/<case>/teaching_design.json
```

模板版式如已在 [references/known_template_layouts.md](references/known_template_layouts.md) 注册，跳过模板解析。

详细 schema：[references/teaching_design_schema.md](references/teaching_design_schema.md) | [references/handout_schema.md](references/handout_schema.md)

### Phase 2：决定映射模式（核心新增！）

读取教学设计稿，**先扫一遍是否含"对应课件 P" 字符串**：

```python
has_pn = "对应课件P" in teaching_design_text or "对应课件 P" in teaching_design_text
```

| 模式 | 触发条件 | 映射策略 |
|---|---|---|
| **严格 P 号模式** | 教学设计有"对应课件 PN"标注 | slide_num 直接等于 PN；每个范围（P9-P13 = 5 页）必须拆成对应张数 |
| **LLM 推断模式** | 教学设计无 PN 标注 | 按 [references/activity_to_layout_rules.md](references/activity_to_layout_rules.md) 的规则展开活动 → 页 |

实践中**严格模式应优先**：高途的标准教学设计稿都带 PN 标注（牛吃草、必胜策略、递推计数均如此）。详细规则见 [references/p_number_mapping.md](references/p_number_mapping.md)。

输出：`output/<case>/slide_structure.json`（每张 slide 含 `slide_num + slide_role + source + placeholders + add_text_boxes / add_tables / add_images`）。

**slide_structure.json 必须落盘**让老师能审 + 失败可恢复。

### Phase 3：抽取讲义 / 教学设计资产（脚本）

#### 知识阶梯图（从讲义抽）

```python
import zipfile, os
# 讲义 docx 内嵌图片，最大那张（~200KB）通常就是知识阶梯（横向流程图：前铺/本讲/后续）
# 详见 references/table_and_asset_extraction.md
```

#### 教学设计里的表格（按 docx body 顺序）

教学设计稿通常有 5-9 个表格，分两类：
- **教学内容表格**（在"对应课件 PN"范围内出现）→ 插入到对应 P 号的 slide
- **元数据表格**（基本信息 / 课件内容结构 / 学情分析 / 教学方法）→ 作为"教师备课参考"附录，放在课后任务之后、结束页之前

详细抽取代码 + cell 文本去重问题见 [references/table_and_asset_extraction.md](references/table_and_asset_extraction.md)。

### Phase 4：渲染 PPT（脚本）

```bash
python3 {SKILL_DIR}/scripts/render_pptx.py \
  --slide-structure output/<case>/slide_structure.json \
  --template "<template.pptx>" \
  -o output/<case>/最终课件.pptx
```

`render_pptx.py` 五阶段渲染（顺序很重要）：

```
duplicate（基于原始 sample，先于 modify）
   ↓
template_idx（修改既有 sample slide 字段）
   ↓
add_layout（用 layout 新建无装饰页）
   ↓
删除未使用的 sample slide
   ↓
按 slide_num 严格重排（用 sldId element 跟踪，不靠文本匹配）
```

**⚠️ duplicate 必须在 modify 之前**，否则副本会复制到已被修改过的内容，造成内容叠加。

支持的字段：`placeholders`（占位符填值）、`free_shapes`（命名 shape 填值）、`add_text_boxes`（自由文本框）、`add_tables`（表格）、`add_images`（图片）、`resize_placeholders`（调整占位符位置/大小）。

## 关键避坑（v1 经验，必读）

| 现象 | 根因 | 处理 |
|---|---|---|
| **表格里 `2ⁿ` 显示成"2쒊"等韩文乱码** | 微软雅黑没有 U+207F 上标字形，被 fallback 成其他 CJK 字符集 | 一律改用 ASCII：`2^n`、`2^7`、`2ⁿ+1` → `2^n + 1`。详见 [references/typography_rules.md](references/typography_rules.md) |
| **拆页过粗**（教学设计标 P9-P13 但只做 1 页） | 没按 P 号严格拆 | 强制 slide_num == P 号，详见 [references/p_number_mapping.md](references/p_number_mapping.md) |
| **本讲总结页文字重叠** | 模板自带"本讲总结"标题在 top≈1.0，自定义文本框 top=1.0 覆盖了它 | 自定义文本框 top ≥ 1.7 |
| **例题题干靠左偏** | 模板 ph[16] 默认 left=2.3 width=10.5（非居中） | 用 `resize_placeholders: {"16": {"left": 0.65, "width": 12.0}}` |
| **副本和原页内容叠加** | duplicate 在 modify 之后跑 | duplicate 必须先 |
| **slide_num 重排错位** | 靠文本匹配定位 slide，多个 slide 含同一关键词时找错 | 用 `sldId` element 跟踪 |
| **加 add_slide(layout) 后丢装饰** | layout 不含 sample slide 上的装饰图 | 用 `template_idx` 修改既有 sample（不要新建） |
| **知识阶梯页背景图删不掉** | 装饰图在 layout 上（不是 slide 上），不能 deletion | 换无装饰 layout（"标题-文字少"） |
| **cell 文本重复 3 次**（"项目项目项目"） | docx 单元格有 3 段相同内容 | 取 `cell.paragraphs[0].text` |
| **教学设计自添加"课后任务"** | LLM 自由发挥 | 严格按教学设计原文写的 P 号 + 内容 |
| **多小点应拆页但被并到一张** | 例如"破解思路/尝试探索/方法总结"必须 3 页（题目要常驻顶部） | 按教学设计的 1/2/3 编号拆，题目作为标题保持 |

## 内容长度 vs 字号 24 的兼容

24 号字下，文本框 `width=10.5 height=5.7` 大约能装 200-250 汉字字符。如果超出：
- 优先**拆 P**（教学设计若给了 P 范围，往往本意就是分多页）
- 其次**精简内容**（去除"师：xxx 生：xxx"对话冗余，保留核心观点）
- 最后才**调字号**（如 22→20，但要先和老师确认）

同页**文本框 + 表格**的标准布局：

```
文本框：top=1.0, height=3.0  （上半）
表格  ：top=4.3, height=3.0  （下半）
```

## 附录顺序（v1 固定）

教师备课参考表（共 4 张）的固定顺序，严格按教学设计原文出现顺序：

```
... 主体内容 ...
P_n   课后任务（最后一张内容页）
P_n+1 附录·基本信息
P_n+2 附录·课件内容结构
P_n+3 附录·学情分析
P_n+4 附录·教学方法
P_n+5 结束页（"下节课再见！"）  ← 永远最后一张
```

## 如何使用此 skill

### 完整命令

```bash
# Step 1: 解析
python3 {SKILL_DIR}/scripts/parse_handout.py "<handout.docx>" \
  -o output/<case>/handout.json
python3 {SKILL_DIR}/scripts/parse_teaching_design.py "<design.docx>" \
  -o output/<case>/teaching_design.json

# Step 2: 检测模式 → 生成 slide_structure.json
#   - 严格模式：LLM 按教学设计的"对应课件PN"标注 + 表 2 模块范围生成（参考 references/p_number_mapping.md）
#   - LLM 模式：按 references/activity_to_layout_rules.md 展开

# Step 3: 抽资产（知识阶梯图 + 教学设计表格）
#   参考 references/table_and_asset_extraction.md 的脚本片段

# ⚠️ 老师审 slide_structure.json，确认 OK 再继续

# Step 4: 渲染
python3 {SKILL_DIR}/scripts/render_pptx.py \
  --slide-structure output/<case>/slide_structure.json \
  --template "<template.pptx>" \
  -o output/<case>/最终课件.pptx
```

## 已注册的模板

详见 [references/known_template_layouts.md](references/known_template_layouts.md)。

| 模板名 | 用途 | 注册状态 |
|---|---|---|
| 三年级A+ 春季 课件模板 | 小学三/四/五年级 数学 春季 | ✅ 已注册（用例：分解质因数、牛吃草、必胜策略、递推计数）|

## 跟其他 skill 的关系

- **`ppt-new-template`**：注册新模板（一次性）。本 skill 复用它的字体/公式规则。
- **`ppt-script-fix`**：排版校验 + 修复。本 skill 渲染完后调它的 check_typography.py。

## 版本

- **v1.0**（当前）：经过 4 个 case 验证（分解质因数、牛吃草、必胜策略、递推计数），含双 mode 检测、字号 24 硬约束、表格抽取、字符兼容性、附录顺序。
- v0.x：仅分解质因数验证，已废弃。
