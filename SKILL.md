---
name: handout-to-courseware
description: 小学数学讲义+教学设计稿→课件 PPT 的渲染 skill。输入「讲义 docx + 教学设计稿 docx + PPT 模板 pptx」三个文件，输出完整 .pptx 课件。当用户说"讲义转课件"、"出课件"、"做课件 PPT"、"渲染课件"、"按教学设计稿生成 PPT"、"把讲义转成 PPT"、"小学数学课件生成"时触发此 skill。注意：这个 skill 不创造教学内容（故事/思考题/笔记），那些必须从教学设计稿里来；它的核心任务是把"剧本（教学设计稿）+ 题库（讲义）+ 舞台（模板）"映射并渲染成最终 PPT。
---

# 讲义+教学设计稿 → 课件 PPT

## 这个 skill 在做什么

```
讲义 docx          ┐
教学设计稿 docx   ─┼─→  解析  →  映射页结构  →  渲染  →  .pptx
PPT 模板 pptx     ┘
```

**核心定位**：剧本 → 舞台渲染。教学设计稿是剧本（有故事、思考题、笔记），讲义是题库（提供例题/练习的题目本身），模板是舞台（提供版式）。skill 不创造内容，只做映射 + 渲染。

## 三件必须确保的事（其他都可妥协）

1. **schema 一致**：教学设计稿和讲义必须按预定 schema 解析。schema 不符直接报错让老师补，**不要让 LLM 静默补全**。
2. **页结构可见**：Phase 2 产出的中间 JSON（每页用哪个版式、填什么）**必须落盘** `output/slide_structure.json`，让老师能审 + 失败可恢复。
3. **题目编号严格对齐**：教学设计稿引用"例 2"必须对应讲义的 example_num=2，不匹配直接报错。**禁止 LLM 编造题目**。

## 四阶段流程

### Phase 1：解析三个输入（确定性，全脚本）

```bash
# 解析讲义 → handout.json
python3 {SKILL_DIR}/scripts/parse_handout.py <handout.docx> -o output/handout.json

# 解析教学设计稿 → teaching_design.json
python3 {SKILL_DIR}/scripts/parse_teaching_design.py <design.docx> -o output/teaching_design.json

# 解析 PPT 模板 → 列出版式（人工检查一次即可，结果写进 known_template_layouts.md）
python3 {SKILL_DIR}/scripts/parse_template.py <template.pptx>
```

如果模板已在 [references/known_template_layouts.md](references/known_template_layouts.md) 注册（如三年级A+ 春季），跳过第三步。

详细 schema 见 [references/teaching_design_schema.md](references/teaching_design_schema.md) 和 [references/handout_schema.md](references/handout_schema.md)。

### Phase 2：映射剧本 → 页结构（混合：规则 + LLM）

读教学设计稿的 segments + activities，决定每个 activity 展开成几页 PPT、用什么版式。

**主版式由规则决定**（见 [references/activity_to_layout_rules.md](references/activity_to_layout_rules.md)）：

```
活动类型           →  主版式      →  默认页数
情境创设/故事推进   →  内文2       →  1
探究活动/思考       →  内文2       →  1-3 (按子思考数)
技能学习/例题解析   →  内文1       →  1-2 (题干+解析)
学生练习            →  1_内文3     →  1
方法总结/笔记       →  空白页1     →  1-2
模块封面            →  空白页5     →  1
```

**LLM 在边缘情况上发挥**：
- "情境创设"段含两个分立场景 → 展开 2 页（LLM 判断）
- "探究活动"含 3 个子思考 → 展开 3 页（按 markdown 子弹列表数）
- "学生练习"题目超长 → 自动分页

**LLM 决策的 prompt 结构**详见 [references/mapping_prompt.md](references/mapping_prompt.md)。

输出：`output/slide_structure.json`（每张 slide 含 layout_name + slide_role + placeholder_data + 引用关系）。

**老师可在这一步审 slide_structure.json**（页数对不对、版式选得对不对、内容齐不齐）。审完才进 Phase 3。

### Phase 3：内容回填（确定性，全脚本）

```bash
python3 {SKILL_DIR}/scripts/fill_content.py \
  --slide-structure output/slide_structure.json \
  --handout output/handout.json \
  --teaching-design output/teaching_design.json \
  -o output/slide_structure_filled.json
```

**回填规则**：
- 故事段（slide_role=story）→ 用教学设计稿原文
- 思考题（slide_role=thinking）→ 教学设计稿"探究活动"的子段
- **例题题干 → 讲义** (固定来源；教学设计稿如果不一致也不允许覆盖)
- **例题解析 → 教学设计稿** (per ④A 原则)
- 练习题干 → 讲义
- 练习解析 → 教学设计稿（如果有），否则讲义
- 笔记 → 教学设计稿"方法总结"

**严格匹配**：所有"例 N / 练 N"引用必须在讲义里能找到对应的 example/practice，否则报错。

### Phase 4：渲染 PPT（确定性，全脚本）

```bash
python3 {SKILL_DIR}/scripts/render_pptx.py \
  --slide-structure output/slide_structure_filled.json \
  --template <template.pptx> \
  -o output/final.pptx
```

渲染规则（继承自 ppt-new-template）：
- 中文字体必须显式设置 `ea typeface="微软雅黑"`
- 公式必须用 OMML 渲染，不能纯文本
- 占位符缺失时从 layout 复制 sp 元素补齐

渲染完执行排版校验：

```bash
python3 /Users/gaotu/product-design-space/.claude/skills/ppt-new-template/scripts/check_typography.py output/final.pptx
```

## 如何使用此 skill

### 完整命令

```bash
# 一键跑通 4 阶段
bash {SKILL_DIR}/scripts/run_pipeline.sh \
  --handout "<path-to-handout.docx>" \
  --teaching-design "<path-to-design.docx>" \
  --template "<path-to-template.pptx>" \
  --output-dir output/<case-name>/
```

### 分步骤跑（推荐第一次用）

```bash
# Step 1: 解析（任意一步失败立即报错）
python3 {SKILL_DIR}/scripts/parse_handout.py "<handout.docx>" \
  -o output/<case>/handout.json
python3 {SKILL_DIR}/scripts/parse_teaching_design.py "<design.docx>" \
  -o output/<case>/teaching_design.json

# Step 2: 映射（LLM 调用，可能需要几十秒）
python3 {SKILL_DIR}/scripts/map_to_slides.py \
  --teaching-design output/<case>/teaching_design.json \
  --handout output/<case>/handout.json \
  --template <template.pptx> \
  -o output/<case>/slide_structure.json

# ⚠️ 老师审 slide_structure.json，确认 OK 再继续

# Step 3: 回填
python3 {SKILL_DIR}/scripts/fill_content.py \
  --slide-structure output/<case>/slide_structure.json \
  --handout output/<case>/handout.json \
  --teaching-design output/<case>/teaching_design.json \
  -o output/<case>/slide_structure_filled.json

# Step 4: 渲染
python3 {SKILL_DIR}/scripts/render_pptx.py \
  --slide-structure output/<case>/slide_structure_filled.json \
  --template <template.pptx> \
  -o output/<case>/final.pptx
```

## 常见失败 & 防御

| 失败信号 | 原因 | 处理 |
|---|---|---|
| `KeyError: 'example_3'` 在 Phase 3 | 教学设计稿引用了讲义里没有的题号 | 报错让老师补讲义题，或修教学设计稿 |
| `LayoutNotFound: 内文2` 在 Phase 4 | 模板版式名不一致（如"内文 2"vs"内文2"）| 检查 references/known_template_layouts.md 是否注册 |
| PPT 渲染后字体是黑体不是微软雅黑 | `_set_run_lang` 没调用 | 复用 ppt-new-template 的字体规则 |
| 公式显示乱码 | OMML 没正确渲染 | 检查 limUpp 展开、oMath 不含 U+200B |
| 同一 activity 被映射成 5 页 PPT（过多）| LLM 边缘判断失误 | 在 mapping_prompt.md 加约束 "单 activity ≤ 3 页" |
| `slide_structure.json` 里 placeholder 为空 | 教学设计稿对应段落为空或解析失误 | 在 fill_content.py 加 assert，报错让老师补内容 |

## 已注册的模板

详见 [references/known_template_layouts.md](references/known_template_layouts.md)。

| 模板名 | 用途 | 注册状态 |
|---|---|---|
| 三年级A+ 春季 课件模板 | 小学三/四/五年级 数学 春季 | ✅ 已注册 |
| 其他 | - | 用 parse_template.py 注册 |

## 跟其他 skill 的关系

- **`ppt-new-template`**：用于注册新模板（一次性）。本 skill 复用它的字体/公式规则。
- **`ppt-script-fix`**：用于排版校验 + 修复脚本。本 skill 渲染完后调它的 check_typography.py。

## 当前状态：v0（only 分解质因数 case 验证）

本 skill v0 仅经过分解质因数一个 case 验证。准备拿更多 case 后再迭代。
