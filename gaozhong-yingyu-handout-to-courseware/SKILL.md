---
name: gaozhong-yingyu-handout-to-courseware
description: 高中英语「讲义 docx + PPT模版 → 课件 pptx」生成 skill（第9讲读后续写+第10讲阅读理解双讲次验证）。输入讲义+原始课件模板即可出整套课件：自动识别讲次类型（主观题·读后续写 / 客观题·阅读），按题型分页（阅读/七选五/语法填空/主观题翻译每页1题、完形每页3-4题），讲义内容100%转入，讲义没有的教学脚手架用AI生成并标"待师审"。标杆PPT仅用于一次性提炼结构与设计token，不克隆。当用户提到"英语讲义转课件"、"高中英语出课件/做PPT"、"读后续写/阅读讲义转PPT"、"第N讲英语课件"时使用。注意与 handout-to-courseware（小学数学）、yuwen-handout-to-courseware（高中语文）区分。
---

# 高中英语 · 讲义 → 课件 PPT（v1）

```
讲义.docx ──extract_handout.py──▶ content.json(A类) ─┐
AI 生成(C类脚手架) ────────────▶ ai_scaffold.json ──┼─ build_pptx.py ──▶ 课件.pptx
原始课件模板.pptx(B类fixture+品牌背景+设计token) ────┘                 + slide_structure.json
```

## 三条铁律（来自双讲次对标验证）

1. **讲义有的内容必须 100% 转入**（按模版格式、参考标杆样式）。
2. **讲义没有、标杆有的人工内容**：测评(课前/课后测)、页码标、配图插画 → **不生成**；
   教学脚手架(方法框架/翻译/概括/推断/推理题/凝练) → **AI 生成**，页面标注"待老师审核"，并出《AI生成内容审核清单》。
3. **没有图片就不要插入图片**：只允许讲义自带图（如 Knowledge Map）与模版自带装饰。

## A/B/C 内容来源配方（每页先判来源）

| 来源 | 判定 | 页型 |
|---|---|---|
| **A 讲义直出** | grep 讲义有 | 封面信息/Preview表/Leading-in/语篇/题源+难度★/题目+选项/答案+解析/翻译(基础·升级·情节描写)/方法表/情节切分内容/范文(由讲义答案拼接) |
| **B 模版直出** | 模版有 | 封面/PART分隔/结束页/品牌背景（克隆模版 fixture） |
| **C AI生成** | 都没有但教学需要 | 目录/固化情节概览/引导(你会怎么表达)/解题三步框架/全文中文翻译/原文概括/整体方向/细节推理(原文细节+推测题)/描写类型凝练 |

## 结构骨架（自动识别讲次类型）

**通用头尾**：封面 → 目录 → Knowledge Map(讲义图) → Preview → … → 结束页

**continuation 读后续写(主观题)**——讲义含「固化情节」：
```
PART1 固化情节: 分隔 → Leading-in → 5步概览(C) → 每情节[引导(C) → 基础版本译 → 升级版本译]
PART2/3 篇章:   分隔 → 题源 → 语篇+续写提示 →
   步骤一(高亮) → 全文对照(C译) → 整体走向[概括→揭示走向(C)]   ← 两页揭示
   步骤二(高亮) → 细节推理[题→答(C)]                          ← 两页揭示
   步骤三(高亮) → 情节切分[留白→填好①–⑧]                      ← 两页揭示
   情节描写 → 每条[中文→揭示英文]，左上=序号+情节名+首/次段+描写类型(C)
   参考范文(部分末尾)
```

**reading 阅读理解(客观题)**——讲义含「阅读X篇」：
```
PART1 辅助方法: 分隔 → 每方法[大字节名 → 例篇题源 → 语篇 → 方法表 → 每题(题→答含解析)]
PART2 篇章训练: 分隔 → 每篇[题源+难度★ → 语篇 → 每题(题→答含解析)]
```

**普适规律**：分析块前放"步骤分隔页"且高亮当前步；讲解内容一律**"先题/留白 → 再揭示"两页式**；
**分页规则**：阅读理解/七选五/语法填空/主观题翻译=每页1题；完形填空=每页3-4题。

## 执行流程

```bash
# 0.【Gate·必做】拆解模版 → template_spec（字体/字号/颜色/坐标/版式全清单）
#    新模版必跑；build_pptx.py 生成前会自动核对 spec，不一致即中止
python3 {SKILL_DIR}/scripts/dissect_template.py "<原始课件模板.pptx>" -o {SKILL_DIR}/references/
#    人工读一遍 references/template_spec.md，确认：标题槽样式(默认sz40/#BA7AC2/普惠体B@y0.98)、
#    fixture页索引(封面/带标题页/空白页/分隔/结束)、各版式用途 → 与 token/T 常量一致

# 1. 解析讲义（自动识别类型 + 抽 Knowledge Map 图）
python3 {SKILL_DIR}/scripts/extract_handout.py "<讲义.docx>" -o output/<讲次>/

# 2. 生成 C 类 AI 脚手架 → output/<讲次>/ai_scaffold.json（schema 见 references/content_schema.md）
#    continuation: plot_overview/step_framework/overall_summary/overall_direction/
#                  reasoning_q(含干扰项)/desc_types/passage_zh(全文翻译)
#    reading:     一般无需（讲义已含答案+解析）；可选方法讲解
#    ⚠️ AI 内容是草稿：逐项列入《AI生成内容审核清单.md》交老师审

# 3. 渲染（设计token + 模版fixture，不碰标杆）
python3 {SKILL_DIR}/scripts/build_pptx.py \
  --content output/<讲次>/content.json \
  --template "<原始课件模板.pptx>" \
  -o "output/<讲次>/<讲次>_课件.pptx"

# 4. 自检：python3 {SKILL_DIR}/scripts/dump_pptx.py pptx <课件> -o dump.txt
#    检查：页数/每题1页/图片只在KM·分隔·结束/无标杆残留文字
```

## 人工确认 Gate

1. **slide_structure.json**（页序+角色）→ 老师确认结构再继续
2. **AI生成内容审核清单**（所有 C 类）→ 老师审核修改
3. 成片在 PowerPoint 逐页核对换行/溢出（本机无渲染器时必做）

## 设计 token / schema / 骨架明细

见 [references/design_tokens.md](references/design_tokens.md) ·
[references/content_schema.md](references/content_schema.md) ·
[references/structure_recipe.md](references/structure_recipe.md)

## 已验证用例

| 讲次 | 类型 | 产出 | 验证 |
|---|---|---|---|
| 第9讲 情节模板克服心魔 | continuation | 83页 | 与标杆逐页对标(差异=测评/配图/页码等人工增量) |
| 第10讲 辅助方法阅读精练 | reading | 52页 | **盲跑**(未碰其标杆)结构正确、每页1题 |

## 经验坑（必读）

- **字体字号铁律**：模版 shape 没写显式字号的（如封面标题/结束页副题）靠继承——回填文字时**绝不补写默认值**（`_style/setname` 传 None=不覆盖）；带标题内容页一律克隆模版"课前测页"换字（`_titled`），**禁止自创标题样式**（曾错用 sz36 加粗紫 #7A2E8E，模版实为 sz40 不加粗 #BA7AC2 @y0.98）

- 标杆里"组合(Group)"形状可能是**内容框**(如P41原文概括)而非装饰，删除前先看文字 → 本 skill 不克隆标杆，规避此坑
- 讲义 Word 的内嵌图要从 `word/media/` 抽（跳过 header/footer），它是 Knowledge Map=内容图，可放
- 答案块两种排布都要兼容：题后紧跟(主旨辅助式) / 全部题后集中块(B/C/D篇式，首条【答案】为空=主题概述要跳过)
- 模版 fixture 索引：封面=2、空白内容页=4、PART分隔=5/7/9、结束=12（换模版需重测）
- 知识标签是题库元数据 → 写 PPT 备注，不上正文
