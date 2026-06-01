# handout-to-courseware

> 小学数学 **讲义 + 教学设计稿 → 课件 PPT** 的渲染 skill（Claude Code 扩展）

把"剧本（教学设计稿）+ 题库（讲义）+ 舞台（PPT 模板）" 渲染为最终 .pptx 课件。**不创造内容**，严格按教学设计稿原文映射。

---

## 这是什么

一个 Claude Code 的 [skill](https://docs.claude.com/en/docs/build-with-claude-code/skills)，安装后可以让 Claude：

1. 解析讲义 docx（提取例题/练习）
2. 解析教学设计稿 docx（提取故事/思考题/笔记/总结）
3. 按教学设计稿的"对应课件 PN"映射或 LLM 推断，组织为 slide_structure.json
4. 用 python-pptx 渲染成最终 .pptx 课件（保留模板装饰）

---

## 安装到 Claude Code

```bash
# clone 到 Claude Code 的 skills 目录
mkdir -p ~/.claude/skills
git clone git@github.com:<你的用户名>/handout-to-courseware.git ~/.claude/skills/handout-to-courseware
```

然后在 Claude Code 里说"讲义转课件"或"按教学设计稿生成 PPT"即可触发。

---

## 4 阶段流程

```
讲义 docx ┐
教学设计稿 docx ─→ 解析 → 映射页结构 → 内容回填 → 渲染 → .pptx
PPT 模板 pptx ─┘
```

### Phase 1：解析三个输入（确定性）

```bash
python3 scripts/parse_handout.py 第1讲分解质因数.docx -o output/handout.json
python3 scripts/parse_teaching_design.py 分解质因数110分钟教学设计.docx -o output/teaching_design.json
```

### Phase 2：映射剧本 → 页结构（混合：规则 + LLM）

LLM（Claude）读 teaching_design.json + handout.json，按 [references/activity_to_layout_rules.md](references/activity_to_layout_rules.md) 的映射表 + 自己的判断，产出 `slide_structure.json`。

老师可以在这一步审中间产物，确认后再进入下一步。

### Phase 3-4：渲染 .pptx

```bash
python3 scripts/render_pptx.py \
  --slide-structure output/slide_structure.json \
  --template "三年级A+ 春季 课件模板.pptx" \
  -o output/最终课件.pptx
```

---

## 渲染策略

**关键决策**：不用 `add_slide(layout)` 新建（会丢失装饰元素），而是**修改模板已有 23 张 sample slide** 的字段。

slide_structure.json 里每张 slide 指定 source 类型：

| source | 用途 |
|---|---|
| `template_idx` | 修改模板第 N 张 sample slide 的字段（保留装饰）|
| `duplicate` | 复制模板第 N 张 sample slide → 改字段（用于需要多张同版式时）|
| `add_layout` | 用 layout 新建空白 slide（仅适用无装饰的 layout，如 `空白页5`）|

更多细节见 [SKILL.md](SKILL.md)。

---

## 案例

仓库内附 2 个真实案例的 `slide_structure.json`，可作为对照模板：

- [examples/分解质因数/](examples/分解质因数/) — 五年级数学，33 张 slide
- [examples/牛吃草/](examples/牛吃草/) — 四年级数学，34 张 slide（教学设计稿明确标了"对应课件 PN"）

---

## 限制 & 已知问题

- **目前仅小学数学场景验证**（分解质因数 + 牛吃草两个 case）
- 模板必须是高途的 "三年级A+ 春季 课件模板.pptx"。其他模板需要先用 `references/known_template_layouts.md` 重新登记版式映射
- 教学设计稿格式：建议严格按"`六、详细教学过程设计 → 第 N 环节 → 6.X.Y 小节（对应课件 P 号）`"组织，便于精准映射
- 不支持公式 LaTeX → OMML 渲染（公式直接显示为文本）

---

## 依赖

- Python 3.11+
- `pip install python-docx python-pptx lxml`

---

## License

MIT — 详见 [LICENSE](LICENSE)
